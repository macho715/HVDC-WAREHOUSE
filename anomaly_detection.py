#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
HVDC Warehouse 물류 이상 감지 시스템
- 장기체류 품목 자동으로 식별 (180일, 365일 등)
- 현재 미출고 상태 및 체류일수 계산
- 실무에서 바로 사용 가능한 엑셀 리포트
"""

import os
import sys
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import warnings
warnings.filterwarnings('ignore')

class AnomalyDetection:
    """물류 이상 감지 클래스"""
    
    def __init__(self, excel_path, sheet_name='CASE LIST'):
        """
        이상 감지기 초기화
        
        Args:
            excel_path: 엑셀 파일 경로
            sheet_name: 시트 이름
        """
        self.excel_path = excel_path
        self.sheet_name = sheet_name
        self.df = None
        self.warehouse_cols = []
        self.site_cols = []
        
        self._load_data()
        self._identify_columns()
        self._preprocess_data()
    
    def _load_data(self):
        """데이터 로드"""
        print(f"📁 데이터 파일 로드 중: {self.excel_path}")
        self.df = pd.read_excel(self.excel_path, sheet_name=self.sheet_name)
        print(f"✅ 데이터 로드 완료: {len(self.df)}행")
    
    def _identify_columns(self):
        """입고/출고 컬럼 식별"""
        # 입고 컬럼 (창고) 식별
        warehouse_patterns = ['DSV', 'Hauler', 'MOSB']
        self.warehouse_cols = []
        for col in self.df.columns:
            if any(pattern in str(col) for pattern in warehouse_patterns):
                self.warehouse_cols.append(col)
        
        # 출고 컬럼 (현장) 식별
        site_patterns = ['MIR', 'SHU', 'DAS', 'AGI']
        self.site_cols = []
        for col in self.df.columns:
            if any(pattern in str(col) for pattern in site_patterns):
                self.site_cols.append(col)
        
        print(f"🏭 입고 컬럼 ({len(self.warehouse_cols)}개): {self.warehouse_cols}")
        print(f"🏗️  출고 컬럼 ({len(self.site_cols)}개): {self.site_cols}")
    
    def _preprocess_data(self):
        """데이터 전처리"""
        print("\n🔧 데이터 전처리 중...")
        
        # 1) 날짜 컬럼들을 datetime으로 변환
        date_cols = self.warehouse_cols + self.site_cols
        for col in date_cols:
            if col in self.df.columns:
                self.df[col] = pd.to_datetime(self.df[col], errors='coerce')
        
        # 2) 입고일/출고일 계산
        self.df['입고일'] = self.df[self.warehouse_cols].min(axis=1)
        self.df['출고일'] = self.df[self.site_cols].max(axis=1)
        
        # 3) 출고여부 판단 (하나라도 출고일이 있으면 True)
        self.df['출고여부'] = self.df[self.site_cols].notna().any(axis=1)
        
        # 4) 현재 날짜 설정
        today = pd.Timestamp(datetime.today().date())
        
        # 5) 입고 후 경과일 계산
        self.df['입고후경과일'] = (today - self.df['입고일']).dt.days
        
        # 6) 현재 보관 창고 식별 (가장 마지막으로 입고된 창고)
        self.df['현재창고'] = None
        for idx, row in self.df.iterrows():
            warehouse_dates = row[self.warehouse_cols]
            valid_dates = warehouse_dates.dropna()
            if len(valid_dates) > 0:
                max_date_idx = valid_dates.idxmax()
                self.df.at[idx, '현재창고'] = max_date_idx
        
        # 7) 초기 입고 창고 식별
        self.df['초기창고'] = None
        for idx, row in self.df.iterrows():
            warehouse_dates = row[self.warehouse_cols]
            valid_dates = warehouse_dates.dropna()
            if len(valid_dates) > 0:
                min_date_idx = valid_dates.idxmin()
                self.df.at[idx, '초기창고'] = min_date_idx
        
        # 8) Case 상태 분류
        self.df['상태'] = self.df.apply(self._classify_case_status, axis=1)
        
        print("✅ 데이터 전처리 완료")
    
    def _classify_case_status(self, row):
        """Case 상태 분류"""
        if pd.isna(row['입고일']):
            return "미입고"
        elif pd.isna(row['출고일']):
            return "미출고"
        else:
            return "출고완료"
    
    def detect_long_stay_anomalies(self, threshold_days=180):
        """
        장기체류 이상 감지
        
        Args:
            threshold_days: 임계일수 (기본값: 180일)
            
        Returns:
            DataFrame: 장기체류 이상 Case 목록
        """
        print(f"\n🚨 장기체류 이상 감지 중 (임계값: {threshold_days}일)...")
        
        # 조건: 미출고이고 임계일수 이상 체류
        anomalies = self.df[
            (self.df['출고여부'] == False) & 
            (self.df['입고후경과일'] >= threshold_days)
        ].copy()
        
        # 체류일수 기준 내림차순 정렬
        anomalies = anomalies.sort_values('입고후경과일', ascending=False)
        
        print(f"✅ 장기체류 이상 감지 완료: {len(anomalies)}건")
        
        if len(anomalies) > 0:
            print(f"  📊 최장 체류일수: {anomalies['입고후경과일'].max()}일")
            print(f"  📊 평균 체류일수: {anomalies['입고후경과일'].mean():.1f}일")
            print(f"  📊 중앙값 체류일수: {anomalies['입고후경과일'].median():.1f}일")
        
        return anomalies
    
    def analyze_by_warehouse(self, anomalies_df):
        """
        창고별 장기체류 분석
        
        Args:
            anomalies_df: 장기체류 이상 DataFrame
            
        Returns:
            DataFrame: 창고별 분석 결과
        """
        print(f"\n🏭 창고별 장기체류 분석 중...")
        
        if len(anomalies_df) == 0:
            print("  ⚠️  장기체류 이상 Case가 없습니다.")
            return pd.DataFrame()
        
        warehouse_analysis = anomalies_df.groupby('현재창고').agg({
            'Case No.': 'count',
            '입고후경과일': ['mean', 'median', 'min', 'max']
        }).reset_index()
        
        warehouse_analysis.columns = ['창고', '건수', '평균체류일', '중앙값체류일', '최소체류일', '최대체류일']
        
        # 소수점 정리
        numeric_cols = warehouse_analysis.select_dtypes(include=[np.number]).columns
        warehouse_analysis[numeric_cols] = warehouse_analysis[numeric_cols].round(1)
        
        warehouse_analysis = warehouse_analysis.sort_values('건수', ascending=False)
        
        print(f"✅ 창고별 분석 완료: {len(warehouse_analysis)}개 창고")
        return warehouse_analysis
    
    def analyze_by_time_period(self, anomalies_df):
        """
        체류기간별 분석
        
        Args:
            anomalies_df: 장기체류 이상 DataFrame
            
        Returns:
            DataFrame: 체류기간별 분석 결과
        """
        print(f"\n⏰ 체류기간별 분석 중...")
        
        if len(anomalies_df) == 0:
            print("  ⚠️  장기체류 이상 Case가 없습니다.")
            return pd.DataFrame()
        
        # 체류기간 구간 설정
        def classify_stay_period(days):
            if days >= 365:
                return "1년 이상"
            elif days >= 270:
                return "9개월-1년"
            elif days >= 180:
                return "6개월-9개월"
            else:
                return "기타"
        
        anomalies_df['체류기간구분'] = anomalies_df['입고후경과일'].apply(classify_stay_period)
        
        period_analysis = anomalies_df.groupby('체류기간구분').agg({
            'Case No.': 'count',
            '입고후경과일': ['mean', 'median', 'min', 'max']
        }).reset_index()
        
        period_analysis.columns = ['체류기간', '건수', '평균체류일', '중앙값체류일', '최소체류일', '최대체류일']
        
        # 소수점 정리
        numeric_cols = period_analysis.select_dtypes(include=[np.number]).columns
        period_analysis[numeric_cols] = period_analysis[numeric_cols].round(1)
        
        period_analysis = period_analysis.sort_values('건수', ascending=False)
        
        print(f"✅ 체류기간별 분석 완료: {len(period_analysis)}개 구간")
        return period_analysis
    
    def get_urgent_cases(self, anomalies_df, urgent_threshold=365):
        """
        긴급 조치 필요 Case 식별 (1년 이상)
        
        Args:
            anomalies_df: 장기체류 이상 DataFrame
            urgent_threshold: 긴급 임계일수 (기본값: 365일)
            
        Returns:
            DataFrame: 긴급 조치 필요 Case 목록
        """
        print(f"\n🚨 긴급 조치 필요 Case 식별 중 (임계값: {urgent_threshold}일)...")
        
        urgent_cases = anomalies_df[anomalies_df['입고후경과일'] >= urgent_threshold].copy()
        urgent_cases = urgent_cases.sort_values('입고후경과일', ascending=False)
        
        print(f"✅ 긴급 조치 필요 Case: {len(urgent_cases)}건")
        
        if len(urgent_cases) > 0:
            print(f"  📊 최장 체류일수: {urgent_cases['입고후경과일'].max()}일")
            print(f"  📊 평균 체류일수: {urgent_cases['입고후경과일'].mean():.1f}일")
        
        return urgent_cases
    
    def generate_anomaly_report(self, threshold_days=180, urgent_threshold=365, output_file=None):
        """
        물류 이상 감지 리포트 생성
        
        Args:
            threshold_days: 장기체류 임계일수
            urgent_threshold: 긴급 조치 임계일수
            output_file: 출력 파일 경로
            
        Returns:
            str: 생성된 파일 경로
        """
        print(f"\n📋 물류 이상 감지 리포트 생성 중...")
        
        if output_file is None:
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            output_dir = os.path.join(os.path.dirname(__file__), 'outputs')
            os.makedirs(output_dir, exist_ok=True)
            output_file = os.path.join(output_dir, f'물류이상감지_{threshold_days}일+_{timestamp}.xlsx')
        
        # 1. 장기체류 이상 감지
        anomalies = self.detect_long_stay_anomalies(threshold_days)
        
        # 2. 긴급 조치 필요 Case 식별
        urgent_cases = self.get_urgent_cases(anomalies, urgent_threshold)
        
        # 3. 창고별 분석
        warehouse_analysis = self.analyze_by_warehouse(anomalies)
        
        # 4. 체류기간별 분석
        period_analysis = self.analyze_by_time_period(anomalies)
        
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            # 1. 전체 장기체류 이상 Case 목록
            print("  📋 전체 장기체류 이상 Case 목록 저장 중...")
            if len(anomalies) > 0:
                display_cols = ['Case No.', '현재창고', '초기창고', '입고일', '입고후경과일', '상태']
                if 'Material Category' in anomalies.columns:
                    display_cols.insert(2, 'Material Category')
                anomalies[display_cols].to_excel(writer, sheet_name=f'장기체류_{threshold_days}일+', index=False)
            
            # 2. 긴급 조치 필요 Case 목록
            print("  📋 긴급 조치 필요 Case 목록 저장 중...")
            if len(urgent_cases) > 0:
                urgent_cols = ['Case No.', '현재창고', '초기창고', '입고일', '입고후경과일']
                if 'Material Category' in urgent_cases.columns:
                    urgent_cols.insert(2, 'Material Category')
                urgent_cases[urgent_cols].to_excel(writer, sheet_name=f'긴급조치_{urgent_threshold}일+', index=False)
            
            # 3. 창고별 분석
            print("  📋 창고별 분석 저장 중...")
            if len(warehouse_analysis) > 0:
                warehouse_analysis.to_excel(writer, sheet_name='창고별_장기체류분석', index=False)
            
            # 4. 체류기간별 분석
            print("  📋 체류기간별 분석 저장 중...")
            if len(period_analysis) > 0:
                period_analysis.to_excel(writer, sheet_name='체류기간별_분석', index=False)
            
            # 5. 전체 미출고 Case 요약
            print("  📋 전체 미출고 Case 요약 저장 중...")
            pending_cases = self.df[self.df['출고여부'] == False].copy()
            pending_summary = {
                '분석 항목': [
                    '전체 Case 수',
                    '미출고 Case 수',
                    '미출고 비율 (%)',
                    f'{threshold_days}일 이상 장기체류 Case 수',
                    f'{threshold_days}일 이상 비율 (%)',
                    f'{urgent_threshold}일 이상 긴급 Case 수',
                    f'{urgent_threshold}일 이상 비율 (%)',
                    '미출고 Case 평균 체류일수',
                    '미출고 Case 중앙값 체류일수',
                    '미출고 Case 최대 체류일수',
                    '분석 창고 수'
                ],
                '값': [
                    len(self.df),
                    len(pending_cases),
                    round(len(pending_cases) / len(self.df) * 100, 1),
                    len(anomalies),
                    round(len(anomalies) / len(pending_cases) * 100, 1) if len(pending_cases) > 0 else 0,
                    len(urgent_cases),
                    round(len(urgent_cases) / len(pending_cases) * 100, 1) if len(pending_cases) > 0 else 0,
                    round(pending_cases['입고후경과일'].mean(), 1) if len(pending_cases) > 0 else 0,
                    round(pending_cases['입고후경과일'].median(), 1) if len(pending_cases) > 0 else 0,
                    pending_cases['입고후경과일'].max() if len(pending_cases) > 0 else 0,
                    len(pending_cases['현재창고'].dropna().unique())
                ]
            }
            summary_df = pd.DataFrame(pending_summary)
            summary_df.to_excel(writer, sheet_name='이상감지_요약', index=False)
        
        print(f"✅ 물류 이상 감지 리포트 생성 완료!")
        print(f"📄 파일명: {os.path.basename(output_file)}")
        print(f"📄 파일 크기: {os.path.getsize(output_file) / 1024:.1f} KB")
        
        return output_file

def main():
    """메인 실행 함수"""
    print("=== HVDC Warehouse 물류 이상 감지 시스템 ===")
    print(f"실행 시간: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    
    # 1. 데이터 파일 경로 설정
    current_dir = os.path.dirname(os.path.abspath(__file__))
    excel_path = os.path.join(current_dir, 'data', 'HVDC WAREHOUSE_HITACHI(HE).xlsx')
    
    if not os.path.exists(excel_path):
        print(f"❌ 오류: 데이터 파일을 찾을 수 없습니다: {excel_path}")
        return
    
    try:
        # 2. 이상 감지기 초기화
        print(f"\n📁 데이터 파일: {excel_path}")
        detector = AnomalyDetection(excel_path, sheet_name='CASE LIST')
        
        # 3. 다양한 임계값으로 이상 감지
        thresholds = [90, 180, 365]
        
        for threshold in thresholds:
            print(f"\n🔍 {threshold}일 이상 장기체류 이상 감지 중...")
            
            # 장기체류 이상 감지
            anomalies = detector.detect_long_stay_anomalies(threshold)
            
            if len(anomalies) > 0:
                print(f"\n📊 {threshold}일 이상 장기체류 상위 10건:")
                display_cols = ['Case No.', '현재창고', '입고일', '입고후경과일']
                print(anomalies[display_cols].head(10).to_string(index=False))
                
                # 창고별 분석
                warehouse_analysis = detector.analyze_by_warehouse(anomalies)
                if len(warehouse_analysis) > 0:
                    print(f"\n🏭 {threshold}일 이상 창고별 분석:")
                    print(warehouse_analysis.to_string(index=False))
        
        # 4. 긴급 조치 필요 Case 분석
        print(f"\n🚨 긴급 조치 필요 Case 분석 (365일 이상)...")
        urgent_cases = detector.get_urgent_cases(
            detector.detect_long_stay_anomalies(180), 
            365
        )
        
        if len(urgent_cases) > 0:
            print(f"\n📊 긴급 조치 필요 상위 10건:")
            display_cols = ['Case No.', '현재창고', '입고일', '입고후경과일']
            print(urgent_cases[display_cols].head(10).to_string(index=False))
        
        # 5. 종합 이상 감지 리포트 생성
        output_file = detector.generate_anomaly_report(threshold_days=180, urgent_threshold=365)
        
        # 6. 엑셀 파일 자동 열기
        try:
            os.startfile(output_file)
            print(f"\n🔓 엑셀 파일을 자동으로 열었습니다.")
        except:
            print(f"\n💡 엑셀 파일을 수동으로 열어주세요: {output_file}")
        
        print(f"\n📋 생성된 시트 목록:")
        print("  - 장기체류_180일+: 180일 이상 체류 Case 목록")
        print("  - 긴급조치_365일+: 365일 이상 긴급 조치 필요 Case")
        print("  - 창고별_장기체류분석: 창고별 장기체류 통계")
        print("  - 체류기간별_분석: 체류기간 구간별 분석")
        print("  - 이상감지_요약: 전체 이상 감지 요약 정보")
        
        return output_file
        
    except Exception as e:
        print(f"\n❌ 오류 발생: {str(e)}")
        import traceback
        traceback.print_exc()
        return None

if __name__ == "__main__":
    main() 