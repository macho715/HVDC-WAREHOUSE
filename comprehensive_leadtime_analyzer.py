#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
HVDC Warehouse 종합 리드타임 분석 시스템
- 전체 Case를 빠짐없이 포함 (미출고, 미입고 포함)
- Case별 상태 분류 (출고완료/미출고/미입고)
- 실무에서 바로 사용 가능한 엑셀 리포트
"""

import os
import sys
import pandas as pd
import numpy as np
from datetime import datetime
import warnings
warnings.filterwarnings('ignore')

class ComprehensiveLeadtimeAnalyzer:
    """종합 리드타임 분석 클래스"""
    
    def __init__(self, excel_path, sheet_name='CASE LIST'):
        """
        종합 리드타임 분석기 초기화
        
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
    
    def calculate_comprehensive_leadtime(self):
        """
        종합 리드타임 계산 (전체 Case 포함)
        
        Returns:
            DataFrame: 모든 Case가 포함된 리드타임 DataFrame
        """
        print("\n📊 종합 리드타임 계산 중...")
        
        # 1) 날짜 컬럼들을 datetime으로 변환
        date_cols = self.warehouse_cols + self.site_cols
        for col in date_cols:
            if col in self.df.columns:
                self.df[col] = pd.to_datetime(self.df[col], errors='coerce')
        
        # 2) 입고일 계산 (각 Case의 모든 입고 창고 일자 중 가장 이른 날짜)
        print("  📅 입고일 계산 중...")
        self.df['입고일'] = self.df[self.warehouse_cols].min(axis=1)
        
        # 3) 출고일 계산 (각 Case의 모든 출고 현장 일자 중 가장 늦은 날짜)
        print("  📅 출고일 계산 중...")
        self.df['출고일'] = self.df[self.site_cols].max(axis=1)
        
        # 4) 리드타임(일) 계산 (출고완료된 Case만)
        print("  ⏱️  리드타임 계산 중...")
        self.df['리드타임(일)'] = (self.df['출고일'] - self.df['입고일']).dt.days
        
        # 5) 초기 입고 창고명 식별
        print("  🏭 초기 입고 창고 식별 중...")
        self.df['초기창고'] = None
        for idx, row in self.df.iterrows():
            warehouse_dates = row[self.warehouse_cols]
            valid_dates = warehouse_dates.dropna()
            if len(valid_dates) > 0:
                min_date_idx = valid_dates.idxmin()
                self.df.at[idx, '초기창고'] = min_date_idx
        
        # 6) Case 상태 분류
        print("  📋 Case 상태 분류 중...")
        self.df['상태'] = self.df.apply(self._classify_case_status, axis=1)
        
        # 7) 현재 체류일수 계산 (미출고 Case용)
        print("  📅 현재 체류일수 계산 중...")
        today = pd.Timestamp.now()
        self.df['현재체류일수'] = (today - self.df['입고일']).dt.days
        
        print(f"✅ 종합 리드타임 계산 완료!")
        print(f"  📊 총 Case 수: {len(self.df)}")
        
        # 상태별 통계 출력
        status_counts = self.df['상태'].value_counts()
        print(f"  📊 상태별 분포:")
        for status, count in status_counts.items():
            print(f"    - {status}: {count}건")
        
        return self.df
    
    def _classify_case_status(self, row):
        """Case 상태 분류"""
        if pd.isna(row['입고일']):
            return "미입고"
        elif pd.isna(row['출고일']):
            return "미출고"
        else:
            return "출고완료"
    
    def analyze_by_status(self):
        """
        상태별 리드타임 통계 분석
        
        Returns:
            dict: 상태별 통계 데이터
        """
        print("\n📊 상태별 분석 중...")
        
        # 출고완료 Case만 필터링
        completed_cases = self.df[self.df['상태'] == '출고완료'].copy()
        
        if len(completed_cases) > 0:
            print(f"  📊 출고완료 Case: {len(completed_cases)}건")
            print(f"    - 평균 리드타임: {completed_cases['리드타임(일)'].mean():.1f}일")
            print(f"    - 중앙값 리드타임: {completed_cases['리드타임(일)'].median():.1f}일")
            print(f"    - 최댓값 리드타임: {completed_cases['리드타임(일)'].max():.0f}일")
        
        # 미출고 Case 분석
        pending_cases = self.df[self.df['상태'] == '미출고'].copy()
        if len(pending_cases) > 0:
            print(f"  📊 미출고 Case: {len(pending_cases)}건")
            print(f"    - 평균 체류일수: {pending_cases['현재체류일수'].mean():.1f}일")
            print(f"    - 중앙값 체류일수: {pending_cases['현재체류일수'].median():.1f}일")
            print(f"    - 최댓값 체류일수: {pending_cases['현재체류일수'].max():.0f}일")
        
        return {
            'completed': completed_cases,
            'pending': pending_cases,
            'no_inbound': self.df[self.df['상태'] == '미입고']
        }
    
    def analyze_by_warehouse(self):
        """
        창고별 리드타임 통계 분석 (출고완료 Case만)
        
        Returns:
            DataFrame: 창고별 리드타임 통계
        """
        print("\n🏭 창고별 리드타임 분석 중...")
        
        completed_cases = self.df[self.df['상태'] == '출고완료'].copy()
        
        if len(completed_cases) == 0:
            print("  ⚠️  출고완료된 Case가 없습니다.")
            return pd.DataFrame()
        
        warehouse_stats = completed_cases.groupby('초기창고')['리드타임(일)'].agg([
            'count', 'mean', 'median', 'std', 'min', 'max'
        ]).reset_index()
        warehouse_stats.columns = ['창고', '건수', '평균(일)', '중앙값(일)', '표준편차', '최솟값(일)', '최댓값(일)']
        
        # 소수점 정리
        numeric_cols = warehouse_stats.select_dtypes(include=[np.number]).columns
        warehouse_stats[numeric_cols] = warehouse_stats[numeric_cols].round(1)
        
        print(f"✅ 창고별 분석 완료: {len(warehouse_stats)}개 창고")
        return warehouse_stats
    
    def get_dead_stock_cases(self, threshold_days=90):
        """
        Dead Stock Case 목록 조회 (미출고 + 체류일수 임계값 초과)
        
        Args:
            threshold_days: 임계값 (일)
            
        Returns:
            DataFrame: Dead Stock Case 목록
        """
        print(f"\n⏰ Dead Stock Case 분석 (임계값: {threshold_days}일)...")
        
        dead_stock = self.df[
            (self.df['상태'] == '미출고') & 
            (self.df['현재체류일수'] >= threshold_days)
        ].copy()
        dead_stock = dead_stock.sort_values('현재체류일수', ascending=False)
        
        print(f"✅ Dead Stock Case: {len(dead_stock)}건")
        return dead_stock
    
    def get_long_leadtime_cases(self, threshold_days=90):
        """
        긴 리드타임 Case 목록 조회 (출고완료 Case만)
        
        Args:
            threshold_days: 임계값 (일)
            
        Returns:
            DataFrame: 긴 리드타임 Case 목록
        """
        print(f"\n⏰ 긴 리드타임 Case 분석 (임계값: {threshold_days}일)...")
        
        long_leadtime = self.df[
            (self.df['상태'] == '출고완료') & 
            (self.df['리드타임(일)'] >= threshold_days)
        ].copy()
        long_leadtime = long_leadtime.sort_values('리드타임(일)', ascending=False)
        
        print(f"✅ 긴 리드타임 Case: {len(long_leadtime)}건")
        return long_leadtime
    
    def generate_comprehensive_report(self, output_file=None):
        """
        종합 리드타임 분석 리포트 생성
        
        Args:
            output_file: 출력 파일 경로
            
        Returns:
            str: 생성된 파일 경로
        """
        print("\n📋 종합 리드타임 분석 리포트 생성 중...")
        
        if output_file is None:
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            output_dir = os.path.join(os.path.dirname(__file__), 'outputs')
            os.makedirs(output_dir, exist_ok=True)
            output_file = os.path.join(output_dir, f'종합_리드타임_분석_{timestamp}.xlsx')
        
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            # 1. 전체 Case 리드타임 데이터 (모든 Case 포함)
            print("  📋 전체 Case 리드타임 데이터 저장 중...")
            display_cols = ['Case No.', '초기창고', '입고일', '출고일', '리드타임(일)', '현재체류일수', '상태']
            self.df[display_cols].to_excel(writer, sheet_name='전체_Case_리드타임', index=False)
            
            # 2. 상태별 Case 목록
            print("  📋 상태별 Case 목록 저장 중...")
            status_analysis = self.analyze_by_status()
            
            # 출고완료 Case
            if len(status_analysis['completed']) > 0:
                completed_cols = ['Case No.', '초기창고', '입고일', '출고일', '리드타임(일)']
                status_analysis['completed'][completed_cols].to_excel(writer, sheet_name='출고완료_Case', index=False)
            
            # 미출고 Case
            if len(status_analysis['pending']) > 0:
                pending_cols = ['Case No.', '초기창고', '입고일', '현재체류일수']
                status_analysis['pending'][pending_cols].to_excel(writer, sheet_name='미출고_Case', index=False)
            
            # 미입고 Case
            if len(status_analysis['no_inbound']) > 0:
                no_inbound_cols = ['Case No.', '초기창고']
                status_analysis['no_inbound'][no_inbound_cols].to_excel(writer, sheet_name='미입고_Case', index=False)
            
            # 3. 창고별 통계
            print("  📋 창고별 통계 저장 중...")
            warehouse_stats = self.analyze_by_warehouse()
            if len(warehouse_stats) > 0:
                warehouse_stats.to_excel(writer, sheet_name='창고별_리드타임_통계', index=False)
            
            # 4. Dead Stock Case 목록
            print("  📋 Dead Stock Case 목록 저장 중...")
            dead_stock = self.get_dead_stock_cases(90)
            if len(dead_stock) > 0:
                dead_stock_cols = ['Case No.', '초기창고', '입고일', '현재체류일수']
                dead_stock[dead_stock_cols].to_excel(writer, sheet_name='DeadStock_90일+', index=False)
            
            # 5. 긴 리드타임 Case 목록
            print("  📋 긴 리드타임 Case 목록 저장 중...")
            long_leadtime = self.get_long_leadtime_cases(90)
            if len(long_leadtime) > 0:
                long_leadtime_cols = ['Case No.', '초기창고', '입고일', '출고일', '리드타임(일)']
                long_leadtime[long_leadtime_cols].to_excel(writer, sheet_name='긴리드타임_90일+', index=False)
            
            # 6. 종합 요약 정보
            print("  📋 종합 요약 정보 저장 중...")
            summary_data = {
                '분석 항목': [
                    '총 Case 수',
                    '출고완료 Case 수',
                    '미출고 Case 수',
                    '미입고 Case 수',
                    '출고완료 Case 평균 리드타임(일)',
                    '출고완료 Case 중앙값 리드타임(일)',
                    '출고완료 Case 최댓값 리드타임(일)',
                    '미출고 Case 평균 체류일수',
                    '미출고 Case 중앙값 체류일수',
                    '미출고 Case 최댓값 체류일수',
                    '90일 이상 리드타임 Case 수',
                    '90일 이상 체류 미출고 Case 수 (Dead Stock)',
                    '분석 창고 수'
                ],
                '값': [
                    len(self.df),
                    len(self.df[self.df['상태'] == '출고완료']),
                    len(self.df[self.df['상태'] == '미출고']),
                    len(self.df[self.df['상태'] == '미입고']),
                    round(self.df[self.df['상태'] == '출고완료']['리드타임(일)'].mean(), 1) if len(self.df[self.df['상태'] == '출고완료']) > 0 else 0,
                    round(self.df[self.df['상태'] == '출고완료']['리드타임(일)'].median(), 1) if len(self.df[self.df['상태'] == '출고완료']) > 0 else 0,
                    self.df[self.df['상태'] == '출고완료']['리드타임(일)'].max() if len(self.df[self.df['상태'] == '출고완료']) > 0 else 0,
                    round(self.df[self.df['상태'] == '미출고']['현재체류일수'].mean(), 1) if len(self.df[self.df['상태'] == '미출고']) > 0 else 0,
                    round(self.df[self.df['상태'] == '미출고']['현재체류일수'].median(), 1) if len(self.df[self.df['상태'] == '미출고']) > 0 else 0,
                    self.df[self.df['상태'] == '미출고']['현재체류일수'].max() if len(self.df[self.df['상태'] == '미출고']) > 0 else 0,
                    len(self.get_long_leadtime_cases(90)),
                    len(self.get_dead_stock_cases(90)),
                    len(self.df['초기창고'].dropna().unique())
                ]
            }
            summary_df = pd.DataFrame(summary_data)
            summary_df.to_excel(writer, sheet_name='종합_분석_요약', index=False)
        
        print(f"✅ 종합 리드타임 분석 리포트 생성 완료!")
        print(f"📄 파일명: {os.path.basename(output_file)}")
        print(f"📄 파일 크기: {os.path.getsize(output_file) / 1024:.1f} KB")
        
        return output_file

def main():
    """메인 실행 함수"""
    print("=== HVDC Warehouse 종합 리드타임 분석 시스템 ===")
    print(f"실행 시간: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    
    # 1. 데이터 파일 경로 설정
    current_dir = os.path.dirname(os.path.abspath(__file__))
    excel_path = os.path.join(current_dir, 'data', 'HVDC WAREHOUSE_HITACHI(HE).xlsx')
    
    if not os.path.exists(excel_path):
        print(f"❌ 오류: 데이터 파일을 찾을 수 없습니다: {excel_path}")
        return
    
    try:
        # 2. 종합 리드타임 분석기 초기화
        print(f"\n📁 데이터 파일: {excel_path}")
        analyzer = ComprehensiveLeadtimeAnalyzer(excel_path, sheet_name='CASE LIST')
        
        # 3. 종합 리드타임 계산
        df_comprehensive = analyzer.calculate_comprehensive_leadtime()
        
        # 4. 상세 분석
        print("\n📊 상세 분석 결과:")
        
        # 상태별 분석
        status_analysis = analyzer.analyze_by_status()
        
        # 창고별 분석
        warehouse_stats = analyzer.analyze_by_warehouse()
        if len(warehouse_stats) > 0:
            print(f"\n🏭 창고별 리드타임 통계 (출고완료 Case만):")
            print(warehouse_stats.to_string(index=False))
        
        # Dead Stock 분석
        dead_stock = analyzer.get_dead_stock_cases(90)
        print(f"\n⏰ Dead Stock Case (90일 이상 체류, 상위 10건):")
        if len(dead_stock) > 0:
            display_cols = ['Case No.', '초기창고', '입고일', '현재체류일수']
            print(dead_stock[display_cols].head(10).to_string(index=False))
        
        # 긴 리드타임 분석
        long_leadtime = analyzer.get_long_leadtime_cases(90)
        print(f"\n⏰ 긴 리드타임 Case (90일 이상, 상위 10건):")
        if len(long_leadtime) > 0:
            display_cols = ['Case No.', '초기창고', '입고일', '출고일', '리드타임(일)']
            print(long_leadtime[display_cols].head(10).to_string(index=False))
        
        # 5. 종합 엑셀 리포트 생성
        output_file = analyzer.generate_comprehensive_report()
        
        # 6. 엑셀 파일 자동 열기
        try:
            os.startfile(output_file)
            print(f"\n🔓 엑셀 파일을 자동으로 열었습니다.")
        except:
            print(f"\n💡 엑셀 파일을 수동으로 열어주세요: {output_file}")
        
        print(f"\n📋 생성된 시트 목록:")
        print("  - 전체_Case_리드타임: 모든 Case (출고완료/미출고/미입고 포함)")
        print("  - 출고완료_Case: 출고완료된 Case만")
        print("  - 미출고_Case: 아직 출고되지 않은 Case")
        print("  - 미입고_Case: 입고되지 않은 Case")
        print("  - 창고별_리드타임_통계: 창고별 평균/중앙값/최댓값")
        print("  - DeadStock_90일+: 90일 이상 체류 미출고 Case")
        print("  - 긴리드타임_90일+: 90일 이상 리드타임 Case")
        print("  - 종합_분석_요약: 전체 분석 요약 정보")
        
        return output_file
        
    except Exception as e:
        print(f"\n❌ 오류 발생: {str(e)}")
        import traceback
        traceback.print_exc()
        return None

if __name__ == "__main__":
    main() 