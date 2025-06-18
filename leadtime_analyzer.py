#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
HVDC Warehouse 리드타임 분석 시스템
- Case별 입고일부터 출고일까지의 리드타임 계산
- 창고별/자재군별 리드타임 통계 분석
- 평균/중앙값/최댓값 등 통계 산출
"""

import os
import sys
import pandas as pd
import numpy as np
from datetime import datetime
import warnings
warnings.filterwarnings('ignore')

# scripts 폴더의 모듈들을 import하기 위한 경로 설정
sys.path.append(os.path.join(os.path.dirname(__file__), 'scripts'))

class LeadtimeAnalyzer:
    """리드타임 분석 클래스"""
    
    def __init__(self, excel_path, sheet_name='CASE LIST'):
        """
        리드타임 분석기 초기화
        
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
    
    def calculate_leadtime(self):
        """
        리드타임 계산
        
        Returns:
            DataFrame: 리드타임이 계산된 DataFrame
        """
        print("\n📊 리드타임 계산 중...")
        
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
        
        # 4) 리드타임(일) 계산
        print("  ⏱️  리드타임 계산 중...")
        self.df['리드타임(일)'] = (self.df['출고일'] - self.df['입고일']).dt.days
        
        # 5) 초기 입고 창고명 식별 (최초 입고일을 제공한 창고)
        print("  🏭 초기 입고 창고 식별 중...")
        # 안전한 초기 창고 식별
        self.df['초기창고'] = None
        for idx, row in self.df.iterrows():
            warehouse_dates = row[self.warehouse_cols]
            valid_dates = warehouse_dates.dropna()
            if len(valid_dates) > 0:
                min_date_idx = valid_dates.idxmin()
                self.df.at[idx, '초기창고'] = min_date_idx
        
        # 6) 유효한 리드타임만 필터링 (입고일과 출고일이 모두 있는 경우)
        valid_mask = self.df['입고일'].notna() & self.df['출고일'].notna()
        self.df_valid = self.df[valid_mask].copy()
        
        print(f"✅ 리드타임 계산 완료!")
        print(f"  📊 총 Case 수: {len(self.df)}")
        print(f"  📊 유효한 Case 수: {len(self.df_valid)}")
        
        if len(self.df_valid) > 0:
            print(f"  📊 리드타임 평균: {self.df_valid['리드타임(일)'].mean():.1f}일")
            print(f"  📊 리드타임 중앙값: {self.df_valid['리드타임(일)'].median():.1f}일")
            print(f"  📊 리드타임 최댓값: {self.df_valid['리드타임(일)'].max():.0f}일")
        else:
            print("  ⚠️  유효한 리드타임 데이터가 없습니다.")
        
        return self.df_valid
    
    def analyze_by_warehouse(self):
        """
        창고별 리드타임 통계 분석
        
        Returns:
            DataFrame: 창고별 리드타임 통계
        """
        print("\n🏭 창고별 리드타임 분석 중...")
        
        if 'Material Category' in self.df_valid.columns:
            # 자재군별로도 그룹화
            warehouse_stats = self.df_valid.groupby(['초기창고', 'Material Category'])['리드타임(일)'].agg([
                'count', 'mean', 'median', 'std', 'min', 'max'
            ]).reset_index()
            warehouse_stats.columns = ['창고', '자재군', '건수', '평균(일)', '중앙값(일)', '표준편차', '최솟값(일)', '최댓값(일)']
        else:
            # 창고별로만 그룹화
            warehouse_stats = self.df_valid.groupby('초기창고')['리드타임(일)'].agg([
                'count', 'mean', 'median', 'std', 'min', 'max'
            ]).reset_index()
            warehouse_stats.columns = ['창고', '건수', '평균(일)', '중앙값(일)', '표준편차', '최솟값(일)', '최댓값(일)']
        
        # 소수점 정리
        numeric_cols = warehouse_stats.select_dtypes(include=[np.number]).columns
        warehouse_stats[numeric_cols] = warehouse_stats[numeric_cols].round(1)
        
        print(f"✅ 창고별 분석 완료: {len(warehouse_stats)}개 그룹")
        return warehouse_stats
    
    def analyze_by_material(self):
        """
        자재군별 리드타임 통계 분석
        
        Returns:
            DataFrame: 자재군별 리드타임 통계
        """
        if 'Material Category' not in self.df_valid.columns:
            print("⚠️  Material Category 컬럼이 없어 자재군별 분석을 건너뜁니다.")
            return None
        
        print("\n📦 자재군별 리드타임 분석 중...")
        
        material_stats = self.df_valid.groupby('Material Category')['리드타임(일)'].agg([
            'count', 'mean', 'median', 'std', 'min', 'max'
        ]).reset_index()
        material_stats.columns = ['자재군', '건수', '평균(일)', '중앙값(일)', '표준편차', '최솟값(일)', '최댓값(일)']
        
        # 소수점 정리
        numeric_cols = material_stats.select_dtypes(include=[np.number]).columns
        material_stats[numeric_cols] = material_stats[numeric_cols].round(1)
        
        print(f"✅ 자재군별 분석 완료: {len(material_stats)}개 그룹")
        return material_stats
    
    def get_long_leadtime_cases(self, threshold_days=90):
        """
        긴 리드타임 Case 목록 조회
        
        Args:
            threshold_days: 임계값 (일)
            
        Returns:
            DataFrame: 긴 리드타임 Case 목록
        """
        print(f"\n⏰ 긴 리드타임 Case 분석 (임계값: {threshold_days}일)...")
        
        long_leadtime = self.df_valid[self.df_valid['리드타임(일)'] >= threshold_days].copy()
        long_leadtime = long_leadtime.sort_values('리드타임(일)', ascending=False)
        
        print(f"✅ 긴 리드타임 Case: {len(long_leadtime)}건")
        return long_leadtime
    
    def generate_report(self, output_file=None):
        """
        리드타임 분석 리포트 생성
        
        Args:
            output_file: 출력 파일 경로
            
        Returns:
            str: 생성된 파일 경로
        """
        print("\n📋 리드타임 분석 리포트 생성 중...")
        
        if output_file is None:
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            output_dir = os.path.join(os.path.dirname(__file__), 'outputs')
            os.makedirs(output_dir, exist_ok=True)
            output_file = os.path.join(output_dir, f'리드타임_분석_{timestamp}.xlsx')
        
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            # 1. 전체 Case 리드타임 데이터
            print("  📋 전체 Case 리드타임 데이터 저장 중...")
            self.df_valid.to_excel(writer, sheet_name='전체_Case_리드타임', index=False)
            
            # 2. 창고별 통계
            print("  📋 창고별 통계 저장 중...")
            warehouse_stats = self.analyze_by_warehouse()
            warehouse_stats.to_excel(writer, sheet_name='창고별_리드타임_통계', index=False)
            
            # 3. 자재군별 통계
            print("  📋 자재군별 통계 저장 중...")
            material_stats = self.analyze_by_material()
            if material_stats is not None:
                material_stats.to_excel(writer, sheet_name='자재군별_리드타임_통계', index=False)
            
            # 4. 긴 리드타임 Case 목록
            print("  📋 긴 리드타임 Case 목록 저장 중...")
            long_leadtime = self.get_long_leadtime_cases(90)
            long_leadtime.to_excel(writer, sheet_name='긴리드타임_90일+', index=False)
            
            # 5. 요약 정보
            print("  📋 요약 정보 저장 중...")
            summary_data = {
                '분석 항목': [
                    '총 Case 수',
                    '유효한 Case 수',
                    '평균 리드타임(일)',
                    '중앙값 리드타임(일)',
                    '최댓값 리드타임(일)',
                    '90일 이상 리드타임 Case 수',
                    '분석 창고 수',
                    '분석 자재군 수'
                ],
                '값': [
                    len(self.df),
                    len(self.df_valid),
                    round(self.df_valid['리드타임(일)'].mean(), 1),
                    round(self.df_valid['리드타임(일)'].median(), 1),
                    self.df_valid['리드타임(일)'].max(),
                    len(self.get_long_leadtime_cases(90)),
                    len(self.df_valid['초기창고'].unique()),
                    len(self.df_valid['Material Category'].unique()) if 'Material Category' in self.df_valid.columns else 0
                ]
            }
            summary_df = pd.DataFrame(summary_data)
            summary_df.to_excel(writer, sheet_name='분석_요약', index=False)
        
        print(f"✅ 리드타임 분석 리포트 생성 완료!")
        print(f"📄 파일명: {os.path.basename(output_file)}")
        print(f"📄 파일 크기: {os.path.getsize(output_file) / 1024:.1f} KB")
        
        return output_file

def main():
    """메인 실행 함수"""
    print("=== HVDC Warehouse 리드타임 분석 시스템 ===")
    print(f"실행 시간: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    
    # 1. 데이터 파일 경로 설정
    current_dir = os.path.dirname(os.path.abspath(__file__))
    excel_path = os.path.join(current_dir, 'data', 'HVDC WAREHOUSE_HITACHI(HE).xlsx')
    
    if not os.path.exists(excel_path):
        print(f"❌ 오류: 데이터 파일을 찾을 수 없습니다: {excel_path}")
        return
    
    try:
        # 2. 리드타임 분석기 초기화
        print(f"\n📁 데이터 파일: {excel_path}")
        analyzer = LeadtimeAnalyzer(excel_path, sheet_name='CASE LIST')
        
        # 3. 리드타임 계산
        df_valid = analyzer.calculate_leadtime()
        
        # 4. 상세 분석
        print("\n📊 상세 분석 결과:")
        
        # 창고별 분석
        warehouse_stats = analyzer.analyze_by_warehouse()
        print(f"\n🏭 창고별 리드타임 통계:")
        print(warehouse_stats.to_string(index=False))
        
        # 자재군별 분석
        material_stats = analyzer.analyze_by_material()
        if material_stats is not None:
            print(f"\n📦 자재군별 리드타임 통계:")
            print(material_stats.to_string(index=False))
        
        # 긴 리드타임 Case 분석
        long_leadtime = analyzer.get_long_leadtime_cases(90)
        print(f"\n⏰ 90일 이상 리드타임 Case (상위 10건):")
        if len(long_leadtime) > 0:
            display_cols = ['Case No.', '초기창고', '입고일', '출고일', '리드타임(일)']
            if 'Material Category' in long_leadtime.columns:
                display_cols.insert(2, 'Material Category')
            
            print(long_leadtime[display_cols].head(10).to_string(index=False))
        
        # 5. 엑셀 리포트 생성
        output_file = analyzer.generate_report()
        
        # 6. 엑셀 파일 자동 열기
        try:
            os.startfile(output_file)
            print(f"\n🔓 엑셀 파일을 자동으로 열었습니다.")
        except:
            print(f"\n💡 엑셀 파일을 수동으로 열어주세요: {output_file}")
        
        print(f"\n📋 생성된 시트 목록:")
        print("  - 전체_Case_리드타임: 모든 Case의 리드타임 데이터")
        print("  - 창고별_리드타임_통계: 창고별 평균/중앙값/최댓값")
        print("  - 자재군별_리드타임_통계: 자재군별 통계 (Material Category 컬럼이 있는 경우)")
        print("  - 긴리드타임_90일+: 90일 이상 리드타임 Case 목록")
        print("  - 분석_요약: 전체 분석 요약 정보")
        
        return output_file
        
    except Exception as e:
        print(f"\n❌ 오류 발생: {str(e)}")
        import traceback
        traceback.print_exc()
        return None

if __name__ == "__main__":
    main() 