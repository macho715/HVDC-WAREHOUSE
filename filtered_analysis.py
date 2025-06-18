#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
HVDC Warehouse 조건별 필터링 및 월별 변화 재계산 시스템
- 특정 창고/현장/자재군/보관형태별 필터링
- 필터링된 데이터의 월별 입출고/재고 재계산
- 실무에서 바로 사용 가능한 엑셀 리포트
"""

import os
import sys
import pandas as pd
import numpy as np
from datetime import datetime
from pandas.tseries.offsets import MonthEnd
import warnings
warnings.filterwarnings('ignore')

class FilteredAnalysis:
    """조건별 필터링 및 월별 변화 재계산 클래스"""
    
    def __init__(self, excel_path, sheet_name='CASE LIST'):
        """
        필터링 분석기 초기화
        
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
        
        # 3) 초기 입고 창고명 식별
        self.df['초기창고'] = None
        for idx, row in self.df.iterrows():
            warehouse_dates = row[self.warehouse_cols]
            valid_dates = warehouse_dates.dropna()
            if len(valid_dates) > 0:
                min_date_idx = valid_dates.idxmin()
                self.df.at[idx, '초기창고'] = min_date_idx
        
        # 4) 최종 출고 현장 식별
        self.df['최종출고현장'] = None
        for idx, row in self.df.iterrows():
            site_dates = row[self.site_cols]
            valid_dates = site_dates.dropna()
            if len(valid_dates) > 0:
                max_date_idx = valid_dates.idxmax()
                self.df.at[idx, '최종출고현장'] = max_date_idx
        
        # 5) Storage Type 분류 (창고명 기반)
        self.df['Storage_Type'] = self.df['초기창고'].apply(self._classify_storage_type)
        
        # 6) Case 상태 분류
        self.df['상태'] = self.df.apply(self._classify_case_status, axis=1)
        
        print("✅ 데이터 전처리 완료")
    
    def _classify_storage_type(self, warehouse):
        """창고명 기반 Storage Type 분류"""
        if pd.isna(warehouse):
            return "Unknown"
        elif 'Indoor' in str(warehouse):
            return "Indoor"
        elif 'Outdoor' in str(warehouse):
            return "Outdoor"
        else:
            return "Other"
    
    def _classify_case_status(self, row):
        """Case 상태 분류"""
        if pd.isna(row['입고일']):
            return "미입고"
        elif pd.isna(row['출고일']):
            return "미출고"
        else:
            return "출고완료"
    
    def filter_by_conditions(self, filters=None):
        """
        조건별 데이터 필터링
        
        Args:
            filters: 필터 조건 딕셔너리
                {
                    'warehouse': '창고명',
                    'site': '현장명', 
                    'storage_type': 'Indoor/Outdoor/Other',
                    'material_category': '자재군명',
                    'status': '출고완료/미출고/미입고'
                }
        
        Returns:
            DataFrame: 필터링된 데이터
        """
        if filters is None:
            filters = {}
        
        print(f"\n🔍 조건별 필터링 중...")
        print(f"  📋 필터 조건: {filters}")
        
        filtered_df = self.df.copy()
        
        # 창고별 필터링
        if 'warehouse' in filters and filters['warehouse']:
            warehouse = filters['warehouse']
            if warehouse in self.warehouse_cols:
                # 특정 창고에 입고된 Case
                filtered_df = filtered_df[filtered_df[warehouse].notna()]
                print(f"    ✅ 창고 필터: {warehouse} (입고된 Case)")
            else:
                # 초기 입고 창고 기준
                filtered_df = filtered_df[filtered_df['초기창고'] == warehouse]
                print(f"    ✅ 창고 필터: {warehouse} (초기 입고 창고)")
        
        # 현장별 필터링
        if 'site' in filters and filters['site']:
            site = filters['site']
            if site in self.site_cols:
                # 특정 현장으로 출고된 Case
                filtered_df = filtered_df[filtered_df[site].notna()]
                print(f"    ✅ 현장 필터: {site} (출고된 Case)")
            else:
                # 최종 출고 현장 기준
                filtered_df = filtered_df[filtered_df['최종출고현장'] == site]
                print(f"    ✅ 현장 필터: {site} (최종 출고 현장)")
        
        # Storage Type 필터링
        if 'storage_type' in filters and filters['storage_type']:
            storage_type = filters['storage_type']
            filtered_df = filtered_df[filtered_df['Storage_Type'] == storage_type]
            print(f"    ✅ Storage Type 필터: {storage_type}")
        
        # 자재군 필터링
        if 'material_category' in filters and filters['material_category']:
            material = filters['material_category']
            if 'Material Category' in filtered_df.columns:
                filtered_df = filtered_df[filtered_df['Material Category'] == material]
                print(f"    ✅ 자재군 필터: {material}")
            else:
                print(f"    ⚠️  Material Category 컬럼이 없어 자재군 필터링을 건너뜁니다.")
        
        # 상태별 필터링
        if 'status' in filters and filters['status']:
            status = filters['status']
            filtered_df = filtered_df[filtered_df['상태'] == status]
            print(f"    ✅ 상태 필터: {status}")
        
        print(f"✅ 필터링 완료: {len(filtered_df)}건 (원본: {len(self.df)}건)")
        return filtered_df
    
    def calculate_monthly_trends(self, filtered_df):
        """
        필터링된 데이터의 월별 변화 계산
        
        Args:
            filtered_df: 필터링된 DataFrame
            
        Returns:
            dict: 월별 변화 데이터
        """
        print(f"\n📊 월별 변화 계산 중...")
        
        if len(filtered_df) == 0:
            print("  ⚠️  필터링된 데이터가 없습니다.")
            return {}
        
        # 1) 입고월/출고월 계산
        filtered_df['입고월'] = filtered_df['입고일'].dt.to_period('M').dt.to_timestamp() + MonthEnd(0)
        filtered_df['출고월'] = filtered_df['출고일'].dt.to_period('M').dt.to_timestamp() + MonthEnd(0)
        
        # 2) 월별 입고/출고 건수 계산
        monthly_in = filtered_df.groupby('입고월').size()
        monthly_out = filtered_df.groupby('출고월').size()
        
        # 3) 모든 월 범위 생성 (입고월과 출고월 합집합)
        all_months = pd.concat([pd.Series(monthly_in.index), pd.Series(monthly_out.index)]).unique()
        all_months = pd.DatetimeIndex(all_months).sort_values()
        
        # 4) 월별 입고/출고 건수 정리
        monthly_in_filled = monthly_in.reindex(all_months, fill_value=0)
        monthly_out_filled = monthly_out.reindex(all_months, fill_value=0)
        
        # 5) 누적 입출고 차이를 통해 월말 재고 계산
        cumulative_in = monthly_in_filled.cumsum()
        cumulative_out = monthly_out_filled.cumsum()
        monthly_stock = cumulative_in - cumulative_out
        
        # 6) 결과를 데이터프레임으로 정리
        monthly_trend = pd.DataFrame({
            '월': all_months,
            '입고': monthly_in_filled.values,
            '출고': monthly_out_filled.values,
            '재고': monthly_stock.values,
            '누적입고': cumulative_in.values,
            '누적출고': cumulative_out.values
        })
        
        # 7) 창고별 월별 재고 계산
        warehouse_monthly = {}
        for warehouse in filtered_df['초기창고'].dropna().unique():
            warehouse_df = filtered_df[filtered_df['초기창고'] == warehouse]
            warehouse_trend = self._calculate_warehouse_monthly(warehouse_df)
            warehouse_monthly[warehouse] = warehouse_trend
        
        # 8) 현장별 월별 누적입고 계산
        site_monthly = {}
        for site in filtered_df['최종출고현장'].dropna().unique():
            site_df = filtered_df[filtered_df['최종출고현장'] == site]
            site_trend = self._calculate_site_monthly(site_df)
            site_monthly[site] = site_trend
        
        print(f"✅ 월별 변화 계산 완료!")
        print(f"  📊 분석 기간: {monthly_trend['월'].min()} ~ {monthly_trend['월'].max()}")
        print(f"  📊 총 월 수: {len(monthly_trend)}개월")
        print(f"  📊 창고별 분석: {len(warehouse_monthly)}개 창고")
        print(f"  📊 현장별 분석: {len(site_monthly)}개 현장")
        
        return {
            'monthly_trend': monthly_trend,
            'warehouse_monthly': warehouse_monthly,
            'site_monthly': site_monthly,
            'filtered_data': filtered_df
        }
    
    def _calculate_warehouse_monthly(self, warehouse_df):
        """창고별 월별 재고 계산"""
        if len(warehouse_df) == 0:
            return pd.DataFrame()
        
        warehouse_df['입고월'] = warehouse_df['입고일'].dt.to_period('M').dt.to_timestamp() + MonthEnd(0)
        warehouse_df['출고월'] = warehouse_df['출고일'].dt.to_period('M').dt.to_timestamp() + MonthEnd(0)
        
        monthly_in = warehouse_df.groupby('입고월').size()
        monthly_out = warehouse_df.groupby('출고월').size()
        
        all_months = pd.concat([pd.Series(monthly_in.index), pd.Series(monthly_out.index)]).unique()
        all_months = pd.DatetimeIndex(all_months).sort_values()
        
        monthly_in_filled = monthly_in.reindex(all_months, fill_value=0)
        monthly_out_filled = monthly_out.reindex(all_months, fill_value=0)
        
        cumulative_in = monthly_in_filled.cumsum()
        cumulative_out = monthly_out_filled.cumsum()
        monthly_stock = cumulative_in - cumulative_out
        
        return pd.DataFrame({
            '월': all_months,
            '입고': monthly_in_filled.values,
            '출고': monthly_out_filled.values,
            '재고': monthly_stock.values
        })
    
    def _calculate_site_monthly(self, site_df):
        """현장별 월별 누적입고 계산"""
        if len(site_df) == 0:
            return pd.DataFrame()
        
        site_df['출고월'] = site_df['출고일'].dt.to_period('M').dt.to_timestamp() + MonthEnd(0)
        
        monthly_out = site_df.groupby('출고월').size()
        all_months = pd.DatetimeIndex(monthly_out.index).sort_values()
        
        monthly_out_filled = monthly_out.reindex(all_months, fill_value=0)
        cumulative_out = monthly_out_filled.cumsum()
        
        return pd.DataFrame({
            '월': all_months,
            '입고': monthly_out_filled.values,  # 현장은 출고=입고 개념
            '누적재고': cumulative_out.values
        })
    
    def generate_filtered_report(self, filters=None, output_file=None):
        """
        필터링된 데이터 분석 리포트 생성
        
        Args:
            filters: 필터 조건
            output_file: 출력 파일 경로
            
        Returns:
            str: 생성된 파일 경로
        """
        print(f"\n📋 필터링된 데이터 분석 리포트 생성 중...")
        
        if output_file is None:
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            output_dir = os.path.join(os.path.dirname(__file__), 'outputs')
            os.makedirs(output_dir, exist_ok=True)
            
            # 필터 조건을 파일명에 반영
            filter_name = "_".join([f"{k}_{v}" for k, v in (filters or {}).items() if v])
            if filter_name:
                output_file = os.path.join(output_dir, f'필터링_분석_{filter_name}_{timestamp}.xlsx')
            else:
                output_file = os.path.join(output_dir, f'필터링_분석_전체_{timestamp}.xlsx')
        
        # 1. 데이터 필터링
        filtered_df = self.filter_by_conditions(filters)
        
        # 2. 월별 변화 계산
        monthly_data = self.calculate_monthly_trends(filtered_df)
        
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            # 1. 필터링된 전체 데이터
            print("  📋 필터링된 전체 데이터 저장 중...")
            display_cols = ['Case No.', '초기창고', '최종출고현장', 'Storage_Type', '입고일', '출고일', '상태']
            if 'Material Category' in filtered_df.columns:
                display_cols.insert(4, 'Material Category')
            filtered_df[display_cols].to_excel(writer, sheet_name='필터링된_전체데이터', index=False)
            
            # 2. 월별 전체 추이
            if 'monthly_trend' in monthly_data and len(monthly_data['monthly_trend']) > 0:
                print("  📋 월별 전체 추이 저장 중...")
                monthly_data['monthly_trend'].to_excel(writer, sheet_name='월별_전체추이', index=False)
            
            # 3. 창고별 월별 추이
            if 'warehouse_monthly' in monthly_data:
                print("  📋 창고별 월별 추이 저장 중...")
                for warehouse, trend_df in monthly_data['warehouse_monthly'].items():
                    if len(trend_df) > 0:
                        sheet_name = f'창고_{warehouse}'[:31]  # Excel 시트명 제한
                        trend_df.to_excel(writer, sheet_name=sheet_name, index=False)
            
            # 4. 현장별 월별 추이
            if 'site_monthly' in monthly_data:
                print("  📋 현장별 월별 추이 저장 중...")
                for site, trend_df in monthly_data['site_monthly'].items():
                    if len(trend_df) > 0:
                        sheet_name = f'Site_{site}'[:31]  # Excel 시트명 제한
                        trend_df.to_excel(writer, sheet_name=sheet_name, index=False)
            
            # 5. 필터링 요약 정보
            print("  📋 필터링 요약 정보 저장 중...")
            summary_data = {
                '분석 항목': [
                    '원본 데이터 총 Case 수',
                    '필터링된 Case 수',
                    '필터링 비율 (%)',
                    '분석 기간 시작',
                    '분석 기간 종료',
                    '분석 월 수',
                    '창고별 분석 수',
                    '현장별 분석 수'
                ],
                '값': [
                    len(self.df),
                    len(filtered_df),
                    round(len(filtered_df) / len(self.df) * 100, 1),
                    monthly_data.get('monthly_trend', pd.DataFrame())['월'].min() if 'monthly_trend' in monthly_data and len(monthly_data['monthly_trend']) > 0 else "N/A",
                    monthly_data.get('monthly_trend', pd.DataFrame())['월'].max() if 'monthly_trend' in monthly_data and len(monthly_data['monthly_trend']) > 0 else "N/A",
                    len(monthly_data.get('monthly_trend', pd.DataFrame())),
                    len(monthly_data.get('warehouse_monthly', {})),
                    len(monthly_data.get('site_monthly', {}))
                ]
            }
            summary_df = pd.DataFrame(summary_data)
            summary_df.to_excel(writer, sheet_name='필터링_요약', index=False)
        
        print(f"✅ 필터링된 데이터 분석 리포트 생성 완료!")
        print(f"📄 파일명: {os.path.basename(output_file)}")
        print(f"📄 파일 크기: {os.path.getsize(output_file) / 1024:.1f} KB")
        
        return output_file

def main():
    """메인 실행 함수"""
    print("=== HVDC Warehouse 조건별 필터링 및 월별 변화 재계산 시스템 ===")
    print(f"실행 시간: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    
    # 1. 데이터 파일 경로 설정
    current_dir = os.path.dirname(os.path.abspath(__file__))
    excel_path = os.path.join(current_dir, 'data', 'HVDC WAREHOUSE_HITACHI(HE).xlsx')
    
    if not os.path.exists(excel_path):
        print(f"❌ 오류: 데이터 파일을 찾을 수 없습니다: {excel_path}")
        return
    
    try:
        # 2. 필터링 분석기 초기화
        print(f"\n📁 데이터 파일: {excel_path}")
        analyzer = FilteredAnalysis(excel_path, sheet_name='CASE LIST')
        
        # 3. 다양한 필터링 조건 예시
        filter_examples = [
            {
                'name': 'DSV_Outdoor_창고',
                'filters': {'warehouse': 'DSV Outdoor'}
            },
            {
                'name': '실내보관_자재',
                'filters': {'storage_type': 'Indoor'}
            },
            {
                'name': 'DAS_현장_출고',
                'filters': {'site': 'DAS'}
            },
            {
                'name': 'DSV_Indoor_미출고',
                'filters': {'warehouse': 'DSV Indoor', 'status': '미출고'}
            }
        ]
        
        # 4. 각 필터링 조건별로 분석 실행
        for example in filter_examples:
            print(f"\n🔍 {example['name']} 분석 중...")
            
            # 필터링된 데이터 분석
            output_file = analyzer.generate_filtered_report(
                filters=example['filters']
            )
            
            print(f"✅ {example['name']} 분석 완료: {os.path.basename(output_file)}")
        
        # 5. 전체 데이터 분석 (필터 없음)
        print(f"\n🔍 전체 데이터 분석 중...")
        output_file = analyzer.generate_filtered_report()
        print(f"✅ 전체 데이터 분석 완료: {os.path.basename(output_file)}")
        
        # 6. 엑셀 파일 자동 열기
        try:
            os.startfile(output_file)
            print(f"\n🔓 최종 엑셀 파일을 자동으로 열었습니다.")
        except:
            print(f"\n💡 엑셀 파일을 수동으로 열어주세요: {output_file}")
        
        print(f"\n📋 생성된 분석 결과:")
        print("  - DSV Outdoor 창고별 분석")
        print("  - 실내보관 자재별 분석")
        print("  - DAS 현장 출고별 분석")
        print("  - DSV Indoor 미출고별 분석")
        print("  - 전체 데이터 종합 분석")
        
        return output_file
        
    except Exception as e:
        print(f"\n❌ 오류 발생: {str(e)}")
        import traceback
        traceback.print_exc()
        return None

if __name__ == "__main__":
    main() 