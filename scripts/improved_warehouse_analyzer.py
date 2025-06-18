import pandas as pd
from collections import Counter
from datetime import datetime
from pandas.tseries.offsets import MonthEnd

class ImprovedWarehouseAnalyzer:
    def __init__(self, excel_path, sheet_name='CASE LIST'):
        self.df = pd.read_excel(excel_path, sheet_name=sheet_name)
        self._preprocess()
        self._identify_columns()
        
    def _preprocess(self):
        """데이터 전처리"""
        # Case No. 정규화
        self.df['Case No.'] = self.df['Case No.'].astype(str)
        
        # 날짜 컬럼들을 datetime으로 변환
        date_columns = self.df.select_dtypes(include=['object']).columns
        for col in date_columns:
            if col != 'Case No.' and col != 'Site':
                self.df[col] = pd.to_datetime(self.df[col], errors='coerce')
    
    def _identify_columns(self):
        """입고/출고 컬럼 자동 식별"""
        inbound_keywords = ['DSV Outdoor', 'DSV Indoor', 'DSV Al Markaz', 'Hauler Indoor', 'DSV MZP', 'MOSB']
        outbound_keywords = ['DAS', 'MIR', 'SHU', 'AGI']
        
        self.inbound_cols = [col for col in self.df.columns if col in inbound_keywords]
        self.outbound_cols = [col for col in self.df.columns if col in outbound_keywords]
        
        print(f"입고 컬럼: {self.inbound_cols}")
        print(f"출고 컬럼: {self.outbound_cols}")
    
    def compute_monthly_inventory(self, start_date='2023-01-01', end_date='2025-12-31'):
        """개선된 월별 입출고/재고 계산 (정확한 로직)"""
        # 월별 집계를 위한 자료구조 초기화
        warehouse_stats = {loc: {'in': Counter(), 'out': Counter()} for loc in self.inbound_cols}
        site_stats = {site: Counter() for site in self.outbound_cols}
        
        # 데이터 행별 이벤트 처리
        for _, row in self.df.iterrows():
            # 해당 케이스의 모든 이벤트 수집 (입고: 창고명, 출고: 현장명)
            events = []
            for loc in self.inbound_cols:
                if pd.notna(row[loc]):  # 창고 입고 이벤트
                    events.append((pd.to_datetime(row[loc]), loc, 'warehouse_in'))
            for site in self.outbound_cols:
                if pd.notna(row[site]):  # 현장 출고 이벤트
                    events.append((pd.to_datetime(row[site]), site, 'site_out'))
            # 입출고 이벤트가 없다면 (미도착 품목), skip
            if not events:
                continue
            # 시간 순 정렬
            events.sort(key=lambda x: x[0])
            
            # 이벤트 순차 처리
            prev_loc = None  # 이전 이벤트 위치
            for date, loc, ev_type in events:
                month = date.strftime("%Y-%m")
                if ev_type == 'warehouse_in':
                    if prev_loc is None:
                        # 최초 외부->창고 입고
                        warehouse_stats[loc]['in'][month] += 1
                    else:
                        # 이전 이벤트가 존재 => 창고 간 이동으로 판단
                        # 이전 위치에서 출고, 새로운 loc로 입고
                        warehouse_stats[prev_loc]['out'][month] += 1
                        warehouse_stats[loc]['in'][month] += 1
                    prev_loc = loc  # 현재 위치를 이전 위치로 갱신
                elif ev_type == 'site_out':
                    # 현장 출고 이벤트 (최종 목적지 인도)
                    if prev_loc is not None:
                        # 마지막 머물던 창고에서 출고 처리
                        warehouse_stats[prev_loc]['out'][month] += 1
                    # 현장 입고량 집계
                    site_stats[loc][month] += 1
                    # 출고 후에는 더 이상 재고로 남지 않으므로 prev_loc 리셋
                    prev_loc = None
            # 루프 종료. prev_loc에 창고명이 남아있다면 (출고 안 된 케이스),
            # 해당 케이스는 현재 prev_loc 창고에 재고로 남아있음.
            # (별도 처리 불필요: 출고 이벤트 없었으므로 재고 계산 시 자동 반영됨)
        
        return warehouse_stats, site_stats
    
    def calculate_monthly_stock(self, warehouse_stats, site_stats, start_date='2023-01-01', end_date='2025-12-31'):
        """월별 재고 계산 (정확한 로직)"""
        # 월 범위 생성
        months = pd.date_range(start=start_date, end=end_date, freq='M')
        month_strs = [m.strftime("%Y-%m") for m in months]
        
        # 창고별 월별 재고 계산
        warehouse_stock = {}
        for warehouse in self.inbound_cols:
            stock_data = []
            current_stock = 0
            
            for month in month_strs:
                inbound = warehouse_stats[warehouse]['in'].get(month, 0)
                outbound = warehouse_stats[warehouse]['out'].get(month, 0)
                current_stock = current_stock + inbound - outbound
                
                stock_data.append({
                    '월': month,
                    '입고': inbound,
                    '출고': outbound,
                    '재고': current_stock
                })
            
            warehouse_stock[warehouse] = pd.DataFrame(stock_data)
        
        # 현장별 월별 누적 입고 계산
        site_stock = {}
        for site in self.outbound_cols:
            stock_data = []
            cumulative_stock = 0
            
            for month in month_strs:
                inbound = site_stats[site].get(month, 0)
                cumulative_stock += inbound
                
                stock_data.append({
                    '월': month,
                    '입고': inbound,
                    '누적재고': cumulative_stock
                })
            
            site_stock[site] = pd.DataFrame(stock_data)
        
        return warehouse_stock, site_stock
    
    def get_dead_stock_analysis(self, days=90):
        """Dead Stock 분석 (개선된 버전)"""
        today = pd.Timestamp(datetime.today())
        dead_stock_list = []
        
        for _, row in self.df.iterrows():
            # 해당 케이스의 모든 이벤트 수집
            events = []
            for loc in self.inbound_cols:
                if pd.notna(row[loc]):
                    events.append((pd.to_datetime(row[loc]), loc, 'warehouse_in'))
            for site in self.outbound_cols:
                if pd.notna(row[site]):
                    events.append((pd.to_datetime(row[site]), site, 'site_out'))
            
            if not events:
                continue
            
            # 시간 순 정렬
            events.sort(key=lambda x: x[0])
            
            # 마지막 이벤트 확인
            last_event_date, last_location, last_event_type = events[-1]
            
            # 마지막 이벤트가 창고 입고이고, 출고 이벤트가 없다면 Dead Stock 후보
            if last_event_type == 'warehouse_in':
                # 출고 이벤트가 있는지 확인
                has_outbound = any(ev[2] == 'site_out' for ev in events)
                
                if not has_outbound:
                    days_since_last_event = (today - last_event_date).days
                    if days_since_last_event > days:
                        dead_stock_list.append({
                            'Case No.': row['Case No.'],
                            '마지막입고일': last_event_date,
                            '마지막위치': last_location,
                            '입고후경과일': days_since_last_event,
                            'Site': row.get('Site', '')
                        })
        
        return pd.DataFrame(dead_stock_list)
    
    def generate_comprehensive_report(self, start_date='2023-01-01', end_date='2025-12-31'):
        """종합 리포트 생성"""
        print("=== HVDC Warehouse 월별 입출고/재고 분석 리포트 ===")
        print(f"분석 기간: {start_date} ~ {end_date}")
        print(f"총 케이스 수: {len(self.df)}")
        print()
        
        # 월별 입출고/재고 계산
        warehouse_stats, site_stats = self.compute_monthly_inventory(start_date, end_date)
        warehouse_stock, site_stock = self.calculate_monthly_stock(warehouse_stats, site_stats, start_date, end_date)
        
        # 창고별 요약
        print("=== 창고별 최근 12개월 요약 ===")
        for warehouse, stock_df in warehouse_stock.items():
            recent_12 = stock_df.tail(12)
            total_in = recent_12['입고'].sum()
            total_out = recent_12['출고'].sum()
            current_stock = recent_12['재고'].iloc[-1]
            
            print(f"{warehouse}:")
            print(f"  - 최근 12개월 입고: {total_in}건")
            print(f"  - 최근 12개월 출고: {total_out}건")
            print(f"  - 현재 재고: {current_stock}건")
            print()
        
        # 현장별 요약
        print("=== 현장별 최근 12개월 요약 ===")
        for site, stock_df in site_stock.items():
            recent_12 = stock_df.tail(12)
            total_in = recent_12['입고'].sum()
            current_stock = recent_12['누적재고'].iloc[-1]
            
            print(f"{site}:")
            print(f"  - 최근 12개월 입고: {total_in}건")
            print(f"  - 누적 재고: {current_stock}건")
            print()
        
        # Dead Stock 분석
        dead_stock_df = self.get_dead_stock_analysis(days=90)
        print(f"=== Dead Stock 분석 (90일 이상 미출고) ===")
        print(f"총 {len(dead_stock_df)}건")
        if len(dead_stock_df) > 0:
            print("상위 10건:")
            print(dead_stock_df.head(10)[['Case No.', '마지막위치', '입고후경과일']])
        print()
        
        return {
            'warehouse_stats': warehouse_stats,
            'site_stats': site_stats,
            'warehouse_stock': warehouse_stock,
            'site_stock': site_stock,
            'dead_stock': dead_stock_df
        } 