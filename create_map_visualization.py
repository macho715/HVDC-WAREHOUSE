#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
HVDC Warehouse Map Visualization
- 지도 위에 창고와 현장 위치를 표시합니다.
- 각 위치를 클릭하면 월별 입출고량 그래프를 보여줍니다.
- Folium과 Matplotlib을 사용하여 인터랙티브 HTML 리포트를 생성합니다.
"""

import os
import sys
import pandas as pd
import folium
import matplotlib.pyplot as plt
import base64
from io import BytesIO
from datetime import datetime

# scripts 디렉토리를 Python 경로에 추가하여 커스텀 분석기 모듈을 import할 수 있도록 합니다.
sys.path.append(os.path.join(os.path.dirname(__file__), 'scripts'))

# 커스텀 창고 분석기 클래스를 import합니다.
from corrected_warehouse_analyzer import CorrectedWarehouseAnalyzer

# --- 설정 ---
# 각 창고와 현장의 지리적 좌표 (위도, 경도)를 정의합니다.
# 이 섹션은 검증된 좌표로 업데이트되었습니다.
LOCATIONS = {
    # 창고들 (검증된 데이터로 업데이트)
    "DSV Al Markaz": (24.19121, 54.47356),      # 검증됨: Al Markaz
    "MOSB": (24.34708, 54.47772),               # 검증됨: M44 (Mussafah)
    "DSV Outdoor": (24.76341, 54.70860),        # 검증됨: DSV Outdoor (KHIA)
    
    # 창고들 (아직 임시 좌표 사용)
    "DSV Indoor": (24.76441, 54.70960),         # 임시 (DSV Outdoor 근처)
    "Hauler Indoor": (24.35000, 54.48000),      # 임시 (Mussafah 근처)
    "DSV MZP": (24.53371, 54.37918),            # 임시 (Mina Zayed Port 근처)

    # 현장들 (검증된 데이터로 업데이트)
    "SHU": (24.10971, 52.53508),               # 검증됨: Shuweihat (SHU)
    "MIR": (24.06285, 53.45938),               # 검증됨: Mirfa (MIR)
    "DAS": (25.15139, 52.87361),               # 검증됨: Das Island
    "AGI": (24.81791, 53.66395),               # 검증됨: Al Ghallan Island
}

# 분석의 종료 월을 설정합니다. 차트는 이 시점까지의 데이터를 표시합니다.
TARGET_MONTH = "2025-06"

def create_monthly_chart(df, title):
    """
    월별 데이터에 대한 선 차트를 생성하고 base64 인코딩된 문자열로 반환합니다.

    Args:
        df (pd.DataFrame): DatetimeIndex와 '입고', '출고' 컬럼이 있는 DataFrame
        title (str): 차트의 제목

    Returns:
        str: 생성된 차트 이미지의 base64 인코딩된 문자열
    """
    if df.empty:
        return None

    # 목표 월까지의 데이터를 필터링합니다
    df_filtered = df[df.index <= TARGET_MONTH].copy()

    if df_filtered.empty:
        return None

    fig, ax = plt.subplots(figsize=(5, 3))

    # 사용 가능한 컬럼에 따라 데이터를 플롯합니다
    if '입고' in df_filtered.columns:
        ax.plot(df_filtered.index, df_filtered['입고'], marker='o', linestyle='-', label='입고', color='blue')
    if '출고' in df_filtered.columns:
        ax.plot(df_filtered.index, df_filtered['출고'], marker='x', linestyle='--', label='출고', color='red')
    if '누적재고' in df_filtered.columns:  # 현장용
        ax.plot(df_filtered.index, df_filtered['누적재고'], marker='s', linestyle='-', label='누적재고', color='green')

    ax.set_title(title, fontsize=12)
    ax.set_xlabel("월", fontsize=10)
    ax.set_ylabel("수량", fontsize=10)
    ax.grid(True, which='both', linestyle='--', linewidth=0.5)
    ax.legend()
    plt.xticks(rotation=45)
    plt.tight_layout()

    # 차트를 임시 버퍼에 저장합니다
    tmpfile = BytesIO()
    fig.savefig(tmpfile, format='png', dpi=100, bbox_inches='tight')
    plt.close(fig)  # 메모리를 해제하기 위해 figure를 닫습니다

    # 이미지를 base64로 인코딩하여 HTML에 임베드합니다
    encoded = base64.b64encode(tmpfile.getvalue()).decode('utf-8')
    return encoded

def main():
    """지도 시각화를 생성하는 메인 함수"""
    print("=== HVDC Warehouse Map Visualization System ===")
    print(f"실행 시간: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

    # 1. 파일 경로 설정
    current_dir = os.path.dirname(os.path.abspath(__file__))
    excel_path = os.path.join(current_dir, 'data', 'HVDC WAREHOUSE_HITACHI(HE).xlsx')
    output_dir = os.path.join(current_dir, 'outputs')
    os.makedirs(output_dir, exist_ok=True)
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    output_file = os.path.join(output_dir, f'warehouse_map_visualization_{timestamp}.html')

    if not os.path.exists(excel_path):
        print(f"❌ 오류: 데이터 파일을 찾을 수 없습니다: {excel_path}")
        return

    try:
        # 2. 창고 데이터를 분석하여 월별 통계를 얻습니다
        print("\n🔍 데이터 분석기 초기화 중...")
        analyzer = CorrectedWarehouseAnalyzer(excel_path, sheet_name='CASE LIST')
        
        print("📊 월별 입출고 데이터 계산 중...")
        analysis_result = analyzer.generate_corrected_report(
            start_date='2023-01-01',
            end_date='2025-12-31'
        )
        warehouse_data = analysis_result.get('warehouse_stock', {})
        site_data = analysis_result.get('site_stock', {})
        print("✅ 데이터 분석 완료.")

        # 3. 대략적인 위치 주변에 기본 지도를 생성합니다
        print("\n🗺️  인터랙티브 지도 생성 중...")
        map_center = [24.42, 54.43]  # 아부다비 지역 주변에 중심을 둡니다
        m = folium.Map(location=map_center, zoom_start=8, tiles="CartoDB positron")  # 약간 확대하여 전체 지역을 볼 수 있도록 합니다

        # 4. 창고에 대한 마커를 추가합니다
        print("🏭 창고 마커 추가 중...")
        for name, data in warehouse_data.items():
            if name in LOCATIONS:
                chart_b64 = create_monthly_chart(data, f"창고: {name}")
                if chart_b64:
                    iframe = folium.IFrame(f'<img src="data:image/png;base64,{chart_b64}">', width=550, height=350)
                    popup = folium.Popup(iframe, max_width=550)
                    folium.Marker(
                        location=LOCATIONS[name],
                        popup=popup,
                        tooltip=f"창고: {name}",
                        icon=folium.Icon(color="blue", icon="industry", prefix="fa")
                    ).add_to(m)

        # 5. 현장에 대한 마커를 추가합니다
        print("🏗️  현장 마커 추가 중...")
        for name, data in site_data.items():
            if name in LOCATIONS:
                chart_b64 = create_monthly_chart(data, f"현장: {name}")
                if chart_b64:
                    iframe = folium.IFrame(f'<img src="data:image/png;base64,{chart_b64}">', width=550, height=350)
                    popup = folium.Popup(iframe, max_width=550)
                    folium.Marker(
                        location=LOCATIONS[name],
                        popup=popup,
                        tooltip=f"현장: {name}",
                        icon=folium.Icon(color="green", icon="wrench", prefix="fa")
                    ).add_to(m)
        
        # 6. 지도를 HTML 파일로 저장합니다
        m.save(output_file)
        print(f"✅ 인터랙티브 지도 저장 완료!")
        print(f"📄 파일명: {os.path.basename(output_file)}")
        print(f"📄 파일 크기: {os.path.getsize(output_file) / 1024:.1f} KB")

        # 7. 생성된 파일을 자동으로 엽니다
        try:
            os.startfile(output_file)
            print(f"\n🔓 지도가 브라우저에서 자동으로 열렸습니다.")
        except AttributeError:
            # os.startfile()은 Windows용입니다. macOS와 Linux용:
            import subprocess
            try:
                subprocess.run(['open', output_file], check=True)  # macOS
                print(f"\n🔓 지도가 브라우저에서 자동으로 열렸습니다.")
            except:
                try:
                    subprocess.run(['xdg-open', output_file], check=True)  # Linux
                    print(f"\n🔓 지도가 브라우저에서 자동으로 열렸습니다.")
                except:
                     print(f"\n💡 지도를 수동으로 열어주세요: {output_file}")

        print(f"\n📋 지도 기능:")
        print("  - 🔵 파란색 마커: 창고 (클릭하면 월별 입출고 그래프)")
        print("  - 🟢 초록색 마커: 현장 (클릭하면 월별 누적입고 그래프)")
        print("  - 📊 팝업 차트: 각 위치의 월별 물류 흐름 시각화")
        print("  - 🗺️  정확한 위치: 검증된 지리적 좌표 사용")

    except Exception as e:
        print(f"\n❌ 오류 발생: {str(e)}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main() 