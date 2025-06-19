# Warehouse Analytics System

## 개요
HVDC 창고 데이터를 분석하여 월별 재고 현황, 공급사별 요약, 창고별 재고 현황, 피벗 테이블 분석, 데드스톡 분석을 제공하는 시스템입니다.

## 파일 구조
```
warehouse_analytics/
├── data/                                    # 원본 엑셀 데이터 파일
│   ├── HVDC WAREHOUSE_HITACHI(HE).xlsx
│   ├── HVDC WAREHOUSE_HITACHI(HE_LOCAL).xlsx
│   ├── HVDC WAREHOUSE_HITACHI(HE-0214,0252).xlsx
│   └── HVDC WAREHOUSE_SIMENSE(SIM).xlsx
├── scripts/                                 # 핵심 분석 모듈
│   ├── main.py                             # 메인 실행 파일
│   ├── improved_warehouse_analyzer.py       # 개선된 분석기 (정확한 이벤트 추적)
│   └── corrected_warehouse_analyzer.py      # 수정된 분석기
├── outputs/                                 # 생성된 분석 결과 파일
├── case_based_analysis_english_pivot.py    # 메인 분석 스크립트 (영어 버전 + 피벗 테이블)
├── comprehensive_analysis_report.py        # 종합 분석 리포트
├── comprehensive_leadtime_analyzer.py      # 종합 리드타임 분석
├── anomaly_detection.py                    # 이상 감지 시스템
├── create_advanced_dashboard.py            # 고급 대시보드 생성
├── create_map_visualization.py             # 지도 시각화
├── generate_practical_report.py            # 실무용 리포트 생성
└── README.md                               # 이 파일
```

## 주요 기능

### 1. 데이터 처리
- 4개 공급사 데이터 통합 처리
  - HITACHI
  - HITACHI_LOCAL
  - HITACHI_LOT
  - SIEMENS
- 창고별 입고/출고/재고 추적
- 현장별 누적 입고 추적

### 2. 창고 분류
- **Indoor**: DSV Indoor, Hauler Indoor, DSV Al Markaz, AAA Storage, DHL WH
- **Outdoor**: DSV Outdoor, DSV MZP, MOSB
- **Dangerous**: AAA Storage

### 3. 생성되는 엑셀 시트

#### Consolidated_Status
- 월별 상세 재고 현황
- 공급사별 창고별 입고/출고/재고 데이터
- 현장별 입고/누적 입고 데이터

#### Overall_Supplier_Summary
- 공급사별 전체 요약
- 총 창고 입고/출고/재고
- 총 현장 누적 입고

#### Warehouse_Stock_Summary
- 창고별 상세 재고 현황
- 창고 분류(Indoor/Outdoor/Dangerous) 포함
- 입고/출고/현재 재고

#### Pivoted_Monthly_Summary (NEW)
- 월별 피벗 테이블 분석
- 창고 분류별 공급사별 데이터 집계
- In/Out/Stock 메트릭별 분석

#### DeadStock_Analysis
- 90일 이상 창고에 보관된 데드스톡 분석
- 공급사별 케이스별 상세 정보

## 실행 방법

### 1. 환경 설정
```bash
# 필요한 라이브러리 설치
pip install pandas openpyxl xlsxwriter
```

### 2. 메인 스크립트 실행
```bash
cd warehouse_analytics/scripts
python main.py
```

### 3. 특화 분석 스크립트 실행
```bash
cd warehouse_analytics

# 영어 버전 + 피벗 테이블 분석
python case_based_analysis_english_pivot.py

# 종합 분석 리포트
python comprehensive_analysis_report.py

# 리드타임 분석
python comprehensive_leadtime_analyzer.py

# 이상 감지
python anomaly_detection.py

# 고급 대시보드
python create_advanced_dashboard.py

# 지도 시각화
python create_map_visualization.py

# 실무용 리포트
python generate_practical_report.py
```

### 4. 결과 확인
- `outputs/` 폴더에 타임스탬프가 포함된 엑셀 파일 생성
- 파일명: `Consolidated_Inventory_Report_YYYYMMDD_HHMMSS.xlsx`

## 설정 옵션

### 파일 경로 설정
```python
file_map = {
    'HITACHI': 'data/HVDC WAREHOUSE_HITACHI(HE).xlsx',
    'HITACHI_LOCAL': 'data/HVDC WAREHOUSE_HITACHI(HE_LOCAL).xlsx',
    'HITACHI_LOT': 'data/HVDC WAREHOUSE_HITACHI(HE-0214,0252).xlsx',
    'SIEMENS': 'data/HVDC WAREHOUSE_SIMENSE(SIM).xlsx',
}
```

### 창고 컬럼 설정
```python
warehouse_cols_map = {
    'HITACHI': ['DSV Outdoor', 'DSV Indoor', 'DSV Al Markaz', 'Hauler Indoor', 'DSV MZP', 'MOSB'],
    'HITACHI_LOCAL': ['DSV Outdoor', 'DSV Al Markaz', 'DSV MZP', 'MOSB'],
    'HITACHI_LOT': ['DSV Indoor', 'DHL WH', 'DSV Al Markaz', 'AAA Storage'],
    'SIEMENS': ['DSV Outdoor', 'DSV Indoor', 'DSV Al Markaz', 'MOSB', 'AAA Storage'],
}
```

### 현장 컬럼 설정
```python
site_cols = ['DAS', 'MIR', 'SHU', 'AGI']
```

### 분석 기준 설정
```python
DEADSTOCK_DAYS = 90  # 데드스톡 기준 (일)
target_month = "2025-06"  # 분석 대상 월
```

## 주요 개선사항

### 1. 피벗 테이블 기능 추가
- 월별 데이터를 창고 분류별로 피벗하여 분석
- 공급사별 비교 분석 가능
- In/Out/Stock 메트릭별 집계

### 2. 영어 버전 완성
- 모든 컬럼명과 값이 영어로 표시
- 국제 표준에 맞는 분석 리포트

### 3. 자동 파일 열기
- 분석 완료 후 자동으로 엑셀 파일 열기
- Windows/Mac/Linux 호환

### 4. 정확한 이벤트 추적
- 입고→창고간이동→출고의 정확한 이벤트 추적
- 누적 재고 계산의 정확성 향상

## 문제 해결

### 파일 경로 오류
- 상대 경로 문제 해결: `warehouse_analytics/data/` → `data/`
- 현재 작업 디렉토리 기준으로 경로 설정

### Python 실행 문제
- Microsoft Store 버전 Python 런처 문제 해결
- `py` 명령어 사용으로 해결

## 향후 개선 계획

1. **대시보드 기능**
   - 웹 기반 대시보드 추가
   - 실시간 데이터 시각화

2. **자동화 기능**
   - 스케줄러를 통한 자동 분석
   - 이메일 알림 기능

3. **고급 분석**
   - 예측 분석 기능
   - 이상치 탐지 기능

## 문의사항
추가 기능이나 수정 사항이 필요하시면 언제든 말씀해 주세요.
