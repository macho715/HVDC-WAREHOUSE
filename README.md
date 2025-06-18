# HVDC Warehouse Analytics

## 프로젝트 개요
- HVDC 자재 입출고/재고/리드타임 등 실무 자동화 분석 시스템
- 창고별/현장별 월별 입출고/재고 자동 집계 및 엑셀 리포트 생성

## 주요 기능
- **창고별 월별 입출고/재고 집계**: DSV Indoor, DSV Outdoor, MOSB 등
- **현장별(Site) 월별 입고/누적재고**: DAS, MIR, SHU, AGI 등
- **Dead Stock 분석**: 90일 이상 미출고 자재 자동 식별
- **KPI 자동 산출**: Site별 도달률, 평균 리드타임
- **회전율 분석**: 창고별 월별 회전율 계산
- **엑셀 자동 저장**: 모든 분석 결과를 하나의 엑셀 파일로 저장

## 폴더 구조
```
warehouse_analytics/
├── data/           # 원본 데이터 (HVDC WAREHOUSE_HITACHI(HE).xlsx)
├── scripts/        # 분석 코드
│   ├── warehouse_analyzer.py      # 기본 분석 클래스
│   ├── warehouse_monthly_analyzer.py  # 월별 집계 클래스
│   └── main.py                    # 실행 파일
├── outputs/        # 결과 리포트 (엑셀 파일)
├── README.md       # 프로젝트 설명
└── requirements.txt # 패키지 목록
```

## 실행 방법
1. **패키지 설치**
   ```bash
   pip install -r requirements.txt
   ```

2. **데이터 파일 준비**
   - `data/` 폴더에 `HVDC WAREHOUSE_HITACHI(HE).xlsx` 파일 배치
   - 시트명: 'CASE LIST'

3. **분석 실행**
   ```bash
   python scripts/main.py
   ```

## 출력 결과
- **터미널 출력**: 월별 요약, Dead Stock, KPI, 회전율 등
- **엑셀 파일**: `outputs/월별_창고_현장_입출고재고_집계.xlsx`
  - 창고별 시트: 각 창고의 월별 입출고/재고
  - 현장별 시트: 각 Site의 월별 입고/누적재고
  - 기타 시트: 전체 요약, Dead Stock, KPI, 회전율

## 기술 스택
- Python 3.x
- pandas: 데이터 처리 및 분석
- matplotlib/seaborn: 시각화
- openpyxl: 엑셀 파일 처리

## 개발자
- GitHub: [macho715](https://github.com/macho715)
