# HVDC Warehouse Analytics System

정확한 Case별 이벤트 기반 창고 분석 시스템

## 🎯 프로젝트 개요

HVDC Warehouse의 물자 입출고, 재고, KPI, 리드타임을 Excel 데이터 기반으로 자동 분석하는 시스템입니다.

## ✨ 주요 기능

### 1. 정확한 Case별 이벤트 기반 재고 집계
- 각 Case의 실제 이동 순서를 추적
- 창고 간 이동 시 이중 카운트 방지
- 정확한 월별 입출고/재고 산출

### 2. 월별 창고/현장 분석
- 창고별 월별 입출고/재고 집계
- 현장별 월별 누적 입고량 집계
- Dead Stock 분석 (90일 이상 미출고)

### 3. 자동 엑셀 리포트 생성
- 타임스탬프가 포함된 고유 파일명
- 다중 시트 구성 (창고별, 현장별, 요약, Dead Stock)
- 자동 파일 열기 기능

## 🚀 사용법

### 1. 기본 실행
```bash
python case_based_inventory.py
```

### 2. 개선된 분석 실행
```bash
python run_improved_analysis.py
```

### 3. 정확한 분석 실행
```bash
python run_corrected_analysis.py
```

## 📁 프로젝트 구조

```
warehouse_analytics/
├── data/                          # 원본 데이터 파일
│   └── HVDC WAREHOUSE_HITACHI(HE).xlsx
├── scripts/                       # 핵심 분석 모듈
│   ├── improved_warehouse_analyzer.py
│   ├── corrected_warehouse_analyzer.py
│   └── ...
├── outputs/                       # 생성된 엑셀 파일들
├── case_based_inventory.py        # 메인 실행 스크립트
├── run_improved_analysis.py       # 개선된 분석 실행
├── run_corrected_analysis.py      # 정확한 분석 실행
├── debug_inventory.py             # 재고 오류 진단
└── README.md
```

## 📊 분석 결과

### 창고별 최종 재고 (Case별 이벤트 기반)
- DSV Outdoor: 826건
- DSV Indoor: 414건
- DSV Al Markaz: 812건
- Hauler Indoor: 392건
- DSV MZP: 10건
- MOSB: 43건

### 현장별 최종 누적입고
- DAS: 678건
- MIR: 753건
- SHU: 1221건
- AGI: 34건

## 🔧 기술 스택

- **Python 3.x**
- **pandas**: 데이터 처리 및 분석
- **openpyxl**: 엑셀 파일 생성
- **datetime**: 날짜/시간 처리

## 📋 요구사항

```bash
pip install -r requirements.txt
```

## 🎯 핵심 개선사항

1. **정확한 재고 산출**: Case별 실제 마지막 위치 기준
2. **이중 카운트 방지**: 창고 간 이동 시 중복 집계 제거
3. **자동화**: 엑셀 파일 자동 생성 및 열기
4. **검증 가능**: 디버깅 도구 포함

## 📝 라이선스

이 프로젝트는 내부 사용을 위한 것입니다.
