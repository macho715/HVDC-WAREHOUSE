# Changelog

모든 주요 변경사항이 이 파일에 기록됩니다.

## [1.3.0] - 2025-06-19

### Changed
- **프로젝트 구조 정리**
  - 불필요한 중복 파일 삭제
  - 존재하지 않는 모듈 import 제거
  - outputs 폴더 중복 파일 정리
- **main.py 수정**
  - 존재하지 않는 `warehouse_analyzer`, `warehouse_monthly_analyzer` import 제거
  - `improved_warehouse_analyzer`만 사용하도록 단순화

### Removed
- `leadtime_analyzer.py` (comprehensive_leadtime_analyzer.py로 대체)
- `create_dashboard.py` (create_advanced_dashboard.py로 대체)
- `filtered_analysis.py` (기능이 다른 모듈에 통합됨)
- `scripts/preprocessing.py` (사용되지 않음)
- `scripts/visualization.py` (사용되지 않음)
- outputs 폴더의 중복 파일들

## [1.2.0] - 2025-06-19

### Added
- **피벗 테이블 기능 추가**
  - 월별 데이터를 창고 분류별로 피벗하여 분석
  - 공급사별 비교 분석 가능
  - In/Out/Stock 메트릭별 집계
- **영어 버전 완성**
  - 모든 컬럼명과 값이 영어로 표시
  - 국제 표준에 맞는 분석 리포트
- **자동 파일 열기 기능**
  - 분석 완료 후 자동으로 엑셀 파일 열기
  - Windows/Mac/Linux 호환

### Changed
- 파일 경로 수정: 상대 경로 문제 해결
- Python 실행 방식 개선: `py` 명령어 사용

### Fixed
- Microsoft Store 버전 Python 런처 문제 해결
- 파일 경로 오류 수정

## [1.1.0] - 2025-06-19

### Added
- **데드스톡 분석 기능**
  - 90일 이상 창고에 보관된 재고 분석
  - 공급사별 케이스별 상세 정보
- **창고 분류 시스템**
  - Indoor/Outdoor/Dangerous 창고 구분
  - 분류별 집계 분석

### Changed
- 다중 시트 출력 형식으로 변경
- 엑셀 포맷팅 개선

## [1.0.0] - 2025-06-19

### Added
- **기본 창고 분석 시스템**
  - 4개 공급사 데이터 통합 처리
  - 월별 재고 현황 분석
  - 창고별 입고/출고/재고 추적
  - 현장별 누적 입고 추적
- **통합 엑셀 리포트 생성**
  - Consolidated_Status 시트
  - Overall_Supplier_Summary 시트
  - Warehouse_Stock_Summary 시트

### Technical
- Python 기반 데이터 처리 시스템
- pandas, openpyxl, xlsxwriter 라이브러리 활용
- 자동화된 엑셀 파일 생성 및 포맷팅 