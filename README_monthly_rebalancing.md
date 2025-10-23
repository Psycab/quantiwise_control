# DeepSearch 외인수급Top20 지수 (PR) 매달 리밸런싱 시스템

## 📋 개요
DeepSearch Net Foreign BuyingTop20 Index PR의 매달 리밸런싱을 자동화하는 시스템입니다.

## 🚀 주요 기능

### 1. 자동 파일 복사 및 날짜 업데이트
- 이전 월의 `deepsearch_foreign_buying_top20_index_raw_data_YYYYMMDD.xlsx` 파일을 다음 달 마지막 날로 복사
- 파일명의 날짜를 다음 달 마지막 날로 자동 변경
- Excel 시트 내의 B5, B6 셀 날짜를 새로운 날짜로 업데이트

### 2. 자동 분석 실행
- 업데이트된 파일로 DeepSearch 외인수급Top20 지수 분석 자동 실행
- 결과 파일을 `deepsearch_foreign_buying_top20_index_result_YYYYMMDD.xlsx` 형식으로 생성

## 📁 파일 구조

```
excel_data/
├── deepsearch_foreign_buying_top20_index_raw_data_20241031.xlsx    # 10월 데이터
├── deepsearch_foreign_buying_top20_index_raw_data_20250831.xlsx    # 8월 데이터  
├── deepsearch_foreign_buying_top20_index_raw_data_20250930.xlsx    # 9월 데이터 (자동 생성)
└── deepsearch_foreign_buying_top20_index_result_20250930.xlsx     # 9월 분석 결과
```

## 🔧 사용법

### PowerShell에서 실행
```powershell
# 가상환경 활성화 후 실행
.venv/Scripts/python.exe -c "import sys; sys.stdout.reconfigure(encoding='utf-8'); exec(open('monthly_rebalancing_scheduler.py', encoding='utf-8').read())"
```

### 실행 과정
1. **사용자 입력**: 기존 파일 날짜와 새 파일 날짜를 각각 입력 (YYYY-MM-DD 형식)
2. **기존 파일 검색**: 입력된 기존 날짜에 해당하는 raw_data 파일 찾기
3. **파일 복사**: 새 날짜 형식으로 새 파일 생성
4. **날짜 업데이트**: Excel 시트 내 B5, B6 셀의 날짜를 새로운 날짜로 변경
5. **Excel 열기**: 복사된 파일을 Excel로 열기
6. **수동 refresh**: Quantiwise의 refresh 버튼을 클릭하여 데이터 새로고침
7. **파일 저장**: 업데이트된 데이터를 저장하고 Excel 닫기
8. **분석 실행**: 새로고침된 파일로 DeepSearch 시스템 분석 실행
9. **결과 생성**: 분석 결과를 새로운 결과 파일로 저장

## 📊 실행 결과 예시

```
================================================================================
DeepSearch 외인수급Top20 지수 (PR) 매달 리밸런싱 시작
영문명: DeepSearch Net Foreign BuyingTop20 Index PR
================================================================================

📅 날짜 정보를 입력해주세요.
형식: YYYY-MM-DD

1️⃣ 기존 raw_data 파일의 날짜를 입력하세요.
예: 2024-10-31
기존 파일 날짜: 2024-10-31

2️⃣ 새로 생성할 raw_data 파일의 날짜를 입력하세요.
예: 2024-11-30
새 파일 날짜: 2024-11-30

✅ 입력된 정보:
   기존 파일 날짜: 2024년 10월 31일
   새 파일 날짜: 2024년 11월 30일

이 정보로 진행하시겠습니까? (y/n): y

📁 기존 파일(20241031) 검색 중...
✅ 기존 파일 발견: deepsearch_foreign_buying_top20_index_raw_data_20241031.xlsx
📋 새 파일로 복사 중...
✅ 파일 복사 완료: deepsearch_foreign_buying_top20_index_raw_data_20241031.xlsx → deepsearch_foreign_buying_top20_index_raw_data_20241130.xlsx
📅 Excel 파일 내 날짜 업데이트 중...
✅ Excel 파일 날짜 업데이트 완료: deepsearch_foreign_buying_top20_index_raw_data_20241130.xlsx
📂 Excel 파일 열기 및 데이터 새로고침 중...
📂 Excel 파일 열기: deepsearch_foreign_buying_top20_index_raw_data_20241130.xlsx
⚠️  수동 작업이 필요합니다:
   1. Excel 파일이 열리면 각 시트의 데이터를 확인하세요
   2. Quantiwise의 refresh 버튼을 클릭하여 데이터를 새로고침하세요
   3. 모든 시트의 데이터가 업데이트되었는지 확인하세요
   4. 파일을 저장하고 Excel을 닫으세요
   5. 완료되면 Enter 키를 눌러 계속하세요...

📋 위 작업을 완료한 후 Enter 키를 눌러 계속하세요...
✅ Excel 파일 처리 완료: deepsearch_foreign_buying_top20_index_raw_data_20241130.xlsx
🔍 데이터 분석 실행 중...
✅ 분석 완료: deepsearch_foreign_buying_top20_index_result_20241130.xlsx
================================================================================
DeepSearch 외인수급Top20 지수 (PR) 매달 리밸런싱 완료!
- 기존 파일: deepsearch_foreign_buying_top20_index_raw_data_20241031.xlsx
- 새 파일: deepsearch_foreign_buying_top20_index_raw_data_20241130.xlsx
- 결과 파일: deepsearch_foreign_buying_top20_index_result_20241130.xlsx
- 실행 시간: 15.52초
================================================================================
```

## 📅 월별 리밸런싱 스케줄

- **사용자 지정 날짜**: 기준이 되는 마지막 날을 직접 입력
- **자동 날짜 계산**: 입력된 날짜의 다음 달 마지막 날로 자동 설정
- **윤년 처리**: 자동으로 윤년의 2월 29일까지 고려
- **확인 절차**: 입력된 날짜와 계산된 다음 달 날짜를 확인 후 진행

## 💡 사용자 입력 가이드

### 날짜 입력 형식
- **형식**: `YYYY-MM-DD`
- **예시**: `2024-10-31`, `2024-12-31`
- **기존 파일**: 복사할 원본 파일의 날짜
- **새 파일**: 생성할 새 파일의 날짜
- **주의**: 기존 파일이 반드시 존재해야 함

### 확인 절차
1. 기존 파일 날짜 입력
2. 새 파일 날짜 입력
3. 입력된 정보 확인 및 승인
4. 기존 파일 존재 여부 확인
5. 프로세스 진행

## 🤖 자동화 모드 선택

### 1. 수동 모드
- 사용자가 직접 Excel을 열고 Quantiwise refresh 버튼 클릭
- 가장 안정적이고 확실한 방법

### 2. 자동화 모드 (권장)
- **퀀티와이즈 담당자님 가이드 기반**
- `win32com.client`를 사용하여 Excel 자동 제어
- 각 시트의 A1 셀에 있는 refresh 버튼을 자동으로 클릭
- 완전 자동화 가능

## ⚙️ 시스템 요구사항

### 기본 요구사항
- Python 3.7+
- openpyxl 라이브러리
- PowerShell 환경
- UTF-8 인코딩 지원

### 자동화 모드 추가 요구사항
- `pywin32` 라이브러리
- 설치: `pip install pywin32`
- Excel이 매크로를 지원하는 환경

## 🧹 자동 파일 정리 기능

### 에러 발생 시 자동 정리
- **작업 중 에러 발생**: 새로 생성된 raw_data 파일 자동 삭제
- **부분 완료 상태 방지**: 에러로 인한 불완전한 파일 남김 방지
- **디스크 공간 절약**: 불필요한 파일 자동 정리

### 정리 대상
- 새로 생성된 `deepsearch_foreign_buying_top20_index_raw_data_YYYYMMDD.xlsx` 파일
- 에러 발생 시점에 따라 생성된 파일만 삭제
- 기존 원본 파일은 보호

## 🔍 주요 클래스 및 메서드

### `MonthlyRebalancingScheduler`
- `get_user_input_dates()`: 사용자로부터 날짜 입력 받기
- `copy_file_with_custom_date()`: 사용자 지정 날짜로 파일 복사
- `update_dates_in_excel()`: Excel 파일 내 날짜 업데이트
- `open_excel_and_refresh_data()`: Excel 파일 열기 및 Quantiwise refresh
- `run_analysis()`: 분석 실행
- `run_monthly_rebalancing()`: 전체 프로세스 실행 (에러 시 파일 정리 포함)

## 📝 주의사항

1. **파일 형식**: 반드시 `deepsearch_foreign_buying_top20_index_raw_data_YYYYMMDD.xlsx` 형식 준수
2. **시트 구조**: eps_sheet, market_cap_sheet, foreign_sheet 시트 존재 필요
3. **날짜 셀**: B5, B6 셀에 날짜 정보 저장 필요
4. **백업**: 중요한 데이터는 사전에 백업 권장

## 🚨 오류 처리

- 파일이 없을 경우: "❌ 기존 raw_data 파일을 찾을 수 없습니다."
- 날짜 형식 오류: "❌ 유효한 날짜 형식의 파일을 찾을 수 없습니다."
- 복사 실패: "❌ 파일 복사 중 오류 발생"
- 분석 실패: "❌ 분석 실행 중 오류 발생"
- **에러 시 파일 정리**: "🧹 에러 발생으로 인한 파일 정리: 파일명 삭제 완료"
