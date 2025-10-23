# 기존 Excel 파일의 Refresh 버튼 클릭 기능 구현

이 프로젝트는 **기존 Excel 파일의 Refresh 버튼을 그대로 유지**하면서 클릭 기능만 파이썬으로 구현한 것입니다. 원본 파일을 수정하지 않고 A1 셀의 기존 Refresh 버튼을 감지하여 데이터 새로고침을 수행합니다.

## 주요 특징

### ✅ **기존 파일 보존**
- 원본 Excel 파일의 Refresh 버튼을 그대로 유지
- A1 셀의 스타일, 폰트, 색상 등 모든 속성 보존
- 데이터 영역만 업데이트하고 버튼은 건드리지 않음

### ✅ **안전한 데이터 업데이트**
- 기존 버튼 상태 백업 및 복원 기능
- 데이터 영역만 선택적으로 업데이트
- 원본 파일 손상 방지

## 주요 기능

### 1. 기본 기존 버튼 핸들러 (`existing_button_handler.py`)
- 기존 Refresh 버튼 구조 분석 및 원본 상태 저장
- 안전한 데이터 새로고침 (버튼 보존)
- 클릭 시뮬레이션 및 데이터 업데이트

### 2. 고급 기존 버튼 시스템 (`advanced_existing_button_system.py`)
- 파일 모니터링을 통한 자동 버튼 클릭 감지
- 설정 파일 기반 동작 (`existing_button_config.json`)
- 다중 데이터 소스 지원
- 실시간 데이터 새로고침

## 설치 및 실행

### 1. 필요한 패키지 설치
```bash
pip install openpyxl watchdog
```

### 2. 기본 기존 버튼 핸들러 실행
```bash
python existing_button_handler.py
```

### 3. 고급 기존 버튼 시스템 실행
```bash
python advanced_existing_button_system.py
```

## 사용법

### 기본 사용법
1. **기존 Excel 파일 준비**: `excel_data/refresh_test_foreign.xlsx` 파일에 A1 셀에 Refresh 버튼이 있어야 함
2. **파이썬 스크립트 실행**: `python existing_button_handler.py`
3. **자동 처리**: 기존 버튼 분석 → 클릭 시뮬레이션 → 데이터 새로고침

### 고급 사용법
1. **파일 모니터링 모드**: Excel 파일 변경을 실시간 감지
2. **수동 시뮬레이션 모드**: 즉시 버튼 클릭 시뮬레이션 실행
3. **설정 파일 조정**: `existing_button_config.json`에서 동작 방식 설정

## 동작 원리

### 1. 기존 버튼 분석
```
Excel 파일 로드 → A1 셀 분석 → 원본 상태 백업 → 버튼 감지
```

### 2. 안전한 데이터 업데이트
```
버튼 클릭 감지 → 데이터 영역만 클리어 → 새 데이터 입력 → 파일 저장
```

### 3. 원본 보존
- A1 셀의 모든 속성 (값, 폰트, 색상, 테두리) 보존
- 데이터 영역 (B2부터)만 업데이트
- 원본 버튼 상태 복원 기능 제공

## 설정 파일 (`existing_button_config.json`)

```json
{
  "button_cell": "A1",
  "data_start_row": 2,
  "data_start_col": 2,
  "monitor_file_changes": true,
  "auto_refresh_interval": 0,
  "preserve_original_button": true,
  "data_sources": {
    "stock_data": {
      "enabled": true,
      "symbols": ["삼성전자", "SK하이닉스", "LG화학", "NAVER", "카카오"]
    },
    "market_data": {
      "enabled": true,
      "indices": ["KOSPI", "KOSDAQ", "S&P500"]
    }
  },
  "refresh_triggers": {
    "text_changes": ["refresh", "새로고침", "reload", "update"],
    "value_changes": true,
    "file_modifications": true
  }
}
```

## 파일 구조

```
quantiwise_control/
├── excel_data/
│   └── refresh_test_foreign.xlsx     # 기존 Refresh 버튼이 있는 Excel 파일
├── existing_button_handler.py        # 기본 기존 버튼 핸들러
├── advanced_existing_button_system.py # 고급 기존 버튼 시스템
├── existing_button_config.json        # 기존 버튼 시스템 설정 (자동 생성)
├── requirements.txt                   # 필요한 패키지
└── README.md                          # 사용법 설명서
```

## 실제 사용 시나리오

### 시나리오 1: 기존 파일 보존하며 새로고침
1. 기존 Excel 파일에 A1 셀에 Refresh 버튼이 있는 상태
2. `python existing_button_handler.py` 실행
3. 기존 버튼은 그대로 유지되고 데이터만 새로고침

### 시나리오 2: 실시간 모니터링
1. `python advanced_existing_button_system.py` 실행
2. 파일 모니터링 모드 선택
3. Excel에서 A1 셀의 "Refresh"를 "새로고침"으로 수동 변경
4. 자동으로 클릭 감지 및 데이터 새로고침 실행

### 시나리오 3: 안전한 백업 및 복원
1. 원본 버튼 상태 자동 백업
2. 데이터 새로고침 실행
3. 필요시 원본 버튼 상태 복원

## 안전 기능

### 1. 원본 보존
- 기존 버튼의 모든 속성 백업
- 데이터 영역만 선택적 업데이트
- 원본 상태 복원 기능

### 2. 오류 방지
- 파일 접근 권한 확인
- 백업 파일 생성 옵션
- 안전한 데이터 업데이트

### 3. 복원 기능
- 원본 버튼 상태 자동 복원
- 설정 파일 기반 복원 옵션
- 수동 복원 기능

## 확장 방법

### 실제 API 연동
```python
def fetch_real_data(self, data_source):
    import requests
    
    if data_source == "stock_data":
        api_url = "https://api.finance.com/stocks"
        response = requests.get(api_url)
        return response.json()
```

### 데이터베이스 연동
```python
def fetch_real_data(self, data_source):
    import sqlite3
    
    conn = sqlite3.connect('financial_data.db')
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM stock_prices WHERE date = ?", (datetime.now().date(),))
    return cursor.fetchall()
```

## 문제 해결

### 기존 버튼이 감지되지 않는 경우
1. A1 셀에 정확히 "Refresh" 텍스트가 있는지 확인
2. 설정 파일의 `button_cell` 설정 확인
3. 파일 권한 확인

### 데이터 업데이트가 되지 않는 경우
1. Excel 파일이 읽기 전용이 아닌지 확인
2. 파일이 다른 프로그램에서 열려있지 않은지 확인
3. 데이터 소스 설정이 활성화되어 있는지 확인

### 원본 버튼이 변경된 경우
1. 설정 파일에서 `preserve_original_button: true` 확인
2. 원본 상태 복원 기능 사용
3. 백업 파일에서 복원

## 주의사항

1. **파일 접근**: Excel 파일이 다른 프로그램에서 열려있으면 업데이트가 실패할 수 있습니다.
2. **권한**: 파일 쓰기 권한이 필요합니다.
3. **백업**: 중요한 데이터는 미리 백업하세요.
4. **원본 보존**: 설정에서 `preserve_original_button: true`로 설정하여 원본 버튼을 보존하세요.

## 라이선스

이 프로젝트는 MIT 라이선스 하에 배포됩니다.
