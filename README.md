# Quantiwise Excel Add-in Refresh 기능 구현

이 프로젝트는 Quantiwise Excel Add-in의 Refresh 기능을 파이썬으로 구현한 것입니다. **A1 셀에 있는 실제 Refresh 버튼**을 감지하고 클릭 시 데이터가 자동으로 재조회됩니다.

## 주요 기능

### 1. 기본 버튼 클릭 시뮬레이션 (`button_click_simulator.py`)
- A1 셀의 실제 Refresh 버튼 구조 분석
- 버튼 클릭 시뮬레이션 및 데이터 새로고침
- 버튼 상태 업데이트

### 2. 고급 버튼 시스템 (`advanced_button_system.py`)
- 파일 모니터링을 통한 자동 버튼 클릭 감지
- 설정 파일 기반 동작 (`button_config.json`)
- 버튼 상태 변화 감지 및 처리
- 실시간 데이터 새로고침

### 3. 기본 기능 (`quantiwise_refresh.py`)
- 텍스트 기반 Refresh 감지
- 시뮬레이션 데이터로 Excel 시트 업데이트

## 설치 및 실행

### 1. 필요한 패키지 설치
```bash
pip install openpyxl watchdog
```

### 2. 실제 Refresh 버튼이 있는 Excel 파일 생성
```bash
python create_button_excel.py
```

### 3. 버튼 클릭 시뮬레이션 실행
```bash
python button_click_simulator.py
```

### 4. 고급 버튼 시스템 실행
```bash
python advanced_button_system.py
```

## 사용법

### 실제 Refresh 버튼 사용법
1. **Excel 파일 생성**: `create_button_excel.py`로 실제 버튼이 있는 Excel 파일 생성
2. **버튼 분석**: A1 셀의 버튼 구조와 스타일 자동 분석
3. **클릭 시뮬레이션**: 파이썬 스크립트로 버튼 클릭 시뮬레이션
4. **자동 새로고침**: 클릭 감지 시 데이터 자동 업데이트

### 파일 모니터링 모드
```bash
python advanced_button_system.py
# 선택: 1 (파일 모니터링 모드)
```
- Excel 파일 변경을 실시간 모니터링
- 버튼 상태 변화 자동 감지
- 수동으로 Excel에서 버튼을 수정하여 클릭 시뮬레이션

### 수동 시뮬레이션 모드
```bash
python advanced_button_system.py
# 선택: 2 (수동 시뮬레이션 모드)
```
- 자동으로 버튼을 클릭된 상태로 변경
- 즉시 데이터 새로고침 실행

## 버튼 동작 원리

### 1. 버튼 상태 감지
- **정상 상태**: "Refresh" 텍스트
- **클릭된 상태**: "Refresh ✓" 또는 배경색 변경
- **상태 변화 감지**: 정상 → 클릭된 상태 변화 시 새로고침 실행

### 2. 데이터 새로고침 프로세스
```
버튼 클릭 감지 → 데이터 소스에서 데이터 가져오기 → Excel 시트 업데이트 → 버튼 상태 복귀
```

### 3. 시각적 피드백
- 클릭 시 버튼 배경색 변경 (파란색 → 연한 녹색)
- 새로고침 시간 표시
- 데이터 업데이트 완료 알림

## 설정 파일 (`button_config.json`)

```json
{
  "button_cell": "A1",
  "button_text": "Refresh",
  "data_start_row": 2,
  "data_start_col": 2,
  "monitor_file_changes": true,
  "auto_refresh_interval": 0,
  "button_styles": {
    "clicked_text": "Refresh ✓",
    "normal_text": "Refresh",
    "clicked_color": "90EE90",
    "normal_color": "FFFFFF"
  },
  "data_sources": {
    "stock_data": {
      "enabled": true,
      "symbols": ["삼성전자", "SK하이닉스", "LG화학", "NAVER", "카카오"]
    },
    "market_data": {
      "enabled": true,
      "indices": ["KOSPI", "KOSDAQ", "S&P500"]
    }
  }
}
```

## 파일 구조

```
quantiwise_control/
├── excel_data/
│   └── refresh_test.xlsx          # 실제 Refresh 버튼이 있는 Excel 파일
├── button_click_simulator.py      # 기본 버튼 클릭 시뮬레이션
├── advanced_button_system.py      # 고급 버튼 시스템 (파일 모니터링)
├── create_button_excel.py         # 실제 버튼이 있는 Excel 파일 생성
├── quantiwise_refresh.py          # 텍스트 기반 구현
├── advanced_quantiwise_refresh.py # 고급 텍스트 기반 구현
├── requirements.txt               # 필요한 패키지
├── button_config.json             # 버튼 시스템 설정 (자동 생성)
└── quantiwise_config.json         # 텍스트 기반 설정 (자동 생성)
```

## 실제 사용 시나리오

### 시나리오 1: 수동 새로고침
1. Excel 파일을 열고 A1 셀의 Refresh 버튼 확인
2. 파이썬 스크립트 실행: `python button_click_simulator.py`
3. 자동으로 버튼 클릭 시뮬레이션 및 데이터 새로고침

### 시나리오 2: 실시간 모니터링
1. 고급 시스템 실행: `python advanced_button_system.py`
2. 파일 모니터링 모드 선택
3. Excel에서 A1 셀의 "Refresh"를 "Refresh ✓"로 수동 변경
4. 자동으로 클릭 감지 및 데이터 새로고침 실행

### 시나리오 3: 자동화된 새로고침
1. 설정 파일에서 `auto_refresh_interval` 설정
2. 주기적으로 버튼 클릭 시뮬레이션 실행
3. 실시간 데이터 업데이트

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

### 웹 인터페이스 연동
```python
from flask import Flask, request, jsonify

app = Flask(__name__)

@app.route('/refresh', methods=['POST'])
def trigger_refresh():
    button_system = AdvancedQuantiwiseButtonSystem("excel_data/refresh_test.xlsx")
    button_system.simulate_button_click()
    return jsonify({"status": "success"})
```

## 문제 해결

### 버튼이 감지되지 않는 경우
1. A1 셀에 정확히 "Refresh" 텍스트가 있는지 확인
2. 설정 파일의 `button_cell` 설정 확인
3. Excel 파일이 다른 프로그램에서 열려있지 않은지 확인

### 파일 모니터링이 작동하지 않는 경우
1. `watchdog` 패키지 설치 확인
2. 파일 권한 확인
3. 바이러스 백신 소프트웨어의 실시간 보호 확인

### 데이터 업데이트가 되지 않는 경우
1. Excel 파일이 읽기 전용이 아닌지 확인
2. 파일 경로가 올바른지 확인
3. 데이터 소스 설정이 활성화되어 있는지 확인

## 주의사항

1. **파일 접근**: Excel 파일이 다른 프로그램에서 열려있으면 업데이트가 실패할 수 있습니다.
2. **권한**: 파일 쓰기 권한이 필요합니다.
3. **백업**: 중요한 데이터는 미리 백업하세요.
4. **성능**: 대용량 데이터의 경우 처리 시간이 오래 걸릴 수 있습니다.

## 라이선스

이 프로젝트는 MIT 라이선스 하에 배포됩니다.
