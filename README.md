# BPT 부킹 선반입 관리 자동화 (v4)

BPT(부산항만터미널) 선반입 부킹 현황을 조회하고, Sinokor Portal에서 ETD를 대조해
ETD가 지난 건을 수량 수정·삭제 처리한 뒤, 작업 결과를 Excel 보고서와 Outlook 메일로 정리하는 Windows용 GUI 도구입니다.

---

## 1. 동작 순서

| 단계 | 버튼 | 동작 |
|---|---|---|
| ① | 선반입 부킹 조회 | BPT 사전반입 화면에 접속해 선사코드(SKR) 기준 부킹 목록을 수집합니다. 부킹번호·요청수량·반입수량·확인상태를 표로 표시합니다. |
| ② | ETD 확인 (전체) | 수집된 전체 부킹의 ETD를 Sinokor Portal 창에서 조회합니다. |
| ③ | ETD 확인 (선택) | 체크한 항목만 ETD를 조회합니다. |
| ④ | 선택 항목 처리 | 체크한 항목을 BPT에 수량 수정 / 삭제 처리합니다. |
| ⑤ | 전체 처리 | 분류된 전체 항목을 일괄 처리합니다. |
| ⑥ | 결과 보고 | 처리 결과를 Excel로 저장하고 Outlook 메일 초안을 만듭니다. |

조회 직후 **반입수량이 없거나 요청수량과 다른 건**이 자동으로 체크됩니다.
표에서 항목을 더블클릭하거나 Ctrl+C를 누르면 부킹번호가 클립보드에 복사됩니다.
조회한 ETD는 `etd_cache.json`에 당일 기준으로 캐시되며, 날짜가 바뀌면 자동으로 무시됩니다.

### 보고 메일에 대하여

작업이 끝나면 `BPT_Delete_Report_YYYYMMDD_HHMM.xlsx`가 실행 폴더에 생성되고,
이 파일을 첨부한 Outlook 메일이 만들어집니다.

- 수신: `bcsto@sinokor.co.kr`
- 제목: `[BPT] 선반입 개수 관리 작업 결과 (날짜)`
- **자동 발송되지 않습니다.** Outlook 임시보관함(Drafts)에만 저장되므로,
  내용을 확인한 뒤 직접 발송하시면 됩니다.

Excel 보고서는 `전체 부킹`, `ETD 지난 대상`, `처리 결과` 3개 시트로 구성됩니다.

---

## 2. 실행 환경

- **Windows 전용**입니다. `pywinauto`로 Sinokor Portal 데스크톱 창을 직접 제어하므로 macOS/Linux에서는 동작하지 않습니다.
- Python 3.9 이상
- Sinokor Portal V2가 **실행되어 로그인된 상태**여야 ETD 조회가 됩니다.
- 보고 메일 기능은 **Outlook 데스크톱 앱**이 설치·로그인되어 있어야 합니다. (웹 Outlook 불가)
- ETD 조회·처리 중에는 `pyautogui`가 마우스·키보드를 점유합니다. **실행 중에는 PC를 조작하지 마세요.**

---

## 3. 설치

```bat
:: 1) 소스 폴더로 이동
cd C:\경로\bpt_booking_app

:: 2) 패키지 설치
pip install -r requirements.txt

:: 3) Playwright 브라우저 설치 (최초 1회)
python -m playwright install chromium
```

> 앱 최초 실행 시 누락된 패키지와 Chromium을 자동 설치하도록 되어 있으나,
> 사내망 프록시 환경에서는 실패할 수 있으므로 위와 같이 미리 설치하는 것을 권장합니다.

---

## 4. 계정 설정 (필수)

보안을 위해 BPT 계정은 소스와 저장소에 들어 있지 않습니다.
`config.example.json`을 복사해 **`config.json`** 으로 저장하고, 공용 BPT 계정을 입력하세요
(계정은 담당자에게 따로 받으세요). `config.json`은 `.gitignore`에 포함돼 커밋되지 않습니다.

```json
{
  "BPT_ID": "사용할_계정_ID",
  "BPT_PW": "비밀번호"
}
```

`config.json`이 없으면 최초 실행 시 빈 양식으로 만들어지고, 계정을 넣으라는 안내 창이 뜹니다.

---

## 5. 실행

```bat
python bpt_booking_app_v4.py
```

### exe로 배포하려면

```bat
pip install pyinstaller
pyinstaller --noconsole --onefile bpt_booking_app_v4.py
```

생성된 `dist\bpt_booking_app_v4.exe`를 그대로 배포하면 됩니다.
계정을 바꿔야 할 때만 exe 옆에 `config.json`을 두면 되며, 앱이 exe 실행 위치를 기준으로 설정 파일을 찾습니다.

---

## 6. 자주 발생하는 문제

| 증상 | 원인 / 조치 |
|---|---|
| BPT 로그인 실패 | 공용 계정 비밀번호가 변경되었을 수 있습니다. `config.json`에 새 계정을 입력하세요. |
| 부킹 목록이 0건 | BPT 로그인 실패 또는 사전반입 화면 구조 변경. 로그 창 메시지를 확인하세요. |
| ETD 조회가 전부 실패 | Sinokor Portal이 실행·로그인되어 있는지 확인하세요. 창이 최소화되어 있으면 인식되지 않을 수 있습니다. |
| Chromium 실행 오류 | `python -m playwright install chromium` 재실행. |
| 메일 초안이 안 만들어짐 | Outlook 데스크톱 앱이 실행 중인지, `pywin32`가 설치되었는지 확인하세요. 보고서 Excel 파일은 정상 생성됩니다. |
| 처리 중 엉뚱한 곳이 클릭됨 | 자동화 중 마우스를 움직였을 때 발생합니다. 실행 중에는 PC를 조작하지 마세요. |

---

## 7. 파일 구성

```
bpt_booking_app_v4.py     메인 프로그램
config.example.json       BPT 계정 설정 양식 (복사해서 config.json으로 사용)
requirements.txt          의존 패키지
.gitignore                실행 산출물 제외 규칙
README.md                 이 문서
```

실행 중 자동 생성되는 파일

```
etd_cache.json                          당일 ETD 캐시
BPT_Delete_Report_YYYYMMDD_HHMM.xlsx    작업 결과 보고서
```

> 위 생성 파일들은 실제 부킹 데이터가 들어 있으므로, 프로그램을 다른 사람에게 다시 전달할 때는
> 함께 압축되지 않도록 주의하세요.

---

## 8. 주의

- BPT 및 Sinokor Portal 화면 구조가 변경되면 셀렉터 수정이 필요합니다.
- 대량 처리 전에는 **③ ETD 확인(선택)** 으로 소수 건을 먼저 검증한 뒤 ⑤ 전체 처리를 진행하는 것을 권장합니다.
- 보고 메일은 초안으로만 저장되므로, 발송 전 반드시 내용을 확인하세요.
