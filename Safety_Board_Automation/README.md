# 무재해·무사고 기록판 자동 갱신

`template/AJOU_무재해기록판.pptx`(업로드하신 원본)의 아래 두 값을 매일 오늘 날짜 기준으로
자동 갱신하고, 슬라이드를 PNG 이미지로 바탕화면(Desktop)에 저장하는 스크립트입니다.

- **달성일수** (예: `51일`) — 시작일(2026-07-08)부터 오늘까지 경과일수. 목표일수 `730일`은 그대로 둡니다.
- **현재 날짜** (예: `2026년 08월 28일`) — 오늘 날짜로 갱신. 목표일(시작일, 2026년 07월 08일)은 그대로 둡니다.

원본 템플릿 파일은 수정하지 않고, 실행할 때마다 값을 새로 계산해 이미지만 만들어냅니다.
그래서 하루 실행을 건너뛰어도 다음 실행 때 자동으로 맞는 값이 채워집니다.

## 1. 준비물 (최초 1회)

1. **Python 3.9 이상** 설치: https://www.python.org (설치 시 "Add python.exe to PATH" 체크)
2. 이 `Safety_Board_Automation` 폴더를 원하는 위치(예: `C:\SafetyBoard`)에 복사
3. PowerPoint 이미지 변환을 위해 아래 중 하나가 PC에 필요합니다.
   - **Microsoft PowerPoint** (이미 설치돼 있다면 별도 설정 불필요, 가장 화질이 좋습니다)
   - 또는 **LibreOffice** (무료, https://www.libreoffice.org) — PowerPoint가 없는 PC용 대안
4. 폴더에서 명령 프롬프트(또는 PowerShell)를 열고 아래 명령으로 필요한 패키지를 설치합니다.

   ```
   pip install -r requirements.txt
   ```

## 2. 수동으로 한 번 실행해보기

```
python update_safety_board.py
```

정상적으로 실행되면 바탕화면에 `무재해기록판_YYYYMMDD.png` 파일이 생성됩니다.
(`완료: C:\Users\...\Desktop\무재해기록판_20260829.png` 메시지가 뜨면 성공입니다.)

## 3. 매일 자동 실행 등록 (Windows)

같은 폴더에서 PowerShell을 열고 아래 명령을 실행하면, **매일 오전 7:30**에 자동으로
스크립트가 실행되도록 Windows 작업 스케줄러에 등록됩니다. (최초 1회만 실행하면 됩니다.)

```
powershell -ExecutionPolicy Bypass -File .\register_task_scheduler.ps1
```

- 실행 시간을 바꾸고 싶으면 `register_task_scheduler.ps1` 파일의 `-At 7:30am` 부분을 원하는 시간으로
  수정한 뒤 다시 실행하세요.
- 등록된 작업은 Windows "작업 스케줄러" 앱에서 `무재해기록판_자동갱신` 이름으로 확인/삭제할 수 있습니다.

### macOS / Linux에서 자동 실행하려면

`cron`으로 매일 실행하도록 등록할 수 있습니다 (`crontab -e` 후 아래 한 줄 추가, 매일 07:30 기준):

```
30 7 * * * /usr/bin/python3 /경로/Safety_Board_Automation/update_safety_board.py
```

## 4. 설정 변경 (`config.json`)

| 항목 | 설명 |
|---|---|
| `template_pptx` | 원본 PPTX 템플릿 경로 (수정하지 않고 읽기만 함) |
| `start_date` | 무사고 시작일 (`YYYY-MM-DD`), 달성일수 계산 기준 |
| `output_dir` | 이미지 저장 위치. `"desktop"`이면 자동으로 바탕화면을 찾습니다 (OneDrive로 리디렉션된 바탕화면도 인식). 특정 폴더를 쓰려면 절대경로로 변경 |
| `filename_pattern` | 저장 파일명 형식 (`strftime` 형식, 기본값은 `무재해기록판_20260829.png`처럼 생성) |
| `export_scale` | PowerPoint로 내보낼 때 이미지 배율 (기본 2배, 더 선명하게 하려면 값을 올리세요) |

## 5. 문제 해결

- **"이미지 저장에 실패했습니다" 메시지가 뜰 때**: PowerPoint 또는 LibreOffice가 설치돼 있는지, 그리고
  `pip install -r requirements.txt`가 정상 완료됐는지 확인해 주세요.
- **PowerPoint가 실행 중일 때 오류가 나면**: 스크립트가 백그라운드에서 PowerPoint를 여닫는 동안 다른
  PowerPoint 창과 충돌할 수 있습니다. 스크립트 실행 전에는 PowerPoint를 닫아두는 것을 권장합니다.
- **템플릿 구조를 바꾼(직접 표/텍스트 박스를 수정한) 경우**: `update_safety_board.py`가 안전 장치로
  "템플릿 구조가 변경된 것 같습니다" 오류를 내며 멈춥니다. 이 경우 원본 템플릿으로 되돌리거나
  코드의 `shape_id`(730일/51일/날짜 텍스트박스 4개)를 새 구조에 맞게 조정해야 합니다.
