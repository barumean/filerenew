# 파일 정리기 (File Renew)

엑셀과 PowerPoint 파일에서 **오류를 자주 일으키는 요소들을 한 번에 정리**해
정상 파일로 되돌려 주는 통합 프로그램입니다. 기존의 **엑셀 정리기(excelrenew)**
와 **PPT 폰트 정리기(pptrenew)** 를 하나로 합쳐, 같은 창·같은 명령에서 두 종류의
파일을 함께 처리합니다.

- **엑셀** — 깨진 정의된 이름 · 외부 링크 · 숨겨진 시트 정리
- **PPT** — 비표준 폰트 · 임베드(끼워넣은) 폰트 정리

GUI는 **[엑셀 정리기] / [PPT 정리기] 두 탭으로 기능이 분리**되어 있고,
명령줄은 확장자를 보고 **자동으로 알맞은 정리**를 적용합니다. 결과 보고
형식·저장 규칙·입력 옵션은 두 포맷에서 **동일하게 통일**되어 있습니다.

## 지원 파일 형식

| 형식 | 종류 | 비고 |
| --- | --- | --- |
| `.xlsx` / `.xlsm` / `.xltx` / `.xltm` | 엑셀 | 매크로·템플릿 포함 (매크로 보존) |
| `.pptx` / `.pptm` / `.potx` | PPT | 매크로·템플릿 포함 |
| `.xls` / `.ppt` | ❌ | 구형 이진 포맷. 최신 형식으로 저장 후 사용 |

## 정리하는 항목

### 엑셀

1. **정의된 이름(Defined Names) 전부 삭제** — 숨겨진 이름, `#REF!` 로 깨진 이름,
   전역/시트범위 이름까지 모두 제거. (VBA 의 `Names.Delete` 와 동일한 효과)
2. **외부 링크/연결 제거** — 다른 통합문서를 참조하는 외부 링크와 연결을 제거.
3. **숨겨진 시트 다시 표시** — `hidden` / `veryHidden` 상태의 시트를 모두 복구.
4. **인쇄 영역 보존(옵션)** — 인쇄 영역/제목(`_xlnm.Print_Area`,
   `_xlnm.Print_Titles`)은 기본적으로 남겨 둡니다.
5. **계산 캐시(calcChain) 정리** — "제거된 레코드(계산 속성)" 경고의 원인인
   `xl/calcChain.xml` 을 미리 제거합니다(엑셀이 자동 재생성하므로 무해).

### PPT

1. **비표준 폰트 치환** — PC에 없어 글자가 깨지는 폰트를 기본 폰트(맑은 고딕)로
   치환합니다. 한글/영문 기본 폰트와 기호·불릿 폰트(Wingdings/Symbol…)는
   **그대로 보존**해 기호가 깨지지 않습니다. `+mn-lt` 같은 테마 참조도 보존.
2. **임베드 폰트 제거** — `ppt/fonts/*.fntdata` 와 `presentation.xml` 의 임베드
   폰트 목록/옵션, 관련 관계(rels)를 제거합니다. OTF 임베드 폰트의 *"TrueType
   폰트가 아니므로 대치할 수 없습니다"* 경고를 근본적으로 해결합니다.

## 동작 원리 (왜 안전한가)

`.xlsx`·`.pptx` 등은 사실 여러 XML 파일을 담은 **ZIP 압축 파일**입니다. 이
프로그램은 `openpyxl` · `python-pptx` 같은 라이브러리로 파일을 다시 저장하지
않고, ZIP 안에서 **문제되는 부분만 외과적으로 수정**합니다(`filerenew/core.py`
의 `rewrite_zip`).

덕분에 **차트·피벗테이블·매크로·슬라이드·서식 등 나머지 내용은 원본 그대로
보존**되며, **필수 외부 라이브러리가 없습니다(파이썬 표준 라이브러리만으로
동작).** 드래그 앤 드롭만 선택적으로 `tkinterdnd2` 를 사용합니다(없어도 정상 동작).

> ⚠️ 구형 `.xls` / `.ppt`(이진 포맷)는 지원하지 않습니다. 엑셀/파워포인트에서
> 최신 형식(`.xlsx` / `.pptx`)으로 먼저 저장한 뒤 사용해 주세요.

## 사용법 (GUI)

파이썬이 설치된 환경에서:

```bash
python run_gui.py
```

창 위쪽의 **[엑셀 정리기] / [PPT 정리기] 탭**으로 기능이 분리되어 있습니다.
각 탭은 자기만의 파일 목록·정리 옵션·실행 버튼·결과 로그를 독립적으로 가집니다.

1. 사용할 탭 선택 (엑셀 또는 PPT)
2. 파일 추가 — **드래그 앤 드롭** 또는 **파일 추가…** 버튼 (탭에 맞는 형식만 추가됨)
3. **정리 항목** 선택 (탭별 옵션)
4. **저장 방식** 선택
   - *새 파일로 저장* (기본/권장): 원본은 그대로 두고 `<이름>_정리됨.<확장자>` 생성
   - *원본 덮어쓰기*: 자동으로 `.bak` 백업을 만든 뒤 원본을 교체
5. **정리 실행** 클릭

> 💡 **드래그 앤 드롭**은 `tkinterdnd2` 패키지가 설치돼 있을 때 활성화됩니다.
> `pip install tkinterdnd2` 로 설치하세요. 없으면 자동으로 버튼 방식으로만
> 동작합니다(프로그램은 정상 실행됨).

## 사용법 (명령줄)

확장자를 보고 엑셀/PPT를 자동 판별하므로 섞어서 한 번에 처리할 수 있습니다.

```bash
# 새 파일로 저장 (원본 보존) — 엑셀·PPT 함께
python -m filerenew "보고서.xlsx" "발표자료.pptx"

# 원본 덮어쓰기(.bak 백업 생성)
python -m filerenew --overwrite "파일.xlsx"

# 엑셀: 특정 작업만 끄기 / 인쇄 영역도 삭제 / 계산 캐시 유지
python -m filerenew --no-links --no-unhide "파일.xlsx"
python -m filerenew --delete-print-areas --keep-calc-chain "파일.xlsx"

# PPT: 대치 기본 폰트 지정
python -m filerenew --font "맑은 고딕" "발표.pptx"
```

엑셀 옵션은 PPT 파일에, PPT 옵션은 엑셀 파일에 아무 영향을 주지 않습니다.

## Windows 실행 파일(.exe) 만들기

파이썬이 없는 PC에서도 더블클릭으로 쓰도록 단일 `.exe` 로 빌드할 수 있습니다.
(빌드는 **배포 대상과 같은 OS**에서 진행하세요. Windows exe는 Windows에서 빌드.)

**가장 쉬운 방법 (Windows):** `build.bat` 더블클릭
가상환경 생성 → 의존성 설치 → 빌드까지 자동으로 진행되고,
결과물은 `dist\파일정리기.exe` (단일 파일)로 나옵니다.

**수동 빌드:**

```bash
pip install -r requirements.txt
pip install pyinstaller
pyinstaller build.spec
```

`build.spec` 은 드래그앤드롭 라이브러리(`tkinterdnd2`)의 네이티브 바이너리까지
자동으로 포함(`collect_all`)합니다. 아이콘을 넣으려면 spec 파일의 `icon=` 줄
주석을 해제하세요.

## 테스트

```bash
python -m unittest discover -s tests -v
```

엑셀·PPT 각각 합성 파일을 만들어 치환/보존/제거가 올바른지 검증합니다.

## 프로젝트 구성

```
filerenew/
├── filerenew/                # 핵심 패키지
│   ├── core.py               # 공통: CleanResult, ZIP 처리·백업·저장 헬퍼
│   ├── excel_cleaner.py      # 엑셀 정리 로직 (clean_workbook)
│   ├── ppt_cleaner.py        # PPT 폰트 정리 로직 (clean_presentation)
│   ├── dispatcher.py         # 확장자별 라우팅 (clean_file)
│   ├── cli.py / __main__.py  # 통합 명령줄 인터페이스
│   └── gui.py                # GUI — [엑셀 정리기]/[PPT 정리기] 탭으로 분리
├── tests/                    # 단위 테스트 (unittest)
│   ├── test_excel_cleaner.py
│   └── test_ppt_cleaner.py
├── run_gui.py                # GUI 실행 진입점 (PyInstaller 진입점)
├── build.spec / build.bat    # Windows .exe 빌드
└── requirements.txt
```

### 통일된 점 (excelrenew + pptrenew → filerenew)

두 프로그램의 **정리 메커니즘(ZIP 외과 수술 방식)은 그대로 유지**하면서,
서로 달랐던 다음 항목들을 하나로 통일했습니다.

| 항목 | 이전(엑셀 / PPT) | 통일 |
| --- | --- | --- |
| 결과 보고 | `CleanResult` / `(경로, dict)` | 공통 `CleanResult` + `.summary()` |
| 저장 접미사 | `_정리됨` / `_폰트정리` | `_정리됨` |
| 덮어쓰기 백업 | `.bak` 생성 / 없음 | 양쪽 모두 `.bak` 생성 |
| 명령줄/입력 옵션 | 별도 스크립트 | `python -m filerenew` 하나로 통합 |
| 테스트 형식 | unittest / 커스텀 | unittest 로 통일 |
