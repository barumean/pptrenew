# PPT 폰트 정리기

PowerPoint(`.pptx`) 파일을 작업하다 보면 **PC에 없는 폰트** 때문에 글자가 깨지거나,
OTF 임베드 폰트가 *"TrueType 폰트가 아니므로 대치할 수 없습니다"* 같은 경고를 내며
교체조차 안 되는 경우가 있습니다.

이 프로그램은 PPTX 안의 **비표준 폰트 참조**와 **끼워넣어진(임베드) 폰트**를
한 번에 찾아서 모두 **기본 폰트(맑은 고딕)** 로 대치해 이 문제를 해결합니다.

## 동작 원리

`.pptx`는 사실 ZIP 압축 안에 여러 XML이 들어있는 구조입니다. 이 프로그램은:

1. 테마(`ppt/theme/*.xml`), 슬라이드/마스터/레이아웃/노트의 `typeface` 속성을 모두 검사
2. **표준(안전) 폰트가 아닌** 폰트를 기본 폰트로 치환
   - 한글 기본(맑은 고딕/굴림/돋움/바탕…), 영문 기본(Arial/Calibri/Times…),
     기호 폰트(Wingdings/Symbol…)는 **그대로 보존** → 기호·불릿이 깨지지 않음
   - `+mn-lt` 같은 테마 참조도 보존 (테마 정의 자체를 고치므로 자동 반영)
3. `ppt/fonts/*.fntdata` 임베드 폰트 바이너리 제거
4. `presentation.xml`의 임베드 폰트 목록/옵션, 관련 관계(rels) 제거

원본은 건드리지 않고 `원본이름_폰트정리.pptx`로 저장합니다(덮어쓰기 옵션 가능).

## 사용 방법

### 1) GUI (드래그앤드롭)

```bash
python ppt_font_fixer.py
```

창에 PPTX 파일을 끌어다 놓으면 처리됩니다. 여러 개를 한꺼번에 넣어도 됩니다.

드래그앤드롭을 쓰려면 한 번만 설치하세요(선택):

```bash
pip install tkinterdnd2
```

설치하지 않아도 **'파일 선택' 버튼**으로 동일하게 동작합니다.

### 2) 명령줄 (일괄 처리/자동화)

```bash
python ppt_font_fixer_cli.py 발표자료.pptx
python ppt_font_fixer_cli.py *.pptx                # 폴더 내 전체
python ppt_font_fixer_cli.py 발표.pptx --font "맑은 고딕"
python ppt_font_fixer_cli.py 발표.pptx --overwrite # 원본 덮어쓰기
```

### 3) 실행 파일(.exe)로 빌드 — 파이썬 없는 PC 배포용

파이썬이 설치되지 않은 PC에서도 쓰려면 단일 `.exe`로 빌드하세요.

**가장 쉬운 방법 (Windows):** `build.bat` 더블클릭
가상환경 생성 → 의존성 설치 → 빌드까지 자동으로 진행되고,
결과물은 `dist\PPT폰트정리기.exe` (단일 파일)로 나옵니다.

**수동 빌드:**

```bash
pip install -r requirements.txt
pyinstaller build.spec
```

`build.spec`은 드래그앤드롭 라이브러리(`tkinterdnd2`)의 네이티브
바이너리까지 자동으로 포함(`collect_all`)하므로, 별도 옵션 없이 그대로 빌드하면
드래그앤드롭이 동작하는 exe가 만들어집니다. 아이콘을 넣고 싶으면 spec 파일의
`icon=` 줄 주석을 해제하세요.

> 빌드는 **배포 대상과 같은 OS**에서 해야 합니다(Windows exe는 Windows에서 빌드).

## 구성 파일

| 파일 | 설명 |
|------|------|
| `font_replacer.py` | 핵심 변환 로직 (PPTX 파싱·치환·임베드 제거) |
| `ppt_font_fixer.py` | GUI(드래그앤드롭) 프로그램 |
| `ppt_font_fixer_cli.py` | 명령줄 버전 |
| `test_font_replacer.py` | 핵심 로직 검증 테스트 |
| `requirements.txt` | 실행/빌드 의존성 |
| `build.spec` | PyInstaller 빌드 스펙 |
| `build.bat` | Windows 원클릭 빌드 스크립트 |

## 테스트

```bash
python test_font_replacer.py
```

합성 PPTX를 만들어 치환/보존/임베드 제거가 올바른지 검증합니다.

## 참고 / 한계

- 안전 폰트 목록은 `font_replacer.py`의 `SAFE_FONTS`에서 조정할 수 있습니다.
- 기본 동작은 "비표준 폰트만" 치환합니다. 문서 내 **모든** 폰트를 강제로 통일하고
  싶다면 `SAFE_FONTS`를 비우면 됩니다.
- 처리 후에도 PowerPoint에서 한 번 열어 확인하시길 권장합니다.
```
