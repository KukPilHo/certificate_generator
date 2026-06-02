# 수료증 자동 생성기 (크로스플랫폼)

엑셀의 각 행을 PPTX 템플릿의 placeholder에 채워 **1인 1파일**로 저장하고,
LibreOffice가 있으면 같은 결과를 **PDF**로도 변환합니다.
macOS · Windows · Linux에서 **같은 명령**으로 동일하게 동작합니다.

기존 Windows 전용 도구(`certificate_gui_app.py`, pywin32 COM 의존)를 대체합니다.

---

## 1. 설치

### (1) Python 패키지 — 양 OS 공통

```bash
pip3 install openpyxl python-pptx
```

### (2) LibreOffice — PDF 변환용 (PPTX만 필요하면 생략 가능)

- macOS: `brew install --cask libreoffice`
- Windows: https://www.libreoffice.org 에서 설치 (기본 경로면 자동 인식)

PDF가 필요 없으면 실행 시 `--no-pdf`를 붙이면 됩니다.

### (3) 한글 폰트 — **추가 설치 불필요** (기본값)

모든 글자를 **실행 OS의 기본 한글 폰트**로 자동 통일하므로 폰트를 따로 받을 필요가 없습니다.

- macOS → **Apple SD Gothic Neo** (기본 설치)
- Windows → **맑은 고딕 (Malgun Gothic)** (기본 설치)

> 양 OS에서 **완전히 같은 글꼴**을 쓰고 싶을 때만 공통 폰트를 설치하고
> `generate_certificates.py` 상단의 `UNIFIED_FONT`를 그 폰트명으로 지정하세요.
> 예) `UNIFIED_FONT = "NanumGothic"` + `brew install --cask font-nanum-gothic`

---

## 2. 실행

### A. 더블클릭 (가장 쉬움)

- macOS: **`수료증_생성_Mac.command`** 더블클릭
- Windows: **`수료증_생성_Windows.bat`** 더블클릭

폴더 안의 첫 번째 엑셀(.xlsx)과 템플릿(.pptx)을 자동으로 찾아 실행하고,
결과를 **`회사이름_교육과정명` 폴더로 묶어** 저장한 뒤 창에 결과를 보여줍니다.
(venv가 있으면 자동 사용, 필요한 패키지가 없으면 자동 설치)

> macOS 첫 실행 시 "확인되지 않은 개발자" 경고가 뜨면,
> 파일을 **마우스 우클릭 → 열기 → 열기**로 한 번만 허용하면 됩니다.

### B. 터미널 (한 줄)

```bash
python3 generate_certificates.py \
  --excel "certificate excel/수료증 엑셀 파일.xlsx" \
  --template "certificate ppt/certificate.pptx"
```

결과는 `output ppt/`(PPTX)와 `output pdf/`(PDF) 아래
**`회사이름_교육과정명` 폴더로 묶여서** 저장됩니다:

```
output ppt/
  테크빌교육㈜_생성형 AI 마케팅 과정/
    이승연.pptx
    조현진.pptx
output pdf/
  테크빌교육㈜_생성형 AI 마케팅 과정/
    이승연.pdf
    조현진.pdf
```

회사·과정이 다른 행은 각각 다른 폴더로 자동 분리됩니다.
묶는 기준을 바꾸려면 `generate_certificates.py` 상단의 `GROUP_PLACEHOLDERS`를 수정하세요
(예: 과정명만으로 묶으려면 `("_ProjectName",)`).

### 옵션

| 옵션 | 설명 | 기본값 |
|---|---|---|
| `--excel PATH` | 입력 엑셀(.xlsx) 경로 (필수) | — |
| `--template PATH` | PPTX 템플릿 경로 (필수) | — |
| `--out-ppt DIR` | PPTX 출력 폴더 | `output ppt` |
| `--out-pdf DIR` | PDF 출력 폴더 | `output pdf` |
| `--no-pdf` | PDF 변환 생략(PPTX만 생성) | 꺼짐 |
| `--limit N` | 앞에서 N행만 처리(테스트용) | 0=전체 |

예) 먼저 2명만 테스트: `--limit 2` 추가
예) PDF 없이 PPTX만: `--no-pdf` 추가

---

## 3. 엑셀 컬럼 → PPT placeholder 매핑

| 엑셀 컬럼 | placeholder | 현재 템플릿 | 비고 |
|---|---|---|---|
| No | `_No` | ✅ | |
| 이름 | `_Name` | ✅ | 출력 파일명에 사용 |
| 소속 | `_Organization` | ✅ | |
| 생년월일 | `_BirthDay` | ❌ 없음 | 매핑은 있으나 템플릿에 placeholder 없어 무시 |
| 프로젝트명 | `_ProjectName` | ✅ | |
| 부제 | `_SubTitle` | ❌ 없음 | 위와 동일 |
| 진행기간 | `_Date` | ✅ | 날짜 → `2026년 5월 12일` 형식 |
| 날짜[수료일] | `_FinDate` | ✅ | 날짜 → `2026년 5월 21일` 형식 |
| Email | (매핑 없음) | — | 사용 안 함 |

매핑·날짜 형식은 `generate_certificates.py` 상단 `COLUMN_MAP` / `format_date()`에서 수정합니다.

---

## 4. 한글 폰트 정책

- 한글 글자 자체(유니코드)는 어떤 OS에서도 깨지지 않습니다. 위험은 **렌더 시 폰트 치환**뿐입니다.
- 원본 템플릿은 궁서·나눔 등 **여러 폰트가 섞여** 있어 글씨가 제각각으로 보였습니다.
- 그래서 이 도구는 생성 시 슬라이드·레이아웃·마스터의 **모든 글자를 단일 폰트로 통일**합니다.
  기본값(`UNIFIED_FONT = "AUTO"`)은 **실행 OS의 기본 한글 폰트**라 추가 설치가 필요 없습니다.
  - macOS → Apple SD Gothic Neo, Windows → 맑은 고딕, Linux → NanumGothic
- 같은 OS 안에서는 모든 글자가 완전히 한 폰트로 통일됩니다.
  Mac과 Windows에서 **글꼴까지 완전히 동일**하게 맞추려면, 공통 폰트(예: 나눔고딕)를
  양쪽에 설치하고 `UNIFIED_FONT`를 그 폰트명으로 지정하세요.

---

## 5. 동작 검증 (제작 시 확인됨)

샘플 엑셀로 실제 생성·렌더링하여 다음을 확인했습니다.

- 한글 모두 정상 출력(□ 두부/깨짐 없음), 날짜 `2026년 5월 13일` 형식 정상
- 슬라이드·레이아웃·마스터 전체에서 **글꼴이 단 1종으로 통일**됨(혼용 폰트 0개)
- 텍스트 위치/레이아웃은 템플릿 그대로 유지
- 결과가 `회사이름_교육과정명` 폴더로 묶여 저장됨

> 참고: 검증 샌드박스엔 나눔명조를 설치할 수 없어 나눔고딕으로 대체 렌더했습니다.
> 실제 Mac에서는 위 (3)으로 나눔명조까지 설치하면 제목이 명조(세리프)로 렌더됩니다.
