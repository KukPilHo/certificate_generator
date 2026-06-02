#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
수료증 자동 생성기 (크로스플랫폼: macOS / Windows / Linux 동일 동작)

엑셀의 각 행을 PPTX 템플릿의 placeholder(_Name 등)에 채워 1인 1파일로 저장하고,
LibreOffice가 있으면 같은 결과를 PDF로도 변환한다.

설계 핵심
  - 데이터 읽기 : openpyxl (pandas 불필요 → 패키징 가벼움)
  - 템플릿 병합 : python-pptx (run 분할 안전 치환)
  - 폰트 통일   : 궁서(Gungsuh) 등 Windows 전용 한글 폰트를 나눔 계열로 리맵
                  → Mac·Windows 양쪽에 나눔 폰트만 깔면 렌더 결과가 동일
  - PDF 변환    : LibreOffice headless (OS별 경로 자동 탐색). 없으면 PPTX만 생성.

사용 예 (한 줄):
  python3 generate_certificates.py \
    --excel "certificate excel/수료증 엑셀 파일.xlsx" \
    --template "certificate ppt/certificate.pptx"
"""

import argparse
import datetime as _dt
import os
import pathlib
import platform
import re
import shutil
import subprocess
import sys
import tempfile

from openpyxl import load_workbook
from pptx import Presentation
from pptx.oxml.ns import qn

# Windows 콘솔(cp949 등)에서 한글 로그가 깨지지 않도록 출력 인코딩을 UTF-8로 고정
for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding="utf-8")
    except (AttributeError, ValueError):
        pass

# ─────────────────────────────────────────────────────────────────────────────
# CONFIG  (비개발자도 여기만 고치면 됨)
# ─────────────────────────────────────────────────────────────────────────────

# 엑셀 컬럼 헤더 → PPT 템플릿 placeholder
# (템플릿에 placeholder가 없는 항목은 자동으로 건너뜀)
COLUMN_MAP = {
    "No": "_No",
    "이름": "_Name",
    "소속": "_Organization",
    "생년월일": "_BirthDay",     # 현재 템플릿엔 placeholder 없음 → 무시됨
    "프로젝트명": "_ProjectName",
    "부제": "_SubTitle",         # 현재 템플릿엔 placeholder 없음 → 무시됨
    "진행기간": "_Date",
    "날짜[수료일]": "_FinDate",
}

# 파일명에 쓸 placeholder (이름)
NAME_PLACEHOLDER = "_Name"

# 출력 폴더 묶음 기준: 아래 placeholder 값들을 '_'로 이어 폴더명을 만든다.
# 예) _Organization="주식회사에이비씨", _ProjectName="생성형 AI 입문과정"
#     → 폴더 "주식회사에이비씨_생성형 AI 입문과정"
# 결과: output ppt/<폴더>/이름.pptx , output pdf/<폴더>/이름.pdf
GROUP_PLACEHOLDERS = ("_Organization", "_ProjectName")

# 날짜 출력 형식: 2026년 5월 21일  (월/일 앞자리 0 없음)
def format_date(value: _dt.date) -> str:
    return f"{value.year}년 {value.month}월 {value.day}일"

# ── 폰트 통일 정책 ───────────────────────────────────────────────────────────
# 수료증의 모든 글자를 단일 폰트로 통일한다.
#   UNIFIED_FONT = "AUTO" : 실행 OS의 '기본 설치' 한글 폰트 사용 (다운로드 불필요)
#                           - macOS   → Apple SD Gothic Neo (맑은 고딕과 같은 산돌 계열)
#                           - Windows → 맑은 고딕 (Malgun Gothic)
#                           - Linux   → NanumGothic
#   UNIFIED_FONT = "폰트명": 모든 OS에서 그 폰트로 강제 (해당 폰트가 설치돼 있어야 함)
#   UNIFIED_FONT = None    : 통일하지 않고 아래 FONT_MAP 치환 방식 사용(구버전 동작)
UNIFIED_FONT = "AUTO"


def _os_default_korean_font() -> str:
    s = platform.system()
    if s == "Windows":
        return "맑은 고딕"            # Malgun Gothic (Windows 기본 설치)
    if s == "Darwin":
        return "Apple SD Gothic Neo"  # macOS 기본 설치
    return "NanumGothic"              # Linux 등


def resolve_unified_font():
    """통일할 폰트명 반환. None이면 FONT_MAP 치환 방식."""
    if UNIFIED_FONT is None:
        return None
    return _os_default_korean_font() if UNIFIED_FONT == "AUTO" else UNIFIED_FONT


# UNIFIED_FONT = None 일 때만 쓰이는 치환표(OS 종속 폰트 → 대체 폰트)
FONT_MAP = {
    "Gungsuh": "NanumMyeongjo",
    "Batang": "NanumMyeongjo",
    "Malgun Gothic": "NanumGothic",
    "맑은 고딕": "NanumGothic",
    "Gulim": "NanumGothic",
    "Dotum": "NanumGothic",
    "Nanum Gothic": "NanumGothic",
}

DEFAULT_PPT_OUT = "output ppt"
DEFAULT_PDF_OUT = "output pdf"

# ─────────────────────────────────────────────────────────────────────────────
# 유틸
# ─────────────────────────────────────────────────────────────────────────────

def sanitize_filename(name: str) -> str:
    """파일명 안전 처리: 양끝 공백 제거 → 내부 공백을 _ → 금지문자 제거."""
    n = (name or "").strip().replace(" ", "_")
    n = re.sub(r'[\\/:*?"<>|]', "_", n)
    return n or "noname"


def sanitize_dirname(name: str) -> str:
    """폴더명 안전 처리: 공백은 유지하되 금지문자 제거 + 끝 공백/점 제거(Windows 호환)."""
    n = (name or "").strip()
    n = re.sub(r'[\\/:*?"<>|]', "_", n)
    n = n.rstrip(". ")
    return n or "미지정"


def group_name(mapping: dict) -> str:
    """GROUP_PLACEHOLDERS 값들을 '_'로 이어 그룹 폴더명을 만든다."""
    parts = [mapping.get(ph, "").strip() for ph in GROUP_PLACEHOLDERS]
    parts = [p for p in parts if p]
    return sanitize_dirname("_".join(parts)) if parts else "미지정"


def cell_to_text(value) -> str:
    """엑셀 셀 값을 placeholder에 넣을 문자열로 변환."""
    if value is None:
        return ""
    if isinstance(value, (_dt.datetime, _dt.date)):
        return format_date(value)
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


# ─────────────────────────────────────────────────────────────────────────────
# 엑셀 읽기
# ─────────────────────────────────────────────────────────────────────────────

def read_rows(excel_path: str):
    """엑셀 첫 시트를 읽어 [{헤더: 값}, ...] 반환. 빈 행은 건너뜀."""
    wb = load_workbook(excel_path, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return []
    headers = [str(h).strip() if h is not None else "" for h in rows[0]]
    records = []
    for r in rows[1:]:
        if all(c is None or str(c).strip() == "" for c in r):
            continue  # 완전히 빈 행 skip
        records.append(dict(zip(headers, r)))
    return records


# ─────────────────────────────────────────────────────────────────────────────
# PPTX 치환 (run 분할 안전)
# ─────────────────────────────────────────────────────────────────────────────

def _replace_in_paragraph(paragraph, mapping):
    """문단 내 placeholder 치환.
    1) run 단위 치환 (서식 완전 보존) — placeholder가 한 run 안에 온전할 때
    2) run에 걸쳐 쪼개진 경우엔 문단을 첫 run으로 병합 후 치환 (첫 run 서식 유지)
    """
    runs = paragraph.runs
    if not runs:
        return
    # 1) run 단위
    for run in runs:
        t = run.text
        for k, v in mapping.items():
            if k in t:
                t = t.replace(k, v)
        if t != run.text:
            run.text = t
    # 2) run에 걸쳐 쪼개진 placeholder 처리
    full = "".join(r.text for r in runs)
    if any(k in full for k in mapping):
        for k, v in mapping.items():
            full = full.replace(k, v)
        runs[0].text = full
        for r in runs[1:]:
            r.text = ""


_FONT_TAGS = {qn("a:latin"), qn("a:ea"), qn("a:cs"), qn("a:sym"), qn("a:buFont")}


def _apply_fonts(element, unified_font):
    """element 하위 트리의 모든 폰트 지정(run rPr, endParaRPr, defRPr, 리스트 스타일 등)을 처리.
    - unified_font가 있으면: 모든 글자를 그 폰트 하나로 통일.
    - None이면: FONT_MAP에 따라 OS 종속 폰트만 치환(구버전 동작).
    run 단위만 보면 endParaRPr 등에 남는 폰트를 놓치므로 트리 전체를 순회한다."""
    for el in element.iter():
        if el.tag in _FONT_TAGS:
            if unified_font:
                el.set("typeface", unified_font)
            else:
                face = el.get("typeface")
                if face in FONT_MAP:
                    el.set("typeface", FONT_MAP[face])


def _iter_text_frames(shapes):
    """슬라이드 내 모든 텍스트 프레임(표 셀 포함)을 순회."""
    for shape in shapes:
        if shape.has_text_frame:
            yield shape.text_frame
        if shape.has_table:
            for row in shape.table.rows:
                for cell in row.cells:
                    yield cell.text_frame
        if shape.shape_type == 6:  # GROUP
            yield from _iter_text_frames(shape.shapes)


def fill_template(template_path: str, mapping: dict):
    """템플릿을 열어 placeholder 치환 + 폰트 통일 후 Presentation 반환."""
    prs = Presentation(template_path)
    font = resolve_unified_font()
    for slide in prs.slides:
        for tf in _iter_text_frames(slide.shapes):
            for p in tf.paragraphs:
                _replace_in_paragraph(p, mapping)
        _apply_fonts(slide._element, font)
    # 슬라이드가 상속하는 레이아웃/마스터의 폰트 지정도 함께 처리
    for master in prs.slide_masters:
        _apply_fonts(master._element, font)
        for layout in master.slide_layouts:
            _apply_fonts(layout._element, font)
    return prs


# ─────────────────────────────────────────────────────────────────────────────
# LibreOffice PDF 변환
# ─────────────────────────────────────────────────────────────────────────────

def find_soffice():
    """OS별 LibreOffice 실행 파일 경로 자동 탐색. 없으면 None."""
    for name in ("soffice", "libreoffice", "soffice.com"):
        p = shutil.which(name)
        if p:
            return p
    candidates = {
        "Darwin": ["/Applications/LibreOffice.app/Contents/MacOS/soffice"],
        "Windows": [
            r"C:\Program Files\LibreOffice\program\soffice.exe",
            r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        ],
        "Linux": ["/usr/bin/soffice", "/opt/libreoffice/program/soffice"],
    }
    for p in candidates.get(platform.system(), []):
        if os.path.exists(p):
            return p
    return None


def convert_to_pdf(soffice: str, pptx_path: str, pdf_out: str, log, profile_uri=None):
    """LibreOffice headless로 PPTX → PDF 변환.
    profile_uri: 격리된 사용자 프로필(file:// URI). LibreOffice가 이미 열려 있어도
    변환이 충돌하지 않도록 별도 프로필을 사용한다."""
    os.makedirs(pdf_out, exist_ok=True)
    cmd = [soffice, "--headless"]
    if profile_uri:
        cmd.append(f"-env:UserInstallation={profile_uri}")
    cmd += ["--convert-to", "pdf", "--outdir", pdf_out, pptx_path]
    try:
        subprocess.run(cmd, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        produced = os.path.join(
            pdf_out, os.path.splitext(os.path.basename(pptx_path))[0] + ".pdf"
        )
        log(f"  PDF  → {produced}")
        return True
    except subprocess.CalledProcessError as e:
        log(f"  [PDF 실패] {pptx_path}: {e.stderr.decode(errors='ignore')[:200]}")
        return False


# ─────────────────────────────────────────────────────────────────────────────
# 메인 처리
# ─────────────────────────────────────────────────────────────────────────────

def process(excel, template, ppt_out, pdf_out, make_pdf, limit, log):
    os.makedirs(ppt_out, exist_ok=True)
    records = read_rows(excel)
    if limit:
        records = records[:limit]
    if not records:
        log("엑셀에서 처리할 데이터 행을 찾지 못했습니다.")
        return

    soffice = find_soffice() if make_pdf else None
    if make_pdf and not soffice:
        log("⚠ LibreOffice를 찾지 못했습니다. PPTX만 생성합니다.")
        log("  설치: brew install --cask libreoffice  (또는 README 참고)")

    # LibreOffice 변환용 격리 프로필(이미 LibreOffice가 열려 있어도 충돌 방지)
    profile_dir = profile_uri = None
    if soffice:
        profile_dir = tempfile.mkdtemp(prefix="lo_profile_")
        profile_uri = pathlib.Path(profile_dir).as_uri()

    log(f"총 {len(records)}건 처리 시작\n")
    made_ppt, made_pdf = 0, 0
    try:
        for i, rec in enumerate(records, 1):
            # 매핑 만들기 (엑셀에 있는 컬럼만)
            mapping = {}
            for col, ph in COLUMN_MAP.items():
                if col in rec:
                    mapping[ph] = cell_to_text(rec[col])

            prs = fill_template(template, mapping)

            raw_name = mapping.get(NAME_PLACEHOLDER) or f"row{i}"
            safe = sanitize_filename(raw_name)
            # 회사_과정 그룹 폴더로 묶어 저장
            group = group_name(mapping)
            ppt_dir = os.path.join(ppt_out, group)
            os.makedirs(ppt_dir, exist_ok=True)
            pptx_file = os.path.join(ppt_dir, f"{safe}.pptx")
            prs.save(pptx_file)
            made_ppt += 1
            log(f"[{i}/{len(records)}] {raw_name}  ({group})")
            log(f"  PPTX → {pptx_file}")

            if soffice:
                pdf_dir = os.path.join(pdf_out, group)
                if convert_to_pdf(soffice, pptx_file, pdf_dir, log, profile_uri):
                    made_pdf += 1
    finally:
        if profile_dir:
            shutil.rmtree(profile_dir, ignore_errors=True)

    log(f"\n완료: PPTX {made_ppt}건" + (f", PDF {made_pdf}건" if soffice else ""))


def build_parser():
    p = argparse.ArgumentParser(
        description="엑셀 데이터로 수료증 PPTX(+PDF)를 일괄 생성합니다.",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    p.add_argument("--excel", required=True, help="입력 엑셀(.xlsx) 경로")
    p.add_argument("--template", required=True, help="PPTX 템플릿 경로")
    p.add_argument("--out-ppt", default=DEFAULT_PPT_OUT, help="PPTX 출력 폴더")
    p.add_argument("--out-pdf", default=DEFAULT_PDF_OUT, help="PDF 출력 폴더")
    p.add_argument("--no-pdf", action="store_true", help="PDF 변환 생략(PPTX만)")
    p.add_argument("--limit", type=int, default=0, help="앞에서 N행만 처리(0=전체, 테스트용)")
    return p


def main(argv=None):
    args = build_parser().parse_args(argv)
    for path, label in ((args.excel, "엑셀"), (args.template, "템플릿")):
        if not os.path.isfile(path):
            print(f"오류: {label} 파일을 찾을 수 없습니다 → {path}", file=sys.stderr)
            return 2
    process(
        excel=args.excel,
        template=args.template,
        ppt_out=args.out_ppt,
        pdf_out=args.out_pdf,
        make_pdf=not args.no_pdf,
        limit=args.limit,
        log=print,
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
