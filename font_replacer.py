# -*- coding: utf-8 -*-
"""
PPTX 폰트 정리 핵심 로직.

PPTX 파일은 내부적으로 ZIP 압축 안에 여러 XML 문서가 들어있는 구조입니다.
폰트는 다음 위치에서 참조됩니다.

  - ppt/theme/themeN.xml        : 테마의 majorFont / minorFont 정의
  - ppt/slides/slideN.xml 등    : <a:latin>, <a:ea>, <a:cs>, <a:sym>, <a:buFont> 의 typeface 속성
  - ppt/slideMasters, slideLayouts, notesSlides, handoutMasters
  - ppt/presentation.xml        : <p:embeddedFontLst> (임베드된 폰트 목록)
  - ppt/fonts/*.fntdata         : 실제 임베드된 폰트 바이너리

이 모듈은 위 참조들을 모두 찾아서
  1) 표준(안전) 폰트가 아닌 폰트를 기본 폰트(기본값: 맑은 고딕)로 치환하고
  2) 파일에 끼워넣어진(임베드) 폰트를 제거합니다.

OTF 임베드 폰트가 "TrueType 폰트가 아니므로 대치할 수 없습니다" 류의 경고를 내며
교체가 안 되는 문제를, 참조 자체를 기본 폰트로 바꿔버려서 한 번에 해결합니다.
"""

import os
import re
import shutil
import zipfile

# 기본 대치 폰트
DEFAULT_FONT = "맑은 고딕"

# 치환하지 않고 그대로 둘 "안전한" 폰트 목록(소문자 비교).
# - 대부분의 Windows / Office 환경에 기본 설치되어 있는 폰트
# - 기호/불릿용 폰트(이걸 바꾸면 기호가 깨지므로 반드시 보존)
SAFE_FONTS = {
    # 한글 기본
    "맑은 고딕", "맑은 고딕 semilight", "malgun gothic", "malgun gothic semilight",
    "굴림", "굴림체", "돋움", "돋움체", "바탕", "바탕체", "궁서", "궁서체",
    "gulim", "gulimche", "dotum", "dotumche", "batang", "batangche",
    "gungsuh", "gungsuhche",
    # 영문 기본 (Office 기본 테마 폰트 포함)
    "arial", "arial black", "arial narrow", "calibri", "calibri light",
    "cambria", "cambria math", "times new roman", "verdana", "tahoma",
    "segoe ui", "segoe ui light", "segoe ui semibold", "courier new",
    "georgia", "comic sans ms", "consolas", "trebuchet ms", "impact",
    "lucida sans unicode", "lucida console", "palatino linotype", "garamond",
    "century gothic", "book antiqua", "candara", "constantia", "corbel",
    "franklin gothic medium", "gabriola", "sylfaen",
    # 기호/불릿 폰트 (절대 치환 금지)
    "wingdings", "wingdings 2", "wingdings 3", "webdings", "symbol",
    "marlett", "mt extra",
}

# typeface="..." 속성을 찾는 정규식 (네임스페이스 접두사와 무관하게 동작)
_TYPEFACE_RE = re.compile(r'typeface="([^"]*)"')

# presentation.xml 의 임베드 폰트 목록 블록
_EMBEDDED_FONT_LST_RE = re.compile(
    r"<p:embeddedFontLst>.*?</p:embeddedFontLst>", re.DOTALL
)
# 폰트 임베드 관련 속성 (자기닫힘/공백 포함 다양한 형태 제거)
_EMBED_ATTR_RE = re.compile(r'\s+(?:embedTrueTypeFonts|saveSubsetFonts)="[^"]*"')

# 폰트 참조가 들어있을 수 있는 XML 파트 경로 패턴
_FONT_BEARING_PATHS = (
    "ppt/theme/",
    "ppt/slides/",
    "ppt/slideMasters/",
    "ppt/slideLayouts/",
    "ppt/notesSlides/",
    "ppt/notesMasters/",
    "ppt/handoutMasters/",
    "ppt/charts/",
    "ppt/diagrams/",
)


def _should_replace(font_name, safe_lower):
    """이 폰트 이름을 치환해야 하는가?"""
    if font_name == "":
        return False
    # '+mn-lt', '+mj-ea' 같은 테마 참조는 그대로 둔다(테마 정의를 이미 고치므로).
    if font_name.startswith("+"):
        return False
    return font_name.strip().lower() not in safe_lower


def _replace_typefaces(xml_text, default_font, safe_lower, stats):
    def repl(m):
        name = m.group(1)
        if _should_replace(name, safe_lower):
            stats[name] = stats.get(name, 0) + 1
            return 'typeface="%s"' % default_font
        return m.group(0)

    return _TYPEFACE_RE.sub(repl, xml_text)


def _is_font_bearing_xml(name):
    if not name.endswith(".xml"):
        return False
    return any(name.startswith(p) for p in _FONT_BEARING_PATHS)


def _clean_presentation_xml(xml_text):
    """presentation.xml 에서 임베드 폰트 목록과 임베드 옵션을 제거."""
    xml_text = _EMBEDDED_FONT_LST_RE.sub("", xml_text)
    xml_text = _EMBED_ATTR_RE.sub("", xml_text)
    return xml_text


def _clean_presentation_rels(xml_text):
    """presentation.xml.rels 에서 폰트 파트를 가리키는 관계(Relationship) 제거."""
    # <Relationship ... Type=".../font" ... Target="fonts/..."/> 형태 제거
    rel_re = re.compile(r"<Relationship\b[^>]*/>")

    def repl(m):
        tag = m.group(0)
        if "/font" in tag or "fonts/" in tag.replace("\\", "/"):
            return ""
        return tag

    return rel_re.sub(repl, xml_text)


def replace_fonts(input_path, output_path=None, default_font=DEFAULT_FONT,
                  log=None):
    """
    PPTX 파일의 폰트를 정리한다.

    input_path  : 원본 .pptx 경로
    output_path : 결과 저장 경로. None 이면 '<원본>_폰트정리.pptx' 로 저장.
    default_font: 비표준 폰트를 치환할 기본 폰트.
    log         : 진행 메시지를 받을 콜백 함수(str). None 이면 무시.

    반환값: (output_path, summary_dict)
    """
    def _log(msg):
        if log:
            log(msg)

    if not os.path.isfile(input_path):
        raise FileNotFoundError(input_path)

    ext = os.path.splitext(input_path)[1].lower()
    if ext not in (".pptx", ".pptm", ".potx"):
        raise ValueError("지원하지 않는 파일 형식입니다: %s (pptx/pptm/potx 만 가능)" % ext)

    if output_path is None:
        base, ext0 = os.path.splitext(input_path)
        output_path = base + "_폰트정리" + ext0

    safe_lower = {f.lower() for f in SAFE_FONTS}
    # 기본 폰트 자신도 안전 목록에 포함(자기 자신으로 무한 치환 방지/일관성)
    safe_lower.add(default_font.strip().lower())

    replaced_stats = {}
    removed_fonts = []
    cleaned_xml_count = 0

    _log("파일 열기: %s" % os.path.basename(input_path))

    # 임시 출력 파일에 새 zip 작성
    tmp_output = output_path + ".tmp"
    with zipfile.ZipFile(input_path, "r") as zin:
        names = zin.namelist()
        with zipfile.ZipFile(tmp_output, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                name = item.filename
                norm = name.replace("\\", "/")

                # 1) 임베드 폰트 바이너리 제거
                if norm.startswith("ppt/fonts/"):
                    removed_fonts.append(norm)
                    _log("  임베드 폰트 제거: %s" % norm)
                    continue

                data = zin.read(name)

                # 2) presentation.xml : 임베드 폰트 목록/옵션 제거
                if norm == "ppt/presentation.xml":
                    text = data.decode("utf-8")
                    text = _clean_presentation_xml(text)
                    text = _replace_typefaces(text, default_font, safe_lower,
                                              replaced_stats)
                    data = text.encode("utf-8")
                    cleaned_xml_count += 1

                # 3) presentation.xml.rels : 폰트 관계 제거
                elif norm == "ppt/_rels/presentation.xml.rels":
                    text = data.decode("utf-8")
                    text = _clean_presentation_rels(text)
                    data = text.encode("utf-8")

                # 4) 폰트 참조가 있는 XML : typeface 치환
                elif _is_font_bearing_xml(norm):
                    text = data.decode("utf-8")
                    text = _replace_typefaces(text, default_font, safe_lower,
                                              replaced_stats)
                    data = text.encode("utf-8")
                    cleaned_xml_count += 1

                # 원본 압축 정보를 유지하며 기록
                zout.writestr(item, data)

    # 원자적 교체
    shutil.move(tmp_output, output_path)

    total_replaced = sum(replaced_stats.values())
    summary = {
        "output": output_path,
        "default_font": default_font,
        "replaced_total": total_replaced,
        "replaced_fonts": replaced_stats,      # {원래폰트: 횟수}
        "removed_embedded": removed_fonts,     # 제거된 임베드 폰트 파일들
        "cleaned_xml": cleaned_xml_count,
    }

    _log("---- 완료 ----")
    if replaced_stats:
        _log("치환된 폰트(%d종, 총 %d곳):" % (len(replaced_stats), total_replaced))
        for fname, cnt in sorted(replaced_stats.items(),
                                 key=lambda x: -x[1]):
            _log("  '%s' -> '%s'  (%d곳)" % (fname, default_font, cnt))
    else:
        _log("치환할 비표준 폰트가 없었습니다.")
    if removed_fonts:
        _log("제거된 임베드 폰트: %d개" % len(removed_fonts))
    _log("저장됨: %s" % output_path)

    return output_path, summary
