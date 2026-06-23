# -*- coding: utf-8 -*-
"""font_replacer 핵심 로직 검증용 테스트 (합성 PPTX 생성)."""
import os
import tempfile
import zipfile

from font_replacer import replace_fonts

PRESENTATION_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<p:presentation xmlns:p="ns" embedTrueTypeFonts="1" saveSubsetFonts="1">'
    '<p:embeddedFontLst>'
    '<p:embeddedFont><p:font typeface="SomeWeirdOTF"/>'
    '<p:regular r:id="rId99"/></p:embeddedFont>'
    '</p:embeddedFontLst>'
    '<p:sldIdLst/></p:presentation>'
)

PRESENTATION_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="r">'
    '<Relationship Id="rId1" Type="http://x/slide" Target="slides/slide1.xml"/>'
    '<Relationship Id="rId99" Type="http://x/font" Target="fonts/font1.fntdata"/>'
    '</Relationships>'
)

THEME_XML = (
    '<?xml version="1.0"?>'
    '<a:theme xmlns:a="ns"><a:themeElements><a:fontScheme>'
    '<a:majorFont><a:latin typeface="Weird OTF Font"/>'
    '<a:ea typeface=""/><a:cs typeface=""/>'
    '<a:font script="Hang" typeface="이상한폰트"/></a:majorFont>'
    '<a:minorFont><a:latin typeface="Calibri"/>'
    '<a:ea typeface=""/><a:cs typeface=""/></a:minorFont>'
    '</a:fontScheme></a:themeElements></a:theme>'
)

SLIDE_XML = (
    '<?xml version="1.0"?>'
    '<p:sld xmlns:p="ns" xmlns:a="ns2"><p:cSld><p:spTree>'
    '<a:p><a:r><a:rPr><a:latin typeface="HY견고딕"/>'
    '<a:ea typeface="HY견고딕"/><a:cs typeface="+mn-cs"/></a:rPr>'
    '<a:t>안녕</a:t></a:r></a:p>'
    '<a:p><a:pPr><a:buFont typeface="Wingdings"/></a:pPr>'
    '<a:r><a:rPr><a:latin typeface="Arial"/></a:rPr><a:t>hi</a:t></a:r></a:p>'
    '</p:spTree></p:cSld></p:sld>'
)


def build_sample(path):
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", "<Types/>")
        z.writestr("ppt/presentation.xml", PRESENTATION_XML)
        z.writestr("ppt/_rels/presentation.xml.rels", PRESENTATION_RELS)
        z.writestr("ppt/theme/theme1.xml", THEME_XML)
        z.writestr("ppt/slides/slide1.xml", SLIDE_XML)
        z.writestr("ppt/fonts/font1.fntdata", b"\x00\x01\x02FAKEFONT")


def read(path, name):
    with zipfile.ZipFile(path) as z:
        return z.read(name).decode("utf-8")


def names(path):
    with zipfile.ZipFile(path) as z:
        return z.namelist()


def main():
    tmpdir = tempfile.mkdtemp()
    src = os.path.join(tmpdir, "sample.pptx")
    build_sample(src)

    out, summary = replace_fonts(src, log=print)

    print("\n=== 검증 ===")
    ok = True

    # 1) 임베드 폰트 파일 제거됨
    if "ppt/fonts/font1.fntdata" in names(out):
        print("FAIL: 임베드 폰트가 제거되지 않음"); ok = False
    else:
        print("OK: 임베드 폰트 제거됨")

    # 2) presentation.xml 에서 embeddedFontLst 제거 + 임베드 속성 제거
    pres = read(out, "ppt/presentation.xml")
    if "embeddedFontLst" in pres or "embedTrueTypeFonts" in pres:
        print("FAIL: presentation.xml 임베드 잔존:", pres); ok = False
    else:
        print("OK: presentation.xml 임베드 목록/옵션 제거됨")

    # 3) rels 에서 폰트 관계 제거 (slide 관계는 유지)
    rels = read(out, "ppt/_rels/presentation.xml.rels")
    if "fonts/font1.fntdata" in rels or "rId99" in rels:
        print("FAIL: 폰트 관계가 남음:", rels); ok = False
    elif "slides/slide1.xml" not in rels:
        print("FAIL: 슬라이드 관계가 잘못 삭제됨:", rels); ok = False
    else:
        print("OK: 폰트 관계만 제거, 슬라이드 관계 유지")

    # 4) theme : 비표준 폰트 치환, Calibri/빈값/테마참조 유지
    theme = read(out, "ppt/theme/theme1.xml")
    checks = {
        '맑은 고딕': '비표준 latin 치환',
        'typeface="Calibri"': 'Calibri 유지',
        'typeface=""': '빈 typeface 유지',
    }
    if "Weird OTF Font" in theme or "이상한폰트" in theme:
        print("FAIL: 테마 비표준 폰트 잔존:", theme); ok = False
    for needle, desc in checks.items():
        if needle not in theme:
            print("FAIL: %s 안됨:" % desc, theme); ok = False
        else:
            print("OK:", desc)

    # 5) slide : HY견고딕 치환, Arial/Wingdings/테마참조(+mn-cs) 유지
    slide = read(out, "ppt/slides/slide1.xml")
    if "HY견고딕" in slide:
        print("FAIL: 슬라이드 비표준 폰트 잔존"); ok = False
    else:
        print("OK: 슬라이드 비표준 폰트(HY견고딕) 치환됨")
    for needle, desc in {
        'typeface="Arial"': 'Arial 유지',
        'typeface="Wingdings"': 'Wingdings(불릿) 유지',
        'typeface="+mn-cs"': '테마참조 +mn-cs 유지',
    }.items():
        if needle not in slide:
            print("FAIL: %s 안됨:" % desc, slide); ok = False
        else:
            print("OK:", desc)

    # 6) 요약 통계
    print("\n요약:", summary["replaced_total"], "곳 치환,",
          len(summary["removed_embedded"]), "임베드 폰트 제거")

    print("\n결과:", "ALL PASS" if ok else "FAIL 있음")
    return 0 if ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
