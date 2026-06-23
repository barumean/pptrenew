# -*- coding: utf-8 -*-
"""
PPT 폰트 정리기 (명령줄 버전)

사용 예:
    python ppt_font_fixer_cli.py 발표자료.pptx
    python ppt_font_fixer_cli.py *.pptx
    python ppt_font_fixer_cli.py 발표자료.pptx --font "맑은 고딕"
    python ppt_font_fixer_cli.py 발표자료.pptx --overwrite
"""

import argparse
import glob
import os
import sys

from font_replacer import DEFAULT_FONT, replace_fonts


def main(argv=None):
    parser = argparse.ArgumentParser(
        description="PPTX의 비표준/임베드 폰트를 기본 폰트로 대치합니다.")
    parser.add_argument("files", nargs="+",
                        help="처리할 .pptx 파일들 (와일드카드 가능)")
    parser.add_argument("--font", default=DEFAULT_FONT,
                        help="대치할 기본 폰트 (기본값: %s)" % DEFAULT_FONT)
    parser.add_argument("--suffix", default="_폰트정리",
                        help="저장 파일 접미사 (기본값: _폰트정리)")
    parser.add_argument("--overwrite", action="store_true",
                        help="원본 파일을 덮어씀")
    args = parser.parse_args(argv)

    # 와일드카드 확장
    paths = []
    for pattern in args.files:
        matched = glob.glob(pattern)
        paths.extend(matched if matched else [pattern])

    ok, fail = 0, 0
    for path in paths:
        print("\n========== %s ==========" % path)
        try:
            if args.overwrite:
                tmp = path + ".fixtmp"
                replace_fonts(path, output_path=tmp, default_font=args.font,
                              log=print)
                os.replace(tmp, path)
                print("덮어씀:", path)
            else:
                base, ext = os.path.splitext(path)
                out = base + args.suffix + ext
                replace_fonts(path, output_path=out, default_font=args.font,
                              log=print)
            ok += 1
        except Exception as e:
            fail += 1
            print("오류:", e, file=sys.stderr)

    print("\n완료: 성공 %d개, 실패 %d개" % (ok, fail))
    return 1 if fail else 0


if __name__ == "__main__":
    raise SystemExit(main())
