# -*- coding: utf-8 -*-
"""
PPT 폰트 정리기 (GUI)
=====================

PPTX 파일을 드래그앤드롭(또는 파일 선택)하면, 시스템에 없어서 문제를 일으키는
비표준 폰트와 OTF 임베드 폰트를 모두 제거하고 기본 폰트(맑은 고딕)로 대치합니다.

실행:
    python ppt_font_fixer.py

드래그앤드롭을 쓰려면(선택):
    pip install tkinterdnd2
설치되어 있지 않으면 자동으로 '파일 선택' 버튼 방식으로 동작합니다.
"""

import os
import sys
import threading
import traceback

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from font_replacer import DEFAULT_FONT, replace_fonts

# 드래그앤드롭 지원 여부 확인
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    _HAS_DND = True
except Exception:
    _HAS_DND = False


PPT_EXTS = (".pptx", ".pptm", ".potx")


class App:
    def __init__(self, root):
        self.root = root
        self.root.title("PPT 폰트 정리기")
        self.root.geometry("680x520")
        self.root.minsize(560, 440)

        self.font_var = tk.StringVar(value=DEFAULT_FONT)
        self.suffix_var = tk.StringVar(value="_폰트정리")
        self.overwrite_var = tk.BooleanVar(value=False)
        self._busy = False

        self._build_ui()

    # ---------------- UI 구성 ----------------
    def _build_ui(self):
        pad = {"padx": 10, "pady": 6}

        top = ttk.Frame(self.root)
        top.pack(fill="x", **pad)

        ttk.Label(top, text="대치할 기본 폰트:").grid(row=0, column=0, sticky="w")
        ttk.Entry(top, textvariable=self.font_var, width=20).grid(
            row=0, column=1, sticky="w", padx=(4, 16))

        ttk.Label(top, text="저장 접미사:").grid(row=0, column=2, sticky="w")
        self._suffix_entry = ttk.Entry(top, textvariable=self.suffix_var,
                                       width=12)
        self._suffix_entry.grid(row=0, column=3, sticky="w", padx=(4, 16))

        ttk.Checkbutton(top, text="원본 덮어쓰기",
                        variable=self.overwrite_var,
                        command=self._toggle_overwrite).grid(
            row=0, column=4, sticky="w")

        # 드롭 영역
        drop_text = ("여기로 PPTX 파일을 끌어다 놓으세요"
                     if _HAS_DND else "아래 '파일 선택'으로 PPTX 파일을 고르세요")
        self.drop = tk.Label(
            self.root, text=drop_text,
            relief="ridge", borderwidth=2, height=5,
            bg="#f4f6f8", fg="#333333",
            font=("맑은 고딕", 12))
        self.drop.pack(fill="x", **pad)

        if _HAS_DND:
            self.drop.drop_target_register(DND_FILES)
            self.drop.dnd_bind("<<Drop>>", self._on_drop)

        btns = ttk.Frame(self.root)
        btns.pack(fill="x", **pad)
        ttk.Button(btns, text="파일 선택…", command=self._choose_files).pack(
            side="left")
        ttk.Button(btns, text="로그 지우기", command=self._clear_log).pack(
            side="left", padx=6)
        self._status = ttk.Label(btns, text="대기 중")
        self._status.pack(side="right")

        # 로그 영역
        logframe = ttk.Frame(self.root)
        logframe.pack(fill="both", expand=True, **pad)
        self.log = tk.Text(logframe, wrap="word", height=12,
                           state="disabled", bg="#1e1e1e", fg="#d4d4d4",
                           font=("Consolas", 10))
        self.log.pack(side="left", fill="both", expand=True)
        sb = ttk.Scrollbar(logframe, command=self.log.yview)
        sb.pack(side="right", fill="y")
        self.log.config(yscrollcommand=sb.set)

        self._log("준비 완료. PPTX 파일을 추가하세요.")
        if not _HAS_DND:
            self._log("(드래그앤드롭을 쓰려면 'pip install tkinterdnd2' 후 다시 실행)")

    def _toggle_overwrite(self):
        state = "disabled" if self.overwrite_var.get() else "normal"
        self._suffix_entry.config(state=state)

    # ---------------- 로그 ----------------
    def _log(self, msg):
        self.log.config(state="normal")
        self.log.insert("end", msg + "\n")
        self.log.see("end")
        self.log.config(state="disabled")
        self.root.update_idletasks()

    def _clear_log(self):
        self.log.config(state="normal")
        self.log.delete("1.0", "end")
        self.log.config(state="disabled")

    def _set_status(self, text):
        self._status.config(text=text)
        self.root.update_idletasks()

    # ---------------- 파일 입력 ----------------
    def _on_drop(self, event):
        paths = self._parse_drop(event.data)
        self._process_files(paths)

    @staticmethod
    def _parse_drop(data):
        # tkdnd 는 공백 포함 경로를 {중괄호}로 감싼다.
        paths, buf, in_brace = [], "", False
        for ch in data:
            if ch == "{":
                in_brace = True
            elif ch == "}":
                in_brace = False
                paths.append(buf); buf = ""
            elif ch == " " and not in_brace:
                if buf:
                    paths.append(buf); buf = ""
            else:
                buf += ch
        if buf:
            paths.append(buf)
        return [p for p in paths if p]

    def _choose_files(self):
        paths = filedialog.askopenfilenames(
            title="PPTX 파일 선택",
            filetypes=[("PowerPoint 파일", "*.pptx *.pptm *.potx"),
                       ("모든 파일", "*.*")])
        if paths:
            self._process_files(list(paths))

    # ---------------- 처리 ----------------
    def _process_files(self, paths):
        if self._busy:
            self._log("처리 중입니다. 잠시 기다려 주세요.")
            return
        ppts = [p for p in paths
                if os.path.splitext(p)[1].lower() in PPT_EXTS]
        skipped = [p for p in paths if p not in ppts]
        for s in skipped:
            self._log("건너뜀(지원 안 함): %s" % os.path.basename(s))
        if not ppts:
            self._log("처리할 PPTX 파일이 없습니다.")
            return

        self._busy = True
        self._set_status("처리 중…")
        t = threading.Thread(target=self._worker, args=(ppts,), daemon=True)
        t.start()

    def _worker(self, ppts):
        default_font = self.font_var.get().strip() or DEFAULT_FONT
        overwrite = self.overwrite_var.get()
        suffix = self.suffix_var.get().strip() or "_폰트정리"

        ok_count, fail_count, total_replaced, total_removed = 0, 0, 0, 0
        for path in ppts:
            self._log("\n========== %s ==========" % os.path.basename(path))
            try:
                if overwrite:
                    out_path = path
                    # 안전을 위해 임시 파일로 만든 뒤 교체
                    tmp = path + ".fixtmp"
                    _, summary = replace_fonts(
                        path, output_path=tmp, default_font=default_font,
                        log=self._log)
                    os.replace(tmp, out_path)
                    summary["output"] = out_path
                else:
                    base, ext = os.path.splitext(path)
                    out_path = base + suffix + ext
                    _, summary = replace_fonts(
                        path, output_path=out_path,
                        default_font=default_font, log=self._log)
                ok_count += 1
                total_replaced += summary["replaced_total"]
                total_removed += len(summary["removed_embedded"])
            except Exception as e:
                fail_count += 1
                self._log("오류: %s" % e)
                self._log(traceback.format_exc())

        self._log("\n==================================================")
        self._log("완료: 성공 %d개, 실패 %d개" % (ok_count, fail_count))
        self._log("총 %d곳 폰트 치환, 임베드 폰트 %d개 제거"
                  % (total_replaced, total_removed))
        self._busy = False
        self._set_status("완료 (성공 %d / 실패 %d)" % (ok_count, fail_count))
        if fail_count == 0:
            self.root.after(
                0, lambda: messagebox.showinfo(
                    "완료",
                    "%d개 파일 처리 완료.\n폰트 %d곳 대치, 임베드 폰트 %d개 제거."
                    % (ok_count, total_replaced, total_removed)))


def main():
    if _HAS_DND:
        root = TkinterDnD.Tk()
    else:
        root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
