# -*- coding: utf-8 -*-
"""
gui.py
======
파일 정리기(엑셀 + PPT)의 그래픽 사용자 인터페이스(GUI).

엑셀 정리기와 PPT 정리기를 하나의 창으로 통합했습니다. 파일을 추가하면
확장자에 따라 자동으로 알맞은 정리가 적용됩니다(엑셀 ↔ PPT 옵션은
서로의 파일에는 영향을 주지 않습니다).

파이썬 표준 라이브러리(tkinter)만 사용하므로 별도 설치가 필요 없으며,
Windows 에서는 PyInstaller 로 .exe 단일 실행 파일을 만들 수 있습니다.

사용법:
    python -m filerenew.gui      또는      python run_gui.py
"""

from __future__ import annotations

import os
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from filerenew.dispatcher import (
    DEFAULT_FONT,
    EXCEL_EXTS,
    PPT_EXTS,
    SUPPORTED_EXTS,
    clean_file,
)

# 드래그 앤 드롭(tkinterdnd2)이 설치돼 있으면 사용하고, 없으면 버튼 방식으로
# 자동 폴백한다. (tkinterdnd2 는 OS 파일 끌어다 놓기를 지원하는 외부 라이브러리)
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD

    DND_AVAILABLE = True
except Exception:  # noqa: BLE001
    DND_AVAILABLE = False

APP_TITLE = "파일 정리기 (File Renew)"

# 지원하는 파일 형식(확장자) 표시 문자열.
SUPPORTED_DISPLAY = " ".join(sorted(SUPPORTED_EXTS))
FILETYPES = [
    ("지원 파일 (엑셀·PPT)", " ".join(f"*{e}" for e in sorted(SUPPORTED_EXTS))),
    ("엑셀 파일", " ".join(f"*{e}" for e in sorted(EXCEL_EXTS))),
    ("PowerPoint 파일", " ".join(f"*{e}" for e in sorted(PPT_EXTS))),
    ("모든 파일", "*.*"),
]


class FileRenewApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.files: list[str] = []

        root.title(APP_TITLE)
        root.geometry("760x720")
        root.minsize(680, 640)

        # 공통 저장 옵션
        self.opt_overwrite = tk.BooleanVar(value=False)

        # 엑셀 작업 옵션
        self.opt_names = tk.BooleanVar(value=True)
        self.opt_links = tk.BooleanVar(value=True)
        self.opt_unhide = tk.BooleanVar(value=True)
        self.opt_keep_print = tk.BooleanVar(value=True)

        # PPT 작업 옵션
        self.opt_replace_fonts = tk.BooleanVar(value=True)
        self.opt_remove_embedded = tk.BooleanVar(value=True)
        self.font_var = tk.StringVar(value=DEFAULT_FONT)

        self._build_ui()

    # ------------------------------------------------------------------ UI
    def _build_ui(self):
        pad = {"padx": 10, "pady": 6}

        # --- 상단 안내 ---
        ttk.Label(
            self.root,
            text="엑셀과 PPT를 한 번에 정리합니다 — 엑셀의 깨진 이름·외부 링크·숨겨진 시트,"
            " PPT의 비표준·임베드 폰트.",
            font=("", 11, "bold"),
        ).pack(anchor="w", **pad)

        ttk.Label(
            self.root,
            text=f"지원 형식: {SUPPORTED_DISPLAY}  "
            "(구형 .xls 는 미지원 — 엑셀에서 .xlsx 로 저장 후 사용)",
            foreground="#555",
        ).pack(anchor="w", padx=10)

        # --- 파일 목록 영역 ---
        list_frame = ttk.LabelFrame(self.root, text="정리할 파일")
        list_frame.pack(fill="both", expand=True, padx=10, pady=4)

        if DND_AVAILABLE:
            hint = "↓ 여기로 엑셀·PPT 파일을 끌어다 놓거나, '파일 추가…' 버튼을 누르세요"
        else:
            hint = "'파일 추가…' 버튼으로 파일을 선택하세요 (드래그 앤 드롭은 tkinterdnd2 설치 시 활성화)"
        ttk.Label(list_frame, text=hint, foreground="#777").pack(
            anchor="w", padx=10, pady=(6, 0)
        )

        inner = ttk.Frame(list_frame)
        inner.pack(fill="both", expand=True)

        self.listbox = tk.Listbox(inner, selectmode=tk.EXTENDED, height=8)
        self.listbox.pack(side="left", fill="both", expand=True, padx=(8, 0), pady=8)
        sb = ttk.Scrollbar(inner, orient="vertical", command=self.listbox.yview)
        sb.pack(side="left", fill="y", pady=8)
        self.listbox.config(yscrollcommand=sb.set)

        if DND_AVAILABLE:
            self.listbox.drop_target_register(DND_FILES)
            self.listbox.dnd_bind("<<Drop>>", self._on_drop)

        btn_col = ttk.Frame(inner)
        btn_col.pack(side="left", fill="y", padx=8, pady=8)
        ttk.Button(btn_col, text="파일 추가…", command=self.add_files).pack(fill="x", pady=2)
        ttk.Button(btn_col, text="선택 제거", command=self.remove_selected).pack(fill="x", pady=2)
        ttk.Button(btn_col, text="목록 비우기", command=self.clear_files).pack(fill="x", pady=2)

        # --- 옵션 영역: 엑셀/PPT 를 좌우로 나란히 배치 ---
        opt_row = ttk.Frame(self.root)
        opt_row.pack(fill="x", padx=10, pady=4)

        excel_frame = ttk.LabelFrame(opt_row, text="엑셀 정리 항목")
        excel_frame.pack(side="left", fill="both", expand=True, padx=(0, 4))
        ttk.Checkbutton(
            excel_frame, text="정의된 이름(Names) 전부 삭제 — 깨진/숨겨진 이름 포함",
            variable=self.opt_names,
        ).pack(anchor="w", padx=10, pady=2)
        ttk.Checkbutton(
            excel_frame, text="외부 링크/연결 제거 — 다른 통합문서 참조 제거",
            variable=self.opt_links,
        ).pack(anchor="w", padx=10, pady=2)
        ttk.Checkbutton(
            excel_frame, text="숨겨진 시트 다시 표시 — hidden/veryHidden 복구",
            variable=self.opt_unhide,
        ).pack(anchor="w", padx=10, pady=2)
        ttk.Checkbutton(
            excel_frame, text="인쇄 영역(Print Area)은 유지",
            variable=self.opt_keep_print,
        ).pack(anchor="w", padx=10, pady=2)

        ppt_frame = ttk.LabelFrame(opt_row, text="PPT 정리 항목")
        ppt_frame.pack(side="left", fill="both", expand=True, padx=(4, 0))
        ttk.Checkbutton(
            ppt_frame, text="비표준 폰트를 기본 폰트로 치환",
            variable=self.opt_replace_fonts,
        ).pack(anchor="w", padx=10, pady=2)
        ttk.Checkbutton(
            ppt_frame, text="임베드(끼워넣은) 폰트 제거",
            variable=self.opt_remove_embedded,
        ).pack(anchor="w", padx=10, pady=2)
        font_row = ttk.Frame(ppt_frame)
        font_row.pack(anchor="w", fill="x", padx=10, pady=2)
        ttk.Label(font_row, text="대치 기본 폰트:").pack(side="left")
        ttk.Entry(font_row, textvariable=self.font_var, width=18).pack(
            side="left", padx=(4, 0))

        # --- 저장 방식 (공통) ---
        save_frame = ttk.LabelFrame(self.root, text="저장 방식 (엑셀·PPT 공통)")
        save_frame.pack(fill="x", padx=10, pady=4)
        ttk.Radiobutton(
            save_frame,
            text="새 파일로 저장 (원본 보존, '<이름>_정리됨' 으로 저장)",
            variable=self.opt_overwrite, value=False,
        ).pack(anchor="w", padx=10, pady=2)
        ttk.Radiobutton(
            save_frame,
            text="원본 덮어쓰기 (자동으로 .bak 백업 생성)",
            variable=self.opt_overwrite, value=True,
        ).pack(anchor="w", padx=10, pady=2)

        # --- 실행 버튼 ---
        self.run_btn = ttk.Button(self.root, text="정리 실행", command=self.run)
        self.run_btn.pack(fill="x", padx=10, pady=(8, 4))

        # --- 결과 로그 ---
        log_frame = ttk.LabelFrame(self.root, text="결과")
        log_frame.pack(fill="both", expand=True, padx=10, pady=(4, 10))
        self.log = tk.Text(log_frame, height=8, wrap="word", state="disabled")
        self.log.pack(side="left", fill="both", expand=True, padx=(8, 0), pady=8)
        lsb = ttk.Scrollbar(log_frame, orient="vertical", command=self.log.yview)
        lsb.pack(side="left", fill="y", pady=8)
        self.log.config(yscrollcommand=lsb.set)

    # -------------------------------------------------------------- actions
    def add_files(self):
        paths = filedialog.askopenfilenames(title="파일 선택", filetypes=FILETYPES)
        self._add_paths(paths)

    def _on_drop(self, event):
        """드래그 앤 드롭으로 들어온 파일 경로 처리."""
        # tk.splitlist 가 중괄호({})로 묶인 공백 포함 경로까지 올바로 분리해 준다.
        paths = self.root.tk.splitlist(event.data)
        self._add_paths(paths)

    def _add_paths(self, paths):
        """경로 목록을 지원 형식만 걸러 중복 없이 추가한다."""
        skipped = []
        for p in paths:
            p = os.path.normpath(p)
            ext = os.path.splitext(p)[1].lower()
            if ext not in SUPPORTED_EXTS:
                skipped.append(os.path.basename(p))
                continue
            if p not in self.files:
                self.files.append(p)
                self.listbox.insert(tk.END, p)
        if skipped:
            messagebox.showwarning(
                APP_TITLE,
                "지원하지 않는 형식이라 제외했습니다:\n"
                + "\n".join(skipped)
                + f"\n\n지원 형식: {SUPPORTED_DISPLAY}",
            )

    def remove_selected(self):
        for idx in reversed(self.listbox.curselection()):
            self.listbox.delete(idx)
            del self.files[idx]

    def clear_files(self):
        self.listbox.delete(0, tk.END)
        self.files.clear()

    def _log(self, text: str):
        self.log.config(state="normal")
        self.log.insert(tk.END, text + "\n")
        self.log.see(tk.END)
        self.log.config(state="disabled")
        self.root.update_idletasks()

    def run(self):
        if not self.files:
            messagebox.showwarning(APP_TITLE, "먼저 정리할 파일을 추가해 주세요.")
            return

        self.run_btn.config(state="disabled")
        self.log.config(state="normal")
        self.log.delete("1.0", tk.END)
        self.log.config(state="disabled")

        # UI 멈춤 방지를 위해 별도 스레드에서 처리
        threading.Thread(target=self._run_worker, daemon=True).start()

    def _run_worker(self):
        files = list(self.files)
        font = self.font_var.get().strip() or DEFAULT_FONT
        ok = 0
        fail = 0
        self._log(f"총 {len(files)}개 파일 처리 시작…\n")
        for p in files:
            r = clean_file(
                p,
                overwrite=self.opt_overwrite.get(),
                backup=True,
                # 엑셀
                delete_names=self.opt_names.get(),
                remove_external_links=self.opt_links.get(),
                unhide_sheets=self.opt_unhide.get(),
                keep_print_areas=self.opt_keep_print.get(),
                # PPT
                default_font=font,
                replace_fonts=self.opt_replace_fonts.get(),
                remove_embedded=self.opt_remove_embedded.get(),
            )
            self._log(r.summary())
            if r.ok and os.path.abspath(r.dst_path) != os.path.abspath(r.src_path):
                self._log(f"        → 저장: {r.dst_path}")
            for w in r.warnings:
                self._log(f"        ⚠ {w}")
            if r.ok:
                ok += 1
            else:
                fail += 1

        self._log(f"\n완료: 성공 {ok}개, 실패 {fail}개")
        self.run_btn.config(state="normal")
        if fail == 0:
            messagebox.showinfo(APP_TITLE, f"{ok}개 파일을 모두 정리했습니다.")
        else:
            messagebox.showwarning(
                APP_TITLE, f"성공 {ok}개 / 실패 {fail}개. 결과 창을 확인해 주세요."
            )


def main():
    # 드래그 앤 드롭 지원 시 전용 Tk 루트를 사용한다.
    root = TkinterDnD.Tk() if DND_AVAILABLE else tk.Tk()
    # ttk 테마(가능하면 보기 좋은 테마 사용)
    try:
        style = ttk.Style()
        if "vista" in style.theme_names():
            style.theme_use("vista")
        elif "clam" in style.theme_names():
            style.theme_use("clam")
    except Exception:
        pass
    FileRenewApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
