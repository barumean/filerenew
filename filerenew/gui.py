# -*- coding: utf-8 -*-
"""
gui.py
======
파일 정리기(엑셀 + PPT)의 그래픽 사용자 인터페이스(GUI).

하나의 창 안에서 **탭(Notebook)으로 [엑셀 정리기] 와 [PPT 정리기] 기능을
분리**합니다. 각 탭은 자기만의 파일 목록·정리 옵션·실행 버튼·결과 로그를
독립적으로 가집니다. 두 탭은 같은 레이아웃 골격(CleanerTab)을 공유하므로
모양과 사용법이 통일되어 있습니다.

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

from filerenew.dispatcher import DEFAULT_FONT, EXCEL_EXTS, PPT_EXTS
from filerenew.excel_cleaner import clean_workbook
from filerenew.ppt_cleaner import clean_presentation

# 드래그 앤 드롭(tkinterdnd2)이 설치돼 있으면 사용하고, 없으면 버튼 방식으로
# 자동 폴백한다. (tkinterdnd2 는 OS 파일 끌어다 놓기를 지원하는 외부 라이브러리)
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD

    DND_AVAILABLE = True
except Exception:  # noqa: BLE001
    DND_AVAILABLE = False

APP_TITLE = "파일 정리기 (File Renew)"


# ===========================================================================
# 탭 공통 골격
# ===========================================================================
class CleanerTab(ttk.Frame):
    """
    엑셀/PPT 탭이 공유하는 공통 UI 골격.

    파일 목록(추가/제거/비우기 + 드래그 앤 드롭) · 저장 방식 · 실행 버튼 ·
    결과 로그를 만들어 준다. 탭별로 다른 부분은 다음 두 가지뿐이다.
      - ``_build_options(parent)`` : 탭 고유 정리 옵션 패널
      - ``_clean_one(path, overwrite)`` : 파일 한 개 처리(→ CleanResult)
    """

    # 서브클래스에서 채운다.
    intro = ""
    supported_exts: set = set()
    filetypes: list = []

    def __init__(self, master):
        super().__init__(master)
        self.files: list[str] = []
        self.opt_overwrite = tk.BooleanVar(value=False)
        self._build_ui()

    @property
    def supported_display(self) -> str:
        return " ".join(sorted(self.supported_exts))

    # ------------------------------------------------------------------ UI
    def _build_ui(self):
        pad = {"padx": 10, "pady": 6}

        ttk.Label(self, text=self.intro, font=("", 11, "bold")).pack(
            anchor="w", **pad)
        ttk.Label(
            self,
            text=f"지원 형식: {self.supported_display}",
            foreground="#555",
        ).pack(anchor="w", padx=10)

        # --- 파일 목록 ---
        list_frame = ttk.LabelFrame(self, text="정리할 파일")
        list_frame.pack(fill="both", expand=True, padx=10, pady=4)

        if DND_AVAILABLE:
            hint = "↓ 여기로 파일을 끌어다 놓거나, '파일 추가…' 버튼을 누르세요"
        else:
            hint = "'파일 추가…' 버튼으로 파일을 선택하세요 (드래그 앤 드롭은 tkinterdnd2 설치 시 활성화)"
        ttk.Label(list_frame, text=hint, foreground="#777").pack(
            anchor="w", padx=10, pady=(6, 0))

        inner = ttk.Frame(list_frame)
        inner.pack(fill="both", expand=True)

        self.listbox = tk.Listbox(inner, selectmode=tk.EXTENDED, height=7)
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

        # --- 정리 항목(탭 고유) ---
        opt_frame = ttk.LabelFrame(self, text="정리 항목")
        opt_frame.pack(fill="x", padx=10, pady=4)
        self._build_options(opt_frame)

        # --- 저장 방식 ---
        save_frame = ttk.LabelFrame(self, text="저장 방식")
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
        self.run_btn = ttk.Button(self, text="정리 실행", command=self.run)
        self.run_btn.pack(fill="x", padx=10, pady=(8, 4))

        # --- 결과 로그 ---
        log_frame = ttk.LabelFrame(self, text="결과")
        log_frame.pack(fill="both", expand=True, padx=10, pady=(4, 10))
        self.log = tk.Text(log_frame, height=7, wrap="word", state="disabled")
        self.log.pack(side="left", fill="both", expand=True, padx=(8, 0), pady=8)
        lsb = ttk.Scrollbar(log_frame, orient="vertical", command=self.log.yview)
        lsb.pack(side="left", fill="y", pady=8)
        self.log.config(yscrollcommand=lsb.set)

    # --------------------------------------------------- 서브클래스 구현부
    def _build_options(self, parent):  # pragma: no cover - UI
        raise NotImplementedError

    def _clean_one(self, path: str, overwrite: bool):  # pragma: no cover - UI
        raise NotImplementedError

    # -------------------------------------------------------------- actions
    def add_files(self):
        paths = filedialog.askopenfilenames(title="파일 선택", filetypes=self.filetypes)
        self._add_paths(paths)

    def _on_drop(self, event):
        # tk.splitlist 가 중괄호({})로 묶인 공백 포함 경로까지 올바로 분리해 준다.
        self._add_paths(self.winfo_toplevel().tk.splitlist(event.data))

    def _add_paths(self, paths):
        skipped = []
        for p in paths:
            p = os.path.normpath(p)
            ext = os.path.splitext(p)[1].lower()
            if ext not in self.supported_exts:
                skipped.append(os.path.basename(p))
                continue
            if p not in self.files:
                self.files.append(p)
                self.listbox.insert(tk.END, p)
        if skipped:
            messagebox.showwarning(
                APP_TITLE,
                "이 탭에서 지원하지 않는 형식이라 제외했습니다:\n"
                + "\n".join(skipped)
                + f"\n\n지원 형식: {self.supported_display}",
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
        self.update_idletasks()

    def run(self):
        if not self.files:
            messagebox.showwarning(APP_TITLE, "먼저 정리할 파일을 추가해 주세요.")
            return
        self.run_btn.config(state="disabled")
        self.log.config(state="normal")
        self.log.delete("1.0", tk.END)
        self.log.config(state="disabled")
        threading.Thread(target=self._run_worker, daemon=True).start()

    def _run_worker(self):
        files = list(self.files)
        overwrite = self.opt_overwrite.get()
        ok = fail = 0
        self._log(f"총 {len(files)}개 파일 처리 시작…\n")
        for p in files:
            r = self._clean_one(p, overwrite)
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


# ===========================================================================
# 엑셀 정리기 탭
# ===========================================================================
class ExcelTab(CleanerTab):
    intro = "엑셀의 깨진 이름·외부 링크·숨겨진 시트를 한 번에 정리합니다."
    supported_exts = EXCEL_EXTS
    filetypes = [
        ("엑셀 파일", " ".join(f"*{e}" for e in sorted(EXCEL_EXTS))),
        ("모든 파일", "*.*"),
    ]

    def __init__(self, master):
        self.opt_names = tk.BooleanVar(value=True)
        self.opt_links = tk.BooleanVar(value=True)
        self.opt_unhide = tk.BooleanVar(value=True)
        self.opt_keep_print = tk.BooleanVar(value=True)
        super().__init__(master)

    def _build_options(self, parent):
        ttk.Checkbutton(
            parent, text="정의된 이름(Names) 전부 삭제 — 깨진/숨겨진 이름 포함",
            variable=self.opt_names,
        ).pack(anchor="w", padx=10, pady=2)
        ttk.Checkbutton(
            parent, text="외부 링크/연결 제거 — 다른 통합문서 참조 제거",
            variable=self.opt_links,
        ).pack(anchor="w", padx=10, pady=2)
        ttk.Checkbutton(
            parent, text="숨겨진 시트 다시 표시 — hidden/veryHidden 복구",
            variable=self.opt_unhide,
        ).pack(anchor="w", padx=10, pady=2)
        ttk.Checkbutton(
            parent, text="인쇄 영역(Print Area)은 유지 — 이름 삭제 시 인쇄 영역/제목 보존",
            variable=self.opt_keep_print,
        ).pack(anchor="w", padx=10, pady=2)

    def _clean_one(self, path, overwrite):
        return clean_workbook(
            path,
            delete_names=self.opt_names.get(),
            remove_external_links=self.opt_links.get(),
            unhide_sheets=self.opt_unhide.get(),
            keep_print_areas=self.opt_keep_print.get(),
            overwrite=overwrite,
            backup=True,
        )


# ===========================================================================
# PPT 정리기 탭
# ===========================================================================
class PptTab(CleanerTab):
    intro = "PPT의 비표준 폰트와 임베드(끼워넣은) 폰트를 기본 폰트로 정리합니다."
    supported_exts = PPT_EXTS
    filetypes = [
        ("PowerPoint 파일", " ".join(f"*{e}" for e in sorted(PPT_EXTS))),
        ("모든 파일", "*.*"),
    ]

    def __init__(self, master):
        self.opt_replace_fonts = tk.BooleanVar(value=True)
        self.opt_remove_embedded = tk.BooleanVar(value=True)
        self.font_var = tk.StringVar(value=DEFAULT_FONT)
        super().__init__(master)

    def _build_options(self, parent):
        ttk.Checkbutton(
            parent, text="비표준 폰트를 기본 폰트로 치환 — 표준/기호 폰트는 보존",
            variable=self.opt_replace_fonts,
        ).pack(anchor="w", padx=10, pady=2)
        ttk.Checkbutton(
            parent, text="임베드(끼워넣은) 폰트 제거 — ppt/fonts 및 관련 관계 정리",
            variable=self.opt_remove_embedded,
        ).pack(anchor="w", padx=10, pady=2)
        font_row = ttk.Frame(parent)
        font_row.pack(anchor="w", fill="x", padx=10, pady=2)
        ttk.Label(font_row, text="대치 기본 폰트:").pack(side="left")
        ttk.Entry(font_row, textvariable=self.font_var, width=18).pack(
            side="left", padx=(4, 0))

    def _clean_one(self, path, overwrite):
        font = self.font_var.get().strip() or DEFAULT_FONT
        return clean_presentation(
            path,
            default_font=font,
            replace_fonts=self.opt_replace_fonts.get(),
            remove_embedded=self.opt_remove_embedded.get(),
            overwrite=overwrite,
            backup=True,
        )


# ===========================================================================
# 메인 윈도우
# ===========================================================================
class FileRenewApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        root.title(APP_TITLE)
        root.geometry("720x680")
        root.minsize(640, 600)

        notebook = ttk.Notebook(root)
        notebook.pack(fill="both", expand=True, padx=8, pady=8)
        notebook.add(ExcelTab(notebook), text="엑셀 정리기")
        notebook.add(PptTab(notebook), text="PPT 정리기")


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
