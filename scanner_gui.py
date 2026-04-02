"""FILE SCANNER GUI — 검색은 별도 프로세스에서 실행."""
from __future__ import annotations

import multiprocessing
import os
import queue
import string
import threading
import time
import tkinter as tk
from datetime import datetime
from multiprocessing.connection import Connection
from tkinter import filedialog, messagebox, ttk

import customtkinter as ctk

import scanner_engine
import scanner_worker

# ──────────────────────────────────────────────
# 상수
# ──────────────────────────────────────────────

DEFAULT_KEYWORDS: list[str] = [
    "접대", "상품권", "선물", "골프", "대외비", "평가", "교수",
    "영업비", "할인", "무상", "컴프", "COMP", "민원", "보상",
    "대여", "심의", "합의서", "검찰", "경찰", "국세청", "세무서",
    "공정위", "감사원",
]

_POLL_INTERVAL_MS = 200


# ──────────────────────────────────────────────
# 결과 행 생성 (GUI 객체 불필요)
# ──────────────────────────────────────────────

def _make_row(
    keyword: str, file_path: str, location: str, context: str,
    tags: tuple[str, ...] = (),
) -> dict[str, object]:
    fp = str(file_path)
    filename = os.path.basename(fp) or fp
    ext = os.path.splitext(fp)[1].lower()
    folder = os.path.dirname(fp)
    ctx = (str(context).replace("\n", " ").strip()[:80]) if context else ""
    return {
        "keyword": str(keyword),
        "filename": filename,
        "extension": ext,
        "filepath_display": folder,
        "location": str(location),
        "context_display": ctx,
        "fullpath": fp,
        "tags": tuple(tags),
    }


# ──────────────────────────────────────────────
# 메인 윈도우
# ──────────────────────────────────────────────

class FileScannerApp(ctk.CTk):

    def __init__(self) -> None:
        super().__init__()
        self.title("FILE SCANNER v2.0")
        try:
            self.iconbitmap(self._resource_path("app_icon.ico"))
        except Exception:
            pass
        self.geometry("1400x900")
        self.minsize(1100, 700)

        # 상태
        self._is_searching = False
        self._search_proc: multiprocessing.Process | None = None
        self._conn: Connection | None = None
        self._poll_id: str | None = None

        self._all_results: list[dict[str, object]] = []
        self._results_for_report: list[dict] = []
        self._fail_count = 0
        self._skip_count = 0
        self._base_summary = ""

        self._keyword_filter: set[str] = set()
        self._ext_filter: set[str] = set()
        self._keyword_select_all = True
        self._ext_select_all = True
        self._filter_popup: tk.Toplevel | None = None
        self._sort_state: dict[str, bool] = {}

        self._path_vars: dict[str, ctk.BooleanVar] = {}
        self._path_values: dict[str, str] = {}
        self._search_mode = tk.StringVar(value="content")

        # 폰트
        self._font = ctk.CTkFont(family="맑은 고딕", size=14, weight="bold")
        self._title_font = ctk.CTkFont(family="맑은 고딕", size=15, weight="bold")
        self._summary_font = ctk.CTkFont(family="맑은 고딕", size=13)
        self._drive_font = ctk.CTkFont(family="맑은 고딕", size=15, weight="bold")

        self._build_ui()
        self._load_drives()
        self._load_default_keywords()
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    # ──────────────────────────────────────────
    # UI 빌드
    # ──────────────────────────────────────────

    def _build_ui(self) -> None:
        self.configure(fg_color="#f0f2f5")
        self.grid_rowconfigure(0, weight=0)
        self.grid_rowconfigure(1, weight=0)
        self.grid_rowconfigure(2, weight=1)
        self.grid_columnconfigure(0, weight=1)

        # 탭 바
        tab_bar = ctk.CTkFrame(self, fg_color="#ffffff", corner_radius=0, height=50)
        tab_bar.grid(row=0, column=0, sticky="ew")
        tab_bar.grid_propagate(False)
        tab_bar.grid_columnconfigure(0, weight=1)
        tab_bar.grid_columnconfigure(1, weight=0)
        tab_bar.grid_columnconfigure(2, weight=1)

        tab_btn = ctk.CTkButton(
            tab_bar, text="파일/문서",
            font=ctk.CTkFont(family="맑은 고딕", size=17, weight="bold"),
            fg_color="#e0edff", hover_color="#d0e3ff",
            text_color="#1a56db", corner_radius=8, height=50, width=160,
        )
        tab_btn.grid(row=0, column=1, sticky="ns")

        sep = ctk.CTkFrame(self, fg_color="#e5e7eb", height=1, corner_radius=0)
        sep.grid(row=1, column=0, sticky="ew")

        content = ctk.CTkFrame(self, fg_color="#f0f2f5", corner_radius=0)
        content.grid(row=2, column=0, padx=16, pady=(12, 16), sticky="nsew")
        content.grid_rowconfigure(0, weight=1)
        content.grid_rowconfigure(1, weight=0)
        content.grid_columnconfigure(0, weight=1)

        paned = tk.PanedWindow(
            content, orient=tk.HORIZONTAL, sashwidth=6,
            sashrelief="flat", bg="#e5e7eb", opaqueresize=True,
        )
        paned.grid(row=0, column=0, sticky="nsew")

        left = ctk.CTkFrame(paned, width=320, fg_color="#ffffff", corner_radius=8)
        left.grid_propagate(False)
        right = ctk.CTkFrame(paned, fg_color="#ffffff", corner_radius=8)
        paned.add(left, minsize=280, stretch="never")
        paned.add(right, minsize=400, stretch="always")

        bottom = ctk.CTkFrame(content, fg_color="#ffffff", corner_radius=8)
        bottom.grid(row=1, column=0, pady=(10, 0), sticky="sew")

        self._build_left(left)
        self._build_right(right)
        self._build_bottom(bottom)

    def _build_left(self, p: ctk.CTkFrame) -> None:
        p.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(p, text="검색 대상", font=self._title_font,
                     anchor="w", text_color="#1f2937").grid(
            row=0, column=0, padx=12, pady=(12, 8), sticky="ew")

        self.path_frame = ctk.CTkScrollableFrame(
            p, width=296, height=170, fg_color="#f9fafb", corner_radius=6,
            scrollbar_button_color="#d1d5db", scrollbar_button_hover_color="#9ca3af",
        )
        self.path_frame.grid(row=1, column=0, padx=12, pady=(0, 8), sticky="ew")

        ctk.CTkButton(
            p, text="폴더 추가", command=self._on_add_folder,
            font=self._font, height=32, fg_color="#1a56db",
            hover_color="#1648c0", text_color="#ffffff", corner_radius=6,
        ).grid(row=2, column=0, padx=12, pady=(0, 14), sticky="ew")

        ctk.CTkLabel(p, text="키워드 관리", font=self._title_font,
                     anchor="w", text_color="#1f2937").grid(
            row=3, column=0, padx=12, pady=(0, 8), sticky="ew")

        self.keyword_entry = ctk.CTkEntry(
            p, font=self._font, placeholder_text="키워드 입력",
            fg_color="#f9fafb", border_color="#d1d5db", text_color="#1f2937",
            placeholder_text_color="#9ca3af", corner_radius=6, height=32,
        )
        self.keyword_entry.grid(row=4, column=0, padx=12, pady=(0, 8), sticky="ew")
        self.keyword_entry.bind("<Return>", self._on_add_keyword)

        lc = ctk.CTkFrame(p, fg_color="#f9fafb", corner_radius=6)
        lc.grid(row=5, column=0, padx=12, pady=(0, 8), sticky="nsew")
        p.grid_rowconfigure(5, weight=1)
        lc.grid_rowconfigure(0, weight=1)
        lc.grid_columnconfigure(0, weight=1)

        self.keyword_listbox = tk.Listbox(
            lc, selectmode=tk.EXTENDED, exportselection=False,
            font=("맑은 고딕", 13, "bold"), activestyle="none", height=10,
            bg="#f9fafb", fg="#1f2937", selectbackground="#1a56db",
            selectforeground="#ffffff", highlightthickness=1,
            highlightbackground="#d1d5db", borderwidth=0, relief="flat",
        )
        self.keyword_listbox.grid(row=0, column=0, sticky="nsew")
        sc = tk.Scrollbar(lc, orient="vertical", command=self.keyword_listbox.yview)
        sc.grid(row=0, column=1, sticky="ns")
        self.keyword_listbox.config(yscrollcommand=sc.set)

        br = ctk.CTkFrame(p, fg_color="transparent")
        br.grid(row=6, column=0, padx=12, pady=(0, 10), sticky="ew")
        br.grid_columnconfigure(0, weight=1)
        br.grid_columnconfigure(1, weight=1)

        ctk.CTkButton(
            br, text="추가", command=self._on_add_keyword,
            font=self._font, height=30, corner_radius=6,
            fg_color="#1a56db", hover_color="#1648c0", text_color="#ffffff",
        ).grid(row=0, column=0, padx=(0, 4), sticky="ew")

        ctk.CTkButton(
            br, text="삭제", command=self._on_remove_keyword,
            font=self._font, height=30, corner_radius=6,
            fg_color="#e5e7eb", hover_color="#d1d5db", text_color="#374151",
        ).grid(row=0, column=1, padx=(4, 0), sticky="ew")

        ctk.CTkLabel(p, text="검색 범위", font=self._font,
                     anchor="w", text_color="#1f2937").grid(
            row=7, column=0, padx=12, pady=(0, 6), sticky="ew")

        sr = ctk.CTkFrame(p, fg_color="transparent")
        sr.grid(row=8, column=0, padx=12, pady=(0, 12), sticky="ew")
        sr.grid_columnconfigure(0, weight=1)
        sr.grid_columnconfigure(1, weight=1)

        ctk.CTkRadioButton(
            sr, text="파일명", variable=self._search_mode, value="filename",
            font=self._font, text_color="#374151", fg_color="#1a56db",
            hover_color="#d0e3ff", border_color="#9ca3af",
        ).grid(row=0, column=0, padx=(0, 8), sticky="w")

        ctk.CTkRadioButton(
            sr, text="문서(내용)", variable=self._search_mode, value="content",
            font=self._font, text_color="#374151", fg_color="#1a56db",
            hover_color="#d0e3ff", border_color="#9ca3af",
        ).grid(row=0, column=1, sticky="w")

    def _build_right(self, p: ctk.CTkFrame) -> None:
        p.grid_rowconfigure(1, weight=1)
        p.grid_rowconfigure(2, weight=0)
        p.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(p, text="검색 결과", font=self._title_font,
                     anchor="w", text_color="#1f2937").grid(
            row=0, column=0, padx=12, pady=(12, 8), sticky="ew")

        tf = ctk.CTkFrame(p, fg_color="#f9fafb", corner_radius=6)
        tf.grid(row=1, column=0, padx=12, pady=(0, 8), sticky="nsew")
        tf.grid_rowconfigure(0, weight=1)
        tf.grid_columnconfigure(0, weight=1)
        self._style_tree()

        self.tree = ttk.Treeview(
            tf,
            columns=("keyword", "filename", "extension",
                     "filepath", "location", "context", "fullpath"),
            show="headings",
            displaycolumns=("keyword", "filename", "extension",
                           "filepath", "location", "context"),
            style="R.Treeview",
        )
        self.tree.grid(row=0, column=0, sticky="nsew")

        cols = [
            ("keyword", "키워드", 80, False, self._show_kw_filter),
            ("filename", "파일명 ↕", 180, False, lambda: self._sort("filename")),
            ("extension", "확장자", 70, False, self._show_ext_filter),
            ("filepath", "파일경로 ↕", 250, False, lambda: self._sort("filepath")),
            ("location", "위치 ↕", 80, False, lambda: self._sort("location")),
            ("context", "해당 문장 ↕", 350, True, lambda: self._sort("context")),
        ]
        for cid, text, w, stretch, cmd in cols:
            self.tree.heading(cid, text=text, command=cmd)
            self.tree.column(cid, width=w, minwidth=w, anchor="w", stretch=stretch)
        self.tree.heading("fullpath", text="전체경로")
        self.tree.column("fullpath", width=0, minwidth=0, stretch=False)
        self.tree.tag_configure("fail", foreground="#dc2626")

        vs = ttk.Scrollbar(tf, orient="vertical", command=self.tree.yview)
        vs.grid(row=0, column=1, sticky="ns")
        hs = ttk.Scrollbar(tf, orient="horizontal", command=self.tree.xview)
        hs.grid(row=1, column=0, sticky="ew")
        self.tree.configure(yscrollcommand=vs.set, xscrollcommand=hs.set)

        self.ctx_menu = tk.Menu(self, tearoff=0)
        self.ctx_menu.add_command(label="파일 열기", command=self._open_file)
        self.ctx_menu.add_command(label="경로 열기", command=self._open_folder)
        self.ctx_menu.add_command(label="경로 복사", command=self._copy_path)
        self.tree.bind("<Double-1>", self._on_dbl_click)
        self.tree.bind("<Button-3>", self._on_right_click)

        self.summary_lbl = ctk.CTkLabel(
            p, text="", font=self._summary_font, anchor="w", text_color="#6b7280",
        )
        self.summary_lbl.grid(row=2, column=0, padx=12, pady=(0, 10), sticky="ew")

    def _build_bottom(self, p: ctk.CTkFrame) -> None:
        p.grid_columnconfigure(1, weight=1)

        self.prog_lbl = ctk.CTkLabel(
            p, text="진행: 0%", font=self._font, text_color="#374151",
        )
        self.prog_lbl.grid(row=0, column=0, padx=(12, 8), pady=10, sticky="w")

        self.prog_bar = ctk.CTkProgressBar(
            p, fg_color="#e5e7eb", progress_color="#1a56db", corner_radius=4, height=14,
        )
        self.prog_bar.grid(row=0, column=1, padx=(0, 12), pady=10, sticky="ew")
        self.prog_bar.set(0)

        self.search_btn = ctk.CTkButton(
            p, text="검색 시작", command=self._on_search_toggle,
            font=self._font, width=110, height=34, corner_radius=6,
            fg_color="#1a56db", hover_color="#1648c0", text_color="#ffffff",
        )
        self.search_btn.grid(row=0, column=2, padx=(0, 8), pady=10)

        self.save_btn = ctk.CTkButton(
            p, text="리포트 저장(Excel)", command=self._on_save,
            font=self._font, width=150, height=34, corner_radius=6,
            fg_color="#059669", hover_color="#047857", text_color="#ffffff",
        )
        self.save_btn.grid(row=0, column=3, padx=(0, 12), pady=10)

    # ──────────────────────────────────────────
    # 드라이브 / 키워드
    # ──────────────────────────────────────────

    def _load_drives(self) -> None:
        def work():
            drives = []
            try:
                import ctypes
                bitmask = ctypes.windll.kernel32.GetLogicalDrives()
                for i in range(26):
                    if bitmask & (1 << i):
                        letter = chr(65 + i)
                        d = f"{letter}:\\"
                        if ctypes.windll.kernel32.GetDriveTypeW(d) in (2, 3):
                            drives.append(d)
            except Exception:
                for c in string.ascii_uppercase:
                    d = f"{c}:\\"
                    if os.path.isdir(d):
                        drives.append(d)
            try:
                self.after(0, lambda: [self._add_path(d) for d in drives])
            except RuntimeError:
                pass
        threading.Thread(target=work, daemon=True).start()

    def _load_default_keywords(self) -> None:
        for kw in DEFAULT_KEYWORDS:
            self._insert_kw(kw)

    def _on_add_folder(self) -> None:
        f = filedialog.askdirectory()
        if f:
            self._add_path(f, checked=True)

    def _add_path(self, path: str, checked: bool = False) -> None:
        norm = os.path.abspath(path)
        key = os.path.normcase(os.path.normpath(norm))
        if key in self._path_vars:
            return
        var = ctk.BooleanVar(value=checked)
        ctk.CTkCheckBox(
            self.path_frame, text=norm, variable=var,
            font=self._drive_font, text_color="#374151",
            fg_color="#1a56db", hover_color="#d0e3ff",
            border_color="#9ca3af", checkmark_color="#ffffff",
        ).pack(anchor="w", fill="x", padx=4, pady=2)
        self._path_vars[key] = var
        self._path_values[key] = norm

    def _on_add_keyword(self, _e=None) -> str | None:
        kw = self.keyword_entry.get().strip()
        if self._insert_kw(kw):
            self.keyword_entry.delete(0, tk.END)
        return "break"

    def _insert_kw(self, kw: str) -> bool:
        if not kw:
            return False
        existing = {self.keyword_listbox.get(i).lower()
                    for i in range(self.keyword_listbox.size())}
        if kw.lower() in existing:
            return False
        self.keyword_listbox.insert(tk.END, kw)
        return True

    def _on_remove_keyword(self) -> None:
        for i in reversed(self.keyword_listbox.curselection()):
            self.keyword_listbox.delete(i)

    def _get_keywords(self) -> list[str]:
        return [self.keyword_listbox.get(i).strip()
                for i in range(self.keyword_listbox.size())
                if self.keyword_listbox.get(i).strip()]

    def _get_paths(self) -> list[str]:
        return [self._path_values[k] for k, v in self._path_vars.items() if v.get()]

    # ──────────────────────────────────────────
    # 검색 시작 / 중지
    # ──────────────────────────────────────────

    def _on_search_toggle(self) -> None:
        if self._is_searching:
            self._stop_search()
        else:
            self._start_search()

    def _start_search(self) -> None:
        paths = self._get_paths()
        if not paths:
            messagebox.showwarning("경고", "검색 대상을 선택하세요.")
            return
        keywords = self._get_keywords()
        if not keywords:
            messagebox.showwarning("경고", "키워드를 추가하세요.")
            return

        self._is_searching = True
        self._all_results.clear()
        self._results_for_report.clear()
        self._fail_count = 0
        self._skip_count = 0
        self._sort_state.clear()
        self._reset_filters()
        self._clear_tree()
        self._set_summary("검색 시작")
        self._set_progress(0, 0)
        self.search_btn.configure(text="검색 중지")
        self.save_btn.configure(state="disabled")

        parent_conn, child_conn = multiprocessing.Pipe()
        self._conn = parent_conn
        self._search_proc = multiprocessing.Process(
            target=scanner_worker.run_search,
            args=(child_conn, paths, keywords, self._search_mode.get()),
            daemon=True,
        )
        self._search_proc.start()
        child_conn.close()

        self._poll_id = self.after(_POLL_INTERVAL_MS, self._poll)

    def _stop_search(self) -> None:
        if self._conn:
            try:
                self._conn.send(scanner_worker.MSG_STOP)
            except (BrokenPipeError, OSError):
                pass
        self._set_summary("검색 중지 요청...")

    # ──────────────────────────────────────────
    # 폴링 — 별도 프로세스에서 결과 수신
    # ──────────────────────────────────────────

    def _poll(self) -> None:
        if self._conn is None:
            return

        count = 0
        done_received = False

        try:
            while self._conn.poll(0) and count < 200:
                msg = self._conn.recv()
                count += 1
                t = msg.get("type")

                if t == scanner_worker.MSG_PROGRESS:
                    self._set_progress(msg["done"], msg["total"])

                elif t == scanner_worker.MSG_RESULT:
                    row = _make_row(
                        msg.get("keyword", ""),
                        msg.get("file_path", ""),
                        msg.get("location", ""),
                        msg.get("context", ""),
                    )
                    self._all_results.append(row)
                    self._results_for_report.append({
                        "keyword": msg.get("keyword", ""),
                        "file_path": msg.get("file_path", ""),
                        "location": msg.get("location", ""),
                        "context": msg.get("context", ""),
                        "searched_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    })

                elif t == scanner_worker.MSG_FAIL:
                    row = _make_row(
                        "실패", msg.get("file_path", ""), "-",
                        msg.get("error", ""), tags=("fail",),
                    )
                    self._all_results.append(row)
                    self._fail_count += 1

                elif t == scanner_worker.MSG_DONE:
                    done_received = True
                    break

        except (EOFError, OSError):
            done_received = True

        # 건수 갱신
        total = len(self._all_results)
        if self._base_summary:
            self.summary_lbl.configure(
                text=f"{self._base_summary} | 수집 {total:,}건"
            )

        if done_received:
            self._finish_search()
        else:
            self._poll_id = self.after(_POLL_INTERVAL_MS, self._poll)

    def _finish_search(self) -> None:
        self._is_searching = False
        if self._conn:
            try:
                self._conn.close()
            except Exception:
                pass
            self._conn = None
        if self._search_proc and self._search_proc.is_alive():
            self._search_proc.join(timeout=2)
            if self._search_proc.is_alive():
                self._search_proc.kill()
        self._search_proc = None

        total = len(self._all_results)
        found = sum(1 for r in self._all_results if r.get("tags") != ("fail",))
        self._set_summary(
            f"검색 완료 | {found:,}건 발견 | 실패 {self._fail_count:,}개"
        )
        self.search_btn.configure(text="검색 시작")
        self.save_btn.configure(state="normal")

        # 필터 수집
        for row in self._all_results:
            self._keyword_filter.add(str(row.get("keyword", "")))
            self._ext_filter.add(str(row.get("extension", "")))

        # 트리뷰 배치 삽입
        self._flush_idx = 0
        self._flush_batch()

    def _flush_batch(self) -> None:
        end = min(self._flush_idx + 500, len(self._all_results))
        for i in range(self._flush_idx, end):
            row = self._all_results[i]
            if self._passes_filter(row):
                self._insert_row(row)
        self._flush_idx = end
        if self._flush_idx < len(self._all_results):
            self.after(1, self._flush_batch)
        else:
            self._refresh_summary()

    # ──────────────────────────────────────────
    # 진행률 / 요약
    # ──────────────────────────────────────────

    def _set_progress(self, done: int, total: int) -> None:
        if total <= 0:
            self.prog_bar.set(0)
            self.prog_lbl.configure(text="진행: 0%")
        else:
            r = min(1.0, done / total)
            self.prog_bar.set(r)
            self.prog_lbl.configure(
                text=f"진행: {int(r * 100)}% ({done:,}/{total:,}파일)"
            )

    def _set_summary(self, text: str) -> None:
        self._base_summary = text
        self._refresh_summary()

    def _refresh_summary(self) -> None:
        shown = len(self.tree.get_children())
        total = len(self._all_results)
        ft = f"표시 {shown:,}건 / 전체 {total:,}건"
        if self._base_summary:
            self.summary_lbl.configure(text=f"{self._base_summary} | {ft}")
        else:
            self.summary_lbl.configure(text=ft)

    # ──────────────────────────────────────────
    # 트리뷰
    # ──────────────────────────────────────────

    def _insert_row(self, row: dict) -> None:
        self.tree.insert("", tk.END, values=(
            str(row.get("keyword", "")),
            str(row.get("filename", "")),
            str(row.get("extension", "")),
            str(row.get("filepath_display", "")),
            str(row.get("location", "")),
            str(row.get("context_display", "")),
            str(row.get("fullpath", "")),
        ), tags=row.get("tags", ()))

    def _clear_tree(self) -> None:
        for c in self.tree.get_children():
            self.tree.delete(c)

    def _passes_filter(self, row: dict) -> bool:
        if self._keyword_filter:
            if not self._keyword_select_all:
                if str(row.get("keyword", "")) not in self._keyword_filter:
                    return False
        if self._ext_filter:
            if not self._ext_select_all:
                if str(row.get("extension", "")) not in self._ext_filter:
                    return False
        return True

    def _apply_filters(self) -> None:
        self._clear_tree()
        for row in self._all_results:
            if self._passes_filter(row):
                self._insert_row(row)
        self._refresh_summary()

    def _reset_filters(self) -> None:
        self._keyword_filter.clear()
        self._ext_filter.clear()
        self._keyword_select_all = True
        self._ext_select_all = True
        if self._filter_popup:
            try:
                self._filter_popup.destroy()
            except Exception:
                pass
            self._filter_popup = None

    # ──────────────────────────────────────────
    # 정렬
    # ──────────────────────────────────────────

    _COL_KEY = {
        "filename": "filename", "filepath": "filepath_display",
        "location": "location", "context": "context_display",
    }

    def _sort(self, col: str) -> None:
        if self._is_searching:
            return
        asc = not self._sort_state.get(col, False)
        self._sort_state[col] = asc
        key = self._COL_KEY.get(col, col)
        self._all_results.sort(key=lambda r: str(r.get(key, "")).lower(), reverse=not asc)
        names = {"filename": "파일명", "filepath": "파일경로",
                 "location": "위치", "context": "해당 문장"}
        for c, base in names.items():
            arr = " ↑" if c == col and asc else " ↓" if c == col else " ↕"
            self.tree.heading(c, text=f"{base}{arr}")
        self._apply_filters()

    # ──────────────────────────────────────────
    # 필터 팝업
    # ──────────────────────────────────────────

    def _show_kw_filter(self) -> None:
        values = sorted({str(r.get("keyword", "")) for r in self._all_results})
        self._show_filter("keyword", "키워드", values)

    def _show_ext_filter(self) -> None:
        values = sorted({str(r.get("extension", "")) for r in self._all_results})
        self._show_filter("extension", "확장자", values)

    def _show_filter(self, target: str, title: str, values: list[str]) -> None:
        if self._filter_popup:
            try:
                self._filter_popup.destroy()
            except Exception:
                pass

        popup = tk.Toplevel(self)
        popup.title(f"{title} 필터")
        popup.geometry("250x350")
        popup.resizable(False, False)
        popup.configure(bg="#ffffff")
        self._filter_popup = popup

        is_kw = target == "keyword"
        all_var = tk.BooleanVar(value=True)
        item_vars: dict[str, tk.BooleanVar] = {}

        def apply():
            selected = {v for v, var in item_vars.items() if var.get()}
            if is_kw:
                self._keyword_filter = selected
                self._keyword_select_all = all_var.get()
            else:
                self._ext_filter = selected
                self._ext_select_all = all_var.get()
            self._apply_filters()

        def toggle_all():
            state = all_var.get()
            for var in item_vars.values():
                var.set(state)
            apply()

        tk.Label(popup, text=f"{title} 필터", bg="#ffffff", fg="#1f2937",
                 font=("맑은 고딕", 12, "bold")).pack(fill="x", padx=10, pady=8)

        tk.Checkbutton(popup, text="전체", variable=all_var, command=toggle_all,
                       bg="#ffffff", font=("맑은 고딕", 12)).pack(fill="x", padx=10)

        for v in values:
            iv = tk.BooleanVar(value=True)
            item_vars[v] = iv
            tk.Checkbutton(popup, text=v, variable=iv,
                           command=apply, bg="#ffffff",
                           font=("맑은 고딕", 12)).pack(fill="x", padx=22)

        popup.focus_set()
        popup.bind("<FocusOut>", lambda e: None)

    # ──────────────────────────────────────────
    # 트리뷰 이벤트
    # ──────────────────────────────────────────

    def _on_dbl_click(self, event) -> None:
        item = self.tree.identify_row(event.y)
        col = self.tree.identify_column(event.x)
        if not item:
            return
        vals = self.tree.item(item, "values")
        if not vals:
            return
        fp = vals[6]  # fullpath
        if col == "#4":  # filepath
            folder = os.path.dirname(fp)
            if os.path.isdir(folder):
                os.startfile(folder)
        else:
            if os.path.isfile(fp):
                os.startfile(fp)

    def _on_right_click(self, event) -> None:
        item = self.tree.identify_row(event.y)
        if item:
            self.tree.selection_set(item)
            self.ctx_menu.post(event.x_root, event.y_root)

    def _get_sel_path(self) -> str | None:
        sel = self.tree.selection()
        if not sel:
            return None
        vals = self.tree.item(sel[0], "values")
        return vals[6] if vals else None

    def _open_file(self) -> None:
        p = self._get_sel_path()
        if p and os.path.isfile(p):
            os.startfile(p)

    def _open_folder(self) -> None:
        p = self._get_sel_path()
        if p:
            folder = os.path.dirname(p)
            if os.path.isdir(folder):
                os.startfile(folder)

    def _copy_path(self) -> None:
        p = self._get_sel_path()
        if p:
            self.clipboard_clear()
            self.clipboard_append(p)

    # ──────────────────────────────────────────
    # 리포트 저장
    # ──────────────────────────────────────────

    def _on_save(self) -> None:
        if self._is_searching:
            messagebox.showwarning("경고", "검색 중에는 저장할 수 없습니다.")
            return
        if not self._results_for_report:
            messagebox.showwarning("경고", "저장할 결과가 없습니다.")
            return
        path = filedialog.asksaveasfilename(
            title="리포트 저장", defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            initialfile=f"scan_report_{datetime.now():%Y%m%d_%H%M%S}.xlsx",
        )
        if path:
            scanner_engine.save_report(self._results_for_report, path)
            self._set_summary(f"리포트 저장: {path}")

    # ──────────────────────────────────────────
    # 스타일 / 유틸
    # ──────────────────────────────────────────

    def _style_tree(self) -> None:
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except Exception:
            pass
        style.configure("R.Treeview", background="#ffffff", foreground="#1f2937",
                        fieldbackground="#ffffff", rowheight=28,
                        font=("맑은 고딕", 11), borderwidth=1, relief="solid")
        style.configure("R.Treeview.Heading", background="#f3f4f6",
                        foreground="#374151", font=("맑은 고딕", 12, "bold"),
                        borderwidth=1, relief="raised")
        style.map("R.Treeview",
                  background=[("selected", "#dbeafe")],
                  foreground=[("selected", "#1e40af")])

    @staticmethod
    def _resource_path(rel: str) -> str:
        import sys
        try:
            base = sys._MEIPASS
        except AttributeError:
            base = os.path.dirname(os.path.abspath(__file__))
        return os.path.join(base, rel)

    def _on_close(self) -> None:
        if self._is_searching:
            self._stop_search()
            self.after(500, self._on_close)
            return
        if self._poll_id:
            self.after_cancel(self._poll_id)
        self.destroy()


def main() -> None:
    ctk.set_appearance_mode("light")
    ctk.set_default_color_theme("blue")
    app = FileScannerApp()
    app.mainloop()


if __name__ == "__main__":
    multiprocessing.freeze_support()
    main()
