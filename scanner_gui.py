"""FILE SCANNER GUI 애플리케이션."""

from __future__ import annotations

from concurrent.futures import ThreadPoolExecutor, as_completed
import os
import queue
import subprocess
import string
import threading
import time
import tkinter as tk
from datetime import datetime
from tkinter import filedialog, messagebox, ttk
from typing import Callable

import customtkinter as ctk

import scanner_engine


TARGET_EXTENSIONS: set[str] = {
    ".pdf", ".xlsx", ".xls", ".docx", ".doc",
    ".hwp", ".hwpx", ".csv", ".txt",
}

DEFAULT_KEYWORDS: list[str] = [
    "접대", "상품권", "선물", "골프", "대외비", "평가", "교수",
    "영업비", "할인", "무상", "컴프", "COMP", "민원", "보상",
    "대여", "심의", "합의서", "검찰", "경찰", "국세청", "세무서",
    "공정위", "감사원",
]

_HELP_TEXTS: dict[str, str] = {
    "파일/문서": (
        "■ 파일/문서 검색\n\n"
        "1. 검색 대상: 드라이브 또는 폴더를 선택합니다\n"
        "2. 키워드 관리: 검색할 키워드를 추가/삭제합니다\n"
        "3. 검색 범위:\n"
        "   - 파일명: 파일 이름에서만 키워드 검색\n"
        "   - 문서(내용): 문서 내부 텍스트에서 키워드 검색\n"
        "4. 지원 확장자: pdf, xlsx, xls, docx, doc, hwp, hwpx, csv, txt\n\n"
        "■ 검색 결과\n"
        "- 키워드/확장자 헤더 클릭: 필터링\n"
        "- 파일명 더블클릭: 파일 열기\n"
        "- 파일경로 더블클릭: 폴더 열기\n"
        "- 우클릭: 파일열기/경로열기/경로복사\n"
        "- 컬럼 헤더 클릭: 정렬 (키워드/확장자는 필터)\n\n"
        "■ 화면 조작\n"
        "- 좌측/우측 패널 경계선 드래그: 패널 크기 조절"
    ),
}

_UI_BATCH_INTERVAL_MS = 2000
_UI_BATCH_MAX_ROWS = 20


class FileScannerApp(ctk.CTk):
    """키워드 기반 문서 검색 GUI를 제공하는 메인 윈도우."""

    def __init__(self) -> None:
        super().__init__()

        self.title("FILE SCANNER v1.2")
        try:
            self.iconbitmap(self._resource_path("app_icon.ico"))
        except Exception:
            pass
        self.geometry("1400x900")
        self.minsize(1100, 700)
        self.resizable(True, True)

        self._stop_flag = False
        self._is_searching = False
        self._search_start_time: float = 0.0
        self._progress_history: list[tuple[float, int]] = []
        self._results: list[dict] = []
        self._all_results: list[dict[str, object]] = []
        self._search_thread: threading.Thread | None = None
        self._executor: ThreadPoolExecutor | None = None
        self._fail_count = 0
        self._skip_count = 0
        self._base_summary_text = ""

        self._keyword_filter: set[str] = set()
        self._ext_filter: set[str] = set()
        self._keyword_select_all = True
        self._ext_select_all = True
        self._filter_popup: tk.Toplevel | None = None

        self._sort_state: dict[str, bool] = {}

        self._path_vars: dict[str, ctk.BooleanVar] = {}
        self._path_values: dict[str, str] = {}
        self._search_mode = tk.StringVar(value="content")

        self._ui_queue: queue.Queue[dict[str, object]] = queue.Queue()
        self._batch_after_id: str | None = None

        self._font = ctk.CTkFont(family="맑은 고딕", size=14, weight="bold")
        self._title_font = ctk.CTkFont(family="맑은 고딕", size=15, weight="bold")
        self._summary_font = ctk.CTkFont(family="맑은 고딕", size=13)
        self._drive_font = ctk.CTkFont(family="맑은 고딕", size=15, weight="bold")

        self._build_ui()
        self._load_system_drives()
        self._load_default_keywords()
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    # 기본 키워드 로드
    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

    def _load_default_keywords(self) -> None:
        """기본 키워드를 리스트박스에 로드한다."""
        for keyword in DEFAULT_KEYWORDS:
            self._insert_keyword_if_new(keyword)

    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    # 배치 UI 갱신
    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

    def _enqueue_result(self, result: dict) -> None:
        self._ui_queue.put({"type": "result", "data": result})

    def _enqueue_fail(self, file_path: str, error_message: str) -> None:
        self._ui_queue.put(
            {"type": "fail", "file_path": file_path, "error": error_message}
        )

    def _enqueue_progress(self, done: int, total: int) -> None:
        self._ui_queue.put({"type": "progress", "done": done, "total": total})

    def _start_batch_timer(self) -> None:
        self._stop_batch_timer()
        self._batch_after_id = self.after(
            _UI_BATCH_INTERVAL_MS, self._process_ui_queue
        )

    def _stop_batch_timer(self) -> None:
        if self._batch_after_id is not None:
            self.after_cancel(self._batch_after_id)
            self._batch_after_id = None

    def _process_ui_queue(self) -> None:
        """메인 스레드에서 일정 간격으로 큐를 꺼내 배치 처리한다."""
        last_progress: dict[str, object] | None = None

        count = 0
        while count < 500:
            try:
                item = self._ui_queue.get_nowait()
            except queue.Empty:
                break
            count += 1

            item_type = item.get("type")

            if item_type == "progress":
                last_progress = item
            elif item_type == "result":
                self._all_results.append(
                    self._build_tree_row(
                        str(item["data"].get("keyword", "")),
                        str(item["data"].get("file_path", "")),
                        str(item["data"].get("location", "")),
                        str(item["data"].get("context", "")).replace("\n", " ").strip(),
                    )
                )
            elif item_type == "fail":
                self._all_results.append(
                    self._build_tree_row(
                        "실패", str(item["file_path"]), "-",
                        str(item["error"]), tags=("fail",),
                    )
                )

        if last_progress is not None:
            self._update_progress(last_progress["done"], last_progress["total"])

        total = len(self._all_results)
        if self._base_summary_text:
            self.summary_label.configure(
                text=f"{self._base_summary_text} | 수집 {total:,}건",
            )

        if self._is_searching or not self._ui_queue.empty():
            self._batch_after_id = self.after(
                _UI_BATCH_INTERVAL_MS, self._process_ui_queue
            )
        else:
            self._batch_after_id = None

    def _flush_ui_queue(self) -> None:
        """검색 종료 시 큐에 남은 항목을 모두 처리하고 트리뷰에 일괄 삽입한다."""
        while True:
            try:
                item = self._ui_queue.get_nowait()
            except queue.Empty:
                break
            item_type = item.get("type")
            if item_type == "result":
                self._all_results.append(
                    self._build_tree_row(
                        str(item["data"].get("keyword", "")),
                        str(item["data"].get("file_path", "")),
                        str(item["data"].get("location", "")),
                        str(item["data"].get("context", "")).replace("\n", " ").strip(),
                    )
                )
            elif item_type == "fail":
                self._all_results.append(
                    self._build_tree_row(
                        "실패", str(item["file_path"]), "-",
                        str(item["error"]), tags=("fail",),
                    )
                )
            elif item_type == "progress":
                self._update_progress(item["done"], item["total"])

        for row in self._all_results:
            self._keyword_filter.add(str(row.get("keyword", "")))
            self._ext_filter.add(str(row.get("extension", "")))

        self._flush_insert_index = 0
        self._flush_batch_insert()

    def _flush_batch_insert(self) -> None:
        """트리뷰에 200건씩 나눠서 삽입한다."""
        batch = 200
        end = min(self._flush_insert_index + batch, len(self._all_results))
        for i in range(self._flush_insert_index, end):
            row = self._all_results[i]
            if self._row_passes_filters(row):
                self._insert_tree_row(row)
        self._flush_insert_index = end
        self._refresh_summary_text()
        if self._flush_insert_index < len(self._all_results):
            self.after(10, self._flush_batch_insert)


    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    # UI 빌드
    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

    def _build_ui(self) -> None:
        self.configure(fg_color="#f0f2f5")
        self.grid_rowconfigure(0, weight=0)
        self.grid_rowconfigure(1, weight=0)
        self.grid_rowconfigure(2, weight=1)
        self.grid_columnconfigure(0, weight=1)

        tab_bar = ctk.CTkFrame(self, fg_color="#ffffff", corner_radius=0, height=50)
        tab_bar.grid(row=0, column=0, sticky="ew")
        tab_bar.grid_propagate(False)
        tab_bar.grid_columnconfigure(0, weight=1)
        tab_bar.grid_columnconfigure(1, weight=0)
        tab_bar.grid_columnconfigure(2, weight=0)
        tab_bar.grid_columnconfigure(3, weight=0)
        tab_bar.grid_columnconfigure(4, weight=1)

        self._tab_buttons: dict[str, ctk.CTkButton] = {}
        self._tab_frames: dict[str, ctk.CTkFrame] = {}
        self._current_tab: str = ""

        tab_names = ["파일/문서"]
        for col, name in enumerate(tab_names):
            btn = ctk.CTkButton(
                tab_bar,
                text=name,
                font=ctk.CTkFont(family="맑은 고딕", size=17, weight="bold"),
                fg_color="transparent",
                hover_color="#e8eaed",
                text_color="#5f6368",
                corner_radius=8,
                height=50,
                width=160,
                anchor="center",
                command=lambda n=name: self._switch_tab(n),
            )
            btn.grid(row=0, column=col + 1, sticky="ns")
            self._tab_buttons[name] = btn

        help_btn = ctk.CTkButton(
            tab_bar,
            text="?",
            width=40,
            height=36,
            font=ctk.CTkFont(family="맑은 고딕", size=16, weight="bold"),
            fg_color="#e5e7eb",
            hover_color="#d1d5db",
            text_color="#374151",
            corner_radius=8,
            command=self._show_help,
        )
        help_btn.grid(row=0, column=4, padx=(0, 16), sticky="e")

        separator = ctk.CTkFrame(self, fg_color="#e5e7eb", height=1, corner_radius=0)
        separator.grid(row=1, column=0, sticky="ew")

        content_area = ctk.CTkFrame(self, fg_color="#f0f2f5", corner_radius=0)
        content_area.grid(row=2, column=0, padx=16, pady=(12, 16), sticky="nsew")
        content_area.grid_rowconfigure(0, weight=1)
        content_area.grid_columnconfigure(0, weight=1)

        tab1 = ctk.CTkFrame(content_area, fg_color="#f0f2f5", corner_radius=0)
        self._tab_frames["파일/문서"] = tab1
        self._build_keyword_tab(tab1)

        self._switch_tab("파일/문서")

    def _switch_tab(self, name: str) -> None:
        if name == self._current_tab:
            return

        if self._is_searching:
            messagebox.showwarning("경고", "검색 중에는 탭을 전환할 수 없습니다.")
            return

        for frame in self._tab_frames.values():
            frame.grid_forget()

        self._tab_frames[name].grid(row=0, column=0, sticky="nsew")

        for btn_name, btn in self._tab_buttons.items():
            if btn_name == name:
                btn.configure(
                    fg_color="#e0edff",
                    text_color="#1a56db",
                    hover_color="#d0e3ff",
                )
            else:
                btn.configure(
                    fg_color="transparent",
                    text_color="#5f6368",
                    hover_color="#e8eaed",
                )

        self._current_tab = name

    def _show_help(self) -> None:
        text = _HELP_TEXTS.get(self._current_tab, "도움말이 없습니다.")
        messagebox.showinfo(f"도움말 - {self._current_tab}", text)

    def _build_keyword_tab(self, tab: ctk.CTkFrame) -> None:
        tab.grid_rowconfigure(0, weight=1)
        tab.grid_rowconfigure(1, weight=0)
        tab.grid_columnconfigure(0, weight=1)

        paned = tk.PanedWindow(
            tab,
            orient=tk.HORIZONTAL,
            sashwidth=6,
            sashrelief="flat",
            bg="#e5e7eb",
            opaqueresize=True,
        )
        paned.grid(row=0, column=0, sticky="nsew")

        left_panel = ctk.CTkFrame(paned, width=320, fg_color="#ffffff", corner_radius=8)
        left_panel.grid_propagate(False)

        right_panel = ctk.CTkFrame(paned, fg_color="#ffffff", corner_radius=8)

        paned.add(left_panel, minsize=280, stretch="never")
        paned.add(right_panel, minsize=400, stretch="always")

        bottom_bar = ctk.CTkFrame(tab, fg_color="#ffffff", corner_radius=8)
        bottom_bar.grid(row=1, column=0, pady=(10, 0), sticky="sew")

        self._build_left_panel(left_panel)
        self._build_right_panel(right_panel)
        self._build_bottom_bar(bottom_bar)

    def _build_left_panel(self, panel: ctk.CTkFrame) -> None:
        panel.grid_columnconfigure(0, weight=1)

        title = ctk.CTkLabel(
            panel, text="검색 대상", font=self._title_font,
            anchor="w", text_color="#1f2937",
        )
        title.grid(row=0, column=0, padx=12, pady=(12, 8), sticky="ew")

        self.path_frame = ctk.CTkScrollableFrame(
            panel, width=296, height=170,
            fg_color="#f9fafb", corner_radius=6,
            scrollbar_button_color="#d1d5db",
            scrollbar_button_hover_color="#9ca3af",
        )
        self.path_frame.grid(row=1, column=0, padx=12, pady=(0, 8), sticky="ew")

        self.add_folder_button = ctk.CTkButton(
            panel, text="폴더 추가", command=self._on_add_folder,
            font=self._font, height=32,
            fg_color="#1a56db", hover_color="#1648c0", text_color="#ffffff",
            corner_radius=6,
        )
        self.add_folder_button.grid(row=2, column=0, padx=12, pady=(0, 14), sticky="ew")

        keyword_title = ctk.CTkLabel(
            panel, text="키워드 관리", font=self._title_font,
            anchor="w", text_color="#1f2937",
        )
        keyword_title.grid(row=3, column=0, padx=12, pady=(0, 8), sticky="ew")

        self.keyword_entry = ctk.CTkEntry(
            panel, font=self._font, placeholder_text="키워드 입력",
            fg_color="#f9fafb", border_color="#d1d5db", text_color="#1f2937",
            placeholder_text_color="#9ca3af", corner_radius=6, height=32,
        )
        self.keyword_entry.grid(row=4, column=0, padx=12, pady=(0, 8), sticky="ew")
        self.keyword_entry.bind("<Return>", self._on_add_keyword)

        list_container = ctk.CTkFrame(panel, fg_color="#f9fafb", corner_radius=6)
        list_container.grid(row=5, column=0, padx=12, pady=(0, 8), sticky="nsew")
        panel.grid_rowconfigure(5, weight=1)
        list_container.grid_rowconfigure(0, weight=1)
        list_container.grid_columnconfigure(0, weight=1)

        self.keyword_listbox = tk.Listbox(
            list_container, selectmode=tk.EXTENDED, exportselection=False,
            font=("맑은 고딕", 13, "bold"), activestyle="none", height=10,
            bg="#f9fafb", fg="#1f2937",
            selectbackground="#1a56db", selectforeground="#ffffff",
            highlightthickness=1, highlightbackground="#d1d5db",
            borderwidth=0, relief="flat",
        )
        self.keyword_listbox.grid(row=0, column=0, sticky="nsew")

        keyword_scroll = tk.Scrollbar(
            list_container, orient="vertical", command=self.keyword_listbox.yview,
        )
        keyword_scroll.grid(row=0, column=1, sticky="ns")
        self.keyword_listbox.config(yscrollcommand=keyword_scroll.set)

        button_row = ctk.CTkFrame(panel, fg_color="transparent")
        button_row.grid(row=6, column=0, padx=12, pady=(0, 10), sticky="ew")
        for col in range(2):
            button_row.grid_columnconfigure(col, weight=1)

        add_button = ctk.CTkButton(
            button_row, text="추가", command=self._on_add_keyword,
            font=self._font, height=30, corner_radius=6,
            fg_color="#1a56db", hover_color="#1648c0", text_color="#ffffff",
        )
        add_button.grid(row=0, column=0, padx=(0, 4), sticky="ew")

        remove_button = ctk.CTkButton(
            button_row, text="삭제", command=self._on_remove_keyword,
            font=self._font, height=30, corner_radius=6,
            fg_color="#e5e7eb", hover_color="#d1d5db", text_color="#374151",
        )
        remove_button.grid(row=0, column=1, padx=(4, 0), sticky="ew")

        search_mode_title = ctk.CTkLabel(
            panel, text="검색 범위", font=self._font,
            anchor="w", text_color="#1f2937",
        )
        search_mode_title.grid(row=7, column=0, padx=12, pady=(0, 6), sticky="ew")

        search_mode_row = ctk.CTkFrame(panel, fg_color="transparent")
        search_mode_row.grid(row=8, column=0, padx=12, pady=(0, 12), sticky="ew")
        search_mode_row.grid_columnconfigure(0, weight=1)
        search_mode_row.grid_columnconfigure(1, weight=1)

        filename_radio = ctk.CTkRadioButton(
            search_mode_row, text="파일명",
            variable=self._search_mode, value="filename",
            font=self._font, text_color="#374151",
            fg_color="#1a56db", hover_color="#d0e3ff", border_color="#9ca3af",
        )
        filename_radio.grid(row=0, column=0, padx=(0, 8), sticky="w")

        content_radio = ctk.CTkRadioButton(
            search_mode_row, text="문서(내용)",
            variable=self._search_mode, value="content",
            font=self._font, text_color="#374151",
            fg_color="#1a56db", hover_color="#d0e3ff", border_color="#9ca3af",
        )
        content_radio.grid(row=0, column=1, sticky="w")

    def _build_right_panel(self, panel: ctk.CTkFrame) -> None:
        panel.grid_rowconfigure(1, weight=1)
        panel.grid_rowconfigure(2, weight=0)
        panel.grid_columnconfigure(0, weight=1)

        title = ctk.CTkLabel(
            panel, text="검색 결과", font=self._title_font,
            anchor="w", text_color="#1f2937",
        )
        title.grid(row=0, column=0, padx=12, pady=(12, 8), sticky="ew")

        tree_frame = ctk.CTkFrame(panel, fg_color="#f9fafb", corner_radius=6)
        tree_frame.grid(row=1, column=0, padx=12, pady=(0, 8), sticky="nsew")
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)
        self._configure_treeview_style()

        self.result_tree = ttk.Treeview(
            tree_frame,
            columns=(
                "keyword", "filename", "extension",
                "filepath", "location", "context", "fullpath",
            ),
            show="headings",
            displaycolumns=(
                "keyword", "filename", "extension",
                "filepath", "location", "context",
            ),
            style="Result.Treeview",
        )
        self.result_tree.grid(row=0, column=0, sticky="nsew")

        # ── 검색 진행 오버레이 ──
        self._overlay_label = ctk.CTkLabel(
            tree_frame,
            text="",
            font=ctk.CTkFont(family="맑은 고딕", size=20, weight="bold"),
            text_color="#374151",
            fg_color="#ffffff",
            corner_radius=12,
        )

     
        self.result_tree.heading(
            "keyword", text="키워드", command=self._show_keyword_filter_popup,
        )
        self.result_tree.heading(
            "filename", text="파일명 ↕",
            command=lambda: self._sort_by_column("filename"),
        )
        self.result_tree.heading(
            "extension", text="확장자", command=self._show_ext_filter_popup,
        )
        self.result_tree.heading(
            "filepath", text="파일경로 ↕",
            command=lambda: self._sort_by_column("filepath"),
        )
        self.result_tree.heading(
            "location", text="위치 ↕",
            command=lambda: self._sort_by_column("location"),
        )
        self.result_tree.heading(
            "context", text="해당 문장 ↕",
            command=lambda: self._sort_by_column("context"),
        )
        self.result_tree.heading("fullpath", text="전체경로")

        self.result_tree.column("keyword", width=80, minwidth=80, anchor="w", stretch=False)
        self.result_tree.column("filename", width=180, minwidth=180, anchor="w", stretch=False)
        self.result_tree.column("extension", width=70, minwidth=70, anchor="w", stretch=False)
        self.result_tree.column("filepath", width=250, minwidth=250, anchor="w", stretch=False)
        self.result_tree.column("location", width=80, minwidth=80, anchor="w", stretch=False)
        self.result_tree.column("context", width=350, minwidth=220, anchor="w", stretch=True)
        self.result_tree.column("fullpath", width=0, minwidth=0, stretch=False)
        self.result_tree.tag_configure("fail", foreground="#dc2626")

        tree_v_scroll = ttk.Scrollbar(tree_frame, orient="vertical", command=self.result_tree.yview)
        tree_v_scroll.grid(row=0, column=1, sticky="ns")
        tree_h_scroll = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.result_tree.xview)
        tree_h_scroll.grid(row=1, column=0, sticky="ew")
        self.result_tree.configure(yscrollcommand=tree_v_scroll.set, xscrollcommand=tree_h_scroll.set)

        self.result_menu = tk.Menu(self, tearoff=0)
        self.result_menu.add_command(label="파일 열기", command=self._open_selected_file)
        self.result_menu.add_command(label="경로 열기", command=self._open_selected_path)
        self.result_menu.add_command(label="경로 복사", command=self._copy_selected_path)
        self.result_tree.bind("<Double-1>", self._on_tree_double_click)
        self.result_tree.bind("<Button-3>", self._on_tree_right_click)

        self.summary_label = ctk.CTkLabel(
            panel, text="", font=self._summary_font,
            anchor="w", text_color="#6b7280",
        )
        self.summary_label.grid(row=2, column=0, padx=12, pady=(0, 10), sticky="ew")

    def _build_bottom_bar(self, panel: ctk.CTkFrame) -> None:
        panel.grid_columnconfigure(1, weight=1)

        self.progress_label = ctk.CTkLabel(
            panel, text="진행: 0% (0/0파일)", font=self._font, text_color="#374151",
        )
        self.progress_label.grid(row=0, column=0, padx=(12, 8), pady=10, sticky="w")

        self.progress_bar = ctk.CTkProgressBar(
            panel, fg_color="#e5e7eb", progress_color="#1a56db",
            corner_radius=4, height=14,
        )
        self.progress_bar.grid(row=0, column=1, padx=(0, 12), pady=10, sticky="ew")
        self.progress_bar.set(0)

        self.search_button = ctk.CTkButton(
            panel, text="검색 시작", command=self._on_search_toggle,
            font=self._font, width=110, height=34, corner_radius=6,
            fg_color="#1a56db", hover_color="#1648c0", text_color="#ffffff",
        )
        self.search_button.grid(row=0, column=2, padx=(0, 8), pady=10)

        self.save_report_button = ctk.CTkButton(
            panel, text="리포트 저장(Excel)", command=self._on_save_report,
            font=self._font, width=150, height=34, corner_radius=6,
            fg_color="#059669", hover_color="#047857", text_color="#ffffff",
        )
        self.save_report_button.grid(row=0, column=3, padx=(0, 12), pady=10)

    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    # 정렬
    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

    _COLUMN_TO_KEY: dict[str, str] = {
        "keyword": "keyword",
        "filename": "filename",
        "extension": "extension",
        "filepath": "filepath_display",
        "location": "location",
        "context": "context_display",
    }

    def _sort_by_column(self, column: str) -> None:
        if self._is_searching:
            return

        ascending = not self._sort_state.get(column, False)
        self._sort_state[column] = ascending

        key_name = self._COLUMN_TO_KEY.get(column, column)
        self._all_results.sort(
            key=lambda row: str(row.get(key_name, "")).lower(),
            reverse=not ascending,
        )

        arrow = " ↑" if ascending else " ↓"
        base_names = {
            "filename": "파일명",
            "filepath": "파일경로",
            "location": "위치",
            "context": "해당 문장",
        }
        for col_name, base_text in base_names.items():
            if col_name == column:
                self.result_tree.heading(col_name, text=f"{base_text}{arrow}")
            else:
                self.result_tree.heading(col_name, text=f"{base_text} ↕")

        self._apply_filters()

    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    # 데이터 / 이벤트
    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

    def _load_system_drives(self) -> None:
        """시스템 드라이브 목록을 백그라운드에서 수집한다."""
        def _detect_drives() -> list[str]:
            drives: list[str] = []
            try:
                import ctypes
                bitmask = ctypes.windll.kernel32.GetLogicalDrives()
                for i in range(26):
                    if bitmask & (1 << i):
                        letter = chr(ord("A") + i)
                        drive = f"{letter}:\\"
                        drive_type = ctypes.windll.kernel32.GetDriveTypeW(drive)
                        if drive_type in (2, 3):
                            drives.append(drive)
            except Exception:
                for letter in string.ascii_uppercase:
                    drive = f"{letter}:\\"
                    try:
                        if os.path.isdir(drive):
                            drives.append(drive)
                    except OSError:
                        continue
            return drives

        def _on_drives_detected(drives: list[str]) -> None:
            for drive in drives:
                self._add_path_option(drive, checked=False)

        def _worker() -> None:
            drives = _detect_drives()
            try:
                self.after(0, lambda: _on_drives_detected(drives))
            except RuntimeError:
                pass

        threading.Thread(target=_worker, daemon=True).start()

    def _on_add_folder(self) -> None:
        folder = filedialog.askdirectory()
        if not folder:
            return
        self._add_path_option(folder, checked=True)

    def _add_path_option(self, path: str, checked: bool = False) -> None:
        normalized_path = os.path.abspath(path)
        canonical = self._canonical_path(normalized_path)
        if canonical in self._path_vars:
            return

        var = ctk.BooleanVar(value=checked)
        checkbox = ctk.CTkCheckBox(
            self.path_frame, text=normalized_path,
            variable=var, onvalue=True, offvalue=False,
            font=self._drive_font, text_color="#374151",
            fg_color="#1a56db", hover_color="#d0e3ff",
            border_color="#9ca3af", checkmark_color="#ffffff",
        )
        checkbox.pack(anchor="w", fill="x", padx=4, pady=2)

        self._path_vars[canonical] = var
        self._path_values[canonical] = normalized_path

    def _on_add_keyword(self, _event: tk.Event | None = None) -> str | None:
        keyword = self.keyword_entry.get().strip()
        if self._insert_keyword_if_new(keyword):
            self.keyword_entry.delete(0, tk.END)
        return "break"

    def _insert_keyword_if_new(self, keyword: str) -> bool:
        if not keyword:
            return False

        lowered = keyword.lower()
        existing = {item.lower() for item in self._get_keywords()}
        if lowered in existing:
            return False

        self.keyword_listbox.insert(tk.END, keyword)
        return True

    def _on_remove_keyword(self) -> None:
        selected = list(self.keyword_listbox.curselection())
        if not selected:
            return
        for index in reversed(selected):
            self.keyword_listbox.delete(index)

    def _on_search_toggle(self) -> None:
        if self._is_searching:
            self._request_stop_search()
            return
        self._start_search()

    def _start_search(self) -> None:
        paths = self._get_selected_paths()
        if not paths:
            messagebox.showwarning("경고", "검색 대상 드라이브/폴더를 선택하세요.")
            return

        keywords = self._get_keywords()
        if not keywords:
            messagebox.showwarning("경고", "키워드를 1개 이상 추가하세요.")
            return
        search_mode = self._search_mode.get()

        self._results = []
        self._stop_flag = False
        self._is_searching = True
        self._search_start_time = time.time()
        self._progress_history.clear()
        self._overlay_label.configure(text="파일 목록 수집 중...")
        self._overlay_label.place(relx=0.5, rely=0.4, anchor="center")
        self._overlay_label.lift()
        self._fail_count = 0
        self._skip_count = 0
        self._sort_state.clear()
        self._reset_filters()
        self._set_search_controls(is_searching=True)
        self._clear_tree_results()
        self._set_summary_text("검색 시작")
        self._update_progress(0, 0)

        while not self._ui_queue.empty():
            try:
                self._ui_queue.get_nowait()
            except queue.Empty:
                break
        self._start_batch_timer()

        self._search_thread = threading.Thread(
            target=self._search_worker,
            args=(paths, keywords, search_mode),
            daemon=True,
        )
        self._search_thread.start()

    def _request_stop_search(self) -> None:
        if self._is_searching and not self._stop_flag:
            self._stop_flag = True
            self._set_summary_text("검색 중지 요청...")
            if self._executor is not None:
                try:
                    self._executor.shutdown(wait=False, cancel_futures=True)
                except Exception:
                    pass

    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    # 검색 워커
    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

    def _search_worker(
        self, paths: list[str], keywords: list[str], search_mode: str,
    ) -> None:
        start_time = time.perf_counter()
        found_count = 0
        completed_count = 0
        total_files = 0
        files: list[str] = []

        try:
            self._call_ui(self._set_summary_text, "파일 목록 수집 중...")

            if search_mode == "filename":
                files = scanner_engine.scan_files(paths, extensions=None)
            else:
                files = scanner_engine.scan_files(
                    paths, extensions=scanner_engine.DEFAULT_EXTENSIONS,
                )
            total_files = len(files)

            if self._stop_flag:
                return

            self._call_ui(
                self._set_summary_text, f"검색 중 | 대상 파일: {total_files:,}개",
            )
            self._enqueue_progress(0, total_files)

            if search_mode == "filename":
                for index, file_path in enumerate(files, start=1):
                    if self._stop_flag:
                        break

                    completed_count = index
                    try:
                        matches = scanner_engine.search_file_by_name(
                            file_path, keywords,
                        )
                        error_message = ""
                    except Exception:
                        self._fail_count += 1
                        matches = []
                        error_message = "검색 처리 중 예외 발생"

                    if error_message:
                        self._enqueue_fail(file_path, error_message)

                    if matches:
                        for match in matches:
                            result = {
                                "keyword": match.get("keyword", ""),
                                "file_path": match.get("file", file_path),
                                "location": match.get("location", ""),
                                "context": "",
                                "searched_at": datetime.now().strftime(
                                    "%Y-%m-%d %H:%M:%S"
                                ),
                            }
                            self._results.append(result)
                            found_count += 1
                            self._enqueue_result(result)

                    if index % 500 == 0 or index == total_files:
                        self._enqueue_progress(index, total_files)
            else:
                worker_count = min(
                    scanner_engine.MAX_WORKERS,
                    max(
                        scanner_engine.MIN_WORKERS,
                        (os.cpu_count() or 4) // scanner_engine.CPU_DIVISOR,
                    ),
                )
                with ThreadPoolExecutor(max_workers=worker_count) as executor:
                    self._executor = executor
                    future_to_file = {
                        executor.submit(
                            scanner_engine.search_file, file_path, keywords,
                        ): file_path
                        for file_path in files
                    }

                    for future in as_completed(future_to_file):
                        if self._stop_flag:
                            break

                        file_path = future_to_file[future]
                        completed_count += 1

                        future_failed = False
                        try:
                            matches = future.result(
                                timeout=scanner_engine.FILE_PROCESS_TIMEOUT,
                            )
                            error_message = ""
                        except TimeoutError:
                            future_failed = True
                            future.cancel()
                            matches = []
                            error_message = f"처리 시간 초과 ({scanner_engine.FILE_PROCESS_TIMEOUT}s)"
                        except Exception:
                            future_failed = True
                            matches = []
                            error_message = "검색 처리 중 예외 발생"

                        meta = scanner_engine.consume_search_file_meta(file_path)
                        if meta.get("skipped", False):
                            self._skip_count += 1

                        if future_failed:
                            self._fail_count += 1
                        elif meta.get("failed", False):
                            self._fail_count += 1
                            if not error_message:
                                error_message = "텍스트 추출 실패"

                        if error_message:
                            self._enqueue_fail(file_path, error_message)

                        if matches:
                            for match in matches:
                                result = {
                                    "keyword": match.get("keyword", ""),
                                    "file_path": match.get("file", file_path),
                                    "location": match.get("location", ""),
                                    "context": match.get("context", ""),
                                    "searched_at": datetime.now().strftime(
                                        "%Y-%m-%d %H:%M:%S"
                                    ),
                                }
                                self._results.append(result)
                                found_count += 1
                                self._enqueue_result(result)

                        self._enqueue_progress(completed_count, total_files)

        except Exception:
            self._fail_count += 1
        finally:
            scanner_engine.clear_all_search_file_meta()
            self._executor = None
            elapsed = time.perf_counter() - start_time
            if self._stop_flag:
                summary = (
                    f"검색 중지됨 | {completed_count:,}/{total_files:,}파일 처리"
                )
            else:
                summary = (
                    f"검색 완료 | {total_files:,}개 스캔 | "
                    f"스킵 {self._skip_count:,}개 | "
                    f"{found_count:,}건 발견 | "
                    f"실패 {self._fail_count:,}개 | "
                    f"소요시간: {elapsed:.1f}s"
                )
            self._call_ui(self._on_search_worker_done, summary)

    def _on_search_worker_done(self, summary: str) -> None:
        self._is_searching = False
        self._overlay_label.place_forget()
        self._search_start_time: float = 0.0
        self._stop_batch_timer()
        self._flush_ui_queue()
        self._set_summary_text(summary)
        self._finish_search()

    def _finish_search(self) -> None:
        self._is_searching = False
        self._stop_flag = False
        self._executor = None
        self._search_thread = None
        self._set_search_controls(is_searching=False)

    def _set_search_controls(self, is_searching: bool) -> None:
        self.search_button.configure(
            text="검색 중지" if is_searching else "검색 시작",
        )
        self.save_report_button.configure(
            state="disabled" if is_searching else "normal",
        )

    def _on_save_report(self) -> None:
        if self._is_searching:
            messagebox.showwarning("경고", "검색 중에는 리포트를 저장할 수 없습니다.")
            return

        if not self._results:
            messagebox.showwarning("경고", "저장할 검색 결과가 없습니다.")
            return

        path = filedialog.asksaveasfilename(
            title="리포트 저장",
            defaultextension=".xlsx",
            filetypes=[("Excel 파일", "*.xlsx"), ("모든 파일", "*.*")],
            initialfile=f"scan_report_{datetime.now():%Y%m%d_%H%M%S}.xlsx",
        )
        if not path:
            return

        scanner_engine.save_report(self._results, path)
        self._set_summary_text(f"리포트 저장: {path}")

    def _update_progress(self, done: int, total: int) -> None:
        if total <= 0:
            ratio = 0.0
            percent = 0
            eta_text = ""
            overlay_text = ""
        else:
            ratio = min(1.0, max(0.0, done / total))
            percent = int(ratio * 100)
            now = time.time()
            self._progress_history.append((now, done))
            if len(self._progress_history) > 100:
                self._progress_history = self._progress_history[-100:]

            if len(self._progress_history) >= 2 and ratio < 1.0:
                oldest_time, oldest_done = self._progress_history[0]
                dt = now - oldest_time
                dd = done - oldest_done
                if dd > 0 and dt > 0:
                    speed = dd / dt
                    remaining = (total - done) / speed
                    if remaining >= 60:
                        m = int(remaining // 60)
                        s = int(remaining % 60)
                        eta_text = f" | 약 {m}분 {s}초 남음"
                        overlay_text = f"검색 중... {percent}%\n약 {m}분 {s}초 남음"
                    else:
                        eta_text = f" | 약 {int(remaining)}초 남음"
                        overlay_text = f"검색 중... {percent}%\n약 {int(remaining)}초 남음"
                else:
                    eta_text = " | 계산 중..."
                    overlay_text = f"검색 중... {percent}%\n계산 중..."
            elif ratio >= 1.0:
                elapsed = now - self._search_start_time
                if elapsed >= 60:
                    m = int(elapsed // 60)
                    s = int(elapsed % 60)
                    eta_text = f" | 완료 (소요 {m}분 {s}초)"
                else:
                    eta_text = f" | 완료 (소요 {int(elapsed)}초)"
                overlay_text = ""
            else:
                eta_text = " | 계산 중..."
                overlay_text = "검색 준비 중..."

        self.progress_bar.set(ratio)
        self.progress_label.configure(
            text=f"진행: {percent}% ({done:,}/{total:,}파일){eta_text}",
        )

        # 오버레이 표시/숨김
        if overlay_text:
            if self._overlay_label.cget("text") != f"  {overlay_text}  ":
                self._overlay_label.configure(text=f"  {overlay_text}  ")
            if not self._overlay_label.winfo_ismapped():
                self._overlay_label.place(relx=0.5, rely=0.4, anchor="center")
                self._overlay_label.lift()
        else:
            if self._overlay_label.winfo_ismapped():
                self._overlay_label.place_forget()

    def _get_keywords(self) -> list[str]:
        items = self.keyword_listbox.get(0, tk.END)
        return [str(item).strip() for item in items if str(item).strip()]

    def _get_selected_paths(self) -> list[str]:
        selected: list[str] = []
        for canonical, var in self._path_vars.items():
            if var.get():
                selected.append(self._path_values[canonical])
        return selected

    def _canonical_path(self, path: str) -> str:
        return os.path.normcase(os.path.normpath(path))

    def _set_summary_text(self, text: str) -> None:
        self._base_summary_text = text
        self._refresh_summary_text()

    def _refresh_summary_text(self) -> None:
        shown = len(self.result_tree.get_children())
        total = len(self._all_results)
        filter_text = f"표시 {shown:,}건 / 전체 {total:,}건"
        if self._base_summary_text:
            self.summary_label.configure(
                text=f"{self._base_summary_text} | {filter_text}",
            )
        else:
            self.summary_label.configure(text=filter_text)

    def _reset_filters(self) -> None:
        self._keyword_filter.clear()
        self._ext_filter.clear()
        self._keyword_select_all = True
        self._ext_select_all = True
        self._close_filter_popup()

    def _clear_tree_widget(self) -> None:
        for item_id in self.result_tree.get_children():
            self.result_tree.delete(item_id)

    def _clear_tree_results(self) -> None:
        self._all_results.clear()
        self._clear_tree_widget()
        self._refresh_summary_text()

    def _build_tree_row(
        self, keyword: str, file_path: str, location: str, context: str,
        tags: tuple[str, ...] = (),
    ) -> dict[str, object]:
        normalized_file_path = str(file_path)
        filename = os.path.basename(normalized_file_path) or normalized_file_path
        extension = os.path.splitext(normalized_file_path)[1].lower()
        folder_path = os.path.dirname(normalized_file_path)
        table_context = self._truncate_text(
            str(context).replace("\n", " ").strip(), 80,
        )
        return {
            "keyword": str(keyword),
            "filename": filename,
            "extension": extension,
            "filepath_display": folder_path,
            "location": str(location),
            "context_display": table_context,
            "fullpath": normalized_file_path,
            "tags": tuple(tags),
        }

    def _insert_tree_row(self, row: dict[str, object]) -> None:
        self.result_tree.insert(
            "", tk.END,
            values=(
                str(row.get("keyword", "")),
                str(row.get("filename", "")),
                str(row.get("extension", "")),
                str(row.get("filepath_display", "")),
                str(row.get("location", "")),
                str(row.get("context_display", "")),
                str(row.get("fullpath", "")),
            ),
            tags=tuple(row.get("tags", ())),
        )

    def _row_passes_filters(self, row: dict[str, object]) -> bool:
        keyword = str(row.get("keyword", ""))
        extension = str(row.get("extension", ""))
        if not self._keyword_select_all and keyword not in self._keyword_filter:
            return False
        if not self._ext_select_all and extension not in self._ext_filter:
            return False
        return True

    def _append_tree_row(self, row: dict[str, object]) -> None:
        self._all_results.append(row)
        keyword = str(row.get("keyword", ""))
        extension = str(row.get("extension", ""))
        if self._keyword_select_all:
            self._keyword_filter.add(keyword)
        if self._ext_select_all:
            self._ext_filter.add(extension)

        if self._row_passes_filters(row):
            self._insert_tree_row(row)

    def _apply_filters(self) -> None:
        self._clear_tree_widget()
        for row in self._all_results:
            if self._row_passes_filters(row):
                self._insert_tree_row(row)
        self._refresh_summary_text()

    def _get_available_filter_values(self, target: str) -> list[str]:
        values = {str(row.get(target, "")) for row in self._all_results}
        return sorted(values, key=lambda v: (v == "", v.lower()))

    def _format_filter_value(self, target: str, value: str) -> str:
        if target == "extension" and value == "":
            return "(확장자 없음)"
        return value

    def _get_filter_popup_position(self, column_name: str) -> tuple[int, int]:
        display_columns = list(self.result_tree["displaycolumns"])
        offset_x = 0
        for dc in display_columns:
            if dc == column_name:
                break
            offset_x += int(self.result_tree.column(dc, "width"))

        root_x = self.result_tree.winfo_rootx()
        root_y = self.result_tree.winfo_rooty()
        return root_x + offset_x, root_y + 28

    def _close_filter_popup(self) -> None:
        if self._filter_popup is None:
            return
        try:
            self._filter_popup.destroy()
        except tk.TclError:
            pass
        finally:
            self._filter_popup = None

    def _show_keyword_filter_popup(self) -> None:
        self._show_filter_popup(
            target="keyword", column_name="keyword", title="키워드",
        )

    def _show_ext_filter_popup(self) -> None:
        self._show_filter_popup(
            target="extension", column_name="extension", title="확장자",
        )

    def _show_filter_popup(
        self, target: str, column_name: str, title: str,
    ) -> None:
        self._close_filter_popup()

        values = self._get_available_filter_values(target)
        all_values = set(values)
        selected_values = (
            set(values)
            if (target == "keyword" and self._keyword_select_all)
            or (target == "extension" and self._ext_select_all)
            else set(
                self._keyword_filter if target == "keyword" else self._ext_filter,
            )
        )
        selected_values &= all_values

        popup = tk.Toplevel(self)
        self._filter_popup = popup
        popup.withdraw()
        popup.transient(self)
        popup.overrideredirect(True)
        popup.configure(bg="#ffffff")

        container = tk.Frame(
            popup, bg="#ffffff", bd=0, relief="flat",
            highlightthickness=1, highlightbackground="#d1d5db",
        )
        container.pack(fill="both", expand=True)

        title_label = tk.Label(
            container, text=f"{title} 필터",
            bg="#ffffff", fg="#1f2937",
            font=("맑은 고딕", 12, "bold"),
            anchor="w", padx=10, pady=8,
        )
        title_label.pack(fill="x")

        all_var = tk.BooleanVar(value=selected_values == all_values)
        item_vars: dict[str, tk.BooleanVar] = {}

        def apply_selection() -> None:
            chosen = {v for v, var in item_vars.items() if var.get()}
            all_selected = chosen == all_values
            if target == "keyword":
                self._keyword_filter = chosen
                self._keyword_select_all = all_selected
            else:
                self._ext_filter = chosen
                self._ext_select_all = all_selected
            self._apply_filters()

        def on_toggle_all() -> None:
            selected = all_var.get()
            for var in item_vars.values():
                var.set(selected)
            apply_selection()

        def on_toggle_item() -> None:
            all_var.set(
                all(v.get() for v in item_vars.values()) if item_vars else False,
            )
            apply_selection()

        all_checkbox = tk.Checkbutton(
            container, text="전체", variable=all_var,
            command=on_toggle_all,
            bg="#ffffff", fg="#1f2937",
            font=("맑은 고딕", 12), selectcolor="#ffffff",
            activebackground="#f3f4f6", activeforeground="#1f2937",
            anchor="w", padx=10, pady=3,
            relief="flat", highlightthickness=0,
        )
        all_checkbox.pack(fill="x")

        if values:
            for value in values:
                item_var = tk.BooleanVar(value=value in selected_values)
                item_vars[value] = item_var
                checkbox = tk.Checkbutton(
                    container,
                    text=self._format_filter_value(target, value),
                    variable=item_var, command=on_toggle_item,
                    bg="#ffffff", fg="#374151",
                    font=("맑은 고딕", 12), selectcolor="#ffffff",
                    activebackground="#f3f4f6", activeforeground="#374151",
                    anchor="w", padx=22, pady=3,
                    relief="flat", highlightthickness=0,
                )
                checkbox.pack(fill="x")
        else:
            empty_label = tk.Label(
                container, text="필터 대상이 없습니다",
                bg="#ffffff", fg="#9ca3af",
                font=("맑은 고딕", 11),
                anchor="w", padx=10, pady=8,
            )
            empty_label.pack(fill="x")
            all_checkbox.configure(state="disabled")

        def _on_popup_focus_out(event: tk.Event) -> None:
            try:
                focus_widget = popup.focus_get()
                if focus_widget is not None:
                    widget_path = str(focus_widget)
                    popup_path = str(popup)
                    if widget_path == popup_path or widget_path.startswith(
                        popup_path + "."
                    ):
                        return
            except (tk.TclError, KeyError):
                pass
            self._close_filter_popup()

        popup.bind("<FocusOut>", _on_popup_focus_out)
        popup.bind("<Escape>", lambda _e: self._close_filter_popup())

        x_pos, y_pos = self._get_filter_popup_position(column_name)
        popup.update_idletasks()
        popup.geometry(f"+{x_pos}+{y_pos}")
        popup.deiconify()
        popup.lift()
        popup.focus_force()

    def _insert_tree_result_from_result(self, result: dict) -> None:
        keyword = str(result.get("keyword", ""))
        file_path = str(result.get("file_path", ""))
        location = str(result.get("location", ""))
        context = str(result.get("context", "")).replace("\n", " ").strip()
        self._insert_tree_result(keyword, file_path, location, context)

    def _insert_tree_result(
        self, keyword: str, file_path: str, location: str, context: str,
    ) -> None:
        row = self._build_tree_row(keyword, file_path, location, context)
        self._append_tree_row(row)

    def _insert_tree_fail(self, file_path: str, error_message: str) -> None:
        row = self._build_tree_row(
            "실패", file_path, "-", error_message, tags=("fail",),
        )
        self._append_tree_row(row)

    def _on_tree_double_click(self, event: tk.Event) -> None:
        col_id = self.result_tree.identify_column(event.x)
        row_id = self.result_tree.identify_row(event.y)
        if not row_id:
            return
        self.result_tree.selection_set(row_id)
        self.result_tree.focus(row_id)

        if col_id == "#2":
            self._open_file_by_item_id(row_id)
        elif col_id == "#4":
            self._open_path_by_item_id(row_id)

    def _on_tree_right_click(self, event: tk.Event) -> None:
        row_id = self.result_tree.identify_row(event.y)
        if not row_id:
            return
        self.result_tree.selection_set(row_id)
        self.result_tree.focus(row_id)
        self.result_menu.post(event.x_root, event.y_root)

    def _get_item_file_path(self, item_id: str) -> str:
        if not item_id:
            return ""
        values = self.result_tree.item(item_id, "values")
        if len(values) < 7:
            return ""
        return str(values[6])

    def _get_selected_tree_file_path(self) -> str:
        selected = self.result_tree.selection()
        if not selected:
            return ""
        return self._get_item_file_path(selected[0])

    def _open_file_by_item_id(self, item_id: str) -> None:
        file_path = (
            self._get_item_file_path(item_id)
            if item_id
            else self._get_selected_tree_file_path()
        )
        if not file_path or not os.path.exists(file_path):
            messagebox.showwarning("경고", "파일을 찾을 수 없습니다")
            return
        try:
            os.startfile(file_path)
        except Exception:
            messagebox.showwarning("경고", "파일을 찾을 수 없습니다")

    def _open_path_by_item_id(self, item_id: str) -> None:
        file_path = (
            self._get_item_file_path(item_id)
            if item_id
            else self._get_selected_tree_file_path()
        )
        if not file_path or not os.path.exists(file_path):
            messagebox.showwarning("경고", "파일을 찾을 수 없습니다")
            return
        subprocess.Popen(["explorer", "/select,", file_path])

    def _open_selected_file(self) -> None:
        selected = self.result_tree.selection()
        item_id = selected[0] if selected else ""
        self._open_file_by_item_id(item_id)

    def _open_selected_path(self) -> None:
        selected = self.result_tree.selection()
        item_id = selected[0] if selected else ""
        self._open_path_by_item_id(item_id)

    def _copy_selected_path(self) -> None:
        file_path = self._get_selected_tree_file_path()
        if not file_path:
            return
        self.clipboard_clear()
        self.clipboard_append(file_path)

    def _truncate_text(self, text: str, limit: int) -> str:
        if len(text) <= limit:
            return text
        if limit <= 3:
            return "." * limit
        return f"{text[:limit - 3]}..."

    def _configure_treeview_style(self) -> None:
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure(
            "Result.Treeview",
            background="#ffffff", foreground="#1f2937",
            fieldbackground="#ffffff", rowheight=28,
            font=("맑은 고딕", 11),
            bordercolor="#e5e7eb", lightcolor="#e5e7eb", darkcolor="#e5e7eb",
            borderwidth=1, relief="solid",
        )
        style.configure(
            "Result.Treeview.Heading",
            background="#f3f4f6", foreground="#374151",
            font=("맑은 고딕", 12, "bold"),
            bordercolor="#d1d5db", lightcolor="#d1d5db", darkcolor="#d1d5db",
            borderwidth=1, relief="raised",
        )
        style.map(
            "Result.Treeview",
            background=[("selected", "#dbeafe")],
            foreground=[("selected", "#1e40af")],
        )
        style.map(
            "Result.Treeview.Heading",
            background=[("active", "#e5e7eb")],
            foreground=[("active", "#1f2937")],
        )
        style.layout(
            "Result.Treeview", [("Treeview.treearea", {"sticky": "nswe"})],
        )

    def _call_ui(self, callback: Callable[..., None], *args: object) -> None:
        try:
            self.after(0, lambda: callback(*args))
        except RuntimeError:
            return

    @staticmethod
    def _resource_path(relative_path: str) -> str:
        import sys
        try:
            base = sys._MEIPASS
        except AttributeError:
            base = os.path.abspath(".")
        return os.path.join(base, relative_path)

    def _on_close(self) -> None:
        self._stop_flag = True
        self._stop_batch_timer()
        self._close_filter_popup()
        if self._executor is not None:
            try:
                self._executor.shutdown(wait=False, cancel_futures=True)
            except Exception:
                pass
        scanner_engine.clear_all_search_file_meta()
        self.destroy()


def main() -> None:
    """애플리케이션을 실행한다."""
    ctk.set_appearance_mode("Light")
    ctk.set_default_color_theme("blue")
    app = FileScannerApp()
    app.mainloop()


if __name__ == "__main__":
    main()