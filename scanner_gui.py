"""FILE SCANNER GUI 애플리케이션."""

from __future__ import annotations

from concurrent.futures import ThreadPoolExecutor, TimeoutError as FuturesTimeoutError, as_completed
import os
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
    ".pdf",
    ".xlsx",
    ".xls",
    ".docx",
    ".doc",
    ".hwp",
    ".hwpx",
    ".csv",
    ".txt",
}


class FileScannerApp(ctk.CTk):
    """키워드 기반 문서 검색 GUI를 제공하는 메인 윈도우."""

    def __init__(self) -> None:
        super().__init__()

        self.title("FILE SCANNER v1.0")
        self.geometry("1400x900")
        self.minsize(1100, 700)
        self.resizable(True, True)

        self._stop_flag = False
        self._is_searching = False
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

        self._path_vars: dict[str, ctk.BooleanVar] = {}
        self._path_values: dict[str, str] = {}
        self._search_mode = tk.StringVar(value="content")

        self._font = ctk.CTkFont(family="맑은 고딕", size=13)
        self._title_font = ctk.CTkFont(family="맑은 고딕", size=15, weight="bold")
        self._tab_font = ctk.CTkFont(family="맑은 고딕", size=16)
        self._summary_font = ctk.CTkFont(family="맑은 고딕", size=13)

        self._build_ui()
        self._load_system_drives()
        self._load_keywords()
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    def _build_ui(self) -> None:
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        self.tabview = ctk.CTkTabview(
            self,
            segmented_button_fg_color="#1a1d22",
            segmented_button_selected_color="#1f6aa5",
            segmented_button_selected_hover_color="#2980b9",
            segmented_button_unselected_color="#2b2f38",
            segmented_button_unselected_hover_color="#3a3f4b",
            corner_radius=30,
        )
        self.tabview.grid(row=0, column=0, padx=12, pady=(0, 12), sticky="nsew")

        self.keyword_tab = self.tabview.add("파일/문서")
        self.pi_tab = self.tabview.add("개인정보")
        self.illegal_sw_tab = self.tabview.add("불법SW파일")

        self.tabview.configure(anchor="center", border_width=0)
        self.tabview._segmented_button.configure(
            font=ctk.CTkFont(family="맑은 고딕", size=18, weight="bold"),
            dynamic_resizing=False,
            width=750,
            height=42,
            corner_radius=6,
        )
        for tab_button in self.tabview._segmented_button._buttons_dict.values():
            tab_button.configure(anchor="center", width=240, height=42)
        self.tabview._segmented_button.grid_configure(pady=(10, 0))

        self._build_keyword_tab(self.keyword_tab)
        self._build_placeholder_tab(self.pi_tab)
        self._build_placeholder_tab(self.illegal_sw_tab)

    def _build_placeholder_tab(self, tab: ctk.CTkFrame) -> None:
        tab.grid_rowconfigure(0, weight=1)
        tab.grid_columnconfigure(0, weight=1)
        label = ctk.CTkLabel(tab, text="준비 중", font=self._title_font)
        label.grid(row=0, column=0, pady=(12, 0), sticky="nsew")

    def _build_keyword_tab(self, tab: ctk.CTkFrame) -> None:
        tab.grid_rowconfigure(0, weight=1)
        tab.grid_rowconfigure(1, weight=0)
        tab.grid_columnconfigure(0, weight=0)
        tab.grid_columnconfigure(1, weight=1)

        left_panel = ctk.CTkFrame(tab, width=300)
        left_panel.grid(row=0, column=0, padx=(0, 10), pady=(12, 10), sticky="nsew")
        left_panel.grid_propagate(False)

        right_panel = ctk.CTkFrame(tab)
        right_panel.grid(row=0, column=1, padx=(0, 0), pady=(12, 10), sticky="nsew")

        bottom_bar = ctk.CTkFrame(tab)
        bottom_bar.grid(row=1, column=0, columnspan=2, sticky="sew")

        self._build_left_panel(left_panel)
        self._build_right_panel(right_panel)
        self._build_bottom_bar(bottom_bar)

    def _build_left_panel(self, panel: ctk.CTkFrame) -> None:
        panel.grid_columnconfigure(0, weight=1)

        title = ctk.CTkLabel(panel, text="검색 대상", font=self._title_font, anchor="w")
        title.grid(row=0, column=0, padx=10, pady=(10, 8), sticky="ew")

        self.path_frame = ctk.CTkScrollableFrame(panel, width=280, height=180)
        self.path_frame.grid(row=1, column=0, padx=10, pady=(0, 8), sticky="ew")

        self.add_folder_button = ctk.CTkButton(
            panel,
            text="폴더 추가",
            command=self._on_add_folder,
            font=self._font,
        )
        self.add_folder_button.grid(row=2, column=0, padx=10, pady=(0, 16), sticky="ew")

        keyword_title = ctk.CTkLabel(panel, text="키워드 관리", font=self._title_font, anchor="w")
        keyword_title.grid(row=3, column=0, padx=10, pady=(0, 8), sticky="ew")

        self.keyword_entry = ctk.CTkEntry(panel, font=self._font, placeholder_text="키워드 입력")
        self.keyword_entry.grid(row=4, column=0, padx=10, pady=(0, 8), sticky="ew")
        self.keyword_entry.bind("<Return>", self._on_add_keyword)

        list_container = ctk.CTkFrame(panel)
        list_container.grid(row=5, column=0, padx=10, pady=(0, 8), sticky="nsew")
        panel.grid_rowconfigure(5, weight=1)
        list_container.grid_rowconfigure(0, weight=1)
        list_container.grid_columnconfigure(0, weight=1)

        self.keyword_listbox = tk.Listbox(
            list_container,
            selectmode=tk.EXTENDED,
            exportselection=False,
            font=("맑은 고딕", 13),
            activestyle="none",
            height=10,
            bg="#121212",
            fg="#f5f7fa",
            selectbackground="#1f6aa5",
            selectforeground="#ffffff",
            highlightthickness=1,
            highlightbackground="#3a3f46",
            borderwidth=0,
            relief="flat",
        )
        self.keyword_listbox.grid(row=0, column=0, sticky="nsew")

        keyword_scroll = tk.Scrollbar(list_container, orient="vertical", command=self.keyword_listbox.yview)
        keyword_scroll.grid(row=0, column=1, sticky="ns")
        self.keyword_listbox.config(yscrollcommand=keyword_scroll.set)

        button_row = ctk.CTkFrame(panel, fg_color="transparent")
        button_row.grid(row=6, column=0, padx=10, pady=(0, 10), sticky="ew")
        for col in range(3):
            button_row.grid_columnconfigure(col, weight=1)

        add_button = ctk.CTkButton(button_row, text="추가", command=self._on_add_keyword, font=self._font)
        add_button.grid(row=0, column=0, padx=(0, 4), sticky="ew")

        remove_button = ctk.CTkButton(button_row, text="선택삭제", command=self._on_remove_keyword, font=self._font)
        remove_button.grid(row=0, column=1, padx=4, sticky="ew")

        save_button = ctk.CTkButton(button_row, text="전체삭제", command=self._on_save_keywords, font=self._font)
        save_button.grid(row=0, column=2, padx=(4, 0), sticky="ew")

        search_mode_title = ctk.CTkLabel(panel, text="검색 범위", font=self._font, anchor="w")
        search_mode_title.grid(row=7, column=0, padx=10, pady=(0, 6), sticky="ew")

        search_mode_row = ctk.CTkFrame(panel, fg_color="transparent")
        search_mode_row.grid(row=8, column=0, padx=10, pady=(0, 10), sticky="ew")
        search_mode_row.grid_columnconfigure(0, weight=1)
        search_mode_row.grid_columnconfigure(1, weight=1)

        filename_radio = ctk.CTkRadioButton(
            search_mode_row,
            text="파일명",
            variable=self._search_mode,
            value="filename",
            font=self._font,
        )
        filename_radio.grid(row=0, column=0, padx=(0, 8), sticky="w")

        content_radio = ctk.CTkRadioButton(
            search_mode_row,
            text="문서(내용)",
            variable=self._search_mode,
            value="content",
            font=self._font,
        )
        content_radio.grid(row=0, column=1, sticky="w")

    def _build_right_panel(self, panel: ctk.CTkFrame) -> None:
        panel.grid_rowconfigure(1, weight=1)
        panel.grid_rowconfigure(2, weight=0)
        panel.grid_columnconfigure(0, weight=1)

        title = ctk.CTkLabel(panel, text="검색 결과", font=self._title_font, anchor="w")
        title.grid(row=0, column=0, padx=10, pady=(10, 8), sticky="ew")

        tree_frame = ctk.CTkFrame(panel)
        tree_frame.grid(row=1, column=0, padx=10, pady=(0, 8), sticky="nsew")
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)
        self._configure_treeview_style()

        self.result_tree = ttk.Treeview(
            tree_frame,
            columns=("keyword", "filename", "extension", "filepath", "location", "context", "fullpath"),
            show="headings",
            displaycolumns=("keyword", "filename", "extension", "filepath", "location", "context"),
            style="Result.Treeview",
        )
        self.result_tree.grid(row=0, column=0, sticky="nsew")

        self.result_tree.heading("keyword", text="키워드", command=self._show_keyword_filter_popup)
        self.result_tree.heading("filename", text="파일명")
        self.result_tree.heading("extension", text="확장자", command=self._show_ext_filter_popup)
        self.result_tree.heading("filepath", text="파일경로")
        self.result_tree.heading("location", text="위치")
        self.result_tree.heading("context", text="해당 문장")
        self.result_tree.heading("fullpath", text="전체경로")

        self.result_tree.column("keyword", width=80, minwidth=80, anchor="w", stretch=False)
        self.result_tree.column("filename", width=180, minwidth=180, anchor="w", stretch=False)
        self.result_tree.column("extension", width=70, minwidth=70, anchor="w", stretch=False)
        self.result_tree.column("filepath", width=250, minwidth=250, anchor="w", stretch=False)
        self.result_tree.column("location", width=80, minwidth=80, anchor="w", stretch=False)
        self.result_tree.column("context", width=350, minwidth=220, anchor="w", stretch=True)
        self.result_tree.column("fullpath", width=0, minwidth=0, stretch=False)
        self.result_tree.tag_configure("fail", foreground="#FF6666")

        tree_v_scroll = ttk.Scrollbar(
            tree_frame,
            orient="vertical",
            command=self.result_tree.yview,
        )
        tree_v_scroll.grid(row=0, column=1, sticky="ns")
        tree_h_scroll = ttk.Scrollbar(
            tree_frame,
            orient="horizontal",
            command=self.result_tree.xview,
        )
        tree_h_scroll.grid(row=1, column=0, sticky="ew")
        self.result_tree.configure(yscrollcommand=tree_v_scroll.set, xscrollcommand=tree_h_scroll.set)

        self.result_menu = tk.Menu(self, tearoff=0)
        self.result_menu.add_command(label="파일 열기", command=self._open_selected_file)
        self.result_menu.add_command(label="경로 열기", command=self._open_selected_path)
        self.result_menu.add_command(label="경로 복사", command=self._copy_selected_path)
        self.result_tree.bind("<Double-1>", self._on_tree_double_click)
        self.result_tree.bind("<Button-3>", self._on_tree_right_click)

        self.summary_label = ctk.CTkLabel(panel, text="", font=self._summary_font, anchor="w")
        self.summary_label.grid(row=2, column=0, padx=10, pady=(0, 10), sticky="ew")

    def _build_bottom_bar(self, panel: ctk.CTkFrame) -> None:
        panel.grid_columnconfigure(1, weight=1)

        self.progress_label = ctk.CTkLabel(panel, text="진행: 0% (0/0파일)", font=self._font)
        self.progress_label.grid(row=0, column=0, padx=(10, 8), pady=10, sticky="w")

        self.progress_bar = ctk.CTkProgressBar(panel)
        self.progress_bar.grid(row=0, column=1, padx=(0, 12), pady=10, sticky="ew")
        self.progress_bar.set(0)

        self.search_button = ctk.CTkButton(
            panel,
            text="검색 시작",
            command=self._on_search_toggle,
            font=self._font,
            width=110,
        )
        self.search_button.grid(row=0, column=2, padx=(0, 8), pady=10)

        self.save_report_button = ctk.CTkButton(
            panel,
            text="리포트 저장(Excel)",
            command=self._on_save_report,
            font=self._font,
            width=150,
        )
        self.save_report_button.grid(row=0, column=3, padx=(0, 10), pady=10)

    def _load_system_drives(self) -> None:
        for letter in string.ascii_uppercase:
            drive = f"{letter}:\\"
            if os.path.exists(drive):
                self._add_path_option(drive, checked=False)

    def _load_keywords(self) -> None:
        try:
            keywords = scanner_engine.load_keywords()
        except Exception:
            return

        for keyword in keywords:
            self._insert_keyword_if_new(keyword)

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
            self.path_frame,
            text=normalized_path,
            variable=var,
            onvalue=True,
            offvalue=False,
            font=self._font,
        )
        checkbox.pack(anchor="w", fill="x", padx=4, pady=2)

        self._path_vars[canonical] = var
        self._path_values[canonical] = normalized_path

    def _on_add_keyword(self, _event: tk.Event | None = None) -> str | None:
        keyword = self.keyword_entry.get().strip()
        if self._insert_keyword_if_new(keyword):
            self.keyword_entry.delete(0, tk.END)
            scanner_engine.save_keywords(self._get_keywords())
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
        scanner_engine.save_keywords(self._get_keywords())

    def _on_save_keywords(self) -> None:
        if not messagebox.askyesno("확인", "키워드를 전체 삭제하시겠습니까?"):
            return
        self.keyword_listbox.delete(0, tk.END)
        scanner_engine.save_keywords([])
        self._set_summary_text("키워드 전체 삭제 완료")

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
        self._fail_count = 0
        self._skip_count = 0
        self._reset_filters()
        self._set_search_controls(is_searching=True)
        self._clear_tree_results()
        self._set_summary_text("검색 시작")
        self._update_progress(0, 0)

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

    def _search_worker(self, paths: list[str], keywords: list[str], search_mode: str) -> None:
        start_time = time.perf_counter()
        found_count = 0
        completed_count = 0
        total_files = 0
        files: list[str] = []

        try:
            if search_mode == "filename":
                files = scanner_engine.scan_files(paths, extensions=None)
            else:
                files = scanner_engine.scan_files(paths, extensions=scanner_engine.DEFAULT_EXTENSIONS)
            total_files = len(files)
            self._call_ui(self._set_summary_text, f"검색 중 | 대상 파일: {total_files:,}개")
            self._call_ui(self._update_progress, 0, total_files)
            if search_mode == "filename":
                for index, file_path in enumerate(files, start=1):
                    if self._stop_flag:
                        break

                    completed_count = index
                    try:
                        matches = scanner_engine.search_file_by_name(file_path, keywords)
                        error_message = ""
                    except Exception:
                        self._fail_count += 1
                        matches = []
                        error_message = "검색 처리 중 예외 발생"

                    if error_message:
                        self._call_ui(self._insert_tree_fail, file_path, error_message)

                    if matches:
                        for match in matches:
                            result = {
                                "keyword": match.get("keyword", ""),
                                "file_path": match.get("file", file_path),
                                "location": match.get("location", ""),
                                "context": match.get("context", ""),
                                "searched_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                            }
                            self._results.append(result)
                            found_count += 1
                            self._call_ui(self._insert_tree_result_from_result, result)

                    if index % 1000 == 0 or index == total_files:
                        self._call_ui(self._update_progress, index, total_files)
            else:
                with ThreadPoolExecutor(max_workers=4) as executor:
                    self._executor = executor
                    future_to_file = {
                        executor.submit(scanner_engine.search_file, file_path, keywords): file_path
                        for file_path in files
                    }
                    pending = set(future_to_file.keys())

                    while pending:
                        if self._stop_flag:
                            executor.shutdown(wait=False, cancel_futures=True)
                            break

                        try:
                            completed = next(as_completed(pending, timeout=0.1))
                            done_now = [completed]
                        except FuturesTimeoutError:
                            continue

                        for future in done_now:
                            if future in pending:
                                pending.remove(future)

                            file_path = future_to_file[future]
                            completed_count += 1

                            try:
                                matches = future.result()
                                error_message = ""
                            except Exception:
                                self._fail_count += 1
                                matches = []
                                error_message = "검색 처리 중 예외 발생"

                            meta = scanner_engine.consume_search_file_meta(file_path)
                            if meta.get("skipped", False):
                                self._skip_count += 1
                            if meta.get("failed", False):
                                self._fail_count += 1
                                if not error_message:
                                    error_message = "텍스트 추출 실패"

                            if error_message:
                                self._call_ui(self._insert_tree_fail, file_path, error_message)

                            if matches:
                                for match in matches:
                                    result = {
                                        "keyword": match.get("keyword", ""),
                                        "file_path": match.get("file", file_path),
                                        "location": match.get("location", ""),
                                        "context": match.get("context", ""),
                                        "searched_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                                    }
                                    self._results.append(result)
                                    found_count += 1
                                    self._call_ui(self._insert_tree_result_from_result, result)

                            self._call_ui(self._update_progress, completed_count, total_files)

                    if self._stop_flag:
                        executor.shutdown(wait=False, cancel_futures=True)

        except Exception:
            self._fail_count += 1
        finally:
            if search_mode == "content":
                for file_path in files:
                    scanner_engine.consume_search_file_meta(file_path)
            self._executor = None
            elapsed = time.perf_counter() - start_time
            if self._stop_flag:
                summary = f"검색 중지됨 | {completed_count:,}/{total_files:,}파일 처리"
            else:
                summary = (
                    f"검색 완료 | {total_files:,}개 스캔 | 스킵 {self._skip_count:,}개 | "
                    f"{found_count:,}건 발견 | 실패 {self._fail_count:,}개 | 소요시간: {elapsed:.1f}s"
                )
            self._call_ui(self._set_summary_text, summary)
            self._call_ui(self._finish_search)

    def _finish_search(self) -> None:
        self._is_searching = False
        self._stop_flag = False
        self._executor = None
        self._search_thread = None
        self._set_search_controls(is_searching=False)

    def _set_search_controls(self, is_searching: bool) -> None:
        self.search_button.configure(text="검색 중지" if is_searching else "검색 시작")
        self.save_report_button.configure(state="disabled" if is_searching else "normal")

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
        else:
            ratio = min(1.0, max(0.0, done / total))
            percent = int(ratio * 100)

        self.progress_bar.set(ratio)
        self.progress_label.configure(text=f"진행: {percent}% ({done:,}/{total:,}파일)")

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
            self.summary_label.configure(text=f"{self._base_summary_text} | {filter_text}")
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
        self,
        keyword: str,
        file_path: str,
        location: str,
        context: str,
        tags: tuple[str, ...] = (),
    ) -> dict[str, object]:
        normalized_file_path = str(file_path)
        filename = os.path.basename(normalized_file_path) or normalized_file_path
        extension = os.path.splitext(normalized_file_path)[1].lower()
        folder_path = os.path.dirname(normalized_file_path)
        table_context = self._truncate_text(str(context).replace("\n", " ").strip(), 80)
        return {
            "keyword": str(keyword),
            "filename": filename,
            "extension": extension,
            "filepath_display": f"📁{folder_path}",
            "location": str(location),
            "context_display": table_context,
            "fullpath": normalized_file_path,
            "tags": tuple(tags),
        }

    def _insert_tree_row(self, row: dict[str, object]) -> None:
        self.result_tree.insert(
            "",
            tk.END,
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
        self._refresh_summary_text()

    def _apply_filters(self) -> None:
        self._clear_tree_widget()
        for row in self._all_results:
            if self._row_passes_filters(row):
                self._insert_tree_row(row)
        self._refresh_summary_text()

    def _get_available_filter_values(self, target: str) -> list[str]:
        values = {str(row.get(target, "")) for row in self._all_results}
        return sorted(values, key=lambda value: (value == "", value.lower()))

    def _format_filter_value(self, target: str, value: str) -> str:
        if target == "extension" and value == "":
            return "(확장자 없음)"
        return value

    def _get_filter_popup_position(self, column_name: str) -> tuple[int, int]:
        display_columns = list(self.result_tree["displaycolumns"])
        offset_x = 0
        for display_column in display_columns:
            if display_column == column_name:
                break
            offset_x += int(self.result_tree.column(display_column, "width"))

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
        self._show_filter_popup(target="keyword", column_name="keyword", title="키워드")

    def _show_ext_filter_popup(self) -> None:
        self._show_filter_popup(target="extension", column_name="extension", title="확장자")

    def _show_filter_popup(self, target: str, column_name: str, title: str) -> None:
        self._close_filter_popup()

        values = self._get_available_filter_values(target)
        all_values = set(values)
        selected_values = (
            set(values)
            if (target == "keyword" and self._keyword_select_all) or (target == "extension" and self._ext_select_all)
            else set(self._keyword_filter if target == "keyword" else self._ext_filter)
        )
        selected_values &= all_values

        popup = tk.Toplevel(self)
        self._filter_popup = popup
        popup.withdraw()
        popup.transient(self)
        popup.overrideredirect(True)
        popup.configure(bg="#2b2b2b")

        container = tk.Frame(
            popup,
            bg="#2b2b2b",
            bd=1,
            relief="solid",
            highlightthickness=1,
            highlightbackground="#555555",
        )
        container.pack(fill="both", expand=True)

        title_label = tk.Label(
            container,
            text=f"{title} 필터",
            bg="#2b2b2b",
            fg="#ffffff",
            font=("맑은 고딕", 12, "bold"),
            anchor="w",
            padx=8,
            pady=6,
        )
        title_label.pack(fill="x")

        all_var = tk.BooleanVar(value=selected_values == all_values)
        item_vars: dict[str, tk.BooleanVar] = {}

        def apply_selection() -> None:
            chosen = {value for value, var in item_vars.items() if var.get()}
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
            all_var.set(all(item.get() for item in item_vars.values()) if item_vars else False)
            apply_selection()

        all_checkbox = tk.Checkbutton(
            container,
            text="전체",
            variable=all_var,
            command=on_toggle_all,
            bg="#2b2b2b",
            fg="#ffffff",
            font=("맑은 고딕", 12),
            selectcolor="#1f1f1f",
            activebackground="#2b2b2b",
            activeforeground="#ffffff",
            anchor="w",
            padx=8,
            pady=2,
            relief="flat",
            highlightthickness=0,
        )
        all_checkbox.pack(fill="x")

        if values:
            for value in values:
                item_var = tk.BooleanVar(value=value in selected_values)
                item_vars[value] = item_var
                checkbox = tk.Checkbutton(
                    container,
                    text=self._format_filter_value(target, value),
                    variable=item_var,
                    command=on_toggle_item,
                    bg="#2b2b2b",
                    fg="#ffffff",
                    font=("맑은 고딕", 12),
                    selectcolor="#1f1f1f",
                    activebackground="#2b2b2b",
                    activeforeground="#ffffff",
                    anchor="w",
                    padx=20,
                    pady=2,
                    relief="flat",
                    highlightthickness=0,
                )
                checkbox.pack(fill="x")
        else:
            empty_label = tk.Label(
                container,
                text="필터 대상이 없습니다",
                bg="#2b2b2b",
                fg="#b5b5b5",
                font=("맑은 고딕", 11),
                anchor="w",
                padx=8,
                pady=6,
            )
            empty_label.pack(fill="x")
            all_checkbox.configure(state="disabled")

        popup.bind("<FocusOut>", lambda _event: self._close_filter_popup())
        popup.bind("<Escape>", lambda _event: self._close_filter_popup())

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

    def _insert_tree_result(self, keyword: str, file_path: str, location: str, context: str) -> None:
        row = self._build_tree_row(keyword, file_path, location, context)
        self._append_tree_row(row)

    def _insert_tree_fail(self, file_path: str, error_message: str) -> None:
        row = self._build_tree_row("실패", file_path, "-", error_message, tags=("fail",))
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
        file_path = self._get_item_file_path(item_id) if item_id else self._get_selected_tree_file_path()
        if not file_path or not os.path.exists(file_path):
            messagebox.showwarning("경고", "파일을 찾을 수 없습니다")
            return
        try:
            os.startfile(file_path)
        except Exception:
            messagebox.showwarning("경고", "파일을 찾을 수 없습니다")

    def _open_path_by_item_id(self, item_id: str) -> None:
        file_path = self._get_item_file_path(item_id) if item_id else self._get_selected_tree_file_path()
        if not file_path or not os.path.exists(file_path):
            messagebox.showwarning("경고", "파일을 찾을 수 없습니다")
            return
        escaped = file_path.replace('"', '""')
        subprocess.Popen(f'explorer /select,"{escaped}"')

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
            background="#101215",
            foreground="#f3f5f7",
            fieldbackground="#101215",
            rowheight=26,
            font=("맑은 고딕", 12),
            bordercolor="#2f343b",
            lightcolor="#2f343b",
            darkcolor="#2f343b",
            borderwidth=1,
            relief="solid",
        )
        style.configure(
            "Result.Treeview.Heading",
            background="#1a1d22",
            foreground="#f8fafc",
            font=("맑은 고딕", 13, "bold"),
            bordercolor="#2f343b",
            lightcolor="#2f343b",
            darkcolor="#2f343b",
            borderwidth=1,
            relief="raised",
        )
        style.map(
            "Result.Treeview",
            background=[("selected", "#1e5f8f")],
            foreground=[("selected", "#ffffff")],
        )
        style.layout("Result.Treeview", [("Treeview.treearea", {"sticky": "nswe"})])

    def _call_ui(self, callback: Callable[..., None], *args: object) -> None:
        try:
            self.after(0, lambda: callback(*args))
        except RuntimeError:
            return

    def _on_close(self) -> None:
        self._stop_flag = True
        self._close_filter_popup()
        self.destroy()


def main() -> None:
    """애플리케이션을 실행한다."""
    ctk.set_appearance_mode("Dark")
    ctk.set_default_color_theme("blue")
    app = FileScannerApp()
    app.mainloop()


if __name__ == "__main__":
    main()
