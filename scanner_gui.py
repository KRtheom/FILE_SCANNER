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
        self.geometry("1100x700")
        self.resizable(True, True)

        self._stop_flag = False
        self._is_searching = False
        self._results: list[dict] = []
        self._search_thread: threading.Thread | None = None
        self._executor: ThreadPoolExecutor | None = None
        self._fail_count = 0
        self._skip_count = 0

        self._path_vars: dict[str, ctk.BooleanVar] = {}
        self._path_values: dict[str, str] = {}

        self._font = ctk.CTkFont(size=13)
        self._title_font = ctk.CTkFont(size=16, weight="bold")
        self._tab_font = ctk.CTkFont(size=17)

        self._build_ui()
        self._load_system_drives()
        self._load_keywords()
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    def _build_ui(self) -> None:
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        self.tabview = ctk.CTkTabview(self)
        self.tabview.grid(row=0, column=0, padx=12, pady=(2, 12), sticky="nsew")

        self.keyword_tab = self.tabview.add("키워드 검색")
        self.pi_tab = self.tabview.add("개인정보 검색")
        self.illegal_sw_tab = self.tabview.add("불법S/W파일")

        # 탭 헤더 가독성 개선: 17pt 폰트 + 균등 배분 고정 폭
        self.tabview.configure(anchor="center")
        self.tabview._segmented_button.configure(
            font=self._tab_font,
            dynamic_resizing=False,
            width=660,
        )
        for tab_button in self.tabview._segmented_button._buttons_dict.values():
            tab_button.configure(anchor="center", width=220)

        self._build_keyword_tab(self.keyword_tab)
        self._build_placeholder_tab(self.pi_tab)
        self._build_placeholder_tab(self.illegal_sw_tab)

    def _build_placeholder_tab(self, tab: ctk.CTkFrame) -> None:
        tab.grid_rowconfigure(0, weight=1)
        tab.grid_columnconfigure(0, weight=1)
        label = ctk.CTkLabel(tab, text="준비 중", font=self._title_font)
        label.grid(row=0, column=0, sticky="nsew")

    def _build_keyword_tab(self, tab: ctk.CTkFrame) -> None:
        tab.grid_rowconfigure(0, weight=1)
        tab.grid_rowconfigure(1, weight=0)
        tab.grid_columnconfigure(0, weight=0)
        tab.grid_columnconfigure(1, weight=1)

        left_panel = ctk.CTkFrame(tab, width=300)
        left_panel.grid(row=0, column=0, padx=(0, 10), pady=(0, 10), sticky="nsew")
        left_panel.grid_propagate(False)

        right_panel = ctk.CTkFrame(tab)
        right_panel.grid(row=0, column=1, padx=(0, 0), pady=(0, 10), sticky="nsew")

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
            columns=("keyword", "filename", "filepath", "location", "context", "fullpath"),
            show="headings",
            displaycolumns=("keyword", "filename", "filepath", "location", "context"),
            style="Result.Treeview",
        )
        self.result_tree.grid(row=0, column=0, sticky="nsew")

        self.result_tree.heading("keyword", text="키워드")
        self.result_tree.heading("filename", text="파일명")
        self.result_tree.heading("filepath", text="파일경로")
        self.result_tree.heading("location", text="위치")
        self.result_tree.heading("context", text="해당 문장")
        self.result_tree.heading("fullpath", text="전체경로")

        self.result_tree.column("keyword", width=80, minwidth=80, anchor="w", stretch=False)
        self.result_tree.column("filename", width=200, minwidth=180, anchor="w", stretch=False)
        self.result_tree.column("filepath", width=350, minwidth=300, anchor="w", stretch=True)
        self.result_tree.column("location", width=100, minwidth=100, anchor="w", stretch=False)
        self.result_tree.column("context", width=350, minwidth=250, anchor="w", stretch=True)
        self.result_tree.column("fullpath", width=0, minwidth=0, stretch=False)
        self.result_tree.tag_configure("fail", foreground="#FF6666")

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

        self.summary_label = ctk.CTkLabel(panel, text="", font=ctk.CTkFont(size=12), anchor="w")
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
            # 시작 시 로그 영역은 비어 있어야 하므로 로드 실패는 조용히 무시한다.
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

        self._results = []
        self._stop_flag = False
        self._is_searching = True
        self._fail_count = 0
        self._skip_count = 0
        self._set_search_controls(is_searching=True)
        self._clear_tree_results()
        self._set_summary_text("검색 시작")
        self._update_progress(0, 0)

        self._search_thread = threading.Thread(
            target=self._search_worker,
            args=(paths, keywords),
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

    def _search_worker(self, paths: list[str], keywords: list[str]) -> None:
        start_time = time.perf_counter()
        found_count = 0
        completed_count = 0
        total_files = 0
        files: list[str] = []

        try:
            files = scanner_engine.scan_files(paths, TARGET_EXTENSIONS)
            total_files = len(files)
            self._call_ui(self._set_summary_text, f"검색 중 | 대상 파일: {total_files:,}개")
            self._call_ui(self._update_progress, 0, total_files)

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
        self.summary_label.configure(text=text)

    def _clear_tree_results(self) -> None:
        for item_id in self.result_tree.get_children():
            self.result_tree.delete(item_id)

    def _insert_tree_result_from_result(self, result: dict) -> None:
        keyword = str(result.get("keyword", ""))
        file_path = str(result.get("file_path", ""))
        location = str(result.get("location", ""))
        context = str(result.get("context", "")).replace("\n", " ").strip()
        self._insert_tree_result(keyword, file_path, location, context)

    def _insert_tree_result(self, keyword: str, file_path: str, location: str, context: str) -> None:
        filename = os.path.basename(file_path) or file_path
        folder_path = os.path.dirname(file_path)
        table_context = self._truncate_text(context, 80)
        self.result_tree.insert(
            "",
            tk.END,
            values=(keyword, filename, f"📁 {folder_path}", location, table_context, file_path),
        )

    def _insert_tree_fail(self, file_path: str, error_message: str) -> None:
        filename = os.path.basename(file_path) or file_path
        folder_path = os.path.dirname(file_path)
        context = self._truncate_text(error_message, 80)
        self.result_tree.insert(
            "",
            tk.END,
            values=("실패", filename, f"📁 {folder_path}", "-", context, file_path),
            tags=("fail",),
        )

    def _on_tree_double_click(self, event: tk.Event) -> None:
        col_id = self.result_tree.identify_column(event.x)
        row_id = self.result_tree.identify_row(event.y)
        if not row_id:
            return
        self.result_tree.selection_set(row_id)
        self.result_tree.focus(row_id)

        if col_id == "#2":
            self._open_file_by_item_id(row_id)
        elif col_id == "#3":
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
        if len(values) < 6:
            return ""
        return str(values[5])

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
        return f"{text[:limit]}..."

    def _configure_treeview_style(self) -> None:
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure("Result.Treeview", rowheight=25, borderwidth=1, relief="solid")
        style.configure("Result.Treeview.Heading", borderwidth=1, relief="raised")
        style.layout("Result.Treeview", [("Treeview.treearea", {"sticky": "nswe"})])

    def _call_ui(self, callback: Callable[..., None], *args: object) -> None:
        try:
            self.after(0, lambda: callback(*args))
        except RuntimeError:
            # 창 종료 중에는 after 등록이 실패할 수 있으므로 조용히 무시한다.
            return

    def _on_close(self) -> None:
        self._stop_flag = True
        self.destroy()


def main() -> None:
    """애플리케이션을 실행한다."""
    ctk.set_appearance_mode("System")
    ctk.set_default_color_theme("blue")
    app = FileScannerApp()
    app.mainloop()


if __name__ == "__main__":
    main()
