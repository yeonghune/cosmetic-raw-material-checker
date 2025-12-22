from __future__ import annotations

import os
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox

from tkinter import ttk

from app.excel import (
    TablePayload,
    ValidationPayload,
    download_template_file,
    table_cross_validation,
    upload_template_file,
    data_merge,
    RESULT_HEADER,
    Header,
    export_validation_result,
)


class TemplateValidatorApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("엑셀 템플릿 검증 도구")
        self.geometry("800x800")
        self.minsize(800, 600)
        self.maxsize(1920, 1000)
        self.configure(padx=16, pady=16)

        self.selected_file = tk.StringVar(value="템플릿이 로드되지 않았습니다.")
        self.status_var = tk.StringVar(value="준비 완료")

        # 현재 로드된 테이블 데이터 (없으면 None)
        self.table1_payload: TablePayload | None = None
        self.table2_payload: TablePayload | None = None

        # Treeview를 감싸는 프레임 및 Treeview 참조
        self.table1_frame: ttk.Frame | None = None
        self.table2_frame: ttk.Frame | None = None
        self.table1_tree: ttk.Treeview | None = None
        self.table2_tree: ttk.Treeview | None = None
        self.result_tree: ttk.Treeview | None = None
        self.result_frame: ttk.LabelFrame | None = None
        self.summary_label: ttk.Label | None = None
        # 각 테이블에서만 존재하는 (RM, INCI) 키 집합
        self.unique_keys_table1: set[tuple[str, str]] = set()
        self.unique_keys_table2: set[tuple[str, str]] = set()
        # 마지막 검증 결과(엑셀 다운로드용)
        self.last_result_rows: list[ValidationPayload] | None = None
        # 셀 편집용 엔트리 상태
        self._cell_editor: ttk.Entry | None = None
        self._cell_editor_info: tuple[ttk.Treeview, str, int] | None = None

        self._build_header()
        self._build_table_preview()
        self._build_result_panel()

        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)  # 미리보기 영역
        self.rowconfigure(2, weight=1)  # 결과 영역

    # --------------------------- UI builders --------------------------- #
    def _build_header(self) -> None:
        header = ttk.Frame(self)
        header.grid(row=0, column=0, sticky="ew")
        header.columnconfigure(3, weight=1)

        download_btn = ttk.Button(header, text="템플릿 다운로드", command=self.download_template)
        upload_btn = ttk.Button(header, text="템플릿 불러오기", command=self.upload_template)
        download_result_btn = ttk.Button(header, text="검증 결과 다운로드", command=self.download_result)
        file_label = ttk.Label(header, textvariable=self.selected_file, foreground="#444444")

        download_btn.grid(row=0, column=0, padx=(0, 8))
        upload_btn.grid(row=0, column=1, padx=(0, 8))
        download_result_btn.grid(row=0, column=2, padx=(0, 8))
        file_label.grid(row=0, column=3, sticky="ew")

    def _build_table_preview(self) -> None:
        preview = ttk.LabelFrame(self, text="템플릿 미리보기")
        preview.grid(row=1, column=0, sticky="nsew", pady=12)
        preview.columnconfigure(0, weight=1)
        preview.columnconfigure(1, weight=1)
        preview.rowconfigure(0, weight=1)

        table1_frame = ttk.Frame(preview)
        table2_frame = ttk.Frame(preview)
        table1_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 6))
        table2_frame.grid(row=0, column=1, sticky="nsew", padx=(6, 0))
        table1_frame.columnconfigure(0, weight=1)
        table1_frame.rowconfigure(1, weight=1)
        table2_frame.columnconfigure(0, weight=1)
        table2_frame.rowconfigure(1, weight=1)

        ttk.Label(table1_frame, text="테이블 1", font=("", 10, "bold")).grid(
            row=0,
            column=0,
            sticky="w",
            pady=(0, 4),
        )
        ttk.Label(table2_frame, text="테이블 2", font=("", 10, "bold")).grid(
            row=0,
            column=0,
            sticky="w",
            pady=(0, 4),
        )

        self.table1_frame = table1_frame
        self.table2_frame = table2_frame

        # 초기에는 placeholder 데이터로 Treeview 구성
        self.table1_tree = self._create_treeview(None, table1_frame)
        self.table2_tree = self._create_treeview(None, table2_frame)

        self.table1_tree.master.grid(row=1, column=0, sticky="nsew")
        self.table2_tree.master.grid(row=1, column=0, sticky="nsew")
        self._bind_cell_editing(self.table1_tree)
        self._bind_cell_editing(self.table2_tree)

    def _build_result_panel(self) -> None:
        container = ttk.Frame(self)
        container.grid(row=2, column=0, sticky="nsew")
        container.columnconfigure(0, weight=1)
        container.rowconfigure(0, weight=1)

        result_frame = ttk.LabelFrame(container, text="검증 결과 (테이블 비교)")
        result_frame.grid(row=0, column=0, sticky="nsew")
        result_frame.columnconfigure(0, weight=1)
        result_frame.rowconfigure(0, weight=1)
        self.result_frame = result_frame

        # 결과용 Treeview: ValidationPayload 리스트를 표시하는 전용 뷰
        self.result_tree = self._create_result_view([], result_frame)
        self.result_tree.master.grid(row=0, column=0, sticky="nsew")

        summary_frame = ttk.Frame(container, padding=(0, 12, 0, 0))
        summary_frame.grid(row=3, column=0, sticky="ew")
        summary_frame.columnconfigure(0, weight=1)

        self.summary_label = ttk.Label(summary_frame, text="불일치 0건 / 총 0건")
        self.summary_label.grid(row=0, column=0, sticky="w")

    # --------------------------- Treeview helpers --------------------------- #
    def _to_str(self, value: object) -> str:
        if value is None:
            return ""
        return str(value)

    def _create_treeview(
        self,
        payload: TablePayload | None,
        parent: tk.Misc,
        unique_keys: set[tuple[str, str]] | None = None,
    ) -> ttk.Treeview:
        """payload가 None이면 placeholder, 아니면 payload 기준으로 Treeview 생성."""
        container = ttk.Frame(parent)
        container.columnconfigure(0, weight=1)
        container.rowconfigure(0, weight=1)

        if not payload or payload.is_empty or not payload.header:
            columns = ("RM", "% RM/FP", "INCI", "% INCI/RM")
            headings = ("RM", "% RM/FP", "INCI", "% INCI/RM")
        else:
            columns = tuple(str(h) for h in payload.header)
            headings = columns

        tree = ttk.Treeview(container, columns=columns, show="headings", height=10)

        for col, heading_text in zip(columns, headings):
            tree.heading(col, text=heading_text)
            tree.column(col, anchor="center", stretch=True, width=120)

        # 고유 row 하이라이트용 태그
        tree.tag_configure("unique", background="#ffcccc")

        scroll_y = ttk.Scrollbar(container, orient="vertical", command=tree.yview)
        scroll_x = ttk.Scrollbar(container, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)

        tree.grid(row=0, column=0, sticky="nsew")
        scroll_y.grid(row=0, column=1, sticky="ns")
        scroll_x.grid(row=1, column=0, sticky="ew")

        # payload가 있으면 데이터 채우기
        if payload and not payload.is_empty and payload.header:
            headers = list(payload.header)
            tree._payload_headers = headers
            for row_index, row_dict in enumerate(payload.data):
                values = [row_dict.get(h, "") for h in headers]
                if unique_keys is not None:
                    rm = self._to_str(row_dict.get(Header.RM))
                    inci = self._to_str(row_dict.get(Header.INCI))
                    key = (rm, inci)
                    if key in unique_keys:
                        tree.insert(
                            "",
                            "end",
                            iid=str(row_index),
                            values=values,
                            tags=("unique",),
                        )
                        continue
                tree.insert("", "end", iid=str(row_index), values=values)

        return tree

    def _bind_cell_editing(self, tree: ttk.Treeview) -> None:
        tree.bind("<Double-1>", lambda event, t=tree: self._start_cell_edit(event, t))

    def _get_payload_for_tree(self, tree: ttk.Treeview) -> TablePayload | None:
        if tree is self.table1_tree:
            return self.table1_payload
        if tree is self.table2_tree:
            return self.table2_payload
        return None

    def _start_cell_edit(self, event: tk.Event, tree: ttk.Treeview) -> None:
        payload = self._get_payload_for_tree(tree)
        if not payload or payload.is_empty:
            return

        if tree.identify_region(event.x, event.y) != "cell":
            return

        row_id = tree.identify_row(event.y)
        col_id = tree.identify_column(event.x)
        if not row_id or not col_id or col_id == "#0":
            return

        try:
            col_index = int(col_id[1:]) - 1
        except ValueError:
            return

        columns = tree["columns"]
        if col_index < 0 or col_index >= len(columns):
            return

        bbox = tree.bbox(row_id, col_id)
        if not bbox:
            return

        self._destroy_cell_editor()

        value = tree.set(row_id, columns[col_index])
        entry = ttk.Entry(tree)
        entry.insert(0, value)
        entry.select_range(0, tk.END)
        entry.focus()

        x, y, width, height = bbox
        entry.place(x=x, y=y, width=width, height=height)

        self._cell_editor = entry
        self._cell_editor_info = (tree, row_id, col_index)

        entry.bind("<Return>", self._commit_cell_edit)
        entry.bind("<Escape>", self._cancel_cell_edit)
        entry.bind("<FocusOut>", self._commit_cell_edit)

    def _commit_cell_edit(self, event: tk.Event | None = None) -> None:
        if self._cell_editor is None or self._cell_editor_info is None:
            return

        tree, row_id, col_index = self._cell_editor_info
        columns = tree["columns"]
        if col_index < 0 or col_index >= len(columns):
            self._destroy_cell_editor()
            return

        col_name = columns[col_index]
        new_value = self._cell_editor.get()
        old_value = tree.set(row_id, col_name)
        if new_value == old_value:
            self._destroy_cell_editor()
            return

        tree.set(row_id, col_name, new_value)

        payload = self._get_payload_for_tree(tree)
        if payload and not payload.is_empty:
            try:
                row_index = int(row_id)
            except ValueError:
                row_index = tree.index(row_id)

            if 0 <= row_index < len(payload.data):
                headers = getattr(tree, "_payload_headers", None)
                if headers and col_index < len(headers):
                    header_key = headers[col_index]
                else:
                    header_key = col_name
                payload.data[row_index][header_key] = new_value

        self._destroy_cell_editor()
        self._run_validation()

    def _cancel_cell_edit(self, event: tk.Event | None = None) -> None:
        self._destroy_cell_editor()

    def _destroy_cell_editor(self) -> None:
        if self._cell_editor is not None:
            self._cell_editor.destroy()
        self._cell_editor = None
        self._cell_editor_info = None

    def _capture_tree_scroll(self, tree: ttk.Treeview | None) -> tuple[float, float]:
        if tree is None:
            return (0.0, 0.0)
        xview = tree.xview()
        yview = tree.yview()
        x_pos = xview[0] if xview else 0.0
        y_pos = yview[0] if yview else 0.0
        return (x_pos, y_pos)

    def _restore_tree_scroll(self, tree: ttk.Treeview | None, x_pos: float, y_pos: float) -> None:
        if tree is None:
            return
        tree.xview_moveto(x_pos)
        tree.yview_moveto(y_pos)

    def _capture_tree_selection(self, tree: ttk.Treeview | None) -> list[str]:
        if tree is None:
            return []
        return list(tree.selection())

    def _restore_tree_selection(self, tree: ttk.Treeview | None, selection: list[str]) -> None:
        if tree is None or not selection:
            return
        existing = [item for item in selection if tree.exists(item)]
        if not existing:
            return
        tree.selection_set(existing)
        tree.focus(existing[0])
        tree.see(existing[0])

    def _rebuild_table_views(self) -> None:
        """로드된 테이블 1/2 Treeview를 payload 기준으로 다시 생성."""
        table1_scroll = self._capture_tree_scroll(self.table1_tree)
        table2_scroll = self._capture_tree_scroll(self.table2_tree)
        table1_selection = self._capture_tree_selection(self.table1_tree)
        table2_selection = self._capture_tree_selection(self.table2_tree)
        if self.table1_frame is not None:
            if self.table1_tree is not None:
                self.table1_tree.master.destroy()
            self.table1_tree = self._create_treeview(
                self.table1_payload,
                self.table1_frame,
                self.unique_keys_table1,
            )
            self.table1_tree.master.grid(row=1, column=0, sticky="nsew")
            self._bind_cell_editing(self.table1_tree)
            self._restore_tree_scroll(self.table1_tree, *table1_scroll)
            self._restore_tree_selection(self.table1_tree, table1_selection)

        if self.table2_frame is not None:
            if self.table2_tree is not None:
                self.table2_tree.master.destroy()
            self.table2_tree = self._create_treeview(
                self.table2_payload,
                self.table2_frame,
                self.unique_keys_table2,
            )
            self.table2_tree.master.grid(row=1, column=0, sticky="nsew")
            self._bind_cell_editing(self.table2_tree)
            self._restore_tree_scroll(self.table2_tree, *table2_scroll)
            self._restore_tree_selection(self.table2_tree, table2_selection)

    def _compute_unique_keys(self) -> None:
        """각 테이블에만 존재하는 (RM, INCI) 키 집합을 계산한다."""
        def _keys_from_payload(payload: TablePayload | None) -> set[tuple[str, str]]:
            keys: set[tuple[str, str]] = set()
            if not payload or payload.is_empty:
                return keys
            for row in payload.data:
                rm = self._to_str(row.get(Header.RM))
                inci = self._to_str(row.get(Header.INCI))
                keys.add((rm, inci))
            return keys

        keys1 = _keys_from_payload(self.table1_payload)
        keys2 = _keys_from_payload(self.table2_payload)

        self.unique_keys_table1 = keys1 - keys2
        self.unique_keys_table2 = keys2 - keys1

    def _get_result_row_tag(self, payload: ValidationPayload) -> str:
        """결과 테이블에서 사용할 행 색상 태그를 반환한다."""
        # INCI/RM이 다르면 빨간색
        if payload.inci_rm_table1 != payload.inci_rm_table2:
            return "row_inci_rm_diff"
        # INCI/RM은 같고 RM/FP가 다르면 노란색
        if payload.rm_fp_table1 != payload.rm_fp_table2:
            return "row_rm_fp_diff"
        return ""

    def _create_result_view(
        self,
        rows: list[ValidationPayload] | None,
        parent: tk.Misc,
    ) -> ttk.Treeview:
        """ValidationPayload 리스트를 RESULT_HEADER 구조로 출력하는 Treeview 생성."""
        container = ttk.Frame(parent)
        container.columnconfigure(0, weight=1)
        container.rowconfigure(0, weight=1)

        tree = ttk.Treeview(container, columns=RESULT_HEADER, show="headings", height=10)

        for col in RESULT_HEADER:
            tree.heading(col, text=col)
            tree.column(col, anchor="center", stretch=True, width=120)

        scroll_y = ttk.Scrollbar(container, orient="vertical", command=tree.yview)
        scroll_x = ttk.Scrollbar(container, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)

        tree.grid(row=0, column=0, sticky="nsew")
        scroll_y.grid(row=0, column=1, sticky="ns")
        scroll_x.grid(row=1, column=0, sticky="ew")

        # 행 색상 태그 정의
        tree.tag_configure("row_inci_rm_diff", background="#ffcccc")  # INCI/RM 불일치 (빨간색)
        tree.tag_configure("row_rm_fp_diff", background="#fff2cc")    # RM/FP 불일치 (노란색)

        # ValidationPayload 리스트 채우기
        if rows:
            for payload in rows:
                # RM은 두 테이블이 동일하다고 가정하고 하나만 표시
                values = [
                    payload.rm,
                    payload.inci,
                    payload.rm_fp_table1,
                    payload.rm_fp_table2,
                    payload.inci_rm_table1,
                    payload.inci_rm_table2,
                ]

                tag = self._get_result_row_tag(payload)
                item_id = f"{payload.rm}||{payload.inci}"

                if tag:
                    tree.insert("", "end", iid=item_id, values=values, tags=(tag,))
                else:
                    tree.insert("", "end", iid=item_id, values=values)

        return tree

    # --------------------------- event handlers --------------------------- #
    def download_template(self) -> None:
        file_path = filedialog.asksaveasfilename(
            title="템플릿 저장 위치/파일명 선택",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialfile="템플릿_샘플.xlsx",
        )
        if not file_path:
            return

        output_path = Path(file_path)
        try:
            saved_path = download_template_file(output_path)
        except Exception as exc:  # pragma: no cover - GUI error path
            messagebox.showerror("템플릿 다운로드 실패", f"파일을 저장하는 중 오류가 발생했습니다.\n")
            return

        self.status_var.set("템플릿 다운로드 완료")
        msg = f"템플릿이 저장되었습니다.\n{saved_path}\n\n폴더를 여시겠습니까?"
        if messagebox.askyesno("템플릿 다운로드", msg):
            self._open_folder(saved_path.parent)

    def upload_template(self) -> None:
        file_path = filedialog.askopenfilename(
            title="템플릿 파일 선택",
            filetypes=[("Excel files", "*.xlsx *.xlsm"), ("All files", "*.*")],
        )
        if not file_path:
            return
        
        self.selected_file.set(Path(file_path).name)

        try:
            table1_payload, table2_payload = upload_template_file(input_path=file_path)
        except Exception as exc:  # pragma: no cover - GUI error path
            messagebox.showerror("템플릿 불러오기 실패", f"파일을 여는 중 오류가 발생했습니다.\n{exc}")
            return

        ok, msg = table_cross_validation(table1_payload, table2_payload)
        if not ok:
            messagebox.showerror("기본 검증 실패", msg)
            self.table1_payload = None
            self.table2_payload = None
            # 고유 키 및 뷰 초기화
            self.unique_keys_table1.clear()
            self.unique_keys_table2.clear()
            self.last_result_rows = None
            self._rebuild_table_views()
            if self.summary_label is not None:
                self.summary_label.configure(text="불일치 0건 / 총 0건")
            self.status_var.set("템플릿 불러오기 실패")
            return

        self.table1_payload = table1_payload
        self.table2_payload = table2_payload
        # 새 템플릿을 불러오면 고유 키와 요약 정보를 초기화
        self.unique_keys_table1.clear()
        self.unique_keys_table2.clear()
        self.last_result_rows = None
        self._rebuild_table_views()
        if self.summary_label is not None:
            self.summary_label.configure(text="불일치 0건 / 총 0건")
        self.status_var.set("템플릿 불러오기 완료")

        self._run_validation()

    def _run_validation(self) -> None:
        if not self.table1_payload or not self.table2_payload:
            messagebox.showwarning("검증", "먼저 템플릿을 불러와 주세요.")
            return

        try:
            rows = data_merge(self.table1_payload, self.table2_payload)
        except Exception as exc:  # pragma: no cover - GUI error path
            self.status_var.set("검증 실패")
            messagebox.showerror("검증 실패", f"검증 중 오류가 발생했습니다.\n{exc}")
            return

        # 마지막 검증 결과 저장 (엑셀 다운로드용)
        self.last_result_rows = list(rows)

        # 검증 시, 각 테이블에만 존재하는 (RM, INCI) row 하이라이트 갱신
        self._compute_unique_keys()
        self._rebuild_table_views()

        result_scroll = self._capture_tree_scroll(self.result_tree)
        result_selection = self._capture_tree_selection(self.result_tree)
        if self.result_frame is not None:
            if self.result_tree is not None:
                self.result_tree.master.destroy()
            self.result_tree = self._create_result_view(rows, self.result_frame)
            self.result_tree.master.grid(row=0, column=0, sticky="nsew")
            self._restore_tree_scroll(self.result_tree, *result_scroll)
            self._restore_tree_selection(self.result_tree, result_selection)

        total = len(rows)
        if self.summary_label is not None:
            red = 0
            yellow = 0
            for payload in rows:
                tag = self._get_result_row_tag(payload)
                if tag == "row_inci_rm_diff":
                    red += 1
                elif tag == "row_rm_fp_diff":
                    yellow += 1
            self.summary_label.configure(
                text=f"INCI/RM 불일치 {red}건, RM/FP 불일치 {yellow}건 / 총 {total}건",
            )

        self.status_var.set("검증 완료")

    def download_result(self) -> None:
        """검증 결과를 엑셀 파일로 다운로드."""
        if (
            not self.table1_payload
            or not self.table2_payload
            or not self.last_result_rows
        ):
            messagebox.showwarning("검증 결과 다운로드", "먼저 검증을 실행해 주세요.")
            return

        file_path = filedialog.asksaveasfilename(
            title="검증 결과 저장",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            initialfile="검증_결과.xlsx",
        )
        if not file_path:
            return

        try:
            saved_path = export_validation_result(
                output_path=file_path,
                table1=self.table1_payload,
                table2=self.table2_payload,
                rows=self.last_result_rows,
                unique_keys_table1=self.unique_keys_table1,
                unique_keys_table2=self.unique_keys_table2,
            )
        except Exception as exc:  # pragma: no cover - GUI error path
            messagebox.showerror("검증 결과 다운로드 실패", f"엑셀 파일 생성 중 오류가 발생했습니다.\n{exc}")
            return

        msg = f"검증 결과가 저장되었습니다.\n{saved_path}\n\n폴더를 여시겠습니까?"
        if messagebox.askyesno("검증 결과 다운로드", msg):
            self._open_folder(Path(saved_path).parent)

    # --------------------------- helpers --------------------------- #
    def _open_folder(self, folder: Path) -> None:
        try:
            os.startfile(folder)
        except OSError as exc:
            messagebox.showerror("폴더 열기 실패", f"폴더를 여는 중 오류가 발생했습니다.\n{exc}")
