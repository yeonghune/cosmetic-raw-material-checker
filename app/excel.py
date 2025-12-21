from __future__ import annotations

from enum import Enum
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict, List, Sequence

from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill


FIXED_HEADER = ("RM", "% RM/FP", "INCI", "% INCI/RM")


class Header(str, Enum):
    RM = "RM"
    RM_FP = "% RM/FP"
    INCI = "INCI"
    INCI_RM = "% INCI/RM"

RESULT_HEADER = (
    "RM",
    "INCI",
    "% RM/FP (TABLE 1)",
    "% RM/FP (TABLE 2)",
    "% INCI/RM (TABLE 1)",
    "% INCI/RM (TABLE 2)",
)

@dataclass
class TablePayload:
    header: tuple[str] | None = None
    data: List[Dict[str, Any]] = field(default_factory=list)
    is_empty: bool = True
    header_len: int | None = None
    data_len: int | None = None


@dataclass
class ValidationPayload:
    rm: str
    inci: str
    rm_fp_table1: str
    rm_fp_table2: str
    inci_rm_table1: str
    inci_rm_table2: str


def download_template_file(output_path: str | Path = "다운로드/output.xlsx") -> Path:
    wb = Workbook()
    ws1 = wb.active
    ws1.title = "Table1"

    ws2 = wb.create_sheet(title="Table2")
    for col_idx, header in enumerate(FIXED_HEADER, start=1):
        ws1.cell(row=1, column=col_idx).value = header
        ws2.cell(row=1, column=col_idx).value = header

    output = Path(output_path)
    output.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output)
    return output


def upload_template_file(input_path: str | Path) -> tuple[TablePayload, TablePayload]:
    wb = load_workbook(input_path)

    table1 = wb["Table1"]
    table2 = wb["Table2"]
    result1 = _validate_input_excel_file(table1)
    result2 = _validate_input_excel_file(table2)

    return result1, result2


def table_cross_validation(table1: TablePayload, table2: TablePayload) -> tuple[bool, str]:
    """Perform simple cross validation between two tables."""
    if table1.is_empty:
        return False, "테이블1 시트가 비었습니다."
    
    if table2.is_empty:
        return False, "테이블2 시트가 비었습니다."

    if table1.header != FIXED_HEADER or table2.header != FIXED_HEADER:
        return False, "테이블1, 테이블2 시트의 헤더가 일치하지 않습니다."

    if table1.header_len != table2.header_len:
        return False, "테이블1, 테이블2 시트의 헤더가 다릅니다."

    return True, "업로드를 수행했습니다."


def data_merge(table1: TablePayload, table2: TablePayload) -> list[ValidationPayload]:
    """(RM, INCI) 키 기준으로 table1, table2를 inner join 한다.

    - 기본 키: (RM, INCI)
    - 두 테이블 모두에 존재하는 키만 결과에 포함된다.
    - 결과에는 "X" 와 같은 placeholder를 사용하지 않는다.
    """

    def _to_str(value: Any) -> str:
        if value is None:
            return ""
        return str(value)

    def _key_from_row(row: Dict[str, Any]) -> tuple[str, str]:
        rm = _to_str(row.get(Header.RM))
        inci = _to_str(row.get(Header.INCI))
        return rm, inci

    # (RM, INCI) -> row dict for table2
    index2: Dict[tuple[str, str], Dict[str, Any]] = {}
    for row in table2.data:
        key = _key_from_row(row)
        if key not in index2:
            index2[key] = row

    result: list[ValidationPayload] = []
    seen_keys: set[tuple[str, str]] = set()

    # table1 순서를 기준으로 inner join 수행
    for row1 in table1.data:
        key = _key_from_row(row1)
        # 동일 키가 여러 번 나와도 한 번만 처리
        if key in seen_keys:
            continue
        seen_keys.add(key)

        row2 = index2.get(key)
        if row2 is None:
            # inner join이므로 table2에 없는 키는 결과에서 제외
            continue

        rm_str, inci_str = key
        payload = ValidationPayload(
            rm=rm_str,
            inci=inci_str,
            rm_fp_table1=_to_str(row1.get(Header.RM_FP)),
            rm_fp_table2=_to_str(row2.get(Header.RM_FP)),
            inci_rm_table1=_to_str(row1.get(Header.INCI_RM)),
            inci_rm_table2=_to_str(row2.get(Header.INCI_RM)),
        )

        result.append(payload)

    # 결과도 RM, INCI 기준으로 정렬
    result.sort(key=lambda p: (p.rm, p.inci))

    return result


def _validate_input_excel_file(ws) -> TablePayload:
    rows = ws.iter_rows(values_only=True)
    result = TablePayload()
    try:
        header = next(rows)
    except StopIteration:
        return result

    value = [dict(zip(header, row)) for row in rows]

    # RM 오름차순, RM 같으면 INCI 오름차순으로 정렬
    def _sort_key(row: Dict[str, Any]) -> tuple[str, str]:
        rm = row.get(Header.RM)
        inci = row.get(Header.INCI)
        rm_key = "" if rm is None else str(rm)
        inci_key = "" if inci is None else str(inci)
        return rm_key, inci_key

    value.sort(key=_sort_key)

    result.header = header
    result.data = value

    result.header_len = len(header)
    result.data_len = len(value)
    result.is_empty = False

    return result


def export_validation_result(
    output_path: str | Path,
    table1: TablePayload | None,
    table2: TablePayload | None,
    rows: Sequence[ValidationPayload],
    unique_keys_table1: Sequence[tuple[str, str]] | None = None,
    unique_keys_table2: Sequence[tuple[str, str]] | None = None,
) -> Path:
    """검증 결과를 새로운 엑셀 파일로 내보낸다.

    - Table1 시트: 테이블1 데이터, table1에만 있는 (RM, INCI) row는 빨간 배경.
    - Table2 시트: 테이블2 데이터, table2에만 있는 (RM, INCI) row는 빨간 배경.
    - Result 시트: inner join 결과, INCI/RM 다르면 빨간, RM/FP만 다르면 노란 배경.
    """

    output = Path(output_path)

    red_fill = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")
    yellow_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")

    def _to_str(value: Any) -> str:
        if value is None:
            return ""
        return str(value)

    def _write_table_sheet(
        ws,
        payload: TablePayload | None,
        unique_keys: Sequence[tuple[str, str]] | None,
    ) -> None:
        # 헤더
        if payload and not payload.is_empty and payload.header:
            headers = list(payload.header)
        else:
            headers = list(FIXED_HEADER)

        for col_idx, header in enumerate(headers, start=1):
            ws.cell(row=1, column=col_idx, value=str(header))

        unique_set = set(unique_keys or [])

        if not payload or payload.is_empty or not payload.data:
            return

        for row_idx, row_dict in enumerate(payload.data, start=2):
            rm = _to_str(row_dict.get(Header.RM))
            inci = _to_str(row_dict.get(Header.INCI))
            is_unique = (rm, inci) in unique_set

            for col_idx, header in enumerate(headers, start=1):
                value = row_dict.get(header, "")
                cell = ws.cell(row=row_idx, column=col_idx, value=value)
                if is_unique:
                    cell.fill = red_fill

    def _result_row_tag(payload: ValidationPayload) -> str:
        """결과 행 색상 구분: 'red', 'yellow', ''."""
        if payload.inci_rm_table1 != payload.inci_rm_table2:
            return "red"
        if payload.rm_fp_table1 != payload.rm_fp_table2:
            return "yellow"
        return ""

    wb = Workbook()

    ws1 = wb.active
    ws1.title = "Table1"
    ws2 = wb.create_sheet(title="Table2")
    ws_result = wb.create_sheet(title="Result")

    _write_table_sheet(ws1, table1, unique_keys_table1)
    _write_table_sheet(ws2, table2, unique_keys_table2)

    # Result 시트 헤더
    for col_idx, header in enumerate(RESULT_HEADER, start=1):
        ws_result.cell(row=1, column=col_idx, value=header)

    # Result 시트 데이터
    for row_idx, payload in enumerate(rows, start=2):
        values = [
            payload.rm,
            payload.inci,
            payload.rm_fp_table1,
            payload.rm_fp_table2,
            payload.inci_rm_table1,
            payload.inci_rm_table2,
        ]

        tag = _result_row_tag(payload)

        for col_idx, value in enumerate(values, start=1):
            cell = ws_result.cell(row=row_idx, column=col_idx, value=value)
            if tag == "red":
                cell.fill = red_fill
            elif tag == "yellow":
                cell.fill = yellow_fill

    output.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output)
    return output
