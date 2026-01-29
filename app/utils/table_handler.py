from PyQt5.QtWidgets import QTableWidgetItem, QHeaderView
from app.models import IngredientRow
from app.utils.excel_handler import load_data_from_excel

FIXED_HEADER = ("RM", "% RM/FP", "INCI", "% INCI/RM")

def setup_table_header(table):
    """테이블의 헤더와 컬럼 설정을 초기화합니다."""
    table.setColumnCount(len(FIXED_HEADER))
    table.setHorizontalHeaderLabels(FIXED_HEADER)
    
    header = table.horizontalHeader()
    header.setSectionResizeMode(QHeaderView.Interactive)
    header.setSectionResizeMode(2, QHeaderView.Stretch)

def extract_data_from_table(table) -> list[IngredientRow]:
    """QTableWidget의 내용을 읽어 IngredientRow 리스트로 변환합니다."""
    rows = table.rowCount()
    raw_data = []
    
    for r in range(rows):
        item_rm = table.item(r, 0)
        item_rm_pct = table.item(r, 1)
        item_inci = table.item(r, 2)
        item_inci_pct = table.item(r, 3)

        raw_data.append(IngredientRow(
            rm_name=item_rm.text() if item_rm else "",
            rm_percent=item_rm_pct.text() if item_rm_pct else "",
            inci_name=item_inci.text() if item_inci else "",
            inci_percent=item_inci_pct.text() if item_inci_pct else ""
        ))
    return raw_data

def render_table(table, data_list: list[IngredientRow]):
    """데이터 리스트를 테이블에 그리고, 자동 병합을 수행합니다."""
    # 1. 정렬 (RM 이름 -> INCI 이름 순)
    if data_list:
        data_list.sort(key=lambda x: (x.rm_name.lower(), x.inci_name.lower()))

    # 2. 초기화
    table.clearContents()
    table.clearSpans()  # [신규] 기존 병합 정보 초기화 (필수)
    table.setRowCount(len(data_list))
    
    if not data_list:
        return

    # 3. 데이터 쓰기 및 병합
    merge_start_idx = 0
    prev_rm = data_list[0].rm_name

    for i, item in enumerate(data_list):
        # 아이템 생성 및 삽입
        table.setItem(i, 0, QTableWidgetItem(item.rm_name))
        table.setItem(i, 1, QTableWidgetItem(item.rm_percent))
        table.setItem(i, 2, QTableWidgetItem(item.inci_name))
        table.setItem(i, 3, QTableWidgetItem(item.inci_percent))

        # RM이 바뀌면 직전 그룹 병합 처리
        if item.rm_name != prev_rm:
            _apply_merge(table, merge_start_idx, i - merge_start_idx)
            
            # 상태 업데이트
            prev_rm = item.rm_name
            merge_start_idx = i

    # 마지막 그룹 병합 처리
    _apply_merge(table, merge_start_idx, len(data_list) - merge_start_idx)

def _apply_merge(table, start_row, span_count):
    """내부 헬퍼: 조건에 맞으면 셀 병합 수행"""
    if span_count > 1:
        table.setSpan(start_row, 0, span_count, 1) # RM 컬럼
        table.setSpan(start_row, 1, span_count, 1) # % RM/FP 컬럼

def re_sort_table(table):
    """현재 테이블 내용을 읽어서 다시 정렬하고 그립니다."""
    current_data = extract_data_from_table(table)
    render_table(table, current_data)

def make_table(table, file_path, sheet_name):
    """엑셀 파일을 읽어 테이블을 구성합니다."""
    setup_table_header(table)
    data = load_data_from_excel(file_path, sheet_name)
    render_table(table, data)
