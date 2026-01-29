# Re-export functions to maintain backward compatibility and simplify imports
from app.utils.excel_handler import (
    download_template_file,
    load_data_from_excel,
    export_to_excel,
    export_comparison_table
)
from app.utils.table_handler import (
    setup_table_header,
    render_table,
    re_sort_table,
    extract_data_from_table
)
from app.utils.diff_logic import generate_diff_report

# Main function used by CheckerPage
def make_table(table, file_path, sheet_name):
    """엑셀 파일을 읽어 테이블을 구성합니다."""
    setup_table_header(table)
    data = load_data_from_excel(file_path, sheet_name)
    render_table(table, data)
