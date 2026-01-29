from app.models import DiffType, DiffItem, IngredientRow

def generate_diff_report(source_data: list[IngredientRow], ref_data: list[IngredientRow]) -> list[DiffItem]:
    """
    Source(내꺼) 기준으로 Ref(상대방)와 비교하여 스타일링(Diff) 정보를 생성합니다.
    """
    diffs = []
    
    # 데이터 구조화
    struct_source = _parse_structured_data_from_list(source_data)
    struct_ref = _parse_structured_data_from_list(ref_data)

    # 비교 루프
    for rm_name, rm_info in struct_source.items():
        # Case 1.3: 내 RM이 상대방에게 아예 없음 -> 전체 행 배경 빨강
        if rm_name not in struct_ref:
            for r in rm_info["rows"]:
                for c in range(4): # 0~3 컬럼 전체
                    diffs.append(DiffItem(r, c, DiffType.MISSING_ROW))
            continue

        ref_rm = struct_ref[rm_name]

        # Case 1.1: RM 함량이 다름 -> 첫 번째 행의 % 컬럼 글자 빨강
        if rm_info["percent"] != ref_rm["percent"]:
            first_row = rm_info["rows"][0]
            diffs.append(DiffItem(first_row, 1, DiffType.CONTENT_MISMATCH))
            
        # INCI 레벨 비교
        for inci_name, inci_info in rm_info["incis"].items():
            # Case 1.4: 내 INCI가 상대방 RM 안에 없음 -> 부분 배경 빨강
            if inci_name not in ref_rm["incis"]:
                r = inci_info["row"]
                diffs.append(DiffItem(r, 2, DiffType.MISSING_INCI))
                diffs.append(DiffItem(r, 3, DiffType.MISSING_INCI))
                continue
            
            # Case 1.2: INCI 함량이 다름 -> 글자 빨강
            ref_inci = ref_rm["incis"][inci_name]
            if inci_info["percent"] != ref_inci["percent"]:
                diffs.append(DiffItem(inci_info["row"], 3, DiffType.CONTENT_MISMATCH))
                
    return diffs

def _parse_structured_data_from_list(data_list: list[IngredientRow]):
    """IngredientRow 리스트를 딕셔너리 구조로 변환"""
    data = {}
    for i, row in enumerate(data_list):
        rm_name = row.rm_name
        
        if rm_name not in data:
            data[rm_name] = {
                "percent": row.rm_percent,
                "rows": [], # Row Index
                "incis": {}
            }
        
        data[rm_name]["rows"].append(i)
        
        if row.inci_name:
            # 대소문자 무시 비교를 위해 Key를 소문자로 변환
            inci_key = row.inci_name.strip().lower()
            data[rm_name]["incis"][inci_key] = {
                "percent": row.inci_percent,
                "row": i
            }
    return data
