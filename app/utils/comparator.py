from typing import List, Tuple
from itertools import zip_longest

# (Original_Text_A, Original_Text_B, Status)
# Status: "MATCH", "DIFF"
ComparisonRow = Tuple[str, str, str]

def compare_ingredients(list1: List[str], list2: List[str]) -> List[ComparisonRow]:
    """
    두 성분 리스트를 순서대로 1:1 비교합니다.
    
    Logic:
    1. 두 리스트를 순서대로 나란히 배치 (zip_longest)
    2. 소문자 변환 및 공백 제거 후 단순 비교 (==)
    3. 다르면 DIFF, 같으면 MATCH 반환
    """
    
    rows: List[ComparisonRow] = []
    
    # 길이가 다른 경우 빈 문자열로 채움
    for item1, item2 in zip_longest(list1, list2, fillvalue=""):
        val1 = item1 if item1 else ""
        val2 = item2 if item2 else ""
        
        norm1 = val1.strip().lower()
        norm2 = val2.strip().lower()
        
        if norm1 == norm2:
            status = "MATCH"
        else:
            status = "DIFF"
            
        rows.append((val1, val2, status))
        
    return rows
