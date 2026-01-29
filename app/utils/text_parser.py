import re

def parse_ingredients(text: str) -> list[str]:
    """
    텍스트를 파싱하여 성분 리스트로 변환합니다.
    
    Rules:
    1. 줄바꿈(\n)은 공백으로 치환
    2. 쉼표(,)로 분리하되, 다음의 경우는 분리하지 않음:
       - 쉼표 바로 뒤에 숫자([0-9])가 오는 경우 (예: 1,2-Hexanediol)
    3. 앞뒤 공백 제거
    """
    if not text:
        return []
        
    # 1. 줄바꿈을 공백으로 치환
    cleaned_text = text.replace('\n', ' ')
    
    # 2. 쉼표로 분리 (Regex 사용)
    # ,(?![0-9]) : 쉼표 뒤에 숫자가 오지 않는 경우에만 매칭
    # 예: "Water, Glycerin" -> 매칭됨
    # 예: "1,2-Hexanediol" -> 매칭안됨 (뒤에 2가 옴)
    # 예: "Peptide-1, 2-..." -> 매칭됨 (뒤에 공백이 옴)
    
    # Note: 사용자의 "붙어있는 특수 문자" 요구사항은 구체적인 예시(1,2-...)를 기반으로
    # 화학명에서 흔히 쓰이는 "숫자 앞 쉼표"를 예외처리하는 것으로 해석함.
    # 필요 시 정규표현식을 수정하여 확장 가능.
    
    pattern = r',(?![0-9])'
    
    raw_ingredients = re.split(pattern, cleaned_text)
    
    # 3. 정제 (공백 제거 및 빈 항목 제외)
    ingredients = [
        item.strip() 
        for item in raw_ingredients 
        if item.strip()
    ]
    
    return ingredients
