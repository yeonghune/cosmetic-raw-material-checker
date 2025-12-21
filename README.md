# 엑셀 템플릿 검증기 (Windows)

화장품 원료/성분 비율을 담은 두 개의 엑셀 시트를 비교·검증하는
윈도우 데스크톱 앱입니다. Table1과 Table2를 비교해 불일치 항목을
하이라이트하고, 결과를 새 엑셀 파일로 저장합니다.

## 주요 기능
- 엑셀 템플릿 다운로드 (Table1, Table2)
- 템플릿 업로드 및 두 테이블 미리보기
- (RM, INCI) 키 기준 검증 및 비율 비교
- 불일치 하이라이트
  - INCI/RM 불일치 = 빨강
  - RM/FP 불일치 = 노랑
- 동일 하이라이트로 결과 엑셀 저장

## 용어 정의
- RM (Raw Material): 원료명
- INCI: 성분명(국제 명칭)
- % RM/FP: 완성품 대비 해당 원료 비율
- % INCI/RM: 원료 내 해당 성분 비율

## 동작 흐름
1. 템플릿 다운로드 후 Table1 / Table2 작성
2. 파일 업로드
3. 검증 실행
   - (RM, INCI) 기준 inner join
   - % RM/FP, % INCI/RM 비교
4. 하이라이트 확인 후 결과 저장

## 요구사항
- Python 3.10+
- openpyxl

의존성 설치:
```
pip install openpyxl
```

## 실행 (개발)
프로젝트 루트에서:
```
python -m app.main
```
