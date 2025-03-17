import requests
import pandas as pd
from bs4 import BeautifulSoup
import os
from datetime import datetime

def generate_output_filename():
    """
    '입찰공고_전체페이지_상세정보_YYYYMMDD_HHMMSS.xlsx' 형태 파일명을 만든 뒤,
    이미 동일 이름의 파일이 존재하면 (1), (2)를 붙여 유니크하게 만듭니다.
    """
    now_str = datetime.now().strftime("%Y%m%d_%H%M%S")
    base_filename = f"입찰공고_전체페이지_상세정보_{now_str}"
    ext = ".xlsx"
    output_filename = base_filename + ext

    counter = 1
    while os.path.exists(output_filename):
        output_filename = f"{base_filename}_{counter}{ext}"
        counter += 1

    return output_filename

def crawl_detail_page(url: str) -> dict:
    """
    주어진 상세정보 링크(url)에 GET 요청을 보내서,
    1) 첫 번째 테이블 (class="contTbl txtC")에서 
       [주택관리업자, 단지명, 관리사무소 주소, 전화번호, 팩스번호, 동수, 세대수]
    2) 두 번째 테이블 (class="contTbl" but txtC 제외)에서 
       [입찰번호, 입찰방법, 입찰서 제출 마감일, 입찰제목, 긴급입찰여부, 입찰종류,
        낙찰방법, 입찰분류, 신용평가등급확인서 제출여부, 현장설명, 관리(공사용역) 실적증명서 제출여부,
        현장설명일시, 현장설명장소, 서류제출마감일, 입찰보증금, 지급조건, 내용]

    ※ '파일첨부' 항목은 크롤링하지 않음.
    """
    response = requests.get(url)
    response.raise_for_status()
    soup = BeautifulSoup(response.text, "lxml")

    data = {
        '주택관리업자': '',
        '단지명': '',
        '관리사무소 주소': '',
        '전화번호': '',
        '팩스번호': '',
        '동수': '',
        '세대수': '',
        '입찰번호': '',
        '입찰방법': '',
        '입찰서 제출 마감일': '',
        '입찰제목': '',
        '긴급입찰여부': '',
        '입찰종류': '',
        '낙찰방법': '',
        '입찰분류': '',
        '신용평가등급확인서 제출여부': '',
        '현장설명': '',
        '관리(공사용역) 실적증명서 제출여부': '',
        '현장설명일시': '',
        '현장설명장소': '',
        '서류제출마감일': '',
        '입찰보증금': '',
        '지급조건': '',
        '내용': ''
    }

    # 첫 번째 테이블 (단지 정보)
    table1 = soup.find("table", class_="contTbl txtC")
    if table1:
        tbody1 = table1.find("tbody")
        if tbody1:
            t1_rows = tbody1.find_all("tr", recursive=False)
            if len(t1_rows) >= 2:
                header_ths = t1_rows[0].find_all("th", recursive=False)
                value_tds = t1_rows[1].find_all("td", recursive=False)

                headers = [th.get_text(strip=True) for th in header_ths]
                values = [td.get_text(strip=True) for td in value_tds]

                for h, v in zip(headers, values):
                    if h in data:
                        data[h] = v

    # 두 번째 테이블 (입찰 정보, 파일첨부 제외)
    table2_candidates = soup.find_all("table", class_="contTbl")
    table2 = None
    for t in table2_candidates:
        # "txtC" class 가 없는 table 만 검색
        if "txtC" not in t.get("class", []):
            table2 = t
            break

    if table2:
        tbody2 = table2.find("tbody")
        if tbody2:
            t2_rows = tbody2.find_all("tr", recursive=False)
            for row in t2_rows:
                th_list = row.find_all("th", recursive=False)
                td_list = row.find_all("td", recursive=False)

                pair_count = min(len(th_list), len(td_list))
                for i in range(pair_count):
                    key = th_list[i].get_text(strip=True)
                    val = td_list[i].get_text(strip=True)

                    # '파일첨부' 항목은 무시
                    if key == '파일첨부':
                        continue

                    if key in data:
                        data[key] = val

    return data

def main():
    # 1) 입력 엑셀 파일 (상세정보링크 컬럼 포함)
    #    예시: "입찰공고_전체페이지/입찰공고_전체페이지_20250223_160924.xlsx"
    input_excel_path = "입찰공고_전체페이지/입찰공고_전체페이지_20250223_160924.xlsx"

    # 2) 아웃풋을 저장할 폴더가 없으면 생성
    output_dir = "입찰공고_전체페이지_상세정보"
    os.makedirs(output_dir, exist_ok=True)

    # 3) 결과 엑셀파일 경로 설정 (폴더 + 파일명)
    output_filename = generate_output_filename()
    output_excel_path = os.path.join(output_dir, output_filename)

    # 4) 입력 엑셀 읽기
    #    (기존 목록에는 "순번", "입찰공고명", "상세정보링크" 등이 있을 것)
    df_input = pd.read_excel(input_excel_path)
    total_count = len(df_input)

    # 결과를 담을 리스트
    results = []

    # 5) 크롤링 반복
    for idx, row in df_input.iterrows():
        detail_url = row.get('상세정보링크')
        if not detail_url:
            print(f"[{idx+1}/{total_count}] 링크가 비어있습니다. (건너뛰기)")
            # 기존 행 데이터(목록)만 살리고, 상세 컬럼은 채우지 않음
            integrated_data = row.to_dict()  # 원본 정보를 dict로
            results.append(integrated_data)
            continue

        print(f"[{idx+1}/{total_count}] 크롤링 시도 중: {detail_url}")

        max_retries = 3
        crawled_data = None
        for attempt in range(1, max_retries + 1):
            try:
                crawled_data = crawl_detail_page(detail_url)
                print(f"  [성공] (시도 {attempt}/{max_retries}) {detail_url}")
                break
            except Exception as e:
                print(f"  [오류] (시도 {attempt}/{max_retries}) {detail_url}: {e}")

        # (A) 정상 크롤링 성공
        if crawled_data:
            # 기존 행 정보 + 상세정보 merge
            integrated_data = row.to_dict()  # 원본 목록 정보를 dict로 변환
            integrated_data.update(crawled_data)  # 상세정보 덮어씌우기
            results.append(integrated_data)

        # (B) 3번 모두 실패
        else:
            print(f"  [실패] 3번 재시도 후 포기: {detail_url}")
            # 원본 목록 정보 + 상세정보를 FAILED로
            failed_data = {
                '주택관리업자': 'FAILED', '단지명': 'FAILED', '관리사무소 주소': 'FAILED',
                '전화번호': 'FAILED', '팩스번호': 'FAILED', '동수': 'FAILED', '세대수': 'FAILED',
                '입찰번호': 'FAILED', '입찰방법': 'FAILED', '입찰서 제출 마감일': 'FAILED',
                '입찰제목': 'FAILED', '긴급입찰여부': 'FAILED', '입찰종류': 'FAILED',
                '낙찰방법': 'FAILED', '입찰분류': 'FAILED', '신용평가등급확인서 제출여부': 'FAILED',
                '현장설명': 'FAILED', '관리(공사용역) 실적증명서 제출여부': 'FAILED',
                '현장설명일시': 'FAILED', '현장설명장소': 'FAILED', '서류제출마감일': 'FAILED',
                '입찰보증금': 'FAILED', '지급조건': 'FAILED', '내용': 'FAILED'
            }
            integrated_data = row.to_dict()
            integrated_data.update(failed_data)  # 상세 컬럼만 FAILED로 채움
            results.append(integrated_data)

    # 6) 결과 엑셀 저장
    if results:
        df_result = pd.DataFrame(results)
        df_result.to_excel(output_excel_path, index=False)
        print(f"\n크롤링 완료! 총 {len(results)}건 데이터를 '{output_excel_path}' 에 저장했습니다.")
    else:
        print("\n크롤링 결과가 없습니다.")

if __name__ == "__main__":
    main()
