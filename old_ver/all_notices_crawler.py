import requests
from bs4 import BeautifulSoup
from urllib.parse import urlparse, parse_qs
import openpyxl
import re
import os
from datetime import datetime

def get_user_input_url():
    """
    사용자로부터 크롤링할 URL을 입력받아 문자열로 반환하는 함수.
    """
    user_input = input("크롤링할 URL을 입력하세요: ")
    return user_input.strip()

def get_soup_by_page(user_input_url, page_no):
    """
    사용자 입력 URL을 파싱하고, pageNo를 page_no로 강제 세팅한 뒤
    해당 페이지의 BeautifulSoup 객체를 반환
    """
    parsed_url = urlparse(user_input_url)
    base_url = parsed_url.scheme + "://" + parsed_url.netloc + parsed_url.path

    query_dict = parse_qs(parsed_url.query)
    query_dict["pageNo"] = [str(page_no)]

    params = {}
    for key, value_list in query_dict.items():
        if len(value_list) == 1:
            params[key] = value_list[0]
        else:
            params[key] = value_list

    response = requests.get(base_url, params=params)
    response.encoding = 'utf-8'
    if response.status_code != 200:
        print(f"{page_no} 페이지 접근 실패, 상태코드:", response.status_code)
        return None

    return BeautifulSoup(response.text, "html.parser")

def get_last_page_number(soup):
    """
    BeautifulSoup 객체에서 페이지네이션 div를 찾고,
    끝 페이지( class='last' )의 goList(xxx)에서 xxx를 추출하여 반환.
    페이지네이션이 없으면 1 반환.
    """
    pagination_div = soup.find("div", class_="pagination")
    if not pagination_div:
        return 1

    last_link = pagination_div.find("a", class_="last")
    if last_link and last_link.get("href"):
        match = re.search(r"goList\((\d+)\)", last_link["href"])
        if match:
            return int(match.group(1))

    # last_link가 없으면, page 링크 중 가장 큰 번호 탐색
    page_links = pagination_div.find_all("a", class_="page")
    page_numbers = []
    for link in page_links:
        href_val = link.get("href", "")
        match = re.search(r"goList\((\d+)\)", href_val)
        if match:
            page_numbers.append(int(match.group(1)))

    if page_numbers:
        return max(page_numbers)
    return 1

def parse_bid_table(soup):
    """
    주어진 BeautifulSoup 객체에서 id='tblBidList' 테이블을 찾아
    각 행의 컬럼을 추출한 뒤 dict 리스트 형태로 반환.
    """
    table = soup.find("table", id="tblBidList")
    if not table:
        return []

    tbody = table.find("tbody")
    if not tbody:
        return []

    rows = tbody.find_all("tr")
    data_list = []
    for row in rows:
        tds = row.find_all("td")
        if len(tds) < 8:
            continue

        seq = tds[0].get_text(strip=True)          # 순번
        bid_type = tds[1].get_text(strip=True)     # 종류
        award_method = tds[2].get_text(strip=True) # 낙찰방법
        bid_title = tds[3].get_text(strip=True)    # 입찰공고명
        bid_deadline = tds[4].get_text(strip=True) # 입찰마감일
        status = tds[5].get_text(strip=True)       # 상태
        apt_name = tds[6].get_text(strip=True)     # 단지명
        reg_date = tds[7].get_text(strip=True)     # 공고일

        onclick_attr = tds[0].get("onclick", "")
        match = re.search(r"goView\('(.+?)'\)", onclick_attr)
        detail_id = match.group(1) if match else ""

        # 상세 페이지 링크는 K-APT 구조에 맞게 세팅
        detail_link = f"https://www.k-apt.go.kr/bid/bidDetail.do?bidNum={detail_id}"

        data_list.append({
            "순번": seq,
            "종류": bid_type,
            "낙찰방법": award_method,
            "입찰공고명": bid_title,
            "입찰마감일": bid_deadline,
            "상태": status,
            "단지명": apt_name,
            "공고일": reg_date,
            "상세정보링크": detail_link
        })

    return data_list

def crawl_all_pages(user_input_url):
    """
    1) 첫 페이지 soup 얻기 -> 마지막 페이지 번호 파악
    2) 1부터 마지막 페이지까지 순회하며 테이블 크롤링
    3) 결과를 데이터 리스트로 반환
    """
    first_page_soup = get_soup_by_page(user_input_url, page_no=1)
    if not first_page_soup:
        print("첫 페이지 로드 실패")
        return []

    last_page = get_last_page_number(first_page_soup)
    print(f"확인된 마지막 페이지: {last_page}")

    all_data = []
    for page in range(1, last_page + 1):
        print(f"{page}/{last_page} 페이지 처리 중...")
        page_soup = get_soup_by_page(user_input_url, page_no=page)
        if not page_soup:
            print(f"{page} 페이지 로드 실패. 넘어갑니다.")
            continue

        page_data = parse_bid_table(page_soup)
        all_data.extend(page_data)

    print(f"총 {len(all_data)}개의 데이터 수집 완료")
    return all_data

def make_unique_filename(base_name="입찰공고_전체페이지"):
    """
    폴더 '입찰공고_전체페이지'가 없다면 생성.
    파일명: base_name_YYYYMMDD_HHMMSS.xlsx 형태.
    중복 시 _1, _2 등 넘버링.
    """
    # 1) 폴더 명
    folder_name = "입찰공고_전체페이지"
    os.makedirs(folder_name, exist_ok=True)  # 폴더가 없으면 생성

    # 2) 현재 시각 -> "YYYYMMDD_HHMMSS"
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

    # 3) 파일 이름
    filename = f"{base_name}_{timestamp}.xlsx"
    full_path = os.path.join(folder_name, filename)  # "입찰공고_전체페이지/입찰공고_전체페이지_YYYYMMDD_HHMMSS.xlsx"

    # 4) 이미 존재하면 _1, _2 증가
    if os.path.exists(full_path):
        counter = 1
        while True:
            new_filename = f"{base_name}_{timestamp}_{counter}.xlsx"
            new_full_path = os.path.join(folder_name, new_filename)
            if not os.path.exists(new_full_path):
                full_path = new_full_path
                break
            counter += 1

    return full_path

def save_to_excel(data_list, filename):
    """
    수집한 딕셔너리 리스트를 엑셀로 저장
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "입찰공고"

    headers = ["순번", "종류", "낙찰방법", "입찰공고명", "입찰마감일", "상태", "단지명", "공고일", "상세정보링크"]
    ws.append(headers)

    for item in data_list:
        row = [
            item["순번"],
            item["종류"],
            item["낙찰방법"],
            item["입찰공고명"],
            item["입찰마감일"],
            item["상태"],
            item["단지명"],
            item["공고일"],
            item["상세정보링크"]
        ]
        ws.append(row)

    wb.save(filename)
    print(f"엑셀 파일 생성 완료: {filename}")

def main():
    # 1) 사용자에게 URL 입력 받기
    user_url = get_user_input_url()
    # 2) 전체 페이지에 대해 크롤링 (데이터 리스트)
    all_data = crawl_all_pages(user_url)

    if not all_data:
        print("가져올 데이터가 없습니다. 종료합니다.")
        return

    # 3) 파일명 생성 (중복 방지 + 타임스탬프), 폴더 생성 포함
    final_filepath = make_unique_filename("입찰공고_전체페이지")
    # 4) 엑셀로 저장
    save_to_excel(all_data, final_filepath)

if __name__ == "__main__":
    main()
