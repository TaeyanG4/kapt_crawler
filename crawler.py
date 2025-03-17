import re
import requests
from bs4 import BeautifulSoup
from urllib.parse import urlparse, parse_qs

class BaseCrawler:
    """
    기본 크롤러 클래스: URL 요청 및 BeautifulSoup 객체 생성을 담당합니다.
    """
    def __init__(self, base_url):
        self.base_url = base_url

    def fetch_page(self, url, params=None):
        try:
            response = requests.get(url, params=params, timeout=10)
            response.encoding = 'utf-8'
            if response.status_code != 200:
                return None
            return response.text
        except Exception as e:
            # 네트워크 오류 발생 시 None 반환
            return None

    def get_soup(self, url, params=None):
        html = self.fetch_page(url, params)
        if html:
            # lxml 파서를 일관되게 사용
            return BeautifulSoup(html, "lxml")
        return None

class SummaryCrawler(BaseCrawler):
    """
    목록 데이터 크롤러: 지정된 URL에서 페이지별 데이터를 수집합니다.
    """
    def __init__(self, base_url, page_type_index=0):
        super().__init__(base_url)
        self.page_type_index = page_type_index

    def get_soup_by_page(self, user_input_url, page_no):
        parsed_url = urlparse(user_input_url)
        base_url = f"{parsed_url.scheme}://{parsed_url.netloc}{parsed_url.path}"
        query_dict = parse_qs(parsed_url.query)
        query_dict["pageNo"] = [str(page_no)]
        # 단일 값인 경우 그냥 문자열로, 다중 값이면 리스트로 처리
        params = {key: (value[0] if len(value)==1 else value) for key, value in query_dict.items()}
        return self.get_soup(base_url, params)

    def get_last_page_number(self, soup):
        """
        페이지 네비게이션 영역을 파싱하여 마지막 페이지 번호를 반환합니다.
        """
        pagination_div = soup.find("div", class_="pagination")
        if not pagination_div:
            return 1

        last_link = pagination_div.find("a", class_="last")
        if last_link and last_link.get("href"):
            match = re.search(r"goList\((\d+)\)", last_link["href"])
            if match:
                return int(match.group(1))

        page_links = pagination_div.find_all("a", class_="page")
        page_numbers = []
        for link in page_links:
            href_val = link.get("href", "")
            match = re.search(r"goList\((\d+)\)", href_val)
            if match:
                page_numbers.append(int(match.group(1)))
        return max(page_numbers) if page_numbers else 1

    def parse_bid_table(self, soup):
        data_list = []
        if self.page_type_index == 0:
            table = soup.find("table", {"class": "contTbl txtC"})
        else:
            table = soup.find("table", id="tblBidList")
        
        if not table:
            return data_list

        tbody = table.find("tbody")
        if not tbody:
            return data_list

        rows = tbody.find_all("tr")
        for row in rows:
            tds = row.find_all("td")
            if self.page_type_index == 0:
                if len(tds) < 7:
                    continue
                seq = tds[0].get_text(strip=True)
                apt_name = tds[1].get_text(strip=True)
                contract_company = tds[2].get_text(strip=True)
                contract_name = tds[3].get_text(strip=True)
                contract_date = tds[4].get_text(strip=True)
                contract_amount = tds[5].get_text(strip=True)
                contract_period = tds[6].get_text(strip=True)
                
                detail_link = ""
                onclick_attr = tds[0].get("onclick", "")
                match = re.search(r"goView\('(.+?)'\)", onclick_attr)
                detail_id = match.group(1) if match else ""
                if detail_id:
                    detail_link = f"https://www.k-apt.go.kr/bid/privateContractDetail.do?pcNum={detail_id}"

                data_list.append({
                    "순번": seq,
                    "단지명": apt_name,
                    "계약업체": contract_company,
                    "계약명": contract_name,
                    "계약일": contract_date,
                    "계약금액": contract_amount,
                    "계약기간": contract_period,
                    "상세정보링크": detail_link
                })
            else:
                if len(tds) < 8:
                    continue
                seq = tds[0].get_text(strip=True)
                bid_type = tds[1].get_text(strip=True)
                award_method = tds[2].get_text(strip=True)
                bid_title = tds[3].get_text(strip=True)
                bid_deadline = tds[4].get_text(strip=True)
                status = tds[5].get_text(strip=True)
                apt_name = tds[6].get_text(strip=True)
                reg_date = tds[7].get_text(strip=True)

                onclick_attr = tds[0].get("onclick", "")
                match = re.search(r"goView\('(.+?)'\)", onclick_attr)
                detail_id = match.group(1) if match else ""
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

    def crawl_all_pages(self, user_input_url, log_callback=None, max_items=50):
        def _log(msg):
            if log_callback:
                log_callback(msg)

        first_page_soup = self.get_soup_by_page(user_input_url, page_no=1)
        if not first_page_soup:
            _log("첫 페이지 로드 실패")
            return []
        last_page = self.get_last_page_number(first_page_soup)
        _log(f"확인된 마지막 페이지: {last_page}")

        all_data = []
        for page in range(1, last_page + 1):
            _log(f"{page}/{last_page} 페이지 처리 중...")
            page_soup = self.get_soup_by_page(user_input_url, page_no=page)
            if not page_soup:
                _log(f"{page} 페이지 로드 실패. 넘어갑니다.")
                continue

            page_data = self.parse_bid_table(page_soup)
            all_data.extend(page_data)
            if len(all_data) >= max_items:
                all_data = all_data[:max_items]
                break

        _log(f"총 {len(all_data)}개의 데이터 수집 완료")
        return all_data

class DetailCrawler(BaseCrawler):
    """
    상세정보 크롤러: 항목의 상세페이지에서 추가 정보를 수집합니다.
    """
    def __init__(self, page_type_index=0):
        self.page_type_index = page_type_index
        # base_url은 사용하지 않으므로 빈 문자열로 초기화
        super().__init__(base_url="")

    def crawl_detail_page(self, url):
        try:
            response = requests.get(url, timeout=10)
            response.raise_for_status()
        except Exception as e:
            raise Exception(f"상세 페이지 로드 실패: {e}")
        # 일관되게 lxml 파서를 사용
        soup = BeautifulSoup(response.text, "lxml")
        if self.page_type_index == 0:
            data = {
                '주택관리업자': '',
                '아파트명': '',
                '관리사무소 주소': '',
                '전화번호': '',
                '팩스번호': '',
                '동수': '',
                '세대수': '',
                '계약번호': '',
                '계약명': '',
                '계약업체명': '',
                '업체대표자명': '',
                '업체전화번호': '',
                '사업자등록번호': '',
                '업체주소': '',
                '계약(예정)일': '',
                '계약금액': '',
                '계약기간': '',
                '등록일': '',
                '분류': '',
                '수의계약 체결사유': ''
            }
            mapping = {
                "주택관리업자": "주택관리업자",
                "아파트명": "아파트명",
                "단지명": "아파트명",
                "관리사무소 주소": "관리사무소 주소",
                "전화번호": "전화번호",
                "팩스번호": "팩스번호",
                "동수": "동수",
                "세대수": "세대수",
                "계약번호": "계약번호",
                "계약명": "계약명",
                "계약업체명": "계약업체명",
                "업체대표자명": "업체대표자명",
                "업체전화번호": "업체전화번호",
                "사업자등록번호": "사업자등록번호",
                "업체주소": "업체주소",
                "계약(예정)일": "계약(예정)일",
                "계약금액": "계약금액",
                "계약기간": "계약기간",
                "등록일": "등록일",
                "분 류": "분류",
                "분류": "분류",
                "수의계약 체결사유": "수의계약 체결사유"
            }
            table_common = soup.find("table", class_="contTbl txtC")
            if table_common:
                tbody = table_common.find("tbody")
                if tbody:
                    row = tbody.find("tr")
                    if row:
                        cells = row.find_all("td")
                        if len(cells) >= 7:
                            data['주택관리업자'] = cells[0].get_text(strip=True)
                            data['아파트명'] = cells[1].get_text(strip=True)
                            data['관리사무소 주소'] = cells[2].get_text(strip=True)
                            data['전화번호'] = cells[3].get_text(strip=True)
                            data['팩스번호'] = cells[4].get_text(strip=True)
                            data['동수'] = cells[5].get_text(strip=True)
                            data['세대수'] = cells[6].get_text(strip=True)
            table_contract = None
            tables = soup.find_all("table", class_="contTbl")
            for t in tables:
                if "txtC" not in t.get("class", []):
                    table_contract = t
                    break
            if table_contract:
                tbody = table_contract.find("tbody")
                if tbody:
                    rows = tbody.find_all("tr")
                    for row in rows:
                        cells = row.find_all(["th", "td"])
                        if len(cells) < 2:
                            continue
                        for i in range(0, len(cells) - 1, 2):
                            key_text = re.sub(r'\s+', ' ', cells[i].get_text(strip=True))
                            val_text = cells[i+1].get_text(strip=True)
                            if key_text in mapping:
                                data[mapping[key_text]] = val_text
            return data
        else:
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
                '내용': '',
                '계약번호': '',
                '계약명': '',
                '계약업체명': '',
                '업체대표자명': '',
                '업체전화번호': '',
                '사업자등록번호': '',
                '업체주소': '',
                '계약(예정)일': '',
                '계약기간': '',
                '계약금액': '',
                '등록일': '',
                '분류': '',
                '수의계약 체결사유': ''
            }
            tables = soup.find_all("table", class_="contTbl")
            for table in tables:
                tbody = table.find("tbody")
                if not tbody:
                    continue
                rows = tbody.find_all("tr")
                for row in rows:
                    cells = row.find_all(["th", "td"])
                    if len(cells) < 2:
                        continue
                    for i in range(0, len(cells) - 1, 2):
                        key_text = cells[i].get_text(strip=True)
                        val_text = cells[i+1].get_text(strip=True)
                        if key_text == '파일첨부':
                            continue
                        if key_text in data:
                            data[key_text] = val_text
            return data
