import sys
import os
import re
import json
import requests
import pandas as pd
import openpyxl

from bs4 import BeautifulSoup
from datetime import datetime
from urllib.parse import urlparse, parse_qs

# PyQt 임포트
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QPushButton, QTextEdit, QRadioButton, QButtonGroup,
    QFileDialog, QCheckBox, QGroupBox, QGridLayout, QMessageBox, QComboBox, QSpinBox, QFormLayout, QDialog
)
from PyQt5.QtCore import Qt, QThread, QObject, pyqtSignal
from PyQt5.QtGui import QTextCursor

# =====================
# == 크롤링 함수 (원본) ==
# =====================

def get_soup_by_page(user_input_url, page_no):
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
        return None
    return BeautifulSoup(response.text, "html.parser")

def get_last_page_number(soup):
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

def parse_bid_table(soup, page_type_index=0):
    data_list = []
    if page_type_index == 0:
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
        if len(tds) < 7:
            continue

        if page_type_index == 0:
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

def crawl_all_pages(user_input_url, page_type_index=0, log_callback=None, max_items=50):
    def _log(msg):
        if log_callback:
            log_callback(msg)

    first_page_soup = get_soup_by_page(user_input_url, page_no=1)
    if not first_page_soup:
        _log("첫 페이지 로드 실패")
        return []

    last_page = get_last_page_number(first_page_soup)
    _log(f"확인된 마지막 페이지: {last_page}")

    all_data = []
    for page in range(1, last_page + 1):
        _log(f"{page}/{last_page} 페이지 처리 중...")
        page_soup = get_soup_by_page(user_input_url, page_no=page)
        if not page_soup:
            _log(f"{page} 페이지 로드 실패. 넘어갑니다.")
            continue

        page_data = parse_bid_table(page_soup, page_type_index=page_type_index)
        all_data.extend(page_data)
        if len(all_data) >= max_items:
            all_data = all_data[:max_items]
            break

    _log(f"총 {len(all_data)}개의 데이터 수집 완료")
    return all_data

def make_unique_filename(base_name="추출데이터"):
    folder_name = "추출데이터"
    os.makedirs(folder_name, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"{base_name}_{timestamp}.xlsx"
    full_path = os.path.join(folder_name, filename)

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

def save_to_excel(data_list, filename, page_type_index=0):
    wb = openpyxl.Workbook()
    ws = wb.active

    if page_type_index == 0:
        ws.title = "수의계약"
        headers = ["순번", "단지명", "계약업체", "계약명", "계약일", "계약금액", "계약기간", "상세정보링크"]
        ws.append(headers)

        for item in data_list:
            row = [
                item.get("순번", ""),
                item.get("단지명", ""),
                item.get("계약업체", ""),
                item.get("계약명", ""),
                item.get("계약일", ""),
                item.get("계약금액", ""),
                item.get("계약기간", ""),
                item.get("상세정보링크", "")
            ]
            ws.append(row)

    else:
        ws.title = "입찰공고"
        headers = ["순번", "종류", "낙찰방법", "입찰공고명", "입찰마감일", "상태", "단지명", "공고일", "상세정보링크"]
        ws.append(headers)

        for item in data_list:
            row = [
                item.get("순번", ""),
                item.get("종류", ""),
                item.get("낙찰방법", ""),
                item.get("입찰공고명", ""),
                item.get("입찰마감일", ""),
                item.get("상태", ""),
                item.get("단지명", ""),
                item.get("공고일", ""),
                item.get("상세정보링크", "")
            ]
            ws.append(row)

    wb.save(filename)

def crawl_detail_page(url: str, page_type_index=0) -> dict:
    response = requests.get(url)
    response.raise_for_status()
    soup = BeautifulSoup(response.text, "lxml")

    if page_type_index == 0:
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
                        key_text = cells[i].get_text(strip=True)
                        val_text = cells[i+1].get_text(strip=True)
                        key_text = re.sub(r'\s+', ' ', key_text)
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

def generate_output_filename():
    now_str = datetime.now().strftime("%Y%m%d_%H%M%S")
    base_filename = f"추출데이터_상세정보_{now_str}"
    ext = ".xlsx"
    output_filename = base_filename + ext

    counter = 1
    while os.path.exists(output_filename):
        output_filename = f"{base_filename}_{counter}{ext}"
        counter += 1
    return output_filename

def crawl_detail_info_from_excel(input_excel_path, selected_columns, log_callback=None, page_type_index=0):
    def _log(msg):
        if log_callback:
            log_callback(msg)

    if not os.path.exists(input_excel_path):
        _log(f"엑셀 파일이 존재하지 않습니다: {input_excel_path}")
        return None

    output_dir = "추출데이터_상세정보"
    os.makedirs(output_dir, exist_ok=True)

    output_filename = generate_output_filename()
    output_excel_path = os.path.join(output_dir, output_filename)

    df_input = pd.read_excel(input_excel_path)
    total_count = len(df_input)
    _log(f"총 {total_count} 건에 대해 상세정보 크롤링 시작...")

    results = []
    for idx, row in df_input.iterrows():
        detail_url = row.get('상세정보링크')
        if not detail_url:
            _log(f"[{idx+1}/{total_count}] 링크 없음 (건너뛰기)")
            integrated_data = row.to_dict()
            results.append(integrated_data)
            continue

        _log(f"[{idx+1}/{total_count}] 상세정보 크롤링 중: {detail_url}")
        max_retries = 3
        crawled_data = None
        for attempt in range(1, max_retries + 1):
            try:
                crawled_data = crawl_detail_page(detail_url, page_type_index=page_type_index)
                _log(f"  [성공] (시도 {attempt}/{max_retries})")
                break
            except Exception as e:
                _log(f"  [오류] (시도 {attempt}/{max_retries}): {e}")

        if crawled_data:
            integrated_data = row.to_dict()
            integrated_data.update(crawled_data)
            results.append(integrated_data)
        else:
            _log("  [실패] 3번 재시도 후 포기.")
            failed_data = {}
            for col in selected_columns:
                failed_data[col] = 'FAILED'
            integrated_data = row.to_dict()
            integrated_data.update(failed_data)
            results.append(integrated_data)

    df_result = pd.DataFrame(results)

    if page_type_index == 0:
        original_summary_cols = ["순번", "단지명", "계약업체", "계약명", "계약일", "계약금액", "계약기간", "상세정보링크"]
    else:
        original_summary_cols = ["순번", "종류", "낙찰방법", "입찰공고명", "입찰마감일", "상태", "단지명", "공고일", "상세정보링크"]
    final_cols = original_summary_cols + [col for col in selected_columns if col not in original_summary_cols]
    final_cols = [c for c in final_cols if c in df_result.columns]

    df_result = df_result[final_cols]
    df_result.to_excel(output_excel_path, index=False)
    _log(f"\n상세정보 크롤링 완료! 결과: {output_excel_path}")
    return output_excel_path

# ============================
# == QThread 활용한 Worker ==
# ============================
class CrawlerWorker(QObject):
    log_signal = pyqtSignal(str)          # 로그 문자열
    finished_signal = pyqtSignal(str)     # 최종 결과(파일 경로 등) 알림

    def __init__(self, mode, url_text, excel_path, selected_columns, extraction_count, page_type_index=0, parent=None):
        super().__init__(parent)
        self.mode = mode
        self.url_text = url_text
        self.excel_path = excel_path
        self.selected_columns = selected_columns
        self.page_type_index = page_type_index
        self.extraction_count = extraction_count

    def run(self):
        try:
            if self.mode == 1:
                self._run_summary_plus_detail()
            elif self.mode == 2:
                self._run_summary_only()
            else:
                self._run_detail_only()
        except Exception as e:
            self.finished_signal.emit(f"ERROR: {e}")

    def _log(self, msg):
        self.log_signal.emit(msg)

    def _make_auto_url(self):
        base = "https://www.k-apt.go.kr"
        if self.page_type_index == 0:
            return f"{base}/bid/privateContractList.do"
        elif self.page_type_index == 1:
            return f"{base}/bid/bidList.do?type=3"
        else:
            return f"{base}/bid/bidList.do"

    def _get_final_url(self):
        final_url = self.url_text.strip() if self.url_text.strip() else self._make_auto_url()
        if not self._check_url_page_match(final_url):
            raise ValueError("URL과 선택한 페이지 유형이 일치하지 않습니다!")
        return final_url

    def _check_url_page_match(self, url_text):
        parsed = urlparse(url_text)
        path = parsed.path
        qs = parse_qs(parsed.query)

        if self.page_type_index == 0:
            return path.endswith("/privateContractList.do")
        elif self.page_type_index == 1:
            if path.endswith("/bidList.do"):
                param_type = qs.get("type", [""])[0]
                return (param_type == "3")
            else:
                return False
        else:
            if path.endswith("/bidList.do"):
                param_type = qs.get("type", [""])[0]
                return (param_type != "3")
            else:
                return False

    def _run_summary_plus_detail(self):
        final_url = self._get_final_url()
        if not final_url:
            self._log("URL이 없습니다.")
            self.finished_signal.emit("ERROR: URL이 비어있음.")
            return

        self._log("[전체 페이지 + 상세정보] 크롤링을 시작합니다...")
        all_data = crawl_all_pages(final_url, page_type_index=self.page_type_index, log_callback=self._log, max_items=self.extraction_count)
        if not all_data:
            self._log("크롤링할 데이터가 없습니다.")
            self.finished_signal.emit("완료: 데이터 없음")
            return

        summary_filename = make_unique_filename()
        save_to_excel(all_data, summary_filename, page_type_index=self.page_type_index)
        self._log(f"전체 페이지 크롤링 완료. 파일 저장: {summary_filename}")

        detail_output_path = crawl_detail_info_from_excel(summary_filename, self.selected_columns, log_callback=self._log, page_type_index=self.page_type_index)
        if detail_output_path:
            self._log(f"상세 정보 크롤링 완료. 결과 파일: {detail_output_path}")

        self.finished_signal.emit(detail_output_path if detail_output_path else "상세 정보 없음")

    def _run_summary_only(self):
        final_url = self._get_final_url()
        if not final_url:
            self._log("URL이 없습니다.")
            self.finished_signal.emit("ERROR: URL이 비어있음.")
            return

        self._log("[전체 페이지만] 크롤링을 시작합니다...")
        all_data = crawl_all_pages(final_url, page_type_index=self.page_type_index, log_callback=self._log, max_items=self.extraction_count)
        if not all_data:
            self._log("크롤링할 데이터가 없습니다.")
            self.finished_signal.emit("완료: 데이터 없음")
            return

        summary_filename = make_unique_filename()
        save_to_excel(all_data, summary_filename, page_type_index=self.page_type_index)
        self._log(f"전체 페이지 크롤링 완료. 파일 저장: {summary_filename}")
        self.finished_signal.emit(summary_filename)

    def _run_detail_only(self):
        if not self.excel_path or not os.path.exists(self.excel_path):
            self._log(f"엑셀 파일이 존재하지 않습니다: {self.excel_path}")
            self.finished_signal.emit("ERROR: 엑셀 파일 경로 문제")
            return

        self._log("[기존 엑셀 -> 상세정보] 크롤링을 시작합니다...")
        detail_output_path = crawl_detail_info_from_excel(self.excel_path, self.selected_columns, log_callback=self._log, page_type_index=self.page_type_index)
        if detail_output_path:
            self._log(f"상세 정보 크롤링 완료. 결과 파일: {detail_output_path}")
            self.finished_signal.emit(detail_output_path)
        else:
            self.finished_signal.emit("상세정보 크롤링 실패 또는 데이터 없음")

# =====================
# == PyQt GUI 코드  ==
# =====================
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("K-APT 크롤러")
        self.setGeometry(100, 100, 900, 700)
        self.auto_exit = False  # 자동 종료 설정 (기본: 비활성)

        # 메뉴바 설정
        menubar = self.menuBar()
        help_menu = menubar.addMenu("도움말")
        help_action = help_menu.addAction("도움말")
        help_action.triggered.connect(self.show_help)

        settings_menu = menubar.addMenu("설정")
        settings_action = settings_menu.addAction("설정")
        settings_action.triggered.connect(self.show_settings_dialog)

        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        main_layout = QVBoxLayout()
        main_widget.setLayout(main_layout)

        # [크롤링 기본 설정] 그룹 (URL, 추출 갯수, 모드)
        crawl_setting_group = QGroupBox("크롤링 기본 설정")
        crawl_setting_layout = QFormLayout()
        crawl_setting_group.setLayout(crawl_setting_layout)
        self.url_edit = QLineEdit()
        self.url_edit.setPlaceholderText("크롤링할 URL을 입력하세요. (빈 칸이면 기본 URL 사용)")
        self.count_spin = QSpinBox()
        self.count_spin.setMinimum(1)
        self.count_spin.setMaximum(10000)
        self.count_spin.setValue(50)
        crawl_setting_layout.addRow("크롤링할 URL:", self.url_edit)
        crawl_setting_layout.addRow("추출 갯수:", self.count_spin)

        # [모드 선택] 그룹 (라디오버튼)
        mode_group = QGroupBox("크롤링 모드 선택")
        mode_layout = QHBoxLayout()
        mode_group.setLayout(mode_layout)
        self.radio_group = QButtonGroup(self)
        self.radio_summary_detail = QRadioButton("전체 페이지 + 상세정보 크롤링")
        self.radio_summary_only = QRadioButton("전체 페이지만 크롤링")
        self.radio_detail_only = QRadioButton("기존 엑셀 -> 상세정보만")
        self.radio_group.addButton(self.radio_summary_detail, 1)
        self.radio_group.addButton(self.radio_summary_only, 2)
        self.radio_group.addButton(self.radio_detail_only, 3)
        self.radio_summary_detail.setChecked(True)
        mode_layout.addWidget(self.radio_summary_detail)
        mode_layout.addWidget(self.radio_summary_only)
        mode_layout.addWidget(self.radio_detail_only)

        # [페이지 유형 선택] 그룹
        page_type_groupbox = QGroupBox("크롤링할 페이지 유형 선택")
        page_type_layout = QHBoxLayout()
        page_type_groupbox.setLayout(page_type_layout)
        self.page_type_combo = QComboBox()
        self.page_type_combo.addItems([
            "사업자 선정(수의계약) 결과 공개",  # index=0
            "사업자 선정(경쟁입찰) 결과 공개",  # index=1
            "전국 입찰공고"                # index=2
        ])
        self.page_type_combo.setCurrentIndex(0)
        page_type_layout.addWidget(self.page_type_combo)
        self.page_type_combo.currentIndexChanged.connect(self.update_detail_checkboxes)

        # [엑셀 파일 선택] 그룹
        file_group = QGroupBox("엑셀 파일 선택 (상세정보 크롤링 모드에서 사용)")
        file_layout = QHBoxLayout()
        file_group.setLayout(file_layout)
        self.file_label = QLabel("(선택된 엑셀 파일 없음)")
        self.file_button = QPushButton("엑셀 파일 선택")
        self.file_button.clicked.connect(self.select_excel_file)
        file_layout.addWidget(self.file_button)
        file_layout.addWidget(self.file_label)

        # [상세 컬럼 선택] 그룹
        self.detail_groupbox = QGroupBox("상세 컬럼 선택")
        self.detail_grid = QGridLayout()
        self.detail_groupbox.setLayout(self.detail_grid)
        self.checkboxes = []
        self.detail_columns = []
        self.update_detail_checkboxes()

        # [설정 관련] 버튼 영역 (설정 저장, 불러오기)
        settings_layout = QHBoxLayout()
        self.save_button = QPushButton("설정 저장")
        self.save_button.clicked.connect(self.save_favorites)
        self.load_button = QPushButton("설정 불러오기")
        self.load_button.clicked.connect(self.load_favorites)
        settings_layout.addWidget(self.save_button)
        settings_layout.addWidget(self.load_button)

        # 로그 출력창
        self.log_edit = QTextEdit()
        self.log_edit.setReadOnly(True)
        self.log_edit.setStyleSheet("background-color: #F0F0F0;")
        
        # 크롤링 시작 버튼 (맨 아래, 중앙 정렬, 큰 글씨)
        self.run_button = QPushButton("크롤링 시작")
        self.run_button.setStyleSheet("font-size: 18pt; font-weight: bold;")
        self.run_button.clicked.connect(self.on_run_clicked)

        # 최종 레이아웃 구성
        main_layout.addWidget(crawl_setting_group)
        main_layout.addWidget(mode_group)
        main_layout.addWidget(page_type_groupbox)
        main_layout.addWidget(file_group)
        main_layout.addWidget(self.detail_groupbox)
        main_layout.addLayout(settings_layout)
        main_layout.addWidget(self.log_edit)
        # 하단에 버튼 중앙 배치 (여백 추가)
        button_layout = QHBoxLayout()
        button_layout.addStretch()
        button_layout.addWidget(self.run_button)
        button_layout.addStretch()
        main_layout.addLayout(button_layout)

        self.thread = None
        self.worker = None
        self.selected_excel_path = ""

        # 모드에 따른 UI 상태 업데이트
        self.radio_summary_detail.toggled.connect(self.update_ui_state)
        self.radio_summary_only.toggled.connect(self.update_ui_state)
        self.radio_detail_only.toggled.connect(self.update_ui_state)
        self.update_ui_state()

        # 프로그램 시작 시 favorites/default.json 자동 로드 (로그 메시지 없이)
        default_settings_path = os.path.join("favorites", "default.json")
        if os.path.exists(default_settings_path):
            try:
                with open(default_settings_path, "r", encoding="utf-8") as f:
                    settings = json.load(f)
                self.apply_settings(settings)
            except Exception as e:
                self.log("기본 설정 불러오기 실패: " + str(e))

    def apply_settings(self, settings):
        self.url_edit.setText(settings.get("url", ""))
        self.count_spin.setValue(settings.get("extraction_count", 50))
        mode = settings.get("mode", 1)
        if mode == 1:
            self.radio_summary_detail.setChecked(True)
        elif mode == 2:
            self.radio_summary_only.setChecked(True)
        elif mode == 3:
            self.radio_detail_only.setChecked(True)
        self.page_type_combo.setCurrentIndex(settings.get("page_type_index", 0))
        self.selected_excel_path = settings.get("selected_excel_path", "")
        if self.selected_excel_path:
            self.file_label.setText(os.path.basename(self.selected_excel_path))
        else:
            self.file_label.setText("(선택된 엑셀 파일 없음)")
        self.update_detail_checkboxes()
        stored_cols = settings.get("selected_detail_columns", [])
        for cb in self.checkboxes:
            cb.setChecked(cb.text() in stored_cols)
        self.auto_exit = settings.get("auto_exit", False)

    def update_detail_checkboxes(self):
        page_type_index = self.page_type_combo.currentIndex()
        if page_type_index == 0:
            columns = [
                '주택관리업자', '아파트명', '관리사무소 주소', '전화번호', '팩스번호',
                '동수', '세대수', '계약번호', '계약명', '계약업체명', '업체대표자명',
                '업체전화번호', '사업자등록번호', '업체주소', '계약(예정)일',
                '계약금액', '계약기간', '등록일', '분류', '수의계약 체결사유'
            ]
            default_checked = {
                '아파트명', '전화번호', '동수', '세대수', '계약명',
                '계약업체명', '업체대표자명', '업체전화번호', '계약금액',
                '계약기간', '수의계약 체결사유'
            }
        else:
            columns = [
                '주택관리업자', '단지명', '관리사무소 주소', '전화번호', '팩스번호',
                '동수', '세대수', '입찰번호', '입찰방법', '입찰서 제출 마감일',
                '입찰제목', '긴급입찰여부', '입찰종류', '낙찰방법', '입찰분류',
                '신용평가등급확인서 제출여부', '현장설명',
                '관리(공사용역) 실적증명서 제출여부', '현장설명일시', '현장설명장소',
                '서류제출마감일', '입찰보증금', '지급조건', '내용'
            ]
            unchecked = {'내용', '현장설명장소', '신용평가등급확인서 제출여부', '긴급입찰여부', '현장설명일시'}

        for i in reversed(range(self.detail_grid.count())):
            widget = self.detail_grid.itemAt(i).widget()
            if widget is not None:
                widget.setParent(None)
        self.checkboxes = []
        self.detail_columns = columns

        for i, col in enumerate(columns):
            cb = QCheckBox(col)
            if self.page_type_combo.currentIndex() == 0:
                cb.setChecked(col in default_checked)
            else:
                cb.setChecked(col not in unchecked)
            self.checkboxes.append(cb)
            row_pos = i // 4
            col_pos = i % 4
            self.detail_grid.addWidget(cb, row_pos, col_pos)

    def select_excel_file(self):
        fname, _ = QFileDialog.getOpenFileName(self, "엑셀 파일 선택", "", "Excel Files (*.xlsx *.xls)")
        if fname:
            self.file_label.setText(os.path.basename(fname))
            self.selected_excel_path = fname
        else:
            self.file_label.setText("(선택된 엑셀 파일 없음)")
            self.selected_excel_path = ""

    def update_ui_state(self):
        mode_id = self.radio_group.checkedId()
        if mode_id == 1:  # 전체 + 상세
            self.url_edit.setEnabled(True)
            self.file_button.setEnabled(False)
            for cb in self.checkboxes:
                cb.setEnabled(True)
        elif mode_id == 2:  # 전체만
            self.url_edit.setEnabled(True)
            self.file_button.setEnabled(False)
            for cb in self.checkboxes:
                cb.setEnabled(False)
        else:  # 상세만
            self.url_edit.setEnabled(False)
            self.file_button.setEnabled(True)
            for cb in self.checkboxes:
                cb.setEnabled(True)

    def log(self, msg):
        self.log_edit.append(msg)
        self.log_edit.moveCursor(QTextCursor.End)
        QApplication.processEvents()

    def on_run_clicked(self):
        mode_id = self.radio_group.checkedId()
        url_text = self.url_edit.text().strip()
        excel_path = self.selected_excel_path
        selected_cols = [cb.text() for cb in self.checkboxes if cb.isChecked()]
        page_type_index = self.page_type_combo.currentIndex()
        extraction_count = self.count_spin.value()

        if hasattr(self, 'thread') and self.thread is not None:
            if self.thread.isRunning():
                QMessageBox.warning(self, "안내", "이미 크롤링 작업이 진행 중입니다.")
                return

        self.thread = QThread(self)
        self.worker = CrawlerWorker(mode_id, url_text, excel_path, selected_cols, extraction_count, page_type_index=page_type_index)
        self.worker.moveToThread(self.thread)

        self.thread.started.connect(self.worker.run)
        self.worker.log_signal.connect(self.log)
        self.worker.finished_signal.connect(self.on_crawl_finished)
        self.worker.finished_signal.connect(self.thread.quit)

        self.thread.start()

    def on_crawl_finished(self, result):
        if result.startswith("ERROR:"):
            QMessageBox.warning(self, "에러 발생", result)
        else:
            if self.auto_exit:
                QMessageBox.information(self, "작업 완료", f"크롤링이 종료되었습니다.\n결과: {result}\n버튼을 누르면 종료됩니다.")
                QApplication.quit()
            else:
                QMessageBox.information(self, "작업 완료", f"크롤링이 종료되었습니다.\n결과: {result}")

    def save_favorites(self):
        settings = {
            "url": self.url_edit.text(),
            "extraction_count": self.count_spin.value(),
            "mode": self.radio_group.checkedId(),
            "page_type_index": self.page_type_combo.currentIndex(),
            "selected_excel_path": self.selected_excel_path,
            "selected_detail_columns": [cb.text() for cb in self.checkboxes if cb.isChecked()],
            "auto_exit": self.auto_exit
        }
        favorites_folder = "favorites"
        os.makedirs(favorites_folder, exist_ok=True)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"favorite_{timestamp}.json"
        full_path = os.path.join(favorites_folder, filename)
        try:
            with open(full_path, "w", encoding="utf-8") as f:
                json.dump(settings, f, ensure_ascii=False, indent=4)
            # 기본 설정(default.json)에도 저장
            default_path = os.path.join(favorites_folder, "default.json")
            with open(default_path, "w", encoding="utf-8") as f:
                json.dump(settings, f, ensure_ascii=False, indent=4)
            QMessageBox.information(self, "설정 저장", f"설정이 저장되었습니다 ({full_path}).")
        except Exception as e:
            QMessageBox.warning(self, "설정 저장 실패", str(e))

    def load_favorites(self):
        fname, _ = QFileDialog.getOpenFileName(self, "설정 파일 선택", "favorites", "JSON Files (*.json)")
        if not fname:
            QMessageBox.warning(self, "설정 불러오기", "선택된 설정 파일이 없습니다.")
            return
        try:
            with open(fname, "r", encoding="utf-8") as f:
                settings = json.load(f)
        except Exception as e:
            QMessageBox.warning(self, "설정 불러오기 실패", str(e))
            return

        self.apply_settings(settings)
        QMessageBox.information(self, "설정 불러오기", "설정이 불러와졌습니다.")

    def show_settings_dialog(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("설정")
        layout = QVBoxLayout(dialog)
        auto_exit_checkbox = QCheckBox("크롤링 완료 후 자동 종료")
        auto_exit_checkbox.setChecked(self.auto_exit)
        layout.addWidget(auto_exit_checkbox)
        btn_box = QHBoxLayout()
        ok_button = QPushButton("확인")
        cancel_button = QPushButton("취소")
        btn_box.addWidget(ok_button)
        btn_box.addWidget(cancel_button)
        layout.addLayout(btn_box)

        ok_button.clicked.connect(lambda: dialog.accept())
        cancel_button.clicked.connect(lambda: dialog.reject())
        result = dialog.exec_()
        if result == QDialog.Accepted:
            self.auto_exit = auto_exit_checkbox.isChecked()
            QMessageBox.information(self, "설정", f"설정이 저장되었습니다. 자동 종료: {'활성화' if self.auto_exit else '비활성화'}")

    def show_help(self):
        help_text = (
            "=== K-APT 크롤러 사용법 ===\n\n"
            "[1] 크롤링 기본 설정\n"
            "  - 크롤링할 URL: 크롤링할 웹페이지 주소를 입력합니다 (빈 칸일 경우 기본 URL 사용).\n"
            "  - 추출 갯수: 한 번에 수집할 데이터 건수를 지정합니다 (예: 50건).\n\n"
            "[2] 크롤링 모드 선택\n"
            "  - 전체 페이지 + 상세정보 크롤링: 목록 데이터와 함께 각 항목의 상세정보도 함께 크롤링합니다.\n"
            "  - 전체 페이지만 크롤링: 목록 데이터만 수집합니다.\n"
            "  - 기존 엑셀 -> 상세정보만: 기존에 저장된 엑셀 파일의 '상세정보링크'를 이용해 상세 정보를 추가로 크롤링합니다.\n\n"
            "[3] 페이지 유형 선택\n"
            "  - 사업자 선정(수의계약) 결과 공개\n"
            "  - 사업자 선정(경쟁입찰) 결과 공개\n"
            "  - 전국 입찰공고\n\n"
            "[4] 엑셀 파일 선택 (상세정보 크롤링 모드 전용)\n"
            "  - 상세정보 크롤링 모드일 경우, 기존 엑셀 파일을 선택하여 상세 데이터를 추가로 수집합니다.\n\n"
            "[5] 상세 컬럼 선택\n"
            "  - 크롤링 시 추가로 추출할 상세 정보를 선택할 수 있습니다.\n\n"
            "[6] 설정 저장 및 불러오기\n"
            "  - 설정 저장: 현재의 URL, 추출 갯수, 모드, 페이지 유형, 선택한 상세 컬럼, 자동 종료 설정 등을 'favorites' 폴더 내에 저장합니다.\n"
            "    마지막 상태는 favorites/default.json에도 저장됩니다.\n"
            "  - 설정 불러오기: 파일 다이얼로그를 통해 저장된 설정 파일을 선택하여 설정을 불러옵니다.\n\n"
            "[7] 크롤링 시작\n"
            "  - 모든 설정 후 하단의 '크롤링 시작' 버튼을 클릭하면 작업이 실행되며, 진행 상황은 로그 창에 표시됩니다.\n\n"
            "※ 주의사항\n"
            "  - URL과 선택한 페이지 유형이 일치해야 정상적으로 크롤링됩니다.\n"
            "  - 네트워크 문제나 페이지 구조 변경 시 일부 데이터 수집에 실패할 수 있습니다.\n\n"
            "※ CMD 사용법\n"
            "  - 커맨드라인에서 'app.exe help'를 입력하면 콘솔에 도움말이 출력됩니다.\n"
        )
        QMessageBox.information(self, "도움말", help_text)

def main():
    if len(sys.argv) > 1 and sys.argv[1].lower() == "help":
        print("K-APT 크롤러 도움말:")
        print("사용법:")
        print("  app.exe                   : GUI 모드로 프로그램 실행")
        print("  app.exe help              : 도움말 출력")
        print("")
        print("GUI 사용:")
        print("  - 크롤링할 URL 입력")
        print("  - 추출 갯수 설정")
        print("  - 라디오 버튼으로 크롤링 모드 선택 (전체 페이지 + 상세정보, 전체 페이지만, 기존 엑셀 -> 상세정보만)")
        print("  - 크롤링할 페이지 유형 선택 (수의계약, 경쟁입찰, 전국 입찰공고)")
        print("  - 상세 컬럼 선택 (필요한 컬럼 체크)")
        print("  - 설정 저장/불러오기 버튼으로 즐겨찾기 기능 사용 가능 (마지막 상태는 default.json에 저장)")
        print("  - 하단의 '크롤링 시작' 버튼 클릭으로 실행 (버튼은 크게 표시됨)")
        sys.exit(0)
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()
