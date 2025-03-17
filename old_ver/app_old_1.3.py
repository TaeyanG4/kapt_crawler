import sys
import os
import re
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
    QFileDialog, QCheckBox, QGroupBox, QGridLayout, QMessageBox,
    QComboBox
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

def parse_bid_table(soup):
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

def crawl_all_pages(user_input_url, log_callback=None):
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

        page_data = parse_bid_table(page_soup)
        all_data.extend(page_data)

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

def save_to_excel(data_list, filename):
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

def crawl_detail_page(url: str) -> dict:
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

    # (1) 첫 번째 테이블 (단지정보)
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

    # (2) 두 번째 테이블 (입찰정보)
    table2_candidates = soup.find_all("table", class_="contTbl")
    table2 = None
    for t in table2_candidates:
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
                    if key == '파일첨부':
                        continue
                    if key in data:
                        data[key] = val

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

def crawl_detail_info_from_excel(input_excel_path, selected_columns, log_callback=None):
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
                crawled_data = crawl_detail_page(detail_url)
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

    original_summary_cols = [
        "순번", "종류", "낙찰방법", "입찰공고명",
        "입찰마감일", "상태", "단지명", "공고일", "상세정보링크"
    ]
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

    def __init__(self, mode, url_text, excel_path, selected_columns, page_type_index=0, parent=None):
        super().__init__(parent)
        self.mode = mode
        self.url_text = url_text
        self.excel_path = excel_path
        self.selected_columns = selected_columns
        self.page_type_index = page_type_index

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
        # page_type_index에 따라 URL 달리 반환
        # 0 => 수의계약 결과
        # 1 => 경쟁입찰 결과(type=3)
        # 2 => 전국 입찰공고
        if self.page_type_index == 0:
            return f"{base}/bid/privateContractList.do"
        elif self.page_type_index == 1:
            return f"{base}/bid/bidList.do?type=3"
        else:
            return f"{base}/bid/bidList.do"

    def _get_final_url(self):
        """
        url_text가 입력되어 있으면 그것을 사용,
        없으면 _make_auto_url로 생성한 url 사용.
        
        이후, page_type_index와 final_url이 '잘 맞는지' 검사.
        틀리면 예외 발생.
        """
        final_url = self.url_text.strip() if self.url_text.strip() else self._make_auto_url()

        if not self._check_url_page_match(final_url):
            raise ValueError("URL과 선택한 페이지 유형이 일치하지 않습니다!")
        return final_url

    def _check_url_page_match(self, url_text):
        """
        page_type_index와 url_text가 적절히 매칭되는지 검사.
        - 0 => /bid/privateContractList.do
        - 1 => /bid/bidList.do?type=3
        - 2 => /bid/bidList.do (type != 3)
        """
        parsed = urlparse(url_text)
        path = parsed.path
        qs = parse_qs(parsed.query)

        if self.page_type_index == 0:
            # 수의계약 => /privateContractList.do
            return path.endswith("/privateContractList.do")
        elif self.page_type_index == 1:
            # 경쟁입찰 => /bidList.do?type=3
            if path.endswith("/bidList.do"):
                param_type = qs.get("type", [""])[0]
                return (param_type == "3")
            else:
                return False
        else:
            # 전국입찰 => /bidList.do (type != 3)
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
        all_data = crawl_all_pages(final_url, log_callback=self._log)
        if not all_data:
            self._log("크롤링할 데이터가 없습니다.")
            self.finished_signal.emit("완료: 데이터 없음")
            return

        summary_filename = make_unique_filename()
        save_to_excel(all_data, summary_filename)
        self._log(f"전체 페이지 크롤링 완료. 파일 저장: {summary_filename}")

        detail_output_path = crawl_detail_info_from_excel(summary_filename, self.selected_columns, log_callback=self._log)
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
        all_data = crawl_all_pages(final_url, log_callback=self._log)
        if not all_data:
            self._log("크롤링할 데이터가 없습니다.")
            self.finished_signal.emit("완료: 데이터 없음")
            return

        summary_filename = make_unique_filename()
        save_to_excel(all_data, summary_filename)
        self._log(f"전체 페이지 크롤링 완료. 파일 저장: {summary_filename}")
        self.finished_signal.emit(summary_filename)

    def _run_detail_only(self):
        if not self.excel_path or not os.path.exists(self.excel_path):
            self._log(f"엑셀 파일이 존재하지 않습니다: {self.excel_path}")
            self.finished_signal.emit("ERROR: 엑셀 파일 경로 문제")
            return

        self._log("[전체 페이지 엑셀로부터 상세정보] 크롤링을 시작합니다...")
        detail_output_path = crawl_detail_info_from_excel(self.excel_path, self.selected_columns, log_callback=self._log)
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

        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        main_layout = QVBoxLayout()
        main_widget.setLayout(main_layout)

        # 1) URL 입력
        url_layout = QHBoxLayout()
        url_label = QLabel("크롤링할 URL:")
        self.url_edit = QLineEdit()
        url_layout.addWidget(url_label)
        url_layout.addWidget(self.url_edit)

        # 라디오버튼 그룹
        self.radio_group = QButtonGroup(self)

        self.radio_summary_detail = QRadioButton("전체 페이지 + 상세정보 크롤링")
        self.radio_summary_only = QRadioButton("전체 페이지만 크롤링")
        self.radio_detail_only = QRadioButton("기존 엑셀 -> 상세정보만")

        self.radio_group.addButton(self.radio_summary_detail, 1)
        self.radio_group.addButton(self.radio_summary_only, 2)
        self.radio_group.addButton(self.radio_detail_only, 3)

        # 기본 선택: "전체 페이지 + 상세정보 크롤링"
        self.radio_summary_detail.setChecked(True)

        radio_layout = QHBoxLayout()
        radio_layout.addWidget(self.radio_summary_detail)
        radio_layout.addWidget(self.radio_summary_only)
        radio_layout.addWidget(self.radio_detail_only)

        # 유형 선택 콤보박스
        page_type_groupbox = QGroupBox("크롤링할 페이지 유형 선택")
        page_type_layout = QHBoxLayout()
        page_type_groupbox.setLayout(page_type_layout)

        # 순서 변경 및 기본 선택 0 => 수의계약
        self.page_type_combo = QComboBox()
        self.page_type_combo.addItems([
            "사업자 선정(수의계약) 결과 공개",  # index=0
            "사업자 선정(경쟁입찰) 결과 공개",  # index=1
            "전국 입찰공고"                   # index=2
        ])
        # 기본 선택은 index=0 (수의계약)
        self.page_type_combo.setCurrentIndex(0)

        page_type_layout.addWidget(self.page_type_combo)

        # 엑셀 파일 선택
        file_layout = QHBoxLayout()
        self.file_label = QLabel("(선택된 엑셀 파일 없음)")
        self.file_button = QPushButton("엑셀 파일 선택")
        self.file_button.clicked.connect(self.select_excel_file)
        file_layout.addWidget(self.file_button)
        file_layout.addWidget(self.file_label)

        # 상세 컬럼 체크박스
        detail_groupbox = QGroupBox("상세 컬럼 선택")
        detail_grid = QGridLayout()
        detail_groupbox.setLayout(detail_grid)

        self.detail_columns = [
            '주택관리업자', '단지명', '관리사무소 주소', '전화번호', '팩스번호',
            '동수', '세대수', '입찰번호', '입찰방법', '입찰서 제출 마감일',
            '입찰제목', '긴급입찰여부', '입찰종류', '낙찰방법', '입찰분류',
            '신용평가등급확인서 제출여부', '현장설명',
            '관리(공사용역) 실적증명서 제출여부', '현장설명일시', '현장설명장소',
            '서류제출마감일', '입찰보증금', '지급조건', '내용'
        ]

        # 기본 체크 해제할 컬럼들
        unchecked_columns = {
            '내용', '현장설명장소', '신용평가등급확인서 제출여부', '긴급입찰여부', '현장설명일시'
        }

        self.checkboxes = []
        for i, col in enumerate(self.detail_columns):
            cb = QCheckBox(col)
            # 지정된 컬럼이면 체크 해제, 아니면 체크
            if col in unchecked_columns:
                cb.setChecked(False)
            else:
                cb.setChecked(True)
            self.checkboxes.append(cb)
            row_pos = i // 4
            col_pos = i % 4
            detail_grid.addWidget(cb, row_pos, col_pos)

        self.radio_summary_detail.toggled.connect(self.update_ui_state)
        self.radio_summary_only.toggled.connect(self.update_ui_state)
        self.radio_detail_only.toggled.connect(self.update_ui_state)
        self.update_ui_state()

        self.log_edit = QTextEdit()
        self.log_edit.setReadOnly(True)
        self.log_edit.setStyleSheet("background-color: #F0F0F0;")

        self.run_button = QPushButton("크롤링 시작")
        self.run_button.clicked.connect(self.on_run_clicked)

        main_layout.addLayout(url_layout)
        main_layout.addLayout(radio_layout)
        main_layout.addWidget(page_type_groupbox)
        main_layout.addLayout(file_layout)
        main_layout.addWidget(detail_groupbox)
        main_layout.addWidget(self.run_button)
        main_layout.addWidget(self.log_edit)

        self.thread = None
        self.worker = None
        self.selected_excel_path = ""

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

        if hasattr(self, 'thread') and self.thread is not None:
            if self.thread.isRunning():
                QMessageBox.warning(self, "안내", "이미 크롤링 작업이 진행 중입니다.")
                return

        self.thread = QThread(self)
        self.worker = CrawlerWorker(mode_id, url_text, excel_path, selected_cols, page_type_index=page_type_index)
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
            QMessageBox.information(self, "작업 완료", f"크롤링이 종료되었습니다.\n결과: {result}")


def main():
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()
