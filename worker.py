import os, json
from crawler import SummaryCrawler, DetailCrawler
from excel_handler import make_unique_filename, save_to_excel, crawl_detail_info_from_excel
from urllib.parse import urlparse, parse_qs
from PyQt5.QtCore import QObject, pyqtSignal

def read_json_with_encoding(file_path):
    """다양한 인코딩으로 JSON 파일을 읽는 함수"""
    encodings = ['utf-8', 'euc-kr']
    
    for encoding in encodings:
        try:
            with open(file_path, "r", encoding=encoding) as f:
                return json.load(f)
        except UnicodeDecodeError:
            continue
        except json.JSONDecodeError as e:
            print(f"{encoding} 인코딩으로 읽었으나 JSON 파싱 실패: {e}")
            continue
    
    raise ValueError(f"파일을 읽을 수 없습니다. 지원되는 인코딩: {', '.join(encodings)}")

class CrawlerWorker:
    def __init__(self, mode, url_text, excel_path, selected_columns, extraction_count, page_type_index=0, log_callback=None):
        self.mode = mode
        self.url_text = url_text
        self.excel_path = excel_path
        self.selected_columns = selected_columns
        self.page_type_index = page_type_index
        self.extraction_count = extraction_count
        self.log_callback = log_callback

    def _log(self, msg):
        if self.log_callback:
            self.log_callback(msg)

    def _make_auto_url(self):
        base = "https://www.k-apt.go.kr"
        if self.page_type_index == 0:
            return f"{base}/bid/privateContractList.do"
        elif self.page_type_index == 1:
            return f"{base}/bid/bidList.do?type=3"
        else:
            return f"{base}/bid/bidList.do"

    def _get_final_url(self):
        final_url = self.url_text.strip()
        # 사용자가 입력한 URL이 있고, 선택된 페이지 유형과 일치하면 그대로 사용
        if final_url and self._check_url_page_match(final_url):
            return final_url
        else:
            # URL이 없거나, 페이지 유형이 맞지 않는 경우 기본 URL 사용
            self._log("입력된 URL이 선택된 페이지 유형과 일치하지 않습니다. 기본 URL로 대체합니다.")
            return self._make_auto_url()


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

    def run(self):
        results = {}
        json_files = [os.path.join(self.folder_path, f) for f in os.listdir(self.folder_path) if f.endswith('.json')]
        if not json_files:
            self._log("선택한 폴더에 JSON 파일이 없습니다.")
            self.finished_signal.emit("실행된 크롤링 없음")
            return
        for json_file in json_files:
            try:
                settings = read_json_with_encoding(json_file)
            except Exception as e:
                self._log(f"파일 {json_file} 읽기 실패: {e}")
                continue
            self._log(f"설정 파일 처리 중: {os.path.basename(json_file)}")
            mode = settings.get("mode", 1)
            url_text = settings.get("url", "")
            excel_path = settings.get("selected_excel_path", "")
            selected_columns = settings.get("selected_detail_columns", [])
            extraction_count = settings.get("extraction_count", 50)
            page_type_index = settings.get("page_type_index", 0)
            worker = CrawlerWorker(mode, url_text, excel_path, selected_columns, extraction_count, page_type_index, log_callback=self._log)
            try:
                result = worker.run()
                self._log(f"크롤링 완료 ({os.path.basename(json_file)}): 결과 파일 -> {result}")
                results[os.path.basename(json_file)] = result
            except Exception as e:
                self._log(f"크롤링 실패 ({os.path.basename(json_file)}): {e}")
        self.finished_signal.emit("모든 크롤링 작업 완료")

    def _run_summary_plus_detail(self):
        final_url = self._get_final_url()
        if not final_url:
            self._log("URL이 없습니다.")
            raise ValueError("URL이 비어있음.")
        self._log("[전체 페이지 + 상세정보] 크롤링을 시작합니다...")
        summary_crawler = SummaryCrawler(final_url, page_type_index=self.page_type_index)
        all_data = summary_crawler.crawl_all_pages(final_url, log_callback=self._log, max_items=self.extraction_count)
        if not all_data:
            self._log("크롤링할 데이터가 없습니다.")
            return "완료: 데이터 없음"
        summary_filename = make_unique_filename()
        save_to_excel(all_data, summary_filename, page_type_index=self.page_type_index)
        self._log(f"전체 페이지 크롤링 완료. 파일 저장: {summary_filename}")
        detail_crawler = DetailCrawler(page_type_index=self.page_type_index)
        detail_output_path = crawl_detail_info_from_excel(summary_filename, self.selected_columns, detail_crawler, log_callback=self._log, page_type_index=self.page_type_index)
        if detail_output_path:
            self._log(f"상세 정보 크롤링 완료. 결과 파일: {detail_output_path}")
        return detail_output_path if detail_output_path else "상세 정보 없음"

    def _run_summary_only(self):
        final_url = self._get_final_url()
        if not final_url:
            self._log("URL이 없습니다.")
            raise ValueError("URL이 비어있음.")
        self._log("[전체 페이지만] 크롤링을 시작합니다...")
        summary_crawler = SummaryCrawler(final_url, page_type_index=self.page_type_index)
        all_data = summary_crawler.crawl_all_pages(final_url, log_callback=self._log, max_items=self.extraction_count)
        if not all_data:
            self._log("크롤링할 데이터가 없습니다.")
            return "완료: 데이터 없음"
        summary_filename = make_unique_filename()
        save_to_excel(all_data, summary_filename, page_type_index=self.page_type_index)
        self._log(f"전체 페이지 크롤링 완료. 파일 저장: {summary_filename}")
        return summary_filename

    def _run_detail_only(self):
        if not self.excel_path or not os.path.exists(self.excel_path):
            self._log(f"엑셀 파일이 존재하지 않습니다: {self.excel_path}")
            raise ValueError("엑셀 파일 경로 문제")
        self._log("[기존 엑셀 -> 상세정보] 크롤링을 시작합니다...")
        detail_crawler = DetailCrawler(page_type_index=self.page_type_index)
        detail_output_path = crawl_detail_info_from_excel(self.excel_path, self.selected_columns, detail_crawler, log_callback=self._log, page_type_index=self.page_type_index)
        if detail_output_path:
            self._log(f"상세 정보 크롤링 완료. 결과 파일: {detail_output_path}")
            return detail_output_path
        else:
            return "상세정보 크롤링 실패 또는 데이터 없음"

# --- 새롭게 추가된 MultiCrawlerWorker 클래스 ---
class MultiCrawlerWorker(QObject):
    log_signal = pyqtSignal(str)
    finished_signal = pyqtSignal(str)

    def __init__(self, folder_path):
        super().__init__()
        self.folder_path = folder_path

    def _log(self, msg):
        self.log_signal.emit(msg)

    def run(self):
        import json, os
        results = {}
        json_files = [os.path.join(self.folder_path, f) for f in os.listdir(self.folder_path) if f.endswith('.json')]
        if not json_files:
            self._log("선택한 폴더에 JSON 파일이 없습니다.")
            self.finished_signal.emit("실행된 크롤링 없음")
            return
        for json_file in json_files:
            try:
                with open(json_file, 'r', encoding='utf-8') as f:
                    settings = json.load(f)
            except Exception as e:
                self._log(f"파일 {json_file} 읽기 실패: {e}")
                continue
            self._log(f"설정 파일 처리 중: {os.path.basename(json_file)}")
            mode = settings.get("mode", 1)
            url_text = settings.get("url", "")
            excel_path = settings.get("selected_excel_path", "")
            selected_columns = settings.get("selected_detail_columns", [])
            extraction_count = settings.get("extraction_count", 50)
            page_type_index = settings.get("page_type_index", 0)
            worker = CrawlerWorker(mode, url_text, excel_path, selected_columns, extraction_count, page_type_index, log_callback=self._log)
            try:
                result = worker.run()
                self._log(f"크롤링 완료 ({os.path.basename(json_file)}): 결과 파일 -> {result}")
                results[os.path.basename(json_file)] = result
            except Exception as e:
                self._log(f"크롤링 실패 ({os.path.basename(json_file)}): {e}")
        self.finished_signal.emit("모든 크롤링 작업 완료")
