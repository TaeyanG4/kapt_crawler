import os, json, sys
from datetime import datetime
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QPushButton, QTextEdit, QRadioButton, QButtonGroup,
    QFileDialog, QCheckBox, QGroupBox, QGridLayout, QMessageBox, QComboBox, QSpinBox, QFormLayout, QDialog
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QObject
from PyQt5.QtGui import QTextCursor
from worker import CrawlerWorker, MultiCrawlerWorker
from worker import CrawlerWorker

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


class WorkerWrapper(QObject):
    log_signal = pyqtSignal(str)
    finished_signal = pyqtSignal(str)
    
    def __init__(self, mode, url_text, excel_path, selected_columns, extraction_count, page_type_index=0):
        super().__init__()
        self.worker = CrawlerWorker(mode, url_text, excel_path, selected_columns, extraction_count, page_type_index)
        self.worker.log_callback = self.log_signal.emit

    def run(self):
        try:
            result = self.worker.run()
            self.finished_signal.emit(result)
        except Exception as e:
            self.finished_signal.emit(f"ERROR: {e}")

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("K-APT 크롤러")
        self.setGeometry(100, 100, 900, 700)
        self.auto_exit = False

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

        # 크롤링 기본 설정 그룹
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

        # 모드 선택 그룹
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

        # 페이지 유형 선택 그룹
        page_type_groupbox = QGroupBox("크롤링할 페이지 유형 선택")
        page_type_layout = QHBoxLayout()
        page_type_groupbox.setLayout(page_type_layout)
        self.page_type_combo = QComboBox()
        self.page_type_combo.addItems([
            "사업자 선정(수의계약) 결과 공개",
            "사업자 선정(경쟁입찰) 결과 공개",
            "전국 입찰공고"
        ])
        self.page_type_combo.setCurrentIndex(0)
        page_type_layout.addWidget(self.page_type_combo)
        self.page_type_combo.currentIndexChanged.connect(self.update_detail_checkboxes)

        # 엑셀 파일 선택 그룹
        file_group = QGroupBox("엑셀 파일 선택 (상세정보 크롤링 모드에서 사용)")
        file_layout = QHBoxLayout()
        file_group.setLayout(file_layout)
        self.file_label = QLabel("(선택된 엑셀 파일 없음)")
        self.file_button = QPushButton("엑셀 파일 선택")
        self.file_button.clicked.connect(self.select_excel_file)
        file_layout.addWidget(self.file_button)
        file_layout.addWidget(self.file_label)

        # 상세 컬럼 선택 그룹
        self.detail_groupbox = QGroupBox("상세 컬럼 선택")
        self.detail_grid = QGridLayout()
        self.detail_groupbox.setLayout(self.detail_grid)
        self.checkboxes = []
        self.detail_columns = []
        self.update_detail_checkboxes()

        # 설정 저장/불러오기 버튼
        settings_layout = QHBoxLayout()
        self.save_button = QPushButton("설정 저장")
        self.save_button.clicked.connect(self.save_favorites)
        self.load_button = QPushButton("설정 불러오기")
        self.load_button.clicked.connect(self.load_favorites)
        settings_layout.addWidget(self.save_button)
        settings_layout.addWidget(self.load_button)
        
        # 새롭게 추가된 "폴더 설정 크롤링 실행" 버튼
        self.folder_crawl_button = QPushButton("폴더 설정 크롤링 실행")
        self.folder_crawl_button.clicked.connect(self.run_folder_crawling)

        # 로그 출력창
        self.log_edit = QTextEdit()
        self.log_edit.setReadOnly(True)
        self.log_edit.setStyleSheet("background-color: #F0F0F0;")

        # 크롤링 시작 버튼
        self.run_button = QPushButton("크롤링 시작")
        self.run_button.setStyleSheet("font-size: 18pt; font-weight: bold;")
        self.run_button.clicked.connect(self.on_run_clicked)

        # 레이아웃 구성
        main_layout.addWidget(crawl_setting_group)
        main_layout.addWidget(mode_group)
        main_layout.addWidget(page_type_groupbox)
        main_layout.addWidget(file_group)
        main_layout.addWidget(self.detail_groupbox)
        main_layout.addLayout(settings_layout)
        # 폴더 크롤링 버튼 추가
        main_layout.addWidget(self.folder_crawl_button)
        main_layout.addWidget(self.log_edit)
        button_layout = QHBoxLayout()
        button_layout.addStretch()
        button_layout.addWidget(self.run_button)
        button_layout.addStretch()
        main_layout.addLayout(button_layout)

        self.thread = None
        self.worker_wrapper = None
        self.selected_excel_path = ""
        self.radio_summary_detail.toggled.connect(self.update_ui_state)
        self.radio_summary_only.toggled.connect(self.update_ui_state)
        self.radio_detail_only.toggled.connect(self.update_ui_state)
        self.update_ui_state()

        default_settings_path = os.path.join("favorites", "default.json")
        if os.path.exists(default_settings_path):
            try:
                settings = read_json_with_encoding(default_settings_path)
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

        # 기존 위젯 제거
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
        if mode_id == 1:
            self.url_edit.setEnabled(True)
            self.file_button.setEnabled(False)
            for cb in self.checkboxes:
                cb.setEnabled(True)
        elif mode_id == 2:
            self.url_edit.setEnabled(True)
            self.file_button.setEnabled(False)
            for cb in self.checkboxes:
                cb.setEnabled(False)
        else:
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

        if self.thread is not None and self.thread.isRunning():
            QMessageBox.warning(self, "안내", "이미 크롤링 작업이 진행 중입니다.")
            return

        self.thread = QThread(self)
        self.worker_wrapper = WorkerWrapper(mode_id, url_text, excel_path, selected_cols, extraction_count, page_type_index)
        self.worker_wrapper.moveToThread(self.thread)
        self.thread.started.connect(self.worker_wrapper.run)
        self.worker_wrapper.log_signal.connect(self.log)
        self.worker_wrapper.finished_signal.connect(self.on_crawl_finished)
        self.worker_wrapper.finished_signal.connect(self.thread.quit)
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
            settings = read_json_with_encoding(fname)
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
            "  - 크롤링할 URL: 크롤링할 웹페이지 주소를 입력합니다 (빈 칸이면 기본 URL 사용).\n"
            "  - 추출 갯수: 한 번에 수집할 데이터 건수를 지정합니다.\n\n"
            "[2] 크롤링 모드 선택\n"
            "  - 전체 페이지 + 상세정보 크롤링: 목록 데이터와 각 항목의 상세정보를 함께 수집합니다.\n"
            "  - 전체 페이지만 크롤링: 목록 데이터만 수집합니다.\n"
            "  - 기존 엑셀 -> 상세정보만: 기존 엑셀 파일의 '상세정보링크'를 통해 추가 상세 정보를 수집합니다.\n\n"
            "[3] 페이지 유형 선택\n"
            "  - 사업자 선정(수의계약) 결과 공개\n"
            "  - 사업자 선정(경쟁입찰) 결과 공개\n"
            "  - 전국 입찰공고\n\n"
            "[4] 엑셀 파일 선택\n"
            "  - 상세정보 크롤링 모드일 경우, 기존 엑셀 파일을 선택하여 상세 데이터를 추가로 수집합니다.\n\n"
            "[5] 상세 컬럼 선택\n"
            "  - 크롤링 시 추가로 추출할 상세 정보를 선택할 수 있습니다.\n\n"
            "[6] 설정 저장 및 불러오기\n"
            "  - 현재 설정(URL, 추출 건수, 모드, 페이지 유형, 선택한 상세 컬럼, 자동 종료)을 favorites 폴더에 저장하고, 불러올 수 있습니다.\n\n"
            "[7] 크롤링 시작\n"
            "  - 모든 설정 후 하단의 '크롤링 시작' 버튼을 클릭하면 작업이 실행되며, 진행 상황은 로그 창에 표시됩니다.\n\n"
            "※ URL과 선택한 페이지 유형이 일치해야 정상 작동합니다.\n"
        )
        QMessageBox.information(self, "도움말", help_text)
    
    def run_folder_crawling(self):
        folder_path = QFileDialog.getExistingDirectory(self, "설정 파일 폴더 선택", "")
        if not folder_path:
            QMessageBox.warning(self, "폴더 선택", "폴더를 선택하지 않았습니다.")
            return
        self.multi_thread = QThread(self)
        self.multi_worker = MultiCrawlerWorker(folder_path)
        self.multi_worker.moveToThread(self.multi_thread)
        self.multi_thread.started.connect(self.multi_worker.run)
        self.multi_worker.log_signal.connect(self.log)
        self.multi_worker.finished_signal.connect(self.on_multi_crawl_finished)
        self.multi_worker.finished_signal.connect(self.multi_thread.quit)
        self.multi_thread.start()

    def on_multi_crawl_finished(self, result):
        if self.auto_exit:
            QMessageBox.information(self, "폴더 크롤링 완료", f"{result}\n프로그램을 종료합니다.")
            QApplication.quit()
        else:
            QMessageBox.information(self, "폴더 크롤링 완료", f"{result}")

def run_app():
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
