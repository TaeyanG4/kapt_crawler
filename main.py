import sys
import os
import json
from ui import run_app

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

def main():
    if len(sys.argv) > 1:
        arg = sys.argv[1]
        if arg.lower() == "help":
            print("=== K-APT 크롤러 도움말 ===\n")
            print("사용법:")
            print("  python main.py              : GUI 모드로 실행")
            print("  python main.py help         : 도움말 출력")
            print("  python main.py <설정파일.json> : 설정 파일에 따라 CLI 모드로 크롤링 실행\n")
            print("GUI 사용 방법:")
            print("  1. 크롤링할 URL 입력 (빈 칸이면 기본 URL 사용)")
            print("  2. 추출할 데이터 건수 설정")
            print("  3. 크롤링 모드 선택:")
            print("     - 전체 페이지 + 상세정보 크롤링")
            print("     - 전체 페이지만 크롤링")
            print("     - 기존 엑셀 -> 상세정보만 크롤링")
            print("  4. 페이지 유형 선택 (수의계약, 경쟁입찰, 전국 입찰공고)")
            print("  5. 상세 컬럼 선택 및 설정 저장/불러오기 기능 활용")
            print("  6. 하단의 '크롤링 시작' 버튼 클릭")
            print("\n※ URL과 선택한 페이지 유형이 일치해야 정상 작동합니다.")
            sys.exit(0)
        elif arg.endswith(".json"):
            json_file = arg
            if not os.path.exists(json_file):
                print(f"설정 파일이 존재하지 않습니다: {json_file}")
                sys.exit(1)
            try:
                settings = read_json_with_encoding(json_file)
            except Exception as e:
                print(f"설정 파일 읽기 실패: {e}")
                sys.exit(1)
            
            url_text = settings.get("url", "").strip()
            extraction_count = settings.get("extraction_count", 50)
            mode = settings.get("mode", 1)
            page_type_index = settings.get("page_type_index", 0)
            selected_excel_path = settings.get("selected_excel_path", "")
            selected_detail_columns = settings.get("selected_detail_columns", [])
            
            def log_callback(msg):
                print(msg)
            
            from worker import CrawlerWorker
            print("CLI 모드 크롤링을 시작합니다...")
            worker = CrawlerWorker(mode, url_text, selected_excel_path, selected_detail_columns, extraction_count, page_type_index, log_callback=log_callback)
            result = worker.run()
            print("크롤링 결과:", result)
            sys.exit(0)
    # 인자가 없으면 기본적으로 GUI 모드 실행
    run_app()

if __name__ == "__main__":
    main()