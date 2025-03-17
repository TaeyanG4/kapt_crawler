import sys
from ui import run_app

def main():
    if len(sys.argv) > 1 and sys.argv[1].lower() == "help":
        print("=== K-APT 크롤러 도움말 ===\n")
        print("사용법:")
        print("  app.exe               : GUI 모드로 실행")
        print("  app.exe help          : 도움말 출력\n")
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
    run_app()

if __name__ == "__main__":
    main()
