import os
from datetime import datetime
import openpyxl
import pandas as pd

def make_unique_filename(base_name="추출데이터", folder_name="추출데이터"):
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
    else:
        ws.title = "입찰공고"
        headers = ["순번", "종류", "낙찰방법", "입찰공고명", "입찰마감일", "상태", "단지명", "공고일", "상세정보링크"]
    
    ws.append(headers)
    for item in data_list:
        if page_type_index == 0:
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
        else:
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
    try:
        wb.save(filename)
    except Exception as e:
        raise Exception(f"엑셀 파일 저장 실패: {e}")

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

def crawl_detail_info_from_excel(input_excel_path, selected_columns, detail_crawler, log_callback=None, page_type_index=0):
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

    try:
        df_input = pd.read_excel(input_excel_path)
    except Exception as e:
        _log(f"엑셀 파일 읽기 실패: {e}")
        return None

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
                crawled_data = detail_crawler.crawl_detail_page(detail_url)
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
            failed_data = {col: 'FAILED' for col in selected_columns}
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
    try:
        df_result.to_excel(output_excel_path, index=False)
    except Exception as e:
        _log(f"결과 엑셀 파일 저장 실패: {e}")
        return None
    _log(f"\n상세정보 크롤링 완료! 결과: {output_excel_path}")
    return output_excel_path
