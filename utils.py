import json

def read_json_with_encoding(file_path: str, encodings=None):
    """
    다양한 인코딩 방식으로 JSON 파일을 읽는 함수.
    :param file_path: JSON 파일 경로.
    :param encodings: 시도할 인코딩 리스트 (기본: ['utf-8', 'euc-kr']).
    :return: 파싱된 JSON 데이터.
    :raises ValueError: 모든 인코딩 방식으로 읽기에 실패한 경우.
    """
    if encodings is None:
        encodings = ['utf-8', 'euc-kr']
    for encoding in encodings:
        try:
            with open(file_path, "r", encoding=encoding) as f:
                return json.load(f)
        except (UnicodeDecodeError, json.JSONDecodeError):
            continue
    raise ValueError(f"파일을 읽을 수 없습니다. 지원되는 인코딩: {', '.join(encodings)}")
