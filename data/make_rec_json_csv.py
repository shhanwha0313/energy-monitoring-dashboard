# 파일명: make_rec_json_csv.py
# 목적: 한국전력거래소_오늘의 REC 시장 CSV 파일에서 대시보드용 rec_index_data.json 생성
#
# 입력 CSV에서 사용하는 컬럼:
# - 거래일
# - 거래량(REC)
# - 가격(원)
# - 대비
# - 등락률
#
# 출력 JSON 구조:
# [
#   {
#     "date": "2025-03-31",
#     "price": 76000,
#     "change": 100,
#     "change_rate": 0.13,
#     "volume": 123456
#   }
# ]
#
# 사용 방법:
# 1) 이 파일과 CSV 파일을 같은 폴더에 둡니다.
#    예: C:\dashboard\data\make_rec_json_csv.py
#        C:\dashboard\data\한국전력거래소_오늘의 REC 시장.csv
# 2) 아래 INPUT_FILE 파일명을 실제 CSV 파일명과 맞춥니다.
# 3) 실행합니다.
#    python make_rec_json_csv.py
#
# 별도 확장자 변환은 필요 없습니다. CSV 그대로 사용합니다.

from pathlib import Path
from datetime import datetime, timedelta
import csv
import json
import math
import re


# =========================
# 1. 경로 설정
# =========================
BASE_DIR = Path(__file__).resolve().parent

INPUT_FILE = BASE_DIR / "한국전력거래소_오늘의 REC 시장.csv"
OUTPUT_FILE = BASE_DIR / "rec_index_data.json"

# 최신 거래일 기준 최근 1년치만 저장
DAYS_TO_KEEP = 365


# =========================
# 2. 변환 함수
# =========================
def read_csv_with_encoding(path: Path):
    """REC CSV는 보통 CP949 형식입니다. UTF-8도 함께 시도합니다."""
    encodings = ["cp949", "utf-8-sig", "utf-8"]

    last_error = None
    for enc in encodings:
        try:
            with open(path, "r", encoding=enc, newline="") as f:
                return list(csv.DictReader(f))
        except UnicodeDecodeError as exc:
            last_error = exc

    raise UnicodeDecodeError(
        "unknown", b"", 0, 1,
        f"CSV 인코딩을 읽지 못했습니다. 마지막 오류: {last_error}"
    )


def parse_date(value):
    """문자 날짜를 YYYY-MM-DD로 변환합니다."""
    if value is None:
        return None

    text = str(value).strip()
    if not text:
        return None

    text = text.replace(".", "-").replace("/", "-")
    text = re.sub(r"\s+", "", text)

    for fmt in ("%Y-%m-%d", "%y-%m-%d", "%Y%m%d"):
        try:
            return datetime.strptime(text, fmt).date().isoformat()
        except ValueError:
            pass

    return None


def parse_number(value):
    """쉼표, %, +, ▲, ▼ 등이 포함된 값을 숫자로 변환합니다."""
    if value is None:
        return None

    if isinstance(value, float) and math.isnan(value):
        return None

    if isinstance(value, (int, float)):
        num = float(value)
    else:
        text = str(value).strip()

        if text in {"", "-", "NaN", "nan", "N/A", "n/a", "None"}:
            return None

        is_negative = False
        if "▼" in text or text.startswith("-") or text.startswith("−"):
            is_negative = True

        if text.startswith("(") and text.endswith(")"):
            is_negative = True
            text = text[1:-1]

        text = (
            text.replace(",", "")
                .replace("%", "")
                .replace("▲", "")
                .replace("▼", "")
                .replace("+", "")
                .replace("−", "-")
                .strip()
        )

        try:
            num = float(text)
        except ValueError:
            return None

        if is_negative and num > 0:
            num = -num

    if float(num).is_integer():
        return int(num)

    return round(num, 4)


def first_existing(row, keys):
    """여러 후보 컬럼명 중 존재하는 값을 반환합니다."""
    for key in keys:
        if key in row:
            return row.get(key)
    return None


# =========================
# 3. 메인 로직
# =========================
def main():
    if not INPUT_FILE.exists():
        raise FileNotFoundError(
            f"입력 파일을 찾을 수 없습니다: {INPUT_FILE}\n"
            "make_rec_json_csv.py와 CSV 파일을 같은 폴더에 두었는지 확인해 주세요."
        )

    raw_rows = read_csv_with_encoding(INPUT_FILE)
    rows = []

    for row in raw_rows:
        item = {
            "date": parse_date(first_existing(row, ["거래일", "date", "Date"])),
            "price": parse_number(first_existing(row, ["가격(원)", "가격", "price"])),
            "change": parse_number(first_existing(row, ["대비", "change"])),
            "change_rate": parse_number(first_existing(row, ["등락률", "등락율", "change_rate"])),
            "volume": parse_number(first_existing(row, ["거래량(REC)", "거래량", "volume"])),
        }

        if not item["date"]:
            continue

        # 가격이 없으면 대시보드 표시 대상에서 제외
        if item["price"] is None:
            continue

        rows.append(item)

    if not rows:
        raise RuntimeError("변환 가능한 REC 데이터가 없습니다.")

    # 최신일자 우선 정렬
    rows.sort(key=lambda x: x["date"], reverse=True)

    # 최신 거래일 기준 최근 1년치만 추출
    latest_date = datetime.strptime(rows[0]["date"], "%Y-%m-%d").date()
    start_date = latest_date - timedelta(days=DAYS_TO_KEEP)

    rows_1y = [
        row for row in rows
        if datetime.strptime(row["date"], "%Y-%m-%d").date() >= start_date
    ]

    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(rows_1y, f, ensure_ascii=False, indent=2)

    print(f"저장 완료: {OUTPUT_FILE}")
    print(f"기간: {rows_1y[-1]['date']} ~ {rows_1y[0]['date']}")
    print(f"건수: {len(rows_1y):,}건")

    print("\n미리보기:")
    for item in rows_1y[:5]:
        print(item)


if __name__ == "__main__":
    main()
