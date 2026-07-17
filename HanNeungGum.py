# -*- coding: utf-8 -*-
"""
한능검 시험장 소요시간 계산기 ㅎㅎㅎㅎ하하하하
==================================

이 스크립트는:
1. 엑셀 파일(시험장 명 / 주소 / 여부)을 읽어 "여부"가 "가능"인 시험장만 추립니다.
2. 각 주소를 네이버 지오코딩(Geocoding) API로 위경도 좌표로 변환합니다.
3. 지정한 출발지(예: 서울역, 동서울터미널)에서 각 시험장까지의 "대중교통" 소요시간을
   ODsay 대중교통 길찾기 API로 조회합니다.
   (※ 참고: 네이버 클라우드 플랫폼은 '자동차 길찾기(Directions)' API는 공식 제공하지만
    '대중교통(버스+지하철) 길찾기'는 별도의 공개 API를 제공하지 않습니다.
    그래서 실제 서비스 현장에서 대중교통 경로/소요시간 계산에는 국내 표준으로 통용되는
    ODsay Lab의 대중교통 API(https://lab.odsay.com)를 함께 사용합니다.
    주소→좌표 변환(지오코딩)만 네이버 API를 쓰고, 대중교통 소요시간 계산은 ODsay를 씁니다.)
4. 결과를 출발지별로 소요시간 오름차순 정렬하여 보기 좋은 HTML 파일로 저장합니다.

사전 준비물
------------
1) 네이버 클라우드 플랫폼(NCP) 계정 -> Console -> AI·NAVER API -> Application 등록
   -> "Maps" 서비스 중 "Geocoding" 활성화 -> Client ID / Client Secret 발급
   https://www.ncloud.com/product/applicationService/maps

2) ODsay Lab 계정 -> API 키 발급 (무료 티어 존재, 일일/월 호출 한도 있음)
   https://lab.odsay.com

3) 파이썬 패키지 설치
   pip install openpyxl requests --break-system-packages

사용법
------
python transit_time_calculator.py

(엑셀 파일 경로 / 출발지 좌표 / API 키는 아래 CONFIG 영역에서 수정하세요)
"""

import json
import os
import sys
import time
from datetime import datetime

import openpyxl
import requests

# ============================================================
# CONFIG - 아래 값들을 본인 환경에 맞게 수정하세요
# ============================================================

SHEET_NAME = None                      # None이면 첫 번째 시트 사용
STATUS_COLUMN_OK_VALUE = "가능"        # 이 값을 가진 행만 대상으로 함
if getattr(sys, "frozen", False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_PATH = os.path.join(BASE_DIR, "한능검_시험장.xlsx")     # 엑셀 파일 경로
OUTPUT_HTML_PATH = os.path.join(BASE_DIR, "시험장_소요시간.html")
CACHE_PATH = os.path.join(BASE_DIR, "transit_cache.json")      # 지오코딩/경로 결과 캐시 (재실행 시 API 절약)

# 네이버 지오코딩 API 인증 정보
NAVER_CLIENT_ID = os.environ.get("NAVER_CLIENT_ID", "여기를 CLIENT ID 값으로 바꾸세요") #두번째 인수를 실제 CLIENT ID 값으로 수정
NAVER_CLIENT_SECRET = os.environ.get("NAVER_CLIENT_SECRET", "여기를 CLIENT SECRET 값으로 바꾸세요") #두번째 인수를 CLIENT SECRET 값으로 수정

# ODsay 대중교통 API 인증 정보
ODSAY_API_KEY = os.environ.get("ODSAY_API_KEY", "여기를 API KEY로 바꾸세요") #두번째 인수를 실제 API KEY 값으로 수정

# 출발지 목록: 이름 -> 주소(지오코딩으로 좌표 변환)
ORIGINS = {
    "서울역": "서울특별시 중구 한강대로 405",
    "동서울터미널": "서울 광진구 강변역로 50 동서울종합터미널",
}

REQUEST_DELAY_SEC = 0.3   # API 호출 사이 딜레이(과호출 방지)
REQUEST_TIMEOUT_SEC = 15

# ============================================================
# 캐시 유틸
# ============================================================

def load_cache():
    if os.path.exists(CACHE_PATH):
        with open(CACHE_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    return {"geocode": {}, "transit": {}}


def save_cache(cache):
    with open(CACHE_PATH, "w", encoding="utf-8") as f:
        json.dump(cache, f, ensure_ascii=False, indent=2)

def get_available_filename(filepath):
    """
    같은 이름의 파일이 있으면
    시험장_소요시간.html
    -> 시험장_소요시간 (1).html
    -> 시험장_소요시간 (2).html
    ...
    형태의 파일명을 반환
    """
    if not os.path.exists(filepath):
        return filepath

    base, ext = os.path.splitext(filepath)

    i = 1
    while True:
        new_path = f"{base} ({i}){ext}"
        if not os.path.exists(new_path):
            return new_path
        i += 1


# ============================================================
# 1) 엑셀에서 "가능" 시험장만 읽기
# ============================================================

def load_available_venues(excel_path, sheet_name=None):
    if not os.path.exists(excel_path):
        raise FileNotFoundError(f"엑셀 파일을 찾을 수 없습니다.\n{excel_path}")

    wb = openpyxl.load_workbook(excel_path, data_only=True)
    ws = wb[sheet_name] if sheet_name else wb.active

    rows = list(ws.iter_rows(values_only=True))
    header = [str(h).strip() if h else "" for h in rows[0]]

    def col_index(*candidates):
        for cand in candidates:
            if cand in header:
                return header.index(cand)
        return None

    idx_name = col_index("시험장 명", "시험장명", "이름")
    idx_addr = col_index("주소")
    idx_status = col_index("여부", "상태")

    if idx_name is None or idx_addr is None or idx_status is None:
        raise ValueError(
        f"필요한 열(시험장 명/주소/여부)을 찾지 못했습니다.\n"
        f"현재 헤더: {header}"
        )

    venues = []
    for row in rows[1:]:
        if row[idx_name] is None:
            continue
        status = row[idx_status]
        if status is None or str(status).strip() != STATUS_COLUMN_OK_VALUE:
            continue
        venues.append({
            "name": str(row[idx_name]).strip(),
            "address": str(row[idx_addr]).strip(),
        })
    return venues


# ============================================================
# 2) 네이버 지오코딩: 주소 -> (lat, lon)
# ============================================================

def geocode_address(address, cache):
    if address in cache["geocode"]:
        return cache["geocode"][address]

    url = "https://maps.apigw.ntruss.com/map-geocode/v2/geocode"
    headers = {
        "X-NCP-APIGW-API-KEY-ID": NAVER_CLIENT_ID,
        "X-NCP-APIGW-API-KEY": NAVER_CLIENT_SECRET,
    }
    params = {"query": address}

    try:
        resp = requests.get(url, headers=headers, params=params, timeout=REQUEST_TIMEOUT_SEC)
        print(resp.status_code)
        print(resp.headers.get("Content-Type"))
        print(resp.text)
        resp.raise_for_status()
        data = resp.json()
    except requests.RequestException as e:
        print(f"  [지오코딩 실패] {address} -> {e}")
        cache["geocode"][address] = None
        return None

    addresses = data.get("addresses") or []
    if not addresses:
        print(f"  [지오코딩 결과 없음] {address}")
        cache["geocode"][address] = None
        return None

    best = addresses[0]
    result = {"lat": float(best["y"]), "lon": float(best["x"])}
    cache["geocode"][address] = result
    return result


# ============================================================
# 3) ODsay 대중교통 경로 조회: (olon, olat) -> (dlon, dlat)
# ============================================================

def get_transit_route(o_lon, o_lat, d_lon, d_lat, cache):
    key = f"{o_lon},{o_lat}->{d_lon},{d_lat}"
    if key in cache["transit"]:
        return cache["transit"][key]

    url = "https://api.odsay.com/v1/api/searchPubTransPathT"
    params = {
        "SX": o_lon, "SY": o_lat,
        "EX": d_lon, "EY": d_lat,
        "apiKey": ODSAY_API_KEY,
    }

    try:
        resp = requests.get(url, params=params, timeout=REQUEST_TIMEOUT_SEC)
        resp.raise_for_status()
        data = resp.json()
    except requests.RequestException as e:
        print(f"  [경로 조회 실패] {e}")
        cache["transit"][key] = None
        return None

    if "error" in data:
        # ODsay는 실패 시에도 200을 반환하며 error 필드를 넣는 경우가 있음
        print(f"  [ODsay 오류] {data['error']}")
        cache["transit"][key] = None
        return None

    try:
        paths = data["result"]["path"]
    except (KeyError, TypeError):
        cache["transit"][key] = None
        return None

    if not paths:
        cache["transit"][key] = None
        return None

    # totalTime(분) 기준으로 가장 빠른 경로 선택
    fastest = min(paths, key=lambda p: p["info"]["totalTime"])
    info = fastest["info"]

    result = {
        "total_time_min": info.get("totalTime"),
        "transfer_count": info.get("busTransitCount", 0) + info.get("subwayTransitCount", 0),
        "total_distance_m": info.get("totalDistance"),
        "total_walk_m": info.get("totalWalk"),
        "path_type": fastest.get("pathType"),  # 1:지하철, 2:버스, 3:버스+지하철
    }
    cache["transit"][key] = result
    return result


# ============================================================
# 4) 메인 로직
# ============================================================

def main():
    print("1) 엑셀에서 '가능' 시험장 목록을 읽는 중...")
    venues = load_available_venues(EXCEL_PATH, SHEET_NAME)
    print(f"   -> 대상 시험장 {len(venues)}곳")

    cache = load_cache()

    print("2) 출발지 좌표 지오코딩 중...")
    origin_coords = {}
    for label, addr in ORIGINS.items():
        coord = geocode_address(addr, cache)
        if coord is None:
            raise RuntimeError(
                f"출발지 '{label}' 좌표를 확인할 수 없습니다.\n주소를 확인하세요: {addr}"
            )
        origin_coords[label] = coord
        print(f"   - {label}: {coord}")
        time.sleep(REQUEST_DELAY_SEC)

    print("3) 시험장 좌표 지오코딩 + 대중교통 소요시간 조회 중...")
    results = {label: [] for label in origin_coords}

    for i, v in enumerate(venues, 1):
        print(f"   ({i}/{len(venues)}) {v['name']}")
        coord = geocode_address(v["address"], cache)
        if coord is None:
            for label in results:
                results[label].append({**v, "minutes": None, "detail": None})
            continue

        v["lat"], v["lon"] = coord["lat"], coord["lon"]

        for label, ocoord in origin_coords.items():
            route = get_transit_route(
                ocoord["lon"], ocoord["lat"], coord["lon"], coord["lat"], cache
            )
            minutes = route["total_time_min"] if route else None
            results[label].append({**v, "minutes": minutes, "detail": route})
            time.sleep(REQUEST_DELAY_SEC)

        # 중간 저장 (중단되어도 캐시로 이어서 실행 가능)
        save_cache(cache)

    print("4) 결과 정렬 및 HTML 생성 중...")
    for label in results:
        results[label].sort(key=lambda x: (x["minutes"] is None, x["minutes"]))

    html = build_html(results, origin_coords)

    output_path = get_available_filename(OUTPUT_HTML_PATH)

    with open(output_path, "w", encoding="utf-8") as f:
        f.write(html)

    print(f"완료! -> {os.path.abspath(output_path)}")


# ============================================================
# 5) HTML 생성
# ============================================================

def build_html(results, origin_coords):
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M")

    sections_html = ""
    for label, venues in results.items():
        rows_html = ""
        for rank, v in enumerate(venues, 1):
            minutes = v["minutes"]
            detail = v["detail"] or {}
            time_str = f"{minutes}분" if minutes is not None else "조회 실패"
            transfer = detail.get("transfer_count", "-")
            walk_m = detail.get("total_walk_m")
            walk_str = f"{walk_m}m" if walk_m is not None else "-"
            highlight = ' class="best"' if rank == 1 and minutes is not None else ""
            rows_html += f"""
            <tr{highlight}>
              <td>{rank}</td>
              <td>{v['name']}</td>
              <td>{v['address']}</td>
              <td class="time">{time_str}</td>
              <td>{transfer}</td>
              <td>{walk_str}</td>
            </tr>"""

        sections_html += f"""
        <section>
          <h2>출발지: {label}</h2>
          <table>
            <thead>
              <tr>
                <th>순위</th>
                <th>시험장</th>
                <th>주소</th>
                <th>소요시간</th>
                <th>환승 횟수</th>
                <th>도보 거리</th>
              </tr>
            </thead>
            <tbody>
              {rows_html}
            </tbody>
          </table>
        </section>"""

    return f"""<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="UTF-8">
<title>한능검 시험장 대중교통 소요시간</title>
<style>
  body {{ font-family: -apple-system, "Malgun Gothic", sans-serif; background:#f7f7f8; margin:0; padding:32px; color:#222; }}
  h1 {{ font-size: 22px; margin-bottom:4px; }}
  .meta {{ color:#888; font-size:13px; margin-bottom:24px; }}
  section {{ background:#fff; border-radius:12px; padding:20px 24px; margin-bottom:28px; box-shadow:0 1px 3px rgba(0,0,0,0.08); }}
  h2 {{ font-size:17px; margin-top:0; }}
  table {{ width:100%; border-collapse: collapse; font-size:14px; }}
  th, td {{ text-align:left; padding:8px 10px; border-bottom:1px solid #eee; }}
  th {{ color:#666; font-weight:600; background:#fafafa; }}
  td.time {{ font-weight:600; }}
  tr.best {{ background:#eef8f0; }}
  tr.best td.time {{ color:#1a8a3c; }}
</style>
</head>
<body>
  <h1>한능검 시험장 대중교통 소요시간 비교</h1>
  <div class="meta">생성 시각: {now_str} · 소요시간 출처: ODsay 대중교통 API (지오코딩: 네이버 지도 API) · "여부=가능" 시험장만 포함</div>
  {sections_html}
</body>
</html>"""


if __name__ == "__main__":
    try:
        main()
        print()
        print("모든 작업이 완료되었습니다.")
    except Exception as e:
        print()
        print("=" * 60)
        print("프로그램 실행 중 오류가 발생했습니다.")
        print(e)
        print("=" * 60)
        
    input("\n엔터를 누르면 종료됩니다.")
