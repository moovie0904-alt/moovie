# -*- coding: utf-8 -*-
"""
핵심키워드(C열)마다 연관 키워드를 찾아, 경쟁강도(상품수 ÷ 월검색량)가 낮은 순으로
브랜드가 아닌 키워드 3개를 D열에 "과일트레이 트레이 나무트레이" 처럼 넣습니다.
E열에는 확인용으로 각 키워드의 경쟁강도·검색량·상품수를 적습니다.

준비: 이 폴더의 naver_key.txt 에 5줄을 적어 저장
    CUSTOMER_ID=      (검색광고 > SA API 사용 관리)
    API_KEY=          (검색광고 액세스라이선스)
    SECRET_KEY=       (검색광고 비밀키)
    CLIENT_ID=        (네이버 개발자센터 > 애플리케이션 Client ID)
    CLIENT_SECRET=    (네이버 개발자센터 > 애플리케이션 Client Secret)

사용법:
    python recommend_keywords.py                 (기본 파일, D열이 빈 줄만 처리)
    python recommend_keywords.py "D:\파일.xlsx"
"""
import base64
import hashlib
import hmac
import html
import re
import sys
import time
from collections import Counter
from pathlib import Path

import requests
from openpyxl import load_workbook

# ───────── 설정 ─────────
DEFAULT_INPUT = r"C:\자동화시스템\합친결과_키워드.xlsx"
KEYWORD_COL = 3          # C열 = 핵심키워드
RESULT_COL = 4           # D열 = 추천키워드 3개
DETAIL_COL = 5           # E열 = 상세
PICK = 3                 # 뽑을 키워드 수
MIN_SEARCH = 100         # 월 검색량이 이보다 적은 키워드는 제외 (찾는 사람이 너무 적음)
MAX_CANDIDATES = 15      # 핵심키워드 하나당 상품수를 조회할 후보 수 (검색량 많은 순)
# ────────────────────────

HERE = Path(__file__).resolve().parent
AD_URL, AD_URI = "https://api.searchad.naver.com", "/keywordstool"
SHOP_URL = "https://openapi.naver.com/v1/search/shop.json"


def load_keys() -> dict:
    path = HERE / "naver_key.txt"
    if not path.exists():
        sys.exit(f"키 파일이 없습니다: {path}")
    keys = {}
    for line in path.read_text(encoding="utf-8-sig").splitlines():
        if "=" in line:
            k, v = line.split("=", 1)
            keys[k.strip().upper()] = v.strip()
    missing = [k for k in ("CUSTOMER_ID", "API_KEY", "SECRET_KEY", "CLIENT_ID", "CLIENT_SECRET") if not keys.get(k)]
    if missing:
        sys.exit(f"naver_key.txt 에 값이 없습니다: {', '.join(missing)}")
    return keys


def load_brands() -> set[str]:
    """브랜드목록.txt: 직접 추가하고 싶은 브랜드 (한 줄에 하나)."""
    path = HERE / "브랜드목록.txt"
    if not path.exists():
        return set()
    return {norm(l) for l in path.read_text(encoding="utf-8-sig").splitlines() if l.strip()}


def norm(s) -> str:
    return "".join(str(s).split()).upper()


def to_int(v) -> int:
    if isinstance(v, (int, float)):
        return int(v)
    return 5 if "<" in str(v) else int(str(v).replace(",", "") or 0)


def get(url, **kw):
    for attempt in range(4):
        r = requests.get(url, timeout=20, **kw)
        if r.status_code == 429:
            time.sleep(2 * (attempt + 1))
            continue
        if r.status_code in (401, 403):
            sys.exit(f"인증 실패 ({r.status_code}) - naver_key.txt 의 키를 확인하세요: {url}\n{r.text[:200]}")
        r.raise_for_status()
        return r.json()
    raise RuntimeError("요청이 계속 거절됩니다. 잠시 후 다시 실행하세요.")


def related_keywords(keys, core: str) -> list[tuple[str, int]]:
    """검색광고 API: 핵심키워드의 연관 키워드와 월 검색량(PC+모바일)."""
    ts = str(int(time.time() * 1000))
    sig = base64.b64encode(hmac.new(keys["SECRET_KEY"].encode(), f"{ts}.GET.{AD_URI}".encode(),
                                    hashlib.sha256).digest()).decode()
    data = get(AD_URL + AD_URI,
               headers={"X-Timestamp": ts, "X-API-KEY": keys["API_KEY"], "X-Customer": keys["CUSTOMER_ID"],
                        "X-Signature": sig},
               params={"hintKeywords": core.replace(" ", ""), "showDetail": "1"})
    return [(i["relKeyword"], to_int(i.get("monthlyPcQcCnt", 0)) + to_int(i.get("monthlyMobileQcCnt", 0)))
            for i in data.get("keywordList", [])]


def shopping(keys, kw: str) -> tuple[int, list[str]]:
    """쇼핑검색 API: 상품 수와 상위 상품들의 브랜드/제조사."""
    data = get(SHOP_URL, headers={"X-Naver-Client-Id": keys["CLIENT_ID"], "X-Naver-Client-Secret": keys["CLIENT_SECRET"]},
               params={"query": kw, "display": 20})
    names = []
    for it in data.get("items", []):
        for f in ("brand", "maker"):
            v = html.unescape(re.sub("<.*?>", "", it.get(f, "") or "")).strip()
            if len(norm(v)) >= 2:
                names.append(norm(v))
    return int(data.get("total", 0)), names


def is_brand(kw: str, shop_brands: list[str], my_brands: set[str]) -> bool:
    k = norm(kw)
    if any(b in k for b in my_brands):
        return True
    # 상위 상품 여러 개(4개 이상)가 같은 브랜드이고, 그 브랜드 이름이 키워드에 들어 있으면 브랜드 키워드로 봄
    return any(cnt >= 4 and b in k for b, cnt in Counter(shop_brands).items())


def is_related(cand: str, core: str) -> bool:
    """같은 종류의 상품인지: 끝부분(제품 종류)이 2글자 이상 같거나, 핵심키워드 안에 들어 있으면 관련 있음."""
    a, b = norm(cand), norm(core)
    if len(a) >= 2 and a in b:
        return True
    n = 0
    while n < min(len(a), len(b)) and a[-1 - n] == b[-1 - n]:
        n += 1
    return n >= 2


def recommend(keys, core: str, my_brands: set[str]) -> tuple[str, str]:
    cands = {}
    for kw, vol in related_keywords(keys, core):
        if vol >= MIN_SEARCH and is_related(kw, core):
            cands.setdefault(norm(kw), (kw, vol))
    top = sorted(cands.values(), key=lambda x: -x[1])[:MAX_CANDIDATES]

    scored = []
    for kw, vol in top:
        total, brands = shopping(keys, kw)
        if is_brand(kw, brands, my_brands):
            continue
        scored.append((total / vol, kw, vol, total))
        time.sleep(0.1)
    scored.sort()
    best = scored[:PICK]
    if not best:
        return "", "조회된 연관키워드 없음"
    return (" ".join(kw for _, kw, _, _ in best),
            " / ".join(f"{kw} 경쟁{r:.1f} (검색 {v:,}, 상품 {t:,})" for r, kw, v, t in best))


def main() -> None:
    in_path = Path(sys.argv[1] if len(sys.argv) > 1 else DEFAULT_INPUT)
    if not in_path.exists():
        sys.exit(f"파일을 찾을 수 없습니다: {in_path}")
    keys, my_brands = load_keys(), load_brands()

    print(f"읽는 중: {in_path} (상품이 많으면 1~2분 걸립니다)")
    wb = load_workbook(in_path)
    ws = wb.active
    ws.cell(1, RESULT_COL).value = "추천키워드(경쟁강도 낮은순)"
    ws.cell(1, DETAIL_COL).value = "상세"

    todo = {}
    for r in range(2, ws.max_row + 1):
        core, done = ws.cell(r, KEYWORD_COL).value, ws.cell(r, RESULT_COL).value
        if core and str(core).strip() and not done:
            todo.setdefault(str(core).strip(), []).append(r)
    print(f"처리할 핵심키워드 {len(todo)}개 (D열이 이미 채워진 줄은 건너뜀)")

    out = in_path.with_name(in_path.stem + "_추천.xlsx") if "_추천" not in in_path.stem else in_path
    try:
        for n, (core, rows) in enumerate(todo.items(), 1):
            try:
                result, detail = recommend(keys, core, my_brands)
            except requests.RequestException as e:
                result, detail = "", f"오류: {e}"
            for r in rows:
                ws.cell(r, RESULT_COL).value = result
                ws.cell(r, DETAIL_COL).value = detail
            print(f"  [{n}/{len(todo)}] {core} → {result or detail}")
    except KeyboardInterrupt:
        print("\n중단됨 - 지금까지 결과를 저장합니다.")
    finally:
        print("저장 중...")
        wb.save(out)
        print(f"저장 완료: {out}")


if __name__ == "__main__":
    main()
