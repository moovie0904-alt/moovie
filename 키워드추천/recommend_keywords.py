# -*- coding: utf-8 -*-
"""
핵심키워드(C열)마다 연관 키워드를 찾아, 경쟁강도(상품수 ÷ 월검색량)가 낮은 순으로
브랜드가 아닌 키워드 3개를 D열에 "과일트레이 트레이 나무트레이" 처럼 넣습니다.
E열에는 확인용으로 각 키워드의 경쟁강도·검색량·상품수를 적습니다.

준비: 이 폴더의 naver_key.txt 에 5줄을 적어 저장
    CUSTOMER_ID=      (검색광고 > SA API 사용 관리)
    API_KEY=          (검색광고 액세스라이선스)
    SECRET_KEY=       (검색광고 비밀키)
    CLIENT_ID=        (네이버 개발자센터 > 애플리케이션 Client ID)      ← 없어도 됨
    CLIENT_SECRET=    (네이버 개발자센터 > 애플리케이션 Client Secret)  ← 없어도 됨

CLIENT_ID/SECRET 이 없으면 상품 수를 못 구하므로, 검색광고 API의 경쟁정도(낮음<중간<높음)가
낮은 순 → 같은 등급이면 검색량 많은 순으로 고릅니다. 브랜드는 브랜드목록.txt 로만 걸러집니다.

사용법:
    python recommend_keywords.py                 (기본 파일, D열이 빈 줄만 처리)
    python recommend_keywords.py "D:\파일.xlsx"
    python recommend_keywords.py --limit 5       (시험용: 5개만)
"""
import base64
import datetime
import hashlib
import shutil
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
BASE_DIR = Path(r"C:\자동화시스템")
FALLBACK_INPUT = BASE_DIR / "합친결과_키워드.xlsx"
KEYWORD_COL = 3          # C열 = 핵심키워드
RESULT_COL = 4           # D열 = 추천키워드 3개
DETAIL_COL = 5           # E열 = 상세
PICK = 3                 # 뽑을 키워드 수
MIN_SEARCH = 100         # 월 검색량이 이보다 적은 키워드는 제외 (찾는 사람이 너무 적음)
MAX_CANDIDATES = 15      # 핵심키워드 하나당 상품수를 조회할 후보 수 (검색량 많은 순)
# ────────────────────────

HERE = Path(__file__).resolve().parent
BACKUP_KEEP = 5          # 백업 폴더에 남겨둘 개수
AD_URL, AD_URI = "https://api.searchad.naver.com", "/keywordstool"
SHOP_URL = "https://openapi.naver.com/v1/search/shop.json"


def default_input() -> Path:
    """C:\자동화시스템 에서 이름에 '핵심키워드'가 들어간 엑셀을 찾음. 없으면 합친결과_키워드.xlsx."""
    skip = ("~$", )
    found = sorted(p for p in BASE_DIR.glob("*핵심키워드*.xlsx")
                   if not p.name.startswith(skip) and "_새로저장" not in p.stem)
    if len(found) == 1:
        return found[0]
    if len(found) > 1:
        sys.exit("'핵심키워드'가 들어간 엑셀이 여러 개입니다. 쓸 파일을 bat 위로 끌어다 놓아 주세요:\n  "
                 + "\n  ".join(p.name for p in found))
    return FALLBACK_INPUT


def backup_and_save(wb, path: Path) -> None:
    """저장 전에 원본을 백업 폴더에 복사한 뒤, 원본 파일에 그대로 저장."""
    bdir = path.parent / "백업"
    bdir.mkdir(exist_ok=True)
    stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    shutil.copy2(path, bdir / f"{path.stem}_{stamp}{path.suffix}")
    for old in sorted(bdir.glob(f"{path.stem}_*{path.suffix}"))[:-BACKUP_KEEP]:
        old.unlink()
    try:
        wb.save(path)
        print(f"저장 완료: {path}  (백업: {bdir})")
    except PermissionError:
        alt = path.with_name(path.stem + "_새로저장.xlsx")
        wb.save(alt)
        print(f"※ 엑셀 파일이 열려 있어 원본에 저장하지 못했습니다. 다른 이름으로 저장: {alt}")
        print("   엑셀을 닫고 다시 실행하면 원본에 반영됩니다.")


def load_keys() -> dict:
    path = HERE / "naver_key.txt"
    if not path.exists():
        sys.exit(f"키 파일이 없습니다: {path}")
    keys = {}
    for line in path.read_text(encoding="utf-8-sig").splitlines():
        if "=" in line:
            k, v = line.split("=", 1)
            keys[k.strip().upper()] = v.strip()
    missing = [k for k in ("CUSTOMER_ID", "API_KEY", "SECRET_KEY") if not keys.get(k)]
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
    return [(i["relKeyword"], to_int(i.get("monthlyPcQcCnt", 0)) + to_int(i.get("monthlyMobileQcCnt", 0)),
             i.get("compIdx", ""))
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


COMP_RANK = {"낮음": 0, "중간": 1, "높음": 2}


def has_shopping(keys) -> bool:
    return bool(keys.get("CLIENT_ID") and keys.get("CLIENT_SECRET"))


def recommend(keys, core: str, my_brands: set[str]) -> tuple[str, str]:
    cands = {}
    for kw, vol, comp in related_keywords(keys, core):
        if vol >= MIN_SEARCH and is_related(kw, core):
            cands.setdefault(norm(kw), (kw, vol, comp))

    if not has_shopping(keys):  # 쇼핑 API 없이: 광고 경쟁정도 낮은 순 → 검색량 많은 순
        ranked = sorted((c for c in cands.values() if not is_brand(c[0], [], my_brands)),
                        key=lambda c: (COMP_RANK.get(c[2], 3), -c[1]))[:PICK]
        if not ranked:
            return "", "조회된 연관키워드 없음"
        return (" ".join(kw for kw, _, _ in ranked),
                " / ".join(f"{kw} 경쟁{comp} (검색 {vol:,})" for kw, vol, comp in ranked))

    top = sorted(cands.values(), key=lambda x: -x[1])[:MAX_CANDIDATES]
    scored = []
    for kw, vol, _ in top:
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
    args = sys.argv[1:]
    limit = None
    if "--limit" in args:
        i = args.index("--limit")
        limit = int(args[i + 1])
        del args[i:i + 2]
    in_path = Path(args[0]) if args else default_input()
    if not in_path.exists():
        sys.exit(f"파일을 찾을 수 없습니다: {in_path}")
    keys, my_brands = load_keys(), load_brands()
    if not has_shopping(keys):
        print("※ 쇼핑 API 키(CLIENT_ID/SECRET)가 없어 '광고 경쟁정도' 기준으로 고릅니다.")

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
    if limit:
        todo = dict(list(todo.items())[:limit])
    print(f"처리할 핵심키워드 {len(todo)}개 (D열이 이미 채워진 줄은 건너뜀)")

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
        backup_and_save(wb, in_path)


if __name__ == "__main__":
    main()
