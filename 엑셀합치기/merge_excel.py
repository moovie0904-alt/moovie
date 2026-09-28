# -*- coding: utf-8 -*-
"""
폴더 안의 모든 엑셀(.xlsx/.xlsm/.xls)의 '디셀' 시트와 CSV 파일을 하나로 합치고,
완전히 똑같은 행(중복)은 한 번만 남깁니다.

사용법:
    python merge_excel.py                      (기본 폴더 사용)
    python merge_excel.py "D:\다른폴더"         (폴더 직접 지정)
"""
import sys
from pathlib import Path

import pandas as pd

DEFAULT_FOLDER = r"C:\자동화시스템\디셀양식"
OUTPUT_NAME = "합친결과.xlsx"
SOURCE_COL = "원본파일"
SHEET_NAME = "디셀"        # 이 이름의 시트만 합침 (없으면 첫 번째 시트)
KEY_COLUMN = "상품명"      # 이 칸이 있는 줄을 제목 줄로 봄
VERSION = "3"


def clean_header(value, idx: int) -> str:
    """제목 칸 이름 정리: 공백/줄바꿈 제거 ('상품 명' -> '상품명'). 비어있으면 '열N'."""
    if value is None or pd.isna(value) or str(value).strip() == "":
        return f"열{idx + 1}"
    return "".join(str(value).split())


def to_table(raw: pd.DataFrame) -> pd.DataFrame:
    """제목 줄 위에 '발주양식' 같은 제목이 있어도, 실제 제목 줄을 찾아서 표로 만든다."""
    raw = raw.dropna(how="all").dropna(axis=1, how="all")
    if raw.empty:
        return raw
    # '상품명' 칸이 있는 줄을 제목 줄로 본다. 없으면 위쪽 30줄 중 가장 많이 채워진 줄.
    top = raw.head(30)
    header_pos = None
    for pos in range(len(top)):
        if any(clean_header(v, 0) == KEY_COLUMN for v in top.iloc[pos]):
            header_pos = pos
            break
    if header_pos is None:
        header_pos = int(top.notna().sum(axis=1).to_numpy().argmax())

    names, seen = [], {}
    for i, v in enumerate(raw.iloc[header_pos]):
        name = clean_header(v, i)
        if name in seen:  # 같은 이름의 열이 또 있으면 뒤에 번호 붙임
            seen[name] += 1
            name = f"{name}_{seen[name]}"
        else:
            seen[name] = 1
        names.append(name)

    table = raw.iloc[header_pos + 1:].copy()
    table.columns = names
    return table


def read_file(path: Path) -> list[pd.DataFrame]:
    """파일 하나를 읽어 시트별 표 목록을 돌려준다."""
    if path.suffix.lower() == ".csv":
        for enc in ("utf-8-sig", "cp949"):
            try:
                raw = pd.read_csv(path, dtype=str, encoding=enc, header=None)
                return [to_table(raw).assign(**{SOURCE_COL: path.name})]
            except UnicodeDecodeError:
                continue
        raise ValueError("CSV 인코딩을 알 수 없습니다")

    names = pd.ExcelFile(path).sheet_names
    sheet = SHEET_NAME if SHEET_NAME in names else names[0]
    raw = pd.read_excel(path, sheet_name=sheet, dtype=str, header=None)
    table = to_table(raw)
    if KEY_COLUMN not in table.columns:
        print(f"  ※ 주의: {path.name} [{sheet}] 시트에서 '{KEY_COLUMN}' 칸을 찾지 못했습니다")
    return [table.assign(**{SOURCE_COL: path.name})]


def main() -> None:
    print(f"엑셀 합치기 (버전 {VERSION})")
    folder = Path(sys.argv[1] if len(sys.argv) > 1 else DEFAULT_FOLDER)
    if not folder.is_dir():
        sys.exit(f"폴더를 찾을 수 없습니다: {folder}")

    files = sorted(
        p for p in folder.iterdir()
        if p.suffix.lower() in (".xlsx", ".xlsm", ".xls", ".csv")
        and p.name != OUTPUT_NAME
        and not p.name.startswith("~$")  # 엑셀이 열려 있을 때 생기는 임시파일 제외
    )
    if not files:
        sys.exit(f"엑셀 파일이 없습니다: {folder}")

    frames = []
    for f in files:
        try:
            got = read_file(f)
            frames.extend(got)
            print(f"  읽음: {f.name} ({sum(len(d) for d in got)}행)")
        except Exception as e:
            print(f"  건너뜀: {f.name} - {e}")

    merged = pd.concat(frames, ignore_index=True, sort=False)

    # 값 정리: 앞뒤 공백 제거, 빈 칸은 비어있는 값으로 통일
    data_cols = [c for c in merged.columns if c != SOURCE_COL]
    merged[data_cols] = merged[data_cols].apply(lambda col: col.str.strip()).replace("", pd.NA)

    # 완전히 빈 행 제거
    merged = merged.dropna(how="all", subset=data_cols)

    # 중복 제거: 원본파일 이름은 빼고, 내용이 모두 같은 행은 처음 것만 남김
    before = len(merged)
    result = merged.drop_duplicates(subset=data_cols, keep="first")
    removed = before - len(result)

    # 원본파일 열은 맨 뒤로
    result = result[data_cols + [SOURCE_COL]]
    print(f"합친 열: {', '.join(map(str, data_cols))}")

    out = folder / OUTPUT_NAME
    result.to_excel(out, index=False)

    print()
    print(f"파일 {len(files)}개 합침 -> 총 {before}행, 중복 {removed}행 제거, 최종 {len(result)}행")
    print(f"저장 위치: {out}")


if __name__ == "__main__":
    main()
