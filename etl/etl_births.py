# /// script
# requires-python = ">=3.11"
# dependencies = ["pandas", "xlrd"]
# ///
"""raw/ -> data/ 長表 CSV(UTF-8 with BOM)。

統計基準:全部採「按發生日期」(出生當年),不與「按登記日期」混用。
  - births_city / births_district: 鄉鎮市區出生死亡結婚離婚(96).xls 的「按發生日期統計」欄
  - births_mother_age: 縣市出生按生母五齡組(按發生)(96).xls(該表只有按發生)
  - 粗出生率原值: 縣市育齡婦女生育率(按發生)(96).xls 的 C.B.R(按發生);
                 各區沒有按發生的原始粗出生率,一律計算。
  新北市民政局「各區粗出生率」xlsx 經核對(板橋區114年: 3.6418‰×年中人口≈2,009人=按登記出生數,
  非按發生1,988人)是按登記基準,依規則不混用,故不納入;原檔留在 raw/ntpc_ca/ 僅供對照。
  因此涵蓋民國96年(2007)起;民國87-95年只有按登記資料,依規則不納入。
  縣市出生按登記逐年檔(87-114年)仍保留在 raw/,僅作對照,本程式不讀。

年份:民國年一律 +1911 轉西元年。尚未滿一整年的年份(>= 今年)整列移除。
臺北縣:民國99年(2010)以前對應到新北市,name_at_time 保留來源表當時印的名稱。
  其他改名/合併(桃園縣→桃園市、臺中縣等)不對應,保留來源名稱,避免重複計算。
年中人口: raw 沒有直接提供,以「前一年底與當年底人口的平均」估算
  (2010 年直轄市合併:前一年底人口為原市+原縣之和)。 population_source 欄標明。
粗出生率: 原始檔有提供就用原值(rate_source=原始),沒有才用 出生數/年中人口*1000(計算)。
缺值:原始檔的「－」「…」「-」等一律留空(不是 0),並在 status 標註。
"""
from __future__ import annotations

import re
import sys
from datetime import date
from pathlib import Path

import pandas as pd
import xlrd

ROOT = Path(__file__).resolve().parent.parent
RAW = ROOT / "raw"
OUT = ROOT / "data"
RIS = RAW / "ris_fastdownload"

F_DISTRICT = RIS / "鄉鎮市區出生死亡結婚離婚(按登記及發生)(96).xls"
F_AGE = RIS / "縣市出生按生母五齡組(按發生)(96).xls"
F_CBR = RIS / "縣市育齡婦女生育率(按發生)(96).xls"
F_POP = RIS / "縣市人口按性別及五齡組(63).xls"
DIR_ODRP013 = RAW / "ris_api" / "ODRP013"
DIR_TOWN_POP = RAW / "ris_town_pop"  # 鄉鎮戶數及人口數,民國97-107年12月底

MISSING_TOKENS = {"－", "—", "-", "…", "...", "…。", "x", "X", "", "―"}
COUNTY_NAMES = {"臺北縣", "新北市", "臺北市", "基隆市", "宜蘭縣", "桃園縣", "桃園市", "新竹縣", "新竹市", "苗栗縣", "臺中縣",
                "臺中市", "彰化縣", "南投縣", "雲林縣", "嘉義縣", "嘉義市", "臺南縣", "臺南市", "高雄縣", "高雄市", "屏東縣",
                "臺東縣", "花蓮縣", "澎湖縣", "金門縣", "連江縣"}
AGGREGATES = {"臺灣省", "福建省", "臺灣地區", "臺閩地區", "總計"}
MUNICIPALITIES = {"臺北市", "新北市", "桃園市", "臺中市", "臺南市", "高雄市"}
RENAME_FOR_POP = {"臺北縣": "新北市", "桃園縣": "桃園市"}  # 只用於人口對接,不改輸出名稱
MERGE_2010 = {"臺中市": ["臺中縣"], "臺南市": ["臺南縣"], "高雄市": ["高雄縣"]}  # 民國99年合併
NTPC_NAMES = {"臺北縣", "新北市"}
FIRST_YEAR_ROC = 96  # 按發生資料起點
TODAY_YEAR = date.today().year

STATUS_OK = "正常"


def clean(s) -> str:
    return re.sub(r"[\s　\xa0]+", "", str(s))


def to_ad(roc: int) -> int:
    return roc + 1911


def num(v):
    """回傳 (值, 是否缺值)。「－」「…」空白都視為缺值,不轉成 0。"""
    if isinstance(v, (int, float)):
        return float(v), False
    if isinstance(v, str):
        t = v.strip()
        if t in MISSING_TOKENS:
            return None, True
        try:
            return float(t.replace(",", "")), False
        except ValueError:
            return None, True
    return None, True


def sheet_rows(path: Path, name: str):
    wb = xlrd.open_workbook(str(path), encoding_override="cp950")
    s = wb.sheet_by_name(name)
    return [[s.cell_value(r, c) for c in range(s.ncols)] for r in range(s.nrows)]


def sheet_names(path: Path):
    return xlrd.open_workbook(str(path), encoding_override="cp950").sheet_names()


def roc_sheets(path: Path):
    return sorted(int(n) for n in sheet_names(path) if n.isdigit())


def usable(roc: int) -> bool:
    return roc >= FIRST_YEAR_ROC and to_ad(roc) < TODAY_YEAR


def map_area(name: str) -> str:
    return "新北市" if name == "臺北縣" else name


def area_type(name_at_time: str) -> str:
    return "直轄市" if name_at_time in MUNICIPALITIES else "縣市"


# ------------------------------------------------------------------ 人口(年底)
def load_population():
    """pop[(roc, area_printed_clean)] = {'總計':..,'男':..,'女':..}  年底人口。"""
    pop = {}
    for roc in roc_sheets(F_POP):
        rows = sheet_rows(F_POP, str(roc))
        for i, r in enumerate(rows):
            if clean(r[1]) == "計" and i + 2 < len(rows) and clean(rows[i + 1][1]) == "男" and clean(rows[i + 2][1]) == "女":
                name = clean(rows[i + 1][0])
                vals = {}
                for lab, rr in (("總計", r), ("男", rows[i + 1]), ("女", rows[i + 2])):
                    v, miss = num(rr[2])
                    vals[lab] = v
                pop[(roc, name)] = vals
    return pop


def year_end_pop(pop, roc, area, sex, merge=False):
    key = RENAME_FOR_POP.get(area, area)
    names = [key] + (MERGE_2010.get(key, []) if merge else [])
    total = 0.0
    for n in names:
        cands = [n] + [k for k, v in RENAME_FOR_POP.items() if v == n]
        got = None
        for c in cands:
            if (roc, c) in pop and pop[(roc, c)][sex] is not None:
                got = pop[(roc, c)][sex]
                break
        if got is None:
            return None
        total += got
    return total


def mid_year_pop(pop, roc, area, sex):
    cur = year_end_pop(pop, roc, area, sex)
    prev = year_end_pop(pop, roc - 1, area, sex, merge=(roc == 99))
    if cur is None or prev is None:
        return None
    return (cur + prev) / 2


# ------------------------------------------------------------------ 粗出生率原值
def load_cbr():
    cbr = {}
    for roc in roc_sheets(F_CBR):
        for r in sheet_rows(F_CBR, str(roc)):
            name = clean(r[0])
            if not name or name in ("區域別",):
                continue
            v, miss = num(r[1])
            if v is not None:
                cbr.setdefault((roc, name), v)  # 只取第一次出現(縣市摘要區)
    return cbr


# ------------------------------------------------------------------ births_city
def build_city(pop, cbr):
    out = []
    for roc in roc_sheets(F_DISTRICT):
        if not usable(roc):
            continue
        rows = sheet_rows(F_DISTRICT, str(roc))
        seen = set()
        for r in rows:
            name = clean(r[10])  # 按發生日期統計的區域別欄
            if not name or name in seen or name in ("區域別",):
                continue
            if name in AGGREGATES and name != "總計":
                continue
            # 只取縣市摘要區:名稱必須存在於當年生育率表(縣市層級)
            if name != "總計" and (roc, name) not in cbr:
                continue
            seen.add(name)
            is_nat = name == "總計"
            area = "全國" if is_nat else map_area(name)
            at_time = "總計" if is_nat else name
            a_type = "全國" if is_nat else area_type(name)
            for sex, col in (("總計", 11), ("男", 12), ("女", 13)):
                births, miss = num(r[col])
                status = STATUS_OK
                if miss:
                    status = f"出生數缺值(原始檔為「{r[col]}」)" if str(r[col]).strip() else "出生數缺值(原始檔空白)"
                mid = None if is_nat else mid_year_pop(pop, roc, name, sex)
                if is_nat:
                    mid = mid_year_pop(pop, roc, "總計", sex)
                rate, src = None, ""
                if sex == "總計" and (roc, name) in cbr:
                    rate, src = cbr[(roc, name)], "原始"
                elif births is not None and mid:
                    rate, src = births / mid * 1000, "計算"
                else:
                    if status == STATUS_OK:
                        status = "缺年中人口,粗出生率未計算" if births is not None else status
                out.append(dict(year=to_ad(roc), area=area, name_at_time=at_time, area_type=a_type, sex=sex,
                                births=births, mid_year_population=mid,
                                population_source="年底人口平均(估)" if mid else "",
                                crude_birth_rate=rate, rate_source=src, status=status))
    return pd.DataFrame(out)


# ------------------------------------------------------------------ births_district
def district_name(n: str) -> str:
    return re.sub(r"[市鎮鄉區]$", "區", clean(n).lstrip("※"))


def load_odrp013():
    """Dec snapshot: pop[(roc, 區)] = 人口(里加總)。"""
    import json
    tot: dict = {}
    for f in sorted(DIR_ODRP013.glob("*_p*.json")):
        ym = f.name.split("_")[0]
        if not ym.endswith("12"):
            continue
        for rec in json.loads(f.read_text(encoding="utf-8"))["responseData"]:
            site = rec["site_id"]
            if not site.startswith("新北市"):
                continue
            v, miss = num(rec["people_total"])
            if v is None:
                continue
            key = (int(ym[:3]), site[3:])
            tot[key] = tot.get(key, 0.0) + v
    return tot


def load_town_pop_xls():
    """民國97-107年12月底各區人口(戶政司「鄉鎮戶數及人口數」,每縣市一張工作表)。"""
    tot = {}
    for f in sorted(DIR_TOWN_POP.glob("*年12月.xls")):
        roc = int(re.search(r"-(\d+)年12月", f.name).group(1))
        wb = xlrd.open_workbook(str(f), encoding_override="cp950")
        name = next(n for n in wb.sheet_names() if n in NTPC_NAMES)
        sh = wb.sheet_by_name(name)
        for r in range(4, sh.nrows):  # 第3列為新北市(臺北縣)合計,不納入
            d = clean(sh.cell_value(r, 0))
            if not d or d.startswith("說明"):
                continue
            v, miss = num(sh.cell_value(r, 2))
            if v is not None:
                tot[(roc, district_name(d))] = v
    return tot


def load_district_pop():
    pop = load_odrp013()          # 民國107-114年
    xls = load_town_pop_xls()     # 民國97-107年(107年兩來源重疊,以 xls 為準並檢核)
    for k, v in xls.items():
        if k in pop and abs(pop[k] - v) > 0.5:
            print(f"警告:{k} 兩來源人口不一致 xls={v} api={pop[k]}", file=sys.stderr)
    pop.update(xls)
    return pop


def build_district():
    dpop = load_district_pop()
    births = {}
    for roc in roc_sheets(F_DISTRICT):
        if not usable(roc):
            continue
        cur_county = None
        for r in sheet_rows(F_DISTRICT, str(roc)):
            name = clean(r[10])
            if not name or name == "區域別" or name in AGGREGATES:
                continue
            if name in COUNTY_NAMES:
                cur_county = name
                continue
            if cur_county in NTPC_NAMES:
                v, miss = num(r[11])
                births.setdefault((roc, district_name(name)), (name.lstrip("※"), v, miss, r[11]))
    out = []
    for (roc, d), (at_time, nb, bmiss, braw) in sorted(births.items()):
        cur = dpop.get((roc, d))
        prv = dpop.get((roc - 1, d))
        mid = (cur + prv) / 2 if cur is not None and prv is not None else None
        status = STATUS_OK
        if bmiss:
            status = f"出生數缺值(原始檔為「{braw}」)"
        rate, src = None, ""
        if nb is not None and mid:
            rate, src = nb / mid * 1000, "計算"
        elif status == STATUS_OK:
            status = "缺年中人口(需前一年底人口,區級人口自民國97年12月起),粗出生率未計算"
        out.append(dict(year=to_ad(roc), district=d, name_at_time=at_time, births=nb,
                        mid_year_population=mid, population_source="年底人口平均(估)" if mid else "",
                        crude_birth_rate=rate, rate_source=src, status=status))
    return pd.DataFrame(out)


# ------------------------------------------------------------------ births_mother_age
AGE_LABELS = [("-20", "20歲以下"), ("20～24", "20-24歲"), ("25～29", "25-29歲"), ("30～34", "30-34歲"),
              ("35～39", "35-39歲"), ("40～44", "40-44歲"), ("45+", "45歲以上")]


def norm_age_hdr(v) -> str:
    if isinstance(v, (int, float)):
        return str(int(v)) if float(v).is_integer() else str(v)
    return clean(v).replace("－", "-").replace("~", "～")


def build_mother_age():
    out = []
    for roc in roc_sheets(F_AGE):
        if not usable(roc):
            continue
        rows = sheet_rows(F_AGE, str(roc))
        cols = {}
        for r in rows[:12]:
            hdr = {norm_age_hdr(v): c for c, v in enumerate(r) if v != ""}
            if "20～24" in hdr and "25～29" in hdr:
                cols = {lab: hdr.get(src) for src, lab in AGE_LABELS if src in hdr}
                u20 = [c for v, c in ((norm_age_hdr(x), i) for i, x in enumerate(r)) if v in ("20", "-20", "0")]
                # 「-20」在部分年度被存成負數 -20.0 → norm 後為 "-20"
                if "20歲以下" not in cols and u20:
                    cols["20歲以下"] = u20[0]
                break
        assert len(cols) == 7, f"民國{roc}年找不到生母年齡欄位: {cols}"
        row = None
        for r in rows:
            n = clean(r[0])
            if n in NTPC_NAMES:
                row = r
                break
        assert row is not None, f"民國{roc}年找不到新北市/臺北縣"
        total, _ = num(row[1])
        vals = {}
        for _, lab in AGE_LABELS:
            v, miss = num(row[cols[lab]])
            vals[lab] = (v, miss, row[cols[lab]])
        known = [v for v, m, _ in vals.values() if v is not None]
        denom = sum(known)
        any_missing = any(m for _, m, _ in vals.values())
        for _, lab in AGE_LABELS:
            v, miss, raw = vals[lab]
            status = STATUS_OK
            if miss:
                status = f"出生數缺值(原始檔為「{raw}」)"
            elif any_missing:
                status = "同年有其他年齡組缺值,占比分母不完整"
            elif total is not None and abs(denom - total) > 0.5:
                status = f"各年齡組合計({denom:.0f})與總數({total:.0f})不符"
            out.append(dict(year=to_ad(roc), age_group=lab, births=v,
                            share=(v / denom if v is not None and denom else None), status=status))
    return pd.DataFrame(out)


def write(df: pd.DataFrame, name: str):
    OUT.mkdir(exist_ok=True)
    p = OUT / name
    df.to_csv(p, index=False, encoding="utf-8-sig")
    print(f"wrote {p.relative_to(ROOT)}  rows={len(df)}")


def main():
    pop = load_population()
    cbr = load_cbr()
    city = build_city(pop, cbr)
    cols_city = ["year", "area", "name_at_time", "area_type", "sex", "births", "mid_year_population",
                 "population_source", "crude_birth_rate", "rate_source", "status"]
    write(city[cols_city], "births_city.csv")
    dist = build_district()
    cols_d = ["year", "district", "name_at_time", "births", "mid_year_population", "population_source",
              "crude_birth_rate", "rate_source", "status"]
    write(dist[cols_d], "births_district.csv")
    age = build_mother_age()
    write(age[["year", "age_group", "births", "share", "status"]], "births_mother_age.csv")
    return city, dist, age


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main()
