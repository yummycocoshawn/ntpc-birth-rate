# /// script
# requires-python = ">=3.11"
# dependencies = ["pandas", "xlrd"]
# ///
"""驗證 data/ 的 CSV。原始檔由本檔自行重新讀取(不 import etl_births),避免兩邊共用同一個 bug。
全部通過 exit 0;任何一項失敗印出是哪幾筆並 exit 1。
"""
from __future__ import annotations

import re
import sys
from datetime import date
from pathlib import Path

import pandas as pd
import xlrd

ROOT = Path(__file__).resolve().parent.parent
DATA = ROOT / "data"
RIS = ROOT / "raw" / "ris_fastdownload"
F_DIST = RIS / "鄉鎮市區出生死亡結婚離婚(按登記及發生)(96).xls"
F_AGE = RIS / "縣市出生按生母五齡組(按發生)(96).xls"
F_CBR = RIS / "縣市育齡婦女生育率(按發生)(96).xls"

COUNTIES = {"臺北縣", "新北市", "臺北市", "基隆市", "宜蘭縣", "桃園縣", "桃園市", "新竹縣", "新竹市", "苗栗縣", "臺中縣", "臺中市",
            "彰化縣", "南投縣", "雲林縣", "嘉義縣", "嘉義市", "臺南縣", "臺南市", "高雄縣", "高雄市", "屏東縣", "臺東縣",
            "花蓮縣", "澎湖縣", "金門縣", "連江縣"}
NTPC = {"臺北縣", "新北市"}
FIRST_AD, LAST_AD = 2007, date.today().year - 1
RATE_TOL = 0.05
EXPECTED_DISTRICTS = 29
AGE_GROUPS = ["20歲以下", "20-24歲", "25-29歲", "30-34歲", "35-39歲", "40-44歲", "45歲以上"]

failures: list[str] = []


def clean(s) -> str:
    return re.sub(r"[\s　\xa0]+", "", str(s))


def rows_of(path: Path, sheet: str):
    s = xlrd.open_workbook(str(path), encoding_override="cp950").sheet_by_name(sheet)
    return [[s.cell_value(r, c) for c in range(s.ncols)] for r in range(s.nrows)]


def roc_years(path: Path):
    wb = xlrd.open_workbook(str(path), encoding_override="cp950")
    return sorted(int(n) for n in wb.sheet_names() if n.isdigit() and int(n) + 1911 <= LAST_AD and int(n) + 1911 >= FIRST_AD)


def is_county(n: str, ad: int) -> bool:
    # 「桃園市」在 2014 年以前是桃園縣下的縣轄市(區級),不是縣市層級
    return n in COUNTIES and not (n == "桃園市" and ad < 2014)


def isnum(v) -> bool:
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def check(name: str, bad: pd.DataFrame | list, note: str = ""):
    n = len(bad)
    if n == 0:
        print(f"[PASS] {name}")
        return
    print(f"[FAIL] {name}: {n} 筆 {note}")
    if isinstance(bad, pd.DataFrame):
        print(bad.head(30).to_string(index=False))
        if n > 30:
            print(f"  ...(另有 {n - 30} 筆)")
    else:
        for b in bad[:30]:
            print("  ", b)
    failures.append(name)


# ------------------------------------------------------------------ 原始檔
def raw_city_and_district():
    """city[(ad, name_clean, sex)] = 值(只取縣市摘要區第一次出現);dist[(ad, name)] = 值(新北市/臺北縣下各區)。"""
    city, dist = {}, {}
    for roc in roc_years(F_DIST):
        ad = roc + 1911
        seen, cur = set(), None
        for r in rows_of(F_DIST, str(roc)):
            n = clean(r[10])
            if not n or n == "區域別":
                continue
            if n == "總計" or is_county(n, ad):
                if n not in seen:  # 第一次出現 = 縣市摘要區
                    seen.add(n)
                    for sex, c in (("總計", 11), ("男", 12), ("女", 13)):
                        city[(ad, n, sex)] = r[c]
                cur = n if is_county(n, ad) else cur
                continue
            if n in ("臺灣省", "福建省", "臺灣地區", "臺閩地區"):
                continue
            if cur in NTPC:
                d = re.sub(r"[市鎮鄉區]$", "區", n.lstrip("※"))
                dist.setdefault((ad, d), r[11])
    return city, dist


def raw_age():
    """age[(ad, group)] = 值;total[ad] = 來源表新北市出生總數。"""
    labels = {"-20": "20歲以下", "20～24": "20-24歲", "25～29": "25-29歲", "30～34": "30-34歲",
              "35～39": "35-39歲", "40～44": "40-44歲", "45+": "45歲以上"}
    age, total = {}, {}
    for roc in roc_years(F_AGE):
        ad = roc + 1911
        rows = rows_of(F_AGE, str(roc))
        cols = {}
        for r in rows[:12]:
            h = {}
            for c, v in enumerate(r):
                k = str(int(v)) if isnum(v) and float(v).is_integer() else clean(v).replace("－", "-")
                h[k] = c
            if "20～24" in h and "25～29" in h:
                cols = {labels[k]: h[k] for k in labels if k in h}
                if "20歲以下" not in cols:
                    for k in ("-20", "20", "0"):
                        if k in h:
                            cols["20歲以下"] = h[k]
                            break
                break
        row = next(r for r in rows if clean(r[0]) in NTPC)
        total[ad] = row[1]
        for g, c in cols.items():
            age[(ad, g)] = row[c]
    return age, total


def raw_cbr():
    out = {}
    for roc in roc_years(F_CBR):
        for r in rows_of(F_CBR, str(roc)):
            n = clean(r[0])
            if n and (n == "總計" or n in COUNTIES) and (roc + 1911, n) not in out and isnum(r[1]):
                out[(roc + 1911, n)] = r[1]
    return out


def main() -> int:
    city = pd.read_csv(DATA / "births_city.csv", encoding="utf-8-sig")
    dist = pd.read_csv(DATA / "births_district.csv", encoding="utf-8-sig")
    age = pd.read_csv(DATA / "births_mother_age.csv", encoding="utf-8-sig")
    rcity, rdist = raw_city_and_district()
    rage, rage_total = raw_age()
    rcbr = raw_cbr()

    # 1. 有數值的筆數與總和 vs 原始檔 ------------------------------------------------
    print("== 1. 筆數與總和對照原始檔")
    c = city.assign(key_name=city.name_at_time.replace({"總計": "總計"}))
    for sex in ("總計", "男", "女"):
        cs = c[c.sex == sex]
        raw_vals = {k: v for k, v in rcity.items() if k[2] == sex and isnum(v)}
        csv_n, csv_sum = int(cs.births.notna().sum()), float(cs.births.sum())
        raw_n, raw_sum = len(raw_vals), float(sum(raw_vals.values()))
        # 原始檔含「臺灣省」之外全部縣市與總計;逐年比對
        bad = []
        for ad, g in cs.groupby("year"):
            rv = [v for k, v in raw_vals.items() if k[0] == ad]
            if int(g.births.notna().sum()) != len(rv) or abs(float(g.births.sum()) - sum(rv)) > 0.5:
                bad.append(dict(year=ad, sex=sex, csv_n=int(g.births.notna().sum()), raw_n=len(rv),
                                csv_sum=float(g.births.sum()), raw_sum=sum(rv)))
        check(f"city 出生數({sex}) 逐年筆數與總和一致 (CSV {csv_n}筆/{csv_sum:.0f} vs 原始 {raw_n}筆/{raw_sum:.0f})", pd.DataFrame(bad))
    raw_d = {k: v for k, v in rdist.items() if isnum(v)}
    bad = []
    for ad, g in dist.groupby("year"):
        rv = [v for k, v in raw_d.items() if k[0] == ad]
        if int(g.births.notna().sum()) != len(rv) or abs(float(g.births.sum()) - sum(rv)) > 0.5:
            bad.append(dict(year=ad, csv_n=int(g.births.notna().sum()), raw_n=len(rv), csv_sum=float(g.births.sum()), raw_sum=sum(rv)))
    check(f"district 出生數逐年筆數與總和一致 (CSV {int(dist.births.notna().sum())}筆 vs 原始 {len(raw_d)}筆)", pd.DataFrame(bad))
    raw_a = {k: v for k, v in rage.items() if isnum(v)}
    bad = []
    for ad, g in age.groupby("year"):
        rv = [v for k, v in raw_a.items() if k[0] == ad]
        if int(g.births.notna().sum()) != len(rv) or abs(float(g.births.sum()) - sum(rv)) > 0.5:
            bad.append(dict(year=ad, csv_n=int(g.births.notna().sum()), raw_n=len(rv), csv_sum=float(g.births.sum()), raw_sum=sum(rv)))
    check(f"mother_age 出生數逐年筆數與總和一致 (CSV {int(age.births.notna().sum())}筆 vs 原始 {len(raw_a)}筆)", pd.DataFrame(bad))

    # 2. 重複鍵 ----------------------------------------------------------------------
    print("== 2. 重複鍵")
    check("city (year, area, sex) 無重複", city[city.duplicated(["year", "area", "sex"], keep=False)])
    check("district (year, district) 無重複", dist[dist.duplicated(["year", "district"], keep=False)])
    check("mother_age (year, age_group) 無重複", age[age.duplicated(["year", "age_group"], keep=False)])

    # 3. 年份連續 ---------------------------------------------------------------------
    print("== 3. 年份連續、無缺漏")
    want = set(range(FIRST_AD, LAST_AD + 1))
    bad = []
    for name, df, key in (("city", city, None), ("district", dist, None), ("mother_age", age, None)):
        miss = sorted(want - set(df.year))
        extra = sorted(set(df.year) - want)
        if miss or extra:
            bad.append(f"{name}: 缺 {miss} 多 {extra}")
    for area, g in city.groupby("name_at_time"):  # 以當時名稱:每個名稱自己的區間內不可斷年
        ys = sorted(set(g.year))
        gaps = sorted(set(range(ys[0], ys[-1] + 1)) - set(ys))
        if gaps:
            bad.append(f"city {area}: {ys[0]}-{ys[-1]} 中斷 {gaps}")
    for sex, g in city.groupby("sex"):
        pass
    cnt = dist.groupby("year").district.nunique()
    bad += [f"district {y}: 行政區 {n} 個(應為 {EXPECTED_DISTRICTS})" for y, n in cnt.items() if n != EXPECTED_DISTRICTS]
    cnt = age.groupby("year").age_group.nunique()
    bad += [f"mother_age {y}: 年齡組 {n} 個(應為 {len(AGE_GROUPS)})" for y, n in cnt.items() if n != len(AGE_GROUPS)]
    bad += [f"mother_age 年齡組名稱異常 {sorted(set(age.age_group) ^ set(AGE_GROUPS))}"] if set(age.age_group) != set(AGE_GROUPS) else []
    check("各表年份連續且完整", bad)

    # 4. 男 + 女 = 總計 -----------------------------------------------------------------
    print("== 4. 男嬰 + 女嬰 = 總計")
    p = city.pivot_table(index=["year", "area", "name_at_time"], columns="sex", values="births", aggfunc="first", dropna=False).reset_index()
    p["diff"] = p["男"] + p["女"] - p["總計"]
    check("city 男+女=總計", p[(p["diff"].abs() > 0.5) | p["diff"].isna() & p["總計"].notna()][["year", "area", "name_at_time", "男", "女", "總計", "diff"]])

    # 5. 29 區加總 = 新北市 ----------------------------------------------------------------
    print("== 5. 29 區出生數加總 vs 新北市全市(逐年差異)")
    ds = dist.groupby("year").births.sum()
    nt = city[(city.area == "新北市") & (city.sex == "總計")].set_index("year").births
    cmp = pd.DataFrame({"區加總": ds, "新北市全市": nt})
    cmp["差異"] = cmp["區加總"] - cmp["新北市全市"]
    print(cmp.to_string())
    check("29 區加總 = 新北市全市", cmp[cmp["差異"].abs() > 0.5].reset_index().rename(columns={"index": "year"}))
    chk_raw = pd.Series({ad: sum(v for (a, _), v in raw_d.items() if a == ad) for ad in sorted({k[0] for k in raw_d})})
    check("原始檔 29 區加總 = 原始檔新北市全市",
          pd.DataFrame([dict(year=ad, raw_區加總=chk_raw[ad], raw_全市=rcity[(ad, n, "總計")]) for ad in chk_raw.index
                        for n in NTPC if (ad, n, "總計") in rcity and abs(chk_raw[ad] - rcity[(ad, n, "總計")]) > 0.5]))

    # 6. 自算粗出生率 vs 官方 ---------------------------------------------------------------
    print(f"== 6. 自行計算的粗出生率 vs 官方公布值(容許 {RATE_TOL}‰)")
    tot = city[(city.sex == "總計") & city.mid_year_population.notna() & city.births.notna()].copy()
    tot["official"] = [rcbr.get((y, n)) for y, n in zip(tot.year, tot.name_at_time.replace({"總計": "總計"}))]
    tot["own"] = tot.births / tot.mid_year_population * 1000
    tot["gap"] = (tot.own - tot.official).abs()
    tot = tot[tot.official.notna()]
    print(f"   比對 {len(tot)} 筆,最大差距 {tot.gap.max():.4f}‰,平均 {tot.gap.mean():.4f}‰")
    check("自算粗出生率與官方值差距 <= 0.05‰", tot[tot.gap > RATE_TOL][["year", "area", "name_at_time", "births", "mid_year_population", "own", "official", "gap"]].sort_values("gap", ascending=False))
    mism = city[(city.rate_source == "原始") & city.crude_birth_rate.notna()].copy()
    mism["official"] = [rcbr.get((y, n)) for y, n in zip(mism.year, mism.name_at_time)]
    check("CSV 的「原始」粗出生率等於原始檔數值", mism[(mism.official.isna()) | ((mism.crude_birth_rate - mism.official).abs() > 1e-9)][["year", "area", "crude_birth_rate", "official"]])
    # 各區加總算出的新北市粗出生率 vs 官方(區級沒有官方值,以全市官方值作對照)
    dd = dist.dropna(subset=["mid_year_population", "births"]).groupby("year").agg(b=("births", "sum"), p=("mid_year_population", "sum"))
    dd["own"] = dd.b / dd.p * 1000
    dd["official"] = [rcbr.get((y, n)) for y in dd.index for n in ["臺北縣" if y <= 2009 else "新北市"]]
    dd["gap"] = (dd.own - dd.official).abs()
    check("各區加總算出的新北市粗出生率與官方全市值差距 <= 0.05‰", dd[dd.gap > RATE_TOL].reset_index())

    # 7. 生母年齡組加總 = 當年出生數 -------------------------------------------------------
    print("== 7. 各生母年齡組加總 = 當年出生數")
    s = age.groupby("year").births.sum().rename("年齡組加總")
    cmp = pd.DataFrame({"年齡組加總": s, "新北市(city表)": nt, "原始檔總數": pd.Series(rage_total)})
    cmp["差(對city)"] = cmp["年齡組加總"] - cmp["新北市(city表)"]
    cmp["差(對原始)"] = cmp["年齡組加總"] - cmp["原始檔總數"]
    bad = cmp[(cmp["差(對city)"].abs() > 0.5) | (cmp["差(對原始)"].abs() > 0.5)].reset_index().rename(columns={"index": "year"})
    check("年齡組加總 = 當年出生數", bad)
    sh = age.groupby("year").share.sum()
    check("占比每年加總 = 1", pd.DataFrame({"year": sh[(sh - 1).abs() > 1e-9].index, "share_sum": sh[(sh - 1).abs() > 1e-9].values}))

    # 8. 粗出生率範圍 ---------------------------------------------------------------------
    print("== 8. 粗出生率落在 0 到 50‰")
    for name, df, cols in (("city", city, ["year", "area", "sex", "crude_birth_rate"]), ("district", dist, ["year", "district", "crude_birth_rate"])):
        r = df.crude_birth_rate
        check(f"{name} 粗出生率 0 <= r <= 50", df[r.notna() & ((r < 0) | (r > 50))][cols])

    print()
    if failures:
        print(f"驗證失敗 {len(failures)} 項:")
        for f in failures:
            print("  -", f)
        return 1
    print("全部驗證通過")
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    sys.exit(main())
