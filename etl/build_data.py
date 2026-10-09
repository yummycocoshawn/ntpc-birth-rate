# /// script
# requires-python = ">=3.11"
# dependencies = ["pandas"]
# ///
"""data/*.csv -> docs/data.js(window.DATA = {...},缺值一律輸出 null,前端不會畫成 0)。"""
from __future__ import annotations

import json
import sys
from datetime import date
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
DATA = ROOT / "data"
OUT = ROOT / "docs" / "data.js"

# 對照對象:全國與其他五都。桃園在 2014 年由桃園縣升格為桃園市,圖上接成同一條線(atTime 保留當時名稱)。
COMPARE = ["全國", "臺北市", "桃園市", "臺中市", "臺南市", "高雄市"]
CITY_ALIASES = {"桃園縣": "桃園市"}
MAIN = "新北市"


def val(x, nd=4):
    if pd.isna(x):
        return None
    return round(float(x), nd)


def main() -> None:
    city = pd.read_csv(DATA / "births_city.csv", encoding="utf-8-sig")
    dist = pd.read_csv(DATA / "births_district.csv", encoding="utf-8-sig")
    age = pd.read_csv(DATA / "births_mother_age.csv", encoding="utf-8-sig")

    years = sorted(int(y) for y in city.year.unique())
    assert years == list(range(years[0], years[-1] + 1)), "年份不連續"
    idx = {y: i for i, y in enumerate(years)}
    n = len(years)

    city = city.assign(key=city.area.replace(CITY_ALIASES))
    city_out = {}
    for name in [MAIN] + COMPARE:
        g = city[city.key == name]
        series = {k: [None] * n for k in ("births", "male", "female", "rate", "pop")}
        series["rateSrc"] = [""] * n
        series["atTime"] = [""] * n
        for _, r in g.iterrows():
            i = idx[int(r.year)]
            if r.sex == "總計":
                series["births"][i] = val(r.births, 0)
                series["rate"][i] = val(r.crude_birth_rate)
                series["pop"][i] = val(r.mid_year_population, 1)
                series["rateSrc"][i] = "" if pd.isna(r.rate_source) else r.rate_source
                series["atTime"][i] = r.name_at_time if r.name_at_time != "總計" else "全國"
            elif r.sex == "男":
                series["male"][i] = val(r.births, 0)
            elif r.sex == "女":
                series["female"][i] = val(r.births, 0)
        city_out[name] = series

    dist_out = {}
    for d, g in dist.groupby("district"):
        s = {k: [None] * n for k in ("births", "rate", "pop")}
        s["atTime"] = [""] * n
        for _, r in g.iterrows():
            i = idx[int(r.year)]
            s["births"][i] = val(r.births, 0)
            s["rate"][i] = val(r.crude_birth_rate)
            s["pop"][i] = val(r.mid_year_population, 1)
            s["atTime"][i] = r.name_at_time
        dist_out[d] = s
    # 以最近一年出生數由多到少排序,方便選單閱讀
    order = sorted(dist_out, key=lambda d: -(next((v for v in reversed(dist_out[d]["births"]) if v is not None), 0)))

    groups = list(dict.fromkeys(age.age_group))
    age_out = {}
    for y, g in age.groupby("year"):
        g = g.set_index("age_group")
        age_out[str(int(y))] = {
            "births": [val(g.births.get(k), 0) for k in groups],
            "share": [val(g.share.get(k), 6) for k in groups],
        }

    data = {
        "meta": {
            "years": years, "nYears": n, "first": years[0], "last": years[-1],
            "built": date.today().isoformat(), "basis": "按發生日期",
            "main": MAIN, "compare": COMPARE, "reformYear": 2010,
        },
        "city": city_out,
        "districts": order,
        "district": dist_out,
        "ageGroups": groups,
        "age": age_out,
    }
    OUT.parent.mkdir(exist_ok=True)
    OUT.write_text("window.DATA = " + json.dumps(data, ensure_ascii=False, separators=(",", ":")) + ";\n", encoding="utf-8")
    print(f"wrote {OUT.relative_to(ROOT)}  {OUT.stat().st_size / 1024:.1f} KB  years={years[0]}-{years[-1]} (N={n})")


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main()
