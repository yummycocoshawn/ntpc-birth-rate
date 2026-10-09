#!/bin/bash
# Raw download of RIS open-data API pages, unmodified. Usage: download_ris_api.sh ENDPOINT PERIOD
# Saves raw/ris_api/<ENDPOINT>/<PERIOD>_p<N>.json
EP=$1; Y=$2
OUT=/c/Users/wudin/ntpc-birth-etl/raw/ris_api/$EP; mkdir -p "$OUT"
URL="https://www.ris.gov.tw/rs-opendata/api/v1/datastore/$EP/$Y"
f="$OUT/${Y}_p1.json"
[ -s "$f" ] || curl -s --retry 3 -m 120 "$URL?PAGE=1" -o "$f"
tp=$(grep -o '"totalPage":"[0-9]*"' "$f" | grep -o '[0-9]*')
[ -z "$tp" ] && { echo "$EP $Y: no data"; rm -f "$f"; exit 0; }
for ((p=2;p<=tp;p++)); do
  g="$OUT/${Y}_p$p.json"
  [ -s "$g" ] || curl -s --retry 3 -m 120 "$URL?PAGE=$p" -o "$g"
done
echo "$EP $Y: $tp pages"
