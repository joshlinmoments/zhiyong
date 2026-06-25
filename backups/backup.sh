#!/usr/bin/env bash
# 智勇網站一鍵備份：抓取 Firebase 全部內容，存成完整 JSON + 可讀文字檔。
# 用法：bash backups/backup.sh
set -euo pipefail

BASE="https://zhiyong-71ee9-default-rtdb.asia-southeast1.firebasedatabase.app"
DIR="$(cd "$(dirname "$0")" && pwd)"
TS="$(date +%Y%m%d_%H%M%S)"

echo "備份時間：$TS"
curl -s "$BASE/content.json"     -o "$DIR/content_$TS.json"     -w "content     HTTP %{http_code}  %{size_download} bytes\n"
curl -s "$BASE/annotations.json" -o "$DIR/annotations_$TS.json" -w "annotations HTTP %{http_code}  %{size_download} bytes\n"

python3 - "$DIR" "$TS" <<'PYEOF'
import json, sys
DIR, ts = sys.argv[1], sys.argv[2]
c = json.load(open(f"{DIR}/content_{ts}.json")) or {}
blocks = c.get("blocks", {}) or {}
anns = json.load(open(f"{DIR}/annotations_{ts}.json")) or {}

items = sorted(blocks.items(), key=lambda kv: (kv[1].get("order") or 0))
nch  = sum(1 for _, v in items if v.get("level") == "章")
nsec = sum(1 for _, v in items if v.get("level") == "節")
npb  = sum(1 for _, v in items if v.get("level") == "換頁")
ncon = sum(1 for _, v in items if v.get("level") in ("內文", "引用"))

out = ["# 智勇人才班教材 — 內容備份",
       f"備份時間：{ts}",
       f"統計：章 {nch}　節 {nsec}　內文/引用 {ncon}　換頁 {npb}　註解 {len(anns)}", ""]
for k, v in items:
    lv = v.get("level"); t = (v.get("text") or "").rstrip()
    if lv == "章":   out.append(f"\n\n# {t}")
    elif lv == "節": out.append(f"\n## {t}")
    elif lv == "換頁": out.append("\n———（換頁）———")
    elif lv == "引用": out.append("> " + t.replace("\n", "\n> "))
    else: out.append(t)
    out.append("")

path = f"{DIR}/智勇教材_文字備份_{ts}.md"
open(path, "w", encoding="utf-8").write("\n".join(out))
print(f"已輸出可讀文字檔：{path}")
print(f"統計 → 章 {nch}  節 {nsec}  內文/引用 {ncon}  換頁 {npb}  註解 {len(anns)}  區塊總數 {len(blocks)}")
PYEOF

echo "完成。備份檔位於：$DIR"
