#!/usr/bin/env python3
"""Generate content.json for the website from the cleaned Word document.

Structure: front matter (motto / preface) + chapters -> sections -> units.
Units merge content the same way as the Excel: a colon lead-in starts a new
unit; classical citations (聖訓/云/曰/篇/經…) are flagged kind="quote".
Stable ids: chNN / chNN-sMM / chNN-sMM-uKK.
"""
import json
import re
import sys
from docx import Document

SRC = sys.argv[1]
OUT = sys.argv[2]

MOTTO = ["特別謹言慎行", "忍耐低心下氣", "注重和眾樂群", "勤儉克己復禮"]
PREFACE = (
    "古聖云：「國家興亡，匹夫有責。」道場興亡，人才亦有責，身為道中人才所擔當的是"
    "引領天下芸芸眾生，遠離塵緣迴轉，這也是我們神聖的使命。人才是道場的重要角色，"
    "是上方的好助手，也是後學們的好前賢。承上啟下、尊師重道、聽師調遣是修辦者應有"
    "的涵養。何其榮幸生逢三期末劫得道修道，更當把握千載難逢的佳機，救人救己善盡職責。"
    "因任重道遠必好學精進增益其所不能。孔子曰：工欲善其事，必先利其器。所以欲完成"
    "渡世的重責大任必好學，廣博自己的知識，精進自己品德的修養，篤厚自己信實的行為，"
    "精神奮發而不懈，作為道中的中流砥柱，成就萬八的聖業。"
)
PREFACE_SIGN = "黃世妍　謹識"
FRONT_H1 = {"樞紐四訓", "樞紐慈諭"}

re_cite = re.compile(r"[聖訓聖諭慈諭云曰篇經譚語詩偈錄傳銘賦頌書記論章句子].{0,8}[：:]$")


def is_leadin(s):
    return s.rstrip().endswith(("：", ":"))


def is_citation(text):
    first = text.split("\n", 1)[0].rstrip()
    return re_cite.search(first) is not None


doc = Document(SRC)

chapters = []
cur_ch = None
cur_sec = None
buf = []


def flush_unit():
    global buf
    if buf and cur_sec is not None:
        text = "\n".join(buf)
        uid = f"{cur_sec['id']}-u{len(cur_sec['units'])+1:02d}"
        cur_sec["units"].append({
            "id": uid,
            "text": text,
            "kind": "quote" if is_citation(text) else "prose",
        })
    buf = []


in_body = False
for p in doc.paragraphs:
    t = p.text.strip()
    if not t:
        continue
    style = p.style.name
    if style == "Title" or t.startswith("［請在 Word") or t in ("目　錄", "目錄"):
        continue
    if style == "Heading 1":
        flush_unit()
        if t in FRONT_H1:
            in_body = False
            cur_ch = cur_sec = None
            continue
        in_body = True
        cur_ch = {"id": f"ch{len(chapters)+1:02d}", "title": t, "sections": []}
        chapters.append(cur_ch)
        cur_sec = None
    elif not in_body:
        continue
    elif style == "Heading 2":
        flush_unit()
        cur_sec = {"id": f"{cur_ch['id']}-s{len(cur_ch['sections'])+1:02d}",
                   "title": t, "units": []}
        cur_ch["sections"].append(cur_sec)
    else:
        # content paragraph (Normal / Quote) -> 內文 unit, colon lead-in delimits
        if cur_sec is None:  # content before any 節 -> implicit section
            flush_unit()
            cur_sec = {"id": f"{cur_ch['id']}-s{len(cur_ch['sections'])+1:02d}",
                       "title": "", "units": []}
            cur_ch["sections"].append(cur_sec)
        if buf and is_leadin(t):
            flush_unit()
        buf.append(t)
flush_unit()

data = {
    "title": "智勇人才班教材",
    "subtitle": "ZHI YONG ‧ 道學講堂",
    "frontMatter": {
        "motto": {"title": "樞紐四訓", "lines": MOTTO},
        "preface": {"title": "序文〈樞紐慈諭〉", "text": PREFACE, "sign": PREFACE_SIGN},
    },
    "chapters": chapters,
}

with open(OUT, "w", encoding="utf-8") as f:
    json.dump(data, f, ensure_ascii=False, indent=1)

nu = sum(len(s["units"]) for c in chapters for s in c["sections"])
ns = sum(len(c["sections"]) for c in chapters)
print(f"Saved {OUT}")
print(f"chapters={len(chapters)} sections={ns} units={nu}")
