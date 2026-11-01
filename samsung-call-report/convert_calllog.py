#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
تبدیل فایل پشتیبان تماس‌ها به گزارش اکسل (کاری / شخصی)
ورودی پشتیبانی‌شده:
  - فایل XML اپ «SMS Backup & Restore»  (calls-*.xml)
  - فایل CSV خروجی اپ‌های Call Log Backup
استفاده:
  python convert_calllog.py calls.xml
خروجی: دو فایل اکسل کنار همان فایل ورودی.
"""
import sys, os, re, csv, io
import xml.etree.ElementTree as ET
from datetime import datetime, timezone

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter
except ImportError:
    sys.exit("کتابخانه openpyxl نصب نیست. اجرا کنید:  pip install openpyxl")

# ---------- تنظیمات (در صورت نیاز ویرایش کنید) ----------
DAYS = 90
WORK_KEYWORDS = ["کار", "شرکت", "مشتری", "همکار", "دفتر", "اداره", "مهندس",
                 "work", "office", "client"]
WORK_NUMBERS = []          # شماره‌های کاری اضافی، مثل: ["09121234567"]
UNKNOWN_IS_WORK = False     # شماره‌های ذخیره‌نشده کاری حساب شوند؟

TYPE_FA = {1: "ورودی (مخاطب زنگ زده)", 2: "خروجی (شما زنگ زده‌اید)",
           3: "بی‌پاسخ (از دست رفته)", 4: "پیغام صوتی", 5: "رد شده",
           6: "مسدود شده", 7: "پاسخ در دستگاه دیگر"}
WEEK_FA = ["دوشنبه", "سه‌شنبه", "چهارشنبه", "پنجشنبه", "جمعه", "شنبه", "یکشنبه"]  # Mon=0

# ---------- تبدیل میلادی به شمسی ----------
def to_jalali(gy, gm, gd):
    g_d_m = [0,31,59,90,120,151,181,212,243,273,304,334]
    gy2 = gy-1600; gm2 = gm-1; gd2 = gd-1
    days = 365*gy2 + (gy2+3)//4 - (gy2+99)//100 + (gy2+399)//400
    days += g_d_m[gm2] + gd2
    if gm2>1 and ((gy%4==0 and gy%100!=0) or gy%400==0): days += 1
    days -= 79
    j_np = days//12053; days %= 12053
    jy = 979 + 33*j_np + 4*(days//1461); days %= 1461
    if days >= 366:
        jy += (days-1)//365; days = (days-1)%365
    if days < 186:
        jm = 1 + days//31; jd = 1 + days%31
    else:
        jm = 7 + (days-186)//30; jd = 1 + (days-186)%30
    return jy, jm, jd

def jdate(dt):
    jy, jm, jd = to_jalali(dt.year, dt.month, dt.day)
    return "%04d/%02d/%02d" % (jy, jm, jd)

def normalize(n):
    d = re.sub(r"\D", "", n or "")
    return d[-10:] if len(d) > 10 else d

def fmt_dur(s):
    return "%d:%02d:%02d" % (s//3600, (s%3600)//60, s%60)

# ---------- خواندن ورودی ----------
def parse_xml(path):
    calls = []
    root = ET.parse(path).getroot()
    for c in root.iter("call"):
        try:
            calls.append(dict(
                number=c.get("number","") or "",
                name=(c.get("contact_name") or "").strip(),
                date=int(c.get("date","0")),
                duration=int(c.get("duration","0")),
                type=int(c.get("type","0")),
            ))
        except (TypeError, ValueError):
            continue
    return calls

def parse_csv(path):
    calls = []
    with open(path, encoding="utf-8-sig", newline="") as f:
        rd = csv.DictReader(f)
        cols = {k.lower().strip(): k for k in (rd.fieldnames or [])}
        def col(*names):
            for n in names:
                if n in cols: return cols[n]
            return None
        cn, cname = col("number","phone number","phone"), col("name","contact name","contact_name")
        cdate, cdur, ctype = col("date","time","call date"), col("duration","duration (sec)","sec"), col("type","call type")
        for r in rd:
            raw = (r.get(cdate) or "").strip() if cdate else ""
            ts = None
            if raw.isdigit():
                ts = int(raw); ts = ts if ts > 10**12 else ts*1000
            else:
                for fmt in ("%Y-%m-%d %H:%M:%S","%d/%m/%Y %H:%M:%S","%m/%d/%Y %H:%M:%S","%Y/%m/%d %H:%M"):
                    try: ts = int(datetime.strptime(raw, fmt).timestamp()*1000); break
                    except ValueError: pass
            if ts is None: continue
            tp = (r.get(ctype) or "").strip().lower() if ctype else ""
            tmap = {"incoming":1,"1":1,"outgoing":2,"2":2,"missed":3,"3":3,"rejected":5,"5":5}
            calls.append(dict(
                number=(r.get(cn) or "") if cn else "",
                name=((r.get(cname) or "").strip()) if cname else "",
                date=ts,
                duration=int(re.sub(r"\D","", r.get(cdur) or "0") or 0) if cdur else 0,
                type=tmap.get(tp, 0),
            ))
    return calls

def classify(c):
    n = normalize(c["number"])
    if n and n in {normalize(x) for x in WORK_NUMBERS}: return "کاری"
    nm = (c["name"] or "")
    for k in WORK_KEYWORDS:
        if k and nm and k.lower() in nm.lower(): return "کاری"
    if not nm and UNKNOWN_IS_WORK: return "کاری"
    return "شخصی"

# ---------- ساخت اکسل ----------
HDR = ["ردیف","تاریخ (شمسی)","تاریخ (میلادی)","روز هفته","ساعت","نام مخاطب","شماره",
       "نوع تماس / تماس‌گیرنده","مدت (س:د:ث)","مدت (ثانیه)","دسته"]
WID = [7,14,14,11,9,26,18,26,14,12,9]
HEAD_FILL = PatternFill("solid", fgColor="2F5597")
HEAD_FONT = Font(bold=True, color="FFFFFF", name="Tahoma")

def style_header(ws, ncol):
    ws.sheet_view.rightToLeft = True
    for j in range(1, ncol+1):
        c = ws.cell(1, j); c.fill = HEAD_FILL; c.font = HEAD_FONT
        c.alignment = Alignment(horizontal="center", vertical="center")
    ws.freeze_panes = "A2"

def rows_for(calls):
    out = []
    for i, c in enumerate(sorted(calls, key=lambda x: -x["date"]), 1):
        dt = datetime.fromtimestamp(c["date"]/1000)
        who = TYPE_FA.get(c["type"], "نوع %d" % c["type"])
        out.append([i, jdate(dt), dt.strftime("%Y-%m-%d"), WEEK_FA[dt.weekday()],
                    dt.strftime("%H:%M"), c["name"] or "(ناشناس)", c["number"],
                    who, fmt_dur(c["duration"]), c["duration"], c["category"]])
    return out

def write_calls_sheet(ws, calls):
    ws.append(HDR)
    for r in rows_for(calls): ws.append(r)
    for j, w in enumerate(WID, 1): ws.column_dimensions[get_column_letter(j)].width = w
    style_header(ws, len(HDR))
    if calls: ws.auto_filter.ref = "A1:%s%d" % (get_column_letter(len(HDR)), len(calls)+1)

def contact_sheet(ws, calls):
    hdr = ["نام مخاطب","شماره","تعداد تماس","ورودی","خروجی","بی‌پاسخ/رد","مجموع مدت","مجموع (دقیقه)","آخرین تماس"]
    ws.append(hdr)
    groups = {}
    for c in calls:
        k = normalize(c["number"]) or c["number"]
        groups.setdefault(k, []).append(c)
    rows = []
    for k, g in groups.items():
        name = next((x["name"] for x in g if x["name"]), "(ناشناس)")
        tot = sum(x["duration"] for x in g)
        last = max(g, key=lambda x: x["date"]); ld = datetime.fromtimestamp(last["date"]/1000)
        rows.append([name, g[0]["number"], len(g),
                     sum(1 for x in g if x["type"]==1), sum(1 for x in g if x["type"]==2),
                     sum(1 for x in g if x["type"] in (3,5)), fmt_dur(tot), round(tot/60,1),
                     jdate(ld)+" "+ld.strftime("%H:%M"), tot])
    rows.sort(key=lambda r: -r[-1])
    for r in rows: ws.append(r[:-1])
    for j, w in enumerate([26,18,11,9,9,12,14,13,20], 1): ws.column_dimensions[get_column_letter(j)].width = w
    style_header(ws, len(hdr))
    if rows: ws.auto_filter.ref = "A1:%s%d" % (get_column_letter(len(hdr)), len(rows)+1)

def stats(calls):
    return dict(total=len(calls),
                inc=sum(1 for c in calls if c["type"]==1),
                out=sum(1 for c in calls if c["type"]==2),
                miss=sum(1 for c in calls if c["type"] in (3,5)),
                sec=sum(c["duration"] for c in calls),
                contacts=len({normalize(c["number"]) for c in calls}))

def main():
    if len(sys.argv) < 2:
        sys.exit("استفاده: python convert_calllog.py <فایل calls.xml یا .csv>")
    path = sys.argv[1]
    if not os.path.exists(path): sys.exit("فایل پیدا نشد: " + path)
    calls = parse_xml(path) if path.lower().endswith(".xml") else parse_csv(path)
    if not calls: sys.exit("هیچ تماسی در فایل خوانده نشد.")

    now = datetime.now(); cutoff = (now.timestamp() - DAYS*86400) * 1000
    calls = [c for c in calls if c["date"] >= cutoff]
    for c in calls: c["category"] = classify(c)
    work = [c for c in calls if c["category"]=="کاری"]
    personal = [c for c in calls if c["category"]=="شخصی"]
    ws_, from_dt = stats(work), datetime.fromtimestamp(cutoff/1000)
    rng = "از %s تا %s" % (jdate(from_dt), jdate(now))

    base = os.path.splitext(os.path.basename(path))[0]
    outdir = os.path.dirname(os.path.abspath(path))
    stamp = now.strftime("%Y-%m-%d")

    # فایل مدیر (فقط کاری)
    wb = openpyxl.Workbook(); s = wb.active; s.title = "خلاصه گزارش"; s.sheet_view.rightToLeft = True
    for r in [["بازه گزارش", rng],
              ["تعداد کل تماس‌های کاری", ws_["total"]],
              ["تماس‌های ورودی", ws_["inc"]], ["تماس‌های خروجی", ws_["out"]],
              ["بی‌پاسخ / رد شده", ws_["miss"]],
              ["مجموع زمان مکالمه (س:د:ث)", fmt_dur(ws_["sec"])],
              ["مجموع زمان مکالمه (دقیقه)", round(ws_["sec"]/60,1)],
              ["تعداد مخاطبین کاری", ws_["contacts"]],
              ["تاریخ تهیه گزارش", jdate(now)+" "+now.strftime("%H:%M")]]:
        s.append(r)
    s.column_dimensions["A"].width = 38; s.column_dimensions["B"].width = 30
    for i in range(1, 10): s.cell(i,1).font = Font(bold=True, name="Tahoma")
    write_calls_sheet(wb.create_sheet("تماس‌های کاری"), work)
    contact_sheet(wb.create_sheet("مخاطبین کاری"), work)
    work_file = os.path.join(outdir, "CallReport_Work_%s.xlsx" % stamp)
    wb.save(work_file)

    # فایل کامل (همه + شخصی)
    wb2 = openpyxl.Workbook(); write_calls_sheet(wb2.active, calls); wb2.active.title = "همه تماس‌ها"
    write_calls_sheet(wb2.create_sheet("تماس‌های شخصی"), personal)
    contact_sheet(wb2.create_sheet("همه مخاطبین"), calls)
    all_file = os.path.join(outdir, "CallReport_All_%s.xlsx" % stamp)
    wb2.save(all_file)

    print("✔ گزارش ساخته شد:")
    print("   مدیر (فقط کاری):", work_file)
    print("   کامل:", all_file)
    print("   تماس‌های ۹۰ روز اخیر: %d  (کاری: %d | شخصی: %d)" % (len(calls), len(work), len(personal)))

if __name__ == "__main__":
    main()
