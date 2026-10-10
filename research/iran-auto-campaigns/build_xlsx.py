# -*- coding: utf-8 -*-
"""Builds Iran_Auto_Campaigns_2025-2026.xlsx and evidence_library.csv from data.py"""
import csv, datetime as dt
from collections import Counter, defaultdict
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, Reference
from openpyxl.worksheet.table import Table, TableStyleInfo
import data as D

OUT = "Iran_Auto_Campaigns_2025-2026.xlsx"
HDR_FILL = PatternFill("solid", fgColor="1F3864"); HDR_FONT = Font(bold=True, color="FFFFFF", name="Tahoma", size=10)
BODY = Font(name="Tahoma", size=9); LINK = Font(name="Tahoma", size=9, color="0563C1", underline="single")
THIN = Side(style="thin", color="BFBFBF"); BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
WRAP = Alignment(wrap_text=True, vertical="top")
GROUP = {"A": "A — خودروساز داخلی", "B": "B — مونتاژی", "C": "C — وارداتی"}
brand_by_id = {b[0]: b for b in D.BRANDS}
co_by_id = {c[0]: c for c in D.COMPANIES}

def to_date(s):
    try: return dt.date.fromisoformat(s)
    except Exception: return None

def sheet(wb, title, headers, rows, widths=None, rtl=True, table=True, link_cols=()):
    ws = wb.create_sheet(title)
    ws.sheet_view.rightToLeft = rtl
    ws.append(headers)
    for r in rows: ws.append(list(r))
    for c in ws[1]:
        c.fill, c.font, c.alignment, c.border = HDR_FILL, HDR_FONT, Alignment(wrap_text=True, vertical="center", horizontal="center"), BORDER
    for row in ws.iter_rows(min_row=2):
        for c in row:
            c.font, c.alignment, c.border = BODY, WRAP, BORDER
            if isinstance(c.value, dt.date): c.number_format = "yyyy-mm-dd"
    for col in link_cols:
        for r in range(2, ws.max_row + 1):
            c = ws.cell(r, col)
            if isinstance(c.value, str) and c.value.startswith("http"):
                c.hyperlink = c.value; c.font = LINK
    ws.freeze_panes = "B2"
    for i, w in enumerate(widths or [], 1): ws.column_dimensions[get_column_letter(i)].width = w
    ws.row_dimensions[1].height = 36
    if table and ws.max_row > 1:
        ref = f"A1:{get_column_letter(ws.max_column)}{ws.max_row}"
        t = Table(displayName="T_" + title.split("_", 1)[1][:20].replace("_", ""), ref=ref)
        t.tableStyleInfo = TableStyleInfo(name="TableStyleLight9", showRowStripes=True)
        ws.add_table(t)  # tables carry autofilter
    return ws

camps = D.C
brand_camps = defaultdict(list)
for c in camps:
    for b in filter(None, c["brands"].split(";")): brand_camps[b].append(c)
ev_camps = defaultdict(list)
for c in camps:
    for e in c["ev"].split(";"): ev_camps[e].append(c["id"])
for b in D.BRANDS:
    for e in filter(None, b[8].split(";")): ev_camps[e].append(b[0])

def is_brand_campaign(c):  # counted as a documented brand campaign (not pure sales notice / ecosystem)
    return c["rt"] in (D.RT_CAMP, D.RT_EXPO)

FOUND = "بررسی انجام شده؛ کمپین پیدا شد"
ONLY_SALES = "بررسی انجام شده؛ کمپین تأییدشده پیدا نشد (فقط طرح فروش)"
def coverage_status(b):
    if b[9] == FOUND and not any(is_brand_campaign(c) for c in brand_camps.get(b[0], [])):
        return ONLY_SALES
    return b[9]

wb = Workbook(); wb.remove(wb.active)

# 01 master brands
rows = []
for b in D.BRANDS:
    cos = "؛ ".join(co_by_id[x][1] for x in b[4].split(";") if x in co_by_id) or D.U
    cs = brand_camps.get(b[0], [])
    rows.append([b[0], b[1], b[2], GROUP[b[3]], b[4], cos, b[5], b[6], b[7], b[8] or "—", coverage_status(b),
                 sum(is_brand_campaign(c) for c in cs), sum(c["rt"] == D.RT_SALES for c in cs)])
sheet(wb, "01_Master_Brands",
      ["Brand_ID", "برند", "کشور مبدأ", "دسته بازار", "Company_IDs", "شرکت(های) مرتبط", "عرضه‌کننده در ایران", "مدل‌های مرتبط",
       "وضعیت فعالیت در بازه", "شواهد فعالیت (Evidence_ID)", "وضعیت بررسی تبلیغات", "تعداد کمپین/رویداد مستند", "تعداد طرح فروش ثبت‌شده"],
      rows, [9, 34, 18, 18, 14, 28, 28, 34, 18, 22, 28, 12, 12])

# 01b companies
rows = []
for co in D.COMPANIES:
    linked = [b[0] for b in D.BRANDS if co[0] in b[4].split(";")]
    rows.append([co[0], co[1], co[2], co[3], GROUP.get(co[4], co[4]), ";".join(linked) or "—"])
sheet(wb, "01b_Companies", ["Company_ID", "نام شرکت", "English", "نقش", "دسته", "برندهای مرتبط"], rows, [11, 36, 30, 30, 18, 30])

# 02 campaigns
rows = []
for c in camps:
    bnames = "؛ ".join(brand_by_id[x][1] for x in filter(None, c["brands"].split(";"))) or "—"
    evs = c["ev"].split(";"); ev_by = {e[0]: e for e in D.EVIDENCE}
    first_url = ev_by[evs[0]][2]
    pubs = "؛ ".join(f"{e}: {ev_by[e][4]}" for e in evs)
    rows.append([c["id"], c["rt"], c["cat"], c["brands"] or "—", bnames, c["ex"], c["name"], c["slogan"], c["obj"], c["aud"], c["prod"],
                 to_date(c["start"]) or D.U, to_date(c["end"]) or D.U, c["jal"], pubs, c["geo"], c["ch"], c["creative"], c["fmt"], c["offer"],
                 c["agency"], c["ev"], first_url, D.U, f"WebSearch — {D.RUN_DATE}", c["ver"], c["conf"], c["lim"],
                 "داده عمومی قابل اتکا پیدا نشد"])
sheet(wb, "02_Campaigns",
      ["Campaign_ID", "نوع رکورد", "دسته کمپین", "Brand_IDs", "برند", "شرکت اجراکننده/سفارش‌دهنده", "نام/عنوان توصیفی", "شعار و پیام", "هدف بازاریابی",
       "مخاطب هدف", "محصول/مدل", "تاریخ شروع (میلادی)", "تاریخ پایان (میلادی)", "بازه اجرا (شمسی، طبق منبع)", "تاریخ انتشار شواهد (اولین/آخرین)",
       "جغرافیا", "کانال‌ها", "ایده خلاقه", "فرمت محتوا", "نوع پیشنهاد فروش", "آژانس/مجری", "Evidence_IDs", "لینک شاهد اصلی",
       "لینک تصویر/ویدئو", "منبع و تاریخ دسترسی", "وضعیت تأیید", "میزان اطمینان", "محدودیت‌های اطلاعاتی", "بودجه و عملکرد"],
      rows, [10, 20, 20, 14, 26, 22, 40, 28, 22, 22, 28, 12, 12, 22, 34, 14, 22, 28, 14, 24, 10, 16, 30, 10, 18, 20, 9, 32, 18], link_cols=(23,))

# 03 evidence
rows = []
for e in D.EVIDENCE:
    rows.append([e[0], e[1], e[2], e[3], e[4], to_date(D.RUN_DATE), e[5], ";".join(ev_camps.get(e[0], [])) or "زمینه‌ای (روش‌شناسی)", e[6]])
sheet(wb, "03_Evidence_Sources", ["Evidence_ID", "رسانه/منبع", "URL", "نوع منبع", "تاریخ انتشار (طبق نتیجه)", "تاریخ دسترسی", "ادعای پشتیبانی‌شده",
      "رکوردهای مرتبط", "روش دسترسی"], rows, [10, 20, 50, 18, 20, 12, 50, 22, 26], link_cols=(3,))

# 04 channel analysis (brand x channel) + timeline (brand x month)
CH = ["نمایشگاه", "رویداد", "سامانه فروش", "رسانه", "بیلبورد", "شوروم", "نمایندگی", "چهره", "رپورتاژ", "تست‌درایو"]
rows = []
for b in D.BRANDS:
    cs = brand_camps.get(b[0], [])
    if not cs: continue
    cnt = [sum(k in c["ch"] for c in cs) for k in CH]
    rows.append([b[0], b[1], GROUP[b[3]], len(cs)] + cnt + [sum(1 for x in cnt if x)])
for c in camps:
    if c["rt"] == D.RT_X:
        cnt = [int(k in c["ch"]) for k in CH]
        rows.append([c["id"], c["ex"], "X — اکوسیستم", 1] + cnt + [sum(cnt)])
ws4 = sheet(wb, "04_Channel_Analysis", ["ID", "برند/بازیگر", "دسته", "تعداد رکورد"] + CH + ["گستردگی کانال (تعداد کانال متمایز)"], rows,
            [10, 34, 18, 10] + [11] * len(CH) + [16])
# timeline matrix
months = [("2025-10", "مهر/آبان ۱۴۰۴"), ("2025-11", "آبان/آذر"), ("2025-12", "آذر/دی"), ("2026-01", "دی/بهمن"), ("2026-02", "بهمن/اسفند"),
          ("2026-03", "اسفند/فروردین"), ("2026-04", "فروردین/اردیبهشت"), ("2026-05", "اردیبهشت/خرداد"), ("2026-06", "خرداد/تیر"),
          ("2026-07", "تیر/مرداد"), ("2026-08", "مرداد/شهریور"), ("2026-09", "شهریور/مهر"), ("2026-10", "مهر ۱۴۰۵")]
JM = {"آبان ۱۴۰۴": "2025-11", "آذر ۱۴۰۴": "2025-12", "دی ۱۴۰۴": "2026-01", "بهمن ۱۴۰۴": "2026-02", "اسفند ۱۴۰۴": "2026-03",
      "اردیبهشت ۱۴۰۵": "2026-05", "خرداد ۱۴۰۵": "2026-06", "تیر ۱۴۰۵": "2026-07", "مرداد ۱۴۰۵": "2026-08", "شهریور ۱۴۰۵": "2026-09", "مهر ۱۴۰۵": "2026-10", "مهر ۱۴۰۴": "2025-10", "پاییز ۱۴۰۴": "2025-11"}
def month_of(c):
    d = to_date(c["start"])
    if d: return d.strftime("%Y-%m")
    for k, v in JM.items():
        if k in c["jal"]: return v
    return None
start = ws4.max_row + 3
ws4.cell(start, 1, "ماتریس برند × زمان (ماه شروع رکورد؛ بر اساس تاریخ میلادی یا ماه شمسی ذکرشده)").font = Font(bold=True, name="Tahoma")
ws4.cell(start + 1, 1, "Brand_ID").font = HDR_FONT; ws4.cell(start + 1, 1).fill = HDR_FILL
ws4.cell(start + 1, 2, "برند").font = HDR_FONT; ws4.cell(start + 1, 2).fill = HDR_FILL
for j, (m, lab) in enumerate(months):
    cl = ws4.cell(start + 1, 3 + j, f"{m}\n{lab}"); cl.font = HDR_FONT; cl.fill = HDR_FILL; cl.alignment = Alignment(wrap_text=True)
r = start + 2
for b in D.BRANDS:
    cs = brand_camps.get(b[0], [])
    if not cs: continue
    ws4.cell(r, 1, b[0]).font = BODY; ws4.cell(r, 2, b[1]).font = BODY
    mc = Counter(month_of(c) for c in cs)
    for j, (m, _) in enumerate(months):
        if mc.get(m): ws4.cell(r, 3 + j, mc[m]).font = BODY
    r += 1
ws4.cell(r + 1, 1, "نکته: رکوردهای با تاریخ نامشخص در ماتریس زمانی شمرده نشده‌اند.").font = BODY

# 05 creative analysis
rows = []
for c in camps:
    rows.append([c["id"], "؛ ".join(brand_by_id[x][1] for x in filter(None, c["brands"].split(";"))) or c["ex"], c["rt"], c["name"], c["slogan"],
                 c["creative"], c["offer"], c["aud"], c["cta"], c["funnel"], c["emph"], c["integ"],
                 "تحلیل (نه واقعیت اثبات‌شده) بر پایه شواهد ثانویه" if c["conf"] != "Low" else "شواهد ناکافی برای تحلیل خلاقه معتبر"])
sheet(wb, "05_Creative_Analysis", ["Campaign_ID", "برند/بازیگر", "نوع رکورد", "عنوان", "شعار/پیام", "ایده خلاقه", "پیشنهاد فروش", "مخاطب", "CTA",
      "مرحله قیف (تحلیل)", "تأکید پیام (تحلیل)", "یکپارچگی چندکاناله (تحلیل)", "یادداشت"], rows,
      [10, 24, 18, 36, 28, 32, 26, 22, 16, 16, 18, 18, 26])

# 06 agencies & partners
rows = [
 ["AG-001", "ایده تجارت هرمس", "برگزارکننده نمایشگاه", "هفتمین نمایشگاه خودرو تهران (۲۲-۲۶ شهریور ۱۴۰۵)",
  "CMP-011;CMP-028;CMP-029;CMP-031;CMP-033;CMP-038;CMP-039;CMP-040;CMP-041;CMP-042;CMP-043;CMP-044", "EV-003;EV-004", "Medium"],
 ["AG-002", "ایران‌مال", "محل برگزاری رویداد رونمایی", "رونمایی تیسان", "CMP-032", "EV-100", "Medium"],
 ["AG-003", "سیدجواد هاشمی، مجید واشقانی، جواد داوری", "چهره‌های حاضر در رونمایی (نوع قرارداد نامشخص)", "رونمایی تیسان", "CMP-032", "EV-100", "Low"],
 ["AG-004", "اتونوین / سامانه یکپارچه واردات", "پلتفرم فروش (دولتی)", "عرضه وارداتی‌ها", "CMP-045;CMP-046", "EV-070;EV-074", "Medium"],
 ["AG-005", "آژانس تبلیغاتی خلاقه/رسانه‌ای خودروسازان", "نامشخص", "برای هیچ‌یک از کمپین‌های خودروسازان، آژانس یا مجری خلاقه در منابع قابل دسترس شناسایی نشد", "—", "—", "—"],
]
sheet(wb, "06_Agencies_Partners", ["Partner_ID", "نام", "نقش", "پروژه/کمپین", "Campaign_IDs", "Evidence_IDs", "اطمینان"], rows, [10, 34, 30, 44, 40, 18, 10])

# 07 coverage
QUERIES = {"BR-001": 7, "BR-002": 1, "BR-003": 6, "BR-004": 1, "BR-006": 4, "BR-007": 5, "BR-008": 2, "BR-009": 3, "BR-010": 2, "BR-011": 3, "BR-012": 2,
           "BR-013": 3, "BR-014": 1, "BR-015": 1, "BR-016": 1, "BR-017": 2, "BR-018": 1, "BR-019": 2, "BR-020": 1, "BR-021": 2, "BR-022": 1, "BR-023": 1,
           "BR-024": 1, "BR-025": 2, "BR-026": 1, "BR-027": 1, "BR-028": 2, "BR-029": 1, "BR-030": 1, "BR-031": 1, "BR-032": 1, "BR-033": 1, "BR-034": 1,
           "BR-035": 1, "BR-036": 1, "BR-037": 1, "BR-038": 1, "BR-039": 1, "BR-040": 1, "BR-046": 1, "BR-047": 1}
rows = []
for b in D.BRANDS:
    nsrc = len({e for c in brand_camps.get(b[0], []) for e in c["ev"].split(";")} | set(filter(None, b[8].split(";"))))
    q = QUERIES.get(b[0], 0)
    gaps = "وب‌سایت رسمی، اینستاگرام/تلگرام/آپارات و کتابخانه‌های تبلیغاتی بررسی نشدند (عدم دسترسی)؛ شواهد خلاقه (تصویر/ویدئو) و آژانس یافت نشد"
    if q == 0: gaps = "هیچ جست‌وجوی اختصاصی انجام نشد"
    rows.append([b[0], b[1], GROUP[b[3]], coverage_status(b), q, nsrc,
                 "جست‌وجوی وب (فارسی/انگلیسی) — رسانه‌های خبری و تخصصی" if q else "—",
                 to_date(D.RUN_DATE) if q else D.U,
                 "صفحات منبع مستقیم باز نشدند؛ فقط خلاصه نتایج موتور جست‌وجو" if q else "بررسی‌نشده", gaps])
sheet(wb, "07_Brand_Coverage", ["Brand_ID", "برند", "دسته", "وضعیت تحقیق", "تعداد جست‌وجوی اختصاصی", "تعداد منابع مرتبط ثبت‌شده", "کانال‌های بررسی‌شده",
      "آخرین تاریخ بررسی", "محدودیت‌ها", "شکاف‌های تحقیق"], rows, [9, 34, 16, 30, 12, 12, 30, 12, 34, 50])

# 08 dashboard (formulas against 02/01 so it updates with data)
ws = wb.create_sheet("08_Executive_Dashboard"); ws.sheet_view.rightToLeft = True
ws["A1"] = "داشبورد مدیریتی — کمپین‌های خودرویی ایران (۱۸ مهر ۱۴۰۴ تا ۱۸ مهر ۱۴۰۵)"; ws["A1"].font = Font(bold=True, size=14, name="Tahoma")
ws["A2"] = f"تاریخ اجرا: {D.RUN_DATE} — پیش‌نویس تحقیق؛ پوشش ناقص (به شیت 09 مراجعه کنید)"; ws["A2"].font = BODY
n = len(camps) + 1
kpis = [
 ("تعداد برندهای فهرست مادر", "=COUNTA('01_Master_Brands'!A2:A200)"),
 ("برندهای دارای کمپین/رویداد مستند", "=COUNTIF('01_Master_Brands'!L2:L200,\">0\")"),
 ("کل رکوردهای ثبت‌شده (شامل اکوسیستم)", f"=COUNTA('02_Campaigns'!A2:A{n})"),
 ("کمپین‌های تبلیغاتی/ترویجی", f"=COUNTIF('02_Campaigns'!B2:B{n},\"{D.RT_CAMP}\")"),
 ("کمپین‌های نمایشگاهی/رویدادی", f"=COUNTIF('02_Campaigns'!B2:B{n},\"{D.RT_EXPO}\")"),
 ("طرح‌های فروش (مرز با کمپین)", f"=COUNTIF('02_Campaigns'!B2:B{n},\"{D.RT_SALES}\")"),
 ("رکوردهای اکوسیستم (خارج از جامعه)", f"=COUNTIF('02_Campaigns'!B2:B{n},\"{D.RT_X}\")"),
 ("رکوردهای «تأییدشده»", f"=COUNTIF('02_Campaigns'!Z2:Z{n},\"تأییدشده*\")"),
 ("رکوردهای «نیازمند تأیید»", f"=COUNTIF('02_Campaigns'!Z2:Z{n},\"نیازمند تأیید*\")"),
 ("اطمینان High", f"=COUNTIF('02_Campaigns'!AA2:AA{n},\"High\")"),
 ("اطمینان Medium", f"=COUNTIF('02_Campaigns'!AA2:AA{n},\"Medium\")"),
 ("اطمینان Low", f"=COUNTIF('02_Campaigns'!AA2:AA{n},\"Low\")"),
 ("منابع/شواهد ثبت‌شده", "=COUNTA('03_Evidence_Sources'!A2:A500)"),
 ("کمپین‌های دارای آژانس شناسایی‌شده", f"=COUNTIFS('02_Campaigns'!U2:U{n},\"<>نامشخص\")"),
]
ws["A4"], ws["B4"] = "شاخص", "مقدار"
for c in (ws["A4"], ws["B4"]): c.font, c.fill = HDR_FONT, HDR_FILL
for i, (k, f) in enumerate(kpis, 5):
    ws.cell(i, 1, k).font = BODY; ws.cell(i, 2, f).font = Font(bold=True, name="Tahoma")
ws.column_dimensions["A"].width = 42; ws.column_dimensions["B"].width = 12
# helper: by group
r0 = 5 + len(kpis) + 2
ws.cell(r0, 1, "کمپین/رویداد مستند به تفکیک دسته بازار").font = Font(bold=True, name="Tahoma")
grp_counts = Counter()
for c in camps:
    if not is_brand_campaign(c): continue
    gs = {brand_by_id[b][3] for b in filter(None, c["brands"].split(";"))} or {"C"}
    for g in gs: grp_counts[g] += 1
for i, g in enumerate("ABC", 1):
    ws.cell(r0 + i, 1, GROUP[g]).font = BODY; ws.cell(r0 + i, 2, grp_counts[g]).font = BODY
ch = BarChart(); ch.type = "bar"; ch.title = "کمپین/رویداد مستند بر حسب دسته"; ch.legend = None
ch.add_data(Reference(ws, min_col=2, min_row=r0 + 1, max_row=r0 + 3)); ch.set_categories(Reference(ws, min_col=1, min_row=r0 + 1, max_row=r0 + 3))
ch.height, ch.width = 7, 14; ws.add_chart(ch, "D4")
# helper: top brands
r1 = r0 + 6
ws.cell(r1, 1, "برندها بر اساس تعداد کمپین/رویداد مستند (بدون طرح فروش)").font = Font(bold=True, name="Tahoma")
top = sorted(((sum(is_brand_campaign(c) for c in brand_camps[b[0]]), b[1]) for b in D.BRANDS if brand_camps.get(b[0])), reverse=True)
top = [t for t in top if t[0] > 0][:12]
for i, (v, name) in enumerate(top, 1):
    ws.cell(r1 + i, 1, name).font = BODY; ws.cell(r1 + i, 2, v).font = BODY
ch2 = BarChart(); ch2.type = "bar"; ch2.title = "برندهای پرفعالیت (کمپین/رویداد مستند)"; ch2.legend = None
ch2.add_data(Reference(ws, min_col=2, min_row=r1 + 1, max_row=r1 + len(top))); ch2.set_categories(Reference(ws, min_col=1, min_row=r1 + 1, max_row=r1 + len(top)))
ch2.height, ch2.width = 9, 14; ws.add_chart(ch2, "D20")
# helper: by record type
r2 = r1 + len(top) + 3
ws.cell(r2, 1, "رکوردها بر حسب نوع").font = Font(bold=True, name="Tahoma")
for i, rt in enumerate([D.RT_CAMP, D.RT_EXPO, D.RT_SALES, D.RT_X], 1):
    ws.cell(r2 + i, 1, rt).font = BODY; ws.cell(r2 + i, 2, f"=COUNTIF('02_Campaigns'!B2:B{n},A{r2+i})").font = BODY
ch3 = BarChart(); ch3.type = "col"; ch3.title = "رکوردها بر حسب نوع"; ch3.legend = None
ch3.add_data(Reference(ws, min_col=2, min_row=r2 + 1, max_row=r2 + 4)); ch3.set_categories(Reference(ws, min_col=1, min_row=r2 + 1, max_row=r2 + 4))
ch3.height, ch3.width = 7, 14; ws.add_chart(ch3, "D39")
r3 = r2 + 7
for i, t in enumerate([
 "یافته کلیدی ۱: پررنگ‌ترین شاهد فعالیت تبلیغاتی دوره، هفتمین نمایشگاه خودرو تهران (۲۲-۲۶ شهریور ۱۴۰۵) است که از بهمن ۱۴۰۴ به تعویق افتاده بود.",
 "یافته کلیدی ۲: بخش بزرگ ارتباطات بازاریابی خودروسازان در قالب بخشنامه‌های فروش با پیام «بدون قرعه‌کشی / قیمت قطعی / تحویل ۳۰ روزه» دیده می‌شود.",
 "یافته کلیدی ۳: برای هیچ کمپین خودروسازی آژانس خلاقه، بودجه یا شاخص عملکرد عمومی پیدا نشد.",
 "یافته کلیدی ۴: تنها شواهد خلاقه OOH مستند در دوره مربوط به بازیگران اکوسیستم (خودرو۴۵، کاسپین) است، نه خودروسازان.",
 "هشدار: صفحات منبع مستقیماً باز نشدند؛ شبکه‌های اجتماعی و سایت‌های رسمی ایرانی از محیط اجرا در دسترس نبودند."], 0):
    ws.cell(r3 + i, 1, t).font = BODY

# 09 methodology
meth = [
 ("بازه تحقیق", "۱۰ اکتبر ۲۰۲۵ (۱۸ مهر ۱۴۰۴) تا ۱۰ اکتبر ۲۰۲۶ (۱۸ مهر ۱۴۰۵)"),
 ("تاریخ اجرا", D.RUN_DATE),
 ("ابزارهای واقعاً استفاده‌شده", "ابزار جست‌وجوی وب (حدود ۴۵ جست‌وجوی فارسی و انگلیسی، با حالت extended برای اغلب پرسش‌ها). خروجی این ابزار عنوان، URL و خلاصه متنی نتایج است."),
 ("ابزارهای ناموفق", "WebFetch: دامنه‌های ایرانی resolve نشدند (ENOTFOUND). curl مستقیم: timeout. Browser Control (browser-use): daemon بالا نیامد. بنابراین هیچ صفحه‌ای مستقیماً باز نشد و هیچ اسکرین‌شات یا رسانه‌ای ذخیره نشد."),
 ("منابع بررسی‌نشده", "وب‌سایت‌های رسمی خودروسازان، اینستاگرام، تلگرام، آپارات، یوتیوب، ایتا، Meta Ad Library/Google Ads Transparency، آرشیو وب، نمونه‌کار آژانس‌ها."),
 ("تعریف رکوردها", f"«{D.RT_CAMP}»: فعالیت ترویجی/ارتباطی با هدف بازاریابی. «{D.RT_EXPO}»: حضور/رونمایی در نمایشگاه یا رویداد. «{D.RT_SALES}»: بخشنامه/طرح فروش که فعالیت فروش عادی است و فقط در صورت شواهد ترویجی مستقل کمپین محسوب می‌شود. «{D.RT_X}»: بازیگران غیرخودروساز (پلتفرم، روانکار) برای مقایسه خلاقه."),
 ("دسته‌بندی کمپین", "۱۲ دسته دستورالعمل (آگاهی، رونمایی، فروش، پیش‌فروش، تخفیف/اقساط، OOH، سوشال/اینفلوئنسر، ویدئو/TV، روابط عمومی، نمایشگاه، پس از فروش، مسئولیت اجتماعی)."),
 ("قواعد تاریخ", "تاریخ اجرا از تاریخ انتشار جدا ثبت شده است. تاریخ میلادی فقط وقتی پر شده که روز مشخص در منبع آمده باشد؛ تبدیل شمسی به میلادی با تقویم رسمی (۱ فروردین ۱۴۰۵ = ۲۱ مارس ۲۰۲۶). موارد شروع‌شده پیش از بازه فقط در صورت ادامه در بازه ثبت شده‌اند."),
 ("سطح اطمینان", "High: شاهد رسمی مستقیم (در این اجرا هیچ موردی، چون صفحات باز نشدند). Medium: دو یا چند منبع مستقل هم‌راستا در نتایج. Low: یک منبع، خلاصه ناقص یا تاریخ نامطمئن."),
 ("انتساب", "برند، شرکت عرضه‌کننده، واردکننده و پلتفرم شناسه مستقل دارند (BR-/CO-). مثال: تیسان (BR-021) برند تیگارد موتور (CO-009) و ری‌بج Jetour (BR-046) است، نه ایران خودرو."),
 ("حذف تکرار", "گزارش‌های متعدد درباره یک رویداد به یک رکورد متصل شده‌اند (مثلاً رونمایی تیسان: سه منبع ← CMP-032)."),
 ("بودجه و عملکرد", "برای هیچ رکوردی داده عمومی قابل اتکا پیدا نشد؛ هیچ تخمینی ساخته نشده است."),
 ("زمینه دوره", "اختلال اینترنت و افت شدید تبلیغات محیطی و سوشال پس از دی ۱۴۰۴ (EV-001، فقط عنوان)؛ جنگ از اواخر فوریه ۲۰۲۶ و تعویق نمایشگاه تهران (EV-005)؛ توقف تبلیغ خودرو در تلویزیون (EV-002، تاریخ نامشخص)."),
 ("وضعیت پروژه", "پیش‌نویس اول. ادعای جامعیت نمی‌شود. «کمپین پیدا نشد» به معنای «کمپینی اجرا نشده» نیست."),
 ("گام‌های بعدی پیشنهادی", "۱) اجرای مجدد از محیطی با دسترسی به دامنه‌های ایرانی و مرورگر؛ ۲) بررسی اینستاگرام/آپارات رسمی هر برند؛ ۳) پایش میدانی/عکس بیلبوردهای تهران؛ ۴) مصاحبه با آژانس‌ها و رسانه‌فروشان محیطی؛ ۵) تکمیل برندهای «بررسی‌نشده»."),
]
sheet(wb, "09_Research_Methodology", ["موضوع", "شرح"], meth, [26, 140], table=False)
wb.active = 0
wb.save(OUT)

with open("evidence_library.csv", "w", newline="", encoding="utf-8-sig") as f:
    w = csv.writer(f)
    w.writerow(["evidence_id", "outlet", "url", "source_type", "publish_date_as_seen", "access_date", "claim", "linked_records", "access_method", "media_rights_note"])
    for e in D.EVIDENCE:
        w.writerow([e[0], e[1], e[2], e[3], e[4], D.RUN_DATE, e[5], ";".join(ev_camps.get(e[0], [])), e[6], "فقط لینک؛ هیچ تصویر/ویدئویی بازنشر یا ذخیره نشده"])
print("ok", OUT)
