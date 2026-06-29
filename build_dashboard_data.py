#!/usr/bin/env python3
"""
Read the store workbook and compute every dashboard metric into data.js.

The dashboard reads ONLY what this script produces — there is no hardcoded
data anywhere in the page. Re-run this after editing the spreadsheet to
refresh the dashboard:

    python3 build_dashboard_data.py [path-to.xlsx]

Default input: store-data.xlsx   Output: data.js  (window.STORE_DATA = {...})
"""
import sys
import json
from collections import defaultdict
from datetime import datetime
import openpyxl

SRC = sys.argv[1] if len(sys.argv) > 1 else "store-data.xlsx"
OUT = "data.js"

wb = openpyxl.load_workbook(SRC, data_only=True)


def rows(sheet):
    """Yield each data row of a sheet as a dict keyed by header."""
    ws = wb[sheet]
    headers = [c.value for c in ws[1]]
    for r in ws.iter_rows(min_row=2, values_only=True):
        if all(v is None for v in r):
            continue
        yield dict(zip(headers, r))


def d(x):
    return x.date().isoformat() if isinstance(x, datetime) else x


def f(x):
    return round(float(x), 2) if x is not None else 0.0


# ---------------------------------------------------------------------------
orders = list(rows("Orders"))
items = list(rows("Order_Items"))
products = list(rows("Products"))
customers = list(rows("Customers"))
daily_src = list(rows("Daily_Metrics"))
channel_src = list(rows("Channel_Summary"))
returns = list(rows("Returns"))

dates = [o["Date"] for o in orders if o["Date"]]
start, end = min(dates), max(dates)

# ---- Headline KPIs (computed from the rows, not copied) -------------------
total_rev = sum(f(o["Total"]) for o in orders)
total_orders = len(orders)
total_margin = sum(f(o["Gross Margin"]) for o in orders)
total_cogs = sum(f(o["COGS"]) for o in orders)
total_ad = sum(f(o["Ad Cost"]) for o in orders)
total_discount = sum(f(o["Discount"]) for o in orders)
aov = total_rev / total_orders if total_orders else 0
roas = total_rev / total_ad if total_ad else 0
returned = sum(1 for o in orders if o["Status"] == "Returned")
refund_total = sum(f(r["Refund Amount"]) for r in returns)
return_rate = returned / total_orders if total_orders else 0

active_customers = [c for c in customers if (c["Orders"] or 0) > 0]
repeat_customers = [c for c in active_customers if (c["Orders"] or 0) >= 2]
repeat_rate = len(repeat_customers) / len(active_customers) if active_customers else 0
avg_ltv = (sum(f(c["Lifetime Value"]) for c in active_customers) /
           len(active_customers)) if active_customers else 0
avg_cvr = (sum(f(r["Conversion Rate"]) for r in daily_src) /
           len(daily_src)) if daily_src else 0

kpis = {
    "revenue": round(total_rev, 2),
    "orders": total_orders,
    "aov": round(aov, 2),
    "grossMargin": round(total_margin, 2),
    "grossMarginPct": round(total_margin / total_rev, 4) if total_rev else 0,
    "adSpend": round(total_ad, 2),
    "roas": round(roas, 2),
    "returnRate": round(return_rate, 4),
    "refundTotal": round(refund_total, 2),
    "conversionRate": round(avg_cvr, 4),
    "repeatRate": round(repeat_rate, 4),
    "customers": len(active_customers),
    "avgLtv": round(avg_ltv, 2),
    "discountTotal": round(total_discount, 2),
}

# ---- Daily series (with 7-day moving average of revenue) ------------------
daily_src_sorted = sorted(daily_src, key=lambda r: r["Date"])
daily = []
rev_window = []
for r in daily_src_sorted:
    rev = f(r["Revenue"])
    rev_window.append(rev)
    if len(rev_window) > 7:
        rev_window.pop(0)
    ma = sum(rev_window) / len(rev_window)
    daily.append({
        "date": d(r["Date"]),
        "revenue": rev,
        "margin": f(r["Gross Margin"]),
        "orders": int(r["Orders"] or 0),
        "aov": f(r["AOV"]),
        "adSpend": f(r["Ad Spend"]),
        "sessions": int(r["Sessions"] or 0),
        "cvr": round(float(r["Conversion Rate"] or 0), 4),
        "ma7": round(ma, 2),
    })

# ---- Monthly rollup -------------------------------------------------------
mon = defaultdict(lambda: {"revenue": 0.0, "margin": 0.0, "orders": 0,
                           "adSpend": 0.0})
for o in orders:
    key = o["Date"].strftime("%Y-%m")
    mon[key]["revenue"] += f(o["Total"])
    mon[key]["margin"] += f(o["Gross Margin"])
    mon[key]["orders"] += 1
    mon[key]["adSpend"] += f(o["Ad Cost"])
monthly = []
for k in sorted(mon):
    m = mon[k]
    monthly.append({
        "month": k,
        "label": datetime.strptime(k, "%Y-%m").strftime("%b %Y"),
        "revenue": round(m["revenue"], 2),
        "margin": round(m["margin"], 2),
        "orders": m["orders"],
        "aov": round(m["revenue"] / m["orders"], 2) if m["orders"] else 0,
        "adSpend": round(m["adSpend"], 2),
    })

# ---- Channels -------------------------------------------------------------
channels = []
for r in channel_src:
    channels.append({
        "channel": r["Channel"],
        "orders": int(r["Orders"] or 0),
        "revenue": f(r["Revenue"]),
        "share": round(float(r["% of Revenue"] or 0), 4),
        "adSpend": f(r["Ad Spend"]),
        "roas": f(r["ROAS"]),
        "cac": f(r["CAC (per order)"]),
        "grossMargin": f(r["Gross Margin"]),
        "netMargin": f(r["Net Margin (after ad)"]),
    })
channels.sort(key=lambda c: c["revenue"], reverse=True)

# ---- Products -------------------------------------------------------------
prods = []
for p in products:
    prods.append({
        "sku": p["SKU"],
        "name": p["Product"],
        "category": p["Category"],
        "price": f(p["Price"]),
        "marginPct": round(float(p["Margin %"] or 0), 4),
        "units": int(p["Units Sold"] or 0),
        "revenue": f(p["Revenue"]),
        "grossMargin": f(p["Gross Margin"]),
        "inventory": int(p["Inventory On Hand"] or 0),
    })
prods.sort(key=lambda p: p["revenue"], reverse=True)

# ---- Category mix (from line items) ---------------------------------------
cat = defaultdict(lambda: {"revenue": 0.0, "units": 0})
for it in items:
    c = cat[it["Category"]]
    c["revenue"] += f(it["Line Revenue"])
    c["units"] += int(it["Qty"] or 0)
categories = sorted(
    [{"category": k, "revenue": round(v["revenue"], 2), "units": v["units"]}
     for k, v in cat.items()],
    key=lambda x: x["revenue"], reverse=True)

# ---- Customer segments ----------------------------------------------------
seg = defaultdict(lambda: {"count": 0, "revenue": 0.0})
for c in customers:
    s = c["Segment"] or "Unknown"
    seg[s]["count"] += 1
    seg[s]["revenue"] += f(c["Lifetime Value"])
segments = [{"segment": k, "count": v["count"], "revenue": round(v["revenue"], 2)}
            for k, v in seg.items()]
seg_order = {"VIP": 0, "Repeat": 1, "One-time": 2, "No purchase yet": 3}
segments.sort(key=lambda s: seg_order.get(s["segment"], 9))

# ---- New vs returning revenue --------------------------------------------
nr = defaultdict(lambda: {"orders": 0, "revenue": 0.0})
for o in orders:
    k = o["Type"]
    nr[k]["orders"] += 1
    nr[k]["revenue"] += f(o["Total"])
new_returning = {
    "new": {"orders": nr["New"]["orders"], "revenue": round(nr["New"]["revenue"], 2)},
    "returning": {"orders": nr["Returning"]["orders"], "revenue": round(nr["Returning"]["revenue"], 2)},
}

# ---- Geography ------------------------------------------------------------
geo = defaultdict(lambda: {"orders": 0, "revenue": 0.0})
for o in orders:
    g = geo[o["Country"]]
    g["orders"] += 1
    g["revenue"] += f(o["Total"])
geography = sorted(
    [{"country": k, "orders": v["orders"], "revenue": round(v["revenue"], 2)}
     for k, v in geo.items()],
    key=lambda x: x["revenue"], reverse=True)

# ---- Returns analysis -----------------------------------------------------
ret_reason = defaultdict(lambda: {"count": 0, "amount": 0.0})
for r in returns:
    rr = ret_reason[r["Reason"]]
    rr["count"] += 1
    rr["amount"] += f(r["Refund Amount"])
returns_by_reason = sorted(
    [{"reason": k, "count": v["count"], "amount": round(v["amount"], 2)}
     for k, v in ret_reason.items()],
    key=lambda x: x["count"], reverse=True)

data = {
    "meta": {
        "store": "Lather & Co.",
        "subtitle": "Shampoo & haircare — performance overview",
        "start": d(start),
        "end": d(end),
        "days": (end - start).days + 1,
        "source": SRC,
    },
    "kpis": kpis,
    "daily": daily,
    "monthly": monthly,
    "channels": channels,
    "products": prods,
    "categories": categories,
    "segments": segments,
    "newReturning": new_returning,
    "geography": geography,
    "returns": {
        "rate": round(return_rate, 4),
        "total": round(refund_total, 2),
        "count": len(returns),
        "byReason": returns_by_reason,
    },
}

with open(OUT, "w") as fh:
    fh.write("// Auto-generated from " + SRC + " by build_dashboard_data.py.\n")
    fh.write("// Do not edit by hand — edit the spreadsheet and re-run the script.\n")
    fh.write("window.STORE_DATA = ")
    json.dump(data, fh, separators=(",", ":"))
    fh.write(";\n")

print(f"Wrote {OUT} from {SRC}")
print(f"  revenue=${kpis['revenue']:,.0f}  orders={kpis['orders']:,}  "
      f"AOV=${kpis['aov']:.2f}  margin={kpis['grossMarginPct']*100:.1f}%  "
      f"ROAS={kpis['roas']:.1f}x  return_rate={kpis['returnRate']*100:.1f}%")
print(f"  channels={len(channels)} products={len(prods)} "
      f"categories={len(categories)} months={len(monthly)} "
      f"countries={len(geography)}")
