#!/usr/bin/env python3
"""
Generate a highly detailed sample ecommerce dataset for a shampoo store
(Lather & Co.) as a multi-sheet .xlsx workbook.

All data is synthetic but internally consistent: order line items roll up to
orders, orders roll up to daily KPIs, products carry cost/margin, customers
carry lifetime value, etc. Deterministic (fixed seed) so it regenerates the
same every time.

Sheets:
  README          – what each sheet/column means
  Orders          – one row per order (header level)
  Order_Items     – one row per product line within an order
  Products        – catalog with cost, price, margin, inventory
  Customers       – customer list with cohort, LTV, orders
  Daily_Metrics   – per-day revenue, orders, sessions, CVR, AOV, ad spend, ROAS
  Channel_Summary – revenue/orders/spend by marketing channel
  Returns         – returned/refunded orders with reasons
"""
import random
from datetime import date, timedelta
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

random.seed(42)

# ----------------------------------------------------------------------------
# Reference data
# ----------------------------------------------------------------------------
START = date(2025, 7, 1)
END = date(2026, 6, 29)          # ~1 year of history, ends "today" in the demo
DAYS = (END - START).days + 1

PRODUCTS = [
    # sku, name, category, size, price, unit_cost, popularity weight
    ("SH-ROS-300", "Rosemary Mint Strengthening Shampoo", "Anti-hair-loss", "300ml", 24.00, 6.40, 30),
    ("SH-ARG-300", "Argan Repair Shampoo",                "Dry / damaged",  "300ml", 22.00, 5.90, 24),
    ("SH-CLR-500", "Daily Clarifying Wash",               "All hair types", "500ml", 18.00, 4.20, 20),
    ("KT-VOL-DUO", "Volume Boost Duo (Shampoo + Cond.)",  "Volumizing",     "2x250ml", 38.00, 10.10, 16),
    ("SH-CLR-300", "Color-Safe Sulfate-Free Shampoo",     "Color-treated",  "300ml", 26.00, 7.30, 10),
    ("CN-ARG-300", "Argan Repair Conditioner",            "Dry / damaged",  "300ml", 22.00, 5.80, 14),
    ("TR-SCL-100", "Scalp Detox Treatment",               "Treatment",      "100ml", 32.00, 8.90, 8),
    ("AC-BRS-001", "Bamboo Detangling Brush",             "Accessories",    "-",     14.00, 3.10, 9),
]

COUNTRIES = [("United States", 0.55), ("Canada", 0.12), ("United Kingdom", 0.11),
             ("Australia", 0.08), ("Germany", 0.07), ("France", 0.04), ("Other", 0.03)]

CHANNELS = [  # name, share of orders, CAC (avg ad cost per order; 0 = organic)
    ("Online store (direct/SEO)", 0.34, 0.0),
    ("Instagram / Meta Ads",      0.24, 9.50),
    ("Google Ads",                0.16, 8.20),
    ("Amazon",                    0.14, 6.00),
    ("TikTok Shop",               0.07, 7.40),
    ("Email / SMS",               0.05, 0.40),
]

RETURN_REASONS = ["Didn't like the scent", "Caused irritation", "Arrived damaged",
                  "Wrong item shipped", "Changed mind", "Found cheaper elsewhere"]

FIRST_NAMES = ["Emma", "Liam", "Olivia", "Noah", "Ava", "Sophia", "Isabella", "Mia",
               "Amelia", "Harper", "Ethan", "Mason", "Lucas", "Aiden", "Chloe",
               "Grace", "Zoe", "Layla", "Nora", "Hannah", "Maya", "Ruby", "Leo", "Ivy"]
LAST_NAMES = ["Smith", "Johnson", "Williams", "Brown", "Jones", "Garcia", "Miller",
              "Davis", "Rodriguez", "Martinez", "Lee", "Walker", "Hall", "Allen",
              "Young", "King", "Wright", "Scott", "Green", "Adams", "Baker", "Nelson"]


def weighted_choice(pairs):
    r = random.random()
    cum = 0.0
    for item, w in pairs:
        cum += w
        if r <= cum:
            return item
    return pairs[-1][0]


def product_weighted():
    total = sum(p[6] for p in PRODUCTS)
    r = random.random() * total
    cum = 0
    for p in PRODUCTS:
        cum += p[6]
        if r <= cum:
            return p
    return PRODUCTS[-1]


# Seasonality multiplier: weekends up, slow growth over the year, Nov/Dec spike
def day_multiplier(d, idx):
    growth = 1 + 0.45 * (idx / DAYS)                 # store grows ~45% over the year
    weekend = 1.22 if d.weekday() >= 5 else 1.0
    holiday = 1.0
    if d.month == 11 and d.day >= 24:                # Black Friday / Cyber week
        holiday = 1.9
    elif d.month == 12 and d.day <= 20:              # holiday shopping
        holiday = 1.5
    elif d.month == 1:                               # post-holiday slump
        holiday = 0.82
    noise = random.uniform(0.85, 1.15)
    return growth * weekend * holiday * noise


# ----------------------------------------------------------------------------
# Build customers
# ----------------------------------------------------------------------------
N_CUSTOMERS = 1400
customers = []
for i in range(N_CUSTOMERS):
    fn = random.choice(FIRST_NAMES)
    ln = random.choice(LAST_NAMES)
    cust_id = 10000 + i
    signup = START + timedelta(days=random.randint(0, DAYS - 1))
    customers.append({
        "id": cust_id,
        "name": f"{fn} {ln}",
        "email": f"{fn.lower()}.{ln.lower()}{random.randint(1,99)}@example.com",
        "country": weighted_choice(COUNTRIES),
        "signup": signup,
        "acq_channel": weighted_choice([(c[0], c[1]) for c in CHANNELS]),
        "orders": 0,
        "revenue": 0.0,
        "first_order": None,
        "last_order": None,
    })

# ----------------------------------------------------------------------------
# Generate orders + line items day by day
# ----------------------------------------------------------------------------
orders = []
order_items = []
returns = []
order_seq = 1000

channel_cac = {c[0]: c[2] for c in CHANNELS}

for idx in range(DAYS):
    d = START + timedelta(days=idx)
    base_orders = 6.5                                 # avg orders/day baseline
    n_orders = max(0, int(round(base_orders * day_multiplier(d, idx))))
    for _ in range(n_orders):
        order_seq += 1
        order_id = f"LC-{order_seq}"
        cust = random.choice(customers)

        # 1–4 line items
        n_lines = random.choices([1, 2, 3, 4], weights=[55, 28, 12, 5])[0]
        chosen = {}
        for _ in range(n_lines):
            p = product_weighted()
            qty = random.choices([1, 2, 3], weights=[78, 18, 4])[0]
            chosen[p[0]] = chosen.get(p[0], (p, 0))
            chosen[p[0]] = (p, chosen[p[0]][1] + qty)

        subtotal = 0.0
        cogs = 0.0
        for sku, (p, qty) in chosen.items():
            line_total = p[4] * qty
            line_cost = p[5] * qty
            subtotal += line_total
            cogs += line_cost
            order_items.append({
                "order_id": order_id, "date": d, "sku": sku, "product": p[1],
                "category": p[2], "qty": qty, "unit_price": p[4],
                "unit_cost": p[5], "line_revenue": round(line_total, 2),
                "line_cost": round(line_cost, 2),
                "line_margin": round(line_total - line_cost, 2),
            })

        # discount: ~30% of orders use a code
        discount_pct = random.choice([0, 0, 0, 0.10, 0.15, 0.20]) if random.random() < 0.30 else 0
        discount = round(subtotal * discount_pct, 2)
        # shipping: free over $35
        shipping = 0.0 if (subtotal - discount) >= 35 else 4.95
        tax = round((subtotal - discount) * 0.07, 2)
        total = round(subtotal - discount + shipping + tax, 2)

        channel = cust["acq_channel"] if cust["orders"] == 0 else weighted_choice([(c[0], c[1]) for c in CHANNELS])
        ad_cost = round(channel_cac.get(channel, 0.0) * (1 if cust["orders"] == 0 else 0.15), 2)

        status = "Completed"
        is_return = random.random() < 0.042             # ~4.2% return rate
        if is_return:
            status = "Returned"

        order = {
            "order_id": order_id, "date": d, "customer_id": cust["id"],
            "customer": cust["name"], "country": cust["country"],
            "channel": channel, "items": sum(q for _, q in chosen.values()),
            "subtotal": round(subtotal, 2), "discount": discount,
            "discount_code": f"SAVE{int(discount_pct*100)}" if discount_pct else "",
            "shipping": shipping, "tax": tax, "total": total,
            "cogs": round(cogs, 2), "gross_margin": round(subtotal - discount - cogs, 2),
            "ad_cost": ad_cost,
            "new_or_returning": "New" if cust["orders"] == 0 else "Returning",
            "status": status,
        }
        orders.append(order)

        cust["orders"] += 1
        cust["revenue"] += total
        cust["first_order"] = cust["first_order"] or d
        cust["last_order"] = d

        if is_return:
            returns.append({
                "order_id": order_id, "date": d + timedelta(days=random.randint(3, 14)),
                "customer": cust["name"], "amount": total,
                "reason": random.choice(RETURN_REASONS),
                "channel": channel, "country": cust["country"],
            })

# ----------------------------------------------------------------------------
# Roll up daily metrics
# ----------------------------------------------------------------------------
daily = {}
for o in orders:
    d = o["date"]
    if d not in daily:
        daily[d] = {"orders": 0, "revenue": 0.0, "discount": 0.0, "cogs": 0.0,
                    "margin": 0.0, "ad_cost": 0.0, "items": 0, "new": 0}
    daily[d]["orders"] += 1
    daily[d]["revenue"] += o["total"]
    daily[d]["discount"] += o["discount"]
    daily[d]["cogs"] += o["cogs"]
    daily[d]["margin"] += o["gross_margin"]
    daily[d]["ad_cost"] += o["ad_cost"]
    daily[d]["items"] += o["items"]
    if o["new_or_returning"] == "New":
        daily[d]["new"] += 1

# ----------------------------------------------------------------------------
# Workbook styling helpers
# ----------------------------------------------------------------------------
wb = openpyxl.Workbook()

HEADER_FILL = PatternFill("solid", fgColor="1F2330")
HEADER_FONT = Font(color="FFFFFF", bold=True, size=11)
TITLE_FONT = Font(bold=True, size=14, color="2A2F3D")
SUB_FONT = Font(color="6B7280", size=10)
THIN = Side(style="thin", color="E2E5EC")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
MONEY = '"$"#,##0.00'
MONEY0 = '"$"#,##0'
PCT = '0.0%'
PCT2 = '0.00%'
ALT_FILL = PatternFill("solid", fgColor="F7F8FA")


def write_sheet(ws, headers, rows, formats=None, freeze="A2", widths=None):
    formats = formats or {}
    for c, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=c, value=h)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="left", vertical="center")
    for r, row in enumerate(rows, 2):
        for c, val in enumerate(row, 1):
            cell = ws.cell(row=r, column=c, value=val)
            if c in formats:
                cell.number_format = formats[c]
            if r % 2 == 0:
                cell.fill = ALT_FILL
    # widths
    for c, h in enumerate(headers, 1):
        letter = get_column_letter(c)
        if widths and c in widths:
            ws.column_dimensions[letter].width = widths[c]
        else:
            maxlen = max([len(str(h))] + [len(str(row[c-1])) for row in rows[:200]] + [8])
            ws.column_dimensions[letter].width = min(maxlen + 3, 38)
    ws.freeze_panes = freeze
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}1"


# ---------- README ----------
ws = wb.active
ws.title = "README"
ws.sheet_view.showGridLines = False
ws["A1"] = "Lather & Co. — Sample Ecommerce Dataset"
ws["A1"].font = TITLE_FONT
ws["A2"] = "Synthetic but internally consistent data for a shampoo store. Use it to experiment with your dashboard, pivot tables, and charts."
ws["A2"].font = SUB_FONT
readme_rows = [
    ("", ""),
    ("Period covered", f"{START.isoformat()}  →  {END.isoformat()}  ({DAYS} days)"),
    ("Total orders", f"{len(orders):,}"),
    ("Total customers", f"{len(customers):,}"),
    ("", ""),
    ("SHEET", "WHAT'S IN IT"),
    ("Orders", "One row per order — totals, channel, country, margin, new vs returning, status."),
    ("Order_Items", "One row per product line inside an order — qty, unit price/cost, line margin."),
    ("Products", "Catalog: price, unit cost, margin %, units sold, revenue, inventory on hand."),
    ("Customers", "Customer list with signup date, acquisition channel, # orders, lifetime value."),
    ("Daily_Metrics", "Per-day revenue, orders, AOV, sessions, conversion rate, ad spend, ROAS, new-customer %."),
    ("Channel_Summary", "Revenue / orders / ad spend / ROAS / CAC grouped by marketing channel."),
    ("Returns", "Returned orders with refund amount and reason."),
    ("", ""),
    ("TIPS", ""),
    ("•", "Every sheet has filter dropdowns (row 1) and frozen headers."),
    ("•", "Orders.total = subtotal − discount + shipping + tax. gross_margin = subtotal − discount − cogs."),
    ("•", "Daily_Metrics.ROAS = revenue ÷ ad_spend. Conversion rate = orders ÷ sessions."),
    ("•", "Feed Daily_Metrics into the dashboard's revenue chart; Products into the top-products table."),
]
for r, (a, b) in enumerate(readme_rows, 4):
    ca = ws.cell(row=r, column=1, value=a)
    cb = ws.cell(row=r, column=2, value=b)
    if a in ("SHEET", "TIPS"):
        ca.font = Font(bold=True, color="2A2F3D")
        cb.font = Font(bold=True, color="2A2F3D")
ws.column_dimensions["A"].width = 18
ws.column_dimensions["B"].width = 95

# ---------- Orders ----------
ws = wb.create_sheet("Orders")
headers = ["Order ID", "Date", "Customer ID", "Customer", "Country", "Channel",
           "Items", "Subtotal", "Discount", "Code", "Shipping", "Tax", "Total",
           "COGS", "Gross Margin", "Ad Cost", "Type", "Status"]
rows = [[o["order_id"], o["date"], o["customer_id"], o["customer"], o["country"],
         o["channel"], o["items"], o["subtotal"], o["discount"], o["discount_code"],
         o["shipping"], o["tax"], o["total"], o["cogs"], o["gross_margin"],
         o["ad_cost"], o["new_or_returning"], o["status"]] for o in orders]
fmt = {2: "yyyy-mm-dd", 8: MONEY, 9: MONEY, 11: MONEY, 12: MONEY, 13: MONEY,
       14: MONEY, 15: MONEY, 16: MONEY}
write_sheet(ws, headers, rows, fmt,
            widths={1: 11, 2: 12, 4: 18, 5: 16, 6: 22, 10: 9})

# ---------- Order_Items ----------
ws = wb.create_sheet("Order_Items")
headers = ["Order ID", "Date", "SKU", "Product", "Category", "Qty",
           "Unit Price", "Unit Cost", "Line Revenue", "Line Cost", "Line Margin"]
rows = [[it["order_id"], it["date"], it["sku"], it["product"], it["category"],
         it["qty"], it["unit_price"], it["unit_cost"], it["line_revenue"],
         it["line_cost"], it["line_margin"]] for it in order_items]
fmt = {2: "yyyy-mm-dd", 7: MONEY, 8: MONEY, 9: MONEY, 10: MONEY, 11: MONEY}
write_sheet(ws, headers, rows, fmt, widths={2: 12, 3: 12, 4: 34, 5: 16})

# ---------- Products ----------
ws = wb.create_sheet("Products")
prod_stats = {}
for it in order_items:
    s = prod_stats.setdefault(it["sku"], {"units": 0, "revenue": 0.0, "margin": 0.0})
    s["units"] += it["qty"]
    s["revenue"] += it["line_revenue"]
    s["margin"] += it["line_margin"]
headers = ["SKU", "Product", "Category", "Size", "Price", "Unit Cost",
           "Margin %", "Units Sold", "Revenue", "Gross Margin", "Inventory On Hand"]
rows = []
for p in PRODUCTS:
    sku = p[0]
    st = prod_stats.get(sku, {"units": 0, "revenue": 0.0, "margin": 0.0})
    margin_pct = (p[4] - p[5]) / p[4]
    inv = random.randint(120, 900)
    rows.append([sku, p[1], p[2], p[3], p[4], p[5], margin_pct,
                 st["units"], round(st["revenue"], 2), round(st["margin"], 2), inv])
rows.sort(key=lambda r: r[8], reverse=True)
fmt = {5: MONEY, 6: MONEY, 7: PCT, 9: MONEY, 10: MONEY}
write_sheet(ws, headers, rows, fmt, widths={1: 12, 2: 34, 3: 16, 5: 9})

# ---------- Customers ----------
ws = wb.create_sheet("Customers")
headers = ["Customer ID", "Name", "Email", "Country", "Signup Date",
           "Acquisition Channel", "Orders", "Lifetime Value", "Avg Order Value",
           "First Order", "Last Order", "Segment"]
rows = []
for c in customers:
    if c["orders"] == 0:
        seg = "No purchase yet"
        aov = 0.0
    else:
        aov = c["revenue"] / c["orders"]
        if c["orders"] >= 4:
            seg = "VIP"
        elif c["orders"] >= 2:
            seg = "Repeat"
        else:
            seg = "One-time"
    rows.append([c["id"], c["name"], c["email"], c["country"], c["signup"],
                 c["acq_channel"], c["orders"], round(c["revenue"], 2),
                 round(aov, 2), c["first_order"], c["last_order"], seg])
rows.sort(key=lambda r: r[7], reverse=True)
fmt = {5: "yyyy-mm-dd", 8: MONEY, 9: MONEY, 10: "yyyy-mm-dd", 11: "yyyy-mm-dd"}
write_sheet(ws, headers, rows, fmt,
            widths={2: 18, 3: 30, 4: 16, 6: 22, 12: 14})

# ---------- Daily_Metrics ----------
ws = wb.create_sheet("Daily_Metrics")
headers = ["Date", "Day of Week", "Orders", "Units", "Revenue", "Discounts",
           "COGS", "Gross Margin", "AOV", "Ad Spend", "ROAS", "Sessions",
           "Conversion Rate", "New Customers", "New Customer %"]
rows = []
for d in sorted(daily.keys()):
    m = daily[d]
    aov = m["revenue"] / m["orders"] if m["orders"] else 0
    # sessions implied from a realistic ~2.4% conversion rate with noise
    cvr = random.uniform(0.019, 0.031)
    sessions = int(m["orders"] / cvr) if m["orders"] else 0
    roas = m["revenue"] / m["ad_cost"] if m["ad_cost"] else 0
    new_pct = m["new"] / m["orders"] if m["orders"] else 0
    rows.append([d, d.strftime("%a"), m["orders"], m["items"],
                 round(m["revenue"], 2), round(m["discount"], 2),
                 round(m["cogs"], 2), round(m["margin"], 2), round(aov, 2),
                 round(m["ad_cost"], 2), round(roas, 2), sessions,
                 m["orders"] / sessions if sessions else 0, m["new"], new_pct])
fmt = {1: "yyyy-mm-dd", 5: MONEY, 6: MONEY, 7: MONEY, 8: MONEY, 9: MONEY,
       10: MONEY, 11: '0.00"x"', 13: PCT2, 15: PCT}
write_sheet(ws, headers, rows, fmt, widths={1: 12, 2: 11, 13: 15, 15: 15})

# ---------- Channel_Summary ----------
ws = wb.create_sheet("Channel_Summary")
ch_stats = {}
for o in orders:
    s = ch_stats.setdefault(o["channel"], {"orders": 0, "revenue": 0.0,
                                            "ad": 0.0, "margin": 0.0})
    s["orders"] += 1
    s["revenue"] += o["total"]
    s["ad"] += o["ad_cost"]
    s["margin"] += o["gross_margin"]
headers = ["Channel", "Orders", "Revenue", "% of Revenue", "Ad Spend",
           "ROAS", "CAC (per order)", "Gross Margin", "Net Margin (after ad)"]
total_rev = sum(s["revenue"] for s in ch_stats.values())
rows = []
for ch, s in sorted(ch_stats.items(), key=lambda kv: kv[1]["revenue"], reverse=True):
    roas = s["revenue"] / s["ad"] if s["ad"] else 0
    cac = s["ad"] / s["orders"] if s["orders"] else 0
    rows.append([ch, s["orders"], round(s["revenue"], 2), s["revenue"] / total_rev,
                 round(s["ad"], 2), round(roas, 2), round(cac, 2),
                 round(s["margin"], 2), round(s["margin"] - s["ad"], 2)])
fmt = {3: MONEY, 4: PCT, 5: MONEY, 6: '0.00"x"', 7: MONEY, 8: MONEY, 9: MONEY}
write_sheet(ws, headers, rows, fmt, widths={1: 26, 4: 13, 7: 16, 9: 22})

# ---------- Returns ----------
ws = wb.create_sheet("Returns")
headers = ["Order ID", "Return Date", "Customer", "Refund Amount",
           "Reason", "Channel", "Country"]
rows = [[r["order_id"], r["date"], r["customer"], round(r["amount"], 2),
         r["reason"], r["channel"], r["country"]] for r in
        sorted(returns, key=lambda x: x["date"])]
fmt = {2: "yyyy-mm-dd", 4: MONEY}
write_sheet(ws, headers, rows, fmt, widths={1: 11, 2: 13, 3: 18, 5: 26, 6: 22, 7: 16})

# ----------------------------------------------------------------------------
out = "lather-and-co-sample-data.xlsx"
wb.save(out)
print(f"Saved {out}")
print(f"  Orders:       {len(orders):,}")
print(f"  Order items:  {len(order_items):,}")
print(f"  Customers:    {len(customers):,}")
print(f"  Daily rows:   {len(daily):,}")
print(f"  Returns:      {len(returns):,}")
print(f"  Total revenue: ${sum(o['total'] for o in orders):,.2f}")
