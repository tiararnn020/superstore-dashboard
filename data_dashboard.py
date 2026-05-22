import pandas as pd
import json
import os

# Buat folder data jika belum ada
os.makedirs('data', exist_ok=True)

df = pd.read_excel('superstore.xls', engine='xlrd')
df['Order Date'] = pd.to_datetime(df['Order Date'])
df['Year'] = df['Order Date'].dt.year
df['Month'] = df['Order Date'].dt.month
df['Month_Name'] = df['Order Date'].dt.strftime('%b %Y')
df['Quarter'] = df['Month'].apply(lambda m: f"Q{(m-1)//3 + 1}")
df['Year_Quarter'] = df['Year'].astype(str) + ' ' + df['Quarter']

# ── 1. KPI CARDS ─────────────────────────────────────────
kpi = {
    "total_sales":    round(df['Sales'].sum(), 2),
    "total_profit":   round(df['Profit'].sum(), 2),
    "total_orders":   int(df['Order ID'].nunique()),
    "total_customers":int(df['Customer ID'].nunique()),
    "profit_margin":  round(df['Profit'].sum() / df['Sales'].sum() * 100, 2),
    "avg_order_value":round(df.groupby('Order ID')['Sales'].sum().mean(), 2)
}
with open('data/kpi.json', 'w') as f:
    json.dump(kpi, f, indent=2)
print("✓ kpi.json")

# ── 2. REVENUE PER KUARTAL (Bar Chart) ───────────────────
rq = df.groupby('Year_Quarter').agg(
    sales=('Sales','sum'),
    profit=('Profit','sum'),
    orders=('Order ID','nunique')
).reset_index().sort_values('Year_Quarter')

revenue_quarterly = []
for _, row in rq.iterrows():
    revenue_quarterly.append({
        "period": row['Year_Quarter'],
        "sales":  round(row['sales'], 2),
        "profit": round(row['profit'], 2),
        "orders": int(row['orders'])
    })
with open('data/revenue_quarterly.json', 'w') as f:
    json.dump(revenue_quarterly, f, indent=2)
print("✓ revenue_quarterly.json")

# ── 3. REVENUE PER BULAN (Line Chart) ────────────────────
rm = df.groupby(['Year','Month','Month_Name']).agg(
    sales=('Sales','sum')
).reset_index().sort_values(['Year','Month'])

revenue_monthly = []
for _, row in rm.iterrows():
    revenue_monthly.append({
        "label": row['Month_Name'],
        "sales": round(row['sales'], 2),
        "year":  int(row['Year'])
    })
with open('data/revenue_monthly.json', 'w') as f:
    json.dump(revenue_monthly, f, indent=2)
print("✓ revenue_monthly.json")

# ── 4. REVENUE PER KATEGORI (Horizontal Bar) ─────────────
rc = df.groupby('Category').agg(
    sales=('Sales','sum'),
    profit=('Profit','sum'),
    orders=('Order ID','nunique')
).reset_index().sort_values('sales', ascending=False)

revenue_category = []
for _, row in rc.iterrows():
    revenue_category.append({
        "category": row['Category'],
        "sales":    round(row['sales'], 2),
        "profit":   round(row['profit'], 2),
        "margin":   round(row['profit'] / row['sales'] * 100, 2),
        "orders":   int(row['orders'])
    })
with open('data/revenue_category.json', 'w') as f:
    json.dump(revenue_category, f, indent=2)
print("✓ revenue_category.json")

# ── 5. PROFIT MARGIN PER SUB-CATEGORY (Horizontal Bar) ───
rsc = df.groupby('Sub-Category').agg(
    sales=('Sales','sum'),
    profit=('Profit','sum')
).reset_index()
rsc['margin'] = rsc['profit'] / rsc['sales'] * 100
rsc = rsc.sort_values('margin', ascending=False)

profit_margin = []
for _, row in rsc.iterrows():
    profit_margin.append({
        "subcategory": row['Sub-Category'],
        "sales":  round(row['sales'], 2),
        "profit": round(row['profit'], 2),
        "margin": round(row['margin'], 2)
    })
with open('data/profit_margin.json', 'w') as f:
    json.dump(profit_margin, f, indent=2)
print("✓ profit_margin.json")

# ── 6. REVENUE PER REGION (Horizontal Bar + Tabel) ───────
rr = df.groupby('Region').agg(
    sales=('Sales','sum'),
    profit=('Profit','sum'),
    orders=('Order ID','nunique'),
    customers=('Customer ID','nunique')
).reset_index().sort_values('sales', ascending=False)

revenue_region = []
for _, row in rr.iterrows():
    revenue_region.append({
        "region":    row['Region'],
        "sales":     round(row['sales'], 2),
        "profit":    round(row['profit'], 2),
        "margin":    round(row['profit'] / row['sales'] * 100, 2),
        "orders":    int(row['orders']),
        "customers": int(row['customers'])
    })
with open('data/revenue_region.json', 'w') as f:
    json.dump(revenue_region, f, indent=2)
print("✓ revenue_region.json")

# ── 7. REVENUE PER SEGMENT (Pie / Doughnut Chart) ────────
rs = df.groupby('Segment').agg(
    sales=('Sales','sum'),
    profit=('Profit','sum'),
    customers=('Customer ID','nunique')
).reset_index().sort_values('sales', ascending=False)

revenue_segment = []
for _, row in rs.iterrows():
    revenue_segment.append({
        "segment":   row['Segment'],
        "sales":     round(row['sales'], 2),
        "profit":    round(row['profit'], 2),
        "customers": int(row['customers'])
    })
with open('data/revenue_segment.json', 'w') as f:
    json.dump(revenue_segment, f, indent=2)
print("✓ revenue_segment.json")

# ── 8. TOP 10 SUB-CATEGORY BY SALES (untuk tabel) ────────
top_sub = df.groupby('Sub-Category').agg(
    sales=('Sales','sum'),
    profit=('Profit','sum'),
    quantity=('Quantity','sum')
).reset_index().sort_values('sales', ascending=False).head(10)

revenue_subcategory = []
for _, row in top_sub.iterrows():
    revenue_subcategory.append({
        "subcategory": row['Sub-Category'],
        "sales":    round(row['sales'], 2),
        "profit":   round(row['profit'], 2),
        "quantity": int(row['quantity']),
        "margin":   round(row['profit'] / row['sales'] * 100, 2)
    })
with open('data/revenue_subcategory.json', 'w') as f:
    json.dump(revenue_subcategory, f, indent=2)
print("✓ revenue_subcategory.json")

# ── 9. REVENUE PER REGION PER TAHUN (untuk tren & scatter) ──
rry = df.groupby(['Region', 'Year']).agg(
    sales=('Sales', 'sum'),
    profit=('Profit', 'sum'),
    orders=('Order ID', 'nunique'),
    customers=('Customer ID', 'nunique')
).reset_index()

revenue_region_yearly = []
for _, row in rry.iterrows():
    revenue_region_yearly.append({
        "region":    row['Region'],
        "year":      int(row['Year']),
        "sales":     round(row['sales'], 2),
        "profit":    round(row['profit'], 2),
        "margin":    round(row['profit'] / row['sales'] * 100, 2),
        "orders":    int(row['orders']),
        "customers": int(row['customers']),
        "aov":       round(row['sales'] / row['orders'], 2)
    })
with open('data/revenue_region_yearly.json', 'w') as f:
    json.dump(revenue_region_yearly, f, indent=2)
print("✓ revenue_region_yearly.json")

print("\n✅ Semua file JSON berhasil dibuat di folder data/")