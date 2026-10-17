"""Generate a realistic synthetic retail sales dataset (seeded, reproducible)."""
import numpy as np
import pandas as pd
from pathlib import Path
ROOT = Path(__file__).resolve().parent.parent

rng = np.random.default_rng(2711)

regions = {
    "North": ["Delhi", "Chandigarh", "Jaipur", "Lucknow"],
    "South": ["Bengaluru", "Chennai", "Hyderabad", "Kochi"],
    "East": ["Kolkata", "Bhubaneswar", "Patna", "Guwahati"],
    "West": ["Mumbai", "Pune", "Ahmedabad", "Aurangabad"],
}
region_w = {"North": 0.27, "South": 0.30, "East": 0.16, "West": 0.27}

# category -> subcategory -> list of (product, base price INR, cost ratio)
catalog = {
    "Electronics": {
        "Phones": [("Nova X5 Smartphone", 18999, 0.82), ("Nova Lite Smartphone", 9999, 0.80)],
        "Audio": [("BassPro Wireless Earbuds", 2499, 0.55), ("SoundBar 200", 7499, 0.68)],
        "Accessories": [("FastCharge 65W Adapter", 1499, 0.45), ("PowerBank 20000mAh", 1799, 0.50)],
        "Computers": [("SlimBook 14 Laptop", 52999, 0.86), ("Tab S Tablet", 21999, 0.80)],
    },
    "Home & Kitchen": {
        "Appliances": [("AirFry 4L", 5499, 0.62), ("MixMaster Grinder", 3299, 0.60)],
        "Cookware": [("Cast Iron Kadai", 1899, 0.48), ("Non-stick Pan Set", 2799, 0.52)],
        "Decor": [("Brass Table Lamp", 2199, 0.42), ("Cotton Cushion Set", 999, 0.38)],
    },
    "Fashion": {
        "Men": [("Linen Kurta", 1499, 0.40), ("Slim Fit Chinos", 1799, 0.44)],
        "Women": [("Silk Saree", 4999, 0.50), ("Cotton Kurti", 1199, 0.38)],
        "Footwear": [("Running Shoes", 3499, 0.52), ("Leather Sandals", 1699, 0.46)],
    },
    "Grocery": {
        "Staples": [("Basmati Rice 5kg", 699, 0.78), ("Cold-pressed Oil 1L", 349, 0.74)],
        "Beverages": [("Assam Tea 500g", 299, 0.65), ("Filter Coffee 250g", 349, 0.62)],
        "Snacks": [("Roasted Makhana", 249, 0.58), ("Millet Cookies", 179, 0.60)],
    },
}
cat_w = {"Electronics": 0.24, "Home & Kitchen": 0.20, "Fashion": 0.26, "Grocery": 0.30}
segments, seg_w = ["Consumer", "Corporate", "Small Business"], [0.62, 0.22, 0.16]
channels, ch_w = ["Online", "In-Store", "Marketplace"], [0.45, 0.33, 0.22]

# monthly seasonality (festive Oct-Nov peak, Jan sale bump) and yearly growth
season = np.array([1.05, 0.85, 0.90, 0.88, 0.92, 0.90, 0.95, 1.00, 1.08, 1.35, 1.40, 1.15])
growth = {2023: 1.00, 2024: 1.14, 2025: 1.27}
base_orders_per_month = 165

rows = []
oid = 100000
for year in (2023, 2024, 2025):
    for m in range(1, 13):
        n_orders = rng.poisson(base_orders_per_month * season[m - 1] * growth[year])
        days_in_month = pd.Period(f"{year}-{m:02d}").days_in_month
        for _ in range(n_orders):
            oid += 1
            date = pd.Timestamp(year, m, int(rng.integers(1, days_in_month + 1)))
            region = rng.choice(list(region_w), p=list(region_w.values()))
            city = rng.choice(regions[region])
            seg = rng.choice(segments, p=seg_w)
            ch = rng.choice(channels, p=ch_w)
            n_lines = rng.choice([1, 1, 1, 2, 2, 3])
            for _ in range(n_lines):
                cat = rng.choice(list(cat_w), p=list(cat_w.values()))
                sub = rng.choice(list(catalog[cat]))
                prod, price, cost_ratio = catalog[cat][sub][rng.integers(len(catalog[cat][sub]))]
                units = int(rng.choice([1, 1, 1, 2, 2, 3, 4])) if cat != "Electronics" else int(rng.choice([1, 1, 1, 2]))
                if seg != "Consumer":
                    units += int(rng.integers(0, 3))
                # discounts deeper in festive months and on marketplace
                disc_pool = [0, 0, 0.05, 0.10, 0.15, 0.20]
                if m in (10, 11):
                    disc_pool += [0.20, 0.25, 0.30]
                if ch == "Marketplace":
                    disc_pool += [0.15, 0.20]
                disc = float(rng.choice(disc_pool))
                unit_price = round(price * rng.uniform(0.97, 1.03))
                revenue = round(units * unit_price * (1 - disc), 2)
                cost = round(units * price * cost_ratio * rng.uniform(0.97, 1.03), 2)
                rows.append({
                    "Order ID": f"ORD-{oid}",
                    "Order Date": date.date(),
                    "Year": str(year),
                    "Month": m,
                    "Region": region,
                    "City": city,
                    "Segment": seg,
                    "Channel": ch,
                    "Category": cat,
                    "Sub-Category": sub,
                    "Product": prod,
                    "Units": units,
                    "Unit Price": unit_price,
                    "Discount": disc,
                    "Revenue": revenue,
                    "Cost": cost,
                    "Profit": round(revenue - cost, 2),
                })

df = pd.DataFrame(rows)
df.to_csv(ROOT / "data" / "sales_data.csv", index=False)
print(len(df), "lines,", df["Order ID"].nunique(), "orders")
print(df.groupby("Year")[["Revenue", "Profit"]].sum().round(0))
print("margin", (df.Profit.sum() / df.Revenue.sum()).round(3))
print(df.groupby("Category")[["Revenue", "Profit"]].sum().assign(m=lambda d: d.Profit / d.Revenue).round(2))
