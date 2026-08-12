# zepto-sales-dashboard
An interactive **Excel PivotTable / PivotChart dashboard** analyzing sales performance for Zepto (quick-commerce) across categories, cities, customer segments, and payment methods.
!https://github.com/Jeshwanth999/zepto-sales-dashboard/blob/main/zepto.jpeg

---

## 📌 Overview

This dashboard gives a single-page, slicer-driven view of sales performance — built entirely with native Excel features (PivotTables, PivotCharts, Slicers, and linked KPI cards). No VBA or add-ins required.

---

## 🧮 KPIs (displayed below the title)

| KPI | Value |
|---|---|
| Total Sales | ₹563.33K |
| Total Orders | 1,499 |
| Total Customers | 301 |
| Total Products | 201 |
| Total Quantity | 4,199 |
| Average Sales | ₹0.376K |

Each KPI card is a **formula-linked cell** (using `SUM`, `COUNTA`/`DISTINCTCOUNT`, or `GETPIVOTDATA`) placed inside a styled shape/rectangle, so the numbers update automatically whenever a slicer is clicked.

---

## 📈 Charts / Visuals Used

| Visual | Type | Field(s) |
|---|---|---|
| Total | Pie Chart | Sum of Sales by Payment Method |
| Payment Mode | Column Chart | Sum of Sales by City, split by Customer Type |
| Payment vs Customer | Column Chart | Sum of Sales by Customer Type & Payment Method |
| Top 10 Orders | Bar Chart | Sum of Sales by Order ID (Top 10 filter) |
| Top 10 Customers | Bar Chart | Sum of Sales by Customer (Top 10 filter) |
| Customer vs Payment | Column Chart | Count of Customer Type by Payment Method |
| Sales vs Delivery | Column Chart | Count of Delivery Status by Customer Segment |

---

## 🎛️ Slicers Used

- **Category** (Baby, Bakery, Beverages, Diary, Essentials, Frozen…)
- **City** (Bangalore, Delhi, Mumbai, Not Available)
- **Customer Segment** (Bakery, Café, Cloud Kitchen, Household, Office, Tea Stall)

All slicers are connected to **every relevant PivotTable** on the sheet, so selecting one filter (e.g. City = Mumbai) updates the KPIs and all charts simultaneously.

Full step-by-step build instructions: **https://github.com/Jeshwanth999/zepto-sales-dashboard/blob/main/SLICER_GUIDE.md)**

---

## 🗂️ Repository Structure

```
zepto-sales-analysis-dashboard/
├── README.md

├── data/
│   └── zepto_sales_data.xlsx        # raw dataset (add your own)
├── dashboard/
│   └── Zepto_Sales_Analysis.xlsx    # final Excel dashboard file
├── assets/
│   └── dashboard-preview.jpeg       # screenshot used in this README
└── docs/
    └── SLICER_GUIDE.md              # step-by-step slicer & KPI build guide
```

---

## 🛠️ Tools Used

- Microsoft Excel (PivotTables, PivotCharts, Slicers, Power Query for data cleaning)
- Git & GitHub for version control

---

## 🚀 How to Use

1. Clone the repository:
   ```bash
   git clone https://github.com/<your-username>/zepto-sales-analysis-dashboard.git
   ```
2. Open `Zepto.xlsx` in Excel.
3. Click any slicer button to filter the entire dashboard.
4. Refer to `docs/SLICER_GUIDE.md` to rebuild the slicers/KPIs from scratch or adapt them to your own dataset.

---

## 📄 License

This project is licensed under the [MIT License](LICENSE).

---

## 🙋 Author

Maintained by **[Your Name]** — feel free to connect or raise an issue if you spot something to improve.

