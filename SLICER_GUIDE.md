# Build Guide: KPI Cards & Slicers

This guide walks through recreating the KPI section and slicers exactly as used in the Zepto Sales Analysis dashboard.

---

## 1. Prepare the Data

1. Load your raw sales data into a sheet (e.g. `RawData`).
2. Select the full range → **Insert → Table** (Ctrl+T) so it becomes a structured Table (e.g. `Tbl_Sales`).
3. Clean the data with **Power Query** if needed (remove duplicates, fix blanks, standardize City/Category names).

---

## 2. Build the PivotTables

1. Select the Table → **Insert → PivotTable → New Worksheet**.
2. Create **one PivotTable per visual** (Top 10 Orders, Top 10 Customers, Payment Mode, Customer vs Payment, Sales vs Delivery, etc.) — this makes it easy to connect the right slicers to the right charts later.
3. Rename each PivotTable (PivotTable Analyze → PivotTable Name) so they're easy to identify, e.g. `PT_TopOrders`, `PT_PaymentMode`.

---

## 3. Add the KPI Cards (below the title)

KPI cards are just formula-driven cells styled to look like tiles.

1. Create a small **helper area** (can be hidden or on a separate calc sheet) with formulas such as:
   ```excel
   Total Sales     = SUM(Tbl_Sales[Sales])
   Total Orders    = COUNTA(UNIQUE(Tbl_Sales[Order ID]))
   Total Customers = COUNTA(UNIQUE(Tbl_Sales[Customer ID]))
   Total Products  = COUNTA(UNIQUE(Tbl_Sales[Product ID]))
   Total Quantity  = SUM(Tbl_Sales[Quantity])
   Average Sales   = AVERAGE(Tbl_Sales[Sales])
   ```
   > If your Excel version doesn't support `UNIQUE`, use `SUMPRODUCT(1/COUNTIF(range,range))` instead.
2. Draw six rounded rectangles below the title (**Insert → Shapes → Rounded Rectangle**).
3. Click each shape, then click into the **Formula Bar** and type `=` followed by the cell reference (e.g. `=CalcSheet!B2`) to link the shape's text to the KPI cell. This makes the KPI update live.
4. Format each shape: fill color, white bold text, small caption label ("total sales", "total orders", etc.) using a text box or a second line inside the shape.
5. Optional: use custom number formatting (`₹#,##0.00,"K"` or `0.0,"k"`) so large numbers display as `563.33k` etc.

---

## 4. Insert Slicers (step-by-step)

1. Click anywhere inside a PivotTable.
2. Go to **PivotTable Analyze (or Insert) → Insert Slicer**.
3. In the dialog, check the fields you want as slicers, e.g.:
   - Category
   - City
   - Customer Segment
   - Customer Type
   - Payment Method
   - Delivery Status
4. Click **OK** — a slicer box appears for each selected field.
5. Arrange the slicers along the left side of the dashboard (as in the screenshot).

### Style a slicer
- Click the slicer → **Slicer → Options (ribbon) → Slicer Styles** → pick a style, or customize with **New Slicer Style**.
- Adjust **Columns** (Slicer settings) if you want multi-column layout.
- Resize by dragging the corner handles, or set exact height/width in **Size** on the ribbon.

### Connect one slicer to multiple PivotTables (critical step)
By default a slicer only filters the PivotTable it was created from. To make one slicer control *all* charts:

1. Right-click the slicer → **Report Connections** (Excel) / **PivotTable Connections**.
2. Check every PivotTable that should respond to this slicer (e.g. `PT_TopOrders`, `PT_TopCustomers`, `PT_PaymentMode`, `PT_CustomerVsPayment`, `PT_SalesVsDelivery`).
3. Click **OK**.
4. Repeat for each slicer (Category, City, Customer Segment, Customer Type, Payment Method, Delivery Status).

> Tip: All PivotTables you want linked together must be built from the **same data source** (same Table or same Data Model) for Report Connections to work.

---

## 5. Build the Charts

1. Click inside each PivotTable → **Insert → PivotChart**.
2. Choose chart type:
   - Pie → for "Total" (Sales by Payment Method)
   - Clustered Column → for Payment Mode, Payment vs Customer, Customer vs Payment, Sales vs Delivery
   - Bar (horizontal) → for Top 10 Orders / Top 10 Customers
3. For Top 10 charts: right-click the field in the **Filters/Rows** area → **Value Filters → Top 10** → set to Top 10 by Sum of Sales.
4. Remove field buttons and gridlines for a cleaner look: **PivotChart Analyze → Field Buttons → Hide All**.
5. Match your dashboard's color theme via **Chart Design → Change Colors**.

---

## 6. Assemble the Dashboard Sheet

1. Create a new sheet named `Dashboard`.
2. Cut/paste (or move) each PivotChart onto this sheet.
3. Add the KPI shapes at the top, and the title banner using a rounded rectangle shape with bold white text.
4. Position slicers on the left, charts in a grid on the right.
5. Go to **View → Uncheck Gridlines** and **Uncheck Headings** for a clean, presentation-style look.
6. Protect the sheet (optional) so users can only interact with slicers: **Review → Protect Sheet**, keep "Use PivotTable & PivotChart" and "Use AutoFilter" checked, uncheck editing of cells.

---

## 7. Final Checks

- [ ] Click each slicer and confirm every chart + KPI updates
- [ ] Clear all filters (Slicer → right-click → Clear Filter, or the eraser icon) to confirm the default/base view is correct
- [ ] Save as `.xlsx` (or `.xlsm` only if you added macros)
