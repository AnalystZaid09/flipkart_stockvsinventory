import streamlit as st
import pandas as pd
from io import BytesIO
import warnings
warnings.filterwarnings("ignore", message="Workbook contains no default style, apply openpyxl's default")

st.set_page_config(page_title="Sales vs Inventory vs Returns Analysis", layout="wide")

# Title
st.title("📊 Sales vs Inventory vs Returns Analysis ")

# Sidebar for file uploads
st.sidebar.header("Upload Files")
sales_file = st.sidebar.file_uploader("Upload Sales Report (Excel/CSV)", type=['xlsx', 'xls', 'csv'])
pm_file = st.sidebar.file_uploader("Upload Product Master (Excel/CSV)", type=['xlsx', 'xls', 'csv'])
inventory_file = st.sidebar.file_uploader("Upload Inventory Report (Excel/CSV)", type=['xlsx', 'xls', 'csv'])
returns_file = st.sidebar.file_uploader("Upload Returns Report (Excel/CSV)", type=['xlsx', 'xls', 'csv'])

# Function to convert DataFrame to Excel for download
def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
    return output.getvalue()

def read_any(uploaded):
    name = uploaded.name.lower()
    if name.endswith('.csv'):
        return pd.read_csv(uploaded)
    else:
        return pd.read_excel(uploaded)

def find_column(columns, candidates):
    for c in candidates:
        if c in columns:
            return c
    lowered = {col.lower(): col for col in columns}
    for c in candidates:
        if c.lower() in lowered:
            return lowered[c.lower()]
    return None

if sales_file and pm_file and inventory_file and returns_file:
    try:
        # Load data
        with st.spinner("Loading data..."):
            SalesReport = read_any(sales_file)
            PM = read_any(pm_file)
            # Try inventory with header=1 (original), fallback to normal read
            try:
                inv_name = inventory_file.name.lower()
                if inv_name.endswith('.csv'):
                    inventory = read_any(inventory_file)
                else:
                    inventory = pd.read_excel(inventory_file, header=1)
            except Exception:
                inventory = read_any(inventory_file)
            Return = read_any(returns_file)

        st.success("✅ All files loaded successfully!")

        # Data Processing
        with st.spinner("Processing data..."):
            # Clean Sales Report
            if "Final Sale Units" not in SalesReport.columns:
                SalesReport["Final Sale Units"] = 0
            SalesReport["Final Sale Units"] = pd.to_numeric(SalesReport["Final Sale Units"], errors="coerce").fillna(0)
            SalesReport["Final Sale Units"] = SalesReport["Final Sale Units"].clip(lower=0)

            # Ensure Product Id exists in SalesReport
            if "Product Id" not in SalesReport.columns:
                alt = find_column(list(SalesReport.columns), ["FSN", "FNS", "ProductID", "Product Id", "Identifier"])
                if alt:
                    SalesReport = SalesReport.rename(columns={alt: "Product Id"})
                else:
                    SalesReport["Product Id"] = SalesReport.iloc[:, 0].astype(str)

            # Merge with Product Master (robust mapping)
            pm_cols = list(PM.columns)
            fns_col = find_column(pm_cols, ["FNS", "FSN", "fsn", "Product Id", "ProductID", "Identifier"]) or pm_cols[0]
            brand_col = find_column(pm_cols, ["Brand", "brand", "Brand Name", "BRAND", "brand_name"]) or (pm_cols[5] if len(pm_cols) > 5 else pm_cols[-1])
            brand_manager_col = find_column(pm_cols, ["Brand Manager", "Brand_Manager", "Brand Manager Name"]) or (pm_cols[4] if len(pm_cols) > 4 else brand_col)

            lookup = PM.rename(columns={
                fns_col: "FNS",
                brand_manager_col: "Brand Manager",
                brand_col: "Brand"
            })

            SalesReport = SalesReport.merge(
                lookup[["FNS", "Brand", "Brand Manager"]].drop_duplicates(subset=["FNS"]),
                left_on="Product Id",
                right_on="FNS",
                how="left"
            )

            # Ensure Brand & Brand Manager exist
            if "Brand" not in SalesReport.columns:
                SalesReport["Brand"] = "Unknown"
            else:
                SalesReport["Brand"] = SalesReport["Brand"].fillna("Unknown").astype(str)
            if "Brand Manager" not in SalesReport.columns:
                SalesReport["Brand Manager"] = ""

            # Reorder columns (safe)
            cols = list(SalesReport.columns)
            if "Brand" in cols and "Brand Manager" in cols and "Product Id" in cols:
                try:
                    cols.remove("Brand")
                    cols.remove("Brand Manager")
                    insert_at = cols.index("Product Id") + 1
                    new_cols = cols[:insert_at] + ["Brand", "Brand Manager"] + cols[insert_at:]
                    SalesReport = SalesReport[new_cols]
                except Exception:
                    pass

            # Create pivot sales (no margins; we'll add explicit Grand Total)
            pivot_sales = pd.pivot_table(
                SalesReport,
                index=["Brand", "Product Id"],
                values="Final Sale Units",
                aggfunc="sum",
                fill_value=0
            ).reset_index()

            # Append explicit Grand Total row
            pivot_sales = pd.concat([pivot_sales, pd.DataFrame([{
                "Brand": "Grand Total",
                "Product Id": "",
                "Final Sale Units": float(pivot_sales["Final Sale Units"].sum())
            }])], ignore_index=True)

            # Process Inventory
            inv_cols = list(inventory.columns)
            inv_id_col = find_column(inv_cols, ["Flipkart's Identifier of the product", "Identifier", "Product Id", "FSN", inv_cols[0]]) or inv_cols[0]
            inv_qty_col = find_column(inv_cols, ["Current stock count for your product", "Current stock count", "Inventory", "Quantity", inv_cols[1] if len(inv_cols) > 1 else inv_cols[-1]]) or inv_cols[-1]

            pivot_inventory = (
                inventory.groupby(inv_id_col, dropna=False)[inv_qty_col]
                .sum()
                .reset_index()
                .rename(columns={inv_id_col: "Product Id", inv_qty_col: "Inventory"})
            )

            pivot_inventory["Inventory"] = pd.to_numeric(pivot_inventory["Inventory"], errors="coerce").fillna(0)
            total_inv = int(pivot_inventory["Inventory"].sum())
            pivot_inventory = pd.concat([pivot_inventory, pd.DataFrame([{"Product Id": "Grand Total", "Inventory": total_inv}])], ignore_index=True)

            # Create Sales vs Inventory
            salesvsinventory = pivot_sales.merge(
                pivot_inventory[["Product Id", "Inventory"]],
                on="Product Id",
                how="left"
            )

            salesvsinventory["Inventory"] = pd.to_numeric(salesvsinventory["Inventory"], errors="coerce").fillna(0).astype(int)

            product_rows = salesvsinventory["Product Id"].notna() & (salesvsinventory["Product Id"].astype(str) != "")
            if product_rows.any():
                brand_inv_sum = salesvsinventory.loc[product_rows].groupby("Brand", dropna=False)["Inventory"].sum()
            else:
                brand_inv_sum = pd.Series(dtype="int64")

            brand_total_mask = salesvsinventory["Brand"].astype(str).str.endswith(" (Total)", na=False)
            if brand_total_mask.any():
                base_brand = salesvsinventory.loc[brand_total_mask, "Brand"].str.replace(" (Total)", "", regex=False)
                salesvsinventory.loc[brand_total_mask, "Inventory"] = base_brand.map(brand_inv_sum).fillna(0).astype(int)

            grand_mask = salesvsinventory["Brand"] == "Grand Total"
            if grand_mask.any():
                salesvsinventory.loc[grand_mask, "Inventory"] = int(brand_inv_sum.sum())
            salesvsinventory["Inventory"] = salesvsinventory["Inventory"].fillna(0).astype(int)

            # Process Returns
            if "Completion Status" in Return.columns:
                Return["Completion Status"] = (
                    Return["Completion Status"].astype(str)
                    .str.lower()
                    .replace({"delivered": "closed", "open": "in_transit"})
                )
            else:
                Return["Completion Status"] = "closed"

            if "Quantity" not in Return.columns:
                q_alt = find_column(list(Return.columns), ["Quantity", "Qty", "COUNT", "Count"])
                if q_alt:
                    Return = Return.rename(columns={q_alt: "Quantity"})
                else:
                    Return["Quantity"] = 0
            Return["Quantity"] = pd.to_numeric(Return["Quantity"], errors="coerce").fillna(0)

            if "FSN" not in Return.columns:
                fsn_alt = find_column(list(Return.columns), ["FSN", "FNS", "Product Id", "ProductID"])
                if fsn_alt:
                    Return = Return.rename(columns={fsn_alt: "FSN"})
                else:
                    Return["FSN"] = Return.get("Product Id", "").astype(str) if "Product Id" in Return.columns else ""

            returns_pivot = pd.pivot_table(
                Return,
                index="FSN",
                columns="Completion Status",
                values="Quantity",
                aggfunc="sum",
                fill_value=0
            ).reset_index()

            if returns_pivot.shape[1] > 1:
                returns_pivot["Grand Total"] = returns_pivot.iloc[:, 1:].sum(axis=1)
            else:
                returns_pivot["Grand Total"] = 0

            bottom_total = {"FSN": "Grand Total"}
            for col in returns_pivot.columns[1:]:
                bottom_total[col] = int(returns_pivot[col].sum())
            returns_pivot = pd.concat([returns_pivot, pd.DataFrame([bottom_total])], ignore_index=True)

            returns_pivot["FSN"] = returns_pivot["FSN"].astype(str).str.strip()
            closed_dict = dict(zip(returns_pivot["FSN"], returns_pivot.get("closed", [0] * len(returns_pivot))))
            transit_dict = dict(zip(returns_pivot["FSN"], returns_pivot.get("in_transit", [0] * len(returns_pivot))))

            salesvsinventory["closed"] = salesvsinventory["Product Id"].astype(str).map(closed_dict).fillna(0).astype(int)
            salesvsinventory["in_transit"] = salesvsinventory["Product Id"].astype(str).map(transit_dict).fillna(0).astype(int)

            mask_total = salesvsinventory["Brand"] == "Grand Total"
            if mask_total.any():
                totals = salesvsinventory.loc[~mask_total, ["Final Sale Units", "Inventory", "closed", "in_transit"]].sum(numeric_only=True)
                salesvsinventory.loc[mask_total, ["Final Sale Units", "Inventory", "closed", "in_transit"]] = totals.values

        st.success("✅ Data processing complete!")

        # Display Tabs (No Visualizations)
        tab1, tab2, tab3, tab4 = st.tabs(["📦 Sales Report", "🏪 Inventory", "↩️ Returns", "📊 Sales vs Inventory vs Return"])

        with tab1:
            st.header("Sales Report")
            st.dataframe(SalesReport, width='stretch', height=450)
            st.download_button("📥 Download Sales Report", to_excel(SalesReport), "sales_report.xlsx")

        with tab2:
            st.header("Inventory Report")
            st.dataframe(pivot_inventory, width='stretch', height=450)
            st.download_button("📥 Download Inventory Report", to_excel(pivot_inventory), "inventory.xlsx")

        with tab3:
            st.header("Returns Pivot")
            st.dataframe(returns_pivot, width='stretch', height=450)
            st.download_button("📥 Download Returns Pivot", to_excel(returns_pivot), "returns_pivot.xlsx")

        with tab4:
            st.header("Sales vs Inventory vs Returns Final Output")
            st.dataframe(salesvsinventory, width='stretch', height=450)
            st.download_button("📥 Download Sales vs Inventory vs Returns", to_excel(salesvsinventory), "sales_vs_inventory_vs_returns.xlsx")

    except Exception as e:
        st.error(f"❌ Error processing files: {str(e)}")
        st.exception(e)

else:
    st.info("👈 Please upload all required files in the sidebar to begin analysis.")
