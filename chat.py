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
sales_file = st.sidebar.file_uploader("Upload Sales Report (Excel/CSV)", type=['xlsx', 'xls','csv'])
pm_file = st.sidebar.file_uploader("Upload Product Master (Excel/CSV)", type=['xlsx', 'xls','csv'])
inventory_file = st.sidebar.file_uploader("Upload Inventory Report (Excel/XLS/CSV)", type=['xlsx', 'xls','csv'])
returns_file = st.sidebar.file_uploader("Upload Returns Report (Excel/CSV)", type=['xlsx', 'xls','csv'])

# Function to convert DataFrame to Excel for download
def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
    return output.getvalue()

def load_file(file, header=None, sheet_name=None):
    if file.name.endswith('.csv'):
        return pd.read_csv(file, header=header)
    else:
        return pd.read_excel(file, header=header, sheet_name=sheet_name)

if sales_file and pm_file and inventory_file and returns_file:
    try:
        # Load data
        with st.spinner("Loading data..."):
            SalesReport = load_file(sales_file)
            PM = load_file(pm_file)
            inventory = load_file(inventory_file, header=1)
            Return = load_file(returns_file, sheet_name='Sheet1')
        
        st.success("✅ All files loaded successfully!")
        
        # Data Processing
        with st.spinner("Processing data..."):
            # Clean Sales Report
            SalesReport['Final Sale Units'] = SalesReport['Final Sale Units'].apply(lambda x: 0 if x < 0 else x)
            
            if "Brand" in SalesReport.columns:
                SalesReport.rename(columns={"Brand": "Brand1"}, inplace=True)
            
            # Merge with Product Master
            lookup = PM.rename(columns={
                PM.columns[0]: "FNS",
                PM.columns[4]: "Brand Manager",
                PM.columns[5]: "Brand"
            })
            
            SalesReport = SalesReport.merge(
                lookup[["FNS", "Brand", "Brand Manager"]], 
                left_on="Product Id", 
                right_on="FNS", 
                how="left"
            )
            
            # Reorder columns
            cols = list(SalesReport.columns)
            cols.remove("Brand")
            cols.remove("Brand Manager")
            insert_at = cols.index("Product Id") + 1
            new_cols = cols[:insert_at] + ["Brand", "Brand Manager"] + cols[insert_at:]
            SalesReport = SalesReport[new_cols]
            
            # Create pivot sales
            pivot_sales = pd.pivot_table(
                SalesReport,
                index=["Brand", "Product Id"],
                values="Final Sale Units",
                aggfunc="sum",
                fill_value=0,
                margins=True,
                margins_name="Grand Total"
            ).reset_index()
            
            # Process Inventory
            pivot_inventory = (
                inventory.groupby("Flipkart's Identifier of the product")["Current stock count for your product"]
                .sum()
                .reset_index()
            )
            
            grand_total = pd.DataFrame({
                "Flipkart's Identifier of the product": ["Grand Total"],
                "Current stock count for your product": [pivot_inventory["Current stock count for your product"].sum()]
            })
            
            inventory_pivot = pd.concat([pivot_inventory, grand_total], ignore_index=True)
            
            # Create Sales vs Inventory
            salesvsinventory = pivot_sales.copy()
            
            inventory_lookup = inventory_pivot.rename(columns={
                "Flipkart's Identifier of the product": "Product Id",
                "Current stock count for your product": "Inventory"
            })
            
            salesvsinventory = salesvsinventory.merge(
                inventory_lookup[["Product Id", "Inventory"]],
                on="Product Id",
                how="left"
            )
            
            salesvsinventory["Inventory"] = pd.to_numeric(salesvsinventory["Inventory"], errors="coerce")
            
            product_rows = salesvsinventory["Product Id"] != ""
            brand_inv_sum = salesvsinventory[product_rows].groupby("Brand")["Inventory"].sum()
            
            brand_total_mask = salesvsinventory["Brand"].str.endswith(" (Total)", na=False)
            if brand_total_mask.any():
                base_brand = salesvsinventory.loc[brand_total_mask, "Brand"].str.replace(" (Total)", "", regex=False)
                salesvsinventory.loc[brand_total_mask, "Inventory"] = base_brand.map(brand_inv_sum)
            
            grand_mask = salesvsinventory["Brand"] == "Grand Total"
            salesvsinventory.loc[grand_mask, "Inventory"] = brand_inv_sum.sum()
            salesvsinventory["Inventory"] = salesvsinventory["Inventory"].fillna(0).astype(int)
            
            # Process Returns
            Return["Completion Status"] = (
                Return["Completion Status"]
                .str.lower()
                .replace({"delivered": "closed", "open": "in_transit"})
            )
            
            returns_pivot = pd.pivot_table(
                Return,
                index="FSN",
                columns="Completion Status",
                values="Quantity",
                aggfunc="sum",
                fill_value=0
            ).reset_index()
            
            returns_pivot["Grand Total"] = returns_pivot.iloc[:, 1:].sum(axis=1)
            
            bottom_total = {"FSN": "Grand Total"}
            for col in returns_pivot.columns[1:]:
                bottom_total[col] = returns_pivot[col].sum()
            
            returns_pivot = pd.concat([returns_pivot, pd.DataFrame([bottom_total])], ignore_index=True)
            
            # Map returns to sales vs inventory
            returns_pivot["FSN"] = returns_pivot["FSN"].astype(str).str.strip()
            closed_dict = dict(zip(returns_pivot["FSN"], returns_pivot.get("closed", [0]*len(returns_pivot))))
            transit_dict = dict(zip(returns_pivot["FSN"], returns_pivot.get("in_transit", [0]*len(returns_pivot))))
            
            salesvsinventory["closed"] = salesvsinventory["Product Id"].map(closed_dict).fillna(0).astype(int)
            salesvsinventory["in_transit"] = salesvsinventory["Product Id"].map(transit_dict).fillna(0).astype(int)
            
            # Update Grand Total
            mask_total = salesvsinventory["Brand"] == "Grand Total"
            totals = salesvsinventory.loc[~mask_total, ["Final Sale Units", "Inventory", "closed", "in_transit"]].sum()
            salesvsinventory.loc[mask_total, ["Final Sale Units", "Inventory", "closed", "in_transit"]] = totals.values
        
        st.success("✅ Data processing complete!")
        
        # Display Tabs (No Visualizations)
        tab1, tab2, tab3, tab4 = st.tabs(["📦 Sales Report", "🏪 Inventory", "↩️ Returns", "📊 Sales vs Inventory vs Return"])
        
        with tab1:
            st.header("Sales Report")
            st.dataframe(SalesReport, width="stretch", height=450)
            st.download_button("📥 Download Sales Report", to_excel(SalesReport), "sales_report.xlsx")
        
        with tab2:
            st.header("Inventory Report")
            st.dataframe(inventory_pivot, width="stretch", height=450)
            st.download_button("📥 Download Inventory Report", to_excel(inventory_pivot), "inventory.xlsx")
        
        with tab3:
            st.header("Returns Pivot")
            st.dataframe(returns_pivot, width="stretch", height=450)
            st.download_button("📥 Download Returns Pivot", to_excel(returns_pivot), "returns_pivot.xlsx")
        
        with tab4:
            st.header("Sales vs Inventory vs Returns Final Output")
            st.dataframe(salesvsinventory, width="stretch", height=450)
            st.download_button("📥 Download Sales vs Inventory vs Returns", to_excel(salesvsinventory), "sales_vs_inventory_vs_returns.xlsx")
            
    except Exception as e:
        st.error(f"❌ Error processing files: {str(e)}")
        st.exception(e)

else:
    st.info("👈 Please upload all required files in the sidebar to begin analysis.")


