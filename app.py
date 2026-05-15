import streamlit as st
import pandas as pd
from io import BytesIO
import plotly.express as px

# --- PAGE SETUP ---
st.set_page_config(page_title="Sourcing Controller AI", layout="wide")
st.title("🏭 Automated Sourcing Dashboard")
st.markdown("### Upload Purchase Register for 5-Sheet Detailed Category Analysis")

uploaded_file = st.file_uploader("Upload CSV/Excel", type=['csv', 'xlsx'])

if uploaded_file:
    # 1. Load Data (Skipping headers as per your file)
    df = pd.read_csv(uploaded_file, skiprows=3) if uploaded_file.name.endswith('.csv') else pd.read_excel(uploaded_file, skiprows=3)
    
    # Cleaning
    df['Value IN INR'] = pd.to_numeric(df['Value IN INR'], errors='coerce').fillna(0)
    df['INR Price'] = pd.to_numeric(df['INR Price'], errors='coerce').fillna(0)
    df['Quantity'] = pd.to_numeric(df['Quantity'], errors='coerce').fillna(0)

    # UI DASHBOARD METRICS
    st.header("📊 High-Level Metrics")
    c1, c2, c3 = st.columns(3)
    c1.metric("Total Spend", f"₹{df['Value IN INR'].sum():,.0f}")
    c2.metric("Total Quantity Purchased", f"{df['Quantity'].sum():,.0f}")
    c3.metric("Total Vendors", df['Vendor Name'].nunique())

    # --- THE 5 SPECIFIC LEVER CALCULATIONS ---
    
    # 1) Category vs Value Spent
    req1 = df.groupby('Category Description')['Value IN INR'].sum().reset_index()
    req1 = req1.sort_values(by='Value IN INR', ascending=False)
    
    # 2) Category vs Country (India/Domestic vs Import Percentage)
    df['Origin'] = df['Country'].apply(lambda x: 'Domestic (India)' if str(x).strip().lower() == 'india' else 'Import')
    req2 = df.pivot_table(index='Category Description', columns='Origin', values='Value IN INR', aggfunc='sum').fillna(0)
    req2['Total Value'] = req2.sum(axis=1)
    if 'Domestic (India)' in req2.columns:
        req2['Domestic %'] = (req2['Domestic (India)'] / req2['Total Value']) * 100
    if 'Import' in req2.columns:
        req2['Import %'] = (req2['Import'] / req2['Total Value']) * 100
    req2 = req2.reset_index()

    # 3) Category vs Average Price & Total Value IN INR
    req3 = df.groupby('Category Description').agg({'INR Price': 'mean', 'Value IN INR': 'sum'}).reset_index()
    req3.rename(columns={'INR Price': 'Average Price (INR)', 'Value IN INR': 'Total Value (INR)'}, inplace=True)
    
    # 4) Vendors for each category (Desc by Quantity, Value, Avg Price)
    req4 = df.groupby(['Category Description', 'Vendor Name']).agg({
        'Quantity': 'sum', 
        'Value IN INR': 'sum', 
        'INR Price': 'mean'
    }).reset_index()
    req4.rename(columns={'INR Price': 'Avg Price Per Unit/KG'}, inplace=True)
    # Sorting by Category Name (A-Z), then Quantity (High-Low), then Value (High-Low)
    req4 = req4.sort_values(by=['Category Description', 'Quantity', 'Value IN INR'], ascending=[True, False, False])
    
    # 5) RM vs BP (Vendor Count, Total Value, Avg Purchase per Vendor)
    req5 = df.groupby('RM Category').agg({
        'Vendor Name': 'nunique',
        'Value IN INR': 'sum'
    }).reset_index()
    req5.rename(columns={'Vendor Name': 'Vendor Count', 'Value IN INR': 'Total Value of Purchases'}, inplace=True)
    req5['Avg Purchase per Vendor (INR)'] = req5['Total Value of Purchases'] / req5['Vendor Count']

    # --- DOWNLOADABLE EXCEL GENERATION (5 SHEETS) ---
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        req1.to_excel(writer, sheet_name='1. Category vs Value', index=False)
        req2.to_excel(writer, sheet_name='2. Import-Domestic Split', index=False)
        req3.to_excel(writer, sheet_name='3. Category Avg Price & Value', index=False)
        req4.to_excel(writer, sheet_name='4. Vendor Category Ranking', index=False)
        req5.to_excel(writer, sheet_name='5. RM vs BP Analysis', index=False)
        
        # Auto-adjust column widths for better readability
        for sheet_name in writer.sheets:
            worksheet = writer.sheets[sheet_name]
            worksheet.set_column('A:Z', 22) # Makes columns wider automatically

    st.success("✅ Your 5-Sheet Excel Analysis is ready!")
    st.download_button("📥 Download 5-Sheet Analysis Report", output.getvalue(), "Sourcing_Detailed_Analysis.xlsx", type="primary")

    # Show a preview of the biggest sheet in the UI
    st.subheader("Preview: Top Vendors by Category (Sheet 4)")
    st.dataframe(req4.head(15), use_container_width=True)
