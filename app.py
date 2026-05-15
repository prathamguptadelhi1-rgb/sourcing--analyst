import streamlit as st
import pandas as pd
import plotly.express as px
from io import BytesIO

st.set_page_config(page_title="Sourcing Controller AI", layout="wide")

st.title("🏭 Automated Sourcing Controller")
st.markdown("### Upload Purchase Register for 6-Lever Cost Analysis")

uploaded_file = st.file_uploader("Upload CSV/Excel", type=['csv', 'xlsx'])

if uploaded_file:
    # 1. Load Data (Skipping headers as per your file)
    df = pd.read_csv(uploaded_file, skiprows=3) if uploaded_file.name.endswith('.csv') else pd.read_excel(uploaded_file, skiprows=3)
    
    # Cleaning
    df['Value IN INR'] = pd.to_numeric(df['Value IN INR'], errors='coerce').fillna(0)
    df['INR Price'] = pd.to_numeric(df['INR Price'], errors='coerce').fillna(0)
    df['Quantity'] = pd.to_numeric(df['Quantity'], errors='coerce').fillna(0)

    # LEVER CALCULATIONS
    # Lever 3: Price Discovery (Negotiation Gap)
    price_gap = df.groupby('Item No.').agg({'INR Price': ['min', 'mean'], 'Quantity': 'sum'})
    price_gap.columns = ['Min_Price', 'Avg_Price', 'Total_Qty']
    price_gap['Savings'] = (price_gap['Avg_Price'] - price_gap['Min_Price']) * price_gap['Total_Qty']
    
    # Lever 6: Localization
    df['Origin'] = df['Country'].apply(lambda x: 'Domestic' if str(x).lower() == 'india' else 'Import')
    import_val = df[df['Origin'] == 'Import']['Value IN INR'].sum()

    # UI DASHBOARD
    st.header("📊 Consolidated Findings")
    c1, c2, c3 = st.columns(3)
    c1.metric("Total Spend", f"₹{df['Value IN INR'].sum():,.0f}")
    c2.metric("Savings Opportunity", f"₹{price_gap['Savings'].sum() + (import_val*0.05):,.0f}")
    c3.metric("Import Share", f"{(import_val/df['Value IN INR'].sum())*100:.1f}%")

    st.subheader("💡 Recommendations")
    st.info(f"1. **Price Discovery:** Negotiate with vendors for top SKUs to save ₹{price_gap['Savings'].sum():,.0f}.\n"
            f"2. **Localization:** Move high-value imports to India to save approx. ₹{import_val*0.05:,.0f}.\n"
            f"3. **SOB Allocation:** Shift 15% volume to the 'Min Price' vendors identified in the report.")

    # DOWNLOADABLE EXCEL
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name='Raw Data', index=False)
        price_gap.to_excel(writer, sheet_name='Price Discovery Analysis')
        
    st.download_button("📥 Download Full Analysis Report", output.getvalue(), "Sourcing_Analysis.xlsx")