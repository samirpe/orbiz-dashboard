import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from datetime import datetime, timedelta
import base64
import io
from PIL import Image
import os

st.set_page_config(layout="wide", page_title="Orbiz Sales Dashboard")

# Custom styling with Poppins font and improved headings
st.markdown("""
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Poppins:wght@400;600;700&display=swap');
        html, body, [class*="css"]  {
            font-family: 'Poppins', sans-serif;
        }
        .main-title {
            font-size: 42px;
            font-weight: 700;
            margin-bottom: 0.5rem;
            text-align: center;
        }
        .section-heading {
            font-size: 32px;
            font-weight: 600;
            margin-top: 2rem;
            margin-bottom: 1rem;
            border-bottom: 2px solid #eee;
            padding-bottom: 0.3rem;
        }
        th, td {
            text-align: right !important;
        }
    </style>
""", unsafe_allow_html=True)

# Display logo and title
logo_path = os.path.join(os.path.dirname(__file__), "Orbiz Logo.jpg")
logo = Image.open(logo_path)
col_logo, col_title, col_spacer = st.columns([1, 4, 1])
with col_logo:
    st.image(logo, width=300)
with col_title:
    st.markdown("<div class='main-title'>Sales & Fulfillment Dashboard</div>", unsafe_allow_html=True)

uploaded_file = st.file_uploader("Upload your Sales Order Excel file", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    df['Order Date'] = pd.to_datetime(df['Order Date'], errors='coerce').dt.date

    # Date filter - default to Month-to-date
    st.markdown("<div class='section-heading'>📅 Filter by Date</div>", unsafe_allow_html=True)
    today = datetime.today().date()
    month_start = datetime(today.year, today.month, 1).date()
    start_date, end_date = st.date_input("Select date range", [month_start, today])
    df = df[(df['Order Date'] >= start_date) & (df['Order Date'] <= end_date)]

    # Format currency to Indian system
    def format_inr(val):
        val = float(val)
        if val >= 1e7:
            return f"₹{val/1e7:.2f} Cr"
        elif val >= 1e5:
            return f"₹{val/1e5:.2f} Lakh"
        else:
            return f"₹{val:,.0f}"

    # KPI Summary
    total_sales = df['Total'].sum()
    total_paid = df['Paid'].sum()
    pending_to_pay = total_sales - total_paid

    st.markdown("<div class='section-heading'>💰 Payment Summary</div>", unsafe_allow_html=True)
    kpi1, spacer1, kpi2, spacer2, kpi3 = st.columns([1, 0.1, 1, 0.1, 1])
    kpi1.metric("Total Sales", format_inr(total_sales))
    kpi2.metric("Total Paid", format_inr(total_paid))
    kpi3.metric("Pending to Pay", format_inr(pending_to_pay))

    # Order Fulfillment
    total_orders = len(df)
    fully_dispatched = (df['Delivery Status'] == 'Fully Dispatched').sum()
    partially_dispatched = (df['Delivery Status'] == 'Partially Dispatched').sum()
    pending_dispatch = (df['Delivery Status'] == 'Not Dispatched').sum()

    st.markdown("<div class='section-heading'>📦 Order Fulfillment Overview</div>", unsafe_allow_html=True)
    o1, o2, o3, o4 = st.columns(4)
    o1.metric("Total Orders", total_orders)
    o2.metric("Fully Dispatched", fully_dispatched)
    o3.metric("Partially Dispatched", partially_dispatched)
    o4.metric("Pending Dispatch", pending_dispatch)

    # Salesperson Insights
    df['First Name'] = df['Salesperson'].str.split().str[0]
    df['Payment to Collect'] = df['Total'] - df['Paid']

    sales_by_person = df.groupby('First Name').agg({
        'Total': 'sum',
        'Paid': 'sum',
        'Payment to Collect': 'sum'
    }).sort_values(by='Total', ascending=False).reset_index()

    df_top_sales = sales_by_person.copy()
    df_top_sales['Total'] = df_top_sales['Total'].apply(format_inr)
    df_top_sales['Paid'] = df_top_sales['Paid'].apply(format_inr)

    df_outstanding = sales_by_person[['First Name', 'Total', 'Payment to Collect']].copy()
    df_outstanding['Total'] = df_outstanding['Total'].apply(format_inr)
    df_outstanding['Payment to Collect'] = df_outstanding['Payment to Collect'].apply(format_inr)

    st.markdown("<div class='section-heading'>👤 Salesperson Performance</div>", unsafe_allow_html=True)
    c1, c2, c3 = st.columns(3)
    with c1:
        st.markdown("**Top 5 Salesman**")
        st.dataframe(df_top_sales.head(5).style.set_properties(**{'text-align': 'right'}), use_container_width=True, hide_index=True)
    with c2:
        st.markdown("**Bottom 5 Salesman**")
        st.dataframe(df_top_sales.tail(5).style.set_properties(**{'text-align': 'right'}), use_container_width=True, hide_index=True)
    with c3:
        st.markdown("**Top 5 by Outstanding Payments**")
        st.dataframe(df_outstanding.sort_values(by='Payment to Collect', ascending=False).head(5).style.set_properties(**{'text-align': 'right'}), use_container_width=True, hide_index=True)

    # Customer Insights
    st.markdown("<div class='section-heading'>🧾 Customer Insights</div>", unsafe_allow_html=True)
    top_customers_by_value = df.groupby('Customer')['Total'].sum().sort_values(ascending=False).head(5).reset_index()
    top_customers_by_value['Sales Value'] = top_customers_by_value['Total'].apply(format_inr)
    top_customers_by_value.drop(columns='Total', inplace=True)

    top_customers_by_orders = df['Customer'].value_counts().head(5).reset_index()
    top_customers_by_orders.columns = ['Customer', 'Order Count']

    cu1, cu2 = st.columns(2)
    with cu1:
        st.markdown("**Top 5 Customers by Sales Value**")
        st.dataframe(top_customers_by_value.style.set_properties(**{'text-align': 'right'}), use_container_width=True, hide_index=True)
    with cu2:
        st.markdown("**Top 5 Customers by Order Count**")
        st.dataframe(top_customers_by_orders.style.set_properties(**{'text-align': 'right'}), use_container_width=True, hide_index=True)

    # Order Dispatch Check
    st.markdown("<div class='section-heading'>📋 Orders Older Than 3 Days & Not Fully Dispatched</div>", unsafe_allow_html=True)
    cutoff_date = today - timedelta(days=3)
    filtered_df = df[(df['Order Date'] <= cutoff_date) & (df['Delivery Status'] != 'Fully Dispatched')].copy()
    filtered_df = filtered_df.sort_values(by='Order Date', ascending=True)
    filtered_df['Order Date'] = pd.to_datetime(filtered_df['Order Date']).dt.strftime('%d-%b-%y')
    display_df = filtered_df[['Order Reference', 'Order Date', 'Customer', 'Salesperson', 'Delivery Status']]
    st.dataframe(display_df.style.set_properties(**{'text-align': 'right'}), use_container_width=True, hide_index=True)

    try:
        from fpdf import FPDF

        def create_pdf(dataframe):
            pdf = FPDF()
            pdf.add_page()
            pdf.set_font("Arial", size=12)
            pdf.cell(200, 10, txt="Orders Older Than 3 Days & Not Fully Dispatched", ln=True, align='C')
            pdf.ln(10)
            for index, row in dataframe.iterrows():
                pdf.cell(0, 10, txt=f"{row['Order Reference']} | {row['Order Date']} | {row['Customer']} | {row['Salesperson']} | {row['Delivery Status']}", ln=True)
            return pdf.output(dest='S').encode('latin-1')

        if not filtered_df.empty:
            pdf_bytes = create_pdf(display_df)
            b64 = base64.b64encode(pdf_bytes).decode()
            href = f'<a href="data:application/pdf;base64,{b64}" download="pending_orders.pdf">📅 Download as PDF</a>'
            st.markdown(href, unsafe_allow_html=True)

    except ModuleNotFoundError:
        st.warning("PDF export is disabled because the 'fpdf' package is not installed. To enable it, run: pip install fpdf")

    # Line chart for Sales over Time
    st.markdown("<div class='section-heading'>📈 Sales Trend Over Time</div>", unsafe_allow_html=True)
    sales_trend = df.groupby(df['Order Date'])['Total'].sum().reset_index()
    sales_trend.columns = ['Date', 'Total Sales']
    fig, ax = plt.subplots()
    ax.plot(sales_trend['Date'], sales_trend['Total Sales'], marker='o', linestyle='-', color='skyblue')
    ax.set_title('Sales vs Date')
    ax.set_xlabel('Date')
    ax.set_ylabel('Total Sales')
    plt.xticks(rotation=45)
    st.pyplot(fig)

else:
    st.info("📂 Please upload an Excel file to view the dashboard.")
