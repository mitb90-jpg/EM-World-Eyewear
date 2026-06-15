import streamlit as st
import pandas as pd
import io

# ---------------- PAGE CONFIG ----------------
st.set_page_config(
    page_title="Smart Transaction Categorizer",
    page_icon="📊",
    layout="wide"
)

# ---------------- CSS ----------------
st.markdown("""
<style>
.export-box {
    background-color: #1f4e79;
    padding: 18px;
    border-radius: 12px;
    text-align: center;
    margin-top: 15px;
    margin-bottom: 10px;
}

.export-title {
    color: white;
    font-size: 18px;
    font-weight: bold;
}

div.stDownloadButton > button {
    background-color: white;
    color: #1f4e79;
    font-weight: bold;
    border-radius: 8px;
    padding: 0.6em 1.2em;
}
div.stDownloadButton > button:hover {
    background-color: #e6f0ff;
}
</style>
""", unsafe_allow_html=True)

# ---------------- HEADER ----------------
col1, col2 = st.columns([1, 6])

with col1:
    st.image("Logo.jpeg", width=100)

with col2:
    st.markdown("<h1 style='color:#1f4e79;'>Prime Accounting and Tax</h1>", unsafe_allow_html=True)
    st.markdown("<p style='font-size:24px;color:gray;'>World Eyewear</p>", unsafe_allow_html=True)

# ---------------- SIDEBAR ----------------
company = st.sidebar.selectbox(
    "Select Account",
    ["Scotia Bank", "Triangle Master Card", "Visa - 6023", "Visa - 7866"]
)

uploaded_file = st.sidebar.file_uploader("Upload Excel File", type=["xlsx"])

# ================= MAIN =================
if uploaded_file is not None:

    df = pd.read_excel(uploaded_file)

    # ---------------- CLEAN ----------------
    df.columns = df.columns.astype(str).str.strip()
    df = df.loc[:, ~df.columns.str.contains("^Unnamed", na=False)]
    df = df.dropna(how="all")

    if "Date" in df.columns:
        df["Date"] = pd.to_datetime(df["Date"], errors="coerce").dt.date

    df["Category"] = ""

    # ---------------- RULES ----------------
    df.loc[df["Credit"].notna() &
           df["Description"].astype(str).str.contains("DEPOSIT|TRANSFER|MISC PAYMENT", case=False, na=False),
           "Category"] = "Revenue"

    df.loc[df["Debit"].notna() &
           df["Description"].astype(str).str.contains("INSURANCE", case=False, na=False),
           "Category"] = "Insurance"

    df.loc[df["Debit"].notna() &
           df["Description"].astype(str).str.strip().str.lower().eq("misc payment"),
           "Category"] = "Misc Expenses"

    df.loc[df["Debit"].notna() &
           df["Description"].astype(str).str.contains("LOAN", case=False, na=False),
           "Category"] = "Car Loan"

    df.loc[df["Debit"].notna() &
           df["Description"].astype(str).str.contains("HIGHWAY", case=False, na=False),
           "Category"] = "Parking and Toll"

    # ---------------- CLEAN DISPLAY ----------------
    df = df.reset_index(drop=True)
    df.insert(0, "Sr. No", range(1, len(df) + 1))

    st.subheader("📊 Categorized Transactions")
    st.dataframe(df, use_container_width=True, hide_index=True)

    # ---------------- AMOUNTS (WITH COMMAS DISPLAY) ----------------
    revenue = df.loc[df["Category"] == "Revenue", "Credit"].fillna(0).sum()
    insurance = df.loc[df["Category"] == "Insurance", "Debit"].fillna(0).sum()
    misc = df.loc[df["Category"] == "Misc Expenses", "Debit"].fillna(0).sum()
    loan = df.loc[df["Category"] == "Car Loan", "Debit"].fillna(0).sum()
    toll = df.loc[df["Category"] == "Parking and Toll", "Debit"].fillna(0).sum()

    amounts = {
        "Revenue": revenue,
        "Insurance": insurance,
        "Misc Expenses": misc,
        "Car Loan": loan,
        "Parking and Toll": toll
    }

    summary_df = pd.DataFrame({
        "Category": amounts.keys(),
        "Amount": [f"{v:,.2f}" for v in amounts.values()]
    })

    st.subheader("📋 Category Summary")
    st.dataframe(summary_df, use_container_width=True, hide_index=True)

    # ---------------- PROFIT & LOSS ----------------
    expense_total = insurance + misc + loan + toll
    net_profit = revenue - expense_total

    pl_df = pd.DataFrame([
        ["Revenue", revenue],
        ["Total Expenses", expense_total],
        ["Net Profit", net_profit]
    ], columns=["Description", "Amount"])

    pl_display = pl_df.copy()
    pl_display["Amount"] = pl_display["Amount"].apply(lambda x: f"{x:,.2f}")

    st.subheader("📊 Profit & Loss Statement")
    st.dataframe(pl_display, use_container_width=True, hide_index=True)

    # ---------------- DOWNLOAD P&L (FIXED) ----------------
    pl_output = io.BytesIO()
    with pd.ExcelWriter(pl_output, engine="openpyxl") as writer:
        pl_df.to_excel(writer, index=False, sheet_name="P&L")
    pl_output.seek(0)

    st.markdown("""
    <div class="export-box">
        <div class="export-title">📊 Export Profit & Loss Statement</div>
    </div>
    """, unsafe_allow_html=True)

    st.download_button(
        "⬇️ Export P&L Excel",
        data=pl_output,
        file_name="Profit_and_Loss.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # ---------------- DOWNLOAD MAIN FILE ----------------
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Transactions")
    output.seek(0)

    st.markdown("""
    <div class="export-box">
        <div class="export-title">📤 Export Categorized Data</div>
    </div>
    """, unsafe_allow_html=True)

    st.download_button(
        "⬇️ Export Excel File",
        data=output,
        file_name="Categorized_Data.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

else:
    st.info("Please upload an Excel file to begin.")
