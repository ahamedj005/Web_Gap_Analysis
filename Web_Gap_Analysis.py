import streamlit as st
import pandas as pd
import plotly.express as px
import os
import io

# ─── Page Config ─────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Business Analytics Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─── Custom CSS ──────────────────────────────────────────────────────────────
st.markdown("""
<style>
    .stApp { background-color: #0f172a; }
    section[data-testid="stSidebar"] {
        background: linear-gradient(180deg, #1e1b4b 0%, #312e81 100%);
    }
    section[data-testid="stSidebar"] * { color: #e0e7ff !important; }
    div[data-testid="metric-container"] {
        background: linear-gradient(135deg, #1e293b, #0f172a);
        border: 1px solid #334155;
        border-radius: 12px;
        padding: 16px;
        box-shadow: 0 4px 15px rgba(0,0,0,0.3);
    }
    div[data-testid="metric-container"] label {
        color: #94a3b8 !important; font-size: 13px !important;
    }
    div[data-testid="metric-container"] div[data-testid="stMetricValue"] {
        color: #f1f5f9 !important; font-size: 26px !important; font-weight: 700 !important;
    }
    .section-header {
        background: linear-gradient(90deg, #4f46e5, #7c3aed);
        color: white !important;
        padding: 10px 20px;
        border-radius: 8px;
        font-size: 18px;
        font-weight: 700;
        margin: 20px 0 12px 0;
    }
    button[data-baseweb="tab"] {
        font-size: 15px !important; font-weight: 600 !important; color: #94a3b8 !important;
    }
    button[data-baseweb="tab"][aria-selected="true"] {
        color: #818cf8 !important; border-bottom: 3px solid #818cf8 !important;
    }
    h1, h2, h3 { color: #e2e8f0 !important; }
    p, li { color: #cbd5e1 !important; }
    hr { border-color: #1e293b !important; }
</style>
""", unsafe_allow_html=True)

# ─── Chart Helpers ────────────────────────────────────────────────────────────
CHART_BG   = "#0f172a"
PAPER_BG   = "#1e293b"
FONT_COLOR = "#e2e8f0"
GRID_COLOR = "#334155"
PALETTE    = ["#818cf8","#34d399","#f472b6","#fbbf24","#60a5fa",
               "#a78bfa","#fb923c","#4ade80","#f87171","#38bdf8"]

def chart_layout(fig, height=380):
    fig.update_layout(
        plot_bgcolor=CHART_BG, paper_bgcolor=PAPER_BG,
        font=dict(color=FONT_COLOR, family="Inter"),
        height=height, margin=dict(l=30, r=30, t=40, b=30),
        legend=dict(bgcolor="rgba(0,0,0,0)", font=dict(color=FONT_COLOR)),
        xaxis=dict(gridcolor=GRID_COLOR, showgrid=True, linecolor=GRID_COLOR),
        yaxis=dict(gridcolor=GRID_COLOR, showgrid=True, linecolor=GRID_COLOR),
    )
    return fig

def section(title):
    st.markdown(f'<div class="section-header">{title}</div>', unsafe_allow_html=True)

# ─── Data Loader ──────────────────────────────────────────────────────────────
def load_data(bytes_data):
    """Load and clean both sheets from the Excel bytes."""
    buf = io.BytesIO(bytes_data)

    # ── Ecommerce ──────────────────────────────────────────────────────────
    ec_raw = pd.read_excel(buf, sheet_name="Ecommerce Orders", header=None)
    ec_df  = ec_raw.iloc[2:].copy()
    ec_df.columns = ec_raw.iloc[1].tolist()
    ec_df  = ec_df[ec_df["Order ID"].astype(str).str.startswith("ORD")].reset_index(drop=True)
    ec_df["Order Date"]     = pd.to_datetime(ec_df["Order Date"], errors="coerce")
    ec_df["Qty"]            = pd.to_numeric(ec_df["Qty"],            errors="coerce")
    ec_df["Unit Price (₹)"] = pd.to_numeric(ec_df["Unit Price (₹)"], errors="coerce")
    ec_df["Total (₹)"]      = pd.to_numeric(ec_df["Total (₹)"],      errors="coerce")

    # ── IT Solution ────────────────────────────────────────────────────────
    buf.seek(0)
    it_raw = pd.read_excel(buf, sheet_name="IT Solution Clients", header=None)
    it_df  = it_raw.iloc[2:].copy()
    it_df.columns = it_raw.iloc[1].tolist()
    it_df  = it_df[it_df["Client ID"].astype(str).str.startswith("CLT")].reset_index(drop=True)
    it_df["Start Date"] = pd.to_datetime(it_df["Start Date"], errors="coerce")
    it_df["End Date"]   = pd.to_datetime(it_df["End Date"],   errors="coerce")
    it_df["Budget (₹)"] = pd.to_numeric(it_df["Budget (₹)"], errors="coerce")

    return ec_df, it_df

# ─── Auto-detect OR upload ────────────────────────────────────────────────────
def find_local_file():
    script_dir = os.path.dirname(os.path.abspath(__file__))
    names = ["sample_dataset.xlsx", "Sample_Dataset.xlsx",
             "Sample_Datasets.xlsx", "dataset.xlsx"]
    for name in names:
        for folder in [script_dir, os.getcwd()]:
            p = os.path.join(folder, name)
            if os.path.exists(p):
                return p
    return None

# Try auto-load from disk first
local_path = find_local_file()

if local_path:
    with open(local_path, "rb") as f:
        file_bytes = f.read()
    st.sidebar.success(f"✅ File loaded: {os.path.basename(local_path)}")
    ec_df, it_df = load_data(file_bytes)

else:
    # Show uploader if file not found on disk
    st.markdown("## 📂 Upload Your Dataset")
    st.info("The Excel file was not found automatically. Please upload **sample_dataset.xlsx** below.")
    uploaded = st.file_uploader("Upload Excel File", type=["xlsx", "xls"])

    if uploaded is None:
        st.warning("👆 Please upload the Excel file to continue.")
        st.markdown("""
        **Required sheet names inside the Excel file:**
        - `Ecommerce Orders`
        - `IT Solution Clients`
        """)
        st.stop()

    file_bytes = uploaded.read()
    ec_df, it_df = load_data(file_bytes)

# ─── Sidebar Navigation ──────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## 📊 Dashboard Controls")
    st.markdown("---")
    page = st.radio("Navigate to", ["🏠 Overview", "🛒 Ecommerce", "💻 IT Solutions"])
    st.markdown("---")
    st.markdown("**📁 Data Sources**")
    st.markdown("🛒 [Ecommerce](https://hariniselvan-bot.github.io/ecom/)")
    st.markdown("💻 [IT Solution 1](https://mathumuthu.github.io/it-solution-26-03-2026/)")
    st.markdown("💻 [IT Solution 2](https://navaneethan733.github.io/IT---STACKLY/)")
    st.markdown("---")
    st.caption("Built with Streamlit + Plotly")


# ══════════════════════════════════════════════════════════════════════════════
# PAGE 1 — OVERVIEW
# ══════════════════════════════════════════════════════════════════════════════
if page == "🏠 Overview":
    st.markdown("# 📊 Business Analytics Dashboard")
    st.markdown("Combined insights across **Ecommerce** and **IT Solutions** platforms.")
    st.markdown("---")

    c1, c2, c3, c4, c5, c6 = st.columns(6)
    c1.metric("🛒 Total Orders",       len(ec_df))
    c2.metric("💰 Total Revenue",      f"₹{ec_df['Total (₹)'].sum():,.0f}")
    c3.metric("📦 Avg Order Value",    f"₹{ec_df['Total (₹)'].mean():,.0f}")
    c4.metric("💻 IT Clients",         len(it_df))
    c5.metric("🏦 Total IT Budget",    f"₹{it_df['Budget (₹)'].sum()/1e5:,.1f}L")
    c6.metric("✅ Completed Projects", len(it_df[it_df["Status"] == "Completed"]))

    st.markdown("---")
    col1, col2 = st.columns(2)

    with col1:
        section("🛒 Ecommerce — Revenue by Category")
        cat_rev = ec_df.groupby("Category")["Total (₹)"].sum().reset_index()
        fig = px.pie(cat_rev, names="Category", values="Total (₹)",
                     color_discrete_sequence=PALETTE, hole=0.45)
        fig.update_traces(textposition="outside", textinfo="percent+label",
                          textfont_color=FONT_COLOR)
        chart_layout(fig, 360)
        st.plotly_chart(fig, use_container_width=True)

    with col2:
        section("💻 IT Solutions — Budget by Service")
        svc_b = (it_df.groupby("Service Type")["Budget (₹)"].sum()
                 .reset_index().sort_values("Budget (₹)", ascending=True))
        fig = px.bar(svc_b, x="Budget (₹)", y="Service Type", orientation="h",
                     color="Budget (₹)", color_continuous_scale="Blues")
        fig.update_coloraxes(showscale=False)
        chart_layout(fig, 360)
        st.plotly_chart(fig, use_container_width=True)

    col3, col4 = st.columns(2)

    with col3:
        section("🛒 Order Status Breakdown")
        s = ec_df["Status"].value_counts().reset_index()
        s.columns = ["Status", "Count"]
        fig = px.bar(s, x="Status", y="Count", color="Status",
                     color_discrete_sequence=PALETTE, text="Count")
        fig.update_traces(textposition="outside")
        fig.update_layout(showlegend=False)
        chart_layout(fig, 320)
        st.plotly_chart(fig, use_container_width=True)

    with col4:
        section("💻 IT Project Status")
        si = it_df["Status"].value_counts().reset_index()
        si.columns = ["Status", "Count"]
        cmap = {"Completed": "#34d399", "In Progress": "#60a5fa", "Planned": "#fbbf24"}
        fig = px.pie(si, names="Status", values="Count",
                     color="Status", color_discrete_map=cmap, hole=0.5)
        fig.update_traces(textposition="outside", textinfo="percent+label",
                          textfont_color=FONT_COLOR)
        chart_layout(fig, 320)
        st.plotly_chart(fig, use_container_width=True)


# ══════════════════════════════════════════════════════════════════════════════
# PAGE 2 — ECOMMERCE
# ══════════════════════════════════════════════════════════════════════════════
elif page == "🛒 Ecommerce":
    st.markdown("# 🛒 Ecommerce Analytics")
    st.markdown("**Source:** [hariniselvan-bot.github.io/ecom](https://hariniselvan-bot.github.io/ecom/)")
    st.markdown("---")

    with st.expander("🔽 Filters", expanded=True):
        fc1, fc2, fc3 = st.columns(3)
        cats     = ["All"] + sorted(ec_df["Category"].dropna().unique().tolist())
        states   = ["All"] + sorted(ec_df["State"].dropna().unique().tolist())
        statuses = ["All"] + sorted(ec_df["Status"].dropna().unique().tolist())
        sel_cat    = fc1.selectbox("Category",     cats)
        sel_state  = fc2.selectbox("State",        states)
        sel_status = fc3.selectbox("Order Status", statuses)

    fdf = ec_df.copy()
    if sel_cat    != "All": fdf = fdf[fdf["Category"] == sel_cat]
    if sel_state  != "All": fdf = fdf[fdf["State"]    == sel_state]
    if sel_status != "All": fdf = fdf[fdf["Status"]   == sel_status]

    k1, k2, k3, k4, k5 = st.columns(5)
    k1.metric("📦 Orders",     len(fdf))
    k2.metric("💰 Revenue",    f"₹{fdf['Total (₹)'].sum():,.0f}")
    k3.metric("📊 Avg Value",  f"₹{fdf['Total (₹)'].mean():,.0f}" if len(fdf) else "—")
    k4.metric("🛍️ Items Sold", int(fdf["Qty"].sum()))
    k5.metric("✅ Delivered",   len(fdf[fdf["Status"] == "Delivered"]))

    st.markdown("---")
    tab1, tab2, tab3, tab4 = st.tabs(
        ["📈 Revenue Trends", "🗂️ Category Analysis", "🗺️ Geography", "📋 Orders Table"])

    with tab1:
        col1, col2 = st.columns(2)
        with col1:
            section("Monthly Revenue")
            mon = fdf.copy()
            mon["Month"] = mon["Order Date"].dt.to_period("M").astype(str)
            mon_rev = mon.groupby("Month")["Total (₹)"].sum().reset_index()
            fig = px.area(mon_rev, x="Month", y="Total (₹)",
                          color_discrete_sequence=["#818cf8"], markers=True)
            fig.update_traces(fill="tozeroy", fillcolor="rgba(129,140,248,0.15)")
            chart_layout(fig, 320)
            st.plotly_chart(fig, use_container_width=True)

        with col2:
            section("Revenue by Category")
            cat = (fdf.groupby("Category")["Total (₹)"].sum()
                   .reset_index().sort_values("Total (₹)", ascending=False))
            fig = px.bar(cat, x="Category", y="Total (₹)",
                         color="Category", color_discrete_sequence=PALETTE, text_auto=".2s")
            fig.update_traces(textposition="outside")
            fig.update_layout(showlegend=False)
            chart_layout(fig, 320)
            st.plotly_chart(fig, use_container_width=True)

        col3, col4 = st.columns(2)
        with col3:
            section("Order Status Distribution")
            s = fdf["Status"].value_counts().reset_index()
            s.columns = ["Status", "Count"]
            cmap = {"Delivered": "#34d399", "Shipped": "#60a5fa",
                    "Processing": "#fbbf24", "Cancelled": "#f87171", "Returned": "#fb923c"}
            fig = px.pie(s, names="Status", values="Count",
                         color="Status", color_discrete_map=cmap, hole=0.45)
            fig.update_traces(textposition="outside", textinfo="percent+label",
                              textfont_color=FONT_COLOR)
            chart_layout(fig, 330)
            st.plotly_chart(fig, use_container_width=True)

        with col4:
            section("Top 5 Products by Revenue")
            top5 = fdf.groupby("Product Name")["Total (₹)"].sum().nlargest(5).reset_index()
            fig = px.bar(top5, x="Total (₹)", y="Product Name", orientation="h",
                         color="Total (₹)", color_continuous_scale="Purples", text_auto=".2s")
            fig.update_coloraxes(showscale=False)
            chart_layout(fig, 330)
            st.plotly_chart(fig, use_container_width=True)

    with tab2:
        col1, col2 = st.columns(2)
        with col1:
            section("Qty Sold by Category")
            qty_cat = (fdf.groupby("Category")["Qty"].sum()
                       .reset_index().sort_values("Qty", ascending=True))
            fig = px.bar(qty_cat, x="Qty", y="Category", orientation="h",
                         color="Qty", color_continuous_scale="Teal", text="Qty")
            fig.update_coloraxes(showscale=False)
            chart_layout(fig, 360)
            st.plotly_chart(fig, use_container_width=True)

        with col2:
            section("Avg Unit Price by Category")
            avg_p = (fdf.groupby("Category")["Unit Price (₹)"].mean()
                     .reset_index().sort_values("Unit Price (₹)", ascending=False))
            fig = px.bar(avg_p, x="Category", y="Unit Price (₹)",
                         color="Category", color_discrete_sequence=PALETTE, text_auto=".2s")
            fig.update_traces(textposition="outside")
            fig.update_layout(showlegend=False)
            chart_layout(fig, 360)
            st.plotly_chart(fig, use_container_width=True)

        section("Category Summary Table")
        cat_sum = fdf.groupby("Category").agg(
            Orders=("Order ID", "count"),
            Total_Qty=("Qty", "sum"),
            Total_Revenue=("Total (₹)", "sum"),
            Avg_Price=("Unit Price (₹)", "mean")
        ).reset_index().sort_values("Total_Revenue", ascending=False)
        cat_sum.columns = ["Category", "Orders", "Total Qty",
                           "Total Revenue (₹)", "Avg Unit Price (₹)"]
        cat_sum["Total Revenue (₹)"]  = cat_sum["Total Revenue (₹)"].map("₹{:,.0f}".format)
        cat_sum["Avg Unit Price (₹)"] = cat_sum["Avg Unit Price (₹)"].map("₹{:,.0f}".format)
        st.dataframe(cat_sum, use_container_width=True, hide_index=True)

    with tab3:
        col1, col2 = st.columns(2)
        with col1:
            section("Revenue by State")
            state_rev = (fdf.groupby("State")["Total (₹)"].sum()
                         .reset_index().sort_values("Total (₹)", ascending=True))
            fig = px.bar(state_rev, x="Total (₹)", y="State", orientation="h",
                         color="Total (₹)", color_continuous_scale="Blues", text_auto=".2s")
            fig.update_coloraxes(showscale=False)
            chart_layout(fig, 380)
            st.plotly_chart(fig, use_container_width=True)

        with col2:
            section("Orders by City")
            city_o = fdf["City"].value_counts().reset_index()
            city_o.columns = ["City", "Orders"]
            fig = px.pie(city_o, names="City", values="Orders",
                         color_discrete_sequence=PALETTE, hole=0.4)
            fig.update_traces(textposition="outside", textinfo="percent+label",
                              textfont_color=FONT_COLOR)
            chart_layout(fig, 380)
            st.plotly_chart(fig, use_container_width=True)

    with tab4:
        section("📋 All Orders")
        show = fdf[["Order ID", "Order Date", "Customer Name", "City", "State",
                     "Category", "Product Name", "Qty",
                     "Unit Price (₹)", "Total (₹)", "Status"]].copy()
        show["Order Date"]     = show["Order Date"].dt.strftime("%Y-%m-%d")
        show["Total (₹)"]      = show["Total (₹)"].map("₹{:,.0f}".format)
        show["Unit Price (₹)"] = show["Unit Price (₹)"].map("₹{:,.0f}".format)
        st.dataframe(show, use_container_width=True, hide_index=True)
        st.download_button("⬇️ Download CSV",
                           fdf.to_csv(index=False).encode("utf-8"),
                           "ecommerce_orders.csv", "text/csv")


# ══════════════════════════════════════════════════════════════════════════════
# PAGE 3 — IT SOLUTIONS
# ══════════════════════════════════════════════════════════════════════════════
elif page == "💻 IT Solutions":
    st.markdown("# 💻 IT Solutions Analytics")
    st.markdown("**Sources:** [Mathumuthu IT](https://mathumuthu.github.io/it-solution-26-03-2026/) · "
                "[IT Stackly](https://navaneethan733.github.io/IT---STACKLY/)")
    st.markdown("---")

    with st.expander("🔽 Filters", expanded=True):
        fc1, fc2 = st.columns(2)
        svcs     = ["All"] + sorted(it_df["Service Type"].dropna().unique().tolist())
        sit_list = ["All"] + sorted(it_df["Status"].dropna().unique().tolist())
        sel_svc = fc1.selectbox("Service Type", svcs)
        sel_sit = fc2.selectbox("Status",        sit_list)

    fdf2 = it_df.copy()
    if sel_svc != "All": fdf2 = fdf2[fdf2["Service Type"] == sel_svc]
    if sel_sit != "All": fdf2 = fdf2[fdf2["Status"]       == sel_sit]

    k1, k2, k3, k4, k5 = st.columns(5)
    k1.metric("💼 Total Clients", len(fdf2))
    k2.metric("💰 Total Budget",  f"₹{fdf2['Budget (₹)'].sum()/1e5:,.1f}L")
    k3.metric("📊 Avg Budget",    f"₹{fdf2['Budget (₹)'].mean():,.0f}" if len(fdf2) else "—")
    k4.metric("✅ Completed",      len(fdf2[fdf2["Status"] == "Completed"]))
    k5.metric("🔄 In Progress",   len(fdf2[fdf2["Status"] == "In Progress"]))

    st.markdown("---")
    tab1, tab2, tab3 = st.tabs(
        ["📈 Budget Analysis", "📊 Service & Status", "📋 Client Table"])

    with tab1:
        col1, col2 = st.columns(2)
        with col1:
            section("Budget by Service Type")
            svc_b2 = (fdf2.groupby("Service Type")["Budget (₹)"].sum()
                      .reset_index().sort_values("Budget (₹)", ascending=False))
            fig = px.bar(svc_b2, x="Service Type", y="Budget (₹)",
                         color="Service Type", color_discrete_sequence=PALETTE, text_auto=".2s")
            fig.update_traces(textposition="outside")
            fig.update_layout(showlegend=False)
            chart_layout(fig, 340)
            st.plotly_chart(fig, use_container_width=True)

        with col2:
            section("Budget Distribution by Status")
            sb = fdf2.groupby("Status")["Budget (₹)"].sum().reset_index()
            cmap2 = {"Completed": "#34d399", "In Progress": "#60a5fa", "Planned": "#fbbf24"}
            fig = px.pie(sb, names="Status", values="Budget (₹)",
                         color="Status", color_discrete_map=cmap2, hole=0.45)
            fig.update_traces(textposition="outside", textinfo="percent+label",
                              textfont_color=FONT_COLOR)
            chart_layout(fig, 340)
            st.plotly_chart(fig, use_container_width=True)

        col3, col4 = st.columns(2)
        with col3:
            section("Top 5 Projects by Budget")
            top5_it = fdf2.nlargest(5, "Budget (₹)")[["Project Title", "Budget (₹)"]].copy()
            top5_it["Project Title"] = top5_it["Project Title"].str[:35] + "…"
            fig = px.bar(top5_it, x="Budget (₹)", y="Project Title", orientation="h",
                         color="Budget (₹)", color_continuous_scale="Purples", text_auto=".2s")
            fig.update_coloraxes(showscale=False)
            chart_layout(fig, 320)
            st.plotly_chart(fig, use_container_width=True)

        with col4:
            section("Clients by City")
            city_c = fdf2["City"].value_counts().reset_index()
            city_c.columns = ["City", "Clients"]
            fig = px.bar(city_c, x="City", y="Clients",
                         color="City", color_discrete_sequence=PALETTE, text="Clients")
            fig.update_traces(textposition="outside")
            fig.update_layout(showlegend=False)
            chart_layout(fig, 320)
            st.plotly_chart(fig, use_container_width=True)

    with tab2:
        col1, col2 = st.columns(2)
        with col1:
            section("Projects per Service Type")
            sc = fdf2["Service Type"].value_counts().reset_index()
            sc.columns = ["Service Type", "Count"]
            fig = px.pie(sc, names="Service Type", values="Count",
                         color_discrete_sequence=PALETTE, hole=0.4)
            fig.update_traces(textposition="outside", textinfo="percent+label",
                              textfont_color=FONT_COLOR)
            chart_layout(fig, 360)
            st.plotly_chart(fig, use_container_width=True)

        with col2:
            section("Status vs Service Type (Stacked)")
            cross = fdf2.groupby(["Service Type", "Status"]).size().reset_index(name="Count")
            cmap3 = {"Completed": "#34d399", "In Progress": "#60a5fa", "Planned": "#fbbf24"}
            fig = px.bar(cross, x="Service Type", y="Count", color="Status",
                         color_discrete_map=cmap3, barmode="stack", text="Count")
            fig.update_traces(textposition="inside")
            chart_layout(fig, 360)
            st.plotly_chart(fig, use_container_width=True)

        section("Service Type Budget Summary")
        svc_sum = fdf2.groupby("Service Type").agg(
            Clients=("Client ID", "count"),
            Total_Budget=("Budget (₹)", "sum"),
            Avg_Budget=("Budget (₹)", "mean"),
        ).reset_index().sort_values("Total_Budget", ascending=False)
        svc_sum.columns = ["Service Type", "Clients", "Total Budget (₹)", "Avg Budget (₹)"]
        svc_sum["Total Budget (₹)"] = svc_sum["Total Budget (₹)"].map("₹{:,.0f}".format)
        svc_sum["Avg Budget (₹)"]   = svc_sum["Avg Budget (₹)"].map("₹{:,.0f}".format)
        st.dataframe(svc_sum, use_container_width=True, hide_index=True)

    with tab3:
        section("📋 All Clients")
        show2 = fdf2[["Client ID", "Company Name", "Contact Person", "City",
                       "Service Type", "Project Title", "Start Date", "End Date",
                       "Budget (₹)", "Status"]].copy()
        show2["Start Date"] = show2["Start Date"].dt.strftime("%Y-%m-%d")
        show2["End Date"]   = show2["End Date"].dt.strftime("%Y-%m-%d")
        show2["Budget (₹)"] = show2["Budget (₹)"].map("₹{:,.0f}".format)
        st.dataframe(show2, use_container_width=True, hide_index=True)
        st.download_button("⬇️ Download CSV",
                           fdf2.to_csv(index=False).encode("utf-8"),
                           "it_solution_clients.csv", "text/csv")
        

