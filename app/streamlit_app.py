import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import os
from pathlib import Path

# Works on both Windows and Linux
BASE_DIR = Path(__file__).resolve().parent.parent

st.set_page_config(
    page_title="Churn Analytics Dashboard",
    page_icon="📊",
    layout="wide"
)

@st.cache_data
def load_data():
    df = pd.read_csv(BASE_DIR / 'outputs' / 'customers_scored.csv')
    return df

df = load_data()

# ── Sidebar ────────────────────────────────────────────────────────────────
st.sidebar.title("Filters")
selected_plans = st.sidebar.multiselect(
    "Plan", options=df['plan'].unique(),
    default=df['plan'].unique()
)
selected_contracts = st.sidebar.multiselect(
    "Contract Type", options=df['contract_type'].unique(),
    default=df['contract_type'].unique()
)
risk_threshold = st.sidebar.slider(
    "High Risk Threshold", 
    min_value=0.1, max_value=0.9, 
    value=0.5, step=0.05
)

filtered = df[
    (df['plan'].isin(selected_plans)) &
    (df['contract_type'].isin(selected_contracts))
]

# ── Header ─────────────────────────────────────────────────────────────────
st.title("Customer Churn Analytics")
st.caption("UK Telecoms · Predictive Retention Dashboard")
st.divider()

# ── Page selector ──────────────────────────────────────────────────────────
page = st.radio("", ["Overview", "Segment Explorer", 
                     "High-Risk Customers", "Model Insights"],
                horizontal=True)
st.divider()

# ══════════════════════════════════════════════════════════════════════════
# PAGE 1 — OVERVIEW
# ══════════════════════════════════════════════════════════════════════════
if page == "Overview":

    col1, col2, col3, col4, col5 = st.columns(5)

    total    = len(filtered)
    churned  = filtered['churned_1yes'].sum()
    churn_rt = churned / total * 100
    rev_risk = filtered[filtered['predicted_churn_prob'] > risk_threshold]['monthly_fee_'].sum()
    avg_nps  = filtered['nps_score_010'].mean()

    col1.metric("Total Customers",   f"{total:,}")
    col2.metric("Churned",           f"{int(churned):,}")
    col3.metric("Churn Rate",        f"{churn_rt:.1f}%")
    col4.metric("Revenue at Risk",   f"£{rev_risk:,.0f}/mo")
    col5.metric("Avg NPS Score",     f"{avg_nps:.1f}")

    st.subheader("Churn Rate by Plan")
    plan_churn = (filtered.groupby('plan')['churned_1yes']
                  .agg(['mean','count','sum'])
                  .reset_index())
    plan_churn.columns = ['Plan','Churn Rate','Total','Churned']
    plan_churn['Churn Rate'] = plan_churn['Churn Rate'] * 100

    fig = px.bar(plan_churn, x='Plan', y='Churn Rate',
                 color='Churn Rate',
                 color_continuous_scale=['#16A34A','#D97706','#DC2626'],
                 text=plan_churn['Churn Rate'].apply(lambda x: f"{x:.1f}%"))
    fig.update_traces(textposition='outside')
    fig.update_layout(showlegend=False, coloraxis_showscale=False,
                      plot_bgcolor='white', yaxis_title="Churn Rate (%)")
    st.plotly_chart(fig, width='stretch')

    col_a, col_b = st.columns(2)

    with col_a:
        st.subheader("Churn by Contract Type")
        contract_churn = (filtered.groupby('contract_type')['churned_1yes']
                          .mean().reset_index())
        contract_churn.columns = ['Contract','Churn Rate']
        contract_churn['Churn Rate'] *= 100
        fig2 = px.bar(contract_churn, x='Contract', y='Churn Rate',
                      color='Churn Rate',
                      color_continuous_scale=['#16A34A','#D97706','#DC2626'],
                      text=contract_churn['Churn Rate'].apply(lambda x: f"{x:.1f}%"))
        fig2.update_traces(textposition='outside')
        fig2.update_layout(showlegend=False, coloraxis_showscale=False,
                           plot_bgcolor='white', yaxis_title="Churn Rate (%)")
        st.plotly_chart(fig2, width='stretch')

    with col_b:
        st.subheader("Churn by Region")
        region_churn = (filtered.groupby('region')['churned_1yes']
                        .mean().reset_index())
        region_churn.columns = ['Region','Churn Rate']
        region_churn['Churn Rate'] *= 100
        region_churn = region_churn.sort_values('Churn Rate', ascending=True)
        fig3 = px.bar(region_churn, x='Churn Rate', y='Region',
                      orientation='h',
                      color='Churn Rate',
                      color_continuous_scale=['#16A34A','#D97706','#DC2626'])
        fig3.update_layout(showlegend=False, coloraxis_showscale=False,
                           plot_bgcolor='white', xaxis_title="Churn Rate (%)")
        st.plotly_chart(fig3, width='stretch')

# ══════════════════════════════════════════════════════════════════════════
# PAGE 2 — SEGMENT EXPLORER
# ══════════════════════════════════════════════════════════════════════════
elif page == "Segment Explorer":

    st.subheader("Customer Segment Analysis")

    col1, col2 = st.columns(2)

    with col1:
        segment_by = st.selectbox("Segment by", 
            ["plan", "contract_type", "region", 
             "age_band", "acquisition_channel", "gender"])

    with col2:
        metric = st.selectbox("Show metric",
            ["Churn Rate", "Customer Count", "Avg NPS"])

    seg = filtered.groupby(segment_by).agg(
        churn_rate=('churned_1yes','mean'),
        count=('churned_1yes','count'),
        avg_nps=('nps_score_010','mean')
    ).reset_index()
    seg['churn_rate'] *= 100

    metric_map = {
        "Churn Rate": ("churn_rate", "Churn Rate (%)"),
        "Customer Count": ("count", "Number of Customers"),
        "Avg NPS": ("avg_nps", "Average NPS Score")
    }
    y_col, y_label = metric_map[metric]
    seg = seg.sort_values(y_col, ascending=False)

    fig4 = px.bar(seg, x=segment_by, y=y_col,
                  text=seg[y_col].apply(lambda x: f"{x:.1f}"),
                  color=y_col,
                  color_continuous_scale=['#16A34A','#D97706','#DC2626'])
    fig4.update_traces(textposition='outside')
    fig4.update_layout(showlegend=False, coloraxis_showscale=False,
                       plot_bgcolor='white', yaxis_title=y_label)
    st.plotly_chart(fig4, width='stretch')

    st.subheader("Filtered Customer Data")
    st.dataframe(
        filtered[['customer_id','plan','contract_type','region',
                  'tenure_months','nps_score_010',
                  'predicted_churn_prob','churned_1yes']]
        .sort_values('predicted_churn_prob', ascending=False)
        .style.format({'predicted_churn_prob': '{:.1%}'}),
        width='stretch', height=400
    )

# ══════════════════════════════════════════════════════════════════════════
# PAGE 3 — HIGH RISK CUSTOMERS
# ══════════════════════════════════════════════════════════════════════════
elif page == "High-Risk Customers":

    high_risk = (filtered[filtered['predicted_churn_prob'] > risk_threshold]
                 .sort_values('monthly_fee_', ascending=False))

    st.subheader(f"High-Risk Customers  —  {len(high_risk)} flagged "
                 f"(threshold: {risk_threshold:.0%})")

    col1, col2, col3 = st.columns(3)
    col1.metric("High Risk Count",     f"{len(high_risk):,}")
    col2.metric("Monthly Revenue at Risk", f"£{high_risk['monthly_fee_'].sum():,.0f}")
    col3.metric("Avg Churn Probability",   
                f"{high_risk['predicted_churn_prob'].mean():.1%}")

    st.divider()

    display_cols = {
        'customer_id':           'Customer ID',
        'plan':                  'Plan',
        'contract_type':         'Contract',
        'region':                'Region',
        'tenure_months':         'Tenure (mo)',
        'nps_score_010':         'NPS',
        'complaints_raised':     'Complaints',
        'monthly_fee_':          'Monthly Fee (£)',
        'predicted_churn_prob':  'Churn Probability',
    }

    display_df = high_risk[list(display_cols.keys())].rename(columns=display_cols)

    st.dataframe(
        display_df.style.format({
            'Churn Probability': '{:.1%}',
            'Monthly Fee (£)': '£{:.2f}'
        }).background_gradient(
            subset=['Churn Probability'], 
            cmap='RdYlGn_r'
        ),
        width='stretch', height=450
    )

    csv = display_df.to_csv(index=False).encode('utf-8')
    st.download_button(
        label="Download High-Risk List (CSV)",
        data=csv,
        file_name='high_risk_customers.csv',
        mime='text/csv'
    )

# ══════════════════════════════════════════════════════════════════════════
# PAGE 4 — MODEL INSIGHTS
# ══════════════════════════════════════════════════════════════════════════
elif page == "Model Insights":

    st.subheader("Model Performance")

    col1, col2, col3 = st.columns(3)
    col1.metric("XGBoost AUC",          "0.784")
    col2.metric("Logistic Regression",  "0.753")
    col3.metric("Baseline (Random)",    "0.500")

    st.divider()

    col_a, col_b = st.columns(2)

    with col_a:
        st.subheader("Top Churn Predictors")
        st.image(str(BASE_DIR / 'outputs' / '07_feature_importance.png'))

    with col_b:
        st.subheader("Confusion Matrix")
        st.image(str(BASE_DIR / 'outputs' / '06_confusion_matrix.png'))

    st.divider()
    st.subheader("Key Findings")

    st.error("**Finding 1 — Early tenure is the highest risk window**  \n"
             "Customers in their first 6 months churn at 60.3% — nearly 3× the rate "
             "of established customers. Recommendation: implement a structured "
             "90-day onboarding sequence with a check-in call at day 30.")

    st.warning("**Finding 2 — Detractors on Month-to-Month contracts need immediate action**  \n"
               "This segment churns at 80%. At £689.75 monthly revenue at risk, "
               "a targeted retention call campaign would pay for itself if it retains "
               "even 2–3 customers.")

    st.success("**Finding 3 — Multi-product customers are significantly stickier**  \n"
               "Customers holding 3+ products churn at materially lower rates. "
               "A cross-sell campaign targeting single-product Basic plan customers "
               "at month 3 is the highest-ROI retention lever available.")