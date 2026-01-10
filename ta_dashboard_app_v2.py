import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from datetime import datetime
import io

# -------------------------------------------------
# Page Config
# -------------------------------------------------
st.set_page_config(
    page_title="LTE Optimization",
    layout="wide",
    initial_sidebar_state="expanded"
)

# -------------------------------------------------
# Custom CSS
# -------------------------------------------------
st.markdown("""
<style>
    .metric-card { background-color: #f0f2f6; padding: 15px; border-radius: 5px; border-left: 4px solid #1f77b4; }
    .alarm-critical { border-left-color: #d62728 !important; }
    .alarm-warning { border-left-color: #ff7f0e !important; }
    .alarm-normal { border-left-color: #2ca02c !important; }
    .stTabs [data-baseweb="tab-list"] button [data-testid="stMarkdownContainer"] p {
        font-size: 16px; font-weight: 600;
    }
</style>
""", unsafe_allow_html=True)

# -------------------------------------------------
# Helper Functions
# -------------------------------------------------

def load_data(uploaded_file):
    """Load and process complete LTE KPI data"""
    try:
        if uploaded_file.name.endswith(".csv"):
            df_raw = pd.read_csv(uploaded_file)
        else:
            df_raw = pd.read_excel(uploaded_file, sheet_name="PRS DATA", engine='openpyxl')
        
        # Map columns
        df_raw['Site_ID'] = df_raw['eNodeB Name']
        df_raw['Cell_ID'] = df_raw['Cell Name']
        df_raw['Date'] = pd.to_datetime(df_raw['Date : Time'], format='%d/%m/%Y', errors='coerce')
        
        # Convert all numeric columns to proper numeric types
        numeric_columns = [
            'LTE Network Availability (%)', 'Cell Downtime with SON(min)', 'Cell Downtime(min)',
            'UL Interference(dBm)', 'PDSCH IBLER(%)', 'PUSCH IBLER(%)',
            'Traffic User(Avg)', 'Traffic User(Max)', 'VoLTE User',
            'DL Throughput(Mbit/s)', 'DL Traffic Volume(GB)', 'DL PRB Utilization(%)',
            'UL Throughput(Mbit/s)', 'UL Traffic Volume(GB)', 'UL PRB Utilization(%)',
            'VoLTE CSSR(%)', 'VoLTE DCR(%)', 'SRVCC SR(%)',
            'ERAB CSSR(%)', 'ERAB DCR(%)', 'RRC CSSR(%)', 'HO SR(%)', 'CSFB SR(%)',
            'VoLTE Traffic (Erl)', 'Avg CQI', 'Avg TA Distance(m)',
            'RRC Redirection E2G', 'RRC Redirection E2G (Blind)',
            'CSFB Attempt E2G', 'CSFB Attempt E2G (Flash)',
            'Smart Carrier Feature', 'Paging Discarded', 'MIMO Rank2',
            'VoLTE Drop due Radio', 'VoLTE Drop due Congestion',
            'VoLTE Drop due TNL', 'VoLTE Drop due MME', 'VoLTE Drop due EUtranGen'
        ]
        
        # TA columns
        ta_columns = [
            'TA (0 -78m)', 'TA (78m - 234m)', 'TA (234m - 546m)', 'TA (546m - 1014m)',
            'TA (1014m-1950m)', 'TA (1950m - 3510m)', 'TA (3510m - 6630m)',
            'TA (6630m-14430m)', 'TA (14430m - 30030m)', 'TA (30030m - 53430m)',
            'TA (53430m - 76830m)', 'TA (>76830m)'
        ]
        
        # Convert all numeric columns
        for col in numeric_columns + ta_columns:
            if col in df_raw.columns:
                df_raw[col] = pd.to_numeric(df_raw[col], errors='coerce')
        
        return df_raw
        
    except Exception as e:
        st.error(f"Error loading data: {str(e)}")
        return None

def process_ta_data(df_raw):
    """Process TA data for coverage analysis"""
    ta_mapping = {
        'TA (0 -78m)': 0, 'TA (78m - 234m)': 2, 'TA (234m - 546m)': 5,
        'TA (546m - 1014m)': 10, 'TA (1014m-1950m)': 19, 'TA (1950m - 3510m)': 35,
        'TA (3510m - 6630m)': 65, 'TA (6630m-14430m)': 135, 'TA (14430m - 30030m)': 285,
        'TA (30030m - 53430m)': 535, 'TA (53430m - 76830m)': 835, 'TA (>76830m)': 1000
    }
    
    ta_columns = [col for col in df_raw.columns if col in ta_mapping.keys()]
    
    df_melted = pd.melt(
        df_raw,
        id_vars=['Site_ID', 'Cell_ID'],
        value_vars=ta_columns,
        var_name='TA_Range',
        value_name='Samples'
    )
    
    df_melted['TA'] = df_melted['TA_Range'].map(ta_mapping)
    df_all = df_melted[df_melted['Samples'] > 0].copy()
    
    return df_all

def calculate_ta_kpis(df, percentile=90, planned_radius=3.0):
    """Calculate TA/Coverage KPIs"""
    if len(df) == 0:
        return None
    
    df = df.copy()
    df["Distance_m"] = df["TA"] * 78
    df["Distance_km"] = df["Distance_m"] / 1000
    
    total_samples = df["Samples"].sum()
    avg_ta = (df["TA"] * df["Samples"]).sum() / total_samples
    
    df_sorted = df.sort_values('TA')
    df_sorted['cumsum'] = df_sorted['Samples'].cumsum()
    df_sorted['cumsum_pct'] = df_sorted['cumsum'] / total_samples
    
    p50_ta = df_sorted[df_sorted['cumsum_pct'] >= 0.50]['TA'].iloc[0]
    p90_ta = df_sorted[df_sorted['cumsum_pct'] >= 0.90]['TA'].iloc[0]
    pxx_ta = df_sorted[df_sorted['cumsum_pct'] >= (percentile/100)]['TA'].iloc[0]
    
    p90_distance = p90_ta * 0.078
    pxx_distance = pxx_ta * 0.078
    
    overshoot_ratio = (df[df["TA"] >= 16]["Samples"].sum() / total_samples) * 100
    
    samples_within_radius = df[df["Distance_km"] <= planned_radius]["Samples"].sum()
    coverage_efficiency = (samples_within_radius / total_samples) * 100
    
    return {
        'avg_ta': avg_ta, 'p50_ta': p50_ta, 'p90_ta': p90_ta, 'pxx_ta': pxx_ta,
        'p90_distance': p90_distance, 'pxx_distance': pxx_distance,
        'overshoot_ratio': overshoot_ratio, 'coverage_efficiency': coverage_efficiency,
        'total_samples': total_samples
    }

def calculate_overall_health(row, thresholds):
    """Calculate overall cell health score (0-100)"""
    score = 100
    
    # Helper function to safely get numeric value
    def safe_get(key, default=None):
        val = row.get(key)
        if pd.isna(val):
            return default
        try:
            return float(val)
        except (TypeError, ValueError):
            return default
    
    # Coverage (30 points)
    avg_ta_dist = safe_get('Avg TA Distance(m)')
    if avg_ta_dist is not None:
        avg_ta_dist_km = avg_ta_dist / 1000
        if avg_ta_dist_km > 1.5:
            score -= 30
        elif avg_ta_dist_km > 1.0:
            score -= 15
    
    # Availability (20 points)
    availability = safe_get('LTE Network Availability (%)')
    if availability is not None:
        if availability < 99.0:
            score -= 20
        elif availability < 99.5:
            score -= 10
    
    # Accessibility (20 points)
    rrc_cssr = safe_get('RRC CSSR(%)')
    if rrc_cssr is not None:
        if rrc_cssr < 95.0:
            score -= 20
        elif rrc_cssr < 98.0:
            score -= 10
    
    # Retainability (15 points)
    erab_dcr = safe_get('ERAB DCR(%)')
    if erab_dcr is not None:
        if erab_dcr > 2.0:
            score -= 15
        elif erab_dcr > 1.0:
            score -= 7
    
    # Mobility (15 points)
    ho_sr = safe_get('HO SR(%)')
    if ho_sr is not None:
        if ho_sr < 95.0:
            score -= 15
        elif ho_sr < 98.0:
            score -= 7
    
    return max(0, score)

def get_verdict(health_score):
    """Get verdict based on health score"""
    if health_score >= 80:
        return "Excellent", "🟢"
    elif health_score >= 60:
        return "Good", "🟡"
    elif health_score >= 40:
        return "Monitor", "🟠"
    else:
        return "Critical", "🔴"

# -------------------------------------------------
# Main App Header
# -------------------------------------------------
st.title("📡 LTE Network Optimization")
st.caption("Complete Performance Analysis & Optimization Platform | All KPIs Integration")

# -------------------------------------------------
# Sidebar - Configuration
# -------------------------------------------------
st.sidebar.header("⚙️ Configuration")

# KPI Thresholds
with st.sidebar.expander("🎯 KPI Thresholds", expanded=False):
    st.subheader("Accessibility")
    rrc_cssr_target = st.number_input("RRC CSSR Target (%)", value=98.0, min_value=90.0, max_value=100.0)
    
    st.subheader("Retainability")
    erab_dcr_target = st.number_input("ERAB DCR Target (%)", value=1.0, min_value=0.0, max_value=5.0)
    
    st.subheader("Mobility")
    ho_sr_target = st.number_input("HO SR Target (%)", value=98.0, min_value=90.0, max_value=100.0)
    
    st.subheader("VoLTE Quality")
    volte_cssr_target = st.number_input("VoLTE CSSR Target (%)", value=98.0, min_value=90.0, max_value=100.0)
    volte_dcr_target = st.number_input("VoLTE DCR Target (%)", value=1.0, min_value=0.0, max_value=5.0)

thresholds = {
    'rrc_cssr': rrc_cssr_target,
    'erab_dcr': erab_dcr_target,
    'ho_sr': ho_sr_target,
    'volte_cssr': volte_cssr_target,
    'volte_dcr': volte_dcr_target
}

percentile = st.sidebar.selectbox("Coverage Percentile", [85, 90, 95], index=1)
planned_radius = st.sidebar.number_input("Planned Cell Radius (km)", min_value=0.5, max_value=50.0, value=3.0, step=0.5)

# -------------------------------------------------
# File Upload Section
# -------------------------------------------------
st.header("📂 Data Upload")

col1, col2 = st.columns(2)

with col1:
    st.subheader("📊 Before Optimization")
    uploaded_before = st.file_uploader(
        "Upload BEFORE data (Required)",
        type=["csv", "xlsx", "xlsm"],
        key="before"
    )

with col2:
    st.subheader("📈 After Optimization")
    uploaded_after = st.file_uploader(
        "Upload AFTER data (Optional)",
        type=["csv", "xlsx", "xlsm"],
        key="after"
    )

# Load data
df_before = None
df_after = None

if uploaded_before:
    with st.spinner("Loading BEFORE data..."):
        df_before = load_data(uploaded_before)
        if df_before is not None:
            st.success(f"✅ BEFORE: {df_before['Site_ID'].nunique()} sites, {df_before['Cell_ID'].nunique()} cells loaded")

if uploaded_after:
    with st.spinner("Loading AFTER data..."):
        df_after = load_data(uploaded_after)
        if df_after is not None:
            st.success(f"✅ AFTER: {df_after['Site_ID'].nunique()} sites, {df_after['Cell_ID'].nunique()} cells loaded")

if df_before is None:
    st.info("⬆️ Upload at least the BEFORE optimization file to start analysis.")
    st.stop()

st.divider()

# -------------------------------------------------
# Main Analysis Tabs
# -------------------------------------------------
tabs = st.tabs([
    "📊 Executive Dashboard",
    "📶 Coverage & TA Analysis", 
    "🎯 Accessibility & Quality",
    "📡 Traffic & Capacity",
    "🔊 VoLTE Performance",
    "📻 RF Performance",
    "⚡ Availability Analysis",
    "🔄 Inter-RAT Performance",
    "📋 Multi-Cell Comparison",
    "💾 Export & Reports"
])

# -------------------------------------------------
# TAB 1: Executive Dashboard
# -------------------------------------------------
with tabs[0]:
    st.header("Executive Network Dashboard")
    
    try:
        # Overall Network Metrics
        total_cells = len(df_before)
        
        # Calculate health scores for all cells
        df_before['Health_Score'] = df_before.apply(lambda row: calculate_overall_health(row, thresholds), axis=1)
        
        avg_health = df_before['Health_Score'].mean()
        excellent_cells = len(df_before[df_before['Health_Score'] >= 80])
        critical_cells = len(df_before[df_before['Health_Score'] < 40])
        
        # Top Row Metrics
        col1, col2, col3, col4, col5 = st.columns(5)
        
        with col1:
            st.metric("Total Cells", total_cells)
        with col2:
            st.metric("Network Health", f"{avg_health:.1f}/100")
        with col3:
            st.metric("Excellent Cells", excellent_cells, delta=f"{(excellent_cells/total_cells*100):.1f}%")
        with col4:
            st.metric("Critical Cells", critical_cells, delta=f"-{(critical_cells/total_cells*100):.1f}%", delta_color="inverse")
        with col5:
            avg_availability = df_before['LTE Network Availability (%)'].mean()
            st.metric("Avg Availability", f"{avg_availability:.2f}%")
        
        st.divider()
        
        # Key Performance Indicators
        st.subheader("Network-Wide KPI Performance")
        
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            st.markdown("**Accessibility**")
            avg_rrc_cssr = df_before['RRC CSSR(%)'].mean()
            st.metric("Avg RRC CSSR", f"{avg_rrc_cssr:.2f}%", 
                     delta=f"{avg_rrc_cssr - thresholds['rrc_cssr']:+.2f}%")
            
            below_target = len(df_before[df_before['RRC CSSR(%)'] < thresholds['rrc_cssr']])
            st.metric("Cells Below Target", below_target)
        
        with col2:
            st.markdown("**Retainability**")
            avg_erab_dcr = df_before['ERAB DCR(%)'].mean()
            st.metric("Avg ERAB DCR", f"{avg_erab_dcr:.2f}%",
                     delta=f"{avg_erab_dcr - thresholds['erab_dcr']:+.2f}%",
                     delta_color="inverse")
            
            above_target = len(df_before[df_before['ERAB DCR(%)'] > thresholds['erab_dcr']])
            st.metric("Cells Above Target", above_target)
        
        with col3:
            st.markdown("**Mobility**")
            avg_ho_sr = df_before['HO SR(%)'].mean()
            st.metric("Avg HO SR", f"{avg_ho_sr:.2f}%",
                     delta=f"{avg_ho_sr - thresholds['ho_sr']:+.2f}%")
            
            below_target = len(df_before[df_before['HO SR(%)'] < thresholds['ho_sr']])
            st.metric("Cells Below Target", below_target)
        
        with col4:
            st.markdown("**VoLTE Quality**")
            avg_volte_cssr = df_before['VoLTE CSSR(%)'].mean()
            st.metric("Avg VoLTE CSSR", f"{avg_volte_cssr:.2f}%",
                     delta=f"{avg_volte_cssr - thresholds['volte_cssr']:+.2f}%")
            
            avg_volte_dcr = df_before['VoLTE DCR(%)'].mean()
            st.metric("Avg VoLTE DCR", f"{avg_volte_dcr:.2f}%",
                     delta=f"{avg_volte_dcr - thresholds['volte_dcr']:+.2f}%",
                     delta_color="inverse")
        
        st.divider()
        
        # Visualizations
        col1, col2 = st.columns(2)
        
        with col1:
            st.subheader("Health Score Distribution")
            
            # Apply verdict to all cells
            df_before['Verdict'] = df_before['Health_Score'].apply(lambda x: get_verdict(x)[0])
            verdict_counts = df_before['Verdict'].value_counts()
            
            fig_pie = px.pie(
                values=verdict_counts.values,
                names=verdict_counts.index,
                color=verdict_counts.index,
                color_discrete_map={
                    'Excellent': '#2ca02c', 'Good': '#ffd700',
                    'Monitor': '#ff7f0e', 'Critical': '#d62728'
                }
            )
            st.plotly_chart(fig_pie, use_container_width=True)
        
        with col2:
            st.subheader("Traffic & Capacity Overview")
            
            fig_traffic = go.Figure()
            fig_traffic.add_trace(go.Bar(
                name='DL PRB Util %',
                x=df_before['Cell_ID'][:10],
                y=df_before['DL PRB Utilization(%)'][:10],
                marker_color='lightblue'
            ))
            fig_traffic.add_trace(go.Bar(
                name='UL PRB Util %',
                x=df_before['Cell_ID'][:10],
                y=df_before['UL PRB Utilization(%)'][:10],
                marker_color='lightcoral'
            ))
            fig_traffic.update_layout(barmode='group', xaxis_tickangle=-45)
            st.plotly_chart(fig_traffic, use_container_width=True)
        
        st.divider()
        
        # Comprehensive KPI Summary Table
        st.subheader("📊 Complete Network KPI Summary - All Cells")
        
        # Check if there are multiple rows per cell (multiple dates)
        cells_count = len(df_before)
        unique_cells = df_before.groupby(['Site_ID', 'Cell_ID']).size().shape[0]
        
        if cells_count > unique_cells:
            st.caption(f"📅 Note: Data contains {cells_count} records for {unique_cells} unique cells. Showing latest values per cell.")
            # Take the latest record for each cell (assuming Date column exists)
            if 'Date' in df_before.columns:
                df_summary_base = df_before.sort_values('Date').groupby(['Site_ID', 'Cell_ID']).last().reset_index()
            else:
                # If no date column, take the first occurrence
                df_summary_base = df_before.groupby(['Site_ID', 'Cell_ID']).first().reset_index()
        else:
            st.caption("Comprehensive view of all key performance indicators across the network")
            df_summary_base = df_before.copy()
        
        # Create comprehensive summary
        summary_data = []
        
        for idx, row in df_summary_base.iterrows():
            cell_summary = {
                # Identifiers
                'Site': row['Site_ID'],
                'Cell': row['Cell_ID'],
                'Health': row['Health_Score'],
                'Status': row['Verdict'],
                
                # Availability
                'Availability (%)': row['LTE Network Availability (%)'],
                'Downtime (min)': row['Cell Downtime(min)'],
                
                # Traffic & Users
                'Avg Users': row['Traffic User(Avg)'],
                'Max Users': row['Traffic User(Max)'],
                'VoLTE Users': row['VoLTE User'],
                
                # Throughput
                'DL Tput (Mbps)': row['DL Throughput(Mbit/s)'],
                'UL Tput (Mbps)': row['UL Throughput(Mbit/s)'],
                
                # Traffic Volume
                'DL Volume (GB)': row['DL Traffic Volume(GB)'],
                'UL Volume (GB)': row['UL Traffic Volume(GB)'],
                
                # PRB Utilization
                'DL PRB (%)': row['DL PRB Utilization(%)'],
                'UL PRB (%)': row['UL PRB Utilization(%)'],
                
                # RF Quality
                'UL Interf (dBm)': row['UL Interference(dBm)'],
                'PDSCH IBLER (%)': row['PDSCH IBLER(%)'],
                'PUSCH IBLER (%)': row['PUSCH IBLER(%)'],
                'Avg CQI': row['Avg CQI'],
                'MIMO Rank2 (%)': row['MIMO Rank2'],
                
                # Quality KPIs
                'RRC CSSR (%)': row['RRC CSSR(%)'],
                'ERAB CSSR (%)': row['ERAB CSSR(%)'],
                'ERAB DCR (%)': row['ERAB DCR(%)'],
                'HO SR (%)': row['HO SR(%)'],
                
                # VoLTE
                'VoLTE CSSR (%)': row['VoLTE CSSR(%)'],
                'VoLTE DCR (%)': row['VoLTE DCR(%)'],
                'VoLTE Traffic (Erl)': row['VoLTE Traffic (Erl)'],
                
                # Coverage
                'Avg TA Dist (m)': row['Avg TA Distance(m)']
            }
            summary_data.append(cell_summary)
        
        summary_df = pd.DataFrame(summary_data)
        
        # Display options
        col1, col2, col3 = st.columns([2, 2, 1])
        
        with col1:
            sort_column = st.selectbox(
                "Sort by",
                options=['Health', 'Availability (%)', 'DL PRB (%)', 'UL PRB (%)', 
                        'RRC CSSR (%)', 'ERAB DCR (%)', 'Avg Users', 'DL Tput (Mbps)'],
                key="summary_sort"
            )
        
        with col2:
            filter_status = st.multiselect(
                "Filter by Status",
                options=['All'] + sorted(summary_df['Status'].unique().tolist()),
                default=['All'],
                key="summary_filter"
            )
        
        with col3:
            show_rows = st.selectbox("Show rows", [10, 20, 50, 100, "All"], index=1, key="summary_rows")
        
        # Apply filters
        filtered_summary = summary_df.copy()
        if 'All' not in filter_status:
            filtered_summary = filtered_summary[filtered_summary['Status'].isin(filter_status)]
        
        # Sort
        filtered_summary = filtered_summary.sort_values(by=sort_column, ascending=False)
        
        # Limit rows
        if show_rows != "All":
            filtered_summary = filtered_summary.head(show_rows)
        
        # Round numeric columns
        numeric_cols = filtered_summary.select_dtypes(include=[np.number]).columns
        filtered_summary[numeric_cols] = filtered_summary[numeric_cols].round(2)
        
        # Color coding function
        def color_status(val):
            if val == 'Excellent':
                return 'background-color: #d4edda'
            elif val == 'Good':
                return 'background-color: #fff3cd'
            elif val == 'Monitor':
                return 'background-color: #f8d7da'
            elif val == 'Critical':
                return 'background-color: #f5c6cb'
            return ''
        
        # Apply styling
        styled_summary = filtered_summary.style.applymap(color_status, subset=['Status'])
        
        # Display table
        st.dataframe(styled_summary, use_container_width=True, hide_index=True, height=400)
        
        # Summary Statistics Row
        st.markdown("**Network-Wide Statistics:**")
        
        stat_col1, stat_col2 = st.columns(2)
        
        with stat_col1:
            st.markdown("**Average Values Across Network:**")
            avg_stats = pd.DataFrame({
                'KPI': [
                    'Availability', 'DL PRB Util', 'UL PRB Util', 'Avg Users',
                    'RRC CSSR', 'ERAB DCR', 'HO SR', 'Avg CQI',
                    'DL Throughput', 'UL Throughput', 'VoLTE CSSR'
                ],
                'Average': [
                    f"{summary_df['Availability (%)'].mean():.2f}%",
                    f"{summary_df['DL PRB (%)'].mean():.1f}%",
                    f"{summary_df['UL PRB (%)'].mean():.1f}%",
                    f"{summary_df['Avg Users'].mean():.1f}",
                    f"{summary_df['RRC CSSR (%)'].mean():.2f}%",
                    f"{summary_df['ERAB DCR (%)'].mean():.2f}%",
                    f"{summary_df['HO SR (%)'].mean():.2f}%",
                    f"{summary_df['Avg CQI'].mean():.2f}",
                    f"{summary_df['DL Tput (Mbps)'].mean():.1f} Mbps",
                    f"{summary_df['UL Tput (Mbps)'].mean():.1f} Mbps",
                    f"{summary_df['VoLTE CSSR (%)'].mean():.2f}%"
                ]
            })
            st.dataframe(avg_stats, use_container_width=True, hide_index=True)
        
        with stat_col2:
            st.markdown("**Max Values Across Network:**")
            max_stats = pd.DataFrame({
                'KPI': [
                    'Max Users', 'Max DL PRB', 'Max UL PRB', 'Max Downtime',
                    'Worst RRC CSSR', 'Worst ERAB DCR', 'Worst HO SR', 'Worst CQI',
                    'Max DL Throughput', 'Max UL Throughput', 'Best Cell Health'
                ],
                'Value': [
                    f"{summary_df['Max Users'].max():.0f}",
                    f"{summary_df['DL PRB (%)'].max():.1f}%",
                    f"{summary_df['UL PRB (%)'].max():.1f}%",
                    f"{summary_df['Downtime (min)'].max():.1f} min",
                    f"{summary_df['RRC CSSR (%)'].min():.2f}%",
                    f"{summary_df['ERAB DCR (%)'].max():.2f}%",
                    f"{summary_df['HO SR (%)'].min():.2f}%",
                    f"{summary_df['Avg CQI'].min():.2f}",
                    f"{summary_df['DL Tput (Mbps)'].max():.1f} Mbps",
                    f"{summary_df['UL Tput (Mbps)'].max():.1f} Mbps",
                    f"{summary_df['Health'].max():.0f}/100"
                ],
                'Cell': [
                    summary_df.loc[summary_df['Max Users'].idxmax(), 'Cell'],
                    summary_df.loc[summary_df['DL PRB (%)'].idxmax(), 'Cell'],
                    summary_df.loc[summary_df['UL PRB (%)'].idxmax(), 'Cell'],
                    summary_df.loc[summary_df['Downtime (min)'].idxmax(), 'Cell'],
                    summary_df.loc[summary_df['RRC CSSR (%)'].idxmin(), 'Cell'],
                    summary_df.loc[summary_df['ERAB DCR (%)'].idxmax(), 'Cell'],
                    summary_df.loc[summary_df['HO SR (%)'].idxmin(), 'Cell'],
                    summary_df.loc[summary_df['Avg CQI'].idxmin(), 'Cell'],
                    summary_df.loc[summary_df['DL Tput (Mbps)'].idxmax(), 'Cell'],
                    summary_df.loc[summary_df['UL Tput (Mbps)'].idxmax(), 'Cell'],
                    summary_df.loc[summary_df['Health'].idxmax(), 'Cell']
                ]
            })
            st.dataframe(max_stats, use_container_width=True, hide_index=True)
        
        # Download option
        csv = filtered_summary.to_csv(index=False)
        st.download_button(
            label="📥 Download Summary as CSV",
            data=csv,
            file_name=f"network_summary_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
            mime="text/csv"
        )
        
        st.divider()
        
        # KPI Heatmap Visualization
        st.subheader("🔥 KPI Heatmap - Quick Problem Identification")
        st.caption("Color intensity shows performance: Green=Good, Yellow=Warning, Red=Critical")
        
        # Select KPIs for heatmap
        heatmap_kpis = st.multiselect(
            "Select KPIs to display in heatmap",
            options=['Health', 'Availability (%)', 'RRC CSSR (%)', 'ERAB DCR (%)', 'HO SR (%)',
                    'DL PRB (%)', 'UL PRB (%)', 'Avg CQI', 'PDSCH IBLER (%)', 'PUSCH IBLER (%)'],
            default=['Health', 'RRC CSSR (%)', 'ERAB DCR (%)', 'DL PRB (%)', 'Avg CQI'],
            key="heatmap_kpis"
        )
        
        if heatmap_kpis:
            try:
                # Create unique cell identifier (Site_Cell)
                summary_df_heatmap = summary_df.copy()
                summary_df_heatmap['Cell_ID_Full'] = summary_df_heatmap['Site'] + '_' + summary_df_heatmap['Cell']
                
                # Aggregate duplicate cells by taking the mean (in case of multiple time periods)
                # Group by Cell_ID_Full and take mean of numeric columns
                heatmap_data = summary_df_heatmap[['Cell_ID_Full'] + heatmap_kpis].groupby('Cell_ID_Full').mean()
                
                # Remove any rows with NaN values in selected KPIs
                heatmap_data = heatmap_data.dropna()
                
                if len(heatmap_data) == 0:
                    st.warning("⚠️ No valid data available for selected KPIs. Please check your data.")
                else:
                    # Normalize data for better color scale (invert for DCR and IBLER where lower is better)
                    heatmap_normalized = heatmap_data.copy()
                    
                    # For DCR and IBLER, invert the scale (lower is better)
                    inverse_kpis = ['ERAB DCR (%)', 'PDSCH IBLER (%)', 'PUSCH IBLER (%)']
                    for kpi in inverse_kpis:
                        if kpi in heatmap_normalized.columns:
                            max_val = heatmap_normalized[kpi].max()
                            if max_val > 0:
                                heatmap_normalized[kpi] = max_val - heatmap_normalized[kpi]
                    
                    # Limit display to max 50 cells for readability
                    if len(heatmap_normalized) > 50:
                        st.info(f"ℹ️ Displaying top 50 cells (sorted by health). Your data has {len(heatmap_normalized)} cells total.")
                        heatmap_normalized = heatmap_normalized.nlargest(50, 'Health' if 'Health' in heatmap_normalized.columns else heatmap_normalized.columns[0])
                    
                    # Create heatmap
                    fig_heatmap = px.imshow(
                        heatmap_normalized.T,
                        labels=dict(x="Cell ID (Site_Cell)", y="KPI", color="Performance"),
                        x=heatmap_normalized.index,
                        y=heatmap_normalized.columns,
                        color_continuous_scale='RdYlGn',
                        aspect='auto'
                    )
                    
                    fig_heatmap.update_layout(
                        height=400,
                        xaxis={'side': 'bottom'},
                        xaxis_tickangle=-45
                    )
                    
                    st.plotly_chart(fig_heatmap, use_container_width=True)
                    
                    st.info("💡 Tip: Red cells indicate problem areas, Green cells are performing well. Click and drag to zoom into specific cells.")
            
            except Exception as e:
                st.error(f"Error generating heatmap: {str(e)}")
                st.info("Please check your data format and ensure all numeric columns contain valid numbers.")
        
        st.divider()
        
        # Issue Categorization & Priority
        st.subheader("🚨 Network Issues Summary - By Category")
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.markdown("**Quality Issues**")
            quality_issues = []
            
            # RRC CSSR issues
            rrc_issues = len(df_before[df_before['RRC CSSR(%)'] < thresholds['rrc_cssr']])
            if rrc_issues > 0:
                quality_issues.append(f"🔴 RRC CSSR: {rrc_issues} cells")
            
            # ERAB DCR issues
            dcr_issues = len(df_before[df_before['ERAB DCR(%)'] > thresholds['erab_dcr']])
            if dcr_issues > 0:
                quality_issues.append(f"🔴 ERAB DCR: {dcr_issues} cells")
            
            # HO SR issues
            ho_issues = len(df_before[df_before['HO SR(%)'] < thresholds['ho_sr']])
            if ho_issues > 0:
                quality_issues.append(f"🔴 HO SR: {ho_issues} cells")
            
            # VoLTE issues
            volte_cssr_issues = len(df_before[df_before['VoLTE CSSR(%)'] < thresholds['volte_cssr']])
            if volte_cssr_issues > 0:
                quality_issues.append(f"🔴 VoLTE CSSR: {volte_cssr_issues} cells")
            
            if quality_issues:
                for issue in quality_issues:
                    st.markdown(issue)
            else:
                st.success("✅ No quality issues")
        
        with col2:
            st.markdown("**Capacity Issues**")
            capacity_issues = []
            
            # DL PRB overload
            dl_overload = len(df_before[df_before['DL PRB Utilization(%)'] > 70])
            if dl_overload > 0:
                capacity_issues.append(f"🟠 DL Overload: {dl_overload} cells")
            
            # UL PRB overload
            ul_overload = len(df_before[df_before['UL PRB Utilization(%)'] > 50])
            if ul_overload > 0:
                capacity_issues.append(f"🟠 UL Overload: {ul_overload} cells")
            
            # High traffic
            high_traffic = len(df_before[df_before['Traffic User(Avg)'] > 20])
            if high_traffic > 0:
                capacity_issues.append(f"🟡 High Traffic: {high_traffic} cells")
            
            if capacity_issues:
                for issue in capacity_issues:
                    st.markdown(issue)
            else:
                st.success("✅ No capacity issues")
        
        with col3:
            st.markdown("**RF Issues**")
            rf_issues = []
            
            # High interference
            high_interference = len(df_before[df_before['UL Interference(dBm)'] > -110])
            if high_interference > 0:
                rf_issues.append(f"🔴 High Interference: {high_interference} cells")
            
            # High IBLER
            high_pdsch_ibler = len(df_before[df_before['PDSCH IBLER(%)'] > 10])
            if high_pdsch_ibler > 0:
                rf_issues.append(f"🔴 High PDSCH IBLER: {high_pdsch_ibler} cells")
            
            high_pusch_ibler = len(df_before[df_before['PUSCH IBLER(%)'] > 10])
            if high_pusch_ibler > 0:
                rf_issues.append(f"🔴 High PUSCH IBLER: {high_pusch_ibler} cells")
            
            # Low CQI
            low_cqi = len(df_before[df_before['Avg CQI'] < 7])
            if low_cqi > 0:
                rf_issues.append(f"🟠 Low CQI: {low_cqi} cells")
            
            if rf_issues:
                for issue in rf_issues:
                    st.markdown(issue)
            else:
                st.success("✅ No RF issues")
        
        # Priority Action Items
        st.markdown("**🎯 Top Priority Actions:**")
        priority_actions = []
        
        # Critical health cells
        if critical_cells > 0:
            priority_actions.append(f"1. **URGENT:** {critical_cells} cells with critical health (<40) - Immediate investigation required")
        
        # Quality issues
        if rrc_issues > 0:
            priority_actions.append(f"2. **HIGH:** Address RRC CSSR issues in {rrc_issues} cells - Impacts accessibility")
        
        # Capacity issues
        if dl_overload > 0:
            priority_actions.append(f"3. **MEDIUM:** {dl_overload} cells with DL overload (>70%) - Plan capacity expansion")
        
        # RF issues
        if high_interference > 0:
            priority_actions.append(f"4. **MEDIUM:** {high_interference} cells with high interference - RF optimization needed")
        
        if not priority_actions:
            st.success("✅ **Excellent!** Network is performing optimally with no critical issues.")
        else:
            for action in priority_actions[:5]:  # Show top 5 priorities
                st.markdown(action)
        
        st.divider()
        
        # Top/Bottom Performers
        col1, col2 = st.columns(2)
        
        with col1:
            st.subheader("🏆 Top 10 Best Cells")
            best_cells = df_before.nlargest(10, 'Health_Score')[
                ['Cell_ID', 'Health_Score', 'Verdict', 'RRC CSSR(%)', 'ERAB DCR(%)', 'HO SR(%)']
            ].round(2)
            best_cells.columns = ['Cell', 'Health', 'Status', 'RRC CSSR', 'ERAB DCR', 'HO SR']
            st.dataframe(best_cells, use_container_width=True, hide_index=True)
        
        with col2:
            st.subheader("⚠️ Top 10 Worst Cells")
            worst_cells = df_before.nsmallest(10, 'Health_Score')[
                ['Cell_ID', 'Health_Score', 'Verdict', 'RRC CSSR(%)', 'ERAB DCR(%)', 'HO SR(%)']
            ].round(2)
            worst_cells.columns = ['Cell', 'Health', 'Status', 'RRC CSSR', 'ERAB DCR', 'HO SR']
            st.dataframe(worst_cells, use_container_width=True, hide_index=True)
    
    except Exception as e:
        st.error(f"Error in Executive Dashboard: {str(e)}")
        st.info("Please check your data format and ensure all numeric columns contain valid numbers.")

# -------------------------------------------------
# TAB 2: Coverage & TA Analysis
# -------------------------------------------------
with tabs[1]:
    st.header("📶 Coverage & Timing Advance Analysis")
    
    try:
        # Process TA data
        df_ta = process_ta_data(df_before)
        
        # Cell Selection
        col1, col2 = st.columns(2)
        
        with col1:
            selected_site = st.selectbox("Select Site", sorted(df_before["Site_ID"].unique()), key="ta_site")
        
        with col2:
            available_cells = sorted(df_before[df_before["Site_ID"] == selected_site]["Cell_ID"].unique())
            selected_cell = st.selectbox("Select Cell", available_cells, key="ta_cell")
        
        # Get cell data
        cell_data_full = df_before[
            (df_before["Site_ID"] == selected_site) &
            (df_before["Cell_ID"] == selected_cell)
        ].iloc[0]
        
        cell_ta = df_ta[
            (df_ta["Site_ID"] == selected_site) &
            (df_ta["Cell_ID"] == selected_cell)
        ]
        
        if len(cell_ta) > 0:
            # Calculate TA KPIs
            ta_kpis = calculate_ta_kpis(cell_ta, percentile, planned_radius)
            
            st.subheader(f"{selected_site} - {selected_cell}")
            
            # TA Metrics
            col1, col2, col3, col4, col5 = st.columns(5)
            col1.metric("Avg TA", f"{ta_kpis['avg_ta']:.1f}")
            col2.metric("P50 TA", f"{ta_kpis['p50_ta']:.0f}")
            col3.metric("P90 TA", f"{ta_kpis['p90_ta']:.0f}")
            col4.metric("P90 Distance", f"{ta_kpis['p90_distance']:.2f} km")
            col5.metric("Overshoot", f"{ta_kpis['overshoot_ratio']:.1f}%")
            
            # From raw data
            col1, col2, col3 = st.columns(3)
            col1.metric("Avg TA Distance", f"{cell_data_full['Avg TA Distance(m)']:.0f} m")
            col2.metric("Coverage Efficiency", f"{ta_kpis['coverage_efficiency']:.1f}%")
            col3.metric("Total Samples", f"{int(ta_kpis['total_samples']):,}")
            
            st.divider()
            
            # TA Distribution
            col1, col2 = st.columns(2)
            
            with col1:
                st.subheader("TA Distribution")
                
                bins = [0, 5, 10, 15, 20, 999]
                labels = ["0-5", "6-10", "11-15", "16-20", ">20"]
                cell_ta_copy = cell_ta.copy()
                cell_ta_copy["TA_Bucket"] = pd.cut(cell_ta_copy["TA"], bins=bins, labels=labels)
                
                bucket_dist = cell_ta_copy.groupby("TA_Bucket")["Samples"].sum().reset_index()
                
                fig_bucket = px.bar(bucket_dist, x="TA_Bucket", y="Samples",
                                   labels={"TA_Bucket": "TA Range", "Samples": "UE Samples"})
                
                if ta_kpis['overshoot_ratio'] > 10:
                    fig_bucket.update_traces(marker_color=[
                        "red" if ("16" in str(b) or ">" in str(b)) else "#1f77b4"
                        for b in bucket_dist["TA_Bucket"]
                    ])
                
                st.plotly_chart(fig_bucket, use_container_width=True)
            
            with col2:
                st.subheader("Coverage Efficiency")
                
                within = ta_kpis['coverage_efficiency']
                beyond = 100 - within
                
                fig_donut = go.Figure(data=[go.Pie(
                    labels=['Within Planned', 'Beyond Planned'],
                    values=[within, beyond],
                    hole=.4,
                    marker_colors=['#2ca02c', '#d62728']
                )])
                fig_donut.update_layout(
                    annotations=[dict(text=f'{within:.1f}%', x=0.5, y=0.5, font_size=20, showarrow=False)]
                )
                st.plotly_chart(fig_donut, use_container_width=True)
        else:
            st.warning("No TA data available for selected cell")
    
    except Exception as e:
        st.error(f"Error in Coverage & TA Analysis: {str(e)}")
        st.info("Please check your data format.")

# -------------------------------------------------
# TAB 3: Accessibility & Quality
# -------------------------------------------------
with tabs[2]:
    st.header("🎯 Accessibility & Quality KPIs")
    
    # Cell Selection
    col1, col2 = st.columns(2)
    
    with col1:
        site_acc = st.selectbox("Select Site", sorted(df_before["Site_ID"].unique()), key="acc_site")
    
    with col2:
        cells_acc = sorted(df_before[df_before["Site_ID"] == site_acc]["Cell_ID"].unique())
        cell_acc = st.selectbox("Select Cell", cells_acc, key="acc_cell")
    
    cell_data = df_before[
        (df_before["Site_ID"] == site_acc) &
        (df_before["Cell_ID"] == cell_acc)
    ].iloc[0]
    
    st.subheader(f"{site_acc} - {cell_acc}")
    
    # KPI Display
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.markdown("**Accessibility (RRC)**")
        rrc_cssr = cell_data['RRC CSSR(%)']
        delta_rrc = rrc_cssr - thresholds['rrc_cssr']
        st.metric("RRC CSSR", f"{rrc_cssr:.2f}%", delta=f"{delta_rrc:+.2f}%",
                 delta_color="normal" if delta_rrc >= 0 else "inverse")
    
    with col2:
        st.markdown("**Session Setup (ERAB)**")
        erab_cssr = cell_data['ERAB CSSR(%)']
        st.metric("ERAB CSSR", f"{erab_cssr:.2f}%")
        
        erab_dcr = cell_data['ERAB DCR(%)']
        delta_erab = erab_dcr - thresholds['erab_dcr']
        st.metric("ERAB DCR", f"{erab_dcr:.2f}%", delta=f"{delta_erab:+.2f}%",
                 delta_color="inverse" if delta_erab > 0 else "normal")
    
    with col3:
        st.markdown("**Mobility (Handover)**")
        ho_sr = cell_data['HO SR(%)']
        delta_ho = ho_sr - thresholds['ho_sr']
        st.metric("HO SR", f"{ho_sr:.2f}%", delta=f"{delta_ho:+.2f}%",
                 delta_color="normal" if delta_ho >= 0 else "inverse")
    
    with col4:
        st.markdown("**CSFB Performance**")
        csfb_sr = cell_data['CSFB SR(%)']
        st.metric("CSFB SR", f"{csfb_sr:.2f}%")
        
        srvcc_sr = cell_data['SRVCC SR(%)']
        st.metric("SRVCC SR", f"{srvcc_sr:.2f}%")
    
    st.divider()
    
    # Network-wide comparison
    st.subheader("Network-Wide Quality Performance")
    
    col1, col2 = st.columns(2)
    
    with col1:
        fig_acc = go.Figure()
        fig_acc.add_trace(go.Box(y=df_before['RRC CSSR(%)'], name='RRC CSSR'))
        fig_acc.add_trace(go.Box(y=df_before['ERAB CSSR(%)'], name='ERAB CSSR'))
        fig_acc.add_trace(go.Box(y=df_before['HO SR(%)'], name='HO SR'))
        fig_acc.update_layout(title="Accessibility & Mobility Distribution", yaxis_title="Success Rate (%)")
        st.plotly_chart(fig_acc, use_container_width=True)
    
    with col2:
        fig_dcr = go.Figure()
        fig_dcr.add_trace(go.Box(y=df_before['ERAB DCR(%)'], name='ERAB DCR'))
        fig_dcr.update_layout(title="Drop Call Rate Distribution", yaxis_title="Drop Rate (%)")
        st.plotly_chart(fig_dcr, use_container_width=True)
    
    # Cells not meeting targets
    st.subheader("⚠️ Cells Not Meeting Targets")
    
    problem_cells = df_before[
        (df_before['RRC CSSR(%)'] < thresholds['rrc_cssr']) |
        (df_before['ERAB DCR(%)'] > thresholds['erab_dcr']) |
        (df_before['HO SR(%)'] < thresholds['ho_sr'])
    ][['Cell_ID', 'RRC CSSR(%)', 'ERAB CSSR(%)', 'ERAB DCR(%)', 'HO SR(%)']].round(2)
    
    if len(problem_cells) > 0:
        st.dataframe(problem_cells, use_container_width=True)
    else:
        st.success("✅ All cells meeting quality targets!")

# -------------------------------------------------
# TAB 4: Traffic & Capacity
# -------------------------------------------------
with tabs[3]:
    st.header("📡 Traffic & Capacity Analysis")
    
    # Cell Selection
    col1, col2 = st.columns(2)
    
    with col1:
        site_cap = st.selectbox("Select Site", sorted(df_before["Site_ID"].unique()), key="cap_site")
    
    with col2:
        cells_cap = sorted(df_before[df_before["Site_ID"] == site_cap]["Cell_ID"].unique())
        cell_cap = st.selectbox("Select Cell", cells_cap, key="cap_cell")
    
    cell_data = df_before[
        (df_before["Site_ID"] == site_cap) &
        (df_before["Cell_ID"] == cell_cap)
    ].iloc[0]
    
    st.subheader(f"{site_cap} - {cell_cap}")
    
    # Traffic Metrics
    col1, col2, col3, col4, col5 = st.columns(5)
    
    col1.metric("Avg Users", f"{cell_data['Traffic User(Avg)']:.1f}")
    col2.metric("Max Users", f"{cell_data['Traffic User(Max)']:.0f}")
    col3.metric("VoLTE Users", f"{cell_data['VoLTE User']:.2f}")
    col4.metric("DL PRB Util", f"{cell_data['DL PRB Utilization(%)']:.1f}%")
    col5.metric("UL PRB Util", f"{cell_data['UL PRB Utilization(%)']:.1f}%")
    
    st.divider()
    
    col1, col2, col3, col4 = st.columns(4)
    
    col1.metric("DL Throughput", f"{cell_data['DL Throughput(Mbit/s)']:.1f} Mbps")
    col2.metric("UL Throughput", f"{cell_data['UL Throughput(Mbit/s)']:.1f} Mbps")
    col3.metric("DL Volume", f"{cell_data['DL Traffic Volume(GB)']:.1f} GB")
    col4.metric("UL Volume", f"{cell_data['UL Traffic Volume(GB)']:.1f} GB")
    
    st.divider()
    
    # Capacity Analysis
    st.subheader("Capacity Utilization Analysis")
    
    col1, col2 = st.columns(2)
    
    with col1:
        # PRB Utilization Gauge
        dl_prb = cell_data['DL PRB Utilization(%)']
        
        fig_gauge = go.Figure(go.Indicator(
            mode="gauge+number+delta",
            value=dl_prb,
            domain={'x': [0, 1], 'y': [0, 1]},
            title={'text': "DL PRB Utilization (%)"},
            delta={'reference': 70},
            gauge={
                'axis': {'range': [None, 100]},
                'bar': {'color': "darkblue"},
                'steps': [
                    {'range': [0, 50], 'color': "lightgreen"},
                    {'range': [50, 70], 'color': "yellow"},
                    {'range': [70, 100], 'color': "red"}
                ],
                'threshold': {
                    'line': {'color': "red", 'width': 4},
                    'thickness': 0.75,
                    'value': 70
                }
            }
        ))
        st.plotly_chart(fig_gauge, use_container_width=True)
    
    with col2:
        # UL PRB Utilization Gauge
        ul_prb = cell_data['UL PRB Utilization(%)']
        
        fig_gauge2 = go.Figure(go.Indicator(
            mode="gauge+number+delta",
            value=ul_prb,
            domain={'x': [0, 1], 'y': [0, 1]},
            title={'text': "UL PRB Utilization (%)"},
            delta={'reference': 50},
            gauge={
                'axis': {'range': [None, 100]},
                'bar': {'color': "darkred"},
                'steps': [
                    {'range': [0, 30], 'color': "lightgreen"},
                    {'range': [30, 50], 'color': "yellow"},
                    {'range': [50, 100], 'color': "red"}
                ],
                'threshold': {
                    'line': {'color': "red", 'width': 4},
                    'thickness': 0.75,
                    'value': 50
                }
            }
        ))
        st.plotly_chart(fig_gauge2, use_container_width=True)
    
    # Capacity Recommendations
    st.subheader("📋 Capacity Recommendations")
    
    recommendations = []
    
    if dl_prb > 70:
        recommendations.append(f"🔴 **High DL Load:** DL PRB at {dl_prb:.1f}% - Consider carrier addition or offload")
    elif dl_prb > 50:
        recommendations.append(f"🟡 **Moderate DL Load:** DL PRB at {dl_prb:.1f}% - Monitor for growth")
    else:
        recommendations.append(f"🟢 **Normal DL Load:** DL PRB at {dl_prb:.1f}% - Adequate capacity")
    
    if ul_prb > 50:
        recommendations.append(f"🔴 **High UL Load:** UL PRB at {ul_prb:.1f}% - Review UL traffic patterns")
    elif ul_prb > 30:
        recommendations.append(f"🟡 **Moderate UL Load:** UL PRB at {ul_prb:.1f}% - Monitor for growth")
    else:
        recommendations.append(f"🟢 **Normal UL Load:** UL PRB at {ul_prb:.1f}% - Adequate capacity")
    
    for rec in recommendations:
        st.markdown(rec)
    
    st.divider()
    
    # Network-wide Traffic Patterns
    st.subheader("Network-Wide Traffic Patterns")
    
    fig_scatter = px.scatter(
        df_before,
        x='DL PRB Utilization(%)',
        y='UL PRB Utilization(%)',
        size='Traffic User(Avg)',
        color='Health_Score',
        hover_data=['Cell_ID'],
        labels={
            'DL PRB Utilization(%)': 'DL PRB Utilization (%)',
            'UL PRB Utilization(%)': 'UL PRB Utilization (%)',
            'Traffic User(Avg)': 'Avg Users'
        },
        color_continuous_scale='RdYlGn'
    )
    fig_scatter.add_hline(y=50, line_dash="dash", line_color="red", annotation_text="UL Threshold")
    fig_scatter.add_vline(x=70, line_dash="dash", line_color="red", annotation_text="DL Threshold")
    
    st.plotly_chart(fig_scatter, use_container_width=True)

# -------------------------------------------------
# TAB 5: VoLTE Performance
# -------------------------------------------------
with tabs[4]:
    st.header("🔊 VoLTE Performance Analysis")
    
    # Cell Selection
    col1, col2 = st.columns(2)
    
    with col1:
        site_volte = st.selectbox("Select Site", sorted(df_before["Site_ID"].unique()), key="volte_site")
    
    with col2:
        cells_volte = sorted(df_before[df_before["Site_ID"] == site_volte]["Cell_ID"].unique())
        cell_volte = st.selectbox("Select Cell", cells_volte, key="volte_cell")
    
    cell_data = df_before[
        (df_before["Site_ID"] == site_volte) &
        (df_before["Cell_ID"] == cell_volte)
    ].iloc[0]
    
    st.subheader(f"{site_volte} - {cell_volte}")
    
    # VoLTE KPIs
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        volte_cssr = cell_data['VoLTE CSSR(%)']
        delta_cssr = volte_cssr - thresholds['volte_cssr']
        st.metric("VoLTE CSSR", f"{volte_cssr:.2f}%", delta=f"{delta_cssr:+.2f}%",
                 delta_color="normal" if delta_cssr >= 0 else "inverse")
    
    with col2:
        volte_dcr = cell_data['VoLTE DCR(%)']
        delta_dcr = volte_dcr - thresholds['volte_dcr']
        st.metric("VoLTE DCR", f"{volte_dcr:.2f}%", delta=f"{delta_dcr:+.2f}%",
                 delta_color="inverse" if delta_dcr > 0 else "normal")
    
    with col3:
        volte_traffic = cell_data['VoLTE Traffic (Erl)']
        st.metric("VoLTE Traffic", f"{volte_traffic:.2f} Erl")
    
    with col4:
        volte_users = cell_data['VoLTE User']
        st.metric("VoLTE Users", f"{volte_users:.2f}")
    
    st.divider()
    
    # VoLTE Drop Cause Analysis
    st.subheader("VoLTE Drop Root Cause Analysis")
    
    drop_causes = {
        'Radio': cell_data['VoLTE Drop due Radio'],
        'Congestion': cell_data['VoLTE Drop due Congestion'],
        'TNL': cell_data['VoLTE Drop due TNL'],
        'MME': cell_data['VoLTE Drop due MME'],
        'EUtranGen': cell_data['VoLTE Drop due EUtranGen']
    }
    
    # Remove zero values
    drop_causes = {k: v for k, v in drop_causes.items() if v > 0}
    
    if drop_causes:
        col1, col2 = st.columns(2)
        
        with col1:
            fig_drops = px.pie(
                values=list(drop_causes.values()),
                names=list(drop_causes.keys()),
                title="VoLTE Drop Causes Distribution"
            )
            st.plotly_chart(fig_drops, use_container_width=True)
        
        with col2:
            st.markdown("**Drop Cause Analysis:**")
            total_drops = sum(drop_causes.values())
            
            for cause, count in sorted(drop_causes.items(), key=lambda x: x[1], reverse=True):
                pct = (count / total_drops) * 100
                st.write(f"**{cause}:** {count:.0f} drops ({pct:.1f}%)")
                
                if cause == 'Radio' and pct > 50:
                    st.warning("⚠️ High radio drops - Check RF coverage and interference")
                elif cause == 'Congestion' and pct > 30:
                    st.warning("⚠️ High congestion drops - Check capacity")
                elif cause == 'TNL' and pct > 20:
                    st.warning("⚠️ High TNL drops - Check transport network")
    else:
        st.success("✅ No VoLTE drops recorded")
    
    st.divider()
    
    # Network-wide VoLTE Performance
    st.subheader("Network-Wide VoLTE Performance")
    
    col1, col2 = st.columns(2)
    
    with col1:
        fig_volte_cssr = go.Figure()
        fig_volte_cssr.add_trace(go.Histogram(
            x=df_before['VoLTE CSSR(%)'],
            nbinsx=20,
            name='VoLTE CSSR'
        ))
        fig_volte_cssr.add_vline(x=thresholds['volte_cssr'], line_dash="dash", 
                                line_color="red", annotation_text="Target")
        fig_volte_cssr.update_layout(title="VoLTE CSSR Distribution", 
                                     xaxis_title="VoLTE CSSR (%)")
        st.plotly_chart(fig_volte_cssr, use_container_width=True)
    
    with col2:
        fig_volte_dcr = go.Figure()
        fig_volte_dcr.add_trace(go.Histogram(
            x=df_before['VoLTE DCR(%)'],
            nbinsx=20,
            name='VoLTE DCR'
        ))
        fig_volte_dcr.add_vline(x=thresholds['volte_dcr'], line_dash="dash", 
                               line_color="red", annotation_text="Target")
        fig_volte_dcr.update_layout(title="VoLTE DCR Distribution", 
                                    xaxis_title="VoLTE DCR (%)")
        st.plotly_chart(fig_volte_dcr, use_container_width=True)

# -------------------------------------------------
# TAB 6: RF Performance
# -------------------------------------------------
with tabs[5]:
    st.header("📻 RF Performance Analysis")
    
    # Cell Selection
    col1, col2 = st.columns(2)
    
    with col1:
        site_rf = st.selectbox("Select Site", sorted(df_before["Site_ID"].unique()), key="rf_site")
    
    with col2:
        cells_rf = sorted(df_before[df_before["Site_ID"] == site_rf]["Cell_ID"].unique())
        cell_rf = st.selectbox("Select Cell", cells_rf, key="rf_cell")
    
    cell_data = df_before[
        (df_before["Site_ID"] == site_rf) &
        (df_before["Cell_ID"] == cell_rf)
    ].iloc[0]
    
    st.subheader(f"{site_rf} - {cell_rf}")
    
    # RF Metrics
    col1, col2, col3, col4, col5 = st.columns(5)
    
    col1.metric("UL Interference", f"{cell_data['UL Interference(dBm)']:.2f} dBm")
    col2.metric("PDSCH IBLER", f"{cell_data['PDSCH IBLER(%)']:.2f}%")
    col3.metric("PUSCH IBLER", f"{cell_data['PUSCH IBLER(%)']:.2f}%")
    col4.metric("Avg CQI", f"{cell_data['Avg CQI']:.2f}")
    col5.metric("MIMO Rank2", f"{cell_data['MIMO Rank2']:.2f}%")
    
    st.divider()
    
    # RF Analysis
    st.subheader("RF Quality Assessment")
    
    rf_issues = []
    
    # Interference check
    ul_int = cell_data['UL Interference(dBm)']
    if ul_int > -110:
        rf_issues.append(f"🔴 **High UL Interference:** {ul_int:.2f} dBm - Investigate interference sources")
    elif ul_int > -115:
        rf_issues.append(f"🟡 **Moderate UL Interference:** {ul_int:.2f} dBm - Monitor")
    else:
        rf_issues.append(f"🟢 **Good UL Interference:** {ul_int:.2f} dBm")
    
    # IBLER check
    pdsch_ibler = cell_data['PDSCH IBLER(%)']
    if pdsch_ibler > 10:
        rf_issues.append(f"🔴 **High PDSCH IBLER:** {pdsch_ibler:.2f}% - Check DL coverage/quality")
    elif pdsch_ibler > 5:
        rf_issues.append(f"🟡 **Moderate PDSCH IBLER:** {pdsch_ibler:.2f}%")
    else:
        rf_issues.append(f"🟢 **Good PDSCH IBLER:** {pdsch_ibler:.2f}%")
    
    pusch_ibler = cell_data['PUSCH IBLER(%)']
    if pusch_ibler > 10:
        rf_issues.append(f"🔴 **High PUSCH IBLER:** {pusch_ibler:.2f}% - Check UL coverage/quality")
    elif pusch_ibler > 5:
        rf_issues.append(f"🟡 **Moderate PUSCH IBLER:** {pusch_ibler:.2f}%")
    else:
        rf_issues.append(f"🟢 **Good PUSCH IBLER:** {pusch_ibler:.2f}%")
    
    # CQI check
    avg_cqi = cell_data['Avg CQI']
    if avg_cqi < 7:
        rf_issues.append(f"🔴 **Low Avg CQI:** {avg_cqi:.2f} - Poor DL quality")
    elif avg_cqi < 10:
        rf_issues.append(f"🟡 **Moderate Avg CQI:** {avg_cqi:.2f}")
    else:
        rf_issues.append(f"🟢 **Good Avg CQI:** {avg_cqi:.2f}")
    
    # MIMO check
    mimo_rank2 = cell_data['MIMO Rank2']
    if mimo_rank2 < 30:
        rf_issues.append(f"🟡 **Low MIMO Rank2:** {mimo_rank2:.2f}% - Check MIMO configuration")
    else:
        rf_issues.append(f"🟢 **Good MIMO Rank2:** {mimo_rank2:.2f}%")
    
    for issue in rf_issues:
        st.markdown(issue)
    
    st.divider()
    
    # Network-wide RF Quality
    st.subheader("Network-Wide RF Quality")
    
    col1, col2 = st.columns(2)
    
    with col1:
        fig_cqi = go.Figure()
        fig_cqi.add_trace(go.Box(y=df_before['Avg CQI'], name='Avg CQI'))
        fig_cqi.update_layout(title="CQI Distribution Across Network", yaxis_title="CQI")
        fig_cqi.add_hline(y=10, line_dash="dash", line_color="green", annotation_text="Good")
        fig_cqi.add_hline(y=7, line_dash="dash", line_color="orange", annotation_text="Fair")
        st.plotly_chart(fig_cqi, use_container_width=True)
    
    with col2:
        fig_interference = px.scatter(
            df_before,
            x='UL Interference(dBm)',
            y='PUSCH IBLER(%)',
            color='Health_Score',
            hover_data=['Cell_ID'],
            title="Interference vs UL Quality",
            color_continuous_scale='RdYlGn'
        )
        st.plotly_chart(fig_interference, use_container_width=True)

# -------------------------------------------------
# TAB 7: Availability Analysis
# -------------------------------------------------
with tabs[6]:
    st.header("⚡ Availability Analysis")
    
    # Cell Selection
    col1, col2 = st.columns(2)
    
    with col1:
        site_avail = st.selectbox("Select Site", sorted(df_before["Site_ID"].unique()), key="avail_site")
    
    with col2:
        cells_avail = sorted(df_before[df_before["Site_ID"] == site_avail]["Cell_ID"].unique())
        cell_avail = st.selectbox("Select Cell", cells_avail, key="avail_cell")
    
    cell_data = df_before[
        (df_before["Site_ID"] == site_avail) &
        (df_before["Cell_ID"] == cell_avail)
    ].iloc[0]
    
    st.subheader(f"{site_avail} - {cell_avail}")
    
    # Availability Metrics
    col1, col2, col3 = st.columns(3)
    
    with col1:
        availability = cell_data['LTE Network Availability (%)']
        if availability >= 99.9:
            st.success(f"**Network Availability:** {availability:.3f}%")
        elif availability >= 99.0:
            st.warning(f"**Network Availability:** {availability:.3f}%")
        else:
            st.error(f"**Network Availability:** {availability:.3f}%")
    
    with col2:
        downtime = cell_data['Cell Downtime(min)']
        st.metric("Cell Downtime", f"{downtime:.1f} min")
    
    with col3:
        downtime_son = cell_data['Cell Downtime with SON(min)']
        st.metric("Downtime with SON", f"{downtime_son:.1f} min")
    
    st.divider()
    
    # Availability Analysis
    if availability < 99.9:
        st.subheader("📋 Availability Issues")
        
        issues = []
        
        if availability < 99.0:
            issues.append(f"🔴 **Critical:** Availability at {availability:.3f}% - Below 99% threshold")
        elif availability < 99.5:
            issues.append(f"🟡 **Warning:** Availability at {availability:.3f}% - Below 99.5% threshold")
        else:
            issues.append(f"🟢 **Minor:** Availability at {availability:.3f}% - Below 99.9% target")
        
        if downtime > 0:
            issues.append(f"⚠️ **Downtime Recorded:** {downtime:.1f} minutes of total downtime")
        
        if downtime_son > downtime:
            son_extra = downtime_son - downtime
            issues.append(f"ℹ️ **SON Impact:** Additional {son_extra:.1f} minutes downtime with SON")
        
        for issue in issues:
            st.markdown(issue)
        
        st.markdown("**Recommendations:**")
        st.markdown("• Review cell outage logs for root cause")
        st.markdown("• Check power supply stability")
        st.markdown("• Verify transmission links")
        st.markdown("• Review SON configuration if SON downtime is high")
    else:
        st.success("✅ Excellent availability - Cell meeting 99.9% target")
    
    st.divider()
    
    # Network-wide Availability
    st.subheader("Network-Wide Availability")
    
    col1, col2 = st.columns(2)
    
    with col1:
        fig_avail = go.Figure()
        fig_avail.add_trace(go.Histogram(
            x=df_before['LTE Network Availability (%)'],
            nbinsx=20
        ))
        fig_avail.add_vline(x=99.9, line_dash="dash", line_color="green", annotation_text="Target")
        fig_avail.add_vline(x=99.0, line_dash="dash", line_color="red", annotation_text="Critical")
        fig_avail.update_layout(title="Availability Distribution", xaxis_title="Availability (%)")
        st.plotly_chart(fig_avail, use_container_width=True)
    
    with col2:
        # Cells by availability category
        excellent = len(df_before[df_before['LTE Network Availability (%)'] >= 99.9])
        good = len(df_before[(df_before['LTE Network Availability (%)'] >= 99.0) & 
                            (df_before['LTE Network Availability (%)'] < 99.9)])
        poor = len(df_before[df_before['LTE Network Availability (%)'] < 99.0])
        
        fig_cat = px.pie(
            values=[excellent, good, poor],
            names=['Excellent (≥99.9%)', 'Good (99-99.9%)', 'Poor (<99%)'],
            title="Availability Categories",
            color_discrete_sequence=['#2ca02c', '#ffd700', '#d62728']
        )
        st.plotly_chart(fig_cat, use_container_width=True)

# -------------------------------------------------
# TAB 8: Inter-RAT Performance
# -------------------------------------------------
with tabs[7]:
    st.header("🔄 Inter-RAT Performance")
    
    # Cell Selection
    col1, col2 = st.columns(2)
    
    with col1:
        site_rat = st.selectbox("Select Site", sorted(df_before["Site_ID"].unique()), key="rat_site")
    
    with col2:
        cells_rat = sorted(df_before[df_before["Site_ID"] == site_rat]["Cell_ID"].unique())
        cell_rat = st.selectbox("Select Cell", cells_rat, key="rat_cell")
    
    cell_data = df_before[
        (df_before["Site_ID"] == site_rat) &
        (df_before["Cell_ID"] == cell_rat)
    ].iloc[0]
    
    st.subheader(f"{site_rat} - {cell_rat}")
    
    # Inter-RAT Metrics
    col1, col2, col3, col4 = st.columns(4)
    
    col1.metric("RRC Redir E2G", f"{cell_data['RRC Redirection E2G']:.0f}")
    col2.metric("RRC Redir E2G (Blind)", f"{cell_data['RRC Redirection E2G (Blind)']:.0f}")
    col3.metric("CSFB Attempt E2G", f"{cell_data['CSFB Attempt E2G']:.0f}")
    col4.metric("CSFB Attempt E2G (Flash)", f"{cell_data['CSFB Attempt E2G (Flash)']:.0f}")
    
    st.divider()
    
    # Analysis
    st.subheader("Inter-RAT Analysis")
    
    total_redir = cell_data['RRC Redirection E2G'] + cell_data['RRC Redirection E2G (Blind)']
    total_csfb = cell_data['CSFB Attempt E2G'] + cell_data['CSFB Attempt E2G (Flash)']
    
    col1, col2 = st.columns(2)
    
    with col1:
        if total_redir > 0:
            fig_redir = px.pie(
                values=[cell_data['RRC Redirection E2G'], cell_data['RRC Redirection E2G (Blind)']],
                names=['Normal Redirection', 'Blind Redirection'],
                title="RRC Redirection Types"
            )
            st.plotly_chart(fig_redir, use_container_width=True)
            
            blind_pct = (cell_data['RRC Redirection E2G (Blind)'] / total_redir) * 100
            if blind_pct > 50:
                st.warning(f"⚠️ High blind redirection: {blind_pct:.1f}% - Review neighbor configuration")
        else:
            st.info("No RRC redirections recorded")
    
    with col2:
        if total_csfb > 0:
            fig_csfb = px.pie(
                values=[cell_data['CSFB Attempt E2G'], cell_data['CSFB Attempt E2G (Flash)']],
                names=['Normal CSFB', 'Flash CSFB'],
                title="CSFB Types"
            )
            st.plotly_chart(fig_csfb, use_container_width=True)
        else:
            st.info("No CSFB attempts recorded")
    
    # Additional Features
    col1, col2 = st.columns(2)
    
    col1.metric("Smart Carrier Feature", f"{cell_data['Smart Carrier Feature']:.0f}")
    col2.metric("Paging Discarded", f"{cell_data['Paging Discarded']:.0f}")
    
    if cell_data['Paging Discarded'] > 0:
        st.warning(f"⚠️ Paging discards detected: {cell_data['Paging Discarded']:.0f} - May indicate congestion")

# -------------------------------------------------
# TAB 9: Multi-Cell Comparison
# -------------------------------------------------
with tabs[8]:
    st.header("📋 Multi-Cell Comprehensive Comparison")
    
    # KPI Category Selection
    kpi_category = st.selectbox(
        "Select KPI Category",
        ["Overview", "Coverage", "Quality", "Capacity", "VoLTE", "RF", "Availability"]
    )
    
    # Prepare comparison data
    if kpi_category == "Overview":
        compare_cols = ['Cell_ID', 'Health_Score', 'Verdict', 'LTE Network Availability (%)',
                       'RRC CSSR(%)', 'ERAB DCR(%)', 'HO SR(%)', 'DL PRB Utilization(%)',
                       'Avg TA Distance(m)']
        display_names = ['Cell', 'Health', 'Status', 'Availability', 'RRC CSSR', 
                        'ERAB DCR', 'HO SR', 'DL PRB Util', 'Avg TA Dist']
    
    elif kpi_category == "Coverage":
        compare_cols = ['Cell_ID', 'Avg TA Distance(m)', 'DL PRB Utilization(%)', 
                       'UL PRB Utilization(%)', 'Avg CQI']
        display_names = ['Cell', 'Avg TA Distance', 'DL PRB Util', 'UL PRB Util', 'Avg CQI']
    
    elif kpi_category == "Quality":
        compare_cols = ['Cell_ID', 'RRC CSSR(%)', 'ERAB CSSR(%)', 'ERAB DCR(%)', 
                       'HO SR(%)', 'CSFB SR(%)']
        display_names = ['Cell', 'RRC CSSR', 'ERAB CSSR', 'ERAB DCR', 'HO SR', 'CSFB SR']
    
    elif kpi_category == "Capacity":
        compare_cols = ['Cell_ID', 'Traffic User(Avg)', 'Traffic User(Max)',
                       'DL PRB Utilization(%)', 'UL PRB Utilization(%)',
                       'DL Throughput(Mbit/s)', 'UL Throughput(Mbit/s)']
        display_names = ['Cell', 'Avg Users', 'Max Users', 'DL PRB', 'UL PRB', 
                        'DL Tput', 'UL Tput']
    
    elif kpi_category == "VoLTE":
        compare_cols = ['Cell_ID', 'VoLTE CSSR(%)', 'VoLTE DCR(%)', 'VoLTE Traffic (Erl)',
                       'VoLTE Drop due Radio', 'VoLTE Drop due Congestion']
        display_names = ['Cell', 'VoLTE CSSR', 'VoLTE DCR', 'VoLTE Traffic',
                        'Radio Drops', 'Congestion Drops']
    
    elif kpi_category == "RF":
        compare_cols = ['Cell_ID', 'UL Interference(dBm)', 'PDSCH IBLER(%)', 
                       'PUSCH IBLER(%)', 'Avg CQI', 'MIMO Rank2']
        display_names = ['Cell', 'UL Interference', 'PDSCH IBLER', 'PUSCH IBLER',
                        'Avg CQI', 'MIMO Rank2']
    
    else:  # Availability
        compare_cols = ['Cell_ID', 'LTE Network Availability (%)', 
                       'Cell Downtime(min)', 'Cell Downtime with SON(min)']
        display_names = ['Cell', 'Availability', 'Downtime', 'Downtime w/ SON']
    
    # Display comparison table
    compare_df = df_before[compare_cols].round(2)
    compare_df.columns = display_names
    
    # Sort options
    col1, col2 = st.columns([3, 1])
    with col1:
        sort_by = st.selectbox("Sort by", display_names[1:])
    with col2:
        sort_order = st.radio("Order", ["Descending", "Ascending"], horizontal=True)
    
    ascending = (sort_order == "Ascending")
    compare_df_sorted = compare_df.sort_values(by=sort_by, ascending=ascending)
    
    st.dataframe(compare_df_sorted, use_container_width=True, hide_index=True)
    
    st.divider()
    
    # Visualization
    st.subheader("Visual Comparison")
    
    if kpi_category == "Overview":
        fig = px.scatter(
            df_before,
            x='RRC CSSR(%)',
            y='ERAB DCR(%)',
            size='DL PRB Utilization(%)',
            color='Health_Score',
            hover_data=['Cell_ID'],
            labels={'RRC CSSR(%)': 'RRC CSSR (%)', 'ERAB DCR(%)': 'ERAB DCR (%)'},
            color_continuous_scale='RdYlGn',
            title="Quality Overview: RRC CSSR vs ERAB DCR"
        )
        st.plotly_chart(fig, use_container_width=True)
    
    else:
        # Bar chart for selected category
        plot_col = compare_cols[1]  # First metric after Cell_ID
        fig = px.bar(
            compare_df_sorted.head(20),
            x='Cell',
            y=display_names[1],
            title=f"Top 20 Cells by {display_names[1]}"
        )
        fig.update_layout(xaxis_tickangle=-45)
        st.plotly_chart(fig, use_container_width=True)

# -------------------------------------------------
# TAB 10: Export & Reports
# -------------------------------------------------
with tabs[9]:
    st.header("💾 Export & Reports")
    
    st.subheader("Generate Comprehensive Report")
    
    report_type = st.selectbox(
        "Report Type",
        ["Complete Network Report", "Coverage Analysis", "Quality Report", 
         "Capacity Report", "VoLTE Analysis", "Problem Cells Only"]
    )
    
    if st.button("📊 Generate Report", type="primary"):
        with st.spinner("Generating comprehensive report..."):
            
            output = io.BytesIO()
            
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                # Executive Summary
                summary_data = {
                    'Metric': ['Total Cells', 'Avg Health Score', 'Excellent Cells', 'Critical Cells',
                              'Avg RRC CSSR', 'Avg ERAB DCR', 'Avg HO SR', 'Avg Availability'],
                    'Value': [
                        len(df_before),
                        f"{df_before['Health_Score'].mean():.1f}",
                        len(df_before[df_before['Health_Score'] >= 80]),
                        len(df_before[df_before['Health_Score'] < 40]),
                        f"{df_before['RRC CSSR(%)'].mean():.2f}%",
                        f"{df_before['ERAB DCR(%)'].mean():.2f}%",
                        f"{df_before['HO SR(%)'].mean():.2f}%",
                        f"{df_before['LTE Network Availability (%)'].mean():.2f}%"
                    ]
                }
                pd.DataFrame(summary_data).to_excel(writer, sheet_name='Executive Summary', index=False)
                
                # Full Data
                df_before.to_excel(writer, sheet_name='Complete Data', index=False)
                
                # Problem Cells
                problem_cells = df_before[df_before['Health_Score'] < 60]
                if len(problem_cells) > 0:
                    problem_cells.to_excel(writer, sheet_name='Problem Cells', index=False)
                
                # Quality Issues
                quality_issues = df_before[
                    (df_before['RRC CSSR(%)'] < thresholds['rrc_cssr']) |
                    (df_before['ERAB DCR(%)'] > thresholds['erab_dcr']) |
                    (df_before['HO SR(%)'] < thresholds['ho_sr'])
                ]
                if len(quality_issues) > 0:
                    quality_issues.to_excel(writer, sheet_name='Quality Issues', index=False)
                
                # Capacity Issues
                capacity_issues = df_before[
                    (df_before['DL PRB Utilization(%)'] > 70) |
                    (df_before['UL PRB Utilization(%)'] > 50)
                ]
                if len(capacity_issues) > 0:
                    capacity_issues.to_excel(writer, sheet_name='Capacity Issues', index=False)
            
            output.seek(0)
            
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            
            st.success("✅ Report generated successfully!")
            
            st.download_button(
                label="⬇️ Download Excel Report",
                data=output,
                file_name=f"LTE_Complete_Report_{timestamp}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

# -------------------------------------------------
# Footer
# -------------------------------------------------
st.divider()
st.caption("Professional Network Analysis Platform | Fadzli Abdullah | Huawei Technologies.")