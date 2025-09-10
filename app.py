import streamlit as st
import pandas as pd
import datetime
from datetime import timedelta
import json
import os
from io import BytesIO
import plotly.express as px
import plotly.graph_objects as go
import numpy as np

# Page config
st.set_page_config(
    page_title="OCM Team Travel Tracker - 2025",
    page_icon="🌍",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Custom CSS for modern design
st.markdown("""
    <style>
    /* Clean white background */
    .stApp {
        background-color: #ffffff;
    }
    
    /* Modern button styling for day cells */
    .stButton > button {
        background: linear-gradient(135deg, #f3f4f6, #e5e7eb) !important;
        border: 2px solid #d1d5db !important;
        border-radius: 16px !important;
        padding: 20px !important;
        font-size: 28px !important;
        height: 80px !important;
        transition: all 0.2s ease !important;
        box-shadow: 0 2px 8px rgba(0,0,0,0.05) !important;
        color: #374151 !important;
    }
    
    .stButton > button:hover {
        transform: scale(1.05) !important;
        box-shadow: 0 4px 16px rgba(0,0,0,0.15) !important;
    }
    
    /* Business travel style - primary buttons */
    .stButton > button[kind="primary"] {
        background: linear-gradient(135deg, #6B46C1, #9333EA) !important;
        border: none !important;
        color: white !important;
        box-shadow: 0 4px 12px rgba(147, 51, 234, 0.3) !important;
    }
    
    /* Vacation style - secondary buttons */
    .stButton > button[kind="secondary"] {
        background: linear-gradient(135deg, #F59E0B, #FCD34D) !important;
        border: none !important;
        color: white !important;
        box-shadow: 0 4px 12px rgba(245, 158, 11, 0.3) !important;
    }
    
    /* Metrics styling */
    [data-testid="metric-container"] {
        background: white;
        padding: 20px;
        border-radius: 12px;
        border: 1px solid #f0f0f0;
        box-shadow: 0 2px 8px rgba(0,0,0,0.04);
    }
    
    /* Tab styling */
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
        background-color: #f9fafb;
        padding: 4px;
        border-radius: 12px;
    }
    
    .stTabs [data-baseweb="tab"] {
        height: 40px;
        padding: 0 16px;
        background: white;
        border-radius: 8px;
        color: #6b7280;
        font-weight: 500;
    }
    
    .stTabs [aria-selected="true"] {
        background: linear-gradient(135deg, #6B46C1, #9333EA);
        color: white;
    }
    
    /* Hide Streamlit menu */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    </style>
    """, unsafe_allow_html=True)

# Initialize session state
if 'travel_data' not in st.session_state:
    st.session_state.travel_data = {}

if 'budget_data' not in st.session_state:
    st.session_state.budget_data = {
        'annual_budget': 150000,
        'default_daily_rate': 500
    }

if 'team_members' not in st.session_state:
    st.session_state.team_members = [
        {"name": "Torsten Gadfelt", "role": "Manager (MD)"},
        {"name": "Niels Woersaa", "role": "Senior OCM Analyst"},
        {"name": "Billie Hoey", "role": "Senior OCM Analyst"},
        {"name": "Tommi Sjcholm", "role": "OCM PMO"},
        {"name": "Tomi Ruippo", "role": "OCM Advisor"},
        {"name": "Annalisa Giotta", "role": "OCM Advisor"},
        {"name": "Jesse Hupkens", "role": "OCM Advisor"}
    ]

# Helper functions
def get_week_dates(date):
    """Get Monday to Friday dates for a given week"""
    start = date - timedelta(days=date.weekday())
    return [(start + timedelta(days=i)) for i in range(5)]

def get_week_range(date):
    """Get formatted week range string"""
    monday = date - timedelta(days=date.weekday())
    friday = monday + timedelta(days=4)
    return f"{monday.strftime('%B %d')} - {friday.strftime('%B %d, %Y')}"

def cycle_status(current_data):
    """Cycle: Office → Business Travel → Vacation → Office"""
    if not current_data or current_data.get('status') == 'office':
        return 'business'
    elif current_data.get('status') == 'business':
        return 'vacation'
    else:
        return 'office'

def calculate_weekly_stats(week_start):
    """Calculate statistics for the current week"""
    week_dates = get_week_dates(week_start)
    business_days = 0
    vacation_days = 0
    business_travelers = set()
    week_cost = 0
    
    for member in st.session_state.team_members:
        has_business_travel = False
        for date in week_dates:
            date_str = date.strftime('%Y-%m-%d')
            key = f"{member['name']}_{date_str}"
            data = st.session_state.travel_data.get(key, {})
            
            if data.get('status') == 'business':
                business_days += 1
                has_business_travel = True
                week_cost += data.get('daily_cost', 500)
            elif data.get('status') == 'vacation':
                vacation_days += 1
        
        # Only count as "traveling" if they have business travel
        if has_business_travel:
            business_travelers.add(member['name'])
    
    # Month stats
    month_start = week_start.replace(day=1)
    month_end = (month_start + timedelta(days=32)).replace(day=1) - timedelta(days=1)
    month_business_days = 0
    month_cost = 0
    
    current = month_start
    while current <= month_end:
        if current.weekday() < 5:  # Weekdays only
            for member in st.session_state.team_members:
                date_str = current.strftime('%Y-%m-%d')
                key = f"{member['name']}_{date_str}"
                data = st.session_state.travel_data.get(key, {})
                if data.get('status') == 'business':
                    month_business_days += 1
                    month_cost += data.get('daily_cost', 500)
        current += timedelta(days=1)
    
    return {
        'business_days': business_days,
        'vacation_days': vacation_days,
        'business_travelers': len(business_travelers),
        'month_business_days': month_business_days,
        'week_cost': week_cost,
        'month_cost': month_cost
    }

def export_to_excel():
    """Export data to Excel format"""
    records = []
    for key, value in st.session_state.travel_data.items():
        if value.get('status') in ['business', 'vacation']:
            name, date = key.rsplit('_', 1)
            records.append({
                'Team Member': name,
                'Date': date,
                'Status': value.get('status', '').title(),
                'Cost': value.get('daily_cost', 0) if value.get('status') == 'business' else 0
            })
    
    if records:
        df = pd.DataFrame(records)
        df['Date'] = pd.to_datetime(df['Date'])
        df = df.sort_values(['Date', 'Team Member'])
        
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Travel Data', index=False)
        output.seek(0)
        return output
    return None

# MAIN APP
st.markdown("<h1 style='text-align: center; color: #6B46C1; margin-bottom: 0;'>🌍 OCM Team Travel Tracker</h1>", unsafe_allow_html=True)
st.markdown(f"<p style='text-align: center; color: #9CA3AF; margin-top: 0;'>{datetime.date.today().strftime('%A, %B %d, %Y')}</p>", unsafe_allow_html=True)

# Tabs
tab1, tab2, tab3, tab4 = st.tabs(["📅 Calendar", "📊 Analytics", "💰 Budget", "⚙️ Settings"])

with tab1:
    # Week navigation
    col1, col2, col3 = st.columns([1, 3, 1])
    
    if 'current_week' not in st.session_state:
        st.session_state.current_week = datetime.date.today()
    
    with col1:
        if st.button("← Previous", use_container_width=True):
            st.session_state.current_week -= timedelta(weeks=1)
    
    with col2:
        st.markdown(f"<h3 style='text-align: center; color: #374151;'>Week of {get_week_range(st.session_state.current_week)}</h3>", unsafe_allow_html=True)
    
    with col3:
        if st.button("Next →", use_container_width=True):
            st.session_state.current_week += timedelta(weeks=1)
    
    # Today button
    if st.button("📍 Today", use_container_width=True):
        st.session_state.current_week = datetime.date.today()
    
    st.markdown("---")
    
    # Calendar header
    week_dates = get_week_dates(st.session_state.current_week)
    weekdays = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]
    
    header_cols = st.columns([2] + [1]*5)
    header_cols[0].markdown("**Team Member**")
    
    for i, (day, date) in enumerate(zip(weekdays, week_dates)):
        is_today = date == datetime.date.today()
        if is_today:
            header_cols[i+1].markdown(f"**{day}**<br>📍 {date.strftime('%b %d')}", unsafe_allow_html=True)
        else:
            header_cols[i+1].markdown(f"**{day}**<br>{date.strftime('%b %d')}", unsafe_allow_html=True)
    
    # Team member rows
    for member in st.session_state.team_members:
        cols = st.columns([2] + [1]*5)
        
        # Name and role
        cols[0].markdown(f"**{member['name']}**<br><small>{member['role']}</small>", unsafe_allow_html=True)
        
        # Day buttons
      # Day buttons
        for i, date in enumerate(week_dates):
            date_str = date.strftime('%Y-%m-%d')
            key = f"{member['name']}_{date_str}"
            data = st.session_state.travel_data.get(key, {'status': 'office'})
            
            # Determine button appearance
            status = data.get('status', 'office')
            if status == 'business':
                icon = "✈️"
                button_clicked = cols[i+1].button(icon, key=f"btn_{key}", use_container_width=True, type="primary")
            elif status == 'vacation':
                icon = "🏖️"
                button_clicked = cols[i+1].button(icon, key=f"btn_{key}", use_container_width=True, type="secondary")
            else:
                icon = "🏢"
                # Don't specify type for office days - use default
                button_clicked = cols[i+1].button(icon, key=f"btn_{key}", use_container_width=True)
            
            if button_clicked:
                # Cycle to next status
                next_status = cycle_status(data)
                
                if next_status == 'office':
                    st.session_state.travel_data[key] = {'status': 'office'}
                elif next_status == 'business':
                    st.session_state.travel_data[key] = {
                        'status': 'business',
                        'daily_cost': st.session_state.budget_data['default_daily_rate']
                    }
                else:  # vacation
                    st.session_state.travel_data[key] = {'status': 'vacation'}
                
                st.rerun()
    
    # Statistics
    st.markdown("---")
    stats = calculate_weekly_stats(st.session_state.current_week)
    
    col1, col2, col3, col4, col5, col6 = st.columns(6)
    col1.metric("Business Days", stats['business_days'])
    col2.metric("Vacation Days", stats['vacation_days'])
    col3.metric("Team Traveling", stats['business_travelers'], help="People on business travel only")
    col4.metric("Month Travel", stats['month_business_days'])
    col5.metric("Week Cost", f"${stats['week_cost']:,.0f}")
    col6.metric("Month Cost", f"${stats['month_cost']:,.0f}")
    
    # Legend
    with st.expander("📖 How to Use"):
        st.markdown("""
        **Click any day to cycle through:**
        - 🏢 **Office/Remote** (default)
        - ✈️ **Business Travel** (purple)
        - 🏖️ **Vacation** (yellow)
        
        **Note:** "Team Traveling" only counts business travel, not vacations.
        """)

with tab2:
    st.header("📊 Analytics")
    
    # Prepare data for visualizations
    travel_by_person = {}
    vacation_by_person = {}
    
    for key, value in st.session_state.travel_data.items():
        name = key.rsplit('_', 1)[0]
        if value.get('status') == 'business':
            travel_by_person[name] = travel_by_person.get(name, 0) + 1
        elif value.get('status') == 'vacation':
            vacation_by_person[name] = vacation_by_person.get(name, 0) + 1
    
    col1, col2 = st.columns(2)
    
    with col1:
        # Business travel chart
        if travel_by_person:
            df = pd.DataFrame(list(travel_by_person.items()), columns=['Person', 'Days'])
            fig = px.bar(df, x='Person', y='Days', title='Business Travel Days',
                        color_discrete_sequence=['#9333EA'])
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("No business travel data yet")
    
    with col2:
        # Vacation chart
        if vacation_by_person:
            df = pd.DataFrame(list(vacation_by_person.items()), columns=['Person', 'Days'])
            fig = px.bar(df, x='Person', y='Days', title='Vacation Days',
                        color_discrete_sequence=['#F59E0B'])
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("No vacation data yet")

with tab3:
    st.header("💰 Budget Management")
    
    col1, col2 = st.columns(2)
    
    with col1:
        # Budget settings
        new_budget = st.number_input("Annual Budget ($)", 
                                    value=st.session_state.budget_data['annual_budget'],
                                    step=1000)
        if new_budget != st.session_state.budget_data['annual_budget']:
            st.session_state.budget_data['annual_budget'] = new_budget
        
        daily_rate = st.number_input("Daily Rate ($)",
                                    value=st.session_state.budget_data['default_daily_rate'],
                                    step=50)
        if daily_rate != st.session_state.budget_data['default_daily_rate']:
            st.session_state.budget_data['default_daily_rate'] = daily_rate
    
    with col2:
        # Calculate spending
        total_spent = sum(v.get('daily_cost', 0) for v in st.session_state.travel_data.values() 
                         if v.get('status') == 'business')
        remaining = st.session_state.budget_data['annual_budget'] - total_spent
        
        # Budget chart
        fig = go.Figure(data=[go.Pie(
            labels=['Spent', 'Remaining'],
            values=[total_spent, max(0, remaining)],
            hole=.4,
            marker_colors=['#9333EA', '#E5E7EB']
        )])
        fig.update_layout(title="Budget Status")
        st.plotly_chart(fig, use_container_width=True)
        
        st.metric("Total Spent", f"${total_spent:,.0f}")
        st.metric("Remaining", f"${remaining:,.0f}")

with tab4:
    st.header("⚙️ Settings")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("Export Data")
        
        excel_file = export_to_excel()
        if excel_file:
            st.download_button(
                "📥 Download Excel Report",
                data=excel_file,
                file_name=f"travel_report_{datetime.date.today()}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        
        # Save/Load JSON
        if st.button("💾 Save Configuration"):
            config = {
                'travel_data': st.session_state.travel_data,
                'budget_data': st.session_state.budget_data,
                'team_members': st.session_state.team_members
            }
            json_str = json.dumps(config, indent=2)
            st.download_button(
                "Download Config",
                data=json_str,
                file_name=f"config_{datetime.date.today()}.json",
                mime="application/json"
            )
    
    with col2:
        st.subheader("Team Management")
        
        # Add member
        with st.expander("Add Team Member"):
            new_name = st.text_input("Name")
            new_role = st.text_input("Role")
            if st.button("Add"):
                if new_name and new_role:
                    st.session_state.team_members.append({"name": new_name, "role": new_role})
                    st.success(f"Added {new_name}")
                    st.rerun()

# Sidebar
with st.sidebar:
    st.header("📊 Year Summary")
    
    # Calculate totals
    total_business = sum(1 for v in st.session_state.travel_data.values() 
                        if v.get('status') == 'business')
    total_vacation = sum(1 for v in st.session_state.travel_data.values() 
                        if v.get('status') == 'vacation')
    total_cost = sum(v.get('daily_cost', 0) for v in st.session_state.travel_data.values() 
                    if v.get('status') == 'business')
    
    st.metric("YTD Business Travel", f"{total_business} days")
    st.metric("YTD Vacation", f"{total_vacation} days")
    st.metric("YTD Cost", f"${total_cost:,.0f}")
    
    st.markdown("---")
    st.markdown("""
    **Legend:**
    - 🏢 Office/Remote
    - ✈️ Business Travel  
    - 🏖️ Vacation
    
    Click days to cycle through statuses.
    """)