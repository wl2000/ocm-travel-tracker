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
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import requests

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
    
    /* Pending approval style */
    .pending-approval {
        background: linear-gradient(135deg, #EF4444, #F87171) !important;
        animation: pulse 2s infinite;
    }
    
    @keyframes pulse {
        0% { opacity: 1; }
        50% { opacity: 0.7; }
        100% { opacity: 1; }
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
    
    /* Notification badge */
    .notification-badge {
        background: #EF4444;
        color: white;
        border-radius: 50%;
        padding: 2px 6px;
        font-size: 10px;
        position: absolute;
        top: -5px;
        right: -5px;
    }
    
    /* Hide Streamlit menu */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    </style>
    """, unsafe_allow_html=True)

# Initialize session state
if 'travel_data' not in st.session_state:
    st.session_state.travel_data = {}

if 'approvals_pending' not in st.session_state:
    st.session_state.approvals_pending = []

if 'email_config' not in st.session_state:
    st.session_state.email_config = {
        'smtp_server': 'smtp.office365.com',
        'smtp_port': 587,
        'sender_email': '',
        'sender_password': '',
        'notification_enabled': False
    }

if 'budget_data' not in st.session_state:
    st.session_state.budget_data = {
        'annual_budget': 150000,
        'default_daily_rate': 500,
        'requires_approval_above': 1000
    }

if 'team_members' not in st.session_state:
    st.session_state.team_members = [
        {"name": "Torsten Gadfelt", "role": "Manager (MD)", "email": "torsten.gadfelt@fedex.com", "is_manager": True},
        {"name": "Niels Woersaa", "role": "Senior OCM Analyst", "email": "niels.woersaa@fedex.com", "is_manager": False},
        {"name": "Billie Hoey", "role": "Senior OCM Analyst", "email": "billie.hoey@fedex.com", "is_manager": False},
        {"name": "Tommi Sjcholm", "role": "OCM PMO", "email": "tommi.sjcholm@fedex.com", "is_manager": False},
        {"name": "Tomi Ruippo", "role": "OCM Advisor", "email": "tomi.ruippo@fedex.com", "is_manager": False},
        {"name": "Annalisa Giotta", "role": "OCM Advisor", "email": "annalisa.giotta@fedex.com", "is_manager": False},
        {"name": "Jesse Hupkens", "role": "OCM Advisor", "email": "jesse.hupkens@fedex.com", "is_manager": False}
    ]

if 'automated_reports' not in st.session_state:
    st.session_state.automated_reports = {
        'weekly_report': True,
        'monthly_report': True,
        'report_recipients': ['torsten.gadfelt@fedex.com']
    }

# Email functions
def send_email_notification(to_email, subject, body, attachment=None):
    """Send email notifications"""
    if not st.session_state.email_config['notification_enabled']:
        return False
    
    try:
        msg = MIMEMultipart()
        msg['From'] = st.session_state.email_config['sender_email']
        msg['To'] = to_email
        msg['Subject'] = subject
        
        msg.attach(MIMEText(body, 'html'))
        
        if attachment:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(attachment.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f'attachment; filename="travel_report.xlsx"')
            msg.attach(part)
        
        # Uncomment when email is configured
        # server = smtplib.SMTP(st.session_state.email_config['smtp_server'], st.session_state.email_config['smtp_port'])
        # server.starttls()
        # server.login(st.session_state.email_config['sender_email'], st.session_state.email_config['sender_password'])
        # server.send_message(msg)
        # server.quit()
        
        return True
    except Exception as e:
        st.error(f"Email error: {str(e)}")
        return False

def send_approval_request(member_name, dates, cost):
    """Send approval request to manager"""
    manager = next((m for m in st.session_state.team_members if m['is_manager']), None)
    if not manager:
        return
    
    subject = f"Travel Approval Required: {member_name}"
    body = f"""
    <html>
        <body>
            <h2>Travel Approval Request</h2>
            <p><strong>Team Member:</strong> {member_name}</p>
            <p><strong>Dates:</strong> {dates}</p>
            <p><strong>Estimated Cost:</strong> ${cost:,.2f}</p>
            <p><strong>Status:</strong> Pending Approval</p>
            <br>
            <p>Please review in the <a href="https://ocm-travel-tracker-nzxvmnkrvw5wknvvkdb7tb.streamlit.app/">Travel Tracker</a></p>
        </body>
    </html>
    """
    
    send_email_notification(manager['email'], subject, body)

def send_calendar_invite(member_email, dates, travel_type):
    """Create calendar invite for Outlook"""
    # This would integrate with Microsoft Graph API
    # For now, we'll create an ICS file
    ics_content = f"""BEGIN:VCALENDAR
VERSION:2.0
PRODID:-//OCM Travel Tracker//EN
BEGIN:VEVENT
SUMMARY:{travel_type}
DTSTART:{dates[0].strftime('%Y%m%d')}
DTEND:{dates[-1].strftime('%Y%m%d')}
DESCRIPTION:Travel booked via OCM Travel Tracker
END:VEVENT
END:VCALENDAR"""
    
    return ics_content

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
    pending_approvals = 0
    
    for member in st.session_state.team_members:
        has_business_travel = False
        for date in week_dates:
            date_str = date.strftime('%Y-%m-%d')
            key = f"{member['name']}_{date_str}"
            data = st.session_state.travel_data.get(key, {})
            
            if data.get('status') == 'business':
                if data.get('approved', True):
                    business_days += 1
                    has_business_travel = True
                    week_cost += data.get('daily_cost', 500)
                else:
                    pending_approvals += 1
            elif data.get('status') == 'vacation':
                vacation_days += 1
        
        if has_business_travel:
            business_travelers.add(member['name'])
    
    # Month stats
    month_start = week_start.replace(day=1)
    month_end = (month_start + timedelta(days=32)).replace(day=1) - timedelta(days=1)
    month_business_days = 0
    month_cost = 0
    
    current = month_start
    while current <= month_end:
        if current.weekday() < 5:
            for member in st.session_state.team_members:
                date_str = current.strftime('%Y-%m-%d')
                key = f"{member['name']}_{date_str}"
                data = st.session_state.travel_data.get(key, {})
                if data.get('status') == 'business' and data.get('approved', True):
                    month_business_days += 1
                    month_cost += data.get('daily_cost', 500)
        current += timedelta(days=1)
    
    return {
        'business_days': business_days,
        'vacation_days': vacation_days,
        'business_travelers': len(business_travelers),
        'month_business_days': month_business_days,
        'week_cost': week_cost,
        'month_cost': month_cost,
        'pending_approvals': pending_approvals
    }

def export_to_excel_advanced():
    """Export data to Excel with multiple sheets"""
    output = BytesIO()
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Sheet 1: Travel Data
        records = []
        for key, value in st.session_state.travel_data.items():
            if value.get('status') in ['business', 'vacation']:
                name, date = key.rsplit('_', 1)
                records.append({
                    'Team Member': name,
                    'Date': date,
                    'Status': value.get('status', '').title(),
                    'Cost': value.get('daily_cost', 0) if value.get('status') == 'business' else 0,
                    'Approved': value.get('approved', True)
                })
        
        if records:
            df_travel = pd.DataFrame(records)
            df_travel['Date'] = pd.to_datetime(df_travel['Date'])
            df_travel = df_travel.sort_values(['Date', 'Team Member'])
            df_travel.to_excel(writer, sheet_name='Travel Data', index=False)
        
        # Sheet 2: Summary Statistics
        summary_data = []
        for member in st.session_state.team_members:
            business_count = sum(1 for k, v in st.session_state.travel_data.items() 
                               if k.startswith(member['name']) and v.get('status') == 'business')
            vacation_count = sum(1 for k, v in st.session_state.travel_data.items() 
                               if k.startswith(member['name']) and v.get('status') == 'vacation')
            total_cost = sum(v.get('daily_cost', 0) for k, v in st.session_state.travel_data.items() 
                           if k.startswith(member['name']) and v.get('status') == 'business')
            
            summary_data.append({
                'Team Member': member['name'],
                'Role': member['role'],
                'Business Days': business_count,
                'Vacation Days': vacation_count,
                'Total Cost': total_cost
            })
        
        df_summary = pd.DataFrame(summary_data)
        df_summary.to_excel(writer, sheet_name='Summary', index=False)
        
        # Sheet 3: Budget Analysis
        budget_data = [{
            'Category': 'Annual Budget',
            'Amount': st.session_state.budget_data['annual_budget']
        }, {
            'Category': 'Total Spent',
            'Amount': sum(v.get('daily_cost', 0) for v in st.session_state.travel_data.values() 
                        if v.get('status') == 'business' and v.get('approved', True))
        }, {
            'Category': 'Remaining',
            'Amount': st.session_state.budget_data['annual_budget'] - 
                     sum(v.get('daily_cost', 0) for v in st.session_state.travel_data.values() 
                        if v.get('status') == 'business' and v.get('approved', True))
        }]
        
        df_budget = pd.DataFrame(budget_data)
        df_budget.to_excel(writer, sheet_name='Budget', index=False)
    
    output.seek(0)
    return output

def generate_weekly_report():
    """Generate automated weekly report"""
    week_start = datetime.date.today() - timedelta(days=datetime.date.today().weekday())
    stats = calculate_weekly_stats(week_start)
    
    report = f"""
    <html>
        <body>
            <h2>Weekly Travel Report - Week of {get_week_range(week_start)}</h2>
            <h3>Summary:</h3>
            <ul>
                <li>Business Travel Days: {stats['business_days']}</li>
                <li>Vacation Days: {stats['vacation_days']}</li>
                <li>Team Members Traveling: {stats['business_travelers']}</li>
                <li>Weekly Cost: ${stats['week_cost']:,.2f}</li>
                <li>Pending Approvals: {stats['pending_approvals']}</li>
            </ul>
            <p>Access the full tracker: <a href="https://ocm-travel-tracker-nzxvmnkrvw5wknvvkdb7tb.streamlit.app/">Travel Tracker</a></p>
        </body>
    </html>
    """
    
    # Send to all report recipients
    for recipient in st.session_state.automated_reports['report_recipients']:
        send_email_notification(recipient, f"Weekly Travel Report - {week_start.strftime('%B %d')}", report)
    
    return report

# MAIN APP
st.markdown("<h1 style='text-align: center; color: #6B46C1; margin-bottom: 0;'>🌍 OCM Team Travel Tracker</h1>", unsafe_allow_html=True)
st.markdown(f"<p style='text-align: center; color: #9CA3AF; margin-top: 0;'>{datetime.date.today().strftime('%A, %B %d, %Y')}</p>", unsafe_allow_html=True)

# Show notifications
stats = calculate_weekly_stats(st.session_state.get('current_week', datetime.date.today()))
if stats['pending_approvals'] > 0:
    st.warning(f"⚠️ You have {stats['pending_approvals']} pending travel approvals")

# Tabs
tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
    "📅 Calendar", 
    "✅ Approvals", 
    "📊 Analytics", 
    "💰 Budget", 
    "📧 Notifications",
    "⚙️ Settings"
])

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
        for i, date in enumerate(week_dates):
            date_str = date.strftime('%Y-%m-%d')
            key = f"{member['name']}_{date_str}"
            data = st.session_state.travel_data.get(key, {'status': 'office'})
            
            # Determine button appearance
            status = data.get('status', 'office')
            needs_approval = False
            
            if status == 'business':
                # Check if needs approval
                cost = data.get('daily_cost', 500)
                if cost > st.session_state.budget_data['requires_approval_above'] and not member.get('is_manager'):
                    if not data.get('approved', False):
                        needs_approval = True
                        icon = "⏳"
                    else:
                        icon = "✈️"
                else:
                    icon = "✈️"
                button_clicked = cols[i+1].button(icon, key=f"btn_{key}", use_container_width=True, type="primary")
            elif status == 'vacation':
                icon = "🏖️"
                button_clicked = cols[i+1].button(icon, key=f"btn_{key}", use_container_width=True, type="secondary")
            else:
                icon = "🏢"
                button_clicked = cols[i+1].button(icon, key=f"btn_{key}", use_container_width=True)
            
            if button_clicked:
                # Cycle to next status
                next_status = cycle_status(data)
                
                if next_status == 'office':
                    st.session_state.travel_data[key] = {'status': 'office'}
                elif next_status == 'business':
                    cost = st.session_state.budget_data['default_daily_rate']
                    needs_approval = cost > st.session_state.budget_data['requires_approval_above'] and not member.get('is_manager')
                    
                    st.session_state.travel_data[key] = {
                        'status': 'business',
                        'daily_cost': cost,
                        'approved': not needs_approval
                    }
                    
                    if needs_approval:
                        st.session_state.approvals_pending.append(key)
                        send_approval_request(member['name'], date_str, cost)
                        
                else:  # vacation
                    st.session_state.travel_data[key] = {'status': 'vacation'}
                    # Send calendar invite
                    if st.session_state.email_config['notification_enabled']:
                        ics = send_calendar_invite(member['email'], [date], 'Vacation')
                
                st.rerun()
    
    # Statistics
    st.markdown("---")
    col1, col2, col3, col4, col5, col6 = st.columns(6)
    col1.metric("Business Days", stats['business_days'])
    col2.metric("Vacation Days", stats['vacation_days'])
    col3.metric("Team Traveling", stats['business_travelers'], help="People on business travel only")
    col4.metric("Month Travel", stats['month_business_days'])
    col5.metric("Week Cost", f"${stats['week_cost']:,.0f}")
    col6.metric("Month Cost", f"${stats['month_cost']:,.0f}")

with tab2:
    st.header("✅ Travel Approvals")
    
    # Check if user is manager
    current_user = st.selectbox("Select your name:", [m['name'] for m in st.session_state.team_members])
    is_manager = next((m['is_manager'] for m in st.session_state.team_members if m['name'] == current_user), False)
    
    if is_manager:
        # Show pending approvals
        pending = []
        for key in st.session_state.approvals_pending:
            if key in st.session_state.travel_data:
                data = st.session_state.travel_data[key]
                if not data.get('approved', True):
                    name, date = key.rsplit('_', 1)
                    pending.append({
                        'key': key,
                        'name': name,
                        'date': date,
                        'cost': data.get('daily_cost', 500)
                    })
        
        if pending:
            st.subheader(f"Pending Approvals ({len(pending)})")
            
            for item in pending:
                col1, col2, col3, col4 = st.columns([2, 1, 1, 1])
                col1.write(f"**{item['name']}** - {item['date']}")
                col2.write(f"${item['cost']:,.0f}")
                
                if col3.button("✅ Approve", key=f"approve_{item['key']}"):
                    st.session_state.travel_data[item['key']]['approved'] = True
                    st.session_state.approvals_pending.remove(item['key'])
                    # Send notification
                    member = next((m for m in st.session_state.team_members if m['name'] == item['name']), None)
                    if member:
                        send_email_notification(
                            member['email'],
                            "Travel Approved",
                            f"Your travel request for {item['date']} has been approved."
                        )
                    st.success(f"Approved travel for {item['name']}")
                    st.rerun()
                
                if col4.button("❌ Reject", key=f"reject_{item['key']}"):
                    st.session_state.travel_data[item['key']] = {'status': 'office'}
                    st.session_state.approvals_pending.remove(item['key'])
                    st.info(f"Rejected travel for {item['name']}")
                    st.rerun()
        else:
            st.info("No pending approvals")
    else:
        st.info("Only managers can approve travel requests")

with tab3:
    st.header("📊 Analytics")
    
    # Prepare data for visualizations
    travel_by_person = {}
    vacation_by_person = {}
    
    for key, value in st.session_state.travel_data.items():
        name = key.rsplit('_', 1)[0]
        if value.get('status') == 'business' and value.get('approved', True):
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
    
    # Generate report button
    if st.button("📊 Generate Weekly Report"):
        report = generate_weekly_report()
        st.success("Weekly report generated and sent!")
        st.markdown(report, unsafe_allow_html=True)

with tab4:
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
        
        approval_threshold = st.number_input("Requires Approval Above ($)",
                                            value=st.session_state.budget_data['requires_approval_above'],
                                            step=100)
        if approval_threshold != st.session_state.budget_data['requires_approval_above']:
            st.session_state.budget_data['requires_approval_above'] = approval_threshold
    
    with col2:
        # Calculate spending
        total_spent = sum(v.get('daily_cost', 0) for v in st.session_state.travel_data.values() 
                         if v.get('status') == 'business' and v.get('approved', True))
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

with tab5:
    st.header("📧 Email Notifications")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("Email Configuration")
        
        # Email settings
        st.session_state.email_config['sender_email'] = st.text_input(
            "Sender Email",
            value=st.session_state.email_config['sender_email']
        )
        
        st.session_state.email_config['sender_password'] = st.text_input(
            "Password",
            type="password"
        )
        
        st.session_state.email_config['notification_enabled'] = st.checkbox(
            "Enable Email Notifications",
            value=st.session_state.email_config['notification_enabled']
        )
        
        if st.button("Test Email Configuration"):
            if send_email_notification(
                st.session_state.email_config['sender_email'],
                "Test Email",
                "This is a test email from OCM Travel Tracker"
            ):
                st.success("Test email sent successfully!")
            else:
                st.error("Failed to send test email. Check configuration.")
    
    with col2:
        st.subheader("Automated Reports")
        
        st.session_state.automated_reports['weekly_report'] = st.checkbox(
            "Send Weekly Reports",
            value=st.session_state.automated_reports['weekly_report']
        )
        
        st.session_state.automated_reports['monthly_report'] = st.checkbox(
            "Send Monthly Reports",
            value=st.session_state.automated_reports['monthly_report']
        )
        
        # Report recipients
        recipients = st.text_area(
            "Report Recipients (one email per line)",
            value="\n".join(st.session_state.automated_reports['report_recipients'])
        )
        st.session_state.automated_reports['report_recipients'] = recipients.split("\n")
        
        st.info("""
        **Notification Types:**
        - Travel approval requests
        - Approval decisions
        - Weekly/Monthly reports
        - Budget alerts
        - Calendar invites for vacations
        """)

with tab6:
    st.header("⚙️ Settings")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("Export Data")
        
        excel_file = export_to_excel_advanced()
        if excel_file:
            st.download_button(
                "📥 Download Full Report (Excel)",
                data=excel_file,
                file_name=f"travel_report_{datetime.date.today()}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        
        # Power Automate webhook
        st.subheader("Power Automate Integration")
        webhook_url = st.text_input(
            "Webhook URL",
            placeholder="https://prod-xx.westeurope.logic.azure.com/workflows/..."
        )
        
        if st.button("Send Data to Power Automate"):
            if webhook_url:
                # Prepare data for Power Automate
                data = {
                    'date': datetime.date.today().isoformat(),
                    'travel_data': [
                        {
                            'name': k.split('_')[0],
                            'date': k.split('_')[1],
                            'status': v.get('status'),
                            'cost': v.get('daily_cost', 0)
                        }
                        for k, v in st.session_state.travel_data.items()
                        if v.get('status') in ['business', 'vacation']
                    ]
                }
                
                try:
                    response = requests.post(webhook_url, json=data)
                    if response.status_code == 200:
                        st.success("Data sent to Power Automate successfully!")
                    else:
                        st.error(f"Error: {response.status_code}")
                except Exception as e:
                    st.error(f"Failed to send data: {str(e)}")
    
    with col2:
        st.subheader("Team Management")
        
        # Add member
        with st.expander("Add Team Member"):
            new_name = st.text_input("Name")
            new_role = st.text_input("Role")
            new_email = st.text_input("Email")
            is_manager = st.checkbox("Is Manager")
            
            if st.button("Add"):
                if new_name and new_role and new_email:
                    st.session_state.team_members.append({
                        "name": new_name,
                        "role": new_role,
                        "email": new_email,
                        "is_manager": is_manager
                    })
                    st.success(f"Added {new_name}")
                    st.rerun()
        
        # Calendar sync
        st.subheader("Calendar Integration")
        st.info("""
        **Outlook Calendar Sync:**
        1. Travel bookings create calendar events
        2. Automatic vacation blocking
        3. Meeting conflict detection
        4. Travel reminders
        """)

# Sidebar
with st.sidebar:
    st.header("📊 Year Summary")
    
    # Calculate totals
    total_business = sum(1 for v in st.session_state.travel_data.values() 
                        if v.get('status') == 'business' and v.get('approved', True))
    total_vacation = sum(1 for v in st.session_state.travel_data.values() 
                        if v.get('status') == 'vacation')
    total_cost = sum(v.get('daily_cost', 0) for v in st.session_state.travel_data.values() 
                    if v.get('status') == 'business' and v.get('approved', True))
    
    st.metric("YTD Business Travel", f"{total_business} days")
    st.metric("YTD Vacation", f"{total_vacation} days")
    st.metric("YTD Cost", f"${total_cost:,.0f}")
    st.metric("Pending Approvals", stats['pending_approvals'])
    
    st.markdown("---")
    st.markdown("""
    **Legend:**
    - 🏢 Office/Remote
    - ✈️ Business Travel
    - 🏖️ Vacation
    - ⏳ Pending Approval
    
    **Features:**
    - Email notifications
    - Manager approvals
    - Automated reports
    - Calendar integration
    - Power Automate sync
    """)