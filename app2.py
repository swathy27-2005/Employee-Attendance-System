from flask import Flask, render_template, request, send_file, redirect, url_for,session,flash
import pandas as pd
import os
from datetime import datetime, timedelta
import json
import matplotlib.pyplot as plt
import io
import base64
import plotly.express as px
import plotly.io as pio
from collections import Counter

app = Flask(__name__)
ATTENDANCE_FILE = "Live_Attendance_Log.xlsx"

@app.route("/", methods=["GET", "POST"])
def home():
    if request.method == "POST":
        role = request.form.get("role")
        emp_id = request.form.get("emp_id").strip()
        emp_name = request.form.get("emp_name").strip()
        password = request.form.get("password").strip()
        shift = request.form.get("shift").strip()
        action = request.form.get("action")
        status = request.form.get("status")

        if action == "Login":
            return login(role, emp_id, emp_name, password, shift)
        elif action == "Logout":
            return logout(emp_id, emp_name, status)
    return render_template("index.html")


def login(role, emp_id, emp_name, password, shift):
    if not emp_id or not emp_name or not password or not shift:
        return "⚠️ Fill all fields to login."

    login_time = datetime.now()

    entry = {
        'Date': login_time.strftime('%Y-%m-%d'),
        'EMP NUM': emp_id,
        'EMP NAME': emp_name,
        'Shift': shift,
        'IN TIME': login_time.strftime('%H:%M:%S'),
        'OUT TIME': '',
        'Status': '',
        'Role': role
    }

    df = pd.read_excel(ATTENDANCE_FILE) if os.path.exists(ATTENDANCE_FILE) else pd.DataFrame()
    df = pd.concat([df, pd.DataFrame([entry])], ignore_index=True)
    df.to_excel(ATTENDANCE_FILE, index=False)

    return f"✅ {role} {emp_name} logged in at {entry['IN TIME']}."


def logout(emp_id, emp_name, status):
    if not os.path.exists(ATTENDANCE_FILE):
        return "❌ No login data found."

    df = pd.read_excel(ATTENDANCE_FILE)

    mask = (df['EMP NUM'] == emp_id) & (df['EMP NAME'] == emp_name) & (df['OUT TIME'].isna() | (df['OUT TIME'] == '') | (df['OUT TIME'] == 'nan'))

    if df[mask].empty:
        return "❌ No matching login entry found or already logged out."

    logout_time = datetime.now().strftime('%H:%M:%S')
    df.loc[mask, 'OUT TIME'] = logout_time
    df.loc[mask, 'Status'] = status if status else 'Logged out'

    df.to_excel(ATTENDANCE_FILE, index=False)
    return f"✅ {emp_name} logged out at {logout_time} with status {status}."

@app.route('/logout')
def logout_route():
    emp_id = session.get('emp_id')
    emp_name = session.get('emp_name')

    if not emp_id or not emp_name:
        return redirect(url_for('home'))  # Redirect to login if session is missing

    # Optional: Customize status or get from query param
    status = request.args.get('status', 'Logged out')
    message = logout(emp_id, emp_name, status)

    session.clear()  # Clear the session after logging out
    return redirect(url_for('home'))  # or render a goodbye page with message

@app.route("/report", methods=["GET", "POST"])
def report():
    if request.method == "POST":
        role = request.form.get("role")
        emp_id = request.form.get("emp_id")
        emp_name = request.form.get("emp_name")
        start_date = request.form.get("start_date")
        end_date = request.form.get("end_date")

        if not os.path.exists(ATTENDANCE_FILE):
            return "❌ No data available."

        df = pd.read_excel(ATTENDANCE_FILE)

        if role == "Employee":
            df = df[(df['EMP NUM'] == emp_id) & (df['EMP NAME'] == emp_name)]

        df['Date'] = pd.to_datetime(df['Date'])

        if start_date and end_date:
            df = df[(df['Date'] >= pd.to_datetime(start_date)) & (df['Date'] <= pd.to_datetime(end_date))]

        if df.empty:
            return "ℹ️ No records found."

        # Fill NaN values to avoid errors
        df.fillna('', inplace=True)

        # Extract necessary columns for display
        df_html = df[['Date', 'EMP NUM', 'EMP NAME', 'Shift', 'IN TIME', 'OUT TIME', 'Status']].to_html(index=False)

        # Calculate counts of each status
        status_counts = df['Status'].value_counts()

        # Ensure all status values (P, A, CL, WO, PL) are in the counts (if they exist)
        status_counts = status_counts.reindex(['P', 'A', 'CL', 'WO', 'PL'], fill_value=0)

        # Get counts for each status
        total_counts = {
            'Present (P)': status_counts['P'],
            'Absent (A)': status_counts['A'],
            'Casual Leave (CL)': status_counts['CL'],
            'W/O (WO)': status_counts['WO'],
            'Privilege Leave (PL)': status_counts['PL'],
        }
        fig, ax = plt.subplots(figsize=(8, 6))
        ax.bar(total_counts.keys(), total_counts.values(), color=['green', 'red', 'blue', 'orange', 'purple'])

        ax.set_xlabel('Status')
        ax.set_ylabel('Count')
        ax.set_title('Attendance Status Distribution')

        # Saving the plot to a BytesIO object and encoding it in base64 for embedding in HTML
        img = io.BytesIO()
        plt.savefig(img, format='png')
        img.seek(0)
        graph_url = base64.b64encode(img.getvalue()).decode('utf-8')
        img.close()

        total_days = len(df)
        present_days = status_counts['P']
        encouragement_msg = ""

        # Identify role: assuming role is stored in the Excel file
        user_role = df['Role'].iloc[0].lower()  # "employee" or "admin"
        show_popup = False

        if user_role == "employee" and total_days > 0:
            attendance_ratio = present_days / total_days
            show_popup = True

            if attendance_ratio >= 0.95:
                encouragement_msg = f"🎉 Outstanding, {emp_name}! You’ve been extremely punctual. Keep leading by example!"
            elif attendance_ratio >= 0.85:
                encouragement_msg = f"👏 Good work, {emp_name}! Your attendance is impressive."
            elif attendance_ratio >= 0.70:
                encouragement_msg = f"🙂 {emp_name}, you're doing well. Stay consistent and keep pushing!"
            else:
                encouragement_msg = f"🛠 {emp_name}, let’s aim for better attendance. Every effort counts!"


        # Return the table and graph
        return render_template("report.html",
                            table=df_html,
                            total_counts=total_counts,
                            graph_url=graph_url,encouragement_msg=encouragement_msg,
                            show_popup=show_popup)

    return render_template("report.html",
                           table=None,
                           total_counts=None,
                           graph_url=None,
                           encouragement_msg=None,
                           show_popup=False)


@app.route("/download", methods=["POST"])
def download():
    role = request.form.get("role")
    emp_id = request.form.get("emp_id")
    emp_name = request.form.get("emp_name")
    start_date = request.form.get("start_date")
    end_date = request.form.get("end_date")

    if not os.path.exists(ATTENDANCE_FILE):
        return "❌ No data available."

    df = pd.read_excel(ATTENDANCE_FILE)

    if role == "Employee":
        df = df[(df['EMP NUM'] == emp_id) & (df['EMP NAME'] == emp_name)]

    df['Date'] = pd.to_datetime(df['Date'])

    if start_date and end_date:
        df = df[(df['Date'] >= pd.to_datetime(start_date)) & (df['Date'] <= pd.to_datetime(end_date))]

    if df.empty:
        return "ℹ️ No records found."

    df['Date'] = df['Date'].dt.strftime('%Y-%m-%d')
    df.fillna('', inplace=True)
    df[['Date', 'EMP NUM', 'EMP NAME', 'Shift', 'IN TIME', 'OUT TIME', 'Status']].to_excel("Filtered_Report.xlsx", index=False)
    return send_file("Filtered_Report.xlsx", as_attachment=True)

@app.route("/dashboard")
def dashboard():
    if not os.path.exists(ATTENDANCE_FILE):
        return "❌ No attendance data found."

    df = pd.read_excel(ATTENDANCE_FILE)
    df.fillna('', inplace=True)

    # Parse datetime
    df['TIME IN'] = pd.to_datetime(df['IN TIME'], errors='coerce')
    df['DATE'] = pd.to_datetime(df['Date'], errors='coerce')

    # --- 🎖️ BADGES CALCULATION ---
    badge_threshold = pd.to_datetime("11:30:00").time()
    punctual_df = df[(df['TIME IN'].notna()) & (df['TIME IN'].dt.time <= badge_threshold)]
    badge_counts = punctual_df['EMP NAME'].value_counts()
    badges = {}

    for name, count in badge_counts.items():
        if count >= 20:
            badges[name] = "🏅 Gold"
        elif count >= 10:
            badges[name] = "🥈 Silver"
        elif count >= 1:
            badges[name] = "🥉 Bronze"

    # --- 📊 HEATMAP ---
    df['day'] = df['DATE'].dt.day_name()
    df['month'] = df['DATE'].dt.to_period('M').astype(str)
    heatmap_data = df.groupby(['month', 'day']).size().reset_index(name='count')

    heatmap_fig = px.density_heatmap(
        heatmap_data,
        x='day', y='month', z='count',
        color_continuous_scale='Blues',
        title='Attendance Activity Heatmap'
    )
    heatmap_html = heatmap_fig.to_html(full_html=False)

    # --- ⏱️ PUNCTUALITY SCORE ---
    df['is_late'] = df['TIME IN'].dt.time > badge_threshold
    punctuality = df.groupby('EMP NAME')['is_late'].apply(lambda x: 100 - (x.sum() / x.count()) * 100).reset_index()
    punctuality.columns = ['EMP NAME', 'Punctuality Score']

    punctual_fig = px.bar(
        punctuality.sort_values(by='Punctuality Score', ascending=False),
        x='EMP NAME', y='Punctuality Score',
        color='Punctuality Score', color_continuous_scale='Teal',
        title='Punctuality Score (%)'
    )
    punctual_html = punctual_fig.to_html(full_html=False)

    # --- 📅 MONTHLY ATTENDANCE PIE CHART ---
    df['month'] = df['DATE'].dt.strftime('%B %Y')
    pie_data = df.groupby(['month', 'Status']).size().reset_index(name='count')

    pie_fig = px.sunburst(
        pie_data,
        path=['month', 'Status'],
        values='count',
        color='Status',
        color_continuous_scale='Viridis',
        color_discrete_map={
            'P': 'green',
            'A': 'red',
            'CL': 'blue',
            'WO': 'orange',
            'PL': 'purple'
        },
        title='Monthly Attendance Overview'
    )
    pie_html = pie_fig.to_html(full_html=False)

    return render_template("dashboard.html",
                           badges=badges,
                           heatmap_html=heatmap_html,
                           punctual_html=punctual_html,
                           pie_html=pie_html)



@app.route("/mark", methods=["GET", "POST"])
def mark():
    if request.method == "POST":
        emp_id = request.form.get("mark_emp_id")
        emp_name = request.form.get("mark_emp_name")
        shift = request.form.get("mark_shift")
        date = request.form.get("mark_date")
        status = request.form.get("mark_status")

        if not emp_id or not emp_name or not shift or not date:
            return "⚠️ Please fill all fields."

        entry = {
            'Date': pd.to_datetime(date).strftime('%Y-%m-%d'),
            'EMP NUM': emp_id,
            'EMP NAME': emp_name,
            'Shift': shift,
            'IN TIME': '',
            'OUT TIME': '',
            'Status': status,
            'Role': 'Employee'
        }

        df = pd.read_excel(ATTENDANCE_FILE) if os.path.exists(ATTENDANCE_FILE) else pd.DataFrame()
        df = pd.concat([df, pd.DataFrame([entry])], ignore_index=True)
        df.to_excel(ATTENDANCE_FILE, index=False)

        return f"✅ Status {status} marked for {emp_name} on {entry['Date']}."

    return render_template("mark.html")

@app.route('/trend/<emp_name>')
def attendance_trend(emp_name):
    df = pd.read_excel('attendance_data.xlsx')
    df['Date'] = pd.to_datetime(df['Date'])
    emp_df = df[df['Employee Name'] == emp_name]

    emp_df['Week'] = emp_df['Date'].dt.strftime('%Y-W%U')
    weekly_summary = emp_df.groupby(['Week', 'Status']).size().unstack(fill_value=0).reset_index()

    for status in ['Present', 'Absent', 'Late']:
        if status not in weekly_summary.columns:
            weekly_summary[status] = 0

    chart_labels = weekly_summary['Week'].tolist()
    chart_datasets = {
        'Present': weekly_summary['Present'].tolist(),
        'Absent': weekly_summary['Absent'].tolist(),
        'Late': weekly_summary['Late'].tolist()
    }

    return render_template('attendance_trend.html',
                           emp_name=emp_name,
                           chart_labels=chart_labels,
                           chart_datasets=chart_datasets)


if __name__ == "__main__":
    app.run(debug=True)
