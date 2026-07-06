import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime, timedelta
import os
import time
import io
from sqlalchemy import create_engine, text

# --- الأمان والسحابة ---
DB_URL = st.secrets["DB_URL"]

def get_engine():
    return create_engine(DB_URL)

engine = get_engine()

# ضبط إعدادات الصفحة والتصميم الداعم لـ RTL بالكامل
st.set_page_config(page_title="منظومة إدارة الأداء الفني - شركة المكتب الرقمي", page_icon="🛠️", layout="wide")

st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Tajawal:wght@400;500;700&display=swap');
    
    html, body, [data-testid="stSidebar"], .stApp, p, div, span, label {
        font-family: 'Tajawal', sans-serif;
        direction: RTL !important;
        text-align: right !important;
    }
    
    [data-testid="stSidebarNav"] {
        direction: RTL !important;
        text-align: right !important;
    }
    
    .stHeadingContainer h1, .stHeadingContainer h2, .stHeadingContainer h3, .stHeadingContainer h4 {
        color: #1E3A8A;
        text-align: right !important;
    }
    
    th {
        background-color: #1E3A8A !important;
        color: white !important;
        text-align: right !important;
    }
    
    td {
        text-align: right !important;
    }
    
    .star-card {
        background-color: #FFFDF0;
        padding: 20px;
        border-radius: 10px;
        border: 2px solid #FBBF24;
        border-right: 10px solid #FBBF24;
        margin-bottom: 25px;
    }
    
    .developer-card {
        background-color: #f8fafc;
        padding: 15px;
        border-radius: 8px;
        border-right: 4px solid #1E3A8A;
        margin-top: 20px;
    }
    </style>
""", unsafe_allow_html=True)

# --- الدوال المساعدة ---
def calculate_sla_status(exp_date_str):
    if not exp_date_str: 
        return "غير محدد"
    try:
        exp_date = datetime.strptime(str(exp_date_str), "%Y-%m-%d").date()
        today = datetime.now().date()
        delta = (exp_date - today).days
        if delta > 0: 
            return f"🟢 ساري (متبقي {delta} يوم)"
        elif delta == 0: 
            return "🟡 ينتهي اليوم"
        else: 
            return f"🔴 منتهي (منذ {abs(delta)} يوم)"
    except: 
        return "صيغة غير صحيحة"

def to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='التقرير المصدّر')
    return output.getvalue()

def to_csv_printable(df):
    return df.to_csv(index=False).encode('utf-8-sig')

def get_star_technician():
    query = "SELECT t.*, tech.name as tech_name FROM tickets t JOIN technicians tech ON t.tech_id = tech.id"
    df_all = pd.read_sql_query(text(query), engine)
    if df_all.empty:
        return None, pd.DataFrame()
        
    df_all['time_reported'] = pd.to_datetime(df_all['time_reported'])
    df_all['time_resolved'] = pd.to_datetime(df_all['time_resolved'])
    df_all['time_arrived'] = pd.to_datetime(df_all['time_arrived'])
    df_all['repair_time_hrs'] = (df_all['time_resolved'] - df_all['time_arrived']).dt.total_seconds() / 3600.0
    
    current_month = datetime.now().strftime("%Y-%m")
    df_current = df_all[df_all['time_reported'].dt.strftime("%Y-%m") == current_month]
    if df_current.empty:
        return None, pd.DataFrame()
        
    tech_stats = df_current.groupby('tech_name').agg(
        total_visits=('id', 'count'),
        avg_repair_time=('repair_time_hrs', 'mean'),
        first_time_fixes=('first_time_fix', 'sum')
    ).reset_index()
    
    tech_stats['score'] = (tech_stats['total_visits'] * 3.0) + (tech_stats['first_time_fixes'] * 4.0) - (tech_stats['avg_repair_time'].fillna(24) * 1.5)
    if tech_stats.empty:
        return None, pd.DataFrame()
        
    winner = tech_stats.sort_values(by='score', ascending=False).iloc[0]['tech_name']
    
    df_all['year_month'] = df_all['time_reported'].dt.strftime("%Y-%m")
    historical_winners = []
    for month, group in df_all.groupby('year_month'):
        g_stats = group.groupby('tech_name').agg(total_v=('id', 'count'), avg_r=('repair_time_hrs', 'mean'), ftf=('first_time_fix', 'sum')).reset_index()
        g_stats['score'] = (g_stats['total_v'] * 3.0) + (g_stats['ftf'] * 4.0) - (g_stats['avg_r'].fillna(24) * 1.5)
        if not g_stats.empty:
            historical_winners.append(g_stats.sort_values(by='score', ascending=False).iloc[0]['tech_name'])
            
    history_df = pd.DataFrame(historical_winners, columns=['tech_name']).value_counts().reset_index(name='star_count')
    return winner, history_df

# --- نظام التحقق من الدخول ---
if 'logged_in' not in st.session_state: 
    st.session_state['logged_in'] = False

if not st.session_state['logged_in']:
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        if os.path.exists("MAC logo resized.png"): 
            st.image("MAC logo resized.png", use_container_width=True)
        st.markdown("<h2 style='text-align: center;'>تسجيل الدخول للمنظومة</h2>", unsafe_allow_html=True)
        with st.form("login_form"):
            username = st.text_input("اسم المستخدم")
            password = st.text_input("كلمة المرور", type="password")
            if st.form_submit_button("دخول"):
                if username == "Ahmed" and password == "admin123":
                    st.session_state['logged_in'] = True
                    st.rerun()
                else: 
                    st.error("❌ بيانات الدخول غير صحيحة.")
    st.stop()

# --- القائمة الجانبية ---
if os.path.exists("MAC logo resized.png"):
    st.sidebar.image("MAC logo resized.png", use_container_width=True)

st.sidebar.title("🛠️ المكتب الرقمي")
st.sidebar.subheader("نظام إدارة الأداء الفني v9.1")

menu = st.sidebar.radio("انتقل إلى القائمة:", [
    "📊 لوحة التحكم والأداء الشهري", 
    "🔍 البحث الشامل عن جهاز (S/N)",
    "➕ تسجيل بلاغ صيانة جديد", 
    "🖥️ إدارة البلاغات والتذاكر",
    "📅 الزيارات الدورية (PM)",
    "🏢 إدارة العملاء والأجهزة (Profile)",
    "👨‍💻 إدارة فريق الفنيين",
    "⚙️ إعدادات أنواع العقود"
])

if st.sidebar.button("🚪 تسجيل الخروج"):
    st.session_state['logged_in'] = False
    st.rerun()

st.sidebar.markdown("""
<div class="developer-card">
    <h4>⚙️ تصميم وتطوير المنظومة:</h4>
    <b>م. أحمد عثمان</b><br>
    مدير إدارة الدعم الفني وتقنية المعلومات<br>
    📧 AhmedE@almactab.com<br>
    📱 0923009907
</div>
""", unsafe_allow_html=True)

# --- 1. 📊 لوحة التحكم والأداء الشهري ---
if menu == "📊 لوحة التحكم والأداء الشهري":
    st.title("📊 الملخص الشهري وتحليل الأداء الفني")
    winner_name, star_history_df = get_star_technician()
    
    if winner_name:
        stars_count = 0
        if not star_history_df.empty and winner_name in star_history_df['tech_name'].values:
            stars_count = star_history_df.loc[star_history_df['tech_name'] == winner_name, 'star_count'].values[0]
        st.markdown(f"""
        <div class="star-card">
            <h3 style='margin:0; color:#B45309;'>🌟 الموظف المثالي لهذا الشهر الحالي: {winner_name} ⭐</h3>
            <p style='margin:5px 0 0 0; font-size:16px; color:#4B5563;'>
                تم اختيار المهندس <b>{winner_name}</b> تلقائياً بناءً على سرعة الاستجابة، عدد الإصلاحات، ومعدل الإصلاح من أول زيارة.
                <br>🎖️ <b>عدد مرات الحصول على اللقب تاريخياً:</b> {stars_count} مرة.
            </p>
        </div>
        """, unsafe_allow_html=True)
    else:
        st.info("💡 سيظهر الفني المثالي هنا تلقائياً بمجرد إدخال بلاغات صيانة وإغلاقها للشهر الحالي.")

    query_tickets = "SELECT t.*, c.name as client_name, e.brand, e.model, tech.name as tech_name FROM tickets t JOIN equipment e ON t.equipment_id = e.id JOIN clients c ON e.client_id = c.id JOIN technicians tech ON t.tech_id = tech.id ORDER BY t.id DESC"
    df_tickets = pd.read_sql_query(text(query_tickets), engine)
    
    if not df_tickets.empty:
        df_tickets['time_reported'] = pd.to_datetime(df_tickets['time_reported'])
        df_tickets['time_arrived'] = pd.to_datetime(df_tickets['time_arrived'])
        df_tickets['time_resolved'] = pd.to_datetime(df_tickets['time_resolved'])
        df_tickets['response_time_hrs'] = (df_tickets['time_arrived'] - df_tickets['time_reported']).dt.total_seconds() / 3600.0
        df_tickets['repair_time_hrs'] = (df_tickets['time_resolved'] - df_tickets['time_arrived']).dt.total_seconds() / 3600.0
        df_tickets['month_year'] = df_tickets['time_reported'].dt.to_period('M').astype(str)
        
        months_available = sorted(df_tickets['month_year'].dropna().unique(), reverse=True)
        selected_month = st.selectbox("📅 اختر الشهر المالي للمراقبة:", months_available)
        df_filtered = df_tickets[df_tickets['month_year'] == selected_month]
        
        m_col1, m_col2, m_col3, m_col4 = st.columns(4)
        m_col1.metric("📥 إجمالي البلاغات المسجلة", len(df_filtered))
        m_col2.metric("🔒 البلاغات المغلقة بنجاح", len(df_filtered[df_filtered['status'] == "مغلق"]))
        m_col3.metric("⏱️ متوسط الاستجابة (ساعة)", f"{df_filtered['response_time_hrs'].mean():.1f}" if not df_filtered['response_time_hrs'].isna().all() else "---")
        m_col4.metric("⚙️ متوسط الإصلاح (ساعة)", f"{df_filtered['repair_time_hrs'].mean():.1f}" if not df_filtered['repair_time_hrs'].isna().all() else "---")
        
        st.markdown("### 📋 سجل البلاغات التفصيلي للشهر المختار")
        display_df = df_filtered[['id', 'client_name', 'brand', 'model', 'tech_name', 'status']].copy()
        if winner_name:
            display_df['tech_name'] = display_df['tech_name'].apply(lambda x: f"{x} ⭐" if x == winner_name else x)
        display_df.columns = ['رقم البلاغ', 'اسم العميل', 'الماركة', 'الموديل', 'المهندس المختص', 'حالة البلاغ']
        
        styled_dash = display_df.style.set_properties(**{'text-align': 'right', 'direction': 'rtl'}).set_table_styles([{'selector': 'th', 'props': [('text-align', 'right'), ('direction', 'rtl')]}])
        st.dataframe(styled_dash, use_container_width=True)
        
        exp_col1, exp_col2, exp_col3 = st.columns([1, 1, 2])
        with exp_col1:
            st.download_button(label="📥 تصدير الجدول إلى Excel", data=to_excel(df_filtered), file_name=f"تقرير_الأداء_{selected_month}.xlsx", mime="application/vnd.ms-excel")
        with exp_col2:
            st.download_button(label="🖨️ تجهيز ملف للطباعة (CSV)", data=to_csv_printable(df_filtered), file_name=f"طباعة_تقرير_{selected_month}.csv", mime="text/csv")
            
        st.markdown("---")
        st.markdown("### 📊 الرسوم البيانية المقارنة وتحليل الأداء")
        g_col1, g_col2 = st.columns(2)
        with g_col1:
            st.subheader("🏆 السجل التاريخي لعدد نجوم التميز")
            if not star_history_df.empty:
                fig_stars = px.bar(star_history_df, x='tech_name', y='star_count', labels={'tech_name': 'اسم الفني', 'star_count': 'مرات التميز'}, title="عدد مرات الحصول على لقب الموظف المثالي تاريخياً", color='star_count', color_continuous_scale=px.colors.sequential.YlOrBr)
                st.plotly_chart(fig_stars, use_container_width=True)
            else:
                st.info("لا توجد بيانات كافية لرسم بياني النجوم التاريخي بعد.")
        with g_col2:
            st.subheader("📈 مستوى الأداء الإجمالي للفنيين (الشهر الحالي)")
            current_stats = df_filtered.groupby('tech_name').agg(total_v=('id', 'count'), ftf=('first_time_fix', 'sum')).reset_index()
            current_stats['مستوى الأداء التقديري'] = (current_stats['total_v'] * 10) + (current_stats['ftf'] * 15)
            if not current_stats.empty:
                fig_perf = px.line(current_stats, x='tech_name', y='مستوى الأداء التقديري', markers=True, title="منحنى الأداء التنافسي للفريق خلال الشهر")
                st.plotly_chart(fig_perf, use_container_width=True)
            else:
                st.info("لا توجد بلاغات كافية لحساب منحنى الأداء.")
    else:
        st.info("📥 لا توجد أي بلاغات صيانة مسجلة في قاعدة البيانات السحابية حتى الآن.")

# --- 2. 🔍 البحث الشامل عن جهاز (S/N) ---
elif menu == "🔍 البحث الشامل عن جهاز (S/N)":
    st.title("🔍 البحث الشامل (بروفايل الجهاز الذكي)")
    search_sn = st.text_input("أدخل الرقم التسلسلي (S/N) للبحث:", placeholder="مثال: XER8145001").strip()
    
    sla_df_options = pd.read_sql_query(text("SELECT name FROM sla_types ORDER BY id"), engine)
    sla_list_db = sla_df_options['name'].tolist() if not sla_df_options.empty else ["عقد افتراضي"]
    extended_sla_list = sla_list_db + ["+ إضافة نوع عقد جديد يدوياً..."]

    if search_sn:
        query_eq = "SELECT e.*, c.name as client_name FROM equipment e JOIN clients c ON e.client_id = c.id WHERE e.serial_number = :sn"
        eq_df = pd.read_sql_query(text(query_eq), engine, params={"sn": search_sn})
        
        if not eq_df.empty:
            eq = eq_df.iloc[0]
            eq_id = eq['id']
            query_tk = "SELECT t.*, tech.name as tech_name FROM tickets t JOIN technicians tech ON t.tech_id = tech.id WHERE t.equipment_id = :eq_id ORDER BY t.time_reported DESC"
            tickets_df = pd.read_sql_query(text(query_tk), engine, params={"eq_id": int(eq_id)})
            
            status_text = "🟢 يعمل (لا توجد بلاغات مفتوحة)" if tickets_df[tickets_df['status'] != 'مغلق'].empty else "🔴 معطل / قيد الصيانة"
            st.markdown(f"""
            <div style="background-color:#F8FAFC; padding:20px; border-radius:10px; border-right:8px solid #1E3A8A; margin-bottom:20px;">
                <h3>🖥️ بطاقة تعريف الجهاز: {eq['brand']} {eq['model']}</h3>
                <b>🏢 العميل:</b> {eq['client_name']} | <b>🏷️ السيريال:</b> {eq['serial_number']} | <b>⚙️ الحالة:</b> {status_text}<br>
                <b>📜 العقد:</b> {eq['sla_type']} | <b>⏳ الصلاحية:</b> {calculate_sla_status(eq['sla_expiration_date'])}
            </div>
            """, unsafe_allow_html=True)
            
            if not tickets_df.empty:
                st.subheader("🛠️ السجل التاريخي لصيانة الجهاز")
                display_hist = tickets_df[['id', 'tech_name', 'issue_description', 'time_reported', 'status', 'parts_replaced']].copy()
                display_hist.columns = ['رقم البلاغ', 'اسم المهندس', 'العطل', 'تاريخ البلاغ', 'الحالة', 'القطع المستبدلة']
                styled_hist = display_hist.style.set_properties(**{'text-align': 'right', 'direction': 'rtl'}).set_table_styles([{'selector': 'th', 'props': [('text-align', 'right'), ('direction', 'rtl')]}])
                st.dataframe(styled_hist, use_container_width=True)
        else:
            st.error("⚠️ لم يتم العثور على هذا الرقم التسلسلي.")

# --- 3. ➕ تسجيل بلاغ صيانة جديد ---
elif menu == "➕ تسجيل بلاغ صيانة جديد":
    st.title("➕ تسجيل بلاغ صيانة جديد")
    eq_rows = pd.read_sql_query(text("SELECT e.id, c.name as client_name, e.brand, e.model, e.serial_number FROM equipment e JOIN clients c ON e.client_id = c.id"), engine)
    tech_rows = pd.read_sql_query(text("SELECT id, name FROM technicians WHERE status='متاح'"), engine)
    
    if eq_rows.empty or tech_rows.empty:
        st.error("⚠️ تأكد من توافر أجهزة وفنيين بحالة متاح.")
    else:
        equip_options = {f"{r['client_name']} - {r['brand']} {r['model']} ({r['serial_number']})": r['id'] for _, r in eq_rows.iterrows()}
        tech_options = {r['name']: r['id'] for _, r in tech_rows.iterrows()}
        with st.form("add_ticket_form"):
            selected_equip = st.selectbox("الجهاز المتضرر:", list(equip_options.keys()))
            selected_tech = st.selectbox("المهندس المختص:", list(tech_options.keys()))
            issue_desc = st.text_area("وصف العطل:")
            if st.form_submit_button("حفظ وفتح البلاغ"):
                full_rep = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                with engine.begin() as conn:
                    conn.execute(text("INSERT INTO tickets (equipment_id, tech_id, issue_description, time_reported, status, first_time_fix) VALUES (:eq_id, :tech_id, :desc, :time_rep, 'مفتوح / قيد المتابعة', 0)"),
                                 {"eq_id": int(equip_options[selected_equip]), "tech_id": int(tech_options[selected_tech]), "desc": issue_desc, "time_rep": full_rep})
                st.success("🎉 تم تسجيل بلاغ الصيانة بنجاح!")

# --- 4. 🖥️ إدارة البلاغات والتذاكر ---
elif menu == "🖥️ إدارة البلاغات والتذاكر":
    st.title("🖥️ الإدارة الشاملة وتعديل تذاكر الصيانة")
    query_all_tk = "SELECT t.*, c.name as client_name, e.brand, e.model, tech.name as tech_name FROM tickets t JOIN equipment e ON t.equipment_id = e.id JOIN clients c ON e.client_id = c.id JOIN technicians tech ON t.tech_id = tech.id ORDER BY t.id DESC"
    df_all_tickets = pd.read_sql_query(text(query_all_tk), engine)
    
    if df_all_tickets.empty:
        st.info("لا توجد بلاغات مسجلة.")
    else:
        ticket_list = {f"بلاغ رقم {r['id']} - لعميل: {r['client_name']} ({r['status']})": r['id'] for _, r in df_all_tickets.iterrows()}
        selected_ticket_label = st.selectbox("🔍 اختر بلاغ الصيانة لتحديثه:", list(ticket_list.keys()))
        t_id = ticket_list[selected_ticket_label]
        current_ticket = df_all_tickets[df_all_tickets['id'] == t_id].iloc[0]
        
        with st.form("edit_ticket_form"):
            st.info(f"📝 وصف العطل: {current_ticket['issue_description']}")
            new_status = st.selectbox("تحديث الحالة:", ["مفتوح / قيد المتابعة", "قيد الانتظار لقطع الغيار", "مغلق"])
            parts_replaced = st.text_input("قطع الغيار المستبدلة:", value=current_ticket['parts_replaced'] or "")
            ftfr_new = st.checkbox("⚙️ تم الإصلاح من الزيارة الأولى (First Time Fix)", value=bool(current_ticket['first_time_fix']))
            if st.form_submit_button("حفظ التحديثات"):
                res_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S") if new_status == "مغلق" else None
                with engine.begin() as conn:
                    conn.execute(text("UPDATE tickets SET status = :status, time_resolved = :time_res, first_time_fix = :ftf, parts_replaced = :parts WHERE id = :id"),
                                 {"status": new_status, "time_res": res_time, "ftf": 1 if ftfr_new else 0, "parts": parts_replaced, "id": int(t_id)})
                st.success("✅ تم التحديث بنجاح!")
                st.rerun()

# --- 5. 📅 الزيارات الدورية (PM) ---
elif menu == "📅 الزيارات الدورية (PM)":
    st.title("📅 لوحة تحكم الزيارات الوقائية الدورية (PM)")
    query_pm = "SELECT p.id as pm_id, c.name as client_name, e.brand, e.model, e.serial_number, p.scheduled_date, p.status FROM pm_visits p JOIN equipment e ON p.equipment_id = e.id JOIN clients c ON e.client_id = c.id ORDER BY p.scheduled_date ASC"
    pm_df = pd.read_sql_query(text(query_pm), engine)
    
    if not pm_df.empty:
        pm_df.columns = ['رقم الزيارة', 'العميل', 'الماركة', 'الموديل', 'الرقم التسلسلي', 'تاريخ الزيارة', 'الحالة']
        styled_pm = pm_df.style.set_properties(**{'text-align': 'right', 'direction': 'rtl'}).set_table_styles([{'selector': 'th', 'props': [('text-align', 'right'), ('direction', 'rtl')]}])
        st.dataframe(styled_pm, use_container_width=True)
    else:
        st.info("لا توجد زيارات دورية مسجلة.")

# --- 6. 🏢 إدارة العملاء والأجهزة (Profile) ---
elif menu == "🏢 إدارة العملاء والأجهزة (Profile)":
    st.title("🗂️ ملفات العملاء والأجهزة الذكية")
    tab1, tab2, tab3 = st.tabs(["🗂️ بروفايل العميل وأجهزته", "➕ عميل جديد", "📋 قائمة ومستودع كل الأجهزة"])
    
    with tab1:
        clients_df = pd.read_sql_query(text("SELECT * FROM clients ORDER BY id"), engine)
        if not clients_df.empty:
            client_dict = {row['name']: row['id'] for _, row in clients_df.iterrows()}
            selected_client_id = client_dict[st.selectbox("🔍 ابحث عن عميل بالمنظومة لاستعراض ملفه:", list(client_dict.keys()))]
            c_info = clients_df[clients_df['id'] == selected_client_id].iloc[0]
            
            st.markdown("---")
            equip_df = pd.read_sql_query(text("SELECT * FROM equipment WHERE client_id = :cid"), engine, params={"cid": int(selected_client_id)})
            
            if not equip_df.empty:
                st.markdown(f"### 🖨️ الأجهزة والمعدات التابعة لـ {c_info['name']}")
                equip_df['حالة العقد'] = equip_df['sla_expiration_date'].apply(calculate_sla_status)
                
                display_equip_df = equip_df[['brand', 'model', 'serial_number', 'sla_type', 'pm_visits_count', 'حالة العقد']].copy()
                display_equip_df.columns = ['الماركة', 'الموديل', 'الرقم التسلسلي (S/N)', 'نوع العقد (SLA)', 'الزيارات الدورية السنوية', 'حالة العقد']
                
                styled_df = display_equip_df.style.set_properties(**{'text-align': 'right', 'direction': 'rtl'}).set_table_styles([{'selector': 'th', 'props': [('text-align', 'right'), ('direction', 'rtl')]}])
                st.dataframe(styled_df, use_container_width=True)
                st.download_button(f"📥 تصدير قائمة أجهزة {c_info['name']} لـ Excel", to_excel(equip_df), f"أجهزة_{c_info['name']}.xlsx", "application/vnd.ms-excel")
            else:
                st.info("لا توجد أجهزة مسجلة ل هذا العميل.")
        else:
            st.warning("المستودع فارغ من العملاء.")

    with tab2:
        with st.form("add_new_client"):
            name = st.text_input("اسم العميل / المؤسسة بالكامل *:")
            addr = st.text_input("العنوان:")
            if st.form_submit_button("إدراج وحفظ ملف العميل"):
                if name:
                    try:
                        with engine.begin() as conn:
                            conn.execute(text("INSERT INTO clients (name, address) VALUES (:name, :addr)"), {"name": name, "addr": addr})
                        st.success("🎉 تم الإضافة بنجاح!")
                        st.rerun()
                    except:
                        st.error("❌ الاسم مكرر.")

    with tab3:
        all_eq = pd.read_sql_query(text("SELECT c.name as client_name, e.* FROM equipment e JOIN clients c ON e.client_id = c.id ORDER BY e.id DESC"), engine)
        if not all_eq.empty:
            all_eq['حالة العقد'] = all_eq['sla_expiration_date'].apply(calculate_sla_status)
            display_all_eq = all_eq[['client_name', 'brand', 'model', 'serial_number', 'sla_type', 'pm_visits_count', 'حالة العقد']].copy()
            display_all_eq.columns = ['العميل المالك', 'الماركة', 'الموديل', 'الرقم التسلسلي', 'نوع العقد', 'الزيارات السنوية', 'حالة العقد']
            styled_all_df = display_all_eq.style.set_properties(**{'text-align': 'right', 'direction': 'rtl'}).set_table_styles([{'selector': 'th', 'props': [('text-align', 'right'), ('direction', 'rtl')]}])
            st.dataframe(styled_all_df, use_container_width=True)
        else:
            st.info("لم يتم إدراج معدات حتى الآن.")

# --- 7. 👨‍💻 إدارة فريق الفنيين ---
elif menu == "👨‍💻 إدارة فريق الفنيين":
    st.title("👨‍💻 سجل كادر الفنيين والمهندسين والتحكم بالحالة الميدانية")
    tech_tab1, tech_tab2, tech_tab3 = st.tabs(["📋 قائمة الفنيين ومتابعة الحالة", "➕ إضافة فني جديد للفريق", "✏️ تعديل بيانات مهندس"])
    df_techs = pd.read_sql_query(text("SELECT * FROM technicians ORDER BY id"), engine)

    with tech_tab1:
        if not df_techs.empty:
            display_df = df_techs[['id', 'name', 'specialty', 'phone', 'email', 'status']].copy()
            winner_name, _ = get_star_technician()
            if winner_name:
                display_df['name'] = display_df['name'].apply(lambda x: f"{x} ⭐" if x == winner_name else x)
            display_df.columns = ['رقم القيد', 'الاسم بالكامل', 'التخصص التقني', 'رقم الهاتف', 'البريد الإلكتروني', 'الحالة']
            styled_tech_df = display_df.style.set_properties(**{'text-align': 'right', 'direction': 'rtl'}).set_table_styles([{'selector': 'th', 'props': [('text-align', 'right'), ('direction', 'rtl')]}])
            st.dataframe(styled_tech_df, use_container_width=True)
        else:
            st.info("لا توجد أسماء مقيدة بسجل الفنيين.")
            
    with tech_tab2:
        with st.form("add_tech_form"):
            t_name = st.text_input("اسم المهندس / الفني الثلاثي *:")
            t_spec = st.text_input("التخصص الفني الدقيق:")
            if st.form_submit_button("اعتماد وإدراج الفني"):
                if t_name:
                    with engine.begin() as conn:
                        conn.execute(text("INSERT INTO technicians (name, specialty, status) VALUES (:name, :spec, 'متاح')"), {"name": t_name, "spec": t_spec})
                    st.success("🎉 تم إضافة المهندس الفني بنجاح!")
                    st.rerun()

    with tech_tab3:
        st.info("اختر التبويبات الأخرى لمتابعة العمل الفني.")

# --- 8. ⚙️ إعدادات أنواع العقود ---
elif menu == "⚙️ إعدادات أنواع العقود":
    st.title("⚙️ إدارة مسميات وأنواع العقود والـ SLAs المركزية")
    sla_df = pd.read_sql_query(text("SELECT id, name as \"نوع العقد المعتمد\" FROM sla_types ORDER BY id"), engine)
    
    col_s1, col_s2 = st.columns([1, 2])
    with col_s1:
        with st.form("add_sla"):
            new_sla_name = st.text_input("اكتب مسمى العقد الجديد:")
            if st.form_submit_button("حفظ وتأكيد الإضافة"):
                if new_sla_name:
                    try:
                        with engine.begin() as conn:
                            conn.execute(text("INSERT INTO sla_types (name) VALUES (:name)"), {"name": new_sla_name})
                        st.success("تمت الإضافة بنجاح!")
                        st.rerun()
                    except:
                        st.error("❌ المسمى مسجل مسبقاً.")
    with col_s2:
        styled_sla_df = sla_df.style.set_properties(**{'text-align': 'right', 'direction': 'rtl'}).set_table_styles([{'selector': 'th', 'props': [('text-align', 'right'), ('direction', 'rtl')]}])
        st.dataframe(styled_sla_df, use_container_width=True)