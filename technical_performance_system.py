import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime, timedelta
import os
import time
import io
from sqlalchemy import create_engine, text

# --- الاتصال السحابي الآمن عبر التجميع ---
DB_URL = st.secrets["DB_URL"]

def get_engine():
    return create_engine(DB_URL)

engine = get_engine()

# تأسيس وتأمين جداول المستخدمين والعقود المتقدمة
with engine.begin() as conn:
    conn.execute(text("""
        CREATE TABLE IF NOT EXISTS system_users (
            id SERIAL PRIMARY KEY,
            username TEXT UNIQUE,
            password TEXT,
            full_name TEXT,
            role TEXT
        );
    """))
    # إضافة المستخدم الافتراضي الإداري إذا لم يكن موجوداً
    conn.execute(text("""
        INSERT INTO system_users (username, password, full_name, role)
        VALUES ('Ahmed', 'admin123', 'م. أحمد عثمان', 'Admin')
        ON CONFLICT (username) DO NOTHING;
    """))
    conn.execute(text("ALTER TABLE technicians ADD COLUMN IF NOT EXISTS city TEXT;"))
    conn.execute(text("ALTER TABLE technicians ADD COLUMN IF NOT EXISTS image_path TEXT;"))
    conn.execute(text("ALTER TABLE equipment ADD COLUMN IF NOT EXISTS contract_coverage TEXT;"))
    conn.execute(text("ALTER TABLE equipment ADD COLUMN IF NOT EXISTS contract_value TEXT;"))
    conn.execute(text("ALTER TABLE equipment ADD COLUMN IF NOT EXISTS contract_duration TEXT;"))

st.set_page_config(page_title="منظومة إدارة الأداء الفني - شركة المكتب الرقمي", page_icon="🛠️", layout="wide")

st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Tajawal:wght@400;500;700&display=swap');
    
    html, body, [data-testid="stSidebar"], .stApp, p, div, span, label, input, select, textarea {
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
    
    .star-card {
        background-color: #FFFDF0;
        padding: 20px;
        border-radius: 12px;
        border: 2px solid #FBBF24;
        border-right: 12px solid #FBBF24;
        margin-bottom: 25px;
    }
    
    .worst-card {
        background-color: #FEF2F2;
        padding: 20px;
        border-radius: 12px;
        border: 2px solid #EF4444;
        border-right: 12px solid #EF4444;
        margin-bottom: 25px;
    }
    
    .kpi-box {
        background-color: #F8FAFC;
        padding: 15px;
        border-radius: 8px;
        border: 1px solid #E2E8F0;
        text-align: center !important;
    }
    
    .developer-card {
        background-color: #f1f5f9;
        padding: 15px;
        border-radius: 8px;
        border-right: 5px solid #1E3A8A;
        margin-top: 25px;
    }
    </style>
""", unsafe_allow_html=True)

if not os.path.exists("uploads/tech_images"):
    os.makedirs("uploads/tech_images")

# --- إدارة حالات الجلسة المحورية ---
if 'view_tech_id' not in st.session_state: st.session_state['view_tech_id'] = None
if 'tech_active_tab' not in st.session_state: st.session_state['tech_active_tab'] = 0
if 'logged_in' not in st.session_state: st.session_state['logged_in'] = False
if 'user_fullname' not in st.session_state: st.session_state['user_fullname'] = ""

# --- الدوال المساعدة ---
def calculate_sla_status(exp_date_str):
    if not exp_date_str: return "غير محدد"
    try:
        exp_date = pd.to_datetime(exp_date_str).date()
        delta = (exp_date - datetime.now().date()).days
        if delta > 30: return f"🟢 ساري المفعول (متبقي {delta} يوم)"
        elif 0 < delta <= 30: return f"🟡 شارف على الانتهاء (متبقي {delta} يوم)"
        else: return f"🔴 منتهي الصلاحية (منذ {abs(delta)} يوم)"
    except: return "صيغة غير صحيحة"

def to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='التقرير')
    return output.getvalue()

def to_csv_printable(df):
    return df.to_csv(index=False).encode('utf-8-sig')

def get_performance_extremes():
    df_tk = pd.read_sql_query(text("SELECT * FROM tickets"), engine)
    df_th = pd.read_sql_query(text("SELECT * FROM technicians"), engine)
    if df_tk.empty or df_th.empty:
        if not df_th.empty: return {"name": df_th.iloc[0]['name'], "image": df_th.iloc[0].get('image_path', ''), "is_estimated": True}, None, pd.DataFrame()
        return None, None, pd.DataFrame()
    df_all = df_tk.merge(df_th, left_on='tech_id', right_on='id', suffixes=('_tk', '_tech'))
    df_all['time_reported'] = pd.to_datetime(df_all['time_reported'])
    df_filtered = df_all[df_all['time_reported'].dt.strftime("%Y-%m") == datetime.now().strftime("%Y-%m")]
    if df_filtered.empty: return {"name": df_th.iloc[0]['name'], "image": df_th.iloc[0].get('image_path', ''), "is_estimated": True}, None, pd.DataFrame()
    stats = df_filtered.groupby('name').agg(total_visits=('id_tk', 'count'), first_time_fixes=('first_time_fix', 'sum'), pending_count=('status_tk', lambda x: x.str.contains('انتظار').sum())).reset_index().merge(df_th[['name', 'image_path']], on='name', how='left')
    stats['score'] = (stats['total_visits'] * 5) + (stats['first_time_fixes'] * 10) - (stats['pending_count'] * 4)
    sorted_stats = stats.sort_values(by='score', ascending=False)
    winner = {"name": sorted_stats.iloc[0]['name'], "image": sorted_stats.iloc[0].get('image_path', ''), "visits": sorted_stats.iloc[0]['total_visits'], "ftf": sorted_stats.iloc[0]['first_time_fixes']}
    worst = None
    if len(sorted_stats) > 1: worst = {"name": sorted_stats.iloc[-1]['name'], "image": sorted_stats.iloc[-1].get('image_path', ''), "visits": sorted_stats.iloc[-1]['total_visits'], "pending": sorted_stats.iloc[-1]['pending_count']}
    return winner, worst, sorted_stats

# --- نافذة الدخول السحابية المحمية ---
if not st.session_state['logged_in']:
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        if os.path.exists("MAC logo resized.png"): st.image("MAC logo resized.png", use_container_width=True)
        st.markdown("<h2 style='text-align: center;'>تسجيل الدخول للمنظومة الموحدة</h2>", unsafe_allow_html=True)
        with st.form("login_form"):
            user_input = st.text_input("اسم المستخدم")
            pass_input = st.text_input("كلمة المرور", type="password")
            if st.form_submit_button("دخول آمن للمنظومة"):
                chk_user = pd.read_sql_query(text("SELECT * FROM system_users WHERE username = :u AND password = :p"), engine, params={"u": user_input, "p": pass_input})
                if not chk_user.empty:
                    st.session_state['logged_in'] = True
                    st.session_state['user_fullname'] = chk_user.iloc[0]['full_name']
                    st.rerun()
                else: st.error("❌ بيانات المستخدم أو كلمة السر غير صحيحة.")
    st.stop()

# --- القائمة الجانبية وإعادة بناء الهوية المفقودة وميزات الخروج ---
if os.path.exists("MAC logo resized.png"): st.sidebar.image("MAC logo resized.png", use_container_width=True)
st.sidebar.title("🛠️ المكتب الرقمي")
st.sidebar.markdown(f"مرحباً بك: **{st.session_state['user_fullname']}**")

menu = st.sidebar.radio("القائمة الميدانية الرئيسية:", [
    "📊 لوحة التحكم والأداء الشهري", 
    "🔍 البحث الشامل عن جهاز (S/N)",
    "➕ تسجيل بلاغ صيانة جديد", 
    "🖥️ إدارة البلاغات والتذاكر",
    "📅 الزيارات الدورية (PM)",
    "🏢 إدارة العملاء والأجهزة (Profile)",
    "👨‍💻 إدارة فريق الفنيين",
    "⚙️ إعدادات أنواع العقود",
    "🔐 إدارة مستخدمي المنظومة"
])

if st.sidebar.button("🚪 تسجيل الخروج من المنظومة"):
    st.session_state['logged_in'] = False
    st.session_state['user_fullname'] = ""
    st.rerun()

st.sidebar.markdown("""
<div class="developer-card">
    <h4>⚙️ تصميم وإعداد المنظومة:</h4>
    <b>م. أحمد عثمان</b><br>
    مدير إدارة الدعم الفني وتقنية المعلومات<br>
    📧 AhmedE@almactab.com<br>
    📱 0923009907
</div>
""", unsafe_allow_html=True)

# --- 1. لوحة التحكم والأداء الشهري ---
if menu == "📊 لوحة التحكم والأداء الشهري":
    st.title("📊 الملخص الشهري وتحليل الأداء الفني")
    winner, worst, performance_df = get_performance_extremes()
    col_w1, col_w2 = st.columns(2)
    with col_w1:
        if winner:
            st.markdown('<div class="star-card">', unsafe_allow_html=True)
            img_c, txt_c = st.columns([1, 2])
            with img_c:
                if winner.get('image') and os.path.exists(str(winner['image'])): st.image(str(winner['image']), width=110)
                else: st.markdown("<h1 style='font-size:55px; margin:0;'>🌟</h1>", unsafe_allow_html=True)
            with txt_c:
                st.markdown(f"### 🌟 المهندس المتميز: {winner['name']} ⭐")
                if "is_estimated" not in winner: st.write(f"📊 زيارات: {winner['visits']} | إصلاح فوري: {winner['ftf']}")
            st.markdown('</div>', unsafe_allow_html=True)
    with col_w2:
        if worst:
            st.markdown('<div class="worst-card">', unsafe_allow_html=True)
            img_c2, txt_c2 = st.columns([1, 2])
            with img_c2:
                if worst.get('image') and os.path.exists(str(worst['image'])): st.image(str(worst['image']), width=110)
                else: st.markdown("<h1 style='font-size:55px; margin:0;'>⚠️</h1>", unsafe_allow_html=True)
            with txt_c2:
                st.markdown(f"### ⚠️ الأقل أداءً: {worst['name']} 👎")
                st.write(f"📊 زيارات: {worst['visits']} | معلقة لقطع الغيار: {worst['pending']}")
            st.markdown('</div>', unsafe_allow_html=True)

    if not performance_df.empty:
        fig = px.bar(performance_df, x='name', y='score', title="منحنى الكفاءة الإجمالي للفريق الحالي", color='score')
        st.plotly_chart(fig, use_container_width=True)

# --- 2. البحث الشامل ---
elif menu == "🔍 البحث الشامل عن جهاز (S/N)":
    st.title("🔍 بروفايل وبطاقة تعريف الآلة الذكية")
    sn_input = st.text_input("أدخل الرقم التسلسلي (S/N) للآلة:").strip()
    if sn_input:
        eq_df = pd.read_sql_query(text("SELECT e.*, c.name as client_name FROM equipment e JOIN clients c ON e.client_id = c.id WHERE e.serial_number = :sn"), engine, params={"sn": sn_input})
        if not eq_df.empty:
            eq = eq_df.iloc[0]
            st.markdown(f"""
            <div style="background-color:#F8FAFC; padding:20px; border-radius:10px; border-right:8px solid #1E3A8A; line-height:1.8;">
                <h3>🖥️ الطراز: {eq['brand']} {eq['model']} (S/N: {eq['serial_number']})</h3>
                <b>🏢 العميل:</b> {eq['client_name']} | <b>📜 نوع العقد:</b> {eq['sla_type']}<br>
                <b>⏳ صلاحية العقد الفعلي:</b> {calculate_sla_status(eq['sla_expiration_date'])}<br>
                <b>📦 ما يشمله العقد بالتفصيل:</b> {eq.get('contract_coverage', 'غير مححدد')}<br>
                <b>💰 القيمة والمدة السنوية:</b> {eq.get('contract_value', '---')} د.ل / {eq.get('contract_duration', '---')}
            </div>
            """, unsafe_allow_html=True)
        else: st.error("⚠️ لم يتم العثور على الرقم التسلسلي.")

# --- 3. تسجيل بلاغ صيانة جديد ---
elif menu == "➕ تسجيل بلاغ صيانة جديد":
    st.title("➕ فتح وتوثيق بلاغ صيانة جديد")
    eq_rows = pd.read_sql_query(text("SELECT e.id, c.name as client_name, e.brand, e.model, e.serial_number FROM equipment e JOIN clients c ON e.client_id = c.id"), engine)
    tech_rows = pd.read_sql_query(text("SELECT id, name FROM technicians"), engine)
    if not eq_rows.empty and not tech_rows.empty:
        equip_options = {f"{r['client_name']} - {r['brand']} {r['model']} ({r['serial_number']})": r['id'] for _, r in eq_rows.iterrows()}
        selected_label = st.selectbox("اختر الآلة المشكو منها:", list(equip_options.keys()))
        eq_id = equip_options[selected_label]
        target = pd.read_sql_query(text("SELECT * FROM equipment WHERE id = :id"), engine, params={"id": int(eq_id)}).iloc[0]
        st.warning(f"🔍 التحقق التلقائي للعقد: {target['sla_type']} ({calculate_sla_status(target['sla_expiration_date'])})")
        with st.form("add_ticket_form"):
            col1, col2 = st.columns(2)
            with col1:
                m_type = st.selectbox("نوع الآلة الآلي:", ["طابعة إنتاج ليزر", "طابعة عريضة ProStream", "منظومة أرشفة"])
                m_model = st.text_input("تأكيد الموديل:", value=f"{target['brand']} {target['model']}")
                i_type = st.text_input("نوع ونطاق العطل الحالي:")
            with col2:
                priority = st.selectbox("مستوى حساسية الاستجابة والطلب:", ["استجابة سريعة فورية", "استجابة عادية"])
                t_name = st.selectbox("المهندس المسؤول عن التنفيذ *:", [r['name'] for _, r in tech_rows.iterrows()])
                desc = st.text_area("وصف مفصل للمشتكى التقني العابر:")
            if st.form_submit_button("إصدار وتثبيت التذكرة"):
                t_id = tech_rows[tech_rows['name'] == t_name].iloc[0]['id']
                with engine.begin() as conn:
                    conn.execute(text("INSERT INTO tickets (equipment_id, tech_id, issue_description, time_reported, status, first_time_fix, parts_replaced) VALUES (:eid, :tid, :desc, :time, 'مفتوح / قيد المتابعة', 0, '')"),
                                 {"eid": int(eq_id), "tid": int(t_id), "desc": f"[{priority}] {m_type} - {i_type}: {desc}", "time": datetime.now().strftime("%Y-%m-%d %H:%M:%S")})
                st.success("🎉 تم فتح وتثبيت بلاغ الصيانة بنجاح!")

# --- 4. إدارة التذاكر والبلاغات ---
elif menu == "🖥️ إدارة البلاغات والتذاكر":
    st.title("🖥️ إدارة البلاغات وتعديل التذاكر")
    df_tk = pd.read_sql_query(text("SELECT t.*, c.name as client_name, e.serial_number FROM tickets t JOIN equipment e ON t.equipment_id = e.id JOIN clients c ON e.client_id = c.id ORDER BY t.id DESC"), engine)
    if not df_tk.empty:
        tk_dict = {f"تذكرة رقم {r['id']} - {r['client_name']} ({r['serial_number']})": r['id'] for _, r in df_tk.iterrows()}
        selected_tk = st.selectbox("اختر تذكرة صيانة لتحديثها:", list(tk_dict.keys()))
        tk_id = tk_dict[selected_tk]
        tk_data = df_tk[df_tk['id'] == tk_id].iloc[0]
        with st.form("edit_ticket_form"):
            u_desc = st.text_area("تعديل تفاصيل وصف العطل والموقع المعين:", value=tk_data['issue_description'])
            u_status = st.selectbox("الحالة الحالية للبلاغ:", ["مفتوح / قيد المتابعة", "قيد الانتظار لقطع الغيار", "مشغول لدى عميل", "مغلق"], index=["مفتوح / قيد المتابعة", "قيد الانتظار لقطع الغيار", "مشغول لدى عميل", "مغلق"].index(tk_data['status']) if tk_data['status'] in ["مفتوح / قيد المتابعة", "قيد الانتظار لقطع الغيار", "مشغول لدى عميل", "مغلق"] else 0)
            u_parts = st.text_input("تحديث قطع الغيار المستبدلة إن وجدت:", value=tk_data['parts_replaced'] or "")
            if st.form_submit_button("اعتماد وحفظ البيانات المحدثة للتذكرة"):
                with engine.begin() as conn:
                    conn.execute(text("UPDATE tickets SET issue_description=:desc, status=:status, parts_replaced=:parts WHERE id=:id"), {"desc": u_desc, "status": u_status, "parts": u_parts, "id": int(tk_id)})
                st.success("✅ تم تحديث كافة حقول البلاغ بنجاح!")
                st.rerun()

# --- 5. الزيارات الدورية ---
elif menu == "📅 الزيارات الدورية (PM)":
    st.title("📅 لوحة تحكم ومراقبة الزيارات الوقائية (PM)")
    df_pm = pd.read_sql_query(text("SELECT p.id, c.name as client_name, e.brand, e.model, e.serial_number, e.pm_visits_count, p.scheduled_date, p.status FROM pm_visits p JOIN equipment e ON p.equipment_id = e.id JOIN clients c ON e.client_id = c.id ORDER BY p.scheduled_date ASC"), engine)
    if not df_pm.empty:
        df_pm.columns = ['رقم الزيارة', 'العميل', 'الماركة', 'الموديل', 'السيريال (S/N)', 'الزيارات الدورية السنوية المجدولة', 'تاريخ الزيارة', 'الحالة']
        st.dataframe(df_pm, use_container_width=True)
    else: st.info("لا توجد زيارات وقائية مجدولة حالياً.")

# --- 6. إدارة العملاء والأجهزة ---
elif menu == "🏢 إدارة العملاء والأجهزة (Profile)":
    st.title("🗂️ ملفات العملاء والأجهزة الذكية")
    tab1, tab2 = st.tabs(["🗂️ ملفات العملاء وجرد الأجهزة التابعة لهم", "➕ إضافة عميل/شركة جديدة"])
    with tab1:
        clients_df = pd.read_sql_query(text("SELECT * FROM clients ORDER BY id"), engine)
        if not clients_df.empty:
            client_dict = {row['name']: row['id'] for _, row in clients_df.iterrows()}
            selected_client_id = client_dict[st.selectbox("ابحث عن عميل للوصول لكافة حقوله وأجهزته المعنية:", list(client_dict.keys()))]
            c_info = clients_df[clients_df['id'] == selected_client_id].iloc[0]
            with st.form("edit_client_advanced"):
                u_name = st.text_input("اسم العميل أو الجهة المختصة *:", value=c_info['name'])
                u_addr = st.text_input("العنوان والمقر الميداني المعتمد:", value=c_info['address'] or "")
                if st.form_submit_button("حفظ التغييرات الشاملة لبروفايل العميل"):
                    with engine.begin() as conn:
                        conn.execute(text("UPDATE clients SET name=:name, address=:addr WHERE id=:id"), {"name": u_name, "addr": u_addr, "id": int(selected_client_id)})
                    st.success("✅ تم تحديث حقول ملف العميل بنج