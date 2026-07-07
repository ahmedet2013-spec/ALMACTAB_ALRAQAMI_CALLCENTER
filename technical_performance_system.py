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

# --- القائمة الجانبية الحاكمة ---
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
                <b>📦 ما يشمله العقد بالتفصيل:</b> {eq.get('contract_coverage', 'غير محدد')}<br>
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
            
            # تم تأمين وإصلاح السطر 319 المبتور بالكامل هنا
            with st.form("edit_client_advanced"):
                st.subheader("📝 تعديل وتحديث أي حقل من حقول بروفايل العميل الحالي:")
                u_name = st.text_input("اسم العميل أو الجهة المختصة *:", value=c_info['name'])
                u_addr = st.text_input("العنوان والمقر الميداني المعتمد:", value=c_info['address'] or "")
                if st.form_submit_button("حفظ التغييرات الشاملة لبروفايل العميل"):
                    with engine.begin() as conn:
                        conn.execute(text("UPDATE clients SET name=:name, address=:addr WHERE id=:id"), {"name": u_name, "addr": u_addr, "id": int(selected_client_id)})
                    st.success("✅ تم تحديث حقول ملف العميل بنجاح!")
                    st.rerun()

            equip_df = pd.read_sql_query(text("SELECT * FROM equipment WHERE client_id = :cid"), engine, params={"cid": int(selected_client_id)})
            if not equip_df.empty:
                equip_df['حالة العقد'] = equip_df['sla_expiration_date'].apply(calculate_sla_status)
                for _, machine in equip_df.iterrows():
                    st.markdown(f"""
                    <div style="background-color:#F8FAFC; padding:15px; border-radius:10px; border-right:6px solid #1E3A8A; margin-bottom:12px; line-height:1.7;">
                        <b>🖥️ موديل ونوع الآلة:</b> {machine['brand']} {machine['model']} | <b>🏷️ الرقم التسلسلي (S/N):</b> {machine['serial_number']}<br>
                        <b>📅 تاريخ التركيب:</b> {machine['installation_date']} | <b>📜 نوع العقد:</b> {machine['sla_type']}<br>
                        <b>⏳ حالة العقد والسريان بالأيام:</b> {machine['حالة العقد']}<br>
                        <b>⚙️ الزيارات الدورية (PM):</b> {machine.get('pm_visits_count', 0)} زيارة | <b>💰 قيمة العقد:</b> {machine.get('contract_value', '0')} د.ل
                    </div>
                    """, unsafe_allow_html=True)

# --- 7. 👨‍💻 إدارة فريق الفنيين ---
elif menu == "👨‍💻 إدارة فريق الفنيين":
    st.title("👨‍💻 سجل كادر الفنيين والمهندسين والتحكم بالحالة الميدانية")
    df_techs = pd.read_sql_query(text("SELECT * FROM technicians ORDER BY id"), engine)
    df_all_tk = pd.read_sql_query(text("SELECT t.*, c.name as client_name, e.brand, e.model, e.serial_number FROM tickets t JOIN equipment e ON t.equipment_id = e.id JOIN clients c ON e.client_id = c.id"), engine)

    t_tab1, t_tab2, t_tab3, t_tab4 = st.tabs(["📋 متابعة التوافر الفوري للفريق", "🔍 بروفايل وبطاقة المهندس الذكية", "➕ إضافة مهندس/فني جديد", "✏️ تعديل وتحديث ملف مهندس"])
    status_options = ["متاح", "إجازة سنوية", "إجازة مرضية", "مهمة عمل", "غياب", "استقال", "مشغول لدى عميل"]

    with t_tab1:
        if not df_techs.empty:
            for _, r_tech in df_techs.iterrows():
                st.markdown(f"""
                <div style="background-color:#F8FAFC; padding:15px; border-radius:10px; border-right:6px solid #1E3A8A; margin-bottom:12px;">
                    <h4>👨‍🔧 المهندس بالشركة: {r_tech['name']} ({r_tech.get('specialty', 'عام')})</h4>
                    <b>📍 المدينة المغطاة:</b> {r_tech.get('city', 'طرابلس')} | <b>📱 رقم الهاتف للعمل:</b> {r_tech.get('phone', '---')}<br>
                    <b>🟢 الحالة الميدانية الجارية:</b> <span style="color:#1E3A8A; font-weight:bold;">{r_tech['status']}</span>
                </div>
                """, unsafe_allow_html=True)
                
                if st.button(f"🔍 استعراض البروفايل وبطاقة الأداء لـ {r_tech['name']}", key=f"nav_tech_{r_tech['id']}"):
                    st.session_state['view_tech_id'] = r_tech['id']
                    st.info("🔄 تم اختيار المهندس. فضلاً اضغط على تبويب [🔍 بروفايل وبطاقة المهندس الذكية] بالأعلى لمشاهدة التحليل الكامل.")
        else: st.info("لا توجد أسماء مسجلة.")

    with t_tab2:
        if not df_techs.empty:
            tech_selection_dict = {row['name']: row['id'] for _, row in df_techs.iterrows()}
            default_index = 0
            if st.session_state['view_tech_id'] in tech_selection_dict.values():
                default_index = list(tech_selection_dict.values()).index(st.session_state['view_tech_id'])
            
            selected_tech_profile = st.selectbox("اختر بروفايل المهندس المستهدف لاستعراض مؤشراته الحالية:", list(tech_selection_dict.keys()), index=default_index)
            active_tech_id = tech_selection_dict[selected_tech_profile]
            tech_profile_data = df_techs[df_techs['id'] == active_tech_id].iloc[0]
            
            tech_tickets = df_all_tk[df_all_tk['tech_id'] == active_tech_id]
            completed_tk = tech_tickets[tech_tickets['status'] == 'مغلق']
            ongoing_tk = tech_tickets[tech_tickets['status'] != 'مغلق']
            ftf_count = tech_tickets['first_time_fix'].sum()
            
            st.markdown(f"## 📊 وثيقة تقييم الأداء والمهام للمهندس: {tech_profile_data['name']}")
            kpi_c1, kpi_c2, kpi_c3, kpi_c4 = st.columns(4)
            with kpi_c1: st.markdown(f'<div class="kpi-box">📥 <b>إجمالي البلاغات المسندة</b><br><h3>{len(tech_tickets)}</h3></div>', unsafe_allow_html=True)
            with kpi_c2: st.markdown(f'<div class="kpi-box">✅ <b>بلاغات أنجزت ومغلقة</b><br><h3>{len(completed_tk)}</h3></div>', unsafe_allow_html=True)
            with kpi_c3: st.markdown(f'<div class="kpi-box">⏳ <b>بلاغات قيد الإنجاز</b><br><h3>{len(ongoing_tk)}</h3></div>', unsafe_allow_html=True)
            with kpi_c4: st.markdown(f'<div class="kpi-box">⚙️ <b>الإصلاح من أول زيارة (FTF)</b><br><h3>{ftf_count}</h3></div>', unsafe_allow_html=True)
            
            st.markdown("---")
            if not tech_tickets.empty:
                display_tech_tk = tech_tickets[['id', 'client_name', 'brand', 'model', 'serial_number', 'status', 'time_reported']].copy()
                display_tech_tk.columns = ['رقم التذكرة', 'العميل المستفيد', 'الماركة', 'الموديل', 'السيريال (S/N)', 'حالة البلاغ الحالية', 'تاريخ وتوقيت البلاغ']
                st.dataframe(display_tech_tk.style.set_properties(**{'text-align': 'right', 'direction': 'rtl'}), use_container_width=True)
            else: st.info("💡 لا توجد تذاكر صيانة مسندة لهذا المهندس حالياً.")
        else: st.info("المستودع فارغ.")

    with t_tab3:
        with st.form("add_technician_form"):
            st.subheader("➕ قيد وإدراج مهندس جديد بالبيانات الرسمية:")
            col_f1, col_f2 = st.columns(2)
            with col_f1:
                nt_name = st.text_input("اسم المهندس / الفني الثلاثي بالكامل *:")
                nt_spec = st.text_input("التخصص الفني التقني المعتمد *:")
                nt_city = st.text_input("المدينة أو المنطقة المتواجد بها التغطية *:", value="طرابلس")
            with col_f2:
                nt_phone = st.text_input("رقم الهاتف المحمول المباشر للعمل *:")
                nt_email = st.text_input("البريد الإلكتروني المهني بالشركة:")
                nt_status = st.selectbox("حالة التوافر الميدانية الأولية المعتمدة للجدولة:", status_options)
            uploaded_img = st.file_uploader("قم برفع وإرفاق الصورة الشخصية الرسمية للمهندس بالفولدر السحابي التعريفي:", type=['jpg', 'png', 'jpeg'])
            
            if st.form_submit_button("اعتماد وتثبيت قيد الفني بالفريق"):
                if nt_name and nt_spec and nt_phone:
                    img_path_save = ""
                    if uploaded_img:
                        img_path_save = f"uploads/tech_images/tech_{int(time.time())}.png"
                        with open(img_path_save, "wb") as f: f.write(uploaded_img.getbuffer())
                    with engine.begin() as conn:
                        conn.execute(text("INSERT INTO technicians (name, specialty, phone, email, status, city, image_path) VALUES (:n, :s, :p, :e, :st, :c, :i)"),
                                     {"n": nt_name, "s": nt_spec, "p": nt_phone, "e": nt_email, "st": nt_status, "c": nt_city, "i": img_path_save})
                    st.success("🎉 تم حفظ ملف المهندس الجديد بنجاح!")
                    st.rerun()
                else: st.error("❌ يرجى ملء الحقول الإجبارية.")

    with t_tab4:
        if not df_techs.empty:
            tech_dict_up = {r['name']: r['id'] for _, r in df_techs.iterrows()}
            sel_t_up = st.selectbox("اختر فني أو مهندس لتعديل ملفه الإداري الفوري:", list(tech_dict_up.keys()))
            t_edit_id = tech_dict_up[sel_t_up]
            t_info = df_techs[df_techs['id'] == t_edit_id].iloc[0]
            with st.form("edit_tech_advanced_form"):
                col_ue1, col_ue2 = st.columns(2)
                with col_ue1:
                    u_t_name = st.text_input("تعديل الاسم بالكامل وبدقة الحروف:", value=t_info['name'])
                    u_t_spec = st.text_input("تحديث التخصص التقني المعين للعمل الميداني:", value=t_info.get('specialty', ''))
                    u_t_city = st.text_input("تعديل مدينة ونطاق التغطية الدورية للآلات:", value=t_info.get('city', 'طرابلس'))
                with col_ue2:
                    u_t_phone = st.text_input("تحديث رقم هاتف الاتصال السريع المقيد:", value=t_info.get('phone', ''))
                    u_t_email = st.text_input("تحديث البريد الإلكتروني الرسمي التابع للمكتب:", value=t_info.get('email', ''))
                    u_t_status = st.selectbox("تغيير وتحديث الحالة الفورية الجارية للفني:", status_options, index=status_options.index(t_info['status']) if t_info['status'] in status_options else 0)
                if st.form_submit_button("حفظ وتثبيت كافة التعديلات المحدثة بملف المهندس"):
                    with engine.begin() as conn:
                        conn.execute(text("UPDATE technicians SET name=:name, specialty=:spec, phone=:phone, email=:email, status=:status, city=:city WHERE id=:id"),
                                     {"name": u_t_name, "spec": u_t_spec, "phone": u_t_phone, "email": u_t_email, "status": u_t_status, "city": u_t_city, "id": int(t_edit_id)})
                    st.success("✅ تم تحديث وتطهير ملف المهندس بنجاح!")
                    st.rerun()

# --- 8. إعدادات العقود ---
elif menu == "⚙️ إعدادات أنواع العقود":
    st.title("⚙️ إدارة مسميات وأنواع العقود المركزية (SLAs)")
    sla_df = pd.read_sql_query(text("SELECT id, name as \"نوع العقد المعتمد\" FROM sla_types ORDER BY id"), engine)
    col_s1, col_s2 = st.columns([1, 2])
    with col_s1:
        with st.form("add_sla_advanced"):
            new_name = st.text_input("اكتب تصنيف ومسمى العقد المركزي الجديد ليكون متاحاً بالخيارات:")
            if st.form_submit_button("حفظ وتأكيد الإضافة"):
                if new_name:
                    try:
                        with engine.begin() as conn: conn.execute(text("INSERT INTO sla_types (name) VALUES (:name)"), {"name": new_name})
                        st.success("تمت الإضافة بنجاح!"); st.rerun()
                    except: st.error("❌ هذا المسمى موجود مسبقاً.")
    with col_s2: st.dataframe(sla_df, use_container_width=True)

# --- 9. 🔐 إدارة مستخدمي المنظومة ---
elif menu == "🔐 إدارة مستخدمي المنظومة":
    st.title("🔐 الإدارة المركزية لمستخدمي المنظومة وتغيير كلمات السر")
    u_tab1, u_tab2 = st.tabs(["🔑 تغيير كلمة السر الحالية", "➕ إضافة مستخدم جديد للمنظومة"])
    
    with u_tab1:
        st.subheader("تعديل كلمة مرور الحساب النشط حالياً")
        with st.form("change_password_form"):
            current_user = st.session_state['user_fullname']
            old_pass = st.text_input("كلمة المرور الحالية للتحقق الأمن:", type="password")
            new_pass = st.text_input("كلمة المرور البديلة الجديدة *:", type="password")
            confirm_pass = st.text_input("تأكيد كتابة كلمة المرور الجديدة *:", type="password")
            
            if st.form_submit_button("تحديث كلمة السر ومزامنة الحساب"):
                if new_pass == confirm_pass and new_pass != "":
                    chk_p = pd.read_sql_query(text("SELECT id FROM system_users WHERE full_name = :f AND password = :p"), engine, params={"f": current_user, "p": old_pass})
                    if not chk_p.empty:
                        with engine.begin() as conn:
                            conn.execute(text("UPDATE system_users SET password = :np WHERE full_name = :f"), {"np": new_pass, "f": current_user})
                        st.success("✅ تم تحديث كلمة السر بنجاح! سيتم تطبيق التغييرات فوراً.")
                    else: st.error("❌ كلمة المرور الحالية غير صحيحة، يرجى التثبت.")
                else: st.warning("تأكد من تطابق كلمتي المرور الجديدة وعدم ترك الخانات فارغة.")
                
    with u_tab2:
        st.subheader("إضافة مستخدم جديد وصلاحية وصول مستقلة للسحابة")
        with st.form("create_new_user_form"):
            new_username = st.text_input("اسم المستخدم الجديد للدخول (بإلإنجليزية بدون فراغات) *:")
            new_fullname = st.text_input("الاسم الثلاثي الكامل للموظف/المدير *:")
            new_password = st.text_input("كلمة المرور المبدئية المخصصة للحساب *:", type="password")
            new_role = st.selectbox("الصلاحية والامتيازات الوظيفية:", ["Admin", "Technical Manager", "Call Center Staff"])
            
            if st.form_submit_button("إصدار وتأصيل الحساب السحابي الجديد"):
                if new_username and new_fullname and new_password:
                    try:
                        with engine.begin() as conn:
                            conn.execute(text("INSERT INTO system_users (username, password, full_name, role) VALUES (:u, :p, :f, :r)"),
                                         {"u": new_username, "p": new_password, "f": new_fullname, "r": new_role})
                        st.success(f"🎉 تم إنشاء حساب المستخدم الموظف [{new_fullname}] بنجاح، وبإمكانه تسجيل الدخول الآن!")
                    except: st.error("❌ اسم المستخدم هذا مسجل مسبقاً لموظف آخر، اختر اسماً فريداً.")
                else: st.error("جميع الحقول الإجبارية معلم عليها بنجمة (*) ويجب ملؤها.")