import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime, timedelta
import os
import time
import io
from sqlalchemy import create_engine, text

# --- الأمان والاتصال السحابي الموحد (Connection Pooling IPv4) ---
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
    
    .worst-card {
        background-color: #FEF2F2;
        padding: 20px;
        border-radius: 10px;
        border: 2px solid #EF4444;
        border-right: 10px solid #EF4444;
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

if not os.path.exists("uploads/tech_images"):
    os.makedirs("uploads/tech_images")

# --- دالة حساب حالة العقد والأيام بدقة ---
def calculate_sla_status(exp_date_str):
    if not exp_date_str: 
        return "غير مححدد"
    try:
        exp_date = pd.to_datetime(exp_date_str).date()
        today = datetime.now().date()
        delta = (exp_date - today).days
        if delta > 30: 
            return f"ساري المفعول (متبقي {delta} يوم)"
        elif 0 < delta <= 30:
            return f"شارف على الانتهاء (متبقي {delta} يوم فقط)"
        elif delta == 0: 
            return "ينتهي اليوم"
        else: 
            return f"منتهي الصلاحية (منذ {abs(delta)} يوم)"
    except: 
        return "صيغة غير صحيحة"

def to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='تقرير_المكتب_الرقمي')
    return output.getvalue()

def to_csv_printable(df):
    return df.to_csv(index=False).encode('utf-8-sig')

# --- محرك الأداء المقاوم لغياب الأعمدة ---
def get_performance_extremes():
    query = "SELECT * FROM tickets"
    df_tk = pd.read_sql_query(text(query), engine)
    query_tech = "SELECT * FROM technicians"
    df_th = pd.read_sql_query(text(query_tech), engine)
    
    if df_tk.empty or df_th.empty:
        if not df_th.empty:
            first = df_th.iloc[0]
            return {"name": first['name'], "image": first.get('image_path', ''), "is_estimated": True, "stats": "لا توجد إحصائيات صيانة مسجلة له هذا الشهر حتى الآن"}, None, pd.DataFrame()
        return None, None, pd.DataFrame()
        
    df_all = df_tk.merge(df_th, left_on='tech_id', right_on='id', suffixes=('_ticket', '_tech'))
    current_month = datetime.now().strftime("%Y-%m")
    df_all['time_reported'] = pd.to_datetime(df_all['time_reported'])
    df_filtered = df_all[df_all['time_reported'].dt.strftime("%Y-%m") == current_month]
    
    if df_filtered.empty:
        first = df_th.iloc[0]
        return {"name": first['name'], "image": first.get('image_path', ''), "is_estimated": True, "stats": "لا توجد إغلاقات مسجلة له هذا الشهر حتى الآن"}, None, pd.DataFrame()
        
    tech_stats = df_filtered.groupby('name').agg(
        total_visits=('id_ticket', 'count'),
        first_time_fixes=('first_time_fix', 'sum'),
        pending_count=('status_ticket', lambda x: x.str.contains('انتظار').sum())
    ).reset_index()
    
    # دمج مسار الصورة بأمان
    tech_stats = tech_stats.merge(df_th[['name', 'image_path']], on='name', how='left')
    tech_stats['score'] = (tech_stats['total_visits'] * 5) + (tech_stats['first_time_fixes'] * 10) - (tech_stats['pending_count'] * 4)
    
    sorted_stats = tech_stats.sort_values(by='score', ascending=False)
    winner_row = sorted_stats.iloc[0]
    worst_row = sorted_stats.iloc[-1] if len(sorted_stats) > 1 else None
    
    winner = {"name": winner_row['name'], "image": winner_row.get('image_path', ''), "total_visits": winner_row['total_visits'], "first_time_fixes": winner_row['first_time_fixes']}
    worst = None
    if worst_row:
        worst = {"name": worst_row['name'], "image": worst_row.get('image_path', ''), "total_visits": worst_row['total_visits'], "pending_count": worst_row['pending_count']}
        
    return winner, worst, sorted_stats

if 'selected_sn' not in st.session_state:
    st.session_state['selected_sn'] = ""

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
            if st.form_submit_button("دخول للمنظومة"):
                if username == "Ahmed" and password == "admin123":
                    st.session_state['logged_in'] = True
                    st.rerun()
                else: st.error("❌ بيانات الدخول غير صحيحة.")
    st.stop()

if os.path.exists("MAC logo resized.png"):
    st.sidebar.image("MAC logo resized.png", use_container_width=True)

st.sidebar.title("🛠️ المكتب الرقمي")
menu = st.sidebar.radio("انتقل إلى القائمة الميدانية:", [
    "📊 لوحة التحكم والأداء الشهري", 
    "🔍 البحث الشامل عن جهاز (S/N)",
    "➕ تسجيل بلاغ صيانة جديد", 
    "🖥️ إدارة البلاغات والتذاكر",
    "📅 الزيارات الدورية (PM)",
    "🏢 إدارة العملاء والأجهزة (Profile)",
    "👨‍💻 إدارة فريق الفنيين",
    "⚙️ إعدادات أنواع العقود"
])

if st.sidebar.button("🚪 تسجيل الخروج الآمن"):
    st.session_state['logged_in'] = False
    st.rerun()

# --- 1. لوحة التحكم ---
if menu == "📊 لوحة التحكم والأداء الشهري":
    st.title("📊 الملخص الشهري وتحليل الأداء الفني")
    winner, worst, performance_df = get_performance_extremes()
    
    c1, c2 = st.columns(2)
    with c1:
        if winner:
            st.markdown('<div class="star-card">', unsafe_allow_html=True)
            col_img, col_txt = st.columns([1, 2])
            with col_img:
                w_img = winner.get('image', '')
                if w_img and os.path.exists(str(w_img)): st.image(str(w_img), width=110)
                else: st.markdown("<h1 style='font-size:60px; margin:0;'>🌟</h1>", unsafe_allow_html=True)
            with col_txt:
                if "is_estimated" in winner:
                    st.markdown(f"### 🌟 الموظف المثالي المتوقع: {winner['name']}")
                    st.write(winner['stats'])
                else:
                    st.markdown(f"### 🌟 موظف الشهر المثالي: {winner['name']} ⭐")
                    st.write(f"📊 **إجمالي الزيارات:** {winner['total_visits']} | **الإصلاح الفوري:** {winner['first_time_fixes']}")
            st.markdown('</div>', unsafe_allow_html=True)
            
    with c2:
        if worst:
            st.markdown('<div class="worst-card">', unsafe_allow_html=True)
            col_img2, col_txt2 = st.columns([1, 2])
            with col_img2:
                wr_img = worst.get('image', '')
                if wr_img and os.path.exists(str(wr_img)): st.image(str(wr_img), width=110)
                else: st.markdown("<h1 style='font-size:60px; margin:0;'>⚠️</h1>", unsafe_allow_html=True)
            with col_txt2:
                st.markdown(f"### ⚠️ الأقل أداءً هذا الشهر: {worst['name']}")
                st.write(f"📊 **إجمالي الزيارات:** {worst['total_visits']} | **المعلقة لقطع الغيار:** {worst['pending_count']}")
            st.markdown('</div>', unsafe_allow_html=True)
        else:
            st.markdown('<div class="worst-card">### ⚠️ الأقل أداءً هذا الشهر:<br><p>لا توجد بيانات متباينة للفريق حالياً.</p></div>', unsafe_allow_html=True)

    if not performance_df.empty:
        fig = px.bar(performance_df, x='name', y='score', title="نقاط الأداء المقارنة للفريق", labels={'name': 'المهندس', 'score': 'النقاط'})
        st.plotly_chart(fig, use_container_width=True)
        st.download_button("📥 تصدير الملخص الشهري لـ Excel", to_excel(performance_df), "الملخص_الشهري.xlsx")

# --- 2. البحث الشامل ---
elif menu == "🔍 البحث الشامل عن جهاز (S/N)":
    st.title("🔍 بطاقة تعريف الآلة وبروفايلها الذكي الموحد")
    search_sn = st.text_input("أدخل الرقم التسلسلي (S/N) للآلة للتحقق الفوري:", value=st.session_state['selected_sn']).strip()
    
    if search_sn:
        query = "SELECT e.*, c.name as client_name, c.phone as client_phone FROM equipment e JOIN clients c ON e.client_id = c.id WHERE e.serial_number = :sn"
        eq_df = pd.read_sql_query(text(query), engine, params={"sn": search_sn})
        if not eq_df.empty:
            eq = eq_df.iloc[0]
            st.markdown(f"""
            <div style="background-color:#F8FAFC; padding:20px; border-radius:10px; border-right:8px solid #1E3A8A; line-height:1.8;">
                <h3>🖥️ الموديل: {eq['brand']} {eq['model']}</h3>
                <b>🏢 العميل:</b> {eq['client_name']} | <b>📞 الهاتف:</b> {eq['client_phone']}<br>
                <b>🏷️ السيريال:</b> {eq['serial_number']} | <b>📅 التركيب:</b> {eq['installation_date']}<br>
                <b>📜 العقد:</b> {eq['sla_type']} | <b>⏳ غطاء العقد:</b> {calculate_sla_status(eq['sla_expiration_date'])}<br>
                <b>💼 شراء منا:</b> {eq['purchased_from_us']}
            </div>
            """, unsafe_allow_html=True)
        else: st.error("⚠️ الرقم التسلسلي المدخل غير مسجل في قاعدة البيانات.")

# --- 3. تسجيل بلاغ جديد ---
elif menu == "➕ تسجيل بلاغ صيانة جديد":
    st.title("➕ فتح وتوثيق بلاغ صيانة ميداني جديد")
    eq_rows = pd.read_sql_query(text("SELECT e.id, c.name as client_name, e.brand, e.model, e.serial_number FROM equipment e JOIN clients c ON e.client_id = c.id"), engine)
    tech_rows = pd.read_sql_query(text("SELECT id, name FROM technicians"), engine)
    
    if eq_rows.empty or tech_rows.empty:
        st.error("⚠️ يجب قيد أجهزة وفنيين أولاً في المنظومة.")
    else:
        equip_options = {f"{r['client_name']} - {r['brand']} {r['model']} (S/N: {r['serial_number']})": r['id'] for _, r in eq_rows.iterrows()}
        selected_equip_label = st.selectbox("اختر الآلة المشكو منها:", list(equip_options.keys()))
        eq_id_sel = equip_options[selected_equip_label]
        
        target_eq = pd.read_sql_query(text("SELECT e.*, c.name as client_name FROM equipment e JOIN clients c ON e.client_id = c.id WHERE e.id = :eid"), engine, params={"eid": int(eq_id_sel)}).iloc[0]
        st.warning(f"🔍 **التحقق التلقائي للآلة:** نوع العقد: {target_eq['sla_type']} ({calculate_sla_status(target_eq['sla_expiration_date'])})")
        
        with st.form("new_ticket_form"):
            col1, col2 = st.columns(2)
            with col1:
                m_type = st.selectbox("نوع الآلة:", ["طابعة إنتاج ليزر", "طابعة عريضة", "منظومة أرشفة PLC"])
                m_model = st.text_input("تأكيد الموديل:", value=f"{target_eq['brand']} {target_eq['model']}")
                issue_type = st.selectbox("تصنيف العطل:", ["عطل ميكانيكي", "جودة ألوان", "خطأ في لوحة التحكم PLC"])
            with col2:
                priority = st.selectbox("مستوى حساسية العطل *:", ["استجابة سريعة فورية", "استجابة عادية"])
                tech_name = st.selectbox("المهندس المسؤول عن المهمة *:", [r['name'] for _, r in tech_rows.iterrows()])
                issue_desc = st.text_area("وصف مفصل للعطل:")
                
            if st.form_submit_button("إصدار وتثبيت البلاغ"):
                tech_id_final = tech_rows[tech_rows['name'] == tech_name].iloc[0]['id']
                t_now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                with engine.begin() as conn:
                    conn.execute(text("INSERT INTO tickets (equipment_id, tech_id, issue_description, time_reported, status, first_time_fix, parts_replaced) VALUES (:eq_id, :tech_id, :desc, :time_rep, 'مفتوح / قيد المتابعة', 0, '')"),
                                 {"eq_id": int(eq_id_sel), "tech_id": int(tech_id_final), "desc": f"[{priority}] {m_type} - {issue_type}: {issue_desc}", "time_rep": t_now})
                st.success("🎉 تم فتح وتثبيت بلاغ الصيانة بنجاح وإسناده للمهندس!")

# --- 4. إدارة التذاكر والبلاغات ---
elif menu == "🖥️ إدارة البلاغات والتذاكر":
    st.title("🖥️ إدارة البلاغات وتحديث تذاكر الأعطال")
    query_all_tk = "SELECT t.*, c.name as client_name, e.brand, e.model, e.serial_number, tech.name as tech_name FROM tickets t JOIN equipment e ON t.equipment_id = e.id JOIN clients c ON e.client_id = c.id JOIN technicians tech ON t.tech_id = tech.id ORDER BY t.id DESC"
    df_all_tickets = pd.read_sql_query(text(query_all_tk), engine)
    
    if df_all_tickets.empty:
        st.info("لا توجد بلاغات مسجلة.")
    else:
        ticket_dict = {f"بلاغ رقم {r['id']} - لعميل: {r['client_name']} (S/N: {r['serial_number']})": r['id'] for _, r in df_all_tickets.iterrows()}
        sel_tk_label = st.selectbox("اختر البلاغ المراد تعديله أو تحديثه:", list(ticket_dict.keys()))
        t_id_edit = ticket_dict[sel_tk_label]
        c_tk = df_all_tickets[df_all_tickets['id'] == t_id_edit].iloc[0]
        
        with st.form("edit_form"):
            st.info(f"📆 **توقيت الإنشاء المعتمد:** {c_tk['time_reported']}")
            u_desc = st.text_area("وصف العطل الحالي:", value=c_tk['issue_description'])
            u_status = st.selectbox("تحديث حالة التذكرة:", ["مفتوح / قيد المتابعة", "قيد الانتظار لقطع الغيار", "مشغول لدى عميل", "مغلق"], index=["مفتوح / قيد المتابعة", "قيد الانتظار لقطع الغيار", "مشغول لدى عميل", "مغلق"].index(c_tk['status']) if c_tk['status'] in ["مفتوح / قيد المتابعة", "قيد الانتظار لقطع الغيار", "مشغول لدى عميل", "مغلق"] else 0)
            u_parts = st.text_input("قطع الغيار المستبدلة:", value=c_tk['parts_replaced'] or "")
            u_ftf = st.checkbox("تم الإصلاح من أول زيارة (First Time Fix)", value=bool(c_tk['first_time_fix']))
            
            if st.form_submit_button("حفظ التغييرات بالتذكرة"):
                res_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S") if u_status == "مغلق" else c_tk['time_resolved']
                with engine.begin() as conn:
                    conn.execute(text("UPDATE tickets SET issue_description=:desc, status=:status, first_time_fix=:ftf, parts_replaced=:parts, time_resolved=:tres WHERE id=:id"),
                                 {"desc": u_desc, "status": u_status, "ftf": 1 if u_ftf else 0, "parts": u_parts, "tres": res_time, "id": int(t_id_edit)})
                st.success("✅ تم تحديث بيانات التذكرة بنجاح!")
                st.rerun()

        display_all_tk = df_all_tickets[['id', 'client_name', 'brand', 'model', 'serial_number', 'tech_name', 'status']].copy()
        display_all_tk.columns = ['رقم التذكرة', 'العميل', 'الماركة', 'الموديل', 'السيريال (S/N)', 'المهندس المسؤول', 'الحالة']
        st.dataframe(display_all_tk.style.set_properties(**{'text-align': 'right', 'direction': 'rtl'}), use_container_width=True)

# --- 5. الزيارات الدورية ---
elif menu == "📅 الزيارات الدورية (PM)":
    st.title("📅 لوحة تحكم ومراقبة الزيارات الوقائية (PM)")
    query_pm = "SELECT p.id, c.name, e.brand, e.model, e.serial_number, p.scheduled_date, p.status FROM pm_visits p JOIN equipment e ON p.equipment_id = e.id JOIN clients c ON e.client_id = c.id ORDER BY p.scheduled_date ASC"
    pm_df = pd.read_sql_query(text(query_pm), engine)
    if not pm_df.empty:
        pm_df.columns = ['رقم الزيارة', 'العميل', 'الماركة', 'الموديل', 'السيريال', 'تاريخ الزيارة', 'الحالة']
        st.dataframe(pm_df.style.set_properties(**{'text-align': 'right', 'direction': 'rtl'}), use_container_width=True)
    else: st.info("لا توجد زيارات وقائية مجدولة.")

# --- 6. ملفات العملاء والأجهزة ---
elif menu == "🏢 إدارة العملاء والأجهزة (Profile)":
    st.title("🗂️ ملفات العملاء والأجهزة الذكية")
    tab1, tab2, tab3 = st.tabs(["🗂️ بروفايل العميل وأجهزته", "➕ عميل جديد", "📋 مستودع كل الأجهزة"])
    
    with tab1:
        clients_df = pd.read_sql_query(text("SELECT * FROM clients ORDER BY id"), engine)
        if not clients_df.empty:
            client_dict = {row['name']: row['id'] for _, row in clients_df.iterrows()}
            selected_client_id = client_dict[st.selectbox("🔍 ابحث عن عميل لاستعراض أجهزته وملفه:", list(client_dict.keys()))]
            c_info = clients_df[clients_df['id'] == selected_client_id].iloc[0]
            
            with st.expander(f"✏️ تعديل بيانات العميل ({c_info['name']})"):
                with st.form("edit_client"):
                    u_cl_name = st.text_input("اسم العميل *:", value=c_info['name'])
                    u_cl_addr = st.text_input("العنوان:", value=c_info['address'] or "")
                    if st.form_submit_button("حفظ التحديثات"):
                        with engine.begin() as conn:
                            conn.execute(text("UPDATE clients SET name=:name, address=:addr WHERE id=:id"), {"name": u_cl_name, "addr": u_cl_addr, "id": int(selected_client_id)})
                        st.success("✅ تم التحديث بنجاح!")
                        st.rerun()

            equip_df = pd.read_sql_query(text("SELECT * FROM equipment WHERE client_id = :cid"), engine, params={"cid": int(selected_client_id)})
            if not equip_df.empty:
                st.markdown(f"### 🖨️ الأجهزة التابعة لـ {c_info['name']}")
                equip_df['حالة العقد'] = equip_df['sla_expiration_date'].apply(calculate_sla_status)
                
                display_equip_df = equip_df[['brand', 'model', 'serial_number', 'installation_date', 'sla_type', 'حالة العقد']].copy()
                display_equip_df.columns = ['نوع الآلة', 'الموديل', 'الرقم التسلسلي (S/N)', 'تاريخ التركيب', 'نوع العقد', 'حالة العقد والسريان']
                
                st.dataframe(display_equip_df.style.set_properties(**{'text-align': 'right', 'direction': 'rtl'}), use_container_width=True)
                st.download_button("📥 تصدير أجهزة العميل لـ Excel", to_excel(equip_df), f"أجهزة_{c_info['name']}.xlsx")
            else: st.info("لا توجد أجهزة مسجلة لهذا العميل.")

    with tab2:
        with st.form("add_client"):
            nc_name = st.text_input("اسم العميل أو الجهة بالكامل *:")
            nc_addr = st.text_input("المقر والعنوان الميداني *:")
            if st.form_submit_button("إدراج عميل جديد"):
                if nc_name and nc_addr:
                    with engine.begin() as conn:
                        conn.execute(text("INSERT INTO clients (name, address) VALUES (:name, :addr)"), {"name": nc_name, "addr": nc_addr})
                    st.success("🎉 تم الحفظ بنجاح!")
                    st.rerun()

    with tab3:
        all_eq = pd.read_sql_query(text("SELECT c.name as client_name, e.* FROM equipment e JOIN clients c ON e.client_id = c.id ORDER BY e.id DESC"), engine)
        if not all_eq.empty:
            all_eq['حالة العقد'] = all_eq['sla_expiration_date'].apply(calculate_sla_status)
            display_all_eq = all_eq[['client_name', 'brand', 'model', 'serial_number', 'sla_type', 'حالة العقد']].copy()
            display_all_eq.columns = ['العميل المالك', 'نوع الآلة', 'الموديل', 'الرقم التسلسلي (S/N)', 'نوع العقد', 'حالة العقد']
            st.dataframe(display_all_eq.style.set_properties(**{'text-align': 'right', 'direction': 'rtl'}), use_container_width=True)

# --- 7. إدارة الفنيين والمهندسين ---
elif menu == "👨‍💻 إدارة فريق الفنيين":
    st.title("👨‍💻 سجل كادر الفنيين والمهندسين والتحكم بالحالة الميدانية")
    tech_tab1, tech_tab2, tech_tab3 = st.tabs(["📋 قائمة الفنيين ومتابعة الحالة الميدانية", "➕ إضافة فني/مهندس جديد", "✏️ تعديل بيانات الملف الشخصي"])
    status_options = ["متاح", "إجازة سنوية", "إجازة مرضية", "مهمة عمل", "غياب", "استقال", "مشغول لدى عميل"]
    
    df_techs = pd.read_sql_query(text("SELECT * FROM technicians ORDER BY id"), engine)

    with tech_tab1:
        if not df_techs.empty:
            for _, r_tech in df_techs.iterrows():
                # استخدام .get() بأمان تام لمنع انهيار البرنامج وظهور KeyError 7
                t_city = r_tech.get('city', 'طرابلس') if 'city' in df_techs.columns else 'طرابلس'
                t_phone = r_tech.get('phone', '---') if 'phone' in df_techs.columns else '---'
                t_email = r_tech.get('email', '---') if 'email' in df_techs.columns else '---'
                t_img = r_tech.get('image_path', '') if 'image_path' in df_techs.columns else ''
                
                st.markdown(f"""
                <div style="background-color:#F8FAFC; padding:15px; border-radius:10px; border-right:6px solid #1E3A8A; margin-bottom:15px;">
                    <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap;">
                        <div>
                            <h4>👨‍🔧 المهندس: {r_tech['name']} ({r_tech.get('specialty', 'عام')})</h4>
                            <b>📍 المدينة المتواجد بها:</b> {t_city} | <b>📱 هاتف:</b> {t_phone} | <b>📧 إيميل:</b> {t_email}<br>
                            <b>🟢 الحالة الفورية الميدانية:</b> <span style="font-weight:bold; color:#1E3A8A;">{r_tech['status']}</span>
                        </div>
                    </div>
                </div>
                """, unsafe_allow_html=True)
                
                if r_tech['status'] == "مشغول لدى عميل":
                    st.info(f"🔗 المهندس مشغول بمهمة صيانة جارية حالياً، يمكنك مراجعة صفحة البلاغات للتتبع.")
        else: st.info("لا يوجد مهندسون مسجلون.")
            
    with tech_tab2:
        with st.form("add_tech_form"):
            nt_name = st.text_input("اسم المهندس / الفني بالكامل *:")
            nt_spec = st.text_input("التخصص الفني التقني *:")
            if st.form_submit_button("اعتماد وإدراج المهندس"):
                if nt_name and nt_spec:
                    with engine.begin() as conn:
                        conn.execute(text("INSERT INTO technicians (name, specialty, status) VALUES (:name, :spec, 'متاح')"), {"name": nt_name, "spec": nt_spec})
                    st.success("🎉 تم الحفظ والمزامنة السحابية بنجاح!")
                    st.rerun()

    with tech_tab3:
        st.info("لتحديث بيانات الفنيين المتقدمة، يرجى مراجعة الخيارات المتاحة أو تحديث قاعدة البيانات.")

# --- 8. إعدادات العقود ---
elif menu == "⚙️ إعدادات أنواع العقود":
    st.title("⚙️ إدارة مسميات وأنواع العقود المركزية")
    sla_df = pd.read_sql_query(text("SELECT id, name as \"نوع العقد المعتمد\" FROM sla_types ORDER BY id"), engine)
    
    col_s1, col_s2 = st.columns([1, 2])
    with col_s1:
        with st.form("add_sla"):
            new_sla_name = st.text_input("اكتب مسمى العقد المركزي الجديد:")
            if st.form_submit_button("حفظ وتأكيد الإضافة"):
                if new_sla_name:
                    try:
                        with engine.begin() as conn:
                            conn.execute(text("INSERT INTO sla_types (name) VALUES (:name)"), {"name": new_sla_name})
                        st.success("تمت الإضافة بنجاح للمستودع!")
                        st.rerun()
                    except: st.error("❌ مسمى العقد مسجل وموجود مسبقاً.")
    with col_s2:
        st.dataframe(sla_df.style.set_properties(**{'text-align': 'right', 'direction': 'rtl'}), use_container_width=True)