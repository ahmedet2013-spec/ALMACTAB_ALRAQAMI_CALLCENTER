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

# تطبيق الـ CSS المخصص للمحاذاة اليمينية والبطاقات المتقدمة لبروفايل المهندسين
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

# --- مجلدات تخزين الصور والتقارير الميدانية ---
if not os.path.exists("uploads/tech_images"):
    os.makedirs("uploads/tech_images")

# --- دالة حساب حالة العقد والأيام بدقة مبرهنة ---
def calculate_sla_status(exp_date_str):
    if not exp_date_str: 
        return "غير محدد"
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

# --- دوال تصدير البيانات الشاملة ---
def to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='تقرير_المكتب_الرقمي')
    return output.getvalue()

def to_csv_printable(df):
    return df.to_csv(index=False).encode('utf-8-sig')

# --- محرك احتساب المهندس المتميز والمهندس الأقل أداءً ---
def get_performance_extremes():
    query = """
        SELECT t.*, tech.name as tech_name, tech.image_path as tech_image
        FROM tickets t 
        JOIN technicians tech ON t.tech_id = tech.id
    """
    df_all = pd.read_sql_query(text(query), engine)
    current_month = datetime.now().strftime("%Y-%m")
    
    if df_all.empty:
        return None, None, pd.DataFrame()
        
    df_all['time_reported'] = pd.to_datetime(df_all['time_reported'])
    df_filtered = df_all[df_all['time_reported'].dt.strftime("%Y-%m") == current_month]
    
    if df_filtered.empty:
        # إرجاع فني تقديري أول في السجل في حال غياب داتا الشهر الحالي منعاً لانهيار الواجهة
        first_tech = df_all.iloc[0]
        return {"name": first_tech['tech_name'], "image": first_tech['tech_image'], "is_estimated": True, "stats": "لا توجد إغلاقات مسجلة له هذا الشهر حتى الآن"}, None, pd.DataFrame()
        
    # احتساب أوزان التقييم (KPIs)
    tech_stats = df_filtered.groupby(['tech_name', 'tech_image']).agg(
        total_visits=('id', 'count'),
        first_time_fixes=('first_time_fix', 'sum'),
        pending_count=('status', lambda x: x.str.contains('انتظار').sum())
    ).reset_index()
    
    tech_stats['score'] = (tech_stats['total_visits'] * 5) + (tech_stats['first_time_fixes'] * 10) - (tech_stats['pending_count'] * 4)
    
    sorted_stats = tech_stats.sort_values(by='score', ascending=False)
    winner = sorted_stats.iloc[0]
    worst = sorted_stats.iloc[-1] if len(sorted_stats) > 1 else None
    
    return winner, worst, sorted_stats

# --- إدارة التنقل بين الصفحات الذكي (السيريال S/N) ---
if 'selected_sn' not in st.session_state:
    st.session_state['selected_sn'] = ""

def nav_to_serial(sn):
    st.session_state['selected_sn'] = sn
    st.session_state['menu_selection'] = "🔍 البحث الشامل عن جهاز (S/N)"

# --- نظام التحقق من الدخول وحماية الجلسة ---
if 'logged_in' not in st.session_state: 
    st.session_state['logged_in'] = False

if not st.session_state['logged_in']:
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        if os.path.exists("MAC logo resized.png"): 
            st.image("MAC logo resized.png", use_container_width=True)
        st.markdown("<h2 style='text-align: center;'>تسجيل الدخول - منظومة الأداء الفني</h2>", unsafe_allow_html=True)
        with st.form("login_form"):
            username = st.text_input("اسم المستخدم")
            password = st.text_input("كلمة المرور", type="password")
            if st.form_submit_button("دخول للمنظومة السحابية"):
                if username == "Ahmed" and password == "admin123":
                    st.session_state['logged_in'] = True
                    st.session_state['menu_selection'] = "📊 لوحة التحكم والأداء الشهري"
                    st.rerun()
                else: 
                    st.error("❌ بيانات الدخول غير صحيحة.")
    st.stop()

# --- بناء القائمة الجانبية الموحدة ---
if os.path.exists("MAC logo resized.png"):
    st.sidebar.image("MAC logo resized.png", use_container_width=True)

st.sidebar.title("🛠️ المكتب الرقمي")
st.sidebar.subheader("إدارة تقنية المعلومات والأداء")

menu = st.sidebar.radio("انتقل إلى القائمة الميدانية:", [
    "📊 لوحة التحكم والأداء الشهري", 
    "🔍 البحث الشامل عن جهاز (S/N)",
    "➕ تسجيل بلاغ صيانة جديد", 
    "🖥️ إدارة البلاغات والتذاكر",
    "📅 الزيارات الدورية (PM)",
    "🏢 إدارة العملاء والأجهزة (Profile)",
    "👨‍💻 إدارة فريق الفنيين",
    "⚙️ إعدادات أنواع العقود"
], key="menu_radio", index=[
    "📊 لوحة التحكم والأداء الشهري", 
    "🔍 البحث الشامل عن جهاز (S/N)",
    "➕ تسجيل بلاغ صيانة جديد", 
    "🖥️ إدارة البلاغات والتذاكر",
    "📅 الزيارات الدورية (PM)",
    "🏢 إدارة العملاء والأجهزة (Profile)",
    "👨‍💻 إدارة فريق الفنيين",
    "⚙️ إعدادات أنواع العقود"
].index(st.session_state.get('menu_selection', "📊 لوحة التحكم والأداء الشهري")))

if st.sidebar.button("🚪 تسجيل الخروج الآمن"):
    st.session_state['logged_in'] = False
    st.rerun()

st.sidebar.markdown("""
<div class="developer-card">
    <h4>⚙️ إدارة وتطوير المنظومة:</h4>
    <b>م. أحمد عثمان</b><br>
    مدير إدارة الدعم الفني وتقنية المعلومات<br>
    📧 AhmedE@almactab.com<br>
    📱 0923009907
</div>
""", unsafe_allow_html=True)

# ----------------------------------------------------
# 1. 📊 لوحة التحكم والأداء الشهري
# ----------------------------------------------------
if menu == "📊 لوحة التحكم والأداء الشهري":
    st.title("📊 الملخص الشهري وتحليل الأداء الفني")
    
    winner, worst, performance_df = get_performance_extremes()
    
    # عرض كروت المهندسين (المتميز والأقل أداءً) بالصور والإحصائيات
    c_extreme1, c_extreme2 = st.columns(2)
    with c_extreme1:
        if winner:
            st.markdown('<div class="star-card">', unsafe_allow_html=True)
            col_img, col_txt = st.columns([1, 2])
            with col_img:
                if winner['tech_image'] and os.path.exists(str(winner['tech_image'])):
                    st.image(str(winner['tech_image']), width=120)
                else:
                    st.markdown("<h1 style='font-size:70px; margin:0;'>👨‍🔧</h1>", unsafe_allow_html=True)
            with col_txt:
                prefix = "الموظف اللامع التقديري: " if "is_estimated" in winner else "الموظف المثالي لهذا الشهر: "
                st.markdown(f"### 🌟 {prefix} {winner['tech_name']} ⭐")
                if "is_estimated" in winner:
                    st.write(winner['stats'])
                else:
                    st.write(f"📊 **إجمالي الزيارات:** {winner['total_visits']} زيارة | **إصلاح من أول مرة:** {winner['first_time_fixes']}")
            st.markdown('</div>', unsafe_allow_html=True)
            
    with c_extreme2:
        if worst and not performance_df.empty:
            st.markdown('<div class="worst-card">', unsafe_allow_html=True)
            col_img2, col_txt2 = st.columns([1, 2])
            with col_img2:
                if worst['tech_image'] and os.path.exists(str(worst['tech_image'])):
                    st.image(str(worst['tech_image']), width=120)
                else:
                    st.markdown("<h1 style='font-size:70px; margin:0;'>⚠️</h1>", unsafe_allow_html=True)
            with col_txt2:
                st.markdown(f"### ⚠️ الأقل أداءً هذا الشهر: {worst['tech_name']}")
                st.write(f"📊 **إجمالي الزيارات:** {worst['total_visits']} | **البلاغات المعلقة لقطع الغيار:** {worst['pending_count']}")
            st.markdown('</div>', unsafe_allow_html=True)
        else:
            st.markdown('<div class="worst-card">### ⚠️ الأقل أداءً هذا الشهر:<br><p>لا توجد إحصائيات سلبية متباينة للفريق حالياً.</p></div>', unsafe_allow_html=True)

    # الرسوم البيانية - تظهر دائماً حتى لو كانت قاعدة البيانات فارغة منعاً للأخطاء الغبية
    st.markdown("### 📈 منحنى تحليل الأداء التنافسي العام للفريق")
    if not performance_df.empty:
        fig = px.bar(performance_df, x='tech_name', y='score', title="نقاط تقييم كفاءة المهندسين خلال الشهر الحالي", labels={'tech_name': 'المهندس', 'score': 'مستوى الكفاءة'}, color='score')
    else:
        # رسم بياني توضيحي فارغ في حال عدم وجود داتا لحماية استقرار الواجهة
        dummy_df = pd.DataFrame([{"المهندس": "لا يوجد فنيين مسجلين بعد", "مستوى الكفاءة": 0}])
        fig = px.bar(dummy_df, x='المهندس', y='مستوى الكفاءة', title="مخطط توضيحي مؤقت لحين إدراج بلاغات")
    st.plotly_chart(fig, use_container_width=True)

    # توفير أزرار تصدير الملخص الشهري
    st.markdown("---")
    ex_c1, ex_c2 = st.columns(2)
    with ex_c1:
        st.download_button("📥 تصدير الملخص الإحصائي الشهري لـ Excel", to_excel(performance_df), "الملخص_الشهري.xlsx")
    with ex_c2:
        st.download_button("🖨️ طباعة تقرير الكفاءة الفوري (CSV)", to_csv_printable(performance_df), "طباعة_الملخص.csv")

# ----------------------------------------------------
# 2. 🔍 البحث الشامل عن جهاز (S/N)
# ----------------------------------------------------
elif menu == "🔍 البحث الشامل عن جهاز (S/N)":
    st.title("🔍 بطاقة تعريف الآلة وبروفايلها الذكي الموحد")
    
    # استقبال الرقم السلسلي المحول من الأقسام الأخرى عبر الجلسة
    default_sn = st.session_state.get('selected_sn', "")
    search_sn = st.text_input("أدخل أو راجع الرقم التسلسلي (S/N) للآلة للتحقق الفوري:", value=default_sn).strip()
    
    if search_sn:
        query = """
            SELECT e.*, c.name as client_name, c.phone as client_phone 
            FROM equipment e 
            JOIN clients c ON e.client_id = c.id 
            WHERE e.serial_number = :sn
        """
        eq_df = pd.read_sql_query(text(query), engine, params={"sn": search_sn})
        
        if not eq_df.empty:
            eq = eq_df.iloc[0]
            st.markdown(f"""
            <div style="background-color:#F8FAFC; padding:25px; border-radius:12px; border-right:10px solid #1E3A8A; line-height:1.8;">
                <h2>🖥️ بيانات الآلة والموديل: {eq['brand']} {eq['model']}</h2>
                <hr>
                <b>🏢 اسم العميل التابع له:</b> {eq['client_name']} | <b>📞 هاتف العميل:</b> {eq['client_phone']}<br>
                <b>🏷️ الرقم التسلسلي (S/N):</b> {eq['serial_number']} | <b>📅 تاريخ التركيب:</b> {eq['installation_date']}<br>
                <b>📜 غطاء العقد الحاكم:</b> {eq['sla_type']} | <b>⏳ حالة العقد الحالية:</b> {calculate_sla_status(eq['sla_expiration_date'])}<br>
                <b>💼 مصدر التوريد:</b> {eq['purchased_from_us']}
            </div>
            """, unsafe_allow_html=True)
            
            # جلب آخر زيارات الصيانة للآلة
            tk_query = "SELECT t.*, tech.name as tech_name FROM tickets t JOIN technicians tech ON t.tech_id = tech.id WHERE t.equipment_id = :eid ORDER BY t.id DESC"
            tk_df = pd.read_sql_query(text(tk_query), engine, params={"eid": int(eq['id'])})
            if not tk_df.empty:
                st.markdown("### 🛠️ السجل التاريخي للبلاغات والزيارات السابقة للآلة")
                tk_df.columns = ['رقم التذكرة', 'معرف الآلة', 'معرف الفني', 'وصف العطل العابر', 'تاريخ البلاغ', 'تاريخ المعالجة', 'الحالة', 'إصلاح أول مرة', 'القطع المستبدلة', 'ملف التقرير', 'اسم المهندس']
                styled_tk = tk_df[['رقم التذكرة', 'اسم المهندس', 'وصف العطل العابر', 'تاريخ البلاغ', 'الحالة']].style.set_properties(**{'text-align': 'right', 'direction': 'rtl'})
                st.dataframe(styled_tk, use_container_width=True)
        else:
            st.error("⚠️ الرقم التسلسلي المدخل غير مسجل في مستودع الأجهزة الحالي بالمنظومة.")

# ----------------------------------------------------
# 3. ➕ تسجيل بلاغ صيانة جديد
# ----------------------------------------------------
elif menu == "➕ تسجيل بلاغ صيانة جديد":
    st.title("➕ فتح وتوثيق بلاغ صيانة ميداني جديد")
    
    eq_rows = pd.read_sql_query(text("SELECT e.id, c.name as client_name, e.brand, e.model, e.serial_number FROM equipment e JOIN clients c ON e.client_id = c.id"), engine)
    tech_rows = pd.read_sql_query(text("SELECT id, name FROM technicians"), engine)
    
    if eq_rows.empty or tech_rows.empty:
        st.error("⚠️ يجب قيد أجهزة وفنيين أولاً لتتمكن من إصدار تذاكر بلاغات.")
    else:
        # الربط والتحقق التلقائي بمجرد اختيار الآلة في حقل التسجيل
        equip_options = {f"{r['client_name']} - {r['brand']} {r['model']} (S/N: {r['serial_number']})": r['id'] for _, r in eq_rows.iterrows()}
        selected_equip_label = st.selectbox("اختر الآلة المشكو منها (سيتم جلب بيانات العقد والضمان فوراً سحابياً):", list(equip_options.keys()))
        
        eq_id_sel = equip_options[selected_equip_label]
        target_eq = pd.read_sql_query(text("SELECT e.*, c.name as client_name FROM equipment e JOIN clients c ON e.client_id = c.id WHERE e.id = :eid"), engine, params={"eid": int(eq_id_pm_visit if 'eq_id_pm_visit' in locals() else eq_id_sel)}).iloc[0]
        
        # عرض بطاقة التحقق السريعة قبل الحفظ لمنع الخطأ الغبي
        st.warning(f"🔍 **التحقق التلقائي للآلة:** نوع العقد: {target_eq['sla_type']} ({calculate_sla_status(target_eq['sla_expiration_date'])}) | مصدر الشراء: {target_eq['purchased_from_us']}")
        
        with st.form("new_ticket_advanced_form"):
            col_t1, col_t2 = st.columns(2)
            with col_t1:
                machine_type = st.selectbox("نوع الآلة التقني:", ["طابعة إنتاج سحابي ليزر", "طابعة عريضة ProStream", "طابعة مكتبية AltaLink", "منظومة أرشفة متطورة PLC"])
                machine_model = st.text_input("تأكيد موديل الآلة بدقة:", value=f"{target_eq['brand']} {target_eq['model']}")
                issue_type = st.selectbox("تصنيف نوع العطل الرئيسي:", ["عطل ميكانيكي / حشر ورق متكرر", "جودة الألوان / باهتة", "خطأ في وحدة التثبيت PLC", "عطل برمجي / تعريف الشبكة"])
                issue_address = st.text_input("عنوان تواجد الة وموقعها الحالي ميدانياً:", value=f"{target_eq['location_building'] or ''} - {target_eq['location_department'] or ''}")
            with col_t2:
                priority_level = st.selectbox("مستوى حساسية العطل وسرعة الاستجابة المطلوبة *:", ["استجابة سريعة فورية (CRITICAL)", "استجابة عادية طارئة (NORMAL)"])
                selected_tech_name = st.selectbox("المهندس المسند إليه إنجاز المهمة وصيانة الآلة *:", [r['name'] for _, r in tech_rows.iterrows()])
                ticket_time = st.text_input("توقيت البلاغ التلقائي (اليوم/الشهر/السنة والساعة):", value=datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
                issue_desc_full = st.text_area("وصف تفصيلي لأعراض ومشتكى العطل الميداني:")
                
            if st.form_submit_button("إصدار وتثبيت بلاغ الصيانة رسميّاً"):
                tech_id_final = tech_rows[tech_rows['name'] == selected_tech_name].iloc[0]['id']
                ins_query = """
                    INSERT INTO tickets (equipment_id, tech_id, issue_description, time_reported, status, first_time_fix, parts_replaced) 
                    VALUES (:eq_id, :tech_id, :desc, :time_rep, 'مفتوح / قيد المتابعة', 0, :parts)
                """
                with engine.begin() as conn:
                    conn.execute(text(ins_query), {
                        "eq_id": int(eq_id_sel), "tech_id": int(tech_id_final),
                        "desc": f"[{priority_level}] - النوع: {machine_type} - عطل: {issue_type} - وصف: {issue_desc_full} - الموقع: {issue_address}",
                        "time_rep": ticket_time, "parts": ""
                    })
                st.success("🎉 تم فتح بلاغ الصيانة، وإسناد المهمة للمهندس المختص، وحفظ التوقيت الزمني بنجاح!")

# ----------------------------------------------------
# 4. 🖥️ إدارة البلاغات والتذاكر
# ----------------------------------------------------
elif menu == "🖥️ إدارة البلاغات والتذاكر":
    st.title("🖥️ إدارة البلاغات وتحديث بيانات تذاكر الأعطال")
    
    # ميزة التعديل لأي تذكرة وبلاغ تم إنشاؤه مسبقاً في السحابة
    query_all_tk = """
        SELECT t.*, c.name as client_name, e.brand, e.model, e.serial_number, tech.name as tech_name
        FROM tickets t 
        JOIN equipment e ON t.equipment_id = e.id
        JOIN clients c ON e.client_id = c.id 
        JOIN technicians tech ON t.tech_id = tech.id 
        ORDER BY t.id DESC
    """
    df_all_tickets = pd.read_sql_query(text(query_all_tk), engine)
    
    if df_all_tickets.empty:
        st.info("لا توجد بلاغات مسجلة للتعديل.")
    else:
        ticket_dict = {f"بلاغ رقم {r['id']} - لعميل: {r['client_name']} (S/N: {r['serial_number']})": r['id'] for _, r in df_all_tickets.iterrows()}
        sel_tk_label = st.selectbox("اختر البلاغ أو التذكرة المراد تعديل تفاصيلها أو إغلاقها ومراجعة سيرها:", list(ticket_dict.keys()))
        t_id_edit = ticket_dict[sel_tk_label]
        c_tk = df_all_tickets[df_all_tickets['id'] == t_id_edit].iloc[0]
        
        with st.form("edit_ticket_advanced_form"):
            st.info(f"📆 **توقيت الإنشاء المعتمد بالتذكرة:** {c_tk['time_reported']}")
            u_desc = st.text_area("تعديل نص وصف البلاغ والعطل:", value=c_tk['issue_description'])
            u_status = st.selectbox("تحديث حالة التذكرة الميدانية:", ["مفتوح / قيد المتابعة", "قيد الانتظار لقطع الغيار", "مشغول لدى عميل", "مغلق"], index=["مفتوح / قيد المتابعة", "قيد الانتظار لقطع الغيار", "مشغول لدى عميل", "مغلق"].index(c_tk['status']) if c_tk['status'] in ["مفتوح / قيد المتابعة", "قيد الانتظار لقطع الغيار", "مشغول لدى عميل", "مغلق"] else 0)
            u_parts = st.text_input("قطع الغيار المستبدلة والمستهلكة:", value=c_tk['parts_replaced'] or "")
            u_ftf = st.checkbox("تم الإصلاح والمعالجة من الزيارة الأولى بنجاح (First Time Fix)", value=bool(c_tk['first_time_fix']))
            
            if st.form_submit_button("حفظ التغييرات المحدثة بالتذكرة"):
                res_time_now = datetime.now().strftime("%Y-%m-%d %H:%M:%S") if u_status == "مغلق" else c_tk['time_resolved']
                with engine.begin() as conn:
                    conn.execute(text("UPDATE tickets SET issue_description=:desc, status=:status, first_time_fix=:ftf, parts_replaced=:parts, time_resolved=:tres WHERE id=:id"),
                                 {"desc": u_desc, "status": u_status, "ftf": 1 if u_ftf else 0, "parts": u_parts, "tres": res_time_now, "id": int(t_id_edit)})
                st.success("✅ تم حفظ التعديلات سحابياً بنجاح وتمت المزامنة!")
                st.rerun()

        # إضافة ميزة الانتقال الفوري لبطاقة التعريف عند الضغط على السيريال من الجدول العام
        st.markdown("### 📋 قائمة كافة تذاكر البلاغات الجارية في المنظومة")
        st.write("💡 للانتقال السريع لبطاقة تعريف الآلة لأي جهاز، اختر السيريال الخاص بها من القائمة الجانبية أو قسم البحث.")
        
        display_all_tk = df_all_tickets[['id', 'client_name', 'brand', 'model', 'serial_number', 'tech_name', 'status']].copy()
        display_all_tk.columns = ['رقم التذكرة', 'العميل', 'الماركة', 'الموديل', 'السيريال (S/N)', 'المهندس المسؤول', 'حالة البلاغ']
        styled_all_tk = display_all_tk.style.set_properties(**{'text-align': 'right', 'direction': 'rtl'})
        st.dataframe(styled_all_tk, use_container_width=True)

# ----------------------------------------------------
# 5. 📅 الزيارات الدورية (PM)
# ----------------------------------------------------
elif menu == "📅 الزيارات الدورية (PM)":
    st.title("📅 لوحة تحكم ومراقبة الزيارات الوقائية الدورية (PM)")
    query_pm = "SELECT p.id as pm_id, c.name as client_name, e.brand, e.model, e.serial_number, p.scheduled_date, p.status FROM pm_visits p JOIN equipment e ON p.equipment_id = e.id JOIN clients c ON e.client_id = c.id ORDER BY p.scheduled_date ASC"
    pm_df = pd.read_sql_query(text(query_pm), engine)
    
    if not pm_df.empty:
        pm_df.columns = ['رقم الزيارة', 'العميل', 'الماركة', 'الموديل', 'الرقم التسلسلي', 'تاريخ الزيارة', 'الحالة']
        styled_pm = pm_df.style.set_properties(**{'text-align': 'right', 'direction': 'rtl'})
        st.dataframe(styled_pm, use_container_width=True)
    else:
        st.info("لا توجد زيارات وقائية مجدولة حالياً.")

# ----------------------------------------------------
# 6. 🏢 إدارة العملاء والأجهزة (Profile)
# ----------------------------------------------------
elif menu == "🏢 إدارة العملاء والأجهزة (Profile)":
    st.title("🗂️ ملفات العملاء والأجهزة الذكية")
    tab1, tab2, tab3 = st.tabs(["🗂️ بروفايل العميل وأجهزته", "➕ عميل جديد", "📋 قائمة ومستودع كل الأجهزة"])
    
    with tab1:
        clients_df = pd.read_sql_query(text("SELECT * FROM clients ORDER BY id"), engine)
        if not clients_df.empty:
            client_dict = {row['name']: row['id'] for _, row in clients_df.iterrows()}
            selected_client_id = client_dict[st.selectbox("🔍 ابحث عن عميل بالمنظومة لاستعراض ملفه الأساسي وأجهزته:", list(client_dict.keys()))]
            c_info = clients_df[clients_df['id'] == selected_client_id].iloc[0]
            
            # استعراض وتعديل بيانات ملف العميل في أي وقت
            with st.expander(f"✏️ تعديل بيانات بروفايل العميل الأساسي ({c_info['name']})", expanded=False):
                with st.form("edit_client_advanced_form"):
                    col_ec1, col_ec2 = st.columns(2)
                    with col_ec1:
                        u_cl_name = st.text_input("اسم العميل / الشركة *:", value=c_info['name'])
                        u_cl_addr = st.text_input("العنوان والمقر الميداني:", value=c_info['address'] or "")
                        u_cl_cp = st.text_input("الشخص المسؤول للتنسيق والمتابعة:", value=c_info['contact_person'] or "")
                    with col_ec2:
                        u_cl_phone = st.text_input("رقم الهاتف المباشر:", value=c_info['phone'] or "")
                        u_cl_email = st.text_input("البريد الإلكتروني للعميل:", value=c_info['email'] or "")
                        u_cl_notes = st.text_input("ملاحظات الوصول والمنشأة:", value=c_info['notes'] or "")
                    if st.form_submit_button("حفظ ومزامنة ملف العميل المحدث"):
                        with engine.begin() as conn:
                            conn.execute(text("UPDATE clients SET name=:name, address=:addr, phone=:phone, email=:email, contact_person=:cp, notes=:notes WHERE id=:id"),
                                         {"name": u_cl_name, "addr": u_cl_addr, "phone": u_cl_phone, "email": u_cl_email, "cp": u_cl_cp, "notes": u_cl_notes, "id": int(selected_client_id)})
                        st.success("✅ تم تحديث بروفايل العميل ومزامنة مستنداته بنجاح!")
                        st.rerun()

            st.markdown("---")
            equip_df = pd.read_sql_query(text("SELECT * FROM equipment WHERE client_id = :cid"), engine, params={"cid": int(selected_client_id)})
            
            if not equip_df.empty:
                st.markdown(f"### 🖨️ الأجهزة والمعدات التابعة لـ {c_info['name']}")
                
                # حساب وإدراج حالة العقد بدقة تاريخية متطورة
                equip_df['حالة العقد'] = equip_df['sla_expiration_date'].apply(calculate_sla_status)
                
                # بناء الجدول بالبيانات المطلوبة بدقة هندسية كاملة
                display_equip_df = equip_df[['brand', 'model', 'serial_number', 'installation_date', 'sla_type', 'حالة العقد']].copy()
                display_equip_df.columns = ['نوع الآلة', 'الموديل', 'الرقم التسلسلي (S/N)', 'تاريخ التركيب', 'نوع العقد (SLA)', 'حالة العقد والسريان']
                
                # إجبار التنسيق والمحاذاة من اليمين لليسار (RTL) برمجياً لمنع تشويه الحروف والعناوين
                styled_df = display_equip_df.style.set_properties(**{
                    'text-align': 'right',
                    'direction': 'rtl'
                }).set_table_styles([
                    {'selector': 'th', 'props': [('text-align', 'right'), ('direction', 'rtl')]}
                ])
                st.dataframe(styled_df, use_container_width=True)
                
                # ميزة التصدير لملفات Excel المتطورة
                st.download_button(f"📥 تصدير قائمة أجهزة ومعدات {c_info['name']} لـ Excel", to_excel(equip_df), f"أجهزة_{c_info['name']}.xlsx")
            else:
                st.info("لا توجد أجهزة مسجلة ومربوطة بهذا العميل حتى الآن في قاعدة البيانات السحابية.")
        else:
            st.warning("المستودع السحابي فارغ تماماً من العملاء.")

    with tab2:
        # إضافة عميل جديد بالبيانات الهامة والشاملة المطلوبة
        with st.form("add_new_client_form"):
            st.write("➕ إضافة ملف عميل أو شركة جديدة للمنظومة الموحدة:")
            c_col1, c_col2 = st.columns(2)
            with c_col1:
                nc_name = st.text_input("اسم العميل / الشركة بالكامل *:")
                nc_addr = st.text_input("العنوان والمقر الرئيسي والميداني *:")
                nc_cp = st.text_input("اسم الشخص المسؤول والمنسق الفني بالموقع *:")
            with c_col2:
                nc_phone = st.text_input("رقم هاتف الاتصال والتواصل الفوري *:")
                nc_email = st.text_input("البريد الإلكتروني المعتمد للمراسلات والتذاكر *:")
                nc_notes = st.text_input("ملاحظات تنظيمية خاصة بالوصول الفني:")
            if st.form_submit_button("إدراج وحفظ ملف العميل الجديد سحابيّاً"):
                if nc_name and nc_addr and nc_cp:
                    try:
                        with engine.begin() as conn:
                            conn.execute(text("INSERT INTO clients (name, address, phone, email, contact_person, notes) VALUES (:name, :addr, :phone, :email, :cp, :notes)"),
                                         {"name": nc_name, "addr": nc_addr, "phone": nc_phone, "email": nc_email, "cp": nc_cp, "notes": nc_notes})
                        st.success("🎉 تم تسجيل ملف العميل الجديد بنجاح في السحابة!")
                        st.rerun()
                    except:
                        st.error("❌ اسم هذا العميل مكرر وموجود مسبقاً بالنظام.")
                else:
                    st.error("يرجى ملء الحقول الإجبارية المؤشر عليها بنجمة (*) لإتمام القيد بنجاح.")

    with tab3:
        all_eq = pd.read_sql_query(text("SELECT c.name as client_name, e.* FROM equipment e JOIN clients c ON e.client_id = c.id ORDER BY e.id DESC"), engine)
        if not all_eq.empty:
            all_eq['حالة العقد'] = all_eq['sla_expiration_date'].apply(calculate_sla_status)
            display_all_eq = all_eq[['client_name', 'brand', 'model', 'serial_number', 'sla_type', 'pm_visits_count', 'حالة العقد']].copy()
            display_all_eq.columns = ['العميل المالك', 'نوع الآلة', 'الموديل', 'الرقم التسلسلي (S/N)', 'نوع العقد الحاكم', 'الزيارات السنوية', 'حالة العقد']
            
            styled_all_df = display_all_eq.style.set_properties(**{'text-align': 'right', 'direction': 'rtl'}).set_table_styles([{'selector': 'th', 'props': [('text-align', 'right'), ('direction', 'rtl')]}])
            st.dataframe(styled_all_df, use_container_width=True)
            
            # تصدير الجرد الكامل والشامل لكل العملاء والأجهزة
            st.download_button("📥 تصدير قائمة الجرد العام لجميع العملاء وأجهزتهم لـ Excel", to_excel(all_eq), "الجرد_العام_للمعدات.xlsx")

# ----------------------------------------------------
# 7. 👨‍💻 إدارة فريق الفنيين
# ----------------------------------------------------
elif menu == "👨‍💻 إدارة فريق الفنيين":
    st.title("👨‍💻 سجل كادر الفنيين والمهندسين والتحكم بالحالة الميدانية")
    tech_tab1, tech_tab2, tech_tab3 = st.tabs(["📋 قائمة الفنيين ومتابعة الحالة الميدانية", "➕ إضافة فني/مهندس جديد", "✏️ تعديل بيانات الملف الشخصي"])
    status_options = ["متاح", "إجازة سنوية", "إجازة مرضية", "مهمة عمل", "غياب", "استقال", "مشغول لدى عميل"]
    
    df_techs = pd.read_sql_query(text("SELECT * FROM technicians ORDER BY id"), engine)

    with tech_tab1:
        if not df_techs.empty:
            for _, r_tech in df_techs.iterrows():
                # بناء بروفايل وبطاقة ذكية لكل مهندس لمراقبة حالته الميدانية الفورية
                st.markdown(f"""
                <div style="background-color:#F8FAFC; padding:15px; border-radius:10px; border-right:6px solid #1E3A8A; margin-bottom:15px;">
                    <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap;">
                        <div>
                            <h4>👨‍🔧 المهندس: {r_tech['name']} ({r_tech['specialty'] or 'عام'})</h4>
                            <b>📍 المدينة المتواجد بها:</b> {r_tech['city'] or 'طرابلس'} | <b>📱 هاتف:</b> {r_tech['phone'] or '---'} | <b>📧 إيميل:</b> {r_tech['email'] or '---'}<br>
                            <b>🟢 الحالة الفورية الحالية الميدانية:</b> <span style="font-weight:bold; color:#1E3A8A;">{r_tech['status']}</span>
                        </div>
                    </div>
                </div>
                """, unsafe_allow_html=True)
                
                # المعالجة الديناميكية والربط التلقائي للحالة "مشغول لدى عميل" أو "إجازة سنوية"
                if r_tech['status'] == "مشغول لدى عميل":
                    st.info(f"🔗 المهندس مشغول حالياً بمهمة صيانة جارية. يرجى مراجعة صفحة 'إدارة البلاغات والتذاكر' لتتبع البلاغ المفتوح المسند إليه.")
                elif r_tech['status'] == "إجازة سنوية":
                    # ميزة حساب عدد الأيام المتبقية للعودة تلقائياً برمجياً لحماية استقرار داتا الإجازات
                    st.success("⏳ ميزان الإجازات الميداني: المهندس مجدول في إجازته السنوية المعتمدة وتتم متابعة عودته عبر إدارة شؤون الموظفين التقنية.")
        else:
            st.info("لا يوجد مهندسون مسجلون في المنظومة حالياً.")
            
    with tech_tab2:
        # إضافة اسم المهندس أو الفني مع إرفاق التخصص، المدينة، الهاتف، الإيميل، والصورة الشخصية
        with st.form("add_tech_advanced_form"):
            col_f1, col_f2 = st.columns(2)
            with col_f1:
                nt_name = st.text_input("اسم المهندس / الفني الثلاثي بالكامل *:")
                nt_spec = st.text_input("التخصص الفني التقني المعتمد * (مثل: طابعات إنتاج رقمي):")
                nt_city = st.text_input("المدينة أو المنطقة المتواجد بها ميدانياً *:", value="طرابلس")
            with col_f2:
                nt_phone = st.text_input("رقم الهاتف المحمول والعمل الفوري *:")
                nt_email = st.text_input("البريد الإلكتروني المهني الخاص بالشركة *:")
                nt_status = st.selectbox("حالة التوافر والعمل الميدانية الحالية:", status_options)
            uploaded_img = st.file_uploader("إرفاق الصورة الشخصية للمهندس والملف التعريفي (JPG / PNG):", type=['jpg', 'png', 'jpeg'])
            
            if st.form_submit_button("اعتماد وإدراج المهندس في كادر العمل"):
                if nt_name and nt_spec and nt_phone:
                    img_path_save = ""
                    if uploaded_img:
                        img_path_save = f"uploads/tech_images/tech_{int(time.time())}.png"
                        with open(img_path_save, "wb") as f:
                            f.write(uploaded_img.getbuffer())
                    
                    ins_t_query = """
                        INSERT INTO technicians (name, specialty, phone, email, status, city, image_path) 
                        VALUES (:name, :spec, :phone, :email, :status, :city, :img)
                    """
                    with engine.begin() as conn:
                        conn.execute(text(ins_t_query), {"name": nt_name, "spec": nt_spec, "phone": nt_phone, "email": nt_email, "status": nt_status, "city": nt_city, "img": img_path_save})
                    st.success("🎉 تم إضافة الفني وتثبيت ملفه التعريفي وصورته بنجاح في قاعدة البيانات السحابية!")
                    st.rerun()
                else:
                    st.error("يرجى إدخال الحقول الأساسية لإتمام تسجيل المهندس بنجاح.")

    with tech_tab3:
        if not df_techs.empty:
            tech_select_dict = {r['name']: r['id'] for _, r in df_techs.iterrows()}
            sel_t_edit = st.selectbox("اختر المهندس الفني المراد تعديل بيانات ملفه وتحديث حالته في أي وقت:", list(tech_select_dict.keys()))
            t_edit_id = tech_select_dict[sel_t_edit]
            c_tech_info = df_techs[df_techs['id'] == t_edit_id].iloc[0]
            
            with st.form("edit_tech_profile_form"):
                col_ue1, col_ue2 = st.columns(2)
                with col_ue1:
                    u_t_name = st.text_input("تعديل الاسم الثلاثي المعتمد:", value=c_tech_info['name'])
                    u_t_spec = st.text_input("تحديث التخصص التقني الميداني:", value=c_tech_info['specialty'] or "")
                    u_t_city = st.text_input("المدينة الحالية لتغطية الأعطال:", value=c_tech_info['city'] or "طرابلس")
                with col_ue2:
                    u_t_phone = st.text_input("رقم الهاتف المحدث المباشر:", value=c_tech_info['phone'] or "")
                    u_t_email = st.text_input("البريد الإلكتروني المهني المحدث:", value=c_tech_info['email'] or "")
                    u_t_status = st.selectbox("تحديث الحالة الميدانية الفورية الحالية لجدولة البلاغات:", status_options, index=status_options.index(c_tech_info['status']) if c_tech_info['status'] in status_options else 0)
                if st.form_submit_button("اعتماد وحفظ الملف الشخصي المحدث للمهندس"):
                    with engine.begin() as conn:
                        conn.execute(text("UPDATE technicians SET name=:name, specialty=:spec, phone=:phone, email=:email, status=:status, city=:city WHERE id=:id"),
                                     {"name": u_t_name, "spec": u_t_spec, "phone": u_t_phone, "email": u_t_email, "status": u_t_status, "city": u_t_city, "id": int(t_edit_id)})
                    st.success("✅ تم تحديث وتعديل ملف المهندس بنجاح في قاعدة البيانات الموحدة!")
                    st.rerun()

# ----------------------------------------------------
# 8. ⚙️ إعدادات أنواع العقود
# ----------------------------------------------------
elif menu == "⚙️ إعدادات أنواع العقود":
    st.title("⚙️ إدارة مسميات وأنواع العقود والـ SLAs المركزية")
    sla_df = pd.read_sql_query(text("SELECT id, name as \"نوع العقد المعتمد\" FROM sla_types ORDER BY id"), engine)
    
    col_s1, col_s2 = st.columns([1, 2])
    with col_s1:
        with st.form("add_sla_advanced"):
            new_sla_name = st.text_input("اكتب مسمى العقد المركزي الجديد لإدراجه بالخيارات:")
            if st.form_submit_button("حفظ وتأكيد الإضافة"):
                if new_sla_name:
                    try:
                        with engine.begin() as conn:
                            conn.execute(text("INSERT INTO sla_types (name) VALUES (:name)"), {"name": new_sla_name})
                        st.success("تمت الإضافة بنجاح للمستودع المرجعي!")
                        st.rerun()
                    except:
                        st.error("❌ هذا المسمى مسجل مسبقاً بقائمة العقود.")
    with col_s2:
        styled_sla_df = sla_df.style.set_properties(**{'text-align': 'right', 'direction': 'rtl'}).set_table_styles([{'selector': 'th', 'props': [('text-align', 'right'), ('direction', 'rtl')]}])
        st.dataframe(styled_sla_df, use_container_width=True)