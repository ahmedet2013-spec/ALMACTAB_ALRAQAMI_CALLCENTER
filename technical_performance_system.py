import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime, timedelta
import os
import time
import io
from sqlalchemy import create_engine, text

# --- الأمان والسحابة ---
# قراءة الرابط السحابي بأمان من Streamlit Secrets بدلاً من الكود المباشر
DB_URL = st.secrets["DB_URL"]

def get_engine():
    return create_engine(DB_URL)

engine = get_engine()

# ضبط إعدادات الصفحة والتصميم الداعم لـ RTL بالكامل
st.set_page_config(page_title="منظومة إدارة الأداء الفني - شركة المكتب الرقمي", page_icon="🛠️", layout="wide")

st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Tajawal:wght@400;500;700&display=swap');
    
    /* تطبيق محاذاة RTL والخط العربي على التطبيق بالكامل */
    html, body, [data-testid="stSidebar"], .stApp, p, div, span, label {
        font-family: 'Tajawal', sans-serif;
        direction: RTL !important;
        text-align: right !important;
    }
    
    /* تعديل اتجاه القائمة الجانبية وعناصرها */
    [data-testid="stSidebarNav"] {
        direction: RTL !important;
        text-align: right !important;
    }
    
    .stHeadingContainer h1, .stHeadingContainer h2, .stHeadingContainer h3, .stHeadingContainer h4 {
        color: #1E3A8A;
        text-align: right !important;
    }
    
    /* تنسيق الجداول لتناسب القراءة العربية */
    th {
        background-color: #1E3A8A !important;
        color: white !important;
        text-align: right !important;
    }
    
    td {
        text-align: right !important;
    }
    
    /* بطاقات مؤشرات الأداء والتميز */
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

# --- نظام التحقق من الدخول ---
if 'logged_in' not in st.session_state: st.session_state['logged_in'] = False
if not st.session_state['logged_in']:
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        if os.path.exists("MAC logo resized.png"): st.image("MAC logo resized.png", use_container_width=True)
        st.markdown("<h2 style='text-align: center;'>تسجيل الدخول للمنظومة</h2>", unsafe_allow_html=True)
        with st.form("login_form"):
            username = st.text_input("اسم المستخدم")
            password = st.text_input("كلمة المرور", type="password")
            if st.form_submit_button("دخول"):
                if username == "Ahmed" and password == "admin123":
                    st.session_state['logged_in'] = True; st.rerun()
                else: st.error("❌ بيانات الدخول غير صحيحة.")
    st.stop()

# --- دالة تصدير البيانات إلى ملف Excel ---
def to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='التقرير المصدّر')
    return output.getvalue()

# --- دالة التصدير كـ CSV (مناسبة للطباعة الفورية والمحاذاة الرأسية) ---
def to_csv_printable(df):
    return df.to_csv(index=False).encode('utf-8-sig')

# --- محرك احتساب الفني المتميز والنجم السريع (KPIs Engine) ---
def get_star_technician():
    query = """
        SELECT t.*, tech.name as tech_name 
        FROM tickets t 
        JOIN technicians tech ON t.tech_id = tech.id
    """
    df_all = pd.read_sql_query(text(query), engine)
    if df_all.empty:
        return None, pd.DataFrame()
        
    df_all['time_reported'] = pd.to_datetime(df_all['time_reported'])
    df_all['time_resolved'] = pd.to_datetime(df_all['time_resolved'])
    df_all['time_arrived'] = pd.to_datetime(df_all['time_arrived'])
    
    # حساب فترات الصيانة بالساعات
    df_all['repair_time_hrs'] = (df_all['time_resolved'] - df_all['time_arrived']).dt.total_seconds() / 3600.0
    
    # تصفية الشهر الحالي للتحديث التلقائي
    current_month = datetime.now().strftime("%Y-%m")
    df_current = df_all[df_all['time_reported'].dt.strftime("%Y-%m") == current_month]
    
    if df_current.empty:
        return None, pd.DataFrame()
        
    # تجميع مؤشرات الأداء لكل فني للشهر الحالي
    tech_stats = df_current.groupby('tech_name').agg(
        total_visits=('id', 'count'),
        avg_repair_time=('repair_time_hrs', 'mean'),
        first_time_fixes=('first_time_fix', 'sum')
    ).reset_index()
    
    # معادلة التقييم التلقائي (الوزن النسبي لسرعة الصيانة، عدد الإغلاقات، والإصلاح الفوري)
    # نقوم بعمل ترتيب عكسي لوقت الصيانة (الأقل هو الأفضل)
    tech_stats['score'] = (tech_stats['total_visits'] * 3.0) + (tech_stats['first_time_fixes'] * 4.0) - (tech_stats['avg_repair_time'].fillna(24) * 1.5)
    
    if tech_stats.empty:
        return None, pd.DataFrame()
        
    winner = tech_stats.sort_values(by='score', ascending=False).iloc[0]['tech_name']
    
    # حساب كم مرة كان الفني مثالي تاريخياً (على مر الشهور)
    df_all['year_month'] = df_all['time_reported'].dt.strftime("%Y-%m")
    historical_winners = []
    
    for month, group in df_all.groupby('year_month'):
        g_stats = group.groupby('tech_name').agg(
            total_v=('id', 'count'), avg_r=('repair_time_hrs', 'mean'), ftf=('first_time_fix', 'sum')
        ).reset_index()
        g_stats['score'] = (g_stats['total_v'] * 3.0) + (g_stats['ftf'] * 4.0) - (g_stats['avg_r'].fillna(24) * 1.5)
        if not g_stats.empty:
            historical_winners.append(g_stats.sort_values(by='score', ascending=False).iloc[0]['tech_name'])
            
    # توليد جدول التكرار التاريخي
    history_df = pd.DataFrame(historical_winners, columns=['tech_name']).value_counts().reset_index(name='star_count')
    
    return winner, history_df
# ====================================================
# الجزء الثاني: القائمة الجانبية وصفحة لوحة التحكم والأداء
# ====================================================

# إعداد القائمة الجانبية بشكل متوافق تماماً مع التنسيق اليميني (RTL)
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

# بطاقة تعريف المطور في القائمة الجانبية
st.sidebar.markdown("""
<div class="developer-card">
    <h4>⚙️ تصميم وتطوير المنظومة:</h4>
    <b>م. أحمد عثمان</b><br>
    مدير إدارة الدعم الفني وتقنية المعلومات<br>
    📧 AhmedE@almactab.com<br>
    📱 0923009907
</div>
""", unsafe_allow_html=True)

# ----------------------------------------------------
# 📊 لوحة التحكم والأداء الشهري
# ----------------------------------------------------
if menu == "📊 لوحة التحكم والأداء الشهري":
    st.title("📊 الملخص الشهري وتحليل الأداء الفني")
    
    # 🌟 حساب وعرض الفني المتميز (الموظف المثالي) تلقائياً
    winner_name, star_history_df = get_star_technician()
    
    if winner_name:
        # حساب كم نجمة تاريخياً لهذا الفني
        stars_count = 0
        if not star_history_df.empty and winner_name in star_history_df['tech_name'].values:
            stars_count = star_history_df.loc[star_history_df['tech_name'] == winner_name, 'star_count'].values[0]
            
        st.markdown(f"""
        <div class="star-card">
            <h3 style='margin:0; color:#B45309;'>🌟 الموظف المثالي لهذا الشهر الحالي: {winner_name} ⭐</h3>
            <p style='margin:5px 0 0 0; font-size:16px; color:#4B5563;'>
                تم اختيار المهندس <b>{winner_name}</b> تلقائياً بناءً على سرعة الاستجابة، عدد الإصلاحات، ومعدل الإصلاح من أول زيارة (First Time Fix).
                <br>🎖️ <b>عدد مرات الحصول على اللقب تاريخياً:</b> {stars_count} مرة.
            </p>
        </div>
        """, unsafe_allow_html=True)
    else:
        st.info("💡 سيظهر الفني المثالي هنا تلقائياً بمجرد إدخال بلاغات صيانة وإغلاقها للشهر الحالي.")

    # جلب بيانات البلاغات الكاملة للعرض والرسوم البيانية
    query_tickets = """
        SELECT t.*, c.name as client_name, e.brand, e.model, tech.name as tech_name
        FROM tickets t 
        JOIN equipment e ON t.equipment_id = e.id
        JOIN clients c ON e.client_id = c.id 
        JOIN technicians tech ON t.tech_id = tech.id
        ORDER BY t.id DESC
    """
    df_tickets = pd.read_sql_query(text(query_tickets), engine)
    
    if not df_tickets.empty:
        df_tickets['time_reported'] = pd.to_datetime(df_tickets['time_reported'])
        df_tickets['time_arrived'] = pd.to_datetime(df_tickets['time_arrived'])
        df_tickets['time_resolved'] = pd.to_datetime(df_tickets['time_resolved'])
        
        # حساب أوقات الاستجابة والإصلاح بالساعات
        df_tickets['response_time_hrs'] = (df_tickets['time_arrived'] - df_tickets['time_reported']).dt.total_seconds() / 3600.0
        df_tickets['repair_time_hrs'] = (df_tickets['time_resolved'] - df_tickets['time_arrived']).dt.total_seconds() / 3600.0
        df_tickets['month_year'] = df_tickets['time_reported'].dt.to_period('M').astype(str)
        
        months_available = sorted(df_tickets['month_year'].dropna().unique(), reverse=True)
        
        # فلترة البيانات حسب الشهر المختار من قبل المستخدم
        selected_month = st.selectbox("📅 اختر الشهر المالي للمراقبة:", months_available)
        df_filtered = df_tickets[df_tickets['month_year'] == selected_month]
        
        # عرض المؤشرات الرقمية السريعة (Metrics)
        m_col1, m_col2, m_col3, m_col4 = st.columns(4)
        m_col1.metric("📥 إجمالي البلاغات المسجلة", len(df_filtered))
        m_col2.metric("🔒 البلاغات المغلقة بنجاح", len(df_filtered[df_filtered['status'] == "مغلق"]))
        m_col3.metric("⏱️ متوسط الاستجابة (ساعة)", f"{df_filtered['response_time_hrs'].mean():.1f}" if not df_filtered['response_time_hrs'].isna().all() else "---")
        m_col4.metric("⚙️ متوسط الإصلاح (ساعة)", f"{df_filtered['repair_time_hrs'].mean():.1f}" if not df_filtered['repair_time_hrs'].isna().all() else "---")
        
        st.markdown("### 📋 سجل البلاغات التفصيلي للشهر المختار")
        
        # تجهيز جدول العرض باللغة العربية وميزة النجمة بجانب اسم الفني المثالي الحالي
        display_df = df_filtered[['id', 'client_name', 'brand', 'model', 'tech_name', 'status']].copy()
        if winner_name:
            display_df['tech_name'] = display_df['tech_name'].apply(lambda x: f"{x} ⭐" if x == winner_name else x)
            
        display_df.columns = ['رقم البلاغ', 'اسم العميل', 'الماركة', 'الموديل', 'المهندس المختص', 'حالة البلاغ']
        st.dataframe(display_df, use_container_width=True)
        
        # --- أدوات التصدير والطباعة أسفل الجدول مباشرة ---
        exp_col1, exp_col2, exp_col3 = st.columns([1, 1, 2])
        with exp_col1:
            st.download_button(
                label="📥 تصدير الجدول إلى Excel",
                data=to_excel(df_filtered),
                file_name=f"تقرير_الأداء_{selected_month}.xlsx",
                mime="application/vnd.ms-excel"
            )
        with exp_col2:
            st.download_button(
                label="🖨️ تجهيز ملف للطباعة الفورية (CSV)",
                data=to_csv_printable(df_filtered),
                file_name=f"طباعة_تقرير_{selected_month}.csv",
                mime="text/csv"
            )
            
        # --- قسم الرسوم البيانية المتطورة وتحليل مستوى الأداء ---
        st.markdown("---")
        st.markdown("### 📊 الرسوم البيانية المقارنة وتحليل الأداء")
        
        g_col1, g_col2 = st.columns(2)
        
        with g_col1:
            st.subheader("🏆 السجل التاريخي لعدد نجوم التميز")
            if not star_history_df.empty:
                fig_stars = px.bar(
                    star_history_df, 
                    x='tech_name', 
                    y='star_count',
                    labels={'tech_name': 'اسم الفني', 'star_count': 'مرات التميز'},
                    title="عدد مرات الحصول على لقب الموظف المثالي تاريخياً",
                    color='star_count',
                    color_continuous_scale=px.colors.sequential.YlOrBr
                )
                fig_stars.update_layout(xaxis_title="الفني", yaxis_title="عدد النجوم")
                st.plotly_chart(fig_stars, use_container_width=True)
            else:
                st.info("لا توجد بيانات كافية لرسم بياني النجوم التاريخي بعد.")
                
        with g_col2:
            st.subheader("📈 مستوى الأداء الإجمالي للفنيين (الشهر الحالي)")
            # حساب وزن نسبي سريع للعرض البياني لمستويات الأداء الحالية
            current_stats = df_filtered.groupby('tech_name').agg(
                total_v=('id', 'count'), ftf=('first_time_fix', 'sum')
            ).reset_index()
            current_stats['مستوى الأداء التقديري'] = (current_stats['total_v'] * 10) + (current_stats['ftf'] * 15)
            
            if not current_stats.empty:
                fig_perf = px.line(
                    current_stats, 
                    x='tech_name', 
                    y='مستوى الأداء التقديري',
                    markers=True,
                    title="منحنى الأداء التنافسي للفريق خلال الشهر",
                )
                fig_perf.update_traces(line_color='#1E3A8A', lw=3)
                st.plotly_chart(fig_perf, use_container_width=True)
            else:
                st.info("لا توجد بلاغات كافية لحساب منحنى الأداء.")
    else:
        st.info("📥 لا توجد أي بلاغات صيانة مسجلة في قاعدة البيانات السحابية حتى الآن.")
# ====================================================
# الجزء الثالث: صفحات البحث، تسجيل البلاغات، وإدارتها
# ====================================================

# ----------------------------------------------------
# 🔍 البحث الشامل عن جهاز (S/N)
# ----------------------------------------------------
elif menu == "🔍 البحث الشامل عن جهاز (S/N)":
    st.title("🔍 البحث الشامل (بروفايل الجهاز الذكي)")
    st.write("أدخل الرقم التسلسلي للجهاز لاستخراج بطاقته الذكية ومعرفة كافة تفاصيله وحالته وتاريخه.")
    
    search_sn = st.text_input("أدخل الرقم التسلسلي (S/N) للبحث:", placeholder="مثال: XER8145001").strip()
    
    # تجهيز قائمة العقود لاستخدامها في نموذج الإضافة السريع بالأسفل
    sla_df_options = pd.read_sql_query(text("SELECT name FROM sla_types ORDER BY id"), engine)
    sla_list_db = sla_df_options['name'].tolist() if not sla_df_options.empty else ["عقد افتراضي"]
    extended_sla_list = sla_list_db + ["+ إضافة نوع عقد جديد يدوياً..."]

    if search_sn:
        query_eq = """
            SELECT e.*, c.name as client_name, c.phone as client_phone, c.address as client_address 
            FROM equipment e 
            JOIN clients c ON e.client_id = c.id 
            WHERE e.serial_number = :sn
        """
        eq_df = pd.read_sql_query(text(query_eq), engine, params={"sn": search_sn})
        
        if not eq_df.empty:
            eq = eq_df.iloc[0]
            eq_id = eq['id']
            
            query_tk = """
                SELECT t.*, tech.name as tech_name 
                FROM tickets t 
                JOIN technicians tech ON t.tech_id = tech.id 
                WHERE t.equipment_id = :eq_id 
                ORDER BY t.time_reported DESC
            """
            tickets_df = pd.read_sql_query(text(query_tk), engine, params={"eq_id": int(eq_id)})
            
            open_tickets = tickets_df[tickets_df['status'] != 'مغلق']
            status_text = "🟢 يعمل (لا توجد بلاغات مفتوحة)" if open_tickets.empty else "🔴 معطل / قيد الصيانة"
            
            maintenance_count = len(tickets_df)
            last_tech = tickets_df.iloc[0]['tech_name'] if maintenance_count > 0 else "لم تتم صيانته بعد"
            last_visit = str(tickets_df.iloc[0]['time_reported']) if maintenance_count > 0 else "---"
            
            st.markdown(f"""
            <div class="device-profile-card">
                <h3>🖥️ بطاقة تعريف الجهاز: {eq['brand']} {eq['model']}</h3>
                <hr>
                <div style="display: flex; justify-content: space-between; flex-wrap: wrap;">
                    <div style="width: 48%; margin-bottom: 15px;">
                        <b>🏢 مملوك للعميل:</b> {eq['client_name']}<br>
                        <b>📍 موقع الجهاز:</b> {eq['location_building'] or 'غير محدد'} - {eq['location_department'] or 'غير محدد'}<br>
                        <b>🏷️ الرقم التسلسلي:</b> {eq['serial_number']}<br>
                        <b>📅 تاريخ التركيب:</b> {eq['installation_date']}<br>
                    </div>
                    <div style="width: 48%; margin-bottom: 15px;">
                        <b>⚙️ حالة الجهاز الحالية:</b> {status_text}<br>
                        <b>👨‍🔧 عدد مرات الصيانة:</b> {maintenance_count} مرة<br>
                        <b>🛠️ آخر مهندس زار الجهاز:</b> {last_tech} (بتاريخ: {last_visit})<br>
                        <b>💼 مصدر الشراء:</b> {eq['purchased_from_us']} {f"({eq['third_party_vendor']})" if eq['third_party_vendor'] else ""}<br>
                    </div>
                </div>
                <hr>
                <div style="text-align: center;">
                    <b>📜 نوع العقد:</b> {eq['sla_type']} &nbsp;|&nbsp; 
                    <b>⏳ صلاحية العقد:</b> {calculate_sla_status(eq['sla_expiration_date'])}
                </div>
            </div>
            """, unsafe_allow_html=True)
            
            if maintenance_count > 0:
                st.subheader("🛠️ السجل التاريخي لصيانة الجهاز")
                display_hist = tickets_df[['id', 'tech_name', 'issue_description', 'time_reported', 'status', 'parts_replaced']].copy()
                display_hist.columns = ['رقم البلاغ', 'اسم المهندس', 'العطل المشتكى منه', 'تاريخ البلاغ', 'الحالة', 'القطع المستبدلة']
                st.dataframe(display_hist, use_container_width=True)
                
                # تصدير سجل صيانة الجهاز المحدد
                ex_c1, ex_c2 = st.columns(2)
                with ex_c1:
                    st.download_button("📥 تصدير سجل صيانة الجهاز لـ Excel", to_excel(tickets_df), f"سجل_جهاز_{search_sn}.xlsx", "application/vnd.ms-excel")
                with ex_c2:
                    st.download_button("🖨️ تجهيز طباعة السجل (CSV)", to_csv_printable(tickets_df), f"طباعة_سجل_{search_sn}.csv", "text/csv")
        else:
            st.error("⚠️ لم يتم العثور على هذا الرقم التسلسلي في النظام. يمكنك إضافته كجهاز جديد بالأسفل:")
            
            with st.expander("➕ إضافة هذا الجهاز الجديد للمنظومة الآن", expanded=True):
                clients_list = pd.read_sql_query(text("SELECT id, name FROM clients"), engine)
                if not clients_list.empty:
                    client_dict = {row['name']: row['id'] for _, row in clients_list.iterrows()}
                    with st.form("add_new_searched_eq"):
                        sel_client = st.selectbox("اختر العميل المالك للجهاز *:", list(client_dict.keys()))
                        c1, c2 = st.columns(2)
                        with c1:
                            brand = st.selectbox("الماركة *:", ["Xerox", "Canon", "Konica Minolta", "OKI", "Kardex Remstar"])
                            model = st.text_input("الموديل *:")
                            s_num = st.text_input("الرقم التسلسلي *:", value=search_sn)
                            purchased_from = st.selectbox("هل تم شراء الجهاز منا؟ *", ["نعم (من المكتب الرقمي)", "لا (من مورد آخر)"])
                            vendor_name = st.text_input("اسم المورد الآخر (إن وجد):")
                        with c2:
                            building = st.text_input("المبنى (اختياري):")
                            dept = st.text_input("القسم (اختياري):")
                            pm_count = st.number_input("الزيارات الدورية سنوياً:", min_value=0, max_value=12, value=0)
                            sla_choice_new = st.selectbox("نوع العقد:", extended_sla_list)
                            final_sla_new = st.text_input("اكتب نوع العقد الجديد:") if sla_choice_new == "+ إضافة نوع عقد جديد يدوياً..." else sla_choice_new
                        
                        ins_date = st.date_input("تاريخ التركيب:")
                        exp_date = st.date_input("تاريخ انتهاء العقد:", datetime.now().date() + timedelta(days=365))
                        
                        if st.form_submit_button("حفظ وإضافة الجهاز"):
                            if model and s_num and final_sla_new:
                                if final_sla_new not in sla_list_db and sla_choice_new == "+ إضافة نوع عقد جديد يدوياً...":
                                    with engine.begin() as conn_sla:
                                        conn_sla.execute(text("INSERT INTO sla_types (name) VALUES (:name) ON CONFLICT DO NOTHING"), {"name": final_sla_new})
                                try:
                                    query_ins = """
                                        INSERT INTO equipment (client_id, brand, model, serial_number, installation_date, sla_type, sla_expiration_date, 
                                        location_building, location_department, pm_visits_count, purchased_from_us, third_party_vendor) 
                                        VALUES (:client_id, :brand, :model, :s_num, :ins_date, :sla, :exp_date, :building, :dept, :pm_count, :purchased, :vendor)
                                    """
                                    with engine.begin() as conn_eq:
                                        conn_eq.execute(text(query_ins), {
                                            "client_id": client_dict[sel_client], "brand": brand, "model": model, "s_num": s_num,
                                            "ins_date": str(ins_date), "sla": final_sla_new, "exp_date": str(exp_date),
                                            "building": building, "dept": dept, "pm_count": int(pm_count), "purchased": purchased_from, "vendor": vendor_name
                                        })
                                    st.success("✅ تم إضافة الجهاز بنجاح! أعد البحث عنه لعرض بطاقته.")
                                except IntegrityError:
                                    st.error("❌ الرقم التسلسلي مكرر لجهاز آخر بالفعل!")
                            else:
                                st.warning("الرجاء ملء الحقول الإجبارية.")
                else:
                    st.warning("يجب إضافة عملاء في المنظومة أولاً.")

# ----------------------------------------------------
# ➕ تسجيل بلاغ صيانة جديد
# ----------------------------------------------------
elif menu == "➕ تسجيل بلاغ صيانة جديد":
    st.title("➕ تسجيل بلاغ صيانة جديد")
    
    eq_rows = pd.read_sql_query(text("SELECT e.id, c.name as client_name, e.brand, e.model, e.serial_number FROM equipment e JOIN clients c ON e.client_id = c.id"), engine)
    tech_rows = pd.read_sql_query(text("SELECT id, name FROM technicians WHERE status='متاح'"), engine)
    
    if eq_rows.empty or tech_rows.empty:
        st.error("⚠️ يجب التأكد من تسجيل أجهزة في النظام وتوافر فنيين بحالة 'متاح' أولاً لتتمكن من فتح تذكرة صيانة.")
    else:
        equip_options = {f"{r['client_name']} - {r['brand']} {r['model']} (S/N: {r['serial_number']})": r['id'] for _, r in eq_rows.iterrows()}
        tech_options = {r['name']: r['id'] for _, r in tech_rows.iterrows()}
        
        with st.form("add_ticket_form"):
            selected_equip = st.selectbox("الجهاز المتضرر:", list(equip_options.keys()))
            selected_tech = st.selectbox("المهندس/الفني المختص والمتاح حالياً:", list(tech_options.keys()))
            issue_desc = st.text_area("وصف دقيق ومفصل للعطل المشكو منه:")
            col_t1, col_t2 = st.columns(2)
            with col_t1: date_rep = st.date_input("تاريخ فتح البلاغ:", datetime.now())
            with col_t2: time_rep = st.time_input("وقت الإبلاغ التقديري:", datetime.now().time())
            status = st.selectbox("حالة البلاغ المبدئية:", ["مفتوح / قيد المتابعة", "قيد الانتظار لقطع الغيار"])
            
            if st.form_submit_button("حفظ وفتح البلاغ رسميّاً"):
                full_rep = datetime.combine(date_rep, time_rep).strftime("%Y-%m-%d %H:%M:%S")
                query_add_ticket = """
                    INSERT INTO tickets (equipment_id, tech_id, issue_description, time_reported, status, first_time_fix) 
                    VALUES (:eq_id, :tech_id, :desc, :time_rep, :status, 0)
                """
                with engine.begin() as conn_tk:
                    conn_tk.execute(text(query_add_ticket), {
                        "eq_id": int(equip_options[selected_equip]), "tech_id": int(tech_options[selected_tech]),
                        "desc": issue_desc, "time_rep": full_rep, "status": status
                    })
                st.success("🎉 تم تسجيل بلاغ الصيانة وإسناده للمهندس بنجاح!")

# ----------------------------------------------------
# 🖥️ إدارة البلاغات والتذاكر
# ----------------------------------------------------
elif menu == "🖥️ إدارة البلاغات والتذاكر":
    st.title("🖥️ الإدارة الشاملة وتعديل تذاكر الصيانة")
    
    query_all_tk = """
        SELECT t.*, c.name as client_name, e.brand, e.model, tech.name as tech_name
        FROM tickets t 
        JOIN equipment e ON t.equipment_id = e.id
        JOIN clients c ON e.client_id = c.id 
        JOIN technicians tech ON t.tech_id = tech.id 
        ORDER BY t.id DESC
    """
    df_all_tickets = pd.read_sql_query(text(query_all_tk), engine)
    
    if df_all_tickets.empty:
        st.info("لا توجد أي بلاغات مسجلة في المنظومة حالياً.")
    else:
        ticket_list = {f"بلاغ رقم {r['id']} - لعميل: {r['client_name']} ({r['status']})": r['id'] for _, r in df_all_tickets.iterrows()}
        selected_ticket_label = st.selectbox("🔍 اختر بلاغ الصيانة المراد تحديثه أو إغلاقه:", list(ticket_list.keys()))
        t_id = ticket_list[selected_ticket_label]
        
        current_ticket = df_all_tickets[df_all_tickets['id'] == t_id].iloc[0]
        
        if current_ticket['report_file'] and os.path.exists(str(current_ticket['report_file'])):
            with open(str(current_ticket['report_file']), "rb") as f:
                st.download_button(label="📥 تحميل تقرير الصيانة المرفق سابقاً", data=f, file_name=os.path.basename(str(current_ticket['report_file'])))
        
        with st.form("edit_ticket_form"):
            st.info(f"📝 **وصف العطل الأصلي المقيد:** {current_ticket['issue_description']}")
            status_list = ["مفتوح / قيد المتابعة", "قيد الانتظار لقطع الغيار", "مغلق"]
            
            cur_status = current_ticket['status']
            def_idx = status_list.index(cur_status) if cur_status in status_list else 0
            
            new_status = st.selectbox("تحديث حالة التذكرة الحالية:", status_list, index=def_idx)
            
            col_r1, col_r2 = st.columns(2)
            with col_r1: date_res = st.date_input("تاريخ المعالجة/الإغلاق الفعلي:", datetime.now())
            with col_r2: time_res = st.time_input("ساعة الإغلاق والإنهاء الفعلي:", datetime.now().time())
            
            ftfr_new = st.checkbox("⚙️ تم الإصلاح بنجاح من الزيارة الأولى (First Time Fix)", value=bool(current_ticket['first_time_fix']))
            parts_replaced = st.text_input("قطع الغيار المستبدلة (إن وجدت):", value=current_ticket['parts_replaced'] or "")
            uploaded_report = st.file_uploader("إرفاق تقرير المهندس الموقع والجاهز (PDF / الصور)", type=['pdf', 'jpg', 'png'])
            
            if st.form_submit_button("حفظ وتحديث بيانات التذكرة"):
                res_time = datetime.combine(date_res, time_res).strftime("%Y-%m-%d %H:%M:%S") if new_status == "مغلق" else None
                report_path = current_ticket['report_file']
                
                if uploaded_report:
                    report_path = save_uploaded_file(uploaded_report, "uploads/reports", f"ticket_{t_id}")
                
                query_update_ticket = """
                    UPDATE tickets 
                    SET status = :status, time_resolved = :time_res, first_time_fix = :ftf, parts_replaced = :parts, report_file = :report 
                    WHERE id = :id
                """
                with engine.begin() as conn_up_tk:
                    conn_up_tk.execute(text(query_update_ticket), {
                        "status": new_status, "time_res": res_time, "ftf": 1 if ftfr_new else 0,
                        "parts": parts_replaced, "report": report_path, "id": int(t_id)
                    })
                st.success("✅ تم تحديث بيانات التذكرة السحابية بنجاح!")
                time.sleep(0.5)
                st.rerun()
                
        st.markdown("### 📋 قائمة كافة التذاكر الحالية وتصديرها")
        display_all_tk = df_all_tickets[['id', 'client_name', 'brand', 'model', 'tech_name', 'status', 'time_reported']].copy()
        display_all_tk.columns = ['رقم التذكرة', 'العميل', 'الماركة', 'الموديل', 'المهندس المختص', 'الحالة', 'تاريخ التبليغ']
        st.dataframe(display_all_tk, use_container_width=True)
        
        ex_col_tk1, ex_col_tk2 = st.columns(2)
        with ex_col_tk1:
            st.download_button("📥 تصدير كل التذاكر لـ Excel", to_excel(df_all_tickets), "سجل_كافة_البلاغات.xlsx", "application/vnd.ms-excel")
        with ex_col_tk2:
            st.download_button("🖨️ طباعة نموذج البلاغات العام (CSV)", to_csv_printable(df_all_tickets), "طباعة_البلاغات.csv", "text/csv")
# ====================================================
# الجزء الرابع والأخير: الزيارات الوقائية، العملاء، الفنيين، والعقود
# ====================================================

# ----------------------------------------------------
# 📅 الزيارات الدورية (PM)
# ----------------------------------------------------
elif menu == "📅 الزيارات الدورية (PM)":
    st.title("📅 لوحة تحكم الزيارات الوقائية الدورية (PM)")
    pm_tab1, pm_tab2, pm_tab3 = st.tabs(["📊 المراقبة والجدول الزمني", "➕ جدولة تلقائية لجهاز", "✅ تحديث حالة زيارة"])
    
    with pm_tab1:
        query_pm = """
            SELECT p.id as pm_id, c.name as client_name, e.brand, e.model, e.serial_number, e.sla_type, 
                   p.scheduled_date, p.status, p.notes
            FROM pm_visits p 
            JOIN equipment e ON p.equipment_id = e.id
            JOIN clients c ON e.client_id = c.id
            ORDER BY p.scheduled_date ASC
        """
        pm_df = pd.read_sql_query(text(query_pm), engine)
        
        if not pm_df.empty:
            def color_status(val):
                color = 'green' if val == 'مكتملة' else 'orange' if val == 'مجدولة' else 'red'
                return f'color: {color}; font-weight: bold;'
                
            st.dataframe(pm_df[['client_name', 'brand', 'model', 'serial_number', 'sla_type', 'scheduled_date', 'status', 'notes']].rename(columns={
                'client_name': 'العميل', 'brand': 'الماركة', 'model': 'الموديل', 'serial_number': 'الرقم التسلسلي',
                'sla_type': 'نوع العقد', 'scheduled_date': 'تاريخ الزيارة', 'status': 'الحالة', 'notes': 'ملاحظات'
            }).style.map(color_status, subset=['الحالة']), use_container_width=True)
            
            # أزرار تصدير جدول الزيارات الوقائية
            st.download_button("📥 تصدير جدول الزيارات الوقائية لـ Excel", to_excel(pm_df), "جدول_الزيارات_الوقائية.xlsx", "application/vnd.ms-excel")
        else:
            st.info("لا توجد زيارات دورية مجدولة حالياً في المنظومة.")

    with pm_tab2:
        st.subheader("توليد التواريخ تلقائياً للزيارات الدورية السنوية")
        eq_data = pd.read_sql_query(text("SELECT e.id, c.name as client_name, e.brand, e.model, e.serial_number, e.installation_date, e.pm_visits_count FROM equipment e JOIN clients c ON e.client_id = c.id WHERE e.pm_visits_count > 0"), engine)
        
        if not eq_data.empty:
            eq_dict = {f"{r['client_name']} - {r['brand']} {r['model']} (S/N: {r['serial_number']}) - {r['pm_visits_count']} زيارة": r['id'] for _, r in eq_data.iterrows()}
            selected_eq_pm = st.selectbox("اختر الجهاز لجدولة زياراته الوقائية:", list(eq_dict.keys()))
            eq_id_pm = eq_dict[selected_eq_pm]
            c_eq = eq_data[eq_data['id'] == eq_id_pm].iloc[0]
            
            try:
                def_date = datetime.strptime(str(c_eq['installation_date']), "%Y-%m-%d").date()
            except:
                def_date = datetime.now().date()
                
            start_date_input = st.date_input("تاريخ بداية الجدولة التلقائية:", def_date)
            
            if st.button("🔄 جدولة وتوزيع الزيارات تلقائياً على السحابة"):
                visits_count = int(c_eq['pm_visits_count'])
                days_interval = 365 / visits_count
                
                with engine.begin() as conn_pm:
                    for i in range(1, visits_count + 1):
                        v_date = start_date_input + timedelta(days=int(days_interval * i))
                        # التحقق من عدم تكرار الزيارة لنفس اليوم والجهاز
                        chk = conn_pm.execute(text("SELECT id FROM pm_visits WHERE equipment_id = :eid AND scheduled_date = :sdate"), {"eid": int(eq_id_pm), "sdate": str(v_date)}).fetchone()
                        if not chk:
                            conn_pm.execute(text("INSERT INTO pm_visits (equipment_id, scheduled_date, status) VALUES (:eid, :sdate, 'مجدولة')"), {"eid": int(eq_id_pm), "sdate": str(v_date)})
                st.success(f"✅ تمت جدولة وتوزيع {visits_count} زيارات بنجاح على قاعدة البيانات!")
                time.sleep(0.5)
                st.rerun()
        else:
            st.warning("لا توجد أجهزة مبرمجة تتطلب زيارات دورية حالياً (تأكد من تعديل الأجهزة ورفع خانة 'الزيارات الدورية سنوياً' عن الرقم 0).")

    with pm_tab3:
        if not pm_df.empty:
            pending_pm = pm_df[pm_df['status'] != 'مكتملة']
            if not pending_pm.empty:
                pm_list = {f"{r['client_name']} - {r['brand']} {r['model']} - ميعاد: {r['scheduled_date']}": r['pm_id'] for _, r in pending_pm.iterrows()}
                pm_id_to_update = pm_list[st.selectbox("اختر الزيارة الوقائية لإغلاقها وتأكيد تنفيذها:", list(pm_list.keys()))]
                
                with st.form("close_pm_form"):
                    comp_date = st.date_input("تاريخ التنفيذ الفعلي على أرض الواقع:", datetime.now().date())
                    pm_status_new = st.selectbox("حالة الزيارة الجديدة بعد المرور التقني:", ["مكتملة", "ملغاة"])
                    pm_notes = st.text_area("ملاحظات الفني وتقرير الصيانة الوقائية الحالية:")
                    
                    if st.form_submit_button("حفظ وإغلاق ملف الزيارة الدوري"):
                        query_up_pm = "UPDATE pm_visits SET status = :status, completed_date = :cdate, notes = :notes WHERE id = :id"
                        with engine.begin() as conn_up_pm:
                            conn_up_pm.execute(text(query_up_pm), {"status": pm_status_new, "cdate": str(comp_date), "notes": pm_notes, "id": int(pm_id_to_update)})
                        st.success("✅ تم إغلاق الزيارة الوقائية وتثبيتها في السجل السحابي الموحد!")
                        time.sleep(0.5)
                        st.rerun()
            else:
                st.info("جميع الزيارات الوقائية المقيدة مكتملة ومغلقة تماماً.")

# ----------------------------------------------------
# 🏢 إدارة العملاء والأجهزة (Profile)
# ----------------------------------------------------
elif menu == "🏢 إدارة العملاء والأجهزة (Profile)":
    st.title("🗂️ ملفات العملاء والأجهزة الذكية")
    tab1, tab2, tab3 = st.tabs(["🗂️ بروفايل العميل وأجهزته", "➕ عميل جديد", "📋 قائمة ومستودع كل الأجهزة"])
    
    # قائمة العقود المركزية لملء النماذج بالأسفل
    sla_df_opts = pd.read_sql_query(text("SELECT name FROM sla_types ORDER BY id"), engine)
    sla_list_db = sla_df_opts['name'].tolist() if not sla_df_opts.empty else ["عقد افتراضي"]
    extended_sla_list = sla_list_db + ["+ إضافة نوع عقد جديد يدوياً..."]

    with tab1:
        clients_df = pd.read_sql_query(text("SELECT * FROM clients ORDER BY id"), engine)
        if not clients_df.empty:
            client_dict = {row['name']: row['id'] for _, row in clients_df.iterrows()}
            selected_client_id = client_dict[st.selectbox("🔍 ابحث عن عميل بالمنظومة لاستعراض ملفه:", list(client_dict.keys()))]
            c_info = clients_df[clients_df['id'] == selected_client_id].iloc[0]
            
            with st.expander(f"✏️ تعديل بيانات الملف الأساسي للعميل ({c_info['name']})", expanded=False):
                with st.form("edit_client_form"):
                    col_c1, col_c2 = st.columns(2)
                    with col_c1:
                        new_name = st.text_input("اسم العميل / الشركة *:", value=c_info['name'])
                        new_address = st.text_input("العنوان:", value=c_info['address'] or "")
                        new_contact = st.text_input("الشخص المسؤول للتنسيق:", value=c_info['contact_person'] or "")
                    with col_c2:
                        new_phone = st.text_input("رقم الهاتف الأرضي أو المحمول:", value=c_info['phone'] or "")
                        new_email = st.text_input("البريد الإلكتروني للعميل:", value=c_info['email'] or "")
                        new_notes = st.text_input("ملاحظات تنظيمية خاصة بالمنشأة:", value=c_info['notes'] or "")
                    if st.form_submit_button("حفظ الملف المحدث سحابيّاً"):
                        if new_name:
                            try:
                                query_cl_up = "UPDATE clients SET name=:name, address=:addr, phone=:phone, email=:email, contact_person=:cp, notes=:notes WHERE id=:id"
                                with engine.begin() as conn_cl_up:
                                    conn_cl_up.execute(text(query_cl_up), {"name": new_name, "addr": new_address, "phone": new_phone, "email": new_email, "cp": new_contact, "notes": new_notes, "id": int(selected_client_id)})
                                st.success("✅ تم تحديث بروفايل العميل بنجاح!")
                                time.sleep(0.3)
                                st.rerun()
                            except:
                                st.error("❌ هذا الاسم مسجل لجهة أو عميل آخر بالفعل في النظام.")
                        else:
                            st.error("اسم العميل أو الجهة إلزامي.")
            
            st.markdown("---")
            equip_df = pd.read_sql_query(text("SELECT * FROM equipment WHERE client_id = :cid"), engine, params={"cid": int(selected_client_id)})
            
            st.subheader(f"🖨️ الأجهزة والمعدات التابعة لـ {c_info['name']}")
            if not equip_df.empty:
                equip_df['حالة العقد'] = equip_df['sla_expiration_date'].apply(calculate_sla_status)
                st.dataframe(equip_df[['brand', 'model', 'serial_number', 'sla_type', 'pm_visits_count', 'حالة العقد']].rename(columns={'pm_visits_count': 'الزيارات الدورية السنوية'}), use_container_width=True)
                
                # تصدير أجهزة العميل المختار
                st.download_button(f"📥 تصدير قائمة أجهزة {c_info['name']} لـ Excel", to_excel(equip_df), f"أجهزة_{c_info['name']}.xlsx", "application/vnd.ms-excel")

                with st.expander("⚙️ تعديل بيانات تفصيلية أو نوع عقد أحد هذه الأجهزة"):
                    eq_dict = {f"{r['brand']} {r['model']} - S/N: {r['serial_number']}": r['id'] for _, r in equip_df.iterrows()}
                    eq_id = eq_dict[st.selectbox("اختر الجهاز المراد تعديله بدقة:", list(eq_dict.keys()))]
                    eq_info = equip_df[equip_df['id'] == eq_id].iloc[0]
                    
                    if eq_info['contract_file'] and os.path.exists(str(eq_info['contract_file'])):
                        with open(str(eq_info['contract_file']), "rb") as f:
                            st.download_button("📥 تحميل نسخة العقد المؤرشفة للمعدّة", f, file_name=os.path.basename(str(eq_info['contract_file'])))
                            
                    with st.form("edit_eq_form"):
                        col_u1, col_u2 = st.columns(2)
                        with col_u1:
                            u_brand = st.selectbox("الماركة المحورية *:", ["Xerox", "Canon", "Konica Minolta", "OKI", "Kardex Remstar"], index=["Xerox", "Canon", "Konica Minolta", "OKI", "Kardex Remstar"].index(eq_info['brand']) if eq_info['brand'] in ["Xerox", "Canon", "Konica Minolta", "OKI", "Kardex Remstar"] else 0)
                            u_model = st.text_input("الموديل الدقيق *:", value=eq_info['model'])
                            u_serial = st.text_input("الرقم التسلسلي المقيد *:", value=eq_info['serial_number'])
                            idx = sla_list_db.index(eq_info['sla_type']) if eq_info['sla_type'] in sla_list_db else 0
                            sla_choice = st.selectbox("نوع العقد الفعلي الحالي بقاعدة البيانات:", extended_sla_list, index=idx)
                            final_sla = st.text_input("اكتب مسمى العقد البديل هنا:") if sla_choice == "+ إضافة نوع عقد جديد يدوياً..." else sla_choice
                            
                            p_options = ["نعم (من المكتب الرقمي)", "لا (من مورد آخر)"]
                            u_purchased = st.selectbox("هل الشراء تم من المكتب الرقمي؟", p_options, index=p_options.index(eq_info['purchased_from_us']) if eq_info['purchased_from_us'] in p_options else 0)
                            u_vendor = st.text_input("مورد خارجي بديل (إن وجد):", value=eq_info['third_party_vendor'] or "")
                        with col_u2:
                            u_building = st.text_input("المبنى/البلوك الميداني:", value=eq_info['location_building'] or "")
                            u_floor = st.text_input("الطابق/الدور المتموضع به:", value=eq_info['location_floor'] or "")
                            u_dept = st.text_input("القسم/الغرفة الإدارية المحددة:", value=eq_info['location_department'] or "")
                            u_pm_count = st.number_input("الزيارات الوقائية المطلوبة سنوياً:", min_value=0, max_value=12, value=int(eq_info['pm_visits_count'] or 0))
                        
                        try:
                            curr_exp = datetime.strptime(str(eq_info['sla_expiration_date']), "%Y-%m-%d").date()
                        except:
                            curr_exp = datetime.now().date() + timedelta(days=365)
                            
                        u_exp_date = st.date_input("تاريخ انتهاء غطاء صيانة العقد المبرم:", curr_exp)
                        uploaded_contract = st.file_uploader("تغيير/تحديث الملف الرقمي لنسخة العقد الممسوحة ضوئياً", type=['pdf', 'jpg', 'png'])
                        
                        if st.form_submit_button("اعتماد وحفظ تعديلات المعدّة"):
                            if u_model and u_serial and final_sla:
                                if final_sla not in sla_list_db and sla_choice == "+ إضافة نوع عقد جديد يدوياً...":
                                    with engine.begin() as conn_s_add:
                                        conn_s_add.execute(text("INSERT INTO sla_types (name) VALUES (:name) ON CONFLICT DO NOTHING"), {"name": final_sla})
                                contract_path = eq_info['contract_file']
                                if uploaded_contract:
                                    contract_path = save_uploaded_file(uploaded_contract, "uploads/contracts", f"eq_{eq_id}")
                                try:
                                    query_up_eq = """
                                        UPDATE equipment SET brand=:brand, model=:model, serial_number=:sn, sla_type=:sla, sla_expiration_date=:exp, contract_file=:file,
                                        location_building=:bld, location_floor=:flr, location_department=:dept, pm_visits_count=:pm, purchased_from_us=:pf, third_party_vendor=:tp WHERE id=:id
                                    """
                                    with engine.begin() as conn_eq_up:
                                        conn_eq_up.execute(text(query_up_eq), {
                                            "brand": u_brand, "model": u_model, "sn": u_serial, "sla": final_sla, "exp": str(u_exp_date), "file": contract_path,
                                            "bld": u_building, "flr": u_floor, "dept": u_dept, "pm": int(u_pm_count), "pf": u_purchased, "tp": u_vendor, "id": int(eq_id)
                                        })
                                    st.success("✅ تم تحديث تفاصيل كرت المعدّة على قاعدة البيانات الموحدة!")
                                    time.sleep(0.3)
                                    st.rerun()
                                except IntegrityError:
                                    st.error("❌ الرقم التسلسلي (S/N) مكرر ومسجل لجهاز آخر بالمنظومة، يرجى التثبت!")
            else:
                st.info("لا توجد أجهزة مسجلة ومربوطة بهذا العميل حتى الآن.")
                
            with st.expander("➕ إضافة معدّة/جهاز جديد وربطه بهذا العميل مباشرة"):
                with st.form("add_eq_to_client"):
                    c1, c2 = st.columns(2)
                    with c1:
                        brand = st.selectbox("الماركة الفنية *:", ["Xerox", "Canon", "Konica Minolta", "OKI", "Kardex Remstar"])
                        model = st.text_input("الموديل *:")
                        s_num = st.text_input("الرقم التسلسلي المحفور (S/N) *:")
                        purchased = st.selectbox("هل تم التوريد عبر المكتب الرقمي؟ *", ["نعم (من المكتب الرقمي)", "لا (من مورد آخر)"])
                        vendor = st.text_input("اسم المورد البديل إن وجد:")
                    with c2:
                        building = st.text_input("المبنى / البلوك الإنشائي:")
                        dept = st.text_input("القسم الوظيفي الداخلي:")
                        pm_count = st.number_input("الزيارات الدورية المطلوبة سنوياً لتفعيل الـ PM:", min_value=0, max_value=12, value=0)
                    
                    col_sla1, col_sla2 = st.columns(2)
                    with col_sla1:
                        sla_choice_new = st.selectbox("تأطير نوع العقد الحاكم:", extended_sla_list)
                        final_sla_new = st.text_input("اكتب اسم العقد غير المدرج بالقائمة:") if sla_choice_new == "+ إضافة نوع عقد جديد يدوياً..." else sla_choice_new
                    with col_sla2:
                        ins_date = st.date_input("تاريخ التركيب الفعلي وتشغيل الضمان:")
                        exp_date = st.date_input("تاريخ انتهاء غطاء الصيانة والعقد المبرم المعتمد:", datetime.now().date() + timedelta(days=365))
                    new_contract_file = st.file_uploader("تحميل المنسوخ الضوئي لملف العقد الرسمي إن توفر", type=['pdf', 'jpg', 'png'])
                    
                    if st.form_submit_button("حفظ وإدراج كرت المعدّة الجديد"):
                        if model and s_num and final_sla_new:
                            if final_sla_new not in sla_list_db and sla_choice_new == "+ إضافة نوع عقد جديد يدوياً...":
                                with engine.begin() as conn_s_in:
                                    conn_s_in.execute(text("INSERT INTO sla_types (name) VALUES (:name) ON CONFLICT DO NOTHING"), {"name": final_sla_new})
                            try:
                                contract_path = save_uploaded_file(new_contract_file, "uploads/contracts", f"new_eq_{int(time.time())}") if new_contract_file else None
                                query_in_eq = """
                                    INSERT INTO equipment (client_id, brand, model, serial_number, installation_date, sla_type, sla_expiration_date, contract_file, 
                                    location_building, location_department, pm_visits_count, purchased_from_us, third_party_vendor) 
                                    VALUES (:cid, :brand, :model, :sn, :ins, :sla, :exp, :file, :bld, :dept, :pm, :pf, :tp)
                                """
                                with engine.begin() as conn_eq_in:
                                    conn_eq_in.execute(text(query_in_eq), {
                                        "cid": int(selected_client_id), "brand": brand, "model": model, "sn": s_num, "ins": str(ins_date),
                                        "sla": final_sla_new, "exp": str(exp_date), "file": contract_path, "bld": building, "dept": dept,
                                        "pm": int(pm_count), "pf": purchased, "tp": vendor
                                    })
                                st.success("✅ تم حفظ وإدراج المعدّة وربطها بملف العميل بنجاح!")
                                time.sleep(0.3)
                                st.rerun()
                            except IntegrityError:
                                st.error("❌ الرقم التسلسلي مكرر ومستعمل مع معدّة أخرى، يرجى المراجعة والتأكد!")
                        else:
                            st.warning("يرجى التأكد من ملء الموديل، السيريال، ونوع العقد لإتمام الإجراء.")
        else:
            st.warning("المستودع الرقمي السحابي فارغ من العملاء، يرجى إضافة عميل أولاً بالتبويب الجانبي.")

    with tab2:
        with st.form("add_new_client"):
            st.write("تسجيل وإدراج عميل/شركة جديدة بملف متكامل ومستقل:")
            col1, col2 = st.columns(2)
            with col1:
                name = st.text_input("اسم العميل / المؤسسة بالكامل *:")
                addr = st.text_input("عنوان الإدارة الرئيسي والمقر الميداني:")
                contact = st.text_input("اسم نقطة التواصل (الشخص المسؤول بالموقع):")
            with col2:
                phone = st.text_input("أرقام الهواتف الرسمية للاتصال المباشر:")
                email = st.text_input("البريد الإلكتروني المعتمد للمراسلات المباشرة:")
                notes = st.text_input("ملاحظات إضافية حول شروط الوصول الميداني:")
            if st.form_submit_button("إدراج وحفظ ملف العميل الجديد"):
                if name:
                    try:
                        with engine.begin() as conn_cl_new:
                            conn_cl_new.execute(text("INSERT INTO clients (name, address, phone, email, contact_person, notes) VALUES (:name, :addr, :phone, :email, :cp, :notes)"),
                                                {"name": name, "addr": addr, "phone": phone, "email": email, "cp": contact, "notes": notes})
                        st.success("🎉 تم إضافة وتأصيل ملف العميل بنجاح بالمستودع السحابي الموحد!")
                        time.sleep(0.3)
                        st.rerun()
                    except:
                        st.error("❌ اسم هذا العميل مكرر ومسجل مسبقاً بالنظام السحابي.")
                else:
                    st.warning("اسم العميل أو الجهة حقل إلزامي لا يمكن تركه فارغاً.")

    with tab3:
        all_eq = pd.read_sql_query(text("SELECT c.name as client_name, e.* FROM equipment e JOIN clients c ON e.client_id = c.id ORDER BY e.id DESC"), engine)
        if not all_eq.empty:
            all_eq['حالة العقد'] = all_eq['sla_expiration_date'].apply(calculate_sla_status)
            all_eq['موقع الجهاز بالتفصيل'] = all_eq['location_building'].fillna('') + " - " + all_eq['location_department'].fillna('')
            
            st.subheader("📋 الجرد العام والشامل لجميع المعدّات السحابية")
            st.dataframe(all_eq[['client_name', 'brand', 'model', 'serial_number', 'موقع الجهاز بالتفصيل', 'sla_type', 'pm_visits_count', 'حالة العقد']].rename(columns={'client_name': 'العميل المالك', 'pm_visits_count': 'الزيارات الدورية السنوية'}), use_container_width=True)
            
            # أدوات تصدير الجرد العام لجميع الأجهزة والشركات
            ex_eq1, ex_eq2 = st.columns(2)
            with ex_eq1:
                st.download_button("📥 تصدير الجرد العام للأجهزة لـ Excel", to_excel(all_eq), "الجرد_العام_للمعدات.xlsx", "application/vnd.ms-excel")
            with ex_eq2:
                st.download_button("🖨️ طباعة نموذج الجرد والتوريد بالكامل (CSV)", to_csv_printable(all_eq), "طباعة_الجرد_العام.csv", "text/csv")
        else:
            st.info("لم يتم إدراج أو قيد أي معدّات أو أجهزة في النظام حتى الآن.")

# ----------------------------------------------------
# 👨‍💻 إدارة فريق الفنيين
# ----------------------------------------------------
elif menu == "👨‍💻 إدارة فريق الفنيين":
    st.title("👨‍💻 سجل كادر الفنيين والمهندسين والتحكم بالحالة الميدانية")
    tech_tab1, tech_tab2, tech_tab3 = st.tabs(["📋 قائمة الفنيين ومتابعة الحالة", "➕ إضافة فني جديد للفريق", "✏️ تعديل بيانات مهندس"])
    status_options = ["متاح", "في إجازة سنوية", "في إجازة طارئة", "في إجازة مرضية", "في رحلة عمل"]
    
    df_techs = pd.read_sql_query(text("SELECT * FROM technicians ORDER BY id"), engine)

    with tech_tab1:
        if not df_techs.empty:
            def color_tech_status(val):
                if val == 'متاح': color = 'green'
                elif val in ['في إجازة سنوية', 'في إجازة طارئة', 'في إجازة مرضية']: color = 'red'
                elif val == 'في رحلة عمل': color = 'orange'
                else: color = 'black'
                return f'color: {color}; font-weight: bold;'
                
            display_df = df_techs[['id', 'name', 'specialty', 'phone', 'email', 'status']].copy()
            # 🌟 إبراز اسم الموظف المثالي الحالي بنجمة في قائمة الفنيين أيضاً
            winner_name, _ = get_star_technician()
            if winner_name:
                display_df['name'] = display_df['name'].apply(lambda x: f"{x} ⭐" if x == winner_name else x)
                
            display_df.columns = ['رقم القيد', 'الاسم بالكامل', 'التخصص التقني الميداني', 'رقم الهاتف المباشر', 'البريد الإلكتروني', 'حالة التوافر الحالية']
            st.dataframe(display_df.style.map(color_tech_status, subset=['حالة التوافر الحالية']), use_container_width=True)
            
            # خيارات التصدير لفريق العمل
            st.download_button("📥 تصدير قائمة وسجل كادر الفنيين لـ Excel", to_excel(df_techs), "فريق_العمل_الفني.xlsx", "application/vnd.ms-excel")
        else:
            st.info("لا توجد أسماء مقيدة بسجل الفنيين التابعين للمنظومة المركزية.")
            
    with tech_tab2:
        with st.form("add_tech_form"):
            col1, col2 = st.columns(2)
            with col1:
                t_name = st.text_input("اسم المهندس / الفني الثلاثي *:")
                t_spec = st.text_input("التخصص الفني الدقيق (مثل: أنظمة إنتاج رقمي، PLC، ميكانيكا طابعات):")
                t_status = st.selectbox("حالة العمل والتوافر الأولية:", status_options)
            with col2:
                t_phone = st.text_input("رقم الهاتف المباشر والخاص بالعمل:")
                t_email = st.text_input("البريد الإلكتروني المهني (المكتب الرقمي):")
            if st.form_submit_button("اعتماد وإدراج الفني بالفريق"):
                if t_name:
                    with engine.begin() as conn_tc_in:
                        conn_tc_in.execute(text("INSERT INTO technicians (name, specialty, phone, email, status) VALUES (:name, :spec, :phone, :email, :status)"),
                                            {"name": t_name, "spec": t_spec, "phone": t_phone, "email": t_email, "status": t_status})
                    st.success("🎉 تم إضافة وتثبيت المهندس الفني الجديد بنجاح في قاعدة البيانات!")
                    time.sleep(0.3)
                    st.rerun()
                else:
                    st.warning("يرجى كتابة الاسم الثلاثي للفني لإكمال القيد بنجاح.")
                    
    with tech_tab3:
        if not df_techs.empty:
            tech_dict = {f"{row['name']} - ({row['specialty'] or 'دون تخصص'})": row['id'] for _, row in df_techs.iterrows()}
            selected_tech = st.selectbox("🔍 اختر الفني أو المهندس المراد تعديل ملفه وحالته التوافرية:", list(tech_dict.keys()))
            tech_id = tech_dict[selected_tech]
            t_info = df_techs[df_techs['id'] == tech_id].iloc[0]
            
            with st.form("edit_tech_form"):
                col_e1, col_e2 = st.columns(2)
                with col_e1:
                    u_name = st.text_input("الاسم بالكامل وبدقة *:", value=t_info['name'])
                    u_spec = st.text_input("التخصص التقني الفعلي المعتمد:", value=t_info['specialty'] or "")
                    cur_st_t = t_info['status'] if t_info['status'] in status_options else "متاح"
                    u_status = st.selectbox("تحديث الحالة الحالية الفورية للعمل:", status_options, index=status_options.index(cur_st_t))
                with col_e2:
                    u_phone = st.text_input("رقم الهاتف المعدّل للتواصل السريع:", value=t_info['phone'] or "")
                    u_email = st.text_input("البريد الإلكتروني الرسمي المحدث للعمل:", value=t_info['email'] or "")
                if st.form_submit_button("حفظ التحديثات بملف الفني"):
                    if u_name:
                        query_tech_up = "UPDATE technicians SET name=:name, specialty=:spec, phone=:phone, email=:email, status=:status WHERE id=:id"
                        with engine.begin() as conn_tc_up:
                            conn_tc_up.execute(text(query_tech_up), {"name": u_name, "spec": u_spec, "phone": u_phone, "email": u_email, "status": u_status, "id": int(tech_id)})
                        st.success("✅ تم تحديث ملف المهندس وتثبيت حالته الميدانية الفورية بنجاح!")
                        time.sleep(0.3)
                        st.rerun()
                    else:
                        st.error("الاسم الثلاثي للفني حقل إلزامي لتوثيق الهوية البرمجية.")

# ----------------------------------------------------
# ⚙️ إعدادات أنواع العقود
# ----------------------------------------------------
elif menu == "⚙️ إعدادات أنواع العقود":
    st.title("⚙️ إدارة مسميات وأنواع العقود والـ SLAs المركزية")
    
    sla_df = pd.read_sql_query(text("SELECT id, name as \"نوع العقد المعتمد\" FROM sla_types ORDER BY id"), engine)
    col_s1, col_s2 = st.columns([1, 2])
    
    with col_s1:
        with st.form("add_sla"):
            st.subheader("➕ إضافة مسمى عقد مركزي جديد")
            new_sla_name = st.text_input("اكتب اسم أو تصنيف العقد الجديد بدقة:")
            if st.form_submit_button("حفظ وتأكيد الإضافة"):
                if new_sla_name:
                    try:
                        with engine.begin() as conn_sla_in:
                            conn_sla_in.execute(text("INSERT INTO sla_types (name) VALUES (:name)"), {"name": new_sla_name})
                        st.success("تمت الإضافة بنجاح للمستودع المرجعي!")
                        time.sleep(0.3)
                        st.rerun()
                    except:
                        st.error("❌ هذا المسمى مسجل وموجود بالفعل في القائمة.")
                else:
                    st.warning("يرجى إدخال مسمى صحيح.")
        
        with st.expander("✏️ تعديل اسم أو صياغة عقد متاح حالياً"):
            if not sla_df.empty:
                with st.form("edit_sla_form"):
                    old_name = st.selectbox("اختر العقد المراد تعديل مسمّاه:", sla_df["نوع العقد المعتمد"].tolist())
                    updated_name = st.text_input("اكتب الاسم المركزي الجديد بدقة ومحاذاة:")
                    if st.form_submit_button("تحديث الاسم الحاكم بجميع الجداول"):
                        if updated_name:
                            try:
                                with engine.begin() as conn_sla_up:
                                    conn_sla_up.execute(text("UPDATE sla_types SET name = :new_name WHERE name = :old_name"), {"new_name": updated_name, "old_name": old_name})
                                    conn_sla_up.execute(text("UPDATE equipment SET sla_type = :new_name WHERE sla_type = :old_name"), {"new_name": updated_name, "old_name": old_name})
                                st.success("تم تحديث مسمى العقد بنجاح وتعديله بكافة كروت الأجهزة المرتبطة!")
                                time.sleep(0.3)
                                st.rerun()
                            except:
                                st.error("❌ الاسم الجديد مستعمل مع نوع عقد آخر متاح.")
                        else:
                            st.warning("أدخل المسمى البديل الجديد.")
        
        with st.expander("🗑️ حذف نوع عقد نهائياً من الجداول المرجعية"):
            if not sla_df.empty:
                with st.form("delete_sla_form"):
                    del_name = st.selectbox("اختر العقد لإزالته نهائياً من النظام السحابي:", sla_df["نوع العقد المعتمد"].tolist())
                    st.warning("⚠️ تنبيه تقني: سيتم حذفه من قائمة الخيارات المتاحة للأجهزة الجديدة فقط ولن يحذف الأجهزة التاريخية.")
                    if st.form_submit_button("تأكيد الحذف النهائي السحابي"):
                        with engine.begin() as conn_sla_del:
                            conn_sla_del.execute(text("DELETE FROM sla_types WHERE name = :name"), {"name": del_name})
                        st.success("🗑️ تم حذف وإلغاء تصنيف العقد من السجل المرجعي بنجاح.")
                        time.sleep(0.3)
                        st.rerun()

    with col_s2:
        st.subheader("📋 قائمة مسميات العقود المسجلة مركزياً")
        st.dataframe(sla_df, use_container_width=True)
        
        # أداة تصدير أنواع العقود
        st.download_button("📥 تصدير مسميات العقود لـ Excel", to_excel(sla_df), "أنواع_العقود_الرسمية.xlsx", "application/vnd.ms-excel")
