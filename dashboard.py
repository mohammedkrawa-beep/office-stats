import pandas as pd
import streamlit as st
import plotly.express as px
import base64
import os

# 1. إعدادات الصفحة
st.set_page_config(page_title="لوحة تحكم المكاتب", layout="wide")

# دالة لقراءة الصور
def get_base64_image(file_path):
    if os.path.exists(file_path):
        with open(file_path, "rb") as f:
            return base64.b64encode(f.read()).decode()
    return ""

# المسارات
logo_path = r"C:\Users\MSI\OneDrive\Desktop\MyDashboard\شعار المركز ابيض\شعار المركز ابيض.png"
excel_path = "جدول_المكاتب_المدمج_النهائي.xlsx" 
logo_base64 = get_base64_image(logo_path)

# 2. التنسيق الجمالي (CSS) والهيدر
st.markdown(f"""
    <style>
    .stApp {{ background-color: #f8f9fa; }}
    
    /* الهيدر العلوي */
    .top-header {{
        background-color: #1e3d59;
        margin: -75px -100px 30px -100px;
        padding: 40px 100px;
        border-bottom: 5px solid #C5A059;
        display: flex;
        align-items: center;
        justify-content: space-between;
    }}
    
    /* إصلاح صندوق البحث - نص كحلي واضح */
    input[type="text"] {{
        color: #1e3d59 !important;
        background-color: white !important;
    }}
    
    /* تنسيق زر العودة للوضع العام */
    div.stButton > button {{
        background-color: #C5A059 !important; 
        color: white !important;
        border-radius: 5px;
        border: none;
        width: 100%;
        font-weight: bold;
    }}
    div.stButton > button:hover {{
        background-color: white !important;
        color: #1e3d59 !important;
        border: 2px solid #1e3d59 !important;
    }}

    /* القائمة الجانبية */
    [data-testid="stSidebar"] {{ background-color: #162a3d !important; }}
    [data-testid="stSidebar"] * {{ color: white !important; }}
    </style>

    <div class="top-header">
        <img src="data:image/png;base64,{logo_base64}" style="width:300px;">
        <div style="color: white; text-align: right;">
            <h1 style="margin:0;">مكتب شؤون المكاتب</h1>
            <p style="margin:0; color: #C5A059;">شاشة عرض المراسلات الصادرة</p>
        </div>
    </div>
    """, unsafe_allow_html=True)

# 3. تحميل البيانات
try:
    df = pd.read_excel(excel_path)
    df.columns = [str(c).strip() for c in df.columns]
    if 'تاريخ البريد' in df.columns:
        df['تاريخ البريد'] = pd.to_datetime(df['تاريخ البريد'])
except Exception as e:
    st.error(f"خطأ في تحميل الملف: {e}")
    st.stop()

# --- 4. القائمة الجانبية (الفلاتر) ---
st.sidebar.title("🔍 خيارات التصفية")

if st.sidebar.button("🔄 العودة للوضع العام"):
    st.rerun()

st.sidebar.markdown("---")
search_query = st.sidebar.text_input("🔍 بحث شامل في المستند:")

office_col = 'الجهات المرسل لها'
subject_col = 'الموضوع'

unique_offices = sorted(df[office_col].dropna().unique())
selected_offices = st.sidebar.multiselect("🏢 اختر المكتب:", unique_offices)

unique_subjects = sorted(df[subject_col].dropna().unique())
selected_subjects = st.sidebar.multiselect("📝 اختر الموضوع:", unique_subjects)

if 'تاريخ البريد' in df.columns:
    min_d = df['تاريخ البريد'].min().date()
    max_d = df['تاريخ البريد'].max().date()
    date_range = st.sidebar.date_input("📅 النطاق الزمني:", [min_d, max_d])

# --- 5. تطبيق الفلترة ---
df_filtered = df.copy()

if 'تاريخ البريد' in df.columns and len(date_range) == 2:
    df_filtered = df_filtered[(df_filtered['تاريخ البريد'].dt.date >= date_range[0]) & 
                              (df_filtered['تاريخ البريد'].dt.date <= date_range[1])]

if selected_offices:
    df_filtered = df_filtered[df_filtered[office_col].isin(selected_offices)]

if selected_subjects:
    df_filtered = df_filtered[df_filtered[subject_col].isin(selected_subjects)]

if search_query:
    mask = df_filtered.astype(str).apply(lambda x: x.str.contains(search_query, case=False, na=False)).any(axis=1)
    df_filtered = df_filtered[mask]

# --- 6. الإجماليات (KPIs) ---
st.markdown("### 📈 ملخص الإحصائيات العامة")
kpi1, kpi2, kpi3 = st.columns(3)
with kpi1:
    st.metric(label="إجمالي المراسلات", value=len(df_filtered))
with kpi2:
    st.metric(label="عدد المواضيع", value=df_filtered[subject_col].nunique())
with kpi3:
    st.metric(label="عدد الجهات", value=df_filtered[office_col].nunique())

st.markdown("---")

# --- 7. الرسوم البيانية ---
col1, col2 = st.columns(2)

with col1:
    st.subheader("📊 إحصائيات المواضيع")
    if not df_filtered.empty:
        top_sub = df_filtered.groupby(subject_col)['الجهة الصادرة'].count().reset_index()
        fig = px.pie(top_sub, names=subject_col, values='الجهة الصادرة', hole=0.5,
                     color_discrete_sequence=px.colors.sequential.Blues_r)
        
        # إظهار الرقم داخل بوكس كبير (داخل شريحة الدائرة)
        fig.update_traces(textinfo='value', textposition='inside', textfont_size=22)
        st.plotly_chart(fig, use_container_width=True)

with col2:
    st.subheader("🏢 توزيع المراسلات حسب الجهة")
    if not df_filtered.empty:
        office_counts = df_filtered[office_col].value_counts().head(10).reset_index()
        office_counts.columns = ['الجهة', 'العدد']
        
        # وضع الرقم داخل العمود ليعطي مظهر البوكس الملون
        fig2 = px.bar(office_counts, x='العدد', y='الجهة', orientation='h', 
                      color_discrete_sequence=['#C5A059'], text='العدد')
        fig2.update_traces(textposition='inside', textfont_size=18)
        st.plotly_chart(fig2, use_container_width=True)

# التسلسل الزمني
st.markdown("### 📅 التسلسل الزمني للمراسلات")
if not df_filtered.empty:
    timeline_data = df_filtered.groupby(df_filtered['تاريخ البريد'].dt.date).size().reset_index(name='العدد')
    fig_line = px.line(timeline_data, x='تاريخ البريد', y='العدد', markers=True, 
                       color_discrete_sequence=['#1e3d59'], text='العدد')
    fig_line.update_traces(textposition='top center', textfont_size=14)
    st.plotly_chart(fig_line, use_container_width=True)

# 8. عرض الجدول
st.markdown("### 📄 تفاصيل البيانات المفلترة")
st.dataframe(df_filtered.sort_values(by='تاريخ البريد', ascending=False), use_container_width=True)