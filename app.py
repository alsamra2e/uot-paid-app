import streamlit as st
import pandas as pd
import plotly.express as px
import re
import io
import difflib

# --- Page Configuration & Modern Styling ---
st.set_page_config(page_title="نظام متابعة تسديدات الطلاب", layout="wide")

st.markdown("""
    <style>
        body, .stApp { direction: rtl; text-align: right; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; }
        .stMetric { background-color: #f8f9fa; padding: 15px; border-radius: 10px; border-left: 5px solid #2ecc71; box-shadow: 0 4px 6px rgba(0,0,0,0.05); }
        h1, h2, h3 { color: #2c3e50; }
        .stButton>button { border-radius: 8px; font-weight: bold; }
    </style>
""", unsafe_allow_html=True)

st.title("📊 نظام مطابقة القاعات والتسديدات")

# --- Helper Functions ---
def clean_arabic_name(name):
    if pd.isna(name): return ""
    name = str(name).strip()
    name = re.sub(r'[أإآ]', 'ا', name)
    name = re.sub(r'ة', 'ه', name)
    name = re.sub(r'ى', 'ي', name)
    name = re.sub(r'\s+', ' ', name)
    return name

def clean_currency(val):
    if pd.isna(val): return 0.0
    val = str(val).replace('IQD', '').replace(',', '').strip()
    try: return float(val)
    except: return 0.0

def process_percentage(val):
    if pd.isna(val): return 0.0
    if isinstance(val, (int, float)):
        if 0 <= val <= 1:
            return val * 100
        return float(val)
    val = str(val).replace('%', '').strip()
    try: return float(val)
    except: return 0.0

# --- Load Static Main Data (Cached for speed) ---
@st.cache_data
def load_main_data():
    try:
        df = pd.read_excel("main_database.xlsx")
    except FileNotFoundError:
        st.warning("يرجى التأكد من وجود ملف main_database.xlsx في نفس المجلد.")
        return pd.DataFrame()
        
    df.columns = df.columns.str.strip()
    
    if 'اسم الطالب' not in df.columns:
        st.error(f"⚠️ خطأ في ملف الإكسل: لم أتمكن من العثور على عمود 'اسم الطالب'. الأعمدة التي وجدتها هي: {list(df.columns)}")
        return pd.DataFrame()

    df['Match_Key'] = df['اسم الطالب'].apply(clean_arabic_name)
    df = df.drop_duplicates(subset=['Match_Key'], keep='first')
    return df

df_main = load_main_data()

# --- Daily File Uploader ---
st.sidebar.header("رفع البيانات اليومية")
uploaded_file = st.sidebar.file_uploader("ارفع ملف التسديدات اليومي (Excel)", type=["xlsx", "xls"])

if uploaded_file and not df_main.empty:
    df_daily = pd.read_excel(uploaded_file)
    df_daily.columns = df_daily.columns.str.strip()
    
    if df_daily.empty:
        st.error("⚠️ الملف المرفوع فارغ.")
    else:
        # --- DYNAMIC DEPARTMENT DETECTION ---
        if 'المرحلة' in df_daily.columns:
            detected_departments = df_daily['المرحلة'].dropna().unique()
            st.success(f"✅ تم اكتشاف الأقسام/المراحل التالية في الملف: {', '.join(detected_departments)}")
        
        # Clean and prepare daily data
        if 'اسم الطالب' not in df_daily.columns:
            st.error(f"⚠️ خطأ: الملف اليومي لا يحتوي على عمود 'اسم الطالب'. الأعمدة الموجودة: {list(df_daily.columns)}")
            st.stop()
            
        df_daily['Match_Key_Daily'] = df_daily['اسم الطالب'].apply(clean_arabic_name)
        
        # Fallbacks for currency and percentage in case column names vary
        currency_col = 'المبلغ المتبقي' if 'المبلغ المتبقي' in df_daily.columns else df_daily.columns[-2]
        percent_col = 'نسبة الدفع' if 'نسبة الدفع' in df_daily.columns else df_daily.columns[-1]
        
        df_daily['المبلغ المتبقي_رقم'] = df_daily[currency_col].apply(clean_currency)
        df_daily['نسبة الدفع_رقم'] = df_daily[percent_col].apply(process_percentage)
        df_daily['نسبة الدفع_معروضة'] = df_daily['نسبة الدفع_رقم'].apply(lambda x: f"{int(x) if x.is_integer() else x}%")
        
        # --- Fuzzy Matching Logic ---
        main_names_list = df_main['Match_Key'].tolist()
        
        def find_best_match(daily_name):
            if daily_name in main_names_list:
                return daily_name
            matches = difflib.get_close_matches(daily_name, main_names_list, n=1, cutoff=0.85)
            return matches[0] if matches else None

        with st.spinner('جاري مطابقة الأسماء...'):
            df_daily['Matched_DB_Name'] = df_daily['Match_Key_Daily'].apply(find_best_match)

        # Merge Data
        merged_df = pd.merge(df_daily, df_main[['Match_Key', 'الرمز', 'القسم والمرحلة', 'رقم القاعة']], 
                             left_on='Matched_DB_Name', right_on='Match_Key', 
                             how='left')
        
        # --- Dashboard Filters ---
        st.sidebar.markdown("---")
        st.sidebar.header("تصفية البيانات")
        
        # Safe Multiselect
        if 'المرحلة' in merged_df.columns:
            stage_options = merged_df['المرحلة'].dropna().unique()
            selected_stage = st.sidebar.multiselect("اختر القسم / المرحلة لعرضها:", options=stage_options)
        else:
            selected_stage = []

        payment_threshold = st.sidebar.slider("نسبة الدفع أقل من أو تساوي (%)", 0, 100, 100)
        
        # Apply filters
        filtered_df = merged_df[merged_df['نسبة الدفع_رقم'] <= payment_threshold]
        
        if selected_stage:
            filtered_df = filtered_df[filtered_df['المرحلة'].isin(selected_stage)]

        # --- Metrics & Infographics ---
        st.markdown("### 📈 ملخص الإحصائيات")
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("إجمالي الطلاب", f"{len(filtered_df)}")
        col2.metric("إجمالي الديون المتبقية", f"{filtered_df['المبلغ المتبقي_رقم'].sum():,.0f} IQD")
        col3.metric("متوسط نسبة الدفع", f"{filtered_df['نسبة الدفع_رقم'].mean():.1f}%")
        col4.metric("حالات غير مطابقة", f"{filtered_df['رقم القاعة'].isna().sum()}")

        # Charts
        c1, c2 = st.columns(2)
        with c1:
            fig1 = px.histogram(filtered_df, x="نسبة الدفع_رقم", nbins=10, 
                                title="توزيع نسب الدفع",
                                color_discrete_sequence=['#3498db'])
            fig1.update_layout(xaxis_title="نسبة الدفع (%)", yaxis_title="عدد الطلاب")
            st.plotly_chart(fig1, use_container_width=True)
            
        with c2:
            # --- DYNAMIC CHART FIX ---
            # Automatically choose a grouping column depending on what is available
            chart_col = 'المرحلة' if 'المرحلة' in filtered_df.columns else ('القسم والمرحلة' if 'القسم والمرحلة' in filtered_df.columns else None)
            
            if chart_col:
                debt_by_stage = filtered_df.groupby(chart_col)['المبلغ المتبقي_رقم'].sum().reset_index()
                fig2 = px.pie(debt_by_stage, values='المبلغ المتبقي_رقم', names=chart_col, 
                              title="حجم الديون المتبقية حسب المرحلة الدراسية", hole=0.4)
                st.plotly_chart(fig2, use_container_width=True)
            else:
                st.info("📊 المخطط الدائري غير متاح: لا يوجد عمود يحدد المرحلة أو القسم.")

        # --- Data Display ---
        st.markdown("### 📝 جدول المطابقة النهائي")
        
        # --- DYNAMIC TABLE COLUMNS FIX ---
        # Only display columns that actually exist in the merged dataframe
        desired_cols = ['التسلسل', 'اسم الطالب', 'المرحلة', 'القسم والمرحلة', 'رقم القاعة', currency_col, 'نسبة الدفع_معروضة']
        display_cols = [col for col in desired_cols if col in filtered_df.columns]
        
        missing_matches = filtered_df[filtered_df['رقم القاعة'].isna()]
        if not missing_matches.empty:
            st.error(f"⚠️ يوجد {len(missing_matches)} طالب لم يتم العثور على قاعاتهم. يرجى مراجعة أسمائهم.")
            with st.expander("عرض الطلاب غير المطابقين"):
                missing_display = [col for col in ['التسلسل', 'اسم الطالب', 'المرحلة'] if col in missing_matches.columns]
                st.dataframe(missing_matches[missing_display], hide_index=True)
        
        # Rename the displayed percentage column to look clean
        final_table = filtered_df[display_cols].rename(columns={'نسبة الدفع_معروضة': 'نسبة الدفع'})
        st.dataframe(final_table, use_container_width=True, hide_index=True)

        # --- Export to Excel ---
        st.markdown("---")
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            final_table.to_excel(writer, index=False, sheet_name='تقرير المطابقة')
        
        st.download_button(
            label="📥 تحميل التقرير النهائي (Excel)",
            data=buffer.getvalue(),
            file_name="Matched_Students_Report.xlsx",
            mime="application/vnd.ms-excel"
        )

else:
    st.info("👈 يرجى رفع ملف التسديدات اليومي من القائمة الجانبية للبدء.")
