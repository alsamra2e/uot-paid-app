import streamlit as st
import pandas as pd
import plotly.express as px
import re
import io
import difflib

# --- Page Configuration & Modern Styling ---
st.set_page_config(page_title="نظام متابعة تسديدات قسم القانون", layout="wide")

st.markdown("""
    <style>
        body, .stApp { direction: rtl; text-align: right; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; }
        .stMetric { background-color: #f8f9fa; padding: 15px; border-radius: 10px; border-left: 5px solid #2ecc71; box-shadow: 0 4px 6px rgba(0,0,0,0.05); }
        h1, h2, h3 { color: #2c3e50; }
        .stButton>button { border-radius: 8px; font-weight: bold; }
    </style>
""", unsafe_allow_html=True)

st.title("⚖️ نظام مطابقة القاعات والتسديدات - قسم القانون")

# --- Helper Functions ---
def clean_arabic_name(name):
    """Normalize Arabic text but KEEP the full name to ensure exact identity."""
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
    """Handles Excel's 0.7 decimal format and converts it to a clean 70.0"""
    if pd.isna(val): return 0.0
    # If Excel passes 0.7 for 70%
    if isinstance(val, (int, float)):
        if 0 <= val <= 1:
            return val * 100
        return float(val)
    # If it is read as a string like "70%"
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
        
    # Remove invisible spaces from Excel headers to prevent KeyError
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
    
    # --- AUTOMATIC DEPARTMENT FILTER: Keep ONLY Law (قانون) ---
    if 'المرحلة' in df_daily.columns:
        df_daily = df_daily[df_daily['المرحلة'].astype(str).str.contains('قانون', na=False)].copy()
        
    if df_daily.empty:
        st.error("⚠️ لم يتم العثور على أي طلاب لقسم 'قانون' في الملف المرفوع.")
    else:
        # Clean and prepare daily data
        df_daily['Match_Key_Daily'] = df_daily['اسم الطالب'].apply(clean_arabic_name)
        df_daily['المبلغ المتبقي_رقم'] = df_daily['المبلغ المتبقي'].apply(clean_currency)
        
        # Process the percentage logic
        df_daily['نسبة الدفع_رقم'] = df_daily['نسبة الدفع'].apply(process_percentage)
        # Format it back to look exactly like "70%" for the final table
        df_daily['نسبة الدفع'] = df_daily['نسبة الدفع_رقم'].apply(lambda x: f"{int(x) if x.is_integer() else x}%")
        
        # --- Fuzzy Matching Logic ---
        main_names_list = df_main['Match_Key'].tolist()
        
        def find_best_match(daily_name):
            if daily_name in main_names_list:
                return daily_name
            matches = difflib.get_close_matches(daily_name, main_names_list, n=1, cutoff=0.85)
            return matches[0] if matches else None

        with st.spinner('جاري مطابقة أسماء طلاب القانون...'):
            df_daily['Matched_DB_Name'] = df_daily['Match_Key_Daily'].apply(find_best_match)

        # Merge Data
        merged_df = pd.merge(df_daily, df_main[['Match_Key', 'الرمز', 'القسم والمرحلة', 'رقم القاعة']], 
                             left_on='Matched_DB_Name', right_on='Match_Key', 
                             how='left')
        
        # --- Dashboard Filters ---
        st.sidebar.markdown("---")
        st.sidebar.header("تصفية البيانات")
        payment_threshold = st.sidebar.slider("نسبة الدفع أقل من أو تساوي (%)", 0, 100, 100)
        
        # Apply filters
        filtered_df = merged_df[merged_df['نسبة الدفع_رقم'] <= payment_threshold]

        # --- Metrics & Infographics ---
        st.markdown("### 📈 ملخص الإحصائيات (قسم القانون فقط)")
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
            debt_by_stage = filtered_df.groupby('المرحلة')['المبلغ المتبقي_رقم'].sum().reset_index()
            fig2 = px.pie(debt_by_stage, values='المبلغ المتبقي_رقم', names='المرحلة', 
                          title="حجم الديون المتبقية حسب المرحلة الدراسية", hole=0.4)
            st.plotly_chart(fig2, use_container_width=True)

        # --- Data Display ---
        st.markdown("### 📝 جدول المطابقة النهائي")
        display_cols = ['التسلسل', 'اسم الطالب', 'المرحلة', 'القسم والمرحلة', 'رقم القاعة', 'المبلغ المتبقي', 'نسبة الدفع']
        
        missing_matches = filtered_df[filtered_df['رقم القاعة'].isna()]
        if not missing_matches.empty:
            st.error(f"⚠️ يوجد {len(missing_matches)} طالب لم يتم العثور على قاعاتهم. يرجى مراجعة أسمائهم.")
            with st.expander("عرض الطلاب غير المطابقين"):
                st.dataframe(missing_matches[['التسلسل', 'اسم الطالب', 'المرحلة']], hide_index=True)
        
        st.dataframe(filtered_df[display_cols], use_container_width=True, hide_index=True)

        # --- Export to Excel ---
        st.markdown("---")
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            filtered_df[display_cols].to_excel(writer, index=False, sheet_name='تقرير المطابقة')
        
        st.download_button(
            label="📥 تحميل التقرير النهائي (Excel)",
            data=buffer.getvalue(),
            file_name="Law_Students_Report.xlsx",
            mime="application/vnd.ms-excel"
        )

else:
    st.info("👈 يرجى رفع ملف التسديدات اليومي من القائمة الجانبية للبدء.")
