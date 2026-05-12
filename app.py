import streamlit as st
import pandas as pd
import plotly.express as px
import re
import io
import difflib

# --- Page Configuration & Modern Styling ---
st.set_page_config(page_title="نظام مطابقة التسديدات", layout="wide")

st.markdown("""
    <style>
        body, .stApp { direction: rtl; text-align: right; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; }
        .stMetric { background-color: #f8f9fa; padding: 15px; border-radius: 10px; border-left: 5px solid #2ecc71; box-shadow: 0 4px 6px rgba(0,0,0,0.05); }
        h1, h2, h3 { color: #2c3e50; }
        .stButton>button { border-radius: 8px; font-weight: bold; border: 1px solid #3498db; }
    </style>
""", unsafe_allow_html=True)

st.title("📊 نظام مطابقة القاعات والتسديدات المتقدم")

# --- Helper Functions ---
def clean_arabic_name(name):
    """Deep cleans names: removes commas, standardizes characters, fixes spacing for 4-5 word names."""
    if pd.isna(name): return ""
    name = str(name).strip()
    name = re.sub(r'[,،_.-]', ' ', name) # DESTROY commas and punctuation
    name = re.sub(r'[أإآ]', 'ا', name)
    name = re.sub(r'ة', 'ه', name)
    name = re.sub(r'ى', 'ي', name)
    name = re.sub(r'\s+', ' ', name).strip() # Fix multiple spaces between words
    return name

def clean_currency(val):
    if pd.isna(val): return 0.0
    val = str(val).replace('IQD', '').replace(',', '').strip()
    try: return float(val)
    except: return 0.0

def process_percentage(val):
    if pd.isna(val): return 0.0
    if isinstance(val, (int, float)):
        if 0 <= val <= 1: return val * 100
        return float(val)
    val = str(val).replace('%', '').strip()
    try: return float(val)
    except: return 0.0

# --- Load Static Main Data ---
@st.cache_data
def load_main_data():
    try:
        df = pd.read_excel("main_database.xlsx")
    except FileNotFoundError:
        st.warning("يرجى التأكد من وجود ملف main_database.xlsx في نفس المجلد.")
        return pd.DataFrame()
        
    df.columns = df.columns.str.strip()
    
    if 'اسم الطالب' not in df.columns:
        st.error(f"⚠️ لم أتمكن من العثور على 'اسم الطالب' في القاعدة الرئيسية. الأعمدة: {list(df.columns)}")
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
        st.stop()

    if 'اسم الطالب' not in df_daily.columns:
        st.error(f"⚠️ الملف اليومي لا يحتوي على 'اسم الطالب'. الأعمدة: {list(df_daily.columns)}")
        st.stop()
        
    # --- Dynamic Column Detection (القسم vs القسم والمرحلة) ---
    dept_col_main = 'القسم' if 'القسم' in df_main.columns else ('القسم والمرحلة' if 'القسم والمرحلة' in df_main.columns else None)
    dept_col_daily = 'المرحلة' if 'المرحلة' in df_daily.columns else ('القسم' if 'القسم' in df_daily.columns else None)
    
    if dept_col_daily:
        detected_departments = df_daily[dept_col_daily].dropna().unique()
        st.success(f"✅ تم اكتشاف الأقسام: {', '.join(detected_departments)}")

    df_daily['Match_Key_Daily'] = df_daily['اسم الطالب'].apply(clean_arabic_name)
    
    currency_col = 'المبلغ المتبقي' if 'المبلغ المتبقي' in df_daily.columns else df_daily.columns[-2]
    percent_col = 'نسبة الدفع' if 'نسبة الدفع' in df_daily.columns else df_daily.columns[-1]
    
    df_daily['المبلغ المتبقي_رقم'] = df_daily[currency_col].apply(clean_currency)
    df_daily['نسبة الدفع_رقم'] = df_daily[percent_col].apply(process_percentage)
    df_daily['نسبة الدفع_معروضة'] = df_daily['نسبة الدفع_رقم'].apply(lambda x: f"{int(x) if x.is_integer() else x}%")
    
    # --- Advanced Fuzzy Matching (Lowered Cutoff for 4-5 words) ---
    main_names_list = df_main['Match_Key'].tolist()
    
    def find_best_match(daily_name):
        if daily_name in main_names_list: return daily_name
        # Lowered to 0.75 to allow for missing 5th names or commas that were replaced by spaces
        matches = difflib.get_close_matches(daily_name, main_names_list, n=1, cutoff=0.75)
        return matches[0] if matches else None

    with st.spinner('جاري مطابقة الأسماء المعقدة...'):
        df_daily['Matched_DB_Name'] = df_daily['Match_Key_Daily'].apply(find_best_match)

    # Columns to merge from Main DB
    cols_to_merge = ['Match_Key', 'الرمز', 'رقم القاعة']
    if dept_col_main: cols_to_merge.append(dept_col_main)

    merged_df = pd.merge(df_daily, df_main[cols_to_merge], 
                         left_on='Matched_DB_Name', right_on='Match_Key', 
                         how='left')
    
    # --- Filters ---
    st.sidebar.markdown("---")
    st.sidebar.header("تصفية البيانات")
    
    selected_stage = []
    if dept_col_daily:
        stage_options = merged_df[dept_col_daily].dropna().unique()
        selected_stage = st.sidebar.multiselect("اختر القسم / المرحلة:", options=stage_options)

    payment_threshold = st.sidebar.slider("نسبة الدفع أقل من أو تساوي (%)", 0, 100, 100)
    filtered_df = merged_df[merged_df['نسبة الدفع_رقم'] <= payment_threshold]
    if selected_stage:
        filtered_df = filtered_df[filtered_df[dept_col_daily].isin(selected_stage)]

    # --- Dashboards ---
    st.markdown("### 📈 ملخص الإحصائيات")
    total_debt = filtered_df['المبلغ المتبقي_رقم'].sum()
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("إجمالي الطلاب", f"{len(filtered_df)}")
    c2.metric("إجمالي الديون", f"{total_debt:,.0f} IQD")
    c3.metric("متوسط نسبة الدفع", f"{filtered_df['نسبة الدفع_رقم'].mean():.1f}%")
    c4.metric("غير مطابقين", f"{filtered_df['رقم القاعة'].isna().sum()}")

    col_chart1, col_chart2 = st.columns(2)
    with col_chart1:
        fig1 = px.histogram(filtered_df, x="نسبة الدفع_رقم", nbins=10, title="توزيع نسب الدفع", color_discrete_sequence=['#2ecc71'])
        st.plotly_chart(fig1, use_container_width=True)
    with col_chart2:
        if dept_col_daily:
            debt_pie = filtered_df.groupby(dept_col_daily)['المبلغ المتبقي_رقم'].sum().reset_index()
            fig2 = px.pie(debt_pie, values='المبلغ المتبقي_رقم', names=dept_col_daily, title="حجم الديون حسب القسم", hole=0.4)
            st.plotly_chart(fig2, use_container_width=True)

    # --- Tables ---
    st.markdown("### 📝 جدول المطابقة النهائي")
    desired_cols = ['التسلسل', 'اسم الطالب', dept_col_daily, dept_col_main, 'رقم القاعة', currency_col, 'نسبة الدفع_معروضة']
    display_cols = [col for col in desired_cols if col and col in filtered_df.columns]
    
    missing_matches = filtered_df[filtered_df['رقم القاعة'].isna()]
    if not missing_matches.empty:
        st.error(f"⚠️ يوجد {len(missing_matches)} طالب غير مطابق.")
    
    final_table = filtered_df[display_cols].rename(columns={'نسبة الدفع_معروضة': 'نسبة الدفع'})
    st.dataframe(final_table, use_container_width=True, hide_index=True)

    # --- PROFESSIONAL EXCEL EXPORT ---
    st.markdown("---")
    st.markdown("### 📥 تصدير التقرير الاحترافي")
    
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        workbook = writer.book
        
        # Formats
        title_fmt = workbook.add_format({'bold': True, 'font_size': 16, 'font_color': '#ffffff', 'bg_color': '#2c3e50', 'align': 'center', 'valign': 'vcenter', 'border': 1})
        header_fmt = workbook.add_format({'bold': True, 'bg_color': '#ecf0f1', 'align': 'center', 'valign': 'vcenter', 'border': 1})
        cell_fmt = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1})
        money_fmt = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'num_format': '#,##0'})
        
        department_title = f"تقرير الأقسام: {', '.join(selected_stage) if selected_stage else 'جميع الأقسام'}"

        # 1. Matched Sheet
        matched_df = final_table.dropna(subset=['رقم القاعة'])
        matched_df.to_excel(writer, index=False, sheet_name='الطلاب المطابقين', startrow=2)
        ws_match = writer.sheets['الطلاب المطابقين']
        ws_match.right_to_left()
        ws_match.merge_range(0, 0, 0, len(matched_df.columns)-1, department_title, title_fmt)
        ws_match.set_row(0, 30)
        
        for col_num, value in enumerate(matched_df.columns):
            ws_match.write(1, col_num, value, header_fmt)
            ws_match.set_column(col_num, col_num, 20) # Set column width
            
        for row_num in range(len(matched_df)):
            for col_num in range(len(matched_df.columns)):
                val = matched_df.iloc[row_num, col_num]
                fmt = money_fmt if pd.api.types.is_numeric_dtype(type(val)) and val > 1000 else cell_fmt
                if pd.isna(val): val = ""
                ws_match.write(row_num + 2, col_num, val, fmt)

        # 2. Missing Sheet (الغير مطابقين)
        if not missing_matches.empty:
            missing_export = missing_matches[[c for c in ['التسلسل', 'اسم الطالب', dept_col_daily, currency_col] if c in missing_matches.columns]]
            missing_export.to_excel(writer, index=False, sheet_name='الطلاب غير المطابقين', startrow=2)
            ws_miss = writer.sheets['الطلاب غير المطابقين']
            ws_miss.right_to_left()
            ws_miss.merge_range(0, 0, 0, len(missing_export.columns)-1, "⚠️ الطلاب غير المطابقين (يرجى المراجعة يدوياً)", title_fmt)
            
            for col_num, value in enumerate(missing_export.columns):
                ws_miss.write(1, col_num, value, header_fmt)
                ws_miss.set_column(col_num, col_num, 25)
            for row_num in range(len(missing_export)):
                for col_num in range(len(missing_export.columns)):
                    val = missing_export.iloc[row_num, col_num]
                    ws_miss.write(row_num + 2, col_num, "" if pd.isna(val) else val, cell_fmt)

        # 3. Dashboard / Summary Sheet
        ws_sum = workbook.add_worksheet('الملخص والإحصائيات')
        ws_sum.right_to_left()
        ws_sum.set_column('A:B', 25)
        ws_sum.merge_range('A1:B1', 'ملخص التقرير المالي', title_fmt)
        
        stats = [
            ('إجمالي الطلاب', len(filtered_df)),
            ('إجمالي الديون (IQD)', total_debt),
            ('الطلاب المطابقين', len(matched_df)),
            ('الطلاب غير المطابقين', len(missing_matches))
        ]
        for row, (label, val) in enumerate(stats, start=1):
            ws_sum.write(row, 0, label, header_fmt)
            ws_sum.write(row, 1, val, money_fmt if 'الديون' in label else cell_fmt)

        # Create Native Excel Chart if department column exists
        if dept_col_daily and len(filtered_df) > 0:
            dept_data = filtered_df.groupby(dept_col_daily)['المبلغ المتبقي_رقم'].sum().reset_index()
            ws_sum.write(7, 0, 'القسم', header_fmt)
            ws_sum.write(7, 1, 'إجمالي الدين', header_fmt)
            for i, row in dept_data.iterrows():
                ws_sum.write(8 + i, 0, row[dept_col_daily], cell_fmt)
                ws_sum.write(8 + i, 1, row['المبلغ المتبقي_رقم'], money_fmt)
                
            chart = workbook.add_chart({'type': 'pie'})
            chart.add_series({
                'name': 'الديون حسب القسم',
                'categories': ['الملخص والإحصائيات', 8, 0, 8 + len(dept_data) - 1, 0],
                'values':     ['الملخص والإحصائيات', 8, 1, 8 + len(dept_data) - 1, 1],
            })
            chart.set_title({'name': 'توزيع الديون'})
            ws_sum.insert_chart('D2', chart)

    st.download_button(
        label="📥 تحميل التقرير الاحترافي (Excel) مع التنسيق والجداول",
        data=buffer.getvalue(),
        file_name="Advanced_Students_Report.xlsx",
        mime="application/vnd.ms-excel"
    )

else:
    st.info("👈 يرجى رفع ملف التسديدات اليومي من القائمة الجانبية للبدء.")
