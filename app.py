import streamlit as st
import pandas as pd
import plotly.express as px
import re
import io
import difflib

# --- Page Configuration & Modern Styling ---
st.set_page_config(page_title="نظام مطابقة التسديدات", layout="wide", initial_sidebar_state="expanded")

st.markdown("""
    <style>
        body, .stApp { direction: rtl; text-align: right; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background-color: #f4f7f6; }
        
        /* Modern Floating Cards for Metrics */
        div[data-testid="metric-container"] {
            background-color: #ffffff;
            border: 1px solid #e0e6ed;
            padding: 20px;
            border-radius: 15px;
            box-shadow: 0 4px 10px rgba(0,0,0,0.04);
            border-right: 5px solid #3498db;
            transition: transform 0.2s ease-in-out;
        }
        div[data-testid="metric-container"]:hover {
            transform: translateY(-5px);
            box-shadow: 0 8px 15px rgba(0,0,0,0.1);
        }
        
        /* Stylish Buttons */
        .stButton>button { 
            border-radius: 10px; 
            font-weight: bold; 
            background: linear-gradient(135deg, #2ecc71, #27ae60);
            color: white;
            border: none;
            padding: 10px 20px;
            transition: all 0.3s ease;
        }
        .stButton>button:hover {
            background: linear-gradient(135deg, #27ae60, #2ecc71);
            box-shadow: 0 4px 10px rgba(46, 204, 113, 0.4);
        }
        
        h1, h2, h3 { color: #2c3e50; font-weight: 700; }
        
        /* Modern Selectbox/Input focus */
        .stSelectbox div[data-baseweb="select"] {
            border-radius: 10px;
        }
    </style>
""", unsafe_allow_html=True)

# --- HEADER WITH LOGO ---
col_logo, col_title = st.columns([1, 11])
with col_logo:
    st.image("https://uoturath.edu.iq/wp-content/uploads/2025/03/shield-1.png", width=80)
with col_title:
    st.title("✨ لوحة البيانات الذكية لمطابقة التسديدات المالية")
st.markdown("---")

# --- Helper Functions ---
def clean_arabic_name(name):
    if pd.isna(name): return ""
    name = str(name).strip()
    name = re.sub(r'[,،_.-]', ' ', name)
    name = re.sub(r'[أإآ]', 'ا', name)
    name = re.sub(r'ة', 'ه', name)
    name = re.sub(r'ى', 'ي', name)
    name = re.sub(r'\s+', ' ', name).strip()
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

def categorize_payment(pct):
    if pct >= 100: return "مكتمل (100%)"
    elif pct > 0: return "دفع جزئي"
    else: return "لم يدفع (0%)"

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
st.sidebar.header("📂 رفع البيانات اليومية")
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
        
    # --- Dynamic Column Detection ---
    dept_col_main = 'القسم' if 'القسم' in df_main.columns else ('القسم والمرحلة' if 'القسم والمرحلة' in df_main.columns else None)
    dept_col_daily = 'المرحلة' if 'المرحلة' in df_daily.columns else ('القسم' if 'القسم' in df_daily.columns else None)

    df_daily['Match_Key_Daily'] = df_daily['اسم الطالب'].apply(clean_arabic_name)
    
    currency_col = 'المبلغ المتبقي' if 'المبلغ المتبقي' in df_daily.columns else df_daily.columns[-2]
    percent_col = 'نسبة الدفع' if 'نسبة الدفع' in df_daily.columns else df_daily.columns[-1]
    
    df_daily['المبلغ المتبقي_رقم'] = df_daily[currency_col].apply(clean_currency)
    df_daily['نسبة الدفع_رقم'] = df_daily[percent_col].apply(process_percentage)
    df_daily['نسبة الدفع_معروضة'] = df_daily['نسبة الدفع_رقم'].apply(lambda x: f"{int(x) if x.is_integer() else x}%")
    df_daily['فئة الدفع'] = df_daily['نسبة الدفع_رقم'].apply(categorize_payment)
    
    # --- Advanced Fuzzy Matching ---
    main_names_list = df_main['Match_Key'].tolist()
    
    def find_best_match(daily_name):
        if daily_name in main_names_list: return daily_name
        matches = difflib.get_close_matches(daily_name, main_names_list, n=1, cutoff=0.75)
        return matches[0] if matches else None

    with st.spinner('جاري تحليل ومطابقة البيانات...'):
        df_daily['Matched_DB_Name'] = df_daily['Match_Key_Daily'].apply(find_best_match)

    cols_to_merge = ['Match_Key', 'الرمز', 'رقم القاعة']
    if dept_col_main: cols_to_merge.append(dept_col_main)

    merged_df = pd.merge(df_daily, df_main[cols_to_merge], 
                         left_on='Matched_DB_Name', right_on='Match_Key', 
                         how='left')
    
    # --- Interactive Filters ---
    st.sidebar.markdown("---")
    st.sidebar.header("🎯 أدوات التصفية المتقدمة")
    
    selected_stage = []
    if dept_col_daily:
        stage_options = merged_df[dept_col_daily].dropna().unique()
        selected_stage = st.sidebar.multiselect("📚 اختر القسم / المرحلة:", options=stage_options)

    selected_hall = []
    if 'رقم القاعة' in merged_df.columns:
        hall_options = sorted([x for x in merged_df['رقم القاعة'].unique() if pd.notna(x)], key=lambda x: str(x))
        selected_hall = st.sidebar.multiselect("🚪 اختر رقم القاعة:", options=hall_options)

    payment_threshold = st.sidebar.slider("💰 نسبة الدفع أقل من أو تساوي (%)", 0, 100, 100)
    
    # Apply filters
    filtered_df = merged_df[merged_df['نسبة الدفع_رقم'] <= payment_threshold]
    if selected_stage:
        filtered_df = filtered_df[filtered_df[dept_col_daily].isin(selected_stage)]
    if selected_hall:
        filtered_df = filtered_df[filtered_df['رقم القاعة'].isin(selected_hall)]

    # --- Metrics Dashboards ---
    st.markdown("### 📊 ملخص الأداء المالي")
    total_debt = filtered_df['المبلغ المتبقي_رقم'].sum()
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("إجمالي الطلاب", f"{len(filtered_df)}")
    c2.metric("إجمالي الديون (المتبقي)", f"{total_debt:,.0f} IQD")
    c3.metric("متوسط نسبة الدفع", f"{filtered_df['نسبة الدفع_رقم'].mean():.1f}%")
    
    missing_matches = filtered_df[filtered_df['رقم القاعة'].isna()]
    c4.metric("الحالات غير المطابقة", f"{len(missing_matches)}")

    # --- Modern Interactive Charts ---
    st.markdown("<br>", unsafe_allow_html=True)
    col_chart1, col_chart2 = st.columns(2)
    
    color_map = {"مكتمل (100%)": "#2ecc71", "دفع جزئي": "#f1c40f", "لم يدفع (0%)": "#e74c3c"}
    
    with col_chart1:
        pie_data = filtered_df['فئة الدفع'].value_counts().reset_index()
        pie_data.columns = ['فئة الدفع', 'العدد']
        
        fig1 = px.pie(pie_data, values='العدد', names='فئة الدفع', 
                      title="موقف تسديدات الطلاب (الإجمالي)", 
                      hole=0.5, 
                      color='فئة الدفع', color_discrete_map=color_map)
        fig1.update_traces(textposition='inside', textinfo='percent+label', hoverinfo='label+percent+value')
        fig1.update_layout(showlegend=False, margin=dict(t=50, b=0, l=0, r=0))
        st.plotly_chart(fig1, use_container_width=True)
        
    with col_chart2:
        if dept_col_daily:
            # NEW CHART: Payment Status Distribution by Department (Stacked Bar)
            payment_by_dept = filtered_df.groupby([dept_col_daily, 'فئة الدفع']).size().reset_index(name='العدد')
            
            fig2 = px.bar(payment_by_dept, x='العدد', y=dept_col_daily, color='فئة الدفع',
                          orientation='h', 
                          title="توزيع حالات التسديد حسب القسم", 
                          color_discrete_map=color_map, 
                          text_auto=True)
            
            fig2.update_layout(xaxis_title="عدد الطلاب", yaxis_title="", 
                               barmode='stack', margin=dict(t=50, b=0, l=0, r=0),
                               legend_title_text='حالة الدفع')
            st.plotly_chart(fig2, use_container_width=True)

    st.markdown("---")

    # --- Smart Dropdown for Unmatched Students ---
    st.markdown("### ⚠️ مراجعة الطلاب غير المطابقين")
    if not missing_matches.empty:
        unmatched_names = missing_matches['اسم الطالب'].tolist()
        
        selected_unmatched = st.selectbox(
            "ابحث أو اختر اسم الطالب غير المطابق لعرض تفاصيله:", 
            options=["-- اختر طالب من القائمة --"] + unmatched_names
        )
        
        if selected_unmatched != "-- اختر طالب من القائمة --":
            student_info = missing_matches[missing_matches['اسم الطالب'] == selected_unmatched].iloc[0]
            st.info(f"**الاسم:** {student_info['اسم الطالب']} | **القسم المُدخل:** {student_info.get(dept_col_daily, 'غير محدد')} | **المبلغ المتبقي:** {student_info['المبلغ المتبقي_رقم']:,.0f} IQD")
    else:
        st.success("🎉 ممتاز! جميع الطلاب في هذا التقرير مطابقون لقاعدة البيانات الرئيسية.")

    st.markdown("---")

    # --- Final Data Table ---
    st.markdown("### 📝 جدول المطابقة النهائي")
    desired_cols = ['التسلسل', 'اسم الطالب', dept_col_daily, dept_col_main, 'رقم القاعة', currency_col, 'نسبة الدفع_معروضة']
    display_cols = [col for col in desired_cols if col and col in filtered_df.columns]
    
    final_table = filtered_df[display_cols].rename(columns={'نسبة الدفع_معروضة': 'نسبة الدفع'})
    
    st.dataframe(final_table, use_container_width=True, hide_index=True, height=400)

    # --- PROFESSIONAL EXCEL EXPORT (Row 1 Header Fix) ---
    st.markdown("<br>", unsafe_allow_html=True)
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        workbook = writer.book
        
        header_fmt = workbook.add_format({'bold': True, 'bg_color': '#2c3e50', 'font_color': '#ffffff', 'align': 'center', 'valign': 'vcenter', 'border': 1})
        cell_fmt = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1})
        money_fmt = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'num_format': '#,##0'})
        title_fmt = workbook.add_format({'bold': True, 'font_size': 16, 'bg_color': '#ecf0f1', 'align': 'center', 'valign': 'vcenter', 'border': 1})

        # 1. Matched Sheet (Headers start exactly at row 0)
        matched_df = final_table.dropna(subset=['رقم القاعة'])
        matched_df.to_excel(writer, index=False, sheet_name='الطلاب المطابقين', startrow=0)
        ws_match = writer.sheets['الطلاب المطابقين']
        ws_match.right_to_left()
        
        for col_num, value in enumerate(matched_df.columns):
            ws_match.write(0, col_num, value, header_fmt) # Header on row 0
            ws_match.set_column(col_num, col_num, 20)
            
        for row_num in range(len(matched_df)):
            for col_num in range(len(matched_df.columns)):
                val = matched_df.iloc[row_num, col_num]
                fmt = money_fmt if pd.api.types.is_numeric_dtype(type(val)) and val > 1000 else cell_fmt
                if pd.isna(val): val = ""
                ws_match.write(row_num + 1, col_num, val, fmt) # Data starts at row 1

        # 2. Missing Sheet
        if not missing_matches.empty:
            missing_export = missing_matches[[c for c in ['التسلسل', 'اسم الطالب', dept_col_daily, currency_col] if c in missing_matches.columns]]
            missing_export.to_excel(writer, index=False, sheet_name='الطلاب غير المطابقين', startrow=0)
            ws_miss = writer.sheets['الطلاب غير المطابقين']
            ws_miss.right_to_left()
            
            for col_num, value in enumerate(missing_export.columns):
                ws_miss.write(0, col_num, value, header_fmt) # Header on row 0
                ws_miss.set_column(col_num, col_num, 25)
            for row_num in range(len(missing_export)):
                for col_num in range(len(missing_export.columns)):
                    val = missing_export.iloc[row_num, col_num]
                    ws_miss.write(row_num + 1, col_num, "" if pd.isna(val) else val, cell_fmt) # Data starts at row 1

        # 3. Dashboard / Summary Sheet
        ws_sum = workbook.add_worksheet('الملخص والإحصائيات')
        ws_sum.right_to_left()
        ws_sum.set_column('A:B', 25)
        ws_sum.merge_range('A1:B1', 'ملخص التقرير المالي', title_fmt)
        
        stats = [
            ('إجمالي الطلاب', len(filtered_df)),
            ('إجمالي الديون المتبقية (IQD)', total_debt),
            ('الطلاب المطابقين', len(matched_df)),
            ('الطلاب غير المطابقين', len(missing_matches))
        ]
        for row, (label, val) in enumerate(stats, start=1):
            ws_sum.write(row, 0, label, header_fmt)
            ws_sum.write(row, 1, val, money_fmt if 'الديون' in label else cell_fmt)

        # Excel Chart - Showing counts by payment category
        if len(pie_data) > 0:
            ws_sum.write(7, 0, 'فئة الدفع', header_fmt)
            ws_sum.write(7, 1, 'عدد الطلاب', header_fmt)
            for i, row in pie_data.iterrows():
                ws_sum.write(8 + i, 0, row['فئة الدفع'], cell_fmt)
                ws_sum.write(8 + i, 1, row['العدد'], cell_fmt)
                
            chart = workbook.add_chart({'type': 'pie'})
            chart.add_series({
                'name': 'حالات الدفع',
                'categories': ['الملخص والإحصائيات', 8, 0, 8 + len(pie_data) - 1, 0],
                'values':     ['الملخص والإحصائيات', 8, 1, 8 + len(pie_data) - 1, 1],
            })
            chart.set_title({'name': 'توزيع حالات الدفع'})
            ws_sum.insert_chart('D2', chart)

    st.download_button(
        label="📥 تصدير التقرير الاحترافي (Excel)",
        data=buffer.getvalue(),
        file_name="Smart_Matched_Report.xlsx",
        mime="application/vnd.ms-excel",
        use_container_width=True
    )

else:
    st.info("👈 يرجى رفع ملف التسديدات اليومي من القائمة الجانبية للبدء.")
