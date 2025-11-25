
# -*- coding: utf-8 -*-
# Arabic CX Dashboard (3 Dimensions) — Streamlit
# Files expected in the same folder:
#   - MN.csv                          ← raw survey data
#   - Digital_Data_tables.xlsx         ← lookup/metadata tables
#
# Run:
#   streamlit run Arabic_Dashboard.py

import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import io, re
from datetime import datetime
from pathlib import Path
USER_KEYS = {
    "بلدية رأس الخيمة": {
        "password": st.secrets["users"]["MN"],
        "role": "center",
        "file": "MN.csv"
    },
    "محاكم رأس الخيمة": {
        "password": st.secrets["users"]["CR"],
        "role": "center",
        "file": "CR.csv"
    },
    "النيابة العامة في رأس الخيمة": {
        "password": st.secrets["users"]["PR"],
        "role": "center",
        "file": "PR.csv"
    },
    "دائرة التنمية الاقتصادية": {
        "password": st.secrets["users"]["EC"],
        "role": "center",
        "file": "EC.csv"
    },
    "جمارك رأس الخيمة": {
        "password": st.secrets["users"]["CU"],
        "role": "center",
        "file": "CU.csv"
    },
    "هيئة حماية البيئة والتنمية": {
        "password": st.secrets["users"]["EN"],
        "role": "center",
        "file": "EN.csv"
    },
    "الأمانة العامة للمجلس التنفيذي": {
        "password": st.secrets["users"]["GS"],
        "role": "admin",
        "file": "Centers_Master.csv"   # غيّر الاسم إذا كان لديك ملف مختلف للإدارة العامة
    }
}
# =========================================================
# إعداد الصفحة + اتجاه RTL
# =========================================================
st.set_page_config(page_title="تقرير تجربة المتعامل في الخدمات الرقمية 2025", layout="wide")
PASTEL = px.colors.qualitative.Pastel

# شعار أعلى الصفحة (استبدل بالرابط المناسب إذا رغبت)
LOGO_URL = "https://raw.githubusercontent.com/gsecrak/rakdcx2025/main/assets/mini_header3.png"
st.markdown(f"""
    <div style="text-align:center; margin-top:-40px;">
        <img src="{LOGO_URL}" alt="Logo" style="width:950px; max-width:95%; height:auto;">
    </div>
    <hr style="margin-top:20px; margin-bottom:10px;">
""", unsafe_allow_html=True)

# اتجاه عربي وخط مناسب
st.markdown("""
    <style>
        html, body, [class*="css"] {
            direction: rtl;
            text-align: right;
            font-family: "Tajawal","Cairo","Segoe UI";
        }

        /* شريط التبويبات: اتجاه عربي وثابت في اليمين */
        .stTabs [data-baseweb="tab-list"] {
            direction: rtl !important;          /* أول تبويب يكون عند اليمين */
            display: flex !important;
            justify-content: flex-start !important;  /* يبدأ من اليمين */
            width: 100% !important;             /* يأخذ عرض السطر بالكامل */
        }

        /* نص كل تبويب يكون RTL ومحاذى يمين */
        .stTabs [data-baseweb="tab"] > div {
            direction: rtl !important;
            text-align: right !important;
        }

        .stDownloadButton, .stButton > button {
            font-weight: 600;
        }
    </style>
""", unsafe_allow_html=True)


# قاموس الجهات والملفات
ENTITIES = {
    "بلدية رأس الخيمة": {
        "csv": "MN.csv",
        "xlsx": "Data_tables_MN.xlsx",
    },
    "محاكم رأس الخيمة": {
        "csv": "CR.csv",
        "xlsx": "Data_tables_CR.xlsx",
    },
    "النيابة العامة في رأس الخيمة": {
        "csv": "PR.csv",
        "xlsx": "Data_tables_PR.xlsx",
    },
    "دائرة التنمية الاقتصادية": {
        "csv": "EC.csv",
        "xlsx": "Data_tables_EC.xlsx",
    },
    "جمارك رأس الخيمة": {
        "csv": "CU.csv",
        "xlsx": "Data_tables_CU.xlsx",
    },
    "هيئة حماية البيئة والتنمية": {
        "csv": "EN.csv",
        "xlsx": "Data_tables_EN.xlsx",
    },
     # 👇 جهة الأدمن (تجميع كل الجهات)
    "الأمانة العامة للمجلس التنفيذي": {
        "csv": "Centers_Master.csv",         # لن نستخدمها
        "xlsx": "Data_tables_MASTER.xlsx",        # لن نستخدمها
        #"aggregated": True,  # علامة أنها جهة تجميع
    },
}

# =========================================================
# تحميل البيانات
# =========================================================
# تحميل البيانات مع إضافة سطر المعاني (Arabic Labels)
# =========================================================
#@st.cache_data(show_spinner=False)
def load_data(csv_name: str, xlsx_name: str):
    # البيانات الرئيسية
    df = pd.read_csv(csv_name, encoding="utf-8", low_memory=False)
    df.columns = [c.strip().upper() for c in df.columns]
    df.columns = [c.replace('DIM', 'Dim') for c in df.columns]

    # الجداول الوصفية
    lookup_catalog = {}
    xls_path = Path(xlsx_name)
    if xls_path.exists():
        xls = pd.ExcelFile(xls_path)
        for sheet in xls.sheet_names:
            tbl = pd.read_excel(xls, sheet_name=sheet)
            tbl.columns = [str(c).strip().upper() for c in tbl.columns]
            lookup_catalog[sheet.strip().upper()] = tbl

        # 🔹 محاولة جلب ورقة "Questions" لإضافة معاني الأعمدة
        qsheet_key = next((k for k in lookup_catalog.keys() if "QUESTION" in k), None)
        if qsheet_key:
            qtbl = lookup_catalog[qsheet_key]
            qtbl.columns = [str(c).strip().upper() for c in qtbl.columns]
            code_col = next((c for c in qtbl.columns if "DIM" in c or "QUESTION" in c or "CODE" in c), None)
            ar_col = next((c for c in qtbl.columns if "ARAB" in c), None)
            if code_col and ar_col:
                code_to_arabic = dict(zip(qtbl[code_col].astype(str).str.upper(),
                                          qtbl[ar_col].astype(str)))
                # إنشاء سطر معاني عربية للأعمدة الموجودة في df
                arabic_row = []
                for c in df.columns:
                    key = c.strip().upper()
                    arabic_row.append(code_to_arabic.get(key, ""))
                # إدراج السطر العربي في الأعلى (اختياري)
                arabic_df = pd.DataFrame([arabic_row], columns=df.columns)
                # df = pd.concat([arabic_df, df], ignore_index=True)

    return df, lookup_catalog
def load_all_entities():
    """تحميل بيانات جميع الجهات ودمجها في DataFrame واحد مع عمود ENTITY_NAME"""
    frames = []
    combined_lookup = {}

    for name, conf in ENTITIES.items():
        # نتخطى جهة الأدمن نفسها
        if conf.get("aggregated"):
            continue

        csv_name = conf["csv"]
        xlsx_name = conf["xlsx"]
        df_i, lookup_i = load_data(csv_name, xlsx_name)

        if df_i is None or df_i.empty:
            continue

        df_i = df_i.copy()
        # نضيف عمود باسم الجهة
        df_i.insert(0, "ENTITY_NAME", name)

        frames.append(df_i)

        # دمج lookup_catalog (نأخذ أول نسخة من كل شيت)
        for k, v in lookup_i.items():
            if k not in combined_lookup:
                combined_lookup[k] = v

    if frames:
        df_all = pd.concat(frames, ignore_index=True)
    else:
        df_all = pd.DataFrame()

    return df_all, combined_lookup


def series_to_percent(vals: pd.Series):
    vals = pd.to_numeric(vals, errors="coerce").dropna()
    if len(vals) == 0:
        return np.nan
    mx = vals.max()
    if mx <= 5:   # سلم 1-5
        return ((vals - 1) / 4 * 100).mean()
    elif mx <= 10:  # سلم 1-10
        return ((vals - 1) / 9 * 100).mean()
    else:        # بيانات جاهزة كنسب
        return vals.mean()

def detect_nps(df: pd.DataFrame):
    cand_cols = [c for c in df.columns if ("NPS" in c.upper()) or ("RECOMMEND" in c.upper()) or ("NETPROMOTER" in c.upper())]
    if not cand_cols:
        return np.nan, 0, 0, 0, None
    col = cand_cols[0]
    s = pd.to_numeric(df[col], errors="coerce").dropna()
    if len(s) == 0:
        return np.nan, 0, 0, 0, col
    promoters = (s >= 9).sum()
    passives  = ((s >= 7) & (s <= 8)).sum()
    detract   = (s <= 6).sum()
    total     = len(s)
    promoters_pct = promoters / total * 100
    passives_pct  = passives  / total * 100
    detract_pct   = detract   / total * 100
    nps = promoters_pct - detract_pct
    return nps, promoters_pct, passives_pct, detract_pct, col

def autodetect_metric_cols(df: pd.DataFrame):
    # نحاول التعرف على أعمدة CSAT و CES (قد تكون Dim6.1/Dim6.2 أو CSAT/CES أو FEES)
    cols_upper = {c.upper(): c for c in df.columns}
    # CSAT
    csat_candidates = [c for c in df.columns if "CSAT" in c.upper()] 

    csat_col = csat_candidates[0] if csat_candidates else None

    #  Fees
    ces_candidates = [c for c in df.columns if "FEES" in c.upper()]
    ces_col = ces_candidates[0] if ces_candidates else None

    # NPS
    nps_candidates = [c for c in df.columns if "NPS" in c.upper()] 
    nps_col = nps_candidates[0] if nps_candidates else None

    return csat_col, ces_col, nps_col

# اختيار الجهة من الشريط الجانبي
st.sidebar.title("اختيار الجهة")
selected_entity = st.sidebar.selectbox("الرجاء اختيار الجهة:", list(ENTITIES.keys()))

# إعداد إعدادات الجهة المختارة
entity_conf = ENTITIES[selected_entity]       # هنا نأخذ ملفات الجهة (csv/xlsx)
user_conf   = USER_KEYS[selected_entity]      # وهنا نأخذ كلمة السر والدور

correct_password = user_conf["password"]      # ← من USER_KEYS
is_aggregated    = entity_conf.get("aggregated", False)

# إدخال كلمة المرور
password_input = st.sidebar.text_input(
    "🔐 كلمة المرور للجهة المختارة:",
    type="password",
    help="لن يتم عرض التقرير إلا بعد إدخال كلمة المرور الصحيحة."
)

# التحقق من كلمة المرور قبل تحميل البيانات
if not password_input:
    st.warning("⚠️ الرجاء إدخال كلمة المرور لعرض تقرير الجهة المختارة.")
    st.stop()
elif password_input != correct_password:
    st.error("❌ كلمة المرور غير صحيحة. الرجاء المحاولة مرة أخرى.")
    st.stop()
else:
    # بعد التحقق من كلمة المرور
    if is_aggregated:
        # جهة الأدمن: تحميل كل الجهات معًا
        df, lookup_catalog = load_all_entities()
    else:
        # جهة عادية: تحميل ملف واحد فقط
        csv_name = entity_conf["csv"]
        xlsx_name = entity_conf["xlsx"]
        df, lookup_catalog = load_data(csv_name, xlsx_name)

    st.sidebar.markdown(f"**الجهة الحالية:** {selected_entity}")

# عناوين عربية للفلاتر
ARABIC_FILTER_TITLES = {
    "AGE": "العمر",
    "SERVICE": "الخدمة",
    "LANGUAGE": "اللغة",
    "PERIOD": "الفترة",
    "CHANNEL": "القناة",
    "ENTITY_NAME": "الجهة"
}

st.sidebar.header("🎛️ الفلاتر")
# نحاول تطبيق ترجمة للأبعاد/المتغيرات باستخدام جداول الـ lookup إذا وجدت
df_filtered = df.copy()

# سنعرض فلاتر لأكثر الحقول شيوعًا؛ ويمكن التوسع تلقائيًا إذا وُجدت جداول مطابقة في الـ lookup
candidate_filter_cols = []
# أبعاد ديموغرافية أو وصفية شائعة
common_keys = ["Language", "SERVICE", "AGE", "PERIOD", "CHANNEL", "ENTITY_NAME"]
candidate_filter_cols = [c for c in df.columns if any(k in c.upper() for k in common_keys)]

# وظيفة لتطبيق جدول lookup إذا توفّر باسم العمود

# وظيفة لتطبيق جدول lookup (تربط تلقائيًا بين الأكواد والأسماء العربية)
def apply_lookup(column_name: str, s: pd.Series) -> pd.Series:
    key = column_name.strip().upper()

    # 1) تطابق تام بين اسم العمود واسم الشيت
    match_key = None
    for k in lookup_catalog.keys():
        if k.strip().upper() == key:
            match_key = k
            break

    # 2) إذا لم نجد تطابق تام → نحاول تطابق جزئي
    if match_key is None:
        for k in lookup_catalog.keys():
            if key in k or k in key:
                match_key = k
                break

    if match_key is None:
        return s

    tbl = lookup_catalog[match_key].copy()
    tbl.columns = [str(c).strip().upper() for c in tbl.columns]
    if len(tbl.columns) < 2:
        return s

    code_col = tbl.columns[0]
    name_col = tbl.columns[1]
    map_dict = dict(zip(tbl[code_col].astype(str), tbl[name_col].astype(str)))
    return s.astype(str).map(map_dict).fillna(s)


# نُحضّر نسخة مترجمة للعرض في الفلاتر
df_filtered_display = df_filtered.copy()
for col in candidate_filter_cols:
    df_filtered_display[col] = apply_lookup(col, df_filtered[col])

with st.sidebar.expander("تطبيق/إزالة الفلاتر"):
    applied_filters = {}

    for col in candidate_filter_cols:

        # طبّق الترجمة على القيم داخل الفلاتر
        df_filtered[col] = apply_lookup(col, df_filtered[col])

        # الخيارات المتاحة للقائمة
        options = df_filtered_display[col].dropna().unique().tolist()
        options_sorted = sorted(options, key=lambda x: str(x))
        default = options_sorted

        # اختيار العنوان العربي إذا كان موجودًا
        label = ARABIC_FILTER_TITLES.get(col.upper(), col)

        # عرض الفلتر باستخدام الاسم العربي
        sel = st.multiselect(label, options_sorted, default=default)

        applied_filters[col] = sel


# تطبيق الفلاتر
for col, selected in applied_filters.items():
    if selected:
        df_filtered = df_filtered[df_filtered[col].isin(selected)]

# البيانات النهائية للعرض
df_view = df_filtered.copy()

# عناوين عربية للأعمدة التي نريد رسم توزيعها
AR_DIST_TITLES = {
    "AGE": "العمر",
    "SERVICE": "الخدمة",
    "LANGUAGE": "اللغة",
    "PERIOD": "الفترة",
    "CHANNEL": "القناة",
}

# =========================================================
# التبويبات
# =========================================================
if is_aggregated:
    # جهة الأدمن: نضيف تبويب المقارنات
    tab_data, tab_sample, tab_kpis, tab_dimensions, tab_services, tab_pareto, tab_admin = st.tabs([
        "📁 البيانات",
        "📈 توزيع العينة",
        "📊 المؤشرات",
        "🧩 الأبعاد",
        "📋 الخدمات",
        "💬 المزعجات",
        "📊 المقارنات بين الجهات"
    ])
else:
    # باقي الجهات: بدون تبويب المقارنات
    tab_data, tab_sample, tab_kpis, tab_dimensions, tab_services, tab_pareto = st.tabs([
        "📁 البيانات",
        "📈 توزيع العينة",
        "📊 المؤشرات",
        "🧩 الأبعاد",
        "📋 الخدمات",
        "💬 المزعجات"
    ])
    
# =========================================================
# تبويب البيانات + تنزيل
# =========================================================
with tab_data:
    # st.subheader("📁 البيانات")
    st.dataframe(df_view, use_container_width=True)
    ts = datetime.now().strftime("%Y-%m-%d_%H%M")
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df_view.to_excel(writer, index=False, sheet_name="Filtered_Data")
    st.download_button("📥 تنزيل البيانات (Excel)", data=buf.getvalue(),
                       file_name=f"Filtered_Data_{ts}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# =========================================================
# تبويب توزيع العينة
# =========================================================
with tab_sample:
    st.subheader("📈 توزيع العينة")
    total = len(df_view)
    st.markdown(
        f"### 🧮 إجمالي الردود: <span style='color:#1E88E5;'>{total:,}</span>",
        unsafe_allow_html=True,
    )

    # نوع الرسم
    chart_type = st.radio(
        "📊 نوع الرسم", ["مخطط أعمدة", "مخطط دائري"], index=0, horizontal=True
    )

    # خيار عرض العدد أو النسبة أو كليهما
    display_mode = st.radio(
        "📋 طريقة العرض:",
        ["العدد فقط", "النسبة فقط", "العدد + النسبة"],
        horizontal=True,
        index=1,
    )

    # الأعمدة التي نريد لها توزيع (5 فقط)
    dist_base = ["AGE", "SERVICE", "LANGUAGE", "PERIOD", "CHANNEL"]
    dist_cols = [c for c in candidate_filter_cols if c.upper() in dist_base]

    for col in dist_cols:
        if col not in df_view.columns:
            continue

        counts = (
            df_view[col]
            .value_counts(dropna=True)
            .reset_index()
        )
        counts.columns = [col, "Count"]
        if counts.empty:
            continue

        counts["Percentage"] = (
            counts["Count"] / counts["Count"].sum() * 100
        )

        # تحديد العمود المستخدم حسب اختيار المستخدم
        if display_mode == "العدد فقط":
            y_col = "Count"
            y_label = "عدد الردود"
            text_col = counts["Count"].astype(str)
        elif display_mode == "النسبة فقط":
            y_col = "Percentage"
            y_label = "النسبة (%)"
            text_col = counts["Percentage"].map("{:.1f}%".format)
        else:  # العدد + النسبة
            y_col = "Count"
            y_label = "عدد الردود"
            text_col = counts.apply(
                lambda x: f"{x['Count']} ({x['Percentage']:.1f}%)", axis=1
            )

        # عنوان عربي للمخطط
        col_key = col.upper()
        col_label = AR_DIST_TITLES.get(col_key, col)
        title_text = f"توزيع {col_label}"

        st.markdown(f"### {title_text}")

        # ===== رسم المخطط =====
        if chart_type == "مخطط أعمدة":
            fig = px.bar(
                counts,
                x=col,
                y=y_col,
                text=text_col,
                color=col,
                color_discrete_sequence=PASTEL,
                title=title_text,
            )
            fig.update_traces(textposition="outside")
            fig.update_layout(
                title={"text": title_text, "x": 0.5},
                xaxis_title="الفئة",
                yaxis_title=y_label,
                showlegend=False,
                height=500,
            )
            fig.update_layout(title_font_size=20)
            st.plotly_chart(fig, use_container_width=True)

        else:  # === مخطط دائري ===
            fig = px.pie(
                counts,
                names=col,
                values="Count",
                hole=0.3,
                color=col,
                color_discrete_sequence=PASTEL,
                title=title_text,
            )

            fig.update_layout(
                title={"text": title_text, "x": 0.5},
                height=500,
            )

            fig.update_layout(title_font_size=20)
            
            # تعديل النص حسب اختيار المستخدم
            if display_mode == "العدد فقط":
                fig.update_traces(
                    textposition="inside",
                    texttemplate="%{label}<br>%{value}",
                )
            elif display_mode == "النسبة فقط":
                fig.update_traces(
                    textposition="inside",
                    texttemplate="%{label}<br>%{percent:.1%}",
                )
            else:  # كلاهما
                fig.update_traces(
                    textposition="inside",
                    texttemplate="%{label}<br>%{value} (%{percent:.1%})",
                )

            st.plotly_chart(fig, use_container_width=True)

        # ===== جدول ملخص تحت المخطط =====
        st.dataframe(
            counts[[col, "Count", "Percentage"]]
            .rename(
                columns={
                    col: "الفئة",
                    "Count": "عدد الردود",
                    "Percentage": "النسبة (%)",
                }
            )
            .style.format({"النسبة (%)": "{:.1f}%"}),
            use_container_width=True,
            hide_index=True,
        )
        st.markdown("---")

# =========================================================
# تبويب المؤشرات (CSAT / CES / NPS)
# =========================================================
with tab_kpis:
    st.subheader("📊 مؤشرات الأداء الرئيسية")
    csat_col, ces_col, nps_col = autodetect_metric_cols(df_view)

    # حساب CSAT
    csat = series_to_percent(df_view.get(csat_col, pd.Series(dtype=float))) if csat_col else np.nan
    # حساب CES/Value
    ces  = series_to_percent(df_view.get(ces_col,  pd.Series(dtype=float))) if ces_col else np.nan
    # حساب NPS
    nps, p_pct, s_pct, d_pct, nps_col = detect_nps(df_view)

    def color_label(score, metric_type):
        if metric_type in ["CSAT", "CES"]:
            if pd.isna(score):           return "#bdc3c7", "غير متاح"
            if score < 70:               return "#FF6B6B", "ضعيف جدًا"
            elif score < 80:             return "#FFD93D", "بحاجة إلى تحسين"
            elif score < 90:             return "#6BCB77", "جيد"
            else:                        return "#4D96FF", "ممتاز"
        else:  # NPS
            if pd.isna(score):           return "#bdc3c7", "غير متاح"
            if score < 0:                return "#FF6B6B", "ضعيف جدًا"
            elif score < 30:             return "#FFD93D", "ضعيف"
            elif score < 60:             return "#6BCB77", "جيد"
            else:                        return "#4D96FF", "ممتاز"

    def gauge(score, title, metric_type):
        color, label = color_label(score, metric_type)
        axis_range = [0, 100] if metric_type in ["CSAT", "CES"] else [-100, 100]
        steps = (
            [{'range': [0, 70], 'color': '#FF6B6B'},
             {'range': [70, 80], 'color': '#FFD93D'},
             {'range': [80, 90], 'color': '#6BCB77'},
             {'range': [90, 100], 'color': '#4D96FF'}]
            if metric_type in ["CSAT", "CES"]
            else [{'range': [-100, 0], 'color': '#FF6B6B'},
                  {'range': [0, 30], 'color': '#FFD93D'},
                  {'range': [30, 60], 'color': '#6BCB77'},
                  {'range': [60, 100], 'color': '#4D96FF'}]
        )
        fig = go.Figure(go.Indicator(
            mode="gauge+number",
            value=0 if pd.isna(score) else float(score),
            number={'suffix': "٪" if metric_type != "NPS" else ""},
            title={'text': title, 'font': {'size': 18}},
            gauge={
                'axis': {'range': axis_range},
                'bar': {'color': color},
                'steps': steps
            }
        ))
        fig.update_layout(height=300, margin=dict(l=30, r=30, t=60, b=30))
        return fig, label

    c1, c2, c3 = st.columns(3)
    fig1, lab1 = gauge(csat, "السعادة العامة (CSAT)", "CSAT")
    fig2, lab2 = gauge(ces,  "القيمة مقابل الجهد/التكلفة (CES/Value)", "CES")
    fig3, lab3 = gauge(nps,  "صافي نقاط الترويج (NPS)", "NPS")
    c1.plotly_chart(fig1, use_container_width=True)
    c1.markdown(f"**التفسير:** {lab1}")
    if csat_col: c1.caption(f"المصدر: {csat_col}")
    c2.plotly_chart(fig2, use_container_width=True)
    c2.markdown(f"**التفسير:** {lab2}")
    if ces_col: c2.caption(f"المصدر: {ces_col}")
    c3.plotly_chart(fig3, use_container_width=True)
    c3.markdown(f"**التفسير:** {lab3}")
    if nps_col: c3.caption(f"المصدر: {nps_col}")
    c3.markdown(f"المروجون: {p_pct:.1f}% | المحايدون: {s_pct:.1f}% | المعارضون: {d_pct:.1f}%", unsafe_allow_html=True)

    # =========================================================
    # 🎨 وسيلتا الإيضاح (Legends)
    # =========================================================
    legend_html = """
    <div style='background-color:#f9f9f9;border:1px solid #ddd;border-radius:10px;padding:12px;margin-top:15px;'>
        <h4 style='margin-bottom:8px;'>🎨 وسيلة الإيضاح — السعادة / القيمة</h4>
        🔴 أقل من 70٪ — ضعيف جدًا<br>
        🟡 من 70 إلى أقل من 80٪ — بحاجة إلى تحسين<br>
        🟢 من 80 إلى أقل من 90٪ — جيد<br>
        🔵 90٪ فأكثر — ممتاز
    </div>
    <div style='background-color:#f9f9f9;border:1px solid #ddd;border-radius:10px;padding:12px;margin-top:10px;'>
        <h4 style='margin-bottom:8px;'>🎯 وسيلة الإيضاح — صافي نقاط الترويج (NPS)</h4>
        🔴 أقل من 0 — ضعيف جدًا (عدد المعارضين أكبر من المروجين)<br>
        🟡 من 0 إلى أقل من 30 — ضعيف (رضا محدود)<br>
        🟢 من 30 إلى أقل من 60 — جيد (رضا عام)<br>
        🔵 60 فأكثر — ممتاز (ولاء مرتفع جدًا)
    </div>
    """
    st.markdown(legend_html, unsafe_allow_html=True)

# =========================================================
# تبويب الأبعاد (3 أبعاد فقط)
# =========================================================
with tab_dimensions:
    # st.subheader("🧩 تحليل الأبعاد")

    # نبحث عن الأعمدة التي تبدأ بـ "DimX." (الأسئلة الفرعية داخل كل بعد)
    dim_subcols = [c for c in df_view.columns if re.match(r"Dim\d+\.", str(c).strip())]
    if not dim_subcols:
        st.info("لا توجد أعمدة فرعية للأبعاد (مثل Dim1.1 أو Dim2.3).")
    else:
        # بناء المتوسط لكل بعد رئيسي (Dim1, Dim2, Dim3...) — نلتقط ما هو متاح
        main_dim_map = {}
        for i in range(1, 6):
            sub = [c for c in df_view.columns if str(c).startswith(f"Dim{i}.")]
            if sub:
                main_dim_map[f"Dim{i}"] = df_view[sub].apply(pd.to_numeric, errors="coerce").mean(axis=1)

        # إنشاء ملخص بنتائج الأبعاد
        summary = []
        for dim, series in main_dim_map.items():
            score = series_to_percent(series)
            summary.append({"Dimension": dim, "Score": score})

        dims = pd.DataFrame(summary).dropna()
        if dims.empty:
            st.info("لا توجد نتائج كافية للأبعاد.")
        else:
            # ترتيب الأبعاد حسب الرقم (Dim1, Dim2...)
            dims["Order"] = dims["Dimension"].str.extract(r"(\d+)").astype(float)
            dims = dims.sort_values("Order").reset_index(drop=True)

            # 🔄 استبدال أسماء الأبعاد من ورقة "Questions" في ملف Excel إذا وُجدت
            for sheet_name in lookup_catalog.keys():
                if "QUESTION" in sheet_name:  # يلتقط Question أو Questions
                    qtbl = lookup_catalog[sheet_name].copy()
                    qtbl.columns = [str(c).strip().upper() for c in qtbl.columns]

                    # محاولة تحديد عمود الأكواد وعمود الاسم العربي
                    code_col = next((c for c in qtbl.columns if any(k in c for k in ["DIM", "CODE", "QUESTION", "ID"])), None)
                    name_col = next((c for c in qtbl.columns if any(k in c for k in ["ARABIC", "NAME", "LABEL", "TEXT"])), None)

                    if code_col and name_col:
                        def _norm(s):
                            return s.astype(str).str.upper().str.replace(r"\s+", "", regex=True)

                        code_series = _norm(qtbl[code_col])
                        name_series = qtbl[name_col].astype(str)
                        map_dict = dict(zip(code_series, name_series))

                        # استبدال الأكواد بالأسماء العربية
                        dims["Dimension"] = (
                            _norm(dims["Dimension"])
                            .map(map_dict)
                            .fillna(dims["Dimension"])
                        )
                    break  # توقف بعد العثور على الورقة المطابقة

            # تصنيف الأبعاد حسب التقييم
            def cat(score):
                if score < 70:  return "🔴 ضعيف"
                elif score < 80: return "🟡 متوسط"
                elif score < 90: return "🟢 جيد"
                else:            return "🔵 ممتاز"
            dims["Category"] = dims["Score"].apply(cat)

            # رسم بياني للأبعاد
            fig = px.bar(
                dims, x="Dimension", y="Score", text="Score", color="Category",
                color_discrete_map={
                    "🔴 ضعيف": "#FF6B6B",
                    "🟡 متوسط": "#FFD93D",
                    "🟢 جيد":   "#6BCB77",
                    "🔵 ممتاز": "#4D96FF"
                },
                title="<span style='font-size:28px; font-weight:bold;'>📊 تحليل متوسط الأبعاد</span>"
            )
            fig.update_traces(texttemplate="%{text:.1f}%", textposition="outside")
            fig.update_layout(
                title={
                    'text': "<span style='font-size:22px; font-weight:bold;'>📊 تحليل متوسط الأبعاد</span>",
                    'x': 0.5,  # المنتصف
                    'xanchor': 'center'
                },
                yaxis=dict(range=[0, 100]),
                xaxis_title="البعد",
                yaxis_title="النسبة المئوية (%)"
            )

            st.plotly_chart(fig, use_container_width=True)
            # وسيلة الإيضاح ثنائية اللغة
            st.markdown(
                """
                **🗂️ وسيلة الإيضاح:**
                - 🔴 أقل من 70٪ — ضعيف الأداء  
                - 🟡 من 70٪ إلى أقل من 80٪ — متوسط  
                - 🟢 من 80٪ إلى أقل من 90٪ — جيد  
                - 🔵 90٪ فأكثر — ممتاز  
                """,
            unsafe_allow_html=True)
            # عرض جدول الأبعاد
            st.dataframe(
                dims[["Dimension", "Score"]]
                .rename(columns={"Dimension": "البعد", "Score": "النسبة (%)"})
                .style.format({"النسبة (%)": "{:.1f}%"}),
                use_container_width=True,
                hide_index=True
            )

# =========================================================
# تبويب الخدمات
# =========================================================
with tab_services:
    st.subheader("📋 تحليل الخدمات")
    if "SERVICE" not in df_view.columns:
        st.warning("⚠️ لا توجد بيانات خاصة بالخدمات (SERVICE).")
    else:
        csat_col, ces_col, _ = autodetect_metric_cols(df_view)
        work = df_view.copy()
        if csat_col:
            work["سعادة (%)"] = (pd.to_numeric(work[csat_col], errors="coerce") - 1) * 25
        if ces_col:
            work["قيمة (%)"] = (pd.to_numeric(work[ces_col], errors="coerce") - 1) * 25

        # NPS لكل خدمة إن وُجد عمود NPS
        nps_cols = [c for c in df_view.columns if "NPS" in c.upper() or "RECOMMEND" in c.upper()]
        if nps_cols:
            work["NPS_VAL"] = pd.to_numeric(work[nps_cols[0]], errors="coerce")
            nps_summary = []
            for svc, g in work.groupby("SERVICE"):
                s = g["NPS_VAL"].dropna()
                if len(s) == 0:
                    nps_summary.append((svc, np.nan))
                    continue
                promoters = (s >= 9).sum()
                detractors = (s <= 6).sum()
                total = len(s)
                nps_value = ((promoters - detractors) / total) * 100
                nps_summary.append((svc, nps_value))
            nps_df = pd.DataFrame(nps_summary, columns=["SERVICE", "NPS (%)"])
        else:
            nps_df = pd.DataFrame(columns=["SERVICE", "NPS (%)"])

        # حساب المتوسط وعدد الردود
        agg_dict = {}
        if "سعادة (%)" in work.columns: agg_dict["سعادة (%)"] = "mean"
        if "قيمة (%)" in work.columns:  agg_dict["قيمة (%)"]  = "mean"
        if csat_col:                   agg_dict[csat_col]    = "count"

        if not agg_dict:
            st.info("لا توجد أعمدة كافية لحساب مؤشرات الخدمة.")
        else:
            summary = work.groupby("SERVICE").agg(agg_dict).reset_index()
            if csat_col and csat_col in summary.columns:
                summary.rename(columns={csat_col: "عدد الردود"}, inplace=True)

            # دمج NPS
            if not nps_df.empty:
                summary = summary.merge(nps_df, on="SERVICE", how="left")

            # ترجمة اسم الخدمة عبر lookup (إن وجد sheet باسم SERVICE)
            if "SERVICE" in lookup_catalog:
                tbl = lookup_catalog["SERVICE"].copy()
                tbl.columns = [str(c).strip().upper() for c in tbl.columns]
                code_col = next((c for c in tbl.columns if "CODE" in c or "SERVICE" in c), None)
                ar_col   = next((c for c in tbl.columns if ("ARABIC" in c) or ("SERVICE2" in c)), None)
                if code_col and ar_col:
                    name_map = dict(zip(tbl[code_col].astype(str), tbl[ar_col].astype(str)))
                    summary["SERVICE"] = summary["SERVICE"].astype(str).map(name_map).fillna(summary["SERVICE"])

            # فلترة إلى خدمات بعدد ردود كافٍ (اختياري: 30)
            if "عدد الردود" in summary.columns:
                summary = summary[summary["عدد الردود"] >= 30]

            # ترتيب
            sort_key = "سعادة (%)" if "سعادة (%)" in summary.columns else ("قيمة (%)" if "قيمة (%)" in summary.columns else None)
            if sort_key:
                summary = summary.sort_values(sort_key, ascending=False)

            # ✅ تلوين الخلايا في الجدول (السعادة والقيمة فقط)
            def color_cells(val):
                try:
                    v = float(val)
                    if v < 70:
                        color = "#FF6B6B"  # أحمر
                    elif v < 80:
                        color = "#FFD93D"  # أصفر
                    elif v < 90:
                        color = "#6BCB77"  # أخضر
                    else:
                        color = "#4D96FF"  # أزرق
                    return f"background-color:{color};color:black"
                except:
                    return ""

            # 📋 إعداد الـ format حسب الأعمدة المتوفرة
            format_dict = {}
            if "سعادة (%)" in summary.columns:
                format_dict["سعادة (%)"] = "{:.1f}%"
            if "قيمة (%)" in summary.columns:
                format_dict["قيمة (%)"] = "{:.1f}%"
            if "NPS (%)" in summary.columns:
                format_dict["NPS (%)"] = "{:.1f}%"
            if "عدد الردود" in summary.columns:
                format_dict["عدد الردود"] = "{:,.0f}"

            subset_cols = [c for c in ["سعادة (%)", "قيمة (%)"] if c in summary.columns]

            # 📋 عرض الجدول
            styled_table = (
                summary.style
                .format(format_dict)
                .applymap(color_cells, subset=subset_cols)
            )
            st.dataframe(styled_table, use_container_width=True)

            # 🛈 ملاحظة توضيحية باللغتين
            st.markdown(
                """
                **ℹ️ ملاحظة:**  
                يتم عرض الخدمات التي تحتوي على **30 ردًا أو أكثر فقط** لضمان دقة النتائج.  
                """
            )

            # رسم مقارنة (سعادة/قيمة)
            if "سعادة (%)" in summary.columns or "قيمة (%)" in summary.columns:
                melted = summary.melt(
                    id_vars=["SERVICE"],
                    value_vars=[v for v in ["سعادة (%)", "قيمة (%)"] if v in summary.columns],
                    var_name="المؤشر",
                    value_name="القيمة"
                )

                fig = px.bar(
                    melted,
                    x="SERVICE",
                    y="القيمة",
                    color="المؤشر",
                    barmode="group",
                    text="القيمة",
                    color_discrete_sequence=PASTEL,
                    title="مقارنة مؤشري السعادة والقيمة حسب الخدمة"
                )
                fig.update_traces(texttemplate="%{text:.1f}%", textposition="outside")

                fig.update_layout(
                    yaxis=dict(range=[0, 100]),
                    xaxis_title="الخدمة",
                    yaxis_title="النسبة (%)"
                )

                # 🔥 تكبير العنوان + توسيطه
                fig.update_layout(
                    title={
                        "text": "📊 مقارنة مؤشري السعادة والقيمة حسب الخدمة",
                        "x": 0.5,
                        "y": 0.95,
                        "xanchor": "center",
                        "yanchor": "top"
                    },
                    title_font_size=20
                )
                st.plotly_chart(fig, use_container_width=True)


# =========================================================
# 💬 تحليل أسباب عدم الرضا (Most_Unsat) بطريقة Pareto
# =========================================================
with tab_pareto:
    st.subheader("💬 تحليل المزعجات")

    unsat_col = next((c for c in df_view.columns if "MOST_UNSAT" in c.upper()), None)
    if not unsat_col:
        st.warning("⚠️ لم يتم العثور على العمود Most_Unsat في البيانات.")
    else:
        data_unsat = df_view[[unsat_col]].copy()
        data_unsat.columns = ["Comment"]
        data_unsat["Comment"] = data_unsat["Comment"].astype(str).str.strip()

        # استثناء الإجابات العامة
        exclude_terms = ["", " ", "لا يوجد", "لايوجد", "لاشيء", "لا شيء",
                         "none", "no", "nothing", "nil", "جيد", "ممتاز", "ok", "تمام", "great"]
        data_unsat = data_unsat[~data_unsat["Comment"].str.lower().isin([t.lower() for t in exclude_terms])]
        data_unsat = data_unsat[data_unsat["Comment"].apply(lambda x: len(x.split()) >= 2)]

        if data_unsat.empty:
            st.info("لا توجد ملاحظات نصية كافية بعد التنظيف.")
        else:
            # 🔹 تصنيف التعليقات حسب المحاور
            themes = {
                "السرعة / الأداء": ["بطء", "تأخير", "انتظار", "delay", "slow", "زمن", "وقت"],
                "التطبيق / المنصة": ["تطبيق", "app", "منصة", "system", "موقع", "بوابة", "صفحة"],
                "الإجراءات / الخطوات": ["إجراء", "اجراء", "عملية", "process", "خطوات", "مراحل", "نموذج"],
                "الرسوم / الدفع": ["رسوم", "دفع", "fee", "تكلفة", "سداد", "pay"],
                "التواصل / الدعم الفني": ["رد", "تواصل", "اتصال", "support", "response", "مساندة", "مساعدة"],
                "الوضوح / المعلومات": ["معلومة", "إيضاح", "clarity", "instructions", "بيانات", "شرح"],
                "الأمان / الدخول": ["كلمة مرور", "دخول", "login", "تحقق", "أمان"]
            }

            def classify_text(txt):
                t = txt.lower()
                for theme, keywords in themes.items():
                    if any(k.lower() in t for k in keywords):
                        return theme
                return "غير مصنّف"

            data_unsat["المحور"] = data_unsat["Comment"].apply(classify_text)
            data_unsat = data_unsat[data_unsat["المحور"] != "غير مصنّف"]

            # 🔢 تجميع حسب المحور + ضمّ التعليقات بفاصل "/"
            summary = data_unsat.groupby("المحور").agg({
                "Comment": lambda x: " / ".join(x.tolist())
            }).reset_index()

            summary["عدد الملاحظات"] = summary["Comment"].apply(lambda x: len(x.split("/")))
            summary = summary.sort_values("عدد الملاحظات", ascending=False).reset_index(drop=True)
            summary["النسبة (%)"] = summary["عدد الملاحظات"] / summary["عدد الملاحظات"].sum() * 100
            summary["النسبة التراكمية (%)"] = summary["النسبة (%)"].cumsum()
            summary["اللون"] = np.where(summary["النسبة التراكمية (%)"] <= 80, "#E74C3C", "#BDC3C7")

            # ✅ أول بند يتجاوز 80٪ يكون أحمر أيضًا
            if not summary[summary["النسبة التراكمية (%)"] > 80].empty:
                first_above = summary[summary["النسبة التراكمية (%)"] > 80].index[0]
                summary.loc[first_above, "اللون"] = "#E74C3C"

            # 🧾 عرض الجدول
            st.dataframe(
                summary[["المحور", "عدد الملاحظات", "النسبة (%)", "النسبة التراكمية (%)", "Comment"]]
                .rename(columns={"Comment": "التعليقات (مجمعة)"}).style.format({
                    "النسبة (%)": "{:.1f}%",
                    "النسبة التراكمية (%)": "{:.1f}%"
                }),
                use_container_width=True,
                hide_index=True
            )

            # 📊 رسم Pareto
            fig = go.Figure()
            fig.add_bar(
                x=summary["المحور"],
                y=summary["عدد الملاحظات"],
                marker_color=summary["اللون"],
                name="عدد الملاحظات"
            )
            fig.add_scatter(
                x=summary["المحور"],
                y=summary["النسبة التراكمية (%)"],
                yaxis="y2",
                mode="lines+markers+text",
                name="النسبة التراكمية (%)",
                text=[f"{v:.1f}%" for v in summary["النسبة التراكمية (%)"]],
                textposition="top center",
                line=dict(color="#2E86DE", width=3)
            )
            fig.update_layout(
                title={
                "text": "📊 تحليل باريتو - المحاور الرئيسية",
                    "x": 0.5,
                    "y": 0.95,
                    "xanchor": "center",
                    "yanchor": "top"
                },
                title_font_size=20,
                xaxis=dict(title="المحور", tickangle=-15),
                yaxis=dict(title="عدد الملاحظات"),
                yaxis2=dict(title="النسبة التراكمية (%)", overlaying="y", side="right", range=[0, 110]),
                height=600,
                bargap=0.3,
                legend=dict(orientation="h", y=-0.2)
            )
            st.plotly_chart(fig, use_container_width=True)
            # 📥 زر تنزيل جدول Pareto (Excel)
            pareto_buffer = io.BytesIO()
            with pd.ExcelWriter(pareto_buffer, engine="openpyxl") as writer:
                summary.to_excel(writer, index=False, sheet_name="Pareto_Results")

            pareto_buffer.seek(0)  # لضمان القراءة من البداية

            st.download_button(
                label="📥 تنزيل جدول Pareto (Excel)",
                data=pareto_buffer.getvalue(),
                file_name=f"Pareto_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
# =========================================================
# تبويب خاص للأمانة العامة: مقارنة الجهات في مؤشرات الأداء والأبعاد
# =========================================================
if is_aggregated:
    with tab_admin:
        st.subheader("📊 مقارنة الجهات في مؤشرات الأداء الرئيسية والأبعاد")

        # تأكد أن عمود اسم الجهة موجود
        if "ENTITY_NAME" not in df_view.columns:
            st.warning("⚠️ لا يوجد عمود ENTITY_NAME في البيانات المجمّعة.")
        else:
            # كشف أعمدة المقاييس (CSAT / CES / NPS) تلقائياً
            csat_col, ces_col, nps_col = autodetect_metric_cols(df_view)

            work = df_view.copy()

            # 🔹 تجميع مؤشرات الأداء الرئيسية لكل جهة
            rows = []
            for ent, g in work.groupby("ENTITY_NAME"):
                row = {"الجهة": ent, "عدد الردود": len(g)}

                if csat_col:
                    row["سعادة (%)"] = series_to_percent(g[csat_col])
                if ces_col:
                    row["قيمة (%)"] = series_to_percent(g[ces_col])

                nps_val, _, _, _, _ = detect_nps(g)
                row["NPS (%)"] = nps_val

                rows.append(row)

            kpi_df = pd.DataFrame(rows)

            if kpi_df.empty:
                st.info("لا توجد بيانات كافية لحساب مؤشرات الأداء الرئيسية.")
            else:
                # 📋 عرض الجدول مع تنسيقات بسيطة
                kpi_display = kpi_df.copy()
                for c in ["سعادة (%)", "قيمة (%)", "NPS (%)"]:
                    if c in kpi_display.columns:
                        kpi_display[c] = kpi_display[c].round(1)

                st.markdown("### 🔍 مقارنة مؤشرات الأداء الرئيسية حسب الجهة")
                st.dataframe(
                    kpi_display.style.format({
                        "سعادة (%)": "{:.1f}%",
                        "قيمة (%)": "{:.1f}%",
                        "NPS (%)": "{:.1f}%",
                        "عدد الردود": "{:,.0f}"
                    }),
                    use_container_width=True,
                    hide_index=True
                )

                # 📊 رسم مقارنة سعادة/قيمة/NPS حسب الجهة
                metric_cols = [c for c in ["سعادة (%)", "قيمة (%)", "NPS (%)"] if c in kpi_df.columns]
                if metric_cols:
                    melted_kpi = kpi_df.melt(
                        id_vars=["الجهة"],
                        value_vars=metric_cols,
                        var_name="المؤشر",
                        value_name="القيمة"
                    )

                    fig_kpi = px.bar(
                        melted_kpi,
                        x="الجهة",
                        y="القيمة",
                        color="المؤشر",
                        barmode="group",
                        text="القيمة",
                        title="مقارنة مؤشرات الأداء الرئيسية حسب الجهة"
                    )
                    fig_kpi.update_traces(texttemplate="%{text:.1f}%", textposition="outside")
                    fig_kpi.update_layout(
                        yaxis=dict(range=[0, 100]),
                        xaxis_title="الجهة",
                        yaxis_title="النسبة (%)",
                        legend=dict(orientation="h", y=-0.2)
                    )
                    st.plotly_chart(fig_kpi, use_container_width=True)
if is_aggregated:
    with tab_admin:
        st.subheader("📊 المقارنات بين الجهات")

        # هنا تضع منطق المقارنات في KPIs والأبعاد
        # مثال بسيط:
        if "ENTITY_NAME" not in df_view.columns:
            st.warning("⚠️ لا يوجد عمود ENTITY_NAME في البيانات.")
        else:
            st.write("هنا سيتم عرض مقارنة مؤشرات الأداء الرئيسية بين الجهات...")
            # ضع كود الجداول والرسوم الخاصة بالمقارنات
# =========================================================
# تبويب الأدمن: مقارنة الجهات حسب الأبعاد الرئيسية (Dim1, Dim2, ...)
# =========================================================
if is_aggregated:
    with tab_admin:
        st.subheader("📊 مقارنة الجهات حسب الأبعاد الرئيسية (Dim1, Dim2, ...)")

        if "ENTITY_NAME" not in df_view.columns:
            st.warning("⚠️ لا يوجد عمود ENTITY_NAME في البيانات المجمّعة.")
        else:
            # 1️⃣ نبحث عن الأعمدة الفرعية للأبعاد DimX. (مثل Dim1.1 / Dim2.3)
            dim_subcols = [c for c in df_view.columns if re.match(r"Dim\d+\.", str(c).strip())]

            if not dim_subcols:
                st.info("لا توجد أعمدة فرعية للأبعاد (مثل Dim1.1 أو Dim2.3) في البيانات.")
            else:
                # نستخرج أرقام الأبعاد الرئيسية الموجودة (1,2,3,...) من DimX.Y
                main_ids = sorted({
                    int(re.match(r"Dim(\d+)\.", str(c).strip()).group(1))
                    for c in dim_subcols
                    if re.match(r"Dim(\d+)\.", str(c).strip())
                })

                # 2️⃣ حساب نتيجة كل بُعد رئيسي لكل جهة
                rows = []
                for ent, g in df_view.groupby("ENTITY_NAME"):
                    for i in main_ids:
                        # كل الأعمدة الفرعية التي تبدأ بـ Dim{i}.
                        sub = [c for c in g.columns if str(c).startswith(f"Dim{i}.")]
                        if not sub:
                            continue

                        # متوسط الأسئلة الفرعية لهذا البعد
                        dim_series = g[sub].apply(pd.to_numeric, errors="coerce").mean(axis=1)
                        score = series_to_percent(dim_series)

                        rows.append({
                            "الجهة": ent,
                            "Dimension": f"Dim{i}",
                            "Score": score
                        })

                dim_comp_df = pd.DataFrame(rows).dropna(subset=["Score"])

                if dim_comp_df.empty:
                    st.info("لا توجد نتائج كافية لحساب الأبعاد لكل جهة.")
                else:
                    # 3️⃣ استبدال أسماء الأبعاد من ورقة Questions (نفس منطق تبويب الأبعاد)
                    for sheet_name in lookup_catalog.keys():
                        if "QUESTION" in sheet_name.upper():  # Question / Questions
                            qtbl = lookup_catalog[sheet_name].copy()
                            qtbl.columns = [str(c).strip().upper() for c in qtbl.columns]

                            code_col = next(
                                (c for c in qtbl.columns if any(k in c for k in ["DIM", "CODE", "QUESTION", "ID"])),
                                None
                            )
                            name_col = next(
                                (c for c in qtbl.columns if any(k in c for k in ["ARABIC", "NAME", "LABEL", "TEXT"])),
                                None
                            )

                            if code_col and name_col:
                                def _norm(s):
                                    return s.astype(str).str.upper().str.replace(r"\s+", "", regex=True)

                                code_series = _norm(qtbl[code_col])
                                name_series = qtbl[name_col].astype(str)
                                map_dict = dict(zip(code_series, name_series))

                                dim_comp_df["Dimension_label"] = (
                                    _norm(dim_comp_df["Dimension"])
                                    .map(map_dict)
                                    .fillna(dim_comp_df["Dimension"])
                                )
                            else:
                                dim_comp_df["Dimension_label"] = dim_comp_df["Dimension"]

                            break
                    else:
                        # لو ما لقينا ورقة Questions
                        dim_comp_df["Dimension_label"] = dim_comp_df["Dimension"]

                    # تقريب النسب
                    dim_comp_df["Score"] = dim_comp_df["Score"].round(1)

                                     # 4️⃣ عرض جدول المقارنات
                    st.markdown("### 📋 جدول مقارنة الأبعاد الرئيسية بين الجهات")
                    st.dataframe(
                        dim_comp_df[["Dimension", "Dimension_label", "الجهة", "Score"]]
                        .rename(columns={
                            "Dimension": "رمز البعد",
                            "Dimension_label": "اسم البعد",
                            "Score": "النسبة (%)"
                        })
                        .style.format({"النسبة (%)": "{:.1f}%"}),
                        use_container_width=True,
                        hide_index=True
                    )

                    # 5️⃣ رسم جميع الأبعاد مرة واحدة (لكل الجهات)
                    st.markdown("### 📊 مقارنة جميع الأبعاد بين الجهات")

                    # نرتب الأبعاد بالترتيب الرقمي Dim1, Dim2, ...
                    dim_comp_df["Order"] = dim_comp_df["Dimension"].str.extract(r"(\d+)").astype(float)
                    dim_comp_df_sorted = dim_comp_df.sort_values(["Order", "الجهة"])

                    fig_all = px.bar(
                        dim_comp_df_sorted,
                        x="Dimension_label",
                        y="Score",
                        color="الجهة",
                        barmode="group",
                        text="Score",
                        title="مقارنة الجهات في جميع الأبعاد الرئيسية"
                    )
                    fig_all.update_traces(texttemplate="%{text:.1f}%", textposition="outside")
                    fig_all.update_layout(
                        xaxis_title="البعد",
                        yaxis_title="النتيجة (%)",
                        yaxis=dict(range=[0, 100]),
                        xaxis_tickangle=-20,
                        legend=dict(orientation="h", y=-0.25)
                    )
                    st.plotly_chart(fig_all, use_container_width=True)

# =========================================================
# تحسينات شكلية
# =========================================================
st.markdown("""
    <style>
    #MainMenu {visibility: hidden;}
    footer, [data-testid="stFooter"] {opacity: 0.03 !important; height: 1px !important; overflow: hidden !important;}
    </style>
""", unsafe_allow_html=True)















































































