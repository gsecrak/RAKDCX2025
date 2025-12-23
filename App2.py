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
    
    "هيئة حماية البيئة والتنمية": {
        "password": st.secrets["users"]["EN"],
        "role": "center",
        "file": "EN.csv"
    },
    "حكومة رأس الخيمة": {
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

# =========================================================
# ✅ RTL مضبوط: الواجهة RTL، لكن الجداول والرسومات تبقى سليمة (LTR)
# =========================================================
st.markdown("""
<style>
/* =========================================
   1) RTL للواجهة (النصوص/السايدبار/العناوين)
========================================= */
[data-testid="stAppViewContainer"],
[data-testid="block-container"],
[data-testid="stSidebar"]{
    direction: rtl !important;
    text-align: right !important;
    font-family: "Tajawal","Cairo","Segoe UI", sans-serif !important;
}

[data-testid="stAppViewContainer"] *,
[data-testid="stSidebar"] *{
    direction: rtl !important;
    text-align: right !important;
}

/* العناوين والنصوص */
h1, h2, h3, h4, h5, h6, p, label, span, li{
    direction: rtl !important;
    text-align: right !important;
}

/* حقول الإدخال */
[data-baseweb="input"] input,
[data-baseweb="textarea"] textarea{
    direction: rtl !important;
    text-align: right !important;
}

/* القوائم المنسدلة */
[data-baseweb="select"] *{
    direction: rtl !important;
    text-align: right !important;
}

/* =========================================
   2) التبويبات: تبدأ من اليمين بدون قلب الترتيب
   (مهم جدًا: لا نستخدم row-reverse)
========================================= */
.stTabs [data-baseweb="tab-list"]{
    direction: rtl !important;
    display: flex !important;
    justify-content: flex-start !important; /* مع RTL: البداية يمين بدون قلب الترتيب */
    width: 100% !important;
}

.stTabs [data-baseweb="tab"] > div{
    direction: rtl !important;
    text-align: right !important;
}

.stDownloadButton, .stButton > button{
    font-weight: 600;
}

/* =========================================
   3) استثناء الجداول (DataFrame/Table) من RTL
   عشان ترجع واضحة مثل السابق
========================================= */
[data-testid="stDataFrame"],
[data-testid="stDataFrame"] *,
.stDataFrame, .stDataFrame *,
.stTable, .stTable *{
    direction: ltr !important;
    text-align: left !important;
}

/* =========================================
   4) استثناء الرسومات (Plotly) من RTL
   لمنع تداخل نص x مع الأعمدة وتغطية الـ legend
========================================= */
[data-testid="stPlotlyChart"],
[data-testid="stPlotlyChart"] *{
    direction: ltr !important;
    text-align: left !important;
}

/* أحيانًا نصوص SVG تتأثر بمحاذاة RTL */
[data-testid="stPlotlyChart"] svg,
[data-testid="stPlotlyChart"] svg *{
    direction: ltr !important;
    unicode-bidi: plaintext !important;
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
     "هيئة حماية البيئة والتنمية": {
        "csv": "EN.csv",
        "xlsx": "Data_tables_EN.xlsx",
    },
     # 👇 جهة الأدمن (تجميع كل الجهات)
    "حكومة رأس الخيمة": {
        "csv": "Centers_Master.csv",         # لن نستخدمها
        "xlsx": "Data_tables_MASTER.xlsx",        # لن نستخدمها
    },
}

# =========================================================
# تحميل البيانات
# =========================================================
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
                arabic_row = []
                for c in df.columns:
                    key = c.strip().upper()
                    arabic_row.append(code_to_arabic.get(key, ""))
                arabic_df = pd.DataFrame([arabic_row], columns=df.columns)
                # df = pd.concat([arabic_df, df], ignore_index=True)

    return df, lookup_catalog

def load_all_entities():
    """تحميل بيانات جميع الجهات ودمجها في DataFrame واحد مع عمود ENTITY_NAME"""
    frames = []
    combined_lookup = {}

    for name, conf in ENTITIES.items():
        if conf.get("aggregated"):
            continue

        csv_name = conf["csv"]
        xlsx_name = conf["xlsx"]
        df_i, lookup_i = load_data(csv_name, xlsx_name)

        if df_i is None or df_i.empty:
            continue

        df_i = df_i.copy()
        df_i.insert(0, "ENTITY_NAME", name)

        frames.append(df_i)

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
    if mx <= 5:
        return ((vals - 1) / 4 * 100).mean()
    elif mx <= 10:
        return ((vals - 1) / 9 * 100).mean()
    else:
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
    csat_candidates = [c for c in df.columns if "CSAT" in c.upper()]
    csat_col = csat_candidates[0] if csat_candidates else None

    Fees_candidates = [c for c in df.columns if "FEES" in c.upper()]
    Fees_col = Fees_candidates[0] if Fees_candidates else None

    nps_candidates = [c for c in df.columns if "NPS" in c.upper()]
    nps_col = nps_candidates[0] if nps_candidates else None

    return csat_col, Fees_col, nps_col

# ============================
# تحسين شكل الشريط الجانبي
# ============================
st.markdown("""
    <style>
        section[data-testid="stSidebar"] .stSelectbox,
        section[data-testid="stSidebar"] .stTextInput {
            margin-top: -15px !important;
        }
        section[data-testid="stSidebar"] label {
            margin-bottom: -3px !important;
        }
    </style>
""", unsafe_allow_html=True)

# ============================
# اختيار الجهة من الشريط الجانبي
# ============================
st.sidebar.markdown(
    """
    <div style='font-size:20px; font-weight:700; margin-bottom:-10px;'>🏢 اختر الجهة</div>
    """,
    unsafe_allow_html=True
)

selected_entity = st.sidebar.selectbox("", list(ENTITIES.keys()))

entity_conf = ENTITIES[selected_entity]
user_conf   = USER_KEYS[selected_entity]

correct_password = user_conf["password"]
is_admin = (user_conf.get("role") == "admin")

# ============================
# إدخال كلمة المرور
# ============================
st.sidebar.markdown(
    """
    <div style='font-size:20px; font-weight:700; margin-bottom:-10px;'>🔐 كلمة المرور</div>
    """,
    unsafe_allow_html=True
)

password_input = st.sidebar.text_input(
    "",
    type="password",
    help="لن يتم عرض التقرير إلا بعد إدخال كلمة المرور الصحيحة."
)

if not password_input:
    st.warning("⚠️ الرجاء إدخال كلمة المرور لعرض تقرير الجهة المختارة.")
    st.stop()
elif password_input != correct_password:
    st.error("❌ كلمة المرور غير صحيحة. الرجاء المحاولة مرة أخرى.")
    st.stop()
else:
    csv_name = entity_conf["csv"]
    xlsx_name = entity_conf["xlsx"]
    df, lookup_catalog = load_data(csv_name, xlsx_name)
    st.sidebar.markdown(f"**الجهة الحالية:** {selected_entity}")

ARABIC_FILTER_TITLES = {
    "AGE": "العمر",
    "SERVICE": "الخدمة",
    "LANGUAGE": "اللغة",
    "PERIOD": "الفترة",
    "CHANNEL": "القناة",
    "ENTITY_NAME": "الجهة"
}

st.sidebar.header("🎛️ الفلاتر")
df_filtered = df.copy()

common_keys = ["Language", "SERVICE", "AGE", "PERIOD", "CHANNEL", "ENTITY_NAME"]
candidate_filter_cols = [c for c in df.columns if any(k in c.upper() for k in common_keys)]

def apply_lookup(column_name: str, s: pd.Series) -> pd.Series:
    key = column_name.strip().upper()

    match_key = None
    for k in lookup_catalog.keys():
        if k.strip().upper() == key:
            match_key = k
            break

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

df_filtered_display = df_filtered.copy()
for col in candidate_filter_cols:
    df_filtered_display[col] = apply_lookup(col, df_filtered[col])

with st.sidebar.expander("تطبيق/إزالة الفلاتر"):
    applied_filters = {}

    for col in candidate_filter_cols:
        df_filtered[col] = apply_lookup(col, df_filtered[col])

        options = df_filtered_display[col].dropna().unique().tolist()
        options_sorted = sorted(options, key=lambda x: str(x))
        default = options_sorted

        label = ARABIC_FILTER_TITLES.get(col.upper(), col)
        sel = st.multiselect(label, options_sorted, default=default)

        applied_filters[col] = sel

for col, selected in applied_filters.items():
    if selected:
        df_filtered = df_filtered[df_filtered[col].isin(selected)]

df_view = df_filtered.copy()

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
if is_admin:
    tab_data, tab_sample, tab_kpis, tab_dimensions, tab_pareto, tab_admin = st.tabs([
        "📁 البيانات",
        "📈 توزيع العينة",
        "📊 المؤشرات",
        "🧩 الأبعاد",
        "💬 المزعجات",
        "📊 المقارنات بين الجهات"
    ])
else:
    tab_data, tab_sample, tab_kpis, tab_dimensions, tab_pareto = st.tabs([
        "📁 البيانات",
        "📈 توزيع العينة",
        "📊 المؤشرات",
        "🧩 الأبعاد",
        "💬 المزعجات"
    ])

# =========================================================
# تبويب البيانات + تنزيل
# =========================================================
with tab_data:
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

    chart_type = st.radio("📊 نوع الرسم", ["مخطط أعمدة", "مخطط دائري"], index=0, horizontal=True)

    display_mode = st.radio(
        "📋 طريقة العرض:",
        ["العدد فقط", "النسبة فقط", "العدد + النسبة"],
        horizontal=True,
        index=1,
    )

    dist_base = ["AGE", "SERVICE", "LANGUAGE", "PERIOD", "CHANNEL"]
    dist_cols = [c for c in candidate_filter_cols if c.upper() in dist_base]

    for col in dist_cols:
        if col not in df_view.columns:
            continue

        counts = df_view[col].value_counts(dropna=True).reset_index()
        counts.columns = [col, "Count"]
        if counts.empty:
            continue

        counts["Percentage"] = counts["Count"] / counts["Count"].sum() * 100

        if display_mode == "العدد فقط":
            y_col = "Count"
            y_label = "عدد الردود"
            text_col = counts["Count"].astype(str)
        elif display_mode == "النسبة فقط":
            y_col = "Percentage"
            y_label = "النسبة (%)"
            text_col = counts["Percentage"].map("{:.1f}%".format)
        else:
            y_col = "Count"
            y_label = "عدد الردود"
            text_col = counts.apply(lambda x: f"{x['Count']} ({x['Percentage']:.1f}%)", axis=1)

        col_key = col.upper()
        col_label = AR_DIST_TITLES.get(col_key, col)
        title_text = f"توزيع {col_label}"

        st.markdown(f"### {title_text}")

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
        else:
            fig = px.pie(
                counts,
                names=col,
                values="Count",
                hole=0.3,
                color=col,
                color_discrete_sequence=PASTEL,
                title=title_text,
            )
            fig.update_layout(title={"text": title_text, "x": 0.5}, height=500)
            fig.update_layout(title_font_size=20)

            if display_mode == "العدد فقط":
                fig.update_traces(textposition="inside", texttemplate="%{label}<br>%{value}")
            elif display_mode == "النسبة فقط":
                fig.update_traces(textposition="inside", texttemplate="%{label}<br>%{percent:.1%}")
            else:
                fig.update_traces(textposition="inside", texttemplate="%{label}<br>%{value} (%{percent:.1%})")

            st.plotly_chart(fig, use_container_width=True)

        st.dataframe(
            counts[[col, "Count", "Percentage"]]
            .rename(columns={col: "الفئة", "Count": "عدد الردود", "Percentage": "النسبة (%)"})
            .style.format({"النسبة (%)": "{:.1f}%"}),
            use_container_width=True,
            hide_index=True,
        )
        st.markdown("---")

# =========================================================
# تبويب المؤشرات (CSAT / Fees / NPS)
# =========================================================
with tab_kpis:
    st.subheader("📊 مؤشرات الأداء الرئيسية")
    csat_col, Fees_col, nps_col = autodetect_metric_cols(df_view)

    csat = series_to_percent(df_view.get(csat_col, pd.Series(dtype=float))) if csat_col else np.nan
    Fees  = series_to_percent(df_view.get(Fees_col,  pd.Series(dtype=float))) if Fees_col else np.nan
    nps, p_pct, s_pct, d_pct, nps_col = detect_nps(df_view)

    def color_label(score, metric_type):
        if metric_type in ["CSAT", "Fees"]:
            if pd.isna(score):           return "#bdc3c7", "غير متاح"
            if score < 70:               return "#FF6B6B", "ضعيف جدًا"
            elif score < 80:             return "#FFD93D", "بحاجة إلى تحسين"
            elif score < 90:             return "#6BCB77", "جيد"
            else:                        return "#4D96FF", "ممتاز"
        else:
            if pd.isna(score):           return "#bdc3c7", "غير متاح"
            if score < 0:                return "#FF6B6B", "ضعيف جدًا"
            elif score < 30:             return "#FFD93D", "ضعيف"
            elif score < 60:             return "#6BCB77", "جيد"
            else:                        return "#4D96FF", "ممتاز"

    def gauge(score, title, metric_type):
        color, label = color_label(score, metric_type)
        axis_range = [0, 100] if metric_type in ["CSAT", "Fees"] else [-100, 100]
        steps = (
            [{'range': [0, 70], 'color': '#FF6B6B'},
             {'range': [70, 80], 'color': '#FFD93D'},
             {'range': [80, 90], 'color': '#6BCB77'},
             {'range': [90, 100], 'color': '#4D96FF'}]
            if metric_type in ["CSAT", "Fees"]
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
            gauge={'axis': {'range': axis_range}, 'bar': {'color': color}, 'steps': steps}
        ))
        fig.update_layout(height=300, margin=dict(l=30, r=30, t=60, b=30))
        return fig, label

    c1, c2, c3 = st.columns(3)
    fig1, lab1 = gauge(csat, "السعادة العامة", "CSAT")
    fig2, lab2 = gauge(Fees,  "الرضا عن الرسوم", "Fees")
    fig3, lab3 = gauge(nps,  "صافي نقاط الترويج", "NPS")
    c1.plotly_chart(fig1, use_container_width=True)
    c1.markdown(f"**التفسير:** {lab1}")
    if csat_col: c1.caption(f"المصدر: {csat_col}")
    c2.plotly_chart(fig2, use_container_width=True)
    c2.markdown(f"**التفسير:** {lab2}")
    if Fees_col: c2.caption(f"المصدر: {Fees_col}")
    c3.plotly_chart(fig3, use_container_width=True)
    c3.markdown(f"**التفسير:** {lab3}")
    if nps_col: c3.caption(f"المصدر: {nps_col}")
    c3.markdown(f"المروجون: {p_pct:.1f}% | المحايدون: {s_pct:.1f}% | المعارضون: {d_pct:.1f}%", unsafe_allow_html=True)

    legend_html = """
    <div style='background-color:#f9f9f9;border:1px solid #ddd;border-radius:10px;padding:12px;margin-top:15px;'>
        <h4 style='margin-bottom:8px;'>🎨 وسيلة الإيضاح — السعادة العامة / الرضا عن الرسوم</h4>
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
    dim_subcols = [c for c in df_view.columns if re.match(r"Dim\d+\.", str(c).strip())]
    if not dim_subcols:
        st.info("لا توجد أعمدة فرعية للأبعاد (مثل Dim1.1 أو Dim2.3).")
    else:
        main_dim_map = {}
        for i in range(1, 6):
            sub = [c for c in df_view.columns if str(c).startswith(f"Dim{i}.")]
            if sub:
                main_dim_map[f"Dim{i}"] = df_view[sub].apply(pd.to_numeric, errors="coerce").mean(axis=1)

        summary = []
        for dim, series in main_dim_map.items():
            score = series_to_percent(series)
            summary.append({"Dimension": dim, "Score": score})

        dims = pd.DataFrame(summary).dropna()
        if dims.empty:
            st.info("لا توجد نتائج كافية للأبعاد.")
        else:
            dims["Order"] = dims["Dimension"].str.extract(r"(\d+)").astype(float)
            dims = dims.sort_values("Order").reset_index(drop=True)

            for sheet_name in lookup_catalog.keys():
                if "QUESTION" in sheet_name:
                    qtbl = lookup_catalog[sheet_name].copy()
                    qtbl.columns = [str(c).strip().upper() for c in qtbl.columns]

                    code_col = next((c for c in qtbl.columns if any(k in c for k in ["DIM", "CODE", "QUESTION", "ID"])), None)
                    name_col = next((c for c in qtbl.columns if any(k in c for k in ["ARABIC", "NAME", "LABEL", "TEXT"])), None)

                    if code_col and name_col:
                        def _norm(s):
                            return s.astype(str).str.upper().str.replace(r"\\s+", "", regex=True)

                        code_series = _norm(qtbl[code_col])
                        name_series = qtbl[name_col].astype(str)
                        map_dict = dict(zip(code_series, name_series))

                        dims["Dimension"] = (
                            _norm(dims["Dimension"])
                            .map(map_dict)
                            .fillna(dims["Dimension"])
                        )
                    break

            def cat(score):
                if score < 70:  return "🔴 ضعيف"
                elif score < 80: return "🟡 متوسط"
                elif score < 90: return "🟢 جيد"
                else:            return "🔵 ممتاز"
            dims["Category"] = dims["Score"].apply(cat)

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
                title={'text': "<span style='font-size:22px; font-weight:bold;'>تحليل متوسط الأبعاد 📊</span>", 'x': 0.5, 'xanchor': 'center'},
                yaxis=dict(range=[0, 100]),
                xaxis_title="البعد",
                yaxis_title="النسبة المئوية (%)"
            )

            st.plotly_chart(fig, use_container_width=True)
            st.markdown(
                """
                **🗂️ وسيلة الإيضاح:**
                - 🔴 أقل من 70٪ — ضعيف الأداء  
                - 🟡 من 70٪ إلى أقل من 80٪ — متوسط  
                - 🟢 من 80٪ إلى أقل من 90٪ — جيد  
                - 🔵 90٪ فأكثر — ممتاز  
                """,
                unsafe_allow_html=True
            )
            st.dataframe(
                dims[["Dimension", "Score"]]
                .rename(columns={"Dimension": "البعد", "Score": "النسبة (%)"})
                .style.format({"النسبة (%)": "{:.1f}%"}),
                use_container_width=True,
                hide_index=True
            )

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

        exclude_terms = ["", " ", "لا يوجد", "لايوجد", "لاشيء", "لا شيء",
                         "none", "no", "nothing", "nil", "جيد", "ممتاز", "ok", "تمام", "great"]
        data_unsat = data_unsat[~data_unsat["Comment"].str.lower().isin([t.lower() for t in exclude_terms])]
        data_unsat = data_unsat[data_unsat["Comment"].apply(lambda x: len(x.split()) >= 2)]

        if data_unsat.empty:
            st.info("لا توجد ملاحظات نصية كافية بعد التنظيف.")
        else:
            themes = {
                "السرعة / زمن الإنجاز": [
                    "بطء", "البطء", "بطيء", "احيانا البرنامج بطئاً", "Site loading", "loading", "Delay", "Late", "Long delay",
                    "التأخير الكثير", "تاخير المعامله", "تاخير المعاملات", "طول وقت المتابعة", "طول فترة الانجاز", "التاخير", "تاخير", "التأخير"
                ],
                "الإجراءات / الخطوات": [
                    "إجراء", "اجراء", "عملية", "process","خطوات", "مراحل", "نموذج","كثرة الإجراءات", "كثرة التعقيدات","صعوبة الاجراءات",
                    "كثرة تغيير الاجراءات", "كثرة المدخلات المطلوبة", "عدم وجود خطوات واضحة", "كثرة إرسال الرسائل", "الاشعارات المتكررة",
                    "صعوبة التعديل على الإدخال بعد التقديم"
                ],
                "الرسوم / الدفع الرقمي": [
                    "رسوم", "دفع الرسوم", "دفع رسوم بدون نتيجة", "خصم المبلغ", "اخسر فلوس", "المبالغ المالية", "التكاليف", "النسبة العالية",
                    "يرجى تسهيل عمليات الدفع", "عملية الدفع بطيئة جدا", "عدم استرجاع المبلغ", "رسوم الدفع الاضافية للبوابة",
                    "payment issues occurred most of the time"
                ],
                "التواصل / الدعم الفني": [
                    "تواصل", "اتصال", "رد", "response", "support", "customer support", "customer service", "صعوبة التواصل",
                    "صعوبة التواصل في حال وجود مشكلة", "عدم استجابة فريق الدعم الفني", "عدم استجابة الدعم الفني لحل مشاكل النظام",
                    "خدمات الدعم الفني ليست سلسه", "عدم الاستجابة", "عدم الاستجابه السريعه", "لم أتلقى أي رد",
                    "لم احصل على معلومات من الشكوى", "NO PROPER CUSTOMER SUPPORT", "NEED TO EASLY CONTACT TO CUSTOMER SUPPORT ONLINE"
                ],
                "الوضوح / المعلومات": [
                    "There is not proper information in English", "معلومة", "معلومات", "تفاصيل", "بيانات", "غير واضحه المعلومه",
                    "صعوبة الحصول على المعلومات", "قلة وضوح المعلومات", "عدم وضوح المتطلبات", "عدم وضوح الملاحظات عند الرفض",
                    "Properly information not giving", "court communication in Arabic only"
                ],
                "الأمان / الدخول": [
                    "دخول", "login", "تحقق", "كلمة مرور", "أمان", "عدم القدرة على الدخول عبر الهاتف",
                    "Some issues when accessing with UAE pass", "Some bug with app access", "عدم القدرة على الطلب"
                ],
                "الأعطال التقنية عامة": [
                    "مشاكل النظام", "مشكلة تقنية", "عدم فتح الرابط للموقع", "المتصفح بطئ جدا", "الموقع يحتاج إلى تحديث",
                    "توقف الموقع الالكتروني عن العمل", "Errors for the service", "Bug", "Some bug with app access","التطبيق يحتاج إلى تعديلات"
                ],
                "رفع وتحميل المستندات": [
                    "طريقة تحميل المستندات", "صعوبة رفع المستندات", "المتصفح لا يحفظ مستندات", "the repeat upload of papers",
                    "No option for attaching the photo", "تم ارفاق الاوراق ولم تظهر", "عدم القدرة على تخليص المعاملة بسبب المستندات",
                    "صعوبة تقديم الخدمات عبر الموقع/التطبيق"
                ],
            }

            def classify_text(txt):
                t = txt.lower()
                for theme, keywords in themes.items():
                    if any(k.lower() in t for k in keywords):
                        return theme
                return "غير مصنّف"

            data_unsat["المحور"] = data_unsat["Comment"].apply(classify_text)
            data_unsat = data_unsat[data_unsat["المحور"] != "غير مصنّف"]

            summary = data_unsat.groupby("المحور").agg({"Comment": lambda x: " / ".join(x.tolist())}).reset_index()

            summary["عدد الملاحظات"] = summary["Comment"].apply(lambda x: len(x.split("/")))
            summary = summary.sort_values("عدد الملاحظات", ascending=False).reset_index(drop=True)
            summary["النسبة (%)"] = summary["عدد الملاحظات"] / summary["عدد الملاحظات"].sum() * 100
            summary["النسبة التراكمية (%)"] = summary["النسبة (%)"].cumsum()
            summary["اللون"] = np.where(summary["النسبة التراكمية (%)"] <= 80, "#E74C3C", "#BDC3C7")

            if not summary[summary["النسبة التراكمية (%)"] > 80].empty:
                first_above = summary[summary["النسبة التراكمية (%)"] > 80].index[0]
                summary.loc[first_above, "اللون"] = "#E74C3C"

            st.dataframe(
                summary[["المحور", "عدد الملاحظات", "النسبة (%)", "النسبة التراكمية (%)", "Comment"]]
                .rename(columns={"Comment": "التعليقات (مجمعة)"}).style.format({
                    "النسبة (%)": "{:.1f}%",
                    "النسبة التراكمية (%)": "{:.1f}%"
                }),
                use_container_width=True,
                hide_index=True
            )

            fig = go.Figure()
            fig.add_bar(x=summary["المحور"], y=summary["عدد الملاحظات"], marker_color=summary["اللون"], name="عدد الملاحظات")
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
                title={"text": "📊 تحليل باريتو - المحاور الرئيسية", "x": 0.5, "y": 0.95, "xanchor": "center", "yanchor": "top"},
                title_font_size=20,
                xaxis=dict(title="المحور", tickangle=-15),
                yaxis=dict(title="عدد الملاحظات"),
                yaxis2=dict(title="النسبة التراكمية (%)", overlaying="y", side="right", range=[0, 110]),
                height=600,
                bargap=0.3,
                legend=dict(orientation="h", y=-0.2)
            )
            st.plotly_chart(fig, use_container_width=True)

            pareto_buffer = io.BytesIO()
            with pd.ExcelWriter(pareto_buffer, engine="openpyxl") as writer:
                summary.to_excel(writer, index=False, sheet_name="Pareto_Results")

            pareto_buffer.seek(0)

            st.download_button(
                label="📥 تنزيل جدول Pareto (Excel)",
                data=pareto_buffer.getvalue(),
                file_name=f"Pareto_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

# =========================================================
# تبويب خاص لحكومة رأس الخيمة: مقارنة الجهات في مؤشرات الأداء والأبعاد
# =========================================================
if is_admin:
    df_all, _ = load_all_entities()

    with tab_admin:

        if "ENTITY_NAME" not in df_all.columns:
            st.warning("⚠️ لا يوجد عمود ENTITY_NAME في البيانات المجمّعة.")
        else:
            # =========================================================
            # 1) جدول مقارنة مؤشرات الأداء الرئيسية حسب الجهة + تنزيل
            # =========================================================
            csat_col, Fees_col, nps_col = autodetect_metric_cols(df_all)

            rows = []
            for ent, g in df_all.groupby("ENTITY_NAME"):
                row = {"الجهة": ent, "عدد الردود": len(g)}

                if csat_col:
                    row["السعادة العامة (%)"] = series_to_percent(g[csat_col])
                if Fees_col:
                    row["الرضا عن الرسوم (%)"] = series_to_percent(g[Fees_col])

                nps_val, _, _, _, _ = detect_nps(g)
                row["NPS (%)"] = nps_val

                rows.append(row)

            kpi_df = pd.DataFrame(rows)

            if kpi_df.empty:
                st.info("لا توجد بيانات كافية لحساب مؤشرات الأداء الرئيسية.")
            else:
                st.markdown(
                    """
                    <h3 style='text-align:center; font-size:22px; font-weight:bold;'>
                    🔍 مقارنة مؤشرات الأداء الرئيسية حسب الجهة
                    </h3>
                    """,
                    unsafe_allow_html=True
                )

                kpi_display = kpi_df.copy()
                for c in ["السعادة العامة (%)", "الرضا عن الرسوم (%)", "NPS (%)"]:
                    if c in kpi_display.columns:
                        kpi_display[c] = pd.to_numeric(kpi_display[c], errors="coerce").round(1)

                st.dataframe(
                    kpi_display.style.format({
                        "السعادة العامة (%)": "{:.1f}%",
                        "الرضا عن الرسوم (%)": "{:.1f}%",
                        "NPS (%)": "{:.1f}%",
                        "عدد الردود": "{:,.0f}"
                    }),
                    use_container_width=True,
                    hide_index=True
                )

                # ✅ تنزيل جدول المؤشرات
                kpi_buf = io.BytesIO()
                with pd.ExcelWriter(kpi_buf, engine="openpyxl") as writer:
                    kpi_display.to_excel(writer, index=False, sheet_name="KPI_Comparison")
                st.download_button(
                    "📥 تنزيل جدول المؤشرات (Excel)",
                    data=kpi_buf.getvalue(),
                    file_name=f"KPI_Comparison_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            st.markdown("---")

            # =========================================================
            # 2) جدول مقارنة الأبعاد (مرتب): الجهات صفوف × الأبعاد أعمدة + تنزيل
            # =========================================================
            dim_subcols = [c for c in df_all.columns if re.match(r"Dim\d+\.", str(c).strip())]

            if not dim_subcols:
                st.info("لا توجد أعمدة فرعية للأبعاد (مثل Dim1.1 أو Dim2.3) في البيانات.")
            else:
                main_ids = sorted({
                    int(re.match(r"Dim(\d+)\.", str(c).strip()).group(1))
                    for c in dim_subcols
                    if re.match(r"Dim(\d+)\.", str(c).strip())
                })

                # حساب نتيجة كل بُعد رئيسي لكل جهة
                rows = []
                for ent, g in df_all.groupby("ENTITY_NAME"):
                    for i in main_ids:
                        sub = [c for c in g.columns if str(c).startswith(f"Dim{i}.")]
                        if not sub:
                            continue

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
                    # تسمية افتراضية للأبعاد (سنحاول استبدالها من Questions إن توفرت)
                    dim_comp_df["Dimension_label"] = dim_comp_df["Dimension"]

                    # محاولة استبدال أسماء الأبعاد من ورقة Questions إن وُجدت
                    for sheet_name in lookup_catalog.keys():
                        if "QUESTION" in sheet_name.upper():
                            qtbl = lookup_catalog[sheet_name].copy()
                            qtbl.columns = [str(c).strip().upper() for c in qtbl.columns]

                            code_col = next((c for c in qtbl.columns if any(k in c for k in ["DIM", "CODE", "QUESTION", "ID"])), None)
                            name_col = next((c for c in qtbl.columns if any(k in c for k in ["ARABIC", "NAME", "LABEL", "TEXT"])), None)

                            if code_col and name_col:
                                def _norm(s):
                                    return s.astype(str).str.upper().str.replace(r"\s+", "", regex=True)

                                map_dict = dict(zip(_norm(qtbl[code_col]), qtbl[name_col].astype(str)))
                                dim_comp_df["Dimension_label"] = (
                                    _norm(dim_comp_df["Dimension"]).map(map_dict).fillna(dim_comp_df["Dimension"])
                                )
                            break

                    dim_comp_df["Score"] = pd.to_numeric(dim_comp_df["Score"], errors="coerce").round(1)

                    st.markdown(
                        """
                        <h3 style='text-align:center; font-size:22px; font-weight:bold;'>
                        📋 مقارنة الأبعاد الرئيسية بين الجهات
                        </h3>
                        """,
                        unsafe_allow_html=True
                    )

                    # ✅ Pivot: الجهات صفوف × الأبعاد أعمدة
                    dim_pivot = (
                        dim_comp_df
                        .pivot_table(
                            index="الجهة",
                            columns="Dimension_label",
                            values="Score",
                            aggfunc="mean"
                        )
                        .reset_index()
                    )

                    # ترتيب الأعمدة: (الجهة أولاً) ثم الأبعاد حسب ترتيب Dim1, Dim2, ...
                    label_order = (
                        dim_comp_df[["Dimension", "Dimension_label"]]
                        .drop_duplicates()
                        .assign(Order=lambda d: d["Dimension"].str.extract(r"(\d+)").astype(float))
                        .sort_values("Order")["Dimension_label"]
                        .tolist()
                    )
                    ordered_cols = ["الجهة"] + [c for c in label_order if c in dim_pivot.columns]
                    dim_pivot = dim_pivot[ordered_cols]

                    st.dataframe(
                        dim_pivot.style.format({c: "{:.1f}%" for c in dim_pivot.columns if c != "الجهة"}),
                        use_container_width=True,
                        hide_index=True
                    )

                    # ✅ تنزيل جدول الأبعاد (Pivot)
                    dim_buf = io.BytesIO()
                    with pd.ExcelWriter(dim_buf, engine="openpyxl") as writer:
                        dim_pivot.to_excel(writer, index=False, sheet_name="Dimensions_Comparison")
                    st.download_button(
                        "📥 تنزيل جدول الأبعاد (Excel)",
                        data=dim_buf.getvalue(),
                        file_name=f"Dimensions_Comparison_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
             
# =========================================================
# تحسينات شكلية
# =========================================================
st.markdown("""
    <style>
    #MainMenu {visibility: hidden;}
    footer, [data-testid="stFooter"] {opacity: 0.03 !important; height: 1px !important; overflow: hidden !important;}
    </style>
""", unsafe_allow_html=True)
#نضيف العام المقبل نقطتين من شات جي بي تي، نقطتي التوصيات وإعداد تقرير كامل. ممكن أن نعطي نموذج تقرير ونطلب منه أن يقوم بإعداد تقرير نفسه. 



