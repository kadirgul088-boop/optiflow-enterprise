import os
import json
from datetime import datetime

import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
import streamlit as st

from modules.scoring import calculate_optiflow_score
from modules.financial import calculate_financial_impact
from modules.maturity import get_maturity_comment
from modules.recommendations import generate_recommendations
from modules.ai_engine import generate_consulting_report
from modules.report_engine import create_enterprise_pdf
from modules.ai_copilot import ask_real_ai_copilot

try:
    from modules.excel_engine import process_uploaded_excel, create_template_excel
except Exception:
    process_uploaded_excel = None
    create_template_excel = None

try:
    from modules.ppt_engine import create_enterprise_pptx
except Exception:
    create_enterprise_pptx = None


st.set_page_config(
    page_title="OptiFlow Enterprise V8",
    page_icon="📊",
    layout="wide"
)


PROJECT_DIR = "client_projects"
EXPORT_DIR = "exports"
DATA_DIR = "data"
BENCHMARK_DB_PATH = os.path.join(DATA_DIR, "benchmark_database.json")

os.makedirs(PROJECT_DIR, exist_ok=True)
os.makedirs(EXPORT_DIR, exist_ok=True)
os.makedirs(DATA_DIR, exist_ok=True)


DEFAULT_BENCHMARK_DB = {
    "Tekstil": {"wait_rate": 18, "oee": 74, "defect_rate": 4, "line_balance_loss": 15},
    "Otomotiv": {"wait_rate": 12, "oee": 82, "defect_rate": 2, "line_balance_loss": 10},
    "Gida": {"wait_rate": 15, "oee": 78, "defect_rate": 3, "line_balance_loss": 12},
    "Lojistik": {"wait_rate": 20, "oee": 70, "defect_rate": 5, "line_balance_loss": 18},
    "Genel Uretim": {"wait_rate": 17, "oee": 75, "defect_rate": 4, "line_balance_loss": 14}
}


def ensure_benchmark_db():
    if not os.path.exists(BENCHMARK_DB_PATH):
        with open(BENCHMARK_DB_PATH, "w", encoding="utf-8") as file:
            json.dump(DEFAULT_BENCHMARK_DB, file, ensure_ascii=False, indent=4)


def load_benchmark_db():
    ensure_benchmark_db()
    with open(BENCHMARK_DB_PATH, "r", encoding="utf-8") as file:
        return json.load(file)


def save_benchmark_db(db):
    with open(BENCHMARK_DB_PATH, "w", encoding="utf-8") as file:
        json.dump(db, file, ensure_ascii=False, indent=4)


def money_fmt(value):
    try:
        return f"{float(value):,.0f} TL".replace(",", ".")
    except Exception:
        return f"{value} TL"


def pct_fmt(value):
    try:
        return f"%{float(value):.1f}"
    except Exception:
        return f"%{value}"


def risk_status(wait_rate, oee, defect_rate, line_balance_loss):
    risk_score = round(
        (
            float(wait_rate)
            + float(line_balance_loss)
            + float(defect_rate)
            + max(0, 100 - float(oee))
        ) / 4,
        1
    )

    if risk_score >= 25:
        return risk_score, "Yüksek Risk", "#dc2626", "#fff1f2"

    if risk_score >= 15:
        return risk_score, "Orta Risk", "#d97706", "#fffbeb"

    return risk_score, "Düşük Risk", "#059669", "#ecfdf5"


def kpi_status(label, value):
    value = float(value)

    if label == "wait":
        if value > 20:
            return "Kritik"
        if value > 12:
            return "Dikkat"
        return "İyi"

    if label == "oee":
        if value >= 80:
            return "Güçlü"
        if value >= 65:
            return "Dikkat"
        return "Kritik"

    if label == "defect":
        if value <= 3:
            return "İyi"
        if value <= 7:
            return "Dikkat"
        return "Kritik"

    if label == "balance":
        if value > 15:
            return "Kritik"
        if value > 8:
            return "Dikkat"
        return "İyi"

    if label == "roi":
        if value >= 150:
            return "Güçlü"
        if value >= 75:
            return "Dikkat"
        return "Zayıf"

    return "Normal"


def benchmark_from_database(company_metrics, sector):
    db = load_benchmark_db()
    sector_ref = db.get(sector, db.get("Genel Uretim", {}))

    return {
        "Firma Bekleme Oranı": company_metrics.get("wait_rate", 0),
        "Sektör Bekleme Oranı": sector_ref.get("wait_rate", 0),
        "Firma OEE": company_metrics.get("oee", 0),
        "Sektör OEE": sector_ref.get("oee", 0),
        "Firma Hata Oranı": company_metrics.get("defect_rate", 0),
        "Sektör Hata Oranı": sector_ref.get("defect_rate", 0),
        "Firma Hat Denge Kaybı": company_metrics.get("line_balance_loss", 0),
        "Sektör Hat Denge Kaybı": sector_ref.get("line_balance_loss", 0)
    }


def save_project_record(company_name, sector, score, maturity, company_metrics, financial_result, recommendations):
    safe_company = company_name.replace(" ", "_").replace("/", "_").replace("\\", "_")
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"{safe_company}_{timestamp}.json"
    path = os.path.join(PROJECT_DIR, filename)

    data = {
        "company_name": company_name,
        "sector": sector,
        "created_at": datetime.now().isoformat(),
        "score": score,
        "maturity": maturity,
        "company_metrics": company_metrics,
        "financial_result": financial_result,
        "recommendations": recommendations
    }

    with open(path, "w", encoding="utf-8") as file:
        json.dump(data, file, ensure_ascii=False, indent=4)

    return path


def load_client_records():
    files = sorted(
        [file for file in os.listdir(PROJECT_DIR) if file.endswith(".json")],
        reverse=True
    )

    records = []

    for file_name in files:
        try:
            with open(os.path.join(PROJECT_DIR, file_name), "r", encoding="utf-8") as file:
                record = json.load(file)

            record["_file"] = file_name
            records.append(record)
        except Exception:
            pass

    return records


def plot_gauge(title, value, color_mode="score"):
    value = float(value)

    if color_mode == "risk":
        steps = [
            {"range": [0, 15], "color": "#dcfce7"},
            {"range": [15, 25], "color": "#fef3c7"},
            {"range": [25, 100], "color": "#fee2e2"}
        ]
        bar_color = "#dc2626" if value >= 25 else "#d97706" if value >= 15 else "#059669"
    else:
        steps = [
            {"range": [0, 50], "color": "#fee2e2"},
            {"range": [50, 75], "color": "#fef3c7"},
            {"range": [75, 100], "color": "#dcfce7"}
        ]
        bar_color = "#059669" if value >= 75 else "#d97706" if value >= 50 else "#dc2626"

    fig = go.Figure(
        go.Indicator(
            mode="gauge+number",
            value=value,
            title={"text": title, "font": {"size": 20}},
            number={"font": {"size": 34}},
            gauge={
                "axis": {"range": [0, 100]},
                "bar": {"color": bar_color, "thickness": 0.35},
                "steps": steps,
                "borderwidth": 1,
                "bordercolor": "#cbd5e1"
            }
        )
    )

    fig.update_layout(
        height=300,
        margin=dict(l=20, r=20, t=55, b=15),
        paper_bgcolor="white",
        font=dict(color="#0f172a")
    )

    return fig


def plot_benchmark(benchmark_result):
    data = pd.DataFrame({
        "KPI": ["Bekleme", "OEE", "Hata", "Denge Kaybı"],
        "Firma": [
            float(benchmark_result.get("Firma Bekleme Oranı", 0)),
            float(benchmark_result.get("Firma OEE", 0)),
            float(benchmark_result.get("Firma Hata Oranı", 0)),
            float(benchmark_result.get("Firma Hat Denge Kaybı", 0))
        ],
        "Sektör": [
            float(benchmark_result.get("Sektör Bekleme Oranı", 0)),
            float(benchmark_result.get("Sektör OEE", 0)),
            float(benchmark_result.get("Sektör Hata Oranı", 0)),
            float(benchmark_result.get("Sektör Hat Denge Kaybı", 0))
        ]
    })

    long_df = data.melt(id_vars="KPI", var_name="Kategori", value_name="Değer")

    fig = px.bar(
        long_df,
        x="KPI",
        y="Değer",
        color="Kategori",
        barmode="group",
        text="Değer",
        color_discrete_map={
            "Firma": "#2563eb",
            "Sektör": "#94a3b8"
        }
    )

    fig.update_layout(
        title="Benchmark Comparison",
        height=360,
        plot_bgcolor="white",
        paper_bgcolor="white",
        margin=dict(l=20, r=20, t=55, b=30),
        legend_title_text=""
    )

    fig.update_traces(texttemplate="%{text:.1f}", textposition="outside")

    return fig


def plot_financial_waterfall(financial_result):
    total_loss = float(financial_result.get("Toplam Operasyonel Kayıp", 0))
    improvement = float(financial_result.get("İyileştirme Potansiyeli", financial_result.get("Tahmini Yıllık Tasarruf", 0)))
    saving = float(financial_result.get("Tahmini Yıllık Tasarruf", 0))

    fig = go.Figure(
        go.Waterfall(
            name="Financial Impact",
            orientation="v",
            measure=["absolute", "relative", "total"],
            x=["Operasyonel Kayıp", "İyileştirme Etkisi", "Net Tasarruf"],
            y=[total_loss, -improvement, saving],
            connector={"line": {"color": "#64748b"}},
            increasing={"marker": {"color": "#dc2626"}},
            decreasing={"marker": {"color": "#059669"}},
            totals={"marker": {"color": "#2563eb"}}
        )
    )

    fig.update_layout(
        title="Financial Impact Waterfall",
        height=360,
        plot_bgcolor="white",
        paper_bgcolor="white",
        margin=dict(l=20, r=20, t=55, b=30)
    )

    return fig


def plot_risk_heatmap(wait_rate, oee, defect_rate, line_balance_loss):
    risk_matrix = pd.DataFrame(
        [
            [float(wait_rate), float(defect_rate)],
            [float(line_balance_loss), max(0, 100 - float(oee))]
        ],
        index=["Akış / Kapasite", "Denge / OEE"],
        columns=["Bekleme Riski", "Kalite / Etkinlik Riski"]
    )

    fig = px.imshow(
        risk_matrix,
        text_auto=True,
        color_continuous_scale=["#dcfce7", "#fef3c7", "#fee2e2"],
        aspect="auto"
    )

    fig.update_layout(
        title="Enterprise Risk Heatmap",
        height=350,
        margin=dict(l=20, r=20, t=55, b=30),
        paper_bgcolor="white"
    )

    return fig


def render_plotly_dashboard(
    company_name,
    score,
    maturity,
    company_metrics,
    benchmark_result,
    financial_result,
    risk_score,
    risk_level
):
    wait_rate = company_metrics.get("wait_rate", 0)
    oee = company_metrics.get("oee", 0)
    defect_rate = company_metrics.get("defect_rate", 0)
    line_balance_loss = company_metrics.get("line_balance_loss", 0)

    st.markdown("## Executive Enterprise Dashboard")

    col1, col2, col3, col4, col5 = st.columns(5)

    with col1:
        st.metric("OptiFlow Score", f"{score}/100")
    with col2:
        st.metric("Risk Level", risk_level)
    with col3:
        st.metric("Annual Saving", money_fmt(financial_result.get("Tahmini Yıllık Tasarruf", 0)))
    with col4:
        st.metric("ROI", pct_fmt(financial_result.get("ROI (%)", 0)))
    with col5:
        st.metric("Payback", f"{financial_result.get('Geri Dönüş Süresi (Ay)', 0)} Ay")

    gauge_col1, gauge_col2 = st.columns(2)

    with gauge_col1:
        st.plotly_chart(plot_gauge("OptiFlow Score", score, "score"), use_container_width=True)

    with gauge_col2:
        st.plotly_chart(plot_gauge("Enterprise Risk", risk_score, "risk"), use_container_width=True)

    st.markdown(
        f"""
        <div style="padding:20px;border-radius:18px;background:#eff6ff;border-left:6px solid #2563eb;margin-bottom:20px;">
        <b>CEO Insight:</b><br>
        {company_name} için analiz sonucu, operasyonel iyileştirme potansiyelinin yönetim seviyesinde takip edilmesi gerektiğini göstermektedir.
        OptiFlow Score <b>{score}/100</b>, risk seviyesi <b>{risk_level}</b> ve yıllık tasarruf potansiyeli
        <b>{money_fmt(financial_result.get("Tahmini Yıllık Tasarruf", 0))}</b> seviyesindedir.
        </div>
        """,
        unsafe_allow_html=True
    )

    chart_col1, chart_col2 = st.columns(2)

    with chart_col1:
        st.plotly_chart(plot_benchmark(benchmark_result), use_container_width=True)

    with chart_col2:
        st.plotly_chart(plot_financial_waterfall(financial_result), use_container_width=True)

    st.plotly_chart(
        plot_risk_heatmap(wait_rate, oee, defect_rate, line_balance_loss),
        use_container_width=True
    )

    st.markdown("### KPI Health Matrix")

    kpi_rows = [
        ["Bekleme Oranı", pct_fmt(wait_rate), kpi_status("wait", wait_rate)],
        ["OEE", pct_fmt(oee), kpi_status("oee", oee)],
        ["Hata / Fire Oranı", pct_fmt(defect_rate), kpi_status("defect", defect_rate)],
        ["Hat Denge Kaybı", pct_fmt(line_balance_loss), kpi_status("balance", line_balance_loss)],
        ["ROI", pct_fmt(financial_result.get("ROI (%)", 0)), kpi_status("roi", financial_result.get("ROI (%)", 0))]
    ]

    st.dataframe(
        pd.DataFrame(kpi_rows, columns=["KPI", "Değer", "Durum"]),
        use_container_width=True,
        hide_index=True
    )

    st.markdown("### Operasyonel Olgunluk")
    st.info(f"{maturity.get('Seviye', '-')} — {maturity.get('Yorum', '-')}")




def generate_copilot_answer(
    question,
    company_name,
    sector,
    score,
    maturity,
    company_metrics,
    financial_result,
    recommendations,
    risk_score,
    risk_level
):
    return ask_real_ai_copilot(
        question=question,
        company_name=company_name,
        sector=sector,
        score=score,
        maturity=maturity,
        company_metrics=company_metrics,
        financial_result=financial_result,
        recommendations=recommendations,
        risk_score=risk_score,
        risk_level=risk_level
    )

def render_ai_copilot(
    company_name,
    sector,
    score,
    maturity,
    company_metrics,
    financial_result,
    recommendations,
    risk_score,
    risk_level
):
    st.markdown("## Executive AI Copilot")
    st.write("OpenAI destekli gerçek yönetim danışmanlığı copilot'u. Mevcut analiz sonuçlarına göre karar desteği üretir.")

    quick_questions = [
        "En büyük operasyonel risk nedir?",
        "ROI ve tasarruf açısından en güçlü fırsat nedir?",
        "İlk 90 gün yol haritası nasıl olmalı?",
        "Öncelikli ilk 3 aksiyon nedir?",
        "Operasyonel olgunluk seviyesi ne ifade ediyor?"
    ]

    selected_question = st.selectbox("Hazır yönetici sorusu seç", quick_questions)

    custom_question = st.text_input(
        "Veya kendi sorunuzu yazın",
        placeholder="Örn: Bu müşteride ilk hangi iyileştirme projesi başlatılmalı?"
    )

    question = custom_question.strip() if custom_question.strip() else selected_question

    if st.button("Copilot Yanıtı Üret", type="primary"):
        answer = generate_copilot_answer(
            question=question,
            company_name=company_name,
            sector=sector,
            score=score,
            maturity=maturity,
            company_metrics=company_metrics,
            financial_result=financial_result,
            recommendations=recommendations,
            risk_score=risk_score,
            risk_level=risk_level
        )

        st.markdown("### Copilot Yanıtı")
        st.markdown(
            f"""
            <div style="padding:22px;border-radius:18px;background:#f8fafc;border-left:6px solid #2563eb;">
                <b>Soru:</b> {question}<br><br>
                <b>Yanıt:</b><br>
                {answer}
            </div>
            """,
            unsafe_allow_html=True
        )

    st.markdown("### Copilot Context")
    st.dataframe(
        pd.DataFrame(
            [
                ["Firma", company_name],
                ["Sektör", sector],
                ["OptiFlow Score", f"{score}/100"],
                ["Risk", f"{risk_level} - {risk_score}/100"],
                ["Yıllık Tasarruf", money_fmt(financial_result.get("Tahmini Yıllık Tasarruf", 0))],
                ["ROI", pct_fmt(financial_result.get("ROI (%)", 0))],
                ["Geri Dönüş", f"{financial_result.get('Geri Dönüş Süresi (Ay)', 0)} Ay"]
            ],
            columns=["Alan", "Değer"]
        ),
        use_container_width=True,
        hide_index=True
    )


# ---------- UI STYLE ----------

st.markdown(
    """
<style>
.hero {
    padding: 34px 34px;
    border-radius: 26px;
    background: linear-gradient(135deg, #0f172a 0%, #1e3a8a 58%, #2563eb 100%);
    color: white;
    margin-bottom: 28px;
}

.hero-title {
    font-size: 44px;
    font-weight: 900;
    margin-bottom: 8px;
}

.hero-subtitle {
    font-size: 17px;
    opacity: 0.92;
    margin-bottom: 18px;
}

.hero-badge {
    display: inline-block;
    padding: 8px 13px;
    border-radius: 999px;
    background-color: rgba(255,255,255,0.14);
    margin-right: 8px;
    font-size: 13px;
    font-weight: 700;
}
</style>
    """,
    unsafe_allow_html=True
)


with st.sidebar:
    if os.path.exists("assets/logo.png"):
        st.image("assets/logo.png", width=220)

    st.title("OptiFlow Menu")

    page = st.radio(
        "Navigation",
        [
            "Landing Page",
            "Dashboard",
            "AI Copilot",
            "Analysis",
            "Benchmark Center",
            "Clients",
            "Report Center"
        ]
    )


if page == "Landing Page":
    st.markdown(
        """
<div class="hero">
    <div class="hero-title">OptiFlow Enterprise V8</div>
    <div class="hero-subtitle">
        Operational Excellence Intelligence Platform with Plotly Executive Dashboards, KPI Diagnostics and Enterprise Reporting.
    </div>
    <span class="hero-badge">Plotly Dashboard</span>
    <span class="hero-badge">KPI Diagnostics</span>
    <span class="hero-badge">Financial Impact</span>
    <span class="hero-badge">AI Consulting</span>
    <span class="hero-badge">PDF & PPT Export</span>
</div>
        """,
        unsafe_allow_html=True
    )

    col_a, col_b, col_c = st.columns(3)

    with col_a:
        st.info("1. Upload Excel or enter operational KPIs.")
    with col_b:
        st.info("2. Quantify financial impact and operational risk.")
    with col_c:
        st.info("3. Export executive PDF and PPT reports.")

    st.stop()


st.markdown(
    """
<div class="hero">
    <div class="hero-title">OptiFlow Enterprise V8</div>
    <div class="hero-subtitle">
        Commercial Operations Excellence Platform | Plotly Dashboard | Benchmark Intelligence | PDF & PPT Export
    </div>
</div>
    """,
    unsafe_allow_html=True
)


with st.sidebar:
    st.markdown("---")
    st.subheader("Analiz Ayarları")

    company_name = st.text_input("Müşteri Firma", "Demo Firma A.Ş.")

    sector = st.selectbox(
        "Sektör",
        ["Tekstil", "Otomotiv", "Gida", "Lojistik", "Genel Uretim"]
    )

    input_mode = st.radio(
        "Veri Giriş Tipi",
        ["Manuel Giriş", "Excel Upload"]
    )

    uploaded_excel = None
    excel_result = None

    if input_mode == "Excel Upload":
        uploaded_excel = st.file_uploader("Müşteri Excel dosyası yükle", type=["xlsx", "xls"])

        if create_template_excel:
            template_path = create_template_excel()
            with open(template_path, "rb") as file:
                st.download_button(
                    "Excel Şablonu İndir",
                    data=file,
                    file_name="OptiFlow_Excel_Template.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

        if uploaded_excel and process_uploaded_excel:
            excel_result = process_uploaded_excel(uploaded_excel)
            st.success("Excel başarıyla okundu.")

    st.markdown("---")

    if excel_result:
        metrics = excel_result["metrics"]
        financial_inputs = excel_result["financial_inputs"]

        wait_rate = float(metrics["wait_rate"])
        oee = float(metrics["oee"])
        defect_rate = float(metrics["defect_rate"])
        line_balance_loss = float(metrics["line_balance_loss"])
        capacity_score = float(metrics["capacity_score"])

        total_wait_minutes = float(financial_inputs["total_wait_minutes"])
        hourly_labor_cost = float(financial_inputs["hourly_labor_cost"])
        working_days = int(financial_inputs["working_days"])
        improvement_rate = int(financial_inputs["improvement_rate"])

        st.info("KPI değerleri Excel verisine göre hesaplandı.")
    else:
        st.subheader("Operasyonel KPI")
        wait_rate = st.number_input("Bekleme Oranı (%)", min_value=0.0, max_value=100.0, value=27.0, step=1.0)
        oee = st.number_input("OEE (%)", min_value=0.0, max_value=100.0, value=68.0, step=1.0)
        defect_rate = st.number_input("Hata / Fire Oranı (%)", min_value=0.0, max_value=100.0, value=5.0, step=1.0)
        line_balance_loss = st.number_input("Hat Denge Kaybı (%)", min_value=0.0, max_value=100.0, value=22.0, step=1.0)
        capacity_score = st.number_input("Kapasite Skoru (%)", min_value=0.0, max_value=100.0, value=75.0, step=1.0)

        st.markdown("---")
        st.subheader("Finansal Varsayımlar")
        total_wait_minutes = st.number_input("Toplam Bekleme Süresi (dk/gün)", min_value=0.0, value=120.0, step=10.0)
        hourly_labor_cost = st.number_input("Saatlik İşçilik Maliyeti (TL)", min_value=0.0, value=250.0, step=25.0)
        working_days = st.number_input("Aylık Çalışma Günü", min_value=1, max_value=31, value=22)
        improvement_rate = st.slider("Bekleme İyileştirme Oranı (%)", 5, 60, 20)

    st.markdown("---")
    save_project = st.checkbox("Analizi müşteri arşivine kaydet", value=True)


company_metrics = {
    "wait_rate": wait_rate,
    "oee": oee,
    "defect_rate": defect_rate,
    "line_balance_loss": line_balance_loss
}

benchmark_result = benchmark_from_database(company_metrics, sector)

score = calculate_optiflow_score(
    efficiency_score=100 - wait_rate,
    oee_score=oee,
    capacity_score=capacity_score,
    flow_score=100 - line_balance_loss,
    quality_score=100 - defect_rate
)

maturity = get_maturity_comment(score)

financial_result = calculate_financial_impact(
    total_wait_minutes=total_wait_minutes,
    improvement_rate=improvement_rate,
    hourly_labor_cost=hourly_labor_cost,
    working_days_per_month=working_days
)

recommendations = generate_recommendations(
    wait_rate=wait_rate,
    oee=oee,
    line_balance_loss=line_balance_loss
)

risk_score, risk_level, risk_color, risk_bg = risk_status(
    wait_rate,
    oee,
    defect_rate,
    line_balance_loss
)


if page == "Dashboard":
    render_plotly_dashboard(
        company_name=company_name,
        score=score,
        maturity=maturity,
        company_metrics=company_metrics,
        benchmark_result=benchmark_result,
        financial_result=financial_result,
        risk_score=risk_score,
        risk_level=risk_level
    )


elif page == "AI Copilot":
    render_ai_copilot(
        company_name=company_name,
        sector=sector,
        score=score,
        maturity=maturity,
        company_metrics=company_metrics,
        financial_result=financial_result,
        recommendations=recommendations,
        risk_score=risk_score,
        risk_level=risk_level
    )


elif page == "Analysis":
    left, right = st.columns(2)

    with left:
        st.markdown("### Benchmark Karşılaştırması")
        st.dataframe(benchmark_result, use_container_width=True)

    with right:
        st.markdown("### Finansal Etki Analizi")
        st.dataframe(financial_result, use_container_width=True)

    if excel_result:
        st.markdown("### Yüklenen Excel Verisi")
        st.dataframe(excel_result["dataframe"], use_container_width=True)

        if excel_result.get("bottleneck"):
            st.markdown("### Otomatik Darboğaz Tespiti")
            st.info(
                f"Kritik darboğaz: {excel_result['bottleneck']['process']} - "
                f"{excel_result['bottleneck']['wait_minutes']} dk bekleme"
            )

        if excel_result.get("pareto_data") is not None:
            st.markdown("### Pareto Verisi")
            st.dataframe(excel_result["pareto_data"], use_container_width=True)

    st.markdown("### Yönetim İçin Öncelikli Aksiyonlar")
    for i, rec in enumerate(recommendations, start=1):
        st.markdown(f"**{i}.** {rec}")


elif page == "Benchmark Center":
    st.markdown("### Benchmark Center")

    db = load_benchmark_db()
    st.write("Sektör benchmark değerlerini buradan güncelleyebilirsin.")

    selected_sector = st.selectbox("Benchmark sektörü seç", list(db.keys()))

    c1, c2, c3, c4 = st.columns(4)

    with c1:
        new_wait = st.number_input("Sektör Bekleme Oranı", value=float(db[selected_sector].get("wait_rate", 0)))
    with c2:
        new_oee = st.number_input("Sektör OEE", value=float(db[selected_sector].get("oee", 0)))
    with c3:
        new_defect = st.number_input("Sektör Hata Oranı", value=float(db[selected_sector].get("defect_rate", 0)))
    with c4:
        new_balance = st.number_input("Sektör Hat Denge Kaybı", value=float(db[selected_sector].get("line_balance_loss", 0)))

    if st.button("Benchmark Verisini Güncelle"):
        db[selected_sector] = {
            "wait_rate": new_wait,
            "oee": new_oee,
            "defect_rate": new_defect,
            "line_balance_loss": new_balance
        }
        save_benchmark_db(db)
        st.success("Benchmark verisi güncellendi.")

    st.markdown("### Benchmark Veritabanı")
    st.json(db)


elif page == "Clients":
    st.markdown("### Client Management")

    records = load_client_records()

    if not records:
        st.info("Henüz kayıtlı müşteri analizi bulunmuyor.")
    else:
        client_names = sorted(list(set([record.get("company_name", "-") for record in records])))
        selected_client = st.selectbox("Müşteri seç", client_names)

        filtered = [record for record in records if record.get("company_name") == selected_client]

        st.markdown(f"### {selected_client} Analiz Geçmişi")

        for record in filtered:
            with st.expander(f"{record.get('created_at')} | Score: {record.get('score')}/100"):
                st.write(f"Sektör: {record.get('sector')}")
                st.write(f"Olgunluk: {record.get('maturity', {}).get('Seviye', '-')}")
                st.json(record.get("financial_result", {}))


elif page == "Report Center":
    st.markdown("### Enterprise Report Center")
    st.write("PDF veya PowerPoint raporu oluşturmak için aşağıdaki butonları kullan.")

    col_pdf, col_ppt = st.columns(2)

    with col_pdf:
        if st.button("Enterprise PDF Raporu Oluştur", type="primary"):
            with st.spinner("OptiFlow Enterprise PDF raporu hazırlanıyor..."):
                consulting_report = generate_consulting_report(
                    sector=sector,
                    company_metrics=company_metrics,
                    benchmark_result=benchmark_result,
                    financial_result=financial_result,
                    maturity=maturity,
                    recommendations=recommendations
                )

                pdf_file = create_enterprise_pdf(
                    company_name=company_name,
                    sector=sector,
                    score=score,
                    maturity=maturity,
                    company_metrics=company_metrics,
                    benchmark_result=benchmark_result,
                    financial_result=financial_result,
                    recommendations=recommendations,
                    consulting_report=consulting_report
                )

                project_path = None
                if save_project:
                    project_path = save_project_record(
                        company_name=company_name,
                        sector=sector,
                        score=score,
                        maturity=maturity,
                        company_metrics=company_metrics,
                        financial_result=financial_result,
                        recommendations=recommendations
                    )

            st.success("Enterprise PDF raporu başarıyla oluşturuldu.")

            if project_path:
                st.info(f"Müşteri analizi arşive kaydedildi: {project_path}")

            with open(pdf_file, "rb") as file:
                st.download_button(
                    label="Enterprise PDF İndir",
                    data=file,
                    file_name=f"OptiFlow_{company_name.replace(' ', '_')}_Enterprise_Report.pdf",
                    mime="application/pdf"
                )

    with col_ppt:
        if st.button("Executive PPTX Oluştur"):
            if create_enterprise_pptx is None:
                st.error("PPTX modülü bulunamadı. modules/ppt_engine.py dosyasını ekle.")
            else:
                with st.spinner("PowerPoint raporu hazırlanıyor..."):
                    ppt_file = os.path.join(
                        EXPORT_DIR,
                        f"OptiFlow_{company_name.replace(' ', '_')}_Executive_Deck.pptx"
                    )

                    ppt_file = create_enterprise_pptx(
                        company_name=company_name,
                        sector=sector,
                        score=score,
                        maturity=maturity,
                        company_metrics=company_metrics,
                        financial_result=financial_result,
                        recommendations=recommendations,
                        output_file=ppt_file
                    )

                st.success("PowerPoint raporu oluşturuldu.")

                with open(ppt_file, "rb") as file:
                    st.download_button(
                        label="Executive PPTX İndir",
                        data=file,
                        file_name=os.path.basename(ppt_file),
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                    )
