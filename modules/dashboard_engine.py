import pandas as pd
import streamlit as st


def _to_float(value, default=0):
    try:
        return float(value)
    except Exception:
        return default


def _money(value):
    try:
        return f"{float(value):,.0f} TL".replace(",", ".")
    except Exception:
        return f"{value} TL"


def _pct(value):
    try:
        return f"%{float(value):.1f}"
    except Exception:
        return f"%{value}"


def _status_color(status):
    if status in ["Kritik", "Yüksek Risk", "Zayıf"]:
        return "#dc2626"
    if status in ["Dikkat", "Orta Risk"]:
        return "#d97706"
    return "#059669"


def kpi_status(label, value):
    value = _to_float(value)

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


def inject_enterprise_css():
    st.markdown(
        """
<style>
.ceo-grid {
    display: grid;
    grid-template-columns: repeat(5, 1fr);
    gap: 16px;
    margin-bottom: 22px;
}

.ceo-card {
    background: white;
    border: 1px solid #e2e8f0;
    border-radius: 20px;
    padding: 20px;
    box-shadow: 0 10px 28px rgba(15, 23, 42, 0.07);
    min-height: 118px;
}

.ceo-title {
    font-size: 12px;
    font-weight: 800;
    color: #64748b;
    letter-spacing: .02em;
    text-transform: uppercase;
    margin-bottom: 8px;
}

.ceo-value {
    font-size: 26px;
    font-weight: 900;
    color: #0f172a;
    line-height: 1.1;
}

.ceo-sub {
    font-size: 12px;
    color: #64748b;
    margin-top: 8px;
}

.gauge-wrap {
    background: white;
    border: 1px solid #e2e8f0;
    border-radius: 22px;
    padding: 22px;
    box-shadow: 0 10px 28px rgba(15, 23, 42, 0.07);
    margin-bottom: 20px;
}

.gauge-title {
    font-size: 14px;
    font-weight: 900;
    color: #0f172a;
    margin-bottom: 12px;
}

.progress-bg {
    height: 16px;
    border-radius: 999px;
    background: #e2e8f0;
    overflow: hidden;
}

.progress-fill {
    height: 16px;
    border-radius: 999px;
}

.insight-box {
    background: linear-gradient(135deg, #eff6ff 0%, #ffffff 100%);
    border: 1px solid #bfdbfe;
    border-left: 6px solid #2563eb;
    border-radius: 18px;
    padding: 20px;
    margin-bottom: 20px;
}

.insight-title {
    font-size: 15px;
    font-weight: 900;
    color: #0f172a;
    margin-bottom: 8px;
}

.matrix-card {
    background: white;
    border: 1px solid #e2e8f0;
    border-radius: 18px;
    padding: 18px;
    box-shadow: 0 8px 22px rgba(15, 23, 42, 0.06);
}

.risk-pill {
    display: inline-block;
    padding: 7px 13px;
    border-radius: 999px;
    color: white;
    font-size: 12px;
    font-weight: 900;
}
</style>
        """,
        unsafe_allow_html=True
    )


def render_ceo_cards(score, risk_level, risk_score, financial_result, maturity):
    annual_saving = financial_result.get("Tahmini Yıllık Tasarruf", 0)
    roi = financial_result.get("ROI (%)", 0)
    payback = financial_result.get("Geri Dönüş Süresi (Ay)", 0)

    st.markdown(
        f"""
<div class="ceo-grid">
    <div class="ceo-card">
        <div class="ceo-title">OptiFlow Score</div>
        <div class="ceo-value">{score}/100</div>
        <div class="ceo-sub">Operational performance score</div>
    </div>
    <div class="ceo-card">
        <div class="ceo-title">Risk Level</div>
        <div class="ceo-value">{risk_level}</div>
        <div class="ceo-sub">Risk score: {risk_score}/100</div>
    </div>
    <div class="ceo-card">
        <div class="ceo-title">Annual Saving</div>
        <div class="ceo-value">{_money(annual_saving)}</div>
        <div class="ceo-sub">Estimated financial impact</div>
    </div>
    <div class="ceo-card">
        <div class="ceo-title">ROI</div>
        <div class="ceo-value">{_pct(roi)}</div>
        <div class="ceo-sub">Return potential</div>
    </div>
    <div class="ceo-card">
        <div class="ceo-title">Maturity</div>
        <div class="ceo-value">{str(maturity.get("Seviye", "-")).split(" - ")[0]}</div>
        <div class="ceo-sub">Operational maturity level</div>
    </div>
</div>
        """,
        unsafe_allow_html=True
    )


def render_score_gauge(score, risk_score):
    score = max(0, min(100, _to_float(score)))
    risk_score = max(0, min(100, _to_float(risk_score)))

    score_color = "#059669" if score >= 75 else "#d97706" if score >= 55 else "#dc2626"
    risk_color = "#dc2626" if risk_score >= 25 else "#d97706" if risk_score >= 15 else "#059669"

    col1, col2 = st.columns(2)

    with col1:
        st.markdown(
            f"""
<div class="gauge-wrap">
    <div class="gauge-title">OptiFlow Performance Gauge</div>
    <div style="display:flex; justify-content:space-between; margin-bottom:8px;">
        <b style="font-size:28px; color:#0f172a;">{score:.0f}/100</b>
        <span style="color:#64748b; font-size:13px;">Target: 80+</span>
    </div>
    <div class="progress-bg">
        <div class="progress-fill" style="width:{score}%; background:{score_color};"></div>
    </div>
</div>
            """,
            unsafe_allow_html=True
        )

    with col2:
        st.markdown(
            f"""
<div class="gauge-wrap">
    <div class="gauge-title">Enterprise Risk Gauge</div>
    <div style="display:flex; justify-content:space-between; margin-bottom:8px;">
        <b style="font-size:28px; color:#0f172a;">{risk_score:.1f}/100</b>
        <span style="color:#64748b; font-size:13px;">Lower is better</span>
    </div>
    <div class="progress-bg">
        <div class="progress-fill" style="width:{risk_score}%; background:{risk_color};"></div>
    </div>
</div>
            """,
            unsafe_allow_html=True
        )


def render_kpi_health(company_metrics, financial_result):
    wait_rate = _to_float(company_metrics.get("wait_rate", 0))
    oee = _to_float(company_metrics.get("oee", 0))
    defect_rate = _to_float(company_metrics.get("defect_rate", 0))
    line_balance_loss = _to_float(company_metrics.get("line_balance_loss", 0))
    roi = _to_float(financial_result.get("ROI (%)", 0))

    rows = [
        ["Bekleme Oranı", _pct(wait_rate), kpi_status("wait", wait_rate)],
        ["OEE", _pct(oee), kpi_status("oee", oee)],
        ["Hata / Fire Oranı", _pct(defect_rate), kpi_status("defect", defect_rate)],
        ["Hat Denge Kaybı", _pct(line_balance_loss), kpi_status("balance", line_balance_loss)],
        ["ROI", _pct(roi), kpi_status("roi", roi)]
    ]

    df = pd.DataFrame(rows, columns=["KPI", "Değer", "Durum"])

    st.markdown("### KPI Health Matrix")
    st.dataframe(df, use_container_width=True, hide_index=True)


def render_benchmark_chart(company_metrics, benchmark_result):
    st.markdown("### Benchmark Intelligence")

    categories = []
    firm_values = []
    sector_values = []

    pairs = [
        ("Bekleme", "Firma Bekleme Oranı", "Sektör Bekleme Oranı"),
        ("OEE", "Firma OEE", "Sektör OEE"),
        ("Hata", "Firma Hata Oranı", "Sektör Hata Oranı"),
        ("Denge Kaybı", "Firma Hat Denge Kaybı", "Sektör Hat Denge Kaybı")
    ]

    for label, firm_key, sector_key in pairs:
        categories.append(label)
        firm_values.append(_to_float(benchmark_result.get(firm_key, 0)))
        sector_values.append(_to_float(benchmark_result.get(sector_key, 0)))

    chart_data = pd.DataFrame({
        "KPI": categories,
        "Firma": firm_values,
        "Sektör": sector_values
    })

    st.bar_chart(chart_data.set_index("KPI"), use_container_width=True)


def render_financial_chart(financial_result):
    st.markdown("### Financial Impact View")

    chart_data = pd.DataFrame({
        "Gösterge": ["Operasyonel Kayıp", "İyileştirme Potansiyeli", "Yıllık Tasarruf"],
        "Tutar": [
            _to_float(financial_result.get("Toplam Operasyonel Kayıp", 0)),
            _to_float(financial_result.get("İyileştirme Potansiyeli", 0)),
            _to_float(financial_result.get("Tahmini Yıllık Tasarruf", 0))
        ]
    })

    st.bar_chart(chart_data.set_index("Gösterge"), use_container_width=True)


def render_management_insight(company_name, score, risk_level, financial_result):
    annual_saving = financial_result.get("Tahmini Yıllık Tasarruf", 0)

    st.markdown(
        f"""
<div class="insight-box">
    <div class="insight-title">CEO-Level Management Insight</div>
    {company_name} için mevcut analiz, operasyonel verimlilikte yönetilebilir fakat önceliklendirilmesi gereken kayıplar olduğunu göstermektedir.
    OptiFlow Score <b>{score}/100</b>, risk seviyesi <b>{risk_level}</b> ve yıllık tasarruf potansiyeli
    <b>{_money(annual_saving)}</b> seviyesindedir. Yönetim için ilk odak alanı bekleme süreleri, hat dengeleme ve KPI ritminin kurulmasıdır.
</div>
        """,
        unsafe_allow_html=True
    )


def render_enterprise_dashboard(
    company_name,
    score,
    maturity,
    company_metrics,
    benchmark_result,
    financial_result,
    risk_score,
    risk_level
):
    inject_enterprise_css()
    render_ceo_cards(score, risk_level, risk_score, financial_result, maturity)
    render_score_gauge(score, risk_score)
    render_management_insight(company_name, score, risk_level, financial_result)

    col1, col2 = st.columns(2)

    with col1:
        render_kpi_health(company_metrics, financial_result)

    with col2:
        render_financial_chart(financial_result)

    render_benchmark_chart(company_metrics, benchmark_result)
