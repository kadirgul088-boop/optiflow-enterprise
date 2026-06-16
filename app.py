import os
import json
from datetime import datetime

import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
import streamlit as st

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from supabase import create_client

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
    page_title="OptiFlow Enterprise SaaS V11",
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







# ============================================================
# SAAS LAYER: AUTH + PLANS + SUPABASE
# ============================================================

SUPABASE_URL = st.secrets.get("SUPABASE_URL", "")
SUPABASE_ANON_KEY = st.secrets.get("SUPABASE_ANON_KEY", "")
SUPABASE_SERVICE_KEY = st.secrets.get("SUPABASE_SERVICE_KEY", "")


@st.cache_resource
def get_supabase_client():
    key = SUPABASE_SERVICE_KEY or SUPABASE_ANON_KEY
    if not SUPABASE_URL or not key:
        return None
    return create_client(SUPABASE_URL, key)


PLAN_RULES = {
    "demo": {
        "label": "Demo",
        "analysis_limit": 1,
        "pdf": False,
        "ppt": False,
        "excel": False,
        "ai": False,
        "price": "Free Demo"
    },
    "starter": {
        "label": "Starter",
        "analysis_limit": 10,
        "pdf": True,
        "ppt": False,
        "excel": True,
        "ai": False,
        "price": "$49 / month"
    },
    "professional": {
        "label": "Professional",
        "analysis_limit": 100,
        "pdf": True,
        "ppt": True,
        "excel": True,
        "ai": True,
        "price": "$149 / month"
    },
    "enterprise": {
        "label": "Enterprise",
        "analysis_limit": 999999,
        "pdf": True,
        "ppt": True,
        "excel": True,
        "ai": True,
        "price": "Custom"
    }
}


def get_app_url():
    return "https://optiflow-enterprise-ewxicdsl96crsq2ycortqo.streamlit.app"


def is_logged_in():
    try:
        return bool(st.user.is_logged_in)
    except Exception:
        return False


def current_user_email():
    if st.session_state.get("demo_mode", False):
        return "demo@optiflow.ai"

    try:
        return st.user.email
    except Exception:
        return None


def current_user_name():
    if st.session_state.get("demo_mode", False):
        return "Demo User"

    try:
        return st.user.name
    except Exception:
        return current_user_email() or "User"


def get_user_profile(email):
    if st.session_state.get("demo_mode", False):
        return {
            "email": "demo@optiflow.ai",
            "full_name": "Demo User",
            "plan": "demo",
            "plan_status": "active",
            "analysis_limit": 1
        }

    supabase = get_supabase_client()

    if not supabase or not email:
        return {
            "email": email or "demo@optiflow.ai",
            "full_name": "Demo User",
            "plan": "demo",
            "plan_status": "active",
            "analysis_limit": 1
        }

    try:
        result = (
            supabase
            .table("users")
            .select("*")
            .eq("email", email)
            .limit(1)
            .execute()
        )

        if result.data:
            return result.data[0]

        payload = {
            "email": email,
            "full_name": current_user_name(),
            "plan": "demo",
            "plan_status": "active",
            "analysis_limit": 1
        }

        supabase.table("users").insert(payload).execute()
        return payload

    except Exception:
        return {
            "email": email,
            "full_name": current_user_name(),
            "plan": "demo",
            "plan_status": "active",
            "analysis_limit": 1
        }


def get_subscription(email, fallback_plan="demo"):
    supabase = get_supabase_client()

    if st.session_state.get("demo_mode", False):
        return {
            "plan": "demo",
            "status": "active"
        }

    if not supabase or not email:
        return {
            "plan": fallback_plan,
            "status": "inactive"
        }

    try:
        result = supabase.table("subscriptions").select("*").eq("user_email", email).limit(1).execute()

        if result.data:
            return result.data[0]

        # First login: create inactive/demo subscription.
        payload = {
            "user_email": email,
            "plan": fallback_plan,
            "status": "inactive"
        }

        supabase.table("subscriptions").insert(payload).execute()
        return payload

    except Exception as exc:
        st.warning(f"Supabase subscription error: {exc}")
        return {
            "plan": fallback_plan,
            "status": "inactive"
        }


def resolve_active_plan(email):
    profile = get_user_profile(email)
    subscription = get_subscription(email, profile.get("plan", "demo"))

    profile_plan = str(profile.get("plan", "demo")).lower()
    sub_plan = str(subscription.get("plan", profile_plan)).lower()
    sub_status = str(subscription.get("status", "inactive")).lower()

    # Paid plans only unlock features when subscription status is active/trialing.
    if sub_plan in ["starter", "professional", "enterprise"] and sub_status in ["active", "trialing"]:
        plan = sub_plan
    else:
        plan = profile_plan if profile_plan in PLAN_RULES else "demo"

    if st.session_state.get("demo_mode", False):
        plan = "demo"

    return plan, PLAN_RULES.get(plan, PLAN_RULES["demo"]), profile, subscription


def count_user_projects(email):
    supabase = get_supabase_client()

    if not supabase or not email:
        return 0

    try:
        result = supabase.table("projects").select("id", count="exact").eq("user_email", email).execute()
        return result.count or 0
    except Exception:
        return 0


def can_create_analysis(email, plan_rules):
    used = count_user_projects(email)
    limit = int(plan_rules.get("analysis_limit", 1))
    return used < limit, used, limit


def save_project_to_supabase(
    user_email,
    company_name,
    sector,
    score,
    risk_level,
    financial_result
):
    supabase = get_supabase_client()

    if not supabase or not user_email:
        return None

    try:
        payload = {
            "user_email": user_email,
            "company_name": company_name,
            "sector": sector,
            "score": float(score),
            "risk_level": risk_level,
            "annual_saving": float(financial_result.get("Tahmini Yıllık Tasarruf", 0))
        }

        result = supabase.table("projects").insert(payload).execute()

        if result.data:
            return result.data[0].get("id")

    except Exception as exc:
        st.warning(f"Supabase project save error: {exc}")

    return None


def save_report_to_supabase(user_email, project_id, report_type, file_name, file_path):
    supabase = get_supabase_client()

    if not supabase or not user_email or not project_id:
        return None

    try:
        payload = {
            "user_email": user_email,
            "project_id": project_id,
            "report_type": report_type,
            "file_name": file_name,
            "file_path": file_path
        }

        result = supabase.table("reports").insert(payload).execute()

        if result.data:
            return result.data[0].get("id")

    except Exception as exc:
        st.warning(f"Supabase report save error: {exc}")

    return None


def load_user_projects(user_email):
    supabase = get_supabase_client()

    if not supabase or not user_email:
        return []

    try:
        result = (
            supabase.table("projects")
            .select("*")
            .eq("user_email", user_email)
            .order("created_at", desc=True)
            .execute()
        )
        return result.data or []
    except Exception as exc:
        st.warning(f"Supabase project load error: {exc}")
        return []


def render_public_landing():
    st.markdown(
        """
<div class="hero">
    <div class="hero-title">OptiFlow Enterprise SaaS</div>
    <div class="hero-subtitle">
        AI-powered Operational Excellence Platform for KPI diagnostics, financial impact analysis and executive reporting.
    </div>
    <span class="hero-badge">AI Consulting</span>
    <span class="hero-badge">PDF / PPT / Excel Reports</span>
    <span class="hero-badge">Benchmark Intelligence</span>
    <span class="hero-badge">Client Portal</span>
</div>
        """,
        unsafe_allow_html=True
    )

    st.markdown("## Choose your access")

    col1, col2, col3, col4 = st.columns(4)

    plans = [
        ("Demo", "Free Demo", "1 sample analysis", "No exports", "No AI"),
        ("Starter", "$49 / month", "10 analyses", "PDF + Excel", "No PPT / AI"),
        ("Professional", "$149 / month", "100 analyses", "PDF + PPT + Excel", "AI Copilot"),
        ("Enterprise", "Custom", "Unlimited", "Team support", "Custom deployment")
    ]

    for col, plan in zip([col1, col2, col3, col4], plans):
        with col:
            st.markdown(
                f"""
                <div style="padding:22px;border:1px solid #e2e8f0;border-radius:18px;background:white;min-height:235px;">
                    <h3>{plan[0]}</h3>
                    <h2>{plan[1]}</h2>
                    <p>{plan[2]}</p>
                    <p>{plan[3]}</p>
                    <p>{plan[4]}</p>
                </div>
                """,
                unsafe_allow_html=True
            )

    st.markdown("---")

    col_login, col_demo = st.columns(2)

    with col_login:
        st.markdown("### Customer Login")
        st.write("Paid users and trial customers can sign in with Google.")
        if st.button("Sign in with Google", type="primary"):
            st.login()

    with col_demo:
        st.markdown("### Try Demo")
        st.write("Use the limited demo mode without exports and without AI Copilot.")
        if st.button("Continue with Demo"):
            st.session_state.demo_mode = True
            st.rerun()

    st.info("Payments are controlled by subscription plan records. Users without an active plan stay in Demo mode.")


def render_user_bar(user_email, plan, rules, used, limit):
    st.sidebar.markdown("---")
    st.sidebar.markdown("### Account")

    if st.session_state.get("demo_mode", False):
        st.sidebar.info("Demo Mode")
        if st.sidebar.button("Exit Demo"):
            st.session_state.demo_mode = False
            st.rerun()
    else:
        st.sidebar.success(current_user_name())
        st.sidebar.caption(user_email)
        if st.sidebar.button("Log out"):
            st.logout()

    st.sidebar.markdown(f"**Plan:** {rules.get('label', plan)}")
    st.sidebar.markdown(f"**Usage:** {used}/{limit if limit < 999999 else 'Unlimited'} analyses")

    if plan == "demo":
        st.sidebar.warning("Exports and AI Copilot are locked in Demo.")
    elif plan == "starter":
        st.sidebar.info("Starter: PDF + Excel enabled.")
    elif plan == "professional":
        st.sidebar.success("Professional: Full reporting + AI enabled.")
    elif plan == "enterprise":
        st.sidebar.success("Enterprise: Full access enabled.")


def render_locked_feature(feature_name, plan):
    st.warning(
        f"{feature_name} is locked for your current plan ({plan}). "
        "Upgrade to unlock this feature."
    )

    starter_link = st.secrets.get("STRIPE_STARTER_LINK", "")
    professional_link = st.secrets.get("STRIPE_PROFESSIONAL_LINK", "")
    enterprise_link = st.secrets.get("ENTERPRISE_CONTACT_LINK", "")

    col1, col2, col3 = st.columns(3)

    with col1:
        if starter_link:
            st.link_button("Upgrade to Starter", starter_link)

    with col2:
        if professional_link:
            st.link_button("Upgrade to Professional", professional_link)

    with col3:
        if enterprise_link:
            st.link_button("Contact Enterprise Sales", enterprise_link)




def create_excel_report(
    company_name,
    sector,
    score,
    maturity,
    company_metrics,
    benchmark_result,
    financial_result,
    recommendations,
    risk_score,
    risk_level
):
    safe_company = (
        str(company_name)
        .replace(" ", "_")
        .replace("/", "_")
        .replace("\\", "_")
        .replace(".", "")
    )

    file_name = os.path.join(
        EXPORT_DIR,
        f"OptiFlow_{safe_company}_Enterprise_Data.xlsx"
    )

    wb = Workbook()

    header_fill = PatternFill("solid", fgColor="0F172A")
    blue_fill = PatternFill("solid", fgColor="EFF6FF")
    green_fill = PatternFill("solid", fgColor="ECFDF5")
    red_fill = PatternFill("solid", fgColor="FFF1F2")
    white_font = Font(color="FFFFFF", bold=True)
    bold_font = Font(bold=True)
    title_font = Font(bold=True, size=14, color="0F172A")
    thin = Side(border_style="thin", color="CBD5E1")

    def style_sheet(ws):
        for row in ws.iter_rows():
            for cell in row:
                cell.border = Border(top=thin, left=thin, right=thin, bottom=thin)
                cell.alignment = Alignment(vertical="top", wrap_text=True)
        for col in ws.columns:
            max_len = 0
            col_letter = col[0].column_letter
            for cell in col:
                try:
                    max_len = max(max_len, len(str(cell.value)))
                except Exception:
                    pass
            ws.column_dimensions[col_letter].width = min(max(max_len + 2, 14), 42)

    def write_table(ws, start_row, title, headers, rows, fill=None):
        ws.cell(start_row, 1, title)
        ws.cell(start_row, 1).font = title_font

        header_row = start_row + 2
        for c, header in enumerate(headers, start=1):
            cell = ws.cell(header_row, c, header)
            cell.fill = header_fill
            cell.font = white_font

        for r, row in enumerate(rows, start=header_row + 1):
            for c, value in enumerate(row, start=1):
                cell = ws.cell(r, c, value)
                if fill:
                    cell.fill = fill

        return header_row + len(rows) + 3

    # DASHBOARD
    ws = wb.active
    ws.title = "Executive Dashboard"

    dashboard_rows = [
        ["Company", company_name],
        ["Sector", sector],
        ["OptiFlow Score", f"{score}/100"],
        ["Maturity", maturity.get("Seviye", "-")],
        ["Risk Level", risk_level],
        ["Risk Score", f"{risk_score}/100"],
        ["Annual Saving", financial_result.get("Tahmini Yıllık Tasarruf", 0)],
        ["ROI (%)", financial_result.get("ROI (%)", 0)],
        ["Payback Month", financial_result.get("Geri Dönüş Süresi (Ay)", 0)],
    ]

    write_table(
        ws,
        1,
        "OptiFlow Enterprise Dashboard",
        ["Metric", "Value"],
        dashboard_rows,
        blue_fill
    )

    # KPI
    ws2 = wb.create_sheet("KPI Metrics")
    kpi_rows = [
        ["Wait Rate (%)", company_metrics.get("wait_rate", 0)],
        ["OEE (%)", company_metrics.get("oee", 0)],
        ["Defect Rate (%)", company_metrics.get("defect_rate", 0)],
        ["Line Balance Loss (%)", company_metrics.get("line_balance_loss", 0)],
    ]

    write_table(
        ws2,
        1,
        "Operational KPI Metrics",
        ["KPI", "Value"],
        kpi_rows,
        blue_fill
    )

    # BENCHMARK
    ws3 = wb.create_sheet("Benchmark")
    benchmark_rows = [[str(k), str(v)] for k, v in benchmark_result.items()]

    write_table(
        ws3,
        1,
        "Benchmark Comparison",
        ["Benchmark Field", "Value"],
        benchmark_rows,
        blue_fill
    )

    # FINANCIAL
    ws4 = wb.create_sheet("Financial Impact")
    financial_rows = [[str(k), str(v)] for k, v in financial_result.items()]

    write_table(
        ws4,
        1,
        "Financial Impact Analysis",
        ["Financial Field", "Value"],
        financial_rows,
        green_fill
    )

    # RECOMMENDATIONS
    ws5 = wb.create_sheet("Recommendations")
    recommendation_rows = [[i, rec] for i, rec in enumerate(recommendations, start=1)]

    write_table(
        ws5,
        1,
        "Management Recommendations",
        ["#", "Recommendation"],
        recommendation_rows,
        blue_fill
    )

    # ROADMAP
    ws6 = wb.create_sheet("30-60-90 Roadmap")
    roadmap_rows = [
        ["0-30 Days", "Data validation, bottleneck confirmation, KPI baseline", "Quick-win list and measurement setup"],
        ["30-60 Days", "Flow improvement, line balancing, standard work", "Improved flow and measurable capacity gain"],
        ["60-90 Days", "KPI rhythm, responsibility matrix, performance board", "Sustainable management system"],
    ]

    write_table(
        ws6,
        1,
        "30-60-90 Day Transformation Roadmap",
        ["Period", "Focus", "Expected Output"],
        roadmap_rows,
        blue_fill
    )

    for sheet in wb.worksheets:
        style_sheet(sheet)
        sheet.freeze_panes = "A4"

    wb.save(file_name)
    return file_name


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
            "Report Center",
            "Billing",
            "Admin Panel"
        ]
    )






# ============================================================
# ADMIN PANEL LAYER
# ============================================================

def get_admin_emails():
    raw = st.secrets.get("ADMIN_EMAILS", "")
    if not raw:
        return []
    return [email.strip().lower() for email in raw.split(",") if email.strip()]


def is_admin_user(email):
    return str(email or "").lower() in get_admin_emails()


def admin_load_table(table_name):
    supabase = get_supabase_client()
    if not supabase:
        return []

    try:
        result = supabase.table(table_name).select("*").execute()
        return result.data or []
    except Exception as exc:
        st.error(f"Admin table load error ({table_name}): {exc}")
        return []


def admin_update_user_plan(email, plan, status):
    supabase = get_supabase_client()
    if not supabase:
        return False

    rules = PLAN_RULES.get(plan, PLAN_RULES["demo"])

    try:
        supabase.table("users").update({
            "plan": plan,
            "plan_status": status,
            "analysis_limit": int(rules.get("analysis_limit", 1))
        }).eq("email", email).execute()

        supabase.table("subscriptions").upsert({
            "user_email": email,
            "plan": plan,
            "status": status
        }).execute()

        return True
    except Exception as exc:
        st.error(f"Plan update error: {exc}")
        return False


def render_admin_dashboard(user_email):
    if not is_admin_user(user_email):
        st.error("Bu sayfa sadece admin kullanıcılar içindir.")
        st.stop()

    st.markdown("## Admin Dashboard")
    st.caption("OptiFlow SaaS yönetim paneli")

    users = admin_load_table("users")
    projects = admin_load_table("projects")
    reports = admin_load_table("reports")
    subscriptions = admin_load_table("subscriptions")

    total_users = len(users)
    total_projects = len(projects)
    total_reports = len(reports)
    active_subs = len([s for s in subscriptions if str(s.get("status", "")).lower() in ["active", "trialing"]])
    demo_users = len([u for u in users if str(u.get("plan", "demo")).lower() == "demo"])
    professional_users = len([u for u in users if str(u.get("plan", "")).lower() == "professional"])
    enterprise_users = len([u for u in users if str(u.get("plan", "")).lower() == "enterprise"])

    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Toplam Kullanıcı", total_users)
    c2.metric("Aktif Abone", active_subs)
    c3.metric("Demo Kullanıcı", demo_users)
    c4.metric("Toplam Analiz", total_projects)
    c5.metric("Toplam Rapor", total_reports)

    st.markdown("### Plan Dağılımı")
    plan_rows = []
    for plan in ["demo", "starter", "professional", "enterprise"]:
        count = len([u for u in users if str(u.get("plan", "demo")).lower() == plan])
        plan_rows.append({"Plan": plan, "Kullanıcı Sayısı": count})

    st.dataframe(plan_rows, use_container_width=True, hide_index=True)

    st.markdown("### Kullanıcı Yönetimi")

    if not users:
        st.info("Henüz kullanıcı yok.")
    else:
        user_options = [u.get("email") for u in users if u.get("email")]
        selected_email = st.selectbox("Kullanıcı seç", user_options)

        selected_user = next((u for u in users if u.get("email") == selected_email), {})
        st.json(selected_user)

        col_plan, col_status, col_button = st.columns([2, 2, 1])

        with col_plan:
            new_plan = st.selectbox(
                "Plan",
                ["demo", "starter", "professional", "enterprise"],
                index=["demo", "starter", "professional", "enterprise"].index(
                    str(selected_user.get("plan", "demo")).lower()
                    if str(selected_user.get("plan", "demo")).lower() in ["demo", "starter", "professional", "enterprise"]
                    else "demo"
                )
            )

        with col_status:
            new_status = st.selectbox(
                "Status",
                ["inactive", "active", "trialing", "canceled"],
                index=1 if str(selected_user.get("plan_status", "inactive")).lower() == "active" else 0
            )

        with col_button:
            st.write("")
            st.write("")
            if st.button("Güncelle", type="primary"):
                ok = admin_update_user_plan(selected_email, new_plan, new_status)
                if ok:
                    st.success("Kullanıcı planı güncellendi.")
                    st.rerun()

    st.markdown("### Son Analizler")
    if projects:
        project_df = pd.DataFrame(projects)
        columns = [c for c in ["user_email", "company_name", "sector", "score", "risk_level", "annual_saving", "created_at"] if c in project_df.columns]
        st.dataframe(project_df[columns], use_container_width=True, hide_index=True)
    else:
        st.info("Henüz analiz kaydı yok.")

    st.markdown("### Rapor Kayıtları")
    if reports:
        report_df = pd.DataFrame(reports)
        columns = [c for c in ["user_email", "report_type", "file_name", "created_at"] if c in report_df.columns]
        st.dataframe(report_df[columns], use_container_width=True, hide_index=True)
    else:
        st.info("Henüz rapor kaydı yok.")



# ============================================================
# AUTH GATE
# ============================================================

if "demo_mode" not in st.session_state:
    st.session_state.demo_mode = False

if not is_logged_in() and not st.session_state.get("demo_mode", False):
    render_public_landing()
    st.stop()

user_email = current_user_email()
active_plan, active_rules, user_profile, subscription = resolve_active_plan(user_email)
can_analyze, analyses_used, analyses_limit = can_create_analysis(user_email, active_rules)

render_user_bar(user_email, active_plan, active_rules, analyses_used, analyses_limit)


if page == "Landing Page":
    st.markdown(
        """
<div class="hero">
    <div class="hero-title">OptiFlow Enterprise SaaS V11</div>
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
    <div class="hero-title">OptiFlow Enterprise SaaS V11</div>
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
    if not active_rules.get("ai", False):
        render_locked_feature("AI Copilot", active_plan)
    else:
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
    st.markdown("### My Client Projects")

    projects = load_user_projects(user_email)

    if not projects:
        st.info("No saved projects yet. Create a report from Report Center to save a project.")
    else:
        st.dataframe(
            pd.DataFrame(projects)[[
                "company_name",
                "sector",
                "score",
                "risk_level",
                "annual_saving",
                "created_at"
            ]],
            use_container_width=True,
            hide_index=True
        )

        selected_company = st.selectbox(
            "Select company",
            sorted(list(set([p.get("company_name", "-") for p in projects])))
        )

        for project in [p for p in projects if p.get("company_name") == selected_company]:
            with st.expander(f"{project.get('created_at')} | Score: {project.get('score')}/100"):
                st.json(project)


elif page == "Report Center":
    st.markdown("### Enterprise Report Center")
    st.write("PDF, PowerPoint ve Excel çıktıları plan seviyesine göre açılır.")

    if not can_analyze:
        st.error("Analysis limit reached for your current plan.")
        render_locked_feature("Additional analyses", active_plan)
        st.stop()

    col_pdf, col_ppt = st.columns(2)

    with col_pdf:
        if not active_rules.get("pdf", False):
            render_locked_feature("PDF / Excel Export", active_plan)
        elif st.button("Enterprise PDF Raporu Oluştur", type="primary"):
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

                excel_file = create_excel_report(
                    company_name=company_name,
                    sector=sector,
                    score=score,
                    maturity=maturity,
                    company_metrics=company_metrics,
                    benchmark_result=benchmark_result,
                    financial_result=financial_result,
                    recommendations=recommendations,
                    risk_score=risk_score,
                    risk_level=risk_level
                )

                supabase_project_id = save_project_to_supabase(
                    user_email=user_email,
                    company_name=company_name,
                    sector=sector,
                    score=score,
                    risk_level=risk_level,
                    financial_result=financial_result
                )

                if supabase_project_id:
                    save_report_to_supabase(
                        user_email=user_email,
                        project_id=supabase_project_id,
                        report_type="pdf",
                        file_name=os.path.basename(pdf_file),
                        file_path=pdf_file
                    )
                    save_report_to_supabase(
                        user_email=user_email,
                        project_id=supabase_project_id,
                        report_type="excel",
                        file_name=os.path.basename(excel_file),
                        file_path=excel_file
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

            if "excel_file" in locals() and excel_file:
                with open(excel_file, "rb") as file:
                    st.download_button(
                        label="Enterprise Excel Data İndir",
                        data=file,
                        file_name=os.path.basename(excel_file),
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

    with col_ppt:
        if not active_rules.get("ppt", False):
            render_locked_feature("PowerPoint Export", active_plan)
        elif st.button("Executive PPTX Oluştur"):
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


elif page == "Billing":
    st.markdown("### Billing & Subscription")

    st.write(f"Current user: **{user_email}**")
    st.write(f"Current plan: **{active_rules.get('label', active_plan)}**")
    st.write(f"Subscription status: **{subscription.get('status', 'inactive')}**")
    st.write(f"Usage: **{analyses_used}/{analyses_limit if analyses_limit < 999999 else 'Unlimited'} analyses**")

    st.markdown("### Upgrade Plans")

    c1, c2, c3 = st.columns(3)

    with c1:
        st.markdown("#### Starter")
        st.write("$49/month")
        st.write("10 analyses/month")
        st.write("PDF + Excel exports")
        link = st.secrets.get("STRIPE_STARTER_LINK", "")
        if link:
            st.link_button("Buy Starter", link)
        else:
            st.info("Stripe link not configured yet.")

    with c2:
        st.markdown("#### Professional")
        st.write("$149/month")
        st.write("100 analyses/month")
        st.write("PDF + PPT + Excel + AI")
        link = st.secrets.get("STRIPE_PROFESSIONAL_LINK", "")
        if link:
            st.link_button("Buy Professional", link)
        else:
            st.info("Stripe link not configured yet.")

    with c3:
        st.markdown("#### Enterprise")
        st.write("Custom")
        st.write("Unlimited usage")
        st.write("Team support + custom deployment")
        link = st.secrets.get("ENTERPRISE_CONTACT_LINK", "")
        if link:
            st.link_button("Contact Sales", link)
        else:
            st.info("Contact link not configured yet.")

    st.caption("After payment, subscription activation can be automated with Stripe webhooks. For now, update the subscriptions table manually or via webhook service.")


elif page == "Admin Panel":
    render_admin_dashboard(user_email)
