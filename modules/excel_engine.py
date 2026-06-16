import pandas as pd


def _find_col(df, candidates):
    lower_map = {str(c).strip().lower(): c for c in df.columns}
    for candidate in candidates:
        key = candidate.strip().lower()
        if key in lower_map:
            return lower_map[key]
    return None


def _mean_if_exists(df, candidates, default=0):
    col = _find_col(df, candidates)
    if col is None:
        return default
    try:
        return float(pd.to_numeric(df[col], errors="coerce").dropna().mean())
    except Exception:
        return default


def _sum_if_exists(df, candidates, default=0):
    col = _find_col(df, candidates)
    if col is None:
        return default
    try:
        return float(pd.to_numeric(df[col], errors="coerce").dropna().sum())
    except Exception:
        return default


def process_uploaded_excel(uploaded_file):
    df = pd.read_excel(uploaded_file)
    wait_minutes_total = _sum_if_exists(df, ["Bekleme Süresi", "Bekleme Süresi (dk)", "Toplam Bekleme", "Wait Time", "Wait Minutes"], 120)
    wait_rate = _mean_if_exists(df, ["Bekleme Oranı", "Bekleme Oranı (%)", "Wait Rate", "Wait Rate (%)"], 27)
    oee = _mean_if_exists(df, ["OEE", "OEE (%)", "Equipment Effectiveness"], 68)
    defect_rate = _mean_if_exists(df, ["Hata Oranı", "Fire Oranı", "Fire Oranı (%)", "Defect Rate", "Defect Rate (%)"], 5)
    line_balance_loss = _mean_if_exists(df, ["Hat Denge Kaybı", "Hat Denge Kaybı (%)", "Line Balance Loss", "Balance Loss"], 22)
    hourly_labor_cost = _mean_if_exists(df, ["Saatlik İşçilik Maliyeti", "Saatlik İşçilik Maliyeti (TL)", "Hourly Labor Cost"], 250)
    working_days = _mean_if_exists(df, ["Aylık Çalışma Günü", "Çalışma Günü", "Working Days"], 22)
    improvement_rate = _mean_if_exists(df, ["İyileştirme Oranı", "İyileştirme Oranı (%)", "Improvement Rate"], 20)
    capacity_score = _mean_if_exists(df, ["Kapasite Skoru", "Kapasite Skoru (%)", "Capacity Score"], 75)

    process_col = _find_col(df, ["Süreç Adımı", "Proses", "Process", "Process Step"])
    wait_col = _find_col(df, ["Bekleme Süresi", "Bekleme Süresi (dk)", "Wait Time", "Wait Minutes"])
    bottleneck = None
    pareto_data = None
    if process_col is not None and wait_col is not None:
        temp = df[[process_col, wait_col]].copy()
        temp[wait_col] = pd.to_numeric(temp[wait_col], errors="coerce")
        temp = temp.dropna().sort_values(wait_col, ascending=False)
        if not temp.empty:
            bottleneck = {"process": str(temp.iloc[0][process_col]), "wait_minutes": float(temp.iloc[0][wait_col])}
            pareto_data = temp.rename(columns={process_col: "Süreç Adımı", wait_col: "Bekleme Süresi"}).head(10)

    return {
        "dataframe": df,
        "metrics": {
            "wait_rate": round(wait_rate, 2),
            "oee": round(oee, 2),
            "defect_rate": round(defect_rate, 2),
            "line_balance_loss": round(line_balance_loss, 2),
            "capacity_score": round(capacity_score, 2)
        },
        "financial_inputs": {
            "total_wait_minutes": round(wait_minutes_total, 2),
            "hourly_labor_cost": round(hourly_labor_cost, 2),
            "working_days": int(round(working_days)),
            "improvement_rate": round(improvement_rate, 2)
        },
        "bottleneck": bottleneck,
        "pareto_data": pareto_data
    }


def create_template_excel(path="OptiFlow_Excel_Template.xlsx"):
    data = {
        "Süreç Adımı": ["Kesim", "Dikim", "Kalite Kontrol", "Paketleme"],
        "Bekleme Süresi": [30, 45, 25, 20],
        "Bekleme Oranı": [22, 28, 18, 15],
        "OEE": [70, 65, 75, 80],
        "Hata Oranı": [4, 6, 5, 3],
        "Hat Denge Kaybı": [18, 24, 14, 10],
        "Kapasite Skoru": [75, 72, 78, 80],
        "Saatlik İşçilik Maliyeti": [250, 250, 250, 250],
        "Aylık Çalışma Günü": [22, 22, 22, 22],
        "İyileştirme Oranı": [20, 20, 20, 20]
    }
    df = pd.DataFrame(data)
    df.to_excel(path, index=False)
    return path
