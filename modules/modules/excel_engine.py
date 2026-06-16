import pandas as pd


def process_excel(uploaded_file):

    df = pd.read_excel(uploaded_file)

    total_process_time = df["Süre (dk)"].sum()
    total_wait_time = df["Bekleme Süresi"].sum()

    wait_rate = round(
        (total_wait_time / (total_process_time + total_wait_time)) * 100,
        2
    )

    avg_quality = round(df["Kalite Oranı"].mean(), 2)

    avg_defect = round(df["Fire Oranı"].mean(), 2)

    operator_count = df["Operatör Sayısı"].sum()

    machine_count = df["Makine Sayısı"].sum()

    daily_output = df["Günlük Adet"].sum()

    oee_estimate = round(
        max(
            40,
            100
            - wait_rate * 0.6
            - avg_defect * 2
        ),
        2
    )

    line_balance_loss = round(
        min(
            40,
            wait_rate * 0.7
        ),
        2
    )

    return {
        "wait_rate": wait_rate,
        "oee": oee_estimate,
        "defect_rate": avg_defect,
        "quality_rate": avg_quality,
        "line_balance_loss": line_balance_loss,
        "daily_output": daily_output,
        "operator_count": operator_count,
        "machine_count": machine_count
    }