def calculate_risk_profile(
    wait_rate,
    oee,
    defect_rate,
    line_balance_loss,
    capacity_utilization
):

    delivery_risk = min(
        100,
        (
            wait_rate * 0.8 +
            line_balance_loss * 0.7
        )
    )

    quality_risk = min(
        100,
        (
            defect_rate * 8
        )
    )

    capacity_risk = min(
        100,
        (
            100 - capacity_utilization
        )
    )

    productivity_risk = min(
        100,
        (
            100 - oee
        )
    )

    financial_risk = min(
        100,
        (
            delivery_risk * 0.3 +
            quality_risk * 0.3 +
            productivity_risk * 0.4
        )
    )

    strategic_risk = min(
        100,
        (
            financial_risk * 0.4 +
            capacity_risk * 0.3 +
            quality_risk * 0.3
        )
    )

    total_risk = round(
        (
            delivery_risk +
            quality_risk +
            capacity_risk +
            productivity_risk +
            financial_risk +
            strategic_risk
        ) / 6,
        2
    )

    if total_risk < 30:
        risk_level = "Düşük Risk"

    elif total_risk < 60:
        risk_level = "Orta Risk"

    else:
        risk_level = "Yüksek Risk"

    return {

        "Teslimat Riski": round(
            delivery_risk,
            2
        ),

        "Kalite Riski": round(
            quality_risk,
            2
        ),

        "Kapasite Riski": round(
            capacity_risk,
            2
        ),

        "Verimlilik Riski": round(
            productivity_risk,
            2
        ),

        "Finansal Risk": round(
            financial_risk,
            2
        ),

        "Stratejik Risk": round(
            strategic_risk,
            2
        ),

        "Toplam Risk Skoru": total_risk,

        "Risk Seviyesi": risk_level
    }