def calculate_optiflow_score(
    efficiency_score,
    oee_score,
    capacity_score,
    flow_score,
    quality_score
):

    score = (
        efficiency_score * 0.25
        + oee_score * 0.25
        + capacity_score * 0.20
        + flow_score * 0.20
        + quality_score * 0.10
    )

    return round(score, 2)


def maturity_level(score):

    if score >= 85:
        return "Seviye 5 - Operasyonel Mükemmellik"

    elif score >= 70:
        return "Seviye 4 - Yönetilen"

    elif score >= 55:
        return "Seviye 3 - Kontrollü"

    elif score >= 40:
        return "Seviye 2 - Temel"

    else:
        return "Seviye 1 - Kritik"