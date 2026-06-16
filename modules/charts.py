import matplotlib.pyplot as plt
import numpy as np


def create_radar_chart(scores, filename="radar_chart.png"):
    labels = list(scores.keys())
    values = list(scores.values())
    values += values[:1]

    angles = np.linspace(0, 2 * np.pi, len(labels), endpoint=False).tolist()
    angles += angles[:1]

    fig = plt.figure(figsize=(7, 7))
    ax = plt.subplot(111, polar=True)

    ax.plot(angles, values, linewidth=3)
    ax.fill(angles, values, alpha=0.25)

    ax.set_xticks(angles[:-1])
    ax.set_xticklabels(labels)
    ax.set_ylim(0, 100)

    plt.tight_layout()
    plt.savefig(filename, dpi=300)
    plt.close()
    return filename


def create_benchmark_chart(benchmark_result, filename="benchmark_chart.png"):
    labels = []
    firm_values = []
    sector_values = []

    for key, value in benchmark_result.items():
        if str(key).startswith("Firma"):
            labels.append(str(key).replace("Firma ", ""))
            firm_values.append(value)
            sector_key = str(key).replace("Firma", "Sektör")
            sector_values.append(benchmark_result.get(sector_key, 0))

    if not labels:
        labels = ["Bekleme", "OEE", "Hata", "Denge"]
        firm_values = [0, 0, 0, 0]
        sector_values = [0, 0, 0, 0]

    x = np.arange(len(labels))

    plt.figure(figsize=(10, 6))
    plt.bar(x - 0.2, firm_values, width=0.4, label="Firma")
    plt.bar(x + 0.2, sector_values, width=0.4, label="Sektör")

    plt.xticks(x, labels, rotation=20)
    plt.legend()
    plt.tight_layout()
    plt.savefig(filename, dpi=300)
    plt.close()
    return filename


def create_financial_chart(financial_result, filename="financial_chart.png"):
    labels = []
    values = []

    keys = [
        "Toplam Operasyonel Kayıp",
        "İyileştirme Potansiyeli",
        "Tahmini Yıllık Tasarruf"
    ]

    for key in keys:
        if key in financial_result:
            labels.append(key)
            values.append(float(financial_result.get(key, 0)))

    if not labels:
        labels = ["Tasarruf"]
        values = [0]

    plt.figure(figsize=(10, 6))
    plt.bar(labels, values)
    plt.xticks(rotation=20)
    plt.tight_layout()
    plt.savefig(filename, dpi=300)
    plt.close()
    return filename


def create_risk_matrix(company_metrics, filename="risk_matrix.png"):
    matrix = np.array([
        [company_metrics.get("wait_rate", 0), company_metrics.get("defect_rate", 0)],
        [company_metrics.get("line_balance_loss", 0), company_metrics.get("oee", 0)]
    ])

    plt.figure(figsize=(6, 5))
    plt.imshow(matrix)
    plt.colorbar()
    plt.xticks([0, 1], ["Bekleme", "Kalite"])
    plt.yticks([0, 1], ["Denge", "OEE"])
    plt.title("Risk Matrix")
    plt.tight_layout()
    plt.savefig(filename, dpi=300)
    plt.close()
    return filename


def create_pareto_chart(pareto_data, filename="pareto_chart.png"):
    plt.figure(figsize=(10, 6))
    plt.bar(pareto_data["Süreç Adımı"], pareto_data["Bekleme Süresi"])
    plt.xticks(rotation=35)
    plt.title("Pareto Analizi")
    plt.tight_layout()
    plt.savefig(filename, dpi=300)
    plt.close()
    return filename


def create_risk_heatmap(risk_data, filename="risk_heatmap.png"):
    matrix = np.array([
        [risk_data.get("Teslimat Riski", 0), risk_data.get("Kalite Riski", 0)],
        [risk_data.get("Kapasite Riski", 0), risk_data.get("Finansal Risk", 0)]
    ])

    plt.figure(figsize=(6, 5))
    plt.imshow(matrix)
    plt.colorbar()
    plt.xticks([0, 1], ["Teslimat", "Kalite"])
    plt.yticks([0, 1], ["Kapasite", "Finans"])
    plt.title("Risk Heatmap")
    plt.tight_layout()
    plt.savefig(filename, dpi=300)
    plt.close()
    return filename


def create_roi_chart(financial_result, filename="roi_chart.png"):
    labels = ["Kayıp", "Tasarruf"]
    values = [
        financial_result.get("Toplam Operasyonel Kayıp", 0),
        financial_result.get("İyileştirme Potansiyeli", 0)
    ]

    plt.figure(figsize=(8, 5))
    plt.bar(labels, values)
    plt.title("ROI ve Tasarruf Potansiyeli")
    plt.tight_layout()
    plt.savefig(filename, dpi=300)
    plt.close()
    return filename
