import json

def load_benchmarks():
    with open("database/benchmark_data.json", "r", encoding="utf-8") as file:
        return json.load(file)


def get_sector_benchmark(sector):
    benchmarks = load_benchmarks()
    return benchmarks.get(sector, benchmarks["Genel Uretim"])


def compare_to_benchmark(company_metrics, sector):
    benchmark = get_sector_benchmark(sector)

    return {
        "Firma Bekleme Oranı": company_metrics["wait_rate"],
        "Sektör Bekleme Oranı": benchmark["bekleme_orani"],
        "Firma OEE": company_metrics["oee"],
        "Sektör OEE": benchmark["oee"],
        "Firma Hata Oranı": company_metrics["defect_rate"],
        "Sektör Hata Oranı": benchmark["hata_orani"],
        "Firma Hat Denge Kaybı": company_metrics["line_balance_loss"],
        "Sektör Hat Denge Kaybı": benchmark["hat_denge_kaybi"]
    }