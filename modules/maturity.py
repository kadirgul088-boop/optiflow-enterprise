def get_maturity_comment(score):
    if score >= 85:
        return {
            "Seviye": "Seviye 5 - Operasyonel Mükemmellik",
            "Yorum": "Süreçler yüksek olgunluk seviyesinde. Odak noktası sürekli iyileştirme ve benchmark liderliği olmalıdır."
        }
    elif score >= 70:
        return {
            "Seviye": "Seviye 4 - Yönetilen",
            "Yorum": "Süreçler genel olarak kontrol altında ancak darboğaz ve bekleme kaynaklı iyileştirme alanları bulunmaktadır."
        }
    elif score >= 55:
        return {
            "Seviye": "Seviye 3 - Kontrollü",
            "Yorum": "Süreçlerde temel kontrol mekanizmaları var ancak performans sürdürülebilirliği için KPI takibi güçlendirilmelidir."
        }
    elif score >= 40:
        return {
            "Seviye": "Seviye 2 - Temel",
            "Yorum": "Operasyonel yapı temel seviyede. Süreç standardizasyonu ve veri temelli yönetim önceliklidir."
        }
    else:
        return {
            "Seviye": "Seviye 1 - Kritik",
            "Yorum": "Süreçler kritik risk seviyesinde. Öncelikle darboğaz, bekleme ve kalite kayıpları ele alınmalıdır."
        }