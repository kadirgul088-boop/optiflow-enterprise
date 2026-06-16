def generate_recommendations(
    wait_rate,
    oee,
    line_balance_loss
):

    recommendations = []

    if wait_rate > 20:
        recommendations.append(
            "Bekleme oranı kritik seviyededir. Süreçler arası akış yeniden tasarlanmalıdır."
        )

    elif wait_rate > 10:
        recommendations.append(
            "Bekleme süreleri sektör ortalamasının üzerindedir. Süreç senkronizasyonu önerilir."
        )

    if oee < 70:
        recommendations.append(
            "OEE seviyesi düşük. Kullanılabilirlik ve performans kayıpları analiz edilmelidir."
        )

    if line_balance_loss > 15:
        recommendations.append(
            "Hat dengeleme çalışması yapılmalıdır."
        )

    if len(recommendations) == 0:
        recommendations.append(
            "Operasyonel performans sektör ortalamasının üzerindedir. Sürekli iyileştirme önerilir."
        )

    return recommendations