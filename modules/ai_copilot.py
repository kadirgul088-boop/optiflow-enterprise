import os

import streamlit as st
from openai import OpenAI


def _get_api_key():
    try:
        key = st.secrets.get("OPENAI_API_KEY", None)
        if key:
            return key
    except Exception:
        pass

    return os.getenv("OPENAI_API_KEY")


def build_context(
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
    return f"""
Firma: {company_name}
Sektör: {sector}

OptiFlow Score: {score}/100
Operasyonel Olgunluk: {maturity.get("Seviye", "-")}
Olgunluk Yorumu: {maturity.get("Yorum", "-")}

Operasyonel KPI:
- Bekleme Oranı: %{company_metrics.get("wait_rate", 0)}
- OEE: %{company_metrics.get("oee", 0)}
- Hata / Fire Oranı: %{company_metrics.get("defect_rate", 0)}
- Hat Denge Kaybı: %{company_metrics.get("line_balance_loss", 0)}

Risk:
- Risk Skoru: {risk_score}/100
- Risk Seviyesi: {risk_level}

Finansal Etki:
- Toplam Operasyonel Kayıp: {financial_result.get("Toplam Operasyonel Kayıp", 0)} TL
- İyileştirme Potansiyeli: {financial_result.get("İyileştirme Potansiyeli", 0)} TL
- Tahmini Yıllık Tasarruf: {financial_result.get("Tahmini Yıllık Tasarruf", 0)} TL
- ROI: %{financial_result.get("ROI (%)", 0)}
- Geri Dönüş Süresi: {financial_result.get("Geri Dönüş Süresi (Ay)", 0)} ay

Öneriler:
{chr(10).join([f"- {r}" for r in recommendations])}
"""


def ask_real_ai_copilot(
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
    api_key = _get_api_key()

    if not api_key:
        return (
            "OPENAI_API_KEY bulunamadı. Lütfen lokal kullanım için .streamlit/secrets.toml dosyasına "
            'OPENAI_API_KEY="sk-..." ekleyin veya Streamlit Cloud > App settings > Secrets alanına girin.'
        )

    client = OpenAI(api_key=api_key)

    context = build_context(
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

    system_prompt = """
Sen kıdemli bir operasyonel mükemmellik, endüstri mühendisliği ve yönetim danışmanlığı uzmanısın.
Cevaplarını CEO, COO ve fabrika müdürü seviyesinde; net, aksiyon odaklı ve finansal etkiyi vurgulayarak ver.
Gereksiz uzun yazma. Sayısal verileri kullan. Risk, ROI, bekleme, OEE ve 30-60-90 gün aksiyon planı mantığıyla konuş.
Varsayım yaparsan bunu açıkça belirt.
"""

    user_prompt = f"""
Aşağıdaki OptiFlow analiz verilerine göre yönetici sorusunu yanıtla.

ANALİZ VERİLERİ:
{context}

YÖNETİCİ SORUSU:
{question}

Cevap formatı:
1. Kısa Yönetici Cevabı
2. Kritik Bulgular
3. Finansal Etki
4. İlk 3 Aksiyon
5. 30-60-90 Gün Önerisi
"""

    response = client.chat.completions.create(
        model="gpt-4o-mini",
        messages=[
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt}
        ],
        temperature=0.35,
        max_tokens=900
    )

    return response.choices[0].message.content
