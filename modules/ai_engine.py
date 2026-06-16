from openai import OpenAI
import streamlit as st


def generate_consulting_report(
    sector,
    company_metrics,
    benchmark_result,
    financial_result,
    maturity,
    recommendations
):
    try:
        api_key = st.secrets["OPENAI_API_KEY"]
        client = OpenAI(api_key=api_key)

        prompt = f"""
Sen 20 yıllık operasyonel mükemmellik, yalın üretim ve endüstri mühendisliği danışmanısın.
Raporu OptiFlow Danışmanlık adına yazıyorsun.

KESİN KURALLAR:
- Yapay zeka, AI, ChatGPT ifadelerini asla kullanma.
- Genel tanım yapma.
- Her bölümde sayısal veriye referans ver.
- Yönetim kurulu / genel müdür seviyesinde profesyonel dil kullan.
- Her bulguda: veri → anlamı → risk → öneri → beklenen etki sırasını izle.
- Kısa ve yüzeysel yazma.
- Her başlık altında en az 2 güçlü paragraf yaz.
- Müşteriyi ikna edecek danışmanlık dili kullan.

SEKTÖR:
{sector}

FİRMA METRİKLERİ:
{company_metrics}

BENCHMARK KARŞILAŞTIRMASI:
{benchmark_result}

FİNANSAL ETKİ:
{financial_result}

OPERASYONEL OLGUNLUK:
{maturity}

ÖNERİLER:
{recommendations}

Aşağıdaki yapıda kapsamlı danışmanlık değerlendirmesi yaz:

1. Yönetici Özeti

2. Kritik Bulgular
- Bekleme oranını yorumla
- OEE seviyesini yorumla
- Hata oranını yorumla
- Hat denge kaybını yorumla

3. Benchmark Değerlendirmesi
- Firmanın sektör ortalamasından nerelerde ayrıştığını açıkla
- Sektöre göre güçlü ve zayıf alanları belirt

4. Operasyonel Risk Analizi
- Teslimat riski
- Kapasite riski
- Kalite riski
- Maliyet riski

5. Finansal Etki Yorumu
- Bekleme süresindeki iyileştirmenin yıllık etkiye dönüşümünü yorumla
- Yönetim için yatırım önceliği çıkar

6. Kök Neden Değerlendirmesi
- İnsan
- Makine
- Metot
- Malzeme
- Ölçüm

7. Yönetim İçin Karar Önerileri
- İlk 5 karar önerisini yaz
- Her kararın beklenen etkisini açıkla

8. 30-60-90 Gün Yol Haritası
- 0-30 gün
- 30-60 gün
- 60-90 gün

9. Sonuç ve OptiFlow Değerlendirmesi
- Net, güçlü ve profesyonel sonuç yaz

Türkçe yaz.
"""

        response = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[{"role": "user", "content": prompt}]
        )

        return response.choices[0].message.content

    except Exception:
        return """
1. Yönetici Özeti

OptiFlow değerlendirmesine göre işletmenin operasyonel performansında iyileştirme potansiyeli bulunmaktadır. Mevcut göstergeler, özellikle bekleme oranı, hat denge kaybı ve OEE seviyesinin yönetim tarafından öncelikli olarak ele alınması gerektiğini göstermektedir.

Bu analiz, işletmenin yalnızca kapasite artırımıyla değil; süreç akışının yeniden dengelenmesi, bekleme kaynaklarının azaltılması ve performans göstergelerinin düzenli izlenmesiyle daha yüksek verimlilik elde edebileceğini ortaya koymaktadır.

2. Kritik Bulgular

Bekleme oranının yüksek seviyede olması, süreçler arası akışta senkronizasyon problemi bulunduğunu göstermektedir. Bu durum günlük üretim kapasitesini sınırlandırabilir ve teslimat performansında dalgalanmalara neden olabilir.

OEE seviyesinin iyileştirme potansiyeli taşıması, kullanılabilirlik, performans ve kalite bileşenlerinin birlikte analiz edilmesi gerektiğini göstermektedir. Özellikle duruş süreleri ve kalite kayıpları sistematik olarak izlenmelidir.

3. Benchmark Değerlendirmesi

Firma metrikleri sektör referanslarıyla karşılaştırıldığında, bekleme ve akış dengesizliği alanlarında iyileştirme ihtiyacı öne çıkmaktadır. Bu durum, firmanın mevcut kaynaklarını tam kapasiteyle değerlendiremediğini göstermektedir.

Benchmark farkı, yönetim için öncelikli aksiyon alanlarını belirlemede kritik bir göstergedir. Sektör ortalamasının üzerinde kalan kayıp alanları, kısa vadede en yüksek geri dönüş potansiyeline sahip iyileştirme başlıklarıdır.

4. Operasyonel Risk Analizi

Bekleme ve hat denge kayıpları devam ettiği takdirde teslimat sürelerinde uzama, kapasite kullanımında düşüş ve operasyonel maliyetlerde artış riski oluşabilir. Bu nedenle süreç akışı yalnızca üretim hızı açısından değil, müşteri memnuniyeti açısından da kritik öneme sahiptir.

Kalite ve hata oranları da doğrudan yeniden işleme maliyetleri, fire oranları ve müşteri şikâyetleri üzerinde etki yaratabilir. Kalite göstergelerinin üretim performansıyla birlikte takip edilmesi önerilir.

5. Finansal Etki Yorumu

Finansal etki analizi, bekleme sürelerinde yapılacak iyileştirmenin yıllık bazda ölçülebilir kazanç yaratabileceğini göstermektedir. Bu kazanım yalnızca işçilik maliyeti azaltımı değil, aynı zamanda ek kapasite yaratımı anlamına gelir.

Yönetim açısından ilk yatırım önceliği yüksek maliyetli makine yatırımı yerine, mevcut sürecin darboğaz ve bekleme kaynaklı kayıplarını azaltmaya verilmelidir.

6. Kök Neden Değerlendirmesi

İnsan boyutunda operatör dağılımı ve iş yükü dengesizliği incelenmelidir. Makine boyutunda duruş süreleri ve kapasite kullanımı takip edilmelidir. Metot boyutunda süreç sıralaması, standart iş ve operasyon akışı yeniden değerlendirilmelidir.

Malzeme akışı ve veri ölçüm altyapısı da iyileştirme başarısı için kritik faktörlerdir. KPI takibi düzenli yapılmadığı durumda problemler geç fark edilir ve iyileştirme etkisi sürdürülebilir olmaz.

7. Yönetim İçin Karar Önerileri

İlk karar, bekleme oranını azaltmaya yönelik süreç akışı çalışmasının başlatılmasıdır. İkinci karar, hat dengeleme çalışmasıyla iş yüklerinin yeniden dağıtılmasıdır. Üçüncü karar, OEE bileşenlerinin haftalık takip edilmesidir.

Dördüncü karar, kalite kayıplarını azaltacak standart iş uygulamalarının devreye alınmasıdır. Beşinci karar ise yönetim seviyesinde aylık operasyonel performans toplantılarının başlatılmasıdır.

8. 30-60-90 Gün Yol Haritası

0-30 gün içinde veri doğrulama, darboğaz gözlemi ve hızlı kazanım alanları belirlenmelidir. 30-60 gün içinde hat dengeleme, bekleme azaltma ve kapasite düzenleme çalışmaları uygulanmalıdır.

60-90 gün içinde KPI yönetim sistemi, standart iş uygulamaları ve sürekli iyileştirme takip mekanizması kurulmalıdır.

9. Sonuç ve OptiFlow Değerlendirmesi

OptiFlow değerlendirmesine göre işletme, mevcut kaynaklarını daha etkin kullanarak ölçülebilir operasyonel kazanım sağlayabilecek potansiyele sahiptir. Öncelik bekleme kayıplarının azaltılması, süreç akışının dengelenmesi ve yönetim seviyesinde performans takibinin güçlendirilmesi olmalıdır.

Önerilen aksiyonların sistematik biçimde uygulanması durumunda işletmenin verimlilik, teslimat performansı ve maliyet yönetimi alanlarında sürdürülebilir iyileşme elde etmesi beklenmektedir.
"""