import os
from datetime import date

from reportlab.platypus import (
    SimpleDocTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
    PageBreak,
    Image,
    KeepTogether
)
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

from modules.charts import (
    create_radar_chart,
    create_benchmark_chart,
    create_financial_chart,
    create_risk_matrix
)


FONT_NAME = "Helvetica"

for path in [
    "/System/Library/Fonts/Supplemental/Arial.ttf",
    "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"
]:
    try:
        pdfmetrics.registerFont(TTFont("OptiFont", path))
        FONT_NAME = "OptiFont"
        break
    except Exception:
        pass


PRIMARY = colors.HexColor("#0f172a")
SECONDARY = colors.HexColor("#2563eb")
SLATE = colors.HexColor("#334155")
LIGHT_BG = colors.HexColor("#f8fafc")
BORDER = colors.HexColor("#cbd5e1")
GREEN = colors.HexColor("#065f46")
GREEN_BG = colors.HexColor("#ecfdf5")
RED = colors.HexColor("#991b1b")
RED_BG = colors.HexColor("#fff1f2")
AMBER = colors.HexColor("#92400e")
AMBER_BG = colors.HexColor("#fffbeb")
BLUE_BG = colors.HexColor("#eff6ff")



def _clean_text(value):
    """
    ReportLab bazı emoji, ikon ve özel sembolleri PDF'te siyah kutu olarak gösterebilir.
    Bu fonksiyon Türkçe karakterleri korur; emoji/ikonları ve riskli XML karakterlerini temizler.
    """
    import unicodedata
    from xml.sax.saxutils import escape

    if value is None:
        return ""

    text = str(value)

    replacements = {
        "🔴": "KRITIK",
        "🟡": "DIKKAT",
        "🟢": "IYI",
        "✅": "",
        "❌": "",
        "⚠️": "DIKKAT:",
        "⚠": "DIKKAT:",
        "📊": "",
        "📈": "",
        "📉": "",
        "🚀": "",
        "💰": "",
        "🎯": "",
        "🧠": "",
        "📄": "",
        "📌": "",
        "•": "-",
        "–": "-",
        "—": "-",
        "“": '"',
        "”": '"',
        "’": "'",
        "‘": "'",
        "©": "(c)",
        "®": "(R)",
        "™": "(TM)",
    }

    for old, new in replacements.items():
        text = text.replace(old, new)

    cleaned = []
    for ch in text:
        cat = unicodedata.category(ch)

        # Turkish letters are normal letters and are preserved.
        # Remove emoji/symbol categories that often render as black boxes in PDF.
        if cat in {"So", "Cs", "Co", "Cn"}:
            continue

        cleaned.append(ch)

    text = "".join(cleaned)
    text = escape(text)
    text = text.replace("\n", "<br/>")

    return text


def _clean_table_data(data):
    return [[_clean_text(cell) for cell in row] for row in data]


def _P(text, style):
    return Paragraph(_clean_text(text), style)



def _safe_get(data, key, default=0):
    try:
        value = data.get(key, default)
        if value is None:
            return default
        return value
    except Exception:
        return default


def _to_float(value, default=0):
    try:
        return float(value)
    except Exception:
        return default


def _fmt_money(value):
    try:
        return f"{float(value):,.0f} TL".replace(",", ".")
    except Exception:
        return f"{value} TL"


def _fmt_pct(value):
    try:
        return f"%{float(value):.1f}"
    except Exception:
        return f"%{value}"


def _status_wait(value):
    value = _to_float(value)
    if value > 20:
        return "KRITIK"
    if value > 12:
        return "DIKKAT"
    return "IYI"


def _status_oee(value):
    value = _to_float(value)
    if value >= 80:
        return "GUCLU"
    if value >= 65:
        return "DIKKAT"
    return "KRITIK"


def _status_defect(value):
    value = _to_float(value)
    if value <= 3:
        return "IYI"
    if value <= 7:
        return "DIKKAT"
    return "KRITIK"


def _status_roi(value):
    value = _to_float(value)
    if value >= 150:
        return "GUCLU"
    if value >= 75:
        return "DIKKAT"
    return "ZAYIF"


def _table_style(header_color=PRIMARY, body_color=LIGHT_BG):
    return TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), header_color),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("BACKGROUND", (0, 1), (-1, -1), body_color),
        ("GRID", (0, 0), (-1, -1), 0.35, BORDER),
        ("FONTNAME", (0, 0), (-1, -1), FONT_NAME),
        ("FONTSIZE", (0, 0), (-1, -1), 8.4),
        ("LEADING", (0, 0), (-1, -1), 10),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("PADDING", (0, 0), (-1, -1), 7),
    ])


def _styled_table(data, header_color=PRIMARY, body_color=LIGHT_BG, col_widths=None):
    table = Table(_clean_table_data(data), colWidths=col_widths, repeatRows=1)
    table.setStyle(_table_style(header_color=header_color, body_color=body_color))
    return table


def _section_title(story, title_text, h1):
    story.append(Spacer(1, 4))
    story.append(_P(title_text, h1))
    story.append(Spacer(1, 3))


def add_footer(canvas, doc):
    canvas.saveState()
    canvas.setFont(FONT_NAME, 7)
    canvas.setFillColor(colors.HexColor("#64748b"))
    canvas.drawString(
        1.35 * cm,
        0.8 * cm,
        "OPTIFLOW ENTERPRISE REPORT | CONFIDENTIAL"
    )
    canvas.drawRightString(
        A4[0] - 1.35 * cm,
        0.8 * cm,
        f"Page {doc.page}"
    )
    canvas.restoreState()


def _safe_image(path, width, height):
    try:
        if path and os.path.exists(path):
            return Image(path, width=width, height=height)
    except Exception:
        return None
    return None


def create_enterprise_pdf(
    company_name,
    sector,
    score,
    maturity,
    company_metrics,
    benchmark_result,
    financial_result,
    recommendations,
    consulting_report
):
    pdf_file = "OptiFlow_Enterprise_Report.pdf"

    wait_rate = _to_float(_safe_get(company_metrics, "wait_rate", 0))
    oee = _to_float(_safe_get(company_metrics, "oee", 0))
    defect_rate = _to_float(_safe_get(company_metrics, "defect_rate", 0))
    line_balance_loss = _to_float(_safe_get(company_metrics, "line_balance_loss", 0))

    yearly_saving = _safe_get(
        financial_result,
        "Tahmini Yıllık Tasarruf",
        _safe_get(financial_result, "İyileştirme Potansiyeli", 0)
    )
    total_loss = _safe_get(financial_result, "Toplam Operasyonel Kayıp", 0)
    roi = _safe_get(financial_result, "ROI (%)", 0)
    payback = _safe_get(financial_result, "Geri Dönüş Süresi (Ay)", 0)

    risk_score = round((wait_rate + line_balance_loss + defect_rate + max(0, 100 - oee)) / 4, 1)
    risk_level = "Yüksek" if risk_score >= 25 else "Orta" if risk_score >= 15 else "Düşük"

    radar_chart = create_radar_chart({
        "Verimlilik": max(0, 100 - wait_rate),
        "OEE": max(0, oee),
        "Kalite": max(0, 100 - defect_rate),
        "Akış": max(0, 100 - line_balance_loss),
        "Kapasite": 75
    })

    benchmark_chart = create_benchmark_chart(benchmark_result)
    financial_chart = create_financial_chart(financial_result)
    risk_chart = create_risk_matrix(company_metrics)

    doc = SimpleDocTemplate(
        pdf_file,
        pagesize=A4,
        rightMargin=1.35 * cm,
        leftMargin=1.35 * cm,
        topMargin=1.25 * cm,
        bottomMargin=1.25 * cm
    )

    styles = getSampleStyleSheet()

    title = ParagraphStyle(
        "title",
        parent=styles["Title"],
        fontName=FONT_NAME,
        fontSize=24,
        leading=30,
        textColor=PRIMARY,
        alignment=1
    )

    subtitle = ParagraphStyle(
        "subtitle",
        parent=styles["Normal"],
        fontName=FONT_NAME,
        fontSize=11,
        leading=16,
        textColor=SLATE,
        alignment=1
    )

    h1 = ParagraphStyle(
        "h1",
        parent=styles["Heading1"],
        fontName=FONT_NAME,
        fontSize=15.5,
        leading=19,
        textColor=PRIMARY,
        spaceAfter=8,
        spaceBefore=4
    )

    h2 = ParagraphStyle(
        "h2",
        parent=styles["Heading2"],
        fontName=FONT_NAME,
        fontSize=12.5,
        leading=16,
        textColor=SECONDARY,
        spaceAfter=6,
        spaceBefore=4
    )

    normal = ParagraphStyle(
        "normal",
        parent=styles["Normal"],
        fontName=FONT_NAME,
        fontSize=9.2,
        leading=13.2,
        textColor=colors.HexColor("#111827"),
        spaceAfter=5
    )

    small = ParagraphStyle(
        "small",
        parent=styles["Normal"],
        fontName=FONT_NAME,
        fontSize=8.2,
        leading=11.5,
        textColor=SLATE
    )

    cover_note = ParagraphStyle(
        "cover_note",
        parent=normal,
        alignment=1,
        fontName=FONT_NAME,
        fontSize=9.4,
        leading=14,
        textColor=colors.HexColor("#111827")
    )

    story = []

    # COVER
    story.append(Spacer(1, 30))
    logo = _safe_image("assets/logo.png", 9.0 * cm, 3.4 * cm)
    if logo:
        story.append(logo)
        story.append(Spacer(1, 22))
    else:
        story.append(_P("OPTIFLOW", title))
        story.append(Spacer(1, 16))

    story.append(_P("ENTERPRISE OPERATIONAL EXCELLENCE REPORT", title))
    story.append(Spacer(1, 10))
    story.append(_P(
        "Executive Consulting Assessment | Financial Impact | Risk & Transformation Roadmap",
        subtitle
    ))
    story.append(Spacer(1, 35))

    cover_table = _styled_table([
        ["CLIENT", company_name],
        ["INDUSTRY", sector],
        ["PREPARED BY", "OptiFlow Consulting"],
        ["REPORT DATE", str(date.today())],
        ["CLASSIFICATION", "CONFIDENTIAL"],
        ["REPORT TYPE", "Enterprise Assessment"]
    ], header_color=PRIMARY, col_widths=[5 * cm, 11 * cm])

    story.append(cover_table)
    story.append(Spacer(1, 26))
    story.append(_P(
        "This report provides an executive-level assessment of operational performance, process efficiency, financial impact, risk exposure and management priorities. It is designed to support data-driven decisions and identify measurable improvement opportunities.",
        cover_note
    ))
    story.append(Spacer(1, 14))
    story.append(_P(
        "Prepared exclusively for executive management use.",
        small
    ))
    story.append(Spacer(1, 110))
    story.append(_P(
        "Operational Excellence • Lean Transformation • KPI Management • Process Optimization",
        small
    ))
    story.append(PageBreak())

    # CONTENTS
    _section_title(story, "İçindekiler", h1)
    contents = [
        "1. Yönetici Özeti",
        "2. Executive Dashboard",
        "3. OptiFlow Score ve Operasyonel Olgunluk",
        "4. Operasyonel Performans Analizi",
        "5. Benchmark Analizi",
        "6. Finansal Etki Analizi",
        "7. ROI ve Yatırım Geri Dönüş Analizi",
        "8. Risk Yönetimi ve Risk Matrisi",
        "9. Kök Neden Analizi (5M)",
        "10. Yönetim Karar Matrisi",
        "11. 30-60-90 Gün Yol Haritası",
        "12. OptiFlow Danışmanlık Değerlendirmesi",
        "13. Sonuç ve Yönetim Tavsiyesi"
    ]

    for item in contents:
        story.append(_P(item, normal))
        story.append(Spacer(1, 3))

    story.append(PageBreak())

    # 1 SUMMARY
    _section_title(story, "1. Yönetici Özeti", h1)

    summary_table = _styled_table([
        ["Ana Gösterge", "Değer"],
        ["OptiFlow Score", f"{score}/100"],
        ["Operasyonel Olgunluk", maturity.get("Seviye", "-")],
        ["Bekleme Oranı", _fmt_pct(wait_rate)],
        ["OEE", _fmt_pct(oee)],
        ["Tahmini Yıllık Tasarruf", _fmt_money(yearly_saving)],
        ["Toplam Operasyonel Kayıp", _fmt_money(total_loss)],
        ["ROI", _fmt_pct(roi)],
        ["Geri Dönüş Süresi", f"{payback} Ay"],
        ["Kurumsal Risk Skoru", f"{risk_score}/100 - {risk_level}"]
    ], header_color=PRIMARY, col_widths=[7 * cm, 9 * cm])

    story.append(summary_table)
    story.append(Spacer(1, 12))

    story.append(_P(
        f"{company_name} için gerçekleştirilen OptiFlow değerlendirmesi, firmanın operasyonel performansında ölçülebilir iyileştirme potansiyeli bulunduğunu göstermektedir. Bekleme oranı {_fmt_pct(wait_rate)}, OEE seviyesi {_fmt_pct(oee)}, hata oranı {_fmt_pct(defect_rate)} ve hat denge kaybı {_fmt_pct(line_balance_loss)} olarak değerlendirilmiştir.",
        normal
    ))
    story.append(_P(
        f"Finansal analiz sonuçlarına göre tahmini yıllık tasarruf potansiyeli {_fmt_money(yearly_saving)} seviyesindedir. Bu nedenle önerilen iyileştirme aksiyonları yalnızca operasyonel verimlilik açısından değil, yatırım geri dönüşü açısından da yönetim için güçlü bir karar alanı oluşturmaktadır.",
        normal
    ))
    story.append(_P(
        "Yönetim açısından ilk öncelik; bekleme sürelerinin azaltılması, hat dengeleme çalışmalarının başlatılması, OEE takibinin sistematik hale getirilmesi ve KPI yönetim ritminin kurulmasıdır.",
        normal
    ))

    executive_table = _styled_table([
        ["Öncelikli Yönetim Bulgusu", "Değerlendirme"],
        ["En kritik performans alanı", "Bekleme ve akış dengesi"],
        ["En yüksek finansal fırsat", _fmt_money(yearly_saving)],
        ["Risk seviyesi", risk_level],
        ["Önerilen yaklaşım", "30-60-90 gün kontrollü dönüşüm programı"],
        ["Yönetim odağı", "Hızlı kazanım + sürdürülebilir KPI sistemi"]
    ], header_color=SECONDARY, body_color=BLUE_BG, col_widths=[6 * cm, 10 * cm])

    story.append(Spacer(1, 8))
    story.append(executive_table)
    story.append(PageBreak())

    # 2 DASHBOARD
    _section_title(story, "2. Executive Dashboard", h1)

    kpi_cards = _styled_table([
        ["KPI", "Değer", "Durum"],
        ["Bekleme Oranı", _fmt_pct(wait_rate), _status_wait(wait_rate)],
        ["OEE", _fmt_pct(oee), _status_oee(oee)],
        ["Hata Oranı", _fmt_pct(defect_rate), _status_defect(defect_rate)],
        ["Hat Denge Kaybı", _fmt_pct(line_balance_loss), _status_wait(line_balance_loss)],
        ["ROI", _fmt_pct(roi), _status_roi(roi)]
    ], header_color=PRIMARY, body_color=LIGHT_BG, col_widths=[5 * cm, 4 * cm, 7 * cm])

    story.append(_P("Executive KPI Summary", h2))
    story.append(kpi_cards)
    story.append(Spacer(1, 10))

    dashboard_data = [
        ["KPI", "Değer", "Yönetim Yorumu"],
        ["OptiFlow Score", f"{score}/100", "Genel operasyonel olgunluk göstergesi"],
        ["Bekleme Oranı", _fmt_pct(wait_rate), "Süreçler arası akış kaybının ana göstergesi"],
        ["OEE", _fmt_pct(oee), "Ekipman ve üretim etkinliği göstergesi"],
        ["Hata Oranı", _fmt_pct(defect_rate), "Kalite maliyeti ve yeniden işleme riski"],
        ["Hat Denge Kaybı", _fmt_pct(line_balance_loss), "İş yükü dağılımı ve darboğaz riski"],
        ["Yıllık Tasarruf Potansiyeli", _fmt_money(yearly_saving), "İyileştirme projelerinin finansal etkisi"],
        ["Kurumsal Risk Skoru", f"{risk_score}/100", "Operasyonel riskin birleşik görünümü"],
        ["Geri Dönüş Süresi", f"{payback} ay", "İyileştirme programının amortisman süresi"]
    ]

    story.append(_styled_table(
        dashboard_data,
        header_color=SLATE,
        col_widths=[4.8 * cm, 4.2 * cm, 7 * cm]
    ))
    story.append(Spacer(1, 10))
    story.append(_P(
        "Dashboard göstergeleri, yönetimin ilk bakışta operasyonun nerede değer kaybettiğini ve hangi iyileştirmelerin finansal dönüş sağlayabileceğini görmesini sağlar. Burada öne çıkan ana bulgu, bekleme ve akış kaynaklı kayıpların hem üretkenliği hem de finansal sonuçları sınırlamasıdır.",
        normal
    ))
    story.append(Spacer(1, 10))
    story.append(Image(radar_chart, width=13.5 * cm, height=10.5 * cm))
    story.append(PageBreak())

    # 3 SCORE
    _section_title(story, "3. OptiFlow Score ve Operasyonel Olgunluk", h1)

    score_table = _styled_table([
        ["Gösterge", "Değer"],
        ["OptiFlow Score", f"{score}/100"],
        ["Operasyonel Olgunluk", maturity.get("Seviye", "-")],
        ["Olgunluk Yorumu", maturity.get("Yorum", "-")]
    ], header_color=PRIMARY, col_widths=[5 * cm, 11 * cm])

    story.append(score_table)
    story.append(Spacer(1, 10))
    story.append(_P(
        "OptiFlow Score; verimlilik, kalite, kapasite, akış ve ekipman etkinliği boyutlarını birlikte değerlendirir. Skorun amacı, yalnızca mevcut performansı ölçmek değil; hangi alanların yönetim önceliği olması gerektiğini de ortaya koymaktır.",
        normal
    ))
    story.append(_P(
        "Skorun 70 puanın altında kalması durumunda iyileştirme programının yalnızca operasyonel ekipler tarafından değil, üst yönetim sponsorluğunda yürütülmesi önerilir. Çünkü bekleme, dengeleme ve kalite problemleri doğrudan maliyet, teslimat ve müşteri memnuniyeti etkisi yaratır.",
        normal
    ))

    # 4 OPERATIONS
    _section_title(story, "4. Operasyonel Performans Analizi", h1)

    perf_table = _styled_table([
        ["Metrik", "Firma Değeri", "Analitik Yorum"],
        ["Bekleme Oranı", _fmt_pct(wait_rate), "Yüksek değer süreçler arası senkronizasyon kaybına işaret eder."],
        ["OEE", _fmt_pct(oee), "Kullanılabilirlik, performans ve kalite bileşenlerinin birlikte yönetilmesi gerekir."],
        ["Hata Oranı", _fmt_pct(defect_rate), "Kalite kaybı doğrudan fire ve yeniden işleme maliyetine dönüşür."],
        ["Hat Denge Kaybı", _fmt_pct(line_balance_loss), "İş yükü dengesizliği darboğaz ve kapasite kaybı yaratır."]
    ], header_color=SLATE, col_widths=[4.2 * cm, 4 * cm, 7.8 * cm])

    story.append(perf_table)
    story.append(Spacer(1, 8))
    story.append(_P(
        "Operasyonel performans analizinde en kritik bakış açısı, metriklerin ayrı ayrı değil birbirleriyle etkileşimli değerlendirilmesidir. Bekleme oranı yüksekken hat denge kaybının da yüksek olması, sistemin yalnızca belirli istasyonlarda değil tüm akışta verimlilik kaybettiğini gösterir.",
        normal
    ))
    story.append(PageBreak())

    # 5 BENCHMARK
    _section_title(story, "5. Benchmark Analizi", h1)

    bench_data = [["Gösterge", "Firma", "Sektör"]]
    try:
        for key, value in benchmark_result.items():
            if str(key).startswith("Firma"):
                sector_key = str(key).replace("Firma", "Sektör")
                bench_data.append([
                    str(key).replace("Firma ", ""),
                    str(value),
                    str(benchmark_result.get(sector_key, "-"))
                ])
    except Exception:
        pass

    if len(bench_data) == 1:
        bench_data.extend([
            ["Bekleme Oranı", str(wait_rate), "-"],
            ["OEE", str(oee), "-"],
            ["Hata Oranı", str(defect_rate), "-"],
            ["Hat Denge Kaybı", str(line_balance_loss), "-"]
        ])

    story.append(_styled_table(bench_data, header_color=SLATE, col_widths=[6 * cm, 5 * cm, 5 * cm]))
    story.append(Spacer(1, 8))
    story.append(Image(benchmark_chart, width=15.5 * cm, height=8.2 * cm))
    story.append(Spacer(1, 8))
    story.append(_P(
        "Benchmark karşılaştırması, firmanın performansını kendi iç verileriyle sınırlı bırakmaz; sektör referanslarına göre konumlandırır. Sektör ortalamasından olumsuz ayrışan metrikler, yönetim için doğrudan yatırım ve iyileştirme önceliği anlamına gelir.",
        normal
    ))

    # 6 FINANCIAL
    _section_title(story, "6. Finansal Etki Analizi", h1)

    fin_rows = [["Finansal Gösterge", "Değer"]]
    try:
        for key, value in financial_result.items():
            if any(x in str(key) for x in ["Tasarruf", "Kayıp", "Potansiyeli"]):
                fin_rows.append([str(key), _fmt_money(value)])
            else:
                fin_rows.append([str(key), str(value)])
    except Exception:
        fin_rows.append(["Finansal veri", "-"])

    story.append(_styled_table(fin_rows, header_color=GREEN, body_color=GREEN_BG, col_widths=[8 * cm, 8 * cm]))
    story.append(Spacer(1, 8))
    story.append(Image(financial_chart, width=15.5 * cm, height=8 * cm))
    story.append(_P(
        "Finansal etki analizi, operasyonel problemlerin parasal karşılığını görünür hale getirir. Bu sayede iyileştirme projeleri yalnızca verimlilik çalışması olarak değil, yatırım geri dönüşü olan yönetim projeleri olarak değerlendirilebilir.",
        normal
    ))
    story.append(PageBreak())

    # 7 ROI
    _section_title(story, "7. ROI ve Yatırım Geri Dönüş Analizi", h1)

    roi_table = _styled_table([
        ["Gösterge", "Sonuç", "Yönetim Anlamı"],
        ["Toplam Operasyonel Kayıp", _fmt_money(total_loss), "Mevcut sistemin yıllık tahmini kayıp büyüklüğü"],
        ["İyileştirme Potansiyeli", _fmt_money(yearly_saving), "Hedeflenen iyileştirme programının yıllık kazanım alanı"],
        ["ROI", _fmt_pct(roi), "Yatırımın finansal geri dönüş oranı"],
        ["Geri Dönüş Süresi", f"{payback} ay", "Yatırımın kendini amorti etme süresi"]
    ], header_color=GREEN, body_color=GREEN_BG, col_widths=[5.5 * cm, 4 * cm, 6.5 * cm])

    story.append(roi_table)
    story.append(Spacer(1, 8))
    story.append(_P(
        f"Mevcut analiz, {_fmt_money(yearly_saving)} seviyesindeki yıllık iyileştirme potansiyelinin işletme lehine doğrudan finansal değer yaratabileceğini göstermektedir. Bu kazanım yalnızca maliyet düşüşü değil, aynı zamanda ek kapasite, daha iyi teslimat performansı ve daha düşük operasyonel dalgalanma anlamına gelir.",
        normal
    ))
    story.append(_P(
        "Yönetim açısından önerilen yaklaşım, yüksek yatırım gerektiren kapasite artışı kararlarından önce düşük maliyetli süreç dengeleme, bekleme azaltma ve KPI izleme projelerinin başlatılmasıdır.",
        normal
    ))

    # 8 RISK
    _section_title(story, "8. Risk Yönetimi ve Risk Matrisi", h1)

    risk_table = _styled_table([
        ["Risk Alanı", "Risk Göstergesi", "Yönetim Yorumu"],
        ["Teslimat Riski", _fmt_pct(wait_rate), "Bekleme ve akış kaybı termin performansını etkileyebilir."],
        ["Kalite Riski", _fmt_pct(defect_rate), "Hata oranı maliyet ve müşteri memnuniyeti riski yaratır."],
        ["Kapasite Riski", _fmt_pct(line_balance_loss), "Dengesiz hat kullanımı kaynak verimliliğini düşürür."],
        ["Finansal Risk", _fmt_money(total_loss), "Operasyonel kayıpların parasal etkisi yönetim gündemine alınmalıdır."]
    ], header_color=RED, body_color=RED_BG, col_widths=[4.2 * cm, 4 * cm, 7.8 * cm])

    story.append(risk_table)
    story.append(Spacer(1, 8))
    story.append(Image(risk_chart, width=13.5 * cm, height=8.5 * cm))
    story.append(_P(
        "Risk matrisi, operasyonel problemlerin finansal, kalite ve teslimat performansına olan etkisini birlikte değerlendirir. Risk seviyesi yükseldikçe iyileştirme projesinin üst yönetim sponsorluğu ile yürütülmesi önem kazanır.",
        normal
    ))
    story.append(PageBreak())

    # 9 ROOT CAUSE
    _section_title(story, "9. Kök Neden Analizi (5M)", h1)

    root_table = _styled_table([
        ["5M Alanı", "Olası Kök Neden", "Önerilen Müdahale"],
        ["İnsan", "Operatör dağılımı, eğitim seviyesi ve iş yükü dengesizliği", "Çoklu beceri matrisi, iş yükü dengeleme ve vardiya bazlı performans takibi"],
        ["Makine", "Plansız duruşlar, kapasite kısıtları ve bakım eksikleri", "OEE bileşen takibi, önleyici bakım planı ve duruş neden analizi"],
        ["Metot", "Standart iş eksikliği ve süreç sırası farklılıkları", "Standart iş dokümanı, iş etüdü ve hat dengeleme çalışması"],
        ["Malzeme", "Malzeme akışı gecikmeleri ve ara stok yönetimi", "Malzeme besleme planı, Kanban ve süreç içi stok kontrolü"],
        ["Ölçüm", "KPI sisteminin yetersizliği ve veri doğruluğu riski", "Günlük KPI panosu, haftalık performans toplantısı ve dijital veri doğrulama"]
    ], header_color=PRIMARY, col_widths=[3 * cm, 6.5 * cm, 6.5 * cm])

    story.append(root_table)
    story.append(Spacer(1, 8))
    story.append(_P(
        "5M analizi, görünen performans problemlerinin arkasındaki sistematik nedenleri ortaya koymak için kullanılır. Bu değerlendirme, yalnızca semptomları değil, tekrar eden kayıp kaynaklarını yönetmeyi hedefler.",
        normal
    ))
    story.append(PageBreak())

    # 10 DECISION MATRIX
    _section_title(story, "10. Yönetim Karar Matrisi", h1)

    decision_final = _styled_table([
        ["Karar Alanı", "Yönetim Tavsiyesi"],
        ["Kısa Vadeli Öncelik", "Bekleme sürelerinin azaltılması ve darboğaz noktalarının kontrol altına alınması"],
        ["Orta Vadeli Öncelik", "Hat dengeleme, OEE takibi ve standart iş uygulamalarının yaygınlaştırılması"],
        ["Finansal Öncelik", "En yüksek tasarruf potansiyeli taşıyan süreçlere yatırım yapılması"],
        ["Risk Önceliği", "Teslimat, kalite ve kapasite risklerinin haftalık KPI sistemi ile izlenmesi"],
        ["Stratejik Tavsiye", "OptiFlow çıktılarının aylık yönetim toplantılarında karar destek aracı olarak kullanılması"]
    ], header_color=PRIMARY, col_widths=[5 * cm, 11 * cm])

    story.append(decision_final)
    story.append(Spacer(1, 14))
    story.append(_P(
        "OptiFlow değerlendirmesine göre işletmenin kısa vadede en yüksek kazanımı, mevcut kaynakları daha etkin kullanarak bekleme ve akış kayıplarını azaltmasıyla elde etmesi beklenmektedir. Bu yaklaşım, yüksek maliyetli yatırım kararlarından önce operasyonel disiplinin güçlendirilmesini sağlar.",
        normal
    ))
    story.append(_P(
        "Yönetimin önerilen aksiyonları sistematik biçimde uygulaması durumunda, operasyonel performansta ölçülebilir iyileşme, maliyetlerde düşüş, kapasite kullanımında artış ve daha sürdürülebilir bir performans yönetim yapısı elde edilmesi beklenmektedir.",
        normal
    ))

    # 11 ROADMAP
    _section_title(story, "11. 30-60-90 Gün Yol Haritası", h1)

    roadmap = _styled_table([
        ["Dönem", "Odak Alanı", "Ana Faaliyetler", "Beklenen Çıktı"],
        ["0-30 Gün", "Veri doğrulama ve hızlı teşhis", "Süreç gözlemi, KPI baz çizgisi, darboğaz doğrulama", "Hızlı kazanım listesi ve ölçüm altyapısı"],
        ["30-60 Gün", "Akış ve kapasite iyileştirme", "Hat dengeleme, bekleme azaltma, standart iş", "Daha dengeli akış ve ölçülebilir kapasite kazanımı"],
        ["60-90 Gün", "Sürdürülebilir yönetim sistemi", "KPI toplantı ritmi, sorumluluk matrisi, performans panosu", "Sürekli iyileştirme sistemi ve yönetim kontrolü"]
    ], header_color=PRIMARY, col_widths=[3 * cm, 4 * cm, 5 * cm, 4 * cm])

    story.append(roadmap)
    story.append(Spacer(1, 8))
    story.append(_P(
        "Yol haritası, iyileştirme çalışmalarını kısa vadeli hızlı kazanımlar ve orta vadeli sürdürülebilir sistem tasarımı olarak iki eksende ele alır. İlk 30 gün veri doğruluğu ve önceliklendirme, sonraki 60 gün ise uygulama ve kalıcı yönetim ritmi için kritik dönemdir.",
        normal
    ))

    # 12 CONSULTING
    _section_title(story, "12. OptiFlow Danışmanlık Değerlendirmesi", h1)

    clean_report = str(consulting_report).replace("**", "").replace("#", "")
    paragraphs = [p.strip() for p in clean_report.split("\n") if p.strip()]

    for paragraph in paragraphs:
        if len(paragraph) < 90 and any(char.isdigit() for char in paragraph[:3]):
            story.append(_P(paragraph, h2))
        else:
            story.append(_P(paragraph, normal))

    story.append(PageBreak())

    # 13 CONCLUSION
    _section_title(story, "13. Sonuç ve Yönetim Tavsiyesi", h1)

    story.append(_P(
        f"OptiFlow değerlendirmesi, {company_name} için en önemli iyileştirme alanlarının bekleme süreleri, hat dengeleme, kalite kayıpları ve performans izleme sistemi olduğunu göstermektedir. Mevcut göstergeler, operasyonel kayıpların yönetim gündemine alınması durumunda ölçülebilir finansal kazanım üretilebileceğini ortaya koymaktadır.",
        normal
    ))
    story.append(_P(
        f"Yıllık {_fmt_money(yearly_saving)} seviyesindeki tasarruf potansiyeli, iyileştirme programının yalnızca operasyonel değil, finansal olarak da anlamlı olduğunu göstermektedir. Öncelikli öneri, ilk 90 gün içinde veri doğrulama, darboğaz yönetimi, hat dengeleme ve KPI takip sisteminin birlikte devreye alınmasıdır.",
        normal
    ))
    story.append(_P(
        "Bu rapor, yönetimin hangi aksiyonları önce başlatması gerektiğini, hangi riskleri önceliklendirmesi gerektiğini ve hangi finansal sonuçların hedeflenebileceğini göstermek üzere hazırlanmıştır. Uygulama başarısı için üst yönetim sahipliği, düzenli takip ritmi ve ölçülebilir KPI sistemi kritik önemdedir.",
        normal
    ))

    doc.build(
        story,
        onFirstPage=add_footer,
        onLaterPages=add_footer
    )

    return pdf_file
