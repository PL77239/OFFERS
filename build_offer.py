#!/usr/bin/env python3
"""
Build a ready-to-send B2B offer for Alce Nero Organic Gluten-Free
Corn & Rice Spaghettoni (250 g), targeted at retail / distribution
buyers in Poland and other CEE markets.

Outputs (into ./offer):
  - Offer_AlceNero_Spaghettoni_EN.pdf   (English offer sheet)
  - Offer_AlceNero_Spaghettoni_PL.pdf   (Polish offer sheet)
  - Offer_AlceNero_Spaghettoni_EN.eml   (ready-to-send email, attachments included)
  - Offer_AlceNero_Spaghettoni_PL.eml   (ready-to-send email, attachments included)

The .eml files can be opened directly in Outlook / Thunderbird / Apple Mail,
reviewed, and sent. They carry the two product photos and the matching PDF
offer sheet as attachments.

Sender contact details are placeholders (SENDER_* constants below) - edit
them once here and re-run, or edit them directly in the generated .eml.
"""

import os
from datetime import datetime, date
from email.message import EmailMessage
from email.utils import formatdate, make_msgid

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image as RLImage,
    HRFlowable, KeepTogether,
)
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# Register a Unicode font family so Polish diacritics (ł, ż, ś, ą, ę, ...) render.
_FDIR = "/usr/share/fonts/truetype/liberation"
pdfmetrics.registerFont(TTFont("LSans", os.path.join(_FDIR, "LiberationSans-Regular.ttf")))
pdfmetrics.registerFont(TTFont("LSans-Bold", os.path.join(_FDIR, "LiberationSans-Bold.ttf")))
pdfmetrics.registerFont(TTFont("LSans-Italic", os.path.join(_FDIR, "LiberationSans-Italic.ttf")))
pdfmetrics.registerFont(TTFont("LSans-BoldItalic", os.path.join(_FDIR, "LiberationSans-BoldItalic.ttf")))
pdfmetrics.registerFontFamily(
    "LSans", normal="LSans", bold="LSans-Bold",
    italic="LSans-Italic", boldItalic="LSans-BoldItalic",
)
FONT = "LSans"
FONT_B = "LSans-Bold"

# --------------------------------------------------------------------------
# Source assets
# --------------------------------------------------------------------------
ROOT = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(ROOT, "offer")
os.makedirs(OUT, exist_ok=True)

PHOTO_FRONT = os.path.join(ROOT, "photo_front_Alce Nero_spaghettoni.jpeg")
PHOTO_BACK = os.path.join(ROOT, "photo_Alce Nero_spaghettoni.jpeg")

# --------------------------------------------------------------------------
# Offer data (from Offer_Alce Nero_spaghettoni_ENG.xlsx + product pack)
# --------------------------------------------------------------------------
PRODUCT_NAME = "Alce Nero Organic Gluten-Free Corn & Rice Spaghettoni"
NET_WEIGHT = "250 g"
CODE = "PSG718"
EAN = "8009004811522"
BEST_BEFORE = date(2028, 5, 13)
UNITS_TOTAL = 100_000
UNITS_PER_PALLET = 2064
PALLETS = UNITS_TOTAL / UNITS_PER_PALLET          # 48.45
RRP_UNIT = 1.95                                    # reference retail price / unit
OFFER_UNIT = 0.65                                  # our offer, EXW Milan, net VAT
OFFER_PALLET = UNITS_PER_PALLET * OFFER_UNIT       # 1,341.60
OFFER_TOTAL = UNITS_TOTAL * OFFER_UNIT             # 65,000
DISCOUNT_PCT = round((1 - OFFER_UNIT / RRP_UNIT) * 100)   # 67 %
MARKUP_X = RRP_UNIT / OFFER_UNIT                   # 3.0x
INCOTERMS = "EXW Milan (Italy)"
OFFER_VALID_DAYS = 7

# --------------------------------------------------------------------------
# Sender details - EDIT THESE (kept as clearly-marked placeholders)
# --------------------------------------------------------------------------
SENDER_NAME = "[Your Name]"
SENDER_TITLE_EN = "Marketing & Sales Specialist"
SENDER_TITLE_PL = "Specjalista ds. Marketingu i Sprzedaży"
SENDER_COMPANY = "[Your Company]"
SENDER_EMAIL = "sales@example.com"
SENDER_PHONE = "[+48 ...]"
SENDER_WEB = "www.example.com"
RECIPIENT_EMAIL = "buyer@example.com"

BRAND_GREEN = colors.HexColor("#5B7A1E")
BRAND_DARK = colors.HexColor("#3E3E3E")
LIGHT_BG = colors.HexColor("#F1F3E8")
ACCENT = colors.HexColor("#B23A1E")


def eur(x, dec=2):
    s = f"{x:,.{dec}f}"
    return "€" + s


# ==========================================================================
# PDF builder
# ==========================================================================
def build_pdf(lang):
    fname = os.path.join(OUT, f"Offer_AlceNero_Spaghettoni_{lang}.pdf")
    doc = SimpleDocTemplate(
        fname, pagesize=A4,
        leftMargin=18 * mm, rightMargin=18 * mm,
        topMargin=15 * mm, bottomMargin=15 * mm,
        title=f"{PRODUCT_NAME} - Offer",
        author=SENDER_COMPANY,
    )
    styles = getSampleStyleSheet()
    h1 = ParagraphStyle("h1", parent=styles["Title"], textColor=BRAND_GREEN,
                        fontName=FONT_B, fontSize=20, leading=23, spaceAfter=2)
    sub = ParagraphStyle("sub", parent=styles["Normal"], textColor=BRAND_DARK,
                         fontName=FONT, fontSize=10.5, leading=13)
    h2 = ParagraphStyle("h2", parent=styles["Heading2"], textColor=BRAND_GREEN,
                        fontName=FONT_B, fontSize=12.5, leading=15, spaceBefore=8, spaceAfter=4)
    body = ParagraphStyle("body", parent=styles["Normal"], fontName=FONT, fontSize=9.5,
                          leading=13, textColor=BRAND_DARK)
    small = ParagraphStyle("small", parent=styles["Normal"], fontName=FONT, fontSize=7.5,
                           leading=9.5, textColor=colors.HexColor("#7A7A7A"))
    cell = ParagraphStyle("cell", parent=styles["Normal"], fontName=FONT, fontSize=9, leading=11)
    cellb = ParagraphStyle("cellb", parent=cell, fontName=FONT_B)

    bb = BEST_BEFORE.strftime("%d.%m.%Y" if lang == "PL" else "%d %b %Y")

    T = {
        "EN": dict(
            kicker="COMMERCIAL OFFER  |  Italian Organic Pasta",
            title="Alce Nero Organic Gluten-Free<br/>Corn &amp; Rice Spaghettoni, 250 g",
            subtitle=("Premium Italian organic brand — a limited stock lot offered "
                      "well below retail for the Polish and CEE market."),
            spec_h="Product specification",
            specs=[
                ("Product", f"{PRODUCT_NAME}, {NET_WEIGHT}"),
                ("Brand", "Alce Nero (Italy) — premium organic"),
                ("Product code", CODE),
                ("EAN / barcode", EAN),
                ("Ingredients", "Corn flour 80%*, rice flour 20%*  (*organic)"),
                ("Certification", "EU Organic (ICEA), Gluten-Free / suitable for coeliacs, HALAL"),
                ("Nutrition", "355 kcal / 100 g · cooking time 16–18 min"),
                ("Origin", "Product of Italy"),
                ("Best before", bb),
            ],
            terms_h="Commercial terms",
            terms_head=["", "Per unit", "Per pallet", "Full lot"],
            terms_rows=[
                ["Units", "1", f"{UNITS_PER_PALLET:,}", f"{UNITS_TOTAL:,}"],
                ["Reference retail price", eur(RRP_UNIT), eur(RRP_UNIT*UNITS_PER_PALLET), eur(RRP_UNIT*UNITS_TOTAL, 0)],
                ["OUR OFFER PRICE (net VAT)", eur(OFFER_UNIT), eur(OFFER_PALLET), eur(OFFER_TOTAL, 0)],
            ],
            terms_note=(f"Incoterms: <b>{INCOTERMS}</b> · Price net of VAT · "
                        f"Availability: <b>{UNITS_TOTAL:,} units (~{PALLETS:.0f} pallets, {UNITS_PER_PALLET:,} pcs/pallet)</b> · "
                        f"Offer valid {OFFER_VALID_DAYS} days, subject to prior sale."),
            why_h="Why this is a strong buy",
            why=[
                f"<b>~{DISCOUNT_PCT}% below reference retail</b> — buy at {eur(OFFER_UNIT)}, "
                f"suggested RRP {eur(RRP_UNIT)} (≈ {MARKUP_X:.1f}× mark-up potential).",
                "<b>Premium Italian organic brand</b> Alce Nero — trusted in retail, delis and HORECA.",
                "<b>Two fast-growing segments in one SKU:</b> certified organic <i>and</i> gluten-free "
                "(suitable for coeliacs).",
                f"<b>Long shelf life</b> — best before {bb}, ample runway for retail rotation and promotions.",
                "<b>Flexible volumes</b> — from single pallets up to the full 100,000-unit lot.",
            ],
            photos_h="Product photos",
            cta=(f"To reserve stock or request a pro-forma invoice, reply to this offer "
                 f"or contact {SENDER_NAME} at {SENDER_EMAIL}."),
            foot=(f"{SENDER_COMPANY} · {SENDER_NAME}, {SENDER_TITLE_EN} · "
                  f"{SENDER_EMAIL} · {SENDER_PHONE} · {SENDER_WEB}"),
            date_lbl="Offer date",
        ),
        "PL": dict(
            kicker="OFERTA HANDLOWA  |  Włoski makaron ekologiczny",
            title="Alce Nero Ekologiczne Bezglutenowe<br/>Spaghettoni Kukurydziano-Ryżowe, 250 g",
            subtitle=("Renomowana włoska marka ekologiczna — ograniczona partia towaru "
                      "w cenie znacznie poniżej detalu, dla rynku Polski i Europy Środkowo-Wschodniej."),
            spec_h="Specyfikacja produktu",
            specs=[
                ("Produkt", f"{PRODUCT_NAME}, {NET_WEIGHT}"),
                ("Marka", "Alce Nero (Włochy) — segment premium/eko"),
                ("Kod produktu", CODE),
                ("EAN / kod kreskowy", EAN),
                ("Składniki", "Mąka kukurydziana 80%*, mąka ryżowa 20%*  (*ekologiczne)"),
                ("Certyfikaty", "Ekologiczny UE (ICEA), bezglutenowy / dla celiaków, HALAL"),
                ("Wartości odżywcze", "355 kcal / 100 g · czas gotowania 16–18 min"),
                ("Pochodzenie", "Produkt Włoch"),
                ("Najlepiej spożyć przed", bb),
            ],
            terms_h="Warunki handlowe",
            terms_head=["", "Za sztukę", "Za paletę", "Cała partia"],
            terms_rows=[
                ["Ilość (szt.)", "1", f"{UNITS_PER_PALLET:,}", f"{UNITS_TOTAL:,}"],
                ["Referencyjna cena detaliczna", eur(RRP_UNIT), eur(RRP_UNIT*UNITS_PER_PALLET), eur(RRP_UNIT*UNITS_TOTAL, 0)],
                ["NASZA CENA OFERTOWA (netto)", eur(OFFER_UNIT), eur(OFFER_PALLET), eur(OFFER_TOTAL, 0)],
            ],
            terms_note=(f"Incoterms: <b>{INCOTERMS}</b> · Ceny netto (bez VAT) · "
                        f"Dostępność: <b>{UNITS_TOTAL:,} szt. (~{PALLETS:.0f} palet, {UNITS_PER_PALLET:,} szt./paleta)</b> · "
                        f"Oferta ważna {OFFER_VALID_DAYS} dni, do wyczerpania zapasów."),
            why_h="Dlaczego warto",
            why=[
                f"<b>~{DISCOUNT_PCT}% poniżej ceny detalicznej</b> — zakup po {eur(OFFER_UNIT)}, "
                f"sugerowana cena detaliczna {eur(RRP_UNIT)} (potencjał marży ≈ {MARKUP_X:.1f}×).",
                "<b>Renomowana włoska marka eko</b> Alce Nero — rozpoznawalna w handlu, delikatesach i HORECA.",
                "<b>Dwa dynamicznie rosnące segmenty w jednym produkcie:</b> certyfikat ekologiczny "
                "<i>oraz</i> produkt bezglutenowy (dla celiaków).",
                f"<b>Długi termin przydatności</b> — najlepiej spożyć przed {bb}, dużo czasu na rotację i promocje.",
                "<b>Elastyczne wolumeny</b> — od pojedynczych palet do pełnej partii 100 000 szt.",
            ],
            photos_h="Zdjęcia produktu",
            cta=(f"Aby zarezerwować towar lub otrzymać fakturę pro forma, prosimy o odpowiedź "
                 f"na tę ofertę lub kontakt: {SENDER_NAME}, {SENDER_EMAIL}."),
            foot=(f"{SENDER_COMPANY} · {SENDER_NAME}, {SENDER_TITLE_PL} · "
                  f"{SENDER_EMAIL} · {SENDER_PHONE} · {SENDER_WEB}"),
            date_lbl="Data oferty",
        ),
    }[lang]

    story = []

    # Header band
    story.append(Paragraph(T["kicker"], sub))
    story.append(Spacer(1, 2))
    story.append(Paragraph(T["title"], h1))
    story.append(Paragraph(T["subtitle"], sub))
    story.append(Spacer(1, 4))
    story.append(HRFlowable(width="100%", thickness=1.4, color=BRAND_GREEN,
                            spaceBefore=2, spaceAfter=8))

    # Two-column: front photo + specs
    front = RLImage(PHOTO_FRONT)
    fw, fh = front.imageWidth, front.imageHeight
    target_w = 46 * mm
    front.drawWidth = target_w
    front.drawHeight = target_w * fh / fw

    spec_data = [[Paragraph(f"<b>{k}</b>", cell), Paragraph(str(v), cell)]
                 for k, v in T["specs"]]
    spec_tbl = Table(spec_data, colWidths=[38 * mm, 70 * mm])
    spec_tbl.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LINEBELOW", (0, 0), (-1, -2), 0.4, colors.HexColor("#DDDDDD")),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("BACKGROUND", (0, 0), (-1, -1), colors.white),
    ]))

    spec_block = [Paragraph(T["spec_h"], h2), spec_tbl]
    top = Table([[front, spec_block]], colWidths=[50 * mm, 110 * mm])
    top.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (0, 0), 0),
        ("LEFTPADDING", (1, 0), (1, 0), 8),
    ]))
    story.append(top)
    story.append(Spacer(1, 8))

    # Commercial terms table
    story.append(Paragraph(T["terms_h"], h2))
    terms_data = [[Paragraph(f"<b>{c}</b>", cellb if i else cell) for i, c in enumerate(T["terms_head"])]]
    for r in T["terms_rows"]:
        terms_data.append([Paragraph(str(c), cell) for c in r])
    terms_tbl = Table(terms_data, colWidths=[62 * mm, 30 * mm, 34 * mm, 34 * mm])
    terms_tbl.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), BRAND_GREEN),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), FONT_B),
        ("BACKGROUND", (0, 3), (-1, 3), LIGHT_BG),
        ("FONTNAME", (0, 3), (-1, 3), FONT_B),
        ("TEXTCOLOR", (0, 3), (-1, 3), ACCENT),
        ("ALIGN", (1, 0), (-1, -1), "RIGHT"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#CCCCCC")),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]))
    story.append(terms_tbl)
    story.append(Spacer(1, 3))
    story.append(Paragraph(T["terms_note"], small))
    story.append(Spacer(1, 8))

    # Why buy
    story.append(Paragraph(T["why_h"], h2))
    for w in T["why"]:
        story.append(Paragraph("•&nbsp;&nbsp;" + w, body))
        story.append(Spacer(1, 2))
    story.append(Spacer(1, 6))

    # Photos row (keep heading + images together)
    imgs = []
    for p in (PHOTO_FRONT, PHOTO_BACK):
        im = RLImage(p)
        w, h = im.imageWidth, im.imageHeight
        tw = 42 * mm
        im.drawWidth = tw
        im.drawHeight = tw * h / w
        imgs.append(im)
    photo_tbl = Table([imgs], colWidths=[52 * mm, 52 * mm])
    photo_tbl.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP")]))
    story.append(KeepTogether([Paragraph(T["photos_h"], h2), photo_tbl]))
    story.append(Spacer(1, 8))

    # CTA box
    cta_tbl = Table([[Paragraph(f"<b>{T['cta']}</b>", body)]], colWidths=[164 * mm])
    cta_tbl.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), LIGHT_BG),
        ("BOX", (0, 0), (-1, -1), 0.6, BRAND_GREEN),
        ("TOPPADDING", (0, 0), (-1, -1), 8),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
        ("LEFTPADDING", (0, 0), (-1, -1), 10),
        ("RIGHTPADDING", (0, 0), (-1, -1), 10),
    ]))
    story.append(cta_tbl)
    story.append(Spacer(1, 6))

    story.append(HRFlowable(width="100%", thickness=0.6,
                            color=colors.HexColor("#CCCCCC"), spaceAfter=4))
    _dfmt = "%d.%m.%Y" if lang == "PL" else "%d %b %Y"
    story.append(Paragraph(
        f"{T['date_lbl']}: {datetime.now().strftime(_dfmt)}", small))
    story.append(Paragraph(T["foot"], small))

    doc.build(story)
    return fname


# ==========================================================================
# Email (.eml) builder
# ==========================================================================
def build_email(lang, pdf_path):
    bb = BEST_BEFORE.strftime("%d.%m.%Y" if lang == "PL" else "%d %b %Y")

    if lang == "EN":
        subject = (f"Offer: Alce Nero Organic Gluten-Free Spaghettoni 250 g "
                   f"— {eur(OFFER_UNIT)}/unit EXW, ~{DISCOUNT_PCT}% below retail")
        text = f"""Dear Buyer,

I am pleased to present an exclusive stock offer of a premium Italian organic
product for the Polish and CEE market:

  {PRODUCT_NAME}, {NET_WEIGHT}
  Brand: Alce Nero (Italy)  |  EAN: {EAN}  |  Code: {CODE}

WHY IT SELLS
  - Certified EU Organic AND gluten-free (suitable for coeliacs) - two of the
    fastest-growing grocery segments in one SKU.
  - Premium, well-recognised Italian brand for retail, delis and HORECA.
  - Product of Italy: corn flour 80% + rice flour 20% (organic).
  - Long shelf life: best before {bb}.

COMMERCIAL TERMS
  Our offer price:      {eur(OFFER_UNIT)} / unit  (net of VAT, {INCOTERMS})
  Reference retail:     {eur(RRP_UNIT)} / unit  -> about {DISCOUNT_PCT}% below retail
                        (~{MARKUP_X:.1f}x mark-up potential)
  Per pallet:           {UNITS_PER_PALLET:,} units = {eur(OFFER_PALLET)}
  Total available:      {UNITS_TOTAL:,} units (~{PALLETS:.0f} pallets) = {eur(OFFER_TOTAL,0)}
  Availability:         from single pallets up to the full lot
  Offer validity:       {OFFER_VALID_DAYS} days, subject to prior sale

Attached you will find:
  - the detailed offer sheet (PDF),
  - two product photos (front and back / label).

I would be glad to prepare a pro-forma invoice or reserve a specific volume for
you. Just let me know the quantity and delivery destination.

Best regards,
{SENDER_NAME}
{SENDER_TITLE_EN}
{SENDER_COMPANY}
{SENDER_EMAIL} | {SENDER_PHONE} | {SENDER_WEB}
"""
        html = f"""\
<html><body style="font-family:Arial,Helvetica,sans-serif;color:#3E3E3E;font-size:14px;line-height:1.5">
<p>Dear Buyer,</p>
<p>I am pleased to present an <b>exclusive stock offer</b> of a premium Italian
organic product for the Polish and CEE market:</p>
<p style="background:#F1F3E8;border-left:4px solid #5B7A1E;padding:10px 14px;margin:0 0 14px">
<b style="color:#5B7A1E;font-size:15px">{PRODUCT_NAME}, {NET_WEIGHT}</b><br/>
Brand: Alce Nero (Italy)&nbsp;&nbsp;|&nbsp;&nbsp;EAN: {EAN}&nbsp;&nbsp;|&nbsp;&nbsp;Code: {CODE}
</p>
<p style="margin:0 0 4px"><b style="color:#5B7A1E">Why it sells</b></p>
<ul style="margin:0 0 14px">
<li>Certified <b>EU Organic</b> <i>and</i> <b>gluten-free</b> (suitable for coeliacs) — two of the fastest-growing grocery segments in one SKU.</li>
<li>Premium, well-recognised Italian brand for retail, delis and HORECA.</li>
<li>Product of Italy: corn flour 80% + rice flour 20% (organic).</li>
<li>Long shelf life: best before <b>{bb}</b>.</li>
</ul>
<p style="margin:0 0 4px"><b style="color:#5B7A1E">Commercial terms</b></p>
<table cellpadding="6" cellspacing="0" style="border-collapse:collapse;font-size:13px;margin-bottom:12px">
<tr style="background:#5B7A1E;color:#fff"><th align="left">&nbsp;</th><th align="right">Per unit</th><th align="right">Per pallet ({UNITS_PER_PALLET:,})</th><th align="right">Full lot ({UNITS_TOTAL:,})</th></tr>
<tr><td>Reference retail</td><td align="right">{eur(RRP_UNIT)}</td><td align="right">{eur(RRP_UNIT*UNITS_PER_PALLET)}</td><td align="right">{eur(RRP_UNIT*UNITS_TOTAL,0)}</td></tr>
<tr style="background:#F1F3E8;color:#B23A1E;font-weight:bold"><td>OUR OFFER (net VAT)</td><td align="right">{eur(OFFER_UNIT)}</td><td align="right">{eur(OFFER_PALLET)}</td><td align="right">{eur(OFFER_TOTAL,0)}</td></tr>
</table>
<p style="font-size:12px;color:#777;margin:0 0 14px">
Incoterms: <b>{INCOTERMS}</b> · net of VAT · ~{DISCOUNT_PCT}% below retail (~{MARKUP_X:.1f}× mark-up) ·
availability from single pallets up to {UNITS_TOTAL:,} units (~{PALLETS:.0f} pallets) ·
offer valid {OFFER_VALID_DAYS} days, subject to prior sale.</p>
<p><b>Attached:</b> detailed offer sheet (PDF) and two product photos (front &amp; label).</p>
<p>I would be glad to prepare a pro-forma invoice or reserve a specific volume for you —
just let me know the quantity and delivery destination.</p>
<p>Best regards,<br/>
<b>{SENDER_NAME}</b><br/>{SENDER_TITLE_EN}<br/>{SENDER_COMPANY}<br/>
{SENDER_EMAIL} | {SENDER_PHONE} | {SENDER_WEB}</p>
</body></html>"""
    else:  # PL
        subject = (f"Oferta: Alce Nero Bezglutenowe Spaghettoni EKO 250 g — "
                   f"{eur(OFFER_UNIT)}/szt. EXW, ~{DISCOUNT_PCT}% poniżej detalu")
        text = f"""Szanowni Państwo,

z przyjemnością przedstawiam ekskluzywną ofertę magazynową renomowanego
włoskiego produktu ekologicznego dla rynku Polski i Europy Środkowo-Wschodniej:

  {PRODUCT_NAME}, {NET_WEIGHT}
  Marka: Alce Nero (Włochy)  |  EAN: {EAN}  |  Kod: {CODE}

DLACZEGO WARTO
  - Certyfikat ekologiczny UE ORAZ produkt bezglutenowy (dla celiaków) - dwa
    najszybciej rosnące segmenty spożywcze w jednym produkcie.
  - Rozpoznawalna włoska marka premium do handlu, delikatesów i HORECA.
  - Produkt Włoch: maka kukurydziana 80% + maka ryzowa 20% (ekologiczne).
  - Długi termin przydatności: najlepiej spożyć przed {bb}.

WARUNKI HANDLOWE
  Nasza cena ofertowa:  {eur(OFFER_UNIT)} / szt.  (netto, {INCOTERMS})
  Cena detaliczna ref.: {eur(RRP_UNIT)} / szt.  -> ok. {DISCOUNT_PCT}% poniżej detalu
                        (potencjał marzy ~{MARKUP_X:.1f}x)
  Za palete:            {UNITS_PER_PALLET:,} szt. = {eur(OFFER_PALLET)}
  Dostepnosc calkowita: {UNITS_TOTAL:,} szt. (~{PALLETS:.0f} palet) = {eur(OFFER_TOTAL,0)}
  Wolumen:              od pojedynczych palet do calej partii
  Waznosc oferty:       {OFFER_VALID_DAYS} dni, do wyczerpania zapasow

W zalaczeniu:
  - szczegolowy arkusz oferty (PDF),
  - dwa zdjecia produktu (front oraz etykieta).

Chetnie przygotuje fakture pro forma lub zarezerwuje wybrany wolumen. Prosze
o informacje dotyczaca ilosci oraz miejsca dostawy.

Z powazaniem,
{SENDER_NAME}
{SENDER_TITLE_PL}
{SENDER_COMPANY}
{SENDER_EMAIL} | {SENDER_PHONE} | {SENDER_WEB}
"""
        html = f"""\
<html><body style="font-family:Arial,Helvetica,sans-serif;color:#3E3E3E;font-size:14px;line-height:1.5">
<p>Szanowni Państwo,</p>
<p>z przyjemnością przedstawiam <b>ekskluzywną ofertę magazynową</b> renomowanego
włoskiego produktu ekologicznego dla rynku Polski i Europy Środkowo-Wschodniej:</p>
<p style="background:#F1F3E8;border-left:4px solid #5B7A1E;padding:10px 14px;margin:0 0 14px">
<b style="color:#5B7A1E;font-size:15px">{PRODUCT_NAME}, {NET_WEIGHT}</b><br/>
Marka: Alce Nero (Włochy)&nbsp;&nbsp;|&nbsp;&nbsp;EAN: {EAN}&nbsp;&nbsp;|&nbsp;&nbsp;Kod: {CODE}
</p>
<p style="margin:0 0 4px"><b style="color:#5B7A1E">Dlaczego warto</b></p>
<ul style="margin:0 0 14px">
<li>Certyfikat <b>ekologiczny UE</b> <i>oraz</i> produkt <b>bezglutenowy</b> (dla celiaków) — dwa najszybciej rosnące segmenty spożywcze w jednym produkcie.</li>
<li>Rozpoznawalna włoska marka premium do handlu, delikatesów i HORECA.</li>
<li>Produkt Włoch: mąka kukurydziana 80% + mąka ryżowa 20% (ekologiczne).</li>
<li>Długi termin przydatności: najlepiej spożyć przed <b>{bb}</b>.</li>
</ul>
<p style="margin:0 0 4px"><b style="color:#5B7A1E">Warunki handlowe</b></p>
<table cellpadding="6" cellspacing="0" style="border-collapse:collapse;font-size:13px;margin-bottom:12px">
<tr style="background:#5B7A1E;color:#fff"><th align="left">&nbsp;</th><th align="right">Za szt.</th><th align="right">Za paletę ({UNITS_PER_PALLET:,})</th><th align="right">Cała partia ({UNITS_TOTAL:,})</th></tr>
<tr><td>Cena detaliczna ref.</td><td align="right">{eur(RRP_UNIT)}</td><td align="right">{eur(RRP_UNIT*UNITS_PER_PALLET)}</td><td align="right">{eur(RRP_UNIT*UNITS_TOTAL,0)}</td></tr>
<tr style="background:#F1F3E8;color:#B23A1E;font-weight:bold"><td>NASZA OFERTA (netto)</td><td align="right">{eur(OFFER_UNIT)}</td><td align="right">{eur(OFFER_PALLET)}</td><td align="right">{eur(OFFER_TOTAL,0)}</td></tr>
</table>
<p style="font-size:12px;color:#777;margin:0 0 14px">
Incoterms: <b>{INCOTERMS}</b> · ceny netto · ok. {DISCOUNT_PCT}% poniżej detalu (potencjał marży ~{MARKUP_X:.1f}×) ·
dostępność od pojedynczych palet do {UNITS_TOTAL:,} szt. (~{PALLETS:.0f} palet) ·
oferta ważna {OFFER_VALID_DAYS} dni, do wyczerpania zapasów.</p>
<p><b>W załączeniu:</b> szczegółowy arkusz oferty (PDF) oraz dwa zdjęcia produktu (front i etykieta).</p>
<p>Chętnie przygotuję fakturę pro forma lub zarezerwuję wybrany wolumen — proszę
o informację dotyczącą ilości oraz miejsca dostawy.</p>
<p>Z poważaniem,<br/>
<b>{SENDER_NAME}</b><br/>{SENDER_TITLE_PL}<br/>{SENDER_COMPANY}<br/>
{SENDER_EMAIL} | {SENDER_PHONE} | {SENDER_WEB}</p>
</body></html>"""

    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = f"{SENDER_NAME} <{SENDER_EMAIL}>"
    msg["To"] = RECIPIENT_EMAIL
    msg["Date"] = formatdate(localtime=True)
    msg["Message-ID"] = make_msgid(domain="example.com")
    msg.set_content(text)
    msg.add_alternative(html, subtype="html")

    # Attach the two product photos + the PDF offer sheet
    attachments = [
        (PHOTO_FRONT, "image", "jpeg", "AlceNero_Spaghettoni_front.jpeg"),
        (PHOTO_BACK, "image", "jpeg", "AlceNero_Spaghettoni_label.jpeg"),
        (pdf_path, "application", "pdf", os.path.basename(pdf_path)),
    ]
    for path, maintype, subtype, fname in attachments:
        with open(path, "rb") as fh:
            msg.add_attachment(fh.read(), maintype=maintype,
                               subtype=subtype, filename=fname)

    out = os.path.join(OUT, f"Offer_AlceNero_Spaghettoni_{lang}.eml")
    with open(out, "wb") as fh:
        fh.write(bytes(msg))
    return out


if __name__ == "__main__":
    for lang in ("EN", "PL"):
        pdf = build_pdf(lang)
        eml = build_email(lang, pdf)
        print("built:", os.path.basename(pdf), "+", os.path.basename(eml))
    print("Done ->", OUT)
