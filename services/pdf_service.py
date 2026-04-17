# -*- coding: utf-8 -*-
"""PDF dışa aktarma: yatay A4, özet tablo, çizelge tablosu, Türkçe font."""

from datetime import datetime
from pathlib import Path
from typing import Any, Dict, Optional, Tuple, Union

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer

from constants import FONT_PATH, GUNLER, SAATLER


def _parse_entry(entry: Any) -> Tuple[str, bool]:
    if isinstance(entry, dict):
        return entry.get("text", ""), bool(entry.get("done", False))
    if entry is None:
        return "", False
    return str(entry), False


def _setup_styles(font_path: Optional[str] = None):
    """PDF font ve stillerini hazırla. font_path verilmezse constants.FONT_PATH kullanılır."""
    path = font_path or FONT_PATH
    try:
        pdfmetrics.registerFont(TTFont("AnaFont", path))
        styles = getSampleStyleSheet()
        normal_style = styles["Normal"]
        normal_style.fontName = "AnaFont"
    except Exception:
        styles = getSampleStyleSheet()
        normal_style = styles["Normal"]
    return styles, normal_style


class PdfService:
    """Haftalık çizelge PDF'i oluşturur."""

    def __init__(self, font_path: Optional[str] = None):
        self._styles, self._normal_style = _setup_styles(font_path)

    def build_pdf(
        self,
        path: Union[str, Path],
        program: Dict[str, Dict[str, Any]],
        ogrenci_adi: Optional[str] = None,
        hafta_adi: Optional[str] = None,
    ) -> None:
        """
        program[gun][saat] = {text, done}.
        PDF: başlık, tarih, özet (toplam/tamamlanan/oran, en yoğun/en boş gün), günlük özet, çizelge tablosu.
        """
        path = Path(path)
        doc = SimpleDocTemplate(
            str(path),
            pagesize=landscape(A4),
            leftMargin=20,
            rightMargin=20,
            topMargin=20,
            bottomMargin=20,
        )
        story = []
        normal = self._normal_style
        styles = self._styles

        # İstatistik
        gunluk_toplam = {}
        gunluk_yapilan = {}
        for gun in GUNLER:
            gun_prog = program.get(gun, {})
            gunluk_toplam[gun] = len(gun_prog)
            yapilan = sum(1 for _, e in gun_prog.items() if _parse_entry(e)[1])
            gunluk_yapilan[gun] = yapilan
        haftalik_toplam = sum(gunluk_toplam.values())
        haftalik_yapilan = sum(gunluk_yapilan.values())
        oran = int(round(100 * haftalik_yapilan / haftalik_toplam)) if haftalik_toplam else 0
        en_yogun = max(gunluk_toplam, key=lambda g: gunluk_toplam[g]) if gunluk_toplam else None
        en_bos = min(gunluk_toplam, key=lambda g: gunluk_toplam[g]) if gunluk_toplam else None

        # Başlık
        baslik_text = "Haftalık Çalışma Programı"
        if ogrenci_adi:
            baslik_text += f" - {ogrenci_adi}"
        if hafta_adi:
            baslik_text += f" ({hafta_adi})"
        baslik_style = styles["Heading2"]
        baslik_style.fontName = getattr(normal, "fontName", "Helvetica")
        story.append(Paragraph(baslik_text, baslik_style))
        story.append(Paragraph(datetime.now().strftime("%d.%m.%Y"), normal))
        story.append(Spacer(1, 8))

        # Özet tablo
        summary_data = [
            [Paragraph("Haftalık Planlanan", normal), Paragraph(f"{haftalik_toplam} saat", normal)],
            [Paragraph("Haftalık Tamamlanan", normal), Paragraph(f"{haftalik_yapilan} saat", normal)],
            [Paragraph("Tamamlama Oranı", normal), Paragraph(f"%{oran}", normal)],
        ]
        if en_yogun is not None:
            summary_data.append([
                Paragraph("En Yoğun Gün", normal),
                Paragraph(f"{en_yogun} ({gunluk_toplam[en_yogun]} saat)", normal),
            ])
            summary_data.append([
                Paragraph("En Boş Gün", normal),
                Paragraph(f"{en_bos} ({gunluk_toplam[en_bos]} saat)", normal),
            ])
        summary_table = Table(summary_data, colWidths=[150, 200])
        summary_table.setStyle(TableStyle([
            ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
            ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("LEFTPADDING", (0, 0), (-1, -1), 4),
            ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ]))
        story.append(summary_table)
        story.append(Spacer(1, 12))

        # Günlük özet metni
        gunluk_par = []
        for gun in GUNLER:
            plan = gunluk_toplam.get(gun, 0)
            yapilan = gunluk_yapilan.get(gun, 0)
            oran_gun = int(round(100 * yapilan / plan)) if plan else 0
            gunluk_par.append(f"{gun}: {yapilan}/{plan} saat (%{oran_gun})")
        story.append(Paragraph("<br/>".join(gunluk_par), normal))
        story.append(Spacer(1, 10))

        # Ana tablo: Saat + günler
        data = []
        data.append([Paragraph(h, normal) for h in ["Saat"] + GUNLER])
        for saat in SAATLER:
            satir = [Paragraph(saat, normal)]
            for gun in GUNLER:
                gun_prog = program.get(gun, {})
                metin = ""
                if saat in gun_prog:
                    text, done = _parse_entry(gun_prog[saat])
                    metin = ("✓ " if done else "") + text
                metin_pdf = metin.replace(" | ", "<br/>• ")
                satir.append(Paragraph(metin_pdf or " ", normal))
            data.append(satir)
        table = Table(data, repeatRows=1)
        table.setStyle(TableStyle([
            ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
            ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
            ("ALIGN", (0, 0), (0, -1), "CENTER"),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("FONTSIZE", (0, 0), (-1, -1), 9),
            ("LEFTPADDING", (0, 0), (-1, -1), 4),
            ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ]))
        story.append(table)
        doc.build(story)
