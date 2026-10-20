"""
Pembuat PDF Laporan Penjualan YN Hijab Fashion (ReportLab).

Modul ini HANYA membuat PDF dari data yang diberikan.
Tidak ada akses database di sini, jadi aman dan mudah diuji.
"""

from datetime import datetime, date
from io import BytesIO
from xml.sax.saxutils import escape

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas
from reportlab.platypus import (
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)

# ---------------------------------------------------------
# Warna tema (sama dengan laporan.css)
# ---------------------------------------------------------
PINK = colors.HexColor("#D58A94")
PINK_DARK = colors.HexColor("#B5606C")
PINK_LIGHT = colors.HexColor("#FAEBED")
PINK_SOFT = colors.HexColor("#F5D9DD")
TEXT = colors.HexColor("#4A3338")
TEXT_SOFT = colors.HexColor("#85696F")
BORDER = colors.HexColor("#F0E2E4")

NAMA_TOKO = "YN Hijab Fashion"

BULAN = [
    "Jan", "Feb", "Mar", "Apr", "Mei", "Jun",
    "Jul", "Agu", "Sep", "Okt", "Nov", "Des",
]


# ---------------------------------------------------------
# Format
# ---------------------------------------------------------
def format_rupiah(nilai):
    """1250000 -> 'Rp 1.250.000'."""
    try:
        angka = int(round(float(nilai or 0)))
    except (TypeError, ValueError):
        angka = 0
    return "Rp " + f"{angka:,}".replace(",", ".")


def format_tanggal(nilai, dengan_jam=True):
    """datetime/date/string -> '19 Okt 2026 14:30'."""
    if nilai is None or nilai == "":
        return "-"
    if isinstance(nilai, str):
        for pola in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%Y-%m-%d"):
            try:
                nilai = datetime.strptime(nilai, pola)
                break
            except ValueError:
                continue
        else:
            return nilai
    teks = f"{nilai.day:02d} {BULAN[nilai.month - 1]} {nilai.year}"
    if dengan_jam and isinstance(nilai, datetime):
        teks += f" {nilai.hour:02d}:{nilai.minute:02d}"
    return teks


def teks_periode(dari, sampai):
    """dari/sampai berupa date atau None."""
    if dari and sampai:
        return f"{format_tanggal(dari, False)} s/d {format_tanggal(sampai, False)}"
    if dari:
        return f"Mulai {format_tanggal(dari, False)}"
    if sampai:
        return f"Sampai {format_tanggal(sampai, False)}"
    return "Semua tanggal"


def _aman(nilai):
    """Escape teks untuk Paragraph; kosong -> '-'."""
    if nilai is None or str(nilai).strip() == "":
        return "-"
    return escape(str(nilai))


# ---------------------------------------------------------
# Canvas dengan "Halaman X dari Y"
# ---------------------------------------------------------
class _NumberedCanvas(canvas.Canvas):
    def __init__(self, *args, waktu_cetak="", **kwargs):
        super().__init__(*args, **kwargs)
        self._saved_pages = []
        self._waktu_cetak = waktu_cetak

    def showPage(self):
        self._saved_pages.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        total = len(self._saved_pages)
        for state in self._saved_pages:
            self.__dict__.update(state)
            self._footer(total)
            super().showPage()
        super().save()

    def _footer(self, total):
        lebar, _ = landscape(A4)
        self.setStrokeColor(BORDER)
        self.setLineWidth(0.6)
        self.line(15 * mm, 12 * mm, lebar - 15 * mm, 12 * mm)
        self.setFont("Helvetica", 8)
        self.setFillColor(TEXT_SOFT)
        self.drawString(
            15 * mm, 7.5 * mm,
            f"{NAMA_TOKO}  |  Dicetak: {self._waktu_cetak}",
        )
        self.drawRightString(
            lebar - 15 * mm, 7.5 * mm,
            f"Halaman {self._pageNumber} dari {total}",
        )


# ---------------------------------------------------------
# Fungsi utama
# ---------------------------------------------------------
def buat_pdf_laporan_penjualan(
    data,
    dari=None,
    sampai=None,
    status=None,
    pencarian=None,
    waktu_cetak=None,
):
    """
    data      : list dict dengan kunci id_pesanan, no_pesanan, tanggal,
                nama_customor, no_hp, total_harga, STATUS,
                metode_pengambilan.
    Mengembalikan BytesIO berisi PDF.
    """
    waktu_cetak = waktu_cetak or datetime.now()
    waktu_teks = format_tanggal(waktu_cetak) + " WIB"

    buffer = BytesIO()
    halaman = landscape(A4)
    margin = 15 * mm
    lebar_isi = halaman[0] - 2 * margin

    doc = SimpleDocTemplate(
        buffer,
        pagesize=halaman,
        leftMargin=margin,
        rightMargin=margin,
        topMargin=14 * mm,
        bottomMargin=18 * mm,
        title="Laporan Penjualan - " + NAMA_TOKO,
        author=NAMA_TOKO,
    )

    # ---- gaya teks ----
    s_toko = ParagraphStyle(
        "toko", fontName="Helvetica-Bold", fontSize=11,
        textColor=PINK_DARK, leading=14,
    )
    s_judul = ParagraphStyle(
        "judul", fontName="Helvetica-Bold", fontSize=18,
        textColor=TEXT, leading=22,
    )
    s_meta = ParagraphStyle(
        "meta", fontName="Helvetica", fontSize=9,
        textColor=TEXT_SOFT, leading=13,
    )
    s_meta_kanan = ParagraphStyle("metak", parent=s_meta, alignment=TA_RIGHT)
    s_head = ParagraphStyle(
        "head", fontName="Helvetica-Bold", fontSize=9,
        textColor=colors.white, leading=11,
    )
    s_head_kanan = ParagraphStyle("headk", parent=s_head, alignment=TA_RIGHT)
    s_cell = ParagraphStyle(
        "cell", fontName="Helvetica", fontSize=8.5,
        textColor=TEXT, leading=11, alignment=TA_LEFT,
    )
    s_cell_kanan = ParagraphStyle("cellk", parent=s_cell, alignment=TA_RIGHT)
    s_kosong = ParagraphStyle(
        "kosong", fontName="Helvetica-Oblique", fontSize=10,
        textColor=TEXT_SOFT, leading=14, alignment=TA_CENTER,
    )
    s_kartu_label = ParagraphStyle(
        "kl", fontName="Helvetica", fontSize=8.5,
        textColor=TEXT_SOFT, leading=11,
    )
    s_kartu_nilai = ParagraphStyle(
        "kn", fontName="Helvetica-Bold", fontSize=14,
        textColor=PINK_DARK, leading=18,
    )
    s_catatan = ParagraphStyle(
        "catatan", fontName="Helvetica-Oblique", fontSize=8,
        textColor=TEXT_SOFT, leading=11,
    )

    story = []

    # ---- kepala laporan ----
    filter_teks = []
    filter_teks.append(f"<b>Status:</b> {_aman(status) if status else 'Semua status'}")
    if pencarian:
        filter_teks.append(f"<b>Pencarian:</b> {_aman(pencarian)}")

    kiri = [
        Paragraph(NAMA_TOKO.upper(), s_toko),
        Paragraph("Laporan Penjualan", s_judul),
    ]
    kanan = [
        Paragraph(f"<b>Periode:</b> {teks_periode(dari, sampai)}", s_meta_kanan),
        Paragraph("&nbsp;&nbsp;|&nbsp;&nbsp;".join(filter_teks), s_meta_kanan),
        Paragraph(f"<b>Dibuat:</b> {waktu_teks}", s_meta_kanan),
    ]
    kepala = Table([[kiri, kanan]], colWidths=[lebar_isi * 0.45, lebar_isi * 0.55])
    kepala.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "BOTTOM"),
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 0),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ("LINEBELOW", (0, 0), (-1, 0), 1.6, PINK),
    ]))
    story.append(kepala)
    story.append(Spacer(1, 8))

    # ---- ringkasan (dihitung dari baris yang diekspor) ----
    total_pesanan = len(data)
    total_penjualan = sum(
        float(r.get("total_harga") or 0)
        for r in data
        if str(r.get("STATUS") or "").lower() != "dibatalkan"
    )

    def kartu(label, nilai):
        return [Paragraph(label, s_kartu_label), Paragraph(nilai, s_kartu_nilai)]

    ringkasan = Table(
        [[kartu("Total Pesanan", str(total_pesanan)),
          kartu("Total Penjualan", format_rupiah(total_penjualan))]],
        colWidths=[lebar_isi * 0.25, lebar_isi * 0.30],
        hAlign="LEFT",
    )
    ringkasan.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), PINK_LIGHT),
        ("BOX", (0, 0), (0, 0), 0.6, PINK_SOFT),
        ("BOX", (1, 0), (1, 0), 0.6, PINK_SOFT),
        ("LEFTPADDING", (0, 0), (-1, -1), 10),
        ("TOPPADDING", (0, 0), (-1, -1), 7),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
    ]))
    story.append(ringkasan)
    story.append(Spacer(1, 3))
    story.append(Paragraph(
        "Total penjualan tidak termasuk pesanan berstatus dibatalkan.",
        s_catatan,
    ))
    story.append(Spacer(1, 8))

    # ---- tabel ----
    header = [
        Paragraph("No Pesanan", s_head),
        Paragraph("Tanggal", s_head),
        Paragraph("Pelanggan", s_head),
        Paragraph("No HP", s_head),
        Paragraph("Total", s_head_kanan),
        Paragraph("Status", s_head),
        Paragraph("Pengambilan", s_head),
    ]
    proporsi = [0.12, 0.13, 0.22, 0.13, 0.13, 0.12, 0.15]
    lebar_kolom = [lebar_isi * p for p in proporsi]

    if not data:
        story.append(Spacer(1, 18))
        kosong = Table(
            [[Paragraph(
                "Tidak ada data penjualan untuk filter / periode yang dipilih.",
                s_kosong,
            )]],
            colWidths=[lebar_isi],
        )
        kosong.setStyle(TableStyle([
            ("BOX", (0, 0), (-1, -1), 0.6, BORDER),
            ("BACKGROUND", (0, 0), (-1, -1), PINK_LIGHT),
            ("TOPPADDING", (0, 0), (-1, -1), 26),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 26),
        ]))
        story.append(kosong)
    else:
        baris = [header]
        for r in data:
            no = r.get("no_pesanan") or f"#{r.get('id_pesanan')}"
            baris.append([
                Paragraph(_aman(no), s_cell),
                Paragraph(format_tanggal(r.get("tanggal")), s_cell),
                Paragraph(_aman(r.get("nama_customor")), s_cell),
                Paragraph(_aman(r.get("no_hp")), s_cell),
                Paragraph(format_rupiah(r.get("total_harga")), s_cell_kanan),
                Paragraph(_aman(r.get("STATUS")), s_cell),
                Paragraph(_aman(r.get("metode_pengambilan")), s_cell),
            ])

        tabel = Table(baris, colWidths=lebar_kolom, repeatRows=1)
        tabel.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), PINK),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("TOPPADDING", (0, 0), (-1, -1), 5),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
            ("LEFTPADDING", (0, 0), (-1, -1), 6),
            ("RIGHTPADDING", (0, 0), (-1, -1), 6),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, PINK_LIGHT]),
            ("LINEBELOW", (0, 1), (-1, -1), 0.4, BORDER),
            ("BOX", (0, 0), (-1, -1), 0.6, PINK_SOFT),
        ]))
        story.append(tabel)

    doc.build(
        story,
        canvasmaker=lambda *a, **k: _NumberedCanvas(*a, waktu_cetak=waktu_teks, **k),
    )
    buffer.seek(0)
    return buffer