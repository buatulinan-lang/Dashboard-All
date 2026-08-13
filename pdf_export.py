"""
Pembuat laporan analisa dalam bentuk PDF.

Dipakai oleh app.py: setiap dashboard punya tombol "Unduh Analisa (PDF)" yang
memanggil build_pdf(...) dan mengembalikan berkas dalam bentuk bytes.

Isi laporan: judul & periode, ringkasan angka (KPI), tabel perbandingan
bulan/tahun, lalu daftar temuan beserta tindak lanjutnya.
"""
import io
import re
from datetime import date
from pathlib import Path

from reportlab.lib import colors
from reportlab.lib.enums import TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import (BaseDocTemplate, Frame, KeepTogether, PageBreak,
                                PageTemplate, Paragraph, Spacer, Table, TableStyle)

ASSETS = Path(__file__).parent / "assets"

NAVY = colors.HexColor("#1F3864")
NAVY_L = colors.HexColor("#2E5394")
INK = colors.HexColor("#20242E")
MUTED = colors.HexColor("#6B7280")
LINE = colors.HexColor("#DDE2EC")
CARD = colors.HexColor("#F6F8FC")
GREEN = colors.HexColor("#16A34A")
RED = colors.HexColor("#C0392B")
AMBER = colors.HexColor("#D97706")

WARNA_JENIS = {
    'baik':      (colors.HexColor("#F2FBF5"), colors.HexColor("#B9E8C9"),
                  colors.HexColor("#0F5132"), "BAIK"),
    'perhatian': (colors.HexColor("#FFF8EC"), colors.HexColor("#F0D9A8"),
                  colors.HexColor("#7A5B18"), "PERHATIAN"),
    'aksi':      (colors.HexColor("#FDF3F2"), colors.HexColor("#EBCFCB"),
                  colors.HexColor("#7A2A24"), "TINDAKAN"),
    'info':      (colors.HexColor("#F7F9FD"), colors.HexColor("#E3E7F0"),
                  colors.HexColor("#1F3864"), "CATATAN"),
}

_ss = getSampleStyleSheet()


def _st(name, **kw):
    base = dict(name=name, fontName="Helvetica", fontSize=9.5, leading=13,
                textColor=INK, alignment=TA_LEFT)
    base.update(kw)
    return ParagraphStyle(**base)


S_JUDUL = _st("judul", fontName="Helvetica-Bold", fontSize=17, leading=21, textColor=NAVY)
S_SUB = _st("sub", fontSize=9.5, textColor=MUTED, leading=13)
S_H2 = _st("h2", fontName="Helvetica-Bold", fontSize=11.5, leading=15, textColor=NAVY,
           spaceBefore=10, spaceAfter=5)
S_BODY = _st("body", fontSize=9.5, leading=13.5)
S_KECIL = _st("kecil", fontSize=8, leading=11, textColor=MUTED)
S_KPI_L = _st("kpil", fontSize=7, leading=9, textColor=colors.white,
              fontName="Helvetica-Bold")
S_KPI_V = _st("kpiv", fontSize=13, leading=16, textColor=colors.white,
              fontName="Helvetica-Bold")
S_KPI_S = _st("kpis", fontSize=7, leading=9, textColor=colors.white)


def _warna(v, bawaan=None):
    """Terima warna berupa kode hex ('#1F3864') maupun objek warna ReportLab."""
    if v is None:
        return bawaan if bawaan is not None else NAVY
    if isinstance(v, str):
        try:
            return colors.HexColor(v if v.startswith('#') else f"#{v}")
        except Exception:
            return bawaan if bawaan is not None else NAVY
    return v


def _bersih(teks):
    """Ubah penanda HTML sederhana jadi tag yang dikenali ReportLab."""
    t = str(teks)
    t = t.replace("<b>", "|B|").replace("</b>", "|/B|")
    t = re.sub(r"<[^>]+>", "", t)
    t = t.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    t = t.replace("|B|", "<b>").replace("|/B|", "</b>")
    return t


def _kop(canvas, doc):
    """Kop & kaki halaman."""
    canvas.saveState()
    w, h = A4

    # garis kop
    canvas.setStrokeColor(LINE)
    canvas.setLineWidth(0.6)
    canvas.line(18 * mm, h - 22 * mm, w - 18 * mm, h - 22 * mm)

    # logo kanan atas
    x = w - 18 * mm
    for nm, lebar in (("logo_mflash.png", 13 * mm), ("logo_madinah.png", 15 * mm)):
        p = ASSETS / nm
        if p.exists():
            try:
                from reportlab.lib.utils import ImageReader
                img = ImageReader(str(p))
                iw, ih = img.getSize()
                tinggi = lebar * ih / iw
                x -= lebar
                canvas.drawImage(img, x, h - 20 * mm - tinggi / 2 + 2 * mm,
                                 width=lebar, height=tinggi, mask='auto')
                x -= 3 * mm
            except Exception:
                pass

    canvas.setFont("Helvetica-Bold", 8)
    canvas.setFillColor(NAVY)
    canvas.drawString(18 * mm, h - 19 * mm, "LAPORAN ANALISA DASHBOARD")

    canvas.setFont("Helvetica", 7.5)
    canvas.setFillColor(MUTED)
    canvas.drawString(18 * mm, 12 * mm,
                      f"Dibuat otomatis dari dashboard · {date.today():%d/%m/%Y}")
    canvas.drawRightString(w - 18 * mm, 12 * mm, f"Halaman {doc.page}")
    canvas.restoreState()


def _kpi_grid(kpis, lebar_total):
    """Baris kartu KPI berwarna."""
    if not kpis:
        return None
    per_baris = 3
    baris = [kpis[i:i + per_baris] for i in range(0, len(kpis), per_baris)]
    tabel_baris = []
    for grup in baris:
        sel = []
        for k in grup:
            isi = [[Paragraph(_bersih(k['label']).upper(), S_KPI_L)],
                   [Paragraph(_bersih(k['value']), S_KPI_V)]]
            if k.get('sub'):
                isi.append([Paragraph(_bersih(k['sub']), S_KPI_S)])
            t = Table(isi, colWidths=[lebar_total / per_baris - 4 * mm])
            t.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, -1), _warna(k.get('warna'), NAVY)),
                ('LEFTPADDING', (0, 0), (-1, -1), 6),
                ('RIGHTPADDING', (0, 0), (-1, -1), 6),
                ('TOPPADDING', (0, 0), (-1, 0), 6),
                ('BOTTOMPADDING', (0, -1), (-1, -1), 6),
                ('TOPPADDING', (0, 1), (-1, -1), 1),
                ('BOTTOMPADDING', (0, 0), (-1, -2), 1),
                ('ROUNDEDCORNERS', [4, 4, 4, 4]),
            ]))
            sel.append(t)
        while len(sel) < per_baris:
            sel.append("")
        tabel_baris.append(sel)

    luar = Table(tabel_baris, colWidths=[lebar_total / per_baris] * per_baris)
    luar.setStyle(TableStyle([
        ('LEFTPADDING', (0, 0), (-1, -1), 0),
        ('RIGHTPADDING', (0, 0), (-1, -1), 0),
        ('TOPPADDING', (0, 0), (-1, -1), 3),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
    ]))
    return luar


def _tabel_banding(judul, baris, nama_baru, nama_lama, catatan, lebar):
    """Tabel perbandingan: metrik, nilai baru, nilai lama, selisih."""
    data = [[Paragraph("<b>Metrik</b>", S_KECIL),
             Paragraph(f"<b>{_bersih(nama_baru)}</b>", S_KECIL),
             Paragraph(f"<b>{_bersih(nama_lama)}</b>", S_KECIL),
             Paragraph("<b>Selisih</b>", S_KECIL)]]
    gaya = [
        ('GRID', (0, 0), (-1, -1), 0.4, LINE),
        ('BACKGROUND', (0, 0), (-1, 0), CARD),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('LEFTPADDING', (0, 0), (-1, -1), 5),
        ('RIGHTPADDING', (0, 0), (-1, -1), 5),
        ('TOPPADDING', (0, 0), (-1, -1), 3.5),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 3.5),
        ('ALIGN', (3, 0), (3, -1), 'RIGHT'),
    ]
    for i, (nama, v_baru, v_lama, delta, warna) in enumerate(baris, start=1):
        data.append([
            Paragraph(_bersih(nama), S_BODY),
            Paragraph(f"<b>{_bersih(v_baru)}</b>", S_BODY),
            Paragraph(_bersih(v_lama),
                      _st(f"m{i}", fontSize=9.5, textColor=MUTED)),
            Paragraph(f"<b>{_bersih(delta)}</b>",
                      _st(f"d{i}", fontSize=9.5, textColor=_warna(warna, MUTED),
                          alignment=2)),
        ])
    t = Table(data, colWidths=[lebar * 0.34, lebar * 0.22, lebar * 0.22, lebar * 0.22])
    t.setStyle(TableStyle(gaya))
    return KeepTogether([Paragraph(judul, S_H2), t,
                         Spacer(1, 2), Paragraph(_bersih(catatan), S_KECIL),
                         Spacer(1, 6)])


def _kotak_temuan(jenis, judul, isi, lebar):
    bg, br, tx, tag = WARNA_JENIS.get(jenis, WARNA_JENIS['info'])
    kepala = Paragraph(
        f"<b>[{tag}] {_bersih(judul)}</b>",
        _st("kt", fontName="Helvetica-Bold", fontSize=9.5, leading=13, textColor=tx))
    badan = Paragraph(_bersih(isi),
                      _st("ki", fontSize=9, leading=12.5, textColor=tx))
    t = Table([[kepala], [badan]], colWidths=[lebar])
    t.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), bg),
        ('BOX', (0, 0), (-1, -1), 0.5, br),
        ('LEFTPADDING', (0, 0), (-1, -1), 7),
        ('RIGHTPADDING', (0, 0), (-1, -1), 7),
        ('TOPPADDING', (0, 0), (-1, 0), 6),
        ('BOTTOMPADDING', (0, -1), (-1, -1), 6),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 2),
        ('TOPPADDING', (0, 1), (-1, -1), 0),
    ]))
    return KeepTogether([t, Spacer(1, 5)])


def build_pdf(*, judul, periode, cabang, kpis=None, banding=None, temuan=None,
              metodologi="", penyusun="", ringkasan=""):
    """Susun laporan analisa jadi berkas PDF (bytes).

    judul      : nama dashboard, mis. "Dashboard Pending"
    periode    : keterangan periode data
    cabang     : keterangan cabang
    kpis       : list dict {label, value, sub, warna}
    banding    : list dict {judul, baris, nama_baru, nama_lama, catatan}
                 baris = list (metrik, nilai_baru, nilai_lama, selisih, warna)
    temuan     : list (jenis, judul, isi)
    """
    buf = io.BytesIO()
    doc = BaseDocTemplate(
        buf, pagesize=A4,
        leftMargin=18 * mm, rightMargin=18 * mm,
        topMargin=26 * mm, bottomMargin=18 * mm,
        title=f"Analisa — {judul}", author=penyusun or "Dashboard")
    lebar = doc.width
    doc.addPageTemplates([PageTemplate(
        id='utama',
        frames=[Frame(doc.leftMargin, doc.bottomMargin, doc.width, doc.height, id='f')],
        onPage=_kop)])

    cerita = []
    cerita.append(Paragraph(_bersih(judul), S_JUDUL))
    ket = f"{_bersih(periode)} &nbsp;·&nbsp; {_bersih(cabang)}"
    if penyusun:
        ket += f" &nbsp;·&nbsp; disiapkan oleh {_bersih(penyusun)}"
    cerita.append(Paragraph(ket, S_SUB))
    cerita.append(Spacer(1, 8))

    if ringkasan:
        cerita.append(Paragraph(_bersih(ringkasan), S_BODY))
        cerita.append(Spacer(1, 6))

    if kpis:
        cerita.append(Paragraph("Ringkasan Angka", S_H2))
        g = _kpi_grid(kpis, lebar)
        if g:
            cerita.append(g)
        cerita.append(Spacer(1, 6))

    for b in (banding or []):
        if not b.get('baris'):
            # tetap tampilkan keterangannya supaya pembaca tahu mengapa kosong
            ket = b.get('catatan') or "Data pembanding tidak tersedia."
            kotak = Table([[Paragraph(_bersih(ket),
                                      _st("kk", fontSize=9, leading=12.5,
                                          textColor=colors.HexColor("#7A5B18")))]],
                          colWidths=[lebar])
            kotak.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor("#FFF8EC")),
                ('BOX', (0, 0), (-1, -1), 0.5, colors.HexColor("#F0D9A8")),
                ('LEFTPADDING', (0, 0), (-1, -1), 7),
                ('RIGHTPADDING', (0, 0), (-1, -1), 7),
                ('TOPPADDING', (0, 0), (-1, -1), 6),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
            ]))
            cerita.append(KeepTogether([
                Paragraph(b.get('judul', 'Perbandingan'), S_H2), kotak, Spacer(1, 6)]))
            continue
        cerita.append(_tabel_banding(
            b.get('judul', 'Perbandingan'), b['baris'],
            b.get('nama_baru', 'Periode ini'), b.get('nama_lama', 'Pembanding'),
            b.get('catatan', ''), lebar))

    if temuan:
        cerita.append(Paragraph("Analisa &amp; Tindak Lanjut", S_H2))
        for jenis, jd, isi in temuan:
            cerita.append(_kotak_temuan(jenis, jd, isi, lebar))

    if metodologi:
        cerita.append(Spacer(1, 4))
        cerita.append(Paragraph("Catatan Metodologi", S_H2))
        cerita.append(Paragraph(_bersih(metodologi), S_KECIL))

    doc.build(cerita)
    buf.seek(0)
    return buf.getvalue()
