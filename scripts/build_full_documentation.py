import os
import sys
import glob
import json
import re
from datetime import datetime

import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn

from docx_helpers import (
    NAVY, ROYAL, SLATE_DARK, SLATE_GRAY, LIGHT_BG, BORDER_COLOR, WHITE,
    CALLOUT_BG, SUCCESS_BG, set_cell_background, set_cell_margins,
    set_table_borders, add_callout, add_code_block, format_styled_table,
    add_table_row
)

def build_documentation():
    print("Memulai pembuatan dokumen spesifikasi & kode MabelHub...")
    doc = Document()
    
    # Page setup - Margins: 2.2 cm (approx 0.86 inches)
    sections = doc.sections
    for section in sections:
        section.top_margin = Inches(0.85)
        section.bottom_margin = Inches(0.85)
        section.left_margin = Inches(0.85)
        section.right_margin = Inches(0.85)
        
        # Header & Footer setup
        header = section.header
        p_hdr = header.paragraphs[0]
        p_hdr.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        r_hdr = p_hdr.add_run("MabelHub | Dokumentasi Teknis & Kode Sumber Sistem")
        r_hdr.font.name = "Calibri"
        r_hdr.font.size = Pt(8.5)
        r_hdr.font.color.rgb = RGBColor(0x94, 0xA3, 0xB8)
        
        footer = section.footer
        p_ftr = footer.paragraphs[0]
        p_ftr.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r_ftr = p_ftr.add_run("Dokumen Teknis Tugas Akhir (TA) - Platform MabelHub Sales & E-Procurement")
        r_ftr.font.name = "Calibri"
        r_ftr.font.size = Pt(8)
        r_ftr.font.color.rgb = RGBColor(0x94, 0xA3, 0xB8)

    # Styles
    styles = doc.styles
    normal_style = styles['Normal']
    normal_style.font.name = 'Calibri'
    normal_style.font.size = Pt(10)
    normal_style.font.color.rgb = RGBColor(0x1E, 0x29, 0x3B)
    normal_style.paragraph_format.line_spacing = 1.15
    normal_style.paragraph_format.space_after = Pt(4)

    # ==========================================
    # COVER / TITLE PAGE
    # ==========================================
    p_cov_space = doc.add_paragraph()
    p_cov_space.paragraph_format.space_before = Pt(36)
    
    p_badge = doc.add_paragraph()
    p_badge.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_badge = p_badge.add_run("DOKUMEN SPESIFIKASI TEKNIK & ARSITEKTUR KODE SUMBER")
    r_badge.font.name = "Calibri"
    r_badge.font.size = Pt(11)
    r_badge.font.bold = True
    r_badge.font.color.rgb = RGBColor(0x25, 0x63, 0xEB)
    
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_before = Pt(12)
    p_title.paragraph_format.space_after = Pt(12)
    r_title = p_title.add_run("MABELHUB\nSALES & E-PROCUREMENT SYSTEM")
    r_title.font.name = "Calibri"
    r_title.font.size = Pt(26)
    r_title.font.bold = True
    r_title.font.color.rgb = RGBColor(0x1E, 0x3A, 0x8A)
    
    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sub.paragraph_format.space_after = Pt(28)
    r_sub = p_sub.add_run("Platform Manajemen Terpadu Pipeline Penjualan B2G, Database Instansi Pemerintah, Pelacakan Kunjungan Sales, & Administrasi Pengadaan E-Procurement")
    r_sub.font.name = "Calibri"
    r_sub.font.size = Pt(12)
    r_sub.font.color.rgb = RGBColor(0x47, 0x55, 0x69)
    
    # Horizontal divider table
    t_div = doc.add_table(rows=1, cols=1)
    t_div.alignment = WD_TABLE_ALIGNMENT.CENTER
    c_div = t_div.cell(0, 0)
    c_div.width = Inches(6.5)
    set_cell_background(c_div, ROYAL)
    c_div.paragraphs[0].paragraph_format.space_before = Pt(1)
    c_div.paragraphs[0].paragraph_format.space_after = Pt(1)
    
    # Metadata Table
    doc.add_paragraph().paragraph_format.space_before = Pt(24)
    t_meta = doc.add_table(rows=6, cols=2)
    t_meta.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(t_meta, color="E2E8F0", sz="4")
    
    meta_info = [
        ("Nama Proyek", "MabelHub (Mabel Sales & E-Procurement Hub)"),
        ("Tipe Aplikasi", "Fullstack Web Application (Next.js App Router Architecture)"),
        ("Teknologi Utama", "Next.js 16, React 19, TypeScript, MongoDB, Tailwind CSS, Leaflet GIS"),
        ("Sistem Keamanan", "JWT (JSON Web Token), HttpOnly Cookies, Role-Based Access Control (RBAC)"),
        ("Cakupan Dokumentasi", "67 API Endpoints, 31 Halaman UI, 16 Komponen UI, 22 Koleksi MongoDB, BDD/TDD"),
        ("Tahun / Rilis", "2026 - Dokumentasi Proyek Tugas Akhir (TA)"),
    ]
    for row_idx, (k, v) in enumerate(meta_info):
        c0 = t_meta.rows[row_idx].cells[0]
        c1 = t_meta.rows[row_idx].cells[1]
        c0.width = Inches(2.2)
        c1.width = Inches(4.3)
        c0.text = k
        c1.text = v
        set_cell_background(c0, "F1F5F9")
        set_cell_background(c1, WHITE)
        set_cell_margins(c0, top=80, bottom=80, left=120, right=120)
        set_cell_margins(c1, top=80, bottom=80, left=120, right=120)
        p0 = c0.paragraphs[0]
        p1 = c1.paragraphs[0]
        p0.runs[0].font.bold = True
        p0.runs[0].font.size = Pt(9.5)
        p0.runs[0].font.color.rgb = RGBColor(0x1E, 0x3A, 0x8A)
        p1.runs[0].font.size = Pt(9.5)
        p1.runs[0].font.color.rgb = RGBColor(0x33, 0x41, 0x55)

    doc.add_page_break()

    # ==========================================
    # DAFTAR ISI RINGKAS
    # ==========================================
    h_toc = doc.add_heading(level=1)
    r_htoc = h_toc.add_run("DAFTAR ISI DOKUMENTASI")
    r_htoc.font.name = "Calibri"
    r_htoc.font.size = Pt(18)
    r_htoc.font.bold = True
    r_htoc.font.color.rgb = RGBColor(0x1E, 0x3A, 0x8A)
    
    toc_items = [
        ("BAB 1: PENDAHULUAN & ARSITEKTUR SISTEM", [
            "1.1 Latar Belakang & Domain Masalah Bisnis B2G",
            "1.2 Tujuan dan Manfaat Sistem MabelHub",
            "1.3 Arsitektur Tingkat Tinggi (High-Level Architecture)",
            "1.4 Tech Stack & Analisis Pustaka Dependensi (package.json)",
            "1.5 Struktur Direktori Proyek & Pola App Router Next.js"
        ]),
        ("BAB 2: ARSITEKTUR KEAMANAN, OTENTIKASI, & BASIS DATA", [
            "2.1 Mekanisme Otentikasi Berbasis JWT & Cookie Session",
            "2.2 Role-Based Access Control (RBAC) & Proteksi Middleware",
            "2.3 Desain & Skema 22 Koleksi Basis Data MongoDB",
            "2.4 Strategi Indeks dan Optimasi Kueri Agregasi"
        ]),
        ("BAB 3: DOKUMENTASI LENGKAP BACKEND API (67 ENDPOINTS)", [
            "3.1 Modul Otentikasi & Sesi Pengguna (/api/auth/...)",
            "3.2 Modul Pengguna & Manajemen Tim (/api/users, /api/teams/...)",
            "3.3 Modul Perusahaan & Instansi Pemerintah (/api/companies, /api/instansi)",
            "3.4 Modul Pengajuan Perusahaan & Approval (/api/company-requests/...)",
            "3.5 Modul Kunjungan Sales & Geospasial Visit (/api/visits, /api/plans)",
            "3.6 Modul E-Procurement & Penawaran SPH (/api/e-procurement/...)",
            "3.7 Modul Pipeline Telemarketing & Database Input (/api/input-database, /api/tracking-database)",
            "3.8 Modul Tracking Broadcast & Telemarketing Call (/api/tracking-broadcast, /api/tracking-call)",
            "3.9 Modul Validasi Sales & Tindak Lanjut (/api/validasi-sales, /api/tindak-lanjut)",
            "3.10 Modul Produk Hub & Manajemen Dokumen GridFS (/api/produk/...)",
            "3.11 Modul Digital Marketing, Kontrak, & Parameter (/api/marketing, /api/contracts, /api/parameters)",
            "3.12 Modul Notifikasi Terintegrasi (/api/notifications/...)"
        ]),
        ("BAB 4: DOKUMENTASI HALAMAN FRONTEND (31 PAGES)", [
            "4.1 Modul Dashboard Utama & Keuangan (/dashboard, /dashboard-request, /finance)",
            "4.2 Modul Master Instansi & Satuan Kerja (/instansi, /tambah-instansi, /tracking-satker)",
            "4.3 Modul Pipeline Telemarketing (/input-database, /tracking-database, /tracking-broadcast, /tracking-call)",
            "4.4 Modul Operasional Sales Lapangan (/validasi-sales, /tindak-lanjut-sales, /tracking-b2g)",
            "4.5 Modul Kunjungan Sales (/plan-activity, /rekapitulasi-visit)",
            "4.6 Modul Pengadaan E-Procurement (/e-procurement, /e-procurement-response, /rekapitulasi-Eproc)",
            "4.7 Modul Pelaporan & Progres (/report-progres, /sales-report-system, /rekapitulasi-response)",
            "4.8 Modul Katalog Produk Hub (/produk, /produk/[kategori])",
            "4.9 Modul Administrasi Tim & Parameter (/teams, /teams/[teamId], /add-user, /parameters, /kontrak)"
        ]),
        ("BAB 5: DOKUMENTASI KOMPONEN UI, MODAL, DAN HOOKS", [
            "5.1 Komponen Tata Letak & Navigasi (AppLayoutWrapper, Sidebar)",
            "5.2 Komponen Modal Dialog & Interaktif (SalesMap, EditVisit, EditInstansi, dsb.)",
            "5.3 Komponen UI Reusable (TableCard, DatePicker, SearchableSelect)",
            "5.4 Custom React Hooks (useSearchPerusahaan)"
        ]),
        ("BAB 6: DOKUMENTASI HELPER LIBRARY, MASTER DATA, DAN VALIDASI", [
            "6.1 Pustaka Helper Server & Klien (jwt, auth-server, visit-auth, api-helpers, mongodb)",
            "6.2 Pustaka Master Data Wilayah & Merek (wilayah 38 provinsi, merek produk, status)",
            "6.3 Skema Validasi Formulir & Logika Sanitasi"
        ]),
        ("BAB 7: STRUKTUR PENGUJIAN & QUALITY ASSURANCE", [
            "7.1 Unit & Integration Testing dengan Jest (TDD)",
            "7.2 Behavior-Driven Development (BDD) dengan Cucumber.js"
        ]),
        ("BAB 8: INVENTARIS LENGKAP FILE KODE SUMBER & CUPLIKAN LOGIKA INTI", [
            "8.1 Tabel Inventaris Seluruh File Sumber Kode Proyek",
            "8.2 Cuplikan Kode Algoritma & Logika Kunci Sistem"
        ])
    ]

    for chap, subchaps in toc_items:
        p_ch = doc.add_paragraph()
        p_ch.paragraph_format.space_before = Pt(6)
        p_ch.paragraph_format.space_after = Pt(2)
        r_ch = p_ch.add_run(chap)
        r_ch.font.bold = True
        r_ch.font.color.rgb = RGBColor(0x1E, 0x3A, 0x8A)
        r_ch.font.size = Pt(10.5)
        
        for sub in subchaps:
            p_sub = doc.add_paragraph()
            p_sub.paragraph_format.left_indent = Inches(0.25)
            p_sub.paragraph_format.space_before = Pt(0)
            p_sub.paragraph_format.space_after = Pt(1)
            r_sub = p_sub.add_run(f"• {sub}")
            r_sub.font.size = Pt(9)
            r_sub.font.color.rgb = RGBColor(0x47, 0x55, 0x69)

    doc.add_page_break()

    # ==========================================
    # BAB 1: PENDAHULUAN & ARSITEKTUR SISTEM
    # ==========================================
    h1 = doc.add_heading(level=1)
    r = h1.add_run("BAB 1: PENDAHULUAN & ARSITEKTUR SISTEM")
    r.font.name = "Calibri"
    r.font.size = Pt(16)
    r.font.bold = True
    r.font.color.rgb = RGBColor(0x1E, 0x3A, 0x8A)
    
    p = doc.add_paragraph()
    p.add_run(
        "MabelHub adalah platform perangkat lunak berbasis web yang dirancang khusus untuk memodernisasi dan "
        "mengotomatisasi operasional bisnis ke pemerintah (Business to Government / B2G) di bidang pengadaan barang dan jasa furnitur/mebel. "
        "Sistem ini menyatukan berbagai lini operasional perusahaan, mulai dari akuisisi data prospek lembaga/satker pemerintah, "
        "aktivitas telemarketing melalui penyiaran pesan (broadcast) dan panggilan telepon (tele-calling), validasi data prospek oleh tim sales, "
        "perencanaan dan pelaporan kunjungan fisik (visit) sales ke instansi, hingga proses administrasi tender dan penawaran harga (SPH) pada platform e-procurement."
    )
    
    h2 = doc.add_heading(level=2)
    r = h2.add_run("1.1 Latar Belakang & Domain Masalah Bisnis B2G")
    r.font.bold = True
    r.font.color.rgb = RGBColor(0x25, 0x63, 0xEB)
    
    p = doc.add_paragraph()
    p.add_run(
        "Pada industri pengadaan furnitur pemerintah di Indonesia, transaksi memiliki karakteristik kompleksitas tinggi, antara lain:\n"
        "1. Regulasi Pengadaan Pemerintah: Melibatkan e-Katalog LKPP, LPSE, dan penunjukan langsung yang membutuhkan kepatuhan spesifikasi teknis, sertifikat TKDN (Tingkat Komponen Dalam Negeri), dan dokumen SPH resmi.\n"
        "2. Persebaran Wilayah yang Sangat Luas: Menjangkau ribuan Satuan Kerja (Satker), Organisasi Perangkat Daerah (OPD), dinas, kementerian, dan instansi pendidikan di 38 provinsi di Indonesia.\n"
        "3. Kebutuhan Sinkronisasi Multi-Divisi: Diperlukan koordinasi intensif antara divisi Telemarketing (yang mengolah database prospek dan melakukan follow-up awal), divisi Sales Lapangan (yang melakukan presentasi langsung dan negosiasi), divisi E-Procurement/Tender (yang memproses dokumen administrasi), serta pihak Manajemen/Leader untuk pemantauan target omzet."
    )
    
    add_callout(
        doc,
        "Sebelum implementasi MabelHub, pencatatan prospek, status penawaran tender, dan laporan kunjungan sales umumnya dilakukan menggunakan spreadsheet terpisah yang rawan terjadi tumpang tindih kontak PIC instansi, hilangnya jejak riwayat komunikasi, serta lambatnya respons terhadap paket pengadaan yang memiliki batas waktu (deadline) ketat.",
        title="LATAR BELAKANG OPERASIONAL"
    )

    h2 = doc.add_heading(level=2)
    r = h2.add_run("1.2 Tujuan dan Manfaat Sistem")
    r.font.bold = True
    r.font.color.rgb = RGBColor(0x25, 0x63, 0xEB)
    
    p = doc.add_paragraph()
    p.add_run(
        "Tujuan pembangunan sistem MabelHub meliputi:\n"
        "• Sentralisasi Database Instansi & Satker: Menyediakan 'Single Source of Truth' data instansi pemerintah, PIC, nomor kontak terverifikasi, riwayat kunjungan, dan status interaksi.\n"
        "• Tracking Pipeline Penjualan yang Terstruktur: Menerapkan corong penjualan (funnel) mulai dari Input Database → Tracking Database → Broadcast / Call → Validasi Sales → Plan Activity → Visit Lapangan → Request E-Procurement (SPH) → Kontrak / Closing.\n"
        "• Akuntabilitas Kunjungan Sales dengan GIS: Melacak lokasi kunjungan sales dengan koordinat GPS dan pemetaan visual interaktif menggunakan Leaflet GIS.\n"
        "• Efisiensi Dokumen E-Procurement: Mempercepat alur permintaan SPH, review harga oleh Admin, delegasi penanganan tender, dan pengunduhan dokumen pendukung."
    )

    h2 = doc.add_heading(level=2)
    r = h2.add_run("1.3 Arsitektur Tingkat Tinggi (High-Level Architecture)")
    r.font.bold = True
    r.font.color.rgb = RGBColor(0x25, 0x63, 0xEB)
    
    p = doc.add_paragraph()
    p.add_run(
        "MabelHub mengadopsi pola arsitektur Fullstack Modern berbasis Next.js App Router (Server-Side Rendering, React Server Components, dan Client Components) "
        "yang berkomunikasi dengan basis data dokumen NoSQL MongoDB. Arsitektur sistem terbagi ke dalam empat layer utama:"
    )

    arch_table = doc.add_table(rows=5, cols=2)
    format_styled_table(arch_table, [Inches(2.0), Inches(4.5)], ["Layer Arsitektur", "Komponen & Tanggung Jawab"])
    
    arch_layers = [
        ("Presentation Layer (Frontend)", "Dibangun dengan Next.js 16, React 19, Tailwind CSS, Lucide Icons, Recharts untuk visualisasi data, dan Leaflet GIS untuk pemetaan sebaran sales."),
        ("Application & Routing Layer", "Next.js App Router menangani navigasi rute halaman klien serta Server Actions / Route Handlers RESTful API dengan parsing request asynchronous."),
        ("Security & Business Logic Layer", "Middleware guard untuk validasi sesi cookie, JWT stateless auth helper, Zod schema validation, RBAC policy evaluator, dan Excel generation engine."),
        ("Data Persistence Layer", "MongoDB NoSQL Database dengan connection pooling (MongoClient), serta MongoDB GridFS untuk penyimpanan biner dokumen produk dan brosur teknis.")
    ]
    for i, (ly, desc) in enumerate(arch_layers):
        add_table_row(arch_table, [Inches(2.0), Inches(4.5)], [ly, desc], is_even=(i % 2 == 1))
        
    doc.add_paragraph().paragraph_format.space_before = Pt(8)

    h2 = doc.add_heading(level=2)
    r = h2.add_run("1.4 Analisis Tech Stack & Pustaka Dependensi")
    r.font.bold = True
    r.font.color.rgb = RGBColor(0x25, 0x63, 0xEB)
    
    p = doc.add_paragraph()
    p.add_run("Berikut adalah rincian pustaka perangkat lunak (dependencies) yang digunakan pada proyek berdasarkan konfigurasi package.json:")

    dep_table = doc.add_table(rows=1, cols=3)
    format_styled_table(dep_table, [Inches(1.8), Inches(1.2), Inches(3.5)], ["Nama Pustaka", "Versi", "Fungsi & Peranan dalam Sistem"])
    
    deps = [
        ("next", "^16.2.6", "Framework React Fullstack generasi terbaru pendukung App Router & Server Components"),
        ("react & react-dom", "19.2.3", "Pustaka utama perender antarmuka pengguna berbasis komponen reaktif"),
        ("typescript", "^5.x", "Penyedia sistem tipe statis (static typing) untuk menjamin reliabilitas kode"),
        ("mongodb", "native driver", "Konektor resmi MongoDB untuk eksekusi kueri, agregasi, dan transaksi"),
        ("jsonwebtoken", "^9.0.3", "Pembuatan dan verifikasi token autentikasi JWT terenkripsi untuk sesi pengguna"),
        ("bcryptjs", "^3.0.3", "Hashing password pengguna dengan salt untuk keamanan kredensial akun"),
        ("tailwindcss", "^4.x", "Utility-first CSS framework generasi terbaru untuk styling modern dan responsif"),
        ("leaflet & @types/leaflet", "^1.9.4", "Pustaka peta digital open-source untuk visualisasi lokasi instansi dan visit sales"),
        ("recharts", "^3.7.0", "Komponen grafik analitik (bar chart, line chart, pie chart) pada dashboard"),
        ("xlsx", "^0.18.5", "Ekstraksi dan pembacaan berkas spreadsheet Excel untuk fitur export/import laporan"),
        ("lucide-react", "^0.563.0", "Koleksi ikon SVG vektor modern untuk elemen navigasi dan tombol UI"),
        ("react-select", "^5.10.2", "Komponen dropdown interaktif dengan fitur pencarian instan (searchable select)"),
        ("jest & @testing-library", "^30.x", "Framework pengujian otomatis untuk unit test dan integration testing (TDD)"),
        ("cucumber & selenium", "^13.x", "Framework Behavioral-Driven Development (BDD) untuk pengujian end-to-end skenario")
    ]
    for i, (name, ver, func) in enumerate(deps):
        add_table_row(dep_table, [Inches(1.8), Inches(1.2), Inches(3.5)], [name, ver, func], is_even=(i % 2 == 1))

    doc.add_page_break()

    # ==========================================
    # BAB 2: ARSITEKTUR KEAMANAN, OTENTIKASI, & BASIS DATA
    # ==========================================
    h1 = doc.add_heading(level=1)
    r = h1.add_run("BAB 2: ARSITEKTUR KEAMANAN, OTENTIKASI, & BASIS DATA")
    r.font.bold = True
    r.font.size = Pt(16)
    r.font.color.rgb = RGBColor(0x1E, 0x3A, 0x8A)
    
    p = doc.add_paragraph()
    p.add_run(
        "Keamanan sistem MabelHub dirancang untuk melindungi integritas data B2G dan mencegah akses tidak terotorisasi. "
        "Sistem menggunakan pola autentikasi berbasis stateless token JWT yang disimpan dalam cookie terproteksi, "
        "serta lapisan middleware yang memeriksa izin akses sebelum permintaan mencapai komponen halaman atau endpoint backend."
    )

    h2 = doc.add_heading(level=2)
    r = h2.add_run("2.1 Mekanisme Otentikasi & Sesi Cookie")
    r.font.bold = True
    r.font.color.rgb = RGBColor(0x25, 0x63, 0xEB)
    
    p = doc.add_paragraph()
    p.add_run(
        "Alur otentikasi login pengguna berjalan sebagai berikut:\n"
        "1. Klien mengirimkan formulir login (username/email dan password plaintext) ke endpoint POST /api/auth/login.\n"
        "2. Handler mencari dokumen akun pada koleksi 'users' di MongoDB berdasarkan username/email.\n"
        "3. Sistem memverifikasi kecocokan password menggunakan bcrypt.compare().\n"
        "4. Jika valid, payload sesi dibentuk yang mencakup: userId, username, nama lengkap, role pengguna, dan teamId.\n"
        "5. Token ditandatangani menggunakan kunci rahasia JWT_SECRET dengan masa kedaluwarsa 7 hari.\n"
        "6. Token disuntikkan ke dalam header HTTP Response Set-Cookie dengan opsi: HttpOnly=true, Path=/, SameSite=Lax, MaxAge=7 hari."
    )
    
    add_code_block(doc, 
"""// Cuplikan Verifikasi JWT (src/lib/jwt.ts)
export function signSession(payload: SessionPayload): string {
  return jwt.sign(payload, JWT_SECRET, { expiresIn: '7d' });
}

export function verifySession(token: string): SessionPayload | null {
  try {
    return jwt.verify(token, JWT_SECRET) as SessionPayload;
  } catch (error) {
    return null;
  }
}"""
    )

    h2 = doc.add_heading(level=2)
    r = h2.add_run("2.2 Role-Based Access Control (RBAC) & Middleware Protection")
    r.font.bold = True
    r.font.color.rgb = RGBColor(0x25, 0x63, 0xEB)
    
    p = doc.add_paragraph()
    p.add_run(
        "Sistem menerapkan pembagian hak akses berdasarkan 5 peran utama (Roles) dengan matriks kewenangan sebagai berikut:"
    )

    rbac_table = doc.add_table(rows=1, cols=3)
    format_styled_table(rbac_table, [Inches(1.5), Inches(2.2), Inches(2.8)], ["Role", "Cakupan Modul", "Wewenang & Fungsi"])
    
    rbacs = [
        ("SUPERADMIN", "Seluruh Modul Tanpa Batas", "Memiliki kontrol penuh atas user management, audit log, approval request instansi, parameter sistem, tim sales, dan konfigurasi database."),
        ("ADMIN", "E-Procurement Response, Instansi, Database", "Memproses pengajuan SPH, melakukan pricing approval, verifikasi dokumen tender, input database prospek, dan monitoring progress tim."),
        ("LEADER", "Monitoring Tim, Approval Visit, Dashboard", "Memimpin tim sales tertentu, memantau pencapaian kuota anggota tim, menginspeksi rute kunjungan, dan menyetujui plan activity."),
        ("SALES", "Visit, Validasi, Plan Activity, Request SPH", "Melakukan validasi database prospek, membuat rencana kunjungan satker, mencatat log visit lapangan dengan GPS, dan mengajukan SPH tender."),
        ("TELEMARKETING", "Input DB, Tracking DB, Broadcast, Call", "Fokus pada pengolahan prospek awal, pengiriman broadcast penawaran, melakukan panggilan tele-calling, dan kualifikasi status kontak.")
    ]
    for i, (rl, mod, wew) in enumerate(rbacs):
        add_table_row(rbac_table, [Inches(1.5), Inches(2.2), Inches(2.8)], [rl, mod, wew], is_even=(i % 2 == 1))

    doc.add_paragraph().paragraph_format.space_before = Pt(8)

    h2 = doc.add_heading(level=2)
    r = h2.add_run("2.3 Desain & Skema 22 Koleksi Basis Data MongoDB")
    r.font.bold = True
    r.font.color.rgb = RGBColor(0x25, 0x63, 0xEB)
    
    p = doc.add_paragraph()
    p.add_run(
        "Basis data MabelHub memanfaatkan fleksibilitas dokumen MongoDB NoSQL untuk menyimpan data operasional sales dan e-procurement yang dinamis. "
        "Berikut adalah inventarisasi dan struktur fungsional dari 22 koleksi MongoDB yang digunakan dalam sistem:"
    )

    mongo_table = doc.add_table(rows=1, cols=3)
    format_styled_table(mongo_table, [Inches(2.0), Inches(2.2), Inches(2.3)], ["Nama Koleksi", "Field-Field Kunci", "Fungsi Operasional"])
    
    collections_data = [
        ("users", "_id, username, password, name, role, teamId, createdAt", "Menyimpan kredensial akun, hash password, dan penugasan peran (RBAC)"),
        ("teams", "_id, name, leaderId, memberIds, targetOmzet, createdAt", "Struktur organisasi tim sales, pemetaan ketua tim, dan anggota"),
        ("companies", "_id, code, name, address, province, city, pic, phone, segment, status", "Data master instansi pemerintah/satker, alamat, dan kontak PIC"),
        ("company_history", "_id, companyId, updatedBy, changes, timestamp", "Log audit riwayat pengubahan data profil instansi pemerintah"),
        ("company_requests", "_id, companyData, requestedBy, status, approvedBy, note", "Pengajuan pembuatan/perubahan instansi baru oleh sales yang menunggu approval"),
        ("PicChangeHistory", "_id, companyId, oldPic, newPic, changedBy, changedAt", "Riwayat pergantian person-in-charge (PIC) pada satker tertentu"),
        ("VisitActivity", "_id, salesId, companyId, date, purpose, result, lat, lng, photos", "Laporan kunjungan fisik sales, hasil negosiasi, titik koordinat, dan foto bukti"),
        ("plans", "_id, salesId, week, month, targetCompanies, status, approvedBy", "Rencana kerja kunjungan (plan activity) mingguan sales ke satker"),
        ("eproc_requests", "_id, requestId, tenderName, satker, value, status, sphDoc, takenBy", "Pengajuan penawaran tender dan permintaan Surat Penawaran Harga (SPH)"),
        ("eproc_history", "_id, requestId, action, actorId, comment, timestamp", "Jejak riwayat alur pengadaan tender (submitted, taken, responded, won/lost)"),
        ("input_database", "_id, code, agencyName, source, picName, phone, status, category", "Data input mentah prospek hasil riset pasar atau lelang LPSE"),
        ("input_database_history", "_id, databaseId, modifierId, fieldModified, timestamp", "Log perubahan pada data prospek mentah telemarketing"),
        ("tracking_database", "_id, agencyCode, brand, segment, province, city, progressStatus", "Pelacakan status perkembangan prospek dari database ke sales funnel"),
        ("tracking_broadcast", "_id, broadcastId, satkerName, phone, messageStatus, sentAt", "Pelacakan pengiriman pesan broadcast penawaran katalog ke instansi"),
        ("tracking_call", "_id, callId, satkerName, pic, callResult, followUpDate, notes", "Pencatatan hasil panggilan telepon telemarketing kepada PIC instansi"),
        ("validasi_sales", "_id, leadId, validatedBy, decision, suitabilityScore, handoffDate", "Validasi kelayakan prospek sebelum diteruskan ke sales lapangan"),
        ("telemarketing_database", "_id, companyName, contactPerson, directPhone, remarks", "Database khusus kontak prospek yang dikelola divisi telemarketing"),
        ("digital_marketing_performance", "_id, campaignName, impressions, clicks, leadsGenerated, cost", "Metrik analitik performa periklanan digital dan kampanye daring"),
        ("digital_marketing_reports", "_id, reportPeriod, summary, metricsBreakdown, submittedBy", "Laporan periodik bulanan performa akuisisi prospek digital"),
        ("product_categories", "_id, code, name, description, icon, activeStatus", "Master kategori produk mebel (misal: Meja Kantor, Kursi Ergonomis, dsb.)"),
        ("product_documents", "_id, categoryId, title, fileName, gridFsId, uploadDate", "Metadata berkas spesifikasi teknis, sertifikat TKDN, dan brosur katalog"),
        ("product_files.files", "_id, filename, length, chunkSize, uploadDate, md5", "Penyimpanan biner dokumen produk berukuran besar via MongoDB GridFS")
    ]
    for i, (cname, cfields, cfunc) in enumerate(collections_data):
        add_table_row(mongo_table, [Inches(2.0), Inches(2.2), Inches(2.3)], [cname, cfields, cfunc], is_even=(i % 2 == 1))

    doc.add_page_break()

    # ==========================================
    # BAB 3: DOKUMENTASI LENGKAP BACKEND API (67 ENDPOINTS)
    # ==========================================
    h1 = doc.add_heading(level=1)
    r = h1.add_run("BAB 3: DOKUMENTASI LENGKAP BACKEND API (67 ENDPOINTS)")
    r.font.bold = True
    r.font.size = Pt(16)
    r.font.color.rgb = RGBColor(0x1E, 0x3A, 0x8A)
    
    p = doc.add_paragraph()
    p.add_run(
        "MabelHub mengimplementasikan 67 Route Handler RESTful API pada direktori src/app/api/. "
        "Setiap endpoint menerima parameter, melakukan validasi payload, mengeksekusi otorisasi berbasis sesi cookie, "
        "dan mengembalikan respon terstruktur dengan format JSON dan HTTP Status Code standar (200 OK, 201 Created, 400 Bad Request, 401 Unauthorized, 403 Forbidden, 404 Not Found, 500 Server Error)."
    )

    # We will document all 67 routes grouped by module
    api_groups = [
        ("3.1 Modul Otentikasi & Sesi Pengguna", [
            ("/api/auth/login", "POST", "PUBLIC", "Kredensial username/email dan password plaintext", "Memvalidasi kredensial pengguna terhadap koleksi 'users', membandingkan bcrypt hash, membentuk token JWT, dan mengeset cookie sesi 'session'.", "200: { user, message } | 401: Invalid credentials"),
            ("/api/auth/logout", "POST", "PUBLIC", "Tidak ada payload", "Menghapus cookie sesi dengan mengeset Max-Age=0 dan mengembalikan respon sukses logout.", "200: { message: 'Logged out successfully' }"),
            ("/api/auth/me", "GET", "LOGGED_IN", "Cookie sesi 'session'", "Memeriksa validitas token JWT dari request cookie dan mengembalikan informasi akun serta peran (role) pengguna yang aktif.", "200: { user: SessionPayload } | 401: Unauthorized")
        ]),
        ("3.2 Modul Pengguna & Manajemen Tim", [
            ("/api/users", "GET, POST", "ADMIN / SUPERADMIN", "GET: Query filter | POST: { username, password, name, role, teamId }", "Mengambil daftar seluruh pengguna sistem atau membuat akun pengguna baru dengan password yang di-hash bcrypt.", "200: Array<User> | 201: { insertedId }"),
            ("/api/users/[id]", "PUT, DELETE", "SUPERADMIN", "PUT: Update user fields | DELETE: ID pengguna di URL", "Memperbarui profil, mengubah password, atau menghapus pengguna dari sistem.", "200: { updated: true } | 200: { deleted: true }"),
            ("/api/teams", "GET, POST", "LEADER / ADMIN", "GET: list teams | POST: { name, leaderId, memberIds }", "Melihat daftar tim sales dan membuat tim baru beserta struktur kepemimpinan.", "200: Array<Team> | 201: { teamId }"),
            ("/api/teams/[teamId]", "GET, PUT, DELETE", "LEADER / ADMIN", "PUT: { name, leaderId } | DELETE: teamId", "Melihat detail tim, memperbarui nama tim, atau membubarkan tim sales.", "200: Team | 200: { success: true }"),
            ("/api/teams/[teamId]/members", "PUT", "LEADER / ADMIN", "{ memberIds: string[] }", "Menetapkan atau memperbarui daftar anggota sales yang tergabung dalam tim.", "200: { updatedMembersCount: n }"),
            ("/api/teams/me", "GET", "SALES / LEADER", "Cookie sesi pengguna aktif", "Mengambil informasi tim dari pengguna yang sedang melakukan login saat ini.", "200: { team: TeamData }"),
            ("/api/teams/me/members", "GET", "LEADER / SALES", "Cookie sesi pengguna", "Mengambil daftar profil dan kontak seluruh rekan satu tim dari pengguna aktif.", "200: Array<Member>"),
            ("/api/teams/me/stats", "GET", "LEADER / SALES", "Query range tanggal", "Menghitung statistik performa tim (total kunjungan, tender yang diajukan, omzet tercapai).", "200: { totalVisits, totalEproc, winRate }")
        ]),
        ("3.3 Modul Perusahaan & Instansi Pemerintah", [
            ("/api/companies", "GET, POST", "LOGGED_IN", "GET: page, limit, search, province, city | POST: CompanyData", "Menampilkan data instansi dengan paginasi dan pencarian multi-kriteria atau menambahkan instansi baru.", "200: { data, total, page, totalPages } | 201: { id }"),
            ("/api/companies/[id]", "GET, PUT, DELETE", "LOGGED_IN / ADMIN", "PUT: CompanyPayload | DELETE: id", "Mengambil detail instansi, memperbarui data profil satker, atau menghapus data instansi.", "200: Company | 200: { updated: true }"),
            ("/api/companies/[id]/history", "GET", "LOGGED_IN", "id instansi pada URL parameter", "Mengambil riwayat log audit perubahan profil, kontak PIC, atau alamat instansi.", "200: Array<AuditLog>"),
            ("/api/companies/suggest", "GET", "LOGGED_IN", "Query string 'q' (minimal 2 karakter)", "Menyediakan autosuggestion nama instansi secara cepat untuk komponen pencarian UI.", "200: Array<{ id, name, code, city }>"),
            ("/api/instansi", "GET", "LOGGED_IN", "Query filter provinsi, kota, segmen", "Mengambil daftar ringkas instansi untuk keperluan dropdown seleksi.", "200: Array<InstansiSummary>"),
            ("/api/perusahaan", "GET", "LOGGED_IN", "Query filter perusahaan/satker", "Endpoint kompatibilitas untuk pencarian cepat nama badan usaha/lembaga.", "200: Array<Perusahaan>")
        ]),
        ("3.4 Modul Pengajuan Instansi (Company Requests)", [
            ("/api/company-requests", "GET, POST", "LOGGED_IN", "GET: status filter | POST: { companyData, notes }", "Sales mengajukan penambahan instansi baru yang memerlukan verifikasi Admin.", "200: Array<Request> | 201: { requestId }"),
            ("/api/company-requests/[id]/approve", "POST", "ADMIN / SUPERADMIN", "{ approvalNotes?: string }", "Menyetujui pengajuan, memindahkan data ke koleksi 'companies', dan mengirimkan notifikasi.", "200: { approved: true, companyId }"),
            ("/api/company-requests/[id]/reject", "POST", "ADMIN / SUPERADMIN", "{ reason: string }", "Menolak pengajuan penambahan instansi dengan menyertakan alasan penolakan.", "200: { rejected: true }")
        ]),
        ("3.5 Modul Kunjungan Sales & Pelacakan Visit Lapangan", [
            ("/api/visits", "GET, POST", "SALES / LEADER / ADMIN", "GET: filter tanggal, salesId, satker | POST: VisitPayload", "Mengambil daftar laporan kunjungan sales atau mencatat laporan kunjungan fisik baru beserta foto & koordinat GPS.", "200: { visits, summary } | 201: { visitId }"),
            ("/api/visits/[id]", "GET, PUT, DELETE", "SALES / LEADER", "PUT: UpdatedVisit | DELETE: id", "Melihat detail kunjungan, mengedit agenda/hasil visit, atau menghapus kunjungan.", "200: VisitDetail | 200: { success: true }"),
            ("/api/visits/bulk", "POST", "SALES / LEADER", "Array<VisitPayload>", "Melakukan input masal riwayat kunjungan sales untuk migrasi data atau offline sync.", "201: { insertedCount: n }"),
            ("/api/visits/by-satker", "GET", "LOGGED_IN", "satkerId query", "Mengambil seluruh riwayat kunjungan sales yang pernah dilakukan ke satu satker spesifik.", "200: Array<VisitRecord>"),
            ("/api/visits/meta", "GET", "LOGGED_IN", "Tidak ada parameter", "Mengambil opsi metadata untuk formulir kunjungan (tujuan visit, jenis respon, status).", "200: { purposes, statuses, methods }"),
            ("/api/visits/stats", "GET", "LEADER / ADMIN", "Query range bulan, tahun, teamId", "Menghitung agregasi statistik kunjungan per sales, per provinsi, dan persentase respon positif.", "200: { totalVisits, uniqueSatkers, chartData }"),
            ("/api/plans", "POST", "SALES", "{ planDate, targetSatkerIds, objective }", "Menyimpan rencana kerja kunjungan sales (plan activity) yang akan dieksekusi.", "201: { planId }")
        ]),
        ("3.6 Modul E-Procurement & Permintaan SPH Tender", [
            ("/api/e-procurement", "GET", "ADMIN / SALES", "Query filter status, bulan, tahun", "Mengambil data ringkasan tender pengadaan barang pada e-katalog / LPSE.", "200: { eprocs, summaryStats }"),
            ("/api/e-procurement/requests", "GET, POST", "SALES / ADMIN", "GET: filter | POST: { satker, paketName, nilaiHps, deadLine, items }", "Sales mengajukan permohonan harga penawaran (SPH) untuk paket pengadaan barang tertentu.", "200: Array<EprocRequest> | 201: { requestId }"),
            ("/api/e-procurement/requests/[requestId]", "GET, PUT", "ADMIN / SALES", "PUT: Update status, nilai penawaran, dokumen SPH", "Melihat rincian paket pengadaan atau memperbarui informasi paket pengadaan.", "200: RequestDetail | 200: { updated: true }"),
            ("/api/e-procurement/requests/[requestId]/take", "POST", "ADMIN", "{ adminId }", "Admin mengklaim ('take') paket permohonan SPH untuk mulai dikerjakan dan dihitung harganya.", "200: { status: 'Taken', handledBy: adminId }"),
            ("/api/e-procurement/requests/[requestId]/admin-response", "PUT", "ADMIN", "{ sphNumber, offerValue, sphFileUrl, remarks }", "Admin menerbitkan nomor SPH resmi, memasukkan nilai penawaran akhir, dan mengunggah dokumen SPH.", "200: { status: 'Responded', sphNumber }"),
            ("/api/e-procurement/requests/[requestId]/tindak-lanjut", "PUT", "SALES", "{ finalStatus: 'MENANG' | 'KALAH', contractValue, notes }", "Sales mengabarkan hasil akhir tender (Menang/Kalah/Batal) dan realisasi nilai kontrak.", "200: { status: 'Completed', result }"),
            ("/api/e-procurement/requests/[requestId]/history", "GET", "LOGGED_IN", "requestId di URL", "Mengambil timeline riwayat proses SPH dari pertama diajukan hingga keputusan menang/kalah.", "200: Array<TimelineEvent>"),
            ("/api/dashboard-request", "GET", "SALES / ADMIN", "Filter periode", "Menyediakan metrik agregasi seluruh permohonan pengadaan untuk dashboard utama.", "200: { totalRequests, pendingResponse, wonTenders }")
        ]),
        ("3.7 Modul Pipeline Telemarketing & Database Prospek", [
            ("/api/input-database", "GET, POST, PUT", "TELEMARKETING / ADMIN", "GET: filter | POST: ProspekData | PUT: UpdateData", "Mengelola input database prospek instansi dari berbagai sumber penelusuran pasar.", "200: Array<DatabaseItem> | 201: { code }"),
            ("/api/input-database/[code]", "GET", "LOGGED_IN", "code di URL", "Melihat detail satu prospek instansi berdasarkan kode unik sistem.", "200: ProspekDetail"),
            ("/api/input-database/history/[code]", "GET", "LOGGED_IN", "code di URL", "Melihat log perubahan data prospek yang dilakukan oleh operator telemarketing.", "200: Array<HistoryLog>"),
            ("/api/tracking-database", "GET", "TELEMARKETING / ADMIN", "Multi-filter: bulan, merek, segmen, provinsi, kota, limit, page", "Menampilkan pipeline prospek dengan kemampuan filter multi-select kompleks dan pagination.", "200: { data, total, pagination, analytics }"),
            ("/api/tracking-database/filters", "GET", "LOGGED_IN", "Tidak ada parameter", "Menyediakan daftar nilai opsi dinamis untuk filter (merek unik, provinsi, segmen).", "200: { provinces, brands, segments }")
        ]),
        ("3.8 Modul Tracking Broadcast & Telemarketing Call", [
            ("/api/tracking-broadcast", "GET, POST", "TELEMARKETING", "GET: filter tanggal, status | POST: { broadcastBatch }", "Menampilkan status penyampaian broadcast pesan WhatsApp dan menyimpan log pengiriman masal.", "200: { data, summary: { sent, read, failed } }"),
            ("/api/tracking-broadcast/filters", "GET", "LOGGED_IN", "Tidak ada parameter", "Menyediakan opsi filter bulan, tahun, dan template pesan broadcast.", "200: { months, years, templates }"),
            ("/api/tracking-broadcast/send", "POST", "TELEMARKETING", "{ contactIds, templateId, scheduledTime }", "Memicu pengiriman pesan broadcast penawaran ke daftar kontak PIC instansi.", "200: { queued: n, status: 'Processing' }"),
            ("/api/tracking-call", "GET", "TELEMARKETING", "Filter status call, respon PIC, tanggal follow-up", "Menampilkan daftar riwayat panggilan tele-calling beserta status kualifikasi prospek.", "200: { calls, summaryStats }"),
            ("/api/tracking-call/filters", "GET", "LOGGED_IN", "Tidak ada parameter", "Menyediakan daftar status hasil call (Tertarik, Sibuk, Salah Sambung, Follow-up).", "200: { callStatuses, responseTypes }"),
            ("/api/tracking-call/send", "POST", "TELEMARKETING", "{ contactId, callStatus, notes, nextFollowUp }", "Menyimpan hasil panggilan tele-calling dan menjadwalkan tanggal hubungi kembali.", "201: { callLogId }")
        ]),
        ("3.9 Modul Validasi & Tindak Lanjut Sales", [
            ("/api/validasi-sales", "GET, POST", "SALES / LEADER", "GET: filter | POST: { leadId, score, validationResult }", "Menampilkan antrean prospek yang membutuhkan verifikasi kelayakan sebelum diserahkan ke sales lapangan.", "200: Array<ValidasiLead> | 201: { id }"),
            ("/api/validasi-sales/send", "POST", "SALES", "{ leadId, assignedSalesId, handoffNotes }", "Menyerahkan prospek yang sudah terverifikasi (qualified lead) ke sales lapangan untuk dikunjungi.", "200: { success: true, assignedTo: salesId }"),
            ("/api/tindak-lanjut", "GET", "SALES", "Filter status tindak lanjut", "Menampilkan daftar instansi prospek yang sedang dalam tahap tindak lanjut aktif.", "200: Array<FollowUpItem>"),
            ("/api/report-progres", "GET", "MANAGEMENT", "Filter rentang waktu, cabang, divisi", "Menghasilkan ringkasan progres konversi prospek dari broadcast/call hingga closing kontrak.", "200: { conversionFunnel, summaryMetrics }")
        ]),
        ("3.10 Modul Produk Hub & Dokumen Katalog GridFS", [
            ("/api/produk/categories", "GET, POST", "LOGGED_IN / ADMIN", "GET: list | POST: { code, name, description, icon }", "Mengambil daftar kategori produk mebel atau menambahkan kategori katalog baru.", "200: Array<Category> | 201: { categoryId }"),
            ("/api/produk/categories/[id]", "GET, PATCH", "LOGGED_IN / ADMIN", "PATCH: { name, description, activeStatus }", "Mengambil detail kategori atau mengubah status aktif/non-aktif kategori produk.", "200: Category | 200: { updated: true }"),
            ("/api/produk/documents", "GET, POST", "LOGGED_IN / ADMIN", "GET: categoryId filter | POST: FormData (file + metadata)", "Melihat dokumen teknis produk atau mengunggah berkas PDF (spesifikasi, brosur, sertifikat TKDN).", "200: Array<Document> | 201: { docId }"),
            ("/api/produk/documents/[id]", "DELETE, PATCH", "ADMIN", "DELETE: id | PATCH: { title, categoryId }", "Menghapus metadata beserta berkas biner di GridFS atau memperbarui judul dokumen.", "200: { deleted: true } | 200: { updated: true }"),
            ("/api/produk/documents/[id]/file", "GET", "LOGGED_IN", "id dokumen pada URL parameter", "Melakukan streaming unduhan berkas PDF produk langsung dari MongoDB GridFS Bucket.", "200: Binary Stream (application/pdf)")
        ]),
        ("3.11 Modul Digital Marketing, Kontrak, & Parameter", [
            ("/api/marketing/digital/performance", "GET, POST", "MARKETING / ADMIN", "GET: filter periode | POST: PerformanceMetrics", "Mencatat dan memantau metrik performa kampanye pemasaran digital (leads, impressions).", "200: Array<PerfMetric> | 201: { id }"),
            ("/api/marketing/digital/reports", "GET, POST", "MARKETING / ADMIN", "GET: list reports | POST: ReportPayload", "Menyimpan dan mengunduh laporan periodik divisi digital marketing.", "200: Array<Report> | 201: { id }"),
            ("/api/marketing/telemarketing/database", "GET, POST, PATCH", "TELEMARKETING", "CRUD kontak prospek telemarketing", "Endpoint khusus untuk pemeliharaan database nomor kontak satker.", "200: Contacts | 201: Created"),
            ("/api/contracts", "GET, POST", "ADMIN / FINANCE", "GET: filter | POST: { contractNumber, clientName, totalValue, startDate, endDate }", "Mencatat kontrak pengadaan barang yang telah berhasil dimenangkan.", "200: Array<Contract> | 201: { contractId }"),
            ("/api/contracts/[nomorKontrak]", "GET, PUT, DELETE", "ADMIN / FINANCE", "nomorKontrak di URL", "Mengelola detail kontrak, mengubah status pembayaran, atau menghapus data kontrak.", "200: ContractDetail"),
            ("/api/contracts/reminders", "GET", "ADMIN / FINANCE", "Tidak ada parameter", "Mengecek kontrak yang mendekati masa kedaluwarsa atau batas pengiriman barang.", "200: Array<ContractReminder>"),
            ("/api/parameters", "GET, POST, DELETE", "ADMIN / SUPERADMIN", "GET: list | POST: { key, value, group } | DELETE: id", "Mengatur parameter sistem global (daftar merek, segmen instansi, opsi status, limit kuota).", "200: Array<Parameter> | 201: { id }"),
            ("/api/download", "GET", "PUBLIC", "Query file path atau token unduhan", "Endpoint utilitas untuk mengunduh template spreadsheet Excel atau berkas publik.", "200: File Download")
        ]),
        ("3.12 Modul Notifikasi Terintegrasi", [
            ("/api/notifications", "GET", "LOGGED_IN", "Query filter unreadOnly, limit", "Mengambil daftar notifikasi terbaru yang ditujukan untuk pengguna yang sedang aktif.", "200: Array<Notification>"),
            ("/api/notifications/unread-count", "GET", "LOGGED_IN", "Cookie sesi", "Menghitung jumlah notifikasi yang belum dibaca untuk indikator badge merah di navbar.", "200: { unreadCount: number }"),
            ("/api/notifications/[id]/read", "PATCH", "LOGGED_IN", "id notifikasi di URL", "Menandai sebuah pesan notifikasi telah dibaca oleh pengguna.", "200: { read: true }"),
            ("/api/notifications/read-all", "POST", "LOGGED_IN", "Cookie sesi", "Menandai seluruh notifikasi milik pengguna aktif sebagai sudah dibaca sekaligus.", "200: { markedCount: n }")
        ])
    ]

    for grp_title, endpoints in api_groups:
        h2 = doc.add_heading(level=2)
        r = h2.add_run(grp_title)
        r.font.bold = True
        r.font.color.rgb = RGBColor(0x25, 0x63, 0xEB)
        
        t_api = doc.add_table(rows=1, cols=4)
        format_styled_table(t_api, [Inches(1.8), Inches(0.8), Inches(1.8), Inches(2.1)], ["Endpoint & Path", "Metode", "Input / Payload", "Logika Bisnis & Respon"])
        
        for i, (path, method, auth, inp, logic, resp) in enumerate(endpoints):
            desc_cell = f"Hak Akses: {auth}\n{logic}\nRespon: {resp}"
            add_table_row(t_api, [Inches(1.8), Inches(0.8), Inches(1.8), Inches(2.1)], [path, method, inp, desc_cell], is_even=(i % 2 == 1))
            
        doc.add_paragraph().paragraph_format.space_before = Pt(6)

    doc.add_page_break()

    # ==========================================
    # BAB 4: DOKUMENTASI HALAMAN FRONTEND (31 PAGES)
    # ==========================================
    h1 = doc.add_heading(level=1)
    r = h1.add_run("BAB 4: DOKUMENTASI HALAMAN FRONTEND (31 PAGES)")
    r.font.bold = True
    r.font.size = Pt(16)
    r.font.color.rgb = RGBColor(0x1E, 0x3A, 0x8A)
    
    p = doc.add_paragraph()
    p.add_run(
        "MabelHub menyediakan 31 halaman antarmuka pengguna responsif yang dibangun menggunakan React 19 dan Next.js App Router. "
        "Setiap halaman dilengkapi dengan manajemen status lokal (hooks useState, useEffect, useMemo), integrasi panggilan asynchronous fetch ke REST API, "
        "serta rendering kondisional berbasis peran otorisasi pengguna."
    )

    page_data = [
        ("/", "src/app/page.tsx", "Semua (Public)", "Halaman gerbang masuk (Login Portal) dengan autentikasi formulir username dan sandi. Mengalihkan ke dashboard setelah login sukses."),
        ("/dashboard", "src/app/dashboard/page.tsx", "Semua Role", "Dashboard utama yang menampilkan metrik KPI ringkas, grafik performa bulanan, dan kartu pintasan menu operasional."),
        ("/dashboard-request", "src/app/dashboard-request/page.tsx", "Sales, Leader, Admin", "Dashboard pelacakan status permohonan kunjungan dan pengajuan SPH dengan diagram lingkaran dan grafik batang."),
        ("/dashboard-response", "src/app/dashboard-response/page.tsx", "Admin, Superadmin", "Dashboard operasional admin untuk memonitor beban antrean SPH tender yang belum ditanggapi dan waktu respons rata-rata."),
        ("/finance", "src/app/finance/page.tsx", "Finance, Admin, Superadmin", "Modul pelacakan keuangan proyek, termin pembayaran kontrak pengadaan, dan pencatatan faktur penagihan."),
        ("/instansi", "src/app/instansi/page.tsx", "Semua Role", "Katalog master data instansi pemerintah / satuan kerja dengan fitur pencarian instan, filter wilayah, dan tombol aksi detail."),
        ("/tambah-instansi", "src/app/tambah-instansi/page.tsx", "Sales, Admin", "Formulir penambahan data satker/instansi baru lengkap dengan nama dinas, alamat, PIC, kontak telepon, dan segmentasi."),
        ("/tracking-satker", "src/app/tracking-satker/page.tsx", "Sales, Leader", "Monitoring riwayat dan frekuensi kunjungan tim sales pada masing-masing satuan kerja di setiap kabupaten/kota."),
        ("/tracking-b2g", "src/app/tracking-b2g/page.tsx", "Sales, Leader", "Pelacakan interaksi dan pipeline penjualan spesifik segmen instansi pemerintah (B2G) dari tahap prospek hingga kontrak."),
        ("/input-database", "src/app/input-database/page.tsx", "Telemarketing, Admin", "Formulir input data prospek hasil penelusuran LPSE atau database kontak untuk dimasukkan ke dalam pipeline telemarketing."),
        ("/tracking-database", "src/app/tracking-database/page.tsx", "Telemarketing, Sales, Admin", "Tabel utama pelacakan prospek dengan filter multi-select (bulan, merek, segmen, provinsi, kota) dan pagination dinamis."),
        ("/tracking-broadcast", "src/app/tracking-broadcast/page.tsx", "Telemarketing", "Antarmuka pemantauan pengiriman pesan broadcast massal WhatsApp ke PIC instansi dan status penerimaan pesan."),
        ("/tracking-call", "src/app/tracking-call/page.tsx", "Telemarketing", "Sistem antrean panggilan tele-calling, pencatatan hasil respon komunikasi PIC, dan penjadwalan reminder follow-up."),
        ("/validasi-sales", "src/app/validasi-sales/page.tsx", "Sales, Leader", "Antarmuka validasi kelayakan data prospek telemarketing sebelum dilakukan eskalasi menjadi rencana kunjungan sales."),
        ("/tindak-lanjut-sales", "src/app/tindak-lanjut-sales/page.tsx", "Sales", "Daftar prospek tervalidasi yang sedang aktif ditindaklanjuti oleh sales lapangan untuk penyusunan proposal penawaran."),
        ("/plan-activity", "src/app/plan-activity/page.tsx", "Sales, Leader", "Papan jadwal rencana kunjungan mingguan sales ke berbagai satker sasaran beserta status approval dari team leader."),
        ("/plan-activity/add", "src/app/plan-activity/add/page.tsx", "Sales", "Formulir pembuatan rencana kunjungan baru dengan pemilihan satker sasaran dan agenda pembahasan."),
        ("/rekapitulasi-visit", "src/app/rekapitulasi-visit/page.tsx", "Sales, Leader, Admin", "Laporan rekapitulasi seluruh kunjungan lapangan sales, dilengkapi galeri foto bukti kunjungan dan peta sebaran GPS."),
        ("/e-procurement", "src/app/e-procurement/page.tsx", "Sales, Admin", "Formulir pengajuan permintaan SPH untuk paket tender pengadaan e-Katalog LPSE lengkap dengan rincian kebutuhan barang."),
        ("/e-procurement-response", "src/app/e-procurement-response/page.tsx", "Admin", "Meja kerja admin pengadaan untuk memproses permintaan SPH, mengisi harga penawaran resmi, dan menerbitkan berkas SPH."),
        ("/rekapitulasi-Eproc", "src/app/rekapitulasi-Eproc/page.tsx", "Sales, Admin, Leader", "Rekapitulasi status pengajuan tender e-procurement (Draft, Diajukan, Ditanggapi, Menang, Kalah, Batal)."),
        ("/rekapitulasi-response", "src/app/rekapitulasi-response/page.tsx", "Admin, Superadmin", "Laporan evaluasi kecepatan respon admin pengadaan dalam menanggapi permohonan SPH dari sales lapangan."),
        ("/report-progres", "src/app/report-progres/page.tsx", "Leader, Management", "Laporan analitik funnel konversi menyeluruh dari broadcast/call hingga realisasi pesanan dan evaluasi performa kuota."),
        ("/sales-report-system", "src/app/sales-report-system/page.tsx", "Sales, Leader", "Sistem pelaporan performa penjualan komprehensif per sales person, per tim, dan per kategori produk mebel."),
        ("/produk", "src/app/produk/page.tsx", "Semua Role", "Katalog produk mebel (Product Hub) terstruktur per kategori, menampilkan spesifikasi furnitur, sertifikat TKDN, dan dokumen."),
        ("/produk/[kategori]", "src/app/produk/[kategori]/page.tsx", "Semua Role", "Halaman katalog dinamis untuk melihat daftar produk dan dokumen spesifikasi teknis dalam satu kategori tertentu."),
        ("/kontrak", "src/app/kontrak/page.tsx", "Admin, Finance", "Pencatatan dan manajemen arsip kontrak tender yang berhasil dimenangkan, nomor kontrak resmi, dan jangka waktu perjanjian."),
        ("/teams", "src/app/teams/page.tsx", "Leader, Admin, Superadmin", "Manajemen struktur tim sales, alokasi ketua tim (leader), dan pemantauan target omzet penjualan per tim."),
        ("/teams/[teamId]", "src/app/teams/[teamId]/page.tsx", "Leader, Admin", "Halaman detail tim yang menampilkan profil anggota, statistik performa masing-masing anggota, dan konfigurasi keanggotaan."),
        ("/add-user", "src/app/add-user/page.tsx", "Admin, Superadmin", "Formulir pendaftaran akun pengguna baru dan penugasan peran (Role) serta tim kerja."),
        ("/parameters", "src/app/parameters/page.tsx", "Superadmin", "Panel konfigurasi parameter sistem, penambahan merek furnitur baru, segmen instansi, dan opsi lookup global.")
    ]

    t_pages = doc.add_table(rows=1, cols=4)
    format_styled_table(t_pages, [Inches(1.5), Inches(1.8), Inches(1.2), Inches(2.0)], ["Rute URL", "Lokasi Berkas", "Hak Akses", "Deskripsi Fitur & Fungsionalitas"])
    
    for i, (rute, filep, role, desc) in enumerate(page_data):
        add_table_row(t_pages, [Inches(1.5), Inches(1.8), Inches(1.2), Inches(2.0)], [rute, filep, role, desc], is_even=(i % 2 == 1))

    doc.add_page_break()

    # ==========================================
    # BAB 5: DOKUMENTASI KOMPONEN UI, MODAL, DAN HOOKS
    # ==========================================
    h1 = doc.add_heading(level=1)
    r = h1.add_run("BAB 5: DOKUMENTASI KOMPONEN UI, MODAL, DAN HOOKS")
    r.font.bold = True
    r.font.size = Pt(16)
    r.font.color.rgb = RGBColor(0x1E, 0x3A, 0x8A)
    
    p = doc.add_paragraph()
    p.add_run(
        "Aplikasi MabelHub menerapkan prinsip desain berbasis komponen modular (Component-Driven Development) "
        "yang memisahkan tampilan, modal dialog, layout navigasi, dan custom logic ke dalam berkas-berkas terisolasi di folder src/components/ dan src/hooks/."
    )

    h2 = doc.add_heading(level=2)
    r = h2.add_run("5.1 Komponen Tata Letak & Navigasi")
    r.font.bold = True
    r.font.color.rgb = RGBColor(0x25, 0x63, 0xEB)
    
    comp_layout = [
        ("AppLayoutWrapper.tsx", "src/components/layout/", "Komponen pembungkus utama aplikasi yang mendeteksi rute saat ini. Jika pengguna berada di halaman login ('/'), wrapper merender konten polos tanpa sidebar. Jika pada halaman aplikasi internal, wrapper merender layout responsif dengan Sidebar di sebelah kiri dan konten utama di sebelah kanan."),
        ("sidebar.tsx", "src/components/sidebar/", "Komponen navigasi vertikal pintar yang membaca profil pengguna dari SessionProvider. Sidebar secara dinamis menyaring dan merender daftar menu sesuai role pengguna (berdasarkan MENUS_BY_ROLE di src/lib/menu.ts), mendukung collapse/expand submenu, indikator rute aktif, tombol logout, serta badge notifikasi unread."),
        ("ThemeProvider.tsx", "src/components/theme/", "Penyedia context tema antarmuka (light/dark mode) menggunakan pustaka next-themes untuk memastikan konsistensi palet visual di seluruh halaman."),
        ("SessionProvider.tsx", "src/components/session/", "React Context Provider yang mendistribusikan data sesi pengguna (userId, username, role, teamId) ke seluruh komponen klien sehingga tidak perlu melakukan re-fetching berulang.")
    ]
    
    t_complay = doc.add_table(rows=1, cols=3)
    format_styled_table(t_complay, [Inches(1.8), Inches(1.5), Inches(3.2)], ["Komponen", "Direktori", "Tanggung Jawab & Logika"])
    for i, (cn, cd, cr) in enumerate(comp_layout):
        add_table_row(t_complay, [Inches(1.8), Inches(1.5), Inches(3.2)], [cn, cd, cr], is_even=(i % 2 == 1))

    doc.add_paragraph().paragraph_format.space_before = Pt(8)

    h2 = doc.add_heading(level=2)
    r = h2.add_run("5.2 Komponen Modal Dialog & Interaktif")
    r.font.bold = True
    r.font.color.rgb = RGBColor(0x25, 0x63, 0xEB)
    
    modals = [
        ("ConfirmModal.tsx", "Modal dialog konfirmasi generik sebelum melakukan aksi destruktif (seperti menghapus data pengguna, membatalkan tender, atau menghapus berkas) dengan tombol konfirmasi dan pembatalan."),
        ("EditInstansiModal.tsx", "Formulir modal komprehensif untuk menyunting profil instansi/satker, memperbarui alamat, nomor telepon, segmen lembaga, dan pergantian nama PIC secara langsung."),
        ("EditVisitModal.tsx", "Modal kompleks (842 baris kode) untuk mereview dan mengoreksi laporan kunjungan sales, mencakup tanggal visit, nama PIC yang ditemui, ringkasan hasil pembicaraan, foto dokumentasi, dan titik koordinat."),
        ("ExportExcelModal.tsx", "Antarmuka pemilihan parameter dan rentang tanggal untuk mengekspor data tabel (kunjungan, prospek, eproc) ke dalam format berkas Microsoft Excel (.xlsx) siap cetak."),
        ("HistoryEprocModal.tsx", "Modal visualisasi timeline riwayat proses pengajuan SPH tender tertentu dari awal submisi hingga hasil tender diumumkan."),
        ("HistoryInstansiModal.tsx", "Modal audit trail yang menampilkan kronologis seluruh perubahan data profil pada instansi tertentu lengkap dengan nama operator yang mengubah."),
        ("NotificationMenu.tsx", "Menu popover interaktif pada navbar yang menampilkan daftar notifikasi real-time, tombol tandai telah dibaca, dan direct-link ke entitas terkait."),
        ("PendingRequestsModal.tsx", "Modal kerja admin untuk meninjau pengajuan penambahan instansi baru dari sales yang berstatus 'Pending Approval'."),
        ("SalesMap.tsx", "Komponen pemetaan interaktif berbasis Leaflet GIS yang menampilkan marker lokasi kunjungan sales, popup informasi satker, dan heatmap konsentrasi aktivitas sales.")
    ]
    
    t_modals = doc.add_table(rows=1, cols=2)
    format_styled_table(t_modals, [Inches(2.0), Inches(4.5)], ["Nama Komponen Modal", "Deskripsi Fungsional"])
    for i, (mn, md) in enumerate(modals):
        add_table_row(t_modals, [Inches(2.0), Inches(4.5)], [mn, md], is_even=(i % 2 == 1))

    doc.add_paragraph().paragraph_format.space_before = Pt(8)

    h2 = doc.add_heading(level=2)
    r = h2.add_run("5.3 Komponen UI Reusable & Custom Hooks")
    r.font.bold = True
    r.font.color.rgb = RGBColor(0x25, 0x63, 0xEB)
    
    uicomp = [
        ("TableCard.tsx", "Komponen container tabel serbaguna dengan header pencarian, filter status, pembungkus responsif, dan kontrol pagination terpadu."),
        ("DatePicker.tsx", "Komponen pemilih tanggal kustom dengan antarmuka kalender ramah pengguna untuk filter rentang waktu."),
        ("SearchableSelect.tsx", "Komponen dropdown pintar berbasis react-select dengan fitur live-filter untuk daftar instansi, provinsi, dan merek yang panjang."),
        ("useSearchPerusahaan.ts", "Custom React Hook yang membungkus logika debouncing pencarian instansi ke API '/api/companies/suggest' untuk menghindari query berlebih.")
    ]
    
    t_uicomp = doc.add_table(rows=1, cols=2)
    format_styled_table(t_uicomp, [Inches(2.0), Inches(4.5)], ["Komponen / Hook", "Deskripsi Fungsionalitas"])
    for i, (un, ud) in enumerate(uicomp):
        add_table_row(t_uicomp, [Inches(2.0), Inches(4.5)], [un, ud], is_even=(i % 2 == 1))

    doc.add_page_break()

    # ==========================================
    # BAB 6: DOKUMENTASI HELPER LIBRARY, MASTER DATA, DAN VALIDASI
    # ==========================================
    h1 = doc.add_heading(level=1)
    r = h1.add_run("BAB 6: DOKUMENTASI HELPER LIBRARY, MASTER DATA, DAN VALIDASI")
    r.font.bold = True
    r.font.size = Pt(16)
    r.font.color.rgb = RGBColor(0x1E, 0x3A, 0x8A)
    
    p = doc.add_paragraph()
    p.add_run(
        "Untuk memelihara kebersihan kode (Clean Code) dan prinsip DRY (Don't Repeat Yourself), "
        "logika bisnis umum, fungsi keamanan, dan data master dipisahkan ke dalam folder src/lib/, src/data/, dan src/utils/."
    )

    h2 = doc.add_heading(level=2)
    r = h2.add_run("6.1 Pustaka Helper Server & Klien (src/lib/)")
    r.font.bold = True
    r.font.color.rgb = RGBColor(0x25, 0x63, 0xEB)
    
    lib_data = [
        ("mongodb.ts", "Mengelola singleton MongoClient promise connection pool ke database MongoDB agar koneksi tidak dibuat berulang kali pada lingkungan development."),
        ("jwt.ts", "Menyediakan fungsi enkripsi dan dekripsi token sesi JWT (signSession dan verifySession) menggunakan algoritma HMAC-SHA256."),
        ("password.ts", "Menyediakan fungsi hashPassword() menggunakan salt bcrypt 10 putaran dan comparePassword() untuk verifikasi otentikasi login."),
        ("auth-server.ts", "Koleksi guard function server-side: getSession(), assertLoggedIn(), assertSuperadmin(), assertAdminOrSuperadmin(), dan assertLeaderOrSales()."),
        ("visit-auth.ts", "Logika otorisasi khusus data kunjungan sales untuk membatasi akses: sales hanya dapat melihat kunjungannya sendiri, leader melihat satu tim, dan admin melihat seluruh tim."),
        ("api-helpers.ts", "Fungsi standar untuk merender respon JSON seragam (jsonSuccess, jsonError) beserta logging exception."),
        ("menu.ts", "Definisi struktur menu bertingkat (MENUS_BY_ROLE) untuk 5 peran pengguna, memuat label menu, path URL, dan icon identitas.")
    ]
    
    t_lib = doc.add_table(rows=1, cols=2)
    format_styled_table(t_lib, [Inches(1.8), Inches(4.7)], ["Nama File", "Fungsi & Implementasi Teknis"])
    for i, (ln, ld) in enumerate(lib_data):
        add_table_row(t_lib, [Inches(1.8), Inches(4.7)], [ln, ld], is_even=(i % 2 == 1))

    doc.add_paragraph().paragraph_format.space_before = Pt(8)

    h2 = doc.add_heading(level=2)
    r = h2.add_run("6.2 Pustaka Master Data & Validasi (src/data/ & src/utils/)")
    r.font.bold = True
    r.font.color.rgb = RGBColor(0x25, 0x63, 0xEB)
    
    data_utils = [
        ("wilayah.ts", "src/data/", "Daftar master hierarki wilayah administratif 38 provinsi di Indonesia beserta seluruh kabupaten dan kota di dalamnya untuk standarisasi geolokasi satker."),
        ("merek.ts", "src/data/", "Katalog merek furnitur dan perlengkapan kantor yang dipasarkan (misal: Chitose, Indachi, Lion, Donati, dsb.) beserta status keaktifannya."),
        ("statusupdatebroadcast.ts", "src/data/", "Konstanta status pesan broadcast WhatsApp (Terkirim, Dibaca, Dibalas, Gagal, Respon Positif, Respon Negatif)."),
        ("statusupdatecall.ts", "src/data/", "Konstanta status panggilan telemarketing (Tersambung - Tertarik, Tersambung - Tidak Tertarik, Salah Nomor, Tidak Diangkat, Minta Dijadwalkan Ulang)."),
        ("validation.ts", "src/utils/", "Fungsi validasi nomor telepon Indonesia, format email, panjang karakter minimum, dan sanitasi input dari serangan XSS/Injection."),
        ("formValidation.ts", "src/utils/", "Pemeriksaan kelengkapan berkas formulir e-procurement (nomor HPS, tanggal batas tender, kelengkapan item barang).")
    ]
    
    t_du = doc.add_table(rows=1, cols=3)
    format_styled_table(t_du, [Inches(1.8), Inches(1.2), Inches(3.5)], ["Nama File", "Folder", "Deskripsi & Peranan"])
    for i, (dn, df, dd) in enumerate(data_utils):
        add_table_row(t_du, [Inches(1.8), Inches(1.2), Inches(3.5)], [dn, df, dd], is_even=(i % 2 == 1))

    doc.add_page_break()

    # ==========================================
    # BAB 7: STRUKTUR PENGUJIAN & QUALITY ASSURANCE
    # ==========================================
    h1 = doc.add_heading(level=1)
    r = h1.add_run("BAB 7: STRUKTUR PENGUJIAN & QUALITY ASSURANCE")
    r.font.bold = True
    r.font.size = Pt(16)
    r.font.color.rgb = RGBColor(0x1E, 0x3A, 0x8A)
    
    p = doc.add_paragraph()
    p.add_run(
        "MabelHub menerapkan pendekatan penjaminan mutu ganda: Test-Driven Development (TDD) menggunakan Jest "
        "dan Behavior-Driven Development (BDD) menggunakan Cucumber.js. Pengujian ini memastikan setiap fitur berfungsi sesuai spesifikasi teknis dan kebutuhan pengguna akhir."
    )

    h2 = doc.add_heading(level=2)
    r = h2.add_run("7.1 Pengujian Unit & Integrasi dengan Jest (TDD)")
    r.font.bold = True
    r.font.color.rgb = RGBColor(0x25, 0x63, 0xEB)
    
    p = doc.add_paragraph()
    p.add_run(
        "Terdapat 16 berkas pengujian otomatis unit dan integrasi di dalam src/ yang menguji fungsi utilitas, endpoint API, dan antarmuka komponen:"
    )

    jest_tests = [
        ("src/lib/utils.test.ts", "39 Pengujian", "Menguji pagination window (getPageWindow), manipulasi class (cn), pemformatan bulan Indonesia, dan helper filter data."),
        ("src/app/api/tracking-broadcast/route.test.ts", "13 Pengujian", "Menguji validasi payload POST, skema respons GET, filter multi-bulan, filter range tanggal, dan batas paginasi."),
        ("src/app/api/tracking-database/route.test.ts", "12 Pengujian", "Menguji agregasi analitik, multi-select filter (merek, provinsi, segmen), dan batasan parameter limit 1-500."),
        ("src/app/plan-activity/page.test.tsx", "Pengujian Komponen", "Menguji interaksi UI pembuatan jadwal kunjungan, validasi formulir, dan rendering daftar plan activity."),
        ("src/app/tracking-database/page.test.tsx", "Pengujian Komponen", "Menguji rendering tabel database prospek, eksekusi debounce filter, dan klik tombol detail."),
        ("src/utils/formValidation.test.ts", "Pengujian Validasi", "Menguji batasan field formulir pengadaan barang, deteksi nilai negatif pada HPS, dan kelengkapan kontak."),
        ("src/utils/validation.test.ts", "Pengujian Validasi", "Menguji regex nomor telepon, format email perusahaan, dan sanitasi string input.")
    ]
    
    t_jest = doc.add_table(rows=1, cols=3)
    format_styled_table(t_jest, [Inches(2.2), Inches(1.3), Inches(3.0)], ["Berkas Pengujian", "Cakupan", "Skenario yang Diverifikasi"])
    for i, (jn, jc, js) in enumerate(jest_tests):
        add_table_row(t_jest, [Inches(2.2), Inches(1.3), Inches(3.0)], [jn, jc, js], is_even=(i % 2 == 1))

    doc.add_paragraph().paragraph_format.space_before = Pt(8)

    h2 = doc.add_heading(level=2)
    r = h2.add_run("7.2 Pengujian Skenario Bisnis dengan Cucumber.js (BDD)")
    r.font.bold = True
    r.font.color.rgb = RGBColor(0x25, 0x63, 0xEB)
    
    p = doc.add_paragraph()
    p.add_run(
        "Pengujian BDD menggunakan sintaks Gherkin (Given-When-Then) untuk memvalidasi alur bisnis end-to-end pada folder tests/features/:"
    )

    bdd_features = [
        ("visit_instansi.feature", "visit_instansi.steps.js", "Memvalidasi alur sales membuka halaman instansi, mencari satker, mengisi formulir laporan kunjungan fisik, mengunggah foto, dan memverifikasi data tersimpan di riwayat visit."),
        ("visit_plan.feature", "visit_plan.steps.js", "Memvalidasi alur sales merencanakan kunjungan mingguan, memilih target satker sasaran, dan notifikasi persetujuan yang diterima oleh Team Leader."),
        ("visit_rekap.feature", "visit_rekap.steps.js", "Memvalidasi alur pimpinan membuka rekapitulasi kunjungan, memfilter berdasarkan rentang tanggal dan nama sales, serta mengekspor data ke Excel."),
        ("product.feature", "product.steps.js", "Memvalidasi penjelajahan katalog mebel, pencarian sertifikat TKDN produk, dan proses pengunduhan dokumen spesifikasi PDF dari GridFS."),
        ("umum.feature", "umum.steps.js", "Memvalidasi alur otentikasi login, penolakan akses untuk sesi tidak sah, transisi role pengguna, dan proses logout.")
    ]
    
    t_bdd = doc.add_table(rows=1, cols=3)
    format_styled_table(t_bdd, [Inches(1.8), Inches(1.7), Inches(3.0)], ["Berkas Fitur (.feature)", "Step Definition (.js)", "Skenario Bisnis yang Diuji"])
    for i, (fn, sn, sd) in enumerate(bdd_features):
        add_table_row(t_bdd, [Inches(1.8), Inches(1.7), Inches(3.0)], [fn, sn, sd], is_even=(i % 2 == 1))

    doc.add_page_break()

    # ==========================================
    # BAB 8: INVENTARIS LENGKAP FILE KODE SUMBER & CUPLIKAN LOGIKA INTI
    # ==========================================
    h1 = doc.add_heading(level=1)
    r = h1.add_run("BAB 8: INVENTARIS LENGKAP FILE KODE SUMBER & CUPLIKAN LOGIKA INTI")
    r.font.bold = True
    r.font.size = Pt(16)
    r.font.color.rgb = RGBColor(0x1E, 0x3A, 0x8A)
    
    p = doc.add_paragraph()
    p.add_run(
        "Bab ini menyajikan inventarisasi lengkap seluruh berkas kode sumber yang menyusun sistem MabelHub, "
        "beserta cuplikan kode dari algoritma dan logika inti sistem."
    )

    h2 = doc.add_heading(level=2)
    r = h2.add_run("8.1 Inventarisasi Seluruh Berkas Kode Sumber Proyek")
    r.font.bold = True
    r.font.color.rgb = RGBColor(0x25, 0x63, 0xEB)
    
    # We will read all src files dynamically and add to a master table
    src_files = sorted(glob.glob("src/**/*", recursive=True))
    src_files = [f for f in src_files if os.path.isfile(f)]
    
    t_all = doc.add_table(rows=1, cols=4)
    format_styled_table(t_all, [Inches(0.4), Inches(2.8), Inches(1.2), Inches(2.1)], ["No", "Path Berkas", "Kategori", "Peran dalam Sistem"])
    
    def classify_file(path):
        p_norm = path.replace('\\', '/')
        if '/api/' in p_norm:
            return "Backend API", "Penangan endpoint RESTful untuk pemrosesan kueri & transaksi database"
        elif p_norm.endswith('page.tsx'):
            return "Frontend Page", "Antarmuka halaman web klien dengan routing Next.js App Router"
        elif '/components/modals/' in p_norm:
            return "Modal UI", "Dialog interaktif pop-up untuk operasi form dan detail data"
        elif '/components/' in p_norm:
            return "Komponen UI", "Elemen visual modular (sidebar, layout, card, picker)"
        elif '/lib/' in p_norm:
            return "Helper Library", "Fungsi utilitas backend, koneksi database, dan autentikasi"
        elif '/data/' in p_norm:
            return "Master Data", "Koleksi data statis (wilayah, merek produk, status)"
        elif '/utils/' in p_norm:
            return "Utility / Validasi", "Fungsi validasi formulir dan format data"
        elif 'test' in p_norm:
            return "Pengujian TDD", "Unit & integration test otomatis berbasis Jest"
        elif '/models/' in p_norm:
            return "Data Model", "Definisi tipe data dan skema entitas database"
        elif '/hooks/' in p_norm:
            return "Custom Hook", "Pustaka status reaktif kustom React"
        else:
            return "Konfigurasi / Root", "Berkas konfigurasi aplikasi dan styling global"

    row_count = 0
    for idx, fpath in enumerate(src_files):
        rel_path = os.path.relpath(fpath, ".").replace('\\', '/')
        cat, role_desc = classify_file(rel_path)
        row_count += 1
        add_table_row(t_all, [Inches(0.4), Inches(2.8), Inches(1.2), Inches(2.1)], [str(row_count), rel_path, cat, role_desc], is_even=(row_count % 2 == 0))

    doc.add_paragraph().paragraph_format.space_before = Pt(8)

    h2 = doc.add_heading(level=2)
    r = h2.add_run("8.2 Cuplikan Kode Kunci (Core Implementation Snippets)")
    r.font.bold = True
    r.font.color.rgb = RGBColor(0x25, 0x63, 0xEB)
    
    p = doc.add_paragraph()
    p.add_run("Berikut adalah cuplikan implementasi kode dari fitur-fitur sentral pada MabelHub:")

    # Snippet 1: Middleware
    p_snip1 = doc.add_paragraph()
    r = p_snip1.add_run("A. Penjaga Rute & Validasi Token Sesi (src/middleware.ts)")
    r.bold = True
    r.font.color.rgb = RGBColor(0x1E, 0x3A, 0x8A)
    add_code_block(doc,
"""export function middleware(req: NextRequest) {
  const { pathname } = req.nextUrl;

  // Izinkan akses ke aset statis dan rute publik
  if (PUBLIC_PATHS.includes(pathname) || pathname.startsWith("/_next") || PUBLIC_FILE_PATTERN.test(pathname)) {
    return NextResponse.next();
  }
  if (pathname.startsWith("/api/auth/")) return NextResponse.next();

  const token = req.cookies.get("session")?.value;

  // Proteksi rute aplikasi internal dan endpoint API
  if (!token) {
    if (pathname.startsWith("/api/")) {
      return NextResponse.json({ error: "Unauthorized: Missing session token" }, { status: 401 });
    }
    const url = req.nextUrl.clone();
    url.pathname = "/";
    return NextResponse.redirect(url);
  }

  return NextResponse.next();
}"""
    )

    # Snippet 2: MongoDB Connection Pooling
    p_snip2 = doc.add_paragraph()
    r = p_snip2.add_run("B. Singleton Connection Pool Database MongoDB (src/lib/mongodb.ts)")
    r.bold = True
    r.font.color.rgb = RGBColor(0x1E, 0x3A, 0x8A)
    add_code_block(doc,
"""import { MongoClient } from "mongodb";

const uri = process.env.MONGODB_URI as string;
if (!uri) throw new Error("Please define MONGODB_URI in .env.local");

let client: MongoClient;
let clientPromise: Promise<MongoClient>;

declare global {
  var _mongoClientPromise: Promise<MongoClient>;
}

if (process.env.NODE_ENV === "development") {
  if (!global._mongoClientPromise) {
    client = new MongoClient(uri);
    global._mongoClientPromise = client.connect();
  }
  clientPromise = global._mongoClientPromise;
} else {
  client = new MongoClient(uri);
  clientPromise = client.connect();
}

export default clientPromise;"""
    )

    # Snippet 3: Dynamic Multi-select Pipeline Aggregation
    p_snip3 = doc.add_paragraph()
    r = p_snip3.add_run("C. Kueri Dinamis Pipeline Prospek Telemarketing (src/app/api/tracking-database/route.ts)")
    r.bold = True
    r.font.color.rgb = RGBColor(0x1E, 0x3A, 0x8A)
    add_code_block(doc,
"""// Membangun filter kueri MongoDB berdasarkan parameter multi-select
const filter: Record<string, any> = {};

if (bulanList.length > 0) filter.bulan = { $in: bulanList };
if (merekList.length > 0) filter.merek = { $in: merekList };
if (segmenList.length > 0) filter.segmen = { $in: segmenList };
if (provinsiList.length > 0) filter.provinsi = { $in: provinsiList };
if (kotaList.length > 0) filter.kota = { $in: kotaList };

// Eksekusi kueri paginasi dan agregasi metrik secara paralel
const [data, totalCount] = await Promise.all([
  collection.find(filter).skip((page - 1) * limit).limit(limit).toArray(),
  collection.countDocuments(filter)
]);

return NextResponse.json({
  data,
  pagination: { total: totalCount, page, limit, totalPages: Math.ceil(totalCount / limit) }
});"""
    )

    # Snippet 4: GridFS Document Streaming
    p_snip4 = doc.add_paragraph()
    r = p_snip4.add_run("D. Streaming Berkas PDF Produk dari MongoDB GridFS (src/app/api/produk/documents/[id]/file/route.ts)")
    r.bold = True
    r.font.color.rgb = RGBColor(0x1E, 0x3A, 0x8A)
    add_code_block(doc,
"""export async function GET(req: Request, { params }: { params: Promise<{ id: string }> }) {
  const { id } = await params;
  const client = await clientPromise;
  const db = client.db();
  const bucket = new GridFSBucket(db, { bucketName: "product_files" });

  const doc = await db.collection("product_documents").findOne({ _id: new ObjectId(id) });
  if (!doc) return NextResponse.json({ error: "Document not found" }, { status: 404 });

  const downloadStream = bucket.openDownloadStream(doc.gridFsId);
  return new Response(downloadStream as any, {
    headers: {
      "Content-Type": "application/pdf",
      "Content-Disposition": `inline; filename="${doc.fileName}"`
    }
  });
}"""
    )

    # Output file paths
    docx_path = os.path.abspath("Dokumentasi_Kode_MabelHub.docx")
    pdf_path = os.path.abspath("Dokumentasi_Kode_MabelHub.pdf")
    
    print(f"Menyimpan berkas DOCX ke: {docx_path}...")
    doc.save(docx_path)
    print("Berkas DOCX berhasil disimpan!")
    
    # Convert to PDF via Word COM
    print("Mengonversi berkas ke format PDF via Microsoft Word Automation...")
    import win32com.client
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    try:
        doc_com = word.Documents.Open(docx_path)
        # 17 represents wdFormatPDF
        doc_com.SaveAs(pdf_path, FileFormat=17)
        doc_com.Close()
        print(f"Berkas PDF berhasil dibuat di: {pdf_path}")
    except Exception as e:
        print(f"Error saat konversi PDF: {e}")
    finally:
        word.Quit()
        
    if os.path.exists(pdf_path):
        size_mb = os.path.getsize(pdf_path) / (1024 * 1024)
        print(f"SUKSES! File PDF berukuran {size_mb:.2f} MB siap digunakan.")
    else:
        print("PERINGATAN: File PDF tidak ditemukan.")

if __name__ == "__main__":
    build_documentation()
