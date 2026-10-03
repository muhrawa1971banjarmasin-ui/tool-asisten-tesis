import streamlit as st
import pandas as pd
import PyPDF2
from datetime import datetime
import json
import urllib.request
import urllib.error
try:
    import docx
except ImportError:
    docx = None


# ============================================================
# KONFIGURASI
# ============================================================
st.set_page_config(
    page_title="Asisten Akademik AI",
    page_icon="🎓",
    layout="wide"
)

APP_NAME = "Asisten Akademik AI"
APP_SUBTITLE = (
    "Perkuliahan • Riset • Referensi • Tesis • "
    "Disertasi • Publikasi • Presentasi • Sidang"
)


# ============================================================
# SESSION STATE
# Catatan: penyimpanan ini masih sementara selama sesi.
# Database multi-user permanen dipasang pada tahap berikutnya.
# ============================================================
if "bank_karya" not in st.session_state:
    st.session_state.bank_karya = []

if "bank_referensi" not in st.session_state:
    st.session_state.bank_referensi = []

if "proyek_aktif" not in st.session_state:
    st.session_state.proyek_aktif = "Proyek Utama"


# ============================================================
# FUNGSI DASAR
# ============================================================
def format_ukuran(byte):
    if byte < 1024:
        return f"{byte} B"
    if byte < 1024 * 1024:
        return f"{byte / 1024:.1f} KB"
    return f"{byte / (1024 * 1024):.1f} MB"


def baca_pdf(file):
    try:
        file.seek(0)
        reader = PyPDF2.PdfReader(file)
        teks = []

        for halaman in reader.pages:
            isi = halaman.extract_text()
            if isi:
                teks.append(isi)

        return "\n".join(teks)

    except Exception as e:
        return f"ERROR: {e}"


def baca_docx(file):
    if docx is None:
        return "ERROR: Library python-docx belum tersedia."

    try:
        file.seek(0)
        dokumen = docx.Document(file)
        return "\n".join(
            paragraf.text for paragraf in dokumen.paragraphs
            if paragraf.text.strip()
        )

    except Exception as e:
        return f"ERROR: {e}"


def baca_txt(file):
    try:
        file.seek(0)
        data = file.getvalue()

        try:
            return data.decode("utf-8")
        except UnicodeDecodeError:
            return data.decode("latin-1")

    except Exception as e:
        return f"ERROR: {e}"


def ekstrak_teks(file):
    nama = file.name.lower()

    if nama.endswith(".pdf"):
        return baca_pdf(file)

    if nama.endswith(".docx"):
        return baca_docx(file)

    if nama.endswith(".txt"):
        return baca_txt(file)

    return ""


def deteksi_struktur(teks):
    teks = teks.lower()

    struktur = {
        "Latar Belakang": ["latar belakang"],
        "Rumusan Masalah": [
            "rumusan masalah",
            "fokus penelitian"
        ],
        "Tujuan Penelitian": ["tujuan penelitian"],
        "Kajian Pustaka/Teori": [
            "kajian pustaka",
            "kajian teori",
            "landasan teori"
        ],
        "Metode Penelitian": [
            "metode penelitian",
            "metodologi penelitian"
        ],
        "Hasil/Temuan": [
            "hasil penelitian",
            "hasil dan pembahasan",
            "temuan penelitian"
        ],
        "Kesimpulan": [
            "kesimpulan",
            "simpulan"
        ],
        "Keterbatasan": [
            "keterbatasan penelitian"
        ],
        "Daftar Pustaka": [
            "daftar pustaka",
            "references"
        ]
    }

    hasil = {}

    for bagian, kata_kunci in struktur.items():
        hasil[bagian] = any(
            kata in teks for kata in kata_kunci
        )

    return hasil


def simpan_karya(nama, jenis, ukuran):
    data = {
        "Proyek": st.session_state.proyek_aktif,
        "Nama File": nama,
        "Jenis": jenis,
        "Ukuran": ukuran,
        "Tanggal": datetime.now().strftime("%d-%m-%Y %H:%M")
    }
    st.session_state.bank_karya.append(data)
# ============================================================
# KONEKSI AI 9ROUTER
# ============================================================
def analisis_dengan_9router(teks, jenis_karya, fokus_analisis):
    try:
        api_key = st.secrets["NINEROUTER_API_KEY"]
        base_url = st.secrets["NINEROUTER_BASE_URL"].rstrip("/")
    except Exception:
        return "Konfigurasi 9Router belum ditemukan di Streamlit Secrets."

    # Batasi teks pada tahap awal agar permintaan tetap stabil
    teks_dokumen = teks[:60000]

    fokus = ", ".join(fokus_analisis) if fokus_analisis else "Analisis akademik menyeluruh"

    prompt = f"""
Anda adalah Asisten Akademik AI untuk mahasiswa S2 dan S3.

Analisis dokumen akademik berikut secara teliti dan hanya berdasarkan
isi dokumen yang diberikan.

Jenis karya yang dipilih:
{jenis_karya}

Fokus analisis:
{fokus}

ATURAN WAJIB:
1. Jangan mengarang informasi.
2. Jika informasi tidak ditemukan, tulis: "Tidak ditemukan dalam dokumen."
3. Jangan membuat nama penulis, teori, metode, hasil, referensi, DOI,
   research gap, atau novelty yang tidak terdapat dalam dokumen.
4. Bedakan antara novelty yang diklaim penulis dengan novelty yang
   benar-benar telah diverifikasi melalui literatur.
5. Pada tahap ini Anda hanya menganalisis dokumen, bukan membuktikan
   novelty terhadap seluruh literatur ilmiah.
6. Gunakan bahasa Indonesia akademik yang jelas.
7. Berikan bukti atau bagian dokumen yang mendukung analisis jika tersedia.

Susun hasil dengan bagian:

A. IDENTITAS DAN JENIS DOKUMEN
B. TOPIK UTAMA
C. LATAR BELAKANG / MASALAH
D. TUJUAN
E. KONSEP ATAU LANDASAN TEORI
F. METODOLOGI
G. TEMUAN / HASIL UTAMA
H. KETERBATASAN
I. RESEARCH GAP YANG TERIDENTIFIKASI
J. NOVELTY / KEBAHARUAN YANG DIKLAIM
K. KONTRIBUSI AKADEMIK
L. RELEVANSI UNTUK PENELITIAN LANJUTAN
M. KESIMPULAN ANALISIS

DOKUMEN:
--------------------
{teks_dokumen}
--------------------
"""

    payload = {
        "model": "auto",
        "messages": [
            {
                "role": "system",
                "content": "Anda adalah asisten analisis akademik yang akurat dan tidak mengarang data."
            },
            {
                "role": "user",
                "content": prompt
            }
        ],
        "temperature": 0.2
    }

    request = urllib.request.Request(
        base_url + "/chat/completions",
        data=json.dumps(payload).encode("utf-8"),
        headers={
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json"
        },
        method="POST"
    )

    try:
        with urllib.request.urlopen(request, timeout=180) as response:
            hasil = json.loads(response.read().decode("utf-8"))

        return hasil["choices"][0]["message"]["content"]

    except urllib.error.HTTPError as e:
        try:
            detail = e.read().decode("utf-8")
        except Exception:
            detail = ""
        return f"9Router belum berhasil memproses permintaan. HTTP {e.code}. {detail}"

    except urllib.error.URLError as e:
        return f"Tidak dapat terhubung ke 9Router: {e.reason}"

    except Exception as e:
        return f"Terjadi kesalahan saat menjalankan analisis AI: {str(e)}"
    sudah_ada = any(
        x["Nama File"] == nama
        and x["Proyek"] == st.session_state.proyek_aktif
        for x in st.session_state.bank_karya
    )

    if not sudah_ada:
        st.session_state.bank_karya.append(data)
        return True

    return False


# ============================================================
# HEADER
# ============================================================
st.title("🎓 Asisten Akademik AI")
st.caption(APP_SUBTITLE)

st.info(
    "Tahap 1 membangun fondasi aplikasi. "
    "Fitur AI, database multi-user, pencarian sumber akademik, "
    "Zotero/Mendeley, statistik lanjutan, transkripsi, pembayaran, "
    "dan kredit AI akan dipasang bertahap setelah fondasi stabil."
)


# ============================================================
# SIDEBAR
# ============================================================
st.sidebar.title("🎓 ASISTEN AKADEMIK AI")

st.sidebar.text_input(
    "Proyek Aktif",
    key="proyek_aktif"
)

menu = st.sidebar.radio(
    "Menu Utama",
    [
        "🏠 Beranda",
        "📚 Perkuliahan",
        "🔬 Analisis Karya Akademik",
        "🔎 Literatur & Referensi",
        "🎓 Tesis S2",
        "🧑‍🎓 Disertasi S3",
        "🧭 Metodologi Penelitian",
        "📝 Instrumen Penelitian",
        "📊 Statistik & SPSS",
        "🔤 Analisis Kualitatif",
        "🎤 Audio & Video",
        "✍️ Penulisan Akademik",
        "👨‍🏫 Bimbingan & Revisi",
        "📑 Publikasi Jurnal",
        "📂 Perpustakaan Akademik",
        "✅ Audit Akademik",
        "📈 Progres Penelitian",
        "🖥️ Presentasi",
        "🎓 Simulasi Sidang",
        "💚 Donasi & Akses",
        "⚙️ Admin"
    ]
)

st.sidebar.divider()
st.sidebar.caption(
    "Data permanen per pengguna dan per proyek "
    "akan menggunakan database pada tahap berikutnya."
)


# ============================================================
# BERANDA
# ============================================================
if menu == "🏠 Beranda":

    st.header("🏠 Pusat Asisten Akademik")

    st.write(
        """
        Satu ruang kerja akademik untuk mendampingi proses
        dari tugas perkuliahan sampai tesis, disertasi,
        publikasi dan ujian akademik.
        """
    )

    c1, c2, c3, c4 = st.columns(4)

    c1.metric(
        "Karya dalam Sesi",
        len(st.session_state.bank_karya)
    )

    c2.metric(
        "Referensi dalam Sesi",
        len(st.session_state.bank_referensi)
    )

    c3.metric(
        "Proyek Aktif",
        st.session_state.proyek_aktif
    )

    c4.metric(
        "Tahap Sistem",
        "Fondasi"
    )

    st.divider()

    st.subheader("🧭 Alur Utama")

    st.success(
        "Perkuliahan → Ide Penelitian → Literatur → Metodologi → "
        "Proposal → Penelitian → Analisis → Tesis/Disertasi → "
        "Publikasi → Presentasi → Sidang"
    )

    col1, col2, col3 = st.columns(3)

    with col1:
        st.subheader("📚 Kuliah")
        st.write(
            """
            Tugas kuliah  
            Makalah  
            Resume  
            Review buku  
            Review jurnal  
            Critical review  
            Mini riset  
            Artikel  
            Presentasi
            """
        )

    with col2:
        st.subheader("🎓 Tesis S2")
        st.write(
            """
            Ide dan topik  
            Research gap  
            Judul  
            BAB I–V  
            Instrumen  
            Analisis data  
            Publikasi  
            Sidang
            """
        )

    with col3:
        st.subheader("🧑‍🎓 Disertasi S3")
        st.write(
            """
            Jembatan S2–S3  
            State of the art  
            Novelty doktoral  
            Kontribusi ilmiah  
            Disertasi  
            Publikasi  
            Ujian doktoral
            """
        )


# ============================================================
# PERKULIAHAN
# ============================================================
elif menu == "📚 Perkuliahan":

    st.header("📚 Asisten Perkuliahan")

    fitur = st.selectbox(
        "Pilih pekerjaan",
        [
            "Pahami Instruksi Dosen",
            "Tugas Kuliah",
            "Makalah",
            "Resume",
            "Review Buku",
            "Review Jurnal",
            "Critical Review",
            "Mini Riset",
            "Artikel/Jurnal",
            "Laporan",
            "PPT / Presentasi",
            "Periksa Tugas Saya"
        ]
    )

    st.subheader(fitur)

    instruksi = st.text_area(
        "Masukkan instruksi dosen, tema, atau kebutuhan tugas"
    )

    files = st.file_uploader(
        "Unggah bahan tugas",
        type=[
            "pdf", "docx", "txt", "csv", "xlsx",
            "pptx", "jpg", "jpeg", "png"
        ],
        accept_multiple_files=True,
        key="kuliah"
    )

    if files:
        st.success(f"{len(files)} file berhasil dipilih.")

        for file in files:
            st.write(
                f"📄 **{file.name}** — "
                f"{format_ukuran(file.size)}"
            )

    st.info(
        "Pada tahap AI, aplikasi akan membaca instruksi dosen, "
        "membandingkannya dengan tugas Anda, menunjukkan bagian "
        "yang sudah sesuai dan bagian yang masih perlu diperbaiki."
    )


# ============================================================
# ANALISIS KARYA AKADEMIK
# ============================================================
elif menu == "🔬   # PERSIAPAN AI":

    st.header("🔬 # SIMPAN DAN EKSPOR")

    st.caption(
        "Analisis satu atau banyak karya akademik. "
        "PDF, DOCX, dan TXT dapat dibaca langsung."
    )

    # --------------------------------------------------------
    # PILIH MODE
    # --------------------------------------------------------
    mode = st.radio(
        "Mode Analisis",
        [
            "Analisis 1 Dokumen",
            "Analisis Banyak Dokumen"
        ],
        horizontal=True,
        key="mode_analisis_karya"
    )

    banyak = mode == "Analisis Banyak Dokumen"

    # --------------------------------------------------------
    # PILIH JENIS KARYA
    # --------------------------------------------------------
    jenis = st.selectbox(
        "Jenis karya",
        [
            "Deteksi Otomatis",
            "Tugas Kuliah",
            "Makalah",
            "Artikel Jurnal",
            "Buku",
            "Bab Buku",
            "Proposal",
            "Tesis",
            "Disertasi",
            "Laporan Penelitian",
            "Regulasi",
            "Dokumen Lain"
        ],
        key="jenis_karya_analisis"
    )

    # --------------------------------------------------------
    # UPLOAD
    # --------------------------------------------------------
    files = st.file_uploader(
        "Unggah dokumen",
        type=["pdf", "docx", "txt"],
        accept_multiple_files=banyak,
        key="upload_analisis_karya"
    )

    daftar_file = []

    if files:
        daftar_file = files if banyak else [files]

    # ========================================================
    # JIKA BELUM ADA DOKUMEN
    # ========================================================
    if not daftar_file:

        st.info(
            "Unggah minimal satu dokumen untuk memulai analisis."
        )

        st.markdown("### 🧠 Kemampuan Analisis")

        st.write(
            """
            Sistem ini disiapkan untuk menganalisis:

            - tugas kuliah dan makalah
            - artikel jurnal
            - buku dan bab buku
            - proposal penelitian
            - tesis
            - disertasi
            - laporan penelitian
            - regulasi dan dokumen akademik lainnya
            """
        )

        st.markdown("### 🔜 Mesin AI")

        st.caption(
            "Analisis semantik penuh akan menggunakan 9Router. "
            "Struktur aplikasi dan alur analisis disiapkan terlebih dahulu."
        )

    # ========================================================
    # JIKA DOKUMEN SUDAH DIUNGGAH
    # ========================================================
    else:

        hasil_dokumen = []

        # ----------------------------------------------------
        # BACA SEMUA DOKUMEN
        # ----------------------------------------------------
        for nomor, file in enumerate(daftar_file, start=1):

            st.divider()
            st.subheader(f"📄 Dokumen {nomor}: {file.name}")

            teks = ekstrak_teks(file)

            if teks.startswith("ERROR:"):
                st.error(teks)
                continue

            if not teks.strip():
                st.warning(
                    "Teks tidak berhasil dibaca dari dokumen ini."
                )
                continue

            jumlah_kata = len(teks.split())
            jumlah_karakter = len(teks)

            col1, col2, col3 = st.columns(3)

            col1.metric(
                "Jumlah Kata",
                jumlah_kata
            )

            col2.metric(
                "Jumlah Karakter",
                jumlah_karakter
            )

            col3.metric(
                "Ukuran File",
                format_ukuran(file.size)
            )

            # -----------------------------------------------
            # DETEKSI STRUKTUR
            # -----------------------------------------------
            struktur = deteksi_struktur(teks)

            st.markdown("#### 🔎 Deteksi Struktur Dasar")

            for bagian, ada in struktur.items():

                if ada:
                    st.success(f"✓ {bagian}")
                else:
                    st.caption(
                        f"○ {bagian} belum terdeteksi secara otomatis"
                    )

            # -----------------------------------------------
            # TEKS ASLI
            # -----------------------------------------------
            with st.expander("📖 Lihat teks dokumen"):

                st.text_area(
                    "Teks hasil ekstraksi",
                    teks,
                    height=350,
                    key=f"teks_dokumen_{nomor}"
                )

            # -----------------------------------------------
            # SIMPAN KE BANK KARYA
            # -----------------------------------------------
            if st.button(
                "💾 Simpan ke Bank Karya",
                key=f"simpan_bank_karya_{nomor}"
            ):

                berhasil = simpan_karya(
                    file.name,
                    jenis,
                    format_ukuran(file.size)
                )

                if berhasil:
                    st.success(
                        "Dokumen berhasil disimpan ke Bank Karya."
                    )
                else:
                    st.warning(
                        "Dokumen ini sudah tercatat pada proyek aktif."
                    )

            hasil_dokumen.append(
                {
                    "nama": file.name,
                    "teks": teks,
                    "kata": jumlah_kata,
                    "karakter": jumlah_karakter,
                    "ukuran": format_ukuran(file.size),
                    "struktur": struktur
                }
            )

        # ====================================================
        # ANALISIS AKADEMIK CERDAS
        # ====================================================
        if hasil_dokumen:

            st.divider()

            st.header("🧠 Analisis Akademik Cerdas")

            st.caption(
                "Kerangka analisis menyesuaikan jenis karya. "
                "Mesin AI 9Router akan dihubungkan pada tahap integrasi AI."
            )

            # ------------------------------------------------
            # FOKUS ANALISIS
            # ------------------------------------------------
            st.markdown("### 🎯 Fokus Analisis")

            fokus_analisis = st.multiselect(
                "Pilih bagian yang ingin dianalisis",
                [
                    "Identitas Dokumen",
                    "Latar Belakang / Masalah",
                    "Rumusan Masalah",
                    "Tujuan Penelitian",
                    "Teori / Konsep Utama",
                    "Penelitian Terdahulu",
                    "Metodologi",
                    "Populasi / Sampel / Informan",
                    "Instrumen Penelitian",
                    "Teknik Pengumpulan Data",
                    "Teknik Analisis Data",
                    "Temuan / Hasil",
                    "Pembahasan",
                    "Kesimpulan",
                    "Keterbatasan",
                    "Research Gap",
                    "Novelty / Kebaruan",
                    "Kontribusi Penelitian",
                    "Relevansi dengan Penelitian Saya"
                ],
                default=[
                    "Latar Belakang / Masalah",
                    "Tujuan Penelitian",
                    "Metodologi",
                    "Temuan / Hasil",
                    "Keterbatasan",
                    "Research Gap",
                    "Novelty / Kebaruan"
                ],
                key="fokus_analisis_akademik"
            )

            # ------------------------------------------------
            # KERANGKA BERDASARKAN JENIS
            # ------------------------------------------------
            st.markdown("### 📋 Kerangka Analisis")

            if jenis == "Artikel Jurnal":

                kerangka = [
                    "Identitas artikel",
                    "Topik penelitian",
                    "Masalah penelitian",
                    "Tujuan penelitian",
                    "Teori atau konsep utama",
                    "Penelitian terdahulu",
                    "Metode penelitian",
                    "Populasi, sampel, atau informan",
                    "Instrumen penelitian",
                    "Teknik pengumpulan data",
                    "Teknik analisis data",
                    "Temuan utama",
                    "Pembahasan",
                    "Kesimpulan",
                    "Keterbatasan penelitian",
                    "Research gap",
                    "Novelty atau kebaruan",
                    "Kontribusi penelitian",
                    "Peluang penelitian lanjutan",
                    "Relevansi dengan penelitian pengguna"
                ]

            elif jenis in ["Buku", "Bab Buku"]:

                kerangka = [
                    "Identitas buku",
                    "Pokok bahasan",
                    "Gagasan utama",
                    "Konsep atau teori penting",
                    "Argumentasi penulis",
                    "Bagian penting",
                    "Kekuatan pembahasan",
                    "Keterbatasan pembahasan",
                    "Relevansi dengan penelitian pengguna"
                ]

            elif jenis == "Proposal":

                kerangka = [
                    "Judul penelitian",
                    "Latar belakang",
                    "Identifikasi masalah",
                    "Research gap",
                    "Rumusan masalah",
                    "Tujuan penelitian",
                    "Manfaat penelitian",
                    "Kajian teori",
                    "Penelitian terdahulu",
                    "Kerangka berpikir",
                    "Hipotesis atau fokus penelitian",
                    "Metodologi",
                    "Populasi, sampel, atau informan",
                    "Instrumen penelitian",
                    "Teknik pengumpulan data",
                    "Teknik analisis data",
                    "Kelayakan rancangan penelitian"
                ]

            elif jenis == "Tesis":

                kerangka = [
                    "Identitas tesis",
                    "Judul penelitian",
                    "Latar belakang",
                    "Masalah penelitian",
                    "Rumusan masalah",
                    "Tujuan penelitian",
                    "Teori utama",
                    "Penelitian terdahulu",
                    "Research gap",
                    "Kerangka berpikir",
                    "Hipotesis atau fokus penelitian",
                    "Metodologi",
                    "Populasi, sampel, atau informan",
                    "Instrumen penelitian",
                    "Teknik analisis data",
                    "Hasil penelitian",
                    "Pembahasan",
                    "Kesimpulan",
                    "Keterbatasan",
                    "Novelty",
                    "Kontribusi penelitian",
                    "Peluang penelitian lanjutan"
                ]

            elif jenis == "Disertasi":

                kerangka = [
                    "Identitas disertasi",
                    "Judul penelitian",
                    "Latar belakang",
                    "Masalah penelitian",
                    "Rumusan masalah",
                    "Tujuan penelitian",
                    "State of the Art",
                    "Landasan teori",
                    "Penelitian terdahulu",
                    "Research gap",
                    "Kerangka konseptual",
                    "Metodologi",
                    "Populasi, sampel, atau informan",
                    "Instrumen penelitian",
                    "Teknik analisis data",
                    "Temuan utama",
                    "Pembahasan",
                    "Originalitas yang diklaim",
                    "Novelty",
                    "Kontribusi teoretis",
                    "Kontribusi metodologis",
                    "Kontribusi praktis",
                    "Keterbatasan",
                    "Peluang penelitian doktoral lanjutan"
                ]

            else:

                kerangka = [
                    "Identitas dokumen",
                    "Topik utama",
                    "Masalah utama",
                    "Tujuan",
                    "Konsep penting",
                    "Metode atau pendekatan",
                    "Temuan atau gagasan utama",
                    "Kesimpulan",
                    "Keterbatasan",
                    "Relevansi"
                ]

            for no, item in enumerate(kerangka, start=1):
                st.write(f"{no}. {item}")

                       # =================================================
            # MESIN ANALISIS AI 9ROUTER
            # =================================================
            st.markdown("### 🤖 Mesin Analisis AI")

            st.info(
                "Dokumen sudah berhasil dibaca. AI 9Router siap "
                "menganalisis isi dokumen berdasarkan fokus yang dipilih."
            )

            if "hasil_ai_9router" not in st.session_state:
                st.session_state.hasil_ai_9router = ""

            if st.button(
                "🤖 Analisis dengan AI",
                type="primary",
                use_container_width=True,
                key="tombol_analisis_9router"
            ):
                if not hasil_dokumen:
                    st.warning(
                        "Belum ada dokumen yang berhasil dibaca."
                    )
                else:
                    dokumen_ai = hasil_dokumen[0]
                    teks_ai = dokumen_ai["teks"]

                    with st.spinner(
                        "9Router sedang membaca dan menganalisis dokumen..."
                    ):
                        hasil_ai = analisis_dengan_9router(
                            teks_ai,
                            jenis,
                            fokus_analisis
                        )

                    st.session_state.hasil_ai_9router = hasil_ai

            # =================================================
            # HASIL ANALISIS AI
            # =================================================
            st.markdown("### 📝 Hasil Analisis")

            if st.session_state.hasil_ai_9router:

                st.success("✅ Analisis AI selesai.")

                hasil_edit = st.text_area(
                    "Hasil analisis dapat diedit sebelum diekspor",
                    value=st.session_state.hasil_ai_9router,
                    height=600,
                    key="editor_hasil_ai_9router"
                )

                st.session_state.hasil_ai_9router = hasil_edit

            else:
                st.text_area(
                    "Hasil analisis AI akan tampil di sini",
                    value="",
                    height=300,
                    disabled=True,
                    key="hasil_ai_kosong"
                )
                        # =================================================
            # SIMPAN DAN EKSPOR HASIL AI
            # =================================================
            st.markdown("### 💾 Simpan & Ekspor")

            hasil_final = st.session_state.get(
                "hasil_ai_9router",
                ""
            )

            if not hasil_final:
                st.info(
                    "Jalankan Analisis dengan AI terlebih dahulu. "
                    "Setelah hasil tersedia, tombol simpan dan ekspor "
                    "akan aktif."
                )

            # =================================================
            # SIMPAN HASIL KE PROYEK
            # =================================================
            col_simpan, col_word = st.columns(2)

            with col_simpan:

                if st.button(
                    "💾 Simpan Hasil ke Proyek",
                    disabled=not bool(hasil_final),
                    use_container_width=True,
                    key="simpan_hasil_proyek"
                ):
                    if "hasil_proyek" not in st.session_state:
                        st.session_state.hasil_proyek = []

                    data_hasil = {
                        "Proyek": st.session_state.proyek_aktif,
                        "Jenis": jenis,
                        "Mode": mode,
                        "Fokus": fokus_analisis,
                        "Hasil": hasil_final,
                        "Tanggal": datetime.now().strftime(
                            "%d-%m-%Y %H:%M"
                        )
                    }

                    st.session_state.hasil_proyek.append(
                        data_hasil
                    )

                    st.success(
                        "✅ Hasil analisis berhasil disimpan "
                        "ke proyek aktif."
                    )

            # =================================================
            # EKSPOR WORD
            # =================================================
            with col_word:

                try:
                    from io import BytesIO
                    from docx import Document

                    dokumen_word = Document()

                    dokumen_word.add_heading(
                        "Analisis Karya Akademik",
                        level=1
                    )

                    dokumen_word.add_paragraph(
                        f"Proyek: {st.session_state.proyek_aktif}"
                    )

                    dokumen_word.add_paragraph(
                        f"Jenis karya: {jenis}"
                    )

                    dokumen_word.add_paragraph(
                        f"Mode analisis: {mode}"
                    )

                    dokumen_word.add_heading(
                        "Dokumen yang Dianalisis",
                        level=2
                    )

                    for item in hasil_dokumen:
                        dokumen_word.add_paragraph(
                            f"Nama file: {item['nama']}"
                        )
                        dokumen_word.add_paragraph(
                            f"Jumlah kata: {item['kata']}"
                        )
                        dokumen_word.add_paragraph(
                            f"Ukuran: {item['ukuran']}"
                        )

                    dokumen_word.add_heading(
                        "Fokus Analisis",
                        level=2
                    )

                    for fokus in fokus_analisis:
                        dokumen_word.add_paragraph(
                            fokus,
                            style="List Bullet"
                        )

                    dokumen_word.add_heading(
                        "Hasil Analisis AI",
                        level=2
                    )

                    dokumen_word.add_paragraph(
                        hasil_final if hasil_final
                        else "Belum ada hasil analisis AI."
                    )

                    buffer_word = BytesIO()
                    dokumen_word.save(buffer_word)
                    buffer_word.seek(0)

                    st.download_button(
                        "📄 Ekspor Word",
                        data=buffer_word.getvalue(),
                        file_name="analisis_karya_akademik.docx",
                        mime=(
                            "application/vnd.openxmlformats-"
                            "officedocument.wordprocessingml.document"
                        ),
                        disabled=not bool(hasil_final),
                        use_container_width=True,
                        key="download_analisis_word"
                    )

                except Exception as e:
                    st.warning(
                        f"Ekspor Word belum dapat dibuat: {e}"
                    )

            # =================================================
            # EKSPOR PDF
            # =================================================
            col_pdf, col_csv = st.columns(2)

            with col_pdf:

                try:
                    from io import BytesIO
                    from reportlab.lib.pagesizes import A4
                    from reportlab.lib.styles import getSampleStyleSheet
                    from reportlab.platypus import (
                        SimpleDocTemplate,
                        Paragraph,
                        Spacer
                    )

                    buffer_pdf = BytesIO()

                    pdf = SimpleDocTemplate(
                        buffer_pdf,
                        pagesize=A4,
                        rightMargin=50,
                        leftMargin=50,
                        topMargin=50,
                        bottomMargin=50
                    )

                    styles = getSampleStyleSheet()
                    isi_pdf = []

                    isi_pdf.append(
                        Paragraph(
                            "Analisis Karya Akademik",
                            styles["Title"]
                        )
                    )

                    isi_pdf.append(Spacer(1, 12))

                    isi_pdf.append(
                        Paragraph(
                            f"Jenis karya: {jenis}",
                            styles["Normal"]
                        )
                    )

                    isi_pdf.append(
                        Paragraph(
                            f"Mode analisis: {mode}",
                            styles["Normal"]
                        )
                    )

                    isi_pdf.append(Spacer(1, 12))

                    isi_pdf.append(
                        Paragraph(
                            "Hasil Analisis AI",
                            styles["Heading2"]
                        )
                    )

                    if hasil_final:
                        for paragraf in hasil_final.split("\n"):
                            if paragraf.strip():
                                isi_pdf.append(
                                    Paragraph(
                                        paragraf.replace(
                                            "&", "&amp;"
                                        ).replace(
                                            "<", "&lt;"
                                        ).replace(
                                            ">", "&gt;"
                                        ),
                                        styles["Normal"]
                                    )
                                )
                                isi_pdf.append(
                                    Spacer(1, 6)
                                )

                    pdf.build(isi_pdf)
                    buffer_pdf.seek(0)

                    st.download_button(
                        "📕 Ekspor PDF",
                        data=buffer_pdf.getvalue(),
                        file_name="analisis_karya_akademik.pdf",
                        mime="application/pdf",
                        disabled=not bool(hasil_final),
                        use_container_width=True,
                        key="download_analisis_pdf"
                    )

                except Exception:
                    st.info(
                        "Ekspor PDF memerlukan reportlab. "
                        "Jika tombol PDF belum tersedia, "
                        "kita aktifkan dependensinya."
                    )

            # =================================================
            # EKSPOR CSV
            # =================================================
            with col_csv:

                data_ekspor = []

                for item in hasil_dokumen:
                    data_ekspor.append(
                        {
                            "Nama File": item["nama"],
                            "Jenis Karya": jenis,
                            "Jumlah Kata": item["kata"],
                            "Ukuran": item["ukuran"],
                            "Fokus Analisis": "; ".join(
                                fokus_analisis
                            ),
                            "Hasil Analisis AI": hasil_final
                        }
                    )

                df_ekspor = pd.DataFrame(data_ekspor)

                st.download_button(
                    "📊 Ekspor CSV",
                    data=df_ekspor.to_csv(
                        index=False
                    ).encode("utf-8-sig"),
                    file_name="hasil_analisis_akademik.csv",
                    mime="text/csv",
                    disabled=not bool(hasil_final),
                    use_container_width=True,
                    key="download_analisis_csv"
                )

            # =================================================
            # EKSPOR EXCEL
            # =================================================
            try:
                from io import BytesIO

                buffer_excel = BytesIO()

                with pd.ExcelWriter(
                    buffer_excel,
                    engine="openpyxl"
                ) as writer:
                    df_ekspor.to_excel(
                        writer,
                        sheet_name="Hasil Analisis",
                        index=False
                    )

                buffer_excel.seek(0)

                st.download_button(
                    "📗 Ekspor Excel",
                    data=buffer_excel.getvalue(),
                    file_name="hasil_analisis_akademik.xlsx",
                    mime=(
                        "application/vnd.openxmlformats-officedocument."
                        "spreadsheetml.sheet"
                    ),
                    disabled=not bool(hasil_final),
                    use_container_width=True,
                    key="download_analisis_excel"
                )

            except Exception as e:
                st.info(
                    f"Ekspor Excel belum tersedia: {e}"
                )

            # =================================================
            # ANALISIS BANYAK DOKUMEN
            # =================================================
            if banyak:

                st.divider()
                st.header("📚 Matriks Literatur")

                data_matriks = []

                for item in hasil_dokumen:
                    data_matriks.append(
                        {
                            "Nama File": item["nama"],
                            "Penulis": "",
                            "Tahun": "",
                            "Judul": "",
                            "Masalah": "",
                            "Teori": "",
                            "Metode": "",
                            "Sampel / Informan": "",
                            "Instrumen": "",
                            "Analisis Data": "",
                            "Temuan": "",
                            "Keterbatasan": "",
                            "Research Gap": "",
                            "Novelty": "",
                            "Relevansi": "",
                            "DOI / URL": ""
                        }
                    )

                df_matriks = pd.DataFrame(
                    data_matriks
                )

                st.dataframe(
                    df_matriks,
                    use_container_width=True,
                    hide_index=True
                )

                st.download_button(
                    "📥 Ekspor Matriks CSV",
                    data=df_matriks.to_csv(
                        index=False
                    ).encode("utf-8-sig"),
                    file_name="matriks_literatur.csv",
                    mime="text/csv",
                    key="download_matriks_csv"
                )
# ============================================================
# TESIS S2
# ============================================================
elif menu == "🎓 Tesis S2":

    st.header("🎓 Asisten Tesis S2")

    tahap = st.selectbox(
        "Tahap Tesis",
        [
            "Ide & Topik",
            "Identifikasi Masalah",
            "Research Gap",
            "State of the Art",
            "Novelty",
            "Alternatif Judul",
            "Rumusan Masalah",
            "Tujuan Penelitian",
            "BAB I",
            "BAB II",
            "Kerangka Berpikir",
            "Hipotesis / Fokus Penelitian",
            "BAB III",
            "Instrumen",
            "Pengumpulan Data",
            "BAB IV",
            "BAB V",
            "Tesis Lengkap",
            "Persiapan Sidang"
        ]
    )

    st.subheader(tahap)

    st.text_area(
        "Tuliskan ide, masalah, atau kebutuhan Anda"
    )

    st.file_uploader(
        "Unggah bahan tesis",
        type=[
            "pdf", "docx", "txt",
            "csv", "xlsx",
            "jpg", "jpeg", "png"
        ],
        accept_multiple_files=True,
        key="tesis"
    )

    st.info(
        "Penulisan AI nantinya menggunakan alur "
        "outline → sumber → draf → sitasi → verifikasi → revisi."
    )


# ============================================================
# DISERTASI S3
# ============================================================
elif menu == "🧑‍🎓 Disertasi S3":

    st.header("🧑‍🎓 Asisten Disertasi S3")

    tahap = st.selectbox(
        "Tahap Disertasi",
        [
            "Jembatan Tesis S2 → S3",
            "Analisis Tesis S2",
            "Keterbatasan Penelitian S2",
            "Pertanyaan Penelitian Lanjutan",
            "Topik Doktoral",
            "State of the Art",
            "Research Gap",
            "Novelty Doktoral",
            "Kontribusi Teoretis",
            "Kontribusi Metodologis",
            "Kontribusi Praktis",
            "Proposal Disertasi",
            "Metodologi Doktoral",
            "Instrumen",
            "Pengumpulan Data",
            "Analisis Data",
            "Penulisan Disertasi",
            "Publikasi",
            "Persiapan Ujian Doktoral"
        ]
    )

    st.subheader(tahap)

    st.file_uploader(
        "Unggah tesis S2, artikel, jurnal atau bahan S3",
        type=[
            "pdf", "docx", "txt",
            "csv", "xlsx"
        ],
        accept_multiple_files=True,
        key="disertasi"
    )

    st.info(
        "Untuk S3, aplikasi nantinya tidak hanya mencari gap, "
        "tetapi membantu menelusuri dasar bukti untuk novelty "
        "dan kontribusi doktoral."
    )


# ============================================================
# METODOLOGI
# ============================================================
elif menu == "🧭 Metodologi Penelitian":

    st.header("🧭 Penentu Jenis & Metodologi Penelitian")

    jenjang = st.radio(
        "Jenjang",
        ["Tesis S2", "Disertasi S3"],
        horizontal=True
    )

    masalah = st.text_area(
        "Apa masalah utama yang ingin diteliti?"
    )

    tujuan = st.text_area(
        "Apa yang ingin diketahui, diuji, dipahami, "
        "atau dikembangkan?"
    )

    data = st.multiselect(
        "Data yang kemungkinan digunakan",
        [
            "Angka / skor",
            "Angket",
            "Wawancara",
            "Observasi",
            "Dokumen",
            "Eksperimen",
            "Produk / model",
            "Literatur",
            "Gabungan kuantitatif dan kualitatif"
        ]
    )

    hasil = st.selectbox(
        "Hasil utama yang diharapkan",
        [
            "Pilih...",
            "Menggambarkan fenomena",
            "Mengetahui hubungan",
            "Mengetahui perbedaan",
            "Menguji pengaruh",
            "Memahami pengalaman/fenomena",
            "Mengembangkan produk",
            "Mengembangkan model",
            "Mengembangkan teori",
            "Menggabungkan kuantitatif dan kualitatif",
            "Mengkaji literatur secara sistematis"
        ]
    )

    if st.button("🔍 Analisis Alternatif Metodologi"):

        if not masalah.strip():
            st.warning(
                "Tuliskan masalah penelitian terlebih dahulu."
            )

        elif hasil == "Pilih...":
            st.warning(
                "Pilih hasil utama yang diharapkan."
            )

        else:
            st.subheader("Alternatif Awal")

            if hasil == "Mengetahui hubungan":
                st.success(
                    "Kuantitatif korelasional dapat dipertimbangkan."
                )

            elif hasil == "Mengetahui perbedaan":
                st.success(
                    "Kuantitatif komparatif atau desain eksperimen "
                    "dapat dipertimbangkan sesuai masalah."
                )

            elif hasil == "Menguji pengaruh":
                st.success(
                    "Kuantitatif eksplanatori/regresi, eksperimen "
                    "atau quasi eksperimen dapat dipertimbangkan."
                )

            elif hasil == "Memahami pengalaman/fenomena":
                st.success(
                    "Pendekatan kualitatif dapat dipertimbangkan, "
                    "misalnya fenomenologi atau studi kasus "
                    "sesuai pertanyaan penelitian."
                )

            elif hasil == "Mengembangkan produk":
                st.success(
                    "Research & Development dapat dipertimbangkan."
                )

            elif hasil in [
                "Mengembangkan model",
                "Mengembangkan teori"
            ]:
                st.success(
                    "Pengembangan model, grounded theory, "
                    "mixed methods atau desain lain dapat "
                    "dipertimbangkan sesuai tujuan."
                )

            elif hasil == (
                "Menggabungkan kuantitatif dan kualitatif"
            ):
                st.success(
                    "Mixed Methods dapat dipertimbangkan."
                )

            elif hasil == (
                "Mengkaji literatur secara sistematis"
            ):
                st.success(
                    "Systematic Literature Review atau "
                    "penelitian kepustakaan dapat dipertimbangkan."
                )

            else:
                st.success(
                    "Pendekatan deskriptif dapat dipertimbangkan."
                )

            st.caption(
                "Ini masih rekomendasi awal. Pada tahap AI, "
                "aplikasi akan membandingkan beberapa alternatif "
                "beserta alasan, kebutuhan data, instrumen, "
                "analisis dan tingkat kelayakannya."
            )


# ============================================================
# INSTRUMEN
# ============================================================
elif menu == "📝 Instrumen Penelitian":

    st.header("📝 Instrumen Penelitian")

    st.selectbox(
        "Jenis Instrumen",
        [
            "Angket/Kuesioner",
            "Pedoman Wawancara",
            "Lembar Observasi",
            "Dokumentasi",
            "Tes",
            "Rubrik",
            "Instrumen R&D"
        ]
    )

    st.write(
        """
        Alur yang akan dikembangkan:

        **Konstruk → Dimensi → Indikator → Butir →
        Sumber Teori → Validasi Ahli → Uji Coba →
        Validitas → Reliabilitas**
        """
    )


# ============================================================
# STATISTIK
# ============================================================
elif menu == "📊 Statistik & SPSS":

    st.header("📊 Laboratorium Statistik & SPSS")

    analisis = st.selectbox(
        "Pilih Analisis",
        [
            "Asisten Pemilihan Uji Statistik",
            "Data Cleaning",
            "Missing Data",
            "Outlier",
            "Statistik Deskriptif",
            "Uji Validitas",
            "Uji Reliabilitas",
            "Uji Normalitas",
            "Uji Homogenitas",
            "Uji Linearitas",
            "Korelasi Pearson",
            "Korelasi Spearman",
            "Uji t",
            "ANOVA",
            "Chi-Square",
            "Regresi Linear Sederhana",
            "Regresi Linear Berganda",
            "Multikolinearitas",
            "Heteroskedastisitas",
            "Uji Nonparametrik",
            "Effect Size",
            "Confidence Interval",
            "Interpretasi Output SPSS",
            "Mediasi / Moderasi",
            "Analisis Faktor / SEM"
        ]
    )

    st.file_uploader(
        "Unggah data / output",
        type=["csv", "xlsx", "pdf", "docx"],
        key="statistik"
    )

    st.info(
        f"Modul **{analisis}** akan dikembangkan bertahap. "
        "Aplikasi nantinya juga menunjukkan langkah ekuivalen "
        "di SPSS dan membantu membuat narasi BAB IV."
    )


# ============================================================
# KUALITATIF
# ============================================================
elif menu == "🔤 Analisis Kualitatif":

    st.header("🔤 Laboratorium Analisis Kualitatif")

    st.write(
        """
        **Transkripsi → Coding → Codebook → Kategori →
        Tema → Kutipan Bukti → Triangulasi →
        Temuan → Pembahasan**
        """
    )

    st.file_uploader(
        "Unggah transkrip atau dokumen penelitian",
        type=["pdf", "docx", "txt"],
        accept_multiple_files=True,
        key="kualitatif"
    )


# ============================================================
# AUDIO VIDEO
# ============================================================
elif menu == "🎤 Audio & Video":

    st.header("🎤 Audio & Video Research Lab")

    files = st.file_uploader(
        "Unggah audio/video",
        type=[
            "mp3", "wav", "m4a",
            "mp4", "mov"
        ],
        accept_multiple_files=True
    )

    if files:
        for file in files:

            st.write(f"**{file.name}**")

            if file.name.lower().endswith(
                (".mp3", ".wav", ".m4a")
            ):
                st.audio(file)

            elif file.name.lower().endswith(
                (".mp4", ".mov")
            ):
                st.video(file)

    st.info(
        "Transkripsi, identifikasi pembicara, timestamp, coding "
        "dan analisis wawancara akan dipasang pada tahap AI."
    )


# ============================================================
# PENULISAN
# ============================================================
elif menu == "✍️ Penulisan Akademik":

    st.header("✍️ Asisten Penulisan Akademik")

    st.selectbox(
        "Kebutuhan",
        [
            "Membuat Outline",
            "Mengembangkan Paragraf",
            "Parafrase Akademik",
            "Sintesis Literatur",
            "Cari Bukti untuk Kalimat",
            "Periksa Klaim Tanpa Sumber",
            "Sitasi dalam Teks",
            "Daftar Pustaka",
            "Abstrak",
            "Ringkasan Akademik"
        ]
    )

    st.text_area(
        "Masukkan teks / gagasan"
    )


# ============================================================
# BIMBINGAN
# ============================================================
elif menu == "👨‍🏫 Bimbingan & Revisi":

    st.header("👨‍🏫 Bimbingan & Revisi")

    st.text_area(
        "Masukkan catatan pembimbing / promotor"
    )

    st.file_uploader(
        "Unggah dokumen revisi",
        type=["pdf", "docx"],
        accept_multiple_files=True
    )

    st.selectbox(
        "Status",
        [
            "Belum Dikerjakan",
            "Sedang Dikerjakan",
            "Selesai"
        ]
    )

    st.info(
        "Riwayat versi akan dikembangkan agar naskah sebelum "
        "dan sesudah revisi tetap dapat dilacak."
    )


# ============================================================
# PUBLIKASI
# ============================================================
elif menu == "📑 Publikasi Jurnal":

    st.header("📑 Asisten Publikasi Jurnal")

    st.selectbox(
        "Tahap Publikasi",
        [
            "Ubah Tesis menjadi Artikel",
            "Pilih Temuan Utama",
            "Struktur IMRaD",
            "Abstrak",
            "Tabel & Gambar",
            "Referensi",
            "Cari Jurnal yang Sesuai",
            "Checklist Submission",
            "Cover Letter",
            "Revisi Reviewer"
        ]
    )

    st.file_uploader(
        "Unggah tesis / artikel",
        type=["pdf", "docx"]
    )


# ============================================================
# PERPUSTAKAAN
# ============================================================
elif menu == "📂 Perpustakaan Akademik":

    st.header("📂 Perpustakaan Akademik Pribadi")

    tab1, tab2 = st.tabs(
        [
            "📁 Bank Karya",
            "📚 Bank Referensi"
        ]
    )

    with tab1:

        if st.session_state.bank_karya:

            df_karya = pd.DataFrame(
                st.session_state.bank_karya
            )

            st.dataframe(
                df_karya,
                use_container_width=True
            )

        else:
            st.info("Bank Karya masih kosong.")

    with tab2:

        if st.session_state.bank_referensi:

            df_ref = pd.DataFrame(
                st.session_state.bank_referensi
            )

            st.dataframe(
                df_ref,
                use_container_width=True
            )

        else:
            st.info("Bank Referensi masih kosong.")

    st.warning(
        "Pada Tahap 1 data ini masih tersimpan selama sesi. "
        "Database permanen nanti memisahkan data berdasarkan "
        "User ID → Project ID → File/Reference ID."
    )


# ============================================================
# AUDIT
# ============================================================
elif menu == "✅ Audit Akademik":

    st.header("✅ Audit Akademik")

    st.selectbox(
        "Jenis Audit",
        [
            "Audit Lengkap",
            "Konsistensi Judul",
            "Rumusan Masalah ↔ Tujuan",
            "Teori ↔ Variabel/Fokus",
            "Metode ↔ Instrumen",
            "Instrumen ↔ Data",
            "Temuan ↔ Kesimpulan",
            "Sitasi ↔ Daftar Pustaka",
            "Validasi Referensi",
            "Klaim Tanpa Sumber",
            "Research Gap",
            "Novelty",
            "Kesiapan Tesis",
            "Kesiapan Disertasi"
        ]
    )

    st.file_uploader(
        "Unggah naskah",
        type=["pdf", "docx"],
        key="audit"
    )

    st.write(
        """
        Rantai konsistensi utama:

        **Judul → Masalah → Rumusan Masalah → Tujuan →
        Teori → Metode → Instrumen → Data →
        Temuan → Kesimpulan**
        """
    )


# ============================================================
# PROGRES
# ============================================================
elif menu == "📈 Progres Penelitian":

    st.header("📈 Dashboard Progres Penelitian")

    st.write(
        """
        Dashboard permanen nantinya menampilkan:

        **Judul → Proposal → Seminar → Instrumen →
        Pengumpulan Data → Analisis → BAB IV →
        BAB V → Publikasi → Sidang**
        """
    )

    st.progress(0)

    st.caption(
        "Progres aktual akan dihitung dari proyek pengguna "
        "setelah database dipasang."
    )


# ============================================================
# PRESENTASI
# ============================================================
elif menu == "🖥️ Presentasi":

    st.header("🖥️ Asisten Presentasi Akademik")

    st.selectbox(
        "Jenis Presentasi",
        [
            "Tugas Kuliah",
            "Presentasi Artikel",
            "Seminar Proposal",
            "Seminar Hasil",
            "Sidang Tesis",
            "Proposal Disertasi",
            "Ujian Disertasi"
        ]
    )

    st.file_uploader(
        "Unggah sumber presentasi",
        type=["pdf", "docx", "pptx"],
        key="presentasi"
    )

    st.write(
        """
        Nantinya aplikasi membantu membuat:

        • Struktur slide  
        • Isi slide  
        • Naskah presentasi  
        • Catatan pembicara  
        • Ringkasan waktu  
        • Prediksi pertanyaan
        """
    )


# ============================================================
# SIMULASI SIDANG
# ============================================================
elif menu == "🎓 Simulasi Sidang":

    st.header("🎓 Simulasi Seminar & Sidang")

    st.selectbox(
        "Jenis Ujian",
        [
            "Seminar Proposal",
            "Seminar Hasil",
            "Sidang Tesis S2",
            "Ujian Proposal Disertasi",
            "Ujian Disertasi / Doktoral"
        ]
    )

    st.multiselect(
        "Mode Penguji",
        [
            "Ketua Sidang",
            "Penguji Substansi",
            "Penguji Teori",
            "Penguji Metodologi",
            "Penguji Statistik",
            "Penguji Referensi",
            "Penguji Novelty",
            "Penguji Kritis"
        ]
    )

    st.select_slider(
        "Tingkat",
        options=[
            "Mudah",
            "Sedang",
            "Kritis",
            "Sangat Kritis"
        ]
    )

    st.file_uploader(
        "Unggah naskah ujian",
        type=["pdf", "docx"],
        key="sidang"
    )

    st.write(
        """
        Fitur AI selanjutnya:

        **Prediksi Pertanyaan → Pertanyaan Satu per Satu →
        Jawaban Pengguna → Analisis Jawaban →
        Pertanyaan Lanjutan → Catatan Perbaikan →
        Laporan Latihan Sidang**

        Termasuk pencarian bagian naskah yang berpotensi
        mendapat pertanyaan sulit dari penguji.
        """
    )


# ============================================================
# DONASI
# ============================================================
elif menu == "💚 Donasi & Akses":

    st.header("💚 Donasi & Akses")

    st.write(
        """
        Rancangan akses:

        **Daftar → Verifikasi → Donasi → Aktivasi →
        Masa Aktif → Kredit AI → Perpanjangan Kredit**
        """
    )

    c1, c2, c3 = st.columns(3)

    c1.metric("Status Akun", "Tahap Pengembangan")
    c2.metric("Masa Aktif", "-")
    c3.metric("Kredit AI", "-")

    st.info(
        "Pembayaran/QRIS belum diaktifkan pada Tahap 1. "
        "Rahasia pembayaran dan API nantinya tidak ditempatkan "
        "di kode yang dapat dilihat pengguna."
    )


# ============================================================
# ADMIN
# ============================================================
elif menu == "⚙️ Admin":

    st.header("⚙️ Admin & Sistem")

    st.write(
        """
        Dashboard admin nantinya mengelola:

        • Pengguna  
        • Verifikasi akun  
        • Aktivasi akses  
        • Donasi  
        • Masa aktif  
        • Kredit AI  
        • Perpanjangan kredit  
        • Fitur aplikasi  
        • Log penggunaan  
        • Backup  
        • Keamanan
        """
    )

    st.warning(
        "Menu Admin pada Tahap 1 belum memiliki autentikasi. "
        "Jangan memasukkan data pembayaran atau API key di sini."
    )


# ============================================================
# FOOTER
# ============================================================
st.divider()

st.caption(
    "Asisten Akademik AI — fondasi pengembangan "
    "multi-user untuk perkuliahan, penelitian, tesis, "
    "disertasi, publikasi dan sidang."
)
