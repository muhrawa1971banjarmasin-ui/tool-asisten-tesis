import streamlit as st
import pandas as pd
import PyPDF2
from datetime import datetime

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
elif menu == "🔬 Analisis Karya Akademik":

    st.header("🔬 Analisis Karya Akademik")

    mode = st.radio(
        "Mode Analisis",
        [
            "Analisis 1 Dokumen",
            "Analisis Banyak Dokumen"
        ],
        horizontal=True
    )

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
        ]
    )

    banyak = mode == "Analisis Banyak Dokumen"

    files = st.file_uploader(
        "Unggah dokumen",
        type=["pdf", "docx", "txt"],
        accept_multiple_files=banyak,
        key="analisis_karya"
    )

    daftar_file = []

    if files:
        daftar_file = files if banyak else [files]

    for nomor, file in enumerate(daftar_file, start=1):

        st.divider()
        st.subheader(f"Dokumen {nomor}: {file.name}")

        teks = ekstrak_teks(file)

        if teks.startswith("ERROR:"):
            st.error(teks)
            continue

        if teks.strip():

            kata = len(teks.split())
            karakter = len(teks)

            a, b, c = st.columns(3)
            a.metric("Jumlah Kata", kata)
            b.metric("Jumlah Karakter", karakter)
            c.metric("Ukuran File", format_ukuran(file.size))

            struktur = deteksi_struktur(teks)

            st.write("**Deteksi struktur dasar:**")

            for bagian, ada in struktur.items():
                if ada:
                    st.success(f"✓ {bagian}")
                else:
                    st.caption(f"○ {bagian} belum terdeteksi")

            with st.expander("Lihat teks dokumen"):
                st.text_area(
                    "Teks",
                    teks,
                    height=350,
                    key=f"teks_{nomor}"
                )

            if st.button(
                "💾 Simpan ke Bank Karya",
                key=f"simpan_{nomor}"
            ):
                berhasil = simpan_karya(
                    file.name,
                    jenis,
                    format_ukuran(file.size)
                )

                if berhasil:
                    st.success("Disimpan ke Bank Karya.")
                else:
                    st.warning(
                        "File ini sudah tercatat pada proyek aktif."
                    )

    if banyak:
        st.info(
            "Tahap berikutnya akan menambahkan matriks lintas dokumen: "
            "penulis, tahun, masalah, teori, metode, sampel/informan, "
            "instrumen, hasil, keterbatasan, gap dan relevansi."
        )


# ============================================================
# LITERATUR & REFERENSI
# ============================================================
elif menu == "🔎 Literatur & Referensi":

    st.header("🔎 Literatur & Referensi")

    tab1, tab2, tab3, tab4 = st.tabs(
        [
            "🔎 Pencarian",
            "📚 Bank Referensi",
            "💬 Bank Kutipan",
            "✅ Validasi"
        ]
    )

    with tab1:

        st.text_input(
            "Topik / judul / DOI / kata kunci"
        )

        st.multiselect(
            "Sumber yang dibutuhkan",
            [
                "Jurnal Indonesia",
                "Jurnal Internasional",
                "Buku",
                "Tesis",
                "Disertasi",
                "Prosiding",
                "Repository"
            ]
        )

        st.info(
            "Tahap berikutnya: pencarian sumber nyata dan "
            "verifikasi metadata. Referensi tidak akan dibuat-buat."
        )

    with tab2:

        st.subheader("Tambah Referensi Manual")

        judul = st.text_input(
            "Judul",
            key="ref_judul"
        )

        penulis = st.text_input(
            "Penulis",
            key="ref_penulis"
        )

        tahun = st.text_input(
            "Tahun",
            key="ref_tahun"
        )

        sumber = st.text_input(
            "Jurnal / Penerbit / Universitas",
            key="ref_sumber"
        )

        doi = st.text_input(
            "DOI / URL",
            key="ref_doi"
        )

        if st.button("➕ Tambahkan ke Bank Referensi"):

            if not judul.strip():
                st.warning("Judul referensi belum diisi.")

            else:
                st.session_state.bank_referensi.append(
                    {
                        "Proyek": st.session_state.proyek_aktif,
                        "Judul": judul,
                        "Penulis": penulis,
                        "Tahun": tahun,
                        "Sumber": sumber,
                        "DOI/URL": doi,
                        "Status": "Belum diverifikasi"
                    }
                )

                st.success("Referensi ditambahkan.")

        if st.session_state.bank_referensi:

            df_ref = pd.DataFrame(
                st.session_state.bank_referensi
            )

            st.dataframe(
                df_ref,
                use_container_width=True
            )

            csv = df_ref.to_csv(
                index=False
            ).encode("utf-8")

            st.download_button(
                "⬇️ Unduh Bank Referensi CSV",
                csv,
                "bank_referensi.csv",
                "text/csv"
            )

        else:
            st.info("Bank Referensi masih kosong.")

    with tab3:

        st.write(
            """
            Bank Kutipan nantinya menyimpan:

            **Kutipan → Halaman → Sumber → DOI/ISBN →
            Topik → Digunakan pada BAB mana**
            """
        )

    with tab4:

        st.write(
            """
            Status pemeriksaan referensi:

            • Terverifikasi  
            • Perlu diperiksa  
            • Tidak ditemukan  
            • Duplikat  
            • Metadata tidak lengkap  
            • Sitasi tanpa daftar pustaka  
            • Daftar pustaka tanpa sitasi
            """
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
