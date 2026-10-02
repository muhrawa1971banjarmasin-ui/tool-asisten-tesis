import streamlit as st
import re
from io import BytesIO


import PyPDF2

try:
    import docx
except ImportError:
    docx = None


# =========================
# KONFIGURASI APLIKASI
# =========================
st.set_page_config(
    page_title="Asisten Tesis & Pengelola Referensi",
    page_icon="🎓",
    layout="wide"
)

st.title("🎓 Asisten Tesis & Pengelola Referensi")
st.caption(
    "Membantu analisis dokumen, pemeriksaan sitasi, "
    "dan penyusunan referensi akademik."
)

st.divider()


# =========================
# FUNGSI MEMBACA DOKUMEN
# =========================
def baca_pdf(file):
    if pypdf is None:
        return "Library pypdf belum tersedia."

    reader = PyPDF2.PdfReader(file)
    teks = []

    for halaman in reader.pages:
        isi = halaman.extract_text()
        if isi:
            teks.append(isi)

    return "\n".join(teks)


def baca_docx(file):
    if docx is None:
        return "Library python-docx belum tersedia."

    dokumen = docx.Document(file)
    return "\n".join(
        paragraf.text
        for paragraf in dokumen.paragraphs
        if paragraf.text.strip()
    )


# =========================
# MENU UTAMA
# =========================
menu = st.sidebar.radio(
    "Pilih Menu",
    [
        "🏠 Beranda",
        "📄 Analisis Dokumen",
        "🔎 Pemeriksa Sitasi",
        "📚 Format Referensi"
    ]
)


# =========================
# BERANDA
# =========================
if menu == "🏠 Beranda":

    st.header("Selamat Datang")

    st.write(
        """
        Aplikasi ini dirancang untuk membantu mahasiswa dalam
        penyusunan proposal, skripsi, tesis, dan karya ilmiah.
        """
    )

    col1, col2, col3 = st.columns(3)

    with col1:
        st.subheader("📄 Analisis Dokumen")
        st.write(
            "Unggah PDF atau DOCX untuk membaca dan "
            "mengidentifikasi bagian penting dokumen."
        )

    with col2:
        st.subheader("🔎 Pemeriksa Sitasi")
        st.write(
            "Membantu memeriksa kesesuaian sitasi dalam teks "
            "dengan daftar pustaka."
        )

    with col3:
        st.subheader("📚 Format Referensi")
        st.write(
            "Membantu menyusun format referensi akademik "
            "secara lebih terstruktur."
        )


# =========================
# ANALISIS DOKUMEN
# =========================
elif menu == "📄 Analisis Dokumen":

    st.header("📄 Analisis Dokumen Tesis / Skripsi")

    file = st.file_uploader(
        "Unggah dokumen PDF atau DOCX",
        type=["pdf", "docx"]
    )

    if file is not None:

        try:
            if file.name.lower().endswith(".pdf"):
                teks = baca_pdf(file)
            else:
                teks = baca_docx(file)

            st.success("Dokumen berhasil dibaca.")

            jumlah_kata = len(teks.split())
            jumlah_karakter = len(teks)

            col1, col2 = st.columns(2)

            col1.metric("Jumlah Kata", jumlah_kata)
            col2.metric("Jumlah Karakter", jumlah_karakter)

            st.subheader("Isi Dokumen")

            st.text_area(
                "Teks hasil ekstraksi",
                teks,
                height=400
            )

            st.subheader("Identifikasi Struktur")

            bagian = {
                "Latar Belakang": [
                    "latar belakang",
                    "pendahuluan"
                ],
                "Metode Penelitian": [
                    "metode penelitian",
                    "metodologi penelitian"
                ],
                "Hasil Penelitian": [
                    "hasil penelitian",
                    "hasil dan pembahasan"
                ],
                "Keterbatasan Penelitian": [
                    "keterbatasan penelitian",
                    "keterbatasan"
                ]
            }

            teks_kecil = teks.lower()

            for nama, kata_kunci in bagian.items():
                ditemukan = any(
                    kata in teks_kecil
                    for kata in kata_kunci
                )

                if ditemukan:
                    st.success(f"✓ {nama} terdeteksi")
                else:
                    st.warning(f"⚠ {nama} belum terdeteksi")

        except Exception as e:
            st.error(f"Dokumen gagal dibaca: {e}")


# =========================
# PEMERIKSA SITASI
# =========================
elif menu == "🔎 Pemeriksa Sitasi":

    st.header("🔎 Pemeriksa Konsistensi Sitasi")

    st.write(
        "Tempel bagian naskah yang memuat sitasi "
        "dan daftar pustaka."
    )

    teks_sitasi = st.text_area(
        "Teks naskah",
        height=250,
        placeholder="Contoh: Menurut Sugiyono (2022)..."
    )

    daftar_pustaka = st.text_area(
        "Daftar pustaka",
        height=250,
        placeholder="Tempel daftar pustaka di sini..."
    )

    if st.button("Periksa Sitasi"):

        if not teks_sitasi.strip():
            st.warning("Masukkan teks naskah terlebih dahulu.")

        elif not daftar_pustaka.strip():
            st.warning("Masukkan daftar pustaka terlebih dahulu.")

        else:

            pola = r"\(([A-Z][A-Za-zÀ-ÿ'’\-]+),?\s*(\d{4})\)"

            hasil = re.findall(pola, teks_sitasi)

            if hasil:

                st.subheader("Sitasi yang Terdeteksi")

                for penulis, tahun in hasil:

                    sitasi = f"{penulis} ({tahun})"

                    if (
                        penulis.lower() in daftar_pustaka.lower()
                        and tahun in daftar_pustaka
                    ):
                        st.success(
                            f"✓ {sitasi} ditemukan dalam daftar pustaka"
                        )
                    else:
                        st.error(
                            f"✗ {sitasi} belum ditemukan dalam daftar pustaka"
                        )

            else:
                st.info(
                    "Belum ditemukan pola sitasi "
                    "(Nama, Tahun) pada teks."
                )


# =========================
# FORMAT REFERENSI
# =========================
elif menu == "📚 Format Referensi":

    st.header("📚 Penyusun Format Referensi")

    gaya = st.selectbox(
        "Pilih gaya sitasi",
        [
            "APA 7th",
            "Harvard",
            "MLA",
            "IEEE"
        ]
    )

    penulis = st.text_input(
        "Nama penulis",
        placeholder="Contoh: Ahmad Fauzi"
    )

    tahun = st.text_input(
        "Tahun terbit",
        placeholder="2026"
    )

    judul = st.text_input(
        "Judul buku / artikel"
    )

    penerbit = st.text_input(
        "Nama penerbit / jurnal"
    )

    if st.button("Buat Referensi"):

        if not penulis or not tahun or not judul:

            st.warning(
                "Nama penulis, tahun, dan judul harus diisi."
            )

        else:

            if gaya == "APA 7th":
                hasil = (
                    f"{penulis}. ({tahun}). "
                    f"*{judul}*. {penerbit}."
                )

            elif gaya == "Harvard":
                hasil = (
                    f"{penulis} ({tahun}) "
                    f"*{judul}*. {penerbit}."
                )

            elif gaya == "MLA":
                hasil = (
                    f'{penulis}. *{judul}*. '
                    f'{penerbit}, {tahun}.'
                )

            else:
                hasil = (
                    f'{penulis}, "{judul}," '
                    f'{penerbit}, {tahun}.'
                )

            st.subheader("Hasil Referensi")

            st.success(hasil)

            st.code(hasil, language=None)


st.divider()

st.caption(
    "Asisten Tesis & Pengelola Referensi | "
    "Versi Awal"
)
