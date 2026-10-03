# ============================================================
# KONEKSI AI GEMINI LANGSUNG
# ============================================================

def analisis_dengan_gemini(teks, jenis_karya, fokus_analisis):
    try:
        api_key = st.secrets["GEMINI_API_KEY"]
    except Exception:
        return {
            "sukses": False,
            "hasil": "",
            "error": "GEMINI_API_KEY belum ditemukan di Streamlit Secrets."
        }

    model_id = "gemini-3.8-flash"

    teks_dokumen = teks[:60000]

    fokus = (
        ", ".join(fokus_analisis)
        if fokus_analisis
        else "Analisis akademik menyeluruh"
    )

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
4. Bedakan novelty yang diklaim penulis dengan novelty yang benar-benar
   telah diverifikasi melalui literatur.
5. Pada tahap ini hanya analisis dokumen, bukan pembuktian novelty
   terhadap seluruh literatur ilmiah.
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
        "contents": [
            {
                "parts": [
                    {
                        "text": prompt
                    }
                ]
            }
        ],
        "generationConfig": {
            "temperature": 0.2
        }
    }

    url = (
        "https://generativelanguage.googleapis.com/v1beta/models/"
        + model_id
        + ":generateContent?key="
        + api_key
    )

    data = json.dumps(payload).encode("utf-8")

    request = urllib.request.Request(
        url,
        data=data,
        headers={
            "Content-Type": "application/json"
        },
        method="POST"
    )

    try:
        with urllib.request.urlopen(request, timeout=180) as response:
            hasil_data = json.loads(
                response.read().decode("utf-8")
            )

        candidates = hasil_data.get("candidates", [])

        if not candidates:
            return {
                "sukses": False,
                "hasil": "",
                "error": "Gemini tidak mengembalikan hasil analisis."
            }

        parts = (
            candidates[0]
            .get("content", {})
            .get("parts", [])
        )

        hasil_teks = "\n".join(
            part.get("text", "")
            for part in parts
            if part.get("text")
        ).strip()

        if not hasil_teks:
            return {
                "sukses": False,
                "hasil": "",
                "error": "Gemini merespons tetapi hasil analisis kosong."
            }

        return {
            "sukses": True,
            "hasil": hasil_teks,
            "model": model_id
        }

    except urllib.error.HTTPError as e:
        try:
            detail = e.read().decode("utf-8")
        except Exception:
            detail = ""

        return {
            "sukses": False,
            "hasil": "",
            "error": (
                f"Gemini belum berhasil memproses permintaan. "
                f"HTTP {e.code}. {detail}"
            )
        }

    except urllib.error.URLError as e:
        return {
            "sukses": False,
            "hasil": "",
            "error": f"Tidak dapat terhubung ke Gemini: {e.reason}"
        }

    except Exception as e:
        return {
            "sukses": False,
            "hasil": "",
            "error": f"Terjadi kesalahan saat menjalankan Gemini: {str(e)}"
        }
