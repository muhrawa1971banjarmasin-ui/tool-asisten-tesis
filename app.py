from pathlib import Path
from copy import deepcopy

import streamlit as st
import shutil
import pandas as pd
import PyPDF2
from datetime import datetime
import json
import urllib.request
import urllib.error
import urllib.parse
import re
import base64
import html
import difflib
import csv
import time
import random
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
    "S1 • S2 • S3 • OBE • Riset • Referensi • Publikasi • Buku • Sidang"
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

if "naskah_aktif" not in st.session_state:
    st.session_state.naskah_aktif = ""

if "hasil_penulisan_ai" not in st.session_state:
    st.session_state.hasil_penulisan_ai = ""
if "mata_kuliah_tambahan" not in st.session_state:
    st.session_state.mata_kuliah_tambahan = []
if "naskah_upload_ref" not in st.session_state:
    st.session_state.naskah_upload_ref = ""

if "kredit_ai" not in st.session_state:
    st.session_state.kredit_ai = 10

if "riwayat_kredit" not in st.session_state:
    st.session_state.riwayat_kredit = []

if "sumber_online_user" not in st.session_state:
    st.session_state.sumber_online_user = []
if "gaya_sitasi" not in st.session_state:
    st.session_state.gaya_sitasi = "Chicago Notes & Bibliography"


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
# KONEKSI AI GEMINI LANGSUNG
# ============================================================

def _panggil_gemini_rest_aman(prompt, temperature=0.25, max_output_tokens=8192):
    """Panggilan Gemini REST dengan retry eksponensial dan fallback model stabil."""
    import time
    import random

    try:
        api_key = st.secrets["GEMINI_API_KEY"]
    except Exception:
        return {
            "sukses": False,
            "hasil": "",
            "error": "GEMINI_API_KEY belum ditemukan di Streamlit Secrets.",
            "model": ""
        }

    # Model utama tetap model yang sudah digunakan aplikasi.
    # Fallback hanya dipakai setelah gangguan sementara pada model utama.
    model_ids = ["gemini-3.8-flash", "gemini-3.5-flash-lite"]
    transient_codes = {408, 429, 500, 502, 503, 504}
    payload = {
        "contents": [{"parts": [{"text": prompt[:90000]}]}],
        "generationConfig": {
            "temperature": temperature,
            "maxOutputTokens": max_output_tokens
        }
    }
    error_terakhir = ""

    for indeks_model, model_id in enumerate(model_ids):
        url = (
            "https://generativelanguage.googleapis.com/v1beta/models/"
            + model_id + ":generateContent?key=" + api_key
        )

        # 4 percobaan per model: 1 panggilan awal + 3 retry.
        for percobaan in range(4):
            req = urllib.request.Request(
                url,
                data=json.dumps(payload).encode("utf-8"),
                headers={"Content-Type": "application/json"},
                method="POST"
            )
            try:
                with urllib.request.urlopen(req, timeout=180) as response:
                    data = json.loads(response.read().decode("utf-8"))

                candidates = data.get("candidates") or []
                parts = (
                    candidates[0].get("content", {}).get("parts", [])
                    if candidates else []
                )
                isi = "\n".join(
                    part.get("text", "")
                    for part in parts
                    if isinstance(part, dict) and part.get("text")
                ).strip()

                if isi:
                    return {
                        "sukses": True,
                        "hasil": isi,
                        "error": "",
                        "model": model_id,
                        "fallback": indeks_model > 0
                    }

                error_terakhir = f"{model_id} merespons tetapi hasil kosong."
                break

            except urllib.error.HTTPError as e:
                try:
                    detail = e.read().decode("utf-8")
                except Exception:
                    detail = ""
                error_terakhir = f"Gemini HTTP {e.code}. {detail}".strip()

                # Jangan retry/fallback untuk kesalahan permanen seperti API key/bad request.
                if e.code not in transient_codes:
                    return {
                        "sukses": False,
                        "hasil": "",
                        "error": error_terakhir,
                        "model": model_id
                    }

                if percobaan < 3:
                    # Exponential backoff 2, 4, 8 detik + jitter kecil.
                    delay = (2 ** (percobaan + 1)) + random.uniform(0.2, 1.0)
                    time.sleep(delay)
                    continue
                break

            except (urllib.error.URLError, TimeoutError) as e:
                error_terakhir = f"Gangguan koneksi ke Gemini: {e}"
                if percobaan < 3:
                    delay = (2 ** (percobaan + 1)) + random.uniform(0.2, 1.0)
                    time.sleep(delay)
                    continue
                break

            except Exception as e:
                error_terakhir = f"Terjadi kesalahan saat menjalankan Gemini: {e}"
                if percobaan < 3:
                    delay = (2 ** (percobaan + 1)) + random.uniform(0.2, 1.0)
                    time.sleep(delay)
                    continue
                break

        # Setelah model utama gagal karena gangguan sementara, lanjut model fallback.

    return {
        "sukses": False,
        "hasil": "",
        "error": (
            "Layanan Gemini sedang sibuk atau belum dapat dijangkau setelah retry otomatis "
            "dan fallback aman. Silakan coba lagi beberapa saat. Detail terakhir: "
            + error_terakhir
        ),
        "model": ""
    }


def analisis_dengan_gemini(teks, jenis_karya, fokus_analisis):
    """Analisis dokumen menggunakan Gemini API dengan retry dan fallback aman."""
    teks_dokumen = teks[:60000]
    fokus = ", ".join(fokus_analisis) if fokus_analisis else "Analisis akademik menyeluruh"

    prompt = f"""
Anda adalah Asisten Akademik AI untuk mahasiswa S1, S2, dan S3.

Analisis dokumen akademik berikut secara teliti dan hanya berdasarkan
isi dokumen yang diberikan.

Jenis karya yang dipilih:
{jenis_karya}

Fokus analisis:
{fokus}

ATURAN WAJIB:
1. Jangan mengarang informasi.
2. Jika informasi tidak ditemukan, tulis: "Tidak ditemukan dalam dokumen."
3. Jangan membuat nama penulis, teori, metode, hasil, referensi, DOI, research gap, atau novelty yang tidak terdapat dalam dokumen.
4. Bedakan novelty yang diklaim penulis dengan novelty yang benar-benar telah diverifikasi melalui literatur.
5. Pada tahap ini hanya analisis dokumen, bukan pembuktian novelty terhadap seluruh literatur ilmiah.
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


    return _panggil_gemini_rest_aman(
        prompt,
        temperature=0.2,
        max_output_tokens=8192
    )


# ============================================================
# MENJALANKAN DAN MENAMPILKAN HASIL AI
# ============================================================
def jalankan_analisis_ai(
    teks_ai,
    jenis,
    fokus_analisis
):
    hasil_ai = analisis_dengan_gemini(
        teks_ai,
        jenis,
        fokus_analisis
    )

    if hasil_ai["sukses"]:
        st.session_state.hasil_ai_gemini = hasil_ai["hasil"]

        st.success(
            "✅ Analisis AI berhasil."
        )

        if hasil_ai.get("model"):
            st.caption(
                f"Model AI: {hasil_ai['model']}"
            )

        return hasil_ai["hasil"]

    else:
        st.session_state.hasil_ai_gemini = ""

        st.error(
            "❌ Analisis AI belum berhasil."
        )

        st.warning(
            hasil_ai["error"]
        )

        return ""



# ============================================================
# MESIN AI UMUM, REFERENSI, SITASI, DAN EKSPOR
# ============================================================
def panggil_gemini(prompt, temperature=0.25):
    """Panggilan Gemini umum. API key tetap hanya di Streamlit Secrets."""
    return _panggil_gemini_rest_aman(
        prompt,
        temperature=temperature,
        max_output_tokens=8192
    )


def _http_json(url, timeout=30):
    req = urllib.request.Request(url, headers={"User-Agent":"AsistenAkademikAI/2.0 (academic reference tool)"})
    with urllib.request.urlopen(req, timeout=timeout) as response:
        return json.loads(response.read().decode("utf-8"))

def _tahun_crossref(item):
    dp=item.get("published-print") or item.get("published-online") or item.get("issued") or {}
    parts=dp.get("date-parts") or [[]]
    return str(parts[0][0]) if parts and parts[0] else ""

def _crossref_to_ref(item, sumber="Crossref"):
    authors=[]
    for a in item.get("author",[]):
        nama=(" ".join([a.get("given",""),a.get("family","")])).strip()
        if nama: authors.append(nama)
    return {"Judul":(item.get("title") or [""])[0],"Penulis":"; ".join(authors),"Tahun":_tahun_crossref(item),"Jurnal":(item.get("container-title") or [""])[0],"Volume":item.get("volume",""),"Nomor":item.get("issue",""),"Halaman":item.get("page",""),"DOI":item.get("DOI",""),"URL":item.get("URL",""),"Sumber":sumber,"Status":"✅ Metadata terverifikasi Crossref" if item.get("DOI") else "⚠️ DOI belum tersedia"}

def cari_crossref(kata_kunci, jumlah=10):
    if not kata_kunci.strip(): return []
    q=urllib.parse.quote(kata_kunci.strip())
    url=f"https://api.crossref.org/works?query.bibliographic={q}&rows={jumlah}&select=DOI,title,author,published-print,published-online,issued,container-title,volume,issue,page,URL,type"
    try: return [_crossref_to_ref(x) for x in _http_json(url).get("message",{}).get("items",[])]
    except Exception: return []

def cari_openalex(kata_kunci, jumlah=10):
    if not kata_kunci.strip(): return []
    try:
        data=_http_json(f"https://api.openalex.org/works?search={urllib.parse.quote(kata_kunci.strip())}&per-page={jumlah}")
        out=[]
        for x in data.get("results",[]):
            authors=[a.get("author",{}).get("display_name","") for a in x.get("authorships",[]) if a.get("author",{}).get("display_name")]
            loc=x.get("primary_location") or {}; src=loc.get("source") or {}; b=x.get("biblio") or {}; doi=(x.get("doi") or "").replace("https://doi.org/","")
            pages="-".join([str(v) for v in [b.get("first_page"),b.get("last_page")] if v])
            out.append({"Judul":x.get("title","") or "","Penulis":"; ".join(authors),"Tahun":str(x.get("publication_year") or ""),"Jurnal":src.get("display_name","") or "","Volume":b.get("volume","") or "","Nomor":b.get("issue","") or "","Halaman":pages,"DOI":doi,"URL":x.get("id","") or "","Sumber":"OpenAlex","Status":"✅ Metadata teridentifikasi OpenAlex" if doi else "⚠️ DOI belum tersedia"})
        return out
    except Exception: return []


def cari_semantic_scholar(kata_kunci, jumlah=10):
    """Pencarian metadata publik Semantic Scholar. Gagal diam-diam bila rate-limit."""
    if not kata_kunci.strip(): return []
    try:
        q=urllib.parse.quote(kata_kunci.strip())
        fields="title,authors,year,venue,externalIds,url"
        data=_http_json(f"https://api.semanticscholar.org/graph/v1/paper/search?query={q}&limit={min(jumlah,100)}&fields={fields}")
        out=[]
        for x in data.get("data",[]):
            ext=x.get("externalIds") or {}
            doi=ext.get("DOI","") or ""
            authors=[a.get("name","") for a in x.get("authors",[]) if a.get("name")]
            out.append({"Judul":x.get("title","") or "","Penulis":"; ".join(authors),
                        "Tahun":str(x.get("year") or ""),"Jurnal":x.get("venue","") or "",
                        "Volume":"","Nomor":"","Halaman":"","DOI":doi,
                        "URL":x.get("url","") or "","Sumber":"Semantic Scholar",
                        "Status":"✅ Metadata teridentifikasi Semantic Scholar" if doi else "🔎 Metadata ditemukan — DOI belum tersedia"})
        return out
    except Exception:
        return []

def cari_library_of_congress(kata_kunci, jumlah=8):
    """Pencarian koleksi digital Library of Congress melalui JSON API resmi."""
    if not kata_kunci.strip(): return []
    try:
        q=urllib.parse.quote(kata_kunci.strip())
        data=_http_json(f"https://www.loc.gov/search/?q={q}&fo=json&c={min(jumlah,25)}")
        out=[]
        for x in data.get("results",[])[:jumlah]:
            title=x.get("title","") or ""
            date=str(x.get("date","") or "")
            creator=x.get("contributor") or x.get("creator") or []
            if isinstance(creator,str): creator=[creator]
            authors="; ".join([str(a) for a in creator[:8]])
            out.append({"Judul":title,"Penulis":authors,"Tahun":date[:4] if date else "",
                        "Jurnal":"","Volume":"","Nomor":"","Halaman":"","DOI":"",
                        "URL":x.get("id","") or x.get("url","") or "",
                        "Sumber":"Library of Congress",
                        "Status":"🏛️ Metadata katalog teridentifikasi Library of Congress"})
        return out
    except Exception:
        return []

def cari_multi_sumber(kata_kunci, jumlah=12):
    """Federated search: metadata sources that legally expose machine-readable APIs."""
    unik=[]; seen=set()
    gabungan=(cari_crossref(kata_kunci,jumlah)+cari_openalex(kata_kunci,jumlah)
              +cari_semantic_scholar(kata_kunci,min(jumlah,10))
              +cari_library_of_congress(kata_kunci,min(jumlah,8)))
    for r in gabungan:
        k=(r.get("DOI") or re.sub(r"[^a-z0-9]+","",r.get("Judul","").lower())).lower()
        if k and k not in seen:
            seen.add(k); unik.append(r)
    return unik

def cari_crossref_doi(doi):
    doi=(doi or "").strip().strip(".,;:) ]}")
    if not doi: return None
    try: return _crossref_to_ref(_http_json("https://api.crossref.org/works/"+urllib.parse.quote(doi,safe="")).get("message") or {})
    except Exception: return None

def ekstrak_doi(teks):
    m=re.search(r"10\.\d{4,9}/[-._;()/:A-Z0-9]+",teks or "",re.I)
    return m.group(0).rstrip(".,;:) ]}") if m else ""

def normal_judul(x): return re.sub(r"[^a-z0-9]+"," ",(x or "").lower()).strip()

def verifikasi_judul_crossref(judul):
    best=None; score=0.0; a=normal_judul(judul)
    for r in cari_crossref(judul,5):
        sc=difflib.SequenceMatcher(None,a,normal_judul(r.get("Judul",""))).ratio() if a else 0
        if sc>score: best,score=r,sc
    return best,score

def ekstrak_metadata_gemini(teks,nama_file=""):
    if not teks or teks.startswith("ERROR:"): return None
    prompt=("Ekstrak metadata bibliografis dari dokumen berikut. Jangan menebak. Jika tidak ada isi string kosong. "
            "Kembalikan HANYA JSON valid dengan kunci Judul, Penulis, Tahun, Jurnal, Volume, Nomor, Halaman, DOI, URL. "
            "Penulis dipisahkan titik koma. Nama file: "+nama_file+"\nDOKUMEN:\n"+teks[:15000])
    h=panggil_gemini(prompt,0.05)
    if not h.get("sukses"): return None
    raw=re.sub(r"^```(?:json)?\s*|\s*```$","",h.get("hasil","").strip(),flags=re.I|re.S).strip()
    try:
        d=json.loads(raw); return {k:str(d.get(k,"") or "").strip() for k in ["Judul","Penulis","Tahun","Jurnal","Volume","Nomor","Halaman","DOI","URL"]}
    except Exception: return None

def kunci_ref(ref):
    doi=(ref.get("DOI") or "").strip().lower()
    return "doi:"+doi if doi else "title:"+re.sub(r"[^a-z0-9]+","",(ref.get("Judul") or "").lower())

def tambah_bank_referensi(ref):
    key=kunci_ref(ref)
    if any(kunci_ref(x)==key for x in st.session_state.bank_referensi): return False
    data=dict(ref); data["Proyek"]=st.session_state.proyek_aktif; data["Tanggal"]=datetime.now().strftime("%d-%m-%Y %H:%M"); st.session_state.bank_referensi.append(data); return True

def _nama_chicago(penulis):
    n=[x.strip() for x in (penulis or "").split(";") if x.strip()]
    return "Tanpa penulis" if not n else n[0] if len(n)==1 else f"{n[0]} dan {n[1]}" if len(n)==2 else f"{n[0]} et al."

def format_chicago_note(ref,halaman_kutip=""):
    pen=_nama_chicago(ref.get("Penulis")); jud=ref.get("Judul") or "Tanpa judul"; jur=ref.get("Jurnal") or ""; vol=ref.get("Volume") or ""; no=ref.get("Nomor") or ""; th=ref.get("Tahun") or "n.d."; doi=ref.get("DOI") or ""; url=ref.get("URL") or ""
    pub=jur + (f" {vol}" if vol else "") + (f", no. {no}" if no else "") + f" ({th})"; loc=halaman_kutip or ref.get("Halaman") or ""
    if loc: pub+=f": {loc}"
    return f'{pen}, “{jud},” {pub}'+(f", https://doi.org/{doi}" if doi else f", {url}" if url else "")+"."

def format_chicago_bibliography(ref):
    pen=_nama_chicago(ref.get("Penulis")); jud=ref.get("Judul") or "Tanpa judul"; jur=ref.get("Jurnal") or ""; vol=ref.get("Volume") or ""; no=ref.get("Nomor") or ""; th=ref.get("Tahun") or "n.d."; hal=ref.get("Halaman") or ""; doi=ref.get("DOI") or ""; url=ref.get("URL") or ""
    s=f'{pen}. “{jud}.”'+(f" {jur}" if jur else "")+(f" {vol}" if vol else "")+(f", no. {no}" if no else "")+f" ({th})"+(f": {hal}" if hal else "")
    return s+(f". https://doi.org/{doi}" if doi else f". {url}" if url else "")+"."

def format_apa(ref):
    pen=ref.get("Penulis") or "Tanpa penulis"; th=ref.get("Tahun") or "n.d."; jud=ref.get("Judul") or "Tanpa judul"; jur=ref.get("Jurnal") or ""; vol=ref.get("Volume") or ""; no=ref.get("Nomor") or ""; hal=ref.get("Halaman") or ""; doi=ref.get("DOI") or ""; tail=""
    if jur: tail+=f" {jur}"
    if vol: tail+=f", {vol}"
    if no: tail+=f"({no})"
    if hal: tail+=f", {hal}"
    if doi: tail+=f". https://doi.org/{doi}"
    return f"{pen}. ({th}). {jud}.{tail}".strip()

def format_referensi(ref,gaya=None):
    gaya=gaya or st.session_state.get("gaya_sitasi","Chicago Notes & Bibliography")
    return format_chicago_bibliography(ref) if gaya.startswith("Chicago") else format_apa(ref)



def deteksi_gaya_sitasi_otomatis(teks):
    """Deteksi konservatif gaya sitasi/footnote dari pola naskah; hasil dapat dikoreksi pengguna."""
    t=(teks or "").strip()
    if not t:
        return {"gaya":"Tidak terdeteksi","keyakinan":"Rendah","alasan":"Naskah belum tersedia."}
    tl=t.lower()
    scores={"Chicago Notes & Bibliography":0,"Turabian Notes-Bibliography":0,"OSCOLA":0,"APA 7":0,"Harvard":0,"MLA":0,"IEEE":0,"Vancouver":0,"AMA":0,"ACS":0,"CSE":0,"APSA":0}
    reasons=[]
    bracket=len(re.findall(r"\[(?:\d{1,3})(?:\s*[-,]\s*\d{1,3})*\]",t))
    supers=len(re.findall(r"[¹²³⁴⁵⁶⁷⁸⁹⁰]+",t))
    if bracket>=2:
        scores["IEEE"]+=4; scores["Vancouver"]+=3; scores["AMA"]+=2; reasons.append("ditemukan pola sitasi numerik")
    if supers>=2:
        scores["Vancouver"]+=3; scores["AMA"]+=3; scores["Chicago Notes & Bibliography"]+=2; reasons.append("ditemukan penanda angka superscript")
    paren_apa=len(re.findall(r"\([A-ZÀ-ÖØ-Þ][A-Za-zÀ-ÿ'’.-]+(?:\s+et\s+al\.)?,?\s+(?:19|20)\d{2}[a-z]?\)",t))
    author_date=len(re.findall(r"[A-ZÀ-ÖØ-Þ][A-Za-zÀ-ÿ'’.-]+\s+\((?:19|20)\d{2}[a-z]?\)",t))
    if paren_apa+author_date>=2:
        scores["APA 7"]+=4; scores["Harvard"]+=3; scores["APSA"]+=2; reasons.append("ditemukan pola penulis-tahun")
    mla=len(re.findall(r"\([A-ZÀ-ÖØ-Þ][A-Za-zÀ-ÿ'’.-]+\s+\d{1,4}(?:[-–]\d{1,4})?\)",t))
    if mla>=2:
        scores["MLA"]+=5; reasons.append("ditemukan pola penulis-halaman")
    note_clues=sum(tl.count(x) for x in ["ibid.","ibid,","op. cit","loc. cit"])
    if note_clues:
        scores["Chicago Notes & Bibliography"]+=5; scores["Turabian Notes-Bibliography"]+=4; reasons.append("ditemukan pola catatan kaki bibliografis")
    if re.search(r"\bvol\.\s*\d+.*\bno\.\s*\d+",tl):
        scores["Chicago Notes & Bibliography"]+=2; scores["Turabian Notes-Bibliography"]+=2
    if re.search(r"\b(v\.|vs\.|case|court|act\s+\d{4}|statute|regulation)\b",tl):
        scores["OSCOLA"]+=5; reasons.append("ditemukan pola rujukan hukum")
    best=max(scores,key=scores.get); val=scores[best]
    if val<=1:
        return {"gaya":"Gaya bawaan/Custom","keyakinan":"Rendah","alasan":"Pola belum cukup kuat; pertahankan format asli dan lakukan audit."}
    sorted_vals=sorted(scores.values(),reverse=True); gap=val-(sorted_vals[1] if len(sorted_vals)>1 else 0)
    conf="Tinggi" if val>=5 and gap>=2 else "Sedang" if val>=3 else "Rendah"
    return {"gaya":best,"keyakinan":conf,"alasan":"; ".join(dict.fromkeys(reasons)) or "pola sitasi terdeteksi"}


def prompt_perbaiki_footnote_kutipan(teks, gaya, refs, mode="Pertahankan gaya bawaan"):
    daftar="\n".join(f"[{i}] {format_referensi(r, gaya if gaya not in ['Gaya bawaan/Custom','Deteksi Otomatis'] else None)} | DOI: {r.get('DOI','')} | STATUS: {r.get('Status','')}" for i,r in enumerate(refs[:80],1)) or "LIBRARY KOSONG"
    return f"""Anda adalah editor sitasi akademik yang sangat konservatif.
MODE: {mode}
GAYA TARGET: {gaya}

TUGAS:
1. Pertahankan isi, struktur, data, argumen, dan urutan naskah. Jangan menulis ulang substansi.
2. Audit semua footnote/endnote/in-text citation dan cocokkan dengan sumber di LIBRARY.
3. Jangan mengarang penulis, judul, DOI, tahun, halaman, kutipan langsung, atau sumber.
4. Jika sumber tidak dapat diverifikasi dari Library, beri label [PERLU VERIFIKASI], jangan menggantinya dengan tebakan.
4a. Untuk sumber yang belum ditemukan, berikan SARAN PENCARIAN berdasarkan penulis/judul/tahun/DOI/topik yang benar-benar terbaca dari naskah.
4b. Jika ada sumber Library yang relevan tetapi bukan sumber asli, labeli jelas sebagai [REFERENSI ALTERNATIF - BUKAN SUMBER ASLI] dan jangan mengganti otomatis.
4c. Gunakan tiga status: ✅ TERVERIFIKASI, ⚠️ KANDIDAT/PERLU KONFIRMASI, ❌ TIDAK DITEMUKAN.
5. Jika mode mempertahankan gaya bawaan, ikuti pola footnote/sitasi yang sudah dominan di naskah. Jangan memaksa Chicago.
6. Jika gaya target memakai footnote/endnote, rapikan nomor dan konsistensinya. Jika gaya target memakai sitasi dalam teks, pertahankan sistem in-text tersebut.
7. Sinkronkan daftar pustaka hanya dengan sumber yang benar-benar digunakan/teridentifikasi.
8. Kalimat utama hanya boleh diperbaiki bila tidak cocok dengan sumber atau sangat tidak efektif. Setiap perubahan kalimat wajib ditampilkan sebagai SEBELUM -> SESUDAH dan jangan diterapkan diam-diam.
9. Untuk halaman yang tidak diketahui, tulis [halaman perlu verifikasi], jangan menciptakan nomor halaman.

KELUARAN WAJIB:
A. Gaya yang digunakan/dipertahankan dan alasan singkat.
B. Tabel audit: No | Lokasi/Penanda | Sumber Lama | Status | Sumber Benar/Usulan | Tindakan.
C. Naskah hasil perbaikan sitasi/footnote.
D. Daftar pustaka tersinkron.
E. Daftar bagian [PERLU VERIFIKASI].
F. SEBELUM -> SESUDAH hanya jika ada kalimat utama yang perlu perubahan.

LIBRARY REFERENSI:
{daftar}

NASKAH:
{teks[:70000]}"""


def ekspor_ris(refs):
    out=[]
    for r in refs:
        out += ["TY  - JOUR",f"TI  - {r.get('Judul','')}",f"PY  - {r.get('Tahun','')}"]
        for a in [x.strip() for x in (r.get("Penulis") or "").split(";") if x.strip()]: out.append(f"AU  - {a}")
        for tag,key in [("JO","Jurnal"),("VL","Volume"),("IS","Nomor"),("SP","Halaman"),("DO","DOI"),("UR","URL")]:
            if r.get(key): out.append(f"{tag}  - {r[key]}")
        out += ["ER  - ",""]
    return "\n".join(out)

def ekspor_bibtex(refs):
    out=[]
    for i,r in enumerate(refs,1):
        out.append("@article{ref"+str(i)+",\n"+f"  author = {{{(r.get('Penulis') or 'Unknown').replace(';',' and')}}},\n  title = {{{r.get('Judul','')}}},\n  year = {{{r.get('Tahun') or 'n.d.'}}},\n  journal = {{{r.get('Jurnal','')}}},\n  doi = {{{r.get('DOI','')}}}\n}}")
    return "\n\n".join(out)

def ekspor_endnote_tagged(refs):
    """EndNote Tagged format; dapat diimpor oleh EndNote dan beberapa reference manager."""
    out=[]
    for r in refs:
        out.append("%0 Journal Article" if r.get("Jurnal") else "%0 Book")
        for a in [x.strip() for x in (r.get("Penulis") or "").split(";") if x.strip()]:
            out.append(f"%A {a}")
        if r.get("Tahun"): out.append(f"%D {r['Tahun']}")
        if r.get("Judul"): out.append(f"%T {r['Judul']}")
        if r.get("Jurnal"): out.append(f"%J {r['Jurnal']}")
        if r.get("Volume"): out.append(f"%V {r['Volume']}")
        if r.get("Nomor"): out.append(f"%N {r['Nomor']}")
        if r.get("Halaman"): out.append(f"%P {r['Halaman']}")
        if r.get("DOI"): out.append(f"%R {r['DOI']}")
        if r.get("URL"): out.append(f"%U {r['URL']}")
        if r.get("ISBN"): out.append(f"%@ {r['ISBN']}")
        out.append("")
    return "\n".join(out)

def ekspor_csv_referensi(refs):
    import io
    buf=io.StringIO()
    fields=["Judul","Penulis","Tahun","Jurnal","Volume","Nomor","Halaman","DOI","ISBN","URL","Sumber","Status"]
    w=csv.DictWriter(buf,fieldnames=fields,extrasaction="ignore")
    w.writeheader()
    for r in refs: w.writerow({k:r.get(k,"") for k in fields})
    return buf.getvalue()


def deteksi_sitasi_author_year(teks):
    pola=r"\(([A-ZÀ-ÖØ-Ý][A-Za-zÀ-ÿ'’.-]+(?:\s+(?:&|dan)\s+[A-ZÀ-ÖØ-Ý][A-Za-zÀ-ÿ'’.-]+|\s+et\s+al\.)?),\s*((?:19|20)\d{2})[a-z]?\)"
    return sorted(set((a.strip(),y) for a,y in re.findall(pola,teks)))

def status_sitasi(teks,refs):
    rows=[]
    for author,year in deteksi_sitasi_author_year(teks):
        surname=author.split()[0].lower(); cocok=[r for r in refs if str(r.get("Tahun",""))==year and surname in (r.get("Penulis") or "").lower()]
        rows.append({"Sitasi":f"({author}, {year})","Status":"✅ Ada di Library" if cocok else "⚠️ Belum cocok dengan Library","Referensi":cocok[0].get("Judul","") if cocok else ""})
    return rows


def _bersihkan_entri_daftar_pustaka(x):
    x=re.sub(r"\s+"," ",x or "").strip()
    return re.sub(r"^\s*(?:\[\d+\]|\d+[\.\)]|[-•▪■])\s*","",x).strip()

def ekstrak_daftar_pustaka(teks):
    """
    Ekstraksi bibliografi lebih tahan terhadap:
    - nama lembaga (UNESCO)
    - nama keluarga berawalan huruf kecil (van/de/al-)
    - entri buku ber-ISBN
    - artikel ber-DOI
    - entri yang terbungkus menjadi beberapa baris
    """
    if not teks or teks.startswith("ERROR:"): return []
    raw=teks.replace("\r","\n")
    m=re.search(r"(?im)^\s*(DAFTAR\s+PUSTAKA|REFERENCES|BIBLIOGRAPHY)\s*$",raw)
    if not m: return []
    bagian=raw[m.end():]
    bagian=re.split(r"(?im)^\s*(LAMPIRAN|APPENDIX|BAB\s+[IVXLCDM]+)\b",bagian,maxsplit=1)[0]

    lines=[_bersihkan_entri_daftar_pustaka(x) for x in bagian.splitlines()]
    lines=[x for x in lines if x]

    year_re=re.compile(r"\b(?:19|20)\d{2}[a-z]?\b",re.I)
    doi_re=re.compile(r"(?:https?://(?:dx\.)?doi\.org/|doi\s*:\s*)10\.\d{4,9}/\S+",re.I)
    isbn_re=re.compile(r"\bISBN(?:-1[03])?\s*:?\s*(97[89][\-\dXx ]{10,20})",re.I)

    entries=[]; buf=""
    for ln in lines:
        # A new bibliography entry normally contains a publication year.
        # This deliberately does NOT require the author to start with A-Z,
        # so "van Niekerk..." and institutional authors are retained.
        new_entry=bool(year_re.search(ln))
        # Strong signals that a line belongs to the current reference.
        continuation=bool(re.match(r"^(https?://|doi\b|ISBN\b)",ln,re.I))

        if buf and new_entry and not continuation:
            entries.append(_bersihkan_entri_daftar_pustaka(buf))
            buf=ln
        else:
            buf=(buf+" "+ln).strip()
    if buf:
        entries.append(_bersihkan_entri_daftar_pustaka(buf))

    # Rescue: if two references were merged, split after DOI/URL when another
    # author+year sequence follows. This is common after DOCX text extraction.
    rescued=[]
    for e in entries:
        # split before a likely next author if there are >=2 years in one entry
        years=list(year_re.finditer(e))
        if len(years)<=1:
            rescued.append(e); continue
        # Prefer boundaries following DOI URL / terminal period.
        cuts=[0]
        for ym in years[1:]:
            pre=e[:ym.start()]
            # walk backward to likely author start after previous sentence/URL
            candidates=[pre.rfind(". "), pre.rfind("  ")]
            pos=max(candidates)
            if pos>0 and len(e[pos+2:ym.start()].strip())<140:
                cuts.append(pos+2)
        if len(cuts)==1:
            rescued.append(e)
        else:
            cuts.append(len(e))
            rescued.extend(e[cuts[i]:cuts[i+1]].strip() for i in range(len(cuts)-1))

    out=[]; seen=set()
    for e in rescued:
        e=_bersihkan_entri_daftar_pustaka(e)
        if len(e)<20: continue
        # A valid candidate must have year, DOI, or ISBN; no invented content.
        if not (year_re.search(e) or doi_re.search(e) or isbn_re.search(e)): continue
        k=re.sub(r"[^a-z0-9]+","",e.lower())[:220]
        if k and k not in seen:
            seen.add(k); out.append(e)
    return out


def ekstrak_daftar_pustaka_docx(file):
    """
    Parser utama untuk DOCX. Tidak memecah berdasarkan banyaknya tahun/DOI.
    Prinsip: satu paragraf bibliografi Word = satu sumber.
    Baris lanjutan DOI/URL/ISBN digabung ke sumber sebelumnya.
    """
    if docx is None:
        return []
    try:
        file.seek(0)
        d=docx.Document(file)
        paras=[p.text.strip() for p in d.paragraphs]
        start=None
        for i,t in enumerate(paras):
            if re.fullmatch(r"\s*(DAFTAR\s+PUSTAKA(?:\s+AWAL)?|REFERENCES|BIBLIOGRAPHY)\s*",t,re.I):
                start=i+1
        if start is None:
            return []

        out=[]
        for t in paras[start:]:
            t=_bersihkan_entri_daftar_pustaka(t)
            if not t:
                continue
            if re.match(r"^(LAMPIRAN|APPENDIX|BAB\s+[IVXLCDM]+)\b",t,re.I):
                break
            if re.match(r"^(https?://|doi\s*:|ISBN(?:-1[03])?\s*:?)",t,re.I) and out:
                out[-1]=(out[-1]+" "+t).strip()
                continue
            # Abaikan heading/non-entri sesudah bibliografi.
            if len(t)<18:
                continue
            out.append(t)

        # Dedup konservatif: DOI, lalu keseluruhan entri ternormalisasi.
        hasil=[]; seen=set()
        for e in out:
            doi=ekstrak_doi(e)
            key=("doi:"+doi.lower()) if doi else ("txt:"+re.sub(r"[^a-z0-9]+","",e.lower())[:260])
            if key not in seen:
                seen.add(key); hasil.append(e)
        return hasil
    except Exception:
        return []


def _judul_dari_entri(entri):
    doi=ekstrak_doi(entri)
    if doi:
        r=cari_crossref_doi(doi)
        if r: return r.get("Judul","")
    x=re.sub(r"^.*?\b(?:19|20)\d{2}[a-z]?\b[\.,]?\s*","",entri,count=1,flags=re.I).strip()
    q=re.search(r'[“"]([^”"]{8,400})[”"]',x)
    if q: return q.group(1).strip()
    # Remove ISBN tail before guessing title.
    x=re.sub(r"\bISBN(?:-1[03])?\s*:?.*$","",x,flags=re.I).strip()
    parts=[z.strip(" .“”\"") for z in re.split(r"\.\s+",x) if z.strip()]
    # Skip bare identifiers/publisher fragments.
    for z in parts:
        if len(z)>=8 and not re.match(r"^(ISBN|https?://|doi\b)",z,re.I):
            return z[:400]
    return entri[:300]


def verifikasi_entri_bibliografi(entri):
    doi=ekstrak_doi(entri)
    if doi:
        r=cari_crossref_doi(doi)
        if r:
            r["Sumber"]="Daftar Pustaka dokumen + Crossref"; r["Status"]="✅ Metadata terverifikasi Crossref"; return r
    judul=_judul_dari_entri(entri)
    if judul:
        cand,score=verifikasi_judul_crossref(judul)
        if cand and score>=0.82:
            cand["Sumber"]="Daftar Pustaka dokumen + Crossref"
            cand["Status"]=f"✅ Metadata terverifikasi Crossref (kemiripan judul {score:.0%})"
            return cand
    th=re.search(r"\b((?:19|20)\d{2})\b",entri)
    isbnm=re.search(r"\bISBN(?:-1[03])?\s*:?\s*(97[89][\-\dXx ]{10,20})",entri,re.I)
    isbn=re.sub(r"[^0-9Xx]","",isbnm.group(1)) if isbnm else ""
    status="📘 Buku/laporan teridentifikasi dari bibliografi — verifikasi katalog/ISBN" if isbn else "🔍 Belum terverifikasi — cek metadata/sumber asli"
    return {"Judul":judul or entri[:220],"Penulis":"","Tahun":th.group(1) if th else "",
            "Jurnal":"","Volume":"","Nomor":"","Halaman":"","DOI":doi,"ISBN":isbn,"URL":"",
            "Sumber":"Daftar Pustaka dokumen","Status":status,
            "Entri Asli":entri}

def unggah_referensi_ke_bank(files):
    jumlah=0; laporan=[]
    for f in files or []:
        nama=f.name; teks=ekstrak_teks(f) if nama.lower().endswith((".pdf",".docx",".txt")) else ""
        if nama.lower().endswith(".docx"):
            entries=ekstrak_daftar_pustaka_docx(f)
            if not entries:
                entries=ekstrak_daftar_pustaka(teks)
        else:
            entries=ekstrak_daftar_pustaka(teks)
        if entries:
            for no,entri in enumerate(entries,1):
                ref=verifikasi_entri_bibliografi(entri)
                ref["Nama File"]=nama; ref["Entri Asli"]=entri; ref["Cuplikan"]=entri[:1500]
                added=tambah_bank_referensi(ref); jumlah+=1 if added else 0
                laporan.append({"File":nama,"No.":no,"Judul":ref.get("Judul",""),"DOI":ref.get("DOI",""),
                                "Status":ref.get("Status",""),"Masuk Library":"Ya" if added else "Sudah ada"})
            continue
        # Jika file adalah satu artikel/buku, tetap gunakan alur lama.
        ref=None; doi=ekstrak_doi(teks)
        if doi: ref=cari_crossref_doi(doi)
        if ref:
            ref["Sumber"]="Unggahan pengguna + Crossref"; ref["Status"]="✅ Metadata terverifikasi Crossref"
        else:
            meta=ekstrak_metadata_gemini(teks,nama)
            if meta:
                doi2=ekstrak_doi(meta.get("DOI","")); ref=cari_crossref_doi(doi2) if doi2 else None
                if not ref:
                    cand,score=verifikasi_judul_crossref(meta.get("Judul","")); ref=cand if cand and score>=0.82 else None
                if ref:
                    ref["Sumber"]="Unggahan pengguna + Crossref"; ref["Status"]="✅ Metadata terverifikasi Crossref"
                else:
                    ref=dict(meta); ref.update({"Sumber":"Unggahan pengguna + ekstraksi AI","Status":"🔍 Metadata belum terverifikasi — cek manual"})
            else:
                ref={"Judul":nama.rsplit(".",1)[0],"Penulis":"","Tahun":"","Jurnal":"","Volume":"","Nomor":"","Halaman":"","DOI":"","URL":"","Sumber":"Unggahan pengguna","Status":"🔍 Metadata belum terbaca — cek manual"}
        ref["Nama File"]=nama; ref["Cuplikan"]=teks[:5000] if teks and not teks.startswith("ERROR:") else ""
        added=tambah_bank_referensi(ref); jumlah+=1 if added else 0
        laporan.append({"File":nama,"No.":1,"Judul":ref.get("Judul",""),"DOI":ref.get("DOI",""),"Status":ref.get("Status",""),"Masuk Library":"Ya" if added else "Sudah ada"})
    return jumlah,laporan

def ringkasan_laporan_referensi(laporan, jumlah_baru):
    ditemukan=len(laporan or [])
    sudah_ada=sum(1 for x in (laporan or []) if x.get("Masuk Library")=="Sudah ada")
    terv=sum(1 for x in (laporan or []) if str(x.get("Status","")).startswith("✅"))
    perlu=ditemukan-terv
    return {"Ditemukan":ditemukan,"Baru masuk Library":jumlah_baru,
            "Sudah ada":sudah_ada,"Terverifikasi":terv,"Perlu verifikasi":perlu}

def sumber_online_default():
    """
    Registry portal resmi. {q} diganti kata kunci.
    API langsung hanya dipakai bila layanan memang menyediakan akses mesin.
    Portal login/berlisensi dibuka resmi tanpa melewati autentikasi.
    """
    return [
        # INDONESIA - NASIONAL
        ("🇮🇩 Nasional","Indonesia OneSearch","https://onesearch.id/Search/Results?lookfor={q}&type=AllFields","Agregator katalog dan repository perpustakaan Indonesia"),
        ("🇮🇩 Nasional","OPAC Perpusnas RI","https://opac.perpusnas.go.id/Search/Results?lookfor={q}&type=AllFields","Katalog Perpustakaan Nasional RI"),
        ("🇮🇩 Nasional","e-Resources Perpusnas","https://e-resources.perpusnas.go.id/","Jurnal, ebook, dan basis data berlangganan; login anggota"),
        ("🇮🇩 Nasional","iPusnas","https://ipusnas.id/","Perpustakaan digital Perpusnas"),
        ("🇮🇩 Nasional","GARUDA","https://garuda.kemdiktisaintek.go.id/documents?q={q}","Publikasi ilmiah Indonesia"),
        ("🇮🇩 Nasional","Neliti","https://www.neliti.com/search?q={q}","Repository publikasi dan kebijakan Indonesia"),
        # DAERAH / PROVINSI - titik awal + agregator nasional
        ("🏢 Daerah/Provinsi","Perpustakaan Provinsi Kalimantan Selatan","https://inlislite.dispersip.my.id/opac/search?q={q}","OPAC daerah; akses mengikuti layanan resmi"),
        ("🏫 Daerah/Provinsi","Perpustakaan Kabupaten Banjar","https://perpustakaan.banjarkab.go.id/opac/index.php?keywords={q}&search=search","OPAC Kabupaten Banjar"),
        ("🏢 Daerah/Provinsi","Direktori melalui Indonesia OneSearch","https://onesearch.id/Search/Results?lookfor={q}&type=AllFields","Menjangkau banyak perpustakaan provinsi, kabupaten/kota, kampus, dan repository Indonesia"),
        # INTERNASIONAL - API/OPEN DISCOVERY
        ("🌍 Internasional","Crossref","https://search.crossref.org/?q={q}","Metadata DOI; terhubung langsung ke pencarian aplikasi"),
        ("🌍 Internasional","OpenAlex","https://openalex.org/works?page=1&filter=default.search:{q}","Indeks karya ilmiah terbuka; terhubung langsung"),
        ("🌍 Internasional","Semantic Scholar","https://www.semanticscholar.org/search?q={q}","Artikel, penulis, sitasi; terhubung langsung"),
        ("🌍 Internasional","Google Scholar","https://scholar.google.com/scholar?q={q}","Pencarian akademik; dibuka resmi, tidak di-scrape"),
        ("🌍 Internasional","DOAJ","https://doaj.org/search/articles?ref=homepage-box&q={q}","Jurnal dan artikel open access"),
        ("🌍 Internasional","CORE","https://core.ac.uk/search?q={q}","Agregator karya ilmiah open access"),
        ("🌍 Internasional","BASE","https://www.base-search.net/Search/Results?lookfor={q}&type=all&oaboost=1","Mesin pencari repository akademik global"),
        ("🌍 Internasional","OpenAIRE Explore","https://explore.openaire.eu/search/find?keyword={q}","Publikasi dan keluaran riset Eropa/global"),
        ("🌍 Internasional","WorldCat","https://search.worldcat.org/search?q={q}","Katalog kolektif perpustakaan dunia"),
        ("🌍 Internasional","Library of Congress","https://www.loc.gov/search/?q={q}","Koleksi digital; terhubung ke JSON API resmi"),
        ("🌍 Internasional","Europe PMC","https://europepmc.org/search?query={q}","Literatur biomedis dan life sciences"),
        ("🌍 Internasional","PubMed","https://pubmed.ncbi.nlm.nih.gov/?term={q}","Literatur biomedis"),
        ("🌍 Internasional","ERIC","https://eric.ed.gov/?q={q}","Literatur pendidikan"),
        ("🌍 Internasional","arXiv","https://arxiv.org/search/?query={q}&searchtype=all","Preprint ilmiah"),
        ("🌍 Internasional","Google Books","https://books.google.com/books?q={q}","Pencarian buku dan metadata"),
        ("🌍 Internasional","Internet Archive","https://archive.org/search?query={q}","Koleksi digital buku dan arsip"),
    ]



def cari_referensi_penguat_ide_s2(masalah, arah="", jumlah_per_query=8):
    dasar = " ".join([str(masalah or "").strip(), str(arah or "").strip()]).strip()
    if not dasar:
        return [], []
    p = f"""Buat kata kunci pencarian akademik untuk masalah tesis S2 berikut.
Kembalikan HANYA 8 baris kata kunci, tanpa nomor dan tanpa penjelasan.
Campurkan Bahasa Indonesia dan Bahasa Inggris. Perluas konsep bila istilah utama masih baru.
Jangan membuat judul artikel, nama penulis, DOI, atau referensi.
MASALAH:
{dasar[:6000]}"""
    h = panggil_gemini(p, 0.15)
    queries = []
    if h.get("sukses"):
        for baris in str(h.get("hasil", "")).splitlines():
            q = re.sub(r"^\s*[-*•\d\.\)\:]+\s*", "", baris).strip().strip('"“”')
            if len(q) >= 4 and q not in queries:
                queries.append(q)
    if not queries:
        queries = [dasar[:220]]
    hasil, seen = [], set()
    for q in queries[:8]:
        for r in cari_multi_sumber(q, jumlah_per_query):
            k = kunci_ref(r)
            if k and k not in seen:
                seen.add(k)
                rr = dict(r); rr["Kata Kunci"] = q; hasil.append(rr)
    return queries[:8], hasil

def analisis_ketersediaan_referensi_judul_s2(judul, jumlah=10):
    refs = cari_multi_sumber(judul, jumlah)
    tahun_sekarang = datetime.now().year
    terbaru = 0
    for r in refs:
        try:
            if int(str(r.get("Tahun", ""))[:4]) >= tahun_sekarang - 5:
                terbaru += 1
        except Exception:
            pass
    n = len(refs)
    status = "🟢 Kuat / relatif mudah" if n >= 15 else "🟡 Cukup / perlu perluasan kata kunci" if n >= 6 else "🔴 Terbatas / perlu pencarian lebih luas"
    return {"Judul":judul, "Kandidat":n, "Literatur 5 Tahun":terbaru,
            "Sumber Internasional":n, "Status":status, "Referensi":refs}


def panel_ai_penulisan(konteks, jenis_output, instruksi, referensi=None, key="ai"):
    refs=referensi or []
    gaya=st.session_state.get("gaya_sitasi","Chicago Notes & Bibliography")
    daftar="\n".join(f"[{i}] {format_referensi(r,gaya)} | DOI: {r.get('DOI','')} | STATUS: {r.get('Status','')}" for i,r in enumerate(refs[:60],1)) or "Belum ada referensi terverifikasi di Library."
    prompt=f"""Anda adalah Asisten Akademik AI S1-S3.
Tugas: {jenis_output}
Instruksi pengguna: {instruksi}
Konteks/bahan:
{konteks[:60000]}

GAYA SITASI: {gaya}
REFERENSI YANG BOLEH DIPAKAI:
{daftar}

ATURAN WAJIB:
- Jangan membuat referensi, DOI, data penelitian, hasil uji, nomor halaman sumber, kutipan, atau fakta yang tidak tersedia.
- Prioritaskan referensi berstatus terverifikasi.
- Jika Chicago Notes & Bibliography: beri penanda footnote [^1], [^2], dst. pada klaim; setelah naskah buat CATATAN KAKI bernomor sama dan DAFTAR PUSTAKA. Jangan mengarang halaman spesifik; tulis [halaman perlu verifikasi] jika belum diketahui.
- Daftar pustaka hanya memuat sumber yang benar-benar dipakai dan tanpa duplikasi.
- Jika referensi belum cukup, tandai [PERLU REFERENSI TERVERIFIKASI].
- Untuk BAB hasil penelitian, jangan menciptakan data. Jika data belum ada, buat struktur analisis saja.
- Pertahankan integritas akademik.
"""
    with st.spinner("Gemini sedang menyusun naskah dan menghubungkan referensi..."):
        h=panggil_gemini(prompt)
    if h["sukses"]:
        st.session_state.hasil_penulisan_ai=h["hasil"]
        st.success("✅ Draf AI selesai dengan aturan referensi.")
    else: st.error(h["error"])
    return h


# ============================================================
# WORD HASIL REVISI - PROTEKSI FORMAT NASKAH ASLI
# ============================================================
def buat_word_hasil_revisi(file_asli, teks_hasil_ai):
    """
    Salinan DOCX asli. Hanya footnote/kutipan yang ditambahkan.
    - marker ditempatkan tepat setelah klaim, bukan di akhir paragraf;
    - nomor referensi di naskah dan nomor di footnote dibuat superscript;
    - mencegah footnote ganda pada posisi/isi yang sama;
    - package Word asli tetap dipertahankan.
    """
    if file_asli is None or not str(getattr(file_asli, "name", "")).lower().endswith(".docx"):
        return None, "Fitur ini memerlukan naskah sumber DOCX."

    import io, zipfile, re
    from lxml import etree

    W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    REL = "http://schemas.openxmlformats.org/package/2006/relationships"
    CT = "http://schemas.openxmlformats.org/package/2006/content-types"
    XML = "http://www.w3.org/XML/1998/namespace"
    ns = {"w": W}
    def q(uri, tag): return f"{{{uri}}}{tag}"

    def clean_md(s):
        s = re.sub(r'\[([^\]]+)\]\(([^)]+)\)', r'\1', s or "")
        s = re.sub(r'\\([*_`])', r'\1', s)
        s = re.sub(r'[*_`#>]', '', s)
        return re.sub(r'\s+', ' ', s).strip()

    def para_text(p):
        return "".join(p.xpath(".//w:t/text()", namespaces=ns))

    def add_superscript_props(run):
        rpr = run.find(q(W, "rPr"))
        if rpr is None:
            rpr = etree.Element(q(W, "rPr"))
            run.insert(0, rpr)
        va = rpr.find(q(W, "vertAlign"))
        if va is None:
            va = etree.SubElement(rpr, q(W, "vertAlign"))
        va.set(q(W, "val"), "superscript")
        rs = rpr.find(q(W, "rStyle"))
        if rs is None:
            rs = etree.SubElement(rpr, q(W, "rStyle"))
        rs.set(q(W, "val"), "FootnoteReference")

    def insert_ref_after_char(p, char_pos, fid):
        """
        Sisipkan footnoteReference tepat sesudah char_pos pada teks paragraf.
        Run teks yang terkena dibelah; formatting run disalin.
        """
        cursor = 0
        text_nodes = p.xpath(".//w:t", namespaces=ns)
        for tnode in text_nodes:
            txt = tnode.text or ""
            end = cursor + len(txt)
            if cursor <= char_pos <= end:
                local = max(0, min(len(txt), char_pos - cursor))
                run = tnode.getparent()
                if run.tag != q(W, "r"):
                    break

                before, after = txt[:local], txt[local:]
                tnode.text = before
                if before.startswith(" ") or before.endswith(" "):
                    tnode.set(q(XML, "space"), "preserve")

                parent = run.getparent()
                idx = parent.index(run)

                ref_run = etree.Element(q(W, "r"))
                # Salin properti karakter dari run asli lalu paksa superscript.
                old_rpr = run.find(q(W, "rPr"))
                if old_rpr is not None:
                    ref_run.append(etree.fromstring(etree.tostring(old_rpr)))
                add_superscript_props(ref_run)
                ref = etree.SubElement(ref_run, q(W, "footnoteReference"))
                ref.set(q(W, "id"), str(fid))
                parent.insert(idx + 1, ref_run)

                if after:
                    after_run = etree.Element(q(W, "r"))
                    if old_rpr is not None:
                        after_run.append(etree.fromstring(etree.tostring(old_rpr)))
                    nt = etree.SubElement(after_run, q(W, "t"))
                    nt.text = after
                    if after.startswith(" ") or after.endswith(" "):
                        nt.set(q(XML, "space"), "preserve")
                    parent.insert(idx + 2, after_run)
                return True
            cursor = end
        return False

    try:
        file_asli.seek(0)
        original = file_asli.read()
        ai = str(teks_hasil_ai or "")

        # Catatan kaki AI.
        note_pat = re.compile(
            r'(?ms)^\[\^(\d+)\]\s*:?\s*(.+?)(?=^\[\^\d+\]\s*:|^\[\^\d+\]\s+|\n---|\n\*\*\*|\n#{1,6}\s|\Z)'
        )
        ai_notes = {}
        for m in note_pat.finditer(ai):
            ai_notes[int(m.group(1))] = clean_md(m.group(2))

        if not ai_notes:
            return None, "Hasil AI belum memiliki marker footnote [^1], [^2], dan seterusnya."

        # Ambil kalimat/klaim tepat sebelum masing-masing marker dari bagian SESUDAH.
        contexts = {}
        for n in sorted(ai_notes):
            marker = f"[^{n}]"
            pos = ai.find(marker)
            if pos < 0:
                continue
            before = ai[:pos]
            if "**SESUDAH:**" in before:
                before = before.rsplit("**SESUDAH:**", 1)[-1]
            elif "SESUDAH:" in before:
                before = before.rsplit("SESUDAH:", 1)[-1]
            # buang marker sebelumnya agar tidak mengganggu pencocokan
            before = re.sub(r'\[\^\d+\]', '', before)
            before = re.sub(r'\[PERLU REFERENSI TERVERIFIKASI\]', '', before)
            before = clean_md(before)
            # Klaim adalah segmen sesudah tanda akhir kalimat terakhir.
            parts = re.split(r'(?<=[.!?])\s+', before)
            ctx = clean_md(parts[-1] if parts else before)
            contexts[n] = ctx

        with zipfile.ZipFile(io.BytesIO(original), "r") as zin:
            infos = zin.infolist()
            files = {i.filename: zin.read(i.filename) for i in infos}

        parser = etree.XMLParser(remove_blank_text=False, recover=False)
        doc_root = etree.fromstring(files["word/document.xml"], parser)

        fn_path = "word/footnotes.xml"
        if fn_path in files:
            fn_root = etree.fromstring(files[fn_path], parser)
        else:
            fn_root = etree.Element(q(W, "footnotes"), nsmap={"w": W})
            for fid, tag in [(-1, "separator"), (0, "continuationSeparator")]:
                fn = etree.SubElement(fn_root, q(W, "footnote"))
                fn.set(q(W, "id"), str(fid))
                p = etree.SubElement(fn, q(W, "p"))
                r = etree.SubElement(p, q(W, "r"))
                etree.SubElement(r, q(W, tag))

        existing_ids = []
        existing_note_text = {}
        for fn in fn_root.xpath("./w:footnote", namespaces=ns):
            try:
                fid = int(fn.get(q(W, "id")))
                existing_ids.append(fid)
                if fid > 0:
                    existing_note_text[clean_md("".join(fn.xpath(".//w:t/text()", namespaces=ns)))] = fid
            except Exception:
                pass

        next_id = max([x for x in existing_ids if x > 0], default=0) + 1
        inserted, skipped, missing = [], [], []

        paragraphs = doc_root.xpath(".//w:body//w:p", namespaces=ns)

        for marker_no in sorted(ai_notes):
            ctx = contexts.get(marker_no, "")
            note_text = ai_notes[marker_no]
            if not ctx:
                missing.append(marker_no)
                continue

            best_p = None
            match_start = match_end = -1
            ctx_norm = clean_md(ctx)

            # Utamakan kecocokan kalimat persis.
            for p in paragraphs:
                ptxt = para_text(p)
                pnorm = re.sub(r'\s+', ' ', ptxt)
                idx = pnorm.find(ctx_norm)
                if idx >= 0:
                    best_p = p
                    # Karena pnorm bisa mengubah whitespace, cari suffix unik pada teks asli.
                    suffix = " ".join(ctx_norm.split()[-10:])
                    raw_idx = ptxt.find(suffix)
                    if raw_idx >= 0:
                        match_end = raw_idx + len(suffix)
                    else:
                        raw_idx = ptxt.find(ctx_norm)
                        match_end = raw_idx + len(ctx_norm) if raw_idx >= 0 else len(ptxt)
                    break

            # Fallback: cari 10/7 kata terakhir klaim.
            if best_p is None:
                words = ctx_norm.split()
                for count in (10, 7, 5):
                    if len(words) < count:
                        continue
                    phrase = " ".join(words[-count:])
                    for p in paragraphs:
                        ptxt = para_text(p)
                        idx = ptxt.find(phrase)
                        if idx >= 0:
                            best_p = p
                            match_end = idx + len(phrase)
                            break
                    if best_p is not None:
                        break

            if best_p is None or match_end < 0:
                missing.append(marker_no)
                continue

            # Jangan menambahkan lagi bila tepat sesudah lokasi itu sudah ada footnoteReference.
            existing_refs = best_p.xpath(".//w:footnoteReference", namespaces=ns)
            # Jika isi catatan identik sudah ada di dokumen, jangan membuat nomor kedua.
            if note_text in existing_note_text:
                skipped.append(marker_no)
                continue

            fid = next_id
            next_id += 1

            if not insert_ref_after_char(best_p, match_end, fid):
                missing.append(marker_no)
                continue

            # True footnote body.
            fn = etree.SubElement(fn_root, q(W, "footnote"))
            fn.set(q(W, "id"), str(fid))
            p = etree.SubElement(fn, q(W, "p"))
            ppr = etree.SubElement(p, q(W, "pPr"))
            pstyle = etree.SubElement(ppr, q(W, "pStyle"))
            pstyle.set(q(W, "val"), "FootnoteText")

            # Nomor pada footnote: hanya SATU footnoteRef, superscript.
            rnum = etree.SubElement(p, q(W, "r"))
            add_superscript_props(rnum)
            etree.SubElement(rnum, q(W, "footnoteRef"))

            rtxt = etree.SubElement(p, q(W, "r"))
            rpr = etree.SubElement(rtxt, q(W, "rPr"))
            fonts = etree.SubElement(rpr, q(W, "rFonts"))
            fonts.set(q(W, "ascii"), "Times New Roman")
            fonts.set(q(W, "hAnsi"), "Times New Roman")
            sz = etree.SubElement(rpr, q(W, "sz"))
            sz.set(q(W, "val"), "20")
            t = etree.SubElement(rtxt, q(W, "t"))
            t.set(q(XML, "space"), "preserve")
            t.text = " " + note_text

            existing_note_text[note_text] = fid
            inserted.append(marker_no)

        if not inserted and not skipped:
            return None, "Tidak ada marker yang dapat dicocokkan secara aman dengan naskah Word asli."

        # Pastikan SEMUA nomor footnote, lama maupun baru, benar-benar superscript.
        # Di badan naskah: w:footnoteReference
        for ref_el in doc_root.xpath(".//w:footnoteReference", namespaces=ns):
            run = ref_el.getparent()
            if run is not None and run.tag == q(W, "r"):
                add_superscript_props(run)

        # Di bagian catatan kaki: w:footnoteRef
        for ref_el in fn_root.xpath(".//w:footnoteRef", namespaces=ns):
            run = ref_el.getparent()
            if run is not None and run.tag == q(W, "r"):
                add_superscript_props(run)

        files["word/document.xml"] = etree.tostring(
            doc_root, xml_declaration=True, encoding="UTF-8", standalone="yes"
        )
        files[fn_path] = etree.tostring(
            fn_root, xml_declaration=True, encoding="UTF-8", standalone="yes"
        )

        # Relationship footnotes.
        rel_path = "word/_rels/document.xml.rels"
        rel_type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes"
        rel_root = etree.fromstring(files[rel_path], parser) if rel_path in files else etree.Element(q(REL, "Relationships"))
        if not any(x.get("Type") == rel_type for x in rel_root):
            used = {x.get("Id") for x in rel_root}
            i = 1
            while f"rId{i}" in used:
                i += 1
            rel = etree.SubElement(rel_root, q(REL, "Relationship"))
            rel.set("Id", f"rId{i}")
            rel.set("Type", rel_type)
            rel.set("Target", "footnotes.xml")
            files[rel_path] = etree.tostring(rel_root, xml_declaration=True, encoding="UTF-8", standalone="yes")

        # Content type footnotes.
        ct_path = "[Content_Types].xml"
        ct_root = etree.fromstring(files[ct_path], parser)
        if not any(x.get("PartName") == "/word/footnotes.xml" for x in ct_root):
            ov = etree.SubElement(ct_root, q(CT, "Override"))
            ov.set("PartName", "/word/footnotes.xml")
            ov.set("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml")
            files[ct_path] = etree.tostring(ct_root, xml_declaration=True, encoding="UTF-8", standalone="yes")

        output = io.BytesIO()
        with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as zout:
            written = set()
            for info in infos:
                if info.filename in files and info.filename not in written:
                    zout.writestr(info, files[info.filename])
                    written.add(info.filename)
            for name, data in files.items():
                if name not in written:
                    zout.writestr(name, data)

        result = output.getvalue()

        # Validasi ZIP/XML dan pasangan ID reference ↔ footnote.
        with zipfile.ZipFile(io.BytesIO(result), "r") as z:
            if z.testzip():
                return None, "Validasi paket DOCX gagal."
            droot = etree.fromstring(z.read("word/document.xml"), parser)
            froot = etree.fromstring(z.read("word/footnotes.xml"), parser)
            refs = [int(x.get(q(W, "id"))) for x in droot.xpath(".//w:footnoteReference", namespaces=ns)]
            fids = {int(x.get(q(W, "id"))) for x in froot.xpath("./w:footnote", namespaces=ns)}
            orphan = [x for x in refs if x not in fids]
            if orphan:
                return None, "Validasi footnote gagal: ada reference tanpa footnote."

        msg = f"Berhasil menambahkan {len(inserted)} footnote Word pada posisi kutipan."
        if skipped:
            msg += f" {len(skipped)} footnote identik tidak digandakan."
        if missing:
            msg += " Marker yang tidak aman dicocokkan: " + ", ".join(map(str, missing)) + "."
        return result, msg

    except Exception as e:
        return None, f"Gagal membuat salinan Word: {e}"

# ============================================================
# HEADER
# ============================================================
st.title("🎓 Asisten Akademik AI")
st.caption(APP_SUBTITLE)

st.info(
    "Asisten Akademik AI S1–S3: OBE, penelitian, referensi tervalidasi, "
    "sitasi, publikasi, buku, penyuntingan, presentasi dan sidang. "
    "Gemini digunakan untuk pekerjaan generatif; metadata referensi diverifikasi terpisah."
)


# ============================================================
# SIDEBAR
# ============================================================
st.sidebar.title("🎓 ASISTEN AKADEMIK AI")

st.sidebar.text_input(
    "Proyek Aktif",
    key="proyek_aktif"
)

menu_utama = st.sidebar.radio(
    "Menu Utama",
    [
        "🎓 Perkuliahan",
        "🔎 Analisis Karya Akademik",
        "🎓 Skripsi S1",
        "🎓 Tesis S2",
        "🎓 Disertasi S3",
        "📚 Literatur & Referensi",
        "✨ Penyunting Akademik AI",
        "📝 Jurnal Akademik",
        "❤️ Donasi & Akses",
        "⚙️ Admin"
    ]
)

# Pemetaan 10 menu utama ke modul yang sudah ada.
# Fitur penelitian lama tetap dipakai, tetapi ditempatkan di dalam S1, S2, dan S3.
if menu_utama == "🎓 Perkuliahan":
    menu = "📚 Perkuliahan & OBE"
elif menu_utama == "🔎 Analisis Karya Akademik":
    menu = "🔬 Analisis Karya Akademik"
elif menu_utama == "📚 Literatur & Referensi":
    menu = "🔎 Literatur & Referensi"
elif menu_utama == "✨ Penyunting Akademik AI":
    menu = "✨ Penyunting Akademik AI"
elif menu_utama == "📝 Jurnal Akademik":
    menu = "📑 Publikasi Jurnal"
elif menu_utama == "❤️ Donasi & Akses":
    menu = "💚 Donasi & Akses"
elif menu_utama == "⚙️ Admin":
    menu = "⚙️ Admin"
else:
    fitur_penelitian = [
        "Ruang Utama",
        "🧭 Metodologi Penelitian",
        "📝 Instrumen Penelitian",
        "📊 Statistik & SPSS",
        "🔤 Analisis Kualitatif",
        "🎤 Audio & Video",
        "✍️ Penulisan Akademik",
        "👨‍🏫 Bimbingan & Revisi",
        "📂 Perpustakaan Akademik",
        "✅ Audit Akademik",
        "📈 Progres Penelitian",
        "🖥️ Presentasi",
        "🎓 Simulasi Sidang",
        "📘 Penulis Buku AI"
    ]
    bagian_penelitian = st.sidebar.radio(
        "Bagian",
        fitur_penelitian,
        key=f"bagian_{menu_utama}"
    )
    if bagian_penelitian == "Ruang Utama":
        if menu_utama == "🎓 Skripsi S1":
            menu = "🎓 Penelitian S1 • S2 • S3"
        elif menu_utama == "🎓 Tesis S2":
            menu = "🎓 Tesis S2"
        else:
            menu = "🧑‍🎓 Disertasi S3"
    else:
        menu = bagian_penelitian

st.sidebar.divider()
st.sidebar.caption(
    "Data permanen per pengguna dan per proyek "
    "akan menggunakan database pada tahap berikutnya."
)


# ============================================================
# PENDAMPING AI UNTUK MODUL PENELITIAN S1/S2/S3
# ============================================================
modul_ai_penelitian = {
    "🧭 Metodologi Penelitian", "📝 Instrumen Penelitian", "📊 Statistik & SPSS",
    "🔤 Analisis Kualitatif", "🎤 Audio & Video", "✍️ Penulisan Akademik",
    "👨‍🏫 Bimbingan & Revisi", "📂 Perpustakaan Akademik", "✅ Audit Akademik",
    "📈 Progres Penelitian", "🖥️ Presentasi", "🎓 Simulasi Sidang", "📘 Penulis Buku AI"
}
if menu in modul_ai_penelitian and menu_utama in ["🎓 Skripsi S1", "🎓 Tesis S2", "🎓 Disertasi S3"]:
    with st.expander("🤖 Asisten AI Bagian Ini", expanded=False):
        bahan_ai_modul = st.file_uploader(
            "Unggah bahan PDF/DOCX/TXT (opsional)",
            type=["pdf", "docx", "txt"],
            key=f"upload_ai_modul_{menu_utama}_{menu}"
        )
        teks_bahan_ai = ekstrak_teks(bahan_ai_modul) if bahan_ai_modul else ""
        instruksi_ai_modul = st.text_area(
            "Apa yang ingin dikerjakan AI?",
            key=f"instruksi_ai_modul_{menu_utama}_{menu}"
        )
        if st.button(
            "🤖 Generate AI",
            type="primary",
            disabled=not bool(instruksi_ai_modul.strip() or teks_bahan_ai.strip()),
            key=f"generate_ai_modul_{menu_utama}_{menu}"
        ):
            konteks_ai = f"JENJANG: {menu_utama}\nBAGIAN: {menu}\nBAHAN:\n{teks_bahan_ai[:60000]}"
            panel_ai_penulisan(konteks_ai, f"{menu_utama} - {menu}", instruksi_ai_modul or "Kerjakan bagian ini berdasarkan bahan yang tersedia.", st.session_state.bank_referensi, "modul_penelitian")
        if st.session_state.get("hasil_penulisan_ai"):
            st.text_area("Hasil AI — dapat diedit", st.session_state.hasil_penulisan_ai, height=400, key=f"hasil_ai_modul_{menu_utama}_{menu}")


# ============================================================
# BERANDA
# ============================================================
if menu == "🏠 Beranda":

    st.header("🏠 Pusat Asisten Akademik")

    st.write(
        """
        Satu ruang kerja akademik untuk mendampingi proses
        dari tugas S1–S3 berbasis OBE sampai skripsi, tesis, disertasi,
        publikasi ilmiah, penulisan buku dan ujian akademik.
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
        "Mode AI",
        "Gemini Langsung"
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
elif menu == "📚 Perkuliahan & OBE":

    st.header("📚 Asisten Perkuliahan & OBE")

    st.caption(
        "Ruang kerja mata kuliah untuk menyusun tugas akademik, "
        "mencari referensi ilmiah, menghubungkan sitasi, "
        "dan menghasilkan naskah yang dapat diedit."
    )

    # ========================================================
    # MATA KULIAH SAYA
    # ========================================================
    st.subheader("📚 Mata Kuliah Saya")

    daftar_mata_kuliah = [
        "Seminar Proposal Tesis",
        "Kepemimpinan dan Supervisi PAI",
    ] + st.session_state.mata_kuliah_tambahan + ["➕ Tambah Mata Kuliah"]

    mata_kuliah = st.selectbox(
        "Pilih Mata Kuliah",
        daftar_mata_kuliah,
        key="pilih_mata_kuliah"
    )

    # ========================================================
    # TAMBAH MATA KULIAH
    # ========================================================
    if mata_kuliah == "➕ Tambah Mata Kuliah":

        st.markdown("### ➕ Tambah Mata Kuliah")

        nama_mk = st.text_input(
            "Nama Mata Kuliah",
            key="nama_mata_kuliah_baru"
        )

        if st.button(
            "Tambahkan Mata Kuliah",
            key="tambah_mata_kuliah"
        ):
            if nama_mk.strip():
                nama_baru = nama_mk.strip()
                daftar_bawaan = ["Seminar Proposal Tesis", "Kepemimpinan dan Supervisi PAI"]
                if nama_baru in daftar_bawaan or nama_baru in st.session_state.mata_kuliah_tambahan:
                    st.warning("Mata kuliah tersebut sudah ada.")
                else:
                    st.session_state.mata_kuliah_tambahan.append(nama_baru)
                    st.success(f"Mata kuliah '{nama_baru}' berhasil ditambahkan.")
                    st.rerun()
            else:
                st.warning(
                    "Masukkan nama mata kuliah terlebih dahulu."
                )

    # ========================================================
    # RUANG KERJA MATA KULIAH
    # ========================================================
    else:

        if mata_kuliah == "Seminar Proposal Tesis":
            st.markdown("### 🎓 Seminar Proposal Tesis")
        else:
            st.markdown("### 👨‍🏫 Kepemimpinan dan Supervisi PAI")

        st.info(
            "Pilih Tugas Saya untuk membuat tugas perkuliahan. "
            "RPS dan materi kuliah dapat disimpan sebagai bahan pendukung, "
            "tetapi tidak perlu dimasukkan setiap kali membuat tugas."
        )

        bagian = st.selectbox(
            "Pilih Bagian",
            [
                "📝 Tugas Saya",
                "📄 RPS / Modul",
                "📚 Materi Perkuliahan",
                "📌 Catatan Dosen",
                "📖 Referensi Mata Kuliah"
            ],
            key="bagian_mata_kuliah"
        )

        # ====================================================
        # TUGAS SAYA
        # ====================================================
        if bagian == "📝 Tugas Saya":

            st.markdown("## 📝 Tugas Saya")

            jenis_tugas = st.selectbox(
                "Jenis Tugas",
                [
                    "Makalah",
                    "Paper",
                    "Resume",
                    "Review Jurnal",
                    "Review Buku",
                    "Critical Review",
                    "Studi Kasus",
                    "Laporan",
                    "Mini Riset",
                    "Presentasi"
                ],
                key="jenis_tugas_mata_kuliah"
            )

            judul_tugas = st.text_input(
                "Judul / Tema Tugas",
                placeholder="Masukkan judul atau tema tugas",
                key="judul_tugas_mata_kuliah"
            )

            st.markdown("### 📎 Bahan Tugas")

            files_tugas = st.file_uploader(
                "Unggah tugas dari dosen atau bahan pendukung bila ada",
                type=[
                    "pdf",
                    "docx",
                    "txt",
                    "csv",
                    "xlsx",
                    "pptx",
                    "jpg",
                    "jpeg",
                    "png"
                ],
                accept_multiple_files=True,
                key="upload_tugas_mata_kuliah"
            )

            bahan_teks = ""

            if files_tugas:

                st.success(
                    f"{len(files_tugas)} file berhasil dipilih."
                )

                for file in files_tugas:

                    st.write(
                        f"📄 **{file.name}** — "
                        f"{format_ukuran(file.size)}"
                    )

                    if file.name.lower().endswith(
                        (".pdf", ".docx", ".txt")
                    ):
                        teks_file = ekstrak_teks(file)

                        if (
                            teks_file
                            and not teks_file.startswith("ERROR:")
                        ):
                            bahan_teks += (
                                f"\n\n===== {file.name} =====\n"
                                + teks_file
                            )

            # =================================================
            # REFERENSI OTOMATIS
            # =================================================
            st.divider()

            st.markdown("## 🔎 Referensi Akademik")

            st.caption(
                "Akademia AI dapat mencari metadata referensi ilmiah "
                "dari sumber yang sudah terhubung dengan aplikasi. "
                "Referensi tidak dibuat atau dikarang oleh AI."
            )

            kata_kunci_ref = st.text_input(
                "Kata Kunci Pencarian Referensi",
                value=judul_tugas,
                placeholder="Contoh: kepemimpinan pendidikan Islam",
                key="kata_kunci_ref_tugas"
            )

            jumlah_ref = st.slider(
                "Jumlah referensi yang dicari",
                min_value=5,
                max_value=20,
                value=10,
                step=1,
                key="jumlah_ref_tugas"
            )

            if "hasil_ref_tugas" not in st.session_state:
                st.session_state.hasil_ref_tugas = []

            if st.button(
                "🔎 Cari Referensi Ilmiah",
                key="cari_ref_tugas",
                use_container_width=True
            ):

                if not kata_kunci_ref.strip():

                    st.warning(
                        "Masukkan judul/tema atau kata kunci terlebih dahulu."
                    )

                else:

                    with st.spinner(
                        "Mencari referensi ilmiah..."
                    ):

                        hasil_ref = cari_multi_sumber(
                            kata_kunci_ref,
                            jumlah_ref
                        )

                    st.session_state.hasil_ref_tugas = hasil_ref

                    if hasil_ref:
                        st.success(
                            f"✅ Ditemukan {len(hasil_ref)} "
                            "referensi yang dapat diperiksa."
                        )
                    else:
                        st.warning(
                            "Referensi belum ditemukan. "
                            "Coba gunakan kata kunci yang lebih spesifik."
                        )

            hasil_ref = st.session_state.get(
                "hasil_ref_tugas",
                []
            )

            referensi_dipilih = []

            if hasil_ref:

                st.markdown("### 📚 Hasil Pencarian")

                st.caption(
                    "Centang referensi yang ingin digunakan dalam tugas."
                )

                for i, ref in enumerate(hasil_ref):

                    judul_ref = (
                        ref.get("Judul")
                        or "Tanpa judul"
                    )

                    tahun_ref = (
                        ref.get("Tahun")
                        or "Tanpa tahun"
                    )

                    sumber_ref = (
                        ref.get("Sumber")
                        or "Sumber tidak diketahui"
                    )

                    status_ref = (
                        ref.get("Status")
                        or "Belum terverifikasi"
                    )

                    label_ref = (
                        f"{judul_ref} "
                        f"({tahun_ref})"
                    )

                    pilih_ref = st.checkbox(
                        label_ref,
                        key=f"pilih_ref_tugas_{i}"
                    )

                    st.caption(
                        f"{sumber_ref} | {status_ref}"
                    )

                    if ref.get("Penulis"):
                        st.write(
                            f"**Penulis:** {ref.get('Penulis')}"
                        )

                    if ref.get("Jurnal"):
                        st.write(
                            f"**Jurnal/Sumber:** "
                            f"{ref.get('Jurnal')}"
                        )

                    if ref.get("DOI"):
                        st.write(
                            f"**DOI:** {ref.get('DOI')}"
                        )

                    if pilih_ref:
                        referensi_dipilih.append(ref)

                    st.divider()

            # =================================================
            # REFERENSI DARI BANK REFERENSI
            # =================================================
            st.markdown("### 📚 Bank Referensi")

            bank_ref = st.session_state.get(
                "bank_referensi",
                []
            )

            if bank_ref:

                st.caption(
                    "Referensi yang sebelumnya sudah disimpan "
                    "juga dapat digunakan."
                )

                pilihan_bank = []

                for i, ref in enumerate(bank_ref):

                    label = (
                        f"{ref.get('Judul', 'Tanpa judul')} "
                        f"({ref.get('Tahun', 'n.d.')})"
                    )

                    if st.checkbox(
                        label,
                        key=f"bank_ref_tugas_{i}"
                    ):
                        pilihan_bank.append(ref)

                for ref in pilihan_bank:
                    if (
                        kunci_ref(ref)
                        not in [
                            kunci_ref(x)
                            for x in referensi_dipilih
                        ]
                    ):
                        referensi_dipilih.append(ref)

            else:

                st.info(
                    "Bank Referensi masih kosong. "
                    "Anda dapat mencari referensi di atas."
                )

            # =================================================
            # SIMPAN REFERENSI TERPILIH
            # =================================================
            if referensi_dipilih:

                st.success(
                    f"{len(referensi_dipilih)} referensi "
                    "dipilih untuk tugas ini."
                )

                if st.button(
                    "💾 Simpan Referensi Terpilih ke Bank Referensi",
                    key="simpan_ref_tugas",
                    use_container_width=True
                ):

                    jumlah_baru = 0

                    for ref in referensi_dipilih:
                        if tambah_bank_referensi(ref):
                            jumlah_baru += 1

                    if jumlah_baru:
                        st.success(
                            f"✅ {jumlah_baru} referensi baru "
                            "masuk ke Bank Referensi."
                        )
                    else:
                        st.info(
                            "Referensi tersebut sudah ada "
                            "di Bank Referensi."
                        )

            # =================================================
            # PENYUSUNAN TUGAS
            # =================================================
            st.divider()

            st.markdown("## 🤖 Penyusunan Tugas dengan AI")

            if jenis_tugas == "Makalah":

                st.info(
                    "Makalah akan disusun lengkap mulai dari cover, "
                    "kata pengantar, daftar isi, BAB I, BAB II, BAB III, "
                    "catatan kaki/sitasi, sampai daftar pustaka."
                )

            elif jenis_tugas == "Paper":

                st.info(
                    "Paper akan disusun dalam format akademik "
                    "sesuai tema dan bahan yang tersedia."
                )

            elif jenis_tugas == "Resume":

                st.info(
                    "Resume akan merangkum bahan secara sistematis "
                    "tanpa mengubah substansi utama."
                )

            elif jenis_tugas == "Review Jurnal":

                st.info(
                    "Review akan membahas identitas artikel, masalah, "
                    "metode, temuan, kekuatan, kelemahan, dan kesimpulan."
                )

            elif jenis_tugas == "Review Buku":

                st.info(
                    "Review akan membahas identitas buku, isi utama, "
                    "kelebihan, kekurangan, analisis, dan kesimpulan."
                )

            elif jenis_tugas == "Critical Review":

                st.info(
                    "Critical review akan menekankan analisis kritis, "
                    "argumentasi, kekuatan, kelemahan, dan relevansi."
                )

            elif jenis_tugas == "Studi Kasus":

                st.info(
                    "Studi kasus akan disusun dari masalah, analisis, "
                    "alternatif solusi, rekomendasi, dan kesimpulan."
                )

            elif jenis_tugas == "Presentasi":

                st.info(
                    "Hasil akan disusun menjadi kerangka presentasi "
                    "yang ringkas dan sistematis."
                )

            # =================================================
            # GENERATE AI
            # =================================================
            if st.button(
                "✨ Susun Tugas Lengkap",
                type="primary",
                key="generate_tugas_lengkap",
                use_container_width=True
            ):

                if not judul_tugas.strip():

                    st.warning(
                        "Masukkan judul atau tema tugas terlebih dahulu."
                    )

                else:

                    konteks_tugas = (
                        f"MATA KULIAH:\n{mata_kuliah}\n\n"
                        f"JENIS TUGAS:\n{jenis_tugas}\n\n"
                        f"JUDUL/TEMA:\n{judul_tugas}\n\n"
                    )

                    if bahan_teks.strip():

                        konteks_tugas += (
                            "BAHAN DARI PENGGUNA:\n"
                            + bahan_teks[:60000]
                            + "\n\n"
                        )

                    if jenis_tugas == "Makalah":

                        instruksi_ai = """
Susun MAKALAH AKADEMIK LENGKAP berdasarkan judul,
bahan pengguna, dan referensi yang tersedia.

STRUKTUR WAJIB:

1. COVER
   - Judul makalah
   - Mata kuliah
   - Sediakan tempat untuk nama mahasiswa
   - Sediakan tempat untuk NIM
   - Sediakan tempat untuk nama dosen
   - Sediakan tempat untuk program studi/institusi
   - Tahun

2. KATA PENGANTAR

3. DAFTAR ISI

4. BAB I PENDAHULUAN
   A. Latar Belakang
   B. Rumusan Masalah
   C. Tujuan Penulisan

5. BAB II PEMBAHASAN
   - Buat subbab berdasarkan fokus pembahasan.
   - Pembahasan harus akademik, sistematis, dan mendalam.
   - Hubungkan teori dengan tema makalah.
   - Gunakan referensi yang tersedia untuk mendukung klaim akademik.

6. BAB III PENUTUP
   A. Kesimpulan
   B. Saran

7. CATATAN KAKI bila gaya sitasi menggunakan
   Chicago Notes & Bibliography.

8. DAFTAR PUSTAKA

ATURAN AKADEMIK WAJIB:
- Jangan mengarang sumber.
- Jangan mengarang DOI.
- Jangan mengarang nama penulis.
- Jangan mengarang nomor halaman.
- Jangan membuat kutipan langsung jika teks asli sumber
  tidak tersedia.
- Bila nomor halaman tidak tersedia, beri tanda
  [halaman perlu verifikasi].
- Gunakan hanya referensi yang diberikan sistem.
- Setiap sumber dalam daftar pustaka harus benar-benar
  digunakan dalam naskah.
- Jangan memasukkan sumber yang tidak digunakan.
- Jangan membuat data penelitian fiktif.
- Gunakan bahasa Indonesia akademik tingkat pascasarjana.
- Hasil harus berupa makalah utuh dan dapat diedit,
  bukan sekadar jawaban singkat.
"""

                    elif jenis_tugas == "Paper":

                        instruksi_ai = """
Susun paper akademik lengkap berdasarkan judul,
bahan, dan referensi yang tersedia.
Bangun argumentasi yang sistematis, analitis,
dan menggunakan referensi yang benar-benar tersedia.
Jangan mengarang referensi, DOI, kutipan, halaman,
atau data penelitian.
"""

                    elif jenis_tugas == "Resume":

                        instruksi_ai = """
Susun resume akademik yang sistematis.
Pertahankan gagasan utama bahan.
Jangan menambahkan fakta atau sumber yang tidak tersedia.
Buat bagian pokok bahasan, uraian inti,
dan kesimpulan.
"""

                    elif jenis_tugas == "Review Jurnal":

                        instruksi_ai = """
Susun review jurnal akademik yang meliputi:
identitas artikel jika tersedia, masalah penelitian,
tujuan, teori, metode, hasil, kekuatan,
kelemahan, analisis kritis, relevansi,
dan kesimpulan.
Jangan mengarang informasi yang tidak ada
dalam bahan atau referensi.
"""

                    elif jenis_tugas == "Review Buku":

                        instruksi_ai = """
Susun review buku akademik yang meliputi:
identitas buku jika tersedia, pokok isi,
gagasan utama, kekuatan, kelemahan,
analisis kritis, relevansi, dan kesimpulan.
Jangan mengarang isi buku yang tidak tersedia.
"""

                    elif jenis_tugas == "Critical Review":

                        instruksi_ai = """
Susun critical review akademik.
Bedakan ringkasan isi dengan analisis kritis.
Bahas argumentasi, kekuatan, kelemahan,
relevansi, dan kontribusi.
Gunakan hanya bahan dan referensi yang tersedia.
"""

                    elif jenis_tugas == "Studi Kasus":

                        instruksi_ai = """
Susun studi kasus akademik dengan struktur:
latar kasus, identifikasi masalah,
analisis masalah, landasan teori,
alternatif solusi, solusi yang direkomendasikan,
alasan pemilihan solusi, dan kesimpulan.
Jangan menciptakan fakta kasus yang tidak diberikan.
"""

                    elif jenis_tugas == "Laporan":

                        instruksi_ai = """
Susun laporan akademik secara sistematis
berdasarkan bahan yang tersedia.
Gunakan struktur pendahuluan, isi/hasil,
pembahasan, kesimpulan, dan rekomendasi
sesuai konteks tugas.
Jangan menciptakan data.
"""

                    elif jenis_tugas == "Mini Riset":

                        instruksi_ai = """
Susun kerangka mini riset akademik berdasarkan
informasi yang tersedia.
Jangan menciptakan data penelitian atau hasil penelitian.
Jika data belum tersedia, buat rancangan analisis
dan tandai bagian yang masih harus diisi pengguna.
"""

                    elif jenis_tugas == "Presentasi":

                        instruksi_ai = """
Susun materi presentasi akademik.
Buat urutan slide yang logis mulai dari judul,
latar belakang, pokok pembahasan,
analisis, kesimpulan, dan referensi.
Isi setiap slide harus ringkas dan siap dipindahkan
ke PowerPoint.
"""

                    else:

                        instruksi_ai = (
                            "Susun tugas akademik lengkap berdasarkan "
                            "judul, bahan, dan referensi yang tersedia. "
                            "Jangan mengarang sumber atau data."
                        )

                    panel_ai_penulisan(
                        konteks_tugas,
                        jenis_tugas,
                        instruksi_ai,
                        referensi_dipilih,
                        "tugas_mata_kuliah"
                    )

            # =================================================
            # HASIL TUGAS
            # =================================================
            if st.session_state.get(
                "hasil_penulisan_ai"
            ):

                st.divider()

                st.markdown("## 📄 Hasil Tugas")

                hasil_edit = st.text_area(
                    "Hasil dapat diedit",
                    st.session_state.hasil_penulisan_ai,
                    height=700,
                    key="hasil_tugas_mata_kuliah"
                )

                st.session_state.naskah_aktif = hasil_edit

                st.caption(
                    "Periksa kembali nama, NIM, dosen, institusi, "
                    "kutipan, halaman sumber, dan ketentuan tugas "
                    "sebelum digunakan sebagai naskah final."
                )

        # ====================================================
        # RPS / MODUL
        # ====================================================
        elif bagian == "📄 RPS / Modul":

            st.markdown("## 📄 RPS / Modul")

            st.info(
                "Bagian ini untuk menyimpan atau membaca RPS dan modul "
                "mata kuliah sebagai bahan pendukung. "
                "RPS tidak perlu dimasukkan setiap kali membuat tugas."
            )

            files_rps = st.file_uploader(
                "Unggah RPS / Modul",
                type=["pdf", "docx", "txt"],
                accept_multiple_files=True,
                key="upload_rps_mata_kuliah"
            )

            if files_rps:

                for file in files_rps:

                    st.write(
                        f"📄 **{file.name}** — "
                        f"{format_ukuran(file.size)}"
                    )

                    if st.button(
                        f"💾 Simpan {file.name} ke Bank Karya",
                        key=f"simpan_rps_{file.name}"
                    ):

                        berhasil = simpan_karya(
                            file.name,
                            f"RPS/Modul - {mata_kuliah}",
                            format_ukuran(file.size)
                        )

                        if berhasil:
                            st.success(
                                "Dokumen tercatat di Bank Karya."
                            )
                        else:
                            st.info(
                                "Dokumen sudah tercatat."
                            )

        # ====================================================
        # MATERI PERKULIAHAN
        # ====================================================
        elif bagian == "📚 Materi Perkuliahan":

            st.markdown("## 📚 Materi Perkuliahan")

            files_materi = st.file_uploader(
                "Unggah materi kuliah",
                type=["pdf", "docx", "txt", "pptx"],
                accept_multiple_files=True,
                key="upload_materi_mata_kuliah"
            )

            if files_materi:

                for file in files_materi:

                    st.write(
                        f"📄 **{file.name}** — "
                        f"{format_ukuran(file.size)}"
                    )

        # ====================================================
        # CATATAN DOSEN
        # ====================================================
        elif bagian == "📌 Catatan Dosen":

            st.markdown("## 📌 Catatan Dosen")

            st.text_area(
                "Catatan, arahan, atau revisi dari dosen",
                height=300,
                key="catatan_dosen_mata_kuliah"
            )

        # ====================================================
        # REFERENSI MATA KULIAH
        # ====================================================
        elif bagian == "📖 Referensi Mata Kuliah":

            st.markdown("## 📖 Referensi Mata Kuliah")

            st.info(
                "Referensi mata kuliah menggunakan Bank Referensi "
                "yang sama dengan menu Literatur & Referensi."
            )

            refs_mk = st.session_state.get(
                "bank_referensi",
                []
            )

            if refs_mk:

                for i, ref in enumerate(refs_mk, 1):

                    st.markdown(
                        f"**{i}. {ref.get('Judul', 'Tanpa judul')}**"
                    )

                    st.write(
                        format_referensi(ref)
                    )

                    st.caption(
                        ref.get(
                            "Status",
                            "Status belum tersedia"
                        )
                    )

                    st.divider()

            else:

                st.info(
                    "Bank Referensi masih kosong. "
                    "Gunakan pencarian referensi pada Tugas Saya "
                    "atau menu Literatur & Referensi."
                )
# ============================================================
# ANALISIS KARYA AKADEMIK
# ============================================================
elif menu == "🔬 Analisis Karya Akademik":
    st.header("🔬 Analisis Karya Akademik")
    st.caption("Analisis dokumen, teks, video, atau audio dengan indikator bawaan opsional dan indikator manual tanpa batas.")

    if "indikator_manual_analisis" not in st.session_state:
        st.session_state.indikator_manual_analisis = []
    if "hasil_ai_gemini" not in st.session_state:
        st.session_state.hasil_ai_gemini = ""

    sumber = st.radio(
        "Sumber yang dianalisis",
        ["📄 Dokumen", "✍️ Teks langsung", "🎥 Video", "🎧 Audio"],
        horizontal=True,
        key="sumber_analisis_karya"
    )
    jenis = st.selectbox(
        "Jenis karya / tayangan",
        ["Deteksi Otomatis", "Tugas Kuliah", "Makalah", "Artikel Jurnal", "Buku", "Bab Buku", "Proposal", "Tesis", "Disertasi", "Laporan Penelitian", "Pembelajaran / Supervisi", "Seminar / Presentasi", "Regulasi", "Dokumen Lain"],
        key="jenis_karya_analisis"
    )

    bahan_teks = ""
    media_bytes = None
    media_mime = None
    nama_sumber = "Teks langsung"

    if sumber == "📄 Dokumen":
        file = st.file_uploader("Unggah dokumen", type=["pdf", "docx", "txt"], key="upload_analisis_karya_baru")
        if file is not None:
            nama_sumber = file.name
            bahan_teks = ekstrak_teks(file)
            if bahan_teks.startswith("ERROR:"):
                st.error(bahan_teks); bahan_teks = ""
            elif bahan_teks.strip():
                c1,c2,c3=st.columns(3)
                c1.metric("Jumlah Kata", len(bahan_teks.split()))
                c2.metric("Jumlah Karakter", len(bahan_teks))
                c3.metric("Ukuran File", format_ukuran(file.size))
                with st.expander("📖 Lihat teks hasil ekstraksi"):
                    st.text_area("Teks sumber", bahan_teks, height=300, key="preview_teks_analisis_baru")
                if st.button("💾 Simpan ke Bank Karya", key="simpan_bank_karya_analisis_baru"):
                    if simpan_karya(file.name, jenis, format_ukuran(file.size)):
                        st.success("Dokumen berhasil disimpan ke Bank Karya.")
                    else:
                        st.warning("Dokumen ini sudah tercatat pada proyek aktif.")
    elif sumber == "✍️ Teks langsung":
        bahan_teks = st.text_area("Tempel atau ketik teks yang akan dianalisis", height=350, key="teks_langsung_analisis")
    elif sumber == "🎥 Video":
        file = st.file_uploader("Unggah video", type=["mp4", "mov", "m4v", "webm", "mpeg", "mpg"], key="upload_video_analisis")
        if file is not None:
            nama_sumber=file.name; media_bytes=file.getvalue(); media_mime=getattr(file,"type",None) or "video/mp4"
            st.video(media_bytes)
            st.caption(f"Ukuran video: {format_ukuran(file.size)}. File besar akan diproses otomatis melalui Gemini Files API.")
    else:
        file = st.file_uploader("Unggah audio", type=["mp3", "wav", "m4a", "aac", "ogg", "flac"], key="upload_audio_analisis")
        if file is not None:
            nama_sumber=file.name; media_bytes=file.getvalue(); media_mime=getattr(file,"type",None) or "audio/mpeg"
            st.audio(media_bytes)
            st.caption(f"Ukuran audio: {format_ukuran(file.size)}. File besar akan diproses otomatis melalui Gemini Files API.")

    # Bersihkan hasil lama otomatis bila sumber dihapus, diganti, atau isi teks berubah.
    if bahan_teks.strip():
        _sig_isi = str(hash(bahan_teks))
    elif media_bytes is not None:
        _sig_isi = f"{nama_sumber}|{len(media_bytes)}"
    else:
        _sig_isi = "KOSONG"
    _sig_sumber = f"{sumber}|{_sig_isi}"
    if "signature_sumber_analisis" not in st.session_state:
        st.session_state.signature_sumber_analisis = _sig_sumber
    elif st.session_state.signature_sumber_analisis != _sig_sumber:
        st.session_state.hasil_ai_gemini = ""
        st.session_state.signature_sumber_analisis = _sig_sumber

    st.divider()
    st.subheader("📋 Indikator Analisis")
    st.caption("Indikator bawaan tidak dipilih otomatis. Centang hanya yang diperlukan. Indikator manual dapat ditambahkan tanpa batas aplikasi.")

    indikator_bawaan = [
        "Identitas dan Jenis Sumber", "Topik Utama", "Latar Belakang / Masalah", "Rumusan Masalah", "Tujuan",
        "Konsep / Landasan Teori", "Penelitian Terdahulu", "Metodologi", "Populasi / Sampel / Informan",
        "Instrumen", "Teknik Pengumpulan Data", "Teknik Analisis Data", "Temuan / Hasil", "Pembahasan", "Kesimpulan",
        "Keterbatasan", "Research Gap", "Novelty / Kebaruan", "Kontribusi Akademik", "Validitas / Kecukupan Referensi",
        "Kelebihan Karya", "Kekurangan Karya", "Kritik Akademik", "Rekomendasi Perbaikan", "Peluang Penelitian Lanjutan",
        "Relevansi dengan Penelitian Pengguna", "Bagaimana Sumber Ini Dapat Digunakan"
    ]
    if jenis == "Pembelajaran / Supervisi":
        indikator_bawaan += [
            "Tujuan Pembelajaran", "Kegiatan Pembukaan", "Penguasaan Materi", "Strategi / Metode Pembelajaran",
            "Aktivitas Guru", "Aktivitas Peserta Didik", "Komunikasi Pembelajaran", "Penggunaan Media",
            "Pertanyaan / HOTS", "Penerapan Deep Learning", "Integrasi KBC / Panca Cinta", "Asesmen Formatif",
            "Pengelolaan Kelas", "Umpan Balik", "Kegiatan Penutup", "Tindak Lanjut Supervisi"
        ]

    pilih_bawaan = st.multiselect("Indikator bawaan (opsional)", indikator_bawaan, default=[], key="indikator_bawaan_analisis_v2")

    st.markdown("#### ✍️ Indikator Manual")
    cman1,cman2=st.columns([4,1])
    with cman1:
        indikator_baru=st.text_input("Ketik indikator baru", key="indikator_manual_baru", placeholder="Contoh: Kesesuaian materi dengan tujuan pembelajaran")
    with cman2:
        st.write(""); st.write("")
        if st.button("➕ Tambah", key="tambah_indikator_manual", use_container_width=True):
            x=indikator_baru.strip()
            if x and x not in st.session_state.indikator_manual_analisis:
                st.session_state.indikator_manual_analisis.append(x); st.rerun()

    if st.session_state.indikator_manual_analisis:
        for i,item in enumerate(list(st.session_state.indikator_manual_analisis)):
            c1,c2=st.columns([8,1])
            with c1:
                nilai=st.text_input(f"Indikator manual {i+1}", value=item, key=f"edit_indikator_manual_{i}")
                st.session_state.indikator_manual_analisis[i]=nilai.strip() or item
            with c2:
                st.write(""); st.write("")
                if st.button("🗑️", key=f"hapus_indikator_manual_{i}", help="Hapus indikator"):
                    st.session_state.indikator_manual_analisis.pop(i); st.rerun()

    indikator_aktif = pilih_bawaan + [x for x in st.session_state.indikator_manual_analisis if x.strip()]
    st.caption(f"Indikator aktif: {len(indikator_aktif)}")
    mode_tambahan = st.checkbox("🤖 Jika tidak memilih indikator, izinkan AI membuat indikator yang paling relevan", value=True, key="ai_buat_indikator_otomatis")

    def _upload_media_gemini_files(api_key, media_data, mime_type, display_name):
        """Upload media besar ke Gemini Files API dan tunggu sampai siap dipakai."""
        start_url = "https://generativelanguage.googleapis.com/upload/v1beta/files?key=" + urllib.parse.quote(api_key)
        metadata = json.dumps({"file": {"display_name": display_name}}).encode("utf-8")
        req = urllib.request.Request(
            start_url, data=metadata, method="POST",
            headers={
                "Content-Type": "application/json",
                "X-Goog-Upload-Protocol": "resumable",
                "X-Goog-Upload-Command": "start",
                "X-Goog-Upload-Header-Content-Length": str(len(media_data)),
                "X-Goog-Upload-Header-Content-Type": mime_type,
            }
        )
        with urllib.request.urlopen(req, timeout=60) as resp:
            upload_url = resp.headers.get("X-Goog-Upload-URL") or resp.headers.get("x-goog-upload-url")
        if not upload_url:
            raise RuntimeError("Gemini Files API tidak memberikan URL upload.")

        req2 = urllib.request.Request(
            upload_url, data=media_data, method="POST",
            headers={
                "Content-Type": mime_type,
                "Content-Length": str(len(media_data)),
                "X-Goog-Upload-Offset": "0",
                "X-Goog-Upload-Command": "upload, finalize",
            }
        )
        with urllib.request.urlopen(req2, timeout=300) as resp:
            info = json.loads(resp.read().decode("utf-8"))
        f = info.get("file", {})
        name = f.get("name", "")
        uri = f.get("uri", "")
        state = f.get("state", "")
        if not name or not uri:
            raise RuntimeError("Upload media selesai tetapi metadata file Gemini tidak lengkap.")

        # Video biasanya perlu waktu pemrosesan. Poll sampai ACTIVE.
        deadline = time.time() + 300
        while state == "PROCESSING" and time.time() < deadline:
            time.sleep(5)
            get_url = "https://generativelanguage.googleapis.com/v1beta/" + name + "?key=" + urllib.parse.quote(api_key)
            with urllib.request.urlopen(get_url, timeout=60) as resp:
                f = json.loads(resp.read().decode("utf-8"))
            uri = f.get("uri", uri)
            state = f.get("state", state)
        if state == "FAILED":
            raise RuntimeError("Gemini gagal memproses file media.")
        if state == "PROCESSING":
            raise RuntimeError("Pemrosesan media belum selesai setelah 5 menit. Coba Generate Analisis AI kembali.")
        return uri

    def _analisis_multimodal_akademik(teks_sumber, media_data, mime_type, nama, jenis_sumber, indikator):
        try:
            api_key=st.secrets["GEMINI_API_KEY"]
        except Exception:
            return {"sukses":False,"hasil":"","error":"GEMINI_API_KEY belum ditemukan di Streamlit Secrets.","model":""}
        daftar = "\n".join(f"{i+1}. {x}" for i,x in enumerate(indikator)) if indikator else "Tentukan sendiri indikator paling relevan berdasarkan isi sumber."
        prompt=f'''Anda adalah Asisten Akademik AI yang kritis dan teliti. Analisis sumber berikut hanya berdasarkan isi yang benar-benar tersedia.

Jenis sumber: {jenis_sumber}
Jenis karya/tayangan: {jenis}
Nama sumber: {nama}

INDIKATOR YANG HARUS DINILAI:
{daftar}

ATURAN WAJIB UNTUK SETIAP INDIKATOR:
1. Indikator tetap harus ditampilkan walaupun unsur tidak ditemukan.
2. Gunakan tepat salah satu status: ✅ ADA, ❌ TIDAK ADA, atau ⚠️ TIDAK EKSPLISIT / PERLU INTERPRETASI.
3. Jika ADA, tuliskan bukti/narasi yang ditemukan. Bila lokasi/halaman/waktu dapat dikenali, sebutkan; jangan mengarang.
4. Jika TIDAK ADA, tulis bahwa unsur tidak ditemukan. Jangan membuat isi pengganti.
5. Jika TIDAK EKSPLISIT, pisahkan dengan jelas ANALISIS AI dari pernyataan sumber.
6. Untuk setiap indikator berikan: Status; Bukti/Narasi; Komentar Analisis; Kelebihan; Kekurangan; Rekomendasi. Jika suatu unsur penilaian tidak relevan, tulis Tidak relevan, bukan mengarang.
7. Setelah semua indikator, berikan: Ringkasan Kelebihan Utama; Kekurangan Utama; Kritik Akademik; Rekomendasi Prioritas; Kesimpulan; dan Bagaimana sumber ini dapat digunakan.
8. Jangan mengarang nama, data, teori, metode, hasil, referensi, DOI, halaman, kutipan, research gap, novelty, atau fakta.
9. Untuk research gap/novelty yang hanya merupakan inferensi, beri label ANALISIS AI dan jangan menyatakannya sebagai klaim penulis.
10. Gunakan bahasa Indonesia akademik yang jelas dan konkret.
'''
        if teks_sumber:
            prompt += "\nTEKS SUMBER:\n--------------------\n" + teks_sumber[:70000] + "\n--------------------\n"
        parts=[{"text":prompt}]
        if media_data is not None:
            try:
                if len(media_data) > 18*1024*1024:
                    file_uri = _upload_media_gemini_files(api_key, media_data, mime_type, nama)
                    parts.append({"file_data":{"mime_type":mime_type,"file_uri":file_uri}})
                else:
                    parts.append({"inline_data":{"mime_type":mime_type,"data":base64.b64encode(media_data).decode("ascii")}})
            except Exception as e:
                return {"sukses":False,"hasil":"","error":"Media belum berhasil diproses melalui Gemini Files API. "+str(e),"model":""}
        payload={"contents":[{"parts":parts}],"generationConfig":{"temperature":0.2,"maxOutputTokens":8192}}
        model_ids=["gemini-3.8-flash","gemini-3.5-flash-lite"]
        transient={408,429,500,502,503,504}; err=""
        for mi,model_id in enumerate(model_ids):
            url="https://generativelanguage.googleapis.com/v1beta/models/"+model_id+":generateContent?key="+api_key
            for attempt in range(4):
                req=urllib.request.Request(url,data=json.dumps(payload).encode("utf-8"),headers={"Content-Type":"application/json"},method="POST")
                try:
                    with urllib.request.urlopen(req,timeout=240) as resp: data=json.loads(resp.read().decode("utf-8"))
                    out="\n".join(p.get("text","") for p in ((data.get("candidates") or [{}])[0].get("content",{}).get("parts",[])) if isinstance(p,dict) and p.get("text")).strip()
                    if out: return {"sukses":True,"hasil":out,"error":"","model":model_id,"fallback":mi>0}
                    err=f"{model_id} merespons tetapi hasil kosong."; break
                except urllib.error.HTTPError as e:
                    try: detail=e.read().decode("utf-8")
                    except Exception: detail=""
                    err=f"Gemini HTTP {e.code}. {detail}".strip()
                    if e.code not in transient: return {"sukses":False,"hasil":"","error":err,"model":model_id}
                    if attempt<3: time.sleep((2**(attempt+1))+random.uniform(.2,1.0)); continue
                    break
                except Exception as e:
                    err=str(e)
                    if attempt<3: time.sleep((2**(attempt+1))+random.uniform(.2,1.0)); continue
                    break
        return {"sukses":False,"hasil":"","error":"Analisis belum berhasil setelah retry/fallback. "+err,"model":""}

    siap = bool(bahan_teks.strip()) or media_bytes is not None
    if st.button("🤖 Generate Analisis AI", type="primary", use_container_width=True, disabled=not siap, key="generate_analisis_karya_lengkap"):
        if not indikator_aktif and not mode_tambahan:
            st.warning("Pilih minimal satu indikator atau aktifkan pembuatan indikator otomatis oleh AI.")
        else:
            with st.spinner("AI sedang menganalisis sumber dan setiap indikator..."):
                hasil=_analisis_multimodal_akademik(bahan_teks,media_bytes,media_mime,nama_sumber,sumber,indikator_aktif)
            if hasil.get("sukses"):
                st.session_state.hasil_ai_gemini=hasil["hasil"]
                st.success("✅ Analisis AI berhasil.")
                st.caption(f"Model AI: {hasil.get('model','')}")
            else:
                st.error("❌ Analisis AI belum berhasil.")
                st.warning(hasil.get("error","Kesalahan tidak diketahui."))

    st.markdown("### 📝 Hasil Analisis")
    if st.session_state.hasil_ai_gemini:
        edit=st.text_area("Hasil analisis dapat diedit", value=st.session_state.hasil_ai_gemini, height=700, key="editor_hasil_ai_gemini_baru")
        st.session_state.hasil_ai_gemini=edit
        c_unduh, c_hapus = st.columns([3, 1])
        with c_unduh:
            st.download_button("📥 Unduh Hasil Analisis (.txt)", edit.encode("utf-8"), "hasil_analisis_akademik.txt", "text/plain", use_container_width=True, key="download_hasil_analisis_baru")
        with c_hapus:
            if st.button("🗑️ Hapus Hasil", use_container_width=True, key="hapus_hasil_analisis_karya"):
                st.session_state.hasil_ai_gemini = ""
                st.rerun()
    else:
        st.info("Unggah/masukkan sumber lalu jalankan Generate Analisis AI.")

# ============================================================
# LITERATUR & REFERENSI
# ============================================================
elif menu == "🔎 Literatur & Referensi":
    st.header("🔎 Literatur, Sitasi & Library Referensi")
    c1,c2=st.columns([2,1])
    with c1: mode_ref=st.radio("Mode Referensi",["🤖 Otomatis Terverifikasi","🔍 Verifikasi Dulu","📚 Referensi Saya"],horizontal=True)
    with c2:
        gaya_list=["Chicago Notes & Bibliography","Turabian Notes-Bibliography","OSCOLA","APA 7","Harvard","MLA","IEEE","Vancouver","AMA","ACS","CSE","APSA","Pedoman Kampus/Jurnal","Custom"]
        st.session_state.gaya_sitasi=st.selectbox("Gaya sitasi default",gaya_list,index=gaya_list.index(st.session_state.gaya_sitasi))
    st.caption("Chicago Notes & Bibliography menjadi default. Artikel jurnal tetap mengikuti gaya rumah jurnal/template yang diunggah.")
    bagian_ref = st.radio(
        "Bagian Literatur & Referensi",
        ["🔎 Cari Terintegrasi","🌐 Sumber Online","📤 Unggah Referensi","📚 Library","📝 Generasi Sitasi & Daftar Pustaka","✍️ Pakai di Naskah","🔧 Perbaiki Footnote & Kutipan","✅ Audit Sitasi"],
        key="bagian_literatur_referensi"
    )

    if bagian_ref == "🔎 Cari Terintegrasi":
        q=st.text_input("Topik / judul / kata kunci",key="q_ref")
        if st.button("🔎 Cari 4 Sumber Terintegrasi",type="primary",disabled=not bool(q.strip())):
            with st.spinner("Mencari Crossref, OpenAlex, Semantic Scholar, dan Library of Congress..."): st.session_state.hasil_cari_ref=cari_multi_sumber(q,12)
        hasil=st.session_state.get("hasil_cari_ref",[])
        if hasil:
            st.success(f"Ditemukan {len(hasil)} kandidat unik dari sumber terintegrasi.")
            st.caption("Status metadata menunjukkan asal verifikasi/identifikasi. Referensi tanpa DOI tetap harus diperiksa sebelum dipakai sebagai sumber final.")
            for i,r in enumerate(hasil):
                with st.expander(f"{i+1}. {r['Judul']} ({r['Tahun']}) — {r['Sumber']}"):
                    st.write(f"**Penulis:** {r['Penulis'] or '-'}")
                    st.write(f"**Jurnal:** {r['Jurnal'] or '-'}")
                    st.write(f"**DOI:** {r['DOI'] or 'Belum tersedia'}")
                    st.write(r["Status"])
                    if st.button("➕ Simpan ke Library",key=f"addref_new_{i}"): st.success("Disimpan." if tambah_bank_referensi(r) else "Sudah ada di Library.")
                    if r.get("DOI"): st.link_button("🔗 Buka DOI","https://doi.org/"+r["DOI"])

    elif bagian_ref == "🌐 Sumber Online":
        st.subheader("🌐 Perpustakaan & Sumber Referensi Online")
        oq=st.text_input("Kata kunci pencarian",key="q_online")
        st.info("Pencarian langsung aplikasi: Crossref, OpenAlex, Semantic Scholar, dan Library of Congress. Sumber lain dibuka melalui portal resminya. Login, lisensi, dan hak akses perpustakaan tetap dihormati.")
        kategori_pilih=st.multiselect("Wilayah sumber",["🇮🇩 Nasional","🏢 Daerah/Provinsi","🌍 Internasional"],
                                     default=["🇮🇩 Nasional","🏢 Daerah/Provinsi","🌍 Internasional"],key="kategori_sumber_online")
        sumber=[x for x in sumber_online_default() if x[0] in kategori_pilih]
        for kategori,nama,url,ket in sumber:
            target=url.replace("{q}",urllib.parse.quote_plus(oq.strip())) if "{q}" in url else url
            a,b=st.columns([4,1]); a.write(f"{kategori} **{nama}** — {ket}"); b.link_button("Buka / Cari",target,use_container_width=True)
        st.divider(); st.subheader("➕ Tambahkan Perpustakaan / Repository Sendiri")
        nm=st.text_input("Nama sumber",key="src_name"); ur=st.text_input("Link katalog/repository",placeholder="https://...",key="src_url")
        if st.button("💾 Simpan Sumber",disabled=not(nm.strip() and ur.strip())):
            if ur.startswith(("http://","https://")):
                item={"Nama":nm.strip(),"URL":ur.strip()}
                if item not in st.session_state.sumber_online_user: st.session_state.sumber_online_user.append(item)
                st.success("Sumber ditambahkan untuk sesi ini.")
            else: st.error("Link harus diawali http:// atau https://")
        for x in st.session_state.sumber_online_user:
            a,b=st.columns([3,1]); a.write("**"+x["Nama"]+"**"); b.link_button("Buka",x["URL"],use_container_width=True)

    elif bagian_ref == "📤 Unggah Referensi":
        uprefs=st.file_uploader("Unggah satu atau banyak PDF/DOCX/TXT referensi",type=["pdf","docx","txt"],accept_multiple_files=True,key="upload_refs")
        st.checkbox("Utamakan referensi yang saya unggah",value=True,key="prioritas_upload")
        st.caption("Otomatis: baca dokumen → cari DOI → verifikasi Crossref. Jika DOI tidak terbaca, Gemini mengekstrak metadata lalu judul diverifikasi kembali.")
        if st.button("📥 Baca, Verifikasi & Masukkan ke Library",type="primary",disabled=not bool(uprefs)):
            with st.spinner("Membaca dan memverifikasi metadata..."): n,lap=unggah_referensi_ke_bank(uprefs)
            st.session_state.laporan_upload_ref=lap; st.success(f"{n} referensi baru masuk Library.")
        if st.session_state.get("laporan_upload_ref"): st.dataframe(pd.DataFrame(st.session_state.laporan_upload_ref),use_container_width=True,hide_index=True)

    elif bagian_ref == "📚 Library":
        refs=st.session_state.bank_referensi
        if refs:
            df=pd.DataFrame(refs); kol=[x for x in ["Judul","Penulis","Tahun","Jurnal","DOI","Sumber","Status"] if x in df.columns]
            st.dataframe(df[kol],use_container_width=True,hide_index=True)

            st.markdown("#### 🗑️ Kelola / Hapus Referensi")
            st.caption("Menghapus di Akademia AI tidak menghapus referensi yang sudah diimpor ke Zotero, Mendeley, EndNote, atau file ekspor yang sudah disimpan.")

            # Hapus per item tetap tersedia.
            with st.expander("🗑️ Hapus satu referensi"):
                for i, r in enumerate(list(refs)):
                    cjudul, chapus = st.columns([8,2])
                    cjudul.write(f"**{i+1}. {r.get('Judul','Tanpa judul')}** ({r.get('Tahun','')})")
                    if chapus.button("🗑️ Hapus", key=f"hapus_ref_item_{i}", use_container_width=True):
                        st.session_state.bank_referensi.pop(i)
                        st.success("Referensi dihapus dari Library Akademia AI.")
                        st.rerun()

            opsi_hapus = [
                f"{i+1}. {r.get('Judul','Tanpa judul')} ({r.get('Tahun','')})"
                for i, r in enumerate(refs)
            ]
            pilih_semua = st.checkbox("☑️ Pilih semua referensi untuk penghapusan massal", key="pilih_semua_hapus_ref")
            if pilih_semua:
                pilih_hapus = opsi_hapus
                st.caption(f"{len(pilih_hapus)} referensi dipilih.")
            else:
                pilih_hapus = st.multiselect(
                    "Pilih beberapa referensi yang akan dihapus",
                    opsi_hapus,
                    key="pilih_hapus_referensi"
                )

            konfirmasi_hapus = st.checkbox(
                "Saya mengerti referensi terpilih akan dihapus dari Library Akademia AI.",
                key="konfirmasi_hapus_massal",
                value=False
            )
            if st.button(
                "🗑️ Hapus Referensi Terpilih",
                disabled=not bool(pilih_hapus) or not konfirmasi_hapus,
                key="hapus_referensi_library"
            ):
                indeks_hapus = {opsi_hapus.index(x) for x in pilih_hapus}
                st.session_state.bank_referensi = [
                    r for i, r in enumerate(refs) if i not in indeks_hapus
                ]
                st.success(f"{len(indeks_hapus)} referensi dihapus dari Library Akademia AI.")
                st.rerun()

            st.markdown("#### 🔄 Pengelola & Ekspor Referensi")
            peta_manager = {
                "Zotero": "RIS (.ris)",
                "Mendeley": "RIS (.ris)",
                "EndNote": "EndNote Tagged (.enw)",
                "RefWorks": "RIS (.ris)",
                "Paperpile": "RIS (.ris)",
                "Citavi": "RIS (.ris)",
                "BibTeX / LaTeX / JabRef": "BibTeX (.bib)",
                "Excel / Spreadsheet": "CSV (.csv)",
                "Lainnya / format universal": "RIS (.ris)"
            }
            ec1,ec2=st.columns([2,1])
            with ec1:
                manager=st.selectbox("Pilih aplikasi tujuan", list(peta_manager.keys()), key="reference_manager")
            with ec2:
                # Jangan gunakan text_input ber-key untuk nilai turunan karena Session State dapat menahan nilai lama.
                st.markdown("**Format otomatis**")
                st.info(peta_manager[manager])

            opsi_ref=[f"{i+1}. {r.get('Judul','Tanpa judul')} ({r.get('Tahun','')})" for i,r in enumerate(refs)]
            cakupan=st.radio("Referensi yang diekspor", ["Semua Referensi","Referensi yang Dipilih"], horizontal=True, key="cakupan_ekspor_ref")
            refs_ekspor=refs
            if cakupan=="Referensi yang Dipilih":
                pilihan_ekspor=st.multiselect("Pilih referensi", opsi_ref, key="pilihan_ekspor_ref")
                idx={opsi_ref.index(x) for x in pilihan_ekspor}
                refs_ekspor=[r for i,r in enumerate(refs) if i in idx]

            if st.button("🤖 Periksa Metadata dengan AI Sebelum Ekspor", disabled=not bool(refs_ekspor), key="ai_periksa_sebelum_ekspor"):
                daftar_ai="\n".join(
                    f"{i+1}. {format_referensi(r)} | DOI: {r.get('DOI','')} | STATUS: {r.get('Status','')}"
                    for i,r in enumerate(refs_ekspor[:100])
                )
                prompt_ai = """Audit metadata referensi berikut sebelum ekspor.
Jangan mengarang metadata, DOI, halaman, penulis, tahun, jurnal, atau URL.
Untuk setiap item beri status:
✅ SIAP EKSPOR
⚠️ PERLU KONFIRMASI
❌ DATA PENTING TIDAK DITEMUKAN
Jika ada kekurangan, berikan saran pemeriksaan/pencarian. Jangan mengganti data secara otomatis.
REFERENSI:
""" + daftar_ai
                with st.spinner("AI memeriksa metadata tanpa mengubah Library..."):
                    h_ai=panggil_gemini(prompt_ai)
                if h_ai["sukses"]:
                    st.session_state.hasil_ai_ekspor_ref=h_ai["hasil"]
                else:
                    st.error(h_ai["error"])
            if st.session_state.get("hasil_ai_ekspor_ref"):
                st.text_area("Hasil pemeriksaan AI sebelum ekspor", st.session_state.hasil_ai_ekspor_ref, height=320, key="hasil_ai_ekspor_ref_tampil")

            fmt=peta_manager[manager]
            if fmt.startswith("RIS"):
                data_ekspor=ekspor_ris(refs_ekspor).encode("utf-8"); nama_ekspor="library_referensi.ris"; mime="application/x-research-info-systems"
            elif fmt.startswith("EndNote"):
                data_ekspor=ekspor_endnote_tagged(refs_ekspor).encode("utf-8"); nama_ekspor="library_referensi.enw"; mime="text/plain"
            elif fmt.startswith("BibTeX"):
                data_ekspor=ekspor_bibtex(refs_ekspor).encode("utf-8"); nama_ekspor="library_referensi.bib"; mime="application/x-bibtex"
            else:
                data_ekspor=ekspor_csv_referensi(refs_ekspor).encode("utf-8-sig"); nama_ekspor="library_referensi.csv"; mime="text/csv"
            st.download_button(f"📥 Ekspor ke {manager} — {fmt}", data_ekspor, nama_ekspor, mime, use_container_width=True, disabled=not bool(refs_ekspor), key="ekspor_manager_otomatis")
            with st.expander("⚙️ Format ekspor manual"):
                c1,c2,c3,c4=st.columns(4)
                with c1: st.download_button("RIS",ekspor_ris(refs_ekspor).encode("utf-8"),"library_referensi.ris","application/x-research-info-systems",disabled=not bool(refs_ekspor),key="exp_ris_manual")
                with c2: st.download_button("BibTeX",ekspor_bibtex(refs_ekspor).encode("utf-8"),"library_referensi.bib","application/x-bibtex",disabled=not bool(refs_ekspor),key="exp_bib_manual")
                with c3: st.download_button("EndNote",ekspor_endnote_tagged(refs_ekspor).encode("utf-8"),"library_referensi.enw","text/plain",disabled=not bool(refs_ekspor),key="exp_enw_manual")
                with c4: st.download_button("CSV",ekspor_csv_referensi(refs_ekspor).encode("utf-8-sig"),"library_referensi.csv","text/csv",disabled=not bool(refs_ekspor),key="exp_csv_manual")
            st.caption("Metadata yang belum terverifikasi tetap ditandai. Akademia AI tidak membuat metadata yang tidak ditemukan.")
        else:
            st.info("Library Referensi masih kosong.")

    elif bagian_ref == "📝 Generasi Sitasi & Daftar Pustaka":
        st.subheader("📝 Generasi Sitasi & Daftar Pustaka")
        refs=st.session_state.bank_referensi
        if not refs:
            st.info("Library Referensi masih kosong. Tambahkan atau validasi sumber terlebih dahulu.")
        else:
            opsi=[f"{i+1}. {r.get('Judul','Tanpa judul')} ({r.get('Tahun','')})" for i,r in enumerate(refs)]
            pilihan=st.multiselect("Pilih referensi; kosong = semua referensi Library", opsi, key="pilih_ref_generasi_sitasi")
            refs_pakai=[refs[opsi.index(x)] for x in pilihan] if pilihan else refs
            gaya_gen=st.selectbox("Gaya sitasi", gaya_list, index=gaya_list.index(st.session_state.gaya_sitasi), key="gaya_generasi_sitasi")
            konteks_gen=st.text_area("Konteks/klaim yang akan diberi sitasi (opsional)", height=180, key="konteks_generasi_sitasi")
            if st.button("🤖 Generate Sitasi & Daftar Pustaka", type="primary", key="btn_generasi_sitasi"):
                daftar="\n".join(
                    f"[{i}] {format_referensi(r,gaya_gen)} | DOI: {r.get('DOI','')} | STATUS: {r.get('Status','')}"
                    for i,r in enumerate(refs_pakai,1)
                )
                prompt=f"""Buat sitasi/footnote dan daftar pustaka dari REFERENSI YANG DISEDIAKAN SAJA.
GAYA: {gaya_gen}
ATURAN:
- Jangan membuat referensi, DOI, penulis, tahun, halaman, kutipan, atau metadata baru.
- Jika gaya menggunakan footnote, buat nomor/penanda konsisten.
- Jika gaya menggunakan in-text citation, gunakan pola gaya tersebut.
- Jika halaman sumber tidak tersedia, tulis [halaman perlu verifikasi].
- Daftar pustaka hanya memuat sumber yang benar-benar dipakai.
KONTEKS/KLAIM:
{konteks_gen or "Tidak ada konteks khusus; buat contoh sitasi bibliografis tanpa mengarang isi sumber."}
REFERENSI:
{daftar}"""
                with st.spinner("Menyusun sitasi dari referensi Library..."):
                    h=panggil_gemini(prompt)
                if h["sukses"]:
                    st.session_state.hasil_generasi_sitasi=h["hasil"]
                else:
                    st.error(h["error"])
            if st.session_state.get("hasil_generasi_sitasi"):
                st.text_area("Hasil — dapat diperiksa dan diedit", st.session_state.hasil_generasi_sitasi, height=500, key="hasil_generasi_sitasi_edit")
    elif bagian_ref == "✍️ Pakai di Naskah":
        st.subheader("✍️ Masukkan Referensi ke BAB / Naskah")
        file_naskah_ref = st.file_uploader(
            "📤 Unggah naskah PDF/DOCX/TXT",
            type=["pdf", "docx", "txt"],
            key="upload_naskah_pakai_referensi"
        )

        teks_upload=""
        sumber_id="manual"
        if file_naskah_ref is not None:
            try:
                teks_upload = ekstrak_teks(file_naskah_ref)
                sumber_id=f"{file_naskah_ref.name}_{getattr(file_naskah_ref,'size',0)}"
                if teks_upload and teks_upload.strip():
                    st.session_state.naskah_upload_ref = teks_upload
                    st.session_state.naskah_aktif = teks_upload
                    st.success(f"Naskah '{file_naskah_ref.name}' berhasil dibaca.")
                else:
                    st.warning("Teks naskah belum dapat dibaca. Anda tetap dapat menempel teks secara manual.")
            except Exception as e:
                st.error(f"Naskah gagal dibaca: {e}")

        isi_awal = teks_upload if teks_upload.strip() else st.session_state.get("naskah_upload_ref","")
        st.markdown("#### 👁️ Naskah / hasil ekstraksi")
        naskah_awal=st.text_area(
            "Naskah / hasil ekstraksi",
            value=isi_awal,
            height=360,
            key=f"naskah_ref_{abs(hash(sumber_id))}",
            label_visibility="collapsed"
        )
        st.session_state.naskah_aktif = naskah_awal

        refs=st.session_state.bank_referensi
        opsi=[f"{i+1}. {r.get('Judul','')} ({r.get('Tahun','')})" for i,r in enumerate(refs)]
        pilihan=st.multiselect("Pilih referensi; kosong = semua yang terverifikasi",opsi,key="pilih_ref_naskah")
        dipilih=[refs[opsi.index(x)] for x in pilihan] if pilihan else [r for r in refs if str(r.get("Status","")).startswith("✅")]
        bagian_target=st.text_input("Bagian/BAB target (opsional)", placeholder="Contoh: BAB I Latar Belakang", key="bagian_target_ref")
        arahan=st.text_area("Arahan",placeholder="Perkuat bagian yang membutuhkan dukungan sumber, tanpa mengubah format dan substansi naskah asli.",key="arah_ref")

        st.info("AI hanya memberi usulan berdasarkan referensi Library. Perubahan tidak diterapkan diam-diam ke naskah asli.")
        if st.button("🤖 Generate AI — Pasang Referensi, Sitasi & Footnote",type="primary",disabled=not bool(naskah_awal.strip()),key="btn_generate_pakai_naskah"):
            arahan_final=(arahan or "Pasang sumber relevan pada klaim yang membutuhkan dukungan.") + (
                f"\nFokus bagian: {bagian_target}." if bagian_target.strip() else ""
            )
            panel_ai_penulisan(
                naskah_awal,
                "Pemasangan sitasi/footnote pada naskah dengan proteksi substansi asli",
                arahan_final + "\nTampilkan usulan SEBELUM -> SESUDAH. Jangan mengubah bagian lain. Jika sumber asli tidak ditemukan, beri saran pencarian atau referensi alternatif berlabel jelas dan jangan mengganti otomatis.",
                dipilih,
                "pasang_ref"
            )
        hasil_pakai_naskah = str(st.session_state.get("hasil_penulisan_ai") or "").strip()
        if hasil_pakai_naskah:
            st.markdown("#### ✨ Hasil AI — periksa sebelum digunakan")

            # Sinkronkan hasil terbaru ke widget khusus menu ini.
            # Ini mencegah nilai lama/kosong pada key Streamlit menutupi hasil AI baru.
            hasil_signature = str(hash(hasil_pakai_naskah))
            if st.session_state.get("_sig_hasil_ref_naskah") != hasil_signature:
                st.session_state["hasil_ref_naskah_editor"] = hasil_pakai_naskah
                st.session_state["_sig_hasil_ref_naskah"] = hasil_signature

            h = st.text_area(
                "Hasil — dapat diedit",
                key="hasil_ref_naskah_editor",
                height=600
            )
            st.session_state.naskah_aktif = h

            if not str(h).strip():
                st.warning("Hasil AI diterima tetapi editor kosong. Klik Generate AI kembali.")
            else:
                st.success("Hasil AI tampil dan siap diperiksa.")

                # Buat SALINAN Word dari DOCX asli. Mesin tidak membangun ulang naskah.
                # Yang ditanam hanya marker kutipan/footnote dari hasil AI.
                if file_naskah_ref is not None and str(file_naskah_ref.name).lower().endswith(".docx"):
                    word_bytes_ref, word_error_ref = buat_word_hasil_revisi(file_naskah_ref, h)
                    if word_bytes_ref:
                        nama_word_ref = Path(file_naskah_ref.name).stem + "_SALINAN_HASIL_REFERENSI.docx"
                        st.download_button(
                            "📥 Unduh Salinan Word Hasil Referensi (.docx)",
                            word_bytes_ref,
                            nama_word_ref,
                            "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            use_container_width=True,
                            key="unduh_salinan_word_pakai_referensi"
                        )
                        st.caption(
                            "Salinan dibuat dari Word asli. Bagian lain tetap berasal dari dokumen asli; "
                            "yang ditambahkan hanya kutipan/footnote yang berhasil dicocokkan."
                        )
                    elif word_error_ref:
                        st.warning(word_error_ref)
                else:
                    st.info("Untuk membuat salinan Word dengan format asli, unggah naskah dalam format DOCX.")
        elif st.session_state.get("hasil_penulisan_ai") is not None:
            st.warning("AI belum mengembalikan isi naskah. Silakan klik Generate AI kembali.")

    elif bagian_ref == "🔧 Perbaiki Footnote & Kutipan":
        st.subheader("🔧 Perbaiki Footnote & Kutipan")
        st.caption("Mendukung footnote, endnote, dan sitasi dalam teks. Mode otomatis mempertahankan gaya bawaan dokumen dan tidak memaksa Chicago.")
        file_perbaikan = st.file_uploader("📤 Unggah naskah PDF/DOCX/TXT",type=["pdf","docx","txt"],key="upload_perbaiki_footnote")
        teks_perbaikan=""
        sumber_foot="manual"
        if file_perbaikan is not None:
            try:
                teks_perbaikan=ekstrak_teks(file_perbaikan)
                sumber_foot=f"{file_perbaikan.name}_{getattr(file_perbaikan,'size',0)}"
                if teks_perbaikan.strip():
                    st.success(f"Naskah '{file_perbaikan.name}' berhasil dibaca.")
                else:
                    st.warning("Teks belum dapat dibaca. Tempel teks secara manual di bawah.")
            except Exception as e:
                st.error(f"Naskah gagal dibaca: {e}")

        st.markdown("#### 👁️ Naskah / hasil ekstraksi")
        teks_perbaikan=st.text_area(
            "Naskah / hasil ekstraksi",
            value=teks_perbaikan,
            height=320,
            key=f"teks_perbaikan_footnote_{abs(hash(sumber_foot))}",
            label_visibility="collapsed"
        )

        mode_gaya=st.radio("Cara menentukan gaya sitasi/footnote",["🔍 Deteksi Otomatis & Pertahankan Gaya Bawaan","✍️ Pilih Gaya Manual"],key="mode_gaya_footnote")
        gaya_target="Gaya bawaan/Custom"
        if mode_gaya.startswith("🔍"):
            if teks_perbaikan.strip():
                hasil_deteksi=deteksi_gaya_sitasi_otomatis(teks_perbaikan); gaya_target=hasil_deteksi["gaya"]
                st.info(f"Gaya terdeteksi: **{gaya_target}** | Keyakinan: **{hasil_deteksi['keyakinan']}** | {hasil_deteksi['alasan']}")
                st.caption("Deteksi otomatis adalah bantuan awal. Jika pola naskah khusus kampus/jurnal, sistem mempertahankan pola tersebut dan menandai bagian yang perlu verifikasi.")
            else:
                st.info("Unggah atau tempel naskah untuk mendeteksi gaya bawaan.")
        else:
            gaya_target=st.selectbox("Pilih gaya",["Chicago Notes & Bibliography","Turabian Notes-Bibliography","OSCOLA","APA 7","Harvard","MLA","IEEE","Vancouver","AMA","ACS","CSE","APSA","Pedoman Kampus/Jurnal","Custom"],key="gaya_manual_footnote")

        refs=st.session_state.bank_referensi
        st.write(f"**Library tersedia:** {len(refs)} referensi")
        hanya_verified=st.checkbox("Utamakan hanya referensi terverifikasi",value=True,key="verified_footnote")
        refs_pakai=[r for r in refs if str(r.get("Status","")).startswith("✅")] if hanya_verified else refs
        if hanya_verified and refs and not refs_pakai:
            st.warning("Belum ada referensi berstatus terverifikasi. Sistem tidak akan menebak sumber pengganti.")

        st.info("Jika sumber footnote tidak ditemukan, AI memberi saran pencarian/kandidat atau referensi alternatif yang diberi label jelas. Tidak ada penggantian otomatis.")
        if st.button("🤖 Generate AI — Audit & Perbaiki Footnote/Kutipan",type="primary",disabled=not bool(teks_perbaikan.strip()),key="btn_perbaiki_footnote"):
            prompt=prompt_perbaiki_footnote_kutipan(teks_perbaikan,gaya_target,refs_pakai,"Pertahankan gaya bawaan" if mode_gaya.startswith("🔍") else "Gaya manual")
            with st.spinner("Mengaudit sitasi, footnote, dan sumber tanpa mengubah substansi naskah..."):
                h=panggil_gemini(prompt)
            if h["sukses"]:
                st.session_state.hasil_perbaikan_footnote=h["hasil"]
                st.success("Audit dan usulan perbaikan selesai. Periksa bagian PERLU VERIFIKASI sebelum digunakan.")
            else:
                st.error(h["error"])
        if st.session_state.get("hasil_perbaikan_footnote"):
            hasil_edit=st.text_area("Hasil — dapat diedit dan diperiksa sebelum dipakai",st.session_state.hasil_perbaikan_footnote,height=650,key="hasil_perbaikan_footnote_edit")
            c_txt, c_word = st.columns(2)
            with c_txt:
                st.download_button(
                    "📥 Unduh Hasil Audit (.txt)",
                    hasil_edit.encode("utf-8"),
                    "hasil_perbaikan_footnote_kutipan.txt",
                    "text/plain",
                    use_container_width=True,
                    key="unduh_audit_footnote_txt"
                )
            with c_word:
                if file_perbaikan is not None and str(file_perbaikan.name).lower().endswith(".docx"):
                    word_bytes, word_error = buat_word_hasil_revisi(file_perbaikan, hasil_edit)
                    if word_bytes:
                        nama_word = Path(file_perbaikan.name).stem + "_HASIL_AKADEMIA_AI.docx"
                        st.download_button(
                            "📥 Unduh Word dengan Footnote Asli (.docx)",
                            word_bytes,
                            nama_word,
                            "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            use_container_width=True,
                            key="unduh_word_footnote"
                        )
                    elif word_error:
                        st.warning(word_error)
                else:
                    st.caption("Unggah naskah DOCX untuk menghasilkan Word dengan format asli dipertahankan.")
            st.warning(
                "File Word dibuat dari salinan naskah asli. Hanya penanda kutipan/footnote yang ditanam sebagai true Word footnote; "
                "bagian lain tetap berasal dari DOCX asli."
            )

    elif bagian_ref == "✅ Audit Sitasi":
        naskah=st.file_uploader("Unggah naskah PDF/DOCX/TXT",type=["pdf","docx","txt"],key="audit_ref_file")
        if naskah:
            teks=ekstrak_teks(naskah); rows=status_sitasi(teks,st.session_state.bank_referensi)
            if rows: st.dataframe(pd.DataFrame(rows),use_container_width=True,hide_index=True)
            else: st.info("Untuk Chicago footnote, gunakan Audit Semantik.")
            if st.button("🤖 Audit Semantik Sitasi dengan Gemini"):
                refs="\n".join(format_referensi(r) for r in st.session_state.bank_referensi)
                h=panggil_gemini("Audit sitasi/footnote. Jangan menyatakan sumber mendukung klaim bila isi sumber tidak tersedia. Jangan membuat DOI/referensi.\nNASKAH:\n"+teks[:60000]+"\nLIBRARY:\n"+refs)
                if h["sukses"]: st.text_area("Hasil Audit",h["hasil"],height=500)
                else: st.error(h["error"])

# ============================================================
# PENELITIAN S1-S3 TERPADU
# ============================================================
elif menu == "🎓 Penelitian S1 • S2 • S3":
    st.header("🎓 Asisten Skripsi S1")
    metode=st.selectbox("Jenis Penelitian",[
        "Belum menentukan metode","Kuantitatif","Kualitatif","Mixed Methods",
        "R&D / Pengembangan","PTK","Studi Literatur / Library Research",
        "Systematic Literature Review (SLR)","Penelitian Evaluatif","Analisis Dokumen / Analisis Isi"
    ], key="metode_s1")
    tahap=st.selectbox("Tahap Skripsi",[
        "Ide & Topik","Identifikasi Masalah","Alternatif Judul","Rumusan Masalah","Tujuan Penelitian",
        "Research Gap","BAB I — Pendahuluan","BAB II — Kajian Teori","Kerangka Berpikir",
        "Hipotesis / Fokus Penelitian","BAB III — Metode","Instrumen Penelitian",
        "Pengumpulan Data","BAB IV — Hasil & Pembahasan","BAB V — Penutup",
        "Skripsi Lengkap","Bimbingan & Revisi","Presentasi","Persiapan Sidang"
    ], key="tahap_s1")
    pedoman=st.file_uploader("📄 Unggah pedoman kampus (opsional)",type=["pdf","docx","txt"],key="pedoman_s1")
    sumber=st.file_uploader("📚 Unggah bahan/referensi/data",type=["pdf","docx","txt","csv","xlsx"],accept_multiple_files=True,key="sumber_s1")
    arah=st.text_area("Ide, masalah, arahan dosen, atau pekerjaan yang ingin dibuat",key="arah_s1")
    konteks=""
    if pedoman:
        t=ekstrak_teks(pedoman)
        if not t.startswith("ERROR:"): konteks+="\nPEDOMAN KAMPUS:\n"+t
    for f in sumber or []:
        if f.name.lower().endswith((".pdf",".docx",".txt")):
            t=ekstrak_teks(f)
            if not t.startswith("ERROR:"): konteks+=f"\nSUMBER {f.name}:\n{t}"
    st.info("BAB IV hanya dibuat dari data nyata. AI tidak boleh menciptakan data penelitian.")
    if st.button("🤖 Generate AI Skripsi S1",type="primary",key="generate_s1"):
        instr=f"""Jenjang: S1 — Skripsi
Metode: {metode}
Tahap: {tahap}
Arahan: {arah}
Gunakan pedoman kampus bila tersedia. Gunakan hanya referensi yang tersedia/terverifikasi. Jangan membuat data, DOI, kutipan, atau nomor halaman palsu."""
        panel_ai_penulisan(konteks,tahap,instr,st.session_state.bank_referensi,"s1")
    if st.session_state.get("hasil_penulisan_ai"):
        edit=st.text_area("Hasil AI — dapat diedit",st.session_state.hasil_penulisan_ai,height=650,key="hasil_s1")
        st.session_state.naskah_aktif=edit


# ============================================================
# PENULIS BUKU AI
# ============================================================
elif menu == "📘 Penulis Buku AI":
    st.header("📘 Penulis Buku AI")
    jenis_buku=st.selectbox("Jenis Buku",["Buku Ajar","Buku Referensi","Monograf","Modul","Buku Akademik"])
    tahap_buku=st.selectbox("Tahap",["Konsep & Pembaca","Outline Buku","Susun BAB","Kembangkan Subbab","Sitasi & Daftar Pustaka","Penyuntingan Buku","Sinopsis & Kata Pengantar","Naskah Buku Lengkap"])
    tema=st.text_area("Tema, tujuan, pembaca sasaran, dan arahan")
    bahan=st.file_uploader("Unggah bahan buku",type=["pdf","docx","txt"],accept_multiple_files=True,key="bahan_buku")
    konteks=""
    for f in bahan or []:
        t=ekstrak_teks(f)
        if not t.startswith("ERROR:"): konteks+=f"\nBAHAN {f.name}:\n"+t
    if st.button("📘 Susun Buku dengan AI",type="primary"):
        panel_ai_penulisan(konteks,f"{jenis_buku} — {tahap_buku}",tema,st.session_state.bank_referensi,"buku")
    if st.session_state.get("hasil_penulisan_ai"):
        st.text_area("Naskah buku — dapat diedit",st.session_state.hasil_penulisan_ai,height=650,key="edit_buku")


# ============================================================
# PENYUNTING AKADEMIK AI
# ============================================================
elif menu == "✨ Penyunting Akademik AI":
    st.header("✨ Penyunting Akademik AI")
    mode_edit=st.selectbox("Mode",[
        "Koreksi Ejaan & Typo","Rapikan Kalimat","Bahasa Akademik",
        "Perkuat Paragraf","Koherensi Antarparagraf",
        "Parafrasa Akademik Bertanggung Jawab","Sunting Naskah Lengkap"
    ])
    file_edit=st.file_uploader("Unggah naskah",type=["pdf","docx","txt"],key="file_editor")
    teks_edit=st.text_area("Atau tempel teks",height=250,key="teks_editor")
    if file_edit:
        t=ekstrak_teks(file_edit)
        if not t.startswith("ERROR:"): teks_edit=t
    if st.button("✨ Sunting dengan AI",type="primary",disabled=not bool(teks_edit.strip())):
        h=panggil_gemini(f"""Sunting teks berikut dengan mode: {mode_edit}.
Pertahankan makna, data, sitasi, nama, dan substansi. Jangan menghapus sitasi untuk menurunkan kemiripan.
Jangan membuat referensi baru. Untuk parafrasa, ubah secara akademik dan wajar, bukan untuk mengelabui pemeriksa plagiarisme.
Tampilkan naskah hasil suntingan dan ringkas perubahan penting.
TEKS:
{teks_edit[:70000]}""")
        if h["sukses"]: st.session_state.hasil_editor=h["hasil"]
        else: st.error(h["error"])
    if st.session_state.get("hasil_editor"):
        st.text_area("Hasil suntingan",st.session_state.hasil_editor,height=650,key="hasil_editor_area")


# ============================================================
# TESIS S2
# ============================================================
elif menu == "🎓 Tesis S2":
    st.header("🎓 Asisten Tesis S2")
    # MENU VERTIKAL TESIS S2
    # Hanya menambah/menata navigasi. Mesin AI dan fitur lama tetap dipertahankan.
    submenu_s2 = st.radio(
        "Menu Tesis S2",
        [
            "💡 Pencarian Ide & Pengajuan Judul 🌟",
            "📑 Proposal Tesis",
            "🔎 Literatur & Penelitian Terdahulu",
            "🎯 Metodologi Penelitian",
            "📋 Instrumen Penelitian",
            "📊 Statistik & SPSS",
            "🧩 Analisis Data Kualitatif",
            "✍️ Penulisan Tesis",
            "🤖 Review & Bimbingan AI",
            "📈 Progres & Timeline Penelitian",
            "🖥️ Presentasi",
            "🎓 Simulasi Sidang",
        ],
        key="submenu_s2_vertikal",
    )

    # Tahap lama tidak dihapus. Semuanya tetap tersedia di bawah submenu yang sesuai.
    tahap_per_submenu_s2 = {
        "💡 Pencarian Ide & Pengajuan Judul 🌟": [
            "Ide & Topik","Identifikasi Masalah","Research Gap","State of the Art",
            "Novelty","Alternatif Judul","Rumusan Masalah","Tujuan Penelitian"
        ],
        "📑 Proposal Tesis": [
            "BAB I","BAB II","Kerangka Berpikir","Hipotesis / Fokus Penelitian","BAB III"
        ],
        "🔎 Literatur & Penelitian Terdahulu": [
            "Research Gap","State of the Art","Novelty","BAB II"
        ],
        "🎯 Metodologi Penelitian": [
            "BAB III","Pengumpulan Data"
        ],
        "📋 Instrumen Penelitian": [
            "Instrumen"
        ],
        "📊 Statistik & SPSS": [
            "BAB IV"
        ],
        "🧩 Analisis Data Kualitatif": [
            "BAB IV"
        ],
        "✍️ Penulisan Tesis": [
            "BAB I","BAB II","Kerangka Berpikir","Hipotesis / Fokus Penelitian",
            "BAB III","BAB IV","BAB V","Tesis Lengkap"
        ],
        "🤖 Review & Bimbingan AI": [
            "Bimbingan & Revisi","Tesis Lengkap"
        ],
        "📈 Progres & Timeline Penelitian": [
            "Tesis Lengkap","Bimbingan & Revisi"
        ],
        "🖥️ Presentasi": [
            "Presentasi"
        ],
        "🎓 Simulasi Sidang": [
            "Persiapan Sidang"
        ],
    }

    # SUBMENU 1 — Pencarian Ide & Pengajuan Judul
    # ============================================================
    # PEDOMAN PENULISAN TESIS - ACUAN UNGGAHAN, BUKAN PERMANEN
    # Berlaku untuk proyek/sesi Tesis S2 dan menjadi konteks AI.
    # ============================================================
    st.markdown("## 📘 Pedoman Penulisan Tesis")
    st.caption("Unggah pedoman kampus sebagai acuan Tesis S2.")

    _pedoman_file_s2 = st.file_uploader(
        "📤 Unggah Pedoman Penulisan Tesis (PDF/DOCX/TXT)",
        type=["pdf", "docx", "txt"],
        key="unggah_pedoman_tesis_s2",
        help="Gunakan pedoman resmi perguruan tinggi/institut/universitas Anda.",
    )

    _c1_ped, _c2_ped = st.columns(2)
    with _c1_ped:
        if st.button("🤖 Analisis & Aktifkan Pedoman", key="analisis_pedoman_tesis_s2", use_container_width=True):
            if _pedoman_file_s2 is None:
                st.warning("Unggah file pedoman terlebih dahulu.")
            else:
                try:
                    _teks_pedoman = ekstrak_teks(_pedoman_file_s2)
                    if not str(_teks_pedoman).strip():
                        st.error("Teks pedoman belum dapat dibaca. Gunakan PDF/DOCX/TXT yang berisi teks.")
                    else:
                        _prompt_pedoman = f"""
Anda adalah analis pedoman akademik. Analisis HANYA dokumen pedoman yang diberikan.
DILARANG mengarang ketentuan yang tidak terdapat dalam dokumen.
Jika suatu ketentuan tidak ditemukan, tulis: "Tidak ditemukan dalam pedoman".

Susun PROFIL PEDOMAN TESIS dengan bagian:
1. Identitas institusi/pedoman jika tercantum.
2. Sistematika proposal tesis.
3. Jenis/metode penelitian yang diatur dan sistematika masing-masing jika ada.
4. Aturan sitasi, footnote/innote, dan daftar pustaka.
5. Ketentuan jumlah/jenis referensi jika ada.
6. Format pengetikan: kertas, margin, font, ukuran, spasi, paragraf, penomoran.
7. Ketentuan transliterasi jika ada.
8. Ketentuan abstrak, tabel, gambar, lampiran, dan bagian awal/akhir jika ada.
9. Ketentuan seminar/ujian/pengesahan jika ada.
10. Checklist kepatuhan yang dapat digunakan AI saat menulis dan mengaudit tesis.

DOKUMEN PEDOMAN:
{str(_teks_pedoman)[:60000]}
"""
                        _hasil_pedoman = panggil_gemini(_prompt_pedoman)
                        if _hasil_pedoman:
                            st.session_state["pedoman_tesis_s2_teks"] = str(_teks_pedoman)
                            st.session_state["pedoman_tesis_s2_analisis"] = str(_hasil_pedoman)
                            st.session_state["pedoman_tesis_s2_nama"] = getattr(_pedoman_file_s2, "name", "Pedoman Tesis")
                            st.session_state["pedoman_tesis_s2_aktif"] = True
                            st.success("Pedoman berhasil dianalisis dan dijadikan acuan aktif.")
                            st.rerun()
                        else:
                            st.error("Analisis pedoman belum menghasilkan keluaran.")
                except Exception as _e_ped:
                    st.error(f"Pedoman belum dapat diproses: {_e_ped}")

    with _c2_ped:
        if st.button("🗑️ Hapus / Ganti Pedoman", key="hapus_pedoman_tesis_s2", use_container_width=True):
            for _kped in [
                "pedoman_tesis_s2_teks",
                "pedoman_tesis_s2_analisis",
                "pedoman_tesis_s2_nama",
                "pedoman_tesis_s2_aktif",
                "unggah_pedoman_tesis_s2",
            ]:
                if _kped in st.session_state:
                    del st.session_state[_kped]
            st.rerun()

    if st.session_state.get("pedoman_tesis_s2_aktif"):
        st.success(f"🟢 Pedoman aktif: {st.session_state.get('pedoman_tesis_s2_nama', 'Pedoman Tesis')}")
    else:
        st.caption("📘 Unggah pedoman kampus jika ingin menjadikannya acuan.")

    # Konteks ringkas pedoman untuk dipakai prompt Tesis S2.
    pedoman_aktif_s2 = ""
    if st.session_state.get("pedoman_tesis_s2_aktif"):
        pedoman_aktif_s2 = (
            "\n\nPEDOMAN INSTITUSI YANG WAJIB DIIKUTI:\n"
            + st.session_state.get("pedoman_tesis_s2_analisis", "")
            + "\nATURAN: Jangan membuat ketentuan yang tidak terdapat dalam pedoman. "
              "Jika pedoman tidak mengatur suatu hal, nyatakan bahwa ketentuan tersebut tidak ditemukan dalam pedoman.\n"
        )

    if submenu_s2 == "💡 Pencarian Ide & Pengajuan Judul 🌟":
        st.markdown("### 🌟 Ruang Kerja Ide, Permasalahan & Judul")
        mode_ide_s2 = st.radio(
            "Cara memulai",
            [
                "🧠 Mulai dari Permasalahan Penelitian",
                "📄 Adaptasi Tesis ke Lokasi/Objek Baru",
                "🔄 Penelitian Lanjutan dari Tesis",
                "🆕 Penelitian Baru dari Tesis Referensi",
                "🔍 Analisis Beberapa Tesis",
            ],
            key="mode_ide_s2",
        )
        st.markdown("### 🔍 Analisis Kelayakan Judul yang Sudah Ada")
        with st.expander("Uji judul milik sendiri/orang lain", expanded=False):
            _judul_uji = st.text_input("Masukkan judul tesis yang ingin diuji", key="judul_uji_kelayakan_s2")
            if st.button("🔎 Analisis Kelayakan Judul", key="analisis_judul_orang_s2"):
                if not _judul_uji.strip():
                    st.warning("Masukkan judul terlebih dahulu.")
                else:
                    with st.spinner("Mencari referensi nyata dan menilai kelayakan judul..."):
                        _data_uji = analisis_ketersediaan_referensi_judul_s2(_judul_uji, 12)
                        _hu = panggil_gemini(f"""Nilai kelayakan judul tesis S2 berikut.
JUDUL: {_judul_uji}
PENCARIAN NYATA: kandidat={_data_uji['Kandidat']}; literatur 5 tahun={_data_uji['Literatur 5 Tahun']}; status={_data_uji['Status']}.
Nilai fokus, masalah ilmiah, researchability, potensi gap, novelty, level S2, metode, data, keluasan judul, dan risiko.
Keputusan: 🟢 LAYAK / 🟡 LAYAK DENGAN REVISI / 🔴 BELUM LAYAK.
Jika perlu beri maksimal 5 perbaikan judul. Jangan mengarang referensi/data/DOI.""", 0.2)
                    if _hu.get("sukses"):
                        st.session_state["hasil_uji_judul_s2"] = _hu.get("hasil","")
                        st.session_state["data_uji_judul_s2"] = _data_uji
                    else:
                        st.error(_hu.get("error","Analisis gagal."))
            if st.session_state.get("data_uji_judul_s2"):
                _du=st.session_state["data_uji_judul_s2"]
                st.info(f"Kandidat referensi: {_du['Kandidat']} | 5 tahun terakhir: {_du['Literatur 5 Tahun']} | {_du['Status']}")
            if st.session_state.get("hasil_uji_judul_s2"):
                st.text_area("Hasil Analisis Kelayakan Judul", st.session_state["hasil_uji_judul_s2"], height=420, key="hasil_uji_judul_s2_area")

        # ------------------------------------------------------------
        # PEMBERSIHAN TOTAL NO.1
        # Diletakkan tepat di bawah pilihan cara memulai dan Analisis Kelayakan Judul.
        # Tidak menghapus Library utama atau fitur/menu lain.
        # ------------------------------------------------------------
        st.markdown("### 🗑️ Pembersihan No. 1")
        if st.session_state.pop("pesan_hapus_total_no1_s2", False):
            st.success("Semua hasil No. 1 dan Bank Bahan sudah dibersihkan. Library utama tetap aman.")
        st.caption(
            "Gunakan tombol ini untuk menghapus seluruh hasil kerja pada Pencarian Ide & Pengajuan Judul "
            "dan memulai dari kondisi bersih. Library utama aplikasi tidak ikut dihapus."
        )
        if st.button(
            "🗑️ Hapus Semua Hasil No. 1 & Mulai Bersih",
            key="hapus_total_no1_s2",
            use_container_width=True,
        ):
            # Simpan hanya state yang memang harus tetap hidup di luar No.1.
            # Library utama sengaja TIDAK dihapus.
            _hapus_pasti = {
                "masalah_ide_s2", "lokasi_ide_s2", "tesis_ide_s2", "arah_ide_s2", "metode_ide_s2",
                "hasil_ai_ide_judul_s2", "hasil_koreksi_ide_s2",
                "judul_alternatif_s2", "judul_tesis_s2_terpilih",
                "dasar_proposal_tesis_s2", "bank_bahan_ide_s2",
                "masalah_terakhir_ide_s2", "versi_naskah_ide_s2",
                "muat_judul_ai_s2", "judul_utama_pilihan_s2",
                "referensi_penguat_ide_s2", "kata_kunci_ref_ide_s2",
                "pilih_ref_penguat_ide_s2", "kelayakan_ref_5_judul_s2",
                "judul_uji_kelayakan_s2", "hasil_uji_judul_s2",
                "data_uji_judul_s2", "hasil_penulisan_ai", "naskah_aktif",
                "hasil_s2", "edit_s2", "hasil_koreksi_s2",
            }
            _potongan_kunci_no1 = (
                "ide_s2", "judul_s2", "judul_alt_s2", "bahan_ide",
                "ref_penguat", "kelayakan_ref", "uji_judul",
                "editor_naskah_ide", "upload_ide", "unggah_ide",
            )

            for _key in list(st.session_state.keys()):
                _ks = str(_key).lower()
                if _key in _hapus_pasti or any(_frag in _ks for _frag in _potongan_kunci_no1):
                    # Jangan sentuh bank referensi / Library utama.
                    if _ks not in ("bank_referensi", "library_referensi"):
                        del st.session_state[_key]

            # Paksa widget 5 judul kosong pada render berikutnya.
            st.session_state["judul_alternatif_s2"] = ["", "", "", "", ""]
            for _i in range(5):
                st.session_state[f"judul_alt_s2_{_i}"] = ""

            st.session_state["bank_bahan_ide_s2"] = []
            st.session_state["hasil_ai_ide_judul_s2"] = ""
            st.session_state["hasil_koreksi_ide_s2"] = ""
            st.session_state["masalah_terakhir_ide_s2"] = ""
            st.session_state["versi_naskah_ide_s2"] = 0
            st.session_state["versi_input_masalah_ide_s2"] = int(
                st.session_state.get("versi_input_masalah_ide_s2", 0)
            ) + 1
            st.session_state["pesan_hapus_total_no1_s2"] = True
            st.rerun()


        st.markdown("### 🆕 Mulai / Ganti Masalah Penelitian")
        st.caption("Gunakan tombol ini jika ingin memulai topik baru agar hasil lama tidak terbawa.")
        if st.session_state.pop("pesan_reset_topik_ide_s2", False):
            st.success("Topik lama, gagasan/masalah, Bank Bahan, hasil AI, dan alternatif judul sudah dibersihkan. Silakan mulai topik baru.")
        if st.button("🧹 Bersihkan Hasil Ide Lama & Mulai Topik Baru", key="reset_topik_ide_s2"):
            # Reset TOPIK harus membersihkan input gagasan juga, bukan hanya hasil AI.
            # Library utama tetap dipertahankan.
            _hapus_reset_topik = {
                "masalah_ide_s2", "lokasi_ide_s2", "tesis_ide_s2", "arah_ide_s2", "metode_ide_s2",
                "hasil_ai_ide_judul_s2", "hasil_koreksi_ide_s2",
                "judul_alternatif_s2", "judul_tesis_s2_terpilih",
                "dasar_proposal_tesis_s2", "bank_bahan_ide_s2",
                "masalah_terakhir_ide_s2", "versi_naskah_ide_s2",
                "muat_judul_ai_s2", "judul_utama_pilihan_s2",
                "referensi_penguat_ide_s2", "kata_kunci_ref_ide_s2",
                "pilih_ref_penguat_ide_s2", "kelayakan_ref_5_judul_s2",
                "hasil_penulisan_ai", "naskah_aktif", "hasil_s2",
            }
            _frag_reset_topik = (
                "judul_alt_s2_", "editor_naskah_ide_s2_", "bahan_ide",
                "upload_ide", "unggah_ide", "file_bahan_ide",
            )
            for _key in list(st.session_state.keys()):
                _ks = str(_key)
                if _key in _hapus_reset_topik or any(_frag in _ks for _frag in _frag_reset_topik):
                    if mode_bank_bahan_ide_s2 != "🤖 Otomatis dengan AI" and (_ks not in ("bank_referensi", "library_referensi")):
                        del st.session_state[_key]
            st.session_state["versi_input_masalah_ide_s2"] = int(
                st.session_state.get("versi_input_masalah_ide_s2", 0)
            ) + 1
            st.session_state["pesan_reset_topik_ide_s2"] = True
            st.rerun()


        if "versi_input_masalah_ide_s2" not in st.session_state:
            st.session_state["versi_input_masalah_ide_s2"] = 0
        _key_masalah_ide_s2 = f"masalah_ide_s2_{st.session_state['versi_input_masalah_ide_s2']}"
        masalah_ide_s2 = st.text_area(
            "Permasalahan/gagasan awal",
            placeholder="Tuliskan masalah nyata yang ingin diteliti...",
            height=160,
            key=_key_masalah_ide_s2,
        )

        st.markdown("### 🎯 Penentuan Metode Penelitian")
        st.caption(
            "Pilih metode sebelum Generate AI. Jika belum yakin, pilih Rekomendasi AI. "
            "AI akan menyarankan metode yang paling sesuai berdasarkan masalah dan bahan penelitian."
        )
        metode_ide_s2 = st.selectbox(
            "Pilih metode penelitian",
            [
                "🤖 Rekomendasi AI",
                "🔵 Kualitatif",
                "🟢 Kuantitatif",
                "🟣 Mixed Methods",
                "🟠 R&D / Research and Development",
                "🔴 Penelitian Tindakan / Action Research",
                "📚 Penelitian Literatur",
            ],
            key="metode_ide_s2",
        )
        lokasi_ide_s2 = ""
        if mode_ide_s2 == "📄 Adaptasi Tesis ke Lokasi/Objek Baru":
            lokasi_ide_s2 = st.text_input(
                "Lokasi/objek penelitian baru",
                key="lokasi_ide_s2",
            )
        tesis_ide_s2 = st.file_uploader(
            "Unggah tesis selesai sebagai bahan (PDF/DOCX)",
            type=["pdf","docx"],
            accept_multiple_files=True,
            key="tesis_ide_s2",
        )
        arah_ide_s2 = st.text_area(
            "Arah penelitian yang diinginkan (opsional)",
            placeholder="Contoh: pertahankan variabel, ubah lokasi, tambah variabel, atau kembangkan penelitian.",
            key="arah_ide_s2",
        )


        # ============================================================
        # BANK BAHAN PENGUAT PERMASALAHAN - dapat ditambah berkali-kali
        # ============================================================
        st.markdown("### 🧠 Cara Menyiapkan Bahan Penguat")
        mode_bank_bahan_ide_s2 = st.radio(
            "Pilih cara menyiapkan bahan penguat permasalahan:",
            [
                "🤖 Otomatis dengan AI",
                "📂 Gunakan Dokumen Saya",
                "🔄 Gabungkan AI + Dokumen Saya",
            ],
            key="mode_bank_bahan_ide_s2",
            help=(
                "Otomatis dengan AI: tidak wajib unggah dokumen. "
                "Dokumen Saya: gunakan bahan yang Anda miliki. "
                "Gabungkan: AI melengkapi dokumen Anda."
            ),
        )

        if mode_bank_bahan_ide_s2 == "🤖 Otomatis dengan AI":
            st.caption("✅ Unggahan 1-4 bersifat opsional pada mode Otomatis dengan AI.")
            st.info(
                "Anda tidak wajib mengunggah dokumen. AI akan membantu memetakan kebutuhan bahan, "
                "kata kunci, teori, penelitian terdahulu, regulasi, dan data yang perlu diperkuat. "
                "Referensi akademik tetap dicari dari sumber nyata dan tidak boleh dibuat-buat."
            )
        elif mode_bank_bahan_ide_s2 == "📂 Gunakan Dokumen Saya":
            st.info(
                "Unggah bahan yang Anda miliki pada Bank Bahan di bawah. "
                "AI akan memprioritaskan isi dokumen tersebut sebagai penguat permasalahan."
            )
        else:
            st.info(
                "AI akan menggunakan dokumen Anda terlebih dahulu, lalu membantu mencari dan "
                "memetakan bahan tambahan yang masih kurang."
            )

        st.markdown("### 📂 Bank Bahan Penguat Permasalahan")
        st.caption(
            "Tambahkan bahan secara bertahap. Bahan yang sudah dimasukkan tetap tersimpan selama sesi "
            "dan digunakan AI untuk menguatkan masalah, gap, novelty, serta alternatif judul."
        )

        if "bank_bahan_ide_s2" not in st.session_state:
            st.session_state["bank_bahan_ide_s2"] = []

        jenis_bahan_ide = st.selectbox(
            "Jenis bahan yang akan ditambahkan",
            [
                "📚 Teori & Literatur",
                "🔎 Penelitian Terdahulu / Tesis / Disertasi",
                "📊 Data, Persentase & Fakta Lapangan",
                "📑 Regulasi & Dokumen Resmi",
                "🖥️ Presentasi / Materi Seminar / Bimtek",
                "📝 Observasi / Wawancara / Catatan Pengawas",
                "🎥 Transkrip Video / Audio",
                "📎 Dokumen Pendukung Lainnya",
            ],
            key="jenis_bahan_ide_s2",
        )
        ket_bahan_ide = st.text_input(
            "Keterangan bahan (opsional)",
            placeholder="Contoh: hasil supervisi MTs 2026, teori kokurikuler, transkrip bimtek, data persentase...",
            key="ket_bahan_ide_s2",
        )
        st.markdown("#### 📥 Unggahan Pendukung (boleh banyak dan bertahap)")
        st.caption("Gunakan beberapa slot di bawah. Setiap slot dapat memuat banyak file sekaligus.")
        _uploads_ide = []
        for _slot in range(1, 5):
            _u = st.file_uploader(
                f"Unggahan {_slot}",
                type=["pdf", "docx", "txt", "xlsx", "xls", "csv", "pptx"],
                accept_multiple_files=True,
                key=f"unggah_bahan_ide_s2_slot_{_slot}",
            )
            if _u:
                _uploads_ide.extend(_u)

        if mode_bank_bahan_ide_s2 in ["📂 Gunakan Dokumen Saya", "🔄 Gabungkan AI + Dokumen Saya"]:
            if st.button("➕ Masukkan Semua Unggahan ke Bank Bahan Ide", key="tambah_bank_bahan_ide_s2", type="primary"):
                if not _uploads_ide:
                    st.warning("Pilih minimal satu file pada Unggahan 1-4.")
                else:
                    _tambah = 0
                    for _fb in _uploads_ide:
                        _nama = getattr(_fb, "name", "Bahan")
                        _identitas = f"{jenis_bahan_ide}|{_nama}|{getattr(_fb, 'size', 0)}"
                        if any(x.get("id") == _identitas for x in st.session_state["bank_bahan_ide_s2"]):
                            continue

                        _teks_bahan = ""
                        try:
                            # Gunakan mesin ekstraksi Akademia AI yang sudah ada.
                            _teks_bahan = ekstrak_teks(_fb)
                        except Exception as _e:
                            _teks_bahan = f"ERROR: {_e}"

                        if not str(_teks_bahan).startswith("ERROR:") and str(_teks_bahan).strip():
                            st.session_state["bank_bahan_ide_s2"].append({
                                "id": _identitas,
                                "nama": _nama,
                                "jenis": jenis_bahan_ide,
                                "keterangan": ket_bahan_ide.strip(),
                                "teks": str(_teks_bahan),
                                "status": "✅ Siap dianalisis AI",
                            })
                            _tambah += 1
                        else:
                            st.warning(
                                f"{_nama}: teks belum dapat diekstrak otomatis. "
                                "Jika ini bahan video/audio, gunakan transkrip dari modul Analisis Karya Akademik lalu unggah transkripnya."
                            )
                    if _tambah:
                        st.success(f"✅ {_tambah} bahan ditambahkan ke Bank Bahan Ide.")

        # Bahan teks/manual sangat penting untuk data persentase, hasil observasi, atau transkrip pendek.
        with st.expander("✍️ Tambahkan bahan dalam bentuk teks / transkrip manual"):
            jenis_teks_ide = st.selectbox(
                "Jenis bahan teks",
                [
                    "📊 Data, Persentase & Fakta Lapangan",
                    "📝 Observasi / Wawancara / Catatan Pengawas",
                    "🎥 Transkrip Video / Audio",
                    "📚 Teori & Literatur",
                    "📑 Regulasi & Dokumen Resmi",
                    "📎 Catatan Lainnya",
                ],
                key="jenis_teks_ide_s2",
            )
            judul_teks_ide = st.text_input(
                "Nama/Judul bahan",
                placeholder="Contoh: Transkrip presentasi kokurikuler 2026",
                key="judul_teks_ide_s2",
            )
            isi_teks_ide = st.text_area(
                "Tempel isi bahan / transkrip / data",
                height=180,
                key="isi_teks_ide_s2",
            )
            if st.button("➕ Simpan Teks ke Bank Bahan", key="simpan_teks_bank_ide_s2"):
                if not isi_teks_ide.strip():
                    st.warning("Isi bahan teks masih kosong.")
                else:
                    _nama_manual = judul_teks_ide.strip() or "Bahan teks manual"
                    _id_manual = f"manual|{jenis_teks_ide}|{_nama_manual}|{len(isi_teks_ide)}"
                    st.session_state["bank_bahan_ide_s2"].append({
                        "id": _id_manual,
                        "nama": _nama_manual,
                        "jenis": jenis_teks_ide,
                        "keterangan": "Input teks/transkrip manual",
                        "teks": isi_teks_ide.strip(),
                        "status": "✅ Siap dianalisis AI",
                    })
                    st.success("✅ Bahan teks/transkrip ditambahkan.")

        if st.session_state["bank_bahan_ide_s2"]:
            st.markdown("#### 📚 Bahan yang Sudah Terkumpul")
            for _idx, _b in enumerate(st.session_state["bank_bahan_ide_s2"]):
                _c1, _c2 = st.columns([8, 1])
                with _c1:
                    st.markdown(
                        f"**{_idx + 1}. {_b['nama']}**  \n"
                        f"{_b['jenis']} · {_b.get('status', '✅ Siap')}  \n"
                        f"{_b.get('keterangan', '')}"
                    )
                with _c2:
                    if st.button("🗑️", key=f"hapus_bahan_ide_s2_{_idx}", help="Hapus bahan ini"):
                        st.session_state["bank_bahan_ide_s2"].pop(_idx)
                        st.rerun()

            if st.button("🗑️ Kosongkan Bank Bahan Ide", key="kosongkan_bank_bahan_ide_s2"):
                st.session_state["bank_bahan_ide_s2"] = []
                st.rerun()

        # No.1 memakai state KHUSUS. Tidak memakai naskah_aktif/hasil_penulisan_ai
        # agar hasil Proposal atau modul lain tidak dapat muncul di ruang Ide/Judul.
        if "hasil_ai_ide_judul_s2" not in st.session_state:
            st.session_state["hasil_ai_ide_judul_s2"] = ""
        if "versi_naskah_ide_s2" not in st.session_state:
            st.session_state["versi_naskah_ide_s2"] = 0

        st.markdown("### 🔎 Referensi Penguat Otomatis")
        st.caption("AI membuat kata kunci Indonesia/Inggris. Referensi kemudian dicari dari sumber akademik nyata yang sudah terhubung, bukan dibuat oleh AI.")
        if st.button("🔎 Cari Referensi Penguat Otomatis", key="cari_ref_penguat_ide_s2", use_container_width=True):
            if not masalah_ide_s2.strip():
                st.warning("Isi permasalahan/gagasan awal terlebih dahulu.")
            else:
                with st.spinner("Menyusun kata kunci dan mencari referensi akademik..."):
                    _kw, _refs = cari_referensi_penguat_ide_s2(masalah_ide_s2, arah_ide_s2)
                st.session_state["kata_kunci_ref_ide_s2"] = _kw
                st.session_state["referensi_penguat_ide_s2"] = _refs
        if st.session_state.get("kata_kunci_ref_ide_s2"):
            with st.expander("🔑 Kata kunci pencarian yang digunakan", expanded=False):
                for _q in st.session_state["kata_kunci_ref_ide_s2"]:
                    st.write("• "+_q)
        _refs_penguat = st.session_state.get("referensi_penguat_ide_s2", [])
        if _refs_penguat:
            st.success(f"{len(_refs_penguat)} kandidat referensi ditemukan dari pencarian nyata.")
            _ops_ref = list(range(len(_refs_penguat)))
            _pilih_ref = st.multiselect(
                "Pilih referensi yang akan dimasukkan ke Library", _ops_ref,
                format_func=lambda i: f"{_refs_penguat[i].get('Tahun','')} | {_refs_penguat[i].get('Judul','')} | {_refs_penguat[i].get('Sumber','')}",
                key="pilih_ref_penguat_ide_s2")
            if st.button("➕ Masukkan Referensi Terpilih ke Library", key="masuk_library_ref_ide_s2"):
                _baru=0
                for _i in _pilih_ref:
                    if tambah_bank_referensi(_refs_penguat[_i]): _baru+=1
                st.success(f"{_baru} referensi baru masuk ke Library. Duplikat dilewati.")

        if st.button("🤖 Analisis Ide, Gap, Novelty & 5 Judul", key="gen_ide_s2", type="primary"):
            if not masalah_ide_s2.strip():
                st.warning("Tuliskan permasalahan/gagasan penelitian terlebih dahulu.")
            else:
                bahan_ide_s2 = ""
                for _f in tesis_ide_s2 or []:
                    _t = ekstrak_teks(_f)
                    if not _t.startswith("ERROR:"):
                        bahan_ide_s2 += f"\n\nTESIS SUMBER {_f.name}:\n{_t[:30000]}"

                # Gabungkan seluruh bahan penguat yang telah dikumpulkan.
                _bank_bahan_prompt = ""
                for _no_b, _b in enumerate(st.session_state.get("bank_bahan_ide_s2", []), 1):
                    _bank_bahan_prompt += (
                        f"\n\n--- BAHAN PENGUAT {_no_b} ---\n"
                        f"Jenis: {_b.get('jenis', '')}\n"
                        f"Nama: {_b.get('nama', '')}\n"
                        f"Keterangan: {_b.get('keterangan', '')}\n"
                        f"Isi:\n{_b.get('teks', '')[:30000]}"
                    )

                _prompt_ide_s2 = f"""Anda adalah dosen pembimbing akademik tesis Magister (S2).

INI HANYA TAHAP NOMOR 1: PENCARIAN IDE DAN PENGAJUAN JUDUL.
DILARANG MENULIS PROPOSAL TESIS.
DILARANG MENULIS BAB I, BAB II, BAB III.
DILARANG membuat bagian Latar Belakang, Rumusan Masalah, Tujuan Penelitian,
Kerangka Teoretis, Metodologi lengkap, Catatan Kaki, atau Daftar Pustaka.
DILARANG menetapkan satu judul sebagai judul final.
PENGGUNA SENDIRI yang memilih dan menetapkan judul.

MODE:
{mode_ide_s2}

MASALAH/GAGASAN BARU PENGGUNA:
{masalah_ide_s2}

LOKASI/OBJEK BARU:
{lokasi_ide_s2}

ARAH PENELITIAN YANG DIINGINKAN:
{arah_ide_s2}

PILIHAN METODE PENELITIAN:
{metode_ide_s2}

MODE BANK BAHAN PENGUAT:
{mode_bank_bahan_ide_s2}

ATURAN BANK BAHAN:
- Jika mode "🤖 Otomatis dengan AI", pengguna tidak wajib memiliki dokumen. Petakan bahan penguat yang diperlukan berdasarkan permasalahan, lalu gunakan pencarian referensi nyata yang tersedia di aplikasi. Jangan mengarang judul artikel, penulis, DOI, regulasi, angka, persentase, atau fakta lapangan.
- Jika mode "📂 Gunakan Dokumen Saya", prioritaskan bukti dari Bank Bahan pengguna. Jika bukti belum cukup, jelaskan kekurangannya tanpa membuat data baru.
- Jika mode "🔄 Gabungkan AI + Dokumen Saya", gunakan dokumen pengguna sebagai bukti utama dan lengkapi kekurangannya melalui pemetaan AI serta pencarian sumber nyata.
- Bedakan fakta dari dokumen pengguna, hasil pencarian referensi nyata, dan saran AI tentang bahan yang masih perlu dicari.

PEDOMAN INSTITUSI AKTIF:
{pedoman_aktif_s2 if pedoman_aktif_s2 else "Tidak ada pedoman institusi yang diunggah/diaktifkan."}

ATURAN PENENTUAN METODE:
- Jika pilihan pengguna adalah "🤖 Rekomendasi AI", tentukan SATU metode yang paling sesuai setelah membaca masalah dan seluruh bahan. Jelaskan alasan singkat pada bagian analisis, lalu buat 5 alternatif judul yang konsisten dengan metode rekomendasi tersebut.
- Jika pengguna memilih metode tertentu, JANGAN menggantinya dengan metode lain. Analisis masalah, research gap, novelty, dan 5 alternatif judul harus konsisten dengan metode pilihan pengguna.
- Jangan memaksakan variabel kuantitatif jika metode yang dipilih bukan Kuantitatif.
- Jangan otomatis mengubah Penelitian Tindakan menjadi eksperimen hanya karena tujuan menggunakan kata "meningkatkan".
- Untuk R&D, judul harus benar-benar mencerminkan pengembangan/validasi produk, model, media, modul, atau aplikasi yang relevan.
- Untuk Kualitatif, utamakan fenomena, proses, pengalaman, implementasi, strategi, atau makna sesuai masalah.
- Untuk Mixed Methods, judul harus layak menggunakan kombinasi data kuantitatif dan kualitatif.
- Untuk Penelitian Literatur, jangan membuat seolah-olah ada intervensi lapangan.
- Rekomendasi metode AI bersifat saran. Keputusan akhir tetap milik pengguna.

BAHAN TESIS SUMBER (jika ada):
{bahan_ide_s2 if bahan_ide_s2 else "Tidak ada tesis sumber yang diunggah."}

BANK BAHAN PENGUAT PERMASALAHAN:
{_bank_bahan_prompt if _bank_bahan_prompt else "Belum ada bahan penguat tambahan."}

ATURAN MEMBACA BAHAN:
- Bedakan fakta lapangan, data/persentase, teori, regulasi, penelitian terdahulu, presentasi, dan transkrip.
- Gunakan bahan untuk MENGUATKAN atau MENGUJI masalah pengguna, bukan sekadar merangkum file.
- Jangan menyatakan research gap sudah terbukti jika penelitian terdahulu yang tersedia belum cukup.
- Jika bukti belum cukup, tulis jelas: "Research gap masih perlu diperkuat dengan literatur."
- Jangan mengubah data/persentase dari bahan.
- Jangan menganggap isi presentasi/transkrip sebagai teori ilmiah kecuali sumber ilmiahnya memang tersedia.

PAGAR KONTEKS:
- Gunakan HANYA masalah/gagasan baru di atas dan bahan yang diunggah pada proses ini.
- Abaikan seluruh hasil pekerjaan AI sebelumnya.
- Jangan membawa topik dari sesi/modul lain.
- Jangan mengarang data, temuan, DOI, kutipan, nomor halaman, atau referensi.
- Research gap dan novelty pada tahap ini adalah HIPOTESIS AWAL yang masih perlu dibuktikan melalui literatur.
- Ikuti PILIHAN METODE PENELITIAN di atas.
- Jika pengguna memilih Rekomendasi AI, rekomendasikan satu metode yang paling cocok tetapi jangan menetapkannya sebagai keputusan final pengguna.

OUTPUT WAJIB HANYA:
A. ANALISIS MASALAH
- Nyatakan masalah inti.
- Bedakan gejala, akar masalah, dampak, dan kebutuhan penelitian.
- Jangan menambah fakta yang tidak ada pada input/bahan.

B. PETA BUKTI PENGUAT
Buat tabel ringkas: Bahan/Sumber | Fakta/Teori yang Didukung | Hubungannya dengan Masalah | Kekuatan Bukti | Yang Masih Kurang.
Jika Bank Bahan kosong, katakan bukti pendukung belum tersedia.

C. RESEARCH GAP AWAL
- Bedakan gap empiris, teoretis, metodologis, dan kontekstual bila memang didukung bahan.
- Jangan menyatakan "sebagian besar penelitian" tanpa bukti dari bahan.
- Jika belum cukup, tulis: "Research gap masih perlu diperkuat dengan literatur."

D. POTENSI NOVELTY
Berikan 2-4 kemungkinan novelty yang logis dan tandai sebagai POTENSI, bukan klaim final.

E. 5 ALTERNATIF JUDUL TESIS
Tulis tepat dengan format:
[JUDUL 1] ...
[JUDUL 2] ...
[JUDUL 3] ...
[JUDUL 4] ...
[JUDUL 5] ...

F. CATATAN PEMILIHAN
Untuk setiap judul jelaskan: fokus, kelebihan, risiko/kebutuhan data, dan metode yang mungkin.
Beri rekomendasi maksimal 2 judul terkuat beserta alasan, tetapi keputusan tetap milik pengguna.

BERHENTI setelah bagian F. JANGAN LANJUT KE PROPOSAL."""

                with st.spinner("AI menganalisis masalah dan menyiapkan 5 alternatif judul..."):
                    _h_ide = panggil_gemini(_prompt_ide_s2)

                if _h_ide.get("sukses") and str(_h_ide.get("hasil", "")).strip():
                    _hasil_baru = str(_h_ide["hasil"]).strip()

                    # Pagar kedua: jika model tetap menghasilkan proposal, jangan tampilkan hasil itu.
                    _indikator_proposal = [
                        "PROPOSAL AWAL TESIS",
                        "BAB I",
                        "BAB II",
                        "BAB III",
                        "DAFTAR PUSTAKA",
                        "CATATAN KAKI",
                    ]
                    _terdeteksi_proposal = sum(
                        1 for _x in _indikator_proposal if _x in _hasil_baru.upper()
                    ) >= 2

                    if _terdeteksi_proposal:
                        st.session_state["hasil_ai_ide_judul_s2"] = ""
                        st.error(
                            "Hasil AI terdeteksi masuk ke penyusunan Proposal, sehingga tidak ditampilkan. "
                            "Silakan klik Generate sekali lagi. Tahap No.1 hanya boleh menghasilkan analisis dan 5 alternatif judul."
                        )
                    else:
                        st.session_state["hasil_ai_ide_judul_s2"] = _hasil_baru
                        st.session_state["masalah_terakhir_ide_s2"] = masalah_ide_s2.strip()
                        st.session_state["versi_naskah_ide_s2"] += 1

                        _judul_ai = re.findall(
                            r"(?im)^\s*\[JUDUL\s*[1-5]\]\s*[:\-]?\s*(.+?)\s*$",
                            _hasil_baru,
                        )
                        _judul_ai = [
                            re.sub(r'^[\"\'“”]+|[\"\'“”]+$', "", j.strip())
                            for j in _judul_ai
                            if j.strip()
                        ]

                        if _judul_ai:
                            _judul_ai = (_judul_ai + ["", "", "", "", ""])[:5]
                            st.session_state["judul_alternatif_s2"] = _judul_ai
                            # Bank judul membaca state ini sebagai nilai awal pada rerun.
                            st.session_state["muat_judul_ai_s2"] = True
                            st.success(
                                f"✅ Analisis selesai. {sum(bool(j) for j in _judul_ai)} alternatif judul ditemukan. "
                                "Anda yang memilih judul."
                            )
                        else:
                            st.warning(
                                "Analisis selesai, tetapi 5 judul belum terbaca otomatis. "
                                "Hasil tetap ditampilkan untuk Anda edit."
                            )
                else:
                    st.error(_h_ide.get("error", "AI belum mengembalikan hasil."))

        # Tombol hapus hasil AI No.1, tidak menghapus Bank Bahan.
        if st.session_state.get("hasil_ai_ide_judul_s2"):
            if st.button("🗑️ Hapus Hasil Analisis AI", key="hapus_hasil_ai_ide_s2"):
                for _k in [
                    "hasil_ai_ide_judul_s2",
                    "hasil_koreksi_ide_s2",
                    "judul_alternatif_s2",
                    "masalah_terakhir_ide_s2",
                ]:
                    if _k in st.session_state:
                        del st.session_state[_k]
                st.session_state["judul_alternatif_s2"] = ["", "", "", "", ""]
                st.success("Hasil analisis AI dihapus. Bank Bahan tetap tersimpan.")
                st.rerun()

        _naskah_ide = str(st.session_state.get("hasil_ai_ide_judul_s2", "") or "").strip()
        _masalah_hasil = str(st.session_state.get("masalah_terakhir_ide_s2", "") or "").strip()
        # Hasil lama tidak boleh muncul untuk masalah yang berbeda.
        if _naskah_ide and _masalah_hasil and _masalah_hasil != masalah_ide_s2.strip():
            st.info("Masalah penelitian telah berubah. Klik Generate untuk membuat analisis baru; hasil lama tidak ditampilkan.")
            _naskah_ide = ""
        if _naskah_ide:
            st.divider()
            st.subheader("📊 Hasil Analisis Ide, Gap, Novelty & 5 Alternatif Judul")
            _v_ide = st.session_state.get("versi_naskah_ide_s2", 0)
            _edit_ide = st.text_area(
                "Hasil AI dapat diedit langsung di sini",
                value=_naskah_ide,
                height=650,
                key=f"editor_naskah_ide_s2_{_v_ide}",
            )
            st.session_state["hasil_ai_ide_judul_s2"] = _edit_ide

            if st.button("🔍 Koreksi Ulang Hasil Edit", key="koreksi_ide_s2"):
                _prompt_koreksi_ide = f"""Anda adalah dosen pembimbing tesis S2.
Koreksi HANYA tahap Ide & Pengajuan Judul berikut.
Jangan membuat proposal, BAB I-III, daftar pustaka, atau catatan kaki.
Pertahankan topik pengguna.
Periksa masalah, gap awal, potensi novelty, 5 alternatif judul, dan kelayakan arah metode.
Pengguna sendiri yang memilih judul final.
Jangan membuat data atau referensi palsu.

NASKAH TERBARU:
{_edit_ide}
"""
                _k = panggil_gemini(_prompt_koreksi_ide)
                if _k.get("sukses") and str(_k.get("hasil", "")).strip():
                    st.session_state["hasil_koreksi_ide_s2"] = _k["hasil"]
                else:
                    st.error(_k.get("error", "Koreksi AI gagal."))

            if st.session_state.get("hasil_koreksi_ide_s2"):
                st.markdown("#### 🔍 Hasil Koreksi AI")
                st.text_area(
                    "Saran dan versi perbaikan",
                    value=st.session_state["hasil_koreksi_ide_s2"],
                    height=450,
                    key="hasil_koreksi_ide_s2_area",
                )


        # ------------------------------------------------------------
        # BANK 5 ALTERNATIF JUDUL — dapat diedit sebelum masuk proposal
        # ------------------------------------------------------------
        st.markdown("#### 🏷️ Bank Alternatif Judul")
        st.caption(
            "Setelah analisis masalah/ide, susun hingga 5 judul. "
            "Semua judul dapat diedit manual sebelum satu judul ditetapkan."
        )

        if "judul_alternatif_s2" not in st.session_state:
            st.session_state["judul_alternatif_s2"] = ["", "", "", "", ""]

        # Muat 5 judul hasil Generate ke widget hanya setelah Generate baru berhasil.
        if st.session_state.pop("muat_judul_ai_s2", False):
            for _i, _j in enumerate(st.session_state["judul_alternatif_s2"][:5]):
                st.session_state[f"judul_alt_s2_{_i}"] = _j

        # Pengguna bebas mengedit 5 judul sebelum memilih satu.
        judul_edit_s2 = []
        for _i in range(5):
            _key_judul = f"judul_alt_s2_{_i}"
            if _key_judul not in st.session_state:
                st.session_state[_key_judul] = st.session_state["judul_alternatif_s2"][_i]
            _j = st.text_input(
                f"Alternatif Judul {_i+1}",
                key=_key_judul,
                placeholder=f"Tulis/edit alternatif judul {_i+1}",
            )
            judul_edit_s2.append(_j)
        st.session_state["judul_alternatif_s2"] = judul_edit_s2

        judul_tersedia_s2 = [j.strip() for j in judul_edit_s2 if j.strip()]
        if any(j.strip() for j in judul_edit_s2):
            if st.button("📚 Cek Ketersediaan Referensi untuk 5 Judul", key="cek_ref_5_judul_s2", use_container_width=True):
                _kel=[]
                with st.spinner("Mengecek ketersediaan referensi nyata untuk setiap judul..."):
                    for _j in judul_edit_s2:
                        if _j.strip():
                            _kel.append(analisis_ketersediaan_referensi_judul_s2(_j.strip(),10))
                st.session_state["kelayakan_ref_5_judul_s2"]=_kel
        if st.session_state.get("kelayakan_ref_5_judul_s2"):
            st.markdown("#### 📚 Analisis Kelayakan Referensi")
            st.caption("Jumlah berasal dari hasil pencarian metadata nyata, bukan perkiraan AI.")
            for _no,_d in enumerate(st.session_state["kelayakan_ref_5_judul_s2"],1):
                with st.expander(f"Alternatif {_no} — {_d['Status']}", expanded=False):
                    st.write(f"**Judul:** {_d['Judul']}")
                    st.write(f"**Kandidat referensi:** {_d['Kandidat']}")
                    st.write(f"**Literatur 5 tahun terakhir:** {_d['Literatur 5 Tahun']}")
                    st.write(f"**Status:** {_d['Status']}")
                    for _r in _d.get("Referensi",[])[:5]:
                        st.write(f"• {_r.get('Tahun','')} — {_r.get('Judul','')} [{_r.get('Sumber','')}]")

        if judul_tersedia_s2:
            judul_pilihan_s2 = st.selectbox(
                "⭐ Anda Pilih Judul yang Akan Ditetapkan",
                judul_tersedia_s2,
                key="judul_utama_pilihan_s2",
            )

            cjudul1, cjudul2 = st.columns(2)
            with cjudul1:
                if st.button(
                    "🔍 Koreksi Judul Terpilih dengan AI",
                    key="koreksi_judul_terpilih_s2",
                    use_container_width=True,
                ):
                    instr_koreksi_judul_s2 = f"""Periksa kelayakan judul tesis S2 berikut:
{judul_pilihan_s2}

Konteks masalah:
{masalah_ide_s2}

Arah penelitian:
{arah_ide_s2}

Periksa:
1. kesesuaian judul dengan masalah;
2. ketajaman fokus;
3. research gap dan potensi novelty;
4. variabel/fokus, subjek/objek, dan lokasi bila relevan;
5. kelayakan metodologis;
6. kelayakan untuk tesis S2;
7. risiko terlalu mirip dengan penelitian sumber.

Berikan:
A. penilaian singkat;
B. bagian yang perlu diperbaiki;
C. maksimal 3 versi judul perbaikan.
Jangan membuat data atau referensi palsu."""
                    panel_ai_penulisan(
                        judul_pilihan_s2,
                        "Koreksi Judul Terpilih Tesis S2",
                        instr_koreksi_judul_s2,
                        st.session_state.bank_referensi,
                        "s2_koreksi_judul",
                    )

            with cjudul2:
                if st.button(
                    "✅ Tetapkan Judul & Lanjutkan ke Proposal",
                    key="tetapkan_judul_s2",
                    use_container_width=True,
                ):
                    st.session_state["judul_tesis_s2_terpilih"] = judul_pilihan_s2

                    # DATA PROYEK TESIS S2: sumber utama antar-submenu.
                    st.session_state["proyek_tesis_s2"] = {
                        "judul": judul_pilihan_s2,
                        "masalah": masalah_ide_s2,
                        "arah": arah_ide_s2,
                        "mode": mode_ide_s2,
                        "metode": metode_ide_s2,
                    }

                    # Tetap simpan format lama agar fitur yang sudah berjalan tidak berubah.
                    st.session_state["dasar_proposal_tesis_s2"] = {
                        "judul": judul_pilihan_s2,
                        "masalah": masalah_ide_s2,
                        "arah": arah_ide_s2,
                        "mode": mode_ide_s2,
                    }

                    # Isi langsung state widget Proposal sehingga tidak bergantung pada value=.
                    st.session_state["judul_proposal_s2"] = judul_pilihan_s2
                    st.session_state["masalah_proposal_s2"] = masalah_ide_s2
                    st.session_state["_judul_proposal_s2_sumber"] = judul_pilihan_s2
                    st.session_state["_masalah_proposal_s2_sumber"] = masalah_ide_s2
                    st.success(
                        "Judul utama sudah ditetapkan sebagai dasar Proposal Tesis. "
                        "Permasalahan dan arah penelitian ikut disimpan."
                    )
        else:
            st.info(
                "Isi atau salin terlebih dahulu beberapa alternatif judul dari hasil analisis AI."
            )


    # ============================================================
    # SUBMENU 2 — PROPOSAL TESIS
    # Alur: Pendahuluan → Kajian Pustaka → Metode → Siap Seminar
    # Setiap bagian: Generate AI → Edit → Koreksi Ulang → Simpan
    # ============================================================
    if submenu_s2 == "📑 Proposal Tesis":
        st.markdown("### 📑 Ruang Kerja Proposal Tesis")

        # Semua dasar Proposal diambil otomatis dari hasil final Submenu 1.
        _proyek = st.session_state.get("proyek_tesis_s2", {})
        _dasar = st.session_state.get("dasar_proposal_tesis_s2", {})
        if not isinstance(_proyek, dict):
            _proyek = {}
        if not isinstance(_dasar, dict):
            _dasar = {}

        _judul_prop = _proyek.get("judul") or _dasar.get("judul") or st.session_state.get("judul_tesis_s2_terpilih", "")
        _masalah_prop = _proyek.get("masalah") or _dasar.get("masalah", "")
        _arah_prop = _proyek.get("arah") or _dasar.get("arah", "")
        _metode_prop = _proyek.get("metode") or st.session_state.get("metode_ide_s2", "Belum ditentukan")
        _pedoman_prop = st.session_state.get("pedoman_tesis_s2_analisis", "") if st.session_state.get("pedoman_tesis_s2_aktif") else ""
        _bank_prop = st.session_state.get("bank_bahan_ide_s2", [])
        _refs_prop = st.session_state.get("bank_referensi", [])

        if not str(_judul_prop).strip():
            st.warning("Tetapkan judul terlebih dahulu pada Submenu 1.")
        else:
            st.success(f"🎓 Judul aktif: {_judul_prop}")
            st.caption(f"Metode: {_metode_prop} • Judul, masalah, pedoman, Bank Bahan, dan referensi digunakan otomatis.")

            with st.expander("➕ Tambah bahan khusus Proposal (opsional)", expanded=False):
                _tambahan_prop = st.file_uploader(
                    "PDF, DOCX, TXT",
                    type=["pdf","docx","txt"],
                    accept_multiple_files=True,
                    key="tambahan_bahan_proposal_s2",
                )

            _tahap_prop = [
                "I. Pendahuluan Penelitian",
                "II. Kajian Pustaka & Kerangka Pikir",
                "III. Metode Penelitian & Sistematika",
                "IV. Simulasi Seminar Proposal",
            ]
            st.markdown("#### Tahapan Proposal")
            for _i,_nama in enumerate(_tahap_prop,1):
                _ok="✅" if st.session_state.get(f"proposal_s2_tahap_{_i}") else "⬜"
                st.write(f"{_ok} {_nama}")

            if st.button("🤖 Generate Proposal Otomatis",key="generate_proposal_otomatis_s2",type="primary",use_container_width=True):
                _tambahan_teks=[]
                for _f in (_tambahan_prop or []):
                    try:
                        _t=ekstrak_teks(_f)
                        if _t and not str(_t).startswith("ERROR:"):
                            _tambahan_teks.append(f"{getattr(_f,'name','dokumen')}:\n{str(_t)[:12000]}")
                    except Exception:
                        pass

                _bank_ringkas="\n\n".join(str(x)[:5000] for x in _bank_prop[:12]) if isinstance(_bank_prop,list) else str(_bank_prop)[:30000]
                _ref_ringkas="\n".join(str(x)[:1200] for x in _refs_prop[:40]) if isinstance(_refs_prop,list) else str(_refs_prop)[:30000]

                _prompt=f"""Anda adalah Asisten Akademik AI untuk menyusun PROPOSAL TESIS S2.
Gunakan seluruh data proyek yang sudah ditetapkan. Jangan meminta pengguna mengisi ulang.

JUDUL:
{_judul_prop}

PERMASALAHAN/KONTEKS:
{_masalah_prop}

ARAH PENELITIAN:
{_arah_prop}

METODE:
{_metode_prop}

PEDOMAN INSTITUSI:
{_pedoman_prop if _pedoman_prop else "Tidak ada pedoman institusi khusus yang aktif."}

BANK BAHAN:
{_bank_ringkas if _bank_ringkas else "Tidak ada dokumen khusus."}

REFERENSI TERSEDIA:
{_ref_ringkas if _ref_ringkas else "Belum ada referensi terverifikasi di Bank Referensi."}

BAHAN TAMBAHAN:
{chr(10).join(_tambahan_teks) if _tambahan_teks else "Tidak ada."}

Susun satu draf Proposal Tesis yang utuh dan saling konsisten dengan struktur:

I. PENDAHULUAN PENELITIAN
- Latar Belakang Masalah
- Rumusan Masalah/Fokus Penelitian
- Tujuan Penelitian
- Signifikansi/Manfaat Penelitian
- Definisi Operasional/Istilah bila relevan

II. KAJIAN PUSTAKA & KERANGKA PIKIR
- Penelitian Terdahulu
- Kajian Teori
- Research Gap
- Novelty
- Kerangka Pikir/Kerangka Konseptual sesuai metode

III. METODE PENELITIAN & SISTEMATIKA
- Pendekatan dan jenis penelitian
- Subjek/objek atau populasi/sampel sesuai metode
- Data/sumber data atau variabel sesuai metode
- Teknik pengumpulan data
- Instrumen bila relevan
- Teknik analisis data
- Keabsahan data/uji instrumen sesuai metode
- Sistematika penulisan

IV. PERSIAPAN SIMULASI SEMINAR PROPOSAL
- Ringkasan proposal
- Alasan pemilihan judul
- Gap dan novelty
- Alasan pemilihan metode
- 10 prediksi pertanyaan penguji dan poin jawaban

ATURAN:
- Ikuti pedoman institusi yang diunggah bila tersedia.
- Sesuaikan metodologi dengan metode yang sudah dipilih.
- Jangan mengarang data lapangan, persentase, hasil penelitian, DOI, halaman, kutipan, atau referensi.
- Gunakan referensi nyata yang tersedia. Bila bukti kurang, beri tanda bahwa bagian perlu diperkuat.
- Jangan membuat BAB IV hasil penelitian atau BAB V kesimpulan penelitian.
- Hasil harus dapat diedit pengguna.
"""
                try:
                    _hasil=panggil_gemini(_prompt)
                    if _hasil:
                        st.session_state["proposal_s2_draf_otomatis"]=str(_hasil)
                        st.session_state["proposal_s2_editor_otomatis"]=str(_hasil)
                        for _i in range(1,5):
                            st.session_state[f"proposal_s2_tahap_{_i}"]=True
                        st.rerun()
                    else:
                        st.error("AI belum menghasilkan draf proposal.")
                except Exception as _e:
                    st.error(f"Proposal belum dapat dibuat: {_e}")

            if st.session_state.get("proposal_s2_draf_otomatis"):
                st.success("✅ Draf proposal selesai.")
                _edit=st.text_area(
                    "✍️ Draf Proposal — dapat diedit",
                    key="proposal_s2_editor_otomatis",
                    height=700,
                )
                _c1,_c2=st.columns(2)
                with _c1:
                    if st.button("💾 Simpan Revisi",key="simpan_revisi_proposal_s2",use_container_width=True):
                        st.session_state["proposal_s2_draf_otomatis"]=_edit
                        st.success("Revisi disimpan.")
                with _c2:
                    if st.button("🤖 Koreksi Ulang AI",key="koreksi_proposal_s2",use_container_width=True):
                        _prompt_k=f"""Review dan perbaiki Proposal Tesis S2 berikut secara utuh.
Pertahankan judul. Periksa konsistensi masalah, rumusan/fokus, tujuan, teori, gap, novelty, metode, dan sistematika.
Ikuti pedoman institusi bila tersedia. Jangan mengarang data, referensi, DOI, halaman, kutipan, atau hasil penelitian.

JUDUL:
{_judul_prop}

PEDOMAN:
{_pedoman_prop if _pedoman_prop else "Tidak ada pedoman khusus."}

DRAF:
{_edit}
"""
                        try:
                            _k=panggil_gemini(_prompt_k)
                            if _k:
                                st.session_state["proposal_s2_draf_otomatis"]=str(_k)
                                st.session_state["proposal_s2_editor_otomatis"]=str(_k)
                                st.rerun()
                        except Exception as _e:
                            st.error(f"Koreksi belum dapat dilakukan: {_e}")

                if st.button("✅ Finalisasi Proposal",key="finalisasi_proposal_s2",type="primary",use_container_width=True):
                    st.session_state["proposal_s2_final"]=_edit
                    st.success("🔒 Proposal ditetapkan sebagai versi final.")

    # No.1 dan Proposal memakai workspace khusus di atas.
    # Generator generik hanya tampil pada submenu 3-12.
    if submenu_s2 not in [
        "💡 Pencarian Ide & Pengajuan Judul 🌟",
        "📑 Proposal Tesis",
    ]:
        pilihan_tahap_s2 = tahap_per_submenu_s2[submenu_s2]
        if len(pilihan_tahap_s2) == 1:
            tahap = pilihan_tahap_s2[0]
            st.caption(f"Tahap aktif: {tahap}")
        else:
            tahap = st.selectbox(
                "Pilih bagian yang dikerjakan",
                pilihan_tahap_s2,
                key=f"tahap_s2_{submenu_s2}",
            )
        metode=st.selectbox("Jenis Penelitian",["Belum menentukan metode","Kuantitatif","Kualitatif","Mixed Methods","R&D / Pengembangan","PTK","Studi Literatur / Library Research","SLR","Evaluatif","Analisis Isi"],key="metode_s2")
        pedoman=st.file_uploader("📄 Unggah pedoman kampus (opsional)",type=["pdf","docx","txt"],key="pedoman_s2")
        sumber=st.file_uploader("📚 Unggah tesis terdahulu, jurnal, bahan, atau data",type=["pdf","docx","txt","csv","xlsx"],accept_multiple_files=True,key="tesis")
        arah=st.text_area("Ide, masalah, arahan pembimbing, atau pekerjaan yang ingin dibuat",key="arah_s2")
        konteks=""
        if pedoman:
            t=ekstrak_teks(pedoman)
            if not t.startswith("ERROR:"): konteks+="\nPEDOMAN KAMPUS:\n"+t
        for f in sumber or []:
            if f.name.lower().endswith((".pdf",".docx",".txt")):
                t=ekstrak_teks(f)
                if not t.startswith("ERROR:"): konteks+=f"\nSUMBER {f.name}:\n{t}"
        st.info("AI menggunakan alur sumber → analisis → draf → sitasi → verifikasi → revisi. BAB IV hanya dari data nyata.")
        if st.button("🤖 Generate AI Tesis S2",type="primary",key="generate_s2"):
            instr=f"""Jenjang: S2 — Tesis
    Metode: {metode}
    Tahap: {tahap}
    Arahan: {arah}
    Tunjukkan kedalaman analisis tingkat magister. Untuk gap/novelty, dasarkan pada bahan dan referensi yang tersedia. Jangan membuat data, DOI, kutipan, atau halaman palsu."""
            panel_ai_penulisan(konteks,tahap,instr,st.session_state.bank_referensi,"s2")
        _hasil_aktif_s2 = (
            st.session_state.get("hasil_ai_ide_judul_s2", "")
            if submenu_s2 == "💡 Pencarian Ide & Pengajuan Judul 🌟"
            else st.session_state.get("hasil_penulisan_ai", "")
        )
        if _hasil_aktif_s2:
            edit = st.text_area(
                "✍️ Hasil AI — dapat diedit",
                _hasil_aktif_s2,
                height=650,
                key="hasil_s2"
            )
            st.session_state.naskah_aktif = edit

            st.caption(
                "Edit hasil AI langsung di atas. Setelah selesai, gunakan Koreksi Ulang AI. "
                "AI menilai versi terbaru dan tidak mengganti tulisan Anda secara otomatis."
            )

            col_koreksi, col_final = st.columns(2)

            with col_koreksi:
                if st.button(
                    "🔍 Koreksi Ulang Hasil Edit",
                    key="koreksi_ulang_s2",
                    use_container_width=True
                ):
                    if not edit.strip():
                        st.warning("Belum ada teks yang dapat dikoreksi.")
                    else:
                        instr_koreksi = f"""Jenjang: S2 — Tesis
    Submenu: {submenu_s2}
    Tahap: {tahap}
    Jenis penelitian: {metode}

    Tugas Anda adalah menjadi reviewer akademik tesis tingkat magister.
    Periksa NASKAH VERSI TERBARU yang sudah diedit pengguna.

    Periksa secara menyeluruh:
    1. kesesuaian isi dengan fokus/judul dan tahap tesis;
    2. struktur akademik dan kelogisan argumentasi;
    3. koherensi antarparagraf dan konsistensi istilah;
    4. bahasa akademik, tata bahasa, dan kejelasan kalimat;
    5. kesesuaian metodologi bila bagian berkaitan dengan metode;
    6. konsistensi rumusan masalah, tujuan, teori, metode, hasil, dan kesimpulan bila tersedia;
    7. klaim yang membutuhkan referensi;
    8. relevansi referensi/sitasi terhadap klaim;
    9. konsistensi sitasi, footnote, dan daftar pustaka bila tersedia;
    10. kelemahan substantif yang masih perlu diperbaiki.

    ATURAN:
    - Jangan membuat data, DOI, halaman, kutipan, atau referensi palsu.
    - Jangan mengubah fakta penelitian pengguna.
    - Pertahankan maksud asli naskah.
    - Berikan saran terlebih dahulu. Jangan mengganti naskah pengguna secara diam-diam.

    Susun hasil:
    A. Ringkasan penilaian
    B. Bagian yang perlu diperbaiki
    C. Usulan perbaikan
    D. Versi revisi yang disarankan
    E. Catatan referensi/sitasi yang perlu diverifikasi

    NASKAH VERSI TERBARU:
    {edit}
    """
                        # Tetap memakai mesin AI lama yang sudah berfungsi.
                        panel_ai_penulisan(
                            edit,
                            f"Koreksi Ulang — {tahap}",
                            instr_koreksi,
                            st.session_state.bank_referensi,
                            "s2_koreksi"
                        )
                        st.session_state["hasil_koreksi_s2"] = st.session_state.get(
                            "hasil_penulisan_ai", ""
                        )

            with col_final:
                if st.button(
                    "✅ Tetapkan Versi Edit sebagai Final",
                    key="final_s2",
                    use_container_width=True
                ):
                    if edit.strip():
                        st.session_state["final_tesis_s2"] = edit
                        st.success("Versi edit terbaru ditetapkan sebagai versi final.")
                    else:
                        st.warning("Belum ada teks untuk difinalisasi.")

            if st.session_state.get("hasil_koreksi_s2"):
                st.markdown("### 🤖 Hasil Koreksi Ulang AI")
                hasil_koreksi_edit = st.text_area(
                    "Hasil koreksi juga dapat diedit",
                    st.session_state["hasil_koreksi_s2"],
                    height=500,
                    key="hasil_koreksi_s2_edit"
                )
                st.session_state["hasil_koreksi_s2"] = hasil_koreksi_edit

            if st.session_state.get("final_tesis_s2"):
                st.markdown("### ✅ Versi Final")
                final_edit = st.text_area(
                    "Versi final tetap dapat diedit bila masih diperlukan",
                    st.session_state["final_tesis_s2"],
                    height=500,
                    key="final_tesis_s2_edit"
                )
                st.session_state["final_tesis_s2"] = final_edit


# ============================================================
# DISERTASI S3
# ============================================================
elif menu == "🧑‍🎓 Disertasi S3":
    st.header("🧑‍🎓 Asisten Disertasi S3")
    tahap=st.selectbox("Tahap Disertasi",[
        "Jembatan Tesis S2 → S3","Analisis Tesis S2","Keterbatasan Penelitian S2","Pertanyaan Penelitian Lanjutan",
        "Topik Doktoral","State of the Art","Research Gap","Novelty Doktoral","Kontribusi Teoretis",
        "Kontribusi Metodologis","Kontribusi Praktis","Proposal Disertasi","Metodologi Doktoral","Instrumen",
        "Pengumpulan Data","Analisis Data","Penulisan Disertasi","Publikasi","Bimbingan & Revisi","Presentasi","Persiapan Ujian Doktoral"
    ],key="tahap_s3")
    sumber=st.file_uploader("📚 Unggah tesis S2, artikel, jurnal, pedoman, atau data",type=["pdf","docx","txt","csv","xlsx"],accept_multiple_files=True,key="disertasi")
    arah=st.text_area("Masalah doktoral, arahan promotor, atau pekerjaan yang ingin dibuat",key="arah_s3")
    konteks=""
    for f in sumber or []:
        if f.name.lower().endswith((".pdf",".docx",".txt")):
            t=ekstrak_teks(f)
            if not t.startswith("ERROR:"): konteks+=f"\nSUMBER {f.name}:\n{t}"
    st.info("Novelty dan kontribusi doktoral harus ditelusuri dari bukti/sumber yang tersedia, bukan dibuat oleh AI.")
    if st.button("🤖 Generate AI Disertasi S3",type="primary",key="generate_s3"):
        instr=f"""Jenjang: S3 — Disertasi
Tahap: {tahap}
Arahan: {arah}
Gunakan analisis doktoral yang kritis. Bedakan state of the art, research gap, novelty, serta kontribusi teoretis/metodologis/praktis. Jangan membuat data, DOI, kutipan, atau halaman palsu."""
        panel_ai_penulisan(konteks,tahap,instr,st.session_state.bank_referensi,"s3")
    if st.session_state.get("hasil_penulisan_ai"):
        edit=st.text_area("Hasil AI — dapat diedit",st.session_state.hasil_penulisan_ai,height=650,key="hasil_s3")
        st.session_state.naskah_aktif=edit


# ============================================================
# METODOLOGI
# ============================================================
elif menu == "🧭 Metodologi Penelitian":

    st.header("🧭 Penentu Jenis & Metodologi Penelitian")

    jenjang = st.radio(
        "Jenjang",
        ["S1 — Skripsi", "S2 — Tesis", "S3 — Disertasi"],
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
            "Ringkasan Akademik",
            "Pemeriksa Kemiripan Internal",
            "Penyunting Akademik"
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
# PUBLIKASI / JURNAL AKADEMIK
# ============================================================
elif menu == "📑 Publikasi Jurnal":

    st.header("📝 Jurnal Akademik")
    st.caption("Ruang kerja artikel jurnal dari penentuan rumah jurnal sampai audit naskah sebelum submission.")

    bagian_jurnal = st.radio(
        "Bagian Jurnal Akademik",
        [
            "🏠 Rumah Jurnal",
            "📤 Template & Author Guidelines",
            "✍️ Tulis / Adaptasi Artikel",
            "🔎 Cek Kesesuaian Rumah Jurnal",
            "📚 Referensi & Sitasi",
            "✅ Audit Artikel"
        ],
        key="bagian_jurnal_akademik"
    )

    if bagian_jurnal == "🏠 Rumah Jurnal":
        st.subheader("🏠 Rumah Jurnal")
        nama_jurnal = st.text_input("Nama jurnal tujuan", key="nama_rumah_jurnal")
        url_jurnal = st.text_input("URL jurnal / halaman author guidelines", key="url_rumah_jurnal")
        scope_jurnal = st.text_area("Focus & Scope / ketentuan utama jurnal", height=180, key="scope_rumah_jurnal")
        target_jurnal = st.selectbox("Target / indeks", ["Belum ditentukan","Jurnal Nasional","SINTA 6","SINTA 5","SINTA 4","SINTA 3","SINTA 2","SINTA 1","Scopus"], key="target_rumah_jurnal")
        if st.button("🤖 Analisis Kesesuaian Rumah Jurnal", type="primary", key="ai_rumah_jurnal"):
            konteks = f"Nama jurnal: {nama_jurnal}\nURL: {url_jurnal}\nTarget: {target_jurnal}\nFocus & Scope/Ketentuan:\n{scope_jurnal}"
            panel_ai_penulisan(konteks, "Analisis rumah jurnal", "Analisis kesesuaian topik, scope, struktur, gaya penulisan, dan hal yang perlu dipenuhi. Jangan mengarang ketentuan yang tidak diberikan.", st.session_state.bank_referensi, "rumah_jurnal")

    elif bagian_jurnal == "📤 Template & Author Guidelines":
        st.subheader("📤 Template & Author Guidelines")
        files_jurnal = st.file_uploader("Unggah template / author guidelines PDF, DOCX, atau TXT", type=["pdf","docx","txt"], accept_multiple_files=True, key="template_jurnal_upload")
        teks_template = ""
        if files_jurnal:
            for f in files_jurnal:
                try:
                    teks_template += f"\n\nFILE: {f.name}\n" + ekstrak_teks(f)
                except Exception as e:
                    st.warning(f"{f.name} belum dapat dibaca: {e}")
        if st.button("🤖 Analisis Template Jurnal", type="primary", disabled=not bool(teks_template.strip()), key="ai_template_jurnal"):
            panel_ai_penulisan(teks_template, "Analisis template dan author guidelines jurnal", "Ekstrak struktur artikel, batasan, format, gaya sitasi, tabel/gambar, abstrak, kata kunci, dan checklist penulisan. Gunakan hanya ketentuan yang tersedia dalam dokumen.", [], "template_jurnal")

    elif bagian_jurnal == "✍️ Tulis / Adaptasi Artikel":
        st.subheader("✍️ Tulis / Adaptasi Artikel")
        sumber_artikel = st.selectbox("Sumber artikel", ["Artikel baru","Tugas Kuliah","Skripsi S1","Tesis S2","Disertasi S3"], key="sumber_artikel_jurnal")
        file_artikel = st.file_uploader("Unggah naskah sumber PDF/DOCX/TXT (opsional)", type=["pdf","docx","txt"], key="sumber_artikel_upload")
        teks_artikel = ekstrak_teks(file_artikel) if file_artikel else ""
        tema_artikel = st.text_area("Judul/tema, temuan utama, atau arahan penulisan", height=180, key="tema_artikel_jurnal")
        if st.button("🤖 Generate Artikel Jurnal", type="primary", disabled=not bool(tema_artikel.strip() or teks_artikel.strip()), key="generate_artikel_jurnal"):
            konteks = f"SUMBER: {sumber_artikel}\nARAHAN: {tema_artikel}\nNASKAH SUMBER:\n{teks_artikel[:60000]}"
            panel_ai_penulisan(konteks, "Artikel jurnal akademik", "Susun artikel akademik yang koheren. Jangan menciptakan data penelitian. Gunakan referensi Library yang relevan dan terverifikasi.", st.session_state.bank_referensi, "artikel_jurnal")

    elif bagian_jurnal == "🔎 Cek Kesesuaian Rumah Jurnal":
        st.subheader("🔎 Cek Kesesuaian Rumah Jurnal")
        fcek = st.file_uploader("Unggah artikel PDF/DOCX/TXT", type=["pdf","docx","txt"], key="cek_jurnal_file")
        aturan = st.text_area("Tempel ketentuan / focus & scope rumah jurnal", height=180, key="cek_jurnal_aturan")
        if st.button("🤖 Cek Kesesuaian dengan AI", type="primary", disabled=fcek is None, key="ai_cek_jurnal"):
            teks = ekstrak_teks(fcek)
            panel_ai_penulisan(teks, "Audit kesesuaian artikel dengan rumah jurnal", f"Bandingkan artikel dengan ketentuan berikut dan buat tabel Sesuai/Perlu Revisi/Tidak Ditemukan. Jangan mengarang aturan.\n{aturan}", st.session_state.bank_referensi, "cek_jurnal")

    elif bagian_jurnal == "📚 Referensi & Sitasi":
        st.subheader("📚 Referensi & Sitasi")
        st.info("Menggunakan Library yang sama dengan menu Literatur & Referensi.")
        if st.session_state.bank_referensi:
            st.dataframe(pd.DataFrame(st.session_state.bank_referensi), use_container_width=True, hide_index=True)
        else:
            st.warning("Library Referensi masih kosong. Tambahkan referensi melalui menu Literatur & Referensi.")

    elif bagian_jurnal == "✅ Audit Artikel":
        st.subheader("✅ Audit Artikel")
        faudit = st.file_uploader("Unggah artikel PDF/DOCX/TXT", type=["pdf","docx","txt"], key="audit_artikel_jurnal")
        if st.button("🤖 Audit Artikel dengan AI", type="primary", disabled=faudit is None, key="ai_audit_artikel"):
            teks = ekstrak_teks(faudit)
            panel_ai_penulisan(teks, "Audit artikel jurnal", "Audit judul, abstrak, kata kunci, pendahuluan, metode, hasil/pembahasan, simpulan, sitasi, daftar pustaka, konsistensi bahasa, dan kesiapan submission. Jangan membuat data atau sumber baru.", st.session_state.bank_referensi, "audit_artikel")

    if st.session_state.get("hasil_penulisan_ai"):
        st.divider()
        st.text_area("📄 Hasil AI — dapat diedit", st.session_state.hasil_penulisan_ai, height=550, key="hasil_ai_jurnal_tampil")


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
    "Asisten Akademik AI — S1 • S2 • S3 • OBE • penelitian • "
    "referensi tervalidasi • publikasi • buku • presentasi • sidang."
)
