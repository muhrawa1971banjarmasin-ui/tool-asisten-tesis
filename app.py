rom pathlib import Path
from copy import deepcopy
from io import BytesIO

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
    from docx.shared import Cm, Pt
    from docx.enum.text import WD_ALIGN_PARAGRAPH
except ImportError:
    docx = None
    Cm = Pt = WD_ALIGN_PARAGRAPH = None

try:
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.enums import TA_CENTER
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
    from reportlab.lib import colors
    from reportlab.lib.units import cm as rl_cm
except ImportError:
    A4 = getSampleStyleSheet = ParagraphStyle = TA_CENTER = None
    SimpleDocTemplate = Paragraph = Spacer = Table = TableStyle = PageBreak = None
    colors = rl_cm = None


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


def hasil_ai_teks(hasil):
    """Ambil teks dari respons AI tanpa pernah menampilkan dict internal ke naskah."""
    if isinstance(hasil, dict):
        return str(hasil.get("hasil", "") or "").strip() if hasil.get("sukses") else ""
    return str(hasil or "").strip()

def error_ai_teks(hasil):
    if isinstance(hasil, dict):
        return str(hasil.get("error", "") or "").strip()
    return ""


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

def _jumlah_hasil_crossref(kata_kunci):
    if not str(kata_kunci or "").strip(): return 0
    try:
        q=urllib.parse.quote(str(kata_kunci).strip())
        data=_http_json(f"https://api.crossref.org/works?query.bibliographic={q}&rows=0")
        return int(data.get("message",{}).get("total-results",0) or 0)
    except Exception:
        return 0

def _jumlah_hasil_openalex(kata_kunci):
