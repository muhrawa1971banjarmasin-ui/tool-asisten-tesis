from pathlib import Path
from copy import deepcopy
from io import BytesIO

import streamlit as st
import streamlit.components.v1 as components
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



def tampilkan_preview_word_rapi(teks, tinggi=900, judul="Preview Dokumen"):
    """Pratinjau bergaya Print Layout Word langsung di Streamlit."""
    raw = _bersihkan_hasil_sunting_ai(str(teks or "")) if '_bersihkan_hasil_sunting_ai' in globals() else str(teks or "")
    lines = [x.rstrip() for x in raw.splitlines()]
    blocks=[]
    for line in lines:
        t=line.strip()
        if not t:
            blocks.append('<div class="spacer"></div>'); continue
        plain=re.sub(r'[*_#]','',t).strip(); up=plain.upper(); esc=html.escape(plain)
        if re.match(r'^BAB\s+[IVXLCDM]+\b',up): blocks.append(f'<div class="bab">{esc}</div>')
        elif up in ('PENDAHULUAN','KAJIAN PUSTAKA DAN KERANGKA PIKIR','METODE PENELITIAN','KATA PENGANTAR','DAFTAR ISI','DAFTAR TABEL','DAFTAR GAMBAR','DAFTAR LAMPIRAN','SISTEMATIKA PENULISAN','DAFTAR PUSTAKA','DAFTAR PUSTAKA SEMENTARA'): blocks.append(f'<div class="heading">{esc}</div>')
        elif re.match(r'^[A-Z]\.\s+',plain): blocks.append(f'<div class="subheading">{esc}</div>')
        elif re.match(r'^\d+[\.)]\s+',plain): blocks.append(f'<div class="numbered">{esc}</div>')
        elif re.search(r'[\u0600-\u06FF]',plain): blocks.append(f'<div class="arabic">{esc}</div>')
        else: blocks.append(f'<div class="para">{esc}</div>')
    body=''.join(f'<div class="src-block">{b}</div>' for b in blocks)
    doc=f'''<!doctype html><html><head><meta charset="utf-8"><style>
    *{{box-sizing:border-box}} body{{margin:0;background:#e7e9ed;font-family:"Times New Roman",serif;color:#111}}
    .toolbar{{position:sticky;top:0;z-index:5;background:#fff;border-bottom:1px solid #cfd3d8;padding:10px 16px;font-family:Arial,sans-serif;font-size:14px}}
    #source{{display:none}} #pages{{padding:22px 0 40px}}
    .page{{width:794px;height:1123px;margin:0 auto 22px;background:#fff;box-shadow:0 2px 10px rgba(0,0,0,.18);padding:151px 113px 113px 151px;overflow:hidden;position:relative}}
    .page-no{{position:absolute;right:55px;top:42px;font:12px Arial;color:#777}}
    .para,.numbered,.subheading,.heading,.bab,.arabic{{font-size:16px;line-height:2;text-align:justify;margin:0}}
    .para{{text-indent:48px}} .numbered{{padding-left:28px;text-indent:-28px}}
    .subheading{{font-weight:bold;text-align:left;margin-top:2px}} .heading,.bab{{font-weight:bold;text-align:center;text-indent:0}}
    .arabic{{direction:rtl;text-align:right;font-family:"Traditional Arabic","Amiri","Noto Naskh Arabic",serif;font-size:20px;line-height:1.8}}
    .spacer{{height:10px}} .src-block{{break-inside:avoid}}
    @media(max-width:900px){{.page{{transform-origin:top center;zoom:.78}}}}
    </style></head><body><div class="toolbar"><b>👁️ {html.escape(judul)}</b> &nbsp; • &nbsp; Tampilan A4 sebelum file Word diunduh</div>
    <div id="source">{body}</div><div id="pages"></div>
    <script>
    const src=[...document.querySelectorAll('#source .src-block')]; const pages=document.getElementById('pages');
    function newPage(){{const p=document.createElement('div');p.className='page';pages.appendChild(p);return p;}}
    let page=newPage();
    src.forEach((b)=>{{let c=b.cloneNode(true);page.appendChild(c);if(page.scrollHeight>page.clientHeight){{page.removeChild(c);page=newPage();page.appendChild(c);}}}});
    [...pages.children].forEach((p,i)=>{{let n=document.createElement('div');n.className='page-no';n.textContent=(i+1);p.appendChild(n);}});
    </script></body></html>'''
    components.html(doc, height=tinggi, scrolling=True)

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
    model_ids = ["gemini-3.8-flash", "gemini-3.5-flash-lite", "gemini-2.5-flash"]
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
5. NASKAH UNGGAHAN ADALAH MASTER GAYA SITASI. Deteksi gaya sitasi yang benar-benar dipakai pada naskah asli (misalnya Chicago footnote, APA author-date, atau gaya lain), lalu KUNCI gaya tersebut. Jangan memaksa Chicago, APA, atau gaya lain jika berbeda dari naskah asli.
5a. Semua kutipan/sitasi baru dan kutipan yang formatnya menyimpang akibat proses AI WAJIB dinormalisasi mengikuti gaya sitasi asli naskah, TANPA mengubah identitas sumber, isi kutipan, atau sumber yang sudah ada.
6. Jika gaya asli memakai footnote/endnote, pertahankan semua footnote lama dan lanjutkan penomoran secara unik sesuai urutan kemunculan; jangan mulai lagi dari 1 dan jangan membuat nomor ganda. Jika gaya asli memakai sitasi dalam teks, pertahankan sistem in-text tersebut.
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
    if not str(kata_kunci or "").strip(): return 0
    try:
        q=urllib.parse.quote(str(kata_kunci).strip())
        data=_http_json(f"https://api.openalex.org/works?search={q}&per-page=1")
        return int((data.get("meta") or {}).get("count",0) or 0)
    except Exception:
        return 0

def _jumlah_hasil_semantic_scholar(kata_kunci):
    if not str(kata_kunci or "").strip(): return 0
    try:
        q=urllib.parse.quote(str(kata_kunci).strip())
        data=_http_json(f"https://api.semanticscholar.org/graph/v1/paper/search?query={q}&limit=1&fields=title")
        return int(data.get("total",0) or 0)
    except Exception:
        return 0

def analisis_ketersediaan_referensi_judul_s2(judul, jumlah=20):
    """Ketersediaan literatur nyata. Angka basis data tidak dijumlahkan karena bisa tumpang tindih."""
    refs = cari_multi_sumber(judul, jumlah)
    tahun_sekarang = datetime.now().year
    terbaru = 0
    internasional = 0
    terverifikasi = 0
    for r in refs:
        try:
            if int(str(r.get("Tahun", ""))[:4]) >= tahun_sekarang - 5:
                terbaru += 1
        except Exception:
            pass
        if r.get("Sumber") in ("Crossref","OpenAlex","Semantic Scholar","Library of Congress"):
            internasional += 1
        if "terverifikasi" in str(r.get("Status","")).lower() or "teridentifikasi" in str(r.get("Status","")).lower():
            terverifikasi += 1
    cr=_jumlah_hasil_crossref(judul)
    oa=_jumlah_hasil_openalex(judul)
    ss=_jumlah_hasil_semantic_scholar(judul)
    # Jangan menjumlahkan basis data karena duplikasi lintas indeks. Gunakan angka terbesar sebagai indikator cakupan.
    indikator=max(cr,oa,ss,len(refs))
    if indikator >= 100:
        status="🟢 Banyak / sangat mendukung"
    elif indikator >= 30:
        status="🟢 Cukup banyak / mendukung"
    elif indikator >= 10:
        status="🟡 Cukup / perlu perluasan kata kunci"
    else:
        status="🔴 Terbatas / perlu pencarian lebih luas"
    return {
        "Judul":judul,
        "Kandidat":len(refs),
        "Literatur 5 Tahun":terbaru,
        "Sumber Internasional":internasional,
        "Terverifikasi":terverifikasi,
        "Crossref":cr,
        "OpenAlex":oa,
        "Semantic Scholar":ss,
        "Indikator Ketersediaan":indikator,
        "Status":status,
        "Referensi":refs,
    }

def _parse_analisis_10_judul(teks, judul_list):
    teks=str(teks or "")
    out=[]
    for i in range(1,11):
        def ambil(label):
            m=re.search(rf"(?ims)^\s*\[{label}\s*{i}\]\s*[:\-]?\s*(.+?)(?=^\s*\[(?:JUDUL|GAP|NOVELTY|KEKUATAN|ALASAN)\s*\d+\]|\Z)",teks)
            return re.sub(r"\s+"," ",m.group(1)).strip() if m else ""
        j=ambil("JUDUL") or (judul_list[i-1] if i-1 < len(judul_list) else "")
        out.append({"No":i,"Judul":j,"Gap":ambil("GAP"),"Novelty":ambil("NOVELTY"),"Kekuatan":ambil("KEKUATAN"),"Alasan":ambil("ALASAN")})
    return out

def _teks_laporan_analisis_judul_s2(masalah, metode, pedoman, analisis, data10, dipilih=""):
    lines=["LAPORAN ANALISIS IDE DAN ALTERNATIF JUDUL TESIS S2", "", f"Tanggal analisis: {datetime.now().strftime('%d-%m-%Y %H:%M')}", f"Pedoman aktif: {pedoman or 'Belum ada pedoman aktif'}", f"Metode: {metode}", "", "PERMASALAHAN/GAGASAN AWAL", masalah or "-", "", "URAIAN ANALISIS IDE, RESEARCH GAP, DAN NOVELTY", analisis or "-"]
    lines += ["", "ANALISIS 10 ALTERNATIF JUDUL"]
    for d in data10:
        av=d.get("Ketersediaan",{}) or {}
        lines += ["", f"{d.get('No')}. {d.get('Judul','')}", f"Research Gap: {d.get('Gap','-') or '-'}", f"Novelty: {d.get('Novelty','-') or '-'}", f"Kekuatan: {d.get('Kekuatan','-') or '-'}", f"Alasan: {d.get('Alasan','-') or '-'}", f"Ketersediaan literatur terindeks: Crossref {av.get('Crossref',0)} | OpenAlex {av.get('OpenAlex',0)} | Semantic Scholar {av.get('Semantic Scholar',0)}", f"Kandidat unik yang diperiksa: {av.get('Kandidat',0)} | 5 tahun terakhir: {av.get('Literatur 5 Tahun',0)} | Terverifikasi/teridentifikasi: {av.get('Terverifikasi',0)}", f"Status ketersediaan: {av.get('Status','Belum diperiksa')}"]
    if dipilih:
        lines += ["", "JUDUL YANG DIPILIH PENGGUNA", dipilih]
    lines += ["", "CATATAN", "Jumlah hasil Crossref, OpenAlex, dan Semantic Scholar ditampilkan per basis data dan tidak dijumlahkan karena satu karya dapat terindeks pada lebih dari satu basis data."]
    return "\n".join(lines)

def buat_docx_analisis_judul_s2(masalah, metode, pedoman, analisis, data10, dipilih=""):
    if docx is None: return None
    bio=BytesIO(); d=docx.Document(); sec=d.sections[0]
    sec.top_margin=Cm(4); sec.left_margin=Cm(4); sec.bottom_margin=Cm(3); sec.right_margin=Cm(3)
    styles=d.styles
    styles['Normal'].font.name='Times New Roman'; styles['Normal'].font.size=Pt(12)
    p=d.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER; r=p.add_run('LAPORAN ANALISIS IDE DAN ALTERNATIF JUDUL TESIS S2'); r.bold=True; r.font.name='Times New Roman'; r.font.size=Pt(14)
    for label,val in [("Tanggal analisis",datetime.now().strftime('%d-%m-%Y %H:%M')),("Pedoman aktif",pedoman or 'Belum ada pedoman aktif'),("Metode",metode)]:
        p=d.add_paragraph(); p.add_run(label+': ').bold=True; p.add_run(str(val))
    d.add_heading('A. Permasalahan/Gagasan Awal',level=1); d.add_paragraph(masalah or '-')
    d.add_heading('B. Analisis Ide, Research Gap, dan Novelty',level=1)
    for blok in str(analisis or '-').splitlines():
        if blok.strip(): d.add_paragraph(blok.strip())
    d.add_heading('C. Analisis 10 Alternatif Judul',level=1)
    table=d.add_table(rows=1, cols=7); table.style='Table Grid'
    hdr=['No','Judul','Research Gap','Novelty','Kekuatan','Ketersediaan Literatur','Status']
    for i,h in enumerate(hdr): table.rows[0].cells[i].text=h
    for x in data10:
        av=x.get('Ketersediaan',{}) or {}; cells=table.add_row().cells
        vals=[str(x.get('No','')),x.get('Judul',''),x.get('Gap','-') or '-',x.get('Novelty','-') or '-',x.get('Kekuatan','-') or '-',f"Crossref {av.get('Crossref',0)}; OpenAlex {av.get('OpenAlex',0)}; S2 {av.get('Semantic Scholar',0)}; sampel unik {av.get('Kandidat',0)}",av.get('Status','Belum diperiksa')]
        for i,v in enumerate(vals): cells[i].text=str(v)
    d.add_paragraph('Catatan: jumlah hasil tiap basis data tidak dijumlahkan karena satu karya dapat terindeks pada lebih dari satu basis data.')
    if dipilih:
        d.add_heading('D. Judul yang Dipilih Pengguna',level=1); d.add_paragraph(dipilih)
    d.save(bio); bio.seek(0); return bio.getvalue()

def buat_pdf_analisis_judul_s2(masalah, metode, pedoman, analisis, data10, dipilih=""):
    # Utamakan ReportLab. Jika belum dipasang di Streamlit Cloud, gunakan PDF teks sederhana
    # agar tombol PDF tetap berfungsi tanpa menambah dependency baru.
    if SimpleDocTemplate is None:
        teks=_teks_laporan_analisis_judul_s2(masalah,metode,pedoman,analisis,data10,dipilih)
        # PDF core font Helvetica mendukung Latin-1. Karakter di luar itu dinormalisasi aman.
        teks=(teks.replace('–','-').replace('—','-').replace('“','"').replace('”','"').replace('’',"'")
                  .replace('🟢','[KUAT]').replace('🟡','[CUKUP]').replace('🔴','[LEMAH]'))
        raw=teks.encode('latin-1','replace').decode('latin-1')
        # Bungkus baris agar tidak terpotong di sisi kanan.
        lines=[]
        for para in raw.splitlines():
            words=para.split(); cur=''
            if not words: lines.append(''); continue
            for w in words:
                if len(cur)+len(w)+1>92:
                    lines.append(cur); cur=w
                else: cur=(cur+' '+w).strip()
            if cur: lines.append(cur)
        pages=[lines[i:i+48] for i in range(0,len(lines),48)] or [[]]
        objects=[]
        # 1 catalog, 2 pages, 3 font, then page/content pairs
        kids=[]
        for idx,page_lines in enumerate(pages):
            page_obj=4+idx*2; content_obj=page_obj+1; kids.append(f'{page_obj} 0 R')
            y=790; cmds=['BT','/F1 10 Tf','14 TL']
            for line in page_lines:
                esc=line.replace('\\','\\\\').replace('(','\\(').replace(')','\\)')
                cmds.append(f'1 0 0 1 70 {y} Tm ({esc}) Tj'); y-=14
            cmds.append('ET'); stream='\n'.join(cmds).encode('latin-1','replace')
            objects.append((page_obj,f'<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] /Resources << /Font << /F1 3 0 R >> >> /Contents {content_obj} 0 R >>'.encode('latin-1')))
            objects.append((content_obj,b'<< /Length '+str(len(stream)).encode()+b' >>\nstream\n'+stream+b'\nendstream'))
        base=[(1,b'<< /Type /Catalog /Pages 2 0 R >>'),(2,f'<< /Type /Pages /Kids [{" ".join(kids)}] /Count {len(pages)} >>'.encode('latin-1')),(3,b'<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>')]
        allobjs=sorted(base+objects,key=lambda x:x[0]); out=bytearray(b'%PDF-1.4\n'); offsets={0:0}
        for num,body in allobjs:
            offsets[num]=len(out); out.extend(f'{num} 0 obj\n'.encode()); out.extend(body); out.extend(b'\nendobj\n')
        xref=len(out); maxobj=max(offsets)
        out.extend(f'xref\n0 {maxobj+1}\n'.encode()); out.extend(b'0000000000 65535 f \n')
        for i in range(1,maxobj+1): out.extend(f'{offsets.get(i,0):010d} 00000 n \n'.encode())
        out.extend(f'trailer\n<< /Size {maxobj+1} /Root 1 0 R >>\nstartxref\n{xref}\n%%EOF'.encode())
        return bytes(out)
    bio=BytesIO(); doc=SimpleDocTemplate(bio,pagesize=A4,rightMargin=3*rl_cm,leftMargin=4*rl_cm,topMargin=4*rl_cm,bottomMargin=3*rl_cm)
    styles=getSampleStyleSheet(); normal=ParagraphStyle('BodyTNR',parent=styles['BodyText'],fontName='Times-Roman',fontSize=11,leading=16); title=ParagraphStyle('TitleTNR',parent=styles['Title'],fontName='Times-Bold',fontSize=14,alignment=TA_CENTER,leading=18)
    story=[Paragraph('LAPORAN ANALISIS IDE DAN ALTERNATIF JUDUL TESIS S2',title),Spacer(1,10),Paragraph(f"Tanggal analisis: {datetime.now().strftime('%d-%m-%Y %H:%M')}",normal),Paragraph(f"Pedoman aktif: {html.escape(pedoman or 'Belum ada pedoman aktif')}",normal),Paragraph(f"Metode: {html.escape(str(metode))}",normal),Spacer(1,8),Paragraph('<b>A. Permasalahan/Gagasan Awal</b>',normal),Paragraph(html.escape(masalah or '-').replace('\n','<br/>'),normal),Spacer(1,8),Paragraph('<b>B. Analisis Ide, Research Gap, dan Novelty</b>',normal),Paragraph(html.escape(analisis or '-').replace('\n','<br/>'),normal),Spacer(1,8),Paragraph('<b>C. Analisis 10 Alternatif Judul</b>',normal)]
    for x in data10:
        av=x.get('Ketersediaan',{}) or {}
        story += [Spacer(1,7),Paragraph(f"<b>{x.get('No')}. {html.escape(x.get('Judul',''))}</b>",normal),Paragraph(f"Research Gap: {html.escape(x.get('Gap','-') or '-')}",normal),Paragraph(f"Novelty: {html.escape(x.get('Novelty','-') or '-')}",normal),Paragraph(f"Kekuatan: {html.escape(x.get('Kekuatan','-') or '-')}",normal),Paragraph(f"Ketersediaan: Crossref {av.get('Crossref',0)} | OpenAlex {av.get('OpenAlex',0)} | Semantic Scholar {av.get('Semantic Scholar',0)} | sampel unik {av.get('Kandidat',0)}",normal),Paragraph(f"Status: {html.escape(av.get('Status','Belum diperiksa'))}",normal)]
    story += [Spacer(1,8),Paragraph('Catatan: jumlah hasil tiap basis data tidak dijumlahkan karena satu karya dapat terindeks pada lebih dari satu basis data.',normal)]
    if dipilih: story += [Spacer(1,8),Paragraph('<b>D. Judul yang Dipilih Pengguna</b>',normal),Paragraph(html.escape(dipilih),normal)]
    doc.build(story); bio.seek(0); return bio.getvalue()


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


def buat_docx_proposal_final(teks, judul, nama="", npm="", prodi="Pendidikan Agama Islam", tahun=None,
                              gaya_sitasi="Chicago Notes & Bibliography (Footnote)",
                              font_naskah="Times New Roman", ukuran_naskah="12 pt"):
    """Membuat DOCX proposal yang rapi; true footnote dipakai untuk Chicago/Turabian."""
    if docx is None:
        return None
    import re, io, zipfile
    from lxml import etree

    def _pt_size(v, default=12):
        m = re.search(r'(\d+)', str(v or ''))
        return int(m.group(1)) if m else default

    font_body = str(font_naskah or 'Times New Roman')
    size_body = _pt_size(ukuran_naskah, 12)
    is_notes_style = ('chicago' in str(gaya_sitasi).lower() or 'turabian' in str(gaya_sitasi).lower())

    bio = BytesIO()
    d = docx.Document()
    sec = d.sections[0]
    sec.top_margin = Cm(4); sec.left_margin = Cm(4); sec.bottom_margin = Cm(3); sec.right_margin = Cm(3)

    normal = d.styles['Normal']
    normal.font.name = font_body; normal.font.size = Pt(size_body)
    normal.paragraph_format.line_spacing = 2
    normal.paragraph_format.first_line_indent = Cm(1.27)
    normal.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    # Sampul dibuat oleh Word exporter agar tidak bercampur dengan badan proposal.
    p0 = d.add_paragraph(); p0.alignment = WD_ALIGN_PARAGRAPH.CENTER; p0.paragraph_format.first_line_indent = Cm(0)
    r = p0.add_run('PROPOSAL TESIS'); r.bold = True; r.font.name = font_body; r.font.size = Pt(14)
    for _ in range(2): d.add_paragraph('')
    pj = d.add_paragraph(); pj.alignment = WD_ALIGN_PARAGRAPH.CENTER; pj.paragraph_format.first_line_indent = Cm(0)
    rr = pj.add_run(str(judul or '').upper()); rr.bold = True; rr.font.name = font_body; rr.font.size = Pt(14)
    for _ in range(5): d.add_paragraph('')
    cover_lines = [
        f'Oleh: {nama}' if nama else 'Oleh:', f'NPM: {npm}' if npm else 'NPM:', '',
        'PROGRAM PASCASARJANA', f'PROGRAM STUDI {prodi.upper()}',
        'INSTITUT AGAMA ISLAM DARUSSALAM MARTAPURA', 'MARTAPURA', str(tahun or datetime.now().year)
    ]
    for line in cover_lines:
        pp = d.add_paragraph(); pp.alignment = WD_ALIGN_PARAGRAPH.CENTER; pp.paragraph_format.first_line_indent = Cm(0)
        run = pp.add_run(line); run.font.name = font_body; run.font.size = Pt(size_body)
        run.bold = line in ['PROGRAM PASCASARJANA', f'PROGRAM STUDI {prodi.upper()}', 'INSTITUT AGAMA ISLAM DARUSSALAM MARTAPURA']
    d.add_page_break()

    raw = str(teks or '').strip()
    # Format marker yang diwajibkan untuk Chicago/Turabian: [^1] pada narasi dan [^1]: isi catatan.
    note_pat = re.compile(r'(?m)^\s*\[\^(\d+)\]\s*:?[ \t]*(.+)$')
    notes = {int(m.group(1)): m.group(2).strip() for m in note_pat.finditer(raw)} if is_notes_style else {}
    body = note_pat.sub('', raw) if is_notes_style else raw
    body = re.sub(r'(?im)^\s*(CATATAN KAKI|FOOTNOTES?)\s*$', '', body)
    body = re.sub(r'(?m)^\s*---+\s*$', '', body)

    # Hindari sampul ganda bila AI ikut menulis blok sampul.
    lines = body.splitlines()
    first_bab = next((i for i,x in enumerate(lines) if re.match(r'^\s*#{0,6}\s*BAB\s+[IVXLCDM]+\b', x, re.I)), None)
    first_kata = next((i for i,x in enumerate(lines) if re.match(r'^\s*#{0,6}\s*KATA\s+PENGANTAR\b', x, re.I)), None)
    if first_kata is not None:
        lines = lines[first_kata:]
    elif first_bab is not None:
        lines = lines[first_bab:]

    bab_seen = 0
    for raw_line in lines:
        t = raw_line.strip()
        if not t:
            continue
        t = re.sub(r'^#{1,6}\s*', '', t)
        t = t.replace('**', '').replace('__', '')
        t = re.sub(r'^>\s*', '', t)

        is_bab = bool(re.match(r'^BAB\s+[IVXLCDM]+\b', t, re.I))
        is_front_heading = bool(re.match(r'^(KATA PENGANTAR|DAFTAR ISI|DAFTAR TABEL|DAFTAR PUSTAKA(?: SEMENTARA)?|SISTEMATIKA PENULISAN(?: TESIS)?(?: \(RENCANA\))?)\s*$', t, re.I))
        is_sub = bool(re.match(r'^[A-Z]\.\s+\S', t))
        is_numbered_item = bool(re.match(r'^\d+[\.)]\s+\S', t))
        is_bullet = bool(re.match(r'^[-*•]\s+\S', t))

        if is_bab:
            if bab_seen > 0:
                d.add_page_break()
            bab_seen += 1
            pp = d.add_paragraph(); pp.alignment = WD_ALIGN_PARAGRAPH.CENTER
            pp.paragraph_format.first_line_indent = Cm(0); pp.paragraph_format.line_spacing = 2
            run = pp.add_run(t.upper()); run.bold = True; run.font.name = font_body; run.font.size = Pt(size_body)
            continue

        if is_front_heading:
            # Bagian utama non-BAB dibuat tegas dan tidak dianggap paragraf naratif.
            pp = d.add_paragraph(); pp.alignment = WD_ALIGN_PARAGRAPH.CENTER
            pp.paragraph_format.first_line_indent = Cm(0); pp.paragraph_format.line_spacing = 2
            run = pp.add_run(t.upper()); run.bold = True; run.font.name = font_body; run.font.size = Pt(size_body)
            continue

        if is_sub:
            pp = d.add_paragraph(); pp.alignment = WD_ALIGN_PARAGRAPH.LEFT
            pp.paragraph_format.first_line_indent = Cm(0); pp.paragraph_format.line_spacing = 2
            run = pp.add_run(t); run.bold = True; run.font.name = font_body; run.font.size = Pt(size_body)
            continue

        pp = d.add_paragraph(); pp.paragraph_format.line_spacing = 2
        if is_numbered_item or is_bullet:
            pp.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY; pp.paragraph_format.first_line_indent = Cm(0)
            pp.paragraph_format.left_indent = Cm(0.75); pp.paragraph_format.hanging_indent = Cm(0.5)
        else:
            pp.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY; pp.paragraph_format.first_line_indent = Cm(1.27)
        run = pp.add_run(t); run.font.name = font_body; run.font.size = Pt(size_body)

    d.save(bio)
    data = bio.getvalue()
    if not (is_notes_style and notes):
        return data

    # Marker [^n] -> true Word footnote. Footnote: Times New Roman 10 pt, spasi 1.
    try:
        W='http://schemas.openxmlformats.org/wordprocessingml/2006/main'
        REL='http://schemas.openxmlformats.org/package/2006/relationships'
        CT='http://schemas.openxmlformats.org/package/2006/content-types'
        XML='http://www.w3.org/XML/1998/namespace'
        q=lambda uri,tag:f'{{{uri}}}{tag}'
        ns={'w':W}
        with zipfile.ZipFile(io.BytesIO(data),'r') as zin:
            infos=zin.infolist(); files={i.filename:zin.read(i.filename) for i in infos}
        parser=etree.XMLParser(remove_blank_text=False)
        root=etree.fromstring(files['word/document.xml'],parser)
        used=set()
        for tn in list(root.xpath('.//w:t',namespaces=ns)):
            txt=tn.text or ''
            mm=list(re.finditer(r'\[\^(\d+)\]',txt))
            if not mm: continue
            run=tn.getparent(); parent=run.getparent(); idx=parent.index(run); pos=0; additions=[]
            old_rpr=run.find(q(W,'rPr'))
            for m in mm:
                before=txt[pos:m.start()]
                if before:
                    nr=etree.Element(q(W,'r'))
                    if old_rpr is not None: nr.append(etree.fromstring(etree.tostring(old_rpr)))
                    nt=etree.SubElement(nr,q(W,'t')); nt.text=before
                    if before.startswith(' ') or before.endswith(' '): nt.set(q(XML,'space'),'preserve')
                    additions.append(nr)
                n=int(m.group(1)); used.add(n)
                rr=etree.Element(q(W,'r')); rpr=etree.SubElement(rr,q(W,'rPr'))
                va=etree.SubElement(rpr,q(W,'vertAlign')); va.set(q(W,'val'),'superscript')
                ref=etree.SubElement(rr,q(W,'footnoteReference')); ref.set(q(W,'id'),str(n)); additions.append(rr)
                pos=m.end()
            after=txt[pos:]
            if after:
                nr=etree.Element(q(W,'r'))
                if old_rpr is not None: nr.append(etree.fromstring(etree.tostring(old_rpr)))
                nt=etree.SubElement(nr,q(W,'t')); nt.text=after
                if after.startswith(' ') or after.endswith(' '): nt.set(q(XML,'space'),'preserve')
                additions.append(nr)
            parent.remove(run)
            for off,nr in enumerate(additions): parent.insert(idx+off,nr)

        fnroot=etree.Element(q(W,'footnotes'),nsmap={'w':W})
        for fid,tag in [(-1,'separator'),(0,'continuationSeparator')]:
            fn=etree.SubElement(fnroot,q(W,'footnote')); fn.set(q(W,'id'),str(fid))
            p1=etree.SubElement(fn,q(W,'p')); r1=etree.SubElement(p1,q(W,'r')); etree.SubElement(r1,q(W,tag))
        for n in sorted(used):
            if n not in notes: continue
            fn=etree.SubElement(fnroot,q(W,'footnote')); fn.set(q(W,'id'),str(n))
            p1=etree.SubElement(fn,q(W,'p'))
            ppr=etree.SubElement(p1,q(W,'pPr'))
            spacing=etree.SubElement(ppr,q(W,'spacing')); spacing.set(q(W,'line'),'240'); spacing.set(q(W,'lineRule'),'auto'); spacing.set(q(W,'before'),'0'); spacing.set(q(W,'after'),'0')
            rnum=etree.SubElement(p1,q(W,'r')); rpr=etree.SubElement(rnum,q(W,'rPr'))
            fonts0=etree.SubElement(rpr,q(W,'rFonts')); fonts0.set(q(W,'ascii'),'Times New Roman'); fonts0.set(q(W,'hAnsi'),'Times New Roman')
            sz0=etree.SubElement(rpr,q(W,'sz')); sz0.set(q(W,'val'),'20')
            va=etree.SubElement(rpr,q(W,'vertAlign')); va.set(q(W,'val'),'superscript'); etree.SubElement(rnum,q(W,'footnoteRef'))
            rt=etree.SubElement(p1,q(W,'r')); rp=etree.SubElement(rt,q(W,'rPr'))
            fonts=etree.SubElement(rp,q(W,'rFonts')); fonts.set(q(W,'ascii'),'Times New Roman'); fonts.set(q(W,'hAnsi'),'Times New Roman')
            sz=etree.SubElement(rp,q(W,'sz')); sz.set(q(W,'val'),'20')
            tt=etree.SubElement(rt,q(W,'t')); tt.set(q(XML,'space'),'preserve'); tt.text=' '+notes[n]

        files['word/document.xml']=etree.tostring(root,xml_declaration=True,encoding='UTF-8',standalone='yes')
        files['word/footnotes.xml']=etree.tostring(fnroot,xml_declaration=True,encoding='UTF-8',standalone='yes')
        relp='word/_rels/document.xml.rels'; relroot=etree.fromstring(files[relp],parser)
        reltype='http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes'
        if not any(x.get('Type')==reltype for x in relroot):
            ids={x.get('Id') for x in relroot}; i=1
            while f'rId{i}' in ids: i+=1
            rel=etree.SubElement(relroot,q(REL,'Relationship')); rel.set('Id',f'rId{i}'); rel.set('Type',reltype); rel.set('Target','footnotes.xml')
        files[relp]=etree.tostring(relroot,xml_declaration=True,encoding='UTF-8',standalone='yes')
        ctp='[Content_Types].xml'; ctroot=etree.fromstring(files[ctp],parser)
        if not any(x.get('PartName')=='/word/footnotes.xml' for x in ctroot):
            ov=etree.SubElement(ctroot,q(CT,'Override')); ov.set('PartName','/word/footnotes.xml'); ov.set('ContentType','application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml')
        files[ctp]=etree.tostring(ctroot,xml_declaration=True,encoding='UTF-8',standalone='yes')
        out=io.BytesIO()
        with zipfile.ZipFile(out,'w',zipfile.ZIP_DEFLATED) as zout:
            written=set()
            for info in infos:
                if info.filename in files and info.filename not in written:
                    zout.writestr(info,files[info.filename]); written.add(info.filename)
            for name,b in files.items():
                if name not in written: zout.writestr(name,b)
        return out.getvalue()
    except Exception:
        return data

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
# ============================================================
# WORD HASIL PENYUNTINGAN - PEDOMAN AKADEMIK AKTIF
# ============================================================
def _bersihkan_hasil_sunting_ai(teks):
    """Bersihkan marker teknis tanpa membuang substansi naskah."""
    import html as _html
    t = _html.unescape(str(teks or ""))
    t = t.replace("\r\n", "\n").replace("\r", "\n")
    t = t.replace("\u00a0", " ").replace("\ufeff", "")
    t = t.replace("\u00ad", "").replace("\ufffe", "-").replace("\ufffd", "")

    # Hapus LABEL AI saja. Isi sesudah label tetap dipertahankan.
    t = re.sub(r'(?im)^\s*#{0,6}\s*HASIL\s+SUNTINGAN\s*:?\s*$', '', t)
    t = re.sub(r'(?im)^\s*#{0,6}\s*CATATAN\s+PERUBAHAN\s+PENTING\s*:?\s*$', '', t)
    t = re.sub(r'(?m)^\s*```[^\n]*$', '', t)
    t = re.sub(r'(?m)^\s*---+\s*$', '', t)
    t = re.sub(r'(?m)^\s*#{1,6}\s+', '', t)

    # Bersihkan pembungkus LaTeX/AI tanpa menghapus isi.
    t = t.replace('\\(', '').replace('\\)', '').replace('\\[', '').replace('\\]', '')
    t = re.sub(r'\$([^$\n]+)\$', r'\1', t)
    t = t.replace('$', '')
    replacements = {
        'H_0':'H₀', 'H_{0}':'H₀', 'H_a':'Hₐ', 'H_{a}':'Hₐ', 'H_1':'H₁', 'H_{1}':'H₁',
        'r_{count}':'r_count', 'r{count}':'r_count', 'r_{hitung}':'r_hitung', 'r{hitung}':'r_hitung',
        'r_{table}':'r_table', 'r{table}':'r_table', 'r_{tabel}':'r_tabel', 'r{tabel}':'r_tabel',
        'r_{11}':'r₁₁', 'r{11}':'r₁₁',
    }
    for x,y in replacements.items(): t=t.replace(x,y)
    t = re.sub(r'\\text\s*\{([^{}]*)\}', r'\1', t)
    t = re.sub(r'\\mathrm\s*\{([^{}]*)\}', r'\1', t)
    t = re.sub(r'\\mathbf\s*\{([^{}]*)\}', r'\1', t)
    t = re.sub(r'\{([^{}\n]+)\}', r'\1', t)

    # Metadata pedoman yang terseret ke isi proposal.
    t = re.sub(r'(?im)^\s*[-*•]?\s*Pedoman\s+Penulisan\s+Tesis\s+\d+\s*\|\s*(?=BAB\b).*?\*?\s*$', '', t)

    # Hapus bintang tunggal yatim, tetapi pasangan Markdown tetap untuk formatter Word.
    fixed=[]
    for line in t.split('\n'):
        if line.count('*') == 1:
            line=line.replace('*','')
        fixed.append(line)
    t='\n'.join(fixed)

    t = re.sub(r'(?i)\bDAFTRA\s+ISI\b', 'DAFTAR ISI', t)
    t = re.sub(r'(?i)\bpemehaman\b', 'pemahaman', t)
    t = re.sub(r'[ \t]+\n', '\n', t)
    t = re.sub(r'\n{3,}', '\n\n', t)
    return t.strip()


def _judul_dari_hasil_sunting(teks):
    """Ambil judul proposal/tesis dari blok awal hasil suntingan."""
    t=_bersihkan_hasil_sunting_ai(teks)
    lines=[re.sub(r'[*_#]', '', x).strip() for x in t.splitlines() if x.strip()]
    for i,x in enumerate(lines):
        if x.upper() in ('PROPOSAL TESIS','TESIS','PROPOSAL SKRIPSI','SKRIPSI','DISERTASI'):
            acc=[]
            for y in lines[i+1:i+8]:
                if re.match(r'(?i)^OLEH\s*:?', y) or re.match(r'(?i)^NPM\s*:', y): break
                if y.upper() in ('KATA PENGANTAR','PROGRAM PASCASARJANA'): break
                acc.append(y)
            if acc: return ' '.join(acc)
    return 'Naskah Akademik Hasil Penyuntingan'


def _add_page_field(paragraph, align=None):
    """Tambahkan PAGE field Word asli."""
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    if align is not None: paragraph.alignment = align
    run=paragraph.add_run()
    fldChar1=OxmlElement('w:fldChar'); fldChar1.set(qn('w:fldCharType'),'begin')
    instr=OxmlElement('w:instrText'); instr.set(qn('xml:space'),'preserve'); instr.text=' PAGE '
    fldChar2=OxmlElement('w:fldChar'); fldChar2.set(qn('w:fldCharType'),'end')
    run._r.extend([fldChar1,instr,fldChar2])


def _set_page_number_format(section, fmt='decimal', start=None):
    """Atur format nomor halaman pada section Word."""
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    sectPr=section._sectPr
    old=sectPr.find(qn('w:pgNumType'))
    if old is not None: sectPr.remove(old)
    el=OxmlElement('w:pgNumType'); el.set(qn('w:fmt'),fmt)
    if start is not None: el.set(qn('w:start'),str(start))
    sectPr.append(el)


def _run_markdown_inline(paragraph, text, font_name='Times New Roman', font_size=12):
    """Tulis teks dengan *italic* dan **bold** sebagai format Word, bukan marker mentah."""
    text=str(text or '')
    pat=re.compile(r'(\*\*[^*]+\*\*|\*[^*]+\*)')
    pos=0
    for m in pat.finditer(text):
        if m.start()>pos:
            r=paragraph.add_run(text[pos:m.start()]); r.font.name=font_name; r.font.size=Pt(font_size)
        token=m.group(0)
        if token.startswith('**'):
            val=token[2:-2]; r=paragraph.add_run(val); r.bold=True
        else:
            val=token[1:-1]; r=paragraph.add_run(val); r.italic=True
        r.font.name=font_name; r.font.size=Pt(font_size)
        pos=m.end()
    if pos<len(text):
        r=paragraph.add_run(text[pos:]); r.font.name=font_name; r.font.size=Pt(font_size)


def buat_docx_hasil_sunting_pedoman(teks, jenis_naskah='Proposal', font_name='Times New Roman', font_size=12):
    """DOCX hasil penyuntingan yang rapi mengikuti ketentuan Pedoman Tesis aktif.
    A4; margin 4-4-3-3 cm; TNR 12; spasi ganda; justify; indent/tab bertingkat;
    cover tanpa nomor; bagian awal Romawi kecil; BAB angka Latin (1, 2, 3, ...); halaman pertama BAB
    di tengah bawah dan halaman lanjutan di kanan atas.
    """
    if docx is None: return None
    from docx.enum.section import WD_SECTION
    from docx.enum.text import WD_BREAK
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Inches

    clean=_bersihkan_hasil_sunting_ai(teks)
    lines=clean.splitlines()
    d=docx.Document()

    def setup_section(sec):
        sec.page_width=Cm(21); sec.page_height=Cm(29.7)
        sec.top_margin=Cm(4); sec.left_margin=Cm(4); sec.bottom_margin=Cm(3); sec.right_margin=Cm(3)
        sec.header_distance=Cm(1.5); sec.footer_distance=Cm(1.5)

    setup_section(d.sections[0])
    normal=d.styles['Normal']
    normal.font.name=font_name; normal.font.size=Pt(font_size)
    normal.paragraph_format.line_spacing=2
    normal.paragraph_format.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY
    normal.paragraph_format.first_line_indent=Cm(1.27)
    normal.paragraph_format.space_before=Pt(0); normal.paragraph_format.space_after=Pt(0)

    # Tentukan batas cover/front/body dari isi naskah.
    idx_kata=next((i for i,x in enumerate(lines) if re.sub(r'[*_#]','',x).strip().upper()=='KATA PENGANTAR'), None)
    idx_bab1=next((i for i,x in enumerate(lines) if re.match(r'^\s*\**\s*BAB\s+I\b',x,re.I)), None)
    current_part='cover' if idx_kata is not None else ('front' if idx_bab1 is not None else 'body')
    bab_count=0

    def configure_front(sec):
        setup_section(sec); sec.header.is_linked_to_previous=False; sec.footer.is_linked_to_previous=False
        sec.different_first_page_header_footer=False
        _set_page_number_format(sec,'lowerRoman',1)
        # Romawi kecil di tengah bawah.
        sec.header.paragraphs[0].clear()
        fp=sec.footer.paragraphs[0]; fp.clear(); _add_page_field(fp,WD_ALIGN_PARAGRAPH.CENTER)

    def configure_body(sec, first=False):
        setup_section(sec); sec.header.is_linked_to_previous=False; sec.footer.is_linked_to_previous=False
        sec.different_first_page_header_footer=True
        if first: _set_page_number_format(sec,'decimal',1)
        else: _set_page_number_format(sec,'decimal',None)
        # Halaman pertama BAB: tengah bawah.
        p=sec.first_page_footer.paragraphs[0]; p.clear(); _add_page_field(p,WD_ALIGN_PARAGRAPH.CENTER)
        sec.first_page_header.paragraphs[0].clear()
        # Halaman lanjutan BAB: kanan atas.
        hp=sec.header.paragraphs[0]; hp.clear(); _add_page_field(hp,WD_ALIGN_PARAGRAPH.RIGHT)
        sec.footer.paragraphs[0].clear()

    # Cover section: tanpa nomor.
    d.sections[0].header.is_linked_to_previous=False; d.sections[0].footer.is_linked_to_previous=False
    d.sections[0].header.paragraphs[0].clear(); d.sections[0].footer.paragraphs[0].clear()

    for i,raw in enumerate(lines):
        t=raw.strip()
        if not t: continue
        plain=re.sub(r'[*_#]','',t).strip()
        up=plain.upper()

        # Mulai bagian awal pada Kata Pengantar.
        if idx_kata is not None and i==idx_kata:
            sec=d.add_section(WD_SECTION.NEW_PAGE); configure_front(sec); current_part='front'
        # Setiap BAB menjadi section baru agar posisi nomor halaman tepat.
        if re.match(r'^BAB\s+[IVXLCDM]+\b',up):
            sec=d.add_section(WD_SECTION.NEW_PAGE); bab_count+=1; configure_body(sec,first=(bab_count==1)); current_part='body'

        # Jangan tampilkan garis pemisah/code marker.
        if plain in ('```','---'): continue

        # Heading utama.
        if re.match(r'^BAB\s+[IVXLCDM]+\b',up):
            p=d.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER; p.paragraph_format.first_line_indent=Cm(0)
            p.paragraph_format.line_spacing=2; p.paragraph_format.keep_with_next=True
            r=p.add_run(up); r.bold=True; r.font.name=font_name; r.font.size=Pt(font_size)
            continue
        if up in ('PENDAHULUAN','KAJIAN PUSTAKA DAN KERANGKA PIKIR','METODE PENELITIAN','KATA PENGANTAR','DAFTAR ISI','DAFTAR TABEL','DAFTAR GAMBAR','DAFTAR LAMPIRAN','SISTEMATIKA PENULISAN','DAFTAR PUSTAKA','DAFTAR PUSTAKA SEMENTARA'):
            p=d.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER; p.paragraph_format.first_line_indent=Cm(0)
            p.paragraph_format.line_spacing=2; p.paragraph_format.keep_with_next=True
            r=p.add_run(plain.upper()); r.bold=True; r.font.name=font_name; r.font.size=Pt(font_size)
            continue

        # Cover dibuat rata tengah sampai Kata Pengantar.
        if current_part=='cover':
            p=d.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER; p.paragraph_format.first_line_indent=Cm(0); p.paragraph_format.line_spacing=2
            _run_markdown_inline(p,t,font_name,font_size)
            if up in ('PROPOSAL TESIS','TESIS') or (len(plain)>35 and plain==plain.upper()):
                for r in p.runs: r.bold=True
            continue

        # Subjudul A., B., C. memakai tab/hanging indent, bukan spasi manual.
        mA=re.match(r'^([A-Z])\.\s*(.+)$',plain)
        if mA:
            p=d.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.LEFT; p.paragraph_format.first_line_indent=Cm(0)
            p.paragraph_format.left_indent=Cm(0); p.paragraph_format.line_spacing=2; p.paragraph_format.keep_with_next=True
            p.paragraph_format.tab_stops.add_tab_stop(Cm(1.0))
            r=p.add_run(mA.group(1)+'.\t'+mA.group(2)); r.bold=True; r.font.name=font_name; r.font.size=Pt(font_size)
            continue

        # Nomor 1., 2. dst. dengan hanging indent dan tab.
        mn=re.match(r'^(\d+)[\.)]\s*(.+)$',plain)
        if mn:
            p=d.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY; p.paragraph_format.first_line_indent=Cm(-0.7)
            p.paragraph_format.left_indent=Cm(0.7); p.paragraph_format.line_spacing=2
            p.paragraph_format.tab_stops.add_tab_stop(Cm(0.7))
            _run_markdown_inline(p,mn.group(1)+'.\t'+mn.group(2),font_name,font_size)
            continue

        # Terjemahan ayat/hadis: tanpa label "Artinya", menjorok, spasi 1.
        # Ayat diakhiri (Q.S. ...), tanpa footnote baru. Hadis diakhiri (HR. ...) lalu nomor footnote.
        _kutip=t.strip()
        _kutip=re.sub(r'^[-*•]\s*', '', _kutip).strip()
        _kutip=re.sub(r'^(?:\*{0,2})?Artinya(?:\*{0,2})?\s*:\s*', '', _kutip, flags=re.I).strip()
        _is_terjemah=bool(re.match(r'^["“]', _kutip) and (re.search(r'\(Q\.S\.\s*[^)]*\)\s*[¹²³⁴⁵⁶⁷⁸⁹⁰\d]*[.!?]?["”]?\s*$', _kutip, re.I) or re.search(r'\(HR\.\s*[^)]*\)\s*[¹²³⁴⁵⁶⁷⁸⁹⁰\d]+[.!?]?["”]?\s*$', _kutip, re.I)))
        if _is_terjemah:
            p=d.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY
            p.paragraph_format.left_indent=Cm(1.27); p.paragraph_format.right_indent=Cm(1.27)
            p.paragraph_format.first_line_indent=Cm(0); p.paragraph_format.line_spacing=1
            p.paragraph_format.space_before=Pt(0); p.paragraph_format.space_after=Pt(0)
            _run_markdown_inline(p,_kutip,font_name,font_size)
            continue

        # Bullet menjadi daftar menjorok rapi.
        mb=re.match(r'^[-*•]\s*(.+)$',t)
        if mb:
            p=d.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY; p.paragraph_format.left_indent=Cm(0.7); p.paragraph_format.first_line_indent=Cm(-0.4); p.paragraph_format.line_spacing=2
            _run_markdown_inline(p,'•\t'+mb.group(1),font_name,font_size)
            continue

        # Ayat/hadis Arab sesuai Pedoman Tesis S2: Traditional Arabic 16 pt, RTL, rata kanan.
        if re.search(r'[\u0600-\u06FF]',plain):
            p=d.add_paragraph()
            p.alignment=WD_ALIGN_PARAGRAPH.RIGHT
            p.paragraph_format.first_line_indent=Cm(0)
            p.paragraph_format.left_indent=Cm(0)
            p.paragraph_format.right_indent=Cm(0)
            p.paragraph_format.line_spacing=1.5
            # Aktifkan BiDi/RTL pada paragraf Word agar huruf Arab tidak terbalik/pecah.
            pPr=p._p.get_or_add_pPr()
            bidi=OxmlElement('w:bidi')
            bidi.set(qn('w:val'),'1')
            pPr.append(bidi)
            r=p.add_run(plain)
            r.font.name='Traditional Arabic'
            r.font.size=Pt(16)
            rPr=r._r.get_or_add_rPr()
            rFonts=rPr.find(qn('w:rFonts'))
            if rFonts is None:
                rFonts=OxmlElement('w:rFonts'); rPr.insert(0,rFonts)
            for attr in ('ascii','hAnsi','eastAsia','cs'):
                rFonts.set(qn('w:'+attr),'Traditional Arabic')
            rtl=OxmlElement('w:rtl'); rtl.set(qn('w:val'),'1'); rPr.append(rtl)
            cs=OxmlElement('w:cs'); cs.set(qn('w:val'),'1'); rPr.append(cs)
            szCs=OxmlElement('w:szCs'); szCs.set(qn('w:val'),'32'); rPr.append(szCs)
            continue

        # Paragraf naratif.
        p=d.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY; p.paragraph_format.first_line_indent=Cm(1.27); p.paragraph_format.line_spacing=2
        _run_markdown_inline(p,t,font_name,font_size)

    bio=BytesIO(); d.save(bio); return bio.getvalue()


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
    st.caption("Bengkel naskah akademik untuk proposal, tesis, skripsi, disertasi, dan artikel. Substansi penelitian dilindungi.")

    # Pedoman aktif global: memakai sumber pedoman yang SUDAH ADA di aplikasi.
    # Status tetap ditampilkan di Penyunting Akademik AI dan tidak bergantung hanya
    # pada satu flag, agar pedoman yang sudah tersimpan tidak tampak hilang.
    _ped_nama = str(st.session_state.get("pedoman_tesis_s2_nama", "") or "")
    _ped_teks = str(st.session_state.get("pedoman_tesis_s2_teks", "") or "")
    _ped_analisis = str(st.session_state.get("pedoman_tesis_s2_analisis", "") or "")
    _ped_flag = bool(st.session_state.get("pedoman_tesis_s2_aktif"))
    _ped_aktif = bool(_ped_flag or _ped_teks.strip() or _ped_analisis.strip())

    st.markdown("### 📘 Pedoman Aktif")
    if _ped_aktif:
        # Pulihkan flag bila isi pedoman sebenarnya sudah tersedia.
        st.session_state["pedoman_tesis_s2_aktif"] = True
        st.success(f"Pedoman Aktif: {_ped_nama or 'Pedoman Penulisan Tesis yang telah diaktifkan'} ✓")
        st.caption("Pedoman ini menjadi acuan global Penyunting Akademik AI. Jika suatu aturan tidak ditemukan di pedoman, AI tidak boleh menebaknya.")
    else:
        st.warning("📘 Belum ada Pedoman Aktif. Penyunting tetap dapat digunakan, tetapi aturan institusi yang tidak tersedia tidak akan ditebak oleh AI.")

    jenis_naskah_editor = st.selectbox(
        "Jenis Naskah",
        ["Deteksi Otomatis", "Proposal", "Tesis", "Skripsi", "Disertasi", "Artikel Jurnal"],
        key="jenis_naskah_penyunting_global"
    )

    submenu_editor = st.radio(
        "Bagian Penyunting Akademik AI",
        [
            "✍️ Sunting & Parafrase",
            "🤖 Asisten Penulisan BAB",
            "👨‍🏫 Revisi Dosen/Penguji",
            "🔎 Audit Akademik",
            "📑 Finalisasi Word",
        ],
        horizontal=True,
        key="submenu_penyunting_akademik_ai"
    )

    # --------------------------------------------------------
    # 1. SUNTING & PARAFRASE
    # --------------------------------------------------------
    if submenu_editor == "✍️ Sunting & Parafrase":
        st.subheader("✍️ Sunting & Parafrase")
        mode_edit = st.selectbox(
            "Mode Penyuntingan",
            [
                "Koreksi Ejaan & Typo",
                "Rapikan Kalimat",
                "Bahasa Akademik",
                "Perkuat Paragraf",
                "Koherensi Antarparagraf",
                "Parafrasa Ringan",
                "Parafrasa Akademik",
                "Parafrasa Mendalam",
                "✨ Penyuntingan Akademik Lengkap (Mode Aman)",
                "Sunting Naskah Lengkap",
            ],
            key="mode_sunting_parafrase_ai"
        )
        file_edit = st.file_uploader("Unggah naskah", type=["pdf", "docx", "txt"], key="file_sunting_parafrase_ai")
        teks_edit = st.text_area("Atau tempel teks", height=300, key="teks_sunting_parafrase_ai")
        if file_edit:
            _t = ekstrak_teks(file_edit)
            if not _t.startswith("ERROR:"):
                teks_edit = _t

        st.info("🔒 Judul, fakta, data, angka, variabel, hasil penelitian, kutipan, sumber, ayat/hadis, tabel, dan makna asli tidak boleh diubah tanpa perintah pengguna.")

        # Naskah aktif memungkinkan penyuntingan berulang tanpa download-upload.
        _naskah_aktif_editor = str(st.session_state.get("naskah_aktif_penyunting_ai", "") or "").strip()
        if _naskah_aktif_editor:
            st.success("✅ Naskah Aktif Terbaru tersedia. Penyuntingan berikutnya otomatis memakai versi aktif ini.")
            _sumber_sunting = _naskah_aktif_editor
        else:
            _sumber_sunting = str(teks_edit or "").strip()

        _siap_sunting = bool(_sumber_sunting)
        if not _siap_sunting:
            st.caption("Unggah naskah atau tempel teks terlebih dahulu agar tombol Sunting aktif.")

        if st.button("✨ Sunting dengan AI", type="primary", disabled=not _siap_sunting, key="btn_sunting_parafrase_ai"):
            # Simpan naskah pertama sebagai versi asli. Tidak pernah ditimpa otomatis.
            if not st.session_state.get("naskah_asli_penyunting_ai"):
                st.session_state["naskah_asli_penyunting_ai"] = _sumber_sunting
            if not st.session_state.get("riwayat_naskah_penyunting_ai"):
                st.session_state["riwayat_naskah_penyunting_ai"] = [
                    {"versi": "Versi Asli", "teks": _sumber_sunting}
                ]
            _ped_prompt = (_ped_teks[:45000] + "\n\nRINGKASAN PEDOMAN:\n" + _ped_analisis[:12000]) if _ped_aktif else "Pedoman institusi belum aktif. Jangan menebak aturan institusi."
            _prompt = f"""Anda adalah Penyunting Akademik AI.
JENIS NASKAH: {jenis_naskah_editor}
MODE: {mode_edit}

PEDOMAN AKTIF:
{_ped_prompt}

ATURAN WAJIB:
1. Pertahankan makna, fakta, angka, data, judul, variabel, hasil, tabel, kutipan, sitasi, nama sumber, ayat dan hadis.
2. Jangan menciptakan referensi, DOI, halaman, data, kutipan langsung, hasil penelitian, atau fakta baru.
3. Perbaiki hanya sesuai mode yang dipilih.
4. Parafrase bertujuan memperjelas bahasa akademik, bukan mengelabui pemeriksa plagiarisme.
5. Sitasi yang sudah ada harus tetap melekat pada klaim yang sama.
6. Bila ada bagian meragukan, tandai [PERLU VERIFIKASI], jangan menebak.
7. Ikuti pedoman aktif bila tersedia. Jika pedoman tidak mengatur sesuatu, jangan membuat aturan institusi sendiri.
8. Keluarkan dua bagian: HASIL SUNTINGAN dan CATATAN PERUBAHAN PENTING. Jangan menambah pembahasan di luar naskah.
9. Jika MODE adalah "✨ Penyuntingan Akademik Lengkap (Mode Aman)", lakukan sekaligus: koreksi ejaan/typo, rapikan kalimat, bahasa akademik, perkuat paragraf, dan koherensi antarparagraf. DILARANG melakukan parafrasa mendalam atau mengubah substansi.
10. Dalam Mode Aman, rumusan masalah, tujuan, hipotesis, nama variabel, istilah metodologis, angka/data, kutipan langsung, sitasi/footnote, daftar pustaka, ayat/hadis, nama tokoh/lembaga, dan temuan penelitian harus dipertahankan.
11. Jangan menghapus bagian naskah hanya karena dianggap berulang. Perbaiki bahasanya tanpa mengurangi informasi substantif.

NASKAH:
{_sumber_sunting[:90000]}"""
            with st.spinner("AI sedang menyunting naskah. Mohon tunggu sampai hasil tampil..."):
                _h = panggil_gemini(_prompt, temperature=0.20)
            if _h.get("sukses"):
                st.session_state["hasil_sunting_parafrase_ai"] = str(_h.get("hasil", "") or "").strip()
                st.session_state["model_sunting_parafrase_ai"] = str(_h.get("model", "") or "")
                st.success("✅ Penyuntingan selesai. Hasil tampil di bawah.")
            else:
                st.session_state["hasil_sunting_parafrase_ai"] = ""
                st.error("❌ Penyuntingan belum berhasil.")
                st.warning(str(_h.get("error", "Kesalahan AI tidak diketahui.")))

        if st.session_state.get("hasil_sunting_parafrase_ai"):
            _model_sunting = st.session_state.get("model_sunting_parafrase_ai", "")
            if _model_sunting:
                st.caption(f"Model AI: {_model_sunting}")
            _hasil_sunting = st.session_state["hasil_sunting_parafrase_ai"]

            # Pisahkan isi naskah dari catatan AI agar versi aktif berikutnya tetap bersih.
            _hasil_bersih_aktif = str(_hasil_sunting or "").strip()
            _m_hasil = re.search(r"(?is)HASIL\s+SUNTINGAN\s*:?\s*(.*?)(?=\n\s*CATATAN\s+PERUBAHAN\s+PENTING\s*:?|$)", _hasil_bersih_aktif)
            if _m_hasil:
                _hasil_bersih_aktif = _m_hasil.group(1).strip()
            _hasil_bersih_aktif = re.sub(r"^```(?:text|markdown)?\s*|\s*```$", "", _hasil_bersih_aktif, flags=re.I|re.S).strip()

            st.markdown("#### 👁️ Preview Dokumen Sebelum Download")
            st.caption("Preview ditampilkan seperti Print Layout Word: lembar A4, margin, Times New Roman, spasi, heading, paragraf, daftar, dan teks Arab. Periksa hasil di sini sebelum menjadikannya naskah aktif atau mengunduh Word.")
            tampilkan_preview_word_rapi(_hasil_bersih_aktif, tinggi=920, judul="Preview Hasil Penyuntingan")
            with st.expander("🔎 Bandingkan dengan naskah sebelum disunting", expanded=False):
                _naskah_sebelum = str(st.session_state.get("naskah_aktif_penyunting_ai") or st.session_state.get("naskah_asli_penyunting_ai") or teks_edit or "")
                tampilkan_preview_word_rapi(_naskah_sebelum, tinggi=650, judul="Naskah Sebelum Disunting")

            # Terapkan pembaruan hasil referensi SEBELUM widget text_area dibuat.
            # Streamlit melarang perubahan session_state sebuah widget setelah widget
            # dengan key yang sama sudah diinstansiasi pada run yang sama.
            _pending_area = st.session_state.pop("pending_hasil_sunting_parafrase_area", None)
            if _pending_area is not None:
                st.session_state["hasil_sunting_parafrase_area"] = _pending_area
            elif "hasil_sunting_parafrase_area" not in st.session_state:
                st.session_state["hasil_sunting_parafrase_area"] = _hasil_sunting
            with st.expander("✏️ Edit teks manual (opsional)", expanded=False):
                st.caption("Bagian ini hanya jika Anda ingin memperbaiki kata tertentu secara manual. Preview rapi tetap menjadi tampilan utama.")
                st.text_area("Hasil Suntingan", height=500, key="hasil_sunting_parafrase_area")
            # Gunakan isi editor TERKINI untuk Word. Jika pengguna memperbaiki teks di kotak
            # Hasil Suntingan, perubahan itu ikut masuk ke file unduhan.
            _hasil_sunting_untuk_word = st.session_state.get("hasil_sunting_parafrase_area", _hasil_sunting) or _hasil_sunting

            st.markdown("#### ✅ Tetapkan Hasil")
            _a1, _a2 = st.columns(2)
            if _a1.button("✅ Jadikan Naskah Aktif Terbaru", key="btn_jadikan_naskah_aktif_penyunting", type="primary", use_container_width=True):
                _teks_aktif_baru = str(st.session_state.get("hasil_sunting_parafrase_area", _hasil_bersih_aktif) or _hasil_bersih_aktif).strip()
                _m_aktif = re.search(r"(?is)HASIL\s+SUNTINGAN\s*:?\s*(.*?)(?=\n\s*CATATAN\s+PERUBAHAN\s+PENTING\s*:?|$)", _teks_aktif_baru)
                if _m_aktif:
                    _teks_aktif_baru = _m_aktif.group(1).strip()
                if _teks_aktif_baru:
                    _riwayat = list(st.session_state.get("riwayat_naskah_penyunting_ai", []))
                    _nomor = 1 + sum(1 for _v in _riwayat if str(_v.get("versi", "")).startswith("Suntingan"))
                    _riwayat.append({"versi": f"Suntingan {_nomor}", "teks": _teks_aktif_baru})
                    st.session_state["riwayat_naskah_penyunting_ai"] = _riwayat
                    st.session_state["naskah_aktif_penyunting_ai"] = _teks_aktif_baru
                    st.session_state["hasil_sunting_parafrase_ai"] = ""
                    st.session_state.pop("hasil_sunting_parafrase_area", None)
                    st.success(f"✅ Suntingan {_nomor} sekarang menjadi Naskah Aktif Terbaru. Tidak perlu unggah ulang.")
                    st.rerun()
            if _a2.button("↩️ Batalkan Hasil Ini", key="btn_batalkan_hasil_penyunting", use_container_width=True):
                st.session_state["hasil_sunting_parafrase_ai"] = ""
                st.session_state.pop("hasil_sunting_parafrase_area", None)
                st.info("Hasil suntingan dibatalkan. Naskah aktif sebelumnya tetap aman.")
                st.rerun()

            _riwayat_now = st.session_state.get("riwayat_naskah_penyunting_ai", [])
            if _riwayat_now:
                with st.expander("📜 Riwayat Versi Naskah", expanded=False):
                    _nama_versi = [str(_v.get("versi", f"Versi {i+1}")) for i, _v in enumerate(_riwayat_now)]
                    _pilih_versi = st.selectbox("Lihat versi", _nama_versi, index=len(_nama_versi)-1, key="pilih_riwayat_naskah_penyunting")
                    _idx_versi = _nama_versi.index(_pilih_versi)
                    st.text_area("Isi versi", value=str(_riwayat_now[_idx_versi].get("teks", "")), height=300, disabled=True, key="lihat_riwayat_naskah_penyunting")
                    if st.button("♻️ Jadikan Versi Ini Naskah Aktif", key="btn_pulihkan_versi_penyunting"):
                        st.session_state["naskah_aktif_penyunting_ai"] = str(_riwayat_now[_idx_versi].get("teks", ""))
                        st.success(f"{_pilih_versi} dipulihkan sebagai Naskah Aktif Terbaru.")
                        st.rerun()

            st.markdown("#### 📚 Perkuat Referensi & Kutipan")
            st.caption("Akademia AI membaca isi naskah untuk menentukan kata kunci dan menyaring referensi yang relevan. Kata kunci tambahan bersifat opsional.")
            _kata_ref = st.text_input(
                "Kata kunci tambahan (opsional)",
                key="kata_ref_penyunting",
                placeholder="Kosongkan untuk pencarian otomatis dari isi proposal/tesis"
            )
            if st.button("🔎 Cari Referensi Sesuai Proposal", key="btn_cari_ref_penyunting", type="primary", use_container_width=True):
                _naskah_ref = str(_hasil_sunting_untuk_word or "").strip()
                if not _naskah_ref:
                    st.warning("Belum ada naskah proposal/tesis yang dapat dianalisis.")
                else:
                    with st.spinner("Membaca isi naskah, menyusun kata kunci, mencari, lalu menyaring relevansi referensi..."):
                        # 1. AI menyusun query dari substansi naskah, bukan dari kata umum seperti 'kutipan'.
                        _prompt_kw = f"""Anda adalah asisten penelusuran literatur akademik.
Baca naskah penelitian berikut dan buat tepat 6 QUERY PENCARIAN yang paling mewakili substansi penelitian.
Fokus pada: judul/topik, variabel atau fokus utama, teori/konsep inti, objek/subjek, konteks pendidikan, dan istilah padanan bahasa Inggris.
Jangan gunakan kata generik seperti referensi, kutipan, daftar pustaka, proposal, tesis, penelitian, atau metodologi kecuali memang merupakan konsep yang diteliti.
Jika ada KATA KUNCI TAMBAHAN, gunakan hanya sebagai penguat bila selaras dengan naskah.
Kembalikan HANYA 6 baris query tanpa nomor dan tanpa penjelasan.

KATA KUNCI TAMBAHAN:
{_kata_ref.strip() or '(tidak ada)'}

NASKAH:
{_naskah_ref[:45000]}"""
                        _hkw = panggil_gemini(_prompt_kw, temperature=0.10)
                        _queries = []
                        if _hkw.get("sukses"):
                            for _baris in str(_hkw.get("hasil", "")).splitlines():
                                _q = re.sub(r"^\s*[-*•\d\.\)\:]+\s*", "", _baris).strip().strip('"“”')
                                if len(_q) >= 8 and _q.lower() not in [x.lower() for x in _queries]:
                                    _queries.append(_q)
                        if _kata_ref.strip() and _kata_ref.strip().lower() not in [x.lower() for x in _queries]:
                            _queries.append(_kata_ref.strip())
                        if not _queries:
                            _queries = [_kata_ref.strip()] if _kata_ref.strip() else [_naskah_ref[:180]]

                        # 2. Cari lebih luas pada tiga indeks, lalu deduplikasi.
                        _refs = []
                        for _q in _queries[:7]:
                            _refs.extend(cari_crossref(_q, 5))
                            _refs.extend(cari_openalex(_q, 5))
                            _refs.extend(cari_semantic_scholar(_q, 5))
                        _seen, _uniq = set(), []
                        for _r in _refs:
                            _k = kunci_ref(_r)
                            if _k and _k not in _seen and str(_r.get("Judul", "")).strip():
                                _seen.add(_k); _uniq.append(_r)

                        # 3. AI hanya menilai relevansi metadata. Tidak boleh mengarang isi artikel.
                        _candidates = []
                        for _i, _r in enumerate(_uniq[:60], 1):
                            _candidates.append(
                                f"[{_i}] Judul: {_r.get('Judul','')} | Penulis: {_r.get('Penulis','')} | Tahun: {_r.get('Tahun','')} | Jurnal: {_r.get('Jurnal','')} | Sumber: {_r.get('Sumber','')}"
                            )
                        _ranked = []
                        if _candidates:
                            _prompt_rank = f"""Anda adalah penyaring relevansi literatur untuk tesis.
Nilai HANYA kecocokan metadata kandidat dengan substansi naskah. Jangan mengarang abstrak, isi, temuan, atau klaim artikel.
Nilai maksimal 25 kandidat terbaik dan klasifikasikan berdasarkan kecocokannya dengan ISI NASKAH, bukan sekadar kemiripan kata pada judul. Kandidat yang hanya kebetulan memiliki kata yang sama tetapi topiknya berbeda harus dinyatakan tidak disarankan.

Kembalikan HANYA JSON array valid seperti:
[{{"no": 3, "skor": 92, "status": "Sangat Relevan", "keputusan": "Disarankan", "bagian_proposal": "Latar belakang tentang kompetensi guru dan AI", "alasan": "..."}}, {{"no": 7, "skor": 42, "status": "Tidak Relevan", "keputusan": "Tidak Disarankan", "bagian_proposal": "Tidak ada bagian yang cocok", "alasan": "Topik artikel berbeda dari fokus penelitian"}}]

Aturan keputusan:
85-100 = Sangat Relevan, Disarankan
70-84 = Relevan, Disarankan
0-69 = Tidak Relevan, Tidak Disarankan
Wajib isi bagian_proposal dan alasan secara ringkas. Jangan mengarang isi artikel; keputusan hanya berdasarkan metadata yang tersedia dibandingkan dengan naskah.

NASKAH:
{_naskah_ref[:30000]}

KANDIDAT:
{chr(10).join(_candidates)}"""
                            _hrank = panggil_gemini(_prompt_rank, temperature=0.05)
                            if _hrank.get("sukses"):
                                _raw = str(_hrank.get("hasil", "")).strip()
                                _raw = re.sub(r"^```(?:json)?\s*|\s*```$", "", _raw, flags=re.I|re.S).strip()
                                try:
                                    _nilai = json.loads(_raw)
                                except Exception:
                                    _nilai = []
                                if isinstance(_nilai, list):
                                    for _v in _nilai:
                                        try:
                                            _idx = int(_v.get("no", 0)) - 1
                                            _skor = int(float(_v.get("skor", 0)))
                                        except Exception:
                                            continue
                                        if 0 <= _idx < len(_uniq):
                                            _rr = dict(_uniq[_idx])
                                            _rr["Skor Relevansi"] = _skor
                                            _rr["Relevansi"] = str(_v.get("status", "Relevan" if _skor >= 70 else "Tidak Relevan"))
                                            _keputusan = str(_v.get("keputusan", "Disarankan" if _skor >= 70 else "Tidak Disarankan")).strip()
                                            if _skor >= 70:
                                                _keputusan = "Disarankan"
                                            else:
                                                _keputusan = "Tidak Disarankan"
                                            _rr["Keputusan AI"] = _keputusan
                                            _rr["Bagian Proposal"] = str(_v.get("bagian_proposal", "")).strip()
                                            _rr["Alasan Relevansi"] = str(_v.get("alasan", "")).strip()
                                            _ranked.append(_rr)
                        _ranked.sort(key=lambda x: int(x.get("Skor Relevansi", 0)), reverse=True)
                        st.session_state["refs_penyunting_terverifikasi"] = _ranked[:25]
                        st.session_state["query_ref_penyunting_otomatis"] = _queries[:7]
                        st.session_state["tampilan_review_ref"] = "Disarankan"

                    if st.session_state.get("refs_penyunting_terverifikasi"):
                        _semua_ref = st.session_state["refs_penyunting_terverifikasi"]
                        _jml_saran = sum(1 for _r in _semua_ref if _r.get("Keputusan AI") == "Disarankan")
                        _jml_tolak = sum(1 for _r in _semua_ref if _r.get("Keputusan AI") == "Tidak Disarankan")
                        st.success(f"Review selesai: {_jml_saran} referensi disarankan dan {_jml_tolak} tidak disarankan berdasarkan kecocokan dengan isi naskah.")
                        with st.expander("🔎 Kata kunci otomatis yang digunakan"):
                            for _q in st.session_state.get("query_ref_penyunting_otomatis", []):
                                st.write(f"• {_q}")
                    else:
                        st.warning("Belum ditemukan referensi yang cukup relevan dengan isi naskah. Coba tambahkan kata kunci khusus secara opsional.")

            _refs_now = st.session_state.get("refs_penyunting_terverifikasi", [])
            if "keputusan_manual_ref_penyunting" not in st.session_state:
                st.session_state["keputusan_manual_ref_penyunting"] = {}

            _selected = []
            if _refs_now:
                st.markdown("##### 🧭 Review Kesesuaian Referensi dengan Isi Proposal/Tesis")
                st.caption("Tinjau setiap referensi. Klik ✅ Disarankan jika boleh digunakan pada proposal, atau ❌ Tidak Disarankan jika tidak boleh digunakan. Keputusan Anda menjadi kontrol akhir.")

                _keputusan_manual = st.session_state["keputusan_manual_ref_penyunting"]
                for _i, _r in enumerate(_refs_now, 1):
                    _sk = _r.get("Skor Relevansi", "")
                    _judul = _r.get("Judul", "")
                    _penulis = _r.get("Penulis", "")
                    _tahun = _r.get("Tahun", "")
                    _sumber = _r.get("Sumber", "")
                    _bagian = _r.get("Bagian Proposal", "") or "Belum ditentukan"
                    _alasan = _r.get("Alasan Relevansi", "") or "Tidak ada alasan tambahan."
                    _ai = _r.get("Keputusan AI", "Belum dinilai")
                    _rid = str(_r.get("DOI", "") or _r.get("URL", "") or f"{_judul}|{_penulis}|{_tahun}|{_sumber}")
                    _status_manual = _keputusan_manual.get(_rid, "Belum dipilih")

                    with st.container(border=True):
                        st.markdown(f"**{_i}. [{_sk}%] {_judul}**")
                        st.caption(f"{_penulis} ({_tahun}) • {_sumber}")
                        st.write(f"**Cocok untuk bagian:** {_bagian}")
                        st.write(f"**Alasan:** {_alasan}")
                        st.caption(f"Penilaian AI: {_ai} | Keputusan Anda: {_status_manual}")

                        _b1, _b2 = st.columns(2)
                        if _b1.button("✅ Disarankan", key=f"ref_ok_{_i}_{abs(hash(_rid))}", use_container_width=True, type="primary" if _status_manual == "Disarankan" else "secondary"):
                            st.session_state["keputusan_manual_ref_penyunting"][_rid] = "Disarankan"
                            st.rerun()
                        if _b2.button("❌ Tidak Disarankan", key=f"ref_no_{_i}_{abs(hash(_rid))}", use_container_width=True, type="primary" if _status_manual == "Tidak Disarankan" else "secondary"):
                            st.session_state["keputusan_manual_ref_penyunting"][_rid] = "Tidak Disarankan"
                            st.rerun()

                    if _status_manual == "Disarankan":
                        _selected.append(_r)

                _jml_pilih = len(_selected)
                _jml_tolak = sum(1 for _v in _keputusan_manual.values() if _v == "Tidak Disarankan")
                st.info(f"Referensi yang Anda setujui: {_jml_pilih} | Tidak disarankan: {_jml_tolak}")
                if _selected:
                    with st.expander(f"📚 Lihat {_jml_pilih} referensi yang akan diterapkan", expanded=False):
                        for _r in _selected:
                            st.write(f"• {_r.get('Penulis','')} ({_r.get('Tahun','')}). {_r.get('Judul','')}")
            else:
                _selected = []

            st.markdown("##### 📌 Tindakan Setelah Review")
            st.info("🔒 Mode aman aktif: kutipan dan footnote lama tidak boleh dihapus. Footnote baru harus menyambung nomor terakhir tanpa penomoran ganda; daftar pustaka lama tetap dipertahankan.")
            _aksi_ref_1, _aksi_ref_2 = st.columns(2)
            if _aksi_ref_1.button("📚 Terapkan Referensi yang Disarankan ke Proposal", key="btn_perkuat_ref_kutipan", type="primary", use_container_width=True):
                if not bool(str(_hasil_sunting_untuk_word).strip()):
                    st.warning("Belum ada naskah hasil suntingan yang dapat diperkuat.")
                elif not _selected:
                    st.warning("Pilih minimal satu referensi dari kategori Disarankan terlebih dahulu.")
                else:
                    _meta = "\n".join([f"- Penulis: {r.get('Penulis','')}; Tahun: {r.get('Tahun','')}; Judul: {r.get('Judul','')}; Jurnal/Penerbit: {r.get('Jurnal','')}; DOI: {r.get('DOI','')}; Halaman metadata: {r.get('Halaman','')}; Status: {r.get('Status','')}" for r in _selected])
                    _prompt_ref = f"""Anda adalah penyunting referensi akademik.
    JENIS NASKAH: {jenis_naskah_editor}

    NASKAH:
    {str(_hasil_sunting_untuk_word)[:90000]}

    REFERENSI METADATA YANG DIIZINKAN:
    {_meta}

    TUGAS WAJIB - MODE AMAN PENYUNTING AKADEMIK AI:
    1. NASKAH ASLI ADALAH MASTER. Jangan menulis ulang, meringkas, menghapus, memindahkan, atau mengganti paragraf yang tidak perlu. Pertahankan urutan BAB, subbab, tabel, ayat, hadis, data, angka, istilah, dan substansi.
    2. PERTAHANKAN 100% SEMUA SUMBER DAN KUTIPAN/FOOTNOTE ASLI. Dilarang menghapus nomor catatan kaki, teks catatan kaki, DOI, sumber kitab, sumber tafsir, sumber hadis, atau penanda [PERLU VERIFIKASI] yang sudah ada.
    3. DETEKSI DAN KUNCI GAYA SITASI NASKAH ASLI. Proposal/naskah unggahan adalah master. Jika gaya asli Chicago footnote, semua sitasi tambahan atau sitasi yang terlanjur berbentuk APA/author-date harus dirapikan menjadi Chicago footnote dengan sumber yang sama. Jika gaya asli APA, tetap APA. Jika gaya lain, ikuti gaya asli. DILARANG mengubah naskah ke gaya pilihan AI.
    4. Jika gaya asli menggunakan footnote, identifikasi seluruh nomor footnote yang SUDAH ADA. Pertahankan catatan lama, rapikan formatnya, dan sisipkan footnote baru sesuai posisi kutipan sehingga penomoran akhir unik, berurutan, dan tidak ganda. Jangan membuat rangkaian nomor kedua yang dimulai lagi dari 1.
    4a. AYAT AL-QURAN: hapus label/bullet 'Artinya:' pada terjemahan. Tulis terjemahan langsung di dalam tanda kutip, sebagai kutipan menjorok dan spasi 1, lalu akhiri dengan identitas ayat seperti (Q.S. Al-Hasyr/59: 18). Identitas ayat tersebut TIDAK diberi footnote baru. Tafsir/penjelasan setelah ayat kembali menjadi paragraf biasa dan menggunakan sitasi sesuai gaya asli naskah.
    4b. HADIS: hapus label/bullet 'Artinya:' pada terjemahan. Tulis terjemahan langsung di dalam tanda kutip, sebagai kutipan menjorok dan spasi 1. Pada akhir terjemahan tulis (HR. Nama Perawi) lalu nomor footnote sesuai urutan gaya asli. Footnote hadis memuat sumber/takhrij yang benar-benar tersedia. Jangan mengarang nomor hadis, halaman, sanad, atau data yang belum terverifikasi.
    5. Rapikan footnote lama dan baru secara konsisten tanpa mengubah identitas sumber. Jika metadata kurang, pertahankan sumber dan tandai [PERLU VERIFIKASI] atau [halaman perlu verifikasi], jangan mengarang.
    6. Tambahkan sumber baru HANYA pada klaim yang benar-benar relevan dengan metadata referensi yang DIIZINKAN. Jangan memaksakan semua referensi masuk ke naskah.
    7. Jangan mengarang isi artikel, DOI, volume, nomor, halaman, kutipan langsung, hasil penelitian, nomor hadis, sanad, atau halaman kitab/tafsir. Jika dukungan substantif belum dapat dipastikan, tandai [PERLU VERIFIKASI SUMBER].
    8. Untuk tafsir, pertahankan footnote tafsir yang sudah ada. Tambahan baru hanya jika benar-benar diperlukan; halaman yang tidak diketahui ditulis [halaman perlu verifikasi]. Untuk hadis, pertahankan sumber/takhrij lama dan jangan membuat nomor hadis atau sanad.
    9. Sinkronkan DAFTAR PUSTAKA secara ADITIF: jangan menghapus entri lama. Tambahkan hanya sumber baru yang benar-benar digunakan. Jangan membuat entri ganda; jika sumber sudah ada, rapikan entri yang sama, bukan menambah duplikat.
    10. HASIL WAJIB berupa naskah utuh dengan isi asli tetap lengkap. Jangan menambahkan judul buatan seperti 'NASKAH DIPERKUAT'. Jangan menambahkan laporan analisis ke dalam badan naskah.
    11. Lakukan pemeriksaan akhir sebelum mengeluarkan hasil: jumlah kutipan lama tidak boleh berkurang; nomor footnote harus unik dan berurutan; footnote baru harus menyambung; tidak boleh ada dua nomor sama untuk catatan berbeda.
    """
                    with st.spinner("Memperkuat referensi dan memeriksa kutipan..."):
                        _hr = panggil_gemini(_prompt_ref, temperature=0.15)
                    if _hr.get("sukses"):
                        st.session_state["hasil_sunting_parafrase_ai"] = str(_hr.get("hasil", "")).strip()
                        # Jangan menulis langsung ke key widget yang sudah dibuat pada run ini.
                        # Simpan sebagai pending, lalu terapkan pada awal rerun berikutnya.
                        st.session_state["pending_hasil_sunting_parafrase_area"] = st.session_state["hasil_sunting_parafrase_ai"]
                        st.success("Referensi terpilih selesai diterapkan. Kutipan/footnote lama wajib dipertahankan, footnote baru disambung tanpa nomor ganda, dan daftar pustaka disinkronkan secara aditif.")
                        st.rerun()
                    else:
                        st.error(str(_hr.get("error", "Pemeriksaan referensi gagal.")))

            if _aksi_ref_2.button("🗑️ Hapus Hasil Review Referensi", key="btn_hapus_penyunting", use_container_width=True):
                for _k in ["hasil_sunting_parafrase_ai", "hasil_sunting_parafrase_area", "model_sunting_parafrase_ai", "refs_penyunting_terverifikasi", "pilih_ref_penyunting", "kata_ref_penyunting", "pending_hasil_sunting_parafrase_area", "tampilan_review_ref", "query_ref_penyunting_otomatis"]:
                    st.session_state.pop(_k, None)
                st.success("Naskah hasil dan referensi sementara dibersihkan. Pedoman dan bank referensi tetap aman.")
                st.rerun()

            st.markdown("#### 📥 Download Hasil Terakhir")
            st.caption("Word dibersihkan dari marker AI/Markdown dan ditata mengikuti Pedoman Tesis aktif: A4, margin 4-4-3-3 cm, Times New Roman 12, spasi ganda, justify, indent/tab bertingkat, serta nomor halaman sesuai bagian naskah.")
            try:
                _word_sunting = buat_docx_hasil_sunting_pedoman(
                    _hasil_sunting_untuk_word, jenis_naskah=jenis_naskah_editor, font_name="Times New Roman", font_size=12
                )
                if _word_sunting:
                    st.download_button(
                        "📥 Download Word Hasil Suntingan Rapi", data=_word_sunting,
                        file_name="Hasil_Suntingan_Akademik_Rapi.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        type="primary", key="download_word_hasil_sunting_rapi"
                    )
                    st.success("✅ Word hasil terakhir siap diunduh. Cover tanpa nomor; bagian awal memakai Romawi kecil; BAB memakai angka Latin (1, 2, 3, ...), halaman pertama BAB di tengah bawah dan halaman berikutnya di kanan atas.")
            except Exception as _e_word_sunting:
                st.error(f"Word belum dapat dibuat: {_e_word_sunting}")

    # --------------------------------------------------------
    # 2. ASISTEN PENULISAN BAB
    # --------------------------------------------------------
    elif submenu_editor == "🤖 Asisten Penulisan BAB":
        st.subheader("🤖 Asisten Penulisan BAB")
        st.caption("Membantu mengembangkan uraian akademik dari proposal hingga tesis BAB I–V. Data, angka, hasil penelitian, kutipan, dan sumber asli dilindungi.")
        _tahap = st.selectbox("Tahap Naskah", ["Proposal sebelum Sempro", "Tesis setelah Sempro / penelitian", "Naskah tesis lengkap"], key="tahap_asisten_bab")
        _bab = st.selectbox("Pilih BAB", ["BAB I — Pendahuluan", "BAB II — Kajian Pustaka / Landasan Teori", "BAB III — Metode Penelitian", "BAB IV — Hasil Penelitian dan Pembahasan", "BAB V — Penutup"], key="bab_asisten_penulisan")
        _subbab = st.text_input("Subbab / bagian yang sedang ditulis", key="subbab_asisten_penulisan", placeholder="Contoh: Latar Belakang, Kajian Teori, Pembahasan Hasil Uji Hipotesis")
        _mode_bab = st.selectbox("Bantuan yang Dibutuhkan", ["Kembangkan paragraf/deskripsi", "Buat paragraf akademik dari poin-poin", "Perkuat hubungan dengan teori", "Buat transisi antarparagraf", "Jelaskan tabel/hasil penelitian tanpa mengubah angka", "Susun pembahasan: temuan → teori → penelitian terdahulu", "Rapikan argumentasi dan alur", "Bantu simpulan/implikasi tanpa membuat temuan baru"], key="mode_asisten_bab")
        _bahan_bab = st.text_area("Bahan / paragraf / poin / hasil yang akan diolah", height=320, key="bahan_asisten_bab", placeholder="Tempel paragraf kasar, poin-poin, tabel yang sudah diekstrak, atau hasil analisis Anda di sini.")
        _instruksi_bab = st.text_area("Arahan tambahan (opsional)", height=100, key="instruksi_asisten_bab")
        st.info("🔒 Asisten tidak boleh menciptakan data, hasil statistik, responden, kutipan, DOI, halaman, teori, atau sumber yang tidak diberikan/terverifikasi.")
        if st.button("🤖 Bantu Tulis BAB", type="primary", key="btn_asisten_penulisan_bab", disabled=not bool(str(_bahan_bab).strip())):
            _ped_prompt = (_ped_teks[:45000] + "\n\nRINGKASAN PEDOMAN:\n" + _ped_analisis[:12000]) if _ped_aktif else "Pedoman institusi belum aktif. Jangan menebak aturan institusi."
            _bank = st.session_state.get("bank_referensi", [])
            _bank_txt = "\n".join([str(x) for x in _bank[-30:]]) if _bank else "Belum ada bank referensi aktif."
            _prompt_bab = f"""Anda adalah Asisten Penulisan BAB Akademia AI.
TAHAP: {_tahap}
JENIS NASKAH: {jenis_naskah_editor}
BAB: {_bab}
SUBBAB: {_subbab}
TUGAS: {_mode_bab}
ARAHAN TAMBAHAN: {_instruksi_bab}

PEDOMAN AKTIF:
{_ped_prompt}

BANK REFERENSI YANG TERSEDIA:
{_bank_txt[:20000]}

ATURAN:
1. Kembangkan hanya dari bahan pengguna dan sumber yang tersedia.
2. Jangan mengubah atau menciptakan angka, data, hasil statistik, populasi, sampel, instrumen, temuan, kutipan, atau kesimpulan penelitian.
3. Jangan membuat referensi, DOI, halaman, kutipan langsung, tafsir, atau hadis. Jika dukungan sumber diperlukan tetapi belum tersedia, beri [REFERENSI DIPERLUKAN].
4. Untuk BAB IV, bedakan HASIL (deskriptif objektif) dan PEMBAHASAN (interpretasi dengan teori/penelitian terdahulu). Jangan mengubah angka.
5. Untuk tafsir, uraian penafsiran wajib memiliki sumber; jika halaman tidak tersedia tandai [halaman perlu verifikasi].
6. Ikuti pedoman aktif. Jika aturan tidak ditemukan, jangan mengarang aturan institusi.
7. Hasil harus siap disalin ke naskah akademik, koheren, tidak bertele-tele, dan mempertahankan maksud penulis.

BAHAN PENGGUNA:
{str(_bahan_bab)[:90000]}"""
            with st.spinner("Asisten sedang mengolah bagian BAB..."):
                _hb = panggil_gemini(_prompt_bab, temperature=0.20)
            if _hb.get("sukses"):
                st.session_state["hasil_asisten_penulisan_bab"] = str(_hb.get("hasil", "")).strip()
            else:
                st.error(str(_hb.get("error", "Asisten BAB belum berhasil.")))
        if st.session_state.get("hasil_asisten_penulisan_bab"):
            st.text_area("Hasil Asisten Penulisan BAB", st.session_state["hasil_asisten_penulisan_bab"], height=650, key="hasil_asisten_penulisan_bab_area")
            _bc1, _bc2 = st.columns(2)
            if _bc1.button("➡️ Kirim ke Sunting & Parafrase", key="btn_bab_ke_sunting"):
                st.session_state["hasil_sunting_parafrase_ai"] = st.session_state.get("hasil_asisten_penulisan_bab_area", st.session_state["hasil_asisten_penulisan_bab"])
                st.success("Hasil disiapkan untuk tahap penyuntingan.")
            if _bc2.button("🗑️ Hapus Hasil BAB", key="btn_hapus_hasil_bab"):
                st.session_state.pop("hasil_asisten_penulisan_bab", None)
                st.session_state.pop("hasil_asisten_penulisan_bab_area", None)
                st.rerun()

    # --------------------------------------------------------
    # 3. REVISI DOSEN / PENGUJI
    # --------------------------------------------------------
    elif submenu_editor == "👨‍🏫 Revisi Dosen/Penguji":
        st.subheader("👨‍🏫 Revisi Dosen/Penguji")
        st.caption("AI mencari sendiri BAB, subbab, dan paragraf yang terkait dengan catatan pembimbing. Bagian lain dipertahankan.")
        file_revisi = st.file_uploader("Unggah naskah yang akan direvisi", type=["pdf", "docx", "txt"], key="file_revisi_dosen_ai")
        teks_revisi = st.text_area("Atau tempel naskah", height=250, key="teks_revisi_dosen_ai")
        if file_revisi:
            _t = ekstrak_teks(file_revisi)
            if not _t.startswith("ERROR:"):
                teks_revisi = _t
        catatan_dosen = st.text_area("Catatan / arahan dosen pembimbing", height=180, key="catatan_dosen_pembimbing_ai", placeholder="Contoh: Perkuat research gap pada latar belakang dan jangan mengubah metode penelitian.")
        batas_revisi = st.selectbox("Lokasi Revisi", ["Otomatis oleh AI", "BAB I", "BAB II", "BAB III", "BAB IV", "BAB V", "Bagian/Paragraf yang disebut dalam catatan"], key="batas_revisi_dosen_ai")

        st.info("🔒 AI hanya merevisi bagian yang diminta. Perubahan substantif besar harus ditandai terlebih dahulu, bukan dilakukan diam-diam.")
        if st.button("🔎 Analisis & Kerjakan Revisi", type="primary", disabled=not (bool(str(teks_revisi).strip()) and bool(str(catatan_dosen).strip())), key="btn_revisi_dosen_ai"):
            _ped_prompt = (_ped_teks[:45000] + "\n\nRINGKASAN PEDOMAN:\n" + _ped_analisis[:12000]) if _ped_aktif else "Pedoman institusi belum aktif. Jangan menebak aturan institusi."
            _prompt = f"""Anda adalah asisten revisi dosen pembimbing untuk naskah akademik.
JENIS NASKAH: {jenis_naskah_editor}
BATAS LOKASI: {batas_revisi}

PEDOMAN AKTIF:
{_ped_prompt}

CATATAN DOSEN:
{catatan_dosen}

TUGAS:
1. Temukan sendiri BAB, subbab, dan paragraf yang paling relevan dengan catatan dosen, termasuk bila dosen tidak menyebut nomor paragraf.
2. Jangan merevisi bagian yang tidak berkaitan dengan catatan dosen.
3. Jangan mengubah judul, rumusan masalah, tujuan, variabel, metode, populasi/sampel, data, hasil, kutipan, atau sumber kecuali catatan dosen secara eksplisit memerintahkannya.
4. Jika catatan meminta perubahan substantif besar yang berdampak lintas bagian, beri PERINGATAN PERUBAHAN SUBSTANTIF dan jelaskan bagian terdampak. Jangan diam-diam mengubah seluruh naskah.
5. Jangan menciptakan data, referensi, DOI, halaman, kutipan, nomor hadis, atau fakta baru.
6. Ikuti pedoman aktif bila tersedia.
7. Tampilkan: LOKASI DITEMUKAN, SEBELUM, SESUDAH, ALASAN REVISI, lalu NASKAH REVISI LENGKAP.
8. Pertahankan bagian naskah lain apa adanya sejauh mungkin.

NASKAH:
{str(teks_revisi)[:100000]}"""
            _h = panggil_gemini(_prompt)
            if _h["sukses"]:
                st.session_state["hasil_revisi_dosen_ai"] = _h["hasil"]
            else:
                st.error(_h["error"])
        if st.session_state.get("hasil_revisi_dosen_ai"):
            st.text_area("Hasil Revisi Dosen", st.session_state["hasil_revisi_dosen_ai"], height=750, key="hasil_revisi_dosen_area")

    # --------------------------------------------------------
    # 3. AUDIT AKADEMIK
    # --------------------------------------------------------
    elif submenu_editor == "🔎 Audit Akademik":
        st.subheader("🔎 Audit Akademik")
        file_audit = st.file_uploader("Unggah naskah untuk diaudit", type=["pdf", "docx", "txt"], key="file_audit_akademik_ai")
        teks_audit = st.text_area("Atau tempel naskah", height=260, key="teks_audit_akademik_ai")
        if file_audit:
            _t = ekstrak_teks(file_audit)
            if not _t.startswith("ERROR:"):
                teks_audit = _t
        cakupan_audit = st.multiselect(
            "Cakupan Audit",
            ["Struktur & Pedoman", "Konsistensi Judul-Rumusan-Tujuan", "Teori & Penelitian Terdahulu", "Metode Penelitian", "Data/Angka yang Belum Terverifikasi", "Sitasi & Footnote", "Daftar Pustaka", "Bahasa & Typo", "Artefak AI/Markdown/LaTeX"],
            default=["Struktur & Pedoman", "Konsistensi Judul-Rumusan-Tujuan", "Metode Penelitian", "Sitasi & Footnote", "Daftar Pustaka", "Bahasa & Typo", "Artefak AI/Markdown/LaTeX"],
            key="cakupan_audit_akademik_ai"
        )
        if st.button("🔎 Audit Naskah", type="primary", disabled=not bool(str(teks_audit).strip()), key="btn_audit_akademik_ai"):
            _ped_prompt = (_ped_teks[:50000] + "\n\nRINGKASAN PEDOMAN:\n" + _ped_analisis[:15000]) if _ped_aktif else "Pedoman institusi belum aktif. Jangan menebak aturan institusi."
            _prompt = f"""Lakukan audit akademik ketat terhadap naskah berikut.
JENIS NASKAH: {jenis_naskah_editor}
CAKUPAN: {', '.join(cakupan_audit)}

PEDOMAN AKTIF:
{_ped_prompt}

ATURAN:
- Jangan memperbaiki naskah secara diam-diam. Audit dahulu.
- Bedakan TEMUAN KRITIS, TEMUAN PENTING, dan TEMUAN MINOR.
- Sebutkan lokasi BAB/subbab/paragraf atau frasa agar mudah ditemukan.
- Periksa konsistensi judul, rumusan/fokus, tujuan, variabel, hipotesis, metode, populasi/sampel, instrumen, dan analisis sesuai jenis penelitian.
- Tandai angka/data lapangan yang tampak tidak memiliki dasar dari naskah sebagai [PERLU VERIFIKASI], bukan dianggap benar.
- Tandai referensi/DOI/halaman/kutipan yang tidak dapat dipastikan sebagai [PERLU VERIFIKASI]. Jangan menciptakan pengganti.
- Deteksi typo, ejaan, tanda baca, istilah tidak konsisten, serta artefak $, *, #, ```, [^n], atau LaTeX mentah yang tidak semestinya tampil di Word.
- Jika pedoman tidak mengatur suatu hal, tulis 'Tidak ditemukan dalam pedoman aktif'.
- Akhiri dengan STATUS KESIAPAN: Belum Siap / Perlu Revisi / Hampir Siap / Siap Difinalisasi, disertai alasan singkat.

NASKAH:
{str(teks_audit)[:110000]}"""
            _h = panggil_gemini(_prompt)
            if _h["sukses"]:
                st.session_state["hasil_audit_akademik_ai"] = _h["hasil"]
            else:
                st.error(_h["error"])
        if st.session_state.get("hasil_audit_akademik_ai"):
            st.markdown(st.session_state["hasil_audit_akademik_ai"])

    # --------------------------------------------------------
    # 4. FINALISASI WORD
    # --------------------------------------------------------
    elif submenu_editor == "📑 Finalisasi Word":
        st.subheader("📑 Format Akademik & Finalisasi Word")
        st.caption("Tahap akhir untuk membersihkan naskah dan membentuk Word akademik tanpa mengubah substansi penelitian.")
        file_final = st.file_uploader("Unggah naskah Word / PDF / TXT", type=["pdf", "docx", "txt"], key="file_finalisasi_word_ai")
        teks_final = st.text_area("Atau tempel naskah", height=280, key="teks_finalisasi_word_ai")
        if file_final:
            _t = ekstrak_teks(file_final)
            if not _t.startswith("ERROR:"):
                teks_final = _t

        c1, c2, c3 = st.columns(3)
        with c1:
            gaya_final = st.selectbox("Gaya Sitasi", ["Chicago Notes & Bibliography (Footnote)", "Turabian Notes-Bibliography", "APA 7th", "IEEE"], key="gaya_final_word_ai")
            font_final = st.selectbox("Jenis Huruf", ["Times New Roman", "Cambria", "Arial", "Calibri"], key="font_final_word_ai")
        with c2:
            size_final = st.selectbox("Ukuran Isi", ["12 pt", "11 pt"], key="size_final_word_ai")
            judul_final = st.text_input("Judul Penelitian", value=st.session_state.get("mesin2_judul_terkunci", ""), key="judul_final_word_ai")
        with c3:
            nama_final = st.text_input("Nama Mahasiswa", key="nama_final_word_ai")
            npm_final = st.text_input("NPM", key="npm_final_word_ai")

        if _ped_aktif:
            st.success(f"Format akan mengikuti Pedoman Aktif: {_ped_nama or 'Pedoman Penulisan'}")
        st.info("Target Word: BAB halaman baru; daftar A./1./a. memakai tab dan hanging indent; tidak ada spasi manual; teks/tabel tidak melewati margin; sampul tanpa nomor; bagian awal Romawi kecil; halaman pertama BAB angka Latin (1, 2, 3, ...) di tengah bawah dan halaman berikutnya di kanan atas; Chicago/Turabian memakai footnote Word dan superscript.")

        if st.button("✨ Finalisasi Naskah", type="primary", disabled=not bool(str(teks_final).strip()), key="btn_finalisasi_word_ai"):
            _ped_prompt = (_ped_teks[:50000] + "\n\nRINGKASAN PEDOMAN:\n" + _ped_analisis[:15000]) if _ped_aktif else "Pedoman institusi belum aktif. Jangan menebak aturan institusi."
            _prompt = f"""Anda adalah finalisator naskah akademik.
JENIS NASKAH: {jenis_naskah_editor}
GAYA SITASI: {gaya_final}
FONT: {font_final} {size_final}

PEDOMAN AKTIF:
{_ped_prompt}

ATURAN WAJIB:
1. Jangan mengubah judul, jenis penelitian, variabel, hipotesis, lokasi, subjek, data angka, hasil, atau substansi ilmiah.
2. Perbaiki typo, ejaan, tanda baca, kalimat tidak efektif, pengulangan, dan konsistensi istilah.
3. Jangan menciptakan data lapangan, jumlah populasi/sampel, DOI, halaman sumber, nama penulis, judul sumber, nomor hadis, metadata bibliografis, atau kutipan. Tandai [PERLU VERIFIKASI] bila perlu.
4. Bersihkan artefak Markdown/LaTeX yang tidak semestinya: $, ```, #, ##, ###, **, dan marker format mentah. Jangan menghapus tanda kurung yang diperlukan secara ilmiah.
5. Hipotesis seperti $H_0$ dan $H_a$ harus menjadi H₀ dan Hₐ, tanpa tanda dolar.
6. Strukturkan BAB, subbab A./B./C., daftar 1./2./3., dan a./b./c. secara konsisten. Jangan memakai spasi manual untuk indentasi.
7. Chicago/Turabian: gunakan marker [^1] di narasi dan definisi [^1]: catatan di akhir teks agar mesin Word dapat mengubahnya menjadi footnote. APA 7 gunakan author-date. IEEE gunakan [1]. Jangan mencampur gaya.
8. Bedakan buku, jurnal, tesis/disertasi, regulasi, website, YouTube, Al-Qur'an, tafsir, hadis, kitab hadis/syarah, kitab klasik, dan wawancara.
9. Pertahankan teks Arab. Bila rusak karena ekstraksi, tandai [TEKS ARAB PERLU DIPERIKSA], jangan menebak.
10. Jangan menambahkan ringkasan perubahan ke dalam naskah final. Keluarkan hanya naskah final yang bersih.

NASKAH:
{str(teks_final)[:110000]}"""
            _h = panggil_gemini(_prompt)
            if _h["sukses"]:
                st.session_state["hasil_finalisasi_word_ai"] = _h["hasil"]
                st.success("Finalisasi teks selesai. Periksa preview sebelum Download Word.")
            else:
                st.error(_h["error"])

        if st.session_state.get("hasil_finalisasi_word_ai"):
            hasil_final = st.text_area("Preview Naskah Final", st.session_state["hasil_finalisasi_word_ai"], height=750, key="preview_finalisasi_word_ai")
            st.session_state["hasil_finalisasi_word_ai"] = hasil_final
            if judul_final.strip():
                try:
                    word_final = buat_docx_proposal_final(
                        hasil_final, judul_final, nama_final, npm_final,
                        "Pendidikan Agama Islam", datetime.now().year,
                        gaya_sitasi=gaya_final, font_naskah=font_final, ukuran_naskah=size_final
                    )
                    if word_final:
                        st.download_button(
                            "📥 Download Word Final (.docx)", data=word_final,
                            file_name="Naskah_Akademik_Final.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            key="download_final_word_ai"
                        )
                except Exception as e:
                    st.error(f"Word belum dapat dibuat: {e}")
            else:
                st.warning("Isi Judul Penelitian agar Word final dapat dibuat.")


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
            "🔎 Literatur & Penelitian Terdahulu",
            "📑 Proposal Tesis",
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
                        _hasil_pedoman_raw = panggil_gemini(_prompt_pedoman)
                        _hasil_pedoman = hasil_ai_teks(_hasil_pedoman_raw)
                        if _hasil_pedoman:
                            st.session_state["pedoman_tesis_s2_teks"] = str(_teks_pedoman)
                            st.session_state["pedoman_tesis_s2_analisis"] = _hasil_pedoman
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
                "pilih_ref_penguat_ide_s2", "kelayakan_ref_10_judul_s2",
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
            st.session_state["judul_alternatif_s2"] = [""] * 10
            for _i in range(10):
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
                "pilih_ref_penguat_ide_s2", "kelayakan_ref_10_judul_s2",
                "hasil_penulisan_ai", "naskah_aktif", "hasil_s2",
            }
            _frag_reset_topik = (
                "judul_alt_s2_", "editor_naskah_ide_s2_", "bahan_ide",
                "upload_ide", "unggah_ide", "file_bahan_ide",
            )
            for _key in list(st.session_state.keys()):
                _ks = str(_key)
                if _key in _hapus_reset_topik or any(_frag in _ks for _frag in _frag_reset_topik):
                    if _ks not in ("bank_referensi", "library_referensi"):
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

        if st.button("🤖 Analisis Ide, Gap, Novelty & 10 Judul", key="gen_ide_s2", type="primary"):
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
- Jika pilihan pengguna adalah "🤖 Rekomendasi AI", tentukan SATU metode yang paling sesuai setelah membaca masalah dan seluruh bahan. Jelaskan alasan singkat pada bagian analisis, lalu buat 10 alternatif judul yang konsisten dengan metode rekomendasi tersebut.
- Jika pengguna memilih metode tertentu, JANGAN menggantinya dengan metode lain. Analisis masalah, research gap, novelty, dan 10 alternatif judul harus konsisten dengan metode pilihan pengguna.
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

E. 10 ALTERNATIF JUDUL TESIS
Tulis TEPAT 10 alternatif judul. Untuk SETIAP judul gunakan format berikut agar dapat dibaca aplikasi:
[JUDUL 1] ...
[GAP 1] Uraikan research gap judul 1 secara ringkas dan spesifik.
[NOVELTY 1] Uraikan potensi novelty judul 1 secara ringkas dan spesifik.
[KEKUATAN 1] KUAT / CUKUP / LEMAH
[ALASAN 1] Jelaskan alasan penilaian kekuatan, fokus, kelayakan metode, dan risiko utama.
Ulangi pola yang sama sampai [JUDUL 10], [GAP 10], [NOVELTY 10], [KEKUATAN 10], [ALASAN 10].

F. CATATAN PEMILIHAN
Bandingkan 10 judul dan rekomendasikan 3 judul terkuat beserta alasan. Keputusan final tetap milik pengguna.

BERHENTI setelah bagian F. JANGAN LANJUT KE PROPOSAL."""

                with st.spinner("AI menganalisis masalah dan menyiapkan 10 alternatif judul..."):
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
                            "Silakan klik Generate sekali lagi. Tahap No.1 hanya boleh menghasilkan analisis dan 10 alternatif judul."
                        )
                    else:
                        st.session_state["hasil_ai_ide_judul_s2"] = _hasil_baru
                        st.session_state["masalah_terakhir_ide_s2"] = masalah_ide_s2.strip()
                        st.session_state["versi_naskah_ide_s2"] += 1

                        _judul_ai = re.findall(
                            r"(?im)^\s*\[JUDUL\s*(?:10|[1-9])\]\s*[:\-]?\s*(.+?)\s*$",
                            _hasil_baru,
                        )
                        _judul_ai = [
                            re.sub(r'^[\"\'“”]+|[\"\'“”]+$', "", j.strip())
                            for j in _judul_ai
                            if j.strip()
                        ]

                        if _judul_ai:
                            _judul_ai = (_judul_ai + [""] * 10)[:10]
                            st.session_state["judul_alternatif_s2"] = _judul_ai
                            # Bank judul membaca state ini sebagai nilai awal pada rerun.
                            st.session_state["muat_judul_ai_s2"] = True
                            st.success(
                                f"✅ Analisis selesai. {sum(bool(j) for j in _judul_ai)} alternatif judul ditemukan. "
                                "Anda yang memilih judul."
                            )
                        else:
                            st.warning(
                                "Analisis selesai, tetapi 10 judul belum terbaca otomatis. "
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
                st.session_state["judul_alternatif_s2"] = [""] * 10
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
            st.subheader("📊 Hasil Analisis Ide, Gap, Novelty & 10 Alternatif Judul")
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
Periksa masalah, gap awal, potensi novelty, 10 alternatif judul, dan kelayakan arah metode.
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
        # BANK 10 ALTERNATIF JUDUL — dapat diedit sebelum masuk proposal
        # ------------------------------------------------------------
        st.markdown("#### 🏷️ Bank Alternatif Judul")
        st.caption(
            "Setelah analisis masalah/ide, susun hingga 10 judul. "
            "Semua judul dapat diedit manual sebelum satu judul ditetapkan."
        )

        if "judul_alternatif_s2" not in st.session_state:
            st.session_state["judul_alternatif_s2"] = [""] * 10

        # Muat 10 judul hasil Generate ke widget hanya setelah Generate baru berhasil.
        if st.session_state.pop("muat_judul_ai_s2", False):
            for _i, _j in enumerate(st.session_state["judul_alternatif_s2"][:10]):
                st.session_state[f"judul_alt_s2_{_i}"] = _j

        # Pengguna bebas mengedit 10 judul sebelum memilih satu.
        judul_edit_s2 = []
        for _i in range(10):
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
            if st.button("📚 Cek Banyaknya Literatur untuk 10 Judul", key="cek_ref_10_judul_s2", use_container_width=True):
                _kel=[]
                with st.spinner("Mengecek ketersediaan referensi nyata untuk setiap judul..."):
                    for _j in judul_edit_s2:
                        if _j.strip():
                            _kel.append(analisis_ketersediaan_referensi_judul_s2(_j.strip(),20))
                st.session_state["kelayakan_ref_10_judul_s2"]=_kel
        if st.session_state.get("kelayakan_ref_10_judul_s2"):
            st.markdown("#### 📚 Banyaknya Literatur yang Tersedia untuk 10 Judul")
            st.caption("Angka berasal dari basis data nyata. Crossref, OpenAlex, dan Semantic Scholar ditampilkan terpisah dan tidak dijumlahkan karena dapat berisi karya yang sama.")
            for _no,_d in enumerate(st.session_state["kelayakan_ref_10_judul_s2"],1):
                with st.expander(f"Alternatif {_no} — {_d['Status']}", expanded=False):
                    st.write(f"**Judul:** {_d['Judul']}")
                    st.write(f"**Crossref:** {_d.get('Crossref',0):,} hasil")
                    st.write(f"**OpenAlex:** {_d.get('OpenAlex',0):,} hasil")
                    st.write(f"**Semantic Scholar:** {_d.get('Semantic Scholar',0):,} hasil")
                    st.write(f"**Kandidat unik yang diperiksa aplikasi:** {_d.get('Kandidat',0)}")
                    st.write(f"**Literatur 5 tahun terakhir dalam sampel:** {_d.get('Literatur 5 Tahun',0)}")
                    st.write(f"**Terverifikasi/teridentifikasi dalam sampel:** {_d.get('Terverifikasi',0)}")
                    st.write(f"**Status ketersediaan:** {_d['Status']}")
                    for _r in _d.get("Referensi",[])[:10]:
                        st.write(f"• {_r.get('Tahun','')} — {_r.get('Judul','')} [{_r.get('Sumber','')}]")

        # Gabungkan uraian Gap/Novelty/Kekuatan AI dengan data ketersediaan literatur nyata.
        _analisis_10 = _parse_analisis_10_judul(st.session_state.get("hasil_ai_ide_judul_s2", ""), judul_edit_s2)
        _kel_map = {str(x.get("Judul","")).strip(): x for x in st.session_state.get("kelayakan_ref_10_judul_s2", []) if isinstance(x,dict)}
        for _x in _analisis_10:
            _x["Ketersediaan"] = _kel_map.get(str(_x.get("Judul","")).strip(), {})

        if any(x.get("Judul") for x in _analisis_10):
            st.markdown("#### 🧾 Uraian Analisis 10 Judul")
            st.caption("Setiap judul dilengkapi Research Gap, Novelty, kekuatan, alasan, dan banyaknya literatur yang tersedia.")
            for _x in _analisis_10:
                if not _x.get("Judul"): continue
                _av=_x.get("Ketersediaan",{}) or {}
                _label=f"Judul {_x['No']} | {_x.get('Kekuatan','Belum dinilai') or 'Belum dinilai'} | {_av.get('Status','Literatur belum dicek')}"
                with st.expander(_label, expanded=False):
                    st.write(f"**Judul:** {_x.get('Judul','')}")
                    st.write(f"**Research Gap:** {_x.get('Gap','Belum dianalisis') or 'Belum dianalisis'}")
                    st.write(f"**Novelty:** {_x.get('Novelty','Belum dianalisis') or 'Belum dianalisis'}")
                    st.write(f"**Kekuatan:** {_x.get('Kekuatan','Belum dinilai') or 'Belum dinilai'}")
                    st.write(f"**Alasan:** {_x.get('Alasan','Belum dianalisis') or 'Belum dianalisis'}")
                    if _av:
                        st.write(f"**Banyaknya literatur:** Crossref {_av.get('Crossref',0):,} | OpenAlex {_av.get('OpenAlex',0):,} | Semantic Scholar {_av.get('Semantic Scholar',0):,}")
                        st.write(f"**Sampel unik diperiksa:** {_av.get('Kandidat',0)} | **5 tahun terakhir:** {_av.get('Literatur 5 Tahun',0)} | **Terverifikasi/teridentifikasi:** {_av.get('Terverifikasi',0)}")
                    else:
                        st.info("Klik 'Cek Banyaknya Literatur untuk 10 Judul' agar jumlah literatur nyata ditampilkan.")

            st.markdown("#### 💾 Simpan / Cetak Analisis Submenu 1")
            _ped_nama=str(st.session_state.get("pedoman_tesis_s2_nama","") or "")
            _judul_saat_ini=str(st.session_state.get("judul_utama_pilihan_s2","") or "")
            _docx=buat_docx_analisis_judul_s2(masalah_ide_s2,metode_ide_s2,_ped_nama,st.session_state.get("hasil_ai_ide_judul_s2",""),_analisis_10,_judul_saat_ini)
            _pdf=buat_pdf_analisis_judul_s2(masalah_ide_s2,metode_ide_s2,_ped_nama,st.session_state.get("hasil_ai_ide_judul_s2",""),_analisis_10,_judul_saat_ini)
            cexp1,cexp2,cexp3=st.columns(3)
            with cexp1:
                if _docx:
                    st.download_button("📄 Unduh Word",data=_docx,file_name="Analisis_10_Judul_Tesis_S2.docx",mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",use_container_width=True,key="unduh_word_analisis10_s2")
                else: st.warning("python-docx belum tersedia.")
            with cexp2:
                if _pdf:
                    st.download_button("📕 Unduh PDF",data=_pdf,file_name="Analisis_10_Judul_Tesis_S2.pdf",mime="application/pdf",use_container_width=True,key="unduh_pdf_analisis10_s2")
                else: st.warning("PDF belum dapat dibuat.")
            with cexp3:
                if st.button("💾 Simpan Analisis",key="simpan_analisis10_s2",use_container_width=True):
                    _hist=st.session_state.setdefault("riwayat_analisis_judul_s2",[])
                    _hist.append({"waktu":datetime.now().strftime("%d-%m-%Y %H:%M"),"masalah":masalah_ide_s2,"metode":metode_ide_s2,"pedoman":_ped_nama,"analisis":st.session_state.get("hasil_ai_ide_judul_s2",""),"judul":deepcopy(_analisis_10),"judul_dipilih":_judul_saat_ini})
                    st.success(f"Analisis tersimpan sebagai versi {_hist.__len__()} dalam proyek sesi ini.")
            if st.session_state.get("riwayat_analisis_judul_s2"):
                st.caption(f"Tersimpan {len(st.session_state['riwayat_analisis_judul_s2'])} versi analisis pada sesi/proyek aktif.")

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
                    "✅ Tetapkan Judul & Lanjutkan ke Literatur",
                    key="tetapkan_judul_s2",
                    use_container_width=True,
                ):
                    st.session_state["judul_tesis_s2_terpilih"] = judul_pilihan_s2

                    # Jalur transfer sederhana dan stabil khusus Submenu Proposal.
                    # State lama tetap dipertahankan agar fitur yang sudah berjalan tidak terganggu.
                    st.session_state["proposal_judul"] = judul_pilihan_s2
                    st.session_state["proposal_masalah"] = masalah_ide_s2
                    st.session_state["proposal_metode"] = metode_ide_s2
                    st.session_state["proposal_arah"] = arah_ide_s2

                    # DATA PROYEK TESIS S2: sumber utama antar-submenu.
                    st.session_state["proyek_tesis_s2"] = {
                        "judul": judul_pilihan_s2,
                        "masalah": masalah_ide_s2,
                        "arah": arah_ide_s2,
                        "mode": mode_ide_s2,
                        "metode": metode_ide_s2,
                        "analisis_ide": st.session_state.get("hasil_ai_ide_judul_s2", ""),
                        "analisis_10_judul": deepcopy(_analisis_10),
                        "ketersediaan_literatur_10_judul": deepcopy(st.session_state.get("kelayakan_ref_10_judul_s2", [])),
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
                        "Judul utama sudah ditetapkan dan siap diteruskan ke Literatur & Penelitian Terdahulu. "
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

        st.markdown("## 🧠 Dua Mesin Proposal Tesis")
        _tab_analisis, _tab_buat = st.tabs([
            "🔎 Mesin 1 · Analisis Proposal/Tesis",
            "✨ Mesin 2 · Buat Proposal Tesis AI",
        ])

        # ========================================================
        # MESIN 1 — ANALISIS DOKUMEN YANG SUDAH ADA
        # ========================================================
        with _tab_analisis:
            st.caption(
                "Untuk proposal milik sendiri, proposal orang lain sebagai bahan analisis, "
                "atau tesis lengkap. Mesin ini tidak menganggap judul dokumen sebagai judul yang harus dipertahankan."
            )
            _file_m1 = st.file_uploader(
                "Unggah proposal atau tesis",
                type=["pdf", "docx", "txt"],
                key="mesin1_upload_proposal_tesis",
            )
            _mode_m1 = st.radio(
                "Tujuan",
                [
                    "Analisis proposal/tesis dan berikan saran judul ulang",
                    "Buat proposal berdasarkan tesis lengkap yang diunggah",
                ],
                key="mesin1_tujuan_proposal_tesis",
            )

            if st.button(
                "🔎 Jalankan Mesin 1",
                key="jalankan_mesin1_proposal_tesis",
                use_container_width=True,
                disabled=(_file_m1 is None),
            ):
                try:
                    _teks_m1 = ekstrak_teks(_file_m1)
                    if not _teks_m1 or str(_teks_m1).startswith("ERROR:"):
                        st.error("Dokumen belum dapat dibaca.")
                    else:
                        _ped_m1 = st.session_state.get("pedoman_tesis_s2_teks", "")
                        if _mode_m1.startswith("Analisis"):
                            _prompt_m1 = f"""Anda adalah reviewer akademik tingkat magister.
Analisis dokumen yang diunggah secara kritis dan konstruktif. Dokumen boleh milik pengguna
atau milik orang lain sebagai bahan kajian. Jangan menganggap judul asli wajib dipertahankan.

PEDOMAN TESIS AKTIF:
{_ped_m1[:45000] if _ped_m1 else "Pedoman aktif tidak tersedia. Jangan mengarang aturan institusi."}

DOKUMEN:
{str(_teks_m1)[:100000]}

TUGAS:
1. Identifikasi jenis dokumen dan topik utamanya.
2. Analisis kualitas JUDUL ASLI: fokus, variabel/fokus, subjek/objek, lokasi bila relevan,
   keluasan, kebaruan, dan kesesuaian dengan metode.
3. Berikan 5 SARAN JUDUL ULANG yang lebih kuat. Untuk setiap judul, jelaskan alasan singkat.
   Pilih 1 judul yang paling direkomendasikan.
4. Analisis latar belakang, rumusan/fokus masalah, tujuan, signifikansi, definisi istilah,
   penelitian terdahulu, kajian teori, research gap, novelty, kerangka pikir, dan metode.
5. Periksa konsistensi: judul -> masalah -> tujuan -> teori -> metode.
6. Periksa kutipan, catatan kaki/footnote, dan daftar pustaka. Tunjukkan sumber yang tidak sinkron,
   kurang lengkap, atau perlu diverifikasi. Jangan mengarang metadata pengganti.
7. Jika Pedoman aktif tersedia, nilai kesesuaiannya dengan Pedoman.
8. Bedakan dengan jelas: KEKUATAN, KELEMAHAN, SARAN PERBAIKAN, SARAN JUDUL ULANG,
   dan PRIORITAS REVISI.
Jangan menulis ulang seluruh proposal kecuali diminta.
"""
                        else:
                            _prompt_m1 = f"""Anda adalah penyusun proposal tesis S2.
Sumber utama adalah TESIS LENGKAP yang diunggah. Ubah tesis tersebut menjadi PROPOSAL PENELITIAN
yang logis seolah penelitian BELUM dilakukan.

PEDOMAN TESIS AKTIF:
{_ped_m1[:45000] if _ped_m1 else "Pedoman aktif tidak tersedia. Gunakan sistematika proposal yang diberikan di bawah."}

TESIS SUMBER:
{str(_teks_m1)[:115000]}

ATURAN:
- Pertahankan substansi, judul, masalah, teori, metode, dan referensi yang memang berasal dari tesis,
  tetapi judul boleh disarankan ulang jika secara akademik lebih kuat.
- HAPUS/UBAH semua hasil penelitian, temuan, angka hasil, simpulan hasil, dan bahasa retrospektif
  menjadi bahasa RENCANA PENELITIAN. Jangan membawa hasil tesis sebagai hasil proposal.
- Jangan menciptakan data, DOI, halaman, kutipan, atau sumber yang tidak ada.
- Susun: BAGIAN AWAL; BAB I PENDAHULUAN; BAB II KAJIAN PUSTAKA DAN KERANGKA PIKIR;
  BAB III METODE PENELITIAN; SISTEMATIKA PENULISAN; DAFTAR PUSTAKA.
- Pertahankan kutipan/catatan kaki yang dapat ditelusuri dari tesis sumber.
- Gunakan bahasa akademik tingkat S2.
Keluarkan proposal lengkap, bukan laporan analisis.
"""
                        with st.spinner("Mesin 1 sedang membaca dokumen..."):
                            _hasil_m1 = hasil_ai_teks(
                                _panggil_gemini_rest_aman(_prompt_m1, temperature=0.2, max_output_tokens=16384)
                            )
                        if _hasil_m1:
                            st.session_state["mesin1_hasil_proposal_tesis"] = _hasil_m1
                            st.success("✅ Mesin 1 selesai.")
                        else:
                            st.error("AI belum menghasilkan analisis.")
                except Exception as _e:
                    st.error(f"Mesin 1 belum dapat dijalankan: {_e}")

            if st.session_state.get("mesin1_hasil_proposal_tesis"):
                st.text_area(
                    "Hasil Mesin 1",
                    value=st.session_state["mesin1_hasil_proposal_tesis"],
                    height=650,
                    key="mesin1_hasil_tampil_proposal_tesis",
                )

        # ========================================================
        # MESIN 2 — PROMPT AI FINAL: INPUT + FORMAT + SUMBER DIKUNCI
        # ========================================================
        with _tab_buat:
            st.caption(
                "Isi masalah, judul, jenis penelitian, gaya sitasi, dan format naskah. "
                "Semua pilihan menjadi acuan tetap sebelum AI menyusun proposal."
            )

            _masalah_m2 = st.text_area(
                "1. 📝 Masalah Penelitian",
                height=180,
                placeholder="Tuliskan masalah utama yang ingin diteliti...",
                key="mesin2_masalah_proposal_tesis",
            )

            _judul_m2 = st.text_input(
                "2. 📌 Judul Penelitian",
                placeholder="Ketik judul tesis yang akan digunakan...",
                key="mesin2_judul_baru_proposal_tesis",
            )

            _jenis_m2 = st.selectbox(
                "3. 🔬 Jenis Penelitian yang Digunakan",
                [
                    "R&D / Research and Development",
                    "Kualitatif",
                    "Kuantitatif",
                    "Action Research / PTK",
                    "Library Research / Penelitian Literatur",
                ],
                key="mesin2_jenis_proposal_tesis",
            )

            _gaya_m2 = st.selectbox(
                "4. 📚 Gaya Sitasi / Catatan Kaki",
                [
                    "Chicago Notes & Bibliography (Footnote)",
                    "Turabian Notes-Bibliography",
                    "APA 7th Edition",
                    "IEEE",
                ],
                key="mesin2_gaya_sitasi_proposal_tesis",
            )

            _c1_m2, _c2_m2 = st.columns(2)
            with _c1_m2:
                _font_m2 = st.selectbox(
                    "5. 🔤 Jenis Huruf",
                    ["Times New Roman", "Cambria", "Arial", "Calibri"],
                    key="mesin2_font_proposal_tesis",
                )
            with _c2_m2:
                _ukuran_m2 = st.selectbox(
                    "6. 🔢 Ukuran Huruf",
                    ["12 pt", "11 pt", "10 pt"],
                    key="mesin2_ukuran_font_proposal_tesis",
                )

            _kedalaman_m2 = st.selectbox(
                "7. 📖 Kedalaman Proposal",
                ["Mendalam S2", "Standar"],
                key="mesin2_kedalaman_proposal_tesis",
            )

            st.info(
                "🔒 Judul dan jenis penelitian dikunci. Referensi ilmiah kontemporer diprioritaskan "
                "tahun 2020–2026. Sumber sebelum 2020 hanya untuk sumber primer, teori seminal/original, "
                "Al-Qur'an, hadis, kitab/tafsir klasik, karya ulama klasik, atau regulasi yang masih relevan."
            )

            if st.button(
                "✨ Generik Prompt AI Proposal Tesis",
                key="jalankan_mesin2_proposal_tesis",
                type="primary",
                use_container_width=True,
            ):
                _kurang_m2 = []
                if not str(_masalah_m2 or "").strip():
                    _kurang_m2.append("Masalah Penelitian")
                if not str(_judul_m2 or "").strip():
                    _kurang_m2.append("Judul Penelitian")

                if _kurang_m2:
                    st.warning("Lengkapi terlebih dahulu: " + ", ".join(_kurang_m2) + ".")
                else:
                    _ped_m2 = st.session_state.get("pedoman_tesis_s2_teks", "")
                    _prompt_m2 = f"""Anda adalah penulis akademik dan asisten riset tingkat MAGISTER.
Susun proposal tesis S2 yang lengkap, mendalam, argumentatif, dan siap dibawa ke bimbingan.

INPUT PENGGUNA YANG SUDAH FINAL:
MASALAH PENELITIAN:
{_masalah_m2}

JUDUL PENELITIAN FINAL:
{_judul_m2}

JENIS PENELITIAN FINAL:
{_jenis_m2}

GAYA SITASI / RUJUKAN:
{_gaya_m2}

FORMAT NASKAH:
Jenis huruf: {_font_m2}
Ukuran huruf: {_ukuran_m2}
Kedalaman: {_kedalaman_m2}
Bahasa: Bahasa Indonesia akademik tingkat S2

PEDOMAN TESIS AKTIF:
{_ped_m2[:50000] if _ped_m2 else "Tidak tersedia. Jangan mengarang ketentuan institusi yang tidak diketahui."}

PENGUNCIAN:
1. JANGAN membuat alternatif judul.
2. JANGAN mengganti, memperpendek, memperluas, atau memparafrasekan JUDUL PENELITIAN FINAL.
3. JANGAN mengganti JENIS PENELITIAN FINAL.
4. Judul -> masalah -> rumusan/fokus -> tujuan -> teori -> gap -> novelty -> metode wajib konsisten.
5. Jangan membuat hasil penelitian seolah-olah penelitian telah dilaksanakan.

SISTEMATIKA PROPOSAL:

BAGIAN AWAL
- HALAMAN SAMPUL: PROPOSAL TESIS, judul final persis seperti input,
  [LOGO IAI DARUSSALAM MARTAPURA], Nama:, NPM:, Institut Agama Islam Darussalam Martapura,
  Pascasarjana, Program Studi Pendidikan Agama Islam, Martapura, tahun.
- KATA PENGANTAR.
- DAFTAR ISI.
- DAFTAR TABEL hanya jika benar-benar terdapat tabel. Jangan membuat halaman Daftar Tabel kosong.

BAB I PENDAHULUAN
A. Latar Belakang Masalah.
   Kembangkan secara mendalam: kondisi ideal, konteks ilmiah, masalah pengguna, urgensi,
   bukti literatur, kesenjangan, dan alasan penelitian. Jangan mengarang data lapangan.
B. Rumusan Masalah/Fokus Penelitian.
C. Tujuan Penelitian.
D. Signifikansi/Manfaat Penelitian.
E. Definisi Operasional/Istilah.

BAB II KAJIAN PUSTAKA DAN KERANGKA PIKIR
A. Penelitian Terdahulu.
   Gunakan beberapa penelitian yang benar-benar relevan dan mutakhir; jelaskan persamaan,
   perbedaan, keterbatasan, dan posisi penelitian.
B. Kajian Teori.
   Uraikan teori utama, konsep, dimensi/indikator bila relevan, serta hubungan antarkonsep.
C. Landasan Normatif Islam.
   Gunakan Al-Qur'an dan/atau hadis yang benar-benar relevan, terjemah, sumber tafsir/penjelasan,
   kemudian hubungkan DALIL -> MAKNA/TAFSIR -> KONSEP PENELITIAN.
   Perspektif Ahlussunnah wal Jamaah harus akademik dan relevan, bukan tempelan.
D. Research Gap dan Posisi Penelitian.
E. Kebaruan/Novelty Penelitian bila relevan.
F. Kerangka Pikir/Kerangka Konseptual.
G. Asumsi Dasar dan Hipotesis hanya jika sesuai jenis penelitian.

BAB III METODE PENELITIAN
Gunakan HANYA keluarga metode pada JENIS PENELITIAN FINAL.
- R&D: harus R&D, bukan diubah menjadi kualitatif. Tentukan model pengembangan yang tepat,
  alasan pemilihan, tahap pengembangan, produk/model, subjek uji coba/pengguna, validator bila perlu,
  instrumen, pengumpulan data, validasi/uji coba, dan analisis data.
- Kualitatif: pendekatan/jenis, lokasi, subjek/informan, objek/fokus, data/sumber,
  pengumpulan, instrumen, analisis, keabsahan, tahapan, dan etika bila relevan.
- Kuantitatif: desain, variabel, populasi/sampel, definisi operasional, instrumen,
  validitas/reliabilitas, pengumpulan data, dan analisis statistik.
- Action Research/PTK: setting/subjek, desain/siklus, tindakan, observasi, instrumen,
  indikator keberhasilan, dan analisis.
- Library Research: pendekatan, sumber primer/sekunder, penelusuran literatur,
  kritik/validasi sumber, analisis, dan sintesis.
JANGAN mencampur metode yang tidak kompatibel.

SETELAH BAB III
- SISTEMATIKA PENULISAN sesuai Pedoman aktif.
- DAFTAR PUSTAKA SEMENTARA.

ATURAN SUMBER DAN TAHUN:
1. Referensi ilmiah kontemporer untuk jurnal, buku akademik, penelitian terdahulu,
   teknologi, pendidikan, kurikulum, dan kajian mutakhir diprioritaskan tahun 2020–2026.
2. Penelitian terdahulu WAJIB mengutamakan tahun 2020–2026.
3. Sumber sebelum 2020 hanya diperbolehkan jika memang diperlukan sebagai:
   sumber primer, teori seminal/original theory, Al-Qur'an, hadis, kitab/tafsir klasik,
   karya ulama klasik, atau regulasi yang masih berlaku/relevan.
4. Jangan menggunakan sumber lama hanya karena lebih mudah diingat.
5. Targetkan dukungan referensi proposal yang kaya dan relevan, sekitar 20–30 sumber bila
   dapat dipertanggungjawabkan. Kualitas dan keterverifikasian lebih penting daripada jumlah.
6. Jangan mengarang penulis, judul, tahun, penerbit, jurnal, volume, nomor, halaman, DOI, URL,
   kutipan langsung, atau metadata apa pun.
7. Jika detail bibliografis belum cukup pasti, jangan membuat detail palsu.
   Catat sebagai kandidat yang memerlukan verifikasi di luar naskah final.
8. Regulasi harus disebut dengan identitas yang benar dan jangan mengarang nomor regulasi.

ATURAN KUTIPAN, CATATAN KAKI, DAN JENIS SUMBER:
1. Terapkan SATU gaya sitasi secara konsisten sesuai GAYA SITASI FINAL. Jangan mencampur Chicago, Turabian, APA, dan IEEE.
2. Jika Chicago Notes & Bibliography atau Turabian dipilih:
   - Pada narasi gunakan marker [^1], [^2], [^3], dst. tepat setelah klaim/kutipan yang dirujuk.
   - Setelah naskah, tulis definisi catatan: [^1]: isi catatan kaki lengkap. Marker ini akan diubah aplikasi menjadi nomor superscript dan true Word footnote.
   - Catatan pertama suatu sumber ditulis lengkap; pengulangan berikutnya gunakan bentuk singkat yang sesuai gaya, tanpa mengarang halaman.
3. Jika APA 7 dipilih: gunakan author-date dalam teks, misalnya (Nama, 2024) atau (Nama, 2024, p. 25) hanya jika halaman benar-benar diketahui. Jangan membuat footnote bibliografis Chicago.
4. Jika IEEE dipilih: gunakan [1], [2], dst. menurut urutan kemunculan dan daftar referensi IEEE. Jangan membuat footnote Chicago.
5. BEDAKAN FORMAT MENURUT JENIS SUMBER:
   a. BUKU: penulis, judul buku, data penerbitan, dan halaman bila halaman benar-benar diketahui.
   b. ARTIKEL JURNAL: penulis, judul artikel, nama jurnal, volume, nomor, tahun, rentang/halaman yang terverifikasi, DOI bila benar-benar ada.
   c. TESIS/DISERTASI: penulis, judul, jenis karya, institusi, tahun, dan halaman bila diketahui.
   d. WEBSITE: penulis/lembaga, judul halaman, nama situs, tanggal publikasi/pembaruan bila tersedia, URL yang benar, serta tanggal akses hanya jika diwajibkan gaya/pedoman.
   e. VIDEO YOUTUBE: pembuat/nama kanal, judul video, YouTube, tanggal publikasi, URL; gunakan timestamp bila mengutip bagian tertentu dan timestamp benar-benar diketahui.
   f. PERATURAN: nama resmi, nomor dan tahun peraturan, serta pasal/bagian bila relevan dan terverifikasi.
   g. AL-QUR'AN: jangan diperlakukan sebagai jurnal/buku biasa. Tulis nama surah dan nomor ayat pada kutipan sesuai Pedoman aktif; tampilkan teks Arab/terjemah hanya bila relevan. Sumber terjemahan mengikuti pedoman institusi.
   h. KITAB TAFSIR: nama mufasir, judul kitab tafsir, jilid/volume, data edisi/penerbitan, halaman yang benar-benar digunakan. Tafsir adalah sumber berbeda dari ayat Al-Qur'an.
   i. HADIS: sebutkan sumber hadis/riwayat, kitab/bab/nomor hadis hanya jika terverifikasi. Jangan mengarang sanad, nomor, derajat, atau lokasi hadis.
   j. KITAB HADIS/SYARAH HADIS: penyusun/pensyarah, judul kitab, jilid, data edisi/penerbitan, halaman yang digunakan bila terverifikasi.
   k. KITAB KLASIK: nama ulama, judul kitab, jilid, data edisi/penerbitan dan halaman berdasarkan edisi yang benar-benar digunakan.
   l. WAWANCARA: nama narasumber, jenis wawancara, tempat/media dan tanggal; perlakuan dalam Daftar Pustaka mengikuti Pedoman aktif.
6. Semua klaim teori, definisi, regulasi, penelitian terdahulu, tafsir, hadis, dan fakta ilmiah yang memerlukan rujukan harus memiliki sumber.
7. Kutipan langsung hanya jika teks dan halaman benar-benar diketahui. Jika halaman tidak diketahui, lakukan parafrase dan JANGAN mengarang halaman.
8. Semua sumber yang dikutip harus sinkron dengan Daftar Pustaka, kecuali jenis sumber yang oleh Pedoman aktif diperlakukan khusus.
9. Jangan membuat sumber, DOI, URL, halaman, nomor hadis, volume, penerbit, atau metadata palsu. Bila belum terverifikasi, jangan menyamarkannya sebagai sumber final.
10. PEDOMAN TESIS AKTIF mengalahkan aturan umum gaya sitasi apabila terdapat perbedaan format institusional.

KUALITAS AKADEMIK:
- Tulis sebagai proposal tesis S2, bukan ringkasan atau outline.
- BAB I dan BAB II harus mendalam dan argumentatif.
- Penelitian terdahulu harus benar-benar membantu membangun research gap.
- Novelty tidak boleh diklaim hanya dengan kalimat "belum pernah diteliti"; jelaskan posisi kebaruannya.
- Jangan mengarang kondisi empiris, jumlah responden/subjek, hasil observasi, wawancara, statistik,
  temuan, efektivitas, atau hasil uji yang belum dilakukan.
- Informasi yang belum diketahui ditulis sebagai rencana penelitian.
- Nama dan NPM tetap kosong jika belum diberikan.
- Jangan memasukkan Simulasi Seminar Proposal.
- Keluarkan naskah proposal lengkap, bukan penjelasan proses AI.

FORMAT:
- Gunakan {_font_m2} {_ukuran_m2} sebagai metadata format yang harus diterapkan saat ekspor Word.
- Jangan menulis instruksi format ini sebagai isi proposal.
"""
                    try:
                        with st.spinner("Mesin 2 sedang menyusun proposal, kutipan, dan referensi..."):
                            _hasil_m2 = hasil_ai_teks(
                                _panggil_gemini_rest_aman(
                                    _prompt_m2,
                                    temperature=0.30,
                                    max_output_tokens=16384
                                )
                            )
                        if _hasil_m2:
                            st.session_state["mesin2_hasil_proposal_tesis"] = _hasil_m2
                            st.session_state["proposal_s2_draf_otomatis"] = _hasil_m2
                            st.session_state["mesin2_judul_terkunci"] = str(_judul_m2).strip()
                            st.session_state["mesin2_jenis_terkunci"] = _jenis_m2
                            st.session_state["mesin2_gaya_terkunci"] = _gaya_m2
                            st.session_state["mesin2_font_terkunci"] = _font_m2
                            st.session_state["mesin2_ukuran_terkunci"] = _ukuran_m2
                            st.success("✅ Proposal Tesis AI selesai dibuat.")
                        else:
                            st.error("AI belum menghasilkan proposal.")
                    except Exception as _e:
                        st.error(f"Mesin 2 belum dapat dijalankan: {_e}")

            if st.session_state.get("mesin2_hasil_proposal_tesis"):
                _judul_kunci = st.session_state.get("mesin2_judul_terkunci", "")
                _jenis_kunci = st.session_state.get("mesin2_jenis_terkunci", "")
                _gaya_kunci = st.session_state.get("mesin2_gaya_terkunci", "")
                _font_kunci = st.session_state.get("mesin2_font_terkunci", "")
                _ukuran_kunci = st.session_state.get("mesin2_ukuran_terkunci", "")

                st.markdown("#### 🔒 Pengaturan Proposal")
                st.caption(f"Judul: {_judul_kunci}")
                st.caption(f"Jenis penelitian: {_jenis_kunci}")
                st.caption(f"Gaya sitasi: {_gaya_kunci}")
                st.caption(f"Format: {_font_kunci} {_ukuran_kunci}")
                st.caption("Referensi kontemporer: diprioritaskan 2020–2026")

                st.text_area(
                    "Proposal Tesis Hasil Mesin 2",
                    value=st.session_state["mesin2_hasil_proposal_tesis"],
                    height=800,
                    key="mesin2_hasil_tampil_proposal_tesis",
                )

                _hasil_lower = str(st.session_state["mesin2_hasil_proposal_tesis"]).lower()
                _ada_dp = ("daftar pustaka" in _hasil_lower) or ("references" in _hasil_lower)
                _ada_rujukan = any(x in _hasil_lower for x in [
                    "catatan kaki", "footnote", "doi.org", "et al.", "vol.", "no.",
                    "[1]", "(2020", "(2021", "(2022", "(2023", "(2024", "(2025", "(2026"
                ])
                if _ada_dp and _ada_rujukan:
                    st.success("🟢 Indikator awal: naskah sudah memuat unsur rujukan dan daftar pustaka. Tetap verifikasi sumber di Literatur & Referensi.")
                else:
                    st.warning("🟡 Perlu penyempurnaan: unsur kutipan/rujukan atau daftar pustaka belum terdeteksi lengkap.")

                # Satu tombol unduh Word untuk hasil Mesin 2. Format mengikuti pilihan pengguna.
                try:
                    _docx_m2 = buat_docx_proposal_final(
                        st.session_state["mesin2_hasil_proposal_tesis"],
                        _judul_kunci,
                        nama="",
                        npm="",
                        prodi="Pendidikan Agama Islam",
                        tahun=datetime.now().year,
                        gaya_sitasi=_gaya_kunci,
                        font_naskah=_font_kunci or "Times New Roman",
                        ukuran_naskah=_ukuran_kunci or "12 pt",
                    )
                    if _docx_m2:
                        _safe_title = re.sub(r"[^A-Za-z0-9_-]+", "_", str(_judul_kunci or "Proposal_Tesis"))[:80].strip("_")
                        st.download_button(
                            "📥 Download Proposal Word (.docx)",
                            data=_docx_m2,
                            file_name=f"{_safe_title or 'Proposal_Tesis'}.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            key="download_word_mesin2_proposal_tesis",
                            use_container_width=True,
                        )
                        if "chicago" in str(_gaya_kunci).lower() or "turabian" in str(_gaya_kunci).lower():
                            st.caption("Word: nomor kutipan superscript + footnote asli di bawah halaman, Times New Roman 10 pt, spasi 1.")
                        else:
                            st.caption("Word mengikuti gaya sitasi yang dipilih; footnote Chicago tidak dipaksakan pada APA/IEEE.")
                except Exception as _e_word:
                    st.warning(f"Word belum dapat dibuat: {_e_word}")

        st.divider()
        st.caption("Ruang kerja proposal lama tetap tersedia di bawah untuk menjaga fungsi aplikasi yang sudah berjalan.")

        # Semua dasar Proposal diambil otomatis dari hasil final Submenu 1.
        _proyek = st.session_state.get("proyek_tesis_s2", {})
        _dasar = st.session_state.get("dasar_proposal_tesis_s2", {})
        if not isinstance(_proyek, dict):
            _proyek = {}
        if not isinstance(_dasar, dict):
            _dasar = {}

        # Pulihkan data lama tanpa meminta pengguna mengulang Submenu 1.
        # Prioritas: data proyek -> dasar proposal -> judul final -> pilihan judul yang tersimpan.
        _judul_prop = (
            st.session_state.get("proposal_judul", "")
            or _proyek.get("judul")
            or _dasar.get("judul")
            or st.session_state.get("judul_tesis_s2_terpilih", "")
            or st.session_state.get("judul_utama_pilihan_s2", "")
        )
        _masalah_prop = (
            st.session_state.get("proposal_masalah", "")
            or _proyek.get("masalah")
            or _dasar.get("masalah", "")
            or st.session_state.get("masalah_ide_s2", "")
        )
        # Dukungan untuk widget masalah yang memakai nonce versi.
        if not _masalah_prop:
            _versi_masalah = st.session_state.get("versi_input_masalah_ide_s2", 0)
            _masalah_prop = st.session_state.get(f"masalah_ide_s2_{_versi_masalah}", "")

        _arah_prop = (
            st.session_state.get("proposal_arah", "")
            or _proyek.get("arah")
            or _dasar.get("arah", "")
            or st.session_state.get("arah_ide_s2", "")
        )
        _metode_prop = (
            st.session_state.get("proposal_metode", "")
            or _proyek.get("metode")
            or st.session_state.get("metode_ide_s2", "Belum ditentukan")
        )

        # Migrasikan otomatis data lama ke format proyek baru agar submenu berikutnya stabil.
        if str(_judul_prop).strip() and not _proyek.get("judul"):
            st.session_state["proyek_tesis_s2"] = {
                "judul": _judul_prop,
                "masalah": _masalah_prop,
                "arah": _arah_prop,
                "mode": _dasar.get("mode", st.session_state.get("mode_ide_s2", "")),
                "metode": _metode_prop,
            }
        _pedoman_prop = st.session_state.get("pedoman_tesis_s2_analisis", "") if st.session_state.get("pedoman_tesis_s2_aktif") else ""
        _bank_prop = st.session_state.get("bank_bahan_ide_s2", [])
        _refs_prop = st.session_state.get("bank_referensi", [])

        # Proposal tidak boleh terkunci walaupun state Submenu 1 belum terbaca.
        # Bila data otomatis tersedia, kolom langsung terisi. Bila belum, pengguna tetap dapat mengetik.
        if "judul_proposal_s2_edit" not in st.session_state:
            st.session_state["judul_proposal_s2_edit"] = str(_judul_prop or "")
        elif _judul_prop and not st.session_state.get("judul_proposal_s2_edit"):
            st.session_state["judul_proposal_s2_edit"] = str(_judul_prop)

        if "masalah_proposal_s2_edit" not in st.session_state:
            st.session_state["masalah_proposal_s2_edit"] = str(_masalah_prop or "")
        elif _masalah_prop and not st.session_state.get("masalah_proposal_s2_edit"):
            st.session_state["masalah_proposal_s2_edit"] = str(_masalah_prop)

        _judul_prop = st.text_input(
            "Judul Tesis",
            key="judul_proposal_s2_edit",
            placeholder="Otomatis dari Submenu 1, atau dapat diketik di sini",
        )
        _masalah_prop = st.text_area(
            "Permasalahan / konteks awal",
            key="masalah_proposal_s2_edit",
            height=120,
            placeholder="Otomatis dari Submenu 1, atau dapat diketik di sini",
        )

        if str(_judul_prop).strip():
            st.session_state["judul_tesis_s2_terpilih"] = _judul_prop
            st.session_state["proposal_judul"] = _judul_prop
            st.session_state["proposal_masalah"] = _masalah_prop
            st.session_state["proposal_metode"] = _metode_prop
            st.session_state["proposal_arah"] = _arah_prop
            _proyek_baru = st.session_state.get("proyek_tesis_s2", {})
            if not isinstance(_proyek_baru, dict):
                _proyek_baru = {}
            _proyek_baru.update({
                "judul": _judul_prop,
                "masalah": _masalah_prop,
                "arah": _arah_prop,
                "metode": _metode_prop,
            })
            st.session_state["proyek_tesis_s2"] = _proyek_baru
            st.success(f"🎓 Judul aktif: {_judul_prop}")
        else:
            st.info("Judul dari Submenu 1 belum terbaca. Kolom tetap dapat diketik tanpa kembali ke Submenu 1.")

        _ada_bahan_proposal = bool(
            str(_judul_prop or "").strip()
            or str(_masalah_prop or "").strip()
            or str(_arah_prop or "").strip()
            or (isinstance(_bank_prop, list) and len(_bank_prop) > 0)
            or (isinstance(_refs_prop, list) and len(_refs_prop) > 0)
        )

        if _ada_bahan_proposal:
            _nama_pedoman_prop = st.session_state.get("pedoman_tesis_s2_nama", "")
            if _pedoman_prop:
                st.success(f"🟢 Pedoman aktif otomatis: {_nama_pedoman_prop or 'Pedoman Tesis yang telah diaktifkan'}")
            else:
                st.caption("Pedoman institusi belum diaktifkan pada bagian Pedoman Penulisan Tesis.")

            st.caption(f"Metode: {_metode_prop} • Judul, masalah, pedoman, Bank Bahan, dan referensi digunakan otomatis.")

            with st.expander("➕ Tambah bahan khusus Proposal (opsional)", expanded=False):
                _tambahan_prop = st.file_uploader(
                    "PDF, DOCX, TXT",
                    type=["pdf","docx","txt"],
                    accept_multiple_files=True,
                    key="tambahan_bahan_proposal_s2",
                )

            _tahap_prop = [
                "Struktur isi sesuai Pedoman aktif",
                "Literatur final & penelitian terdahulu",
                "Metode sesuai jenis penelitian dalam Pedoman",
                "Audit kepatuhan Pedoman",
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

                _profil_pedoman = st.session_state.get("pedoman_tesis_s2_analisis", "")
                _refs_final = st.session_state.get("referensi_final_tesis_s2", []) or st.session_state.get("bank_referensi", [])
                _ref_ringkas = "\n".join(format_referensi(x) if isinstance(x,dict) else str(x) for x in _refs_final[:60])
                _prompt=f"""Anda adalah mesin penyusun PROPOSAL TESIS S2 berbasis PEDOMAN AKTIF.

PRINSIP MUTLAK:
1. PEDOMAN AKTIF adalah aturan tertinggi. Jangan membuat struktur sendiri.
2. Susun HANYA bagian yang diwajibkan/diizinkan oleh Pedoman untuk PROPOSAL, bukan struktur tesis lengkap.
3. DILARANG menambahkan Simulasi Seminar, prediksi pertanyaan penguji, BAB hasil penelitian, BAB penutup, atau bagian lain yang tidak termasuk proposal menurut Pedoman.
4. Sesuaikan metode dengan ketentuan khusus jenis penelitian dalam Pedoman.
5. Gunakan HANYA referensi final/terverifikasi yang diberikan. Jangan menciptakan penulis, tahun, DOI, halaman, teori, regulasi, kutipan, atau referensi baru.
6. Bila halaman sumber belum terverifikasi, jangan menebak nomor halaman. Tandai [HALAMAN PERLU VERIFIKASI] hanya bila footnote memerlukan halaman.
7. Gaya kutipan, footnote, daftar pustaka, transliterasi, Arab, Al-Qur'an dan hadis mengikuti Pedoman aktif. Bukan pilihan penulis.
8. Jangan tampilkan JSON, dict Python, metadata internal, kata 'sukses', 'hasil', atau informasi model AI.
9. Jangan mengarang data lapangan. Klaim empiris lokal yang belum memiliki data harus ditandai [DATA LAPANGAN PERLU DILENGKAPI].
10. Bahasa akademik harus alami, koheren, dan siap diaudit terhadap Pedoman.

JUDUL FINAL:
{_judul_prop}

PERMASALAHAN/KONTEKS:
{_masalah_prop}

ARAH PENELITIAN:
{_arah_prop}

METODE TERPILIH:
{_metode_prop}

TEKS PEDOMAN AKTIF (SUMBER ATURAN):
{st.session_state.get('pedoman_tesis_s2_teks','')[:50000] if st.session_state.get('pedoman_tesis_s2_teks') else 'Pedoman belum aktif.'}

PROFIL ATURAN PEDOMAN:
{_profil_pedoman if _profil_pedoman else 'Belum tersedia.'}

REFERENSI FINAL/TERVERIFIKASI:
{_ref_ringkas if _ref_ringkas else 'Belum ada referensi final. Jangan mengarang referensi; tandai kebutuhan sumber.'}

BANK BAHAN:
{_bank_ringkas if _bank_ringkas else 'Tidak ada dokumen khusus.'}

BAHAN TAMBAHAN:
{chr(10).join(_tambahan_teks) if _tambahan_teks else 'Tidak ada.'}

TUGAS:
A. Identifikasi terlebih dahulu sistematika PROPOSAL yang benar dari Pedoman aktif dan jenis penelitian terpilih.
B. Tulis proposal mengikuti urutan tersebut secara persis.
C. Gunakan referensi final pada klaim yang memang didukung. Jangan memaksakan sumber pada klaim yang tidak didukung.
D. Penelitian terdahulu harus berupa sintesis sumber final yang relevan, bukan daftar rekaan.
E. Akhiri hanya dengan bagian akhir proposal yang diwajibkan Pedoman.
F. Jangan menambahkan komentar AI setelah naskah proposal.
"""
                try:
                    _hasil_raw=panggil_gemini(_prompt)
                    _hasil=hasil_ai_teks(_hasil_raw)
                    if _hasil:
                        st.session_state["proposal_s2_draf_otomatis"]=_hasil
                        st.session_state["proposal_s2_editor_otomatis"]=_hasil
                        for _i in range(1,5):
                            st.session_state[f"proposal_s2_tahap_{_i}"]=True
                        st.rerun()
                    else:
                        st.error("AI belum menghasilkan draf proposal.")
                except Exception as _e:
                    st.error(f"Proposal belum dapat dibuat: {_e}")

            if st.session_state.get("proposal_s2_draf_otomatis"):
                st.success("✅ Draf proposal selesai.")
                # Terapkan hasil koreksi AI SEBELUM widget editor dibuat.
                # Streamlit melarang key widget diubah setelah widget terinstansiasi.
                _pending_editor = st.session_state.pop("proposal_s2_editor_pending", None)
                if _pending_editor is not None:
                    st.session_state["proposal_s2_editor_otomatis"] = _pending_editor
                elif "proposal_s2_editor_otomatis" not in st.session_state:
                    st.session_state["proposal_s2_editor_otomatis"] = st.session_state.get("proposal_s2_draf_otomatis", "")

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
                            _k_raw=panggil_gemini(_prompt_k)
                            _k=hasil_ai_teks(_k_raw)
                            if _k:
                                st.session_state["proposal_s2_draf_otomatis"]=_k
                                # Jangan menulis langsung ke key widget yang sudah dibuat pada run ini.
                                # Simpan sementara, lalu terapkan pada awal rerun sebelum text_area dibuat.
                                st.session_state["proposal_s2_editor_pending"]=_k
                                st.rerun()
                        except Exception as _e:
                            st.error(f"Koreksi belum dapat dilakukan: {_e}")

                st.markdown("### ✅ Audit Kepatuhan Pedoman")
                if st.button("🔍 Audit Proposal terhadap Pedoman",key="btn_audit_pedoman_proposal_s2",use_container_width=True):
                    _paudit=f"""Audit proposal berikut HANYA terhadap Pedoman aktif. Jangan mengarang aturan.
PEDOMAN:
{st.session_state.get('pedoman_tesis_s2_teks','')[:50000]}

PROPOSAL:
{_edit[:70000]}

Periksa: struktur dan urutan bagian proposal; kesesuaian dengan jenis penelitian; bagian wajib yang hilang; bagian terlarang/tidak ada dalam Pedoman; gaya kutipan/footnote; daftar pustaka; transliterasi/Arab/Al-Qur'an/Hadis bila ada; dan klaim/referensi yang belum terverifikasi.
Berikan status akhir tepat salah satu: LULUS PEDOMAN atau BELUM LULUS PEDOMAN. Jangan menyatakan format fisik DOCX (margin/font/spasi) lulus hanya dari teks preview."""
                    _audit=hasil_ai_teks(panggil_gemini(_paudit,0.05))
                    if _audit: st.session_state["audit_pedoman_proposal_s2"]=_audit
                if st.session_state.get("audit_pedoman_proposal_s2"):
                    st.markdown(st.session_state["audit_pedoman_proposal_s2"])
                _audit_state = hasil_ai_teks(st.session_state.get("audit_pedoman_proposal_s2", ""))
                _audit_ok = ("LULUS PEDOMAN" in _audit_state) and ("BELUM LULUS PEDOMAN" not in _audit_state)
                if st.button("✅ Finalisasi Proposal",key="finalisasi_proposal_s2",type="primary",use_container_width=True,disabled=not _audit_ok):
                    st.session_state["proposal_s2_final"]=_edit
                    st.success("🔒 Proposal ditetapkan sebagai versi final setelah lulus audit Pedoman.")
                if not _audit_ok:
                    st.caption("Finalisasi dikunci sampai audit menyatakan LULUS PEDOMAN.")

                st.info(
                    "Proposal Tesis memiliki dua jalur kerja: "
                    "🔎 Analisis Proposal Tesis untuk proposal yang sudah ada, dan "
                    "✨ Buat Proposal Tesis AI untuk membuat proposal baru dari bahan minimal."
                )
                st.markdown("### ✨ Buat Proposal Tesis AI")
                st.caption(
                    "Mesin pembuat proposal baru. Cukup masukkan masalah penelitian, pilih jenjang Tesis S2, "
                    "dan tentukan jenis penelitian. Judul boleh belum ada. Artikel/bahan tambahan bersifat opsional."
                )

                _masalah_baru_ai = st.text_area(
                    "Masalah penelitian",
                    value=str(_masalah_prop or ""),
                    height=140,
                    placeholder="Tuliskan masalah utama yang ingin diteliti...",
                    key="masalah_buat_proposal_tesis_ai",
                )
                _jenis_baru_ai = st.selectbox(
                    "Jenis penelitian",
                    [
                        "R&D / Research and Development",
                        "Kualitatif",
                        "Kuantitatif",
                        "Action Research / PTK",
                        "Library Research / Penelitian Literatur",
                    ],
                    index=0,
                    key="jenis_buat_proposal_tesis_ai",
                )
                st.text_input(
                    "Jenjang",
                    value="Tesis S2",
                    disabled=True,
                    key="jenjang_buat_proposal_tesis_ai",
                )

                with st.expander("➕ Bahan tambahan untuk Prompt Akhir (opsional)", expanded=False):
                    _bahan_prompt_akhir = st.file_uploader(
                        "Tambahkan artikel, jurnal, proposal awal, PDF, DOCX, atau TXT jika diperlukan",
                        type=["pdf", "docx", "txt"],
                        accept_multiple_files=True,
                        key="bahan_prompt_akhir_proposal_s2",
                    )
                    st.caption("Tidak wajib. Semua bahan yang sudah tersimpan di aplikasi tetap dibaca otomatis.")

                if st.button(
                    "✨ Buat Proposal Tesis AI",
                    key="prompt_akhir_proposal_tesis_s2",
                    type="primary",
                    use_container_width=True,
                ):
                    if not str(_masalah_baru_ai or "").strip():
                        st.warning("Masukkan masalah penelitian terlebih dahulu.")
                        st.stop()
                    _ped_prompt = st.session_state.get("pedoman_tesis_s2_teks", "")
                    _profil_prompt = st.session_state.get("pedoman_tesis_s2_analisis", "")
                    _refs_prompt = (
                        st.session_state.get("referensi_final_tesis_s2", [])
                        or st.session_state.get("bank_referensi", [])
                    )
                    _bank_prompt = st.session_state.get("bank_bahan_ide_s2", [])

                    _teks_bahan_prompt = []
                    for _bf in (_bahan_prompt_akhir or []):
                        try:
                            _bt = ekstrak_teks(_bf)
                            if _bt and not str(_bt).startswith("ERROR:"):
                                _teks_bahan_prompt.append(
                                    f"FILE: {getattr(_bf, 'name', 'dokumen')}\n{str(_bt)[:18000]}"
                                )
                        except Exception:
                            pass

                    _bank_prompt_teks = "\n\n".join(
                        str(x)[:9000] for x in (_bank_prompt[:20] if isinstance(_bank_prompt, list) else [])
                    ) if isinstance(_bank_prompt, list) else str(_bank_prompt)[:50000]

                    _refs_prompt_teks = "\n".join(
                        format_referensi(x) if isinstance(x, dict) else str(x)
                        for x in (_refs_prompt[:100] if isinstance(_refs_prompt, list) else [])
                    ) if isinstance(_refs_prompt, list) else str(_refs_prompt)[:60000]

                    _naskah_sebelumnya = (
                        st.session_state.get("proposal_s2_hasil_penyempurnaan_ai")
                        or st.session_state.get("proposal_s2_draf_otomatis")
                        or _edit
                        or ""
                    )

                    _judul_prompt = str(_judul_prop or "").strip()
                    _masalah_prompt = str(_masalah_baru_ai or _masalah_prop or "").strip()
                    _arah_prompt = str(_arah_prop or "").strip()
                    _metode_prompt = str(_jenis_baru_ai or _metode_prop or "").strip()

                    _prompt_akhir = f"""Anda adalah ASISTEN AKADEMIK AI tingkat magister.
Tugas Anda adalah MEMBUAT PROPOSAL TESIS S2 LENGKAP seperti ketika seorang mahasiswa meminta:
"Berdasarkan bahan yang saya miliki, buatkan proposal tesis lengkap yang layak diajukan."

Anda BUKAN sekadar editor naskah lama. Gunakan seluruh bahan yang tersedia sebagai ACUAN,
kemudian kembangkan proposal secara luas, mendalam, argumentatif, akademik, dan koheren.

============================================================
A. CARA MEMBACA INPUT
============================================================
Input pengguna boleh TIDAK LENGKAP.
Pengguna mungkin hanya mempunyai:
- masalah penelitian; atau
- judul; atau
- jenis penelitian; atau
- beberapa artikel/bahan; atau
- proposal setengah jadi; atau
- kombinasi sebagian dari semuanya.

Jika JUDUL belum tersedia tetapi masalah/bahan cukup jelas:
- rumuskan satu judul tesis yang paling konsisten dengan masalah, bidang PAI/madrasah, dan metode;
- jangan meminta pengguna mengisi ulang seluruh proposal.

Jika METODE sudah dipilih pengguna, hormati metode tersebut dan susun BAB III secara konsisten.
Jika metode benar-benar belum ditentukan, pilih metode yang paling logis dari tujuan penelitian
dan nyatakan pilihan tersebut secara akademik tanpa mencampur beberapa desain yang tidak kompatibel.

Naskah lama adalah BAHAN, bukan batas kreativitas akademik.
Pertahankan gagasan yang baik, tetapi Anda boleh menata ulang, memperluas, memperdalam,
dan menulis kembali bagian yang lemah agar menjadi satu proposal tesis yang utuh.

============================================================
B. ACUAN DAN HIERARKI
============================================================
1. PEDOMAN TESIS S2 AKTIF, bila tersedia, adalah rambu penulisan institusi.
2. SISTEMATIKA WAJIB di bawah menjadi kerangka proposal, kecuali Pedoman aktif secara tegas
   mensyaratkan penamaan/urutan berbeda.
3. Bahan aplikasi, artikel, bank bahan, referensi, masalah, judul, arah penelitian,
   metode, dan proposal lama menjadi sumber konteks.
4. Jangan mengubah proposal menjadi laporan hasil penelitian.

============================================================
C. SISTEMATIKA PROPOSAL WAJIB
============================================================
BAGIAN AWAL
1. HALAMAN SAMPUL
   Tampilkan elemen:
   - PROPOSAL TESIS
   - Judul
   - tempat untuk Logo IAI Darussalam Martapura
   - Nama
   - NPM
   - Institut Agama Islam Darussalam Martapura
   - Pascasarjana
   - Program Studi Pendidikan Agama Islam
   - Martapura
   - Tahun
   Jika Nama/NPM belum tersedia, gunakan label "Nama:" dan "NPM:" tanpa mengarang identitas.
2. KATA PENGANTAR
3. DAFTAR ISI
4. DAFTAR TABEL, hanya jika memang ada tabel.

BAB I PENDAHULUAN
A. Latar Belakang Masalah
B. Rumusan Masalah/Fokus Penelitian
C. Tujuan Penelitian
D. Signifikansi/Manfaat Penelitian
E. Definisi Operasional/Istilah

BAB II KAJIAN PUSTAKA DAN KERANGKA PIKIR
A. Penelitian Terdahulu
B. Kajian Teori
C. Landasan Normatif Islam yang relevan
   - Al-Qur'an dan/atau hadis yang relevan;
   - terjemah;
   - tafsir/penjelasan akademik;
   - hubungan DALIL -> TAFSIR/MAKNA -> KONSEP PENELITIAN;
   - perspektif Ahlussunnah wal Jamaah secara akademik dan proporsional.
D. Research Gap dan Posisi Penelitian
E. Kebaruan/Novelty Penelitian, bila relevan
F. Kerangka Pikir/Kerangka Konseptual
G. Asumsi Dasar dan Hipotesis, HANYA bila sesuai jenis penelitian.

BAB III METODE PENELITIAN
Susun SUBBAGIAN BAB III sesuai metode yang benar-benar digunakan.
- Kualitatif: pendekatan/jenis, lokasi, subjek/informan, objek/fokus, data/sumber data,
  teknik pengumpulan, instrumen, analisis, keabsahan, tahapan, etika bila relevan.
- Kuantitatif: desain, variabel, populasi/sampel, definisi operasional variabel,
  instrumen, validitas/reliabilitas, pengumpulan data, teknik analisis statistik.
- R&D: model pengembangan yang dipilih, tahapan, subjek uji/validator, produk,
  instrumen, validasi, uji coba, teknik pengumpulan dan analisis data.
- Action Research/PTK: setting/subjek, desain/siklus, tindakan, observasi,
  instrumen, indikator keberhasilan, analisis.
- Library Research: jenis/pendekatan, sumber primer/sekunder, teknik pengumpulan
  literatur, kritik/validasi sumber, teknik analisis.
JANGAN mencampur struktur metode yang tidak kompatibel.

SETELAH BAB III
- SISTEMATIKA PENULISAN, sesuai Pedoman aktif.
- DAFTAR PUSTAKA SEMENTARA.

JANGAN memasukkan Simulasi Seminar Proposal ke dalam naskah.

============================================================
D. STANDAR PENGEMBANGAN AKADEMIK
============================================================
- Latar belakang harus berkembang dari konteks ideal -> realitas/permasalahan ->
  dukungan kajian/literatur -> kesenjangan -> urgensi -> arah solusi/penelitian.
- Jangan membuat latar belakang hanya beberapa paragraf pendek.
- Penelitian terdahulu harus cukup kaya untuk menunjukkan peta penelitian,
  persamaan, perbedaan, keterbatasan studi terdahulu, dan posisi penelitian ini.
- Kajian teori harus benar-benar menjelaskan konsep/teori utama dan hubungan antarkonsep,
  bukan hanya daftar definisi.
- Research gap harus diturunkan dari penelitian terdahulu dan teori, bukan klaim kosong.
- Novelty harus proporsional dengan bukti yang tersedia.
- Kerangka pikir harus menunjukkan alur logis masalah -> teori/konsep -> proses penelitian -> sasaran.
- BAB III harus konsisten dengan rumusan masalah dan tujuan.
- Jaga konsistensi JUDUL -> LATAR BELAKANG -> RUMUSAN/FOKUS -> TUJUAN ->
  TEORI -> GAP -> KERANGKA PIKIR -> METODE.

============================================================
E. REFERENSI: BEBAS MEMPERKAYA, TETAPI JUJUR STATUSNYA
============================================================
Anda BOLEH mengusulkan sebanyak mungkin literatur akademik yang relevan untuk memperkaya proposal:
jurnal nasional/internasional, buku akademik, regulasi, sumber metodologi, tafsir, hadis,
dan sumber ilmiah lain yang relevan.

Bedakan dua kelompok:
1. SUMBER TERSEDIA/TERVERIFIKASI APLIKASI:
   boleh digunakan sebagai sumber yang sudah tersedia.
2. SUMBER TAMBAHAN USULAN AI:
   boleh diusulkan untuk memperkaya proposal, tetapi JANGAN mengarang DOI, URL,
   nomor halaman, volume, nomor jurnal, atau metadata yang tidak benar-benar diketahui.
   Tandai secara wajar sebagai kandidat yang perlu diverifikasi di menu Literatur & Referensi.

Jangan membuat daftar pustaka fiktif hanya untuk terlihat banyak.
Lebih baik memberikan kandidat bibliografis yang jujur untuk diverifikasi daripada metadata palsu.
Referensi yang sudah valid dari bahan pengguna jangan dihilangkan.

============================================================
F. BATAS KEBEBASAN AI
============================================================
AI bebas mengembangkan ARGUMENTASI, STRUKTUR PENJELASAN, ANALISIS TEORITIS,
SINTESIS LITERATUR, GAP, NOVELTY, KERANGKA PIKIR, dan PENJELASAN METODOLOGIS.

AI DILARANG mengarang:
- hasil observasi/wawancara;
- jumlah informan/responden yang belum diberikan;
- nama lokasi/madrasah yang belum diberikan;
- data statistik lapangan;
- hasil uji;
- temuan penelitian;
- kutipan langsung palsu;
- nomor halaman palsu;
- DOI/URL palsu;
- identitas mahasiswa.

Jika fakta lapangan belum tersedia, tulis secara metodologis sebagai rencana penelitian,
bukan sebagai temuan yang sudah terjadi.

============================================================
G. KELUARAN
============================================================
Keluarkan HANYA NASKAH PROPOSAL TESIS LENGKAP.
Jangan keluarkan laporan audit, komentar AI, JSON, instruksi kepada pengguna, atau penjelasan proses.
Gunakan bahasa Indonesia akademik tingkat S2.
Utamakan kelengkapan dan kedalaman substansi, bukan sekadar mengejar jumlah halaman.
Jangan meringkas hanya karena input pengguna sedikit.

============================================================
DATA DAN BAHAN APLIKASI
============================================================
JUDUL YANG TERSEDIA:
{_judul_prompt if _judul_prompt else "Belum tersedia. Rumuskan dari masalah dan bahan jika memungkinkan."}

MASALAH/KONTEKS:
{_masalah_prompt if _masalah_prompt else "Belum tersedia secara khusus."}

ARAH PENELITIAN:
{_arah_prompt if _arah_prompt else "Belum tersedia secara khusus."}

JENIS/METODE:
{_metode_prompt if _metode_prompt else "Belum ditentukan."}

PEDOMAN TESIS S2 AKTIF:
{_ped_prompt[:50000] if _ped_prompt else "Belum tersedia. Jangan mengarang ketentuan kampus."}

PROFIL/ANALISIS PEDOMAN:
{_profil_prompt[:18000] if _profil_prompt else "Belum tersedia."}

REFERENSI/LITERATUR YANG SUDAH ADA DI APLIKASI:
{_refs_prompt_teks[:55000] if _refs_prompt_teks else "Belum ada referensi tersimpan."}

BANK BAHAN APLIKASI:
{_bank_prompt_teks[:50000] if _bank_prompt_teks else "Belum ada Bank Bahan."}

BAHAN TAMBAHAN YANG BARU DIUNGGAH:
{chr(10).join(_teks_bahan_prompt)[:60000] if _teks_bahan_prompt else "Tidak ada."}

PROPOSAL/DRAF SEBELUMNYA, JIKA ADA:
{_naskah_sebelumnya[:85000] if _naskah_sebelumnya else "Belum ada. Susun proposal dari bahan parsial yang tersedia."}
"""

                    try:
                        with st.spinner("AI sedang menyusun Proposal Tesis lengkap dari seluruh bahan aplikasi..."):
                            _prompt_raw = _panggil_gemini_rest_aman(
                                _prompt_akhir,
                                temperature=0.35,
                                max_output_tokens=16384,
                            )
                            _prompt_teks = hasil_ai_teks(_prompt_raw)

                        if not _prompt_teks:
                            st.error("AI belum menghasilkan Proposal Tesis.")
                        else:
                            st.session_state["proposal_s2_prompt_akhir"] = _prompt_teks
                            st.session_state["proposal_s2_hasil_penyempurnaan_ai"] = _prompt_teks
                            st.session_state["proposal_s2_draf_otomatis"] = _prompt_teks
                            st.session_state["proposal_s2_editor_pending"] = _prompt_teks
                            st.session_state.pop("audit_pedoman_proposal_s2", None)
                            st.success(
                                "✅ Pembuatan Proposal Tesis AI selesai. Proposal disusun dari bahan yang tersedia "
                                "di aplikasi dan dapat dilanjutkan ke Literatur & Referensi untuk verifikasi sumber."
                            )
                            st.rerun()
                    except Exception as _e:
                        st.error(f"Prompt Akhir Proposal Tesis belum dapat dijalankan: {_e}")

                if st.session_state.get("proposal_s2_prompt_akhir"):
                    with st.expander("📖 Lihat Hasil Proposal Tesis AI", expanded=False):
                        st.markdown(st.session_state["proposal_s2_prompt_akhir"])

                st.markdown("### 📄 Word Proposal untuk Bimbingan")
                with st.expander("Identitas untuk sampul Word", expanded=False):
                    _nama_word = st.text_input("Nama mahasiswa", key="proposal_final_nama_s2")
                    _npm_word = st.text_input("NPM", key="proposal_final_npm_s2")
                    _prodi_word = st.text_input("Program Studi", value="Pendidikan Agama Islam", key="proposal_final_prodi_s2")
                    _tahun_word = st.number_input("Tahun", min_value=2020, max_value=2100, value=datetime.now().year, step=1, key="proposal_final_tahun_s2")

                _naskah_word = (
                    st.session_state.get("proposal_s2_prompt_akhir")
                    or st.session_state.get("proposal_s2_hasil_penyempurnaan_ai")
                    or _edit
                )
                if st.button("📄 Siapkan Word Proposal untuk Bimbingan", key="siapkan_word_proposal_bimbingan_s2", use_container_width=True):
                    try:
                        _word = buat_docx_proposal_final(_naskah_word, _judul_prop, _nama_word, _npm_word, _prodi_word, int(_tahun_word))
                        st.session_state["proposal_s2_word_bimbingan"] = _word
                        st.success("✅ Word proposal untuk bimbingan berhasil disiapkan tanpa membuat ulang atau meringkas proposal.")
                    except Exception as _e:
                        st.error(f"Word belum dapat disiapkan: {_e}")

                _word_bimbingan = st.session_state.get("proposal_s2_word_bimbingan")
                if _word_bimbingan:
                    st.download_button(
                        "📥 Unduh Proposal Tesis untuk Bimbingan (.docx)",
                        data=_word_bimbingan,
                        file_name="Proposal_Tesis_Hasil_Prompt_Akhir_IAID.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        key="unduh_proposal_bimbingan_s2",
                        use_container_width=True,
                    )

    # ============================================================
    # SUBMENU 2 — LITERATUR & PENELITIAN TERDAHULU
    # Literatur disiapkan SEBELUM Proposal dan dikendalikan Pedoman aktif.
    # ============================================================
    if submenu_s2 == "🔎 Literatur & Penelitian Terdahulu":
        st.markdown("## 🔎 Literatur, Penelitian Terdahulu & Gaya Penulisan")
        _judul_lit = st.session_state.get("judul_tesis_s2_terpilih", "") or st.session_state.get("proposal_judul", "")
        _ped_teks = st.session_state.get("pedoman_tesis_s2_teks", "")
        _ped_analisis = st.session_state.get("pedoman_tesis_s2_analisis", "")
        if not _judul_lit:
            st.warning("Tetapkan satu judul final pada Submenu 1 terlebih dahulu.")
        else:
            st.success(f"🎓 Judul final: {_judul_lit}")
        if _ped_teks:
            st.success(f"🔒 Pedoman aktif: {st.session_state.get('pedoman_tesis_s2_nama','Pedoman Tesis')}")
            st.caption("Gaya penulisan tidak dipilih penulis. Footnote, daftar pustaka, transliterasi, Arab, Al-Qur'an/Hadis, struktur dan format mengikuti Pedoman aktif.")
            with st.expander("📐 Aturan Penulisan dari Pedoman Aktif", expanded=False):
                st.markdown(_ped_analisis or "Profil Pedoman belum dianalisis.")
        else:
            st.error("Pedoman belum aktif. Aktifkan Pedoman terlebih dahulu agar literatur dan gaya penulisan tidak menggunakan aturan generik.")

        if _judul_lit and _ped_teks and st.button("🤖 Analisis Kebutuhan Literatur Proposal", type="primary", use_container_width=True, key="analisis_kebutuhan_lit_s2"):
            _p=f"""Analisis kebutuhan literatur untuk proposal tesis berikut berdasarkan PEDOMAN AKTIF.
JUDUL: {_judul_lit}
METODE: {st.session_state.get('proposal_metode') or st.session_state.get('metode_terpilih_ide_s2') or ''}
PEDOMAN:
{_ped_teks[:45000]}

Keluarkan:
1. kebutuhan teori/konsep utama;
2. kebutuhan penelitian terdahulu;
3. kebutuhan regulasi/dokumen primer;
4. kebutuhan sumber metodologi sesuai jenis penelitian;
5. kebutuhan landasan Al-Qur'an/Hadis/Tafsir HANYA jika benar-benar relevan dengan substansi;
6. kata kunci pencarian sumber;
7. aturan footnote/daftar pustaka yang terdeteksi dari Pedoman.
Jangan membuat referensi. Jangan membuat ayat. Ini hanya peta kebutuhan sumber."""
            _hr=hasil_ai_teks(panggil_gemini(_p,0.15))
            if _hr: st.session_state["kebutuhan_literatur_tesis_s2"]=_hr
        if st.session_state.get("kebutuhan_literatur_tesis_s2"):
            st.markdown("### 🎯 Kebutuhan Kutipan Proposal")
            st.markdown(st.session_state["kebutuhan_literatur_tesis_s2"])

        st.markdown("### 🔍 Pencarian Referensi Terverifikasi")

        # Sinkronkan kata kunci Literatur setiap kali judul final dari Submenu 1 berubah.
        # Streamlit mempertahankan nilai widget berdasarkan key, sehingga parameter value=
        # saja tidak cukup untuk mengganti judul lama yang sudah tersimpan.
        _judul_lit_norm = str(_judul_lit or "").strip()
        _judul_lit_sumber_lama = str(st.session_state.get("_judul_literatur_s2_sumber", "") or "").strip()
        if _judul_lit_norm and _judul_lit_norm != _judul_lit_sumber_lama:
            st.session_state["kw_literatur_tesis_s2"] = _judul_lit_norm
            st.session_state["_judul_literatur_s2_sumber"] = _judul_lit_norm
            # Hasil pencarian judul lama tidak boleh terbawa ke judul baru.
            st.session_state.pop("kandidat_literatur_tesis_s2", None)
            st.session_state.pop("kebutuhan_literatur_tesis_s2", None)

        _kw=st.text_input("Kata kunci pencarian", key="kw_literatur_tesis_s2")
        _tahun_min=st.number_input("Prioritas tahun minimal", min_value=1900, max_value=datetime.now().year, value=2020, key="tahun_min_lit_s2")
        if st.button("🔎 Cari Referensi", use_container_width=True, key="cari_lit_tesis_s2"):
            _hasil_cari=cari_multi_sumber(_kw,12)
            _baru=[]
            for _r in _hasil_cari:
                try: _th=int(str(_r.get('Tahun',''))[:4])
                except Exception: _th=0
                _r=dict(_r); _r["Kelayakan Awal"]="🟢 Mutakhir" if _th>=int(_tahun_min) else "🟡 Pertimbangkan tahun/fondasional"
                _baru.append(_r)
            st.session_state["kandidat_literatur_tesis_s2"]=_baru
        _kands=st.session_state.get("kandidat_literatur_tesis_s2",[])
        if _kands:
            for _i,_r in enumerate(_kands):
                with st.container(border=True):
                    st.markdown(f"**{_r.get('Judul','Tanpa judul')}**")
                    st.caption(f"{_r.get('Penulis','')} | {_r.get('Tahun','')} | {_r.get('Sumber','')} | {_r.get('Kelayakan Awal','')}")
                    st.write(_r.get('Status',''))
                    _a,_b=st.columns(2)
                    with _a:
                        if st.button("✅ Gunakan sebagai Referensi Final",key=f"final_lit_{_i}",use_container_width=True):
                            _final=st.session_state.setdefault("referensi_final_tesis_s2",[])
                            if not any(kunci_ref(x)==kunci_ref(_r) for x in _final): _final.append(dict(_r))
                            tambah_bank_referensi(_r); st.rerun()
                    with _b:
                        if st.button("🔄 Cari Pengganti",key=f"ganti_lit_{_i}",use_container_width=True):
                            st.session_state["kw_literatur_tesis_s2"]=_r.get('Judul','') or _kw
                            st.session_state["kandidat_literatur_tesis_s2"]=cari_multi_sumber(_r.get('Judul','') or _kw,12)
                            st.rerun()

        _final=st.session_state.get("referensi_final_tesis_s2",[])
        st.markdown("### ✅ Referensi Final untuk Proposal")
        if _final:
            st.dataframe(pd.DataFrame(_final),use_container_width=True,hide_index=True)
            st.caption(f"Referensi final terpilih: {len(_final)}. Proposal hanya boleh memakai sumber final ini atau sumber primer yang telah diverifikasi.")
        else:
            st.info("Belum ada referensi final.")

        st.markdown("### 🕌 Ayat, Hadis & Tafsir")
        st.caption("AI hanya merekomendasikan kebutuhan/tema. Teks ayat, terjemahan, hadis, tafsir, edisi dan halaman tidak boleh ditebak dan harus diverifikasi sebelum menjadi sumber final.")
        if _judul_lit and _ped_teks and st.button("🕌 Analisis Perlu/Tidaknya Landasan Ayat",key="analisis_ayat_lit_s2",use_container_width=True):
            _pa=f"""Nilai apakah proposal berjudul berikut MEMERLUKAN landasan Al-Qur'an/Hadis/Tafsir secara substantif.
JUDUL: {_judul_lit}
PEDOMAN: {_ped_analisis}
Jika tidak perlu, katakan TIDAK PERLU dan alasannya. Jika perlu, sebutkan tema/konsep ayat yang perlu dicari dan bagian proposal tempat landasan itu relevan. Jangan menulis nomor ayat, teks Arab, terjemahan, hadis, tafsir, jilid atau halaman kecuali tersedia dalam sumber terverifikasi yang diberikan."""
            _ha=hasil_ai_teks(panggil_gemini(_pa,0.1))
            if _ha: st.session_state["rekomendasi_ayat_tesis_s2"]=_ha
        if st.session_state.get("rekomendasi_ayat_tesis_s2"):
            st.markdown(st.session_state["rekomendasi_ayat_tesis_s2"])

    # No.1 dan Proposal memakai workspace khusus di atas.
    # Generator generik hanya tampil pada submenu 3-12.
    if submenu_s2 not in [
        "💡 Pencarian Ide & Pengajuan Judul 🌟",
        "🔎 Literatur & Penelitian Terdahulu",
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
