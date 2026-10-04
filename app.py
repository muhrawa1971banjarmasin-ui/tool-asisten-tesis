
import streamlit as st
import pandas as pd
import PyPDF2
from datetime import datetime
import json
import urllib.request
import urllib.error
import urllib.parse
import re
import html
import difflib
import csv
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

def analisis_dengan_gemini(teks, jenis_karya, fokus_analisis):
    """Analisis dokumen menggunakan Gemini API langsung."""
    import time

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

    payload = {
        "contents": [
            {"parts": [{"text": prompt}]}
        ],
        "generationConfig": {
            "temperature": 0.2,
            "maxOutputTokens": 8192
        }
    }

    url = (
        "https://generativelanguage.googleapis.com/v1beta/models/"
        + model_id
        + ":generateContent?key="
        + api_key
    )

    transient_codes = {429, 500, 502, 503, 504}

    for percobaan in range(3):
        request = urllib.request.Request(
            url,
            data=json.dumps(payload).encode("utf-8"),
            headers={"Content-Type": "application/json"},
            method="POST"
        )

        try:
            with urllib.request.urlopen(request, timeout=180) as response:
                data = json.loads(response.read().decode("utf-8"))

            candidates = data.get("candidates", [])
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
                    "model": model_id
                }

            return {
                "sukses": False,
                "hasil": "",
                "error": "Gemini merespons, tetapi hasil analisis kosong."
            }

        except urllib.error.HTTPError as e:
            try:
                detail = e.read().decode("utf-8")
            except Exception:
                detail = ""

            if e.code in transient_codes and percobaan < 2:
                time.sleep(3 * (percobaan + 1))
                continue

            return {
                "sukses": False,
                "hasil": "",
                "error": f"Gemini belum berhasil memproses permintaan. HTTP {e.code}. {detail}"
            }

        except urllib.error.URLError as e:
            if percobaan < 2:
                time.sleep(3 * (percobaan + 1))
                continue
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

    return {
        "sukses": False,
        "hasil": "",
        "error": "Gemini belum berhasil setelah beberapa percobaan."
    }


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
    import time
    try:
        api_key = st.secrets["GEMINI_API_KEY"]
    except Exception:
        return {"sukses": False, "hasil": "", "error": "GEMINI_API_KEY belum ditemukan di Streamlit Secrets."}

    model_id = "gemini-3.8-flash"
    payload = {
        "contents": [{"parts": [{"text": prompt[:90000]}]}],
        "generationConfig": {"temperature": temperature, "maxOutputTokens": 8192}
    }
    url = (
        "https://generativelanguage.googleapis.com/v1beta/models/"
        + model_id + ":generateContent?key=" + api_key
    )
    for percobaan in range(3):
        req = urllib.request.Request(
            url, data=json.dumps(payload).encode("utf-8"),
            headers={"Content-Type": "application/json"}, method="POST"
        )
        try:
            with urllib.request.urlopen(req, timeout=180) as response:
                data = json.loads(response.read().decode("utf-8"))
            parts = (data.get("candidates") or [{}])[0].get("content", {}).get("parts", [])
            isi = "\n".join(x.get("text","") for x in parts if isinstance(x,dict)).strip()
            if isi:
                return {"sukses": True, "hasil": isi, "error": "", "model": model_id}
            return {"sukses": False, "hasil": "", "error": "Gemini merespons tetapi hasil kosong."}
        except urllib.error.HTTPError as e:
            detail = ""
            try: detail = e.read().decode("utf-8")
            except Exception: pass
            if e.code in {429,500,502,503,504} and percobaan < 2:
                time.sleep(3*(percobaan+1)); continue
            return {"sukses":False,"hasil":"","error":f"Gemini HTTP {e.code}. {detail}"}
        except Exception as e:
            if percobaan < 2:
                time.sleep(2*(percobaan+1)); continue
            return {"sukses":False,"hasil":"","error":str(e)}
    return {"sukses":False,"hasil":"","error":"Gemini belum berhasil setelah beberapa percobaan."}


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

menu = st.sidebar.radio(
    "Menu Utama",
    [
        "🏠 Beranda",
        "📚 Perkuliahan & OBE",
        "🔬 Analisis Karya Akademik",
        "🔎 Literatur & Referensi",
        "🎓 Penelitian S1 • S2 • S3",
        "📘 Penulis Buku AI",
        "✨ Penyunting Akademik AI",
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
    # DATA MATA KULIAH TAMBAHAN
    # ========================================================
    if "mata_kuliah_tambahan" not in st.session_state:
        st.session_state.mata_kuliah_tambahan = []

    mata_kuliah_bawaan = [
        "Seminar Proposal Tesis",
        "Kepemimpinan dan Supervisi PAI"
    ]

    daftar_mata_kuliah = []

    for nama in (
        mata_kuliah_bawaan
        + st.session_state.mata_kuliah_tambahan
    ):
        if nama and nama not in daftar_mata_kuliah:
            daftar_mata_kuliah.append(nama)

    daftar_mata_kuliah.append("➕ Tambah Mata Kuliah")

    # ========================================================
    # MATA KULIAH SAYA
    # ========================================================
    st.subheader("📚 Mata Kuliah Saya")

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
            placeholder="Contoh: Filsafat Pendidikan Islam",
            key="nama_mata_kuliah_baru"
        )

        if st.button(
            "➕ Tambahkan Mata Kuliah",
            key="tambah_mata_kuliah",
            use_container_width=True
        ):

            nama_baru = nama_mk.strip()

            if not nama_baru:
                st.warning(
                    "Masukkan nama mata kuliah terlebih dahulu."
                )

            elif nama_baru in daftar_mata_kuliah:
                st.warning(
                    "Mata kuliah tersebut sudah ada."
                )

            else:
                st.session_state.mata_kuliah_tambahan.append(
                    nama_baru
                )

                st.success(
                    f"✅ Mata kuliah '{nama_baru}' berhasil ditambahkan."
                )

                st.rerun()

        st.info(
            "Mata kuliah yang ditambahkan pada tahap ini tersimpan "
            "selama sesi aplikasi aktif. Penyimpanan permanen "
            "akan menggunakan database pada tahap berikutnya."
        )

    # ========================================================
    # RUANG KERJA MATA KULIAH
    # ========================================================
    else:

        st.markdown(
            f"### 📘 {mata_kuliah}"
        )

        st.info(
            "Pilih Tugas Saya untuk membuat tugas perkuliahan. "
            "RPS dan materi kuliah dapat disimpan sebagai bahan "
            "pendukung, tetapi tidak perlu dimasukkan setiap kali "
            "membuat tugas."
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
                    "txt"
                ],
                accept_multiple_files=True,
                key="upload_tugas_mata_kuliah"
            )

            bahan_teks = ""

            if files_tugas:

                st.success(
                    f"✅ {len(files_tugas)} file berhasil dipilih."
                )

                for file in files_tugas:

                    st.write(
                        f"📄 **{file.name}** — "
                        f"{format_ukuran(file.size)}"
                    )

                    teks_file = ekstrak_teks(file)

                    if (
                        teks_file
                        and not teks_file.startswith("ERROR:")
                    ):
                        bahan_teks += (
                            f"\n\n===== {file.name} =====\n"
                            + teks_file
                        )

                    elif teks_file.startswith("ERROR:"):
                        st.warning(
                            f"{file.name}: {teks_file}"
                        )

            # =================================================
            # REFERENSI AKADEMIK
            # =================================================
            st.divider()

            st.markdown("## 🔎 Referensi Akademik")

            st.caption(
                "Akademia AI mencari metadata referensi dari "
                "Crossref, OpenAlex, Semantic Scholar dan sumber "
                "yang sudah terhubung. Referensi tidak dibuat "
                "atau dikarang oleh AI."
            )

            if "hasil_ref_tugas" not in st.session_state:
                st.session_state.hasil_ref_tugas = []

            if "signature_ref_tugas" not in st.session_state:
                st.session_state.signature_ref_tugas = ""

            kata_kunci_ref = st.text_input(
                "Kata Kunci Pencarian Referensi",
                value=judul_tugas,
                placeholder=(
                    "Contoh: kepemimpinan pendidikan Islam"
                ),
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

            if st.button(
                "🔎 Cari Referensi Ilmiah",
                key="cari_ref_tugas",
                use_container_width=True
            ):

                if not kata_kunci_ref.strip():

                    st.warning(
                        "Masukkan judul atau kata kunci terlebih dahulu."
                    )

                else:

                    with st.spinner(
                        "Mencari referensi ilmiah..."
                    ):

                        hasil_pencarian = cari_multi_sumber(
                            kata_kunci_ref,
                            jumlah_ref
                        )

                    st.session_state.hasil_ref_tugas = (
                        hasil_pencarian
                    )

                    st.session_state.signature_ref_tugas = (
                        kata_kunci_ref.strip().lower()
                    )

                    if hasil_pencarian:
                        st.success(
                            f"✅ Ditemukan "
                            f"{len(hasil_pencarian)} referensi."
                        )
                    else:
                        st.warning(
                            "Belum ditemukan referensi yang sesuai."
                        )

            referensi_dipilih = []

            hasil_ref = st.session_state.get(
                "hasil_ref_tugas",
                []
            )

            if hasil_ref:

                st.markdown(
                    "### 📚 Hasil Pencarian Referensi"
                )

                st.caption(
                    "Centang referensi yang akan digunakan "
                    "dalam tugas."
                )

                for i, ref in enumerate(
                    hasil_ref,
                    start=1
                ):

                    judul_ref = (
                        ref.get("Judul")
                        or "Tanpa judul"
                    )

                    penulis_ref = (
                        ref.get("Penulis")
                        or "Penulis belum tersedia"
                    )

                    tahun_ref = (
                        ref.get("Tahun")
                        or "Tahun belum tersedia"
                    )

                    sumber_ref = (
                        ref.get("Sumber")
                        or ""
                    )

                    status_ref = (
                        ref.get("Status")
                        or ""
                    )

                    label_ref = (
                        f"{i}. {judul_ref} "
                        f"({tahun_ref})"
                    )

                    dipilih = st.checkbox(
                        label_ref,
                        key=f"pilih_ref_tugas_{i}"
                    )

                    st.caption(
                        f"{penulis_ref} | "
                        f"{sumber_ref} | "
                        f"{status_ref}"
                    )

                    if ref.get("DOI"):
                        st.caption(
                            f"DOI: {ref.get('DOI')}"
                        )

                    if dipilih:
                        referensi_dipilih.append(
                            ref
                        )

                if referensi_dipilih:

                    if st.button(
                        "📚 Simpan Referensi Terpilih ke Bank Referensi",
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
                                "Referensi terpilih sudah ada "
                                "di Bank Referensi."
                            )

            # =================================================
            # REFERENSI DARI BANK REFERENSI
            # =================================================
            st.divider()

            st.markdown(
                "### 📖 Gunakan Bank Referensi"
            )

            refs_bank = st.session_state.bank_referensi

            if refs_bank:

                gunakan_bank = st.checkbox(
                    "Tambahkan referensi yang sudah tersimpan "
                    "di Bank Referensi",
                    value=False,
                    key="gunakan_bank_ref_tugas"
                )

                if gunakan_bank:

                    for i, ref in enumerate(
                        refs_bank,
                        start=1
                    ):

                        judul_bank = (
                            ref.get("Judul")
                            or "Tanpa judul"
                        )

                        tahun_bank = (
                            ref.get("Tahun")
                            or "n.d."
                        )

                        pilih_bank = st.checkbox(
                            f"{judul_bank} ({tahun_bank})",
                            key=f"bank_ref_tugas_{i}"
                        )

                        if pilih_bank:

                            if not any(
                                kunci_ref(x)
                                == kunci_ref(ref)
                                for x in referensi_dipilih
                            ):
                                referensi_dipilih.append(
                                    ref
                                )

            else:

                st.info(
                    "Bank Referensi masih kosong. "
                    "Anda dapat mencari referensi di atas."
                )

            # =================================================
            # PEMBUATAN TUGAS DENGAN GEMINI
            # =================================================
            st.divider()

            st.markdown("## 🤖 Susun Tugas dengan AI")

            st.caption(
                "Gemini menyusun naskah berdasarkan tema, "
                "bahan yang diunggah dan referensi yang tersedia. "
                "Referensi, DOI, data dan kutipan tidak boleh "
                "dikarang."
            )

            if jenis_tugas == "Makalah":

                instruksi_tugas = """
Susun MAKALAH AKADEMIK yang utuh dan sistematis.

STRUKTUR:
1. COVER
   - Judul makalah
   - Mata kuliah
   - Identitas mahasiswa dibuat sebagai tempat isian
   - Institusi dibuat sebagai tempat isian
   - Tahun

2. KATA PENGANTAR

3. DAFTAR ISI
   Buat struktur daftar isi yang dapat disesuaikan
   setelah naskah dipindahkan ke Word.

4. BAB I PENDAHULUAN
   A. Latar Belakang
   B. Rumusan Masalah
   C. Tujuan Penulisan

5. BAB II PEMBAHASAN
   - Susun subbab sesuai tema.
   - Jelaskan konsep secara akademik.
   - Hubungkan teori dengan konteks pembahasan.
   - Gunakan referensi yang tersedia.
   - Setiap klaim ilmiah penting harus memiliki
     dukungan referensi jika referensinya tersedia.
   - Jangan membuat sumber yang tidak tersedia.

6. BAB III PENUTUP
   A. Kesimpulan
   B. Saran

7. CATATAN KAKI jika gaya sitasi menggunakan
   Chicago Notes & Bibliography.

8. DAFTAR PUSTAKA
   Hanya masukkan sumber yang benar-benar digunakan.

ATURAN:
- Jangan mengarang referensi.
- Jangan mengarang DOI.
- Jangan mengarang nomor halaman.
- Jangan membuat kutipan langsung jika teks asli
  dan halaman sumber tidak tersedia.
- Jika halaman kutipan diperlukan tetapi belum
  terverifikasi, beri tanda:
  [halaman perlu verifikasi].
- Jika referensi tidak cukup, beri tanda:
  [PERLU REFERENSI TERVERIFIKASI].
"""

            elif jenis_tugas == "Paper":

                instruksi_tugas = """
Susun paper akademik berdasarkan tema dan bahan yang
tersedia.

Struktur utama:
- Judul
- Pendahuluan
- Permasalahan atau fokus pembahasan
- Kajian konsep/teori
- Analisis dan pembahasan
- Kesimpulan
- Daftar pustaka

Gunakan referensi yang tersedia dan jangan membuat
referensi, DOI, kutipan, data atau nomor halaman
yang tidak terverifikasi.
"""

            elif jenis_tugas == "Resume":

                instruksi_tugas = """
Susun resume akademik yang sistematis.

Tampilkan:
- Identitas/topik bahan
- Pokok-pokok pembahasan
- Konsep penting
- Penjelasan ringkas setiap pokok
- Kesimpulan

Resume harus setia pada bahan yang tersedia.
Jangan menambahkan fakta yang tidak didukung bahan.
"""

            elif jenis_tugas == "Review Jurnal":

                instruksi_tugas = """
Susun review jurnal akademik.

Struktur:
- Identitas artikel
- Latar belakang
- Tujuan penelitian
- Teori/konsep
- Metode
- Hasil/temuan
- Kekuatan artikel
- Keterbatasan artikel
- Analisis kritis
- Relevansi
- Kesimpulan

Jangan mengarang identitas, metode, hasil atau
temuan yang tidak terdapat pada bahan.
"""

            elif jenis_tugas == "Review Buku":

                instruksi_tugas = """
Susun review buku akademik.

Struktur:
- Identitas buku
- Gambaran umum
- Pokok isi
- Konsep utama
- Kekuatan
- Keterbatasan
- Analisis kritis
- Relevansi akademik
- Kesimpulan

Gunakan hanya informasi yang tersedia dalam bahan.
"""

            elif jenis_tugas == "Critical Review":

                instruksi_tugas = """
Susun critical review akademik.

Struktur:
- Identitas karya
- Ringkasan gagasan utama
- Analisis argumentasi
- Analisis teori
- Analisis metodologis jika relevan
- Kekuatan
- Kelemahan
- Posisi kritis penulis review
- Implikasi
- Kesimpulan

Pisahkan dengan jelas isi sumber dan analisis kritis.
Jangan membuat fakta atau sumber yang tidak tersedia.
"""

            elif jenis_tugas == "Studi Kasus":

                instruksi_tugas = """
Susun studi kasus akademik.

Struktur:
- Judul
- Latar belakang kasus
- Deskripsi kasus
- Identifikasi masalah
- Analisis akar masalah
- Kajian teori yang relevan
- Alternatif solusi
- Solusi yang direkomendasikan
- Rencana tindak lanjut
- Kesimpulan
- Daftar pustaka

Bedakan fakta kasus dengan analisis.
Jangan membuat data kasus yang tidak diberikan.
"""

            elif jenis_tugas == "Laporan":

                instruksi_tugas = """
Susun laporan akademik secara sistematis.

Struktur disesuaikan dengan bahan:
- Judul
- Pendahuluan
- Tujuan
- Pelaksanaan/kegiatan
- Temuan atau hasil
- Pembahasan
- Kesimpulan
- Rekomendasi
- Lampiran bila diperlukan

Jangan membuat kegiatan, data atau hasil yang
tidak tersedia.
"""

            elif jenis_tugas == "Mini Riset":

                instruksi_tugas = """
Susun naskah mini riset akademik.

Struktur:
- Judul
- Latar belakang
- Identifikasi masalah
- Rumusan masalah
- Tujuan
- Kajian teori
- Metode
- Kerangka analisis
- Hasil/temuan jika data tersedia
- Pembahasan
- Kesimpulan
- Rekomendasi
- Daftar pustaka

Jika data penelitian belum tersedia, jangan
menciptakan hasil penelitian. Buat struktur
analisis dan tandai bagian yang memerlukan data.
"""

            else:

                instruksi_tugas = """
Susun bahan presentasi akademik berdasarkan tema,
bahan dan referensi yang tersedia.

Buat:
- Judul presentasi
- Tujuan
- Latar belakang
- Pokok materi
- Konsep/teori utama
- Analisis
- Contoh atau studi kasus jika tersedia
- Kesimpulan
- Rekomendasi
- Referensi

Susun per slide dengan narasi singkat dan jelas.
Jangan membuat data atau referensi yang tidak
tersedia.
"""

            konteks_tugas = (
                f"MATA KULIAH:\n{mata_kuliah}\n\n"
                f"JENIS TUGAS:\n{jenis_tugas}\n\n"
                f"JUDUL/TEMA:\n"
                f"{judul_tugas if judul_tugas else '[belum diisi]'}"
            )

            if bahan_teks:

                konteks_tugas += (
                    "\n\nBAHAN DARI DOSEN / "
                    "BAHAN PENDUKUNG:\n"
                    + bahan_teks[:50000]
                )

            if referensi_dipilih:

                konteks_tugas += (
                    "\n\nREFERENSI TERPILIH TERSEDIA "
                    "PADA SISTEM."
                )

            if st.button(
                f"✨ Buat {jenis_tugas} dengan Gemini",
                key="buat_tugas_gemini",
                type="primary",
                use_container_width=True
            ):

                if (
                    not judul_tugas.strip()
                    and not bahan_teks.strip()
                ):

                    st.warning(
                        "Masukkan judul/tema atau unggah "
                        "bahan tugas terlebih dahulu."
                    )

                else:

                    # -----------------------------------------
                    # CARI REFERENSI OTOMATIS BILA DIPERLUKAN
                    # -----------------------------------------
                    if (
                        not referensi_dipilih
                        and judul_tugas.strip()
                        and jenis_tugas in [
                            "Makalah",
                            "Paper",
                            "Studi Kasus",
                            "Mini Riset",
                            "Presentasi"
                        ]
                    ):

                        with st.spinner(
                            "Mencari referensi ilmiah "
                            "yang relevan..."
                        ):

                            refs_otomatis = cari_multi_sumber(
                                judul_tugas,
                                10
                            )

                        refs_layak = []

                        for ref in refs_otomatis:

                            status = str(
                                ref.get("Status", "")
                            )

                            if (
                                ref.get("DOI")
                                or status.startswith("✅")
                            ):
                                refs_layak.append(ref)

                        referensi_dipilih = (
                            refs_layak[:8]
                        )

                        for ref in referensi_dipilih:
                            tambah_bank_referensi(ref)

                        if referensi_dipilih:

                            st.success(
                                f"✅ {len(referensi_dipilih)} "
                                "referensi ilmiah ditemukan "
                                "dan dihubungkan dengan tugas."
                            )

                        else:

                            st.warning(
                                "Belum ditemukan referensi "
                                "terverifikasi yang cukup. "
                                "Gemini tidak diperbolehkan "
                                "mengarang sumber."
                            )

                    panel_ai_penulisan(
                        konteks=konteks_tugas,
                        jenis_output=(
                            f"{jenis_tugas} untuk mata kuliah "
                            f"{mata_kuliah}"
                        ),
                        instruksi=instruksi_tugas,
                        referensi=referensi_dipilih,
                        key="ai_tugas_perkuliahan"
                    )

            # =================================================
            # HASIL TUGAS
            # =================================================
            hasil_tugas = st.session_state.get(
                "hasil_penulisan_ai",
                ""
            )

            if hasil_tugas:

                st.divider()

                st.markdown("## ✍️ Hasil Tugas")

                hasil_edit = st.text_area(
                    "Naskah dapat diedit",
                    value=hasil_tugas,
                    height=700,
                    key="edit_hasil_tugas"
                )

                st.session_state.naskah_aktif = (
                    hasil_edit
                )

                st.download_button(
                    "⬇️ Unduh Naskah TXT",
                    data=hasil_edit,
                    file_name=(
                        f"{jenis_tugas}_"
                        f"{mata_kuliah}.txt"
                    ),
                    mime="text/plain",
                    key="download_tugas_txt",
                    use_container_width=True
                )

                st.success(
                    "Naskah aktif sudah terhubung dengan "
                    "ruang kerja Akademia AI."
                )

        # ====================================================
        # RPS / MODUL
        # ====================================================
        elif bagian == "📄 RPS / Modul":

            st.markdown("## 📄 RPS / Modul")

            st.caption(
                "RPS atau modul disimpan sebagai bahan "
                "pendukung mata kuliah. Tidak perlu diunggah "
                "setiap kali membuat tugas."
            )

            files_rps = st.file_uploader(
                "Unggah RPS / Modul",
                type=[
                    "pdf",
                    "docx",
                    "txt"
                ],
                accept_multiple_files=True,
                key="upload_rps_mata_kuliah"
            )

            if files_rps:

                for file in files_rps:

                    st.write(
                        f"📄 **{file.name}** — "
                        f"{format_ukuran(file.size)}"
                    )

                    baru = simpan_karya(
                        file.name,
                        f"RPS / Modul - {mata_kuliah}",
                        format_ukuran(file.size)
                    )

                    if baru:
                        st.success(
                            f"✅ {file.name} disimpan "
                            "ke Bank Karya."
                        )
                    else:
                        st.info(
                            f"{file.name} sudah ada "
                            "di Bank Karya."
                        )

        # ====================================================
        # MATERI PERKULIAHAN
        # ====================================================
        elif bagian == "📚 Materi Perkuliahan":

            st.markdown(
                "## 📚 Materi Perkuliahan"
            )

            files_materi = st.file_uploader(
                "Unggah materi kuliah",
                type=[
                    "pdf",
                    "docx",
                    "txt"
                ],
                accept_multiple_files=True,
                key="upload_materi_mata_kuliah"
            )

            if files_materi:

                for file in files_materi:

                    st.write(
                        f"📄 **{file.name}** — "
                        f"{format_ukuran(file.size)}"
                    )

                    baru = simpan_karya(
                        file.name,
                        f"Materi - {mata_kuliah}",
                        format_ukuran(file.size)
                    )

                    if baru:
                        st.success(
                            f"✅ {file.name} disimpan "
                            "ke Bank Karya."
                        )
                    else:
                        st.info(
                            f"{file.name} sudah tersimpan."
                        )

        # ====================================================
        # CATATAN DOSEN
        # ====================================================
        elif bagian == "📌 Catatan Dosen":

            st.markdown("## 📌 Catatan Dosen")

            kunci_catatan = (
                "catatan_dosen_"
                + re.sub(
                    r"[^a-zA-Z0-9]+",
                    "_",
                    mata_kuliah
                ).lower()
            )

            if kunci_catatan not in st.session_state:
                st.session_state[kunci_catatan] = ""

            catatan = st.text_area(
                "Catatan penting dari dosen",
                value=st.session_state[kunci_catatan],
                height=350,
                key=f"input_{kunci_catatan}"
            )

            if st.button(
                "💾 Simpan Catatan",
                key=f"simpan_{kunci_catatan}",
                use_container_width=True
            ):

                st.session_state[kunci_catatan] = (
                    catatan
                )

                st.success(
                    "✅ Catatan dosen tersimpan "
                    "dalam sesi."
                )

        # ====================================================
        # REFERENSI MATA KULIAH
        # ====================================================
        elif bagian == "📖 Referensi Mata Kuliah":

            st.markdown(
                "## 📖 Referensi Mata Kuliah"
            )

            refs = st.session_state.bank_referensi

            if refs:

                st.success(
                    f"📚 Tersedia {len(refs)} referensi "
                    "dalam Bank Referensi."
                )

                for i, ref in enumerate(
                    refs,
                    start=1
                ):

                    st.markdown(
                        f"**{i}. "
                        f"{ref.get('Judul', 'Tanpa judul')}**"
                    )

                    st.caption(
                        f"{ref.get('Penulis', '')} | "
                        f"{ref.get('Tahun', '')} | "
                        f"{ref.get('Sumber', '')} | "
                        f"{ref.get('Status', '')}"
                    )

                    st.write(
                        format_referensi(ref)
                    )

                    if ref.get("DOI"):
                        st.caption(
                            f"DOI: {ref.get('DOI')}"
                        )

                    st.divider()

            else:

                st.info(
                    "Bank Referensi masih kosong. "
                    "Gunakan pencarian referensi pada "
                    "Tugas Saya atau menu "
                    "Literatur & Referensi."
                )
# ============================================================
# ANALISIS KARYA AKADEMIK
# ============================================================
elif menu == "🔬 Analisis Karya Akademik":

    st.header("🔬 Analisis Karya Akademik")

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
            "Analisis semantik penuh menggunakan AI Gemini langsung. "
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
                "Mesin AI dihubungkan langsung ke Gemini API untuk analisis dokumen akademik."
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
                "Dokumen sudah berhasil dibaca. Gemini AI siap "
                "menganalisis isi dokumen berdasarkan fokus yang dipilih."
            )

            if "hasil_ai_gemini" not in st.session_state:
                st.session_state.hasil_ai_gemini = ""

            if st.button(
                "🤖 Analisis dengan AI",
                type="primary",
                use_container_width=True,
                key="tombol_analisis_gemini"
            ):
                if not hasil_dokumen:
                    st.warning(
                        "Belum ada dokumen yang berhasil dibaca."
                    )
                else:
                    dokumen_ai = hasil_dokumen[0]
                    teks_ai = dokumen_ai["teks"]

                    with st.spinner(
                        "Gemini AI sedang membaca dan menganalisis dokumen..."
                    ):
                        hasil_ai = analisis_dengan_gemini(
                            teks_ai,
                            jenis,
                            fokus_analisis
                        )

                    if hasil_ai.get("sukses"):
                        st.session_state.hasil_ai_gemini = hasil_ai["hasil"]
                        st.success("✅ Analisis AI berhasil.")
                        if hasil_ai.get("model"):
                            st.caption(f"Model AI: {hasil_ai['model']}")
                    else:
                        st.session_state.hasil_ai_gemini = ""
                        st.error("❌ Analisis AI belum berhasil.")
                        st.warning(hasil_ai.get("error", "Terjadi kesalahan yang belum diketahui."))

            # =================================================
            # HASIL ANALISIS AI
            # =================================================
            st.markdown("### 📝 Hasil Analisis")

            if st.session_state.hasil_ai_gemini:

                st.success("✅ Analisis AI selesai.")

                hasil_edit = st.text_area(
                    "Hasil analisis dapat diedit sebelum diekspor",
                    value=st.session_state.hasil_ai_gemini,
                    height=600,
                    key="editor_hasil_ai_gemini"
                )

                st.session_state.hasil_ai_gemini = hasil_edit

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
                "hasil_ai_gemini",
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
# LITERATUR & REFERENSI
# ============================================================
elif menu == "🔎 Literatur & Referensi":
    st.header("🔎 Literatur, Sitasi & Library Referensi")
    c1,c2=st.columns([2,1])
    with c1: mode_ref=st.radio("Mode Referensi",["🤖 Otomatis Terverifikasi","🔍 Verifikasi Dulu","📚 Referensi Saya"],horizontal=True)
    with c2:
        gaya_list=["Chicago Notes & Bibliography","APA 7","Harvard","IEEE","MLA"]
        st.session_state.gaya_sitasi=st.selectbox("Gaya sitasi default",gaya_list,index=gaya_list.index(st.session_state.gaya_sitasi))
    st.caption("Chicago Notes & Bibliography menjadi default. Artikel jurnal tetap mengikuti gaya rumah jurnal/template yang diunggah.")
    tab_cari,tab_online,tab_upload,tab_bank,tab_pakai,tab_audit=st.tabs(["🔎 Cari Terintegrasi","🌐 Sumber Online","📤 Unggah Referensi","📚 Library","✍️ Pakai di Naskah","✅ Audit Sitasi"])

    with tab_cari:
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

    with tab_online:
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

    with tab_upload:
        uprefs=st.file_uploader("Unggah satu atau banyak PDF/DOCX/TXT referensi",type=["pdf","docx","txt"],accept_multiple_files=True,key="upload_refs")
        st.checkbox("Utamakan referensi yang saya unggah",value=True,key="prioritas_upload")
        st.caption("Otomatis: baca dokumen → cari DOI → verifikasi Crossref. Jika DOI tidak terbaca, Gemini mengekstrak metadata lalu judul diverifikasi kembali.")
        if st.button("📥 Baca, Verifikasi & Masukkan ke Library",type="primary",disabled=not bool(uprefs)):
            with st.spinner("Membaca dan memverifikasi metadata..."): n,lap=unggah_referensi_ke_bank(uprefs)
            st.session_state.laporan_upload_ref=lap; st.success(f"{n} referensi baru masuk Library.")
        if st.session_state.get("laporan_upload_ref"): st.dataframe(pd.DataFrame(st.session_state.laporan_upload_ref),use_container_width=True,hide_index=True)

    with tab_bank:
        refs=st.session_state.bank_referensi
        if refs:
            df=pd.DataFrame(refs); kol=[x for x in ["Judul","Penulis","Tahun","Jurnal","DOI","Sumber","Status"] if x in df.columns]
            st.dataframe(df[kol],use_container_width=True,hide_index=True)
            st.markdown("#### 🔄 Pengelola & Ekspor Referensi")
            manager=st.selectbox(
                "Pilih pengelola referensi",
                ["Zotero","Mendeley","EndNote","RefWorks","Paperpile","Citavi","JabRef","Lainnya / format universal"],
                key="reference_manager"
            )
            st.caption("Aplikasi menyiapkan file impor standar. Zotero, Mendeley, EndNote, RefWorks, Paperpile, Citavi, dan JabRef tetap merupakan pengelola referensi; Chicago/APA/IEEE/Harvard/MLA adalah gaya sitasi.")
            c1,c2,c3=st.columns(3)
            with c1:
                st.download_button("📥 RIS (universal)",ekspor_ris(refs).encode("utf-8"),"library_referensi.ris","application/x-research-info-systems",use_container_width=True)
            with c2:
                st.download_button("📥 BibTeX",ekspor_bibtex(refs).encode("utf-8"),"library_referensi.bib","application/x-bibtex",use_container_width=True)
            with c3:
                st.download_button("📥 EndNote Tagged",ekspor_endnote_tagged(refs).encode("utf-8"),"library_referensi.enw","text/plain",use_container_width=True)
            c4,c5=st.columns(2)
            with c4:
                st.download_button("📥 CSV Metadata",ekspor_csv_referensi(refs).encode("utf-8-sig"),"library_referensi.csv","text/csv",use_container_width=True)
            with c5:
                st.download_button("📥 Daftar Pustaka — "+st.session_state.gaya_sitasi,"\n\n".join(format_referensi(r) for r in refs).encode("utf-8"),"daftar_pustaka.txt","text/plain",use_container_width=True)
            st.info(f"Pilihan aktif: {manager}. Gunakan RIS sebagai pilihan paling umum; BibTeX cocok untuk JabRef/LaTeX, dan EndNote Tagged untuk EndNote. Metadata yang belum terverifikasi tetap ditandai agar tidak dianggap valid otomatis.")
        else: st.info("Library Referensi masih kosong.")

    with tab_pakai:
        st.subheader("✍️ Masukkan Referensi ke BAB / Naskah")
        naskah_awal=st.text_area("Tempel paragraf atau BAB",value=st.session_state.get("naskah_aktif",""),height=300,key="naskah_ref")
        refs=st.session_state.bank_referensi; opsi=[f"{i+1}. {r.get('Judul','')} ({r.get('Tahun','')})" for i,r in enumerate(refs)]
        pilihan=st.multiselect("Pilih referensi; kosong = semua yang terverifikasi",opsi,key="pilih_ref_naskah")
        dipilih=[refs[opsi.index(x)] for x in pilihan] if pilihan else [r for r in refs if str(r.get("Status","")).startswith("✅")]
        arahan=st.text_area("Arahan",placeholder="Perkuat paragraf ini dengan sumber yang benar-benar relevan.",key="arah_ref")
        if st.button("🧩 Pasang Sitasi & Footnote",type="primary",disabled=not bool(naskah_awal.strip())): panel_ai_penulisan(naskah_awal,"Pemasangan sitasi pada naskah",arahan or "Pasang sumber relevan pada klaim yang membutuhkan dukungan.",dipilih,"pasang_ref")
        if st.session_state.get("hasil_penulisan_ai"):
            h=st.text_area("Hasil — dapat diedit",st.session_state.hasil_penulisan_ai,height=600,key="hasil_ref_naskah"); st.session_state.naskah_aktif=h

    with tab_audit:
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
    st.header("🎓 Skripsi • Tesis • Disertasi")
    c1,c2,c3=st.columns(3)
    with c1:
        jenjang=st.selectbox("Jenjang",["S1 — Skripsi","S2 — Tesis","S3 — Disertasi"])
    with c2:
        metode=st.selectbox("Jenis Penelitian",[
            "Belum menentukan metode","Kuantitatif","Kualitatif","Mixed Methods",
            "R&D / Pengembangan","PTK","Studi Literatur / Library Research",
            "Systematic Literature Review (SLR)","Penelitian Evaluatif",
            "Analisis Dokumen / Analisis Isi"
        ])
    with c3:
        tahap=st.selectbox("Tahap",[
            "Rekonstruksi & Pengembangan Penelitian","Ide & Topik","Judul",
            "BAB I — Pendahuluan","BAB II — Kajian Teori","BAB III — Metode",
            "Instrumen Penelitian","BAB IV — Hasil & Pembahasan",
            "BAB V — Penutup","Naskah Lengkap","Paket Bimbingan"
        ])
    pedoman=st.file_uploader("📄 Pedoman kampus (opsional)",type=["pdf","docx","txt"],key="pedoman_kampus")
    sumber=st.file_uploader(
        "📚 Unggah tesis/skripsi/disertasi referensi, bahan, atau data",
        type=["pdf","docx","txt","csv","xlsx"],accept_multiple_files=True,key="riset_terpadu"
    )
    arah=st.text_area("Masalah, topik, arahan pembimbing, atau pengembangan yang diinginkan")
    konteks=""
    if pedoman:
        t=ekstrak_teks(pedoman)
        if not t.startswith("ERROR:"): konteks+="\nPEDOMAN KAMPUS:\n"+t
    for f in sumber or []:
        if f.name.lower().endswith((".pdf",".docx",".txt")):
            t=ekstrak_teks(f)
            if not t.startswith("ERROR:"): konteks+=f"\nSUMBER {f.name}:\n"+t
    st.info("BAB IV hanya disusun dari data penelitian nyata. Jika data belum tersedia, AI membuat struktur analisis, bukan data fiktif.")
    if st.button("🚀 Susun dengan Asisten Penelitian AI",type="primary"):
        instr=f"""Jenjang: {jenjang}
Metode: {metode}
Tahap: {tahap}
Arahan: {arah}
Jika tahap rekonstruksi, jangan menyalin penelitian lama sebagai karya baru. Analisis penelitian lama, identifikasi keterbatasan/gap, lalu kembangkan rancangan baru.
Ikuti pedoman kampus bila tersedia."""
        panel_ai_penulisan(konteks,tahap,instr,st.session_state.bank_referensi,"riset")
    if st.session_state.get("hasil_penulisan_ai"):
        edit=st.text_area("Draf penelitian — dapat diedit",st.session_state.hasil_penulisan_ai,height=650,key="edit_riset")
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
# PUBLIKASI
# ============================================================
elif menu == "📑 Publikasi Jurnal":

    st.header("📑 Asisten Publikasi Jurnal")
    target_jurnal = st.selectbox(
        "Target Publikasi",
        ["Belum ditentukan","Jurnal Nasional","SINTA 6","SINTA 5","SINTA 4","SINTA 3","SINTA 2","SINTA 1","Scopus"]
    )
    st.caption("Aplikasi membantu menyesuaikan naskah dengan scope/template jurnal target; tidak menjamin penerimaan atau peringkat jurnal.")

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
    "Asisten Akademik AI — S1 • S2 • S3 • OBE • penelitian • "
    "referensi tervalidasi • publikasi • buku • presentasi • sidang."
)
