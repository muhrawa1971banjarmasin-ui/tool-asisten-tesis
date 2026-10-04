
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
        "🎓 Perkuliahan",
        "🔎 Analisis Karya Akademik",
        "🎓 Skripsi S1",
        "🎓 Tesis S2",
        "🎓 Disertasi S3",
        "📚 Literatur & Referensi",
        "🏦 Bank Literatur",
        "✨ Penyunting Akademik AI",
        "📝 Jurnal Akademik",
        "❤️ Donasi & Akses",
        "⚙️ Admin",
    ],
)

st.sidebar.divider()
st.sidebar.caption("Akademia AI • ruang kerja akademik terpadu")

# ============================================================
# PERKULIAHAN
# ============================================================
if menu == "🎓 Perkuliahan":
    st.header("🎓 Perkuliahan")
    st.caption("RPS, materi, tugas kuliah, OBE, paper/makalah, studi kasus, review, presentasi dan artikel jurnal tugas kuliah.")

    if "mata_kuliah_tambahan" not in st.session_state:
        st.session_state.mata_kuliah_tambahan = []
    if "form_tambah_mk" not in st.session_state:
        st.session_state.form_tambah_mk = False

    mk_awal = ["Seminar Proposal Tesis", "Kepemimpinan dan Supervisi PAI"]
    daftar_mk = mk_awal + [x for x in st.session_state.mata_kuliah_tambahan if x not in mk_awal]
    mata_kuliah = st.selectbox("Pilih Mata Kuliah", daftar_mk, key="mk_aktif_final")

    if st.button("➕ Tambah Mata Kuliah", use_container_width=True, key="buka_form_mk"):
        st.session_state.form_tambah_mk = True

    if st.session_state.form_tambah_mk:
        with st.form("form_tambah_mata_kuliah", clear_on_submit=True):
            nama_mk = st.text_input("Nama Mata Kuliah")
            simpan_mk = st.form_submit_button("💾 Simpan Mata Kuliah", use_container_width=True)
        if simpan_mk:
            nama_baru = nama_mk.strip()
            if not nama_baru:
                st.warning("Masukkan nama mata kuliah.")
            elif nama_baru.lower() in [x.lower() for x in daftar_mk]:
                st.warning("Mata kuliah sudah ada.")
            else:
                st.session_state.mata_kuliah_tambahan.append(nama_baru)
                st.session_state.form_tambah_mk = False
                st.rerun()

    st.success(f"Mata kuliah aktif: {mata_kuliah}")
    bagian = st.tabs(["📝 Tugas Saya", "📄 RPS / Modul", "📚 Materi", "📌 Catatan Dosen", "📖 Referensi"])

    with bagian[0]:
        jenis = st.selectbox("Jenis Tugas", ["Makalah", "Paper", "Resume", "Review Jurnal", "Review Buku", "Critical Review", "Studi Kasus", "Laporan", "Mini Riset", "Presentasi", "Artikel Jurnal Tugas Kuliah"], key="jenis_tugas_final")
        judul = st.text_input("Judul / Tema Tugas", key="judul_tugas_final")
        files = st.file_uploader("Unggah instruksi atau bahan dosen bila ada", type=["pdf","docx","txt"], accept_multiple_files=True, key="file_tugas_final")
        bahan = ""
        for f in files or []:
            t = ekstrak_teks(f)
            if t and not t.startswith("ERROR:"):
                bahan += f"\n\nFILE: {f.name}\n{t}"
        if st.button("🔎 Cari Referensi Akademik", key="cari_ref_tugas", use_container_width=True):
            if judul.strip():
                with st.spinner("Mencari metadata referensi..."):
                    st.session_state["hasil_ref_tugas_final"] = cari_multi_sumber(judul, 10)
            else:
                st.warning("Isi judul atau tema tugas terlebih dahulu.")
        hasil_ref = st.session_state.get("hasil_ref_tugas_final", [])
        if hasil_ref:
            st.write(f"Ditemukan {len(hasil_ref)} referensi. Pilih yang akan dipakai:")
            pilihan = []
            for i, r in enumerate(hasil_ref):
                if st.checkbox(f"{r.get('Judul','Tanpa judul')} ({r.get('Tahun','')}) • {r.get('Sumber','')}", key=f"ref_tugas_final_{i}"):
                    pilihan.append(r)
            if st.button("🏦 Simpan Pilihan ke Bank Literatur", key="simpan_ref_tugas_final"):
                n = sum(1 for r in pilihan if tambah_bank_referensi(r))
                st.success(f"{n} referensi baru disimpan ke Bank Literatur.")
        if st.button("🤖 Susun Tugas dengan AI", type="primary", key="buat_tugas_final", use_container_width=True):
            if not judul.strip() and not bahan.strip():
                st.warning("Isi judul/tema atau unggah bahan tugas.")
            else:
                instruksi = f"Susun {jenis} untuk mata kuliah {mata_kuliah}. Tema/judul: {judul}. Gunakan struktur akademik yang sesuai. Untuk makalah gunakan Cover, Kata Pengantar, Daftar Isi, BAB I Pendahuluan, BAB II Pembahasan, BAB III Penutup, dan Daftar Pustaka. Jangan membuat sumber, kutipan langsung, DOI, data, atau nomor halaman yang tidak tersedia."
                panel_ai_penulisan(bahan, jenis, instruksi, st.session_state.bank_referensi, "tugas_final")
        if st.session_state.get("hasil_penulisan_ai"):
            edit = st.text_area("Hasil dapat diedit", st.session_state.hasil_penulisan_ai, height=520, key="edit_tugas_final")
            st.download_button("📥 Unduh TXT", edit.encode("utf-8"), "tugas_kuliah.txt", "text/plain", use_container_width=True)

    with bagian[1]:
        rps = st.file_uploader("Unggah RPS / modul / rancangan perkuliahan", type=["pdf","docx","txt"], accept_multiple_files=True, key="rps_final")
        if rps and st.button("🤖 Analisis RPS", key="analisis_rps_final"):
            teks="\n".join(ekstrak_teks(f) for f in rps)
            panel_ai_penulisan(teks, "Analisis RPS", f"Analisis RPS mata kuliah {mata_kuliah}. Susun materi, strategi, indikator, tugas dan keterkaitan OBE hanya berdasarkan dokumen.", st.session_state.bank_referensi, "rps_final")
    with bagian[2]:
        st.info("Materi perkuliahan dapat dikembangkan dari RPS/modul dan sumber di Bank Literatur.")
    with bagian[3]:
        st.text_area("Catatan Dosen", key=f"catatan_dosen_{mata_kuliah}", height=220)
    with bagian[4]:
        refs=st.session_state.bank_referensi
        st.dataframe(pd.DataFrame(refs), use_container_width=True, hide_index=True) if refs else st.info("Bank Literatur masih kosong.")

# ============================================================
# ANALISIS KARYA AKADEMIK
# ============================================================
elif menu == "🔎 Analisis Karya Akademik":
    st.header("🔎 Analisis Karya Akademik")
    jenis = st.selectbox("Jenis karya", ["Jurnal","Paper","Proposal","Skripsi","Tesis","Disertasi","Buku/Bab Buku","Laporan Penelitian"])
    files = st.file_uploader("Unggah karya", type=["pdf","docx","txt"], accept_multiple_files=True, key="analisis_karya_final")
    if files and st.button("🤖 Analisis Karya", type="primary"):
        teks="\n\n".join(f"DOKUMEN {i+1}: {f.name}\n{ekstrak_teks(f)}" for i,f in enumerate(files))
        panel_ai_penulisan(teks, "Analisis Karya Akademik", f"Analisis {jenis}. Jika lebih dari satu dokumen, bandingkan. Temukan masalah, gap, kelemahan, novelty yang benar-benar didukung dokumen, dan peluang pengembangan penelitian baru. Jangan mengarang.", st.session_state.bank_referensi, "analisis_karya_final")
    if st.session_state.get("hasil_penulisan_ai"):
        st.text_area("Hasil Analisis", st.session_state.hasil_penulisan_ai, height=600, key="hasil_analisis_karya_final")

# ============================================================
# RUANG PENELITIAN S1/S2/S3
# ============================================================
elif menu in ["🎓 Skripsi S1", "🎓 Tesis S2", "🎓 Disertasi S3"]:
    st.header(menu)
    if menu == "🎓 Skripsi S1":
        tahap=["Ide & Topik","Research Gap","Judul","Proposal","Metodologi","Instrumen","Analisis Data","BAB I","BAB II","BAB III","BAB IV","BAB V","Bimbingan & Revisi","Audit","Persiapan Sidang"]
        standar="skripsi S1"
    elif menu == "🎓 Tesis S2":
        tahap=["Seminar Proposal","Ide & Topik","Research Gap","Judul","Metodologi Lanjutan","Instrumen","Audio/Video","Analisis Data","BAB I","BAB II","BAB III","BAB IV","BAB V","Bimbingan & Revisi","Audit Tesis","Ujian Tesis"]
        standar="tesis S2 dengan kedalaman magister"
    else:
        tahap=["State of the Art","Theoretical Gap","Methodological Gap","Novelty Doktoral","Kontribusi Teori","Judul","Proposal","Metodologi Lanjutan","Instrumen","Analisis Data","BAB I","BAB II","BAB III","BAB IV","BAB V","Audit Disertasi","Promosi/Ujian Disertasi"]
        standar="disertasi S3 dengan standar doktoral"
    pilihan=st.selectbox("Tahap", tahap, key=f"tahap_{menu}")
    topik=st.text_area("Topik, masalah, data, atau arahan", key=f"topik_{menu}", height=150)
    files=st.file_uploader("Unggah bahan bila ada", type=["pdf","docx","txt"], accept_multiple_files=True, key=f"file_{menu}")
    bahan="\n".join(ekstrak_teks(f) for f in files or [])
    if st.button("🤖 Kerjakan dengan AI", type="primary", key=f"ai_{menu}"):
        panel_ai_penulisan(bahan, pilihan, f"Kerjakan tahap {pilihan} untuk {standar}. Arahan pengguna: {topik}. Jangan menciptakan data penelitian atau sumber.", st.session_state.bank_referensi, f"riset_{menu}")
    if st.session_state.get("hasil_penulisan_ai"):
        st.text_area("Hasil dapat diedit", st.session_state.hasil_penulisan_ai, height=600, key=f"hasil_{menu}")

# ============================================================
# LITERATUR & REFERENSI - TANPA TOMBOL HAPUS
# ============================================================
elif menu == "📚 Literatur & Referensi":
    st.header("📚 Literatur & Referensi")
    tabs=st.tabs(["🔎 Cari","📤 Unggah","🧠 Analisis","📊 Matriks","🧾 Metadata & Sitasi","✅ Konsistensi"])
    with tabs[0]:
        q=st.text_input("Kata kunci / topik", key="lit_q_final")
        if st.button("🔎 Cari Literatur", key="lit_cari_final"):
            st.session_state["lit_hasil_final"]=cari_multi_sumber(q,12) if q.strip() else []
        hasil=st.session_state.get("lit_hasil_final",[])
        for i,r in enumerate(hasil):
            with st.expander(f"{i+1}. {r.get('Judul','Tanpa judul')}"):
                st.write(format_referensi(r)); st.caption(r.get("Status",""))
                if st.button("🏦 Kirim ke Bank Literatur", key=f"kirim_bank_{i}"):
                    st.success("Disimpan ke Bank Literatur." if tambah_bank_referensi(r) else "Referensi sudah ada di Bank Literatur.")
    with tabs[1]:
        up=st.file_uploader("Unggah artikel/daftar pustaka", type=["pdf","docx","txt"], accept_multiple_files=True, key="unggah_lit_final")
        if up and st.button("📥 Analisis & Kirim ke Bank Literatur", key="proses_unggah_lit_final"):
            n,lap=unggah_referensi_ke_bank(up); st.success(f"{n} referensi baru masuk Bank Literatur."); st.dataframe(pd.DataFrame(lap),use_container_width=True,hide_index=True)
    with tabs[2]:
        st.info("Analisis literatur menggunakan sumber yang sudah berada di Bank Literatur.")
        if st.button("🤖 Analisis Literatur", key="analisis_lit_final"):
            panel_ai_penulisan("", "Analisis Literatur", "Kelompokkan tema, temukan kecenderungan, gap dan peluang penelitian berdasarkan referensi yang tersedia. Jangan mengarang isi sumber yang tidak tersedia.", st.session_state.bank_referensi, "analisis_lit_final")
    with tabs[3]:
        refs=st.session_state.bank_referensi
        if refs: st.dataframe(pd.DataFrame(refs),use_container_width=True,hide_index=True)
        else: st.info("Belum ada sumber di Bank Literatur.")
    with tabs[4]:
        refs=st.session_state.bank_referensi
        if refs:
            gaya=st.selectbox("Gaya sitasi",["Chicago Notes & Bibliography","APA 7"],key="gaya_lit_final")
            st.session_state.gaya_sitasi=gaya
            for r in refs: st.write(format_referensi(r,gaya))
            daftar="\n\n".join(format_referensi(r,gaya) for r in refs)
            st.download_button("📥 Daftar Pustaka",daftar.encode("utf-8"),"daftar_pustaka.txt","text/plain")
        else: st.info("Belum ada sumber di Bank Literatur.")
    with tabs[5]:
        naskah=st.text_area("Tempel naskah untuk pemeriksaan sitasi",height=280,key="cek_konsistensi_final")
        if naskah:
            rows=status_sitasi(naskah,st.session_state.bank_referensi)
            st.dataframe(pd.DataFrame(rows),use_container_width=True,hide_index=True) if rows else st.info("Sitasi pola author-year belum terdeteksi.")

# ============================================================
# BANK LITERATUR - TOMBOL HAPUS HANYA DI SINI
# ============================================================
elif menu == "🏦 Bank Literatur":
    st.header("🏦 Bank Literatur")
    refs=st.session_state.bank_referensi
    if not refs:
        st.info("Bank Literatur masih kosong. Cari atau unggah sumber melalui menu Literatur & Referensi.")
    else:
        st.success(f"{len(refs)} sumber tersimpan.")
        for i,r in enumerate(list(refs)):
            with st.expander(f"{i+1}. {r.get('Judul','Tanpa judul')} ({r.get('Tahun','')})"):
                st.write(format_referensi(r)); st.caption(f"{r.get('Sumber','')} • {r.get('Status','')}")
                if r.get("DOI"): st.write(f"DOI: {r.get('DOI')}")
                if st.button("🗑️ Hapus Referensi",key=f"hapus_bank_final_{i}",use_container_width=True):
                    st.session_state.bank_referensi.pop(i); st.rerun()
        st.divider()
        st.dataframe(pd.DataFrame(st.session_state.bank_referensi),use_container_width=True,hide_index=True)
        c1,c2,c3=st.columns(3)
        c1.download_button("📥 RIS",ekspor_ris(refs).encode("utf-8"),"bank_literatur.ris","application/x-research-info-systems",use_container_width=True)
        c2.download_button("📥 BibTeX",ekspor_bibtex(refs).encode("utf-8"),"bank_literatur.bib","application/x-bibtex",use_container_width=True)
        c3.download_button("📥 CSV",ekspor_csv_referensi(refs).encode("utf-8-sig"),"bank_literatur.csv","text/csv",use_container_width=True)

# ============================================================
# PENYUNTING AKADEMIK AI
# ============================================================
elif menu == "✨ Penyunting Akademik AI":
    st.header("✨ Penyunting Akademik AI")
    mode=st.selectbox("Pilih pekerjaan",["Bahasa Akademik","Typo & Ejaan","Struktur Kalimat","Konsistensi Istilah","Parafrase Ringan","Parafrase Akademik","Parafrase Mendalam","Parafrase per Paragraf","Sitasi","Orisinalitas/Kemiripan","Perapian Naskah"])
    teks=st.text_area("Teks asli",height=300,key="teks_sunting_final")
    if st.button("✨ Buat Usulan Penyuntingan",type="primary"):
        panel_ai_penulisan(teks,mode,"Tampilkan usulan sebelum→sesudah. Pertahankan makna, fakta, sitasi dan sumber. Jangan menerapkan perubahan secara diam-diam; pengguna harus dapat menilai hasilnya.",st.session_state.bank_referensi,"sunting_final")
    if st.session_state.get("hasil_penulisan_ai"):
        st.text_area("Usulan hasil — silakan periksa dan edit",st.session_state.hasil_penulisan_ai,height=500,key="hasil_sunting_final")

# ============================================================
# JURNAL AKADEMIK
# ============================================================
elif menu == "📝 Jurnal Akademik":
    st.header("📝 Jurnal Akademik")
    jenis=st.selectbox("Jenis pekerjaan",["Jurnal Tugas Kuliah","Artikel Publikasi","Artikel dari Skripsi","Artikel dari Tesis","Artikel dari Disertasi","Sesuaikan dengan Template Rumah Jurnal"])
    tema=st.text_input("Judul / tema artikel",key="tema_jurnal_final")
    files=st.file_uploader("Unggah naskah atau template rumah jurnal",type=["pdf","docx","txt"],accept_multiple_files=True,key="file_jurnal_final")
    bahan="\n".join(ekstrak_teks(f) for f in files or [])
    if st.button("🤖 Susun / Sesuaikan Artikel",type="primary"):
        panel_ai_penulisan(bahan,jenis,f"Judul/tema: {tema}. Susun atau sesuaikan artikel akademik. Jika ada template, ikuti struktur yang benar-benar terbaca dari template. Jangan mengarang aturan jurnal yang tidak tersedia.",st.session_state.bank_referensi,"jurnal_final")
    if st.session_state.get("hasil_penulisan_ai"):
        st.text_area("Hasil dapat diedit",st.session_state.hasil_penulisan_ai,height=600,key="hasil_jurnal_final")

# ============================================================
# DONASI & AKSES
# ============================================================
elif menu == "❤️ Donasi & Akses":
    st.header("❤️ Donasi & Akses")
    st.info("Pengelolaan akses, dukungan pengembangan, dan kredit AI ditempatkan di sini.")
    st.metric("Kredit AI",st.session_state.get("kredit_ai",0))

# ============================================================
# ADMIN
# ============================================================
elif menu == "⚙️ Admin":
    st.header("⚙️ Admin")
    st.write("Pengelolaan pengguna, akses AI, master data, template dan administrasi aplikasi.")
    st.metric("Karya sesi",len(st.session_state.bank_karya))
    st.metric("Sumber Bank Literatur",len(st.session_state.bank_referensi))
