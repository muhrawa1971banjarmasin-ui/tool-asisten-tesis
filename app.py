import hashlib
import io
import re
import streamlit as st
from services.literature import search_multi

def _extract_source(uploaded):
    data = uploaded.getvalue()
    name = uploaded.name.lower()
    try:
        if name.endswith(".pdf"):
            try:
                from pypdf import PdfReader
            except ImportError:
                from PyPDF2 import PdfReader
            reader = PdfReader(io.BytesIO(data))
            pages = []
            for i, page in enumerate(reader.pages, 1):
                try:
                    text = page.extract_text() or ""
                except Exception:
                    text = ""
                pages.append({"pdf_page": i, "text": text.strip()})
            return pages, None
        if name.endswith(".docx"):
            from docx import Document
            doc = Document(io.BytesIO(data))
            text = "\n".join(p.text for p in doc.paragraphs if p.text.strip())
            return [{"pdf_page": None, "text": text}], None
        if name.endswith(".txt"):
            for enc in ("utf-8", "utf-8-sig", "latin-1"):
                try:
                    return [{"pdf_page": None, "text": data.decode(enc)}], None
                except Exception:
                    pass
            return [], "Teks tidak dapat dibaca."
    except Exception as e:
        return [], str(e)
    return [], "Format belum didukung."

def _source_id(uploaded):
    return hashlib.sha256(uploaded.getvalue()).hexdigest()[:20]

def _search_book(book, query):
    terms = [x.lower() for x in re.findall(r"\w+", query, flags=re.UNICODE) if len(x) > 1]
    found = []
    for p in book.get("pages", []):
        text = p.get("text", "")
        low = text.lower()
        score = sum(low.count(t) for t in terms)
        if score:
            pos_list = [low.find(t) for t in terms if low.find(t) >= 0]
            pos = min(pos_list) if pos_list else 0
            found.append({
                "score": score,
                "pdf_page": p.get("pdf_page"),
                "excerpt": text[max(0, pos-500):min(len(text), pos+1200)].strip()
            })
    return sorted(found, key=lambda x: x["score"], reverse=True)[:10]

def _as_reference(book):
    m = book.get("metadata", {})
    return {
        "Penulis": m.get("penulis", "").strip() or "Belum terverifikasi",
        "Tahun": m.get("tahun", "").strip() or "t.t.",
        "Judul": m.get("judul", "").strip() or book.get("filename", "Kitab/Buku"),
        "Sumber": m.get("penerbit", "").strip() or "Dokumen unggahan",
        "DOI": "-",
        "Status": "🔵 Terverifikasi dari Dokumen Unggahan",
        "Jenis": m.get("jenis", "Kitab/Buku"),
        "Kota": m.get("kota", ""),
        "Jilid": m.get("jilid", ""),
        "Editor": m.get("editor", ""),
        "File": book.get("filename", ""),
        "source_id": book.get("id", "")
    }

def render():
    st.header("2. 🔎 Literatur & Penelitian Terdahulu")
    judul = st.session_state.get("judul_final_s2", "")
    if not judul:
        st.warning("Tetapkan satu judul di Submenu 1 terlebih dahulu.")
        return
    st.success("Judul final: " + judul)
    st.caption("Judul diterima otomatis dari Submenu 1. Tidak perlu diketik ulang.")

    if st.button("🔎 Cari Literatur untuk Judul Final", type="primary", key="btn_cari_lit_s2"):
        with st.spinner("Mencari Crossref, OpenAlex, dan Semantic Scholar..."):
            st.session_state["kandidat_literatur_s2"] = search_multi(judul, 15)
        st.rerun()

    refs = st.session_state.get("kandidat_literatur_s2", [])
    if refs:
        st.subheader("Kandidat Literatur")
        selected = []
        for i, r in enumerate(refs):
            k = f"ui_ref_pilih_s2_{i}"
            if st.checkbox(f"{r.get('Judul','')} ({r.get('Tahun','')})", key=k):
                selected.append(r)
            st.caption(f"{r.get('Penulis','')} | {r.get('Sumber','')} | DOI: {r.get('DOI','-')} | {r.get('Status','')}")
        if st.button("✅ Simpan sebagai Literatur Final", key="btn_simpan_lit_final_s2"):
            st.session_state["literatur_final_s2"] = selected
            st.success(f"{len(selected)} referensi disimpan sebagai Literatur Final.")

    st.divider()
    st.subheader("📚 Unggah Kitab/Buku sebagai Referensi")
    st.caption("Fitur tambahan. Pencarian jurnal di atas tetap dipertahankan.")

    if "perpustakaan_kitab_s2" not in st.session_state:
        st.session_state["perpustakaan_kitab_s2"] = {}

    uploaded = st.file_uploader(
        "Unggah kitab, tafsir, hadis, syarah, buku atau sumber primer",
        type=["pdf", "docx", "txt"], key="upload_kitab_s2"
    )

    if uploaded is not None:
        sid = _source_id(uploaded)
        old = st.session_state["perpustakaan_kitab_s2"].get(sid, {}).get("metadata", {})
        c1, c2 = st.columns(2)
        with c1:
            jenis = st.selectbox("Jenis sumber", ["Kitab Klasik","Kitab Modern","Tafsir","Hadis","Syarah Hadis","Fikih","Ushul Fikih","Akidah","Tasawuf","Pendidikan Islam","Buku Modern","Lainnya"], key=f"jenis_{sid}")
            jdl = st.text_input("Judul kitab/buku", value=old.get("judul",""), key=f"jdl_{sid}")
            penulis = st.text_input("Penulis/Pengarang", value=old.get("penulis",""), key=f"pen_{sid}")
            jilid = st.text_input("Jilid/Volume", value=old.get("jilid",""), key=f"jilid_{sid}")
        with c2:
            editor = st.text_input("Tahqiq/Editor/Penerjemah", value=old.get("editor",""), key=f"ed_{sid}")
            penerbit = st.text_input("Penerbit", value=old.get("penerbit",""), key=f"pub_{sid}")
            kota = st.text_input("Kota terbit", value=old.get("kota",""), key=f"kota_{sid}")
            tahun = st.text_input("Tahun/Edisi", value=old.get("tahun",""), key=f"th_{sid}")

        if st.button("📥 Analisis & Simpan ke Perpustakaan Kitab", key=f"simpan_{sid}", type="primary"):
            with st.spinner("Membaca sumber asli..."):
                pages, err = _extract_source(uploaded)
            if err:
                st.error("Sumber belum dapat dibaca: " + err)
            else:
                searchable = any(p.get("text","").strip() for p in pages)
                st.session_state["perpustakaan_kitab_s2"][sid] = {
                    "id": sid, "filename": uploaded.name, "pages": pages,
                    "searchable": searchable,
                    "metadata": {"jenis":jenis,"judul":jdl.strip() or uploaded.name,"penulis":penulis.strip(),"jilid":jilid.strip(),"editor":editor.strip(),"penerbit":penerbit.strip(),"kota":kota.strip(),"tahun":tahun.strip()}
                }
                if searchable:
                    st.success(f"Kitab tersimpan. {len(pages)} bagian/halaman berhasil diproses.")
                else:
                    st.warning("File tersimpan tetapi teks tidak dapat dicari. Kemungkinan PDF hasil scan. Kutipan dan halaman tidak akan dibuat otomatis.")
                st.rerun()

    library = st.session_state.get("perpustakaan_kitab_s2", {})
    if library:
        st.subheader("📖 Perpustakaan Kitab Referensi")
        ids = list(library)
        chosen = st.selectbox("Pilih sumber", ids, format_func=lambda x: (library[x].get("metadata",{}).get("judul") or library[x].get("filename")), key="pilih_kitab_s2")
        book = library[chosen]
        m = book.get("metadata", {})
        st.write(f"**Jenis:** {m.get('jenis','')}  \n**Judul:** {m.get('judul','')}  \n**Penulis:** {m.get('penulis','') or 'Belum terverifikasi'}  \n**Jilid:** {m.get('jilid','') or '-'}  \n**Tahqiq/Editor/Penerjemah:** {m.get('editor','') or '-'}  \n**Penerbit:** {m.get('penerbit','') or '-'}  \n**Kota/Tahun:** {m.get('kota','') or '-'} / {m.get('tahun','') or '-'}")
        if book.get("searchable"):
            st.success("🔵 Teks tersedia dari dokumen unggahan.")
        else:
            st.warning("🟡 Dokumen tersimpan, tetapi teks belum dapat dicari secara aman.")

        q = st.text_input("🔎 Cari tema, kata, atau potongan kalimat dalam sumber ini", key=f"q_{chosen}")
        if q:
            matches = _search_book(book, q)
            if not matches:
                st.warning("Belum ditemukan potongan yang cocok.")
            for n, hit in enumerate(matches, 1):
                label = f"Halaman PDF {hit['pdf_page']}" if hit.get("pdf_page") else "Dokumen teks"
                with st.expander(f"Hasil {n} • {label}"):
                    st.text(hit["excerpt"])
                    st.caption("Nomor halaman adalah halaman PDF. Halaman cetak tidak ditebak.")

        c1, c2 = st.columns(2)
        with c1:
            if st.button("➕ Masukkan Kitab ke Literatur Final", key=f"final_{chosen}"):
                final = list(st.session_state.get("literatur_final_s2", []))
                ref = _as_reference(book)
                if not any(r.get("source_id") == chosen for r in final):
                    final.append(ref)
                    st.session_state["literatur_final_s2"] = final
                    st.success("Kitab ditambahkan ke Literatur Final.")
                else:
                    st.info("Kitab ini sudah ada di Literatur Final.")
        with c2:
            if st.button("🗑️ Hapus dari Perpustakaan Kitab", key=f"hapus_{chosen}"):
                del st.session_state["perpustakaan_kitab_s2"][chosen]
                st.rerun()

    st.divider()
    st.metric("Literatur final", len(st.session_state.get("literatur_final_s2", [])))
