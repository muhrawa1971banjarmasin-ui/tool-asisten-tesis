PERINTAH INDUK PENGEMBANGAN
ASISTEN AKADEMIK AI

PENTING:
File Python yang saya berikan adalah MASTER aplikasi.
JANGAN membuat aplikasi baru dari nol.
JANGAN menyederhanakan aplikasi.
JANGAN menghapus fitur yang sudah ada.
JANGAN mengganti fungsi yang sudah terbukti bekerja.

==================================================
A. ATURAN UTAMA
==================================================

1. Gunakan file Python terakhir yang saya berikan sebagai MASTER.

2. Sebelum mengubah kode:
   - baca seluruh kode;
   - identifikasi fungsi yang sudah bekerja;
   - identifikasi fungsi yang bermasalah;
   - ubah HANYA bagian yang diperlukan.

3. Semua menu, fungsi, tombol, dan fitur yang sudah bekerja
   WAJIB dipertahankan.

4. Jangan mengganti nama fungsi yang sudah digunakan bagian lain,
   kecuali semua pemanggilnya diperbarui dengan aman.

5. Jangan membuat fungsi duplikat dengan fungsi yang sudah ada.

6. Jangan membuat ulang parser referensi apabila parser lama
   masih bekerja.

7. Jangan membuat ulang sistem Library jika Library lama masih
   bekerja.

8. Setelah perubahan, periksa hubungan antarfitur agar tidak ada
   tombol mati atau variabel session_state yang terputus.

==================================================
B. SATU DOKUMEN AKTIF
==================================================

Gunakan SATU dokumen aktif.

Dokumen hanya diunggah SATU KALI.

Simpan dokumen aktif ke session_state, misalnya:

st.session_state["dokumen_aktif_bytes"]
st.session_state["dokumen_aktif_nama"]
st.session_state["dokumen_aktif_tipe"]

Semua fitur berikut WAJIB menggunakan dokumen aktif yang sama:

- Unggah Referensi
- Library
- Pakai di Naskah
- Audit Sitasi
- Pemeriksaan Footnote
- Validasi Referensi
- Sinkronisasi Daftar Pustaka

JANGAN membuat file_uploader kedua untuk dokumen yang sama.

Jika dokumen belum diunggah, tampilkan:

"Silakan unggah Dokumen Aktif terlebih dahulu."

==================================================
C. PARSER REFERENSI DIKUNCI
==================================================

Parser daftar pustaka DOCX yang sudah berhasil membaca
13 referensi dari proposal pengujian adalah PARSER MASTER.

JANGAN mengubah algoritma parser tersebut.

Prinsip parser:

SATU PARAGRAF DAFTAR PUSTAKA WORD = SATU REFERENSI.

Baris lanjutan DOI, URL, atau ISBN boleh digabungkan dengan
referensi sebelumnya.

JANGAN memecah referensi berdasarkan banyaknya tahun,
nama penulis, DOI, atau tanda titik.

Sebelum validasi tampilkan:

"Referensi terdeteksi: X"

Untuk proposal pengujian yang telah digunakan,
hasil yang diharapkan = 13.

==================================================
D. LIBRARY REFERENSI PUSAT
==================================================

Gunakan SATU sumber data Library.

Contoh:

st.session_state["library_referensi"]

JANGAN membuat:
bank_referensi,
references,
referensi_library,
library lain
yang menyimpan data berbeda.

Semua menu membaca Library yang sama.

Alur:

Dokumen
↓
Ekstraksi Referensi
↓
Validasi
↓
Library
↓
Pakai di Naskah
↓
Audit Sitasi
↓
Ekspor

Setiap referensi memiliki minimal:

id
judul
penulis
tahun
doi
isbn
jenis
sumber
status_validasi
kekuatan_validasi
alasan_validasi
raw_reference

==================================================
E. VALIDASI REFERENSI
==================================================

Validasi bibliografi tidak boleh bergantung sepenuhnya pada AI.

Prioritas:

1. DOI exact match
2. Crossref
3. ISBN
4. Judul
5. Penulis + tahun

Gemini digunakan untuk membantu ekstraksi atau analisis semantik,
BUKAN satu-satunya validator bibliografi.

Status:

✅ Terverifikasi
⚠️ Perlu Verifikasi
❌ Tidak Cocok

Tambahkan:

KEKUATAN VALIDASI

- Sangat Kuat
- Kuat
- Sedang
- Lemah

Contoh:

DOI + judul + penulis + tahun cocok
= Sangat Kuat.

Judul + penulis + tahun cocok
= Kuat.

Judul sebagian cocok
= Sedang.

Data sangat terbatas
= Lemah.

Tampilkan juga alasan penilaian.

==================================================
F. LIBRARY TIDAK BOLEH HILANG SAAT RERUN
==================================================

Jangan menginisialisasi ulang Library pada setiap rerun.

Gunakan pola:

if "library_referensi" not in st.session_state:
    st.session_state["library_referensi"] = []

JANGAN gunakan:

st.session_state["library_referensi"] = []

di luar kondisi inisialisasi.

Hapus referensi hanya jika pengguna menekan:

🗑️ Hapus

Kosongkan seluruh Library hanya setelah:

🗑️ Kosongkan Library
→ konfirmasi
→ Ya, kosongkan.

==================================================
G. PROTEKSI WORD MASTER
==================================================

FILE WORD YANG DIUNGGAH ADALAH MASTER.

Jangan membangun ulang dokumen dengan python-docx hanya untuk
membaca atau memvalidasi referensi.

Pada proses baca/audit:

JANGAN mengubah:
- judul
- isi naskah
- paragraf
- font
- ukuran font
- bold/italic
- spasi
- indentasi
- margin
- tabel
- gambar
- caption
- header/footer
- nomor halaman
- section
- page break
- orientasi halaman
- penomoran
- layout.

Jangan memperbaiki typo pada proses referensi.

Isi naskah hanya boleh diedit melalui menu
"Penyunting Akademik AI".

==================================================
H. FOOTNOTE
==================================================

Aplikasi harus membaca:

1. True Word Footnote
2. Footnote/catatan manual

Setelah dokumen aktif diunggah:

- baca semua footnote otomatis;
- tampilkan jumlahnya;
- cocokkan dengan Library;
- tampilkan status kecocokan.

Tidak perlu upload Word lagi.

Untuk Chicago Notes & Bibliography:

Footnote:
Times New Roman 10 pt
spasi 1.0
nomor superscript.

Sumber yang berulang menggunakan shortened note jika sesuai.

Perubahan footnote dilakukan HANYA setelah pengguna menekan:

"TERAPKAN PERUBAHAN"

==================================================
I. PAKAI DI NASKAH
==================================================

Menu ini TIDAK memiliki upload ulang.

Gunakan Dokumen Aktif.

Alur:

Dokumen Aktif
→ baca footnote/sitasi
→ cocokkan Library
→ tampilkan hasil
→ preview
→ pengguna menyetujui
→ terapkan.

Sediakan:

👁️ Pratinjau Perubahan
✅ Terapkan ke Seluruh Naskah

==================================================
J. AUDIT SITASI
==================================================

Audit Sitasi juga TIDAK memiliki upload ulang.

Gunakan Dokumen Aktif.

Periksa:

- sitasi ada tetapi sumber tidak ada di Library;
- sumber ada di Library tetapi tidak disitasi;
- footnote tidak cocok;
- DOI/ISBN bermasalah;
- format sitasi tidak sesuai;
- kemungkinan duplikasi referensi.

Audit tidak boleh mengubah dokumen.

==================================================
K. ERROR GEMINI 429
==================================================

Jangan tampilkan JSON/error teknis panjang kepada pengguna.

Jika Gemini mengembalikan HTTP 429, tampilkan:

"⚠️ Kuota AI sementara habis.
Fitur non-AI tetap dapat digunakan.
Silakan gunakan analisis AI kembali setelah kuota tersedia."

Aplikasi TIDAK BOLEH berhenti.

Crossref, Library, parser, audit metadata,
dan fungsi non-AI tetap berjalan.

==================================================
L. GAYA SITASI
==================================================

Pertahankan pilihan:

- Pertahankan format naskah asli
- Pedoman Kampus
- Chicago Notes & Bibliography
- Chicago Author-Date
- APA 7
- MLA
- Harvard
- IEEE
- Vancouver
- AMA
- Turabian
- OSCOLA
- ACS
- CSE
- APSA
- Format Kustom
- Ikuti Template/Author Guidelines Jurnal

Jangan menganggap Zotero/Mendeley sebagai gaya sitasi.

==================================================
M. REFERENCE MANAGER
==================================================

Pertahankan dukungan ekspor untuk:

- Zotero
- Mendeley
- EndNote
- RefWorks
- Paperpile
- Citavi
- JabRef

Format ekspor:

- RIS
- BibTeX
- EndNote Tagged
- CSV metadata
- daftar pustaka terformat.

==================================================
N. PENGUJIAN WAJIB SEBELUM MEMBERIKAN FILE
==================================================

Jangan hanya melakukan syntax/compile test.

Lakukan pemeriksaan logika minimal:

TEST 1
Aplikasi dapat dijalankan.

TEST 2
Upload dokumen hanya satu kali.

TEST 3
Proposal pengujian menghasilkan 13 referensi.

TEST 4
13 referensi dapat masuk ke Library.

TEST 5
Pindah tab/menu tidak mengosongkan Library.

TEST 6
Pakai di Naskah membaca Library yang sama.

TEST 7
Audit Sitasi menggunakan dokumen yang sama tanpa upload ulang.

TEST 8
Tombol-tombol utama memiliki aksi.

TEST 9
File Word MASTER tidak berubah selama proses baca/validasi.

TEST 10
Error Gemini tidak mematikan fungsi non-AI.

Jika salah satu tes gagal:
JANGAN menyebut file FINAL.

==================================================
O. ATURAN OUTPUT
==================================================

Berikan SATU file Python lengkap.

JANGAN memberikan patch terpisah.
JANGAN meminta pengguna menempel kode ke beberapa tempat.
JANGAN menghapus fitur lama untuk memperbaiki fitur baru.

Nama versi baru harus berbeda agar MASTER lama tetap aman.

==================================================
PERUBAHAN YANG DIMINTA SAAT INI
==================================================

[TULIS PERUBAHAN BARU DI SINI]

Semua bagian lain yang sudah bekerja HARUS tetap dipertahankan.
