# HajaArabic (VBA Terbilang Arab) 🚀

**HajaArabic** adalah fungsi kustom (UDF) berbasis VBA untuk Microsoft Excel yang dirancang untuk mengubah angka numerik menjadi teks terbilang dalam bahasa Arab. Engine ini dibangun dengan fokus pada akurasi tata bahasa (Nahwu) dan kemampuan menangani angka "monster" hingga skala **Decilyar** (66 digit).

> **Jujur saja:** Kode ini masih dalam tahap pengembangan. Walaupun sudah melalui berbagai uji ekstrem hingga 65+ digit, harap cek kembali untuk penggunaan pada dokumen keuangan yang bersifat kritis. *Feel free* untuk melaporkan bug jika ditemukan!

---

### ✨ Fitur Unggulan

* **Paham Nahwu:** Mengatur otomatis aturan gender (*Adad Ma'dud*) dan perubahan akhiran kata (*I'rab* Marfu, Mansub, Majrur).
* **Skala Astronomis:** Mendukung konversi hingga 66 digit (**Decilyar**).
* **3 Gaya Penulisan:** Pilih gaya **Modern** (ejaan populer), **Klasik** (standar literatur), atau **Sastra** (urutan angka dari satuan ke ribuan).
* **Unit Universal:** Mendukung mata uang (Rupiah, Riyal, Dinar, dll), satuan fisik (jarak, berat, waktu), hingga angka urutan (Ordinal).
* **Unicode Safe:** Teks Arab menggunakan karakter Unicode murni, dijamin rapi dan tidak korup di berbagai perangkat.

---

### 🚀 Cara Pakai

#### 1. Fungsi Utama: `=TERBILANG_ARAB()`
Gunakan rumus berikut di sel Excel Anda:
`=TERBILANG_ARAB(angka; [mode]; [gender]; [irab]; [gaya]; [parameter]; [harakat])`

* **Mode:** `umum` (default), `uang`, `eja`, `urutan`, `benda`.
* **Gender:** `m` (Muzakkar/Laki-laki), `f` (Muannas/Perempuan).
* **I'rab:** `u` (Marfu/Subjek), `a` (Mansub/Objek), `i` (Majrur/Kepemilikan).
* **Gaya:** `modern`, `klasik`, `sastra`.

#### 2. Contoh Penggunaan
| Rumus | Hasil (Output) |
| :--- | :--- |
| `=TERBILANG_ARAB(110; "umum"; "m"; "u")` | مِائَةٌ وَعَشَرَةٌ |
| `=TERBILANG_ARAB(1500; "uang"; ; ; ; "id"; TRUE)` | أَلْفٌ وَخَمْسُ مِائَةِ رُوْبِيَّةٍ |
| `=TERBILANG_ARAB(5; "urutan"; "m")` | الفَصْلُ الخَامِسُ (Kelas ke-5) |

---

### 📦 Tabel Referensi Parameter

#### A. Kode Mata Uang (Mode: `uang`)
Masukkan kode berikut pada parameter `parameter_negara`:

| Kode | Mata Uang | Satuan Kecil | Gender Utama |
| :--- | :--- | :--- | :--- |
| **id** | Rupiah (روبية) | Sen (سن) | Muannas |
| **sa** | Riyal (ريال) | Halalah (هللة) | Muzakkar |
| **kw** | Dinar (دينار) | Fils (فلس) | Muzakkar |
| **ae** | Dirham (درهم) | Fils (فلس) | Muzakkar |
| **my** | Ringgit (رينجيت) | Sen (سن) | Muzakkar |

#### B. ID Satuan Benda (Mode: `benda`)
Masukkan ID angka berikut pada parameter `parameter_negara`:

| ID | Satuan (Benda/Fisik) | ID | Satuan (Waktu) |
| :--- | :--- | :--- | :--- |
| **21** | Buku (Kitab) | **10** | Jam (Sa'ah) |
| **22** | Halaman (Safhah) | **11** | Hari (Yaum) |
| **2** | Meter (Metr) | **12** | Bulan (Syahr) |
| **5** | Berat (KG) | **13** | Tahun (Sanah) |
| **51** | Persen (Fil Miah) | **20** | Orang (Syakhsh) |

---

### 🌍 Multilingual Documentation

* **[ID] Bahasa Indonesia:** HajaArabic hadir sebagai solusi konversi angka ke teks Arab di Excel. Engine ini otomatis menyesuaikan tata bahasa, mulai dari perubahan gender angka 3-10 sampai akhiran kalimat yang kompleks.
* **[EN] English:** HajaArabic is a robust VBA solution for proper Arabic numeral representation. It handles complex gender rules and grammatical endings (I'rab) automatically, supporting astronomical scales up to 10^63.
* **[AR] العربية:** **HajaArabic** هي أداة قوية لتحويل الأرقام إلى كلمات عربية فصحى مع مراعاة دقيقة لقواعد النحو. يدعم المحرك الأرقام الضخمة والعملات المختلفة، مع إمكانية التبديل بين الأنماط الحديثة والكلاسيكية بكل سهولة.

---

### 🛠 Cara Instalasi

1.  Buka Microsoft Excel, tekan `ALT + F11` untuk membuka VBA Editor.
2.  Pilih menu `Insert` > `Module`.
3.  *Copy-paste* seluruh kode dari file `Module1.bas` ke dalam modul tersebut.
4.  Simpan file Excel Anda dengan format **Excel Macro-Enabled Workbook (.xlsm)**.

---

### 📜 Lisensi & Penulis
* **Penulis:** Rida Rahman DH 96-02
* **GitHub:** [RidhaHaja](https://github.com/RidhaHaja)
* **Lisensi:** MIT (Bebas digunakan dan dimodifikasi untuk tujuan apa pun).

---
*Dibuat dengan dedikasi untuk memudahkan komputasi angka dalam bahasa Arab.*
