# HajaArabic (VBA Terbilang Arab) 🚀

> Fungsi Terbilang Arab untuk Microsoft Excel (Kompatibel 32-bit & 64-bit).

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://github.com/Rida-Haja/VbaTerbilangArab/blob/main/LICENSE)
[![VBA](https://img.shields.io/badge/Language-VBA-blue.svg)](https://learn.microsoft.com/en-us/office/vba/api/overview/)
[![Excel](https://img.shields.io/badge/Platform-Excel-green.svg)](https://www.microsoft.com/en-us/microsoft-365/excel)
![Architecture](https://img.shields.io/badge/Arch-32--bit%20%7C%2064--bit-orange.svg)

---

**HajaArabic** adalah engine konverter angka ke teks bahasa Arab (*Tafqit*) berbasis VBA dengan akurasi tata bahasa (**Nahwu**). / *A high-precision VBA-based Arabic number-to-words (Tafqit) engine with advanced grammatical accuracy.*

> **📢 Catatan:** Kode ini dikembangkan dengan fokus pada presisi linguistik. Harap lakukan verifikasi ulang untuk penggunaan pada dokumen keuangan yang bersifat kritis. *Feel free to contribute!*
---

## 🌟 Fitur Utama / Key Features

| Fitur / Feature | Deskripsi (ID) | Description (EN) |
| :--- | :--- | :--- |
| **Nahwu Engine** | Otomasi Gender, I’rab, dan Idhafah. | Automated Gender, Case (I’rab), and Idhafah rules. |
| **Ism Manqus** | Logika dinamis untuk angka 8 (`ثماني`). | Dynamic morphology for the number 8. |
| **Digit** | Mendukung hingga 1000 digit. | Supports up to 1000 digits. |
| **Scale** | Menggunakan Skala Panjang dan Skala Pendek. | Supports both Long and Short Scale systems. |
| **3 Styles** | Modern, Klasik, & Sastra (Kecil-ke-Besar). | Modern, Classic, & Literary styles. |
| **Vocalized** | Dukungan Harakat otomatis. | Optional automatic vowelization (Harakat). |

---

## 🔧 Instalasi / Installation

1. **ID:** Buka Excel → ALT + F11 → Insert > Module → Paste kode.  
   **EN:** Open Excel → ALT + F11 → Insert > Module → Paste the code.

2. **ID:** Simpan file sebagai Excel Macro-Enabled Workbook (.xlsm).  
   **EN:** Save the file as Excel Macro-Enabled Workbook (.xlsm).

3. **ID:** Jalankan makro RegisterArabFunctions (opsional) untuk deskripsi fungsi.  
   **EN:** Run the RegisterArabFunctions macro (optional) for function descriptions.
---

## 📌 Sintaks Fungsi / Function Syntax

`=TERBILANG_ARAB(Angka; [Mode]; [Gender]; [I'rab]; [Gaya]; [Parameter]; [Harakat]; [isIdhafah])`

### Referensi Parameter / Parameter Reference

* **Angka (Number):** Input numerik / *Numeric input or cell reference.*
* **Mode:** `umum`, `urutan` (ordinal), `eja`, `uang` (currency), `benda` (unit).
* **Gender:** `m` (Muzakkar), `f` (Muannas).
* **I'rab:** `u` (Marfu’), `a` (Mansub/Majrur).
* **Gaya (Style):** `Modern`, `Klasik` (Traditional), `Sastra` (Small-to-Large).
* **Parameter:** Kode negara (`id`, `sa`, dll) atau ID Benda. / *Country code or Unit ID.*
* **Harakat:** `TRUE` / `FALSE`.
* **isIdhafah:** `TRUE` (Peluruhan Nun / *Nun deletion logic*).

---

## 💡 Contoh Penggunaan / Examples

| Konteks / Context | Rumus / Formula | Hasil / Output |
| :--- | :--- | :--- |
| **Currency (IDR)** | `=TERBILANG_ARAB(1250;"uang";"m";;"";"id";TRUE)` | أَلْف وَ مئتان وَ خَمْسُونَ رُوْبِيَّة |
| **Ordinal (Bab 5)** | `=TERBILANG_ARAB(5; "urutan"; "m")` | الفصل الخامس |
| **Percentage** | `=TERBILANG_ARAB(100; "benda"; ; ; ; 51)` | مئة في المائة |
| **Advanced 8 (f)** | `=TERBILANG_ARAB(8; "umum"; "f"; "u"; ""; ; ; FALSE)` | ثمانٍ |

---
<img width="892" height="221" alt="sshot-6" src="https://github.com/user-attachments/assets/e3ba3dae-8989-4ba6-a39c-a03190cd5d6a" />
  
## 📦 ID Satuan Benda / Unit IDs (Mode: "benda")

| ID | Kategori / Category | Unit |
| :--- | :--- | :--- |
| **1-3** | Jarak / Distance | CM, Meter, KM |
| **10-13** | Waktu / Time | Hour, Day, Month, Year |
| **20-22** | Sosial / Literacy | Person, Book, Page |
| **50-53** | Sains / Science | Degree, Percent, Watt, GB |

---

## ⚠️ Batasan / Limitations
* **ID:** Angka ordinal (urutan) hanya didukung 1–999
* **EN:** Ordinal numbers supported for 1–999 only.

---

## 👨‍💻 Author & License
* **Email**: [ridahaja@gmail.com](mailto:ridahaja@gmail.com)

---
**© 2026 Rida Rahman DH 96-02** Licensed under the [MIT License](https://github.com/Rida-Haja/VbaTerbilangArab/blob/main/LICENSE).
