# HajaArabic (VBA Terbilang Arab) 🚀

![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)
![VBA](https://img.shields.io/badge/Language-VBA-blue.svg)
![Excel](https://img.shields.io/badge/Platform-Excel-green.svg)

**HajaArabic** adalah engine konverter angka ke teks bahasa Arab (*Tafqit*) berbasis VBA dengan akurasi tata bahasa (**Nahwu**) tingkat tinggi. / *A high-precision VBA-based Arabic number-to-words (Tafqit) engine with advanced grammatical accuracy.*

> **📢 Catatan Pengembang:** Kode ini dikembangkan dengan fokus pada presisi linguistik. Walaupun sudah melalui berbagai uji ekstrem (65+ digit), harap lakukan verifikasi ulang untuk penggunaan pada dokumen keuangan yang bersifat kritis. *Feel free to contribute!*
---

## 🌟 Fitur Utama / Key Features

| Fitur / Feature | Deskripsi (ID) | Description (EN) |
| :--- | :--- | :--- |
| **Nahwu Engine** | Otomasi Gender, I’rab, dan Idhafah. | Automated Gender, Case (I’rab), and Idhafah rules. |
| **Ism Manqus** | Logika dinamis untuk angka 8 (`ثماني`). | Dynamic morphology for the number 8. |
| **Monster Scale** | Mendukung hingga 66 digit (Decilyar). | Supports up to 66 digits (Decilliard). |
| **Long Scale** | Menggunakan Skala Panjang Eropa. | Engineered for the European Long Scale system. |
| **3 Styles** | Modern, Klasik, & Sastra (Kecil-ke-Besar). | Modern, Classic, & Literary styles. |
| **Vocalized** | Dukungan Harakat otomatis. | Optional automatic vowelization (Harakat). |

---
<img width="892" height="221" alt="sshot-6" src="https://github.com/user-attachments/assets/e3ba3dae-8989-4ba6-a39c-a03190cd5d6a" />
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

## 📦 ID Satuan Benda / Unit IDs (Mode: "benda")

| ID | Kategori / Category | Unit |
| :--- | :--- | :--- |
| **1-3** | Jarak / Distance | CM, Meter, KM |
| **10-13** | Waktu / Time | Hour, Day, Month, Year |
| **20-22** | Sosial / Literacy | Person, Book, Page |
| **50-53** | Sains / Science | Degree, Percent, Watt, GB |

---

## ⚠️ Batasan / Limitations
* **ID:** Angka ordinal (urutan) hanya didukung 1–12. Maksimal input 66 digit.
* **EN:** Ordinal numbers supported for 1–12 only. Maximum input is 66 digits.

---

## 👨‍💻 Author & License
**Developer:** Rida Rahman  
**Email:** [RidaHaja@gmail.com](mailto:RidaHaja@gmail.com)  

© 2026 Rida Rahman. Licensed under the **MIT License**.
