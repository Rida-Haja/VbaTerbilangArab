Attribute VB_Name = "Module1"
' ==============================================================================
'  HajaArabic (VBA Terbilang Arab)
' ==============================================================================
' Copyright (c) 2026 Rida Rahman DH 96-02 — Banjarmasin, Indonesia
'  GitHub: https://github.com/RidhaHaja
'  Email : RidaHaja@gmail.com

' ==============================================================================
' LICENSE / LISENSI
' ==============================================================================
'
' [ENGLISH]
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software (the "ArabPro+"), to deal in the Software without restriction,
' including without limitation the rights to use, copy, modify, merge, publish,
' distribute, sublicense, and/or sell copies of the Software, and to permit
' persons to whom the Software is furnished to do so, subject to the following
' conditions:
'
' The above copyright notice and this permission notice shall be included in
' all copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
' THE SOFTWARE.
'
' ------------------------------------------------------------------------------
' [BAHASA INDONESIA]
' Izin diberikan secara gratis kepada siapa pun yang mendapatkan salinan
' perangkat lunak ini (ArabPro+), untuk menggunakan Perangkat Lunak tanpa
' batasan, termasuk tanpa batasan hak untuk menggunakan, menyalin, mengubah,
' menggabungkan, menerbitkan, mendistribusikan, menyublisensikan, dan/atau
' menjual salinan Perangkat Lunak, dengan syarat sebagai berikut:
'
' Pernyataan hak cipta di atas dan pernyataan izin ini harus dicantumkan dalam
' semua salinan atau bagian substansial dari Perangkat Lunak.
'
' PERANGKAT LUNAK INI DISEDIAKAN "APA ADANYA", TANPA JAMINAN APA PUN, BAIK
' TERSURAT MAUPUN TERSIRAT. DALAM HAL APA PUN, PENULIS ATAU PEMEGANG HAK CIPTA
' TIDAK BERTANGGUNG JAWAB ATAS KLAIM, KERUSAKAN, ATAU KEWAJIBAN LAINNYA DALAM
' TINDAKAN KONTRAK, KERUGIAN, ATAU LAINNYA, YANG MUNCUL DARI, DARI ATAU DALAM
' HUBUNGANNYA DENGAN PERANGKAT LUNAK.
' ==============================================================================
Option Explicit

' ==============================================================================
' DEFINISI TIPE DATA & ENUMERASI
' ==============================================================================
Public Enum GenderArab
    Muzakkar = 1    ' Laki-laki
    Muannas = 2     ' Perempuan
End Enum

Public Enum GayaArab
    Klasik = 1      ' Fusha Turats (Sambung Wawu, istilah klasik)
    Modern = 2      ' Standar Modern (Spasi jelas, istilah umum)
    Sastra = 3      ' Gaya Sastra/Puitis
End Enum

Public Enum IrabArab
    Marfu = 1       ' Nominative (Dhommah/Waw-Nun)
    Mansub = 2      ' Accusative (Fathah/Ya-Nun)
    Majrur = 3      ' Genitive (Kasrah/Ya-Nun)
End Enum

' ==============================================================================
' VARIABEL GLOBAL
' ==============================================================================
Private IsInitialized As Boolean
Private kata(0 To 12, 1 To 2) As String
Private Puluhan(2 To 9) As String
Private SatuanBesar(1 To 30) As String
Private Ordinal(1 To 12, 1 To 2) As String

' ==============================================================================
' INISIALISASI DATA
' ==============================================================================
Private Sub InitializeArabicWords()
    If IsInitialized Then Exit Sub
    
    Dim ALI As String, baa As String, taa As String, tha As String
    Dim jim As String, haa As String, kha As String, dal As String
    Dim rho As String, sin As String, syi As String, sho As String
    Dim ain As String, fha As String, kaf As String, lam As String
    Dim mim As String, NUN As String, WAW As String, YAA As String
    Dim ham As String, mar As String, maq As String
    Dim MAD As String, HZA As String, HAW As String, HAB As String
    Dim i As Integer: Dim v As Variant
    
    ' Array Unicode Dasar
    v = Array(1575, 1576, 1578, 1579, 1580, 1581, 1582, 1583, 1585, 1587, 1588, 1589, _
              1593, 1601, 1603, 1604, 1605, 1606, 1608, 1610, 1571, 1577, 1609, _
              1570, 1574, 1573, 1572)

    For i = 0 To UBound(v): v(i) = ChrW(v(i)): Next i

    ALI = v(0):  baa = v(1):  taa = v(2):  tha = v(3):  jim = v(4)
    haa = v(5):  kha = v(6):  dal = v(7):  rho = v(8):  sin = v(9)
    syi = v(10): sho = v(11): ain = v(12): fha = v(13): kaf = v(14)
    lam = v(15): mim = v(16): NUN = v(17): WAW = v(18): YAA = v(19)
    ham = v(20): mar = v(21): maq = v(22)

    ' ini tidak dipakai disini
    MAD = v(23) 'Alif Maddah / a
    HZA = v(24) ' Hamzah Nabrah / hamzah diatas ya
    HAB = v(25) ' Hamzah Bawah / i
    HAW = v(26) ' Hamzah Waw / u
    
    ' --- 0: Sifrun ---
    kata(0, 1) = sho & fha & rho
    kata(0, 2) = kata(0, 1)

    ' --- 1-9: Satuan Dasar ---
    kata(1, 1) = WAW & ALI & haa & dal           ' Wahid
    kata(2, 1) = ALI & tha & NUN & ALI & NUN     ' Ithnan
    kata(3, 1) = tha & lam & ALI & tha           ' Thalath
    kata(4, 1) = ham & rho & baa & ain           ' Arba
    kata(5, 1) = kha & mim & sin                 ' Khams
    kata(6, 1) = sin & taa                       ' Sitt
    kata(7, 1) = sin & baa & ain                 ' Sab'a
    kata(8, 1) = tha & mim & ALI & NUN           ' Thaman
    kata(9, 1) = taa & sin & ain                 ' Tis'a
    
    ' --- 10: Asyar / Asyarah ---
    kata(10, 1) = ain & syi & rho
    kata(10, 2) = kata(10, 1) & mar
    
    ' Muannas: Penambahan Ta Marbutah (mar)
    kata(1, 2) = kata(1, 1) & mar                ' Wahidah
    For i = 3 To 9
        If i <> 8 Then kata(i, 2) = kata(i, 1) & mar
    Next i
    kata(2, 2) = ALI & tha & NUN & taa & ALI & NUN ' Ithnatan
    kata(8, 2) = kata(8, 1) & YAA & mar            ' Thamaniyah

    ' --- PULUHAN (Isyrun - Tis'un) ---
    Puluhan(2) = ain & syi & rho & WAW & NUN      ' 'Isyrun
    For i = 3 To 9
        Puluhan(i) = kata(i, 1) & WAW & NUN
    Next i

' --- SATUAN BESAR (Skala Panjang Eropa Klasik) ---
    ' Skema: n-ilyun (10^6n) dan n-ilyar (10^6n+3)
    SatuanBesar(1) = ham & lam & fha                                 ' 10^3: Alfun (Ribu)
    SatuanBesar(2) = mim & lam & YAA & WAW & NUN                     ' 10^6: Milyun (Juta)
    SatuanBesar(3) = mim & lam & YAA & ALI & rho                     ' 10^9: Milyar (Miliar)
    SatuanBesar(4) = baa & lam & YAA & WAW & NUN                     ' 10^12: Bilyun (Triliun Skala Panjang)
    SatuanBesar(5) = baa & lam & YAA & ALI & rho                     ' 10^15: Bilyar (Biliard)
    SatuanBesar(6) = taa & rho & YAA & lam & YAA & WAW & NUN         ' 10^18: Trilyun (Kuadriliun)
    SatuanBesar(7) = taa & rho & YAA & lam & YAA & ALI & rho         ' 10^21: Trilyar
    SatuanBesar(8) = kaf & WAW & ALI & dal & rho & YAA & lam & YAA & WAW & NUN ' 10^24: Quadrilyun
    SatuanBesar(9) = kaf & WAW & ALI & dal & rho & YAA & lam & YAA & ALI & rho ' 10^27: Quadrilyar
    SatuanBesar(10) = kaf & WAW & YAA & NUN & taa & YAA & lam & YAA & WAW & NUN ' 10^30: Quintilyun
    SatuanBesar(11) = kaf & WAW & YAA & NUN & taa & YAA & lam & YAA & ALI & rho ' 10^33: Quintilyar
    SatuanBesar(12) = sin & kaf & sin & taa & YAA & lam & YAA & WAW & NUN       ' 10^36: Sextilyun
    SatuanBesar(13) = sin & kaf & sin & taa & YAA & lam & YAA & ALI & rho       ' 10^39: Sextilyar
    SatuanBesar(14) = sin & baa & taa & YAA & lam & YAA & WAW & NUN             ' 10^42: Septilyun
    SatuanBesar(15) = sin & baa & taa & YAA & lam & YAA & ALI & rho             ' 10^45: Septilyar
    SatuanBesar(16) = ham & kaf & taa & YAA & lam & YAA & WAW & NUN             ' 10^48: Octilyun
    SatuanBesar(17) = ham & kaf & taa & YAA & lam & YAA & ALI & rho             ' 10^51: Octilyar
    SatuanBesar(18) = NUN & WAW & NUN & YAA & lam & YAA & WAW & NUN             ' 10^54: Nonilyun
    SatuanBesar(19) = NUN & WAW & NUN & YAA & lam & YAA & ALI & rho             ' 10^57: Nonilyar
    SatuanBesar(20) = dal & YAA & sin & YAA & lam & YAA & WAW & NUN             ' 10^60: Decilyun
    SatuanBesar(21) = dal & YAA & sin & YAA & lam & YAA & ALI & rho             ' 10^63: Decilyar

    ' --- ORDINAL NUMBER (1-12) ---
    Dim Asyar As String: Asyar = ain & syi & rho
    Dim Asyarah As String: Asyarah = Asyar & mar

    ' Pola Fa'il (3,4,5,7,9)
    For i = 3 To 9
        If i <> 6 And i <> 8 Then
            Ordinal(i, 1) = Left(kata(i, 1), 1) & ALI & Mid(kata(i, 1), 2)
            Ordinal(i, 2) = Ordinal(i, 1) & mar
        End If
    Next i

    ' Irregular Ordinals
    Ordinal(1, 1) = ham & WAW & lam               ' Awwal
    Ordinal(1, 2) = ham & WAW & lam & maq         ' Ula
    Ordinal(2, 1) = tha & ALI & NUN & YAA         ' Thani
    Ordinal(2, 2) = Ordinal(2, 1) & mar           ' Thaniyah
    Ordinal(6, 1) = sin & ALI & dal & sin         ' Sadis
    Ordinal(6, 2) = Ordinal(6, 1) & mar
    Ordinal(8, 1) = tha & ALI & mim & NUN         ' Thamin
    Ordinal(8, 2) = Ordinal(8, 1) & mar
    Ordinal(10, 1) = ain & ALI & syi & rho        ' Asyir
    Ordinal(10, 2) = Ordinal(10, 1) & mar
    
    ' Compound Ordinals (11-12)
    Ordinal(11, 1) = haa & ALI & dal & YAA & " " & Asyar
    Ordinal(11, 2) = haa & ALI & dal & YAA & mar & " " & Asyarah
    Ordinal(12, 1) = Ordinal(2, 1) & " " & Asyar
    Ordinal(12, 2) = Ordinal(2, 1) & mar & " " & Asyarah

    IsInitialized = True
End Sub

' ==============================================================================
' FUNGSI UTAMA
' ==============================================================================
Private Function AngkaKeTeksArab(ByVal angka As Variant, _
                                Optional ByVal mode As Byte = 1, _
                                Optional ByVal gender As GenderArab = Muzakkar, _
                                Optional ByVal gaya As GayaArab = Modern, _
                                Optional ByVal isCurrency As Boolean = False, _
                                Optional ByVal irab As IrabArab = Marfu, _
                                Optional ByVal isTunggal As Boolean = True) As String
    
    Dim strNum As String, res As String
    Dim FinalBulat As String, FinalDes As String
    
    ' 1. Normalisasi input angka/string
    If IsNumeric(angka) Then
        ' Format agar tidak ada notasi ilmiah (E+) dan presisi penuh
        strNum = Format(angka, "##############################0.##############################")
        strNum = Trim(Replace(strNum, " ", ""))
    Else
        strNum = Trim(Replace(CStr(angka), " ", ""))
    End If
    
    If strNum = "" Then AngkaKeTeksArab = "": Exit Function
    If Len(strNum) > 66 Then
        AngkaKeTeksArab = "[Khatha': Melebihi Batas]"
        Exit Function
    End If
    
    On Error GoTo SafeExit
    
    ' Pastikan inisialisasi data sekali saja
    If Not IsInitialized Then InitializeArabicWords

    ' 2. Cek Negatif
    Dim prefix As String: prefix = ""
    If Left(strNum, 1) = "-" Then
        prefix = ChrW(1587) & ChrW(1575) & ChrW(1604) & ChrW(1576) & " " ' Salib (Negatif)
        strNum = Mid(strNum, 2)
    End If

    ' 3. Pisahkan Bilangan Bulat dan Desimal
    Dim PosKoma As Long
    PosKoma = InStr(strNum, ",")
    If PosKoma = 0 Then PosKoma = InStr(strNum, ".")
    
    Dim bulat As String, Desimal As String
    If PosKoma > 0 Then
        bulat = Left(strNum, PosKoma - 1)
        Desimal = Mid(strNum, PosKoma + 1)
        Desimal = Replace(Replace(Desimal, ".", ""), ",", "")
    Else
        bulat = strNum
        Desimal = ""
    End If
    bulat = Replace(Replace(bulat, ".", ""), ",", "")

    ' 4. Proses Berdasarkan Mode
    If mode = 2 Then
        ' --- MODE PER DIGIT ---
        res = ""
        Dim i As Long
        For i = 1 To Len(bulat)
            res = res & kata(val(Mid(bulat, i, 1)), gender) & " "
        Next i
        
        If Len(Desimal) > 0 Then
            ' Tambah kata "Fashilah" (Koma)
            res = res & ChrW(1601) & ChrW(1575) & ChrW(1589) & ChrW(1604) & ChrW(1577) & " "
            For i = 1 To Len(Desimal)
                res = res & kata(val(Mid(Desimal, i, 1)), gender) & " "
            Next i
        End If
        AngkaKeTeksArab = Trim(prefix & res)
        Exit Function
        
    Else
        ' --- MODE STANDAR / SASTRA ---
        ' A. Bagian Bulat
        If Trim(Replace(bulat, "0", "")) = "" Then
            FinalBulat = kata(0, gender) ' Sifrun
        Else
            FinalBulat = ProsesBlokRibuan(bulat, gender, gaya, irab, isTunggal)
        End If
        
        ' B. Bagian Desimal
        If Len(Desimal) > 0 Then
            Dim isAllZero As Boolean
            isAllZero = (Trim(Replace(Desimal, "0", "")) = "")
            
            If Not isAllZero Then
                Dim Pemisah As String
                ' Jika mata uang gunakan "Wa" (Dan), jika biasa gunakan "Fashilah" (Koma)
                Pemisah = IIf(isCurrency, " " & ChrW(1608) & " ", " " & ChrW(1601) & ChrW(1575) & ChrW(1589) & ChrW(1604) & ChrW(1577) & " ")
                
                If isCurrency Then
                    ' Logic khusus sen/halalah (2 digit)
                    Dim senVal As String
                    senVal = Format$(val("0." & Desimal), "0.00")
                    FinalDes = Pemisah & ConvertThreeDigitsArab(val(Right(senVal, 2)), Muannas, gaya, irab, isTunggal)
                Else
                    ' Baca digit per digit untuk desimal matematika
                    FinalDes = Pemisah
                    Dim j As Integer
                    For j = 1 To Len(Desimal)
                        FinalDes = FinalDes & kata(val(Mid(Desimal, j, 1)), Muannas) & " "
                    Next j
                End If
            End If
        End If
        
        res = Trim(prefix & FinalBulat & FinalDes)
    End If

    ' 5. Formatting Akhir (Post-Processing)
    If gaya = Klasik Then
        ' A. Rapatkan Ratusan (Contoh: Thalathu Miah -> Thalathumiah)
        res = Replace(res, " " & ChrW(1605) & ChrW(1574) & ChrW(1577), ChrW(1605) & ChrW(1574) & ChrW(1577))

        ' B. Penyesuaian Nun Dual (Hanya jika isTunggal = False)
        If isTunggal = False Then
            res = IdhafahNun(res, False)
        End If

        ' C. Rapatkan 'Wa' (?) dengan kata berikutnya
        res = Replace(res, ChrW(1608) & " ", ChrW(1608))
    End If
    
    ' Pembersihan spasi ganda dan Trim akhir
    Do While InStr(res, "  ") > 0: res = Replace(res, "  ", " "): Loop
    
    AngkaKeTeksArab = Trim(res)
    Exit Function

SafeExit:
    AngkaKeTeksArab = "[Khatha': Error Sistem]"
End Function

' ==============================================================================
' FUNGSI LOGIKA INTI
' ==============================================================================
Private Function ProsesBlokRibuan(ByVal strNum As Variant, _
                                  ByVal PilihGender As GenderArab, _
                                  ByVal gaya As GayaArab, _
                                  Optional irab As IrabArab = Marfu, _
                                  Optional isTunggal As Boolean = True) As String
    Dim Blocks() As String
    Dim wordList() As String
    Dim finalArr() As String
    Dim blockCount As Integer: blockCount = 0
    Dim i As Integer, j As Integer
    Dim WAW As String: WAW = ChrW(1608)

    ' 1. Normalisasi String Murni
    Dim cleanNum As String: cleanNum = ""
    Dim rawInput As String: rawInput = Trim(CStr(strNum))
    
    For i = 1 To Len(rawInput)
        Dim c As String: c = Mid(rawInput, i, 1)
        If c >= "0" And c <= "9" Then cleanNum = cleanNum & c
    Next i
    
    If cleanNum = "" Then Exit Function

    ' 2. Pemecahan Blok Berbasis Kelipatan 3
    ' Contoh: "11" menjadi "011" agar terbaca satu blok utuh
    Do While Len(cleanNum) Mod 3 <> 0
        cleanNum = "0" & cleanNum
    Loop

    blockCount = Len(cleanNum) / 3
    ReDim wordList(0 To blockCount - 1)

    ' 3. Konversi Tiap Blok (PENTING: i harus sampai 0)
    For i = 0 To blockCount - 1
        ' Mengambil 3 digit dari posisi paling kanan secara bertahap
        Dim startPos As Integer
        startPos = Len(cleanNum) - (i * 3) - 2
        
        Dim sBlock As String
        sBlock = Mid(cleanNum, startPos, 3)
        
        Dim vTemp As Integer: vTemp = val(sBlock)
        
        If vTemp > 0 Then
            ' Jika i=0, panggil ConvertThreeDigitsArab (untuk angka 11)
            ' Jika i>0, panggil FormatSatuanBesar (untuk Septiliyar, dst)
            If i = 0 Then
                wordList(i) = ConvertThreeDigitsArab(vTemp, PilihGender, gaya, irab, isTunggal)
            Else
                wordList(i) = FormatSatuanBesar(vTemp, i, gaya, irab, isTunggal)
            End If
        End If
    Next i

    ' 4. Penentuan Kata Hubung (Junction)
    Dim junc As String
    If gaya = Sastra Then
        junc = " "
    ElseIf gaya = Klasik Then
        junc = " " & WAW
    Else
        junc = " " & WAW & " "
    End If

    ' 5. Penyusunan Array Final (Urutan Besar ke Kecil untuk Modern/Klasik)
    Dim countVal As Integer: countVal = 0
    If gaya = Sastra Then
        For i = 0 To blockCount - 1 ' Satuan dulu
            If wordList(i) <> "" Then
                ReDim Preserve finalArr(countVal)
                finalArr(countVal) = wordList(i)
                countVal = countVal + 1
            End If
        Next i
    Else
        For i = blockCount - 1 To 0 Step -1 ' Skala besar dulu, akhiri di satuan (0)
            If wordList(i) <> "" Then
                ReDim Preserve finalArr(countVal)
                finalArr(countVal) = wordList(i)
                countVal = countVal + 1
            End If
        Next i
    End If

    ' 6. Penggabungan Akhir
    If countVal > 0 Then
        Dim res As String
        res = Join(finalArr, junc)
        
        ' Koreksi spasi untuk Gaya Klasik
        If gaya = Klasik Then res = Replace(res, WAW & " ", WAW)
        
        ' Hapus spasi ganda dan Trim
        Do While InStr(res, "  ") > 0: res = Replace(res, "  ", " "): Loop
        ProsesBlokRibuan = Trim(res)
    Else
        ProsesBlokRibuan = ""
    End If
End Function

Private Function ConvertThreeDigitsArab(ByVal num As Variant, _
                                      ByVal PilihGender As GenderArab, _
                                      ByVal gaya As GayaArab, _
                                      Optional irab As IrabArab = Marfu, _
                                      Optional isTunggal As Boolean = True) As String
    Dim s As String, WAW As String
    Dim ALI As String: ALI = ChrW(1575)
    Dim taa As String: taa = ChrW(1578)
    Dim NUN As String: NUN = ChrW(1606)
    Dim YAA As String: YAA = ChrW(1610)
    Dim maq As String: maq = ChrW(1609)
    Dim ham As String: ham = ChrW(1574) ' Hamzah di atas kursi Ya (Miah)
    Dim mar As String: mar = ChrW(1577) ' Ta Marbuthah
    
    ' Kata hubung (Waw) disesuaikan gaya
    ' Klasik:(Spasi sebelum, menempel setelahnya)
    ' Modern:(Spasi sebelum dan sesudah)
    WAW = IIf(gaya = Klasik, " " & ChrW(1608), " " & ChrW(1608) & " ")
    
    Dim v As Long
    v = CLng(num)
    If v = 0 Then Exit Function
    
    ' --- 1. RATUSAN (Mi'ah / Ma'iah) ---
    If v >= 100 Then
        Dim h As Integer: h = Int(v / 100)
        
        Select Case h
            Case 1
                ' 100: Klasik menggunakan Alif Ekstra (Ma'iah)
                If gaya = Klasik Then
                    s = ChrW(1605) & ALI & ham & mar
                Else
                    s = ChrW(1605) & ham & mar
                End If
                
            Case 2
                ' 200: Ta Marbuthah terbuka menjadi Ta Maftuhah
                Dim baseDual As String
                baseDual = IIf(gaya = Klasik, ChrW(1605) & ALI & ham & taa, ChrW(1605) & ham & taa)
                
                s = baseDual & IIf(irab = Marfu, ALI, YAA)
                ' Nun dibuang pada gaya Klasik/Sastra untuk persiapan Idofah
                If gaya = Modern Then s = s & NUN
                
            Case 8
                If gaya = Klasik Then
                    ' 800 Klasik: (Tanpa spasi, dengan Alif)
                    s = ChrW(1579) & ChrW(1605) & ChrW(1575) & NUN & ChrW(1605) & ALI & ham & mar
                Else
                    s = kata(8, 2) & " " & ChrW(1605) & ham & mar
                End If
                
            Case Else
                ' Gaya Klasik: Gabung tanpa spasi + Alif
                ' Gaya Modern: Pisah dengan spasi
                Dim suffixMiah As String
                If gaya = Klasik Then
                    suffixMiah = ChrW(1605) & ALI & ham & mar
                Else
                    suffixMiah = " " & ChrW(1605) & ham & mar
                End If
                s = kata(h, 1) & suffixMiah
        End Select
        
        v = v Mod 100
        If v > 0 Then s = s & WAW
    End If

    ' --- 2. PULUHAN & SATUAN ---
    If v > 0 Then
        ' Satuan Dasar (1-10)
        If v <= 10 Then
            If v = 8 Then
                s = s & GetThaman(PilihGender, isTunggal)
            ElseIf v = 1 Or v = 2 Then
                s = s & GetWahidIthnan(v, PilihGender, irab)
            Else
                ' Kaidah Adad Ma'dud: Gender berlawanan (3-10)
                s = s & kata(v, IIf(PilihGender = Muzakkar, 2, 1))
            End If
            
        ' Belasan (11-19)
        ElseIf v < 20 Then
            Dim unitB As Integer: unitB = v Mod 10
            Dim txtU As String
            
            If unitB = 1 Then
                ' Ahad (M) / Ihda (F)
                txtU = IIf(PilihGender = Muzakkar, ChrW(1571) & ChrW(1581) & ChrW(1583), ChrW(1573) & ChrW(1581) & ChrW(1583) & maq)
            ElseIf unitB = 2 Then
                Dim suffixDual As String: suffixDual = IIf(irab = Marfu, ALI, YAA)
                ' Angka 12: Nun selalu dibuang baik Modern/Klasik karena bersambung dengan Asyar
                If PilihGender = Muzakkar Then
                    txtU = ALI & ChrW(1579) & NUN & suffixDual
                Else
                    txtU = ALI & ChrW(1579) & NUN & taa & suffixDual
                End If
            ElseIf unitB = 8 Then
                txtU = GetThaman(PilihGender, True)
            Else
                ' Satuan berlawanan gender (13-19)
                txtU = kata(unitB, IIf(PilihGender = Muzakkar, 2, 1))
            End If
            
            ' Gabungkan dengan Asyar (Gender Asyar sesuai input)
            s = s & txtU & " " & kata(10, IIf(PilihGender = Muzakkar, 1, 2))
            
        ' Puluhan (20-99)
        Else
            Dim st As Integer: st = v Mod 10
            Dim pul As String: pul = Puluhan(Int(v / 10))
            
            ' Perubahan I'rab Puluhan (uuna -> iina)
            If irab <> Marfu Then
                pul = Replace(pul, ChrW(1608) & NUN, YAA & NUN)
            End If
            
            If st = 0 Then
                s = s & pul
            Else
                Dim txtSat As String
                If st = 8 Then
                    txtSat = GetThaman(PilihGender, True)
                ElseIf st = 1 Or st = 2 Then
                    txtSat = GetWahidIthnan(st, PilihGender, irab)
                Else
                    txtSat = kata(st, IIf(PilihGender = Muzakkar, 2, 1))
                End If
                s = s & txtSat & WAW & pul
            End If
        End If
    End If
    
    ConvertThreeDigitsArab = Trim(s)
End Function

' ==============================================================================
' FUNGSI BANTUAN (Helper)
' ==============================================================================
Private Function FormatSatuanBesar(ByVal vNum As Long, ByVal idx As Integer, _
                                   ByVal gaya As GayaArab, ByVal irab As IrabArab, _
                                   ByVal isTunggal As Boolean) As String
    Dim Cur As String
    Dim ALI As String: ALI = ChrW(1575): Dim YAA As String: YAA = ChrW(1610)
    Dim NUN As String: NUN = ChrW(1606): Dim fat As String: fat = ChrW(1611)
    Dim rho As String: rho = ChrW(1585)
    Dim fha As String: fha = ChrW(1601) ' Tambahan karakter Fha untuk Alfun
    
    ' 1. Blok Satuan/Ratusan murni
    If idx = 0 Then
        FormatSatuanBesar = ConvertThreeDigitsArab(vNum, Muzakkar, gaya, irab, isTunggal)
        Exit Function
    End If

    ' 2. Logika vNum = 1 (Seribu, Sejuta, dsb.)
    If vNum = 1 Then
        Cur = SatuanBesar(idx)

    ' 3. Logika vNum = 2 (Dua Ribu, Dua Juta, dsb.)
    ElseIf vNum = 2 Then
        Dim baseUnit As String: baseUnit = SatuanBesar(idx)
        
        ' Membangun bentuk Dual (Alf + ani / Alf + aini)
        Cur = baseUnit & IIf(irab = Marfu, ALI, YAA)
        
        If isTunggal = True Then
            Cur = Cur & NUN
        End If

    ' 4. Logika 3 Ke Atas (Jamak vs Tamyiz)
    ElseIf vNum >= 3 Then
        Dim unitFinal As String
        Dim sisa As Integer
        sisa = vNum Mod 100

        If sisa >= 3 And sisa <= 10 Then
            unitFinal = GetMa_dud(sisa, SatuanBesar(idx))
        Else
            unitFinal = SatuanBesar(idx)
            
            If gaya <> Sastra Then
                If Right(unitFinal, 1) <> ALI Then
                    If Right(unitFinal, 1) = NUN Or Right(unitFinal, 1) = rho Or Right(unitFinal, 1) = fha Then
                        ' Jika user memilih Harakat = TRUE, Anda bisa menyisipkan ChrW(1611) sebelum Alif
                        unitFinal = unitFinal & ALI
                    End If
                End If
            End If
        End If
        
        Cur = ConvertThreeDigitsArab(vNum, Muzakkar, gaya, irab, isTunggal) & " " & unitFinal
    End If
        
    FormatSatuanBesar = Cur
End Function

Private Function GetThaman(ByVal gender As GenderArab, ByVal isTunggal As Boolean) As String
    ' Penanganan angka 8 yang kompleks
    If gender = Muzakkar Then
        GetThaman = kata(8, 2) ' Thamaniyah
    Else
        ' Thaman / Thamani
        GetThaman = kata(8, 1) & IIf(isTunggal, "", ChrW(1610))
    End If
End Function

Private Function GetWahidIthnan(ByVal val As Variant, ByVal gender As GenderArab, ByVal irab As IrabArab) As String
    Dim ALI As String: ALI = ChrW(1575): Dim NUN As String: NUN = ChrW(1606)
    Dim YAA As String: YAA = ChrW(1610): Dim tha As String: tha = ChrW(1579)
    
    If val = 1 Then
        GetWahidIthnan = kata(1, gender)
    ElseIf val = 2 Then
        ' Dasar: Ithn (M) / Ithnat (F) + Akhiran Dual
        Dim base As String
        base = ALI & tha & NUN & IIf(gender = Muannas, ChrW(1578), "")
        GetWahidIthnan = base & IIf(irab = Marfu, ALI, YAA) & NUN
    End If
End Function

Private Function IdhafahNun(ByVal Teks As String, ByVal isTunggal As Boolean) As String
''' -------------------------------------------------------------------------
''' Fungsi: IdhafahNun
''' Deskripsi: Menghapus Nun pada bilangan dual (Mutsanna) jika angka
'''            tersebut tidak berdiri sendiri (isTunggal = False), sesuai
'''            dengan kaidah Hadzfun Nun lil Idhafah.
''' -------------------------------------------------------------------------
    If isTunggal = True Then IdhafahNun = Teks: Exit Function

    Dim kata() As String, i As Integer
    Dim NUN As String: NUN = ChrW(1606)
    Dim YAA As String: YAA = ChrW(1610)
    Dim WAW As String: WAW = ChrW(1608)
    Dim ALI As String: ALI = ChrW(1575)
    
    kata = Split(Teks, " ")

    For i = 0 To UBound(kata)
        Dim s As String: s = kata(i)
        
        If Right(s, 1) = NUN And Len(s) > 3 Then
            Dim isPuluhan As Boolean: isPuluhan = False
            
            ' A. Proteksi Puluhan (uuna / iina)
            ' Jika sebelum Nun adalah Waw, atau jika pola kata adalah puluhan murni
            If Mid(s, Len(s) - 1, 1) = WAW Then
                isPuluhan = True
            End If
            
            ' B. Filter Inklusi Bilangan Dual
            Dim isDual As Boolean
            isDual = (InStr(s, ALI & ChrW(1579)) > 0) Or _
                     (InStr(s, ChrW(1605) & ALI & ChrW(1574) & ChrW(1578)) > 0) Or _
                     (InStr(s, ChrW(1571) & ChrW(1604) & ChrW(1601)) > 0) Or _
                     (InStr(s, YAA & WAW) > 0) Or _
                     (InStr(s, YAA & ALI & ChrW(1585)) > 0)

            ' C. Eksekusi
            If isDual And Not isPuluhan Then
                Dim preNun As String: preNun = Mid(s, Len(s) - 1, 1)
                ' Pastikan hanya memotong jika diakhiri Alif-Nun atau Ya-Nun (Ciri Mutsanna)
                If preNun = ALI Or preNun = YAA Then
                    ' Cek tambahan: Pastikan bukan kata puluhan seperti 'Isyrina'
                    ' Isyrina mengandung akar 'Ain-Syin-Ra, bukan Ithn/Miat/Alf/Yunu/Yari
                    kata(i) = Left(s, Len(s) - 1)
                End If
            End If
        End If
    Next i
    
    IdhafahNun = Join(kata, " ")
End Function

Private Function GetMa_dud(ByVal angka As Long, ByVal satuan As String) As String
    Dim rho As String, YAA As String, ALI As String, lam As String, taa As String
    Dim dal As String, NUN As String, raa As String, haa As String, mim As String
    Dim WAW As String, baa As String, mar As String

    rho = ChrW(1585): YAA = ChrW(1610): ALI = ChrW(1575): lam = ChrW(1604): taa = ChrW(1578)
    dal = ChrW(1583): NUN = ChrW(1606): raa = ChrW(1585): haa = ChrW(1607): mim = ChrW(1605)
    WAW = ChrW(1608): baa = ChrW(1576): mar = ChrW(1577)

    Dim v As Long: v = angka Mod 100
    
    Dim sRiyal As String:  sRiyal = rho & YAA & ALI & lam
    Dim sDinar As String:  sDinar = dal & YAA & NUN & ALI & raa
    Dim sDirham As String: sDirham = dal & rho & haa & mim
    Dim sRupiah As String: sRupiah = rho & WAW & baa & YAA & mar

    Select Case True
        Case satuan = sRiyal
            GetMa_dud = IIf(v >= 3 And v <= 10, sRiyal & ALI & taa, satuan)
            
        Case satuan = sDinar
            ' Jamak Taktsir: Dananir
            GetMa_dud = IIf(v >= 3 And v <= 10, dal & NUN & ALI & NUN & YAA & raa, satuan)
            
        Case satuan = sDirham
            ' Jamak Taktsir: Darahim
            GetMa_dud = IIf(v >= 3 And v <= 10, dal & raa & ALI & haa & mim, satuan)
                
        Case satuan = sRupiah
            GetMa_dud = IIf(v >= 3 And v <= 10, rho & WAW & baa & YAA & ALI & taa, satuan)
                        
        Case satuan = ChrW(1571) & ChrW(1604) & ChrW(1601) ' Alfun (???)
            ' Jika 3-10 menjadi Aalaaf (????)
            GetMa_dud = IIf(v >= 3 And v <= 10, ChrW(1570) & ChrW(1604) & ChrW(1575) & ChrW(1601), satuan)

        Case satuan = ChrW(1605) & ChrW(1604) & ChrW(1610) & ChrW(1608) & ChrW(1606) ' Milyun (?????)
            ' Jika 3-10 menjadi Malayiin (??????)
            GetMa_dud = IIf(v >= 3 And v <= 10, ChrW(1605) & ChrW(1604) & ChrW(1575) & ChrW(1610) & ChrW(1610) & ChrW(1606), satuan)

        Case satuan = ChrW(1605) & ChrW(1604) & ChrW(1610) & ChrW(1575) & ChrW(1585) ' Milyar (?????)
            ' Jika 3-10 menjadi Milyarat (???????)
            GetMa_dud = IIf(v >= 3 And v <= 10, ChrW(1605) & ChrW(1604) & ChrW(1610) & ChrW(1575) & ChrW(1585) & ChrW(1575) & ChrW(1578), satuan)
            
        Case satuan = ChrW(1578) & ChrW(1585) & ChrW(1610) & ChrW(1604) & ChrW(1610) & ChrW(1608) & ChrW(1606) ' Trilyun (??????)
             ' Jika 3-10 menjadi Trilyunat (????????)
             GetMa_dud = IIf(v >= 3 And v <= 10, ChrW(1578) & ChrW(1585) & ChrW(1610) & ChrW(1604) & ChrW(1610) & ChrW(1608) & ChrW(1606) & ChrW(1575) & ChrW(1578), satuan)
        
        Case Else
            GetMa_dud = satuan
    End Select
End Function





'=================== BAB TAMBAHAN LAINNYA ============================================
Public Function ArabCurrency(ByVal angka As Variant, _
                             Optional ByVal KodeNegara As String = "id") As String
    ' ==================================================================================
    ' FUNGSI: Mengonversi angka ke terbilang mata uang dalam bahasa Arab
    ' PENYESUAIAN: Menambahkan IdhafahNun agar 200/2000 membuang Nun saat bertemu mata uang
    ' ==================================================================================
    
    Dim HasilUtama As String, strAngka As String
    Dim BagianBulat As String, BagianDesimal As String
    Dim PosKoma As Long
    Dim TeksMataUang As String, TeksSen As String, WAW As String
    Dim GenderUtama As GenderArab, GenderKecil As GenderArab
    
    ' kata hubung "dan" (Waw)
    WAW = " " & ChrW(1608) & " "
    KodeNegara = LCase(Trim(KodeNegara))
    
    ' --- SELECT CASE TETAP (TIDAK DIBUANG) ---
    Select Case KodeNegara
        Case "id": TeksMataUang = ChrW(1585) & ChrW(1608) & ChrW(1576) & ChrW(1610) & ChrW(1577): TeksSen = ChrW(1587) & ChrW(1606): GenderUtama = Muannas: GenderKecil = Muzakkar
        Case "sa": TeksMataUang = ChrW(1585) & ChrW(1610) & ChrW(1575) & ChrW(1604): TeksSen = ChrW(1607) & ChrW(1604) & ChrW(1604) & ChrW(1577): GenderUtama = Muzakkar: GenderKecil = Muannas
        Case "kw": TeksMataUang = ChrW(1583) & ChrW(1610) & ChrW(1606) & ChrW(1575) & ChrW(1585): TeksSen = ChrW(1601) & ChrW(1604) & ChrW(1587): GenderUtama = Muzakkar: GenderKecil = Muzakkar
        Case "ae": TeksMataUang = ChrW(1583) & ChrW(1585) & ChrW(1607) & ChrW(1605): TeksSen = ChrW(1601) & ChrW(1604) & ChrW(1587): GenderUtama = Muzakkar: GenderKecil = Muzakkar
        Case "qa": TeksMataUang = ChrW(1585) & ChrW(1610) & ChrW(1575) & ChrW(1604): TeksSen = ChrW(1583) & ChrW(1585) & ChrW(1607) & ChrW(1605): GenderUtama = Muzakkar: GenderKecil = Muzakkar
        Case "my": TeksMataUang = ChrW(1585) & ChrW(1610) & ChrW(1606) & ChrW(1580) & ChrW(1610) & ChrW(1578): TeksSen = ChrW(1587) & ChrW(1606): GenderUtama = Muzakkar: GenderKecil = Muzakkar
        Case "sg", "bn": TeksMataUang = ChrW(1583) & ChrW(1608) & ChrW(1604) & ChrW(1575) & ChrW(1585): TeksSen = ChrW(1587) & ChrW(1606): GenderUtama = Muzakkar: GenderKecil = Muzakkar
        Case "jp": TeksMataUang = ChrW(1610) & ChrW(1606): TeksSen = "": GenderUtama = Muzakkar
        Case "cn": TeksMataUang = ChrW(1610) & ChrW(1608) & ChrW(1575) & ChrW(1606): TeksSen = "": GenderUtama = Muzakkar
        Case Else: TeksMataUang = "": TeksSen = "": GenderUtama = Muzakkar: GenderKecil = Muzakkar
    End Select

    strAngka = CStr(angka)
    strAngka = Replace(strAngka, " ", "")
    strAngka = Replace(strAngka, ",", ".")

    PosKoma = InStr(strAngka, ".")
    If PosKoma > 0 Then
        BagianBulat = Left(strAngka, PosKoma - 1)
        BagianDesimal = Mid(strAngka, PosKoma + 1)
        If Len(BagianDesimal) = 1 Then BagianDesimal = BagianDesimal & "0"
        If Len(BagianDesimal) > 2 Then BagianDesimal = Left(BagianDesimal, 2)
    Else
        BagianBulat = strAngka: BagianDesimal = "0"
    End If

    ' --- BAGIAN BULAT (PERBAIKAN) ---
    If val(BagianBulat) = 0 Then
        HasilUtama = AngkaKeTeksArab(0, 1, GenderUtama, Modern, True, Marfu) & " " & TeksMataUang
    Else
        Dim TeksB As String
        ' 1. Ambil teks dasar dengan isTunggal = FALSE karena diikuti Ma'dud (Mata Uang)
        TeksB = AngkaKeTeksArab(BagianBulat, 1, GenderUtama, Modern, False, Marfu)
        
        ' 2. Jalankan IdhafahNun untuk koreksi Mutsanna (Contoh: Alfani -> Alfa)
        TeksB = IdhafahNun(TeksB, False)
        
        ' 3. Ambil satuan yang benar (Riyal/Riyalat/Riyalan) menggunakan 2 digit terakhir
        HasilUtama = TeksB & " " & GetMa_dud(val(Right(BagianBulat, 2)), TeksMataUang)
    End If
    
    ' --- BAGIAN DESIMAL (PERBAIKAN) ---
    If val(BagianDesimal) > 0 And TeksSen <> "" Then
        Dim TeksD As String
        ' Analog dengan bagian bulat, jalankan logika Idhafah
        TeksD = AngkaKeTeksArab(BagianDesimal, 1, GenderKecil, Modern, False, Marfu)
        TeksD = IdhafahNun(TeksD, False)
        
        HasilUtama = HasilUtama & WAW & TeksD & " " & GetMa_dud(val(BagianDesimal), TeksSen)
    End If
    
    ArabCurrency = Trim(HasilUtama)
End Function

Public Function ArabUniversal(ByVal angka As Variant, _
                              Optional ByVal JenisSatuan As Integer = 0, _
                              Optional ByVal PilihGender As Integer = 1, _
                              Optional ByVal IsOrdinal As Boolean = False) As String
Attribute ArabUniversal.VB_Description = "Menyebutkan satuan benda atau urutan posisi (Ordinal) secara otomatis."
Attribute ArabUniversal.VB_ProcData.VB_Invoke_Func = " \n19"

    ' --- Proteksi Error & Inisialisasi ---
    If IsError(angka) Then ArabUniversal = "[Khatha']": Exit Function
    If Not IsNumeric(angka) And Not IsOrdinal Then ArabUniversal = "[Khatha']": Exit Function

    InitializeArabicWords
    
    Dim ALI As String, baa As String, taa As String, tha As String
    Dim jim As String, haa As String, kha As String, dal As String
    Dim rho As String, sin As String, syi As String, sho As String
    Dim ain As String, fha As String, kaf As String, lam As String
    Dim mim As String, NUN As String, WAW As String, YAA As String
    Dim ham As String, mar As String, maq As String, gha As String
    Dim thoo As String, raa As String, dho As String, qof As String
    Dim zaa As String, dhal As String, HZA As String

    ALI = ChrW(1575): baa = ChrW(1576): taa = ChrW(1578): tha = ChrW(1579)
    jim = ChrW(1580): haa = ChrW(1581): kha = ChrW(1582): dal = ChrW(1583)
    rho = ChrW(1585): sin = ChrW(1587): syi = ChrW(1588): sho = ChrW(1589)
    ain = ChrW(1593): fha = ChrW(1601): kaf = ChrW(1603): lam = ChrW(1604)
    mim = ChrW(1605): NUN = ChrW(1606): WAW = ChrW(1608): YAA = ChrW(1610)
    ham = ChrW(1571): mar = ChrW(1577): maq = ChrW(1609): gha = ChrW(1594)
    thoo = ChrW(1591): raa = ChrW(1585): dho = ChrW(1590): qof = ChrW(1602)
    zaa = ChrW(1586): dhal = ChrW(1584): HZA = ChrW(1569)

    Dim hasilTerbilang As String, TeksSatuan As String
    Dim AL As String: AL = ALI & lam
    Dim GenderFinal As GenderArab: GenderFinal = IIf(PilihGender = 2, Muannas, Muzakkar)
    Dim v As Long: v = Fix(Abs(val(angka))) Mod 100

    Select Case JenisSatuan
        ' === KELOMPOK 1: JARAK & UKURAN ===
        Case 1: TeksSatuan = sin & NUN & taa & mim & taa & rho: GenderFinal = 1 ' CM
        Case 2: ' Meter (Muzakkar)
            If v = 2 Then: ArabUniversal = mim & taa & raa & ALI & NUN: Exit Function
            If v >= 3 And v <= 10 Then
                TeksSatuan = ham & mim & taa & ALI & rho
                GenderFinal = 2
            Else
                TeksSatuan = mim & taa & raa
                GenderFinal = 1
            End If
        Case 3: TeksSatuan = kaf & YAA & lam & WAW & mim & taa & raa: GenderFinal = 1 ' KM

        ' === KELOMPOK 2: BERAT & VOLUME ===
        Case 4: TeksSatuan = gha & raa & ALI & mim               ' Gram
        Case 5: TeksSatuan = kaf & YAA & lam & WAW & gha & raa & ALI & mim ' KG
        Case 6: ' Ton
            If v = 2 Then: ArabUniversal = thoo & NUN & ALI & NUN: Exit Function
            TeksSatuan = thoo & NUN: GenderFinal = 1
        Case 7: TeksSatuan = lam & taa & raa                     ' Liter

        ' === KELOMPOK 3: WAKTU (Muthanna & Jamak) ===
        Case 10: ' Jam (Sa'ah)
            If v = 2 Then
                TeksSatuan = sin & ALI & ain & taa & ALI & NUN: ArabUniversal = TeksSatuan: Exit Function
            ElseIf v >= 3 And v <= 10 Then
                TeksSatuan = sin & ALI & ain & ALI & taa: GenderFinal = Muzakkar
            Else
                TeksSatuan = sin & ALI & ain & mar: GenderFinal = Muannas
            End If
        Case 11: ' Hari (Yaum)
            If v = 2 Then
                TeksSatuan = YAA & WAW & mim & ALI & NUN: ArabUniversal = TeksSatuan: Exit Function
            ElseIf v >= 3 And v <= 10 Then
                TeksSatuan = ham & YAA & ALI & mim: GenderFinal = Muannas
            Else
                TeksSatuan = YAA & WAW & mim: GenderFinal = Muzakkar
            End If
        Case 12: ' Bulan (Syahr)
            If v = 2 Then
                TeksSatuan = syi & haa & rho & ALI & NUN: ArabUniversal = TeksSatuan: Exit Function
            ElseIf v >= 3 And v <= 10 Then
                TeksSatuan = ham & syi & haa & rho: GenderFinal = Muannas
            Else
                TeksSatuan = syi & haa & rho: GenderFinal = Muzakkar
            End If
        Case 13: ' Tahun (Sanah)
            If v = 2 Then
                TeksSatuan = sin & NUN & taa & ALI & NUN: ArabUniversal = TeksSatuan: Exit Function
            ElseIf v >= 3 And v <= 10 Then
                TeksSatuan = sin & NUN & WAW & ALI & taa: GenderFinal = Muzakkar
            Else
                TeksSatuan = sin & NUN & mar: GenderFinal = Muannas
            End If

        ' === KELOMPOK 4: SOSIAL & LITERASI ===
        Case 20: ' Orang (Syakhsh)
            If v = 2 Then
                TeksSatuan = syi & kha & sho & ALI & NUN: ArabUniversal = TeksSatuan: Exit Function
            ElseIf v >= 3 And v <= 10 Then
                TeksSatuan = ham & syi & kha & ALI & sho: GenderFinal = Muannas
            Else
                TeksSatuan = syi & kha & sho: GenderFinal = Muzakkar
            End If
        Case 21: ' Buku (Kitab)
            If v = 2 Then
                TeksSatuan = kaf & taa & ALI & baa & ALI & NUN: ArabUniversal = TeksSatuan: Exit Function
            ElseIf v >= 3 And v <= 10 Then
                TeksSatuan = kaf & taa & baa: GenderFinal = Muannas
            Else
                TeksSatuan = kaf & taa & ALI & baa: GenderFinal = Muzakkar
            End If
        Case 22: ' Halaman (Safhah)
            If v = 2 Then
                TeksSatuan = sho & fha & haa & taa & ALI & NUN: ArabUniversal = TeksSatuan: Exit Function
            ElseIf v >= 3 And v <= 10 Then
                TeksSatuan = sho & fha & haa & ALI & taa: GenderFinal = Muzakkar
            Else
                TeksSatuan = sho & fha & haa & mar: GenderFinal = Muannas
            End If

        ' === KELOMPOK 5: FISIKA & ELEKTRONIKA ===
        Case 40: TeksSatuan = fha & WAW & lam & taa              ' Volt
        Case 41: TeksSatuan = WAW & ALI & taa                    ' Watt
        Case 42: TeksSatuan = ham & mim & baa & YAA & rho        ' Ampere
        Case 43: TeksSatuan = dho & gha & thoo                   ' Press (Daghth)

        ' === KELOMPOK 6: SUHU, PERSEN & DATA ===
        Case 50: ' Derajat (Darajah - Muannas)
            If v = 2 Then: ArabUniversal = dal & rho & jim & taa & ALI & NUN: Exit Function
            If v >= 3 And v <= 10 Then
                TeksSatuan = dal & rho & jim & ALI & taa
                GenderFinal = 1
            Else
                TeksSatuan = dal & rho & jim & mar
                GenderFinal = 2
            End If
        Case 51: ' Persen (Fil-Miah)
        TeksSatuan = fha & YAA & " " & AL & mim & ALI & ChrW(1574) & mar: GenderFinal = 1 ' %
        Case 52: TeksSatuan = baa & ALI & YAA & taa: GenderFinal = 1 ' Byte
        Case 53: TeksSatuan = mim & YAA & gha & ALI & baa & ALI & YAA & taa: GenderFinal = 1  ' GB

        ' === KELOMPOK 7: POSISI & GEOGRAFI ===
        Case 60: TeksSatuan = AL & fha & sho & lam: GenderFinal = 1 ' Bab
        Case 61: TeksSatuan = AL & mim & raa & taa & baa & mar: GenderFinal = 2 ' Peringkat
        Case 62: TeksSatuan = AL & fha & raa & YAA & qof: GenderFinal = 1 ' Tim
        Case 63: TeksSatuan = kha & thoo & " " & AL & ALI & sin & taa & WAW & ALI & HZA: GenderFinal = 2 ' Khatulistiwa

        Case Else: TeksSatuan = ""
    End Select
    
    ' --- PROSES OUTPUT ---
    If IsOrdinal Then
        Dim ordText As String: ordText = GetArabicOrdinal(CInt(angka), GenderFinal)
        If InStr(ordText, " ") > 0 Then
            Dim parts() As String: parts = Split(ordText, " ")
            hasilTerbilang = AL & parts(0) & " " & parts(1)
        Else
            hasilTerbilang = AL & ordText
        End If
    Else
        hasilTerbilang = AngkaKeTeksArab(angka, 1, GenderFinal)
    End If

    ' --- PENGGABUNGAN ---
    If IsOrdinal And TeksSatuan <> "" Then
        ArabUniversal = TeksSatuan & " " & hasilTerbilang
    ElseIf TeksSatuan <> "" Then
        ArabUniversal = hasilTerbilang & " " & TeksSatuan
    Else
        ArabUniversal = hasilTerbilang
    End If
End Function

Private Function GetArabicOrdinal(ByVal n As Integer, ByVal g As GenderArab) As String
    If Not IsInitialized Then InitializeArabicWords
    If n >= 1 And n <= 12 Then
        GetArabicOrdinal = Ordinal(n, g)
    Else
        GetArabicOrdinal = "Error"
    End If
End Function

Sub RegisterArabFunctions()
    Dim ArgHelp(0 To 7) As String
    
    ' --- TERBILANG_ARAB ---
    ArgHelp(0) = "Angka atau referensi sel (Mendukung angka sangat besar)."
    ArgHelp(1) = "Mode: umum, uang, urutan, eja, benda, jarak, waktu."
    ArgHelp(2) = "Gender: m (Laki-laki) | f (Perempuan)."
    ArgHelp(3) = "I'rab: u (Marfu - Subjek) | a (Mansub/Majrur - Objek)."
    ArgHelp(4) = "Gaya: modern (spasi) | klasik (sambung - wa)."
    ArgHelp(5) = "Parameter Tambahan: ID Negara (id, sa) atau ID Benda (1-63)."
    ArgHelp(6) = "True: Dengan Harakat lengkap, False: Polos."
    ArgHelp(7) = "Thaman Lengkap: True (Ya), False (Tanpa Ya)."

    Application.MacroOptions Macro:="TERBILANG_ARAB", _
        Description:="Pintu utama konversi angka ke Arab dengan fitur I'rab, Gender, dan Harakat.", _
        Category:="ArabPro+ Tools", _
        ArgumentDescriptions:=ArgHelp
    
    MsgBox "Sistem ArabPro+ (Abjadun Mode) berhasil didaftarkan ke Excel!", vbInformation, "Sukses"
End Sub

Public Function TERBILANG_ARAB(ByVal angka As Variant, _
    Optional ByVal mode As String = "umum", _
    Optional ByVal gender As String = "m", _
    Optional ByVal irab As String = "u", _
    Optional ByVal gaya As String = "modern", _
    Optional ByVal Parameter_Negara As String = "id", _
    Optional ByVal PakaiHarakat As Boolean = False, _
    Optional ByVal isTunggal As Boolean = True) As String
Attribute TERBILANG_ARAB.VB_Description = "Pintu utama konversi angka ke Arab dengan fitur I'rab, Gender, dan Harakat."
Attribute TERBILANG_ARAB.VB_ProcData.VB_Invoke_Func = " \n19"

    On Error GoTo ErrHandler
    
    If IsEmpty(angka) Or Trim(angka) = "" Then
        TERBILANG_ARAB = "": Exit Function
    End If

' --- 2. Normalisasi I'rab ---
    Dim enIrab As IrabArab
    Select Case LCase(Trim(irab))
        Case "u", "marfu", "1": enIrab = 1 ' Marfu (u) -> Alif
        Case "a", "mansub", "2": enIrab = 2 ' Mansub (a) -> Ya (untuk dual)
        Case "i", "majrur", "3": enIrab = 3 ' Majrur (i) -> Ya (untuk dual)
        Case Else: enIrab = 1
    End Select

    ' --- 1. Normalisasi Gender ---
    Dim enGender As GenderArab
    Select Case LCase(Trim(gender))
        Case "f", "muannas", "p", "perempuan", "2": enGender = 2 ' Muannas
        Case Else: enGender = 1 ' Muzakkar
    End Select

    ' --- 2. Normalisasi I'rab & Gaya ---
    Dim enGaya As GayaArab
    Dim gayaLcase As String: gayaLcase = LCase(Trim(gaya))
    
    Select Case gayaLcase
        Case "klasik", "1": enGaya = 1 ' Klasik
        Case "sastra", "3": enGaya = 3 ' Sastra (INI YANG TADI HILANG)
        Case Else: enGaya = 2          ' Modern (Default)
    End Select

' --- 3. Eksekusi Berdasarkan Mode & Kelompok Satuan ---
    Dim HasilSementara As String
    Dim modeLcase As String: modeLcase = LCase(Trim(mode))
    Dim idUnit As Integer: idUnit = 0 ' Default: Tanpa satuan

    Select Case modeLcase
        ' -- [KELOMPOK DASAR] --
        Case "umum", "angka", "kardinal"
            HasilSementara = AngkaKeTeksArab(angka, 1, enGender, enGaya, False, enIrab, isTunggal)

        Case "eja", "digit"
            HasilSementara = AngkaKeTeksArab(angka, 2, enGender, enGaya, False, enIrab, isTunggal)

        ' -- [KELOMPOK MATA UANG] --
        Case "uang", "currency"
            HasilSementara = ArabCurrency(angka, Parameter_Negara)

        ' -- [KELOMPOK 1 & 2: UKURAN & MASSA] --
        Case "jarak", "km":   idUnit = 3
        Case "meter", "m":    idUnit = 2
        Case "liter", "l":    idUnit = 7
        Case "berat", "kg":   idUnit = 5

        ' -- [KELOMPOK 3: WAKTU] --
        Case "waktu", "hari": idUnit = 11
        Case "jam":           idUnit = 10
        Case "bulan":         idUnit = 12
        Case "tahun":         idUnit = 13

        ' -- [KELOMPOK 4 & 7: LITERASI & POSISI] --
        Case "buku", "kitab": idUnit = 21
        Case "halaman":       idUnit = 22
        Case "urutan", "ordinal"
            ' Otomatis pilih Bab (60) atau Peringkat (61) berdasarkan gender
            idUnit = IIf(enGender = 1, 60, 61)
            HasilSementara = ArabUniversal(angka, idUnit, enGender, True)
            GoTo Finalize

        ' -- [MODE CUSTOM / LAINNYA] --
        Case "benda", "unit", "lainnya"
            idUnit = val(Parameter_Negara)
            
        Case Else
            ' Jika user memasukkan teks yang tidak terdaftar, cek apakah itu angka ID benda
            idUnit = val(mode)
    End Select

    ' Eksekusi Logika Satuan jika idUnit terisi
    If idUnit > 0 Then
            ' Standar Kardinal untuk unit lainnya
            HasilSementara = ArabUniversal(angka, idUnit, enGender, False)
    ElseIf HasilSementara = "" Then
        ' Jika benar-benar kosong, default ke angka kardinal umum
        HasilSementara = AngkaKeTeksArab(angka, 1, enGender, enGaya, False, enIrab, isTunggal)
    End If

Finalize:
    ' --- 4. Proses Akhir: Harakat ---
    If PakaiHarakat Then
        TERBILANG_ARAB = BeriHarakat(HasilSementara, enIrab)
    Else
        TERBILANG_ARAB = HasilSementara
    End If
    Exit Function

ErrHandler:
    TERBILANG_ARAB = "[Khatha']"
End Function






' ================= HARAKAT =========================================================
Private Function MapKataHarakat(ByVal w As String, ByVal irab As Integer) As String
    ' === DEFINISI HARAKAT ===
    Dim f As String: f = ChrW(1614):   Dim k As String: k = ChrW(1616)
    Dim d As String: d = ChrW(1615):   Dim s As String: s = ChrW(1618)
    Dim sy As String: sy = ChrW(1617): Dim dt As String: dt = ChrW(1612)
    
    ' === HARAKAT AKHIR DINAMIS (u/a/i) ===
    Dim ha As String
    Select Case irab
        Case 1: ha = d ' Marfu (u)
        Case 2: ha = f ' Mansub (a)
        Case 3: ha = k ' Majrur (i)
    End Select

    ' === DEFINISI HURUF ===
    Dim ALI As String: ALI = ChrW(1575): Dim baa As String: baa = ChrW(1576)
    Dim taa As String: taa = ChrW(1578): Dim tha As String: tha = ChrW(1579)
    Dim jim As String: jim = ChrW(1580): Dim haa As String: haa = ChrW(1581)
    Dim kha As String: kha = ChrW(1582): Dim dal As String: dal = ChrW(1583)
    Dim rho As String: rho = ChrW(1585): Dim sin As String: sin = ChrW(1587)
    Dim syi As String: syi = ChrW(1588): Dim ain As String: ain = ChrW(1593)
    Dim fha As String: fha = ChrW(1601): Dim kaf As String: kaf = ChrW(1603)
    Dim lam As String: lam = ChrW(1604): Dim mim As String: mim = ChrW(1605)
    Dim NUN As String: NUN = ChrW(1606): Dim WAW As String: WAW = ChrW(1608)
    Dim YAA As String: YAA = ChrW(1610): Dim ham As String: ham = ChrW(1571)
    Dim mar As String: mar = ChrW(1577): Dim maq As String: maq = ChrW(1609)
    Dim HZA As String: HZA = ChrW(1574): Dim sho As String: sho = ChrW(1589)

    ' --- A. LOGIKA DINAMIS PULUHAN (20-90) ---
    If Len(w) >= 4 Then
        Dim akhir As String: akhir = Right(w, 2)
        If akhir = WAW & NUN Or akhir = YAA & NUN Then
            Dim akar As String: akar = Left(w, Len(w) - 2)
            Dim hAkhir As String
            hAkhir = IIf(irab = 1, d & WAW & NUN & f, k & YAA & s & NUN & f)
            
            Select Case akar
                Case ain & syi & rho: MapKataHarakat = ain & k & syi & s & rho & hAkhir: Exit Function
                Case tha & lam & ALI & tha: MapKataHarakat = tha & f & lam & f & ALI & tha & hAkhir: Exit Function
                Case ham & rho & baa & ain: MapKataHarakat = ham & f & rho & s & baa & f & ain & hAkhir: Exit Function
                Case kha & mim & sin: MapKataHarakat = kha & f & mim & s & sin & hAkhir: Exit Function
                Case sin & taa: MapKataHarakat = sin & k & taa & sy & hAkhir: Exit Function
                Case sin & baa & ain: MapKataHarakat = sin & f & baa & s & ain & hAkhir: Exit Function
                Case tha & mim & ALI & NUN: MapKataHarakat = tha & f & mim & f & ALI & NUN & hAkhir: Exit Function
                Case taa & sin & ain: MapKataHarakat = taa & k & sin & s & ain & hAkhir: Exit Function
            End Select
        End If
    End If

    ' --- B. LOGIKA DUAL (Angka 2 / 12) ---
    If w = ALI & tha & NUN & ALI Then
        MapKataHarakat = IIf(irab = 1, ALI & k & tha & s & NUN & f & ALI, ALI & k & tha & s & NUN & f & YAA & s)
        Exit Function
    End If

    ' --- C. KATA STATIS & DINAMIS ---
    Select Case w
        ' Nol
        Case sho & fha & rho: MapKataHarakat = sho & k & fha & s & rho & ha

        ' --- ANGKA 10 ---
        Case ain & syi & rho & mar: MapKataHarakat = ain & f & syi & s & rho & f & mar & ha ' Muzakkar (Asyrah)
        Case ain & syi & rho:       MapKataHarakat = ain & f & syi & s & rho & ha         ' Muannas (Asyr)

        ' --- ANGKA 11 (Mabni Fathah) ---
        Case ham & haa & dal: MapKataHarakat = ham & f & haa & f & dal & f
        Case ChrW(1573) & haa & dal & maq: MapKataHarakat = ChrW(1573) & k & haa & s & dal & f & maq
        
        ' --- BAGIAN BELASAN (11-19) ---
        Case ain & syi & rho & f, ain & syi & rho & mar & f: MapKataHarakat = w: Exit Function

        ' Satuan Muzakkar (3-9)
        Case WAW & ALI & haa & dal: MapKataHarakat = WAW & f & ALI & haa & k & dal & ha
        Case tha & lam & ALI & tha: MapKataHarakat = tha & f & lam & f & ALI & tha & ha
        Case ham & rho & baa & ain: MapKataHarakat = ham & f & rho & s & baa & f & ain & ha
        Case kha & mim & sin:       MapKataHarakat = kha & f & mim & s & sin & ha
        Case sin & taa:             MapKataHarakat = sin & k & taa & sy & ha
        Case sin & baa & ain:       MapKataHarakat = sin & f & baa & s & ain & ha
        Case tha & mim & ALI & NUN: MapKataHarakat = tha & f & mim & f & ALI & NUN & ha
        Case taa & sin & ain:       MapKataHarakat = taa & k & sin & s & ain & ha

        ' Angka 8 Muannas (Thamani)
        Case tha & mim & ALI & NUN & YAA: MapKataHarakat = tha & f & mim & f & ALI & NUN & k & YAA & IIf(irab = 2, f, s)

        ' --- RATUSAN
        ' 100: Mi'ah
        Case mim & ALI & HZA & mar, mim & HZA & mar, mim & ALI & ham & mar, mim & ham & mar
             MapKataHarakat = mim & k & ALI & HZA & f & mar & ha
             
             ' --- RIBUAN (Alf) ---
        ' 1000: Alf (Mufrad)
        Case ham & lam & fha: MapKataHarakat = ham & f & lam & s & fha & ha
        
        ' 2000: Alfaini / Alfaa (Dual)
        Case ham & lam & fha & ALI & NUN:
             MapKataHarakat = IIf(irab = 1, ham & f & lam & f & fha & ALI & NUN & k, ham & f & lam & f & fha & YAA & s & NUN & k)
        
        ' 3000-10.000: Aalaaf (Jamak Taksir)
        Case ham & ALI & lam & ALI & fha:
             MapKataHarakat = ChrW(1570) & lam & f & ALI & fha & ha
             
        ' Jutaan (Milyun)
        Case mim & lam & YAA & WAW & NUN: MapKataHarakat = mim & k & lam & s & YAA & d & WAW & NUN & ha

        ' Indonesia: Rupiah
        Case rho & WAW & baa & YAA & mar:        MapKataHarakat = rho & d & WAW & s & baa & k & YAA & sy & f & mar
        Case rho & WAW & baa & YAA & ALI & taa:  MapKataHarakat = rho & d & WAW & s & baa & k & YAA & sy & f & ALI & taa
        
        Case Else: MapKataHarakat = w
    End Select
End Function

Function HapusHarakat(ByVal Teks As String) As String
    Dim i As Long, hasil As String, c As Long
    For i = 1 To Len(Teks)
        c = AscW(Mid(Teks, i, 1))
        ' Abaikan karakter Unicode harakat (Fathah, Dammah, Kasrah, dll)
        If c < 1611 Or c > 1618 Then hasil = hasil & Mid(Teks, i, 1)
    Next i
    HapusHarakat = hasil
End Function

Private Function BeriHarakat(ByVal Teks As String, ByVal irab As Integer) As String
    Dim kata() As String, i As Integer, hasil() As String
    Dim WAW As String: WAW = ChrW(1608)
    
    kata = Split(Teks, " ")
    ReDim hasil(0 To UBound(kata))
    
    For i = 0 To UBound(kata)
        ' Jika Gaya Klasik (Wau menempel)
        If Left(kata(i), 1) = WAW And Len(kata(i)) > 1 Then
            Dim sisa As String: sisa = Mid(kata(i), 2)
            hasil(i) = (WAW & ChrW(1614)) & MapKataHarakat(sisa, irab)
        Else
            hasil(i) = MapKataHarakat(kata(i), irab)
        End If
    Next i
    
    BeriHarakat = Join(hasil, " ")
End Function


'==============================================================================================
'==================================== TEST +++++++++++++++======================================
