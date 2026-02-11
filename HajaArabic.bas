Attribute VB_Name = "Module1"
' ==============================================================================
'  HajaArabic (VBA Terbilang Arab)
' ==============================================================================
' Copyright (c) 2026 Rida Rahman DH 96-02 — Banjarmasin, Indonesia
'  GitHub: https://github.com/Rida-Haja
'  Email : RidaHaja@gmail.com
' Februari 2026
' ==============================================================================
' LICENSE / LISENSI
' ==============================================================================
'
' [ENGLISH]
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software (the "HajaArabic"), to deal in the Software without restriction,
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
' perangkat lunak ini (HajaArabic), untuk menggunakan Perangkat Lunak tanpa
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
    Modern = 1 ' Standar Modern (Spasi jelas, istilah umum)
    Klasik = 2 ' Fusha Turats (Sambung Wawu, istilah klasik)
    Sastra = 3 ' Gaya Sastra/Puitis
End Enum

Public Enum IrabArab
    Marfu = 1       ' Nominative (Dhommah/Waw-Nun)
    Mansub = 2      ' Accusative (Fathah/Ya-Nun)
    Majrur = 3      ' Genitive (Kasrah/Ya-Nun)
End Enum

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
    
    Dim ALI As String, baa As String, taa As String, THA As String
    Dim jim As String, haa As String, kha As String, dal As String
    Dim rho As String, sin As String, syi As String, sho As String
    Dim ain As String, fha As String, kaf As String, lam As String
    Dim MIM As String, NUN As String, WAW As String, YAA As String
    Dim ham As String, MAR As String, maq As String
    Dim MAD As String, hza As String, HAW As String, HAB As String
    Dim i As Integer: Dim v As Variant
    
    ' Array Unicode
    v = Array(1575, 1576, 1578, 1579, 1580, 1581, 1582, 1583, 1585, 1587, 1588, 1589, _
              1593, 1601, 1603, 1604, 1605, 1606, 1608, 1610, 1571, 1577, 1609) '

    For i = 0 To UBound(v): v(i) = ChrW(v(i)): Next i

    ALI = v(0):  baa = v(1):  taa = v(2):  THA = v(3):  jim = v(4)
    haa = v(5):  kha = v(6):  dal = v(7):  rho = v(8):  sin = v(9)
    syi = v(10): sho = v(11): ain = v(12): fha = v(13): kaf = v(14)
    lam = v(15): MIM = v(16): NUN = v(17): WAW = v(18): YAA = v(19)
    ham = v(20): MAR = v(21): maq = v(22)
    
    ' --- 0: Sifrun ---
    kata(0, 1) = sho & fha & rho
    kata(0, 2) = kata(0, 1)

    ' --- 1-9: Satuan Dasar ---
    kata(1, 1) = WAW & ALI & haa & dal           ' Wahid
    kata(2, 1) = ALI & THA & NUN & ALI & NUN     ' Ithnan
    kata(3, 1) = THA & lam & ALI & THA           ' Thalath
    kata(4, 1) = ham & rho & baa & ain           ' Arba
    kata(5, 1) = kha & MIM & sin                 ' Khams
    kata(6, 1) = sin & taa                       ' Sitt
    kata(7, 1) = sin & baa & ain                 ' Sab'a
    kata(8, 1) = THA & MIM & ALI & NUN & YAA               ' Thamaniy
    kata(9, 1) = taa & sin & ain                 ' Tis'a
    
    ' --- 10: Asyar / Asyarah ---
    kata(10, 1) = ain & syi & rho
    kata(10, 2) = kata(10, 1) & MAR
    
    ' Muannas: Penambahan Ta Marbutah (mar)
    kata(1, 2) = kata(1, 1) & MAR                ' Wahidah
    For i = 3 To 9
        If i <> 8 Then kata(i, 2) = kata(i, 1) & MAR
    Next i
    kata(2, 2) = ALI & THA & NUN & taa & ALI & NUN ' Ithnatan
    kata(8, 2) = kata(8, 1) & MAR             ' Thamaniyah

    ' --- PULUHAN (Isyrun - Tis'un) ---
    Puluhan(2) = ain & syi & rho & WAW & NUN      ' 'Isyrun
    For i = 3 To 9
        Puluhan(i) = kata(i, 1) & WAW & NUN
    Next i

' --- SATUAN BESAR (Skala Panjang Eropa Klasik) ---
    ' Skema: n-ilyun (10^6n) dan n-ilyar (10^6n+3)
    SatuanBesar(1) = ham & lam & fha                                 ' 10^3: Alfun (Ribu)
    SatuanBesar(2) = MIM & lam & YAA & WAW & NUN                     ' 10^6: Milyun (Juta)
    SatuanBesar(3) = MIM & lam & YAA & ALI & rho                     ' 10^9: Milyar (Miliar)
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
    Dim Asyarah As String: Asyarah = Asyar & MAR

    ' Pola Fa'il (3,4,5,7,9)
    For i = 3 To 9
        If i <> 6 And i <> 8 Then
            Ordinal(i, 1) = Left(kata(i, 1), 1) & ALI & Mid(kata(i, 1), 2)
            Ordinal(i, 2) = Ordinal(i, 1) & MAR
        End If
    Next i

    ' Irregular Ordinals
    Ordinal(1, 1) = ham & WAW & lam               ' Awwal
    Ordinal(1, 2) = ham & WAW & lam & maq         ' Ula
    Ordinal(2, 1) = THA & ALI & NUN & YAA         ' Thani
    Ordinal(2, 2) = Ordinal(2, 1) & MAR           ' Thaniyah
    Ordinal(6, 1) = sin & ALI & dal & sin         ' Sadis
    Ordinal(6, 2) = Ordinal(6, 1) & MAR
    Ordinal(8, 1) = THA & ALI & MIM & NUN      ' Thamin
    Ordinal(8, 2) = Ordinal(8, 1) & MAR
    Ordinal(10, 1) = ain & ALI & syi & rho        ' Asyir
    Ordinal(10, 2) = Ordinal(10, 1) & MAR
    
    ' Compound Ordinals (11-12)
    Ordinal(11, 1) = haa & ALI & dal & YAA & " " & Asyar
    Ordinal(11, 2) = haa & ALI & dal & YAA & MAR & " " & Asyarah
    Ordinal(12, 1) = Ordinal(2, 1) & " " & Asyar
    Ordinal(12, 2) = Ordinal(2, 1) & MAR & " " & Asyarah

    IsInitialized = True
End Sub

'=================== BAB TERBILANG ============================================
Private Function ProsesBlokRibuan(ByVal strNum As Variant, _
                                  ByVal PilihGender As GenderArab, _
                                  ByVal gaya As GayaArab, _
                                  Optional irab As IrabArab = Marfu, _
                                  Optional isIdhafah As Boolean = False) As String
    Dim wordList() As String
    Dim finalArr() As String
    Dim blockCount As Integer
    Dim i As Integer, countVal As Integer
    Dim WAW As String: WAW = ChrW(1608)

    Dim cleanNum As String: cleanNum = ""
    Dim rawInput As String: rawInput = Trim(CStr(strNum))
    For i = 1 To Len(rawInput)
        If Mid(rawInput, i, 1) >= "0" And Mid(rawInput, i, 1) <= "9" Then cleanNum = cleanNum & Mid(rawInput, i, 1)
    Next i
    If cleanNum = "" Then Exit Function
    
    Do While Len(cleanNum) Mod 3 <> 0
        cleanNum = "0" & cleanNum
    Loop
    
    blockCount = Len(cleanNum) \ 3
    ReDim wordList(0 To blockCount - 1)

    For i = 0 To blockCount - 1
        Dim startPos As Integer: startPos = Len(cleanNum) - (i * 3) - 2
        Dim strTemp As String: strTemp = Mid(cleanNum, startPos, 3)
        Dim vTemp As Integer: vTemp = CInt(val(strTemp))
        
        If vTemp > 0 Then
            Dim txtResult As String
            If i = 0 Then
                ' Blok Satuan (0-999)
                txtResult = ConvertThreeDigitsArab(vTemp, PilihGender, gaya, irab, isIdhafah)
            Else
                ' Blok Ribuan/Jutaan/Milyar
                txtResult = FormatSatuanBesar(strTemp, i, gaya, irab, isIdhafah)
            End If
            wordList(i) = txtResult
        End If
    Next i

    Dim junc As String
    junc = IIf(gaya = Klasik, " " & WAW, " " & WAW & " ")
    
    countVal = 0
    If gaya = Sastra Then
        For i = 0 To blockCount - 1
            If wordList(i) <> "" Then
                ReDim Preserve finalArr(countVal): finalArr(countVal) = wordList(i)
                countVal = countVal + 1
            End If
        Next i
    Else
        For i = blockCount - 1 To 0 Step -1
            If wordList(i) <> "" Then
                ReDim Preserve finalArr(countVal): finalArr(countVal) = wordList(i)
                countVal = countVal + 1
            End If
        Next i
    End If
    
 Dim THA As String: THA = ChrW(1579)
 Dim ALI As String: ALI = ChrW(1575)
 Dim NUN As String: NUN = ChrW(1606)
 Dim YAA As String: YAA = ChrW(1610)
 Dim lam As String: lam = ChrW(1604)
 Dim FAA As String: FAA = ChrW(1601)
 Dim MIM As String: MIM = ChrW(1605)
 Dim HAMZAH_YA As String: HAMZAH_YA = ChrW(1574)
 Dim MAR As String: MAR = ChrW(1577)
     
' --- Idhafah ---
If isIdhafah Then
    Dim fullText As String: fullText = Join(finalArr, junc)
    Dim arTeks() As String: arTeks = Split(fullText, " ")
    Dim kataTerakhir As String: kataTerakhir = arTeks(UBound(arTeks))
    
    If Right(kataTerakhir, 1) = NUN And Len(kataTerakhir) > 2 Then
        Dim sisaKata As String: sisaKata = Left(kataTerakhir, Len(kataTerakhir) - 1)
        Dim duaHurufTerakhir As String: duaHurufTerakhir = Right(kataTerakhir, 2)
        
        If duaHurufTerakhir = ALI & NUN Or duaHurufTerakhir = YAA & NUN Then
            
            Dim isMalayiin As Boolean: isMalayiin = (InStr(kataTerakhir, MIM & lam & ALI & YAA) > 0)
            Dim isMilyunTunggal As Boolean: isMilyunTunggal = (InStr(kataTerakhir, WAW & NUN) > 0)
            
            ' isDualNumber: Memastikan ini angka 2, 200, 2000, atau skala Dual
            Dim isDualNumber As Boolean
            isDualNumber = (InStr(kataTerakhir, THA) > 0) Or _
                           (InStr(kataTerakhir, MIM & ALI & HAMZAH_YA) > 0) Or _
                           (InStr(kataTerakhir, FAA & ALI) > 0) Or _
                           (InStr(kataTerakhir, FAA & YAA) > 0) Or _
                           (InStr(kataTerakhir, ALI & lam & FAA) > 0) Or _
                           (InStr(kataTerakhir, YAA & NUN & ALI) > 0)
            
            If (Not isMalayiin) And (Not isMilyunTunggal) Then
                If duaHurufTerakhir = YAA & NUN Then
                    If isDualNumber Then kataTerakhir = sisaKata
                Else
                    kataTerakhir = sisaKata
                End If
            End If
        End If
    End If
    
    arTeks(UBound(arTeks)) = kataTerakhir
    ProsesBlokRibuan = Join(arTeks, " ")
Else
    ProsesBlokRibuan = Join(finalArr, junc)
End If
End Function

Private Function ConvertThreeDigitsArab(ByVal num As Variant, _
                                      ByVal PilihGender As GenderArab, _
                                      ByVal gaya As GayaArab, _
                                      Optional irab As IrabArab = Marfu, _
                                      Optional isIdhafah As Boolean = False) As String
    Dim s As String
    Dim MIM As String: MIM = ChrW(1605): Dim ALI As String: ALI = ChrW(1575)
    Dim taa As String: taa = ChrW(1578): Dim NUN As String: NUN = ChrW(1606)
    Dim YAA As String: YAA = ChrW(1610): Dim maq As String: maq = ChrW(1609)
    Dim ham As String: ham = ChrW(1574): Dim MAR As String: MAR = ChrW(1577)
    Dim THA As String: THA = ChrW(1579): Dim WAW As String: WAW = ChrW(1608)
    Dim lam As String: lam = ChrW(1604): Dim FAA As String: FAA = ChrW(1601)
    
    Dim v As Long: v = CLng(num)
    If v = 0 Then Exit Function
    Dim spWaw As String: spWaw = IIf(gaya = Klasik, " " & WAW, " " & WAW & " ")
    ' --- RATUSAN ---
    If v >= 100 Then
        Dim h As Integer: h = Int(v / 100)
        Dim sisaRatusan As Integer: sisaRatusan = v Mod 100
        
        Select Case h
            Case 1: s = IIf(gaya = Modern, MIM & ham & MAR, MIM & ALI & ham & MAR)
            Case 2
                ' 200: Mi'atani (Selalu pakai Nun di sini, nanti dibuang di FormatSatuanBesar jika perlu)
                s = IIf(gaya = Modern, MIM & ham & taa, MIM & ALI & ham & taa)
                s = s & IIf(irab = Marfu, ALI, YAA) & NUN
            Case 3 To 9
                Dim txtH As String: txtH = kata(h, 1) ' Ratusan pakai Muzakkar (Tsalatsu Mi'ah)
                If gaya = Sastra Then
                    s = txtH & " " & MIM & ALI & ham & MAR
                ElseIf gaya = Modern Then
                    s = txtH & " " & MIM & ham & MAR
                Else: s = txtH & MIM & ALI & ham & MAR
                End If
        End Select
        
        v = sisaRatusan
        If v > 0 Then s = s & spWaw
    End If

    ' --- PULUHAN DAN SATUAN ---
    If v > 0 Then
        If Len(s) > 0 And Right(s, 1) <> WAW And Right(s, 1) <> " " Then
            If gaya <> Klasik Then s = s & " "
        End If
        
        Dim idxLawan As Integer: idxLawan = IIf(PilihGender = Muzakkar, 2, 1)

        If v <= 10 Then
                    If v = 1 Or v = 2 Then
                s = s & GetWahidIthnan(v, PilihGender, irab)
            ElseIf v = 8 Then
                ' Logika Ism Manqus untuk Angka 8
                If PilihGender = Muzakkar Then
                    ' KAIDAH: Ma'dud Pria -> Angka Wanita (Thamaniyah)
                    ' KECUALI: Gaya Klasik dalam Idhafah boleh menggunakan Muzakkar (Thamani)
                    'If gaya = Klasik And isIdhafah Then
                    '    s = s & kata(8, 1) ' Thamani
                    'Else
                        s = s & kata(8, 2) ' Thamaniyah
                    'End If
                Else
                    ' Ma'dud Muannas (seperti Rupiah) -> Angka harus Muzakkar
                    If isIdhafah Then
                        s = s & kata(8, 1) ' Thamani
                    Else
                        s = s & IIf(irab = Mansub, THA & MIM & ALI & NUN & YAA & ALI & ChrW(1611), THA & MIM & ALI & NUN & ChrW(1613))
                    End If
                End If
            Else
                s = s & kata(v, idxLawan)
            End If
        ElseIf v < 20 Then
            ' Belasan (11-19)
            Dim unitB As Integer: unitB = v Mod 10
            Dim txtU As String
            If unitB = 1 Then ' 11
                txtU = IIf(PilihGender = Muzakkar, ChrW(1571) & ChrW(1581) & ChrW(1583), ChrW(1573) & ChrW(1581) & ChrW(1583) & maq)
            ElseIf unitB = 2 Then ' 12
                ' 12: KHUSUS - Selalu Tanpa Nun (Ithna/Ithnata)
                Dim suf As String: suf = IIf(irab = Marfu, ALI, YAA)
                txtU = ALI & THA & NUN & IIf(PilihGender = Muannas, taa & suf, suf)
            Else
                txtU = kata(unitB, idxLawan)
            End If
            s = s & txtU & " " & kata(10, PilihGender)
        Else
            ' Puluhan (20-99)
            Dim st As Integer: st = v Mod 10
            Dim pul As String: pul = Puluhan(Int(v / 10))
            If irab <> Marfu Then pul = Replace(pul, WAW & NUN, YAA & NUN)
            
            If st = 0 Then
                s = s & pul
            Else
                Dim txtSat As String: txtSat = IIf(st = 1 Or st = 2, GetWahidIthnan(st, PilihGender, irab), kata(st, idxLawan))
                s = s & txtSat & spWaw & pul
            End If
        End If
    End If

        ConvertThreeDigitsArab = Trim(s)
End Function

Private Function FormatSatuanBesar(ByVal vNum As Variant, ByVal idx As Integer, _
                                   ByVal gaya As GayaArab, ByVal irab As IrabArab, _
                                   Optional isIdhafah As Boolean = False) As String
    Dim Cur As String
    Dim ALI As String: ALI = ChrW(1575): Dim YAA As String: YAA = ChrW(1610): Dim NUN As String: NUN = ChrW(1606)
    Dim v As Integer: v = CInt(val(vNum))
    If v = 0 Then Exit Function

    Dim baseUnit As String: baseUnit = SatuanBesar(idx)

    If v = 1 Then
        Cur = baseUnit
    ElseIf v = 2 Then
        ' --- ANGKA 2.000, 2 JUTA, dsb ---
        Dim root As String: root = IIf(idx = 1, ChrW(1571) & ChrW(1604) & ChrW(1601), baseUnit)
        Cur = root & IIf(irab = Marfu, ALI, YAA)
        
        ' Peluruhan Nun untuk Idhafah pada Mutsanna (2000 -> Alfai / Alfaa)
        If Not isIdhafah Then
            Cur = Cur & NUN
        End If
    Else
    ' --- ANGKA 3 - 999 ---
        Dim unitFinal As String: unitFinal = baseUnit
        Dim resAngka As String
        
        ' 1. Tentukan Bentuk Angka
        ' Pengecekan Khusus: Gaya Klasik hanya untuk Ribuan (idx = 1) dan angka pas 8
        If v = 8 And idx = 1 And gaya = Klasik Or gaya = Sastra Then
            resAngka = kata(8, 1)
        Else
            resAngka = ConvertThreeDigitsArab(v, Muzakkar, gaya, irab, False)
        End If
        
        ' 2. Tentukan Unit (Singular/Plural/Accusative)
        If v >= 3 And v <= 10 Then
             ' 3-10: Idhafah ke Jamak (Aalaaf, Malayiin)
             unitFinal = GetMa_dud(v, baseUnit)
        Else
             ' 11-999: Tamyiz Mufrad Manshub (Alfan, Malyunan)
             If (v Mod 100 >= 11 And v Mod 100 <= 99) Then unitFinal = unitFinal & ALI
        End If

        ' 3. LOGIKA PELULUHAN NUN (IDHAFAH) UNTUK ANGKA 200
        ' Jika angkanya PERSIS 200 (Mi'atani), Nun harus dibuang saat bertemu Satuan Besar (Mi'ata Alfin)
        ' Jika 205 (Mi'atani wa Khamsah), Nun Tetap Ada.
        ' Angka 12 sudah ditangani di ConvertThreeDigitsArab (tanpa Nun).
        
        If v = 200 Then
            ' Hapus karakter terakhir (Nun)
            resAngka = Left(resAngka, Len(resAngka) - 1)
        End If
        
        Cur = resAngka & " " & unitFinal
    End If
    FormatSatuanBesar = Cur
End Function

Private Function GetWahidIthnan(ByVal val As Variant, ByVal gender As GenderArab, ByVal irab As IrabArab) As String
    Dim ALI As String: ALI = ChrW(1575): Dim NUN As String: NUN = ChrW(1606)
    Dim YAA As String: YAA = ChrW(1610): Dim THA As String: THA = ChrW(1579)
    Dim taa As String: taa = ChrW(1578)
    
    If val = 1 Then
        GetWahidIthnan = kata(1, gender)
    ElseIf val = 2 Then
        Dim base As String
        base = ALI & THA & NUN & IIf(gender = Muannas, taa, "")
        GetWahidIthnan = base & IIf(irab = Marfu, ALI, YAA) & NUN
    End If
End Function

Private Function GetMa_dud(ByVal angka As Long, ByVal satuan As String) As String
    Dim rho As String, YAA As String, ALI As String, lam As String, taa As String
    Dim dal As String, NUN As String, raa As String, haa As String, MIM As String
    Dim WAW As String, baa As String, MAR As String

    rho = ChrW(1585): YAA = ChrW(1610): ALI = ChrW(1575): lam = ChrW(1604): taa = ChrW(1578)
    dal = ChrW(1583): NUN = ChrW(1606): raa = ChrW(1585): haa = ChrW(1607): MIM = ChrW(1605)
    WAW = ChrW(1608): baa = ChrW(1576): MAR = ChrW(1577)

    Dim v As Long: v = angka Mod 100
    
    Dim sRiyal As String:  sRiyal = rho & YAA & ALI & lam
    Dim sDinar As String:  sDinar = dal & YAA & NUN & ALI & raa
    Dim sDirham As String: sDirham = dal & rho & haa & MIM
    Dim sRupiah As String: sRupiah = rho & WAW & baa & YAA & MAR

    Select Case True
        Case satuan = sRiyal
            GetMa_dud = IIf(v >= 3 And v <= 10, sRiyal & ALI & taa, satuan)
            
        Case satuan = sDinar
            ' Jamak Taktsir: Dananir
            GetMa_dud = IIf(v >= 3 And v <= 10, dal & NUN & ALI & NUN & YAA & raa, satuan)
            
        Case satuan = sDirham
            ' Jamak Taktsir: Darahim
            GetMa_dud = IIf(v >= 3 And v <= 10, dal & raa & ALI & haa & MIM, satuan)
                
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
        Case satuan = ChrW(1601) & ChrW(1604) & ChrW(1587) ' Fils
             GetMa_dud = IIf(v >= 3 And v <= 10, ChrW(1601) & ChrW(1604) & ChrW(1608) & ChrW(1587), satuan)
        Case Else
            GetMa_dud = satuan
    End Select
End Function

'=================== BAB TAMBAHAN LAINNYA ============================================
Private Function AngkaKeTeksArab(ByVal angka As Variant, _
                                Optional ByVal mode As Byte = 1, _
                                Optional ByVal gender As GenderArab = Muzakkar, _
                                Optional ByVal gaya As GayaArab = Modern, _
                                Optional ByVal isCurrency As Boolean = False, _
                                Optional ByVal irab As IrabArab = Marfu, _
                                Optional ByVal isIdhafah As Boolean = False) As String
    
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
            FinalBulat = ProsesBlokRibuan(bulat, gender, gaya, irab, isIdhafah)
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
                    FinalDes = Pemisah & ProsesBlokRibuan(val(Right(senVal, 2)), Muannas, gaya, irab, True)
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

Public Function ArabCurrency(ByVal angka As Variant, _
                             Optional ByVal KodeNegara As String = "id") As String
Attribute ArabCurrency.VB_Description = "Konverter Nominal Uang ke Teks Bahasa Arab\r\n---------------------------------------------\r\nMenampilkan satuan mata uang dan pecahan (sen/halala) \r\nsecara otomatis sesuai negara yang dipilih."
Attribute ArabCurrency.VB_ProcData.VB_Invoke_Func = " \n19"
    ' ==================================================================================
    ' FUNGSI: Mengonversi angka ke terbilang mata uang dalam bahasa Arab
    ' ==================================================================================
    
    Dim HasilUtama As String, strAngka As String
    Dim BagianBulat As String, BagianDesimal As String
    Dim PosKoma As Long
    Dim TeksMataUang As String, TeksSen As String, WAW As String
    Dim GenderUtama As GenderArab, GenderKecil As GenderArab
    
    ' kata hubung "dan" (Waw)
    WAW = " " & ChrW(1608) & " "
    KodeNegara = LCase(Trim(KodeNegara))
    
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

    ' --- BAGIAN BULAT ---
    If val(BagianBulat) = 0 Then
        HasilUtama = AngkaKeTeksArab(0, 1, GenderUtama, Modern, True, Marfu) & " " & TeksMataUang
    Else
        Dim TeksB As String
        ' 1. Ambil teks dasar dengan isIdhafah = TRUE karena diikuti Ma'dud (Mata Uang)
        TeksB = AngkaKeTeksArab(BagianBulat, 1, GenderUtama, Modern, False, Marfu, True)
        
        
        ' 3. Ambil satuan yang benar (Riyal/Riyalat/Riyalan) menggunakan 2 digit terakhir
        HasilUtama = TeksB & " " & GetMa_dud(val(Right(BagianBulat, 2)), TeksMataUang)
    End If
    
    ' --- BAGIAN DESIMAL ---
    If val(BagianDesimal) > 0 And TeksSen <> "" Then
        Dim TeksD As String
        ' Analog dengan bagian bulat, jalankan logika Idhafah
        TeksD = AngkaKeTeksArab(BagianDesimal, 1, GenderKecil, Modern, False, Marfu)
        
        HasilUtama = HasilUtama & WAW & TeksD & " " & GetMa_dud(val(BagianDesimal), TeksSen)
    End If
    
    ArabCurrency = Trim(HasilUtama)
End Function

Public Function ArabUniversal(ByVal angka As Variant, _
                              Optional ByVal JenisSatuan As Integer = 0, _
                              Optional ByVal PilihGender As Integer = 1, _
                              Optional ByVal IsOrdinal As Boolean = False) As String
Attribute ArabUniversal.VB_Description = "Konverter Angka Arab dengan Satuan Otomatis\r\n------------------------------------------\r\nMendukung perubahan otomatis bentuk Tunggal (Mufrad), \r\nDua (Muthanna), dan Jamak sesuai kaidah bahasa Arab."
Attribute ArabUniversal.VB_ProcData.VB_Invoke_Func = " \n19"

    ' --- Proteksi Error & Inisialisasi ---
    If IsError(angka) Then ArabUniversal = "[Khatha']": Exit Function
    If Not IsNumeric(angka) And Not IsOrdinal Then ArabUniversal = "[Khatha']": Exit Function

    InitializeArabicWords
    
    Dim ALI As String, baa As String, taa As String, THA As String
    Dim jim As String, haa As String, kha As String, dal As String
    Dim rho As String, sin As String, syi As String, sho As String
    Dim ain As String, fha As String, kaf As String, lam As String
    Dim MIM As String, NUN As String, WAW As String, YAA As String
    Dim ham As String, MAR As String, maq As String, gha As String
    Dim thoo As String, raa As String, dho As String, qof As String
    Dim zaa As String, dhal As String, hza As String

    ALI = ChrW(1575): baa = ChrW(1576): taa = ChrW(1578): THA = ChrW(1579)
    jim = ChrW(1580): haa = ChrW(1581): kha = ChrW(1582): dal = ChrW(1583)
    rho = ChrW(1585): sin = ChrW(1587): syi = ChrW(1588): sho = ChrW(1589)
    ain = ChrW(1593): fha = ChrW(1601): kaf = ChrW(1603): lam = ChrW(1604)
    MIM = ChrW(1605): NUN = ChrW(1606): WAW = ChrW(1608): YAA = ChrW(1610)
    ham = ChrW(1571): MAR = ChrW(1577): maq = ChrW(1609): gha = ChrW(1594)
    thoo = ChrW(1591): raa = ChrW(1585): dho = ChrW(1590): qof = ChrW(1602)
    zaa = ChrW(1586): dhal = ChrW(1584): hza = ChrW(1569)

    Dim hasilTerbilang As String, TeksSatuan As String
    Dim AL As String: AL = ALI & lam
    Dim GenderFinal As GenderArab: GenderFinal = IIf(PilihGender = 2, Muannas, Muzakkar)
    Dim v As Long: v = Fix(Abs(val(angka))) Mod 100

    Select Case JenisSatuan
        ' === KELOMPOK 1: JARAK & UKURAN ===
        Case 1: TeksSatuan = sin & NUN & taa & MIM & taa & rho: GenderFinal = 1 ' CM
        Case 2: ' Meter (Muzakkar)
            If angka = 2 Then: ArabUniversal = MIM & taa & raa & ALI & NUN: Exit Function
            If v >= 3 And v <= 10 Then
                TeksSatuan = ham & MIM & taa & ALI & rho
                GenderFinal = 2
            Else
                TeksSatuan = MIM & taa & raa
                GenderFinal = 1
            End If
        Case 3: TeksSatuan = kaf & YAA & lam & WAW & MIM & taa & raa: GenderFinal = 1 ' KM

        ' === KELOMPOK 2: BERAT & VOLUME ===
        Case 4: TeksSatuan = gha & raa & ALI & MIM               ' Gram
        Case 5: TeksSatuan = kaf & YAA & lam & WAW & gha & raa & ALI & MIM ' KG
        Case 6: ' Ton
            If angka = 2 Then: ArabUniversal = thoo & NUN & ALI & NUN
            TeksSatuan = thoo & NUN: GenderFinal = 1
        Case 7: TeksSatuan = lam & taa & raa                     ' Liter

        ' === KELOMPOK 3: WAKTU (Muthanna & Jamak) ===
        Case 10: ' Jam (Sa'ah)
            If angka = 2 Then
                TeksSatuan = sin & ALI & ain & taa & ALI & NUN: ArabUniversal = TeksSatuan: Exit Function
            ElseIf v >= 3 And v <= 10 Then
                TeksSatuan = sin & ALI & ain & ALI & taa: GenderFinal = Muzakkar
            Else
                TeksSatuan = sin & ALI & ain & MAR: GenderFinal = Muannas
            End If
        Case 11: ' Hari (Yaum)
            If angka = 2 Then
                TeksSatuan = YAA & WAW & MIM & ALI & NUN: ArabUniversal = TeksSatuan: Exit Function
            ElseIf v >= 3 And v <= 10 Then
                TeksSatuan = ham & YAA & ALI & MIM: GenderFinal = Muannas
            Else
                TeksSatuan = YAA & WAW & MIM: GenderFinal = Muzakkar
            End If
        Case 12: ' Bulan (Syahr)
            If angka = 2 Then
                TeksSatuan = syi & haa & rho & ALI & NUN: ArabUniversal = TeksSatuan: Exit Function
            ElseIf v >= 3 And v <= 10 Then
                TeksSatuan = ham & syi & haa & rho: GenderFinal = Muannas
            Else
                TeksSatuan = syi & haa & rho: GenderFinal = Muzakkar
            End If
        Case 13: ' Tahun (Sanah)
            If angka = 2 Then
                TeksSatuan = sin & NUN & taa & ALI & NUN: ArabUniversal = TeksSatuan: Exit Function
            ElseIf v >= 3 And v <= 10 Then
                TeksSatuan = sin & NUN & WAW & ALI & taa: GenderFinal = Muzakkar
            Else
                TeksSatuan = sin & NUN & MAR: GenderFinal = Muannas
            End If

        ' === KELOMPOK 4: SOSIAL & LITERASI ===
        Case 20: ' Orang (Syakhsh)
            If angka = 1 Then
                TeksSatuan = syi & kha & sho: ArabUniversal = TeksSatuan: Exit Function
            ElseIf angka = 2 Then
                TeksSatuan = syi & kha & sho & ALI & NUN: ArabUniversal = TeksSatuan: Exit Function
            ElseIf v >= 3 And v <= 10 Then
                TeksSatuan = ham & syi & kha & ALI & sho: GenderFinal = Muannas
            Else
                TeksSatuan = syi & kha & sho: GenderFinal = Muzakkar
            End If
        Case 21: ' Buku (Kitab)
            If angka = 2 Then
                TeksSatuan = kaf & taa & ALI & baa & ALI & NUN: ArabUniversal = TeksSatuan: Exit Function
            ElseIf v >= 3 And v <= 10 Then
                TeksSatuan = kaf & taa & baa: GenderFinal = Muannas
            Else
                TeksSatuan = kaf & taa & ALI & baa: GenderFinal = Muzakkar
            End If
        Case 22: ' Halaman (Safhah)
            If angka = 2 Then
                TeksSatuan = sho & fha & haa & taa & ALI & NUN: ArabUniversal = TeksSatuan: Exit Function
            ElseIf v >= 3 And v <= 10 Then
                TeksSatuan = sho & fha & haa & ALI & taa: GenderFinal = Muzakkar
            Else
                TeksSatuan = sho & fha & haa & MAR: GenderFinal = Muannas
            End If

        ' === KELOMPOK 5: FISIKA & ELEKTRONIKA ===
        Case 40: TeksSatuan = fha & WAW & lam & taa              ' Volt
        Case 41: TeksSatuan = WAW & ALI & taa                    ' Watt
        Case 42: TeksSatuan = ham & MIM & baa & YAA & rho        ' Ampere
        Case 43: TeksSatuan = dho & gha & thoo                   ' Press (Daghth)

        ' === KELOMPOK 6: SUHU, PERSEN & DATA ===
        Case 50: ' Derajat (Darajah - Muannas)
            If angka = 2 Then: ArabUniversal = dal & rho & jim & taa & ALI & NUN: Exit Function
            If v >= 3 And v <= 10 Then
                TeksSatuan = dal & rho & jim & ALI & taa
                GenderFinal = 1
            Else
                TeksSatuan = dal & rho & jim & MAR
                GenderFinal = 2
            End If
        Case 51: ' Persen (Fil-Miah)
        TeksSatuan = fha & YAA & " " & AL & MIM & ALI & ChrW(1574) & MAR: GenderFinal = 1 ' %
        Case 52: TeksSatuan = baa & ALI & YAA & taa: GenderFinal = 1 ' Byte
        Case 53: TeksSatuan = MIM & YAA & gha & ALI & baa & ALI & YAA & taa: GenderFinal = 1  ' GB

        ' === KELOMPOK 7: POSISI & GEOGRAFI ===
        Case 60: TeksSatuan = AL & fha & sho & lam: GenderFinal = 1 ' Bab
        Case 61: TeksSatuan = AL & MIM & raa & taa & baa & MAR: GenderFinal = 2 ' Peringkat
        Case 62: TeksSatuan = AL & fha & raa & YAA & qof: GenderFinal = 1 ' Tim
        Case 63: TeksSatuan = kha & thoo & " " & AL & ALI & sin & taa & WAW & ALI & hza: GenderFinal = 2 ' Khatulistiwa

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
On Error Resume Next
    Dim ArgHelp(0 To 7) As String
    Dim DeskripsiFungsi As String
    
    ' --- TERBILANG_ARAB ---
    ArgHelp(0) = "Angka atau referensi sel (Mendukung angka hingga 65 digit)."
    ArgHelp(1) = "Mode: umum urutan eja uang jarak meter waktu buku lainnya."
    ArgHelp(2) = "Gender: m (Laki-laki) | f (Perempuan)."
    ArgHelp(3) = "I'rab: u (Marfu - Subjek/Predikat) | a (Mansub - Objek) | i (Majrur - Setelah Preposisi/Idhafah)."
    ArgHelp(4) = "Gaya: modern (Standar) | klasik (Wawu sambung) | sastra (Kecil ke besar)."
    ArgHelp(5) = "Parameter Tambahan: ID Negara (id, sa) jika mode 'uang', atau ID Benda (1-63)."
    ArgHelp(6) = "Harakat: True (Dengan Harakat), False (Polos)."
    ArgHelp(7) = "Idhafah: True (Gunakan aturan penyambungan kata benda), False (Tanpa idhafah)."

    DeskripsiFungsi = "Konverter Angka ke Bahasa Arab (Tafqit)" & vbCrLf & _
                      "------------------------------------------" & vbCrLf & _
                      "Dev: Rida Rahman DH 96-02" & vbCrLf & _
                      "GitHub: https://github.com/Rida-Haja" & vbCrLf & _
                      "Mengguakan Skala Panjang (Long Scale)" & vbCrLf & _
                      "Fitur: Grammar, Gaya Tulisan, & Mata Uang."
                      
    Application.MacroOptions Macro:="TERBILANG_ARAB", _
        Description:=DeskripsiFungsi, _
        Category:="HajaArabic (VBA Terbilang Arab)", _
        ArgumentDescriptions:=ArgHelp
        
'=====================================================================================

    Dim ArabUniversal(0 To 3) As String
    
    ' --- ARAB_UNIVERSAL ---
    ArabUniversal(0) = "Angka atau referensi sel (Input nilai numerik)."
    ArabUniversal(1) = "Jenis Satuan: ID Benda (1-63) seperti Meter, Jam, Hari, Bab, dll."
    ArabUniversal(2) = "Pilih Gender: 1 (Muzakkar/Laki-laki) | 2 (Muannas/Perempuan)."
    ArabUniversal(3) = "Ordinal: True (Angka Urutan: ke-1, ke-2) | False (Angka Biasa)."

    DeskripsiFungsi = "Konverter Angka Arab dengan Satuan Otomatis" & vbCrLf & _
                      "------------------------------------------" & vbCrLf & _
                      "Mendukung perubahan otomatis bentuk Tunggal (Mufrad), " & vbCrLf & _
                      "Dua (Muthanna), dan Jamak sesuai kaidah bahasa Arab."
                      
    Application.MacroOptions Macro:="ArabUniversal", _
        Description:=DeskripsiFungsi, _
        Category:="HajaArabic (VBA Terbilang Arab)", _
        ArgumentDescriptions:=ArabUniversal
        
'=====================================================================================

    Dim ArabCurrency(0 To 1) As String
    
    ' --- ARAB_CURRENCY ---
    ArabCurrency(0) = "Angka nominal uang yang akan dikonversi."
    ArabCurrency(1) = "ID Negara: 'id' (Rupiah), 'sa' (Riyal), 'kw' (Dinar), 'ae' (Dirham), dll."

    DeskripsiFungsi = "Konverter Nominal Uang ke Teks Bahasa Arab" & vbCrLf & _
                      "---------------------------------------------" & vbCrLf & _
                      "Menampilkan satuan mata uang dan pecahan (sen/halala) " & vbCrLf & _
                      "secara otomatis sesuai negara yang dipilih."
                      
    Application.MacroOptions Macro:="ArabCurrency", _
        Description:=DeskripsiFungsi, _
        Category:="HajaArabic (VBA Terbilang Arab)", _
        ArgumentDescriptions:=ArabCurrency
        
    If Err.Number = 0 Then
        MsgBox "Sistem HajaArabic (VBA Terbilang Arab) berhasil didaftarkan ke Excel!", vbInformation, "Sukses"
    Else
        MsgBox "Terjadi kesalahan saat pendaftaran.", vbCritical, "Error"
    End If
End Sub

Public Function TERBILANG_ARAB(ByVal angka As Variant, _
    Optional ByVal mode As String = "umum", _
    Optional ByVal gender As String = "m", _
    Optional ByVal irab As String = "u", _
    Optional ByVal gaya As String = "modern", _
    Optional ByVal Parameter_Negara As String = "id", _
    Optional ByVal PakaiHarakat As Boolean = False, _
    Optional ByVal isIdhafah As Boolean = False) As String
Attribute TERBILANG_ARAB.VB_Description = "Konverter Angka ke Bahasa Arab (Tafqit)\r\n------------------------------------------\r\nDev: Rida Rahman DH 96-02\r\nGitHub: https://github.com/Rida-Haja\r\nMengguakan Skala Panjang (Long Scale)\r\nFitur: Grammar, Gaya Tulisan, & Mata Uang."
Attribute TERBILANG_ARAB.VB_ProcData.VB_Invoke_Func = " \n19"

    On Error GoTo ErrHandler
    Application.Volatile
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
        Case "klasik", "2": enGaya = 2 ' Klasik
        Case "sastra", "3": enGaya = 3 ' Sastra
        Case Else: enGaya = 1          ' Modern (Default)
    End Select

' --- 3. Eksekusi Berdasarkan Mode & Kelompok Satuan ---
    Dim HasilSementara As String
    Dim modeLcase As String: modeLcase = LCase(Trim(mode))
    Dim idUnit As Integer: idUnit = 0 ' Default: Tanpa satuan

    Select Case modeLcase
        ' -- [KELOMPOK DASAR] --
        Case "umum", "angka", "kardinal"
            HasilSementara = AngkaKeTeksArab(angka, 1, enGender, enGaya, False, enIrab, isIdhafah)

        Case "eja", "digit"
            HasilSementara = AngkaKeTeksArab(angka, 2, enGender, enGaya, False, enIrab, isIdhafah)

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
        HasilSementara = AngkaKeTeksArab(angka, 1, enGender, enGaya, False, enIrab, isIdhafah)
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

' ================= Harakat =========================================================
Private Function MapKataHarakat(ByVal w As String, ByVal irab As Integer) As String
    ' === DEFINISI HARAKAT ===
    Dim f As String: f = ChrW(1614):  Dim k As String: k = ChrW(1616)
    Dim d As String: d = ChrW(1615):  Dim s As String: s = ChrW(1618)
    Dim sy As String: sy = ChrW(1617): Dim dt As String: dt = ChrW(1612)
    
    ' === DEFINISI HURUF ===
    Dim ALI As String: ALI = ChrW(1575): Dim baa As String: baa = ChrW(1576)
    Dim taa As String: taa = ChrW(1578): Dim THA As String: THA = ChrW(1579)
    Dim jim As String: jim = ChrW(1580): Dim haa As String: haa = ChrW(1581)
    Dim kha As String: kha = ChrW(1582): Dim dal As String: dal = ChrW(1583)
    Dim rho As String: rho = ChrW(1585): Dim sin As String: sin = ChrW(1587)
    Dim syi As String: syi = ChrW(1588): Dim ain As String: ain = ChrW(1593)
    Dim fha As String: fha = ChrW(1601): Dim kaf As String: kaf = ChrW(1603)
    Dim lam As String: lam = ChrW(1604): Dim MIM As String: MIM = ChrW(1605)
    Dim NUN As String: NUN = ChrW(1606): Dim WAW As String: WAW = ChrW(1608)
    Dim YAA As String: YAA = ChrW(1610): Dim ham As String: ham = ChrW(1571)
    Dim MAR As String: MAR = ChrW(1577): Dim maq As String: maq = ChrW(1609)
    Dim hza As String: hza = ChrW(1574): Dim sho As String: sho = ChrW(1589)
    Dim AL As String: AL = ALI & lam
    Dim eL As String: eL = ALI & lam & s

    ' --- A. LOGIKA DINAMIS PULUHAN (20-90) ---
    If Len(w) >= 4 Then
        Dim akhir As String: akhir = Right(w, 2)
        ' Deteksi uuna (Marfu) atau iina (Mansub/Majrur)
        If akhir = WAW & NUN Or akhir = YAA & NUN Then
            Dim akar As String: akar = Left(w, Len(w) - 2)
            Dim hAkhir As String
            ' Jika I'rab Marfu (1) pakai dhommah+waw, jika Mansub/Majrur (2/3) pakai kasrah+ya
            hAkhir = IIf(irab = 1, d & WAW & NUN & f, k & YAA & s & NUN & f)
            
            Select Case akar
                Case ain & syi & rho: MapKataHarakat = ain & k & syi & s & rho & hAkhir: Exit Function ' 20
                Case THA & lam & ALI & THA: MapKataHarakat = THA & f & lam & f & ALI & THA & hAkhir: Exit Function ' 30
                Case ham & rho & baa & ain: MapKataHarakat = ham & f & rho & s & baa & f & ain & hAkhir: Exit Function ' 40
                Case kha & MIM & sin: MapKataHarakat = kha & f & MIM & s & sin & hAkhir: Exit Function ' 50
                Case sin & taa: MapKataHarakat = sin & k & taa & sy & hAkhir: Exit Function ' 60
                Case sin & baa & ain: MapKataHarakat = sin & f & baa & s & ain & hAkhir: Exit Function ' 70
                Case THA & MIM & ALI & NUN: MapKataHarakat = THA & f & MIM & f & ALI & NUN & hAkhir: Exit Function ' 80
                Case taa & sin & ain: MapKataHarakat = taa & k & sin & s & ain & hAkhir: Exit Function ' 90
            End Select
        End If
    End If

    ' --- B. LOGIKA DUAL (Angka 2 / 12) ---
    If w = ALI & THA & NUN & ALI Then
        MapKataHarakat = IIf(irab = 1, ALI & THA & s & NUN & f & ALI, ALI & k & THA & s & NUN & f & YAA & s)
        Exit Function
    End If

    ' --- C. LOGIKA KATA STATIS (Satuan, Belasan, Ribuan) ---
    Select Case w
        'Nol
        Case sho & fha & rho: MapKataHarakat = sho & k & fha & s & rho
        '100
        Case ChrW(1605) & ALI & ChrW(1574) & MAR: MapKataHarakat = ChrW(1605) & k & ALI & ChrW(1574) & f & MAR
        Case ChrW(1605) & ChrW(1574) & MAR: MapKataHarakat = ChrW(1605) & k & ChrW(1574) & f & MAR

        ' Satuan Muzakkar
        Case WAW & ALI & haa & dal: MapKataHarakat = WAW & f & ALI & haa & k & dal '& dt
        Case THA & lam & ALI & THA: MapKataHarakat = THA & f & lam & f & ALI & THA & d
        Case ham & rho & baa & ain: MapKataHarakat = ham & f & rho & s & baa & f & ain & d
        Case kha & MIM & sin:       MapKataHarakat = kha & f & MIM & s & sin & d
        Case sin & taa:             MapKataHarakat = sin & k & taa & sy & d
        Case sin & baa & ain:       MapKataHarakat = sin & f & baa & s & ain & d
        Case THA & MIM & ALI & NUN: MapKataHarakat = THA & f & MIM & f & ALI & NUN & d
        Case taa & sin & ain:       MapKataHarakat = taa & k & sin & s & ain & d

        ' Satuan Muannas
        Case WAW & ALI & haa & dal & MAR: MapKataHarakat = WAW & f & ALI & haa & k & dal & f & MAR '& dt
        Case THA & lam & ALI & THA & MAR: MapKataHarakat = THA & f & lam & f & ALI & THA & f & MAR '& dt
        Case ham & rho & baa & ain & MAR: MapKataHarakat = ham & f & rho & s & baa & f & ain & f & MAR '& dt
        Case kha & MIM & sin & MAR:       MapKataHarakat = kha & f & MIM & s & sin & f & MAR '& dt
        Case sin & taa & MAR:             MapKataHarakat = sin & k & taa & sy & f & MAR '& dt
        Case sin & baa & ain & MAR:       MapKataHarakat = sin & f & baa & s & ain & f & MAR '& dt
        Case THA & MIM & ALI & NUN & YAA & MAR: MapKataHarakat = THA & f & MIM & f & ALI & NUN & k & YAA & f & MAR '& dt
        Case taa & sin & ain & MAR:       MapKataHarakat = taa & k & sin & s & ain & f & MAR '& dt

        ' Belasan
        Case ain & syi & rho:       MapKataHarakat = ain & f & syi & f & rho & f
        Case ain & syi & rho & MAR: MapKataHarakat = ain & f & syi & s & rho & f & MAR & f
        Case ham & haa & dal:       MapKataHarakat = ham & f & haa & f & dal & f
        Case ChrW(1573) & haa & dal & maq: MapKataHarakat = ChrW(1573) & k & haa & s & dal & f & maq

        ' Ratusan, Ribuan, Jutaan
        Case MIM & hza & ALI & MAR: MapKataHarakat = MIM & k & hza & f & MAR '& dt
        Case ham & lam & fha:       MapKataHarakat = ham & f & lam & s & fha '& dt ' Alfun
        Case ALI & lam & ALI & fha: MapKataHarakat = ALI & f & lam & f & ALI & f & fha '& dt ' Alaf
        Case MIM & lam & YAA & WAW & NUN: MapKataHarakat = MIM & k & lam & s & YAA & d & WAW & NUN '& dt ' Milyun
        Case MIM & lam & ALI & YAA & YAA & NUN: MapKataHarakat = MIM & f & lam & f & ALI & YAA & k & YAA & f & NUN & k ' Malayiin
        Case ALI & lam & ALI & fha: MapKataHarakat = ALI & f & lam & f & ALI & f & fha & k ' Alaf (Majrur)

        ' Partikel
        Case WAW: MapKataHarakat = WAW & f
          
          ' Indonesia: Rupiah
        Case rho & WAW & baa & YAA & MAR:        MapKataHarakat = rho & d & WAW & s & baa & k & YAA & sy & f & MAR
        Case rho & WAW & baa & YAA & ALI & taa:        MapKataHarakat = rho & d & WAW & s & baa & k & YAA & sy & f & ALI & taa
        
         ' Jam (Sa'ah)
        Case sin & ALI & ain & taa & ALI & NUN:        MapKataHarakat = sin & f & ALI & ain & f & taa & f & ALI & NUN
        Case sin & ALI & ain & ALI & taa:        MapKataHarakat = sin & f & ALI & ain & f & ALI & taa
        Case sin & ALI & ain & MAR:        MapKataHarakat = sin & f & ALI & ain & f & MAR
         ' Hari (Yaum)
        Case YAA & WAW & MIM & ALI & NUN:         MapKataHarakat = YAA & f & WAW & MIM & f & ALI & NUN
        Case ham & YAA & ALI & MIM:        MapKataHarakat = ham & f & YAA & sy & f & ALI & MIM
        Case YAA & WAW & MIM:        MapKataHarakat = YAA & f & WAW & s & MIM
         'Bulan (Syahr)
        Case syi & haa & rho & ALI & NUN:        MapKataHarakat = syi & f & haa & rho & f & ALI & NUN
        Case ham & syi & haa & rho:        MapKataHarakat = ham & f & syi & s & haa & d & rho
        Case syi & haa & rho:        MapKataHarakat = syi & f & haa & rho
        ' Tahun (Sanah)
        Case sin & NUN & taa & ALI & NUN:        MapKataHarakat = sin & f & NUN & f & taa & f & ALI & NUN
        Case sin & NUN & WAW & ALI & taa:        MapKataHarakat = sin & f & NUN & f & WAW & f & ALI & taa
        Case sin & NUN & MAR:        MapKataHarakat = sin & f & NUN & f & MAR

        
        Case AL & fha & sho & lam:         MapKataHarakat = eL & fha & f & sho & d & lam 'Kelas
        Case AL & MIM & rho & taa & baa & MAR:      MapKataHarakat = eL & MIM & f & rho & d & taa & f & baa & f & MAR ' Peringkat

        
        Case Else: MapKataHarakat = w
    End Select
End Function

' --- FUNGSI PEMBANTU ---
Private Function HapusHarakat(ByVal teks As String) As String
    Dim i As Long, hasil As String, c As Long
    For i = 1 To Len(teks)
        c = AscW(Mid(teks, i, 1))
        ' Abaikan karakter Unicode harakat (Fathah, Dammah, Kasrah, dll)
        If c < 1611 Or c > 1618 Then hasil = hasil & Mid(teks, i, 1)
    Next i
    HapusHarakat = hasil
End Function

Private Function BeriHarakat(ByVal teks As String, ByVal irab As Integer) As String
    Dim kata() As String, i As Integer, hasil() As String
    Dim WAW As String: WAW = ChrW(1608)
    
    kata = Split(teks, " ")
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



Public Sub JalankanUjiLogika()
    InitializeArabicWords
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Sheet2")
    ws.Cells.Clear
    
    ' Header Tabel
    Dim headers As Variant
    'headers = Array("Angka", "Konteks Uji", "Gaya", "Gender", "I'rab", "Hasil Konversi (Arab)")
    headers = Array("Angka", "Konteks Uji", "Hasil Konversi (Arab)")
    ws.Range("A1:C1").Value = headers
    ws.Range("A1:C1").Font.Bold = True
    
    Dim r As Integer: r = 2
    
    ' --- 1. KOMPARASI GAYA (100 & 800) ---
    TambahBarisFull ws, r, 100, "100 Gaya Modern", Modern, Muzakkar, Marfu
    TambahBarisFull ws, r, 100, "100 Gaya Klasik", Klasik, Muzakkar, Marfu
    TambahBarisFull ws, r, 800, "800 Gaya Modern", Modern, Muzakkar, Marfu
    TambahBarisFull ws, r, 800, "800 Gaya Sastra", Sastra, Muzakkar, Marfu
    TambahBarisFull ws, r, 800, "800 Gaya Klasik", Klasik, Muzakkar, Marfu

    ' --- 2. MASALAH ANGKA "2" ---
    TambahBarisFull ws, r, 2, "Satuan 2", Modern, Muzakkar, Marfu
    TambahBarisFull ws, r, 12, "Belasan 12 (Nun Hilang)", Modern, Muannas, Marfu
    TambahBarisFull ws, r, 200, "Ratusan 200 Pas ", Modern, Muzakkar, Marfu
    TambahBarisFull ws, r, 1202, "Ratusan 1202 Idhafah true ", Modern, Muzakkar, Marfu, True
    TambahBarisFull ws, r, 1202, "Ratusan 1202 Idhafah false ", Modern, Muzakkar, Marfu, False
    TambahBarisFull ws, r, 2000, "Ribuan 2000 Pas", Modern, Muzakkar, Marfu
    TambahBarisFull ws, r, 2000, "Ribuan 2000 Majrur", Modern, Muzakkar, Majrur

    ' --- 3. MASALAH ANGKA 8 ---
    TambahBarisFull ws, r, 8, "8 Muzakkar", Modern, Muzakkar, Marfu
    TambahBarisFull ws, r, 8, "8 Muannas", Modern, Muannas, Marfu
    TambahBarisFull ws, r, 8, "8 Muzakkar Idhafah true", Modern, Muzakkar, Marfu, True
    
    ' --- 4. KOMBINASI PULUHAN & SATUAN ---
    TambahBarisFull ws, r, 21, "21 Marfu", Modern, Muzakkar, Marfu
    TambahBarisFull ws, r, 99, "99 Max Unit", Modern, Muannas, Marfu

    ' --- 5. ANGKA "MONSTER" (Skala Besar) ---
    ' Gunakan string untuk angka yang sangat besar agar tidak rusak di Excel/
    ' Gunakan tanda petik dua agar dianggap String murni oleh VBA
    TambahBarisFull ws, r, "5000", "Jamak 5000 (Aalaaf)", Modern, Muzakkar, Marfu
    TambahBarisFull ws, r, "1000000", "1 Juta", Modern, Muzakkar, Marfu
    TambahBarisFull ws, r, "3000000", "3 Juta (Malayiin)", Modern, Muzakkar, Marfu
    TambahBarisFull ws, r, "12200308", "Modern", Modern, Muzakkar, Marfu
    TambahBarisFull ws, r, "12200308", "Sastra Marfu(Kecil-Besar)", Sastra, Muzakkar, Marfu
    TambahBarisFull ws, r, "12200308", "Sastra Mansub(Kecil-Besar)", Sastra, Muzakkar, Mansub
    TambahBarisFull ws, r, "1000000000000", "1 Bilyun (Triliun Skala Panjang)", Modern, Muzakkar, Marfu
    TambahBarisFull ws, r, "1212121212202121211", "20 Digit Test", Klasik, Muzakkar, Mansub

' Test 66 Digit (Vigintillion)
Dim monster66 As String
monster66 = "12121212122021212121000000000000000000000000001212121212121210000"
TambahBarisFull ws, r, monster66, "Decilyar Test", Klasik, Muzakkar, Mansub    ' Formatting
    ws.Columns("A:C").AutoFit
    ws.Columns("C:C").ColumnWidth = 55 ' Beri ruang lebih untuk teks Arab
    ws.Columns("C").HorizontalAlignment = xlRight
    ws.Columns("C").Font.Name = "Arial Italic" ' Opsional: Font yang mendukung Unicode
    ws.Columns("C").Font.Size = 14
    
    MsgBox "Uji logika 'Abjadun' Selesai!", vbInformation
End Sub

Private Sub TambahBarisFull(ws As Worksheet, ByRef r As Integer, ByVal num As Variant, _
                            ByVal konteks As String, ByVal gya As GayaArab, _
                            ByVal gen As GenderArab, ByVal irb As IrabArab, Optional isIdhafah As Boolean)
    
    ' PENTING: Paksa menjadi string murni untuk menghindari scientific notation
    Dim strNum As String
    strNum = Trim(CStr(num))
    
    ' Masukkan ke Excel sebagai teks murni
    ws.Cells(r, 1).Value = "'" & strNum
    ws.Cells(r, 2).Value = konteks
    
    ' ... (Logika Tabel) ...

    On Error Resume Next
    ' Eksekusi konversi dengan parameter String murni
    ws.Cells(r, 3).Value = ProsesBlokRibuan(strNum, gen, gya, irb, isIdhafah)
    
    If Err.Number <> 0 Then
        ws.Cells(r, 3).Value = "Error: " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0
    
    r = r + 1
End Sub



