VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTerbilangArab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Rem ==============================================================================
Rem  HajaArabic (VBA Terbilang Arab) v.001a beta
Rem ==============================================================================
Rem  AUTHOR      : Rida Haja (Rida Rahman DH 96-02)
Rem  LICENSE     : MIT License
Rem  COPYRIGHT   : (c) 2026 Rida Rahman DH 96-02
Rem  URL         : https://github.com/Rida-Haja/VbaTerbilangArab
Rem  EMAIL       : ridahaja@gmail.com
Rem  DATE        : Februari 2026
Rem  SPECS       : Supports up to 1000-digit conversion
Rem                with Full Harakat & I'rab logic.
Rem ==============================================================================
Rem
Rem ==============================================================================
Rem LICENSE / LISENSI
Rem ==============================================================================
Rem
Rem [ENGLISH]
Rem
Rem Permission is hereby granted, free of charge, to any person obtaining a copy of this software
Rem and associated documentation files (the "HajaArabic"), to deal in the Software without restriction,
Rem including without limitation the rights to use, copy, modify, merge, publish, distribute,
Rem sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is
Rem furnished to do so, subject to the following conditions:
Rem
Rem The above copyright notice and this permission notice shall be included in all copies or
Rem substantial portions of the Software.
Rem
Rem THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING
Rem BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
Rem NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM,
Rem DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
Rem OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
Rem
Rem ------------------------------------------------------------------------------
Rem [BAHASA INDONESIA]
Rem
Rem Izin diberikan secara gratis kepada siapa pun yang mendapatkan salinan perangkat lunak ini
Rem untuk menggunakan, menyalin, mengubah, dan mendistribusikan tanpa batasan, dengan syarat
Rem pemberitahuan hak cipta di atas wajib disertakan dalam semua salinan perangkat lunak.
Rem Perangkat lunak ini disediakan "APA ADANYA", tanpa jaminan apa pun dari penulis.
Rem ==============================================================================

' --- Private Variables untuk Karakter ---
Public ALI As String, baa As String, TAA As String, THA As String, jim As String, haa As String
Public kha As String, dal As String, DZA As String, RHO As String, ZAI As String, sin As String
Public syi As String, sho As String, DHA As String, THO As String, ZHO As String, AIN As String
Public GHA As String, FHA As String, QAF As String, kaf As String, LAM As String, mim As String
Public NUN As String, WAW As String, HAA_besar As String, YAA As String
Public HAM As String, MAR As String, MAQ As String, HZA As String

' --- Private Arrays untuk Data Kata ---
Private pKata(0 To 11, 1 To 2) As String
Private pPuluhan(2 To 9) As String
Private pSatuanBesar(1 To 333) As String
Private pPilihanSkala As Integer

' --- Status Inisialisasi ---
Private pIsInitialized As Boolean

' ==============================================================================
' CLASS INITIALIZER (Otomatis jalan saat class dibuat)
' ==============================================================================
Private Sub Class_Initialize()
    InitializeData
End Sub

' ==============================================================================
' PROPERTY GETTERS (Untuk mengakses data dari luar class)
' ==============================================================================
Public Property Get Kata(ByVal Angka As Integer, ByVal gender As Integer) As String
    Kata = pKata(Angka, gender)
End Property

Public Property Get Puluhan(ByVal indeks As Integer) As String
    Puluhan = pPuluhan(indeks)
End Property

Public Property Get SatuanBesar(ByVal indeks As Integer) As String
    SatuanBesar = pSatuanBesar(indeks)
End Property

Public Property Get IsInitialized() As Boolean
    IsInitialized = pIsInitialized
End Property

' ==============================================================================
' METODE INTERNAL: INISIALISASI DATA
' ==============================================================================
Private Sub InitializeData()
    If pIsInitialized Then Exit Sub
    
    Dim i As Integer, V As Variant
    
    ' --- SET UNICODE DASAR ---
    V = Array(1575, 1576, 1578, 1579, 1580, 1581, 1582, 1583, 1584, 1585, 1586, 1587, _
              1588, 1589, 1590, 1591, 1592, 1593, 1594, 1601, 1602, 1603, 1604, 1605, _
              1606, 1608, 1607, 1610, 1571, 1577, 1609, 1574)
              
    For i = 0 To UBound(V): V(i) = ChrW(V(i)): Next i

    ALI = V(0):  baa = V(1):  TAA = V(2):  THA = V(3):  jim = V(4):  haa = V(5)
    kha = V(6):  dal = V(7):  DZA = V(8):  RHO = V(9):  ZAI = V(10): sin = V(11)
    syi = V(12): sho = V(13): DHA = V(14): THO = V(15): ZHO = V(16): AIN = V(17)
    GHA = V(18): FHA = V(19): QAF = V(20): kaf = V(21): LAM = V(22): mim = V(23)
    NUN = V(24): WAW = V(25): HAA_besar = V(26): YAA = V(27): HAM = V(28): MAR = V(29): MAQ = V(30): HZA = V(31)
    
    ' --- 0: Sifrun ---
    pKata(0, 1) = sho & FHA & RHO
    pKata(0, 2) = pKata(0, 1)

    ' --- 1-9: Satuan Dasar ---
    pKata(1, 1) = WAW & ALI & haa & dal
    pKata(2, 1) = ALI & THA & NUN & ALI & NUN
    pKata(3, 1) = THA & LAM & ALI & THA
    pKata(4, 1) = HAM & RHO & baa & AIN
    pKata(5, 1) = kha & mim & sin
    pKata(6, 1) = sin & TAA
    pKata(7, 1) = sin & baa & AIN
    pKata(8, 1) = THA & mim & ALI & NUN & YAA
    pKata(9, 1) = TAA & sin & AIN

    ' Muannas
    pKata(1, 2) = pKata(1, 1) & MAR
    pKata(2, 2) = ALI & THA & NUN & TAA & ALI & NUN
    For i = 3 To 9
        pKata(i, 2) = pKata(i, 1) & MAR
    Next i
   
    ' --- 10-11 ---
    pKata(10, 1) = AIN & syi & RHO
    pKata(10, 2) = pKata(10, 1) & MAR
    pKata(11, 1) = HAM & haa & dal
    pKata(11, 2) = ChrW(1573) & haa & dal & MAQ
    
    ' --- 20-90: Puluhan ---
    pPuluhan(2) = pKata(10, 1) & WAW & NUN
    For i = 3 To 9
        If i = 8 Then
            pPuluhan(i) = Left(pKata(i, 1), Len(pKata(i, 1)) - 1) & WAW & NUN
        Else
            pPuluhan(i) = pKata(i, 1) & WAW & NUN
        End If
    Next i

    ' --- SATUAN BESAR ---
    pSatuanBesar(1) = HAM & LAM & FHA
    pSatuanBesar(2) = mim & LAM & YAA & WAW & NUN
    pSatuanBesar(3) = mim & LAM & YAA & ALI & RHO

    ' --- LOGIKA DINAMIS 1000 DIGIT ---
    Dim lyun As String: lyun = LAM & YAA & WAW & NUN
    Dim lyar As String: lyar = LAM & YAA & ALI & RHO
    
    Dim pU: pU = Array("", ALI & NUN, dal & WAW, TAA & RHO & YAA, kaf & WAW & ALI & dal & RHO, _
                        kaf & WAW & YAA & NUN, sin & kaf & sin, sin & baa & TAA, _
                        HAM & kaf & TAA, NUN & WAW & FHA & mim)
                        
    Dim pT: pT = Array("", dal & YAA & sin & YAA, FHA & YAA & jim & NUN & TAA, _
                        TAA & RHO & YAA & jim & NUN & TAA, kaf & WAW & ALI & dal & RHO & ALI & jim & NUN & TAA, _
                        kaf & WAW & YAA & NUN & kaf & WAW & ALI & jim & NUN & TAA, _
                        sin & kaf & sin & ALI & jim & NUN & TAA, sin & baa & TAA & ALI & jim & NUN & TAA, _
                        HAM & kaf & TAA & ALI & jim & NUN & TAA, NUN & WAW & NUN & ALI & jim & NUN & TAA)
                        
    Dim pC: pC = Array("", sin & NUN & TAA & YAA, dal & WAW & sin & NUN & TAA & YAA, _
                        TAA & RHO & YAA & sin & NUN & TAA & YAA, kaf & WAW & ALI & dal & RHO & sin & NUN & TAA & YAA, _
                        kaf & WAW & YAA & NUN & sin & NUN & TAA & YAA, sin & kaf & sin & sin & NUN & TAA & YAA, _
                        sin & baa & TAA & sin & NUN & TAA & YAA, HAM & kaf & TAA & sin & NUN & TAA & YAA, _
                        NUN & WAW & NUN & sin & NUN & TAA & YAA)

    ' Penentuan Skala
    On Error Resume Next
    Dim vSkala As Variant
    vSkala = ThisWorkbook.Worksheets("Sheet1").Cells(6, 3).Value
    pPilihanSkala = IIf(LCase(CStr(vSkala)) = "panjang", 2, 1)
    On Error GoTo 0

    For i = 4 To 333
        Dim prefixVal As Integer
        If pPilihanSkala = 1 Then prefixVal = i - 1 Else prefixVal = i \ 2
        
        Dim uIdx As Integer: uIdx = prefixVal Mod 10
        Dim tIdx As Integer: tIdx = (prefixVal \ 10) Mod 10
        Dim cIdx As Integer: cIdx = (prefixVal \ 100) Mod 10
        
        Dim buildPrefix As String
        If prefixVal = 2 Then
            buildPrefix = baa
        Else
            buildPrefix = pU(uIdx) & pT(tIdx) & pC(cIdx)
        End If
        
        If pPilihanSkala = 1 Then
            pSatuanBesar(i) = buildPrefix & lyun
        Else
            pSatuanBesar(i) = buildPrefix & IIf(i Mod 2 = 0, lyun, lyar)
        End If
    Next i

    ' Overwrite Manual
    If pPilihanSkala = 1 Then
        pSatuanBesar(11) = dal & YAA & sin & YAA & lyun
    End If

    pIsInitialized = True
End Sub

Public Property Get SkalaAktif() As String
    SkalaAktif = IIf(pPilihanSkala = 1, "Skala Pendek", "Skala Panjang")
End Property

' ==============================================================================
' FUNGSI INTI: PENGGABUNG BLOK ANGKA ARAB (TAFQIT)
' ==============================================================================
Public Function Core_GabungAngkaAr(ByVal strNum As Variant, _
                                  ByVal PilihGender As Integer, _
                                  Optional ByVal gaya As Integer = 1, _
                                  Optional ByVal irab As Integer = 1, _
                                  Optional ByVal isIdhafah As Boolean = False) As String
    
    Dim wordList() As String
    Dim blockCount As Integer, i As Integer
    Dim cleanNum As String: cleanNum = ""
    Dim rawInput As String: rawInput = Trim(CStr(strNum))
    
    ' 1. Ekstraksi Angka Saja
    For i = 1 To Len(rawInput)
        If Mid(rawInput, i, 1) >= "0" And Mid(rawInput, i, 1) <= "9" Then
            cleanNum = cleanNum & Mid(rawInput, i, 1)
        End If
    Next i
    
    ' 2. Cek apakah angka valid atau nol
    If cleanNum = "" Or val(cleanNum) = 0 Then
        Core_GabungAngkaAr = pKata(0, 1) ' Menghasilkan Sifrun
        Exit Function
    End If
    
    ' 3. Normalisasi Panjang String (Padding kelipatan 3)
    ' Menambahkan nol di depan agar string bisa dibagi rata per 3 digit (blok)
    Do While Len(cleanNum) Mod 3 <> 0
       cleanNum = "0" & cleanNum
    Loop
    
    blockCount = Len(cleanNum) \ 3
    ReDim wordList(0 To blockCount - 1)

    ' 4. Identifikasi Blok Terakhir yang Diucapkan (Penting untuk Idhafah)
    ' Index 0 = Satuan/Ratusan, Index 1 = Ribuan, dst.
    Dim lastPronouncedIdx As Integer
    
    If gaya = 3 Then ' GAYA SASTRA (Kecil ke Besar)
        ' Cari blok bukan nol pertama dari kanan (satuan)
        For i = 0 To blockCount - 1
            If CLng(Mid(cleanNum, Len(cleanNum) - (i * 3) - 2, 3)) > 0 Then
                lastPronouncedIdx = i
                Exit For
            End If
        Next i
    Else ' GAYA MODERN & KLASIK (Besar ke Kecil)
        ' Cari blok bukan nol pertama dari kiri (skala terbesar)
        For i = blockCount - 1 To 0 Step -1
            If CLng(Mid(cleanNum, Len(cleanNum) - (i * 3) - 2, 3)) > 0 Then
                lastPronouncedIdx = i
                Exit For
            End If
        Next i
    End If

    ' 5. Pemrosesan per Blok (3 Digit)
    For i = 0 To blockCount - 1
        Dim startPos As Integer: startPos = Len(cleanNum) - (i * 3) - 2
        Dim strTemp As String:   strTemp = Mid(cleanNum, startPos, 3)
        Dim vTemp As Long:       vTemp = CLng(strTemp)
        
        If vTemp <> 0 Then
            Dim txtResult As String
            
            ' Idhafah hanya diterapkan pada blok terakhir yang terdengar secara lisan
            Dim kirimIdhafah As Boolean
            kirimIdhafah = (i = lastPronouncedIdx) And isIdhafah
            
            If i = 0 Then
                ' Blok Dasar: Satuan, Belasan, Puluhan, Ratusan (0-999)
                txtResult = Core_SatuanDasarAr(vTemp, PilihGender, gaya, irab, kirimIdhafah)
            Else
                ' Blok Skala Besar: Ribuan, Jutaan, dst.
                txtResult = Core_SatuanBesarAr(vTemp, i, gaya, irab, kirimIdhafah)
            End If
            
            wordList(i) = txtResult
        End If
    Next i
    
    ' 6. Penggabungan Akhir dengan Kata Sambung (Wawu)
    Dim finalResult As String
    Dim junc As String
    
    ' Junc: Modern menggunakan spasi penuh, Klasik/Sastra menempel ke kata berikutnya
    junc = IIf(gaya = 1, " " & WAW & " ", " " & WAW)
    
    If gaya = 3 Then
        ' --- GAYA SASTRA: Urutan Menaik (Satuan -> Ribuan -> Jutaan) ---
        For i = 0 To blockCount - 1
            If wordList(i) <> "" Then
                finalResult = finalResult & IIf(finalResult <> "", junc, "") & wordList(i)
            End If
        Next i
    Else
        ' --- GAYA MODERN & KLASIK: Urutan Menurun (Jutaan -> Ribuan -> Satuan) ---
        For i = blockCount - 1 To 0 Step -1
            If wordList(i) <> "" Then
                finalResult = finalResult & IIf(finalResult <> "", junc, "") & wordList(i)
            End If
        Next i
    End If
    
    Core_GabungAngkaAr = Trim$(finalResult)
End Function

' ==============================================================================
' FUNGSI: KONVERSI SATUAN DASAR (0 - 999)
' ==============================================================================
Private Function Core_SatuanDasarAr(ByVal num As Variant, _
                                     ByVal PilihGender As Integer, _
                                     ByVal gaya As Integer, _
                                     Optional irab As Integer = 1, _
                                     Optional isIdhafah As Boolean = False) As String
    
    Dim s As String, strSatuan As String, strRatusan As String
    Dim V As Long: V = CLng(num)
    
    If V = 0 Then Exit Function
    
    ' Spasi Waw: Gaya Modern dipisah spasi (" ? "), Klasik/Sastra menempel (" ?")
    Dim spWaw As String: spWaw = IIf(gaya = 1, " " & WAW & " ", " " & WAW)
    
    ' --- 1. PROSES RATUSAN (100 - 900) ---
    Dim tRatusklasik As String: tRatusklasik = mim & ALI & HZA
    Dim tRatusModern As String: tRatusModern = mim & HZA
    
    If V >= 100 Then
        Dim H As Integer: H = Int(V \ 100)
        Dim sisaRatusan As Integer: sisaRatusan = V Mod 100
        
        Select Case H
            Case 1 ' 100: Mi'ah
                strRatusan = IIf(gaya = 1, tRatusModern, tRatusklasik) & MAR
                
            Case 2 ' 200: Mi'atani / Mi'atai
                strRatusan = IIf(gaya = 1, tRatusModern, tRatusklasik) & TAA
                strRatusan = strRatusan & IIf(irab = 1, ALI, YAA)
                
                ' Aturan Nun Luruh: Jika 200 pas dan Idhafah, atau Idhafah dalam Sastra
                Dim luruhNun200 As Boolean
                luruhNun200 = (isIdhafah And sisaRatusan = 0) Or (isIdhafah And gaya = 3)
                
                If Not luruhNun200 Then strRatusan = strRatusan & NUN
                
            Case 3 To 9 ' 300 - 900
                Dim txtH As String: txtH = Kata(H, 1) ' Ratusan selalu Mudzakkar
                
                If gaya = 1 Then ' MODERN
                    strRatusan = txtH & " " & tRatusModern & MAR
                Else ' KLASIK / SASTRA
                    If H = 8 Then
                        ' Kasus Khusus 800: Thamanu-mi'ah (Yaa luruh)
                        strRatusan = Left(Kata(8, 1), Len(Kata(8, 1)) - 1) & tRatusklasik & MAR
                        ' Kondisi Sastra Non-Idhafah
                        If (Not isIdhafah And gaya = 3) Then strRatusan = txtH & " " & tRatusklasik & ALI & TAA
                    Else
                        strRatusan = txtH & tRatusklasik & MAR
                    End If
                End If
        End Select
        V = sisaRatusan
    End If

    ' --- 2. PROSES PULUHAN DAN SATUAN (1 - 99) ---
    If V > 0 Then
        Dim idxLawan As Integer: idxLawan = IIf(PilihGender = 1, 2, 1)

        If V <= 10 Then
            ' --- SATUAN MURNI (1-10) ---
            If V = 1 Or V = 2 Then
                ' Penanganan Wahid/Ithnan dengan potensi luruh Nun
                Dim keIdhafah As Boolean: keIdhafah = (isIdhafah And sisaRatusan = 0)
                strSatuan = GetWahidIthnan(V, PilihGender, irab, keIdhafah)
                
            ElseIf V = 8 Then
                ' Logika Ism Manqus untuk Angka 8 (Thamaniyah vs Thamanin)
                If PilihGender = 1 Then
                    strSatuan = Kata(8, 2) ' Thamaniyah
                Else
                    If isIdhafah Then
                        strSatuan = Kata(8, 1) ' Thamani
                    Else
                        ' Kondisi Tanwin (Thamanin)
                        strSatuan = IIf(irab = 2, Kata(8, 1) & ALI & ChrW(1611), _
                                    Left(Kata(8, 1), Len(Kata(8, 1)) - 1) & ChrW(1613))
                    End If
                End If
            Else
                strSatuan = Kata(V, idxLawan)
            End If
            
        ElseIf V < 20 Then
            ' --- BELASAN (11-19) ---
            Dim unitB As Integer: unitB = V Mod 10
            Dim txtU As String
            
            Select Case unitB
                Case 1: txtU = IIf(PilihGender = 1, Kata(11, 1), Kata(11, 2)) ' Ahad / Ihda
                Case 2: txtU = GetWahidIthnan(2, PilihGender, irab, True) ' Ithna / Ithnata (Nun luruh)
                Case Else: txtU = Kata(unitB, idxLawan)
            End Select
            
            strSatuan = txtU & " " & Kata(10, PilihGender) ' ... 'Asyara / 'Asyarata
            
        Else
            ' --- PULUHAN (20-99) ---
            Dim sT As Integer: sT = V Mod 10
            Dim pul As String: pul = Puluhan(Int(V / 10))
            
            ' I'rab Puluhan: Wawu-Nun (Marfu) -> Ya-Nun (Mansub/Majrur)
            If irab <> 1 Then pul = Replace(pul, WAW & NUN, YAA & NUN)

            If sT > 0 Then
                Dim txtSat As String
                ' Angka 21, 22, dst: Nun pada "Ithnan" TIDAK luruh karena ada pemisah Wawu
                txtSat = IIf(sT = 1 Or sT = 2, GetWahidIthnan(sT, PilihGender, irab, False), Kata(sT, idxLawan))
                strSatuan = txtSat & spWaw & pul
            Else
                strSatuan = pul
            End If
        End If
    End If

    ' --- 3. PENGGABUNGAN AKHIR ---
    If strRatusan <> "" And strSatuan <> "" Then
        ' Gaya Sastra: Satuan dulu baru Ratusan
        ' Gaya Modern/Klasik: Ratusan dulu baru Satuan
        s = IIf(gaya = 3, strSatuan & spWaw & strRatusan, strRatusan & spWaw & strSatuan)
    Else
        s = strRatusan & strSatuan
    End If

    Core_SatuanDasarAr = Trim(s)
End Function

' ==============================================================================
' FUNGSI: KONVERSI SATUAN BESAR (Ribuan, Jutaan, dst.)
' ==============================================================================
Private Function Core_SatuanBesarAr(ByVal vNum As Variant, ByVal idx As Integer, _
                                   ByVal gaya As Integer, ByVal irab As Integer, _
                                   Optional isIdhafah As Boolean = False) As String
    
    Dim Cur As String
    Dim TANWIN_FATHAH As String: TANWIN_FATHAH = ChrW(1611)
    Dim V As Long: V = CLng(vNum)
    
    If V <= 0 Then Exit Function

    ' Mengambil unit dasar dari class (Alfun, Milyun, dsb.)
    Dim baseUnit As String: baseUnit = SatuanBesar(idx)
    
    ' --- 1. LOGIKA GAYA SASTRA (Skala Berulang) ---
    ' Contoh: Sejuta dalam Sastra kuno bisa disebut "Alfu Alfin" (Seribu Seribu)
    Dim isSastraBesar As Boolean: isSastraBesar = (gaya = 3 And idx >= 2)
    If isSastraBesar Then
        Dim i As Integer
        baseUnit = SatuanBesar(1) ' Mulai dari "Alfun"
        For i = 2 To idx
            baseUnit = baseUnit & " " & SatuanBesar(1)
        Next i
    End If
    
    ' --- 2. PENANGANAN ANGKA 1 (Mufrad) ---
    If V = 1 Then
        ' Jika tidak Idhafah dan Mansub, tambahkan Alif Tanwin
        Cur = baseUnit & IIf(Not isIdhafah And irab = 2, ALI, "")

    ' --- 3. PENANGANAN ANGKA 2 (Tasniyah: Alfani / Alfaini) ---
    ElseIf V = 2 Then
        Dim suf As String: suf = IIf(irab = 1, ALI, YAA) ' Ali (Marfu), Ya (Mansub/Majrur)
        
        If isSastraBesar Then
            Dim spacePos As Integer: spacePos = InStr(baseUnit, " ")
            ' Memasukkan akhiran Tasniyah ke kata pertama dalam rangkaian skala
            Cur = Left(baseUnit, spacePos - 1) & suf & _
                  IIf(Not isIdhafah Or (isIdhafah And isSastraBesar), NUN, "") & _
                  Mid(baseUnit, spacePos)
        Else
            Cur = baseUnit & suf
            ' Nun luruh jika dalam posisi Idhafah (Mudhaf)
            If Not isIdhafah Then Cur = Cur & NUN
        End If

    ' --- 4. PENANGANAN ANGKA 3 - 999 ---
    Else
        Dim sisa As Integer: sisa = V Mod 100
        Dim Ratusan As Integer: Ratusan = (V \ 100) * 100
        
        ' --- A. Kasus Kombinasi (Contoh: 150.000) ---
        If (Ratusan > 0) And (sisa > 0) Then
            If gaya = 3 Then ' SASTRA (Satuan/Sisa dulu baru Ratusan)
                Cur = Core_SatuanDasarAr(V, 1, gaya, irab, True) & " " & baseUnit
            ElseIf gaya = 2 Then ' KLASIK (Ratusan - Wawu - Sisa - Unit)
                Cur = Core_SatuanDasarAr(Ratusan, 1, gaya, irab, True) & " " & baseUnit & _
                      " " & WAW & Core_SatuanBesarAr(sisa, idx, gaya, irab, isIdhafah)
            Else ' MODERN (Ratusan - Wawu - Sisa + Unit di setiap blok)
                Cur = Core_SatuanDasarAr(Ratusan, 1, gaya, irab, False) & " " & _
                      WAW & " " & Core_SatuanBesarAr(sisa, idx, gaya, irab, isIdhafah)
            End If
            
        ' --- B. Kasus Angka Tunggal dalam Blok (Contoh: 5.000, 20.000) ---
        Else
            Dim keUnitIdhafah As Boolean: keUnitIdhafah = (sisa = 0 Or sisa = 2)
            Dim resAngka As String: resAngka = Core_SatuanDasarAr(V, 1, gaya, irab, keUnitIdhafah)
            
            ' Aturan 3-10: Ma'dud (Unit) harus Jamak Majrur
            If V >= 3 And V <= 10 Then
                If isSastraBesar Then
                    Cur = resAngka & " " & GetMa_dud(V, SatuanBesar(1)) & Mid(baseUnit, InStr(baseUnit, " "))
                Else
                    Cur = resAngka & " " & GetMa_dud(V, baseUnit)
                End If
                
            ' Aturan 11-99: Tamyiz (Unit) harus Mufrad Manshub
            Else
                Dim unitFinal As String: unitFinal = baseUnit
                
                If sisa >= 11 And sisa <= 99 Then
                    If isSastraBesar Then
                        Dim sPos As Integer: sPos = InStr(unitFinal, " ")
                        unitFinal = Left(unitFinal, sPos - 1) & IIf(Not isIdhafah, TANWIN_FATHAH & ALI, "") & Mid(unitFinal, sPos)
                    Else
                        unitFinal = unitFinal & IIf(Not isIdhafah, TANWIN_FATHAH & ALI, "")
                    End If
                
                ' Aturan 100+: Mufrad Majrur
                ElseIf sisa = 0 And Ratusan > 0 Then
                    unitFinal = unitFinal & IIf(Not isIdhafah, ChrW(1613), "") ' Tanwin Kasrah
                End If
                
                Cur = resAngka & " " & unitFinal
            End If
        End If
    End If
    
    Core_SatuanBesarAr = Cur
End Function

' ==============================================================================
' FUNGSI: PENANGANAN KHUSUS ANGKA 1 (WAHID) DAN 2 (ITHNAN)
' ==============================================================================
Private Function GetWahidIthnan(ByVal val As Variant, _
                               ByVal gender As Integer, _
                               ByVal irab As Integer, _
                               Optional isIdhafah As Boolean = False) As String
    
    ' Pastikan input adalah angka yang valid untuk fungsi ini
    If val = 1 Then
        ' Mengambil Wahid (Mudzakkar) atau Wahidah (Muannas) dari Array Kata
        GetWahidIthnan = Kata(1, gender)
        
    ElseIf val = 2 Then
        Dim s As String
        ' Akar kata dasar: Alif - Tha - Nun (Ithn...)
        s = ALI & THA & NUN
        
        ' 1. Gender: Jika Muannas, tambahkan huruf Taa (menjadi Ithnat...)
        If gender = 2 Then s = s & TAA
        
        ' 2. I'rab: Tambah Alif (Marfu/Subjek) atau Ya (Mansub-Majrur/Objek-Keterangan)
        s = s & IIf(irab = 1, ALI, YAA)
        
        ' 3. Aturan Idhafah:
        ' Nun I'rab HARUS luruh jika angka 2 menjadi Mudhaf (misal: Ithna 'Asyara)
        ' atau jika diikuti oleh Ma'dud dalam struktur Idhafah.
        If Not isIdhafah Then
            s = s & NUN
        End If
        
        GetWahidIthnan = s
    End If
End Function

' ==============================================================================
' FUNGSI: PENANGANAN MA'DUD (KATA BENDA/UNIT) UNTUK ANGKA 3-10
' ==============================================================================
Private Function GetMa_dud(ByVal Angka As Long, ByVal Satuan As String) As String

    Dim V As Long: V = Angka Mod 100
    
    ' --- Definisi Satuan Dasar (Unicode) ---
    Dim sRiyal As String:  sRiyal = RHO & YAA & ALI & LAM
    Dim sDinar As String:  sDinar = dal & YAA & NUN & ALI & RHO
    Dim sDirham As String: sDirham = dal & RHO & haa & mim
    Dim sRupiah As String: sRupiah = RHO & WAW & baa & YAA & MAR
    Dim sAlfun As String:  sAlfun = HAM & LAM & FHA
    Dim sFils As String:   sFils = ChrW(1601) & LAM & sin

    ' --- LOGIKA PERUBAHAN JAMAK (3-10) ---
    ' Dalam tata bahasa Arab: 3-10 menggunakan Jamak Majrur
    If V >= 3 And V <= 10 Then
        Select Case Satuan
            Case sRiyal
                GetMa_dud = sRiyal & ALI & TAA            ' Riyalat
            Case sDinar
                GetMa_dud = dal & NUN & ALI & NUN & YAA & RHO ' Dananir
            Case sDirham
                GetMa_dud = dal & RHO & ALI & haa & mim   ' Darahim
            Case sRupiah
                GetMa_dud = RHO & WAW & baa & YAA & ALI & TAA ' Rubiyaat
            Case sAlfun
                GetMa_dud = ChrW(1570) & LAM & ALI & FHA  ' Aalaaf
            Case sFils
                GetMa_dud = ChrW(1601) & LAM & WAW & sin  ' Fuluus
            
            ' --- Logika Skala Besar (Milyun, Milyar, dst) ---
            Case Else
                ' 1. Pola Milyun (10^6) -> Malayiin (Jamak Taksir)
                If Satuan = mim & LAM & YAA & WAW & NUN Then
                    GetMa_dud = mim & LAM & ALI & YAA & YAA & NUN
                
                ' 2. Pola n-ilyun (Bilyun, Trilyun) -> n-ilyunat
                ElseIf InStr(Satuan, LAM & YAA & WAW & NUN) > 0 Then
                    GetMa_dud = Satuan & ALI & TAA
                
                ' 3. Pola n-ilyar (Milyar, Bilyar) -> n-ilyarat
                ElseIf InStr(Satuan, LAM & YAA & ALI & RHO) > 0 Then
                    GetMa_dud = Satuan & ALI & TAA
                
                Else
                    GetMa_dud = Satuan
                End If
        End Select
    Else
        ' Selain rentang 3-10, Ma'dud kembali ke bentuk Mufrad (Tunggal)
        ' Contoh: 11-99 (Mufrad Manshub), 100+ (Mufrad Majrur)
        GetMa_dud = Satuan
    End If
End Function

' ==============================================================================
' FUNGSI: MENGHASILKAN ANGKA URUTAN/ORDINAL (1 - 999)
' Contoh: Al-Awwal, Al-Hadi wal-'Isyrun, dsb.
' ==============================================================================
Public Function GetArabicOrdinal(ByVal n As Integer, ByVal g As Integer, _
                                 Optional ByVal gaya As Integer = 2, _
                                 Optional ByVal irab As Integer = 1) As String
    
    ' g: 1 = Mudzakkar, 2 = Muannats
    If n < 1 Or n > 999 Then
        GetArabicOrdinal = "Error: Range 1-999"
        Exit Function
    End If

    Dim h_mid As String: h_mid = ChrW(1574) ' Ya dengan Hamzah di atas
    Dim AL As String: AL = ALI & LAM
    Dim WA_AL_SPASI As String: WA_AL_SPASI = " " & WAW & AL

    ' --- 1. INISIALISASI POLA FA'IL (URUTAN 1-10) ---
    Dim Ord(1 To 10, 1 To 2) As String
    Dim i As Byte
    
    ' Pola Umum: Karakter pertama + Alif + Sisa Karakter
    For i = 3 To 10
        If (i <> 6 And i <> 8) Then
            Ord(i, 1) = Left(Kata(i, 1), 1) & ALI & Mid(Kata(i, 1), 2)
            Ord(i, 2) = Ord(i, 1) & MAR
        End If
    Next i

    ' Pengecualian & Kasus Khusus Ordinal
    Ord(1, 1) = HAM & WAW & LAM:              Ord(1, 2) = HAM & WAW & LAM & MAQ   ' Awwal / Ula
    Ord(2, 1) = THA & ALI & NUN & YAA:        Ord(2, 2) = Ord(2, 1) & MAR         ' Thani / Thaniyah
    Ord(3, 1) = THA & ALI & LAM & THA:        Ord(3, 2) = Ord(3, 1) & MAR         ' Thalith / Thalithah
    Ord(4, 1) = RHO & ALI & baa & AIN:        Ord(4, 2) = Ord(4, 1) & MAR         ' Rabi' / Rabi'ah
    Ord(6, 1) = sin & ALI & dal & sin:        Ord(6, 2) = Ord(6, 1) & MAR         ' Sadis / Sadisah
    Ord(8, 1) = THA & ALI & mim & NUN:        Ord(8, 2) = Ord(8, 1) & MAR         ' Thamin / Thaminah
    Ord(10, 1) = AIN & ALI & syi & RHO:       Ord(10, 2) = Ord(10, 1) & MAR        ' Asyir / Asyirah
    
    ' --- 2. LOGIKA RATUSAN ---
    Dim strRatusan As String
    Dim r As Integer: r = n \ 100
    Dim s_p As Integer: s_p = n Mod 100
    
    If r > 0 Then
        Select Case r
            Case 1: strRatusan = IIf(gaya = 1, mim & h_mid & MAR, mim & ALI & h_mid & MAR)
            Case 2
                strRatusan = IIf(gaya = 1, mim & h_mid & TAA, mim & ALI & h_mid & TAA)
                strRatusan = strRatusan & IIf(irab = 1, ALI, YAA)
                If s_p <> 0 Then strRatusan = strRatusan & NUN
            Case 3 To 9
                Dim txtH As String: txtH = Kata(r, 1)
                If gaya = 1 Then
                    strRatusan = txtH & " " & mim & ChrW(1616) & h_mid & ChrW(1614) & MAR
                Else
                    strRatusan = IIf(r <> 8, txtH & mim & ALI & h_mid & MAR, _
                                 THA & ChrW(1614) & mim & ChrW(1614) & ALI & NUN & mim & ALI & h_mid & MAR)
                End If
        End Select
        strRatusan = AL & strRatusan
    End If
    
    ' --- 3. LOGIKA SATUAN & PULUHAN ---
    Dim strSisa As String
    If s_p > 0 Then
        If s_p <= 10 Then
            strSisa = AL & Ord(s_p, g)
        ElseIf s_p = 11 Then
            ' Al-Hadi 'Asyara / Al-Hadiyata 'Asyarata
            strSisa = AL & haa & ALI & dal & YAA & IIf(g = 2, MAR, "") & " " & Kata(10, g)
        ElseIf s_p <= 19 Then
            ' Belasan: Al-X 'Asyara
            strSisa = AL & Ord(s_p - 10, g) & " " & Kata(10, g)
        Else
            ' 20 - 99
            Dim sat As Integer: sat = s_p Mod 10
            Dim pul As Integer: pul = s_p \ 10
            Dim strPulBase As String: strPulBase = Puluhan(pul)

            If irab <> 1 Then strPulBase = Replace(strPulBase, WAW & NUN, YAA & NUN)
             
            If sat = 0 Then
                strSisa = AL & strPulBase
            Else
                Dim strSat As String
                ' Kaidah 21, 31, dst: Menggunakan Al-Haadi (bukan Al-Wahid)
                If sat = 1 Then
                    strSat = AL & haa & ALI & dal & YAA & IIf(g = 2, MAR, "")
                Else
                    strSat = AL & Ord(sat, g)
                End If
                strSisa = strSat & WA_AL_SPASI & strPulBase
            End If
        End If
    End If

    ' --- 4. PENGGABUNGAN AKHIR ---
    If r > 0 And s_p > 0 Then
        ' Menyambung ratusan ke sisa (Misal: Al-Mi'ah wal-Khamisah)
        GetArabicOrdinal = strRatusan & WA_AL_SPASI & Mid(strSisa, Len(AL) + 1)
    ElseIf r > 0 Then
        GetArabicOrdinal = strRatusan
    Else
        GetArabicOrdinal = strSisa
    End If
End Function

' ================= Harakat ===================================================
Function MapKataHarakat(ByVal w As String, ByVal irab As IrabArab, Optional ByVal gaya As GayaArab = Modern, Optional ByVal Mode As String, Optional ByVal isIdhafah As Boolean = False) As String
    ' === 1. DEFINISI HARAKAT (Unicode) ===
    Dim f As String: f = ChrW(1614)    ' Fathah
    Dim k As String: k = ChrW(1616)    ' Kasrah
    Dim d As String: d = ChrW(1615)    ' Dhommah
    Dim s As String: s = ChrW(1618)    ' Sukun
    Dim sy As String: sy = ChrW(1617)  ' Tasydid
    Dim dt As String: dt = ChrW(1612)  ' Dhommatain (uun)
    Dim kt As String: kt = ChrW(1613)  ' Kasratain (in)
    Dim ft As String: ft = ChrW(1611)  ' Fathatain (an)
    Dim ha As String ' baris / harakat akhir
       
    Dim wClean As String: wClean = HapusHarakat(Trim(w))
    
    Dim isAL As Boolean: isAL = (Left(wClean, 2) = ALI & LAM)
    Dim wOmitAL As String: wOmitAL = IIf(isAL, Mid(wClean, 3), wClean)
    
    ' Cek pola ordinal satuan (Alif di huruf kedua) atau kata "Pertama"
    Dim isFailPattern As Boolean: isFailPattern = (InStr(wOmitAL, ALI) = 2)
    Dim isAwwalUla As Boolean: isAwwalUla = (Right(wOmitAL, 3) = HAM & WAW & LAM Or Right(wOmitAL, 4) = HAM & WAW & LAM & MAQ)

    Select Case Mode
        Case "eja", "digit": ha = dt
        Case Else
            ' Jika isIdhafah = TRUE (Artinya: Ada kata setelahnya/Mudhaf)
            ' -> Gunakan Harakat Tunggal (u, a, i)
            If (isIdhafah Or isAL Or isAwwalUla) Then
                Select Case irab
                    Case 1: ha = d ' Marfu (u)
                    Case 2: ha = f ' Mansub (a)
                    Case 3: ha = k ' Majrur (i)
                    Case Else: ha = d
                End Select
            ' Jika isIdhafah = FALSE (Artinya: Kata terakhir/Berdiri sendiri)
            ' -> Gunakan Tanwin (un, an, in)
            Else
                Select Case irab
                    Case 1: ha = dt ' Marfua (un)
                    Case 2: ha = ft ' Mansub (an)
                    Case 3: ha = kt ' Majrur (in)
                    Case Else: ha = dt
                End Select
            End If
    End Select

    Dim res As String: res = ""
    
        ' --- A. SATUAN (0-10), BELASAN (11-19) & KATA STATIS ---
    Select Case wClean
        Case Kata(0, 1): MapKataHarakat = sho & k & FHA & s & RHO & ha
        
        ' --- ANGKA 1 (WAHID / WAHIDAH  -  AHADA / IHDA) ---
        Case Kata(1, 1): MapKataHarakat = WAW & f & ALI & haa & k & dal & IIf(ha = ft And Not isIdhafah, ft & ALI, ha): Exit Function
        Case Kata(1, 2): MapKataHarakat = WAW & f & ALI & haa & k & dal & f & MAR & ha: Exit Function
        Case HAM & haa & dal: MapKataHarakat = HAM & f & haa & f & dal & f
        Case ChrW(1573) & haa & dal & MAQ: MapKataHarakat = ChrW(1573) & k & haa & s & dal & f & MAQ
        
        ' --- ANGKA 2 (ITSNAIN / ITSNA / ITSNATAIN / ITSNATA) ---
        Case Kata(2, 1), ALI & THA & NUN & ALI & NUN, ALI & THA & NUN & YAA & NUN, ALI & THA & NUN & ALI, ALI & THA & NUN & YAA
            Dim isNunExistsM As Boolean: isNunExistsM = (Right(wClean, 1) = NUN)
            res = ALI & THA & s & NUN & f & IIf(irab = 1, ALI, YAA & s) & IIf(isNunExistsM And Not isIdhafah, NUN & k, "")
            
        Case Kata(2, 2), ALI & THA & NUN & TAA & ALI & NUN, ALI & THA & NUN & TAA & YAA & NUN, ALI & THA & NUN & TAA & ALI, ALI & THA & NUN & TAA & YAA
            Dim isNunExistsF As Boolean: isNunExistsF = (Right(wClean, 1) = NUN)
            res = ALI & THA & s & NUN & f & TAA & f & IIf(irab = 1, ALI, YAA & s) & IIf(isNunExistsF And Not isIdhafah, NUN & k, "")

        ' --- ANGKA 3-10  ---
        Case Kata(3, 1): MapKataHarakat = THA & f & LAM & f & ALI & THA & IIf(ha = ft And Not isIdhafah, ft & ALI, ha)
        Case Kata(3, 2): MapKataHarakat = THA & f & LAM & f & ALI & THA & f & MAR & ha
        
        Case Kata(4, 1): MapKataHarakat = HAM & f & RHO & s & baa & f & AIN & IIf(ha = ft And Not isIdhafah, ft & ALI, ha)
        Case Kata(4, 2): MapKataHarakat = HAM & f & RHO & s & baa & f & AIN & f & MAR & ha
        
        Case Kata(5, 1): MapKataHarakat = kha & f & mim & s & sin & IIf(ha = ft And Not isIdhafah, ft & ALI, ha)
        Case Kata(5, 2): MapKataHarakat = kha & f & mim & s & sin & f & MAR & ha
        
        Case Kata(6, 1): MapKataHarakat = sin & k & TAA & sy & IIf(ha = ft And Not isIdhafah, ft & ALI, ha)
        Case Kata(6, 2): MapKataHarakat = sin & k & TAA & sy & f & MAR & ha
        
        Case Kata(7, 1): MapKataHarakat = sin & f & baa & s & AIN & IIf(ha = ft And Not isIdhafah, ft & ALI, ha)
        Case Kata(7, 2): MapKataHarakat = sin & f & baa & s & AIN & f & MAR & ha
        
        ' --- ANGKA 8 (THAMANIYAH) ---
        Case Kata(8, 1): MapKataHarakat = THA & f & mim & f & ALI & NUN & k & YAA & IIf(isIdhafah, ha, IIf(irab = 3, k, s))
        Case Kata(8, 2): MapKataHarakat = THA & f & mim & f & ALI & NUN & k & YAA & f & MAR & ha
        Case THA & mim & ALI & NUN: MapKataHarakat = THA & f & mim & f & ALI & NUN & IIf(isIdhafah, ha, kt)
        Case THA & mim & ALI & NUN & YAA & ALI: MapKataHarakat = THA & f & mim & f & ALI & NUN & k & YAA & IIf(isIdhafah, f & ALI, ft & ALI)

        Case Kata(9, 1): MapKataHarakat = TAA & k & sin & s & AIN & IIf(ha = ft And Not isIdhafah, ft & ALI, ha)
        Case Kata(9, 2): MapKataHarakat = TAA & k & sin & s & AIN & f & MAR & ha
        
        Case Kata(10, 1): MapKataHarakat = AIN & f & syi & f & RHO & IIf(ha = ft And Not isIdhafah, ft & ALI, ha)
        Case Kata(10, 2): MapKataHarakat = AIN & f & syi & s & RHO & f & MAR & ha
         
        Case WAW: MapKataHarakat = WAW & f
        Select Case ha
        Case ha = d: ha = dt
        Case ha = f: ha = ft
        Case ha = k: ha = kt
        End Select
        
        Case FHA & ALI & sho & LAM & MAR: MapKataHarakat = FHA & f & ALI & sho & k & LAM & f & MAR & ha: Exit Function
        Case Else: MapKataHarakat = w
    End Select
    
    If res <> "" Then MapKataHarakat = res: Exit Function
       
        ' --- B. PULUHAN (20-90) ---
    If Len(wClean) >= 4 Then
        Dim akhir As String: akhir = Right(w, 2)
        If akhir = WAW & NUN Or akhir = YAA & NUN Then
            Dim akarP As String: akarP = Left(w, Len(w) - 2)
            Dim hAkhir As String: hAkhir = IIf(irab = 1, d & WAW & s & NUN & f, k & YAA & s & NUN & f)
            Select Case akarP
                Case Kata(10, 1): MapKataHarakat = AIN & k & syi & s & RHO & hAkhir: Exit Function
                Case Kata(3, 1):  MapKataHarakat = THA & f & LAM & f & ALI & THA & hAkhir: Exit Function
                Case Kata(4, 1):  MapKataHarakat = HAM & f & RHO & s & baa & f & AIN & hAkhir: Exit Function
                Case Kata(5, 1):  MapKataHarakat = kha & f & mim & s & sin & hAkhir: Exit Function
                Case Kata(6, 1):  MapKataHarakat = sin & k & TAA & sy & hAkhir: Exit Function
                Case Kata(7, 1):  MapKataHarakat = sin & f & baa & s & AIN & hAkhir: Exit Function
                Case THA & mim & ALI & NUN: MapKataHarakat = THA & f & mim & f & ALI & NUN & hAkhir: Exit Function
                Case Kata(9, 1):  MapKataHarakat = TAA & k & sin & s & AIN & hAkhir: Exit Function
            End Select
        End If
    End If
    
    ' --- C. RATUSAN (100-900) ---
        Dim vD As String
        If irab = 1 Then vD = ALI Else vD = YAA & s
    
    ' 1. 200
    Dim is200 As Boolean
    is200 = InStr(wClean, mim & HZA & TAA & ALI) > 0 Or _
    InStr(wClean, mim & HZA & TAA & YAA) > 0 Or _
    InStr(wClean, mim & ALI & HZA & TAA & ALI) > 0 Or _
    InStr(wClean, mim & ALI & HZA & TAA & YAA) > 0
    
    If is200 Then
        Dim hasAL_200 As Boolean: hasAL_200 = (Left(wClean, 2) = ALI & LAM)
        Dim alz200 As String: alz200 = IIf(InStr(wClean, mim & ALI) > 0, ALI, "")
        Dim hNR As Boolean: hNR = (Right(wClean, 1) = NUN)
        
        ' vD: I'rab dual (Alif untuk Marfu, Ya untuk Manshub/Majrur)
        Dim res200 As String
        res200 = mim & k & alz200 & HZA & f & TAA & f & vD & IIf(hNR, NUN & k, "")
        
        MapKataHarakat = IIf(hasAL_200, ALI & LAM & s, "") & res200
        Exit Function
    End If

    ' 2. 100 dan 300-900
    ' Cari Mim + Hamzah saja agar fleksibel (bisa Ta Marbuthah atau Jamak Ta Maftuhah)
    Dim pMiah As Integer
    pMiah = InStr(wClean, mim & HZA)
    If pMiah = 0 Then pMiah = InStr(wClean, mim & ALI & HZA)
    
    ' 3. EKSEKUSI 100 DAN 300-900
    If pMiah > 0 Then
        ' Tentukan Alif Zaidah (alzM)
        Dim alzM As String
        If Mid$(wClean, pMiah + 1, 1) = ALI Then alzM = ALI Else alzM = ""
        
        ' Cek apakah input aslinya pakai TAA (Jamak) atau MAR (Tunggal)
        Dim isJamakInput As Boolean
        isJamakInput = (InStr(wClean, mim & alzM & HZA & ALI & TAA) > 0) Or _
                       (InStr(wClean, mim & alzM & HZA & TAA) > 0)
        
        ' Tentukan karakter akhiran yang akan digunakan (mengikuti input)
        Dim charAkhir As String
        charAkhir = IIf(isJamakInput, ALI & TAA, MAR)

        ' --- 100 (Atau Kata Ratusan Berdiri Sendiri) ---
        If pMiah = 1 Or (pMiah = 3 And Left$(wClean, 2) = ALI & LAM) Then
            Dim hStandar As String: hStandar = IIf(isJamakInput, kt, ha)
            Dim hasAL_100 As Boolean: hasAL_100 = (Left$(wClean, 2) = ALI & LAM)
            
            MapKataHarakat = IIf(hasAL_100, ALI & LAM & s, "") & mim & k & alzM & HZA & f & charAkhir & hStandar
            Exit Function
            
        ' --- 300-900 (Struktur Idhafah) ---
        Else
            Dim preR As String: preR = Trim$(Left$(wClean, pMiah - 1))
            If preR <> "" Then
                ' Rekursif untuk satuannya (misal: "Sab'u")
                Dim resSatuan As String: resSatuan = MapKataHarakat(preR, irab, gaya, Mode, True)
                
                ' Penentuan Harakat Akhir (Majrur karena Mudhaf Ilaih)
                Dim hMiah As String: hMiah = IIf(isIdhafah Or isAL Or isAwwalUla, k, kt)
                
                ' Rakit miahBase secara dinamis mengikuti input
                Dim miahBase As String
                miahBase = mim & k & alzM & HZA & f & charAkhir & hMiah

                If gaya = Modern Then
                    MapKataHarakat = resSatuan & " " & miahBase
                Else
                    ' Gaya Klasik: Sambung
                    If HapusHarakat(Trim$(resSatuan)) = THA & mim & ALI & NUN Then
                        ' Kasus 800 (Thamani)
                        MapKataHarakat = THA & f & mim & f & ALI & NUN & k & miahBase
                    Else
                        MapKataHarakat = resSatuan & miahBase
                    End If
                End If
                Exit Function
            End If
        End If
    End If

    ' --- D. RIBUAN ---
    If (Left(wClean, 3) = HAM & LAM & FHA Or Left(wClean, 3) = HAM & LAM & FHA & ALI Or InStr(wClean, ChrW(1570)) > 0) Then
       Dim resAlf As String: resAlf = ""
       
              ' Alfun / Alfan / Alfin (1.000) - TUNGGAL
            If Right(wClean, 1) = ALI And Not (irab = 1 And isIdhafah) Then
                resAlf = HAM & f & LAM & s & FHA & ft & ALI
                
              ' Alfani / Alfaini (2.000) - DUAL
            ElseIf InStr(wClean, FHA & ALI) > 0 Or InStr(wClean, FHA & YAA) > 0 Then
            Dim hasNun As Boolean: hasNun = (Right(wClean, 1) = NUN)
            
            ' I'rab: ALI (Marfu), YAA (Manshub/Majrur)
            'vD = IIf(irab = 1, ALI, YAA & s)
            
            resAlf = HAM & f & LAM & s & FHA & f & vD & IIf(hasNun, NUN & k, "")
            
             ' Aalaf (3.000 - 10.000)
            ElseIf InStr(wClean, ChrW(1570) & LAM & ALI & FHA) > 0 Then
            resAlf = ChrW(1570) & LAM & ALI & f & FHA & IIf(isIdhafah, k, kt)
            
            Else
            ' Harakat akhir mengikuti I'rab (Default: ha / dhommah)
            Dim hTunggal As String
                 If isIdhafah Then hTunggal = k Else hTunggal = IIf(ha = ft And Not isIdhafah, ft & ALI, ha)
                 resAlf = HAM & f & LAM & s & FHA & hTunggal
            End If
                    MapKataHarakat = resAlf
    Exit Function
    End If
       
    ' --- E. SKALA BESAR (Milyun s/d 1000 Digit) ---
        Dim wID As String: wID = HapusHarakat(w)
    
    ' 1. Pembersihan Sinyal Akhiran (Normalisasi agar deteksi InStr akurat)
    If Right(wID, 1) = f Then wID = Left(wID, Len(wID) - 1)
    If Right(wID, 1) = ALI Then wID = Left(wID, Len(wID) - 1)
    If Right(wID, 2) = ALI & ft Then wID = Left(wID, Len(wID) - 2)
    
    ' 2. Deteksi Jenis Skala (Boolean)
    Dim isIlyun As Boolean: isIlyun = (InStr(wID, LAM & YAA & WAW & NUN) > 0)
    Dim isIlyar As Boolean: isIlyar = (InStr(wID, LAM & YAA & ALI & RHO) > 0)
    Dim isMalayiin As Boolean: isMalayiin = (wID = mim & LAM & ALI & YAA & YAA & NUN)
    Dim isMilyarat As Boolean: isMilyarat = (InStr(wID, LAM & YAA & ALI & RHO & ALI & TAA) > 0)
    
    ' Deteksi Jamak Skala Besar (Lyun -> Lyunat) untuk Angka 3-10
    Dim isJamakLyun As Boolean: isJamakLyun = (InStr(wID, LAM & YAA & WAW & NUN & ALI & TAA) > 0) Or _
                                             (isIlyun And Right(HapusHarakat(w), 2) = ALI & TAA)
    
    ' Masuk ke logika pemrosesan utama
    If isIlyun Or isIlyar Or isMalayiin Or isMilyarat Or isJamakLyun Then
        
        ' --- A. KASUS KHUSUS: MALAYIIN (Jutaan - Jamak Taksir) ---
        If isMalayiin Then
            ' mim(f) - LAM(f) - ALI - YAA(k) - YAA(s) - NUN(h)
            MapKataHarakat = mim & f & LAM & f & ALI & YAA & k & YAA & s & NUN & f
            Exit Function
        End If
            ' Mencari posisi huruf LAM sebagai batas antara prefiks (Un/Duo) dan inti (Lyun/Lyar)
        Dim bL As Integer: bL = InStr(wID, LAM)
        Dim pRaw As String: pRaw = IIf(bL > 1, Left(wID, bL - 1), "")
        Dim finalPrefix As String: finalPrefix = ""
    ' Loop membedah pRaw dari kiri ke kanan berdasarkan kecocokan string terpanjang
        Do While Len(pRaw) > 0
            ' --- Deteksi 6 Huruf ---
            If Left(pRaw, 6) = kaf & WAW & ALI & TAA & WAW & RHO Then ' Kwator-
                ' Ku-waa-tuu-ri (Menambahkan sukun 's' pada WAW kedua agar terbaca Mad)
                finalPrefix = finalPrefix & kaf & d & WAW & f & ALI & TAA & d & WAW & s & RHO & k: pRaw = Mid(pRaw, 7)
            
            ' --- Deteksi 5 Huruf ---
            ElseIf Left(pRaw, 5) = kaf & WAW & ALI & dal & RHO Then ' Kwadr-
                ' Ku-waa-d-ri (Koreksi harakat dal menjadi sukun 's' dan Rho menjadi kasrah 'k')
                finalPrefix = finalPrefix & kaf & d & WAW & f & ALI & dal & s & RHO & k: pRaw = Mid(pRaw, 6)
            
            ' --- Deteksi 4 Huruf ---
            ElseIf Left(pRaw, 4) = kaf & WAW & YAA & NUN Then ' Kwin
                finalPrefix = finalPrefix & kaf & d & WAW & k & YAA & s & NUN & s: pRaw = Mid(pRaw, 5)
            ElseIf Left(pRaw, 4) = sin & baa & TAA & NUN Then ' Septen-
                finalPrefix = finalPrefix & sin & k & baa & s & TAA & f & NUN & s: pRaw = Mid(pRaw, 5)
            ElseIf Left(pRaw, 4) = sin & baa & TAA & ALI Then ' Sibtaa-
                finalPrefix = finalPrefix & sin & k & baa & s & TAA & f & ALI: pRaw = Mid(pRaw, 5)
            ElseIf Left(pRaw, 4) = HAM & kaf & TAA & WAW Then ' Okto-
                finalPrefix = finalPrefix & HAM & d & kaf & s & TAA & d & WAW & s: pRaw = Mid(pRaw, 5)
            ElseIf Left(pRaw, 4) = NUN & WAW & FHA & mim Then ' Novem-
                finalPrefix = finalPrefix & NUN & d & WAW & s & FHA & f & mim & s: pRaw = Mid(pRaw, 5)
            
            ' --- Deteksi 3 Huruf ---
            ElseIf Left(pRaw, 3) = TAA & RHO & YAA Then ' Tre- / Tri-
                finalPrefix = finalPrefix & TAA & k & RHO & k & YAA & s: pRaw = Mid(pRaw, 4)
            ElseIf Left(pRaw, 3) = sin & YAA & kaf Then ' Seks- (1)
                finalPrefix = finalPrefix & sin & k & YAA & s & kaf & s: pRaw = Mid(pRaw, 4)
            ElseIf Left(pRaw, 3) = sin & kaf & sin Then ' Seks- (2) ' Sik-si
                finalPrefix = finalPrefix & sin & k & kaf & s & sin & k: pRaw = Mid(pRaw, 4)
            ElseIf Left(pRaw, 3) = dal & YAA & sin Then ' Desi-
                finalPrefix = finalPrefix & dal & k & YAA & s & sin & k: pRaw = Mid(pRaw, 4)
            ElseIf Left(pRaw, 3) = sin & NUN & TAA Then ' Senti-
                finalPrefix = finalPrefix & sin & f & NUN & s & TAA & k: pRaw = Mid(pRaw, 4)
            ElseIf Left(pRaw, 3) = jim & NUN & TAA Then ' Jinta-
                finalPrefix = finalPrefix & jim & k & NUN & s & TAA & f: pRaw = Mid(pRaw, 4)
            ElseIf Left(pRaw, 3) = sin & baa & TAA Then ' sibta-
                finalPrefix = finalPrefix & sin & k & baa & s & TAA & f: pRaw = Mid(pRaw, 4)
            ElseIf Left(pRaw, 3) = HAM & kaf & TAA Then ' okta-
                finalPrefix = finalPrefix & HAM & d & kaf & s & TAA & f: pRaw = Mid(pRaw, 4)
            ElseIf Left(pRaw, 3) = NUN & WAW & NUN Then ' noni-
                finalPrefix = finalPrefix & NUN & d & WAW & s & NUN & k: pRaw = Mid(pRaw, 4)
            
            ' --- Deteksi 2 Huruf ---
            ElseIf Left(pRaw, 2) = HAM & NUN Then ' Un- ' Un (U dammah)
                finalPrefix = finalPrefix & HAM & d & NUN & s: pRaw = Mid(pRaw, 3)
            ElseIf Left(pRaw, 2) = dal & WAW Then ' Duo-
                finalPrefix = finalPrefix & dal & d & WAW & f: pRaw = Mid(pRaw, 3)
                
            ' --- Deteksi 1 Huruf (Harakat Default / Fallback) ---
            ElseIf Left(pRaw, 1) = mim Then ' Mi-
                finalPrefix = finalPrefix & mim & k: pRaw = Mid(pRaw, 2)
            ElseIf Left(pRaw, 1) = baa Then ' Bi-
                finalPrefix = finalPrefix & baa & k: pRaw = Mid(pRaw, 2)
            ElseIf Left(pRaw, 1) = YAA Then ' i (Vowel Linker)
                finalPrefix = finalPrefix & YAA & s: pRaw = Mid(pRaw, 2)
            Else
                finalPrefix = finalPrefix & Left(pRaw, 1) & s: pRaw = Mid(pRaw, 2)
            End If
        Loop
    
        ' --- C. PENENTUAN CORE (BODY) & HARAKAT I'RAB ---
        Dim finalCore As String, finalLast As String
    
        If isIlyun Then
            ' Pola: (LAM-kasroh, YAA-sukun/fathah, WAW-sukun, NUN)
            finalCore = LAM & k & YAA & d & WAW & s
            finalLast = NUN
        Else
            ' Pola:  (LAM-kasroh, YAA-fathah, ALI, RHO)
            finalCore = LAM & k & YAA & f & ALI
            finalLast = RHO
        End If
    
        ' --- D. PENYUSUNAN AKHIR (Suffix & Gramatika Arab) ---
        ' Jika bentuk Jamak Muannas Salim (Milyarat / Lyunat)
        If isMilyarat Then
            Dim hAtMilyar As String: hAtMilyar = IIf(irab = 3 Or isIdhafah, k, kt)
            MapKataHarakat = finalPrefix & LAM & s & YAA & f & ALI & RHO & f & ALI & TAA & hAtMilyar
            Exit Function
        ElseIf isJamakLyun Then
            Dim hAtLyun As String: hAtLyun = IIf(irab = 3 Or isIdhafah, kt, dt)
            MapKataHarakat = finalPrefix & finalCore & finalLast & f & ALI & TAA & hAtLyun
            Exit Function
        End If
    
        If Right(w, 3) = YAA & WAW & NUN Then
            MapKataHarakat = finalPrefix & finalCore & finalLast & f
        End If

        ' Jika bentuk Dual (Tatsniyah - Dua Juta/Miliar)
        Dim cleanW As String: cleanW = HapusHarakat(w)
        If Right(cleanW, 2) = ALI & NUN Or Right(cleanW, 2) = YAA & NUN Then
            Dim dSuff As String
            If Not isIdhafah Then
                If irab = 1 Then dSuff = ALI & NUN & k Else dSuff = YAA & s & NUN & k
            Else
                ' Jika Mudhaf (Idhafah), huruf NUN dibuang
                If irab = 1 Then dSuff = ALI Else dSuff = YAA & s
            End If
            MapKataHarakat = finalPrefix & finalCore & finalLast & f & dSuff
            Exit Function
        End If
    
        If Right(w, 1) = ALI Then
            MapKataHarakat = finalPrefix & finalCore & finalLast & IIf(isIdhafah, f & ALI, ft & ALI)
            Exit Function
        End If
    
        ' Bentuk Standar (Mufrad)
        MapKataHarakat = finalPrefix & finalCore & finalLast & ha
        Exit Function
    End If

    ' --- F. ORDINAL ---
    ' ==========================================================
    If (isAL Or isFailPattern Or isAwwalUla) Then
                
        Dim resOrd As String: resOrd = ""
        
        ' 1. Proses Satuan Pola Fa-il
        If isAwwalUla Then
            If Right(wOmitAL, 4) = HAM & WAW & LAM & MAQ Then MapKataHarakat = ALI & LAM & HAM & d & WAW & s & LAM & f & MAQ: Exit Function
            If Right(wOmitAL, 3) = HAM & WAW & LAM Then resOrd = HAM & f & WAW & sy & f & LAM
        ElseIf isFailPattern Then
            Select Case True

            ' Khusus Ke-11 (Hadi)
            Case InStr(wOmitAL, haa & ALI & dal & YAA) > 0 Or InStr(wOmitAL, haa & ALI & dal) > 0:
                resOrd = haa & f & ALI & dal & k & YAA
        
                Case InStr(wOmitAL, THA & ALI & NUN) > 0: resOrd = THA & f & ALI & NUN & k & YAA
                Case InStr(wOmitAL, THA & ALI & LAM) > 0: resOrd = THA & f & ALI & LAM & k & THA
                Case InStr(wOmitAL, RHO & ALI & baa) > 0: resOrd = RHO & f & ALI & baa & k & AIN
                Case InStr(wOmitAL, kha & ALI & mim) > 0: resOrd = kha & f & ALI & mim & k & sin
                Case InStr(wOmitAL, sin & ALI & dal) > 0: resOrd = sin & f & ALI & dal & k & sin
                Case InStr(wOmitAL, sin & ALI & baa) > 0: resOrd = sin & f & ALI & baa & k & AIN
                Case InStr(wOmitAL, THA & ALI & mim) > 0: resOrd = THA & f & ALI & mim & k & NUN
                Case InStr(wOmitAL, TAA & ALI & sin) > 0: resOrd = TAA & f & ALI & sin & k & AIN
                Case InStr(wOmitAL, AIN & ALI & syi) > 0: resOrd = AIN & f & ALI & syi & k & RHO
                Case Else: resOrd = wOmitAL
            End Select
        Else
        
        ' 2. Proses Ordinal Besar (20, 100)
           If Len(wOmitAL) >= 4 Then
         akhir = Right(wOmitAL, 2)
        If akhir = WAW & NUN Or akhir = YAA & NUN Then
            akarP = Left(wOmitAL, Len(wOmitAL) - 2)
            hAkhir = IIf(irab = 1, d & WAW & s & NUN & f, k & YAA & s & NUN & f)
            Select Case akarP
                Case Kata(10, 1): MapKataHarakat = AIN & k & syi & s & RHO & hAkhir: Exit Function
                Case Kata(3, 1):  MapKataHarakat = THA & f & LAM & f & ALI & THA & hAkhir: Exit Function
                Case Kata(4, 1): MapKataHarakat = HAM & f & RHO & s & baa & f & AIN & hAkhir: Exit Function
                Case Kata(5, 1):  MapKataHarakat = kha & f & mim & s & sin & hAkhir: Exit Function
                Case Kata(6, 1):  MapKataHarakat = sin & k & TAA & sy & hAkhir: Exit Function
                Case Kata(7, 1):  MapKataHarakat = sin & f & baa & s & AIN & hAkhir: Exit Function
                Case THA & mim & ALI & NUN: MapKataHarakat = THA & f & mim & f & ALI & NUN & hAkhir: Exit Function
                Case Kata(9, 1):  MapKataHarakat = TAA & k & sin & s & AIN & hAkhir: Exit Function
            End Select
        End If
        End If

            resOrd = MapKataHarakat(wOmitAL, irab, gaya, , True)
            ' Buang harakat i'rab dari kardinal agar tidak double di finalisasi
            If Right(resOrd, 1) = ha Or Right(resOrd, 1) = dt Or Right(resOrd, 1) = f Then
                resOrd = Left(resOrd, Len(resOrd) - 1)
            End If
        End If

    If InStr(wOmitAL, " ") > 0 Then
        ' Cek apakah ini Ordinal Belasan (Contoh: Al-Rabi'a 'Asyara)
        ' atau Ordinal Gabungan (Contoh: Al-Thalith wal-'Isyrun)
        
        Dim partsOrd() As String: partsOrd = Split(wOmitAL, " ")
        Dim resFull As String: resFull = ""
        Dim i As Integer
        
        ' Kita proses tiap kata secara terpisah
        For i = LBound(partsOrd) To UBound(partsOrd)
            Dim word As String: word = partsOrd(i)
            Dim resWord As String: resWord = ""
            
            ' Jika ada huruf WAW (dan) di awal kata (untuk angka gabungan 21-99)
            Dim hasWAW As Boolean: hasWAW = (Left(word, 1) = WAW And Len(word) > 1)
            Dim wordNoWaw As String: wordNoWaw = IIf(hasWAW, Mid(word, 2), word)
            
            ' Deteksi AL di dalam potongan kata (misal: wal-Isyrun)
            Dim innerAL As Boolean: innerAL = (Left(wordNoWaw, 2) = ALI & LAM)
            Dim wordFinal As String: wordFinal = IIf(innerAL, Mid(wordNoWaw, 3), wordNoWaw)
            
            ' Tentukan i'rab untuk potongan kata
            ' Belasan (11-19) selalu Fathah (2). Selain itu ikut i'rab asli.
            Dim TargetIrab As Integer: TargetIrab = irab
            If UBound(partsOrd) = 1 Then ' Potensi Belasan
                If partsOrd(1) = AIN & syi & RHO Or partsOrd(1) = AIN & syi & RHO & MAR Then TargetIrab = 2
            End If
            
            ' Panggil rekursi untuk mengharkati potongan kata
            resWord = MapKataHarakat(wordFinal, TargetIrab)
            
            ' Kembalikan AL dan WAW
            If innerAL Or (isAL And i = 0) Then
                Dim hPot As String: hPot = Left(resWord, 1)
                Dim strSyamPot As String: strSyamPot = TAA & THA & dal & DZA & RHO & ZAI & sin & syi & sho & DHA & THO & ZHO & LAM & NUN
                If InStr(strSyamPot, hPot) > 0 Then
                    resWord = ALI & LAM & hPot & sy & Mid(resWord, 2)
                Else
                    resWord = ALI & LAM & s & resWord
                End If
            End If
            
            ' Tambahkan WAW jika ada
            If hasWAW Then resWord = WAW & f & resWord
            
            ' Gabungkan
            resFull = resFull & IIf(resFull = "", "", " ") & resWord
        Next i
        
        ' Tambahkan harakat i'rab di akhir kalimat jika bukan belasan
        If TargetIrab <> 2 Then
             MapKataHarakat = resFull & ha
        Else
             MapKataHarakat = resFull
        End If
        Exit Function
    End If

        ' --- G. FINALISASI ---
        If resOrd <> "" Then
            If Right(wOmitAL, 1) = MAR And Right(resOrd, 1) <> MAR Then resOrd = resOrd & f & MAR
            If isAL Then
                Dim h1 As String: h1 = Left(resOrd, 1)
                Dim strSyam As String: strSyam = TAA & THA & dal & DZA & RHO & ZAI & sin & syi & sho & DHA & THO & ZHO & LAM & NUN
                If InStr(strSyam, h1) > 0 Then
                    MapKataHarakat = ALI & LAM & h1 & sy & Mid(resOrd, 2) & ha
                Else
                    MapKataHarakat = ALI & LAM & s & resOrd & ha
                End If
            Else
                MapKataHarakat = resOrd & ha
            End If
            Exit Function
        End If
    End If

End Function

Public Function BeriHarakat(ByVal Teks As String, ByVal irab As Integer, _
    Optional ByVal gaya As GayaArab = Modern, _
    Optional ByVal Mode As String = "", _
    Optional ByVal isIdhafah As Boolean = False, _
    Optional ByVal Level As LevelHarakat = 2, Optional ByVal isModeEja As Boolean) As String
    
    ' --- Inisialisasi Karakter Harakat & Huruf Arab ---
    Dim f As String: f = ChrW(1614)    ' Fathah
    Dim d As String: d = ChrW(1615)    ' Dammah
    Dim k As String: k = ChrW(1616)    ' Kasrah
    Dim s As String: s = ChrW(1618)    ' Sukun
    Dim TSY As String: TSY = ChrW(1617) ' Tasydid (Dipindahkan ke atas untuk akses global fungsi)

    Dim AYN As String: AYN = ChrW(1593) ' 'Ain
    Dim SYN As String: SYN = ChrW(1588) ' Syin

    ' --- Konstanta Kata Arab via ChrW (Anti-Corrupt di VBA) ---
    Dim SIFR As String: SIFR = ChrW(1589) & ChrW(1601) & ChrW(1585)
    Dim THAMANI As String: THAMANI = ChrW(1579) & ChrW(1605) & ChrW(1575) & ChrW(1606) & ChrW(1610)
    
    'Dim isModeEja As Boolean: isModeEja = (LCase(Mode) = "eja")

    ' --- Pembersihan Input ---
    Dim sInput As String: sInput = Trim(Teks)
    Do While InStr(sInput, "  ") > 0: sInput = Replace(sInput, "  ", " "): Loop
    If sInput = "" Then BeriHarakat = "": Exit Function

    Dim kataArr() As String: kataArr = Split(sInput, " ")
    Dim i As Integer, j As Integer
    Dim HasilFinal As String: HasilFinal = ""

    ' =========================================================
    ' STEP 1: PRE-PROCESSING (DILEWATI JIKA MODE EJA)
    ' =========================================================
    If UBound(kataArr) > 0 And Not isModeEja Then
        For i = 0 To UBound(kataArr) - 1
            If kataArr(i) = "[SKIP]" Then GoTo NextIteration
            
            Dim k1Clean As String: k1Clean = HapusHarakat(kataArr(i))
            Dim k2Clean As String: k2Clean = HapusHarakat(kataArr(i + 1))
            
            ' A. Kasus Belasan (11-19)
            If k2Clean = Kata(10, 1) Or k2Clean = Kata(10, 2) Then
                Dim prefixW As String: prefixW = IIf(Left(k1Clean, 1) = WAW And Len(k1Clean) > 3, WAW & f, "")
                Dim k1Murni As String: k1Murni = IIf(prefixW <> "", Mid(k1Clean, 2), k1Clean)
                Dim is12 As Boolean: is12 = (InStr(k1Murni, ALI & ChrW(1579) & NUN) > 0)
                
                If is12 Then
                    kataArr(i) = prefixW & MapKataHarakat(k1Murni, irab, gaya, Mode, isIdhafah)
                Else
                    kataArr(i) = prefixW & MapKataHarakat(k1Murni, 2, gaya, Mode, True)
                End If
                
                kataArr(i + 1) = MapKataHarakat(k2Clean, 2, gaya, Mode, True)
                kataArr(i) = kataArr(i) & " " & kataArr(i + 1)
                kataArr(i + 1) = "[SKIP]"
                
            ' B. Kasus Ratusan Modern (300-900 Terpisah)
            ElseIf (InStr(k2Clean, mim & HZA & MAR) > 0 Or InStr(k2Clean, mim & ALI & HZA & MAR) > 0) Then
                Dim isSatuan39 As Boolean: isSatuan39 = False
                Select Case k1Clean
                    Case Kata(3, 1), Kata(3, 2), Kata(4, 1), Kata(4, 2), Kata(5, 1), Kata(5, 2), _
                         Kata(6, 1), Kata(6, 2), Kata(7, 1), Kata(7, 2), Kata(8, 1), Kata(8, 2), _
                         Kata(9, 1), Kata(9, 2): isSatuan39 = True
                End Select
                
                If isSatuan39 Then
                    Dim adaKataLagi As Boolean: adaKataLagi = False
                
                    For j = i + 2 To UBound(kataArr)
                        Dim rawCek As String: rawCek = Trim(kataArr(j))
                        If rawCek <> "[SKIP]" And rawCek <> "" Then
                            Dim cleanCek As String: cleanCek = HapusHarakat(rawCek)
                            If (cleanCek = WAW Or cleanCek = mim & NUN Or cleanCek = mim & NUN) Then adaKataLagi = False Else adaKataLagi = True
                            Exit For
                        End If
                    Next j
                    
                    Dim statusIdhafahBlok As Boolean: statusIdhafahBlok = adaKataLagi
                    kataArr(i) = MapKataHarakat(k1Clean & " " & k2Clean, irab, gaya, Mode, statusIdhafahBlok)
                    kataArr(i + 1) = "[SKIP]"
                End If
            End If
NextIteration:
        Next i
    End If

    ' =========================================================
    ' STEP 2: PROSES FINAL (APPLY HARAKAT & HANDOVER KE MAP)
    ' =========================================================
    Dim lastIdx As Integer: lastIdx = -1
    For i = UBound(kataArr) To 0 Step -1
        If kataArr(i) <> "[SKIP]" And Trim(kataArr(i)) <> "" Then
            lastIdx = i: Exit For
        End If
    Next i


    For i = 0 To UBound(kataArr)
        If kataArr(i) <> "[SKIP]" And Trim(kataArr(i)) <> "" Then
            Dim currentWord As String: currentWord = kataArr(i)
            Dim wordClean As String: wordClean = HapusHarakat(currentWord)
            

            Dim irabLokal As Integer: irabLokal = irab
            Dim paksaIdhafah As Boolean: paksaIdhafah = False
            Dim sendIdhafah As Boolean
            
            ' --- B. LOGIKA KHUSUS MODE EJA ---
            If isModeEja Then
                irabLokal = 1
                sendIdhafah = False
                
                If wordClean = SIFR Then
                    kataArr(i) = ChrW(1589) & k & ChrW(1601) & s & ChrW(1585) & ChrW(1612)
                ElseIf wordClean = THAMANI Then
                    kataArr(i) = ChrW(1579) & f & ChrW(1605) & ChrW(1575) & ChrW(1606) & ChrW(1613)
                ElseIf InStr(currentWord, f) > 0 Or InStr(currentWord, d) > 0 Or InStr(currentWord, k) > 0 Then
                    kataArr(i) = currentWord
                Else
                    kataArr(i) = MapKataHarakat(currentWord, irab, gaya, Mode, sendIdhafah)
                End If
                
                HasilFinal = HasilFinal & Trim(kataArr(i)) & " "
                GoTo NextWord
            End If
            

            ' 2. DETEKSI I'RAB LOKAL
            If i > 0 Then
                Dim prevIdx As Integer: prevIdx = -1
                For j = i - 1 To 0 Step -1
                    If kataArr(j) <> "[SKIP]" Then prevIdx = j: Exit For
                Next j
                
                If prevIdx <> -1 Then
                    Dim prevC As String: prevC = HapusHarakat(kataArr(prevIdx))
                    If Left(wordClean, 1) = WAW Then
                        irabLokal = irab
                    ElseIf (InStr(prevC, mim & HZA & MAR) > 0 Or InStr(prevC, mim & ALI & HZA & MAR) > 0 Or _
                            InStr(prevC, mim & HZA & TAA) > 0 Or InStr(prevC, mim & ALI & HZA & TAA) > 0) Then
                        irabLokal = 3
                    Else
                        Dim isPuluhan As Boolean: isPuluhan = False
                        If Right(prevC, 2) = WAW & NUN Or Right(prevC, 2) = YAA & NUN Then
                            isPuluhan = True
                            If InStr(prevC, LAM & YAA & WAW & NUN) > 0 Then isPuluhan = False
                            If InStr(prevC, LAM & YAA & YAA & NUN) > 0 Then isPuluhan = False
                            If InStr(prevC, HAM & LAM & FHA) > 0 Then isPuluhan = False
                        End If
                        Dim isBelasan As Boolean
                        isBelasan = (prevC = AYN & SYN & RHO) Or (prevC = AYN & SYN & RHO & MAR)
                        If isPuluhan Or isBelasan Then irabLokal = 2
                    End If
                End If
            End If

            ' 3. EKSEKUSI FINAL & DETERMINASI IDHAFAH
            If i = lastIdx Then
                sendIdhafah = isIdhafah ' Ikuti parameter fungsi untuk kata terakhir
            Else
                Dim nextIdx As Integer: nextIdx = -1
                For j = i + 1 To UBound(kataArr)
                    If kataArr(j) <> "[SKIP]" Then nextIdx = j: Exit For
                Next j
                
                If nextIdx <> -1 Then
                    Dim nextC As String: nextC = HapusHarakat(kataArr(nextIdx))
                    Dim isFashilah As Boolean: isFashilah = (nextC = FHA & ALI & sho & LAM & MAR)
                    ' 1. Jika ada "Waw" (Dan) -> Bukan Idhafah (Ada Tanwin/Nun)
                    ' 2. Jika kata berikutnya "Fashilah" (Koma) -> Bukan Idhafah (Ada Tanwin/Nun)
                    
                    If (Left(nextC, 1) = WAW Or isFashilah) Then
                        sendIdhafah = False
                    Else
                        ' Otomatis True untuk menghubungkan [Angka] ke [Skala/Benda]
                        sendIdhafah = True
                    End If
                Else
                    sendIdhafah = isIdhafah
                End If
            End If

            If InStr(currentWord, f) > 0 Or InStr(currentWord, d) > 0 Or InStr(currentWord, k) > 0 Then
                HasilFinal = HasilFinal & Trim(currentWord) & " "
            Else
                Dim isWahid As Boolean: isWahid = (wordClean = Kata(1, 1) Or wordClean = Kata(1, 2))
                
                ' =================================================================
                ' Deteksi Waw + Al dan Syamsiah/Qamariah
                ' =================================================================
                If Left(wordClean, 1) = WAW And Len(wordClean) > 2 And Not isWahid Then
                    Dim sisaW As String
                    Dim isAdaAl As Boolean: isAdaAl = False
                    
                    ' Cek jika formatnya  (Waw + Alif + Lam)
                    If Mid(wordClean, 2, 2) = ALI & LAM Then
                        isAdaAl = True
                        sisaW = Mid(wordClean, 4) ' Ambil akar kata murni (contoh: ?????)
                    Else
                        sisaW = Mid(wordClean, 2) ' Ambil kata tanpa Waw
                    End If

                    ' Eksekusi pemetaan harakat pada akar kata MURNI
                    Dim hasilHarakat As String
                    hasilHarakat = MapKataHarakat(sisaW, irabLokal, gaya, Mode, sendIdhafah)
                    
                    If isAdaAl Then
                        If InStr(hasilHarakat, ChrW(1612)) > 0 Then      ' Dammatain
                            hasilHarakat = Replace(hasilHarakat, ChrW(1612), ChrW(1615)) ' Jadi Dammah
                        ElseIf InStr(hasilHarakat, ChrW(1611)) > 0 Then  ' Fathatain
                            hasilHarakat = Replace(hasilHarakat, ChrW(1611), ChrW(1614)) ' Jadi Fathah
                        ElseIf InStr(hasilHarakat, ChrW(1613)) > 0 Then  ' Kasratain
                            hasilHarakat = Replace(hasilHarakat, ChrW(1613), ChrW(1616)) ' Jadi Kasrah
                        End If
                        ' =========================================================

                        Dim hurufAwal As String: hurufAwal = Left(sisaW, 1)
                        Dim isSyamsiah As Boolean: isSyamsiah = False
                        
                        ' Deteksi 14 Huruf Syamsiah
                        Select Case hurufAwal
                            Case ChrW(1578), ChrW(1579), ChrW(1583), ChrW(1584), ChrW(1585), ChrW(1586), _
                                 ChrW(1587), ChrW(1588), ChrW(1589), ChrW(1590), ChrW(1591), ChrW(1592), _
                                 ChrW(1604), ChrW(1606)
                                isSyamsiah = True
                        End Select
                        
                        If isSyamsiah Then
                            ' SYAMSIAH: Waw Fathah + Alif + Lam (tanpa sukun) + Huruf Awal + TASYDID + Sisa Kata
                            Dim char1 As String: char1 = Left(hasilHarakat, 1)
                            Dim charSisa As String: charSisa = Mid(hasilHarakat, 2)
                            
                            If Left(charSisa, 1) = TSY Then
                                kataArr(i) = (WAW & f) & ALI & LAM & char1 & charSisa
                            Else
                                kataArr(i) = (WAW & f) & ALI & LAM & char1 & TSY & charSisa
                            End If
                        Else
                            ' QAMARIAH: Waw Fathah + Alif + Lam SUKUN + Hasil Kata
                            kataArr(i) = (WAW & f) & ALI & LAM & s & hasilHarakat
                        End If
                    Else
                        ' Normal Waw tanpa Al (Waw Fathah + Hasil)
                        kataArr(i) = (WAW & f) & hasilHarakat
                    End If
                ' =================================================================
                Else
                    kataArr(i) = MapKataHarakat(currentWord, irabLokal, gaya, Mode, sendIdhafah)
                End If
                
                HasilFinal = HasilFinal & Trim(kataArr(i)) & " "
            End If
        End If
NextWord:
    Next i

    ' =========================================================
    ' STEP 3: LEVELING HARAKAT
    ' =========================================================
    Dim teksHasil As String: teksHasil = Trim(HasilFinal)
    If Level >= 2 Then: BeriHarakat = teksHasil: Exit Function
    
    Dim hasilLeveling As String: hasilLeveling = ""
    Dim char As String
    Dim m As Integer
    For m = 1 To Len(teksHasil)
        char = Mid(teksHasil, m, 1)
        If AscW(char) >= 1611 And AscW(char) <= 1618 Then
            If Level = 1 Then
                If char = TSY Then
                    hasilLeveling = hasilLeveling & char
                ElseIf m < Len(teksHasil) Then
                    If Mid(teksHasil, m + 1, 1) = " " Then hasilLeveling = hasilLeveling & char
                Else
                    hasilLeveling = hasilLeveling & char
                End If
            End If
        Else
            hasilLeveling = hasilLeveling & char
        End If
    Next m
    BeriHarakat = Trim(hasilLeveling)
End Function

Function HapusHarakat(ByVal Teks As String) As String
    Dim i As Long, c As Long
    Dim Hasil As String:    Hasil = ""
    
    For i = 1 To Len(Teks)
        c = AscW(Mid(Teks, i, 1))
        If Not ((c >= 1611 And c <= 1631) Or c = 1648) Then
            Hasil = Hasil & Mid(Teks, i, 1)
        End If
    Next i
    
    HapusHarakat = Trim(Hasil)
End Function

Public Function Tanggal(ByVal d As Variant, _
                        Optional ByVal gaya As String = "modern", _
                        Optional ByVal isHijri As Boolean = False, _
                        Optional ByVal vHarakat As Boolean = False) As String

    ' --- 1. Validasi Input (Support TODAY() & Sel Kosong) ---
    If IsEmpty(d) Or d = 0 Then
        d = Date
    ElseIf Not IsDate(d) Then
        Tanggal = "Khatha'": Exit Function
    End If
    
    Dim tgl As Date: tgl = CDate(d)
    Dim dd As Integer, mm As Integer, yy As Long, ww As Integer
    
    ' --- 2. Definisi Harakat ---
    Dim f As String: f = IIf(vHarakat, ChrW(1614), "")
    Dim d_ As String: d_ = IIf(vHarakat, ChrW(1615), "")
    Dim k As String: k = IIf(vHarakat, ChrW(1616), "")
    Dim s As String: s = IIf(vHarakat, ChrW(1618), "")
    Dim sy As String: sy = IIf(vHarakat, ChrW(1617), "")
    Dim ft As String: ft = IIf(vHarakat, ChrW(1611), "")

    ' --- 3. Ekstraksi Data (Ambil Hari di Sini) ---
    If isHijri Then
        VBA.Calendar = vbCalHijri
        dd = Day(tgl): mm = Month(tgl): yy = Year(tgl): ww = Weekday(tgl)
        VBA.Calendar = vbCalGreg
    Else
        dd = Day(tgl): mm = Month(tgl): yy = Year(tgl): ww = Weekday(tgl)
    End If

    ' --- 4. Nama Hari ---
    Static nH As Variant
    If IsEmpty(nH) Or vHarakat Then ' Re-init jika harakat berubah
        nH = Array("", _
             HAM & f & haa & f & dal & d_, _
             ALI & k & THA & s & NUN & f & YAA & s & NUN & k, _
             THA & sy & d_ & LAM & f & ALI & THA & f & ALI & HAM & d_, _
             HAM & f & RHO & s & baa & k & AIN & f & ALI & HZA & d_, _
             kha & f & mim & k & YAA & s & sin & d_, _
             jim & d_ & mim & d_ & AIN & f & MAR & d_, _
             sin & sy & f & baa & s & TAA & d_)
    End If

    ' --- 5. Nama Bulan ---
    Static bM As Variant, bH As Variant
    If IsEmpty(bM) Or vHarakat Then
        bM = Array("", _
            YAA & f & NUN & f & ALI & YAA & k & RHO & f, _
            FHA & k & baa & s & RHO & f & ALI & YAA & k & RHO & f, _
            mim & f & ALI & RHO & k & sin & f, _
            HAM & f & baa & s & RHO & k & YAA & s & LAM & f, _
            mim & f & ALI & YAA & d_ & WAW & f, _
            YAA & d_ & WAW & s & NUN & k & YAA & d_ & WAW & f, _
            YAA & d_ & WAW & s & LAM & k & YAA & d_ & WAW & f, _
            HAM & f & GHA & d_ & sin & s & THO & d_ & sin & f, _
            sin & k & baa & s & TAA & f & mim & s & baa & k & RHO & f, _
            HAM & d_ & kaf & s & TAA & d_ & WAW & s & baa & k & RHO & f, _
            NUN & d_ & WAW & f & f & mim & s & baa & k & RHO & f, _
            dal & k & YAA & s & f & mim & s & baa & k & RHO & f)
            
        bH = Array("", _
            mim & d_ & haa & f & RHO & sy & f & mim & k, _
            sho & f & FHA & f & RHO & k, _
            RHO & f & baa & k & YAA & f & AIN & k & " " & ALI & LAM & f & HAM & s & WAW & sy & f & LAM & k, _
            RHO & f & baa & k & YAA & f & AIN & k & " " & ALI & LAM & f & ALI & f & kha & f & RHO & k, _
            jim & d_ & mim & f & ALI & dal & f & MAQ & " " & ALI & LAM & f & HAM & s & WAW & f & LAM & f & MAQ & k, _
            jim & d_ & mim & f & ALI & dal & f & MAQ & " " & ALI & LAM & f & ALI & f & kha & k & RHO & f & MAR & k, _
            RHO & f & jim & f & baa & k, _
            syi & f & AIN & s & baa & f & ALI & NUN & k, _
            RHO & f & mim & f & DHA & f & ALI & NUN & f, _
            syi & f & WAW & sy & f & ALI & LAM & k, _
            DZA & d_ & WAW & s & " " & ALI & LAM & f & QAF & k & AIN & s & dal & f & MAR & k, _
            DZA & d_ & WAW & s & " " & ALI & LAM & f & haa & k & jim & sy & f & MAR & k)
    End If

    ' --- 6. Terbilang Angka ---
    Dim TargetStyle As Integer
    Select Case LCase(Trim(gaya))
        Case "klasik", "2": TargetStyle = 2
        Case "sastra", "3": TargetStyle = 3
        Case Else: TargetStyle = 1
    End Select

    Dim strH As String, strT As String
    strH = Core_GabungAngkaAr(CStr(dd), 1, TargetStyle, 2)
    strT = Core_GabungAngkaAr(CStr(yy), 1, TargetStyle, 2)
    
    If vHarakat Then
        strH = BeriHarakat(strH, 2, TargetStyle)
        strT = BeriHarakat(strT, 2, TargetStyle)
    End If

    ' --- 7. Label Kalender ---
    Dim lbl As String
    If isHijri Then
        lbl = HAA_besar & k & jim & s & RHO & k & YAA & sy & ft & ALI
    Else
        lbl = mim & k & YAA & s & LAM & f & ALI & dal & k & YAA & sy & ft & ALI
    End If

    ' --- 8. Final Output ---
    Dim EL As String: EL = ALI & LAM & s
    Dim strB As String: strB = IIf(isHijri, bH(mm), bM(mm))
    
    ' Menggunakan variabel ww (weekday) yang sudah diambil sebelumnya
    Tanggal = EL & nH(ww) & ChrW(1548) & " " & strH & " " & strB & " " & strT & " " & lbl
End Function

Public Function Clock_ToTeks(ByVal n As Variant, _
                    Optional ByVal irab As String = "u", _
                    Optional ByVal gaya As String = "modern", _
                    Optional ByVal vHarakat As Boolean = False, _
                    Optional ByVal ShowVisual As Boolean = False, _
                    Optional ByVal StyleVisual As ClockStyle = SolidBlock) As String

    ' ==================================================================================
    ' FUNGSI: Mengonversi jam ke Terbilang Arab (Ordinal Jam + Kardinal Menit)
    ' ==================================================================================

    Dim f As String: f = IIf(vHarakat, ChrW(1614), "")   ' Fathah
    Dim d_ As String: d_ = IIf(vHarakat, ChrW(1615), "")  ' Dammah
    Dim k As String: k = IIf(vHarakat, ChrW(1616), "")   ' Kasrah
    Dim s_ As String: s_ = IIf(vHarakat, ChrW(1618), "")  ' Sukun
    Dim sy As String: sy = IIf(vHarakat, ChrW(1617), "") ' Tasydid
    Dim ft As String: ft = IIf(vHarakat, ChrW(1611), "") ' Tanwin Fathah

    Dim TargetStyle As Integer
    Dim gayaLcase As String: gayaLcase = LCase(Trim(gaya))
    Select Case gayaLcase
        Case "klasik", "2": TargetStyle = 2
        Case "sastra", "3": TargetStyle = 3
        Case Else: TargetStyle = 1
    End Select
    
    Dim TargetIrab As IrabArab
    Select Case LCase(Trim(irab))
        Case "u", "marfu", "1": TargetIrab = 1
        Case "a", "mansub", "2": TargetIrab = 2
        Case "i", "majrur", "3": TargetIrab = 3
        Case Else: TargetIrab = 1
    End Select
    
    ' --- Validasi Waktu (Support NOW(), TODAY(), atau Sel Kosong) ---
    Dim t As Date
    On Error Resume Next
    If IsEmpty(n) Or n = 0 Then
        t = Now ' Default ke waktu saat ini jika input kosong
    ElseIf IsNumeric(n) Then
        t = CDate(n)
    Else
        t = CDate(CStr(n))
    End If
    
    If Err.Number <> 0 Then Clock_ToTeks = "Khatha'": Exit Function
    On Error GoTo 0

    ' --- Komponen Waktu ---
    Dim H As Integer: H = Hour(t)
    Dim m As Integer: m = Minute(t)
    Dim s As Integer: s = Second(t)
    
    ' --- Generate Visual Jam ---
    Dim resVisual As String: resVisual = ""
    If ShowVisual Then
        Dim B As String, Dn As String, T_ As String, U_ As String, sp As String
        Dim N0, N1, N2, N3, N4, N5, N6, N7, N8, N9, NC
        sp = " "

        Select Case StyleVisual
            Case ArabicArqam ' --- STYLE ANGKA ARAB ---
                B = ChrW(9608): T_ = ChrW(9604): U_ = ChrW(9600): sp = " "
                N0 = Array(sp & sp & sp, sp & T_ & sp, B & sp & B, sp & U_ & sp, sp & sp & sp)
                N1 = Array(sp & B & sp, sp & B & sp, sp & B & sp, sp & B & sp, sp & B & sp)
                N2 = Array(B & T_ & T_ & B, sp & sp & sp & B, sp & sp & sp & B, sp & sp & sp & B, sp & sp & sp & B)
                N3 = Array(B & T_ & B & T_ & B, sp & sp & sp & sp & B, sp & sp & sp & sp & B, sp & sp & sp & sp & B, sp & sp & sp & sp & B)
                N4 = Array(sp & T_ & T_ & B, B & sp & sp & U_, U_ & T_ & T_ & B, B & sp & sp & U_, U_ & U_ & U_ & sp)
                N5 = Array(sp & T_ & T_ & sp, B & sp & sp & B, B & sp & sp & B, B & sp & sp & B, sp & U_ & U_ & sp)
                N6 = Array(B & T_ & T_ & B, B & sp & sp & sp, B & sp & sp & sp, B & sp & sp & sp, B & sp & sp & sp)
                N7 = Array(B & sp & sp & B, B & sp & sp & B, sp & B & B & sp, sp & B & B & sp, sp & sp & sp & sp)
                N8 = Array(sp & sp & sp & sp, sp & B & B & sp, sp & B & B & sp, B & sp & sp & B, B & sp & sp & B)
                N9 = Array(sp & T_ & T_ & B, B & sp & sp & B, U_ & T_ & T_ & B, sp & sp & sp & B, sp & sp & sp & B)
                
                ' Colon (:) - Titik Dua Khas Digital
                NC = Array(sp, ChrW(9673), sp, ChrW(9673), sp)
                
            Case DotMatrix ' --- STYLE TITIK-TITIK ---
                B = ChrW(9679): Dn = ChrW(9675)
                N0 = Array(sp & B & B & B & sp, B & Dn & Dn & Dn & B, B & Dn & B & Dn & B, B & Dn & Dn & Dn & B, sp & B & B & B & sp)
                N1 = Array(sp & sp & B & sp & sp, sp & B & B & sp & sp, sp & sp & B & sp & sp, sp & sp & B & sp & sp, sp & B & B & B & sp)
                N2 = Array(sp & B & B & B & sp, sp & Dn & Dn & Dn & B, sp & B & B & B & sp, B & Dn & Dn & Dn & sp, sp & B & B & B & B)
                N3 = Array(sp & B & B & B & sp, sp & Dn & Dn & Dn & B, sp & B & B & B & sp, sp & Dn & Dn & Dn & B, sp & B & B & B & sp)
                N4 = Array(B & sp & sp & sp & B, B & sp & sp & sp & B, sp & B & B & B & B, sp & sp & sp & sp & B, sp & sp & sp & sp & B)
                N5 = Array(B & B & B & B & B, B & sp & sp & sp & sp, B & B & B & B & sp, sp & sp & sp & sp & B, B & B & B & B & sp)
                N6 = Array(sp & B & B & B & sp, B & sp & sp & sp & sp, B & B & B & B & sp, B & sp & sp & sp & B, sp & B & B & B & sp)
                N7 = Array(B & B & B & B & B, sp & sp & sp & B & sp, sp & sp & B & sp & sp, sp & B & sp & sp & sp, B & sp & sp & sp & sp)
                N8 = Array(sp & B & B & B & sp, B & sp & sp & sp & B, sp & B & B & B & sp, B & sp & sp & sp & B, sp & B & B & B & sp)
                N9 = Array(sp & B & B & B & sp, B & sp & sp & sp & B, sp & B & B & B & B, sp & sp & sp & sp & B, sp & B & B & B & sp)
                NC = Array(sp, B, sp, B, sp)
    
            Case DoubleLine ' --- STYLE GARIS GANDA ---
                Dim TL As String: TL = ChrW(9556): Dim TR As String: TR = ChrW(9559)
                Dim bL As String: bL = ChrW(9562): Dim BR As String: BR = ChrW(9565)
                Dim H_ As String: H_ = ChrW(9552)
                Dim V As String: V = ChrW(9553)
                
                N0 = Array(TL & H_ & H_ & H_ & TR, V & "   " & V, V & "   " & V, V & "   " & V, bL & H_ & H_ & H_ & BR)
                N1 = Array("  " & TL & H_ & TR, "   " & V & " ", "   " & V & " ", "   " & V & " ", "  " & bL & H_ & BR)
                N2 = Array(TL & H_ & H_ & H_ & TR, "    " & V, TL & H_ & H_ & H_ & BR, V & "    ", bL & H_ & H_ & H_ & BR)
                N3 = Array(TL & H_ & H_ & H_ & TR, "    " & V, "  " & H_ & H_ & H_ & BR, "    " & V, bL & H_ & H_ & H_ & BR)
                N4 = Array(V & "   " & V, V & "   " & V, bL & H_ & H_ & H_ & TR, "    " & V, "    " & V)
                N5 = Array(TL & H_ & H_ & H_ & TR, V & "    ", bL & H_ & H_ & H_ & TR, "    " & V, bL & H_ & H_ & H_ & BR)
                N6 = Array(TL & H_ & H_ & H_ & TR, V & "    ", TL & H_ & H_ & H_ & TR, V & "   " & V, bL & H_ & H_ & H_ & BR)
                N7 = Array(TL & H_ & H_ & H_ & TR, "    " & V, "   " & V & " ", "  " & V & "  ", " " & V & "   ")
                N8 = Array(TL & H_ & H_ & H_ & TR, V & "   " & V, TL & H_ & H_ & H_ & TR, V & "   " & V, bL & H_ & H_ & H_ & BR)
                N9 = Array(TL & H_ & H_ & H_ & TR, V & "   " & V, bL & H_ & H_ & H_ & TR, "    " & V, bL & H_ & H_ & H_ & BR)
                NC = Array(sp, ChrW(9553), sp, ChrW(9553), sp)
            
            Case ShadowStyle ' --- STYLE BAYANGAN ---
                Dim S1 As String: S1 = ChrW(9617): Dim S2 As String: S2 = ChrW(9618): Dim S3 As String: S3 = ChrW(9619)
                B = S3: sp = " "
                
                N0 = Array(B & B & B & B, B & sp & sp & B, B & sp & sp & B, B & sp & sp & B, B & B & B & B)
                N1 = Array(sp & sp & B & sp, sp & B & B & sp, sp & sp & B & sp, sp & sp & B & sp, sp & B & B & B)
                N2 = Array(B & B & B & B, sp & sp & sp & B, B & B & B & B, B & sp & sp & sp, B & B & B & B)
                N3 = Array(B & B & B & B, sp & sp & sp & B, sp & B & B & B, sp & sp & sp & B, B & B & B & B)
                N4 = Array(B & sp & sp & B, B & sp & sp & B, B & B & B & B, sp & sp & sp & B, sp & sp & sp & B)
                N5 = Array(B & B & B & B, B & sp & sp & sp, B & B & B & B, sp & sp & sp & B, B & B & B & B)
                N6 = Array(B & B & B & B, B & sp & sp & sp, B & B & B & B, B & sp & sp & B, B & B & B & B)
                N7 = Array(B & B & B & B, sp & sp & sp & B, sp & sp & B & sp, sp & B & sp & sp, B & sp & sp & sp)
                N8 = Array(B & B & B & B, B & sp & sp & B, B & B & B & B, B & sp & sp & B, B & B & B & B)
                N9 = Array(B & B & B & B, B & sp & sp & B, B & B & B & B, sp & sp & sp & B, B & B & B & B)
                NC = Array(sp, S2, sp, S2, sp)

            Case Else ' --- STYLE BALOK TEBAL (SolidBlock) ---
                B = ChrW(9608): T_ = ChrW(9604): U_ = ChrW(9600)
                N0 = Array(sp & T_ & T_ & T_ & sp, B & sp & sp & T_ & B, B & sp & B & sp & B, B & U_ & sp & sp & B, sp & U_ & U_ & U_ & sp)
                N1 = Array(sp & T_ & T_ & sp & sp, sp & sp & B & sp & sp, sp & sp & B & sp & sp, sp & sp & B & sp & sp, sp & U_ & U_ & U_ & sp)
                N2 = Array(sp & T_ & T_ & T_ & sp, B & sp & sp & sp & B, sp & sp & T_ & T_ & U_, B & sp & sp & sp & sp, U_ & U_ & U_ & U_ & U_)
                N3 = Array(sp & T_ & T_ & T_ & sp, U_ & sp & sp & sp & B, sp & sp & U_ & U_ & T_, T_ & sp & sp & sp & B, sp & U_ & U_ & U_ & sp)
                N4 = Array(T_ & sp & sp & sp & T_, B & sp & sp & sp & B, U_ & U_ & U_ & U_ & B, sp & sp & sp & sp & B, sp & sp & sp & sp & U_)
                N5 = Array(T_ & T_ & T_ & T_ & T_, B & sp & sp & sp & U_, U_ & U_ & U_ & U_ & T_, sp & sp & sp & sp & B, sp & U_ & U_ & U_ & sp)
                N6 = Array(sp & T_ & T_ & T_ & sp, B & sp & sp & sp & U_, B & U_ & U_ & U_ & T_, B & sp & sp & sp & B, sp & U_ & U_ & U_ & sp)
                N7 = Array(T_ & T_ & T_ & T_ & T_, U_ & sp & sp & T_ & U_, sp & sp & B & sp & sp, sp & sp & B & sp & sp, sp & sp & U_ & sp & sp)
                N8 = Array(sp & T_ & T_ & T_ & sp, B & sp & sp & sp & B, T_ & U_ & U_ & U_ & T_, B & sp & sp & sp & B, sp & U_ & U_ & U_ & sp)
                N9 = Array(sp & T_ & T_ & T_ & sp, B & sp & sp & sp & B, U_ & T_ & T_ & T_ & B, T_ & sp & sp & sp & B, sp & U_ & U_ & U_ & sp)
                NC = Array(sp, T_, sp, U_, sp)
        End Select

        ' --- Loop Konstruksi Visual ---
        Dim vResult(0 To 4) As String
        Dim timeStr As String
        
        ' Paksa format menggunakan ":" tanpa terpengaruh regional setting
        timeStr = Format(H, "00") & ":" & Format(m, "00")
        
        Dim i As Integer, char As String, r As Integer, curr As Variant
        
        For i = 1 To Len(timeStr)
            char = Mid(timeStr, i, 1)
            curr = Empty
            
            Select Case char
                Case "0": curr = N0: Case "1": curr = N1: Case "2": curr = N2
                Case "3": curr = N3: Case "4": curr = N4: Case "5": curr = N5
                Case "6": curr = N6: Case "7": curr = N7: Case "8": curr = N8
                Case "9": curr = N9: Case ":": curr = NC
                Case Else: curr = NC
            End Select
            
            ' Gambar ke array vResult
            If Not IsEmpty(curr) Then
                For r = 0 To 4
                    vResult(r) = vResult(r) & curr(r) & sp
                Next r
            End If
        Next i
        resVisual = Join(vResult, vbCrLf) & vbCrLf & String(30, "-") & vbCrLf
End If
    
    ' Kata "As-Sa'ah" (Jam)
    Dim AS_SAA_AH As String: AS_SAA_AH = ALI & LAM & sy & f & sin & f & ALI & AIN & f & MAR & d_
    
    Dim WA As String: WA = " " & WAW & f & " "
    Dim W_AL As String: W_AL = WA & ALI & LAM
    Dim tMenit As String: tMenit = " " & dal & f & QAF & k & YAA & s_ & QAF & f & MAR & ft
    Dim tDetik As String: tDetik = " " & ChrW(1579) & f & ChrW(1575) & ChrW(1606) & k & ChrW(1610) & f & ChrW(1577) & ft
    Dim tSore As String: tSore = mim & f & sin & f & ALI & ChrW(1569) & ft
    Dim tPagi As String: tPagi = sho & f & baa & f & ALI & haa & ft & ALI
    Dim ILLA As String: ILLA = " " & ALI & k & LAM & sy & f & ALI & " "
    
    ' Cek dulu menitnya sebelum menentukan Jam
    Dim isKurang As Boolean: isKurang = (m = 40 Or m = 45)
    Dim hFasih As Integer: hFasih = IIf(isKurang, H + 1, H)
    
    ' A. Jam (Ordinal Muannats 1-12)
    Dim h12 As Integer: h12 = hFasih Mod 12: If h12 = 0 Then h12 = 12
    Dim resJam As String: resJam = GetArabicOrdinal(h12, 2) ' 2 = Muannats
    resJam = BeriHarakat(resJam, 1, TargetStyle)

    ' B. Menit (Kardinal Muannats 1-59)
    Dim resMenit As String
    Dim pilihMenit As Byte: pilihMenit = 2 'gaya rumahan
    
    Select Case pilihMenit
        Case 1
            If m > 0 Then
                Dim txtMenit As String: txtMenit = Core_GabungAngkaAr(CStr(m), 2, TargetStyle, TargetIrab)
                txtMenit = BeriHarakat(txtMenit, TargetIrab, TargetStyle)
                resMenit = WA & txtMenit & tMenit
            End If
            
     Case 2 ' gaya rumahan
            If m > 0 Then
                Select Case m
                    Case 5 ' Lebih 5 Menit (Khamsu Daqa'iq)
                        resMenit = WAW & kha & f & mim & s_ & sin & d_ & " " & dal & f & QAF & ALI & HAM & k & QAF & f
                    Case 10 ' Lebih 10 Menit (Asyru Daqa'iq)
                        resMenit = WA & AIN & f & syi & s_ & RHO & d_ & " " & dal & f & QAF & ALI & HAM & k & QAF & f
                    Case 15 ' Seperempat (Ar-Rubu'u)
                        resMenit = W_AL & RHO & sy & d_ & baa & d_ & AIN & d_
                    Case 20 ' Sepertiga (At-Tsulutsu)
                        resMenit = W_AL & THA & sy & d_ & LAM & d_ & THA & d_
                    Case 30 ' Setengah (An-Nishfu)
                        resMenit = W_AL & NUN & sy & k & sho & s_ & FHA & d_
                    Case 40 ' Kurang Sepertiga (Illa At-Tsulutsa)
                        ' Harakat akhir Fathah (f) karena Mustatsna
                        resMenit = ILLA & ALI & LAM & THA & sy & d_ & LAM & d_ & THA & f
                    Case 45 ' Kurang Seperempat (Illa Ar-Rubu'a)
                        ' Harakat akhir Fathah (f) karena Mustatsna
                        resMenit = ILLA & ALI & LAM & RHO & sy & d_ & baa & d_ & AIN & f
                    Case 50 ' Kurang 10 Menit (Illa 'Asyra Daqa'iq)
                        resMenit = ILLA & AIN & f & syi & s_ & RHO & f & " " & dal & f & QAF & ALI & HAM & k & QAF & f
                    Case 55 ' Kurang 5 Menit (Illa Khamsa Daqa'iq)
                        resMenit = ILLA & kha & f & mim & s_ & sin & f & " " & dal & f & QAF & ALI & HAM & k & QAF & f
                    Case Else
                        ' Menit Standar Kardinal (1-59 selain di atas)
                        Dim txtM As String: txtM = Core_GabungAngkaAr(CStr(m), 2, TargetStyle, TargetIrab)
                        resMenit = WA & BeriHarakat(txtM, TargetIrab, TargetStyle) & tMenit
                End Select
            End If
    End Select
    
    ' C. Detik (Kardinal Muannats 1-59)
    Dim resDetik As String
    If s > 0 Then
        resDetik = Core_GabungAngkaAr(CStr(s), 2, TargetStyle, TargetIrab)
        resDetik = WA & BeriHarakat(resDetik, TargetIrab, TargetStyle) & tDetik
    End If
    
    ' D. Periode (Shabahan / Masa'an)
    Dim resPeriod As String
    ' Menggunakan H asli untuk periode waktu
    If H >= 12 Then
        resPeriod = " " & tSore
    Else
        resPeriod = " " & tPagi
    End If

    ' --- 5. Output Akhir ---
    'Clock_ToTeks = resVisual & AS_SAA_AH & " " & resJam & resMenit & resDetik & resPeriod
    Clock_ToTeks = resVisual & AS_SAA_AH & " " & resJam & resMenit & resPeriod
End Function

