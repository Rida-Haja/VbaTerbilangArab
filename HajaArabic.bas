Option Explicit

' ==============================================================================
' DEFINISI TIPE DATA & ENUMERASI
' ==============================================================================
Public Enum GenderArab
    Muzakkar = 1    ' Laki-laki
    muannas = 2     ' Perempuan
End Enum

Public Enum GayaArab
    Modern = 1 ' Standar Modern (Spasi jelas, istilah umum)
    klasik = 2 ' Fusha Turats (Sambung Wawu, istilah klasik)
    sastra = 3 ' Gaya Sastra/Puitis
End Enum

Public Enum IrabArab
    Marfu = 1       ' Nominative (Dhommah/Waw-Nun)
    Mansub = 2      ' Accusative (Fathah/Ya-Nun)
    Majrur = 3      ' Genitive (Kasrah/Ya-Nun)
End Enum

Public Enum LevelHarakat
    Gundul = 0      ' Tanpa harakat sama sekali
    IrabSaja = 1    ' Hanya harakat di huruf terakhir setiap kata
    FullMasykul = 2 ' Harakat lengkap (Fathah, Kasrah, Dhommah, Sukun, Tasydid)
End Enum

Public Enum ClockStyle
    SolidBlock = 0
    DotMatrix = 1
    ArabicArqam = 2
    DoubleLine = 3
    ShadowStyle = 4
End Enum

Public ArabData As clsTerbilangArab
Private NextTick As Date

' Prosedur untuk memastikan Class sudah dibuat
Private Sub CekInisialisasi()
    If ArabData Is Nothing Then
        Set ArabData = New clsTerbilangArab
    End If
End Sub

Public Sub ResetDanUpdateSkala()
    Set ArabData = New clsTerbilangArab
    
    MsgBox "Skala Arab berhasil diperbarui ke: " & ArabData.SkalaAktif, _
           vbInformation, "Update Berhasil"
End Sub

Sub teshaja()
Set ArabData = New clsTerbilangArab
 MsgBox ArabData.Core_GabungAngkaAr(1234567890, 2, 1, 2, False)
End Sub
                    
'=================== BAB TERBILANG_ARAB =======================================
Public Function TERBILANG_ARAB(ByVal Angka As Variant, _
    Optional ByVal Mode As String = "umum", _
    Optional ByVal gender As String = "m", _
    Optional ByVal irab As String = "u", _
    Optional ByVal gaya As String = "modern", _
    Optional ByVal Parameter_Negara As String = "id", _
    Optional ByVal PakaiHarakat As Byte = 0, _
    Optional ByVal isIdhafah As Boolean = False) As String

    On Error GoTo ErrHandler
    
    If IsEmpty(Angka) Or Trim(Angka) = "" Then
        TERBILANG_ARAB = ""
        Exit Function
    End If
    
    Application.Volatile
    CekInisialisasi
    Dim Txt As clsTerbilangArab
        Set Txt = ArabData
    
    Dim TargetIrab As IrabArab
    Dim TargetGender As Integer
    Dim TargetStyle As Integer
    
    TargetIrab = Switch(LCase(irab) = "a" Or irab = "2", 2, _
                        LCase(irab) = "i" Or irab = "3", 3, _
                        True, 1)
                        
    TargetGender = IIf(LCase(gender) = "f" Or gender = "p" Or gender = "2", 2, 1)
    
    TargetStyle = Switch(LCase(gaya) = "klasik" Or gaya = "2", 2, _
                         LCase(gaya) = "sastra" Or gaya = "3", 3, _
                         True, 1)

    Dim strNum As String
    strNum = Replace(Replace(Trim(CStr(Angka)), " ", ""), ",", ".")
    
    If strNum = "" Or strNum = "0" Then
        TERBILANG_ARAB = Txt.Kata(0, TargetGender)
        Exit Function
    End If
    
    ' Validasi Batas Maksimal Digit
    Static BatasMaks As Long
    If BatasMaks = 0 Then
        On Error Resume Next ' Jaga-jaga jika Sheet1 tidak ada
        BatasMaks = val(ThisWorkbook.Worksheets("Sheet1").Range("C7").Value)
        If BatasMaks = 0 Then BatasMaks = 1000
        On Error GoTo ErrHandler
    End If
    
    If Len(strNum) > BatasMaks Then
        TERBILANG_ARAB = "[Khatha': Maks " & BatasMaks & " Digit]"
        Exit Function
    End If

    Dim tKoma As String:    tKoma = Txt.FHA & Txt.ALI & Txt.sho & Txt.LAM & Txt.MAR   ' Fashilah
    Dim tNegatif As String: tNegatif = Txt.sin & Txt.LAM & Txt.baa             ' Salib
    Dim tWa As String:      tWa = " " & Txt.WAW & " "                  ' Wa (dan)
    
    ' Cek Negatif
    Dim prefix As String
    If Left$(strNum, 1) = "-" Then
        prefix = tNegatif & " "
        strNum = Mid$(strNum, 2)
    End If

    Dim bulat As String, desimal As String
    Dim parts() As String
    
    parts = Split(strNum, ".")
    bulat = parts(0)
    If UBound(parts) > 0 Then desimal = parts(1)

    Dim res As String, idUnit As Integer
    Dim modeLcase As String: modeLcase = LCase(Trim(Mode))

    Select Case modeLcase
        Case "umum", "angka", "kardinal"
            ' --- MODE KARDINAL (Standar + Koma) ---
            Dim fb As String
            Dim fd As String
            
            fb = IIf(Len(Replace(bulat, "0", "")) = 0, _
                     Txt.Kata(0, TargetGender), _
                     Txt.Core_GabungAngkaAr(bulat, TargetGender, TargetStyle, TargetIrab, isIdhafah))
            
            If Len(desimal) > 0 And Replace(desimal, "0", "") <> "" Then
                fd = " " & tKoma & " " & Txt.Core_GabungAngkaAr(desimal, TargetGender, TargetStyle, TargetIrab, isIdhafah)
            End If
            
            res = prefix & fb & fd
            TERBILANG_ARAB = res
            TERBILANG_ARAB = Txt.BeriHarakat(Trim$(res), TargetIrab, TargetStyle, modeLcase, isIdhafah, PakaiHarakat)

            Exit Function

        Case "eja", "digit", "2"
            ' --- MODE EJA (Per Digit) ---
            Dim i As Long
            For i = 1 To Len(bulat)
                res = res & Txt.Kata(val(Mid$(bulat, i, 1)), TargetGender) & " "
            Next i
            
            If Len(desimal) > 0 Then
                res = res & tKoma & " "
                For i = 1 To Len(desimal)
                    res = res & Txt.Kata(val(Mid$(desimal, i, 1)), TargetGender) & " "
                Next i
            End If
            res = prefix & Trim$(res)
            TERBILANG_ARAB = Txt.BeriHarakat(res, TargetIrab, TargetStyle, modeLcase, False, PakaiHarakat, True)
            Exit Function

        Case "uang", "currency"
            TERBILANG_ARAB = ArabCurrency(Angka, Parameter_Negara, PakaiHarakat, TargetIrab)
            Exit Function

        Case "urutan", "ordinal"
            Dim idOrd As Integer
            idOrd = IIf(Parameter_Negara = "", IIf(TargetGender = 1, 60, 61), val(Parameter_Negara))
            TERBILANG_ARAB = ArabUniversal(Angka, idOrd, True, TargetIrab, TargetStyle, PakaiHarakat)
            Exit Function

        Case "jarak", "km", "meter", "m", "liter", "l", "berat", "kg", "waktu", "hari", "jam", "bulan", "tahun", "buku", "kitab", "halaman", "persen", "benda", "unit"
            ' Mapping ID Unit secara spesifik
            idUnit = Switch(modeLcase = "meter" Or modeLcase = "m", 2, _
                            modeLcase = "jarak" Or modeLcase = "km", 3, _
                            modeLcase = "berat" Or modeLcase = "kg", 5, _
                            modeLcase = "liter" Or modeLcase = "l", 7, _
                            modeLcase = "jam", 10, _
                            modeLcase = "waktu" Or modeLcase = "hari", 11, _
                            modeLcase = "bulan", 12, _
                            modeLcase = "tahun", 13, _
                            modeLcase = "buku", 21, _
                            modeLcase = "halaman", 22, _
                            modeLcase = "persen", 51, _
                            True, val(Parameter_Negara))
            TERBILANG_ARAB = ArabUniversal(Angka, idUnit, False, TargetIrab, TargetStyle, PakaiHarakat)
            Exit Function
            
        Case Else
            If IsNumeric(Mode) Then idUnit = val(Mode)
    End Select

    ' 6. Penanganan Fallback Akhir (Jika tidak masuk Select Case di atas)
    If idUnit > 0 Then
        TERBILANG_ARAB = ArabUniversal(Angka, idUnit, False, TargetIrab, TargetStyle, PakaiHarakat)
    Else
        ' --- Mode Pecahan Umum ---
        Dim mBulat As String, mDes As String
        
        mBulat = IIf(Len(Replace(bulat, "0", "")) = 0, _
                     Txt.Kata(0, TargetGender), _
                     Txt.Core_GabungAngkaAr(bulat, TargetGender, TargetStyle, TargetIrab, isIdhafah))
        
        If Len(desimal) > 0 And Replace(desimal, "0", "") <> "" Then
            mDes = " " & tKoma & " " & Txt.Core_GabungAngkaAr(desimal, TargetGender, TargetStyle, TargetIrab, isIdhafah)
        End If
        
        res = prefix & mBulat & mDes
        TERBILANG_ARAB = Trim$(res)
    End If

    Exit Function

ErrHandler:
TERBILANG_ARAB = "[Khatha': " & Err.Description & "]"
End Function

'=================== BAB ArabCurrency, ArabUniversal, Jam, Tanggal ============
Public Function ArabCurrency(ByVal Angka As Variant, _
                             Optional ByVal CountryCode As String = "id", _
                             Optional ByVal vHarakat As LevelHarakat = 0, _
                             Optional ByVal irab As IrabArab = 1) As String

    ' ==================================================================================
    ' FUNGSI: Mengonversi angka ke terbilang mata uang dalam bahasa Arab
    ' ==================================================================================

    Dim mainResult As String, strAngka As String
    Dim bagianBulat As String, bagianDesimal As String
    Dim posKoma As Long
    Dim currencyText As String, subUnitText As String
    Dim GenderUtama As GenderArab, GenderKecil As GenderArab
    
    ' --- Definisi Harakat ---
    Dim f As String: f = IIf(vHarakat = 2, ChrW(1614), "")   ' Fathah
    Dim d As String: d = IIf(vHarakat = 2, ChrW(1615), "")   ' Dammah
    Dim k As String: k = IIf(vHarakat = 2, ChrW(1616), "")   ' Kasrah
    Dim s As String: s = IIf(vHarakat = 2, ChrW(1618), "")   ' Sukun
    Dim ft As String: ft = IIf(Not vHarakat = 0, ChrW(1611), "") ' Fathatain (an)
    Dim dt As String: dt = IIf(Not vHarakat = 0, ChrW(1612), "") ' Dammatain (un)
    Dim kt As String: kt = IIf(Not vHarakat = 0, ChrW(1613), "") ' Kasratain (in)
    Dim sy As String: sy = ChrW(1617) ' Tasydid
    
    Dim isIdhafah As Boolean, ha As String
    
    CekInisialisasi
    Dim H As clsTerbilangArab
        Set H = ArabData
        
    isIdhafah = False
            ' -> Gunakan Harakat Tunggal (u, a, i)
            If (isIdhafah) Then
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
            

    ' --- Logika Irab untuk Akhiran Dua (Tatsniyah) ---
    Dim IsMarfu As Boolean: IsMarfu = (irab = 1 Or irab = 10)
    Dim AkhiranDua As String: AkhiranDua = f & IIf(IsMarfu, H.ALI, H.YAA & s) & H.NUN & k

    CountryCode = LCase(Trim(CountryCode))
    If CountryCode = "" Then CountryCode = "id"
    
    ' --- Konfigurasi Nama Mata Uang & Ketetapan Gender Gramatikal ---
    Dim wRupiah, wRiyal, wSen, wDinar, wFils, wDirham, wRinggit, wDollar, wYen, wYuan, wHalalah
    wRupiah = H.RHO & d & H.WAW & s & H.baa & k & H.YAA & sy & f & H.MAR
    wSen = H.sin & k & H.NUN & sy
    wRiyal = H.RHO & k & H.YAA & f & H.ALI & H.LAM
    wHalalah = H.HAA_besar & f & H.LAM & f & H.LAM & f & H.MAR
    wDinar = H.dal & k & H.YAA & s & H.NUN & f & H.ALI & H.RHO
    wFils = H.FHA & k & H.LAM & s
    wDirham = H.dal & k & H.RHO & s & H.HAA_besar & f & H.mim
    wRinggit = H.RHO & k & H.YAA & H.NUN & s & H.jim & k & H.YAA & s & H.TAA
    wDollar = H.dal & d & H.WAW & H.LAM & f & H.ALI & H.RHO
    wYen = H.YAA & f & H.NUN
    wYuan = H.YAA & d & H.WAW & f & H.ALI & H.NUN

    ' Default Gender
    GenderUtama = Muzakkar: GenderKecil = Muzakkar

    Select Case CountryCode
        Case "id": currencyText = wRupiah:  subUnitText = wSen:     GenderUtama = muannas
        Case "sa": currencyText = wRiyal:   subUnitText = wHalalah: GenderKecil = muannas
        Case "kw": currencyText = wDinar:   subUnitText = wFils
        Case "ae": currencyText = wDirham:  subUnitText = wFils
        Case "qa": currencyText = wRiyal:   subUnitText = wDirham
        Case "my": currencyText = wRinggit: subUnitText = wSen
        Case "sg", "bn": currencyText = wDollar: subUnitText = wSen
        Case "jp": currencyText = wYen:     subUnitText = ""
        Case "cn": currencyText = wYuan:    subUnitText = ""
        Case Else: currencyText = "":       subUnitText = ""
    End Select

    ' --- Parsing Angka ---
    strAngka = Replace(Replace(CStr(Angka), " ", ""), ",", ".")
    posKoma = InStr(strAngka, ".")
    If posKoma > 0 Then
        bagianBulat = Left(strAngka, posKoma - 1): bagianDesimal = Mid(strAngka, posKoma + 1)
        If Len(bagianDesimal) = 1 Then bagianDesimal = bagianDesimal & "0"
        bagianDesimal = Left(bagianDesimal, 2)
    Else
        bagianBulat = strAngka: bagianDesimal = "0"
    End If

    ' --- LOGIKA BAGIAN BULAT ---
    If bagianBulat = "0" Or bagianBulat = "" Then
        Dim T0 As String: T0 = H.Core_GabungAngkaAr(0, GenderUtama, Modern, irab, True)
        T0 = H.BeriHarakat(T0, irab, Modern, , True, vHarakat)
        mainResult = T0 & " " & currencyText & ha
    Else
        Dim TeksB As String, Ma_dudB As String, uB As Integer
        Dim tHapus As String, tAkhir As String
        
        uB = val(Right("00" & bagianBulat, 2))
        TeksB = H.Core_GabungAngkaAr(bagianBulat, GenderUtama, Modern, irab, False)
        
        Dim THA As String: THA = IIf(ha = ft Or ha = f, IIf(Right(Trim(TeksB), 1) = ChrW(1577), ha, ha & ChrW(1575)), ha)
        TeksB = H.BeriHarakat(TeksB, irab, Modern, , True, vHarakat)
        
        If bagianBulat = "1" Then
        ' 1: Sebutkan Nama Mata Uang saja
            Ma_dudB = currencyText & THA
            mainResult = Ma_dudB
        ElseIf uB = 1 And bagianBulat <> "1" Then
            tHapus = H.Core_GabungAngkaAr(1, GenderUtama, Modern, irab, True)
            tHapus = H.BeriHarakat(tHapus, irab, Modern, , True, vHarakat)
            If Right(TeksB, Len(tHapus)) = tHapus Then TeksB = Trim(Left(TeksB, Len(TeksB) - Len(tHapus)))
            Ma_dudB = currencyText & THA
                tAkhir = H.Core_GabungAngkaAr(1, GenderUtama, Modern, irab, False)
                mainResult = TeksB & " " & Ma_dudB '& " " & H.BeriHarakat(tAkhir, irab, modern, , False, vHarakat)
        ElseIf uB = 2 And bagianBulat <> "12" Then
        ' 2: Gunakan bentuk Dual (Tatsniyah)
            If GenderUtama = muannas Then
                Ma_dudB = Left(currencyText, Len(currencyText) - IIf(vHarakat = Gundul, 1, 2)) & f & ChrW(1578) & AkhiranDua
            Else
                Ma_dudB = currencyText & AkhiranDua
            End If
            
            If bagianBulat = "2" Then
                mainResult = Ma_dudB
            Else
                tHapus = H.Core_GabungAngkaAr(2, GenderUtama, Modern, irab, True)
                tHapus = H.BeriHarakat(tHapus, irab, Modern, , True, vHarakat)
                If Right(TeksB, Len(tHapus)) = tHapus Then TeksB = Trim(Left(TeksB, Len(TeksB) - Len(tHapus)))

                tAkhir = H.Core_GabungAngkaAr(2, GenderUtama, Modern, irab, False)
                mainResult = TeksB & " " & Ma_dudB '& " " & H.BeriHarakat(tAkhir, irab, modern, , False, vHarakat)
            End If
            
        ElseIf uB >= 3 And uB <= 10 Then
        ' 3-10: Gunakan bentuk Jamak (Jam'u Taksir/Muannas Salim)
            Dim plRupiah, plRiyal, plDinar, plDirham
        
            ' Rupiah: Menghapus Ta Marbuta lalu tambah Alif + Taa (Rupiyaat)
            plRupiah = Left(currencyText, Len(currencyText) - IIf(vHarakat = Gundul, 1, 2)) & f & H.ALI & H.TAA
            
            ' Riyal -> Riyalaat
            plRiyal = H.RHO & k & H.YAA & f & H.ALI & H.LAM & H.ALI & f & H.TAA
            
            ' Dinar -> Dananiir
            plDinar = H.dal & f & H.NUN & f & H.ALI & H.NUN & k & H.YAA & f & H.RHO
            
            ' Dirham -> Darahim
            plDirham = H.dal & f & H.RHO & f & H.ALI & H.HAA_besar & k & H.mim
        
            ' --- Tetapkan Ma'dud ---
            Select Case CountryCode
                Case "id":      Ma_dudB = plRupiah
                Case "sa", "qa": Ma_dudB = plRiyal
                Case "kw":      Ma_dudB = plDinar
                Case "ae":      Ma_dudB = plDirham
                Case Else:      Ma_dudB = currencyText
            End Select

        ' Tambahkan suffix kt di akhir secara otomatis untuk semua case
        Ma_dudB = Ma_dudB & kt
        mainResult = TeksB & " " & Ma_dudB
                
        ElseIf uB >= 11 And uB <= 99 Then
            ' 11-99: Mufrad Manshub (Tanwin Fathah)
            Ma_dudB = currencyText & IIf(GenderUtama = Muzakkar, ft & H.ALI, ft)
            mainResult = TeksB & " " & Ma_dudB
            
        Else
            ' Mufrad Majrur (100, 1000, dst menggunakan Kasratain)
            Ma_dudB = currencyText & kt
            mainResult = TeksB & " " & Ma_dudB
        End If
    End If
    
    ' --- LOGIKA BAGIAN DESIMAL (SEN) ---
    Dim vD As Integer: vD = val(bagianDesimal)
    If vD > 0 And subUnitText <> "" Then
        Dim TeksD As String, Ma_dudD As String
        TeksD = H.Core_GabungAngkaAr(bagianDesimal, GenderKecil, Modern, irab)
        THA = IIf(ha = ft Or ha = f, IIf(Right(Trim(TeksD), 1) = ChrW(1577), ha, ha & ChrW(1575)), ha)
        TeksD = H.BeriHarakat(TeksD, irab, Modern, , True, vHarakat)
        
        If vD = 1 Then
            Ma_dudD = subUnitText & THA
        ElseIf vD = 2 Then
            If GenderKecil = muannas Then
                Ma_dudD = Left(subUnitText, Len(subUnitText) - IIf(vHarakat = Gundul, 1, 2)) & f & ChrW(1578) & AkhiranDua
            Else
                Ma_dudD = subUnitText & AkhiranDua
            End If
        ElseIf vD >= 3 And vD <= 10 Then
            ' Penanganan jamak untuk Sen (Jika diperlukan bentuk jamak spesifik bisa ditambahkan)
            Ma_dudD = subUnitText & kt
        ElseIf vD >= 11 And vD <= 99 Then
            Ma_dudD = subUnitText & IIf(GenderKecil = Muzakkar, ft & H.ALI, ft)
        Else
            Ma_dudD = subUnitText & kt
        End If
        
        mainResult = mainResult & " " & H.WAW & f & " " & IIf(vD = 2, Ma_dudD, TeksD & " " & Ma_dudD)
    End If
    
    ArabCurrency = Trim(mainResult)
End Function

Private Function ArabUniversal(ByVal Angka As Variant, _
                               Optional ByVal unitType As Byte = 0, _
                               Optional ByVal IsOrdinal As Boolean = False, _
                               Optional ByVal irab As IrabArab = 1, _
                               Optional ByVal gaya As GayaArab = Modern, _
                               Optional ByVal vHarakat As LevelHarakat = 0) As String

    Dim f As String, d As String, k As String, s As String
    Dim ft As String, dt As String, kt As String, sy As String
    Dim ha As String, i As Integer, v1 As Variant
    
    ' Inisialisasi Harakat
    f = IIf(vHarakat = 2, ChrW(1614), "")
    d = IIf(vHarakat = 2, ChrW(1615), "")
    k = IIf(vHarakat = 2, ChrW(1616), "")
    s = IIf(vHarakat = 2, ChrW(1618), "")
    ft = IIf(vHarakat <> 0, ChrW(1611), "")
    dt = IIf(vHarakat <> 0, ChrW(1612), "")
    kt = IIf(vHarakat <> 0, ChrW(1613), "")
    sy = ChrW(1617)  ' Tasydid
  
    If vHarakat <> 0 Then
        Select Case LCase$(Trim$(CStr(irab)))
            Case "u", "marfu", "1": ha = IIf(IsOrdinal, d, dt)
            Case "a", "mansub", "2": ha = IIf(IsOrdinal, f, ft)
            Case "i", "majrur", "3": ha = IIf(IsOrdinal, k, kt)
            Case Else: ha = IIf(IsOrdinal, d, dt)
        End Select
    End If

    ' Bersihkan input angka
    Dim strAngka As String: strAngka = Trim$(Replace$(CStr(Angka), " ", ""))
    If strAngka = "" Then strAngka = "0"
    
    CekInisialisasi
    Dim H As clsTerbilangArab
        Set H = ArabData

    Dim V As Integer: V = val(Right$("00" & strAngka, 2))
    Dim IsSatu As Boolean: IsSatu = (strAngka = "1")
    Dim IsDua As Boolean: IsDua = (strAngka = "2")
    Dim IsMarfu As Boolean: IsMarfu = (irab = 1 Or irab = 10)
    Dim AL As String: AL = H.ALI & H.LAM & s
    Dim GenderFinal As GenderArab
    Dim FASHILAH As String: FASHILAH = ChrW(1601) & f & H.ALI & H.sho & k & H.LAM & f & H.MAR & ha ' Fashilah (Pemisah Desimal)
    Dim baseUnit As String, pluralUnit As String
    Dim isMuannasUnit As Boolean: isMuannasUnit = False

    ' Default setting
    GenderFinal = Muzakkar
    pluralUnit = ""

    Select Case unitType
        ' --- KELOMPOK MUZAKKAR ---
        Case 1:  baseUnit = H.sin & f & H.NUN & s & H.TAA & k & H.mim & k & H.TAA & H.RHO ' Centimeter
        Case 2
                 baseUnit = H.mim & k & H.TAA & s & H.RHO ' Meter
                 pluralUnit = H.HAM & f & H.mim & s & H.TAA & f & H.ALI & H.RHO
        Case 3:  baseUnit = H.kaf & k & H.YAA & s & H.LAM & d & H.WAW & H.mim & k & H.TAA & s & H.RHO ' Kilometer
        Case 4
                 baseUnit = H.GHA & f & H.RHO & f & H.ALI & H.mim ' Gram
                 pluralUnit = baseUnit & f & H.ALI & H.TAA
        Case 5:  baseUnit = H.kaf & k & H.YAA & s & H.LAM & d & H.WAW & H.GHA & f & H.RHO & f & H.ALI & H.mim ' Kilogram
        Case 6:  baseUnit = H.THO & d & H.NUN ' Ton
        Case 7:  baseUnit = H.LAM & k & H.TAA & s & H.RHO ' Liter
        Case 11
                 baseUnit = H.YAA & f & H.WAW & s & H.mim ' Yaum
                 pluralUnit = H.HAM & f & H.YAA & sy & f & H.ALI & H.mim
        Case 12
                 baseUnit = H.syi & f & H.HAA_besar & s & H.RHO ' Syahr
                 pluralUnit = H.HAM & f & H.syi & s & H.HAA_besar & d & H.RHO
        Case 20
                 baseUnit = H.syi & f & H.kha & s & H.sho ' Syakhs
                 pluralUnit = H.HAM & f & H.syi & s & H.kha & f & H.ALI & H.sho
        Case 21
                 baseUnit = H.kaf & k & H.TAA & f & H.ALI & H.baa ' Kitab
                 pluralUnit = H.kaf & d & H.TAA & d & H.baa
        Case 40: baseUnit = H.FHA & d & H.WAW & s & H.LAM & s & H.TAA ' Volt
        Case 41: baseUnit = H.WAW & f & H.ALI & H.TAA ' Watt
        Case 42: baseUnit = H.HAM & f & H.mim & s & H.baa & k & H.YAA & s & H.RHO ' Ampere
        Case 60: baseUnit = H.baa & f & H.ALI & H.baa ' Bab
        Case 62: baseUnit = H.FHA & f & H.RHO & k & H.YAA & s & H.QAF ' Fariq

        ' --- KELOMPOK MUANNAS ---
        Case 10, 13, 22, 50, 51, 61
            isMuannasUnit = True
            GenderFinal = muannas
            Select Case unitType
                Case 10: baseUnit = H.sin & f & H.ALI & H.AIN & f & H.MAR ' Saah
                Case 13
                         baseUnit = H.sin & f & H.NUN & f & H.MAR ' Sanah
                         pluralUnit = H.sin & f & H.NUN & f & H.WAW & f & H.ALI & H.TAA
                Case 22: baseUnit = H.sho & f & H.FHA & s & H.haa & f & H.MAR ' Shafhah
                Case 50: baseUnit = H.dal & f & H.RHO & f & H.jim & f & H.MAR ' Darajah
                Case 61: baseUnit = H.mim & f & H.RHO & s & H.TAA & f & H.baa & f & H.MAR ' Martabah
                Case 51: baseUnit = H.FHA & k & H.YAA & s & " " & H.ALI & H.LAM & H.mim & k & IIf(gaya = Modern, "", H.ALI) & ChrW(1574) & f & H.MAR & k ' Fialmiah (%)
            End Select
    End Select

    Dim convertedText As String, TeksMa_dud As String
    Dim AkhiranDua As String
    AkhiranDua = IIf(IsMarfu, f & H.ALI & H.NUN & k, f & H.YAA & s & H.NUN & k) & " " & H.BeriHarakat(H.Core_GabungAngkaAr("2", GenderFinal, gaya, irab), irab, , , False, vHarakat)
    Dim isDualSuffix As Boolean: isDualSuffix = False

    ' --- DESIMAL ---
    Dim hasDecimal As Boolean
    hasDecimal = (InStr(strAngka, ",") > 0 Or InStr(strAngka, ".") > 0)

    If hasDecimal And Not IsOrdinal Then
        Dim tempAngka As String: tempAngka = Replace(strAngka, ",", ".")
        Dim posKoma As Integer: posKoma = InStr(tempAngka, ".")
        
        Dim mainPart As String: mainPart = Left(tempAngka, posKoma - 1)
        If mainPart = "" Then mainPart = "0"
        Dim decimalPart As String: decimalPart = Mid(tempAngka, posKoma + 1)
        
        Dim hasilBulat As String, hasilDesimal As String, hubungFashilah As String
        
        If gaya = Modern Then
            ' 1. Ambil Angka Bulat saja (Tanpa Unit)
            hasilBulat = H.BeriHarakat(H.Core_GabungAngkaAr(mainPart, GenderFinal, gaya, irab, False), irab, gaya, , False, vHarakat)
            
            ' 2. Eja Desimal Per Digit
            Dim idx As Integer, digitTeks As String
            For idx = 1 To Len(decimalPart)
                digitTeks = digitTeks & H.Kata(val(Mid(decimalPart, idx, 1)), Muzakkar) & " "
            Next idx
            hasilDesimal = H.BeriHarakat(Trim(digitTeks), irab, gaya, , False, vHarakat, True)
            hubungFashilah = " " & FASHILAH & " "
            
            ' 3. Logika Unit/Penyebut (Khusus Persen vs Unit Lain)
            Dim teksPenyebut As String
            'If unitType = 51 Then
                ' KHUSUS PERSEN: Menggunakan baseUnit langsung (Fi al-Mi'ah) tanpa perubahan jamak
                'teksPenyebut = baseUnit
            'Else
                ' UNIT LAIN: Mengikuti logika Ma'dud (Satuan/Jamak)
                'Dim vBulat As Long: vBulat = val(mainPart) Mod 100
                'If vBulat >= 3 And vBulat <= 10 Then
                    'teksPenyebut = IIf(pluralUnit <> "", pluralUnit, baseUnit & f & h.ALI & h.TAA) & kt
                'ElseIf vBulat >= 11 And vBulat <= 99 Then
                    'teksPenyebut = baseUnit & IIf(isMuannasUnit, ft, ft & h.ALI)
                'Else
                    teksPenyebut = baseUnit & ha '& kt
                'End If
            'End If

            ' GABUNGKAN: [Angka Bulat] + [Fashilah] + [Eja Digit] + [Persen/Unit]
            ArabUniversal = Trim$(hasilBulat & hubungFashilah & hasilDesimal & " " & teksPenyebut)
            
        Else
            ' --- GAYA KLASIK/SASTRA: Unit tetap setelah angka bulat ---
            hasilBulat = ArabUniversal(mainPart, unitType, IsOrdinal, irab, gaya, vHarakat)
            hasilDesimal = H.Core_GabungAngkaAr(decimalPart, Muzakkar, gaya, irab)
            hasilDesimal = H.BeriHarakat(hasilDesimal, irab, gaya, , False, vHarakat)
            hubungFashilah = " " & H.WAW & f  '& FASHILAH & " "
            Dim bagian As String: bagian = H.WAW & f & H.jim & d & H.ZAI & s & ChrW(&H621) & ha & IIf(irab = Mansub, H.ALI, "") & " "

            Dim Makam As String, TargetStyle As Integer, L As Integer: L = Len(decimalPart)
                TargetStyle = Switch(LCase(gaya) = "klasik" Or gaya = "2", 2, LCase(gaya) = "sastra" Or gaya = "3", 3, True, 1)
            Dim minAl As String: minAl = H.mim & ChrW(1616) & H.NUN & ChrW(1614) & " " & H.ALI & H.LAM
            Dim ratus As String: ratus = H.mim & ChrW(1616) & IIf(TargetStyle = 1, "", H.ALI) & ChrW(1574) & ChrW(1614) & H.MAR & ChrW(1616)
            
            Select Case L
                Case 1: Makam = " " & bagian & minAl & H.AIN & ChrW(1614) & H.syi & ChrW(1618) & H.RHO & ChrW(1614) & H.MAR & ChrW(1616)
                Case 2: Makam = " " & bagian & minAl & ratus
                Case 3: Makam = " " & bagian & minAl & H.ALI & ChrW(1614) & H.LAM & ChrW(1618) & H.FHA & ChrW(1616)
                Case Else
                    Dim idxSkala As Integer: idxSkala = (L \ 3)
                    Dim sisaNol As Integer: sisaNol = L Mod 3
                    ' Majrur (3) dipanggil untuk Skala
                    Dim namaSkala As String: namaSkala = H.MapKataHarakat(H.SatuanBesar(idxSkala), 3, TargetStyle, , False)
                    
                    Select Case sisaNol
                        Case 1: Makam = " " & H.mim & ChrW(1616) & H.NUN & " " & H.AIN & ChrW(1614) & H.syi & ChrW(1618) & H.RHO & ChrW(1616) & " " & namaSkala
                        Case 2: Makam = " " & H.mim & ChrW(1616) & H.NUN & " " & ratus & " " & namaSkala
                        Case 0: Makam = " " & minAl & namaSkala
                    End Select
            End Select
            If decimalPart = 1 Then
            ArabUniversal = Trim$(hasilBulat & " " & bagian & hasilDesimal)
            Else
            ArabUniversal = Trim$(hasilBulat & hubungFashilah & hasilDesimal) & Makam
        End If
        End If
        
        Exit Function
    End If
    
    ' --- PERSEN (UNIT 51) ANGKA BULAT ---
    If unitType = 51 And Not hasDecimal Then
        Dim teksAngkaMurni As String
        teksAngkaMurni = H.BeriHarakat(H.Core_GabungAngkaAr(strAngka, GenderFinal, gaya, irab, False), irab, gaya, , True, vHarakat)
        
        ' Gabungkan dengan unit di akhir (Fi al-Mi'ah)
        ArabUniversal = Trim$(teksAngkaMurni & " " & baseUnit)
        Exit Function
    End If

    ' --- LOGIKA ORDINAL ---
    If IsOrdinal Then
    Dim strFinalOrdinal As String
        strFinalOrdinal = H.GetArabicOrdinal(strAngka, GenderFinal, gaya, irab)
        
        ' --- BAB (60) ---
        Dim s_p_cek As Integer: s_p_cek = val(strAngka) Mod 100
        
        ' Cek: Apakah ini Bab (60), Mudzakkar, dan satuannya 1 (21, 31, dst)
        If unitType = 60 And GenderFinal = Muzakkar And s_p_cek > 20 And (s_p_cek Mod 10 = 1) Then
            Dim targetCari As String: targetCari = H.ALI & H.LAM & H.haa & H.ALI & H.dal & H.YAA
            Dim kataGanti As String: kataGanti = H.ALI & H.LAM & H.WAW & f & H.ALI & H.haa & k & H.dal & ha  ' Wahid
            
            If InStr(strFinalOrdinal, targetCari) > 0 Then
                strFinalOrdinal = Replace$(strFinalOrdinal, targetCari, kataGanti)
            End If
        End If

        convertedText = H.BeriHarakat(strFinalOrdinal, irab, gaya, , , vHarakat)
        TeksMa_dud = IIf(baseUnit <> "", H.ALI & H.LAM & s & baseUnit, "")
        ArabUniversal = Trim$(TeksMa_dud) & " " & convertedText
    Else
        ' --- ANGKA SATUAN (1) ---
        If IsSatu And unitType <> 50 Then
            ha = IIf(ha = ft And GenderFinal = Muzakkar, ft & H.ALI, ha)
            ArabUniversal = baseUnit & ha & " " & H.BeriHarakat(H.Core_GabungAngkaAr("1", GenderFinal, gaya, irab), irab, gaya, , False, vHarakat)
        
        ' --- ANGKA DUA (2) ---
        ElseIf (IsDua And unitType <> 50) Then
            ArabUniversal = IIf(isMuannasUnit And Right$(baseUnit, 1) = H.MAR, Left$(baseUnit, Len(baseUnit) - 1) & H.TAA, baseUnit) & AkhiranDua
        
        ' --- ANGKA 3 - 999 ---
        Else
            Dim flagIdhafah As Boolean
            flagIdhafah = IIf(hasDecimal, False, True)
            
            convertedText = H.Core_GabungAngkaAr(strAngka, GenderFinal, gaya, irab, flagIdhafah)
            
            If unitType = 50 Then
                TeksMa_dud = baseUnit & kt
            Else
                If V >= 3 And V <= 10 Then
                    If pluralUnit <> "" Then
                        TeksMa_dud = pluralUnit & kt
                    Else
                        If isMuannasUnit Then
                            TeksMa_dud = Left$(baseUnit, Len(baseUnit) - 1) & H.ALI & H.TAA & kt
                        Else
                            TeksMa_dud = baseUnit & f & H.ALI & H.TAA & kt
                        End If
                    End If
                ElseIf V >= 11 And V <= 99 Then
                    TeksMa_dud = baseUnit & IIf(isMuannasUnit, ft, ft & H.ALI)
                Else
                    TeksMa_dud = baseUnit & kt
                End If
            End If

            ' --- LOGIKA AKHIRAN 1 DAN 2 ---
            If Len(strAngka) > 1 And Not hasDecimal Then
                Dim tW As String: tW = H.Core_GabungAngkaAr("1", GenderFinal, gaya, irab, False)
                Dim tI As String: tI = H.Core_GabungAngkaAr("2", GenderFinal, gaya, irab, False)
                
                Dim unitSuffix As String: unitSuffix = IIf(ha = ft And GenderFinal = Muzakkar, ft & H.ALI, ha)
                convertedText = Trim$(convertedText)
            
                ' --- KASUS AKHIRAN 1 (V = 1) ---
                If V = 1 Then
                    If Right$(convertedText, Len(tW)) = tW Then convertedText = Trim$(Left$(convertedText, Len(convertedText) - Len(tW)))
            
                    Dim blokBesar As String: blokBesar = H.BeriHarakat(convertedText, irab, gaya, , True, vHarakat)
                    Dim wahidMurni As String: wahidMurni = H.BeriHarakat(tW, irab, gaya, , False, vHarakat)
            
                    If gaya = sastra Then
                        Dim OptSastra As String: OptSastra = "A"
                        Dim satuanSatu As String: satuanSatu = baseUnit & IIf(irab = 1, dt, IIf(irab = 2, ft & H.ALI, kt))
            
                        If UCase$(OptSastra) = "A" Then
                            ArabUniversal = satuanSatu & " " & blokBesar & " " & baseUnit & kt
                        Else
                            ArabUniversal = blokBesar & " " & baseUnit & kt
                        End If
                        Exit Function
            
                    ElseIf gaya = klasik Then
                        If Right$(convertedText, 1) = H.WAW Then convertedText = Trim$(Left$(convertedText, Len(convertedText) - 1))
                        ArabUniversal = H.BeriHarakat(Trim$(convertedText & " " & TeksMa_dud), irab, gaya, , True, vHarakat) & " " & H.WAW & f & baseUnit & unitSuffix
                        Exit Function
            
                    Else
                        ArabUniversal = blokBesar & " " & baseUnit & unitSuffix & " " & wahidMurni
                        Exit Function
                    End If
            
                ' --- KASUS AKHIRAN 2 (V = 2) ---
                ElseIf V = 2 Then
                    convertedText = Left$(convertedText, Len(convertedText) - Len(tI))
                    isDualSuffix = True
            
                    If gaya <> Modern Then
                        If Right$(convertedText, 1) = H.WAW Then convertedText = Trim$(Left$(convertedText, Len(convertedText) - 1))
                        If Right$(convertedText, 1) = H.ALI Then convertedText = Trim$(Left$(convertedText, Len(convertedText) - 1))
            
                        If gaya = klasik Then
                            Dim unitDua As String: unitDua = IIf(isMuannasUnit, Replace$(baseUnit, H.MAR, H.TAA), baseUnit) & IIf(IsMarfu, f & H.ALI & H.NUN & k, f & H.YAA & s & H.NUN & k)
                            ArabUniversal = H.BeriHarakat(convertedText, irab, gaya, , True, vHarakat) & " " & baseUnit & kt & " " & H.WAW & f & unitDua
                        Else
                            ArabUniversal = H.BeriHarakat(H.Core_GabungAngkaAr(strAngka, GenderFinal, gaya, irab, flagIdhafah), irab, gaya, , True, vHarakat) & " " & baseUnit & kt
                        End If
                        Exit Function
                    End If
                End If
            End If
            
            ' --- FINALISASI OUTPUT ---
            convertedText = H.BeriHarakat(convertedText, irab, gaya, , True, vHarakat)
            Dim hAkr As String
            hAkr = IIf(isDualSuffix, IIf(isMuannasUnit, Replace$(baseUnit, H.MAR, H.TAA), baseUnit) & AkhiranDua, _
                   baseUnit & IIf(ha = ft And GenderFinal = Muzakkar, ft & H.ALI, ha) & " " & H.BeriHarakat(H.Core_GabungAngkaAr("1", GenderFinal, gaya, irab), irab, gaya, , False, vHarakat))
            
            If (Right$(convertedText, 1) = H.WAW Or Right$(convertedText, 2) = H.WAW & f) And Not hasDecimal Then
                ArabUniversal = convertedText & " " & hAkr
            Else
                ArabUniversal = Trim$(convertedText & " " & TeksMa_dud)
            End If
        End If
    End If
End Function

Public Sub Clock_Start()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Sheet1")
    
    On Error Resume Next
    ws.Unprotect "password" ' Tambahkan password jika ada
    
    ' 1. Update Nilai di Sel
    ws.Range("C3").Value = Jam(Now, "u", "modern", True, True, 0)
    
    ' 2. Jadwalkan Update Berikutnya (1 Menit ke depan)
    NextTick = Now + TimeValue("00:01:00")
    Application.OnTime NextTick, "Clock_Start"
    
    ' 3. Proteksi Kembali
    ws.Protect "password"
    On Error GoTo 0
End Sub

Public Sub Clock_Stop()
    On Error Resume Next
    ' Menghentikan jadwal hanya jika NextTick sudah terisi
    If NextTick <> 0 Then
        Application.OnTime NextTick, "Clock_Start", , False
        NextTick = 0 ' Reset variabel
        MsgBox "Jam telah dihentikan.", vbInformation, "Status"
    End If
    
    ' Pastikan sheet kembali dalam keadaan terproteksi jika diinginkan
    ThisWorkbook.Worksheets("Sheet1").Protect "password"
End Sub

Public Function Jam(ByVal n As Variant, _
                    Optional ByVal irab As String = "u", _
                    Optional ByVal gaya As String = "modern", _
                    Optional ByVal vHarakat As Boolean = False, _
                    Optional ByVal ShowVisual As Boolean = False, _
                    Optional ByVal StyleVisual As ClockStyle = SolidBlock) As String

    Jam = ArabData.Clock_ToTeks(n, irab, gaya, vHarakat, ShowVisual, StyleVisual)

End Function

Public Function Tanggal_Start(ByVal d As Variant, _
                        Optional ByVal gaya As String = "modern", _
                        Optional ByVal isHijri As Boolean = False, _
                        Optional ByVal vHarakat As Boolean = False) As String

Tanggal_Start = ArabData.Tanggal(d, gaya, isHijri, vHarakat)
End Function

'=================== BAB BANTUAN & REGISTRASI =================================
Function TerbilangHelpArab() As String
    Dim NL As String: NL = Chr(10)
    TerbilangHelpArab = "=== PANDUAN TERBILANG ARAB ===" & NL & NL & _
        "[SINTAKS]" & NL & "=TERBILANG_ARAB(RefInput; [Mode]; [Gender]; [I'rab]; [Gaya]; [Parameter Tambahan]; [Harakat]; [Idhafah])" & NL & NL & _
    "[Mode] : umum, urutan, eja uang, jarak, meter, waktu, buku, lainnya." & NL & _
    "[Gender] : m (Laki-laki) | f (Perempuan)." & NL & _
    "[I'rab] : u (Marfu - Subjek/Predikat) | a (Mansub - Objek) | i (Majrur - Setelah Preposisi/Idhafah)." & NL & _
    "[Gaya] : modern (Standar) | klasik (Wawu sambung) | sastra (Kecil ke besar)." & NL & _
    "[Parameter Tambahan] : ID Negara (id, sa) jika mode 'uang', atau ID Benda (1-63)." & NL & _
    "[Harakat] : True (Dengan Harakat), False (Polos)." & NL & _
    "[Idhafah]: True (Gunakan aturan penyambungan kata benda), False (Tanpa idhafah)." & NL & NL & _
        "Catatan: " & NL & _
        "Mendukung hingga 1000 digit." & NL & _
        "Pastikan Format Cell diatur ke 'Wrap Text' agar tulisan ini rapi."
End Function

Sub RegisterArabFunctions()
On Error Resume Next
    Dim DeskripsiFungsi As String, tKategori As String, ketSkala As String

    tKategori = "HajaArabic (VBA Terbilang Arab)"
    
        ' Penentuan Skala
    On Error Resume Next
    Dim vSkala As Variant
    vSkala = ThisWorkbook.Worksheets("Sheet1").Cells(6, 3).Value
    ketSkala = IIf(LCase(CStr(vSkala)) = "panjang", "Skala Panjang", "Skala Pendek")
    On Error GoTo 0
        
    ' --- TERBILANG_ARAB ---
    Dim ArgHelp(0 To 7) As String
    ArgHelp(0) = " : Angka atau referensi sel (Mendukung angka hingga 65 digit)."
    ArgHelp(1) = " : umum urutan eja uang jarak meter waktu buku lainnya."
    ArgHelp(2) = " : m (Laki-laki) | f (Perempuan)."
    ArgHelp(3) = " : u (Marfu - Subjek/Predikat) | a (Mansub - Objek) | i (Majrur - Setelah Preposisi/Idhafah)."
    ArgHelp(4) = " : modern (Standar) | klasik (Wawu sambung) | sastra (Kecil ke besar)."
    ArgHelp(5) = " : ID Negara (id, sa) jika mode 'uang', atau ID Benda (1-63)."
    ArgHelp(6) = " : True (Dengan Harakat), False (Polos)."
    ArgHelp(7) = " : True (Gunakan aturan penyambungan kata benda), False (Tanpa idhafah)."
    
    DeskripsiFungsi = "Konverter Angka ke Bahasa Arab (Tafqit)" & vbCrLf & _
                      "----------------------------------------" & vbCrLf & _
                      "Dev: Rida Haja (Rida Rahman DH 96-02)" & vbCrLf & _
                      "GitHub: https://github.com/Rida-Haja/VbaTerbilangArab" & vbCrLf & _
                      "Menggunakan " & ketSkala
                      
    Application.MacroOptions Macro:="TERBILANG_ARAB", _
        Description:=DeskripsiFungsi, _
        Category:=tKategori, _
        ArgumentDescriptions:=ArgHelp
        
'===============================================================================
 
    ' --- JAM ---
    Dim JamHelp(0 To 4) As String
    JamHelp(0) = " : Nilai yang akan dikonversi, Support NOW(), TODAY(), atau Sel Kosong"
    JamHelp(1) = " : modern / 1 | klasik / 2 | sastra / 3."
    JamHelp(2) = " : true | false "
    JamHelp(3) = " : true | false "
    JamHelp(4) = " : SolidBlock / 1 | DotMatrix / 2 | ArabicArqam / 3."

    DeskripsiFungsi = "Mengonversi jam keTerbilangnya dalam bahasa Arab" & vbCrLf & _
                      "---------------------------------------------" & vbCrLf
                      
    Application.MacroOptions Macro:="Jam", _
        Description:=DeskripsiFungsi, _
        Category:=tKategori, _
        ArgumentDescriptions:=JamHelp
        
'===============================================================================

    ' --- TANGGAL ---
    Dim TanggalHelp(0 To 2) As String
    TanggalHelp(0) = " : Nilai yang akan dikonversi. Support TODAY() & Sel Kosong "
    TanggalHelp(1) = " : modern / 1 | klasik / 2 | sastra / 3."
    TanggalHelp(2) = " : Konversi tanggal Masehi ke Hijriyah."

    DeskripsiFungsi = "Mengonversi tanggal keterbilangnya dalam bahasa Arab" & vbCrLf & _
                      "---------------------------------------------" & vbCrLf
                      
    Application.MacroOptions Macro:="Tanggal_Start", _
        Description:=DeskripsiFungsi, _
        Category:=tKategori, _
        ArgumentDescriptions:=TanggalHelp

'=====================================================================================
    
    ' --- ARAB_CURRENCY ---
    Dim ArabCurrency(0 To 1) As String
    ArabCurrency(0) = " : Angka nominal uang yang akan dikonversi."
    ArabCurrency(1) = " : 'id' (Rupiah), 'sa' (Riyal), 'kw' (Dinar), 'ae' (Dirham), dll."

    DeskripsiFungsi = "Konverter Nominal Uang ke Teks Bahasa Arab" & vbCrLf & _
                      "---------------------------------------------" & vbCrLf & _
                      "Menampilkan satuan mata uang dan pecahan (sen/halala) " & vbCrLf & _
                      "secara otomatis sesuai negara yang dipilih."
                      
    Application.MacroOptions Macro:="ArabCurrency", _
        Description:=DeskripsiFungsi, _
        Category:=tKategori, _
        ArgumentDescriptions:=ArabCurrency
        
    If Err.Number = 0 Then
        MsgBox "Sistem HajaArabic (VBA Terbilang Arab) menggunakan " & ketSkala & " berhasil didaftarkan ke Excel!", vbInformation, "Sukses"
    Else
        MsgBox "Terjadi kesalahan saat pendaftaran.", vbCritical, "Error"
    End If
End Sub
