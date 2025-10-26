Attribute VB_Name = "ModuleReport3"
Option Explicit
Public OpenWordTakip As Boolean
Public OpenWordSay As Integer
Dim IlceSakla As String

Sub KillWordinTaskBar()
Dim ObjWordx As Object

    OpenWordSay = 0
    On Error Resume Next
    Set ObjWordx = GetObject(, "Word.Application")
    Set ObjWordx = GetObject(, "Word.Application")
    Set ObjWordx = GetObject(, "Word.Application")
    Set ObjWordx = GetObject(, "Word.Application")
    Set ObjWordx = GetObject(, "Word.Application")
    Set ObjWordx = GetObject(, "Word.Application")
    Set ObjWordx = GetObject(, "Word.Application")
    Set ObjWordx = GetObject(, "Word.Application")
    Set ObjWordx = GetObject(, "Word.Application")
    Set ObjWordx = GetObject(, "Word.Application")
    If ObjWordx Is Nothing Then
        Set ObjWordx = CreateObject("Word.Application")
        ObjWordx.Visible = False
        GoTo Atla
    End If

    ObjWordx.Visible = True
    ObjWordx.Quit SaveChanges:=True
    Set ObjWordx = Nothing
    GoTo Son

    
Atla:
OpenWordSay = 1

Son:

End Sub

Sub OpenWordControl()

Dim ObjWordx As Object
Dim objDocx As Object

'MsgBox "OpenWordControl prosedürü başlıyor."

    On Error GoTo NoOpenDoc
    Set ObjWordx = GetObject(, "Word.Application")
    Set ObjWordx = GetObject(, "Word.Application")
    Set ObjWordx = GetObject(, "Word.Application")
    Set ObjWordx = GetObject(, "Word.Application")
    Set ObjWordx = GetObject(, "Word.Application")
    OpenWordTakip = True
    GoTo NoOpenDocAtla
NoOpenDoc:
    OpenWordTakip = False
NoOpenDocAtla:
    If OpenWordTakip = True Then
        'MsgBox objWordx.ActiveDocument.Name
        If ObjWordx.ActiveDocument.name <> "" Then
            ObjWordx.Quit SaveChanges:=True
            'MsgBox "Dosya OpenWordControl methodu ile kapatıldı."
        End If
    Else
        'MsgBox "Açık word dokümanı yok."
    End If

Son:
Set ObjWordx = Nothing

End Sub

Sub Rapor3_1Tutanak()

Dim DestOperasyon As String, SourceRapor3_1Normal As String, AutoPath As String, ReNameRapor3_1 As String
Dim fso As Object, objWord As Object, objDoc As Object
Dim OpenKontrolName As String, OpenControl As String, DestOpUserFolder As String, DestOpUserFolderName As String
Dim ContSay As Integer, KontrolFile As String, x As Long, y As Long, i As Long, j As Long
Dim IlkSiraBul As Range, SonSiraBul As Range, IlkSira As Long, SonSira As Long
Dim GelenTema As String, SourceDL As String, ReNameDL As String
Dim TxtFileRapor3_1 As String, TotalSayfaRapor3_1 As String, TxtFileDokum As String, TotalSayfaDokum As String
Dim DokumFileTxt As Object, DokumSayfaGonder As String
Dim SourceRaporNormal As String, SourceTutanak2Normal As String, SourceUstYaziNormal As String
Dim Bolum1 As String, Bolum2 As String, Bolum3 As String, Bolum4 As String, Bolum5 As String, Bolum6 As String
Dim Ek1 As String, Ek2 As String, Ek3 As String, Ek4 As String, Ek5 As String, Ek6 As String, Birlestir As String

Dim ReNameRaporNormal As String, SiraNoIlkSatir As Long, RaporIlgi As String
Dim AltRaporNoIlk As Long, AltRaporNoSon As Long, AdetTopla As Long
Dim TekCogulTipA As String, k As Integer
Dim MyRange As Object, TxtFileRapor As String, TotalSayfaRapor As String
Dim ReNameTutanak2Normal As String, GidenTema As String
Dim TxtFileTutanak2 As String, TotalSayfaTutanak2 As String
Dim ReNameUstYaziNormal As String, PaketTipi As String
Dim TxtFileUstYazi As String, TotalSayfaUstYazi As String
Dim Birimx As String, UserName As String, SourceRapor3_1Farkli As String
Dim Kolluk As String


'TUTANAK1 için prosedürü başlat
'If ActiveCell.Column = 6 Then
'    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    'Application.EnableEvents = False
    
    'Pathfinder...
    AutoPath = ThisWorkbook.Path
    DestOperasyon = AutoPath & "\System Files\Operation\"
    'TUTANAK1 TANIMLARI
    SourceRapor3_1Normal = AutoPath & "\System Files\System Templates\Report 3 Statements\Report 3.1 – Type A.docm"
    
    'Birden fazla kullanıcının operasyon klasöründe çakışmasını önlemek için operasyon klasörünün içinde yeni klasör aç.
    DestOpUserFolderName = "Operation-" & Left(ThisWorkbook.name, InStr(ThisWorkbook.name, ".") - 1) '& " " & Format(Now(), "ddmmyyyyhhmmss")
    DestOpUserFolder = DestOperasyon & DestOpUserFolderName & "\"
    
    'Check the "System Files" folder name.
    If Not Dir(AutoPath & "\System Files\", vbDirectory) <> vbNullString Then
        MsgBox "Cannot access the directory " & AutoPath & "\System Files\" & ". The folder named 'System Files' may have been renamed or deleted.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    
    'Check the "Operation" folder name.
    If Not Dir(DestOperasyon, vbDirectory) <> vbNullString Then
        MsgBox "Cannot access the directory " & DestOperasyon & ". The folder named 'Operation' may have been renamed or deleted.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    
    'Check folder names.
    If Not Dir(SourceRapor3_1Normal, vbDirectory) <> vbNullString Then
        MsgBox "Cannot access the directory " & SourceRapor3_1Normal & ". The names of folders and/or files in this directory may have been changed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If

    'Operation klasörü içinde kullanıcı modülü klasörü yoksa oluştur.
    If Not Dir(DestOpUserFolder, vbDirectory) <> vbNullString Then
        MkDir DestOpUserFolder
    End If
    
    'On Error Resume Next
    ReNameRapor3_1 = Cells(ActiveCell.Row, 5).Value & "-" & Cells(6, 6).Value
    'ReNameDL = Cells(ActiveCell.Row, 5).Value & "-" & "Dispatch List"
'________________________________________

    
    'Close the all Word application
    Call ModuleReport3.OpenWordControl
    
    'Operation klasöründeki docm uzantılı word dosyalarından açık olanları kapat ve temizle.
    OpenKontrolName = Dir(DestOpUserFolder & "*.docm")
    Do While OpenKontrolName <> ""
        OpenControl = IsFileOpen(DestOpUserFolder & OpenKontrolName)
        If OpenControl = True Then 'Açıksa
            On Error Resume Next
            Set objWord = GetObject(, "Word.Application")
            Set objWord = GetObject(, "Word.Application")
            Set objWord = GetObject(, "Word.Application")
            Set objWord = GetObject(, "Word.Application")
            Set objWord = GetObject(, "Word.Application")
            objWord.Quit SaveChanges:=True
            'MsgBox "Dosya OpenKontrol methodu ile kapatıldı."

        End If
        OpenKontrolName = Dir()
    Loop

'    Call ModuleReport3.OpenWordControl
     
    Set objWord = Nothing
    Set objDoc = Nothing

'________________________________________

    On Error Resume Next
'    Klasörün içindeki tüm dosyaları sil (txt, docm vb.)
    ContSay = 0
    KontrolFile = Dir(DestOpUserFolder & "*.???")
    Do While KontrolFile <> ""
        ContSay = ContSay + 1
        KontrolFile = Dir()
    Loop
    If ContSay > 0 Then
        Kill DestOpUserFolder & "*.???"
    End If
    
    'Dosyayı şablondan operasyon klasörüne kopyala ve adını değiştir.
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.CopyFile (SourceRapor3_1Normal), DestOpUserFolder & ReNameRapor3_1 & ".docm", True
    
'________________________________________

    'Oluşturulacak dosyayı aç
    On Error Resume Next
    Set objWord = GetObject(, "Word.Application")
    Set objWord = GetObject(, "Word.Application")
    Set objWord = GetObject(, "Word.Application")
    Set objWord = GetObject(, "Word.Application")
    Set objWord = GetObject(, "Word.Application")
    If objWord Is Nothing Then
        'MsgBox "Dosya oluşturmada CreateObject methodu kullanılacak."
        Set objWord = CreateObject("Word.Application")
        objWord.Visible = False
    End If
    objWord.Documents.Open FileName:=DestOpUserFolder & ReNameRapor3_1 & ".docm"
    objWord.Visible = True
    objWord.Activate 'Ekrana getirir.
    'objDoc.Activate 'Ekrana getirmez.
    objWord.Application.WindowState = wdWindowStateMaximize
    Set objDoc = GetObject(DestOpUserFolder & ReNameRapor3_1 & ".docm")
'________________________________________

    Set IlkSiraBul = Range("FG7:FG100000").Find(What:=Cells(ActiveCell.Row, 5).Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    Set SonSiraBul = Range("FH7:FH100000").Find(What:=Cells(ActiveCell.Row, 5).Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    If Not IlkSiraBul Is Nothing Then
        IlkSira = IlkSiraBul.Row
    Else
        MsgBox "The serial number could not be found, so the operation cannot be completed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    If Not SonSiraBul Is Nothing Then
        SonSira = SonSiraBul.Row
    Else
        MsgBox "The serial number could not be found, so the operation cannot be completed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    
    'Birim
    Birimx = UCase(Replace(Replace(Worksheets(2).Cells(6, 99).Value, "i", "I"), "ı", "I")) & " UNIT"
    objDoc.Tables(1).Cell(Row:=1, Column:=1).Range.Text = Birimx
    
    'Dosyada içerikleri değiştir.
    'Kişinin açık kimliği var/yok
    If Cells(ActiveCell.Row, 126).Value <> "" And Cells(ActiveCell.Row, 126).Value > 0 Then
        'var
    Else
        'yok
        Ek1 = ""
        Set MyRange = objDoc.Tables(1).Cell(Row:=4, Column:=1).Range
        With MyRange.Find
            .Text = "identified as xxxxx xxxxx xx xxxxx, "
            .Replacement.Text = Ek1
            .Execute Replace:=wdReplaceAll
        End With
    End If
    'Ad Soyad
    Ek1 = Cells(ActiveCell.Row, 109).Value
    Set MyRange = objDoc.Tables(1).Cell(Row:=4, Column:=1).Range
    With MyRange.Find
        .Text = "<fullName>"
        .Replacement.Text = Ek1
        .Execute Replace:=wdReplaceAll
    End With
    MyRange.HighlightColorIndex = wdNoHighlight
    'Getirilme amacı
    Ek1 = Cells(ActiveCell.Row, 106).Value
    Set MyRange = objDoc.Tables(1).Cell(Row:=4, Column:=1).Range
    With MyRange.Find
        .Text = "<purpose>"
        .Replacement.Text = Ek1
        .Execute Replace:=wdReplaceAll
    End With
    MyRange.HighlightColorIndex = wdNoHighlight
    'TipA(un)(ların)
    'Tekil-çoğul tipA
    AdetTopla = Application.Sum(Range(Cells(IlkSira, 136), Cells(SonSira, 136)))
    If AdetTopla = 1 Then
        TekCogulTipA = "Type A item"
        Ek2 = "has"
        Ek3 = "Type A Item"
    ElseIf AdetTopla > 1 Then
        TekCogulTipA = "Type A items"
        Ek2 = "have"
        Ek3 = "Type A Items"
    End If
    Set MyRange = objDoc.Tables(1).Cell(Row:=4, Column:=1).Range
    With MyRange.Find
        .Text = "<typeA>"
        .Replacement.Text = TekCogulTipA
        .Execute Replace:=wdReplaceAll
    End With
    MyRange.HighlightColorIndex = wdNoHighlight
    Set MyRange = objDoc.Tables(1).Cell(Row:=4, Column:=1).Range
    With MyRange.Find
        .Text = "<have_has>"
        .Replacement.Text = Ek2
        .Execute Replace:=wdReplaceAll
    End With
    MyRange.HighlightColorIndex = wdNoHighlight
    
    'Kolluk
    If Cells(ActiveCell.Row, 102).Value = "Provincial Directorate B" Then
        Kolluk = Cells(ActiveCell.Row, 91).Value & " Provincial Governorship Provincial Directorate B " & Cells(ActiveCell.Row, 103).Value
    ElseIf Cells(ActiveCell.Row, 102).Value = "District Directorate B" Then
        Kolluk = Cells(ActiveCell.Row, 92).Value & " District Governorship District Directorate B " & Cells(ActiveCell.Row, 103).Value
    ElseIf Cells(ActiveCell.Row, 102).Value = "Provincial Directorate C" Then
        Kolluk = Cells(ActiveCell.Row, 91).Value & " Provincial Governorship Provincial Directorate C " & Cells(ActiveCell.Row, 103).Value
    ElseIf Cells(ActiveCell.Row, 102).Value = "District Directorate C" Then
        Kolluk = Cells(ActiveCell.Row, 92).Value & " District Governorship District Directorate C " & Cells(ActiveCell.Row, 103).Value
    ElseIf Cells(ActiveCell.Row, 102).Value = "Provincial Directorate D" Then
        Kolluk = Cells(ActiveCell.Row, 91).Value & " Provincial Governorship Provincial Directorate D " & Cells(ActiveCell.Row, 103).Value
    ElseIf Cells(ActiveCell.Row, 102).Value = "District Directorate D" Then
        Kolluk = Cells(ActiveCell.Row, 92).Value & " District Governorship District Directorate D " & Cells(ActiveCell.Row, 103).Value
    ElseIf Cells(ActiveCell.Row, 102).Value = "Provincial Directorate E" Then
        Kolluk = Cells(ActiveCell.Row, 91).Value & " Provincial Governorship Provincial Directorate E " & Cells(ActiveCell.Row, 103).Value
    ElseIf Cells(ActiveCell.Row, 102).Value = "District Directorate E" Then
        Kolluk = Cells(ActiveCell.Row, 92).Value & " District Governorship District Directorate E " & Cells(ActiveCell.Row, 103).Value
    ElseIf InStr(Cells(ActiveCell.Row, 102).Value, "General Directorate") <> 0 Or InStr(Cells(ActiveCell.Row, 102).Value, "Regional Directorate") <> 0 Then
        If Cells(ActiveCell.Row, 103).Value <> "" Then
            Kolluk = WorksheetFunction.Proper(ThisWorkbook.Worksheets(2).Cells(6, 111).Value) & " " & Cells(ActiveCell.Row, 102).Value & " " & Cells(ActiveCell.Row, 103).Value
        Else
            Kolluk = WorksheetFunction.Proper(ThisWorkbook.Worksheets(2).Cells(6, 111).Value) & " " & Cells(ActiveCell.Row, 102).Value
        End If
    Else
        If Cells(ActiveCell.Row, 103).Value <> "" Then
            If InStr(Cells(ActiveCell.Row, 102).Value, "X.X. ") > 0 Then
                Kolluk = Mid(Cells(ActiveCell.Row, 102).Value, 6, Len(Cells(ActiveCell.Row, 102).Value)) & " " & Cells(ActiveCell.Row, 103).Value
            Else
                Kolluk = Cells(ActiveCell.Row, 102).Value & " " & Cells(ActiveCell.Row, 103).Value
            End If
        Else
            If InStr(Cells(ActiveCell.Row, 102).Value, "X.X. ") > 0 Then
                Kolluk = Mid(Cells(ActiveCell.Row, 102).Value, 6, Len(Cells(ActiveCell.Row, 102).Value))
            Else
                Kolluk = Cells(ActiveCell.Row, 102).Value
            End If
        End If
    End If
        
    'Recipient
    Set MyRange = objDoc.Tables(2).Cell(Row:=1, Column:=1).Range
    With MyRange.Find
        .Text = "<recipient>"
        .Replacement.Text = Kolluk
        .Execute Replace:=wdReplaceAll
    End With
    MyRange.HighlightColorIndex = wdNoHighlight
    'Tutanak tarihi
    Ek1 = Cells(ActiveCell.Row, 95).Value
    Set MyRange = objDoc.Tables(2).Cell(Row:=1, Column:=1).Range
    With MyRange.Find
        .Text = "<reportDate>"
        .Replacement.Text = Ek1
        .Execute Replace:=wdReplaceAll
    End With
    MyRange.HighlightColorIndex = wdNoHighlight

    'imzalar
    objDoc.Tables(3).Cell(Row:=4, Column:=1).Range.Text = Cells(ActiveCell.Row, 184).Value 'Ad Soyad1
    objDoc.Tables(3).Cell(Row:=5, Column:=1).Range.Text = Cells(ActiveCell.Row, 185).Value 'Unvan1
    objDoc.Tables(3).Cell(Row:=4, Column:=2).Range.Text = Cells(ActiveCell.Row, 187).Value 'Ad Soyad2
    objDoc.Tables(3).Cell(Row:=5, Column:=2).Range.Text = Cells(ActiveCell.Row, 188).Value 'Unvan2
    objDoc.Tables(3).Cell(Row:=4, Column:=3).Range.Text = Cells(ActiveCell.Row, 190).Value 'Ad Soyad3
    objDoc.Tables(3).Cell(Row:=5, Column:=3).Range.Text = Cells(ActiveCell.Row, 191).Value 'Unvan3
    
    'Tablo başlığı
    objDoc.Tables(4).Cell(Row:=1, Column:=1).Range.Text = Ek3 & " Evaluated as Invalid:"
    
    'Tabloya satır ekle
    x = 0
    If SonSira - IlkSira > 0 Then
        With objDoc.Tables(4)
            For i = IlkSira To SonSira - 1
                .Rows.Add BeforeRow:=objDoc.Tables(4).Rows(3)
                x = x + 1
            Next i
        End With
    End If
    For i = 3 To SonSira - IlkSira + 3
        objDoc.Tables(4).Cell(Row:=i, Column:=1).Range.Text = i - 2 'Tablo sıra no
        objDoc.Tables(4).Cell(Row:=i, Column:=2).Range.Text = Cells(IlkSira + i - 3, 130).Value 'Öğe türü
        objDoc.Tables(4).Cell(Row:=i, Column:=3).Range.Text = Cells(IlkSira + i - 3, 133).Value 'Öğe değeri
        objDoc.Tables(4).Cell(Row:=i, Column:=4).Range.Text = Cells(IlkSira + i - 3, 136).Value 'Adet
        objDoc.Tables(4).Cell(Row:=i, Column:=5).Range.Text = Cells(IlkSira + i - 3, 139).Value 'Öğe ID No
    Next i
    
    'Kişi  bilgileri
    'Tablo başlığı
    objDoc.Tables(5).Cell(Row:=1, Column:=1).Range.Text = "The Person Who Delivered the " & Ek3 & ":"
    objDoc.Tables(5).Cell(Row:=2, Column:=3).Range.Text = Cells(ActiveCell.Row, 109).Value 'Ad Soyad
    objDoc.Tables(5).Cell(Row:=3, Column:=3).Range.Text = Cells(ActiveCell.Row, 110).Value 'TCK No
    If Cells(ActiveCell.Row, 117).Value <> "" And Cells(ActiveCell.Row, 118).Value <> "" Then 'Kimlik Türü ve No
        objDoc.Tables(5).Cell(Row:=4, Column:=3).Range.Text = Cells(ActiveCell.Row, 117).Value & " - " & Cells(ActiveCell.Row, 118).Value
    ElseIf Cells(ActiveCell.Row, 117).Value <> "" And Cells(ActiveCell.Row, 118).Value = "" Then
        objDoc.Tables(5).Cell(Row:=4, Column:=3).Range.Text = Cells(ActiveCell.Row, 117).Value
    ElseIf Cells(ActiveCell.Row, 117).Value = "" And Cells(ActiveCell.Row, 118).Value <> "" Then
        objDoc.Tables(5).Cell(Row:=4, Column:=3).Range.Text = Cells(ActiveCell.Row, 118).Value
    End If
    objDoc.Tables(5).Cell(Row:=5, Column:=3).Range.Text = Cells(ActiveCell.Row, 111).Value 'Baba Adı
    If Cells(ActiveCell.Row, 112).Value <> "" And Cells(ActiveCell.Row, 113).Value <> "" Then 'Doğum Yeri ve Tarihi
        objDoc.Tables(5).Cell(Row:=6, Column:=3).Range.Text = Cells(ActiveCell.Row, 112).Value & " - " & Cells(ActiveCell.Row, 113).Value
    ElseIf Cells(ActiveCell.Row, 112).Value <> "" And Cells(ActiveCell.Row, 113).Value = "" Then
        objDoc.Tables(5).Cell(Row:=6, Column:=3).Range.Text = Cells(ActiveCell.Row, 112).Value
    ElseIf Cells(ActiveCell.Row, 112).Value = "" And Cells(ActiveCell.Row, 113).Value <> "" Then
        objDoc.Tables(5).Cell(Row:=6, Column:=3).Range.Text = Cells(ActiveCell.Row, 113).Value
    End If
    objDoc.Tables(5).Cell(Row:=7, Column:=3).Range.Text = Cells(ActiveCell.Row, 119).Value 'Nüfusa Kayıtlı Olduğu Yer
    objDoc.Tables(5).Cell(Row:=8, Column:=3).Range.Text = Cells(ActiveCell.Row, 120).Value 'Cilt No, Aile Sıra No, Sıra No
    If Cells(ActiveCell.Row, 123).Value <> "" And Cells(ActiveCell.Row, 116).Value <> "" Then 'Adres/Telefon Numarası
        objDoc.Tables(5).Cell(Row:=9, Column:=3).Range.Text = Cells(ActiveCell.Row, 123).Value & " - " & Cells(ActiveCell.Row, 116).Value
    ElseIf Cells(ActiveCell.Row, 123).Value <> "" And Cells(ActiveCell.Row, 116).Value = "" Then
        objDoc.Tables(5).Cell(Row:=9, Column:=3).Range.Text = Cells(ActiveCell.Row, 123).Value
    ElseIf Cells(ActiveCell.Row, 123).Value = "" And Cells(ActiveCell.Row, 116).Value <> "" Then
        objDoc.Tables(5).Cell(Row:=9, Column:=3).Range.Text = Cells(ActiveCell.Row, 116).Value
    End If
    'Getiren Kişi/Organization B Mensubu İmza alanı
    If Cells(ActiveCell.Row, 124).Value = "Yes" Then
        objDoc.Tables(6).Cell(Row:=3, Column:=1).Range.Text = "Signature"
        objDoc.Tables(6).Cell(Row:=4, Column:=1).Range.Text = "Full Name"
        objDoc.Tables(6).Cell(Row:=5, Column:=1).Range.Text = "The Person Who Delivered the " & Ek3
        objDoc.Tables(6).Cell(Row:=3, Column:=2).Range.Text = "Signature"
        objDoc.Tables(6).Cell(Row:=4, Column:=2).Range.Text = "Full Name"
        objDoc.Tables(6).Cell(Row:=5, Column:=2).Range.Text = "Officer Who Confiscated the " & Ek3
    ElseIf Cells(ActiveCell.Row, 124).Value = "No" Then
        objDoc.Tables(6).Cell(Row:=3, Column:=2).Range.Text = "Signature"
        objDoc.Tables(6).Cell(Row:=4, Column:=2).Range.Text = "Full Name"
        objDoc.Tables(6).Cell(Row:=5, Column:=2).Range.Text = "The Person Who Delivered the " & Ek3
    End If
    'Ek kimlik fotokopisi
    If Cells(ActiveCell.Row, 126).Value <> "" Then
        objDoc.Tables(7).Cell(Row:=1, Column:=1).Range.Text = "Attachment"
        objDoc.Tables(7).Cell(Row:=1, Column:=2).Range.Text = ":"
        If Cells(ActiveCell.Row, 126) > 1 Then
            objDoc.Tables(7).Cell(Row:=1, Column:=3).Range.Text = "ID Photocopy (" & Cells(ActiveCell.Row, 126) & " pages)"
        Else
            objDoc.Tables(7).Cell(Row:=1, Column:=3).Range.Text = "ID Photocopy (" & Cells(ActiveCell.Row, 126) & " page)"
        End If
    ElseIf Cells(ActiveCell.Row, 126).Value = "" And Cells(ActiveCell.Row, 127).Value = "No" Then
        objDoc.Tables(7).Cell(Row:=1, Column:=1).Range.Text = ""
        objDoc.Tables(7).Cell(Row:=1, Column:=2).Range.Text = ""
        objDoc.Tables(7).Cell(Row:=1, Column:=3).Range.Text = ""
    Else '126 boş ve 127 var ise notu eklemiş olacak.
        '
    End If

    'Olay bilgileri
    objDoc.CheckBox3.Caption = "Condition4"
    objDoc.CheckBox4.Caption = "Condition5"
    objDoc.CheckBox5.Caption = "Condition6"


    'Olay bilgilerini işaretle
    If Mid(Cells(IlkSira, 128).Value, 1, 2) = "10" Then
        objDoc.CheckBox1.Value = False
    ElseIf Mid(Cells(IlkSira, 128).Value, 1, 2) = "11" Then
        objDoc.CheckBox1.Value = True
    End If
    If Mid(Cells(IlkSira, 128).Value, 4, 2) = "20" Then
        objDoc.CheckBox2.Value = False
    ElseIf Mid(Cells(IlkSira, 128).Value, 4, 2) = "21" Then
        objDoc.CheckBox2.Value = True
    End If
    If Mid(Cells(IlkSira, 128).Value, 7, 2) = "30" Then
        objDoc.CheckBox3.Value = False
    ElseIf Mid(Cells(IlkSira, 128).Value, 7, 2) = "31" Then
        objDoc.CheckBox3.Value = True
    End If
    If Mid(Cells(IlkSira, 128).Value, 10, 2) = "40" Then
        objDoc.CheckBox4.Value = False
    ElseIf Mid(Cells(IlkSira, 128).Value, 10, 2) = "41" Then
        objDoc.CheckBox4.Value = True
    End If
    If Mid(Cells(IlkSira, 128).Value, 13, 2) = "50" Then
        objDoc.CheckBox5.Value = False
    ElseIf Mid(Cells(IlkSira, 128).Value, 13, 2) = "51" Then
        objDoc.CheckBox5.Value = True
    End If
    If Mid(Cells(IlkSira, 128).Value, 16, 2) = "60" Then
        objDoc.CheckBox6.Value = False
    ElseIf Mid(Cells(IlkSira, 128).Value, 16, 2) = "61" Then
        objDoc.CheckBox6.Value = True
    End If
    
    'Türkçe karakterleri düzelt
    objDoc.CheckBox1.Enabled = False
    objDoc.CheckBox1.Enabled = True
    objDoc.CheckBox2.Enabled = False
    objDoc.CheckBox2.Enabled = True
    objDoc.CheckBox3.Enabled = False
    objDoc.CheckBox3.Enabled = True
    objDoc.CheckBox4.Enabled = False
    objDoc.CheckBox4.Enabled = True
    objDoc.CheckBox5.Enabled = False
    objDoc.CheckBox5.Enabled = True
    objDoc.CheckBox6.Enabled = False
    objDoc.CheckBox6.Enabled = True

    'Alt bilgi ekle
    objDoc.Sections(1).Footers(wdHeaderFooterPrimary).Range.Tables(1).Cell(Row:=1, Column:=1).Range.Text = ReNameRapor3_1
        
    'Save üstünde iken save bağlı kodlar çalışmıyor; designmod açık kalıyor. Bu sebeple checkbox silme komutu
    'save komutunun altında konumlandırıldı.
    '(Güncelleme 25.11.2018: Bu sorun word'taki kodların yeniden çalıştırılması ile çözüldü.
    If Cells(ActiveCell.Row, 124).Value = "No" Then 'Organization B mensubu yoksa
        'On Error GoTo Hata
        On Error Resume Next
        objDoc.Fields.Item(1).Delete
        objDoc.Fields.Item(5).Delete
    Else
        objDoc.Tables(7).Rows.Add 'BeforeRow:=objDoc.Tables(7).Rows(0)
    End If
    
    objWord.Run "Register_Event_Handler"
    
    'Sayfa sayısı kaydet komutuna bağlandı.
    objDoc.Save
    
    'Sayfayı text dosyasından çek
    TxtFileRapor3_1 = DestOpUserFolder & "Report 3 Page Count.txt"
    #If UseADO Then
        Const adReadLine = -2&
        With CreateObject("ADODB.Stream")
            .Open
            .LoadFromFile TxtFileRapor3_1
            Do Until .EOS
                TotalSayfaRapor3_1 = .ReadText(adReadLine)
            Loop
            .Close
        End With
    #Else 'UseFSO
        Const Rapor3_1TristateTrue = -1&
        With CreateObject("Scripting.FileSystemObject")
            With .OpenTextFile(TxtFileRapor3_1, Format:=Rapor3_1TristateTrue)
                Do Until .AtEndOfStream
                    TotalSayfaRapor3_1 = .ReadLine
                Loop
                .Close
            End With
        End With
    #End If
    'MsgBox "Tutanak1: " & TotalSayfaRapor3_1

    If TumDoc = True Then
        objWord.Activate
'        For i = 1 To SayPrt
'            objDoc.PrintOut
'        Next i
        objDoc.PrintOut Background:=False, Copies:=SayPrt
        'objWord.Documents.Save
        objDoc.Close SaveChanges:=False
        objWord.Visible = False
    End If

    'Tutanak1 ve Döküm sayfa sayısı
    Cells(ActiveCell.Row, 169).Value = TotalSayfaRapor3_1
    
    'Nesneleri temizle
    Set fso = Nothing
    Set objWord = Nothing
    Set objDoc = Nothing
    GoTo Son
    
Hata:
MsgBox "An error occurred while retrieving the number of pages of the statement. Please manually enter the number of statement pages in the attachment section when creating the cover letter.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
'If Err.Number <> 0 Then
'MsgBox "Error # " & Str(Err.Number)
'End If

Son:

'Nesneleri temizle
Set fso = Nothing
Set objWord = Nothing
Set objDoc = Nothing


'Worksheets(5).Activate
Application.ScreenUpdating = True
Application.DisplayAlerts = True
'Application.EnableEvents = True

'End If

End Sub

Sub Rapor3_1Rapor()

Dim DestOperasyon As String, SourceTutanak1Normal As String, AutoPath As String, ReNameTutanak1 As String
Dim fso As Object, objWord As Object, objDoc As Object
Dim OpenKontrolName As String, OpenControl As String, DestOpUserFolder As String, DestOpUserFolderName As String
Dim ContSay As Integer, KontrolFile As String, x As Long, y As Long, i As Long, j As Long
Dim IlkSiraBul As Range, SonSiraBul As Range, IlkSira As Long, SonSira As Long
Dim GelenTema As String, SourceDL As String, ReNameDL As String
Dim TxtFileTutanak1 As String, TotalSayfaTutanak1 As String, TxtFileDokum As String, TotalSayfaDokum As String
Dim DokumFileTxt As Object, DokumSayfaGonder As String
Dim SourceRaporNormal As String, SourceTutanak2Normal As String, SourceUstYaziNormal As String
Dim Bolum1 As String, Bolum2 As String, Bolum3 As String, Bolum4 As String, Bolum5 As String, Bolum6 As String
Dim Ek1 As String, Ek2 As String, Ek3 As String, Ek4 As String, Ek5 As String, Ek6 As String, Birlestir As String

Dim ReNameRaporNormal As String, SiraNoIlkSatir As Long, RaporIlgi As String
Dim AltRaporNoIlk As Long, AltRaporNoSon As Long, AdetTopla As Long
Dim TekCogulTipA As String, k As Integer
Dim MyRange As Object, TxtFileRapor As String, TotalSayfaRapor As String
Dim ReNameTutanak2Normal As String, GidenTema As String
Dim TxtFileTutanak2 As String, TotalSayfaTutanak2 As String
Dim ReNameUstYaziNormal As String
Dim TxtFileUstYazi As String, TotalSayfaUstYazi As String
Dim Birimx As String, UserName As String, NotEkle As String
Dim Explorer As Integer, b As Long, RaporTipi As String, DestNotlar As String, TxtFileNot As String
Dim TextLine As String, StrTeknik_ANotu As String

'RAPOR için prosedürü başlat
'If ActiveCell.Column = 7 Then
'    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    'Application.EnableEvents = False

    'Dokümanların oluşturulmasına soldan sağa doğru bir kontrol mekanizması kur.
    j = ActiveCell.Row
    For i = ActiveCell.Row To 7 Step -1
        If Cells(i, 5).Value <> "" Then
            GoTo SiraNoBulundu
        Else
            j = j - 1
        End If
    Next i
SiraNoBulundu:
    If j < 7 Then
        MsgBox "The serial number could not be found, so the operation cannot be completed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    Set IlkSiraBul = Range("FG7:FG100000").Find(What:=Cells(j, 5).Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    Set SonSiraBul = Range("FH7:FH100000").Find(What:=Cells(j, 5).Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    If Not IlkSiraBul Is Nothing Then
        IlkSira = IlkSiraBul.Row
    Else
        MsgBox "The serial number could not be found, so the operation cannot be completed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    If Not SonSiraBul Is Nothing Then
        SonSira = SonSiraBul.Row
    Else
        MsgBox "The serial number could not be found, so the operation cannot be completed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    
    If Cells(IlkSira, 6).Value = "x" Then
        MsgBox "The Report 3 statement data contains missing or incorrect information, so the operation cannot be started.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If

    
    'Pathfinder...
    AutoPath = ThisWorkbook.Path
    DestOperasyon = AutoPath & "\System Files\Operation\"
    DestNotlar = AutoPath & "\System Files\System Templates\Item Notes\"
    
    'Birden fazla kullanıcının operasyon klasöründe çakışmasını önlemek için operasyon klasörünün içinde yeni klasör aç.
    DestOpUserFolderName = "Operation-" & Left(ThisWorkbook.name, InStr(ThisWorkbook.name, ".") - 1) '& " " & Format(Now(), "ddmmyyyyhhmmss")
    DestOpUserFolder = DestOperasyon & DestOpUserFolderName & "\"
    
    'Check the "System Files" folder name.
    If Not Dir(AutoPath & "\System Files\", vbDirectory) <> vbNullString Then
        MsgBox "Cannot access the directory " & AutoPath & "\System Files\" & ". The folder named 'System Files' may have been renamed or deleted.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    
    'Check the "Operation" folder name.
    If Not Dir(DestOperasyon, vbDirectory) <> vbNullString Then
        MsgBox "Cannot access the directory " & DestOperasyon & ". The folder named 'Operation' may have been renamed or deleted.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If

    
    'Operation klasörü içinde kullanıcı modülü klasörü yoksa oluştur.
    If Not Dir(DestOpUserFolder, vbDirectory) <> vbNullString Then
        MkDir DestOpUserFolder
    End If

'____________________________________
If TumDoc = True Then
    'Worksheets(4).Activate
    b = ActiveCell.Row

    'İlk ve son sıraları bul (For Explorer)
    'On Error Resume Next
    SiraNoIlkSatir = ActiveCell.Row
    If Cells(ActiveCell.Row, 5).Value = "" Then
        For i = ActiveCell.Row To 7 Step -1
            If Cells(i, 5).Value <> "" Then
                ReNameRaporNormal = Cells(i, 5).Value & "-" & Cells(6, 7).Value & " " & Cells(ActiveCell.Row, 13).Value
                SiraNoIlkSatir = i
                GoTo RaporNoDonguSonExplorer
            End If
        Next i
    Else
        ReNameRaporNormal = Cells(ActiveCell.Row, 5).Value & "-" & Cells(6, 7).Value & " " & Cells(ActiveCell.Row, 13).Value
    End If
RaporNoDonguSonExplorer:
   
    'Rapor/Alt Rapor aralıklarının tespiti.
    Set IlkSiraBul = Range("FG7:FG100000").Find(What:=Cells(SiraNoIlkSatir, 5).Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    Set SonSiraBul = Range("FH7:FH100000").Find(What:=Cells(SiraNoIlkSatir, 5).Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    If Not IlkSiraBul Is Nothing Then
        IlkSira = IlkSiraBul.Row
    Else
        MsgBox "The serial number could not be found, so the operation cannot be completed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    If Not SonSiraBul Is Nothing Then
        SonSira = SonSiraBul.Row
    Else
        MsgBox "The serial number could not be found, so the operation cannot be completed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
End If
'____________________________________

If TumDoc = False Then
    IlkSira = 1
    SonSira = 1
End If

For Explorer = IlkSira To SonSira
    If TumDoc = True Then
        If Cells(Explorer, 13) <> "" Then
           Cells(Explorer, 13).Select
        Else
            GoTo ExplorerBos
        End If
    End If
    'Dosyayı isimlendir
    'On Error Resume Next
    SiraNoIlkSatir = ActiveCell.Row
    If Cells(ActiveCell.Row, 5).Value = "" Then
        For i = ActiveCell.Row To 7 Step -1
            If Cells(i, 5).Value <> "" Then
                ReNameRaporNormal = Cells(i, 5).Value & "-" & Cells(6, 7).Value & " " & Cells(ActiveCell.Row, 13).Value
                SiraNoIlkSatir = i
                GoTo RaporNoDonguSon
            End If
        Next i
    Else
        ReNameRaporNormal = Cells(ActiveCell.Row, 5).Value & "-" & Cells(6, 7).Value & " " & Cells(ActiveCell.Row, 13).Value
    End If
RaporNoDonguSon:
    'Rapor/Alt Rapor aralıklarının tespiti.
    Set IlkSiraBul = Range("FG7:FG100000").Find(What:=Cells(SiraNoIlkSatir, 5).Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    Set SonSiraBul = Range("FH7:FH100000").Find(What:=Cells(SiraNoIlkSatir, 5).Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    If Not IlkSiraBul Is Nothing Then
        IlkSira = IlkSiraBul.Row
    Else
        MsgBox "The serial number could not be found, so the operation cannot be completed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    If Not SonSiraBul Is Nothing Then
        SonSira = SonSiraBul.Row
    Else
        MsgBox "The serial number could not be found, so the operation cannot be completed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    'MsgBox IlkSira & "-" & SonSira
    For i = ActiveCell.Row To SonSira
        If Cells(i + 1, 7).Value <> "" And i < SonSira Then
            AltRaporNoIlk = ActiveCell.Row
            AltRaporNoSon = i
            'MsgBox AltRaporNoIlk & "-" & AltRaporNoSon
            GoTo RaporNoDonguSon1
        ElseIf i = SonSira Then
            AltRaporNoIlk = ActiveCell.Row
            AltRaporNoSon = i
            'MsgBox AltRaporNoIlk & "-" & AltRaporNoSon
            GoTo RaporNoDonguSon1
        End If
    Next i
RaporNoDonguSon1:

'________________________________________

    If TumDoc = False Then
        'Close the all Word application
        Call ModuleReport2.OpenWordControl
    
        'Operation klasöründeki docm uzantılı word dosyalarından açık olanları kapat ve temizle.
        OpenKontrolName = Dir(DestOpUserFolder & "*.docm")
        Do While OpenKontrolName <> ""
            OpenControl = IsFileOpen(DestOpUserFolder & OpenKontrolName)
            If OpenControl = True Then 'Açıksa
                On Error Resume Next
                Set objWord = GetObject(, "Word.Application")
                Set objWord = GetObject(, "Word.Application")
                Set objWord = GetObject(, "Word.Application")
                Set objWord = GetObject(, "Word.Application")
                Set objWord = GetObject(, "Word.Application")
                objWord.Quit SaveChanges:=True
                'MsgBox "Dosya OpenKontrol methodu ile kapatıldı."
    
            End If
            OpenKontrolName = Dir()
        Loop
    
    '    Call ModuleReport2.OpenWordControl
    
        Set objWord = Nothing
        Set objDoc = Nothing
    
    '________________________________________
    
        On Error Resume Next
    '    Klasörün içindeki tüm dosyaları sil (txt, docm vb.)
        ContSay = 0
        KontrolFile = Dir(DestOpUserFolder & "*.???")
        Do While KontrolFile <> ""
            ContSay = ContSay + 1
            KontrolFile = Dir()
        Loop
        If ContSay > 0 Then
            Kill DestOpUserFolder & "*.???"
        End If
    End If

'________________________________________Dinamik dosya isimleri

    ThisWorkbook.Activate
    RaporTipi = Cells(ActiveCell.Row, 214).Value
    SourceRaporNormal = AutoPath & "\System Files\System Templates\Report 2 Templates\" & RaporTipi & ".docm"
    'Klasör isimlerini kontrol et.
    If Not Dir(SourceRaporNormal, vbDirectory) <> vbNullString Then
        MsgBox "Cannot access the directory " & SourceRaporNormal & ". The names of folders and/or files in this directory may have been changed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    
    'Dosyayı şablondan operasyon klasörüne kopyala ve adını değiştir.
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.CopyFile (SourceRaporNormal), DestOpUserFolder & ReNameRaporNormal & ".docm", True
'________________________________________

    'Oluşturulacak dosyayı aç
    On Error Resume Next
    Set objWord = GetObject(, "Word.Application")
    Set objWord = GetObject(, "Word.Application")
    Set objWord = GetObject(, "Word.Application")
    Set objWord = GetObject(, "Word.Application")
    Set objWord = GetObject(, "Word.Application")
    If objWord Is Nothing Then
        'MsgBox "Dosya oluşturmada CreateObject methodu kullanılacak."
        Set objWord = CreateObject("Word.Application")
        objWord.Visible = False
    End If

    objWord.Documents.Open FileName:=DestOpUserFolder & ReNameRaporNormal & ".docm"
    objWord.Visible = True
    objWord.Activate 'Ekrana getirir.
    'objDoc.Activate 'Ekrana getirmez.
    objWord.Application.WindowState = wdWindowStateMaximize
    Set objDoc = GetObject(DestOpUserFolder & ReNameRaporNormal & ".docm")
'________________________________________

    'Birim
    Birimx = UCase(Replace(Replace(Worksheets(2).Cells(6, 99).Value, "i", "I"), "ı", "I")) & " UNIT"
    objDoc.Tables(1).Cell(Row:=1, Column:=1).Range.Text = Birimx
    'Rapor tarihi
    objDoc.Tables(1).Cell(Row:=4, Column:=3).Range.Text = Cells(IlkSira, 218).Value
    'Rapor No
    objDoc.Tables(1).Cell(Row:=5, Column:=3).Range.Text = Cells(AltRaporNoIlk, 217).Value

    objDoc.CheckBox1.Value = False 'Rapor Talebi
    objDoc.CheckBox2.Value = True 'Rapor3
    
    'Rapora esas yazı
    GelenTema = "The statement dated " & Cells(IlkSira, 95).Value & " concerning the person named " & Cells(IlkSira, 109).Value & "."
'
    
    'İlgi
    RaporIlgi = GelenTema
    If Cells(ActiveCell.Row, 212).Value <> "valid" Then
        objDoc.Tables(1).Cell(Row:=8, Column:=3).Range.Text = RaporIlgi
        'İlgili birim olmayacak. Çünkü rapor doğrudan KURUM_A tarafından düzenleniyor.
        objDoc.Tables(1).Rows(9).Delete
    Else
        objDoc.Tables(1).Cell(Row:=6, Column:=3).Range.Text = RaporIlgi
        'İlgili birim olmayacak. Çünkü rapor doğrudan KURUM_A tarafından düzenleniyor.
        objDoc.Tables(1).Rows(7).Delete
    End If

    'Tanımlamalar
    AdetTopla = Application.Sum(Range(Cells(AltRaporNoIlk, 136), Cells(AltRaporNoSon, 136)))
    If AdetTopla = 1 Then
        TekCogulTipA = "Type A item"
        Ek3 = "Type A Item"
    ElseIf AdetTopla > 1 Then
        TekCogulTipA = "Type A items"
        Ek3 = "Type A Items"
    End If
    
    objDoc.Tables(2).Cell(Row:=1, Column:=1).Range.Text = "The Examined " & Ek3
    
    'Tabloyu doldur
    j = 2
    For i = AltRaporNoIlk To AltRaporNoSon
        x = i - AltRaporNoIlk + 1
        If x Mod 2 <> 0 Then
            'Başlıklar
            objDoc.Tables(2).Cell(Row:=j + 0, Column:=1).Range.Text = "Item Type"
            objDoc.Tables(2).Cell(Row:=j + 1, Column:=1).Range.Text = "Item Value"
            objDoc.Tables(2).Cell(Row:=j + 2, Column:=1).Range.Text = "Qty"
            objDoc.Tables(2).Cell(Row:=j + 3, Column:=1).Range.Text = "Item ID No."
            objDoc.Tables(2).Cell(Row:=j + 0, Column:=2).Range.Text = ":"
            objDoc.Tables(2).Cell(Row:=j + 1, Column:=2).Range.Text = ":"
            objDoc.Tables(2).Cell(Row:=j + 2, Column:=2).Range.Text = ":"
            objDoc.Tables(2).Cell(Row:=j + 3, Column:=2).Range.Text = ":"
            'Verileri aktar
            objDoc.Tables(2).Cell(Row:=j + 0, Column:=3).Range.Text = Cells(i, 130).Value 'Öğe türü
            objDoc.Tables(2).Cell(Row:=j + 1, Column:=3).Range.Text = Cells(i, 133).Value 'Öğe değeri
            objDoc.Tables(2).Cell(Row:=j + 2, Column:=3).Range.Text = Cells(i, 136).Value 'Adet
            objDoc.Tables(2).Cell(Row:=j + 3, Column:=3).Range.Text = Cells(i, 139).Value 'Öğe ID
            j = j + 0
        ElseIf x Mod 2 = 0 Then
            'Verileri aktar
            objDoc.Tables(2).Cell(Row:=j + 0, Column:=4).Range.Text = Cells(i, 130).Value 'Öğe türü
            objDoc.Tables(2).Cell(Row:=j + 1, Column:=4).Range.Text = Cells(i, 133).Value 'Öğe değeri
            objDoc.Tables(2).Cell(Row:=j + 2, Column:=4).Range.Text = Cells(i, 136).Value 'Adet
            objDoc.Tables(2).Cell(Row:=j + 3, Column:=4).Range.Text = Cells(i, 139).Value 'Öğe ID
            'Satır ekle veya ekleme
            If i <> AltRaporNoSon Then
                For k = 1 To 5
                    objDoc.Tables(2).Rows.Add BeforeRow:=objDoc.Tables(2).Rows(j + 4)
                Next k
                j = j + 5
            End If
        End If
    Next i
    'Tema kodu
    objDoc.Tables(3).Cell(Row:=1, Column:=3).Range.Text = Cells(IlkSira, 98).Value  'Tema no

    'Rapor metin kısmı
    'Tanımlamalar
    AdetTopla = Application.Sum(Range(Cells(AltRaporNoIlk, 136), Cells(AltRaporNoSon, 136)))
    If AdetTopla = 1 Then
        Ek1 = "item"
    ElseIf AdetTopla > 1 Then
        Ek1 = "items"
    End If
    'Dosyada öğe/öğeler düzenlemesini gerçekleştir.
    Set MyRange = objDoc.Tables(3).Cell(Row:=2, Column:=1).Range
    With MyRange.Find
        .Text = "<item>"
        .Replacement.Text = TekCogulTipA
        .Execute Replace:=wdReplaceAll
    End With
    MyRange.HighlightColorIndex = wdNoHighlight
    
    'Artık satırı sil
    objDoc.Tables(2).Rows(j + 4).Delete
    
    'imzalar
    objDoc.Tables(6).Cell(Row:=3, Column:=1).Range.Text = Cells(ActiveCell.Row, 226).Value 'Ad Soyad Directorate
    objDoc.Tables(6).Cell(Row:=4, Column:=1).Range.Text = Cells(ActiveCell.Row, 227).Value 'Unvan
    objDoc.Tables(6).Cell(Row:=3, Column:=2).Range.Text = Cells(ActiveCell.Row, 220).Value 'Ad Soyad1
    objDoc.Tables(6).Cell(Row:=4, Column:=2).Range.Text = Cells(ActiveCell.Row, 221).Value 'Unvan1
    objDoc.Tables(6).Cell(Row:=3, Column:=3).Range.Text = Cells(ActiveCell.Row, 223).Value 'Ad Soyad2
    objDoc.Tables(6).Cell(Row:=4, Column:=3).Range.Text = Cells(ActiveCell.Row, 224).Value 'Unvan2

    'Alt bilgi ekle
    objDoc.Sections(1).Footers(wdHeaderFooterPrimary).Range.Tables(1).Cell(Row:=1, Column:=1).Range.Text = ReNameRaporNormal
    
    'Not ekle
    If Cells(ActiveCell.Row, 216).Value = "Yes" Then
        TxtFileNot = DestNotlar & Cells(ActiveCell.Row, 130).Value & ".txt"
        If Not Dir(TxtFileNot, vbDirectory) <> vbNullString Then
            MsgBox "Cannot access the directory " & TxtFileNot & ". The names of folders and/or files in this directory may have been changed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
            GoTo NotEkleAtla
        End If
        Open TxtFileNot For Input As #1
            Do Until EOF(1)
                Line Input #1, TextLine
                NotEkle = NotEkle & TextLine
            Loop
        Close #1
        objDoc.Tables(7).Cell(Row:=1, Column:=1).Range.Text = NotEkle
    End If
NotEkleAtla:

    'Technique A var/yok (SADECE RAPOR3 işleminde olan kısım)
    StrTeknik_ANotu = ""
    For i = AltRaporNoIlk To AltRaporNoSon
        If Cells(i, 212).Value = "invalid" And Left(Cells(i, 213).Value, 11) = "Technique A" Then
            'Notu ekle
            StrTeknik_ANotu = "*For invalid Type A items produced using Technical A, the Report 2.2 code can be determined based on the report prepared after the items are submitted to the XXX Directorate."
            '6. maddeyi ekle
            objDoc.Tables(4).Rows.Add
            objDoc.Tables(4).Cell(Row:=objDoc.Tables(4).Rows.Count, Column:=2).Range.Text = "Production technique appears to involve Technique A.*"
            'Rapor Talebi ve Rapor3 satırlarını kaldır
            objDoc.Tables(1).Rows(6).Delete
            objDoc.Tables(1).Rows(6).Delete
            GoTo Teknik_AOk
        End If
    Next i
Teknik_AOk:
    objDoc.Tables(8).Cell(Row:=1, Column:=1).Range.Text = StrTeknik_ANotu

    'Rapor sayfa sayısını oluşturmadan önce eskisini sil
    For i = IlkSira To SonSira
        If Cells(i, 13).Value = "" Then
            Cells(i, 174).Value = ""
        End If
    Next i
        
    'Sayfa sayısı kaydet komutuna bağlandı.
    objDoc.Save
    'Sayfayı text dosyasından çek
    TxtFileRapor = DestOpUserFolder & "Report 2 Page Count.txt"
    #If UseADO Then
        Const adReadLine = -2&
        With CreateObject("ADODB.Stream")
            .Open
            .LoadFromFile TxtFileRapor
            Do Until .EOS
                TotalSayfaRapor = .ReadText(adReadLine)
            Loop
            .Close
        End With
    #Else 'UseFSO
        Const RaporTristateTrue = -1&
        With CreateObject("Scripting.FileSystemObject")
            With .OpenTextFile(TxtFileRapor, Format:=RaporTristateTrue)
                Do Until .AtEndOfStream
                    TotalSayfaRapor = .ReadLine
                Loop
                .Close
            End With
        End With
    #End If
    'MsgBox "Rapor: " & TotalSayfaRapor
    Cells(ActiveCell.Row, 174).Value = TotalSayfaRapor

    objWord.Visible = True
    objWord.Activate

    If TumDoc = True Then
        objWord.Activate
'        For i = 1 To SayPrt
'            objDoc.PrintOut
'        Next i
        objDoc.PrintOut Background:=False, Copies:=SayPrt
        'objWord.Documents.Save
        objDoc.Close SaveChanges:=False
        objWord.Visible = False
    End If
    
    'Nesneleri temizle
    Set fso = Nothing
    Set objWord = Nothing
    Set objDoc = Nothing

ExplorerBos:

Next Explorer

If TumDoc = True Then
    'Tümünün bulunduğu butonu tekrar seç.
    Cells(b, 11).Select
End If

Son:

'Worksheets(4).Activate
Application.ScreenUpdating = True
Application.DisplayAlerts = True
'Application.EnableEvents = True

'End If

End Sub

Sub Rapor3_1Tutanak2()

Dim DestOperasyon As String, SourceRapor3_1Normal As String, AutoPath As String, ReNameRapor3_1 As String
Dim fso As Object, objWord As Object, objDoc As Object
Dim OpenKontrolName As String, OpenControl As String, DestOpUserFolder As String, DestOpUserFolderName As String
Dim ContSay As Integer, KontrolFile As String, x As Long, y As Long, i As Long, j As Long
Dim IlkSiraBul As Range, SonSiraBul As Range, IlkSira As Long, SonSira As Long
Dim GelenTema As String, SourceDL As String, ReNameDL As String
Dim TxtFileRapor3_1 As String, TotalSayfaRapor3_1 As String, TxtFileDokum As String, TotalSayfaDokum As String
Dim DokumFileTxt As Object, DokumSayfaGonder As String
Dim SourceRaporNormal As String, SourceTutanak2Normal As String, SourceUstYaziNormal As String
Dim Bolum1 As String, Bolum2 As String, Bolum3 As String, Bolum4 As String, Bolum5 As String, Bolum6 As String
Dim Ek1 As String, Ek2 As String, Ek3 As String, Ek4 As String, Ek5 As String, Ek6 As String, Birlestir As String

Dim ReNameRaporNormal As String, SiraNoIlkSatir As Long, RaporIlgi As String
Dim AltRaporNoIlk As Long, AltRaporNoSon As Long, AdetTopla As Long
Dim TekCogulTipA As String, k As Integer
Dim MyRange As Object, TxtFileRapor As String, TotalSayfaRapor As String
Dim ReNameTutanak2Normal As String, GidenTema As String
Dim TxtFileTutanak2 As String, TotalSayfaTutanak2 As String
Dim ReNameUstYaziNormal As String
Dim TxtFileUstYazi As String, TotalSayfaUstYazi As String
Dim Birimx As String, UserName As String, Kolluk As String

'TUTANAK2 için prosedürü başlat
'If ActiveCell.Column = 8 Then
'    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    'Application.EnableEvents = False
    
    'Dokümanların oluşturulmasına soldan sağa doğru bir kontrol mekanizması kur.
    j = ActiveCell.Row
    For i = ActiveCell.Row To 7 Step -1
        If Cells(i, 5).Value <> "" Then
            GoTo SiraNoBulundu
        Else
            j = j - 1
        End If
    Next i
SiraNoBulundu:
    If j < 7 Then
        MsgBox "The serial number could not be found, so the operation cannot be completed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    Set IlkSiraBul = Range("FG7:FG100000").Find(What:=Cells(j, 5).Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    Set SonSiraBul = Range("FH7:FH100000").Find(What:=Cells(j, 5).Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    If Not IlkSiraBul Is Nothing Then
        IlkSira = IlkSiraBul.Row
    Else
        MsgBox "The serial number could not be found, so the operation cannot be completed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    If Not SonSiraBul Is Nothing Then
        SonSira = SonSiraBul.Row
    Else
        MsgBox "The serial number could not be found, so the operation cannot be completed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    
    If Cells(IlkSira, 6).Value = "x" Then
        MsgBox "The Report 3 statement data contains missing or incorrect information, so the operation cannot be started.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    
    For i = IlkSira To SonSira
        If Cells(i, 7).Value = "x" Then
            MsgBox "The report data contains missing or incorrect information, so the operation cannot be started.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
            GoTo Son
        End If
    Next i

    
    
    'Pathfinder...
    AutoPath = ThisWorkbook.Path
    DestOperasyon = AutoPath & "\System Files\Operation\"

    'TUTANAK2 TANIMLARI
    SourceTutanak2Normal = AutoPath & "\System Files\System Templates\Statement 2 Template\Statement 2.docm"
    'ÜST YAZI TANIMLARI
    'SourceUstYaziNormal = AutoPath & "\System Files\System Templates\Report 2 Cover Letter Templates\Report 2 Cover Letter.docm"
    'Birden fazla kullanıcının operasyon klasöründe çakışmasını önlemek için operasyon klasörünün içinde yeni klasör aç.
    DestOpUserFolderName = "Operation-" & Left(ThisWorkbook.name, InStr(ThisWorkbook.name, ".") - 1) '& " " & Format(Now(), "ddmmyyyyhhmmss")
    DestOpUserFolder = DestOperasyon & DestOpUserFolderName & "\"

    'Check the "System Files" folder name.
    If Not Dir(AutoPath & "\System Files\", vbDirectory) <> vbNullString Then
        MsgBox "Cannot access the directory " & AutoPath & "\System Files\" & ". The folder named 'System Files' may have been renamed or deleted.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    
    'Check the "Operation" folder name.
    If Not Dir(DestOperasyon, vbDirectory) <> vbNullString Then
        MsgBox "Cannot access the directory " & DestOperasyon & ". The folder named 'Operation' may have been renamed or deleted.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    
    'Check folder names.
    If Not Dir(SourceTutanak2Normal, vbDirectory) <> vbNullString Then
        MsgBox "Cannot access the directory " & SourceTutanak2Normal & ". The names of folders and/or files in this directory may have been changed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If

    'Operation klasörü içinde kullanıcı modülü klasörü yoksa oluştur.
    If Not Dir(DestOpUserFolder, vbDirectory) <> vbNullString Then
        MkDir DestOpUserFolder
    End If

    'On Error Resume Next
    ReNameTutanak2Normal = Cells(ActiveCell.Row, 5).Value & "-" & Cells(6, 7).Value
'________________________________________

    If TumDoc = False Then
        'Close the all Word application
        Call ModuleReport3.OpenWordControl

        'Operation klasöründeki docm uzantılı word dosyalarından açık olanları kapat ve temizle.
        OpenKontrolName = Dir(DestOpUserFolder & "*.docm")
        Do While OpenKontrolName <> ""
            OpenControl = IsFileOpen(DestOpUserFolder & OpenKontrolName)
            If OpenControl = True Then 'Açıksa
                On Error Resume Next
                Set objWord = GetObject(, "Word.Application")
                Set objWord = GetObject(, "Word.Application")
                Set objWord = GetObject(, "Word.Application")
                Set objWord = GetObject(, "Word.Application")
                Set objWord = GetObject(, "Word.Application")
                objWord.Quit SaveChanges:=True
                'MsgBox "Dosya OpenKontrol methodu ile kapatıldı."

            End If
            OpenKontrolName = Dir()
        Loop

    '    Call ModuleReport3.OpenWordControl

        Set objWord = Nothing
        Set objDoc = Nothing
    '________________________________________

        On Error Resume Next
    '    Klasörün içindeki tüm dosyaları sil (txt, docm vb.)
        ContSay = 0
        KontrolFile = Dir(DestOpUserFolder & "*.???")
        Do While KontrolFile <> ""
            ContSay = ContSay + 1
            KontrolFile = Dir()
        Loop
        If ContSay > 0 Then
            Kill DestOpUserFolder & "*.???"
        End If
    End If

    'Dosyayı şablondan operasyon klasörüne kopyala ve adını değiştir.
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.CopyFile (SourceTutanak2Normal), DestOpUserFolder & ReNameTutanak2Normal & ".docm", True
'________________________________________

    'Oluşturulacak dosyayı aç
    On Error Resume Next
    Set objWord = GetObject(, "Word.Application")
    Set objWord = GetObject(, "Word.Application")
    Set objWord = GetObject(, "Word.Application")
    Set objWord = GetObject(, "Word.Application")
    Set objWord = GetObject(, "Word.Application")
    If objWord Is Nothing Then
        'MsgBox "Dosya oluşturmada CreateObject methodu kullanılacak."
        Set objWord = CreateObject("Word.Application")
        objWord.Visible = False
    End If
    objWord.Documents.Open FileName:=DestOpUserFolder & ReNameTutanak2Normal & ".docm"
    objWord.Visible = True
    objWord.Activate 'Ekrana getirir.
    'objDoc.Activate 'Ekrana getirmez.
    objWord.Application.WindowState = wdWindowStateMaximize
    Set objDoc = GetObject(DestOpUserFolder & ReNameTutanak2Normal & ".docm")
'________________________________________

    'Birim
    Birimx = UCase(Replace(Replace(Worksheets(2).Cells(6, 99).Value, "i", "I"), "ı", "I")) & " UNIT"
    objDoc.Tables(1).Cell(Row:=1, Column:=1).Range.Text = Birimx
    'Tutanak2 tarihi
    objDoc.Tables(1).Cell(Row:=4, Column:=3).Range.Text = Cells(ActiveCell.Row, 147).Value
    'Belge tarihi ve numarası
    objDoc.Tables(1).Rows(6).Delete

    'Kolluk
    If Cells(ActiveCell.Row, 102).Value = "Provincial Directorate B" Then
        Kolluk = Cells(ActiveCell.Row, 91).Value & " Provincial Governorship Provincial Directorate B " & Cells(ActiveCell.Row, 103).Value
    ElseIf Cells(ActiveCell.Row, 102).Value = "District Directorate B" Then
        Kolluk = Cells(ActiveCell.Row, 92).Value & " District Governorship District Directorate B " & Cells(ActiveCell.Row, 103).Value
    ElseIf Cells(ActiveCell.Row, 102).Value = "Provincial Directorate C" Then
        Kolluk = Cells(ActiveCell.Row, 91).Value & " Provincial Governorship Provincial Directorate C " & Cells(ActiveCell.Row, 103).Value
    ElseIf Cells(ActiveCell.Row, 102).Value = "District Directorate C" Then
        Kolluk = Cells(ActiveCell.Row, 92).Value & " District Governorship District Directorate C " & Cells(ActiveCell.Row, 103).Value
    ElseIf Cells(ActiveCell.Row, 102).Value = "Provincial Directorate D" Then
        Kolluk = Cells(ActiveCell.Row, 91).Value & " Provincial Governorship Provincial Directorate D " & Cells(ActiveCell.Row, 103).Value
    ElseIf Cells(ActiveCell.Row, 102).Value = "District Directorate D" Then
        Kolluk = Cells(ActiveCell.Row, 92).Value & " District Governorship District Directorate D " & Cells(ActiveCell.Row, 103).Value
    ElseIf Cells(ActiveCell.Row, 102).Value = "Provincial Directorate E" Then
        Kolluk = Cells(ActiveCell.Row, 91).Value & " Provincial Governorship Provincial Directorate E " & Cells(ActiveCell.Row, 103).Value
    ElseIf Cells(ActiveCell.Row, 102).Value = "District Directorate E" Then
        Kolluk = Cells(ActiveCell.Row, 92).Value & " District Governorship District Directorate E " & Cells(ActiveCell.Row, 103).Value
    ElseIf InStr(Cells(ActiveCell.Row, 102).Value, "General Directorate") <> 0 Or InStr(Cells(ActiveCell.Row, 102).Value, "Regional Directorate") <> 0 Then
        If Cells(ActiveCell.Row, 103).Value <> "" Then
            Kolluk = WorksheetFunction.Proper(ThisWorkbook.Worksheets(2).Cells(6, 111).Value) & " " & Cells(ActiveCell.Row, 102).Value & " " & Cells(ActiveCell.Row, 103).Value
        Else
            Kolluk = WorksheetFunction.Proper(ThisWorkbook.Worksheets(2).Cells(6, 111).Value) & " " & Cells(ActiveCell.Row, 102).Value
        End If
    Else
        If Cells(ActiveCell.Row, 103).Value <> "" Then
            If InStr(Cells(ActiveCell.Row, 102).Value, "X.X. ") > 0 Then
                Kolluk = Mid(Cells(ActiveCell.Row, 102).Value, 6, Len(Cells(ActiveCell.Row, 102).Value)) & " " & Cells(ActiveCell.Row, 103).Value
            Else
                Kolluk = Cells(ActiveCell.Row, 102).Value & " " & Cells(ActiveCell.Row, 103).Value
            End If
        Else
            If InStr(Cells(ActiveCell.Row, 102).Value, "X.X. ") > 0 Then
                Kolluk = Mid(Cells(ActiveCell.Row, 102).Value, 6, Len(Cells(ActiveCell.Row, 102).Value))
            Else
                Kolluk = Cells(ActiveCell.Row, 102).Value
            End If
        End If
    End If
    
    objDoc.Tables(1).Cell(Row:=5, Column:=3).Range.Text = Kolluk

    'Tabloyu doldur
    Set IlkSiraBul = Range("FG7:FG100000").Find(What:=Cells(ActiveCell.Row, 5).Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    Set SonSiraBul = Range("FH7:FH100000").Find(What:=Cells(ActiveCell.Row, 5).Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    If Not IlkSiraBul Is Nothing Then
        IlkSira = IlkSiraBul.Row
    Else
        MsgBox "The serial number could not be found, so the operation cannot be completed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    If Not SonSiraBul Is Nothing Then
        SonSira = SonSiraBul.Row
    Else
        MsgBox "The serial number could not be found, so the operation cannot be completed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    'Tabloya satır ekle
    x = 0
    If SonSira - IlkSira > 0 Then
        With objDoc.Tables(2)
            For i = IlkSira To SonSira - 1
                .Rows.Add BeforeRow:=objDoc.Tables(2).Rows(2)
                x = x + 1
            Next i
        End With
    End If

    For i = 2 To SonSira - IlkSira + 2
        objDoc.Tables(2).Cell(Row:=i, Column:=1).Range.Text = i - 1 'Tablo sıra no
        objDoc.Tables(2).Cell(Row:=i, Column:=2).Range.Text = Cells(IlkSira + i - 2, 130).Value 'Öğe türü
        objDoc.Tables(2).Cell(Row:=i, Column:=3).Range.Text = Cells(IlkSira + i - 2, 133).Value 'Öğe değeri
        objDoc.Tables(2).Cell(Row:=i, Column:=4).Range.Text = Cells(IlkSira + i - 2, 136).Value 'Adet
        objDoc.Tables(2).Cell(Row:=i, Column:=5).Range.Text = Cells(IlkSira + i - 2, 139).Value 'Öğe ID No
        objDoc.Tables(2).Cell(Row:=i, Column:=6).Range.Text = Cells(IlkSira, 98).Value  'Tema No (Temai her satıra yaz.)
        objDoc.Tables(2).Cell(Row:=i, Column:=7).Range.Text = Cells(IlkSira + i - 2, 142).Value 'Açıklama
    Next i

    'Tutanak2 tutanağının metin kısmı
    'Tanımlamalar
    If Application.Sum(Range(Cells(IlkSira, 136), Cells(SonSira, 136))) > 1 Then
        Ek4 = "items"
    Else
        Ek4 = "item"
    End If
    Bolum1 = vbTab & "A total of "
    Bolum2 = " " & Ek4 & " , as described above, have been placed inside "
    Bolum3 = " "
    Bolum4 = " and enclosed to be sent to the designated unit mentioned above."
    
    Ek1 = Application.Sum(Range(Cells(IlkSira, 136), Cells(SonSira, 136)))
    Ek2 = Cells(ActiveCell.Row, 150).Value
    Ek3 = Application.WorksheetFunction.Proper(Cells(ActiveCell.Row, 149).Value)
    Ek3 = Left(Ek3, InStr(Ek3, "/") - 1)
    
    Birlestir = Bolum1 & Ek1 & Bolum2 & Ek2 & Bolum3 & Ek3 & Bolum4
    objDoc.Tables(3).Cell(Row:=1, Column:=1).Range.Text = Birlestir
    
    Set MyRange = objDoc.Tables(3).Cell(Row:=1, Column:=1).Range
    MyRange.Find.Execute FindText:=Ek1
    MyRange.Font.Bold = True

    Set MyRange = objDoc.Tables(3).Cell(Row:=1, Column:=1).Range
    With MyRange.Find
        .Execute FindText:=Ek2
        .Execute Forward:=True
    End With
    MyRange.Font.Bold = True
    'Aralıkta bulunan karakterleri bold yap
    Set MyRange = objDoc.Tables(3).Cell(Row:=1, Column:=1).Range
    MyRange.Find.Execute FindText:=Ek3
    MyRange.Font.Bold = True

'    'İmza boşluğunu sayfaya sığdırmak için düzenle
    If SonSira - IlkSira + 1 = 17 Then
        For i = 1 To 2
            objDoc.Tables(4).Rows(1).Delete
        Next i
    ElseIf SonSira - IlkSira + 1 = 18 Then
        For i = 1 To 4
            objDoc.Tables(4).Rows(1).Delete
        Next i
    ElseIf SonSira - IlkSira + 1 = 19 Then
        For i = 1 To 6
            objDoc.Tables(4).Rows(1).Delete
        Next i
    ElseIf SonSira - IlkSira + 1 = 20 Then
        For i = 1 To 7
            objDoc.Tables(4).Rows(1).Delete
        Next i
        objDoc.Tables(4).Rows(1).Height = 5
    End If

    'imzalar
    objDoc.Tables(4).Cell(Row:=6, Column:=2).Range.Text = Cells(ActiveCell.Row, 193).Value 'Ad Soyad1
    objDoc.Tables(4).Cell(Row:=7, Column:=2).Range.Text = Cells(ActiveCell.Row, 194).Value 'Unvan1
    objDoc.Tables(4).Cell(Row:=6, Column:=3).Range.Text = Cells(ActiveCell.Row, 196).Value 'Ad Soyad2
    objDoc.Tables(4).Cell(Row:=7, Column:=3).Range.Text = Cells(ActiveCell.Row, 197).Value 'Unvan2
    
    'Alt bilgi ekle
    objDoc.Sections(1).Footers(wdHeaderFooterPrimary).Range.Tables(1).Cell(Row:=1, Column:=1).Range.Text = ReNameTutanak2Normal
    
    'Sayfa sayısı kaydet komutuna bağlandı.
'    objDoc.Close SaveChanges:=True
'    objWord.Quit
    objDoc.Save
    'Sayfayı text dosyasından çek
    TxtFileTutanak2 = DestOpUserFolder & "Statement 2 Page Count.txt"
    #If UseADO Then
        Const adReadLine = -2&
        With CreateObject("ADODB.Stream")
            .Open
            .LoadFromFile TxtFileTutanak2
            Do Until .EOS
                TotalSayfaTutanak2 = .ReadText(adReadLine)
            Loop
            .Close
        End With
    #Else 'UseFSO
        Const Tutanak2TristateTrue = -1&
        With CreateObject("Scripting.FileSystemObject")
            With .OpenTextFile(TxtFileTutanak2, Format:=Tutanak2TristateTrue)
                Do Until .AtEndOfStream
                    TotalSayfaTutanak2 = .ReadLine
                Loop
                .Close
            End With
        End With
    #End If
    'MsgBox "Tutanak2: " & TotalSayfaTutanak2

    'Tutanak2 sayfa sayısı
    Cells(ActiveCell.Row, 171).Value = TotalSayfaTutanak2

    objWord.Visible = True
    objWord.Activate

    If TumDoc = True Then
        objWord.Activate
'        For i = 1 To SayPrt
'            objDoc.PrintOut
'        Next i
        objDoc.PrintOut Background:=False, Copies:=SayPrt
        'objWord.Documents.Save
        objDoc.Close SaveChanges:=False
        objWord.Visible = False
    End If

    'Nesneleri temizle
    Set fso = Nothing
    Set objWord = Nothing
    Set objDoc = Nothing

Son:

'Worksheets(5).Activate
Application.ScreenUpdating = True
Application.DisplayAlerts = True
'Application.EnableEvents = True

'End If

End Sub

Sub Rapor3_1UstYazi()

Dim DestOperasyon As String, SourceRapor3_1Normal As String, AutoPath As String, ReNameRapor3_1 As String
Dim fso As Object, objWord As Object, objDoc As Object
Dim OpenKontrolName As String, OpenControl As String, DestOpUserFolder As String, DestOpUserFolderName As String
Dim ContSay As Integer, KontrolFile As String, x As Long, y As Long, i As Long, j As Long
Dim IlkSiraBul As Range, SonSiraBul As Range, IlkSira As Long, SonSira As Long
Dim GelenTema As String, SourceDL As String, ReNameDL As String
Dim TxtFileRapor3_1 As String, TotalSayfaRapor3_1 As String, TxtFileDokum As String, TotalSayfaDokum As String
Dim DokumFileTxt As Object, DokumSayfaGonder As String
Dim SourceRaporNormal As String, SourceTutanak2Normal As String, SourceUstYaziNormal As String
Dim Bolum1 As String, Bolum2 As String, Bolum3 As String, Bolum4 As String, Bolum5 As String, Bolum6 As String, Bolum7 As String
Dim Ek1 As String, Ek2 As String, Ek3 As String, Ek4 As String, Ek5 As String, Ek6 As String, Birlestir As String
Dim Ek7 As String, Ek8 As String, Ek9 As String, Ek_Paket As String

Dim ReNameRaporNormal As String, SiraNoIlkSatir As Long, RaporIlgi As String
Dim AltRaporNoIlk As Long, AltRaporNoSon As Long, AdetTopla As Long
Dim TekCogulTipA As String, k As Integer
Dim MyRange As Object, TxtFileRapor As String, TotalSayfaRapor As String
Dim ReNameTutanak2Normal As String, GidenTema As String
Dim TxtFileTutanak2 As String, TotalSayfaTutanak2 As String
Dim ReNameUstYaziNormal As String
Dim TxtFileUstYazi As String, TotalSayfaUstYazi As String
Dim Birimx As String, UserName As String
Dim IlBuyukHarf, IlKucukHarf, IlceBuyukHarf, IlceKucukHarf As String

Dim BStr As Integer
Dim M2 As Boolean, M3 As Boolean, M4 As Boolean
Dim Ifv As Boolean, Ify As Boolean
Dim XXXMudNotu As Boolean
Dim IlgiStrSay As Integer, Govde1StrSay As Integer, IlgiRng As Object, Govde1Rng As Object, Govde2Rng As Object
Dim IlgiFarkSay As Integer, Govde1FarkSay As Integer
Dim ustbilgimuhatap As String, ustbilgitarih As String, ustbilgisayi As String
Dim CokluSayfa As Integer
Dim ItemBul As Range, AdSoyadParaf As String, UnvanParaf As String, TelParaf As String
Dim GonderimUsulu As String
Dim StrTeknik_ANotu As String
Dim gecersizSay As Long


IlceSakla = ""
If InStr(Cells(ActiveCell.Row, 92).Value, " Organization A") <> 0 Then
    IlceSakla = Cells(ActiveCell.Row, 92).Value
    Cells(ActiveCell.Row, 92).Value = ""
End If

'ÜST YAZI için prosedürü başlat
'If ActiveCell.Column = 9 Then
'    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    'Application.EnableEvents = False

    'Dokümanların oluşturulmasına soldan sağa doğru bir kontrol mekanizması kur.
    j = ActiveCell.Row
    For i = ActiveCell.Row To 7 Step -1
        If Cells(i, 5).Value <> "" Then
            GoTo SiraNoBulundu
        Else
            j = j - 1
        End If
    Next i
SiraNoBulundu:
    If j < 7 Then
        MsgBox "The serial number could not be found, so the operation cannot be completed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    Set IlkSiraBul = Range("FG7:FG100000").Find(What:=Cells(j, 5).Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    Set SonSiraBul = Range("FH7:FH100000").Find(What:=Cells(j, 5).Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    If Not IlkSiraBul Is Nothing Then
        IlkSira = IlkSiraBul.Row
    Else
        MsgBox "The serial number could not be found, so the operation cannot be completed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    If Not SonSiraBul Is Nothing Then
        SonSira = SonSiraBul.Row
    Else
        MsgBox "The serial number could not be found, so the operation cannot be completed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    
    If Cells(IlkSira, 6).Value = "x" Then
        MsgBox "The Report 3 statement data contains missing or incorrect information, so the operation cannot be started.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    
    For i = IlkSira To SonSira
        If Cells(i, 7).Value = "x" Then
            MsgBox "The report data contains missing or incorrect information, so the operation cannot be started.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
            GoTo Son
        End If
    Next i
    
    If Cells(IlkSira, 8).Value = "x" Then
        MsgBox "The Statement 2 data contains missing or incorrect information, so the operation cannot be started.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If

    
    
    'Pathfinder...
    AutoPath = ThisWorkbook.Path
    DestOperasyon = AutoPath & "\System Files\Operation\"

    'ÜST YAZI TANIMLARI
    SourceUstYaziNormal = AutoPath & "\System Files\System Templates\Report 3 Cover Letter Templates\Report 3.1 – Type A Cover Letter.docm"
    'Birden fazla kullanıcının operasyon klasöründe çakışmasını önlemek için operasyon klasörünün içinde yeni klasör aç.
    DestOpUserFolderName = "Operation-" & Left(ThisWorkbook.name, InStr(ThisWorkbook.name, ".") - 1) '& " " & Format(Now(), "ddmmyyyyhhmmss")
    DestOpUserFolder = DestOperasyon & DestOpUserFolderName & "\"

    'Check the "System Files" folder name.
    If Not Dir(AutoPath & "\System Files\", vbDirectory) <> vbNullString Then
        MsgBox "Cannot access the directory " & AutoPath & "\System Files\" & ". The folder named 'System Files' may have been renamed or deleted.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    
    'Check the "Operation" folder name.
    If Not Dir(DestOperasyon, vbDirectory) <> vbNullString Then
        MsgBox "Cannot access the directory " & DestOperasyon & ". The folder named 'Operation' may have been renamed or deleted.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    
    'Check folder names.
    If Not Dir(SourceUstYaziNormal, vbDirectory) <> vbNullString Then
        MsgBox "Cannot access the directory " & SourceUstYaziNormal & ". The names of folders and/or files in this directory may have been changed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If

    'Operation klasörü içinde kullanıcı modülü klasörü yoksa oluştur.
    If Not Dir(DestOpUserFolder, vbDirectory) <> vbNullString Then
        MkDir DestOpUserFolder
    End If

    'On Error Resume Next
    Set IlkSiraBul = Range("FG7:FG100000").Find(What:=Cells(ActiveCell.Row, 5).Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    Set SonSiraBul = Range("FH7:FH100000").Find(What:=Cells(ActiveCell.Row, 5).Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    If Not IlkSiraBul Is Nothing Then
        IlkSira = IlkSiraBul.Row
    Else
        MsgBox "The serial number could not be found, so the operation cannot be completed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    If Not SonSiraBul Is Nothing Then
        SonSira = SonSiraBul.Row
    Else
        MsgBox "The serial number could not be found, so the operation cannot be completed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If

    'Pre-checks for pages to be specified in attachments
    
    'Report 3 Statement check
    If Cells(IlkSira, 169).Value = "" Then
        MsgBox "Cover letter for sequence number " & Cells(IlkSira, 5).Value & " cannot be prepared because Report 3 Statement has not been created. Please create Report 3 Statement for sequence number " & Cells(IlkSira, 5).Value & " and try again.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    
    'Statement 2 check
    If Cells(IlkSira, 171).Value = "" Then
        MsgBox "Cover letter for sequence number " & Cells(IlkSira, 5).Value & " cannot be prepared because Statement 2 has not been created. Please create Statement 2 for sequence number " & Cells(IlkSira, 5).Value & " and try again.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If


    'PARAF HAZIRLIK İŞLEMİ
    UserName = Environ("UserProfile")
    UserName = UCase(Right(UserName, 7))
    Set ItemBul = Worksheets(2).Range("DR6:DR1000").Find(What:=UserName, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    If Not ItemBul Is Nothing Then
        AdSoyadParaf = "For information only: " & Worksheets(2).Cells(ItemBul.Row, 123).Value
        UnvanParaf = Worksheets(2).Cells(ItemBul.Row, 124).Value
        TelParaf = "Phone / Extension No: " & Worksheets(2).Cells(ItemBul.Row, 126).Value
    Else
        MsgBox UserName & " session is not registered in the system, so the cover letter cannot be created. Please register your session in the system using the Initials Interface located in the Settings Group of the Enterprise Document Automation System menu, then try the operation again.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    
    'On Error Resume Next
    ReNameUstYaziNormal = Cells(ActiveCell.Row, 5).Value & "-" & Cells(6, 10).Value
'________________________________________

    If TumDoc = False Then
        'Close the all Word application
        Call ModuleReport3.OpenWordControl

        'Operation klasöründeki docm uzantılı word dosyalarından açık olanları kapat ve temizle.
        OpenKontrolName = Dir(DestOpUserFolder & "*.docm")
        Do While OpenKontrolName <> ""
            OpenControl = IsFileOpen(DestOpUserFolder & OpenKontrolName)
            If OpenControl = True Then 'Açıksa
    '            On Error Resume Next
    '            Set objWord = GetObject(, "Word.Application")
    '            If objWord Is Nothing Then
    '                Set objWord = CreateObject("Word.Application")
    '                objWord.Visible = True
    '            End If
    '            objWord.Quit SaveChanges:=True

                On Error Resume Next
                Set objWord = GetObject(, "Word.Application")
                Set objWord = GetObject(, "Word.Application")
                Set objWord = GetObject(, "Word.Application")
                Set objWord = GetObject(, "Word.Application")
                Set objWord = GetObject(, "Word.Application")
                objWord.Quit SaveChanges:=True
                'MsgBox "Dosya OpenKontrol methodu ile kapatıldı."

            End If
            OpenKontrolName = Dir()
        Loop

    '    Call ModuleReport3.OpenWordControl

        Set objWord = Nothing
        Set objDoc = Nothing

    '________________________________________

        On Error Resume Next
    '    Klasörün içindeki tüm dosyaları sil (txt, docm vb.)
        ContSay = 0
        KontrolFile = Dir(DestOpUserFolder & "*.???")
        Do While KontrolFile <> ""
            ContSay = ContSay + 1
            KontrolFile = Dir()
        Loop
        If ContSay > 0 Then
            Kill DestOpUserFolder & "*.???"
        End If
    End If

    'Dosyayı şablondan operasyon klasörüne kopyala ve adını değiştir.
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.CopyFile (SourceUstYaziNormal), DestOpUserFolder & ReNameUstYaziNormal & ".docm", True
'________________________________________

    'Oluşturulacak dosyayı aç
    On Error Resume Next
    Set objWord = GetObject(, "Word.Application")
    Set objWord = GetObject(, "Word.Application")
    Set objWord = GetObject(, "Word.Application")
    Set objWord = GetObject(, "Word.Application")
    Set objWord = GetObject(, "Word.Application")
    If objWord Is Nothing Then
        'MsgBox "Dosya oluşturmada CreateObject methodu kullanılacak."
        Set objWord = CreateObject("Word.Application")
        objWord.Visible = False
    End If
    objWord.Documents.Open FileName:=DestOpUserFolder & ReNameUstYaziNormal & ".docm"
    objWord.Visible = True
    objWord.Activate 'Ekrana getirir.
    'objDoc.Activate 'Ekrana getirmez.
    objWord.Application.WindowState = wdWindowStateMaximize
    Set objDoc = GetObject(DestOpUserFolder & ReNameUstYaziNormal & ".docm")
'________________________________________

    'Birim
    Birimx = UCase(Replace(Replace(Worksheets(2).Cells(6, 99).Value, "i", "I"), "ı", "I")) & " UNIT"
    objDoc.Tables(1).Cell(Row:=1, Column:=1).Range.Text = Birimx

    Birimx = UCase(Replace(Replace(Worksheets(2).Cells(6, 99).Value, "i", "I"), "ı", "I"))
    'Üst yazı tarihi
    objDoc.Tables(1).Cell(Row:=1, Column:=2).Range.Text = Birimx & ", " & FormatEnglishDate(Cells(ActiveCell.Row, 155).Value) '(Format(Cells(ActiveCell.Row, 155).Value, "d mmmm yyyy"))
    'Yazı no
    objDoc.Tables(1).Cell(Row:=4, Column:=2).Range.Text = Cells(ActiveCell.Row, 156).Value
    'Sigorta türü
    Ek2 = Cells(ActiveCell.Row, 149).Value
    Ek2 = Right(Ek2, Len(Ek2) - InStr(Ek2, "/"))
    objDoc.Tables(1).Cell(Row:=7, Column:=2).Range.Text = Ek2


    'Muhatap
    IlBuyukHarf = UCase(Replace(Replace(Cells(ActiveCell.Row, 91).Value, "i", "I"), "ı", "I"))
    IlKucukHarf = Cells(ActiveCell.Row, 91).Value
    If Cells(ActiveCell.Row, 92).Value <> "" Then
        IlceBuyukHarf = UCase(Replace(Replace(Cells(ActiveCell.Row, 92).Value, "i", "I"), "ı", "I"))
        IlceKucukHarf = Cells(ActiveCell.Row, 92).Value
    Else
        IlceBuyukHarf = ""
        IlceKucukHarf = ""
    End If
    If Cells(ActiveCell.Row, 92).Value <> "" Then
        Bolum3 = IlceKucukHarf & "/" & IlBuyukHarf
    Else
        Bolum3 = IlBuyukHarf
    End If
    

    'YENİ MUHATAP TEMASI
    
    M2 = False
    M3 = False
    M4 = False
'    Ifv = False
'    Ify = False
    'TipB = False
    BStr = 10
    '1. sayfadan sonraki üst bilgi tanımları
    ustbilgitarih = (Format(Cells(ActiveCell.Row, 155).Value, "dd.mm.yyyy"))
    ustbilgisayi = Cells(ActiveCell.Row, 156).Value
    
    '4'lük
    If Cells(ActiveCell.Row, 103).Value <> "" Then
        If Cells(ActiveCell.Row, 102).Value = "Provincial Directorate B" Or Cells(ActiveCell.Row, 102).Value = "Provincial Directorate C" Or _
        Cells(ActiveCell.Row, 102).Value = "Provincial Directorate D" Or Cells(ActiveCell.Row, 102).Value = "Provincial Directorate E" Then 'VALİLİK
            objDoc.Tables(1).Cell(Row:=BStr, Column:=1).Range.Text = IlBuyukHarf & " PROVINCIAL GOVERNORSHIP"
            objDoc.Tables(1).Cell(Row:=BStr + 1, Column:=1).Range.Text = Cells(ActiveCell.Row, 102).Value  'UCase(Replace(Replace(Cells(ActiveCell.Row, 102).Value, "i", "I"), "ı", "I"))
            objDoc.Tables(1).Cell(Row:=BStr + 2, Column:=1).Range.Text = "(" & Cells(ActiveCell.Row, 103).Value & ")"
            objDoc.Tables(1).Cell(Row:=BStr + 3, Column:=1).Range.Text = Bolum3
            objDoc.Tables(1).Cell(Row:=BStr + 3, Column:=1).Range.Font.Underline = wdUnderlineSingle
            M4 = True
            
            ustbilgimuhatap = IlKucukHarf & " Provincial Governorship " & Cells(ActiveCell.Row, 102).Value
        ElseIf Cells(ActiveCell.Row, 102).Value = "District Directorate B" Or Cells(ActiveCell.Row, 102).Value = "District Directorate C" Or _
        Cells(ActiveCell.Row, 102).Value = "District Directorate D" Or Cells(ActiveCell.Row, 102).Value = "District Directorate E" Then 'KAYMAKAMLIK
            objDoc.Tables(1).Cell(Row:=BStr, Column:=1).Range.Text = IlceBuyukHarf & " DISTRICT GOVERNORSHIP"
            objDoc.Tables(1).Cell(Row:=BStr + 1, Column:=1).Range.Text = Cells(ActiveCell.Row, 102).Value 'UCase(Replace(Replace(Cells(ActiveCell.Row, 102).Value, "i", "I"), "ı", "I"))
            objDoc.Tables(1).Cell(Row:=BStr + 2, Column:=1).Range.Text = "(" & Cells(ActiveCell.Row, 103).Value & ")"
            objDoc.Tables(1).Cell(Row:=BStr + 3, Column:=1).Range.Text = Bolum3
            objDoc.Tables(1).Cell(Row:=BStr + 3, Column:=1).Range.Font.Underline = wdUnderlineSingle
            M4 = True
            
            ustbilgimuhatap = IlceKucukHarf & " District Governorship " & Cells(ActiveCell.Row, 102).Value
        ElseIf InStr(Cells(ActiveCell.Row, 102).Value, "General Directorate") <> 0 Or InStr(Cells(ActiveCell.Row, 102).Value, "Regional Directorate") <> 0 Then
            objDoc.Tables(1).Cell(Row:=BStr, Column:=1).Range.Text = UCase(Replace(Replace(ThisWorkbook.Worksheets(2).Cells(6, 111).Value, "i", "I"), "ı", "I"))
            objDoc.Tables(1).Cell(Row:=BStr + 1, Column:=1).Range.Text = Cells(ActiveCell.Row, 102).Value 'UCase(Replace(Replace(Cells(ActiveCell.Row, 102).Value, "i", "I"), "ı", "I"))
            objDoc.Tables(1).Cell(Row:=BStr + 2, Column:=1).Range.Text = "(" & Cells(ActiveCell.Row, 103).Value & ")"
            objDoc.Tables(1).Cell(Row:=BStr + 3, Column:=1).Range.Text = Bolum3
            objDoc.Tables(1).Cell(Row:=BStr + 3, Column:=1).Range.Font.Underline = wdUnderlineSingle
            M4 = True
            
            ustbilgimuhatap = WorksheetFunction.Proper(ThisWorkbook.Worksheets(2).Cells(6, 111).Value) & " " & Cells(ActiveCell.Row, 102).Value
        Else 'YARGI 3'lük
            Bolum1 = UCase(Replace(Replace(Cells(ActiveCell.Row, 102).Value, "i", "I"), "ı", "I"))
            If InStr(Bolum1, "X.X. ") > 0 Then
                objDoc.Tables(1).Cell(Row:=BStr, Column:=1).Range.Text = Mid(Bolum1, 6, Len(Bolum1))
                objDoc.Tables(1).Cell(Row:=BStr + 1, Column:=1).Range.Text = "(" & Cells(ActiveCell.Row, 103).Value & ")"
                objDoc.Tables(1).Cell(Row:=BStr + 2, Column:=1).Range.Text = Bolum3
                objDoc.Tables(1).Cell(Row:=BStr + 2, Column:=1).Range.Font.Underline = wdUnderlineSingle
                M3 = True
                
                ustbilgimuhatap = Mid(Cells(ActiveCell.Row, 102).Value, 6, Len(Cells(ActiveCell.Row, 102).Value))
            Else
                objDoc.Tables(1).Cell(Row:=BStr, Column:=1).Range.Text = Bolum1
                objDoc.Tables(1).Cell(Row:=BStr + 1, Column:=1).Range.Text = "(" & Cells(ActiveCell.Row, 103).Value & ")"
                objDoc.Tables(1).Cell(Row:=BStr + 2, Column:=1).Range.Text = Bolum3
                objDoc.Tables(1).Cell(Row:=BStr + 2, Column:=1).Range.Font.Underline = wdUnderlineSingle
                M3 = True

                ustbilgimuhatap = Cells(ActiveCell.Row, 102).Value
            End If
        End If
    End If
    '3'lük
    If Cells(ActiveCell.Row, 103).Value = "" Then
        If Cells(ActiveCell.Row, 102).Value = "Provincial Directorate B" Or Cells(ActiveCell.Row, 102).Value = "Provincial Directorate C" Or _
        Cells(ActiveCell.Row, 102).Value = "Provincial Directorate D" Or Cells(ActiveCell.Row, 102).Value = "Provincial Directorate E" Then 'VALİLİK
            objDoc.Tables(1).Cell(Row:=BStr, Column:=1).Range.Text = IlBuyukHarf & " PROVINCIAL GOVERNORSHIP"
            objDoc.Tables(1).Cell(Row:=BStr + 1, Column:=1).Range.Text = "(" & Cells(ActiveCell.Row, 102).Value & ")"  'UCase(Replace(Replace(Cells(ActiveCell.Row, 102).Value, "i", "I"), "ı", "I"))
            objDoc.Tables(1).Cell(Row:=BStr + 2, Column:=1).Range.Text = Bolum3
            objDoc.Tables(1).Cell(Row:=BStr + 2, Column:=1).Range.Font.Underline = wdUnderlineSingle
            M3 = True
            
            ustbilgimuhatap = IlKucukHarf & " Provincial Governorship " & Cells(ActiveCell.Row, 102).Value
        ElseIf Cells(ActiveCell.Row, 102).Value = "District Directorate B" Or Cells(ActiveCell.Row, 102).Value = "District Directorate C" Or _
        Cells(ActiveCell.Row, 102).Value = "District Directorate D" Or Cells(ActiveCell.Row, 102).Value = "District Directorate E" Then
            objDoc.Tables(1).Cell(Row:=BStr, Column:=1).Range.Text = IlceBuyukHarf & " DISTRICT GOVERNORSHIP"
            objDoc.Tables(1).Cell(Row:=BStr + 1, Column:=1).Range.Text = "(" & Cells(ActiveCell.Row, 102).Value & ")"  'UCase(Replace(Replace(Cells(ActiveCell.Row, 102).Value, "i", "I"), "ı", "I"))
            objDoc.Tables(1).Cell(Row:=BStr + 2, Column:=1).Range.Text = Bolum3
            objDoc.Tables(1).Cell(Row:=BStr + 2, Column:=1).Range.Font.Underline = wdUnderlineSingle
            M3 = True
            
            ustbilgimuhatap = IlceKucukHarf & " District Governorship " & Cells(ActiveCell.Row, 102).Value
        ElseIf InStr(Cells(ActiveCell.Row, 102).Value, "General Directorate") <> 0 Or InStr(Cells(ActiveCell.Row, 102).Value, "Regional Directorate") <> 0 Then
            objDoc.Tables(1).Cell(Row:=BStr, Column:=1).Range.Text = UCase(Replace(Replace(ThisWorkbook.Worksheets(2).Cells(6, 111).Value, "i", "I"), "ı", "I"))
            objDoc.Tables(1).Cell(Row:=BStr + 1, Column:=1).Range.Text = "(" & Cells(ActiveCell.Row, 102).Value & ")"  'UCase(Replace(Replace(Cells(ActiveCell.Row, 102).Value, "i", "I"), "ı", "I"))
            objDoc.Tables(1).Cell(Row:=BStr + 2, Column:=1).Range.Text = Bolum3
            objDoc.Tables(1).Cell(Row:=BStr + 2, Column:=1).Range.Font.Underline = wdUnderlineSingle
            M3 = True

            ustbilgimuhatap = WorksheetFunction.Proper(ThisWorkbook.Worksheets(2).Cells(6, 111).Value) & " " & Cells(ActiveCell.Row, 102).Value
        Else 'YARGI 2'lik
            Bolum1 = UCase(Replace(Replace(Cells(ActiveCell.Row, 102).Value, "i", "I"), "ı", "I"))
            If InStr(Bolum1, "X.X. ") > 0 Then
                objDoc.Tables(1).Cell(Row:=BStr, Column:=1).Range.Text = Mid(Bolum1, 6, Len(Bolum1))
                objDoc.Tables(1).Cell(Row:=BStr + 1, Column:=1).Range.Text = Bolum3
                objDoc.Tables(1).Cell(Row:=BStr + 1, Column:=1).Range.Font.Underline = wdUnderlineSingle
                M2 = True

                ustbilgimuhatap = Mid(Cells(ActiveCell.Row, 102).Value, 6, Len(Cells(ActiveCell.Row, 102).Value))
            Else
                objDoc.Tables(1).Cell(Row:=BStr, Column:=1).Range.Text = Bolum1
                objDoc.Tables(1).Cell(Row:=BStr + 1, Column:=1).Range.Text = Bolum3
                objDoc.Tables(1).Cell(Row:=BStr + 1, Column:=1).Range.Font.Underline = wdUnderlineSingle
                M2 = True

                ustbilgimuhatap = Cells(ActiveCell.Row, 102).Value
            End If
        End If
    End If
   
    'tipAnın/tipAların (TipA adedi)
    AdetTopla = Application.Sum(Range(Cells(IlkSira, 136), Cells(SonSira, 136)))
    If AdetTopla = 1 Then
        Ek1 = "item"
        Ek2 = "item was"
        Ek7 = "has"
    ElseIf AdetTopla > 1 Then
        Ek1 = "items"
        Ek2 = "items were"
        Ek7 = "have"
    End If

    ' Report body (English version)
    ' Report(s) control
    Dim ReportList() As String
    ' Rapor numaralarını diziye al
    y = 0
    For i = IlkSira To SonSira
        If Cells(i, 7).Value <> "" Then
            y = y + 1
            ReDim Preserve ReportList(1 To y)
            ReportList(y) = Cells(i, 13).Value
        End If
    Next i
    
    ' Listeyi biçimlendir
    Select Case y
        Case 1
            Ek3 = ReportList(1)
        Case 2
            Ek3 = ReportList(1) & " and " & ReportList(2)
        Case Else
            Ek3 = ""
            For i = 1 To y - 1
                Ek3 = Ek3 & ReportList(i) & ", "
            Next i
            Ek3 = Ek3 & "and " & ReportList(y)
    End Select

    If y = 1 Then
        Ek4 = "Report 2.1"
        Ek5 = "is"
    Else
        Ek4 = "Report 2.1s"
        Ek5 = "are"
    End If

    'Technique A var/yok (SADECE RAPOR3 işleminde olan kısım)
    StrTeknik_ANotu = ""
    For i = IlkSira To SonSira
        If Cells(i, 212).Value = "invalid" And Left(Cells(i, 213).Value, 11) = "Technique A" Then
            StrTeknik_ANotu = "Furthermore, since the Type A " & Ek2 & " determined to be invalid based on the assessment of xxxxxx xxxxxx xxxxxx / xxxxxx xxxxxx xxx xxxxxx, the preparation of the related Report 2.2 is only possible upon submission of the " & Ek1 & " to the XXX Directorate."
            GoTo Teknik_AOk1
        End If
    Next i
Teknik_AOk1:

    Ek8 = Cells(ActiveCell.Row, 95).Value 'Tutanak tarihi
    Ek9 = Cells(ActiveCell.Row, 109).Value 'Ad Soyad
    Ek6 = Cells(ActiveCell.Row, 98).Value 'Tema no
        
    'imzalar
    objDoc.Tables(3).Cell(Row:=4, Column:=2).Range.Text = Cells(ActiveCell.Row, 205).Value 'Ad Soyad1
    objDoc.Tables(3).Cell(Row:=5, Column:=2).Range.Text = Cells(ActiveCell.Row, 206).Value 'Unvan1
    objDoc.Tables(3).Cell(Row:=4, Column:=3).Range.Text = Cells(ActiveCell.Row, 208).Value 'Ad Soyad2
    objDoc.Tables(3).Cell(Row:=5, Column:=3).Range.Text = Cells(ActiveCell.Row, 209).Value 'Unvan2
        
    'Üst yazı gövde metni ve ek senaryoları
    If Cells(IlkSira, 124).Value = "Yes" Then 'Organization B mensubu var.

        If Mid(Cells(IlkSira, 128).Value, 7, 2) = "31" Then 'TipAyı rızası ile verdi.

            Bolum2 = Ek4 & " dated " & Cells(IlkSira, 218).Value & " and numbered " & Ek3 & ", concerning the Type A " & Ek1 & ", " & Ek5 & " hereby attached to be delivered to the relevant Process Monitoring Directorate along with the " & Ek1 & "."

            'Kişinin açık kimliği var/yok
            If Cells(ActiveCell.Row, 126).Value <> "" And Cells(ActiveCell.Row, 126).Value > 0 Then
                Bolum1 = "The Type A " & Ek1 & " delivered to our unit on " & Ek8 & " by the person named " & Ek9 & ", whose full identity is stated in the attached statement, " & _
                        Ek7 & " been evaluated as invalid. The item type, item value, quantity, and item ID number are also specified in the same statement, " & _
                        "and the " & Ek1 & " have been assigned the theme number " & Ek6 & " by our office."
            Else
                Bolum1 = "The Type A " & Ek1 & " delivered to our unit on " & Ek8 & " by the person named " & Ek9 & _
                        " " & Ek7 & " been evaluated as invalid. The item type, item value, quantity, and item ID number are also specified in the same statement, " & _
                        "and the " & Ek1 & " have been assigned the theme number " & Ek6 & " by our office."
            End If
            
            Bolum5 = Chr(9) & "Respectfully submitted for your information."
            
            If StrTeknik_ANotu <> "" Then
                Birlestir = Chr(9) & Bolum1 & vbNewLine & Chr(9) & Bolum2 & vbNewLine & Chr(9) & StrTeknik_ANotu
            Else
                Birlestir = Chr(9) & Bolum1 & vbNewLine & Chr(9) & Bolum2
            End If
            objDoc.Tables(2).Cell(Row:=1, Column:=1).Range.Text = Birlestir
            objDoc.Tables(2).Cell(Row:=3, Column:=1).Range.Text = Bolum5
            
    
            'Ekler
            objDoc.Tables(3).Cell(Row:=8, Column:=1).Range.Text = "Attachment:"

            Ek_Paket = Cells(ActiveCell.Row, 149).Value 'Kapalı Package A
            Ek_Paket = Left(Ek_Paket, InStr(Ek_Paket, "/") - 1)
            If Cells(ActiveCell.Row, 150).Value > 1 Then objDoc.Tables(3).Cell(Row:=8, Column:=2).Range.Text = " 1) Enclosed " & Ek_Paket & " (" & Cells(ActiveCell.Row, 150).Value & " pieces)"
            If Cells(ActiveCell.Row, 150).Value < 2 Then objDoc.Tables(3).Cell(Row:=8, Column:=2).Range.Text = " 1) Enclosed " & Ek_Paket & " (" & Cells(ActiveCell.Row, 150).Value & " piece)"
            
            x = Application.Sum(Range(Cells(IlkSira, 171), Cells(SonSira, 171))) 'Statement 2 toplam sayfa sayısı
            If x > 1 Then objDoc.Tables(3).Cell(Row:=9, Column:=2).Range.Text = " 2) Statement 2 (" & x & " pages)"
            If x < 2 Then objDoc.Tables(3).Cell(Row:=9, Column:=2).Range.Text = " 2) Statement 2 (" & x & " page)"
            
            x = Application.Sum(Range(Cells(IlkSira, 174), Cells(SonSira, 174)))  'Rapor1 toplam sayfa sayısı
            If y = 1 Then
                If x > 1 Then objDoc.Tables(3).Cell(Row:=10, Column:=2).Range.Text = " 3) Report 2.1 (" & x & " pages)"
                If x < 2 Then objDoc.Tables(3).Cell(Row:=10, Column:=2).Range.Text = " 3) Report 2.1 (" & x & " page)"
            Else
                objDoc.Tables(3).Cell(Row:=10, Column:=2).Range.Text = " 3) Report 2.1 (" & y & " reports, " & "total of " & x & " pages)"
            End If
            
            If Cells(ActiveCell.Row, 126).Value <> "" And Cells(ActiveCell.Row, 126).Value > 0 Then 'Kimlik var
                x = Application.Sum(Range(Cells(IlkSira, 169), Cells(SonSira, 169))) + Cells(ActiveCell.Row, 126).Value 'Rapor3 Tutanağı
                objDoc.Tables(3).Cell(Row:=11, Column:=2).Range.Text = " 4) Attached Statement (total of " & x & " pages)"
            Else 'Kimlik yok
                x = Application.Sum(Range(Cells(IlkSira, 169), Cells(SonSira, 169))) 'Rapor3 Tutanağı
                If x > 1 Then objDoc.Tables(3).Cell(Row:=11, Column:=2).Range.Text = " 4) Statement (" & x & " pages)"
                If x < 2 Then objDoc.Tables(3).Cell(Row:=11, Column:=2).Range.Text = " 4) Statement (" & x & " page)"
            End If
    
        ElseIf Mid(Cells(IlkSira, 128).Value, 13, 2) = "51" Then 'tipAya Organization B mensubu el koydu.

            Bolum2 = Ek4 & " dated " & Cells(IlkSira, 218).Value & " and numbered " & Ek3 & ", concerning the Type A " & Ek1 & ", " & Ek5 & " hereby attached to be delivered to the relevant Process Monitoring Directorate."

            'Kişinin açık kimliği var/yok
            If Cells(ActiveCell.Row, 126).Value <> "" And Cells(ActiveCell.Row, 126).Value > 0 Then
                Bolum1 = "The Type A " & Ek1 & " delivered to our unit on " & Ek8 & " by the person named " & Ek9 & ", whose full identity is stated in the attached statement, " & _
                        Ek7 & " been evaluated as invalid and confiscated by " & Cells(ActiveCell.Row, 125).Value & "." & " The item type, item value, quantity, and item ID number are also specified in the same statement, " & _
                        "and the " & Ek1 & " have been assigned the theme number " & Ek6 & " by our office."
            Else
                Bolum1 = "The Type A " & Ek1 & " delivered to our unit on " & Ek8 & " by the person named " & Ek9 & _
                        " " & Ek7 & " been evaluated as invalid and confiscated by " & Cells(ActiveCell.Row, 125).Value & "." & " The item type, item value, quantity, and item ID number are also specified in the same statement, " & _
                        "and the " & Ek1 & " have been assigned the theme number " & Ek6 & " by our office."
            End If

            Bolum5 = Chr(9) & "Respectfully submitted for your information."
            
            If StrTeknik_ANotu <> "" Then
                Birlestir = Chr(9) & Bolum1 & vbNewLine & Chr(9) & Bolum2 & vbNewLine & Chr(9) & StrTeknik_ANotu
            Else
                Birlestir = Chr(9) & Bolum1 & vbNewLine & Chr(9) & Bolum2
            End If
            objDoc.Tables(2).Cell(Row:=1, Column:=1).Range.Text = Birlestir
            objDoc.Tables(2).Cell(Row:=3, Column:=1).Range.Text = Bolum5

            'Ekler
            objDoc.Tables(3).Cell(Row:=8, Column:=1).Range.Text = "Attachment:"
            
            x = Application.Sum(Range(Cells(IlkSira, 174), Cells(SonSira, 174)))  'Rapor1 toplam sayfa sayısı
            If y = 1 Then
                If x > 1 Then objDoc.Tables(3).Cell(Row:=8, Column:=2).Range.Text = " 1) Report 2.1 (" & x & " pages)"
                If x < 2 Then objDoc.Tables(3).Cell(Row:=8, Column:=2).Range.Text = " 1) Report 2.1 (" & x & " page)"
            Else
                objDoc.Tables(3).Cell(Row:=8, Column:=2).Range.Text = " 1) Report 2.1 (" & y & " reports, " & "total of " & x & " pages)"
            End If
            
            If Cells(ActiveCell.Row, 126).Value <> "" And Cells(ActiveCell.Row, 126).Value > 0 Then 'Kimlik var
                x = Application.Sum(Range(Cells(IlkSira, 169), Cells(SonSira, 169))) + Cells(ActiveCell.Row, 126).Value 'Rapor3 Tutanağı
                objDoc.Tables(3).Cell(Row:=9, Column:=2).Range.Text = " 2) Attached Statement (total of " & x & " pages)"
            Else 'Kimlik yok
                x = Application.Sum(Range(Cells(IlkSira, 169), Cells(SonSira, 169))) 'Rapor3 Tutanağı
                If x > 1 Then objDoc.Tables(3).Cell(Row:=9, Column:=2).Range.Text = " 2) Statement (" & x & " pages)"
                If x < 2 Then objDoc.Tables(3).Cell(Row:=9, Column:=2).Range.Text = " 2) Statement (" & x & " page)"
            End If
            
        End If
    Else 'Organization B mensubu yok.
        If Mid(Cells(IlkSira, 128).Value, 7, 2) = "31" Then 'TipAyı rızası ile verdi.

            Bolum2 = Ek4 & " dated " & Cells(IlkSira, 218).Value & " and numbered " & Ek3 & ", concerning the Type A " & Ek1 & ", " & Ek5 & " hereby attached to be delivered to the relevant Process Monitoring Directorate along with the " & Ek1 & "."

            'Kişinin açık kimliği var/yok
            If Cells(ActiveCell.Row, 126).Value <> "" And Cells(ActiveCell.Row, 126).Value > 0 Then
                Bolum1 = "The Type A " & Ek1 & " delivered to our unit on " & Ek8 & " by the person named " & Ek9 & ", whose full identity is stated in the attached statement, " & _
                        Ek7 & " been evaluated as invalid. The item type, item value, quantity, and item ID number are also specified in the same statement, " & _
                        "and the " & Ek1 & " have been assigned the theme number " & Ek6 & " by our office."
            Else
                Bolum1 = "The Type A " & Ek1 & " delivered to our unit on " & Ek8 & " by the person named " & Ek9 & _
                        " " & Ek7 & " been evaluated as invalid. The item type, item value, quantity, and item ID number are also specified in the same statement, " & _
                        "and the " & Ek1 & " have been assigned the theme number " & Ek6 & " by our office."
            End If
            
            Bolum5 = Chr(9) & "Respectfully submitted for your information."
            
            If StrTeknik_ANotu <> "" Then
                Birlestir = Chr(9) & Bolum1 & vbNewLine & Chr(9) & Bolum2 & vbNewLine & Chr(9) & StrTeknik_ANotu
            Else
                Birlestir = Chr(9) & Bolum1 & vbNewLine & Chr(9) & Bolum2
            End If
            objDoc.Tables(2).Cell(Row:=1, Column:=1).Range.Text = Birlestir
            objDoc.Tables(2).Cell(Row:=3, Column:=1).Range.Text = Bolum5

            'Ekler
            objDoc.Tables(3).Cell(Row:=8, Column:=1).Range.Text = "Attachment:"

            Ek_Paket = Cells(ActiveCell.Row, 149).Value 'Kapalı Package A
            Ek_Paket = Left(Ek_Paket, InStr(Ek_Paket, "/") - 1)
            If Cells(ActiveCell.Row, 150).Value > 1 Then objDoc.Tables(3).Cell(Row:=8, Column:=2).Range.Text = " 1) Enclosed " & Ek_Paket & " (" & Cells(ActiveCell.Row, 150).Value & " pieces)"
            If Cells(ActiveCell.Row, 150).Value < 2 Then objDoc.Tables(3).Cell(Row:=8, Column:=2).Range.Text = " 1) Enclosed " & Ek_Paket & " (" & Cells(ActiveCell.Row, 150).Value & " piece)"
            
            x = Application.Sum(Range(Cells(IlkSira, 171), Cells(SonSira, 171))) 'Statement 2 toplam sayfa sayısı
            If x > 1 Then objDoc.Tables(3).Cell(Row:=9, Column:=2).Range.Text = " 2) Statement 2 (" & x & " pages)"
            If x < 2 Then objDoc.Tables(3).Cell(Row:=9, Column:=2).Range.Text = " 2) Statement 2 (" & x & " page)"
            
            x = Application.Sum(Range(Cells(IlkSira, 174), Cells(SonSira, 174)))  'Rapor1 toplam sayfa sayısı
            If y = 1 Then
                If x > 1 Then objDoc.Tables(3).Cell(Row:=10, Column:=2).Range.Text = " 3) Report 2.1 (" & x & " pages)"
                If x < 2 Then objDoc.Tables(3).Cell(Row:=10, Column:=2).Range.Text = " 3) Report 2.1 (" & x & " page)"
            Else
                objDoc.Tables(3).Cell(Row:=10, Column:=2).Range.Text = " 3) Report 2.1 (" & y & " reports, " & "total of " & x & " pages)"
            End If
            
            If Cells(ActiveCell.Row, 126).Value <> "" And Cells(ActiveCell.Row, 126).Value > 0 Then 'Kimlik var
                x = Application.Sum(Range(Cells(IlkSira, 169), Cells(SonSira, 169))) + Cells(ActiveCell.Row, 126).Value 'Rapor3 Tutanağı
                objDoc.Tables(3).Cell(Row:=11, Column:=2).Range.Text = " 4) Attached Statement (total of " & x & " pages)"
            Else 'Kimlik yok
                x = Application.Sum(Range(Cells(IlkSira, 169), Cells(SonSira, 169))) 'Rapor3 Tutanağı
                If x > 1 Then objDoc.Tables(3).Cell(Row:=11, Column:=2).Range.Text = " 4) Statement (" & x & " pages)"
                If x < 2 Then objDoc.Tables(3).Cell(Row:=11, Column:=2).Range.Text = " 4) Statement (" & x & " page)"
            End If
        
        ElseIf Mid(Cells(IlkSira, 128).Value, 10, 2) = "41" Then 'TipAyı rızası ile vermedi.

            'Kişinin açık kimliği var/yok
            If Cells(ActiveCell.Row, 126).Value <> "" And Cells(ActiveCell.Row, 126).Value > 0 Then
                Bolum1 = "The Type A " & Ek1 & " delivered to our unit on " & Ek8 & " by the person named " & Ek9 & ", whose full identity is stated in the attached statement, " & _
                        Ek7 & " been evaluated as invalid. The item type, item value, quantity, and item ID number are also specified in the same statement. " & _
                        "Nevertheless, although the necessary reminder was provided, the individual who delivered the Type A " & Ek1 & " evaluated as invalid refused to hand them over voluntarily."
            Else
                Bolum1 = "The Type A " & Ek1 & " delivered to our unit on " & Ek8 & " by the person named " & Ek9 & _
                        Ek7 & " been evaluated as invalid. The item type, item value, quantity, and item ID number are also specified in the same statement. " & _
                     "Nevertheless, although the necessary reminder was provided, the individual who delivered the Type A " & Ek1 & " evaluated as invalid refused to hand them over voluntarily."
            End If
            
            Bolum5 = Chr(9) & "Respectfully submitted for your information."
            
            If StrTeknik_ANotu <> "" Then
                Birlestir = Chr(9) & Bolum1 & vbNewLine & Chr(9) & StrTeknik_ANotu
            Else
                Birlestir = Chr(9) & Bolum1
            End If
            objDoc.Tables(2).Cell(Row:=1, Column:=1).Range.Text = Birlestir
            objDoc.Tables(2).Cell(Row:=3, Column:=1).Range.Text = Bolum5

            'Ekler
            objDoc.Tables(3).Cell(Row:=8, Column:=1).Range.Text = "Attachment:"
            If Cells(ActiveCell.Row, 126).Value <> "" And Cells(ActiveCell.Row, 126).Value > 0 Then 'Kimlik var
                x = Application.Sum(Range(Cells(IlkSira, 169), Cells(SonSira, 169))) + Cells(ActiveCell.Row, 126).Value 'Rapor3 Tutanağı
                objDoc.Tables(3).Cell(Row:=8, Column:=2).Range.Text = " Attached Statement (total of " & x & " pages)"
            Else 'Kimlik yok
                x = Application.Sum(Range(Cells(IlkSira, 169), Cells(SonSira, 169))) 'Rapor3 Tutanağı
                If x > 1 Then objDoc.Tables(3).Cell(Row:=8, Column:=2).Range.Text = " Statement (" & x & " pages)"
                If x < 2 Then objDoc.Tables(3).Cell(Row:=8, Column:=2).Range.Text = " Statement (" & x & " page)"
            End If
            
        End If
    End If

    ''''
    XXXMudNotu = False
    'Üst yazı notu (XXXMud notu)
    If Cells(ActiveCell.Row, 215).Value = "Yes" Then 'XXXMudNotu = True
        XXXMudNotu = True
    Else
        objDoc.Tables(2).Cell(Row:=2, Column:=1).Range.Text = ""
        'objDoc.Tables(2).Rows(2).Delete
    End If
    

    '__________________________DİNAMİK SAYFA YAPISI
    
    'İlgi ve gövde metninin kaç satırdan oluştuğu bilgisinin elde edilmesinde kullanılıyor.
    'Set IlgiRng = objDoc.Tables(1).Cell(Row:=15, Column:=3).Range
    Set Govde1Rng = objDoc.Tables(2).Cell(Row:=1, Column:=1).Range
    Set Govde2Rng = objDoc.Tables(2).Cell(Row:=3, Column:=1).Range
    
    'IlgiStrSay = Govde1Rng.Information(wdFirstCharacterLineNumber) - IlgiRng.Information(wdFirstCharacterLineNumber) - 1
    Govde1StrSay = Govde2Rng.Information(wdFirstCharacterLineNumber) - Govde1Rng.Information(wdFirstCharacterLineNumber)
    IlgiFarkSay = 0
    Govde1FarkSay = Govde1StrSay - 2
    CokluSayfa = 0
    'MsgBox Govde1FarkSay

    'Dinamik sayfa düzeni
    If StrTeknik_ANotu <> "" Then '
        If XXXMudNotu = True Then
            'yazı tipini küçült
            Set Govde1Rng = objDoc.Tables(2).Cell(Row:=1, Column:=1).Range
            Set Govde2Rng = objDoc.Tables(2).Cell(Row:=2, Column:=1).Range

            objDoc.Range.Font.Size = 11
            objDoc.Tables(1).Cell(Row:=5, Column:=1).Range.Font.Size = 13
            objDoc.Tables(1).Cell(Row:=2, Column:=1).Range.Font.Size = 9
            
            If Mid(Cells(IlkSira, 128).Value, 7, 2) = "31" Then 'TipAyı rızası ile verdi. '4 Ek var
                For i = 1 To 1
                    objDoc.Tables(3).Rows(12).Delete 'Ek sonrası
                Next i
                ' ID photocopy
                If Cells(ActiveCell.Row, 126).Value <> "" Then
                    Govde1FarkSay = Govde1FarkSay - 2
                Else
                    Govde1FarkSay = Govde1FarkSay - 4
                End If
            ElseIf Mid(Cells(IlkSira, 128).Value, 13, 2) = "51" Then 'tipAya Organization B mensubu el koydu. '2 Ek var
                For i = 1 To 3
                    objDoc.Tables(3).Rows(10).Delete 'Ek sonrası
                Next i
                Govde1FarkSay = Govde1FarkSay - 4
            Else 'Tek ek var
                For i = 1 To 4
                    objDoc.Tables(3).Rows(9).Delete 'Ek sonrası
                Next i
                Govde1FarkSay = Govde1FarkSay - 5
            End If
        Else
            objDoc.Tables(2).Rows(2).Delete 'Boş satırı sil
            If Mid(Cells(IlkSira, 128).Value, 7, 2) = "31" Then 'TipAyı rızası ile verdi. '4 Ek var
                For i = 1 To 1
                    objDoc.Tables(3).Rows(12).Delete 'Ek sonrası
                Next i
                Govde1FarkSay = Govde1FarkSay - 1
            ElseIf Mid(Cells(IlkSira, 128).Value, 13, 2) = "51" Then 'tipAya Organization B mensubu el koydu. '2 Ek var
                For i = 1 To 3
                    objDoc.Tables(3).Rows(10).Delete 'Ek sonrası
                Next i
                Govde1FarkSay = Govde1FarkSay - 3
            Else 'Tek ek var
                For i = 1 To 4
                    objDoc.Tables(3).Rows(9).Delete 'Ek sonrası
                Next i
                Govde1FarkSay = Govde1FarkSay - 4
            End If
        End If
    Else
        If XXXMudNotu = True Then
            If Mid(Cells(IlkSira, 128).Value, 7, 2) = "31" Then 'TipAyı rızası ile verdi. '4 Ek var
                For i = 1 To 1
                    objDoc.Tables(3).Rows(12).Delete 'Ek sonrası
                Next i
                Govde1FarkSay = Govde1FarkSay - 0
            ElseIf Mid(Cells(IlkSira, 128).Value, 13, 2) = "51" Then 'tipAya Organization B mensubu el koydu. '2 Ek var
                For i = 1 To 3
                    objDoc.Tables(3).Rows(10).Delete 'Ek sonrası
                Next i
                Govde1FarkSay = Govde1FarkSay - 2
            Else 'Tek ek var
                For i = 1 To 4
                    objDoc.Tables(3).Rows(9).Delete 'Ek sonrası
                Next i
                Govde1FarkSay = Govde1FarkSay - 3
            End If
        Else
            objDoc.Tables(2).Rows(2).Delete 'Boş satırı sil
            If Mid(Cells(IlkSira, 128).Value, 7, 2) = "31" Then 'TipAyı rızası ile verdi. '4 Ek var
                For i = 1 To 1
                    objDoc.Tables(3).Rows(12).Delete 'Ek sonrası
                Next i
                Govde1FarkSay = Govde1FarkSay - 1
            ElseIf Mid(Cells(IlkSira, 128).Value, 13, 2) = "51" Then 'tipAya Organization B mensubu el koydu. '2 Ek var
                For i = 1 To 3
                    objDoc.Tables(3).Rows(10).Delete 'Ek sonrası
                Next i
                Govde1FarkSay = Govde1FarkSay - 3
            Else 'Tek ek var
                For i = 1 To 4
                    objDoc.Tables(3).Rows(9).Delete 'Ek sonrası
                Next i
                Govde1FarkSay = Govde1FarkSay - 4
            End If
        End If
    End If
    
    If M4 = True Then
        For i = 1 To 2
            objDoc.Tables(1).Rows(14).Delete 'Muhatap sonrası
        Next i
        
        If IlgiFarkSay + Govde1FarkSay < 17 Then
            Govde1FarkSay = Govde1FarkSay + 2
            'MsgBox IlgiFarkSay + Govde1FarkSay
            If IlgiFarkSay + Govde1FarkSay > 8 Then
                For i = 9 To IlgiFarkSay + Govde1FarkSay
                    If i = 9 Then
                        objDoc.Tables(3).Rows(6).Delete 'Ek öncesi
                    ElseIf i = 10 Then
                        objDoc.Tables(3).Rows(6).Delete 'Ek öncesi
                    ElseIf i = 11 Then
                        objDoc.Tables(3).Rows(1).Delete 'İmza öncesi
                    ElseIf i = 12 Then
                        objDoc.Tables(3).Rows(1).Delete 'İmza öncesi
                    ElseIf i = 13 Then
                        objDoc.Tables(1).Rows(8).Delete 'Muhatap öncesi
                    ElseIf i = 14 Then
                        objDoc.Tables(1).Rows(8).Delete 'Muhatap öncesi
                    End If
                Next i
            Else
                If (IlgiFarkSay + Govde1FarkSay) < 8 Then
                    For i = 1 To 8 - (IlgiFarkSay + Govde1FarkSay)
                        'Boş satır ekle
                        With objDoc.Tables(3)
                            .Rows.Add BeforeRow:=objDoc.Tables(3).Rows(8)
                        End With
                    Next i
                End If
            End If
        Else
            CokluSayfa = 1
        End If
    ElseIf M3 = True Then
        For i = 1 To 3
            objDoc.Tables(1).Rows(13).Delete 'Muhatap sonrası
        Next i
        
        If IlgiFarkSay + Govde1FarkSay < 16 Then
            Govde1FarkSay = Govde1FarkSay + 1
            'MsgBox IlgiFarkSay + Govde1FarkSay
            If IlgiFarkSay + Govde1FarkSay > 8 Then
                For i = 9 To IlgiFarkSay + Govde1FarkSay
                    If i = 9 Then
                        objDoc.Tables(3).Rows(6).Delete 'Ek öncesi
                    ElseIf i = 10 Then
                        objDoc.Tables(3).Rows(6).Delete 'Ek öncesi
                    ElseIf i = 11 Then
                        objDoc.Tables(3).Rows(1).Delete 'İmza öncesi
                    ElseIf i = 12 Then
                        objDoc.Tables(3).Rows(1).Delete 'İmza öncesi
                    ElseIf i = 13 Then
                        objDoc.Tables(1).Rows(8).Delete 'Muhatap öncesi
                    ElseIf i = 14 Then
                        objDoc.Tables(1).Rows(8).Delete 'Muhatap öncesi
                    End If
                Next i
            Else
                If (IlgiFarkSay + Govde1FarkSay) < 8 Then
                    For i = 1 To 8 - (IlgiFarkSay + Govde1FarkSay)
                        'Boş satır ekle
                        With objDoc.Tables(3)
                            .Rows.Add BeforeRow:=objDoc.Tables(3).Rows(8)
                        End With
                    Next i
                End If
            End If
        Else
            CokluSayfa = 1
        End If
    ElseIf M2 = True Then
        For i = 1 To 4
            objDoc.Tables(1).Rows(12).Delete 'Muhatap sonrası
        Next i
        'Govde1FarkSay = Govde1FarkSay + 0
        'MsgBox IlgiFarkSay + Govde1FarkSay
        If IlgiFarkSay + Govde1FarkSay < 15 Then
            If IlgiFarkSay + Govde1FarkSay > 8 Then
                For i = 9 To IlgiFarkSay + Govde1FarkSay
                    If i = 9 Then
                        objDoc.Tables(3).Rows(6).Delete 'Ek öncesi
                    ElseIf i = 10 Then
                        objDoc.Tables(3).Rows(6).Delete 'Ek öncesi
                    ElseIf i = 11 Then
                        objDoc.Tables(3).Rows(1).Delete 'İmza öncesi
                    ElseIf i = 12 Then
                        objDoc.Tables(3).Rows(1).Delete 'İmza öncesi
                    ElseIf i = 13 Then
                        objDoc.Tables(1).Rows(8).Delete 'Muhatap öncesi
                    ElseIf i = 14 Then
                        objDoc.Tables(1).Rows(8).Delete 'Muhatap öncesi
                    End If
                Next i
            Else
                If (IlgiFarkSay + Govde1FarkSay) < 8 Then
                    For i = 1 To 8 - (IlgiFarkSay + Govde1FarkSay)
                        'Boş satır ekle
                        With objDoc.Tables(3)
                            .Rows.Add BeforeRow:=objDoc.Tables(3).Rows(8)
                        End With
                    Next i
                End If
            End If
        Else
            CokluSayfa = 1
        End If
    End If
    
    objDoc.Paragraphs(objDoc.Paragraphs.Count).Range.Font.Size = 1

    
    'Footer
    If objDoc.ActiveWindow.View.SplitSpecial <> wdPaneNone Then
        objDoc.ActiveWindow.Panes(2).Close
    End If
    If objDoc.ActiveWindow.ActivePane.View.Type = wdNormalView Or objDoc.ActiveWindow.ActivePane.View.Type = wdOutlineView Then
        objDoc.ActiveWindow.ActivePane.View.Type = wdPrintView
    End If

    With objDoc
        With .Sections(1).Headers(wdHeaderFooterPrimary).Range
            .InsertAfter Text:="Page "
            .Fields.Add Range:=.Characters.Last, Type:=wdFieldEmpty, Text:="PAGE", PreserveFormatting:=False
            .InsertAfter Text:=" of the letter dated " & ustbilgitarih & _
                             ", reference number " & ustbilgisayi & _
                             ", addressed to the " & ustbilgimuhatap
        End With
    End With
    
    objDoc.ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument

    
    'PARAF EKLEME İŞLEMİ
    objDoc.Sections(1).Footers(wdHeaderFooterFirstPage).Range.Tables(1).Cell(Row:=1, Column:=2).Range.Text = AdSoyadParaf & vbNewLine & _
                                                                                       UnvanParaf & vbNewLine & _
                                                                                       TelParaf
    objDoc.Sections(1).Footers(wdHeaderFooterPrimary).Range.Tables(1).Cell(Row:=1, Column:=2).Range.Text = AdSoyadParaf & vbNewLine & _
                                                                                       UnvanParaf & vbNewLine & _
                                                                                       TelParaf
                                                                                       

Tekrarla:
    'Sayfa sayısı kaydet komutuna bağlandı.
    objDoc.Save
    'Sayfayı text dosyasından çek
    TxtFileUstYazi = DestOpUserFolder & "Cover Letter Page Count.txt"
    #If UseADO Then
        Const adReadLine = -2&
        With CreateObject("ADODB.Stream")
            .Open
            .LoadFromFile TxtFileUstYazi
            Do Until .EOS
                TotalSayfaUstYazi = .ReadText(adReadLine)
            Loop
            .Close
        End With
    #Else 'UseFSO
        Const UstYaziTristateTrue = -1&
        With CreateObject("Scripting.FileSystemObject")
            With .OpenTextFile(TxtFileUstYazi, Format:=UstYaziTristateTrue)
                Do Until .AtEndOfStream
                    TotalSayfaUstYazi = .ReadLine
                Loop
                .Close
            End With
        End With
    #End If
    'MsgBox "Üst Yazı: " & TotalSayfaUstYazi

    'Üst Yazı sayfa sayısı
    Cells(ActiveCell.Row, 173).Value = TotalSayfaUstYazi

    objWord.Visible = True
    objWord.Activate

    If TumDoc = True Then
        objWord.Activate
'        For i = 1 To SayPrt
'            objDoc.PrintOut
'        Next i
        objDoc.PrintOut Background:=False, Copies:=SayPrt
        'objWord.Documents.Save
        objDoc.Close SaveChanges:=False
        objWord.Visible = False
    End If

    'Nesneleri temizle
    Set fso = Nothing
    Set objWord = Nothing
    Set objDoc = Nothing

Son:

If IlceSakla <> "" Then
    Cells(ActiveCell.Row, 92).Value = IlceSakla
End If

'Worksheets(5).Activate
Application.ScreenUpdating = True
Application.DisplayAlerts = True
'Application.EnableEvents = True

'End If

End Sub


Sub IslemGunluguRapor3_1()
Dim AutoPath As String, IslemGunlugu As String, IslemGunlukleriKlasor As String, WsIslemGunlugu As Object
Dim Kenarlar As Range
Dim BulIslemGunlugu As Range, AralikSay As Integer, KayitDefSiraNo As Long

Dim i As Long, j As Long
Dim YeniIslem As Long, Maxi As Integer
Dim OpenControl As String
Dim WsRapor As Object

Dim IlkSira As Long, SonSira As Long
Dim IslemGunluguIlkSiraBul As Range, IslemGunluguSonSiraBul As Range, IslemGunluguIlkSira As Long, IslemGunluguSonSira As Long
Dim GelenTema As String, Fark As Long, DelControl As Boolean



Application.DisplayAlerts = False
Application.EnableEvents = False
Application.ScreenUpdating = False

'Pathfinder...
AutoPath = ThisWorkbook.Path
IslemGunlukleriKlasor = AutoPath & "\System Files\System Templates\Registry Reports\"
IslemGunlugu = IslemGunlukleriKlasor & "System Registry Report 2.1.xlsx"

'Check the "Registry Reports" folder name.
If Not Dir(IslemGunlukleriKlasor, vbDirectory) <> vbNullString Then
    MsgBox "Cannot access the directory " & IslemGunlukleriKlasor & ". The folder named 'Registry Reports' may have been renamed or deleted.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Out
End If

If Not Dir(IslemGunlugu, vbDirectory) <> vbNullString Then
    MsgBox "Cannot access the directory " & IslemGunlugu & ". The names of folders and/or files in this directory may have been changed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Out
End If


Maxi = MaxiAktar
YeniIslem = YeniIslemAktar
'StrTime = Format(Now, "ddmmyyyyhhmmss")
DelControl = False

'Modülün Rapor sayfasında bulunan başlangıç ve bitiş satır numaraları
IlkSira = YeniIslem
SonSira = YeniIslem + Maxi
Set WsRapor = ThisWorkbook.Worksheets(5)
'WsRapor.Unprotect Password:="123"

IlceSakla = ""
If InStr(WsRapor.Cells(IlkSira, 92).Value, " Organization A") <> 0 Then
    IlceSakla = WsRapor.Cells(IlkSira, 92).Value
    WsRapor.Cells(IlkSira, 92).Value = ""
End If

'Aylık ayraçlar
If WsRapor.Cells(IlkSira, 95).Value <> "" Then
    ModulTarih = WsRapor.Cells(IlkSira, 95).Value
    ModulAyrac = "01" & Right(ModulTarih, 8)
Else 'işlemin yapıldığı günü esas al
    ModulTarih = Format(Date, "dd.mm.yyyy")
    ModulAyrac = "01" & Right(ModulTarih, 8)
End If


'İşlem günlüğü açıksa kaydet ve kapat.
OpenControl = IsWorkBookOpen(IslemGunlugu)
If OpenControl = True Then
    Workbooks("System Registry Report 2.1.xlsx").Save
    Workbooks("System Registry Report 2.1.xlsx").Close SaveChanges:=False
End If

'İşlem günlüğü aç
Workbooks.Open (IslemGunlugu)
Set WsIslemGunlugu = Workbooks("System Registry Report 2.1.xlsx").Worksheets(1)

WsIslemGunlugu.Unprotect Password:="123"

WsIslemGunlugu.Columns("B:C").EntireColumn.Hidden = False


'İşlem günlüğünde yoksa ve işlem tipA değilse prosedürden çık (TipA/TipB)
Set IslemGunluguIlkSiraBul = WsIslemGunlugu.Range("B7:B100000").Find(What:=WsRapor.Cells(IlkSira, 165).Value, SearchDirection:=xlNext, _
                SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole) 'zaman damgasını ara
If Not IslemGunluguIlkSiraBul Is Nothing Then
    'Nothing
Else
    If WsRapor.Cells(IlkSira, 100).Value <> "Type A" Then
        GoTo Son
    End If
End If


'____________HAZIRLIK

'İşlem günlüğünde başlangıç ve bitiş satırlarını tespit et.
Say1IslemGunlugu = WsIslemGunlugu.Range("B100000").End(xlUp).Row
Say2IslemGunlugu = WsIslemGunlugu.Range("C100000").End(xlUp).Row
SayAyracIslemGunlugu = WsIslemGunlugu.Range("E100000").End(xlUp).Row
'İşlem günlüğünde ayraçları oluştur
If Say1IslemGunlugu < 7 And SayAyracIslemGunlugu < 7 Then

    Say1IslemGunlugu = 6
    Say2IslemGunlugu = 6
    SayAyracIslemGunlugu = 6

    i = 6
    IslemGunluguAyrac = "01.01" & Right(ModulAyrac, 5)
    ModulAyrac = DateAdd("m", 1, CDate(ModulAyrac))
    'MsgBox IslemGunluguAyrac & " and " & ModulAyrac '''
    Do Until CDate(IslemGunluguAyrac) >= CDate(ModulAyrac)
        i = i + 1
        With WsIslemGunlugu.Range("E" & i & ":T" & i)
            .WrapText = False
            .Font.Bold = True
            .Font.name = "Open Sans"
            .Font.Size = 9
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = RGB(174, 185, 194) 'RGB(201, 216, 230) 'RGB(174, 185, 194)
            .HorizontalAlignment = xlLeft
            .NumberFormat = "mmmm yyyy"
        End With
        WsIslemGunlugu.Cells(i, 5).Value = CDate(IslemGunluguAyrac)
        IslemGunluguAyrac = DateAdd("m", 1, CDate(IslemGunluguAyrac))
    Loop
    ModulAyrac = DateAdd("m", -1, CDate(ModulAyrac))

ElseIf Say1IslemGunlugu < 7 And SayAyracIslemGunlugu >= 7 Then

    i = SayAyracIslemGunlugu
    IslemGunluguAyrac = WsIslemGunlugu.Cells(SayAyracIslemGunlugu, 5).Value
    'MsgBox IslemGunluguAyrac & " and " & ModulAyrac '''
    Do Until CDate(IslemGunluguAyrac) >= CDate(ModulAyrac)
        i = i + 1
        With WsIslemGunlugu.Range("E" & i & ":T" & i)
            .WrapText = False
            .Font.Bold = True
            .Font.name = "Open Sans"
            .Font.Size = 9
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = RGB(174, 185, 194) 'RGB(201, 216, 230) 'RGB(174, 185, 194)
            .HorizontalAlignment = xlLeft
            .NumberFormat = "mmmm yyyy"
        End With
        IslemGunluguAyrac = DateAdd("m", 1, CDate(IslemGunluguAyrac))
        WsIslemGunlugu.Cells(i, 5).Value = CDate(IslemGunluguAyrac)
    Loop

ElseIf Say1IslemGunlugu >= 7 And SayAyracIslemGunlugu >= 7 Then

    SayMax = WorksheetFunction.Max(Say2IslemGunlugu, SayAyracIslemGunlugu)
    i = SayMax
    IslemGunluguAyrac = WsIslemGunlugu.Cells(SayAyracIslemGunlugu, 5).Value
    'MsgBox IslemGunluguAyrac & " and " & ModulAyrac '''
    Do Until CDate(IslemGunluguAyrac) >= CDate(ModulAyrac)
        i = i + 1
        With WsIslemGunlugu.Range("E" & i & ":T" & i)
            .WrapText = False
            .Font.Bold = True
            .Font.name = "Open Sans"
            .Font.Size = 9
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = RGB(174, 185, 194) 'RGB(201, 216, 230) 'RGB(174, 185, 194)
            .HorizontalAlignment = xlLeft
            .NumberFormat = "mmmm yyyy"
        End With
        IslemGunluguAyrac = DateAdd("m", 1, CDate(IslemGunluguAyrac))
        WsIslemGunlugu.Cells(i, 5).Value = CDate(IslemGunluguAyrac)
    Loop

End If

'MsgBox "IlkSira: " & IlkSira & " ve SonSira: " & SonSira
'GoTo Son

'GELEN TEMA
GelenTema = ""
If WsRapor.Cells(IlkSira, 102).Value = "Provincial Directorate B" Then
    If WsRapor.Cells(IlkSira, 103).Value <> "" Then
        GelenTema = WsRapor.Cells(IlkSira, 91).Value & " Provincial Governorship Provincial Directorate B " & WsRapor.Cells(IlkSira, 103).Value
    Else
        GelenTema = WsRapor.Cells(IlkSira, 91).Value & " Provincial Governorship Provincial Directorate B"
    End If
ElseIf WsRapor.Cells(IlkSira, 102).Value = "District Directorate B" Then
    If WsRapor.Cells(IlkSira, 103).Value <> "" Then
        GelenTema = WsRapor.Cells(IlkSira, 92).Value & " District Governorship District Directorate B " & WsRapor.Cells(IlkSira, 103).Value
    Else
        GelenTema = WsRapor.Cells(IlkSira, 92).Value & " District Governorship District Directorate B"
    End If
ElseIf WsRapor.Cells(IlkSira, 102).Value = "Provincial Directorate C" Then
    If WsRapor.Cells(IlkSira, 103).Value <> "" Then
        GelenTema = WsRapor.Cells(IlkSira, 91).Value & " Provincial Governorship Provincial Directorate C " & WsRapor.Cells(IlkSira, 103).Value
    Else
        GelenTema = WsRapor.Cells(IlkSira, 91).Value & " Provincial Governorship Provincial Directorate C"
    End If
ElseIf WsRapor.Cells(IlkSira, 102).Value = "District Directorate C" Then
    If WsRapor.Cells(IlkSira, 103).Value <> "" Then
        GelenTema = WsRapor.Cells(IlkSira, 92).Value & " District Governorship District Directorate C " & WsRapor.Cells(IlkSira, 103).Value
    Else
        GelenTema = WsRapor.Cells(IlkSira, 92).Value & " District Governorship District Directorate C"
    End If
ElseIf WsRapor.Cells(IlkSira, 102).Value = "Provincial Directorate D" Then
    If WsRapor.Cells(IlkSira, 103).Value <> "" Then
        GelenTema = WsRapor.Cells(IlkSira, 91).Value & " Provincial Governorship Provincial Directorate D " & WsRapor.Cells(IlkSira, 103).Value
    Else
        GelenTema = WsRapor.Cells(IlkSira, 91).Value & " Provincial Governorship Provincial Directorate D"
    End If
ElseIf WsRapor.Cells(IlkSira, 102).Value = "District Directorate D" Then
    If WsRapor.Cells(IlkSira, 103).Value <> "" Then
        GelenTema = WsRapor.Cells(IlkSira, 92).Value & " District Governorship District Directorate D " & WsRapor.Cells(IlkSira, 103).Value
    Else
        GelenTema = WsRapor.Cells(IlkSira, 92).Value & " District Governorship District Directorate D"
    End If
ElseIf WsRapor.Cells(IlkSira, 102).Value = "Provincial Directorate E" Then
    If WsRapor.Cells(IlkSira, 103).Value <> "" Then
        GelenTema = WsRapor.Cells(IlkSira, 91).Value & " Provincial Governorship Provincial Directorate E " & WsRapor.Cells(IlkSira, 103).Value
    Else
        GelenTema = WsRapor.Cells(IlkSira, 91).Value & " Provincial Governorship Provincial Directorate E"
    End If
ElseIf WsRapor.Cells(IlkSira, 102).Value = "District Directorate E" Then
    If WsRapor.Cells(IlkSira, 103).Value <> "" Then
        GelenTema = WsRapor.Cells(IlkSira, 92).Value & " District Governorship District Directorate E " & WsRapor.Cells(IlkSira, 103).Value
    Else
        GelenTema = WsRapor.Cells(IlkSira, 92).Value & " District Governorship District Directorate E"
    End If
ElseIf InStr(WsRapor.Cells(IlkSira, 102).Value, "General Directorate") <> 0 Or InStr(WsRapor.Cells(IlkSira, 102).Value, "Regional Directorate") <> 0 Then
    If WsRapor.Cells(IlkSira, 103).Value <> "" Then
        GelenTema = WorksheetFunction.Proper(ThisWorkbook.Worksheets(2).Cells(6, 111).Value) & " " & WsRapor.Cells(IlkSira, 102).Value & " " & WsRapor.Cells(IlkSira, 103).Value
    Else
        GelenTema = WorksheetFunction.Proper(ThisWorkbook.Worksheets(2).Cells(6, 111).Value) & " " & WsRapor.Cells(IlkSira, 102).Value
    End If
Else
    If WsRapor.Cells(IlkSira, 103).Value <> "" Then
        If InStr(WsRapor.Cells(IlkSira, 102).Value, "X.X. ") > 0 Then
            GelenTema = Mid(WsRapor.Cells(IlkSira, 102).Value, 6, Len(WsRapor.Cells(IlkSira, 102).Value)) & " " & WsRapor.Cells(IlkSira, 103).Value
        Else
            GelenTema = WsRapor.Cells(IlkSira, 102).Value & " " & WsRapor.Cells(IlkSira, 103).Value
        End If
    Else
        If InStr(WsRapor.Cells(IlkSira, 102).Value, "X.X. ") > 0 Then
            GelenTema = Mid(WsRapor.Cells(IlkSira, 102).Value, 6, Len(WsRapor.Cells(IlkSira, 102).Value))
        Else
            GelenTema = WsRapor.Cells(IlkSira, 102).Value
        End If
    End If
End If

'____________OPERASYONLAR

'İşlem günlüğünde başlangıç ve bitiş satırlarını tespit et.
Say1IslemGunlugu = WsIslemGunlugu.Range("B100000").End(xlUp).Row
Say2IslemGunlugu = WsIslemGunlugu.Range("C100000").End(xlUp).Row
SayAyracIslemGunlugu = WsIslemGunlugu.Range("E100000").End(xlUp).Row

Set IslemGunluguIlkSiraBul = WsIslemGunlugu.Range("B7:B100000").Find(What:=WsRapor.Cells(IlkSira, 165).Value, SearchDirection:=xlNext, _
                SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole) 'zaman damgasını ara
Set IslemGunluguSonSiraBul = WsIslemGunlugu.Range("C7:C100000").Find(What:=WsRapor.Cells(IlkSira, 165).Value, SearchDirection:=xlNext, _
                SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
If Not IslemGunluguIlkSiraBul Is Nothing Then
    IslemGunluguIlkSira = IslemGunluguIlkSiraBul.Row
    If Not IslemGunluguSonSiraBul Is Nothing Then
        IslemGunluguSonSira = IslemGunluguSonSiraBul.Row
    End If
End If

    
If Not IslemGunluguIlkSiraBul Is Nothing Then 'DÜZENLEME İŞLEMİ

    '_______________'TipA/TipB (Daha önce tipA olarak kaydettiği bir işlemi tipBye çevirirse işlem günlüğündeki kaydın silinmesi gerekir.)
    
    If WsRapor.Cells(IlkSira, 100).Value <> "Type A" Then
        
        'kayıt def. verileri sil, satırları işaretle
        WsIslemGunlugu.Range(WsIslemGunlugu.Cells(IslemGunluguIlkSira, 2), WsIslemGunlugu.Cells(IslemGunluguSonSira, 20)).ClearContents
        WsIslemGunlugu.Cells(IslemGunluguIlkSira, 2).Value = "Sil" 'ilk satırı silmek üzere işaretle
        WsIslemGunlugu.Cells(IslemGunluguSonSira, 3).Value = "Sil" 'son satırı silmek üzere işaretle
        
        'Dönem sıra no.ları güncelle
        i = IslemGunluguSonSira
        Do Until WsIslemGunlugu.Cells(i, 5).Value <> "" 'silinecek verinin dönemi en alt satırda değilse stop koşulu
            i = i + 1
            If i > Say2IslemGunlugu Then 'silinecek verinin dönemi en alt satırda ise stop koşulu
                GoTo SilDonemSiraNo
            End If
            If WsIslemGunlugu.Cells(i, 6).Value <> "" And IsNumeric(WsIslemGunlugu.Cells(i, 6).Value) Then 'silinen veriden sonraki verileri dönem sıra no.ları 1 azalır
                WsIslemGunlugu.Cells(i, 6).Value = WsIslemGunlugu.Cells(i, 6).Value - 1
            End If
        Loop
SilDonemSiraNo:
    
        'Genel sıra no.ları güncelle
        SayGenel = WsIslemGunlugu.Range("D100000").End(xlUp).Row
        i = IslemGunluguSonSira
        Do Until i > SayGenel
            i = i + 1
            If WsIslemGunlugu.Cells(i, 4).Value <> "" Then
                WsIslemGunlugu.Cells(i, 4).Value = WsIslemGunlugu.Cells(i, 4).Value - 1
            End If
        Loop
        
        Say2IslemGunlugu = WsIslemGunlugu.Range("C100000").End(xlUp).Row
        If IslemGunluguIlkSira > 8 Then 'And IslemGunluguIlkSira < Say2IslemGunlugu Then
            Set Kenarlar = WsIslemGunlugu.Range("D" & IslemGunluguIlkSira - 1 & ":T" & IslemGunluguIlkSira - 1)
            With Kenarlar.Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .Color = RGB(174, 185, 194)
                .Weight = xlThin
            End With
            With Kenarlar.Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .Color = RGB(174, 185, 194)
                .Weight = xlThin
            End With
            With Kenarlar.Borders(xlInsideVertical)
                .LineStyle = xlContinuous
                .Color = RGB(174, 185, 194)
                .Weight = xlThin
            End With
            With Kenarlar.Borders(xlInsideHorizontal)
                .LineStyle = xlContinuous
                .Color = RGB(174, 185, 194)
                .Weight = xlThin
            End With
        End If
    
        'Silinecek dönemde yer alan boş satır aralığını kaldır
        Set BulIslemGunlugu = WsIslemGunlugu.Range("B:B").Find(What:="Sil", SearchDirection:=xlNext, _
        SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
        If Not BulIslemGunlugu Is Nothing Then
            ilkrowx = BulIslemGunlugu.Row
            Set BulIslemGunlugu = WsIslemGunlugu.Range("C:C").Find(What:="Sil", SearchDirection:=xlNext, _
            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not BulIslemGunlugu Is Nothing Then
                sonrowx = BulIslemGunlugu.Row
            End If
            WsIslemGunlugu.Rows(ilkrowx & ":" & sonrowx).EntireRow.Delete
        End If

        
        'İşlem günlüğünde aşağı git
        Say2IslemGunlugu = WsIslemGunlugu.Range("C100000").End(xlUp).Row
        On Error Resume Next
        ActiveWindow.ScrollRow = Say2IslemGunlugu - 10
        On Error GoTo 0
    
        WsIslemGunlugu.Columns("B:C").EntireColumn.Hidden = True
    
        WsIslemGunlugu.Protect Password:="123"
    
        'İşlem günlüğü açıksa kaydet ve kapat.
        OpenControl = IsWorkBookOpen(IslemGunlugu)
        If OpenControl = True Then
            Workbooks("System Registry Report 2.1.xlsx").Save
            Workbooks("System Registry Report 2.1.xlsx").Close SaveChanges:=False
        End If

        If IlceSakla <> "" Then
            WsRapor.Cells(IlkSira, 92).Value = IlceSakla
        End If

        GoTo Out
        
    End If

    '_________________________'TipA/TipB BİTİŞ
    
    
    'DÖNEM AYNI mı FARKLI mı?
    i = IslemGunluguIlkSiraBul.Row
    Do Until WsIslemGunlugu.Cells(i, 5).Value <> "" 'CDate(ModulAyrac)
        i = i - 1
    Loop
    
    If WsIslemGunlugu.Cells(i, 5).Formula = CDate(ModulAyrac) Then
        'MsgBox "Aynı dönem"
        GoTo DonemAyni
    Else
        'MsgBox "Farklı dönem"
        GoTo DonemFarkli
    End If

DonemFarkli:
    '_______________FARKLI DÖNEM (DÜZENLEME İŞLEMİ)

    'Önceki dönemde bulunan veriyi sil.
    WsIslemGunlugu.Range(WsIslemGunlugu.Cells(IslemGunluguIlkSira, 2), WsIslemGunlugu.Cells(IslemGunluguSonSira, 20)).ClearContents
    WsIslemGunlugu.Cells(IslemGunluguIlkSira, 2).Value = "Sil" 'ilk satırı silmek üzere işaretle
    WsIslemGunlugu.Cells(IslemGunluguSonSira, 3).Value = "Sil" 'son satırı silmek üzere işaretle
        
    'Kaynak dönemde bulunan dönem sıra no.ları güncelle
    i = IslemGunluguSonSira
    Do Until WsIslemGunlugu.Cells(i, 5).Value <> "" 'kaynak dönem en alt satırda değilse stop koşulu
        i = i + 1
        If i > Say2IslemGunlugu Then 'kaynak dönem en alt satırda ise stop koşulu
            GoTo KaynakDonemSiraNo
        End If
        If WsIslemGunlugu.Cells(i, 6).Value <> "" And IsNumeric(WsIslemGunlugu.Cells(i, 6).Value) Then 'silinen veriden sonraki verileri dönem sıra no.ları 1 azalır
            WsIslemGunlugu.Cells(i, 6).Value = WsIslemGunlugu.Cells(i, 6).Value - 1
        End If
    Loop
KaynakDonemSiraNo:

    
    Set BulIslemGunlugu = WsIslemGunlugu.Range("E:E").Find(What:=CDate(ModulAyrac), SearchDirection:=xlNext, _
            SearchOrder:=xlByRows, LookIn:=xlFormulas, LookAt:=xlWhole)
    If Not BulIslemGunlugu Is Nothing Then 'HEDEF DÖNEMİ bul
        
        'Hedef dönemin en alt satırı
        i = BulIslemGunlugu.Row + 1
        j = 0
        Do Until WsIslemGunlugu.Cells(i, 5).Value <> "" 'Hedef dönem en alt satırda değilse stop koşulu
            i = i + 1
            If i > Say2IslemGunlugu Then 'Hedef dönem en alt satırda ise stop koşulu
                i = i - 1
                j = 1
                GoTo HedefDonemAltSatir
            End If
        Loop
HedefDonemAltSatir:
YeniDonemAltRow = i

        If j = 1 Then
            ilkrow = YeniDonemAltRow
            sonrow = YeniDonemAltRow + (SonSira - IlkSira)
        Else
            For i = 1 To (SonSira - IlkSira) + 1 'Taşınacak satır aralığı kadar yeni dönemin en altına satır ekle
                WsIslemGunlugu.Rows(YeniDonemAltRow).EntireRow.Insert Shift:=xlUp
            Next i

            ilkrow = YeniDonemAltRow
            sonrow = YeniDonemAltRow + (SonSira - IlkSira)
            
        End If

        'Genel sıra no.ları güncelle
        WsIslemGunlugu.Cells(ilkrow, 4).Value = 1 'Genel sıra no.sunu 1 olarak işaretle (aşağıda, doğru no. ile değiştirilecek)
        SayGenel = WsIslemGunlugu.Range("D100000").End(xlUp).Row
        i = 6
        j = 0
        Do Until i > SayGenel
            i = i + 1
            If WsIslemGunlugu.Cells(i, 4).Value <> "" Then
                j = j + 1
                WsIslemGunlugu.Cells(i, 4).Value = j
            End If
        Loop
  
        'Hedef dönem sıra no.ları güncelle
        i = ilkrow
        Do Until WsIslemGunlugu.Cells(i, 5).Value <> ""
            i = i - 1
            If WsIslemGunlugu.Cells(i, 5).Value <> "" And Not IsNumeric(WsIslemGunlugu.Cells(i, 5).Value) Then
                WsIslemGunlugu.Cells(ilkrow, 6).Value = 1
                GoTo HedefDonemSiraNo
            End If
            If WsIslemGunlugu.Cells(i, 6).Value <> "" And IsNumeric(WsIslemGunlugu.Cells(i, 6).Value) Then
                WsIslemGunlugu.Cells(ilkrow, 6).Value = WsIslemGunlugu.Cells(i, 6).Value + 1
                GoTo HedefDonemSiraNo
            End If
        Loop
HedefDonemSiraNo:
        
        
        'Kaynak dönemde yer alan boş satır aralığını kaldır
        DelControl = True

    End If
    

    GoTo DonemAyniyiAtla

DonemAyni:
    '_______________AYNI DÖNEM (DÜZENLEME İŞLEMİ)

    If Not IslemGunluguSonSiraBul Is Nothing Then
        'MsgBox "Buradayım"
        'Aktarımları yapan kodlar buraya gelecek
        WsIslemGunlugu.Range(WsIslemGunlugu.Cells(IslemGunluguIlkSira, 7), WsIslemGunlugu.Cells(IslemGunluguSonSira, 20)).ClearContents
        WsIslemGunlugu.Range(WsIslemGunlugu.Cells(IslemGunluguIlkSira, 2), WsIslemGunlugu.Cells(IslemGunluguSonSira, 3)).ClearContents
        Fark = (IslemGunluguSonSira - IslemGunluguIlkSira) - (SonSira - IlkSira)
        'MsgBox "Fark: " & Fark
        If Fark > 0 Then 'İşlem günlüğünden satır silinecek
            'MsgBox "Fark: " & Fark & " satır kaldır"
            WsIslemGunlugu.Rows(IslemGunluguSonSira - (Fark - 1) & ":" & IslemGunluguSonSira).EntireRow.Delete
            ilkrow = IslemGunluguIlkSira
            sonrow = IslemGunluguSonSira - Fark
        ElseIf Fark < 0 Then 'İşlem günlüğüne satır eklenecek
            'MsgBox "Fark: " & Fark & " satır ekle"
            Fark = -1 * Fark
            For i = 1 To Fark
                WsIslemGunlugu.Rows(IslemGunluguSonSira + 1).EntireRow.Insert Shift:=xlUp
            Next i
            ilkrow = IslemGunluguIlkSira
            sonrow = IslemGunluguSonSira + Fark
        ElseIf Fark = 0 Then 'İşlem günlüğünde satırlarda değişiklik olmayacak
            'MsgBox "Fark: " & Fark & " değişiklik yok"
            ilkrow = IslemGunluguIlkSira
            sonrow = IslemGunluguSonSira
        End If

    End If

Else 'YENİ İŞLEM

    'MsgBox CDate(ModulAyrac)
    Set BulIslemGunlugu = WsIslemGunlugu.Range("E:E").Find(What:=CDate(ModulAyrac), SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlFormulas, LookAt:=xlWhole)
    If Not BulIslemGunlugu Is Nothing Then
        donemrow = BulIslemGunlugu.Row
        If SayAyracIslemGunlugu = BulIslemGunlugu.Row Then 'Cari dönemin verisi
            If Say2IslemGunlugu > SayAyracIslemGunlugu Then 'Yeni veriyi Say2IslemGunlugu+1'e yaz
                'Genel sıra no
                SayGenel = WsIslemGunlugu.Range("D100000").End(xlUp).Row
                If SayGenel < 7 Then
                    WsIslemGunlugu.Cells(Say2IslemGunlugu + 1, 4).Value = 1
                Else
                    i = Say2IslemGunlugu + 1
                    Do Until WsIslemGunlugu.Cells(i, 4).Value <> ""
                        i = i - 1
                    Loop
                    WsIslemGunlugu.Cells(Say2IslemGunlugu + 1, 4).Value = WsIslemGunlugu.Cells(i, 4).Value + 1
                End If
                'Dönem sıra no
                SayDonem = WsIslemGunlugu.Range("F100000").End(xlUp).Row
                If SayDonem < 7 Then
                    WsIslemGunlugu.Cells(Say2IslemGunlugu + 1, 6).Value = 1
                Else
                    i = Say2IslemGunlugu + 1
                    Do Until WsIslemGunlugu.Cells(i, 6).Value <> ""
                        i = i - 1
                    Loop
                    WsIslemGunlugu.Cells(Say2IslemGunlugu + 1, 6).Value = WsIslemGunlugu.Cells(i, 6).Value + 1
                End If
 
                ilkrow = Say2IslemGunlugu + 1
                sonrow = Say2IslemGunlugu + 1 + (SonSira - IlkSira)
                
            Else 'Yeni veriyi SayAyracIslemGunlugu+1'e yaz
                'Genel sıra no
                SayGenel = WsIslemGunlugu.Range("D100000").End(xlUp).Row
                If SayGenel < 7 Then
                    WsIslemGunlugu.Cells(SayAyracIslemGunlugu + 1, 4).Value = 1
                Else
                    i = SayAyracIslemGunlugu + 1
                    Do Until WsIslemGunlugu.Cells(i, 4).Value <> ""
                        i = i - 1
                    Loop
                    WsIslemGunlugu.Cells(SayAyracIslemGunlugu + 1, 4).Value = WsIslemGunlugu.Cells(i, 4).Value + 1
                End If
                'Dönem sıra no
                SayDonem = WsIslemGunlugu.Range("E100000").End(xlUp).Row
                WsIslemGunlugu.Cells(SayAyracIslemGunlugu + 1, 6).Value = 1

                ilkrow = SayAyracIslemGunlugu + 1
                sonrow = SayAyracIslemGunlugu + 1 + (SonSira - IlkSira)
                
            End If
        Else 'Cari dönemden önceki dönemin verisi
            'MsgBox "Buradayım"
            ModulAyrac = DateAdd("m", 1, CDate(ModulAyrac)) 'Sonraki dönemi bul ve onun üstüne satır ekle
            Set BulIslemGunlugu = WsIslemGunlugu.Range("E:E").Find(What:=CDate(ModulAyrac), SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlFormulas, LookAt:=xlWhole)
            If Not BulIslemGunlugu Is Nothing Then
                For i = 1 To (SonSira - IlkSira) + 1
                    WsIslemGunlugu.Rows(BulIslemGunlugu.Row).EntireRow.Insert Shift:=xlUp
                Next i
'                WsIslemGunlugu.Range("E" & BulIslemGunlugu.Row - 1).Value = "İkincisi"
'                WsIslemGunlugu.Range("E" & BulIslemGunlugu.Row - 2).Value = "İlki"

                ilkrow = BulIslemGunlugu.Row - (SonSira - IlkSira + 1)
                sonrow = BulIslemGunlugu.Row - 1
                
                'Genel sıra no
                SayGenel = WsIslemGunlugu.Range("D100000").End(xlUp).Row
                If SayGenel < 7 Then
                    WsIslemGunlugu.Cells(ilkrow, 4).Value = 1
                    If SayGenel > ilkrow Then 'ilkrow dan sonra gelen sıra no.ları düzelt
                        For i = ilkrow + 1 To SayGenel
                            If WsIslemGunlugu.Cells(i, 4).Value <> "" Then
                                WsIslemGunlugu.Cells(i, 4).Value = WsIslemGunlugu.Cells(i, 4).Value + 1
                            End If
                        Next i
                    End If
                Else
                    i = ilkrow
                    Do Until WsIslemGunlugu.Cells(i, 4).Value <> ""
                        i = i - 1
                        If WsIslemGunlugu.Cells(i, 4).Value <> "" And Not IsNumeric(WsIslemGunlugu.Cells(i, 4).Value) Then
                            WsIslemGunlugu.Cells(ilkrow, 4).Value = 1
                            GoTo LoopSon1
                        End If
                    Loop
                    WsIslemGunlugu.Cells(ilkrow, 4).Value = WsIslemGunlugu.Cells(i, 4).Value + 1
LoopSon1:
                    If SayGenel > ilkrow Then 'ilkrow dan sonra gelen sıra no.ları düzelt
                        For i = ilkrow + 1 To SayGenel
                            If WsIslemGunlugu.Cells(i, 4).Value <> "" Then
                                WsIslemGunlugu.Cells(i, 4).Value = WsIslemGunlugu.Cells(i, 4).Value + 1
                            End If
                        Next i
                    End If
                End If
                
                'Dönem sıra no
                'SayDonem = WsIslemGunlugu.Range("E100000").End(xlUp).Row
                i = ilkrow
                Do Until i < donemrow
                    i = i - 1
                    If WsIslemGunlugu.Cells(i, 5).Value <> "" And Not IsNumeric(WsIslemGunlugu.Cells(i, 5).Value) Then
                        WsIslemGunlugu.Cells(ilkrow, 6).Value = 1
                        GoTo LoopSon2
                    End If
                    If WsIslemGunlugu.Cells(i, 6).Value <> "" And IsNumeric(WsIslemGunlugu.Cells(i, 6).Value) Then
                        WsIslemGunlugu.Cells(ilkrow, 6).Value = WsIslemGunlugu.Cells(i, 6).Value + 1
                        GoTo LoopSon2
                    End If
                Loop
LoopSon2:

            End If
            ModulAyrac = DateAdd("m", -1, CDate(ModulAyrac)) 'Yukarıda atadığın +1 ayı geri al
        End If
    Else
        'İşlem günlüğünde kayıtlı en eski DÖNEMDEN daha eski bir döneme ilişkin veri girişi gerçekleşmesi durumu.
    '    MsgBox "Outbound Qty"
        GoTo Son
    End If
End If

DonemAyniyiAtla:


'Dolguları kaldır/Biçimlendirmeleri düzelt
WsIslemGunlugu.Range("E" & ilkrow & ":T" & sonrow).Interior.Color = xlNone
WsIslemGunlugu.Range("E" & ilkrow & ":T" & sonrow).Font.Color = RGB(0, 0, 0)
WsIslemGunlugu.Range("E" & ilkrow & ":T" & sonrow).Font.Bold = False
WsIslemGunlugu.Range("B" & ilkrow & ":T" & sonrow).NumberFormat = "@"
WsIslemGunlugu.Range("P" & ilkrow & ":Q" & sonrow).NumberFormat = "General"
WsIslemGunlugu.Range("F" & ilkrow & ":F" & sonrow).NumberFormat = "General"
WsIslemGunlugu.Range("D" & ilkrow & ":D" & sonrow).NumberFormat = "General"
WsIslemGunlugu.Range("B" & ilkrow & ":T" & sonrow).WrapText = True
WsIslemGunlugu.Range("F" & ilkrow & ":G" & sonrow).VerticalAlignment = xlCenter
WsIslemGunlugu.Range("P" & ilkrow & ":Q" & sonrow).VerticalAlignment = xlCenter
WsIslemGunlugu.Range("F" & ilkrow & ":G" & sonrow).HorizontalAlignment = xlCenter
WsIslemGunlugu.Range("P" & ilkrow & ":Q" & sonrow).HorizontalAlignment = xlCenter


'Zaman damgaları
WsIslemGunlugu.Cells(ilkrow, 2).Value = WsRapor.Cells(IlkSira, 165).Value
WsIslemGunlugu.Cells(sonrow, 3).Value = WsRapor.Cells(IlkSira, 165).Value
'Verileri yaz
WsIslemGunlugu.Range(WsIslemGunlugu.Cells(ilkrow, 7), WsIslemGunlugu.Cells(sonrow, 7)).Value = WsRapor.Range(WsRapor.Cells(IlkSira, 13), WsRapor.Cells(SonSira, 13)).Value 'Rapor no
WsIslemGunlugu.Cells(ilkrow, 8).Value = WsRapor.Cells(IlkSira, 91).Value 'İl
WsIslemGunlugu.Cells(ilkrow, 9).Value = WsRapor.Cells(IlkSira, 92).Value 'İlçe
WsIslemGunlugu.Cells(ilkrow, 10).Value = GelenTema
WsIslemGunlugu.Cells(ilkrow, 11).Value = WsRapor.Cells(IlkSira, 95).Value 'Belge tarihi
WsIslemGunlugu.Cells(ilkrow, 12).Value = "" 'Belge no
WsIslemGunlugu.Cells(ilkrow, 13).Value = WsRapor.Cells(IlkSira, 95).Value 'FinansalBirimya ulaşma tarihi
WsIslemGunlugu.Cells(ilkrow, 14).Value = WsRapor.Cells(IlkSira, 95).Value 'Tespit tarihi
WsIslemGunlugu.Range(WsIslemGunlugu.Cells(ilkrow, 15), WsIslemGunlugu.Cells(sonrow, 15)).Value = WsRapor.Range(WsRapor.Cells(IlkSira, 130), WsRapor.Cells(SonSira, 130)).Value 'Öğe türü
WsIslemGunlugu.Range(WsIslemGunlugu.Cells(ilkrow, 16), WsIslemGunlugu.Cells(sonrow, 16)).Value = WsRapor.Range(WsRapor.Cells(IlkSira, 133), WsRapor.Cells(SonSira, 133)).Value 'Öğe değeri
WsIslemGunlugu.Range(WsIslemGunlugu.Cells(ilkrow, 17), WsIslemGunlugu.Cells(sonrow, 17)).Value = WsRapor.Range(WsRapor.Cells(IlkSira, 136), WsRapor.Cells(SonSira, 136)).Value 'Adet
WsIslemGunlugu.Cells(ilkrow, 18).Value = WsRapor.Cells(IlkSira, 98).Value 'Tema
WsIslemGunlugu.Range(WsIslemGunlugu.Cells(ilkrow, 19), WsIslemGunlugu.Cells(sonrow, 19)).Value = WsRapor.Range(WsRapor.Cells(IlkSira, 142), WsRapor.Cells(SonSira, 142)).Value 'Açıklama
WsIslemGunlugu.Range(WsIslemGunlugu.Cells(ilkrow, 20), WsIslemGunlugu.Cells(sonrow, 20)).Value = WsRapor.Range(WsRapor.Cells(IlkSira, 213), WsRapor.Cells(SonSira, 213)).Value 'Baskı tekniği


'Kenarlıklar.
Set Kenarlar = WsIslemGunlugu.Range("D" & ilkrow & ":T" & sonrow)
With Kenarlar.Borders(xlEdgeTop)
    .LineStyle = xlContinuous
    .Color = RGB(174, 185, 194)
    .Weight = xlThin
End With
With Kenarlar.Borders(xlEdgeBottom)
    .LineStyle = xlContinuous
    .Color = RGB(174, 185, 194)
    .Weight = xlThin
End With
With Kenarlar.Borders(xlInsideVertical)
    .LineStyle = xlContinuous
    .Color = RGB(174, 185, 194)
    .Weight = xlThin
End With
With Kenarlar.Borders(xlInsideHorizontal)
    .LineStyle = xlContinuous
    .Color = RGB(174, 185, 194)
    .Weight = xlThin
End With

If DelControl = True Then
    'Kaynak dönemde yer alan boş satır aralığını kaldır
    Set BulIslemGunlugu = WsIslemGunlugu.Range("B:B").Find(What:="Sil", SearchDirection:=xlNext, _
    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    If Not BulIslemGunlugu Is Nothing Then
        ilkrowx = BulIslemGunlugu.Row
        Set BulIslemGunlugu = WsIslemGunlugu.Range("C:C").Find(What:="Sil", SearchDirection:=xlNext, _
        SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
        If Not BulIslemGunlugu Is Nothing Then
            sonrowx = BulIslemGunlugu.Row
        End If
        WsIslemGunlugu.Rows(ilkrowx & ":" & sonrowx).EntireRow.Delete
    End If
End If


Son:

'İşlem günlüğünde aşağı git
Say2IslemGunlugu = WsIslemGunlugu.Range("C100000").End(xlUp).Row
On Error Resume Next
ActiveWindow.ScrollRow = Say2IslemGunlugu - 10
On Error GoTo 0

WsIslemGunlugu.Columns("B:C").EntireColumn.Hidden = True

WsIslemGunlugu.Protect Password:="123"

'İşlem günlüğü açıksa kaydet ve kapat.
OpenControl = IsWorkBookOpen(IslemGunlugu)
If OpenControl = True Then
    Workbooks("System Registry Report 2.1.xlsx").Save
    Workbooks("System Registry Report 2.1.xlsx").Close SaveChanges:=False
End If


If IlceSakla <> "" Then
    WsRapor.Cells(IlkSira, 92).Value = IlceSakla
End If


'WsRapor.Protect Password:="123"

Out:

Application.DisplayAlerts = True
Application.EnableEvents = True
Application.ScreenUpdating = True


End Sub

Sub Rapor3_1TeslimTutanaklari()
Dim i As Integer, TarihBul As Range
Dim Say As Long, j As Long, Cont As Long, Sno As Variant
Dim LstBx As MSForms.ListBox, ctl As MSForms.Control
Dim LblSira1 As MSForms.Label
Dim LblSira2 As MSForms.Label
Dim Contx As Long

CalTarih = CalTarihTakip

Say = ThisWorkbook.Worksheets(5).Range("FH100000").End(xlUp).Row
If Say < 7 Then
    GoTo Son
End If

'Cont = 0
Contx = 0
TarihiTekrarla1:
Set TarihBul = ThisWorkbook.Worksheets(5).Range("EY6:EY100000").Find(What:=CalTarih, SearchDirection:=xlNext, _
                SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
If Not TarihBul Is Nothing Then
    'Cells(TarihBul.Row, 31).Value
Else
    CalTarih = CDate(CalTarih)
    CalTarih = DateAdd("d", 1, CalTarih)
    CalTarih = CStr(CalTarih)
    If Mid(CalTarih, 2, 1) = "." Then 'Günün soluna 0 ekle
        CalTarih = "0" & CalTarih
    End If
    If Mid(CalTarih, 5, 1) = "." Then 'Ayın soluna 0 ekle
        CalTarih = Left(CalTarih, 3) & "0" & Mid(CalTarih, 4, 6)
    End If
    'MsgBox CalTarih
    Contx = Contx + 1
    If Contx = 100 Then
        'MsgBox "Belirtilen tarihten sonra herhangi bir tutanak1 işlemi yapılmadığından işleminiz gerçekleştirilemiyor.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    GoTo TarihiTekrarla1
End If


Cont = ContTakip
'Cont = 0
For j = Say To TarihBul.Row Step -1
    'If Cont < 50 Then
        If ThisWorkbook.Worksheets(5).Cells(j, 5).Value <> "" And ThisWorkbook.Worksheets(5).Cells(j, 155).Value <> "" Then
            Cont = Cont + 1
            If ThisWorkbook.Worksheets(5).Cells(j, 103).Value <> "" Then
                If Left(ThisWorkbook.Worksheets(5).Cells(j, 102).Value, 3) = "İl " Then
                    Sno = ThisWorkbook.Worksheets(5).Cells(j, 5).Value & " | G | " & ThisWorkbook.Worksheets(5).Cells(j, 156).Value & " | " & ThisWorkbook.Worksheets(5).Cells(j, 103).Value
                ElseIf Left(ThisWorkbook.Worksheets(5).Cells(j, 102).Value, 3) = "İlç" Then
                    Sno = ThisWorkbook.Worksheets(5).Cells(j, 5).Value & " | G | " & ThisWorkbook.Worksheets(5).Cells(j, 156).Value & " | " & ThisWorkbook.Worksheets(5).Cells(j, 103).Value
                Else
                    Sno = ThisWorkbook.Worksheets(5).Cells(j, 5).Value & " | G | " & ThisWorkbook.Worksheets(5).Cells(j, 156).Value & " | " & ThisWorkbook.Worksheets(5).Cells(j, 102).Value & " " & ThisWorkbook.Worksheets(5).Cells(j, 103).Value
                End If
            ElseIf ThisWorkbook.Worksheets(5).Cells(j, 103).Value = "" Then
                If Left(ThisWorkbook.Worksheets(5).Cells(j, 102).Value, 3) = "İl " Then
                    Sno = ThisWorkbook.Worksheets(5).Cells(j, 5).Value & " | G | " & ThisWorkbook.Worksheets(5).Cells(j, 156).Value & " | " & ThisWorkbook.Worksheets(5).Cells(j, 91).Value & ThisWorkbook.Worksheets(5).Cells(j, 102).Value
                ElseIf Left(ThisWorkbook.Worksheets(5).Cells(j, 102).Value, 3) = "İlç" Then
                    Sno = ThisWorkbook.Worksheets(5).Cells(j, 5).Value & " | G | " & ThisWorkbook.Worksheets(5).Cells(j, 156).Value & " | " & ThisWorkbook.Worksheets(5).Cells(j, 92).Value & ThisWorkbook.Worksheets(5).Cells(j, 102).Value
                Else
                    Sno = ThisWorkbook.Worksheets(5).Cells(j, 5).Value & " | G | " & ThisWorkbook.Worksheets(5).Cells(j, 156).Value & " | " & ThisWorkbook.Worksheets(5).Cells(j, 102).Value
                End If
            End If

            Set LstBx = core_delivery_manager_UI.Frame1.Controls.Add("Forms.ListBox.1", "Frame1LstBx" & Cont)
            With LstBx
                .Top = (Cont - 1) * 12
                .Left = 18
                .Height = 12
                .Width = 300
                .SpecialEffect = fmSpecialEffectFlat
                .MultiSelect = fmMultiSelectMulti
                .TextAlign = fmTextAlignLeft '02092021
                .AddItem (Sno)
            End With

            Set LblSira1 = core_delivery_manager_UI.Frame1.Controls.Add("Forms.Label.1", "Lbl" & Cont)
            With LblSira1
                .Top = (Cont - 1) * 12
                .Left = 0
                .Height = 12
                .Width = 18
                .SpecialEffect = fmSpecialEffectEtched
                .TextAlign = fmTextAlignCenter
                .Caption = Cont
            End With

            ScrollTakip1 = ScrollTakip1 + 12
        End If
    'End If
Next j

ContTakip = Cont

Son:

'MsgBox "Rapor3_1: " & ContTakip

End Sub

''''TipB

Sub Rapor3_1TutanakTipB()

Dim DestOperasyon As String, SourceRapor3_1Normal As String, AutoPath As String, ReNameRapor3_1 As String
Dim fso As Object, objWord As Object, objDoc As Object
Dim OpenKontrolName As String, OpenControl As String, DestOpUserFolder As String, DestOpUserFolderName As String
Dim ContSay As Integer, KontrolFile As String, x As Long, y As Long, i As Long, j As Long
Dim IlkSiraBul As Range, SonSiraBul As Range, IlkSira As Long, SonSira As Long
Dim GelenTema As String, SourceDL As String, ReNameDL As String
Dim TxtFileRapor3_1 As String, TotalSayfaRapor3_1 As String, TxtFileDokum As String, TotalSayfaDokum As String
Dim DokumFileTxt As Object, DokumSayfaGonder As String
Dim SourceRaporNormal As String, SourceTutanak2Normal As String, SourceUstYaziNormal As String
Dim Bolum1 As String, Bolum2 As String, Bolum3 As String, Bolum4 As String, Bolum5 As String, Bolum6 As String
Dim Ek1 As String, Ek2 As String, Ek3 As String, Ek4 As String, Ek5 As String, Ek6 As String, Birlestir As String

Dim ReNameRaporNormal As String, SiraNoIlkSatir As Long, RaporIlgi As String
Dim AltRaporNoIlk As Long, AltRaporNoSon As Long, AdetTopla As Long
Dim TekCogulTipA As String, k As Integer
Dim MyRange As Object, TxtFileRapor As String, TotalSayfaRapor As String
Dim ReNameTutanak2Normal As String, GidenTema As String
Dim TxtFileTutanak2 As String, TotalSayfaTutanak2 As String
Dim ReNameUstYaziNormal As String, PaketTipi As String
Dim TxtFileUstYazi As String, TotalSayfaUstYazi As String
Dim Birimx As String, UserName As String, SourceRapor3_1Farkli As String
Dim Kolluk As String


'    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    'Application.EnableEvents = False
    'Worksheets(5).Unprotect Password:="123"
    'Pathfinder...
    AutoPath = ThisWorkbook.Path
    DestOperasyon = AutoPath & "\System Files\Operation\"
    'TUTANAK1 TANIMLARI
    SourceRapor3_1Normal = AutoPath & "\System Files\System Templates\Report 3 Statements\Report 3.1 – Type B.docm"
    'Birden fazla kullanıcının operasyon klasöründe çakışmasını önlemek için operasyon klasörünün içinde yeni klasör aç.
    DestOpUserFolderName = "Operation-" & Left(ThisWorkbook.name, InStr(ThisWorkbook.name, ".") - 1) '& " " & Format(Now(), "ddmmyyyyhhmmss")
    DestOpUserFolder = DestOperasyon & DestOpUserFolderName & "\"
    
    'Check the "System Files" folder name.
    If Not Dir(AutoPath & "\System Files\", vbDirectory) <> vbNullString Then
        MsgBox "Cannot access the directory " & AutoPath & "\System Files\" & ". The folder named 'System Files' may have been renamed or deleted.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    
    'Check the "Operation" folder name.
    If Not Dir(DestOperasyon, vbDirectory) <> vbNullString Then
        MsgBox "Cannot access the directory " & DestOperasyon & ". The folder named 'Operation' may have been renamed or deleted.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    
    'Check folder names.
    If Not Dir(SourceRapor3_1Normal, vbDirectory) <> vbNullString Then
        MsgBox "Cannot access the directory " & SourceRapor3_1Normal & ". The names of folders and/or files in this directory may have been changed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If

    'Operation klasörü içinde kullanıcı modülü klasörü yoksa oluştur.
    If Not Dir(DestOpUserFolder, vbDirectory) <> vbNullString Then
        MkDir DestOpUserFolder
    End If
    
    'On Error Resume Next
    ReNameRapor3_1 = Cells(ActiveCell.Row, 5).Value & "-" & Cells(6, 6).Value
'________________________________________

    
    'Close the all Word application
    Call ModuleReport3.OpenWordControl
    
    'Operation klasöründeki docm uzantılı word dosyalarından açık olanları kapat ve temizle.
    OpenKontrolName = Dir(DestOpUserFolder & "*.docm")
    Do While OpenKontrolName <> ""
        OpenControl = IsFileOpen(DestOpUserFolder & OpenKontrolName)
        If OpenControl = True Then 'Açıksa
            On Error Resume Next
            Set objWord = GetObject(, "Word.Application")
            Set objWord = GetObject(, "Word.Application")
            Set objWord = GetObject(, "Word.Application")
            Set objWord = GetObject(, "Word.Application")
            Set objWord = GetObject(, "Word.Application")
            objWord.Quit SaveChanges:=True
            'MsgBox "Dosya OpenKontrol methodu ile kapatıldı."

        End If
        OpenKontrolName = Dir()
    Loop
     
    Set objWord = Nothing
    Set objDoc = Nothing

'________________________________________

    On Error Resume Next
'    Klasörün içindeki tüm dosyaları sil (txt, docm vb.)
    ContSay = 0
    KontrolFile = Dir(DestOpUserFolder & "*.???")
    Do While KontrolFile <> ""
        ContSay = ContSay + 1
        KontrolFile = Dir()
    Loop
    If ContSay > 0 Then
        Kill DestOpUserFolder & "*.???"
    End If
    
    'Dosyayı şablondan operasyon klasörüne kopyala ve adını değiştir.
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.CopyFile (SourceRapor3_1Normal), DestOpUserFolder & ReNameRapor3_1 & ".docm", True
    
'________________________________________

    'Oluşturulacak dosyayı aç
    On Error Resume Next
    Set objWord = GetObject(, "Word.Application")
    Set objWord = GetObject(, "Word.Application")
    Set objWord = GetObject(, "Word.Application")
    Set objWord = GetObject(, "Word.Application")
    Set objWord = GetObject(, "Word.Application")
    If objWord Is Nothing Then
        'MsgBox "Dosya oluşturmada CreateObject methodu kullanılacak."
        Set objWord = CreateObject("Word.Application")
        objWord.Visible = False
    End If
    objWord.Documents.Open FileName:=DestOpUserFolder & ReNameRapor3_1 & ".docm"
    objWord.Visible = True
    objWord.Activate 'Ekrana getirir.
    'objDoc.Activate 'Ekrana getirmez.
    objWord.Application.WindowState = wdWindowStateMaximize
    Set objDoc = GetObject(DestOpUserFolder & ReNameRapor3_1 & ".docm")
'________________________________________

    Set IlkSiraBul = Range("FG7:FG100000").Find(What:=Cells(ActiveCell.Row, 5).Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    Set SonSiraBul = Range("FH7:FH100000").Find(What:=Cells(ActiveCell.Row, 5).Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    If Not IlkSiraBul Is Nothing Then
        IlkSira = IlkSiraBul.Row
    Else
        MsgBox "The serial number could not be found, so the operation cannot be completed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    If Not SonSiraBul Is Nothing Then
        SonSira = SonSiraBul.Row
    Else
        MsgBox "The serial number could not be found, so the operation cannot be completed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    
    'Birim
    Birimx = UCase(Replace(Replace(Worksheets(2).Cells(6, 99).Value, "i", "I"), "ı", "I")) & " UNIT"
    objDoc.Tables(1).Cell(Row:=1, Column:=1).Range.Text = Birimx
    
    'Dosyada içerikleri değiştir.
    'Kişinin açık kimliği var/yok
    If Cells(ActiveCell.Row, 126).Value <> "" And Cells(ActiveCell.Row, 126).Value > 0 Then
        'var
    Else
        'yok
        Ek1 = ""
        Set MyRange = objDoc.Tables(1).Cell(Row:=4, Column:=1).Range
        With MyRange.Find
            .Text = "identified as xxxxx xxxxx xx xxxxx, "
            .Replacement.Text = Ek1
            .Execute Replace:=wdReplaceAll
        End With
    End If
    'Ad Soyad
    Ek1 = Cells(ActiveCell.Row, 109).Value
    Set MyRange = objDoc.Tables(1).Cell(Row:=4, Column:=1).Range
    With MyRange.Find
        .Text = "<fullName>"
        .Replacement.Text = Ek1
        .Execute Replace:=wdReplaceAll
    End With
    MyRange.HighlightColorIndex = wdNoHighlight
    'Getirilme amacı
    Ek1 = Cells(ActiveCell.Row, 106).Value
    Set MyRange = objDoc.Tables(1).Cell(Row:=4, Column:=1).Range
    With MyRange.Find
        .Text = "<purpose>"
        .Replacement.Text = Ek1
        .Execute Replace:=wdReplaceAll
    End With
    MyRange.HighlightColorIndex = wdNoHighlight
    'TipA(un)(ların)
    'Tekil-çoğul tipA
    AdetTopla = Application.Sum(Range(Cells(IlkSira, 136), Cells(SonSira, 136)))
    If AdetTopla = 1 Then
        TekCogulTipA = "Type B item"
        Ek2 = "has"
        Ek3 = "Type B Item"
    ElseIf AdetTopla > 1 Then
        TekCogulTipA = "Type B items"
        Ek2 = "have"
        Ek3 = "Type B Items"
    End If
    Set MyRange = objDoc.Tables(1).Cell(Row:=4, Column:=1).Range
    With MyRange.Find
        .Text = "<typeA>"
        .Replacement.Text = TekCogulTipA
        .Execute Replace:=wdReplaceAll
    End With
    MyRange.HighlightColorIndex = wdNoHighlight
    Set MyRange = objDoc.Tables(1).Cell(Row:=4, Column:=1).Range
    With MyRange.Find
        .Text = "<have_has>"
        .Replacement.Text = Ek2
        .Execute Replace:=wdReplaceAll
    End With
    MyRange.HighlightColorIndex = wdNoHighlight
    
    'Kolluk
    If Cells(ActiveCell.Row, 102).Value = "Provincial Directorate B" Then
        Kolluk = Cells(ActiveCell.Row, 91).Value & " Provincial Governorship Provincial Directorate B " & Cells(ActiveCell.Row, 103).Value
    ElseIf Cells(ActiveCell.Row, 102).Value = "District Directorate B" Then
        Kolluk = Cells(ActiveCell.Row, 92).Value & " District Governorship District Directorate B " & Cells(ActiveCell.Row, 103).Value
    ElseIf Cells(ActiveCell.Row, 102).Value = "Provincial Directorate C" Then
        Kolluk = Cells(ActiveCell.Row, 91).Value & " Provincial Governorship Provincial Directorate C " & Cells(ActiveCell.Row, 103).Value
    ElseIf Cells(ActiveCell.Row, 102).Value = "District Directorate C" Then
        Kolluk = Cells(ActiveCell.Row, 92).Value & " District Governorship District Directorate C " & Cells(ActiveCell.Row, 103).Value
    ElseIf Cells(ActiveCell.Row, 102).Value = "Provincial Directorate D" Then
        Kolluk = Cells(ActiveCell.Row, 91).Value & " Provincial Governorship Provincial Directorate D " & Cells(ActiveCell.Row, 103).Value
    ElseIf Cells(ActiveCell.Row, 102).Value = "District Directorate D" Then
        Kolluk = Cells(ActiveCell.Row, 92).Value & " District Governorship District Directorate D " & Cells(ActiveCell.Row, 103).Value
    ElseIf Cells(ActiveCell.Row, 102).Value = "Provincial Directorate E" Then
        Kolluk = Cells(ActiveCell.Row, 91).Value & " Provincial Governorship Provincial Directorate E " & Cells(ActiveCell.Row, 103).Value
    ElseIf Cells(ActiveCell.Row, 102).Value = "District Directorate E" Then
        Kolluk = Cells(ActiveCell.Row, 92).Value & " District Governorship District Directorate E " & Cells(ActiveCell.Row, 103).Value
    ElseIf InStr(Cells(ActiveCell.Row, 102).Value, "General Directorate") <> 0 Or InStr(Cells(ActiveCell.Row, 102).Value, "Regional Directorate") <> 0 Then
        If Cells(ActiveCell.Row, 103).Value <> "" Then
            Kolluk = WorksheetFunction.Proper(ThisWorkbook.Worksheets(2).Cells(6, 111).Value) & " " & Cells(ActiveCell.Row, 102).Value & " " & Cells(ActiveCell.Row, 103).Value
        Else
            Kolluk = WorksheetFunction.Proper(ThisWorkbook.Worksheets(2).Cells(6, 111).Value) & " " & Cells(ActiveCell.Row, 102).Value
        End If
    Else
        If Cells(ActiveCell.Row, 103).Value <> "" Then
            If InStr(Cells(ActiveCell.Row, 102).Value, "X.X. ") > 0 Then
                Kolluk = Mid(Cells(ActiveCell.Row, 102).Value, 6, Len(Cells(ActiveCell.Row, 102).Value)) & " " & Cells(ActiveCell.Row, 103).Value
            Else
                Kolluk = Cells(ActiveCell.Row, 102).Value & " " & Cells(ActiveCell.Row, 103).Value
            End If
        Else
            If InStr(Cells(ActiveCell.Row, 102).Value, "X.X. ") > 0 Then
                Kolluk = Mid(Cells(ActiveCell.Row, 102).Value, 6, Len(Cells(ActiveCell.Row, 102).Value))
            Else
                Kolluk = Cells(ActiveCell.Row, 102).Value
            End If
        End If
    End If
    
    'Recipient
    Set MyRange = objDoc.Tables(2).Cell(Row:=1, Column:=1).Range
    With MyRange.Find
        .Text = "<recipient>"
        .Replacement.Text = Kolluk
        .Execute Replace:=wdReplaceAll
    End With
    MyRange.HighlightColorIndex = wdNoHighlight
    'Tutanak tarihi
    Ek1 = Cells(ActiveCell.Row, 95).Value
    Set MyRange = objDoc.Tables(2).Cell(Row:=1, Column:=1).Range
    With MyRange.Find
        .Text = "<reportDate>"
        .Replacement.Text = Ek1
        .Execute Replace:=wdReplaceAll
    End With
    MyRange.HighlightColorIndex = wdNoHighlight

    'imzalar
    objDoc.Tables(3).Cell(Row:=4, Column:=1).Range.Text = Cells(ActiveCell.Row, 184).Value 'Ad Soyad1
    objDoc.Tables(3).Cell(Row:=5, Column:=1).Range.Text = Cells(ActiveCell.Row, 185).Value 'Unvan1
    objDoc.Tables(3).Cell(Row:=4, Column:=2).Range.Text = Cells(ActiveCell.Row, 187).Value 'Ad Soyad2
    objDoc.Tables(3).Cell(Row:=5, Column:=2).Range.Text = Cells(ActiveCell.Row, 188).Value 'Unvan2
    objDoc.Tables(3).Cell(Row:=4, Column:=3).Range.Text = Cells(ActiveCell.Row, 190).Value 'Ad Soyad3
    objDoc.Tables(3).Cell(Row:=5, Column:=3).Range.Text = Cells(ActiveCell.Row, 191).Value 'Unvan3
    
    'Tablo başlığı
    objDoc.Tables(4).Cell(Row:=1, Column:=1).Range.Text = Ek3 & " Evaluated as Invalid:"
    'Tabloya satır ekle
    x = 0
    If SonSira - IlkSira > 0 Then
        With objDoc.Tables(4)
            For i = IlkSira To SonSira - 1
                .Rows.Add BeforeRow:=objDoc.Tables(4).Rows(3)
                x = x + 1
            Next i
        End With
    End If
    For i = 3 To SonSira - IlkSira + 3
        objDoc.Tables(4).Cell(Row:=i, Column:=1).Range.Text = i - 2 'Tablo sıra no
        objDoc.Tables(4).Cell(Row:=i, Column:=2).Range.Text = Cells(IlkSira + i - 3, 130).Value 'Öğe türü
        objDoc.Tables(4).Cell(Row:=i, Column:=3).Range.Text = Cells(IlkSira + i - 3, 133).Value 'Öğe değeri
        objDoc.Tables(4).Cell(Row:=i, Column:=4).Range.Text = Cells(IlkSira + i - 3, 136).Value 'Adet
        objDoc.Tables(4).Cell(Row:=i, Column:=5).Range.Text = Cells(IlkSira + i - 3, 139).Value 'Öğe ID No
    Next i
    
    'Tablo başlığı
    objDoc.Tables(5).Cell(Row:=1, Column:=1).Range.Text = "The Person Who Delivered the " & Ek3 & ":"
    objDoc.Tables(5).Cell(Row:=2, Column:=3).Range.Text = Cells(ActiveCell.Row, 109).Value 'Ad Soyad
    objDoc.Tables(5).Cell(Row:=3, Column:=3).Range.Text = Cells(ActiveCell.Row, 110).Value 'TCK No
    If Cells(ActiveCell.Row, 117).Value <> "" And Cells(ActiveCell.Row, 118).Value <> "" Then 'Kimlik Türü ve No
        objDoc.Tables(5).Cell(Row:=4, Column:=3).Range.Text = Cells(ActiveCell.Row, 117).Value & " - " & Cells(ActiveCell.Row, 118).Value
    ElseIf Cells(ActiveCell.Row, 117).Value <> "" And Cells(ActiveCell.Row, 118).Value = "" Then
        objDoc.Tables(5).Cell(Row:=4, Column:=3).Range.Text = Cells(ActiveCell.Row, 117).Value
    ElseIf Cells(ActiveCell.Row, 117).Value = "" And Cells(ActiveCell.Row, 118).Value <> "" Then
        objDoc.Tables(5).Cell(Row:=4, Column:=3).Range.Text = Cells(ActiveCell.Row, 118).Value
    End If
    objDoc.Tables(5).Cell(Row:=5, Column:=3).Range.Text = Cells(ActiveCell.Row, 111).Value 'Baba Adı
    If Cells(ActiveCell.Row, 112).Value <> "" And Cells(ActiveCell.Row, 113).Value <> "" Then 'Doğum Yeri ve Tarihi
        objDoc.Tables(5).Cell(Row:=6, Column:=3).Range.Text = Cells(ActiveCell.Row, 112).Value & " - " & Cells(ActiveCell.Row, 113).Value
    ElseIf Cells(ActiveCell.Row, 112).Value <> "" And Cells(ActiveCell.Row, 113).Value = "" Then
        objDoc.Tables(5).Cell(Row:=6, Column:=3).Range.Text = Cells(ActiveCell.Row, 112).Value
    ElseIf Cells(ActiveCell.Row, 112).Value = "" And Cells(ActiveCell.Row, 113).Value <> "" Then
        objDoc.Tables(5).Cell(Row:=6, Column:=3).Range.Text = Cells(ActiveCell.Row, 113).Value
    End If
    objDoc.Tables(5).Cell(Row:=7, Column:=3).Range.Text = Cells(ActiveCell.Row, 119).Value 'Nüfusa Kayıtlı Olduğu Yer
    objDoc.Tables(5).Cell(Row:=8, Column:=3).Range.Text = Cells(ActiveCell.Row, 120).Value 'Cilt No, Aile Sıra No, Sıra No
    If Cells(ActiveCell.Row, 123).Value <> "" And Cells(ActiveCell.Row, 116).Value <> "" Then 'Adres/Telefon Numarası
        objDoc.Tables(5).Cell(Row:=9, Column:=3).Range.Text = Cells(ActiveCell.Row, 123).Value & " - " & Cells(ActiveCell.Row, 116).Value
    ElseIf Cells(ActiveCell.Row, 123).Value <> "" And Cells(ActiveCell.Row, 116).Value = "" Then
        objDoc.Tables(5).Cell(Row:=9, Column:=3).Range.Text = Cells(ActiveCell.Row, 123).Value
    ElseIf Cells(ActiveCell.Row, 123).Value = "" And Cells(ActiveCell.Row, 116).Value <> "" Then
        objDoc.Tables(5).Cell(Row:=9, Column:=3).Range.Text = Cells(ActiveCell.Row, 116).Value
    End If
    'Getiren Kişi/Organization B Mensubu İmza alanı
    If Cells(ActiveCell.Row, 124).Value = "Yes" Then
        objDoc.Tables(6).Cell(Row:=3, Column:=1).Range.Text = "Signature"
        objDoc.Tables(6).Cell(Row:=4, Column:=1).Range.Text = "Full Name"
        objDoc.Tables(6).Cell(Row:=5, Column:=1).Range.Text = "The Person Who Delivered the " & Ek3
        objDoc.Tables(6).Cell(Row:=3, Column:=2).Range.Text = "Signature"
        objDoc.Tables(6).Cell(Row:=4, Column:=2).Range.Text = "Full Name"
        objDoc.Tables(6).Cell(Row:=5, Column:=2).Range.Text = "Officer Who Confiscated the " & Ek3
    ElseIf Cells(ActiveCell.Row, 124).Value = "No" Then
        objDoc.Tables(6).Cell(Row:=3, Column:=2).Range.Text = "Signature"
        objDoc.Tables(6).Cell(Row:=4, Column:=2).Range.Text = "Full Name"
        objDoc.Tables(6).Cell(Row:=5, Column:=2).Range.Text = "The Person Who Delivered the " & Ek3
    End If

    'Ek kimlik fotokopisi
    If Cells(ActiveCell.Row, 126).Value <> "" Then
        objDoc.Tables(7).Cell(Row:=1, Column:=1).Range.Text = "Attachment"
        objDoc.Tables(7).Cell(Row:=1, Column:=2).Range.Text = ":"
        If Cells(ActiveCell.Row, 126) > 1 Then
            objDoc.Tables(7).Cell(Row:=1, Column:=3).Range.Text = "ID Photocopy (" & Cells(ActiveCell.Row, 126) & " pages)"
        Else
            objDoc.Tables(7).Cell(Row:=1, Column:=3).Range.Text = "ID Photocopy (" & Cells(ActiveCell.Row, 126) & " page)"
        End If
    ElseIf Cells(ActiveCell.Row, 126).Value = "" And Cells(ActiveCell.Row, 127).Value = "No" Then
        objDoc.Tables(7).Cell(Row:=1, Column:=1).Range.Text = ""
        objDoc.Tables(7).Cell(Row:=1, Column:=2).Range.Text = ""
        objDoc.Tables(7).Cell(Row:=1, Column:=3).Range.Text = ""
    Else '126 boş ve 127 var ise notu eklemiş olacak.
        '
    End If

    'Olay bilgileri tekil
    objDoc.CheckBox3.Caption = "Condition4"
    objDoc.CheckBox4.Caption = "Condition5"
    objDoc.CheckBox5.Caption = "Condition6"
   
    'Olay bilgilerini işaretle
    If Mid(Cells(IlkSira, 128).Value, 1, 2) = "10" Then
        objDoc.CheckBox1.Value = False
    ElseIf Mid(Cells(IlkSira, 128).Value, 1, 2) = "11" Then
        objDoc.CheckBox1.Value = True
    End If
    If Mid(Cells(IlkSira, 128).Value, 4, 2) = "20" Then
        objDoc.CheckBox2.Value = False
    ElseIf Mid(Cells(IlkSira, 128).Value, 4, 2) = "21" Then
        objDoc.CheckBox2.Value = True
    End If
    If Mid(Cells(IlkSira, 128).Value, 7, 2) = "30" Then
        objDoc.CheckBox3.Value = False
    ElseIf Mid(Cells(IlkSira, 128).Value, 7, 2) = "31" Then
        objDoc.CheckBox3.Value = True
    End If
    If Mid(Cells(IlkSira, 128).Value, 10, 2) = "40" Then
        objDoc.CheckBox4.Value = False
    ElseIf Mid(Cells(IlkSira, 128).Value, 10, 2) = "41" Then
        objDoc.CheckBox4.Value = True
    End If
    If Mid(Cells(IlkSira, 128).Value, 13, 2) = "50" Then
        objDoc.CheckBox5.Value = False
    ElseIf Mid(Cells(IlkSira, 128).Value, 13, 2) = "51" Then
        objDoc.CheckBox5.Value = True
    End If
    If Mid(Cells(IlkSira, 128).Value, 16, 2) = "60" Then
        objDoc.CheckBox6.Value = False
    ElseIf Mid(Cells(IlkSira, 128).Value, 16, 2) = "61" Then
        objDoc.CheckBox6.Value = True
    End If

    'Türkçe karakterleri düzelt
    objDoc.CheckBox1.Enabled = False
    objDoc.CheckBox1.Enabled = True
    objDoc.CheckBox2.Enabled = False
    objDoc.CheckBox2.Enabled = True
    objDoc.CheckBox3.Enabled = False
    objDoc.CheckBox3.Enabled = True
    objDoc.CheckBox4.Enabled = False
    objDoc.CheckBox4.Enabled = True
    objDoc.CheckBox5.Enabled = False
    objDoc.CheckBox5.Enabled = True
    objDoc.CheckBox6.Enabled = False
    objDoc.CheckBox6.Enabled = True

    'Alt bilgi ekle
    objDoc.Sections(1).Footers(wdHeaderFooterPrimary).Range.Tables(1).Cell(Row:=1, Column:=1).Range.Text = ReNameRapor3_1
    
    'Save üstünde iken save bağlı kodlar çalışmıyor; designmod açık kalıyor. Bu sebeple checkbox silme komutu
    'save komutunun altında konumlandırıldı.
    '(Güncelleme 25.11.2018: Bu sorun word'taki kodların yeniden çalıştırılması ile çözüldü.
    If Cells(ActiveCell.Row, 124).Value = "No" Then 'Organization B mensubu yoksa
        'On Error GoTo Hata
        On Error Resume Next
        objDoc.Fields.Item(1).Delete
        objDoc.Fields.Item(5).Delete
    Else
        objDoc.Tables(7).Rows.Add 'BeforeRow:=objDoc.Tables(7).Rows(0)
    End If
    
    
    objWord.Run "Register_Event_Handler"
    
    'Sayfa sayısı kaydet komutuna bağlandı.
    objDoc.Save
    
    'Text dosyasından sayfayı çek.
    TxtFileRapor3_1 = DestOpUserFolder & "Report 3 Page Count.txt"
    #If UseADO Then
        Const adReadLine = -2&
        With CreateObject("ADODB.Stream")
            .Open
            .LoadFromFile TxtFileRapor3_1
            Do Until .EOS
                TotalSayfaRapor3_1 = .ReadText(adReadLine)
            Loop
            .Close
        End With
    #Else 'UseFSO
        Const Rapor3_1TristateTrue = -1&
        With CreateObject("Scripting.FileSystemObject")
            With .OpenTextFile(TxtFileRapor3_1, Format:=Rapor3_1TristateTrue)
                Do Until .AtEndOfStream
                    TotalSayfaRapor3_1 = .ReadLine
                Loop
                .Close
            End With
        End With
    #End If
    'MsgBox "Tutanak1: " & TotalSayfaRapor3_1

    If TumDoc = True Then
        objWord.Activate
'        For i = 1 To SayPrt
'            objDoc.PrintOut
'        Next i
        objDoc.PrintOut Background:=False, Copies:=SayPrt
        'objWord.Documents.Save
        objDoc.Close SaveChanges:=False
        objWord.Visible = False
    End If

    'Tutanak1 ve Döküm sayfa sayısı
    Cells(ActiveCell.Row, 169).Value = TotalSayfaRapor3_1
            
    'Nesneleri temizle
    Set fso = Nothing
    Set objWord = Nothing
    Set objDoc = Nothing
    GoTo Son
    
Hata:
MsgBox "An error occurred while retrieving the number of pages of the statement. Please manually enter the number of statement pages in the attachment section when creating the cover letter.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
'If Err.Number <> 0 Then
'MsgBox "Error # " & Str(Err.Number)
'End If

Son:

'Nesneleri temizle
Set fso = Nothing
Set objWord = Nothing
Set objDoc = Nothing

'Worksheets(5).Protect Password:="123"', DrawingObjects:=False
'Worksheets(5).Activate
Application.ScreenUpdating = True
Application.DisplayAlerts = True
'Application.EnableEvents = True

'End If

End Sub

Sub Rapor3_1Tutanak2TipB()

Dim DestOperasyon As String, SourceRapor3_1Normal As String, AutoPath As String, ReNameRapor3_1 As String
Dim fso As Object, objWord As Object, objDoc As Object
Dim OpenKontrolName As String, OpenControl As String, DestOpUserFolder As String, DestOpUserFolderName As String
Dim ContSay As Integer, KontrolFile As String, x As Long, y As Long, i As Long, j As Long
Dim IlkSiraBul As Range, SonSiraBul As Range, IlkSira As Long, SonSira As Long
Dim GelenTema As String, SourceDL As String, ReNameDL As String
Dim TxtFileRapor3_1 As String, TotalSayfaRapor3_1 As String, TxtFileDokum As String, TotalSayfaDokum As String
Dim DokumFileTxt As Object, DokumSayfaGonder As String
Dim SourceRaporNormal As String, SourceTutanak2Normal As String, SourceUstYaziNormal As String
Dim Bolum1 As String, Bolum2 As String, Bolum3 As String, Bolum4 As String, Bolum5 As String, Bolum6 As String
Dim Ek1 As String, Ek2 As String, Ek3 As String, Ek4 As String, Ek5 As String, Ek6 As String, Birlestir As String

Dim ReNameRaporNormal As String, SiraNoIlkSatir As Long, RaporIlgi As String
Dim AltRaporNoIlk As Long, AltRaporNoSon As Long, AdetTopla As Long
Dim TekCogulTipA As String, k As Integer
Dim MyRange As Object, TxtFileRapor As String, TotalSayfaRapor As String
Dim ReNameTutanak2Normal As String, GidenTema As String
Dim TxtFileTutanak2 As String, TotalSayfaTutanak2 As String
Dim ReNameUstYaziNormal As String
Dim TxtFileUstYazi As String, TotalSayfaUstYazi As String
Dim Birimx As String, UserName As String, Kolluk As String

'TUTANAK2 için prosedürü başlat

'    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    'Application.EnableEvents = False
    'Worksheets(5).Unprotect Password:="123"
    
    'Establish a left-to-right control mechanism for document creation.
    If Cells(ActiveCell.Row, 6).Value = "x" Then
        MsgBox "The Report 3 statement data contains missing or incorrect information, so the operation cannot be started.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If

    
    'Pathfinder...
    AutoPath = ThisWorkbook.Path
    DestOperasyon = AutoPath & "\System Files\Operation\"
    'TUTANAK2 TANIMLARI
    SourceTutanak2Normal = AutoPath & "\System Files\System Templates\Statement 2 Template\Statement 2.docm"
    'Birden fazla kullanıcının operasyon klasöründe çakışmasını önlemek için operasyon klasörünün içinde yeni klasör aç.
    DestOpUserFolderName = "Operation-" & Left(ThisWorkbook.name, InStr(ThisWorkbook.name, ".") - 1) '& " " & Format(Now(), "ddmmyyyyhhmmss")
    DestOpUserFolder = DestOperasyon & DestOpUserFolderName & "\"

    'Check the "System Files" folder name.
    If Not Dir(AutoPath & "\System Files\", vbDirectory) <> vbNullString Then
        MsgBox "Cannot access the directory " & AutoPath & "\System Files\" & ". The folder named 'System Files' may have been renamed or deleted.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    
    'Check the "Operation" folder name.
    If Not Dir(DestOperasyon, vbDirectory) <> vbNullString Then
        MsgBox "Cannot access the directory " & DestOperasyon & ". The folder named 'Operation' may have been renamed or deleted.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    
    'Check folder names.
    If Not Dir(SourceTutanak2Normal, vbDirectory) <> vbNullString Then
        MsgBox "Cannot access the directory " & SourceTutanak2Normal & ". The names of folders and/or files in this directory may have been changed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If

    'Operation klasörü içinde kullanıcı modülü klasörü yoksa oluştur.
    If Not Dir(DestOpUserFolder, vbDirectory) <> vbNullString Then
        MkDir DestOpUserFolder
    End If

    'On Error Resume Next
    ReNameTutanak2Normal = Cells(ActiveCell.Row, 5).Value & "-" & Cells(6, 7).Value
'________________________________________

    If TumDoc = False Then
        'Close the all Word application
        Call ModuleReport3.OpenWordControl

        'Operation klasöründeki docm uzantılı word dosyalarından açık olanları kapat ve temizle.
        OpenKontrolName = Dir(DestOpUserFolder & "*.docm")
        Do While OpenKontrolName <> ""
            OpenControl = IsFileOpen(DestOpUserFolder & OpenKontrolName)
            If OpenControl = True Then 'Açıksa
                On Error Resume Next
                Set objWord = GetObject(, "Word.Application")
                Set objWord = GetObject(, "Word.Application")
                Set objWord = GetObject(, "Word.Application")
                Set objWord = GetObject(, "Word.Application")
                Set objWord = GetObject(, "Word.Application")
                objWord.Quit SaveChanges:=True
                'MsgBox "Dosya OpenKontrol methodu ile kapatıldı."

            End If
            OpenKontrolName = Dir()
        Loop

        Set objWord = Nothing
        Set objDoc = Nothing
    '________________________________________

        On Error Resume Next
    '    Klasörün içindeki tüm dosyaları sil (txt, docm vb.)
        ContSay = 0
        KontrolFile = Dir(DestOpUserFolder & "*.???")
        Do While KontrolFile <> ""
            ContSay = ContSay + 1
            KontrolFile = Dir()
        Loop
        If ContSay > 0 Then
            Kill DestOpUserFolder & "*.???"
        End If
    End If

    'Dosyayı şablondan operasyon klasörüne kopyala ve adını değiştir.
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.CopyFile (SourceTutanak2Normal), DestOpUserFolder & ReNameTutanak2Normal & ".docm", True
'________________________________________

    'Oluşturulacak dosyayı aç
    On Error Resume Next
    Set objWord = GetObject(, "Word.Application")
    Set objWord = GetObject(, "Word.Application")
    Set objWord = GetObject(, "Word.Application")
    Set objWord = GetObject(, "Word.Application")
    Set objWord = GetObject(, "Word.Application")
    If objWord Is Nothing Then
        'MsgBox "Dosya oluşturmada CreateObject methodu kullanılacak."
        Set objWord = CreateObject("Word.Application")
        objWord.Visible = False
    End If
    objWord.Documents.Open FileName:=DestOpUserFolder & ReNameTutanak2Normal & ".docm"
    objWord.Visible = True
    objWord.Activate 'Ekrana getirir.
    'objDoc.Activate 'Ekrana getirmez.
    objWord.Application.WindowState = wdWindowStateMaximize
    Set objDoc = GetObject(DestOpUserFolder & ReNameTutanak2Normal & ".docm")
'________________________________________

    'Birim
    Birimx = UCase(Replace(Replace(Worksheets(2).Cells(6, 99).Value, "i", "I"), "ı", "I")) & " UNIT"
    objDoc.Tables(1).Cell(Row:=1, Column:=1).Range.Text = Birimx
    'Tutanak2 tarihi
    objDoc.Tables(1).Cell(Row:=4, Column:=3).Range.Text = Cells(ActiveCell.Row, 147).Value
    'Belge tarihi ve numarası
'    objDoc.Tables(1).Cell(Row:=6, Column:=3).Range.Text = "" 'Cells(ActiveCell.Row, 20).Value
'    objDoc.Tables(1).Cell(Row:=6, Column:=5).Range.Text = "" 'Cells(ActiveCell.Row, 21).Value
    objDoc.Tables(1).Rows(6).Delete

    'Kolluk
    If Cells(ActiveCell.Row, 102).Value = "Provincial Directorate B" Then
        Kolluk = Cells(ActiveCell.Row, 91).Value & " Provincial Governorship Provincial Directorate B " & Cells(ActiveCell.Row, 103).Value
    ElseIf Cells(ActiveCell.Row, 102).Value = "District Directorate B" Then
        Kolluk = Cells(ActiveCell.Row, 92).Value & " District Governorship District Directorate B " & Cells(ActiveCell.Row, 103).Value
    ElseIf Cells(ActiveCell.Row, 102).Value = "Provincial Directorate C" Then
        Kolluk = Cells(ActiveCell.Row, 91).Value & " Provincial Governorship Provincial Directorate C " & Cells(ActiveCell.Row, 103).Value
    ElseIf Cells(ActiveCell.Row, 102).Value = "District Directorate C" Then
        Kolluk = Cells(ActiveCell.Row, 92).Value & " District Governorship District Directorate C " & Cells(ActiveCell.Row, 103).Value
    ElseIf Cells(ActiveCell.Row, 102).Value = "Provincial Directorate D" Then
        Kolluk = Cells(ActiveCell.Row, 91).Value & " Provincial Governorship Provincial Directorate D " & Cells(ActiveCell.Row, 103).Value
    ElseIf Cells(ActiveCell.Row, 102).Value = "District Directorate D" Then
        Kolluk = Cells(ActiveCell.Row, 92).Value & " District Governorship District Directorate D " & Cells(ActiveCell.Row, 103).Value
    ElseIf Cells(ActiveCell.Row, 102).Value = "Provincial Directorate E" Then
        Kolluk = Cells(ActiveCell.Row, 91).Value & " Provincial Governorship Provincial Directorate E " & Cells(ActiveCell.Row, 103).Value
    ElseIf Cells(ActiveCell.Row, 102).Value = "District Directorate E" Then
        Kolluk = Cells(ActiveCell.Row, 92).Value & " District Governorship District Directorate E " & Cells(ActiveCell.Row, 103).Value
    ElseIf InStr(Cells(ActiveCell.Row, 102).Value, "General Directorate") <> 0 Or InStr(Cells(ActiveCell.Row, 102).Value, "Regional Directorate") <> 0 Then
        If Cells(ActiveCell.Row, 103).Value <> "" Then
            Kolluk = WorksheetFunction.Proper(ThisWorkbook.Worksheets(2).Cells(6, 111).Value) & " " & Cells(ActiveCell.Row, 102).Value & " " & Cells(ActiveCell.Row, 103).Value
        Else
            Kolluk = WorksheetFunction.Proper(ThisWorkbook.Worksheets(2).Cells(6, 111).Value) & " " & Cells(ActiveCell.Row, 102).Value
        End If
    Else
        If Cells(ActiveCell.Row, 103).Value <> "" Then
            If InStr(Cells(ActiveCell.Row, 102).Value, "X.X. ") > 0 Then
                Kolluk = Mid(Cells(ActiveCell.Row, 102).Value, 6, Len(Cells(ActiveCell.Row, 102).Value)) & " " & Cells(ActiveCell.Row, 103).Value
            Else
                Kolluk = Cells(ActiveCell.Row, 102).Value & " " & Cells(ActiveCell.Row, 103).Value
            End If
        Else
            If InStr(Cells(ActiveCell.Row, 102).Value, "X.X. ") > 0 Then
                Kolluk = Mid(Cells(ActiveCell.Row, 102).Value, 6, Len(Cells(ActiveCell.Row, 102).Value))
            Else
                Kolluk = Cells(ActiveCell.Row, 102).Value
            End If
        End If
    End If
    
    objDoc.Tables(1).Cell(Row:=5, Column:=3).Range.Text = Kolluk

    objDoc.Tables(2).Cell(Row:=1, Column:=6).Range.Text = "Theme 2 No."
    
    'Tabloyu doldur
    Set IlkSiraBul = Range("FG7:FG100000").Find(What:=Cells(ActiveCell.Row, 5).Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    Set SonSiraBul = Range("FH7:FH100000").Find(What:=Cells(ActiveCell.Row, 5).Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    If Not IlkSiraBul Is Nothing Then
        IlkSira = IlkSiraBul.Row
    Else
        MsgBox "The serial number could not be found, so the operation cannot be completed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    If Not SonSiraBul Is Nothing Then
        SonSira = SonSiraBul.Row
    Else
        MsgBox "The serial number could not be found, so the operation cannot be completed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    'Tabloya satır ekle
    x = 0
    If SonSira - IlkSira > 0 Then
        With objDoc.Tables(2)
            For i = IlkSira To SonSira - 1
                .Rows.Add BeforeRow:=objDoc.Tables(2).Rows(2)
                x = x + 1
            Next i
        End With
    End If

    For i = 2 To SonSira - IlkSira + 2
        objDoc.Tables(2).Cell(Row:=i, Column:=1).Range.Text = i - 1 'Tablo sıra no
        objDoc.Tables(2).Cell(Row:=i, Column:=2).Range.Text = Cells(IlkSira + i - 2, 130).Value 'Öğe türü
        objDoc.Tables(2).Cell(Row:=i, Column:=3).Range.Text = Cells(IlkSira + i - 2, 133).Value 'Öğe değeri
        objDoc.Tables(2).Cell(Row:=i, Column:=4).Range.Text = Cells(IlkSira + i - 2, 136).Value 'Adet
        objDoc.Tables(2).Cell(Row:=i, Column:=5).Range.Text = Cells(IlkSira + i - 2, 139).Value 'Öğe ID No
        objDoc.Tables(2).Cell(Row:=i, Column:=6).Range.Text = Cells(IlkSira, 98).Value  'Tema No (Temai her satıra yaz.)
        objDoc.Tables(2).Cell(Row:=i, Column:=7).Range.Text = Cells(IlkSira + i - 2, 142).Value 'Açıklama
    Next i
    
    'TipBler için tutanak2 tutanağında bulunan Öğe ID No kolounu sil.
    objDoc.Tables(2).Columns(5).Delete

    'Tutanak2 tutanağının metin kısmı
    'Tanımlamalar
    If Application.Sum(Range(Cells(IlkSira, 136), Cells(SonSira, 136))) > 1 Then
        Ek4 = "items"
    Else
        Ek4 = "item"
    End If
    Bolum1 = vbTab & "A total of "
    Bolum2 = " " & Ek4 & " , as described above, have been placed inside "
    Bolum3 = " "
    Bolum4 = " and enclosed to be sent to the designated unit mentioned above."

    Ek1 = Application.Sum(Range(Cells(IlkSira, 136), Cells(SonSira, 136)))
    Ek2 = Cells(ActiveCell.Row, 150).Value
    Ek3 = Application.WorksheetFunction.Proper(Cells(ActiveCell.Row, 149).Value)
    Ek3 = Left(Ek3, InStr(Ek3, "/") - 1)
    
    Birlestir = Bolum1 & Ek1 & Bolum2 & Ek2 & Bolum3 & Ek3 & Bolum4
    objDoc.Tables(3).Cell(Row:=1, Column:=1).Range.Text = Birlestir
    
    
    Set MyRange = objDoc.Tables(3).Cell(Row:=1, Column:=1).Range
    MyRange.Find.Execute FindText:=Ek1
    MyRange.Font.Bold = True

    Set MyRange = objDoc.Tables(3).Cell(Row:=1, Column:=1).Range
    With MyRange.Find
        .Execute FindText:=Ek2
        .Execute Forward:=True
    End With
    MyRange.Font.Bold = True
    'Aralıkta bulunan karakterleri bold yap
'    objDoc.Range(objDoc.Tables(3).Cell(Row:=1, Column:=1).Range.Characters(5).Start, _
'    objDoc.Tables(3).Cell(Row:=1, Column:=1).Range.Characters(10).End).Font.Bold = True
    Set MyRange = objDoc.Tables(3).Cell(Row:=1, Column:=1).Range
    MyRange.Find.Execute FindText:=Ek3
    MyRange.Font.Bold = True
    
'    'İmza boşluğunu sayfaya sığdırmak için düzenle
    If SonSira - IlkSira + 1 = 17 Then
        For i = 1 To 2
            objDoc.Tables(4).Rows(1).Delete
        Next i
    ElseIf SonSira - IlkSira + 1 = 18 Then
        For i = 1 To 4
            objDoc.Tables(4).Rows(1).Delete
        Next i
    ElseIf SonSira - IlkSira + 1 = 19 Then
        For i = 1 To 6
            objDoc.Tables(4).Rows(1).Delete
        Next i
    ElseIf SonSira - IlkSira + 1 = 20 Then
        For i = 1 To 7
            objDoc.Tables(4).Rows(1).Delete
        Next i
        objDoc.Tables(4).Rows(1).Height = 5
    End If

    'imzalar
    objDoc.Tables(4).Cell(Row:=6, Column:=2).Range.Text = Cells(ActiveCell.Row, 193).Value 'Ad Soyad1
    objDoc.Tables(4).Cell(Row:=7, Column:=2).Range.Text = Cells(ActiveCell.Row, 194).Value 'Unvan1
    objDoc.Tables(4).Cell(Row:=6, Column:=3).Range.Text = Cells(ActiveCell.Row, 196).Value 'Ad Soyad2
    objDoc.Tables(4).Cell(Row:=7, Column:=3).Range.Text = Cells(ActiveCell.Row, 197).Value 'Unvan2

    'Alt bilgi ekle
    objDoc.Sections(1).Footers(wdHeaderFooterPrimary).Range.Tables(1).Cell(Row:=1, Column:=1).Range.Text = ReNameTutanak2Normal
    
    'Sayfa sayısı kaydet komutuna bağlandı.
'    objDoc.Close SaveChanges:=True
'    objWord.Quit
    objDoc.Save
    'Sayfayı text dosyasından çek
    TxtFileTutanak2 = DestOpUserFolder & "Statement 2 Page Count.txt"
    #If UseADO Then
        Const adReadLine = -2&
        With CreateObject("ADODB.Stream")
            .Open
            .LoadFromFile TxtFileTutanak2
            Do Until .EOS
                TotalSayfaTutanak2 = .ReadText(adReadLine)
            Loop
            .Close
        End With
    #Else 'UseFSO
        Const Tutanak2TristateTrue = -1&
        With CreateObject("Scripting.FileSystemObject")
            With .OpenTextFile(TxtFileTutanak2, Format:=Tutanak2TristateTrue)
                Do Until .AtEndOfStream
                    TotalSayfaTutanak2 = .ReadLine
                Loop
                .Close
            End With
        End With
    #End If
    'MsgBox "Tutanak2: " & TotalSayfaTutanak2

    'Tutanak2 sayfa sayısı
    Cells(ActiveCell.Row, 171).Value = TotalSayfaTutanak2

    objWord.Visible = True
    objWord.Activate

    If TumDoc = True Then
        objWord.Activate
'        For i = 1 To SayPrt
'            objDoc.PrintOut
'        Next i
        objDoc.PrintOut Background:=False, Copies:=SayPrt
        'objWord.Documents.Save
        objDoc.Close SaveChanges:=False
        objWord.Visible = False
    End If

    'Nesneleri temizle
    Set fso = Nothing
    Set objWord = Nothing
    Set objDoc = Nothing

Son:


'Worksheets(5).Protect Password:="123"', DrawingObjects:=False

'Worksheets(5).Activate
Application.ScreenUpdating = True
Application.DisplayAlerts = True
'Application.EnableEvents = True

'End If

End Sub

Sub Rapor3_1UstYaziTipB()

Dim DestOperasyon As String, SourceRapor3_1Normal As String, AutoPath As String, ReNameRapor3_1 As String
Dim fso As Object, objWord As Object, objDoc As Object
Dim OpenKontrolName As String, OpenControl As String, DestOpUserFolder As String, DestOpUserFolderName As String
Dim ContSay As Integer, KontrolFile As String, x As Long, y As Long, i As Long, j As Long
Dim IlkSiraBul As Range, SonSiraBul As Range, IlkSira As Long, SonSira As Long
Dim GelenTema As String, SourceDL As String, ReNameDL As String
Dim TxtFileRapor3_1 As String, TotalSayfaRapor3_1 As String, TxtFileDokum As String, TotalSayfaDokum As String
Dim DokumFileTxt As Object, DokumSayfaGonder As String
Dim SourceRaporNormal As String, SourceTutanak2Normal As String, SourceUstYaziNormal As String
Dim Bolum1 As String, Bolum2 As String, Bolum3 As String, Bolum4 As String, Bolum5 As String, Bolum6 As String
Dim Ek1 As String, Ek2 As String, Ek3 As String, Ek4 As String, Ek5 As String, Ek6 As String, Birlestir As String

Dim ReNameRaporNormal As String, SiraNoIlkSatir As Long, RaporIlgi As String
Dim AltRaporNoIlk As Long, AltRaporNoSon As Long, AdetTopla As Long
Dim TekCogulTipA As String, k As Integer
Dim MyRange As Object, TxtFileRapor As String, TotalSayfaRapor As String
Dim ReNameTutanak2Normal As String, GidenTema As String
Dim TxtFileTutanak2 As String, TotalSayfaTutanak2 As String
Dim ReNameUstYaziNormal As String
Dim TxtFileUstYazi As String, TotalSayfaUstYazi As String
Dim Birimx As String, UserName As String
Dim IlBuyukHarf, IlKucukHarf, IlceBuyukHarf, IlceKucukHarf As String

Dim BStr As Integer
Dim M2 As Boolean, M3 As Boolean, M4 As Boolean
Dim Ifv As Boolean, Ify As Boolean
Dim XXXMudNotu As Boolean
Dim IlgiStrSay As Integer, Govde1StrSay As Integer, IlgiRng As Object, Govde1Rng As Object, Govde2Rng As Object
Dim IlgiFarkSay As Integer, Govde1FarkSay As Integer
Dim ustbilgimuhatap As String, ustbilgitarih As String, ustbilgisayi As String
Dim CokluSayfa As Integer
Dim ItemBul As Range, AdSoyadParaf As String, UnvanParaf As String, TelParaf As String


IlceSakla = ""
If InStr(Cells(ActiveCell.Row, 92).Value, " Organization A") <> 0 Then
    IlceSakla = Cells(ActiveCell.Row, 92).Value
    Cells(ActiveCell.Row, 92).Value = ""
End If

'ÜST YAZI için prosedürü başlat
'    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    'Application.EnableEvents = False
    'Worksheets(5).Unprotect Password:="123"

    'Establish a left-to-right control mechanism for document creation.
    If Cells(ActiveCell.Row, 6).Value = "x" Then
        MsgBox "The Report 3 statement data contains missing or incorrect information, so the operation cannot be started.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    
    If Cells(ActiveCell.Row, 8).Value = "x" Then
        MsgBox "The Statement 2 data contains missing or incorrect information, so the operation cannot be started.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If

    
    'Pathfinder...
    AutoPath = ThisWorkbook.Path
    DestOperasyon = AutoPath & "\System Files\Operation\"

    'ÜST YAZI TANIMLARI
    SourceUstYaziNormal = AutoPath & "\System Files\System Templates\Report 3 Cover Letter Templates\Report 3.1 – Type B Cover Letter.docm"
    'Birden fazla kullanıcının operasyon klasöründe çakışmasını önlemek için operasyon klasörünün içinde yeni klasör aç.
    DestOpUserFolderName = "Operation-" & Left(ThisWorkbook.name, InStr(ThisWorkbook.name, ".") - 1) '& " " & Format(Now(), "ddmmyyyyhhmmss")
    DestOpUserFolder = DestOperasyon & DestOpUserFolderName & "\"

    'Check the "System Files" folder name.
    If Not Dir(AutoPath & "\System Files\", vbDirectory) <> vbNullString Then
        MsgBox "Cannot access the directory " & AutoPath & "\System Files\" & ". The folder named 'System Files' may have been renamed or deleted.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    
    'Check the "Operation" folder name.
    If Not Dir(DestOperasyon, vbDirectory) <> vbNullString Then
        MsgBox "Cannot access the directory " & DestOperasyon & ". The folder named 'Operation' may have been renamed or deleted.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    
    'Check folder names.
    If Not Dir(SourceUstYaziNormal, vbDirectory) <> vbNullString Then
        MsgBox "Cannot access the directory " & SourceUstYaziNormal & ". The names of folders and/or files in this directory may have been changed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If

    'Operation klasörü içinde kullanıcı modülü klasörü yoksa oluştur.
    If Not Dir(DestOpUserFolder, vbDirectory) <> vbNullString Then
        MkDir DestOpUserFolder
    End If

    'On Error Resume Next
    Set IlkSiraBul = Range("FG7:FG100000").Find(What:=Cells(ActiveCell.Row, 5).Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    Set SonSiraBul = Range("FH7:FH100000").Find(What:=Cells(ActiveCell.Row, 5).Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    If Not IlkSiraBul Is Nothing Then
        IlkSira = IlkSiraBul.Row
    Else
        MsgBox "The serial number could not be found, so the operation cannot be completed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    If Not SonSiraBul Is Nothing Then
        SonSira = SonSiraBul.Row
    Else
        MsgBox "The serial number could not be found, so the operation cannot be completed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If

    'Pre-checks for pages to be specified in attachments
    'Report 3 Statement check
    If Cells(IlkSira, 169).Value = "" Then
        MsgBox "Cover letter for sequence number " & Cells(IlkSira, 5).Value & " cannot be prepared because Report 3 Statement has not been created. Please create Report 3 Statement for sequence number " & Cells(IlkSira, 5).Value & " and try again.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    
    'Statement 2 check
    If Cells(IlkSira, 171).Value = "" Then
        MsgBox "Cover letter for sequence number " & Cells(IlkSira, 5).Value & " cannot be prepared because Statement 2 has not been created. Please create Statement 2 for sequence number " & Cells(IlkSira, 5).Value & " and try again.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If


    'PARAF HAZIRLIK İŞLEMİ
    UserName = Environ("UserProfile")
    UserName = UCase(Right(UserName, 7))
    Set ItemBul = Worksheets(2).Range("DR6:DR1000").Find(What:=UserName, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    If Not ItemBul Is Nothing Then
        AdSoyadParaf = "For information only: " & Worksheets(2).Cells(ItemBul.Row, 123).Value
        UnvanParaf = Worksheets(2).Cells(ItemBul.Row, 124).Value
        TelParaf = "Phone / Extension No: " & Worksheets(2).Cells(ItemBul.Row, 126).Value
    Else
        MsgBox UserName & " session is not registered in the system, so the cover letter cannot be created. Please register your session in the system using the Initials Interface located in the Settings Group of the Enterprise Document Automation System menu, then try the operation again.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    
    'On Error Resume Next
    ReNameUstYaziNormal = Cells(ActiveCell.Row, 5).Value & "-" & Cells(6, 10).Value
'________________________________________

    If TumDoc = False Then
        'Close the all Word application
        Call ModuleReport3.OpenWordControl

        'Operation klasöründeki docm uzantılı word dosyalarından açık olanları kapat ve temizle.
        OpenKontrolName = Dir(DestOpUserFolder & "*.docm")
        Do While OpenKontrolName <> ""
            OpenControl = IsFileOpen(DestOpUserFolder & OpenKontrolName)
            If OpenControl = True Then 'Açıksa
                On Error Resume Next
                Set objWord = GetObject(, "Word.Application")
                Set objWord = GetObject(, "Word.Application")
                Set objWord = GetObject(, "Word.Application")
                Set objWord = GetObject(, "Word.Application")
                Set objWord = GetObject(, "Word.Application")
                objWord.Quit SaveChanges:=True
                'MsgBox "Dosya OpenKontrol methodu ile kapatıldı."
            End If
            OpenKontrolName = Dir()
        Loop

        Set objWord = Nothing
        Set objDoc = Nothing

    '________________________________________

        On Error Resume Next
    '    Klasörün içindeki tüm dosyaları sil (txt, docm vb.)
        ContSay = 0
        KontrolFile = Dir(DestOpUserFolder & "*.???")
        Do While KontrolFile <> ""
            ContSay = ContSay + 1
            KontrolFile = Dir()
        Loop
        If ContSay > 0 Then
            Kill DestOpUserFolder & "*.???"
        End If
    End If

    'Dosyayı şablondan operasyon klasörüne kopyala ve adını değiştir.
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.CopyFile (SourceUstYaziNormal), DestOpUserFolder & ReNameUstYaziNormal & ".docm", True
'________________________________________

    'Oluşturulacak dosyayı aç
    On Error Resume Next
    Set objWord = GetObject(, "Word.Application")
    Set objWord = GetObject(, "Word.Application")
    Set objWord = GetObject(, "Word.Application")
    Set objWord = GetObject(, "Word.Application")
    Set objWord = GetObject(, "Word.Application")
    If objWord Is Nothing Then
        'MsgBox "Dosya oluşturmada CreateObject methodu kullanılacak."
        Set objWord = CreateObject("Word.Application")
        objWord.Visible = False
    End If
    objWord.Documents.Open FileName:=DestOpUserFolder & ReNameUstYaziNormal & ".docm"
    objWord.Visible = True
    objWord.Activate 'Ekrana getirir.
    'objDoc.Activate 'Ekrana getirmez.
    objWord.Application.WindowState = wdWindowStateMaximize
    Set objDoc = GetObject(DestOpUserFolder & ReNameUstYaziNormal & ".docm")
'________________________________________

    'Birim
    Birimx = UCase(Replace(Replace(Worksheets(2).Cells(6, 99).Value, "i", "I"), "ı", "I")) & " UNIT"
    objDoc.Tables(1).Cell(Row:=1, Column:=1).Range.Text = Birimx

    Birimx = UCase(Replace(Replace(Worksheets(2).Cells(6, 99).Value, "i", "I"), "ı", "I"))
    'Üst yazı tarihi
    objDoc.Tables(1).Cell(Row:=1, Column:=2).Range.Text = Birimx & ", " & FormatEnglishDate(Cells(ActiveCell.Row, 155).Value) '(Format(Cells(ActiveCell.Row, 155).Value, "d mmmm yyyy"))
    'Yazı no
    objDoc.Tables(1).Cell(Row:=4, Column:=2).Range.Text = Cells(ActiveCell.Row, 156).Value
    'Sigorta türü
    Ek2 = Cells(ActiveCell.Row, 149).Value
    Ek2 = Right(Ek2, Len(Ek2) - InStr(Ek2, "/"))
    objDoc.Tables(1).Cell(Row:=7, Column:=2).Range.Text = Ek2


    'Muhatap
    IlBuyukHarf = UCase(Replace(Replace(Cells(ActiveCell.Row, 91).Value, "i", "I"), "ı", "I"))
    IlKucukHarf = Cells(ActiveCell.Row, 91).Value
    If Cells(ActiveCell.Row, 92).Value <> "" Then
        IlceBuyukHarf = UCase(Replace(Replace(Cells(ActiveCell.Row, 92).Value, "i", "I"), "ı", "I"))
        IlceKucukHarf = Cells(ActiveCell.Row, 92).Value
    Else
        IlceBuyukHarf = ""
        IlceKucukHarf = ""
    End If
    If Cells(ActiveCell.Row, 92).Value <> "" Then
        Bolum3 = IlceKucukHarf & "/" & IlBuyukHarf
    Else
        Bolum3 = IlBuyukHarf
    End If
    

    'YENİ MUHATAP TEMASI
    
    M2 = False
    M3 = False
    M4 = False
'    Ifv = False
'    Ify = False
    'TipB = False
    BStr = 10
    '1. sayfadan sonraki üst bilgi tanımları
    ustbilgitarih = (Format(Cells(ActiveCell.Row, 155).Value, "dd.mm.yyyy"))
    ustbilgisayi = Cells(ActiveCell.Row, 156).Value
    
    '4'lük
    If Cells(ActiveCell.Row, 103).Value <> "" Then
        If Cells(ActiveCell.Row, 102).Value = "Provincial Directorate B" Or Cells(ActiveCell.Row, 102).Value = "Provincial Directorate C" Or _
        Cells(ActiveCell.Row, 102).Value = "Provincial Directorate D" Or Cells(ActiveCell.Row, 102).Value = "Provincial Directorate E" Then 'VALİLİK
            objDoc.Tables(1).Cell(Row:=BStr, Column:=1).Range.Text = IlBuyukHarf & " PROVINCIAL GOVERNORSHIP"
            objDoc.Tables(1).Cell(Row:=BStr + 1, Column:=1).Range.Text = Cells(ActiveCell.Row, 102).Value  'UCase(Replace(Replace(Cells(ActiveCell.Row, 102).Value, "i", "I"), "ı", "I"))
            objDoc.Tables(1).Cell(Row:=BStr + 2, Column:=1).Range.Text = "(" & Cells(ActiveCell.Row, 103).Value & ")"
            objDoc.Tables(1).Cell(Row:=BStr + 3, Column:=1).Range.Text = Bolum3
            objDoc.Tables(1).Cell(Row:=BStr + 3, Column:=1).Range.Font.Underline = wdUnderlineSingle
            M4 = True
            
            ustbilgimuhatap = IlKucukHarf & " Provincial Governorship " & Cells(ActiveCell.Row, 102).Value
        ElseIf Cells(ActiveCell.Row, 102).Value = "District Directorate B" Or Cells(ActiveCell.Row, 102).Value = "District Directorate C" Or _
        Cells(ActiveCell.Row, 102).Value = "District Directorate D" Or Cells(ActiveCell.Row, 102).Value = "District Directorate E" Then 'KAYMAKAMLIK
            objDoc.Tables(1).Cell(Row:=BStr, Column:=1).Range.Text = IlceBuyukHarf & " DISTRICT GOVERNORSHIP"
            objDoc.Tables(1).Cell(Row:=BStr + 1, Column:=1).Range.Text = Cells(ActiveCell.Row, 102).Value 'UCase(Replace(Replace(Cells(ActiveCell.Row, 102).Value, "i", "I"), "ı", "I"))
            objDoc.Tables(1).Cell(Row:=BStr + 2, Column:=1).Range.Text = "(" & Cells(ActiveCell.Row, 103).Value & ")"
            objDoc.Tables(1).Cell(Row:=BStr + 3, Column:=1).Range.Text = Bolum3
            objDoc.Tables(1).Cell(Row:=BStr + 3, Column:=1).Range.Font.Underline = wdUnderlineSingle
            M4 = True
            
            ustbilgimuhatap = IlceKucukHarf & " District Governorship " & Cells(ActiveCell.Row, 102).Value
        ElseIf InStr(Cells(ActiveCell.Row, 102).Value, "General Directorate") <> 0 Or InStr(Cells(ActiveCell.Row, 102).Value, "Regional Directorate") <> 0 Then
            objDoc.Tables(1).Cell(Row:=BStr, Column:=1).Range.Text = UCase(Replace(Replace(ThisWorkbook.Worksheets(2).Cells(6, 111).Value, "i", "I"), "ı", "I"))
            objDoc.Tables(1).Cell(Row:=BStr + 1, Column:=1).Range.Text = Cells(ActiveCell.Row, 102).Value 'UCase(Replace(Replace(Cells(ActiveCell.Row, 102).Value, "i", "I"), "ı", "I"))
            objDoc.Tables(1).Cell(Row:=BStr + 2, Column:=1).Range.Text = "(" & Cells(ActiveCell.Row, 103).Value & ")"
            objDoc.Tables(1).Cell(Row:=BStr + 3, Column:=1).Range.Text = Bolum3
            objDoc.Tables(1).Cell(Row:=BStr + 3, Column:=1).Range.Font.Underline = wdUnderlineSingle
            M4 = True
            
            ustbilgimuhatap = WorksheetFunction.Proper(ThisWorkbook.Worksheets(2).Cells(6, 111).Value) & " " & Cells(ActiveCell.Row, 102).Value
        Else 'YARGI 3'lük
            Bolum1 = UCase(Replace(Replace(Cells(ActiveCell.Row, 102).Value, "i", "I"), "ı", "I"))
            If InStr(Bolum1, "X.X. ") > 0 Then
                objDoc.Tables(1).Cell(Row:=BStr, Column:=1).Range.Text = Mid(Bolum1, 6, Len(Bolum1))
                objDoc.Tables(1).Cell(Row:=BStr + 1, Column:=1).Range.Text = "(" & Cells(ActiveCell.Row, 103).Value & ")"
                objDoc.Tables(1).Cell(Row:=BStr + 2, Column:=1).Range.Text = Bolum3
                objDoc.Tables(1).Cell(Row:=BStr + 2, Column:=1).Range.Font.Underline = wdUnderlineSingle
                M3 = True
                
                ustbilgimuhatap = Mid(Cells(ActiveCell.Row, 102).Value, 6, Len(Cells(ActiveCell.Row, 102).Value))
            Else
                objDoc.Tables(1).Cell(Row:=BStr, Column:=1).Range.Text = Bolum1
                objDoc.Tables(1).Cell(Row:=BStr + 1, Column:=1).Range.Text = "(" & Cells(ActiveCell.Row, 103).Value & ")"
                objDoc.Tables(1).Cell(Row:=BStr + 2, Column:=1).Range.Text = Bolum3
                objDoc.Tables(1).Cell(Row:=BStr + 2, Column:=1).Range.Font.Underline = wdUnderlineSingle
                M3 = True

                ustbilgimuhatap = Cells(ActiveCell.Row, 102).Value
            End If
        End If
    End If
    '3'lük
    If Cells(ActiveCell.Row, 103).Value = "" Then
        If Cells(ActiveCell.Row, 102).Value = "Provincial Directorate B" Or Cells(ActiveCell.Row, 102).Value = "Provincial Directorate C" Or _
        Cells(ActiveCell.Row, 102).Value = "Provincial Directorate D" Or Cells(ActiveCell.Row, 102).Value = "Provincial Directorate E" Then 'VALİLİK
            objDoc.Tables(1).Cell(Row:=BStr, Column:=1).Range.Text = IlBuyukHarf & " PROVINCIAL GOVERNORSHIP"
            objDoc.Tables(1).Cell(Row:=BStr + 1, Column:=1).Range.Text = "(" & Cells(ActiveCell.Row, 102).Value & ")" 'UCase(Replace(Replace(Cells(ActiveCell.Row, 102).Value, "i", "I"), "ı", "I"))
            objDoc.Tables(1).Cell(Row:=BStr + 2, Column:=1).Range.Text = Bolum3
            objDoc.Tables(1).Cell(Row:=BStr + 2, Column:=1).Range.Font.Underline = wdUnderlineSingle
            M3 = True
            
            ustbilgimuhatap = IlKucukHarf & " Provincial Governorship " & Cells(ActiveCell.Row, 102).Value
        ElseIf Cells(ActiveCell.Row, 102).Value = "District Directorate B" Or Cells(ActiveCell.Row, 102).Value = "District Directorate C" Or _
        Cells(ActiveCell.Row, 102).Value = "District Directorate D" Or Cells(ActiveCell.Row, 102).Value = "District Directorate E" Then
            objDoc.Tables(1).Cell(Row:=BStr, Column:=1).Range.Text = IlceBuyukHarf & " DISTRICT GOVERNORSHIP"
            objDoc.Tables(1).Cell(Row:=BStr + 1, Column:=1).Range.Text = "(" & Cells(ActiveCell.Row, 102).Value & ")" 'UCase(Replace(Replace(Cells(ActiveCell.Row, 102).Value, "i", "I"), "ı", "I"))
            objDoc.Tables(1).Cell(Row:=BStr + 2, Column:=1).Range.Text = Bolum3
            objDoc.Tables(1).Cell(Row:=BStr + 2, Column:=1).Range.Font.Underline = wdUnderlineSingle
            M3 = True
            
            ustbilgimuhatap = IlceKucukHarf & " District Governorship " & Cells(ActiveCell.Row, 102).Value
        ElseIf InStr(Cells(ActiveCell.Row, 102).Value, "General Directorate") <> 0 Or InStr(Cells(ActiveCell.Row, 102).Value, "Regional Directorate") <> 0 Then
            objDoc.Tables(1).Cell(Row:=BStr, Column:=1).Range.Text = UCase(Replace(Replace(ThisWorkbook.Worksheets(2).Cells(6, 111).Value, "i", "I"), "ı", "I"))
            objDoc.Tables(1).Cell(Row:=BStr + 1, Column:=1).Range.Text = "(" & Cells(ActiveCell.Row, 102).Value & ")"  'UCase(Replace(Replace(Cells(ActiveCell.Row, 102).Value, "i", "I"), "ı", "I"))
            objDoc.Tables(1).Cell(Row:=BStr + 2, Column:=1).Range.Text = Bolum3
            objDoc.Tables(1).Cell(Row:=BStr + 2, Column:=1).Range.Font.Underline = wdUnderlineSingle
            M3 = True

            ustbilgimuhatap = WorksheetFunction.Proper(ThisWorkbook.Worksheets(2).Cells(6, 111).Value) & " " & Cells(ActiveCell.Row, 102).Value
        Else 'YARGI 2'lik
            Bolum1 = UCase(Replace(Replace(Cells(ActiveCell.Row, 102).Value, "i", "I"), "ı", "I"))
            If InStr(Bolum1, "X.X. ") > 0 Then
                objDoc.Tables(1).Cell(Row:=BStr, Column:=1).Range.Text = Mid(Bolum1, 6, Len(Bolum1))
                objDoc.Tables(1).Cell(Row:=BStr + 1, Column:=1).Range.Text = Bolum3
                objDoc.Tables(1).Cell(Row:=BStr + 1, Column:=1).Range.Font.Underline = wdUnderlineSingle
                M2 = True

                ustbilgimuhatap = Mid(Cells(ActiveCell.Row, 102).Value, 6, Len(Cells(ActiveCell.Row, 102).Value))
            Else
                objDoc.Tables(1).Cell(Row:=BStr, Column:=1).Range.Text = Bolum1
                objDoc.Tables(1).Cell(Row:=BStr + 1, Column:=1).Range.Text = Bolum3
                objDoc.Tables(1).Cell(Row:=BStr + 1, Column:=1).Range.Font.Underline = wdUnderlineSingle
                M2 = True

                ustbilgimuhatap = Cells(ActiveCell.Row, 102).Value
            End If
        End If
    End If

    
    'Üst yazı gövde metni ve ek senaryoları
    If Cells(IlkSira, 124).Value = "Yes" Then 'Organization B mensubu var.
        If Mid(Cells(IlkSira, 128).Value, 7, 2) = "31" Then 'TipAyı rızası ile verdi.
            'tipAnın/tipAların (TipA adedi)
            AdetTopla = Application.Sum(Range(Cells(IlkSira, 136), Cells(SonSira, 136)))
            If AdetTopla = 1 Then
                Ek1 = "tipBnin"
            ElseIf AdetTopla > 1 Then
                Ek1 = "tipBlerin"
            End If
            Ek2 = Cells(ActiveCell.Row, 95).Value 'Tutanak tarihi
            Ek3 = Cells(ActiveCell.Row, 109).Value 'Ad Soyad
            Ek4 = Cells(ActiveCell.Row, 98).Value 'Temax no
            Bolum1 = "Birimimize "
            'Kişinin açık kimliği var/yok
            If Cells(ActiveCell.Row, 126).Value <> "" And Cells(ActiveCell.Row, 126).Value > 0 Then
                Bolum2 = " tarihinde ilişik tutanakta xxxxx xxxxx xx xxxxx yazılı " 'var
            Else
                Bolum2 = " tarihinde " 'yok
            End If
            Bolum3 = " isimli kişinin getirmiş olduğu öğe türü, öğe değeri, adedi ilişik tutanakta belirtilen ve tarafımızca "
            Bolum4 = " tema numarası verilen "
            Bolum5 = " invalid olduğu değerlendirilmiştir."
            Bolum6 = Chr(9) & "Bilgilerinize arz ederiz."
            Birlestir = Chr(9) & Bolum1 & Ek2 & Bolum2 & Ek3 & Bolum3 & Ek4 & Bolum4 & Ek1 & Bolum5 '& vbNewLine & Chr(9) & Ek6
            objDoc.Tables(2).Cell(Row:=1, Column:=1).Range.Text = Birlestir
            objDoc.Tables(2).Cell(Row:=3, Column:=1).Range.Text = Bolum6
            'Ekler
            objDoc.Tables(3).Cell(Row:=8, Column:=1).Range.Text = "Attachment:"
            Ek2 = Cells(ActiveCell.Row, 149).Value 'Kapalı Package A
            Ek2 = Left(Ek2, InStr(Ek2, "/") - 1)
            objDoc.Tables(3).Cell(Row:=8, Column:=2).Range.Text = " 1) Kapalı " & Ek2 & " (" & Cells(ActiveCell.Row, 150).Value & " adet)"
            x = Application.Sum(Range(Cells(IlkSira, 171), Cells(SonSira, 171))) 'Statement 2 toplam sayfa sayısı
            objDoc.Tables(3).Cell(Row:=9, Column:=2).Range.Text = " 2) Statement 2 (" & x & " page(s))"
            If Cells(ActiveCell.Row, 126).Value <> "" And Cells(ActiveCell.Row, 126).Value > 0 Then 'Kimlik var
                x = Application.Sum(Range(Cells(IlkSira, 169), Cells(SonSira, 169))) + Cells(ActiveCell.Row, 126).Value 'Rapor3 Tutanağı
                objDoc.Tables(3).Cell(Row:=10, Column:=2).Range.Text = " 3) Ekli Tutanak (Toplam " & x & " page(s))"
            Else 'Kimlik yok
                x = Application.Sum(Range(Cells(IlkSira, 169), Cells(SonSira, 169))) 'Rapor3 Tutanağı
                objDoc.Tables(3).Cell(Row:=10, Column:=2).Range.Text = " 3) Tutanak (" & x & " page(s))"
            End If
        ElseIf Mid(Cells(IlkSira, 128).Value, 13, 2) = "51" Then 'tipAya Organization B mensubu el koydu.
            'tipAnın/tipAların (TipA adedi)
            AdetTopla = Application.Sum(Range(Cells(IlkSira, 136), Cells(SonSira, 136)))
            If AdetTopla = 1 Then
                Ek1 = "tipBnin"
                Ek5 = "tipBye"
            ElseIf AdetTopla > 1 Then
                Ek1 = "tipBlerin"
                Ek5 = "tipBlere"
            End If
            Ek2 = Cells(ActiveCell.Row, 95).Value 'Tutanak tarihi
            Ek3 = Cells(ActiveCell.Row, 109).Value 'Ad Soyad
            Ek4 = Cells(ActiveCell.Row, 98).Value 'Tema no
            Bolum1 = "Birimimize "
            'Kişinin açık kimliği var/yok
            If Cells(ActiveCell.Row, 126).Value <> "" And Cells(ActiveCell.Row, 126).Value > 0 Then
                Bolum2 = " tarihinde ilişik tutanakta xxxxx xxxxx xx xxxxx yazılı " 'var
            Else
                Bolum2 = " tarihinde " 'yok
            End If
            Bolum3 = " isimli kişinin getirmiş olduğu öğe türü, öğe değeri, adedi ilişik tutanakta belirtilen ve tarafımızca "
            Bolum4 = " tema numarası verilen "
            Bolum5 = " invalid olduğu değerlendirilmiş ve " & Ek5 & " " & Cells(ActiveCell.Row, 125).Value & " tarafından el konulmuştur."
            Bolum6 = Chr(9) & "Bilgilerinize arz ederiz."
            Birlestir = Chr(9) & Bolum1 & Ek2 & Bolum2 & Ek3 & Bolum3 & Ek4 & Bolum4 & Ek1 & Bolum5 '& vbNewLine & Chr(9) & Ek6
            objDoc.Tables(2).Cell(Row:=1, Column:=1).Range.Text = Birlestir
            objDoc.Tables(2).Cell(Row:=3, Column:=1).Range.Text = Bolum6
            'Ekler
            objDoc.Tables(3).Cell(Row:=8, Column:=1).Range.Text = "Attachment:"
            If Cells(ActiveCell.Row, 126).Value <> "" And Cells(ActiveCell.Row, 126).Value > 0 Then 'Kimlik var
                x = Application.Sum(Range(Cells(IlkSira, 169), Cells(SonSira, 169))) + Cells(ActiveCell.Row, 126).Value 'Rapor3 Tutanağı
                objDoc.Tables(3).Cell(Row:=8, Column:=2).Range.Text = " Ekli Tutanak (Toplam " & x & " page(s))"
            Else 'Kimlik yok
                x = Application.Sum(Range(Cells(IlkSira, 169), Cells(SonSira, 169))) 'Rapor3 Tutanağı
                objDoc.Tables(3).Cell(Row:=8, Column:=2).Range.Text = " Tutanak (" & x & " page(s))"
            End If
        End If
    Else 'Organization B mensubu yok.
        If Mid(Cells(IlkSira, 128).Value, 7, 2) = "31" Then 'TipAyı rızası ile verdi.
            'tipAnın/tipAların (TipA adedi)
            AdetTopla = Application.Sum(Range(Cells(IlkSira, 136), Cells(SonSira, 136)))
            If AdetTopla = 1 Then
                Ek1 = "tipBnin"
            ElseIf AdetTopla > 1 Then
                Ek1 = "tipBlerin"
            End If
            Ek2 = Cells(ActiveCell.Row, 95).Value 'Tutanak tarihi
            Ek3 = Cells(ActiveCell.Row, 109).Value 'Ad Soyad
            Ek4 = Cells(ActiveCell.Row, 98).Value 'Tema no
            Bolum1 = "Birimimize "
            'Kişinin açık kimliği var/yok
            If Cells(ActiveCell.Row, 126).Value <> "" And Cells(ActiveCell.Row, 126).Value > 0 Then
                Bolum2 = " tarihinde ilişik tutanakta xxxxx xxxxx xx xxxxx yazılı " 'var
            Else
                Bolum2 = " tarihinde " 'yok
            End If
            Bolum3 = " isimli kişinin getirmiş olduğu öğe türü, öğe değeri, adedi ilişik tutanakta belirtilen ve tarafımızca "
            Bolum4 = " tema numarası verilen "
            Bolum5 = " invalid olduğu değerlendirilmiştir."
            Bolum6 = Chr(9) & "Bilgilerinize arz ederiz."
            Birlestir = Chr(9) & Bolum1 & Ek2 & Bolum2 & Ek3 & Bolum3 & Ek4 & Bolum4 & Ek1 & Bolum5 '& vbNewLine & Chr(9) & Ek6
            objDoc.Tables(2).Cell(Row:=1, Column:=1).Range.Text = Birlestir
            objDoc.Tables(2).Cell(Row:=3, Column:=1).Range.Text = Bolum6
            'Ekler
            objDoc.Tables(3).Cell(Row:=8, Column:=1).Range.Text = "Attachment:"
            Ek2 = Cells(ActiveCell.Row, 149).Value 'Kapalı Package A
            Ek2 = Left(Ek2, InStr(Ek2, "/") - 1)
            objDoc.Tables(3).Cell(Row:=8, Column:=2).Range.Text = " 1) Kapalı " & Ek2 & " (" & Cells(ActiveCell.Row, 150).Value & " adet)"
            x = Application.Sum(Range(Cells(IlkSira, 171), Cells(SonSira, 171))) 'Statement 2 toplam sayfa sayısı
            objDoc.Tables(3).Cell(Row:=9, Column:=2).Range.Text = " 2) Statement 2 (" & x & " page(s))"
            If Cells(ActiveCell.Row, 126).Value <> "" And Cells(ActiveCell.Row, 126).Value > 0 Then 'Kimlik var
                x = Application.Sum(Range(Cells(IlkSira, 169), Cells(SonSira, 169))) + Cells(ActiveCell.Row, 126).Value 'Rapor3 Tutanağı
                objDoc.Tables(3).Cell(Row:=10, Column:=2).Range.Text = " 3) Ekli Tutanak (Toplam " & x & " page(s))"
            Else 'Kimlik yok
                x = Application.Sum(Range(Cells(IlkSira, 169), Cells(SonSira, 169))) 'Rapor3 Tutanağı
                objDoc.Tables(3).Cell(Row:=10, Column:=2).Range.Text = " 3) Tutanak (" & x & " page(s))"
            End If
        ElseIf Mid(Cells(IlkSira, 128).Value, 10, 2) = "41" Then 'TipAyı rızası ile vermedi.
            'tipAnın/tipAların (TipA adedi)
            AdetTopla = Application.Sum(Range(Cells(IlkSira, 136), Cells(SonSira, 136)))
            If AdetTopla = 1 Then
                Ek1 = "tipBnin"
                Ek5 = "tipByi"
            ElseIf AdetTopla > 1 Then
                Ek1 = "tipBlerin"
                Ek5 = "tipBleri"
            End If
            Ek2 = Cells(ActiveCell.Row, 95).Value 'Tutanak tarihi
            Ek3 = Cells(ActiveCell.Row, 109).Value 'Ad Soyad
            Ek4 = Cells(ActiveCell.Row, 98).Value 'Tema no
            Bolum1 = "Birimimize "
            'Kişinin açık kimliği var/yok
            If Cells(ActiveCell.Row, 126).Value <> "" And Cells(ActiveCell.Row, 126).Value > 0 Then
                Bolum2 = " tarihinde ilişik tutanakta xxxxx xxxxx xx xxxxx yazılı " 'var
            Else
                Bolum2 = " tarihinde " 'yok
            End If
            Bolum3 = " isimli kişinin getirmiş olduğu öğe türü, öğe değeri, adedi ilişik tutanakta belirtilen " & Ek1 & " geçersizliğinden şüphe edilmiştir."
            Bolum4 = " Ancak gerekli hatırlatma yapıldığı halde, invalid olduğu değerlendirilen "
            Bolum5 = " getiren kişi rızası ile vermemiştir."
            Bolum6 = Chr(9) & "Bilgilerinize arz ederiz."
            Birlestir = Chr(9) & Bolum1 & Ek2 & Bolum2 & Ek3 & Bolum3 & Bolum4 & Ek5 & Bolum5 '& vbNewLine & Chr(9) & Ek6
            objDoc.Tables(2).Cell(Row:=1, Column:=1).Range.Text = Birlestir
            objDoc.Tables(2).Cell(Row:=3, Column:=1).Range.Text = Bolum6
            'Ekler
            objDoc.Tables(3).Cell(Row:=8, Column:=1).Range.Text = "Attachment:"
            If Cells(ActiveCell.Row, 126).Value <> "" And Cells(ActiveCell.Row, 126).Value > 0 Then 'Kimlik var
                x = Application.Sum(Range(Cells(IlkSira, 169), Cells(SonSira, 169))) + Cells(ActiveCell.Row, 126).Value 'Rapor3 Tutanağı
                objDoc.Tables(3).Cell(Row:=8, Column:=2).Range.Text = " Ekli Tutanak (Toplam " & x & " page(s))"
            Else 'Kimlik yok
                x = Application.Sum(Range(Cells(IlkSira, 169), Cells(SonSira, 169))) 'Rapor3 Tutanağı
                objDoc.Tables(3).Cell(Row:=8, Column:=2).Range.Text = " Tutanak (" & x & " page(s))"
            End If
        End If
    End If

    'imzalar (Tüm senaryolar için geçerli.)
    objDoc.Tables(3).Cell(Row:=4, Column:=2).Range.Text = Cells(ActiveCell.Row, 205).Value 'Ad Soyad1
    objDoc.Tables(3).Cell(Row:=5, Column:=2).Range.Text = Cells(ActiveCell.Row, 206).Value 'Unvan1
    objDoc.Tables(3).Cell(Row:=4, Column:=3).Range.Text = Cells(ActiveCell.Row, 208).Value 'Ad Soyad2
    objDoc.Tables(3).Cell(Row:=5, Column:=3).Range.Text = Cells(ActiveCell.Row, 209).Value 'Unvan2


    '__________________________DİNAMİK SAYFA YAPISI
    
    'İlgi ve gövde metninin kaç satırdan oluştuğu bilgisinin elde edilmesinde kullanılıyor.
    'Set IlgiRng = objDoc.Tables(1).Cell(Row:=15, Column:=3).Range
    Set Govde1Rng = objDoc.Tables(2).Cell(Row:=1, Column:=1).Range
    Set Govde2Rng = objDoc.Tables(2).Cell(Row:=2, Column:=1).Range
    
    'IlgiStrSay = Govde1Rng.Information(wdFirstCharacterLineNumber) - IlgiRng.Information(wdFirstCharacterLineNumber) - 1
    Govde1StrSay = Govde2Rng.Information(wdFirstCharacterLineNumber) - Govde1Rng.Information(wdFirstCharacterLineNumber)
    'MsgBox "İlgi Satır Sayısı: " & IlgiStrSay & vbNewLine & vbNewLine & "Gövde 1 Satır Sayısı: " & Govde1StrSay
    'IlgiFarkSay = IlgiStrSay - 1
    IlgiFarkSay = 0
    Govde1FarkSay = Govde1StrSay '- 3 'Varsayılan 1 paragrafta (1 rowda) toplam 3 satıra göre sıfırlandı.
    CokluSayfa = 0
    'MsgBox Govde1FarkSay
 
   
    'Dinamik sayfa düzeni
    objDoc.Tables(3).Rows(12).Delete 'Ek sonrası
    'Dinamik sayfa düzeni
    If Mid(Cells(IlkSira, 128).Value, 7, 2) = "31" Then 'TipAyı rızası ile verdi. ' 3 Ek var
        For i = 1 To 2
            objDoc.Tables(3).Rows(11).Delete 'Ek sonrası
        Next i
    Else 'Tek ek var
        For i = 1 To 4
            objDoc.Tables(3).Rows(9).Delete 'Ek sonrası
        Next i
        Govde1FarkSay = Govde1FarkSay - 2
    End If

    If M4 = True Then
        For i = 1 To 2
            objDoc.Tables(1).Rows(14).Delete 'Muhatap sonrası
        Next i
        If IlgiFarkSay + Govde1FarkSay < 17 Then
            Govde1FarkSay = Govde1FarkSay + 2
            'MsgBox IlgiFarkSay + Govde1FarkSay
            If IlgiFarkSay + Govde1FarkSay > 8 Then
                For i = 1 To IlgiFarkSay + Govde1FarkSay
                    If i = 9 Then
                        objDoc.Tables(3).Rows(6).Delete 'Ek öncesi
                    ElseIf i = 10 Then
                        objDoc.Tables(3).Rows(6).Delete 'Ek öncesi
                    ElseIf i = 11 Then
                        objDoc.Tables(3).Rows(1).Delete 'İmza öncesi
                    ElseIf i = 12 Then
                        objDoc.Tables(3).Rows(1).Delete 'İmza öncesi
                    ElseIf i = 13 Then
                        objDoc.Tables(1).Rows(8).Delete 'Muhatap öncesi
                    ElseIf i = 14 Then
                        objDoc.Tables(1).Rows(8).Delete 'Muhatap öncesi
                    End If
                Next i
            Else
                If (IlgiFarkSay + Govde1FarkSay) < 8 Then
                    For i = 1 To 8 - (IlgiFarkSay + Govde1FarkSay)
                        'Boş satır ekle
                        With objDoc.Tables(3)
                            .Rows.Add BeforeRow:=objDoc.Tables(3).Rows(8)
                        End With
                    Next i
                End If
            End If
        Else
            CokluSayfa = 1
        End If
    ElseIf M3 = True Then
        For i = 1 To 3
            objDoc.Tables(1).Rows(13).Delete 'Muhatap sonrası
        Next i
        If IlgiFarkSay + Govde1FarkSay < 16 Then
            Govde1FarkSay = Govde1FarkSay + 1
            'MsgBox IlgiFarkSay + Govde1FarkSay
            If IlgiFarkSay + Govde1FarkSay > 8 Then
                For i = 9 To IlgiFarkSay + Govde1FarkSay
                    If i = 9 Then
                        objDoc.Tables(3).Rows(6).Delete 'Ek öncesi
                    ElseIf i = 10 Then
                        objDoc.Tables(3).Rows(6).Delete 'Ek öncesi
                    ElseIf i = 11 Then
                        objDoc.Tables(3).Rows(1).Delete 'İmza öncesi
                    ElseIf i = 12 Then
                        objDoc.Tables(3).Rows(1).Delete 'İmza öncesi
                    ElseIf i = 13 Then
                        objDoc.Tables(1).Rows(8).Delete 'Muhatap öncesi
                    ElseIf i = 14 Then
                        objDoc.Tables(1).Rows(8).Delete 'Muhatap öncesi
                    End If
                Next i
            Else
                If (IlgiFarkSay + Govde1FarkSay) < 8 Then
                    For i = 1 To 8 - (IlgiFarkSay + Govde1FarkSay)
                        'Boş satır ekle
                        With objDoc.Tables(3)
                            .Rows.Add BeforeRow:=objDoc.Tables(3).Rows(8)
                        End With
                    Next i
                End If
            End If
        Else
            CokluSayfa = 1
        End If
    ElseIf M2 = True Then
        For i = 1 To 4
            objDoc.Tables(1).Rows(12).Delete 'Muhatap sonrası
        Next i

        If IlgiFarkSay + Govde1FarkSay < 15 Then
            'Govde1FarkSay = Govde1FarkSay + 0
            'MsgBox IlgiFarkSay + Govde1FarkSay
            If IlgiFarkSay + Govde1FarkSay > 8 Then
                For i = 9 To IlgiFarkSay + Govde1FarkSay
                    If i = 9 Then
                        objDoc.Tables(3).Rows(6).Delete 'Ek öncesi
                    ElseIf i = 10 Then
                        objDoc.Tables(3).Rows(6).Delete 'Ek öncesi
                    ElseIf i = 11 Then
                        objDoc.Tables(3).Rows(1).Delete 'İmza öncesi
                    ElseIf i = 12 Then
                        objDoc.Tables(3).Rows(1).Delete 'İmza öncesi
                    ElseIf i = 13 Then
                        objDoc.Tables(1).Rows(8).Delete 'Muhatap öncesi
                    ElseIf i = 14 Then
                        objDoc.Tables(1).Rows(8).Delete 'Muhatap öncesi
                    End If
                Next i
            Else
                If (IlgiFarkSay + Govde1FarkSay) < 8 Then
                    For i = 1 To 8 - (IlgiFarkSay + Govde1FarkSay)
                        'Boş satır ekle
                        With objDoc.Tables(3)
                            .Rows.Add BeforeRow:=objDoc.Tables(3).Rows(8)
                        End With
                    Next i
                End If
            End If
        Else
            CokluSayfa = 1
        End If
    End If

    objDoc.Paragraphs(objDoc.Paragraphs.Count).Range.Font.Size = 1

    
    'Footer
    If objDoc.ActiveWindow.View.SplitSpecial <> wdPaneNone Then
        objDoc.ActiveWindow.Panes(2).Close
    End If
    If objDoc.ActiveWindow.ActivePane.View.Type = wdNormalView Or objDoc.ActiveWindow.ActivePane.View.Type = wdOutlineView Then
        objDoc.ActiveWindow.ActivePane.View.Type = wdPrintView
    End If

    With objDoc
        With .Sections(1).Headers(wdHeaderFooterPrimary).Range
            .InsertAfter Text:="Page "
            .Fields.Add Range:=.Characters.Last, Type:=wdFieldEmpty, Text:="PAGE", PreserveFormatting:=False
            .InsertAfter Text:=" of the letter dated " & ustbilgitarih & _
                             ", reference number " & ustbilgisayi & _
                             ", addressed to the " & ustbilgimuhatap
        End With
    End With
    
    objDoc.ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument


    'PARAF EKLEME İŞLEMİ
    objDoc.Sections(1).Footers(wdHeaderFooterFirstPage).Range.Tables(1).Cell(Row:=1, Column:=2).Range.Text = AdSoyadParaf & vbNewLine & _
                                                                                       UnvanParaf & vbNewLine & _
                                                                                       TelParaf
    objDoc.Sections(1).Footers(wdHeaderFooterPrimary).Range.Tables(1).Cell(Row:=1, Column:=2).Range.Text = AdSoyadParaf & vbNewLine & _
                                                                                       UnvanParaf & vbNewLine & _
                                                                                       TelParaf

    
Tekrarla:
    'Sayfa sayısı kaydet komutuna bağlandı.
    objDoc.Save
    'Sayfayı text dosyasından çek
    TxtFileUstYazi = DestOpUserFolder & "Cover Letter Page Count.txt"
    #If UseADO Then
        Const adReadLine = -2&
        With CreateObject("ADODB.Stream")
            .Open
            .LoadFromFile TxtFileUstYazi
            Do Until .EOS
                TotalSayfaUstYazi = .ReadText(adReadLine)
            Loop
            .Close
        End With
    #Else 'UseFSO
        Const UstYaziTristateTrue = -1&
        With CreateObject("Scripting.FileSystemObject")
            With .OpenTextFile(TxtFileUstYazi, Format:=UstYaziTristateTrue)
                Do Until .AtEndOfStream
                    TotalSayfaUstYazi = .ReadLine
                Loop
                .Close
            End With
        End With
    #End If
    'MsgBox "Üst Yazı: " & TotalSayfaUstYazi

    'Üst Yazı sayfa sayısı
    Cells(ActiveCell.Row, 173).Value = TotalSayfaUstYazi

    objWord.Visible = True
    objWord.Activate

    If TumDoc = True Then
        objWord.Activate
'        For i = 1 To SayPrt
'            objDoc.PrintOut
'        Next i
        objDoc.PrintOut Background:=False, Copies:=SayPrt
        'objWord.Documents.Save
        objDoc.Close SaveChanges:=False
        objWord.Visible = False
    End If

    'Nesneleri temizle
    Set fso = Nothing
    Set objWord = Nothing
    Set objDoc = Nothing

Son:

'Worksheets(5).Protect Password:="123"', DrawingObjects:=False


If IlceSakla <> "" Then
    Cells(ActiveCell.Row, 92).Value = IlceSakla
End If

'Worksheets(5).Activate
Application.ScreenUpdating = True
Application.DisplayAlerts = True
'Application.EnableEvents = True

'End If

End Sub

Sub Rapor3_2Tutanak()

Dim DestOperasyon As String, SourceRapor3_1Normal As String, AutoPath As String, ReNameRapor3_1 As String
Dim fso As Object, objWord As Object, objDoc As Object
Dim OpenKontrolName As String, OpenControl As String, DestOpUserFolder As String, DestOpUserFolderName As String
Dim ContSay As Integer, KontrolFile As String, x As Long, y As Long, i As Long, j As Long
Dim IlkSiraBul As Range, SonSiraBul As Range, IlkSira As Long, SonSira As Long
Dim GelenTema As String, SourceDL As String, ReNameDL As String
Dim TxtFileRapor3_1 As String, TotalSayfaRapor3_1 As String, TxtFileDokum As String, TotalSayfaDokum As String
Dim DokumFileTxt As Object, DokumSayfaGonder As String
Dim SourceRaporNormal As String, SourceTutanak2Normal As String, SourceUstYaziNormal As String
Dim Bolum1 As String, Bolum2 As String, Bolum3 As String, Bolum4 As String, Bolum5 As String, Bolum6 As String
Dim Ek1 As String, Ek2 As String, Ek3 As String, Ek4 As String, Ek5 As String, Ek6 As String, Birlestir As String

Dim ReNameRaporNormal As String, SiraNoIlkSatir As Long, RaporIlgi As String
Dim AltRaporNoIlk As Long, AltRaporNoSon As Long, AdetTopla As Long
Dim TekCogulTipA As String, k As Integer
Dim MyRange As Object, TxtFileRapor As String, TotalSayfaRapor As String
Dim ReNameTutanak2Normal As String, GidenTema As String
Dim TxtFileTutanak2 As String, TotalSayfaTutanak2 As String
Dim ReNameUstYaziNormal As String, PaketTipi As String
Dim TxtFileUstYazi As String, TotalSayfaUstYazi As String
Dim Birimx As String, UserName As String, SourceRapor3_1Farkli As String
Dim Kolluk As String


'TUTANAK1 için prosedürü başlat
'    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    'Application.EnableEvents = False
    'Worksheets(5).Unprotect Password:="123"
    
    'Pathfinder...
    AutoPath = ThisWorkbook.Path
    DestOperasyon = AutoPath & "\System Files\Operation\"
    'TUTANAK1 TANIMLARI
    SourceRapor3_1Normal = AutoPath & "\System Files\System Templates\Report 3 Statements\Report 3.2 – Type A.docm"
    
    'Birden fazla kullanıcının operasyon klasöründe çakışmasını önlemek için operasyon klasörünün içinde yeni klasör aç.
    DestOpUserFolderName = "Operation-" & Left(ThisWorkbook.name, InStr(ThisWorkbook.name, ".") - 1) '& " " & Format(Now(), "ddmmyyyyhhmmss")
    DestOpUserFolder = DestOperasyon & DestOpUserFolderName & "\"
    
    'Check the "System Files" folder name.
    If Not Dir(AutoPath & "\System Files\", vbDirectory) <> vbNullString Then
        MsgBox "Cannot access the directory " & AutoPath & "\System Files\" & ". The folder named 'System Files' may have been renamed or deleted.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    
    'Check the "Operation" folder name.
    If Not Dir(DestOperasyon, vbDirectory) <> vbNullString Then
        MsgBox "Cannot access the directory " & DestOperasyon & ". The folder named 'Operation' may have been renamed or deleted.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If

    'Check folder names.
    If Not Dir(SourceRapor3_1Normal, vbDirectory) <> vbNullString Then
        MsgBox "Cannot access the directory " & SourceRapor3_1Normal & ". The names of folders and/or files in this directory may have been changed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If

    'Operation klasörü içinde kullanıcı modülü klasörü yoksa oluştur.
    If Not Dir(DestOpUserFolder, vbDirectory) <> vbNullString Then
        MkDir DestOpUserFolder
    End If
    
    'On Error Resume Next
    ReNameRapor3_1 = Cells(ActiveCell.Row, 5).Value & "-" & Cells(6, 6).Value
    'ReNameDL = Cells(ActiveCell.Row, 5).Value & "-" & "Dispatch List"
'________________________________________

    
    'Close the all Word application
    Call ModuleReport3.OpenWordControl
    
    'Operation klasöründeki docm uzantılı word dosyalarından açık olanları kapat ve temizle.
    OpenKontrolName = Dir(DestOpUserFolder & "*.docm")
    Do While OpenKontrolName <> ""
        OpenControl = IsFileOpen(DestOpUserFolder & OpenKontrolName)
        If OpenControl = True Then 'Açıksa
            On Error Resume Next
            Set objWord = GetObject(, "Word.Application")
            Set objWord = GetObject(, "Word.Application")
            Set objWord = GetObject(, "Word.Application")
            Set objWord = GetObject(, "Word.Application")
            Set objWord = GetObject(, "Word.Application")
            objWord.Quit SaveChanges:=True
            'MsgBox "Dosya OpenKontrol methodu ile kapatıldı."

        End If
        OpenKontrolName = Dir()
    Loop

'    Call ModuleReport3.OpenWordControl
     
    Set objWord = Nothing
    Set objDoc = Nothing

'________________________________________

    On Error Resume Next
'    Klasörün içindeki tüm dosyaları sil (txt, docm vb.)
    ContSay = 0
    KontrolFile = Dir(DestOpUserFolder & "*.???")
    Do While KontrolFile <> ""
        ContSay = ContSay + 1
        KontrolFile = Dir()
    Loop
    If ContSay > 0 Then
        Kill DestOpUserFolder & "*.???"
    End If
    
    'Dosyayı şablondan operasyon klasörüne kopyala ve adını değiştir.
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.CopyFile (SourceRapor3_1Normal), DestOpUserFolder & ReNameRapor3_1 & ".docm", True
    
'________________________________________

    'Oluşturulacak dosyayı aç
    On Error Resume Next
    Set objWord = GetObject(, "Word.Application")
    Set objWord = GetObject(, "Word.Application")
    Set objWord = GetObject(, "Word.Application")
    Set objWord = GetObject(, "Word.Application")
    Set objWord = GetObject(, "Word.Application")
    If objWord Is Nothing Then
        'MsgBox "Dosya oluşturmada CreateObject methodu kullanılacak."
        Set objWord = CreateObject("Word.Application")
        objWord.Visible = False
    End If
    objWord.Documents.Open FileName:=DestOpUserFolder & ReNameRapor3_1 & ".docm"
    objWord.Visible = True
    objWord.Activate 'Ekrana getirir.
    'objDoc.Activate 'Ekrana getirmez.
    objWord.Application.WindowState = wdWindowStateMaximize
    Set objDoc = GetObject(DestOpUserFolder & ReNameRapor3_1 & ".docm")
'________________________________________

    Set IlkSiraBul = Range("FG7:FG100000").Find(What:=Cells(ActiveCell.Row, 5).Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    Set SonSiraBul = Range("FH7:FH100000").Find(What:=Cells(ActiveCell.Row, 5).Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    If Not IlkSiraBul Is Nothing Then
        IlkSira = IlkSiraBul.Row
    Else
        MsgBox "The serial number could not be found, so the operation cannot be completed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    If Not SonSiraBul Is Nothing Then
        SonSira = SonSiraBul.Row
    Else
        MsgBox "The serial number could not be found, so the operation cannot be completed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    
    'Birim
    Birimx = UCase(Replace(Replace(Worksheets(2).Cells(6, 99).Value, "i", "I"), "ı", "I")) & " UNIT"
    objDoc.Tables(1).Cell(Row:=1, Column:=1).Range.Text = Birimx
    
    'Dosyada içerikleri değiştir.
    'FinansalBirim adı
    Ek1 = Cells(ActiveCell.Row, 30).Value
    Set MyRange = objDoc.Tables(1).Cell(Row:=4, Column:=1).Range
    With MyRange.Find
        .Text = "<financialUnit>"
        .Replacement.Text = Ek1
        .Execute Replace:=wdReplaceAll
    End With
    MyRange.HighlightColorIndex = wdNoHighlight
    'Teslim birimi
    Ek1 = Cells(ActiveCell.Row, 31).Value
    Set MyRange = objDoc.Tables(1).Cell(Row:=4, Column:=1).Range
    With MyRange.Find
        .Text = "<deliveryUnit>"
        .Replacement.Text = Ek1
        .Execute Replace:=wdReplaceAll
    End With
    MyRange.HighlightColorIndex = wdNoHighlight
    'Teslim tarihi
    Ek1 = Cells(ActiveCell.Row, 35).Value
    Set MyRange = objDoc.Tables(1).Cell(Row:=4, Column:=1).Range
    With MyRange.Find
        .Text = "<deliveryDate>"
        .Replacement.Text = Ek1
        .Execute Replace:=wdReplaceAll
    End With
    MyRange.HighlightColorIndex = wdNoHighlight
    'TipA(un)(ların)
    'Tekil-çoğul tipA
    AdetTopla = Application.Sum(Range(Cells(IlkSira, 58), Cells(SonSira, 58)))
    If AdetTopla = 1 Then
        TekCogulTipA = "Type A item"
        Ek3 = "Type A Item"
        Ek2 = "has"
        Ek4 = "Type A item was"
    ElseIf AdetTopla > 1 Then
        TekCogulTipA = "Type A items"
        Ek3 = "Type A Items"
        Ek2 = "have"
        Ek4 = "Type A items were"
    End If
    Set MyRange = objDoc.Tables(1).Cell(Row:=4, Column:=1).Range
    With MyRange.Find
        .Text = "<typeA>"
        .Replacement.Text = TekCogulTipA
        .Execute Replace:=wdReplaceAll
    End With
    MyRange.HighlightColorIndex = wdNoHighlight
    'Sayım tarihi
    Ek1 = Cells(ActiveCell.Row, 23).Value
    Set MyRange = objDoc.Tables(1).Cell(Row:=4, Column:=1).Range
    With MyRange.Find
        .Text = "<countDate>"
        .Replacement.Text = Ek1
        .Execute Replace:=wdReplaceAll
    End With
    MyRange.HighlightColorIndex = wdNoHighlight
    'Point2/Point3
    If Cells(ActiveCell.Row, 12).Value = "Point2" Then
        Ek1 = "Point 2"
    ElseIf Cells(ActiveCell.Row, 12).Value = "Point3" Then
        Ek1 = "Point 3"
    End If
    Set MyRange = objDoc.Tables(1).Cell(Row:=4, Column:=1).Range
    With MyRange.Find
        .Text = "<location>"
        .Replacement.Text = Ek1
        .Execute Replace:=wdReplaceAll
    End With
    MyRange.HighlightColorIndex = wdNoHighlight
    'invalid birimi
    Ek1 = Cells(ActiveCell.Row, 32).Value
    Set MyRange = objDoc.Tables(1).Cell(Row:=4, Column:=1).Range
    With MyRange.Find
        .Text = "<subjectUnit>"
        .Replacement.Text = Ek1
        .Execute Replace:=wdReplaceAll
    End With
    MyRange.HighlightColorIndex = wdNoHighlight
    'Barkod
    Set MyRange = objDoc.Tables(1).Cell(Row:=4, Column:=1).Range
    With MyRange.Find
        .Text = "<barcode>"
        .Replacement.Text = Cells(IlkSira, 51).Value '& " barkod no.lu torbadaki)"
        .Execute Replace:=wdReplaceAll
    End With
    MyRange.HighlightColorIndex = wdNoHighlight
    'have/has
    Set MyRange = objDoc.Tables(1).Cell(Row:=4, Column:=1).Range
    With MyRange.Find
        .Text = "<have_has>"
        .Replacement.Text = Ek2
        .Execute Replace:=wdReplaceAll
    End With
    MyRange.HighlightColorIndex = wdNoHighlight
    
    'Kolluk
    If Cells(ActiveCell.Row, 47).Value = "Provincial Directorate B" Then
        Kolluk = Cells(ActiveCell.Row, 19).Value & " Provincial Governorship Provincial Directorate B " & Cells(ActiveCell.Row, 48).Value
    ElseIf Cells(ActiveCell.Row, 47).Value = "District Directorate B" Then
        Kolluk = Cells(ActiveCell.Row, 20).Value & " District Governorship District Directorate B " & Cells(ActiveCell.Row, 48).Value
    ElseIf Cells(ActiveCell.Row, 47).Value = "Provincial Directorate C" Then
        Kolluk = Cells(ActiveCell.Row, 19).Value & " Provincial Governorship Provincial Directorate C " & Cells(ActiveCell.Row, 48).Value
    ElseIf Cells(ActiveCell.Row, 47).Value = "District Directorate C" Then
        Kolluk = Cells(ActiveCell.Row, 20).Value & " District Governorship District Directorate C " & Cells(ActiveCell.Row, 48).Value
    ElseIf Cells(ActiveCell.Row, 47).Value = "Provincial Directorate D" Then
        Kolluk = Cells(ActiveCell.Row, 19).Value & " Provincial Governorship Provincial Directorate D " & Cells(ActiveCell.Row, 48).Value
    ElseIf Cells(ActiveCell.Row, 47).Value = "District Directorate D" Then
        Kolluk = Cells(ActiveCell.Row, 20).Value & " District Governorship District Directorate D " & Cells(ActiveCell.Row, 48).Value
    ElseIf Cells(ActiveCell.Row, 47).Value = "Provincial Directorate E" Then
        Kolluk = Cells(ActiveCell.Row, 19).Value & " Provincial Governorship Provincial Directorate E " & Cells(ActiveCell.Row, 48).Value
    ElseIf Cells(ActiveCell.Row, 47).Value = "District Directorate E" Then
        Kolluk = Cells(ActiveCell.Row, 20).Value & " District Governorship District Directorate E " & Cells(ActiveCell.Row, 48).Value
    ElseIf InStr(Cells(ActiveCell.Row, 47).Value, "General Directorate") <> 0 Or InStr(Cells(ActiveCell.Row, 47).Value, "Regional Directorate") <> 0 Then
        If Cells(ActiveCell.Row, 48).Value <> "" Then
            Kolluk = WorksheetFunction.Proper(ThisWorkbook.Worksheets(2).Cells(6, 111).Value) & " " & Cells(ActiveCell.Row, 47).Value & " " & Cells(ActiveCell.Row, 48).Value
        Else
            Kolluk = WorksheetFunction.Proper(ThisWorkbook.Worksheets(2).Cells(6, 111).Value) & " " & Cells(ActiveCell.Row, 47).Value
        End If
    Else
        If Cells(ActiveCell.Row, 48).Value <> "" Then
            If InStr(Cells(ActiveCell.Row, 47).Value, "X.X. ") > 0 Then
                Kolluk = Mid(Cells(ActiveCell.Row, 47).Value, 6, Len(Cells(ActiveCell.Row, 47).Value)) & " " & Cells(ActiveCell.Row, 48).Value
            Else
                Kolluk = Cells(ActiveCell.Row, 47).Value & " " & Cells(ActiveCell.Row, 48).Value
            End If
        Else
            If InStr(Cells(ActiveCell.Row, 47).Value, "X.X. ") > 0 Then
                Kolluk = Mid(Cells(ActiveCell.Row, 47).Value, 6, Len(Cells(ActiveCell.Row, 47).Value))
            Else
                Kolluk = Cells(ActiveCell.Row, 47).Value
            End If
        End If
    End If
    
    
    'Recipient
    Set MyRange = objDoc.Tables(2).Cell(Row:=1, Column:=1).Range
    With MyRange.Find
        .Text = "<recipient>"
        .Replacement.Text = Kolluk
        .Execute Replace:=wdReplaceAll
    End With
    MyRange.HighlightColorIndex = wdNoHighlight
    'Tutanak tarihi
    Ek1 = Cells(ActiveCell.Row, 23).Value
    Set MyRange = objDoc.Tables(2).Cell(Row:=1, Column:=1).Range
    With MyRange.Find
        .Text = "<statementDate>"
        .Replacement.Text = Ek1
        .Execute Replace:=wdReplaceAll
    End With
    MyRange.HighlightColorIndex = wdNoHighlight
    
    'imzalar
    objDoc.Tables(3).Cell(Row:=4, Column:=1).Range.Text = Cells(ActiveCell.Row, 184).Value 'Ad Soyad1
    objDoc.Tables(3).Cell(Row:=5, Column:=1).Range.Text = Cells(ActiveCell.Row, 185).Value 'Unvan1
    objDoc.Tables(3).Cell(Row:=4, Column:=2).Range.Text = Cells(ActiveCell.Row, 187).Value 'Ad Soyad2
    objDoc.Tables(3).Cell(Row:=5, Column:=2).Range.Text = Cells(ActiveCell.Row, 188).Value 'Unvan2
    objDoc.Tables(3).Cell(Row:=4, Column:=3).Range.Text = Cells(ActiveCell.Row, 190).Value 'Ad Soyad3
    objDoc.Tables(3).Cell(Row:=5, Column:=3).Range.Text = Cells(ActiveCell.Row, 191).Value 'Unvan3
        
    'Tablo başlığı
    objDoc.Tables(4).Cell(Row:=1, Column:=1).Range.Text = Ek3 & " Evaluated as Invalid:"
    'Tabloya satır ekle
    x = 0
    If SonSira - IlkSira > 0 Then
        With objDoc.Tables(4)
            For i = IlkSira To SonSira - 1
                .Rows.Add BeforeRow:=objDoc.Tables(4).Rows(3)
                x = x + 1
            Next i
        End With
    End If
    For i = 3 To SonSira - IlkSira + 3
        objDoc.Tables(4).Cell(Row:=i, Column:=1).Range.Text = i - 2 'Tablo sıra no
        objDoc.Tables(4).Cell(Row:=i, Column:=2).Range.Text = Cells(IlkSira + i - 3, 52).Value 'Öğe türü
        objDoc.Tables(4).Cell(Row:=i, Column:=3).Range.Text = Cells(IlkSira + i - 3, 55).Value 'Öğe değeri
        objDoc.Tables(4).Cell(Row:=i, Column:=4).Range.Text = Cells(IlkSira + i - 3, 58).Value 'Adet
        objDoc.Tables(4).Cell(Row:=i, Column:=5).Range.Text = Cells(IlkSira + i - 3, 61).Value 'Öğe ID No
    Next i
    'Tablo başlığı
    objDoc.Tables(5).Cell(Row:=1, Column:=1).Range.Text = "The Financial Unit Officer Who Delivered the " & Ek3 & ":"
    objDoc.Tables(5).Cell(Row:=2, Column:=3).Range.Text = Cells(ActiveCell.Row, 36).Value 'Ad Soyad
    objDoc.Tables(5).Cell(Row:=3, Column:=3).Range.Text = Cells(ActiveCell.Row, 38).Value 'TCK No
    objDoc.Tables(5).Cell(Row:=4, Column:=3).Range.Text = Cells(ActiveCell.Row, 39).Value 'Baba adı
    objDoc.Tables(5).Cell(Row:=5, Column:=3).Range.Text = Cells(ActiveCell.Row, 40).Value 'Doğum yeri
    objDoc.Tables(5).Cell(Row:=6, Column:=3).Range.Text = Cells(ActiveCell.Row, 41).Value 'Doğum tarihi
    objDoc.Tables(5).Cell(Row:=7, Column:=3).Range.Text = Cells(ActiveCell.Row, 44).Value 'Tel no
    
'    'Ek teslimat dekontu fotokopisi
    objDoc.Tables(6).Cell(Row:=1, Column:=1).Range.Text = "Attachment"
    objDoc.Tables(6).Cell(Row:=1, Column:=2).Range.Text = ":"

    x = Application.Sum(Range(Cells(IlkSira, 50), Cells(SonSira, 50))) 'Ek teslimat dekontu fotokopisi
    If x > 1 Then objDoc.Tables(6).Cell(Row:=1, Column:=3).Range.Text = "Delivery Receipt (" & x & " pages)"
    If x < 2 Then objDoc.Tables(6).Cell(Row:=1, Column:=3).Range.Text = "Delivery Receipt (" & x & " page)"
            
    objDoc.CheckBox1.Caption = "The invalid " & Ek4 & " received with the approval of the financial unit officer."
    objDoc.CheckBox1.Value = True
    'Türkçe karakterleri düzelt
    objDoc.CheckBox1.Enabled = False
    objDoc.CheckBox1.Enabled = True

    'Alt bilgi ekle
    objDoc.Sections(1).Footers(wdHeaderFooterPrimary).Range.Tables(1).Cell(Row:=1, Column:=1).Range.Text = ReNameRapor3_1
 
    'Sayfa sayısı kaydet komutuna bağlandı.
    objDoc.Save
    'Sayfayı text dosyasından çek
    TxtFileRapor3_1 = DestOpUserFolder & "Report 3 Page Count.txt"
    #If UseADO Then
        Const adReadLine = -2&
        With CreateObject("ADODB.Stream")
            .Open
            .LoadFromFile TxtFileRapor3_1
            Do Until .EOS
                TotalSayfaRapor3_1 = .ReadText(adReadLine)
            Loop
            .Close
        End With
    #Else 'UseFSO
        Const Rapor3_1TristateTrue = -1&
        With CreateObject("Scripting.FileSystemObject")
            With .OpenTextFile(TxtFileRapor3_1, Format:=Rapor3_1TristateTrue)
                Do Until .AtEndOfStream
                    TotalSayfaRapor3_1 = .ReadLine
                Loop
                .Close
            End With
        End With
    #End If
    'MsgBox "Tutanak1: " & TotalSayfaRapor3_1

    If TumDoc = True Then
        objWord.Activate
'        For i = 1 To SayPrt
'            objDoc.PrintOut
'        Next i
        objDoc.PrintOut Background:=False, Copies:=SayPrt
        'objWord.Documents.Save
        objDoc.Close SaveChanges:=False
        objWord.Visible = False
    End If

    'Tutanak1 ve Döküm sayfa sayısı
    Cells(ActiveCell.Row, 169).Value = TotalSayfaRapor3_1
    
    'Nesneleri temizle
    Set fso = Nothing
    Set objWord = Nothing
    Set objDoc = Nothing
    GoTo Son
    
Hata:
MsgBox "An error occurred while retrieving the number of pages of the statement. Please manually enter the number of statement pages in the attachment section when creating the cover letter.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
'If Err.Number <> 0 Then
'MsgBox "Error # " & Str(Err.Number)
'End If

Son:

'Nesneleri temizle
Set fso = Nothing
Set objWord = Nothing
Set objDoc = Nothing

'Columns("CK:CN").EntireColumn.Hidden = True
'Columns("CE:CF").EntireColumn.Hidden = True

'Worksheets(5).Protect Password:="123"', DrawingObjects:=False

'Worksheets(5).Activate
Application.ScreenUpdating = True
Application.DisplayAlerts = True
'Application.EnableEvents = True

'End If

End Sub

Sub Rapor3_2Rapor()

Dim DestOperasyon As String, SourceTutanak1Normal As String, AutoPath As String, ReNameTutanak1 As String
Dim fso As Object, objWord As Object, objDoc As Object
Dim OpenKontrolName As String, OpenControl As String, DestOpUserFolder As String, DestOpUserFolderName As String
Dim ContSay As Integer, KontrolFile As String, x As Long, y As Long, i As Long, j As Long
Dim IlkSiraBul As Range, SonSiraBul As Range, IlkSira As Long, SonSira As Long
Dim GelenTema As String, SourceDL As String, ReNameDL As String
Dim TxtFileTutanak1 As String, TotalSayfaTutanak1 As String, TxtFileDokum As String, TotalSayfaDokum As String
Dim DokumFileTxt As Object, DokumSayfaGonder As String
Dim SourceRaporNormal As String, SourceTutanak2Normal As String, SourceUstYaziNormal As String
Dim Bolum1 As String, Bolum2 As String, Bolum3 As String, Bolum4 As String, Bolum5 As String, Bolum6 As String
Dim Ek1 As String, Ek2 As String, Ek3 As String, Ek4 As String, Ek5 As String, Ek6 As String, Birlestir As String

Dim ReNameRaporNormal As String, SiraNoIlkSatir As Long, RaporIlgi As String
Dim AltRaporNoIlk As Long, AltRaporNoSon As Long, AdetTopla As Long
Dim TekCogulTipA As String, k As Integer
Dim MyRange As Object, TxtFileRapor As String, TotalSayfaRapor As String
Dim ReNameTutanak2Normal As String, GidenTema As String
Dim TxtFileTutanak2 As String, TotalSayfaTutanak2 As String
Dim ReNameUstYaziNormal As String
Dim TxtFileUstYazi As String, TotalSayfaUstYazi As String
Dim Birimx As String, UserName As String, NotEkle As String
Dim Explorer As Integer, b As Long, RaporTipi As String, DestNotlar As String, TxtFileNot As String
Dim TextLine As String, StrTeknik_ANotu As String

'RAPOR için prosedürü başlat
'If ActiveCell.Column = 7 Then
'    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    'Application.EnableEvents = False

    'Dokümanların oluşturulmasına soldan sağa doğru bir kontrol mekanizması kur.
    j = ActiveCell.Row
    For i = ActiveCell.Row To 7 Step -1
        If Cells(i, 5).Value <> "" Then
            GoTo SiraNoBulundu
        Else
            j = j - 1
        End If
    Next i
SiraNoBulundu:
    If j < 7 Then
        MsgBox "The serial number could not be found, so the operation cannot be completed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    Set IlkSiraBul = Range("FG7:FG100000").Find(What:=Cells(j, 5).Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    Set SonSiraBul = Range("FH7:FH100000").Find(What:=Cells(j, 5).Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    If Not IlkSiraBul Is Nothing Then
        IlkSira = IlkSiraBul.Row
    Else
        MsgBox "The serial number could not be found, so the operation cannot be completed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    If Not SonSiraBul Is Nothing Then
        SonSira = SonSiraBul.Row
    Else
        MsgBox "The serial number could not be found, so the operation cannot be completed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    If Cells(IlkSira, 6).Value = "x" Then
        MsgBox "The statement data contains missing or incorrect information, so the operation cannot be started.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    
    'Pathfinder...
    AutoPath = ThisWorkbook.Path
    DestOperasyon = AutoPath & "\System Files\Operation\"
    DestNotlar = AutoPath & "\System Files\System Templates\Item Notes\"
    
    'Birden fazla kullanıcının operasyon klasöründe çakışmasını önlemek için operasyon klasörünün içinde yeni klasör aç.
    DestOpUserFolderName = "Operation-" & Left(ThisWorkbook.name, InStr(ThisWorkbook.name, ".") - 1) '& " " & Format(Now(), "ddmmyyyyhhmmss")
    DestOpUserFolder = DestOperasyon & DestOpUserFolderName & "\"
    
    'Check the "System Files" folder name.
    If Not Dir(AutoPath & "\System Files\", vbDirectory) <> vbNullString Then
        MsgBox "Cannot access the directory " & AutoPath & "\System Files\" & ". The folder named 'System Files' may have been renamed or deleted.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    
    'Check the "Operation" folder name.
    If Not Dir(DestOperasyon, vbDirectory) <> vbNullString Then
        MsgBox "Cannot access the directory " & DestOperasyon & ". The folder named 'Operation' may have been renamed or deleted.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If


    'Operation klasörü içinde kullanıcı modülü klasörü yoksa oluştur.
    If Not Dir(DestOpUserFolder, vbDirectory) <> vbNullString Then
        MkDir DestOpUserFolder
    End If

'____________________________________
If TumDoc = True Then
    'Worksheets(4).Activate
    b = ActiveCell.Row

    'İlk ve son sıraları bul (For Explorer)
    'On Error Resume Next
    SiraNoIlkSatir = ActiveCell.Row
    If Cells(ActiveCell.Row, 5).Value = "" Then
        For i = ActiveCell.Row To 7 Step -1
            If Cells(i, 5).Value <> "" Then
                ReNameRaporNormal = Cells(i, 5).Value & "-" & Cells(6, 7).Value & " " & Cells(ActiveCell.Row, 13).Value
                SiraNoIlkSatir = i
                GoTo RaporNoDonguSonExplorer
            End If
        Next i
    Else
        ReNameRaporNormal = Cells(ActiveCell.Row, 5).Value & "-" & Cells(6, 7).Value & " " & Cells(ActiveCell.Row, 13).Value
    End If
RaporNoDonguSonExplorer:
   
    'Rapor/Alt Rapor aralıklarının tespiti.
    Set IlkSiraBul = Range("FG7:FG100000").Find(What:=Cells(SiraNoIlkSatir, 5).Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    Set SonSiraBul = Range("FH7:FH100000").Find(What:=Cells(SiraNoIlkSatir, 5).Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    If Not IlkSiraBul Is Nothing Then
        IlkSira = IlkSiraBul.Row
    Else
        MsgBox "The serial number could not be found, so the operation cannot be completed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    If Not SonSiraBul Is Nothing Then
        SonSira = SonSiraBul.Row
    Else
        MsgBox "The serial number could not be found, so the operation cannot be completed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
End If
'____________________________________

If TumDoc = False Then
    IlkSira = 1
    SonSira = 1
End If

For Explorer = IlkSira To SonSira
    If TumDoc = True Then
        If Cells(Explorer, 13) <> "" Then
           Cells(Explorer, 13).Select
        Else
            GoTo ExplorerBos
        End If
    End If
    'Dosyayı isimlendir
    'On Error Resume Next
    SiraNoIlkSatir = ActiveCell.Row
    If Cells(ActiveCell.Row, 5).Value = "" Then
        For i = ActiveCell.Row To 7 Step -1
            If Cells(i, 5).Value <> "" Then
                ReNameRaporNormal = Cells(i, 5).Value & "-" & Cells(6, 7).Value & " " & Cells(ActiveCell.Row, 13).Value
                SiraNoIlkSatir = i
                GoTo RaporNoDonguSon
            End If
        Next i
    Else
        ReNameRaporNormal = Cells(ActiveCell.Row, 5).Value & "-" & Cells(6, 7).Value & " " & Cells(ActiveCell.Row, 13).Value
    End If
RaporNoDonguSon:
    'Rapor/Alt Rapor aralıklarının tespiti.
    Set IlkSiraBul = Range("FG7:FG100000").Find(What:=Cells(SiraNoIlkSatir, 5).Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    Set SonSiraBul = Range("FH7:FH100000").Find(What:=Cells(SiraNoIlkSatir, 5).Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    If Not IlkSiraBul Is Nothing Then
        IlkSira = IlkSiraBul.Row
    Else
        MsgBox "The serial number could not be found, so the operation cannot be completed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    If Not SonSiraBul Is Nothing Then
        SonSira = SonSiraBul.Row
    Else
        MsgBox "The serial number could not be found, so the operation cannot be completed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    'MsgBox IlkSira & "-" & SonSira
    For i = ActiveCell.Row To SonSira
        If Cells(i + 1, 7).Value <> "" And i < SonSira Then
            AltRaporNoIlk = ActiveCell.Row
            AltRaporNoSon = i
            'MsgBox AltRaporNoIlk & "-" & AltRaporNoSon
            GoTo RaporNoDonguSon1
        ElseIf i = SonSira Then
            AltRaporNoIlk = ActiveCell.Row
            AltRaporNoSon = i
            'MsgBox AltRaporNoIlk & "-" & AltRaporNoSon
            GoTo RaporNoDonguSon1
        End If
    Next i
RaporNoDonguSon1:

'________________________________________

    If TumDoc = False Then
        'Close the all Word application
        Call ModuleReport2.OpenWordControl
    
        'Operation klasöründeki docm uzantılı word dosyalarından açık olanları kapat ve temizle.
        OpenKontrolName = Dir(DestOpUserFolder & "*.docm")
        Do While OpenKontrolName <> ""
            OpenControl = IsFileOpen(DestOpUserFolder & OpenKontrolName)
            If OpenControl = True Then 'Açıksa
                On Error Resume Next
                Set objWord = GetObject(, "Word.Application")
                Set objWord = GetObject(, "Word.Application")
                Set objWord = GetObject(, "Word.Application")
                Set objWord = GetObject(, "Word.Application")
                Set objWord = GetObject(, "Word.Application")
                objWord.Quit SaveChanges:=True
                'MsgBox "Dosya OpenKontrol methodu ile kapatıldı."
    
            End If
            OpenKontrolName = Dir()
        Loop
    
    '    Call ModuleReport2.OpenWordControl
    
        Set objWord = Nothing
        Set objDoc = Nothing
    
    '________________________________________
    
        On Error Resume Next
    '    Klasörün içindeki tüm dosyaları sil (txt, docm vb.)
        ContSay = 0
        KontrolFile = Dir(DestOpUserFolder & "*.???")
        Do While KontrolFile <> ""
            ContSay = ContSay + 1
            KontrolFile = Dir()
        Loop
        If ContSay > 0 Then
            Kill DestOpUserFolder & "*.???"
        End If
    End If

'________________________________________Dinamik dosya isimleri

    ThisWorkbook.Activate
    RaporTipi = Cells(ActiveCell.Row, 214).Value
    SourceRaporNormal = AutoPath & "\System Files\System Templates\Report 2 Templates\" & RaporTipi & ".docm"
    'Klasör isimlerini kontrol et.
    If Not Dir(SourceRaporNormal, vbDirectory) <> vbNullString Then
        MsgBox "Cannot access the directory " & SourceRaporNormal & ". The names of folders and/or files in this directory may have been changed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    
    'Dosyayı şablondan operasyon klasörüne kopyala ve adını değiştir.
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.CopyFile (SourceRaporNormal), DestOpUserFolder & ReNameRaporNormal & ".docm", True
'________________________________________

    'Oluşturulacak dosyayı aç
    On Error Resume Next
    Set objWord = GetObject(, "Word.Application")
    Set objWord = GetObject(, "Word.Application")
    Set objWord = GetObject(, "Word.Application")
    Set objWord = GetObject(, "Word.Application")
    Set objWord = GetObject(, "Word.Application")
    If objWord Is Nothing Then
        'MsgBox "Dosya oluşturmada CreateObject methodu kullanılacak."
        Set objWord = CreateObject("Word.Application")
        objWord.Visible = False
    End If

    objWord.Documents.Open FileName:=DestOpUserFolder & ReNameRaporNormal & ".docm"
    objWord.Visible = True
    objWord.Activate 'Ekrana getirir.
    'objDoc.Activate 'Ekrana getirmez.
    objWord.Application.WindowState = wdWindowStateMaximize
    Set objDoc = GetObject(DestOpUserFolder & ReNameRaporNormal & ".docm")
'________________________________________

    'Birim
    Birimx = UCase(Replace(Replace(Worksheets(2).Cells(6, 99).Value, "i", "I"), "ı", "I")) & " UNIT"
    objDoc.Tables(1).Cell(Row:=1, Column:=1).Range.Text = Birimx
    'Rapor tarihi
    objDoc.Tables(1).Cell(Row:=4, Column:=3).Range.Text = Cells(IlkSira, 218).Value
    'Rapor No
    objDoc.Tables(1).Cell(Row:=5, Column:=3).Range.Text = Cells(AltRaporNoIlk, 217).Value

    objDoc.CheckBox1.Value = False 'Rapor Talebi
    objDoc.CheckBox2.Value = True 'Rapor3
    
'    'Rapora esas yazı
'    If Cells(ActiveCell.Row, 31).Value <> "" Then
'        GelenTema = Cells(IlkSira, 30).Value & " " & Cells(IlkSira, 31).Value & " için düzenlenen " & Cells(IlkSira, 23).Value & " tarihli tutanak."
'    Else
'        GelenTema = Cells(IlkSira, 30).Value & " için düzenlenen " & Cells(IlkSira, 23).Value & " tarihli tutanak."
'    End If

    'Rapora esas yazı (English version)
    If Cells(ActiveCell.Row, 31).Value <> "" Then
        GelenTema = "The statement dated " & Cells(IlkSira, 23).Value & " issued for " & Cells(IlkSira, 30).Value & " " & Cells(IlkSira, 31).Value & "."
    Else
        GelenTema = "The statement dated " & Cells(IlkSira, 23).Value & " issued for " & Cells(IlkSira, 30).Value & "."
    End If


    'İlgi
    RaporIlgi = GelenTema
    If Cells(ActiveCell.Row, 212).Value <> "valid" Then
        objDoc.Tables(1).Cell(Row:=8, Column:=3).Range.Text = RaporIlgi
        'İlgili birim olmayacak. Çünkü rapor doğrudan KURUM_A tarafından düzenleniyor.
        objDoc.Tables(1).Rows(9).Delete
    Else
        objDoc.Tables(1).Cell(Row:=6, Column:=3).Range.Text = RaporIlgi
        'İlgili birim olmayacak. Çünkü rapor doğrudan KURUM_A tarafından düzenleniyor.
        objDoc.Tables(1).Rows(7).Delete
    End If
    
    'Tekil-çoğul tipA
    AdetTopla = Application.Sum(Range(Cells(AltRaporNoIlk, 58), Cells(AltRaporNoSon, 58)))
    If AdetTopla = 1 Then
        TekCogulTipA = "Type A item"
        Ek3 = "Type A Item"
    ElseIf AdetTopla > 1 Then
        TekCogulTipA = "Type A items"
        Ek3 = "Type A Items"
    End If
    objDoc.Tables(2).Cell(Row:=1, Column:=1).Range.Text = "The Examined " & Ek3

    'Tabloyu doldur
    j = 2
    For i = AltRaporNoIlk To AltRaporNoSon
        x = i - AltRaporNoIlk + 1
        If x Mod 2 <> 0 Then
            'Başlıklar
            objDoc.Tables(2).Cell(Row:=j + 0, Column:=1).Range.Text = "Item Type"
            objDoc.Tables(2).Cell(Row:=j + 1, Column:=1).Range.Text = "Item Value"
            objDoc.Tables(2).Cell(Row:=j + 2, Column:=1).Range.Text = "Qty"
            objDoc.Tables(2).Cell(Row:=j + 3, Column:=1).Range.Text = "Item ID No."
            objDoc.Tables(2).Cell(Row:=j + 0, Column:=2).Range.Text = ":"
            objDoc.Tables(2).Cell(Row:=j + 1, Column:=2).Range.Text = ":"
            objDoc.Tables(2).Cell(Row:=j + 2, Column:=2).Range.Text = ":"
            objDoc.Tables(2).Cell(Row:=j + 3, Column:=2).Range.Text = ":"
            'Verileri aktar
            objDoc.Tables(2).Cell(Row:=j + 0, Column:=3).Range.Text = Cells(i, 52).Value 'Öğe türü
            objDoc.Tables(2).Cell(Row:=j + 1, Column:=3).Range.Text = Cells(i, 55).Value 'Öğe değeri
            objDoc.Tables(2).Cell(Row:=j + 2, Column:=3).Range.Text = Cells(i, 58).Value 'Adet
            objDoc.Tables(2).Cell(Row:=j + 3, Column:=3).Range.Text = Cells(i, 61).Value 'Öğe ID
            j = j + 0
        ElseIf x Mod 2 = 0 Then
            'Verileri aktar
            objDoc.Tables(2).Cell(Row:=j + 0, Column:=4).Range.Text = Cells(i, 52).Value 'Öğe türü
            objDoc.Tables(2).Cell(Row:=j + 1, Column:=4).Range.Text = Cells(i, 55).Value 'Öğe değeri
            objDoc.Tables(2).Cell(Row:=j + 2, Column:=4).Range.Text = Cells(i, 58).Value 'Adet
            objDoc.Tables(2).Cell(Row:=j + 3, Column:=4).Range.Text = Cells(i, 61).Value 'Öğe ID
            'Satır ekle veya ekleme
            If i <> AltRaporNoSon Then
                For k = 1 To 5
                    objDoc.Tables(2).Rows.Add BeforeRow:=objDoc.Tables(2).Rows(j + 4)
                Next k
                j = j + 5
            End If
        End If
    Next i
    'Tema kodu
    objDoc.Tables(3).Cell(Row:=1, Column:=3).Range.Text = Cells(IlkSira, 26).Value  'Tema no

    'Rapor metin kısmı
    'Tanımlamalar
    AdetTopla = Application.Sum(Range(Cells(AltRaporNoIlk, 58), Cells(AltRaporNoSon, 58)))
    If AdetTopla = 1 Then
        Ek1 = "item"
    ElseIf AdetTopla > 1 Then
        Ek1 = "items"
    End If
    'Dosyada öğe/öğeler düzenlemesini gerçekleştir.
    Set MyRange = objDoc.Tables(3).Cell(Row:=2, Column:=1).Range
    With MyRange.Find
        .Text = "<item>"
        .Replacement.Text = Ek1
        .Execute Replace:=wdReplaceAll
    End With
    MyRange.HighlightColorIndex = wdNoHighlight
    
    'Artık satırı sil
    objDoc.Tables(2).Rows(j + 4).Delete
    
    'imzalar
    objDoc.Tables(6).Cell(Row:=3, Column:=1).Range.Text = Cells(ActiveCell.Row, 226).Value 'Ad Soyad Directorate
    objDoc.Tables(6).Cell(Row:=4, Column:=1).Range.Text = Cells(ActiveCell.Row, 227).Value 'Unvan
    objDoc.Tables(6).Cell(Row:=3, Column:=2).Range.Text = Cells(ActiveCell.Row, 220).Value 'Ad Soyad1
    objDoc.Tables(6).Cell(Row:=4, Column:=2).Range.Text = Cells(ActiveCell.Row, 221).Value 'Unvan1
    objDoc.Tables(6).Cell(Row:=3, Column:=3).Range.Text = Cells(ActiveCell.Row, 223).Value 'Ad Soyad2
    objDoc.Tables(6).Cell(Row:=4, Column:=3).Range.Text = Cells(ActiveCell.Row, 224).Value 'Unvan2
    
    'Alt bilgi ekle
    objDoc.Sections(1).Footers(wdHeaderFooterPrimary).Range.Tables(1).Cell(Row:=1, Column:=1).Range.Text = ReNameRaporNormal
    
    'Not ekle
    If Cells(ActiveCell.Row, 216).Value = "Yes" Then
        TxtFileNot = DestNotlar & Cells(ActiveCell.Row, 52).Value & ".txt"
        If Not Dir(TxtFileNot, vbDirectory) <> vbNullString Then
            MsgBox "Cannot access the directory " & TxtFileNot & ". The names of folders and/or files in this directory may have been changed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
            GoTo NotEkleAtla
        End If
        Open TxtFileNot For Input As #1
            Do Until EOF(1)
                Line Input #1, TextLine
                NotEkle = NotEkle & TextLine
            Loop
        Close #1
        objDoc.Tables(7).Cell(Row:=1, Column:=1).Range.Text = NotEkle
    End If
NotEkleAtla:

    'Technique A var/yok (SADECE RAPOR3 işleminde olan kısım)
    StrTeknik_ANotu = ""
    For i = AltRaporNoIlk To AltRaporNoSon
        If Cells(i, 212).Value = "invalid" And Left(Cells(i, 213).Value, 11) = "Technique A" Then
            'Notu ekle
            StrTeknik_ANotu = "*For invalid Type A items produced using Technical A, the Report 2.2 code can be determined based on the report prepared after the items are submitted to the XXX Directorate."
            '6. maddeyi ekle
            objDoc.Tables(4).Rows.Add
            objDoc.Tables(4).Cell(Row:=objDoc.Tables(4).Rows.Count, Column:=2).Range.Text = "Production technique appears to involve Technique A.*"
            'Rapor Talebi ve Rapor3 satırlarını kaldır
            objDoc.Tables(1).Rows(6).Delete
            objDoc.Tables(1).Rows(6).Delete
            GoTo Teknik_AOk
        End If
    Next i
Teknik_AOk:
    objDoc.Tables(8).Cell(Row:=1, Column:=1).Range.Text = StrTeknik_ANotu

    'Rapor sayfa sayısını oluşturmadan önce eskisini sil
    For i = IlkSira To SonSira
        If Cells(i, 13).Value = "" Then
            Cells(i, 174).Value = ""
        End If
    Next i
        
    'Sayfa sayısı kaydet komutuna bağlandı.
    objDoc.Save
    'Sayfayı text dosyasından çek
    TxtFileRapor = DestOpUserFolder & "Report 2 Page Count.txt"
    #If UseADO Then
        Const adReadLine = -2&
        With CreateObject("ADODB.Stream")
            .Open
            .LoadFromFile TxtFileRapor
            Do Until .EOS
                TotalSayfaRapor = .ReadText(adReadLine)
            Loop
            .Close
        End With
    #Else 'UseFSO
        Const RaporTristateTrue = -1&
        With CreateObject("Scripting.FileSystemObject")
            With .OpenTextFile(TxtFileRapor, Format:=RaporTristateTrue)
                Do Until .AtEndOfStream
                    TotalSayfaRapor = .ReadLine
                Loop
                .Close
            End With
        End With
    #End If
    'MsgBox "Rapor: " & TotalSayfaRapor
    Cells(ActiveCell.Row, 174).Value = TotalSayfaRapor

    objWord.Visible = True
    objWord.Activate

    If TumDoc = True Then
        objWord.Activate
'        For i = 1 To SayPrt
'            objDoc.PrintOut
'        Next i
        objDoc.PrintOut Background:=False, Copies:=SayPrt
        'objWord.Documents.Save
        objDoc.Close SaveChanges:=False
        objWord.Visible = False
    End If
    
    'Nesneleri temizle
    Set fso = Nothing
    Set objWord = Nothing
    Set objDoc = Nothing

ExplorerBos:

Next Explorer

If TumDoc = True Then
    'Tümünün bulunduğu butonu tekrar seç.
    Cells(b, 11).Select
End If

Son:

'Worksheets(4).Activate
Application.ScreenUpdating = True
Application.DisplayAlerts = True
'Application.EnableEvents = True

'End If

End Sub


Sub Rapor3_2Tutanak2()

Dim DestOperasyon As String, SourceRapor3_1Normal As String, AutoPath As String, ReNameRapor3_1 As String
Dim fso As Object, objWord As Object, objDoc As Object
Dim OpenKontrolName As String, OpenControl As String, DestOpUserFolder As String, DestOpUserFolderName As String
Dim ContSay As Integer, KontrolFile As String, x As Long, y As Long, i As Long, j As Long
Dim IlkSiraBul As Range, SonSiraBul As Range, IlkSira As Long, SonSira As Long
Dim GelenTema As String, SourceDL As String, ReNameDL As String
Dim TxtFileRapor3_1 As String, TotalSayfaRapor3_1 As String, TxtFileDokum As String, TotalSayfaDokum As String
Dim DokumFileTxt As Object, DokumSayfaGonder As String
Dim SourceRaporNormal As String, SourceTutanak2Normal As String, SourceUstYaziNormal As String
Dim Bolum1 As String, Bolum2 As String, Bolum3 As String, Bolum4 As String, Bolum5 As String, Bolum6 As String
Dim Ek1 As String, Ek2 As String, Ek3 As String, Ek4 As String, Ek5 As String, Ek6 As String, Birlestir As String

Dim ReNameRaporNormal As String, SiraNoIlkSatir As Long, RaporIlgi As String
Dim AltRaporNoIlk As Long, AltRaporNoSon As Long, AdetTopla As Long
Dim TekCogulTipA As String, k As Integer
Dim MyRange As Object, TxtFileRapor As String, TotalSayfaRapor As String
Dim ReNameTutanak2Normal As String, GidenTema As String
Dim TxtFileTutanak2 As String, TotalSayfaTutanak2 As String
Dim ReNameUstYaziNormal As String
Dim TxtFileUstYazi As String, TotalSayfaUstYazi As String
Dim Birimx As String, UserName As String, Kolluk As String

'TUTANAK2 için prosedürü başlat
'    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    'Application.EnableEvents = False
    'Worksheets(5).Unprotect Password:="123"


    'Dokümanların oluşturulmasına soldan sağa doğru bir kontrol mekanizması kur.
    j = ActiveCell.Row
    For i = ActiveCell.Row To 7 Step -1
        If Cells(i, 5).Value <> "" Then
            GoTo SiraNoBulundu
        Else
            j = j - 1
        End If
    Next i
SiraNoBulundu:
    If j < 7 Then
        MsgBox "The serial number could not be found, so the operation cannot be completed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    Set IlkSiraBul = Range("FG7:FG100000").Find(What:=Cells(j, 5).Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    Set SonSiraBul = Range("FH7:FH100000").Find(What:=Cells(j, 5).Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    If Not IlkSiraBul Is Nothing Then
        IlkSira = IlkSiraBul.Row
    Else
        MsgBox "The serial number could not be found, so the operation cannot be completed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    If Not SonSiraBul Is Nothing Then
        SonSira = SonSiraBul.Row
    Else
        MsgBox "The serial number could not be found, so the operation cannot be completed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    If Cells(IlkSira, 6).Value = "x" Then
        MsgBox "The statement data contains missing or incorrect information, so the operation cannot be started.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    For i = IlkSira To SonSira
        If Cells(i, 7).Value = "x" Then
            MsgBox "The report data contains missing or incorrect information, so the operation cannot be started.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
            GoTo Son
        End If
    Next i
    

    'Pathfinder...
    AutoPath = ThisWorkbook.Path
    DestOperasyon = AutoPath & "\System Files\Operation\"
    'TUTANAK2 TANIMLARI
    SourceTutanak2Normal = AutoPath & "\System Files\System Templates\Statement 2 Template\Statement 2.docm"
    'Birden fazla kullanıcının operasyon klasöründe çakışmasını önlemek için operasyon klasörünün içinde yeni klasör aç.
    DestOpUserFolderName = "Operation-" & Left(ThisWorkbook.name, InStr(ThisWorkbook.name, ".") - 1) '& " " & Format(Now(), "ddmmyyyyhhmmss")
    DestOpUserFolder = DestOperasyon & DestOpUserFolderName & "\"

    'Check the "System Files" folder name.
    If Not Dir(AutoPath & "\System Files\", vbDirectory) <> vbNullString Then
        MsgBox "Cannot access the directory " & AutoPath & "\System Files\" & ". The folder named 'System Files' may have been renamed or deleted.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    
    'Check the "Operation" folder name.
    If Not Dir(DestOperasyon, vbDirectory) <> vbNullString Then
        MsgBox "Cannot access the directory " & DestOperasyon & ". The folder named 'Operation' may have been renamed or deleted.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    
    'Check folder names.
    If Not Dir(SourceTutanak2Normal, vbDirectory) <> vbNullString Then
        MsgBox "Cannot access the directory " & SourceTutanak2Normal & ". The names of folders and/or files in this directory may have been changed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    
    'Operation klasörü içinde kullanıcı modülü klasörü yoksa oluştur.
    If Not Dir(DestOpUserFolder, vbDirectory) <> vbNullString Then
        MkDir DestOpUserFolder
    End If

    'On Error Resume Next
    ReNameTutanak2Normal = Cells(ActiveCell.Row, 5).Value & "-" & Cells(6, 8).Value
'________________________________________

    If TumDoc = False Then
        'Close the all Word application
        Call ModuleReport3.OpenWordControl

        'Operation klasöründeki docm uzantılı word dosyalarından açık olanları kapat ve temizle.
        OpenKontrolName = Dir(DestOpUserFolder & "*.docm")
        Do While OpenKontrolName <> ""
            OpenControl = IsFileOpen(DestOpUserFolder & OpenKontrolName)
            If OpenControl = True Then 'Açıksa
                On Error Resume Next
                Set objWord = GetObject(, "Word.Application")
                Set objWord = GetObject(, "Word.Application")
                Set objWord = GetObject(, "Word.Application")
                Set objWord = GetObject(, "Word.Application")
                Set objWord = GetObject(, "Word.Application")
                objWord.Quit SaveChanges:=True
                'MsgBox "Dosya OpenKontrol methodu ile kapatıldı."

            End If
            OpenKontrolName = Dir()
        Loop

    '    Call ModuleReport3.OpenWordControl

        Set objWord = Nothing
        Set objDoc = Nothing
    '________________________________________

        On Error Resume Next
    '    Klasörün içindeki tüm dosyaları sil (txt, docm vb.)
        ContSay = 0
        KontrolFile = Dir(DestOpUserFolder & "*.???")
        Do While KontrolFile <> ""
            ContSay = ContSay + 1
            KontrolFile = Dir()
        Loop
        If ContSay > 0 Then
            Kill DestOpUserFolder & "*.???"
        End If
    End If

    'Dosyayı şablondan operasyon klasörüne kopyala ve adını değiştir.
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.CopyFile (SourceTutanak2Normal), DestOpUserFolder & ReNameTutanak2Normal & ".docm", True
'________________________________________

    'Oluşturulacak dosyayı aç
    On Error Resume Next
    Set objWord = GetObject(, "Word.Application")
    Set objWord = GetObject(, "Word.Application")
    Set objWord = GetObject(, "Word.Application")
    Set objWord = GetObject(, "Word.Application")
    Set objWord = GetObject(, "Word.Application")
    If objWord Is Nothing Then
        'MsgBox "Dosya oluşturmada CreateObject methodu kullanılacak."
        Set objWord = CreateObject("Word.Application")
        objWord.Visible = False
    End If
    objWord.Documents.Open FileName:=DestOpUserFolder & ReNameTutanak2Normal & ".docm"
    objWord.Visible = True
    objWord.Activate 'Ekrana getirir.
    'objDoc.Activate 'Ekrana getirmez.
    objWord.Application.WindowState = wdWindowStateMaximize
    Set objDoc = GetObject(DestOpUserFolder & ReNameTutanak2Normal & ".docm")
'________________________________________

    'Birim
    Birimx = UCase(Replace(Replace(Worksheets(2).Cells(6, 99).Value, "i", "I"), "ı", "I")) & " UNIT"
    objDoc.Tables(1).Cell(Row:=1, Column:=1).Range.Text = Birimx
    'Tutanak2 tarihi
    objDoc.Tables(1).Cell(Row:=4, Column:=3).Range.Text = Cells(ActiveCell.Row, 69).Value
    'Belge tarihi ve numarası
'    objDoc.Tables(1).Cell(Row:=6, Column:=3).Range.Text = "" 'Cells(ActiveCell.Row, 20).Value
'    objDoc.Tables(1).Cell(Row:=6, Column:=5).Range.Text = "" 'Cells(ActiveCell.Row, 21).Value
    objDoc.Tables(1).Rows(6).Delete

    'Kolluk
    If Cells(ActiveCell.Row, 47).Value = "Provincial Directorate B" Then
        Kolluk = Cells(ActiveCell.Row, 19).Value & " Provincial Governorship Provincial Directorate B " & Cells(ActiveCell.Row, 48).Value
    ElseIf Cells(ActiveCell.Row, 47).Value = "District Directorate B" Then
        Kolluk = Cells(ActiveCell.Row, 20).Value & " District Governorship District Directorate B " & Cells(ActiveCell.Row, 48).Value
    ElseIf Cells(ActiveCell.Row, 47).Value = "Provincial Directorate C" Then
        Kolluk = Cells(ActiveCell.Row, 19).Value & " Provincial Governorship Provincial Directorate C " & Cells(ActiveCell.Row, 48).Value
    ElseIf Cells(ActiveCell.Row, 47).Value = "District Directorate C" Then
        Kolluk = Cells(ActiveCell.Row, 20).Value & " District Governorship District Directorate C " & Cells(ActiveCell.Row, 48).Value
    ElseIf Cells(ActiveCell.Row, 47).Value = "Provincial Directorate D" Then
        Kolluk = Cells(ActiveCell.Row, 19).Value & " Provincial Governorship Provincial Directorate D " & Cells(ActiveCell.Row, 48).Value
    ElseIf Cells(ActiveCell.Row, 47).Value = "District Directorate D" Then
        Kolluk = Cells(ActiveCell.Row, 20).Value & " District Governorship District Directorate D " & Cells(ActiveCell.Row, 48).Value
    ElseIf Cells(ActiveCell.Row, 47).Value = "Provincial Directorate E" Then
        Kolluk = Cells(ActiveCell.Row, 19).Value & " Provincial Governorship Provincial Directorate E " & Cells(ActiveCell.Row, 48).Value
    ElseIf Cells(ActiveCell.Row, 47).Value = "District Directorate E" Then
        Kolluk = Cells(ActiveCell.Row, 20).Value & " District Governorship District Directorate E " & Cells(ActiveCell.Row, 48).Value
    ElseIf InStr(Cells(ActiveCell.Row, 47).Value, "General Directorate") <> 0 Or InStr(Cells(ActiveCell.Row, 47).Value, "Regional Directorate") <> 0 Then
        If Cells(ActiveCell.Row, 48).Value <> "" Then
            Kolluk = WorksheetFunction.Proper(ThisWorkbook.Worksheets(2).Cells(6, 111).Value) & " " & Cells(ActiveCell.Row, 47).Value & " " & Cells(ActiveCell.Row, 48).Value
        Else
            Kolluk = WorksheetFunction.Proper(ThisWorkbook.Worksheets(2).Cells(6, 111).Value) & " " & Cells(ActiveCell.Row, 47).Value
        End If
    Else
        If Cells(ActiveCell.Row, 48).Value <> "" Then
            If InStr(Cells(ActiveCell.Row, 47).Value, "X.X. ") > 0 Then
                Kolluk = Mid(Cells(ActiveCell.Row, 47).Value, 6, Len(Cells(ActiveCell.Row, 47).Value)) & " " & Cells(ActiveCell.Row, 48).Value
            Else
                Kolluk = Cells(ActiveCell.Row, 47).Value & " " & Cells(ActiveCell.Row, 48).Value
            End If
        Else
            If InStr(Cells(ActiveCell.Row, 47).Value, "X.X. ") > 0 Then
                Kolluk = Mid(Cells(ActiveCell.Row, 47).Value, 6, Len(Cells(ActiveCell.Row, 47).Value))
            Else
                Kolluk = Cells(ActiveCell.Row, 47).Value
            End If
        End If
    End If
    
    objDoc.Tables(1).Cell(Row:=5, Column:=3).Range.Text = Kolluk

    'Tabloyu doldur
    Set IlkSiraBul = Range("FG7:FG100000").Find(What:=Cells(ActiveCell.Row, 5).Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    Set SonSiraBul = Range("FH7:FH100000").Find(What:=Cells(ActiveCell.Row, 5).Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    If Not IlkSiraBul Is Nothing Then
        IlkSira = IlkSiraBul.Row
    Else
        MsgBox "The serial number could not be found, so the operation cannot be completed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    If Not SonSiraBul Is Nothing Then
        SonSira = SonSiraBul.Row
    Else
        MsgBox "The serial number could not be found, so the operation cannot be completed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    'Tabloya satır ekle
    x = 0
    If SonSira - IlkSira > 0 Then
        With objDoc.Tables(2)
            For i = IlkSira To SonSira - 1
                .Rows.Add BeforeRow:=objDoc.Tables(2).Rows(2)
                x = x + 1
            Next i
        End With
    End If

    For i = 2 To SonSira - IlkSira + 2
        objDoc.Tables(2).Cell(Row:=i, Column:=1).Range.Text = i - 1 'Tablo sıra no
        objDoc.Tables(2).Cell(Row:=i, Column:=2).Range.Text = Cells(IlkSira + i - 2, 52).Value 'Öğe türü
        objDoc.Tables(2).Cell(Row:=i, Column:=3).Range.Text = Cells(IlkSira + i - 2, 55).Value 'Öğe değeri
        objDoc.Tables(2).Cell(Row:=i, Column:=4).Range.Text = Cells(IlkSira + i - 2, 58).Value 'Adet
        objDoc.Tables(2).Cell(Row:=i, Column:=5).Range.Text = Cells(IlkSira + i - 2, 61).Value 'Öğe ID No
        objDoc.Tables(2).Cell(Row:=i, Column:=6).Range.Text = Cells(IlkSira, 26).Value  'Tema No (Temai her satıra yaz.)
        objDoc.Tables(2).Cell(Row:=i, Column:=7).Range.Text = Cells(IlkSira + i - 2, 64).Value 'Açıklama
    Next i

    'Tutanak2 tutanağının metin kısmı
    'Tanımlamalar
    If Application.Sum(Range(Cells(IlkSira, 58), Cells(SonSira, 58))) > 1 Then
        Ek4 = "items"
    Else
        Ek4 = "item"
    End If
    Bolum1 = vbTab & "A total of "
    Bolum2 = " " & Ek4 & " , as described above, have been placed inside "
    Bolum3 = " "
    Bolum4 = " and enclosed to be sent to the designated unit mentioned above."

    Ek1 = Application.Sum(Range(Cells(IlkSira, 58), Cells(SonSira, 58)))
    Ek2 = Cells(ActiveCell.Row, 72).Value
    Ek3 = Application.WorksheetFunction.Proper(Cells(ActiveCell.Row, 71).Value)
    Ek3 = Left(Ek3, InStr(Ek3, "/") - 1)
    
    Birlestir = Bolum1 & Ek1 & Bolum2 & Ek2 & Bolum3 & Ek3 & Bolum4
    objDoc.Tables(3).Cell(Row:=1, Column:=1).Range.Text = Birlestir
    
    
    Set MyRange = objDoc.Tables(3).Cell(Row:=1, Column:=1).Range
    MyRange.Find.Execute FindText:=Ek1
    MyRange.Font.Bold = True

    Set MyRange = objDoc.Tables(3).Cell(Row:=1, Column:=1).Range
    With MyRange.Find
        .Execute FindText:=Ek2
        .Execute Forward:=True
    End With
    MyRange.Font.Bold = True
    'Aralıkta bulunan karakterleri bold yap
'    objDoc.Range(objDoc.Tables(3).Cell(Row:=1, Column:=1).Range.Characters(5).Start, _
    objDoc.Tables(3).Cell(Row:=1, Column:=1).Range.Characters(10).End).Font.Bold = True
    Set MyRange = objDoc.Tables(3).Cell(Row:=1, Column:=1).Range
    MyRange.Find.Execute FindText:=Ek3
    MyRange.Font.Bold = True

    'imzalar
    objDoc.Tables(4).Cell(Row:=6, Column:=2).Range.Text = Cells(ActiveCell.Row, 193).Value 'Ad Soyad1
    objDoc.Tables(4).Cell(Row:=7, Column:=2).Range.Text = Cells(ActiveCell.Row, 194).Value 'Unvan1
    objDoc.Tables(4).Cell(Row:=6, Column:=3).Range.Text = Cells(ActiveCell.Row, 196).Value 'Ad Soyad2
    objDoc.Tables(4).Cell(Row:=7, Column:=3).Range.Text = Cells(ActiveCell.Row, 197).Value 'Unvan2
    
    'Alt bilgi ekle
    objDoc.Sections(1).Footers(wdHeaderFooterPrimary).Range.Tables(1).Cell(Row:=1, Column:=1).Range.Text = ReNameTutanak2Normal
    
'    'İmza boşluğunu sayfaya sığdırmak için düzenle
    If SonSira - IlkSira + 1 = 17 Then
        For i = 1 To 2
            objDoc.Tables(4).Rows(1).Delete
        Next i
    ElseIf SonSira - IlkSira + 1 = 18 Then
        For i = 1 To 4
            objDoc.Tables(4).Rows(1).Delete
        Next i
    ElseIf SonSira - IlkSira + 1 = 19 Then
        For i = 1 To 6
            objDoc.Tables(4).Rows(1).Delete
        Next i
    ElseIf SonSira - IlkSira + 1 = 20 Then
        For i = 1 To 7
            objDoc.Tables(4).Rows(1).Delete
        Next i
        objDoc.Tables(4).Rows(1).Height = 5
    End If

    'Sayfa sayısı kaydet komutuna bağlandı.
'    objDoc.Close SaveChanges:=True
'    objWord.Quit
    objDoc.Save
    'Sayfayı text dosyasından çek
    TxtFileTutanak2 = DestOpUserFolder & "Statement 2 Page Count.txt"
    #If UseADO Then
        Const adReadLine = -2&
        With CreateObject("ADODB.Stream")
            .Open
            .LoadFromFile TxtFileTutanak2
            Do Until .EOS
                TotalSayfaTutanak2 = .ReadText(adReadLine)
            Loop
            .Close
        End With
    #Else 'UseFSO
        Const Tutanak2TristateTrue = -1&
        With CreateObject("Scripting.FileSystemObject")
            With .OpenTextFile(TxtFileTutanak2, Format:=Tutanak2TristateTrue)
                Do Until .AtEndOfStream
                    TotalSayfaTutanak2 = .ReadLine
                Loop
                .Close
            End With
        End With
    #End If
    'MsgBox "Tutanak2: " & TotalSayfaTutanak2

    'Tutanak2 sayfa sayısı
    Cells(ActiveCell.Row, 171).Value = TotalSayfaTutanak2

    objWord.Visible = True
    objWord.Activate

    If TumDoc = True Then
        objWord.Activate
'        For i = 1 To SayPrt
'            objDoc.PrintOut
'        Next i
        objDoc.PrintOut Background:=False, Copies:=SayPrt
        'objWord.Documents.Save
        objDoc.Close SaveChanges:=False
        objWord.Visible = False
    End If

    'Nesneleri temizle
    Set fso = Nothing
    Set objWord = Nothing
    Set objDoc = Nothing

Son:

'Worksheets(5).Protect Password:="123"', DrawingObjects:=False

'Worksheets(5).Activate
Application.ScreenUpdating = True
Application.DisplayAlerts = True
'Application.EnableEvents = True

'End If

End Sub

Sub Rapor3_2FinansalBirimUstYazi()

Dim DestOperasyon As String, SourceRapor3_1Normal As String, AutoPath As String, ReNameRapor3_1 As String
Dim fso As Object, objWord As Object, objDoc As Object
Dim OpenKontrolName As String, OpenControl As String, DestOpUserFolder As String, DestOpUserFolderName As String
Dim ContSay As Integer, KontrolFile As String, x As Long, y As Long, i As Long, j As Long
Dim IlkSiraBul As Range, SonSiraBul As Range, IlkSira As Long, SonSira As Long
Dim GelenTema As String, SourceDL As String, ReNameDL As String
Dim TxtFileRapor3_1 As String, TotalSayfaRapor3_1 As String, TxtFileDokum As String, TotalSayfaDokum As String
Dim DokumFileTxt As Object, DokumSayfaGonder As String
Dim SourceRaporNormal As String, SourceTutanak2Normal As String, SourceFinansalBirimUstYaziNormal As String
Dim Bolum1 As String, Bolum2 As String, Bolum3 As String, Bolum4 As String, Bolum5 As String, Bolum6 As String
Dim Ek1 As String, Ek2 As String, Ek3 As String, Ek4 As String, Ek5 As String, Ek6 As String, Birlestir As String

Dim ReNameRaporNormal As String, SiraNoIlkSatir As Long, RaporIlgi As String
Dim AltRaporNoIlk As Long, AltRaporNoSon As Long, AdetTopla As Long
Dim TekCogulTipA As String, k As Integer
Dim MyRange As Object, TxtFileRapor As String, TotalSayfaRapor As String
Dim ReNameTutanak2Normal As String, GidenTema As String
Dim TxtFileTutanak2 As String, TotalSayfaTutanak2 As String
Dim ReNameFinansalBirimUstYaziNormal As String
Dim TxtFileFinansalBirimUstYazi As String, TotalSayfaFinansalBirimUstYazi As String
Dim Birimx As String, UserName As String, Kolluk As String

Dim BStr As Integer
Dim M2 As Boolean, M3 As Boolean, M4 As Boolean
Dim Ifv As Boolean, Ify As Boolean
Dim XXXMudNotu As Boolean
Dim IlgiStrSay As Integer, Govde1StrSay As Integer, IlgiRng As Object, Govde1Rng As Object, Govde2Rng As Object
Dim IlgiFarkSay As Integer, Govde1FarkSay As Integer
Dim ustbilgimuhatap As String, ustbilgitarih As String, ustbilgisayi As String
Dim CokluSayfa As Integer
Dim ItemBul As Range, AdSoyadParaf As String, UnvanParaf As String, TelParaf As String


IlceSakla = ""
If InStr(Cells(ActiveCell.Row, 78).Value, " Organization A") <> 0 Then
    IlceSakla = Cells(ActiveCell.Row, 78).Value
    Cells(ActiveCell.Row, 78).Value = ""
End If

'ÜST YAZI için prosedürü başlat
'    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    'Application.EnableEvents = False
    'Worksheets(5).Unprotect Password:="123"


    'Dokümanların oluşturulmasına soldan sağa doğru bir kontrol mekanizması kur.
    j = ActiveCell.Row
    For i = ActiveCell.Row To 7 Step -1
        If Cells(i, 5).Value <> "" Then
            GoTo SiraNoBulundu
        Else
            j = j - 1
        End If
    Next i
SiraNoBulundu:
    If j < 7 Then
        MsgBox "The serial number could not be found, so the operation cannot be completed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    Set IlkSiraBul = Range("FG7:FG100000").Find(What:=Cells(j, 5).Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    Set SonSiraBul = Range("FH7:FH100000").Find(What:=Cells(j, 5).Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    If Not IlkSiraBul Is Nothing Then
        IlkSira = IlkSiraBul.Row
    Else
        MsgBox "The serial number could not be found, so the operation cannot be completed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    If Not SonSiraBul Is Nothing Then
        SonSira = SonSiraBul.Row
    Else
        MsgBox "The serial number could not be found, so the operation cannot be completed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    
    If Cells(IlkSira, 6).Value = "x" Then
        MsgBox "The statement data contains missing or incorrect information, so the operation cannot be started.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    For i = IlkSira To SonSira
        If Cells(i, 7).Value = "x" Then
            MsgBox "The report data contains missing or incorrect information, so the operation cannot be started.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
            GoTo Son
        End If
    Next i
    If Cells(IlkSira, 8).Value = "x" Then
        MsgBox "The Statement 2 data contains missing or incorrect information, so the operation cannot be started.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If

    'Pathfinder...
    AutoPath = ThisWorkbook.Path
    DestOperasyon = AutoPath & "\System Files\Operation\"
    'ÜST YAZI TANIMLARI
    SourceFinansalBirimUstYaziNormal = AutoPath & "\System Files\System Templates\Report 3 Cover Letter Templates\Report 3.2 – Type A Cover Letter – Financial Unit.docm"
    'Birden fazla kullanıcının operasyon klasöründe çakışmasını önlemek için operasyon klasörünün içinde yeni klasör aç.
    DestOpUserFolderName = "Operation-" & Left(ThisWorkbook.name, InStr(ThisWorkbook.name, ".") - 1) '& " " & Format(Now(), "ddmmyyyyhhmmss")
    DestOpUserFolder = DestOperasyon & DestOpUserFolderName & "\"

    'Check "System Files" folder.
    If Not Dir(AutoPath & "\System Files\", vbDirectory) <> vbNullString Then
        MsgBox "Cannot access the directory " & AutoPath & "\System Files\" & ". The folder named 'System Files' may have been renamed or deleted.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    
    'Check "Operation" folder.
    If Not Dir(DestOperasyon, vbDirectory) <> vbNullString Then
        MsgBox "Cannot access the directory " & DestOperasyon & ". The folder named 'Operation' may have been renamed or deleted.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    
    'Check folder names.
    If Not Dir(SourceFinansalBirimUstYaziNormal, vbDirectory) <> vbNullString Then
        MsgBox "Cannot access the directory " & SourceFinansalBirimUstYaziNormal & ". The names of folders and/or files in this directory may have been changed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If

    
    'Operation klasörü içinde kullanıcı modülü klasörü yoksa oluştur.
    If Not Dir(DestOpUserFolder, vbDirectory) <> vbNullString Then
        MkDir DestOpUserFolder
    End If

    'On Error Resume Next
    Set IlkSiraBul = Range("FG7:FG100000").Find(What:=Cells(ActiveCell.Row, 5).Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    Set SonSiraBul = Range("FH7:FH100000").Find(What:=Cells(ActiveCell.Row, 5).Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    If Not IlkSiraBul Is Nothing Then
        IlkSira = IlkSiraBul.Row
    Else
        MsgBox "The serial number could not be found, so the operation cannot be completed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    If Not SonSiraBul Is Nothing Then
        SonSira = SonSiraBul.Row
    Else
        MsgBox "The serial number could not be found, so the operation cannot be completed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If

    'Eklerde belirtilecek sayfalar için ön kontroller
    'Bu kısım çıkarıldı. Çünkü bu dokümanın ekinde önce aşamada oluştutulan raporlar yer almıyor.

    'PARAF HAZIRLIK İŞLEMİ
    UserName = Environ("UserProfile")
    UserName = UCase(Right(UserName, 7))
    Set ItemBul = Worksheets(2).Range("DR6:DR1000").Find(What:=UserName, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    If Not ItemBul Is Nothing Then
        AdSoyadParaf = "For information only: " & Worksheets(2).Cells(ItemBul.Row, 123).Value
        UnvanParaf = Worksheets(2).Cells(ItemBul.Row, 124).Value
        TelParaf = "Phone / Extension No: " & Worksheets(2).Cells(ItemBul.Row, 126).Value
    Else
        MsgBox UserName & " session is not registered in the system, so the cover letter cannot be created. Please register your session using the Initials Interface located in the Settings Group of the Enterprise Document Automation System menu and try the operation again.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    
    'On Error Resume Next
    ReNameFinansalBirimUstYaziNormal = Cells(ActiveCell.Row, 5).Value & "-" & Cells(6, 9).Value & " Üst Yazı"
'________________________________________

    If TumDoc = False Then
        'Close the all Word application
        Call ModuleReport3.OpenWordControl

        'Operation klasöründeki docm uzantılı word dosyalarından açık olanları kapat ve temizle.
        OpenKontrolName = Dir(DestOpUserFolder & "*.docm")
        Do While OpenKontrolName <> ""
            OpenControl = IsFileOpen(DestOpUserFolder & OpenKontrolName)
            If OpenControl = True Then 'Açıksa
    '            On Error Resume Next
    '            Set objWord = GetObject(, "Word.Application")
    '            If objWord Is Nothing Then
    '                Set objWord = CreateObject("Word.Application")
    '                objWord.Visible = True
    '            End If
    '            objWord.Quit SaveChanges:=True

                On Error Resume Next
                Set objWord = GetObject(, "Word.Application")
                Set objWord = GetObject(, "Word.Application")
                Set objWord = GetObject(, "Word.Application")
                Set objWord = GetObject(, "Word.Application")
                Set objWord = GetObject(, "Word.Application")
                objWord.Quit SaveChanges:=True
                'MsgBox "Dosya OpenKontrol methodu ile kapatıldı."

            End If
            OpenKontrolName = Dir()
        Loop

    '    Call ModuleReport3.OpenWordControl

        Set objWord = Nothing
        Set objDoc = Nothing

    '________________________________________

        On Error Resume Next
    '    Klasörün içindeki tüm dosyaları sil (txt, docm vb.)
        ContSay = 0
        KontrolFile = Dir(DestOpUserFolder & "*.???")
        Do While KontrolFile <> ""
            ContSay = ContSay + 1
            KontrolFile = Dir()
        Loop
        If ContSay > 0 Then
            Kill DestOpUserFolder & "*.???"
        End If
    End If

    'Dosyayı şablondan operasyon klasörüne kopyala ve adını değiştir.
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.CopyFile (SourceFinansalBirimUstYaziNormal), DestOpUserFolder & ReNameFinansalBirimUstYaziNormal & ".docm", True
'________________________________________

    'Oluşturulacak dosyayı aç
    On Error Resume Next
    Set objWord = GetObject(, "Word.Application")
    Set objWord = GetObject(, "Word.Application")
    Set objWord = GetObject(, "Word.Application")
    Set objWord = GetObject(, "Word.Application")
    Set objWord = GetObject(, "Word.Application")
    If objWord Is Nothing Then
        'MsgBox "Dosya oluşturmada CreateObject methodu kullanılacak."
        Set objWord = CreateObject("Word.Application")
        objWord.Visible = False
    End If
    objWord.Documents.Open FileName:=DestOpUserFolder & ReNameFinansalBirimUstYaziNormal & ".docm"
    objWord.Visible = True
    objWord.Activate 'Ekrana getirir.
    'objDoc.Activate 'Ekrana getirmez.
    objWord.Application.WindowState = wdWindowStateMaximize
    Set objDoc = GetObject(DestOpUserFolder & ReNameFinansalBirimUstYaziNormal & ".docm")
'________________________________________

    'Birim
    Birimx = UCase(Replace(Replace(Worksheets(2).Cells(6, 99).Value, "i", "I"), "ı", "I")) & " UNIT"
    objDoc.Tables(1).Cell(Row:=1, Column:=1).Range.Text = Birimx

    Birimx = UCase(Replace(Replace(Worksheets(2).Cells(6, 99).Value, "i", "I"), "ı", "I"))
    'Üst yazı tarihi
    objDoc.Tables(1).Cell(Row:=1, Column:=2).Range.Text = Birimx & ", " & FormatEnglishDate(Cells(ActiveCell.Row, 75).Value) '(Format(Cells(ActiveCell.Row, 75).Value, "d mmmm yyyy"))
    'Yazı no
    objDoc.Tables(1).Cell(Row:=4, Column:=2).Range.Text = Cells(ActiveCell.Row, 76).Value

    'Gönderi tipi
    Ek2 = Cells(ActiveCell.Row, 85).Value
    Ek2 = UCase(Replace(Replace(Ek2, "i", "I"), "ı", "I"))
    objDoc.Tables(1).Cell(Row:=7, Column:=2).Range.Text = Ek2

    'Kolluk
    If Cells(ActiveCell.Row, 47).Value = "Provincial Directorate B" Then
        Kolluk = Cells(ActiveCell.Row, 19).Value & " Provincial Governorship Provincial Directorate B " & Cells(ActiveCell.Row, 48).Value
    ElseIf Cells(ActiveCell.Row, 47).Value = "District Directorate B" Then
        Kolluk = Cells(ActiveCell.Row, 20).Value & " District Governorship District Directorate B " & Cells(ActiveCell.Row, 48).Value
    ElseIf Cells(ActiveCell.Row, 47).Value = "Provincial Directorate C" Then
        Kolluk = Cells(ActiveCell.Row, 19).Value & " Provincial Governorship Provincial Directorate C " & Cells(ActiveCell.Row, 48).Value
    ElseIf Cells(ActiveCell.Row, 47).Value = "District Directorate C" Then
        Kolluk = Cells(ActiveCell.Row, 20).Value & " District Governorship District Directorate C " & Cells(ActiveCell.Row, 48).Value
    ElseIf Cells(ActiveCell.Row, 47).Value = "Provincial Directorate D" Then
        Kolluk = Cells(ActiveCell.Row, 19).Value & " Provincial Governorship Provincial Directorate D " & Cells(ActiveCell.Row, 48).Value
    ElseIf Cells(ActiveCell.Row, 47).Value = "District Directorate D" Then
        Kolluk = Cells(ActiveCell.Row, 20).Value & " District Governorship District Directorate D " & Cells(ActiveCell.Row, 48).Value
    ElseIf Cells(ActiveCell.Row, 47).Value = "Provincial Directorate E" Then
        Kolluk = Cells(ActiveCell.Row, 19).Value & " Provincial Governorship Provincial Directorate E " & Cells(ActiveCell.Row, 48).Value
    ElseIf Cells(ActiveCell.Row, 47).Value = "District Directorate E" Then
        Kolluk = Cells(ActiveCell.Row, 20).Value & " District Governorship District Directorate E " & Cells(ActiveCell.Row, 48).Value
    ElseIf InStr(Cells(ActiveCell.Row, 47).Value, "General Directorate") <> 0 Or InStr(Cells(ActiveCell.Row, 47).Value, "Regional Directorate") <> 0 Then
        If Cells(ActiveCell.Row, 48).Value <> "" Then
            Kolluk = WorksheetFunction.Proper(ThisWorkbook.Worksheets(2).Cells(6, 111).Value) & " " & Cells(ActiveCell.Row, 47).Value & " " & Cells(ActiveCell.Row, 48).Value
        Else
            Kolluk = WorksheetFunction.Proper(ThisWorkbook.Worksheets(2).Cells(6, 111).Value) & " " & Cells(ActiveCell.Row, 47).Value
        End If
    Else
        If InStr(Cells(ActiveCell.Row, 47).Value, "X.X. ") > 0 Then  'Başında X.X. ifadesi varsa
            If Cells(ActiveCell.Row, 48).Value <> "" Then
                Kolluk = Mid(Cells(ActiveCell.Row, 47).Value, 6, Len(Cells(ActiveCell.Row, 47).Value)) & " " & Cells(ActiveCell.Row, 48).Value
                
            Else
                Kolluk = Mid(Cells(ActiveCell.Row, 47).Value, 6, Len(Cells(ActiveCell.Row, 47).Value))
            End If
        Else
            If Cells(ActiveCell.Row, 48).Value <> "" Then
                Kolluk = Cells(ActiveCell.Row, 47).Value & " " & Cells(ActiveCell.Row, 48).Value
            Else
                Kolluk = Cells(ActiveCell.Row, 47).Value
            End If
        End If
    End If

    'YENİ MUHATAP TEMASI
    M2 = False
    M3 = False
    M4 = False
    BStr = 9
    '1. sayfadan sonraki üst bilgi tanımları
    ustbilgitarih = (Format(Cells(ActiveCell.Row, 75).Value, "dd.mm.yyyy"))
    ustbilgisayi = Cells(ActiveCell.Row, 76).Value
    
    'Muhatap
    If Cells(ActiveCell.Row, 82).Value <> "" Then
        objDoc.Tables(1).Cell(Row:=BStr, Column:=1).Range.Text = UCase(Replace(Replace(Cells(ActiveCell.Row, 30).Value, "i", "I"), "ı", "I")) 'FinansalBirim
        objDoc.Tables(1).Cell(Row:=BStr + 1, Column:=1).Range.Text = "(" & Cells(ActiveCell.Row, 82).Value & ")" 'Birim
        objDoc.Tables(1).Cell(Row:=BStr + 2, Column:=1).Range.Text = Cells(ActiveCell.Row, 79).Value 'Adres
        If Cells(ActiveCell.Row, 78).Value <> "" Then
            objDoc.Tables(1).Cell(Row:=BStr + 3, Column:=1).Range.Text = Cells(ActiveCell.Row, 78).Value & "/" & UCase(Replace(Replace(Cells(ActiveCell.Row, 77).Value, "i", "I"), "ı", "I"))
        Else
            objDoc.Tables(1).Cell(Row:=BStr + 3, Column:=1).Range.Text = UCase(Replace(Replace(Cells(ActiveCell.Row, 77).Value, "i", "I"), "ı", "I"))
        End If
        objDoc.Tables(1).Cell(Row:=BStr + 3, Column:=1).Range.Font.Underline = wdUnderlineSingle
        M4 = True
        
        ustbilgimuhatap = Cells(ActiveCell.Row, 30).Value
    Else
        objDoc.Tables(1).Cell(Row:=BStr, Column:=1).Range.Text = UCase(Replace(Replace(Cells(ActiveCell.Row, 30).Value, "i", "I"), "ı", "I")) 'FinansalBirim
        objDoc.Tables(1).Cell(Row:=BStr + 1, Column:=1).Range.Text = Cells(ActiveCell.Row, 79).Value 'Adres
        If Cells(ActiveCell.Row, 78).Value <> "" Then
            objDoc.Tables(1).Cell(Row:=BStr + 2, Column:=1).Range.Text = Cells(ActiveCell.Row, 78).Value & "/" & UCase(Replace(Replace(Cells(ActiveCell.Row, 77).Value, "i", "I"), "ı", "I"))
        Else
            objDoc.Tables(1).Cell(Row:=BStr + 2, Column:=1).Range.Text = UCase(Replace(Replace(Cells(ActiveCell.Row, 77).Value, "i", "I"), "ı", "I"))
        End If
        objDoc.Tables(1).Cell(Row:=BStr + 2, Column:=1).Range.Font.Underline = wdUnderlineSingle
        M3 = True
        
        ustbilgimuhatap = Cells(ActiveCell.Row, 30).Value
    End If
    
    'Üst yazı gövde metni ve ek senaryoları
    'tipAnın/tipAların (TipA adedi)
    AdetTopla = Application.Sum(Range(Cells(IlkSira, 58), Cells(SonSira, 58)))
    If AdetTopla = 1 Then
        Ek1 = "Type A item"
        Ek4 = "item"
        'Ek5 = "is"
        Ek5 = "is" 'always singular
        Ek6 = "has"
    ElseIf AdetTopla > 1 Then
        Ek1 = "Type A items"
        Ek4 = "item"
        'Ek5 = "are"
        Ek5 = "is" 'always singular
        Ek6 = "have"
    End If
    Ek2 = Cells(ActiveCell.Row, 35).Value 'Teslimat tarihi
    If Cells(ActiveCell.Row, 12).Value = "Point2" Then
        Ek3 = "Point 2"
    ElseIf Cells(ActiveCell.Row, 12).Value = "Point3" Then
        Ek3 = "Point 3"
    End If
    
    Bolum1 = "The labels related to the invalid " & Ek1 & ", received by our unit on " & Ek2 & " via xxxxx xxxxx xxxxx for the purpose of xxxxx xxxxx xxxxx xxxxx " & _
        "and processed at " & Ek3 & ", " & Ek5 & " hereby enclosed and delivered to your office. The details of the " & Ek4 & " are documented in the attached statement."
    Bolum2 = "The " & Ek1 & " " & Ek6 & " been retained by our unit for forwarding to the X1 Process Monitoring Directorate."
    Bolum3 = Chr(9) & "Submitted for your information."
    Birlestir = Chr(9) & Bolum1 & vbNewLine & Chr(9) & Bolum2
    objDoc.Tables(2).Cell(Row:=1, Column:=1).Range.Text = Birlestir
    objDoc.Tables(2).Cell(Row:=3, Column:=1).Range.Text = Bolum3

    'imzalar
    objDoc.Tables(3).Cell(Row:=4, Column:=2).Range.Text = Cells(ActiveCell.Row, 199).Value 'Ad Soyad1
    objDoc.Tables(3).Cell(Row:=5, Column:=2).Range.Text = Cells(ActiveCell.Row, 200).Value 'Unvan1
    objDoc.Tables(3).Cell(Row:=4, Column:=3).Range.Text = Cells(ActiveCell.Row, 202).Value 'Ad Soyad2
    objDoc.Tables(3).Cell(Row:=5, Column:=3).Range.Text = Cells(ActiveCell.Row, 203).Value 'Unvan2
    
    'Ekler
    objDoc.Tables(3).Cell(Row:=8, Column:=1).Range.Text = "Attachment:"
    If Cells(ActiveCell.Row, 80).Value <> "" Then
        'Etiket adedi
        If Cells(ActiveCell.Row, 81).Value > 1 Then objDoc.Tables(3).Cell(Row:=8, Column:=2).Range.Text = " 1) Label (" & Cells(ActiveCell.Row, 81).Value & " pieces)"
        If Cells(ActiveCell.Row, 81).Value < 2 Then objDoc.Tables(3).Cell(Row:=8, Column:=2).Range.Text = " 1) Label (" & Cells(ActiveCell.Row, 81).Value & " piece)"
        'Dekont
        If Cells(ActiveCell.Row, 80).Value > 1 Then objDoc.Tables(3).Cell(Row:=9, Column:=2).Range.Text = " 2) Xxxxx Receipt (" & Cells(ActiveCell.Row, 80).Value & " pages)"
        If Cells(ActiveCell.Row, 80).Value < 2 Then objDoc.Tables(3).Cell(Row:=9, Column:=2).Range.Text = " 2) Xxxxx Receipt (" & Cells(ActiveCell.Row, 80).Value & " page)"
        'Statement
        x = Application.Sum(Range(Cells(IlkSira, 169), Cells(SonSira, 169))) + Application.Sum(Range(Cells(IlkSira, 50), Cells(SonSira, 50)))
        objDoc.Tables(3).Cell(Row:=10, Column:=2).Range.Text = " 3) Attached Statement (" & "total of " & x & " pages)"
    Else
        'Etiket adedi
        If Cells(ActiveCell.Row, 81).Value > 1 Then objDoc.Tables(3).Cell(Row:=8, Column:=2).Range.Text = " 1) Label (" & Cells(ActiveCell.Row, 81).Value & " pieces)"
        If Cells(ActiveCell.Row, 81).Value < 2 Then objDoc.Tables(3).Cell(Row:=8, Column:=2).Range.Text = " 1) Label (" & Cells(ActiveCell.Row, 81).Value & " piece)"
        'Statement
        x = Application.Sum(Range(Cells(IlkSira, 169), Cells(SonSira, 169))) + Application.Sum(Range(Cells(IlkSira, 50), Cells(SonSira, 50)))
        objDoc.Tables(3).Cell(Row:=9, Column:=2).Range.Text = " 2) Attached Statement (" & "total of " & x & " pages)"
    End If



    '__________________________DİNAMİK SAYFA YAPISI
    
    'İlgi ve gövde metninin kaç satırdan oluştuğu bilgisinin elde edilmesinde kullanılıyor.
    'Set IlgiRng = objDoc.Tables(1).Cell(Row:=15, Column:=3).Range
    Set Govde1Rng = objDoc.Tables(2).Cell(Row:=1, Column:=1).Range
    Set Govde2Rng = objDoc.Tables(2).Cell(Row:=3, Column:=1).Range
    
    'IlgiStrSay = Govde1Rng.Information(wdFirstCharacterLineNumber) - IlgiRng.Information(wdFirstCharacterLineNumber) - 1
    Govde1StrSay = Govde2Rng.Information(wdFirstCharacterLineNumber) - Govde1Rng.Information(wdFirstCharacterLineNumber)
    'MsgBox "İlgi Satır Sayısı: " & IlgiStrSay & vbNewLine & vbNewLine & "Gövde 1 Satır Sayısı: " & Govde1StrSay
    'IlgiFarkSay = IlgiStrSay - 1
    IlgiFarkSay = 0
    Govde1FarkSay = Govde1StrSay - 6 'Varsayılan 2 paragrafta (1 rowda) toplam 5 satıra göre sıfırlandı.
    CokluSayfa = 0
    'MsgBox Govde1FarkSay
    
    '__________________________________'Tes.Düzensiz.Dek. var.
   
    'Dinamik sayfa düzeni
    
    objDoc.Tables(2).Rows(2).Delete 'Boş satırı sil
    For i = 1 To 3
        objDoc.Tables(1).Rows(13).Delete 'Muhatap sonrası
    Next i
    
    If Cells(ActiveCell.Row, 80).Value <> "" Then 'Tes.Düzensiz.Dek. var. '3 ek var
        For i = 1 To 2 'Ek sonrası satırları sil
            objDoc.Tables(3).Rows(11).Delete
        Next i
        Govde1FarkSay = Govde1FarkSay + 1
    Else 'Tes.Düzensiz.Dek. yok. '2 ek var
        For i = 1 To 3 'Ek sonrası satırları sil
            objDoc.Tables(3).Rows(10).Delete
        Next i
        Govde1FarkSay = Govde1FarkSay + 0
    End If

    If M4 = True Then
        If IlgiFarkSay + Govde1FarkSay < 17 Then
            Govde1FarkSay = Govde1FarkSay + 2
            'MsgBox IlgiFarkSay + Govde1FarkSay
            If IlgiFarkSay + Govde1FarkSay > 8 Then
                For i = 9 To IlgiFarkSay + Govde1FarkSay
                    If i = 9 Then
                        objDoc.Tables(3).Rows(6).Delete 'Ek öncesi
                    ElseIf i = 10 Then
                        objDoc.Tables(3).Rows(6).Delete 'Ek öncesi
                    ElseIf i = 11 Then
                        objDoc.Tables(3).Rows(1).Delete 'İmza öncesi
                    ElseIf i = 12 Then
                        objDoc.Tables(3).Rows(1).Delete 'İmza öncesi
                    ElseIf i = 13 Then
                        objDoc.Tables(1).Rows(8).Delete 'Muhatap öncesi
                    ElseIf i = 14 Then
                        objDoc.Tables(1).Rows(8).Delete 'Muhatap öncesi
                    End If
                Next i
            Else
                If (IlgiFarkSay + Govde1FarkSay) < 8 Then
                    For i = 1 To 8 - (IlgiFarkSay + Govde1FarkSay)
                        'Boş satır ekle
                        With objDoc.Tables(3)
                            .Rows.Add BeforeRow:=objDoc.Tables(3).Rows(8)
                        End With
                    Next i
                End If
            End If
        Else
            CokluSayfa = 1
        End If
    ElseIf M3 = True Then
        For i = 1 To 1
            objDoc.Tables(1).Rows(13).Delete 'ilgi öncesi
        Next i

        If IlgiFarkSay + Govde1FarkSay < 16 Then
            Govde1FarkSay = Govde1FarkSay + 1
            'MsgBox IlgiFarkSay + Govde1FarkSay
            If IlgiFarkSay + Govde1FarkSay > 8 Then
                For i = 9 To IlgiFarkSay + Govde1FarkSay
                    If i = 9 Then
                        objDoc.Tables(3).Rows(6).Delete 'Ek öncesi
                    ElseIf i = 10 Then
                        objDoc.Tables(3).Rows(6).Delete 'Ek öncesi
                    ElseIf i = 11 Then
                        objDoc.Tables(3).Rows(1).Delete 'İmza öncesi
                    ElseIf i = 12 Then
                        objDoc.Tables(3).Rows(1).Delete 'İmza öncesi
                    ElseIf i = 13 Then
                        objDoc.Tables(1).Rows(8).Delete 'Muhatap öncesi
                    ElseIf i = 14 Then
                        objDoc.Tables(1).Rows(8).Delete 'Muhatap öncesi
                    End If
                Next i
            Else
                If (IlgiFarkSay + Govde1FarkSay) < 8 Then
                    For i = 1 To 8 - (IlgiFarkSay + Govde1FarkSay)
                        'Boş satır ekle
                        With objDoc.Tables(3)
                            .Rows.Add BeforeRow:=objDoc.Tables(3).Rows(8)
                        End With
                    Next i
                End If
            End If
        Else
            CokluSayfa = 1
        End If
    ElseIf M2 = True Then
        For i = 1 To 2
            objDoc.Tables(1).Rows(13).Delete 'ilgi öncesi
        Next i
        
        'Govde1FarkSay = Govde1FarkSay + 0
        'MsgBox IlgiFarkSay + Govde1FarkSay
        If IlgiFarkSay + Govde1FarkSay < 15 Then
            If IlgiFarkSay + Govde1FarkSay > 8 Then
                For i = 9 To IlgiFarkSay + Govde1FarkSay
                    If i = 9 Then
                        objDoc.Tables(3).Rows(6).Delete 'Ek öncesi
                    ElseIf i = 10 Then
                        objDoc.Tables(3).Rows(6).Delete 'Ek öncesi
                    ElseIf i = 11 Then
                        objDoc.Tables(3).Rows(1).Delete 'İmza öncesi
                    ElseIf i = 12 Then
                        objDoc.Tables(3).Rows(1).Delete 'İmza öncesi
                    ElseIf i = 13 Then
                        objDoc.Tables(1).Rows(8).Delete 'Muhatap öncesi
                    ElseIf i = 14 Then
                        objDoc.Tables(1).Rows(8).Delete 'Muhatap öncesi
                    End If
                Next i
            Else
                If (IlgiFarkSay + Govde1FarkSay) < 8 Then
                    For i = 1 To 8 - (IlgiFarkSay + Govde1FarkSay)
                        'Boş satır ekle
                        With objDoc.Tables(3)
                            .Rows.Add BeforeRow:=objDoc.Tables(3).Rows(8)
                        End With
                    Next i
                End If
            End If
        Else
            CokluSayfa = 1
        End If
    End If
    
    objDoc.Paragraphs(objDoc.Paragraphs.Count).Range.Font.Size = 1


    'Footer
    If objDoc.ActiveWindow.View.SplitSpecial <> wdPaneNone Then
        objDoc.ActiveWindow.Panes(2).Close
    End If
    If objDoc.ActiveWindow.ActivePane.View.Type = wdNormalView Or objDoc.ActiveWindow.ActivePane.View.Type = wdOutlineView Then
        objDoc.ActiveWindow.ActivePane.View.Type = wdPrintView
    End If

    With objDoc
        With .Sections(1).Headers(wdHeaderFooterPrimary).Range
            .InsertAfter Text:="Page "
            .Fields.Add Range:=.Characters.Last, Type:=wdFieldEmpty, Text:="PAGE", PreserveFormatting:=False
            .InsertAfter Text:=" of the letter dated " & ustbilgitarih & _
                             ", reference number " & ustbilgisayi & _
                             ", addressed to the " & ustbilgimuhatap
        End With
    End With
    
    objDoc.ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument

    
    'PARAF EKLEME İŞLEMİ
    objDoc.Sections(1).Footers(wdHeaderFooterFirstPage).Range.Tables(1).Cell(Row:=1, Column:=2).Range.Text = AdSoyadParaf & vbNewLine & _
                                                                                       UnvanParaf & vbNewLine & _
                                                                                       TelParaf
    objDoc.Sections(1).Footers(wdHeaderFooterPrimary).Range.Tables(1).Cell(Row:=1, Column:=2).Range.Text = AdSoyadParaf & vbNewLine & _
                                                                                       UnvanParaf & vbNewLine & _
                                                                                       TelParaf


Tekrarla:
    'Sayfa sayısı kaydet komutuna bağlandı.
    objDoc.Save
    'Sayfayı text dosyasından çek
    TxtFileFinansalBirimUstYazi = DestOpUserFolder & "Cover Letter Page Count.txt"
    #If UseADO Then
        Const adReadLine = -2&
        With CreateObject("ADODB.Stream")
            .Open
            .LoadFromFile TxtFileFinansalBirimUstYazi
            Do Until .EOS
                TotalSayfaFinansalBirimUstYazi = .ReadText(adReadLine)
            Loop
            .Close
        End With
    #Else 'UseFSO
        Const FinansalBirimUstYaziTristateTrue = -1&
        With CreateObject("Scripting.FileSystemObject")
            With .OpenTextFile(TxtFileFinansalBirimUstYazi, Format:=FinansalBirimUstYaziTristateTrue)
                Do Until .AtEndOfStream
                    TotalSayfaFinansalBirimUstYazi = .ReadLine
                Loop
                .Close
            End With
        End With
    #End If
    'MsgBox "Üst Yazı: " & TotalSayfaFinansalBirimUstYazi

    'Üst Yazı sayfa sayısı
    Cells(ActiveCell.Row, 172).Value = TotalSayfaFinansalBirimUstYazi

    objWord.Visible = True
    objWord.Activate

    If TumDoc = True Then
        objWord.Activate
'        For i = 1 To SayPrt
'            objDoc.PrintOut
'        Next i
        objDoc.PrintOut Background:=False, Copies:=SayPrt
        'objWord.Documents.Save
        objDoc.Close SaveChanges:=False
        objWord.Visible = False
    End If

    'Nesneleri temizle
    Set fso = Nothing
    Set objWord = Nothing
    Set objDoc = Nothing

Son:

If IlceSakla <> "" Then
    Cells(ActiveCell.Row, 78).Value = IlceSakla
End If

'Worksheets(5).Protect Password:="123"', DrawingObjects:=False
'Worksheets(5).Activate
Application.ScreenUpdating = True
Application.DisplayAlerts = True
'Application.EnableEvents = True


End Sub

Sub Rapor3_2UstYazi()

Dim DestOperasyon As String, SourceRapor3_1Normal As String, AutoPath As String, ReNameRapor3_1 As String
Dim fso As Object, objWord As Object, objDoc As Object
Dim OpenKontrolName As String, OpenControl As String, DestOpUserFolder As String, DestOpUserFolderName As String
Dim ContSay As Integer, KontrolFile As String, x As Long, y As Long, i As Long, j As Long
Dim IlkSiraBul As Range, SonSiraBul As Range, IlkSira As Long, SonSira As Long
Dim GelenTema As String, SourceDL As String, ReNameDL As String
Dim TxtFileRapor3_1 As String, TotalSayfaRapor3_1 As String, TxtFileDokum As String, TotalSayfaDokum As String
Dim DokumFileTxt As Object, DokumSayfaGonder As String
Dim SourceRaporNormal As String, SourceTutanak2Normal As String, SourceUstYaziNormal As String
Dim Bolum1 As String, Bolum2 As String, Bolum3 As String, Bolum4 As String, Bolum5 As String, Bolum6 As String, Bolum7 As String
Dim Ek1 As String, Ek2 As String, Ek3 As String, Ek4 As String, Ek5 As String, Ek6 As String, Birlestir As String
Dim Ek7 As String, Ek8 As String, Ek9 As String, Ek_Paket As String

Dim ReNameRaporNormal As String, SiraNoIlkSatir As Long, RaporIlgi As String
Dim AltRaporNoIlk As Long, AltRaporNoSon As Long, AdetTopla As Long
Dim TekCogulTipA As String, k As Integer
Dim MyRange As Object, TxtFileRapor As String, TotalSayfaRapor As String
Dim ReNameTutanak2Normal As String, GidenTema As String
Dim TxtFileTutanak2 As String, TotalSayfaTutanak2 As String
Dim ReNameUstYaziNormal As String
Dim TxtFileUstYazi As String, TotalSayfaUstYazi As String
Dim Birimx As String, UserName As String
Dim IlBuyukHarf, IlKucukHarf, IlceBuyukHarf, IlceKucukHarf As String
Dim GonderimUsulu As String

Dim BStr As Integer
Dim M2 As Boolean, M3 As Boolean, M4 As Boolean
Dim Ifv As Boolean, Ify As Boolean
Dim XXXMudNotu As Boolean
Dim IlgiStrSay As Integer, Govde1StrSay As Integer, IlgiRng As Object, Govde1Rng As Object, Govde2Rng As Object
Dim IlgiFarkSay As Integer, Govde1FarkSay As Integer
Dim ustbilgimuhatap As String, ustbilgitarih As String, ustbilgisayi As String
Dim CokluSayfa As Integer
Dim ItemBul As Range, AdSoyadParaf As String, UnvanParaf As String, TelParaf As String
Dim StrTeknik_ANotu As String
Dim gecersizSay As Long


IlceSakla = ""
If InStr(Cells(ActiveCell.Row, 20).Value, " Organization A") <> 0 Then
    IlceSakla = Cells(ActiveCell.Row, 20).Value
    Cells(ActiveCell.Row, 20).Value = ""
End If

'ÜST YAZI için prosedürü başlat
'    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    'Application.EnableEvents = False
    'Worksheets(5).Unprotect Password:="123"

    'Dokümanların oluşturulmasına soldan sağa doğru bir kontrol mekanizması kur.
    j = ActiveCell.Row
    For i = ActiveCell.Row To 7 Step -1
        If Cells(i, 5).Value <> "" Then
            GoTo SiraNoBulundu
        Else
            j = j - 1
        End If
    Next i
SiraNoBulundu:
    If j < 7 Then
        MsgBox "The serial number could not be found, so the operation cannot be completed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    Set IlkSiraBul = Range("FG7:FG100000").Find(What:=Cells(j, 5).Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    Set SonSiraBul = Range("FH7:FH100000").Find(What:=Cells(j, 5).Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    If Not IlkSiraBul Is Nothing Then
        IlkSira = IlkSiraBul.Row
    Else
        MsgBox "The serial number could not be found, so the operation cannot be completed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    If Not SonSiraBul Is Nothing Then
        SonSira = SonSiraBul.Row
    Else
        MsgBox "The serial number could not be found, so the operation cannot be completed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    
    If Cells(IlkSira, 6).Value = "x" Then
        MsgBox "Report 3 statement data is incomplete and/or contains errors, so your operation cannot be started.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    For i = IlkSira To SonSira
        If Cells(i, 7).Value = "x" Then
            MsgBox "Report data is incomplete and/or contains errors, so your operation cannot be started.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
            GoTo Son
        End If
    Next i
    If Cells(IlkSira, 8).Value = "x" Then
        MsgBox "Statement 2 data is incomplete and/or contains errors, so your operation cannot be started.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    If Cells(IlkSira, 9).Value = "x" Then
        MsgBox "Financial Unit cover letter data is incomplete and/or contains errors, so your operation cannot be started.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    
    'Pathfinder...
    AutoPath = ThisWorkbook.Path
    DestOperasyon = AutoPath & "\System Files\Operation\"
    'ÜST YAZI TANIMLARI
    SourceUstYaziNormal = AutoPath & "\System Files\System Templates\Report 3 Cover Letter Templates\Report 3.2 – Type A Cover Letter.docm"
    'Birden fazla kullanıcının operasyon klasöründe çakışmasını önlemek için operasyon klasörünün içinde yeni klasör aç.
    DestOpUserFolderName = "Operation-" & Left(ThisWorkbook.name, InStr(ThisWorkbook.name, ".") - 1) '& " " & Format(Now(), "ddmmyyyyhhmmss")
    DestOpUserFolder = DestOperasyon & DestOpUserFolderName & "\"

    'System Files folder name check.
    If Not Dir(AutoPath & "\System Files\", vbDirectory) <> vbNullString Then
        MsgBox AutoPath & "\System Files\" & " directory cannot be accessed. The folder named 'System Files' may have been renamed or deleted.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    
    'Check Operation folder name.
    If Not Dir(DestOperasyon, vbDirectory) <> vbNullString Then
        MsgBox DestOperasyon & " directory cannot be accessed. The folder named 'Operation' may have been renamed or deleted.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    
    'Check folder names.
    If Not Dir(SourceUstYaziNormal, vbDirectory) <> vbNullString Then
        MsgBox SourceUstYaziNormal & " directory cannot be accessed. The names of the folders and/or files in this directory may have been changed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If

    
    'Operation klasörü içinde kullanıcı modülü klasörü yoksa oluştur.
    If Not Dir(DestOpUserFolder, vbDirectory) <> vbNullString Then
        MkDir DestOpUserFolder
    End If

    'On Error Resume Next
    Set IlkSiraBul = Range("FG7:FG100000").Find(What:=Cells(ActiveCell.Row, 5).Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    Set SonSiraBul = Range("FH7:FH100000").Find(What:=Cells(ActiveCell.Row, 5).Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    If Not IlkSiraBul Is Nothing Then
        IlkSira = IlkSiraBul.Row
    Else
        MsgBox "The serial number could not be found, so the operation cannot be completed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    If Not SonSiraBul Is Nothing Then
        SonSira = SonSiraBul.Row
    Else
        MsgBox "The serial number could not be found, so the operation cannot be completed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If

    'Pre-checks for pages to be attached
    'Tutanak1 check
    If Cells(IlkSira, 169).Value = "" Then
        MsgBox Cells(IlkSira, 5).Value & " row number: Report 3.2 statement has not been created, so the summary letter for row " & Cells(IlkSira, 5).Value & " cannot be prepared. Please create Report 3 statement for row " & Cells(IlkSira, 5).Value & " and try again.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    
    'Tutanak2 check
    If Cells(IlkSira, 171).Value = "" Then
        MsgBox Cells(IlkSira, 5).Value & " row number: Statement 2 has not been created, so the summary letter for row " & Cells(IlkSira, 5).Value & " cannot be prepared. Please create Statement 2 for row " & Cells(IlkSira, 5).Value & " and try again.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    
    'Financial Unit Cover Letter check
    If Cells(IlkSira, 172).Value = "" Then
        MsgBox Cells(IlkSira, 5).Value & " row number: Financial Unit cover letter has not been created, so the summary letter for row " & Cells(IlkSira, 5).Value & " cannot be prepared. Please create the Financial Unit cover letter for row " & Cells(IlkSira, 5).Value & " and try again.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If


    'PARAF HAZIRLIK İŞLEMİ
    UserName = Environ("UserProfile")
    UserName = UCase(Right(UserName, 7))
    Set ItemBul = Worksheets(2).Range("DR6:DR1000").Find(What:=UserName, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    If Not ItemBul Is Nothing Then
        AdSoyadParaf = "For information only: " & Worksheets(2).Cells(ItemBul.Row, 123).Value
        UnvanParaf = Worksheets(2).Cells(ItemBul.Row, 124).Value
        TelParaf = "Phone / Extension No: " & Worksheets(2).Cells(ItemBul.Row, 126).Value
    Else
        MsgBox UserName & " session is not registered in the system, so the cover letter cannot be created. Please register your session in the system using the Initials Interface located in the Settings Group of the Enterprise Document Automation System menu, and then try the operation again.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    
    'On Error Resume Next
    ReNameUstYaziNormal = Cells(ActiveCell.Row, 5).Value & "-" & Cells(6, 10).Value
'________________________________________

    If TumDoc = False Then
        'Close the all Word application
        Call ModuleReport3.OpenWordControl

        'Operation klasöründeki docm uzantılı word dosyalarından açık olanları kapat ve temizle.
        OpenKontrolName = Dir(DestOpUserFolder & "*.docm")
        Do While OpenKontrolName <> ""
            OpenControl = IsFileOpen(DestOpUserFolder & OpenKontrolName)
            If OpenControl = True Then 'Açıksa
                On Error Resume Next
                Set objWord = GetObject(, "Word.Application")
                Set objWord = GetObject(, "Word.Application")
                Set objWord = GetObject(, "Word.Application")
                Set objWord = GetObject(, "Word.Application")
                Set objWord = GetObject(, "Word.Application")
                objWord.Quit SaveChanges:=True
                'MsgBox "Dosya OpenKontrol methodu ile kapatıldı."

            End If
            OpenKontrolName = Dir()
        Loop

    '    Call ModuleReport3.OpenWordControl

        Set objWord = Nothing
        Set objDoc = Nothing

    '________________________________________

        On Error Resume Next
    '    Klasörün içindeki tüm dosyaları sil (txt, docm vb.)
        ContSay = 0
        KontrolFile = Dir(DestOpUserFolder & "*.???")
        Do While KontrolFile <> ""
            ContSay = ContSay + 1
            KontrolFile = Dir()
        Loop
        If ContSay > 0 Then
            Kill DestOpUserFolder & "*.???"
        End If
    End If

    'Dosyayı şablondan operasyon klasörüne kopyala ve adını değiştir.
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.CopyFile (SourceUstYaziNormal), DestOpUserFolder & ReNameUstYaziNormal & ".docm", True
'________________________________________

    'Oluşturulacak dosyayı aç
    On Error Resume Next
    Set objWord = GetObject(, "Word.Application")
    Set objWord = GetObject(, "Word.Application")
    Set objWord = GetObject(, "Word.Application")
    Set objWord = GetObject(, "Word.Application")
    Set objWord = GetObject(, "Word.Application")
    If objWord Is Nothing Then
        'MsgBox "Dosya oluşturmada CreateObject methodu kullanılacak."
        Set objWord = CreateObject("Word.Application")
        objWord.Visible = False
    End If
    objWord.Documents.Open FileName:=DestOpUserFolder & ReNameUstYaziNormal & ".docm"
    objWord.Visible = True
    objWord.Activate 'Ekrana getirir.
    'objDoc.Activate 'Ekrana getirmez.
    objWord.Application.WindowState = wdWindowStateMaximize
    Set objDoc = GetObject(DestOpUserFolder & ReNameUstYaziNormal & ".docm")
'________________________________________

    'Birim
    Birimx = UCase(Replace(Replace(Worksheets(2).Cells(6, 99).Value, "i", "I"), "ı", "I")) & " UNIT"
    objDoc.Tables(1).Cell(Row:=1, Column:=1).Range.Text = Birimx

    Birimx = UCase(Replace(Replace(Worksheets(2).Cells(6, 99).Value, "i", "I"), "ı", "I"))
    'Üst yazı tarihi
    objDoc.Tables(1).Cell(Row:=1, Column:=2).Range.Text = Birimx & ", " & FormatEnglishDate(Cells(ActiveCell.Row, 83).Value) '(Format(Cells(ActiveCell.Row, 83).Value, "d mmmm yyyy"))
    'Yazı no
    objDoc.Tables(1).Cell(Row:=4, Column:=2).Range.Text = Cells(ActiveCell.Row, 84).Value
    'Sigorta türü
    Ek2 = Cells(ActiveCell.Row, 71).Value
    Ek2 = Right(Ek2, Len(Ek2) - InStr(Ek2, "/"))
    objDoc.Tables(1).Cell(Row:=7, Column:=2).Range.Text = Ek2

    'Muhatap
    IlBuyukHarf = UCase(Replace(Replace(Cells(ActiveCell.Row, 19).Value, "i", "I"), "ı", "I"))
    IlKucukHarf = Cells(ActiveCell.Row, 19).Value
    If Cells(ActiveCell.Row, 20).Value <> "" Then
        IlceBuyukHarf = UCase(Replace(Replace(Cells(ActiveCell.Row, 20).Value, "i", "I"), "ı", "I"))
        IlceKucukHarf = Cells(ActiveCell.Row, 20).Value
    Else
        IlceBuyukHarf = ""
        IlceKucukHarf = ""
    End If
    If Cells(ActiveCell.Row, 20).Value <> "" Then
        Bolum3 = IlceKucukHarf & "/" & IlBuyukHarf
    Else
        Bolum3 = IlBuyukHarf
    End If
    

    'YENİ MUHATAP TEMASI
    
    M2 = False
    M3 = False
    M4 = False
'    Ifv = False
'    Ify = False
    'TipB = False
    BStr = 10
    '1. sayfadan sonraki üst bilgi tanımları
    ustbilgitarih = (Format(Cells(ActiveCell.Row, 83).Value, "dd.mm.yyyy"))
    ustbilgisayi = Cells(ActiveCell.Row, 84).Value
    
    '4'lük
    If Cells(ActiveCell.Row, 48).Value <> "" Then
        If Cells(ActiveCell.Row, 47).Value = "Provincial Directorate B" Or Cells(ActiveCell.Row, 47).Value = "Provincial Directorate C" Or _
        Cells(ActiveCell.Row, 47).Value = "Provincial Directorate D" Or Cells(ActiveCell.Row, 47).Value = "Provincial Directorate E" Then 'VALİLİK
            objDoc.Tables(1).Cell(Row:=BStr, Column:=1).Range.Text = IlBuyukHarf & " PROVINCIAL GOVERNORSHIP"
            objDoc.Tables(1).Cell(Row:=BStr + 1, Column:=1).Range.Text = Cells(ActiveCell.Row, 47).Value  'UCase(Replace(Replace(Cells(ActiveCell.Row, 47).Value, "i", "I"), "ı", "I"))
            objDoc.Tables(1).Cell(Row:=BStr + 2, Column:=1).Range.Text = "(" & Cells(ActiveCell.Row, 48).Value & ")"
            objDoc.Tables(1).Cell(Row:=BStr + 3, Column:=1).Range.Text = Bolum3
            objDoc.Tables(1).Cell(Row:=BStr + 3, Column:=1).Range.Font.Underline = wdUnderlineSingle
            M4 = True
            
            ustbilgimuhatap = IlKucukHarf & " Provincial Governorship " & Cells(ActiveCell.Row, 47).Value
        ElseIf Cells(ActiveCell.Row, 47).Value = "District Directorate B" Or Cells(ActiveCell.Row, 47).Value = "District Directorate C" Or _
        Cells(ActiveCell.Row, 47).Value = "District Directorate D" Or Cells(ActiveCell.Row, 47).Value = "District Directorate E" Then 'KAYMAKAMLIK
            objDoc.Tables(1).Cell(Row:=BStr, Column:=1).Range.Text = IlceBuyukHarf & " DISTRICT GOVERNORSHIP"
            objDoc.Tables(1).Cell(Row:=BStr + 1, Column:=1).Range.Text = Cells(ActiveCell.Row, 47).Value 'UCase(Replace(Replace(Cells(ActiveCell.Row, 47).Value, "i", "I"), "ı", "I"))
            objDoc.Tables(1).Cell(Row:=BStr + 2, Column:=1).Range.Text = "(" & Cells(ActiveCell.Row, 48).Value & ")"
            objDoc.Tables(1).Cell(Row:=BStr + 3, Column:=1).Range.Text = Bolum3
            objDoc.Tables(1).Cell(Row:=BStr + 3, Column:=1).Range.Font.Underline = wdUnderlineSingle
            M4 = True
            
            ustbilgimuhatap = IlceKucukHarf & " District Governorship " & Cells(ActiveCell.Row, 47).Value
        ElseIf InStr(Cells(ActiveCell.Row, 47).Value, "General Directorate") <> 0 Or InStr(Cells(ActiveCell.Row, 47).Value, "Regional Directorate") <> 0 Then
            objDoc.Tables(1).Cell(Row:=BStr, Column:=1).Range.Text = UCase(Replace(Replace(ThisWorkbook.Worksheets(2).Cells(6, 111).Value, "i", "I"), "ı", "I"))
            objDoc.Tables(1).Cell(Row:=BStr + 1, Column:=1).Range.Text = Cells(ActiveCell.Row, 47).Value 'UCase(Replace(Replace(Cells(ActiveCell.Row, 47).Value, "i", "I"), "ı", "I"))
            objDoc.Tables(1).Cell(Row:=BStr + 2, Column:=1).Range.Text = "(" & Cells(ActiveCell.Row, 48).Value & ")"
            objDoc.Tables(1).Cell(Row:=BStr + 3, Column:=1).Range.Text = Bolum3
            objDoc.Tables(1).Cell(Row:=BStr + 3, Column:=1).Range.Font.Underline = wdUnderlineSingle
            M4 = True
            
            ustbilgimuhatap = WorksheetFunction.Proper(ThisWorkbook.Worksheets(2).Cells(6, 111).Value) & " " & Cells(ActiveCell.Row, 47).Value
        Else 'YARGI 3'lük
            Bolum1 = UCase(Replace(Replace(Cells(ActiveCell.Row, 47).Value, "i", "I"), "ı", "I"))
            If InStr(Bolum1, "X.X. ") > 0 Then
                objDoc.Tables(1).Cell(Row:=BStr, Column:=1).Range.Text = Mid(Bolum1, 6, Len(Bolum1))
                objDoc.Tables(1).Cell(Row:=BStr + 1, Column:=1).Range.Text = "(" & Cells(ActiveCell.Row, 48).Value & ")"
                objDoc.Tables(1).Cell(Row:=BStr + 2, Column:=1).Range.Text = Bolum3
                objDoc.Tables(1).Cell(Row:=BStr + 2, Column:=1).Range.Font.Underline = wdUnderlineSingle
                M3 = True
                
                ustbilgimuhatap = Mid(Cells(ActiveCell.Row, 47).Value, 6, Len(Cells(ActiveCell.Row, 47).Value))
            Else
                objDoc.Tables(1).Cell(Row:=BStr, Column:=1).Range.Text = Bolum1
                objDoc.Tables(1).Cell(Row:=BStr + 1, Column:=1).Range.Text = "(" & Cells(ActiveCell.Row, 48).Value & ")"
                objDoc.Tables(1).Cell(Row:=BStr + 2, Column:=1).Range.Text = Bolum3
                objDoc.Tables(1).Cell(Row:=BStr + 2, Column:=1).Range.Font.Underline = wdUnderlineSingle
                M3 = True

                ustbilgimuhatap = Cells(ActiveCell.Row, 47).Value
            End If
        End If
    End If
    '3'lük
    If Cells(ActiveCell.Row, 48).Value = "" Then
        If Cells(ActiveCell.Row, 47).Value = "Provincial Directorate B" Or Cells(ActiveCell.Row, 47).Value = "Provincial Directorate C" Or _
        Cells(ActiveCell.Row, 47).Value = "Provincial Directorate D" Or Cells(ActiveCell.Row, 47).Value = "Provincial Directorate E" Then 'VALİLİK
            objDoc.Tables(1).Cell(Row:=BStr, Column:=1).Range.Text = IlBuyukHarf & " PROVINCIAL GOVERNORSHIP"
            objDoc.Tables(1).Cell(Row:=BStr + 1, Column:=1).Range.Text = "(" & Cells(ActiveCell.Row, 47).Value & ")"  'UCase(Replace(Replace(Cells(ActiveCell.Row, 47).Value, "i", "I"), "ı", "I"))
            objDoc.Tables(1).Cell(Row:=BStr + 2, Column:=1).Range.Text = Bolum3
            objDoc.Tables(1).Cell(Row:=BStr + 2, Column:=1).Range.Font.Underline = wdUnderlineSingle
            M3 = True
            
            ustbilgimuhatap = IlKucukHarf & " Provincial Governorship " & Cells(ActiveCell.Row, 47).Value
        ElseIf Cells(ActiveCell.Row, 47).Value = "District Directorate B" Or Cells(ActiveCell.Row, 47).Value = "District Directorate C" Or _
        Cells(ActiveCell.Row, 47).Value = "District Directorate D" Or Cells(ActiveCell.Row, 47).Value = "District Directorate E" Then
            objDoc.Tables(1).Cell(Row:=BStr, Column:=1).Range.Text = IlceBuyukHarf & " DISTRICT GOVERNORSHIP"
            objDoc.Tables(1).Cell(Row:=BStr + 1, Column:=1).Range.Text = "(" & Cells(ActiveCell.Row, 47).Value & ")"  'UCase(Replace(Replace(Cells(ActiveCell.Row, 47).Value, "i", "I"), "ı", "I"))
            objDoc.Tables(1).Cell(Row:=BStr + 2, Column:=1).Range.Text = Bolum3
            objDoc.Tables(1).Cell(Row:=BStr + 2, Column:=1).Range.Font.Underline = wdUnderlineSingle
            M3 = True
            
            ustbilgimuhatap = IlceKucukHarf & " District Governorship " & Cells(ActiveCell.Row, 47).Value
        ElseIf InStr(Cells(ActiveCell.Row, 47).Value, "General Directorate") <> 0 Or InStr(Cells(ActiveCell.Row, 47).Value, "Regional Directorate") <> 0 Then
            objDoc.Tables(1).Cell(Row:=BStr, Column:=1).Range.Text = UCase(Replace(Replace(ThisWorkbook.Worksheets(2).Cells(6, 111).Value, "i", "I"), "ı", "I"))
            objDoc.Tables(1).Cell(Row:=BStr + 1, Column:=1).Range.Text = "(" & Cells(ActiveCell.Row, 47).Value & ")"  'UCase(Replace(Replace(Cells(ActiveCell.Row, 47).Value, "i", "I"), "ı", "I"))
            objDoc.Tables(1).Cell(Row:=BStr + 2, Column:=1).Range.Text = Bolum3
            objDoc.Tables(1).Cell(Row:=BStr + 2, Column:=1).Range.Font.Underline = wdUnderlineSingle
            M3 = True

            ustbilgimuhatap = WorksheetFunction.Proper(ThisWorkbook.Worksheets(2).Cells(6, 111).Value) & " " & Cells(ActiveCell.Row, 47).Value
        Else 'YARGI 2'lik
            Bolum1 = UCase(Replace(Replace(Cells(ActiveCell.Row, 47).Value, "i", "I"), "ı", "I"))
            If InStr(Bolum1, "X.X. ") > 0 Then
                objDoc.Tables(1).Cell(Row:=BStr, Column:=1).Range.Text = Mid(Bolum1, 6, Len(Bolum1))
                objDoc.Tables(1).Cell(Row:=BStr + 1, Column:=1).Range.Text = Bolum3
                objDoc.Tables(1).Cell(Row:=BStr + 1, Column:=1).Range.Font.Underline = wdUnderlineSingle
                M2 = True

                ustbilgimuhatap = Mid(Cells(ActiveCell.Row, 47).Value, 6, Len(Cells(ActiveCell.Row, 47).Value))
            Else
                objDoc.Tables(1).Cell(Row:=BStr, Column:=1).Range.Text = Bolum1
                objDoc.Tables(1).Cell(Row:=BStr + 1, Column:=1).Range.Text = Bolum3
                objDoc.Tables(1).Cell(Row:=BStr + 1, Column:=1).Range.Font.Underline = wdUnderlineSingle
                M2 = True

                ustbilgimuhatap = Cells(ActiveCell.Row, 47).Value
            End If
        End If
    End If

    'Üst yazı gövde metni
    'tipAnın/tipAların (TipA adedi)
    AdetTopla = Application.Sum(Range(Cells(IlkSira, 58), Cells(SonSira, 58)))
    If AdetTopla = 1 Then
        Ek1 = "item"
        Ek2 = "has"
    ElseIf AdetTopla > 1 Then
        Ek1 = "items"
        Ek2 = "have"
    End If

    ' Report body (English version)
    ' Report(s) control
    Dim ReportList() As String
    ' Rapor numaralarını diziye al
    y = 0
    For i = IlkSira To SonSira
        If Cells(i, 7).Value <> "" Then
            y = y + 1
            ReDim Preserve ReportList(1 To y)
            ReportList(y) = Cells(i, 13).Value
        End If
    Next i
    
    ' Listeyi biçimlendir
    Select Case y
        Case 1
            Ek3 = ReportList(1)
        Case 2
            Ek3 = ReportList(1) & " and " & ReportList(2)
        Case Else
            Ek3 = ""
            For i = 1 To y - 1
                Ek3 = Ek3 & ReportList(i) & ", "
            Next i
            Ek3 = Ek3 & "and " & ReportList(y)
    End Select

    If y = 1 Then
        Ek4 = "Report 2.1"
        'Ek5 = "is"
    Else
        Ek4 = "Report 2.1s"
        'Ek5 = "are"
    End If
    Ek6 = Cells(IlkSira, 218).Value 'report date
    
    Bolum2 = "The situation has been reported to the relevant financial unit, and copies of the official statement and the letter addressed to the mentioned unit are enclosed herewith, along with the " & Ek1 & " in question and the " & Ek4 & " dated " & Ek6 & " and numbered " & Ek3 & ", for your information."
    
    'Technique A var/yok (SADECE RAPOR3 işleminde olan kısım)
    StrTeknik_ANotu = ""
    For i = IlkSira To SonSira
        If Cells(i, 212).Value = "invalid" And Left(Cells(i, 213).Value, 11) = "Technique A" Then
            StrTeknik_ANotu = "Furthermore, since the Type A " & Ek1 & " determined to be invalid based on the assessment of xxxxxx xxxxxx xxxxxx, the preparation of the related Report 2.2 is only possible upon submission of the " & Ek1 & " to the XXX Directorate."
            GoTo Teknik_AOk1
        End If
    Next i
Teknik_AOk1:
    
    Ek4 = Cells(ActiveCell.Row, 35).Value 'Teslimat tarihi
    Ek5 = Cells(ActiveCell.Row, 30).Value 'FinansalBirim
    Ek6 = Cells(ActiveCell.Row, 26).Value 'Tema no
    Ek7 = Cells(ActiveCell.Row, 32).Value 'bildirime konu birim
    Ek8 = Cells(ActiveCell.Row, 31).Value 'Teslimatı yapan birim
    
    Bolum1 = "As a result of the initial inspection of the " & Ek1 & " delivered to our unit on " & Ek4 & " by " & Ek5 & " – " & Ek8 & " for the purpose of xxxxx xxxxx xxxxx, " & _
    "the Type A " & Ek1 & "—listed in the attached statement, associated with " & Ek5 & " – " & Ek7 & ", and assigned the theme number " & Ek6 & " by our office— " & _
    Ek2 & " been evaluated as invalid."

    Bolum5 = Chr(9) & "Respectfully submitted for your information."
    
    If StrTeknik_ANotu <> "" Then
        Birlestir = Chr(9) & Bolum1 & vbNewLine & Chr(9) & Bolum2 & vbNewLine & Chr(9) & StrTeknik_ANotu
    Else
        Birlestir = Chr(9) & Bolum1 & vbNewLine & Chr(9) & Bolum2
    End If
    objDoc.Tables(2).Cell(Row:=1, Column:=1).Range.Text = Birlestir
    objDoc.Tables(2).Cell(Row:=3, Column:=1).Range.Text = Bolum5

    
    'Üst yazı notu (TipA XXXMudsı notu)
    XXXMudNotu = False
    If Cells(ActiveCell.Row, 215).Value = "Yes" Then
        XXXMudNotu = True
    Else
        XXXMudNotu = False
        objDoc.Tables(2).Cell(Row:=2, Column:=1).Range.Text = ""
        'objDoc.Tables(2).Rows(2).Delete
    End If
    
    'imzalar
    objDoc.Tables(3).Cell(Row:=4, Column:=2).Range.Text = Cells(ActiveCell.Row, 205).Value 'Ad Soyad1
    objDoc.Tables(3).Cell(Row:=5, Column:=2).Range.Text = Cells(ActiveCell.Row, 206).Value 'Unvan1
    objDoc.Tables(3).Cell(Row:=4, Column:=3).Range.Text = Cells(ActiveCell.Row, 208).Value 'Ad Soyad2
    objDoc.Tables(3).Cell(Row:=5, Column:=3).Range.Text = Cells(ActiveCell.Row, 209).Value 'Unvan2
    
    'Ekler
    objDoc.Tables(3).Cell(Row:=8, Column:=1).Range.Text = "Attachment:"
    Ek2 = Cells(ActiveCell.Row, 71).Value 'Kapalı Package A
    Ek2 = Left(Ek2, InStr(Ek2, "/") - 1)
    If Cells(ActiveCell.Row, 72).Value > 1 Then objDoc.Tables(3).Cell(Row:=8, Column:=2).Range.Text = " 1) Enclosed " & Ek2 & " (" & Cells(ActiveCell.Row, 72).Value & " pieces)"
    If Cells(ActiveCell.Row, 72).Value < 2 Then objDoc.Tables(3).Cell(Row:=8, Column:=2).Range.Text = " 1) Enclosed " & Ek2 & " (" & Cells(ActiveCell.Row, 72).Value & " piece)"

    x = Application.Sum(Range(Cells(IlkSira, 171), Cells(SonSira, 171))) 'Statement 2 toplam sayfa sayısı
    If x > 1 Then objDoc.Tables(3).Cell(Row:=9, Column:=2).Range.Text = " 2) Statement 2 (" & x & " pages)"
    If x < 2 Then objDoc.Tables(3).Cell(Row:=9, Column:=2).Range.Text = " 2) Statement 2 (" & x & " page)"
    
    x = Application.Sum(Range(Cells(IlkSira, 172), Cells(SonSira, 172))) 'FinansalBirim üst yazı toplam sayfa sayısı
    If x > 1 Then objDoc.Tables(3).Cell(Row:=10, Column:=2).Range.Text = " 3) Photocopy of the Letter to Relevant Unit (" & x & " pages)"
    If x < 2 Then objDoc.Tables(3).Cell(Row:=10, Column:=2).Range.Text = " 3) Photocopy of the Letter to Relevant Unit (" & x & " page)"

    x = Application.Sum(Range(Cells(IlkSira, 174), Cells(SonSira, 174)))  'Rapor 3.2 toplam sayfa sayısı
    If y = 1 Then
        If x > 1 Then objDoc.Tables(3).Cell(Row:=11, Column:=2).Range.Text = " 4) Report 2.1 (" & x & " pages)"
        If x < 2 Then objDoc.Tables(3).Cell(Row:=11, Column:=2).Range.Text = " 4) Report 2.1 (" & x & " page)"
    Else
        objDoc.Tables(3).Cell(Row:=11, Column:=2).Range.Text = " 4) Report 2.1 (" & y & " reports, " & "total of " & x & " pages)"
    End If
    
    x = Application.Sum(Range(Cells(IlkSira, 169), Cells(SonSira, 169))) + Cells(ActiveCell.Row, 50).Value 'Rapor3 Tutanağı
    objDoc.Tables(3).Cell(Row:=12, Column:=2).Range.Text = " 5) Attached Statement (total of " & x & " pages)"

    x = Application.Sum(Range(Cells(IlkSira, 49), Cells(SonSira, 49))) 'Label Fotokopisi
    If x > 1 Then objDoc.Tables(3).Cell(Row:=13, Column:=2).Range.Text = " 6) Photocopy of the Label (" & x & " pages)"
    If x < 2 Then objDoc.Tables(3).Cell(Row:=13, Column:=2).Range.Text = " 6) Photocopy of the Label (" & x & " page)"


    '__________________________DİNAMİK SAYFA YAPISI
    
    'İlgi ve gövde metninin kaç satırdan oluştuğu bilgisinin elde edilmesinde kullanılıyor.
    'Set IlgiRng = objDoc.Tables(1).Cell(Row:=15, Column:=3).Range
    Set Govde1Rng = objDoc.Tables(2).Cell(Row:=1, Column:=1).Range
    Set Govde2Rng = objDoc.Tables(2).Cell(Row:=2, Column:=1).Range
    
    'IlgiStrSay = Govde1Rng.Information(wdFirstCharacterLineNumber) - IlgiRng.Information(wdFirstCharacterLineNumber) - 1
    Govde1StrSay = Govde2Rng.Information(wdFirstCharacterLineNumber) - Govde1Rng.Information(wdFirstCharacterLineNumber)
    'MsgBox "İlgi Satır Sayısı: " & IlgiStrSay & vbNewLine & vbNewLine & "Gövde 1 Satır Sayısı: " & Govde1StrSay
    'IlgiFarkSay = IlgiStrSay - 1
    IlgiFarkSay = 0
    Govde1FarkSay = Govde1StrSay - 4 + 1 'Varsayılan 2 paragrafta (1 rowda) toplam 3 satıra göre sıfırlandı. (+1) 1 adet ilave ek geldiği için gövdeden düzeltme yapıldı
    CokluSayfa = 0
    'MsgBox Govde1FarkSay
 
   
    'Dinamik sayfa düzeni
    If StrTeknik_ANotu <> "" Then '11-12 / M2,  M3=M2+1, M4=M3+1
        If XXXMudNotu = True Then
            'yazı tipini küçült
            Set Govde1Rng = objDoc.Tables(2).Cell(Row:=1, Column:=1).Range
            Set Govde2Rng = objDoc.Tables(2).Cell(Row:=2, Column:=1).Range

            objDoc.Range.Font.Size = 11
            objDoc.Tables(1).Cell(Row:=5, Column:=1).Range.Font.Size = 13
            objDoc.Tables(1).Cell(Row:=2, Column:=1).Range.Font.Size = 9
            
            Govde1FarkSay = Govde1StrSay + 1
        Else '10-11 / M2
            objDoc.Tables(2).Rows(2).Delete 'Boş satırı sil
            Govde1FarkSay = Govde1StrSay + 0 '- 2 '- 4 + 1 + 1
        End If
    Else
        If XXXMudNotu = True Then '10-11 / M2
            Govde1FarkSay = Govde1StrSay + 3 '- 4 + 1 + 1 + 4
        Else '7-8 / M2
            objDoc.Tables(2).Rows(2).Delete 'Boş satırı sil
            Govde1FarkSay = Govde1StrSay + 0 '- 3 '- 4 + 1 + 1
        End If
    End If


    If M4 = True Then
        For i = 1 To 2
            objDoc.Tables(1).Rows(14).Delete 'Muhatap sonrası
        Next i
        
        If IlgiFarkSay + Govde1FarkSay < 17 Then
            Govde1FarkSay = Govde1FarkSay + 2
            If IlgiFarkSay + Govde1FarkSay > 8 Then
                For i = 9 To IlgiFarkSay + Govde1FarkSay
                    If i = 9 Then
                        objDoc.Tables(3).Rows(6).Delete 'Ek öncesi
                    ElseIf i = 10 Then
                        objDoc.Tables(3).Rows(6).Delete 'Ek öncesi
                    ElseIf i = 11 Then
                        objDoc.Tables(3).Rows(1).Delete 'İmza öncesi
                    ElseIf i = 12 Then
                        objDoc.Tables(3).Rows(1).Delete 'İmza öncesi
                    ElseIf i = 13 Then
                        objDoc.Tables(1).Rows(8).Delete 'Muhatap öncesi
                    ElseIf i = 14 Then
                        objDoc.Tables(1).Rows(8).Delete 'Muhatap öncesi
                    End If
                Next i
            Else
                If (IlgiFarkSay + Govde1FarkSay) < 8 Then
                    For i = 1 To 8 - (IlgiFarkSay + Govde1FarkSay)
                        'Boş satır ekle
                        With objDoc.Tables(3)
                            .Rows.Add BeforeRow:=objDoc.Tables(3).Rows(8)
                        End With
                    Next i
                End If
            End If
        Else
            CokluSayfa = 1
        End If
    ElseIf M3 = True Then
        For i = 1 To 3
            objDoc.Tables(1).Rows(13).Delete 'Muhatap sonrası
        Next i
        
        If IlgiFarkSay + Govde1FarkSay < 16 Then
            Govde1FarkSay = Govde1FarkSay + 1
            'MsgBox IlgiFarkSay + Govde1FarkSay
            If IlgiFarkSay + Govde1FarkSay > 8 Then
                For i = 9 To IlgiFarkSay + Govde1FarkSay
                    If i = 9 Then
                        objDoc.Tables(3).Rows(6).Delete 'Ek öncesi
                    ElseIf i = 10 Then
                        objDoc.Tables(3).Rows(6).Delete 'Ek öncesi
                    ElseIf i = 11 Then
                        objDoc.Tables(3).Rows(1).Delete 'İmza öncesi
                    ElseIf i = 12 Then
                        objDoc.Tables(3).Rows(1).Delete 'İmza öncesi
                    ElseIf i = 13 Then
                        objDoc.Tables(1).Rows(8).Delete 'Muhatap öncesi
                    ElseIf i = 14 Then
                        objDoc.Tables(1).Rows(8).Delete 'Muhatap öncesi
                    End If
                Next i
            Else
                If (IlgiFarkSay + Govde1FarkSay) < 8 Then
                    For i = 1 To 8 - (IlgiFarkSay + Govde1FarkSay)
                        'Boş satır ekle
                        With objDoc.Tables(3)
                            .Rows.Add BeforeRow:=objDoc.Tables(3).Rows(8)
                        End With
                    Next i
                End If
            End If
        Else
            CokluSayfa = 1
        End If
    ElseIf M2 = True Then
        For i = 1 To 4
            objDoc.Tables(1).Rows(12).Delete 'Muhatap sonrası
        Next i
'        MsgBox IlgiFarkSay + Govde1FarkSay
        If IlgiFarkSay + Govde1FarkSay < 15 Then
            If IlgiFarkSay + Govde1FarkSay > 8 Then
                For i = 9 To IlgiFarkSay + Govde1FarkSay
                    If i = 9 Then
                        objDoc.Tables(3).Rows(6).Delete 'Ek öncesi
                    ElseIf i = 10 Then
                        objDoc.Tables(3).Rows(6).Delete 'Ek öncesi
                    ElseIf i = 11 Then
                        objDoc.Tables(3).Rows(1).Delete 'İmza öncesi
                    ElseIf i = 12 Then
                        objDoc.Tables(3).Rows(1).Delete 'İmza öncesi
                    ElseIf i = 13 Then
                        objDoc.Tables(1).Rows(8).Delete 'Muhatap öncesi
                    ElseIf i = 14 Then
                        objDoc.Tables(1).Rows(8).Delete 'Muhatap öncesi
                    End If
                Next i
            Else
                If (IlgiFarkSay + Govde1FarkSay) < 8 Then
                    For i = 1 To 8 - (IlgiFarkSay + Govde1FarkSay)
                        'Boş satır ekle
                        With objDoc.Tables(3)
                            .Rows.Add BeforeRow:=objDoc.Tables(3).Rows(8)
                        End With
                    Next i
                End If
            End If
        Else
            CokluSayfa = 1
        End If
    End If
    
    objDoc.Paragraphs(objDoc.Paragraphs.Count).Range.Font.Size = 1
    
    'Footer
    If objDoc.ActiveWindow.View.SplitSpecial <> wdPaneNone Then
        objDoc.ActiveWindow.Panes(2).Close
    End If
    If objDoc.ActiveWindow.ActivePane.View.Type = wdNormalView Or objDoc.ActiveWindow.ActivePane.View.Type = wdOutlineView Then
        objDoc.ActiveWindow.ActivePane.View.Type = wdPrintView
    End If

    With objDoc
        With .Sections(1).Headers(wdHeaderFooterPrimary).Range
            .InsertAfter Text:="Page "
            .Fields.Add Range:=.Characters.Last, Type:=wdFieldEmpty, Text:="PAGE", PreserveFormatting:=False
            .InsertAfter Text:=" of the letter dated " & ustbilgitarih & _
                             ", reference number " & ustbilgisayi & _
                             ", addressed to the " & ustbilgimuhatap
        End With
    End With
    
    objDoc.ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument


    'PARAF EKLEME İŞLEMİ
    objDoc.Sections(1).Footers(wdHeaderFooterFirstPage).Range.Tables(1).Cell(Row:=1, Column:=2).Range.Text = AdSoyadParaf & vbNewLine & _
                                                                                       UnvanParaf & vbNewLine & _
                                                                                       TelParaf
    objDoc.Sections(1).Footers(wdHeaderFooterPrimary).Range.Tables(1).Cell(Row:=1, Column:=2).Range.Text = AdSoyadParaf & vbNewLine & _
                                                                                       UnvanParaf & vbNewLine & _
                                                                                       TelParaf

Tekrarla:
    'Sayfa sayısı kaydet komutuna bağlandı.
    objDoc.Save
    'Sayfayı text dosyasından çek
    TxtFileUstYazi = DestOpUserFolder & "Cover Letter Page Count.txt"
    #If UseADO Then
        Const adReadLine = -2&
        With CreateObject("ADODB.Stream")
            .Open
            .LoadFromFile TxtFileUstYazi
            Do Until .EOS
                TotalSayfaUstYazi = .ReadText(adReadLine)
            Loop
            .Close
        End With
    #Else 'UseFSO
        Const UstYaziTristateTrue = -1&
        With CreateObject("Scripting.FileSystemObject")
            With .OpenTextFile(TxtFileUstYazi, Format:=UstYaziTristateTrue)
                Do Until .AtEndOfStream
                    TotalSayfaUstYazi = .ReadLine
                Loop
                .Close
            End With
        End With
    #End If
    'MsgBox "Üst Yazı: " & TotalSayfaUstYazi

    'Üst Yazı sayfa sayısı
    Cells(ActiveCell.Row, 173).Value = TotalSayfaUstYazi

    objWord.Visible = True
    objWord.Activate

    If TumDoc = True Then
        objWord.Activate
'        For i = 1 To SayPrt
'            objDoc.PrintOut
'        Next i
        objDoc.PrintOut Background:=False, Copies:=SayPrt
        'objWord.Documents.Save
        objDoc.Close SaveChanges:=False
        objWord.Visible = False
    End If

    'Nesneleri temizle
    Set fso = Nothing
    Set objWord = Nothing
    Set objDoc = Nothing

Son:

If IlceSakla <> "" Then
    Cells(ActiveCell.Row, 20).Value = IlceSakla
End If

'Worksheets(5).Activate
Application.ScreenUpdating = True
Application.DisplayAlerts = True
'Application.EnableEvents = True

'End If

End Sub

Sub Rapor3_2TutanakTipB()

Dim DestOperasyon As String, SourceRapor3_1Normal As String, AutoPath As String, ReNameRapor3_1 As String
Dim fso As Object, objWord As Object, objDoc As Object
Dim OpenKontrolName As String, OpenControl As String, DestOpUserFolder As String, DestOpUserFolderName As String
Dim ContSay As Integer, KontrolFile As String, x As Long, y As Long, i As Long, j As Long
Dim IlkSiraBul As Range, SonSiraBul As Range, IlkSira As Long, SonSira As Long
Dim GelenTema As String, SourceDL As String, ReNameDL As String
Dim TxtFileRapor3_1 As String, TotalSayfaRapor3_1 As String, TxtFileDokum As String, TotalSayfaDokum As String
Dim DokumFileTxt As Object, DokumSayfaGonder As String
Dim SourceRaporNormal As String, SourceTutanak2Normal As String, SourceUstYaziNormal As String
Dim Bolum1 As String, Bolum2 As String, Bolum3 As String, Bolum4 As String, Bolum5 As String, Bolum6 As String
Dim Ek1 As String, Ek2 As String, Ek3 As String, Ek4 As String, Ek5 As String, Ek6 As String, Birlestir As String

Dim ReNameRaporNormal As String, SiraNoIlkSatir As Long, RaporIlgi As String
Dim AltRaporNoIlk As Long, AltRaporNoSon As Long, AdetTopla As Long
Dim TekCogulTipA As String, k As Integer
Dim MyRange As Object, TxtFileRapor As String, TotalSayfaRapor As String
Dim ReNameTutanak2Normal As String, GidenTema As String
Dim TxtFileTutanak2 As String, TotalSayfaTutanak2 As String
Dim ReNameUstYaziNormal As String, PaketTipi As String
Dim TxtFileUstYazi As String, TotalSayfaUstYazi As String
Dim Birimx As String, UserName As String, SourceRapor3_1Farkli As String
Dim Kolluk As String


'TUTANAK1 için prosedürü başlat
'    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    'Application.EnableEvents = False
    'Worksheets(5).Unprotect Password:="123"
    
    'Pathfinder...
    AutoPath = ThisWorkbook.Path
    DestOperasyon = AutoPath & "\System Files\Operation\"
    'TUTANAK1 TANIMLARI
    SourceRapor3_1Normal = AutoPath & "\System Files\System Templates\Report 3 Statements\Report 3.2 – Type B.docm"
    
    'Birden fazla kullanıcının operasyon klasöründe çakışmasını önlemek için operasyon klasörünün içinde yeni klasör aç.
    DestOpUserFolderName = "Operation-" & Left(ThisWorkbook.name, InStr(ThisWorkbook.name, ".") - 1) '& " " & Format(Now(), "ddmmyyyyhhmmss")
    DestOpUserFolder = DestOperasyon & DestOpUserFolderName & "\"
    
    'Check System Files folder existence.
    If Not Dir(AutoPath & "\System Files\", vbDirectory) <> vbNullString Then
        MsgBox "Cannot access the folder " & AutoPath & "\System Files\" & ". The folder may have been renamed or deleted.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo EndProcess
    End If
    
    'Check Operations folder existence.
    If Not Dir(DestOperasyon, vbDirectory) <> vbNullString Then
        MsgBox "Cannot access the folder " & DestOperasyon & ". The folder may have been renamed or deleted.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo EndProcess
    End If
    
    'Check folder names.
    If Not Dir(SourceRapor3_1Normal, vbDirectory) <> vbNullString Then
        MsgBox "Cannot access the folder " & SourceRapor3_1Normal & ". The folder or files inside may have been renamed or deleted.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo EndProcess
    End If

    'Operation klasörü içinde kullanıcı modülü klasörü yoksa oluştur.
    If Not Dir(DestOpUserFolder, vbDirectory) <> vbNullString Then
        MkDir DestOpUserFolder
    End If
    
    'On Error Resume Next
    ReNameRapor3_1 = Cells(ActiveCell.Row, 5).Value & "-" & Cells(6, 6).Value
    'ReNameDL = Cells(ActiveCell.Row, 5).Value & "-" & "Dispatch List"
'________________________________________

    
    'Close the all Word application
    Call ModuleReport3.OpenWordControl
    
    'Operation klasöründeki docm uzantılı word dosyalarından açık olanları kapat ve temizle.
    OpenKontrolName = Dir(DestOpUserFolder & "*.docm")
    Do While OpenKontrolName <> ""
        OpenControl = IsFileOpen(DestOpUserFolder & OpenKontrolName)
        If OpenControl = True Then 'Açıksa
            On Error Resume Next
            Set objWord = GetObject(, "Word.Application")
            Set objWord = GetObject(, "Word.Application")
            Set objWord = GetObject(, "Word.Application")
            Set objWord = GetObject(, "Word.Application")
            Set objWord = GetObject(, "Word.Application")
            objWord.Quit SaveChanges:=True
            'MsgBox "Dosya OpenKontrol methodu ile kapatıldı."

        End If
        OpenKontrolName = Dir()
    Loop

'    Call ModuleReport3.OpenWordControl
     
    Set objWord = Nothing
    Set objDoc = Nothing

'________________________________________

    On Error Resume Next
'    Klasörün içindeki tüm dosyaları sil (txt, docm vb.)
    ContSay = 0
    KontrolFile = Dir(DestOpUserFolder & "*.???")
    Do While KontrolFile <> ""
        ContSay = ContSay + 1
        KontrolFile = Dir()
    Loop
    If ContSay > 0 Then
        Kill DestOpUserFolder & "*.???"
    End If
    
    'Dosyayı şablondan operasyon klasörüne kopyala ve adını değiştir.
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.CopyFile (SourceRapor3_1Normal), DestOpUserFolder & ReNameRapor3_1 & ".docm", True
    
'________________________________________

    'Oluşturulacak dosyayı aç
    On Error Resume Next
    Set objWord = GetObject(, "Word.Application")
    Set objWord = GetObject(, "Word.Application")
    Set objWord = GetObject(, "Word.Application")
    Set objWord = GetObject(, "Word.Application")
    Set objWord = GetObject(, "Word.Application")
    If objWord Is Nothing Then
        'MsgBox "Dosya oluşturmada CreateObject methodu kullanılacak."
        Set objWord = CreateObject("Word.Application")
        objWord.Visible = False
    End If
    objWord.Documents.Open FileName:=DestOpUserFolder & ReNameRapor3_1 & ".docm"
    objWord.Visible = True
    objWord.Activate 'Ekrana getirir.
    'objDoc.Activate 'Ekrana getirmez.
    objWord.Application.WindowState = wdWindowStateMaximize
    Set objDoc = GetObject(DestOpUserFolder & ReNameRapor3_1 & ".docm")
'________________________________________

    Set IlkSiraBul = Range("FG7:FG100000").Find(What:=Cells(ActiveCell.Row, 5).Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    Set SonSiraBul = Range("FH7:FH100000").Find(What:=Cells(ActiveCell.Row, 5).Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    If Not IlkSiraBul Is Nothing Then
        IlkSira = IlkSiraBul.Row
    Else
        MsgBox "The serial number could not be found, so the operation cannot be completed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    If Not SonSiraBul Is Nothing Then
        SonSira = SonSiraBul.Row
    Else
        MsgBox "The serial number could not be found, so the operation cannot be completed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    
    'Birim
    Birimx = UCase(Replace(Replace(Worksheets(2).Cells(6, 99).Value, "i", "I"), "ı", "I")) & " UNIT"
    objDoc.Tables(1).Cell(Row:=1, Column:=1).Range.Text = Birimx
    
    'Dosyada içerikleri değiştir.
    'FinansalBirim adı
    Ek1 = Cells(ActiveCell.Row, 30).Value
    Set MyRange = objDoc.Tables(1).Cell(Row:=4, Column:=1).Range
    With MyRange.Find
        .Text = "<financialUnit>"
        .Replacement.Text = Ek1
        .Execute Replace:=wdReplaceAll
    End With
    MyRange.HighlightColorIndex = wdNoHighlight
    'Teslim birimi
    Ek1 = Cells(ActiveCell.Row, 31).Value
    Set MyRange = objDoc.Tables(1).Cell(Row:=4, Column:=1).Range
    With MyRange.Find
        .Text = "<deliveryUnit>"
        .Replacement.Text = Ek1
        .Execute Replace:=wdReplaceAll
    End With
    MyRange.HighlightColorIndex = wdNoHighlight
    'Teslim tarihi
    Ek1 = Cells(ActiveCell.Row, 35).Value
    Set MyRange = objDoc.Tables(1).Cell(Row:=4, Column:=1).Range
    With MyRange.Find
        .Text = "<deliveryDate>"
        .Replacement.Text = Ek1
        .Execute Replace:=wdReplaceAll
    End With
    MyRange.HighlightColorIndex = wdNoHighlight
    'Tekil-çoğul tipB
    AdetTopla = Application.Sum(Range(Cells(IlkSira, 58), Cells(SonSira, 58)))
    If AdetTopla = 1 Then
        TekCogulTipA = "Type B item"
        Ek3 = "Type B Item"
        Ek2 = "has"
        Ek4 = "Type B item was"
    ElseIf AdetTopla > 1 Then
        TekCogulTipA = "Type B items"
        Ek3 = "Type B Items"
        Ek2 = "have"
        Ek4 = "Type B item were"
    End If
    Set MyRange = objDoc.Tables(1).Cell(Row:=4, Column:=1).Range
    With MyRange.Find
        .Text = "<typeB>"
        .Replacement.Text = TekCogulTipA
        .Execute Replace:=wdReplaceAll
    End With
    MyRange.HighlightColorIndex = wdNoHighlight
    'Sayım tarihi
    Ek1 = Cells(ActiveCell.Row, 23).Value
    Set MyRange = objDoc.Tables(1).Cell(Row:=4, Column:=1).Range
    With MyRange.Find
        .Text = "<countDate>"
        .Replacement.Text = Ek1
        .Execute Replace:=wdReplaceAll
    End With
    MyRange.HighlightColorIndex = wdNoHighlight
    'Point2/Point3
    If Cells(ActiveCell.Row, 12).Value = "Point2" Then
        Ek1 = "Point 2"
    ElseIf Cells(ActiveCell.Row, 12).Value = "Point3" Then
        Ek1 = "Point 3"
    End If
    Set MyRange = objDoc.Tables(1).Cell(Row:=4, Column:=1).Range
    With MyRange.Find
        .Text = "<location>"
        .Replacement.Text = Ek1
        .Execute Replace:=wdReplaceAll
    End With
    MyRange.HighlightColorIndex = wdNoHighlight
    Set MyRange = objDoc.Tables(1).Cell(Row:=4, Column:=1).Range
    With MyRange.Find
        .Text = "<have_has>"
        .Replacement.Text = Ek2
        .Execute Replace:=wdReplaceAll
    End With
    MyRange.HighlightColorIndex = wdNoHighlight
    
    'Kolluk
    If Cells(ActiveCell.Row, 47).Value = "Provincial Directorate B" Then
        Kolluk = Cells(ActiveCell.Row, 19).Value & " Provincial Governorship Provincial Directorate B " & Cells(ActiveCell.Row, 48).Value
    ElseIf Cells(ActiveCell.Row, 47).Value = "District Directorate B" Then
        Kolluk = Cells(ActiveCell.Row, 20).Value & " District Governorship District Directorate B " & Cells(ActiveCell.Row, 48).Value
    ElseIf Cells(ActiveCell.Row, 47).Value = "Provincial Directorate C" Then
        Kolluk = Cells(ActiveCell.Row, 19).Value & " Provincial Governorship Provincial Directorate C " & Cells(ActiveCell.Row, 48).Value
    ElseIf Cells(ActiveCell.Row, 47).Value = "District Directorate C" Then
        Kolluk = Cells(ActiveCell.Row, 20).Value & " District Governorship District Directorate C " & Cells(ActiveCell.Row, 48).Value
    ElseIf Cells(ActiveCell.Row, 47).Value = "Provincial Directorate D" Then
        Kolluk = Cells(ActiveCell.Row, 19).Value & " Provincial Governorship Provincial Directorate D " & Cells(ActiveCell.Row, 48).Value
    ElseIf Cells(ActiveCell.Row, 47).Value = "District Directorate D" Then
        Kolluk = Cells(ActiveCell.Row, 20).Value & " District Governorship District Directorate D " & Cells(ActiveCell.Row, 48).Value
    ElseIf Cells(ActiveCell.Row, 47).Value = "Provincial Directorate E" Then
        Kolluk = Cells(ActiveCell.Row, 19).Value & " Provincial Governorship Provincial Directorate E " & Cells(ActiveCell.Row, 48).Value
    ElseIf Cells(ActiveCell.Row, 47).Value = "District Directorate E" Then
        Kolluk = Cells(ActiveCell.Row, 20).Value & " District Governorship District Directorate E " & Cells(ActiveCell.Row, 48).Value
    ElseIf InStr(Cells(ActiveCell.Row, 47).Value, "General Directorate") <> 0 Or InStr(Cells(ActiveCell.Row, 47).Value, "Regional Directorate") <> 0 Then
        If Cells(ActiveCell.Row, 48).Value <> "" Then
            Kolluk = WorksheetFunction.Proper(ThisWorkbook.Worksheets(2).Cells(6, 111).Value) & " " & Cells(ActiveCell.Row, 47).Value & " " & Cells(ActiveCell.Row, 48).Value
        Else
            Kolluk = WorksheetFunction.Proper(ThisWorkbook.Worksheets(2).Cells(6, 111).Value) & " " & Cells(ActiveCell.Row, 47).Value
        End If
    Else
        If Cells(ActiveCell.Row, 48).Value <> "" Then
            If InStr(Cells(ActiveCell.Row, 47).Value, "X.X. ") > 0 Then
                Kolluk = Mid(Cells(ActiveCell.Row, 47).Value, 6, Len(Cells(ActiveCell.Row, 47).Value)) & " " & Cells(ActiveCell.Row, 48).Value
            Else
                Kolluk = Cells(ActiveCell.Row, 47).Value & " " & Cells(ActiveCell.Row, 48).Value
            End If
        Else
            If InStr(Cells(ActiveCell.Row, 47).Value, "X.X. ") > 0 Then
                Kolluk = Mid(Cells(ActiveCell.Row, 47).Value, 6, Len(Cells(ActiveCell.Row, 47).Value))
            Else
                Kolluk = Cells(ActiveCell.Row, 47).Value
            End If
        End If
    End If
    
    'Recipient
    Set MyRange = objDoc.Tables(2).Cell(Row:=1, Column:=1).Range
    With MyRange.Find
        .Text = "<recipient>"
        .Replacement.Text = Kolluk
        .Execute Replace:=wdReplaceAll
    End With
    MyRange.HighlightColorIndex = wdNoHighlight
    'Tutanak tarihi
    Ek1 = Cells(ActiveCell.Row, 23).Value
    Set MyRange = objDoc.Tables(2).Cell(Row:=1, Column:=1).Range
    With MyRange.Find
        .Text = "<reportDate>"
        .Replacement.Text = Ek1
        .Execute Replace:=wdReplaceAll
    End With
    MyRange.HighlightColorIndex = wdNoHighlight

    
    'imzalar
    objDoc.Tables(3).Cell(Row:=4, Column:=1).Range.Text = Cells(ActiveCell.Row, 184).Value 'Ad Soyad1
    objDoc.Tables(3).Cell(Row:=5, Column:=1).Range.Text = Cells(ActiveCell.Row, 185).Value 'Unvan1
    objDoc.Tables(3).Cell(Row:=4, Column:=2).Range.Text = Cells(ActiveCell.Row, 187).Value 'Ad Soyad2
    objDoc.Tables(3).Cell(Row:=5, Column:=2).Range.Text = Cells(ActiveCell.Row, 188).Value 'Unvan2
    objDoc.Tables(3).Cell(Row:=4, Column:=3).Range.Text = Cells(ActiveCell.Row, 190).Value 'Ad Soyad3
    objDoc.Tables(3).Cell(Row:=5, Column:=3).Range.Text = Cells(ActiveCell.Row, 191).Value 'Unvan3
        
    'Tablo başlığı
    objDoc.Tables(4).Cell(Row:=1, Column:=1).Range.Text = Ek3 & " Evaluated as Invalid:"
    'Tabloya satır ekle
    x = 0
    If SonSira - IlkSira > 0 Then
        With objDoc.Tables(4)
            For i = IlkSira To SonSira - 1
                .Rows.Add BeforeRow:=objDoc.Tables(4).Rows(3)
                x = x + 1
            Next i
        End With
    End If
    For i = 3 To SonSira - IlkSira + 3
        objDoc.Tables(4).Cell(Row:=i, Column:=1).Range.Text = i - 2 'Tablo sıra no
        objDoc.Tables(4).Cell(Row:=i, Column:=2).Range.Text = Cells(IlkSira + i - 3, 52).Value 'Öğe türü
        objDoc.Tables(4).Cell(Row:=i, Column:=3).Range.Text = Cells(IlkSira + i - 3, 55).Value 'Öğe değeri
        objDoc.Tables(4).Cell(Row:=i, Column:=4).Range.Text = Cells(IlkSira + i - 3, 58).Value 'Adet
        'objDoc.Tables(4).Cell(Row:=i, Column:=5).Range.Text = Cells(IlkSira + i - 3, 61).Value 'Öğe ID No
    Next i
    'Öğe ID No kolonunu sil
    'objDoc.Tables(4).Columns(5).Delete
    
    'Tablo başlığı
    objDoc.Tables(5).Cell(Row:=1, Column:=1).Range.Text = "The Financial Unit Officer Who Delivered the " & Ek3 & ":"
    objDoc.Tables(5).Cell(Row:=2, Column:=3).Range.Text = Cells(ActiveCell.Row, 36).Value 'Ad Soyad
    objDoc.Tables(5).Cell(Row:=3, Column:=3).Range.Text = Cells(ActiveCell.Row, 38).Value 'TCK No
    objDoc.Tables(5).Cell(Row:=4, Column:=3).Range.Text = Cells(ActiveCell.Row, 39).Value 'Baba adı
    objDoc.Tables(5).Cell(Row:=5, Column:=3).Range.Text = Cells(ActiveCell.Row, 40).Value 'Doğum yeri
    objDoc.Tables(5).Cell(Row:=6, Column:=3).Range.Text = Cells(ActiveCell.Row, 41).Value 'Doğum tarihi
    objDoc.Tables(5).Cell(Row:=7, Column:=3).Range.Text = Cells(ActiveCell.Row, 44).Value 'Tel no


'    'Ek teslimat dekontu fotokopisi
    objDoc.Tables(6).Cell(Row:=1, Column:=1).Range.Text = "Attachment"
    objDoc.Tables(6).Cell(Row:=1, Column:=2).Range.Text = ":"

    x = Application.Sum(Range(Cells(IlkSira, 50), Cells(SonSira, 50))) 'Ek teslimat dekontu fotokopisi
    If x > 1 Then objDoc.Tables(6).Cell(Row:=1, Column:=3).Range.Text = "Delivery Receipt (" & x & " pages)"
    If x < 2 Then objDoc.Tables(6).Cell(Row:=1, Column:=3).Range.Text = "Delivery Receipt (" & x & " page)"
            
    objDoc.CheckBox1.Value = True
    objDoc.CheckBox1.Caption = "The invalid " & Ek4 & " received with the approval of the financial unit officer."
    'Türkçe karakterleri düzelt
    objDoc.CheckBox1.Enabled = False
    objDoc.CheckBox1.Enabled = True

    'Alt bilgi ekle
    objDoc.Sections(1).Footers(wdHeaderFooterPrimary).Range.Tables(1).Cell(Row:=1, Column:=1).Range.Text = ReNameRapor3_1
    
    'Sayfa sayısı kaydet komutuna bağlandı.
    objDoc.Save
    'Sayfayı text dosyasından çek
    TxtFileRapor3_1 = DestOpUserFolder & "Report 3 Page Count.txt"
    #If UseADO Then
        Const adReadLine = -2&
        With CreateObject("ADODB.Stream")
            .Open
            .LoadFromFile TxtFileRapor3_1
            Do Until .EOS
                TotalSayfaRapor3_1 = .ReadText(adReadLine)
            Loop
            .Close
        End With
    #Else 'UseFSO
        Const Rapor3_1TristateTrue = -1&
        With CreateObject("Scripting.FileSystemObject")
            With .OpenTextFile(TxtFileRapor3_1, Format:=Rapor3_1TristateTrue)
                Do Until .AtEndOfStream
                    TotalSayfaRapor3_1 = .ReadLine
                Loop
                .Close
            End With
        End With
    #End If
    'MsgBox "Tutanak1: " & TotalSayfaRapor3_1

    If TumDoc = True Then
        objWord.Activate
'        For i = 1 To SayPrt
'            objDoc.PrintOut
'        Next i
        objDoc.PrintOut Background:=False, Copies:=SayPrt
        'objWord.Documents.Save
        objDoc.Close SaveChanges:=False
        objWord.Visible = False
    End If

    'Tutanak1 ve Döküm sayfa sayısı
    Cells(ActiveCell.Row, 169).Value = TotalSayfaRapor3_1
    
    'Nesneleri temizle
    Set fso = Nothing
    Set objWord = Nothing
    Set objDoc = Nothing
    GoTo Son
    
Hata:
MsgBox "An error occurred while retrieving the page count of the statement. Please manually enter the statement page count in the appendix section when creating the cover letter.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
'If Err.Number <> 0 Then
'MsgBox "Error # " & Str(Err.Number)
'End If

Son:

'Nesneleri temizle
Set fso = Nothing
Set objWord = Nothing
Set objDoc = Nothing

'Worksheets(5).Activate
Application.ScreenUpdating = True
Application.DisplayAlerts = True
'Application.EnableEvents = True

'End If

End Sub

Sub Rapor3_2Tutanak2TipB()

Dim DestOperasyon As String, SourceRapor3_1Normal As String, AutoPath As String, ReNameRapor3_1 As String
Dim fso As Object, objWord As Object, objDoc As Object
Dim OpenKontrolName As String, OpenControl As String, DestOpUserFolder As String, DestOpUserFolderName As String
Dim ContSay As Integer, KontrolFile As String, x As Long, y As Long, i As Long, j As Long
Dim IlkSiraBul As Range, SonSiraBul As Range, IlkSira As Long, SonSira As Long
Dim GelenTema As String, SourceDL As String, ReNameDL As String
Dim TxtFileRapor3_1 As String, TotalSayfaRapor3_1 As String, TxtFileDokum As String, TotalSayfaDokum As String
Dim DokumFileTxt As Object, DokumSayfaGonder As String
Dim SourceRaporNormal As String, SourceTutanak2Normal As String, SourceUstYaziNormal As String
Dim Bolum1 As String, Bolum2 As String, Bolum3 As String, Bolum4 As String, Bolum5 As String, Bolum6 As String
Dim Ek1 As String, Ek2 As String, Ek3 As String, Ek4 As String, Ek5 As String, Ek6 As String, Birlestir As String

Dim ReNameRaporNormal As String, SiraNoIlkSatir As Long, RaporIlgi As String
Dim AltRaporNoIlk As Long, AltRaporNoSon As Long, AdetTopla As Long
Dim TekCogulTipA As String, k As Integer
Dim MyRange As Object, TxtFileRapor As String, TotalSayfaRapor As String
Dim ReNameTutanak2Normal As String, GidenTema As String
Dim TxtFileTutanak2 As String, TotalSayfaTutanak2 As String
Dim ReNameUstYaziNormal As String
Dim TxtFileUstYazi As String, TotalSayfaUstYazi As String
Dim Birimx As String, UserName As String, Kolluk As String

'TUTANAK2 için prosedürü başlat
'    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    'Application.EnableEvents = False
    'Worksheets(5).Unprotect Password:="123"

    ' Set up a control mechanism from left to right for document creation.
    If Cells(ActiveCell.Row, 6).Value = "x" Then
        MsgBox "Your process cannot proceed due to missing and/or incorrect data in the statement.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If

    
    'Pathfinder...
    AutoPath = ThisWorkbook.Path
    DestOperasyon = AutoPath & "\System Files\Operation\"
    'TUTANAK2 TANIMLARI
    SourceTutanak2Normal = AutoPath & "\System Files\System Templates\Statement 2 Template\Statement 2.docm"
    'Birden fazla kullanıcının operasyon klasöründe çakışmasını önlemek için operasyon klasörünün içinde yeni klasör aç.
    DestOpUserFolderName = "Operation-" & Left(ThisWorkbook.name, InStr(ThisWorkbook.name, ".") - 1) '& " " & Format(Now(), "ddmmyyyyhhmmss")
    DestOpUserFolder = DestOperasyon & DestOpUserFolderName & "\"

    ' System Files klasör adını kontrol et.
    If Not Dir(AutoPath & "\System Files\", vbDirectory) <> vbNullString Then
        MsgBox "Cannot access the folder " & AutoPath & "\System Files\" & ". The folder named 'System Files' might have been renamed or deleted.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    
    ' Operation klasörü adını kontrol et.
    If Not Dir(DestOperasyon, vbDirectory) <> vbNullString Then
        MsgBox "Cannot access the folder " & DestOperasyon & ". The folder named 'Operation' might have been renamed or deleted.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    
    ' Klasör isimlerini kontrol et.
    If Not Dir(SourceTutanak2Normal, vbDirectory) <> vbNullString Then
        MsgBox "Cannot access the folder " & SourceTutanak2Normal & ". The names of folders and/or files in this directory might have been changed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If


    'Operation klasörü içinde kullanıcı modülü klasörü yoksa oluştur.
    If Not Dir(DestOpUserFolder, vbDirectory) <> vbNullString Then
        MkDir DestOpUserFolder
    End If

    'On Error Resume Next
    ReNameTutanak2Normal = Cells(ActiveCell.Row, 5).Value & "-" & Cells(6, 8).Value
'________________________________________

    If TumDoc = False Then
        'Close the all Word application
        Call ModuleReport3.OpenWordControl

        'Operation klasöründeki docm uzantılı word dosyalarından açık olanları kapat ve temizle.
        OpenKontrolName = Dir(DestOpUserFolder & "*.docm")
        Do While OpenKontrolName <> ""
            OpenControl = IsFileOpen(DestOpUserFolder & OpenKontrolName)
            If OpenControl = True Then 'Açıksa
                On Error Resume Next
                Set objWord = GetObject(, "Word.Application")
                Set objWord = GetObject(, "Word.Application")
                Set objWord = GetObject(, "Word.Application")
                Set objWord = GetObject(, "Word.Application")
                Set objWord = GetObject(, "Word.Application")
                objWord.Quit SaveChanges:=True
                'MsgBox "Dosya OpenKontrol methodu ile kapatıldı."

            End If
            OpenKontrolName = Dir()
        Loop

    '    Call ModuleReport3.OpenWordControl

        Set objWord = Nothing
        Set objDoc = Nothing
    '________________________________________

        On Error Resume Next
    '    Klasörün içindeki tüm dosyaları sil (txt, docm vb.)
        ContSay = 0
        KontrolFile = Dir(DestOpUserFolder & "*.???")
        Do While KontrolFile <> ""
            ContSay = ContSay + 1
            KontrolFile = Dir()
        Loop
        If ContSay > 0 Then
            Kill DestOpUserFolder & "*.???"
        End If
    End If

    'Dosyayı şablondan operasyon klasörüne kopyala ve adını değiştir.
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.CopyFile (SourceTutanak2Normal), DestOpUserFolder & ReNameTutanak2Normal & ".docm", True
'________________________________________

    'Oluşturulacak dosyayı aç
    On Error Resume Next
    Set objWord = GetObject(, "Word.Application")
    Set objWord = GetObject(, "Word.Application")
    Set objWord = GetObject(, "Word.Application")
    Set objWord = GetObject(, "Word.Application")
    Set objWord = GetObject(, "Word.Application")
    If objWord Is Nothing Then
        'MsgBox "Dosya oluşturmada CreateObject methodu kullanılacak."
        Set objWord = CreateObject("Word.Application")
        objWord.Visible = False
    End If
    objWord.Documents.Open FileName:=DestOpUserFolder & ReNameTutanak2Normal & ".docm"
    objWord.Visible = True
    objWord.Activate 'Ekrana getirir.
    'objDoc.Activate 'Ekrana getirmez.
    objWord.Application.WindowState = wdWindowStateMaximize
    Set objDoc = GetObject(DestOpUserFolder & ReNameTutanak2Normal & ".docm")
'________________________________________

    'Birim
    Birimx = UCase(Replace(Replace(Worksheets(2).Cells(6, 99).Value, "i", "I"), "ı", "I")) & " UNIT"
    objDoc.Tables(1).Cell(Row:=1, Column:=1).Range.Text = Birimx
    'Tutanak2 tarihi
    objDoc.Tables(1).Cell(Row:=4, Column:=3).Range.Text = Cells(ActiveCell.Row, 69).Value
    'Belge tarihi ve numarası
    objDoc.Tables(1).Rows(6).Delete

    'Kolluk
    If Cells(ActiveCell.Row, 47).Value = "Provincial Directorate B" Then
        Kolluk = Cells(ActiveCell.Row, 19).Value & " Provincial Governorship Provincial Directorate B " & Cells(ActiveCell.Row, 48).Value
    ElseIf Cells(ActiveCell.Row, 47).Value = "District Directorate B" Then
        Kolluk = Cells(ActiveCell.Row, 20).Value & " District Governorship District Directorate B " & Cells(ActiveCell.Row, 48).Value
    ElseIf Cells(ActiveCell.Row, 47).Value = "Provincial Directorate C" Then
        Kolluk = Cells(ActiveCell.Row, 19).Value & " Provincial Governorship Provincial Directorate C " & Cells(ActiveCell.Row, 48).Value
    ElseIf Cells(ActiveCell.Row, 47).Value = "District Directorate C" Then
        Kolluk = Cells(ActiveCell.Row, 20).Value & " District Governorship District Directorate C " & Cells(ActiveCell.Row, 48).Value
    ElseIf Cells(ActiveCell.Row, 47).Value = "Provincial Directorate D" Then
        Kolluk = Cells(ActiveCell.Row, 19).Value & " Provincial Governorship Provincial Directorate D " & Cells(ActiveCell.Row, 48).Value
    ElseIf Cells(ActiveCell.Row, 47).Value = "District Directorate D" Then
        Kolluk = Cells(ActiveCell.Row, 20).Value & " District Governorship District Directorate D " & Cells(ActiveCell.Row, 48).Value
    ElseIf Cells(ActiveCell.Row, 47).Value = "Provincial Directorate E" Then
        Kolluk = Cells(ActiveCell.Row, 19).Value & " Provincial Governorship Provincial Directorate E " & Cells(ActiveCell.Row, 48).Value
    ElseIf Cells(ActiveCell.Row, 47).Value = "District Directorate E" Then
        Kolluk = Cells(ActiveCell.Row, 20).Value & " District Governorship District Directorate E " & Cells(ActiveCell.Row, 48).Value
    ElseIf InStr(Cells(ActiveCell.Row, 47).Value, "General Directorate") <> 0 Or InStr(Cells(ActiveCell.Row, 47).Value, "Regional Directorate") <> 0 Then
        If Cells(ActiveCell.Row, 48).Value <> "" Then
            Kolluk = WorksheetFunction.Proper(ThisWorkbook.Worksheets(2).Cells(6, 111).Value) & " " & Cells(ActiveCell.Row, 47).Value & " " & Cells(ActiveCell.Row, 48).Value
        Else
            Kolluk = WorksheetFunction.Proper(ThisWorkbook.Worksheets(2).Cells(6, 111).Value) & " " & Cells(ActiveCell.Row, 47).Value
        End If
    Else
        If Cells(ActiveCell.Row, 48).Value <> "" Then
            If InStr(Cells(ActiveCell.Row, 47).Value, "X.X. ") > 0 Then
                Kolluk = Mid(Cells(ActiveCell.Row, 47).Value, 6, Len(Cells(ActiveCell.Row, 47).Value)) & " " & Cells(ActiveCell.Row, 48).Value
            Else
                Kolluk = Cells(ActiveCell.Row, 47).Value & " " & Cells(ActiveCell.Row, 48).Value
            End If
        Else
            If InStr(Cells(ActiveCell.Row, 47).Value, "X.X. ") > 0 Then
                Kolluk = Mid(Cells(ActiveCell.Row, 47).Value, 6, Len(Cells(ActiveCell.Row, 47).Value))
            Else
                Kolluk = Cells(ActiveCell.Row, 47).Value
            End If
        End If
    End If
    
    objDoc.Tables(1).Cell(Row:=5, Column:=3).Range.Text = Kolluk
    
    objDoc.Tables(2).Cell(Row:=1, Column:=6).Range.Text = "Theme 2 No."

    'Tabloyu doldur
    Set IlkSiraBul = Range("FG7:FG100000").Find(What:=Cells(ActiveCell.Row, 5).Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    Set SonSiraBul = Range("FH7:FH100000").Find(What:=Cells(ActiveCell.Row, 5).Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    If Not IlkSiraBul Is Nothing Then
        IlkSira = IlkSiraBul.Row
    Else
        MsgBox "The serial number could not be found, so the operation cannot be completed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    If Not SonSiraBul Is Nothing Then
        SonSira = SonSiraBul.Row
    Else
        MsgBox "The serial number could not be found, so the operation cannot be completed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    'Tabloya satır ekle
    x = 0
    If SonSira - IlkSira > 0 Then
        With objDoc.Tables(2)
            For i = IlkSira To SonSira - 1
                .Rows.Add BeforeRow:=objDoc.Tables(2).Rows(2)
                x = x + 1
            Next i
        End With
    End If

    For i = 2 To SonSira - IlkSira + 2
        objDoc.Tables(2).Cell(Row:=i, Column:=1).Range.Text = i - 1 'Tablo sıra no
        objDoc.Tables(2).Cell(Row:=i, Column:=2).Range.Text = Cells(IlkSira + i - 2, 52).Value 'Öğe türü
        objDoc.Tables(2).Cell(Row:=i, Column:=3).Range.Text = Cells(IlkSira + i - 2, 55).Value 'Öğe değeri
        objDoc.Tables(2).Cell(Row:=i, Column:=4).Range.Text = Cells(IlkSira + i - 2, 58).Value 'Adet
        'objDoc.Tables(2).Cell(Row:=i, Column:=5).Range.Text = Cells(IlkSira + i - 2, 61).Value 'Öğe ID No
        objDoc.Tables(2).Cell(Row:=i, Column:=6).Range.Text = Cells(IlkSira, 26).Value  'Tema No (Temai her satıra yaz.)
        objDoc.Tables(2).Cell(Row:=i, Column:=7).Range.Text = Cells(IlkSira + i - 2, 64).Value 'Açıklama
    Next i

    'TipBler için tutanak2 tutanağında bulunan Öğe ID No kolounu sil.
    objDoc.Tables(2).Columns(5).Delete

    'Tutanak2 tutanağının metin kısmı
    'Tanımlamalar
    If Application.Sum(Range(Cells(IlkSira, 58), Cells(SonSira, 58))) > 1 Then
        Ek4 = "items"
    Else
        Ek4 = "item"
    End If
    Bolum1 = vbTab & "A total of "
    Bolum2 = " " & Ek4 & " , as described above, have been placed inside "
    Bolum3 = " "
    Bolum4 = " and enclosed to be sent to the designated unit mentioned above."

    Ek1 = Application.Sum(Range(Cells(IlkSira, 58), Cells(SonSira, 58)))
    Ek2 = Cells(ActiveCell.Row, 72).Value
    Ek3 = Application.WorksheetFunction.Proper(Cells(ActiveCell.Row, 71).Value)
    Ek3 = Left(Ek3, InStr(Ek3, "/") - 1)
    
    Birlestir = Bolum1 & Ek1 & Bolum2 & Ek2 & Bolum3 & Ek3 & Bolum4
    objDoc.Tables(3).Cell(Row:=1, Column:=1).Range.Text = Birlestir
    
    
    Set MyRange = objDoc.Tables(3).Cell(Row:=1, Column:=1).Range
    MyRange.Find.Execute FindText:=Ek1
    MyRange.Font.Bold = True

    Set MyRange = objDoc.Tables(3).Cell(Row:=1, Column:=1).Range
    With MyRange.Find
        .Execute FindText:=Ek2
        .Execute Forward:=True
    End With
    MyRange.Font.Bold = True
    'Aralıkta bulunan karakterleri bold yap
    Set MyRange = objDoc.Tables(3).Cell(Row:=1, Column:=1).Range
    MyRange.Find.Execute FindText:=Ek3
    MyRange.Font.Bold = True

    'imzalar
    objDoc.Tables(4).Cell(Row:=6, Column:=2).Range.Text = Cells(ActiveCell.Row, 193).Value 'Ad Soyad1
    objDoc.Tables(4).Cell(Row:=7, Column:=2).Range.Text = Cells(ActiveCell.Row, 194).Value 'Unvan1
    objDoc.Tables(4).Cell(Row:=6, Column:=3).Range.Text = Cells(ActiveCell.Row, 196).Value 'Ad Soyad2
    objDoc.Tables(4).Cell(Row:=7, Column:=3).Range.Text = Cells(ActiveCell.Row, 197).Value 'Unvan2
    
    'Alt bilgi ekle
    objDoc.Sections(1).Footers(wdHeaderFooterPrimary).Range.Tables(1).Cell(Row:=1, Column:=1).Range.Text = ReNameTutanak2Normal
    
'    'İmza boşluğunu sayfaya sığdırmak için düzenle
    If SonSira - IlkSira + 1 = 17 Then
        For i = 1 To 2
            objDoc.Tables(4).Rows(1).Delete
        Next i
    ElseIf SonSira - IlkSira + 1 = 18 Then
        For i = 1 To 4
            objDoc.Tables(4).Rows(1).Delete
        Next i
    ElseIf SonSira - IlkSira + 1 = 19 Then
        For i = 1 To 6
            objDoc.Tables(4).Rows(1).Delete
        Next i
    ElseIf SonSira - IlkSira + 1 = 20 Then
        For i = 1 To 7
            objDoc.Tables(4).Rows(1).Delete
        Next i
        objDoc.Tables(4).Rows(1).Height = 5
    End If

    'Sayfa sayısı kaydet komutuna bağlandı.
'    objDoc.Close SaveChanges:=True
'    objWord.Quit
    objDoc.Save
    'Sayfayı text dosyasından çek
    TxtFileTutanak2 = DestOpUserFolder & "Statement 2 Page Count.txt"
    #If UseADO Then
        Const adReadLine = -2&
        With CreateObject("ADODB.Stream")
            .Open
            .LoadFromFile TxtFileTutanak2
            Do Until .EOS
                TotalSayfaTutanak2 = .ReadText(adReadLine)
            Loop
            .Close
        End With
    #Else 'UseFSO
        Const Tutanak2TristateTrue = -1&
        With CreateObject("Scripting.FileSystemObject")
            With .OpenTextFile(TxtFileTutanak2, Format:=Tutanak2TristateTrue)
                Do Until .AtEndOfStream
                    TotalSayfaTutanak2 = .ReadLine
                Loop
                .Close
            End With
        End With
    #End If
    'MsgBox "Tutanak2: " & TotalSayfaTutanak2

    'Tutanak2 sayfa sayısı
    Cells(ActiveCell.Row, 171).Value = TotalSayfaTutanak2

    objWord.Visible = True
    objWord.Activate

    If TumDoc = True Then
        objWord.Activate
'        For i = 1 To SayPrt
'            objDoc.PrintOut
'        Next i
        objDoc.PrintOut Background:=False, Copies:=SayPrt
        'objWord.Documents.Save
        objDoc.Close SaveChanges:=False
        objWord.Visible = False
    End If

    'Nesneleri temizle
    Set fso = Nothing
    Set objWord = Nothing
    Set objDoc = Nothing

Son:

'Worksheets(5).Protect Password:="123"', DrawingObjects:=False

'Worksheets(5).Activate
Application.ScreenUpdating = True
Application.DisplayAlerts = True
'Application.EnableEvents = True

'End If

End Sub

Sub Rapor3_2FinansalBirimUstYaziTipB()

Dim DestOperasyon As String, SourceRapor3_1Normal As String, AutoPath As String, ReNameRapor3_1 As String
Dim fso As Object, objWord As Object, objDoc As Object
Dim OpenKontrolName As String, OpenControl As String, DestOpUserFolder As String, DestOpUserFolderName As String
Dim ContSay As Integer, KontrolFile As String, x As Long, y As Long, i As Long, j As Long
Dim IlkSiraBul As Range, SonSiraBul As Range, IlkSira As Long, SonSira As Long
Dim GelenTema As String, SourceDL As String, ReNameDL As String
Dim TxtFileRapor3_1 As String, TotalSayfaRapor3_1 As String, TxtFileDokum As String, TotalSayfaDokum As String
Dim DokumFileTxt As Object, DokumSayfaGonder As String
Dim SourceRaporNormal As String, SourceTutanak2Normal As String, SourceFinansalBirimUstYaziNormal As String
Dim Bolum1 As String, Bolum2 As String, Bolum3 As String, Bolum4 As String, Bolum5 As String, Bolum6 As String
Dim Ek1 As String, Ek2 As String, Ek3 As String, Ek4 As String, Ek5 As String, Ek6 As String, Birlestir As String

Dim ReNameRaporNormal As String, SiraNoIlkSatir As Long, RaporIlgi As String
Dim AltRaporNoIlk As Long, AltRaporNoSon As Long, AdetTopla As Long
Dim TekCogulTipA As String, k As Integer
Dim MyRange As Object, TxtFileRapor As String, TotalSayfaRapor As String
Dim ReNameTutanak2Normal As String, GidenTema As String
Dim TxtFileTutanak2 As String, TotalSayfaTutanak2 As String
Dim ReNameFinansalBirimUstYaziNormal As String
Dim TxtFileFinansalBirimUstYazi As String, TotalSayfaFinansalBirimUstYazi As String
Dim Birimx As String, UserName As String, Kolluk As String

Dim BStr As Integer
Dim M2 As Boolean, M3 As Boolean, M4 As Boolean
Dim Ifv As Boolean, Ify As Boolean
Dim XXXMudNotu As Boolean
Dim IlgiStrSay As Integer, Govde1StrSay As Integer, IlgiRng As Object, Govde1Rng As Object, Govde2Rng As Object
Dim IlgiFarkSay As Integer, Govde1FarkSay As Integer
Dim ustbilgimuhatap As String, ustbilgitarih As String, ustbilgisayi As String
Dim CokluSayfa As Integer
Dim ItemBul As Range, AdSoyadParaf As String, UnvanParaf As String, TelParaf As String


IlceSakla = ""
If InStr(Cells(ActiveCell.Row, 78).Value, " Organization A") <> 0 Then
    IlceSakla = Cells(ActiveCell.Row, 78).Value
    Cells(ActiveCell.Row, 78).Value = ""
End If

'ÜST YAZI için prosedürü başlat
'    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    'Application.EnableEvents = False
    'Worksheets(5).Unprotect Password:="123"

    ' Set up a left-to-right control mechanism for document creation.
    If Cells(ActiveCell.Row, 6).Value = "x" Then
        MsgBox "Your process cannot start due to missing and/or erroneous data in the Report 3 document.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    If Cells(ActiveCell.Row, 8).Value = "x" Then
        MsgBox "Your process cannot start due to missing and/or erroneous data in the Statement 2 document.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    
    'Pathfinder...
    AutoPath = ThisWorkbook.Path
    DestOperasyon = AutoPath & "\System Files\Operation\"
    'ÜST YAZI TANIMLARI
    SourceFinansalBirimUstYaziNormal = AutoPath & "\System Files\System Templates\Report 3 Cover Letter Templates\Report 3.2 – Type B Cover Letter – Financial Unit.docm"
    'Birden fazla kullanıcının operasyon klasöründe çakışmasını önlemek için operasyon klasörünün içinde yeni klasör aç.
    DestOpUserFolderName = "Operation-" & Left(ThisWorkbook.name, InStr(ThisWorkbook.name, ".") - 1) '& " " & Format(Now(), "ddmmyyyyhhmmss")
    DestOpUserFolder = DestOperasyon & DestOpUserFolderName & "\"

    'Check System Files folder.
    If Not Dir(AutoPath & "\System Files\", vbDirectory) <> vbNullString Then
        MsgBox "Cannot access the folder " & AutoPath & "\System Files\" & ". The folder named 'System Files' might have been renamed or deleted.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    
    'Check Operations folder.
    If Not Dir(DestOperasyon, vbDirectory) <> vbNullString Then
        MsgBox "Cannot access the folder " & DestOperasyon & ". The folder named 'Operation' might have been renamed or deleted.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    
    'Check folder names.
    If Not Dir(SourceFinansalBirimUstYaziNormal, vbDirectory) <> vbNullString Then
        MsgBox "Cannot access the folder " & SourceFinansalBirimUstYaziNormal & ". The names of folders and/or files in this directory might have been changed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    
    'Operation klasörü içinde kullanıcı modülü klasörü yoksa oluştur.
    If Not Dir(DestOpUserFolder, vbDirectory) <> vbNullString Then
        MkDir DestOpUserFolder
    End If

    'On Error Resume Next
    Set IlkSiraBul = Range("FG7:FG100000").Find(What:=Cells(ActiveCell.Row, 5).Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    Set SonSiraBul = Range("FH7:FH100000").Find(What:=Cells(ActiveCell.Row, 5).Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    If Not IlkSiraBul Is Nothing Then
        IlkSira = IlkSiraBul.Row
    Else
        MsgBox "The serial number could not be found, so the operation cannot be completed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    If Not SonSiraBul Is Nothing Then
        SonSira = SonSiraBul.Row
    Else
        MsgBox "The serial number could not be found, so the operation cannot be completed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If

    'Eklerde belirtilecek sayfalar için ön kontroller
    'Bu kısım çıkarıldı. Çünkü bu dokümanın ekinde önce aşamada oluştutulan raporlar yer almıyor.

    'PARAF HAZIRLIK İŞLEMİ
    UserName = Environ("UserProfile")
    UserName = UCase(Right(UserName, 7))
    Set ItemBul = Worksheets(2).Range("DR6:DR1000").Find(What:=UserName, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    If Not ItemBul Is Nothing Then
        AdSoyadParaf = "For information only: " & Worksheets(2).Cells(ItemBul.Row, 123).Value
        UnvanParaf = Worksheets(2).Cells(ItemBul.Row, 124).Value
        TelParaf = "Phone / Extension No: " & Worksheets(2).Cells(ItemBul.Row, 126).Value
    Else
        MsgBox UserName & " session is not registered in the system, so the cover letter cannot be created. Please register your session in the system using the Initials Interface located in the Settings Group of the Enterprise Document Automation System menu, and then try the operation again.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    
    'On Error Resume Next
    ReNameFinansalBirimUstYaziNormal = Cells(ActiveCell.Row, 5).Value & "-" & Cells(6, 9).Value & " Üst Yazı"
'________________________________________

    If TumDoc = False Then
        'Close the all Word application
        Call ModuleReport3.OpenWordControl

        'Operation klasöründeki docm uzantılı word dosyalarından açık olanları kapat ve temizle.
        OpenKontrolName = Dir(DestOpUserFolder & "*.docm")
        Do While OpenKontrolName <> ""
            OpenControl = IsFileOpen(DestOpUserFolder & OpenKontrolName)
            If OpenControl = True Then 'Açıksa
    '            On Error Resume Next
    '            Set objWord = GetObject(, "Word.Application")
    '            If objWord Is Nothing Then
    '                Set objWord = CreateObject("Word.Application")
    '                objWord.Visible = True
    '            End If
    '            objWord.Quit SaveChanges:=True

                On Error Resume Next
                Set objWord = GetObject(, "Word.Application")
                Set objWord = GetObject(, "Word.Application")
                Set objWord = GetObject(, "Word.Application")
                Set objWord = GetObject(, "Word.Application")
                Set objWord = GetObject(, "Word.Application")
                objWord.Quit SaveChanges:=True
                'MsgBox "Dosya OpenKontrol methodu ile kapatıldı."

            End If
            OpenKontrolName = Dir()
        Loop

    '    Call ModuleReport3.OpenWordControl

        Set objWord = Nothing
        Set objDoc = Nothing

    '________________________________________

        On Error Resume Next
    '    Klasörün içindeki tüm dosyaları sil (txt, docm vb.)
        ContSay = 0
        KontrolFile = Dir(DestOpUserFolder & "*.???")
        Do While KontrolFile <> ""
            ContSay = ContSay + 1
            KontrolFile = Dir()
        Loop
        If ContSay > 0 Then
            Kill DestOpUserFolder & "*.???"
        End If
    End If

    'Dosyayı şablondan operasyon klasörüne kopyala ve adını değiştir.
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.CopyFile (SourceFinansalBirimUstYaziNormal), DestOpUserFolder & ReNameFinansalBirimUstYaziNormal & ".docm", True
'________________________________________

    'Oluşturulacak dosyayı aç
    On Error Resume Next
    Set objWord = GetObject(, "Word.Application")
    Set objWord = GetObject(, "Word.Application")
    Set objWord = GetObject(, "Word.Application")
    Set objWord = GetObject(, "Word.Application")
    Set objWord = GetObject(, "Word.Application")
    If objWord Is Nothing Then
        'MsgBox "Dosya oluşturmada CreateObject methodu kullanılacak."
        Set objWord = CreateObject("Word.Application")
        objWord.Visible = False
    End If
    objWord.Documents.Open FileName:=DestOpUserFolder & ReNameFinansalBirimUstYaziNormal & ".docm"
    objWord.Visible = True
    objWord.Activate 'Ekrana getirir.
    'objDoc.Activate 'Ekrana getirmez.
    objWord.Application.WindowState = wdWindowStateMaximize
    Set objDoc = GetObject(DestOpUserFolder & ReNameFinansalBirimUstYaziNormal & ".docm")
'________________________________________

    'Birim
    Birimx = UCase(Replace(Replace(Worksheets(2).Cells(6, 99).Value, "i", "I"), "ı", "I")) & " UNIT"
    objDoc.Tables(1).Cell(Row:=1, Column:=1).Range.Text = Birimx

    Birimx = UCase(Replace(Replace(Worksheets(2).Cells(6, 99).Value, "i", "I"), "ı", "I"))
    'Üst yazı tarihi
    objDoc.Tables(1).Cell(Row:=1, Column:=2).Range.Text = Birimx & ", " & FormatEnglishDate(Cells(ActiveCell.Row, 75).Value) '(Format(Cells(ActiveCell.Row, 75).Value, "d mmmm yyyy"))
    'Yazı no
    objDoc.Tables(1).Cell(Row:=4, Column:=2).Range.Text = Cells(ActiveCell.Row, 76).Value

    'Gönderi tipi
    Ek2 = Cells(ActiveCell.Row, 85).Value
    Ek2 = UCase(Replace(Replace(Ek2, "i", "I"), "ı", "I"))
    objDoc.Tables(1).Cell(Row:=7, Column:=2).Range.Text = Ek2
    
    
    'Kolluk
    If Cells(ActiveCell.Row, 47).Value = "Provincial Directorate B" Then
        Kolluk = Cells(ActiveCell.Row, 19).Value & " Provincial Governorship Provincial Directorate B " & Cells(ActiveCell.Row, 48).Value
    ElseIf Cells(ActiveCell.Row, 47).Value = "District Directorate B" Then
        Kolluk = Cells(ActiveCell.Row, 20).Value & " District Governorship District Directorate B " & Cells(ActiveCell.Row, 48).Value
    ElseIf Cells(ActiveCell.Row, 47).Value = "Provincial Directorate C" Then
        Kolluk = Cells(ActiveCell.Row, 19).Value & " Provincial Governorship Provincial Directorate C " & Cells(ActiveCell.Row, 48).Value
    ElseIf Cells(ActiveCell.Row, 47).Value = "District Directorate C" Then
        Kolluk = Cells(ActiveCell.Row, 20).Value & " District Governorship District Directorate C " & Cells(ActiveCell.Row, 48).Value
    ElseIf Cells(ActiveCell.Row, 47).Value = "Provincial Directorate D" Then
        Kolluk = Cells(ActiveCell.Row, 19).Value & " Provincial Governorship Provincial Directorate D " & Cells(ActiveCell.Row, 48).Value
    ElseIf Cells(ActiveCell.Row, 47).Value = "District Directorate D" Then
        Kolluk = Cells(ActiveCell.Row, 20).Value & " District Governorship District Directorate D " & Cells(ActiveCell.Row, 48).Value
    ElseIf Cells(ActiveCell.Row, 47).Value = "Provincial Directorate E" Then
        Kolluk = Cells(ActiveCell.Row, 19).Value & " Provincial Governorship Provincial Directorate E " & Cells(ActiveCell.Row, 48).Value
    ElseIf Cells(ActiveCell.Row, 47).Value = "District Directorate E" Then
        Kolluk = Cells(ActiveCell.Row, 20).Value & " District Governorship District Directorate E " & Cells(ActiveCell.Row, 48).Value
    ElseIf InStr(Cells(ActiveCell.Row, 47).Value, "General Directorate") <> 0 Or InStr(Cells(ActiveCell.Row, 47).Value, "Regional Directorate") <> 0 Then
        If Cells(ActiveCell.Row, 48).Value <> "" Then
            Kolluk = WorksheetFunction.Proper(ThisWorkbook.Worksheets(2).Cells(6, 111).Value) & " " & Cells(ActiveCell.Row, 47).Value & " " & Cells(ActiveCell.Row, 48).Value
        Else
            Kolluk = WorksheetFunction.Proper(ThisWorkbook.Worksheets(2).Cells(6, 111).Value) & " " & Cells(ActiveCell.Row, 47).Value
        End If
    Else
        If InStr(Cells(ActiveCell.Row, 47).Value, "X.X. ") > 0 Then  'Başında X.X. ifadesi varsa
            If Cells(ActiveCell.Row, 48).Value <> "" Then
                Kolluk = Mid(Cells(ActiveCell.Row, 47).Value, 6, Len(Cells(ActiveCell.Row, 47).Value)) & " " & Cells(ActiveCell.Row, 48).Value
            Else
                Kolluk = Mid(Cells(ActiveCell.Row, 47).Value, 6, Len(Cells(ActiveCell.Row, 47).Value))
            End If
        Else
            If Cells(ActiveCell.Row, 48).Value <> "" Then
                Kolluk = Cells(ActiveCell.Row, 47).Value & " " & Cells(ActiveCell.Row, 48).Value
            Else
                Kolluk = Cells(ActiveCell.Row, 47).Value
            End If
        End If
    End If

    'YENİ MUHATAP TEMASI
    M2 = False
    M3 = False
    M4 = False
    BStr = 9
    '1. sayfadan sonraki üst bilgi tanımları
    ustbilgitarih = (Format(Cells(ActiveCell.Row, 75).Value, "dd.mm.yyyy"))
    ustbilgisayi = Cells(ActiveCell.Row, 76).Value
    
    'Muhatap
    If Cells(ActiveCell.Row, 82).Value <> "" Then
        objDoc.Tables(1).Cell(Row:=BStr, Column:=1).Range.Text = UCase(Replace(Replace(Cells(ActiveCell.Row, 30).Value, "i", "I"), "ı", "I")) 'FinansalBirim
        objDoc.Tables(1).Cell(Row:=BStr + 1, Column:=1).Range.Text = "(" & Cells(ActiveCell.Row, 82).Value & ")" 'Birim
        objDoc.Tables(1).Cell(Row:=BStr + 2, Column:=1).Range.Text = Cells(ActiveCell.Row, 79).Value 'Adres
        If Cells(ActiveCell.Row, 78).Value <> "" Then
            objDoc.Tables(1).Cell(Row:=BStr + 3, Column:=1).Range.Text = Cells(ActiveCell.Row, 78).Value & "/" & UCase(Replace(Replace(Cells(ActiveCell.Row, 77).Value, "i", "I"), "ı", "I"))
        Else
            objDoc.Tables(1).Cell(Row:=BStr + 3, Column:=1).Range.Text = UCase(Replace(Replace(Cells(ActiveCell.Row, 77).Value, "i", "I"), "ı", "I"))
        End If
        objDoc.Tables(1).Cell(Row:=BStr + 3, Column:=1).Range.Font.Underline = wdUnderlineSingle
        M4 = True
        
        ustbilgimuhatap = Cells(ActiveCell.Row, 30).Value
    Else
        objDoc.Tables(1).Cell(Row:=BStr, Column:=1).Range.Text = UCase(Replace(Replace(Cells(ActiveCell.Row, 30).Value, "i", "I"), "ı", "I")) 'FinansalBirim
        objDoc.Tables(1).Cell(Row:=BStr + 1, Column:=1).Range.Text = Cells(ActiveCell.Row, 79).Value 'Adres
        If Cells(ActiveCell.Row, 78).Value <> "" Then
            objDoc.Tables(1).Cell(Row:=BStr + 2, Column:=1).Range.Text = Cells(ActiveCell.Row, 78).Value & "/" & UCase(Replace(Replace(Cells(ActiveCell.Row, 77).Value, "i", "I"), "ı", "I"))
        Else
            objDoc.Tables(1).Cell(Row:=BStr + 2, Column:=1).Range.Text = UCase(Replace(Replace(Cells(ActiveCell.Row, 77).Value, "i", "I"), "ı", "I"))
        End If
        objDoc.Tables(1).Cell(Row:=BStr + 2, Column:=1).Range.Font.Underline = wdUnderlineSingle
        M3 = True
        
        ustbilgimuhatap = Cells(ActiveCell.Row, 30).Value
    End If
    
    
    'Üst yazı gövde metni ve ek senaryoları
    'tipAnın/tipAların (TipA adedi)
    AdetTopla = Application.Sum(Range(Cells(IlkSira, 58), Cells(SonSira, 58)))
    If AdetTopla = 1 Then
        Ek1 = "Type B item"
        Ek4 = "item"
        Ek6 = "has"
    ElseIf AdetTopla > 1 Then
        Ek1 = "Type B items"
        Ek4 = "items"
        Ek6 = "have"
    End If
    
    Ek2 = Cells(ActiveCell.Row, 35).Value 'Teslimat tarihi
    
    If Cells(ActiveCell.Row, 12).Value = "Point2" Then
        Ek3 = "Point 2"
    ElseIf Cells(ActiveCell.Row, 12).Value = "Point3" Then
        Ek3 = "Point 3"
    End If
    
    Bolum1 = "The " & Ek1 & ", delivered to our unit on " & Ek2 & " via xxxxx xxxxx xxxxx for the purpose of xxxxx xxxxx xxxxx xxxxx and processed at " & Ek3 & ", " & _
            Ek6 & " been evaluated as invalid, as documented in the attached statement."
   
    Bolum2 = "The " & Ek4 & " " & Ek6 & " been retained by our unit for forwarding to the X1 Process Monitoring Directorate."
    Bolum3 = Chr(9) & "Submitted for your information."
    Birlestir = Chr(9) & Bolum1 & vbNewLine & Chr(9) & Bolum2
    objDoc.Tables(2).Cell(Row:=1, Column:=1).Range.Text = Birlestir
    objDoc.Tables(2).Cell(Row:=3, Column:=1).Range.Text = Bolum3

    'imzalar
    objDoc.Tables(3).Cell(Row:=4, Column:=2).Range.Text = Cells(ActiveCell.Row, 199).Value 'Ad Soyad1
    objDoc.Tables(3).Cell(Row:=5, Column:=2).Range.Text = Cells(ActiveCell.Row, 200).Value 'Unvan1
    objDoc.Tables(3).Cell(Row:=4, Column:=3).Range.Text = Cells(ActiveCell.Row, 202).Value 'Ad Soyad2
    objDoc.Tables(3).Cell(Row:=5, Column:=3).Range.Text = Cells(ActiveCell.Row, 203).Value 'Unvan2
    

    'Ekler
    objDoc.Tables(3).Cell(Row:=8, Column:=1).Range.Text = "Attachment:"
    If Cells(ActiveCell.Row, 80).Value <> "" Then
        'Dekont
        If Cells(ActiveCell.Row, 80).Value > 1 Then objDoc.Tables(3).Cell(Row:=8, Column:=2).Range.Text = " 1) Xxxxx Receipt (" & Cells(ActiveCell.Row, 80).Value & " pages)"
        If Cells(ActiveCell.Row, 80).Value < 2 Then objDoc.Tables(3).Cell(Row:=8, Column:=2).Range.Text = " 1) Xxxxx Receipt (" & Cells(ActiveCell.Row, 80).Value & " page)"
        'Statement
        x = Application.Sum(Range(Cells(IlkSira, 169), Cells(SonSira, 169))) + Cells(ActiveCell.Row, 50).Value
        objDoc.Tables(3).Cell(Row:=9, Column:=2).Range.Text = " 2) Attached Statement (" & "total of " & x & " pages)"
    Else
        'Statement
        x = Application.Sum(Range(Cells(IlkSira, 169), Cells(SonSira, 169))) + Cells(ActiveCell.Row, 50).Value
        objDoc.Tables(3).Cell(Row:=8, Column:=2).Range.Text = " Attached Statement (" & "total of " & x & " pages)"
    End If

'    'Paraf
'    UserName = Environ("UserProfile")
'    objDoc.Tables(5).Cell(Row:=13, Column:=1).Range.Text = UCase(Replace(Replace(Mid(Right(UserName, 7), 4, 2), "i", "I"), "ı", "I"))

    '__________________________DİNAMİK SAYFA YAPISI
    
    'İlgi ve gövde metninin kaç satırdan oluştuğu bilgisinin elde edilmesinde kullanılıyor.
    'Set IlgiRng = objDoc.Tables(1).Cell(Row:=15, Column:=3).Range
    Set Govde1Rng = objDoc.Tables(2).Cell(Row:=1, Column:=1).Range
    Set Govde2Rng = objDoc.Tables(2).Cell(Row:=3, Column:=1).Range
    
    'IlgiStrSay = Govde1Rng.Information(wdFirstCharacterLineNumber) - IlgiRng.Information(wdFirstCharacterLineNumber) - 1
    Govde1StrSay = Govde2Rng.Information(wdFirstCharacterLineNumber) - Govde1Rng.Information(wdFirstCharacterLineNumber)
    IlgiFarkSay = 0
    Govde1FarkSay = Govde1StrSay - 5 'Varsayılan 2 paragrafta (1 rowda) toplam 4 satıra göre sıfırlandı.
    CokluSayfa = 0
    'MsgBox Govde1FarkSay
    
    '__________________________________'Tes.Düzensiz.Dek. var.
   
    'Dinamik sayfa düzeni

    'Dinamik sayfa düzeni
    
    objDoc.Tables(2).Rows(2).Delete 'Boş satırı sil
    For i = 1 To 3
        objDoc.Tables(1).Rows(13).Delete 'Muhatap sonrası
    Next i
    
    If Cells(ActiveCell.Row, 80).Value <> "" Then 'Tes.Düzensiz.Dek. var. '2 ek var
        For i = 1 To 3 'Ek sonrası satırları sil
            objDoc.Tables(3).Rows(10).Delete
        Next i
        Govde1FarkSay = Govde1FarkSay - 1
    Else 'Tes.Düzensiz.Dek. yok. '1 ek var
        For i = 1 To 4 'Ek sonrası satırları sil
            objDoc.Tables(3).Rows(9).Delete
        Next i
        Govde1FarkSay = Govde1FarkSay - 2
    End If

    If M4 = True Then
        If IlgiFarkSay + Govde1FarkSay < 17 Then
            Govde1FarkSay = Govde1FarkSay + 2
            'MsgBox IlgiFarkSay + Govde1FarkSay
            If IlgiFarkSay + Govde1FarkSay > 8 Then
                For i = 9 To IlgiFarkSay + Govde1FarkSay
                    If i = 9 Then
                        objDoc.Tables(3).Rows(6).Delete 'Ek öncesi
                    ElseIf i = 10 Then
                        objDoc.Tables(3).Rows(6).Delete 'Ek öncesi
                    ElseIf i = 11 Then
                        objDoc.Tables(3).Rows(1).Delete 'İmza öncesi
                    ElseIf i = 12 Then
                        objDoc.Tables(3).Rows(1).Delete 'İmza öncesi
                    ElseIf i = 13 Then
                        objDoc.Tables(1).Rows(8).Delete 'Muhatap öncesi
                    ElseIf i = 14 Then
                        objDoc.Tables(1).Rows(8).Delete 'Muhatap öncesi
                    End If
                Next i
            Else
                If (IlgiFarkSay + Govde1FarkSay) < 8 Then
                    For i = 1 To 8 - (IlgiFarkSay + Govde1FarkSay)
                        'Boş satır ekle
                        With objDoc.Tables(3)
                            .Rows.Add BeforeRow:=objDoc.Tables(3).Rows(8)
                        End With
                    Next i
                End If
            End If
        Else
            CokluSayfa = 1
        End If
    ElseIf M3 = True Then
        For i = 1 To 1
            objDoc.Tables(1).Rows(13).Delete 'ilgi öncesi
        Next i

        If IlgiFarkSay + Govde1FarkSay < 16 Then
            Govde1FarkSay = Govde1FarkSay + 1
            'MsgBox IlgiFarkSay + Govde1FarkSay
            If IlgiFarkSay + Govde1FarkSay > 8 Then
                For i = 9 To IlgiFarkSay + Govde1FarkSay
                    If i = 9 Then
                        objDoc.Tables(3).Rows(6).Delete 'Ek öncesi
                    ElseIf i = 10 Then
                        objDoc.Tables(3).Rows(6).Delete 'Ek öncesi
                    ElseIf i = 11 Then
                        objDoc.Tables(3).Rows(1).Delete 'İmza öncesi
                    ElseIf i = 12 Then
                        objDoc.Tables(3).Rows(1).Delete 'İmza öncesi
                    ElseIf i = 13 Then
                        objDoc.Tables(1).Rows(8).Delete 'Muhatap öncesi
                    ElseIf i = 14 Then
                        objDoc.Tables(1).Rows(8).Delete 'Muhatap öncesi
                    End If
                Next i
            Else
                If (IlgiFarkSay + Govde1FarkSay) < 8 Then
                    For i = 1 To 8 - (IlgiFarkSay + Govde1FarkSay)
                        'Boş satır ekle
                        With objDoc.Tables(3)
                            .Rows.Add BeforeRow:=objDoc.Tables(3).Rows(8)
                        End With
                    Next i
                End If
            End If
        Else
            CokluSayfa = 1
        End If
    ElseIf M2 = True Then
        For i = 1 To 2
            objDoc.Tables(1).Rows(13).Delete 'ilgi öncesi
        Next i
        
        'Govde1FarkSay = Govde1FarkSay + 0
        'MsgBox IlgiFarkSay + Govde1FarkSay
        If IlgiFarkSay + Govde1FarkSay < 15 Then
            If IlgiFarkSay + Govde1FarkSay > 8 Then
                For i = 9 To IlgiFarkSay + Govde1FarkSay
                    If i = 9 Then
                        objDoc.Tables(3).Rows(6).Delete 'Ek öncesi
                    ElseIf i = 10 Then
                        objDoc.Tables(3).Rows(6).Delete 'Ek öncesi
                    ElseIf i = 11 Then
                        objDoc.Tables(3).Rows(1).Delete 'İmza öncesi
                    ElseIf i = 12 Then
                        objDoc.Tables(3).Rows(1).Delete 'İmza öncesi
                    ElseIf i = 13 Then
                        objDoc.Tables(1).Rows(8).Delete 'Muhatap öncesi
                    ElseIf i = 14 Then
                        objDoc.Tables(1).Rows(8).Delete 'Muhatap öncesi
                    End If
                Next i
            Else
                If (IlgiFarkSay + Govde1FarkSay) < 8 Then
                    For i = 1 To 8 - (IlgiFarkSay + Govde1FarkSay)
                        'Boş satır ekle
                        With objDoc.Tables(3)
                            .Rows.Add BeforeRow:=objDoc.Tables(3).Rows(8)
                        End With
                    Next i
                End If
            End If
        Else
            CokluSayfa = 1
        End If
    End If
    
    objDoc.Paragraphs(objDoc.Paragraphs.Count).Range.Font.Size = 1


    'footer
    If objDoc.ActiveWindow.View.SplitSpecial <> wdPaneNone Then
        objDoc.ActiveWindow.Panes(2).Close
    End If
    If objDoc.ActiveWindow.ActivePane.View.Type = wdNormalView Or objDoc.ActiveWindow.ActivePane.View.Type = wdOutlineView Then
        objDoc.ActiveWindow.ActivePane.View.Type = wdPrintView
    End If

    With objDoc
        With .Sections(1).Headers(wdHeaderFooterPrimary).Range
            .InsertAfter Text:="Page "
            .Fields.Add Range:=.Characters.Last, Type:=wdFieldEmpty, Text:="PAGE", PreserveFormatting:=False
            .InsertAfter Text:=" of the letter dated " & ustbilgitarih & _
                             ", reference number " & ustbilgisayi & _
                             ", addressed to the " & ustbilgimuhatap
        End With
    End With
    
    objDoc.ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument

    'PARAF EKLEME İŞLEMİ
    objDoc.Sections(1).Footers(wdHeaderFooterFirstPage).Range.Tables(1).Cell(Row:=1, Column:=2).Range.Text = AdSoyadParaf & vbNewLine & _
                                                                                       UnvanParaf & vbNewLine & _
                                                                                       TelParaf
    objDoc.Sections(1).Footers(wdHeaderFooterPrimary).Range.Tables(1).Cell(Row:=1, Column:=2).Range.Text = AdSoyadParaf & vbNewLine & _
                                                                                       UnvanParaf & vbNewLine & _
                                                                                       TelParaf

Tekrarla:
    'Sayfa sayısı kaydet komutuna bağlandı.
    objDoc.Save
    'Sayfayı text dosyasından çek
    TxtFileFinansalBirimUstYazi = DestOpUserFolder & "Cover Letter Page Count.txt"
    #If UseADO Then
        Const adReadLine = -2&
        With CreateObject("ADODB.Stream")
            .Open
            .LoadFromFile TxtFileFinansalBirimUstYazi
            Do Until .EOS
                TotalSayfaFinansalBirimUstYazi = .ReadText(adReadLine)
            Loop
            .Close
        End With
    #Else 'UseFSO
        Const FinansalBirimUstYaziTristateTrue = -1&
        With CreateObject("Scripting.FileSystemObject")
            With .OpenTextFile(TxtFileFinansalBirimUstYazi, Format:=FinansalBirimUstYaziTristateTrue)
                Do Until .AtEndOfStream
                    TotalSayfaFinansalBirimUstYazi = .ReadLine
                Loop
                .Close
            End With
        End With
    #End If
    'MsgBox "Üst Yazı: " & TotalSayfaFinansalBirimUstYazi

    'Üst Yazı sayfa sayısı
    Cells(ActiveCell.Row, 172).Value = TotalSayfaFinansalBirimUstYazi

    objWord.Visible = True
    objWord.Activate

    If TumDoc = True Then
        objWord.Activate
'        For i = 1 To SayPrt
'            objDoc.PrintOut
'        Next i
        objDoc.PrintOut Background:=False, Copies:=SayPrt
        'objWord.Documents.Save
        objDoc.Close SaveChanges:=False
        objWord.Visible = False
    End If

    'Nesneleri temizle
    Set fso = Nothing
    Set objWord = Nothing
    Set objDoc = Nothing

Son:

If IlceSakla <> "" Then
    Cells(ActiveCell.Row, 78).Value = IlceSakla
End If


'Worksheets(5).Protect Password:="123"', DrawingObjects:=False
'Worksheets(5).Activate
Application.ScreenUpdating = True
Application.DisplayAlerts = True
'Application.EnableEvents = True


End Sub

Sub Rapor3_2UstYaziTipB()

Dim DestOperasyon As String, SourceRapor3_1Normal As String, AutoPath As String, ReNameRapor3_1 As String
Dim fso As Object, objWord As Object, objDoc As Object
Dim OpenKontrolName As String, OpenControl As String, DestOpUserFolder As String, DestOpUserFolderName As String
Dim ContSay As Integer, KontrolFile As String, x As Long, y As Long, i As Long, j As Long
Dim IlkSiraBul As Range, SonSiraBul As Range, IlkSira As Long, SonSira As Long
Dim GelenTema As String, SourceDL As String, ReNameDL As String
Dim TxtFileRapor3_1 As String, TotalSayfaRapor3_1 As String, TxtFileDokum As String, TotalSayfaDokum As String
Dim DokumFileTxt As Object, DokumSayfaGonder As String
Dim SourceRaporNormal As String, SourceTutanak2Normal As String, SourceUstYaziNormal As String
Dim Bolum1 As String, Bolum2 As String, Bolum3 As String, Bolum4 As String, Bolum5 As String, Bolum6 As String, Bolum7 As String
Dim Ek1 As String, Ek2 As String, Ek3 As String, Ek4 As String, Ek5 As String, Ek6 As String, Birlestir As String

Dim ReNameRaporNormal As String, SiraNoIlkSatir As Long, RaporIlgi As String
Dim AltRaporNoIlk As Long, AltRaporNoSon As Long, AdetTopla As Long
Dim TekCogulTipA As String, k As Integer
Dim MyRange As Object, TxtFileRapor As String, TotalSayfaRapor As String
Dim ReNameTutanak2Normal As String, GidenTema As String
Dim TxtFileTutanak2 As String, TotalSayfaTutanak2 As String
Dim ReNameUstYaziNormal As String
Dim TxtFileUstYazi As String, TotalSayfaUstYazi As String
Dim Birimx As String, UserName As String
Dim IlBuyukHarf, IlKucukHarf, IlceBuyukHarf, IlceKucukHarf As String
Dim GonderimUsulu As String

Dim BStr As Integer
Dim M2 As Boolean, M3 As Boolean, M4 As Boolean
Dim Ifv As Boolean, Ify As Boolean
Dim XXXMudNotu As Boolean
Dim IlgiStrSay As Integer, Govde1StrSay As Integer, IlgiRng As Object, Govde1Rng As Object, Govde2Rng As Object
Dim IlgiFarkSay As Integer, Govde1FarkSay As Integer
Dim ustbilgimuhatap As String, ustbilgitarih As String, ustbilgisayi As String
Dim CokluSayfa As Integer
Dim ItemBul As Range, AdSoyadParaf As String, UnvanParaf As String, TelParaf As String


IlceSakla = ""
If InStr(Cells(ActiveCell.Row, 20).Value, " Organization A") <> 0 Then
    IlceSakla = Cells(ActiveCell.Row, 20).Value
    Cells(ActiveCell.Row, 20).Value = ""
End If

'ÜST YAZI için prosedürü başlat
'    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    'Application.EnableEvents = False
    'Worksheets(5).Unprotect Password:="123"


    'Set up a left-to-right control mechanism for document creation.
    If Cells(ActiveCell.Row, 6).Value = "x" Then
        MsgBox "Your process cannot be started due to missing and/or erroneous data in Report 3 document.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    If Cells(ActiveCell.Row, 8).Value = "x" Then
        MsgBox "Your process cannot be started due to missing and/or erroneous data in Statement 2 document.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    If Cells(ActiveCell.Row, 9).Value = "x" Then
        MsgBox "Your process cannot be started due to missing and/or erroneous data in Financial Unit cover letter data.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If

    'Pathfinder...
    AutoPath = ThisWorkbook.Path
    DestOperasyon = AutoPath & "\System Files\Operation\"
    'ÜST YAZI TANIMLARI
    SourceUstYaziNormal = AutoPath & "\System Files\System Templates\Report 3 Cover Letter Templates\Report 3.2 – Type B Cover Letter.docm"
    'Birden fazla kullanıcının operasyon klasöründe çakışmasını önlemek için operasyon klasörünün içinde yeni klasör aç.
    DestOpUserFolderName = "Operation-" & Left(ThisWorkbook.name, InStr(ThisWorkbook.name, ".") - 1) '& " " & Format(Now(), "ddmmyyyyhhmmss")
    DestOpUserFolder = DestOperasyon & DestOpUserFolderName & "\"

    'System Files folder name check.
    If Not Dir(AutoPath & "\System Files\", vbDirectory) <> vbNullString Then
        MsgBox "Cannot access the folder " & AutoPath & "\System Files\" & ". The folder named 'System Files' might have been renamed or deleted.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    'Operation folder name check.
    If Not Dir(DestOperasyon, vbDirectory) <> vbNullString Then
        MsgBox "Cannot access the folder " & DestOperasyon & ". The folder named 'Operation' might have been renamed or deleted.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    
    'Check folder names.
    If Not Dir(SourceUstYaziNormal, vbDirectory) <> vbNullString Then
        MsgBox "Cannot access the folder " & SourceUstYaziNormal & ". The names of folders and/or files in this directory might have been changed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If

    'Operation klasörü içinde kullanıcı modülü klasörü yoksa oluştur.
    If Not Dir(DestOpUserFolder, vbDirectory) <> vbNullString Then
        MkDir DestOpUserFolder
    End If

    'On Error Resume Next
    Set IlkSiraBul = Range("FG7:FG100000").Find(What:=Cells(ActiveCell.Row, 5).Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    Set SonSiraBul = Range("FH7:FH100000").Find(What:=Cells(ActiveCell.Row, 5).Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    If Not IlkSiraBul Is Nothing Then
        IlkSira = IlkSiraBul.Row
    Else
        MsgBox "The serial number could not be found, so the operation cannot be completed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    If Not SonSiraBul Is Nothing Then
        SonSira = SonSiraBul.Row
    Else
        MsgBox "The serial number could not be found, so the operation cannot be completed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If

    'Pre-checks for pages to be included in attachments
    'Tutanak1 check
    If Cells(IlkSira, 169).Value = "" Then
        MsgBox Cells(IlkSira, 5).Value & " row number: The statement has not been created, so the cover letter for row " & Cells(IlkSira, 5).Value & " cannot be prepared. Please create the Report 3 Statement for row " & Cells(IlkSira, 5).Value & " and try again.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    
    'Tutanak2 check
    If Cells(IlkSira, 171).Value = "" Then
        MsgBox Cells(IlkSira, 5).Value & " row number: The Statement 2 has not been created, so the cover letter for row " & Cells(IlkSira, 5).Value & " cannot be prepared. Please create the Statement 2 for row " & Cells(IlkSira, 5).Value & " and try again.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    
    'Tutanak2 check (Financial Unit)
    If Cells(IlkSira, 172).Value = "" Then
        MsgBox Cells(IlkSira, 5).Value & " row number: The Financial Unit Cover Letter has not been created, so the cover letter for row " & Cells(IlkSira, 5).Value & " cannot be prepared. Please create the Financial Unit Cover Letter for row " & Cells(IlkSira, 5).Value & " and try again.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If

 
    'PARAF HAZIRLIK İŞLEMİ
    UserName = Environ("UserProfile")
    UserName = UCase(Right(UserName, 7))
    Set ItemBul = Worksheets(2).Range("DR6:DR1000").Find(What:=UserName, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    If Not ItemBul Is Nothing Then
        AdSoyadParaf = "For information only: " & Worksheets(2).Cells(ItemBul.Row, 123).Value
        UnvanParaf = Worksheets(2).Cells(ItemBul.Row, 124).Value
        TelParaf = "Phone / Extension No: " & Worksheets(2).Cells(ItemBul.Row, 126).Value
    Else
        MsgBox UserName & " session is not registered in the system, so the cover letter cannot be created. Please register your session via the Initials Interface located in the Settings Group of the Enterprise Document Automation System menu and try the operation again.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    
    'On Error Resume Next
    ReNameUstYaziNormal = Cells(ActiveCell.Row, 5).Value & "-" & Cells(6, 10).Value
'________________________________________

    If TumDoc = False Then
        'Close the all Word application
        Call ModuleReport3.OpenWordControl

        'Operation klasöründeki docm uzantılı word dosyalarından açık olanları kapat ve temizle.
        OpenKontrolName = Dir(DestOpUserFolder & "*.docm")
        Do While OpenKontrolName <> ""
            OpenControl = IsFileOpen(DestOpUserFolder & OpenKontrolName)
            If OpenControl = True Then 'Açıksa
                On Error Resume Next
                Set objWord = GetObject(, "Word.Application")
                Set objWord = GetObject(, "Word.Application")
                Set objWord = GetObject(, "Word.Application")
                Set objWord = GetObject(, "Word.Application")
                Set objWord = GetObject(, "Word.Application")
                objWord.Quit SaveChanges:=True
                'MsgBox "Dosya OpenKontrol methodu ile kapatıldı."

            End If
            OpenKontrolName = Dir()
        Loop

    '    Call ModuleReport3.OpenWordControl

        Set objWord = Nothing
        Set objDoc = Nothing

    '________________________________________

        On Error Resume Next
    '    Klasörün içindeki tüm dosyaları sil (txt, docm vb.)
        ContSay = 0
        KontrolFile = Dir(DestOpUserFolder & "*.???")
        Do While KontrolFile <> ""
            ContSay = ContSay + 1
            KontrolFile = Dir()
        Loop
        If ContSay > 0 Then
            Kill DestOpUserFolder & "*.???"
        End If
    End If

    'Dosyayı şablondan operasyon klasörüne kopyala ve adını değiştir.
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.CopyFile (SourceUstYaziNormal), DestOpUserFolder & ReNameUstYaziNormal & ".docm", True
'________________________________________

    'Oluşturulacak dosyayı aç
    On Error Resume Next
    Set objWord = GetObject(, "Word.Application")
    Set objWord = GetObject(, "Word.Application")
    Set objWord = GetObject(, "Word.Application")
    Set objWord = GetObject(, "Word.Application")
    Set objWord = GetObject(, "Word.Application")
    If objWord Is Nothing Then
        'MsgBox "Dosya oluşturmada CreateObject methodu kullanılacak."
        Set objWord = CreateObject("Word.Application")
        objWord.Visible = False
    End If
    objWord.Documents.Open FileName:=DestOpUserFolder & ReNameUstYaziNormal & ".docm"
    objWord.Visible = True
    objWord.Activate 'Ekrana getirir.
    'objDoc.Activate 'Ekrana getirmez.
    objWord.Application.WindowState = wdWindowStateMaximize
    Set objDoc = GetObject(DestOpUserFolder & ReNameUstYaziNormal & ".docm")
'________________________________________

    'Birim
    Birimx = UCase(Replace(Replace(Worksheets(2).Cells(6, 99).Value, "i", "I"), "ı", "I")) & " UNIT"
    objDoc.Tables(1).Cell(Row:=1, Column:=1).Range.Text = Birimx

    Birimx = UCase(Replace(Replace(Worksheets(2).Cells(6, 99).Value, "i", "I"), "ı", "I"))
    'Üst yazı tarihi
    objDoc.Tables(1).Cell(Row:=1, Column:=2).Range.Text = Birimx & ", " & FormatEnglishDate(Cells(ActiveCell.Row, 83).Value) '(Format(Cells(ActiveCell.Row, 83).Value, "d mmmm yyyy"))
    'Yazı no
    objDoc.Tables(1).Cell(Row:=4, Column:=2).Range.Text = Cells(ActiveCell.Row, 84).Value
    'Sigorta türü
    Ek2 = Cells(ActiveCell.Row, 71).Value
    Ek2 = Right(Ek2, Len(Ek2) - InStr(Ek2, "/"))
    objDoc.Tables(1).Cell(Row:=7, Column:=2).Range.Text = Ek2

    'Muhatap
    IlBuyukHarf = UCase(Replace(Replace(Cells(ActiveCell.Row, 19).Value, "i", "I"), "ı", "I"))
    IlKucukHarf = Cells(ActiveCell.Row, 19).Value
    If Cells(ActiveCell.Row, 20).Value <> "" Then
        IlceBuyukHarf = UCase(Replace(Replace(Cells(ActiveCell.Row, 20).Value, "i", "I"), "ı", "I"))
        IlceKucukHarf = Cells(ActiveCell.Row, 20).Value
    Else
        IlceBuyukHarf = ""
        IlceKucukHarf = ""
    End If
    If Cells(ActiveCell.Row, 20).Value <> "" Then
        Bolum3 = IlceKucukHarf & "/" & IlBuyukHarf
    Else
        Bolum3 = IlBuyukHarf
    End If
    
    
    'YENİ MUHATAP TEMASI
    
    M2 = False
    M3 = False
    M4 = False
'    Ifv = False
'    Ify = False
    'TipB = False
    BStr = 10
    '1. sayfadan sonraki üst bilgi tanımları
    ustbilgitarih = (Format(Cells(ActiveCell.Row, 83).Value, "dd.mm.yyyy"))
    ustbilgisayi = Cells(ActiveCell.Row, 84).Value
    
    '4'lük
    If Cells(ActiveCell.Row, 48).Value <> "" Then
        If Cells(ActiveCell.Row, 47).Value = "Provincial Directorate B" Or Cells(ActiveCell.Row, 47).Value = "Provincial Directorate C" Or _
        Cells(ActiveCell.Row, 47).Value = "Provincial Directorate D" Or Cells(ActiveCell.Row, 47).Value = "Provincial Directorate E" Then 'VALİLİK
            objDoc.Tables(1).Cell(Row:=BStr, Column:=1).Range.Text = IlBuyukHarf & " PROVINCIAL GOVERNORSHIP"
            objDoc.Tables(1).Cell(Row:=BStr + 1, Column:=1).Range.Text = Cells(ActiveCell.Row, 47).Value  'UCase(Replace(Replace(Cells(ActiveCell.Row, 47).Value, "i", "I"), "ı", "I"))
            objDoc.Tables(1).Cell(Row:=BStr + 2, Column:=1).Range.Text = "(" & Cells(ActiveCell.Row, 48).Value & ")"
            objDoc.Tables(1).Cell(Row:=BStr + 3, Column:=1).Range.Text = Bolum3
            objDoc.Tables(1).Cell(Row:=BStr + 3, Column:=1).Range.Font.Underline = wdUnderlineSingle
            M4 = True
            
            ustbilgimuhatap = IlKucukHarf & " Provincial Governorship " & Cells(ActiveCell.Row, 47).Value
        ElseIf Cells(ActiveCell.Row, 47).Value = "District Directorate B" Or Cells(ActiveCell.Row, 47).Value = "District Directorate C" Or _
        Cells(ActiveCell.Row, 47).Value = "District Directorate D" Or Cells(ActiveCell.Row, 47).Value = "District Directorate E" Then 'KAYMAKAMLIK
            objDoc.Tables(1).Cell(Row:=BStr, Column:=1).Range.Text = IlceBuyukHarf & " DISTRICT GOVERNORSHIP"
            objDoc.Tables(1).Cell(Row:=BStr + 1, Column:=1).Range.Text = Cells(ActiveCell.Row, 47).Value 'UCase(Replace(Replace(Cells(ActiveCell.Row, 47).Value, "i", "I"), "ı", "I"))
            objDoc.Tables(1).Cell(Row:=BStr + 2, Column:=1).Range.Text = "(" & Cells(ActiveCell.Row, 48).Value & ")"
            objDoc.Tables(1).Cell(Row:=BStr + 3, Column:=1).Range.Text = Bolum3
            objDoc.Tables(1).Cell(Row:=BStr + 3, Column:=1).Range.Font.Underline = wdUnderlineSingle
            M4 = True
            
            ustbilgimuhatap = IlceKucukHarf & " District Governorship " & Cells(ActiveCell.Row, 47).Value
        Else 'YARGI 3'lük
            Bolum1 = UCase(Replace(Replace(Cells(ActiveCell.Row, 47).Value, "i", "I"), "ı", "I"))
            If InStr(Bolum1, "X.X. ") > 0 Then
                objDoc.Tables(1).Cell(Row:=BStr, Column:=1).Range.Text = Mid(Bolum1, 6, Len(Bolum1))
                objDoc.Tables(1).Cell(Row:=BStr + 1, Column:=1).Range.Text = "(" & Cells(ActiveCell.Row, 48).Value & ")"
                objDoc.Tables(1).Cell(Row:=BStr + 2, Column:=1).Range.Text = Bolum3
                objDoc.Tables(1).Cell(Row:=BStr + 2, Column:=1).Range.Font.Underline = wdUnderlineSingle
                M3 = True
                
                ustbilgimuhatap = Mid(Cells(ActiveCell.Row, 47).Value, 6, Len(Cells(ActiveCell.Row, 47).Value))
            Else
                objDoc.Tables(1).Cell(Row:=BStr, Column:=1).Range.Text = Bolum1
                objDoc.Tables(1).Cell(Row:=BStr + 1, Column:=1).Range.Text = "(" & Cells(ActiveCell.Row, 48).Value & ")"
                objDoc.Tables(1).Cell(Row:=BStr + 2, Column:=1).Range.Text = Bolum3
                objDoc.Tables(1).Cell(Row:=BStr + 2, Column:=1).Range.Font.Underline = wdUnderlineSingle
                M3 = True

                ustbilgimuhatap = Cells(ActiveCell.Row, 47).Value
            End If
        End If
    End If
    '3'lük
    If Cells(ActiveCell.Row, 48).Value = "" Then
        If Cells(ActiveCell.Row, 47).Value = "Provincial Directorate B" Or Cells(ActiveCell.Row, 47).Value = "Provincial Directorate C" Or _
        Cells(ActiveCell.Row, 47).Value = "Provincial Directorate D" Or Cells(ActiveCell.Row, 47).Value = "Provincial Directorate E" Then 'VALİLİK
            objDoc.Tables(1).Cell(Row:=BStr, Column:=1).Range.Text = IlBuyukHarf & " PROVINCIAL GOVERNORSHIP"
            objDoc.Tables(1).Cell(Row:=BStr + 1, Column:=1).Range.Text = Cells(ActiveCell.Row, 47).Value 'UCase(Replace(Replace(Cells(ActiveCell.Row, 47).Value, "i", "I"), "ı", "I"))
            objDoc.Tables(1).Cell(Row:=BStr + 2, Column:=1).Range.Text = Bolum3
            objDoc.Tables(1).Cell(Row:=BStr + 2, Column:=1).Range.Font.Underline = wdUnderlineSingle
            M3 = True
            
            ustbilgimuhatap = IlKucukHarf & " Provincial Governorship " & Cells(ActiveCell.Row, 47).Value
        ElseIf Cells(ActiveCell.Row, 47).Value = "District Directorate B" Or Cells(ActiveCell.Row, 47).Value = "District Directorate C" Or _
        Cells(ActiveCell.Row, 47).Value = "District Directorate D" Or Cells(ActiveCell.Row, 47).Value = "District Directorate E" Then
            objDoc.Tables(1).Cell(Row:=BStr, Column:=1).Range.Text = IlceBuyukHarf & " DISTRICT GOVERNORSHIP"
            objDoc.Tables(1).Cell(Row:=BStr + 1, Column:=1).Range.Text = Cells(ActiveCell.Row, 47).Value 'UCase(Replace(Replace(Cells(ActiveCell.Row, 47).Value, "i", "I"), "ı", "I"))
            objDoc.Tables(1).Cell(Row:=BStr + 2, Column:=1).Range.Text = Bolum3
            objDoc.Tables(1).Cell(Row:=BStr + 2, Column:=1).Range.Font.Underline = wdUnderlineSingle
            M3 = True
            
            ustbilgimuhatap = IlceKucukHarf & " District Governorship " & Cells(ActiveCell.Row, 47).Value
        Else 'YARGI 2'lik
            Bolum1 = UCase(Replace(Replace(Cells(ActiveCell.Row, 47).Value, "i", "I"), "ı", "I"))
            If InStr(Bolum1, "X.X. ") > 0 Then
                objDoc.Tables(1).Cell(Row:=BStr, Column:=1).Range.Text = Mid(Bolum1, 6, Len(Bolum1))
                objDoc.Tables(1).Cell(Row:=BStr + 1, Column:=1).Range.Text = Bolum3
                objDoc.Tables(1).Cell(Row:=BStr + 1, Column:=1).Range.Font.Underline = wdUnderlineSingle
                M2 = True

                ustbilgimuhatap = Mid(Cells(ActiveCell.Row, 47).Value, 6, Len(Cells(ActiveCell.Row, 47).Value))
            Else
                objDoc.Tables(1).Cell(Row:=BStr, Column:=1).Range.Text = Bolum1
                objDoc.Tables(1).Cell(Row:=BStr + 1, Column:=1).Range.Text = Bolum3
                objDoc.Tables(1).Cell(Row:=BStr + 1, Column:=1).Range.Font.Underline = wdUnderlineSingle
                M2 = True

                ustbilgimuhatap = Cells(ActiveCell.Row, 47).Value
            End If
        End If
    End If
    
    
    'Üst yazı gövde metni
    'tipAnın/tipAların (TipA adedi)
    AdetTopla = Application.Sum(Range(Cells(IlkSira, 58), Cells(SonSira, 58)))
    If AdetTopla = 1 Then
        Ek1 = "item"
        Ek5 = "has"
    ElseIf AdetTopla > 1 Then
        Ek1 = "items"
        Ek5 = "have"
    End If
    
    Ek2 = Cells(ActiveCell.Row, 35).Value 'Teslimat tarihi
    Ek3 = Cells(ActiveCell.Row, 30).Value 'FinansalBirim
    Ek4 = Cells(ActiveCell.Row, 26).Value 'Temax no
    Ek6 = Cells(ActiveCell.Row, 31).Value 'Teslimatı yapan birim

    Bolum1 = "As a result of the initial inspection of the Type B " & Ek1 & " delivered to our unit on " & Ek2 & " by " & Ek3 & " – " & Ek6 & " for the purpose of xxxxx xxxxx xxxxx, the Type B " & _
    Ek1 & ", listed in the attached statement and assigned the theme number " & Ek4 & " by our office, " & Ek5 & " been evaluated as invalid."
    Bolum2 = "The situation has been reported to the relevant financial unit, and copies of the official statement and the letter addressed to the mentioned unit are enclosed herewith, " & _
    "along with the " & Ek1 & " in question, for your information."
    Bolum3 = Chr(9) & "Respectfully submitted for your information."
    Birlestir = Chr(9) & Bolum1 & vbNewLine & Chr(9) & Bolum2
    
    objDoc.Tables(2).Cell(Row:=1, Column:=1).Range.Text = Birlestir
    objDoc.Tables(2).Cell(Row:=3, Column:=1).Range.Text = Bolum3

    'imzalar
    objDoc.Tables(3).Cell(Row:=4, Column:=2).Range.Text = Cells(ActiveCell.Row, 205).Value 'Ad Soyad1
    objDoc.Tables(3).Cell(Row:=5, Column:=2).Range.Text = Cells(ActiveCell.Row, 206).Value 'Unvan1
    objDoc.Tables(3).Cell(Row:=4, Column:=3).Range.Text = Cells(ActiveCell.Row, 208).Value 'Ad Soyad2
    objDoc.Tables(3).Cell(Row:=5, Column:=3).Range.Text = Cells(ActiveCell.Row, 209).Value 'Unvan2
    
    'Ekler
    objDoc.Tables(3).Cell(Row:=8, Column:=1).Range.Text = "Attachment:"

    Ek2 = Cells(ActiveCell.Row, 71).Value 'Kapalı Package A
    Ek2 = Left(Ek2, InStr(Ek2, "/") - 1)
    If Cells(ActiveCell.Row, 72).Value > 1 Then objDoc.Tables(3).Cell(Row:=8, Column:=2).Range.Text = " 1) Enclosed " & Ek2 & " (" & Cells(ActiveCell.Row, 72).Value & " pieces)"
    If Cells(ActiveCell.Row, 72).Value < 2 Then objDoc.Tables(3).Cell(Row:=8, Column:=2).Range.Text = " 1) Enclosed " & Ek2 & " (" & Cells(ActiveCell.Row, 72).Value & " piece)"
    
    x = Application.Sum(Range(Cells(IlkSira, 171), Cells(SonSira, 171))) 'Statement 2 toplam sayfa sayısı
    If x > 1 Then objDoc.Tables(3).Cell(Row:=9, Column:=2).Range.Text = " 2) Statement 2 (" & x & " pages)"
    If x < 2 Then objDoc.Tables(3).Cell(Row:=9, Column:=2).Range.Text = " 2) Statement 2 (" & x & " page)"
    
    x = Application.Sum(Range(Cells(IlkSira, 172), Cells(SonSira, 172))) 'FinansalBirim üst yazı toplam sayfa sayısı
    If x > 1 Then objDoc.Tables(3).Cell(Row:=10, Column:=2).Range.Text = " 3) Photocopy of the Letter to Relevant Unit (" & x & " pages)"
    If x < 2 Then objDoc.Tables(3).Cell(Row:=10, Column:=2).Range.Text = " 3) Photocopy of the Letter to Relevant Unit (" & x & " page)"
    
    x = Application.Sum(Range(Cells(IlkSira, 169), Cells(SonSira, 169))) + Cells(ActiveCell.Row, 50).Value 'Rapor3 Tutanağı
    objDoc.Tables(3).Cell(Row:=11, Column:=2).Range.Text = " 4) Attached Statement (total of " & x & " pages)"



    '__________________________DİNAMİK SAYFA YAPISI
    
    'İlgi ve gövde metninin kaç satırdan oluştuğu bilgisinin elde edilmesinde kullanılıyor.
    'Set IlgiRng = objDoc.Tables(1).Cell(Row:=15, Column:=3).Range
    Set Govde1Rng = objDoc.Tables(2).Cell(Row:=1, Column:=1).Range
    Set Govde2Rng = objDoc.Tables(2).Cell(Row:=2, Column:=1).Range
    
    'IlgiStrSay = Govde1Rng.Information(wdFirstCharacterLineNumber) - IlgiRng.Information(wdFirstCharacterLineNumber) - 1
    Govde1StrSay = Govde2Rng.Information(wdFirstCharacterLineNumber) - Govde1Rng.Information(wdFirstCharacterLineNumber)
    'MsgBox "İlgi Satır Sayısı: " & IlgiStrSay & vbNewLine & vbNewLine & "Gövde 1 Satır Sayısı: " & Govde1StrSay
    'IlgiFarkSay = IlgiStrSay - 1
    IlgiFarkSay = 0
    Govde1FarkSay = Govde1StrSay + 1 '- 3 'Varsayılan 2 paragrafta (1 rowda) toplam 3 satıra göre sıfırlandı.
    CokluSayfa = 0
    'MsgBox Govde1FarkSay

    
    objDoc.Tables(3).Rows(12).Delete 'Ek sonrası
    'Dinamik sayfa düzeni
    If M4 = True Then
        For i = 1 To 2
            objDoc.Tables(1).Rows(14).Delete 'Muhatap sonrası
        Next i
        If IlgiFarkSay + Govde1FarkSay < 17 Then
            Govde1FarkSay = Govde1FarkSay + 2
'            MsgBox IlgiFarkSay + Govde1FarkSay
            If IlgiFarkSay + Govde1FarkSay > 8 Then
                For i = 1 To IlgiFarkSay + Govde1FarkSay
                    If i = 9 Then
                        objDoc.Tables(3).Rows(6).Delete 'Ek öncesi
                    ElseIf i = 10 Then
                        objDoc.Tables(3).Rows(6).Delete 'Ek öncesi
                    ElseIf i = 11 Then
                        objDoc.Tables(3).Rows(1).Delete 'İmza öncesi
                    ElseIf i = 12 Then
                        objDoc.Tables(3).Rows(1).Delete 'İmza öncesi
                    ElseIf i = 13 Then
                        objDoc.Tables(1).Rows(8).Delete 'Muhatap öncesi
                    ElseIf i = 14 Then
                        objDoc.Tables(1).Rows(8).Delete 'Muhatap öncesi
                    End If
                Next i
            Else
                If (IlgiFarkSay + Govde1FarkSay) < 8 Then
                    For i = 1 To 8 - (IlgiFarkSay + Govde1FarkSay)
                        'Boş satır ekle
                        With objDoc.Tables(3)
                            .Rows.Add BeforeRow:=objDoc.Tables(3).Rows(8)
                        End With
                    Next i
                End If
            End If
        Else
            CokluSayfa = 1
        End If
    ElseIf M3 = True Then
        For i = 1 To 3
            objDoc.Tables(1).Rows(13).Delete 'Muhatap sonrası
        Next i
        If IlgiFarkSay + Govde1FarkSay < 16 Then
            Govde1FarkSay = Govde1FarkSay + 1
'            MsgBox IlgiFarkSay + Govde1FarkSay
            If IlgiFarkSay + Govde1FarkSay > 8 Then
                For i = 9 To IlgiFarkSay + Govde1FarkSay
                    If i = 9 Then
                        objDoc.Tables(3).Rows(6).Delete 'Ek öncesi
                    ElseIf i = 10 Then
                        objDoc.Tables(3).Rows(6).Delete 'Ek öncesi
                    ElseIf i = 11 Then
                        objDoc.Tables(3).Rows(1).Delete 'İmza öncesi
                    ElseIf i = 12 Then
                        objDoc.Tables(3).Rows(1).Delete 'İmza öncesi
                    ElseIf i = 13 Then
                        objDoc.Tables(1).Rows(8).Delete 'Muhatap öncesi
                    ElseIf i = 14 Then
                        objDoc.Tables(1).Rows(8).Delete 'Muhatap öncesi
                    End If
                Next i
            Else
                If (IlgiFarkSay + Govde1FarkSay) < 8 Then
                    For i = 1 To 8 - (IlgiFarkSay + Govde1FarkSay)
                        'Boş satır ekle
                        With objDoc.Tables(3)
                            .Rows.Add BeforeRow:=objDoc.Tables(3).Rows(8)
                        End With
                    Next i
                End If
            End If
        Else
            CokluSayfa = 1
        End If
    ElseIf M2 = True Then
        For i = 1 To 4
            objDoc.Tables(1).Rows(12).Delete 'Muhatap sonrası
        Next i

        If IlgiFarkSay + Govde1FarkSay < 15 Then
'            MsgBox IlgiFarkSay + Govde1FarkSay
            If IlgiFarkSay + Govde1FarkSay > 8 Then
                For i = 9 To IlgiFarkSay + Govde1FarkSay
                    If i = 9 Then
                        objDoc.Tables(3).Rows(6).Delete 'Ek öncesi
                    ElseIf i = 10 Then
                        objDoc.Tables(3).Rows(6).Delete 'Ek öncesi
                    ElseIf i = 11 Then
                        objDoc.Tables(3).Rows(1).Delete 'İmza öncesi
                    ElseIf i = 12 Then
                        objDoc.Tables(3).Rows(1).Delete 'İmza öncesi
                    ElseIf i = 13 Then
                        objDoc.Tables(1).Rows(8).Delete 'Muhatap öncesi
                    ElseIf i = 14 Then
                        objDoc.Tables(1).Rows(8).Delete 'Muhatap öncesi
                    End If
                Next i
            Else
                If (IlgiFarkSay + Govde1FarkSay) < 8 Then
                    For i = 1 To 8 - (IlgiFarkSay + Govde1FarkSay)
                        'Boş satır ekle
                        With objDoc.Tables(3)
                            .Rows.Add BeforeRow:=objDoc.Tables(3).Rows(8)
                        End With
                    Next i
                End If
            End If
        Else
            CokluSayfa = 1
        End If
    End If

    objDoc.Paragraphs(objDoc.Paragraphs.Count).Range.Font.Size = 1


    'Footer
    If objDoc.ActiveWindow.View.SplitSpecial <> wdPaneNone Then
        objDoc.ActiveWindow.Panes(2).Close
    End If
    If objDoc.ActiveWindow.ActivePane.View.Type = wdNormalView Or objDoc.ActiveWindow.ActivePane.View.Type = wdOutlineView Then
        objDoc.ActiveWindow.ActivePane.View.Type = wdPrintView
    End If

    With objDoc
        With .Sections(1).Headers(wdHeaderFooterPrimary).Range
            .InsertAfter Text:="Page "
            .Fields.Add Range:=.Characters.Last, Type:=wdFieldEmpty, Text:="PAGE", PreserveFormatting:=False
            .InsertAfter Text:=" of the letter dated " & ustbilgitarih & _
                             ", reference number " & ustbilgisayi & _
                             ", addressed to the " & ustbilgimuhatap
        End With
    End With
    
    objDoc.ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument


    'PARAF EKLEME İŞLEMİ
    objDoc.Sections(1).Footers(wdHeaderFooterFirstPage).Range.Tables(1).Cell(Row:=1, Column:=2).Range.Text = AdSoyadParaf & vbNewLine & _
                                                                                       UnvanParaf & vbNewLine & _
                                                                                       TelParaf
    objDoc.Sections(1).Footers(wdHeaderFooterPrimary).Range.Tables(1).Cell(Row:=1, Column:=2).Range.Text = AdSoyadParaf & vbNewLine & _
                                                                                       UnvanParaf & vbNewLine & _
                                                                                       TelParaf
                                                                                       

Tekrarla:
    'Sayfa sayısı kaydet komutuna bağlandı.
    objDoc.Save
    'Sayfayı text dosyasından çek
    TxtFileUstYazi = DestOpUserFolder & "Cover Letter Page Count.txt"
    #If UseADO Then
        Const adReadLine = -2&
        With CreateObject("ADODB.Stream")
            .Open
            .LoadFromFile TxtFileUstYazi
            Do Until .EOS
                TotalSayfaUstYazi = .ReadText(adReadLine)
            Loop
            .Close
        End With
    #Else 'UseFSO
        Const UstYaziTristateTrue = -1&
        With CreateObject("Scripting.FileSystemObject")
            With .OpenTextFile(TxtFileUstYazi, Format:=UstYaziTristateTrue)
                Do Until .AtEndOfStream
                    TotalSayfaUstYazi = .ReadLine
                Loop
                .Close
            End With
        End With
    #End If
    'MsgBox "Üst Yazı: " & TotalSayfaUstYazi

    'Üst Yazı sayfa sayısı
    Cells(ActiveCell.Row, 173).Value = TotalSayfaUstYazi

    objWord.Visible = True
    objWord.Activate

    If TumDoc = True Then
        objWord.Activate
'        For i = 1 To SayPrt
'            objDoc.PrintOut
'        Next i
        objDoc.PrintOut Background:=False, Copies:=SayPrt
        'objWord.Documents.Save
        objDoc.Close SaveChanges:=False
        objWord.Visible = False
    End If

    'Nesneleri temizle
    Set fso = Nothing
    Set objWord = Nothing
    Set objDoc = Nothing

Son:

If IlceSakla <> "" Then
    Cells(ActiveCell.Row, 20).Value = IlceSakla
End If

'Worksheets(5).Activate
Application.ScreenUpdating = True
Application.DisplayAlerts = True
'Application.EnableEvents = True

'End If

End Sub

Sub IslemGunluguRapor3_2()
Dim AutoPath As String, IslemGunlugu As String, IslemGunlukleriKlasor As String, WsIslemGunlugu As Object
Dim Kenarlar As Range
Dim BulIslemGunlugu As Range, AralikSay As Integer, KayitDefSiraNo As Long

Dim i As Long, j As Long
Dim YeniIslem As Long, Maxi As Integer
Dim OpenControl As String
Dim WsRapor As Object

Dim IlkSira As Long, SonSira As Long
Dim IslemGunluguIlkSiraBul As Range, IslemGunluguSonSiraBul As Range, IslemGunluguIlkSira As Long, IslemGunluguSonSira As Long
Dim GelenTema As String, Fark As Long, DelControl As Boolean


Application.DisplayAlerts = False
Application.EnableEvents = False
Application.ScreenUpdating = False


'Pathfinder...
AutoPath = ThisWorkbook.Path
IslemGunlukleriKlasor = AutoPath & "\System Files\System Templates\Registry Reports\"
IslemGunlugu = IslemGunlukleriKlasor & "System Registry Report 2.1.xlsx"

'Check the Process Logs folder name.
If Not Dir(IslemGunlukleriKlasor, vbDirectory) <> vbNullString Then
    MsgBox IslemGunlukleriKlasor & " folder cannot be accessed. The folder named 'Registry Reports' might have been renamed or deleted.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Out
End If
If Not Dir(IslemGunlugu, vbDirectory) <> vbNullString Then
    MsgBox IslemGunlugu & " folder cannot be accessed. The names of folders and/or files in this directory might have been changed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Out
End If

Maxi = MaxiAktar
YeniIslem = YeniIslemAktar
'StrTime = Format(Now, "ddmmyyyyhhmmss")
DelControl = False

'Modülün Rapor sayfasında bulunan başlangıç ve bitiş satır numaraları
IlkSira = YeniIslem
SonSira = YeniIslem + Maxi
Set WsRapor = ThisWorkbook.Worksheets(5)
'WsRapor.Unprotect Password:="123"

IlceSakla = ""
If InStr(WsRapor.Cells(IlkSira, 20).Value, " Organization A") <> 0 Then
    IlceSakla = WsRapor.Cells(IlkSira, 20).Value
    WsRapor.Cells(IlkSira, 20).Value = ""
End If

'Aylık ayraçlar
If WsRapor.Cells(IlkSira, 23).Value <> "" Then
    ModulTarih = WsRapor.Cells(IlkSira, 23).Value
    ModulAyrac = "01" & Right(ModulTarih, 8)
Else 'işlemin yapıldığı günü esas al
    ModulTarih = Format(Date, "dd.mm.yyyy")
    ModulAyrac = "01" & Right(ModulTarih, 8)
End If


'İşlem günlüğü açıksa kaydet ve kapat.
OpenControl = IsWorkBookOpen(IslemGunlugu)
If OpenControl = True Then
    Workbooks("System Registry Report 2.1.xlsx").Save
    Workbooks("System Registry Report 2.1.xlsx").Close SaveChanges:=False
End If

'İşlem günlüğü aç
Workbooks.Open (IslemGunlugu)
Set WsIslemGunlugu = Workbooks("System Registry Report 2.1.xlsx").Worksheets(1)

WsIslemGunlugu.Unprotect Password:="123"

WsIslemGunlugu.Columns("B:C").EntireColumn.Hidden = False


'İşlem günlüğünde yoksa ve işlem tipA değilse prosedürden çık (TipA/TipB)
Set IslemGunluguIlkSiraBul = WsIslemGunlugu.Range("B7:B100000").Find(What:=WsRapor.Cells(IlkSira, 165).Value, SearchDirection:=xlNext, _
                SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole) 'zaman damgasını ara
If Not IslemGunluguIlkSiraBul Is Nothing Then
    'Nothing
Else
    If WsRapor.Cells(IlkSira, 28).Value <> "Type A" Then
        GoTo Son
    End If
End If

'____________HAZIRLIK

'İşlem günlüğünde başlangıç ve bitiş satırlarını tespit et.
Say1IslemGunlugu = WsIslemGunlugu.Range("B100000").End(xlUp).Row
Say2IslemGunlugu = WsIslemGunlugu.Range("C100000").End(xlUp).Row
SayAyracIslemGunlugu = WsIslemGunlugu.Range("E100000").End(xlUp).Row
'İşlem günlüğünde ayraçları oluştur
If Say1IslemGunlugu < 7 And SayAyracIslemGunlugu < 7 Then

    Say1IslemGunlugu = 6
    Say2IslemGunlugu = 6
    SayAyracIslemGunlugu = 6

    i = 6
    IslemGunluguAyrac = "01.01" & Right(ModulAyrac, 5)
    ModulAyrac = DateAdd("m", 1, CDate(ModulAyrac))
    'MsgBox IslemGunluguAyrac & " and " & ModulAyrac '''
    Do Until CDate(IslemGunluguAyrac) >= CDate(ModulAyrac)
        i = i + 1
        With WsIslemGunlugu.Range("E" & i & ":T" & i)
            .WrapText = False
            .Font.Bold = True
            .Font.name = "Open Sans"
            .Font.Size = 9
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = RGB(174, 185, 194) 'RGB(201, 216, 230) 'RGB(174, 185, 194)
            .HorizontalAlignment = xlLeft
            .NumberFormat = "mmmm yyyy"
        End With
        WsIslemGunlugu.Cells(i, 5).Value = CDate(IslemGunluguAyrac)
        IslemGunluguAyrac = DateAdd("m", 1, CDate(IslemGunluguAyrac))
    Loop
    ModulAyrac = DateAdd("m", -1, CDate(ModulAyrac))

ElseIf Say1IslemGunlugu < 7 And SayAyracIslemGunlugu >= 7 Then

    i = SayAyracIslemGunlugu
    IslemGunluguAyrac = WsIslemGunlugu.Cells(SayAyracIslemGunlugu, 5).Value
    'MsgBox IslemGunluguAyrac & " and " & ModulAyrac '''
    Do Until CDate(IslemGunluguAyrac) >= CDate(ModulAyrac)
        i = i + 1
        With WsIslemGunlugu.Range("E" & i & ":T" & i)
            .WrapText = False
            .Font.Bold = True
            .Font.name = "Open Sans"
            .Font.Size = 9
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = RGB(174, 185, 194) 'RGB(201, 216, 230) 'RGB(174, 185, 194)
            .HorizontalAlignment = xlLeft
            .NumberFormat = "mmmm yyyy"
        End With
        IslemGunluguAyrac = DateAdd("m", 1, CDate(IslemGunluguAyrac))
        WsIslemGunlugu.Cells(i, 5).Value = CDate(IslemGunluguAyrac)
    Loop

ElseIf Say1IslemGunlugu >= 7 And SayAyracIslemGunlugu >= 7 Then

    SayMax = WorksheetFunction.Max(Say2IslemGunlugu, SayAyracIslemGunlugu)
    i = SayMax
    IslemGunluguAyrac = WsIslemGunlugu.Cells(SayAyracIslemGunlugu, 5).Value
    'MsgBox IslemGunluguAyrac & " and " & ModulAyrac '''
    Do Until CDate(IslemGunluguAyrac) >= CDate(ModulAyrac)
        i = i + 1
        With WsIslemGunlugu.Range("E" & i & ":T" & i)
            .WrapText = False
            .Font.Bold = True
            .Font.name = "Open Sans"
            .Font.Size = 9
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = RGB(174, 185, 194) 'RGB(201, 216, 230) 'RGB(174, 185, 194)
            .HorizontalAlignment = xlLeft
            .NumberFormat = "mmmm yyyy"
        End With
        IslemGunluguAyrac = DateAdd("m", 1, CDate(IslemGunluguAyrac))
        WsIslemGunlugu.Cells(i, 5).Value = CDate(IslemGunluguAyrac)
    Loop

End If

'MsgBox "IlkSira: " & IlkSira & " ve SonSira: " & SonSira
'GoTo Son

'GELEN TEMA
GelenTema = ""
If WsRapor.Cells(IlkSira, 47).Value = "Provincial Directorate B" Then
    If WsRapor.Cells(IlkSira, 48).Value <> "" Then
        GelenTema = WsRapor.Cells(IlkSira, 19).Value & " Provincial Governorship Provincial Directorate B " & WsRapor.Cells(IlkSira, 48).Value
    Else
        GelenTema = WsRapor.Cells(IlkSira, 19).Value & " Provincial Governorship Provincial Directorate B"
    End If
ElseIf WsRapor.Cells(IlkSira, 47).Value = "District Directorate B" Then
    If WsRapor.Cells(IlkSira, 48).Value <> "" Then
        GelenTema = WsRapor.Cells(IlkSira, 20).Value & " District Governorship District Directorate B " & WsRapor.Cells(IlkSira, 48).Value
    Else
        GelenTema = WsRapor.Cells(IlkSira, 20).Value & " District Governorship District Directorate B"
    End If
ElseIf WsRapor.Cells(IlkSira, 47).Value = "Provincial Directorate C" Then
    If WsRapor.Cells(IlkSira, 48).Value <> "" Then
        GelenTema = WsRapor.Cells(IlkSira, 19).Value & " Provincial Governorship Provincial Directorate C " & WsRapor.Cells(IlkSira, 48).Value
    Else
        GelenTema = WsRapor.Cells(IlkSira, 19).Value & " Provincial Governorship Provincial Directorate C"
    End If
ElseIf WsRapor.Cells(IlkSira, 47).Value = "District Directorate C" Then
    If WsRapor.Cells(IlkSira, 48).Value <> "" Then
        GelenTema = WsRapor.Cells(IlkSira, 20).Value & " District Governorship District Directorate C " & WsRapor.Cells(IlkSira, 48).Value
    Else
        GelenTema = WsRapor.Cells(IlkSira, 20).Value & " District Governorship District Directorate C"
    End If
ElseIf WsRapor.Cells(IlkSira, 47).Value = "Provincial Directorate D" Then
    If WsRapor.Cells(IlkSira, 48).Value <> "" Then
        GelenTema = WsRapor.Cells(IlkSira, 19).Value & " Provincial Governorship Provincial Directorate D " & WsRapor.Cells(IlkSira, 48).Value
    Else
        GelenTema = WsRapor.Cells(IlkSira, 19).Value & " Provincial Governorship Provincial Directorate D"
    End If
ElseIf WsRapor.Cells(IlkSira, 47).Value = "District Directorate D" Then
    If WsRapor.Cells(IlkSira, 48).Value <> "" Then
        GelenTema = WsRapor.Cells(IlkSira, 20).Value & " District Governorship District Directorate D " & WsRapor.Cells(IlkSira, 48).Value
    Else
        GelenTema = WsRapor.Cells(IlkSira, 20).Value & " District Governorship District Directorate D"
    End If
ElseIf WsRapor.Cells(IlkSira, 47).Value = "Provincial Directorate E" Then
    If WsRapor.Cells(IlkSira, 48).Value <> "" Then
        GelenTema = WsRapor.Cells(IlkSira, 19).Value & " Provincial Governorship Provincial Directorate E " & WsRapor.Cells(IlkSira, 48).Value
    Else
        GelenTema = WsRapor.Cells(IlkSira, 19).Value & " Provincial Governorship Provincial Directorate E"
    End If
ElseIf WsRapor.Cells(IlkSira, 47).Value = "District Directorate E" Then
    If WsRapor.Cells(IlkSira, 48).Value <> "" Then
        GelenTema = WsRapor.Cells(IlkSira, 20).Value & " District Governorship District Directorate E " & WsRapor.Cells(IlkSira, 48).Value
    Else
        GelenTema = WsRapor.Cells(IlkSira, 20).Value & " District Governorship District Directorate E"
    End If
ElseIf InStr(WsRapor.Cells(IlkSira, 47).Value, "General Directorate") <> 0 Or InStr(WsRapor.Cells(IlkSira, 47).Value, "Regional Directorate") <> 0 Then
    If WsRapor.Cells(IlkSira, 48).Value <> "" Then
        GelenTema = WorksheetFunction.Proper(ThisWorkbook.Worksheets(2).Cells(6, 111).Value) & " " & WsRapor.Cells(IlkSira, 47).Value & " " & WsRapor.Cells(IlkSira, 48).Value
    Else
        GelenTema = WorksheetFunction.Proper(ThisWorkbook.Worksheets(2).Cells(6, 111).Value) & " " & WsRapor.Cells(IlkSira, 47).Value
    End If
Else
    If WsRapor.Cells(IlkSira, 48).Value <> "" Then
        If InStr(WsRapor.Cells(IlkSira, 47).Value, "X.X. ") > 0 Then
            GelenTema = Mid(WsRapor.Cells(IlkSira, 47).Value, 6, Len(WsRapor.Cells(IlkSira, 47).Value)) & " " & WsRapor.Cells(IlkSira, 48).Value
        Else
            GelenTema = WsRapor.Cells(IlkSira, 47).Value & " " & WsRapor.Cells(IlkSira, 48).Value
        End If
    Else
        If InStr(WsRapor.Cells(IlkSira, 47).Value, "X.X. ") > 0 Then
            GelenTema = Mid(WsRapor.Cells(IlkSira, 47).Value, 6, Len(WsRapor.Cells(IlkSira, 47).Value))
        Else
            GelenTema = WsRapor.Cells(IlkSira, 47).Value
        End If
    End If
End If

'____________OPERASYONLAR

'İşlem günlüğünde başlangıç ve bitiş satırlarını tespit et.
Say1IslemGunlugu = WsIslemGunlugu.Range("B100000").End(xlUp).Row
Say2IslemGunlugu = WsIslemGunlugu.Range("C100000").End(xlUp).Row
SayAyracIslemGunlugu = WsIslemGunlugu.Range("E100000").End(xlUp).Row

Set IslemGunluguIlkSiraBul = WsIslemGunlugu.Range("B7:B100000").Find(What:=WsRapor.Cells(IlkSira, 165).Value, SearchDirection:=xlNext, _
                SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole) 'zaman damgasını ara
Set IslemGunluguSonSiraBul = WsIslemGunlugu.Range("C7:C100000").Find(What:=WsRapor.Cells(IlkSira, 165).Value, SearchDirection:=xlNext, _
                SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
If Not IslemGunluguIlkSiraBul Is Nothing Then
    IslemGunluguIlkSira = IslemGunluguIlkSiraBul.Row
    If Not IslemGunluguSonSiraBul Is Nothing Then
        IslemGunluguSonSira = IslemGunluguSonSiraBul.Row
    End If
End If

    
If Not IslemGunluguIlkSiraBul Is Nothing Then 'DÜZENLEME İŞLEMİ

    '_______________'TipA/TipB (Daha önce tipA olarak kaydettiği bir işlemi tipBye çevirirse işlem günlüğündeki kaydın silinmesi gerekir.)
    
    If WsRapor.Cells(IlkSira, 28).Value <> "Type A" Then
        
        'kayıt def. verileri sil, satırları işaretle
        WsIslemGunlugu.Range(WsIslemGunlugu.Cells(IslemGunluguIlkSira, 2), WsIslemGunlugu.Cells(IslemGunluguSonSira, 20)).ClearContents
        WsIslemGunlugu.Cells(IslemGunluguIlkSira, 2).Value = "Sil" 'ilk satırı silmek üzere işaretle
        WsIslemGunlugu.Cells(IslemGunluguSonSira, 3).Value = "Sil" 'son satırı silmek üzere işaretle
        
        'Dönem sıra no.ları güncelle
        i = IslemGunluguSonSira
        Do Until WsIslemGunlugu.Cells(i, 5).Value <> "" 'silinecek verinin dönemi en alt satırda değilse stop koşulu
            i = i + 1
            If i > Say2IslemGunlugu Then 'silinecek verinin dönemi en alt satırda ise stop koşulu
                GoTo SilDonemSiraNo
            End If
            If WsIslemGunlugu.Cells(i, 6).Value <> "" And IsNumeric(WsIslemGunlugu.Cells(i, 6).Value) Then 'silinen veriden sonraki verileri dönem sıra no.ları 1 azalır
                WsIslemGunlugu.Cells(i, 6).Value = WsIslemGunlugu.Cells(i, 6).Value - 1
            End If
        Loop
SilDonemSiraNo:
    
        'Genel sıra no.ları güncelle
        SayGenel = WsIslemGunlugu.Range("D100000").End(xlUp).Row
        i = IslemGunluguSonSira
        Do Until i > SayGenel
            i = i + 1
            If WsIslemGunlugu.Cells(i, 4).Value <> "" Then
                WsIslemGunlugu.Cells(i, 4).Value = WsIslemGunlugu.Cells(i, 4).Value - 1
            End If
        Loop
        
        Say2IslemGunlugu = WsIslemGunlugu.Range("C100000").End(xlUp).Row
        If IslemGunluguIlkSira > 8 Then 'And IslemGunluguIlkSira < Say2IslemGunlugu Then
            Set Kenarlar = WsIslemGunlugu.Range("D" & IslemGunluguIlkSira - 1 & ":T" & IslemGunluguIlkSira - 1)
            With Kenarlar.Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .Color = RGB(174, 185, 194)
                .Weight = xlThin
            End With
            With Kenarlar.Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .Color = RGB(174, 185, 194)
                .Weight = xlThin
            End With
            With Kenarlar.Borders(xlInsideVertical)
                .LineStyle = xlContinuous
                .Color = RGB(174, 185, 194)
                .Weight = xlThin
            End With
            With Kenarlar.Borders(xlInsideHorizontal)
                .LineStyle = xlContinuous
                .Color = RGB(174, 185, 194)
                .Weight = xlThin
            End With
        End If
    
        'Silinecek dönemde yer alan boş satır aralığını kaldır
        Set BulIslemGunlugu = WsIslemGunlugu.Range("B:B").Find(What:="Sil", SearchDirection:=xlNext, _
        SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
        If Not BulIslemGunlugu Is Nothing Then
            ilkrowx = BulIslemGunlugu.Row
            Set BulIslemGunlugu = WsIslemGunlugu.Range("C:C").Find(What:="Sil", SearchDirection:=xlNext, _
            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not BulIslemGunlugu Is Nothing Then
                sonrowx = BulIslemGunlugu.Row
            End If
            WsIslemGunlugu.Rows(ilkrowx & ":" & sonrowx).EntireRow.Delete
        End If

        
        'İşlem günlüğünde aşağı git
        Say2IslemGunlugu = WsIslemGunlugu.Range("C100000").End(xlUp).Row
        On Error Resume Next
        ActiveWindow.ScrollRow = Say2IslemGunlugu - 10
        On Error GoTo 0
    
        WsIslemGunlugu.Columns("B:C").EntireColumn.Hidden = True
    
        WsIslemGunlugu.Protect Password:="123"
    
        'İşlem günlüğü açıksa kaydet ve kapat.
        OpenControl = IsWorkBookOpen(IslemGunlugu)
        If OpenControl = True Then
            Workbooks("System Registry Report 2.1.xlsx").Save
            Workbooks("System Registry Report 2.1.xlsx").Close SaveChanges:=False
        End If

        If IlceSakla <> "" Then
            WsRapor.Cells(IlkSira, 20).Value = IlceSakla
        End If

        GoTo Out
        
    End If

    '_________________________'TipA/TipB BİTİŞ
    
    
    'DÖNEM AYNI mı FARKLI mı?
    i = IslemGunluguIlkSiraBul.Row
    Do Until WsIslemGunlugu.Cells(i, 5).Value <> "" 'CDate(ModulAyrac)
        i = i - 1
    Loop
    
    If WsIslemGunlugu.Cells(i, 5).Formula = CDate(ModulAyrac) Then
        'MsgBox "Aynı dönem"
        GoTo DonemAyni
    Else
        'MsgBox "Farklı dönem"
        GoTo DonemFarkli
    End If

DonemFarkli:
    '_______________FARKLI DÖNEM (DÜZENLEME İŞLEMİ)

    'Önceki dönemde bulunan veriyi sil.
    WsIslemGunlugu.Range(WsIslemGunlugu.Cells(IslemGunluguIlkSira, 2), WsIslemGunlugu.Cells(IslemGunluguSonSira, 20)).ClearContents
    WsIslemGunlugu.Cells(IslemGunluguIlkSira, 2).Value = "Sil" 'ilk satırı silmek üzere işaretle
    WsIslemGunlugu.Cells(IslemGunluguSonSira, 3).Value = "Sil" 'son satırı silmek üzere işaretle
        
    'Kaynak dönemde bulunan dönem sıra no.ları güncelle
    i = IslemGunluguSonSira
    Do Until WsIslemGunlugu.Cells(i, 5).Value <> "" 'kaynak dönem en alt satırda değilse stop koşulu
        i = i + 1
        If i > Say2IslemGunlugu Then 'kaynak dönem en alt satırda ise stop koşulu
            GoTo KaynakDonemSiraNo
        End If
        If WsIslemGunlugu.Cells(i, 6).Value <> "" And IsNumeric(WsIslemGunlugu.Cells(i, 6).Value) Then 'silinen veriden sonraki verileri dönem sıra no.ları 1 azalır
            WsIslemGunlugu.Cells(i, 6).Value = WsIslemGunlugu.Cells(i, 6).Value - 1
        End If
    Loop
KaynakDonemSiraNo:

    
    Set BulIslemGunlugu = WsIslemGunlugu.Range("E:E").Find(What:=CDate(ModulAyrac), SearchDirection:=xlNext, _
            SearchOrder:=xlByRows, LookIn:=xlFormulas, LookAt:=xlWhole)
    If Not BulIslemGunlugu Is Nothing Then 'HEDEF DÖNEMİ bul
        
        'Hedef dönemin en alt satırı
        i = BulIslemGunlugu.Row + 1
        j = 0
        Do Until WsIslemGunlugu.Cells(i, 5).Value <> "" 'Hedef dönem en alt satırda değilse stop koşulu
            i = i + 1
            If i > Say2IslemGunlugu Then 'Hedef dönem en alt satırda ise stop koşulu
                i = i - 1
                j = 1
                GoTo HedefDonemAltSatir
            End If
        Loop
HedefDonemAltSatir:
YeniDonemAltRow = i

        If j = 1 Then
            ilkrow = YeniDonemAltRow
            sonrow = YeniDonemAltRow + (SonSira - IlkSira)
        Else
            For i = 1 To (SonSira - IlkSira) + 1 'Taşınacak satır aralığı kadar yeni dönemin en altına satır ekle
                WsIslemGunlugu.Rows(YeniDonemAltRow).EntireRow.Insert Shift:=xlUp
            Next i

            ilkrow = YeniDonemAltRow
            sonrow = YeniDonemAltRow + (SonSira - IlkSira)
            
        End If

        'Genel sıra no.ları güncelle
        WsIslemGunlugu.Cells(ilkrow, 4).Value = 1 'Genel sıra no.sunu 1 olarak işaretle (aşağıda, doğru no. ile değiştirilecek)
        SayGenel = WsIslemGunlugu.Range("D100000").End(xlUp).Row
        i = 6
        j = 0
        Do Until i > SayGenel
            i = i + 1
            If WsIslemGunlugu.Cells(i, 4).Value <> "" Then
                j = j + 1
                WsIslemGunlugu.Cells(i, 4).Value = j
            End If
        Loop
  
        'Hedef dönem sıra no.ları güncelle
        i = ilkrow
        Do Until WsIslemGunlugu.Cells(i, 5).Value <> ""
            i = i - 1
            If WsIslemGunlugu.Cells(i, 5).Value <> "" And Not IsNumeric(WsIslemGunlugu.Cells(i, 5).Value) Then
                WsIslemGunlugu.Cells(ilkrow, 6).Value = 1
                GoTo HedefDonemSiraNo
            End If
            If WsIslemGunlugu.Cells(i, 6).Value <> "" And IsNumeric(WsIslemGunlugu.Cells(i, 6).Value) Then
                WsIslemGunlugu.Cells(ilkrow, 6).Value = WsIslemGunlugu.Cells(i, 6).Value + 1
                GoTo HedefDonemSiraNo
            End If
        Loop
HedefDonemSiraNo:
        
        
        'Kaynak dönemde yer alan boş satır aralığını kaldır
        DelControl = True

    End If
    

    GoTo DonemAyniyiAtla

DonemAyni:
    '_______________AYNI DÖNEM (DÜZENLEME İŞLEMİ)

    If Not IslemGunluguSonSiraBul Is Nothing Then
        'MsgBox "Buradayım"
        'Aktarımları yapan kodlar buraya gelecek
        WsIslemGunlugu.Range(WsIslemGunlugu.Cells(IslemGunluguIlkSira, 7), WsIslemGunlugu.Cells(IslemGunluguSonSira, 20)).ClearContents
        WsIslemGunlugu.Range(WsIslemGunlugu.Cells(IslemGunluguIlkSira, 2), WsIslemGunlugu.Cells(IslemGunluguSonSira, 3)).ClearContents
        Fark = (IslemGunluguSonSira - IslemGunluguIlkSira) - (SonSira - IlkSira)
        'MsgBox "Fark: " & Fark
        If Fark > 0 Then 'İşlem günlüğünden satır silinecek
            'MsgBox "Fark: " & Fark & " satır kaldır"
            WsIslemGunlugu.Rows(IslemGunluguSonSira - (Fark - 1) & ":" & IslemGunluguSonSira).EntireRow.Delete
            ilkrow = IslemGunluguIlkSira
            sonrow = IslemGunluguSonSira - Fark
        ElseIf Fark < 0 Then 'İşlem günlüğüne satır eklenecek
            'MsgBox "Fark: " & Fark & " satır ekle"
            Fark = -1 * Fark
            For i = 1 To Fark
                WsIslemGunlugu.Rows(IslemGunluguSonSira + 1).EntireRow.Insert Shift:=xlUp
            Next i
            ilkrow = IslemGunluguIlkSira
            sonrow = IslemGunluguSonSira + Fark
        ElseIf Fark = 0 Then 'İşlem günlüğünde satırlarda değişiklik olmayacak
            'MsgBox "Fark: " & Fark & " değişiklik yok"
            ilkrow = IslemGunluguIlkSira
            sonrow = IslemGunluguSonSira
        End If

    End If

Else 'YENİ İŞLEM

    'MsgBox CDate(ModulAyrac)
    Set BulIslemGunlugu = WsIslemGunlugu.Range("E:E").Find(What:=CDate(ModulAyrac), SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlFormulas, LookAt:=xlWhole)
    If Not BulIslemGunlugu Is Nothing Then
        donemrow = BulIslemGunlugu.Row
        If SayAyracIslemGunlugu = BulIslemGunlugu.Row Then 'Cari dönemin verisi
            If Say2IslemGunlugu > SayAyracIslemGunlugu Then 'Yeni veriyi Say2IslemGunlugu+1'e yaz
                'Genel sıra no
                SayGenel = WsIslemGunlugu.Range("D100000").End(xlUp).Row
                If SayGenel < 7 Then
                    WsIslemGunlugu.Cells(Say2IslemGunlugu + 1, 4).Value = 1
                Else
                    i = Say2IslemGunlugu + 1
                    Do Until WsIslemGunlugu.Cells(i, 4).Value <> ""
                        i = i - 1
                    Loop
                    WsIslemGunlugu.Cells(Say2IslemGunlugu + 1, 4).Value = WsIslemGunlugu.Cells(i, 4).Value + 1
                End If
                'Dönem sıra no
                SayDonem = WsIslemGunlugu.Range("F100000").End(xlUp).Row
                If SayDonem < 7 Then
                    WsIslemGunlugu.Cells(Say2IslemGunlugu + 1, 6).Value = 1
                Else
                    i = Say2IslemGunlugu + 1
                    Do Until WsIslemGunlugu.Cells(i, 6).Value <> ""
                        i = i - 1
                    Loop
                    WsIslemGunlugu.Cells(Say2IslemGunlugu + 1, 6).Value = WsIslemGunlugu.Cells(i, 6).Value + 1
                End If
 
                ilkrow = Say2IslemGunlugu + 1
                sonrow = Say2IslemGunlugu + 1 + (SonSira - IlkSira)
                
            Else 'Yeni veriyi SayAyracIslemGunlugu+1'e yaz
                'Genel sıra no
                SayGenel = WsIslemGunlugu.Range("D100000").End(xlUp).Row
                If SayGenel < 7 Then
                    WsIslemGunlugu.Cells(SayAyracIslemGunlugu + 1, 4).Value = 1
                Else
                    i = SayAyracIslemGunlugu + 1
                    Do Until WsIslemGunlugu.Cells(i, 4).Value <> ""
                        i = i - 1
                    Loop
                    WsIslemGunlugu.Cells(SayAyracIslemGunlugu + 1, 4).Value = WsIslemGunlugu.Cells(i, 4).Value + 1
                End If
                'Dönem sıra no
                SayDonem = WsIslemGunlugu.Range("E100000").End(xlUp).Row
                WsIslemGunlugu.Cells(SayAyracIslemGunlugu + 1, 6).Value = 1

                ilkrow = SayAyracIslemGunlugu + 1
                sonrow = SayAyracIslemGunlugu + 1 + (SonSira - IlkSira)
                
            End If
        Else 'Cari dönemden önceki dönemin verisi
            'MsgBox "Buradayım"
            ModulAyrac = DateAdd("m", 1, CDate(ModulAyrac)) 'Sonraki dönemi bul ve onun üstüne satır ekle
            Set BulIslemGunlugu = WsIslemGunlugu.Range("E:E").Find(What:=CDate(ModulAyrac), SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlFormulas, LookAt:=xlWhole)
            If Not BulIslemGunlugu Is Nothing Then
                For i = 1 To (SonSira - IlkSira) + 1
                    WsIslemGunlugu.Rows(BulIslemGunlugu.Row).EntireRow.Insert Shift:=xlUp
                Next i
'                WsIslemGunlugu.Range("E" & BulIslemGunlugu.Row - 1).Value = "İkincisi"
'                WsIslemGunlugu.Range("E" & BulIslemGunlugu.Row - 2).Value = "İlki"

                ilkrow = BulIslemGunlugu.Row - (SonSira - IlkSira + 1)
                sonrow = BulIslemGunlugu.Row - 1
                
                'Genel sıra no
                SayGenel = WsIslemGunlugu.Range("D100000").End(xlUp).Row
                If SayGenel < 7 Then
                    WsIslemGunlugu.Cells(ilkrow, 4).Value = 1
                    If SayGenel > ilkrow Then 'ilkrow dan sonra gelen sıra no.ları düzelt
                        For i = ilkrow + 1 To SayGenel
                            If WsIslemGunlugu.Cells(i, 4).Value <> "" Then
                                WsIslemGunlugu.Cells(i, 4).Value = WsIslemGunlugu.Cells(i, 4).Value + 1
                            End If
                        Next i
                    End If
                Else
                    i = ilkrow
                    Do Until WsIslemGunlugu.Cells(i, 4).Value <> ""
                        i = i - 1
                        If WsIslemGunlugu.Cells(i, 4).Value <> "" And Not IsNumeric(WsIslemGunlugu.Cells(i, 4).Value) Then
                            WsIslemGunlugu.Cells(ilkrow, 4).Value = 1
                            GoTo LoopSon1
                        End If
                    Loop
                    WsIslemGunlugu.Cells(ilkrow, 4).Value = WsIslemGunlugu.Cells(i, 4).Value + 1
LoopSon1:
                    If SayGenel > ilkrow Then 'ilkrow dan sonra gelen sıra no.ları düzelt
                        For i = ilkrow + 1 To SayGenel
                            If WsIslemGunlugu.Cells(i, 4).Value <> "" Then
                                WsIslemGunlugu.Cells(i, 4).Value = WsIslemGunlugu.Cells(i, 4).Value + 1
                            End If
                        Next i
                    End If
                End If
                
                'Dönem sıra no
                'SayDonem = WsIslemGunlugu.Range("E100000").End(xlUp).Row
                i = ilkrow
                Do Until i < donemrow
                    i = i - 1
                    If WsIslemGunlugu.Cells(i, 5).Value <> "" And Not IsNumeric(WsIslemGunlugu.Cells(i, 5).Value) Then
                        WsIslemGunlugu.Cells(ilkrow, 6).Value = 1
                        GoTo LoopSon2
                    End If
                    If WsIslemGunlugu.Cells(i, 6).Value <> "" And IsNumeric(WsIslemGunlugu.Cells(i, 6).Value) Then
                        WsIslemGunlugu.Cells(ilkrow, 6).Value = WsIslemGunlugu.Cells(i, 6).Value + 1
                        GoTo LoopSon2
                    End If
                Loop
LoopSon2:

            End If
            ModulAyrac = DateAdd("m", -1, CDate(ModulAyrac)) 'Yukarıda atadığın +1 ayı geri al
        End If
    Else
        'İşlem günlüğünde kayıtlı en eski DÖNEMDEN daha eski bir döneme ilişkin veri girişi gerçekleşmesi durumu.
    '    MsgBox "Outbound Qty"
        GoTo Son
    End If
End If

DonemAyniyiAtla:


'Dolguları kaldır/Biçimlendirmeleri düzelt
WsIslemGunlugu.Range("E" & ilkrow & ":T" & sonrow).Interior.Color = xlNone
WsIslemGunlugu.Range("E" & ilkrow & ":T" & sonrow).Font.Color = RGB(0, 0, 0)
WsIslemGunlugu.Range("E" & ilkrow & ":T" & sonrow).Font.Bold = False
WsIslemGunlugu.Range("B" & ilkrow & ":T" & sonrow).NumberFormat = "@"
WsIslemGunlugu.Range("P" & ilkrow & ":Q" & sonrow).NumberFormat = "General"
WsIslemGunlugu.Range("F" & ilkrow & ":F" & sonrow).NumberFormat = "General"
WsIslemGunlugu.Range("D" & ilkrow & ":D" & sonrow).NumberFormat = "General"
WsIslemGunlugu.Range("B" & ilkrow & ":T" & sonrow).WrapText = True
WsIslemGunlugu.Range("F" & ilkrow & ":G" & sonrow).VerticalAlignment = xlCenter
WsIslemGunlugu.Range("P" & ilkrow & ":Q" & sonrow).VerticalAlignment = xlCenter
WsIslemGunlugu.Range("F" & ilkrow & ":G" & sonrow).HorizontalAlignment = xlCenter
WsIslemGunlugu.Range("P" & ilkrow & ":Q" & sonrow).HorizontalAlignment = xlCenter


'Zaman damgaları
WsIslemGunlugu.Cells(ilkrow, 2).Value = WsRapor.Cells(IlkSira, 165).Value
WsIslemGunlugu.Cells(sonrow, 3).Value = WsRapor.Cells(IlkSira, 165).Value
'Verileri yaz
WsIslemGunlugu.Range(WsIslemGunlugu.Cells(ilkrow, 7), WsIslemGunlugu.Cells(sonrow, 7)).Value = WsRapor.Range(WsRapor.Cells(IlkSira, 13), WsRapor.Cells(SonSira, 13)).Value 'Rapor no
WsIslemGunlugu.Cells(ilkrow, 8).Value = WsRapor.Cells(IlkSira, 19).Value 'İl
WsIslemGunlugu.Cells(ilkrow, 9).Value = WsRapor.Cells(IlkSira, 20).Value 'İlçe
WsIslemGunlugu.Cells(ilkrow, 10).Value = GelenTema
WsIslemGunlugu.Cells(ilkrow, 11).Value = WsRapor.Cells(IlkSira, 23).Value 'Belge tarihi
WsIslemGunlugu.Cells(ilkrow, 12).Value = "" 'Belge no
WsIslemGunlugu.Cells(ilkrow, 13).Value = WsRapor.Cells(IlkSira, 23).Value 'finansal birime ulaşma tarihi
WsIslemGunlugu.Cells(ilkrow, 14).Value = WsRapor.Cells(IlkSira, 23).Value 'Tespit tarihi
WsIslemGunlugu.Range(WsIslemGunlugu.Cells(ilkrow, 15), WsIslemGunlugu.Cells(sonrow, 15)).Value = WsRapor.Range(WsRapor.Cells(IlkSira, 52), WsRapor.Cells(SonSira, 52)).Value 'Öğe türü
WsIslemGunlugu.Range(WsIslemGunlugu.Cells(ilkrow, 16), WsIslemGunlugu.Cells(sonrow, 16)).Value = WsRapor.Range(WsRapor.Cells(IlkSira, 55), WsRapor.Cells(SonSira, 55)).Value 'Öğe değeri
WsIslemGunlugu.Range(WsIslemGunlugu.Cells(ilkrow, 17), WsIslemGunlugu.Cells(sonrow, 17)).Value = WsRapor.Range(WsRapor.Cells(IlkSira, 58), WsRapor.Cells(SonSira, 58)).Value 'Adet
WsIslemGunlugu.Cells(ilkrow, 18).Value = WsRapor.Cells(IlkSira, 26).Value 'Tema
WsIslemGunlugu.Range(WsIslemGunlugu.Cells(ilkrow, 19), WsIslemGunlugu.Cells(sonrow, 19)).Value = WsRapor.Range(WsRapor.Cells(IlkSira, 64), WsRapor.Cells(SonSira, 64)).Value 'Açıklama
WsIslemGunlugu.Range(WsIslemGunlugu.Cells(ilkrow, 20), WsIslemGunlugu.Cells(sonrow, 20)).Value = WsRapor.Range(WsRapor.Cells(IlkSira, 213), WsRapor.Cells(SonSira, 213)).Value 'Baskı tekniği


'Kenarlıklar.
Set Kenarlar = WsIslemGunlugu.Range("D" & ilkrow & ":T" & sonrow)
With Kenarlar.Borders(xlEdgeTop)
    .LineStyle = xlContinuous
    .Color = RGB(174, 185, 194)
    .Weight = xlThin
End With
With Kenarlar.Borders(xlEdgeBottom)
    .LineStyle = xlContinuous
    .Color = RGB(174, 185, 194)
    .Weight = xlThin
End With
With Kenarlar.Borders(xlInsideVertical)
    .LineStyle = xlContinuous
    .Color = RGB(174, 185, 194)
    .Weight = xlThin
End With
With Kenarlar.Borders(xlInsideHorizontal)
    .LineStyle = xlContinuous
    .Color = RGB(174, 185, 194)
    .Weight = xlThin
End With

If DelControl = True Then
    'Kaynak dönemde yer alan boş satır aralığını kaldır
    Set BulIslemGunlugu = WsIslemGunlugu.Range("B:B").Find(What:="Sil", SearchDirection:=xlNext, _
    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    If Not BulIslemGunlugu Is Nothing Then
        ilkrowx = BulIslemGunlugu.Row
        Set BulIslemGunlugu = WsIslemGunlugu.Range("C:C").Find(What:="Sil", SearchDirection:=xlNext, _
        SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
        If Not BulIslemGunlugu Is Nothing Then
            sonrowx = BulIslemGunlugu.Row
        End If
        WsIslemGunlugu.Rows(ilkrowx & ":" & sonrowx).EntireRow.Delete
    End If
End If


Son:

'İşlem günlüğünde aşağı git
Say2IslemGunlugu = WsIslemGunlugu.Range("C100000").End(xlUp).Row
On Error Resume Next
ActiveWindow.ScrollRow = Say2IslemGunlugu - 10
On Error GoTo 0

WsIslemGunlugu.Columns("B:C").EntireColumn.Hidden = True

WsIslemGunlugu.Protect Password:="123"

'İşlem günlüğü açıksa kaydet ve kapat.
OpenControl = IsWorkBookOpen(IslemGunlugu)
If OpenControl = True Then
    Workbooks("System Registry Report 2.1.xlsx").Save
    Workbooks("System Registry Report 2.1.xlsx").Close SaveChanges:=False
End If


If IlceSakla <> "" Then
    WsRapor.Cells(IlkSira, 20).Value = IlceSakla
End If


'WsRapor.Protect Password:="123"

Out:

Application.DisplayAlerts = True
Application.EnableEvents = True
Application.ScreenUpdating = True


End Sub

Sub Rapor3_2TeslimTutanaklari()
Dim i As Integer, TarihBul As Range
Dim Say As Long, j As Long, Cont As Long, Sno As Variant
Dim LstBx As MSForms.ListBox, ctl As MSForms.Control
Dim LblSira1 As MSForms.Label
Dim LblSira2 As MSForms.Label
Dim Contx As Long

CalTarih = CalTarihTakip

Say = ThisWorkbook.Worksheets(5).Range("FH100000").End(xlUp).Row
If Say < 7 Then
    GoTo Son
End If

'Cont = 0
Contx = 0
TarihiTekrarla1:
Set TarihBul = ThisWorkbook.Worksheets(5).Range("CE6:CE100000").Find(What:=CalTarih, SearchDirection:=xlNext, _
                SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
If Not TarihBul Is Nothing Then
    'Cells(TarihBul.Row, 31).Value
Else
    CalTarih = CDate(CalTarih)
    CalTarih = DateAdd("d", 1, CalTarih)
    CalTarih = CStr(CalTarih)
    If Mid(CalTarih, 2, 1) = "." Then 'Günün soluna 0 ekle
        CalTarih = "0" & CalTarih
    End If
    If Mid(CalTarih, 5, 1) = "." Then 'Ayın soluna 0 ekle
        CalTarih = Left(CalTarih, 3) & "0" & Mid(CalTarih, 4, 6)
    End If
    'MsgBox CalTarih
    Contx = Contx + 1
    If Contx = 100 Then
        'MsgBox "Belirtilen tarihten sonra herhangi bir tutanak1 işlemi yapılmadığından işleminiz gerçekleştirilemiyor.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    GoTo TarihiTekrarla1
End If


Cont = ContTakip
'Cont = 0
For j = Say To TarihBul.Row Step -1
    'If Cont < 50 Then
        If ThisWorkbook.Worksheets(5).Cells(j, 5).Value <> "" And ThisWorkbook.Worksheets(5).Cells(j, 83).Value <> "" Then
            Cont = Cont + 1
            If ThisWorkbook.Worksheets(5).Cells(j, 48).Value <> "" Then
                If Left(ThisWorkbook.Worksheets(5).Cells(j, 47).Value, 3) = "İl " Then
                    Sno = ThisWorkbook.Worksheets(5).Cells(j, 5).Value & " | M | " & ThisWorkbook.Worksheets(5).Cells(j, 84).Value & " | " & ThisWorkbook.Worksheets(5).Cells(j, 48).Value
                ElseIf Left(ThisWorkbook.Worksheets(5).Cells(j, 47).Value, 3) = "İlç" Then
                    Sno = ThisWorkbook.Worksheets(5).Cells(j, 5).Value & " | M | " & ThisWorkbook.Worksheets(5).Cells(j, 84).Value & " | " & ThisWorkbook.Worksheets(5).Cells(j, 48).Value
                Else
                    Sno = ThisWorkbook.Worksheets(5).Cells(j, 5).Value & " | M | " & ThisWorkbook.Worksheets(5).Cells(j, 84).Value & " | " & ThisWorkbook.Worksheets(5).Cells(j, 47).Value & " " & ThisWorkbook.Worksheets(5).Cells(j, 48).Value
                End If
            ElseIf ThisWorkbook.Worksheets(5).Cells(j, 48).Value = "" Then
                If Left(ThisWorkbook.Worksheets(5).Cells(j, 47).Value, 3) = "İl " Then
                    Sno = ThisWorkbook.Worksheets(5).Cells(j, 5).Value & " | M | " & ThisWorkbook.Worksheets(5).Cells(j, 84).Value & " | " & ThisWorkbook.Worksheets(5).Cells(j, 19).Value & " " & ThisWorkbook.Worksheets(5).Cells(j, 47).Value
                ElseIf Left(ThisWorkbook.Worksheets(5).Cells(j, 47).Value, 3) = "İlç" Then
                    Sno = ThisWorkbook.Worksheets(5).Cells(j, 5).Value & " | M | " & ThisWorkbook.Worksheets(5).Cells(j, 84).Value & " | " & ThisWorkbook.Worksheets(5).Cells(j, 20).Value & " " & ThisWorkbook.Worksheets(5).Cells(j, 47).Value
                Else
                    Sno = ThisWorkbook.Worksheets(5).Cells(j, 5).Value & " | M | " & ThisWorkbook.Worksheets(5).Cells(j, 84).Value & " | " & ThisWorkbook.Worksheets(5).Cells(j, 47).Value
                End If
            End If

            Set LstBx = core_delivery_manager_UI.Frame1.Controls.Add("Forms.ListBox.1", "Frame1LstBx" & Cont)
            With LstBx
                .Top = (Cont - 1) * 12
                .Left = 18
                .Height = 12
                .Width = 300
                .SpecialEffect = fmSpecialEffectFlat
                .MultiSelect = fmMultiSelectMulti
                .TextAlign = fmTextAlignLeft '02092021
                .AddItem (Sno)
            End With

            Set LblSira1 = core_delivery_manager_UI.Frame1.Controls.Add("Forms.Label.1", "Lbl" & Cont)
            With LblSira1
                .Top = (Cont - 1) * 12
                .Left = 0
                .Height = 12
                .Width = 18
                .SpecialEffect = fmSpecialEffectEtched
                .TextAlign = fmTextAlignCenter
                .Caption = Cont
            End With

            ScrollTakip1 = ScrollTakip1 + 12
        End If
    'End If
Next j

ContTakip = Cont

Son:

'MsgBox "Rapor3_2: " & ContTakip

End Sub

Sub Rapor3_2TeslimTutanaklariFinansalBirim()
Dim i As Integer, TarihBul As Range
Dim Say As Long, j As Long, Cont As Long, Sno As Variant
Dim LstBx As MSForms.ListBox, ctl As MSForms.Control
Dim LblSira1 As MSForms.Label
Dim LblSira2 As MSForms.Label
Dim Contx As Long

CalTarih = CalTarihTakip

Say = ThisWorkbook.Worksheets(5).Range("FH100000").End(xlUp).Row
If Say < 7 Then
    GoTo Son
End If

'Cont = 0
Contx = 0
TarihiTekrarla1:
Set TarihBul = ThisWorkbook.Worksheets(5).Range("BW6:BW100000").Find(What:=CalTarih, SearchDirection:=xlNext, _
                SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
If Not TarihBul Is Nothing Then
    'Cells(TarihBul.Row, 31).Value
Else
    CalTarih = CDate(CalTarih)
    CalTarih = DateAdd("d", 1, CalTarih)
    CalTarih = CStr(CalTarih)
    If Mid(CalTarih, 2, 1) = "." Then 'Günün soluna 0 ekle
        CalTarih = "0" & CalTarih
    End If
    If Mid(CalTarih, 5, 1) = "." Then 'Ayın soluna 0 ekle
        CalTarih = Left(CalTarih, 3) & "0" & Mid(CalTarih, 4, 6)
    End If
    'MsgBox CalTarih
    Contx = Contx + 1
    If Contx = 100 Then
        'MsgBox "Belirtilen tarihten sonra herhangi bir tutanak1 işlemi yapılmadığından işleminiz gerçekleştirilemiyor.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    GoTo TarihiTekrarla1
End If


Cont = ContTakip
'Cont = 0
For j = Say To TarihBul.Row Step -1
    'If Cont < 50 Then
        If ThisWorkbook.Worksheets(5).Cells(j, 5).Value <> "" And ThisWorkbook.Worksheets(5).Cells(j, 75).Value <> "" Then
            Cont = Cont + 1
            If ThisWorkbook.Worksheets(5).Cells(j, 30).Value <> "" Then
                If ThisWorkbook.Worksheets(5).Cells(j, 28).Value = "Type A" Then
                    Sno = ThisWorkbook.Worksheets(5).Cells(j, 5).Value & " | M | " & ThisWorkbook.Worksheets(5).Cells(j, 76).Value & " | " & ThisWorkbook.Worksheets(5).Cells(j, 30).Value & " (FinansalBirim-TipA)"
                ElseIf ThisWorkbook.Worksheets(5).Cells(j, 28).Value = "Type B" Then
                    Sno = ThisWorkbook.Worksheets(5).Cells(j, 5).Value & " | M | " & ThisWorkbook.Worksheets(5).Cells(j, 76).Value & " | " & ThisWorkbook.Worksheets(5).Cells(j, 30).Value & " (FinansalBirim-TipB)"
                End If
            End If

            Set LstBx = core_delivery_manager_UI.Frame1.Controls.Add("Forms.ListBox.1", "Frame1LstBx" & Cont)
            With LstBx
                .Top = (Cont - 1) * 12
                .Left = 18
                .Height = 12
                .Width = 300
                .SpecialEffect = fmSpecialEffectFlat
                .MultiSelect = fmMultiSelectMulti
                .TextAlign = fmTextAlignLeft '02092021
                .AddItem (Sno)
            End With

            Set LblSira1 = core_delivery_manager_UI.Frame1.Controls.Add("Forms.Label.1", "Lbl" & Cont)
            With LblSira1
                .Top = (Cont - 1) * 12
                .Left = 0
                .Height = 12
                .Width = 18
                .SpecialEffect = fmSpecialEffectEtched
                .TextAlign = fmTextAlignCenter
                .Caption = Cont
            End With

            ScrollTakip1 = ScrollTakip1 + 12
        End If
    'End If
Next j

ContTakip = Cont

Son:

'MsgBox "Rapor3_2: " & ContTakip

End Sub

Sub Rapor3_2VarlikHareketleriGiris()
Dim i As Integer, TarihBul As Range
Dim Say As Long, j As Long, Cont As Long, Sno As Variant
Dim LstBx As MSForms.ListBox, ctl As MSForms.Control
Dim LblSira1 As MSForms.Label
Dim LblSira2 As MSForms.Label
Dim Contx As Long

Dim IlkSiraBul As Range, SonSiraBul As Range, IlkSira As Long, SonSira As Long, k As Long
Dim RaporNoTakip As String, RaporNoKontrol As Integer, SiraNo As Long

CalTarih = CalTarihTakip

Say = ThisWorkbook.Worksheets(5).Range("FH100000").End(xlUp).Row
If Say < 7 Then
    GoTo Son
End If

'Cont = 0
Contx = 0
Set TarihBul = ThisWorkbook.Worksheets(5).Range("W6:W100000").Find(What:=CalTarih, SearchDirection:=xlNext, _
                SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole) 'Tutanak tarihi içinde ara
If Not TarihBul Is Nothing Then
    '
Else
    GoTo Son
End If

'MsgBox CalTarih

Cont = ContTakipGiris
'Cont = 0
For j = TarihBul.Row To Say
    If ThisWorkbook.Worksheets(5).Cells(j, 12).Value = "Point2" Or ThisWorkbook.Worksheets(5).Cells(j, 12).Value = "Point3" Then
        If ThisWorkbook.Worksheets(5).Cells(j, 28).Value = "Type A" Then
            If CalTarih = ThisWorkbook.Worksheets(5).Cells(j, 23).Value And CDate(ThisWorkbook.Worksheets(5).Cells(j, 23).Value) <> CDate(ThisWorkbook.Worksheets(5).Cells(j, 176).Value) Then
                If ThisWorkbook.Worksheets(5).Cells(j, 5).Value <> "" And ThisWorkbook.Worksheets(5).Cells(j, 13).Value <> "" Then
                    Cont = Cont + 1
                    
                    Set IlkSiraBul = Nothing
                    Set SonSiraBul = Nothing
                    IlkSira = 0
                    SonSira = 0
                    SiraNo = ThisWorkbook.Worksheets(5).Cells(j, 5).Value
                    Set IlkSiraBul = ThisWorkbook.Worksheets(5).Range("FG7:FG100000").Find(What:=SiraNo, SearchDirection:=xlNext, _
                                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
                    If Not IlkSiraBul Is Nothing Then
                        IlkSira = IlkSiraBul.Row
                    End If
                    Set SonSiraBul = ThisWorkbook.Worksheets(5).Range("FH7:FH100000").Find(What:=SiraNo, SearchDirection:=xlNext, _
                                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
                    If Not SonSiraBul Is Nothing Then
                        SonSira = SonSiraBul.Row
                    End If
                    RaporNoKontrol = 0
                    For k = IlkSira To SonSira
                        If ThisWorkbook.Worksheets(5).Cells(k, 13).Value <> "" Then
                            If RaporNoKontrol = 0 Then
                                RaporNoTakip = ThisWorkbook.Worksheets(5).Cells(k, 13).Value
                                RaporNoKontrol = 1
                            Else
                                RaporNoTakip = RaporNoTakip & ", " & ThisWorkbook.Worksheets(5).Cells(k, 13).Value
                            End If
                        End If
                    Next k
                    
                    
                    If ThisWorkbook.Worksheets(5).Cells(j, 48).Value <> "" Then
                        If Left(ThisWorkbook.Worksheets(5).Cells(j, 47).Value, 3) = "İl " Then
                            Sno = ThisWorkbook.Worksheets(5).Cells(j, 5).Value & " | M | " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(5).Cells(j, 48).Value
                        ElseIf Left(ThisWorkbook.Worksheets(5).Cells(j, 47).Value, 3) = "İlç" Then
                            Sno = ThisWorkbook.Worksheets(5).Cells(j, 5).Value & " | M | " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(5).Cells(j, 48).Value
                        Else
                            Sno = ThisWorkbook.Worksheets(5).Cells(j, 5).Value & " | M | " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(5).Cells(j, 47).Value & " " & ThisWorkbook.Worksheets(5).Cells(j, 48).Value
                        End If
                    ElseIf ThisWorkbook.Worksheets(5).Cells(j, 48).Value = "" Then
                        If Left(ThisWorkbook.Worksheets(5).Cells(j, 47).Value, 3) = "İl " Then
                            Sno = ThisWorkbook.Worksheets(5).Cells(j, 5).Value & " | M | " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(5).Cells(j, 19).Value & " " & ThisWorkbook.Worksheets(5).Cells(j, 47).Value
                        ElseIf Left(ThisWorkbook.Worksheets(5).Cells(j, 47).Value, 3) = "İlç" Then
                            Sno = ThisWorkbook.Worksheets(5).Cells(j, 5).Value & " | M | " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(5).Cells(j, 20).Value & " " & ThisWorkbook.Worksheets(5).Cells(j, 47).Value
                        Else
                            Sno = ThisWorkbook.Worksheets(5).Cells(j, 5).Value & " | M | " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(5).Cells(j, 47).Value
                        End If
                    End If
    
                    Set LstBx = core_asset_manager_UI.FrameGiris.Controls.Add("Forms.ListBox.1", "Frame1LstBx" & Cont)
                    With LstBx
                        .Top = (Cont - 1) * 12
                        .Left = 18
                        .Height = 12
                        .Width = 300
                        .BackColor = &H80000000
                        .SpecialEffect = fmSpecialEffectFlat
                        .MultiSelect = fmMultiSelectMulti
                        .TextAlign = fmTextAlignLeft '02092021
                        .AddItem (Sno)
                    End With
        
                    Set LblSira1 = core_asset_manager_UI.FrameGiris.Controls.Add("Forms.Label.1", "Lbl" & Cont)
                    With LblSira1
                        .Top = (Cont - 1) * 12
                        .Left = 0
                        .Height = 12
                        .Width = 18
                        .SpecialEffect = fmSpecialEffectEtched
                        .TextAlign = fmTextAlignCenter
                        .Caption = Cont
                    End With
        
                    ScrollTakip1 = ScrollTakip1 + 12
                End If
            End If
        Else
            If CalTarih = ThisWorkbook.Worksheets(5).Cells(j, 23).Value And CDate(ThisWorkbook.Worksheets(5).Cells(j, 23).Value) <> CDate(ThisWorkbook.Worksheets(5).Cells(j, 176).Value) Then
                If ThisWorkbook.Worksheets(5).Cells(j, 5).Value <> "" Then 'And ThisWorkbook.Worksheets(5).Cells(j, 13).Value <> "" Then
                    Cont = Cont + 1
                    If ThisWorkbook.Worksheets(5).Cells(j, 48).Value <> "" Then
                        If Left(ThisWorkbook.Worksheets(5).Cells(j, 47).Value, 3) = "İl " Then
                            Sno = ThisWorkbook.Worksheets(5).Cells(j, 5).Value & " | M | " & ThisWorkbook.Worksheets(5).Cells(j, 26).Value & " | " & ThisWorkbook.Worksheets(5).Cells(j, 48).Value
                        ElseIf Left(ThisWorkbook.Worksheets(5).Cells(j, 47).Value, 3) = "İlç" Then
                            Sno = ThisWorkbook.Worksheets(5).Cells(j, 5).Value & " | M | " & ThisWorkbook.Worksheets(5).Cells(j, 26).Value & " | " & ThisWorkbook.Worksheets(5).Cells(j, 48).Value
                        Else
                            Sno = ThisWorkbook.Worksheets(5).Cells(j, 5).Value & " | M | " & ThisWorkbook.Worksheets(5).Cells(j, 26).Value & " | " & ThisWorkbook.Worksheets(5).Cells(j, 47).Value & " " & ThisWorkbook.Worksheets(5).Cells(j, 48).Value
                        End If
                    ElseIf ThisWorkbook.Worksheets(5).Cells(j, 48).Value = "" Then
                        If Left(ThisWorkbook.Worksheets(5).Cells(j, 47).Value, 3) = "İl " Then
                            Sno = ThisWorkbook.Worksheets(5).Cells(j, 5).Value & " | M | " & ThisWorkbook.Worksheets(5).Cells(j, 26).Value & " | " & ThisWorkbook.Worksheets(5).Cells(j, 19).Value & " " & ThisWorkbook.Worksheets(5).Cells(j, 47).Value
                        ElseIf Left(ThisWorkbook.Worksheets(5).Cells(j, 47).Value, 3) = "İlç" Then
                            Sno = ThisWorkbook.Worksheets(5).Cells(j, 5).Value & " | M | " & ThisWorkbook.Worksheets(5).Cells(j, 26).Value & " | " & ThisWorkbook.Worksheets(5).Cells(j, 20).Value & " " & ThisWorkbook.Worksheets(5).Cells(j, 47).Value
                        Else
                            Sno = ThisWorkbook.Worksheets(5).Cells(j, 5).Value & " | M | " & ThisWorkbook.Worksheets(5).Cells(j, 26).Value & " | " & ThisWorkbook.Worksheets(5).Cells(j, 47).Value
                        End If
                    End If
    
                    Set LstBx = core_asset_manager_UI.FrameGiris.Controls.Add("Forms.ListBox.1", "Frame1LstBx" & Cont)
                    With LstBx
                        .Top = (Cont - 1) * 12
                        .Left = 18
                        .Height = 12
                        .Width = 300
                        .BackColor = &H80000000
                        .SpecialEffect = fmSpecialEffectFlat
                        .MultiSelect = fmMultiSelectMulti
                        .TextAlign = fmTextAlignLeft '02092021
                        .AddItem (Sno)
                    End With
        
                    Set LblSira1 = core_asset_manager_UI.FrameGiris.Controls.Add("Forms.Label.1", "Lbl" & Cont)
                    With LblSira1
                        .Top = (Cont - 1) * 12
                        .Left = 0
                        .Height = 12
                        .Width = 18
                        .SpecialEffect = fmSpecialEffectEtched
                        .TextAlign = fmTextAlignCenter
                        .Caption = Cont
                    End With
        
                    ScrollTakip1 = ScrollTakip1 + 12
                End If
            End If
        End If
    End If
Next j

ContTakipGiris = Cont

Son:

End Sub

Sub Rapor3_2VarlikHareketleriMevcut()
Dim i As Integer, TarihBul As Range, SiraNoBul As Range
Dim Say As Long, j As Long, Cont As Long, Sno As Variant
Dim LstBx As MSForms.ListBox, ctl As MSForms.Control
Dim LblSira1 As MSForms.Label
Dim LblSira2 As MSForms.Label
Dim Contx As Long, MevcutSatir As Long
Dim TarihAra As String
Dim SayDevir As Long, TakipCount As Long

Dim IlkSiraBul As Range, SonSiraBul As Range, IlkSira As Long, SonSira As Long, k As Long
Dim RaporNoTakip As String, RaporNoKontrol As Integer, SiraNo As Long


CalTarih = CalTarihTakip

Say = ThisWorkbook.Worksheets(5).Range("FH100000").End(xlUp).Row
If Say < 7 Then
    GoTo Son
End If

'Mevcutta satırı bul
'CalTarih = CDate(CalTarih)
MevcutSatir = 0
For j = Say To 7 Step -1
    If ThisWorkbook.Worksheets(5).Cells(j, 12).Value = "Point2" Or ThisWorkbook.Worksheets(5).Cells(j, 12).Value = "Point3" Then
        If ThisWorkbook.Worksheets(5).Range("W" & j).Value <> "" Then
            TarihAra = ThisWorkbook.Worksheets(5).Range("W" & j)
            If CDate(TarihAra) < CDate(CalTarih) Then
                MevcutSatir = j
                'MsgBox Range("W" & MevcutSatir)
                GoTo DonguSonu
            End If
        End If
    End If
Next j
DonguSonu:
'Mevcutta satır bulunamaz ise sona git
If MevcutSatir = 0 Then
    GoTo Son
End If

Cont = ContTakipMevcut

For j = MevcutSatir To 7 Step -1
    If ThisWorkbook.Worksheets(5).Cells(j, 12).Value = "Point2" Or ThisWorkbook.Worksheets(5).Cells(j, 12).Value = "Point3" Then
        If ThisWorkbook.Worksheets(5).Cells(j, 28).Value = "Type A" Then
            If ThisWorkbook.Worksheets(5).Cells(j, 23).Value <> "" Then
                TarihAra = ThisWorkbook.Worksheets(5).Range("W" & j)
                If CDate(TarihAra) < CDate(CalTarih) Then
                    If (ThisWorkbook.Worksheets(5).Cells(j, 176).Value <> "" And CDate(ThisWorkbook.Worksheets(5).Cells(j, 176).Value) > CDate(CalTarih)) Or _
                        ThisWorkbook.Worksheets(5).Cells(j, 176).Value = "" Then
                        'MsgBox TarihAra & " < " & CalTarih
                        If ThisWorkbook.Worksheets(5).Cells(j, 5).Value <> "" And ThisWorkbook.Worksheets(5).Cells(j, 13).Value <> "" Then
                            Cont = Cont + 1
                            
                            Set IlkSiraBul = Nothing
                            Set SonSiraBul = Nothing
                            IlkSira = 0
                            SonSira = 0
                            SiraNo = ThisWorkbook.Worksheets(5).Cells(j, 5).Value
                            Set IlkSiraBul = ThisWorkbook.Worksheets(5).Range("FG7:FG100000").Find(What:=SiraNo, SearchDirection:=xlNext, _
                                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
                            If Not IlkSiraBul Is Nothing Then
                                IlkSira = IlkSiraBul.Row
                            End If
                            Set SonSiraBul = ThisWorkbook.Worksheets(5).Range("FH7:FH100000").Find(What:=SiraNo, SearchDirection:=xlNext, _
                                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
                            If Not SonSiraBul Is Nothing Then
                                SonSira = SonSiraBul.Row
                            End If
                            RaporNoKontrol = 0
                            For k = IlkSira To SonSira
                                If ThisWorkbook.Worksheets(5).Cells(k, 13).Value <> "" Then
                                    If RaporNoKontrol = 0 Then
                                        RaporNoTakip = ThisWorkbook.Worksheets(5).Cells(k, 13).Value
                                        RaporNoKontrol = 1
                                    Else
                                        RaporNoTakip = RaporNoTakip & ", " & ThisWorkbook.Worksheets(5).Cells(k, 13).Value
                                    End If
                                End If
                            Next k
                            
                            If ThisWorkbook.Worksheets(5).Cells(j, 48).Value <> "" Then
                                If Left(ThisWorkbook.Worksheets(5).Cells(j, 47).Value, 3) = "İl " Then
                                    Sno = ThisWorkbook.Worksheets(5).Cells(j, 5).Value & " | M | " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(5).Cells(j, 48).Value
                                ElseIf Left(ThisWorkbook.Worksheets(5).Cells(j, 47).Value, 3) = "İlç" Then
                                    Sno = ThisWorkbook.Worksheets(5).Cells(j, 5).Value & " | M | " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(5).Cells(j, 48).Value
                                Else
                                    Sno = ThisWorkbook.Worksheets(5).Cells(j, 5).Value & " | M | " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(5).Cells(j, 47).Value & " " & ThisWorkbook.Worksheets(5).Cells(j, 48).Value
                                End If
                            ElseIf ThisWorkbook.Worksheets(5).Cells(j, 48).Value = "" Then
                                If Left(ThisWorkbook.Worksheets(5).Cells(j, 47).Value, 3) = "İl " Then
                                    Sno = ThisWorkbook.Worksheets(5).Cells(j, 5).Value & " | M | " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(5).Cells(j, 19).Value & " " & ThisWorkbook.Worksheets(5).Cells(j, 47).Value
                                ElseIf Left(ThisWorkbook.Worksheets(5).Cells(j, 47).Value, 3) = "İlç" Then
                                    Sno = ThisWorkbook.Worksheets(5).Cells(j, 5).Value & " | M | " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(5).Cells(j, 20).Value & " " & ThisWorkbook.Worksheets(5).Cells(j, 47).Value
                                Else
                                    Sno = ThisWorkbook.Worksheets(5).Cells(j, 5).Value & " | M | " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(5).Cells(j, 47).Value
                                End If
                            End If
                    
                            Set LstBx = core_asset_manager_UI.FrameMevcut.Controls.Add("Forms.ListBox.1", "Frame1LstBx" & Cont)
                            With LstBx
                                .Top = (Cont - 1) * 12
                                .Left = 18
                                .Height = 12
                                .Width = 300
                                .BackColor = &H80000003
                                .SpecialEffect = fmSpecialEffectFlat
                                .MultiSelect = fmMultiSelectMulti
                                .TextAlign = fmTextAlignLeft '02092021
                                .AddItem (Sno)
                            End With
        
                            Set LblSira1 = core_asset_manager_UI.FrameMevcut.Controls.Add("Forms.Label.1", "Lbl" & Cont)
                            With LblSira1
                                .Top = (Cont - 1) * 12
                                .Left = 0
                                .Height = 12
                                .Width = 18
                                .SpecialEffect = fmSpecialEffectEtched
                                .TextAlign = fmTextAlignCenter
                                .Caption = Cont
                            End With
        
                            ScrollTakip2 = ScrollTakip2 + 12
                        End If
                    End If
                End If
            End If
        Else
            If ThisWorkbook.Worksheets(5).Cells(j, 23).Value <> "" Then
                TarihAra = ThisWorkbook.Worksheets(5).Range("W" & j)
                If CDate(TarihAra) < CDate(CalTarih) Then
                    If (ThisWorkbook.Worksheets(5).Cells(j, 176).Value <> "" And CDate(ThisWorkbook.Worksheets(5).Cells(j, 176).Value) > CDate(CalTarih)) Or _
                        ThisWorkbook.Worksheets(5).Cells(j, 176).Value = "" Then
                        'MsgBox TarihAra & " < " & CalTarih
                        If ThisWorkbook.Worksheets(5).Cells(j, 5).Value <> "" Then 'And ThisWorkbook.Worksheets(5).Cells(j, 13).Value <> "" Then
                            Cont = Cont + 1
                            If ThisWorkbook.Worksheets(5).Cells(j, 48).Value <> "" Then
                                If Left(ThisWorkbook.Worksheets(5).Cells(j, 47).Value, 3) = "İl " Then
                                    Sno = ThisWorkbook.Worksheets(5).Cells(j, 5).Value & " | M | " & ThisWorkbook.Worksheets(5).Cells(j, 26).Value & " | " & ThisWorkbook.Worksheets(5).Cells(j, 48).Value
                                ElseIf Left(ThisWorkbook.Worksheets(5).Cells(j, 47).Value, 3) = "İlç" Then
                                    Sno = ThisWorkbook.Worksheets(5).Cells(j, 5).Value & " | M | " & ThisWorkbook.Worksheets(5).Cells(j, 26).Value & " | " & ThisWorkbook.Worksheets(5).Cells(j, 48).Value
                                Else
                                    Sno = ThisWorkbook.Worksheets(5).Cells(j, 5).Value & " | M | " & ThisWorkbook.Worksheets(5).Cells(j, 26).Value & " | " & ThisWorkbook.Worksheets(5).Cells(j, 47).Value & " " & ThisWorkbook.Worksheets(5).Cells(j, 48).Value
                                End If
                            ElseIf ThisWorkbook.Worksheets(5).Cells(j, 48).Value = "" Then
                                If Left(ThisWorkbook.Worksheets(5).Cells(j, 47).Value, 3) = "İl " Then
                                    Sno = ThisWorkbook.Worksheets(5).Cells(j, 5).Value & " | M | " & ThisWorkbook.Worksheets(5).Cells(j, 26).Value & " | " & ThisWorkbook.Worksheets(5).Cells(j, 19).Value & " " & ThisWorkbook.Worksheets(5).Cells(j, 47).Value
                                ElseIf Left(ThisWorkbook.Worksheets(5).Cells(j, 47).Value, 3) = "İlç" Then
                                    Sno = ThisWorkbook.Worksheets(5).Cells(j, 5).Value & " | M | " & ThisWorkbook.Worksheets(5).Cells(j, 26).Value & " | " & ThisWorkbook.Worksheets(5).Cells(j, 20).Value & " " & ThisWorkbook.Worksheets(5).Cells(j, 47).Value
                                Else
                                    Sno = ThisWorkbook.Worksheets(5).Cells(j, 5).Value & " | M | " & ThisWorkbook.Worksheets(5).Cells(j, 26).Value & " | " & ThisWorkbook.Worksheets(5).Cells(j, 47).Value
                                End If
                            End If
                    
                            Set LstBx = core_asset_manager_UI.FrameMevcut.Controls.Add("Forms.ListBox.1", "Frame1LstBx" & Cont)
                            With LstBx
                                .Top = (Cont - 1) * 12
                                .Left = 18
                                .Height = 12
                                .Width = 300
                                .BackColor = &H80000003
                                .SpecialEffect = fmSpecialEffectFlat
                                .MultiSelect = fmMultiSelectMulti
                                .TextAlign = fmTextAlignLeft '02092021
                                .AddItem (Sno)
                            End With
        
                            Set LblSira1 = core_asset_manager_UI.FrameMevcut.Controls.Add("Forms.Label.1", "Lbl" & Cont)
                            With LblSira1
                                .Top = (Cont - 1) * 12
                                .Left = 0
                                .Height = 12
                                .Width = 18
                                .SpecialEffect = fmSpecialEffectEtched
                                .TextAlign = fmTextAlignCenter
                                .Caption = Cont
                            End With
        
                            ScrollTakip2 = ScrollTakip2 + 12
                        End If
                    End If
                End If
            End If
        End If
    End If
Next j

ContTakipMevcut = Cont

Son:

End Sub

Sub Rapor3_2VarlikHareketleriCikis()
Dim i As Integer, TarihBul As Range
Dim Say As Long, j As Long, Cont As Long, Sno As Variant
Dim LstBx As MSForms.ListBox, ctl As MSForms.Control
Dim LblSira1 As MSForms.Label
Dim LblSira2 As MSForms.Label
Dim Contx As Long

Dim IlkSiraBul As Range, SonSiraBul As Range, IlkSira As Long, SonSira As Long, k As Long
Dim RaporNoTakip As String, RaporNoKontrol As Integer, SiraNo As Long


CalTarih = CalTarihTakip

Say = ThisWorkbook.Worksheets(5).Range("FH100000").End(xlUp).Row
If Say < 7 Then
    GoTo Son
End If

'Cont = 0
Contx = 0
Set TarihBul = ThisWorkbook.Worksheets(5).Range("FT6:FT100000").Find(What:=CalTarih, SearchDirection:=xlNext, _
                SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
If Not TarihBul Is Nothing Then
    '
Else
    GoTo Son
End If

'MsgBox CalTarih

Cont = ContTakipCikis
'Cont = 0
For j = TarihBul.Row To Say
    If ThisWorkbook.Worksheets(5).Cells(j, 12).Value = "Point2" Or ThisWorkbook.Worksheets(5).Cells(j, 12).Value = "Point3" Then
        If ThisWorkbook.Worksheets(5).Cells(j, 28).Value = "Type A" Then
            If CalTarih = ThisWorkbook.Worksheets(5).Cells(j, 176).Value And ThisWorkbook.Worksheets(5).Cells(j, 23).Value <> "" Then
                If ThisWorkbook.Worksheets(5).Cells(j, 5).Value <> "" And ThisWorkbook.Worksheets(5).Cells(j, 13).Value <> "" Then
                    Cont = Cont + 1
                    
                    Set IlkSiraBul = Nothing
                    Set SonSiraBul = Nothing
                    IlkSira = 0
                    SonSira = 0
                    SiraNo = ThisWorkbook.Worksheets(5).Cells(j, 5).Value
                    Set IlkSiraBul = ThisWorkbook.Worksheets(5).Range("FG7:FG100000").Find(What:=SiraNo, SearchDirection:=xlNext, _
                                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
                    If Not IlkSiraBul Is Nothing Then
                        IlkSira = IlkSiraBul.Row
                    End If
                    Set SonSiraBul = ThisWorkbook.Worksheets(5).Range("FH7:FH100000").Find(What:=SiraNo, SearchDirection:=xlNext, _
                                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
                    If Not SonSiraBul Is Nothing Then
                        SonSira = SonSiraBul.Row
                    End If
                    RaporNoKontrol = 0
                    For k = IlkSira To SonSira
                        If ThisWorkbook.Worksheets(5).Cells(k, 13).Value <> "" Then
                            If RaporNoKontrol = 0 Then
                                RaporNoTakip = ThisWorkbook.Worksheets(5).Cells(k, 13).Value
                                RaporNoKontrol = 1
                            Else
                                RaporNoTakip = RaporNoTakip & ", " & ThisWorkbook.Worksheets(5).Cells(k, 13).Value
                            End If
                        End If
                    Next k
                    
                    If ThisWorkbook.Worksheets(5).Cells(j, 48).Value <> "" Then
                        If Left(ThisWorkbook.Worksheets(5).Cells(j, 47).Value, 3) = "İl " Then
                            Sno = ThisWorkbook.Worksheets(5).Cells(j, 5).Value & " | M | " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(5).Cells(j, 48).Value
                        ElseIf Left(ThisWorkbook.Worksheets(5).Cells(j, 47).Value, 3) = "İlç" Then
                            Sno = ThisWorkbook.Worksheets(5).Cells(j, 5).Value & " | M | " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(5).Cells(j, 48).Value
                        Else
                            Sno = ThisWorkbook.Worksheets(5).Cells(j, 5).Value & " | M | " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(5).Cells(j, 47).Value & " " & ThisWorkbook.Worksheets(5).Cells(j, 48).Value
                        End If
                    ElseIf ThisWorkbook.Worksheets(5).Cells(j, 48).Value = "" Then
                        If Left(ThisWorkbook.Worksheets(5).Cells(j, 47).Value, 3) = "İl " Then
                            Sno = ThisWorkbook.Worksheets(5).Cells(j, 5).Value & " | M | " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(5).Cells(j, 19).Value & " " & ThisWorkbook.Worksheets(5).Cells(j, 47).Value
                        ElseIf Left(ThisWorkbook.Worksheets(5).Cells(j, 47).Value, 3) = "İlç" Then
                            Sno = ThisWorkbook.Worksheets(5).Cells(j, 5).Value & " | M | " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(5).Cells(j, 20).Value & " " & ThisWorkbook.Worksheets(5).Cells(j, 47).Value
                        Else
                            Sno = ThisWorkbook.Worksheets(5).Cells(j, 5).Value & " | M | " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(5).Cells(j, 47).Value
                        End If
                    End If
                
                    Set LstBx = core_asset_manager_UI.FrameCikis.Controls.Add("Forms.ListBox.1", "Frame1LstBx" & Cont)
                    With LstBx
                        .Top = (Cont - 1) * 12
                        .Left = 18
                        .Height = 12
                        .Width = 300
                        If ThisWorkbook.Worksheets(5).Cells(j, 23).Value = ThisWorkbook.Worksheets(5).Cells(j, 176).Value Then
                            .BackColor = &H80000000 'Giriş
                        Else
                            .BackColor = &H80000003 'Mevcut
                        End If
                        .SpecialEffect = fmSpecialEffectFlat
                        .MultiSelect = fmMultiSelectMulti
                        .TextAlign = fmTextAlignLeft '02092021
                        .AddItem (Sno)
                    End With
        
                    Set LblSira1 = core_asset_manager_UI.FrameCikis.Controls.Add("Forms.Label.1", "Lbl" & Cont)
                    With LblSira1
                        .Top = (Cont - 1) * 12
                        .Left = 0
                        .Height = 12
                        .Width = 18
                        .SpecialEffect = fmSpecialEffectEtched
                        .TextAlign = fmTextAlignCenter
                        .Caption = Cont
                    End With
        
                    ScrollTakip3 = ScrollTakip3 + 12
                End If
            End If
        Else
            If CalTarih = ThisWorkbook.Worksheets(5).Cells(j, 176).Value And ThisWorkbook.Worksheets(5).Cells(j, 23).Value <> "" Then
                If ThisWorkbook.Worksheets(5).Cells(j, 5).Value <> "" Then 'And ThisWorkbook.Worksheets(5).Cells(j, 13).Value <> "" Then
                    Cont = Cont + 1
                    If ThisWorkbook.Worksheets(5).Cells(j, 48).Value <> "" Then
                        If Left(ThisWorkbook.Worksheets(5).Cells(j, 47).Value, 3) = "İl " Then
                            Sno = ThisWorkbook.Worksheets(5).Cells(j, 5).Value & " | M | " & ThisWorkbook.Worksheets(5).Cells(j, 26).Value & " | " & ThisWorkbook.Worksheets(5).Cells(j, 48).Value
                        ElseIf Left(ThisWorkbook.Worksheets(5).Cells(j, 47).Value, 3) = "İlç" Then
                            Sno = ThisWorkbook.Worksheets(5).Cells(j, 5).Value & " | M | " & ThisWorkbook.Worksheets(5).Cells(j, 26).Value & " | " & ThisWorkbook.Worksheets(5).Cells(j, 48).Value
                        Else
                            Sno = ThisWorkbook.Worksheets(5).Cells(j, 5).Value & " | M | " & ThisWorkbook.Worksheets(5).Cells(j, 26).Value & " | " & ThisWorkbook.Worksheets(5).Cells(j, 47).Value & " " & ThisWorkbook.Worksheets(5).Cells(j, 48).Value
                        End If
                    ElseIf ThisWorkbook.Worksheets(5).Cells(j, 48).Value = "" Then
                        If Left(ThisWorkbook.Worksheets(5).Cells(j, 47).Value, 3) = "İl " Then
                            Sno = ThisWorkbook.Worksheets(5).Cells(j, 5).Value & " | M | " & ThisWorkbook.Worksheets(5).Cells(j, 26).Value & " | " & ThisWorkbook.Worksheets(5).Cells(j, 19).Value & " " & ThisWorkbook.Worksheets(5).Cells(j, 47).Value
                        ElseIf Left(ThisWorkbook.Worksheets(5).Cells(j, 47).Value, 3) = "İlç" Then
                            Sno = ThisWorkbook.Worksheets(5).Cells(j, 5).Value & " | M | " & ThisWorkbook.Worksheets(5).Cells(j, 26).Value & " | " & ThisWorkbook.Worksheets(5).Cells(j, 20).Value & " " & ThisWorkbook.Worksheets(5).Cells(j, 47).Value
                        Else
                            Sno = ThisWorkbook.Worksheets(5).Cells(j, 5).Value & " | M | " & ThisWorkbook.Worksheets(5).Cells(j, 26).Value & " | " & ThisWorkbook.Worksheets(5).Cells(j, 47).Value
                        End If
                    End If
                
                    Set LstBx = core_asset_manager_UI.FrameCikis.Controls.Add("Forms.ListBox.1", "Frame1LstBx" & Cont)
                    With LstBx
                        .Top = (Cont - 1) * 12
                        .Left = 18
                        .Height = 12
                        .Width = 300
                        If ThisWorkbook.Worksheets(5).Cells(j, 23).Value = ThisWorkbook.Worksheets(5).Cells(j, 176).Value Then
                            .BackColor = &H80000000 'Giriş
                        Else
                            .BackColor = &H80000003 'Mevcut
                        End If
                        .SpecialEffect = fmSpecialEffectFlat
                        .MultiSelect = fmMultiSelectMulti
                        .TextAlign = fmTextAlignLeft '02092021
                        .AddItem (Sno)
                    End With
        
                    Set LblSira1 = core_asset_manager_UI.FrameCikis.Controls.Add("Forms.Label.1", "Lbl" & Cont)
                    With LblSira1
                        .Top = (Cont - 1) * 12
                        .Left = 0
                        .Height = 12
                        .Width = 18
                        .SpecialEffect = fmSpecialEffectEtched
                        .TextAlign = fmTextAlignCenter
                        .Caption = Cont
                    End With
        
                    ScrollTakip3 = ScrollTakip3 + 12
                End If
            End If
        End If
    End If
Next j

ContTakipCikis = Cont

Son:

End Sub

Sub Rapor3_1VarlikHareketleriGiris()
Dim i As Integer, TarihBul As Range
Dim Say As Long, j As Long, Cont As Long, Sno As Variant
Dim LstBx As MSForms.ListBox, ctl As MSForms.Control
Dim LblSira1 As MSForms.Label
Dim LblSira2 As MSForms.Label
Dim Contx As Long

Dim IlkSiraBul As Range, SonSiraBul As Range, IlkSira As Long, SonSira As Long, k As Long
Dim RaporNoTakip As String, RaporNoKontrol As Integer, SiraNo As Long


CalTarih = CalTarihTakip

Say = ThisWorkbook.Worksheets(5).Range("FH100000").End(xlUp).Row
If Say < 7 Then
    GoTo Son
End If

'Cont = 0
Contx = 0
Set TarihBul = ThisWorkbook.Worksheets(5).Range("CQ6:CQ100000").Find(What:=CalTarih, SearchDirection:=xlNext, _
                SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole) 'Tutanak tarihi içinde ara
If Not TarihBul Is Nothing Then
    '
Else
    GoTo Son
End If

'MsgBox CalTarih

Cont = ContTakipGiris
'Cont = 0
For j = TarihBul.Row To Say
    If ThisWorkbook.Worksheets(5).Cells(j, 12).Value = "Point1" Then
        If ThisWorkbook.Worksheets(5).Cells(j, 100).Value = "Type A" Then
            If CalTarih = ThisWorkbook.Worksheets(5).Cells(j, 95).Value And CDate(ThisWorkbook.Worksheets(5).Cells(j, 95).Value) <> CDate(ThisWorkbook.Worksheets(5).Cells(j, 176).Value) Then
                If ThisWorkbook.Worksheets(5).Cells(j, 5).Value <> "" And ThisWorkbook.Worksheets(5).Cells(j, 13).Value <> "" Then
                    Cont = Cont + 1

                    Set IlkSiraBul = Nothing
                    Set SonSiraBul = Nothing
                    IlkSira = 0
                    SonSira = 0
                    SiraNo = ThisWorkbook.Worksheets(5).Cells(j, 5).Value
                    Set IlkSiraBul = ThisWorkbook.Worksheets(5).Range("FG7:FG100000").Find(What:=SiraNo, SearchDirection:=xlNext, _
                                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
                    If Not IlkSiraBul Is Nothing Then
                        IlkSira = IlkSiraBul.Row
                    End If
                    Set SonSiraBul = ThisWorkbook.Worksheets(5).Range("FH7:FH100000").Find(What:=SiraNo, SearchDirection:=xlNext, _
                                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
                    If Not SonSiraBul Is Nothing Then
                        SonSira = SonSiraBul.Row
                    End If
                    RaporNoKontrol = 0
                    For k = IlkSira To SonSira
                        If ThisWorkbook.Worksheets(5).Cells(k, 13).Value <> "" Then
                            If RaporNoKontrol = 0 Then
                                RaporNoTakip = ThisWorkbook.Worksheets(5).Cells(k, 13).Value
                                RaporNoKontrol = 1
                            Else
                                RaporNoTakip = RaporNoTakip & ", " & ThisWorkbook.Worksheets(5).Cells(k, 13).Value
                            End If
                        End If
                    Next k
                
                    If ThisWorkbook.Worksheets(5).Cells(j, 103).Value <> "" Then
                        If Left(ThisWorkbook.Worksheets(5).Cells(j, 102).Value, 3) = "İl " Then
                            Sno = ThisWorkbook.Worksheets(5).Cells(j, 5).Value & " | G | " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(5).Cells(j, 103).Value
                        ElseIf Left(ThisWorkbook.Worksheets(5).Cells(j, 102).Value, 3) = "İlç" Then
                            Sno = ThisWorkbook.Worksheets(5).Cells(j, 5).Value & " | G | " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(5).Cells(j, 103).Value
                        Else
                            Sno = ThisWorkbook.Worksheets(5).Cells(j, 5).Value & " | G | " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(5).Cells(j, 102).Value & " " & ThisWorkbook.Worksheets(5).Cells(j, 103).Value
                        End If
                    ElseIf ThisWorkbook.Worksheets(5).Cells(j, 103).Value = "" Then
                        If Left(ThisWorkbook.Worksheets(5).Cells(j, 102).Value, 3) = "İl " Then
                            Sno = ThisWorkbook.Worksheets(5).Cells(j, 5).Value & " | G | " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(5).Cells(j, 91).Value & ThisWorkbook.Worksheets(5).Cells(j, 102).Value
                        ElseIf Left(ThisWorkbook.Worksheets(5).Cells(j, 102).Value, 3) = "İlç" Then
                            Sno = ThisWorkbook.Worksheets(5).Cells(j, 5).Value & " | G | " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(5).Cells(j, 92).Value & ThisWorkbook.Worksheets(5).Cells(j, 102).Value
                        Else
                            Sno = ThisWorkbook.Worksheets(5).Cells(j, 5).Value & " | G | " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(5).Cells(j, 102).Value
                        End If
                    End If
                
                    Set LstBx = core_asset_manager_UI.FrameGiris.Controls.Add("Forms.ListBox.1", "Frame1LstBx" & Cont)
                    With LstBx
                        .Top = (Cont - 1) * 12
                        .Left = 18
                        .Height = 12
                        .Width = 300
                        .BackColor = &H80000000
                        .SpecialEffect = fmSpecialEffectFlat
                        .MultiSelect = fmMultiSelectMulti
                        .TextAlign = fmTextAlignLeft '02092021
                        .AddItem (Sno)
                    End With
        
                    Set LblSira1 = core_asset_manager_UI.FrameGiris.Controls.Add("Forms.Label.1", "Lbl" & Cont)
                    With LblSira1
                        .Top = (Cont - 1) * 12
                        .Left = 0
                        .Height = 12
                        .Width = 18
                        .SpecialEffect = fmSpecialEffectEtched
                        .TextAlign = fmTextAlignCenter
                        .Caption = Cont
                    End With
        
                    ScrollTakip1 = ScrollTakip1 + 12
                End If
            End If
        Else
            If CalTarih = ThisWorkbook.Worksheets(5).Cells(j, 95).Value And CDate(ThisWorkbook.Worksheets(5).Cells(j, 95).Value) <> CDate(ThisWorkbook.Worksheets(5).Cells(j, 176).Value) Then
                If ThisWorkbook.Worksheets(5).Cells(j, 5).Value <> "" Then 'And ThisWorkbook.Worksheets(5).Cells(j, 13).Value <> "" Then
                    Cont = Cont + 1
                    If ThisWorkbook.Worksheets(5).Cells(j, 103).Value <> "" Then
                        If Left(ThisWorkbook.Worksheets(5).Cells(j, 102).Value, 3) = "İl " Then
                            Sno = ThisWorkbook.Worksheets(5).Cells(j, 5).Value & " | G | " & ThisWorkbook.Worksheets(5).Cells(j, 98).Value & " | " & ThisWorkbook.Worksheets(5).Cells(j, 103).Value
                        ElseIf Left(ThisWorkbook.Worksheets(5).Cells(j, 102).Value, 3) = "İlç" Then
                            Sno = ThisWorkbook.Worksheets(5).Cells(j, 5).Value & " | G | " & ThisWorkbook.Worksheets(5).Cells(j, 98).Value & " | " & ThisWorkbook.Worksheets(5).Cells(j, 103).Value
                        Else
                            Sno = ThisWorkbook.Worksheets(5).Cells(j, 5).Value & " | G | " & ThisWorkbook.Worksheets(5).Cells(j, 98).Value & " | " & ThisWorkbook.Worksheets(5).Cells(j, 102).Value & " " & ThisWorkbook.Worksheets(5).Cells(j, 103).Value
                        End If
                    ElseIf ThisWorkbook.Worksheets(5).Cells(j, 103).Value = "" Then
                        If Left(ThisWorkbook.Worksheets(5).Cells(j, 102).Value, 3) = "İl " Then
                            Sno = ThisWorkbook.Worksheets(5).Cells(j, 5).Value & " | G | " & ThisWorkbook.Worksheets(5).Cells(j, 98).Value & " | " & ThisWorkbook.Worksheets(5).Cells(j, 91).Value & ThisWorkbook.Worksheets(5).Cells(j, 102).Value
                        ElseIf Left(ThisWorkbook.Worksheets(5).Cells(j, 102).Value, 3) = "İlç" Then
                            Sno = ThisWorkbook.Worksheets(5).Cells(j, 5).Value & " | G | " & ThisWorkbook.Worksheets(5).Cells(j, 98).Value & " | " & ThisWorkbook.Worksheets(5).Cells(j, 92).Value & ThisWorkbook.Worksheets(5).Cells(j, 102).Value
                        Else
                            Sno = ThisWorkbook.Worksheets(5).Cells(j, 5).Value & " | G | " & ThisWorkbook.Worksheets(5).Cells(j, 98).Value & " | " & ThisWorkbook.Worksheets(5).Cells(j, 102).Value
                        End If
                    End If
                
                    Set LstBx = core_asset_manager_UI.FrameGiris.Controls.Add("Forms.ListBox.1", "Frame1LstBx" & Cont)
                    With LstBx
                        .Top = (Cont - 1) * 12
                        .Left = 18
                        .Height = 12
                        .Width = 300
                        .BackColor = &H80000000
                        .SpecialEffect = fmSpecialEffectFlat
                        .MultiSelect = fmMultiSelectMulti
                        .TextAlign = fmTextAlignLeft '02092021
                        .AddItem (Sno)
                    End With
        
                    Set LblSira1 = core_asset_manager_UI.FrameGiris.Controls.Add("Forms.Label.1", "Lbl" & Cont)
                    With LblSira1
                        .Top = (Cont - 1) * 12
                        .Left = 0
                        .Height = 12
                        .Width = 18
                        .SpecialEffect = fmSpecialEffectEtched
                        .TextAlign = fmTextAlignCenter
                        .Caption = Cont
                    End With
        
                    ScrollTakip1 = ScrollTakip1 + 12
                End If
            End If
        End If
    End If
Next j

ContTakipGiris = Cont

Son:

End Sub

Sub Rapor3_1VarlikHareketleriMevcut()
Dim i As Integer, TarihBul As Range, SiraNoBul As Range
Dim Say As Long, j As Long, Cont As Long, Sno As Variant
Dim LstBx As MSForms.ListBox, ctl As MSForms.Control
Dim LblSira1 As MSForms.Label
Dim LblSira2 As MSForms.Label
Dim Contx As Long, MevcutSatir As Long
Dim TarihAra As String
Dim SayDevir As Long, TakipCount As Long

Dim IlkSiraBul As Range, SonSiraBul As Range, IlkSira As Long, SonSira As Long, k As Long
Dim RaporNoTakip As String, RaporNoKontrol As Integer, SiraNo As Long

CalTarih = CalTarihTakip

Say = ThisWorkbook.Worksheets(5).Range("FH100000").End(xlUp).Row
If Say < 7 Then
    GoTo Son
End If

'Mevcutta satırı bul
'CalTarih = CDate(CalTarih)
MevcutSatir = 0
For j = Say To 7 Step -1
    If ThisWorkbook.Worksheets(5).Cells(j, 12).Value = "Point1" Then
        If ThisWorkbook.Worksheets(5).Range("CQ" & j).Value <> "" Then
            TarihAra = ThisWorkbook.Worksheets(5).Range("CQ" & j)
            If CDate(TarihAra) < CDate(CalTarih) Then
                MevcutSatir = j
                'MsgBox Range("CQ" & MevcutSatir)
                GoTo DonguSonu
            End If
        End If
    End If
Next j
DonguSonu:
'Mevcutta satır bulunamaz ise sona git
If MevcutSatir = 0 Then
    GoTo Son
End If

Cont = ContTakipMevcut

For j = MevcutSatir To 7 Step -1
    If ThisWorkbook.Worksheets(5).Cells(j, 12).Value = "Point1" Then
        If ThisWorkbook.Worksheets(5).Cells(j, 100).Value = "Type A" Then
            If ThisWorkbook.Worksheets(5).Cells(j, 95).Value <> "" Then
                TarihAra = ThisWorkbook.Worksheets(5).Range("CQ" & j)
                If CDate(TarihAra) < CDate(CalTarih) Then
                    If (ThisWorkbook.Worksheets(5).Cells(j, 176).Value <> "" And CDate(ThisWorkbook.Worksheets(5).Cells(j, 176).Value) > CDate(CalTarih)) Or _
                        ThisWorkbook.Worksheets(5).Cells(j, 176).Value = "" Then
                        'MsgBox TarihAra & " < " & CalTarih
                        If ThisWorkbook.Worksheets(5).Cells(j, 5).Value <> "" And ThisWorkbook.Worksheets(5).Cells(j, 13).Value <> "" Then
                            Cont = Cont + 1

                            Set IlkSiraBul = Nothing
                            Set SonSiraBul = Nothing
                            IlkSira = 0
                            SonSira = 0
                            SiraNo = ThisWorkbook.Worksheets(5).Cells(j, 5).Value
                            Set IlkSiraBul = ThisWorkbook.Worksheets(5).Range("FG7:FG100000").Find(What:=SiraNo, SearchDirection:=xlNext, _
                                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
                            If Not IlkSiraBul Is Nothing Then
                                IlkSira = IlkSiraBul.Row
                            End If
                            Set SonSiraBul = ThisWorkbook.Worksheets(5).Range("FH7:FH100000").Find(What:=SiraNo, SearchDirection:=xlNext, _
                                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
                            If Not SonSiraBul Is Nothing Then
                                SonSira = SonSiraBul.Row
                            End If
                            RaporNoKontrol = 0
                            For k = IlkSira To SonSira
                                If ThisWorkbook.Worksheets(5).Cells(k, 13).Value <> "" Then
                                    If RaporNoKontrol = 0 Then
                                        RaporNoTakip = ThisWorkbook.Worksheets(5).Cells(k, 13).Value
                                        RaporNoKontrol = 1
                                    Else
                                        RaporNoTakip = RaporNoTakip & ", " & ThisWorkbook.Worksheets(5).Cells(k, 13).Value
                                    End If
                                End If
                            Next k
                    
                            If ThisWorkbook.Worksheets(5).Cells(j, 103).Value <> "" Then
                                If Left(ThisWorkbook.Worksheets(5).Cells(j, 102).Value, 3) = "İl " Then
                                    Sno = ThisWorkbook.Worksheets(5).Cells(j, 5).Value & " | G | " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(5).Cells(j, 103).Value
                                ElseIf Left(ThisWorkbook.Worksheets(5).Cells(j, 102).Value, 3) = "İlç" Then
                                    Sno = ThisWorkbook.Worksheets(5).Cells(j, 5).Value & " | G | " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(5).Cells(j, 103).Value
                                Else
                                    Sno = ThisWorkbook.Worksheets(5).Cells(j, 5).Value & " | G | " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(5).Cells(j, 102).Value & " " & ThisWorkbook.Worksheets(5).Cells(j, 103).Value
                                End If
                            ElseIf ThisWorkbook.Worksheets(5).Cells(j, 103).Value = "" Then
                                If Left(ThisWorkbook.Worksheets(5).Cells(j, 102).Value, 3) = "İl " Then
                                    Sno = ThisWorkbook.Worksheets(5).Cells(j, 5).Value & " | G | " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(5).Cells(j, 91).Value & ThisWorkbook.Worksheets(5).Cells(j, 102).Value
                                ElseIf Left(ThisWorkbook.Worksheets(5).Cells(j, 102).Value, 3) = "İlç" Then
                                    Sno = ThisWorkbook.Worksheets(5).Cells(j, 5).Value & " | G | " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(5).Cells(j, 92).Value & ThisWorkbook.Worksheets(5).Cells(j, 102).Value
                                Else
                                    Sno = ThisWorkbook.Worksheets(5).Cells(j, 5).Value & " | G | " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(5).Cells(j, 102).Value
                                End If
                            End If
                    
                            Set LstBx = core_asset_manager_UI.FrameMevcut.Controls.Add("Forms.ListBox.1", "Frame1LstBx" & Cont)
                            With LstBx
                                .Top = (Cont - 1) * 12
                                .Left = 18
                                .Height = 12
                                .Width = 300
                                .BackColor = &H80000003
                                .SpecialEffect = fmSpecialEffectFlat
                                .MultiSelect = fmMultiSelectMulti
                                .TextAlign = fmTextAlignLeft '02092021
                                .AddItem (Sno)
                            End With
        
                            Set LblSira1 = core_asset_manager_UI.FrameMevcut.Controls.Add("Forms.Label.1", "Lbl" & Cont)
                            With LblSira1
                                .Top = (Cont - 1) * 12
                                .Left = 0
                                .Height = 12
                                .Width = 18
                                .SpecialEffect = fmSpecialEffectEtched
                                .TextAlign = fmTextAlignCenter
                                .Caption = Cont
                            End With
        
                            ScrollTakip2 = ScrollTakip2 + 12
                        End If
                    End If
                End If
            End If
        Else
            If ThisWorkbook.Worksheets(5).Cells(j, 95).Value <> "" Then
                TarihAra = ThisWorkbook.Worksheets(5).Range("CQ" & j)
                If CDate(TarihAra) < CDate(CalTarih) Then
                    If (ThisWorkbook.Worksheets(5).Cells(j, 176).Value <> "" And CDate(ThisWorkbook.Worksheets(5).Cells(j, 176).Value) > CDate(CalTarih)) Or _
                        ThisWorkbook.Worksheets(5).Cells(j, 176).Value = "" Then
                        'MsgBox TarihAra & " < " & CalTarih
                        If ThisWorkbook.Worksheets(5).Cells(j, 5).Value <> "" Then 'And ThisWorkbook.Worksheets(5).Cells(j, 13).Value <> "" Then
                            Cont = Cont + 1
                            If ThisWorkbook.Worksheets(5).Cells(j, 103).Value <> "" Then
                                If Left(ThisWorkbook.Worksheets(5).Cells(j, 102).Value, 3) = "İl " Then
                                    Sno = ThisWorkbook.Worksheets(5).Cells(j, 5).Value & " | G | " & ThisWorkbook.Worksheets(5).Cells(j, 98).Value & " | " & ThisWorkbook.Worksheets(5).Cells(j, 103).Value
                                ElseIf Left(ThisWorkbook.Worksheets(5).Cells(j, 102).Value, 3) = "İlç" Then
                                    Sno = ThisWorkbook.Worksheets(5).Cells(j, 5).Value & " | G | " & ThisWorkbook.Worksheets(5).Cells(j, 98).Value & " | " & ThisWorkbook.Worksheets(5).Cells(j, 103).Value
                                Else
                                    Sno = ThisWorkbook.Worksheets(5).Cells(j, 5).Value & " | G | " & ThisWorkbook.Worksheets(5).Cells(j, 98).Value & " | " & ThisWorkbook.Worksheets(5).Cells(j, 102).Value & " " & ThisWorkbook.Worksheets(5).Cells(j, 103).Value
                                End If
                            ElseIf ThisWorkbook.Worksheets(5).Cells(j, 103).Value = "" Then
                                If Left(ThisWorkbook.Worksheets(5).Cells(j, 102).Value, 3) = "İl " Then
                                    Sno = ThisWorkbook.Worksheets(5).Cells(j, 5).Value & " | G | " & ThisWorkbook.Worksheets(5).Cells(j, 98).Value & " | " & ThisWorkbook.Worksheets(5).Cells(j, 91).Value & ThisWorkbook.Worksheets(5).Cells(j, 102).Value
                                ElseIf Left(ThisWorkbook.Worksheets(5).Cells(j, 102).Value, 3) = "İlç" Then
                                    Sno = ThisWorkbook.Worksheets(5).Cells(j, 5).Value & " | G | " & ThisWorkbook.Worksheets(5).Cells(j, 98).Value & " | " & ThisWorkbook.Worksheets(5).Cells(j, 92).Value & ThisWorkbook.Worksheets(5).Cells(j, 102).Value
                                Else
                                    Sno = ThisWorkbook.Worksheets(5).Cells(j, 5).Value & " | G | " & ThisWorkbook.Worksheets(5).Cells(j, 98).Value & " | " & ThisWorkbook.Worksheets(5).Cells(j, 102).Value
                                End If
                            End If
                    
                            Set LstBx = core_asset_manager_UI.FrameMevcut.Controls.Add("Forms.ListBox.1", "Frame1LstBx" & Cont)
                            With LstBx
                                .Top = (Cont - 1) * 12
                                .Left = 18
                                .Height = 12
                                .Width = 300
                                .BackColor = &H80000003
                                .SpecialEffect = fmSpecialEffectFlat
                                .MultiSelect = fmMultiSelectMulti
                                .TextAlign = fmTextAlignLeft '02092021
                                .AddItem (Sno)
                            End With
        
                            Set LblSira1 = core_asset_manager_UI.FrameMevcut.Controls.Add("Forms.Label.1", "Lbl" & Cont)
                            With LblSira1
                                .Top = (Cont - 1) * 12
                                .Left = 0
                                .Height = 12
                                .Width = 18
                                .SpecialEffect = fmSpecialEffectEtched
                                .TextAlign = fmTextAlignCenter
                                .Caption = Cont
                            End With
        
                            ScrollTakip2 = ScrollTakip2 + 12
                        End If
                    End If
                End If
            End If
        End If
    End If
Next j

ContTakipMevcut = Cont

Son:

End Sub

Sub Rapor3_1VarlikHareketleriCikis()
Dim i As Integer, TarihBul As Range
Dim Say As Long, j As Long, Cont As Long, Sno As Variant
Dim LstBx As MSForms.ListBox, ctl As MSForms.Control
Dim LblSira1 As MSForms.Label
Dim LblSira2 As MSForms.Label
Dim Contx As Long

Dim IlkSiraBul As Range, SonSiraBul As Range, IlkSira As Long, SonSira As Long, k As Long
Dim RaporNoTakip As String, RaporNoKontrol As Integer, SiraNo As Long

CalTarih = CalTarihTakip

Say = ThisWorkbook.Worksheets(5).Range("FH100000").End(xlUp).Row
If Say < 7 Then
    GoTo Son
End If

'Cont = 0
Contx = 0
Set TarihBul = ThisWorkbook.Worksheets(5).Range("FT6:FT100000").Find(What:=CalTarih, SearchDirection:=xlNext, _
                SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
If Not TarihBul Is Nothing Then
    '
Else
    GoTo Son
End If

'MsgBox CalTarih

Cont = ContTakipCikis
'Cont = 0
For j = TarihBul.Row To Say
    If ThisWorkbook.Worksheets(5).Cells(j, 12).Value = "Point1" Then
        If ThisWorkbook.Worksheets(5).Cells(j, 100).Value = "Type A" Then
            If CalTarih = ThisWorkbook.Worksheets(5).Cells(j, 176).Value And ThisWorkbook.Worksheets(5).Cells(j, 95).Value <> "" Then
                If ThisWorkbook.Worksheets(5).Cells(j, 5).Value <> "" And ThisWorkbook.Worksheets(5).Cells(j, 13).Value <> "" Then
                    Cont = Cont + 1
                    
                    Set IlkSiraBul = Nothing
                    Set SonSiraBul = Nothing
                    IlkSira = 0
                    SonSira = 0
                    SiraNo = ThisWorkbook.Worksheets(5).Cells(j, 5).Value
                    Set IlkSiraBul = ThisWorkbook.Worksheets(5).Range("FG7:FG100000").Find(What:=SiraNo, SearchDirection:=xlNext, _
                                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
                    If Not IlkSiraBul Is Nothing Then
                        IlkSira = IlkSiraBul.Row
                    End If
                    Set SonSiraBul = ThisWorkbook.Worksheets(5).Range("FH7:FH100000").Find(What:=SiraNo, SearchDirection:=xlNext, _
                                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
                    If Not SonSiraBul Is Nothing Then
                        SonSira = SonSiraBul.Row
                    End If
                    RaporNoKontrol = 0
                    For k = IlkSira To SonSira
                        If ThisWorkbook.Worksheets(5).Cells(k, 13).Value <> "" Then
                            If RaporNoKontrol = 0 Then
                                RaporNoTakip = ThisWorkbook.Worksheets(5).Cells(k, 13).Value
                                RaporNoKontrol = 1
                            Else
                                RaporNoTakip = RaporNoTakip & ", " & ThisWorkbook.Worksheets(5).Cells(k, 13).Value
                            End If
                        End If
                    Next k
                            
                    If ThisWorkbook.Worksheets(5).Cells(j, 103).Value <> "" Then
                        If Left(ThisWorkbook.Worksheets(5).Cells(j, 102).Value, 3) = "İl " Then
                            Sno = ThisWorkbook.Worksheets(5).Cells(j, 5).Value & " | G | " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(5).Cells(j, 103).Value
                        ElseIf Left(ThisWorkbook.Worksheets(5).Cells(j, 102).Value, 3) = "İlç" Then
                            Sno = ThisWorkbook.Worksheets(5).Cells(j, 5).Value & " | G | " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(5).Cells(j, 103).Value
                        Else
                            Sno = ThisWorkbook.Worksheets(5).Cells(j, 5).Value & " | G | " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(5).Cells(j, 102).Value & " " & ThisWorkbook.Worksheets(5).Cells(j, 103).Value
                        End If
                    ElseIf ThisWorkbook.Worksheets(5).Cells(j, 103).Value = "" Then
                        If Left(ThisWorkbook.Worksheets(5).Cells(j, 102).Value, 3) = "İl " Then
                            Sno = ThisWorkbook.Worksheets(5).Cells(j, 5).Value & " | G | " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(5).Cells(j, 91).Value & ThisWorkbook.Worksheets(5).Cells(j, 102).Value
                        ElseIf Left(ThisWorkbook.Worksheets(5).Cells(j, 102).Value, 3) = "İlç" Then
                            Sno = ThisWorkbook.Worksheets(5).Cells(j, 5).Value & " | G | " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(5).Cells(j, 92).Value & ThisWorkbook.Worksheets(5).Cells(j, 102).Value
                        Else
                            Sno = ThisWorkbook.Worksheets(5).Cells(j, 5).Value & " | G | " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(5).Cells(j, 102).Value
                        End If
                    End If
                
                    Set LstBx = core_asset_manager_UI.FrameCikis.Controls.Add("Forms.ListBox.1", "Frame1LstBx" & Cont)
                    With LstBx
                        .Top = (Cont - 1) * 12
                        .Left = 18
                        .Height = 12
                        .Width = 300
                        If ThisWorkbook.Worksheets(5).Cells(j, 95).Value = ThisWorkbook.Worksheets(5).Cells(j, 176).Value Then
                            .BackColor = &H80000000 'Giriş
                        Else
                            .BackColor = &H80000003 'Mevcut
                        End If
                        .SpecialEffect = fmSpecialEffectFlat
                        .MultiSelect = fmMultiSelectMulti
                        .TextAlign = fmTextAlignLeft '02092021
                        .AddItem (Sno)
                    End With
        
                    Set LblSira1 = core_asset_manager_UI.FrameCikis.Controls.Add("Forms.Label.1", "Lbl" & Cont)
                    With LblSira1
                        .Top = (Cont - 1) * 12
                        .Left = 0
                        .Height = 12
                        .Width = 18
                        .SpecialEffect = fmSpecialEffectEtched
                        .TextAlign = fmTextAlignCenter
                        .Caption = Cont
                    End With
        
                    ScrollTakip3 = ScrollTakip3 + 12
                End If
            End If
        Else
            If CalTarih = ThisWorkbook.Worksheets(5).Cells(j, 176).Value And ThisWorkbook.Worksheets(5).Cells(j, 95).Value <> "" Then
                If ThisWorkbook.Worksheets(5).Cells(j, 5).Value <> "" Then 'And ThisWorkbook.Worksheets(5).Cells(j, 13).Value <> "" Then
                    Cont = Cont + 1
                    If ThisWorkbook.Worksheets(5).Cells(j, 103).Value <> "" Then
                        If Left(ThisWorkbook.Worksheets(5).Cells(j, 102).Value, 3) = "İl " Then
                            Sno = ThisWorkbook.Worksheets(5).Cells(j, 5).Value & " | G | " & ThisWorkbook.Worksheets(5).Cells(j, 98).Value & " | " & ThisWorkbook.Worksheets(5).Cells(j, 103).Value
                        ElseIf Left(ThisWorkbook.Worksheets(5).Cells(j, 102).Value, 3) = "İlç" Then
                            Sno = ThisWorkbook.Worksheets(5).Cells(j, 5).Value & " | G | " & ThisWorkbook.Worksheets(5).Cells(j, 98).Value & " | " & ThisWorkbook.Worksheets(5).Cells(j, 103).Value
                        Else
                            Sno = ThisWorkbook.Worksheets(5).Cells(j, 5).Value & " | G | " & ThisWorkbook.Worksheets(5).Cells(j, 98).Value & " | " & ThisWorkbook.Worksheets(5).Cells(j, 102).Value & " " & ThisWorkbook.Worksheets(5).Cells(j, 103).Value
                        End If
                    ElseIf ThisWorkbook.Worksheets(5).Cells(j, 103).Value = "" Then
                        If Left(ThisWorkbook.Worksheets(5).Cells(j, 102).Value, 3) = "İl " Then
                            Sno = ThisWorkbook.Worksheets(5).Cells(j, 5).Value & " | G | " & ThisWorkbook.Worksheets(5).Cells(j, 98).Value & " | " & ThisWorkbook.Worksheets(5).Cells(j, 91).Value & ThisWorkbook.Worksheets(5).Cells(j, 102).Value
                        ElseIf Left(ThisWorkbook.Worksheets(5).Cells(j, 102).Value, 3) = "İlç" Then
                            Sno = ThisWorkbook.Worksheets(5).Cells(j, 5).Value & " | G | " & ThisWorkbook.Worksheets(5).Cells(j, 98).Value & " | " & ThisWorkbook.Worksheets(5).Cells(j, 92).Value & ThisWorkbook.Worksheets(5).Cells(j, 102).Value
                        Else
                            Sno = ThisWorkbook.Worksheets(5).Cells(j, 5).Value & " | G | " & ThisWorkbook.Worksheets(5).Cells(j, 98).Value & " | " & ThisWorkbook.Worksheets(5).Cells(j, 102).Value
                        End If
                    End If
                
                    Set LstBx = core_asset_manager_UI.FrameCikis.Controls.Add("Forms.ListBox.1", "Frame1LstBx" & Cont)
                    With LstBx
                        .Top = (Cont - 1) * 12
                        .Left = 18
                        .Height = 12
                        .Width = 300
                        If ThisWorkbook.Worksheets(5).Cells(j, 95).Value = ThisWorkbook.Worksheets(5).Cells(j, 176).Value Then
                            .BackColor = &H80000000 'Giriş
                        Else
                            .BackColor = &H80000003 'Mevcut
                        End If
                        .SpecialEffect = fmSpecialEffectFlat
                        .MultiSelect = fmMultiSelectMulti
                        .TextAlign = fmTextAlignLeft '02092021
                        .AddItem (Sno)
                    End With
        
                    Set LblSira1 = core_asset_manager_UI.FrameCikis.Controls.Add("Forms.Label.1", "Lbl" & Cont)
                    With LblSira1
                        .Top = (Cont - 1) * 12
                        .Left = 0
                        .Height = 12
                        .Width = 18
                        .SpecialEffect = fmSpecialEffectEtched
                        .TextAlign = fmTextAlignCenter
                        .Caption = Cont
                    End With
        
                    ScrollTakip3 = ScrollTakip3 + 12
                End If
            End If
        End If
    End If
Next j

ContTakipCikis = Cont

Son:

End Sub


