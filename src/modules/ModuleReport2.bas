Attribute VB_Name = "ModuleReport2"
Option Explicit
Public OpenWordTakip As Boolean
Public OpenWordSay As Integer
Dim IlceSakla As String
Public WsRaporNo As Worksheet, islemNew As Long
Public RnoIlkSiraBul As Range, RnoSonSiraBul As Range, RnoIlkSira As Long, RnoSonSira As Long
Public IlkSiraBulGlobal As Range, SonSiraBulGlobal As Range, IlkSiraGlobal As Long, SonSiraGlobal As Long
Public MyRngGlobal As Range, MyFinderGlobal As Range, StrAramaGlobal As String, IlkAdresGlobal As String, SonrakiAdresGlobal As String
Public StrRaporTarihiGlobal As String, RaporTireTek As Integer, SayGlobal As Long, AktarNoGlobal As Long


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

Sub Rapor2_1Tutanak1()

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
Dim ReNameUstYaziNormal As String, PaketTipi As String
Dim TxtFileUstYazi As String, TotalSayfaUstYazi As String
Dim Birimx As String, UserName As String, SourceTutanak1Farkli As String

Dim IlkSiraBulx As Range, SonSiraBulx As Range, IlkSirax As Long, SonSirax As Long, WsFarkGirisRapor1 As Worksheet



IlceSakla = ""
If InStr(Cells(ActiveCell.Row, 26).Value, " Organization A") <> 0 Then
    IlceSakla = Cells(ActiveCell.Row, 26).Value
    Cells(ActiveCell.Row, 26).Value = ""
End If

'TUTANAK1 için prosedürü başlat
'If ActiveCell.Column = 6 Then
'    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    'Application.EnableEvents = False
    
    'Pathfinder...
    AutoPath = ThisWorkbook.Path
    DestOperasyon = AutoPath & "\System Files\Operation\"
    'TUTANAK1 TANIMLARI
    SourceTutanak1Normal = AutoPath & "\System Files\System Templates\Statement 1 Templates\Statement 1.docm"
    SourceTutanak1Farkli = AutoPath & "\System Files\System Templates\Statement 1 Templates\Discrepancy Statement.docm"
    SourceDL = AutoPath & "\System Files\System Templates\Statement 1 Templates\Dispatch List.docm"
    
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
    If Not Dir(SourceTutanak1Normal, vbDirectory) <> vbNullString Then
        MsgBox "Cannot access the directory " & SourceTutanak1Normal & ". The names of folders and/or files in this directory may have been changed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    
    If Not Dir(SourceTutanak1Farkli, vbDirectory) <> vbNullString Then
        MsgBox "Cannot access the directory " & SourceTutanak1Farkli & ". The names of folders and/or files in this directory may have been changed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    
    If Not Dir(SourceDL, vbDirectory) <> vbNullString Then
        MsgBox "Cannot access the directory " & SourceDL & ". The names of folders and/or files in this directory may have been changed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If

    'Operation klasörü içinde kullanıcı modülü klasörü yoksa oluştur.
    If Not Dir(DestOpUserFolder, vbDirectory) <> vbNullString Then
        MkDir DestOpUserFolder
    End If
    
    'On Error Resume Next
    ReNameTutanak1 = Cells(ActiveCell.Row, 5).Value & "-" & Cells(6, 6).Value
    ReNameDL = Cells(ActiveCell.Row, 5).Value & "-" & "Dispatch List"
'________________________________________

    
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
    
    'Dosyayı şablondan operasyon klasörüne kopyala ve adını değiştir.
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Cells(ActiveCell.Row, 40).Value = "d. Discrepancy Detected" Then
        fso.CopyFile (SourceTutanak1Farkli), DestOpUserFolder & ReNameTutanak1 & ".docm", True
    Else
        fso.CopyFile (SourceTutanak1Normal), DestOpUserFolder & ReNameTutanak1 & ".docm", True
    End If
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
    objWord.Documents.Open FileName:=DestOpUserFolder & ReNameTutanak1 & ".docm"
    objWord.Visible = True
    objWord.Activate 'Ekrana getirir.
    'objDoc.Activate 'Ekrana getirmez.
    objWord.Application.WindowState = wdWindowStateMaximize
    Set objDoc = GetObject(DestOpUserFolder & ReNameTutanak1 & ".docm")
'________________________________________
    
    If Cells(ActiveCell.Row, 40).Value = "d. Discrepancy Detected" Then 'Farklı tutanak1 tutanağı
        'Birim
        Birimx = UCase(Replace(Replace(Worksheets(2).Cells(6, 99).Value, "i", "I"), "ı", "I")) & " UNIT"
        objDoc.Tables(1).Cell(Row:=1, Column:=1).Range.Text = Birimx
        'Tutanak1 tarihi
        objDoc.Tables(1).Cell(Row:=4, Column:=3).Range.Text = Cells(ActiveCell.Row, 39).Value
        'Geliş tarihi
        objDoc.Tables(1).Cell(Row:=5, Column:=3).Range.Text = Cells(ActiveCell.Row, 36).Value
        'By Hand/By Mail
        If Cells(ActiveCell.Row, 38).Value = "By Hand" Then
            objDoc.CheckElden.Value = True
        ElseIf Cells(ActiveCell.Row, 38).Value = "By Mail" Then
            objDoc.CheckPosta.Value = True
        End If
        'Türkçe karakterleri düzelt
        objDoc.CheckElden.Enabled = False
        objDoc.CheckElden.Enabled = True
        objDoc.CheckPosta.Enabled = False
        objDoc.CheckPosta.Enabled = True
        
        'Gelen Yazının Tarih ve Sayısı
        objDoc.Tables(1).Cell(Row:=6, Column:=3).Range.Text = Cells(ActiveCell.Row, 28).Value
        objDoc.Tables(1).Cell(Row:=6, Column:=5).Range.Text = Cells(ActiveCell.Row, 29).Value
        'Gönderen
        If Cells(ActiveCell.Row, 33).Value = "Provincial Directorate B" Then
            GelenTema = Cells(ActiveCell.Row, 25).Value & " Provincial Governorship Provincial Directorate B " & Cells(ActiveCell.Row, 34).Value
        ElseIf Cells(ActiveCell.Row, 33).Value = "District Directorate B" Then
            GelenTema = Cells(ActiveCell.Row, 26).Value & " District Governorship District Directorate B " & Cells(ActiveCell.Row, 34).Value
        ElseIf Cells(ActiveCell.Row, 33).Value = "Provincial Directorate C" Then
            GelenTema = Cells(ActiveCell.Row, 25).Value & " Provincial Governorship Provincial Directorate C " & Cells(ActiveCell.Row, 34).Value
        ElseIf Cells(ActiveCell.Row, 33).Value = "District Directorate C" Then
            GelenTema = Cells(ActiveCell.Row, 26).Value & " District Governorship District Directorate C " & Cells(ActiveCell.Row, 34).Value
        ElseIf Cells(ActiveCell.Row, 33).Value = "Provincial Directorate D" Then
            GelenTema = Cells(ActiveCell.Row, 25).Value & " Provincial Governorship Provincial Directorate D " & Cells(ActiveCell.Row, 34).Value
        ElseIf Cells(ActiveCell.Row, 33).Value = "District Directorate D" Then
            GelenTema = Cells(ActiveCell.Row, 26).Value & " District Governorship District Directorate D " & Cells(ActiveCell.Row, 34).Value
        ElseIf Cells(ActiveCell.Row, 33).Value = "Provincial Directorate E" Then
            GelenTema = Cells(ActiveCell.Row, 25).Value & " Provincial Governorship Provincial Directorate E " & Cells(ActiveCell.Row, 34).Value
        ElseIf Cells(ActiveCell.Row, 33).Value = "District Directorate E" Then
            GelenTema = Cells(ActiveCell.Row, 26).Value & " District Governorship District Directorate E " & Cells(ActiveCell.Row, 34).Value
        ElseIf InStr(Cells(ActiveCell.Row, 33).Value, "General Directorate") <> 0 Or InStr(Cells(ActiveCell.Row, 33).Value, "Regional Directorate") <> 0 Then
            If Cells(ActiveCell.Row, 34).Value <> "" Then
                GelenTema = WorksheetFunction.Proper(ThisWorkbook.Worksheets(2).Cells(6, 111).Value) & " " & Cells(ActiveCell.Row, 33).Value & " " & Cells(ActiveCell.Row, 34).Value
            Else
                GelenTema = WorksheetFunction.Proper(ThisWorkbook.Worksheets(2).Cells(6, 111).Value) & " " & Cells(ActiveCell.Row, 33).Value
            End If
        Else
            If Cells(ActiveCell.Row, 34).Value <> "" Then
                If InStr(Cells(ActiveCell.Row, 33).Value, "X.X. ") > 0 Then
                    GelenTema = Mid(Cells(ActiveCell.Row, 33).Value, 6, Len(Cells(ActiveCell.Row, 33).Value)) & " " & Cells(ActiveCell.Row, 34).Value
                Else
                    GelenTema = Cells(ActiveCell.Row, 33).Value & " " & Cells(ActiveCell.Row, 34).Value
                End If
            Else
                If InStr(Cells(ActiveCell.Row, 33).Value, "X.X. ") > 0 Then
                    GelenTema = Mid(Cells(ActiveCell.Row, 33).Value, 6, Len(Cells(ActiveCell.Row, 33).Value))
                Else
                    GelenTema = Cells(ActiveCell.Row, 33).Value
                End If
            End If
        End If
        
        objDoc.Tables(1).Cell(Row:=7, Column:=3).Range.Text = GelenTema
        
        'Gelen paket tipi(Package A/Package B/Package C)
        PaketTipi = Cells(ActiveCell.Row, 37).Value
        'Tablo başlığı
'        objDoc.Tables(4).Cell(Row:=1, Column:=1).Range.Text = PaketTipi & " İçerisinden Çıkan"
        objDoc.Tables(4).Cell(Row:=1, Column:=1).Range.Text = "Items Found Inside the " & PaketTipi
        PaketTipi = LCase(Cells(ActiveCell.Row, 37).Value)
        'Ek olarak Dispatch List
        If Cells(ActiveCell.Row, 43).Value = "Yes" Then
'            objDoc.Tables(2).Cell(Row:=1, Column:=1).Range.Text = "Yukarıda belirtilen yazı ekindeki " & _
'            PaketTipi & " açılmış ve yazıda gönderildiği belirtilen öğe(ler) ile " & _
'            PaketTipi & " içerisinden çıkan öğe(ler) arasında fark bulunduğu tespit edilmiştir. Farkı gösterir döküm aşağıda belirtilmiş olup, " & _
'            "gönderildiği belirtilen diğer öğelerin belirtildiği şekilde çıktığı tespit edilmiştir. Ayrıca, tespiti yapılan öğelerin genel dökümü ekte verilmiştir."
            objDoc.Tables(2).Cell(Row:=1, Column:=1).Range.Text = _
            "The " & PaketTipi & " attached to the document above has been opened, and a discrepancy was detected between the item(s) stated as sent and the item(s) actually found inside the " & PaketTipi & ". " & _
            "The list of differences is provided below, and the remaining item(s) mentioned in the document were confirmed to be present as described. " & _
            "Additionally, a full summary of the identified item(s) is included in the attachment."

            'Ek olarak Dispatch List
            If Cells(ActiveCell.Row, 43).Value = "Yes" Then
                If Cells(ActiveCell.Row, 44).Value > 1 Then objDoc.Tables(5).Cell(Row:=5, Column:=1).Range.Text = "Attachment: Dispatch List (" & Cells(ActiveCell.Row, 44).Value & " pages)"
                If Cells(ActiveCell.Row, 44).Value < 2 Then objDoc.Tables(5).Cell(Row:=5, Column:=1).Range.Text = "Attachment: Dispatch List (" & Cells(ActiveCell.Row, 44).Value & " page)"
            End If
        ElseIf Cells(ActiveCell.Row, 43).Value = "No" Then
'            objDoc.Tables(2).Cell(Row:=1, Column:=1).Range.Text = "Yukarıda belirtilen yazı ekindeki " & _
'            PaketTipi & " açılmış ve yazıda gönderildiği belirtilen öğe(ler) ile " & _
'            PaketTipi & " içerisinden çıkan öğe(ler) arasında fark bulunduğu tespit edilmiştir. Farkı gösterir döküm aşağıda belirtilmiştir."
            objDoc.Tables(2).Cell(Row:=1, Column:=1).Range.Text = _
            "The " & PaketTipi & " attached to the above-mentioned document has been opened, and a discrepancy has been detected between the item(s) stated as sent and the item(s) found inside the " & PaketTipi & ". " & _
            "A breakdown of the discrepancy is provided below."
        End If

        'Tabloyu doldur
        Set IlkSiraBul = Range("CM7:CM100000").Find(What:=Cells(ActiveCell.Row, 5).Value, SearchDirection:=xlNext, _
                        SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
        Set SonSiraBul = Range("CN7:CN100000").Find(What:=Cells(ActiveCell.Row, 5).Value, SearchDirection:=xlNext, _
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
            With objDoc.Tables(4)
                For i = IlkSira To SonSira - 1
                    .Rows.Add BeforeRow:=objDoc.Tables(4).Rows(4)
                    x = x + 1
                Next i
            End With
        End If
        
        For i = 4 To SonSira - IlkSira + 4
            objDoc.Tables(4).Cell(Row:=i, Column:=1).Range.Text = i - 3 'Tablo sıra no
            objDoc.Tables(4).Cell(Row:=i, Column:=2).Range.Text = Cells(IlkSira + i - 4, 46).Value 'Öğe türü
            objDoc.Tables(4).Cell(Row:=i, Column:=3).Range.Text = Cells(IlkSira + i - 4, 49).Value 'Öğe değeri
            objDoc.Tables(4).Cell(Row:=i, Column:=4).Range.Text = Cells(IlkSira + i - 4, 52).Value 'Adet
            If Cells(IlkSira + i - 4, 55).Value = "Dispatch List" Then
                objDoc.Tables(4).Cell(Row:=i, Column:=5).Range.Text = "Listed in the Dispatch List" 'Öğe ID No
            Else
                objDoc.Tables(4).Cell(Row:=i, Column:=5).Range.Text = Cells(IlkSira + i - 4, 55).Value 'Öğe ID No
            End If
            objDoc.Tables(4).Cell(Row:=i, Column:=6).Range.Text = Cells(IlkSira, 31).Value  'Tema No (Temai her satıra yaz.)
            objDoc.Tables(4).Cell(Row:=i, Column:=7).Range.Text = Cells(IlkSira + i - 4, 58).Value 'Açıklama
        Next i

        'Tabloyu doldur (Fark kısmı)
        Set WsFarkGirisRapor1 = ThisWorkbook.Worksheets(9)
        Set IlkSiraBulx = WsFarkGirisRapor1.Range("P3:P100000").Find(What:=Cells(ActiveCell.Row, 5).Value, SearchDirection:=xlNext, _
                        SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
        Set SonSiraBulx = WsFarkGirisRapor1.Range("Q3:Q100000").Find(What:=Cells(ActiveCell.Row, 5).Value, SearchDirection:=xlNext, _
                        SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
        If Not IlkSiraBulx Is Nothing Then
            IlkSirax = IlkSiraBulx.Row
        Else
            MsgBox "The serial number could not be found, so the operation cannot be completed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
            GoTo Son
        End If
        If Not SonSiraBulx Is Nothing Then
            SonSirax = SonSiraBulx.Row
        Else
            MsgBox "The serial number could not be found, so the operation cannot be completed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
            GoTo Son
        End If
        'Tabloya satır ekle
        x = 0
        If SonSirax - IlkSirax > 0 Then
            With objDoc.Tables(3)
                For i = IlkSirax To SonSirax - 1
                    .Rows.Add BeforeRow:=objDoc.Tables(3).Rows(4)
                    x = x + 1
                Next i
            End With
        End If
        
        For i = 4 To SonSirax - IlkSirax + 4
            objDoc.Tables(3).Cell(Row:=i, Column:=1).Range.Text = i - 3 'Tablo sıra no
            objDoc.Tables(3).Cell(Row:=i, Column:=2).Range.Text = WsFarkGirisRapor1.Cells(IlkSirax + i - 4, 1).Value 'Öğe türü
            objDoc.Tables(3).Cell(Row:=i, Column:=3).Range.Text = WsFarkGirisRapor1.Cells(IlkSirax + i - 4, 4).Value 'Öğe değeri
            objDoc.Tables(3).Cell(Row:=i, Column:=4).Range.Text = WsFarkGirisRapor1.Cells(IlkSirax + i - 4, 7).Value 'Adet
            If WsFarkGirisRapor1.Cells(IlkSirax + i - 4, 10).Value = "Dispatch List" Then
                objDoc.Tables(3).Cell(Row:=i, Column:=5).Range.Text = "Listed in the Dispatch List" 'Öğe ID No
            Else
                objDoc.Tables(3).Cell(Row:=i, Column:=5).Range.Text = WsFarkGirisRapor1.Cells(IlkSirax + i - 4, 10).Value 'Öğe ID No
            End If
            objDoc.Tables(3).Cell(Row:=i, Column:=6).Range.Text = Cells(IlkSira, 31).Value  'Tema No (Temai her satıra yaz.)
            objDoc.Tables(3).Cell(Row:=i, Column:=7).Range.Text = WsFarkGirisRapor1.Cells(IlkSirax + i - 4, 13).Value 'Açıklama
        Next i


        'imzalar
        objDoc.Tables(5).Cell(Row:=2, Column:=2).Range.Text = Cells(ActiveCell.Row, 112).Value 'Ad Soyad1
        objDoc.Tables(5).Cell(Row:=3, Column:=2).Range.Text = Cells(ActiveCell.Row, 113).Value 'Unvan1
        objDoc.Tables(5).Cell(Row:=2, Column:=3).Range.Text = Cells(ActiveCell.Row, 115).Value 'Ad Soyad2
        objDoc.Tables(5).Cell(Row:=3, Column:=3).Range.Text = Cells(ActiveCell.Row, 116).Value 'Unvan2
        
        'Alt bilgi ekle
        objDoc.Sections(1).Footers(wdHeaderFooterPrimary).Range.Tables(1).Cell(Row:=1, Column:=1).Range.Text = ReNameTutanak1
    
        'Sayfa sayısı kaydet komutuna bağlandı.
    '    objDoc.Close SaveChanges:=True
    '    objWord.Quit
        objDoc.Save
        'Sayfayı text dosyasından çek
        TxtFileTutanak1 = DestOpUserFolder & "Statement 1 Page Count.txt"
        #If UseADO Then
            Const adReadLine = -2&
            With CreateObject("ADODB.Stream")
                .Open
                .LoadFromFile TxtFileTutanak1
                Do Until .EOS
                    TotalSayfaTutanak1 = .ReadText(adReadLine)
                Loop
                .Close
            End With
        #Else 'UseFSO
            Const Tutanak1FarkTristateTrue = -1&
            With CreateObject("Scripting.FileSystemObject")
                With .OpenTextFile(TxtFileTutanak1, Format:=Tutanak1FarkTristateTrue)
                    Do Until .AtEndOfStream
                        TotalSayfaTutanak1 = .ReadLine
                    Loop
                    .Close
                End With
            End With
        #End If
        'MsgBox "Tutanak1: " & TotalSayfaTutanak1

        If TumDoc = True Then
            objWord.Activate
'            For i = 1 To SayPrt
'                objDoc.PrintOut
'            Next i
            objDoc.PrintOut Background:=False, Copies:=SayPrt
            'objWord.Documents.Save
            objDoc.Close SaveChanges:=False
            objWord.Visible = False
        End If

    Else 'Normal tutanak1 tutanağı
        'Birim
        Birimx = UCase(Replace(Replace(Worksheets(2).Cells(6, 99).Value, "i", "I"), "ı", "I")) & " UNIT"
        objDoc.Tables(1).Cell(Row:=1, Column:=1).Range.Text = Birimx
        'Tutanak1 tarihi
        objDoc.Tables(1).Cell(Row:=4, Column:=3).Range.Text = Cells(ActiveCell.Row, 39).Value
        'Geliş tarihi
        objDoc.Tables(1).Cell(Row:=5, Column:=3).Range.Text = Cells(ActiveCell.Row, 36).Value
        'By Hand/By Mail
        If Cells(ActiveCell.Row, 38).Value = "By Hand" Then
            objDoc.CheckElden.Value = True
        ElseIf Cells(ActiveCell.Row, 38).Value = "By Mail" Then
            objDoc.CheckPosta.Value = True
        End If
        'Öğe çıktı/çıkmadı/vb.
        If Cells(ActiveCell.Row, 40).Value = "a. Content as Expected" Then
            objDoc.CheckTam.Value = True
        ElseIf Cells(ActiveCell.Row, 40).Value = "b. Content Empty" Then
            objDoc.CheckYok.Value = True
        ElseIf Cells(ActiveCell.Row, 40).Value = "c. Only Specific Content Available" Then
            objDoc.CheckEksik.Value = True
        End If
        'Türkçe karakterleri düzelt
        objDoc.CheckElden.Enabled = False
        objDoc.CheckElden.Enabled = True
        objDoc.CheckPosta.Enabled = False
        objDoc.CheckPosta.Enabled = True
        objDoc.CheckTam.Enabled = False
        objDoc.CheckTam.Enabled = True
        objDoc.CheckYok.Enabled = False
        objDoc.CheckYok.Enabled = True
        objDoc.CheckEksik.Enabled = False
        objDoc.CheckEksik.Enabled = True
        
        'Gönderen
        'Gelen Yazının Tarih ve Sayısı
        objDoc.Tables(1).Cell(Row:=6, Column:=3).Range.Text = Cells(ActiveCell.Row, 28).Value
        objDoc.Tables(1).Cell(Row:=6, Column:=5).Range.Text = Cells(ActiveCell.Row, 29).Value
        'Gönderen
        If Cells(ActiveCell.Row, 33).Value = "Provincial Directorate B" Then
            GelenTema = Cells(ActiveCell.Row, 25).Value & " Provincial Governorship Provincial Directorate B " & Cells(ActiveCell.Row, 34).Value
        ElseIf Cells(ActiveCell.Row, 33).Value = "District Directorate B" Then
            GelenTema = Cells(ActiveCell.Row, 26).Value & " District Governorship District Directorate B " & Cells(ActiveCell.Row, 34).Value
        ElseIf Cells(ActiveCell.Row, 33).Value = "Provincial Directorate C" Then
            GelenTema = Cells(ActiveCell.Row, 25).Value & " Provincial Governorship Provincial Directorate C " & Cells(ActiveCell.Row, 34).Value
        ElseIf Cells(ActiveCell.Row, 33).Value = "District Directorate C" Then
            GelenTema = Cells(ActiveCell.Row, 26).Value & " District Governorship District Directorate C " & Cells(ActiveCell.Row, 34).Value
        ElseIf Cells(ActiveCell.Row, 33).Value = "Provincial Directorate D" Then
            GelenTema = Cells(ActiveCell.Row, 25).Value & " Provincial Governorship Provincial Directorate D " & Cells(ActiveCell.Row, 34).Value
        ElseIf Cells(ActiveCell.Row, 33).Value = "District Directorate D" Then
            GelenTema = Cells(ActiveCell.Row, 26).Value & " District Governorship District Directorate D " & Cells(ActiveCell.Row, 34).Value
        ElseIf Cells(ActiveCell.Row, 33).Value = "Provincial Directorate E" Then
            GelenTema = Cells(ActiveCell.Row, 25).Value & " Provincial Governorship Provincial Directorate E " & Cells(ActiveCell.Row, 34).Value
        ElseIf Cells(ActiveCell.Row, 33).Value = "District Directorate E" Then
            GelenTema = Cells(ActiveCell.Row, 26).Value & " District Governorship District Directorate E " & Cells(ActiveCell.Row, 34).Value
        ElseIf InStr(Cells(ActiveCell.Row, 33).Value, "General Directorate") <> 0 Or InStr(Cells(ActiveCell.Row, 33).Value, "Regional Directorate") <> 0 Then
            If Cells(ActiveCell.Row, 34).Value <> "" Then
                GelenTema = WorksheetFunction.Proper(ThisWorkbook.Worksheets(2).Cells(6, 111).Value) & " " & Cells(ActiveCell.Row, 33).Value & " " & Cells(ActiveCell.Row, 34).Value
            Else
                GelenTema = WorksheetFunction.Proper(ThisWorkbook.Worksheets(2).Cells(6, 111).Value) & " " & Cells(ActiveCell.Row, 33).Value
            End If
        Else
            If Cells(ActiveCell.Row, 34).Value <> "" Then
                If InStr(Cells(ActiveCell.Row, 33).Value, "X.X. ") > 0 Then
                    GelenTema = Mid(Cells(ActiveCell.Row, 33).Value, 6, Len(Cells(ActiveCell.Row, 33).Value)) & " " & Cells(ActiveCell.Row, 34).Value
                Else
                    GelenTema = Cells(ActiveCell.Row, 33).Value & " " & Cells(ActiveCell.Row, 34).Value
                End If
            Else
                If InStr(Cells(ActiveCell.Row, 33).Value, "X.X. ") > 0 Then
                    GelenTema = Mid(Cells(ActiveCell.Row, 33).Value, 6, Len(Cells(ActiveCell.Row, 33).Value))
                Else
                    GelenTema = Cells(ActiveCell.Row, 33).Value
                End If
            End If
        End If
        
        objDoc.Tables(1).Cell(Row:=9, Column:=3).Range.Text = GelenTema
        
        'Gönderenin adresi
'        If Cells(ActiveCell.Row, 18).Value = Cells(ActiveCell.Row, 17).Value & " Organization A" Then
        If Cells(ActiveCell.Row, 26).Value <> "" Then
            objDoc.Tables(1).Cell(Row:=10, Column:=3).Range.Text = Cells(ActiveCell.Row, 26).Value & "/" & Cells(ActiveCell.Row, 25).Value
        Else
            objDoc.Tables(1).Cell(Row:=10, Column:=3).Range.Text = Cells(ActiveCell.Row, 25).Value
        End If

        'Gelen Yazının Tarih ve Sayısı
        objDoc.Tables(1).Cell(Row:=11, Column:=3).Range.Text = Cells(ActiveCell.Row, 28).Value
        objDoc.Tables(1).Cell(Row:=11, Column:=5).Range.Text = Cells(ActiveCell.Row, 29).Value
        'Gönderinin Eki
        If Cells(ActiveCell.Row, 37).Value = "Package A" Then
            objDoc.CheckZarf.Value = True
        ElseIf Cells(ActiveCell.Row, 37).Value = "Package B" Then
            objDoc.CheckTorba.Value = True
        ElseIf Cells(ActiveCell.Row, 37).Value = "Package C" Then
            objDoc.CheckKoli.Value = True
        End If
        'Tabloyu doldur
        Set IlkSiraBul = Range("CM7:CM100000").Find(What:=Cells(ActiveCell.Row, 5).Value, SearchDirection:=xlNext, _
                        SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
        Set SonSiraBul = Range("CN7:CN100000").Find(What:=Cells(ActiveCell.Row, 5).Value, SearchDirection:=xlNext, _
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
            objDoc.Tables(2).Cell(Row:=i, Column:=2).Range.Text = Cells(IlkSira + i - 2, 46).Value 'Öğe türü
            objDoc.Tables(2).Cell(Row:=i, Column:=3).Range.Text = Cells(IlkSira + i - 2, 49).Value 'Öğe değeri
            objDoc.Tables(2).Cell(Row:=i, Column:=4).Range.Text = Cells(IlkSira + i - 2, 52).Value 'Adet
            If Cells(IlkSira + i - 2, 55).Value = "Dispatch List" Then
                objDoc.Tables(2).Cell(Row:=i, Column:=5).Range.Text = "Listed in the Dispatch List" 'Öğe ID No
            Else
                objDoc.Tables(2).Cell(Row:=i, Column:=5).Range.Text = Cells(IlkSira + i - 2, 55).Value 'Öğe ID No
            End If
            objDoc.Tables(2).Cell(Row:=i, Column:=6).Range.Text = Cells(IlkSira, 31).Value  'Tema No (Temai her satıra yaz.)
            objDoc.Tables(2).Cell(Row:=i, Column:=7).Range.Text = Cells(IlkSira + i - 2, 58).Value 'Açıklama
        Next i
        'Ek olarak Dispatch List
        If Cells(ActiveCell.Row, 43).Value = "Yes" Then
            If Cells(ActiveCell.Row, 43).Value = "Yes" Then
                If Cells(ActiveCell.Row, 44).Value > 1 Then objDoc.Tables(5).Cell(Row:=5, Column:=1).Range.Text = "Attachment: Dispatch List (" & Cells(ActiveCell.Row, 44).Value & " pages)"
                If Cells(ActiveCell.Row, 44).Value < 2 Then objDoc.Tables(5).Cell(Row:=5, Column:=1).Range.Text = "Attachment: Dispatch List (" & Cells(ActiveCell.Row, 44).Value & " page)"
            End If
        ElseIf Cells(ActiveCell.Row, 43).Value = "No" Then
            objDoc.Tables(4).Cell(Row:=5, Column:=1).Range.Text = Cells(ActiveCell.Row, 44).Value
        End If

        'imzalar
        objDoc.Tables(4).Cell(Row:=2, Column:=2).Range.Text = Cells(ActiveCell.Row, 112).Value 'Ad Soyad1
        objDoc.Tables(4).Cell(Row:=3, Column:=2).Range.Text = Cells(ActiveCell.Row, 113).Value 'Unvan1
        objDoc.Tables(4).Cell(Row:=2, Column:=3).Range.Text = Cells(ActiveCell.Row, 115).Value 'Ad Soyad2
        objDoc.Tables(4).Cell(Row:=3, Column:=3).Range.Text = Cells(ActiveCell.Row, 116).Value 'Unvan2
        
        'Alt bilgi ekle
        objDoc.Sections(1).Footers(wdHeaderFooterPrimary).Range.Tables(1).Cell(Row:=1, Column:=1).Range.Text = ReNameTutanak1
    
        'İmza boşluğunu sayfaya sığdırmak için düzenle
        If x > 9 And x < 14 Then
            For i = 1 To x - 9
                If i = 1 Then
                    objDoc.Tables(3).Rows(1).Delete
                ElseIf i = 2 Then
                    For j = 1 To 2
                        objDoc.Tables(3).Rows(1).Delete
                    Next j
                ElseIf i = 3 Then
                    For j = 1 To 2
                        objDoc.Tables(4).Rows(1).Delete
                    Next j
                ElseIf i = 4 Then
                    objDoc.Tables(3).Rows.Add BeforeRow:=objDoc.Tables(3).Rows(1)
                    For j = 1 To 2
                        objDoc.Tables(4).Rows(1).Delete
                    Next j
                End If
            Next i
        ElseIf x > 13 Then
            'Nothing
        End If
        
        'Sayfa sayısı kaydet komutuna bağlandı.
    '    objDoc.Close SaveChanges:=True
    '    objWord.Quit
        objDoc.Save
        'Sayfayı text dosyasından çek
        TxtFileTutanak1 = DestOpUserFolder & "Statement 1 Page Count.txt"
        #If UseADO Then
            Const adReadLine = -2&
            With CreateObject("ADODB.Stream")
                .Open
                .LoadFromFile TxtFileTutanak1
                Do Until .EOS
                    TotalSayfaTutanak1 = .ReadText(adReadLine)
                Loop
                .Close
            End With
        #Else 'UseFSO
            Const Tutanak1TristateTrue = -1&
            With CreateObject("Scripting.FileSystemObject")
                With .OpenTextFile(TxtFileTutanak1, Format:=Tutanak1TristateTrue)
                    Do Until .AtEndOfStream
                        TotalSayfaTutanak1 = .ReadLine
                    Loop
                    .Close
                End With
            End With
        #End If
        'MsgBox "Tutanak1: " & TotalSayfaTutanak1
    
        If TumDoc = True Then
            objWord.Activate
'            For i = 1 To SayPrt
'                objDoc.PrintOut
'            Next i
            objDoc.PrintOut Background:=False, Copies:=SayPrt
            'objWord.Documents.Save
            objDoc.Close SaveChanges:=False
            objWord.Visible = False
        End If
    End If
    
    'Dispatch List oluşturulacak.
    If Cells(ActiveCell.Row, 43).Value = "Yes" Then
        'Döküm için sayfayı belirt.(Dispatch List word daosyası bu veriyi işleyecek.)
        Set fso = CreateObject("Scripting.FileSystemObject")
        DokumSayfaGonder = Cells(ActiveCell.Row, 44).Value
        Set DokumFileTxt = fso.CreateTextFile(DestOpUserFolder & "Send Dispatch Page Count.txt", True, True)
        DokumFileTxt.Write DokumSayfaGonder
        DokumFileTxt.Close

        'Dosyayı şablondan operasyon klasörüne kopyala ve adını değiştir.
        Set fso = CreateObject("Scripting.FileSystemObject")
        fso.CopyFile (SourceDL), DestOpUserFolder & ReNameDL & ".docm", True
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
        objWord.Documents.Open FileName:=DestOpUserFolder & ReNameDL & ".docm"
        objWord.Visible = True
        objWord.Activate 'Ekrana getirir.
        'objDoc.Activate 'Ekrana getirmez.
        objWord.Application.WindowState = wdWindowStateMaximize
        Set objDoc = GetObject(DestOpUserFolder & ReNameDL & ".docm")
    '________________________________________
        
        'Birim
        Birimx = UCase(Replace(Replace(Worksheets(2).Cells(6, 99).Value, "i", "I"), "ı", "I")) & " UNIT"
        objDoc.Tables(1).Cell(Row:=1, Column:=1).Range.Text = Birimx
        'Gönderen
        objDoc.Tables(2).Cell(Row:=3, Column:=3).Range.Text = GelenTema
        'Belge tarihi
        objDoc.Tables(2).Cell(Row:=4, Column:=3).Range.Text = Cells(ActiveCell.Row, 28).Value
        'Belge sayısı
        objDoc.Tables(2).Cell(Row:=4, Column:=5).Range.Text = Cells(ActiveCell.Row, 29).Value

        'Alt bilgi ekle
        objDoc.Sections(1).Footers(wdHeaderFooterPrimary).Range.Tables(1).Cell(Row:=1, Column:=1).Range.Text = ReNameDL
    
        'Sayfa sayısı kaydet komutuna bağlandı.
        objDoc.Save
        'Sayfayı text dosyasından çek
        TxtFileDokum = DestOpUserFolder & "Dispatch Page Count.txt"
        #If UseADO Then
            Const adReadLine = -2&
            With CreateObject("ADODB.Stream")
                .Open
                .LoadFromFile TxtFileDokum
                Do Until .EOS
                    TotalSayfaDokum = .ReadText(adReadLine)
                Loop
                .Close
            End With
        #Else 'UseFSO
            Const DokumTristateTrue = -1&
            With CreateObject("Scripting.FileSystemObject")
                With .OpenTextFile(TxtFileDokum, Format:=DokumTristateTrue)
                    Do Until .AtEndOfStream
                        TotalSayfaDokum = .ReadLine
                    Loop
                    .Close
                End With
            End With
        #End If
        'MsgBox "Tutanak1: " & TotalSayfaTutanak1
        'MsgBox "Döküm: " & TotalSayfaDokum
        If TumDoc = True Then
            objWord.Activate
'            For i = 1 To 1 'SayPrt
'                objDoc.PrintOut
'            Next i
            objDoc.PrintOut Background:=False, Copies:=SayPrt
            'objWord.Documents.Save
            objDoc.Close SaveChanges:=False
            objWord.Visible = False
        End If
    End If
    
    'Tutanak1 ve Döküm sayfa sayısı
    Cells(ActiveCell.Row, 97).Value = TotalSayfaTutanak1
    Cells(ActiveCell.Row, 98).Value = TotalSayfaDokum
    
    'Nesneleri temizle
    Set fso = Nothing
    Set objWord = Nothing
    Set objDoc = Nothing

Son:

'Nesneleri temizle
Set fso = Nothing
Set objWord = Nothing
Set objDoc = Nothing

If IlceSakla <> "" Then
    Cells(ActiveCell.Row, 26).Value = IlceSakla
End If

'Worksheets(4).Activate
Application.ScreenUpdating = True
Application.DisplayAlerts = True
'Application.EnableEvents = True

'End If

End Sub

Sub Rapor2_2Tutanak1()

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
Dim ReNameUstYaziNormal As String, PaketTipi As String
Dim TxtFileUstYazi As String, TotalSayfaUstYazi As String
Dim Birimx As String, UserName As String, SourceTutanak1Farkli As String

Dim IlkSiraBulx As Range, SonSiraBulx As Range, IlkSirax As Long, SonSirax As Long, WsFarkGirisRapor1 As Worksheet

IlceSakla = ""
If InStr(Cells(ActiveCell.Row, 26).Value, " Organization A") <> 0 Then
    IlceSakla = Cells(ActiveCell.Row, 26).Value
    Cells(ActiveCell.Row, 26).Value = ""
End If

'TUTANAK1 için prosedürü başlat
'If ActiveCell.Column = 6 Then
'    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    'Application.EnableEvents = False
    
    'Pathfinder...
    AutoPath = ThisWorkbook.Path
    DestOperasyon = AutoPath & "\System Files\Operation\"
    'TUTANAK1 TANIMLARI
    SourceTutanak1Normal = AutoPath & "\System Files\System Templates\Statement 1 Templates\Statement 1.docm"
    SourceTutanak1Farkli = AutoPath & "\System Files\System Templates\Statement 1 Templates\Discrepancy Statement.docm"
    SourceDL = AutoPath & "\System Files\System Templates\Statement 1 Templates\Dispatch List.docm"
    
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
    If Not Dir(SourceTutanak1Normal, vbDirectory) <> vbNullString Then
        MsgBox "Cannot access the directory " & SourceTutanak1Normal & ". The names of folders and/or files in this directory may have been changed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    
    If Not Dir(SourceTutanak1Farkli, vbDirectory) <> vbNullString Then
        MsgBox "Cannot access the directory " & SourceTutanak1Farkli & ". The names of folders and/or files in this directory may have been changed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    
    If Not Dir(SourceDL, vbDirectory) <> vbNullString Then
        MsgBox "Cannot access the directory " & SourceDL & ". The names of folders and/or files in this directory may have been changed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If

    'Operation klasörü içinde kullanıcı modülü klasörü yoksa oluştur.
    If Not Dir(DestOpUserFolder, vbDirectory) <> vbNullString Then
        MkDir DestOpUserFolder
    End If
    
    'On Error Resume Next
    ReNameTutanak1 = Cells(ActiveCell.Row, 5).Value & "-" & Cells(6, 13).Value
    ReNameDL = Cells(ActiveCell.Row, 5).Value & "-" & "Dispatch List"
'________________________________________

    
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
    
    'Dosyayı şablondan operasyon klasörüne kopyala ve adını değiştir.
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Cells(ActiveCell.Row, 40).Value = "d. Discrepancy Detected" Then
        fso.CopyFile (SourceTutanak1Farkli), DestOpUserFolder & ReNameTutanak1 & ".docm", True
    Else
        fso.CopyFile (SourceTutanak1Normal), DestOpUserFolder & ReNameTutanak1 & ".docm", True
    End If
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
    objWord.Documents.Open FileName:=DestOpUserFolder & ReNameTutanak1 & ".docm"
    objWord.Visible = True
    objWord.Activate 'Ekrana getirir.
    'objDoc.Activate 'Ekrana getirmez.
    objWord.Application.WindowState = wdWindowStateMaximize
    Set objDoc = GetObject(DestOpUserFolder & ReNameTutanak1 & ".docm")
'________________________________________
    
    If Cells(ActiveCell.Row, 40).Value = "d. Discrepancy Detected" Then 'Farklı tutanak1 tutanağı
        'Birim
        Birimx = UCase(Replace(Replace(Worksheets(2).Cells(6, 99).Value, "i", "I"), "ı", "I")) & " UNIT"
        objDoc.Tables(1).Cell(Row:=1, Column:=1).Range.Text = Birimx
        'Tutanak1 tarihi
        objDoc.Tables(1).Cell(Row:=4, Column:=3).Range.Text = Cells(ActiveCell.Row, 39).Value
        'Geliş tarihi
        objDoc.Tables(1).Cell(Row:=5, Column:=3).Range.Text = Cells(ActiveCell.Row, 36).Value
        'By Hand/By Mail
        If Cells(ActiveCell.Row, 38).Value = "By Hand" Then
            objDoc.CheckElden.Value = True
        ElseIf Cells(ActiveCell.Row, 38).Value = "By Mail" Then
            objDoc.CheckPosta.Value = True
        End If
        'Türkçe karakterleri düzelt
        objDoc.CheckElden.Enabled = False
        objDoc.CheckElden.Enabled = True
        objDoc.CheckPosta.Enabled = False
        objDoc.CheckPosta.Enabled = True
        
        'Gelen Yazının Tarih ve Sayısı
        objDoc.Tables(1).Cell(Row:=6, Column:=3).Range.Text = Cells(ActiveCell.Row, 28).Value
        objDoc.Tables(1).Cell(Row:=6, Column:=5).Range.Text = Cells(ActiveCell.Row, 29).Value
        'Gönderen
        If Cells(ActiveCell.Row, 33).Value = "Provincial Directorate B" Then
            GelenTema = Cells(ActiveCell.Row, 25).Value & " Provincial Governorship Provincial Directorate B " & Cells(ActiveCell.Row, 34).Value
        ElseIf Cells(ActiveCell.Row, 33).Value = "District Directorate B" Then
            GelenTema = Cells(ActiveCell.Row, 26).Value & " District Governorship District Directorate B " & Cells(ActiveCell.Row, 34).Value
        ElseIf Cells(ActiveCell.Row, 33).Value = "Provincial Directorate C" Then
            GelenTema = Cells(ActiveCell.Row, 25).Value & " Provincial Governorship Provincial Directorate C " & Cells(ActiveCell.Row, 34).Value
        ElseIf Cells(ActiveCell.Row, 33).Value = "District Directorate C" Then
            GelenTema = Cells(ActiveCell.Row, 26).Value & " District Governorship District Directorate C " & Cells(ActiveCell.Row, 34).Value
        ElseIf Cells(ActiveCell.Row, 33).Value = "Provincial Directorate D" Then
            GelenTema = Cells(ActiveCell.Row, 25).Value & " Provincial Governorship Provincial Directorate D " & Cells(ActiveCell.Row, 34).Value
        ElseIf Cells(ActiveCell.Row, 33).Value = "District Directorate D" Then
            GelenTema = Cells(ActiveCell.Row, 26).Value & " District Governorship District Directorate D " & Cells(ActiveCell.Row, 34).Value
        ElseIf Cells(ActiveCell.Row, 33).Value = "Provincial Directorate E" Then
            GelenTema = Cells(ActiveCell.Row, 25).Value & " Provincial Governorship Provincial Directorate E " & Cells(ActiveCell.Row, 34).Value
        ElseIf Cells(ActiveCell.Row, 33).Value = "District Directorate E" Then
            GelenTema = Cells(ActiveCell.Row, 26).Value & " District Governorship District Directorate E " & Cells(ActiveCell.Row, 34).Value
        ElseIf InStr(Cells(ActiveCell.Row, 33).Value, "General Directorate") <> 0 Or InStr(Cells(ActiveCell.Row, 33).Value, "Regional Directorate") <> 0 Then
            If Cells(ActiveCell.Row, 34).Value <> "" Then
                GelenTema = WorksheetFunction.Proper(ThisWorkbook.Worksheets(2).Cells(6, 111).Value) & " " & Cells(ActiveCell.Row, 33).Value & " " & Cells(ActiveCell.Row, 34).Value
            Else
                GelenTema = WorksheetFunction.Proper(ThisWorkbook.Worksheets(2).Cells(6, 111).Value) & " " & Cells(ActiveCell.Row, 33).Value
            End If
        Else
            If Cells(ActiveCell.Row, 34).Value <> "" Then
                If InStr(Cells(ActiveCell.Row, 33).Value, "X.X. ") > 0 Then
                    GelenTema = Mid(Cells(ActiveCell.Row, 33).Value, 6, Len(Cells(ActiveCell.Row, 33).Value)) & " " & Cells(ActiveCell.Row, 34).Value
                Else
                    GelenTema = Cells(ActiveCell.Row, 33).Value & " " & Cells(ActiveCell.Row, 34).Value
                End If
            Else
                If InStr(Cells(ActiveCell.Row, 33).Value, "X.X. ") > 0 Then
                    GelenTema = Mid(Cells(ActiveCell.Row, 33).Value, 6, Len(Cells(ActiveCell.Row, 33).Value))
                Else
                    GelenTema = Cells(ActiveCell.Row, 33).Value
                End If
            End If
        End If
        
        objDoc.Tables(1).Cell(Row:=7, Column:=3).Range.Text = GelenTema
        
        'Gelen paket tipi(Package A/Package B/Package C)
        PaketTipi = Cells(ActiveCell.Row, 37).Value
        'Tablo başlığı
'        objDoc.Tables(4).Cell(Row:=1, Column:=1).Range.Text = PaketTipi & " İçerisinden Çıkan"
        objDoc.Tables(4).Cell(Row:=1, Column:=1).Range.Text = "Items Found Inside the " & PaketTipi
        PaketTipi = LCase(Cells(ActiveCell.Row, 37).Value)
        'Ek olarak Dispatch List
        If Cells(ActiveCell.Row, 43).Value = "Yes" Then
'            objDoc.Tables(2).Cell(Row:=1, Column:=1).Range.Text = "Yukarıda belirtilen yazı ekindeki " & _
'            PaketTipi & " açılmış ve yazıda gönderildiği belirtilen öğe(ler) ile " & _
'            PaketTipi & " içerisinden çıkan öğe(ler) arasında fark bulunduğu tespit edilmiştir. Farkı gösterir döküm aşağıda belirtilmiş olup, " & _
'            "gönderildiği belirtilen diğer öğelerin belirtildiği şekilde çıktığı tespit edilmiştir. Ayrıca, tespiti yapılan öğelerin genel dökümü ekte verilmiştir."
            objDoc.Tables(2).Cell(Row:=1, Column:=1).Range.Text = _
            "The " & PaketTipi & " attached to the document above has been opened, and a discrepancy was detected between the item(s) stated as sent and the item(s) actually found inside the " & PaketTipi & ". " & _
            "The list of differences is provided below, and the remaining item(s) mentioned in the document were confirmed to be present as described. " & _
            "Additionally, a full summary of the identified item(s) is included in the attachment."
            If Cells(ActiveCell.Row, 44).Value > 1 Then objDoc.Tables(5).Cell(Row:=5, Column:=1).Range.Text = "Attachment: Dispatch List (" & Cells(ActiveCell.Row, 44).Value & " pages)"
            If Cells(ActiveCell.Row, 44).Value < 2 Then objDoc.Tables(5).Cell(Row:=5, Column:=1).Range.Text = "Attachment: Dispatch List (" & Cells(ActiveCell.Row, 44).Value & " page)"
        ElseIf Cells(ActiveCell.Row, 43).Value = "No" Then
'            objDoc.Tables(2).Cell(Row:=1, Column:=1).Range.Text = "Yukarıda belirtilen yazı ekindeki " & _
'            PaketTipi & " açılmış ve yazıda gönderildiği belirtilen öğe(ler) ile " & _
'            PaketTipi & " içerisinden çıkan öğe(ler) arasında fark bulunduğu tespit edilmiştir. Farkı gösterir döküm aşağıda belirtilmiştir."
            objDoc.Tables(2).Cell(Row:=1, Column:=1).Range.Text = _
            "The " & PaketTipi & " attached to the above-mentioned document has been opened, and a discrepancy has been detected between the item(s) stated as sent and the item(s) found inside the " & PaketTipi & ". " & _
            "A breakdown of the discrepancy is provided below."
        End If


        'Tabloyu doldur
        Set IlkSiraBul = Range("CM7:CM100000").Find(What:=Cells(ActiveCell.Row, 5).Value, SearchDirection:=xlNext, _
                        SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
        Set SonSiraBul = Range("CN7:CN100000").Find(What:=Cells(ActiveCell.Row, 5).Value, SearchDirection:=xlNext, _
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
            With objDoc.Tables(4)
                For i = IlkSira To SonSira - 1
                    .Rows.Add BeforeRow:=objDoc.Tables(4).Rows(4)
                    x = x + 1
                Next i
            End With
        End If
        
        For i = 4 To SonSira - IlkSira + 4
            objDoc.Tables(4).Cell(Row:=i, Column:=1).Range.Text = i - 3 'Tablo sıra no
            objDoc.Tables(4).Cell(Row:=i, Column:=2).Range.Text = Cells(IlkSira + i - 4, 46).Value 'Öğe türü
            objDoc.Tables(4).Cell(Row:=i, Column:=3).Range.Text = Cells(IlkSira + i - 4, 49).Value 'Öğe değeri
            objDoc.Tables(4).Cell(Row:=i, Column:=4).Range.Text = Cells(IlkSira + i - 4, 52).Value 'Adet
            If Cells(IlkSira + i - 4, 55).Value = "Dispatch List" Then
                objDoc.Tables(4).Cell(Row:=i, Column:=5).Range.Text = "Listed in the Dispatch List" 'Öğe ID No
            Else
                objDoc.Tables(4).Cell(Row:=i, Column:=5).Range.Text = Cells(IlkSira + i - 4, 55).Value 'Öğe ID No
            End If
            objDoc.Tables(4).Cell(Row:=i, Column:=6).Range.Text = Cells(IlkSira, 31).Value  'Tema No (Temai her satıra yaz.)
            objDoc.Tables(4).Cell(Row:=i, Column:=7).Range.Text = Cells(IlkSira + i - 4, 58).Value 'Açıklama
        Next i
        
        'Tabloyu doldur (Fark kısmı)
        Set WsFarkGirisRapor1 = ThisWorkbook.Worksheets(9)
        Set IlkSiraBulx = WsFarkGirisRapor1.Range("P3:P100000").Find(What:=Cells(ActiveCell.Row, 5).Value, SearchDirection:=xlNext, _
                        SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
        Set SonSiraBulx = WsFarkGirisRapor1.Range("Q3:Q100000").Find(What:=Cells(ActiveCell.Row, 5).Value, SearchDirection:=xlNext, _
                        SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
        If Not IlkSiraBulx Is Nothing Then
            IlkSirax = IlkSiraBulx.Row
        Else
            MsgBox "The serial number could not be found, so the operation cannot be completed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
            GoTo Son
        End If
        If Not SonSiraBulx Is Nothing Then
            SonSirax = SonSiraBulx.Row
        Else
            MsgBox "The serial number could not be found, so the operation cannot be completed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
            GoTo Son
        End If
        'Tabloya satır ekle
        x = 0
        If SonSirax - IlkSirax > 0 Then
            With objDoc.Tables(3)
                For i = IlkSirax To SonSirax - 1
                    .Rows.Add BeforeRow:=objDoc.Tables(3).Rows(4)
                    x = x + 1
                Next i
            End With
        End If
        
        For i = 4 To SonSirax - IlkSirax + 4
            objDoc.Tables(3).Cell(Row:=i, Column:=1).Range.Text = i - 3 'Tablo sıra no
            objDoc.Tables(3).Cell(Row:=i, Column:=2).Range.Text = WsFarkGirisRapor1.Cells(IlkSirax + i - 4, 1).Value 'Öğe türü
            objDoc.Tables(3).Cell(Row:=i, Column:=3).Range.Text = WsFarkGirisRapor1.Cells(IlkSirax + i - 4, 4).Value 'Öğe değeri
            objDoc.Tables(3).Cell(Row:=i, Column:=4).Range.Text = WsFarkGirisRapor1.Cells(IlkSirax + i - 4, 7).Value 'Adet
            If WsFarkGirisRapor1.Cells(IlkSirax + i - 4, 10).Value = "Dispatch List" Then
                objDoc.Tables(3).Cell(Row:=i, Column:=5).Range.Text = "Listed in the Dispatch List" 'Öğe ID No
            Else
                objDoc.Tables(3).Cell(Row:=i, Column:=5).Range.Text = WsFarkGirisRapor1.Cells(IlkSirax + i - 4, 10).Value 'Öğe ID No
            End If
            objDoc.Tables(3).Cell(Row:=i, Column:=6).Range.Text = Cells(IlkSira, 31).Value  'Tema No (Temai her satıra yaz.)
            objDoc.Tables(3).Cell(Row:=i, Column:=7).Range.Text = WsFarkGirisRapor1.Cells(IlkSirax + i - 4, 13).Value 'Açıklama
        Next i
        
        
        'imzalar
        objDoc.Tables(5).Cell(Row:=2, Column:=2).Range.Text = Cells(ActiveCell.Row, 112).Value 'Ad Soyad1
        objDoc.Tables(5).Cell(Row:=3, Column:=2).Range.Text = Cells(ActiveCell.Row, 113).Value 'Unvan1
        objDoc.Tables(5).Cell(Row:=2, Column:=3).Range.Text = Cells(ActiveCell.Row, 115).Value 'Ad Soyad2
        objDoc.Tables(5).Cell(Row:=3, Column:=3).Range.Text = Cells(ActiveCell.Row, 116).Value 'Unvan2
        
        'Alt bilgi ekle
        objDoc.Sections(1).Footers(wdHeaderFooterPrimary).Range.Tables(1).Cell(Row:=1, Column:=1).Range.Text = ReNameTutanak1
    
        'Sayfa sayısı kaydet komutuna bağlandı.
    '    objDoc.Close SaveChanges:=True
    '    objWord.Quit
        objDoc.Save
        'Sayfayı text dosyasından çek
        TxtFileTutanak1 = DestOpUserFolder & "Statement 1 Page Count.txt"
        #If UseADO Then
            Const adReadLine = -2&
            With CreateObject("ADODB.Stream")
                .Open
                .LoadFromFile TxtFileTutanak1
                Do Until .EOS
                    TotalSayfaTutanak1 = .ReadText(adReadLine)
                Loop
                .Close
            End With
        #Else 'UseFSO
            Const Tutanak1FarkTristateTrue = -1&
            With CreateObject("Scripting.FileSystemObject")
                With .OpenTextFile(TxtFileTutanak1, Format:=Tutanak1FarkTristateTrue)
                    Do Until .AtEndOfStream
                        TotalSayfaTutanak1 = .ReadLine
                    Loop
                    .Close
                End With
            End With
        #End If
        'MsgBox "Tutanak1: " & TotalSayfaTutanak1

        If TumDoc = True Then
            objWord.Activate
'            For i = 1 To SayPrt
'                objDoc.PrintOut
'            Next i
            objDoc.PrintOut Background:=False, Copies:=SayPrt
            'objWord.Documents.Save
            objDoc.Close SaveChanges:=False
            objWord.Visible = False
        End If

    
    Else 'Normal tutanak1 tutanağı
        'Birim
        Birimx = UCase(Replace(Replace(Worksheets(2).Cells(6, 99).Value, "i", "I"), "ı", "I")) & " UNIT"
        objDoc.Tables(1).Cell(Row:=1, Column:=1).Range.Text = Birimx
        'Tutanak1 tarihi
        objDoc.Tables(1).Cell(Row:=4, Column:=3).Range.Text = Cells(ActiveCell.Row, 39).Value
        'Geliş tarihi
        objDoc.Tables(1).Cell(Row:=5, Column:=3).Range.Text = Cells(ActiveCell.Row, 36).Value
        'By Hand/By Mail
        If Cells(ActiveCell.Row, 38).Value = "By Hand" Then
            objDoc.CheckElden.Value = True
        ElseIf Cells(ActiveCell.Row, 38).Value = "By Mail" Then
            objDoc.CheckPosta.Value = True
        End If
        'Öğe çıktı/çıkmadı/vb.
        If Cells(ActiveCell.Row, 40).Value = "a. Content as Expected" Then
            objDoc.CheckTam.Value = True
        ElseIf Cells(ActiveCell.Row, 40).Value = "b. Content Empty" Then
            objDoc.CheckYok.Value = True
        ElseIf Cells(ActiveCell.Row, 40).Value = "c. Only Specific Content Available" Then
            objDoc.CheckEksik.Value = True
        End If
        'Türkçe karakterleri düzelt
        objDoc.CheckElden.Enabled = False
        objDoc.CheckElden.Enabled = True
        objDoc.CheckPosta.Enabled = False
        objDoc.CheckPosta.Enabled = True
        objDoc.CheckTam.Enabled = False
        objDoc.CheckTam.Enabled = True
        objDoc.CheckYok.Enabled = False
        objDoc.CheckYok.Enabled = True
        objDoc.CheckEksik.Enabled = False
        objDoc.CheckEksik.Enabled = True
        
        'Gönderen
        If Cells(ActiveCell.Row, 33).Value = "Provincial Directorate B" Then
            GelenTema = Cells(ActiveCell.Row, 25).Value & " Provincial Governorship Provincial Directorate B " & Cells(ActiveCell.Row, 34).Value
        ElseIf Cells(ActiveCell.Row, 33).Value = "District Directorate B" Then
            GelenTema = Cells(ActiveCell.Row, 26).Value & " District Governorship District Directorate B " & Cells(ActiveCell.Row, 34).Value
        ElseIf Cells(ActiveCell.Row, 33).Value = "Provincial Directorate C" Then
            GelenTema = Cells(ActiveCell.Row, 25).Value & " Provincial Governorship Provincial Directorate C " & Cells(ActiveCell.Row, 34).Value
        ElseIf Cells(ActiveCell.Row, 33).Value = "District Directorate C" Then
            GelenTema = Cells(ActiveCell.Row, 26).Value & " District Governorship District Directorate C " & Cells(ActiveCell.Row, 34).Value
        ElseIf Cells(ActiveCell.Row, 33).Value = "Provincial Directorate D" Then
            GelenTema = Cells(ActiveCell.Row, 25).Value & " Provincial Governorship Provincial Directorate D " & Cells(ActiveCell.Row, 34).Value
        ElseIf Cells(ActiveCell.Row, 33).Value = "District Directorate D" Then
            GelenTema = Cells(ActiveCell.Row, 26).Value & " District Governorship District Directorate D " & Cells(ActiveCell.Row, 34).Value
        ElseIf Cells(ActiveCell.Row, 33).Value = "Provincial Directorate E" Then
            GelenTema = Cells(ActiveCell.Row, 25).Value & " Provincial Governorship Provincial Directorate E " & Cells(ActiveCell.Row, 34).Value
        ElseIf Cells(ActiveCell.Row, 33).Value = "District Directorate E" Then
            GelenTema = Cells(ActiveCell.Row, 26).Value & " District Governorship District Directorate E " & Cells(ActiveCell.Row, 34).Value
        ElseIf InStr(Cells(ActiveCell.Row, 33).Value, "General Directorate") <> 0 Or InStr(Cells(ActiveCell.Row, 33).Value, "Regional Directorate") <> 0 Then
            If Cells(ActiveCell.Row, 34).Value <> "" Then
                GelenTema = WorksheetFunction.Proper(ThisWorkbook.Worksheets(2).Cells(6, 111).Value) & " " & Cells(ActiveCell.Row, 33).Value & " " & Cells(ActiveCell.Row, 34).Value
            Else
                GelenTema = WorksheetFunction.Proper(ThisWorkbook.Worksheets(2).Cells(6, 111).Value) & " " & Cells(ActiveCell.Row, 33).Value
            End If
        Else
            If Cells(ActiveCell.Row, 34).Value <> "" Then
                If InStr(Cells(ActiveCell.Row, 33).Value, "X.X. ") > 0 Then
                    GelenTema = Mid(Cells(ActiveCell.Row, 33).Value, 6, Len(Cells(ActiveCell.Row, 33).Value)) & " " & Cells(ActiveCell.Row, 34).Value
                Else
                    GelenTema = Cells(ActiveCell.Row, 33).Value & " " & Cells(ActiveCell.Row, 34).Value
                End If
            Else
                If InStr(Cells(ActiveCell.Row, 33).Value, "X.X. ") > 0 Then
                    GelenTema = Mid(Cells(ActiveCell.Row, 33).Value, 6, Len(Cells(ActiveCell.Row, 33).Value))
                Else
                    GelenTema = Cells(ActiveCell.Row, 33).Value
                End If
            End If
        End If
        
        objDoc.Tables(1).Cell(Row:=9, Column:=3).Range.Text = GelenTema
        
        'Gönderenin adresi
'        If Cells(ActiveCell.Row, 18).Value = Cells(ActiveCell.Row, 17).Value & " Organization A" Then
        If Cells(ActiveCell.Row, 26).Value <> "" Then
            objDoc.Tables(1).Cell(Row:=10, Column:=3).Range.Text = Cells(ActiveCell.Row, 26).Value & "/" & Cells(ActiveCell.Row, 25).Value
        Else
            objDoc.Tables(1).Cell(Row:=10, Column:=3).Range.Text = Cells(ActiveCell.Row, 25).Value
        End If

        'Gelen Yazının Tarih ve Sayısı
        objDoc.Tables(1).Cell(Row:=11, Column:=3).Range.Text = Cells(ActiveCell.Row, 28).Value
        objDoc.Tables(1).Cell(Row:=11, Column:=5).Range.Text = Cells(ActiveCell.Row, 29).Value
        'Gönderinin Eki
        If Cells(ActiveCell.Row, 37).Value = "Package A" Then
            objDoc.CheckZarf.Value = True
        ElseIf Cells(ActiveCell.Row, 37).Value = "Package B" Then
            objDoc.CheckTorba.Value = True
        ElseIf Cells(ActiveCell.Row, 37).Value = "Package C" Then
            objDoc.CheckKoli.Value = True
        End If
        'Tabloyu doldur
        Set IlkSiraBul = Range("CM7:CM100000").Find(What:=Cells(ActiveCell.Row, 5).Value, SearchDirection:=xlNext, _
                        SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
        Set SonSiraBul = Range("CN7:CN100000").Find(What:=Cells(ActiveCell.Row, 5).Value, SearchDirection:=xlNext, _
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
            objDoc.Tables(2).Cell(Row:=i, Column:=2).Range.Text = Cells(IlkSira + i - 2, 46).Value 'Öğe türü
            objDoc.Tables(2).Cell(Row:=i, Column:=3).Range.Text = Cells(IlkSira + i - 2, 49).Value 'Öğe değeri
            objDoc.Tables(2).Cell(Row:=i, Column:=4).Range.Text = Cells(IlkSira + i - 2, 52).Value 'Adet
            If Cells(IlkSira + i - 2, 55).Value = "Dispatch List" Then
                objDoc.Tables(2).Cell(Row:=i, Column:=5).Range.Text = "Listed in the Dispatch List" 'Öğe ID No
            Else
                objDoc.Tables(2).Cell(Row:=i, Column:=5).Range.Text = Cells(IlkSira + i - 2, 55).Value 'Öğe ID No
            End If
            objDoc.Tables(2).Cell(Row:=i, Column:=6).Range.Text = Cells(IlkSira, 31).Value  'Tema No (Temai her satıra yaz.)
            objDoc.Tables(2).Cell(Row:=i, Column:=7).Range.Text = Cells(IlkSira + i - 2, 58).Value 'Açıklama
        Next i
        'Ek olarak Dispatch List
        If Cells(ActiveCell.Row, 43).Value = "Yes" Then
            If Cells(ActiveCell.Row, 44).Value > 1 Then objDoc.Tables(5).Cell(Row:=5, Column:=1).Range.Text = "Attachment: Dispatch List (" & Cells(ActiveCell.Row, 44).Value & " pages)"
            If Cells(ActiveCell.Row, 44).Value < 2 Then objDoc.Tables(5).Cell(Row:=5, Column:=1).Range.Text = "Attachment: Dispatch List (" & Cells(ActiveCell.Row, 44).Value & " page)"
        End If

        'imzalar
        objDoc.Tables(4).Cell(Row:=2, Column:=2).Range.Text = Cells(ActiveCell.Row, 112).Value 'Ad Soyad1
        objDoc.Tables(4).Cell(Row:=3, Column:=2).Range.Text = Cells(ActiveCell.Row, 113).Value 'Unvan1
        objDoc.Tables(4).Cell(Row:=2, Column:=3).Range.Text = Cells(ActiveCell.Row, 115).Value 'Ad Soyad2
        objDoc.Tables(4).Cell(Row:=3, Column:=3).Range.Text = Cells(ActiveCell.Row, 116).Value 'Unvan2
        
        'Alt bilgi ekle
        objDoc.Sections(1).Footers(wdHeaderFooterPrimary).Range.Tables(1).Cell(Row:=1, Column:=1).Range.Text = ReNameTutanak1
    
        'İmza boşluğunu sayfaya sığdırmak için düzenle
        If x > 9 And x < 14 Then
            For i = 1 To x - 9
                If i = 1 Then
                    objDoc.Tables(3).Rows(1).Delete
                ElseIf i = 2 Then
                    For j = 1 To 2
                        objDoc.Tables(3).Rows(1).Delete
                    Next j
                ElseIf i = 3 Then
                    For j = 1 To 2
                        objDoc.Tables(4).Rows(1).Delete
                    Next j
                ElseIf i = 4 Then
                    objDoc.Tables(3).Rows.Add BeforeRow:=objDoc.Tables(3).Rows(1)
                    For j = 1 To 2
                        objDoc.Tables(4).Rows(1).Delete
                    Next j
                End If
            Next i
        ElseIf x > 13 Then
            'Nothing
        End If
        
        'Sayfa sayısı kaydet komutuna bağlandı.
    '    objDoc.Close SaveChanges:=True
    '    objWord.Quit
        objDoc.Save
        'Sayfayı text dosyasından çek
        TxtFileTutanak1 = DestOpUserFolder & "Statement 1 Page Count.txt"
        #If UseADO Then
            Const adReadLine = -2&
            With CreateObject("ADODB.Stream")
                .Open
                .LoadFromFile TxtFileTutanak1
                Do Until .EOS
                    TotalSayfaTutanak1 = .ReadText(adReadLine)
                Loop
                .Close
            End With
        #Else 'UseFSO
            Const Tutanak1TristateTrue = -1&
            With CreateObject("Scripting.FileSystemObject")
                With .OpenTextFile(TxtFileTutanak1, Format:=Tutanak1TristateTrue)
                    Do Until .AtEndOfStream
                        TotalSayfaTutanak1 = .ReadLine
                    Loop
                    .Close
                End With
            End With
        #End If
        'MsgBox "Tutanak1: " & TotalSayfaTutanak1
    
        If TumDoc = True Then
            objWord.Activate
'            For i = 1 To SayPrt
'                objDoc.PrintOut
'            Next i
            objDoc.PrintOut Background:=False, Copies:=SayPrt
            'objWord.Documents.Save
            objDoc.Close SaveChanges:=False
            objWord.Visible = False
        End If
        
    End If
    
    'Dispatch List oluşturulacak.
    If Cells(ActiveCell.Row, 43).Value = "Yes" Then
        'Döküm için sayfayı belirt.(Dispatch List word daosyası bu veriyi işleyecek.)
        Set fso = CreateObject("Scripting.FileSystemObject")
        DokumSayfaGonder = Cells(ActiveCell.Row, 44).Value
        Set DokumFileTxt = fso.CreateTextFile(DestOpUserFolder & "Send Dispatch Page Count.txt", True, True)
        DokumFileTxt.Write DokumSayfaGonder
        DokumFileTxt.Close

        'Dosyayı şablondan operasyon klasörüne kopyala ve adını değiştir.
        Set fso = CreateObject("Scripting.FileSystemObject")
        fso.CopyFile (SourceDL), DestOpUserFolder & ReNameDL & ".docm", True
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
        objWord.Documents.Open FileName:=DestOpUserFolder & ReNameDL & ".docm"
        objWord.Visible = True
        objWord.Activate 'Ekrana getirir.
        'objDoc.Activate 'Ekrana getirmez.
        objWord.Application.WindowState = wdWindowStateMaximize
        Set objDoc = GetObject(DestOpUserFolder & ReNameDL & ".docm")
    '________________________________________
        
        'Birim
        Birimx = UCase(Replace(Replace(Worksheets(2).Cells(6, 99).Value, "i", "I"), "ı", "I")) & " UNIT"
        objDoc.Tables(1).Cell(Row:=1, Column:=1).Range.Text = Birimx
        'Gönderen
        objDoc.Tables(2).Cell(Row:=3, Column:=3).Range.Text = GelenTema
        'Belge tarihi
        objDoc.Tables(2).Cell(Row:=4, Column:=3).Range.Text = Cells(ActiveCell.Row, 28).Value
        'Belge sayısı
        objDoc.Tables(2).Cell(Row:=4, Column:=5).Range.Text = Cells(ActiveCell.Row, 29).Value

        'Alt bilgi ekle
        objDoc.Sections(1).Footers(wdHeaderFooterPrimary).Range.Tables(1).Cell(Row:=1, Column:=1).Range.Text = ReNameDL
    
        'Sayfa sayısı kaydet komutuna bağlandı.
        objDoc.Save
        'Sayfayı text dosyasından çek
        TxtFileDokum = DestOpUserFolder & "Dispatch Page Count.txt"
        #If UseADO Then
            Const adReadLine = -2&
            With CreateObject("ADODB.Stream")
                .Open
                .LoadFromFile TxtFileDokum
                Do Until .EOS
                    TotalSayfaDokum = .ReadText(adReadLine)
                Loop
                .Close
            End With
        #Else 'UseFSO
            Const DokumTristateTrue = -1&
            With CreateObject("Scripting.FileSystemObject")
                With .OpenTextFile(TxtFileDokum, Format:=DokumTristateTrue)
                    Do Until .AtEndOfStream
                        TotalSayfaDokum = .ReadLine
                    Loop
                    .Close
                End With
            End With
        #End If
        'MsgBox "Tutanak1: " & TotalSayfaTutanak1
        'MsgBox "Döküm: " & TotalSayfaDokum
        If TumDoc = True Then
            objWord.Activate
'            For i = 1 To 1 'SayPrt
'                objDoc.PrintOut
'            Next i
            objDoc.PrintOut Background:=False, Copies:=SayPrt
            'objWord.Documents.Save
            objDoc.Close SaveChanges:=False
            objWord.Visible = False
        End If
    End If
    
    'Tutanak1 ve Döküm sayfa sayısı
    Cells(ActiveCell.Row, 217).Value = TotalSayfaTutanak1
    Cells(ActiveCell.Row, 218).Value = TotalSayfaDokum
    
    'Nesneleri temizle
    Set fso = Nothing
    Set objWord = Nothing
    Set objDoc = Nothing

Son:

'Nesneleri temizle
Set fso = Nothing
Set objWord = Nothing
Set objDoc = Nothing

If IlceSakla <> "" Then
    Cells(ActiveCell.Row, 26).Value = IlceSakla
End If

'Worksheets(4).Activate
Application.ScreenUpdating = True
Application.DisplayAlerts = True
'Application.EnableEvents = True

'End If

End Sub

Sub Rapor2_1Rapor()

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
Dim TextLine As String

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
    Set IlkSiraBul = Range("CM7:CM100000").Find(What:=Cells(j, 5).Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    Set SonSiraBul = Range("CN7:CN100000").Find(What:=Cells(j, 5).Value, SearchDirection:=xlNext, _
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
        MsgBox "The Statement 1 data contains missing or incorrect information, so the operation cannot be started.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
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
    
    '_______________

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
                ReNameRaporNormal = Cells(i, 5).Value & "-" & Cells(6, 7).Value & " " & Cells(ActiveCell.Row, 11).Value
                SiraNoIlkSatir = i
                GoTo RaporNoDonguSonExplorer
            End If
        Next i
    Else
        ReNameRaporNormal = Cells(ActiveCell.Row, 5).Value & "-" & Cells(6, 7).Value & " " & Cells(ActiveCell.Row, 11).Value
    End If
RaporNoDonguSonExplorer:
   
    'Rapor/Alt Rapor aralıklarının tespiti.
    Set IlkSiraBul = Range("CM7:CM100000").Find(What:=Cells(SiraNoIlkSatir, 5).Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    Set SonSiraBul = Range("CN7:CN100000").Find(What:=Cells(SiraNoIlkSatir, 5).Value, SearchDirection:=xlNext, _
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
        If Cells(Explorer, 11) <> "" Then
           Cells(Explorer, 11).Select
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
                ReNameRaporNormal = Cells(i, 5).Value & "-" & Cells(6, 7).Value & " " & Cells(ActiveCell.Row, 11).Value
                SiraNoIlkSatir = i
                GoTo RaporNoDonguSon
            End If
        Next i
    Else
        ReNameRaporNormal = Cells(ActiveCell.Row, 5).Value & "-" & Cells(6, 7).Value & " " & Cells(ActiveCell.Row, 11).Value
    End If
RaporNoDonguSon:
    'Rapor/Alt Rapor aralıklarının tespiti.
    Set IlkSiraBul = Range("CM7:CM100000").Find(What:=Cells(SiraNoIlkSatir, 5).Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    Set SonSiraBul = Range("CN7:CN100000").Find(What:=Cells(SiraNoIlkSatir, 5).Value, SearchDirection:=xlNext, _
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
    RaporTipi = Cells(ActiveCell.Row, 64).Value
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
    'On Error GoTo 0
    
    'Birim
    Birimx = UCase(Replace(Replace(Worksheets(2).Cells(6, 99).Value, "i", "I"), "ı", "I")) & " UNIT"
    objDoc.Tables(1).Cell(Row:=1, Column:=1).Range.Text = Birimx
    'Rapor tarihi
    objDoc.Tables(1).Cell(Row:=4, Column:=3).Range.Text = Cells(IlkSira, 68).Value
    'Rapor No
    objDoc.Tables(1).Cell(Row:=5, Column:=3).Range.Text = Cells(AltRaporNoIlk, 67).Value
    
    If Cells(ActiveCell.Row, 62).Value <> "valid" Then
        objDoc.CheckBox1.Value = True 'Rapor Talebi
        objDoc.CheckBox2.Value = False 'Rapor3
    End If
    
    'Gönderen tema
    If Cells(IlkSira, 33).Value = "Provincial Directorate B" Then
        If Cells(IlkSira, 34).Value <> "" Then
            GelenTema = Cells(IlkSira, 25).Value & " Provincial Governorship Provincial Directorate B " & Cells(IlkSira, 34).Value
        Else
            GelenTema = Cells(IlkSira, 25).Value & " Provincial Governorship Provincial Directorate B"
        End If
    ElseIf Cells(IlkSira, 33).Value = "District Directorate B" Then
        If Cells(IlkSira, 34).Value <> "" Then
            GelenTema = Cells(IlkSira, 26).Value & " District Governorship District Directorate B " & Cells(IlkSira, 34).Value
        Else
            GelenTema = Cells(IlkSira, 26).Value & " District Governorship District Directorate B"
        End If
    ElseIf Cells(IlkSira, 33).Value = "Provincial Directorate C" Then
        If Cells(IlkSira, 34).Value <> "" Then
            GelenTema = Cells(IlkSira, 25).Value & " Provincial Governorship Provincial Directorate C " & Cells(IlkSira, 34).Value
        Else
            GelenTema = Cells(IlkSira, 25).Value & " Provincial Governorship Provincial Directorate C"
        End If
    ElseIf Cells(IlkSira, 33).Value = "District Directorate C" Then
        If Cells(IlkSira, 34).Value <> "" Then
            GelenTema = Cells(IlkSira, 26).Value & " District Governorship District Directorate C " & Cells(IlkSira, 34).Value
        Else
            GelenTema = Cells(IlkSira, 26).Value & " District Governorship District Directorate C"
        End If
    ElseIf Cells(IlkSira, 33).Value = "Provincial Directorate D" Then
        If Cells(IlkSira, 34).Value <> "" Then
            GelenTema = Cells(IlkSira, 25).Value & " Provincial Governorship Provincial Directorate D " & Cells(IlkSira, 34).Value
        Else
            GelenTema = Cells(IlkSira, 25).Value & " Provincial Governorship Provincial Directorate D"
        End If
    ElseIf Cells(IlkSira, 33).Value = "District Directorate D" Then
        If Cells(IlkSira, 34).Value <> "" Then
            GelenTema = Cells(IlkSira, 26).Value & " District Governorship District Directorate D " & Cells(IlkSira, 34).Value
        Else
            GelenTema = Cells(IlkSira, 26).Value & " District Governorship District Directorate D"
        End If
    ElseIf Cells(IlkSira, 33).Value = "Provincial Directorate E" Then
        If Cells(IlkSira, 34).Value <> "" Then
            GelenTema = Cells(IlkSira, 25).Value & " Provincial Governorship Provincial Directorate E " & Cells(IlkSira, 34).Value
        Else
            GelenTema = Cells(IlkSira, 25).Value & " Provincial Governorship Provincial Directorate E"
        End If
    ElseIf Cells(IlkSira, 33).Value = "District Directorate E" Then
        If Cells(IlkSira, 34).Value <> "" Then
            GelenTema = Cells(IlkSira, 26).Value & " District Governorship District Directorate E " & Cells(IlkSira, 34).Value
        Else
            GelenTema = Cells(IlkSira, 26).Value & " District Governorship District Directorate E"
        End If
    ElseIf InStr(Cells(IlkSira, 33).Value, "General Directorate") <> 0 Or InStr(Cells(IlkSira, 33).Value, "Regional Directorate") <> 0 Then
        If Cells(IlkSira, 34).Value <> "" Then
            GelenTema = WorksheetFunction.Proper(ThisWorkbook.Worksheets(2).Cells(6, 111).Value) & " " & Cells(IlkSira, 33).Value & " " & Cells(IlkSira, 34).Value
        Else
            GelenTema = WorksheetFunction.Proper(ThisWorkbook.Worksheets(2).Cells(6, 111).Value) & " " & Cells(IlkSira, 33).Value
        End If
    Else
        If Cells(IlkSira, 34).Value <> "" Then
            If InStr(Cells(IlkSira, 33).Value, "X.X. ") > 0 Then
                GelenTema = Mid(Cells(IlkSira, 33).Value, 6, Len(Cells(IlkSira, 33).Value)) & " " & Cells(IlkSira, 34).Value
            Else
                GelenTema = Cells(IlkSira, 33).Value & " " & Cells(IlkSira, 34).Value
            End If
        Else
            If InStr(Cells(IlkSira, 33).Value, "X.X. ") > 0 Then
                GelenTema = Mid(Cells(IlkSira, 33).Value, 6, Len(Cells(IlkSira, 33).Value))
            Else
                GelenTema = Cells(IlkSira, 33).Value
            End If
        End If
    End If
    
    'İlgi
    'RaporIlgi = GelenTema & "n" & Right(GelenTema, 1) & "n " & Cells(IlkSira, 28).Value & " tarihli ve " & Cells(IlkSira, 29).Value & " sayılı yazısı."
    RaporIlgi = "The letter from the " & GelenTema & ", dated " & Cells(IlkSira, 28).Value & ", reference number " & Cells(IlkSira, 29).Value & "."
    If Cells(ActiveCell.Row, 62).Value <> "valid" Then
        objDoc.Tables(1).Cell(Row:=8, Column:=3).Range.Text = RaporIlgi
    Else
        objDoc.Tables(1).Cell(Row:=6, Column:=3).Range.Text = RaporIlgi
    End If
    'İlgili birim
    If Cells(ActiveCell.Row, 62).Value <> "valid" Then
        If Cells(IlkSira, 69).Value <> "" Then
            If InStr(Cells(IlkSira, 69).Value, "X.X. ") > 0 Then
                objDoc.Tables(1).Cell(Row:=9, Column:=3).Range.Text = Mid(Cells(IlkSira, 69).Value, 6, Len(Cells(IlkSira, 69).Value))
            Else
                objDoc.Tables(1).Cell(Row:=9, Column:=3).Range.Text = Cells(IlkSira, 69).Value
            End If
        Else
            'objDoc.Tables(1).Cell(Row:=9, Column:=3).Range.Text = Cells(IlkSira, 69).Value
            objDoc.Tables(1).Rows(9).Delete
        End If
    Else
        If Cells(IlkSira, 69).Value <> "" Then
            If InStr(Cells(IlkSira, 69).Value, "X.X. ") > 0 Then
                objDoc.Tables(1).Cell(Row:=7, Column:=3).Range.Text = Mid(Cells(IlkSira, 69).Value, 6, Len(Cells(IlkSira, 69).Value))
            Else
                objDoc.Tables(1).Cell(Row:=7, Column:=3).Range.Text = Cells(IlkSira, 69).Value
            End If
        Else
            'objDoc.Tables(1).Cell(Row:=7, Column:=3).Range.Text = Cells(IlkSira, 69).Value
            objDoc.Tables(1).Rows(7).Delete
        End If
    End If
    
    'Tekil-çoğul tipA
    AdetTopla = Application.Sum(Range(Cells(AltRaporNoIlk, 52), Cells(AltRaporNoSon, 52)))
    If AdetTopla = 1 Then
        TekCogulTipA = "Type A"
    ElseIf AdetTopla > 1 Then
        TekCogulTipA = "Type A items"
    End If
    'objDoc.Tables(2).Cell(Row:=1, Column:=1).Range.Text = "İnceleme Konusu " & WorksheetFunction.Proper(TekCogulTipA)
    objDoc.Tables(2).Cell(Row:=1, Column:=1).Range.Text = "The examined " & TekCogulTipA
    
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
            objDoc.Tables(2).Cell(Row:=j + 0, Column:=3).Range.Text = Cells(i, 46).Value 'Öğe türü
            objDoc.Tables(2).Cell(Row:=j + 1, Column:=3).Range.Text = Cells(i, 49).Value 'Öğe değeri
            objDoc.Tables(2).Cell(Row:=j + 2, Column:=3).Range.Text = Cells(i, 52).Value 'Adet
            If Cells(i, 55).Value = "Dispatch List" Then
                objDoc.Tables(2).Cell(Row:=j + 3, Column:=3).Range.Text = "Listed in the Attachment to Statement 1" 'Öğe ID
            Else
                objDoc.Tables(2).Cell(Row:=j + 3, Column:=3).Range.Text = Cells(i, 55).Value 'Öğe ID
            End If
            j = j + 0
        ElseIf x Mod 2 = 0 Then
            'Verileri aktar
            objDoc.Tables(2).Cell(Row:=j + 0, Column:=4).Range.Text = Cells(i, 46).Value 'Öğe türü
            objDoc.Tables(2).Cell(Row:=j + 1, Column:=4).Range.Text = Cells(i, 49).Value 'Öğe değeri
            objDoc.Tables(2).Cell(Row:=j + 2, Column:=4).Range.Text = Cells(i, 52).Value 'Adet
            If Cells(i, 55).Value = "Dispatch List" Then
                objDoc.Tables(2).Cell(Row:=j + 3, Column:=4).Range.Text = "Listed in the Attachment to Statement 1" 'Öğe ID
            Else
                objDoc.Tables(2).Cell(Row:=j + 3, Column:=4).Range.Text = Cells(i, 55).Value 'Öğe ID
            End If
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
    objDoc.Tables(3).Cell(Row:=1, Column:=3).Range.Text = Cells(IlkSira, 31).Value  'Tema no

    'Rapor metin kısmı
    'Tanımlamalar
    AdetTopla = Application.Sum(Range(Cells(AltRaporNoIlk, 52), Cells(AltRaporNoSon, 52)))
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
    objDoc.Tables(6).Cell(Row:=3, Column:=1).Range.Text = Cells(ActiveCell.Row, 136).Value 'Ad Soyad Directorate
    objDoc.Tables(6).Cell(Row:=4, Column:=1).Range.Text = Cells(ActiveCell.Row, 137).Value 'Unvan
    objDoc.Tables(6).Cell(Row:=3, Column:=2).Range.Text = Cells(ActiveCell.Row, 118).Value 'Ad Soyad1
    objDoc.Tables(6).Cell(Row:=4, Column:=2).Range.Text = Cells(ActiveCell.Row, 119).Value 'Unvan1
    objDoc.Tables(6).Cell(Row:=3, Column:=3).Range.Text = Cells(ActiveCell.Row, 121).Value 'Ad Soyad2
    objDoc.Tables(6).Cell(Row:=4, Column:=3).Range.Text = Cells(ActiveCell.Row, 122).Value 'Unvan2

    'Alt bilgi ekle
    objDoc.Sections(1).Footers(wdHeaderFooterPrimary).Range.Tables(1).Cell(Row:=1, Column:=1).Range.Text = ReNameRaporNormal
    
    'Not ekle
    If Cells(ActiveCell.Row, 66).Value = "Yes" Then
        TxtFileNot = DestNotlar & Cells(ActiveCell.Row, 46).Value & ".txt"
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

    'Rapor sayfa sayısını oluşturmadan önce eskisini sil
    If Cells(i, 174).Value = "No" Then
        For i = IlkSira To SonSira
            If Cells(i, 11).Value = "" Then
                Cells(i, 99).Value = ""
            End If
        Next i
    End If
    
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
    Cells(ActiveCell.Row, 99).Value = TotalSayfaRapor

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
    Cells(b, 10).Select
End If

Son:

'Worksheets(4).Activate
Application.ScreenUpdating = True
Application.DisplayAlerts = True
'Application.EnableEvents = True

'End If

End Sub

Sub Rapor2_2Rapor()

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
Dim TextLine As String

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
    Set IlkSiraBul = Range("CM7:CM100000").Find(What:=Cells(j, 5).Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    Set SonSiraBul = Range("CN7:CN100000").Find(What:=Cells(j, 5).Value, SearchDirection:=xlNext, _
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
    If Cells(IlkSira, 13).Value = "x" Then
        MsgBox "The Statement 1 data contains missing or incorrect information, so the operation cannot be started.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
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
                ReNameRaporNormal = Cells(i, 5).Value & "-" & Cells(6, 14).Value & " " & Cells(ActiveCell.Row, 11).Value
                SiraNoIlkSatir = i
                GoTo RaporNoDonguSonExplorer
            End If
        Next i
    Else
        ReNameRaporNormal = Cells(ActiveCell.Row, 5).Value & "-" & Cells(6, 14).Value & " " & Cells(ActiveCell.Row, 11).Value
    End If
RaporNoDonguSonExplorer:
   
    'Rapor/Alt Rapor aralıklarının tespiti.
    Set IlkSiraBul = Range("CM7:CM100000").Find(What:=Cells(SiraNoIlkSatir, 5).Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    Set SonSiraBul = Range("CN7:CN100000").Find(What:=Cells(SiraNoIlkSatir, 5).Value, SearchDirection:=xlNext, _
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
        If Cells(Explorer, 11) <> "" Then
           Cells(Explorer, 11).Select
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
                ReNameRaporNormal = Cells(i, 5).Value & "-" & Cells(6, 14).Value & " " & Cells(ActiveCell.Row, 11).Value
                SiraNoIlkSatir = i
                GoTo RaporNoDonguSon
            End If
        Next i
    Else
        ReNameRaporNormal = Cells(ActiveCell.Row, 5).Value & "-" & Cells(6, 14).Value & " " & Cells(ActiveCell.Row, 11).Value
    End If
RaporNoDonguSon:
    'Rapor/Alt Rapor aralıklarının tespiti.
    Set IlkSiraBul = Range("CM7:CM100000").Find(What:=Cells(SiraNoIlkSatir, 5).Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    Set SonSiraBul = Range("CN7:CN100000").Find(What:=Cells(SiraNoIlkSatir, 5).Value, SearchDirection:=xlNext, _
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
        If Cells(i + 1, 14).Value <> "" And i < SonSira Then
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
    RaporTipi = Cells(ActiveCell.Row, 64).Value
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
    objDoc.Tables(1).Cell(Row:=4, Column:=3).Range.Text = Cells(IlkSira, 68).Value
    'Rapor No
    objDoc.Tables(1).Cell(Row:=5, Column:=3).Range.Text = Cells(AltRaporNoIlk, 67).Value

    objDoc.CheckBox1.Value = True 'Rapor Talebi
    objDoc.CheckBox2.Value = False 'Rapor3
    
    'Gönderen tema
    If Cells(IlkSira, 33).Value = "Provincial Directorate B" Then
        If Cells(IlkSira, 34).Value <> "" Then
            GelenTema = Cells(IlkSira, 25).Value & " Provincial Governorship Provincial Directorate B " & Cells(IlkSira, 34).Value
        Else
            GelenTema = Cells(IlkSira, 25).Value & " Provincial Governorship Provincial Directorate B"
        End If
    ElseIf Cells(IlkSira, 33).Value = "District Directorate B" Then
        If Cells(IlkSira, 34).Value <> "" Then
            GelenTema = Cells(IlkSira, 26).Value & " District Governorship District Directorate B " & Cells(IlkSira, 34).Value
        Else
            GelenTema = Cells(IlkSira, 26).Value & " District Governorship District Directorate B"
        End If
    ElseIf Cells(IlkSira, 33).Value = "Provincial Directorate C" Then
        If Cells(IlkSira, 34).Value <> "" Then
            GelenTema = Cells(IlkSira, 25).Value & " Provincial Governorship Provincial Directorate C " & Cells(IlkSira, 34).Value
        Else
            GelenTema = Cells(IlkSira, 25).Value & " Provincial Governorship Provincial Directorate C"
        End If
    ElseIf Cells(IlkSira, 33).Value = "District Directorate C" Then
        If Cells(IlkSira, 34).Value <> "" Then
            GelenTema = Cells(IlkSira, 26).Value & " District Governorship District Directorate C " & Cells(IlkSira, 34).Value
        Else
            GelenTema = Cells(IlkSira, 26).Value & " District Governorship District Directorate C"
        End If
    ElseIf Cells(IlkSira, 33).Value = "Provincial Directorate D" Then
        If Cells(IlkSira, 34).Value <> "" Then
            GelenTema = Cells(IlkSira, 25).Value & " Provincial Governorship Provincial Directorate D " & Cells(IlkSira, 34).Value
        Else
            GelenTema = Cells(IlkSira, 25).Value & " Provincial Governorship Provincial Directorate D"
        End If
    ElseIf Cells(IlkSira, 33).Value = "District Directorate D" Then
        If Cells(IlkSira, 34).Value <> "" Then
            GelenTema = Cells(IlkSira, 26).Value & " District Governorship District Directorate D " & Cells(IlkSira, 34).Value
        Else
            GelenTema = Cells(IlkSira, 26).Value & " District Governorship District Directorate D"
        End If
    ElseIf Cells(IlkSira, 33).Value = "Provincial Directorate E" Then
        If Cells(IlkSira, 34).Value <> "" Then
            GelenTema = Cells(IlkSira, 25).Value & " Provincial Governorship Provincial Directorate E " & Cells(IlkSira, 34).Value
        Else
            GelenTema = Cells(IlkSira, 25).Value & " Provincial Governorship Provincial Directorate E"
        End If
    ElseIf Cells(IlkSira, 33).Value = "District Directorate E" Then
        If Cells(IlkSira, 34).Value <> "" Then
            GelenTema = Cells(IlkSira, 26).Value & " District Governorship District Directorate E " & Cells(IlkSira, 34).Value
        Else
            GelenTema = Cells(IlkSira, 26).Value & " District Governorship District Directorate E"
        End If
    ElseIf InStr(Cells(IlkSira, 33).Value, "General Directorate") <> 0 Or InStr(Cells(IlkSira, 33).Value, "Regional Directorate") <> 0 Then
        If Cells(IlkSira, 34).Value <> "" Then
            GelenTema = WorksheetFunction.Proper(ThisWorkbook.Worksheets(2).Cells(6, 111).Value) & " " & Cells(IlkSira, 33).Value & " " & Cells(IlkSira, 34).Value
        Else
            GelenTema = WorksheetFunction.Proper(ThisWorkbook.Worksheets(2).Cells(6, 111).Value) & " " & Cells(IlkSira, 33).Value
        End If
    Else
        If Cells(IlkSira, 34).Value <> "" Then
            If InStr(Cells(IlkSira, 33).Value, "X.X. ") > 0 Then
                GelenTema = Mid(Cells(IlkSira, 33).Value, 6, Len(Cells(IlkSira, 33).Value)) & " " & Cells(IlkSira, 34).Value
            Else
                GelenTema = Cells(IlkSira, 33).Value & " " & Cells(IlkSira, 34).Value
            End If
        Else
            If InStr(Cells(IlkSira, 33).Value, "X.X. ") > 0 Then
                GelenTema = Mid(Cells(IlkSira, 33).Value, 6, Len(Cells(IlkSira, 33).Value))
            Else
                GelenTema = Cells(IlkSira, 33).Value
            End If
        End If
    End If
    
    'İlgi
    'RaporIlgi = GelenTema & "n" & Right(GelenTema, 1) & "n " & Cells(IlkSira, 28).Value & " tarihli ve " & Cells(IlkSira, 29).Value & " sayılı yazısı."
    RaporIlgi = "The letter from the " & GelenTema & ", dated " & Cells(IlkSira, 28).Value & ", reference number " & Cells(IlkSira, 29).Value & "."
    If Cells(ActiveCell.Row, 62).Value <> "valid" Then
        objDoc.Tables(1).Cell(Row:=8, Column:=3).Range.Text = RaporIlgi
    Else
        objDoc.Tables(1).Cell(Row:=6, Column:=3).Range.Text = RaporIlgi
    End If
    'İlgili birim
    If Cells(ActiveCell.Row, 62).Value <> "valid" Then
        If Cells(IlkSira, 69).Value <> "" Then
            If InStr(Cells(IlkSira, 69).Value, "X.X. ") > 0 Then
                objDoc.Tables(1).Cell(Row:=9, Column:=3).Range.Text = Mid(Cells(IlkSira, 69).Value, 6, Len(Cells(IlkSira, 69).Value))
            Else
                objDoc.Tables(1).Cell(Row:=9, Column:=3).Range.Text = Cells(IlkSira, 69).Value
            End If
        Else
            'objDoc.Tables(1).Cell(Row:=9, Column:=3).Range.Text = Cells(IlkSira, 69).Value
            objDoc.Tables(1).Rows(9).Delete
        End If
    Else
        If Cells(IlkSira, 69).Value <> "" Then
            If InStr(Cells(IlkSira, 69).Value, "X.X. ") > 0 Then
                objDoc.Tables(1).Cell(Row:=7, Column:=3).Range.Text = Mid(Cells(IlkSira, 69).Value, 6, Len(Cells(IlkSira, 69).Value))
            Else
                objDoc.Tables(1).Cell(Row:=7, Column:=3).Range.Text = Cells(IlkSira, 69).Value
            End If
        Else
            'objDoc.Tables(1).Cell(Row:=7, Column:=3).Range.Text = Cells(IlkSira, 69).Value
            objDoc.Tables(1).Rows(7).Delete
        End If

    End If

    'Tekil-çoğul tipA
    AdetTopla = Application.Sum(Range(Cells(AltRaporNoIlk, 52), Cells(AltRaporNoSon, 52)))
    If AdetTopla = 1 Then
        TekCogulTipA = "Type A"
    ElseIf AdetTopla > 1 Then
        TekCogulTipA = "Type A items"
    End If
    'objDoc.Tables(2).Cell(Row:=1, Column:=1).Range.Text = "İnceleme Konusu " & WorksheetFunction.Proper(TekCogulTipA)
    objDoc.Tables(2).Cell(Row:=1, Column:=1).Range.Text = "The examined " & TekCogulTipA
    
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
            objDoc.Tables(2).Cell(Row:=j + 0, Column:=3).Range.Text = Cells(i, 46).Value 'Öğe türü
            objDoc.Tables(2).Cell(Row:=j + 1, Column:=3).Range.Text = Cells(i, 49).Value 'Öğe değeri
            objDoc.Tables(2).Cell(Row:=j + 2, Column:=3).Range.Text = Cells(i, 52).Value 'Adet
            If Cells(i, 55).Value = "Dispatch List" Then
                objDoc.Tables(2).Cell(Row:=j + 3, Column:=3).Range.Text = "Listed in the Attachment to Statement 1" 'Öğe ID
            Else
                objDoc.Tables(2).Cell(Row:=j + 3, Column:=3).Range.Text = Cells(i, 55).Value 'Öğe ID
            End If
            j = j + 0
        ElseIf x Mod 2 = 0 Then
            'Verileri aktar
            objDoc.Tables(2).Cell(Row:=j + 0, Column:=4).Range.Text = Cells(i, 46).Value 'Öğe türü
            objDoc.Tables(2).Cell(Row:=j + 1, Column:=4).Range.Text = Cells(i, 49).Value 'Öğe değeri
            objDoc.Tables(2).Cell(Row:=j + 2, Column:=4).Range.Text = Cells(i, 52).Value 'Adet
            If Cells(i, 55).Value = "Dispatch List" Then
                objDoc.Tables(2).Cell(Row:=j + 3, Column:=4).Range.Text = "Listed in the Attachment to Statement 1" 'Öğe ID
            Else
                objDoc.Tables(2).Cell(Row:=j + 3, Column:=4).Range.Text = Cells(i, 55).Value 'Öğe ID
            End If
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
    objDoc.Tables(3).Cell(Row:=1, Column:=3).Range.Text = Cells(IlkSira, 31).Value  'Tema no


    'Rapor metin kısmı
    'Tanımlamalar
    AdetTopla = Application.Sum(Range(Cells(AltRaporNoIlk, 52), Cells(AltRaporNoSon, 52)))
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
    objDoc.Tables(6).Cell(Row:=3, Column:=1).Range.Text = Cells(ActiveCell.Row, 136).Value 'Ad Soyad Directorate
    objDoc.Tables(6).Cell(Row:=4, Column:=1).Range.Text = Cells(ActiveCell.Row, 137).Value 'Unvan
    objDoc.Tables(6).Cell(Row:=3, Column:=2).Range.Text = Cells(ActiveCell.Row, 118).Value 'Ad Soyad1
    objDoc.Tables(6).Cell(Row:=4, Column:=2).Range.Text = Cells(ActiveCell.Row, 119).Value 'Unvan1
    objDoc.Tables(6).Cell(Row:=3, Column:=3).Range.Text = Cells(ActiveCell.Row, 121).Value 'Ad Soyad2
    objDoc.Tables(6).Cell(Row:=4, Column:=3).Range.Text = Cells(ActiveCell.Row, 122).Value 'Unvan2

    'Alt bilgi ekle
    objDoc.Sections(1).Footers(wdHeaderFooterPrimary).Range.Tables(1).Cell(Row:=1, Column:=1).Range.Text = ReNameRaporNormal
    
    'Not ekle
    If Cells(ActiveCell.Row, 66).Value = "Yes" Then
        TxtFileNot = DestNotlar & Cells(ActiveCell.Row, 46).Value & ".txt"
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

    'Rapor sayfa sayısını oluşturmadan önce eskisini sil
    If Cells(i, 174).Value = "Yes" Then
        For i = IlkSira To SonSira
            If Cells(i, 11).Value = "" Then
                Cells(i, 219).Value = ""
            End If
        Next i
    End If
    
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
    Cells(ActiveCell.Row, 219).Value = TotalSayfaRapor

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
    Cells(b, 18).Select
End If

Son:

'Worksheets(4).Activate
Application.ScreenUpdating = True
Application.DisplayAlerts = True
'Application.EnableEvents = True

'End If

End Sub

Sub Rapor2_1Tutanak2()

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
Dim Birimx As String, UserName As String

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
    Set IlkSiraBul = Range("CM7:CM100000").Find(What:=Cells(j, 5).Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    Set SonSiraBul = Range("CN7:CN100000").Find(What:=Cells(j, 5).Value, SearchDirection:=xlNext, _
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
        MsgBox "The Statement 1 data contains missing or incorrect information, so the operation cannot be started.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
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
    ReNameTutanak2Normal = Cells(ActiveCell.Row, 5).Value & "-" & Cells(6, 8).Value
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
    objDoc.Tables(1).Cell(Row:=4, Column:=3).Range.Text = Cells(ActiveCell.Row, 71).Value
    'Belge tarihi ve numarası
    objDoc.Tables(1).Cell(Row:=6, Column:=3).Range.Text = Cells(ActiveCell.Row, 28).Value
    objDoc.Tables(1).Cell(Row:=6, Column:=5).Range.Text = Cells(ActiveCell.Row, 29).Value
    
    'Gönderilen
    If Cells(ActiveCell.Row, 72).Value = "Provincial Directorate B" Then
        GidenTema = Cells(ActiveCell.Row, 77).Value & " Provincial Governorship Provincial Directorate B " & Cells(ActiveCell.Row, 73).Value
    ElseIf Cells(ActiveCell.Row, 72).Value = "District Directorate B" Then
        GidenTema = Cells(ActiveCell.Row, 78).Value & " District Governorship District Directorate B " & Cells(ActiveCell.Row, 73).Value
    ElseIf Cells(ActiveCell.Row, 72).Value = "Provincial Directorate C" Then
        GidenTema = Cells(ActiveCell.Row, 77).Value & " Provincial Governorship Provincial Directorate C " & Cells(ActiveCell.Row, 73).Value
    ElseIf Cells(ActiveCell.Row, 72).Value = "District Directorate C" Then
        GidenTema = Cells(ActiveCell.Row, 78).Value & " District Governorship District Directorate C " & Cells(ActiveCell.Row, 73).Value
    ElseIf Cells(ActiveCell.Row, 72).Value = "Provincial Directorate D" Then
        GidenTema = Cells(ActiveCell.Row, 77).Value & " Provincial Governorship Provincial Directorate D " & Cells(ActiveCell.Row, 73).Value
    ElseIf Cells(ActiveCell.Row, 72).Value = "District Directorate D" Then
        GidenTema = Cells(ActiveCell.Row, 78).Value & " District Governorship District Directorate D " & Cells(ActiveCell.Row, 73).Value
    ElseIf Cells(ActiveCell.Row, 72).Value = "Provincial Directorate E" Then
        GidenTema = Cells(ActiveCell.Row, 77).Value & " Provincial Governorship Provincial Directorate E " & Cells(ActiveCell.Row, 73).Value
    ElseIf Cells(ActiveCell.Row, 72).Value = "District Directorate E" Then
        GidenTema = Cells(ActiveCell.Row, 78).Value & " District Governorship District Directorate E " & Cells(ActiveCell.Row, 73).Value
    ElseIf InStr(Cells(ActiveCell.Row, 72).Value, "General Directorate") <> 0 Or InStr(Cells(ActiveCell.Row, 72).Value, "Regional Directorate") <> 0 Then
        If Cells(ActiveCell.Row, 73).Value <> "" Then
            GidenTema = WorksheetFunction.Proper(ThisWorkbook.Worksheets(2).Cells(6, 111).Value) & " " & Cells(ActiveCell.Row, 72).Value & " " & Cells(ActiveCell.Row, 73).Value
        Else
            GidenTema = WorksheetFunction.Proper(ThisWorkbook.Worksheets(2).Cells(6, 111).Value) & " " & Cells(ActiveCell.Row, 72).Value
        End If
    Else
        If Cells(ActiveCell.Row, 73).Value <> "" Then
            If InStr(Cells(ActiveCell.Row, 72).Value, "X.X. ") > 0 Then
                GidenTema = Mid(Cells(ActiveCell.Row, 72).Value, 6, Len(Cells(ActiveCell.Row, 72).Value)) & " " & Cells(ActiveCell.Row, 73).Value
            Else
                GidenTema = Cells(ActiveCell.Row, 72).Value & " " & Cells(ActiveCell.Row, 73).Value
            End If
        Else
            If InStr(Cells(ActiveCell.Row, 72).Value, "X.X. ") > 0 Then
                GidenTema = Mid(Cells(ActiveCell.Row, 72).Value, 6, Len(Cells(ActiveCell.Row, 72).Value))
            Else
                GidenTema = Cells(ActiveCell.Row, 72).Value
            End If
        End If
    End If
    
    
    objDoc.Tables(1).Cell(Row:=5, Column:=3).Range.Text = GidenTema
    
    'Tabloyu doldur
    Set IlkSiraBul = Range("CM7:CM100000").Find(What:=Cells(ActiveCell.Row, 5).Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    Set SonSiraBul = Range("CN7:CN100000").Find(What:=Cells(ActiveCell.Row, 5).Value, SearchDirection:=xlNext, _
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
        objDoc.Tables(2).Cell(Row:=i, Column:=2).Range.Text = Cells(IlkSira + i - 2, 46).Value 'Öğe türü
        objDoc.Tables(2).Cell(Row:=i, Column:=3).Range.Text = Cells(IlkSira + i - 2, 49).Value 'Öğe değeri
        objDoc.Tables(2).Cell(Row:=i, Column:=4).Range.Text = Cells(IlkSira + i - 2, 52).Value 'Adet
        If Cells(IlkSira + i - 2, 55).Value = "Dispatch List" Then
            objDoc.Tables(2).Cell(Row:=i, Column:=5).Range.Text = "Listed in the Attachment to Statement 1" '"Listed in the Dispatch List" 'Öğe ID No
        Else
            objDoc.Tables(2).Cell(Row:=i, Column:=5).Range.Text = Cells(IlkSira + i - 2, 55).Value 'Öğe ID No
        End If
        objDoc.Tables(2).Cell(Row:=i, Column:=6).Range.Text = Cells(IlkSira, 31).Value  'Tema No (Temai her satıra yaz.)
        objDoc.Tables(2).Cell(Row:=i, Column:=7).Range.Text = Cells(IlkSira + i - 2, 58).Value 'Açıklama
    Next i


    'Tutanak2 tutanağının metin kısmı
    'Tanımlamalar
    If Application.Sum(Range(Cells(IlkSira, 52), Cells(SonSira, 52))) > 1 Then
        Ek4 = "items"
    Else
        Ek4 = "item"
    End If
    Bolum1 = vbTab & "A total of "
    Bolum2 = " " & Ek4 & " , as described above, have been placed inside "
    Bolum3 = " "
    Bolum4 = " and enclosed to be sent to the designated unit mentioned above."

    Ek1 = Application.Sum(Range(Cells(IlkSira, 52), Cells(SonSira, 52)))
    Ek2 = Cells(ActiveCell.Row, 76).Value
    Ek3 = Application.WorksheetFunction.Proper(Cells(ActiveCell.Row, 75).Value)
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

    'imzalar
    objDoc.Tables(4).Cell(Row:=6, Column:=2).Range.Text = Cells(ActiveCell.Row, 124).Value 'Ad Soyad1
    objDoc.Tables(4).Cell(Row:=7, Column:=2).Range.Text = Cells(ActiveCell.Row, 125).Value 'Unvan1
    objDoc.Tables(4).Cell(Row:=6, Column:=3).Range.Text = Cells(ActiveCell.Row, 127).Value 'Ad Soyad2
    objDoc.Tables(4).Cell(Row:=7, Column:=3).Range.Text = Cells(ActiveCell.Row, 128).Value 'Unvan2
    
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
    Cells(ActiveCell.Row, 100).Value = TotalSayfaTutanak2

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

'Worksheets(4).Activate
Application.ScreenUpdating = True
Application.DisplayAlerts = True
'Application.EnableEvents = True

'End If

End Sub

Sub Rapor2_2Tutanak2XXXMudGiden()

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
Dim Birimx As String, UserName As String

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
    Set IlkSiraBul = Range("CM7:CM100000").Find(What:=Cells(j, 5).Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    Set SonSiraBul = Range("CN7:CN100000").Find(What:=Cells(j, 5).Value, SearchDirection:=xlNext, _
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
    If Cells(IlkSira, 13).Value = "x" Then
        MsgBox "The Statement 1 data contains missing or incorrect information, so the operation cannot be started.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    
    For i = IlkSira To SonSira
        If Cells(i, 14).Value = "x" Then
            MsgBox "The report data contains missing or incorrect information, so the operation cannot be started.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
            GoTo Son
        End If
    Next i

    
    'Pathfinder...
    AutoPath = ThisWorkbook.Path
    DestOperasyon = AutoPath & "\System Files\Operation\"

    'TUTANAK2 TANIMLARI
    SourceTutanak2Normal = AutoPath & "\System Files\System Templates\Statement 2 Template\Statement 2.docm"

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
    ReNameTutanak2Normal = Cells(ActiveCell.Row, 5).Value & "-" & Cells(6, 15).Value
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
    objDoc.Tables(1).Cell(Row:=4, Column:=3).Range.Text = Cells(ActiveCell.Row, 183).Value
    'Belge tarihi ve numarası
    objDoc.Tables(1).Cell(Row:=6, Column:=3).Range.Text = Cells(ActiveCell.Row, 28).Value
    objDoc.Tables(1).Cell(Row:=6, Column:=5).Range.Text = Cells(ActiveCell.Row, 29).Value
    
    'Gönderilen
    GidenTema = "ORGANIZATION A XXX Directorate"
    objDoc.Tables(1).Cell(Row:=5, Column:=3).Range.Text = GidenTema
    
    'Tabloyu doldur
    Set IlkSiraBul = Range("CM7:CM100000").Find(What:=Cells(ActiveCell.Row, 5).Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    Set SonSiraBul = Range("CN7:CN100000").Find(What:=Cells(ActiveCell.Row, 5).Value, SearchDirection:=xlNext, _
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
    
    If Cells(IlkSira, 187).Value = "All" Then 'TÜMÜ gönderiliyor
        'Tabloya satır ekle
        If SonSira - IlkSira > 0 Then
            With objDoc.Tables(2)
                For i = IlkSira To SonSira - 1
                    .Rows.Add BeforeRow:=objDoc.Tables(2).Rows(2)
                Next i
            End With
        End If
        'Tabloyu doldur.
        i = 1
        Ek1 = 0
        For j = IlkSira To SonSira
            i = i + 1
            objDoc.Tables(2).Cell(Row:=i, Column:=1).Range.Text = i - 1 'Tablo sıra no
            objDoc.Tables(2).Cell(Row:=i, Column:=2).Range.Text = Cells(j, 46).Value 'Öğe türü
            objDoc.Tables(2).Cell(Row:=i, Column:=3).Range.Text = Cells(j, 49).Value 'Öğe değeri
            objDoc.Tables(2).Cell(Row:=i, Column:=4).Range.Text = Cells(j, 52).Value 'Adet
            Ek1 = Ek1 + Cells(j, 52).Value
            If Cells(j, 55).Value = "Dispatch List" Then
                objDoc.Tables(2).Cell(Row:=i, Column:=5).Range.Text = "Listed in the Attachment to Statement 1" '"Listed in the Dispatch List" 'Öğe ID No
            Else
                objDoc.Tables(2).Cell(Row:=i, Column:=5).Range.Text = Cells(j, 55).Value 'Öğe ID No
            End If
            objDoc.Tables(2).Cell(Row:=i, Column:=6).Range.Text = Cells(IlkSira, 31).Value  'Tema No (Temai her satıra yaz.)
            objDoc.Tables(2).Cell(Row:=i, Column:=7).Range.Text = Cells(j, 58).Value 'Açıklama
        Next j
    ElseIf Cells(IlkSira, 187).Value = "Technique A" Then 'SADECE Technique A gönderiliyor
        'Tabloya satır ekle
        y = 0
        If SonSira - IlkSira > 0 Then
            For i = IlkSira To SonSira
                If Cells(i, 62).Value = "invalid" And Left(Cells(i, 63).Value, 11) = "Technique A" Then
                    y = y + 1
                End If
            Next i
            If y > 1 Then
                With objDoc.Tables(2)
                    For i = IlkSira To IlkSira + y - 2
                        .Rows.Add BeforeRow:=objDoc.Tables(2).Rows(2)
                    Next i
                End With
            End If
        End If
        'Tabloyu doldur.
        i = 1
        Ek1 = 0
        For j = IlkSira To SonSira
            If Cells(j, 62).Value = "invalid" And Left(Cells(j, 63).Value, 11) = "Technique A" Then
                i = i + 1
                objDoc.Tables(2).Cell(Row:=i, Column:=1).Range.Text = i - 1 'Tablo sıra no
                objDoc.Tables(2).Cell(Row:=i, Column:=2).Range.Text = Cells(j, 46).Value 'Öğe türü
                objDoc.Tables(2).Cell(Row:=i, Column:=3).Range.Text = Cells(j, 49).Value 'Öğe değeri
                objDoc.Tables(2).Cell(Row:=i, Column:=4).Range.Text = Cells(j, 52).Value 'Adet
                Ek1 = Ek1 + Cells(j, 52).Value
                If Cells(j, 55).Value = "Dispatch List" Then
                    objDoc.Tables(2).Cell(Row:=i, Column:=5).Range.Text = "Listed in the Attachment to Statement 1" '"Listed in the Dispatch List" 'Öğe ID No
                Else
                    objDoc.Tables(2).Cell(Row:=i, Column:=5).Range.Text = Cells(j, 55).Value 'Öğe ID No
                End If
                objDoc.Tables(2).Cell(Row:=i, Column:=6).Range.Text = Cells(IlkSira, 31).Value  'Tema No (Temai her satıra yaz.)
                objDoc.Tables(2).Cell(Row:=i, Column:=7).Range.Text = Cells(j, 58).Value 'Açıklama
            End If
        Next j
    End If

    'Tutanak2 tutanağının metin kısmı
    'Tanımlamalar
    If CInt(Ek1) > 1 Then
        Ek4 = "items"
    Else
        Ek4 = "item"
    End If
    Bolum1 = vbTab & "A total of "
    Bolum2 = " " & Ek4 & " , as described above, have been placed inside "
    Bolum3 = " "
    Bolum4 = " and enclosed to be sent to the designated unit mentioned above."

    Ek2 = Cells(ActiveCell.Row, 185).Value
    Ek3 = Application.WorksheetFunction.Proper(Cells(ActiveCell.Row, 184).Value)
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
    
    Set MyRange = objDoc.Tables(3).Cell(Row:=1, Column:=1).Range
    MyRange.Find.Execute FindText:=Ek3
    MyRange.Font.Bold = True

    'imzalar
    objDoc.Tables(4).Cell(Row:=6, Column:=2).Range.Text = Cells(ActiveCell.Row, 148).Value 'Ad Soyad1
    objDoc.Tables(4).Cell(Row:=7, Column:=2).Range.Text = Cells(ActiveCell.Row, 149).Value 'Unvan1
    objDoc.Tables(4).Cell(Row:=6, Column:=3).Range.Text = Cells(ActiveCell.Row, 151).Value 'Ad Soyad2
    objDoc.Tables(4).Cell(Row:=7, Column:=3).Range.Text = Cells(ActiveCell.Row, 152).Value 'Unvan2
    
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
    Cells(ActiveCell.Row, 220).Value = TotalSayfaTutanak2

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

'Worksheets(4).Activate
Application.ScreenUpdating = True
Application.DisplayAlerts = True
'Application.EnableEvents = True

'End If

End Sub

Sub Rapor2_2XXXMudUstYazi()

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
Dim Birimx As String, UserName As String
Dim gecersizSay As Long
Dim IlBuyukHarf, IlKucukHarf, IlceBuyukHarf, IlceKucukHarf As String

Dim AdetTakip As Integer, TipATakip As String, NomTakip As Integer
Dim GonderimUsulu As String

Dim BStr As Integer
Dim M2 As Boolean, M3 As Boolean, M4 As Boolean
Dim Ifv As Boolean, Ify As Boolean
Dim XXXMudNotu As Boolean
Dim IlgiStrSay As Integer, Govde1StrSay As Integer, IlgiRng As Object, Govde1Rng As Object, Govde2Rng As Object
Dim IlgiFarkSay As Integer, Govde1FarkSay As Integer
Dim ustbilgimuhatap As String, ustbilgitarih As String, ustbilgisayi As String
Dim CokluSayfa As Integer

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
    Set IlkSiraBul = Range("CM7:CM100000").Find(What:=Cells(j, 5).Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    Set SonSiraBul = Range("CN7:CN100000").Find(What:=Cells(j, 5).Value, SearchDirection:=xlNext, _
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
    
    If Cells(IlkSira, 13).Value = "x" Then
        MsgBox "The Statement 1 data contains missing or incorrect information, so the operation cannot be started.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    
    For i = IlkSira To SonSira
        If Cells(i, 14).Value = "x" Then
            MsgBox "The report data contains missing or incorrect information, so the operation cannot be started.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
            GoTo Son
        End If
    Next i
    
    If Cells(IlkSira, 15).Value = "x" Then
        MsgBox "The Statement 2 data sent to XXXMud contains missing or incorrect information, so the operation cannot be started.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If

    
    'Pathfinder...
    AutoPath = ThisWorkbook.Path
    DestOperasyon = AutoPath & "\System Files\Operation\"

    'ÜST YAZI TANIMLARI
    SourceUstYaziNormal = AutoPath & "\System Files\System Templates\Report 2 Cover Letter Templates\XXX Directorate Cover Letter.docm"
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
    Set IlkSiraBul = Range("CM7:CM100000").Find(What:=Cells(ActiveCell.Row, 5).Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    Set SonSiraBul = Range("CN7:CN100000").Find(What:=Cells(ActiveCell.Row, 5).Value, SearchDirection:=xlNext, _
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
    
    'Tutanak2 check
    If Cells(IlkSira, 220).Value = "" Then
        MsgBox "Cover letter for sequence number " & Cells(IlkSira, 5).Value & " cannot be prepared because the Statement 2 sent to XXXMud has not been created. Please create Statement 2 for sequence number " & Cells(IlkSira, 5).Value & " and try again.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If

    
    'On Error Resume Next
    ReNameUstYaziNormal = Cells(ActiveCell.Row, 5).Value & "-" & Cells(6, 16).Value
'________________________________________

    If TumDoc = False Then
        'Close the all Word application
        Call ModuleReport2.OpenWordControl
        
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
    'Kurum içi
    objDoc.Tables(1).Cell(Row:=4, Column:=1).Range.Text = Worksheets(2).Cells(6, 99).Value & " Unit"

    Birimx = UCase(Replace(Replace(Worksheets(2).Cells(6, 99).Value, "i", "I"), "ı", "I"))
    'Üst yazı tarihi
    objDoc.Tables(1).Cell(Row:=1, Column:=2).Range.Text = Birimx & ", " & FormatEnglishDate(Cells(ActiveCell.Row, 175).Value) '(Format(Cells(ActiveCell.Row, 175).Value, "d mmmm yyyy"))
    'Yazı no
    objDoc.Tables(1).Cell(Row:=6, Column:=2).Range.Text = Cells(ActiveCell.Row, 176).Value
    'Sigorta türü
    Ek2 = Cells(ActiveCell.Row, 184).Value
    Ek2 = Right(Ek2, Len(Ek2) - InStr(Ek2, "/"))
    objDoc.Tables(1).Cell(Row:=9, Column:=2).Range.Text = Ek2

    'Muhatap
    'Muhatap şablonun içinde.
    ustbilgitarih = (Format(Cells(ActiveCell.Row, 175).Value, "dd.mm.yyyy"))
    ustbilgisayi = Cells(ActiveCell.Row, 176).Value
    ustbilgimuhatap = "XXX Directorate"

            
    'İlgi tema
    If Cells(IlkSira, 33).Value = "Provincial Directorate B" Then
        If Cells(IlkSira, 34).Value <> "" Then
            GelenTema = Cells(IlkSira, 25).Value & " Provincial Governorship Provincial Directorate B " & Cells(IlkSira, 34).Value
        Else
            GelenTema = Cells(IlkSira, 25).Value & " Provincial Governorship Provincial Directorate B"
        End If
    ElseIf Cells(IlkSira, 33).Value = "District Directorate B" Then
        If Cells(IlkSira, 34).Value <> "" Then
            GelenTema = Cells(IlkSira, 26).Value & " District Governorship District Directorate B " & Cells(IlkSira, 34).Value
        Else
            GelenTema = Cells(IlkSira, 26).Value & " District Governorship District Directorate B"
        End If
    ElseIf Cells(IlkSira, 33).Value = "Provincial Directorate C" Then
        If Cells(IlkSira, 34).Value <> "" Then
            GelenTema = Cells(IlkSira, 25).Value & " Provincial Governorship Provincial Directorate C " & Cells(IlkSira, 34).Value
        Else
            GelenTema = Cells(IlkSira, 25).Value & " Provincial Governorship Provincial Directorate C"
        End If
    ElseIf Cells(IlkSira, 33).Value = "District Directorate C" Then
        If Cells(IlkSira, 34).Value <> "" Then
            GelenTema = Cells(IlkSira, 26).Value & " District Governorship District Directorate C " & Cells(IlkSira, 34).Value
        Else
            GelenTema = Cells(IlkSira, 26).Value & " District Governorship District Directorate C"
        End If
    ElseIf Cells(IlkSira, 33).Value = "Provincial Directorate D" Then
        If Cells(IlkSira, 34).Value <> "" Then
            GelenTema = Cells(IlkSira, 25).Value & " Provincial Governorship Provincial Directorate D " & Cells(IlkSira, 34).Value
        Else
            GelenTema = Cells(IlkSira, 25).Value & " Provincial Governorship Provincial Directorate D"
        End If
    ElseIf Cells(IlkSira, 33).Value = "District Directorate D" Then
        If Cells(IlkSira, 34).Value <> "" Then
            GelenTema = Cells(IlkSira, 26).Value & " District Governorship District Directorate D " & Cells(IlkSira, 34).Value
        Else
            GelenTema = Cells(IlkSira, 26).Value & " District Governorship District Directorate D"
        End If
    ElseIf Cells(IlkSira, 33).Value = "Provincial Directorate E" Then
        If Cells(IlkSira, 34).Value <> "" Then
            GelenTema = Cells(IlkSira, 25).Value & " Provincial Governorship Provincial Directorate E " & Cells(IlkSira, 34).Value
        Else
            GelenTema = Cells(IlkSira, 25).Value & " Provincial Governorship Provincial Directorate E"
        End If
    ElseIf Cells(IlkSira, 33).Value = "District Directorate E" Then
        If Cells(IlkSira, 34).Value <> "" Then
            GelenTema = Cells(IlkSira, 26).Value & " District Governorship District Directorate E " & Cells(IlkSira, 34).Value
        Else
            GelenTema = Cells(IlkSira, 26).Value & " District Governorship District Directorate E"
        End If
    ElseIf InStr(Cells(IlkSira, 33).Value, "General Directorate") <> 0 Or InStr(Cells(IlkSira, 33).Value, "Regional Directorate") <> 0 Then
        If Cells(IlkSira, 34).Value <> "" Then
            GelenTema = WorksheetFunction.Proper(ThisWorkbook.Worksheets(2).Cells(6, 111).Value) & " " & Cells(IlkSira, 33).Value & " " & Cells(IlkSira, 34).Value
        Else
            GelenTema = WorksheetFunction.Proper(ThisWorkbook.Worksheets(2).Cells(6, 111).Value) & " " & Cells(IlkSira, 33).Value
        End If
    Else
        If InStr(Cells(IlkSira, 33).Value, "X.X. ") > 0 Then
            If Cells(IlkSira, 34).Value <> "" Then
                GelenTema = Mid(Cells(IlkSira, 33).Value, 6, Len(Cells(IlkSira, 33).Value)) & " " & Cells(IlkSira, 34).Value
            Else
                GelenTema = Mid(Cells(IlkSira, 33).Value, 6, Len(Cells(IlkSira, 33).Value))
            End If
        Else
            If Cells(IlkSira, 34).Value <> "" Then
                GelenTema = Cells(IlkSira, 33).Value & " " & Cells(IlkSira, 34).Value
            Else
                GelenTema = Cells(IlkSira, 33).Value
            End If
        End If
    End If

    
    'İlgi
    'RaporIlgi = GelenTema & "n" & Right(GelenTema, 1) & "n " & Cells(IlkSira, 28).Value & " tarihli ve " & Cells(IlkSira, 29).Value & " sayılı yazısı."
    RaporIlgi = "The letter from the " & GelenTema & ", dated " & Cells(IlkSira, 28).Value & ", reference number " & Cells(IlkSira, 29).Value & "."
    objDoc.Tables(1).Cell(Row:=19, Column:=3).Range.Text = RaporIlgi

    'Gövde metni
    AdetTakip = 0
    NomTakip = 0
    TipATakip = ""
    AdetTopla = 0
    y = 0
    For i = IlkSira To SonSira
        If Cells(i, 62).Value = "invalid" And Left(Cells(i, 63).Value, 11) = "Technique A" Then
            y = y + 1
            If y = 1 Then
                AdetTakip = Cells(i, 52).Value
                TipATakip = Cells(i, 46).Value
                TipATakip = Right(TipATakip, Len(TipATakip) - InStr(TipATakip, "-")) 'Bu karakterin solunda bir boşluk oluşuyor.
                NomTakip = Cells(i, 49).Value
                If AdetTakip > 1 Then
                    Ek1 = AdetTakip & " pieces of " & NomTakip & TipATakip
                Else
                    Ek1 = AdetTakip & " piece of " & NomTakip & TipATakip
                End If
                AdetTopla = AdetTopla + Cells(i, 52).Value
                '2 pieces of 100 X1
            ElseIf y > 1 Then
                AdetTakip = Cells(i, 52).Value
                TipATakip = Cells(i, 46).Value
                TipATakip = Right(TipATakip, Len(TipATakip) - InStr(TipATakip, "-"))
                NomTakip = Cells(i, 49).Value
                If AdetTakip > 1 Then
                    Ek1 = Ek1 & ", " & AdetTakip & " pieces of " & NomTakip & TipATakip
                Else
                    Ek1 = Ek1 & ", " & AdetTakip & " piece of " & NomTakip & TipATakip
                End If
                AdetTopla = AdetTopla + Cells(i, 52).Value
                '2 pieces of 100 X1, 3 pieces of 200 X1, 5 pieces of 50 X3
            End If
        End If
    Next i
    'Birlestir = Chr(9) & Ek1
    'objDoc.Tables(2).Cell(Row:=1, Column:=1).Range.Text = Birlestir
    
    'TipA adedi (Tekil-Çoğul)
    If AdetTopla = 1 Then
        Ek2 = "item"
    Else
        Ek2 = "items"
    End If

    'Metni oluştur
    Bolum1 = "The Type A " & Ek2 & " (" & Ek1 & "), sent to our unit as an enclosure to the referenced letter and " & _
    "determined to be invalid following the examination, is being forwarded to you to enable the determination of xxxxxxx xxxxxxx xxxxxxx/xxxxxxx xxxxxxx xxx xxxxxxx xxxxxxx xxxxxxx xxxxxxx xxxxxxx xxxxxxx xxxxxxx xxxxxxx/xxxxxxx xxxxxxx and " & _
    "to facilitate the preparation of the related Report 2.2 concerning the said " & Ek2 & "."
    Bolum2 = "We kindly request that the mentioned Type A " & Ek2 & " be returned to us along with the Report 2.2 to be issued, so that it can be forwarded to the relevant unit."
    Bolum3 = "Respectfully submitted for your information and necessary action."
    
    objDoc.Tables(2).Cell(Row:=1, Column:=1).Range.Text = Chr(9) & Bolum1
    With objDoc.Tables(2).Cell(Row:=3, Column:=1).Range
        .Text = Chr(9) & Bolum2 & vbNewLine & Chr(9) & Bolum3
        .Paragraphs.Last.Range.ParagraphFormat.SpaceBefore = 6 ' 6 pt before spacing
    End With

    'Kurum içi imza
    'Birim
    'Unvan ve ad soyad birleşik mi ayrı mı
    If Cells(IlkSira, 170).Value = "Birleşik" Then
        Birimx = UCase(Replace(Replace(Worksheets(2).Cells(6, 99).Value, "i", "I"), "ı", "I")) & " UNIT"
        If Cells(IlkSira, 169).Value = "Authorized Signature" Then
            objDoc.Tables(3).Cell(Row:=2, Column:=2).Range.Text = Birimx & " MANAGER"
        ElseIf Cells(IlkSira, 169).Value = "Proxy Signature" Then
            objDoc.Tables(3).Cell(Row:=2, Column:=2).Range.Text = Birimx & " MANAGER (P)"
        ElseIf Cells(IlkSira, 169).Value = "Temporary Signature" Then
            objDoc.Tables(3).Cell(Row:=2, Column:=2).Range.Text = Birimx & " MANAGER (T)"
        ElseIf Cells(IlkSira, 169).Value = "Representative Signature" Then
            objDoc.Tables(3).Cell(Row:=2, Column:=2).Range.Text = Birimx & " MANAGER (R)"
        Else
            objDoc.Tables(3).Cell(Row:=2, Column:=2).Range.Text = Birimx & " MANAGER"
        End If
        objDoc.Tables(2).Cell(Row:=4, Column:=2).Range.Text = ""
        objDoc.Tables(3).Cell(Row:=3, Column:=2).Range.Text = UCase(Replace(Replace(Cells(ActiveCell.Row, 139).Value, "i", "I"), "ı", "I")) 'Ad Soyad büyük harf
    Else 'If Cells(IlkSira, 170).Value = "Ayrı" Then
        Birimx = UCase(Replace(Replace(Worksheets(2).Cells(6, 99).Value, "i", "I"), "ı", "I")) & " UNIT"
        If Cells(IlkSira, 169).Value = "Authorized Signature" Then
            objDoc.Tables(2).Cell(Row:=4, Column:=2).Range.Text = Birimx & " MANAGER"
        ElseIf Cells(IlkSira, 169).Value = "Proxy Signature" Then
            objDoc.Tables(2).Cell(Row:=4, Column:=2).Range.Text = Birimx & " MANAGER (P)"
        ElseIf Cells(IlkSira, 169).Value = "Temporary Signature" Then
            objDoc.Tables(2).Cell(Row:=4, Column:=2).Range.Text = Birimx & " MANAGER (T)"
        ElseIf Cells(IlkSira, 169).Value = "Representative Signature" Then
            objDoc.Tables(2).Cell(Row:=4, Column:=2).Range.Text = Birimx & " MANAGER (R)"
        Else
            objDoc.Tables(2).Cell(Row:=4, Column:=2).Range.Text = Birimx & " MANAGER"
        End If
        objDoc.Tables(3).Cell(Row:=4, Column:=2).Range.Text = UCase(Replace(Replace(Cells(ActiveCell.Row, 139).Value, "i", "I"), "ı", "I")) 'Ad Soyad büyük harf
    End If
        
    
    'Ekler
    objDoc.Tables(3).Cell(Row:=8, Column:=1).Range.Text = "Attachment:"
        
    Ek2 = Cells(ActiveCell.Row, 184).Value 'Kapalı Package A
    Ek2 = Left(Ek2, InStr(Ek2, "/") - 1)
    If Cells(ActiveCell.Row, 185).Value > 1 Then objDoc.Tables(3).Cell(Row:=8, Column:=2).Range.Text = " 1) Enclosed " & Ek2 & " (" & Cells(ActiveCell.Row, 185).Value & " pieces)"
    If Cells(ActiveCell.Row, 185).Value < 2 Then objDoc.Tables(3).Cell(Row:=8, Column:=2).Range.Text = " 1) Enclosed " & Ek2 & " (" & Cells(ActiveCell.Row, 185).Value & " piece)"

    x = Application.Sum(Range(Cells(IlkSira, 220), Cells(SonSira, 220))) 'XXXMud Giden Statement 2 toplam sayfa sayısı
    If x > 1 Then objDoc.Tables(3).Cell(Row:=9, Column:=2).Range.Text = " 2) Statement 2 (" & x & " pages)"
    If x < 2 Then objDoc.Tables(3).Cell(Row:=9, Column:=2).Range.Text = " 2) Statement 2 (" & x & " page)"

    'İlgi yazı fotokopisi
    'objDoc.Tables(3).Cell(Row:=8, Column:=2).Range.Text = " 1) İlgi Yazı Fotokopisi " & " (" & Cells(ActiveCell.Row, 206).Value & " page(s))"
    If Cells(ActiveCell.Row, 206).Value > 1 Then objDoc.Tables(3).Cell(Row:=10, Column:=2).Range.Text = " 3) Photocopy of the Referenced Letter (" & Cells(IlkSira, 206).Value & " pages)"
    If Cells(ActiveCell.Row, 206).Value < 2 Then objDoc.Tables(3).Cell(Row:=10, Column:=2).Range.Text = " 3) Photocopy of the Referenced Letter (" & Cells(IlkSira, 206).Value & " page)"
 
    
     '__________________________DİNAMİK SAYFA YAPISI
    
    'İlgi ve gövde metninin kaç satırdan oluştuğu bilgisinin elde edilmesinde kullanılıyor.
    Set IlgiRng = objDoc.Tables(1).Cell(Row:=19, Column:=3).Range
    Set Govde1Rng = objDoc.Tables(2).Cell(Row:=1, Column:=1).Range
    Set Govde2Rng = objDoc.Tables(2).Cell(Row:=3, Column:=1).Range
    
    IlgiStrSay = Govde1Rng.Information(wdFirstCharacterLineNumber) - IlgiRng.Information(wdFirstCharacterLineNumber) - 1
    Govde1StrSay = Govde2Rng.Information(wdFirstCharacterLineNumber) - Govde1Rng.Information(wdFirstCharacterLineNumber)
    'MsgBox "İlgi Satır Sayısı: " & IlgiStrSay & vbNewLine & vbNewLine & "Gövde 1 Satır Sayısı: " & Govde1StrSay
    IlgiFarkSay = IlgiStrSay - 1 'Sıfırlandı, 1 satır varsayılan
    Govde1FarkSay = Govde1StrSay - 4 'Sıfırlandı, 4 satır varsayılan
    CokluSayfa = 0
    'MsgBox Govde1FarkSay + IlgiFarkSay 'Varsayılan ilgi ve 1 paragrafta (2 rowda) toplam 5 satıra göre sıfırlandı.
    
    
    'Dinamik sayfa düzeni
    objDoc.Tables(2).Rows(2).Delete 'gövde ara satır
    For i = 1 To 2
        objDoc.Tables(3).Rows(11).Delete 'Ek sonrası
    Next i

    If Cells(IlkSira, 170).Value = "Birleşik" Then 'Merged signature area
        Govde1FarkSay = Govde1FarkSay + 5
    Else 'Split signature area
        Govde1FarkSay = Govde1FarkSay + 5
    End If

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


 
'MsgBox "Total Sayfa Öncesi: " & TotalSayfaUstYazi

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
    Cells(ActiveCell.Row, 221).Value = TotalSayfaUstYazi
    
    'MsgBox "Total Sayfa Sonrası: " & TotalSayfaUstYazi
    
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

'Worksheets(4).Activate
Application.ScreenUpdating = True
Application.DisplayAlerts = True
'Application.EnableEvents = True

'End If

End Sub

Sub Rapor2_2BilgilendirmeUstYazi()

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
Dim Ek1 As String, Ek2 As String, Ek3 As String, Ek4 As String, Ek5 As String, Ek6 As String, Ek7 As String, Birlestir As String

Dim ReNameRaporNormal As String, SiraNoIlkSatir As Long, RaporIlgi As String
Dim AltRaporNoIlk As Long, AltRaporNoSon As Long, AdetTopla As Long
Dim TekCogulTipA As String, k As Integer
Dim MyRange As Object, TxtFileRapor As String, TotalSayfaRapor As String
Dim ReNameTutanak2Normal As String, GidenTema As String
Dim TxtFileTutanak2 As String, TotalSayfaTutanak2 As String
Dim ReNameUstYaziNormal As String
Dim TxtFileUstYazi As String, TotalSayfaUstYazi As String
Dim Birimx As String, UserName As String
Dim gecersizSay As Long
Dim IlBuyukHarf, IlKucukHarf, IlceBuyukHarf, IlceKucukHarf As String

Dim AdetTakip As Integer, TipATakip As String, NomTakip As Integer, TemaNo As String
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
If InStr(Cells(ActiveCell.Row, 204).Value, " Organization A") <> 0 Then
    IlceSakla = Cells(ActiveCell.Row, 204).Value
    Cells(ActiveCell.Row, 204).Value = ""
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
    Set IlkSiraBul = Range("CM7:CM100000").Find(What:=Cells(j, 5).Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    Set SonSiraBul = Range("CN7:CN100000").Find(What:=Cells(j, 5).Value, SearchDirection:=xlNext, _
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
    
    If Cells(IlkSira, 13).Value = "x" Then
        MsgBox "The Statement 1 data contains missing or incorrect information, so the operation cannot be started.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    
    For i = IlkSira To SonSira
        If Cells(i, 14).Value = "x" Then
            MsgBox "The report data contains missing or incorrect information, so the operation cannot be started.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
            GoTo Son
        End If
    Next i
    
    If Cells(IlkSira, 15).Value = "x" Then
        MsgBox "The Statement 2 data sent to XXXMud contains missing or incorrect information, so the operation cannot be started.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    
    If Cells(IlkSira, 16).Value = "x" Then
        MsgBox "The cover letter data sent to XXXMud contains missing or incorrect information, so the operation cannot be started.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    
    'Pathfinder...
    AutoPath = ThisWorkbook.Path
    DestOperasyon = AutoPath & "\System Files\Operation\"

    'ÜST YAZI TANIMLARI
    SourceUstYaziNormal = AutoPath & "\System Files\System Templates\Report 2 Cover Letter Templates\Informative Cover Letter.docm"
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
    Set IlkSiraBul = Range("CM7:CM100000").Find(What:=Cells(ActiveCell.Row, 5).Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    Set SonSiraBul = Range("CN7:CN100000").Find(What:=Cells(ActiveCell.Row, 5).Value, SearchDirection:=xlNext, _
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
    'Report(s) check
    For i = IlkSira To SonSira
        If Cells(i, 14).Value <> "" Then
            If Cells(i, 219).Value = "" Then
                MsgBox "Report number " & Cells(i, 11).Value & " has not been created, so the cover letter for sequence number " & Cells(IlkSira, 5).Value & " cannot be prepared. Please create Report number " & Cells(i, 11).Value & " and try again.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
                GoTo Son
            End If
        End If
    Next i
    
    'XXXMud outgoing cover letter
    If Cells(IlkSira, 221).Value = "" Then
        MsgBox "Cover letter for sequence number " & Cells(IlkSira, 5).Value & " cannot be prepared because the outgoing cover letter to XXXMud has not been created. Please create the Statement 2 protocol for sequence number " & Cells(IlkSira, 5).Value & " and try again.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
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
        MsgBox UserName & " session is not registered in the system, so the cover letter cannot be created. Please register your session in the system using the Initials Interface located in the Settings Group of the Enterprise Document Automation System menu and try the operation again.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    
    'On Error Resume Next
    ReNameUstYaziNormal = Cells(ActiveCell.Row, 5).Value & "-" & Cells(6, 17).Value
'________________________________________

    If TumDoc = False Then
        'Close the all Word application
        Call ModuleReport2.OpenWordControl
        
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
    Ek2 = Cells(ActiveCell.Row, 81).Value
    Ek2 = Right(Ek2, Len(Ek2) - InStr(Ek2, "/"))
    objDoc.Tables(1).Cell(Row:=7, Column:=2).Range.Text = Ek2

    'Muhatap
    IlBuyukHarf = UCase(Replace(Replace(Cells(ActiveCell.Row, 203).Value, "i", "I"), "ı", "I"))
    IlKucukHarf = Cells(ActiveCell.Row, 203).Value
    If Cells(ActiveCell.Row, 204).Value <> "" Then
        IlceBuyukHarf = UCase(Replace(Replace(Cells(ActiveCell.Row, 204).Value, "i", "I"), "ı", "I"))
        IlceKucukHarf = Cells(ActiveCell.Row, 204).Value
    Else
        IlceBuyukHarf = ""
        IlceKucukHarf = ""
    End If
    If Cells(ActiveCell.Row, 204).Value <> "" Then
        Bolum3 = IlceKucukHarf & "/" & IlBuyukHarf
    Else
        Bolum3 = IlBuyukHarf
    End If
    
    
    'YENİ MUHATAP TEMASI
    
    M2 = False
    M3 = False
    M4 = False
    'TipB = False
    BStr = 10
    '1. sayfadan sonraki üst bilgi tanımları
    ustbilgitarih = (Format(Cells(ActiveCell.Row, 83).Value, "dd.mm.yyyy"))
    ustbilgisayi = Cells(ActiveCell.Row, 84).Value
    
    '4'lük
    If Cells(ActiveCell.Row, 200).Value <> "" Then
        If Cells(ActiveCell.Row, 199).Value = "Provincial Directorate B" Or Cells(ActiveCell.Row, 199).Value = "Provincial Directorate C" Or _
        Cells(ActiveCell.Row, 199).Value = "Provincial Directorate D" Or Cells(ActiveCell.Row, 199).Value = "Provincial Directorate E" Then 'VALİLİK
            objDoc.Tables(1).Cell(Row:=BStr, Column:=1).Range.Text = IlBuyukHarf & " PROVINCIAL GOVERNORSHIP"
            objDoc.Tables(1).Cell(Row:=BStr + 1, Column:=1).Range.Text = Cells(ActiveCell.Row, 199).Value  'UCase(Replace(Replace(Cells(ActiveCell.Row, 199).Value, "i", "I"), "ı", "I"))
            objDoc.Tables(1).Cell(Row:=BStr + 2, Column:=1).Range.Text = "(" & Cells(ActiveCell.Row, 200).Value & ")"
            objDoc.Tables(1).Cell(Row:=BStr + 3, Column:=1).Range.Text = Bolum3
            objDoc.Tables(1).Cell(Row:=BStr + 3, Column:=1).Range.Font.Underline = wdUnderlineSingle
            M4 = True
            
            ustbilgimuhatap = IlKucukHarf & " Provincial Governorship " & Cells(ActiveCell.Row, 199).Value
        ElseIf Cells(ActiveCell.Row, 199).Value = "District Directorate B" Or Cells(ActiveCell.Row, 199).Value = "District Directorate C" Or _
        Cells(ActiveCell.Row, 199).Value = "District Directorate D" Or Cells(ActiveCell.Row, 199).Value = "District Directorate E" Then 'KAYMAKAMLIK
            objDoc.Tables(1).Cell(Row:=BStr, Column:=1).Range.Text = IlceBuyukHarf & " DISTRICT GOVERNORSHIP"
            objDoc.Tables(1).Cell(Row:=BStr + 1, Column:=1).Range.Text = Cells(ActiveCell.Row, 199).Value 'UCase(Replace(Replace(Cells(ActiveCell.Row, 199).Value, "i", "I"), "ı", "I"))
            objDoc.Tables(1).Cell(Row:=BStr + 2, Column:=1).Range.Text = "(" & Cells(ActiveCell.Row, 200).Value & ")"
            objDoc.Tables(1).Cell(Row:=BStr + 3, Column:=1).Range.Text = Bolum3
            objDoc.Tables(1).Cell(Row:=BStr + 3, Column:=1).Range.Font.Underline = wdUnderlineSingle
            M4 = True
            
            ustbilgimuhatap = IlceKucukHarf & " District Governorship " & Cells(ActiveCell.Row, 199).Value
        ElseIf InStr(Cells(ActiveCell.Row, 199).Value, "General Directorate") <> 0 Or InStr(Cells(ActiveCell.Row, 199).Value, "Regional Directorate") <> 0 Then
            objDoc.Tables(1).Cell(Row:=BStr, Column:=1).Range.Text = UCase(Replace(Replace(ThisWorkbook.Worksheets(2).Cells(6, 111).Value, "i", "I"), "ı", "I"))
            objDoc.Tables(1).Cell(Row:=BStr + 1, Column:=1).Range.Text = Cells(ActiveCell.Row, 199).Value 'UCase(Replace(Replace(Cells(ActiveCell.Row, 199).Value, "i", "I"), "ı", "I"))
            objDoc.Tables(1).Cell(Row:=BStr + 2, Column:=1).Range.Text = "(" & Cells(ActiveCell.Row, 200).Value & ")"
            objDoc.Tables(1).Cell(Row:=BStr + 3, Column:=1).Range.Text = Bolum3
            objDoc.Tables(1).Cell(Row:=BStr + 3, Column:=1).Range.Font.Underline = wdUnderlineSingle
            M4 = True
            
            ustbilgimuhatap = WorksheetFunction.Proper(ThisWorkbook.Worksheets(2).Cells(6, 111).Value) & " " & Cells(ActiveCell.Row, 199).Value
        Else 'YARGI 3'lük
            Bolum1 = UCase(Replace(Replace(Cells(ActiveCell.Row, 199).Value, "i", "I"), "ı", "I"))
            If InStr(Bolum1, "X.X. ") > 0 Then
                objDoc.Tables(1).Cell(Row:=BStr, Column:=1).Range.Text = Mid(Bolum1, 6, Len(Bolum1))
                objDoc.Tables(1).Cell(Row:=BStr + 1, Column:=1).Range.Text = "(" & Cells(ActiveCell.Row, 200).Value & ")"
                objDoc.Tables(1).Cell(Row:=BStr + 2, Column:=1).Range.Text = Bolum3
                objDoc.Tables(1).Cell(Row:=BStr + 2, Column:=1).Range.Font.Underline = wdUnderlineSingle
                M3 = True
                
                ustbilgimuhatap = Mid(Cells(ActiveCell.Row, 199).Value, 6, Len(Cells(ActiveCell.Row, 199).Value))
            Else
                objDoc.Tables(1).Cell(Row:=BStr, Column:=1).Range.Text = Bolum1
                objDoc.Tables(1).Cell(Row:=BStr + 1, Column:=1).Range.Text = "(" & Cells(ActiveCell.Row, 200).Value & ")"
                objDoc.Tables(1).Cell(Row:=BStr + 2, Column:=1).Range.Text = Bolum3
                objDoc.Tables(1).Cell(Row:=BStr + 2, Column:=1).Range.Font.Underline = wdUnderlineSingle
                M3 = True

                ustbilgimuhatap = Cells(ActiveCell.Row, 199).Value
            End If
        End If
    End If
    '3'lük
    If Cells(ActiveCell.Row, 200).Value = "" Then
        If Cells(ActiveCell.Row, 199).Value = "Provincial Directorate B" Or Cells(ActiveCell.Row, 199).Value = "Provincial Directorate C" Or _
        Cells(ActiveCell.Row, 199).Value = "Provincial Directorate D" Or Cells(ActiveCell.Row, 199).Value = "Provincial Directorate E" Then 'VALİLİK
            objDoc.Tables(1).Cell(Row:=BStr, Column:=1).Range.Text = IlBuyukHarf & " PROVINCIAL GOVERNORSHIP"
            objDoc.Tables(1).Cell(Row:=BStr + 1, Column:=1).Range.Text = "(" & Cells(ActiveCell.Row, 199).Value & ")" 'UCase(Replace(Replace(Cells(ActiveCell.Row, 199).Value, "i", "I"), "ı", "I"))
            objDoc.Tables(1).Cell(Row:=BStr + 2, Column:=1).Range.Text = Bolum3
            objDoc.Tables(1).Cell(Row:=BStr + 2, Column:=1).Range.Font.Underline = wdUnderlineSingle
            M3 = True
            
            ustbilgimuhatap = IlKucukHarf & " Provincial Governorship " & Cells(ActiveCell.Row, 199).Value
        ElseIf Cells(ActiveCell.Row, 199).Value = "District Directorate B" Or Cells(ActiveCell.Row, 199).Value = "District Directorate C" Or _
        Cells(ActiveCell.Row, 199).Value = "District Directorate D" Or Cells(ActiveCell.Row, 199).Value = "District Directorate E" Then
            objDoc.Tables(1).Cell(Row:=BStr, Column:=1).Range.Text = IlceBuyukHarf & " DISTRICT GOVERNORSHIP"
            objDoc.Tables(1).Cell(Row:=BStr + 1, Column:=1).Range.Text = "(" & Cells(ActiveCell.Row, 199).Value & ")" 'UCase(Replace(Replace(Cells(ActiveCell.Row, 199).Value, "i", "I"), "ı", "I"))
            objDoc.Tables(1).Cell(Row:=BStr + 2, Column:=1).Range.Text = Bolum3
            objDoc.Tables(1).Cell(Row:=BStr + 2, Column:=1).Range.Font.Underline = wdUnderlineSingle
            M3 = True
            
            ustbilgimuhatap = IlceKucukHarf & " District Governorship " & Cells(ActiveCell.Row, 199).Value
        ElseIf InStr(Cells(ActiveCell.Row, 199).Value, "General Directorate") <> 0 Or InStr(Cells(ActiveCell.Row, 199).Value, "Regional Directorate") <> 0 Then
            objDoc.Tables(1).Cell(Row:=BStr, Column:=1).Range.Text = UCase(Replace(Replace(ThisWorkbook.Worksheets(2).Cells(6, 111).Value, "i", "I"), "ı", "I"))
            objDoc.Tables(1).Cell(Row:=BStr + 1, Column:=1).Range.Text = "(" & Cells(ActiveCell.Row, 199).Value & ")" 'UCase(Replace(Replace(Cells(ActiveCell.Row, 199).Value, "i", "I"), "ı", "I"))
            objDoc.Tables(1).Cell(Row:=BStr + 2, Column:=1).Range.Text = Bolum3
            objDoc.Tables(1).Cell(Row:=BStr + 2, Column:=1).Range.Font.Underline = wdUnderlineSingle
            M3 = True
            
            ustbilgimuhatap = WorksheetFunction.Proper(ThisWorkbook.Worksheets(2).Cells(6, 111).Value) & " " & Cells(ActiveCell.Row, 199).Value
        Else 'YARGI 2'lik
            Bolum1 = UCase(Replace(Replace(Cells(ActiveCell.Row, 199).Value, "i", "I"), "ı", "I"))
            If InStr(Bolum1, "X.X. ") > 0 Then
                objDoc.Tables(1).Cell(Row:=BStr, Column:=1).Range.Text = Mid(Bolum1, 6, Len(Bolum1))
                objDoc.Tables(1).Cell(Row:=BStr + 1, Column:=1).Range.Text = Bolum3
                objDoc.Tables(1).Cell(Row:=BStr + 1, Column:=1).Range.Font.Underline = wdUnderlineSingle
                M2 = True

                ustbilgimuhatap = Mid(Cells(ActiveCell.Row, 199).Value, 6, Len(Cells(ActiveCell.Row, 199).Value))
            Else
                objDoc.Tables(1).Cell(Row:=BStr, Column:=1).Range.Text = Bolum1
                objDoc.Tables(1).Cell(Row:=BStr + 1, Column:=1).Range.Text = Bolum3
                objDoc.Tables(1).Cell(Row:=BStr + 1, Column:=1).Range.Font.Underline = wdUnderlineSingle
                M2 = True

                ustbilgimuhatap = Cells(ActiveCell.Row, 199).Value
            End If
        End If
    End If
    
    'İlgi tema
    If Cells(IlkSira, 33).Value = "Provincial Directorate B" Then
        If Cells(IlkSira, 34).Value <> "" Then
            GelenTema = Cells(IlkSira, 25).Value & " Provincial Governorship Provincial Directorate B " & Cells(IlkSira, 34).Value
        Else
            GelenTema = Cells(IlkSira, 25).Value & " Provincial Governorship Provincial Directorate B"
        End If
    ElseIf Cells(IlkSira, 33).Value = "District Directorate B" Then
        If Cells(IlkSira, 34).Value <> "" Then
            GelenTema = Cells(IlkSira, 26).Value & " District Governorship District Directorate B " & Cells(IlkSira, 34).Value
        Else
            GelenTema = Cells(IlkSira, 26).Value & " District Governorship District Directorate B"
        End If
    ElseIf Cells(IlkSira, 33).Value = "Provincial Directorate C" Then
        If Cells(IlkSira, 34).Value <> "" Then
            GelenTema = Cells(IlkSira, 25).Value & " Provincial Governorship Provincial Directorate C " & Cells(IlkSira, 34).Value
        Else
            GelenTema = Cells(IlkSira, 25).Value & " Provincial Governorship Provincial Directorate C"
        End If
    ElseIf Cells(IlkSira, 33).Value = "District Directorate C" Then
        If Cells(IlkSira, 34).Value <> "" Then
            GelenTema = Cells(IlkSira, 26).Value & " District Governorship District Directorate C " & Cells(IlkSira, 34).Value
        Else
            GelenTema = Cells(IlkSira, 26).Value & " District Governorship District Directorate C"
        End If
    ElseIf Cells(IlkSira, 33).Value = "Provincial Directorate D" Then
        If Cells(IlkSira, 34).Value <> "" Then
            GelenTema = Cells(IlkSira, 25).Value & " Provincial Governorship Provincial Directorate D " & Cells(IlkSira, 34).Value
        Else
            GelenTema = Cells(IlkSira, 25).Value & " Provincial Governorship Provincial Directorate D"
        End If
    ElseIf Cells(IlkSira, 33).Value = "District Directorate D" Then
        If Cells(IlkSira, 34).Value <> "" Then
            GelenTema = Cells(IlkSira, 26).Value & " District Governorship District Directorate D " & Cells(IlkSira, 34).Value
        Else
            GelenTema = Cells(IlkSira, 26).Value & " District Governorship District Directorate D"
        End If
    ElseIf Cells(IlkSira, 33).Value = "Provincial Directorate E" Then
        If Cells(IlkSira, 34).Value <> "" Then
            GelenTema = Cells(IlkSira, 25).Value & " Provincial Governorship Provincial Directorate E " & Cells(IlkSira, 34).Value
        Else
            GelenTema = Cells(IlkSira, 25).Value & " Provincial Governorship Provincial Directorate E"
        End If
    ElseIf Cells(IlkSira, 33).Value = "District Directorate E" Then
        If Cells(IlkSira, 34).Value <> "" Then
            GelenTema = Cells(IlkSira, 26).Value & " District Governorship District Directorate E " & Cells(IlkSira, 34).Value
        Else
            GelenTema = Cells(IlkSira, 26).Value & " District Governorship District Directorate E"
        End If
    ElseIf InStr(Cells(IlkSira, 33).Value, "General Directorate") <> 0 Or InStr(Cells(IlkSira, 33).Value, "Regional Directorate") <> 0 Then
        If Cells(IlkSira, 34).Value <> "" Then
            GelenTema = WorksheetFunction.Proper(ThisWorkbook.Worksheets(2).Cells(6, 111).Value) & " " & Cells(IlkSira, 33).Value & " " & Cells(IlkSira, 34).Value
        Else
            GelenTema = WorksheetFunction.Proper(ThisWorkbook.Worksheets(2).Cells(6, 111).Value) & " " & Cells(IlkSira, 33).Value
        End If
    Else
        If InStr(Cells(IlkSira, 33).Value, "X.X. ") > 0 Then
            If Cells(IlkSira, 34).Value <> "" Then
                GelenTema = Mid(Cells(IlkSira, 33).Value, 6, Len(Cells(IlkSira, 33).Value)) & " " & Cells(IlkSira, 34).Value
            Else
                GelenTema = Mid(Cells(IlkSira, 33).Value, 6, Len(Cells(IlkSira, 33).Value))
            End If
        Else
            If Cells(IlkSira, 34).Value <> "" Then
                GelenTema = Cells(IlkSira, 33).Value & " " & Cells(IlkSira, 34).Value
            Else
                GelenTema = Cells(IlkSira, 33).Value
            End If
        End If
    End If


    'İlgi
    Bolum1 = "Delivered to us on "
    Bolum2 = ", with a letter dated "
    Bolum3 = " and reference number "
    Bolum4 = "."

    If Cells(ActiveCell.Row, 33).Value = Cells(ActiveCell.Row, 199).Value And _
    Cells(ActiveCell.Row, 34).Value = Cells(ActiveCell.Row, 200).Value And _
    Cells(ActiveCell.Row, 25).Value = Cells(ActiveCell.Row, 203).Value Then
        ' Gelen ve giden birim aynıysa
        Bolum5 = "a) " & Bolum1 & Cells(ActiveCell.Row, 36).Value & Bolum2 & _
            Cells(ActiveCell.Row, 28).Value & Bolum3 & _
            Cells(ActiveCell.Row, 29).Value & Bolum4
    Else
        ' Gelen ve giden birim farklıysa
        Bolum5 = "a) " & Bolum1 & Cells(ActiveCell.Row, 36).Value & _
                    " from " & GelenTema & Bolum2 & _
                    Cells(IlkSira, 28).Value & Bolum3 & _
                    Cells(IlkSira, 29).Value & Bolum4
    End If
    Bolum6 = "b) Our letter dated " & Cells(ActiveCell.Row, 175).Value & " and numbered " & Cells(ActiveCell.Row, 176).Value & ", addressed to the XXX Directorate of Organization A."
    
    'ilgileri ekle
    objDoc.Tables(1).Cell(Row:=15, Column:=3).Range.Text = Bolum5 & vbNewLine & Bolum6

    'Gövde metni
    AdetTakip = 0
    NomTakip = 0
    TipATakip = ""
    AdetTopla = 0
    y = 0
    For i = IlkSira To SonSira
        If Cells(i, 62).Value = "invalid" Or Cells(i, 62).Value = "valid" Then 'And Left(Cells(i, 63).Value, 11) = "Technique A" Then
            y = y + 1
            If y = 1 Then
                AdetTakip = Cells(i, 52).Value
                TipATakip = Cells(i, 46).Value
                TipATakip = Right(TipATakip, Len(TipATakip) - InStr(TipATakip, "-")) 'Bu karakterin solunda bir boşluk oluşuyor.
                NomTakip = Cells(i, 49).Value
                If AdetTakip > 1 Then
                    Ek1 = AdetTakip & " pieces of " & NomTakip & TipATakip
                Else
                    Ek1 = AdetTakip & " piece of " & NomTakip & TipATakip
                End If
                '2 pieces of 100 X1
                AdetTopla = AdetTopla + Cells(i, 52).Value
            ElseIf y > 1 Then
                AdetTakip = Cells(i, 52).Value
                TipATakip = Cells(i, 46).Value
                TipATakip = Right(TipATakip, Len(TipATakip) - InStr(TipATakip, "-"))
                NomTakip = Cells(i, 49).Value
                If AdetTakip > 1 Then
                    Ek1 = Ek1 & ", " & AdetTakip & " pieces of " & NomTakip & TipATakip
                Else
                    Ek1 = Ek1 & ", " & AdetTakip & " piece of " & NomTakip & TipATakip
                End If
                '2 pieces of 100 X1, 3 pieces of 200 X1, 5 pieces of 50 X3
                AdetTopla = AdetTopla + Cells(i, 52).Value
            End If
        End If
    Next i
    
    'TemaNo
    TemaNo = Cells(ActiveCell.Row, 31).Value

    ' Report body (English version)
    
    ' Report(s) control
    Dim ReportList() As String
    ' Rapor numaralarını diziye al
    y = 0
    For i = IlkSira To SonSira
        If Cells(i, 14).Value <> "" Then
            y = y + 1
            ReDim Preserve ReportList(1 To y)
            ReportList(y) = Cells(i, 11).Value
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
        Ek4 = "report"
        Ek5 = "Report 2.1"
        Ek6 = "Report 2.2"
        Bolum5 = "has"
    Else
        Ek4 = "reports"
        Ek5 = "Report 2.1s"
        Ek6 = "Report 2.2s"
        Bolum5 = "have"
    End If
    
    
    If AdetTopla = 1 Then
        Ek2 = "item"
        Bolum1 = "The Type A " & Ek2 & " (" & Ek1 & ") with theme number " & TemaNo & ", sent to our unit as an enclosure to the referenced letter, has been documented in "
        Ek7 = "has"
    Else
        Ek2 = "items"
        Bolum1 = "The Type A " & Ek2 & " (" & Ek1 & ") with theme number " & TemaNo & ", sent to our unit as an enclosure to the referenced letter, have been documented in "
        Ek7 = "have"
    End If
    

    ' Construct the text
    Bolum1 = Bolum1 & Ek5 & ", "
    Bolum2 = Bolum1 & "dated " & Cells(IlkSira, 68).Value & " and numbered " & Ek3 & ". " & "The corresponding " & Ek4 & " " & Bolum5 & " been sent to your office."
    Bolum3 = "Furthermore, the Type A " & Ek2 & ", found to be invalid following the examination, " & Ek7 & " been sent to Organization A – XXX Directorate as an enclosure to our letter referenced in (b), in order to determine xxxxxx xxxxxx xxxxxx / xxxxxxxxx xxxxxxx xxx xxxxxxxxx xxxxx xxxxx xxxxxx xxxxxx xxxxxx xxxxxx / xxxxxx xxxxxx and to enable the preparation of the related " & Ek6 & "." & " " & _
    "Upon receiving their response, " & Ek6 & " will be completed, and both the " & Ek2 & " and the corresponding " & Ek4 & " will be forwarded to your office accordingly."
    Bolum4 = "Respectfully submitted for your information."
    objDoc.Tables(2).Cell(Row:=1, Column:=1).Range.Text = Chr(9) & Bolum2
    objDoc.Tables(2).Cell(Row:=2, Column:=1).Range.Text = Chr(9) & Bolum3
    objDoc.Tables(2).Cell(Row:=3, Column:=1).Range.Text = Chr(9) & Bolum4

    'imzalar
    objDoc.Tables(3).Cell(Row:=4, Column:=2).Range.Text = Cells(ActiveCell.Row, 130).Value 'Ad Soyad1
    objDoc.Tables(3).Cell(Row:=5, Column:=2).Range.Text = Cells(ActiveCell.Row, 131).Value 'Unvan1
    objDoc.Tables(3).Cell(Row:=4, Column:=3).Range.Text = Cells(ActiveCell.Row, 133).Value 'Ad Soyad2
    objDoc.Tables(3).Cell(Row:=5, Column:=3).Range.Text = Cells(ActiveCell.Row, 134).Value 'Unvan2

    'Ekler
    objDoc.Tables(3).Cell(Row:=8, Column:=1).Range.Text = "Attachment:"
    
    'İlgi (a) yazınız fotokopisi
    If Cells(IlkSira, 206).Value > 1 Then objDoc.Tables(3).Cell(Row:=8, Column:=2).Range.Text = " 1) Photocopy of Your Letter Referenced in (a) (" & Cells(IlkSira, 206).Value & " pages)"
    If Cells(IlkSira, 206).Value < 2 Then objDoc.Tables(3).Cell(Row:=8, Column:=2).Range.Text = " 1) Photocopy of Your Letter Referenced in (a) (" & Cells(IlkSira, 206).Value & " page)"
    
    'İlgi (b) yazımız fotokopisi
    If Cells(IlkSira, 221).Value > 1 Then objDoc.Tables(3).Cell(Row:=9, Column:=2).Range.Text = " 2) Photocopy of Our Letter Referenced in (b) (" & Cells(IlkSira, 221).Value & " pages)"
    If Cells(IlkSira, 221).Value < 2 Then objDoc.Tables(3).Cell(Row:=9, Column:=2).Range.Text = " 2) Photocopy of Our Letter Referenced in (b) (" & Cells(IlkSira, 221).Value & " page)"

    x = Application.Sum(Range(Cells(IlkSira, 219), Cells(SonSira, 219)))  'Rapor 2.1 toplam sayfa sayısı
    If y = 1 Then
        If x > 1 Then objDoc.Tables(3).Cell(Row:=10, Column:=2).Range.Text = " 3) Report 2.1 (" & x & " pages)"
        If x < 2 Then objDoc.Tables(3).Cell(Row:=10, Column:=2).Range.Text = " 3) Report 2.1 (" & x & " page)"
    Else
        objDoc.Tables(3).Cell(Row:=10, Column:=2).Range.Text = " 3) Report 2.1 (" & y & " reports, " & "total of " & x & " pages)"
    End If
  
        
    '__________________________DİNAMİK SAYFA YAPISI


    'İlgi ve gövde metninin kaç satırdan oluştuğu bilgisinin elde edilmesinde kullanılıyor.
    Set IlgiRng = objDoc.Tables(1).Cell(Row:=15, Column:=3).Range
    Set Govde1Rng = objDoc.Tables(2).Cell(Row:=1, Column:=1).Range
    Set Govde2Rng = objDoc.Tables(2).Cell(Row:=2, Column:=1).Range
    
    IlgiStrSay = Govde1Rng.Information(wdFirstCharacterLineNumber) - IlgiRng.Information(wdFirstCharacterLineNumber) - 1
    Govde1StrSay = Govde2Rng.Information(wdFirstCharacterLineNumber) - Govde1Rng.Information(wdFirstCharacterLineNumber)
    'MsgBox "İlgi Satır Sayısı: " & IlgiStrSay & vbNewLine & vbNewLine & "Gövde 1 Satır Sayısı: " & Govde1StrSay
    IlgiFarkSay = IlgiStrSay - 1
    'Govde1FarkSay = Govde1StrSay - 2
    CokluSayfa = 0
    
    Govde1FarkSay = Govde1StrSay - 5
    'MsgBox IlgiFarkSay + Govde1FarkSay
    
    'Dinamik sayfa düzeni
    'Delete rows after attach.
    For i = 1 To 2
        objDoc.Tables(3).Rows(11).Delete 'Ek sonrası
    Next i
    
    If M4 = True Then

        objDoc.Range.Font.Size = 11
        objDoc.Tables(1).Cell(Row:=5, Column:=1).Range.Font.Size = 13
        objDoc.Tables(1).Cell(Row:=2, Column:=1).Range.Font.Size = 9
    
        Set IlgiRng = objDoc.Tables(1).Cell(Row:=15, Column:=3).Range
        Set Govde1Rng = objDoc.Tables(2).Cell(Row:=1, Column:=1).Range
        Set Govde2Rng = objDoc.Tables(2).Cell(Row:=2, Column:=1).Range
      
        IlgiStrSay = Govde1Rng.Information(wdFirstCharacterLineNumber) - IlgiRng.Information(wdFirstCharacterLineNumber) - 1
        Govde1StrSay = Govde2Rng.Information(wdFirstCharacterLineNumber) - Govde1Rng.Information(wdFirstCharacterLineNumber)
        IlgiFarkSay = IlgiStrSay - 1
        Govde1FarkSay = Govde1StrSay - 2
        CokluSayfa = 0

    
        If IlgiFarkSay + Govde1FarkSay < 17 Then
            Govde1FarkSay = Govde1FarkSay + 7
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
            Govde1FarkSay = Govde1FarkSay + 11
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
        
        Govde1FarkSay = Govde1FarkSay + 10
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

    
    'If CokluSayfa = 1 Then
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
    'End If

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
    Cells(ActiveCell.Row, 222).Value = TotalSayfaUstYazi

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
    Cells(ActiveCell.Row, 204).Value = IlceSakla
End If

'Worksheets(4).Activate
Application.ScreenUpdating = True
Application.DisplayAlerts = True
'Application.EnableEvents = True

'End If

End Sub

Sub Rapor2_2XXXMudTutanak1()

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
Dim ReNameUstYaziNormal As String, PaketTipi As String
Dim TxtFileUstYazi As String, TotalSayfaUstYazi As String
Dim Birimx As String, UserName As String, SourceTutanak1Farkli As String


'TUTANAK1 için prosedürü başlat
'If ActiveCell.Column = 6 Then
'    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    'Application.EnableEvents = False
    
    'Pathfinder...
    AutoPath = ThisWorkbook.Path
    DestOperasyon = AutoPath & "\System Files\Operation\"
    'TUTANAK1 TANIMLARI
    SourceTutanak1Normal = AutoPath & "\System Files\System Templates\Statement 1 Templates\Statement 1.docm"
    
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
    If Not Dir(SourceTutanak1Normal, vbDirectory) <> vbNullString Then
        MsgBox "Cannot access the directory " & SourceTutanak1Normal & ". The names of folders and/or files in this directory may have been changed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If

    'Operation klasörü içinde kullanıcı modülü klasörü yoksa oluştur.
    If Not Dir(DestOpUserFolder, vbDirectory) <> vbNullString Then
        MkDir DestOpUserFolder
    End If

    'On Error Resume Next
    Set IlkSiraBul = Range("CM7:CM100000").Find(What:=Cells(ActiveCell.Row, 5).Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    Set SonSiraBul = Range("CN7:CN100000").Find(What:=Cells(ActiveCell.Row, 5).Value, SearchDirection:=xlNext, _
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


    If Cells(IlkSira, 13).Value = "x" Then
        MsgBox "The Statement 1 data contains missing or incorrect information, so the operation cannot be started.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    
    For i = IlkSira To SonSira
        If Cells(i, 14).Value = "x" Then
            MsgBox "The report data contains missing or incorrect information, so the operation cannot be started.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
            GoTo Son
        End If
    Next i
    
    If Cells(IlkSira, 15).Value = "x" Then
        MsgBox "The Statement 2 data sent to XXXMud contains missing or incorrect information, so the operation cannot be started.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    
    If Cells(IlkSira, 16).Value = "x" Then
        MsgBox "The cover letter data sent to XXXMud contains missing or incorrect information, so the operation cannot be started.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    
    If Cells(IlkSira, 17).Value = "x" Then
        MsgBox "The data of the notification cover letter sent to the institution contains missing or incorrect information, so the operation cannot be started.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If

    'Pre-checks for pages to be specified in attachments
    'Statement 2 sent to XXXMud
    If Cells(IlkSira, 220).Value = "" Then
        MsgBox "Statement 1 related to the package received from XXXMud for sequence number " & Cells(IlkSira, 5).Value & " cannot be prepared because Statement 2 sent to XXXMud for sequence number " & Cells(IlkSira, 5).Value & " has not been created. Please create Statement 2 for sequence number " & Cells(IlkSira, 5).Value & " and try again.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If

    
    'On Error Resume Next
    ReNameTutanak1 = Cells(ActiveCell.Row, 5).Value & "-" & Cells(6, 19).Value
'    ReNameDL = Cells(ActiveCell.Row, 5).Value & "-" & "Dispatch List"
'________________________________________

    
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
    
    'Dosyayı şablondan operasyon klasörüne kopyala ve adını değiştir.
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.CopyFile (SourceTutanak1Normal), DestOpUserFolder & ReNameTutanak1 & ".docm", True
'_____________________________________

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
    objWord.Documents.Open FileName:=DestOpUserFolder & ReNameTutanak1 & ".docm"
    objWord.Visible = True
    objWord.Activate 'Ekrana getirir.
    'objDoc.Activate 'Ekrana getirmez.
    objWord.Application.WindowState = wdWindowStateMaximize
    Set objDoc = GetObject(DestOpUserFolder & ReNameTutanak1 & ".docm")
'________________________________________
    
    'Normal tutanak1 tutanağı
    'Birim
    Birimx = UCase(Replace(Replace(Worksheets(2).Cells(6, 99).Value, "i", "I"), "ı", "I")) & " UNIT"
    objDoc.Tables(1).Cell(Row:=1, Column:=1).Range.Text = Birimx
    'Tutanak1 tarihi
    objDoc.Tables(1).Cell(Row:=4, Column:=3).Range.Text = Cells(ActiveCell.Row, 190).Value
    'Geliş tarihi
    objDoc.Tables(1).Cell(Row:=5, Column:=3).Range.Text = Cells(ActiveCell.Row, 191).Value
    'By Hand/By Mail
    If Cells(ActiveCell.Row, 196).Value = "By Hand" Then
        objDoc.CheckElden.Value = True
    ElseIf Cells(ActiveCell.Row, 196).Value = "By Mail" Then
        objDoc.CheckPosta.Value = True
    End If
    'Öğe çıktı/çıkmadı/vb.
    If Cells(ActiveCell.Row, 192).Value = "a. Content as Expected" Then
        objDoc.CheckTam.Value = True
    ElseIf Cells(ActiveCell.Row, 192).Value = "b. Content Empty" Then
        objDoc.CheckYok.Value = True
    ElseIf Cells(ActiveCell.Row, 192).Value = "c. Only Specific Content Available" Then
        objDoc.CheckEksik.Value = True
    End If
    'Türkçe karakterleri düzelt
    objDoc.CheckElden.Enabled = False
    objDoc.CheckElden.Enabled = True
    objDoc.CheckPosta.Enabled = False
    objDoc.CheckPosta.Enabled = True
    objDoc.CheckTam.Enabled = False
    objDoc.CheckTam.Enabled = True
    objDoc.CheckYok.Enabled = False
    objDoc.CheckYok.Enabled = True
    objDoc.CheckEksik.Enabled = False
    objDoc.CheckEksik.Enabled = True

    'Gönderen
    GelenTema = "ORGANIZATION A XXX Directorate"
    objDoc.Tables(1).Cell(Row:=9, Column:=3).Range.Text = GelenTema
    
    'Gönderenin adresi
    objDoc.Tables(1).Cell(Row:=10, Column:=3).Range.Text = "Xxxxxxx/Xxxxxxx"
    'Gelen Yazının Tarih ve Sayısı (XXXMud'den gelen üst yazı)
    objDoc.Tables(1).Cell(Row:=11, Column:=3).Range.Text = Cells(ActiveCell.Row, 193).Value
    objDoc.Tables(1).Cell(Row:=11, Column:=5).Range.Text = Cells(ActiveCell.Row, 194).Value
    'Gönderinin Eki
    If Cells(ActiveCell.Row, 195).Value = "Package A" Then
        objDoc.CheckZarf.Value = True
    ElseIf Cells(ActiveCell.Row, 195).Value = "Package B" Then
        objDoc.CheckTorba.Value = True
    ElseIf Cells(ActiveCell.Row, 195).Value = "Package C" Then
        objDoc.CheckKoli.Value = True
    End If
    'Tabloyu doldur
    Set IlkSiraBul = Range("CM7:CM100000").Find(What:=Cells(ActiveCell.Row, 5).Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    Set SonSiraBul = Range("CN7:CN100000").Find(What:=Cells(ActiveCell.Row, 5).Value, SearchDirection:=xlNext, _
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
    
    
    If Cells(IlkSira, 187).Value = "All" Then 'TÜMÜ gönderilmişse
        'Tabloya satır ekle
        If SonSira - IlkSira > 0 Then
            With objDoc.Tables(2)
                For i = IlkSira To SonSira - 1
                    .Rows.Add BeforeRow:=objDoc.Tables(2).Rows(2)
                Next i
            End With
        End If
        'Tabloyu doldur.
        x = 0
        For i = 2 To SonSira - IlkSira + 2
            objDoc.Tables(2).Cell(Row:=i, Column:=1).Range.Text = i - 1 'Tablo sıra no
            objDoc.Tables(2).Cell(Row:=i, Column:=2).Range.Text = Cells(IlkSira + i - 2, 46).Value 'Öğe türü
            objDoc.Tables(2).Cell(Row:=i, Column:=3).Range.Text = Cells(IlkSira + i - 2, 49).Value 'Öğe değeri
            objDoc.Tables(2).Cell(Row:=i, Column:=4).Range.Text = Cells(IlkSira + i - 2, 52).Value 'Adet
            If Cells(IlkSira + i - 2, 55).Value = "Dispatch List" Then
                x = x + 1
                objDoc.Tables(2).Cell(Row:=i, Column:=5).Range.Text = "Listed in the Attachment to Statement 1" '"Listed in the Dispatch List" 'Öğe ID No
            Else
                objDoc.Tables(2).Cell(Row:=i, Column:=5).Range.Text = Cells(IlkSira + i - 2, 55).Value 'Öğe ID No
            End If
            objDoc.Tables(2).Cell(Row:=i, Column:=6).Range.Text = Cells(IlkSira, 31).Value  'Tema No (Temai her satıra yaz.)
            objDoc.Tables(2).Cell(Row:=i, Column:=7).Range.Text = Cells(IlkSira + i - 2, 58).Value 'Açıklama
        Next i
    ElseIf Cells(IlkSira, 187).Value = "Technique A" Then 'SADECE Technique A gönderilmişse
        'Tabloya satır ekle
        y = 0
        If SonSira - IlkSira > 0 Then
            For i = IlkSira To SonSira
                If Cells(i, 62).Value = "invalid" And Left(Cells(i, 63).Value, 11) = "Technique A" Then
                    y = y + 1
                End If
            Next i
            If y > 1 Then
                With objDoc.Tables(2)
                    For i = IlkSira To IlkSira + y - 2
                        .Rows.Add BeforeRow:=objDoc.Tables(2).Rows(2)
                    Next i
                End With
            End If
        End If
        'Tabloyu doldur.
        x = 0
        i = 1
        For j = IlkSira To SonSira
            If Cells(j, 62).Value = "invalid" And Left(Cells(j, 63).Value, 11) = "Technique A" Then
                i = i + 1
                objDoc.Tables(2).Cell(Row:=i, Column:=1).Range.Text = i - 1 'Tablo sıra no
                objDoc.Tables(2).Cell(Row:=i, Column:=2).Range.Text = Cells(j, 46).Value 'Öğe türü
                objDoc.Tables(2).Cell(Row:=i, Column:=3).Range.Text = Cells(j, 49).Value 'Öğe değeri
                objDoc.Tables(2).Cell(Row:=i, Column:=4).Range.Text = Cells(j, 52).Value 'Adet
                If Cells(j, 55).Value = "Dispatch List" Then
                    x = x + 1
                    objDoc.Tables(2).Cell(Row:=i, Column:=5).Range.Text = "Listed in the Attachment to Statement 1" '"Listed in the Dispatch List" 'Öğe ID No
                Else
                    objDoc.Tables(2).Cell(Row:=i, Column:=5).Range.Text = Cells(j, 55).Value 'Öğe ID No
                End If
                objDoc.Tables(2).Cell(Row:=i, Column:=6).Range.Text = Cells(IlkSira, 31).Value  'Tema No (Temai her satıra yaz.)
                objDoc.Tables(2).Cell(Row:=i, Column:=7).Range.Text = Cells(j, 58).Value 'Açıklama
            End If
        Next j
    End If
    'Ek olarak Dispatch List
    If x > 0 Then
        objDoc.Tables(4).Cell(Row:=5, Column:=1).Range.Text = "Attachment: Dispatch List (" & Cells(ActiveCell.Row, 44).Value & " page(s))"
    Else
        objDoc.Tables(4).Cell(Row:=5, Column:=1).Range.Text = Cells(ActiveCell.Row, 44).Value
    End If
        
    'imzalar
    objDoc.Tables(4).Cell(Row:=2, Column:=2).Range.Text = Cells(ActiveCell.Row, 154).Value 'Ad Soyad1
    objDoc.Tables(4).Cell(Row:=3, Column:=2).Range.Text = Cells(ActiveCell.Row, 155).Value 'Unvan1
    objDoc.Tables(4).Cell(Row:=2, Column:=3).Range.Text = Cells(ActiveCell.Row, 157).Value 'Ad Soyad2
    objDoc.Tables(4).Cell(Row:=3, Column:=3).Range.Text = Cells(ActiveCell.Row, 158).Value 'Unvan2
    
    'Alt bilgi ekle
    objDoc.Sections(1).Footers(wdHeaderFooterPrimary).Range.Tables(1).Cell(Row:=1, Column:=1).Range.Text = ReNameTutanak1

    'İmza boşluğunu sayfaya sığdırmak için düzenle
    If x > 9 And x < 14 Then
        For i = 1 To x - 9
            If i = 1 Then
                objDoc.Tables(3).Rows(1).Delete
            ElseIf i = 2 Then
                For j = 1 To 2
                    objDoc.Tables(3).Rows(1).Delete
                Next j
            ElseIf i = 3 Then
                For j = 1 To 2
                    objDoc.Tables(4).Rows(1).Delete
                Next j
            ElseIf i = 4 Then
                objDoc.Tables(3).Rows.Add BeforeRow:=objDoc.Tables(3).Rows(1)
                For j = 1 To 2
                    objDoc.Tables(4).Rows(1).Delete
                Next j
            End If
        Next i
    ElseIf x > 13 Then
        'Nothing
    End If
    
    'Sayfa sayısı kaydet komutuna bağlandı.
'    objDoc.Close SaveChanges:=True
'    objWord.Quit
    objDoc.Save
    'Sayfayı text dosyasından çek
    TxtFileTutanak1 = DestOpUserFolder & "Statement 1 Page Count.txt"
    #If UseADO Then
        Const adReadLine = -2&
        With CreateObject("ADODB.Stream")
            .Open
            .LoadFromFile TxtFileTutanak1
            Do Until .EOS
                TotalSayfaTutanak1 = .ReadText(adReadLine)
            Loop
            .Close
        End With
    #Else 'UseFSO
        Const Tutanak1TristateTrue = -1&
        With CreateObject("Scripting.FileSystemObject")
            With .OpenTextFile(TxtFileTutanak1, Format:=Tutanak1TristateTrue)
                Do Until .AtEndOfStream
                    TotalSayfaTutanak1 = .ReadLine
                Loop
                .Close
            End With
        End With
    #End If
    'MsgBox "Tutanak1: " & TotalSayfaTutanak1
    
    'Tutanak1 ve Döküm sayfa sayısı
    Cells(ActiveCell.Row, 223).Value = TotalSayfaTutanak1

    'Nesneleri temizle
    Set fso = Nothing
    Set objWord = Nothing
    Set objDoc = Nothing

Son:

'Nesneleri temizle
Set fso = Nothing
Set objWord = Nothing
Set objDoc = Nothing

'Worksheets(4).Activate
Application.ScreenUpdating = True
Application.DisplayAlerts = True
'Application.EnableEvents = True

'End If

End Sub

Sub Rapor2_2Tutanak2XXXMudGelen()

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
Dim Birimx As String, UserName As String

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
    Set IlkSiraBul = Range("CM7:CM100000").Find(What:=Cells(j, 5).Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    Set SonSiraBul = Range("CN7:CN100000").Find(What:=Cells(j, 5).Value, SearchDirection:=xlNext, _
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


    If Cells(IlkSira, 13).Value = "x" Then
        MsgBox "The Statement 1 data contains missing or incorrect information, so the operation cannot be started.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    
    For i = IlkSira To SonSira
        If Cells(i, 14).Value = "x" Then
            MsgBox "The report data contains missing or incorrect information, so the operation cannot be started.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
            GoTo Son
        End If
    Next i
    
    If Cells(IlkSira, 15).Value = "x" Then
        MsgBox "The Statement 2 data sent to XXXMud contains missing or incorrect information, so the operation cannot be started.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    
    If Cells(IlkSira, 16).Value = "x" Then
        MsgBox "The cover letter data sent to XXXMud contains missing or incorrect information, so the operation cannot be started.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    
    If Cells(IlkSira, 17).Value = "x" Then
        MsgBox "The data of the notification cover letter sent to the institution contains missing or incorrect information, so the operation cannot be started.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    
    If Cells(IlkSira, 19).Value = "x" Then
        MsgBox "The Statement 1 data related to the package received from XXXMud contains missing or incorrect information, so the operation cannot be started.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
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
    ReNameTutanak2Normal = Cells(ActiveCell.Row, 5).Value & "-" & Cells(6, 20).Value
'________________________________________
    
    
    'If TumDoc = False Then
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
    'End If
    
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
    objDoc.Tables(1).Cell(Row:=4, Column:=3).Range.Text = Cells(ActiveCell.Row, 198).Value
    'Belge tarihi ve numarası (XXXMud'den gelen yazının tarihi ve sayısı)
    objDoc.Tables(1).Cell(Row:=6, Column:=3).Range.Text = Cells(ActiveCell.Row, 193).Value
    objDoc.Tables(1).Cell(Row:=6, Column:=5).Range.Text = Cells(ActiveCell.Row, 194).Value

    'Gönderilen
    If Cells(ActiveCell.Row, 199).Value = "Provincial Directorate B" Then
        GidenTema = Cells(ActiveCell.Row, 203).Value & " Provincial Governorship Provincial Directorate B " & Cells(ActiveCell.Row, 200).Value
    ElseIf Cells(ActiveCell.Row, 199).Value = "District Directorate B" Then
        GidenTema = Cells(ActiveCell.Row, 204).Value & " District Governorship District Directorate B " & Cells(ActiveCell.Row, 200).Value
    ElseIf Cells(ActiveCell.Row, 199).Value = "Provincial Directorate C" Then
        GidenTema = Cells(ActiveCell.Row, 203).Value & " Provincial Governorship Provincial Directorate C " & Cells(ActiveCell.Row, 200).Value
    ElseIf Cells(ActiveCell.Row, 199).Value = "District Directorate C" Then
        GidenTema = Cells(ActiveCell.Row, 204).Value & " District Governorship District Directorate C " & Cells(ActiveCell.Row, 200).Value
    ElseIf Cells(ActiveCell.Row, 199).Value = "Provincial Directorate D" Then
        GidenTema = Cells(ActiveCell.Row, 203).Value & " Provincial Governorship Provincial Directorate D " & Cells(ActiveCell.Row, 200).Value
    ElseIf Cells(ActiveCell.Row, 199).Value = "District Directorate D" Then
        GidenTema = Cells(ActiveCell.Row, 204).Value & " District Governorship District Directorate D " & Cells(ActiveCell.Row, 200).Value
    ElseIf Cells(ActiveCell.Row, 199).Value = "Provincial Directorate E" Then
        GidenTema = Cells(ActiveCell.Row, 203).Value & " Provincial Governorship Provincial Directorate E " & Cells(ActiveCell.Row, 200).Value
    ElseIf Cells(ActiveCell.Row, 199).Value = "District Directorate E" Then
        GidenTema = Cells(ActiveCell.Row, 204).Value & " District Governorship District Directorate E " & Cells(ActiveCell.Row, 200).Value
    ElseIf InStr(Cells(ActiveCell.Row, 199).Value, "General Directorate") <> 0 Or InStr(Cells(ActiveCell.Row, 199).Value, "Regional Directorate") <> 0 Then
        If Cells(ActiveCell.Row, 200).Value <> "" Then
            GidenTema = WorksheetFunction.Proper(ThisWorkbook.Worksheets(2).Cells(6, 111).Value) & " " & Cells(ActiveCell.Row, 199).Value & " " & Cells(ActiveCell.Row, 200).Value
        Else
            GidenTema = WorksheetFunction.Proper(ThisWorkbook.Worksheets(2).Cells(6, 111).Value) & " " & Cells(ActiveCell.Row, 199).Value
        End If
    Else
        If Cells(ActiveCell.Row, 200).Value <> "" Then
            If InStr(Cells(ActiveCell.Row, 199).Value, "X.X. ") > 0 Then
                GidenTema = Mid(Cells(ActiveCell.Row, 199).Value, 6, Len(Cells(ActiveCell.Row, 199).Value)) & " " & Cells(ActiveCell.Row, 200).Value
            Else
                GidenTema = Cells(ActiveCell.Row, 199).Value & " " & Cells(ActiveCell.Row, 200).Value
            End If
        Else
            If InStr(Cells(ActiveCell.Row, 199).Value, "X.X. ") > 0 Then
                GidenTema = Mid(Cells(ActiveCell.Row, 199).Value, 6, Len(Cells(ActiveCell.Row, 199).Value))
            Else
                GidenTema = Cells(ActiveCell.Row, 199).Value
            End If
        End If
    End If
    
    objDoc.Tables(1).Cell(Row:=5, Column:=3).Range.Text = GidenTema

    'Tabloyu doldur
    Set IlkSiraBul = Range("CM7:CM100000").Find(What:=Cells(ActiveCell.Row, 5).Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    Set SonSiraBul = Range("CN7:CN100000").Find(What:=Cells(ActiveCell.Row, 5).Value, SearchDirection:=xlNext, _
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
    
    If Cells(IlkSira, 187).Value = "All" Then 'TÜMÜ gönderiliyor
        'Tabloya satır ekle
        If SonSira - IlkSira > 0 Then
            With objDoc.Tables(2)
                For i = IlkSira To SonSira - 1
                    .Rows.Add BeforeRow:=objDoc.Tables(2).Rows(2)
                Next i
            End With
        End If
        'Tabloyu doldur.
        i = 1
        Ek1 = 0
        For j = IlkSira To SonSira
            i = i + 1
            objDoc.Tables(2).Cell(Row:=i, Column:=1).Range.Text = i - 1 'Tablo sıra no
            objDoc.Tables(2).Cell(Row:=i, Column:=2).Range.Text = Cells(j, 46).Value 'Öğe türü
            objDoc.Tables(2).Cell(Row:=i, Column:=3).Range.Text = Cells(j, 49).Value 'Öğe değeri
            objDoc.Tables(2).Cell(Row:=i, Column:=4).Range.Text = Cells(j, 52).Value 'Adet
            Ek1 = Ek1 + Cells(j, 52).Value
            If Cells(j, 55).Value = "Dispatch List" Then
                objDoc.Tables(2).Cell(Row:=i, Column:=5).Range.Text = "Listed in the Attachment to Statement 1" '"Listed in the Dispatch List" 'Öğe ID No
            Else
                objDoc.Tables(2).Cell(Row:=i, Column:=5).Range.Text = Cells(j, 55).Value 'Öğe ID No
            End If
            objDoc.Tables(2).Cell(Row:=i, Column:=6).Range.Text = Cells(IlkSira, 31).Value  'Tema No (Temai her satıra yaz.)
            objDoc.Tables(2).Cell(Row:=i, Column:=7).Range.Text = Cells(j, 58).Value 'Açıklama
        Next j
    ElseIf Cells(IlkSira, 187).Value = "Technique A" Then 'SADECE Technique A gönderiliyor
        'Tabloya satır ekle
        y = 0
        If SonSira - IlkSira > 0 Then
            For i = IlkSira To SonSira
                If Cells(i, 62).Value = "invalid" And Left(Cells(i, 63).Value, 11) = "Technique A" Then
                    y = y + 1
                End If
            Next i
            If y > 1 Then
                With objDoc.Tables(2)
                    For i = IlkSira To IlkSira + y - 2
                        .Rows.Add BeforeRow:=objDoc.Tables(2).Rows(2)
                    Next i
                End With
            End If
        End If
        'Tabloyu doldur.
        i = 1
        Ek1 = 0
        For j = IlkSira To SonSira
            If Cells(j, 62).Value = "invalid" And Left(Cells(j, 63).Value, 11) = "Technique A" Then
                i = i + 1
                objDoc.Tables(2).Cell(Row:=i, Column:=1).Range.Text = i - 1 'Tablo sıra no
                objDoc.Tables(2).Cell(Row:=i, Column:=2).Range.Text = Cells(j, 46).Value 'Öğe türü
                objDoc.Tables(2).Cell(Row:=i, Column:=3).Range.Text = Cells(j, 49).Value 'Öğe değeri
                objDoc.Tables(2).Cell(Row:=i, Column:=4).Range.Text = Cells(j, 52).Value 'Adet
                Ek1 = Ek1 + Cells(j, 52).Value
                If Cells(j, 55).Value = "Dispatch List" Then
                    objDoc.Tables(2).Cell(Row:=i, Column:=5).Range.Text = "Listed in the Attachment to Statement 1" '"Listed in the Dispatch List" 'Öğe ID No
                Else
                    objDoc.Tables(2).Cell(Row:=i, Column:=5).Range.Text = Cells(j, 55).Value 'Öğe ID No
                End If
                objDoc.Tables(2).Cell(Row:=i, Column:=6).Range.Text = Cells(IlkSira, 31).Value  'Tema No (Temai her satıra yaz.)
                objDoc.Tables(2).Cell(Row:=i, Column:=7).Range.Text = Cells(j, 58).Value 'Açıklama
            End If
        Next j
    End If

    'Tutanak2 tutanağının metin kısmı
    'Tanımlamalar
    If CInt(Ek1) > 1 Then
        Ek4 = "items"
    Else
        Ek4 = "item"
    End If
    Bolum1 = vbTab & "A total of "
    Bolum2 = " " & Ek4 & " , as described above, have been placed inside "
    Bolum3 = " "
    Bolum4 = " and enclosed to be sent to the designated unit mentioned above."

    Ek2 = Cells(ActiveCell.Row, 202).Value
    Ek3 = Application.WorksheetFunction.Proper(Cells(ActiveCell.Row, 201).Value)
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
    objDoc.Tables(4).Cell(Row:=6, Column:=2).Range.Text = Cells(ActiveCell.Row, 160).Value 'Ad Soyad1
    objDoc.Tables(4).Cell(Row:=7, Column:=2).Range.Text = Cells(ActiveCell.Row, 161).Value 'Unvan1
    objDoc.Tables(4).Cell(Row:=6, Column:=3).Range.Text = Cells(ActiveCell.Row, 163).Value 'Ad Soyad2
    objDoc.Tables(4).Cell(Row:=7, Column:=3).Range.Text = Cells(ActiveCell.Row, 164).Value 'Unvan2
    
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
    Cells(ActiveCell.Row, 224).Value = TotalSayfaTutanak2

    objWord.Visible = True
    objWord.Activate
    
    'Nesneleri temizle
    Set fso = Nothing
    Set objWord = Nothing
    Set objDoc = Nothing
    
Son:

'Worksheets(4).Activate
Application.ScreenUpdating = True
Application.DisplayAlerts = True
'Application.EnableEvents = True

'End If

End Sub

Sub Rapor2_2Tutanak2IlgiliBirim()

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
Dim Birimx As String, UserName As String

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
    Set IlkSiraBul = Range("CM7:CM100000").Find(What:=Cells(j, 5).Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    Set SonSiraBul = Range("CN7:CN100000").Find(What:=Cells(j, 5).Value, SearchDirection:=xlNext, _
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


    If Cells(IlkSira, 13).Value = "x" Then
        MsgBox "The Statement 1 data contains missing or incorrect information, so the operation cannot be started.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    
    For i = IlkSira To SonSira
        If Cells(i, 14).Value = "x" Then
            MsgBox "The report data contains missing or incorrect information, so the operation cannot be started.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
            GoTo Son
        End If
    Next i
    
    If Cells(IlkSira, 15).Value = "x" Then
        MsgBox "The Statement 2 data sent to XXXMud contains missing or incorrect information, so the operation cannot be started.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    
    If Cells(IlkSira, 16).Value = "x" Then
        MsgBox "The cover letter data sent to XXXMud contains missing or incorrect information, so the operation cannot be started.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    
    If Cells(IlkSira, 17).Value = "x" Then
        MsgBox "The notification cover letter data sent to the institution contains missing or incorrect information, so the operation cannot be started.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If


    
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
    ReNameTutanak2Normal = Cells(ActiveCell.Row, 5).Value & "-" & Cells(6, 21).Value
'________________________________________
    
    
    'If TumDoc = False Then
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
    'End If
    
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
    objDoc.Tables(1).Cell(Row:=4, Column:=3).Range.Text = Cells(ActiveCell.Row, 71).Value
    'Belge tarihi ve numarası (Kurumdan gelen yazının tarihi ve sayısı)
    objDoc.Tables(1).Cell(Row:=6, Column:=3).Range.Text = Cells(ActiveCell.Row, 28).Value
    objDoc.Tables(1).Cell(Row:=6, Column:=5).Range.Text = Cells(ActiveCell.Row, 29).Value

    'Gönderilen
    If Cells(ActiveCell.Row, 199).Value = "Provincial Directorate B" Then
        GidenTema = Cells(ActiveCell.Row, 203).Value & " Provincial Governorship Provincial Directorate B " & Cells(ActiveCell.Row, 200).Value
    ElseIf Cells(ActiveCell.Row, 199).Value = "District Directorate B" Then
        GidenTema = Cells(ActiveCell.Row, 204).Value & " District Governorship District Directorate B " & Cells(ActiveCell.Row, 200).Value
    ElseIf Cells(ActiveCell.Row, 199).Value = "Provincial Directorate C" Then
        GidenTema = Cells(ActiveCell.Row, 203).Value & " Provincial Governorship Provincial Directorate C " & Cells(ActiveCell.Row, 200).Value
    ElseIf Cells(ActiveCell.Row, 199).Value = "District Directorate C" Then
        GidenTema = Cells(ActiveCell.Row, 204).Value & " District Governorship District Directorate C " & Cells(ActiveCell.Row, 200).Value
    ElseIf Cells(ActiveCell.Row, 199).Value = "Provincial Directorate D" Then
        GidenTema = Cells(ActiveCell.Row, 203).Value & " Provincial Governorship Provincial Directorate D " & Cells(ActiveCell.Row, 200).Value
    ElseIf Cells(ActiveCell.Row, 199).Value = "District Directorate D" Then
        GidenTema = Cells(ActiveCell.Row, 204).Value & " District Governorship District Directorate D " & Cells(ActiveCell.Row, 200).Value
    ElseIf Cells(ActiveCell.Row, 199).Value = "Provincial Directorate E" Then
        GidenTema = Cells(ActiveCell.Row, 203).Value & " Provincial Governorship Provincial Directorate E " & Cells(ActiveCell.Row, 200).Value
    ElseIf Cells(ActiveCell.Row, 199).Value = "District Directorate E" Then
        GidenTema = Cells(ActiveCell.Row, 204).Value & " District Governorship District Directorate E " & Cells(ActiveCell.Row, 200).Value
    ElseIf InStr(Cells(ActiveCell.Row, 199).Value, "General Directorate") <> 0 Or InStr(Cells(ActiveCell.Row, 199).Value, "Regional Directorate") <> 0 Then
        If Cells(ActiveCell.Row, 200).Value <> "" Then
            GidenTema = WorksheetFunction.Proper(ThisWorkbook.Worksheets(2).Cells(6, 111).Value) & " " & Cells(ActiveCell.Row, 199).Value & " " & Cells(ActiveCell.Row, 200).Value
        Else
            GidenTema = WorksheetFunction.Proper(ThisWorkbook.Worksheets(2).Cells(6, 111).Value) & " " & Cells(ActiveCell.Row, 199).Value
        End If
    Else
        If Cells(ActiveCell.Row, 200).Value <> "" Then
            If InStr(Cells(ActiveCell.Row, 199).Value, "X.X. ") > 0 Then
                GidenTema = Mid(Cells(ActiveCell.Row, 199).Value, 6, Len(Cells(ActiveCell.Row, 199).Value)) & " " & Cells(ActiveCell.Row, 200).Value
            Else
                GidenTema = Cells(ActiveCell.Row, 199).Value & " " & Cells(ActiveCell.Row, 200).Value
            End If
        Else
            If InStr(Cells(ActiveCell.Row, 199).Value, "X.X. ") > 0 Then
                GidenTema = Mid(Cells(ActiveCell.Row, 199).Value, 6, Len(Cells(ActiveCell.Row, 199).Value))
            Else
                GidenTema = Cells(ActiveCell.Row, 199).Value
            End If
        End If
    End If
    
    objDoc.Tables(1).Cell(Row:=5, Column:=3).Range.Text = GidenTema

    'Tabloyu doldur
    Set IlkSiraBul = Range("CM7:CM100000").Find(What:=Cells(ActiveCell.Row, 5).Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    Set SonSiraBul = Range("CN7:CN100000").Find(What:=Cells(ActiveCell.Row, 5).Value, SearchDirection:=xlNext, _
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

    'Tabloyu senaryoya göre doldur
    If Cells(IlkSira, 187).Value = "All" Then 'TÜMÜ gönderilmiş
        'Tabloya satır ekle
        If SonSira - IlkSira > 0 Then
            With objDoc.Tables(2)
                For i = IlkSira To SonSira - 1
                    .Rows.Add BeforeRow:=objDoc.Tables(2).Rows(2)
                Next i
            End With
        End If
        'Tabloyu doldur.
        i = 1
        Ek1 = 0
        For j = IlkSira To SonSira
            i = i + 1
            objDoc.Tables(2).Cell(Row:=i, Column:=1).Range.Text = i - 1 'Tablo sıra no
            objDoc.Tables(2).Cell(Row:=i, Column:=2).Range.Text = Cells(j, 46).Value 'Öğe türü
            objDoc.Tables(2).Cell(Row:=i, Column:=3).Range.Text = Cells(j, 49).Value 'Öğe değeri
            objDoc.Tables(2).Cell(Row:=i, Column:=4).Range.Text = Cells(j, 52).Value 'Adet
            Ek1 = Ek1 + Cells(j, 52).Value
            If Cells(j, 55).Value = "Dispatch List" Then
                objDoc.Tables(2).Cell(Row:=i, Column:=5).Range.Text = "Listed in the Attachment to Statement 1" '"Listed in the Dispatch List" 'Öğe ID No
            Else
                objDoc.Tables(2).Cell(Row:=i, Column:=5).Range.Text = Cells(j, 55).Value 'Öğe ID No
            End If
            objDoc.Tables(2).Cell(Row:=i, Column:=6).Range.Text = Cells(IlkSira, 31).Value  'Tema No (Temai her satıra yaz.)
            objDoc.Tables(2).Cell(Row:=i, Column:=7).Range.Text = Cells(j, 58).Value 'Açıklama
        Next j
    ElseIf Cells(IlkSira, 187).Value = "Technique A" And _
           Cells(IlkSira, 188).Value = "No" Then 'SADECE Technique A gönderilmiş ve tutanak1 olmadığı için sadece varlıkdakiler kapatılacak.
        'Tabloya satır ekle
        y = 0
        If SonSira - IlkSira > 0 Then
            For i = IlkSira To SonSira
                If Left(Cells(i, 63).Value, 8) <> "Technique A" Then 'Hem Technique A olmayan invalidler hem de validleri kapsıyor.
                    y = y + 1
                End If
            Next i
            If y > 1 Then
                With objDoc.Tables(2)
                    For i = IlkSira To IlkSira + y - 2
                        .Rows.Add BeforeRow:=objDoc.Tables(2).Rows(2)
                    Next i
                End With
            End If
        End If
        'Tabloyu doldur.
        i = 1
        Ek1 = 0
        For j = IlkSira To SonSira
            If Left(Cells(j, 63).Value, 8) <> "Technique A" Then
                i = i + 1
                objDoc.Tables(2).Cell(Row:=i, Column:=1).Range.Text = i - 1 'Tablo sıra no
                objDoc.Tables(2).Cell(Row:=i, Column:=2).Range.Text = Cells(j, 46).Value 'Öğe türü
                objDoc.Tables(2).Cell(Row:=i, Column:=3).Range.Text = Cells(j, 49).Value 'Öğe değeri
                objDoc.Tables(2).Cell(Row:=i, Column:=4).Range.Text = Cells(j, 52).Value 'Adet
                Ek1 = Ek1 + Cells(j, 52).Value
                If Cells(j, 55).Value = "Dispatch List" Then
                    objDoc.Tables(2).Cell(Row:=i, Column:=5).Range.Text = "Listed in the Attachment to Statement 1" '"Listed in the Dispatch List" 'Öğe ID No
                Else
                    objDoc.Tables(2).Cell(Row:=i, Column:=5).Range.Text = Cells(j, 55).Value 'Öğe ID No
                End If
                objDoc.Tables(2).Cell(Row:=i, Column:=6).Range.Text = Cells(IlkSira, 31).Value  'Tema No (Temai her satıra yaz.)
                objDoc.Tables(2).Cell(Row:=i, Column:=7).Range.Text = Cells(j, 58).Value 'Açıklama
            End If
        Next j
    ElseIf Cells(IlkSira, 187).Value = "Technique A" And _
           Cells(IlkSira, 188).Value = "Yes" And _
           Cells(IlkSira, 189).Value = "Yes" Then 'SADECE Technique A gönderilmiş, ama tutanak1 sonrasında öğelerin tümü birleştirileceği için tümü kapatılacak.
        'Tabloya satır ekle
        If SonSira - IlkSira > 0 Then
            With objDoc.Tables(2)
                For i = IlkSira To SonSira - 1
                    .Rows.Add BeforeRow:=objDoc.Tables(2).Rows(2)
                Next i
            End With
        End If
        'Tabloyu doldur.
        i = 1
        Ek1 = 0
        For j = IlkSira To SonSira
            i = i + 1
            objDoc.Tables(2).Cell(Row:=i, Column:=1).Range.Text = i - 1 'Tablo sıra no
            objDoc.Tables(2).Cell(Row:=i, Column:=2).Range.Text = Cells(j, 46).Value 'Öğe türü
            objDoc.Tables(2).Cell(Row:=i, Column:=3).Range.Text = Cells(j, 49).Value 'Öğe değeri
            objDoc.Tables(2).Cell(Row:=i, Column:=4).Range.Text = Cells(j, 52).Value 'Adet
            Ek1 = Ek1 + Cells(j, 52).Value
            If Cells(j, 55).Value = "Dispatch List" Then
                objDoc.Tables(2).Cell(Row:=i, Column:=5).Range.Text = "Listed in the Attachment to Statement 1" '"Listed in the Dispatch List" 'Öğe ID No
            Else
                objDoc.Tables(2).Cell(Row:=i, Column:=5).Range.Text = Cells(j, 55).Value 'Öğe ID No
            End If
            objDoc.Tables(2).Cell(Row:=i, Column:=6).Range.Text = Cells(IlkSira, 31).Value  'Tema No (Temai her satıra yaz.)
            objDoc.Tables(2).Cell(Row:=i, Column:=7).Range.Text = Cells(j, 58).Value 'Açıklama
        Next j
    ElseIf Cells(IlkSira, 187).Value = "Technique A" And _
           Cells(IlkSira, 188).Value = "Yes" And _
           Cells(IlkSira, 189).Value = "No" Then 'SADECE Technique A gönderilmiş, ama tutanak1 sonrasında öğelerin tümü birleştirilmeyeceği için sadece varlıkdakiler kapatılacak.
        'Tabloya satır ekle
        y = 0
        If SonSira - IlkSira > 0 Then
            For i = IlkSira To SonSira
                If Left(Cells(i, 63).Value, 8) <> "Technique A" Then 'Hem Technique A olmayan invalidler hem de validleri kapsıyor.
                    y = y + 1
                End If
            Next i
            If y > 1 Then
                With objDoc.Tables(2)
                    For i = IlkSira To IlkSira + y - 2
                        .Rows.Add BeforeRow:=objDoc.Tables(2).Rows(2)
                    Next i
                End With
            End If
        End If
        'Tabloyu doldur.
        i = 1
        Ek1 = 0
        For j = IlkSira To SonSira
            If Left(Cells(j, 63).Value, 8) <> "Technique A" Then
                i = i + 1
                objDoc.Tables(2).Cell(Row:=i, Column:=1).Range.Text = i - 1 'Tablo sıra no
                objDoc.Tables(2).Cell(Row:=i, Column:=2).Range.Text = Cells(j, 46).Value 'Öğe türü
                objDoc.Tables(2).Cell(Row:=i, Column:=3).Range.Text = Cells(j, 49).Value 'Öğe değeri
                objDoc.Tables(2).Cell(Row:=i, Column:=4).Range.Text = Cells(j, 52).Value 'Adet
                Ek1 = Ek1 + Cells(j, 52).Value
                If Cells(j, 55).Value = "Dispatch List" Then
                    objDoc.Tables(2).Cell(Row:=i, Column:=5).Range.Text = "Listed in the Attachment to Statement 1" '"Listed in the Dispatch List" 'Öğe ID No
                Else
                    objDoc.Tables(2).Cell(Row:=i, Column:=5).Range.Text = Cells(j, 55).Value 'Öğe ID No
                End If
                objDoc.Tables(2).Cell(Row:=i, Column:=6).Range.Text = Cells(IlkSira, 31).Value  'Tema No (Temai her satıra yaz.)
                objDoc.Tables(2).Cell(Row:=i, Column:=7).Range.Text = Cells(j, 58).Value 'Açıklama
            End If
        Next j
    End If


    'Tutanak2 tutanağının metin kısmı
    'Tanımlamalar
    If CInt(Ek1) > 1 Then
        Ek4 = "items"
    Else
        Ek4 = "item"
    End If
    Bolum1 = vbTab & "A total of "
    Bolum2 = " " & Ek4 & " , as described above, have been placed inside "
    Bolum3 = " "
    Bolum4 = " and enclosed to be sent to the designated unit mentioned above."

    Ek2 = Cells(ActiveCell.Row, 76).Value
    Ek3 = Application.WorksheetFunction.Proper(Cells(ActiveCell.Row, 75).Value)
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
    objDoc.Tables(4).Cell(Row:=6, Column:=2).Range.Text = Cells(ActiveCell.Row, 124).Value 'Ad Soyad1
    objDoc.Tables(4).Cell(Row:=7, Column:=2).Range.Text = Cells(ActiveCell.Row, 125).Value 'Unvan1
    objDoc.Tables(4).Cell(Row:=6, Column:=3).Range.Text = Cells(ActiveCell.Row, 127).Value 'Ad Soyad2
    objDoc.Tables(4).Cell(Row:=7, Column:=3).Range.Text = Cells(ActiveCell.Row, 128).Value 'Unvan2
    
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
    Cells(ActiveCell.Row, 225).Value = TotalSayfaTutanak2

    objWord.Visible = True
    objWord.Activate
    
    'Nesneleri temizle
    Set fso = Nothing
    Set objWord = Nothing
    Set objDoc = Nothing
    
Son:

'Worksheets(4).Activate
Application.ScreenUpdating = True
Application.DisplayAlerts = True
'Application.EnableEvents = True

'End If

End Sub

Sub Rapor2_2SonucUstYazi()

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
Dim Ek1 As String, Ek2 As String, Ek3 As String, Ek4 As String, Ek5 As String, Ek6 As String, Ek7 As String, Birlestir As String

Dim ReNameRaporNormal As String, SiraNoIlkSatir As Long, RaporIlgi As String
Dim AltRaporNoIlk As Long, AltRaporNoSon As Long, AdetTopla As Long
Dim TekCogulTipA As String, k As Integer
Dim MyRange As Object, TxtFileRapor As String, TotalSayfaRapor As String
Dim ReNameTutanak2Normal As String, GidenTema As String
Dim TxtFileTutanak2 As String, TotalSayfaTutanak2 As String
Dim ReNameUstYaziNormal As String
Dim TxtFileUstYazi As String, TotalSayfaUstYazi As String
Dim Birimx As String, UserName As String
Dim gecersizSay As Long
Dim IlBuyukHarf, IlKucukHarf, IlceBuyukHarf, IlceKucukHarf As String, rng As Range

Dim AdetTakip As Integer, TipATakip As String, NomTakip As Integer, TemaNo As String, Rapor2_2x As Integer
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
If InStr(Cells(ActiveCell.Row, 204).Value, " Organization A") <> 0 Then
    IlceSakla = Cells(ActiveCell.Row, 204).Value
    Cells(ActiveCell.Row, 204).Value = ""
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
    Set IlkSiraBul = Range("CM7:CM100000").Find(What:=Cells(j, 5).Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    Set SonSiraBul = Range("CN7:CN100000").Find(What:=Cells(j, 5).Value, SearchDirection:=xlNext, _
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
    
    If Cells(IlkSira, 13).Value = "x" Then
        MsgBox "The Statement 1 data contains missing or incorrect information, so the operation cannot be started.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    
    For i = IlkSira To SonSira
        If Cells(i, 14).Value = "x" Then
            MsgBox "The report data contains missing or incorrect information, so the operation cannot be started.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
            GoTo Son
        End If
    Next i
    
    If Cells(IlkSira, 15).Value = "x" Then
        MsgBox "The Statement 2 data sent to XXXMud contains missing or incorrect information, so the operation cannot be started.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    
    If Cells(IlkSira, 16).Value = "x" Then
        MsgBox "The cover letter data sent to XXXMud contains missing or incorrect information, so the operation cannot be started.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    
    If Cells(IlkSira, 17).Value = "x" Then
        MsgBox "The notification cover letter data sent to the institution contains missing or incorrect information, so the operation cannot be started.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    
    If Cells(IlkSira, 19).Value <> "" And Cells(IlkSira, 19).Value = "x" Then
        MsgBox "The Statement 1 data related to the package received from XXXMud contains missing or incorrect information, so the operation cannot be started.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    
    If Cells(IlkSira, 20).Value <> "" And Cells(IlkSira, 20).Value = "x" Then
        MsgBox "The data related to Statement 2 of the package received from XXXMud contains missing or incorrect information, so the operation cannot be started.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    
    If Cells(IlkSira, 21).Value <> "" And Cells(IlkSira, 21).Value = "x" Then
        MsgBox "The data related to the institution's Statement 2 contains missing or incorrect information, so the operation cannot be started.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If

    
    'Pathfinder...
    AutoPath = ThisWorkbook.Path
    DestOperasyon = AutoPath & "\System Files\Operation\"

    'ÜST YAZI TANIMLARI
    SourceUstYaziNormal = AutoPath & "\System Files\System Templates\Report 2 Cover Letter Templates\Final Cover Letter.docm"
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
    Set IlkSiraBul = Range("CM7:CM100000").Find(What:=Cells(ActiveCell.Row, 5).Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    Set SonSiraBul = Range("CN7:CN100000").Find(What:=Cells(ActiveCell.Row, 5).Value, SearchDirection:=xlNext, _
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
    
    'Statement 1 check
    If Cells(IlkSira, 217).Value = "" Then
        MsgBox "Cover letter for sequence number " & Cells(IlkSira, 5).Value & " cannot be prepared because Statement 1 has not been created. Please create Statement 1 for sequence number " & Cells(IlkSira, 5).Value & " and try again.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    
    'Report(s) check
    For i = IlkSira To SonSira
        If Cells(i, 14).Value <> "" Then
            If Cells(i, 219).Value = "" Then
                MsgBox "Report number " & Cells(i, 11).Value & " has not been created, so the cover letter for sequence number " & Cells(IlkSira, 5).Value & " cannot be prepared. Please create Report number " & Cells(i, 11).Value & " and try again.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
                GoTo Son
            End If
        End If
    Next i
    
    'Statement 2 sent to XXXMud
    If Cells(IlkSira, 220).Value = "" Then
        MsgBox "Cover letter for sequence number " & Cells(IlkSira, 5).Value & " cannot be prepared because Statement 2 sent to XXXMud has not been created. Please create Statement 2 for sequence number " & Cells(IlkSira, 5).Value & " and try again.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    
    'XXXMud outgoing cover letter
    If Cells(IlkSira, 221).Value = "" Then
        MsgBox "Cover letter for sequence number " & Cells(IlkSira, 5).Value & " cannot be prepared because the outgoing cover letter to XXXMud has not been created. Please create the outgoing cover letter for sequence number " & Cells(IlkSira, 5).Value & " and try again.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    
    'Notification cover letter
    If Cells(IlkSira, 222).Value = "" Then
        MsgBox "Cover letter for sequence number " & Cells(IlkSira, 5).Value & " cannot be prepared because the notification cover letter sent to the institution has not been created. Please create the notification cover letter for sequence number " & Cells(IlkSira, 5).Value & " and try again.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    
    'Statement 1 from XXXMud package
    If Cells(IlkSira, 19).Value <> "" And Cells(IlkSira, 223).Value = "" Then
        MsgBox "Cover letter for sequence number " & Cells(IlkSira, 5).Value & " cannot be prepared because Statement 1 related to the package received from XXXMud has not been created. Please create Statement 1 for sequence number " & Cells(IlkSira, 5).Value & " and try again.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    
    'Statement 2 from XXXMud package
    If Cells(IlkSira, 20).Value <> "" And Cells(IlkSira, 224).Value = "" Then
        MsgBox "Cover letter for sequence number " & Cells(IlkSira, 5).Value & " cannot be prepared because Statement 2 related to the package received from XXXMud has not been created. Please create Statement 2 for sequence number " & Cells(IlkSira, 5).Value & " and try again.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    
    'Institution related Statement 2
    If Cells(IlkSira, 21).Value <> "" And Cells(IlkSira, 225).Value = "" Then
        MsgBox "Cover letter for sequence number " & Cells(IlkSira, 5).Value & " cannot be prepared because Statement 2 related to the institution has not been created. Please create Statement 2 for sequence number " & Cells(IlkSira, 5).Value & " and try again.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
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
    ReNameUstYaziNormal = Cells(ActiveCell.Row, 5).Value & "-" & Cells(6, 22).Value
'________________________________________

    If TumDoc = False Then
        'Close the all Word application
        Call ModuleReport2.OpenWordControl
        
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
    'Üst yazı tarihi (sabit)
    objDoc.Tables(1).Cell(Row:=1, Column:=2).Range.Text = Birimx & ", " & FormatEnglishDate(Cells(ActiveCell.Row, 177).Value) '(Format(Cells(ActiveCell.Row, 177).Value, "d mmmm yyyy"))
    'Yazı no (sabit)
    objDoc.Tables(1).Cell(Row:=4, Column:=2).Range.Text = Cells(ActiveCell.Row, 178).Value
    
    

    'Sigorta türü
    If Cells(ActiveCell.Row, 75).Value = "" Then
        Ek2 = Cells(ActiveCell.Row, 201).Value
        Ek2 = Right(Ek2, Len(Ek2) - InStr(Ek2, "/"))
        objDoc.Tables(1).Cell(Row:=7, Column:=2).Range.Text = Ek2
    Else
        Ek2 = Cells(ActiveCell.Row, 75).Value
        Ek2 = Right(Ek2, Len(Ek2) - InStr(Ek2, "/"))
        objDoc.Tables(1).Cell(Row:=7, Column:=2).Range.Text = Ek2
    End If


    'Gönderim usulünde hangi tutanak2 esas alınacak?
    If Cells(ActiveCell.Row, 75).Value <> "" And Cells(ActiveCell.Row, 201).Value = "" Then 'Sadece 75 esas al
        GonderimUsulu = Cells(ActiveCell.Row, 75).Value 'Giden Paket Tipi
    ElseIf Cells(ActiveCell.Row, 75).Value = "" And Cells(ActiveCell.Row, 201).Value <> "" Then 'Sadece 201 esas al
        GonderimUsulu = Cells(ActiveCell.Row, 201).Value 'Giden Paket Tipi
    ElseIf Cells(ActiveCell.Row, 75).Value <> "" And Cells(ActiveCell.Row, 201).Value <> "" Then '75 ve 201, ancak 75 esas al
        GonderimUsulu = Cells(ActiveCell.Row, 75).Value 'Giden Paket Tipi
    End If
    'Ortak bölüm
    GonderimUsulu = Mid(GonderimUsulu, InStr(GonderimUsulu, "/") + 1, Len(GonderimUsulu) - InStr(GonderimUsulu, "/"))
    If GonderimUsulu = "HAND DELIVERY" Then
        GonderimUsulu = "teslim edilmektedir."
    Else
        GonderimUsulu = "gönderilmektedir."
    End If

    'Muhatap
    IlBuyukHarf = UCase(Replace(Replace(Cells(ActiveCell.Row, 203).Value, "i", "I"), "ı", "I"))
    IlKucukHarf = Cells(ActiveCell.Row, 203).Value
    If Cells(ActiveCell.Row, 204).Value <> "" Then
        IlceBuyukHarf = UCase(Replace(Replace(Cells(ActiveCell.Row, 204).Value, "i", "I"), "ı", "I"))
        IlceKucukHarf = Cells(ActiveCell.Row, 204).Value
    Else
        IlceBuyukHarf = ""
        IlceKucukHarf = ""
    End If
    If Cells(ActiveCell.Row, 204).Value <> "" Then
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
'    XXXMudNotu = False
    BStr = 10
    '1. sayfadan sonraki üst bilgi tanımları
    ustbilgitarih = (Format(Cells(ActiveCell.Row, 177).Value, "dd.mm.yyyy"))
    ustbilgisayi = Cells(ActiveCell.Row, 178).Value
    
    '4'lük
    If Cells(ActiveCell.Row, 200).Value <> "" Then
        If Cells(ActiveCell.Row, 199).Value = "Provincial Directorate B" Or Cells(ActiveCell.Row, 199).Value = "Provincial Directorate C" Or _
        Cells(ActiveCell.Row, 199).Value = "Provincial Directorate D" Or Cells(ActiveCell.Row, 199).Value = "Provincial Directorate E" Then 'VALİLİK
            objDoc.Tables(1).Cell(Row:=BStr, Column:=1).Range.Text = IlBuyukHarf & " PROVINCIAL GOVERNORSHIP"
            objDoc.Tables(1).Cell(Row:=BStr + 1, Column:=1).Range.Text = Cells(ActiveCell.Row, 199).Value  'UCase(Replace(Replace(Cells(ActiveCell.Row, 199).Value, "i", "I"), "ı", "I"))
            objDoc.Tables(1).Cell(Row:=BStr + 2, Column:=1).Range.Text = "(" & Cells(ActiveCell.Row, 200).Value & ")"
            objDoc.Tables(1).Cell(Row:=BStr + 3, Column:=1).Range.Text = Bolum3
            objDoc.Tables(1).Cell(Row:=BStr + 3, Column:=1).Range.Font.Underline = wdUnderlineSingle
            M4 = True
            
            ustbilgimuhatap = IlKucukHarf & " Provincial Governorship " & Cells(ActiveCell.Row, 199).Value
        ElseIf Cells(ActiveCell.Row, 199).Value = "District Directorate B" Or Cells(ActiveCell.Row, 199).Value = "District Directorate C" Or _
        Cells(ActiveCell.Row, 199).Value = "District Directorate D" Or Cells(ActiveCell.Row, 199).Value = "District Directorate E" Then 'KAYMAKAMLIK
            objDoc.Tables(1).Cell(Row:=BStr, Column:=1).Range.Text = IlceBuyukHarf & " DISTRICT GOVERNORSHIP"
            objDoc.Tables(1).Cell(Row:=BStr + 1, Column:=1).Range.Text = Cells(ActiveCell.Row, 199).Value 'UCase(Replace(Replace(Cells(ActiveCell.Row, 199).Value, "i", "I"), "ı", "I"))
            objDoc.Tables(1).Cell(Row:=BStr + 2, Column:=1).Range.Text = "(" & Cells(ActiveCell.Row, 200).Value & ")"
            objDoc.Tables(1).Cell(Row:=BStr + 3, Column:=1).Range.Text = Bolum3
            objDoc.Tables(1).Cell(Row:=BStr + 3, Column:=1).Range.Font.Underline = wdUnderlineSingle
            M4 = True
            
            ustbilgimuhatap = IlceKucukHarf & " District Governorship " & Cells(ActiveCell.Row, 199).Value
        ElseIf InStr(Cells(ActiveCell.Row, 199).Value, "General Directorate") <> 0 Or InStr(Cells(ActiveCell.Row, 199).Value, "Regional Directorate") <> 0 Then
            objDoc.Tables(1).Cell(Row:=BStr, Column:=1).Range.Text = UCase(Replace(Replace(ThisWorkbook.Worksheets(2).Cells(6, 111).Value, "i", "I"), "ı", "I"))
            objDoc.Tables(1).Cell(Row:=BStr + 1, Column:=1).Range.Text = Cells(ActiveCell.Row, 199).Value 'UCase(Replace(Replace(Cells(ActiveCell.Row, 199).Value, "i", "I"), "ı", "I"))
            objDoc.Tables(1).Cell(Row:=BStr + 2, Column:=1).Range.Text = "(" & Cells(ActiveCell.Row, 200).Value & ")"
            objDoc.Tables(1).Cell(Row:=BStr + 3, Column:=1).Range.Text = Bolum3
            objDoc.Tables(1).Cell(Row:=BStr + 3, Column:=1).Range.Font.Underline = wdUnderlineSingle
            M4 = True
            
            ustbilgimuhatap = WorksheetFunction.Proper(ThisWorkbook.Worksheets(2).Cells(6, 111).Value) & " " & Cells(ActiveCell.Row, 199).Value
        Else 'YARGI 3'lük
            Bolum1 = UCase(Replace(Replace(Cells(ActiveCell.Row, 199).Value, "i", "I"), "ı", "I"))
            If InStr(Bolum1, "X.X. ") > 0 Then
                objDoc.Tables(1).Cell(Row:=BStr, Column:=1).Range.Text = Mid(Bolum1, 6, Len(Bolum1))
                objDoc.Tables(1).Cell(Row:=BStr + 1, Column:=1).Range.Text = "(" & Cells(ActiveCell.Row, 200).Value & ")"
                objDoc.Tables(1).Cell(Row:=BStr + 2, Column:=1).Range.Text = Bolum3
                objDoc.Tables(1).Cell(Row:=BStr + 2, Column:=1).Range.Font.Underline = wdUnderlineSingle
                M3 = True
                
                ustbilgimuhatap = Mid(Cells(ActiveCell.Row, 199).Value, 6, Len(Cells(ActiveCell.Row, 199).Value))
            Else
                objDoc.Tables(1).Cell(Row:=BStr, Column:=1).Range.Text = Bolum1
                objDoc.Tables(1).Cell(Row:=BStr + 1, Column:=1).Range.Text = "(" & Cells(ActiveCell.Row, 200).Value & ")"
                objDoc.Tables(1).Cell(Row:=BStr + 2, Column:=1).Range.Text = Bolum3
                objDoc.Tables(1).Cell(Row:=BStr + 2, Column:=1).Range.Font.Underline = wdUnderlineSingle
                M3 = True

                ustbilgimuhatap = Cells(ActiveCell.Row, 199).Value
            End If
        End If
    End If
    '3'lük
    If Cells(ActiveCell.Row, 200).Value = "" Then
        If Cells(ActiveCell.Row, 199).Value = "Provincial Directorate B" Or Cells(ActiveCell.Row, 199).Value = "Provincial Directorate C" Or _
        Cells(ActiveCell.Row, 199).Value = "Provincial Directorate D" Or Cells(ActiveCell.Row, 199).Value = "Provincial Directorate E" Then 'VALİLİK
            objDoc.Tables(1).Cell(Row:=BStr, Column:=1).Range.Text = IlBuyukHarf & " PROVINCIAL GOVERNORSHIP"
            objDoc.Tables(1).Cell(Row:=BStr + 1, Column:=1).Range.Text = "(" & Cells(ActiveCell.Row, 199).Value & ")" 'UCase(Replace(Replace(Cells(ActiveCell.Row, 199).Value, "i", "I"), "ı", "I"))
            objDoc.Tables(1).Cell(Row:=BStr + 2, Column:=1).Range.Text = Bolum3
            objDoc.Tables(1).Cell(Row:=BStr + 2, Column:=1).Range.Font.Underline = wdUnderlineSingle
            M3 = True
            
            ustbilgimuhatap = IlKucukHarf & " Provincial Governorship " & Cells(ActiveCell.Row, 199).Value
        ElseIf Cells(ActiveCell.Row, 199).Value = "District Directorate B" Or Cells(ActiveCell.Row, 199).Value = "District Directorate C" Or _
        Cells(ActiveCell.Row, 199).Value = "District Directorate D" Or Cells(ActiveCell.Row, 199).Value = "District Directorate E" Then
            objDoc.Tables(1).Cell(Row:=BStr, Column:=1).Range.Text = IlceBuyukHarf & " DISTRICT GOVERNORSHIP"
            objDoc.Tables(1).Cell(Row:=BStr + 1, Column:=1).Range.Text = "(" & Cells(ActiveCell.Row, 199).Value & ")" 'UCase(Replace(Replace(Cells(ActiveCell.Row, 199).Value, "i", "I"), "ı", "I"))
            objDoc.Tables(1).Cell(Row:=BStr + 2, Column:=1).Range.Text = Bolum3
            objDoc.Tables(1).Cell(Row:=BStr + 2, Column:=1).Range.Font.Underline = wdUnderlineSingle
            M3 = True
            
            ustbilgimuhatap = IlceKucukHarf & " District Governorship " & Cells(ActiveCell.Row, 199).Value
        ElseIf InStr(Cells(ActiveCell.Row, 199).Value, "General Directorate") <> 0 Or InStr(Cells(ActiveCell.Row, 199).Value, "Regional Directorate") <> 0 Then
            objDoc.Tables(1).Cell(Row:=BStr, Column:=1).Range.Text = UCase(Replace(Replace(ThisWorkbook.Worksheets(2).Cells(6, 111).Value, "i", "I"), "ı", "I"))
            objDoc.Tables(1).Cell(Row:=BStr + 1, Column:=1).Range.Text = "(" & Cells(ActiveCell.Row, 199).Value & ")" 'UCase(Replace(Replace(Cells(ActiveCell.Row, 199).Value, "i", "I"), "ı", "I"))
            objDoc.Tables(1).Cell(Row:=BStr + 2, Column:=1).Range.Text = Bolum3
            objDoc.Tables(1).Cell(Row:=BStr + 2, Column:=1).Range.Font.Underline = wdUnderlineSingle
            M3 = True
            
            ustbilgimuhatap = WorksheetFunction.Proper(ThisWorkbook.Worksheets(2).Cells(6, 111).Value) & " " & Cells(ActiveCell.Row, 199).Value
        Else 'YARGI 2'lik
            Bolum1 = UCase(Replace(Replace(Cells(ActiveCell.Row, 199).Value, "i", "I"), "ı", "I"))
            If InStr(Bolum1, "X.X. ") > 0 Then
                objDoc.Tables(1).Cell(Row:=BStr, Column:=1).Range.Text = Mid(Bolum1, 6, Len(Bolum1))
                objDoc.Tables(1).Cell(Row:=BStr + 1, Column:=1).Range.Text = Bolum3
                objDoc.Tables(1).Cell(Row:=BStr + 1, Column:=1).Range.Font.Underline = wdUnderlineSingle
                M2 = True

                ustbilgimuhatap = Mid(Cells(ActiveCell.Row, 199).Value, 6, Len(Cells(ActiveCell.Row, 199).Value))
            Else
                objDoc.Tables(1).Cell(Row:=BStr, Column:=1).Range.Text = Bolum1
                objDoc.Tables(1).Cell(Row:=BStr + 1, Column:=1).Range.Text = Bolum3
                objDoc.Tables(1).Cell(Row:=BStr + 1, Column:=1).Range.Font.Underline = wdUnderlineSingle
                M2 = True

                ustbilgimuhatap = Cells(ActiveCell.Row, 199).Value
            End If
        End If
    End If

    'İlgi tema
    If Cells(IlkSira, 33).Value = "Provincial Directorate B" Then
        If Cells(IlkSira, 34).Value <> "" Then
            GelenTema = Cells(IlkSira, 25).Value & " Provincial Governorship Provincial Directorate B " & Cells(IlkSira, 34).Value
        Else
            GelenTema = Cells(IlkSira, 25).Value & " Provincial Governorship Provincial Directorate B"
        End If
    ElseIf Cells(IlkSira, 33).Value = "District Directorate B" Then
        If Cells(IlkSira, 34).Value <> "" Then
            GelenTema = Cells(IlkSira, 26).Value & " District Governorship District Directorate B " & Cells(IlkSira, 34).Value
        Else
            GelenTema = Cells(IlkSira, 26).Value & " District Governorship District Directorate B"
        End If
    ElseIf Cells(IlkSira, 33).Value = "Provincial Directorate C" Then
        If Cells(IlkSira, 34).Value <> "" Then
            GelenTema = Cells(IlkSira, 25).Value & " Provincial Governorship Provincial Directorate C " & Cells(IlkSira, 34).Value
        Else
            GelenTema = Cells(IlkSira, 25).Value & " Provincial Governorship Provincial Directorate C"
        End If
    ElseIf Cells(IlkSira, 33).Value = "District Directorate C" Then
        If Cells(IlkSira, 34).Value <> "" Then
            GelenTema = Cells(IlkSira, 26).Value & " District Governorship District Directorate C " & Cells(IlkSira, 34).Value
        Else
            GelenTema = Cells(IlkSira, 26).Value & " District Governorship District Directorate C"
        End If
    ElseIf Cells(IlkSira, 33).Value = "Provincial Directorate D" Then
        If Cells(IlkSira, 34).Value <> "" Then
            GelenTema = Cells(IlkSira, 25).Value & " Provincial Governorship Provincial Directorate D " & Cells(IlkSira, 34).Value
        Else
            GelenTema = Cells(IlkSira, 25).Value & " Provincial Governorship Provincial Directorate D"
        End If
    ElseIf Cells(IlkSira, 33).Value = "District Directorate D" Then
        If Cells(IlkSira, 34).Value <> "" Then
            GelenTema = Cells(IlkSira, 26).Value & " District Governorship District Directorate D " & Cells(IlkSira, 34).Value
        Else
            GelenTema = Cells(IlkSira, 26).Value & " District Governorship District Directorate D"
        End If
    ElseIf Cells(IlkSira, 33).Value = "Provincial Directorate E" Then
        If Cells(IlkSira, 34).Value <> "" Then
            GelenTema = Cells(IlkSira, 25).Value & " Provincial Governorship Provincial Directorate E " & Cells(IlkSira, 34).Value
        Else
            GelenTema = Cells(IlkSira, 25).Value & " Provincial Governorship Provincial Directorate E"
        End If
    ElseIf Cells(IlkSira, 33).Value = "District Directorate E" Then
        If Cells(IlkSira, 34).Value <> "" Then
            GelenTema = Cells(IlkSira, 26).Value & " District Governorship District Directorate E " & Cells(IlkSira, 34).Value
        Else
            GelenTema = Cells(IlkSira, 26).Value & " District Governorship District Directorate E"
        End If
    ElseIf InStr(Cells(IlkSira, 33).Value, "General Directorate") <> 0 Or InStr(Cells(IlkSira, 33).Value, "Regional Directorate") <> 0 Then
        If Cells(IlkSira, 34).Value <> "" Then
            GelenTema = WorksheetFunction.Proper(ThisWorkbook.Worksheets(2).Cells(6, 111).Value) & " " & Cells(IlkSira, 33).Value & " " & Cells(IlkSira, 34).Value
        Else
            GelenTema = WorksheetFunction.Proper(ThisWorkbook.Worksheets(2).Cells(6, 111).Value) & " " & Cells(IlkSira, 33).Value
        End If
    Else
        If InStr(Cells(IlkSira, 33).Value, "X.X. ") > 0 Then
            If Cells(IlkSira, 34).Value <> "" Then
                GelenTema = Mid(Cells(IlkSira, 33).Value, 6, Len(Cells(IlkSira, 33).Value)) & " " & Cells(IlkSira, 34).Value
            Else
                GelenTema = Mid(Cells(IlkSira, 33).Value, 6, Len(Cells(IlkSira, 33).Value))
            End If
        Else
            If Cells(IlkSira, 34).Value <> "" Then
                GelenTema = Cells(IlkSira, 33).Value & " " & Cells(IlkSira, 34).Value
            Else
                GelenTema = Cells(IlkSira, 33).Value
            End If
        End If
    End If


    'İlgi
    Bolum1 = "Delivered to us on "
    Bolum2 = ", with a letter dated "
    Bolum3 = " and reference number "
    Bolum4 = "."

    If Cells(ActiveCell.Row, 33).Value = Cells(ActiveCell.Row, 199).Value And _
    Cells(ActiveCell.Row, 34).Value = Cells(ActiveCell.Row, 200).Value And _
    Cells(ActiveCell.Row, 25).Value = Cells(ActiveCell.Row, 203).Value Then
        ' Gelen ve giden birim aynıysa
        Bolum5 = "a) " & Bolum1 & Cells(ActiveCell.Row, 36).Value & Bolum2 & _
            Cells(ActiveCell.Row, 28).Value & Bolum3 & _
            Cells(ActiveCell.Row, 29).Value & Bolum4
    Else
        ' Gelen ve giden birim farklıysa
        Bolum5 = "a) " & Bolum1 & Cells(ActiveCell.Row, 36).Value & _
                    " from " & GelenTema & Bolum2 & _
                    Cells(IlkSira, 28).Value & Bolum3 & _
                    Cells(IlkSira, 29).Value & Bolum4
    End If
    Bolum6 = "b) Our letter dated " & Cells(ActiveCell.Row, 83).Value & " and numbered " & Cells(ActiveCell.Row, 84).Value & "."
    
    'ilgileri ekle
    objDoc.Tables(1).Cell(Row:=15, Column:=3).Range.Text = Bolum5 & vbNewLine & Bolum6
    
    'Gövde metni (Rapor2_2 değişkeni)
    AdetTakip = 0
    NomTakip = 0
    TipATakip = ""
    AdetTopla = 0
    y = 0
    For i = IlkSira To SonSira
        If Cells(i, 62).Value = "valid" Or Cells(i, 62).Value = "invalid" Then 'And Left(Cells(i, 63).Value, 11) = "Technique A" Then
            y = y + 1
            If y = 1 Then
                AdetTakip = Cells(i, 52).Value
                TipATakip = Cells(i, 46).Value
                TipATakip = Right(TipATakip, Len(TipATakip) - InStr(TipATakip, "-")) 'Bu karakterin solunda bir boşluk oluşuyor.
                NomTakip = Cells(i, 49).Value
                If AdetTakip > 1 Then
                    Ek1 = AdetTakip & " pieces of " & NomTakip & TipATakip
                Else
                    Ek1 = AdetTakip & " piece of " & NomTakip & TipATakip
                End If
                '2 pieces of 100 X1
                AdetTopla = AdetTopla + Cells(i, 52).Value
            ElseIf y > 1 Then
                AdetTakip = Cells(i, 52).Value
                TipATakip = Cells(i, 46).Value
                TipATakip = Right(TipATakip, Len(TipATakip) - InStr(TipATakip, "-"))
                NomTakip = Cells(i, 49).Value
                If AdetTakip > 1 Then
                    Ek1 = Ek1 & ", " & AdetTakip & " pieces of " & NomTakip & TipATakip
                Else
                    Ek1 = Ek1 & ", " & AdetTakip & " piece of " & NomTakip & TipATakip
                End If
                '2 pieces of 100 X1, 3 pieces of 200 X1, 5 pieces of 50 X3
                AdetTopla = AdetTopla + Cells(i, 52).Value
            End If
        End If
    Next i

    'Rapor2_2 rapor adedi için virgüller sayılacak (Rapor2_2x ile temsil edildi.) (Rapor2_2 değişkeni)
    Set rng = Cells(ActiveCell.Row, 180)
    Rapor2_2x = 1
    For i = 1 To Len(rng)
        If Mid(rng, i, 1) = "," Then
            Rapor2_2x = Rapor2_2x + 1
            'MsgBox Mid(rng, i, 1)
        End If
    Next i

    If Rapor2_2x = 1 Then
        Ek4 = "report"
        'Ek5 = "Report 2.1"
        Ek6 = "Report 2.2"
        Bolum5 = "is"
    Else
        Ek4 = "reports"
        'Ek5 = "Report 2.1s"
        Ek6 = "Report 2.2s"
        Bolum5 = "are"
    End If

    Ek3 = Cells(ActiveCell.Row, 180).Value 'report no.

    If AdetTopla = 1 Then
        Ek2 = "item"
        Bolum1 = "The Type A " & Ek2 & " (" & Ek1 & ")" & ", sent to our unit as an enclosure to the referenced letter, has been documented in "
        'Ek7 = "has"
    Else
        Ek2 = "items"
        Bolum1 = "The Type A " & Ek2 & " (" & Ek1 & ")" & ", sent to our unit as an enclosure to the referenced letter, have been documented in "
        'Ek7 = "have"
    End If

    'Rapor2_2 rapor senaryoları
    If Cells(IlkSira, 186).Value = "Yes" Then

        ' Construct the text
        Bolum1 = Bolum1 & Ek6 & ", "
        Bolum2 = Bolum1 & "dated " & Cells(ActiveCell.Row, 179).Value & " and numbered " & Ek3 & ". " & "Both the " & Ek2 & " and the corresponding " & Ek4 & " have been forwarded to your office accordingly."
        Bolum3 = "In the event that the " & Ek6 & " " & Bolum5 & " communicated by your office to the General Directorate of Organization B – Xxxxx Xxxxxxx, " & _
        "it may be possible to obtain the xxxxxxxxxx xxxxxx xxxxxx xxxxxxx / xxxxxxx xxxxxxxxxx xxxxxxxx / xxxxxxxx xxxxxxxx information of the invalid Type A " & _
        Ek2 & " by contacting Xxxxxxx through the mentioned unit."
        Bolum4 = "Respectfully submitted for your information."
        objDoc.Tables(2).Cell(Row:=1, Column:=1).Range.Text = Chr(9) & Bolum2
        objDoc.Tables(2).Cell(Row:=2, Column:=1).Range.Text = Chr(9) & Bolum3
        objDoc.Tables(2).Cell(Row:=3, Column:=1).Range.Text = Chr(9) & Bolum4
    
        Set MyRange = objDoc.Tables(2).Cell(Row:=1, Column:=1).Range
        With MyRange.Find
            .Execute FindText:="dated " & Cells(ActiveCell.Row, 179).Value & " and numbered " & Ek3
            .Execute Forward:=True
        End With
        MyRange.Font.Bold = True
    ElseIf Cells(IlkSira, 186).Value = "No" Then

        ' Construct the text
        Bolum1 = Bolum1 & Ek6 & ", "
        Bolum2 = Bolum1 & "dated " & Cells(ActiveCell.Row, 179).Value & " and numbered " & Ek3 & ". " & "Both the " & Ek2 & " and the corresponding " & Ek4 & " have been forwarded to your office accordingly."
        Bolum3 = "As a result of the examination of the " & Ek2 & ", conducted through xxxxxxx and documented in " & Ek6 & ", it was concluded that no Report 2.2 code could be assigned to the " & Ek2 & "."
        Bolum4 = "Respectfully submitted for your information."
        objDoc.Tables(2).Cell(Row:=1, Column:=1).Range.Text = Chr(9) & Bolum2
        objDoc.Tables(2).Cell(Row:=2, Column:=1).Range.Text = Chr(9) & Bolum3
        objDoc.Tables(2).Cell(Row:=3, Column:=1).Range.Text = Chr(9) & Bolum4
        
        Set MyRange = objDoc.Tables(2).Cell(Row:=1, Column:=1).Range
        With MyRange.Find
            .Execute FindText:="dated " & Cells(ActiveCell.Row, 179).Value & " and numbered " & Ek3
            .Execute Forward:=True
        End With
        MyRange.Font.Bold = True
    ElseIf Cells(IlkSira, 186).Value = "Unresolved" Then
        ' Construct the text
        Bolum1 = Bolum1 & Ek6 & ", "
        Bolum2 = Bolum1 & "dated " & Cells(ActiveCell.Row, 179).Value & " and numbered " & Ek3 & ". " & "Both the " & Ek2 & " and the corresponding " & Ek4 & " have been forwarded to your office accordingly."
        Bolum3 = "Although the xxxxxx xxxxxxx xxxxxx / xxxxxxx xxxxxxxx xxx xxxxxxxxxx of the " & Ek2 & " were identified during the examination conducted through xxxxxxx, the Report 2.2 code could not be determined."
        Bolum4 = "Respectfully submitted for your information."
        objDoc.Tables(2).Cell(Row:=1, Column:=1).Range.Text = Chr(9) & Bolum2
        objDoc.Tables(2).Cell(Row:=2, Column:=1).Range.Text = Chr(9) & Bolum3
        objDoc.Tables(2).Cell(Row:=3, Column:=1).Range.Text = Chr(9) & Bolum4
    
        Set MyRange = objDoc.Tables(2).Cell(Row:=1, Column:=1).Range
        With MyRange.Find
            .Execute FindText:="dated " & Cells(ActiveCell.Row, 179).Value & " and numbered " & Ek3
            .Execute Forward:=True
        End With
        MyRange.Font.Bold = True
    End If
    
    'imzalar (sabit)
    objDoc.Tables(3).Cell(Row:=4, Column:=2).Range.Text = Cells(ActiveCell.Row, 142).Value 'Ad Soyad1
    objDoc.Tables(3).Cell(Row:=5, Column:=2).Range.Text = Cells(ActiveCell.Row, 143).Value 'Unvan1
    objDoc.Tables(3).Cell(Row:=4, Column:=3).Range.Text = Cells(ActiveCell.Row, 145).Value 'Ad Soyad2
    objDoc.Tables(3).Cell(Row:=5, Column:=3).Range.Text = Cells(ActiveCell.Row, 146).Value 'Unvan2

    'Ekler (Rapor2_2 değişken)
    objDoc.Tables(3).Cell(Row:=8, Column:=1).Range.Text = "Attachment:"
    If Cells(IlkSira, 187).Value = "All" And Cells(IlkSira, 188).Value = "No" Then 'XXXMud tutanak2 bilgileri var.
        
        Ek2 = Cells(ActiveCell.Row, 201).Value 'Kapalı Package A
        Ek2 = Left(Ek2, InStr(Ek2, "/") - 1)
        If Cells(ActiveCell.Row, 202).Value > 1 Then objDoc.Tables(3).Cell(Row:=8, Column:=2).Range.Text = " 1) Enclosed " & Ek2 & " (" & Cells(ActiveCell.Row, 202).Value & " pieces)"
        If Cells(ActiveCell.Row, 202).Value < 2 Then objDoc.Tables(3).Cell(Row:=8, Column:=2).Range.Text = " 1) Enclosed " & Ek2 & " (" & Cells(ActiveCell.Row, 202).Value & " piece)"

        x = Application.Sum(Range(Cells(IlkSira, 205), Cells(SonSira, 205))) 'Statement 2 toplam sayfa sayısı
        If x > 1 Then objDoc.Tables(3).Cell(Row:=9, Column:=2).Range.Text = " 2) Statement 2 (" & x & " pages)"
        If x < 2 Then objDoc.Tables(3).Cell(Row:=9, Column:=2).Range.Text = " 2) Statement 2 (" & x & " page)"
        
    ElseIf Cells(IlkSira, 187).Value = "All" And Cells(IlkSira, 188).Value = "Yes" Then 'İlgili birim tutanak2sı var.
        
        Ek2 = Cells(ActiveCell.Row, 75).Value 'Kapalı Package A
        Ek2 = Left(Ek2, InStr(Ek2, "/") - 1)
        If Cells(ActiveCell.Row, 76).Value > 1 Then objDoc.Tables(3).Cell(Row:=8, Column:=2).Range.Text = " 1) Enclosed " & Ek2 & " (" & Cells(ActiveCell.Row, 76).Value & " pieces)"
        If Cells(ActiveCell.Row, 76).Value < 2 Then objDoc.Tables(3).Cell(Row:=8, Column:=2).Range.Text = " 1) Enclosed " & Ek2 & " (" & Cells(ActiveCell.Row, 76).Value & " piece)"

        x = Application.Sum(Range(Cells(IlkSira, 225), Cells(SonSira, 225))) 'Statement 2 toplam sayfa sayısı
        If x > 1 Then objDoc.Tables(3).Cell(Row:=9, Column:=2).Range.Text = " 2) Statement 2 (" & x & " pages)"
        If x < 2 Then objDoc.Tables(3).Cell(Row:=9, Column:=2).Range.Text = " 2) Statement 2 (" & x & " page)"

    ElseIf Cells(IlkSira, 187).Value = "Technique A" And Cells(IlkSira, 188).Value = "No" Then 'İlgili birim tutanak2sı var.

        Ek2 = Cells(ActiveCell.Row, 75).Value 'Kapalı Package A
        Ek2 = Left(Ek2, InStr(Ek2, "/") - 1)
        objDoc.Tables(3).Cell(Row:=8, Column:=2).Range.Text = " 1) Enclosed " & Ek2 & " (" & Cells(ActiveCell.Row, 76).Value + Cells(ActiveCell.Row, 202).Value & " pieces)"

        x = Application.Sum(Range(Cells(IlkSira, 225), Cells(SonSira, 225))) + Application.Sum(Range(Cells(IlkSira, 205), Cells(SonSira, 205))) 'Statement 2 toplam sayfa sayısı
        objDoc.Tables(3).Cell(Row:=9, Column:=2).Range.Text = " 2) Statement 2 (2 statements, " & "total of " & x & " pages)"
        
    ElseIf Cells(IlkSira, 187).Value = "Technique A" And Cells(IlkSira, 188).Value = "Yes" And Cells(IlkSira, 189).Value = "Yes" Then 'İlgili birim tutanak2sı var.
        
        Ek2 = Cells(ActiveCell.Row, 75).Value 'Kapalı Package A
        Ek2 = Left(Ek2, InStr(Ek2, "/") - 1)
        If Cells(ActiveCell.Row, 76).Value > 1 Then objDoc.Tables(3).Cell(Row:=8, Column:=2).Range.Text = " 1) Enclosed " & Ek2 & " (" & Cells(ActiveCell.Row, 76).Value & " pieces)"
        If Cells(ActiveCell.Row, 76).Value < 2 Then objDoc.Tables(3).Cell(Row:=8, Column:=2).Range.Text = " 1) Enclosed " & Ek2 & " (" & Cells(ActiveCell.Row, 76).Value & " piece)"

        x = Application.Sum(Range(Cells(IlkSira, 225), Cells(SonSira, 225))) 'Statement 2 toplam sayfa sayısı
        If x > 1 Then objDoc.Tables(3).Cell(Row:=9, Column:=2).Range.Text = " 2) Statement 2 (" & x & " pages)"
        If x < 2 Then objDoc.Tables(3).Cell(Row:=9, Column:=2).Range.Text = " 2) Statement 2 (" & x & " page)"
        
    ElseIf Cells(IlkSira, 187).Value = "Technique A" And Cells(IlkSira, 188).Value = "Yes" And Cells(IlkSira, 189).Value = "No" Then 'XXXMud ve ilgili birim tutanak2ları var.

        Ek2 = Cells(ActiveCell.Row, 75).Value 'Kapalı Package A
        Ek2 = Left(Ek2, InStr(Ek2, "/") - 1)
        objDoc.Tables(3).Cell(Row:=8, Column:=2).Range.Text = " 1) Enclosed " & Ek2 & " (" & Cells(ActiveCell.Row, 76).Value + Cells(ActiveCell.Row, 202).Value & " pieces)"

        x = Application.Sum(Range(Cells(IlkSira, 225), Cells(SonSira, 225))) + Application.Sum(Range(Cells(IlkSira, 224), Cells(SonSira, 224))) 'Statement 2 toplam sayfa sayısı
        objDoc.Tables(3).Cell(Row:=9, Column:=2).Range.Text = " 2) Statement 2 (2 statements, " & "total of " & x & " pages)"
        
    End If

    'Rapor2_2 rapor senaryoları
    If Cells(IlkSira, 186).Value = "Yes" Then
    
        x = Application.Sum(Range(Cells(IlkSira, 173), Cells(SonSira, 173)))  'Rapor 2.2 toplam sayfa sayısı
        If Rapor2_2x = 1 Then
            If x > 1 Then objDoc.Tables(3).Cell(Row:=10, Column:=2).Range.Text = " 3) Report 2.2 (" & x & " pages)"
            If x < 2 Then objDoc.Tables(3).Cell(Row:=10, Column:=2).Range.Text = " 3) Report 2.2 (" & x & " page)"
        Else
            objDoc.Tables(3).Cell(Row:=10, Column:=2).Range.Text = " 3) Report 2.2 (" & y & " reports, " & "total of " & x & " pages)"
        End If
        
        'Analiz Dijital İçerik'si
        If Cells(ActiveCell.Row, 181).Value > 1 Then objDoc.Tables(3).Cell(Row:=11, Column:=2).Range.Text = " 4) Digital Content (" & Cells(ActiveCell.Row, 181).Value & " pieces)"
        If Cells(ActiveCell.Row, 181).Value < 2 Then objDoc.Tables(3).Cell(Row:=11, Column:=2).Range.Text = " 4) Digital Content (" & Cells(ActiveCell.Row, 181).Value & " piece)"
        
        'Rapor2_2 Analiz Çıktısı
        If Cells(ActiveCell.Row, 181).Value > 1 Then objDoc.Tables(3).Cell(Row:=12, Column:=2).Range.Text = " 5) Report 2.2 Analysis Output (" & Cells(ActiveCell.Row, 182).Value & " pages)"
        If Cells(ActiveCell.Row, 181).Value < 2 Then objDoc.Tables(3).Cell(Row:=12, Column:=2).Range.Text = " 5) Report 2.2 Analysis Output (" & Cells(ActiveCell.Row, 182).Value & " page)"
        
        'Statement 1 (döküm dahil) toplam sayfa sayısı
        If Cells(ActiveCell.Row, 218).Value = "" Then
            x = Application.Sum(Range(Cells(IlkSira, 217), Cells(SonSira, 217)))
            If x > 1 Then objDoc.Tables(3).Cell(Row:=13, Column:=2).Range.Text = " 6) Statement 1 (" & x & " pages)"
            If x < 2 Then objDoc.Tables(3).Cell(Row:=13, Column:=2).Range.Text = " 6) Statement 1 (" & x & " page)"
        Else
            x = Application.Sum(Range(Cells(IlkSira, 217), Cells(SonSira, 218)))
            objDoc.Tables(3).Cell(Row:=13, Column:=2).Range.Text = " 6) Attached Statement 1 (total of " & x & " pages)"
        End If

    ElseIf Cells(IlkSira, 186).Value = "No" Or Cells(IlkSira, 186).Value = "Unresolved" Then

        x = Application.Sum(Range(Cells(IlkSira, 173), Cells(SonSira, 173)))  'Rapor 2.2 toplam sayfa sayısı
        If Rapor2_2x = 1 Then
            If x > 1 Then objDoc.Tables(3).Cell(Row:=10, Column:=2).Range.Text = " 3) Report 2.2 (" & x & " pages)"
            If x < 2 Then objDoc.Tables(3).Cell(Row:=10, Column:=2).Range.Text = " 3) Report 2.2 (" & x & " page)"
        Else
            objDoc.Tables(3).Cell(Row:=10, Column:=2).Range.Text = " 3) Report 2.2 (" & y & " reports, " & "total of " & x & " pages)"
        End If
                
        'Statement 1 (döküm dahil) toplam sayfa sayısı
        If Cells(ActiveCell.Row, 218).Value = "" Then
            x = Application.Sum(Range(Cells(IlkSira, 217), Cells(SonSira, 217)))
            If x > 1 Then objDoc.Tables(3).Cell(Row:=11, Column:=2).Range.Text = " 4) Statement 1 (" & x & " pages)"
            If x < 2 Then objDoc.Tables(3).Cell(Row:=11, Column:=2).Range.Text = " 4) Statement 1 (" & x & " page)"
        Else
            x = Application.Sum(Range(Cells(IlkSira, 217), Cells(SonSira, 218)))
            objDoc.Tables(3).Cell(Row:=11, Column:=2).Range.Text = " 4) Attached Statement 1 (total of " & x & " pages)"
        End If
        
    End If
   
 '__________________________DİNAMİK SAYFA YAPISI
    
    'İlgi ve gövde metninin kaç satırdan oluştuğu bilgisinin elde edilmesinde kullanılıyor.
    Set IlgiRng = objDoc.Tables(1).Cell(Row:=15, Column:=3).Range
    Set Govde1Rng = objDoc.Tables(2).Cell(Row:=1, Column:=1).Range
    Set Govde2Rng = objDoc.Tables(2).Cell(Row:=2, Column:=1).Range

    IlgiStrSay = Govde1Rng.Information(wdFirstCharacterLineNumber) - IlgiRng.Information(wdFirstCharacterLineNumber) - 1
    Govde1StrSay = Govde2Rng.Information(wdFirstCharacterLineNumber) - Govde1Rng.Information(wdFirstCharacterLineNumber)
    'MsgBox "İlgi Satır Sayısı: " & IlgiStrSay & vbNewLine & vbNewLine & "Gövde 1 Satır Sayısı: " & Govde1StrSay
    IlgiFarkSay = IlgiStrSay - 2
    Govde1FarkSay = Govde1StrSay - 3
    CokluSayfa = 0
    'MsgBox IlgiFarkSay + Govde1FarkSay 'Varsayılan ilgi ve 1 paragrafta (2 rowda) toplam 5 satıra göre sıfırlandı.
    
    
    'Dinamik sayfa düzeni
    For i = 14 To 15 'Ek sonrası satırları sil
        objDoc.Tables(3).Rows(14).Delete
    Next i
    
    If Cells(IlkSira, 186).Value = "Yes" Then '6 ek var
        Govde1FarkSay = Govde1FarkSay + 9
    Else '4 ek var
        For i = 1 To 2 'Ek sonrası satırları sil
            objDoc.Tables(3).Rows(12).Delete
        Next i
        Govde1FarkSay = Govde1FarkSay + 6
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
    Cells(ActiveCell.Row, 226).Value = TotalSayfaUstYazi

    objWord.Visible = True
    objWord.Activate
    
    'Nesneleri temizle
    Set fso = Nothing
    Set objWord = Nothing
    Set objDoc = Nothing

Son:

If IlceSakla <> "" Then
    Cells(ActiveCell.Row, 204).Value = IlceSakla
End If

'Worksheets(4).Activate
Application.ScreenUpdating = True
Application.DisplayAlerts = True
'Application.EnableEvents = True

'End If

End Sub

Sub Rapor2_1UstYazi()

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
Dim Birimx As String, UserName As String
Dim gecersizSay As Long
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
Dim StrTemaNotu As String


IlceSakla = ""
If InStr(Cells(ActiveCell.Row, 78).Value, " Organization A") <> 0 Then
    IlceSakla = Cells(ActiveCell.Row, 78).Value
    Cells(ActiveCell.Row, 78).Value = ""
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
    Set IlkSiraBul = Range("CM7:CM100000").Find(What:=Cells(j, 5).Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    Set SonSiraBul = Range("CN7:CN100000").Find(What:=Cells(j, 5).Value, SearchDirection:=xlNext, _
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
        MsgBox "The Statement 1 data contains missing or incorrect information, so the operation cannot be started.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
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
    SourceUstYaziNormal = AutoPath & "\System Files\System Templates\Report 2 Cover Letter Templates\Report 2 Cover Letter.docm"
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
    Set IlkSiraBul = Range("CM7:CM100000").Find(What:=Cells(ActiveCell.Row, 5).Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    Set SonSiraBul = Range("CN7:CN100000").Find(What:=Cells(ActiveCell.Row, 5).Value, SearchDirection:=xlNext, _
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
    
    'Statement 1 check
    If Cells(IlkSira, 97).Value = "" Then
        MsgBox "Cover letter for sequence number " & Cells(IlkSira, 5).Value & " cannot be prepared because Statement 1 has not been created. Please create Statement 1 for sequence number " & Cells(IlkSira, 5).Value & " and try again.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    
    'Report(s) check
    For i = IlkSira To SonSira
        If Cells(i, 7).Value <> "" Then
            If Cells(i, 99).Value = "" Then
                MsgBox "Report number " & Cells(i, 11).Value & " has not been created, so the cover letter for sequence number " & Cells(IlkSira, 5).Value & " cannot be prepared. Please create Report number " & Cells(i, 11).Value & " and try again.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
                GoTo Son
            End If
        End If
    Next i
    
    'Statement 2 check
    If Cells(IlkSira, 100).Value = "" Then
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
    ReNameUstYaziNormal = Cells(ActiveCell.Row, 5).Value & "-" & Cells(6, 9).Value
'________________________________________

    If TumDoc = False Then
        'Close the all Word application
        Call ModuleReport2.OpenWordControl
        
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
    Ek2 = Cells(ActiveCell.Row, 75).Value
    Ek2 = Right(Ek2, Len(Ek2) - InStr(Ek2, "/"))
    objDoc.Tables(1).Cell(Row:=7, Column:=2).Range.Text = Ek2


    'Muhatap
    IlBuyukHarf = UCase(Replace(Replace(Cells(ActiveCell.Row, 77).Value, "i", "I"), "ı", "I"))
    IlKucukHarf = Cells(ActiveCell.Row, 77).Value
    If Cells(ActiveCell.Row, 78).Value <> "" Then
        IlceBuyukHarf = UCase(Replace(Replace(Cells(ActiveCell.Row, 78).Value, "i", "I"), "ı", "I"))
        IlceKucukHarf = Cells(ActiveCell.Row, 78).Value
    Else
        IlceBuyukHarf = ""
        IlceKucukHarf = ""
    End If
    If Cells(ActiveCell.Row, 78).Value <> "" Then
        Bolum3 = IlceKucukHarf & "/" & IlBuyukHarf
    Else
        Bolum3 = IlBuyukHarf
    End If
    

    
    'YENİ MUHATAP TEMASI
    
    M2 = False
    M3 = False
    M4 = False
    Ifv = False
    Ify = False
    XXXMudNotu = False
    BStr = 10
    '1. sayfadan sonraki üst bilgi tanımları
    ustbilgitarih = (Format(Cells(ActiveCell.Row, 83).Value, "dd.mm.yyyy"))
    ustbilgisayi = Cells(ActiveCell.Row, 84).Value
    
    '4'lük
    If Cells(ActiveCell.Row, 73).Value <> "" Then
        If Cells(ActiveCell.Row, 72).Value = "Provincial Directorate B" Or Cells(ActiveCell.Row, 72).Value = "Provincial Directorate C" Or _
        Cells(ActiveCell.Row, 72).Value = "Provincial Directorate D" Or Cells(ActiveCell.Row, 72).Value = "Provincial Directorate E" Then 'VALİLİK
            objDoc.Tables(1).Cell(Row:=BStr, Column:=1).Range.Text = IlBuyukHarf & " PROVINCIAL GOVERNORSHIP"
            objDoc.Tables(1).Cell(Row:=BStr + 1, Column:=1).Range.Text = Cells(ActiveCell.Row, 72).Value  'UCase(Replace(Replace(Cells(ActiveCell.Row, 72).Value, "i", "I"), "ı", "I"))
            objDoc.Tables(1).Cell(Row:=BStr + 2, Column:=1).Range.Text = "(" & Cells(ActiveCell.Row, 73).Value & ")"
            objDoc.Tables(1).Cell(Row:=BStr + 3, Column:=1).Range.Text = Bolum3
            objDoc.Tables(1).Cell(Row:=BStr + 3, Column:=1).Range.Font.Underline = wdUnderlineSingle
            M4 = True
            
            ustbilgimuhatap = IlKucukHarf & " Provincial Governorship " & Cells(ActiveCell.Row, 72).Value
        ElseIf Cells(ActiveCell.Row, 72).Value = "District Directorate B" Or Cells(ActiveCell.Row, 72).Value = "District Directorate C" Or _
        Cells(ActiveCell.Row, 72).Value = "District Directorate D" Or Cells(ActiveCell.Row, 72).Value = "District Directorate E" Then 'KAYMAKAMLIK
            objDoc.Tables(1).Cell(Row:=BStr, Column:=1).Range.Text = IlceBuyukHarf & " DISTRICT GOVERNORSHIP"
            objDoc.Tables(1).Cell(Row:=BStr + 1, Column:=1).Range.Text = Cells(ActiveCell.Row, 72).Value 'UCase(Replace(Replace(Cells(ActiveCell.Row, 72).Value, "i", "I"), "ı", "I"))
            objDoc.Tables(1).Cell(Row:=BStr + 2, Column:=1).Range.Text = "(" & Cells(ActiveCell.Row, 73).Value & ")"
            objDoc.Tables(1).Cell(Row:=BStr + 3, Column:=1).Range.Text = Bolum3
            objDoc.Tables(1).Cell(Row:=BStr + 3, Column:=1).Range.Font.Underline = wdUnderlineSingle
            M4 = True
            
            ustbilgimuhatap = IlceKucukHarf & " District Governorship " & Cells(ActiveCell.Row, 72).Value
        ElseIf InStr(Cells(ActiveCell.Row, 72).Value, "General Directorate") <> 0 Or InStr(Cells(ActiveCell.Row, 72).Value, "Regional Directorate") <> 0 Then
            objDoc.Tables(1).Cell(Row:=BStr, Column:=1).Range.Text = UCase(Replace(Replace(ThisWorkbook.Worksheets(2).Cells(6, 111).Value, "i", "I"), "ı", "I"))
            objDoc.Tables(1).Cell(Row:=BStr + 1, Column:=1).Range.Text = Cells(ActiveCell.Row, 72).Value 'UCase(Replace(Replace(Cells(ActiveCell.Row, 72).Value, "i", "I"), "ı", "I"))
            objDoc.Tables(1).Cell(Row:=BStr + 2, Column:=1).Range.Text = "(" & Cells(ActiveCell.Row, 73).Value & ")"
            objDoc.Tables(1).Cell(Row:=BStr + 3, Column:=1).Range.Text = Bolum3
            objDoc.Tables(1).Cell(Row:=BStr + 3, Column:=1).Range.Font.Underline = wdUnderlineSingle
            M4 = True
            
            ustbilgimuhatap = WorksheetFunction.Proper(ThisWorkbook.Worksheets(2).Cells(6, 111).Value) & " " & Cells(ActiveCell.Row, 72).Value
        Else 'YARGI 3'lük
            Bolum1 = UCase(Replace(Replace(Cells(ActiveCell.Row, 72).Value, "i", "I"), "ı", "I"))
            If InStr(Bolum1, "X.X. ") > 0 Then
                objDoc.Tables(1).Cell(Row:=BStr, Column:=1).Range.Text = Mid(Bolum1, 6, Len(Bolum1))
                objDoc.Tables(1).Cell(Row:=BStr + 1, Column:=1).Range.Text = "(" & Cells(ActiveCell.Row, 73).Value & ")"
                objDoc.Tables(1).Cell(Row:=BStr + 2, Column:=1).Range.Text = Bolum3
                objDoc.Tables(1).Cell(Row:=BStr + 2, Column:=1).Range.Font.Underline = wdUnderlineSingle
                M3 = True
                
                ustbilgimuhatap = Mid(Cells(ActiveCell.Row, 72).Value, 6, Len(Cells(ActiveCell.Row, 72).Value))
            Else
                objDoc.Tables(1).Cell(Row:=BStr, Column:=1).Range.Text = Bolum1
                objDoc.Tables(1).Cell(Row:=BStr + 1, Column:=1).Range.Text = "(" & Cells(ActiveCell.Row, 73).Value & ")"
                objDoc.Tables(1).Cell(Row:=BStr + 2, Column:=1).Range.Text = Bolum3
                objDoc.Tables(1).Cell(Row:=BStr + 2, Column:=1).Range.Font.Underline = wdUnderlineSingle
                M3 = True

                ustbilgimuhatap = Cells(ActiveCell.Row, 72).Value
            End If
        End If
    End If
    '3'lük
    If Cells(ActiveCell.Row, 73).Value = "" Then
        If Cells(ActiveCell.Row, 72).Value = "Provincial Directorate B" Or Cells(ActiveCell.Row, 72).Value = "Provincial Directorate C" Or _
        Cells(ActiveCell.Row, 72).Value = "Provincial Directorate D" Or Cells(ActiveCell.Row, 72).Value = "Provincial Directorate E" Then 'VALİLİK
            objDoc.Tables(1).Cell(Row:=BStr, Column:=1).Range.Text = IlBuyukHarf & " PROVINCIAL GOVERNORSHIP"
            objDoc.Tables(1).Cell(Row:=BStr + 1, Column:=1).Range.Text = "(" & Cells(ActiveCell.Row, 72).Value & ")"  'UCase(Replace(Replace(Cells(ActiveCell.Row, 72).Value, "i", "I"), "ı", "I"))
            objDoc.Tables(1).Cell(Row:=BStr + 2, Column:=1).Range.Text = Bolum3
            objDoc.Tables(1).Cell(Row:=BStr + 2, Column:=1).Range.Font.Underline = wdUnderlineSingle
            M3 = True
            
            ustbilgimuhatap = IlKucukHarf & " Provincial Governorship " & Cells(ActiveCell.Row, 72).Value
        ElseIf Cells(ActiveCell.Row, 72).Value = "District Directorate B" Or Cells(ActiveCell.Row, 72).Value = "District Directorate C" Or _
        Cells(ActiveCell.Row, 72).Value = "District Directorate D" Or Cells(ActiveCell.Row, 72).Value = "District Directorate E" Then
            objDoc.Tables(1).Cell(Row:=BStr, Column:=1).Range.Text = IlceBuyukHarf & " DISTRICT GOVERNORSHIP"
            objDoc.Tables(1).Cell(Row:=BStr + 1, Column:=1).Range.Text = "(" & Cells(ActiveCell.Row, 72).Value & ")"  'UCase(Replace(Replace(Cells(ActiveCell.Row, 72).Value, "i", "I"), "ı", "I"))
            objDoc.Tables(1).Cell(Row:=BStr + 2, Column:=1).Range.Text = Bolum3
            objDoc.Tables(1).Cell(Row:=BStr + 2, Column:=1).Range.Font.Underline = wdUnderlineSingle
            M3 = True
            
            ustbilgimuhatap = IlceKucukHarf & " District Governorship " & Cells(ActiveCell.Row, 72).Value
        ElseIf InStr(Cells(ActiveCell.Row, 72).Value, "General Directorate") <> 0 Or InStr(Cells(ActiveCell.Row, 72).Value, "Regional Directorate") <> 0 Then
            objDoc.Tables(1).Cell(Row:=BStr, Column:=1).Range.Text = UCase(Replace(Replace(ThisWorkbook.Worksheets(2).Cells(6, 111).Value, "i", "I"), "ı", "I"))
            objDoc.Tables(1).Cell(Row:=BStr + 1, Column:=1).Range.Text = "(" & Cells(ActiveCell.Row, 72).Value & ")" 'UCase(Replace(Replace(Cells(ActiveCell.Row, 72).Value, "i", "I"), "ı", "I"))
            objDoc.Tables(1).Cell(Row:=BStr + 2, Column:=1).Range.Text = Bolum3
            objDoc.Tables(1).Cell(Row:=BStr + 2, Column:=1).Range.Font.Underline = wdUnderlineSingle
            M3 = True
            
            ustbilgimuhatap = WorksheetFunction.Proper(ThisWorkbook.Worksheets(2).Cells(6, 111).Value) & " " & Cells(ActiveCell.Row, 72).Value
        Else 'YARGI 2'lik
            Bolum1 = UCase(Replace(Replace(Cells(ActiveCell.Row, 72).Value, "i", "I"), "ı", "I"))
            If InStr(Bolum1, "X.X. ") > 0 Then
                objDoc.Tables(1).Cell(Row:=BStr, Column:=1).Range.Text = Mid(Bolum1, 6, Len(Bolum1))
                objDoc.Tables(1).Cell(Row:=BStr + 1, Column:=1).Range.Text = Bolum3
                objDoc.Tables(1).Cell(Row:=BStr + 1, Column:=1).Range.Font.Underline = wdUnderlineSingle
                M2 = True

                ustbilgimuhatap = Mid(Cells(ActiveCell.Row, 72).Value, 6, Len(Cells(ActiveCell.Row, 72).Value))
            Else
                objDoc.Tables(1).Cell(Row:=BStr, Column:=1).Range.Text = Bolum1
                objDoc.Tables(1).Cell(Row:=BStr + 1, Column:=1).Range.Text = Bolum3
                objDoc.Tables(1).Cell(Row:=BStr + 1, Column:=1).Range.Font.Underline = wdUnderlineSingle
                M2 = True

                ustbilgimuhatap = Cells(ActiveCell.Row, 72).Value
            End If
        End If
    End If
   

    'İlgi tema
    If Cells(IlkSira, 33).Value = "Provincial Directorate B" Then
        If Cells(IlkSira, 34).Value <> "" Then
            GelenTema = Cells(IlkSira, 25).Value & " Provincial Governorship Provincial Directorate B " & Cells(IlkSira, 34).Value
        Else
            GelenTema = Cells(IlkSira, 25).Value & " Provincial Governorship Provincial Directorate B"
        End If
    ElseIf Cells(IlkSira, 33).Value = "District Directorate B" Then
        If Cells(IlkSira, 34).Value <> "" Then
            GelenTema = Cells(IlkSira, 26).Value & " District Governorship District Directorate B " & Cells(IlkSira, 34).Value
        Else
            GelenTema = Cells(IlkSira, 26).Value & " District Governorship District Directorate B"
        End If
    ElseIf Cells(IlkSira, 33).Value = "Provincial Directorate C" Then
        If Cells(IlkSira, 34).Value <> "" Then
            GelenTema = Cells(IlkSira, 25).Value & " Provincial Governorship Provincial Directorate C " & Cells(IlkSira, 34).Value
        Else
            GelenTema = Cells(IlkSira, 25).Value & " Provincial Governorship Provincial Directorate C"
        End If
    ElseIf Cells(IlkSira, 33).Value = "District Directorate C" Then
        If Cells(IlkSira, 34).Value <> "" Then
            GelenTema = Cells(IlkSira, 26).Value & " District Governorship District Directorate C " & Cells(IlkSira, 34).Value
        Else
            GelenTema = Cells(IlkSira, 26).Value & " District Governorship District Directorate C"
        End If
    ElseIf Cells(IlkSira, 33).Value = "Provincial Directorate D" Then
        If Cells(IlkSira, 34).Value <> "" Then
            GelenTema = Cells(IlkSira, 25).Value & " Provincial Governorship Provincial Directorate D " & Cells(IlkSira, 34).Value
        Else
            GelenTema = Cells(IlkSira, 25).Value & " Provincial Governorship Provincial Directorate D"
        End If
    ElseIf Cells(IlkSira, 33).Value = "District Directorate D" Then
        If Cells(IlkSira, 34).Value <> "" Then
            GelenTema = Cells(IlkSira, 26).Value & " District Governorship District Directorate D " & Cells(IlkSira, 34).Value
        Else
            GelenTema = Cells(IlkSira, 26).Value & " District Governorship District Directorate D"
        End If
    ElseIf Cells(IlkSira, 33).Value = "Provincial Directorate E" Then
        If Cells(IlkSira, 34).Value <> "" Then
            GelenTema = Cells(IlkSira, 25).Value & " Provincial Governorship Provincial Directorate E " & Cells(IlkSira, 34).Value
        Else
            GelenTema = Cells(IlkSira, 25).Value & " Provincial Governorship Provincial Directorate E"
        End If
    ElseIf Cells(IlkSira, 33).Value = "District Directorate E" Then
        If Cells(IlkSira, 34).Value <> "" Then
            GelenTema = Cells(IlkSira, 26).Value & " District Governorship District Directorate E " & Cells(IlkSira, 34).Value
        Else
            GelenTema = Cells(IlkSira, 26).Value & " District Governorship District Directorate E"
        End If
    ElseIf InStr(Cells(IlkSira, 33).Value, "General Directorate") <> 0 Or InStr(Cells(IlkSira, 33).Value, "Regional Directorate") <> 0 Then
        If Cells(IlkSira, 34).Value <> "" Then
            GelenTema = WorksheetFunction.Proper(ThisWorkbook.Worksheets(2).Cells(6, 111).Value) & " " & Cells(IlkSira, 33).Value & " " & Cells(IlkSira, 34).Value
        Else
            GelenTema = WorksheetFunction.Proper(ThisWorkbook.Worksheets(2).Cells(6, 111).Value) & " " & Cells(IlkSira, 33).Value
        End If
    Else
        If InStr(Cells(IlkSira, 33).Value, "X.X. ") > 0 Then
            If Cells(IlkSira, 34).Value <> "" Then
                GelenTema = Mid(Cells(IlkSira, 33).Value, 6, Len(Cells(IlkSira, 33).Value)) & " " & Cells(IlkSira, 34).Value
            Else
                GelenTema = Mid(Cells(IlkSira, 33).Value, 6, Len(Cells(IlkSira, 33).Value))
            End If
        Else
            If Cells(IlkSira, 34).Value <> "" Then
                GelenTema = Cells(IlkSira, 33).Value & " " & Cells(IlkSira, 34).Value
            Else
                GelenTema = Cells(IlkSira, 33).Value
            End If
        End If
    End If

    'İlgi
    Bolum1 = "Delivered to us on "
    Bolum2 = ", with a letter dated "
    Bolum3 = " and reference number "
    Bolum4 = "."
    
    If Cells(ActiveCell.Row, 33).Value = Cells(ActiveCell.Row, 72).Value And _
    Cells(ActiveCell.Row, 34).Value = Cells(ActiveCell.Row, 73).Value And _
    Cells(ActiveCell.Row, 25).Value = Cells(ActiveCell.Row, 77).Value Then
           
        ' Gelen ve giden birim aynıysa
        objDoc.Tables(1).Cell(Row:=15, Column:=3).Range.Text = _
            Bolum1 & Cells(ActiveCell.Row, 36).Value & Bolum2 & _
            Cells(ActiveCell.Row, 28).Value & Bolum3 & _
            Cells(ActiveCell.Row, 29).Value & Bolum4
    
        ' İlgi yazı eki kontrolü
        If Trim(Cells(ActiveCell.Row, 82).Value) <> "" Then
            Ifv = True
        Else
            Ify = True
        End If
    
    Else
        ' Gelen ve giden birim farklıysa
        RaporIlgi = Bolum1 & Cells(ActiveCell.Row, 36).Value & _
                    " from " & GelenTema & Bolum2 & _
                    Cells(IlkSira, 28).Value & Bolum3 & _
                    Cells(IlkSira, 29).Value & Bolum4
    
        objDoc.Tables(1).Cell(Row:=15, Column:=3).Range.Text = RaporIlgi
        Ifv = True
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
            ReportList(y) = Cells(i, 11).Value
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
        Ek4 = "report"
        Ek5 = "Report 2.1"
    Else
        Ek4 = "reports"
        Ek5 = "Report 2.1s"
    End If
    
    AdetTopla = Application.Sum(Range(Cells(IlkSira, 52), Cells(SonSira, 52)))
    If AdetTopla = 1 Then
        Ek1 = "The Type A item, sent to our unit as an enclosure to the referenced letter, has been documented in "
        Ek2 = "item"
    Else
        Ek1 = "The Type A items, sent to our unit as an enclosure to the referenced letter, have been documented in "
        Ek2 = "items"
    End If

    ' Delivery method
    GonderimUsulu = Mid(Cells(ActiveCell.Row, 75).Value, InStr(Cells(ActiveCell.Row, 75).Value, "/") + 1)
    If GonderimUsulu = "HAND DELIVERY" Then
        GonderimUsulu = "hand-delivered to you."
    Else
        GonderimUsulu = "sent to your office accordingly."
    End If
    
    ' Construct the text
    Bolum1 = Ek1 & Ek5 & " "
    Bolum2 = "dated " & Cells(ActiveCell.Row, 68).Value & " and numbered " & Ek3 & ". "
    Bolum3 = "Both the " & Ek2 & " and the corresponding " & Ek4 & " have been " & GonderimUsulu
    Bolum6 = Chr(9) & "Respectfully submitted for your information."
    Birlestir = Chr(9) & Bolum1 & Bolum2 & Bolum3
 
    'Word'ten metni çek ve word'teki satırı temizle
    StrTemaNotu = objDoc.Tables(2).Cell(Row:=1, Column:=1).Range.Text
    objDoc.Tables(2).Cell(Row:=1, Column:=1).Range.Text = ""
    'Çekilen metnin sonunda breakline varsa kaldır
    If Len(StrTemaNotu) > 0 Then
        If InStr(StrTemaNotu, Chr(13)) <> 0 Then
            StrTemaNotu = Replace(StrTemaNotu, Chr(13), "")
            'StrTemaNotu = Left(StrTemaNotu, Len(StrTemaNotu) - 2)
        End If
    End If
    'Tema notu varsa Birlestir değeri şöyle olsun
    If Cells(ActiveCell.Row, 86).Value = "Yes" Then
        Birlestir = Birlestir & vbNewLine & StrTemaNotu
    End If
    
    objDoc.Tables(2).Cell(Row:=1, Column:=1).Range.Text = Birlestir
    objDoc.Tables(2).Cell(Row:=3, Column:=1).Range.Text = Bolum6
    
    
    'invalidleri say
    gecersizSay = 0
    For i = IlkSira To SonSira
        If Cells(i, 62).Value = "invalid" Then
            gecersizSay = gecersizSay + Cells(i, 52).Value
        End If
    Next i
    'Üst yazı notu (XXXMud notu)
    If Cells(ActiveCell.Row, 85).Value = "Yes" Then
        XXXMudNotu = True
    Else
        XXXMudNotu = False
        objDoc.Tables(2).Cell(Row:=2, Column:=1).Range.Text = ""
        'objDoc.Tables(2).Rows(2).Delete
    End If
    
    'Tema notu
    If Cells(ActiveCell.Row, 86).Value = "Yes" Then
        'XXXMudNotu = True
        Set MyRange = objDoc.Tables(2).Cell(Row:=1, Column:=1).Range
        MyRange.Find.Execute FindText:="""" & "xxxxx xxxxxxxxx xxxxxxxxxxxxx xx xxxxxxxxxxx xxxxxxxxxx xxxxxxxx xx xxxxxx xxxxxxxx xxxxxxxx" & """"
        MyRange.Font.Bold = True

        If y = 1 Then
            Ek5 = "was"
        Else
            Ek5 = "were"
        End If
        
        'item / items
        Set MyRange = objDoc.Tables(2).Cell(Row:=1, Column:=1).Range
        With MyRange.Find
            .Text = "<item_1>"
            .Replacement.Text = Ek2
            .Execute Replace:=wdReplaceAll
        End With
        'report / reports
        Set MyRange = objDoc.Tables(2).Cell(Row:=1, Column:=1).Range
        With MyRange.Find
            .Text = "<report>"
            .Replacement.Text = Ek4
            .Execute Replace:=wdReplaceAll
        End With
        'Single or multiple report no.
        Set MyRange = objDoc.Tables(2).Cell(Row:=1, Column:=1).Range
        With MyRange.Find
            .Text = "<report no.>"
            .Replacement.Text = Ek3
            .Execute Replace:=wdReplaceAll
        End With
        ' was / were
        Set MyRange = objDoc.Tables(2).Cell(Row:=1, Column:=1).Range
        With MyRange.Find
            .Text = "<was-were>"
            .Replacement.Text = Ek5
            .Execute Replace:=wdReplaceAll
        End With
        'MyRange.Find.Execute FindText:=Ek3
        'MyRange.Font.Bold = True
    Else
        'XXXMudNotu = False
    End If
     
    
    'imzalar
    objDoc.Tables(3).Cell(Row:=4, Column:=2).Range.Text = Cells(ActiveCell.Row, 130).Value 'Ad Soyad1
    objDoc.Tables(3).Cell(Row:=5, Column:=2).Range.Text = Cells(ActiveCell.Row, 131).Value 'Unvan1
    objDoc.Tables(3).Cell(Row:=4, Column:=3).Range.Text = Cells(ActiveCell.Row, 133).Value 'Ad Soyad2
    objDoc.Tables(3).Cell(Row:=5, Column:=3).Range.Text = Cells(ActiveCell.Row, 134).Value 'Unvan2

    'Ekler
    objDoc.Tables(3).Cell(Row:=8, Column:=1).Range.Text = "Attachment:"

    Ek2 = Cells(ActiveCell.Row, 75).Value 'Kapalı Package A
    Ek2 = Left(Ek2, InStr(Ek2, "/") - 1)
    If Cells(ActiveCell.Row, 76).Value > 1 Then objDoc.Tables(3).Cell(Row:=8, Column:=2).Range.Text = " 1) Enclosed " & Ek2 & " (" & Cells(ActiveCell.Row, 76).Value & " pieces)"
    If Cells(ActiveCell.Row, 76).Value < 2 Then objDoc.Tables(3).Cell(Row:=8, Column:=2).Range.Text = " 1) Enclosed " & Ek2 & " (" & Cells(ActiveCell.Row, 76).Value & " piece)"
    
    x = Application.Sum(Range(Cells(IlkSira, 100), Cells(SonSira, 100))) 'Statement 2 toplam sayfa sayısı
    If x > 1 Then objDoc.Tables(3).Cell(Row:=9, Column:=2).Range.Text = " 2) Statement 2 (" & x & " pages)"
    If x < 2 Then objDoc.Tables(3).Cell(Row:=9, Column:=2).Range.Text = " 2) Statement 2 (" & x & " page)"
    
    x = Application.Sum(Range(Cells(IlkSira, 99), Cells(SonSira, 99)))  'Rapor1 toplam sayfa sayısı
    If y = 1 Then
        If x > 1 Then objDoc.Tables(3).Cell(Row:=10, Column:=2).Range.Text = " 3) Report 2.1 (" & x & " pages)"
        If x < 2 Then objDoc.Tables(3).Cell(Row:=10, Column:=2).Range.Text = " 3) Report 2.1 (" & x & " page)"
    Else
        objDoc.Tables(3).Cell(Row:=10, Column:=2).Range.Text = " 3) Report 2.1 (" & y & " reports, " & "total of " & x & " pages)"
    End If
    
    'Statement 1 (döküm dahil) toplam sayfa sayısı
    If Cells(ActiveCell.Row, 98).Value = "" Then
        x = Application.Sum(Range(Cells(IlkSira, 97), Cells(SonSira, 97)))
        If x > 1 Then objDoc.Tables(3).Cell(Row:=11, Column:=2).Range.Text = " 4) Statement 1 (" & x & " pages)"
        If x < 2 Then objDoc.Tables(3).Cell(Row:=11, Column:=2).Range.Text = " 4) Statement 1 (" & x & " page)"
    Else
        x = Application.Sum(Range(Cells(IlkSira, 97), Cells(SonSira, 98)))
        objDoc.Tables(3).Cell(Row:=11, Column:=2).Range.Text = " 4) Attached Statement 1 (total of " & x & " pages)"
    End If
        
    If Cells(IlkSira, 82).Value <> "" Then 'ilgi yazı fotokopisi var
        'İlgi yazı fotokopisi
        If Cells(IlkSira, 74).Value > 1 Then objDoc.Tables(3).Cell(Row:=12, Column:=2).Range.Text = " 5) Photocopy of the Referenced Letter (" & Cells(IlkSira, 82).Value & " pages)"
        If Cells(IlkSira, 74).Value < 2 Then objDoc.Tables(3).Cell(Row:=12, Column:=2).Range.Text = " 5) Photocopy of the Referenced Letter (" & Cells(IlkSira, 82).Value & " page)"
    Else 'ilgi yazı fotokopisi yok
        '
    End If


    'On Error GoTo 0


    '__________________________DİNAMİK SAYFA YAPISI
    
    'İlgi ve gövde metninin kaç satırdan oluştuğu bilgisinin elde edilmesinde kullanılıyor.
    Set IlgiRng = objDoc.Tables(1).Cell(Row:=15, Column:=3).Range
    Set Govde1Rng = objDoc.Tables(2).Cell(Row:=1, Column:=1).Range
    Set Govde2Rng = objDoc.Tables(2).Cell(Row:=2, Column:=1).Range
    
    IlgiStrSay = Govde1Rng.Information(wdFirstCharacterLineNumber) - IlgiRng.Information(wdFirstCharacterLineNumber) - 1
    Govde1StrSay = Govde2Rng.Information(wdFirstCharacterLineNumber) - Govde1Rng.Information(wdFirstCharacterLineNumber)
    'MsgBox "İlgi Satır Sayısı: " & IlgiStrSay & vbNewLine & vbNewLine & "Gövde 1 Satır Sayısı: " & Govde1StrSay
    IlgiFarkSay = IlgiStrSay - 1
    Govde1FarkSay = Govde1StrSay - 2
    CokluSayfa = 0

    'MsgBox IlgiFarkSay
    'MsgBox Govde1FarkSay
    'MsgBox IlgiFarkSay + Govde1FarkSay 'Varsayılan ilgi ve 1 paragrafta (2 rowda) toplam 3 satıra göre sıfırlandı.
    
    'Dinamik sayfa düzeni
    If Cells(ActiveCell.Row, 86).Value = "Yes" Then 'Tema notu var.

        'yazı tipini küçült
        Set Govde1Rng = objDoc.Tables(2).Cell(Row:=1, Column:=1).Range
        Set Govde2Rng = objDoc.Tables(2).Cell(Row:=2, Column:=1).Range

        objDoc.Range.Font.Size = 11
        objDoc.Tables(1).Cell(Row:=5, Column:=1).Range.Font.Size = 13
        objDoc.Tables(1).Cell(Row:=2, Column:=1).Range.Font.Size = 9
            
        If XXXMudNotu = True Then 'XXXMud Notu var.
            If Ifv = True Then 'Ekte ilgi fotokopisi var.
                Govde1FarkSay = Govde1FarkSay + 2
            Else
                Govde1FarkSay = Govde1FarkSay + 1
                objDoc.Tables(3).Rows(12).Delete
            End If
        Else
            objDoc.Tables(2).Rows(2).Delete 'XXXMud Notu
            If Ifv = True Then 'Ekte ilgi fotokopisi var.
                Govde1FarkSay = Govde1FarkSay - 1
            Else
                objDoc.Tables(3).Rows(12).Delete
                Govde1FarkSay = Govde1FarkSay - 2
            End If
        End If
        
    Else

        If XXXMudNotu = True Then 'XXXMud Notu var.
            If Ifv = True Then 'Ekte ilgi fotokopisi var.
                Govde1FarkSay = Govde1FarkSay + 5
            Else
                Govde1FarkSay = Govde1FarkSay + 4
                objDoc.Tables(3).Rows(12).Delete
            End If
        Else
            objDoc.Tables(2).Rows(2).Delete 'XXXMud Notu
            If Ifv = True Then 'Ekte ilgi fotokopisi var.
                Govde1FarkSay = Govde1FarkSay + 2
            Else
                objDoc.Tables(3).Rows(12).Delete
                Govde1FarkSay = Govde1FarkSay + 1
            End If
        End If
        
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
    Cells(ActiveCell.Row, 101).Value = TotalSayfaUstYazi

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

'Worksheets(4).Activate
Application.ScreenUpdating = True
Application.DisplayAlerts = True
'Application.EnableEvents = True

'End If

End Sub

Sub IslemGunluguRapor2_1()
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
Set WsRapor = ThisWorkbook.Worksheets(4)
'WsRapor.Unprotect Password:="123"

IlceSakla = ""
If InStr(WsRapor.Cells(IlkSira, 26).Value, " Organization A") <> 0 Then
    IlceSakla = WsRapor.Cells(IlkSira, 26).Value
    WsRapor.Cells(IlkSira, 26).Value = ""
End If


'Aylık ayraçlar
If WsRapor.Cells(IlkSira, 39).Value <> "" Then
    ModulTarih = WsRapor.Cells(IlkSira, 39).Value
    ModulAyrac = "01" & Right(ModulTarih, 8)
Else 'işlemin yapıldığı günü esas al
    ModulTarih = Format(Date, "dd.mm.yyyy")
    ModulAyrac = "01" & Right(ModulTarih, 8)
End If

'KayitDefSiraNo = 0

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
If WsRapor.Cells(IlkSira, 33).Value = "Provincial Directorate B" Then
    If WsRapor.Cells(IlkSira, 34).Value <> "" Then
        GelenTema = WsRapor.Cells(IlkSira, 25).Value & " Provincial Governorship Provincial Directorate B " & WsRapor.Cells(IlkSira, 34).Value
    Else
        GelenTema = WsRapor.Cells(IlkSira, 25).Value & " Provincial Governorship Provincial Directorate B"
    End If
ElseIf WsRapor.Cells(IlkSira, 33).Value = "District Directorate B" Then
    If WsRapor.Cells(IlkSira, 34).Value <> "" Then
        GelenTema = WsRapor.Cells(IlkSira, 26).Value & " District Governorship District Directorate B " & WsRapor.Cells(IlkSira, 34).Value
    Else
        GelenTema = WsRapor.Cells(IlkSira, 26).Value & " District Governorship District Directorate B"
    End If
ElseIf WsRapor.Cells(IlkSira, 33).Value = "Provincial Directorate C" Then
    If WsRapor.Cells(IlkSira, 34).Value <> "" Then
        GelenTema = WsRapor.Cells(IlkSira, 25).Value & " Provincial Governorship Provincial Directorate C " & WsRapor.Cells(IlkSira, 34).Value
    Else
        GelenTema = WsRapor.Cells(IlkSira, 25).Value & " Provincial Governorship Provincial Directorate C"
    End If
ElseIf WsRapor.Cells(IlkSira, 33).Value = "District Directorate C" Then
    If WsRapor.Cells(IlkSira, 34).Value <> "" Then
        GelenTema = WsRapor.Cells(IlkSira, 26).Value & " District Governorship District Directorate C " & WsRapor.Cells(IlkSira, 34).Value
    Else
        GelenTema = WsRapor.Cells(IlkSira, 26).Value & " District Governorship District Directorate C"
    End If
ElseIf WsRapor.Cells(IlkSira, 33).Value = "Provincial Directorate D" Then
    If WsRapor.Cells(IlkSira, 34).Value <> "" Then
        GelenTema = WsRapor.Cells(IlkSira, 25).Value & " Provincial Governorship Provincial Directorate D " & WsRapor.Cells(IlkSira, 34).Value
    Else
        GelenTema = WsRapor.Cells(IlkSira, 25).Value & " Provincial Governorship Provincial Directorate D"
    End If
ElseIf WsRapor.Cells(IlkSira, 33).Value = "District Directorate D" Then
    If WsRapor.Cells(IlkSira, 34).Value <> "" Then
        GelenTema = WsRapor.Cells(IlkSira, 26).Value & " District Governorship District Directorate D " & WsRapor.Cells(IlkSira, 34).Value
    Else
        GelenTema = WsRapor.Cells(IlkSira, 26).Value & " District Governorship District Directorate D"
    End If
ElseIf WsRapor.Cells(IlkSira, 33).Value = "Provincial Directorate E" Then
    If WsRapor.Cells(IlkSira, 34).Value <> "" Then
        GelenTema = WsRapor.Cells(IlkSira, 25).Value & " Provincial Governorship Provincial Directorate E " & WsRapor.Cells(IlkSira, 34).Value
    Else
        GelenTema = WsRapor.Cells(IlkSira, 25).Value & " Provincial Governorship Provincial Directorate E"
    End If
ElseIf WsRapor.Cells(IlkSira, 33).Value = "District Directorate E" Then
    If WsRapor.Cells(IlkSira, 34).Value <> "" Then
        GelenTema = WsRapor.Cells(IlkSira, 26).Value & " District Governorship District Directorate E " & WsRapor.Cells(IlkSira, 34).Value
    Else
        GelenTema = WsRapor.Cells(IlkSira, 26).Value & " District Governorship District Directorate E"
    End If
ElseIf InStr(WsRapor.Cells(IlkSira, 33).Value, "General Directorate") <> 0 Or InStr(WsRapor.Cells(IlkSira, 33).Value, "Regional Directorate") <> 0 Then
    If WsRapor.Cells(IlkSira, 34).Value <> "" Then
        GelenTema = WorksheetFunction.Proper(ThisWorkbook.Worksheets(2).Cells(6, 111).Value) & " " & WsRapor.Cells(IlkSira, 33).Value & " " & WsRapor.Cells(IlkSira, 34).Value
    Else
        GelenTema = WorksheetFunction.Proper(ThisWorkbook.Worksheets(2).Cells(6, 111).Value) & " " & WsRapor.Cells(IlkSira, 33).Value
    End If
Else
    If WsRapor.Cells(IlkSira, 34).Value <> "" Then
        If InStr(WsRapor.Cells(IlkSira, 33).Value, "X.X. ") > 0 Then
            GelenTema = Mid(WsRapor.Cells(IlkSira, 33).Value, 6, Len(WsRapor.Cells(IlkSira, 33).Value)) & " " & WsRapor.Cells(IlkSira, 34).Value
        Else
            GelenTema = WsRapor.Cells(IlkSira, 33).Value & " " & WsRapor.Cells(IlkSira, 34).Value
        End If
    Else
        If InStr(WsRapor.Cells(IlkSira, 33).Value, "X.X. ") > 0 Then
            GelenTema = Mid(WsRapor.Cells(IlkSira, 33).Value, 6, Len(WsRapor.Cells(IlkSira, 33).Value))
        Else
            GelenTema = WsRapor.Cells(IlkSira, 33).Value
        End If
    End If
End If



'____________OPERASYONLAR

'İşlem günlüğünde başlangıç ve bitiş satırlarını tespit et.
Say1IslemGunlugu = WsIslemGunlugu.Range("B100000").End(xlUp).Row
Say2IslemGunlugu = WsIslemGunlugu.Range("C100000").End(xlUp).Row
SayAyracIslemGunlugu = WsIslemGunlugu.Range("E100000").End(xlUp).Row

Set IslemGunluguIlkSiraBul = WsIslemGunlugu.Range("B7:B100000").Find(What:=WsRapor.Cells(IlkSira, 93).Value, SearchDirection:=xlNext, _
                SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole) 'zaman damgasını ara
Set IslemGunluguSonSiraBul = WsIslemGunlugu.Range("C7:C100000").Find(What:=WsRapor.Cells(IlkSira, 93).Value, SearchDirection:=xlNext, _
                SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
If Not IslemGunluguIlkSiraBul Is Nothing Then
    IslemGunluguIlkSira = IslemGunluguIlkSiraBul.Row
    If Not IslemGunluguSonSiraBul Is Nothing Then
        IslemGunluguSonSira = IslemGunluguSonSiraBul.Row
    End If
End If

    
If Not IslemGunluguIlkSiraBul Is Nothing Then 'DÜZENLEME İŞLEMİ

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
WsIslemGunlugu.Cells(ilkrow, 2).Value = WsRapor.Cells(IlkSira, 93).Value
WsIslemGunlugu.Cells(sonrow, 3).Value = WsRapor.Cells(IlkSira, 93).Value
'Verileri yaz
WsIslemGunlugu.Range(WsIslemGunlugu.Cells(ilkrow, 7), WsIslemGunlugu.Cells(sonrow, 7)).Value = WsRapor.Range(WsRapor.Cells(IlkSira, 11), WsRapor.Cells(SonSira, 11)).Value 'Rapor no
WsIslemGunlugu.Cells(ilkrow, 8).Value = WsRapor.Cells(IlkSira, 25).Value 'İl
WsIslemGunlugu.Cells(ilkrow, 9).Value = WsRapor.Cells(IlkSira, 26).Value 'İlçe
WsIslemGunlugu.Cells(ilkrow, 10).Value = GelenTema
WsIslemGunlugu.Cells(ilkrow, 11).Value = WsRapor.Cells(IlkSira, 28).Value 'Belge tarihi
WsIslemGunlugu.Cells(ilkrow, 12).Value = WsRapor.Cells(IlkSira, 29).Value 'Belge no
WsIslemGunlugu.Cells(ilkrow, 13).Value = WsRapor.Cells(IlkSira, 36).Value 'finansal birime ulaşma tarihi
WsIslemGunlugu.Cells(ilkrow, 14).Value = WsRapor.Cells(IlkSira, 39).Value 'Tespit tarihi
WsIslemGunlugu.Range(WsIslemGunlugu.Cells(ilkrow, 15), WsIslemGunlugu.Cells(sonrow, 15)).Value = WsRapor.Range(WsRapor.Cells(IlkSira, 46), WsRapor.Cells(SonSira, 46)).Value 'Öğe türü
WsIslemGunlugu.Range(WsIslemGunlugu.Cells(ilkrow, 16), WsIslemGunlugu.Cells(sonrow, 16)).Value = WsRapor.Range(WsRapor.Cells(IlkSira, 49), WsRapor.Cells(SonSira, 49)).Value 'Öğe değeri
WsIslemGunlugu.Range(WsIslemGunlugu.Cells(ilkrow, 17), WsIslemGunlugu.Cells(sonrow, 17)).Value = WsRapor.Range(WsRapor.Cells(IlkSira, 52), WsRapor.Cells(SonSira, 52)).Value 'Adet
WsIslemGunlugu.Cells(ilkrow, 18).Value = WsRapor.Cells(IlkSira, 31).Value 'Tema
WsIslemGunlugu.Range(WsIslemGunlugu.Cells(ilkrow, 19), WsIslemGunlugu.Cells(sonrow, 19)).Value = WsRapor.Range(WsRapor.Cells(IlkSira, 58), WsRapor.Cells(SonSira, 58)).Value 'Açıklama
WsIslemGunlugu.Range(WsIslemGunlugu.Cells(ilkrow, 20), WsIslemGunlugu.Cells(sonrow, 20)).Value = WsRapor.Range(WsRapor.Cells(IlkSira, 63), WsRapor.Cells(SonSira, 63)).Value 'Baskı tekniği
'Rapor2_2 işareti (Technique A içeren rapor numaralarını rapor2_2 numarası ile işaretle)
If WsRapor.Cells(IlkSira, 65).Value <> "" Then
    For i = ilkrow To sonrow
        If Left(WsIslemGunlugu.Cells(i, 20).Value, 11) = "Technique A" Then
            If WsIslemGunlugu.Cells(i, 7).Value <> "" Then
                WsIslemGunlugu.Cells(i, 7).Value = WsIslemGunlugu.Cells(i, 7).Value & " (B/" & WsRapor.Cells(IlkSira, 65).Value & ")"
            End If
        End If
    Next i
End If

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
    WsRapor.Cells(IlkSira, 26).Value = IlceSakla
End If


'WsRapor.Protect Password:="123"

Out:

Application.DisplayAlerts = True
Application.EnableEvents = True
Application.ScreenUpdating = True


End Sub

Sub IslemGunluguRapor2_2()

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
Dim StrContent As String


Application.DisplayAlerts = False
Application.EnableEvents = False
Application.ScreenUpdating = False

'Pathfinder...
AutoPath = ThisWorkbook.Path
IslemGunlukleriKlasor = AutoPath & "\System Files\System Templates\Registry Reports\"
IslemGunlugu = IslemGunlukleriKlasor & "System Registry Report 2.2.xlsx"

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
Set WsRapor = ThisWorkbook.Worksheets(4)
'WsRapor.Unprotect Password:="123"


IlceSakla = ""
If InStr(WsRapor.Cells(IlkSira, 26).Value, " Organization A") <> 0 Then
    IlceSakla = WsRapor.Cells(IlkSira, 26).Value
    WsRapor.Cells(IlkSira, 26).Value = ""
End If

'Aylık ayraçlar
If WsRapor.Cells(IlkSira, 39).Value <> "" Then
    ModulTarih = WsRapor.Cells(IlkSira, 39).Value
    ModulAyrac = "01" & Right(ModulTarih, 8)
Else 'işlemin yapıldığı günü esas al
    ModulTarih = Format(Date, "dd.mm.yyyy")
    ModulAyrac = "01" & Right(ModulTarih, 8)
End If


'İşlem günlüğü açıksa kaydet ve kapat.
OpenControl = IsWorkBookOpen(IslemGunlugu)
If OpenControl = True Then
    Workbooks("System Registry Report 2.2.xlsx").Save
    Workbooks("System Registry Report 2.2.xlsx").Close SaveChanges:=False
End If

'İşlem günlüğü aç
Workbooks.Open (IslemGunlugu)
Set WsIslemGunlugu = Workbooks("System Registry Report 2.2.xlsx").Worksheets(1)

WsIslemGunlugu.Unprotect Password:="123"

WsIslemGunlugu.Columns("B:C").EntireColumn.Hidden = False



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
        With WsIslemGunlugu.Range("E" & i & ":S" & i)
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
        With WsIslemGunlugu.Range("E" & i & ":S" & i)
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
        With WsIslemGunlugu.Range("E" & i & ":S" & i)
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


'GELEN TEMA
GelenTema = ""
If WsRapor.Cells(IlkSira, 33).Value = "Provincial Directorate B" Then
    If WsRapor.Cells(IlkSira, 34).Value <> "" Then
        GelenTema = WsRapor.Cells(IlkSira, 25).Value & " Provincial Governorship Provincial Directorate B " & WsRapor.Cells(IlkSira, 34).Value
    Else
        GelenTema = WsRapor.Cells(IlkSira, 25).Value & " Provincial Governorship Provincial Directorate B"
    End If
ElseIf WsRapor.Cells(IlkSira, 33).Value = "District Directorate B" Then
    If WsRapor.Cells(IlkSira, 34).Value <> "" Then
        GelenTema = WsRapor.Cells(IlkSira, 26).Value & " District Governorship District Directorate B " & WsRapor.Cells(IlkSira, 34).Value
    Else
        GelenTema = WsRapor.Cells(IlkSira, 26).Value & " District Governorship District Directorate B"
    End If
ElseIf WsRapor.Cells(IlkSira, 33).Value = "Provincial Directorate C" Then
    If WsRapor.Cells(IlkSira, 34).Value <> "" Then
        GelenTema = WsRapor.Cells(IlkSira, 25).Value & " Provincial Governorship Provincial Directorate C " & WsRapor.Cells(IlkSira, 34).Value
    Else
        GelenTema = WsRapor.Cells(IlkSira, 25).Value & " Provincial Governorship Provincial Directorate C"
    End If
ElseIf WsRapor.Cells(IlkSira, 33).Value = "District Directorate C" Then
    If WsRapor.Cells(IlkSira, 34).Value <> "" Then
        GelenTema = WsRapor.Cells(IlkSira, 26).Value & " District Governorship District Directorate C " & WsRapor.Cells(IlkSira, 34).Value
    Else
        GelenTema = WsRapor.Cells(IlkSira, 26).Value & " District Governorship District Directorate C"
    End If
ElseIf WsRapor.Cells(IlkSira, 33).Value = "Provincial Directorate D" Then
    If WsRapor.Cells(IlkSira, 34).Value <> "" Then
        GelenTema = WsRapor.Cells(IlkSira, 25).Value & " Provincial Governorship Provincial Directorate D " & WsRapor.Cells(IlkSira, 34).Value
    Else
        GelenTema = WsRapor.Cells(IlkSira, 25).Value & " Provincial Governorship Provincial Directorate D"
    End If
ElseIf WsRapor.Cells(IlkSira, 33).Value = "District Directorate D" Then
    If WsRapor.Cells(IlkSira, 34).Value <> "" Then
        GelenTema = WsRapor.Cells(IlkSira, 26).Value & " District Governorship District Directorate D " & WsRapor.Cells(IlkSira, 34).Value
    Else
        GelenTema = WsRapor.Cells(IlkSira, 26).Value & " District Governorship District Directorate D"
    End If
ElseIf WsRapor.Cells(IlkSira, 33).Value = "Provincial Directorate E" Then
    If WsRapor.Cells(IlkSira, 34).Value <> "" Then
        GelenTema = WsRapor.Cells(IlkSira, 25).Value & " Provincial Governorship Provincial Directorate E " & WsRapor.Cells(IlkSira, 34).Value
    Else
        GelenTema = WsRapor.Cells(IlkSira, 25).Value & " Provincial Governorship Provincial Directorate E"
    End If
ElseIf WsRapor.Cells(IlkSira, 33).Value = "District Directorate E" Then
    If WsRapor.Cells(IlkSira, 34).Value <> "" Then
        GelenTema = WsRapor.Cells(IlkSira, 26).Value & " District Governorship District Directorate E " & WsRapor.Cells(IlkSira, 34).Value
    Else
        GelenTema = WsRapor.Cells(IlkSira, 26).Value & " District Governorship District Directorate E"
    End If
ElseIf InStr(WsRapor.Cells(IlkSira, 33).Value, "General Directorate") <> 0 Or InStr(WsRapor.Cells(IlkSira, 33).Value, "Regional Directorate") <> 0 Then
    If WsRapor.Cells(IlkSira, 34).Value <> "" Then
        GelenTema = WorksheetFunction.Proper(ThisWorkbook.Worksheets(2).Cells(6, 111).Value) & " " & WsRapor.Cells(IlkSira, 33).Value & " " & WsRapor.Cells(IlkSira, 34).Value
    Else
        GelenTema = WorksheetFunction.Proper(ThisWorkbook.Worksheets(2).Cells(6, 111).Value) & " " & WsRapor.Cells(IlkSira, 33).Value
    End If
Else
    If WsRapor.Cells(IlkSira, 34).Value <> "" Then
        If InStr(WsRapor.Cells(IlkSira, 33).Value, "X.X. ") > 0 Then
            GelenTema = Mid(WsRapor.Cells(IlkSira, 33).Value, 6, Len(WsRapor.Cells(IlkSira, 33).Value)) & " " & WsRapor.Cells(IlkSira, 34).Value
        Else
            GelenTema = WsRapor.Cells(IlkSira, 33).Value & " " & WsRapor.Cells(IlkSira, 34).Value
        End If
    Else
        If InStr(WsRapor.Cells(IlkSira, 33).Value, "X.X. ") > 0 Then
            GelenTema = Mid(WsRapor.Cells(IlkSira, 33).Value, 6, Len(WsRapor.Cells(IlkSira, 33).Value))
        Else
            GelenTema = WsRapor.Cells(IlkSira, 33).Value
        End If
    End If
End If

'____________OPERASYONLAR

'İşlem günlüğünde başlangıç ve bitiş satırlarını tespit et.
Say1IslemGunlugu = WsIslemGunlugu.Range("B100000").End(xlUp).Row
Say2IslemGunlugu = WsIslemGunlugu.Range("C100000").End(xlUp).Row
SayAyracIslemGunlugu = WsIslemGunlugu.Range("E100000").End(xlUp).Row

Set IslemGunluguIlkSiraBul = WsIslemGunlugu.Range("B7:B100000").Find(What:=WsRapor.Cells(IlkSira, 93).Value, SearchDirection:=xlNext, _
                SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole) 'zaman damgasını ara
Set IslemGunluguSonSiraBul = WsIslemGunlugu.Range("C7:C100000").Find(What:=WsRapor.Cells(IlkSira, 93).Value, SearchDirection:=xlNext, _
                SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
If Not IslemGunluguIlkSiraBul Is Nothing Then
    IslemGunluguIlkSira = IslemGunluguIlkSiraBul.Row
    If Not IslemGunluguSonSiraBul Is Nothing Then
        IslemGunluguSonSira = IslemGunluguSonSiraBul.Row
    End If
End If

    
If Not IslemGunluguIlkSiraBul Is Nothing Then 'DÜZENLEME İŞLEMİ

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
    WsIslemGunlugu.Range(WsIslemGunlugu.Cells(IslemGunluguIlkSira, 2), WsIslemGunlugu.Cells(IslemGunluguSonSira, 19)).ClearContents
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
        WsIslemGunlugu.Range(WsIslemGunlugu.Cells(IslemGunluguIlkSira, 7), WsIslemGunlugu.Cells(IslemGunluguSonSira, 19)).ClearContents
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
WsIslemGunlugu.Range("E" & ilkrow & ":S" & sonrow).Interior.Color = xlNone
WsIslemGunlugu.Range("E" & ilkrow & ":S" & sonrow).Font.Color = RGB(0, 0, 0)
WsIslemGunlugu.Range("E" & ilkrow & ":S" & sonrow).Font.Bold = False
WsIslemGunlugu.Range("B" & ilkrow & ":S" & sonrow).NumberFormat = "@"
WsIslemGunlugu.Range("P" & ilkrow & ":Q" & sonrow).NumberFormat = "General"
WsIslemGunlugu.Range("F" & ilkrow & ":F" & sonrow).NumberFormat = "General"
WsIslemGunlugu.Range("D" & ilkrow & ":D" & sonrow).NumberFormat = "General"
WsIslemGunlugu.Range("B" & ilkrow & ":S" & sonrow).WrapText = True
WsIslemGunlugu.Range("F" & ilkrow & ":G" & sonrow).VerticalAlignment = xlCenter
WsIslemGunlugu.Range("P" & ilkrow & ":Q" & sonrow).VerticalAlignment = xlCenter
WsIslemGunlugu.Range("F" & ilkrow & ":G" & sonrow).HorizontalAlignment = xlCenter
WsIslemGunlugu.Range("P" & ilkrow & ":Q" & sonrow).HorizontalAlignment = xlCenter


'Zaman damgaları
WsIslemGunlugu.Cells(ilkrow, 2).Value = WsRapor.Cells(IlkSira, 93).Value
WsIslemGunlugu.Cells(sonrow, 3).Value = WsRapor.Cells(IlkSira, 93).Value

'Verileri yaz

'Rapor2_2 no
j = 0
For i = IlkSira To SonSira
    If WsRapor.Cells(i, 11).Value <> "" And Left(WsRapor.Cells(i, 63).Value, 11) = "Technique A" Then
        If j = 0 Then
            StrContent = "R/" & WsRapor.Cells(i, 11).Value
        Else
            StrContent = StrContent & ", R/" & WsRapor.Cells(i, 11).Value
        End If
        j = 1
    End If
Next i
WsIslemGunlugu.Cells(ilkrow, 7).Value = WsRapor.Cells(IlkSira, 65).Value & " (" & StrContent & ")"

'Belge bilgilerini aktar
WsIslemGunlugu.Cells(ilkrow, 8).Value = WsRapor.Cells(IlkSira, 25).Value 'İl
WsIslemGunlugu.Cells(ilkrow, 9).Value = WsRapor.Cells(IlkSira, 26).Value 'İlçe
WsIslemGunlugu.Cells(ilkrow, 10).Value = GelenTema
WsIslemGunlugu.Cells(ilkrow, 11).Value = WsRapor.Cells(IlkSira, 28).Value 'Belge tarihi
WsIslemGunlugu.Cells(ilkrow, 12).Value = WsRapor.Cells(IlkSira, 29).Value 'Belge no
WsIslemGunlugu.Cells(ilkrow, 13).Value = WsRapor.Cells(IlkSira, 36).Value 'finansal birime ulaşma tarihi
WsIslemGunlugu.Cells(ilkrow, 14).Value = WsRapor.Cells(IlkSira, 39).Value 'Tespit tarihi

'Öğe bilgilerini aktar
j = ilkrow - 1
For i = IlkSira To SonSira
    If Left(WsRapor.Cells(i, 63).Value, 11) = "Technique A" Then
        j = j + 1
        'Öğe türü
        WsIslemGunlugu.Cells(j, 15).Value = WsRapor.Cells(i, 46).Value
        'Öğe değeri
        WsIslemGunlugu.Cells(j, 16).Value = WsRapor.Cells(i, 49).Value
        'Adet
        WsIslemGunlugu.Cells(j, 17).Value = WsRapor.Cells(i, 52).Value
        'Açıklama
        WsIslemGunlugu.Cells(j, 19).Value = WsRapor.Cells(i, 58).Value
    End If
Next i

WsIslemGunlugu.Cells(ilkrow, 18).Value = WsRapor.Cells(IlkSira, 31).Value 'Tema

'Sadece Technique A içeriklik satırlar aktarılacağı için Technique A olmayan satırlar silinecek; dolayısıyla bitiş damgası yeniden yazıldı.
sonrowx = sonrow
sonrow = j
WsIslemGunlugu.Cells(sonrow, 3).Value = WsRapor.Cells(IlkSira, 93).Value

'Kenarlıklar.
Set Kenarlar = WsIslemGunlugu.Range("D" & ilkrow & ":S" & sonrow)
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

'hedef dönemdeki artık satırı sil
If sonrowx > sonrow Then
    WsIslemGunlugu.Rows(sonrowx & ":" & sonrow + 1).EntireRow.Delete
End If


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
    Workbooks("System Registry Report 2.2.xlsx").Save
    Workbooks("System Registry Report 2.2.xlsx").Close SaveChanges:=False
End If


If IlceSakla <> "" Then
    WsRapor.Cells(IlkSira, 26).Value = IlceSakla
End If


'WsRapor.Protect Password:="123"

Out:

Application.DisplayAlerts = True
Application.EnableEvents = True
Application.ScreenUpdating = True


End Sub

Sub Rapor2GelenBelgeGiris()
Dim i As Integer, TarihBul As Range
Dim Say As Long, j As Long, Cont As Long, Sno As Variant
Dim LstBx As MSForms.ListBox, ctl As MSForms.Control
Dim LblSira1 As MSForms.Label
Dim LblSira2 As MSForms.Label
Dim Contx As Long

CalTarih = CalTarihTakip

Say = ThisWorkbook.Worksheets(4).Range("CN100000").End(xlUp).Row
If Say < 7 Then
    GoTo Son
End If

'Cont = 0
Contx = 0
TarihiTekrarla1:
Set TarihBul = ThisWorkbook.Worksheets(4).Range("AM6:AM100000").Find(What:=CalTarih, SearchDirection:=xlNext, _
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
        If ThisWorkbook.Worksheets(4).Cells(j, 5).Value <> "" And ThisWorkbook.Worksheets(4).Cells(j, 39).Value <> "" Then
            Cont = Cont + 1
            If ThisWorkbook.Worksheets(4).Cells(j, 34).Value = "" Then
                If Left(ThisWorkbook.Worksheets(4).Cells(j, 33).Value, 3) = "İl " Then
                    Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | R | " & ThisWorkbook.Worksheets(4).Cells(j, 29).Value & " | " & ThisWorkbook.Worksheets(4).Cells(j, 25).Value & " " & ThisWorkbook.Worksheets(4).Cells(j, 33).Value
                ElseIf Left(ThisWorkbook.Worksheets(4).Cells(j, 33).Value, 3) = "İlç" Then
                    Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | R | " & ThisWorkbook.Worksheets(4).Cells(j, 29).Value & " | " & ThisWorkbook.Worksheets(4).Cells(j, 26).Value & " " & ThisWorkbook.Worksheets(4).Cells(j, 33).Value
                Else
                    Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | R | " & ThisWorkbook.Worksheets(4).Cells(j, 29).Value & " | " & ThisWorkbook.Worksheets(4).Cells(j, 33).Value
                End If
            Else
                Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | R | " & ThisWorkbook.Worksheets(4).Cells(j, 29).Value & " | " & ThisWorkbook.Worksheets(4).Cells(j, 34).Value
            End If

            Set LstBx = core_acceptance_manager_UI.Frame1.Controls.Add("Forms.ListBox.1", "Frame1LstBx" & Cont)
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

            Set LblSira1 = core_acceptance_manager_UI.Frame1.Controls.Add("Forms.Label.1", "Lbl" & Cont)
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


End Sub

Sub Rapor2_1TeslimTutanaklari()
Dim i As Integer, TarihBul As Range
Dim Say As Long, j As Long, Cont As Long, Sno As Variant
Dim LstBx As MSForms.ListBox, ctl As MSForms.Control
Dim LblSira1 As MSForms.Label
Dim LblSira2 As MSForms.Label
Dim Contx As Long

CalTarih = CalTarihTakip

Say = ThisWorkbook.Worksheets(4).Range("CN100000").End(xlUp).Row
If Say < 7 Then
    GoTo Son
End If

'Cont = 0
Contx = 0
TarihiTekrarla1:
Set TarihBul = ThisWorkbook.Worksheets(4).Range("CE6:CE100000").Find(What:=CalTarih, SearchDirection:=xlNext, _
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

'Bilgilendirme
For j = Say To TarihBul.Row Step -1
    'If Cont < 50 Then
        If ThisWorkbook.Worksheets(4).Cells(j, 5).Value <> "" And ThisWorkbook.Worksheets(4).Cells(j, 83).Value <> "" And ThisWorkbook.Worksheets(4).Cells(j, 174).Value = "No" Then
            Cont = Cont + 1
            If ThisWorkbook.Worksheets(4).Cells(j, 73).Value <> "" Then
                If Left(ThisWorkbook.Worksheets(4).Cells(j, 72).Value, 3) = "İl " Then
                    Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | R | " & ThisWorkbook.Worksheets(4).Cells(j, 84).Value & " | " & ThisWorkbook.Worksheets(4).Cells(j, 73).Value
                ElseIf Left(ThisWorkbook.Worksheets(4).Cells(j, 72).Value, 3) = "İlç" Then
                    Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | R | " & ThisWorkbook.Worksheets(4).Cells(j, 84).Value & " | " & ThisWorkbook.Worksheets(4).Cells(j, 73).Value
                Else
                    Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | R | " & ThisWorkbook.Worksheets(4).Cells(j, 84).Value & " | " & ThisWorkbook.Worksheets(4).Cells(j, 72).Value & " " & ThisWorkbook.Worksheets(4).Cells(j, 73).Value
                End If
            ElseIf ThisWorkbook.Worksheets(4).Cells(j, 73).Value = "" Then
                If Left(ThisWorkbook.Worksheets(4).Cells(j, 72).Value, 3) = "İl " Then
                    Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | R | " & ThisWorkbook.Worksheets(4).Cells(j, 84).Value & " | " & ThisWorkbook.Worksheets(4).Cells(j, 77).Value & " " & ThisWorkbook.Worksheets(4).Cells(j, 72).Value
                ElseIf Left(ThisWorkbook.Worksheets(4).Cells(j, 72).Value, 3) = "İlç" Then
                    Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | R | " & ThisWorkbook.Worksheets(4).Cells(j, 84).Value & " | " & ThisWorkbook.Worksheets(4).Cells(j, 78).Value & " " & ThisWorkbook.Worksheets(4).Cells(j, 72).Value
                Else
                    Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | R | " & ThisWorkbook.Worksheets(4).Cells(j, 84).Value & " | " & ThisWorkbook.Worksheets(4).Cells(j, 72).Value
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


'MsgBox "Rapor2: " & ContTakip

End Sub

Sub Rapor2_1VarlikHareketleriGiris()
Dim i As Integer, TarihBul As Range
Dim Say As Long, j As Long, Cont As Long, Sno As Variant
Dim LstBx As MSForms.ListBox, ctl As MSForms.Control
Dim LblSira1 As MSForms.Label
Dim LblSira2 As MSForms.Label
Dim Contx As Long

Dim IlkSiraBul As Range, SonSiraBul As Range, IlkSira As Long, SonSira As Long, k As Long
Dim RaporNoTakip As String, RaporNoKontrol As Integer, SiraNo As Long


CalTarih = CalTarihTakip

Say = ThisWorkbook.Worksheets(4).Range("CN100000").End(xlUp).Row
If Say < 7 Then
    GoTo Son
End If

'Cont = 0
Contx = 0
Set TarihBul = ThisWorkbook.Worksheets(4).Range("AM6:AM100000").Find(What:=CalTarih, SearchDirection:=xlNext, _
                SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole) 'Tutanak1 tarihi içinde ara
If Not TarihBul Is Nothing Then
    '
Else
    GoTo Son
End If

'MsgBox CalTarih

Cont = ContTakipGiris
'Cont = 0
For j = TarihBul.Row To Say
    If ThisWorkbook.Worksheets(4).Cells(j, 174).Value = "No" Then
        If CalTarih = ThisWorkbook.Worksheets(4).Cells(j, 39).Value And CDate(ThisWorkbook.Worksheets(4).Cells(j, 39).Value) <> CDate(ThisWorkbook.Worksheets(4).Cells(j, 104).Value) Then
            If ThisWorkbook.Worksheets(4).Cells(j, 5).Value <> "" And ThisWorkbook.Worksheets(4).Cells(j, 11).Value <> "" Then
                Cont = Cont + 1

                Set IlkSiraBul = Nothing
                Set SonSiraBul = Nothing
                IlkSira = 0
                SonSira = 0
                SiraNo = ThisWorkbook.Worksheets(4).Cells(j, 5).Value
                Set IlkSiraBul = ThisWorkbook.Worksheets(4).Range("CM7:CM100000").Find(What:=SiraNo, SearchDirection:=xlNext, _
                                SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
                If Not IlkSiraBul Is Nothing Then
                    IlkSira = IlkSiraBul.Row
                End If
                Set SonSiraBul = ThisWorkbook.Worksheets(4).Range("CN7:CN100000").Find(What:=SiraNo, SearchDirection:=xlNext, _
                                SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
                If Not SonSiraBul Is Nothing Then
                    SonSira = SonSiraBul.Row
                End If
                RaporNoKontrol = 0
                For k = IlkSira To SonSira
                    If ThisWorkbook.Worksheets(4).Cells(k, 11).Value <> "" Then
                        If RaporNoKontrol = 0 Then
                            RaporNoTakip = ThisWorkbook.Worksheets(4).Cells(k, 11).Value
                            RaporNoKontrol = 1
                        Else
                            RaporNoTakip = RaporNoTakip & ", " & ThisWorkbook.Worksheets(4).Cells(k, 11).Value
                        End If
                    End If
                Next k
                    
                If ThisWorkbook.Worksheets(4).Cells(j, 73).Value <> "" Then
                    If Left(ThisWorkbook.Worksheets(4).Cells(j, 72).Value, 3) = "İl " Then
                        Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | R | " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(4).Cells(j, 73).Value
                    ElseIf Left(ThisWorkbook.Worksheets(4).Cells(j, 72).Value, 3) = "İlç" Then
                        Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | R | " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(4).Cells(j, 73).Value
                    Else
                        Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | R | " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(4).Cells(j, 72).Value & " " & ThisWorkbook.Worksheets(4).Cells(j, 73).Value
                    End If
                ElseIf ThisWorkbook.Worksheets(4).Cells(j, 73).Value = "" Then
                    If Left(ThisWorkbook.Worksheets(4).Cells(j, 72).Value, 3) = "İl " Then
                        Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | R | " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(4).Cells(j, 77).Value & " " & ThisWorkbook.Worksheets(4).Cells(j, 72).Value
                    ElseIf Left(ThisWorkbook.Worksheets(4).Cells(j, 72).Value, 3) = "İlç" Then
                        Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | R | " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(4).Cells(j, 78).Value & " " & ThisWorkbook.Worksheets(4).Cells(j, 72).Value
                    Else
                        Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | R | " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(4).Cells(j, 72).Value
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
Next j

ContTakipGiris = Cont

Son:

End Sub

Sub Rapor2_1VarlikHareketleriMevcut()
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

Say = ThisWorkbook.Worksheets(4).Range("CN100000").End(xlUp).Row
If Say < 7 Then
    GoTo Son
End If

'Mevcutta satırı bul
'CalTarih = CDate(CalTarih)
MevcutSatir = 0
For j = Say To 7 Step -1
    If ThisWorkbook.Worksheets(4).Cells(j, 174).Value = "No" Then
        If ThisWorkbook.Worksheets(4).Range("AM" & j).Value <> "" Then
            TarihAra = ThisWorkbook.Worksheets(4).Range("AM" & j)
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
    If ThisWorkbook.Worksheets(4).Cells(j, 174).Value = "No" Then
        If ThisWorkbook.Worksheets(4).Cells(j, 39).Value <> "" Then
            TarihAra = ThisWorkbook.Worksheets(4).Range("AM" & j)
            If CDate(TarihAra) < CDate(CalTarih) Then
                If (ThisWorkbook.Worksheets(4).Cells(j, 104).Value <> "" And CDate(ThisWorkbook.Worksheets(4).Cells(j, 104).Value) > CDate(CalTarih)) Or _
                    ThisWorkbook.Worksheets(4).Cells(j, 104).Value = "" Then
                    'MsgBox TarihAra & " < " & CalTarih
                    If ThisWorkbook.Worksheets(4).Cells(j, 5).Value <> "" And ThisWorkbook.Worksheets(4).Cells(j, 11).Value <> "" Then
                        Cont = Cont + 1

                        Set IlkSiraBul = Nothing
                        Set SonSiraBul = Nothing
                        IlkSira = 0
                        SonSira = 0
                        SiraNo = ThisWorkbook.Worksheets(4).Cells(j, 5).Value
                        Set IlkSiraBul = ThisWorkbook.Worksheets(4).Range("CM7:CM100000").Find(What:=SiraNo, SearchDirection:=xlNext, _
                                        SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
                        If Not IlkSiraBul Is Nothing Then
                            IlkSira = IlkSiraBul.Row
                        End If
                        Set SonSiraBul = ThisWorkbook.Worksheets(4).Range("CN7:CN100000").Find(What:=SiraNo, SearchDirection:=xlNext, _
                                        SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
                        If Not SonSiraBul Is Nothing Then
                            SonSira = SonSiraBul.Row
                        End If
                        RaporNoKontrol = 0
                        For k = IlkSira To SonSira
                            If ThisWorkbook.Worksheets(4).Cells(k, 11).Value <> "" Then
                                If RaporNoKontrol = 0 Then
                                    RaporNoTakip = ThisWorkbook.Worksheets(4).Cells(k, 11).Value
                                    RaporNoKontrol = 1
                                Else
                                    RaporNoTakip = RaporNoTakip & ", " & ThisWorkbook.Worksheets(4).Cells(k, 11).Value
                                End If
                            End If
                        Next k
                
                        If ThisWorkbook.Worksheets(4).Cells(j, 73).Value <> "" Then
                            If Left(ThisWorkbook.Worksheets(4).Cells(j, 72).Value, 3) = "İl " Then
                                Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | R | " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(4).Cells(j, 73).Value
                            ElseIf Left(ThisWorkbook.Worksheets(4).Cells(j, 72).Value, 3) = "İlç" Then
                                Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | R | " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(4).Cells(j, 73).Value
                            Else
                                Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | R | " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(4).Cells(j, 72).Value & " " & ThisWorkbook.Worksheets(4).Cells(j, 73).Value
                            End If
                        ElseIf ThisWorkbook.Worksheets(4).Cells(j, 73).Value = "" Then
                            If Left(ThisWorkbook.Worksheets(4).Cells(j, 72).Value, 3) = "İl " Then
                                Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | R | " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(4).Cells(j, 77).Value & " " & ThisWorkbook.Worksheets(4).Cells(j, 72).Value
                            ElseIf Left(ThisWorkbook.Worksheets(4).Cells(j, 72).Value, 3) = "İlç" Then
                                Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | R | " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(4).Cells(j, 78).Value & " " & ThisWorkbook.Worksheets(4).Cells(j, 72).Value
                            Else
                                Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | R | " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(4).Cells(j, 72).Value
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
Next j

ContTakipMevcut = Cont

Son:

End Sub

Sub Rapor2_1VarlikHareketleriCikis()
Dim i As Integer, TarihBul As Range
Dim Say As Long, j As Long, Cont As Long, Sno As Variant
Dim LstBx As MSForms.ListBox, ctl As MSForms.Control
Dim LblSira1 As MSForms.Label
Dim LblSira2 As MSForms.Label
Dim Contx As Long

Dim IlkSiraBul As Range, SonSiraBul As Range, IlkSira As Long, SonSira As Long, k As Long
Dim RaporNoTakip As String, RaporNoKontrol As Integer, SiraNo As Long


CalTarih = CalTarihTakip

Say = ThisWorkbook.Worksheets(4).Range("CN100000").End(xlUp).Row
If Say < 7 Then
    GoTo Son
End If

'Cont = 0
Contx = 0
Set TarihBul = ThisWorkbook.Worksheets(4).Range("CZ6:CZ100000").Find(What:=CalTarih, SearchDirection:=xlNext, _
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
    If ThisWorkbook.Worksheets(4).Cells(j, 174).Value = "No" Then
        If CalTarih = ThisWorkbook.Worksheets(4).Cells(j, 104).Value And ThisWorkbook.Worksheets(4).Cells(j, 39).Value <> "" Then
            If ThisWorkbook.Worksheets(4).Cells(j, 5).Value <> "" And ThisWorkbook.Worksheets(4).Cells(j, 11).Value <> "" Then
                Cont = Cont + 1

                Set IlkSiraBul = Nothing
                Set SonSiraBul = Nothing
                IlkSira = 0
                SonSira = 0
                SiraNo = ThisWorkbook.Worksheets(4).Cells(j, 5).Value
                Set IlkSiraBul = ThisWorkbook.Worksheets(4).Range("CM7:CM100000").Find(What:=SiraNo, SearchDirection:=xlNext, _
                                SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
                If Not IlkSiraBul Is Nothing Then
                    IlkSira = IlkSiraBul.Row
                End If
                Set SonSiraBul = ThisWorkbook.Worksheets(4).Range("CN7:CN100000").Find(What:=SiraNo, SearchDirection:=xlNext, _
                                SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
                If Not SonSiraBul Is Nothing Then
                    SonSira = SonSiraBul.Row
                End If
                RaporNoKontrol = 0
                For k = IlkSira To SonSira
                    If ThisWorkbook.Worksheets(4).Cells(k, 11).Value <> "" Then
                        If RaporNoKontrol = 0 Then
                            RaporNoTakip = ThisWorkbook.Worksheets(4).Cells(k, 11).Value
                            RaporNoKontrol = 1
                        Else
                            RaporNoTakip = RaporNoTakip & ", " & ThisWorkbook.Worksheets(4).Cells(k, 11).Value
                        End If
                    End If
                Next k

                If ThisWorkbook.Worksheets(4).Cells(j, 73).Value <> "" Then
                    If Left(ThisWorkbook.Worksheets(4).Cells(j, 72).Value, 3) = "İl " Then
                        Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | R | " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(4).Cells(j, 73).Value
                    ElseIf Left(ThisWorkbook.Worksheets(4).Cells(j, 72).Value, 3) = "İlç" Then
                        Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | R | " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(4).Cells(j, 73).Value
                    Else
                        Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | R | " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(4).Cells(j, 72).Value & " " & ThisWorkbook.Worksheets(4).Cells(j, 73).Value
                    End If
                ElseIf ThisWorkbook.Worksheets(4).Cells(j, 73).Value = "" Then
                    If Left(ThisWorkbook.Worksheets(4).Cells(j, 72).Value, 3) = "İl " Then
                        Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | R | " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(4).Cells(j, 77).Value & " " & ThisWorkbook.Worksheets(4).Cells(j, 72).Value
                    ElseIf Left(ThisWorkbook.Worksheets(4).Cells(j, 72).Value, 3) = "İlç" Then
                        Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | R | " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(4).Cells(j, 78).Value & " " & ThisWorkbook.Worksheets(4).Cells(j, 72).Value
                    Else
                        Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | R | " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(4).Cells(j, 72).Value
                    End If
                End If
  
                Set LstBx = core_asset_manager_UI.FrameCikis.Controls.Add("Forms.ListBox.1", "Frame1LstBx" & Cont)
                With LstBx
                    .Top = (Cont - 1) * 12
                    .Left = 18
                    .Height = 12
                    .Width = 300
                    If ThisWorkbook.Worksheets(4).Cells(j, 39).Value = ThisWorkbook.Worksheets(4).Cells(j, 104).Value Then
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
Next j

ContTakipCikis = Cont

Son:

End Sub

Sub Rapor2_2VarlikHareketleriGiris()
Dim i As Integer, TarihBul As Range
Dim Say As Long, j As Long, Cont As Long, Sno As Variant
Dim LstBx As MSForms.ListBox, ctl As MSForms.Control
Dim LblSira1 As MSForms.Label
Dim LblSira2 As MSForms.Label
Dim Contx As Long

Dim IlkSiraBul As Range, SonSiraBul As Range, IlkSira As Long, SonSira As Long, k As Long
Dim RaporNoTakip As String, RaporNoKontrol As Integer, SiraNo As Long


CalTarih = CalTarihTakip

Say = ThisWorkbook.Worksheets(4).Range("CN100000").End(xlUp).Row
If Say < 7 Then
    GoTo Son
End If

'Cont = 0
Contx = 0
Set TarihBul = ThisWorkbook.Worksheets(4).Range("AM6:AM100000").Find(What:=CalTarih, SearchDirection:=xlNext, _
                SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole) 'Tutanak1 tarihi içinde ara
If Not TarihBul Is Nothing Then
    '
Else
    GoTo Son
End If

'MsgBox CalTarih
    
Cont = ContTakipGiris
'Cont = 0
For j = TarihBul.Row To Say
    If ThisWorkbook.Worksheets(4).Cells(j, 174).Value = "Yes" Then
        If ThisWorkbook.Worksheets(4).Cells(j, 187).Value = "" Then
            If CalTarih = ThisWorkbook.Worksheets(4).Cells(j, 39).Value And CDate(ThisWorkbook.Worksheets(4).Cells(j, 39).Value) <> CDate(ThisWorkbook.Worksheets(4).Cells(j, 105).Value) Then
                If ThisWorkbook.Worksheets(4).Cells(j, 5).Value <> "" And ThisWorkbook.Worksheets(4).Cells(j, 11).Value <> "" Then
                    Cont = Cont + 1
        
                    Set IlkSiraBul = Nothing
                    Set SonSiraBul = Nothing
                    IlkSira = 0
                    SonSira = 0
                    SiraNo = ThisWorkbook.Worksheets(4).Cells(j, 5).Value
                    Set IlkSiraBul = ThisWorkbook.Worksheets(4).Range("CM7:CM100000").Find(What:=SiraNo, SearchDirection:=xlNext, _
                                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
                    If Not IlkSiraBul Is Nothing Then
                        IlkSira = IlkSiraBul.Row
                    End If
                    Set SonSiraBul = ThisWorkbook.Worksheets(4).Range("CN7:CN100000").Find(What:=SiraNo, SearchDirection:=xlNext, _
                                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
                    If Not SonSiraBul Is Nothing Then
                        SonSira = SonSiraBul.Row
                    End If
                    RaporNoKontrol = 0
                    For k = IlkSira To SonSira
                        If ThisWorkbook.Worksheets(4).Cells(k, 11).Value <> "" Then
                            If RaporNoKontrol = 0 Then
                                RaporNoTakip = ThisWorkbook.Worksheets(4).Cells(k, 11).Value
                                RaporNoKontrol = 1
                            Else
                                RaporNoTakip = RaporNoTakip & ", " & ThisWorkbook.Worksheets(4).Cells(k, 11).Value
                            End If
                        End If
                    Next k
                    
                    If ThisWorkbook.Worksheets(4).Cells(j, 34).Value <> "" Then
                        If Left(ThisWorkbook.Worksheets(4).Cells(j, 33).Value, 3) = "İl " Then
                            Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | B | " & "<T> " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(4).Cells(j, 34).Value
                        ElseIf Left(ThisWorkbook.Worksheets(4).Cells(j, 33).Value, 3) = "İlç" Then
                            Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | B | " & "<T> " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(4).Cells(j, 34).Value
                        Else
                            Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | B | " & "<T> " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(4).Cells(j, 33).Value & " " & ThisWorkbook.Worksheets(4).Cells(j, 34).Value
                        End If
                    ElseIf ThisWorkbook.Worksheets(4).Cells(j, 34).Value = "" Then
                        If Left(ThisWorkbook.Worksheets(4).Cells(j, 33).Value, 3) = "İl " Then
                            Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | B | " & "<T> " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(4).Cells(j, 25).Value & " " & ThisWorkbook.Worksheets(4).Cells(j, 33).Value
                        ElseIf Left(ThisWorkbook.Worksheets(4).Cells(j, 33).Value, 3) = "İlç" Then
                            Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | B | " & "<T> " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(4).Cells(j, 26).Value & " " & ThisWorkbook.Worksheets(4).Cells(j, 33).Value
                        Else
                            Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | B | " & "<T> " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(4).Cells(j, 33).Value
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

        ElseIf ThisWorkbook.Worksheets(4).Cells(j, 187).Value = "All" Then
            If CalTarih = ThisWorkbook.Worksheets(4).Cells(j, 39).Value And CDate(ThisWorkbook.Worksheets(4).Cells(j, 39).Value) <> CDate(ThisWorkbook.Worksheets(4).Cells(j, 105).Value) Then
                If ThisWorkbook.Worksheets(4).Cells(j, 5).Value <> "" And ThisWorkbook.Worksheets(4).Cells(j, 11).Value <> "" Then
                    Cont = Cont + 1

                    Set IlkSiraBul = Nothing
                    Set SonSiraBul = Nothing
                    IlkSira = 0
                    SonSira = 0
                    SiraNo = ThisWorkbook.Worksheets(4).Cells(j, 5).Value
                    Set IlkSiraBul = ThisWorkbook.Worksheets(4).Range("CM7:CM100000").Find(What:=SiraNo, SearchDirection:=xlNext, _
                                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
                    If Not IlkSiraBul Is Nothing Then
                        IlkSira = IlkSiraBul.Row
                    End If
                    Set SonSiraBul = ThisWorkbook.Worksheets(4).Range("CN7:CN100000").Find(What:=SiraNo, SearchDirection:=xlNext, _
                                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
                    If Not SonSiraBul Is Nothing Then
                        SonSira = SonSiraBul.Row
                    End If
                    RaporNoKontrol = 0
                    For k = IlkSira To SonSira
                        If ThisWorkbook.Worksheets(4).Cells(k, 11).Value <> "" Then
                            If RaporNoKontrol = 0 Then
                                RaporNoTakip = ThisWorkbook.Worksheets(4).Cells(k, 11).Value
                                RaporNoKontrol = 1
                            Else
                                RaporNoTakip = RaporNoTakip & ", " & ThisWorkbook.Worksheets(4).Cells(k, 11).Value
                            End If
                        End If
                    Next k
                    
                    If ThisWorkbook.Worksheets(4).Cells(j, 34).Value <> "" Then
                        If Left(ThisWorkbook.Worksheets(4).Cells(j, 33).Value, 3) = "İl " Then
                            Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | B | " & "<T> (Giden) " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(4).Cells(j, 34).Value
                        ElseIf Left(ThisWorkbook.Worksheets(4).Cells(j, 33).Value, 3) = "İlç" Then
                            Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | B | " & "<T> (Giden) " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(4).Cells(j, 34).Value
                        Else
                            Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | B | " & "<T> (Giden) " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(4).Cells(j, 33).Value & " " & ThisWorkbook.Worksheets(4).Cells(j, 34).Value
                        End If
                    ElseIf ThisWorkbook.Worksheets(4).Cells(j, 34).Value = "" Then
                        If Left(ThisWorkbook.Worksheets(4).Cells(j, 33).Value, 3) = "İl " Then
                            Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | B | " & "<T> (Giden) " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(4).Cells(j, 25).Value & " " & ThisWorkbook.Worksheets(4).Cells(j, 33).Value
                        ElseIf Left(ThisWorkbook.Worksheets(4).Cells(j, 33).Value, 3) = "İlç" Then
                            Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | B | " & "<T> (Giden) " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(4).Cells(j, 26).Value & " " & ThisWorkbook.Worksheets(4).Cells(j, 33).Value
                        Else
                            Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | B | " & "<T> (Giden) " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(4).Cells(j, 33).Value
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
            
        ElseIf ThisWorkbook.Worksheets(4).Cells(j, 187).Value = "Technique A" Then
            'If CalTarih = ThisWorkbook.Worksheets(4).Cells(j, 39).Value And CDate(ThisWorkbook.Worksheets(4).Cells(j, 39).Value) <> CDate(ThisWorkbook.Worksheets(4).Cells(j, 105).Value) Then
                If ThisWorkbook.Worksheets(4).Cells(j, 5).Value <> "" And ThisWorkbook.Worksheets(4).Cells(j, 11).Value <> "" Then
                    For i = 1 To 2
                        'Cont = Cont + 1
                        If i = 1 Then
                            If CalTarih = ThisWorkbook.Worksheets(4).Cells(j, 39).Value And CDate(ThisWorkbook.Worksheets(4).Cells(j, 39).Value) <> CDate(ThisWorkbook.Worksheets(4).Cells(j, 104).Value) Then
                                Cont = Cont + 1

                                Set IlkSiraBul = Nothing
                                Set SonSiraBul = Nothing
                                IlkSira = 0
                                SonSira = 0
                                SiraNo = ThisWorkbook.Worksheets(4).Cells(j, 5).Value
                                Set IlkSiraBul = ThisWorkbook.Worksheets(4).Range("CM7:CM100000").Find(What:=SiraNo, SearchDirection:=xlNext, _
                                                SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
                                If Not IlkSiraBul Is Nothing Then
                                    IlkSira = IlkSiraBul.Row
                                End If
                                Set SonSiraBul = ThisWorkbook.Worksheets(4).Range("CN7:CN100000").Find(What:=SiraNo, SearchDirection:=xlNext, _
                                                SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
                                If Not SonSiraBul Is Nothing Then
                                    SonSira = SonSiraBul.Row
                                End If
                                RaporNoKontrol = 0
                                For k = IlkSira To SonSira
                                    If Left(ThisWorkbook.Worksheets(4).Cells(k, 63).Value, 8) <> "Technique A" Then
                                        If ThisWorkbook.Worksheets(4).Cells(k, 11).Value <> "" Then
                                            If RaporNoKontrol = 0 Then
                                                RaporNoTakip = ThisWorkbook.Worksheets(4).Cells(k, 11).Value
                                                RaporNoKontrol = 1
                                            Else
                                                RaporNoTakip = RaporNoTakip & ", " & ThisWorkbook.Worksheets(4).Cells(k, 11).Value
                                            End If
                                        End If
                                    End If
                                Next k

                                If ThisWorkbook.Worksheets(4).Cells(j, 34).Value <> "" Then
                                    If Left(ThisWorkbook.Worksheets(4).Cells(j, 33).Value, 3) = "İl " Then
                                        Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | B | " & "<R> " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(4).Cells(j, 34).Value
                                    ElseIf Left(ThisWorkbook.Worksheets(4).Cells(j, 33).Value, 3) = "İlç" Then
                                        Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | B | " & "<R> " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(4).Cells(j, 34).Value
                                    Else
                                        Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | B | " & "<R> " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(4).Cells(j, 33).Value & " " & ThisWorkbook.Worksheets(4).Cells(j, 34).Value
                                    End If
                                ElseIf ThisWorkbook.Worksheets(4).Cells(j, 34).Value = "" Then
                                    If Left(ThisWorkbook.Worksheets(4).Cells(j, 33).Value, 3) = "İl " Then
                                        Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | B | " & "<R> " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(4).Cells(j, 25).Value & " " & ThisWorkbook.Worksheets(4).Cells(j, 33).Value
                                    ElseIf Left(ThisWorkbook.Worksheets(4).Cells(j, 33).Value, 3) = "İlç" Then
                                        Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | B | " & "<R> " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(4).Cells(j, 26).Value & " " & ThisWorkbook.Worksheets(4).Cells(j, 33).Value
                                    Else
                                        Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | B | " & "<R> " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(4).Cells(j, 33).Value
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

                        ElseIf i = 2 Then
                            If CalTarih = ThisWorkbook.Worksheets(4).Cells(j, 39).Value And CDate(ThisWorkbook.Worksheets(4).Cells(j, 39).Value) <> CDate(ThisWorkbook.Worksheets(4).Cells(j, 105).Value) Then
                                Cont = Cont + 1

                                Set IlkSiraBul = Nothing
                                Set SonSiraBul = Nothing
                                IlkSira = 0
                                SonSira = 0
                                SiraNo = ThisWorkbook.Worksheets(4).Cells(j, 5).Value
                                Set IlkSiraBul = ThisWorkbook.Worksheets(4).Range("CM7:CM100000").Find(What:=SiraNo, SearchDirection:=xlNext, _
                                                SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
                                If Not IlkSiraBul Is Nothing Then
                                    IlkSira = IlkSiraBul.Row
                                End If
                                Set SonSiraBul = ThisWorkbook.Worksheets(4).Range("CN7:CN100000").Find(What:=SiraNo, SearchDirection:=xlNext, _
                                                SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
                                If Not SonSiraBul Is Nothing Then
                                    SonSira = SonSiraBul.Row
                                End If
                                RaporNoKontrol = 0
                                For k = IlkSira To SonSira
                                    If Left(ThisWorkbook.Worksheets(4).Cells(k, 63).Value, 11) = "Technique A" Then
                                        If ThisWorkbook.Worksheets(4).Cells(k, 11).Value <> "" Then
                                            If RaporNoKontrol = 0 Then
                                                RaporNoTakip = ThisWorkbook.Worksheets(4).Cells(k, 11).Value
                                                RaporNoKontrol = 1
                                            Else
                                                RaporNoTakip = RaporNoTakip & ", " & ThisWorkbook.Worksheets(4).Cells(k, 11).Value
                                            End If
                                        End If
                                    End If
                                Next k
                            
                                If ThisWorkbook.Worksheets(4).Cells(j, 34).Value <> "" Then
                                    If Left(ThisWorkbook.Worksheets(4).Cells(j, 33).Value, 3) = "İl " Then
                                        Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | B | " & "<L> (Giden) " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(4).Cells(j, 34).Value
                                    ElseIf Left(ThisWorkbook.Worksheets(4).Cells(j, 33).Value, 3) = "İlç" Then
                                        Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | B | " & "<L> (Giden) " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(4).Cells(j, 34).Value
                                    Else
                                        Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | B | " & "<L> (Giden) " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(4).Cells(j, 33).Value & " " & ThisWorkbook.Worksheets(4).Cells(j, 34).Value
                                    End If
                                ElseIf ThisWorkbook.Worksheets(4).Cells(j, 34).Value = "" Then
                                    If Left(ThisWorkbook.Worksheets(4).Cells(j, 33).Value, 3) = "İl " Then
                                        Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | B | " & "<L> (Giden) " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(4).Cells(j, 25).Value & " " & ThisWorkbook.Worksheets(4).Cells(j, 33).Value
                                    ElseIf Left(ThisWorkbook.Worksheets(4).Cells(j, 33).Value, 3) = "İlç" Then
                                        Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | B | " & "<L> (Giden) " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(4).Cells(j, 26).Value & " " & ThisWorkbook.Worksheets(4).Cells(j, 33).Value
                                    Else
                                        Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | B | " & "<L> (Giden) " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(4).Cells(j, 33).Value
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
                    Next i
                End If
            'End If
        End If
    End If
Next j

ContTakipGiris = Cont

Son:

End Sub

Sub Rapor2_2VarlikHareketleriMevcut()
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

Say = ThisWorkbook.Worksheets(4).Range("CN100000").End(xlUp).Row
If Say < 7 Then
    GoTo Son
End If

'Mevcutta satırı bul
'CalTarih = CDate(CalTarih)
MevcutSatir = 0
For j = Say To 7 Step -1
    If ThisWorkbook.Worksheets(4).Cells(j, 174).Value = "Yes" Then
        If ThisWorkbook.Worksheets(4).Range("AM" & j).Value <> "" Then
            TarihAra = ThisWorkbook.Worksheets(4).Range("AM" & j)
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
    If ThisWorkbook.Worksheets(4).Cells(j, 174).Value = "Yes" Then
        If ThisWorkbook.Worksheets(4).Cells(j, 187).Value = "" Then
            If ThisWorkbook.Worksheets(4).Cells(j, 39).Value <> "" Then
                TarihAra = ThisWorkbook.Worksheets(4).Range("AM" & j)
                If CDate(TarihAra) < CDate(CalTarih) Then
                    If (ThisWorkbook.Worksheets(4).Cells(j, 105).Value <> "" And CDate(ThisWorkbook.Worksheets(4).Cells(j, 105).Value) > CDate(CalTarih)) Or _
                        ThisWorkbook.Worksheets(4).Cells(j, 105).Value = "" Then
                        'MsgBox TarihAra & " < " & CalTarih
                        If ThisWorkbook.Worksheets(4).Cells(j, 5).Value <> "" And ThisWorkbook.Worksheets(4).Cells(j, 11).Value <> "" Then
                            Cont = Cont + 1

                            Set IlkSiraBul = Nothing
                            Set SonSiraBul = Nothing
                            IlkSira = 0
                            SonSira = 0
                            SiraNo = ThisWorkbook.Worksheets(4).Cells(j, 5).Value
                            Set IlkSiraBul = ThisWorkbook.Worksheets(4).Range("CM7:CM100000").Find(What:=SiraNo, SearchDirection:=xlNext, _
                                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
                            If Not IlkSiraBul Is Nothing Then
                                IlkSira = IlkSiraBul.Row
                            End If
                            Set SonSiraBul = ThisWorkbook.Worksheets(4).Range("CN7:CN100000").Find(What:=SiraNo, SearchDirection:=xlNext, _
                                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
                            If Not SonSiraBul Is Nothing Then
                                SonSira = SonSiraBul.Row
                            End If
                            RaporNoKontrol = 0
                            For k = IlkSira To SonSira
                                If ThisWorkbook.Worksheets(4).Cells(k, 11).Value <> "" Then
                                    If RaporNoKontrol = 0 Then
                                        RaporNoTakip = ThisWorkbook.Worksheets(4).Cells(k, 11).Value
                                        RaporNoKontrol = 1
                                    Else
                                        RaporNoTakip = RaporNoTakip & ", " & ThisWorkbook.Worksheets(4).Cells(k, 11).Value
                                    End If
                                End If
                            Next k
                    
                            If ThisWorkbook.Worksheets(4).Cells(j, 34).Value <> "" Then
                                If Left(ThisWorkbook.Worksheets(4).Cells(j, 33).Value, 3) = "İl " Then
                                    Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | B | " & "<T> " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(4).Cells(j, 34).Value
                                ElseIf Left(ThisWorkbook.Worksheets(4).Cells(j, 33).Value, 3) = "İlç" Then
                                    Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | B | " & "<T> " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(4).Cells(j, 34).Value
                                Else
                                    Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | B | " & "<T> " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(4).Cells(j, 33).Value & " " & ThisWorkbook.Worksheets(4).Cells(j, 34).Value
                                End If
                            ElseIf ThisWorkbook.Worksheets(4).Cells(j, 34).Value = "" Then
                                If Left(ThisWorkbook.Worksheets(4).Cells(j, 33).Value, 3) = "İl " Then
                                    Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | B | " & "<T> " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(4).Cells(j, 25).Value & " " & ThisWorkbook.Worksheets(4).Cells(j, 33).Value
                                ElseIf Left(ThisWorkbook.Worksheets(4).Cells(j, 33).Value, 3) = "İlç" Then
                                    Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | B | " & "<T> " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(4).Cells(j, 26).Value & " " & ThisWorkbook.Worksheets(4).Cells(j, 33).Value
                                Else
                                    Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | B | " & "<T> " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(4).Cells(j, 33).Value
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

        ElseIf ThisWorkbook.Worksheets(4).Cells(j, 187).Value = "All" Then
            If ThisWorkbook.Worksheets(4).Cells(j, 39).Value <> "" Then
                TarihAra = ThisWorkbook.Worksheets(4).Range("AM" & j)
                If CDate(TarihAra) < CDate(CalTarih) Then
                    If (ThisWorkbook.Worksheets(4).Cells(j, 105).Value <> "" And CDate(ThisWorkbook.Worksheets(4).Cells(j, 105).Value) > CDate(CalTarih)) Or _
                        ThisWorkbook.Worksheets(4).Cells(j, 105).Value = "" Then
                        'MsgBox TarihAra & " < " & CalTarih
                        If ThisWorkbook.Worksheets(4).Cells(j, 5).Value <> "" And ThisWorkbook.Worksheets(4).Cells(j, 11).Value <> "" Then
                            Cont = Cont + 1

                            Set IlkSiraBul = Nothing
                            Set SonSiraBul = Nothing
                            IlkSira = 0
                            SonSira = 0
                            SiraNo = ThisWorkbook.Worksheets(4).Cells(j, 5).Value
                            Set IlkSiraBul = ThisWorkbook.Worksheets(4).Range("CM7:CM100000").Find(What:=SiraNo, SearchDirection:=xlNext, _
                                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
                            If Not IlkSiraBul Is Nothing Then
                                IlkSira = IlkSiraBul.Row
                            End If
                            Set SonSiraBul = ThisWorkbook.Worksheets(4).Range("CN7:CN100000").Find(What:=SiraNo, SearchDirection:=xlNext, _
                                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
                            If Not SonSiraBul Is Nothing Then
                                SonSira = SonSiraBul.Row
                            End If
                            RaporNoKontrol = 0
                            For k = IlkSira To SonSira
                                If ThisWorkbook.Worksheets(4).Cells(k, 11).Value <> "" Then
                                    If RaporNoKontrol = 0 Then
                                        RaporNoTakip = ThisWorkbook.Worksheets(4).Cells(k, 11).Value
                                        RaporNoKontrol = 1
                                    Else
                                        RaporNoTakip = RaporNoTakip & ", " & ThisWorkbook.Worksheets(4).Cells(k, 11).Value
                                    End If
                                End If
                            Next k
                    
                            If ThisWorkbook.Worksheets(4).Cells(j, 34).Value <> "" Then
                                If Left(ThisWorkbook.Worksheets(4).Cells(j, 33).Value, 3) = "İl " Then
                                    Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | B | " & "<T> (Giden) " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(4).Cells(j, 34).Value
                                ElseIf Left(ThisWorkbook.Worksheets(4).Cells(j, 33).Value, 3) = "İlç" Then
                                    Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | B | " & "<T> (Giden) " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(4).Cells(j, 34).Value
                                Else
                                    Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | B | " & "<T> (Giden) " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(4).Cells(j, 33).Value & " " & ThisWorkbook.Worksheets(4).Cells(j, 34).Value
                                End If
                            ElseIf ThisWorkbook.Worksheets(4).Cells(j, 34).Value = "" Then
                                If Left(ThisWorkbook.Worksheets(4).Cells(j, 33).Value, 3) = "İl " Then
                                    Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | B | " & "<T> (Giden) " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(4).Cells(j, 25).Value & " " & ThisWorkbook.Worksheets(4).Cells(j, 33).Value
                                ElseIf Left(ThisWorkbook.Worksheets(4).Cells(j, 33).Value, 3) = "İlç" Then
                                    Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | B | " & "<T> (Giden) " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(4).Cells(j, 26).Value & " " & ThisWorkbook.Worksheets(4).Cells(j, 33).Value
                                Else
                                    Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | B | " & "<T> (Giden) " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(4).Cells(j, 33).Value
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
            
        ElseIf ThisWorkbook.Worksheets(4).Cells(j, 187).Value = "Technique A" Then
            If ThisWorkbook.Worksheets(4).Cells(j, 39).Value <> "" Then
                TarihAra = ThisWorkbook.Worksheets(4).Range("AM" & j)
                If CDate(TarihAra) < CDate(CalTarih) Then
'                    If (ThisWorkbook.Worksheets(4).Cells(j, 105).Value <> "" And CDate(ThisWorkbook.Worksheets(4).Cells(j, 105).Value) > CDate(CalTarih)) Or _
'                        ThisWorkbook.Worksheets(4).Cells(j, 105).Value = "" Then
                        'MsgBox TarihAra & " < " & CalTarih
                        If ThisWorkbook.Worksheets(4).Cells(j, 5).Value <> "" And ThisWorkbook.Worksheets(4).Cells(j, 11).Value <> "" Then
                            For i = 1 To 2
                                'Cont = Cont + 1
                                If i = 1 Then
                                
                                    If (ThisWorkbook.Worksheets(4).Cells(j, 104).Value <> "" And CDate(ThisWorkbook.Worksheets(4).Cells(j, 104).Value) > CDate(CalTarih)) Or _
                                        ThisWorkbook.Worksheets(4).Cells(j, 104).Value = "" Then
                                        Cont = Cont + 1

                                        Set IlkSiraBul = Nothing
                                        Set SonSiraBul = Nothing
                                        IlkSira = 0
                                        SonSira = 0
                                        SiraNo = ThisWorkbook.Worksheets(4).Cells(j, 5).Value
                                        Set IlkSiraBul = ThisWorkbook.Worksheets(4).Range("CM7:CM100000").Find(What:=SiraNo, SearchDirection:=xlNext, _
                                                        SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
                                        If Not IlkSiraBul Is Nothing Then
                                            IlkSira = IlkSiraBul.Row
                                        End If
                                        Set SonSiraBul = ThisWorkbook.Worksheets(4).Range("CN7:CN100000").Find(What:=SiraNo, SearchDirection:=xlNext, _
                                                        SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
                                        If Not SonSiraBul Is Nothing Then
                                            SonSira = SonSiraBul.Row
                                        End If
                                        RaporNoKontrol = 0
                                        For k = IlkSira To SonSira
                                            If Left(ThisWorkbook.Worksheets(4).Cells(k, 63).Value, 8) <> "Technique A" Then
                                                If ThisWorkbook.Worksheets(4).Cells(k, 11).Value <> "" Then
                                                    If RaporNoKontrol = 0 Then
                                                        RaporNoTakip = ThisWorkbook.Worksheets(4).Cells(k, 11).Value
                                                        RaporNoKontrol = 1
                                                    Else
                                                        RaporNoTakip = RaporNoTakip & ", " & ThisWorkbook.Worksheets(4).Cells(k, 11).Value
                                                    End If
                                                End If
                                            End If
                                        Next k
                                
                                        If ThisWorkbook.Worksheets(4).Cells(j, 34).Value <> "" Then
                                            If Left(ThisWorkbook.Worksheets(4).Cells(j, 33).Value, 3) = "İl " Then
                                                Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | B | " & "<R> " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(4).Cells(j, 34).Value
                                            ElseIf Left(ThisWorkbook.Worksheets(4).Cells(j, 33).Value, 3) = "İlç" Then
                                                Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | B | " & "<R> " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(4).Cells(j, 34).Value
                                            Else
                                                Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | B | " & "<R> " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(4).Cells(j, 33).Value & " " & ThisWorkbook.Worksheets(4).Cells(j, 34).Value
                                            End If
                                        ElseIf ThisWorkbook.Worksheets(4).Cells(j, 34).Value = "" Then
                                            If Left(ThisWorkbook.Worksheets(4).Cells(j, 33).Value, 3) = "İl " Then
                                                Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | B | " & "<R> " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(4).Cells(j, 25).Value & " " & ThisWorkbook.Worksheets(4).Cells(j, 33).Value
                                            ElseIf Left(ThisWorkbook.Worksheets(4).Cells(j, 33).Value, 3) = "İlç" Then
                                                Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | B | " & "<R> " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(4).Cells(j, 26).Value & " " & ThisWorkbook.Worksheets(4).Cells(j, 33).Value
                                            Else
                                                Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | B | " & "<R> " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(4).Cells(j, 33).Value
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
                                
                                ElseIf i = 2 Then

                                    If (ThisWorkbook.Worksheets(4).Cells(j, 105).Value <> "" And CDate(ThisWorkbook.Worksheets(4).Cells(j, 105).Value) > CDate(CalTarih)) Or _
                                        ThisWorkbook.Worksheets(4).Cells(j, 105).Value = "" Then
                                        Cont = Cont + 1

                                        Set IlkSiraBul = Nothing
                                        Set SonSiraBul = Nothing
                                        IlkSira = 0
                                        SonSira = 0
                                        SiraNo = ThisWorkbook.Worksheets(4).Cells(j, 5).Value
                                        Set IlkSiraBul = ThisWorkbook.Worksheets(4).Range("CM7:CM100000").Find(What:=SiraNo, SearchDirection:=xlNext, _
                                                        SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
                                        If Not IlkSiraBul Is Nothing Then
                                            IlkSira = IlkSiraBul.Row
                                        End If
                                        Set SonSiraBul = ThisWorkbook.Worksheets(4).Range("CN7:CN100000").Find(What:=SiraNo, SearchDirection:=xlNext, _
                                                        SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
                                        If Not SonSiraBul Is Nothing Then
                                            SonSira = SonSiraBul.Row
                                        End If
                                        RaporNoKontrol = 0
                                        For k = IlkSira To SonSira
                                            If Left(ThisWorkbook.Worksheets(4).Cells(k, 63).Value, 11) = "Technique A" Then
                                                If ThisWorkbook.Worksheets(4).Cells(k, 11).Value <> "" Then
                                                    If RaporNoKontrol = 0 Then
                                                        RaporNoTakip = ThisWorkbook.Worksheets(4).Cells(k, 11).Value
                                                        RaporNoKontrol = 1
                                                    Else
                                                        RaporNoTakip = RaporNoTakip & ", " & ThisWorkbook.Worksheets(4).Cells(k, 11).Value
                                                    End If
                                                End If
                                            End If
                                        Next k
                                
                                        If ThisWorkbook.Worksheets(4).Cells(j, 34).Value <> "" Then
                                            If Left(ThisWorkbook.Worksheets(4).Cells(j, 33).Value, 3) = "İl " Then
                                                Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | B | " & "<L> (Giden) " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(4).Cells(j, 34).Value
                                            ElseIf Left(ThisWorkbook.Worksheets(4).Cells(j, 33).Value, 3) = "İlç" Then
                                                Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | B | " & "<L> (Giden) " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(4).Cells(j, 34).Value
                                            Else
                                                Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | B | " & "<L> (Giden) " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(4).Cells(j, 33).Value & " " & ThisWorkbook.Worksheets(4).Cells(j, 34).Value
                                            End If
                                        ElseIf ThisWorkbook.Worksheets(4).Cells(j, 34).Value = "" Then
                                            If Left(ThisWorkbook.Worksheets(4).Cells(j, 33).Value, 3) = "İl " Then
                                                Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | B | " & "<L> (Giden) " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(4).Cells(j, 25).Value & " " & ThisWorkbook.Worksheets(4).Cells(j, 33).Value
                                            ElseIf Left(ThisWorkbook.Worksheets(4).Cells(j, 33).Value, 3) = "İlç" Then
                                                Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | B | " & "<L> (Giden) " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(4).Cells(j, 26).Value & " " & ThisWorkbook.Worksheets(4).Cells(j, 33).Value
                                            Else
                                                Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | B | " & "<L> (Giden) " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(4).Cells(j, 33).Value
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
                            Next i
                        End If
                    'End If
                End If
            End If
        End If
    End If
Next j

ContTakipMevcut = Cont

Son:

End Sub

Sub Rapor2_2VarlikHareketleriCikis()
Dim i As Integer, TarihBul As Range
Dim Say As Long, j As Long, Cont As Long, Sno As Variant
Dim LstBx As MSForms.ListBox, ctl As MSForms.Control
Dim LblSira1 As MSForms.Label
Dim LblSira2 As MSForms.Label
Dim Contx As Long, FindKontrol1 As Long, FindKontrol2 As Long, TarihBulSay As Long, TarihBulSay1 As Long, TarihBulSay2 As Long

Dim IlkSiraBul As Range, SonSiraBul As Range, IlkSira As Long, SonSira As Long, k As Long
Dim RaporNoTakip As String, RaporNoKontrol As Integer, SiraNo As Long

CalTarih = CalTarihTakip

Say = ThisWorkbook.Worksheets(4).Range("CN100000").End(xlUp).Row
If Say < 7 Then
    GoTo Son
End If

'Cont = 0
Contx = 0
FindKontrol1 = 0
FindKontrol2 = 0
TarihBulSay1 = 0
TarihBulSay2 = 0
Set TarihBul = ThisWorkbook.Worksheets(4).Range("DA6:DA100000").Find(What:=CalTarih, SearchDirection:=xlNext, _
                SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
If Not TarihBul Is Nothing Then
    TarihBulSay1 = TarihBul.Row
Else
    FindKontrol1 = 1
    'GoTo Son
End If

Set TarihBul = ThisWorkbook.Worksheets(4).Range("CZ6:CZ100000").Find(What:=CalTarih, SearchDirection:=xlNext, _
                SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
If Not TarihBul Is Nothing Then
    TarihBulSay2 = TarihBul.Row
Else
    FindKontrol2 = 1
    'GoTo Son
End If
If FindKontrol1 = 1 And FindKontrol2 = 1 Then
    GoTo Son
End If

If TarihBulSay1 <> 0 And TarihBulSay2 <> 0 And TarihBulSay1 < TarihBulSay2 Then
    TarihBulSay = TarihBulSay1
End If
If TarihBulSay1 <> 0 And TarihBulSay2 <> 0 And TarihBulSay2 < TarihBulSay1 Then
    TarihBulSay = TarihBulSay2
End If
If TarihBulSay1 <> 0 And TarihBulSay2 <> 0 And TarihBulSay2 = TarihBulSay1 Then
    TarihBulSay = TarihBulSay1
End If
If TarihBulSay2 = 0 And TarihBulSay1 <> 0 Then
    TarihBulSay = TarihBulSay1
End If
If TarihBulSay1 = 0 And TarihBulSay2 <> 0 Then
    TarihBulSay = TarihBulSay2
End If
If TarihBulSay1 = 0 And TarihBulSay2 = 0 Then
    GoTo Son
End If

'MsgBox CalTarih

Cont = ContTakipCikis
'Cont = 0
For j = TarihBulSay To Say
    If ThisWorkbook.Worksheets(4).Cells(j, 174).Value = "Yes" Then
        If ThisWorkbook.Worksheets(4).Cells(j, 187).Value = "" Then
            If CalTarih = ThisWorkbook.Worksheets(4).Cells(j, 105).Value And ThisWorkbook.Worksheets(4).Cells(j, 39).Value <> "" Then
                If ThisWorkbook.Worksheets(4).Cells(j, 5).Value <> "" And ThisWorkbook.Worksheets(4).Cells(j, 11).Value <> "" Then
                    Cont = Cont + 1

                    Set IlkSiraBul = Nothing
                    Set SonSiraBul = Nothing
                    IlkSira = 0
                    SonSira = 0
                    SiraNo = ThisWorkbook.Worksheets(4).Cells(j, 5).Value
                    Set IlkSiraBul = ThisWorkbook.Worksheets(4).Range("CM7:CM100000").Find(What:=SiraNo, SearchDirection:=xlNext, _
                                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
                    If Not IlkSiraBul Is Nothing Then
                        IlkSira = IlkSiraBul.Row
                    End If
                    Set SonSiraBul = ThisWorkbook.Worksheets(4).Range("CN7:CN100000").Find(What:=SiraNo, SearchDirection:=xlNext, _
                                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
                    If Not SonSiraBul Is Nothing Then
                        SonSira = SonSiraBul.Row
                    End If
                    RaporNoKontrol = 0
                    For k = IlkSira To SonSira
                        If ThisWorkbook.Worksheets(4).Cells(k, 11).Value <> "" Then
                            If RaporNoKontrol = 0 Then
                                RaporNoTakip = ThisWorkbook.Worksheets(4).Cells(k, 11).Value
                                RaporNoKontrol = 1
                            Else
                                RaporNoTakip = RaporNoTakip & ", " & ThisWorkbook.Worksheets(4).Cells(k, 11).Value
                            End If
                        End If
                    Next k
                    
                    If ThisWorkbook.Worksheets(4).Cells(j, 34).Value <> "" Then
                        If Left(ThisWorkbook.Worksheets(4).Cells(j, 33).Value, 3) = "İl " Then
                            Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | B | " & "<T> " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(4).Cells(j, 34).Value
                        ElseIf Left(ThisWorkbook.Worksheets(4).Cells(j, 33).Value, 3) = "İlç" Then
                            Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | B | " & "<T> " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(4).Cells(j, 34).Value
                        Else
                            Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | B | " & "<T> " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(4).Cells(j, 33).Value & " " & ThisWorkbook.Worksheets(4).Cells(j, 34).Value
                        End If
                    ElseIf ThisWorkbook.Worksheets(4).Cells(j, 34).Value = "" Then
                        If Left(ThisWorkbook.Worksheets(4).Cells(j, 33).Value, 3) = "İl " Then
                            Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | B | " & "<T> " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(4).Cells(j, 25).Value & " " & ThisWorkbook.Worksheets(4).Cells(j, 33).Value
                        ElseIf Left(ThisWorkbook.Worksheets(4).Cells(j, 33).Value, 3) = "İlç" Then
                            Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | B | " & "<T> " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(4).Cells(j, 26).Value & " " & ThisWorkbook.Worksheets(4).Cells(j, 33).Value
                        Else
                            Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | B | " & "<T> " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(4).Cells(j, 33).Value
                        End If
                    End If
    
                    Set LstBx = core_asset_manager_UI.FrameCikis.Controls.Add("Forms.ListBox.1", "Frame1LstBx" & Cont)
                    With LstBx
                      .Top = (Cont - 1) * 12
                      .Left = 18
                      .Height = 12
                      .Width = 300
                      If ThisWorkbook.Worksheets(4).Cells(j, 39).Value = ThisWorkbook.Worksheets(4).Cells(j, 105).Value Then
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

        ElseIf ThisWorkbook.Worksheets(4).Cells(j, 187).Value = "All" Then
            If CalTarih = ThisWorkbook.Worksheets(4).Cells(j, 105).Value And ThisWorkbook.Worksheets(4).Cells(j, 39).Value <> "" Then
                If ThisWorkbook.Worksheets(4).Cells(j, 5).Value <> "" And ThisWorkbook.Worksheets(4).Cells(j, 11).Value <> "" Then
                    Cont = Cont + 1

                    Set IlkSiraBul = Nothing
                    Set SonSiraBul = Nothing
                    IlkSira = 0
                    SonSira = 0
                    SiraNo = ThisWorkbook.Worksheets(4).Cells(j, 5).Value
                    Set IlkSiraBul = ThisWorkbook.Worksheets(4).Range("CM7:CM100000").Find(What:=SiraNo, SearchDirection:=xlNext, _
                                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
                    If Not IlkSiraBul Is Nothing Then
                        IlkSira = IlkSiraBul.Row
                    End If
                    Set SonSiraBul = ThisWorkbook.Worksheets(4).Range("CN7:CN100000").Find(What:=SiraNo, SearchDirection:=xlNext, _
                                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
                    If Not SonSiraBul Is Nothing Then
                        SonSira = SonSiraBul.Row
                    End If
                    RaporNoKontrol = 0
                    For k = IlkSira To SonSira
                        If ThisWorkbook.Worksheets(4).Cells(k, 11).Value <> "" Then
                            If RaporNoKontrol = 0 Then
                                RaporNoTakip = ThisWorkbook.Worksheets(4).Cells(k, 11).Value
                                RaporNoKontrol = 1
                            Else
                                RaporNoTakip = RaporNoTakip & ", " & ThisWorkbook.Worksheets(4).Cells(k, 11).Value
                            End If
                        End If
                    Next k
                    
                    If ThisWorkbook.Worksheets(4).Cells(j, 34).Value <> "" Then
                        If Left(ThisWorkbook.Worksheets(4).Cells(j, 33).Value, 3) = "İl " Then
                            Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | B | " & "<T> (Giden) " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(4).Cells(j, 34).Value
                        ElseIf Left(ThisWorkbook.Worksheets(4).Cells(j, 33).Value, 3) = "İlç" Then
                            Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | B | " & "<T> (Giden) " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(4).Cells(j, 34).Value
                        Else
                            Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | B | " & "<T> (Giden) " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(4).Cells(j, 33).Value & " " & ThisWorkbook.Worksheets(4).Cells(j, 34).Value
                        End If
                    ElseIf ThisWorkbook.Worksheets(4).Cells(j, 34).Value = "" Then
                        If Left(ThisWorkbook.Worksheets(4).Cells(j, 33).Value, 3) = "İl " Then
                            Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | B | " & "<T> (Giden) " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(4).Cells(j, 25).Value & " " & ThisWorkbook.Worksheets(4).Cells(j, 33).Value
                        ElseIf Left(ThisWorkbook.Worksheets(4).Cells(j, 33).Value, 3) = "İlç" Then
                            Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | B | " & "<T> (Giden) " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(4).Cells(j, 26).Value & " " & ThisWorkbook.Worksheets(4).Cells(j, 33).Value
                        Else
                            Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | B | " & "<T> (Giden) " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(4).Cells(j, 33).Value
                        End If
                    End If
    
                    Set LstBx = core_asset_manager_UI.FrameCikis.Controls.Add("Forms.ListBox.1", "Frame1LstBx" & Cont)
                    With LstBx
                      .Top = (Cont - 1) * 12
                      .Left = 18
                      .Height = 12
                      .Width = 300
                      If ThisWorkbook.Worksheets(4).Cells(j, 39).Value = ThisWorkbook.Worksheets(4).Cells(j, 105).Value Then
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
            
        ElseIf ThisWorkbook.Worksheets(4).Cells(j, 187).Value = "Technique A" Then
            'If CalTarih = ThisWorkbook.Worksheets(4).Cells(j, 105).Value And ThisWorkbook.Worksheets(4).Cells(j, 39).Value <> "" Then
                If ThisWorkbook.Worksheets(4).Cells(j, 5).Value <> "" And ThisWorkbook.Worksheets(4).Cells(j, 11).Value <> "" Then
                    For i = 1 To 2
                        'Cont = Cont + 1
                        If i = 1 Then

                            If CalTarih = ThisWorkbook.Worksheets(4).Cells(j, 104).Value And ThisWorkbook.Worksheets(4).Cells(j, 39).Value <> "" Then
                                Cont = Cont + 1

                                Set IlkSiraBul = Nothing
                                Set SonSiraBul = Nothing
                                IlkSira = 0
                                SonSira = 0
                                SiraNo = ThisWorkbook.Worksheets(4).Cells(j, 5).Value
                                Set IlkSiraBul = ThisWorkbook.Worksheets(4).Range("CM7:CM100000").Find(What:=SiraNo, SearchDirection:=xlNext, _
                                                SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
                                If Not IlkSiraBul Is Nothing Then
                                    IlkSira = IlkSiraBul.Row
                                End If
                                Set SonSiraBul = ThisWorkbook.Worksheets(4).Range("CN7:CN100000").Find(What:=SiraNo, SearchDirection:=xlNext, _
                                                SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
                                If Not SonSiraBul Is Nothing Then
                                    SonSira = SonSiraBul.Row
                                End If
                                RaporNoKontrol = 0
                                For k = IlkSira To SonSira
                                    If Left(ThisWorkbook.Worksheets(4).Cells(k, 63).Value, 8) <> "Technique A" Then
                                        If ThisWorkbook.Worksheets(4).Cells(k, 11).Value <> "" Then
                                            If RaporNoKontrol = 0 Then
                                                RaporNoTakip = ThisWorkbook.Worksheets(4).Cells(k, 11).Value
                                                RaporNoKontrol = 1
                                            Else
                                                RaporNoTakip = RaporNoTakip & ", " & ThisWorkbook.Worksheets(4).Cells(k, 11).Value
                                            End If
                                        End If
                                    End If
                                Next k
                                
                                If ThisWorkbook.Worksheets(4).Cells(j, 34).Value <> "" Then
                                    If Left(ThisWorkbook.Worksheets(4).Cells(j, 33).Value, 3) = "İl " Then
                                        Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | B | " & "<R> " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(4).Cells(j, 34).Value
                                    ElseIf Left(ThisWorkbook.Worksheets(4).Cells(j, 33).Value, 3) = "İlç" Then
                                        Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | B | " & "<R> " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(4).Cells(j, 34).Value
                                    Else
                                        Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | B | " & "<R> " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(4).Cells(j, 33).Value & " " & ThisWorkbook.Worksheets(4).Cells(j, 34).Value
                                    End If
                                ElseIf ThisWorkbook.Worksheets(4).Cells(j, 34).Value = "" Then
                                    If Left(ThisWorkbook.Worksheets(4).Cells(j, 33).Value, 3) = "İl " Then
                                        Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | B | " & "<R> " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(4).Cells(j, 25).Value & " " & ThisWorkbook.Worksheets(4).Cells(j, 33).Value
                                    ElseIf Left(ThisWorkbook.Worksheets(4).Cells(j, 33).Value, 3) = "İlç" Then
                                        Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | B | " & "<R> " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(4).Cells(j, 26).Value & " " & ThisWorkbook.Worksheets(4).Cells(j, 33).Value
                                    Else
                                        Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | B | " & "<R> " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(4).Cells(j, 33).Value
                                    End If
                                End If

                                Set LstBx = core_asset_manager_UI.FrameCikis.Controls.Add("Forms.ListBox.1", "Frame1LstBx" & Cont)
                                With LstBx
                                    .Top = (Cont - 1) * 12
                                    .Left = 18
                                    .Height = 12
                                    .Width = 300
                                    If ThisWorkbook.Worksheets(4).Cells(j, 39).Value = ThisWorkbook.Worksheets(4).Cells(j, 104).Value Then
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
                            
                        ElseIf i = 2 Then
                        
                            If CalTarih = ThisWorkbook.Worksheets(4).Cells(j, 105).Value And ThisWorkbook.Worksheets(4).Cells(j, 39).Value <> "" Then
                                Cont = Cont + 1

                                Set IlkSiraBul = Nothing
                                Set SonSiraBul = Nothing
                                IlkSira = 0
                                SonSira = 0
                                SiraNo = ThisWorkbook.Worksheets(4).Cells(j, 5).Value
                                Set IlkSiraBul = ThisWorkbook.Worksheets(4).Range("CM7:CM100000").Find(What:=SiraNo, SearchDirection:=xlNext, _
                                                SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
                                If Not IlkSiraBul Is Nothing Then
                                    IlkSira = IlkSiraBul.Row
                                End If
                                Set SonSiraBul = ThisWorkbook.Worksheets(4).Range("CN7:CN100000").Find(What:=SiraNo, SearchDirection:=xlNext, _
                                                SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
                                If Not SonSiraBul Is Nothing Then
                                    SonSira = SonSiraBul.Row
                                End If
                                RaporNoKontrol = 0
                                For k = IlkSira To SonSira
                                    If Left(ThisWorkbook.Worksheets(4).Cells(k, 63).Value, 11) = "Technique A" Then
                                        If ThisWorkbook.Worksheets(4).Cells(k, 11).Value <> "" Then
                                            If RaporNoKontrol = 0 Then
                                                RaporNoTakip = ThisWorkbook.Worksheets(4).Cells(k, 11).Value
                                                RaporNoKontrol = 1
                                            Else
                                                RaporNoTakip = RaporNoTakip & ", " & ThisWorkbook.Worksheets(4).Cells(k, 11).Value
                                            End If
                                        End If
                                    End If
                                Next k
                                
                                If ThisWorkbook.Worksheets(4).Cells(j, 34).Value <> "" Then
                                    If Left(ThisWorkbook.Worksheets(4).Cells(j, 33).Value, 3) = "İl " Then
                                        Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | B | " & "<L> (Giden) " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(4).Cells(j, 34).Value
                                    ElseIf Left(ThisWorkbook.Worksheets(4).Cells(j, 33).Value, 3) = "İlç" Then
                                        Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | B | " & "<L> (Giden) " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(4).Cells(j, 34).Value
                                    Else
                                        Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | B | " & "<L> (Giden) " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(4).Cells(j, 33).Value & " " & ThisWorkbook.Worksheets(4).Cells(j, 34).Value
                                    End If
                                ElseIf ThisWorkbook.Worksheets(4).Cells(j, 34).Value = "" Then
                                    If Left(ThisWorkbook.Worksheets(4).Cells(j, 33).Value, 3) = "İl " Then
                                        Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | B | " & "<L> (Giden) " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(4).Cells(j, 25).Value & " " & ThisWorkbook.Worksheets(4).Cells(j, 33).Value
                                    ElseIf Left(ThisWorkbook.Worksheets(4).Cells(j, 33).Value, 3) = "İlç" Then
                                        Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | B | " & "<L> (Giden) " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(4).Cells(j, 26).Value & " " & ThisWorkbook.Worksheets(4).Cells(j, 33).Value
                                    Else
                                        Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | B | " & "<L> (Giden) " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(4).Cells(j, 33).Value
                                    End If
                                End If

                                Set LstBx = core_asset_manager_UI.FrameCikis.Controls.Add("Forms.ListBox.1", "Frame1LstBx" & Cont)
                                With LstBx
                                    .Top = (Cont - 1) * 12
                                    .Left = 18
                                    .Height = 12
                                    .Width = 300
                                    If ThisWorkbook.Worksheets(4).Cells(j, 39).Value = ThisWorkbook.Worksheets(4).Cells(j, 105).Value Then
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
                    Next i
                End If
            'End If
        End If
    End If
Next j

ContTakipCikis = Cont

Son:

End Sub

Sub Rapor2_2VarlikHareketleriGirisXXXMud()
Dim i As Integer, TarihBul As Range
Dim Say As Long, j As Long, Cont As Long, Sno As Variant
Dim LstBx As MSForms.ListBox, ctl As MSForms.Control
Dim LblSira1 As MSForms.Label
Dim LblSira2 As MSForms.Label
Dim Contx As Long

Dim IlkSiraBul As Range, SonSiraBul As Range, IlkSira As Long, SonSira As Long, k As Long
Dim RaporNoTakip As String, RaporNoKontrol As Integer, SiraNo As Long


CalTarih = CalTarihTakip

Say = ThisWorkbook.Worksheets(4).Range("CN100000").End(xlUp).Row
If Say < 7 Then
    GoTo Son
End If

'Cont = 0
Contx = 0
Set TarihBul = ThisWorkbook.Worksheets(4).Range("GI6:GI100000").Find(What:=CalTarih, SearchDirection:=xlNext, _
                SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole) 'Tutanak1 tarihi içinde ara
If Not TarihBul Is Nothing Then
    '
Else
    GoTo Son
End If

'MsgBox CalTarih

Cont = ContTakipGiris
'Cont = 0
For j = TarihBul.Row To Say
    If ThisWorkbook.Worksheets(4).Cells(j, 174).Value = "Yes" Then
        If ThisWorkbook.Worksheets(4).Cells(j, 187).Value = "All" Then
            If CalTarih = ThisWorkbook.Worksheets(4).Cells(j, 191).Value And CDate(ThisWorkbook.Worksheets(4).Cells(j, 191).Value) <> CDate(ThisWorkbook.Worksheets(4).Cells(j, 106).Value) Then
                If ThisWorkbook.Worksheets(4).Cells(j, 5).Value <> "" And ThisWorkbook.Worksheets(4).Cells(j, 11).Value <> "" Then
                    Cont = Cont + 1

                    Set IlkSiraBul = Nothing
                    Set SonSiraBul = Nothing
                    IlkSira = 0
                    SonSira = 0
                    SiraNo = ThisWorkbook.Worksheets(4).Cells(j, 5).Value
                    Set IlkSiraBul = ThisWorkbook.Worksheets(4).Range("CM7:CM100000").Find(What:=SiraNo, SearchDirection:=xlNext, _
                                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
                    If Not IlkSiraBul Is Nothing Then
                        IlkSira = IlkSiraBul.Row
                    End If
                    Set SonSiraBul = ThisWorkbook.Worksheets(4).Range("CN7:CN100000").Find(What:=SiraNo, SearchDirection:=xlNext, _
                                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
                    If Not SonSiraBul Is Nothing Then
                        SonSira = SonSiraBul.Row
                    End If
                    RaporNoKontrol = 0
                    For k = IlkSira To SonSira
                        If ThisWorkbook.Worksheets(4).Cells(k, 11).Value <> "" Then
                            If RaporNoKontrol = 0 Then
                                RaporNoTakip = ThisWorkbook.Worksheets(4).Cells(k, 11).Value
                                RaporNoKontrol = 1
                            Else
                                RaporNoTakip = RaporNoTakip & ", " & ThisWorkbook.Worksheets(4).Cells(k, 11).Value
                            End If
                        End If
                    Next k
                    
                    If ThisWorkbook.Worksheets(4).Cells(j, 34).Value <> "" Then
                        If Left(ThisWorkbook.Worksheets(4).Cells(j, 33).Value, 3) = "İl " Then
                            Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | B | " & "<T> (Gelen) " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(4).Cells(j, 34).Value
                        ElseIf Left(ThisWorkbook.Worksheets(4).Cells(j, 33).Value, 3) = "İlç" Then
                            Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | B | " & "<T> (Gelen) " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(4).Cells(j, 34).Value
                        Else
                            Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | B | " & "<T> (Gelen) " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(4).Cells(j, 33).Value & " " & ThisWorkbook.Worksheets(4).Cells(j, 34).Value
                        End If
                    ElseIf ThisWorkbook.Worksheets(4).Cells(j, 34).Value = "" Then
                        If Left(ThisWorkbook.Worksheets(4).Cells(j, 33).Value, 3) = "İl " Then
                            Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | B | " & "<T> (Gelen) " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(4).Cells(j, 25).Value & " " & ThisWorkbook.Worksheets(4).Cells(j, 33).Value
                        ElseIf Left(ThisWorkbook.Worksheets(4).Cells(j, 33).Value, 3) = "İlç" Then
                            Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | B | " & "<T> (Gelen) " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(4).Cells(j, 26).Value & " " & ThisWorkbook.Worksheets(4).Cells(j, 33).Value
                        Else
                            Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | B | " & "<T> (Gelen) " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(4).Cells(j, 33).Value
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
            
        ElseIf ThisWorkbook.Worksheets(4).Cells(j, 187).Value = "Technique A" Then
            If CalTarih = ThisWorkbook.Worksheets(4).Cells(j, 191).Value And CDate(ThisWorkbook.Worksheets(4).Cells(j, 191).Value) <> CDate(ThisWorkbook.Worksheets(4).Cells(j, 106).Value) Then
                If ThisWorkbook.Worksheets(4).Cells(j, 5).Value <> "" And ThisWorkbook.Worksheets(4).Cells(j, 11).Value <> "" Then
                    'If CalTarih = ThisWorkbook.Worksheets(4).Cells(j, 191).Value And CDate(ThisWorkbook.Worksheets(4).Cells(j, 191).Value) <> CDate(ThisWorkbook.Worksheets(4).Cells(j, 106).Value) Then
                        Cont = Cont + 1

                        Set IlkSiraBul = Nothing
                        Set SonSiraBul = Nothing
                        IlkSira = 0
                        SonSira = 0
                        SiraNo = ThisWorkbook.Worksheets(4).Cells(j, 5).Value
                        Set IlkSiraBul = ThisWorkbook.Worksheets(4).Range("CM7:CM100000").Find(What:=SiraNo, SearchDirection:=xlNext, _
                                        SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
                        If Not IlkSiraBul Is Nothing Then
                            IlkSira = IlkSiraBul.Row
                        End If
                        Set SonSiraBul = ThisWorkbook.Worksheets(4).Range("CN7:CN100000").Find(What:=SiraNo, SearchDirection:=xlNext, _
                                        SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
                        If Not SonSiraBul Is Nothing Then
                            SonSira = SonSiraBul.Row
                        End If
                        RaporNoKontrol = 0
                        For k = IlkSira To SonSira
                            If Left(ThisWorkbook.Worksheets(4).Cells(k, 63).Value, 11) = "Technique A" Then
                                If ThisWorkbook.Worksheets(4).Cells(k, 11).Value <> "" Then
                                    If RaporNoKontrol = 0 Then
                                        RaporNoTakip = ThisWorkbook.Worksheets(4).Cells(k, 11).Value
                                        RaporNoKontrol = 1
                                    Else
                                        RaporNoTakip = RaporNoTakip & ", " & ThisWorkbook.Worksheets(4).Cells(k, 11).Value
                                    End If
                                End If
                            End If
                        Next k
                                
                        If ThisWorkbook.Worksheets(4).Cells(j, 34).Value <> "" Then
                            If Left(ThisWorkbook.Worksheets(4).Cells(j, 33).Value, 3) = "İl " Then
                                Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | B | " & "<L> (Gelen) " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(4).Cells(j, 34).Value
                            ElseIf Left(ThisWorkbook.Worksheets(4).Cells(j, 33).Value, 3) = "İlç" Then
                                Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | B | " & "<L> (Gelen) " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(4).Cells(j, 34).Value
                            Else
                                Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | B | " & "<L> (Gelen) " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(4).Cells(j, 33).Value & " " & ThisWorkbook.Worksheets(4).Cells(j, 34).Value
                            End If
                        ElseIf ThisWorkbook.Worksheets(4).Cells(j, 34).Value = "" Then
                            If Left(ThisWorkbook.Worksheets(4).Cells(j, 33).Value, 3) = "İl " Then
                                Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | B | " & "<L> (Gelen) " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(4).Cells(j, 25).Value & " " & ThisWorkbook.Worksheets(4).Cells(j, 33).Value
                            ElseIf Left(ThisWorkbook.Worksheets(4).Cells(j, 33).Value, 3) = "İlç" Then
                                Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | B | " & "<L> (Gelen) " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(4).Cells(j, 26).Value & " " & ThisWorkbook.Worksheets(4).Cells(j, 33).Value
                            Else
                                Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | B | " & "<L> (Gelen) " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(4).Cells(j, 33).Value
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
                    'End If
                End If
            End If
        End If
    End If
Next j

ContTakipGiris = Cont

Son:

End Sub

Sub Rapor2_2VarlikHareketleriMevcutXXXMud()
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

Say = ThisWorkbook.Worksheets(4).Range("CN100000").End(xlUp).Row
If Say < 7 Then
    GoTo Son
End If

'Mevcutta satırı bul
'CalTarih = CDate(CalTarih)
MevcutSatir = 0
For j = Say To 7 Step -1
    If ThisWorkbook.Worksheets(4).Cells(j, 174).Value = "Yes" Then
        If ThisWorkbook.Worksheets(4).Range("GI" & j).Value <> "" Then
            TarihAra = ThisWorkbook.Worksheets(4).Range("GI" & j)
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
    If ThisWorkbook.Worksheets(4).Cells(j, 174).Value = "Yes" Then
        If ThisWorkbook.Worksheets(4).Cells(j, 187).Value = "All" Then
            If ThisWorkbook.Worksheets(4).Cells(j, 191).Value <> "" Then
                TarihAra = ThisWorkbook.Worksheets(4).Range("GI" & j)
                If CDate(TarihAra) < CDate(CalTarih) Then
                    If (ThisWorkbook.Worksheets(4).Cells(j, 106).Value <> "" And CDate(ThisWorkbook.Worksheets(4).Cells(j, 106).Value) > CDate(CalTarih)) Or _
                        ThisWorkbook.Worksheets(4).Cells(j, 106).Value = "" Then
                        'MsgBox TarihAra & " < " & CalTarih
                        If ThisWorkbook.Worksheets(4).Cells(j, 5).Value <> "" And ThisWorkbook.Worksheets(4).Cells(j, 11).Value <> "" Then
                            Cont = Cont + 1

                            Set IlkSiraBul = Nothing
                            Set SonSiraBul = Nothing
                            IlkSira = 0
                            SonSira = 0
                            SiraNo = ThisWorkbook.Worksheets(4).Cells(j, 5).Value
                            Set IlkSiraBul = ThisWorkbook.Worksheets(4).Range("CM7:CM100000").Find(What:=SiraNo, SearchDirection:=xlNext, _
                                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
                            If Not IlkSiraBul Is Nothing Then
                                IlkSira = IlkSiraBul.Row
                            End If
                            Set SonSiraBul = ThisWorkbook.Worksheets(4).Range("CN7:CN100000").Find(What:=SiraNo, SearchDirection:=xlNext, _
                                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
                            If Not SonSiraBul Is Nothing Then
                                SonSira = SonSiraBul.Row
                            End If
                            RaporNoKontrol = 0
                            For k = IlkSira To SonSira
                                If ThisWorkbook.Worksheets(4).Cells(k, 11).Value <> "" Then
                                    If RaporNoKontrol = 0 Then
                                        RaporNoTakip = ThisWorkbook.Worksheets(4).Cells(k, 11).Value
                                        RaporNoKontrol = 1
                                    Else
                                        RaporNoTakip = RaporNoTakip & ", " & ThisWorkbook.Worksheets(4).Cells(k, 11).Value
                                    End If
                                End If
                            Next k
                    
                            If ThisWorkbook.Worksheets(4).Cells(j, 34).Value <> "" Then
                                If Left(ThisWorkbook.Worksheets(4).Cells(j, 33).Value, 3) = "İl " Then
                                    Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | B | " & "<T> (Gelen) " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(4).Cells(j, 34).Value
                                ElseIf Left(ThisWorkbook.Worksheets(4).Cells(j, 33).Value, 3) = "İlç" Then
                                    Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | B | " & "<T> (Gelen) " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(4).Cells(j, 34).Value
                                Else
                                    Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | B | " & "<T> (Gelen) " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(4).Cells(j, 33).Value & " " & ThisWorkbook.Worksheets(4).Cells(j, 34).Value
                                End If
                            ElseIf ThisWorkbook.Worksheets(4).Cells(j, 34).Value = "" Then
                                If Left(ThisWorkbook.Worksheets(4).Cells(j, 33).Value, 3) = "İl " Then
                                    Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | B | " & "<T> (Gelen) " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(4).Cells(j, 25).Value & " " & ThisWorkbook.Worksheets(4).Cells(j, 33).Value
                                ElseIf Left(ThisWorkbook.Worksheets(4).Cells(j, 33).Value, 3) = "İlç" Then
                                    Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | B | " & "<T> (Gelen) " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(4).Cells(j, 26).Value & " " & ThisWorkbook.Worksheets(4).Cells(j, 33).Value
                                Else
                                    Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | B | " & "<T> (Gelen) " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(4).Cells(j, 33).Value
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
        ElseIf ThisWorkbook.Worksheets(4).Cells(j, 187).Value = "Technique A" Then
            If ThisWorkbook.Worksheets(4).Cells(j, 191).Value <> "" Then
                TarihAra = ThisWorkbook.Worksheets(4).Range("GI" & j)
                If CDate(TarihAra) < CDate(CalTarih) Then
'                    If (ThisWorkbook.Worksheets(4).Cells(j, 105).Value <> "" And CDate(ThisWorkbook.Worksheets(4).Cells(j, 105).Value) > CDate(CalTarih)) Or _
'                        ThisWorkbook.Worksheets(4).Cells(j, 105).Value = "" Then
                        'MsgBox TarihAra & " < " & CalTarih
                        If ThisWorkbook.Worksheets(4).Cells(j, 5).Value <> "" And ThisWorkbook.Worksheets(4).Cells(j, 11).Value <> "" Then
                            If (ThisWorkbook.Worksheets(4).Cells(j, 106).Value <> "" And CDate(ThisWorkbook.Worksheets(4).Cells(j, 106).Value) > CDate(CalTarih)) Or _
                                ThisWorkbook.Worksheets(4).Cells(j, 106).Value = "" Then
                                Cont = Cont + 1

                                Set IlkSiraBul = Nothing
                                Set SonSiraBul = Nothing
                                IlkSira = 0
                                SonSira = 0
                                SiraNo = ThisWorkbook.Worksheets(4).Cells(j, 5).Value
                                Set IlkSiraBul = ThisWorkbook.Worksheets(4).Range("CM7:CM100000").Find(What:=SiraNo, SearchDirection:=xlNext, _
                                                SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
                                If Not IlkSiraBul Is Nothing Then
                                    IlkSira = IlkSiraBul.Row
                                End If
                                Set SonSiraBul = ThisWorkbook.Worksheets(4).Range("CN7:CN100000").Find(What:=SiraNo, SearchDirection:=xlNext, _
                                                SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
                                If Not SonSiraBul Is Nothing Then
                                    SonSira = SonSiraBul.Row
                                End If
                                RaporNoKontrol = 0
                                For k = IlkSira To SonSira
                                    If Left(ThisWorkbook.Worksheets(4).Cells(k, 63).Value, 11) = "Technique A" Then
                                        If ThisWorkbook.Worksheets(4).Cells(k, 11).Value <> "" Then
                                            If RaporNoKontrol = 0 Then
                                                RaporNoTakip = ThisWorkbook.Worksheets(4).Cells(k, 11).Value
                                                RaporNoKontrol = 1
                                            Else
                                                RaporNoTakip = RaporNoTakip & ", " & ThisWorkbook.Worksheets(4).Cells(k, 11).Value
                                            End If
                                        End If
                                    End If
                                Next k

                                If ThisWorkbook.Worksheets(4).Cells(j, 34).Value <> "" Then
                                    If Left(ThisWorkbook.Worksheets(4).Cells(j, 33).Value, 3) = "İl " Then
                                        Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | B | " & "<L> (Gelen) " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(4).Cells(j, 34).Value
                                    ElseIf Left(ThisWorkbook.Worksheets(4).Cells(j, 33).Value, 3) = "İlç" Then
                                        Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | B | " & "<L> (Gelen) " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(4).Cells(j, 34).Value
                                    Else
                                        Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | B | " & "<L> (Gelen) " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(4).Cells(j, 33).Value & " " & ThisWorkbook.Worksheets(4).Cells(j, 34).Value
                                    End If
                                ElseIf ThisWorkbook.Worksheets(4).Cells(j, 34).Value = "" Then
                                    If Left(ThisWorkbook.Worksheets(4).Cells(j, 33).Value, 3) = "İl " Then
                                        Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | B | " & "<L> (Gelen) " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(4).Cells(j, 25).Value & " " & ThisWorkbook.Worksheets(4).Cells(j, 33).Value
                                    ElseIf Left(ThisWorkbook.Worksheets(4).Cells(j, 33).Value, 3) = "İlç" Then
                                        Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | B | " & "<L> (Gelen) " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(4).Cells(j, 26).Value & " " & ThisWorkbook.Worksheets(4).Cells(j, 33).Value
                                    Else
                                        Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | B | " & "<L> (Gelen) " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(4).Cells(j, 33).Value
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
                    'End If
                End If
            End If
        End If
    End If
Next j

ContTakipMevcut = Cont

Son:

End Sub

Sub Rapor2_2VarlikHareketleriCikisXXXMud()
Dim i As Integer, TarihBul As Range
Dim Say As Long, j As Long, Cont As Long, Sno As Variant
Dim LstBx As MSForms.ListBox, ctl As MSForms.Control
Dim LblSira1 As MSForms.Label
Dim LblSira2 As MSForms.Label
Dim Contx As Long, FindKontrol1 As Long, FindKontrol2 As Long, TarihBulSay As Long, TarihBulSay1 As Long, TarihBulSay2 As Long

Dim IlkSiraBul As Range, SonSiraBul As Range, IlkSira As Long, SonSira As Long, k As Long
Dim RaporNoTakip As String, RaporNoKontrol As Integer, SiraNo As Long


CalTarih = CalTarihTakip

Say = ThisWorkbook.Worksheets(4).Range("CN100000").End(xlUp).Row
If Say < 7 Then
    GoTo Son
End If

'Cont = 0
Contx = 0
Set TarihBul = ThisWorkbook.Worksheets(4).Range("DB6:DB100000").Find(What:=CalTarih, SearchDirection:=xlNext, _
                SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
If Not TarihBul Is Nothing Then
    TarihBulSay = TarihBul.Row
Else
    GoTo Son
End If


'MsgBox CalTarih

Cont = ContTakipCikis
'Cont = 0
For j = TarihBulSay To Say
    If ThisWorkbook.Worksheets(4).Cells(j, 174).Value = "Yes" Then
        If ThisWorkbook.Worksheets(4).Cells(j, 187).Value = "All" Then
            If CalTarih = ThisWorkbook.Worksheets(4).Cells(j, 106).Value And ThisWorkbook.Worksheets(4).Cells(j, 191).Value <> "" Then
                If ThisWorkbook.Worksheets(4).Cells(j, 5).Value <> "" And ThisWorkbook.Worksheets(4).Cells(j, 11).Value <> "" Then
                    Cont = Cont + 1

                    Set IlkSiraBul = Nothing
                    Set SonSiraBul = Nothing
                    IlkSira = 0
                    SonSira = 0
                    SiraNo = ThisWorkbook.Worksheets(4).Cells(j, 5).Value
                    Set IlkSiraBul = ThisWorkbook.Worksheets(4).Range("CM7:CM100000").Find(What:=SiraNo, SearchDirection:=xlNext, _
                                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
                    If Not IlkSiraBul Is Nothing Then
                        IlkSira = IlkSiraBul.Row
                    End If
                    Set SonSiraBul = ThisWorkbook.Worksheets(4).Range("CN7:CN100000").Find(What:=SiraNo, SearchDirection:=xlNext, _
                                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
                    If Not SonSiraBul Is Nothing Then
                        SonSira = SonSiraBul.Row
                    End If
                    RaporNoKontrol = 0
                    For k = IlkSira To SonSira
                        If ThisWorkbook.Worksheets(4).Cells(k, 11).Value <> "" Then
                            If RaporNoKontrol = 0 Then
                                RaporNoTakip = ThisWorkbook.Worksheets(4).Cells(k, 11).Value
                                RaporNoKontrol = 1
                            Else
                                RaporNoTakip = RaporNoTakip & ", " & ThisWorkbook.Worksheets(4).Cells(k, 11).Value
                            End If
                        End If
                    Next k
                    
                    If ThisWorkbook.Worksheets(4).Cells(j, 34).Value <> "" Then
                        If Left(ThisWorkbook.Worksheets(4).Cells(j, 33).Value, 3) = "İl " Then
                            Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | B | " & "<T> (Gelen) " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(4).Cells(j, 34).Value
                        ElseIf Left(ThisWorkbook.Worksheets(4).Cells(j, 33).Value, 3) = "İlç" Then
                            Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | B | " & "<T> (Gelen) " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(4).Cells(j, 34).Value
                        Else
                            Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | B | " & "<T> (Gelen) " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(4).Cells(j, 33).Value & " " & ThisWorkbook.Worksheets(4).Cells(j, 34).Value
                        End If
                    ElseIf ThisWorkbook.Worksheets(4).Cells(j, 34).Value = "" Then
                        If Left(ThisWorkbook.Worksheets(4).Cells(j, 33).Value, 3) = "İl " Then
                            Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | B | " & "<T> (Gelen) " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(4).Cells(j, 25).Value & " " & ThisWorkbook.Worksheets(4).Cells(j, 33).Value
                        ElseIf Left(ThisWorkbook.Worksheets(4).Cells(j, 33).Value, 3) = "İlç" Then
                            Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | B | " & "<T> (Gelen) " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(4).Cells(j, 26).Value & " " & ThisWorkbook.Worksheets(4).Cells(j, 33).Value
                        Else
                            Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | B | " & "<T> (Gelen) " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(4).Cells(j, 33).Value
                        End If
                    End If
    
                    Set LstBx = core_asset_manager_UI.FrameCikis.Controls.Add("Forms.ListBox.1", "Frame1LstBx" & Cont)
                    With LstBx
                      .Top = (Cont - 1) * 12
                      .Left = 18
                      .Height = 12
                      .Width = 300
                      If ThisWorkbook.Worksheets(4).Cells(j, 191).Value = ThisWorkbook.Worksheets(4).Cells(j, 106).Value Then
                          .BackColor = &H80000000 'Giriş
                      Else
                          .BackColor = &H80000003 'Mevcut
                      End If
                      .SpecialEffect = fmSpecialEffectFlat
                      .MultiSelect = fmMultiSelectMulti
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
        ElseIf ThisWorkbook.Worksheets(4).Cells(j, 187).Value = "Technique A" Then
            'If CalTarih = ThisWorkbook.Worksheets(4).Cells(j, 105).Value And ThisWorkbook.Worksheets(4).Cells(j, 39).Value <> "" Then
                If ThisWorkbook.Worksheets(4).Cells(j, 5).Value <> "" And ThisWorkbook.Worksheets(4).Cells(j, 11).Value <> "" Then
                    If CalTarih = ThisWorkbook.Worksheets(4).Cells(j, 106).Value And ThisWorkbook.Worksheets(4).Cells(j, 191).Value <> "" Then
                        Cont = Cont + 1

                        Set IlkSiraBul = Nothing
                        Set SonSiraBul = Nothing
                        IlkSira = 0
                        SonSira = 0
                        SiraNo = ThisWorkbook.Worksheets(4).Cells(j, 5).Value
                        Set IlkSiraBul = ThisWorkbook.Worksheets(4).Range("CM7:CM100000").Find(What:=SiraNo, SearchDirection:=xlNext, _
                                        SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
                        If Not IlkSiraBul Is Nothing Then
                            IlkSira = IlkSiraBul.Row
                        End If
                        Set SonSiraBul = ThisWorkbook.Worksheets(4).Range("CN7:CN100000").Find(What:=SiraNo, SearchDirection:=xlNext, _
                                        SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
                        If Not SonSiraBul Is Nothing Then
                            SonSira = SonSiraBul.Row
                        End If
                        RaporNoKontrol = 0
                        For k = IlkSira To SonSira
                            If Left(ThisWorkbook.Worksheets(4).Cells(k, 63).Value, 11) = "Technique A" Then
                                If ThisWorkbook.Worksheets(4).Cells(k, 11).Value <> "" Then
                                    If RaporNoKontrol = 0 Then
                                        RaporNoTakip = ThisWorkbook.Worksheets(4).Cells(k, 11).Value
                                        RaporNoKontrol = 1
                                    Else
                                        RaporNoTakip = RaporNoTakip & ", " & ThisWorkbook.Worksheets(4).Cells(k, 11).Value
                                    End If
                                End If
                            End If
                        Next k

                        If ThisWorkbook.Worksheets(4).Cells(j, 34).Value <> "" Then
                            If Left(ThisWorkbook.Worksheets(4).Cells(j, 33).Value, 3) = "İl " Then
                                Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | B | " & "<L> (Gelen) " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(4).Cells(j, 34).Value
                            ElseIf Left(ThisWorkbook.Worksheets(4).Cells(j, 33).Value, 3) = "İlç" Then
                                Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | B | " & "<L> (Gelen) " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(4).Cells(j, 34).Value
                            Else
                                Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | B | " & "<L> (Gelen) " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(4).Cells(j, 33).Value & " " & ThisWorkbook.Worksheets(4).Cells(j, 34).Value
                            End If
                        ElseIf ThisWorkbook.Worksheets(4).Cells(j, 34).Value = "" Then
                            If Left(ThisWorkbook.Worksheets(4).Cells(j, 33).Value, 3) = "İl " Then
                                Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | B | " & "<L> (Gelen) " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(4).Cells(j, 25).Value & " " & ThisWorkbook.Worksheets(4).Cells(j, 33).Value
                            ElseIf Left(ThisWorkbook.Worksheets(4).Cells(j, 33).Value, 3) = "İlç" Then
                                Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | B | " & "<L> (Gelen) " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(4).Cells(j, 26).Value & " " & ThisWorkbook.Worksheets(4).Cells(j, 33).Value
                            Else
                                Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | B | " & "<L> (Gelen) " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(4).Cells(j, 33).Value
                            End If
                        End If

                        Set LstBx = core_asset_manager_UI.FrameCikis.Controls.Add("Forms.ListBox.1", "Frame1LstBx" & Cont)
                        With LstBx
                            .Top = (Cont - 1) * 12
                            .Left = 18
                            .Height = 12
                            .Width = 300
                            If ThisWorkbook.Worksheets(4).Cells(j, 191).Value = ThisWorkbook.Worksheets(4).Cells(j, 106).Value Then
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
            'End If
        End If
    End If
Next j

ContTakipCikis = Cont

Son:

End Sub

Sub Rapor2_2BilgilendirmeTeslimTutanaklari()
Dim i As Integer, TarihBul As Range
Dim Say As Long, j As Long, Cont As Long, Sno As Variant
Dim LstBx As MSForms.ListBox, ctl As MSForms.Control
Dim LblSira1 As MSForms.Label
Dim LblSira2 As MSForms.Label
Dim Contx As Long

CalTarih = CalTarihTakip

Say = ThisWorkbook.Worksheets(4).Range("CN100000").End(xlUp).Row
If Say < 7 Then
    GoTo Son
End If

'Cont = 0
Contx = 0
TarihiTekrarla1:
Set TarihBul = ThisWorkbook.Worksheets(4).Range("CE6:CE100000").Find(What:=CalTarih, SearchDirection:=xlNext, _
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

'Bilgilendirme
For j = Say To TarihBul.Row Step -1
    'If Cont < 50 Then
        If ThisWorkbook.Worksheets(4).Cells(j, 5).Value <> "" And ThisWorkbook.Worksheets(4).Cells(j, 83).Value <> "" And ThisWorkbook.Worksheets(4).Cells(j, 174).Value = "Yes" Then
            Cont = Cont + 1
            If ThisWorkbook.Worksheets(4).Cells(j, 200).Value <> "" Then
                If Left(ThisWorkbook.Worksheets(4).Cells(j, 199).Value, 3) = "İl " Then
                    Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | R | " & ThisWorkbook.Worksheets(4).Cells(j, 84).Value & " | " & "Bilgilendirme (" & ThisWorkbook.Worksheets(4).Cells(j, 200).Value & ")"
                ElseIf Left(ThisWorkbook.Worksheets(4).Cells(j, 199).Value, 3) = "İlç" Then
                    Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | R | " & ThisWorkbook.Worksheets(4).Cells(j, 84).Value & " | " & "Bilgilendirme (" & ThisWorkbook.Worksheets(4).Cells(j, 200).Value & ")"
                Else
                    Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | R | " & ThisWorkbook.Worksheets(4).Cells(j, 84).Value & " | " & "Bilgilendirme (" & ThisWorkbook.Worksheets(4).Cells(j, 199).Value & " " & ThisWorkbook.Worksheets(4).Cells(j, 200).Value & ")"
                End If
            ElseIf ThisWorkbook.Worksheets(4).Cells(j, 200).Value = "" Then
                If Left(ThisWorkbook.Worksheets(4).Cells(j, 199).Value, 3) = "İl " Then
                    Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | R | " & ThisWorkbook.Worksheets(4).Cells(j, 84).Value & " | " & "Bilgilendirme (" & ThisWorkbook.Worksheets(4).Cells(j, 203).Value & " " & ThisWorkbook.Worksheets(4).Cells(j, 199).Value & ")"
                ElseIf Left(ThisWorkbook.Worksheets(4).Cells(j, 199).Value, 3) = "İlç" Then
                    Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | R | " & ThisWorkbook.Worksheets(4).Cells(j, 84).Value & " | " & "Bilgilendirme (" & ThisWorkbook.Worksheets(4).Cells(j, 204).Value & " " & ThisWorkbook.Worksheets(4).Cells(j, 199).Value & ")"
                Else
                    Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | R | " & ThisWorkbook.Worksheets(4).Cells(j, 84).Value & " | " & "Bilgilendirme (" & ThisWorkbook.Worksheets(4).Cells(j, 199).Value & ")"
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

End Sub

Sub Rapor2_2XXXMudTeslimTutanaklari()
Dim i As Integer, TarihBul As Range
Dim Say As Long, j As Long, Cont As Long, Sno As Variant
Dim LstBx As MSForms.ListBox, ctl As MSForms.Control
Dim LblSira1 As MSForms.Label
Dim LblSira2 As MSForms.Label
Dim Contx As Long

CalTarih = CalTarihTakip

Say = ThisWorkbook.Worksheets(4).Range("CN100000").End(xlUp).Row
If Say < 7 Then
    GoTo Son
End If

'XXXMud
Contx = 0
TarihiTekrarla2:
Set TarihBul = ThisWorkbook.Worksheets(4).Range("FS6:FS100000").Find(What:=CalTarih, SearchDirection:=xlNext, _
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
    GoTo TarihiTekrarla2
End If

Cont = ContTakip
'Cont = 0

For j = Say To TarihBul.Row Step -1
    'If Cont < 50 Then
        If ThisWorkbook.Worksheets(4).Cells(j, 5).Value <> "" And ThisWorkbook.Worksheets(4).Cells(j, 175).Value <> "" Then
            Cont = Cont + 1
            If ThisWorkbook.Worksheets(4).Cells(j, 34).Value <> "" Then
                If Left(ThisWorkbook.Worksheets(4).Cells(j, 33).Value, 3) = "İl " Then
                    Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | B | " & ThisWorkbook.Worksheets(4).Cells(j, 176).Value & " | " & "XXXMud (" & ThisWorkbook.Worksheets(4).Cells(j, 34).Value & ")"
                ElseIf Left(ThisWorkbook.Worksheets(4).Cells(j, 33).Value, 3) = "İlç" Then
                    Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | B | " & ThisWorkbook.Worksheets(4).Cells(j, 176).Value & " | " & "XXXMud (" & ThisWorkbook.Worksheets(4).Cells(j, 34).Value & ")"
                Else
                    Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | B | " & ThisWorkbook.Worksheets(4).Cells(j, 176).Value & " | " & "XXXMud (" & ThisWorkbook.Worksheets(4).Cells(j, 33).Value & " " & ThisWorkbook.Worksheets(4).Cells(j, 34).Value & ")"
                End If
            ElseIf ThisWorkbook.Worksheets(4).Cells(j, 34).Value = "" Then
                If Left(ThisWorkbook.Worksheets(4).Cells(j, 33).Value, 3) = "İl " Then
                    Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | B | " & ThisWorkbook.Worksheets(4).Cells(j, 176).Value & " | " & "XXXMud (" & ThisWorkbook.Worksheets(4).Cells(j, 25).Value & " " & ThisWorkbook.Worksheets(4).Cells(j, 33).Value & ")"
                ElseIf Left(ThisWorkbook.Worksheets(4).Cells(j, 33).Value, 3) = "İlç" Then
                    Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | B | " & ThisWorkbook.Worksheets(4).Cells(j, 176).Value & " | " & "XXXMud (" & ThisWorkbook.Worksheets(4).Cells(j, 26).Value & " " & ThisWorkbook.Worksheets(4).Cells(j, 33).Value & ")"
                Else
                    Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | B | " & ThisWorkbook.Worksheets(4).Cells(j, 176).Value & " | " & "XXXMud (" & ThisWorkbook.Worksheets(4).Cells(j, 33).Value & ")"
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

End Sub

Sub Rapor2_2SonucTeslimTutanaklari()
Dim i As Integer, TarihBul As Range
Dim Say As Long, j As Long, Cont As Long, Sno As Variant
Dim LstBx As MSForms.ListBox, ctl As MSForms.Control
Dim LblSira1 As MSForms.Label
Dim LblSira2 As MSForms.Label
Dim Contx As Long

CalTarih = CalTarihTakip

Say = ThisWorkbook.Worksheets(4).Range("CN100000").End(xlUp).Row
If Say < 7 Then
    GoTo Son
End If

'Rapor2_2 sonuç
Contx = 0
TarihiTekrarla3:
Set TarihBul = ThisWorkbook.Worksheets(4).Range("FU6:FU100000").Find(What:=CalTarih, SearchDirection:=xlNext, _
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
    GoTo TarihiTekrarla3
End If

Cont = ContTakip
'Cont = 0

For j = Say To TarihBul.Row Step -1
    'If Cont < 50 Then
        If ThisWorkbook.Worksheets(4).Cells(j, 5).Value <> "" And ThisWorkbook.Worksheets(4).Cells(j, 177).Value <> "" Then 'ThisWorkbook.Worksheets(4).Cells(j, 83).Value <> "" And ThisWorkbook.Worksheets(4).Cells(j, 177).Value <> "" Then
            Cont = Cont + 1
            'Bilgilendirme
            If ThisWorkbook.Worksheets(4).Cells(j, 200).Value <> "" Then
                If Left(ThisWorkbook.Worksheets(4).Cells(j, 199).Value, 3) = "İl " Then
                    Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | B | " & ThisWorkbook.Worksheets(4).Cells(j, 178).Value & " | " & "Sonuç (" & ThisWorkbook.Worksheets(4).Cells(j, 200).Value & ")"
                ElseIf Left(ThisWorkbook.Worksheets(4).Cells(j, 199).Value, 3) = "İlç" Then
                    Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | B | " & ThisWorkbook.Worksheets(4).Cells(j, 178).Value & " | " & "Sonuç (" & ThisWorkbook.Worksheets(4).Cells(j, 200).Value & ")"
                Else
                    Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | B | " & ThisWorkbook.Worksheets(4).Cells(j, 178).Value & " | " & "Sonuç (" & ThisWorkbook.Worksheets(4).Cells(j, 199).Value & " " & ThisWorkbook.Worksheets(4).Cells(j, 200).Value & ")"
                End If
            ElseIf ThisWorkbook.Worksheets(4).Cells(j, 200).Value = "" Then
                If Left(ThisWorkbook.Worksheets(4).Cells(j, 199).Value, 3) = "İl " Then
                    Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | B | " & ThisWorkbook.Worksheets(4).Cells(j, 178).Value & " | " & "Sonuç (" & ThisWorkbook.Worksheets(4).Cells(j, 203).Value & " " & ThisWorkbook.Worksheets(4).Cells(j, 199).Value & ")"
                ElseIf Left(ThisWorkbook.Worksheets(4).Cells(j, 199).Value, 3) = "İlç" Then
                    Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | B | " & ThisWorkbook.Worksheets(4).Cells(j, 178).Value & " | " & "Sonuç (" & ThisWorkbook.Worksheets(4).Cells(j, 204).Value & " " & ThisWorkbook.Worksheets(4).Cells(j, 199).Value & ")"
                Else
                    Sno = ThisWorkbook.Worksheets(4).Cells(j, 5).Value & " | B | " & ThisWorkbook.Worksheets(4).Cells(j, 178).Value & " | " & "Sonuç (" & ThisWorkbook.Worksheets(4).Cells(j, 199).Value
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

End Sub

Sub Rapor2_2Options()
Dim ctl As MSForms.Control

core_report2_entry_UI.Rapor2_2KararFrame.Visible = True

'___________________________

'Tümü seçili iken birleştirme pasif
If core_report2_entry_UI.TumuOption.Value = True Then
    core_report2_entry_UI.Tutanak2BirlestirmeOptionEvet.Enabled = False
    core_report2_entry_UI.Tutanak2BirlestirmeOptionHayir.Enabled = False
    core_report2_entry_UI.Tutanak2BirlestirmeOptionEvet.Value = False
    core_report2_entry_UI.Tutanak2BirlestirmeOptionHayir.Value = False
End If
'Technique A seçili ve tutanak1 yapılacaksa birleştirme aktif
If core_report2_entry_UI.SadeceTeknik_AGecersizlerOption.Value = True And core_report2_entry_UI.XXXMudPaketiOptionEvet.Value = True Then
    core_report2_entry_UI.Tutanak2BirlestirmeOptionEvet.Enabled = True
    core_report2_entry_UI.Tutanak2BirlestirmeOptionHayir.Enabled = True
End If
'Technique A seçili ve tutanak1 yapılmayacaksa birleştirme pasif
If core_report2_entry_UI.SadeceTeknik_AGecersizlerOption.Value = True And core_report2_entry_UI.XXXMudPaketiOptionHayir.Value = True Then
    core_report2_entry_UI.Tutanak2BirlestirmeOptionEvet.Enabled = False
    core_report2_entry_UI.Tutanak2BirlestirmeOptionHayir.Enabled = False
    core_report2_entry_UI.Tutanak2BirlestirmeOptionEvet.Value = False
    core_report2_entry_UI.Tutanak2BirlestirmeOptionHayir.Value = False
End If

'___________________________

'Tümü seçili ve tutanak1 yapılacaksa
If core_report2_entry_UI.TumuOption.Value = True And core_report2_entry_UI.XXXMudPaketiOptionEvet.Value = True Then
    'core_report2_entry_UI.Rapor2_2KararFrame.Visible = True
    core_report2_entry_UI.Rapor1Frame.Visible = True
    core_report2_entry_UI.XXXMudTutanak2Frame.Visible = True
    core_report2_entry_UI.XXXMudUstYaziFrame.Visible = True
    core_report2_entry_UI.GelenXXXMudTutanak1Frame.Visible = True
    core_report2_entry_UI.GelenXXXMudTutanak2Frame.Visible = False 'Kapatıldı
    core_report2_entry_UI.Tutanak2Frame.Visible = True
    core_report2_entry_UI.UstYaziFrame.Visible = True
    core_report2_entry_UI.Rapor2_2UstYaziFrame.Visible = True
    
    core_report2_entry_UI.Rapor2_2KararFrame.Top = 516
    core_report2_entry_UI.XXXMudTutanak2Frame.Top = 576 '379 '399 '339
    core_report2_entry_UI.XXXMudUstYaziFrame.Top = 636 '419 '469 '409
    core_report2_entry_UI.UstYaziFrame.Top = 696 '459 '509 '449
    core_report2_entry_UI.GelenXXXMudTutanak1Frame.Top = 786 '519 '499
    'core_report2_entry_UI.GelenXXXMudTutanak2Frame.Top = 559 '499 '549 '489
    core_report2_entry_UI.Tutanak2Frame.Top = 876 '579 '559 '619 '559 '619 '559
    core_report2_entry_UI.Rapor2_2UstYaziFrame.Top = 936 '619 '679 '619 '689 '629
    core_report2_entry_UI.AltMenuFrame.Top = 1032 '679 '739 '679 '749 '689
    core_report2_entry_UI.TasiyiciFrame.Height = 1052
    
    'Genel front (TasiyiciFrame'den sonra olmazsa çalışmaz.)
    core_report2_entry_UI.Rapor2_2KararFrame.ZOrder msoBringToFront
    core_report2_entry_UI.Rapor1Frame.ZOrder msoBringToFront
    core_report2_entry_UI.XXXMudTutanak2Frame.ZOrder msoBringToFront
    core_report2_entry_UI.XXXMudUstYaziFrame.ZOrder msoBringToFront
    core_report2_entry_UI.GelenXXXMudTutanak1Frame.ZOrder msoBringToFront
    core_report2_entry_UI.GelenXXXMudTutanak2Frame.ZOrder msoBringToFront
    core_report2_entry_UI.Tutanak2Frame.ZOrder msoBringToFront
    core_report2_entry_UI.UstYaziFrame.ZOrder msoBringToFront
    core_report2_entry_UI.Rapor2_2UstYaziFrame.ZOrder msoBringToFront
    
    'Ekrana göre formun ayarlanması
    If EkranKontrol = True Then '760 Then '710 Then '572 Then
        RepYukseklik = 1092
        core_report2_entry_UI.AltMenuFrame.Top = RepYukseklik - 40 - 26
        core_report2_entry_UI.TasiyiciFrame.Height = RepYukseklik - 40
        core_report2_entry_UI.Height = 485
        'core_report2_entry_UI.Width = 1072 '979
        core_report2_entry_UI.ScrollBars = fmScrollBarsVertical
        core_report2_entry_UI.ScrollHeight = RepYukseklik '- (RepYukseklik - Application.Height) '572 '549
        'core_report2_entry_UI.ScrollTop = 12 '549 '449
        'core_report2_entry_UI.Height = Application.Height
    Else
        RepYukseklik = 1092
        core_report2_entry_UI.Height = 716
        'core_report2_entry_UI.Width = 1072 '979
        core_report2_entry_UI.ScrollBars = fmScrollBarsVertical
        core_report2_entry_UI.ScrollHeight = RepYukseklik '- (RepYukseklik - Application.Height) '572 '549
        'core_report2_entry_UI.ScrollTop = 12 '549 '449
        'core_report2_entry_UI.Height = Application.Height
    End If
    'core_report2_entry_UI.IlgiYaziFotokopisiFrame.Visible = False
    'core_report2_entry_UI.UstYaziNotuFrame.Visible = False
End If

'___________________________



'Tümü seçili ve tutanak1 yapılmayacaksa
If core_report2_entry_UI.TumuOption.Value = True And core_report2_entry_UI.XXXMudPaketiOptionHayir.Value = True Then
'    core_report2_entry_UI.Rapor2_2KararFrame.Visible = True
    core_report2_entry_UI.Rapor1Frame.Visible = True
    core_report2_entry_UI.XXXMudTutanak2Frame.Visible = True
    core_report2_entry_UI.XXXMudUstYaziFrame.Visible = True
    core_report2_entry_UI.GelenXXXMudTutanak1Frame.Visible = False 'Kapatıldı
    core_report2_entry_UI.GelenXXXMudTutanak2Frame.Visible = True 'Sadece Paket Tipi, Paket Adedi gösterilecek ve Tutanak2 Sayfası ilave edilecek.
    core_report2_entry_UI.Tutanak2Frame.Visible = False 'Kapatıldı.
    core_report2_entry_UI.UstYaziFrame.Visible = True
    core_report2_entry_UI.Rapor2_2UstYaziFrame.Visible = True
    
    'XXXMud tutanak2 bölümünü düzenle
    core_report2_entry_UI.GelenXXXMudTutanak2TutSayfaFrame.Visible = True
    core_report2_entry_UI.GelenXXXMudTutanak2TarihiFrame.Visible = False
    core_report2_entry_UI.GelenXXXMudTutanak2Imza1Frame.Visible = False
    core_report2_entry_UI.GelenXXXMudTutanak2Imza2Frame.Visible = False
    
    core_report2_entry_UI.Rapor2_2KararFrame.Top = 516
    core_report2_entry_UI.XXXMudTutanak2Frame.Top = 576 '379 '399 '339
    core_report2_entry_UI.XXXMudUstYaziFrame.Top = 636 '419 '469 '409
    core_report2_entry_UI.UstYaziFrame.Top = 696 '459 '509 '449
    'core_report2_entry_UI.GelenXXXMudTutanak1Frame.Top = 499
    core_report2_entry_UI.GelenXXXMudTutanak2Frame.Top = 786 '519 '499 '559 '499 '549 '489
    'core_report2_entry_UI.Tutanak2Frame.Top = 559 '619 '559 '619 '559
    core_report2_entry_UI.Rapor2_2UstYaziFrame.Top = 846 '559 '619 '679 '619 '689 '629
    core_report2_entry_UI.AltMenuFrame.Top = 942 '619 '679 '739 '679 '749 '689
    core_report2_entry_UI.TasiyiciFrame.Height = 962

    'Genel front (TasiyiciFrame'den sonra olmazsa çalışmaz.)
    core_report2_entry_UI.Rapor2_2KararFrame.ZOrder msoBringToFront
    core_report2_entry_UI.Rapor1Frame.ZOrder msoBringToFront
    core_report2_entry_UI.XXXMudTutanak2Frame.ZOrder msoBringToFront
    core_report2_entry_UI.XXXMudUstYaziFrame.ZOrder msoBringToFront
    core_report2_entry_UI.GelenXXXMudTutanak1Frame.ZOrder msoBringToFront
    core_report2_entry_UI.GelenXXXMudTutanak2Frame.ZOrder msoBringToFront
    core_report2_entry_UI.Tutanak2Frame.ZOrder msoBringToFront
    core_report2_entry_UI.UstYaziFrame.ZOrder msoBringToFront
    core_report2_entry_UI.Rapor2_2UstYaziFrame.ZOrder msoBringToFront

    'Ekrana göre formun ayarlanması
    If EkranKontrol = True Then '760 Then '710 Then '572 Then
        RepYukseklik = 1002
        core_report2_entry_UI.AltMenuFrame.Top = RepYukseklik - 40 - 26
        core_report2_entry_UI.TasiyiciFrame.Height = RepYukseklik - 40
        core_report2_entry_UI.Height = 485
        'core_report2_entry_UI.Width = 1072 '979
        core_report2_entry_UI.ScrollBars = fmScrollBarsVertical
        core_report2_entry_UI.ScrollHeight = RepYukseklik '- (RepYukseklik - Application.Height) '572 '549
        'core_report2_entry_UI.ScrollTop = 12 '549 '449
        'core_report2_entry_UI.Height = Application.Height
    Else
        RepYukseklik = 1002
        core_report2_entry_UI.Height = 716
        'core_report2_entry_UI.Width = 1072 '979
        core_report2_entry_UI.ScrollBars = fmScrollBarsVertical
        core_report2_entry_UI.ScrollHeight = RepYukseklik '- (RepYukseklik - Application.Height) '572 '549
        'core_report2_entry_UI.ScrollTop = 12 '549 '449
        'core_report2_entry_UI.Height = Application.Height
    End If
'    core_report2_entry_UI.IlgiYaziFotokopisiFrame.Visible = False
'    core_report2_entry_UI.UstYaziNotuFrame.Visible = False
End If

'___________________________

'Sadece Technique A seçili ve tutanak1 yapılacaksa
If core_report2_entry_UI.SadeceTeknik_AGecersizlerOption.Value = True And core_report2_entry_UI.XXXMudPaketiOptionEvet.Value = True Then
    If core_report2_entry_UI.Tutanak2BirlestirmeOptionEvet.Value = False And core_report2_entry_UI.Tutanak2BirlestirmeOptionHayir.Value = False Then
        core_report2_entry_UI.Rapor1Frame.Visible = True
        'core_report2_entry_UI.Rapor2_2KararFrame.Visible = True
        core_report2_entry_UI.XXXMudTutanak2Frame.Visible = False
        core_report2_entry_UI.XXXMudUstYaziFrame.Visible = False
        core_report2_entry_UI.GelenXXXMudTutanak1Frame.Visible = False
        core_report2_entry_UI.GelenXXXMudTutanak2Frame.Visible = False 'Kapatıldı
        core_report2_entry_UI.Rapor2_2UstYaziFrame.Visible = False
    
        core_report2_entry_UI.UstYaziFrame.Visible = False
        core_report2_entry_UI.Tutanak2Frame.Visible = False
        
        core_report2_entry_UI.Rapor2_2KararFrame.Top = 516
        core_report2_entry_UI.AltMenuFrame.Top = 582 '739 '679 '749 '689
        core_report2_entry_UI.TasiyiciFrame.Height = 606
        RepYukseklik = 674 '431
        core_report2_entry_UI.Height = RepYukseklik '430 '400 '760
        
        core_report2_entry_UI.ScrollTop = 0
        core_report2_entry_UI.ScrollHeight = 0
        core_report2_entry_UI.ScrollBars = fmScrollBarsNone
    End If
    
    'Tutanak2 tutanakları birleştirilecek
    If core_report2_entry_UI.Tutanak2BirlestirmeOptionEvet.Value = True Then
        'Rapor2_2KararFrame.Visible = True
        core_report2_entry_UI.Rapor1Frame.Visible = True
        core_report2_entry_UI.XXXMudTutanak2Frame.Visible = True
        core_report2_entry_UI.XXXMudUstYaziFrame.Visible = True
        core_report2_entry_UI.GelenXXXMudTutanak1Frame.Visible = True
        core_report2_entry_UI.GelenXXXMudTutanak2Frame.Visible = False 'Kapatıldı
        core_report2_entry_UI.Tutanak2Frame.Visible = True
        core_report2_entry_UI.UstYaziFrame.Visible = True
        core_report2_entry_UI.Rapor2_2UstYaziFrame.Visible = True

        core_report2_entry_UI.Rapor2_2KararFrame.Top = 516
        core_report2_entry_UI.XXXMudTutanak2Frame.Top = 576 '379 '399 '339
        core_report2_entry_UI.XXXMudUstYaziFrame.Top = 636 '419 '469 '409
        core_report2_entry_UI.UstYaziFrame.Top = 696 '459 '509 '449
        core_report2_entry_UI.GelenXXXMudTutanak1Frame.Top = 786 '519 '499
        core_report2_entry_UI.Tutanak2Frame.Top = 876 '579 '559 '619 '559 '619 '559
        core_report2_entry_UI.Rapor2_2UstYaziFrame.Top = 936 '619 '679 '619 '689 '629
        core_report2_entry_UI.AltMenuFrame.Top = 1032 '679 '739 '679 '749 '689
        core_report2_entry_UI.TasiyiciFrame.Height = 1052

        'Genel front (TasiyiciFrame'den sonra olmazsa çalışmaz.)
        core_report2_entry_UI.Rapor2_2KararFrame.ZOrder msoBringToFront
        core_report2_entry_UI.Rapor1Frame.ZOrder msoBringToFront
        core_report2_entry_UI.XXXMudTutanak2Frame.ZOrder msoBringToFront
        core_report2_entry_UI.XXXMudUstYaziFrame.ZOrder msoBringToFront
        core_report2_entry_UI.GelenXXXMudTutanak1Frame.ZOrder msoBringToFront
        core_report2_entry_UI.GelenXXXMudTutanak2Frame.ZOrder msoBringToFront
        core_report2_entry_UI.Tutanak2Frame.ZOrder msoBringToFront
        core_report2_entry_UI.UstYaziFrame.ZOrder msoBringToFront
        core_report2_entry_UI.Rapor2_2UstYaziFrame.ZOrder msoBringToFront
    
        'Ekrana göre formun ayarlanması
        If EkranKontrol = True Then '760 Then '710 Then '572 Then
            RepYukseklik = 1092
            core_report2_entry_UI.AltMenuFrame.Top = RepYukseklik - 40 - 26
            core_report2_entry_UI.TasiyiciFrame.Height = RepYukseklik - 40
            core_report2_entry_UI.Height = 485
            core_report2_entry_UI.ScrollBars = fmScrollBarsVertical
            core_report2_entry_UI.ScrollHeight = RepYukseklik '- (RepYukseklik - Application.Height) '572 '549
        Else
            RepYukseklik = 1092
            core_report2_entry_UI.Height = 716
            core_report2_entry_UI.ScrollBars = fmScrollBarsVertical
            core_report2_entry_UI.ScrollHeight = RepYukseklik '- (RepYukseklik - Application.Height) '572 '549
        End If
    End If
    
    'Tutanak2 tutanakları birleştirilmeyecek
    If core_report2_entry_UI.Tutanak2BirlestirmeOptionHayir.Value = True Then
        'core_report2_entry_UI.Rapor2_2KararFrame.Visible = True
        core_report2_entry_UI.Rapor1Frame.Visible = True
        core_report2_entry_UI.XXXMudTutanak2Frame.Visible = True
        core_report2_entry_UI.XXXMudUstYaziFrame.Visible = True
        core_report2_entry_UI.GelenXXXMudTutanak1Frame.Visible = True
        core_report2_entry_UI.GelenXXXMudTutanak2Frame.Visible = True
        core_report2_entry_UI.Tutanak2Frame.Visible = True
        core_report2_entry_UI.UstYaziFrame.Visible = True
        core_report2_entry_UI.Rapor2_2UstYaziFrame.Visible = True

        'XXXMud tutanak2 bölümünü düzenle
        core_report2_entry_UI.GelenXXXMudTutanak2TutSayfaFrame.Visible = False
        core_report2_entry_UI.GelenXXXMudTutanak2TarihiFrame.Visible = True
        core_report2_entry_UI.GelenXXXMudTutanak2Imza1Frame.Visible = True
        core_report2_entry_UI.GelenXXXMudTutanak2Imza2Frame.Visible = True

        core_report2_entry_UI.Rapor2_2KararFrame.Top = 516
        core_report2_entry_UI.XXXMudTutanak2Frame.Top = 576 '379 '399 '339
        core_report2_entry_UI.XXXMudUstYaziFrame.Top = 636 '419 '469 '409
        core_report2_entry_UI.UstYaziFrame.Top = 696 '459 '509 '449
        core_report2_entry_UI.GelenXXXMudTutanak1Frame.Top = 786 '519 '499
        core_report2_entry_UI.GelenXXXMudTutanak2Frame.Top = 876 '559 '499 '549 '489
        core_report2_entry_UI.Tutanak2Frame.Top = 936 '579 '559 '619 '559 '619 '559
        core_report2_entry_UI.Rapor2_2UstYaziFrame.Top = 996 '619 '679 '619 '689 '629
        core_report2_entry_UI.AltMenuFrame.Top = 1092 '679 '739 '679 '749 '689
        core_report2_entry_UI.TasiyiciFrame.Height = 1112

        'Genel front (TasiyiciFrame'den sonra olmazsa çalışmaz.)
        core_report2_entry_UI.Rapor2_2KararFrame.ZOrder msoBringToFront
        core_report2_entry_UI.Rapor1Frame.ZOrder msoBringToFront
        core_report2_entry_UI.XXXMudTutanak2Frame.ZOrder msoBringToFront
        core_report2_entry_UI.XXXMudUstYaziFrame.ZOrder msoBringToFront
        core_report2_entry_UI.GelenXXXMudTutanak1Frame.ZOrder msoBringToFront
        core_report2_entry_UI.GelenXXXMudTutanak2Frame.ZOrder msoBringToFront
        core_report2_entry_UI.Tutanak2Frame.ZOrder msoBringToFront
        core_report2_entry_UI.UstYaziFrame.ZOrder msoBringToFront
        core_report2_entry_UI.Rapor2_2UstYaziFrame.ZOrder msoBringToFront
    
        'Ekrana göre formun ayarlanması
        If EkranKontrol = True Then '760 Then '710 Then '572 Then
            RepYukseklik = 1152
            core_report2_entry_UI.AltMenuFrame.Top = RepYukseklik - 40 - 26
            core_report2_entry_UI.TasiyiciFrame.Height = RepYukseklik - 40
            core_report2_entry_UI.Height = 485
            core_report2_entry_UI.ScrollBars = fmScrollBarsVertical
            core_report2_entry_UI.ScrollHeight = RepYukseklik '- (RepYukseklik - Application.Height) '572 '549
        Else
            RepYukseklik = 1152
            core_report2_entry_UI.Height = 716
            core_report2_entry_UI.ScrollBars = fmScrollBarsVertical
            core_report2_entry_UI.ScrollHeight = RepYukseklik '- (RepYukseklik - Application.Height) '572 '549
        End If
    End If
End If

'___________________________

'Sadece Technique A seçili ve tutanak1 yapılmayacaksa
If core_report2_entry_UI.SadeceTeknik_AGecersizlerOption.Value = True And core_report2_entry_UI.XXXMudPaketiOptionHayir.Value = True Then
    core_report2_entry_UI.Rapor1Frame.Visible = True
    core_report2_entry_UI.XXXMudTutanak2Frame.Visible = True
    core_report2_entry_UI.XXXMudUstYaziFrame.Visible = True
    core_report2_entry_UI.GelenXXXMudTutanak1Frame.Visible = False 'Kapatıldı
    core_report2_entry_UI.GelenXXXMudTutanak2Frame.Visible = True 'Sadece Paket Tipi, Paket Adedi gösterilecek ve Tutanak2 Sayfası ilave edilecek.
    core_report2_entry_UI.Tutanak2Frame.Visible = True
    core_report2_entry_UI.UstYaziFrame.Visible = True
    core_report2_entry_UI.Rapor2_2UstYaziFrame.Visible = True

    'XXXMud tutanak2 bölümünü düzenle
    core_report2_entry_UI.GelenXXXMudTutanak2TutSayfaFrame.Visible = True
    core_report2_entry_UI.GelenXXXMudTutanak2TarihiFrame.Visible = False
    core_report2_entry_UI.GelenXXXMudTutanak2Imza1Frame.Visible = False
    core_report2_entry_UI.GelenXXXMudTutanak2Imza2Frame.Visible = False
    
    core_report2_entry_UI.Rapor2_2KararFrame.Top = 516
    core_report2_entry_UI.XXXMudTutanak2Frame.Top = 576 '379 '399 '339
    core_report2_entry_UI.XXXMudUstYaziFrame.Top = 636 '419 '469 '409
    core_report2_entry_UI.UstYaziFrame.Top = 696 '459 '509 '449
    core_report2_entry_UI.GelenXXXMudTutanak2Frame.Top = 786 '559 '499 '549 '489
    core_report2_entry_UI.Tutanak2Frame.Top = 846 '579 '559 '619 '559 '619 '559
    core_report2_entry_UI.Rapor2_2UstYaziFrame.Top = 906 '619 '679 '619 '689 '629
    core_report2_entry_UI.AltMenuFrame.Top = 1002 '679 '739 '679 '749 '689
    core_report2_entry_UI.TasiyiciFrame.Height = 1022

    'Genel front (TasiyiciFrame'den sonra olmazsa çalışmaz.)
    core_report2_entry_UI.Rapor2_2KararFrame.ZOrder msoBringToFront
    core_report2_entry_UI.Rapor1Frame.ZOrder msoBringToFront
    core_report2_entry_UI.XXXMudTutanak2Frame.ZOrder msoBringToFront
    core_report2_entry_UI.XXXMudUstYaziFrame.ZOrder msoBringToFront
    core_report2_entry_UI.GelenXXXMudTutanak1Frame.ZOrder msoBringToFront
    core_report2_entry_UI.GelenXXXMudTutanak2Frame.ZOrder msoBringToFront
    core_report2_entry_UI.Tutanak2Frame.ZOrder msoBringToFront
    core_report2_entry_UI.UstYaziFrame.ZOrder msoBringToFront
    core_report2_entry_UI.Rapor2_2UstYaziFrame.ZOrder msoBringToFront

    'Ekrana göre formun ayarlanması
    If EkranKontrol = True Then '760 Then '710 Then '572 Then
        RepYukseklik = 1062
        core_report2_entry_UI.AltMenuFrame.Top = RepYukseklik - 40 - 26
        core_report2_entry_UI.TasiyiciFrame.Height = RepYukseklik - 40
        core_report2_entry_UI.Height = 485
        core_report2_entry_UI.ScrollBars = fmScrollBarsVertical
        core_report2_entry_UI.ScrollHeight = RepYukseklik '- (RepYukseklik - Application.Height) '572 '549
    Else
        RepYukseklik = 1062
        core_report2_entry_UI.Height = 716
        core_report2_entry_UI.ScrollBars = fmScrollBarsVertical
        core_report2_entry_UI.ScrollHeight = RepYukseklik '- (RepYukseklik - Application.Height) '572 '549
    End If
End If

'___________________________

'XXXMud'den yazının geliş tarihinin sonuca ilavesi
If core_report2_entry_UI.GelenXXXMudTutanak1Frame.Visible = False Then
    core_report2_entry_UI.SonucGelenXXXMudGelisTarihiFrame.Visible = True
Else
    core_report2_entry_UI.SonucGelenXXXMudGelisTarihiFrame.Visible = False
End If


If (core_report2_entry_UI.TumuOption.Value = False And core_report2_entry_UI.SadeceTeknik_AGecersizlerOption.Value = False And _
    core_report2_entry_UI.XXXMudPaketiOptionEvet.Value = False And core_report2_entry_UI.XXXMudPaketiOptionHayir.Value = False And _
    core_report2_entry_UI.Tutanak2BirlestirmeOptionEvet.Value = False And core_report2_entry_UI.Tutanak2BirlestirmeOptionHayir.Value = False) Or _
   (core_report2_entry_UI.TumuOption.Value = True And core_report2_entry_UI.SadeceTeknik_AGecersizlerOption.Value = False And _
    core_report2_entry_UI.XXXMudPaketiOptionEvet.Value = False And core_report2_entry_UI.XXXMudPaketiOptionHayir.Value = False And _
    core_report2_entry_UI.Tutanak2BirlestirmeOptionEvet.Value = False And core_report2_entry_UI.Tutanak2BirlestirmeOptionHayir.Value = False) Or _
   (core_report2_entry_UI.TumuOption.Value = False And core_report2_entry_UI.SadeceTeknik_AGecersizlerOption.Value = True And _
    core_report2_entry_UI.XXXMudPaketiOptionEvet.Value = False And core_report2_entry_UI.XXXMudPaketiOptionHayir.Value = False And _
    core_report2_entry_UI.Tutanak2BirlestirmeOptionEvet.Value = False And core_report2_entry_UI.Tutanak2BirlestirmeOptionHayir.Value = False) Or _
   (core_report2_entry_UI.TumuOption.Value = False And core_report2_entry_UI.SadeceTeknik_AGecersizlerOption.Value = False And _
    core_report2_entry_UI.XXXMudPaketiOptionEvet.Value = True And core_report2_entry_UI.XXXMudPaketiOptionHayir.Value = False And _
    core_report2_entry_UI.Tutanak2BirlestirmeOptionEvet.Value = False And core_report2_entry_UI.Tutanak2BirlestirmeOptionHayir.Value = False) Or _
   (core_report2_entry_UI.TumuOption.Value = False And core_report2_entry_UI.SadeceTeknik_AGecersizlerOption.Value = False And _
    core_report2_entry_UI.XXXMudPaketiOptionEvet.Value = False And core_report2_entry_UI.XXXMudPaketiOptionHayir.Value = True And _
    core_report2_entry_UI.Tutanak2BirlestirmeOptionEvet.Value = False And core_report2_entry_UI.Tutanak2BirlestirmeOptionHayir.Value = False) Then

    core_report2_entry_UI.UstYaziFrame.Caption = "Informative Cover Letter Entry"
    core_report2_entry_UI.Tutanak2Frame.Caption = "Statement 2 Entry"
    
    core_report2_entry_UI.Rapor1Frame.Visible = True
    'core_report2_entry_UI.Rapor2_2KararFrame.Visible = True
    core_report2_entry_UI.XXXMudTutanak2Frame.Visible = False
    core_report2_entry_UI.XXXMudUstYaziFrame.Visible = False
    core_report2_entry_UI.GelenXXXMudTutanak1Frame.Visible = False
    core_report2_entry_UI.GelenXXXMudTutanak2Frame.Visible = False
    core_report2_entry_UI.Rapor2_2UstYaziFrame.Visible = False

    core_report2_entry_UI.UstYaziFrame.Visible = False
    core_report2_entry_UI.Tutanak2Frame.Visible = False


    For Each ctl In core_report2_entry_UI.UstMenuFrame.Controls
        If TypeName(ctl) = "Label" Then
            ctl.BackColor = RGB(225, 235, 245) 'RGB(201, 216, 230)
            ctl.ForeColor = RGB(30, 30, 30)
        End If
    Next ctl
    
    core_report2_entry_UI.UstYaziGirisi.BackColor = RGB(180, 210, 240)
    core_report2_entry_UI.UstYaziGirisi.ForeColor = RGB(30, 30, 30)

    'Ekrana göre formun ayarlanması
    If EkranKontrol = True Then '760 Then '710 Then '572 Then
        RepYukseklik = 648 '431
        core_report2_entry_UI.AltMenuFrame.Top = RepYukseklik - 40 - 26
        core_report2_entry_UI.TasiyiciFrame.Height = RepYukseklik - 40
        core_report2_entry_UI.Height = 485

        core_report2_entry_UI.ScrollBars = fmScrollBarsVertical
        core_report2_entry_UI.ScrollHeight = RepYukseklik '- (RepYukseklik - Application.Height) '572 '549
    Else
        core_report2_entry_UI.Rapor2_2KararFrame.Top = 516
        core_report2_entry_UI.AltMenuFrame.Top = 582 '739 '679 '749 '689
        core_report2_entry_UI.TasiyiciFrame.Height = 606
        RepYukseklik = 674 '431
        core_report2_entry_UI.Height = RepYukseklik '430 '400 '760
        
        core_report2_entry_UI.ScrollTop = 0
        core_report2_entry_UI.ScrollHeight = 0
        core_report2_entry_UI.ScrollBars = fmScrollBarsNone
    End If
    
End If

End Sub


