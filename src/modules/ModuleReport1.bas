Attribute VB_Name = "ModuleReport1"
Option Explicit
Public OpenWordTakip As Boolean
Public OpenWordSay As Integer
Public ContTakip As Long
Public MaxiAktar As Integer
Public FarkAktar As Integer
Public YeniIslemAktar As Long
Public IlkSiraAktar As Long
Public SonSiraAktar As Long
Public ContTakipGiris As Long
Public ContTakipMevcut As Long
Public ContTakipCikis As Long
Dim IlceSakla As String
Public StrTime As String, IslemGunluguTarih As String, IslemGunluguAyrac As String, ModulTarih As String, ModulAyrac As String
Public Say1IslemGunlugu As Long, Say2IslemGunlugu As Long, SayAyracIslemGunlugu As Long, SayMax As Long, SayGenel As Long, SayDonem As Long
Public ilkrow As Long, sonrow As Long, donemrow As Long, YeniDonemAltRow As Long, ilkrowx As Long, sonrowx As Long

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

Sub Rapor1Tutanak1()

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
Dim TipB, TipA As Boolean

Dim IlkSiraBulx As Range, SonSiraBulx As Range, IlkSirax As Long, SonSirax As Long, WsFarkGirisRapor1 As Worksheet


IlceSakla = ""
If InStr(Cells(ActiveCell.Row, 18).Value, " Organization A") <> 0 Then
    IlceSakla = Cells(ActiveCell.Row, 18).Value
    Cells(ActiveCell.Row, 18).Value = ""
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
    'RAPOR TANIMLARI
    SourceRaporNormal = AutoPath & "\System Files\System Templates\Report 1 Template\Report 1.docm"
    'TUTANAK2 TANIMLARI
    SourceTutanak2Normal = AutoPath & "\System Files\System Templates\Statement 2 Template\Statement 2.docm"
    'ÜST YAZI TANIMLARI
    SourceUstYaziNormal = AutoPath & "\System Files\System Templates\Report 1 Cover Letter Template\Report 1 Cover Letter.docm"
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
    Call ModuleReport1.OpenWordControl
    
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

'    Call ModuleReport1.OpenWordControl
     
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
    If Cells(ActiveCell.Row, 32).Value = "d. Discrepancy Detected" Then
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
    
    If Cells(ActiveCell.Row, 32).Value = "d. Discrepancy Detected" Then 'Farklı tutanak1 tutanağı
        'Birim
        Birimx = UCase(Replace(Replace(Worksheets(2).Cells(6, 99).Value, "i", "I"), "ı", "I")) & " UNIT"
        objDoc.Tables(1).Cell(Row:=1, Column:=1).Range.Text = Birimx
        'Tutanak1 tarihi
        objDoc.Tables(1).Cell(Row:=4, Column:=3).Range.Text = Cells(ActiveCell.Row, 31).Value
        'Geliş tarihi
        objDoc.Tables(1).Cell(Row:=5, Column:=3).Range.Text = Cells(ActiveCell.Row, 28).Value
        'By Hand/By Mail
        If Cells(ActiveCell.Row, 30).Value = "By Hand" Then
            objDoc.CheckElden.Value = True
        ElseIf Cells(ActiveCell.Row, 30).Value = "By Mail" Then
            objDoc.CheckPosta.Value = True
        End If
        'Türkçe karakterleri düzelt
        objDoc.CheckElden.Enabled = False
        objDoc.CheckElden.Enabled = True
        objDoc.CheckPosta.Enabled = False
        objDoc.CheckPosta.Enabled = True
        
        'Gelen Yazının Tarih ve Sayısı
        objDoc.Tables(1).Cell(Row:=6, Column:=3).Range.Text = Cells(ActiveCell.Row, 20).Value
        objDoc.Tables(1).Cell(Row:=6, Column:=5).Range.Text = Cells(ActiveCell.Row, 21).Value
        'Gönderen
        If Cells(ActiveCell.Row, 25).Value = "Provincial Directorate B" Then
            GelenTema = Cells(ActiveCell.Row, 17).Value & " Provincial Governorship Provincial Directorate B " & Cells(ActiveCell.Row, 26).Value
        ElseIf Cells(ActiveCell.Row, 25).Value = "District Directorate B" Then
            GelenTema = Cells(ActiveCell.Row, 18).Value & " District Governorship District Directorate B " & Cells(ActiveCell.Row, 26).Value
        ElseIf Cells(ActiveCell.Row, 25).Value = "Provincial Directorate C" Then
            GelenTema = Cells(ActiveCell.Row, 17).Value & " Provincial Governorship Provincial Directorate C " & Cells(ActiveCell.Row, 26).Value
        ElseIf Cells(ActiveCell.Row, 25).Value = "District Directorate C" Then
            GelenTema = Cells(ActiveCell.Row, 18).Value & " District Governorship District Directorate C " & Cells(ActiveCell.Row, 26).Value
        ElseIf Cells(ActiveCell.Row, 25).Value = "Provincial Directorate D" Then
            GelenTema = Cells(ActiveCell.Row, 17).Value & " Provincial Governorship Provincial Directorate D " & Cells(ActiveCell.Row, 26).Value
        ElseIf Cells(ActiveCell.Row, 25).Value = "District Directorate D" Then
            GelenTema = Cells(ActiveCell.Row, 18).Value & " District Governorship District Directorate D " & Cells(ActiveCell.Row, 26).Value
        ElseIf Cells(ActiveCell.Row, 25).Value = "Provincial Directorate E" Then
            GelenTema = Cells(ActiveCell.Row, 17).Value & " Provincial Governorship Provincial Directorate E " & Cells(ActiveCell.Row, 26).Value
        ElseIf Cells(ActiveCell.Row, 25).Value = "District Directorate E" Then
            GelenTema = Cells(ActiveCell.Row, 18).Value & " District Governorship District Directorate E " & Cells(ActiveCell.Row, 26).Value
        ElseIf InStr(Cells(ActiveCell.Row, 25).Value, "General Directorate") <> 0 Or InStr(Cells(ActiveCell.Row, 25).Value, "Regional Directorate") <> 0 Then
            If Cells(ActiveCell.Row, 26).Value <> "" Then
                GelenTema = WorksheetFunction.Proper(ThisWorkbook.Worksheets(2).Cells(6, 111).Value) & " " & Cells(ActiveCell.Row, 25).Value & " " & Cells(ActiveCell.Row, 26).Value
            Else
                GelenTema = WorksheetFunction.Proper(ThisWorkbook.Worksheets(2).Cells(6, 111).Value) & " " & Cells(ActiveCell.Row, 25).Value
            End If
        Else
            If Cells(ActiveCell.Row, 26).Value <> "" Then
                If InStr(Cells(ActiveCell.Row, 25).Value, "X.X. ") > 0 Then
                    GelenTema = Mid(Cells(ActiveCell.Row, 25).Value, 6, Len(Cells(ActiveCell.Row, 25).Value)) & " " & Cells(ActiveCell.Row, 26).Value
                Else
                    GelenTema = Cells(ActiveCell.Row, 25).Value & " " & Cells(ActiveCell.Row, 26).Value
                End If
            Else
                If InStr(Cells(ActiveCell.Row, 25).Value, "X.X. ") > 0 Then
                    GelenTema = Mid(Cells(ActiveCell.Row, 25).Value, 6, Len(Cells(ActiveCell.Row, 25).Value))
                Else
                    GelenTema = Cells(ActiveCell.Row, 25).Value
                End If
            End If
        End If
        
        objDoc.Tables(1).Cell(Row:=7, Column:=3).Range.Text = GelenTema
        
        'Gelen paket tipi(Package A/Package B/Package C)
        PaketTipi = Cells(ActiveCell.Row, 29).Value
        'Tablo başlığı
'        objDoc.Tables(4).Cell(Row:=1, Column:=1).Range.Text = PaketTipi & " İçerisinden Çıkan"
        objDoc.Tables(4).Cell(Row:=1, Column:=1).Range.Text = "Items Found Inside the " & PaketTipi
        PaketTipi = LCase(Cells(ActiveCell.Row, 29).Value)

        'Ek olarak Dispatch List
        If Cells(ActiveCell.Row, 35).Value = "Yes" Then
            objDoc.Tables(2).Cell(Row:=1, Column:=1).Range.Text = _
            "The " & PaketTipi & " attached to the document above has been opened, and a discrepancy was detected between the item(s) stated as sent and the item(s) actually found inside the " & PaketTipi & ". " & _
            "The list of differences is provided below, and the remaining item(s) mentioned in the document were confirmed to be present as described. " & _
            "Additionally, a full summary of the identified item(s) is included in the attachment."
            
            If Cells(ActiveCell.Row, 36).Value > 1 Then objDoc.Tables(5).Cell(Row:=5, Column:=1).Range.Text = "Attachment: Dispatch List (" & Cells(ActiveCell.Row, 36).Value & " pages)"
            If Cells(ActiveCell.Row, 36).Value < 2 Then objDoc.Tables(5).Cell(Row:=5, Column:=1).Range.Text = "Attachment: Dispatch List (" & Cells(ActiveCell.Row, 36).Value & " page)"
        ElseIf Cells(ActiveCell.Row, 35).Value = "No" Then
            objDoc.Tables(2).Cell(Row:=1, Column:=1).Range.Text = _
            "The " & PaketTipi & " attached to the above-mentioned document has been opened, and a discrepancy has been detected between the item(s) stated as sent and the item(s) found inside the " & PaketTipi & ". " & _
            "A breakdown of the discrepancy is provided below."
        End If

        'Tabloyu doldur
        Set IlkSiraBul = Range("CE7:CE100000").Find(What:=Cells(ActiveCell.Row, 5).Value, SearchDirection:=xlNext, _
                        SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
        Set SonSiraBul = Range("CF7:CF100000").Find(What:=Cells(ActiveCell.Row, 5).Value, SearchDirection:=xlNext, _
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
            objDoc.Tables(4).Cell(Row:=i, Column:=2).Range.Text = Cells(IlkSira + i - 4, 38).Value 'Öğe türü
            objDoc.Tables(4).Cell(Row:=i, Column:=3).Range.Text = Cells(IlkSira + i - 4, 41).Value 'Öğe değeri
            objDoc.Tables(4).Cell(Row:=i, Column:=4).Range.Text = Cells(IlkSira + i - 4, 44).Value 'Adet
            If Cells(IlkSira + i - 4, 47).Value = "Dispatch List" Then
                objDoc.Tables(4).Cell(Row:=i, Column:=5).Range.Text = "Listed in the Dispatch List" 'Öğe ID No
            Else
                objDoc.Tables(4).Cell(Row:=i, Column:=5).Range.Text = Cells(IlkSira + i - 4, 47).Value 'Öğe ID No
            End If
'            'TipB satır tespiti
            objDoc.Tables(4).Cell(Row:=i, Column:=6).Range.Text = Cells(IlkSira, 23).Value  'Tema No (Temai her satıra yaz.)
            objDoc.Tables(4).Cell(Row:=i, Column:=7).Range.Text = Cells(IlkSira + i - 4, 50).Value 'Açıklama
        Next i
    
        
        'Tabloyu doldur (Fark kısmı)
        Set WsFarkGirisRapor1 = ThisWorkbook.Worksheets(8)
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
'            'TipB satır tespiti
            objDoc.Tables(3).Cell(Row:=i, Column:=6).Range.Text = Cells(IlkSira, 23).Value  'Tema No (Temai her satıra yaz.)
            objDoc.Tables(3).Cell(Row:=i, Column:=7).Range.Text = WsFarkGirisRapor1.Cells(IlkSirax + i - 4, 13).Value 'Açıklama
        Next i
    

        'imzalar
        objDoc.Tables(5).Cell(Row:=2, Column:=2).Range.Text = Cells(ActiveCell.Row, 104).Value 'Ad Soyad1
        objDoc.Tables(5).Cell(Row:=3, Column:=2).Range.Text = Cells(ActiveCell.Row, 105).Value 'Unvan1
        objDoc.Tables(5).Cell(Row:=2, Column:=3).Range.Text = Cells(ActiveCell.Row, 107).Value 'Ad Soyad2
        objDoc.Tables(5).Cell(Row:=3, Column:=3).Range.Text = Cells(ActiveCell.Row, 108).Value 'Unvan2
        
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
        objDoc.Tables(1).Cell(Row:=4, Column:=3).Range.Text = Cells(ActiveCell.Row, 31).Value
        'Geliş tarihi
        objDoc.Tables(1).Cell(Row:=5, Column:=3).Range.Text = Cells(ActiveCell.Row, 28).Value
        'By Hand/By Mail
        If Cells(ActiveCell.Row, 30).Value = "By Hand" Then
            objDoc.CheckElden.Value = True
        ElseIf Cells(ActiveCell.Row, 30).Value = "By Mail" Then
            objDoc.CheckPosta.Value = True
        End If
        'Öğe çıktı/çıkmadı/vb.
        If Cells(ActiveCell.Row, 32).Value = "a. Content as Expected" Then
            objDoc.CheckTam.Value = True
        ElseIf Cells(ActiveCell.Row, 32).Value = "b. Content Empty" Then
            objDoc.CheckYok.Value = True
        ElseIf Cells(ActiveCell.Row, 32).Value = "c. Only Specific Content Available" Then
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
        If Cells(ActiveCell.Row, 25).Value = "Provincial Directorate B" Then
            GelenTema = Cells(ActiveCell.Row, 17).Value & " Provincial Governorship Provincial Directorate B " & Cells(ActiveCell.Row, 26).Value
        ElseIf Cells(ActiveCell.Row, 25).Value = "District Directorate B" Then
            GelenTema = Cells(ActiveCell.Row, 18).Value & " District Governorship District Directorate B " & Cells(ActiveCell.Row, 26).Value
        ElseIf Cells(ActiveCell.Row, 25).Value = "Provincial Directorate C" Then
            GelenTema = Cells(ActiveCell.Row, 17).Value & " Provincial Governorship Provincial Directorate C " & Cells(ActiveCell.Row, 26).Value
        ElseIf Cells(ActiveCell.Row, 25).Value = "District Directorate C" Then
            GelenTema = Cells(ActiveCell.Row, 18).Value & " District Governorship District Directorate C " & Cells(ActiveCell.Row, 26).Value
        ElseIf Cells(ActiveCell.Row, 25).Value = "Provincial Directorate D" Then
            GelenTema = Cells(ActiveCell.Row, 17).Value & " Provincial Governorship Provincial Directorate D " & Cells(ActiveCell.Row, 26).Value
        ElseIf Cells(ActiveCell.Row, 25).Value = "District Directorate D" Then
            GelenTema = Cells(ActiveCell.Row, 18).Value & " District Governorship District Directorate D " & Cells(ActiveCell.Row, 26).Value
        ElseIf Cells(ActiveCell.Row, 25).Value = "Provincial Directorate E" Then
            GelenTema = Cells(ActiveCell.Row, 17).Value & " Provincial Governorship Provincial Directorate E " & Cells(ActiveCell.Row, 26).Value
        ElseIf Cells(ActiveCell.Row, 25).Value = "District Directorate E" Then
            GelenTema = Cells(ActiveCell.Row, 18).Value & " District Governorship District Directorate E " & Cells(ActiveCell.Row, 26).Value
        ElseIf InStr(Cells(ActiveCell.Row, 25).Value, "General Directorate") <> 0 Or InStr(Cells(ActiveCell.Row, 25).Value, "Regional Directorate") <> 0 Then
            If Cells(ActiveCell.Row, 26).Value <> "" Then
                GelenTema = WorksheetFunction.Proper(ThisWorkbook.Worksheets(2).Cells(6, 111).Value) & " " & Cells(ActiveCell.Row, 25).Value & " " & Cells(ActiveCell.Row, 26).Value
            Else
                GelenTema = WorksheetFunction.Proper(ThisWorkbook.Worksheets(2).Cells(6, 111).Value) & " " & Cells(ActiveCell.Row, 25).Value
            End If
        Else
            If Cells(ActiveCell.Row, 26).Value <> "" Then
                If InStr(Cells(ActiveCell.Row, 25).Value, "X.X. ") > 0 Then
                    GelenTema = Mid(Cells(ActiveCell.Row, 25).Value, 6, Len(Cells(ActiveCell.Row, 25).Value)) & " " & Cells(ActiveCell.Row, 26).Value
                Else
                    GelenTema = Cells(ActiveCell.Row, 25).Value & " " & Cells(ActiveCell.Row, 26).Value
                End If
            Else
                If InStr(Cells(ActiveCell.Row, 25).Value, "X.X. ") > 0 Then
                    GelenTema = Mid(Cells(ActiveCell.Row, 25).Value, 6, Len(Cells(ActiveCell.Row, 25).Value))
                Else
                    GelenTema = Cells(ActiveCell.Row, 25).Value
                End If
            End If
        End If
        
        objDoc.Tables(1).Cell(Row:=9, Column:=3).Range.Text = GelenTema
        
        'Gönderenin adresi
'        If Cells(ActiveCell.Row, 18).Value = Cells(ActiveCell.Row, 17).Value & " Organization A" Then
        If Cells(ActiveCell.Row, 18).Value <> "" Then
            objDoc.Tables(1).Cell(Row:=10, Column:=3).Range.Text = Cells(ActiveCell.Row, 18).Value & "/" & Cells(ActiveCell.Row, 17).Value
        Else
            objDoc.Tables(1).Cell(Row:=10, Column:=3).Range.Text = Cells(ActiveCell.Row, 17).Value
        End If
        'Gelen Yazının Tarih ve Sayısı
        objDoc.Tables(1).Cell(Row:=11, Column:=3).Range.Text = Cells(ActiveCell.Row, 20).Value
        objDoc.Tables(1).Cell(Row:=11, Column:=5).Range.Text = Cells(ActiveCell.Row, 21).Value
        'Gönderinin Eki
        If Cells(ActiveCell.Row, 29).Value = "Package A" Then
            objDoc.CheckZarf.Value = True
        ElseIf Cells(ActiveCell.Row, 29).Value = "Package B" Then
            objDoc.CheckTorba.Value = True
        ElseIf Cells(ActiveCell.Row, 29).Value = "Package C" Then
            objDoc.CheckKoli.Value = True
        End If
        'Tabloyu doldur
        Set IlkSiraBul = Range("CE7:CE100000").Find(What:=Cells(ActiveCell.Row, 5).Value, SearchDirection:=xlNext, _
                        SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
        Set SonSiraBul = Range("CF7:CF100000").Find(What:=Cells(ActiveCell.Row, 5).Value, SearchDirection:=xlNext, _
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
            objDoc.Tables(2).Cell(Row:=i, Column:=2).Range.Text = Cells(IlkSira + i - 2, 38).Value 'Öğe türü
            objDoc.Tables(2).Cell(Row:=i, Column:=3).Range.Text = Cells(IlkSira + i - 2, 41).Value 'Öğe değeri
            objDoc.Tables(2).Cell(Row:=i, Column:=4).Range.Text = Cells(IlkSira + i - 2, 44).Value 'Adet
            If Cells(IlkSira + i - 2, 47).Value = "Dispatch List" Then
                objDoc.Tables(2).Cell(Row:=i, Column:=5).Range.Text = "Listed in the Dispatch List" 'Öğe ID No
            Else
                objDoc.Tables(2).Cell(Row:=i, Column:=5).Range.Text = Cells(IlkSira + i - 2, 47).Value 'Öğe ID No
            End If
            'TipB satır tespiti
            If Cells(IlkSira + i - 2, 38).Value = "Item Type - X2" Or (Cells(IlkSira + i - 2, 38).Value = "Item Type - X1" And Cells(IlkSira + i - 2, 41).Value = "1") Then
                 objDoc.Tables(2).Cell(Row:=i, Column:=6).Range.Text = "" 'Tema No (TipB ise Temai yazma.)
            Else
                objDoc.Tables(2).Cell(Row:=i, Column:=6).Range.Text = Cells(IlkSira, 23).Value  'Tema No (Temai her satıra yaz.)
            End If
            objDoc.Tables(2).Cell(Row:=i, Column:=7).Range.Text = Cells(IlkSira + i - 2, 50).Value 'Açıklama
        Next i
    
        'TipA ve tipB tespiti
        TipB = False
        TipA = False
        For i = IlkSira To SonSira
            If Cells(i, 38).Value = "Item Type - X2" Or (Cells(i, 38).Value = "Item Type - X1" And Cells(i, 41).Value = "1") Then
                TipB = True
            Else
                TipA = True
            End If
        Next i
        'Sadece tipBler için tutanak2 tutanağında bulunan öğe ID no ile Tema koloununu sil.
        If TipB = True And TipA = False Then
            For i = 1 To 2
                objDoc.Tables(2).Columns(5).Delete
            Next i
        End If
    
        'Ek olarak Dispatch List
        If Cells(ActiveCell.Row, 35).Value = "Yes" Then
            If Cells(ActiveCell.Row, 36).Value > 1 Then objDoc.Tables(4).Cell(Row:=5, Column:=1).Range.Text = "Attachment: Dispatch List (" & Cells(ActiveCell.Row, 36).Value & " pages)"
            If Cells(ActiveCell.Row, 36).Value < 2 Then objDoc.Tables(4).Cell(Row:=5, Column:=1).Range.Text = "Attachment: Dispatch List (" & Cells(ActiveCell.Row, 36).Value & " page)"
        ElseIf Cells(ActiveCell.Row, 35).Value = "No" Then
            'objDoc.Tables(4).Cell(Row:=5, Column:=1).Range.Text = Cells(ActiveCell.Row, 36).Value
        End If
        
        'imzalar
        objDoc.Tables(4).Cell(Row:=2, Column:=2).Range.Text = Cells(ActiveCell.Row, 104).Value 'Ad Soyad1
        objDoc.Tables(4).Cell(Row:=3, Column:=2).Range.Text = Cells(ActiveCell.Row, 105).Value 'Unvan1
        objDoc.Tables(4).Cell(Row:=2, Column:=3).Range.Text = Cells(ActiveCell.Row, 107).Value 'Ad Soyad2
        objDoc.Tables(4).Cell(Row:=3, Column:=3).Range.Text = Cells(ActiveCell.Row, 108).Value 'Unvan2
        
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
    If Cells(ActiveCell.Row, 35).Value = "Yes" Then
        'Döküm için sayfayı belirt.(Dispatch List word daosyası bu veriyi işleyecek.)
        Set fso = CreateObject("Scripting.FileSystemObject")
        DokumSayfaGonder = Cells(ActiveCell.Row, 36).Value
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
        objDoc.Tables(2).Cell(Row:=4, Column:=3).Range.Text = Cells(ActiveCell.Row, 20).Value
        'Belge sayısı
        objDoc.Tables(2).Cell(Row:=4, Column:=5).Range.Text = Cells(ActiveCell.Row, 21).Value

        'Alt bilgi ekle
        objDoc.Sections(1).Footers(wdHeaderFooterPrimary).Range.Tables(1).Cell(Row:=1, Column:=1).Range.Text = ReNameDL
    
        'Sayfa sayısı kaydet komutuna bağlandı.
    '    objDoc.Close SaveChanges:=True
    '    objWord.Quit
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
    Cells(ActiveCell.Row, 89).Value = TotalSayfaTutanak1
    Cells(ActiveCell.Row, 90).Value = TotalSayfaDokum
    
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
    Cells(ActiveCell.Row, 18).Value = IlceSakla
End If

'Worksheets(3).Activate
Application.ScreenUpdating = True
Application.DisplayAlerts = True
'Application.EnableEvents = True

'End If

End Sub

Sub Rapor1Rapor()

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
Dim Explorer As Integer, b As Long
Dim TipBRaporu As Boolean


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
    Set IlkSiraBul = Range("CE7:CE100000").Find(What:=Cells(j, 5).Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    Set SonSiraBul = Range("CF7:CF100000").Find(What:=Cells(j, 5).Value, SearchDirection:=xlNext, _
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
        MsgBox "Your operation cannot be started due to missing and/or incorrect data in Statement 1.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    
    'Pathfinder...
    AutoPath = ThisWorkbook.Path
    DestOperasyon = AutoPath & "\System Files\Operation\"
    'TUTANAK1 TANIMLARI
    SourceTutanak1Normal = AutoPath & "\System Files\System Templates\Statement 1 Templates\Statement 1.docm"
    SourceDL = AutoPath & "\System Files\System Templates\Statement 1 Templates\Dispatch List.docm"
    'RAPOR TANIMLARI
    SourceRaporNormal = AutoPath & "\System Files\System Templates\Report 1 Template\Report 1.docm"
    'TUTANAK2 TANIMLARI
    SourceTutanak2Normal = AutoPath & "\System Files\System Templates\Statement 2 Template\Statement 2.docm"
    'ÜST YAZI TANIMLARI
    SourceUstYaziNormal = AutoPath & "\System Files\System Templates\Report 1 Cover Letter Template\Report 1 Cover Letter.docm"
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
    If Not Dir(SourceRaporNormal, vbDirectory) <> vbNullString Then
        MsgBox "Cannot access the directory " & SourceRaporNormal & ". The names of folders and/or files in this directory may have been changed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If

    'Operation klasörü içinde kullanıcı modülü klasörü yoksa oluştur.
    If Not Dir(DestOpUserFolder, vbDirectory) <> vbNullString Then
        MkDir DestOpUserFolder
    End If

'____________________________________
If TumDoc = True Then
    'Worksheets(3).Activate
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
    Set IlkSiraBul = Range("CE7:CE100000").Find(What:=Cells(SiraNoIlkSatir, 5).Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    Set SonSiraBul = Range("CF7:CF100000").Find(What:=Cells(SiraNoIlkSatir, 5).Value, SearchDirection:=xlNext, _
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
    Set IlkSiraBul = Range("CE7:CE100000").Find(What:=Cells(SiraNoIlkSatir, 5).Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    Set SonSiraBul = Range("CF7:CF100000").Find(What:=Cells(SiraNoIlkSatir, 5).Value, SearchDirection:=xlNext, _
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

    'TipB rapor tespiti
    TipBRaporu = False
    If Cells(ActiveCell.Row, 7).Value <> "" Then
        If Cells(ActiveCell.Row, 38).Value = "Item Type - X2" Or (Cells(ActiveCell.Row, 38).Value = "Item Type - X1" And Cells(ActiveCell.Row, 41).Value = "1") Then
             TipBRaporu = True
        End If
    End If
           
'________________________________________

    If TumDoc = False Then
        'Close the all Word application
        Call ModuleReport1.OpenWordControl
    
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
    
    '    Call ModuleReport1.OpenWordControl
    
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
    If TipBRaporu = True Then
        objDoc.Tables(1).Cell(Row:=2, Column:=1).Range.Text = "Report 1 (Type B)"
    End If
    'Rapor tarihi
    objDoc.Tables(1).Cell(Row:=4, Column:=3).Range.Text = Cells(IlkSira, 60).Value
    'Rapor No
    objDoc.Tables(1).Cell(Row:=5, Column:=3).Range.Text = Cells(AltRaporNoIlk, 59).Value
    'Gönderen tema
    If Cells(IlkSira, 25).Value = "Provincial Directorate B" Then
        If Cells(IlkSira, 26).Value <> "" Then
            GelenTema = Cells(IlkSira, 17).Value & " Provincial Governorship Provincial Directorate B " & Cells(IlkSira, 26).Value
        Else
            GelenTema = Cells(IlkSira, 17).Value & " Provincial Governorship Provincial Directorate B"
        End If
    ElseIf Cells(IlkSira, 25).Value = "District Directorate B" Then
        If Cells(IlkSira, 26).Value <> "" Then
            GelenTema = Cells(IlkSira, 18).Value & " District Governorship District Directorate B " & Cells(IlkSira, 26).Value
        Else
            GelenTema = Cells(IlkSira, 18).Value & " District Governorship District Directorate B"
        End If
    ElseIf Cells(IlkSira, 25).Value = "Provincial Directorate C" Then
        If Cells(IlkSira, 26).Value <> "" Then
            GelenTema = Cells(IlkSira, 17).Value & " Provincial Governorship Provincial Directorate C " & Cells(IlkSira, 26).Value
        Else
            GelenTema = Cells(IlkSira, 17).Value & " Provincial Governorship Provincial Directorate C"
        End If
    ElseIf Cells(IlkSira, 25).Value = "District Directorate C" Then
        If Cells(IlkSira, 26).Value <> "" Then
            GelenTema = Cells(IlkSira, 18).Value & " District Governorship District Directorate C " & Cells(IlkSira, 26).Value
        Else
            GelenTema = Cells(IlkSira, 18).Value & " District Governorship District Directorate C"
        End If
    ElseIf Cells(IlkSira, 25).Value = "Provincial Directorate D" Then
        If Cells(IlkSira, 26).Value <> "" Then
            GelenTema = Cells(IlkSira, 17).Value & " Provincial Governorship Provincial Directorate D " & Cells(IlkSira, 26).Value
        Else
            GelenTema = Cells(IlkSira, 17).Value & " Provincial Governorship Provincial Directorate D"
        End If
    ElseIf Cells(IlkSira, 25).Value = "District Directorate D" Then
        If Cells(IlkSira, 26).Value <> "" Then
            GelenTema = Cells(IlkSira, 18).Value & " District Governorship District Directorate D " & Cells(IlkSira, 26).Value
        Else
            GelenTema = Cells(IlkSira, 18).Value & " District Governorship District Directorate D"
        End If
    ElseIf Cells(IlkSira, 25).Value = "Provincial Directorate E" Then
        If Cells(IlkSira, 26).Value <> "" Then
            GelenTema = Cells(IlkSira, 17).Value & " Provincial Governorship Provincial Directorate E " & Cells(IlkSira, 26).Value
        Else
            GelenTema = Cells(IlkSira, 17).Value & " Provincial Governorship Provincial Directorate E"
        End If
    ElseIf Cells(IlkSira, 25).Value = "District Directorate E" Then
        If Cells(IlkSira, 26).Value <> "" Then
            GelenTema = Cells(IlkSira, 18).Value & " District Governorship District Directorate E " & Cells(IlkSira, 26).Value
        Else
            GelenTema = Cells(IlkSira, 18).Value & " District Governorship District Directorate E"
        End If
    ElseIf InStr(Cells(IlkSira, 25).Value, "General Directorate") <> 0 Or InStr(Cells(IlkSira, 25).Value, "Regional Directorate") <> 0 Then
        If Cells(IlkSira, 26).Value <> "" Then
            GelenTema = WorksheetFunction.Proper(ThisWorkbook.Worksheets(2).Cells(6, 111).Value) & " " & Cells(IlkSira, 25).Value & " " & Cells(IlkSira, 26).Value
        Else
            GelenTema = WorksheetFunction.Proper(ThisWorkbook.Worksheets(2).Cells(6, 111).Value) & " " & Cells(IlkSira, 25).Value
        End If
    Else
        If Cells(IlkSira, 26).Value <> "" Then
            If InStr(Cells(IlkSira, 25).Value, "X.X. ") > 0 Then
                GelenTema = Mid(Cells(IlkSira, 25).Value, 6, Len(Cells(IlkSira, 25).Value)) & " " & Cells(IlkSira, 26).Value
            Else
                GelenTema = Cells(IlkSira, 25).Value & " " & Cells(IlkSira, 26).Value
            End If
        Else
            If InStr(Cells(IlkSira, 25).Value, "X.X. ") > 0 Then
                GelenTema = Mid(Cells(IlkSira, 25).Value, 6, Len(Cells(IlkSira, 25).Value))
            Else
                GelenTema = Cells(IlkSira, 25).Value
            End If
        End If
    End If
    
    'İlgi
'    RaporIlgi = GelenTema & "n" & Right(GelenTema, 1) & "n " & Cells(IlkSira, 20).Value & " tarihli ve " & Cells(IlkSira, 21).Value & " sayılı yazısı."
    RaporIlgi = "The letter from the " & GelenTema & ", dated " & Cells(IlkSira, 20).Value & ", reference number " & Cells(IlkSira, 21).Value & "."

    objDoc.Tables(1).Cell(Row:=6, Column:=3).Range.Text = RaporIlgi
    'Tekil-çoğul tipA
    AdetTopla = Application.Sum(Range(Cells(AltRaporNoIlk, 44), Cells(AltRaporNoSon, 44)))
    If TipBRaporu = True Then
        If AdetTopla = 1 Then
            TekCogulTipA = "Type B"
        ElseIf AdetTopla > 1 Then
            TekCogulTipA = "Type B items"
        End If
    Else
        If AdetTopla = 1 Then
            TekCogulTipA = "Type A"
        ElseIf AdetTopla > 1 Then
            TekCogulTipA = "Type A items"
        End If
    End If
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
            If TipBRaporu = True Then
                objDoc.Tables(2).Cell(Row:=j + 3, Column:=1).Range.Text = ""
            Else
                objDoc.Tables(2).Cell(Row:=j + 3, Column:=1).Range.Text = "Item ID No."
            End If
            objDoc.Tables(2).Cell(Row:=j + 0, Column:=2).Range.Text = ":"
            objDoc.Tables(2).Cell(Row:=j + 1, Column:=2).Range.Text = ":"
            objDoc.Tables(2).Cell(Row:=j + 2, Column:=2).Range.Text = ":"
            If TipBRaporu = True Then
                objDoc.Tables(2).Cell(Row:=j + 3, Column:=2).Range.Text = ""
            Else
                objDoc.Tables(2).Cell(Row:=j + 3, Column:=2).Range.Text = ":"
            End If
            'Verileri aktar
            objDoc.Tables(2).Cell(Row:=j + 0, Column:=3).Range.Text = Cells(i, 38).Value 'Öğe türü
            objDoc.Tables(2).Cell(Row:=j + 1, Column:=3).Range.Text = Cells(i, 41).Value 'Öğe değeri
            objDoc.Tables(2).Cell(Row:=j + 2, Column:=3).Range.Text = Cells(i, 44).Value 'Adet
            If TipBRaporu = True Then
                objDoc.Tables(2).Cell(Row:=j + 3, Column:=3).Range.Text = "" 'Öğe ID
            Else
                If Cells(i, 47).Value = "Dispatch List" Then
                    objDoc.Tables(2).Cell(Row:=j + 3, Column:=3).Range.Text = "Listed in the Attachment to Statement 1" 'Öğe ID
                Else
                    objDoc.Tables(2).Cell(Row:=j + 3, Column:=3).Range.Text = Cells(i, 47).Value 'Öğe ID
                End If
            End If
            j = j + 0
        ElseIf x Mod 2 = 0 Then
            'Verileri aktar
            objDoc.Tables(2).Cell(Row:=j + 0, Column:=4).Range.Text = Cells(i, 38).Value 'Öğe türü
            objDoc.Tables(2).Cell(Row:=j + 1, Column:=4).Range.Text = Cells(i, 41).Value 'Öğe değeri
            objDoc.Tables(2).Cell(Row:=j + 2, Column:=4).Range.Text = Cells(i, 44).Value 'Adet
            If TipBRaporu = True Then
                objDoc.Tables(2).Cell(Row:=j + 3, Column:=4).Range.Text = "" 'Öğe ID
            Else
                If Cells(i, 47).Value = "Dispatch List" Then
                    objDoc.Tables(2).Cell(Row:=j + 3, Column:=4).Range.Text = "Listed in the Attachment to Statement 1" 'Öğe ID
                Else
                    objDoc.Tables(2).Cell(Row:=j + 3, Column:=4).Range.Text = Cells(i, 47).Value 'Öğe ID
                End If
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
    If TipBRaporu = True Then
        objDoc.Tables(3).Cell(Row:=1, Column:=1).Range.Text = ""  'Tema no başık
        objDoc.Tables(3).Cell(Row:=1, Column:=2).Range.Text = ""  'Tema no iki nokta
        objDoc.Tables(3).Cell(Row:=1, Column:=3).Range.Text = ""  'Tema no
    Else
        objDoc.Tables(3).Cell(Row:=1, Column:=3).Range.Text = Cells(IlkSira, 23).Value  'Tema no
    End If
    'Rapor1 metin kısmı
    'Tanımlamalar
    AdetTopla = Application.Sum(Range(Cells(AltRaporNoIlk, 44), Cells(AltRaporNoSon, 44)))
    If TipBRaporu = True Then
        If AdetTopla = 1 Then
'            Ek1 = "tip_Bnin"
            Ek1 = "Tip B item is"
        ElseIf AdetTopla > 1 Then
'            Ek1 = "tip_Blerin"
            Ek1 = "Tip B items are"
        End If
    Else
        If AdetTopla = 1 Then
            Ek1 = "Tip A item is"
        ElseIf AdetTopla > 1 Then
            Ek1 = "Tip A items are"
        End If
    End If
    'İlk kısım
'    If TipBRaporu = True Then
'        Bolum1 = "Yukarıda türü, öğe değeri ve miktarı belirtilen "
'    Else
'        Bolum1 = "Yukarıda türü, öğe değeri, miktarı, öğe ID numarası belirtilen "
'    End If
'    Bolum2 = " ilk inceleme sonucunda "
'    Bolum3 = " olduğu değerlendirilmektedir."
'    Ek2 = LCase(Cells(AltRaporNoIlk, 54).Value)
'    Birlestir = Bolum1 & Ek1 & Bolum2 & Ek2 & Bolum3
'    objDoc.Tables(3).Cell(Row:=2, Column:=3).Range.Text = Birlestir
'    Set MyRange = objDoc.Tables(3).Cell(Row:=2, Column:=3).Range
'    MyRange.Find.Execute FindText:=Ek2
'    MyRange.Font.Bold = True
    
    'İlk kısım
    If TipBRaporu = True Then
        Bolum1 = "The type, item value, and quantity mentioned above indicate that the "
    Else
        Bolum1 = "The type, item value, quantity, and item ID number mentioned above indicate that the "
    End If
    Bolum2 = " preliminarily assessed to be "
    Bolum3 = "."
    Ek2 = LCase(Cells(AltRaporNoIlk, 54).Value)
    Birlestir = Bolum1 & Ek1 & Bolum2 & Ek2 & Bolum3
    objDoc.Tables(3).Cell(Row:=2, Column:=3).Range.Text = Birlestir
    Set MyRange = objDoc.Tables(3).Cell(Row:=2, Column:=3).Range
    MyRange.Find.Execute FindText:=Ek2
    MyRange.Font.Bold = True

    'İkinci kısım
'    Bolum1 = "Ancak "
'    Ek1 = LCase(Cells(AltRaporNoIlk, 54).Value)
'    Ek2 = "validity/invalidity " 'Ek1 & "lik "
'    Bolum2 = "konusundaki kesin sonuç,"
'    If TipBRaporu = True Then
'        Bolum3 = " """ & "invalid TipB XXX Uygulama Kılavuzu" & """"
'        Bolum4 = " in x. maddesi uyarınca yetkili birimler olan Process Monitoring Directorate, Arbitration veya Karar Kurulları tarafından XXX General Directoratenden istenecek olan xxx raporu ile belirlenecektir."
'    Else
'        Bolum3 = " """ & "invalid TipAların XXX Uygulama Kılavuzu" & """"
'        Bolum4 = " in x. maddesi uyarınca yetkili birimler olan Process Monitoring Directorate, Arbitration veya Karar Kurulları tarafından xxx kurumundan istenecek olan rapor2 ile belirlenecektir."
'    End If
'    Birlestir = Bolum1 & Ek2 & Bolum2 & Bolum3 & Bolum4
'    objDoc.Tables(3).Cell(Row:=3, Column:=3).Range.Text = Birlestir
'
'    Set MyRange = objDoc.Tables(3).Cell(Row:=3, Column:=3).Range
'    MyRange.Find.Execute FindText:=Bolum3
'    MyRange.Font.Bold = True
'    objDoc.Tables(3).Cell(Row:=4, Column:=3).Range.Text = "Bilgilerinize sunarız."

    'İkinci kısım
    Bolum1 = "However, "
    Ek1 = LCase(Cells(AltRaporNoIlk, 54).Value)
    Ek2 = "the final conclusion regarding the validity or invalidity,"
    If TipBRaporu = True Then
        Bolum3 = """the Invalid Type B XXX Implementation Guide"""
    Else
        Bolum3 = """the Invalid Type A XXX Implementation Guide"""
    End If
    Bolum4 = " in accordance with Article x of " & Bolum3 & _
             ", will be determined based on the report to be requested from the relevant institution by the authorized bodies, namely the Process Monitoring Directorate, the Arbitration Unit, or the Decision Board."
    Birlestir = Bolum1 & Ek2 & Bolum4
    objDoc.Tables(3).Cell(Row:=3, Column:=3).Range.Text = Birlestir
    Set MyRange = objDoc.Tables(3).Cell(Row:=3, Column:=3).Range
    MyRange.Find.Execute FindText:=Bolum3
    MyRange.Font.Bold = True

        
    'Artık satırı sil
    objDoc.Tables(2).Rows(j + 4).Delete

    'imzalar
    objDoc.Tables(4).Cell(Row:=4, Column:=2).Range.Text = Cells(ActiveCell.Row, 110).Value 'Ad Soyad1
    objDoc.Tables(4).Cell(Row:=5, Column:=2).Range.Text = Cells(ActiveCell.Row, 111).Value 'Unvan1
    objDoc.Tables(4).Cell(Row:=4, Column:=3).Range.Text = Cells(ActiveCell.Row, 113).Value 'Ad Soyad2
    objDoc.Tables(4).Cell(Row:=5, Column:=3).Range.Text = Cells(ActiveCell.Row, 114).Value 'Unvan2
        
    'Alt bilgi ekle
    objDoc.Sections(1).Footers(wdHeaderFooterPrimary).Range.Tables(1).Cell(Row:=1, Column:=1).Range.Text = ReNameRaporNormal
    
'    'İmza boşluğunu sayfaya sığdırmak için düzenle
    If AltRaporNoSon - AltRaporNoIlk + 1 > 8 And AltRaporNoSon - AltRaporNoIlk + 1 < 11 Then
        objDoc.Tables(2).Rows(6).Height = 4
        objDoc.Tables(2).Rows(11).Height = 4
        objDoc.Tables(2).Rows(16).Height = 4
        objDoc.Tables(2).Rows(21).Height = 4
        objDoc.Tables(2).Rows(26).Height = 4
    ElseIf AltRaporNoSon - AltRaporNoIlk + 1 > 10 And AltRaporNoSon - AltRaporNoIlk + 1 < 13 Then
        objDoc.Tables(2).Rows(6).Height = 4
        objDoc.Tables(2).Rows(11).Height = 4
        objDoc.Tables(2).Rows(16).Height = 4
        objDoc.Tables(2).Rows(21).Height = 4
        objDoc.Tables(2).Rows(26).Height = 4
        objDoc.Tables(2).Rows(31).Height = 4
        objDoc.Tables(4).Rows(4).Delete
        objDoc.Tables(4).Rows(3).Delete
        objDoc.Tables(4).Rows(2).Delete
        objDoc.Tables(4).Rows(1).Height = 3
    ElseIf AltRaporNoSon - AltRaporNoIlk + 1 > 12 And AltRaporNoSon - AltRaporNoIlk + 1 < 15 Then
        objDoc.Tables(2).Rows(36).Height = 60
    ElseIf AltRaporNoSon - AltRaporNoIlk + 1 > 14 And AltRaporNoSon - AltRaporNoIlk + 1 < 17 Then
        objDoc.Tables(2).Rows(6).Height = 10
        objDoc.Tables(2).Rows(11).Height = 10
        objDoc.Tables(2).Rows(16).Height = 10
        objDoc.Tables(2).Rows(21).Height = 10
        objDoc.Tables(2).Rows(26).Height = 10
        objDoc.Tables(2).Rows(31).Height = 10
        objDoc.Tables(2).Rows(36).Height = 10
        objDoc.Tables(2).Rows(41).Height = 10
    ElseIf AltRaporNoSon - AltRaporNoIlk + 1 > 16 And AltRaporNoSon - AltRaporNoIlk + 1 < 19 Then
        objDoc.Tables(2).Rows(6).Height = 12
        objDoc.Tables(2).Rows(11).Height = 12
        objDoc.Tables(2).Rows(16).Height = 12
        objDoc.Tables(2).Rows(21).Height = 12
        objDoc.Tables(2).Rows(26).Height = 12
        objDoc.Tables(2).Rows(31).Height = 12
        objDoc.Tables(2).Rows(36).Height = 12
        objDoc.Tables(2).Rows(41).Height = 12
        objDoc.Tables(2).Rows(46).Height = 12
    ElseIf AltRaporNoSon - AltRaporNoIlk + 1 > 18 And AltRaporNoSon - AltRaporNoIlk + 1 < 21 Then
        objDoc.Tables(2).Rows(6).Height = 12
        objDoc.Tables(2).Rows(11).Height = 12
        objDoc.Tables(2).Rows(16).Height = 12
        objDoc.Tables(2).Rows(21).Height = 12
        objDoc.Tables(2).Rows(26).Height = 12
        objDoc.Tables(2).Rows(31).Height = 12
        objDoc.Tables(2).Rows(36).Height = 12
        objDoc.Tables(2).Rows(41).Height = 12
        objDoc.Tables(2).Rows(46).Height = 12
    End If
    
    'Rapor1 sayfa sayısını oluşturmadan önce eskisini sil
    For i = IlkSira To SonSira
        If Cells(i, 11).Value = "" Then
            Cells(i, 91).Value = ""
        End If
    Next i
    
    'Sayfa sayısı kaydet komutuna bağlandı.
'    objDoc.Close SaveChanges:=True
'    objWord.Quit
    objDoc.Save
    'Sayfayı text dosyasından çek
    TxtFileRapor = DestOpUserFolder & "Report 1 Page Count.txt"
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
    Cells(ActiveCell.Row, 91).Value = TotalSayfaRapor

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

'    If TumDoc = True Then
'        'Worksheets(3).Activate
'    End If

Next Explorer

If TumDoc = True Then
    'Tümünün bulunduğu butonu tekrar seç.
    Cells(b, 10).Select
End If

Son:


'Worksheets(3).Activate
Application.ScreenUpdating = True
Application.DisplayAlerts = True
'Application.EnableEvents = True

'End If

End Sub

Sub Rapor1Tutanak2()

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
Dim TipB, TipA As Boolean


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
    Set IlkSiraBul = Range("CE7:CE100000").Find(What:=Cells(j, 5).Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    Set SonSiraBul = Range("CF7:CF100000").Find(What:=Cells(j, 5).Value, SearchDirection:=xlNext, _
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
        MsgBox "Your operation cannot be started due to missing and/or incorrect data in Statement 1.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    For i = IlkSira To SonSira
        If Cells(i, 7).Value = "x" Then
            MsgBox "Your operation cannot be started due to missing and/or incorrect data in Report 1.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
            GoTo Son
        End If
    Next i
        
    'Pathfinder...
    AutoPath = ThisWorkbook.Path
    DestOperasyon = AutoPath & "\System Files\Operation\"
    'TUTANAK1 TANIMLARI
    SourceTutanak1Normal = AutoPath & "\System Files\System Templates\Statement 1 Templates\Statement 1.docm"
    SourceDL = AutoPath & "\System Files\System Templates\Statement 1 Templates\Dispatch List.docm"
    'RAPOR TANIMLARI
    SourceRaporNormal = AutoPath & "\System Files\System Templates\Report 1 Template\Report 1.docm"
    'TUTANAK2 TANIMLARI
    SourceTutanak2Normal = AutoPath & "\System Files\System Templates\Statement 2 Template\Statement 2.docm"
    'ÜST YAZI TANIMLARI
    SourceUstYaziNormal = AutoPath & "\System Files\System Templates\Report 1 Cover Letter Template\Report 1 Cover Letter.docm"
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
        Call ModuleReport1.OpenWordControl
        
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
    
    '    Call ModuleReport1.OpenWordControl
         
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
    objDoc.Tables(1).Cell(Row:=4, Column:=3).Range.Text = Cells(ActiveCell.Row, 63).Value
    'Belge tarihi ve numarası
    objDoc.Tables(1).Cell(Row:=6, Column:=3).Range.Text = Cells(ActiveCell.Row, 20).Value
    objDoc.Tables(1).Cell(Row:=6, Column:=5).Range.Text = Cells(ActiveCell.Row, 21).Value
    
    'Gönderilen
    If Cells(ActiveCell.Row, 64).Value = "Provincial Directorate B" Then
        GidenTema = Cells(ActiveCell.Row, 69).Value & " Provincial Governorship Provincial Directorate B " & Cells(ActiveCell.Row, 65).Value
    ElseIf Cells(ActiveCell.Row, 64).Value = "District Directorate B" Then
        GidenTema = Cells(ActiveCell.Row, 70).Value & " District Governorship District Directorate B " & Cells(ActiveCell.Row, 65).Value
    ElseIf Cells(ActiveCell.Row, 64).Value = "Provincial Directorate C" Then
        GidenTema = Cells(ActiveCell.Row, 69).Value & " Provincial Governorship Provincial Directorate C " & Cells(ActiveCell.Row, 65).Value
    ElseIf Cells(ActiveCell.Row, 64).Value = "District Directorate C" Then
        GidenTema = Cells(ActiveCell.Row, 70).Value & " District Governorship District Directorate C " & Cells(ActiveCell.Row, 65).Value
    ElseIf Cells(ActiveCell.Row, 64).Value = "Provincial Directorate D" Then
        GidenTema = Cells(ActiveCell.Row, 69).Value & " Provincial Governorship Provincial Directorate D " & Cells(ActiveCell.Row, 65).Value
    ElseIf Cells(ActiveCell.Row, 64).Value = "District Directorate D" Then
        GidenTema = Cells(ActiveCell.Row, 70).Value & " District Governorship District Directorate D " & Cells(ActiveCell.Row, 65).Value
    ElseIf Cells(ActiveCell.Row, 64).Value = "Provincial Directorate E" Then
        GidenTema = Cells(ActiveCell.Row, 69).Value & " Provincial Governorship Provincial Directorate E " & Cells(ActiveCell.Row, 65).Value
    ElseIf Cells(ActiveCell.Row, 64).Value = "District Directorate E" Then
        GidenTema = Cells(ActiveCell.Row, 70).Value & " District Governorship District Directorate E " & Cells(ActiveCell.Row, 65).Value
    ElseIf InStr(Cells(ActiveCell.Row, 64).Value, "General Directorate") <> 0 Or InStr(Cells(ActiveCell.Row, 64).Value, "Regional Directorate") <> 0 Then
        If Cells(ActiveCell.Row, 65).Value <> "" Then
            GidenTema = WorksheetFunction.Proper(ThisWorkbook.Worksheets(2).Cells(6, 111).Value) & " " & Cells(ActiveCell.Row, 64).Value & " " & Cells(ActiveCell.Row, 65).Value
        Else
            GidenTema = WorksheetFunction.Proper(ThisWorkbook.Worksheets(2).Cells(6, 111).Value) & " " & Cells(ActiveCell.Row, 64).Value
        End If
    Else
        If Cells(ActiveCell.Row, 65).Value <> "" Then
            If InStr(Cells(ActiveCell.Row, 64).Value, "X.X. ") > 0 Then
                GidenTema = Mid(Cells(ActiveCell.Row, 64).Value, 6, Len(Cells(ActiveCell.Row, 64).Value)) & " " & Cells(ActiveCell.Row, 65).Value
            Else
                GidenTema = Cells(ActiveCell.Row, 64).Value & " " & Cells(ActiveCell.Row, 65).Value
            End If
        Else
            If InStr(Cells(ActiveCell.Row, 64).Value, "X.X. ") > 0 Then
                GidenTema = Mid(Cells(ActiveCell.Row, 64).Value, 6, Len(Cells(ActiveCell.Row, 64).Value))
            Else
                GidenTema = Cells(ActiveCell.Row, 64).Value
            End If
        End If
    End If
    
    objDoc.Tables(1).Cell(Row:=5, Column:=3).Range.Text = GidenTema
    
    'Tabloyu doldur
    Set IlkSiraBul = Range("CE7:CE100000").Find(What:=Cells(ActiveCell.Row, 5).Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    Set SonSiraBul = Range("CF7:CF100000").Find(What:=Cells(ActiveCell.Row, 5).Value, SearchDirection:=xlNext, _
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
        objDoc.Tables(2).Cell(Row:=i, Column:=2).Range.Text = Cells(IlkSira + i - 2, 38).Value 'Öğe türü
        objDoc.Tables(2).Cell(Row:=i, Column:=3).Range.Text = Cells(IlkSira + i - 2, 41).Value 'Öğe değeri
        objDoc.Tables(2).Cell(Row:=i, Column:=4).Range.Text = Cells(IlkSira + i - 2, 44).Value 'Adet
        If Cells(IlkSira + i - 2, 47).Value = "Dispatch List" Then
            objDoc.Tables(2).Cell(Row:=i, Column:=5).Range.Text = "Listed in the Attachment to Statement 1" '"Listed in the Dispatch List" 'Öğe ID No
        Else
            objDoc.Tables(2).Cell(Row:=i, Column:=5).Range.Text = Cells(IlkSira + i - 2, 47).Value 'Öğe ID No
        End If
        'TipB satır tespiti
        If Cells(IlkSira + i - 2, 38).Value = "Item Type - X2" Or (Cells(IlkSira + i - 2, 38).Value = "Item Type - X1" And Cells(IlkSira + i - 2, 41).Value = "1") Then
             objDoc.Tables(2).Cell(Row:=i, Column:=6).Range.Text = "" 'Tema No (TipB ise Temai yazma.)
        Else
            objDoc.Tables(2).Cell(Row:=i, Column:=6).Range.Text = Cells(IlkSira, 23).Value  'Tema No (Temai her satıra yaz.)
        End If
        objDoc.Tables(2).Cell(Row:=i, Column:=7).Range.Text = Cells(IlkSira + i - 2, 50).Value 'Açıklama
    Next i

    'TipA ve tipB tespiti
    TipB = False
    TipA = False
    For i = IlkSira To SonSira
        If Cells(i, 38).Value = "Item Type - X2" Or (Cells(i, 38).Value = "Item Type - X1" And Cells(i, 41).Value = "1") Then
            TipB = True
        Else
            TipA = True
        End If
    Next i
    'Sadece tipBler için tutanak2 tutanağında bulunan öğe ID no ile Tema koloununu sil.
    If TipB = True And TipA = False Then
        For i = 1 To 2
            objDoc.Tables(2).Columns(5).Delete
        Next i
    End If

    'Tutanak2 tutanağının metin kısmı
    'Tanımlamalar
    If Application.Sum(Range(Cells(IlkSira, 44), Cells(SonSira, 44))) > 1 Then
        Ek4 = "items"
    Else
        Ek4 = "item"
    End If
    
    Bolum1 = vbTab & "A total of "
    Bolum2 = " " & Ek4 & " , as described above, have been placed inside "
    Bolum3 = " "
    Bolum4 = " and enclosed to be sent to the designated unit mentioned above."
    
    Ek1 = Application.Sum(Range(Cells(IlkSira, 44), Cells(SonSira, 44)))
    Ek2 = Cells(ActiveCell.Row, 68).Value
    Ek3 = Application.WorksheetFunction.Proper(Cells(ActiveCell.Row, 67).Value)
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
    objDoc.Tables(4).Cell(Row:=6, Column:=2).Range.Text = Cells(ActiveCell.Row, 116).Value 'Ad Soyad1
    objDoc.Tables(4).Cell(Row:=7, Column:=2).Range.Text = Cells(ActiveCell.Row, 117).Value 'Unvan1
    objDoc.Tables(4).Cell(Row:=6, Column:=3).Range.Text = Cells(ActiveCell.Row, 119).Value 'Ad Soyad2
    objDoc.Tables(4).Cell(Row:=7, Column:=3).Range.Text = Cells(ActiveCell.Row, 120).Value 'Unvan2
    
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
    Cells(ActiveCell.Row, 92).Value = TotalSayfaTutanak2

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

'Worksheets(3).Activate
Application.ScreenUpdating = True
Application.DisplayAlerts = True
'Application.EnableEvents = True

'End If

End Sub

Sub Rapor1UstYazi()

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
Dim IlBuyukHarf, IlKucukHarf, IlceBuyukHarf, IlceKucukHarf As String
Dim TipB, TipA As Boolean
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
If InStr(Cells(ActiveCell.Row, 70).Value, " Organization A") <> 0 Then
    IlceSakla = Cells(ActiveCell.Row, 70).Value
    Cells(ActiveCell.Row, 70).Value = ""
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
    Set IlkSiraBul = Range("CE7:CE100000").Find(What:=Cells(j, 5).Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    Set SonSiraBul = Range("CF7:CF100000").Find(What:=Cells(j, 5).Value, SearchDirection:=xlNext, _
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
        MsgBox "Your operation cannot be started due to missing and/or incorrect data in Statement 1.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    For i = IlkSira To SonSira
        If Cells(i, 7).Value = "x" Then
            MsgBox "Your operation cannot be started due to missing and/or incorrect data in Report 1.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
            GoTo Son
        End If
    Next i
    If Cells(IlkSira, 8).Value = "x" Then
        MsgBox "Your operation cannot be started due to missing and/or incorrect data in Statement 2.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    
    'Pathfinder...
    AutoPath = ThisWorkbook.Path
    DestOperasyon = AutoPath & "\System Files\Operation\"
    'TUTANAK1 TANIMLARI
    SourceTutanak1Normal = AutoPath & "\System Files\System Templates\Statement 1 Templates\Statement 1.docm"
    SourceDL = AutoPath & "\System Files\System Templates\Statement 1 Templates\Dispatch List.docm"
    'RAPOR TANIMLARI
    SourceRaporNormal = AutoPath & "\System Files\System Templates\Report 1 Template\Report 1.docm"
    'TUTANAK2 TANIMLARI
    SourceTutanak2Normal = AutoPath & "\System Files\System Templates\Statement 2 Template\Statement 2.docm"
    'ÜST YAZI TANIMLARI
    SourceUstYaziNormal = AutoPath & "\System Files\System Templates\Report 1 Cover Letter Template\Report 1 Cover Letter.docm"
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
    Set IlkSiraBul = Range("CE7:CE100000").Find(What:=Cells(ActiveCell.Row, 5).Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    Set SonSiraBul = Range("CF7:CF100000").Find(What:=Cells(ActiveCell.Row, 5).Value, SearchDirection:=xlNext, _
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
    If Cells(IlkSira, 89).Value = "" Then
        MsgBox "Cover letter for serial number " & Cells(IlkSira, 5).Value & " cannot be prepared because Statement 1 has not been created. Please create Statement 1 for serial number " & Cells(IlkSira, 5).Value & " and try again.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    
    'Report(s) check
    For i = IlkSira To SonSira
        If Cells(i, 7).Value <> "" Then
            If Cells(i, 91).Value = "" Then
                MsgBox "Report 1 with number " & Cells(i, 11).Value & " has not been created, so the cover letter for serial number " & Cells(IlkSira, 5).Value & " cannot be prepared. Please create Report 1 with number " & Cells(i, 11).Value & " and try again.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
                GoTo Son
            End If
        End If
    Next i
    
    'Statement 2 check
    If Cells(IlkSira, 92).Value = "" Then
        MsgBox "Cover letter for serial number " & Cells(IlkSira, 5).Value & " cannot be prepared because Statement 2 has not been created. Please create Statement 2 for serial number " & Cells(IlkSira, 5).Value & " and try again.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
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
    ReNameUstYaziNormal = Cells(ActiveCell.Row, 5).Value & "-" & Cells(6, 9).Value
'________________________________________

    If TumDoc = False Then

        'Close the all Word application
        Call ModuleReport1.OpenWordControl
        
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
    
    '    Call ModuleReport1.OpenWordControl
         
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
    'On Error GoTo 0
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
    objDoc.Tables(1).Cell(Row:=1, Column:=2).Range.Text = Birimx & ", " & FormatEnglishDate(Cells(ActiveCell.Row, 75).Value) '(Format(Cells(ActiveCell.Row, 75).Value, "d mmmm yyyy"))

    'Yazı no
    objDoc.Tables(1).Cell(Row:=4, Column:=2).Range.Text = Cells(ActiveCell.Row, 76).Value
    'Sigorta türü
    Ek2 = Cells(ActiveCell.Row, 67).Value
    Ek2 = Right(Ek2, Len(Ek2) - InStr(Ek2, "/"))
    objDoc.Tables(1).Cell(Row:=7, Column:=2).Range.Text = Ek2


    'Muhatap
    IlBuyukHarf = UCase(Replace(Replace(Cells(ActiveCell.Row, 69).Value, "i", "I"), "ı", "I"))
    IlKucukHarf = Cells(ActiveCell.Row, 69).Value
    If Cells(ActiveCell.Row, 70).Value <> "" Then
        IlceBuyukHarf = UCase(Replace(Replace(Cells(ActiveCell.Row, 70).Value, "i", "I"), "ı", "I"))
        IlceKucukHarf = Cells(ActiveCell.Row, 70).Value
    Else
        IlceBuyukHarf = ""
        IlceKucukHarf = ""
    End If
    If Cells(ActiveCell.Row, 70).Value <> "" Then
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
    TipB = False
    BStr = 10
    '1. sayfadan sonraki üst bilgi tanımları
    ustbilgitarih = (Format(Cells(ActiveCell.Row, 75).Value, "dd.mm.yyyy"))
    ustbilgisayi = Cells(ActiveCell.Row, 76).Value
    
    '4'lük
    If Cells(ActiveCell.Row, 65).Value <> "" Then
        If Cells(ActiveCell.Row, 64).Value = "Provincial Directorate B" Or Cells(ActiveCell.Row, 64).Value = "Provincial Directorate C" Or _
        Cells(ActiveCell.Row, 64).Value = "Provincial Directorate D" Or Cells(ActiveCell.Row, 64).Value = "Provincial Directorate E" Then 'VALİLİK
            objDoc.Tables(1).Cell(Row:=BStr, Column:=1).Range.Text = IlBuyukHarf & " PROVINCIAL GOVERNORSHIP"
            objDoc.Tables(1).Cell(Row:=BStr + 1, Column:=1).Range.Text = Cells(ActiveCell.Row, 64).Value  'UCase(Replace(Replace(Cells(ActiveCell.Row, 64).Value, "i", "I"), "ı", "I"))
            objDoc.Tables(1).Cell(Row:=BStr + 2, Column:=1).Range.Text = "(" & Cells(ActiveCell.Row, 65).Value & ")"
            objDoc.Tables(1).Cell(Row:=BStr + 3, Column:=1).Range.Text = Bolum3
            objDoc.Tables(1).Cell(Row:=BStr + 3, Column:=1).Range.Font.Underline = wdUnderlineSingle
            M4 = True
            
            ustbilgimuhatap = IlKucukHarf & " Provincial Governorship " & Cells(ActiveCell.Row, 64).Value
        ElseIf Cells(ActiveCell.Row, 64).Value = "District Directorate B" Or Cells(ActiveCell.Row, 64).Value = "District Directorate C" Or _
        Cells(ActiveCell.Row, 64).Value = "District Directorate D" Or Cells(ActiveCell.Row, 64).Value = "District Directorate E" Then 'KAYMAKAMLIK
            objDoc.Tables(1).Cell(Row:=BStr, Column:=1).Range.Text = IlceBuyukHarf & " DISTRICT GOVERNORSHIP"
            objDoc.Tables(1).Cell(Row:=BStr + 1, Column:=1).Range.Text = Cells(ActiveCell.Row, 64).Value 'UCase(Replace(Replace(Cells(ActiveCell.Row, 64).Value, "i", "I"), "ı", "I"))
            objDoc.Tables(1).Cell(Row:=BStr + 2, Column:=1).Range.Text = "(" & Cells(ActiveCell.Row, 65).Value & ")"
            objDoc.Tables(1).Cell(Row:=BStr + 3, Column:=1).Range.Text = Bolum3
            objDoc.Tables(1).Cell(Row:=BStr + 3, Column:=1).Range.Font.Underline = wdUnderlineSingle
            M4 = True
            
            ustbilgimuhatap = IlceKucukHarf & " District Governorship " & Cells(ActiveCell.Row, 64).Value
        ElseIf InStr(Cells(ActiveCell.Row, 64).Value, "General Directorate") <> 0 Or InStr(Cells(ActiveCell.Row, 64).Value, "Regional Directorate") <> 0 Then
            objDoc.Tables(1).Cell(Row:=BStr, Column:=1).Range.Text = UCase(Replace(Replace(ThisWorkbook.Worksheets(2).Cells(6, 111).Value, "i", "I"), "ı", "I"))
            objDoc.Tables(1).Cell(Row:=BStr + 1, Column:=1).Range.Text = Cells(ActiveCell.Row, 64).Value 'UCase(Replace(Replace(Cells(ActiveCell.Row, 64).Value, "i", "I"), "ı", "I"))
            objDoc.Tables(1).Cell(Row:=BStr + 2, Column:=1).Range.Text = "(" & Cells(ActiveCell.Row, 65).Value & ")"
            objDoc.Tables(1).Cell(Row:=BStr + 3, Column:=1).Range.Text = Bolum3
            objDoc.Tables(1).Cell(Row:=BStr + 3, Column:=1).Range.Font.Underline = wdUnderlineSingle
            M4 = True

            ustbilgimuhatap = WorksheetFunction.Proper(ThisWorkbook.Worksheets(2).Cells(6, 111).Value) & " " & Cells(ActiveCell.Row, 64).Value
        Else 'YARGI 3'lük
            Bolum1 = UCase(Replace(Replace(Cells(ActiveCell.Row, 64).Value, "i", "I"), "ı", "I"))
            If InStr(Bolum1, "X.X. ") > 0 Then
                objDoc.Tables(1).Cell(Row:=BStr, Column:=1).Range.Text = Mid(Bolum1, 6, Len(Bolum1))
                objDoc.Tables(1).Cell(Row:=BStr + 1, Column:=1).Range.Text = "(" & Cells(ActiveCell.Row, 65).Value & ")"
                objDoc.Tables(1).Cell(Row:=BStr + 2, Column:=1).Range.Text = Bolum3
                objDoc.Tables(1).Cell(Row:=BStr + 2, Column:=1).Range.Font.Underline = wdUnderlineSingle
                M3 = True
                
                ustbilgimuhatap = Mid(Cells(ActiveCell.Row, 64).Value, 6, Len(Cells(ActiveCell.Row, 64).Value))
            Else
                objDoc.Tables(1).Cell(Row:=BStr, Column:=1).Range.Text = Bolum1
                objDoc.Tables(1).Cell(Row:=BStr + 1, Column:=1).Range.Text = "(" & Cells(ActiveCell.Row, 65).Value & ")"
                objDoc.Tables(1).Cell(Row:=BStr + 2, Column:=1).Range.Text = Bolum3
                objDoc.Tables(1).Cell(Row:=BStr + 2, Column:=1).Range.Font.Underline = wdUnderlineSingle
                M3 = True

                ustbilgimuhatap = Cells(ActiveCell.Row, 64).Value
            End If
        End If
    End If
    '3'lük
    If Cells(ActiveCell.Row, 65).Value = "" Then
        If Cells(ActiveCell.Row, 64).Value = "Provincial Directorate B" Or Cells(ActiveCell.Row, 64).Value = "Provincial Directorate C" Or _
        Cells(ActiveCell.Row, 64).Value = "Provincial Directorate D" Or Cells(ActiveCell.Row, 64).Value = "Provincial Directorate E" Then 'VALİLİK
            objDoc.Tables(1).Cell(Row:=BStr, Column:=1).Range.Text = IlBuyukHarf & " PROVINCIAL GOVERNORSHIP"
            objDoc.Tables(1).Cell(Row:=BStr + 1, Column:=1).Range.Text = "(" & Cells(ActiveCell.Row, 64).Value & ")" 'UCase(Replace(Replace(Cells(ActiveCell.Row, 64).Value, "i", "I"), "ı", "I"))
            objDoc.Tables(1).Cell(Row:=BStr + 2, Column:=1).Range.Text = Bolum3
            objDoc.Tables(1).Cell(Row:=BStr + 2, Column:=1).Range.Font.Underline = wdUnderlineSingle
            M3 = True
            ustbilgimuhatap = IlKucukHarf & " Provincial Governorship " & Cells(ActiveCell.Row, 64).Value
        ElseIf Cells(ActiveCell.Row, 64).Value = "District Directorate B" Or Cells(ActiveCell.Row, 64).Value = "District Directorate C" Or _
        Cells(ActiveCell.Row, 64).Value = "District Directorate D" Or Cells(ActiveCell.Row, 64).Value = "District Directorate E" Then
            objDoc.Tables(1).Cell(Row:=BStr, Column:=1).Range.Text = IlceBuyukHarf & " DISTRICT GOVERNORSHIP"
            objDoc.Tables(1).Cell(Row:=BStr + 1, Column:=1).Range.Text = "(" & Cells(ActiveCell.Row, 64).Value & ")"  'UCase(Replace(Replace(Cells(ActiveCell.Row, 64).Value, "i", "I"), "ı", "I"))
            objDoc.Tables(1).Cell(Row:=BStr + 2, Column:=1).Range.Text = Bolum3
            objDoc.Tables(1).Cell(Row:=BStr + 2, Column:=1).Range.Font.Underline = wdUnderlineSingle
            M3 = True
            
            ustbilgimuhatap = IlceKucukHarf & " District Governorship " & Cells(ActiveCell.Row, 64).Value
        ElseIf InStr(Cells(ActiveCell.Row, 64).Value, "General Directorate") <> 0 Or InStr(Cells(ActiveCell.Row, 64).Value, "Regional Directorate") <> 0 Then
            objDoc.Tables(1).Cell(Row:=BStr, Column:=1).Range.Text = UCase(Replace(Replace(ThisWorkbook.Worksheets(2).Cells(6, 111).Value, "i", "I"), "ı", "I"))
            objDoc.Tables(1).Cell(Row:=BStr + 1, Column:=1).Range.Text = "(" & Cells(ActiveCell.Row, 64).Value & ")"  'UCase(Replace(Replace(Cells(ActiveCell.Row, 64).Value, "i", "I"), "ı", "I"))
            objDoc.Tables(1).Cell(Row:=BStr + 2, Column:=1).Range.Text = Bolum3
            objDoc.Tables(1).Cell(Row:=BStr + 2, Column:=1).Range.Font.Underline = wdUnderlineSingle
            M3 = True

            ustbilgimuhatap = WorksheetFunction.Proper(ThisWorkbook.Worksheets(2).Cells(6, 111).Value) & " " & Cells(ActiveCell.Row, 64).Value
        Else 'YARGI 2'lik
            Bolum1 = UCase(Replace(Replace(Cells(ActiveCell.Row, 64).Value, "i", "I"), "ı", "I"))
            If InStr(Bolum1, "X.X. ") > 0 Then
                objDoc.Tables(1).Cell(Row:=BStr, Column:=1).Range.Text = Mid(Bolum1, 6, Len(Bolum1))
                objDoc.Tables(1).Cell(Row:=BStr + 1, Column:=1).Range.Text = Bolum3
                objDoc.Tables(1).Cell(Row:=BStr + 1, Column:=1).Range.Font.Underline = wdUnderlineSingle
                M2 = True

                ustbilgimuhatap = Mid(Cells(ActiveCell.Row, 64).Value, 6, Len(Cells(ActiveCell.Row, 64).Value))
            Else
                objDoc.Tables(1).Cell(Row:=BStr, Column:=1).Range.Text = Bolum1
                objDoc.Tables(1).Cell(Row:=BStr + 1, Column:=1).Range.Text = Bolum3
                objDoc.Tables(1).Cell(Row:=BStr + 1, Column:=1).Range.Font.Underline = wdUnderlineSingle
                M2 = True

                ustbilgimuhatap = Cells(ActiveCell.Row, 64).Value
            End If
        End If
    End If
    
    'İlgi tema
    If Cells(IlkSira, 25).Value = "Provincial Directorate B" Then
        If Cells(IlkSira, 26).Value <> "" Then
            GelenTema = Cells(IlkSira, 17).Value & " Provincial Governorship Provincial Directorate B " & Cells(IlkSira, 26).Value
        Else
            GelenTema = Cells(IlkSira, 17).Value & " Provincial Governorship Provincial Directorate B"
        End If
    ElseIf Cells(IlkSira, 25).Value = "District Directorate B" Then
        If Cells(IlkSira, 26).Value <> "" Then
            GelenTema = Cells(IlkSira, 18).Value & " District Governorship District Directorate B " & Cells(IlkSira, 26).Value
        Else
            GelenTema = Cells(IlkSira, 18).Value & " District Governorship District Directorate B"
        End If
    ElseIf Cells(IlkSira, 25).Value = "Provincial Directorate C" Then
        If Cells(IlkSira, 26).Value <> "" Then
            GelenTema = Cells(IlkSira, 17).Value & " Provincial Governorship Provincial Directorate C " & Cells(IlkSira, 26).Value
        Else
            GelenTema = Cells(IlkSira, 17).Value & " Provincial Governorship Provincial Directorate C"
        End If
    ElseIf Cells(IlkSira, 25).Value = "District Directorate C" Then
        If Cells(IlkSira, 26).Value <> "" Then
            GelenTema = Cells(IlkSira, 18).Value & " District Governorship District Directorate C " & Cells(IlkSira, 26).Value
        Else
            GelenTema = Cells(IlkSira, 18).Value & " District Governorship District Directorate C"
        End If
    ElseIf Cells(IlkSira, 25).Value = "Provincial Directorate D" Then
        If Cells(IlkSira, 26).Value <> "" Then
            GelenTema = Cells(IlkSira, 17).Value & " Provincial Governorship Provincial Directorate D " & Cells(IlkSira, 26).Value
        Else
            GelenTema = Cells(IlkSira, 17).Value & " Provincial Governorship Provincial Directorate D"
        End If
    ElseIf Cells(IlkSira, 25).Value = "District Directorate D" Then
        If Cells(IlkSira, 26).Value <> "" Then
            GelenTema = Cells(IlkSira, 18).Value & " District Governorship District Directorate D " & Cells(IlkSira, 26).Value
        Else
            GelenTema = Cells(IlkSira, 18).Value & " District Governorship District Directorate D"
        End If
    ElseIf Cells(IlkSira, 25).Value = "Provincial Directorate E" Then
        If Cells(IlkSira, 26).Value <> "" Then
            GelenTema = Cells(IlkSira, 17).Value & " Provincial Governorship Provincial Directorate E " & Cells(IlkSira, 26).Value
        Else
            GelenTema = Cells(IlkSira, 17).Value & " Provincial Governorship Provincial Directorate E"
        End If
    ElseIf Cells(IlkSira, 25).Value = "District Directorate E" Then
        If Cells(IlkSira, 26).Value <> "" Then
            GelenTema = Cells(IlkSira, 18).Value & " District Governorship District Directorate E " & Cells(IlkSira, 26).Value
        Else
            GelenTema = Cells(IlkSira, 18).Value & " District Governorship District Directorate E"
        End If
    ElseIf InStr(Cells(IlkSira, 25).Value, "General Directorate") <> 0 Or InStr(Cells(IlkSira, 25).Value, "Regional Directorate") <> 0 Then
        If Cells(IlkSira, 26).Value <> "" Then
            GelenTema = WorksheetFunction.Proper(ThisWorkbook.Worksheets(2).Cells(6, 111).Value) & " " & Cells(IlkSira, 25).Value & " " & Cells(IlkSira, 26).Value
        Else
            GelenTema = WorksheetFunction.Proper(ThisWorkbook.Worksheets(2).Cells(6, 111).Value) & " " & Cells(IlkSira, 25).Value
        End If
    Else
        If InStr(Cells(IlkSira, 25).Value, "X.X. ") > 0 Then
            If Cells(IlkSira, 26).Value <> "" Then
                GelenTema = Mid(Cells(IlkSira, 25).Value, 6, Len(Cells(IlkSira, 25).Value)) & " " & Cells(IlkSira, 26).Value
            Else
                GelenTema = Mid(Cells(IlkSira, 25).Value, 6, Len(Cells(IlkSira, 25).Value))
            End If
        Else
            If Cells(IlkSira, 26).Value <> "" Then
                GelenTema = Cells(IlkSira, 25).Value & " " & Cells(IlkSira, 26).Value
            Else
                GelenTema = Cells(IlkSira, 25).Value
            End If
        End If
    End If

    'İlgi
    Bolum1 = "Delivered to us on "
    Bolum2 = ", with a letter dated "
    Bolum3 = " and reference number "
    Bolum4 = "."
    
    If Cells(ActiveCell.Row, 25).Value = Cells(ActiveCell.Row, 64).Value And _
       Cells(ActiveCell.Row, 26).Value = Cells(ActiveCell.Row, 65).Value And _
       Cells(ActiveCell.Row, 17).Value = Cells(ActiveCell.Row, 69).Value And _
       Cells(ActiveCell.Row, 18).Value = Cells(ActiveCell.Row, 70).Value Then
           
        ' Gelen ve giden birim aynıysa
        objDoc.Tables(1).Cell(Row:=15, Column:=3).Range.Text = _
            Bolum1 & Cells(ActiveCell.Row, 28).Value & Bolum2 & _
            Cells(ActiveCell.Row, 20).Value & Bolum3 & _
            Cells(ActiveCell.Row, 21).Value & Bolum4
    
        ' İlgi yazı eki kontrolü
        If Trim(Cells(ActiveCell.Row, 74).Value) <> "" Then
            Ifv = True
        Else
            Ify = True
        End If
    
    Else
        ' Gelen ve giden birim farklıysa
        RaporIlgi = Bolum1 & Cells(ActiveCell.Row, 28).Value & _
                    " from " & GelenTema & Bolum2 & _
                    Cells(IlkSira, 20).Value & Bolum3 & _
                    Cells(IlkSira, 21).Value & Bolum4
    
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
        Ek4 = "report "
        Ek5 = "Report 1, "
    Else
        Ek4 = "reports "
        Ek5 = "Report 1s, "
    End If
    
    ' Tip A and Tip B
    TipB = False
    TipA = False
    For i = IlkSira To SonSira
        If Cells(i, 38).Value = "Item Type - X2" Or _
           (Cells(i, 38).Value = "Item Type - X1" And Cells(i, 41).Value = "1") Then
            TipB = True
        Else
            TipA = True
        End If
    Next i
    
    AdetTopla = Application.Sum(Range(Cells(IlkSira, 44), Cells(SonSira, 44)))
    
    If TipB = True And TipA = True Then
        Ek1 = "The Type A and Type B items, sent to our unit as an enclosure to the referenced letter, have been documented in "
        Ek2 = "items "
    ElseIf TipB = True Then
        If AdetTopla = 1 Then
            Ek1 = "The Type B item, sent to our unit as an enclosure to the referenced letter, has been documented in "
            Ek2 = "item "
        Else
            Ek1 = "The Type B items, sent to our unit as an enclosure to the referenced letter, have been documented in "
            Ek2 = "items "
        End If
    ElseIf TipA = True Then
        If AdetTopla = 1 Then
            Ek1 = "The Type A item, sent to our unit as an enclosure to the referenced letter, has been documented in "
            Ek2 = "item "
        Else
            Ek1 = "The Type A items, sent to our unit as an enclosure to the referenced letter, have been documented in "
            Ek2 = "items "
        End If
    End If
    
    ' Delivery method
    GonderimUsulu = Mid(Cells(ActiveCell.Row, 67).Value, InStr(Cells(ActiveCell.Row, 67).Value, "/") + 1)
    If GonderimUsulu = "HAND DELIVERY" Then
        GonderimUsulu = "hand-delivered to you."
    Else
        GonderimUsulu = "sent to your office accordingly."
    End If
    
    ' Construct the text
    Bolum1 = Ek1 & Ek5
    Bolum2 = "dated " & Cells(ActiveCell.Row, 60).Value & " and numbered " & Ek3 & ". "
    Bolum3 = "Both the " & Ek2 & "and the corresponding " & Ek4 & "have been " & GonderimUsulu
    Bolum6 = Chr(9) & "Respectfully submitted for your information."
    
    Birlestir = Chr(9) & Bolum1 & Bolum2 & Bolum3
    objDoc.Tables(2).Cell(Row:=1, Column:=1).Range.Text = Birlestir
    objDoc.Tables(2).Cell(Row:=3, Column:=1).Range.Text = Bolum6



    'Üst yazı notu (Darphane notu)
    If TipB = True Then
'        'Exactly
    Else
        objDoc.Tables(2).Cell(Row:=2, Column:=1).Range.Text = ""
        'objDoc.Tables(2).Rows(2).Delete
    End If
    
    'imzalar
    objDoc.Tables(3).Cell(Row:=4, Column:=2).Range.Text = Cells(ActiveCell.Row, 122).Value 'Ad Soyad1
    objDoc.Tables(3).Cell(Row:=5, Column:=2).Range.Text = Cells(ActiveCell.Row, 123).Value 'Unvan1
    objDoc.Tables(3).Cell(Row:=4, Column:=3).Range.Text = Cells(ActiveCell.Row, 125).Value 'Ad Soyad2
    objDoc.Tables(3).Cell(Row:=5, Column:=3).Range.Text = Cells(ActiveCell.Row, 126).Value 'Unvan2

    'Ekler
    objDoc.Tables(3).Cell(Row:=8, Column:=1).Range.Text = "Attachment:"
    
    Ek2 = Cells(ActiveCell.Row, 67).Value 'Kapalı Package A
    Ek2 = Left(Ek2, InStr(Ek2, "/") - 1)
    If Cells(ActiveCell.Row, 68).Value > 1 Then objDoc.Tables(3).Cell(Row:=8, Column:=2).Range.Text = " 1) Enclosed " & Ek2 & " (" & Cells(ActiveCell.Row, 68).Value & " pieces)"
    If Cells(ActiveCell.Row, 68).Value < 2 Then objDoc.Tables(3).Cell(Row:=8, Column:=2).Range.Text = " 1) Enclosed " & Ek2 & " (" & Cells(ActiveCell.Row, 68).Value & " piece)"
    
    x = Application.Sum(Range(Cells(IlkSira, 92), Cells(SonSira, 92))) 'Statement 2 toplam sayfa sayısı
    If x > 1 Then objDoc.Tables(3).Cell(Row:=9, Column:=2).Range.Text = " 2) Statement 2 (" & x & " pages)"
    If x < 2 Then objDoc.Tables(3).Cell(Row:=9, Column:=2).Range.Text = " 2) Statement 2 (" & x & " page)"
    
    x = Application.Sum(Range(Cells(IlkSira, 91), Cells(SonSira, 91)))  'Rapor1 toplam sayfa sayısı
    If y = 1 Then
        If x > 1 Then objDoc.Tables(3).Cell(Row:=10, Column:=2).Range.Text = " 3) Report 1 (" & x & " pages)"
        If x < 2 Then objDoc.Tables(3).Cell(Row:=10, Column:=2).Range.Text = " 3) Report 1 (" & x & " page)"
    Else
        objDoc.Tables(3).Cell(Row:=10, Column:=2).Range.Text = " 3) Report 1 (" & y & " reports, " & "total of " & x & " pages)"
    End If
    
    'Statement 1 (döküm dahil) toplam sayfa sayısı
    If Cells(ActiveCell.Row, 90).Value = "" Then
        x = Application.Sum(Range(Cells(IlkSira, 89), Cells(SonSira, 89)))
        If x > 1 Then objDoc.Tables(3).Cell(Row:=11, Column:=2).Range.Text = " 4) Statement 1 (" & x & " pages)"
        If x < 2 Then objDoc.Tables(3).Cell(Row:=11, Column:=2).Range.Text = " 4) Statement 1 (" & x & " page)"
    Else
        x = Application.Sum(Range(Cells(IlkSira, 89), Cells(SonSira, 90)))
        objDoc.Tables(3).Cell(Row:=11, Column:=2).Range.Text = " 4) Attached Statement 1 (total of " & x & " pages)"
    End If
        
    If Cells(IlkSira, 74).Value <> "" Then 'ilgi yazı fotokopisi var
        'İlgi yazı fotokopisi
        If Cells(IlkSira, 74).Value > 1 Then objDoc.Tables(3).Cell(Row:=12, Column:=2).Range.Text = " 5) Photocopy of the Referenced Letter (" & Cells(IlkSira, 74).Value & " pages)"
        If Cells(IlkSira, 74).Value < 2 Then objDoc.Tables(3).Cell(Row:=12, Column:=2).Range.Text = " 5) Photocopy of the Referenced Letter (" & Cells(IlkSira, 74).Value & " page)"
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
    Govde1FarkSay = Govde1StrSay - 1
    CokluSayfa = 0
    
    'MsgBox IlgiFarkSay
    'MsgBox Govde1FarkSay
    'MsgBox Govde1FarkSay + IlgiFarkSay 'Varsayılan ilgi ve 1 paragrafta (2 rowda) toplam 2 satıra göre sıfırlandı.
    
    '__________________________________'TipB Notu var.


    'Dinamik sayfa düzeni
    If TipB = True Then 'TipB Notu var.
        If Ifv = True Then 'Ekte ilgi fotokopisi var.
            Govde1FarkSay = Govde1FarkSay + 4
        Else
            Govde1FarkSay = Govde1FarkSay + 3
            objDoc.Tables(3).Rows(12).Delete
        End If
    Else
        objDoc.Tables(2).Rows(2).Delete 'TipB notu
        If Ifv = True Then 'Ekte ilgi fotokopisi var.
            Govde1FarkSay = Govde1FarkSay + 1
        Else
            objDoc.Tables(3).Rows(12).Delete
            'Govde1FarkSay = Govde1FarkSay + 0
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
'    objDoc.Close SaveChanges:=True
'    objWord.Quit
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
    Cells(ActiveCell.Row, 93).Value = TotalSayfaUstYazi

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
    Cells(ActiveCell.Row, 70).Value = IlceSakla
End If

'Worksheets(3).Activate
Application.ScreenUpdating = True
Application.DisplayAlerts = True
'Application.EnableEvents = True

'End If


End Sub

Sub IslemGunluguRapor1()
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
IslemGunlugu = IslemGunlukleriKlasor & "System Registry Report 1.xlsx"

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
Set WsRapor = ThisWorkbook.Worksheets(3)
'WsRapor.Unprotect Password:="123"


IlceSakla = ""
If InStr(WsRapor.Cells(IlkSira, 18).Value, " Organization A") <> 0 Then
    IlceSakla = WsRapor.Cells(IlkSira, 18).Value
    WsRapor.Cells(IlkSira, 18).Value = ""
End If


'Aylık ayraçlar
If WsRapor.Cells(IlkSira, 31).Value <> "" Then
    ModulTarih = WsRapor.Cells(IlkSira, 31).Value
    ModulAyrac = "01" & Right(ModulTarih, 8)
Else 'işlemin yapıldığı günü esas al
    ModulTarih = Format(Date, "dd.mm.yyyy")
    ModulAyrac = "01" & Right(ModulTarih, 8)
End If


'İşlem günlüğü açıksa kaydet ve kapat.
OpenControl = IsWorkBookOpen(IslemGunlugu)
If OpenControl = True Then
    Workbooks("System Registry Report 1.xlsx").Save
    Workbooks("System Registry Report 1.xlsx").Close SaveChanges:=False
End If

'İşlem günlüğü aç
Workbooks.Open (IslemGunlugu)
Set WsIslemGunlugu = Workbooks("System Registry Report 1.xlsx").Worksheets(1)

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
If WsRapor.Cells(IlkSira, 25).Value = "Provincial Directorate B" Then
    If WsRapor.Cells(IlkSira, 26).Value <> "" Then
        GelenTema = WsRapor.Cells(IlkSira, 17).Value & " Provincial Governorship Provincial Directorate B " & WsRapor.Cells(IlkSira, 26).Value
    Else
        GelenTema = WsRapor.Cells(IlkSira, 17).Value & " Provincial Governorship Provincial Directorate B"
    End If
ElseIf WsRapor.Cells(IlkSira, 25).Value = "District Directorate B" Then
    If WsRapor.Cells(IlkSira, 26).Value <> "" Then
        GelenTema = WsRapor.Cells(IlkSira, 18).Value & " District Governorship District Directorate B " & WsRapor.Cells(IlkSira, 26).Value
    Else
        GelenTema = WsRapor.Cells(IlkSira, 18).Value & " District Governorship District Directorate B"
    End If
ElseIf WsRapor.Cells(IlkSira, 25).Value = "Provincial Directorate C" Then
    If WsRapor.Cells(IlkSira, 26).Value <> "" Then
        GelenTema = WsRapor.Cells(IlkSira, 17).Value & " Provincial Governorship Provincial Directorate C " & WsRapor.Cells(IlkSira, 26).Value
    Else
        GelenTema = WsRapor.Cells(IlkSira, 17).Value & " Provincial Governorship Provincial Directorate C"
    End If
ElseIf WsRapor.Cells(IlkSira, 25).Value = "District Directorate C" Then
    If WsRapor.Cells(IlkSira, 26).Value <> "" Then
        GelenTema = WsRapor.Cells(IlkSira, 18).Value & " District Governorship District Directorate C " & WsRapor.Cells(IlkSira, 26).Value
    Else
        GelenTema = WsRapor.Cells(IlkSira, 18).Value & " District Governorship District Directorate C"
    End If
ElseIf WsRapor.Cells(IlkSira, 25).Value = "Provincial Directorate D" Then
    If WsRapor.Cells(IlkSira, 26).Value <> "" Then
        GelenTema = WsRapor.Cells(IlkSira, 17).Value & " Provincial Governorship Provincial Directorate D " & WsRapor.Cells(IlkSira, 26).Value
    Else
        GelenTema = WsRapor.Cells(IlkSira, 17).Value & " Provincial Governorship Provincial Directorate D"
    End If
ElseIf WsRapor.Cells(IlkSira, 25).Value = "District Directorate D" Then
    If WsRapor.Cells(IlkSira, 26).Value <> "" Then
        GelenTema = WsRapor.Cells(IlkSira, 18).Value & " District Governorship District Directorate D " & WsRapor.Cells(IlkSira, 26).Value
    Else
        GelenTema = WsRapor.Cells(IlkSira, 18).Value & " District Governorship District Directorate D"
    End If
ElseIf WsRapor.Cells(IlkSira, 25).Value = "Provincial Directorate E" Then
    If WsRapor.Cells(IlkSira, 26).Value <> "" Then
        GelenTema = WsRapor.Cells(IlkSira, 17).Value & " Provincial Governorship Provincial Directorate E " & WsRapor.Cells(IlkSira, 26).Value
    Else
        GelenTema = WsRapor.Cells(IlkSira, 17).Value & " Provincial Governorship Provincial Directorate E"
    End If
ElseIf WsRapor.Cells(IlkSira, 25).Value = "District Directorate E" Then
    If WsRapor.Cells(IlkSira, 26).Value <> "" Then
        GelenTema = WsRapor.Cells(IlkSira, 18).Value & " District Governorship District Directorate E " & WsRapor.Cells(IlkSira, 26).Value
    Else
        GelenTema = WsRapor.Cells(IlkSira, 18).Value & " District Governorship District Directorate E"
    End If
ElseIf InStr(WsRapor.Cells(IlkSira, 25).Value, "General Directorate") <> 0 Or InStr(WsRapor.Cells(IlkSira, 25).Value, "Regional Directorate") <> 0 Then
    If WsRapor.Cells(IlkSira, 26).Value <> "" Then
        GelenTema = WorksheetFunction.Proper(ThisWorkbook.Worksheets(2).Cells(6, 111).Value) & " " & WsRapor.Cells(IlkSira, 25).Value & " " & WsRapor.Cells(IlkSira, 26).Value
    Else
        GelenTema = WorksheetFunction.Proper(ThisWorkbook.Worksheets(2).Cells(6, 111).Value) & " " & WsRapor.Cells(IlkSira, 25).Value
    End If
Else
    If WsRapor.Cells(IlkSira, 26).Value <> "" Then
        If InStr(WsRapor.Cells(IlkSira, 25).Value, "X.X. ") > 0 Then
            GelenTema = Mid(WsRapor.Cells(IlkSira, 25).Value, 6, Len(WsRapor.Cells(IlkSira, 25).Value)) & " " & WsRapor.Cells(IlkSira, 26).Value
        Else
            GelenTema = WsRapor.Cells(IlkSira, 25).Value & " " & WsRapor.Cells(IlkSira, 26).Value
        End If
    Else
        If InStr(WsRapor.Cells(IlkSira, 25).Value, "X.X. ") > 0 Then
            GelenTema = Mid(WsRapor.Cells(IlkSira, 25).Value, 6, Len(WsRapor.Cells(IlkSira, 25).Value))
        Else
            GelenTema = WsRapor.Cells(IlkSira, 25).Value
        End If
    End If
End If


'____________OPERASYONLAR

'İşlem günlüğünde başlangıç ve bitiş satırlarını tespit et.
Say1IslemGunlugu = WsIslemGunlugu.Range("B100000").End(xlUp).Row
Say2IslemGunlugu = WsIslemGunlugu.Range("C100000").End(xlUp).Row
SayAyracIslemGunlugu = WsIslemGunlugu.Range("E100000").End(xlUp).Row

Set IslemGunluguIlkSiraBul = WsIslemGunlugu.Range("B7:B100000").Find(What:=WsRapor.Cells(IlkSira, 85).Value, SearchDirection:=xlNext, _
                SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole) 'zaman damgasını ara
Set IslemGunluguSonSiraBul = WsIslemGunlugu.Range("C7:C100000").Find(What:=WsRapor.Cells(IlkSira, 85).Value, SearchDirection:=xlNext, _
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
WsIslemGunlugu.Cells(ilkrow, 2).Value = WsRapor.Cells(IlkSira, 85).Value
WsIslemGunlugu.Cells(sonrow, 3).Value = WsRapor.Cells(IlkSira, 85).Value
'Verileri yaz
WsIslemGunlugu.Range(WsIslemGunlugu.Cells(ilkrow, 7), WsIslemGunlugu.Cells(sonrow, 7)).Value = WsRapor.Range(WsRapor.Cells(IlkSira, 11), WsRapor.Cells(SonSira, 11)).Value 'Rapor no
WsIslemGunlugu.Cells(ilkrow, 8).Value = WsRapor.Cells(IlkSira, 17).Value 'İl
WsIslemGunlugu.Cells(ilkrow, 9).Value = WsRapor.Cells(IlkSira, 18).Value 'İlçe
WsIslemGunlugu.Cells(ilkrow, 10).Value = GelenTema
WsIslemGunlugu.Cells(ilkrow, 11).Value = WsRapor.Cells(IlkSira, 20).Value 'Belge tarihi
WsIslemGunlugu.Cells(ilkrow, 12).Value = WsRapor.Cells(IlkSira, 21).Value 'Belge no
WsIslemGunlugu.Cells(ilkrow, 13).Value = WsRapor.Cells(IlkSira, 28).Value 'finansal birimya ulaşma tarihi
WsIslemGunlugu.Cells(ilkrow, 14).Value = WsRapor.Cells(IlkSira, 31).Value 'Tespit tarihi
WsIslemGunlugu.Range(WsIslemGunlugu.Cells(ilkrow, 15), WsIslemGunlugu.Cells(sonrow, 15)).Value = WsRapor.Range(WsRapor.Cells(IlkSira, 38), WsRapor.Cells(SonSira, 38)).Value 'Öğe türü
WsIslemGunlugu.Range(WsIslemGunlugu.Cells(ilkrow, 16), WsIslemGunlugu.Cells(sonrow, 16)).Value = WsRapor.Range(WsRapor.Cells(IlkSira, 41), WsRapor.Cells(SonSira, 41)).Value 'Öğe değeri
WsIslemGunlugu.Range(WsIslemGunlugu.Cells(ilkrow, 17), WsIslemGunlugu.Cells(sonrow, 17)).Value = WsRapor.Range(WsRapor.Cells(IlkSira, 44), WsRapor.Cells(SonSira, 44)).Value 'Adet
WsIslemGunlugu.Cells(ilkrow, 18).Value = WsRapor.Cells(IlkSira, 23).Value 'Tema
WsIslemGunlugu.Range(WsIslemGunlugu.Cells(ilkrow, 19), WsIslemGunlugu.Cells(sonrow, 19)).Value = WsRapor.Range(WsRapor.Cells(IlkSira, 50), WsRapor.Cells(SonSira, 50)).Value 'Açıklama

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
    Workbooks("System Registry Report 1.xlsx").Save
    Workbooks("System Registry Report 1.xlsx").Close SaveChanges:=False
End If


If IlceSakla <> "" Then
    WsRapor.Cells(IlkSira, 18).Value = IlceSakla
End If


'WsRapor.Protect Password:="123"

Out:

Application.DisplayAlerts = True
Application.EnableEvents = True
Application.ScreenUpdating = True

End Sub

Sub Rapor1GelenBelgeGiris()
Dim i As Integer, TarihBul As Range
Dim Say As Long, j As Long, Cont As Long, Sno As Variant
Dim LstBx As MSForms.ListBox, ctl As MSForms.Control
Dim LblSira1 As MSForms.Label
Dim LblSira2 As MSForms.Label
Dim Contx As Long

CalTarih = CalTarihTakip

Say = ThisWorkbook.Worksheets(3).Range("CF100000").End(xlUp).Row
If Say < 7 Then
    GoTo Son
End If

'Cont = 0
Contx = 0
TarihiTekrarla1:
Set TarihBul = ThisWorkbook.Worksheets(3).Range("AE6:AE100000").Find(What:=CalTarih, SearchDirection:=xlNext, _
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
        If ThisWorkbook.Worksheets(3).Cells(j, 5).Value <> "" And ThisWorkbook.Worksheets(3).Cells(j, 31).Value <> "" Then
            Cont = Cont + 1
            If ThisWorkbook.Worksheets(3).Cells(j, 26).Value = "" Then
                If Left(ThisWorkbook.Worksheets(3).Cells(j, 25).Value, 3) = "İl " Then
                    Sno = ThisWorkbook.Worksheets(3).Cells(j, 5).Value & " | ÖR | " & ThisWorkbook.Worksheets(3).Cells(j, 21).Value & " | " & ThisWorkbook.Worksheets(3).Cells(j, 69).Value & " " & ThisWorkbook.Worksheets(3).Cells(j, 25).Value
                ElseIf Left(ThisWorkbook.Worksheets(3).Cells(j, 25).Value, 3) = "İlç" Then
                    Sno = ThisWorkbook.Worksheets(3).Cells(j, 5).Value & " | ÖR | " & ThisWorkbook.Worksheets(3).Cells(j, 21).Value & " | " & ThisWorkbook.Worksheets(3).Cells(j, 18).Value & " " & ThisWorkbook.Worksheets(3).Cells(j, 25).Value
                Else
                    Sno = ThisWorkbook.Worksheets(3).Cells(j, 5).Value & " | ÖR | " & ThisWorkbook.Worksheets(3).Cells(j, 21).Value & " | " & ThisWorkbook.Worksheets(3).Cells(j, 25).Value
                End If
            Else
                Sno = ThisWorkbook.Worksheets(3).Cells(j, 5).Value & " | ÖR | " & ThisWorkbook.Worksheets(3).Cells(j, 21).Value & " | " & ThisWorkbook.Worksheets(3).Cells(j, 26).Value
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

Sub Rapor1TeslimTutanaklari()
Dim i As Integer, TarihBul As Range
Dim Say As Long, j As Long, Cont As Long, Sno As Variant
Dim LstBx As MSForms.ListBox, ctl As MSForms.Control
Dim LblSira1 As MSForms.Label
Dim LblSira2 As MSForms.Label
Dim Contx As Long

CalTarih = CalTarihTakip

Say = ThisWorkbook.Worksheets(3).Range("CF100000").End(xlUp).Row
If Say < 7 Then
    GoTo Son
End If

'Cont = 0
Contx = 0
TarihiTekrarla1:
Set TarihBul = ThisWorkbook.Worksheets(3).Range("BW6:BW100000").Find(What:=CalTarih, SearchDirection:=xlNext, _
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
'    MsgBox CalTarih
    Contx = Contx + 1
    If Contx = 100 Then
        GoTo Son
    End If
    GoTo TarihiTekrarla1
End If


'MsgBox CalTarih

Cont = ContTakip
'Cont = 0
For j = Say To TarihBul.Row Step -1
    'If Cont < 50 Then
        If ThisWorkbook.Worksheets(3).Cells(j, 5).Value <> "" And ThisWorkbook.Worksheets(3).Cells(j, 75).Value <> "" Then
            Cont = Cont + 1
            If ThisWorkbook.Worksheets(3).Cells(j, 65).Value <> "" Then
                If Left(ThisWorkbook.Worksheets(3).Cells(j, 64).Value, 3) = "İl " Then
                    Sno = ThisWorkbook.Worksheets(3).Cells(j, 5).Value & " | Ö | " & ThisWorkbook.Worksheets(3).Cells(j, 76).Value & " | " & ThisWorkbook.Worksheets(3).Cells(j, 65).Value
                ElseIf Left(ThisWorkbook.Worksheets(3).Cells(j, 64).Value, 3) = "İlç" Then
                    Sno = ThisWorkbook.Worksheets(3).Cells(j, 5).Value & " | Ö | " & ThisWorkbook.Worksheets(3).Cells(j, 76).Value & " | " & ThisWorkbook.Worksheets(3).Cells(j, 65).Value
                Else
                    Sno = ThisWorkbook.Worksheets(3).Cells(j, 5).Value & " | Ö | " & ThisWorkbook.Worksheets(3).Cells(j, 76).Value & " | " & ThisWorkbook.Worksheets(3).Cells(j, 64).Value & " " & ThisWorkbook.Worksheets(3).Cells(j, 65).Value
                End If
            ElseIf ThisWorkbook.Worksheets(3).Cells(j, 65).Value = "" Then
                If Left(ThisWorkbook.Worksheets(3).Cells(j, 64).Value, 3) = "İl " Then
                    Sno = ThisWorkbook.Worksheets(3).Cells(j, 5).Value & " | Ö | " & ThisWorkbook.Worksheets(3).Cells(j, 76).Value & " | " & ThisWorkbook.Worksheets(3).Cells(j, 69).Value & " " & ThisWorkbook.Worksheets(3).Cells(j, 64).Value
                ElseIf Left(ThisWorkbook.Worksheets(3).Cells(j, 64).Value, 3) = "İlç" Then
                    Sno = ThisWorkbook.Worksheets(3).Cells(j, 5).Value & " | Ö | " & ThisWorkbook.Worksheets(3).Cells(j, 76).Value & " | " & ThisWorkbook.Worksheets(3).Cells(j, 70).Value & " " & ThisWorkbook.Worksheets(3).Cells(j, 64).Value
                Else
                    Sno = ThisWorkbook.Worksheets(3).Cells(j, 5).Value & " | Ö | " & ThisWorkbook.Worksheets(3).Cells(j, 76).Value & " | " & ThisWorkbook.Worksheets(3).Cells(j, 64).Value
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

'MsgBox "Rapor1: " & ContTakip

End Sub

Sub Rapor1VarlikHareketleriGiris()
Dim i As Integer, TarihBul As Range
Dim Say As Long, j As Long, Cont As Long, Sno As Variant
Dim LstBx As MSForms.ListBox, ctl As MSForms.Control
Dim LblSira1 As MSForms.Label
Dim LblSira2 As MSForms.Label
Dim Contx As Long

Dim IlkSiraBul As Range, SonSiraBul As Range, IlkSira As Long, SonSira As Long, k As Long
Dim RaporNoTakip As String, RaporNoKontrol As Integer, SiraNo As Long


CalTarih = CalTarihTakip

Say = ThisWorkbook.Worksheets(3).Range("CF100000").End(xlUp).Row
If Say < 7 Then
    GoTo Son
End If

'Cont = 0
Contx = 0
Set TarihBul = ThisWorkbook.Worksheets(3).Range("AE6:AE100000").Find(What:=CalTarih, SearchDirection:=xlNext, _
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
    If CalTarih = ThisWorkbook.Worksheets(3).Cells(j, 31).Value And CDate(ThisWorkbook.Worksheets(3).Cells(j, 31).Value) <> CDate(ThisWorkbook.Worksheets(3).Cells(j, 96).Value) Then
        If ThisWorkbook.Worksheets(3).Cells(j, 5).Value <> "" And ThisWorkbook.Worksheets(3).Cells(j, 11).Value <> "" Then
            Cont = Cont + 1

            Set IlkSiraBul = Nothing
            Set SonSiraBul = Nothing
            IlkSira = 0
            SonSira = 0
            SiraNo = ThisWorkbook.Worksheets(3).Cells(j, 5).Value
            Set IlkSiraBul = ThisWorkbook.Worksheets(3).Range("CE7:CE100000").Find(What:=SiraNo, SearchDirection:=xlNext, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not IlkSiraBul Is Nothing Then
                IlkSira = IlkSiraBul.Row
            End If
            Set SonSiraBul = ThisWorkbook.Worksheets(3).Range("CF7:CF100000").Find(What:=SiraNo, SearchDirection:=xlNext, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not SonSiraBul Is Nothing Then
                SonSira = SonSiraBul.Row
            End If
            RaporNoKontrol = 0
            For k = IlkSira To SonSira
                If ThisWorkbook.Worksheets(3).Cells(k, 11).Value <> "" Then
                    If RaporNoKontrol = 0 Then
                        RaporNoTakip = ThisWorkbook.Worksheets(3).Cells(k, 11).Value
                        RaporNoKontrol = 1
                    Else
                        RaporNoTakip = RaporNoTakip & ", " & ThisWorkbook.Worksheets(3).Cells(k, 11).Value
                    End If
                End If
            Next k
                        
            If ThisWorkbook.Worksheets(3).Cells(j, 26).Value <> "" Then
                If Left(ThisWorkbook.Worksheets(3).Cells(j, 25).Value, 3) = "İl " Then
                    Sno = ThisWorkbook.Worksheets(3).Cells(j, 5).Value & " | Ö | " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(3).Cells(j, 26).Value
                ElseIf Left(ThisWorkbook.Worksheets(3).Cells(j, 25).Value, 3) = "İlç" Then
                    Sno = ThisWorkbook.Worksheets(3).Cells(j, 5).Value & " | Ö | " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(3).Cells(j, 26).Value
                Else
                    Sno = ThisWorkbook.Worksheets(3).Cells(j, 5).Value & " | Ö | " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(3).Cells(j, 25).Value & " " & ThisWorkbook.Worksheets(3).Cells(j, 26).Value
                End If
            ElseIf ThisWorkbook.Worksheets(3).Cells(j, 26).Value = "" Then
                If Left(ThisWorkbook.Worksheets(3).Cells(j, 25).Value, 3) = "İl " Then
                    Sno = ThisWorkbook.Worksheets(3).Cells(j, 5).Value & " | Ö | " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(3).Cells(j, 17).Value & " " & ThisWorkbook.Worksheets(3).Cells(j, 25).Value
                ElseIf Left(ThisWorkbook.Worksheets(3).Cells(j, 25).Value, 3) = "İlç" Then
                    Sno = ThisWorkbook.Worksheets(3).Cells(j, 5).Value & " | Ö | " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(3).Cells(j, 18).Value & " " & ThisWorkbook.Worksheets(3).Cells(j, 25).Value
                Else
                    Sno = ThisWorkbook.Worksheets(3).Cells(j, 5).Value & " | Ö | " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(3).Cells(j, 25).Value
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
Next j

ContTakipGiris = Cont

Son:

End Sub

Sub Rapor1VarlikHareketleriMevcut()
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

Say = ThisWorkbook.Worksheets(3).Range("CF100000").End(xlUp).Row
If Say < 7 Then
    GoTo Son
End If

'Mevcutta satırı bul
'CalTarih = CDate(CalTarih)
MevcutSatir = 0
For j = Say To 7 Step -1
    If ThisWorkbook.Worksheets(3).Range("AE" & j).Value <> "" Then
        TarihAra = ThisWorkbook.Worksheets(3).Range("AE" & j)
        If CDate(TarihAra) < CDate(CalTarih) Then
            MevcutSatir = j
            'MsgBox Range("AE" & MevcutSatir)
            GoTo DonguSonu
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
    If ThisWorkbook.Worksheets(3).Cells(j, 31).Value <> "" Then 'And ThisWorkbook.Worksheets(3).Cells(j, 96).Value = "" Then
        TarihAra = ThisWorkbook.Worksheets(3).Range("AE" & j)
        If CDate(TarihAra) < CDate(CalTarih) Then
            If (ThisWorkbook.Worksheets(3).Cells(j, 96).Value <> "" And CDate(ThisWorkbook.Worksheets(3).Cells(j, 96).Value) > CDate(CalTarih)) Or _
                ThisWorkbook.Worksheets(3).Cells(j, 96).Value = "" Then
                'MsgBox TarihAra & " < " & CalTarih
                If ThisWorkbook.Worksheets(3).Cells(j, 5).Value <> "" And ThisWorkbook.Worksheets(3).Cells(j, 11).Value <> "" Then
                    Cont = Cont + 1

                    Set IlkSiraBul = Nothing
                    Set SonSiraBul = Nothing
                    IlkSira = 0
                    SonSira = 0
                    SiraNo = ThisWorkbook.Worksheets(3).Cells(j, 5).Value
                    Set IlkSiraBul = ThisWorkbook.Worksheets(3).Range("CE7:CE100000").Find(What:=SiraNo, SearchDirection:=xlNext, _
                                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
                    If Not IlkSiraBul Is Nothing Then
                        IlkSira = IlkSiraBul.Row
                    End If
                    Set SonSiraBul = ThisWorkbook.Worksheets(3).Range("CF7:CF100000").Find(What:=SiraNo, SearchDirection:=xlNext, _
                                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
                    If Not SonSiraBul Is Nothing Then
                        SonSira = SonSiraBul.Row
                    End If
                    RaporNoKontrol = 0
                    For k = IlkSira To SonSira
                        If ThisWorkbook.Worksheets(3).Cells(k, 11).Value <> "" Then
                            If RaporNoKontrol = 0 Then
                                RaporNoTakip = ThisWorkbook.Worksheets(3).Cells(k, 11).Value
                                RaporNoKontrol = 1
                            Else
                                RaporNoTakip = RaporNoTakip & ", " & ThisWorkbook.Worksheets(3).Cells(k, 11).Value
                            End If
                        End If
                    Next k
            
                    If ThisWorkbook.Worksheets(3).Cells(j, 26).Value <> "" Then
                        If Left(ThisWorkbook.Worksheets(3).Cells(j, 25).Value, 3) = "İl " Then
                            Sno = ThisWorkbook.Worksheets(3).Cells(j, 5).Value & " | Ö | " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(3).Cells(j, 26).Value
                        ElseIf Left(ThisWorkbook.Worksheets(3).Cells(j, 25).Value, 3) = "İlç" Then
                            Sno = ThisWorkbook.Worksheets(3).Cells(j, 5).Value & " | Ö | " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(3).Cells(j, 26).Value
                        Else
                            Sno = ThisWorkbook.Worksheets(3).Cells(j, 5).Value & " | Ö | " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(3).Cells(j, 25).Value & " " & ThisWorkbook.Worksheets(3).Cells(j, 26).Value
                        End If
                    ElseIf ThisWorkbook.Worksheets(3).Cells(j, 26).Value = "" Then
                        If Left(ThisWorkbook.Worksheets(3).Cells(j, 25).Value, 3) = "İl " Then
                            Sno = ThisWorkbook.Worksheets(3).Cells(j, 5).Value & " | Ö | " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(3).Cells(j, 17).Value & " " & ThisWorkbook.Worksheets(3).Cells(j, 25).Value
                        ElseIf Left(ThisWorkbook.Worksheets(3).Cells(j, 25).Value, 3) = "İlç" Then
                            Sno = ThisWorkbook.Worksheets(3).Cells(j, 5).Value & " | Ö | " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(3).Cells(j, 18).Value & " " & ThisWorkbook.Worksheets(3).Cells(j, 25).Value
                        Else
                            Sno = ThisWorkbook.Worksheets(3).Cells(j, 5).Value & " | Ö | " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(3).Cells(j, 25).Value
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
Next j

ContTakipMevcut = Cont

Son:


End Sub

Sub Rapor1VarlikHareketleriCikis()
Dim i As Integer, TarihBul As Range
Dim Say As Long, j As Long, Cont As Long, Sno As Variant
Dim LstBx As MSForms.ListBox, ctl As MSForms.Control
Dim LblSira1 As MSForms.Label
Dim LblSira2 As MSForms.Label
Dim Contx As Long

Dim IlkSiraBul As Range, SonSiraBul As Range, IlkSira As Long, SonSira As Long, k As Long
Dim RaporNoTakip As String, RaporNoKontrol As Integer, SiraNo As Long


CalTarih = CalTarihTakip

Say = ThisWorkbook.Worksheets(3).Range("CF100000").End(xlUp).Row
If Say < 7 Then
    GoTo Son
End If

'Cont = 0
Contx = 0
Set TarihBul = ThisWorkbook.Worksheets(3).Range("CR6:CR100000").Find(What:=CalTarih, SearchDirection:=xlNext, _
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
    If CalTarih = ThisWorkbook.Worksheets(3).Cells(j, 96).Value And ThisWorkbook.Worksheets(3).Cells(j, 31).Value <> "" Then
        If ThisWorkbook.Worksheets(3).Cells(j, 5).Value <> "" And ThisWorkbook.Worksheets(3).Cells(j, 11).Value <> "" Then
            Cont = Cont + 1

            Set IlkSiraBul = Nothing
            Set SonSiraBul = Nothing
            IlkSira = 0
            SonSira = 0
            SiraNo = ThisWorkbook.Worksheets(3).Cells(j, 5).Value
            Set IlkSiraBul = ThisWorkbook.Worksheets(3).Range("CE7:CE100000").Find(What:=SiraNo, SearchDirection:=xlNext, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not IlkSiraBul Is Nothing Then
                IlkSira = IlkSiraBul.Row
            End If
            Set SonSiraBul = ThisWorkbook.Worksheets(3).Range("CF7:CF100000").Find(What:=SiraNo, SearchDirection:=xlNext, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not SonSiraBul Is Nothing Then
                SonSira = SonSiraBul.Row
            End If
            RaporNoKontrol = 0
            For k = IlkSira To SonSira
                If ThisWorkbook.Worksheets(3).Cells(k, 11).Value <> "" Then
                    If RaporNoKontrol = 0 Then
                        RaporNoTakip = ThisWorkbook.Worksheets(3).Cells(k, 11).Value
                        RaporNoKontrol = 1
                    Else
                        RaporNoTakip = RaporNoTakip & ", " & ThisWorkbook.Worksheets(3).Cells(k, 11).Value
                    End If
                End If
            Next k
            
            If ThisWorkbook.Worksheets(3).Cells(j, 26).Value <> "" Then
                If Left(ThisWorkbook.Worksheets(3).Cells(j, 25).Value, 3) = "İl " Then
                    Sno = ThisWorkbook.Worksheets(3).Cells(j, 5).Value & " | Ö | " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(3).Cells(j, 26).Value
                ElseIf Left(ThisWorkbook.Worksheets(3).Cells(j, 25).Value, 3) = "İlç" Then
                    Sno = ThisWorkbook.Worksheets(3).Cells(j, 5).Value & " | Ö | " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(3).Cells(j, 26).Value
                Else
                    Sno = ThisWorkbook.Worksheets(3).Cells(j, 5).Value & " | Ö | " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(3).Cells(j, 25).Value & " " & ThisWorkbook.Worksheets(3).Cells(j, 26).Value
                End If
            ElseIf ThisWorkbook.Worksheets(3).Cells(j, 26).Value = "" Then
                If Left(ThisWorkbook.Worksheets(3).Cells(j, 25).Value, 3) = "İl " Then
                    Sno = ThisWorkbook.Worksheets(3).Cells(j, 5).Value & " | Ö | " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(3).Cells(j, 17).Value & " " & ThisWorkbook.Worksheets(3).Cells(j, 25).Value
                ElseIf Left(ThisWorkbook.Worksheets(3).Cells(j, 25).Value, 3) = "İlç" Then
                    Sno = ThisWorkbook.Worksheets(3).Cells(j, 5).Value & " | Ö | " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(3).Cells(j, 18).Value & " " & ThisWorkbook.Worksheets(3).Cells(j, 25).Value
                Else
                    Sno = ThisWorkbook.Worksheets(3).Cells(j, 5).Value & " | Ö | " & RaporNoTakip & " | " & ThisWorkbook.Worksheets(3).Cells(j, 25).Value
                End If
            End If

            Set LstBx = core_asset_manager_UI.FrameCikis.Controls.Add("Forms.ListBox.1", "Frame1LstBx" & Cont)
            With LstBx
                .Top = (Cont - 1) * 12
                .Left = 18
                .Height = 12
                .Width = 300
                If ThisWorkbook.Worksheets(3).Cells(j, 31).Value = ThisWorkbook.Worksheets(3).Cells(j, 96).Value Then
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
Next j

ContTakipCikis = Cont

Son:

End Sub



