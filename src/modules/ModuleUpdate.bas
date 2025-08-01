Attribute VB_Name = "ModuleUpdate"
Option Explicit
Public OpenWordTakip As Boolean
Public ResetControl As Boolean
Dim TasiyiciArray(1 To 6) As String
Dim gecersizKontrol  As Boolean, gecerliKontrol As Boolean

Sub Update()
Dim AutoPath As String, Tanimlar As String, DestTanimlar As String
Dim SourceUpdate As String, SourceUpdateFile As String, FileName As String
Dim DestOpUserFolderName As String, DestOpUserFolder As String, DestOperasyon As String
Dim OpenKontrolName As String, OpenControl As String

Dim DestTanimlarUpdate As String, TanimlarUpdate As String, FileNameUpdate As String
Dim oldName As String, newName As String

Dim WsSKP As Object, WsSkpUpdate As Object, WsTanimlar As Object, WsTanimlarUpdate As Object
Dim NameDuzenleyici As String, i As Long, j As Long, SayUpdate As Long

Dim WsRapor1 As Object, WsRapor As Object, WsRapor3 As Object, WsTutanak As Object
Dim WsRapor1Update As Object, WsRaporUpdate As Object, WsRapor3Update As Object, WsTutanakUpdate As Object
Dim SayA As Long, SayB As Long, SayC As Long, SayD As Long

Dim IslemGunlugu As String, Rapor2_2IslemGunlugu As String, Rapor3IslemGunlugu As String, Rapor1IslemGunlugu As String, RaporIslemGunlugu As String
Dim IslemGunluguUpdate As String, Rapor2_2IslemGunluguUpdate As String, Rapor3IslemGunluguUpdate As String, Rapor1IslemGunluguUpdate As String, RaporIslemGunluguUpdate As String

Dim WsRapor2_2IslemGunlugu As Object, WsRapor2_2IslemGunluguUpdate As Object
Dim WsRapor3IslemGunlugu As Object, WsRapor3IslemGunluguUpdate As Object
Dim WsRapor1IslemGunlugu As Object, WsRapor1IslemGunluguUpdate As Object
Dim WsRaporIslemGunlugu As Object, WsRaporIslemGunluguUpdate As Object
Dim Rapor2Sablonu As String, RaporSablonuUpdate As String, Normalgecersiz As String, Normalgecerli As String
Dim KontrolFile As String, StrNotlar As String, NotlarUpdate As String, AltBilgi As String, RaporUstYaziSablonuUpdate As String

Dim objWord As Object, objDoc As Object
Dim TasiyiciIlk As String, TasiyiciSon As String
Dim fso As Object, MyRange As Object
Dim TextLine As String, HedefFile As String, a As Long, Cont As Long
Dim Bilgi As Variant

Dim Kurum_ANoStr As String, Birimx As String, Kurum_A As String, YeniVersiyon As Boolean

Dim FarkSayfaKontrol As Integer, SayFarkGirisRapor1 As Long, SayFarkGirisRapor As Long
    

'System Reset prosedürü buraya (aşağıya) da gömüldüğü için alttaki yeşil kodlar resetlemeden sonra başlayacak

Call ModuleRibbon.AnaSayfa

'Pathfinder...
AutoPath = ThisWorkbook.Path
DestOperasyon = AutoPath & "\System Files\Operation\"

SourceUpdate = AutoPath & "\System Files\Update\"
SourceUpdateFile = AutoPath & "\System Files\Update\Update.xlsm"

DestTanimlar = AutoPath & "\System Files\System Definitions\"
Tanimlar = AutoPath & "\System Files\System Definitions\Definitions.xlsx"

DestTanimlarUpdate = AutoPath & "\System Files\Update\System Files\System Definitions\"
TanimlarUpdate = AutoPath & "\System Files\Update\System Files\System Definitions\Definitions.xlsx"

IslemGunlugu = AutoPath & "\System Files\System Templates\Registry Reports\"
Rapor2_2IslemGunlugu = AutoPath & "\System Files\System Templates\Registry Reports\System Registry Report 2.2.xlsx"
'Rapor3IslemGunlugu = AutoPath & "\System Files\System Templates\Registry Reports\Rapor3 Sistem İşlem Günlüğü.xlsx"
Rapor1IslemGunlugu = AutoPath & "\System Files\System Templates\Registry Reports\System Registry Report 1.xlsx"
RaporIslemGunlugu = AutoPath & "\System Files\System Templates\Registry Reports\System Registry Report 2.1.xlsx"

IslemGunluguUpdate = AutoPath & "\System Files\Update\System Files\System Templates\Registry Reports\"
Rapor2_2IslemGunluguUpdate = AutoPath & "\System Files\Update\System Files\System Templates\Registry Reports\System Registry Report 2.2.xlsx"
'Rapor3IslemGunluguUpdate = AutoPath & "\System Files\Update\System Files\System Templates\Registry Reports\Rapor3 Sistem İşlem Günlüğü.xlsx"
Rapor1IslemGunluguUpdate = AutoPath & "\System Files\Update\System Files\System Templates\Registry Reports\System Registry Report 1.xlsx"
RaporIslemGunluguUpdate = AutoPath & "\System Files\Update\System Files\System Templates\Registry Reports\System Registry Report 2.1.xlsx"

Rapor2Sablonu = AutoPath & "\System Files\System Templates\Report 2 Templates\"
Normalgecersiz = AutoPath & "\System Files\System Templates\Report 2 Templates\Standard Invalid.docm"
Normalgecerli = AutoPath & "\System Files\System Templates\Report 2 Templates\Standard Valid.docm"
RaporSablonuUpdate = AutoPath & "\System Files\Update\System Files\System Templates\Report 2 Templates\"

NotlarUpdate = AutoPath & "\System Files\Update\System Files\System Templates\Item Notes\"
StrNotlar = AutoPath & "\System Files\System Templates\Item Notes\"

AltBilgi = AutoPath & "\System Files\System Templates\Footer Field\"

RaporUstYaziSablonuUpdate = AutoPath & "\System Files\Update\System Files\System Templates\Report 2 Cover Letter Templates\"


'Birden fazla kullanıcının operasyon klasöründe çakışmasını önlemek için operasyon klasörünün içinde yeni klasör aç.
DestOpUserFolderName = "Operation-" & Left(ThisWorkbook.name, InStr(ThisWorkbook.name, ".") - 1) '& " " & Format(Now(), "ddmmyyyyhhmmss")
DestOpUserFolder = DestOperasyon & DestOpUserFolderName & "\"

'Check the System Files folder name.
If Not Dir(AutoPath & "\System Files\", vbDirectory) <> vbNullString Then
    MsgBox "Cannot access the folder " & AutoPath & "\System Files\" & ". The folder named 'System Files' might have been renamed or deleted.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

'Check the Operations folder name.
If Not Dir(DestOperasyon, vbDirectory) <> vbNullString Then
    MsgBox "Cannot access the folder " & DestOperasyon & ". The folder named 'Operation' might have been renamed or deleted.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If


'Operation klasörü içinde kullanıcı modülü klasörü yoksa oluştur.
If Not Dir(DestOpUserFolder, vbDirectory) <> vbNullString Then
    MkDir DestOpUserFolder
End If

'________________________________

'Check the Definitions folder name.
If Not Dir(DestTanimlar, vbDirectory) <> vbNullString Then
    MsgBox "Cannot access the folder " & DestTanimlar & ". The folder named 'System Definitions' might have been renamed or deleted.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

'Check the Definitions file name.
If Not Dir(Tanimlar, vbDirectory) <> vbNullString Then
    MsgBox "Cannot access the folder " & Tanimlar & ". The file named 'System Definitions' might have been renamed or deleted.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

'Check the Transaction Logs folder name.
If Not Dir(IslemGunlugu, vbDirectory) <> vbNullString Then
    MsgBox "Cannot access the folder " & IslemGunlugu & ". The folder named 'Registry Reports' might have been renamed or deleted.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

'Check the Report2_2 System Transaction Log file name.
If Not Dir(Rapor2_2IslemGunlugu, vbDirectory) <> vbNullString Then
    MsgBox "Cannot access the folder " & Rapor2_2IslemGunlugu & ". The file named 'System Registry Report 2.2' might have been renamed or deleted.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

''Check the Report3 System Transaction Log file name.
'If Not Dir(Rapor3IslemGunlugu, vbDirectory) <> vbNullString Then
'    MsgBox "Cannot access the folder " & Rapor3IslemGunlugu & ". The file named 'Rapor3 Sistem İşlem Günlüğü' might have been renamed or deleted.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
'    GoTo Son
'End If

'Check the Report1 System Transaction Log file name.
If Not Dir(Rapor1IslemGunlugu, vbDirectory) <> vbNullString Then
    MsgBox "Cannot access the folder " & Rapor1IslemGunlugu & ". The file named 'System Registry Report 1' might have been renamed or deleted.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

'Check the Report System Transaction Log file name.
If Not Dir(RaporIslemGunlugu, vbDirectory) <> vbNullString Then
    MsgBox "Cannot access the folder " & RaporIslemGunlugu & ". The file named 'System Registry Report 2.1' might have been renamed or deleted.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

'Check the Report Template folder name.
If Not Dir(Rapor2Sablonu, vbDirectory) <> vbNullString Then
    MsgBox "Cannot access the folder " & Rapor2Sablonu & ". The folder named 'Rapor Şablonu' might have been renamed or deleted.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

'Check the Standard Invalid file name.
If Not Dir(Normalgecersiz, vbDirectory) <> vbNullString Then
    MsgBox "Cannot access the folder " & Normalgecersiz & ". The file named 'Standard Invalid' might have been renamed or deleted.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

'Check the Notes folder name.
If Not Dir(StrNotlar, vbDirectory) <> vbNullString Then
    MsgBox "Cannot access the folder " & StrNotlar & ". The folder named 'Item Notes' might have been renamed or deleted.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

'Check the Footer Area folder name.
If Not Dir(AltBilgi, vbDirectory) <> vbNullString Then
    MsgBox "Cannot access the folder " & AltBilgi & ". The folder named 'Footer Field' might have been renamed or deleted.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If


'________________________________

'Decision for update
Bilgi = MsgBox("Click " & """" & "Yes" & """" & " to start the update process, or " & """" & "No" & """" & " to cancel.", vbYesNo + vbQuestion, "Enterprise Document Automation System")
If Bilgi = vbYes Then
    GoTo UpdateDevam
ElseIf Bilgi = vbNo Then
    GoTo Son
Else
    GoTo Son
End If
UpdateDevam:

'Check Update folder name.
If Not Dir(SourceUpdate, vbDirectory) <> vbNullString Then
    MsgBox "The requested operation concerns updating the system. Please contact İshak Kutlu, who designed and developed the system, for technical support.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

If Not Dir(SourceUpdate, vbDirectory) = vbNullString Then 'If Update folder exists, check further
    'Check Update file name.
    If Not Dir(SourceUpdateFile, vbDirectory) <> vbNullString Then
        MsgBox "The requested operation concerns updating the system. Please contact İshak Kutlu, who designed and developed the system, for technical support.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
End If


'Check the folder name of System Files inside Update.
If Not Dir(SourceUpdate, vbDirectory) = vbNullString Then 'If Update folder exists, check
    If Not Dir(AutoPath & "\System Files\Update\System Files\", vbDirectory) <> vbNullString Then
        MsgBox "Cannot access the folder " & AutoPath & "\System Files\Update\System Files\" & ". The folder named 'System Files' might have been renamed or deleted.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
End If

If Not Dir(SourceUpdate, vbDirectory) = vbNullString Then 'If Update folder exists, check
    'Check the System Definitions folder name inside Update.
    If Not Dir(DestTanimlarUpdate, vbDirectory) <> vbNullString Then
        MsgBox "Cannot access the folder " & DestTanimlarUpdate & ". The folder named 'System Definitions' might have been renamed or deleted.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
End If

If Not Dir(SourceUpdate, vbDirectory) = vbNullString Then 'If Update folder exists, check
    'Check the System Definitions file name inside Update.
    If Not Dir(TanimlarUpdate, vbDirectory) <> vbNullString Then
        MsgBox "Cannot access the folder " & TanimlarUpdate & ". The file named 'System Definitions' might have been renamed or deleted.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
End If

If Not Dir(SourceUpdate, vbDirectory) = vbNullString Then 'If Update folder exists, check
    'Check the Registry Reports folder name.
    If Not Dir(IslemGunluguUpdate, vbDirectory) <> vbNullString Then
        MsgBox "Cannot access the folder " & IslemGunluguUpdate & ". The folder named 'Registry Reports' might have been renamed or deleted.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
End If

If Not Dir(SourceUpdate, vbDirectory) = vbNullString Then 'If Update folder exists, check
    'Check the System Registry Report 2.2 file name.
    If Not Dir(Rapor2_2IslemGunluguUpdate, vbDirectory) <> vbNullString Then
        MsgBox "Cannot access the folder " & Rapor2_2IslemGunluguUpdate & ". The file named 'System Registry Report 2.2' might have been renamed or deleted.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
End If

'If Not Dir(SourceUpdate, vbDirectory) = vbNullString Then 'If Update folder exists, check
'    'Check the Rapor3 Sistem İşlem Günlüğü file name.
'    If Not Dir(Rapor3IslemGunluguUpdate, vbDirectory) <> vbNullString Then
'        MsgBox "Cannot access the folder " & Rapor3IslemGunluguUpdate & ". The file named 'Rapor3 Sistem İşlem Günlüğü' might have been renamed or deleted.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
'        GoTo Son
'    End If
'End If

If Not Dir(SourceUpdate, vbDirectory) = vbNullString Then 'If Update folder exists, check
    'Check the System Registry Report 1 file name.
    If Not Dir(Rapor1IslemGunluguUpdate, vbDirectory) <> vbNullString Then
        MsgBox "Cannot access the folder " & Rapor1IslemGunluguUpdate & ". The file named 'System Registry Report 1' might have been renamed or deleted.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
End If

If Not Dir(SourceUpdate, vbDirectory) = vbNullString Then 'If Update folder exists, check
    'Check the System Registry Report 2.1 file name.
    If Not Dir(RaporIslemGunluguUpdate, vbDirectory) <> vbNullString Then
        MsgBox "Cannot access the folder " & RaporIslemGunluguUpdate & ". The file named 'System Registry Report 2.1' might have been renamed or deleted.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
End If

If Not Dir(SourceUpdate, vbDirectory) = vbNullString Then 'If Update folder exists, check
    'Check the Rapor Şablonu Update folder name.
    If Not Dir(RaporSablonuUpdate, vbDirectory) <> vbNullString Then
        MsgBox "Cannot access the folder " & RaporSablonuUpdate & ". The folder named 'Rapor Şablonu' might have been renamed or deleted.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
End If

If Not Dir(SourceUpdate, vbDirectory) = vbNullString Then 'If Update folder exists, check
    'Check the Item Notes Update folder name.
    If Not Dir(NotlarUpdate, vbDirectory) <> vbNullString Then
        MsgBox "Cannot access the folder " & NotlarUpdate & ". The folder named 'Item Notes' might have been renamed or deleted.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
End If

If Not Dir(SourceUpdate, vbDirectory) = vbNullString Then 'If Update folder exists, check
    'Check the Report 2 Cover Letter Templates Update folder name.
    If Not Dir(RaporUstYaziSablonuUpdate, vbDirectory) <> vbNullString Then
        MsgBox "Cannot access the folder " & RaporUstYaziSablonuUpdate & ". The folder named 'Report 2 Cover Letter Templates' might have been renamed or deleted.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
End If


'GoTo Atla
ResetControl = True
GlobalResetKapsami = 1 'Varlik çıkışı yapılmayanlar dahil, yani tümü
Call ModuleRibbon.ResetProsedur

Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.EnableEvents = False


ThisWorkbook.Unprotect "123"
ThisWorkbook.Worksheets(2).Unprotect Password:="123"


'__________________Eski Tanımlardan Yeni Tanımlara ve Skp sayfasına data transferi (Başlangıç)

'Update içindeki dosya
FileName = "Definitions.xlsx"
OpenControl = IsFileOpen(DestTanimlarUpdate & FileName)
If OpenControl = True Then 'Açıksa
    Workbooks(FileName).Close SaveChanges:=True
ElseIf OpenControl = False Then
    '
End If

'Normal sistem klasörleri içindeki dosya
FileName = "Definitions.xlsx"
OpenControl = IsFileOpen(DestTanimlar & FileName)
If OpenControl = True Then 'Açıksa
    Workbooks(FileName).Close SaveChanges:=True
ElseIf OpenControl = False Then
    '
End If
Workbooks.Open (DestTanimlar & FileName)
'Workbooks(FileName).Worksheets(1).Activate

'Update System Definitions isimli dosya adını değiştir.
oldName = TanimlarUpdate
newName = AutoPath & "\System Files\Update\System Files\System Definitions\Definitions Update.xlsx"
Name oldName As newName
'MsgBox newName

FileNameUpdate = "Definitions Update.xlsx"
OpenControl = IsFileOpen(newName)
If OpenControl = True Then 'Açıksa
    Workbooks(FileNameUpdate).Close SaveChanges:=True
ElseIf OpenControl = False Then
    '
End If
Workbooks.Open (newName)


Set WsSKP = ThisWorkbook.Worksheets(2)
Set WsTanimlar = Workbooks(FileName).Worksheets(1)
Set WsTanimlarUpdate = Workbooks(FileNameUpdate).Worksheets(1)

WsTanimlar.Unprotect Password:="123"
WsTanimlarUpdate.Unprotect Password:="123"


'System Definitions dosyası güncellemesi
WsTanimlarUpdate.Range("C4:EG305").Copy
WsTanimlar.Range("C4:EG305").PasteSpecial xlPasteValues
WsTanimlar.Cells(5, 122).Value = "Session Name (Initials)"
WsTanimlar.Cells(5, 123).Value = "Full Name (Initials)"
WsTanimlar.Cells(5, 124).Value = "Title (Initials)"
WsTanimlar.Cells(5, 125).Value = "Staff ID (Initials)"
WsTanimlar.Cells(5, 126).Value = "Phone (Initials)"
WsTanimlar.Cells(5, 136).Value = "Editor 1"
WsTanimlar.Cells(5, 137).Value = "Editor 2"

WsTanimlar.Cells(5, 115).Value = "Label / Receipt Printer Definitions"

WsTanimlar.Cells(5, 110).Value = "Signature Type"
WsTanimlar.Cells(6, 110).Value = "Authorized Signature"
WsTanimlar.Cells(7, 110).Value = "Proxy Signature"
WsTanimlar.Cells(8, 110).Value = "Temporary Signature"
WsTanimlar.Cells(9, 110).Value = "Representative Signature"

WsTanimlar.Cells(5, 111).Value = "Ministry"
If WsTanimlarUpdate.Cells(6, 111).Value <> "" Then
    WsTanimlarUpdate.Cells(6, 111).Copy
    WsTanimlar.Cells(6, 111).PasteSpecial xlPasteValues
Else
    WsTanimlar.Cells(6, 111).Value = "MINISTRY of XXX"
End If


'Skp'de yer alan name managerların tümünü temizle.
On Error Resume Next
For i = 6 To 95
    If WsSKP.Range("F" & i).Value <> "" Then
        NameDuzenleyici = Replace(WsSKP.Range("F" & i).Value, " ", "_")
        ThisWorkbook.Names(NameDuzenleyici).Delete
    End If
Next i
On Error GoTo 0

'Skp sayfası güncellemesi
WsTanimlarUpdate.Range("C4:EG305").Copy
WsSKP.Range("C4:EG305").PasteSpecial xlPasteValues
WsSKP.Cells(5, 122).Value = "Session Name (Initials)"
WsSKP.Cells(5, 123).Value = "Full Name (Initials)"
WsSKP.Cells(5, 124).Value = "Title (Initials)"
WsSKP.Cells(5, 125).Value = "Staff ID (Initials)"
WsSKP.Cells(5, 126).Value = "Phone (Initials)"
WsSKP.Cells(5, 136).Value = "Editor 1"
WsSKP.Cells(5, 137).Value = "Editor 2"

WsSKP.Cells(5, 115).Value = "Label / Receipt Printer Definitions"

WsSKP.Cells(5, 110).Value = "Signature Type"
WsSKP.Cells(6, 110).Value = "Authorized Signature"
WsSKP.Cells(7, 110).Value = "Proxy Signature"
WsSKP.Cells(8, 110).Value = "Temporary Signature"
WsSKP.Cells(9, 110).Value = "Representative Signature"


WsSKP.Cells(5, 111).Value = "Ministry"
If WsTanimlarUpdate.Cells(6, 111).Value <> "" Then
    WsTanimlarUpdate.Cells(6, 111).Copy
    WsSKP.Cells(6, 111).PasteSpecial xlPasteValues
Else
    WsSKP.Cells(6, 111).Value = "MINISTRY of XXX"
End If


'Skp'de yer alan name managerların tümünü tekrar oluştur.
On Error Resume Next
For i = 6 To 95
    If WsSKP.Range("F" & i).Value <> "" Then
        NameDuzenleyici = Replace(WsSKP.Range("F" & i).Value, " ", "_")
        ThisWorkbook.Names.Add name:=NameDuzenleyici, RefersTo:=WsSKP.Range(WsSKP.Cells(6, i + 1), WsSKP.Cells(55, i + 1))
    End If
Next i
On Error GoTo 0

WsTanimlar.Protect Password:="123"
WsTanimlarUpdate.Protect Password:="123"


'System Definitions Update dosyasını kapat
FileNameUpdate = "Definitions Update.xlsx"
OpenControl = IsFileOpen(newName)
If OpenControl = True Then 'Açıksa
    Workbooks(FileNameUpdate).Close SaveChanges:=True
ElseIf OpenControl = False Then
    '
End If

'System Definitions Update isimli dosya adını eski haline getir.
oldName = AutoPath & "\System Files\Update\System Files\System Definitions\Definitions Update.xlsx"
newName = TanimlarUpdate
Name oldName As newName

'System Definitions dosyasını kapat
FileName = "Definitions.xlsx"
OpenControl = IsFileOpen(DestTanimlar & FileName)
If OpenControl = True Then 'Açıksa
    Workbooks(FileName).Close SaveChanges:=True
ElseIf OpenControl = False Then
    '
End If

'__________________Eski Tanımlardan Yeni Tanımlara ve Skp sayfasına data transferi (Bitiş)


'__________________Eski Uygulama dosyasından yeni uygulama dosyasına data transferi (Başlangıç)


FileNameUpdate = "Update.xlsm"
OpenControl = IsFileOpen(SourceUpdateFile)
If OpenControl = True Then 'Açıksa
    Workbooks(FileNameUpdate).Close SaveChanges:=True
ElseIf OpenControl = False Then
    '
End If
Workbooks.Open (SourceUpdateFile)


Set WsRapor1 = ThisWorkbook.Worksheets(3)
Set WsRapor = ThisWorkbook.Worksheets(4)
Set WsRapor3 = ThisWorkbook.Worksheets(5)
Set WsTutanak = ThisWorkbook.Worksheets(6)

Set WsSkpUpdate = Workbooks(FileNameUpdate).Worksheets(2)
Set WsRapor1Update = Workbooks(FileNameUpdate).Worksheets(3)
Set WsRaporUpdate = Workbooks(FileNameUpdate).Worksheets(4)
Set WsRapor3Update = Workbooks(FileNameUpdate).Worksheets(5)
Set WsTutanakUpdate = Workbooks(FileNameUpdate).Worksheets(6)

WsRapor1.Unprotect Password:="123"
WsRapor.Unprotect Password:="123"
WsRapor3.Unprotect Password:="123"
WsTutanak.Unprotect Password:="123"

WsSkpUpdate.Unprotect Password:="123"
WsRapor1Update.Unprotect Password:="123"
WsRaporUpdate.Unprotect Password:="123"
WsRapor3Update.Unprotect Password:="123"
WsTutanakUpdate.Unprotect Password:="123"

'SKP içindeki Birim Adını güncelle
WsSKP.Cells(6, 99).Value = WsSkpUpdate.Cells(6, 99).Value 'Aslında bu bilgi, tanımlar dosyasından yeni skp'ye zaten aktarılmış olması gerekir ! Ancak aşağıda Footer Field bölümünde bu bilgi kullanılacak. Dolayısıyla kalacak.

SayUpdate = WsRapor1Update.Range("CF100000").End(xlUp).Row
If SayUpdate > 6 Then
    'Rapor1 sayfası güncellemesi
    WsRapor1Update.Range("E7:DW" & SayUpdate).Copy WsRapor1.Range("E7:DW" & SayUpdate)
    With WsRapor1.Range("E7:E" & SayUpdate)
        .Font.name = "Open Sans"
        .Font.Size = 9
    End With
    
    With WsRapor1.Range("F7:J" & SayUpdate)
        .Font.Size = 16
        .Font.Color = RGB(60, 100, 180)
    End With

    With WsRapor1.Range("K7:DW" & SayUpdate)
        .Font.name = "Open Sans"
        .Font.Size = 9
    End With
    With WsRapor1.Range("F7:J" & SayUpdate)
        .Font.Color = RGB(60, 100, 180)
    End With
    With WsRapor1.Range("L7:L" & SayUpdate)
        .Font.Color = RGB(0, 0, 0)
    End With
    With WsRapor1.Range("CG7:CG" & SayUpdate)
        .NumberFormat = "@"
    End With
End If

SayUpdate = WsRaporUpdate.Range("CN100000").End(xlUp).Row
If SayUpdate > 6 Then
    'Rapor sayfası güncellemesi
    WsRaporUpdate.Range("E7:HR" & SayUpdate).Copy WsRapor.Range("E7:HR" & SayUpdate)
    With WsRapor.Range("E7:E" & SayUpdate)
        .Font.name = "Open Sans"
        .Font.Size = 9
    End With

    With WsRapor.Range("F7:J" & SayUpdate)
        .Font.Size = 16
        .Font.Color = RGB(60, 100, 180)
    End With
    With WsRapor.Range("M7:V" & SayUpdate)
        .Font.Size = 16
        .Font.Color = RGB(60, 100, 180)
    End With
    
    With WsRapor.Range("K7:L" & SayUpdate)
        .Font.name = "Open Sans"
        .Font.Size = 9
    End With
    With WsRapor.Range("W7:HR" & SayUpdate)
        .Font.name = "Open Sans"
        .Font.Size = 9
    End With
    With WsRapor.Range("F7:J" & SayUpdate)
        .Font.Color = RGB(60, 100, 180)
    End With
    With WsRapor.Range("M7:V" & SayUpdate)
        .Font.Color = RGB(60, 100, 180)
    End With
    With WsRapor.Range("L7:L" & SayUpdate)
        .Font.Color = RGB(0, 0, 0)
    End With
    With WsRapor.Range("CO7:CO" & SayUpdate)
        .NumberFormat = "@"
    End With
End If

SayUpdate = WsRapor3Update.Range("FH100000").End(xlUp).Row
If SayUpdate > 6 Then
    'Rapor3 sayfası güncellemesi
    
    If WsRapor3Update.Range("G6").Value = "Rapor" Then
        WsRapor3Update.Range("E7:HT" & SayUpdate).Copy WsRapor3.Range("E7:HT" & SayUpdate)
    Else
        WsRapor3Update.Range("E7:F" & SayUpdate).Copy WsRapor3.Range("E7:F" & SayUpdate) 'Sıra ve tutanak
                                                                                       'Rapor kısmı eski dosyada olmadığı için aktarım yapılmayacak
        WsRapor3Update.Range("G7:J" & SayUpdate).Copy WsRapor3.Range("H7:K" & SayUpdate) 'Tutanak2 ve sonrası
        WsRapor3Update.Range("K7:K" & SayUpdate).Copy WsRapor3.Range("L7:L" & SayUpdate)
        WsRapor3Update.Range("L7:L" & SayUpdate).Copy WsRapor3.Range("N7:N" & SayUpdate)
        WsRapor3Update.Range("O7:HB" & SayUpdate).Copy WsRapor3.Range("O7:HB" & SayUpdate)
        
        With WsRapor3.Range("E7:E" & SayUpdate)
            .Font.name = "Open Sans"
            .Font.Size = 9
        End With
        With WsRapor3.Range("L7:HB" & SayUpdate)
            .Font.name = "Open Sans"
            .Font.Size = 9
        End With
        With WsRapor3.Range("F7:K" & SayUpdate)
            .Font.Color = RGB(60, 100, 180)
        End With
        With WsRapor3.Range("N7:N" & SayUpdate)
            .Font.Color = RGB(0, 0, 0)
        End With
    End If
    WsRapor3.Range("AY7:AY" & Rows.Count).NumberFormat = "@" 'Barkod formatı text olmalı
    WsRapor3.Range("HD7:HJ" & Rows.Count).NumberFormat = "@"
    WsRapor3.Range("G7:G" & Rows.Count).NumberFormat = "@"
    WsRapor3.Range("M7:M" & Rows.Count).NumberFormat = "@"
    With WsRapor3.Range("FI7:FI" & Rows.Count)
        .NumberFormat = "@"
    End With

    With WsRapor3.Range("F7:K" & SayUpdate)
        .Font.Size = 16
        .Font.Color = RGB(60, 100, 180)
    End With
End If

'Maksimum değerler.
SayA = WsTutanakUpdate.Range("A100000").End(xlUp).Row
SayB = WsTutanakUpdate.Range("B100000").End(xlUp).Row
SayC = WsTutanakUpdate.Range("C100000").End(xlUp).Row
SayD = WsTutanakUpdate.Range("D100000").End(xlUp).Row
SayUpdate = WorksheetFunction.Max(SayA, SayB, SayC, SayD)
If SayUpdate > 2 Then
    'Tutanak sayfası güncellemesi
    WsTutanakUpdate.Range("A3:D" & SayUpdate).Copy
    WsTutanak.Range("A3:D" & SayUpdate).PasteSpecial xlPasteValues
End If


'___________________________________________

'Fark girişleri sayfası varsa onları da aktar
FarkSayfaKontrol = 0
For i = 1 To Workbooks(FileNameUpdate).Worksheets.Count
    If Workbooks(FileNameUpdate).Worksheets(i).name = "Temp Discrepancies" Then
        FarkSayfaKontrol = FarkSayfaKontrol + 1
        'MsgBox Workbooks(FileNameUpdate).Worksheets(i).Name
    ElseIf Workbooks(FileNameUpdate).Worksheets(i).name = "Discrepancies – Report 1" Then
        FarkSayfaKontrol = FarkSayfaKontrol + 1
        'MsgBox Workbooks(FileNameUpdate).Worksheets(i).Name
    ElseIf Workbooks(FileNameUpdate).Worksheets(i).name = "Discrepancies – Report 2" Then
        FarkSayfaKontrol = FarkSayfaKontrol + 1
        'MsgBox Workbooks(FileNameUpdate).Worksheets(i).Name
    End If
Next i
If FarkSayfaKontrol = 3 Then
    'MsgBox "Yeni uygulama!"

    ThisWorkbook.Worksheets(7).Unprotect Password:="123"
    ThisWorkbook.Worksheets(8).Unprotect Password:="123"
    ThisWorkbook.Worksheets(9).Unprotect Password:="123"

    Workbooks(FileNameUpdate).Worksheets(7).Unprotect Password:="123"
    Workbooks(FileNameUpdate).Worksheets(8).Unprotect Password:="123"
    Workbooks(FileNameUpdate).Worksheets(9).Unprotect Password:="123"
        
    SayFarkGirisRapor1 = Workbooks(FileNameUpdate).Worksheets(8).Range("Q100000").End(xlUp).Row
    SayFarkGirisRapor = Workbooks(FileNameUpdate).Worksheets(9).Range("Q100000").End(xlUp).Row
    If SayFarkGirisRapor1 > 2 Then
        'ThisWorkbook.Worksheets(8).Range("A3:Q" & SayFarkGirisRapor1) = Workbooks(FileNameUpdate).Worksheets(8).Range("A3:Q" & SayFarkGirisRapor1)
        Workbooks(FileNameUpdate).Worksheets(8).Range("A3:Q" & SayFarkGirisRapor1).Copy
        ThisWorkbook.Worksheets(8).Range("A3:Q" & SayFarkGirisRapor1).PasteSpecial xlPasteValues
    End If
    If SayFarkGirisRapor > 2 Then
        'ThisWorkbook.Worksheets(9).Range("A3:Q" & SayFarkGirisRapor1) = Workbooks(FileNameUpdate).Worksheets(9).Range("A3:Q" & SayFarkGirisRapor1)
        Workbooks(FileNameUpdate).Worksheets(9).Range("A3:Q" & SayFarkGirisRapor).Copy
        ThisWorkbook.Worksheets(9).Range("A3:Q" & SayFarkGirisRapor).PasteSpecial xlPasteValues
    End If

    ThisWorkbook.Worksheets(7).Protect Password:="123"
    ThisWorkbook.Worksheets(8).Protect Password:="123"
    ThisWorkbook.Worksheets(9).Protect Password:="123"

    Workbooks(FileNameUpdate).Worksheets(7).Protect Password:="123"
    Workbooks(FileNameUpdate).Worksheets(8).Protect Password:="123"
    Workbooks(FileNameUpdate).Worksheets(9).Protect Password:="123"
    
Else
    'MsgBox "Eski uygulama!"
End If

'___________________________________________



'___________________________________________

'Rapor no sayfası varsa onları da aktar
FarkSayfaKontrol = 0
For i = 1 To Workbooks(FileNameUpdate).Worksheets.Count
    If Workbooks(FileNameUpdate).Worksheets(i).name = "Report 2 Numbers" Then
        FarkSayfaKontrol = FarkSayfaKontrol + 1
        'MsgBox Workbooks(FileNameUpdate).Worksheets(i).Name
    ElseIf Workbooks(FileNameUpdate).Worksheets(i).name = "Report 1 Numbers" Then
        FarkSayfaKontrol = FarkSayfaKontrol + 1
        'MsgBox Workbooks(FileNameUpdate).Worksheets(i).Name
    End If
Next i
If FarkSayfaKontrol = 2 Then
    'MsgBox "Yeni uygulama!"

    ThisWorkbook.Worksheets(10).Unprotect Password:="123"
    ThisWorkbook.Worksheets(11).Unprotect Password:="123"

    Workbooks(FileNameUpdate).Worksheets(10).Unprotect Password:="123"
    Workbooks(FileNameUpdate).Worksheets(11).Unprotect Password:="123"
        
    SayFarkGirisRapor1 = Workbooks(FileNameUpdate).Worksheets(11).Range("E100000").End(xlUp).Row
    SayFarkGirisRapor = Workbooks(FileNameUpdate).Worksheets(10).Range("E100000").End(xlUp).Row
    If SayFarkGirisRapor1 > 6 Then
        Workbooks(FileNameUpdate).Worksheets(11).Range("A7:K" & SayFarkGirisRapor1).Copy
        ThisWorkbook.Worksheets(11).Range("A7:K" & SayFarkGirisRapor1).PasteSpecial xlPasteValues
    End If
    If SayFarkGirisRapor > 6 Then
        Workbooks(FileNameUpdate).Worksheets(10).Range("A7:K" & SayFarkGirisRapor).Copy
        ThisWorkbook.Worksheets(10).Range("A7:K" & SayFarkGirisRapor).PasteSpecial xlPasteValues
    End If


    ThisWorkbook.Worksheets(10).Protect Password:="123"
    ThisWorkbook.Worksheets(11).Protect Password:="123"

    Workbooks(FileNameUpdate).Worksheets(10).Protect Password:="123"
    Workbooks(FileNameUpdate).Worksheets(11).Protect Password:="123"
    
Else
    'MsgBox "Eski uygulama!"
End If

'___________________________________________



WsRapor1.Protect Password:="123"
WsRapor.Protect Password:="123"
WsRapor3.Protect Password:="123"
WsTutanak.Protect Password:="123"

WsRapor1Update.Protect Password:="123"
WsRaporUpdate.Protect Password:="123"
WsRapor3Update.Protect Password:="123"
WsTutanakUpdate.Protect Password:="123"

'Uygulama update dosyasını kapat
FileNameUpdate = "Update.xlsm"
OpenControl = IsFileOpen(SourceUpdateFile)
If OpenControl = True Then 'Açıksa
    Workbooks(FileNameUpdate).Close SaveChanges:=True
ElseIf OpenControl = False Then
    '
End If

'Güvenlik için update dosyasının adını değiştir
oldName = AutoPath & "\System Files\Update\Update.xlsm"
newName = AutoPath & "\System Files\Update\Updatex.xlsm"
Name oldName As newName

'__________________Eski Uygulama dosyasından yeni uygulama dosyasına data transferi (Bitiş)




GoTo IslemGunlukleriniAtla 'Eski sistemin kayıt defetterlerini aktarma (Gelecekte yeni sistemde güncelleme olursa bu bölümü aç. 29.11.2021, 15:25, İshak.

'__________________________________Rapor2_2 işlem günlüğü (Başlangıç)

'Update içindeki dosya
FileName = "System Registry Report 2.2.xlsx"
OpenControl = IsFileOpen(IslemGunluguUpdate & FileName)
If OpenControl = True Then 'Açıksa
    Workbooks(FileName).Close SaveChanges:=True
ElseIf OpenControl = False Then
    '
End If

'Normal sistem klasörü içindeki dosya
FileName = "System Registry Report 2.2.xlsx"
OpenControl = IsFileOpen(IslemGunlugu & FileName)
If OpenControl = True Then 'Açıksa
    Workbooks(FileName).Close SaveChanges:=True
ElseIf OpenControl = False Then
    '
End If
Workbooks.Open (IslemGunlugu & FileName)


'Update işlem günlüğü içindeki dosya adını değiştir.
oldName = Rapor2_2IslemGunluguUpdate
newName = AutoPath & "\System Files\Update\System Files\System Templates\Registry Reports\System Registry Report 2.2 Update.xlsx"
Name oldName As newName
'MsgBox newName

FileNameUpdate = "System Registry Report 2.2 Update.xlsx"
OpenControl = IsFileOpen(newName)
If OpenControl = True Then 'Açıksa
    Workbooks(FileNameUpdate).Close SaveChanges:=True
ElseIf OpenControl = False Then
    '
End If

Workbooks.Open (newName)

Set WsRapor2_2IslemGunlugu = Workbooks(FileName).Worksheets(1)
Set WsRapor2_2IslemGunluguUpdate = Workbooks(FileNameUpdate).Worksheets(1)

WsRapor2_2IslemGunlugu.Unprotect Password:="123"
WsRapor2_2IslemGunluguUpdate.Unprotect Password:="123"

SayUpdate = WsRapor2_2IslemGunluguUpdate.Range("O100000").End(xlUp).Row
If SayUpdate > 6 Then
    'Rapor2_2 işlem günlüğü dosyası güncellemesi
    WsRapor2_2IslemGunluguUpdate.Range("B7:S" & SayUpdate).Copy WsRapor2_2IslemGunlugu.Range("B7:S" & SayUpdate)
    With WsRapor2_2IslemGunlugu.Range("B7:S" & SayUpdate)
        .Font.name = "Open Sans"
        .Font.Size = 9
    End With
End If

WsRapor2_2IslemGunlugu.Protect Password:="123"
WsRapor2_2IslemGunluguUpdate.Protect Password:="123"

'Rapor2_2 İşlem Günlüğü Update dosyasını kapat
FileName = "System Registry Report 2.2 Update.xlsx"
OpenControl = IsFileOpen(newName)
If OpenControl = True Then 'Açıksa
    Workbooks(FileNameUpdate).Close SaveChanges:=True
ElseIf OpenControl = False Then
    '
End If

'Rapor2_2 İşlem Günlüğü Update isimli dosya adını eski haline getir.
oldName = AutoPath & "\System Files\Update\System Files\System Templates\Registry Reports\System Registry Report 2.2 Update.xlsx"
newName = Rapor2_2IslemGunluguUpdate
Name oldName As newName

'Rapor2_2 İşlem Günlüğü dosyasını kapat
FileName = "System Registry Report 2.2.xlsx"
OpenControl = IsFileOpen(IslemGunlugu & FileName)
If OpenControl = True Then 'Açıksa
    Workbooks(FileName).Close SaveChanges:=True
ElseIf OpenControl = False Then
    '
End If

'__________________________________Rapor2_2 işlem günlüğü (Bitiş)


'__________________________________Rapor1 işlem günlüğü (Başlangıç)

'Update içindeki dosya
FileName = "System Registry Report 1.xlsx"
OpenControl = IsFileOpen(IslemGunluguUpdate & FileName)
If OpenControl = True Then 'Açıksa
    Workbooks(FileName).Close SaveChanges:=True
ElseIf OpenControl = False Then
    '
End If

'Normal sistem klasörü içindeki dosya
FileName = "System Registry Report 1.xlsx"
OpenControl = IsFileOpen(IslemGunlugu & FileName)
If OpenControl = True Then 'Açıksa
    Workbooks(FileName).Close SaveChanges:=True
ElseIf OpenControl = False Then
    '
End If
Workbooks.Open (IslemGunlugu & FileName)


'Update işlem günlüğü içindeki dosya adını değiştir.
oldName = Rapor1IslemGunluguUpdate
newName = AutoPath & "\System Files\Update\System Files\System Templates\Registry Reports\System Registry Report 1 Update.xlsx"
Name oldName As newName
'MsgBox newName

FileNameUpdate = "System Registry Report 1 Update.xlsx"
OpenControl = IsFileOpen(newName)
If OpenControl = True Then 'Açıksa
    Workbooks(FileNameUpdate).Close SaveChanges:=True
ElseIf OpenControl = False Then
    '
End If

Workbooks.Open (newName)

Set WsRapor1IslemGunlugu = Workbooks(FileName).Worksheets(1)
Set WsRapor1IslemGunluguUpdate = Workbooks(FileNameUpdate).Worksheets(1)

WsRapor1IslemGunlugu.Unprotect Password:="123"
WsRapor1IslemGunluguUpdate.Unprotect Password:="123"

SayUpdate = WsRapor1IslemGunluguUpdate.Range("O100000").End(xlUp).Row
If SayUpdate > 6 Then
    'Rapor1 işlem günlüğü dosyası güncellemesi
    WsRapor1IslemGunluguUpdate.Range("B7:S" & SayUpdate).Copy WsRapor1IslemGunlugu.Range("B7:S" & SayUpdate)
    With WsRapor1IslemGunlugu.Range("B7:S" & SayUpdate)
        .Font.name = "Open Sans"
        .Font.Size = 9
    End With
End If

WsRapor1IslemGunlugu.Protect Password:="123"
WsRapor1IslemGunluguUpdate.Protect Password:="123"

'Rapor1 İşlem Günlüğü Update dosyasını kapat
FileName = "System Registry Report 1 Update.xlsx"
OpenControl = IsFileOpen(newName)
If OpenControl = True Then 'Açıksa
    Workbooks(FileNameUpdate).Close SaveChanges:=True
ElseIf OpenControl = False Then
    '
End If

'Rapor1 İşlem Günlüğü Update isimli dosya adını eski haline getir.
oldName = AutoPath & "\System Files\Update\System Files\System Templates\Registry Reports\System Registry Report 1 Update.xlsx"
newName = Rapor1IslemGunluguUpdate
Name oldName As newName

'Rapor1 İşlem Günlüğü dosyasını kapat
FileName = "System Registry Report 1.xlsx"
OpenControl = IsFileOpen(IslemGunlugu & FileName)
If OpenControl = True Then 'Açıksa
    Workbooks(FileName).Close SaveChanges:=True
ElseIf OpenControl = False Then
    '
End If

'__________________________________Rapor1 işlem günlüğü (Bitiş)



'__________________________________Rapor işlem günlüğü (Başlangıç)

'Update içindeki dosya
FileName = "System Registry Report 2.1.xlsx"
OpenControl = IsFileOpen(IslemGunluguUpdate & FileName)
If OpenControl = True Then 'Açıksa
    Workbooks(FileName).Close SaveChanges:=True
ElseIf OpenControl = False Then
    '
End If

'Normal sistem klasörü içindeki dosya
FileName = "System Registry Report 2.1.xlsx"
OpenControl = IsFileOpen(IslemGunlugu & FileName)
If OpenControl = True Then 'Açıksa
    Workbooks(FileName).Close SaveChanges:=True
ElseIf OpenControl = False Then
    '
End If
Workbooks.Open (IslemGunlugu & FileName)


'Update işlem günlüğü içindeki dosya adını değiştir.
oldName = RaporIslemGunluguUpdate
newName = AutoPath & "\System Files\Update\System Files\System Templates\Registry Reports\System Registry Report 2.1 Update.xlsx"
Name oldName As newName
'MsgBox newName

FileNameUpdate = "System Registry Report 2.1 Update.xlsx"
OpenControl = IsFileOpen(newName)
If OpenControl = True Then 'Açıksa
    Workbooks(FileNameUpdate).Close SaveChanges:=True
ElseIf OpenControl = False Then
    '
End If

Workbooks.Open (newName)

Set WsRaporIslemGunlugu = Workbooks(FileName).Worksheets(1)
Set WsRaporIslemGunluguUpdate = Workbooks(FileNameUpdate).Worksheets(1)

WsRaporIslemGunlugu.Unprotect Password:="123"
WsRaporIslemGunluguUpdate.Unprotect Password:="123"

SayUpdate = WsRaporIslemGunluguUpdate.Range("O100000").End(xlUp).Row
If SayUpdate > 6 Then
    'Rapor işlem günlüğü dosyası güncellemesi
    WsRaporIslemGunluguUpdate.Range("B7:T" & SayUpdate).Copy WsRaporIslemGunlugu.Range("B7:T" & SayUpdate)
    With WsRaporIslemGunlugu.Range("B7:T" & SayUpdate)
        .Font.name = "Open Sans"
        .Font.Size = 9
    End With
End If

WsRaporIslemGunlugu.Protect Password:="123"
WsRaporIslemGunluguUpdate.Protect Password:="123"

'Rapor İşlem Günlüğü Update dosyasını kapat
FileName = "System Registry Report 2.1 Update.xlsx"
OpenControl = IsFileOpen(newName)
If OpenControl = True Then 'Açıksa
    Workbooks(FileNameUpdate).Close SaveChanges:=True
ElseIf OpenControl = False Then
    '
End If

'Rapor İşlem Günlüğü Update isimli dosya adını eski haline getir.
oldName = AutoPath & "\System Files\Update\System Files\System Templates\Registry Reports\System Registry Report 2.1 Update.xlsx"
newName = RaporIslemGunluguUpdate
Name oldName As newName

'Rapor İşlem Günlüğü dosyasını kapat
FileName = "System Registry Report 2.1.xlsx"
OpenControl = IsFileOpen(IslemGunlugu & FileName)
If OpenControl = True Then 'Açıksa
    Workbooks(FileName).Close SaveChanges:=True
ElseIf OpenControl = False Then
    '
End If

'__________________________________Rapor işlem günlüğü (Bitiş)


IslemGunlukleriniAtla:




'____________________________________________________________________Rapor System Templatesı (Başlangıç)


'Atla:



Call OpenWordControl


'__________________________________Alt Bilgi (Başlangıç)

'Update içindeki dosya

FileNameUpdate = "Report 2 Cover Letter.docm" 'Eski uygulamadan Footer Field alınırken bu üst yazı kullanılacak.
newName = RaporUstYaziSablonuUpdate & FileNameUpdate

'Update dosyasını aç
On Error Resume Next
Set objWord = GetObject(, "Word.Application")
Set objWord = GetObject(, "Word.Application")
Set objWord = GetObject(, "Word.Application")
Set objWord = GetObject(, "Word.Application")
Set objWord = GetObject(, "Word.Application")
If objWord Is Nothing Then
    Set objWord = CreateObject("Word.Application")
    objWord.Visible = False
End If
objWord.Documents.Open FileName:=newName
'objWord.Visible = True
'objWord.Activate 'Ekrana getirir.
Set objDoc = GetObject(newName)

''Taşıyıcıyı resetle"
'TasiyiciIlk = ""
'For i = 1 To 5
'    TasiyiciArray(i) = ""
'Next i



Set WsSKP = ThisWorkbook.Worksheets(2)

'Açık userformlar varsa kapatılsın.
Call ModuleInit.UserFormlariKapat
Load support_update_UI
With support_update_UI
    .Show vbModeless
End With


'Birim adı
support_update_UI.ComboSube.Value = WsSKP.Cells(6, 99).Value & " Unit"

'Kurum_ANo getir
'YeniVersiyon = False
Kurum_ANoStr = objDoc.Tables(1).Cell(Row:=3, Column:=1).Range.Text
If InStr(Kurum_ANoStr, "KURUM_A") > 0 Then 'Kasım 2019 sürümü
    support_update_UI.Kurum_ANo.Value = Mid(Kurum_ANoStr, 6, InStr(Kurum_ANoStr, "-") - 6)
    support_update_UI.Aktar.Value = objDoc.Sections(1).Footers(wdHeaderFooterPrimary).Range.Tables(1).Cell(Row:=1, Column:=2).Range.Text
Else
    Kurum_ANoStr = objDoc.Tables(1).Cell(Row:=4, Column:=1).Range.Text
    support_update_UI.Kurum_ANo.Value = Mid(Kurum_ANoStr, 8, 8)
    support_update_UI.Aktar.Value = objDoc.Sections(1).Footers(wdHeaderFooterPrimary).Range.Tables(1).Cell(Row:=1, Column:=1).Range.Text
'    YeniVersiyon = True
End If


objDoc.Close SaveChanges:=False

objWord.Visible = False
'Nesneleri temizle
Set fso = Nothing
Set objWord = Nothing
Set objDoc = Nothing

support_update_UI.Aktar.SetFocus
For i = 0 To support_update_UI.Aktar.LineCount - 1
    If i = 0 Then
        support_update_UI.Adres.Value = Split(support_update_UI.Aktar.Value, Chr(13))(i)
    ElseIf i = 1 Then
        support_update_UI.Tel.Value = Split(support_update_UI.Aktar.Value, Chr(13))(i)
    ElseIf i = 2 Then
        support_update_UI.Eposta.Value = Split(support_update_UI.Aktar.Value, Chr(13))(i)
    ElseIf i = 3 Then
        support_update_UI.ElektronikAg.Value = Split(support_update_UI.Aktar.Value, Chr(13))(i)
    End If
Next i


WsSKP.Cells(6, 99).Value = Left(support_update_UI.ComboSube.Value, Len(support_update_UI.ComboSube.Value) - 7)

'Birimx = UCase(Replace(Replace(Worksheets(2).Cells(6, 99).Value, "i", "I"), "ı", "I")) & " UNIT"
Birimx = WsSKP.Cells(6, 99).Value & " Unit"
Kurum_A = "ORGANIZATION A" & vbNewLine & Birimx

'Birden fazla boşluk varsa kaldır
'Sağdaki ve soldaki tek boşluğu kaldır
For i = 1 To 50
    support_update_UI.Adres.Value = Replace(support_update_UI.Adres.Value, "  ", " ")
Next i
Do While Left(support_update_UI.Adres.Value, 1) = " "
    support_update_UI.Adres.Value = Right(support_update_UI.Adres.Value, Len(support_update_UI.Adres.Value) - 1)
Loop
Do While Right(support_update_UI.Adres.Value, 1) = " "
    support_update_UI.Adres.Value = Left(support_update_UI.Adres.Value, Len(support_update_UI.Adres.Value) - 1)
Loop
For i = 1 To 50
    support_update_UI.Tel.Value = Replace(support_update_UI.Tel.Value, "  ", " ")
Next i
Do While Left(support_update_UI.Tel.Value, 1) = " "
    support_update_UI.Tel.Value = Right(support_update_UI.Tel.Value, Len(support_update_UI.Tel.Value) - 1)
Loop
Do While Right(support_update_UI.Tel.Value, 1) = " "
    support_update_UI.Tel.Value = Left(support_update_UI.Tel.Value, Len(support_update_UI.Tel.Value) - 1)
Loop
For i = 1 To 50
    support_update_UI.Eposta.Value = Replace(support_update_UI.Eposta.Value, "  ", " ")
Next i
Do While Left(support_update_UI.Eposta.Value, 1) = " "
    support_update_UI.Eposta.Value = Right(support_update_UI.Eposta.Value, Len(support_update_UI.Eposta.Value) - 1)
Loop
Do While Right(support_update_UI.Eposta.Value, 1) = " "
    support_update_UI.Eposta.Value = Left(support_update_UI.Eposta.Value, Len(support_update_UI.Eposta.Value) - 1)
Loop
For i = 1 To 50
    support_update_UI.ElektronikAg.Value = Replace(support_update_UI.ElektronikAg.Value, "  ", " ")
Next i
Do While Left(support_update_UI.ElektronikAg.Value, 1) = " "
    support_update_UI.ElektronikAg.Value = Right(support_update_UI.ElektronikAg.Value, Len(support_update_UI.ElektronikAg.Value) - 1)
Loop
Do While Right(support_update_UI.ElektronikAg.Value, 1) = " "
    support_update_UI.ElektronikAg.Value = Left(support_update_UI.ElektronikAg.Value, Len(support_update_UI.ElektronikAg.Value) - 1)
Loop

'Paragraf işaretlerini kaldır.
If InStr(support_update_UI.Adres.Value, Chr(13)) > 0 Then
    support_update_UI.Adres.Value = Right(support_update_UI.Adres.Value, Len(support_update_UI.Adres.Value) - 2)
End If
If InStr(support_update_UI.Tel.Value, Chr(13)) > 0 Then
    support_update_UI.Tel.Value = Right(support_update_UI.Tel.Value, Len(support_update_UI.Tel.Value) - 2)
End If
If InStr(support_update_UI.Eposta.Value, Chr(13)) > 0 Then
    support_update_UI.Eposta.Value = Right(support_update_UI.Eposta.Value, Len(support_update_UI.Eposta.Value) - 2)
End If
If InStr(support_update_UI.ElektronikAg.Value, Chr(13)) > 0 Then
    support_update_UI.ElektronikAg.Value = Right(support_update_UI.ElektronikAg.Value, Len(support_update_UI.ElektronikAg.Value) - 2)
End If


'HEDEF DOSYA OLAN Footer Field'e taşıyıcılardaki bilgileri aktar

FileName = "External Footer.docm"

'Footer Field
objWord.Documents.Open FileName:=AltBilgi & FileName
Set objDoc = GetObject(AltBilgi & FileName)
'objDoc.ActiveWindow.Visible = False
objDoc.Sections(1).Footers(wdHeaderFooterFirstPage).Range.Tables(1).Cell(Row:=1, Column:=1).Range.Text = support_update_UI.Adres.Value & vbNewLine & _
                                                                                   support_update_UI.Tel.Value & vbNewLine & _
                                                                                   support_update_UI.Eposta.Value & vbNewLine & _
                                                                                   support_update_UI.ElektronikAg.Value
objDoc.Sections(1).Footers(wdHeaderFooterPrimary).Range.Tables(1).Cell(Row:=1, Column:=1).Range.Text = support_update_UI.Adres.Value & vbNewLine & _
                                                                                   support_update_UI.Tel.Value & vbNewLine & _
                                                                                   support_update_UI.Eposta.Value & vbNewLine & _
                                                                                   support_update_UI.ElektronikAg.Value
'Kurum_ANo
Kurum_ANoStr = objDoc.Tables(1).Cell(Row:=4, Column:=1).Range.Text
Kurum_ANoStr = Left(Kurum_ANoStr, 7) & support_update_UI.Kurum_ANo.Value & Mid(Kurum_ANoStr, 16, Len(Kurum_ANoStr) - 17)
'MsgBox Kurum_ANoStr

Unload support_update_UI


objDoc.Tables(1).Cell(Row:=4, Column:=1).Range.Text = Kurum_ANoStr
'KURUM_A
objDoc.Tables(2).Cell(Row:=4, Column:=1).Range.Text = Kurum_A
objDoc.Close SaveChanges:=True
objWord.Visible = False



'__________________________________Alt Bilgi (Bitiş)


'UPDATE KLASÖRÜ
'Yeni versiyonda Standard Valid ve Standard Invalid dışında var olan System Templatesı sil
On Error Resume Next
'ContSay = 0
KontrolFile = Dir(Rapor2Sablonu & "*.???")
Do While KontrolFile <> ""
    'ContSay = ContSay + 1
    If KontrolFile = "Standard Invalid.docm" Then
        '
    ElseIf KontrolFile = "Standard Valid.docm" Then
        '
    Else
        Kill Rapor2Sablonu & KontrolFile
    End If
    KontrolFile = Dir()
Loop
'If ContSay > 0 Then
'    Kill Rapor2Sablonu & "*.???"
'End If
    
a = 1
KontrolFile = Dir(RaporSablonuUpdate & "*.docm")
Do While KontrolFile <> ""
    
    FileName = KontrolFile
    FileName = Replace(FileName, ".docm", "")
    FileName = Replace(FileName, "Ve", "ve")
    FileNameUpdate = FileName & " Update"

    'Standard Valid ve Standard Invalid dosyalarında aktarım yapma!
    If FileName = "Standard Valid" Or FileName = "Standard Invalid" Then
        'MsgBox FileName
        GoTo NormaliAtla
    End If
    
    'Update Rapor Şablonu içindeki dosya adını değiştir.
    oldName = RaporSablonuUpdate & FileName & ".docm"
    newName = RaporSablonuUpdate & FileNameUpdate & ".docm"
    Name oldName As newName
    
    
    'Update dosyasını aç
    On Error Resume Next
    Set objWord = GetObject(, "Word.Application")
    Set objWord = GetObject(, "Word.Application")
    Set objWord = GetObject(, "Word.Application")
    Set objWord = GetObject(, "Word.Application")
    Set objWord = GetObject(, "Word.Application")
    If objWord Is Nothing Then
        Set objWord = CreateObject("Word.Application")
        objWord.Visible = False
    End If
    objWord.Documents.Open FileName:=newName
    'objWord.Visible = True
    'objWord.Activate 'Ekrana getirir.
    Set objDoc = GetObject(newName)
    
    'Taşıyıcıyı resetle
    TasiyiciIlk = ""
    For i = 1 To 6
        TasiyiciArray(i) = ""
    Next i
    TasiyiciSon = ""

    'Taşıyıcı İlk
    If objDoc.Tables(3).Rows.Count > 0 Then
        TasiyiciIlk = objDoc.Tables(3).Cell(Row:=2, Column:=1).Range.Text
        'Boşlukları ve boş satırları kaldır
        TasiyiciIlk = Replace(Replace(TasiyiciIlk, Chr(10), ""), Chr(13), "")
        TextLine = TasiyiciIlk
        Do While Left(TextLine, 1) = " " ' Delete any excess spaces
            TextLine = Right(TextLine, Len(TextLine) - 1)
        Loop
        Do While Right(TextLine, 1) = " " ' Delete any excess spaces
            TextLine = Left(TextLine, Len(TextLine) - 1)
        Loop
        TasiyiciIlk = TextLine
    End If

    'Taşıyıcı 1-2-3-4-5-6
    If objDoc.Tables(4).Rows.Count > 0 Then
        For j = 1 To objDoc.Tables(4).Rows.Count
            TasiyiciArray(j) = objDoc.Tables(4).Cell(Row:=j, Column:=2).Range.Text
            'Boşlukları ve boş satırları kaldır
            TasiyiciArray(j) = Replace(Replace(TasiyiciArray(j), Chr(10), ""), Chr(13), "")
            TextLine = TasiyiciArray(j)
            Do While Left(TextLine, 1) = " " ' Delete any excess spaces
                TextLine = Right(TextLine, Len(TextLine) - 1)
            Loop
            Do While Right(TextLine, 1) = " " ' Delete any excess spaces
                TextLine = Left(TextLine, Len(TextLine) - 1)
            Loop
            TasiyiciArray(j) = TextLine
        Next j
    End If

    'Taşıyıcı Son
    If objDoc.Tables(5).Rows.Count > 0 Then
        TasiyiciSon = objDoc.Tables(5).Cell(Row:=1, Column:=1).Range.Text
        'Boşlukları ve boş satırları kaldır
        TasiyiciSon = Replace(Replace(TasiyiciSon, Chr(10), ""), Chr(13), "")
        TextLine = TasiyiciSon
        Do While Left(TextLine, 1) = " " ' Delete any excess spaces
            TextLine = Right(TextLine, Len(TextLine) - 1)
        Loop
        Do While Right(TextLine, 1) = " " ' Delete any excess spaces
            TextLine = Left(TextLine, Len(TextLine) - 1)
        Loop
        TasiyiciSon = TextLine
    End If
    
    'Kopyalanacak dosya invalid mi valid mi?
    gecersizKontrol = False
    gecerliKontrol = False
    Set MyRange = objDoc.Tables(3).Range
    MyRange.Find.Execute FindText:="invalid"
    If InStr(MyRange.Text, "invalid") <> 0 Or InStr(MyRange.Text, "invalid") <> 0 Then
        gecersizKontrol = True
        GoTo geçersizgecerliKontrolOk
    End If
    Set MyRange = objDoc.Tables(4).Range
    MyRange.Find.Execute FindText:="invalid"
    If InStr(MyRange.Text, "invalid") <> 0 Or InStr(MyRange.Text, "invalid") <> 0 Then
        gecersizKontrol = True
        GoTo geçersizgecerliKontrolOk
    End If
     
    Set MyRange = objDoc.Tables(3).Range
    MyRange.Find.Execute FindText:="valid"
    If InStr(MyRange.Text, "valid") <> 0 Or InStr(MyRange.Text, "valid") <> 0 Then
        gecerliKontrol = True
        GoTo geçersizgecerliKontrolOk
    End If
    Set MyRange = objDoc.Tables(4).Range
    MyRange.Find.Execute FindText:="valid"
    If InStr(MyRange.Text, "valid") <> 0 Or InStr(MyRange.Text, "valid") <> 0 Then
        gecerliKontrol = True
        GoTo geçersizgecerliKontrolOk
    End If
geçersizgecerliKontrolOk:

    'Update dosyasını kapat
    objDoc.Close SaveChanges:=False
    'objWord.Visible = False

    'Update Rapor Şablonu içindeki dosya adını eski haline getir.
    oldName = RaporSablonuUpdate & FileNameUpdate & ".docm"
    newName = RaporSablonuUpdate & FileName & ".docm"
    Name oldName As newName
    
    
    'Yeni dosyaya data transferini başlat
    'Şablon dosyayı Rapor Şablonu klasörüne kopyala ve adını değiştir.
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If gecersizKontrol = True Then
        fso.CopyFile (Normalgecersiz), Rapor2Sablonu & FileName & ".docm", True
    ElseIf gecerliKontrol = True Then
        fso.CopyFile (Normalgecerli), Rapor2Sablonu & FileName & ".docm", True
    Else 'other
        fso.CopyFile (Normalgecersiz), Rapor2Sablonu & FileName & ".docm", True
    End If

    HedefFile = Rapor2Sablonu & FileName & ".docm"

    'Hedef dosyayı aç
    On Error Resume Next
    Set objWord = GetObject(, "Word.Application")
    Set objWord = GetObject(, "Word.Application")
    Set objWord = GetObject(, "Word.Application")
    Set objWord = GetObject(, "Word.Application")
    Set objWord = GetObject(, "Word.Application")
    If objWord Is Nothing Then
        Set objWord = CreateObject("Word.Application")
        objWord.Visible = False
    End If
    objWord.Documents.Open FileName:=HedefFile
    'objWord.Visible = True
    'objWord.Activate 'Ekrana getirir.
    Set objDoc = GetObject(HedefFile)

    'İlk bölümü doldur
    If TasiyiciIlk <> "" Then
        objDoc.Tables(3).Cell(Row:=2, Column:=1).Range.Text = TasiyiciIlk
    End If

    'Önce satırları sil ve teke düşür
    If objDoc.Tables(4).Rows.Count > 1 Then
        For j = 1 To objDoc.Tables(4).Rows.Count - 1
            objDoc.Tables(4).Rows(2).Delete
        Next j
    End If
    'Satır bold olmasın
    objDoc.Tables(4).Cell(Row:=1, Column:=2).Range.Font.Bold = False
    'Sonra tabloya satır ekle

    
    'HEDEF KLASÖR
    Cont = 0
    For j = 1 To 6
        If TasiyiciArray(j) <> "" Then
            Cont = Cont + 1
            If Cont > 1 Then
                With objDoc.Tables(4)
                    .Rows.Add BeforeRow:=objDoc.Tables(4).Rows(1)
                    .Rows(1).Height = 10
                End With
            End If
        End If
    Next j
    'Arayı doldur
    If Cont > 0 Then
        For j = 1 To Cont
            objDoc.Tables(4).Cell(Row:=j, Column:=2).Range.Text = TasiyiciArray(j)
        Next j
    End If
    'Son bölümü doldur
    If TasiyiciSon <> "" Then
        objDoc.Tables(5).Cell(Row:=1, Column:=1).Range.Text = TasiyiciSon
    End If

    'ANAHTAR KELİMELERİ BOLD YAP
    'invalid (ilk satır)
    Set MyRange = objDoc.Tables(3).Cell(Row:=2, Column:=1).Range
    MyRange.Find.Execute FindText:="invalid"
    If InStr(MyRange.Text, "invalid") <> 0 Or InStr(MyRange.Text, "invalid") <> 0 Then
        MyRange.Font.Bold = True
    End If
    'valid (ilk satır)
    Set MyRange = objDoc.Tables(3).Cell(Row:=2, Column:=1).Range
    MyRange.Find.Execute FindText:="valid"
    If InStr(MyRange.Text, "valid") <> 0 Or InStr(MyRange.Text, "valid") <> 0 Then
        MyRange.Font.Bold = True
    End If
    'invalid (1. satır)
    Set MyRange = objDoc.Tables(4).Cell(Row:=1, Column:=2).Range
    MyRange.Find.Execute FindText:="invalid"
    If InStr(MyRange.Text, "invalid") <> 0 Or InStr(MyRange.Text, "invalid") <> 0 Then
        MyRange.Font.Bold = True
    End If
    'valid (1. satır)
    Set MyRange = objDoc.Tables(4).Cell(Row:=1, Column:=2).Range
    MyRange.Find.Execute FindText:="valid"
    If InStr(MyRange.Text, "valid") <> 0 Or InStr(MyRange.Text, "valid") <> 0 Then
        MyRange.Font.Bold = True
    End If

    'Kullanılmayan satırların içeriğini boşalt.
    If TasiyiciArray(1) = "" Then
        objDoc.Tables(4).Cell(Row:=1, Column:=1).Range.Text = ""
        objDoc.Tables(4).Cell(Row:=1, Column:=1).Range.ListFormat.RemoveNumbers
        objDoc.Tables(4).Cell(Row:=1, Column:=2).Range.Text = ""
    End If
    If TasiyiciSon = "" Then
        objDoc.Tables(5).Cell(Row:=1, Column:=1).Range.Text = ""
    End If

    'Öğeyi kaydet
    objDoc.Close SaveChanges:=True

NormaliAtla:

'    If a >= 1 Then
'        GoTo WhileSon
'    End If
'    a = a + 1
    
    KontrolFile = Dir()
Loop



'____________________________________________________________________Rapor System Templatesı (Bitiş)



'__________________________________Item Notes (Başlangıç)

KontrolFile = Dir(NotlarUpdate & "*.txt")
Do While KontrolFile <> ""

    FileName = KontrolFile
    FileName = Replace(FileName, ".txt", "")
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.CopyFile (NotlarUpdate & KontrolFile), StrNotlar & FileName & ".txt", True

    KontrolFile = Dir()
Loop

'__________________________________Item Notes (Bitiş)



Call OpenWordControl

ThisWorkbook.Worksheets(2).Protect Password:="123"
ThisWorkbook.Protect "123"

Application.ScreenUpdating = True
Application.DisplayAlerts = True
Application.EnableEvents = True

'ThisWorkbook.Save

MsgBox "The update process is two-staged, and the first stage of the update has been completed successfully. To proceed to the second stage of the update, please close and reopen Enterprise Document Automation System (this file). Then, select your unit from the Unit Settings Interface and click the Add/Update button. After some time, the appearance of an information message indicating that the process was completed successfully signifies that the second stage of the update has also been successfully executed.", vbOKOnly + vbInformation, "Enterprise Document Automation System"

GoTo UpdateOk


Son:

Call OpenWordControl

ThisWorkbook.Worksheets(2).Protect Password:="123"
ThisWorkbook.Protect "123"


Application.ScreenUpdating = True
Application.DisplayAlerts = True
Application.EnableEvents = True

UpdateOk:


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



