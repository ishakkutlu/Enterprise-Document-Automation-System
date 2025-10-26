Attribute VB_Name = "ModuleRibbon"
Option Explicit
Public OpenWordTakip As Boolean
Public EkranKontrol As Boolean

Sub islemgunlukleriribbon(Control As IRibbonControl)

core_registry_reports_UI.Show vbModeless 'vbModal

End Sub

Sub etiketyaziciribbon(Control As IRibbonControl)

core_label_envelope_printing_UI.Show vbModal 'vbModal 'vbModeless
'MsgBox "The delivery printing interface is temporarily out of service.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"

End Sub

Sub rapor1ribbon(Control As IRibbonControl)

Application.DisplayAlerts = False
Application.EnableEvents = False
Application.ScreenUpdating = False

'Açık userformlar varsa kapatılsın.
Call ModuleInit.UserFormlariKapat

Call Rapor1

Application.DisplayAlerts = True
Application.EnableEvents = True
Application.ScreenUpdating = True

End Sub

Sub rapor1girisribbon(Control As IRibbonControl)

'Açık userformlar varsa kapatılsın.
Call ModuleInit.UserFormlariKapat

If ActiveSheet.Index <> 3 Then

Application.DisplayAlerts = False
Application.EnableEvents = False
Application.ScreenUpdating = False

Call Rapor1

Application.DisplayAlerts = True
Application.EnableEvents = True
Application.ScreenUpdating = True

End If

Call Rapor1Girisler
    
End Sub

Sub rapor1islemgunlukleriribbon(Control As IRibbonControl)

Application.DisplayAlerts = False
Application.ScreenUpdating = False

'Call Rapor1
Call IslemGunluguRapor1

Application.DisplayAlerts = True
Application.ScreenUpdating = True

End Sub
Sub rapor1bilgiribbon(Control As IRibbonControl)

Application.DisplayAlerts = False
Application.EnableEvents = False
Application.ScreenUpdating = False

Call Rapor1
Call YardimRapor1

Application.DisplayAlerts = True
Application.EnableEvents = True
Application.ScreenUpdating = True

End Sub

Sub Rapor1()
Dim ws As Object, i As Integer

ThisWorkbook.Unprotect "123"

i = 0
Worksheets(3).Visible = True
For Each ws In ThisWorkbook.Worksheets
    i = i + 1
    If i <> 3 Then
        ws.Visible = False
    End If
Next ws
Worksheets(3).Activate

With ActiveWindow
    .DisplayHorizontalScrollBar = True
    .DisplayVerticalScrollBar = True
End With

ActiveSheet.DisplayPageBreaks = False

ThisWorkbook.Protect "123"

End Sub

Sub Rapor1Girisler()
Dim ws As Object, i As Integer
Dim yukseklik As Variant
Dim Rep As Variant, genislik As Variant


EkranKontrol = False

Call ModuleSystemSettings.NumLockAc

'Ekrana göre formun ayarlanması
If Application.Height < 766 Then 'core_report1_entry_UI.Height Then

    EkranKontrol = True
    
    With core_report1_entry_UI
        .TasiyiciFrame.Left = 12 '36
        .OgeTurleriFrameUst.Left = 24 '48
        .ScrollFrame.Left = 24 '48
        .Rapor1Frame.Left = 24 '48
        .Tutanak2Frame.Left = 24 '48
        .UstYaziFrame.Left = 24 '48
        .AltMenuFrame.Left = 24 '48
        .Width = 1024 + 12 '1072 '970 '950
        .Height = 475 '766 'Aslında 485, ancak scrollbarda 10 px sapması var.
        .StartUpPosition = 0
        .Left = Application.Left + (0.5 * Application.Width) - (0.5 * (1024 + 12))
        .Top = Application.Top '+ (0.5 * Application.Height) - (0.5 * 485)
        .Show vbModal 'vbModal 'vbModeless
    End With
    
    core_report1_entry_UI.ScrollBars = fmScrollBarsVertical
    core_report1_entry_UI.ScrollHeight = core_report1_entry_UI.Height + 285 - 30 '- (RepYukseklik - Application.Height) '572 '549
    core_report1_entry_UI.ScrollTop = 10 '0 iken tutanak1 butonuna ilk basmada tepki vermiyor.
 
Else
    With core_report1_entry_UI
        .Width = 1072 '970 '950
        .Height = 766 '766
        .StartUpPosition = 0
        .Left = Application.Left + (0.5 * Application.Width) - (0.5 * 1072)
        .Top = Application.Top + (0.5 * Application.Height) - (0.5 * 766)
        .Show vbModal 'vbModal 'vbModeless
    End With
End If

End Sub


Sub IslemGunluguRapor1()
Dim AutoPath As String, IslemGunlukleriKlasor As String, IslemGunlugu As String, OpenControl As String

'Pathfinder...
AutoPath = ThisWorkbook.Path
IslemGunlukleriKlasor = AutoPath & "\System Files\System Templates\Registry Reports\"
IslemGunlugu = IslemGunlukleriKlasor & "System Registry Report 1.xlsx"

'Registry Reports klasör adını kontrol et.
If Not Dir(IslemGunlukleriKlasor, vbDirectory) <> vbNullString Then
    MsgBox "Cannot access the directory " & IslemGunlukleriKlasor & ". The folder named 'Registry Reports' may have been renamed or deleted.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Out
End If

If Not Dir(IslemGunlugu, vbDirectory) <> vbNullString Then
    MsgBox "Cannot access the directory " & IslemGunlugu & ". The names of folders and/or files in this directory may have been changed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Out
End If


'İşlem günlüğü açıksa kaydet ve kapat.
OpenControl = IsWorkBookOpen(IslemGunlugu)
If OpenControl = True Then
    Workbooks("System Registry Report 1.xlsx").Save
    Workbooks("System Registry Report 1.xlsx").Close SaveChanges:=False
End If

Workbooks.Open (IslemGunlugu)

'Scrollbar göster
With ActiveWindow
    .DisplayHorizontalScrollBar = True
    .DisplayVerticalScrollBar = True
End With

Out:

End Sub

Sub YardimRapor1()
Dim a() As Variant, i As Variant, Tanimlar As String, DestTanimlar As String
Dim AutoPath As String, DestRapor3lar As String, OpenControl As String
Dim Rapor3_2File As String, SayHedef As Long, ItemName As String, j As Integer

Dim fso As Object, objWord As Object, objDoc As Object, Birimx As String, Kurum_A As String, Kurum_ANoStr As String
Dim Rapor3_1File As String, Rapor3_1TipBFile As String, FinansalBirimFile As String, RaporFile As String, Rapor1File As String
Dim DestRapor2 As String, DestRapor1 As String

Dim DestOpUserFolderName As String, DestOpUserFolder As String, DestOperasyon As String
Dim OpenKontrolName As String, ContSay As Integer, KontrolFile As String
Dim ReNameTaslak As String, SourceTaslak As String


'Pathfinder...
AutoPath = ThisWorkbook.Path
DestOperasyon = AutoPath & "\System Files\Operation\"
'Draft File
SourceTaslak = AutoPath & "\System Files\Help Documents\Report 1 Workflow – Help.docm"
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
If Not Dir(SourceTaslak, vbDirectory) <> vbNullString Then
    MsgBox "Cannot access the directory " & SourceTaslak & ". The names of folders and/or files in this directory may have been changed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

'Operation klasörü içinde kullanıcı modülü klasörü yoksa oluştur.
If Not Dir(DestOpUserFolder, vbDirectory) <> vbNullString Then
    MkDir DestOpUserFolder
End If

'On Error Resume Next
ReNameTaslak = "Help Documents"
'________________________________________

'Close the all Word application
Call OpenWordControl

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
fso.CopyFile (SourceTaslak), DestOpUserFolder & ReNameTaslak & ".docm", True

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
objWord.Documents.Open FileName:=DestOpUserFolder & ReNameTaslak & ".docm"
objWord.Visible = True
objWord.Activate 'Ekrana getirir.
'objDoc.Activate 'Ekrana getirmez.
objWord.Application.WindowState = wdWindowStateMaximize

Son:

Set objWord = Nothing
Set objDoc = Nothing

End Sub



Sub rapor2ribbon(Control As IRibbonControl)

Application.DisplayAlerts = False
Application.EnableEvents = False
Application.ScreenUpdating = False

'Açık userformlar varsa kapatılsın.
Call ModuleInit.UserFormlariKapat

Call Rapor

Application.DisplayAlerts = True
Application.EnableEvents = True
Application.ScreenUpdating = True

End Sub

Sub rapor2girisribbon(Control As IRibbonControl)

'Açık userformlar varsa kapatılsın.
Call ModuleInit.UserFormlariKapat

If ActiveSheet.Index <> 4 Then

Application.DisplayAlerts = False
Application.EnableEvents = False
Application.ScreenUpdating = False

Call Rapor

Application.DisplayAlerts = True
Application.EnableEvents = True
Application.ScreenUpdating = True

End If

Call RaporGirisler

End Sub

Sub rapor2_1islemgunlukleriribbon(Control As IRibbonControl)

Application.DisplayAlerts = False
Application.ScreenUpdating = False

'Call Rapor
Call IslemGunluguRapor2_1

Application.DisplayAlerts = True
Application.ScreenUpdating = True

End Sub

Sub rapor2_2islemgunlukleriribbon(Control As IRibbonControl)

Application.DisplayAlerts = False
Application.ScreenUpdating = False

'Call Rapor
Call IslemGunluguRapor2_2

Application.DisplayAlerts = True
Application.ScreenUpdating = True

End Sub

Sub rapor2bilgiribbon(Control As IRibbonControl)

Application.DisplayAlerts = False
Application.EnableEvents = False
Application.ScreenUpdating = False

Call Rapor
Call YardimRapor

Application.DisplayAlerts = True
Application.EnableEvents = True
Application.ScreenUpdating = True

End Sub

Sub Rapor()
Dim ws As Object, i As Integer

ThisWorkbook.Unprotect "123"

i = 0
Worksheets(4).Visible = True
For Each ws In ThisWorkbook.Worksheets
    i = i + 1
    If i <> 4 Then
        ws.Visible = False
    End If
Next ws
Worksheets(4).Activate

With ActiveWindow
    .DisplayHorizontalScrollBar = True
    .DisplayVerticalScrollBar = True
End With

ActiveSheet.DisplayPageBreaks = False

ThisWorkbook.Protect "123"

End Sub

Sub RaporGirisler()
Dim ws As Object, i As Integer
Dim yukseklik As Variant
Dim Rep As Variant, genislik As Variant


EkranKontrol = False

Call ModuleSystemSettings.NumLockAc

'Ekrana göre formun ayarlanması
If Application.Height < 766 Then 'core_report2_entry_UI.Height Then

    EkranKontrol = True
    
    With core_report2_entry_UI
        .TasiyiciFrame.Left = 12 '36
        .OgeTurleriFrameUst.Left = 24 '48
        .ScrollFrame.Left = 24 '48
        .Rapor1Frame.Left = 24 '48

        .Rapor2_2KararFrame.Left = 24
        .XXXMudTutanak2Frame.Left = 24
        .XXXMudUstYaziFrame.Left = 24
        .UstYaziFrame.Left = 24
        .GelenXXXMudTutanak1Frame.Left = 24
        .GelenXXXMudTutanak2Frame.Left = 24
        .Tutanak2Frame.Left = 24
        .Rapor2_2UstYaziFrame.Left = 24

        .AltMenuFrame.Left = 24 '48
        .Width = 1024 + 12 '1072 '970 '950
        .Height = 475 '766 'Aslında 485, ancak scrollbarda 10 px sapması var.
        .StartUpPosition = 0
        .Left = Application.Left + (0.5 * Application.Width) - (0.5 * (1024 + 12))
        .Top = Application.Top '+ (0.5 * Application.Height) - (0.5 * 485)
        .Show vbModal 'vbModal 'vbModeless
    End With
    
    core_report2_entry_UI.ScrollBars = fmScrollBarsVertical
    core_report2_entry_UI.ScrollHeight = core_report2_entry_UI.Height + 285 - 30 '- (RepYukseklik - Application.Height) '572 '549
    core_report2_entry_UI.ScrollTop = 10 '0 iken tutanak1 butonuna ilk basmada tepki vermiyor.
Else
    With core_report2_entry_UI
        .Width = 1072 '970 '950
        .Height = 766 '766
        .StartUpPosition = 0
        .Left = Application.Left + (0.5 * Application.Width) - (0.5 * 1072)
        .Top = Application.Top + (0.5 * Application.Height) - (0.5 * 766)
        .Show vbModal 'vbModal 'vbModeless
    End With
End If

End Sub

Sub IslemGunluguRapor2_1()
Dim AutoPath As String, IslemGunlukleriKlasor As String, IslemGunlugu As String, OpenControl As String

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


'İşlem günlüğü açıksa kaydet ve kapat.
OpenControl = IsWorkBookOpen(IslemGunlugu)
If OpenControl = True Then
    Workbooks("System Registry Report 2.1.xlsx").Save
    Workbooks("System Registry Report 2.1.xlsx").Close SaveChanges:=False
End If

Workbooks.Open (IslemGunlugu)

'Scrollbar göster
With ActiveWindow
    .DisplayHorizontalScrollBar = True
    .DisplayVerticalScrollBar = True
End With

Out:

End Sub

Sub IslemGunluguRapor2_2()
Dim AutoPath As String, IslemGunlukleriKlasor As String, IslemGunlugu As String, OpenControl As String

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


'İşlem günlüğü açıksa kaydet ve kapat.
OpenControl = IsWorkBookOpen(IslemGunlugu)
If OpenControl = True Then
    Workbooks("System Registry Report 2.2.xlsx").Save
    Workbooks("System Registry Report 2.2.xlsx").Close SaveChanges:=False
End If

Workbooks.Open (IslemGunlugu)

'Scrollbar göster
With ActiveWindow
    .DisplayHorizontalScrollBar = True
    .DisplayVerticalScrollBar = True
End With

Out:

End Sub

Sub YardimRapor()
Dim a() As Variant, i As Variant, Tanimlar As String, DestTanimlar As String
Dim AutoPath As String, DestRapor3lar As String, OpenControl As String
Dim Rapor3_2File As String, SayHedef As Long, ItemName As String, j As Integer

Dim fso As Object, objWord As Object, objDoc As Object, Birimx As String, Kurum_A As String, Kurum_ANoStr As String
Dim Rapor3_1File As String, Rapor3_1TipBFile As String, FinansalBirimFile As String, RaporFile As String, Rapor1File As String
Dim DestRapor2 As String, DestRapor1 As String

Dim DestOpUserFolderName As String, DestOpUserFolder As String, DestOperasyon As String
Dim OpenKontrolName As String, ContSay As Integer, KontrolFile As String
Dim ReNameTaslak As String, SourceTaslak As String

'Pathfinder...
AutoPath = ThisWorkbook.Path
DestOperasyon = AutoPath & "\System Files\Operation\"
'Draft File
SourceTaslak = AutoPath & "\System Files\Help Documents\Report 2 Workflow – Help.docm"
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
If Not Dir(SourceTaslak, vbDirectory) <> vbNullString Then
    MsgBox "Cannot access the directory " & SourceTaslak & ". The names of folders and/or files in this directory may have been changed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

'Operation klasörü içinde kullanıcı modülü klasörü yoksa oluştur.
If Not Dir(DestOpUserFolder, vbDirectory) <> vbNullString Then
    MkDir DestOpUserFolder
End If

'On Error Resume Next
ReNameTaslak = "Help Documents"
'________________________________________

'Close the all Word application
Call OpenWordControl

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
fso.CopyFile (SourceTaslak), DestOpUserFolder & ReNameTaslak & ".docm", True

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
objWord.Documents.Open FileName:=DestOpUserFolder & ReNameTaslak & ".docm"
objWord.Visible = True
objWord.Activate 'Ekrana getirir.
'objDoc.Activate 'Ekrana getirmez.
objWord.Application.WindowState = wdWindowStateMaximize

'Set objDoc = GetObject(DestOpUserFolder & ReNameTaslak & ".docm")

Son:

Set objWord = Nothing
Set objDoc = Nothing

End Sub

Sub rapor3ribbon(Control As IRibbonControl)

Application.DisplayAlerts = False
Application.EnableEvents = False
Application.ScreenUpdating = False

'Açık userformlar varsa kapatılsın.
Call ModuleInit.UserFormlariKapat

Call Rapor3

Application.DisplayAlerts = True
Application.EnableEvents = True
Application.ScreenUpdating = True

End Sub

Sub rapor3_1girisribbon(Control As IRibbonControl)

'Açık userformlar varsa kapatılsın.
Call ModuleInit.UserFormlariKapat

If ActiveSheet.Index <> 5 Then

Application.DisplayAlerts = False
Application.EnableEvents = False
Application.ScreenUpdating = False

Call Rapor3

Application.DisplayAlerts = True
Application.EnableEvents = True
Application.ScreenUpdating = True

End If

Call Rapor3_1Girisleri

End Sub

Sub rapor3_2girisribbon(Control As IRibbonControl)

'Açık userformlar varsa kapatılsın.
Call ModuleInit.UserFormlariKapat

If ActiveSheet.Index <> 5 Then

Application.DisplayAlerts = False
Application.EnableEvents = False
Application.ScreenUpdating = False

Call Rapor3

Application.DisplayAlerts = True
Application.EnableEvents = True
Application.ScreenUpdating = True

End If

Call Rapor3_2Girisleri

End Sub

Sub Rapor3kayitribbon(Control As IRibbonControl)

Application.DisplayAlerts = False
Application.ScreenUpdating = False

'Call Rapor3
Call IslemGunluguRapor3

Application.DisplayAlerts = True
Application.ScreenUpdating = True

End Sub
Sub rapor3bilgiribbon(Control As IRibbonControl)

Application.DisplayAlerts = False
Application.EnableEvents = False
Application.ScreenUpdating = False

Call Rapor3
Call YardimRapor3

Application.DisplayAlerts = True
Application.EnableEvents = True
Application.ScreenUpdating = True

End Sub

Sub Rapor3()
Dim ws As Object, i As Integer

ThisWorkbook.Unprotect "123"

i = 0
Worksheets(5).Visible = True
For Each ws In ThisWorkbook.Worksheets
    i = i + 1
    If i <> 5 Then
        ws.Visible = False
    End If
Next ws
Worksheets(5).Activate

With ActiveWindow
    .DisplayHorizontalScrollBar = True
    .DisplayVerticalScrollBar = True
End With

ActiveSheet.DisplayPageBreaks = False

ThisWorkbook.Protect "123"

End Sub

Sub Rapor3_1Girisleri()
Dim ws As Object, i As Integer
Dim yukseklik As Variant
Dim Rep As Variant, genislik As Variant

EkranKontrol = False

Call ModuleSystemSettings.NumLockAc

'Ekrana göre formun ayarlanması
If Application.Height < 766 Then 'core_report3_1_entry_UI.Height Then

    EkranKontrol = True

    With core_report3_1_entry_UI
        .TasiyiciFrame.Left = 12 '36
        .OgeTurleriFrameUst.Left = 24 '48
        .ScrollFrame.Left = 24 '48
        .Tutanak2Frame.Left = 24 '48
        .UstYaziFrame.Left = 24 '48
        .AltMenuFrame.Left = 24 '48
        .Width = 1024 + 12 '1072 '970 '950
        .Height = 462 '475 'Aslında 485, ancak scrollbarda 10 px sapması var.
        .StartUpPosition = 0
        .Left = Application.Left + (0.5 * Application.Width) - (0.5 * (1024 + 12))
        .Top = Application.Top '+ (0.5 * Application.Height) - (0.5 * 485)
        .Show vbModal 'vbModal 'vbModeless
    End With
    
    core_report3_1_entry_UI.ScrollBars = fmScrollBarsVertical
    core_report3_1_entry_UI.ScrollHeight = core_report3_1_entry_UI.Height + 246 '- (RepYukseklik - Application.Height) '572 '549
    core_report3_1_entry_UI.ScrollTop = 10 '0 iken tutanak1 butonuna ilk basmada tepki vermiyor.
 
Else
    With core_report3_1_entry_UI
        .Width = 1072 '970 '950
        .Height = 790 '740 '766
        .StartUpPosition = 0
        .Left = Application.Left + (0.5 * Application.Width) - (0.5 * 1072)
        .Top = Application.Top + (0.5 * Application.Height) - (0.5 * 790)
        .Show vbModal 'vbModal 'vbModeless
    End With
End If

End Sub

Sub Rapor3_2Girisleri()
Dim ws As Object, i As Integer
Dim yukseklik As Variant
Dim Rep As Variant, genislik As Variant

EkranKontrol = False

Call ModuleSystemSettings.NumLockAc

'Ekrana göre formun ayarlanması
If Application.Height < 766 Then 'core_report3_2_entry_UI.Height Then

    EkranKontrol = True
    
    With core_report3_2_entry_UI
        .TasiyiciFrame.Left = 12 '36
        .OgeTurleriFrameUst.Left = 24 '48
        .ScrollFrame.Left = 24 '48
        .Tutanak2Frame.Left = 24 '48
        .FinansalBirimUstYaziFrame.Left = 24 '48
        .UstYaziFrame.Left = 24 '48
        .AltMenuFrame.Left = 24 '48
        .Width = 1024 + 12 '1072 '970 '950
        .Height = 518 '442 '766 'Aslında 485, ancak scrollbarda 10 px sapması var.
        .StartUpPosition = 0
        .Left = Application.Left + (0.5 * Application.Width) - (0.5 * (1024 + 12))
        .Top = Application.Top '+ (0.5 * Application.Height) - (0.5 * 485)
        .Show vbModal 'vbModal 'vbModeless
    End With
    
    core_report3_2_entry_UI.ScrollBars = fmScrollBarsVertical
    core_report3_2_entry_UI.ScrollHeight = core_report3_2_entry_UI.Height + 200 - 20 '- (RepYukseklik - Application.Height) '572 '549
    core_report3_2_entry_UI.ScrollTop = 10 '0 iken tutanak1 butonuna ilk basmada tepki vermiyor.
 
Else
    With core_report3_2_entry_UI
        .Width = 1072 '970 '950
        .Height = 790 '730 '766
        .StartUpPosition = 0
        .Left = Application.Left + (0.5 * Application.Width) - (0.5 * 1072)
        .Top = Application.Top + (0.5 * Application.Height) - (0.5 * 790)
        .Show vbModal 'vbModal 'vbModeless
    End With
End If

End Sub

Sub IslemGunluguRapor3()
Dim AutoPath As String, IslemGunlukleriKlasor As String, IslemGunlugu As String, OpenControl As String

'Pathfinder...
AutoPath = ThisWorkbook.Path
IslemGunlukleriKlasor = AutoPath & "\System Files\System Templates\Registry Reports\"
IslemGunlugu = IslemGunlukleriKlasor & "Rapor3 Sistem İşlem Günlüğü.xlsx"

'Check the "Registry Reports" folder name.
If Not Dir(IslemGunlukleriKlasor, vbDirectory) <> vbNullString Then
    MsgBox "Cannot access the directory " & IslemGunlukleriKlasor & ". The folder named 'Registry Reports' may have been renamed or deleted.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Out
End If

If Not Dir(IslemGunlugu, vbDirectory) <> vbNullString Then
    MsgBox "Cannot access the directory " & IslemGunlugu & ". The names of folders and/or files in this directory may have been changed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Out
End If

'İşlem günlüğü açıksa kaydet ve kapat.
OpenControl = IsWorkBookOpen(IslemGunlugu)
If OpenControl = True Then
    Workbooks("Rapor3 Sistem İşlem Günlüğü.xlsx").Save
    Workbooks("Rapor3 Sistem İşlem Günlüğü.xlsx").Close SaveChanges:=False
End If

Workbooks.Open (IslemGunlugu)

'Scrollbar göster
With ActiveWindow
    .DisplayHorizontalScrollBar = True
    .DisplayVerticalScrollBar = True
End With

Out:

End Sub

Sub YardimRapor3()
Dim a() As Variant, i As Variant, Tanimlar As String, DestTanimlar As String
Dim AutoPath As String, DestRapor3lar As String, OpenControl As String
Dim Rapor3_2File As String, SayHedef As Long, ItemName As String, j As Integer

Dim fso As Object, objWord As Object, objDoc As Object, Birimx As String, Kurum_A As String, Kurum_ANoStr As String
Dim Rapor3_1File As String, Rapor3_1TipBFile As String, FinansalBirimFile As String, RaporFile As String, Rapor1File As String
Dim DestRapor2 As String, DestRapor1 As String

Dim DestOpUserFolderName As String, DestOpUserFolder As String, DestOperasyon As String
Dim OpenKontrolName As String, ContSay As Integer, KontrolFile As String
Dim ReNameTaslak As String, SourceTaslak As String

'Pathfinder...
AutoPath = ThisWorkbook.Path
DestOperasyon = AutoPath & "\System Files\Operation\"
'Draft File
SourceTaslak = AutoPath & "\System Files\Help Documents\Report 3 Workflow – Help.docm"
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
If Not Dir(SourceTaslak, vbDirectory) <> vbNullString Then
    MsgBox "Cannot access the directory " & SourceTaslak & ". The names of folders and/or files in this directory may have been changed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

'Operation klasörü içinde kullanıcı modülü klasörü yoksa oluştur.
If Not Dir(DestOpUserFolder, vbDirectory) <> vbNullString Then
    MkDir DestOpUserFolder
End If

'On Error Resume Next
ReNameTaslak = "Help Documents"
'________________________________________

'Close the all Word application
Call OpenWordControl

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
fso.CopyFile (SourceTaslak), DestOpUserFolder & ReNameTaslak & ".docm", True

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
objWord.Documents.Open FileName:=DestOpUserFolder & ReNameTaslak & ".docm"
objWord.Visible = True
objWord.Activate 'Ekrana getirir.
'objDoc.Activate 'Ekrana getirmez.
objWord.Application.WindowState = wdWindowStateMaximize

Son:

Set objWord = Nothing
Set objDoc = Nothing

End Sub

Sub kaynakhareketleriribbon(Control As IRibbonControl)

Application.DisplayAlerts = False
Application.EnableEvents = False
Application.ScreenUpdating = False

'Açık userformlar varsa kapatılsın.
Call ModuleInit.UserFormlariKapat

Call AnaSayfa
Call VarlikHareketleri

Application.DisplayAlerts = True
Application.EnableEvents = True
Application.ScreenUpdating = True

End Sub

Sub gelenbelgeribbon(Control As IRibbonControl)

Application.DisplayAlerts = False
Application.EnableEvents = False
Application.ScreenUpdating = False

'Açık userformlar varsa kapatılsın.
Call ModuleInit.UserFormlariKapat

Call AnaSayfa
Call GelenBelgeTutanak

Application.DisplayAlerts = True
Application.EnableEvents = True
Application.ScreenUpdating = True

End Sub

Sub teslimatlarribbon(Control As IRibbonControl)

Application.DisplayAlerts = False
Application.EnableEvents = False
Application.ScreenUpdating = False

'Açık userformlar varsa kapatılsın.
Call ModuleInit.UserFormlariKapat

Call AnaSayfa
Call TeslimTutanaklari

Application.DisplayAlerts = True
Application.EnableEvents = True
Application.ScreenUpdating = True

End Sub

Sub performansribbon(Control As IRibbonControl)

Application.DisplayAlerts = False
Application.EnableEvents = False
Application.ScreenUpdating = False

'Açık userformlar varsa kapatılsın.
Call ModuleInit.UserFormlariKapat

Call AnaSayfa
Call gecersizPerformans

Application.DisplayAlerts = True
Application.EnableEvents = True
Application.ScreenUpdating = True

End Sub

Sub birimayarlariribbon(Control As IRibbonControl)

Application.DisplayAlerts = False
Application.EnableEvents = False
Application.ScreenUpdating = False

'Açık userformlar varsa kapatılsın.
Call ModuleInit.UserFormlariKapat

Call AnaSayfa
Call Ayarlar

Application.DisplayAlerts = True
Application.EnableEvents = True
Application.ScreenUpdating = True

End Sub
Sub parafribbon(Control As IRibbonControl)

Application.DisplayAlerts = False
Application.EnableEvents = False
Application.ScreenUpdating = False

'Açık userformlar varsa kapatılsın.
Call ModuleInit.UserFormlariKapat

Call AnaSayfa
Call Paraf

Application.DisplayAlerts = True
Application.EnableEvents = True
Application.ScreenUpdating = True

End Sub

Sub resetribbon(Control As IRibbonControl)

Application.DisplayAlerts = False
Application.EnableEvents = False
Application.ScreenUpdating = False

'Açık userformlar varsa kapatılsın.
Call ModuleInit.UserFormlariKapat

Call AnaSayfa

Application.DisplayAlerts = True
Application.EnableEvents = True
Application.ScreenUpdating = True

Call Reset

End Sub

Sub yazismarehberiribbon(Control As IRibbonControl)

Application.DisplayAlerts = False
Application.EnableEvents = False
Application.ScreenUpdating = False

Call AnaSayfa
Call YazismaRehberi

Application.DisplayAlerts = True
Application.EnableEvents = True
Application.ScreenUpdating = True

End Sub

Sub masausturibbon(Control As IRibbonControl)

Application.DisplayAlerts = False
Application.EnableEvents = False
Application.ScreenUpdating = False

Call AnaSayfa
Call Masaustu

Application.DisplayAlerts = True
Application.EnableEvents = True
Application.ScreenUpdating = True

End Sub

Sub updateribbon(Control As IRibbonControl)

'Açık userformlar varsa kapatılsın.
Call ModuleInit.UserFormlariKapat

Call ModuleUpdate.Update


End Sub

Sub bilgiribbon(Control As IRibbonControl)

Application.DisplayAlerts = False
Application.EnableEvents = False
Application.ScreenUpdating = False

Call AnaSayfa
Call Yardim

Application.DisplayAlerts = True
Application.EnableEvents = True
Application.ScreenUpdating = True

End Sub

Sub homeribbon(Control As IRibbonControl)

Application.DisplayAlerts = False
Application.EnableEvents = False
Application.ScreenUpdating = False

'Açık userformlar varsa kapatılsın.
Call ModuleInit.UserFormlariKapat

Call AnaSayfa

Application.DisplayAlerts = True
Application.EnableEvents = True
Application.ScreenUpdating = True

End Sub

Sub AnaSayfa()
Dim ws As Object, i As Integer

ThisWorkbook.Unprotect "123"
ThisWorkbook.Worksheets(1).Unprotect "123"

i = 0
Worksheets(1).Visible = True
For Each ws In ThisWorkbook.Worksheets
    i = i + 1
    If i <> 1 Then
        ws.Visible = False
    End If
Next ws
Worksheets(1).Activate

With ActiveWindow
    .DisplayHorizontalScrollBar = False
    .DisplayVerticalScrollBar = False
End With

ActiveSheet.DisplayPageBreaks = False

ThisWorkbook.Worksheets(1).Protect "123"
ThisWorkbook.Protect "123"

End Sub

Sub GelenBelgeTutanak()

Call ModuleSystemSettings.NumLockAc

If Application.Height < 766 Then
    With core_acceptance_manager_UI
        .StartUpPosition = 0
        .Left = Application.Left + (0.5 * Application.Width) - (0.5 * 828)
        .Top = Application.Top + (0.5 * Application.Height) - (0.5 * 586)
        .Show vbModeless 'vbModal 'vbModeless
    End With
Else
    core_acceptance_manager_UI.Show vbModeless 'vbModeless 'vbModal 'Show vbModeless
End If

End Sub

Sub TeslimTutanaklari()

Call ModuleSystemSettings.NumLockAc

If Application.Height < 766 Then
    With core_delivery_manager_UI
        .StartUpPosition = 0
        .Left = Application.Left + (0.5 * Application.Width) - (0.5 * 828)
        .Top = Application.Top + (0.5 * Application.Height) - (0.5 * 586)
        .Show vbModeless 'vbModal 'vbModeless
    End With
Else
    core_delivery_manager_UI.Show vbModeless 'vbModeless 'vbModal 'Show vbModeless
End If

End Sub

Sub gecersizPerformans()

Call ModuleSystemSettings.NumLockAc

If Application.Height < 766 Then
    With core_performance_report_UI
        .StartUpPosition = 0
        .Left = Application.Left + (0.5 * Application.Width) - (0.5 * 708)
        .Top = Application.Top + (0.5 * Application.Height) - (0.5 * 574)
        .Show vbModeless 'vbModal 'vbModeless
    End With
Else
    core_performance_report_UI.Show vbModeless
End If


End Sub

Sub Ayarlar()

Call ModuleSystemSettings.NumLockAc

If Application.Height < 766 Then
    With core_unit_settings_UI
        .StartUpPosition = 0
        .Left = Application.Left + (0.5 * Application.Width) - (0.5 * 708)
        .Top = Application.Top + (0.5 * Application.Height) - (0.5 * 574)
        .Show vbModeless 'vbModal 'vbModeless
    End With
Else
    core_unit_settings_UI.Show vbModeless
End If

End Sub

Sub Paraf()

Call ModuleSystemSettings.NumLockAc

'Load core_initials_UI
core_initials_UI.Show vbModeless

End Sub

Sub Reset()
Dim AutoPath As String, IslemGunlukleriKlasor As String, OpenControl As String
Dim IslemGunluguOR As String, IslemGunluguR As String, IslemGunluguRapor3 As String, IslemGunluguB As String
Dim ResetlemeFolder As String, Rapor1Resetleme As String, RaporResetleme As String, Rapor3Resetleme As String

'Pathfinder...
AutoPath = ThisWorkbook.Path
IslemGunlukleriKlasor = AutoPath & "\System Files\System Templates\Registry Reports\"
IslemGunluguOR = IslemGunlukleriKlasor & "System Registry Report 1.xlsx"
IslemGunluguR = IslemGunlukleriKlasor & "System Registry Report 2.1.xlsx"
IslemGunluguB = IslemGunlukleriKlasor & "System Registry Report 2.2.xlsx"
'IslemGunluguRapor3 = IslemGunlukleriKlasor & "Rapor3 Sistem İşlem Günlüğü.xlsx"

ResetlemeFolder = AutoPath & "\System Files\System Reset\"
Rapor1Resetleme = ResetlemeFolder & "Reset Report 1.xlsx"
RaporResetleme = ResetlemeFolder & "Reset Report 2.xlsx"
Rapor3Resetleme = ResetlemeFolder & "Reset Report 3.xlsx"

Call ModuleSystemSettings.NumLockAc

'Check the "Registry Reports" folder name.
If Not Dir(IslemGunlukleriKlasor, vbDirectory) <> vbNullString Then
    MsgBox "Cannot access the directory " & IslemGunlukleriKlasor & ". The folder named 'Registry Reports' may have been renamed or deleted.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

If Not Dir(ResetlemeFolder, vbDirectory) <> vbNullString Then
    MsgBox "Cannot access the directory " & ResetlemeFolder & ". The folder named 'System Reset' may have been renamed or deleted.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

If Not Dir(IslemGunluguOR, vbDirectory) <> vbNullString Then
    MsgBox "Cannot access the directory " & IslemGunluguOR & ". The names of folders and/or files in this directory may have been changed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

If Not Dir(IslemGunluguR, vbDirectory) <> vbNullString Then
    MsgBox "Cannot access the directory " & IslemGunluguR & ". The names of folders and/or files in this directory may have been changed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

If Not Dir(IslemGunluguB, vbDirectory) <> vbNullString Then
    MsgBox "Cannot access the directory " & IslemGunluguB & ". The names of folders and/or files in this directory may have been changed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

If Not Dir(Rapor1Resetleme, vbDirectory) <> vbNullString Then
    MsgBox "Cannot access the directory " & Rapor1Resetleme & ". The names of folders and/or files in this directory may have been changed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

If Not Dir(RaporResetleme, vbDirectory) <> vbNullString Then
    MsgBox "Cannot access the directory " & RaporResetleme & ". The names of folders and/or files in this directory may have been changed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

If Not Dir(Rapor3Resetleme, vbDirectory) <> vbNullString Then
    MsgBox "Cannot access the directory " & Rapor3Resetleme & ". The names of folders and/or files in this directory may have been changed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

core_system_reset_wizard1_UI.Show 'vbModeless

Son:

End Sub

Sub ResetProsedur()
Dim Sifre1 As Variant, Sifre2 As Variant
Dim AutoPath As String, IslemGunlukleriKlasor As String, OpenControl As String
Dim IslemGunluguOR As String, IslemGunluguR As String, IslemGunluguRapor3 As String, IslemGunluguB As String
Dim WsIslemGunluguOR As Object, WsIslemGunluguR As Object, WsIslemGunluguB, WsIslemGunluguRapor3 As Object, WsIslemGunluguRapor2_2 As Object
Dim ResetlemeFolder As String, Rapor1Resetleme As String, RaporResetleme As String, Rapor3Resetleme As String
Dim Rapor1File As Object, RaporFile As Object, Rapor3File As Object
Dim SayRapor1 As Long, SayRapor As Long, SayRapor3 As Long, i As Long, j As Long
Dim HedefRapor1 As Long, HedefRapor As Long, HedefRapor3 As Long
Dim Rapor1No As Long, RaporNo As Long, Rapor3No As Long

Dim WsFarkGiris As Worksheet, SayFarkGiris As Integer ', SayA As Long, SayD As Long, SayG As Long, SayJ As Long, SayM As Long
Dim IlkSira As Long, SonSira As Long, WsFarkGirisRapor1 As Worksheet
Dim SayFarkGirisRapor1 As Long, Maxi1 As Integer
Dim IlkSiraBul As Range, SonSiraBul As Range, Cont As Long, k As Long

Application.DisplayAlerts = False
Application.EnableEvents = False
Application.ScreenUpdating = False

'Pathfinder...
AutoPath = ThisWorkbook.Path
IslemGunlukleriKlasor = AutoPath & "\System Files\System Templates\Registry Reports\"
IslemGunluguOR = IslemGunlukleriKlasor & "System Registry Report 1.xlsx"
IslemGunluguR = IslemGunlukleriKlasor & "System Registry Report 2.1.xlsx"
IslemGunluguB = IslemGunlukleriKlasor & "System Registry Report 2.2.xlsx"
'IslemGunluguRapor3 = IslemGunlukleriKlasor & "Rapor3 Sistem İşlem Günlüğü.xlsx"

ResetlemeFolder = AutoPath & "\System Files\System Reset\"
Rapor1Resetleme = ResetlemeFolder & "Reset Report 1.xlsx"
RaporResetleme = ResetlemeFolder & "Reset Report 2.xlsx"
Rapor3Resetleme = ResetlemeFolder & "Reset Report 3.xlsx"


'Kayıt defterleri açıksa kaydet ve kapat.
OpenControl = IsWorkBookOpen(IslemGunluguOR)
If OpenControl = True Then
    Workbooks("System Registry Report 1.xlsx").Close SaveChanges:=True
End If
OpenControl = IsWorkBookOpen(IslemGunluguR)
If OpenControl = True Then
    Workbooks("System Registry Report 2.1.xlsx").Close SaveChanges:=True
End If
OpenControl = IsWorkBookOpen(IslemGunluguB)
If OpenControl = True Then
    Workbooks("System Registry Report 2.2.xlsx").Close SaveChanges:=True
End If

'System Reset dosyaları açıksa kapat
OpenControl = IsWorkBookOpen(Rapor1Resetleme)
If OpenControl = True Then
    Workbooks("Reset Report 1.xlsx").Save
    Workbooks("Reset Report 1.xlsx").Close SaveChanges:=True
End If
OpenControl = IsWorkBookOpen(RaporResetleme)
If OpenControl = True Then
    Workbooks("Reset Report 2.xlsx").Save
    Workbooks("Reset Report 2.xlsx").Close SaveChanges:=True
End If

OpenControl = IsWorkBookOpen(Rapor3Resetleme)
If OpenControl = True Then
    Workbooks("Reset Report 3.xlsx").Save
    Workbooks("Reset Report 3.xlsx").Close SaveChanges:=True
End If


'System Reset işlemlerine başla_____________________________________________

ThisWorkbook.Unprotect "123"
ThisWorkbook.Worksheets(3).Unprotect Password:="123"
ThisWorkbook.Worksheets(4).Unprotect Password:="123"
ThisWorkbook.Worksheets(5).Unprotect Password:="123"
ThisWorkbook.Worksheets(6).Unprotect Password:="123"
ThisWorkbook.Worksheets(7).Unprotect Password:="123"
ThisWorkbook.Worksheets(8).Unprotect Password:="123"
ThisWorkbook.Worksheets(9).Unprotect Password:="123"
ThisWorkbook.Worksheets(10).Unprotect Password:="123"
ThisWorkbook.Worksheets(11).Unprotect Password:="123"
ThisWorkbook.Worksheets(3).Visible = True
ThisWorkbook.Worksheets(4).Visible = True
ThisWorkbook.Worksheets(5).Visible = True
ThisWorkbook.Worksheets(6).Visible = True
ThisWorkbook.Worksheets(7).Visible = True
ThisWorkbook.Worksheets(8).Visible = True
ThisWorkbook.Worksheets(9).Visible = True
ThisWorkbook.Worksheets(10).Visible = True
ThisWorkbook.Worksheets(11).Visible = True

'FARK GİRİŞLERİ HAZIRLIK
Set WsFarkGiris = ThisWorkbook.Worksheets(7)
WsFarkGiris.Rows("3:100000").EntireRow.Delete


'RAPOR1 RESETLEME
Set WsFarkGirisRapor1 = ThisWorkbook.Worksheets(8) 'fark girişleri
Workbooks.Open (Rapor1Resetleme)
Set Rapor1File = Workbooks("Reset Report 1.xlsx").Worksheets(1)
Rapor1File.Unprotect Password:="123"
'System Reset dosyasını temizle.
Rapor1File.Rows("7:100000").EntireRow.Delete
'Modülde son satırı bul.
SayRapor1 = ThisWorkbook.Worksheets(3).Range("CF100000").End(xlUp).Row
If SayRapor1 < 7 Then
    SayRapor1 = 7
    GoTo Rapor1ResetlemeyiAtla
End If
'Modülden varlık çıkışı olmayanları ayıkla ve resetleme dosyasına aktar.
HedefRapor1 = 7
Rapor1No = 1
For i = 7 To SayRapor1
    If ThisWorkbook.Worksheets(3).Range("CE" & i).Value <> "" Then 'Başlangıç no (i)
        For j = i To SayRapor1
            If ThisWorkbook.Worksheets(3).Range("CF" & j).Value <> "" Then 'Bitiş No (j)
                GoTo Rapor1BitisNoBulundu
            End If
        Next j
Rapor1BitisNoBulundu:
        If ThisWorkbook.Worksheets(3).Range("CR" & i).Value = "" Then 'Çıkış tarihi yoksa
            ThisWorkbook.Worksheets(3).Range("E" & i & ":DW" & j).Copy Rapor1File.Range("E" & HedefRapor1 & ":DW" & HedefRapor1 + (j - i))

            '_______________________________
            
            'Fark girişlerini kontrol et.
            Set IlkSiraBul = WsFarkGirisRapor1.Range("P3:P100000").Find(What:=ThisWorkbook.Worksheets(3).Range("E" & i).Value, SearchDirection:=xlNext, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            Set SonSiraBul = WsFarkGirisRapor1.Range("Q3:Q100000").Find(What:=ThisWorkbook.Worksheets(3).Range("E" & i).Value, SearchDirection:=xlNext, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not IlkSiraBul Is Nothing Then
                IlkSira = IlkSiraBul.Row
            Else
                GoTo FarkGirisAtla1
            End If
            If Not SonSiraBul Is Nothing Then
                SonSira = SonSiraBul.Row
            Else
                GoTo FarkGirisAtla1
            End If
            'R kolununu yeni no ile işaretle
            For k = IlkSira To SonSira
                WsFarkGirisRapor1.Range("R" & k).Value = Rapor1No
            Next k
            'başlangıç ve bitişleri S ve T kolonunda takip et; eski no.lar halen kalacak.
            WsFarkGirisRapor1.Range("S" & IlkSira).Value = Rapor1No
            WsFarkGirisRapor1.Range("T" & SonSira).Value = Rapor1No
FarkGirisAtla1:

            '_______________________________
            
            Rapor1File.Range("E" & HedefRapor1).Value = Rapor1No 'Tekrar sıra numarası ver.
            Rapor1File.Range("CE" & HedefRapor1).Value = Rapor1No 'Tekrar Başlangıç no (i) ver.
            Rapor1File.Range("CF" & HedefRapor1 + (j - i)).Value = Rapor1No 'Tekrar Bitiş No (j) ver.

            Rapor1No = Rapor1No + 1
            HedefRapor1 = HedefRapor1 + (j - i) + 1
        End If
    End If
Next i
'Son satir kararları
If HedefRapor1 = 7 Then
    GoTo Rapor1SonSatir7
End If
HedefRapor1 = HedefRapor1 - (j - i) - 1 'System Reset dosyasının son satır numarası
If HedefRapor1 < 7 Then
    HedefRapor1 = 7
End If
Rapor1SonSatir7:
'MsgBox HedefRapor1

'Fark girişleri esas kayıttaki kalacakları, geçici kayıtlar sayfasına aktar
SayFarkGirisRapor1 = WsFarkGirisRapor1.Range("R100000").End(xlUp).Row
Cont = 2
If SayFarkGirisRapor1 > 2 Then
    For k = 3 To SayFarkGirisRapor1
        If WsFarkGirisRapor1.Range("R" & k).Value <> "" Then
            Cont = Cont + 1
            WsFarkGiris.Range("A" & Cont & ":M" & Cont).Value = WsFarkGirisRapor1.Range("A" & k & ":M" & k).Value
            WsFarkGiris.Range("P" & Cont & ":Q" & Cont).Value = WsFarkGirisRapor1.Range("S" & k & ":T" & k).Value
        End If
    Next k
End If

'MODÜLÜ TEMİZLE
ThisWorkbook.Worksheets(3).Rows("7:100000").EntireRow.Delete
WsFarkGirisRapor1.Rows("3:100000").EntireRow.Delete 'fark girişleri
If GlobalResetKapsami = 2 Then 'Varlikdan çıkışı yapılmayanlar hariç
    'Reset dosyasından verileri aktar.
    Rapor1File.Range("E" & 7 & ":DW" & HedefRapor1).Copy ThisWorkbook.Worksheets(3).Range("E" & 7 & ":DW" & HedefRapor1)
    'Geçici kayıtları kalıcı kayıtlara aktar.
    If Cont > 2 Then
        WsFarkGirisRapor1.Range("A" & 3 & ":Q" & Cont).Value = WsFarkGiris.Range("A" & 3 & ":Q" & Cont).Value
    End If
ElseIf GlobalResetKapsami = 1 Then 'Varlikdan çıkışı yapılmayanlar dahil
    '
End If

'Modüldeki verileri rapor1 sayfasına da aktar
Set WsRaporNo = ThisWorkbook.Worksheets(11)
WsRaporNo.Rows("7:100000").EntireRow.Delete 'rapor1 numaraları sayfası
SayGlobal = ThisWorkbook.Worksheets(3).Range("CF100000").End(xlUp).Row
AktarNoGlobal = 7
If SayGlobal < 7 Then
    '
Else
    'Verileri aktar
    WsRaporNo.Range(WsRaporNo.Cells(7, 1), WsRaporNo.Cells(SayGlobal, 1)).Value = Range(Cells(7, 11), Cells(SayGlobal, 11)).Value  'Rapor no
    WsRaporNo.Range(WsRaporNo.Cells(7, 2), WsRaporNo.Cells(SayGlobal, 2)).Value = Range(Cells(7, 60), Cells(SayGlobal, 60)).Value

    For i = 7 To SayGlobal
        If Cells(i, 83).Value <> "" Then
            ilkrow = i
            For j = i To i + 100
                If Cells(j, 84).Value <> "" Then
                    sonrow = j
                    WsRaporNo.Cells(AktarNoGlobal, 3).Value = "Request"
                    WsRaporNo.Cells(AktarNoGlobal, 4).Value = Cells(ilkrow, 85).Value
                    WsRaporNo.Cells(AktarNoGlobal + (sonrow - ilkrow), 5).Value = Cells(ilkrow, 85).Value
                    AktarNoGlobal = AktarNoGlobal + (sonrow - ilkrow) + 1
                    GoTo jDonguOut2
                End If
            Next j
        End If
jDonguOut2:
    Next i
End If

'System Reset dosyasını temizle.
Rapor1File.Rows("7:100000").EntireRow.Delete
WsFarkGiris.Rows("3:100000").EntireRow.Delete


Rapor1ResetlemeyiAtla:
Rapor1File.Protect Password:="123"
Workbooks("Reset Report 1.xlsx").Save
OpenControl = IsWorkBookOpen(Rapor1Resetleme)
If OpenControl = True Then
    Workbooks("Reset Report 1.xlsx").Save
    Workbooks("Reset Report 1.xlsx").Close SaveChanges:=False
End If


'RAPOR RESETLEME
Set WsFarkGirisRapor1 = ThisWorkbook.Worksheets(9) 'fark girişleri
Workbooks.Open (RaporResetleme)
Set RaporFile = Workbooks("Reset Report 2.xlsx").Worksheets(1)
RaporFile.Unprotect Password:="123"
'System Reset dosyasını temizle.
RaporFile.Rows("7:100000").EntireRow.Delete
'Modülde son satırı bul.
SayRapor = ThisWorkbook.Worksheets(4).Range("CN100000").End(xlUp).Row
If SayRapor < 7 Then
    SayRapor = 7
    GoTo RaporResetlemeyiAtla
End If
'Modülden varlık çıkışı olmayanları ayıkla ve resetleme dosyasına aktar.
HedefRapor = 7
RaporNo = 1
For i = 7 To SayRapor
    If ThisWorkbook.Worksheets(4).Range("CM" & i).Value <> "" Then 'Başlangıç no (i)
        For j = i To SayRapor
            If ThisWorkbook.Worksheets(4).Range("CN" & j).Value <> "" Then 'Bitiş No (j)
                GoTo RaporBitisNoBulundu
            End If
        Next j
RaporBitisNoBulundu:
        If ThisWorkbook.Worksheets(4).Range("FR" & i).Value = "Yes" Then 'Rapor2_2
            If ThisWorkbook.Worksheets(4).Range("DB" & i).Value = "" Then 'Çıkış tarihi yoksa
                ThisWorkbook.Worksheets(4).Range("E" & i & ":HR" & j).Copy RaporFile.Range("E" & HedefRapor & ":HR" & HedefRapor + (j - i))

                '_______________________________
                
                'Fark girişlerini kontrol et.
                Set IlkSiraBul = WsFarkGirisRapor1.Range("P3:P100000").Find(What:=ThisWorkbook.Worksheets(4).Range("E" & i).Value, SearchDirection:=xlNext, _
                                SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
                Set SonSiraBul = WsFarkGirisRapor1.Range("Q3:Q100000").Find(What:=ThisWorkbook.Worksheets(4).Range("E" & i).Value, SearchDirection:=xlNext, _
                                SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
                If Not IlkSiraBul Is Nothing Then
                    IlkSira = IlkSiraBul.Row
                Else
                    GoTo FarkGirisAtla2
                End If
                If Not SonSiraBul Is Nothing Then
                    SonSira = SonSiraBul.Row
                Else
                    GoTo FarkGirisAtla2
                End If
                'R kolununu yeni no ile işaretle
                For k = IlkSira To SonSira
                    WsFarkGirisRapor1.Range("R" & k).Value = RaporNo
                Next k
                'başlangıç ve bitişleri S ve T kolonunda takip et; eski no.lar halen kalacak.
                WsFarkGirisRapor1.Range("S" & IlkSira).Value = RaporNo
                WsFarkGirisRapor1.Range("T" & SonSira).Value = RaporNo
FarkGirisAtla2:
    
                '_______________________________
            
                RaporFile.Range("E" & HedefRapor).Value = RaporNo 'Tekrar sıra numarası ver.
                RaporFile.Range("CM" & HedefRapor).Value = RaporNo 'Tekrar Başlangıç no (i) ver.
                RaporFile.Range("CN" & HedefRapor + (j - i)).Value = RaporNo 'Tekrar Bitiş No (j) ver.
                RaporNo = RaporNo + 1
                HedefRapor = HedefRapor + (j - i) + 1
            End If
        ElseIf ThisWorkbook.Worksheets(4).Range("FR" & i).Value = "No" Then 'Rapor
            If ThisWorkbook.Worksheets(4).Range("CZ" & i).Value = "" Then 'Çıkış tarihi yoksa
                ThisWorkbook.Worksheets(4).Range("E" & i & ":HR" & j).Copy RaporFile.Range("E" & HedefRapor & ":HR" & HedefRapor + (j - i))

                '_______________________________
                
                'Fark girişlerini kontrol et.
                Set IlkSiraBul = WsFarkGirisRapor1.Range("P3:P100000").Find(What:=ThisWorkbook.Worksheets(4).Range("E" & i).Value, SearchDirection:=xlNext, _
                                SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
                Set SonSiraBul = WsFarkGirisRapor1.Range("Q3:Q100000").Find(What:=ThisWorkbook.Worksheets(4).Range("E" & i).Value, SearchDirection:=xlNext, _
                                SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
                If Not IlkSiraBul Is Nothing Then
                    IlkSira = IlkSiraBul.Row
                Else
                    GoTo FarkGirisAtla3
                End If
                If Not SonSiraBul Is Nothing Then
                    SonSira = SonSiraBul.Row
                Else
                    GoTo FarkGirisAtla3
                End If
                'R kolununu yeni no ile işaretle
                For k = IlkSira To SonSira
                    WsFarkGirisRapor1.Range("R" & k).Value = RaporNo
                Next k
                'başlangıç ve bitişleri S ve T kolonunda takip et; eski no.lar halen kalacak.
                WsFarkGirisRapor1.Range("S" & IlkSira).Value = RaporNo
                WsFarkGirisRapor1.Range("T" & SonSira).Value = RaporNo
FarkGirisAtla3:
    
                '_______________________________
                
                RaporFile.Range("E" & HedefRapor).Value = RaporNo 'Tekrar sıra numarası ver.
                RaporFile.Range("CM" & HedefRapor).Value = RaporNo 'Tekrar Başlangıç no (i) ver.
                RaporFile.Range("CN" & HedefRapor + (j - i)).Value = RaporNo 'Tekrar Bitiş No (j) ver.
                RaporNo = RaporNo + 1
                HedefRapor = HedefRapor + (j - i) + 1
            End If
        End If
    End If
Next i
'Son satir kararları
If HedefRapor = 7 Then
    GoTo RaporSonSatir7
End If
HedefRapor = HedefRapor - (j - i) - 1 'System Reset dosyasının son satır numarası
If HedefRapor < 7 Then
    HedefRapor = 7
End If
RaporSonSatir7:
'MsgBox HedefRapor

'Fark girişleri esas kayıttaki kalacakları, geçici kayıtlar sayfasına aktar
SayFarkGirisRapor1 = WsFarkGirisRapor1.Range("R100000").End(xlUp).Row
Cont = 2
If SayFarkGirisRapor1 > 2 Then
    For k = 3 To SayFarkGirisRapor1
        If WsFarkGirisRapor1.Range("R" & k).Value <> "" Then
            Cont = Cont + 1
            WsFarkGiris.Range("A" & Cont & ":M" & Cont).Value = WsFarkGirisRapor1.Range("A" & k & ":M" & k).Value
            WsFarkGiris.Range("P" & Cont & ":Q" & Cont).Value = WsFarkGirisRapor1.Range("S" & k & ":T" & k).Value
        End If
    Next k
End If

'MODÜLÜ TEMİZLE
ThisWorkbook.Worksheets(4).Rows("7:100000").EntireRow.Delete
WsFarkGirisRapor1.Rows("3:100000").EntireRow.Delete 'fark girişleri
If GlobalResetKapsami = 2 Then 'Varlikdan çıkışı yapılmayanlar hariç
    'Reset dosyasından verileri aktar.
    RaporFile.Range("E" & 7 & ":HR" & HedefRapor).Copy ThisWorkbook.Worksheets(4).Range("E" & 7 & ":HR" & HedefRapor)
    'Geçici kayıtları kalıcı kayıtlara aktar.
    If Cont > 2 Then
        WsFarkGirisRapor1.Range("A" & 3 & ":Q" & Cont).Value = WsFarkGiris.Range("A" & 3 & ":Q" & Cont).Value
    End If
ElseIf GlobalResetKapsami = 1 Then 'Varlikdan çıkışı yapılmayanlar dahil
    '
End If

'Modüldeki verileri rapor1 sayfasına da aktar
Set WsRaporNo = ThisWorkbook.Worksheets(10)
WsRaporNo.Rows("7:100000").EntireRow.Delete 'rapor1 numaraları sayfası
SayGlobal = ThisWorkbook.Worksheets(4).Range("CN100000").End(xlUp).Row
AktarNoGlobal = 7
If SayGlobal < 7 Then
    '
Else
    'Verileri aktar
    WsRaporNo.Range(WsRaporNo.Cells(7, 1), WsRaporNo.Cells(SayGlobal, 1)).Value = Range(Cells(7, 11), Cells(SayGlobal, 11)).Value  'Rapor no
    WsRaporNo.Range(WsRaporNo.Cells(7, 2), WsRaporNo.Cells(SayGlobal, 2)).Value = Range(Cells(7, 68), Cells(SayGlobal, 68)).Value
    For i = 7 To SayGlobal
        If Cells(i, 91).Value <> "" Then
            ilkrow = i
            For j = i To i + 100
                If Cells(j, 92).Value <> "" Then
                    sonrow = j
                    WsRaporNo.Cells(AktarNoGlobal, 3).Value = "Request"
                    WsRaporNo.Cells(AktarNoGlobal, 4).Value = Cells(ilkrow, 93).Value
                    WsRaporNo.Cells(AktarNoGlobal + (sonrow - ilkrow), 5).Value = Cells(ilkrow, 93).Value
                    AktarNoGlobal = AktarNoGlobal + (sonrow - ilkrow) + 1
                    GoTo jDonguOut1
                End If
            Next j
        End If
jDonguOut1:
    Next i
End If


'System Reset dosyasını temizle.
RaporFile.Rows("7:100000").EntireRow.Delete
WsFarkGiris.Rows("3:100000").EntireRow.Delete


RaporResetlemeyiAtla:
RaporFile.Protect Password:="123"
Workbooks("Reset Report 2.xlsx").Save
OpenControl = IsWorkBookOpen(RaporResetleme)
If OpenControl = True Then
    Workbooks("Reset Report 2.xlsx").Save
    Workbooks("Reset Report 2.xlsx").Close SaveChanges:=True
End If


'RAPOR3 RESETLEME
Workbooks.Open (Rapor3Resetleme)
Set Rapor3File = Workbooks("Reset Report 3.xlsx").Worksheets(1)
Rapor3File.Unprotect Password:="123"
'System Reset dosyasını temizle.
Rapor3File.Rows("7:100000").EntireRow.Delete
'Modülde son satırı bul.
SayRapor3 = ThisWorkbook.Worksheets(5).Range("FH100000").End(xlUp).Row
If SayRapor3 < 7 Then
    SayRapor3 = 7
    GoTo Rapor3ResetlemeyiAtla
End If
'Modülden varlık çıkışı olmayanları ayıkla ve resetleme dosyasına aktar.
HedefRapor3 = 7
Rapor3No = 1
For i = 7 To SayRapor3
    If ThisWorkbook.Worksheets(5).Range("FG" & i).Value <> "" Then 'Başlangıç no (i)
        For j = i To SayRapor3
            If ThisWorkbook.Worksheets(5).Range("FH" & j).Value <> "" Then 'Bitiş No (j)
                GoTo Rapor3BitisNoBulundu
            End If
        Next j
Rapor3BitisNoBulundu:
        If ThisWorkbook.Worksheets(5).Range("FT" & i).Value = "" Then 'Çıkış tarihi yoksa
            ThisWorkbook.Worksheets(5).Range("E" & i & ":HT" & j).Copy Rapor3File.Range("E" & HedefRapor3 & ":HT" & HedefRapor3 + (j - i))
            Rapor3File.Range("E" & HedefRapor3).Value = Rapor3No 'Tekrar sıra numarası ver.
            Rapor3File.Range("FG" & HedefRapor3).Value = Rapor3No 'Tekrar Başlangıç no (i) ver.
            Rapor3File.Range("FH" & HedefRapor3 + (j - i)).Value = Rapor3No 'Tekrar Bitiş No (j) ver.
            'Rapor3File.Range("FI" & HedefRapor3).Value = ""  'Kayıt def. No ver.
            Rapor3No = Rapor3No + 1
            HedefRapor3 = HedefRapor3 + (j - i) + 1
        End If
    End If
Next i
'Son satir kararları
If HedefRapor3 = 7 Then
    GoTo Rapor3SonSatir7
End If
HedefRapor3 = HedefRapor3 - (j - i) - 1 'System Reset dosyasının son satır numarası
If HedefRapor3 < 7 Then
    HedefRapor3 = 7
End If
Rapor3SonSatir7:
'MsgBox HedefRapor3

'MODÜLÜ TEMİZLE
ThisWorkbook.Worksheets(5).Rows("7:100000").EntireRow.Delete
If GlobalResetKapsami = 2 Then 'Varlikdan çıkışı yapılmayanlar hariç
    'Reset dosyasından verileri aktar.
    Rapor3File.Range("E" & 7 & ":HT" & HedefRapor3).Copy ThisWorkbook.Worksheets(5).Range("E" & 7 & ":HT" & HedefRapor3)
ElseIf GlobalResetKapsami = 1 Then 'Varlikdan çıkışı yapılmayanlar dahil
    '
End If



''SADECE TipALAR AKTARILACAK
''Modüldeki verileri rapor1 sayfasına da aktar
Set WsRaporNo = ThisWorkbook.Worksheets(10)
'WsRaporNo.Rows("7:100000").EntireRow.Delete 'rapor1 numaraları sayfası
SayGlobal = ThisWorkbook.Worksheets(5).Range("FH100000").End(xlUp).Row
If SayGlobal < 7 Then
    '
Else

    'raporlardan (varsa) gelen verilerden sonra da Rapor3 verilerini aktar
    AktarNoGlobal = WsRaporNo.Range("E100000").End(xlUp).Row
    If AktarNoGlobal < 7 Then
        AktarNoGlobal = 7
    Else
        AktarNoGlobal = AktarNoGlobal + 1
    End If

    For i = 7 To SayGlobal
        If ThisWorkbook.Worksheets(5).Cells(i, 28).Value = "Type A" Or ThisWorkbook.Worksheets(5).Cells(i, 100).Value = "Type A" Then
            For j = i To i + 100
                If ThisWorkbook.Worksheets(5).Cells(j, 164).Value <> "" Then
                    GoTo jDonguSonBitisNo
                End If
                If j = 100 Then
                    GoTo AktarmaYok
                End If

            Next j
jDonguSonBitisNo:
            'Verileri Aktar
            WsRaporNo.Range(WsRaporNo.Cells(AktarNoGlobal, 1), WsRaporNo.Cells(AktarNoGlobal + (j - i), 1)).Value = Range(Cells(i, 13), Cells(j, 13)).Value 'Rapor no
            WsRaporNo.Cells(AktarNoGlobal, 2).Value = Cells(i, 218).Value
            WsRaporNo.Cells(AktarNoGlobal, 3).Value = "Notification"
            WsRaporNo.Cells(AktarNoGlobal, 4).Value = Cells(i, 165).Value
            WsRaporNo.Cells(AktarNoGlobal + (j - i), 5).Value = Cells(i, 165).Value
            AktarNoGlobal = AktarNoGlobal + (j - i) + 1
        End If
AktarmaYok:
    Next i
End If

'System Reset dosyasını temizle.
Rapor3File.Rows("7:100000").EntireRow.Delete

Rapor3ResetlemeyiAtla:
Rapor3File.Protect Password:="123"
Workbooks("Reset Report 3.xlsx").Save
OpenControl = IsWorkBookOpen(Rapor3Resetleme)
If OpenControl = True Then
    Workbooks("Reset Report 3.xlsx").Save
    Workbooks("Reset Report 3.xlsx").Close SaveChanges:=True
End If

ThisWorkbook.Worksheets(6).Rows("3:100000").EntireRow.Delete

ThisWorkbook.Worksheets(3).Visible = False
ThisWorkbook.Worksheets(4).Visible = False
ThisWorkbook.Worksheets(5).Visible = False
ThisWorkbook.Worksheets(6).Visible = False
ThisWorkbook.Worksheets(7).Visible = False
ThisWorkbook.Worksheets(8).Visible = False
ThisWorkbook.Worksheets(9).Visible = False
ThisWorkbook.Worksheets(10).Visible = False
ThisWorkbook.Worksheets(11).Visible = False

ThisWorkbook.Worksheets(3).Protect Password:="123" ', DrawingObjects:=False
ThisWorkbook.Worksheets(4).Protect Password:="123" ', DrawingObjects:=False
ThisWorkbook.Worksheets(5).Protect Password:="123" ', DrawingObjects:=False
ThisWorkbook.Worksheets(6).Protect Password:="123" ', DrawingObjects:=False
ThisWorkbook.Worksheets(7).Protect Password:="123"
ThisWorkbook.Worksheets(8).Protect Password:="123"
ThisWorkbook.Worksheets(9).Protect Password:="123"
ThisWorkbook.Worksheets(10).Protect Password:="123"
ThisWorkbook.Worksheets(11).Protect Password:="123"
ThisWorkbook.Protect "123"

'Kayıt defterlerinin temizlenmesi.
Workbooks.Open (IslemGunluguOR)
Workbooks.Open (IslemGunluguR)
Workbooks.Open (IslemGunluguB)
'Workbooks.Open (IslemGunluguRapor3)
Set WsIslemGunluguOR = Workbooks("System Registry Report 1.xlsx").Worksheets(1)
Set WsIslemGunluguR = Workbooks("System Registry Report 2.1.xlsx").Worksheets(1)
Set WsIslemGunluguB = Workbooks("System Registry Report 2.2.xlsx").Worksheets(1)
'Set WsIslemGunluguRapor3 = Workbooks("Rapor3 Sistem İşlem Günlüğü.xlsx").Worksheets(1)
WsIslemGunluguOR.Unprotect Password:="123"
WsIslemGunluguR.Unprotect Password:="123"
WsIslemGunluguB.Unprotect Password:="123"
'WsIslemGunluguRapor3.Unprotect Password:="123"

WsIslemGunluguOR.Rows("7:100000").EntireRow.Delete
WsIslemGunluguR.Rows("7:100000").EntireRow.Delete
WsIslemGunluguB.Rows("7:100000").EntireRow.Delete
'WsIslemGunluguRapor3.Rows("7:100000").EntireRow.Delete

WsIslemGunluguOR.Protect Password:="123" ', DrawingObjects:=False
WsIslemGunluguR.Protect Password:="123" ', DrawingObjects:=False
WsIslemGunluguB.Protect Password:="123" ', DrawingObjects:=False
'WsIslemGunluguRapor3.Protect Password:="123" ', DrawingObjects:=False
'Kayıt defterleri açıksa kaydet ve kapat.
OpenControl = IsWorkBookOpen(IslemGunluguOR)
If OpenControl = True Then
    Workbooks("System Registry Report 1.xlsx").Close SaveChanges:=True
End If
OpenControl = IsWorkBookOpen(IslemGunluguR)
If OpenControl = True Then
    Workbooks("System Registry Report 2.1.xlsx").Close SaveChanges:=True
End If
OpenControl = IsWorkBookOpen(IslemGunluguB)
If OpenControl = True Then
    Workbooks("System Registry Report 2.2.xlsx").Close SaveChanges:=True
End If

'System Reset işlemlerini bitir_____________________________________________


If ResetControl = False Then
    'ThisWorkbook.Save
    MsgBox "Reset operation has been successfully completed.", vbOKOnly + vbInformation, "Enterprise Document Automation System"
End If

Son:

Application.DisplayAlerts = True
Application.EnableEvents = True
Application.ScreenUpdating = True

ThisWorkbook.Worksheets(1).Activate

End Sub

Sub YazismaRehberi()
Dim SourceTaslak As String, AutoPath As String
Dim strPDFPath As Variant, sAdobeReader As Variant
Dim SourceEk As String

'Pathfinder...
AutoPath = ThisWorkbook.Path
'Draft File
SourceTaslak = AutoPath & "\System Files\Correspondence Guide\Correspondence Guide.pdf"
SourceEk = AutoPath & "\System Files\Correspondence Guide\Attachment.pdf"

'Check folder names.
If Not Dir(SourceTaslak, vbDirectory) <> vbNullString Then
    MsgBox "Cannot access the directory " & SourceTaslak & ". The names of folders and/or files in this directory may have been changed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

If Not Dir(SourceEk, vbDirectory) <> vbNullString Then
    MsgBox "Cannot access the directory " & SourceEk & ". The names of folders and/or files in this directory may have been changed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If



On Error GoTo Hata
'Dokümanı aç

strPDFPath = SourceEk
ThisWorkbook.FollowHyperlink strPDFPath, NewWindow:=True

strPDFPath = SourceTaslak
ThisWorkbook.FollowHyperlink strPDFPath, NewWindow:=True

GoTo Son

Hata:
MsgBox "An unknown error occurred. Adobe Acrobat Reader, which can open PDF files on your computer, must be installed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
GoTo Son

Son:

End Sub

Sub Masaustu()
Dim oWSH As Object, oShortcut As Object, sPathDeskTop As String
Dim AutoPath As String, NameFinder As String, DestTarget As String
Dim NameLen As Integer

NameLen = Len(ThisWorkbook.name) - 5
NameFinder = Left(ThisWorkbook.name, NameLen)
AutoPath = ThisWorkbook.Path
DestTarget = AutoPath & "\" & NameFinder & ".xlsm"

Set oWSH = CreateObject("WScript.Shell")
sPathDeskTop = oWSH.SpecialFolders("Desktop")
Set oShortcut = oWSH.CreateShortCut(sPathDeskTop & "\" & NameFinder & ".lnk")
With oShortcut
.TargetPath = DestTarget
.IconLocation = AutoPath & "\System Files\Logo" & "\edaslogo.ico, 0"
.Save
End With
Set oWSH = Nothing

MsgBox "The desktop shortcut for " & NameFinder & " has been successfully created. You can quickly access the application using the shortcut created on your desktop.", vbOKOnly + vbInformation, "Enterprise Document Automation System"

Son:

End Sub

Sub Yardim()

Dim a() As Variant, i As Variant, Tanimlar As String, DestTanimlar As String
Dim AutoPath As String, DestRapor3lar As String, OpenControl As String
Dim Rapor3_2File As String, SayHedef As Long, ItemName As String, j As Integer

Dim fso As Object, objWord As Object, objDoc As Object, Birimx As String, Kurum_A As String, Kurum_ANoStr As String
Dim Rapor3_1File As String, Rapor3_1TipBFile As String, FinansalBirimFile As String, RaporFile As String, Rapor1File As String
Dim DestRapor2 As String, DestRapor1 As String

Dim DestOpUserFolderName As String, DestOpUserFolder As String, DestOperasyon As String
Dim OpenKontrolName As String, ContSay As Integer, KontrolFile As String
Dim ReNameTaslak As String, SourceTaslak As String

'Pathfinder...
AutoPath = ThisWorkbook.Path
DestOperasyon = AutoPath & "\System Files\Operation\"
'Draft File
SourceTaslak = AutoPath & "\System Files\Help Documents\Home – Help.docm"
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
If Not Dir(SourceTaslak, vbDirectory) <> vbNullString Then
    MsgBox "Cannot access the directory " & SourceTaslak & ". The names of folders and/or files in this directory may have been changed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

'Operation klasörü içinde kullanıcı modülü klasörü yoksa oluştur.
If Not Dir(DestOpUserFolder, vbDirectory) <> vbNullString Then
    MkDir DestOpUserFolder
End If

'On Error Resume Next
ReNameTaslak = "Help Documents"
'________________________________________

'Close the all Word application
Call OpenWordControl

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
fso.CopyFile (SourceTaslak), DestOpUserFolder & ReNameTaslak & ".docm", True

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
objWord.Documents.Open FileName:=DestOpUserFolder & ReNameTaslak & ".docm"
objWord.Visible = True
objWord.Activate 'Ekrana getirir.
'objDoc.Activate 'Ekrana getirmez.
objWord.Application.WindowState = wdWindowStateMaximize

'Set objDoc = GetObject(DestOpUserFolder & ReNameTaslak & ".docm")

Son:

Set objWord = Nothing
Set objDoc = Nothing

End Sub

Sub VarlikHareketleri()

'core_asset_manager_UI.Show vbModeless 'vbModeless 'vbModal 'Show vbModeless
With core_asset_manager_UI
    .StartUpPosition = 0
    .Left = Application.Left + (0.5 * Application.Width) - (0.5 * 1044)
    .Top = Application.Top + (0.5 * Application.Height) - (0.5 * 586 - 20)
    .Show vbModeless 'vbModal 'vbModeless
End With
    
End Sub

Sub otokapatribbon(Control As IRibbonControl)

Application.DisplayAlerts = False
Application.EnableEvents = False
Application.ScreenUpdating = False

'Açık userformlar varsa kapatılsın.
Call ModuleInit.UserFormlariKapat

Call AnaSayfa
Call otokapat

Application.DisplayAlerts = True
Application.EnableEvents = True
Application.ScreenUpdating = True


End Sub
Sub otokapat()

Call ModuleSystemSettings.NumLockAc

core_auto_close_settings_UI.Show vbModeless

End Sub

Sub ikribbon(Control As IRibbonControl)
Dim chromePath As String

On Error GoTo DefBrowser
chromePath = """C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"""
Shell (chromePath & " -url " & "https://www.ishakkutlu.com")
GoTo JumpingDefBrowser
DefBrowser:
On Error Resume Next
ThisWorkbook.FollowHyperlink ("https://ishakkutlu.com")
On Error GoTo 0
JumpingDefBrowser:

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

Sub timeout(duration_ms As Double)
Dim Start_Time As Variant

Start_Time = Timer
Do
DoEvents
Loop Until (Timer - Start_Time) >= duration_ms

End Sub




