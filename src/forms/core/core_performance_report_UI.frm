VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} core_performance_report_UI 
   Caption         =   "Enterprise Document Automation System"
   ClientHeight    =   10890
   ClientLeft      =   30
   ClientTop       =   375
   ClientWidth     =   13920
   OleObjectBlob   =   "core_performance_report_UI.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "core_performance_report_UI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Public OpenWordTakip As Boolean

Private Sub FirstDateLabel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
FirstDateLabel.BackColor = RGB(60, 100, 180) 'RGB(60, 100, 180) 'RGB(60, 100, 180)
FirstDateLabel.ForeColor = RGB(255, 255, 255)
End Sub
Private Sub FirstDateText_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub

Private Sub LastDateLabel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
LastDateLabel.BackColor = RGB(60, 100, 180) 'RGB(60, 100, 180) 'RGB(60, 100, 180)
LastDateLabel.ForeColor = RGB(255, 255, 255)
End Sub
Private Sub LastDateText_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub

Private Sub Raporla_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
Raporla.BackColor = RGB(60, 100, 180) 'RGB(60, 100, 180) 'RGB(60, 100, 180)
Raporla.ForeColor = RGB(255, 255, 255)
End Sub
Private Sub Kapat_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
Kapat.BackColor = RGB(60, 100, 180) 'RGB(60, 100, 180) 'RGB(60, 100, 180)
Kapat.ForeColor = RGB(255, 255, 255)
End Sub
Private Sub Yardim_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
Yardim.BackColor = RGB(60, 100, 180) 'RGB(60, 100, 180) 'RGB(60, 100, 180)
Yardim.ForeColor = RGB(255, 255, 255)
End Sub

Sub ColorChangerGenel()

If Raporla.BackColor <> RGB(225, 235, 245) Then
Raporla.BackColor = RGB(225, 235, 245)
Raporla.ForeColor = RGB(30, 30, 30)
End If
If Kapat.BackColor <> RGB(225, 235, 245) Then
Kapat.BackColor = RGB(225, 235, 245)
Kapat.ForeColor = RGB(30, 30, 30)
End If
If Yardim.BackColor <> RGB(225, 235, 245) Then
Yardim.BackColor = RGB(225, 235, 245)
Yardim.ForeColor = RGB(30, 30, 30)
End If

If FirstDateLabel.BackColor <> RGB(254, 254, 254) Then
FirstDateLabel.BackColor = RGB(254, 254, 254)
FirstDateLabel.ForeColor = RGB(70, 70, 70)
End If
If LastDateLabel.BackColor <> RGB(254, 254, 254) Then
LastDateLabel.BackColor = RGB(254, 254, 254)
LastDateLabel.ForeColor = RGB(70, 70, 70)
End If

End Sub


Private Sub Frame1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub


Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub
Private Sub TasiyiciFrame_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub
Private Sub BaslikFrame_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub

Private Sub UstMenuFrame_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub
Private Sub BilgilendirmeFrame_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub
Private Sub LblBilgilendirme_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub
Private Sub AltMenuFrame_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub

Private Sub Kapat_Click()
    Unload Me
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

Private Sub Yardim_Click()
Dim a() As Variant, i As Variant, Tanimlar As String, DestTanimlar As String
Dim AutoPath As String, DestRapor3lar As String, OpenControl As String
Dim Rapor3_2File As String, SayHedef As Long, ItemName As String, j As Integer

Dim fso As Object, objWord As Object, objDoc As Object, Birimx As String, Kurum_A As String, Kurum_ANoStr As String
Dim Rapor3_1File As String, Rapor3_1TipBFile As String, FinansalBirimFile As String, RaporFile As String, Rapor1File As String
Dim DestRapor2 As String, DestRapor1 As String

Dim DestOpUserFolderName As String, DestOpUserFolder As String, DestOperasyon As String
Dim OpenKontrolName As String, ContSay As Integer, KontrolFile As String
Dim ReNameTaslak As String, SourceTaslak As String

Application.ScreenUpdating = False
Application.DisplayAlerts = False

'Pathfinder...
AutoPath = ThisWorkbook.Path
DestOperasyon = AutoPath & "\System Files\Operation\"
'Taslak
SourceTaslak = AutoPath & "\System Files\Help Documents\Performance Report Panel – Help.docm"
'Birden fazla kullanıcının operasyon klasöründe çakışmasını önlemek için operasyon klasörünün içinde yeni klasör aç.
DestOpUserFolderName = "Operation-" & Left(ThisWorkbook.name, InStr(ThisWorkbook.name, ".") - 1) '& " " & Format(Now(), "ddmmyyyyhhmmss")
DestOpUserFolder = DestOperasyon & DestOpUserFolderName & "\"

'Check System Files folder name
If Not Dir(AutoPath & "\System Files\", vbDirectory) <> vbNullString Then
    MsgBox "Cannot access the directory: " & AutoPath & "\System Files\" & ". The folder named 'System Files' might have been renamed or deleted.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If
'Check Operations folder name
If Not Dir(DestOperasyon, vbDirectory) <> vbNullString Then
    MsgBox "Cannot access the directory: " & DestOperasyon & ". The folder named 'Operations' might have been renamed or deleted.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If


'RmDir DestOpUserFolder 'Sistem kapanırken DestOpUserFolder klasörünü temizle EKLENECEK!
'_______________

'Check folder names
If Not Dir(SourceTaslak, vbDirectory) <> vbNullString Then
    MsgBox "Cannot access the directory: " & SourceTaslak & ". Folder and/or file names in this directory might have been changed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
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

Application.ScreenUpdating = True
Application.DisplayAlerts = True

End Sub


Private Sub Raporla_Click()
Dim AutoPath As String, PerformansRaporu As String, PerformansRaporuKlasor As String, WsPerfRap As Object
Dim OpenControl As String, DestOperasyon As String, DestOpUserFolderName As String, DestOpUserFolder As String
Dim PerformansRaporuOp As String, ContSay As Integer, KontrolFile As String, fso As Object
Dim SayRapor1 As Long, SayRapor As Long, SayRapor3_2 As Long, SayRapor3_1 As Long, Cont As Integer, AdetSay As Integer
Dim IlkTarihBulRapor1 As Range, SonTarihBulRapor1 As Range, IlkTarihRapor1 As String, SonTarihRapor1 As String
Dim IlkTarihBulRapor As Range, SonTarihBulRapor As Range, IlkTarihRapor As String, SonTarihRapor As String

Dim IlkTarihBulRapor3_2 As Range, SonTarihBulRapor3_2 As Range, IlkTarihRapor3_2 As String, SonTarihRapor3_2 As String
Dim IlkTarihBulRapor3_1 As Range, SonTarihBulRapor3_1 As Range, IlkTarihRapor3_1 As String, SonTarihRapor3_1 As String

Dim Rapor1Kont As Integer, RaporKont As Integer, Rapor3_2Kont As Integer, Rapor3_1Kont As Integer, i As Long, j As Long, WsPerfRapSay As Long
Dim Kenarlar As Range

Dim x As Long, Count As Integer, Countx As Integer, Maxi As Long
Dim sDayName As String, myDate As Date, ToplamTalep As Long
Dim DateCtrl As Integer, Dec As Double
Dim GelenTema As String

Dim BilgilendirmeKont, XXXMudKont, SonucKont As Integer
Dim IlkTarihXXXMud, SonTarihXXXMud As String
Dim IlkTarihBulXXXMud, SonTarihBulXXXMud As Range
Dim IlkTarihSonuc, SonTarihSonuc As String
Dim IlkTarihBulSonuc, SonTarihBulSonuc As Range
Dim StrContent As String

ThisWorkbook.Activate

Application.DisplayAlerts = False
Application.EnableEvents = False
Application.ScreenUpdating = False

'Pathfinder...
AutoPath = ThisWorkbook.Path
DestOperasyon = AutoPath & "\System Files\Operation\"
PerformansRaporuKlasor = AutoPath & "\System Files\System Templates\Performance Report\"
PerformansRaporu = PerformansRaporuKlasor & "Performance Report.xlsx"

'Birden fazla kullanıcının operasyon klasöründe çakışmasını önlemek için operasyon klasörünün içinde yeni klasör aç.
DestOpUserFolderName = "Operation-" & Left(ThisWorkbook.name, InStr(ThisWorkbook.name, ".") - 1) '& " " & Format(Now(), "ddmmyyyyhhmmss")
DestOpUserFolder = DestOperasyon & DestOpUserFolderName & "\"

'Check Perf folder name
If Not Dir(PerformansRaporuKlasor, vbDirectory) <> vbNullString Then
    MsgBox "Cannot access the directory: " & PerformansRaporuKlasor & ". The folder named 'Registry Reports' may have been renamed or deleted.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Out
End If
If Not Dir(PerformansRaporu, vbDirectory) <> vbNullString Then
    MsgBox "Cannot access the directory: " & PerformansRaporu & ". Folder and/or file names in this directory may have been changed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Out
End If
'Check System Files folder name
If Not Dir(AutoPath & "\System Files\", vbDirectory) <> vbNullString Then
    MsgBox "Cannot access the directory: " & AutoPath & "\System Files\" & ". The folder named 'System Files' may have been renamed or deleted.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Out
End If
'Check Operation folder name
If Not Dir(DestOperasyon, vbDirectory) <> vbNullString Then
    MsgBox "Cannot access the directory: " & DestOperasyon & ". The folder named 'Operation' may have been renamed or deleted.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Out
End If

'Operation klasörü içinde kullanıcı modülü klasörü yoksa oluştur.
If Not Dir(DestOpUserFolder, vbDirectory) <> vbNullString Then
    MkDir DestOpUserFolder
End If

'Tarih kontrolleri
If FirstDateText.Value = "" Or LastDateText.Value = "" Then
    GoTo Out
End If

If Year(FirstDateText.Value) > Year(LastDateText.Value) Or _
(Year(FirstDateText.Value) = Year(LastDateText.Value) And Month(FirstDateText.Value) > Month(LastDateText.Value)) Or _
(Year(FirstDateText.Value) = Year(LastDateText.Value) And Month(FirstDateText.Value) = Month(LastDateText.Value) And Day(FirstDateText.Value) > Day(LastDateText.Value)) Then
    MsgBox "The process cannot proceed due to a mismatch between the start and end dates. Please ensure that the end date is not earlier than the start date.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Out
End If

'Invalidty verilerini kontrol et.
SayRapor1 = ThisWorkbook.Worksheets(3).Range("CF100000").End(xlUp).Row
SayRapor3_2 = ThisWorkbook.Worksheets(5).Range("FH100000").End(xlUp).Row
SayRapor = ThisWorkbook.Worksheets(4).Range("CN100000").End(xlUp).Row

If SayRapor1 < 7 And SayRapor < 7 And SayRapor3_2 < 7 Then
    MsgBox "Your process cannot be completed because no data was found in Report 1 or Report 2 workflows.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Out
End If

Rapor1Kont = 0
Rapor3_2Kont = 0
Rapor3_1Kont = 0
RaporKont = 0
'BilgilendirmeKont = 0
XXXMudKont = 0
SonucKont = 0


'______________________RAPOR1

IlkTarihRapor1 = FirstDateText.Value
SonTarihRapor1 = LastDateText.Value

Cont = 0
TarihiTekrarla1:
Set IlkTarihBulRapor1 = ThisWorkbook.Worksheets(3).Range("BW6:BW100000").Find(What:=IlkTarihRapor1, SearchDirection:=xlNext, _
                SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
If Not IlkTarihBulRapor1 Is Nothing Then
    '
Else
    IlkTarihRapor1 = CDate(IlkTarihRapor1)
    IlkTarihRapor1 = DateAdd("d", 1, IlkTarihRapor1)
    IlkTarihRapor1 = CStr(IlkTarihRapor1)
    If Mid(IlkTarihRapor1, 2, 1) = "." Then 'Günün soluna 0 ekle
        IlkTarihRapor1 = "0" & IlkTarihRapor1
    End If
    If Mid(IlkTarihRapor1, 5, 1) = "." Then 'Ayın soluna 0 ekle
        IlkTarihRapor1 = Left(IlkTarihRapor1, 3) & "0" & Mid(IlkTarihRapor1, 4, 6)
    End If
    'MsgBox IlkTarih
    Cont = Cont + 1
    If Cont = 1460 Then '4 yıl
        Rapor1Kont = 1
        GoTo Rapor3danDevam
    End If
    GoTo TarihiTekrarla1
End If


Cont = 0
TarihiTekrarla2:
Set SonTarihBulRapor1 = ThisWorkbook.Worksheets(3).Range("BW6:BW100000").Find(What:=SonTarihRapor1, SearchDirection:=xlPrevious, _
                SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
If Not SonTarihBulRapor1 Is Nothing Then
    '
Else
    SonTarihRapor1 = CDate(SonTarihRapor1)
    SonTarihRapor1 = DateAdd("d", -1, SonTarihRapor1)
    SonTarihRapor1 = CStr(SonTarihRapor1)
    If Mid(SonTarihRapor1, 2, 1) = "." Then 'Günün soluna 0 ekle
        SonTarihRapor1 = "0" & SonTarihRapor1
    End If
    If Mid(SonTarihRapor1, 5, 1) = "." Then 'Ayın soluna 0 ekle
        SonTarihRapor1 = Left(SonTarihRapor1, 3) & "0" & Mid(SonTarihRapor1, 4, 6)
    End If
    'MsgBox SonTarih
    Cont = Cont + 1
    If Cont = 1460 Then '4 yıl
        Rapor1Kont = 1
        GoTo Rapor3danDevam
    End If
    GoTo TarihiTekrarla2
End If


'______________________RAPOR3

Rapor3danDevam:

'____________________RAPOR3_2

IlkTarihRapor3_2 = FirstDateText.Value
SonTarihRapor3_2 = LastDateText.Value

Cont = 0
TarihiTekrarla1x:
Set IlkTarihBulRapor3_2 = ThisWorkbook.Worksheets(5).Range("CE6:CE100000").Find(What:=IlkTarihRapor3_2, SearchDirection:=xlNext, _
                SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
If Not IlkTarihBulRapor3_2 Is Nothing Then
    '
Else
    IlkTarihRapor3_2 = CDate(IlkTarihRapor3_2)
    IlkTarihRapor3_2 = DateAdd("d", 1, IlkTarihRapor3_2)
    IlkTarihRapor3_2 = CStr(IlkTarihRapor3_2)
    If Mid(IlkTarihRapor3_2, 2, 1) = "." Then 'Günün soluna 0 ekle
        IlkTarihRapor3_2 = "0" & IlkTarihRapor3_2
    End If
    If Mid(IlkTarihRapor3_2, 5, 1) = "." Then 'Ayın soluna 0 ekle
        IlkTarihRapor3_2 = Left(IlkTarihRapor3_2, 3) & "0" & Mid(IlkTarihRapor3_2, 4, 6)
    End If
    'MsgBox IlkTarih
    Cont = Cont + 1
    If Cont = 1460 Then '4 yıl
        Rapor3_2Kont = 1
        GoTo Rapor3_1denDevam
    End If
    GoTo TarihiTekrarla1x
End If


Cont = 0
TarihiTekrarla2x:
Set SonTarihBulRapor3_2 = ThisWorkbook.Worksheets(5).Range("CE6:CE100000").Find(What:=SonTarihRapor3_2, SearchDirection:=xlPrevious, _
                SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
If Not SonTarihBulRapor3_2 Is Nothing Then
    '
Else
    SonTarihRapor3_2 = CDate(SonTarihRapor3_2)
    SonTarihRapor3_2 = DateAdd("d", -1, SonTarihRapor3_2)
    SonTarihRapor3_2 = CStr(SonTarihRapor3_2)
    If Mid(SonTarihRapor3_2, 2, 1) = "." Then 'Günün soluna 0 ekle
        SonTarihRapor3_2 = "0" & SonTarihRapor3_2
    End If
    If Mid(SonTarihRapor3_2, 5, 1) = "." Then 'Ayın soluna 0 ekle
        SonTarihRapor3_2 = Left(SonTarihRapor3_2, 3) & "0" & Mid(SonTarihRapor3_2, 4, 6)
    End If
    'MsgBox SonTarih
    Cont = Cont + 1
    If Cont = 1460 Then '4 yıl
        Rapor3_2Kont = 1
        GoTo Rapor3_1denDevam
    End If
    GoTo TarihiTekrarla2x
End If

'____________________RAPOR3_1

Rapor3_1denDevam:

IlkTarihRapor3_1 = FirstDateText.Value
SonTarihRapor3_1 = LastDateText.Value

Cont = 0
TarihiTekrarla1x1:
Set IlkTarihBulRapor3_1 = ThisWorkbook.Worksheets(5).Range("EY6:EY100000").Find(What:=IlkTarihRapor3_1, SearchDirection:=xlNext, _
                SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
If Not IlkTarihBulRapor3_1 Is Nothing Then
    '
Else
    IlkTarihRapor3_1 = CDate(IlkTarihRapor3_1)
    IlkTarihRapor3_1 = DateAdd("d", 1, IlkTarihRapor3_1)
    IlkTarihRapor3_1 = CStr(IlkTarihRapor3_1)
    If Mid(IlkTarihRapor3_1, 2, 1) = "." Then 'Günün soluna 0 ekle
        IlkTarihRapor3_1 = "0" & IlkTarihRapor3_1
    End If
    If Mid(IlkTarihRapor3_1, 5, 1) = "." Then 'Ayın soluna 0 ekle
        IlkTarihRapor3_1 = Left(IlkTarihRapor3_1, 3) & "0" & Mid(IlkTarihRapor3_1, 4, 6)
    End If
    'MsgBox IlkTarih
    Cont = Cont + 1
    If Cont = 1460 Then '4 yıl
        Rapor3_1Kont = 1
        GoTo RapordanDevam
    End If
    GoTo TarihiTekrarla1x1
End If


Cont = 0
TarihiTekrarla2x1:
Set SonTarihBulRapor3_1 = ThisWorkbook.Worksheets(5).Range("EY6:EY100000").Find(What:=SonTarihRapor3_1, SearchDirection:=xlPrevious, _
                SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
If Not SonTarihBulRapor3_1 Is Nothing Then
    '
Else
    SonTarihRapor3_1 = CDate(SonTarihRapor3_1)
    SonTarihRapor3_1 = DateAdd("d", -1, SonTarihRapor3_1)
    SonTarihRapor3_1 = CStr(SonTarihRapor3_1)
    If Mid(SonTarihRapor3_1, 2, 1) = "." Then 'Günün soluna 0 ekle
        SonTarihRapor3_1 = "0" & SonTarihRapor3_1
    End If
    If Mid(SonTarihRapor3_1, 5, 1) = "." Then 'Ayın soluna 0 ekle
        SonTarihRapor3_1 = Left(SonTarihRapor3_1, 3) & "0" & Mid(SonTarihRapor3_1, 4, 6)
    End If
    'MsgBox SonTarih
    Cont = Cont + 1
    If Cont = 1460 Then '4 yıl
        Rapor3_1Kont = 1
        GoTo RapordanDevam
    End If
    GoTo TarihiTekrarla2x1
End If



'______________________RAPOR ve BİLGİLENDİRME

RapordanDevam:

IlkTarihRapor = FirstDateText.Value
SonTarihRapor = LastDateText.Value

Cont = 0
TarihiTekrarla3:
Set IlkTarihBulRapor = ThisWorkbook.Worksheets(4).Range("CE6:CE100000").Find(What:=IlkTarihRapor, SearchDirection:=xlNext, _
                SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
If Not IlkTarihBulRapor Is Nothing Then
    '
Else
    IlkTarihRapor = CDate(IlkTarihRapor)
    IlkTarihRapor = DateAdd("d", 1, IlkTarihRapor)
    IlkTarihRapor = CStr(IlkTarihRapor)
    If Mid(IlkTarihRapor, 2, 1) = "." Then 'Günün soluna 0 ekle
        IlkTarihRapor = "0" & IlkTarihRapor
    End If
    If Mid(IlkTarihRapor, 5, 1) = "." Then 'Ayın soluna 0 ekle
        IlkTarihRapor = Left(IlkTarihRapor, 3) & "0" & Mid(IlkTarihRapor, 4, 6)
    End If
    'MsgBox IlkTarih
    Cont = Cont + 1
    If Cont = 1460 Then '4 yıl
        RaporKont = 1
        GoTo XXXMuddanDevam
    End If
    GoTo TarihiTekrarla3
End If


Cont = 0
TarihiTekrarla4:
Set SonTarihBulRapor = ThisWorkbook.Worksheets(4).Range("CE6:CE100000").Find(What:=SonTarihRapor, SearchDirection:=xlPrevious, _
                SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
If Not SonTarihBulRapor Is Nothing Then
    '
Else
    SonTarihRapor = CDate(SonTarihRapor)
    SonTarihRapor = DateAdd("d", -1, SonTarihRapor)
    SonTarihRapor = CStr(SonTarihRapor)
    If Mid(SonTarihRapor, 2, 1) = "." Then 'Günün soluna 0 ekle
        SonTarihRapor = "0" & SonTarihRapor
    End If
    If Mid(SonTarihRapor, 5, 1) = "." Then 'Ayın soluna 0 ekle
        SonTarihRapor = Left(SonTarihRapor, 3) & "0" & Mid(SonTarihRapor, 4, 6)
    End If
    'MsgBox SonTarih
    Cont = Cont + 1
    If Cont = 1460 Then '4 yıl
        RaporKont = 1
        GoTo XXXMuddanDevam
    End If
    GoTo TarihiTekrarla4
End If


'______________________XXXMud

XXXMuddanDevam:

IlkTarihXXXMud = FirstDateText.Value
SonTarihXXXMud = LastDateText.Value

Cont = 0
TarihiTekrarla5:
Set IlkTarihBulXXXMud = ThisWorkbook.Worksheets(4).Range("FS6:FS100000").Find(What:=IlkTarihXXXMud, SearchDirection:=xlNext, _
                SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
If Not IlkTarihBulXXXMud Is Nothing Then
    '
Else
    IlkTarihXXXMud = CDate(IlkTarihXXXMud)
    IlkTarihXXXMud = DateAdd("d", 1, IlkTarihXXXMud)
    IlkTarihXXXMud = CStr(IlkTarihXXXMud)
    If Mid(IlkTarihXXXMud, 2, 1) = "." Then 'Günün soluna 0 ekle
        IlkTarihXXXMud = "0" & IlkTarihXXXMud
    End If
    If Mid(IlkTarihXXXMud, 5, 1) = "." Then 'Ayın soluna 0 ekle
        IlkTarihXXXMud = Left(IlkTarihXXXMud, 3) & "0" & Mid(IlkTarihXXXMud, 4, 6)
    End If
    'MsgBox IlkTarih
    Cont = Cont + 1
    If Cont = 1460 Then '4 yıl
        XXXMudKont = 1
        GoTo SonucdanDevam
    End If
    GoTo TarihiTekrarla5
End If


Cont = 0
TarihiTekrarla6:
Set SonTarihBulXXXMud = ThisWorkbook.Worksheets(4).Range("FS6:FS100000").Find(What:=SonTarihXXXMud, SearchDirection:=xlPrevious, _
                SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
If Not SonTarihBulXXXMud Is Nothing Then
    '
Else
    SonTarihXXXMud = CDate(SonTarihXXXMud)
    SonTarihXXXMud = DateAdd("d", -1, SonTarihXXXMud)
    SonTarihXXXMud = CStr(SonTarihXXXMud)
    If Mid(SonTarihXXXMud, 2, 1) = "." Then 'Günün soluna 0 ekle
        SonTarihXXXMud = "0" & SonTarihXXXMud
    End If
    If Mid(SonTarihXXXMud, 5, 1) = "." Then 'Ayın soluna 0 ekle
        SonTarihXXXMud = Left(SonTarihXXXMud, 3) & "0" & Mid(SonTarihXXXMud, 4, 6)
    End If
    'MsgBox SonTarih
    Cont = Cont + 1
    If Cont = 1460 Then '4 yıl
        XXXMudKont = 1
        GoTo SonucdanDevam
    End If
    GoTo TarihiTekrarla6
End If



'______________________SONUÇ

SonucdanDevam:

IlkTarihSonuc = FirstDateText.Value
SonTarihSonuc = LastDateText.Value

Cont = 0
TarihiTekrarla7:
Set IlkTarihBulSonuc = ThisWorkbook.Worksheets(4).Range("FU6:FU100000").Find(What:=IlkTarihSonuc, SearchDirection:=xlNext, _
                SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
If Not IlkTarihBulSonuc Is Nothing Then
    '
Else
    IlkTarihSonuc = CDate(IlkTarihSonuc)
    IlkTarihSonuc = DateAdd("d", 1, IlkTarihSonuc)
    IlkTarihSonuc = CStr(IlkTarihSonuc)
    If Mid(IlkTarihSonuc, 2, 1) = "." Then 'Günün soluna 0 ekle
        IlkTarihSonuc = "0" & IlkTarihSonuc
    End If
    If Mid(IlkTarihSonuc, 5, 1) = "." Then 'Ayın soluna 0 ekle
        IlkTarihSonuc = Left(IlkTarihSonuc, 3) & "0" & Mid(IlkTarihSonuc, 4, 6)
    End If
    'MsgBox IlkTarih
    Cont = Cont + 1
    If Cont = 1460 Then '4 yıl
        SonucKont = 1
        GoTo DerlemedenDevam
    End If
    GoTo TarihiTekrarla7
End If


Cont = 0
TarihiTekrarla8:
Set SonTarihBulSonuc = ThisWorkbook.Worksheets(4).Range("FU6:FU100000").Find(What:=SonTarihSonuc, SearchDirection:=xlPrevious, _
                SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
If Not SonTarihBulSonuc Is Nothing Then
    '
Else
    SonTarihSonuc = CDate(SonTarihSonuc)
    SonTarihSonuc = DateAdd("d", -1, SonTarihSonuc)
    SonTarihSonuc = CStr(SonTarihSonuc)
    If Mid(SonTarihSonuc, 2, 1) = "." Then 'Günün soluna 0 ekle
        SonTarihSonuc = "0" & SonTarihSonuc
    End If
    If Mid(SonTarihSonuc, 5, 1) = "." Then 'Ayın soluna 0 ekle
        SonTarihSonuc = Left(SonTarihSonuc, 3) & "0" & Mid(SonTarihSonuc, 4, 6)
    End If
    'MsgBox SonTarih
    Cont = Cont + 1
    If Cont = 1460 Then '4 yıl
        SonucKont = 1
        GoTo DerlemedenDevam
    End If
    GoTo TarihiTekrarla8
End If




DerlemedenDevam:

'______________________DERLEMEYE BAŞLA

If Rapor1Kont = 1 And Rapor3_2Kont = 1 And Rapor3_1Kont = 1 And RaporKont = 1 And XXXMudKont = 1 And SonucKont = 1 Then
    MsgBox "Your process cannot be completed because no response letter date was found within the specified date range in Report1, Report3, and report/report2_2 modules.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Out
End If


On Error Resume Next 'Operation içinde Performance Report.xlsx dosyası yoksa oluşacak hata için
PerformansRaporuOp = DestOpUserFolder & "Performance Report.xlsx"
'İşlem günlüğü açıksa kaydet ve kapat.
OpenControl = IsWorkBookOpen(PerformansRaporuOp)
If OpenControl = True Then
    Workbooks("Performance Report.xlsx").Close SaveChanges:=True
End If

'Klasörün içindeki tüm dosyaları sil (txt, docm vb.)
On Error Resume Next
ContSay = 0
KontrolFile = Dir(DestOpUserFolder & "*.???")
Do While KontrolFile <> ""
    ContSay = ContSay + 1
    KontrolFile = Dir()
Loop
If ContSay > 0 Then
    Kill DestOpUserFolder & "*.???"
End If

'Dosyayı operasyon klasörüne kopyala ve adını değiştir.
Set fso = CreateObject("Scripting.FileSystemObject")
fso.CopyFile (PerformansRaporu), DestOpUserFolder & "Performance Report" & ".xlsx", True
'open the file
Workbooks.Open (PerformansRaporuOp)

Set WsPerfRap = Workbooks("Performance Report.xlsx").Worksheets(1)

WsPerfRap.Unprotect Password:="123"

WsPerfRapSay = 9

'Rapor1 Aktarımı
If Rapor1Kont = 0 Then
    For i = IlkTarihBulRapor1.Row To SonTarihBulRapor1.Row
        If ThisWorkbook.Worksheets(3).Cells(i, 83).Value <> "" Then
            'Gelen tema
            GelenTema = ""
            If ThisWorkbook.Worksheets(3).Cells(i, 25).Value = "Provincial Directorate B" Then
                If ThisWorkbook.Worksheets(3).Cells(i, 26).Value <> "" Then
                    GelenTema = ThisWorkbook.Worksheets(3).Cells(i, 17).Value & " Provincial Governorship Provincial Directorate B " & ThisWorkbook.Worksheets(3).Cells(i, 26).Value
                Else
                    GelenTema = ThisWorkbook.Worksheets(3).Cells(i, 17).Value & " Provincial Governorship Provincial Directorate B"
                End If
            ElseIf ThisWorkbook.Worksheets(3).Cells(i, 25).Value = "District Directorate B" Then
                If ThisWorkbook.Worksheets(3).Cells(i, 26).Value <> "" Then
                    GelenTema = ThisWorkbook.Worksheets(3).Cells(i, 18).Value & " District Governorship District Directorate B " & ThisWorkbook.Worksheets(3).Cells(i, 26).Value
                Else
                    GelenTema = ThisWorkbook.Worksheets(3).Cells(i, 18).Value & " District Governorship District Directorate B"
                End If
            ElseIf ThisWorkbook.Worksheets(3).Cells(i, 25).Value = "Provincial Directorate C" Then
                If ThisWorkbook.Worksheets(3).Cells(i, 26).Value <> "" Then
                    GelenTema = ThisWorkbook.Worksheets(3).Cells(i, 17).Value & " Provincial Governorship Provincial Directorate C " & ThisWorkbook.Worksheets(3).Cells(i, 26).Value
                Else
                    GelenTema = ThisWorkbook.Worksheets(3).Cells(i, 17).Value & " Provincial Governorship Provincial Directorate C"
                End If
            ElseIf ThisWorkbook.Worksheets(3).Cells(i, 25).Value = "District Directorate C" Then
                If ThisWorkbook.Worksheets(3).Cells(i, 26).Value <> "" Then
                    GelenTema = ThisWorkbook.Worksheets(3).Cells(i, 18).Value & " District Governorship District Directorate C " & ThisWorkbook.Worksheets(3).Cells(i, 26).Value
                Else
                    GelenTema = ThisWorkbook.Worksheets(3).Cells(i, 18).Value & " District Governorship District Directorate C"
                End If
            ElseIf ThisWorkbook.Worksheets(3).Cells(i, 25).Value = "Provincial Directorate D" Then
                If ThisWorkbook.Worksheets(3).Cells(i, 26).Value <> "" Then
                    GelenTema = ThisWorkbook.Worksheets(3).Cells(i, 17).Value & " Provincial Governorship Provincial Directorate D " & ThisWorkbook.Worksheets(3).Cells(i, 26).Value
                Else
                    GelenTema = ThisWorkbook.Worksheets(3).Cells(i, 17).Value & " Provincial Governorship Provincial Directorate D"
                End If
            ElseIf ThisWorkbook.Worksheets(3).Cells(i, 25).Value = "District Directorate D" Then
                If ThisWorkbook.Worksheets(3).Cells(i, 26).Value <> "" Then
                    GelenTema = ThisWorkbook.Worksheets(3).Cells(i, 18).Value & " District Governorship District Directorate D " & ThisWorkbook.Worksheets(3).Cells(i, 26).Value
                Else
                    GelenTema = ThisWorkbook.Worksheets(3).Cells(i, 18).Value & " District Governorship District Directorate D"
                End If
            ElseIf ThisWorkbook.Worksheets(3).Cells(i, 25).Value = "Provincial Directorate E" Then
                If ThisWorkbook.Worksheets(3).Cells(i, 26).Value <> "" Then
                    GelenTema = ThisWorkbook.Worksheets(3).Cells(i, 17).Value & " Provincial Governorship Provincial Directorate E " & ThisWorkbook.Worksheets(3).Cells(i, 26).Value
                Else
                    GelenTema = ThisWorkbook.Worksheets(3).Cells(i, 17).Value & " Provincial Governorship Provincial Directorate E"
                End If
            ElseIf ThisWorkbook.Worksheets(3).Cells(i, 25).Value = "District Directorate E" Then
                If ThisWorkbook.Worksheets(3).Cells(i, 26).Value <> "" Then
                    GelenTema = ThisWorkbook.Worksheets(3).Cells(i, 18).Value & " District Governorship District Directorate E " & ThisWorkbook.Worksheets(3).Cells(i, 26).Value
                Else
                    GelenTema = ThisWorkbook.Worksheets(3).Cells(i, 18).Value & " District Governorship District Directorate E"
                End If
            ElseIf InStr(ThisWorkbook.Worksheets(3).Cells(i, 25).Value, "General Directorate") <> 0 Or InStr(ThisWorkbook.Worksheets(3).Cells(i, 25).Value, "Regional Directorate") <> 0 Then
                If ThisWorkbook.Worksheets(3).Cells(i, 26).Value <> "" Then
                    GelenTema = WorksheetFunction.Proper(ThisWorkbook.Worksheets(2).Cells(6, 111).Value) & " " & ThisWorkbook.Worksheets(3).Cells(i, 25).Value & " " & ThisWorkbook.Worksheets(3).Cells(i, 26).Value
                Else
                    GelenTema = WorksheetFunction.Proper(ThisWorkbook.Worksheets(2).Cells(6, 111).Value) & " " & ThisWorkbook.Worksheets(3).Cells(i, 25).Value
                End If
            Else
                If ThisWorkbook.Worksheets(3).Cells(i, 26).Value <> "" Then
                    If InStr(ThisWorkbook.Worksheets(3).Cells(i, 25).Value, "X.X. ") > 0 Then
                        GelenTema = Mid(ThisWorkbook.Worksheets(3).Cells(i, 25).Value, 6, Len(ThisWorkbook.Worksheets(3).Cells(i, 25).Value)) & " " & ThisWorkbook.Worksheets(3).Cells(i, 26).Value
                    Else
                        GelenTema = ThisWorkbook.Worksheets(3).Cells(i, 25).Value & " " & ThisWorkbook.Worksheets(3).Cells(i, 26).Value
                    End If
                Else
                    If InStr(ThisWorkbook.Worksheets(3).Cells(i, 25).Value, "X.X. ") > 0 Then
                        GelenTema = Mid(ThisWorkbook.Worksheets(3).Cells(i, 25).Value, 6, Len(ThisWorkbook.Worksheets(3).Cells(i, 25).Value))
                    Else
                        GelenTema = ThisWorkbook.Worksheets(3).Cells(i, 25).Value
                    End If
                End If
            End If

            'Gönderen
            WsPerfRap.Cells(WsPerfRapSay, 3) = GelenTema
            'Rapor Tipi
            WsPerfRap.Cells(WsPerfRapSay, 4) = "Report 1"
            'Rapor Adedi
            AdetSay = 0
            For j = i To i + 50
                If ThisWorkbook.Worksheets(3).Cells(j, 84).Value = "" Then
                    If ThisWorkbook.Worksheets(3).Cells(j, 11).Value <> "" Then
                        AdetSay = AdetSay + 1
                        WsPerfRap.Cells(WsPerfRapSay, 5) = AdetSay
                    End If
                ElseIf ThisWorkbook.Worksheets(3).Cells(j, 84).Value <> "" Then
                    If ThisWorkbook.Worksheets(3).Cells(j, 11).Value <> "" Then
                        AdetSay = AdetSay + 1
                        WsPerfRap.Cells(WsPerfRapSay, 5) = AdetSay
                        GoTo jDongusuSonRapor1
                    End If
                    GoTo jDongusuSonRapor1
                End If
            Next j
jDongusuSonRapor1:
            'Geliş Tarihi
            WsPerfRap.Cells(WsPerfRapSay, 6) = ThisWorkbook.Worksheets(3).Cells(i, 28)
            'Üst Yazı Tarihi
            WsPerfRap.Cells(WsPerfRapSay, 7) = ThisWorkbook.Worksheets(3).Cells(i, 75)
            'Sıra No
            WsPerfRap.Cells(WsPerfRapSay, 2) = WsPerfRapSay - 8
            
            WsPerfRapSay = WsPerfRapSay + 1
        End If
    Next i
End If


'Rapor3_2 Aktarımı
If Rapor3_2Kont = 0 Then
    For i = IlkTarihBulRapor3_2.Row To SonTarihBulRapor3_2.Row
        If ThisWorkbook.Worksheets(5).Cells(i, 163).Value <> "" And ThisWorkbook.Worksheets(5).Cells(i, 12).Value = "Point2" Then
            'Gelen tema
            GelenTema = ""
            If ThisWorkbook.Worksheets(5).Cells(i, 47).Value = "Provincial Directorate B" Then
                If ThisWorkbook.Worksheets(5).Cells(i, 48).Value <> "" Then
                    GelenTema = ThisWorkbook.Worksheets(5).Cells(i, 19).Value & " Provincial Governorship Provincial Directorate B " & ThisWorkbook.Worksheets(5).Cells(i, 48).Value
                Else
                    GelenTema = ThisWorkbook.Worksheets(5).Cells(i, 19).Value & " Provincial Governorship Provincial Directorate B"
                End If
            ElseIf ThisWorkbook.Worksheets(5).Cells(i, 47).Value = "District Directorate B" Then
                If ThisWorkbook.Worksheets(5).Cells(i, 48).Value <> "" Then
                    GelenTema = ThisWorkbook.Worksheets(5).Cells(i, 20).Value & " District Governorship District Directorate B " & ThisWorkbook.Worksheets(5).Cells(i, 48).Value
                Else
                    GelenTema = ThisWorkbook.Worksheets(5).Cells(i, 20).Value & " District Governorship District Directorate B"
                End If
            ElseIf ThisWorkbook.Worksheets(5).Cells(i, 47).Value = "Provincial Directorate C" Then
                If ThisWorkbook.Worksheets(5).Cells(i, 48).Value <> "" Then
                    GelenTema = ThisWorkbook.Worksheets(5).Cells(i, 19).Value & " Provincial Governorship Provincial Directorate C " & ThisWorkbook.Worksheets(5).Cells(i, 48).Value
                Else
                    GelenTema = ThisWorkbook.Worksheets(5).Cells(i, 19).Value & " Provincial Governorship Provincial Directorate C"
                End If
            ElseIf ThisWorkbook.Worksheets(5).Cells(i, 47).Value = "District Directorate C" Then
                If ThisWorkbook.Worksheets(5).Cells(i, 48).Value <> "" Then
                    GelenTema = ThisWorkbook.Worksheets(5).Cells(i, 20).Value & " District Governorship District Directorate C " & ThisWorkbook.Worksheets(5).Cells(i, 48).Value
                Else
                    GelenTema = ThisWorkbook.Worksheets(5).Cells(i, 20).Value & " District Governorship District Directorate C"
                End If
            ElseIf ThisWorkbook.Worksheets(5).Cells(i, 47).Value = "Provincial Directorate D" Then
                If ThisWorkbook.Worksheets(5).Cells(i, 48).Value <> "" Then
                    GelenTema = ThisWorkbook.Worksheets(5).Cells(i, 19).Value & " Provincial Governorship Provincial Directorate D " & ThisWorkbook.Worksheets(5).Cells(i, 48).Value
                Else
                    GelenTema = ThisWorkbook.Worksheets(5).Cells(i, 19).Value & " Provincial Governorship Provincial Directorate D"
                End If
            ElseIf ThisWorkbook.Worksheets(5).Cells(i, 47).Value = "District Directorate D" Then
                If ThisWorkbook.Worksheets(5).Cells(i, 48).Value <> "" Then
                    GelenTema = ThisWorkbook.Worksheets(5).Cells(i, 20).Value & " District Governorship District Directorate D " & ThisWorkbook.Worksheets(5).Cells(i, 48).Value
                Else
                    GelenTema = ThisWorkbook.Worksheets(5).Cells(i, 20).Value & " District Governorship District Directorate D"
                End If
            ElseIf ThisWorkbook.Worksheets(5).Cells(i, 47).Value = "Provincial Directorate E" Then
                If ThisWorkbook.Worksheets(5).Cells(i, 48).Value <> "" Then
                    GelenTema = ThisWorkbook.Worksheets(5).Cells(i, 19).Value & " Provincial Governorship Provincial Directorate E " & ThisWorkbook.Worksheets(5).Cells(i, 48).Value
                Else
                    GelenTema = ThisWorkbook.Worksheets(5).Cells(i, 19).Value & " Provincial Governorship Provincial Directorate E"
                End If
            ElseIf ThisWorkbook.Worksheets(5).Cells(i, 47).Value = "District Directorate E" Then
                If ThisWorkbook.Worksheets(5).Cells(i, 48).Value <> "" Then
                    GelenTema = ThisWorkbook.Worksheets(5).Cells(i, 20).Value & " District Governorship District Directorate E " & ThisWorkbook.Worksheets(5).Cells(i, 48).Value
                Else
                    GelenTema = ThisWorkbook.Worksheets(5).Cells(i, 20).Value & " District Governorship District Directorate E"
                End If
            ElseIf InStr(ThisWorkbook.Worksheets(5).Cells(i, 47).Value, "General Directorate") <> 0 Or InStr(ThisWorkbook.Worksheets(5).Cells(i, 47).Value, "Regional Directorate") <> 0 Then
                If ThisWorkbook.Worksheets(5).Cells(i, 48).Value <> "" Then
                    GelenTema = WorksheetFunction.Proper(ThisWorkbook.Worksheets(2).Cells(6, 111).Value) & " " & ThisWorkbook.Worksheets(5).Cells(i, 47).Value & " " & ThisWorkbook.Worksheets(5).Cells(i, 48).Value
                Else
                    GelenTema = WorksheetFunction.Proper(ThisWorkbook.Worksheets(2).Cells(6, 111).Value) & " " & ThisWorkbook.Worksheets(5).Cells(i, 47).Value
                End If
            Else
                If ThisWorkbook.Worksheets(5).Cells(i, 48).Value <> "" Then
                    If InStr(ThisWorkbook.Worksheets(5).Cells(i, 47).Value, "X.X. ") > 0 Then
                        GelenTema = Mid(ThisWorkbook.Worksheets(5).Cells(i, 47).Value, 6, Len(ThisWorkbook.Worksheets(5).Cells(i, 47).Value)) & " " & ThisWorkbook.Worksheets(5).Cells(i, 48).Value
                    Else
                        GelenTema = ThisWorkbook.Worksheets(5).Cells(i, 47).Value & " " & ThisWorkbook.Worksheets(5).Cells(i, 48).Value
                    End If
                Else
                    If InStr(ThisWorkbook.Worksheets(5).Cells(i, 47).Value, "X.X. ") > 0 Then
                        GelenTema = Mid(ThisWorkbook.Worksheets(5).Cells(i, 47).Value, 6, Len(ThisWorkbook.Worksheets(5).Cells(i, 47).Value))
                    Else
                        GelenTema = ThisWorkbook.Worksheets(5).Cells(i, 47).Value
                    End If
                End If
            End If

            'Gönderen
            WsPerfRap.Cells(WsPerfRapSay, 3) = GelenTema
            'Rapor Tipi
            WsPerfRap.Cells(WsPerfRapSay, 4) = "Report 3.2"
            'Rapor Adedi
            AdetSay = 0
            For j = i To i + 50
                If ThisWorkbook.Worksheets(5).Cells(j, 164).Value = "" Then
                    If ThisWorkbook.Worksheets(5).Cells(j, 13).Value <> "" Then
                        AdetSay = AdetSay + 1
                        WsPerfRap.Cells(WsPerfRapSay, 5) = AdetSay
                    End If
                ElseIf ThisWorkbook.Worksheets(5).Cells(j, 164).Value <> "" Then
                    If ThisWorkbook.Worksheets(5).Cells(j, 13).Value <> "" Then
                        AdetSay = AdetSay + 1
                        WsPerfRap.Cells(WsPerfRapSay, 5) = AdetSay
                        GoTo jDongusuSonRapor3_2
                    End If
                    GoTo jDongusuSonRapor3_2
                End If
            Next j
jDongusuSonRapor3_2:
            'Geliş Tarihi
            WsPerfRap.Cells(WsPerfRapSay, 6) = ThisWorkbook.Worksheets(5).Cells(i, 23)
            'Üst Yazı Tarihi
            WsPerfRap.Cells(WsPerfRapSay, 7) = ThisWorkbook.Worksheets(5).Cells(i, 83)
            'Sıra No
            WsPerfRap.Cells(WsPerfRapSay, 2) = WsPerfRapSay - 8
            
            WsPerfRapSay = WsPerfRapSay + 1
        End If
    Next i
End If

'Rapor3_1 Aktarımı
If Rapor3_1Kont = 0 Then
    For i = IlkTarihBulRapor3_1.Row To SonTarihBulRapor3_1.Row
        If ThisWorkbook.Worksheets(5).Cells(i, 163).Value <> "" And ThisWorkbook.Worksheets(5).Cells(i, 12).Value = "Point1" Then
            'Gelen tema
            GelenTema = ""
            If ThisWorkbook.Worksheets(5).Cells(i, 102).Value = "Provincial Directorate B" Then
                If ThisWorkbook.Worksheets(5).Cells(i, 103).Value <> "" Then
                    GelenTema = ThisWorkbook.Worksheets(5).Cells(i, 91).Value & " Provincial Governorship Provincial Directorate B " & ThisWorkbook.Worksheets(5).Cells(i, 103).Value
                Else
                    GelenTema = ThisWorkbook.Worksheets(5).Cells(i, 91).Value & " Provincial Governorship Provincial Directorate B"
                End If
            ElseIf ThisWorkbook.Worksheets(5).Cells(i, 102).Value = "District Directorate B" Then
                If ThisWorkbook.Worksheets(5).Cells(i, 103).Value <> "" Then
                    GelenTema = ThisWorkbook.Worksheets(5).Cells(i, 92).Value & " District Governorship District Directorate B " & ThisWorkbook.Worksheets(5).Cells(i, 103).Value
                Else
                    GelenTema = ThisWorkbook.Worksheets(5).Cells(i, 92).Value & " District Governorship District Directorate B"
                End If
            ElseIf ThisWorkbook.Worksheets(5).Cells(i, 102).Value = "Provincial Directorate C" Then
                If ThisWorkbook.Worksheets(5).Cells(i, 103).Value <> "" Then
                    GelenTema = ThisWorkbook.Worksheets(5).Cells(i, 91).Value & " Provincial Governorship Provincial Directorate C " & ThisWorkbook.Worksheets(5).Cells(i, 103).Value
                Else
                    GelenTema = ThisWorkbook.Worksheets(5).Cells(i, 91).Value & " Provincial Governorship Provincial Directorate C"
                End If
            ElseIf ThisWorkbook.Worksheets(5).Cells(i, 102).Value = "District Directorate C" Then
                If ThisWorkbook.Worksheets(5).Cells(i, 103).Value <> "" Then
                    GelenTema = ThisWorkbook.Worksheets(5).Cells(i, 92).Value & " District Governorship District Directorate C " & ThisWorkbook.Worksheets(5).Cells(i, 103).Value
                Else
                    GelenTema = ThisWorkbook.Worksheets(5).Cells(i, 92).Value & " District Governorship District Directorate C"
                End If
            ElseIf ThisWorkbook.Worksheets(5).Cells(i, 102).Value = "Provincial Directorate D" Then
                If ThisWorkbook.Worksheets(5).Cells(i, 103).Value <> "" Then
                    GelenTema = ThisWorkbook.Worksheets(5).Cells(i, 91).Value & " Provincial Governorship Provincial Directorate D " & ThisWorkbook.Worksheets(5).Cells(i, 103).Value
                Else
                    GelenTema = ThisWorkbook.Worksheets(5).Cells(i, 91).Value & " Provincial Governorship Provincial Directorate D"
                End If
            ElseIf ThisWorkbook.Worksheets(5).Cells(i, 102).Value = "District Directorate D" Then
                If ThisWorkbook.Worksheets(5).Cells(i, 103).Value <> "" Then
                    GelenTema = ThisWorkbook.Worksheets(5).Cells(i, 92).Value & " District Governorship District Directorate D " & ThisWorkbook.Worksheets(5).Cells(i, 103).Value
                Else
                    GelenTema = ThisWorkbook.Worksheets(5).Cells(i, 92).Value & " District Governorship District Directorate D"
                End If
            ElseIf ThisWorkbook.Worksheets(5).Cells(i, 102).Value = "Provincial Directorate E" Then
                If ThisWorkbook.Worksheets(5).Cells(i, 103).Value <> "" Then
                    GelenTema = ThisWorkbook.Worksheets(5).Cells(i, 91).Value & " Provincial Governorship Provincial Directorate E " & ThisWorkbook.Worksheets(5).Cells(i, 103).Value
                Else
                    GelenTema = ThisWorkbook.Worksheets(5).Cells(i, 91).Value & " Provincial Governorship Provincial Directorate E"
                End If
            ElseIf ThisWorkbook.Worksheets(5).Cells(i, 102).Value = "District Directorate E" Then
                If ThisWorkbook.Worksheets(5).Cells(i, 103).Value <> "" Then
                    GelenTema = ThisWorkbook.Worksheets(5).Cells(i, 92).Value & " District Governorship District Directorate E " & ThisWorkbook.Worksheets(5).Cells(i, 103).Value
                Else
                    GelenTema = ThisWorkbook.Worksheets(5).Cells(i, 92).Value & " District Governorship District Directorate E"
                End If
            ElseIf InStr(ThisWorkbook.Worksheets(5).Cells(i, 102).Value, "General Directorate") <> 0 Or InStr(ThisWorkbook.Worksheets(5).Cells(i, 102).Value, "Regional Directorate") <> 0 Then
                If ThisWorkbook.Worksheets(5).Cells(i, 103).Value <> "" Then
                    GelenTema = WorksheetFunction.Proper(ThisWorkbook.Worksheets(2).Cells(6, 111).Value) & " " & ThisWorkbook.Worksheets(5).Cells(i, 102).Value & " " & ThisWorkbook.Worksheets(5).Cells(i, 103).Value
                Else
                    GelenTema = WorksheetFunction.Proper(ThisWorkbook.Worksheets(2).Cells(6, 111).Value) & " " & ThisWorkbook.Worksheets(5).Cells(i, 102).Value
                End If
            Else
                If ThisWorkbook.Worksheets(5).Cells(i, 103).Value <> "" Then
                    If InStr(ThisWorkbook.Worksheets(5).Cells(i, 102).Value, "X.X. ") > 0 Then
                        GelenTema = Mid(ThisWorkbook.Worksheets(5).Cells(i, 102).Value, 6, Len(ThisWorkbook.Worksheets(5).Cells(i, 102).Value)) & " " & ThisWorkbook.Worksheets(5).Cells(i, 103).Value
                    Else
                        GelenTema = ThisWorkbook.Worksheets(5).Cells(i, 102).Value & " " & ThisWorkbook.Worksheets(5).Cells(i, 103).Value
                    End If
                Else
                    If InStr(ThisWorkbook.Worksheets(5).Cells(i, 102).Value, "X.X. ") > 0 Then
                        GelenTema = Mid(ThisWorkbook.Worksheets(5).Cells(i, 102).Value, 6, Len(ThisWorkbook.Worksheets(5).Cells(i, 102).Value))
                    Else
                        GelenTema = ThisWorkbook.Worksheets(5).Cells(i, 102).Value
                    End If
                End If
            End If

            'Gönderen
            WsPerfRap.Cells(WsPerfRapSay, 3) = GelenTema
            'Rapor Tipi
            WsPerfRap.Cells(WsPerfRapSay, 4) = "Report 3.1"
            'Rapor Adedi
            AdetSay = 0
            For j = i To i + 50
                If ThisWorkbook.Worksheets(5).Cells(j, 164).Value = "" Then
                    If ThisWorkbook.Worksheets(5).Cells(j, 13).Value <> "" Then
                        AdetSay = AdetSay + 1
                        WsPerfRap.Cells(WsPerfRapSay, 5) = AdetSay
                    End If
                ElseIf ThisWorkbook.Worksheets(5).Cells(j, 164).Value <> "" Then
                    If ThisWorkbook.Worksheets(5).Cells(j, 13).Value <> "" Then
                        AdetSay = AdetSay + 1
                        WsPerfRap.Cells(WsPerfRapSay, 5) = AdetSay
                        GoTo jDongusuSonRapor3_1
                    End If
                    GoTo jDongusuSonRapor3_1
                End If
            Next j
jDongusuSonRapor3_1:
            'Geliş Tarihi
            WsPerfRap.Cells(WsPerfRapSay, 6) = ThisWorkbook.Worksheets(5).Cells(i, 95)
            'Üst Yazı Tarihi
            WsPerfRap.Cells(WsPerfRapSay, 7) = ThisWorkbook.Worksheets(5).Cells(i, 155)
            'Sıra No
            WsPerfRap.Cells(WsPerfRapSay, 2) = WsPerfRapSay - 8
            
            WsPerfRapSay = WsPerfRapSay + 1
        End If
    Next i
End If


'Rapor ve Bilgilendirme Aktarımı
If RaporKont = 0 Then
    For i = IlkTarihBulRapor.Row To SonTarihBulRapor.Row
        If ThisWorkbook.Worksheets(4).Cells(i, 91).Value <> "" And ThisWorkbook.Worksheets(4).Cells(i, 174).Value = "No" Then
            'Gelen tema
            GelenTema = ""
            If ThisWorkbook.Worksheets(4).Cells(i, 33).Value = "Provincial Directorate B" Then
                If ThisWorkbook.Worksheets(4).Cells(i, 34).Value <> "" Then
                    GelenTema = ThisWorkbook.Worksheets(4).Cells(i, 25).Value & " Provincial Governorship Provincial Directorate B " & ThisWorkbook.Worksheets(4).Cells(i, 34).Value
                Else
                    GelenTema = ThisWorkbook.Worksheets(4).Cells(i, 25).Value & " Provincial Governorship Provincial Directorate B"
                End If
            ElseIf ThisWorkbook.Worksheets(4).Cells(i, 33).Value = "District Directorate B" Then
                If ThisWorkbook.Worksheets(4).Cells(i, 34).Value <> "" Then
                    GelenTema = ThisWorkbook.Worksheets(4).Cells(i, 26).Value & " District Governorship District Directorate B " & ThisWorkbook.Worksheets(4).Cells(i, 34).Value
                Else
                    GelenTema = ThisWorkbook.Worksheets(4).Cells(i, 26).Value & " District Governorship District Directorate B"
                End If
            ElseIf ThisWorkbook.Worksheets(4).Cells(i, 33).Value = "Provincial Directorate C" Then
                If ThisWorkbook.Worksheets(4).Cells(i, 34).Value <> "" Then
                    GelenTema = ThisWorkbook.Worksheets(4).Cells(i, 25).Value & " Provincial Governorship Provincial Directorate C " & ThisWorkbook.Worksheets(4).Cells(i, 34).Value
                Else
                    GelenTema = ThisWorkbook.Worksheets(4).Cells(i, 25).Value & " Provincial Governorship Provincial Directorate C"
                End If
            ElseIf ThisWorkbook.Worksheets(4).Cells(i, 33).Value = "District Directorate C" Then
                If ThisWorkbook.Worksheets(4).Cells(i, 34).Value <> "" Then
                    GelenTema = ThisWorkbook.Worksheets(4).Cells(i, 26).Value & " District Governorship District Directorate C " & ThisWorkbook.Worksheets(4).Cells(i, 34).Value
                Else
                    GelenTema = ThisWorkbook.Worksheets(4).Cells(i, 26).Value & " District Governorship District Directorate C"
                End If
            ElseIf ThisWorkbook.Worksheets(4).Cells(i, 33).Value = "Provincial Directorate D" Then
                If ThisWorkbook.Worksheets(4).Cells(i, 34).Value <> "" Then
                    GelenTema = ThisWorkbook.Worksheets(4).Cells(i, 25).Value & " Provincial Governorship Provincial Directorate D " & ThisWorkbook.Worksheets(4).Cells(i, 34).Value
                Else
                    GelenTema = ThisWorkbook.Worksheets(4).Cells(i, 25).Value & " Provincial Governorship Provincial Directorate D"
                End If
            ElseIf ThisWorkbook.Worksheets(4).Cells(i, 33).Value = "District Directorate D" Then
                If ThisWorkbook.Worksheets(4).Cells(i, 34).Value <> "" Then
                    GelenTema = ThisWorkbook.Worksheets(4).Cells(i, 26).Value & " District Governorship District Directorate D " & ThisWorkbook.Worksheets(4).Cells(i, 34).Value
                Else
                    GelenTema = ThisWorkbook.Worksheets(4).Cells(i, 26).Value & " District Governorship District Directorate D"
                End If
            ElseIf ThisWorkbook.Worksheets(4).Cells(i, 33).Value = "Provincial Directorate E" Then
                If ThisWorkbook.Worksheets(4).Cells(i, 34).Value <> "" Then
                    GelenTema = ThisWorkbook.Worksheets(4).Cells(i, 25).Value & " Provincial Governorship Provincial Directorate E " & ThisWorkbook.Worksheets(4).Cells(i, 34).Value
                Else
                    GelenTema = ThisWorkbook.Worksheets(4).Cells(i, 25).Value & " Provincial Governorship Provincial Directorate E"
                End If
            ElseIf ThisWorkbook.Worksheets(4).Cells(i, 33).Value = "District Directorate E" Then
                If ThisWorkbook.Worksheets(4).Cells(i, 34).Value <> "" Then
                    GelenTema = ThisWorkbook.Worksheets(4).Cells(i, 26).Value & " District Governorship District Directorate E " & ThisWorkbook.Worksheets(4).Cells(i, 34).Value
                Else
                    GelenTema = ThisWorkbook.Worksheets(4).Cells(i, 26).Value & " District Governorship District Directorate E"
                End If
            ElseIf InStr(ThisWorkbook.Worksheets(4).Cells(i, 33).Value, "General Directorate") <> 0 Or InStr(ThisWorkbook.Worksheets(4).Cells(i, 33).Value, "Regional Directorate") <> 0 Then
                If ThisWorkbook.Worksheets(4).Cells(i, 34).Value <> "" Then
                    GelenTema = WorksheetFunction.Proper(ThisWorkbook.Worksheets(2).Cells(6, 111).Value) & " " & ThisWorkbook.Worksheets(4).Cells(i, 33).Value & " " & ThisWorkbook.Worksheets(4).Cells(i, 34).Value
                Else
                    GelenTema = WorksheetFunction.Proper(ThisWorkbook.Worksheets(2).Cells(6, 111).Value) & " " & ThisWorkbook.Worksheets(4).Cells(i, 33).Value
                End If
            Else
                If ThisWorkbook.Worksheets(4).Cells(i, 34).Value <> "" Then
                    If InStr(ThisWorkbook.Worksheets(4).Cells(i, 33).Value, "X.X. ") > 0 Then
                        GelenTema = Mid(ThisWorkbook.Worksheets(4).Cells(i, 33).Value, 6, Len(ThisWorkbook.Worksheets(4).Cells(i, 33).Value)) & " " & ThisWorkbook.Worksheets(4).Cells(i, 34).Value
                    Else
                        GelenTema = ThisWorkbook.Worksheets(4).Cells(i, 33).Value & " " & ThisWorkbook.Worksheets(4).Cells(i, 34).Value
                    End If
                Else
                    If InStr(ThisWorkbook.Worksheets(4).Cells(i, 33).Value, "X.X. ") > 0 Then
                        GelenTema = Mid(ThisWorkbook.Worksheets(4).Cells(i, 33).Value, 6, Len(ThisWorkbook.Worksheets(4).Cells(i, 33).Value))
                    Else
                        GelenTema = ThisWorkbook.Worksheets(4).Cells(i, 33).Value
                    End If
                End If
            End If

            'Gönderen
            WsPerfRap.Cells(WsPerfRapSay, 3) = GelenTema
            'Rapor Tipi
            WsPerfRap.Cells(WsPerfRapSay, 4) = "Report 2.1"
            'Rapor Adedi
            AdetSay = 0
            For j = i To i + 50
                If ThisWorkbook.Worksheets(4).Cells(j, 92).Value = "" Then
                    If ThisWorkbook.Worksheets(4).Cells(j, 11).Value <> "" Then
                        AdetSay = AdetSay + 1
                        WsPerfRap.Cells(WsPerfRapSay, 5) = AdetSay
                    End If
                ElseIf ThisWorkbook.Worksheets(4).Cells(j, 92).Value <> "" Then
                    If ThisWorkbook.Worksheets(4).Cells(j, 11).Value <> "" Then
                        AdetSay = AdetSay + 1
                        WsPerfRap.Cells(WsPerfRapSay, 5) = AdetSay
                        GoTo jDongusuSonRapor
                    End If
                    GoTo jDongusuSonRapor
                End If
            Next j
jDongusuSonRapor:
            'Geliş Tarihi
            WsPerfRap.Cells(WsPerfRapSay, 6) = ThisWorkbook.Worksheets(4).Cells(i, 36)
            'Üst Yazı Tarihi
            WsPerfRap.Cells(WsPerfRapSay, 7) = ThisWorkbook.Worksheets(4).Cells(i, 83)
            'Sıra No
            WsPerfRap.Cells(WsPerfRapSay, 2) = WsPerfRapSay - 8
            
            WsPerfRapSay = WsPerfRapSay + 1
        End If
    Next i
End If

'Bilgilendirme Aktarımı
If RaporKont = 0 Then
    For i = IlkTarihBulRapor.Row To SonTarihBulRapor.Row
        If ThisWorkbook.Worksheets(4).Cells(i, 91).Value <> "" And ThisWorkbook.Worksheets(4).Cells(i, 174).Value = "Yes" Then
            'Gelen tema
            GelenTema = ""
            If ThisWorkbook.Worksheets(4).Cells(i, 33).Value = "Provincial Directorate B" Then
                If ThisWorkbook.Worksheets(4).Cells(i, 34).Value <> "" Then
                    GelenTema = ThisWorkbook.Worksheets(4).Cells(i, 25).Value & " Provincial Governorship Provincial Directorate B " & ThisWorkbook.Worksheets(4).Cells(i, 34).Value
                Else
                    GelenTema = ThisWorkbook.Worksheets(4).Cells(i, 25).Value & " Provincial Governorship Provincial Directorate B"
                End If
            ElseIf ThisWorkbook.Worksheets(4).Cells(i, 33).Value = "District Directorate B" Then
                If ThisWorkbook.Worksheets(4).Cells(i, 34).Value <> "" Then
                    GelenTema = ThisWorkbook.Worksheets(4).Cells(i, 26).Value & " District Governorship District Directorate B " & ThisWorkbook.Worksheets(4).Cells(i, 34).Value
                Else
                    GelenTema = ThisWorkbook.Worksheets(4).Cells(i, 26).Value & " District Governorship District Directorate B"
                End If
            ElseIf ThisWorkbook.Worksheets(4).Cells(i, 33).Value = "Provincial Directorate C" Then
                If ThisWorkbook.Worksheets(4).Cells(i, 34).Value <> "" Then
                    GelenTema = ThisWorkbook.Worksheets(4).Cells(i, 25).Value & " Provincial Governorship Provincial Directorate C " & ThisWorkbook.Worksheets(4).Cells(i, 34).Value
                Else
                    GelenTema = ThisWorkbook.Worksheets(4).Cells(i, 25).Value & " Provincial Governorship Provincial Directorate C"
                End If
            ElseIf ThisWorkbook.Worksheets(4).Cells(i, 33).Value = "District Directorate C" Then
                If ThisWorkbook.Worksheets(4).Cells(i, 34).Value <> "" Then
                    GelenTema = ThisWorkbook.Worksheets(4).Cells(i, 26).Value & " District Governorship District Directorate C " & ThisWorkbook.Worksheets(4).Cells(i, 34).Value
                Else
                    GelenTema = ThisWorkbook.Worksheets(4).Cells(i, 26).Value & " District Governorship District Directorate C"
                End If
            ElseIf ThisWorkbook.Worksheets(4).Cells(i, 33).Value = "Provincial Directorate D" Then
                If ThisWorkbook.Worksheets(4).Cells(i, 34).Value <> "" Then
                    GelenTema = ThisWorkbook.Worksheets(4).Cells(i, 25).Value & " Provincial Governorship Provincial Directorate D " & ThisWorkbook.Worksheets(4).Cells(i, 34).Value
                Else
                    GelenTema = ThisWorkbook.Worksheets(4).Cells(i, 25).Value & " Provincial Governorship Provincial Directorate D"
                End If
            ElseIf ThisWorkbook.Worksheets(4).Cells(i, 33).Value = "District Directorate D" Then
                If ThisWorkbook.Worksheets(4).Cells(i, 34).Value <> "" Then
                    GelenTema = ThisWorkbook.Worksheets(4).Cells(i, 26).Value & " District Governorship District Directorate D " & ThisWorkbook.Worksheets(4).Cells(i, 34).Value
                Else
                    GelenTema = ThisWorkbook.Worksheets(4).Cells(i, 26).Value & " District Governorship District Directorate D"
                End If
            ElseIf ThisWorkbook.Worksheets(4).Cells(i, 33).Value = "Provincial Directorate E" Then
                If ThisWorkbook.Worksheets(4).Cells(i, 34).Value <> "" Then
                    GelenTema = ThisWorkbook.Worksheets(4).Cells(i, 25).Value & " Provincial Governorship Provincial Directorate E " & ThisWorkbook.Worksheets(4).Cells(i, 34).Value
                Else
                    GelenTema = ThisWorkbook.Worksheets(4).Cells(i, 25).Value & " Provincial Governorship Provincial Directorate E"
                End If
            ElseIf ThisWorkbook.Worksheets(4).Cells(i, 33).Value = "District Directorate E" Then
                If ThisWorkbook.Worksheets(4).Cells(i, 34).Value <> "" Then
                    GelenTema = ThisWorkbook.Worksheets(4).Cells(i, 26).Value & " District Governorship District Directorate E " & ThisWorkbook.Worksheets(4).Cells(i, 34).Value
                Else
                    GelenTema = ThisWorkbook.Worksheets(4).Cells(i, 26).Value & " District Governorship District Directorate E"
                End If
            ElseIf InStr(ThisWorkbook.Worksheets(4).Cells(i, 33).Value, "General Directorate") <> 0 Or InStr(ThisWorkbook.Worksheets(4).Cells(i, 33).Value, "Regional Directorate") <> 0 Then
                If ThisWorkbook.Worksheets(4).Cells(i, 34).Value <> "" Then
                    GelenTema = WorksheetFunction.Proper(ThisWorkbook.Worksheets(2).Cells(6, 111).Value) & " " & ThisWorkbook.Worksheets(4).Cells(i, 33).Value & " " & ThisWorkbook.Worksheets(4).Cells(i, 34).Value
                Else
                    GelenTema = WorksheetFunction.Proper(ThisWorkbook.Worksheets(2).Cells(6, 111).Value) & " " & ThisWorkbook.Worksheets(4).Cells(i, 33).Value
                End If
            Else
                If ThisWorkbook.Worksheets(4).Cells(i, 34).Value <> "" Then
                    If InStr(ThisWorkbook.Worksheets(4).Cells(i, 33).Value, "X.X. ") > 0 Then
                        GelenTema = Mid(ThisWorkbook.Worksheets(4).Cells(i, 33).Value, 6, Len(ThisWorkbook.Worksheets(4).Cells(i, 33).Value)) & " " & ThisWorkbook.Worksheets(4).Cells(i, 34).Value
                    Else
                        GelenTema = ThisWorkbook.Worksheets(4).Cells(i, 33).Value & " " & ThisWorkbook.Worksheets(4).Cells(i, 34).Value
                    End If
                Else
                    If InStr(ThisWorkbook.Worksheets(4).Cells(i, 33).Value, "X.X. ") > 0 Then
                        GelenTema = Mid(ThisWorkbook.Worksheets(4).Cells(i, 33).Value, 6, Len(ThisWorkbook.Worksheets(4).Cells(i, 33).Value))
                    Else
                        GelenTema = ThisWorkbook.Worksheets(4).Cells(i, 33).Value
                    End If
                End If
            End If

            'Gönderen
            WsPerfRap.Cells(WsPerfRapSay, 3) = GelenTema & " (Information)"
            'Rapor Tipi
            WsPerfRap.Cells(WsPerfRapSay, 4) = "Report 2.1"
            'Rapor Adedi
            AdetSay = 0
            For j = i To i + 50
                If ThisWorkbook.Worksheets(4).Cells(j, 92).Value = "" Then
                    If ThisWorkbook.Worksheets(4).Cells(j, 11).Value <> "" Then
                        AdetSay = AdetSay + 1
                        WsPerfRap.Cells(WsPerfRapSay, 5) = AdetSay
                    End If
                ElseIf ThisWorkbook.Worksheets(4).Cells(j, 92).Value <> "" Then
                    If ThisWorkbook.Worksheets(4).Cells(j, 11).Value <> "" Then
                        AdetSay = AdetSay + 1
                        WsPerfRap.Cells(WsPerfRapSay, 5) = AdetSay
                        GoTo jDongusuSonRaporBilgilendirme
                    End If
                    GoTo jDongusuSonRaporBilgilendirme
                End If
            Next j
jDongusuSonRaporBilgilendirme:
            'Geliş Tarihi
            WsPerfRap.Cells(WsPerfRapSay, 6) = ThisWorkbook.Worksheets(4).Cells(i, 36)
            'Üst Yazı Tarihi
            WsPerfRap.Cells(WsPerfRapSay, 7) = ThisWorkbook.Worksheets(4).Cells(i, 83)
            'Sıra No
            WsPerfRap.Cells(WsPerfRapSay, 2) = WsPerfRapSay - 8
            
            WsPerfRapSay = WsPerfRapSay + 1
        End If
    Next i
End If

'XXXMud Aktarımı
If XXXMudKont = 0 Then
    For i = IlkTarihBulXXXMud.Row To SonTarihBulXXXMud.Row
        If ThisWorkbook.Worksheets(4).Cells(i, 91).Value <> "" And ThisWorkbook.Worksheets(4).Cells(i, 174).Value = "Yes" Then
            'Gelen tema
            GelenTema = ""
            If ThisWorkbook.Worksheets(4).Cells(i, 33).Value = "Provincial Directorate B" Then
                If ThisWorkbook.Worksheets(4).Cells(i, 34).Value <> "" Then
                    GelenTema = ThisWorkbook.Worksheets(4).Cells(i, 25).Value & " Provincial Governorship Provincial Directorate B " & ThisWorkbook.Worksheets(4).Cells(i, 34).Value
                Else
                    GelenTema = ThisWorkbook.Worksheets(4).Cells(i, 25).Value & " Provincial Governorship Provincial Directorate B"
                End If
            ElseIf ThisWorkbook.Worksheets(4).Cells(i, 33).Value = "District Directorate B" Then
                If ThisWorkbook.Worksheets(4).Cells(i, 34).Value <> "" Then
                    GelenTema = ThisWorkbook.Worksheets(4).Cells(i, 26).Value & " District Governorship District Directorate B " & ThisWorkbook.Worksheets(4).Cells(i, 34).Value
                Else
                    GelenTema = ThisWorkbook.Worksheets(4).Cells(i, 26).Value & " District Governorship District Directorate B"
                End If
            ElseIf ThisWorkbook.Worksheets(4).Cells(i, 33).Value = "Provincial Directorate C" Then
                If ThisWorkbook.Worksheets(4).Cells(i, 34).Value <> "" Then
                    GelenTema = ThisWorkbook.Worksheets(4).Cells(i, 25).Value & " Provincial Governorship Provincial Directorate C " & ThisWorkbook.Worksheets(4).Cells(i, 34).Value
                Else
                    GelenTema = ThisWorkbook.Worksheets(4).Cells(i, 25).Value & " Provincial Governorship Provincial Directorate C"
                End If
            ElseIf ThisWorkbook.Worksheets(4).Cells(i, 33).Value = "District Directorate C" Then
                If ThisWorkbook.Worksheets(4).Cells(i, 34).Value <> "" Then
                    GelenTema = ThisWorkbook.Worksheets(4).Cells(i, 26).Value & " District Governorship District Directorate C " & ThisWorkbook.Worksheets(4).Cells(i, 34).Value
                Else
                    GelenTema = ThisWorkbook.Worksheets(4).Cells(i, 26).Value & " District Governorship District Directorate C"
                End If
            ElseIf ThisWorkbook.Worksheets(4).Cells(i, 33).Value = "Provincial Directorate D" Then
                If ThisWorkbook.Worksheets(4).Cells(i, 34).Value <> "" Then
                    GelenTema = ThisWorkbook.Worksheets(4).Cells(i, 25).Value & " Provincial Governorship Provincial Directorate D " & ThisWorkbook.Worksheets(4).Cells(i, 34).Value
                Else
                    GelenTema = ThisWorkbook.Worksheets(4).Cells(i, 25).Value & " Provincial Governorship Provincial Directorate D"
                End If
            ElseIf ThisWorkbook.Worksheets(4).Cells(i, 33).Value = "District Directorate D" Then
                If ThisWorkbook.Worksheets(4).Cells(i, 34).Value <> "" Then
                    GelenTema = ThisWorkbook.Worksheets(4).Cells(i, 26).Value & " District Governorship District Directorate D " & ThisWorkbook.Worksheets(4).Cells(i, 34).Value
                Else
                    GelenTema = ThisWorkbook.Worksheets(4).Cells(i, 26).Value & " District Governorship District Directorate D"
                End If
            ElseIf ThisWorkbook.Worksheets(4).Cells(i, 33).Value = "Provincial Directorate E" Then
                If ThisWorkbook.Worksheets(4).Cells(i, 34).Value <> "" Then
                    GelenTema = ThisWorkbook.Worksheets(4).Cells(i, 25).Value & " Provincial Governorship Provincial Directorate E " & ThisWorkbook.Worksheets(4).Cells(i, 34).Value
                Else
                    GelenTema = ThisWorkbook.Worksheets(4).Cells(i, 25).Value & " Provincial Governorship Provincial Directorate E"
                End If
            ElseIf ThisWorkbook.Worksheets(4).Cells(i, 33).Value = "District Directorate E" Then
                If ThisWorkbook.Worksheets(4).Cells(i, 34).Value <> "" Then
                    GelenTema = ThisWorkbook.Worksheets(4).Cells(i, 26).Value & " District Governorship District Directorate E " & ThisWorkbook.Worksheets(4).Cells(i, 34).Value
                Else
                    GelenTema = ThisWorkbook.Worksheets(4).Cells(i, 26).Value & " District Governorship District Directorate E"
                End If
            ElseIf InStr(ThisWorkbook.Worksheets(4).Cells(i, 33).Value, "General Directorate") <> 0 Or InStr(ThisWorkbook.Worksheets(4).Cells(i, 33).Value, "Regional Directorate") <> 0 Then
                If ThisWorkbook.Worksheets(4).Cells(i, 34).Value <> "" Then
                    GelenTema = WorksheetFunction.Proper(ThisWorkbook.Worksheets(2).Cells(6, 111).Value) & " " & ThisWorkbook.Worksheets(4).Cells(i, 33).Value & " " & ThisWorkbook.Worksheets(4).Cells(i, 34).Value
                Else
                    GelenTema = WorksheetFunction.Proper(ThisWorkbook.Worksheets(2).Cells(6, 111).Value) & " " & ThisWorkbook.Worksheets(4).Cells(i, 33).Value
                End If
            Else
                If ThisWorkbook.Worksheets(4).Cells(i, 34).Value <> "" Then
                    If InStr(ThisWorkbook.Worksheets(4).Cells(i, 33).Value, "X.X. ") > 0 Then
                        GelenTema = Mid(ThisWorkbook.Worksheets(4).Cells(i, 33).Value, 6, Len(ThisWorkbook.Worksheets(4).Cells(i, 33).Value)) & " " & ThisWorkbook.Worksheets(4).Cells(i, 34).Value
                    Else
                        GelenTema = ThisWorkbook.Worksheets(4).Cells(i, 33).Value & " " & ThisWorkbook.Worksheets(4).Cells(i, 34).Value
                    End If
                Else
                    If InStr(ThisWorkbook.Worksheets(4).Cells(i, 33).Value, "X.X. ") > 0 Then
                        GelenTema = Mid(ThisWorkbook.Worksheets(4).Cells(i, 33).Value, 6, Len(ThisWorkbook.Worksheets(4).Cells(i, 33).Value))
                    Else
                        GelenTema = ThisWorkbook.Worksheets(4).Cells(i, 33).Value
                    End If
                End If
            End If


            'Gönderen
            WsPerfRap.Cells(WsPerfRapSay, 3) = GelenTema & " (Outgoing to XXX Directorate)"
            
            'Rapor Tipi
            WsPerfRap.Cells(WsPerfRapSay, 4) = "Report 2.2"
            'Rapor Adedi
            AdetSay = 1
            WsPerfRap.Cells(WsPerfRapSay, 5) = AdetSay
            'Geliş Tarihi
            WsPerfRap.Cells(WsPerfRapSay, 6) = ThisWorkbook.Worksheets(4).Cells(i, 36)
            'Üst Yazı Tarihi
            WsPerfRap.Cells(WsPerfRapSay, 7) = ThisWorkbook.Worksheets(4).Cells(i, 175)
            'Sıra No
            WsPerfRap.Cells(WsPerfRapSay, 2) = WsPerfRapSay - 8
            
            WsPerfRapSay = WsPerfRapSay + 1
        End If
    Next i
End If


'Sonuç Aktarımı
If SonucKont = 0 Then
    For i = IlkTarihBulSonuc.Row To SonTarihBulSonuc.Row
        If ThisWorkbook.Worksheets(4).Cells(i, 91).Value <> "" And ThisWorkbook.Worksheets(4).Cells(i, 174).Value = "Yes" Then
'            'Gelen tema
'            GelenTema = "ORGANIZATION A XXX Directorate"
            'Gelen tema
            GelenTema = ""
            If ThisWorkbook.Worksheets(4).Cells(i, 33).Value = "Provincial Directorate B" Then
                If ThisWorkbook.Worksheets(4).Cells(i, 34).Value <> "" Then
                    GelenTema = ThisWorkbook.Worksheets(4).Cells(i, 25).Value & " Provincial Governorship Provincial Directorate B " & ThisWorkbook.Worksheets(4).Cells(i, 34).Value
                Else
                    GelenTema = ThisWorkbook.Worksheets(4).Cells(i, 25).Value & " Provincial Governorship Provincial Directorate B"
                End If
            ElseIf ThisWorkbook.Worksheets(4).Cells(i, 33).Value = "District Directorate B" Then
                If ThisWorkbook.Worksheets(4).Cells(i, 34).Value <> "" Then
                    GelenTema = ThisWorkbook.Worksheets(4).Cells(i, 26).Value & " District Governorship District Directorate B " & ThisWorkbook.Worksheets(4).Cells(i, 34).Value
                Else
                    GelenTema = ThisWorkbook.Worksheets(4).Cells(i, 26).Value & " District Governorship District Directorate B"
                End If
            ElseIf ThisWorkbook.Worksheets(4).Cells(i, 33).Value = "Provincial Directorate C" Then
                If ThisWorkbook.Worksheets(4).Cells(i, 34).Value <> "" Then
                    GelenTema = ThisWorkbook.Worksheets(4).Cells(i, 25).Value & " Provincial Governorship Provincial Directorate C " & ThisWorkbook.Worksheets(4).Cells(i, 34).Value
                Else
                    GelenTema = ThisWorkbook.Worksheets(4).Cells(i, 25).Value & " Provincial Governorship Provincial Directorate C"
                End If
            ElseIf ThisWorkbook.Worksheets(4).Cells(i, 33).Value = "District Directorate C" Then
                If ThisWorkbook.Worksheets(4).Cells(i, 34).Value <> "" Then
                    GelenTema = ThisWorkbook.Worksheets(4).Cells(i, 26).Value & " District Governorship District Directorate C " & ThisWorkbook.Worksheets(4).Cells(i, 34).Value
                Else
                    GelenTema = ThisWorkbook.Worksheets(4).Cells(i, 26).Value & " District Governorship District Directorate C"
                End If
            ElseIf ThisWorkbook.Worksheets(4).Cells(i, 33).Value = "Provincial Directorate D" Then
                If ThisWorkbook.Worksheets(4).Cells(i, 34).Value <> "" Then
                    GelenTema = ThisWorkbook.Worksheets(4).Cells(i, 25).Value & " Provincial Governorship Provincial Directorate D " & ThisWorkbook.Worksheets(4).Cells(i, 34).Value
                Else
                    GelenTema = ThisWorkbook.Worksheets(4).Cells(i, 25).Value & " Provincial Governorship Provincial Directorate D"
                End If
            ElseIf ThisWorkbook.Worksheets(4).Cells(i, 33).Value = "District Directorate D" Then
                If ThisWorkbook.Worksheets(4).Cells(i, 34).Value <> "" Then
                    GelenTema = ThisWorkbook.Worksheets(4).Cells(i, 26).Value & " District Governorship District Directorate D " & ThisWorkbook.Worksheets(4).Cells(i, 34).Value
                Else
                    GelenTema = ThisWorkbook.Worksheets(4).Cells(i, 26).Value & " District Governorship District Directorate D"
                End If
            ElseIf ThisWorkbook.Worksheets(4).Cells(i, 33).Value = "Provincial Directorate E" Then
                If ThisWorkbook.Worksheets(4).Cells(i, 34).Value <> "" Then
                    GelenTema = ThisWorkbook.Worksheets(4).Cells(i, 25).Value & " Provincial Governorship Provincial Directorate E " & ThisWorkbook.Worksheets(4).Cells(i, 34).Value
                Else
                    GelenTema = ThisWorkbook.Worksheets(4).Cells(i, 25).Value & " Provincial Governorship Provincial Directorate E"
                End If
            ElseIf ThisWorkbook.Worksheets(4).Cells(i, 33).Value = "District Directorate E" Then
                If ThisWorkbook.Worksheets(4).Cells(i, 34).Value <> "" Then
                    GelenTema = ThisWorkbook.Worksheets(4).Cells(i, 26).Value & " District Governorship District Directorate E " & ThisWorkbook.Worksheets(4).Cells(i, 34).Value
                Else
                    GelenTema = ThisWorkbook.Worksheets(4).Cells(i, 26).Value & " District Governorship District Directorate E"
                End If
            ElseIf InStr(ThisWorkbook.Worksheets(4).Cells(i, 33).Value, "General Directorate") <> 0 Or InStr(ThisWorkbook.Worksheets(4).Cells(i, 33).Value, "Regional Directorate") <> 0 Then
                If ThisWorkbook.Worksheets(4).Cells(i, 34).Value <> "" Then
                    GelenTema = WorksheetFunction.Proper(ThisWorkbook.Worksheets(2).Cells(6, 111).Value) & " " & ThisWorkbook.Worksheets(4).Cells(i, 33).Value & " " & ThisWorkbook.Worksheets(4).Cells(i, 34).Value
                Else
                    GelenTema = WorksheetFunction.Proper(ThisWorkbook.Worksheets(2).Cells(6, 111).Value) & " " & ThisWorkbook.Worksheets(4).Cells(i, 33).Value
                End If
            Else
                If ThisWorkbook.Worksheets(4).Cells(i, 34).Value <> "" Then
                    If InStr(ThisWorkbook.Worksheets(4).Cells(i, 33).Value, "X.X. ") > 0 Then
                        GelenTema = Mid(ThisWorkbook.Worksheets(4).Cells(i, 33).Value, 6, Len(ThisWorkbook.Worksheets(4).Cells(i, 33).Value)) & " " & ThisWorkbook.Worksheets(4).Cells(i, 34).Value
                    Else
                        GelenTema = ThisWorkbook.Worksheets(4).Cells(i, 33).Value & " " & ThisWorkbook.Worksheets(4).Cells(i, 34).Value
                    End If
                Else
                    If InStr(ThisWorkbook.Worksheets(4).Cells(i, 33).Value, "X.X. ") > 0 Then
                        GelenTema = Mid(ThisWorkbook.Worksheets(4).Cells(i, 33).Value, 6, Len(ThisWorkbook.Worksheets(4).Cells(i, 33).Value))
                    Else
                        GelenTema = ThisWorkbook.Worksheets(4).Cells(i, 33).Value
                    End If
                End If
            End If
            'Gönderen
            WsPerfRap.Cells(WsPerfRapSay, 3) = GelenTema & " (Final)"

            'Rapor Tipi
            WsPerfRap.Cells(WsPerfRapSay, 4) = "Report 2.2"
'            'Rapor Adedi
'            'Rapor2_2 rapor numaralarını say
'            StrContent = ThisWorkbook.Worksheets(4).Cells(i, 180).Value
'            If StrContent <> "" Then
'                AdetSay = 1
'                For j = 1 To Len(StrContent)
'                    If Mid(StrContent, j, 1) = " " Then
'                        AdetSay = AdetSay + 1
'                    End If
'                Next j
'            End If
            AdetSay = 1
            WsPerfRap.Cells(WsPerfRapSay, 5) = AdetSay
            
            'Geliş Tarihi
            WsPerfRap.Cells(WsPerfRapSay, 6) = ThisWorkbook.Worksheets(4).Cells(i, 191)
            'Üst Yazı Tarihi
            WsPerfRap.Cells(WsPerfRapSay, 7) = ThisWorkbook.Worksheets(4).Cells(i, 177)
            'Sıra No
            WsPerfRap.Cells(WsPerfRapSay, 2) = WsPerfRapSay - 8
            
            WsPerfRapSay = WsPerfRapSay + 1
        End If
    Next i
End If


Unload Me

Son:

'Kenarlıklar.
If WsPerfRapSay = 9 Then
    WsPerfRapSay = 10
End If
Set Kenarlar = WsPerfRap.Range("B" & 9 & ":G" & WsPerfRapSay - 1)
Kenarlar.Borders(xlDiagonalDown).LineStyle = xlNone
Kenarlar.Borders(xlDiagonalUp).LineStyle = xlNone
'With Kenarlar.Borders(xlEdgeLeft)
'    .LineStyle = xlContinuous
'    .Color = RGB(217, 217, 217)
'    .TintAndShade = 0
'    .Weight = xlThin
'End With
With Kenarlar.Borders(xlEdgeTop)
    .LineStyle = xlContinuous
    .Color = RGB(217, 217, 217)
    .TintAndShade = 0
    .Weight = xlThin
End With
With Kenarlar.Borders(xlEdgeBottom)
    .LineStyle = xlContinuous
    .Color = RGB(217, 217, 217)
    .TintAndShade = 0
    .Weight = xlThin
End With
'With Kenarlar.Borders(xlEdgeRight)
'    .LineStyle = xlContinuous
'    .Color = RGB(217, 217, 217)
'    .TintAndShade = 0
'    .Weight = xlThin
'End With
With Kenarlar.Borders(xlInsideVertical)
    .LineStyle = xlContinuous
    .Color = RGB(217, 217, 217)
    .TintAndShade = 0
    .Weight = xlThin
End With
With Kenarlar.Borders(xlInsideHorizontal)
    .LineStyle = xlContinuous
    .Color = RGB(217, 217, 217)
    .TintAndShade = 0
    .Weight = xlThin
End With

Workbooks("Performance Report.xlsx").Save


Out:

Application.DisplayAlerts = True
Application.EnableEvents = True
Application.ScreenUpdating = True

End Sub

Private Sub FirstDateText_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'suppport_calendar_UI
suppport_calendar_UI.Show

If CalTarih = "" Then
    GoTo Son
Else
    FirstDateText.Value = CalTarih
    FirstDateText.Value = Format(FirstDateText.Value, "dd.mm.yyyy")
End If

Son:
CalTarih = ""
Cancel = True

End Sub

Private Sub FirstDateText_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    
    On Error Resume Next
    'Delete ve Backspace tuşları textboxu sil.
    If KeyCode = vbKeyDelete Then
        FirstDateText.Value = ""
    End If
    If KeyCode = vbKeyBack Then
        FirstDateText.Value = ""
    End If

End Sub

Private Sub FirstDateLabel_Click()

'suppport_calendar_UI
suppport_calendar_UI.Show

If CalTarih = "" Then
    GoTo Son
Else
    FirstDateText.Value = CalTarih
    FirstDateText.Value = Format(FirstDateText.Value, "dd.mm.yyyy")
End If

Son:
CalTarih = ""

End Sub

Private Sub LastDateText_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'suppport_calendar_UI
suppport_calendar_UI.Show

If CalTarih = "" Then
    GoTo Son
Else
    LastDateText.Value = CalTarih
    LastDateText.Value = Format(LastDateText.Value, "dd.mm.yyyy")
End If

Son:
CalTarih = ""
Cancel = True

End Sub

Private Sub LastDateText_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    
    On Error Resume Next
    'Delete ve Backspace tuşları textboxu sil.
    If KeyCode = vbKeyDelete Then
        LastDateText.Value = ""
    End If
    If KeyCode = vbKeyBack Then
        LastDateText.Value = ""
    End If

End Sub

Private Sub LastDateLabel_Click()

'suppport_calendar_UI
suppport_calendar_UI.Show

If CalTarih = "" Then
    GoTo Son
Else
    LastDateText.Value = CalTarih
    LastDateText.Value = Format(LastDateText.Value, "dd.mm.yyyy")
End If

Son:
CalTarih = ""

End Sub

Private Sub UserForm_Initialize()
Dim i As Long, WsSKP As Object
Dim ClrLab As MSForms.Control


ThisWorkbook.Activate


For Each ClrLab In core_performance_report_UI.Controls
    If TypeName(ClrLab) = "Label" Then
        ClrLab.BackColor = RGB(254, 254, 254)
        ClrLab.ForeColor = RGB(70, 70, 70)
    End If
    If TypeName(ClrLab) = "CheckBox" Then
        ClrLab.BackColor = RGB(254, 254, 254)
        ClrLab.ForeColor = RGB(70, 70, 70)
    End If
    If TypeName(ClrLab) = "OptionButton" Then
        ClrLab.BackColor = RGB(254, 254, 254)
        ClrLab.ForeColor = RGB(70, 70, 70)
    End If
    If TypeName(ClrLab) = "TextBox" Then
        ClrLab.ForeColor = RGB(30, 30, 30)
    End If
    If TypeName(ClrLab) = "ComboBox" Then
        ClrLab.ForeColor = RGB(30, 30, 30)
    End If
    
    'YENİ
    If TypeName(ClrLab) = "Frame" Then
        ClrLab.BackColor = RGB(254, 254, 254)
        ClrLab.ForeColor = RGB(30, 30, 30)
        ClrLab.BorderColor = RGB(180, 180, 180)
    End If
Next ClrLab

UstMenuFrame.BackColor = RGB(225, 235, 245) 'YENİ
AltMenuFrame.BackColor = RGB(225, 235, 245) 'YENİ
LblBilgilendirme.BackColor = RGB(254, 254, 254)

Raporla.BackColor = RGB(225, 235, 245)
Raporla.ForeColor = RGB(30, 30, 30)
Kapat.BackColor = RGB(225, 235, 245)
Kapat.ForeColor = RGB(30, 30, 30)
Yardim.BackColor = RGB(225, 235, 245)
Yardim.ForeColor = RGB(30, 30, 30)
core_performance_report_UI.BackColor = RGB(230, 230, 230) 'YENİ


End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
Dim yukseklik As Variant, genislik As Variant
Dim Rep As Variant

yukseklik = Me.Height
Rep = Me.Width
Do
DoEvents
Rep = Rep - 60
Call timeout(0.01)
    If Rep > 60 Then
        core_performance_report_UI.Width = Rep
        yukseklik = yukseklik - 60
        core_performance_report_UI.Height = yukseklik
        If yukseklik <= 60 Then
            yukseklik = 60
            core_performance_report_UI.Height = yukseklik
        End If
    ElseIf Rep <= 50 Then
        Rep = 50
        core_performance_report_UI.Width = Rep
        yukseklik = yukseklik - 50
        core_performance_report_UI.Height = yukseklik
        If yukseklik <= 50 Then
            yukseklik = 50
            core_performance_report_UI.Height = yukseklik
        End If
    End If
Loop Until yukseklik = 50

Unload Me

End Sub

Sub timeout(duration_ms As Double)
Dim Start_Time As Variant

Start_Time = Timer
Do
DoEvents
Loop Until (Timer - Start_Time) >= duration_ms

End Sub




