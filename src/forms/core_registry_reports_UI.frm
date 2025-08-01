VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} core_registry_reports_UI 
   Caption         =   "Enterprise Document Automation System"
   ClientHeight    =   8895
   ClientLeft      =   30
   ClientTop       =   375
   ClientWidth     =   13920
   OleObjectBlob   =   "core_registry_reports_UI.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "core_registry_reports_UI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Option Compare Text
Dim StrRaporTarihiAylik As String

Sub ColorChangerGenel()

'RaporlaGunluk
If RaporlaGunluk.BackColor <> RGB(225, 235, 245) Then
    RaporlaGunluk.BackColor = RGB(225, 235, 245)
    RaporlaGunluk.ForeColor = RGB(30, 30, 30)
End If
'RaporlaAylik
If RaporlaAylik.BackColor <> RGB(225, 235, 245) Then
    RaporlaAylik.BackColor = RGB(225, 235, 245)
    RaporlaAylik.ForeColor = RGB(30, 30, 30)
End If
'Kapat
If Kapat.BackColor <> RGB(225, 235, 245) Then
    Kapat.BackColor = RGB(225, 235, 245)
    Kapat.ForeColor = RGB(30, 30, 30)
End If

'CheckBoxRapor1
If CheckBoxRapor1.BackColor <> RGB(254, 254, 254) Then
    CheckBoxRapor1.BackColor = RGB(254, 254, 254)
    CheckBoxRapor1.ForeColor = RGB(70, 70, 70)
End If

'CheckBoxRapor
If CheckBoxRapor.BackColor <> RGB(254, 254, 254) Then
    CheckBoxRapor.BackColor = RGB(254, 254, 254)
    CheckBoxRapor.ForeColor = RGB(70, 70, 70)
End If

'CheckBoxRapor2_2
If CheckBoxRapor2_2.BackColor <> RGB(254, 254, 254) Then
    CheckBoxRapor2_2.BackColor = RGB(254, 254, 254)
    CheckBoxRapor2_2.ForeColor = RGB(70, 70, 70)
End If

If GunlukTarihiLabel.BackColor <> RGB(254, 254, 254) Then
    GunlukTarihiLabel.BackColor = RGB(254, 254, 254)
    GunlukTarihiLabel.ForeColor = RGB(70, 70, 70)
End If


'CheckBoxRapor1Aylik
If CheckBoxRapor1Aylik.BackColor <> RGB(254, 254, 254) Then
    CheckBoxRapor1Aylik.BackColor = RGB(254, 254, 254)
    CheckBoxRapor1Aylik.ForeColor = RGB(70, 70, 70)
End If

'CheckBoxRaporAylik
If CheckBoxRaporAylik.BackColor <> RGB(254, 254, 254) Then
    CheckBoxRaporAylik.BackColor = RGB(254, 254, 254)
    CheckBoxRaporAylik.ForeColor = RGB(70, 70, 70)
End If

If AylikTarihiLabel.BackColor <> RGB(254, 254, 254) Then
    AylikTarihiLabel.BackColor = RGB(254, 254, 254)
    AylikTarihiLabel.ForeColor = RGB(70, 70, 70)
End If


End Sub

Private Sub RaporlaGunluk_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
RaporlaGunluk.BackColor = RGB(60, 100, 180) 'RGB(60, 100, 180) 'RGB(60, 100, 180)
RaporlaGunluk.ForeColor = RGB(255, 255, 255)
End Sub

Private Sub RaporlaAylik_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
RaporlaAylik.BackColor = RGB(60, 100, 180) 'RGB(60, 100, 180) 'RGB(60, 100, 180)
RaporlaAylik.ForeColor = RGB(255, 255, 255)
End Sub


Private Sub Kapat_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
Kapat.BackColor = RGB(60, 100, 180) 'RGB(60, 100, 180) 'RGB(60, 100, 180)
Kapat.ForeColor = RGB(255, 255, 255)
End Sub


Private Sub CheckBoxRapor1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
CheckBoxRapor1.BackColor = RGB(60, 100, 180)
CheckBoxRapor1.ForeColor = RGB(255, 255, 255)
End Sub

Private Sub CheckBoxRapor_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
CheckBoxRapor.BackColor = RGB(60, 100, 180)
CheckBoxRapor.ForeColor = RGB(255, 255, 255)
End Sub

Private Sub CheckBoxRapor2_2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
CheckBoxRapor2_2.BackColor = RGB(60, 100, 180)
CheckBoxRapor2_2.ForeColor = RGB(255, 255, 255)
End Sub

Private Sub GunlukTarihiLabel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
GunlukTarihiLabel.BackColor = RGB(60, 100, 180) 'RGB(60, 100, 180) 'RGB(60, 100, 180)
GunlukTarihiLabel.ForeColor = RGB(255, 255, 255)
End Sub
Private Sub GunlukTarihi_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub

Private Sub CheckBoxRapor1Aylik_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
CheckBoxRapor1Aylik.BackColor = RGB(60, 100, 180)
CheckBoxRapor1Aylik.ForeColor = RGB(255, 255, 255)
End Sub

Private Sub CheckBoxRaporAylik_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
CheckBoxRaporAylik.BackColor = RGB(60, 100, 180)
CheckBoxRaporAylik.ForeColor = RGB(255, 255, 255)
End Sub

Private Sub AylikTarihiLabel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
AylikTarihiLabel.BackColor = RGB(60, 100, 180) 'RGB(60, 100, 180) 'RGB(60, 100, 180)
AylikTarihiLabel.ForeColor = RGB(255, 255, 255)
End Sub
Private Sub AylikTarihi_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
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
Private Sub AylikFrame_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub
Private Sub GunlukFrame_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub

Private Sub LblBilgilendirme_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub
Private Sub AltMenuFrame_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub

Private Sub GunlukTarihi_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'suppport_calendar_UI
suppport_calendar_UI.Show

If CalTarih = "" Then
    GoTo Son
Else
    GunlukTarihi.Value = CalTarih
    GunlukTarihi.Value = Format(GunlukTarihi.Value, "dd.mm.yyyy")
End If

Son:
CalTarih = ""
Cancel = True

End Sub

Private Sub GunlukTarihi_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    
    On Error Resume Next
    'Delete ve Backspace tuşları textboxu sil.
    If KeyCode = vbKeyDelete Then
        GunlukTarihi.Value = ""
    End If
    If KeyCode = vbKeyBack Then
        GunlukTarihi.Value = ""
    End If

End Sub

Private Sub GunlukTarihiLabel_Click()

'suppport_calendar_UI
suppport_calendar_UI.Show

If CalTarih = "" Then
    GoTo Son
Else
    GunlukTarihi.Value = CalTarih
    GunlukTarihi.Value = Format(GunlukTarihi.Value, "dd.mm.yyyy")
End If

Son:
CalTarih = ""

End Sub

Private Sub AylikTarihi_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'suppport_calendar_UI
suppport_calendar_UI.Show

If CalTarih = "" Then
    GoTo Son
Else
    AylikTarihi.Value = CalTarih
    AylikTarihi.Value = Format(AylikTarihi.Value, "dd.mm.yyyy")
End If

Son:
CalTarih = ""
Cancel = True

End Sub

Private Sub AylikTarihi_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    
    On Error Resume Next
    'Delete ve Backspace tuşları textboxu sil.
    If KeyCode = vbKeyDelete Then
        AylikTarihi.Value = ""
    End If
    If KeyCode = vbKeyBack Then
        AylikTarihi.Value = ""
    End If

End Sub

Private Sub AylikTarihiLabel_Click()

'suppport_calendar_UI
suppport_calendar_UI.Show

If CalTarih = "" Then
    GoTo Son
Else
    AylikTarihi.Value = CalTarih
    AylikTarihi.Value = Format(AylikTarihi.Value, "dd.mm.yyyy")
End If

Son:
CalTarih = ""

End Sub

Private Sub RaporlaAylik_Click()
Dim AutoPath As String, IslemGunlukleriKlasor As String, OpenControl As String, KontrolFile As String
Dim i As Long, j As Long, ContSay As Long

Dim SourceRapor1 As String, SourceRapor As String, SourceRapor2_2 As String
Dim StrRapor1 As String, StrRapor As String, StrRapor2_2 As String
Dim Rapor1Op As String, RaporOp As String, Rapor2_2Op As String
Dim WsRapor1 As Worksheet, WsRapor As Worksheet, WsRapor2_2 As Worksheet

Dim SourceSysRapor1 As String, SourceSysRapor As String, SourceSysRapor2_2 As String
Dim SistemRapor1 As String, SistemRapor As String, SistemRapor2_2 As String
Dim WsSysRapor1 As Worksheet, WsSysRapor As Worksheet, WsSysRapor2_2 As Worksheet

Dim DestOperasyon As String, DestOpUserFolderName As String, DestOpUserFolder As String
Dim fso As Object, StrRaporTarihi As String, Kenarlar As Range

Dim IlkSira As Long, SonSira As Long, SiraNo As Long
Dim BulIslemGunlugu As Range, BulIslemGunlugux As Range


Application.DisplayAlerts = False
Application.EnableEvents = False
Application.ScreenUpdating = False


If AylikTarihi.Value <> "" Then
    If CheckBoxRapor1Aylik.Value = False And CheckBoxRaporAylik.Value = False Then
        MsgBox "Please specify the type of daily operation log you want to generate a MONTHLY report for.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Out
    End If
End If

If CheckBoxRapor1Aylik.Value = True Or CheckBoxRaporAylik.Value = True Then
    If AylikTarihi.Value = "" Then
        MsgBox "Please specify a date within the period for the MONTHLY report you want to generate. For example, to create the report for December 2021, you may select any date within December such as December 1, December 7, or December 31.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Out
    End If
End If

If AylikTarihi.Value <> "" Then
    If CheckBoxRapor1Aylik.Value = True Or CheckBoxRaporAylik.Value = True Then
        StrRaporTarihiAylik = AylikTarihi.Value
    Else
        GoTo Out
    End If
Else
    GoTo Out
End If


'____________


'Pathfinder...
AutoPath = ThisWorkbook.Path
DestOperasyon = AutoPath & "\System Files\Operation\"

IslemGunlukleriKlasor = AutoPath & "\System Files\System Templates\Registry Reports\"

StrRapor1 = "Registry Report 1.xlsx"
StrRapor = "Registry Report 2.1.xlsx"
'StrRapor2_2 = "Registry Report 2.2.xlsx"
SourceRapor1 = IslemGunlukleriKlasor & StrRapor1
SourceRapor = IslemGunlukleriKlasor & StrRapor
'SourceRapor2_2 = IslemGunlukleriKlasor & StrRapor2_2

SistemRapor1 = "System Registry Report 1.xlsx"
SistemRapor = "System Registry Report 2.1.xlsx"
'SistemRapor2_2 = "System Registry Report 2.2.xlsx"
SourceSysRapor1 = IslemGunlukleriKlasor & SistemRapor1
SourceSysRapor = IslemGunlukleriKlasor & SistemRapor
'SourceSysRapor2_2 = IslemGunlukleriKlasor & SistemRapor2_2


'Birden fazla kullanıcının operasyon klasöründe çakışmasını önlemek için operasyon klasörünün içinde yeni klasör aç.
DestOpUserFolderName = "Operation-" & Left(ThisWorkbook.name, InStr(ThisWorkbook.name, ".") - 1) '& " " & Format(Now(), "ddmmyyyyhhmmss")
DestOpUserFolder = DestOperasyon & DestOpUserFolderName & "\"

'Check Operation folder name.
If Not Dir(DestOperasyon, vbDirectory) <> vbNullString Then
    MsgBox "Cannot access directory: " & DestOperasyon & ". The folder named 'Operation' may have been renamed or deleted.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Out
End If

'Operation klasörü içinde kullanıcı modülü klasörü yoksa oluştur.
If Not Dir(DestOpUserFolder, vbDirectory) <> vbNullString Then
    MkDir DestOpUserFolder
End If

Rapor1Op = DestOpUserFolder & StrRapor1
RaporOp = DestOpUserFolder & StrRapor
'Rapor2_2Op = DestOpUserFolder & StrRapor2_2


'Check Operation Logs folder name.
If Not Dir(IslemGunlukleriKlasor, vbDirectory) <> vbNullString Then
    MsgBox "Cannot access directory: " & IslemGunlukleriKlasor & ". The folder named 'Registry Reports' may have been renamed or deleted.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Out
End If

If Not Dir(SourceRapor1, vbDirectory) <> vbNullString Then
    MsgBox "Cannot access directory: " & SourceRapor1 & ". Folder or file names within this directory may have been changed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Out
End If

If Not Dir(SourceRapor, vbDirectory) <> vbNullString Then
    MsgBox "Cannot access directory: " & SourceRapor & ". Folder or file names within this directory may have been changed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Out
End If

'If Not Dir(SourceRapor2_2, vbDirectory) <> vbNullString Then
'    MsgBox "Cannot access directory: " & SourceRapor2_2 & ". Folder or file names within this directory may have been changed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
'    GoTo Out
'End If

If Not Dir(SourceSysRapor1, vbDirectory) <> vbNullString Then
    MsgBox "Cannot access directory: " & SourceSysRapor1 & ". Folder or file names within this directory may have been changed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Out
End If

If Not Dir(SourceSysRapor, vbDirectory) <> vbNullString Then
    MsgBox "Cannot access directory: " & SourceSysRapor & ". Folder or file names within this directory may have been changed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Out
End If

'If Not Dir(SourceSysRapor2_2, vbDirectory) <> vbNullString Then
'    MsgBox "Cannot access directory: " & SourceSysRapor2_2 & ". Folder or file names within this directory may have been changed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
'    GoTo Out
'End If



On Error Resume Next 'Operation içinde .xlsx dosyası yoksa oluşacak hata için
'İşlem günlüğü açıksa kaydet ve kapat.
OpenControl = IsWorkBookOpen(Rapor1Op)
If OpenControl = True Then
    Workbooks(StrRapor1).Close SaveChanges:=False
End If
'İşlem günlüğü açıksa kaydet ve kapat.
OpenControl = IsWorkBookOpen(RaporOp)
If OpenControl = True Then
    Workbooks(StrRapor).Close SaveChanges:=False
End If
''İşlem günlüğü açıksa kaydet ve kapat.
'OpenControl = IsWorkBookOpen(Rapor2_2Op)
'If OpenControl = True Then
'    Workbooks(StrRapor2_2).Close SaveChanges:=False
'End If

OpenControl = IsWorkBookOpen(SourceSysRapor1)
If OpenControl = True Then
    Workbooks(SistemRapor1).Close SaveChanges:=False
End If
'İşlem günlüğü açıksa kaydet ve kapat.
OpenControl = IsWorkBookOpen(SourceSysRapor)
If OpenControl = True Then
    Workbooks(SistemRapor).Close SaveChanges:=False
End If
''İşlem günlüğü açıksa kaydet ve kapat.
'OpenControl = IsWorkBookOpen(SourceSysRapor2_2)
'If OpenControl = True Then
'    Workbooks(SistemRapor2_2).Close SaveChanges:=False
'End If


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

On Error GoTo 0

'Dosyayı operasyon klasörüne kopyala ve adını değiştir.
Set fso = CreateObject("Scripting.FileSystemObject")
fso.CopyFile (SourceRapor1), DestOpUserFolder & StrRapor1, True
fso.CopyFile (SourceRapor), DestOpUserFolder & StrRapor, True
'fso.CopyFile (SourceRapor2_2), DestOpUserFolder & StrRapor2_2, True

'open the file
If CheckBoxRapor1Aylik.Value = True Then
    Workbooks.Open (SourceSysRapor1)
    Set WsSysRapor1 = Workbooks(SistemRapor1).Worksheets(1)
    WsSysRapor1.Unprotect Password:="123"
    WsSysRapor1.Columns("B:C").EntireColumn.Hidden = False

    Workbooks.Open (Rapor1Op)
    Set WsRapor1 = Workbooks(StrRapor1).Worksheets(1)

    'İşlem günlüğünde başlangıç ve bitiş satırlarını tespit et.
    'Say1IslemGunlugu = WsSysRapor1.Range("B100000").End(xlUp).Row
    Say2IslemGunlugu = WsSysRapor1.Range("C100000").End(xlUp).Row
    If Say2IslemGunlugu < 8 Then
        Say2IslemGunlugu = 8
    End If
    'SayAyracIslemGunlugu = WsSysRapor1.Range("E100000").End(xlUp).Row

    'Ayracı oluştur
    ModulTarih = StrRaporTarihiAylik
    ModulAyrac = "01" & Right(ModulTarih, 8)
    SiraNo = 0
    
    Set BulIslemGunlugu = WsSysRapor1.Range("E:E").Find(What:=CDate(ModulAyrac), SearchDirection:=xlNext, _
            SearchOrder:=xlByRows, LookIn:=xlFormulas, LookAt:=xlWhole)
    If Not BulIslemGunlugu Is Nothing Then 'HEDEF DÖNEMİ bul
        ilkrow = BulIslemGunlugu.Row + 1
        ModulAyrac = DateAdd("m", 1, CDate(ModulAyrac)) 'Sonraki dönemi bul ve onun üstüne satır ekle
        Set BulIslemGunlugux = WsSysRapor1.Range("E:E").Find(What:=CDate(ModulAyrac), SearchDirection:=xlNext, _
                SearchOrder:=xlByRows, LookIn:=xlFormulas, LookAt:=xlWhole)
        If Not BulIslemGunlugux Is Nothing Then
            sonrow = BulIslemGunlugux.Row - 1
        Else
            sonrow = Say2IslemGunlugu
        End If
        ModulAyrac = DateAdd("m", -1, CDate(ModulAyrac)) 'Yukarıda atadığın +1 ayı geri al
        
        If WsSysRapor1.Cells(ilkrow, 6).Value <> "" Then
            SiraNo = 1
            WsRapor1.Range(WsRapor1.Cells(2, 1), WsRapor1.Cells(2 + sonrow - ilkrow, 14)).Value = WsSysRapor1.Range(WsSysRapor1.Cells(ilkrow, 6), WsSysRapor1.Cells(sonrow, 19)).Value
        Else
            '
        End If
    Else
        '
    End If


    'Kenarlıklar.
    If SiraNo > 0 Then
        Set Kenarlar = WsRapor1.Range("A" & 1 & ":N" & 2 + sonrow - ilkrow)
        Kenarlar.Borders.LineStyle = xlNone
        Kenarlar.Borders(xlDiagonalDown).LineStyle = xlNone
        Kenarlar.Borders(xlDiagonalUp).LineStyle = xlNone
        Kenarlar.Borders.LineStyle = xlContinuous
    End If
    With WsRapor1.PageSetup
        .CenterFooter = "&""Open Sans""&18" & Format(CDate(ModulAyrac), "mmmm yyyy")
    End With



    OpenControl = IsWorkBookOpen(Rapor1Op)
    If OpenControl = True Then
        Workbooks(StrRapor1).Save
    End If
    
    '_________________
    
    WsSysRapor1.Columns("B:C").EntireColumn.Hidden = True
    WsSysRapor1.Protect Password:="123"
    OpenControl = IsWorkBookOpen(SourceSysRapor1)
    If OpenControl = True Then
        Workbooks(SistemRapor1).Close SaveChanges:=True
    End If
    
End If

If CheckBoxRaporAylik.Value = True Then
    Workbooks.Open (SourceSysRapor)
    Set WsSysRapor = Workbooks(SistemRapor).Worksheets(1)
    WsSysRapor.Unprotect Password:="123"
    WsSysRapor.Columns("B:C").EntireColumn.Hidden = False

    Workbooks.Open (RaporOp)
    Set WsRapor = Workbooks(StrRapor).Worksheets(1)

    'İşlem günlüğünde başlangıç ve bitiş satırlarını tespit et.
    'Say1IslemGunlugu = WsSysRapor.Range("B100000").End(xlUp).Row
    Say2IslemGunlugu = WsSysRapor.Range("C100000").End(xlUp).Row
    If Say2IslemGunlugu < 8 Then
        Say2IslemGunlugu = 8
    End If
    'SayAyracIslemGunlugu = WsSysRapor.Range("E100000").End(xlUp).Row

    'Ayracı oluştur
    ModulTarih = StrRaporTarihiAylik
    ModulAyrac = "01" & Right(ModulTarih, 8)
    SiraNo = 0
    
    Set BulIslemGunlugu = WsSysRapor.Range("E:E").Find(What:=CDate(ModulAyrac), SearchDirection:=xlNext, _
            SearchOrder:=xlByRows, LookIn:=xlFormulas, LookAt:=xlWhole)
    If Not BulIslemGunlugu Is Nothing Then 'HEDEF DÖNEMİ bul
        ilkrow = BulIslemGunlugu.Row + 1
        ModulAyrac = DateAdd("m", 1, CDate(ModulAyrac)) 'Sonraki dönemi bul ve onun üstüne satır ekle
        Set BulIslemGunlugux = WsSysRapor.Range("E:E").Find(What:=CDate(ModulAyrac), SearchDirection:=xlNext, _
                SearchOrder:=xlByRows, LookIn:=xlFormulas, LookAt:=xlWhole)
        If Not BulIslemGunlugux Is Nothing Then
            sonrow = BulIslemGunlugux.Row - 1
        Else
            sonrow = Say2IslemGunlugu
        End If
        ModulAyrac = DateAdd("m", -1, CDate(ModulAyrac)) 'Yukarıda atadığın +1 ayı geri al
        
        If WsSysRapor.Cells(ilkrow, 6).Value <> "" Then
            SiraNo = 1
            WsRapor.Range(WsRapor.Cells(2, 1), WsRapor.Cells(2 + sonrow - ilkrow, 15)).Value = WsSysRapor.Range(WsSysRapor.Cells(ilkrow, 6), WsSysRapor.Cells(sonrow, 20)).Value
        Else
            '
        End If
    Else
        '
    End If
    
    'Kenarlıklar.
    If SiraNo > 0 Then
        Set Kenarlar = WsRapor.Range("A" & 1 & ":O" & 2 + sonrow - ilkrow)
        Kenarlar.Borders.LineStyle = xlNone
        Kenarlar.Borders(xlDiagonalDown).LineStyle = xlNone
        Kenarlar.Borders(xlDiagonalUp).LineStyle = xlNone
        Kenarlar.Borders.LineStyle = xlContinuous
    End If
    With WsRapor.PageSetup
        .CenterFooter = "&""Open Sans""&18" & Format(CDate(ModulAyrac), "mmmm yyyy")
    End With



    OpenControl = IsWorkBookOpen(RaporOp)
    If OpenControl = True Then
        Workbooks(StrRapor).Save
    End If
    
    '_________________
    
    WsSysRapor.Columns("B:C").EntireColumn.Hidden = True
    WsSysRapor.Protect Password:="123"
    OpenControl = IsWorkBookOpen(SourceSysRapor)
    If OpenControl = True Then
        Workbooks(SistemRapor).Close SaveChanges:=True
    End If

End If



' Unload Me


Son:


''


Out:

Application.DisplayAlerts = True
Application.EnableEvents = True
Application.ScreenUpdating = True

ThisWorkbook.Activate
ActiveSheet.DisplayPageBreaks = False
    
If Not WsRapor1 Is Nothing Then
    WsRapor1.Activate
ElseIf Not WsRapor Is Nothing Then
    WsRapor.Activate
End If


End Sub

Private Sub RaporlaGunluk_Click()
Dim AutoPath As String, IslemGunlukleriKlasor As String, OpenControl As String, KontrolFile As String
Dim i As Long, j As Long, ContSay As Long

Dim SourceRapor1 As String, SourceRapor As String, SourceRapor2_2 As String
Dim StrRapor1 As String, StrRapor As String, StrRapor2_2 As String
Dim Rapor1Op As String, RaporOp As String, Rapor2_2Op As String
Dim WsRapor1 As Worksheet, WsRapor As Worksheet, WsRapor2_2 As Worksheet

Dim SourceSysRapor1 As String, SourceSysRapor As String, SourceSysRapor2_2 As String
Dim SistemRapor1 As String, SistemRapor As String, SistemRapor2_2 As String
Dim WsSysRapor1 As Worksheet, WsSysRapor As Worksheet, WsSysRapor2_2 As Worksheet

Dim DestOperasyon As String, DestOpUserFolderName As String, DestOpUserFolder As String
Dim fso As Object, StrRaporTarihi As String, Kenarlar As Range

Dim IlkSira As Long, SonSira As Long, SiraNo As Long



Application.DisplayAlerts = False
Application.EnableEvents = False
Application.ScreenUpdating = False


If GunlukTarihi.Value <> "" Then
    If CheckBoxRapor1.Value = False And CheckBoxRapor.Value = False And CheckBoxRapor2_2.Value = False Then
        MsgBox "Please specify the type of daily operation log you want to generate a report for.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Out
    End If
End If

If CheckBoxRapor1.Value = True Or CheckBoxRapor.Value = True Or CheckBoxRapor2_2.Value = True Then
    If GunlukTarihi.Value = "" Then
        MsgBox "Please specify the date of the daily operation log you want to report on.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Out
    End If
End If

If GunlukTarihi.Value <> "" Then
    If CheckBoxRapor1.Value = True Or CheckBoxRapor.Value = True Or CheckBoxRapor2_2.Value = True Then
        StrRaporTarihi = GunlukTarihi.Value
    Else
        GoTo Out
    End If
Else
    GoTo Out
End If

'____________


'Pathfinder...
AutoPath = ThisWorkbook.Path
DestOperasyon = AutoPath & "\System Files\Operation\"

IslemGunlukleriKlasor = AutoPath & "\System Files\System Templates\Registry Reports\"

StrRapor1 = "Registry Report 1.xlsx"
StrRapor = "Registry Report 2.1.xlsx"
StrRapor2_2 = "Registry Report 2.2.xlsx"
SourceRapor1 = IslemGunlukleriKlasor & StrRapor1
SourceRapor = IslemGunlukleriKlasor & StrRapor
SourceRapor2_2 = IslemGunlukleriKlasor & StrRapor2_2

SistemRapor1 = "System Registry Report 1.xlsx"
SistemRapor = "System Registry Report 2.1.xlsx"
SistemRapor2_2 = "System Registry Report 2.2.xlsx"
SourceSysRapor1 = IslemGunlukleriKlasor & SistemRapor1
SourceSysRapor = IslemGunlukleriKlasor & SistemRapor
SourceSysRapor2_2 = IslemGunlukleriKlasor & SistemRapor2_2


'Birden fazla kullanıcının operasyon klasöründe çakışmasını önlemek için operasyon klasörünün içinde yeni klasör aç.
DestOpUserFolderName = "Operation-" & Left(ThisWorkbook.name, InStr(ThisWorkbook.name, ".") - 1) '& " " & Format(Now(), "ddmmyyyyhhmmss")
DestOpUserFolder = DestOperasyon & DestOpUserFolderName & "\"

'Check Operation folder name.
If Not Dir(DestOperasyon, vbDirectory) <> vbNullString Then
    MsgBox "Cannot access directory: " & DestOperasyon & ". The folder named 'Operation' may have been renamed or deleted.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Out
End If

'Operation klasörü içinde kullanıcı modülü klasörü yoksa oluştur.
If Not Dir(DestOpUserFolder, vbDirectory) <> vbNullString Then
    MkDir DestOpUserFolder
End If

Rapor1Op = DestOpUserFolder & StrRapor1
RaporOp = DestOpUserFolder & StrRapor
Rapor2_2Op = DestOpUserFolder & StrRapor2_2


'Check Operation Logs folder name.
If Not Dir(IslemGunlukleriKlasor, vbDirectory) <> vbNullString Then
    MsgBox "Cannot access directory: " & IslemGunlukleriKlasor & ". The folder named 'Registry Reports' may have been renamed or deleted.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Out
End If

If Not Dir(SourceRapor1, vbDirectory) <> vbNullString Then
    MsgBox "Cannot access directory: " & SourceRapor1 & ". Folder or file names within this directory may have been changed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Out
End If

If Not Dir(SourceRapor, vbDirectory) <> vbNullString Then
    MsgBox "Cannot access directory: " & SourceRapor & ". Folder or file names within this directory may have been changed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Out
End If

If Not Dir(SourceRapor2_2, vbDirectory) <> vbNullString Then
    MsgBox "Cannot access directory: " & SourceRapor2_2 & ". Folder or file names within this directory may have been changed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Out
End If

If Not Dir(SourceSysRapor1, vbDirectory) <> vbNullString Then
    MsgBox "Cannot access directory: " & SourceSysRapor1 & ". Folder or file names within this directory may have been changed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Out
End If

If Not Dir(SourceSysRapor, vbDirectory) <> vbNullString Then
    MsgBox "Cannot access directory: " & SourceSysRapor & ". Folder or file names within this directory may have been changed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Out
End If

If Not Dir(SourceSysRapor2_2, vbDirectory) <> vbNullString Then
    MsgBox "Cannot access directory: " & SourceSysRapor2_2 & ". Folder or file names within this directory may have been changed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Out
End If


On Error Resume Next 'Operation içinde .xlsx dosyası yoksa oluşacak hata için
'İşlem günlüğü açıksa kaydet ve kapat.
OpenControl = IsWorkBookOpen(Rapor1Op)
If OpenControl = True Then
    Workbooks(StrRapor1).Close SaveChanges:=False
End If
'İşlem günlüğü açıksa kaydet ve kapat.
OpenControl = IsWorkBookOpen(RaporOp)
If OpenControl = True Then
    Workbooks(StrRapor).Close SaveChanges:=False
End If
'İşlem günlüğü açıksa kaydet ve kapat.
OpenControl = IsWorkBookOpen(Rapor2_2Op)
If OpenControl = True Then
    Workbooks(StrRapor2_2).Close SaveChanges:=False
End If

OpenControl = IsWorkBookOpen(SourceSysRapor1)
If OpenControl = True Then
    Workbooks(SistemRapor1).Close SaveChanges:=False
End If
'İşlem günlüğü açıksa kaydet ve kapat.
OpenControl = IsWorkBookOpen(SourceSysRapor)
If OpenControl = True Then
    Workbooks(SistemRapor).Close SaveChanges:=False
End If
'İşlem günlüğü açıksa kaydet ve kapat.
OpenControl = IsWorkBookOpen(SourceSysRapor2_2)
If OpenControl = True Then
    Workbooks(SistemRapor2_2).Close SaveChanges:=False
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

On Error GoTo 0

'Dosyayı operasyon klasörüne kopyala ve adını değiştir.
Set fso = CreateObject("Scripting.FileSystemObject")
fso.CopyFile (SourceRapor1), DestOpUserFolder & StrRapor1, True
fso.CopyFile (SourceRapor), DestOpUserFolder & StrRapor, True
fso.CopyFile (SourceRapor2_2), DestOpUserFolder & StrRapor2_2, True

'open the file
If CheckBoxRapor1.Value = True Then
    Workbooks.Open (SourceSysRapor1)
    Set WsSysRapor1 = Workbooks(SistemRapor1).Worksheets(1)
    WsSysRapor1.Unprotect Password:="123"
    WsSysRapor1.Columns("B:C").EntireColumn.Hidden = False

    Workbooks.Open (Rapor1Op)
    Set WsRapor1 = Workbooks(StrRapor1).Worksheets(1)
    
            
    '_________________
    
    j = 2
    SiraNo = 0
    StrAramaGlobal = StrRaporTarihi
    Set MyRngGlobal = WsSysRapor1.Range("N:N")
    Set MyFinderGlobal = MyRngGlobal.Find(What:=StrAramaGlobal, _
                    SearchDirection:=xlNext, MatchCase:=False, SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    If Not MyFinderGlobal Is Nothing Then
        IlkAdresGlobal = MyFinderGlobal.Address
        'MsgBox Replace(IlkAdres, "$", ""), vbOKOnly, "ishakkutlu.com"
        'Sonraki satırlarda aramaya devam et
        Do

            '______________

            i = MyFinderGlobal.Row
            Do Until WsSysRapor1.Range("C" & i).Value <> ""
                i = i + 1
                If i - MyFinderGlobal.Row >= 20 Then
                    GoTo DoWhileBitir1
                End If
            Loop
DoWhileBitir1:
            'MsgBox "SonSira: " & i

            IlkSira = MyFinderGlobal.Row
            SonSira = i
            ilkrow = j
            sonrow = ilkrow + SonSira - IlkSira
            WsRapor1.Range(WsRapor1.Cells(ilkrow, 2), WsRapor1.Cells(sonrow, 14)).Value = WsSysRapor1.Range(WsSysRapor1.Cells(IlkSira, 7), WsSysRapor1.Cells(SonSira, 19)).Value
            SiraNo = SiraNo + 1
            WsRapor1.Cells(ilkrow, 1).Value = SiraNo
            j = sonrow + 1
            
            '______________
            
            SonrakiAdresGlobal = MyFinderGlobal.Address
            'MsgBox Replace(SonrakiAdresGlobal, "$", ""), vbOKOnly, "ishakkutlu.com"
            Set MyFinderGlobal = MyRngGlobal.FindNext(MyFinderGlobal)
            SonrakiAdresGlobal = MyFinderGlobal.Address

        Loop While IlkAdresGlobal <> SonrakiAdresGlobal
    End If

    'Kenarlıklar.
    If SiraNo > 0 Then
        Set Kenarlar = WsRapor1.Range("A" & 1 & ":N" & sonrow)
        Kenarlar.Borders.LineStyle = xlNone
        Kenarlar.Borders(xlDiagonalDown).LineStyle = xlNone
        Kenarlar.Borders(xlDiagonalUp).LineStyle = xlNone
        Kenarlar.Borders.LineStyle = xlContinuous
    End If
    With WsRapor1.PageSetup
        .CenterFooter = "&""Open Sans""&18" & "Yeni kayıt girişi yapılan evrak bilgileri kontrol edilmiştir. " & StrRaporTarihi
    End With
    OpenControl = IsWorkBookOpen(Rapor1Op)
    If OpenControl = True Then
        Workbooks(StrRapor1).Save
    End If
    
    '_________________
    
    WsSysRapor1.Columns("B:C").EntireColumn.Hidden = True
    WsSysRapor1.Protect Password:="123"
    OpenControl = IsWorkBookOpen(SourceSysRapor1)
    If OpenControl = True Then
        Workbooks(SistemRapor1).Close SaveChanges:=True
    End If
    
End If

If CheckBoxRapor.Value = True Then
    Workbooks.Open (SourceSysRapor)
    Set WsSysRapor = Workbooks(SistemRapor).Worksheets(1)
    WsSysRapor.Unprotect Password:="123"
    WsSysRapor.Columns("B:C").EntireColumn.Hidden = False

    Workbooks.Open (RaporOp)
    Set WsRapor = Workbooks(StrRapor).Worksheets(1)
    
            
    '_________________
    
    j = 2
    SiraNo = 0
    StrAramaGlobal = StrRaporTarihi
    Set MyRngGlobal = WsSysRapor.Range("N:N")
    Set MyFinderGlobal = MyRngGlobal.Find(What:=StrAramaGlobal, _
                    SearchDirection:=xlNext, MatchCase:=False, SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    If Not MyFinderGlobal Is Nothing Then
        IlkAdresGlobal = MyFinderGlobal.Address
        'MsgBox Replace(IlkAdres, "$", ""), vbOKOnly, "ishakkutlu.com"
        'Sonraki satırlarda aramaya devam et
        Do

            '______________

            i = MyFinderGlobal.Row
            Do Until WsSysRapor.Range("C" & i).Value <> ""
                i = i + 1
                If i - MyFinderGlobal.Row >= 20 Then
                    GoTo DoWhileBitir2
                End If
            Loop
DoWhileBitir2:
            'MsgBox "SonSira: " & i

            IlkSira = MyFinderGlobal.Row
            SonSira = i
            ilkrow = j
            sonrow = ilkrow + SonSira - IlkSira
            WsRapor.Range(WsRapor.Cells(ilkrow, 2), WsRapor.Cells(sonrow, 15)).Value = WsSysRapor.Range(WsSysRapor.Cells(IlkSira, 7), WsSysRapor.Cells(SonSira, 20)).Value
            SiraNo = SiraNo + 1
            WsRapor.Cells(ilkrow, 1).Value = SiraNo
            j = sonrow + 1
            
            '______________
            
            SonrakiAdresGlobal = MyFinderGlobal.Address
            'MsgBox Replace(SonrakiAdresGlobal, "$", ""), vbOKOnly, "ishakkutlu.com"
            Set MyFinderGlobal = MyRngGlobal.FindNext(MyFinderGlobal)
            SonrakiAdresGlobal = MyFinderGlobal.Address

        Loop While IlkAdresGlobal <> SonrakiAdresGlobal
    End If

    'Kenarlıklar.
    If SiraNo > 0 Then
        Set Kenarlar = WsRapor.Range("A" & 1 & ":O" & sonrow)
        Kenarlar.Borders.LineStyle = xlNone
        Kenarlar.Borders(xlDiagonalDown).LineStyle = xlNone
        Kenarlar.Borders(xlDiagonalUp).LineStyle = xlNone
        Kenarlar.Borders.LineStyle = xlContinuous
    End If
    
    With WsRapor.PageSetup
        .CenterFooter = "&""Open Sans""&18" & "Yeni kayıt girişi yapılan evrak bilgileri kontrol edilmiştir. " & StrRaporTarihi
    End With
    
    OpenControl = IsWorkBookOpen(RaporOp)
    If OpenControl = True Then
        Workbooks(StrRapor).Save
    End If
    
    '_________________
    
    WsSysRapor.Columns("B:C").EntireColumn.Hidden = True
    WsSysRapor.Protect Password:="123"
    OpenControl = IsWorkBookOpen(SourceSysRapor)
    If OpenControl = True Then
        Workbooks(SistemRapor).Close SaveChanges:=True
    End If
    
End If

If CheckBoxRapor2_2.Value = True Then
    Workbooks.Open (SourceSysRapor2_2)
    Set WsSysRapor2_2 = Workbooks(SistemRapor2_2).Worksheets(1)
    WsSysRapor2_2.Unprotect Password:="123"
    WsSysRapor2_2.Columns("B:C").EntireColumn.Hidden = False

    Workbooks.Open (Rapor2_2Op)
    Set WsRapor2_2 = Workbooks(StrRapor2_2).Worksheets(1)
    
            
    '_________________
    
    j = 2
    SiraNo = 0
    StrAramaGlobal = StrRaporTarihi
    Set MyRngGlobal = WsSysRapor2_2.Range("N:N")
    Set MyFinderGlobal = MyRngGlobal.Find(What:=StrAramaGlobal, _
                    SearchDirection:=xlNext, MatchCase:=False, SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    If Not MyFinderGlobal Is Nothing Then
        IlkAdresGlobal = MyFinderGlobal.Address
        'MsgBox Replace(IlkAdres, "$", ""), vbOKOnly, "ishakkutlu.com"
        'Sonraki satırlarda aramaya devam et
        Do

            '______________

            i = MyFinderGlobal.Row
            Do Until WsSysRapor2_2.Range("C" & i).Value <> ""
                i = i + 1
                If i - MyFinderGlobal.Row >= 20 Then
                    GoTo DoWhileBitir3
                End If
            Loop
DoWhileBitir3:
            'MsgBox "SonSira: " & i

            IlkSira = MyFinderGlobal.Row
            SonSira = i
            ilkrow = j
            sonrow = ilkrow + SonSira - IlkSira
            WsRapor2_2.Range(WsRapor2_2.Cells(ilkrow, 2), WsRapor2_2.Cells(sonrow, 14)).Value = WsSysRapor2_2.Range(WsSysRapor2_2.Cells(IlkSira, 7), WsSysRapor2_2.Cells(SonSira, 19)).Value
            SiraNo = SiraNo + 1
            WsRapor2_2.Cells(ilkrow, 1).Value = SiraNo
            j = sonrow + 1
            
            '______________
            
            SonrakiAdresGlobal = MyFinderGlobal.Address
            'MsgBox Replace(SonrakiAdresGlobal, "$", ""), vbOKOnly, "ishakkutlu.com"
            Set MyFinderGlobal = MyRngGlobal.FindNext(MyFinderGlobal)
            SonrakiAdresGlobal = MyFinderGlobal.Address

        Loop While IlkAdresGlobal <> SonrakiAdresGlobal
    End If

    'Kenarlıklar.
    If SiraNo > 0 Then
        Set Kenarlar = WsRapor2_2.Range("A" & 1 & ":N" & sonrow)
        Kenarlar.Borders.LineStyle = xlNone
        Kenarlar.Borders(xlDiagonalDown).LineStyle = xlNone
        Kenarlar.Borders(xlDiagonalUp).LineStyle = xlNone
        Kenarlar.Borders.LineStyle = xlContinuous
    End If

    With WsRapor2_2.PageSetup
        .CenterFooter = "&""Open Sans""&18" & "Yeni kayıt girişi yapılan evrak bilgileri kontrol edilmiştir. " & StrRaporTarihi
    End With
    
    OpenControl = IsWorkBookOpen(Rapor2_2Op)
    If OpenControl = True Then
        Workbooks(StrRapor2_2).Save
    End If
    
    '_________________
    
    WsSysRapor2_2.Columns("B:C").EntireColumn.Hidden = True
    WsSysRapor2_2.Protect Password:="123"
    OpenControl = IsWorkBookOpen(SourceSysRapor2_2)
    If OpenControl = True Then
        Workbooks(SistemRapor2_2).Close SaveChanges:=True
    End If
    
End If



'Unload Me


Son:


''



Out:

Application.DisplayAlerts = True
Application.EnableEvents = True
Application.ScreenUpdating = True

ThisWorkbook.Activate
ActiveSheet.DisplayPageBreaks = False
    
If Not WsRapor1 Is Nothing Then
    WsRapor1.Activate
ElseIf Not WsRapor Is Nothing Then
    WsRapor.Activate
ElseIf Not WsRapor2_2 Is Nothing Then
    WsRapor2_2.Activate
End If


End Sub

Private Sub Kapat_Click()
    Unload Me
End Sub


Private Sub UserForm_Initialize()
Dim i As Long, WsSKP As Object
Dim ClrLab As MSForms.Control

ThisWorkbook.Activate

For Each ClrLab In core_registry_reports_UI.Controls

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

RaporlaGunluk.BackColor = RGB(225, 235, 245)
RaporlaGunluk.ForeColor = RGB(30, 30, 30)
RaporlaAylik.BackColor = RGB(225, 235, 245)
RaporlaAylik.ForeColor = RGB(30, 30, 30)
Kapat.BackColor = RGB(225, 235, 245)
Kapat.ForeColor = RGB(30, 30, 30)


core_registry_reports_UI.BackColor = RGB(230, 230, 230) 'YENİ


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
        core_registry_reports_UI.Width = Rep
        yukseklik = yukseklik - 60
        core_registry_reports_UI.Height = yukseklik
        If yukseklik <= 60 Then
            yukseklik = 60
            core_registry_reports_UI.Height = yukseklik
        End If
    ElseIf Rep <= 50 Then
        Rep = 50
        core_registry_reports_UI.Width = Rep
        yukseklik = yukseklik - 50
        core_registry_reports_UI.Height = yukseklik
        If yukseklik <= 50 Then
            yukseklik = 50
            core_registry_reports_UI.Height = yukseklik
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


