VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} core_auto_close_settings_UI 
   Caption         =   "Enterprise Document Automation System"
   ClientHeight    =   8895
   ClientLeft      =   30
   ClientTop       =   375
   ClientWidth     =   13920
   OleObjectBlob   =   "core_auto_close_settings_UI.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "core_auto_close_settings_UI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim Abort As Boolean

Sub ColorChangerGenel()

If Kaydet.BackColor <> RGB(225, 235, 245) Then
Kaydet.BackColor = RGB(225, 235, 245)
Kaydet.ForeColor = RGB(30, 30, 30)
End If

If Kapat.BackColor <> RGB(225, 235, 245) Then
Kapat.BackColor = RGB(225, 235, 245)
Kapat.ForeColor = RGB(30, 30, 30)
End If

End Sub

Private Sub AktifPasif_Click()

'AKTİF/PASİF
If AktifPasif.Value = True Then
    AktifPasif.Caption = "Disable Auto-Close"
Else
    AktifPasif.Caption = "Enable Auto-Close"
End If

End Sub

Private Sub Kaydet_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
Kaydet.BackColor = RGB(60, 100, 180) 'RGB(60, 100, 180)
Kaydet.ForeColor = RGB(255, 255, 255)
End Sub

Private Sub Kapat_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
Kapat.BackColor = RGB(60, 100, 180) 'RGB(60, 100, 180)
Kapat.ForeColor = RGB(255, 255, 255)
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

'Saat color
Private Sub SaatFrame_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub
Private Sub Saat_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
Call SetComboBoxHook(Saat)
End Sub
Private Sub LblSaat_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub
Private Sub Saat_LostFocus()
Call RemoveComboBoxHook
End Sub
Private Sub Saat_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Me.Saat.DropDown
End Sub

'Dakika color
Private Sub DakikaFrame_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub
Private Sub Dakika_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
Call SetComboBoxHook(Dakika)
End Sub
Private Sub LblDakika_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub
Private Sub Dakika_LostFocus()
Call RemoveComboBoxHook
End Sub
Private Sub Dakika_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Me.Dakika.DropDown
End Sub

'Saniye color
Private Sub SaniyeFrame_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub
Private Sub Saniye_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
Call SetComboBoxHook(Saniye)
End Sub
Private Sub LblSaniye_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub
Private Sub Saniye_LostFocus()
Call RemoveComboBoxHook
End Sub
Private Sub Saniye_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Me.Saniye.DropDown
End Sub

'AktifPasif color
Private Sub AktifPasifFrame_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub
Private Sub AktifPasif_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub


Private Sub Saat_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next

   Select Case KeyCode
        Case 38  'Yukarı
            If Saat.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                Saat.ListIndex = Saat.ListIndex - 1
            End If
            Me.Saat.DropDown
            
        Case 40 'Aşağı
            If Saat.ListIndex = Saat.ListCount - 1 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                Saat.ListIndex = Saat.ListIndex + 1
            End If
            Me.Saat.DropDown
    End Select
    Abort = False
    
End Sub

Private Sub Saat_Change()

If Saat.ListIndex = -1 And Saat.Value <> "" Then
   Saat.Value = ""
   GoTo Son
End If

If Saat.Value <> "" Then
    Saat.SelStart = 0
    Saat.SelLength = Len(Saat.Value)
End If

Son:

Saat.DropDown

End Sub

Private Sub Dakika_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next

   Select Case KeyCode
        Case 38  'Yukarı
            If Dakika.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                Dakika.ListIndex = Dakika.ListIndex - 1
            End If
            Me.Dakika.DropDown
            
        Case 40 'Aşağı
            If Dakika.ListIndex = Dakika.ListCount - 1 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                Dakika.ListIndex = Dakika.ListIndex + 1
            End If
            Me.Dakika.DropDown
    End Select
    Abort = False
    
End Sub

Private Sub Dakika_Change()

If Dakika.ListIndex = -1 And Dakika.Value <> "" Then
   Dakika.Value = ""
   GoTo Son
End If

If Dakika.Value <> "" Then
    Dakika.SelStart = 0
    Dakika.SelLength = Len(Dakika.Value)
End If

Son:

Dakika.DropDown

End Sub

Private Sub Saniye_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next

   Select Case KeyCode
        Case 38  'Yukarı
            If Saniye.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                Saniye.ListIndex = Saniye.ListIndex - 1
            End If
            Me.Saniye.DropDown
            
        Case 40 'Aşağı
            If Saniye.ListIndex = Saniye.ListCount - 1 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                Saniye.ListIndex = Saniye.ListIndex + 1
            End If
            Me.Saniye.DropDown
    End Select
    Abort = False
    
End Sub

Private Sub Saniye_Change()

If Saniye.ListIndex = -1 And Saniye.Value <> "" Then
   Saniye.Value = ""
   GoTo Son
End If

If Saniye.Value <> "" Then
    Saniye.SelStart = 0
    Saniye.SelLength = Len(Saniye.Value)
End If

Son:

Saniye.DropDown

End Sub


Private Sub Kaydet_Click()
Dim WsSKP As Object
Dim AutoPath As String, DestTarget As String, OpenControl As String
Dim FileName As String, Tanimlar As String


Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.EnableEvents = False

ThisWorkbook.Unprotect "123"

AutoPath = ThisWorkbook.Path
DestTarget = AutoPath & "\System Files\System Definitions\"
FileName = "Definitions.xlsx"
Tanimlar = AutoPath & "\System Files\System Definitions\Definitions.xlsx"


'Check "System Files" folder name.
If Not Dir(AutoPath & "\System Files\", vbDirectory) <> vbNullString Then
    MsgBox "Cannot access directory: " & AutoPath & "\System Files\" & ". The folder named 'System Files' may have been renamed or deleted.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

'Check Definitions file name.
If Not Dir(Tanimlar, vbDirectory) <> vbNullString Then
    MsgBox "Cannot access directory: " & Tanimlar & ". The file named 'System Definitions' may have been renamed or deleted.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If


OpenControl = IsFileOpen(DestTarget & FileName)
If OpenControl = True Then 'Açıksa
    Workbooks(FileName).Close SaveChanges:=True
ElseIf OpenControl = False Then
    '
End If
Workbooks.Open (DestTarget & FileName)
'Workbooks(FileName).Worksheets(1).Activate


Set WsSKP = ThisWorkbook.Worksheets(2)

xSaat = Saat.Value
xDakika = Dakika.Value
xSaniye = Saniye.Value

If AktifPasif.Value = True Then
    xDurum = "Enable"
Else
    xDurum = "Disable"
    xSaat = ""
    xDakika = ""
    xSaniye = ""
End If

'Bekleme süresi yoksa pasif kabul et
If xSaat = "" And xDakika = "" And xSaniye = "" Then
    xDurum = "Disable"
    'MsgBox "test"
    If AktifPasif.Value = True Then
        AktifPasif.Value = False
    End If

End If

WsSKP.Unprotect "123"
WsSKP.Cells(5, 121).Value = "Automatic Closure Information"
WsSKP.Protect "123"

WsSKP.Cells(6, 121).Value = xDurum
WsSKP.Cells(7, 121).Value = xSaat
WsSKP.Cells(8, 121).Value = xDakika
WsSKP.Cells(9, 121).Value = xSaniye

Workbooks(FileName).Worksheets(1).Unprotect "123"
Workbooks(FileName).Worksheets(1).Cells(5, 121).Value = "Automatic Closure Information"
Workbooks(FileName).Worksheets(1).Protect "123"

Workbooks(FileName).Worksheets(1).Cells(6, 121).Value = xDurum
Workbooks(FileName).Worksheets(1).Cells(7, 121).Value = xSaat
Workbooks(FileName).Worksheets(1).Cells(8, 121).Value = xDakika
Workbooks(FileName).Worksheets(1).Cells(9, 121).Value = xSaniye


If xSaat = "" Then
    xSaat = 0
End If
If xDakika = "" Then
    xDakika = 0
End If
If xSaniye = "" Then
    xSaniye = 0
End If


Workbooks(FileName).Save
OpenControl = IsFileOpen(DestTarget & FileName)
If OpenControl = True Then 'Açıksa
    Workbooks(FileName).Close SaveChanges:=True
ElseIf OpenControl = False Then
    '
End If


ThisWorkbook.Protect "123"
''ThisWorkbook.Save
'Call ModuleSystemSettings.DropDownKapat

If xDurum = "Enable" Then
    Call ModuleSystemSettings.TimeStop
    Call ModuleSystemSettings.TimeSetting
    MsgBox "Automatic shutdown has been ACTIVATED successfully. The wait time is set to " & xSaat & " hours, " & xDakika & " minutes, and " & xSaniye & " seconds.", vbOKOnly + vbInformation, "Enterprise Document Automation System"
    Unload Me
Else
    Call ModuleSystemSettings.TimeStop
    MsgBox "Automatic shutdown has been successfully CANCELED.", vbOKOnly + vbInformation, "Enterprise Document Automation System"
    Unload Me
End If


GoTo Out

Son:

ThisWorkbook.Protect "123"

OpenControl = IsFileOpen(DestTarget & FileName)
If OpenControl = True Then 'Açıksa
    Workbooks(FileName).Close SaveChanges:=True
ElseIf OpenControl = False Then
    '
End If

Out:

'ThisWorkbook.Worksheets(1).Activate

Application.ScreenUpdating = True
Application.DisplayAlerts = True
Application.EnableEvents = True


End Sub


Private Sub Kapat_Click()
    Unload Me
End Sub

Private Sub UserForm_Initialize()
Dim i As Long, WsSKP As Object
Dim ClrLab As MSForms.Control


ThisWorkbook.Activate


For Each ClrLab In core_auto_close_settings_UI.Controls
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

Kaydet.BackColor = RGB(225, 235, 245)
Kaydet.ForeColor = RGB(30, 30, 30)
Kapat.BackColor = RGB(225, 235, 245)
Kapat.ForeColor = RGB(30, 30, 30)
core_auto_close_settings_UI.BackColor = RGB(230, 230, 230) 'YENİ

'Saat liste değerleri
For i = 1 To 23 Step 1
    With Saat
        .AddItem (i)
    End With
Next i

'Dakika liste değerleri
For i = 1 To 59 Step 1
    With Dakika
        .AddItem (i)
    End With
Next i

'Saniye liste değerleri
For i = 1 To 59 Step 1
    With Saniye
        .AddItem (i)
    End With
Next i


Set WsSKP = ThisWorkbook.Worksheets(2)

On Error GoTo PasifeGit
If WsSKP.Cells(6, 121).Value = "Enable" Then
    AktifPasif.Value = True
ElseIf WsSKP.Cells(6, 121).Value = "Disable" Then
    AktifPasif.Value = False
Else
    AktifPasif.Value = False
End If


Saat.Value = WsSKP.Cells(7, 121).Value
Dakika.Value = WsSKP.Cells(8, 121).Value
Saniye.Value = WsSKP.Cells(9, 121).Value


On Error GoTo 0
GoTo PasifiAtla

PasifeGit:
AktifPasif.Value = False
Saat.Value = ""
Dakika.Value = ""
Saniye.Value = ""


PasifiAtla:
'AKTİF/PASİF
If AktifPasif.Value = True Then
    AktifPasif.Caption = "Disable Auto-Close"
Else
    AktifPasif.Caption = "Enable Auto-Close"
    Saat.Value = ""
    Dakika.Value = ""
    Saniye.Value = ""
End If

'Açık dropdown kapat
Call ModuleSystemSettings.DropDownKapat


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
        core_auto_close_settings_UI.Width = Rep
        yukseklik = yukseklik - 60
        core_auto_close_settings_UI.Height = yukseklik
        If yukseklik <= 60 Then
            yukseklik = 60
            core_auto_close_settings_UI.Height = yukseklik
        End If
    ElseIf Rep <= 50 Then
        Rep = 50
        core_auto_close_settings_UI.Width = Rep
        yukseklik = yukseklik - 50
        core_auto_close_settings_UI.Height = yukseklik
        If yukseklik <= 50 Then
            yukseklik = 50
            core_auto_close_settings_UI.Height = yukseklik
        End If
    End If
Loop Until yukseklik = 50

Unload Me
End Sub

