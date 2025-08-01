VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} core_system_reset_wizard1_UI 
   Caption         =   "Enterprise Document Automation System"
   ClientHeight    =   8895
   ClientLeft      =   30
   ClientTop       =   375
   ClientWidth     =   13920
   OleObjectBlob   =   "core_system_reset_wizard1_UI.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "core_system_reset_wizard1_UI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Sub ColorChangerGenel()

'LabelDevam
If LabelDevam.BackColor <> RGB(225, 235, 245) Then
    LabelDevam.BackColor = RGB(225, 235, 245)
    LabelDevam.ForeColor = RGB(30, 30, 30)
End If
'LabelKapat
If LabelKapat.BackColor <> RGB(225, 235, 245) Then
    LabelKapat.BackColor = RGB(225, 235, 245)
    LabelKapat.ForeColor = RGB(30, 30, 30)
End If

'VarlikDahilOption
If VarlikDahilOption.BackColor <> RGB(254, 254, 254) Then
    VarlikDahilOption.BackColor = RGB(254, 254, 254)
    VarlikDahilOption.ForeColor = RGB(70, 70, 70)
End If

'VarlikHaricOption
If VarlikHaricOption.BackColor <> RGB(254, 254, 254) Then
    VarlikHaricOption.BackColor = RGB(254, 254, 254)
    VarlikHaricOption.ForeColor = RGB(70, 70, 70)
End If


End Sub

Private Sub LabelDevam_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
LabelDevam.BackColor = RGB(60, 100, 180) 'RGB(60, 100, 180) 'RGB(60, 100, 180)
LabelDevam.ForeColor = RGB(256, 256, 256)
End Sub

Private Sub LabelKapat_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
LabelKapat.BackColor = RGB(60, 100, 180) 'RGB(60, 100, 180) 'RGB(60, 100, 180)
LabelKapat.ForeColor = RGB(256, 256, 256)
End Sub

Private Sub VarlikDahilOption_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
VarlikDahilOption.BackColor = RGB(60, 100, 180) 'RGB(60, 100, 180)
VarlikDahilOption.ForeColor = RGB(256, 256, 256)
End Sub

Private Sub VarlikHaricOption_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
VarlikHaricOption.BackColor = RGB(60, 100, 180) 'RGB(60, 100, 180)
VarlikHaricOption.ForeColor = RGB(256, 256, 256)
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


Private Sub LabelDevam_Click()

If VarlikDahilOption.Value = False And VarlikHaricOption.Value = False Then
    MsgBox "The operation could not be completed since the reset boundaries were not specified. Please indicate whether the transactions without asset output are included in the reset process.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

If SifreText.Value = "123" Then

    If Application.Height < 766 Then
        'İkinci aşama için formu aç
        With core_system_reset_wizard2_UI
            .StartUpPosition = 0
            .Left = Application.Left + (0.5 * Application.Width) - (0.5 * 708) + 100
            .Top = Application.Top + (0.5 * Application.Height) - (0.5 * 474) '+ 100
            .Show vbModal 'vbModal 'vbModeless
        End With
    Else
        'İkinci aşama için formu aç
        With core_system_reset_wizard2_UI
            .StartUpPosition = 0
            .Left = Application.Left + (0.5 * Application.Width) - (0.5 * 708) + 100
            .Top = Application.Top + (0.5 * Application.Height) - (0.5 * 474) + 100
            .Show vbModal 'vbModal 'vbModeless
        End With
    End If

ElseIf SifreText.Value = "" Then
    MsgBox "Please enter the password to proceed with the reset process.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
ElseIf SifreText.Value <> "" And SifreText.Value <> "123" Then
    MsgBox "The reset process could not be started due to an incorrect password.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If


Son:


End Sub

Private Sub LabelKapat_Click()
    Unload Me
End Sub

Private Sub UserForm_Initialize()
Dim i As Long, WsSKP As Object
Dim ClrLab As MSForms.Control

ThisWorkbook.Activate

For Each ClrLab In core_system_reset_wizard1_UI.Controls
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

LabelDevam.BackColor = RGB(225, 235, 245)
LabelDevam.ForeColor = RGB(30, 30, 30)
LabelKapat.BackColor = RGB(225, 235, 245)
LabelKapat.ForeColor = RGB(30, 30, 30)

core_system_reset_wizard1_UI.BackColor = RGB(230, 230, 230) 'YENİ


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
        core_system_reset_wizard1_UI.Width = Rep
        yukseklik = yukseklik - 60
        core_system_reset_wizard1_UI.Height = yukseklik
        If yukseklik <= 60 Then
            yukseklik = 60
            core_system_reset_wizard1_UI.Height = yukseklik
        End If
    ElseIf Rep <= 50 Then
        Rep = 50
        core_system_reset_wizard1_UI.Width = Rep
        yukseklik = yukseklik - 50
        core_system_reset_wizard1_UI.Height = yukseklik
        If yukseklik <= 50 Then
            yukseklik = 50
            core_system_reset_wizard1_UI.Height = yukseklik
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





