VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} core_system_reset_wizard2_UI 
   Caption         =   "Enterprise Document Automation System"
   ClientHeight    =   8895
   ClientLeft      =   30
   ClientTop       =   375
   ClientWidth     =   13920
   OleObjectBlob   =   "core_system_reset_wizard2_UI.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "core_system_reset_wizard2_UI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Sub ColorChangerGenel()

'LabelResetle
If LabelResetle.BackColor <> RGB(225, 235, 245) Then
    LabelResetle.BackColor = RGB(225, 235, 245)
    LabelResetle.ForeColor = RGB(30, 30, 30)
End If
'LabelKapat
If LabelKapat.BackColor <> RGB(225, 235, 245) Then
    LabelKapat.BackColor = RGB(225, 235, 245)
    LabelKapat.ForeColor = RGB(30, 30, 30)
End If


End Sub

Private Sub LabelResetle_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
LabelResetle.BackColor = RGB(60, 100, 180) 'RGB(60, 100, 180) 'RGB(60, 100, 180)
LabelResetle.ForeColor = RGB(256, 256, 256)
End Sub

Private Sub LabelKapat_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
LabelKapat.BackColor = RGB(60, 100, 180) 'RGB(60, 100, 180) 'RGB(60, 100, 180)
LabelKapat.ForeColor = RGB(256, 256, 256)
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


Private Sub LabelResetle_Click()

If SifreText.Value = "123" Then
    
    Call ModuleRibbon.ResetProsedur
    
    Unload Me
    Unload core_system_reset_wizard1_UI
    
    'GlobalResetKapsami = 1 'dahil
    'GlobalResetKapsami = 2 'hariç
ElseIf SifreText.Value = "" Then
    MsgBox "Please enter the password to proceed the reset process.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
ElseIf SifreText.Value <> "" And SifreText.Value <> "123" Then
    MsgBox "The reset process could not be proceeded due to an incorrect password.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
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

For Each ClrLab In core_system_reset_wizard2_UI.Controls
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
LblSecimler.BackColor = RGB(254, 254, 254)

LabelResetle.BackColor = RGB(225, 235, 245)
LabelResetle.ForeColor = RGB(30, 30, 30)
LabelKapat.BackColor = RGB(225, 235, 245)
LabelKapat.ForeColor = RGB(30, 30, 30)

core_system_reset_wizard2_UI.BackColor = RGB(230, 230, 230) 'YENİ

GlobalResetKapsami = 0
'System Reset kapsamını tespit et
If core_system_reset_wizard1_UI.VarlikDahilOption.Value = True Then
    GlobalResetKapsami = 1
End If
If core_system_reset_wizard1_UI.VarlikHaricOption.Value = True Then
    GlobalResetKapsami = 2
End If
'Kapsama göre uyarı metnini değiştir.
If GlobalResetKapsami = 1 Then
    LblSecimler.Caption = "WARNING!" & vbNewLine & _
    "Except for unit-specific definitions, ALL data in the system — INCLUDING operations WITHOUT asset output — will be deleted within the boundaries you defined in the previous step."
ElseIf GlobalResetKapsami = 2 Then
    LblSecimler.Caption = "WARNING!" & vbNewLine & _
    "Except for unit-specific definitions, ALL data in the system — EXCLUDING operations WITHOUT asset output — will be deleted within the boundaries you defined in the previous step."
End If


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
        core_system_reset_wizard2_UI.Width = Rep
        yukseklik = yukseklik - 60
        core_system_reset_wizard2_UI.Height = yukseklik
        If yukseklik <= 60 Then
            yukseklik = 60
            core_system_reset_wizard2_UI.Height = yukseklik
        End If
    ElseIf Rep <= 50 Then
        Rep = 50
        core_system_reset_wizard2_UI.Width = Rep
        yukseklik = yukseklik - 50
        core_system_reset_wizard2_UI.Height = yukseklik
        If yukseklik <= 50 Then
            yukseklik = 50
            core_system_reset_wizard2_UI.Height = yukseklik
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





