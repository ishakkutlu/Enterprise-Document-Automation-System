VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} core_asset_manager_UI 
   Caption         =   "Enterprise Document Automation System"
   ClientHeight    =   11130
   ClientLeft      =   30
   ClientTop       =   375
   ClientWidth     =   20640
   OleObjectBlob   =   "core_asset_manager_UI.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "core_asset_manager_UI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim core_asset_manager_UIDirektYazdir As Boolean
Dim core_asset_manager_UISayPrt As Variant
Dim IlkSiraGlobal As Long
Dim ctl As MSForms.Control
Dim Abort As Boolean

Private Sub CheckBoxTumVarliklarEsas_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
CheckBoxTumVarliklarEsas.BackColor = RGB(60, 100, 180) 'RGB(60, 100, 180) 'RGB(60, 100, 180)
CheckBoxTumVarliklarEsas.ForeColor = RGB(255, 255, 255)
End Sub
Private Sub CheckBoxGenelVarlikEsas_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
CheckBoxGenelVarlikEsas.BackColor = RGB(60, 100, 180) 'RGB(60, 100, 180) 'RGB(60, 100, 180)
CheckBoxGenelVarlikEsas.ForeColor = RGB(255, 255, 255)
End Sub
Private Sub CheckBoxRaporVarlikEsas_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
CheckBoxRaporVarlikEsas.BackColor = RGB(60, 100, 180) 'RGB(60, 100, 180) 'RGB(60, 100, 180)
CheckBoxRaporVarlikEsas.ForeColor = RGB(255, 255, 255)
End Sub
Private Sub CheckBoxRapor2_2VarlikEsas_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
CheckBoxRapor2_2VarlikEsas.BackColor = RGB(60, 100, 180) 'RGB(60, 100, 180) 'RGB(60, 100, 180)
CheckBoxRapor2_2VarlikEsas.ForeColor = RGB(255, 255, 255)
End Sub
Private Sub CheckBoxRapor3VarlikEsas_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
CheckBoxRapor3VarlikEsas.BackColor = RGB(60, 100, 180) 'RGB(60, 100, 180) 'RGB(60, 100, 180)
CheckBoxRapor3VarlikEsas.ForeColor = RGB(255, 255, 255)
End Sub
Private Sub CheckBoxRapor1VarlikEsas_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
CheckBoxRapor1VarlikEsas.BackColor = RGB(60, 100, 180) 'RGB(60, 100, 180) 'RGB(60, 100, 180)
CheckBoxRapor1VarlikEsas.ForeColor = RGB(255, 255, 255)
End Sub


Private Sub CheckBoxTumVarliklar_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
CheckBoxTumVarliklar.BackColor = RGB(60, 100, 180) 'RGB(60, 100, 180) 'RGB(60, 100, 180)
CheckBoxTumVarliklar.ForeColor = RGB(255, 255, 255)
End Sub
Private Sub CheckBoxGenelVarlik_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
CheckBoxGenelVarlik.BackColor = RGB(60, 100, 180) 'RGB(60, 100, 180) 'RGB(60, 100, 180)
CheckBoxGenelVarlik.ForeColor = RGB(255, 255, 255)
End Sub
Private Sub CheckBoxRaporVarlik_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
CheckBoxRaporVarlik.BackColor = RGB(60, 100, 180) 'RGB(60, 100, 180) 'RGB(60, 100, 180)
CheckBoxRaporVarlik.ForeColor = RGB(255, 255, 255)
End Sub
Private Sub CheckBoxRapor2_2Varlik_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
CheckBoxRapor2_2Varlik.BackColor = RGB(60, 100, 180) 'RGB(60, 100, 180) 'RGB(60, 100, 180)
CheckBoxRapor2_2Varlik.ForeColor = RGB(255, 255, 255)
End Sub
Private Sub CheckBoxRapor3Varlik_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
CheckBoxRapor3Varlik.BackColor = RGB(60, 100, 180) 'RGB(60, 100, 180) 'RGB(60, 100, 180)
CheckBoxRapor3Varlik.ForeColor = RGB(255, 255, 255)
End Sub
Private Sub CheckBoxRapor1Varlik_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
CheckBoxRapor1Varlik.BackColor = RGB(60, 100, 180) 'RGB(60, 100, 180) 'RGB(60, 100, 180)
CheckBoxRapor1Varlik.ForeColor = RGB(255, 255, 255)
End Sub


''''''''''''''''''''Esas Varliklar

Private Sub CheckBoxTumVarliklarEsas_Click()
If CheckBoxTumVarliklarEsas.Value = True Then
    CheckBoxGenelVarlikEsas.Value = True
    CheckBoxRaporVarlikEsas.Value = True
    CheckBoxRapor2_2VarlikEsas.Value = True
    CheckBoxRapor3VarlikEsas.Value = True
    CheckBoxRapor1VarlikEsas.Value = True
ElseIf CheckBoxTumVarliklarEsas.Value = False Then
    CheckBoxGenelVarlikEsas.Value = False
    CheckBoxRaporVarlikEsas.Value = False
    CheckBoxRapor2_2VarlikEsas.Value = False
    CheckBoxRapor3VarlikEsas.Value = False
    CheckBoxRapor1VarlikEsas.Value = False
End If
End Sub

Private Sub CheckBoxGenelVarlikEsas_Click()
Dim Genel As Boolean, Rapor As Boolean, Rapor2_2 As Boolean, Rapor3 As Boolean, Rapor1 As Boolean

Rapor = CheckBoxRaporVarlikEsas.Value
Rapor2_2 = CheckBoxRapor2_2VarlikEsas.Value
Rapor3 = CheckBoxRapor3VarlikEsas.Value
Rapor1 = CheckBoxRapor1VarlikEsas.Value

If CheckBoxGenelVarlikEsas.Value = True Then
    If CheckBoxRaporVarlikEsas.Value = True And CheckBoxRapor2_2VarlikEsas.Value = True And _
        CheckBoxRapor3VarlikEsas.Value = True And CheckBoxRapor1VarlikEsas.Value = True Then
        If CheckBoxTumVarliklarEsas.Value = False Then
            CheckBoxTumVarliklarEsas.Value = True
        End If
    End If
ElseIf CheckBoxGenelVarlikEsas.Value = False Then
    If CheckBoxTumVarliklarEsas.Value = True Then
        CheckBoxTumVarliklarEsas.Value = False
    End If
End If

CheckBoxRaporVarlikEsas.Value = Rapor
CheckBoxRapor2_2VarlikEsas.Value = Rapor2_2
CheckBoxRapor3VarlikEsas.Value = Rapor3
CheckBoxRapor1VarlikEsas.Value = Rapor1

End Sub

Private Sub CheckBoxRaporVarlikEsas_Click()
Dim Genel As Boolean, Rapor As Boolean, Rapor2_2 As Boolean, Rapor3 As Boolean, Rapor1 As Boolean

Genel = CheckBoxGenelVarlikEsas.Value
Rapor2_2 = CheckBoxRapor2_2VarlikEsas.Value
Rapor3 = CheckBoxRapor3VarlikEsas.Value
Rapor1 = CheckBoxRapor1VarlikEsas.Value

If CheckBoxRaporVarlikEsas.Value = True Then
    If CheckBoxGenelVarlikEsas.Value = True And CheckBoxRapor2_2VarlikEsas.Value = True And _
        CheckBoxRapor3VarlikEsas.Value = True And CheckBoxRapor1VarlikEsas.Value = True Then
        If CheckBoxTumVarliklarEsas.Value = False Then
            CheckBoxTumVarliklarEsas.Value = True
        End If
    End If
ElseIf CheckBoxRaporVarlikEsas.Value = False Then
    If CheckBoxTumVarliklarEsas.Value = True Then
        CheckBoxTumVarliklarEsas.Value = False
    End If
End If

CheckBoxGenelVarlikEsas.Value = Genel
CheckBoxRapor2_2VarlikEsas.Value = Rapor2_2
CheckBoxRapor3VarlikEsas.Value = Rapor3
CheckBoxRapor1VarlikEsas.Value = Rapor1

End Sub

Private Sub CheckBoxRapor2_2VarlikEsas_Click()
Dim Genel As Boolean, Rapor As Boolean, Rapor2_2 As Boolean, Rapor3 As Boolean, Rapor1 As Boolean

Genel = CheckBoxGenelVarlikEsas.Value
Rapor = CheckBoxRaporVarlikEsas.Value
Rapor3 = CheckBoxRapor3VarlikEsas.Value
Rapor1 = CheckBoxRapor1VarlikEsas.Value

If CheckBoxRapor2_2VarlikEsas.Value = True Then
    If CheckBoxGenelVarlikEsas.Value = True And CheckBoxRaporVarlikEsas.Value = True And _
        CheckBoxRapor3VarlikEsas.Value = True And CheckBoxRapor1VarlikEsas.Value = True Then
        If CheckBoxTumVarliklarEsas.Value = False Then
            CheckBoxTumVarliklarEsas.Value = True
        End If
    End If
ElseIf CheckBoxRapor2_2VarlikEsas.Value = False Then
    If CheckBoxTumVarliklarEsas.Value = True Then
        CheckBoxTumVarliklarEsas.Value = False
    End If
End If

CheckBoxGenelVarlikEsas.Value = Genel
CheckBoxRaporVarlikEsas.Value = Rapor
CheckBoxRapor3VarlikEsas.Value = Rapor3
CheckBoxRapor1VarlikEsas.Value = Rapor1

End Sub

Private Sub CheckBoxRapor3VarlikEsas_Click()
Dim Genel As Boolean, Rapor As Boolean, Rapor2_2 As Boolean, Rapor3 As Boolean, Rapor1 As Boolean

Genel = CheckBoxGenelVarlikEsas.Value
Rapor = CheckBoxRaporVarlikEsas.Value
Rapor2_2 = CheckBoxRapor2_2VarlikEsas.Value
Rapor1 = CheckBoxRapor1VarlikEsas.Value

If CheckBoxRapor3VarlikEsas.Value = True Then
    If CheckBoxGenelVarlikEsas.Value = True And CheckBoxRaporVarlikEsas.Value = True And _
        CheckBoxRapor2_2VarlikEsas.Value = True And CheckBoxRapor1VarlikEsas.Value = True Then
        If CheckBoxTumVarliklarEsas.Value = False Then
            CheckBoxTumVarliklarEsas.Value = True
        End If
    End If
ElseIf CheckBoxRapor3VarlikEsas.Value = False Then
    If CheckBoxTumVarliklarEsas.Value = True Then
        CheckBoxTumVarliklarEsas.Value = False
    End If
End If

CheckBoxGenelVarlikEsas.Value = Genel
CheckBoxRaporVarlikEsas.Value = Rapor
CheckBoxRapor2_2VarlikEsas.Value = Rapor2_2
CheckBoxRapor1VarlikEsas.Value = Rapor1

End Sub

Private Sub CheckBoxRapor1VarlikEsas_Click()
Dim Genel As Boolean, Rapor As Boolean, Rapor2_2 As Boolean, Rapor3 As Boolean, Rapor1 As Boolean

Genel = CheckBoxGenelVarlikEsas.Value
Rapor = CheckBoxRaporVarlikEsas.Value
Rapor2_2 = CheckBoxRapor2_2VarlikEsas.Value
Rapor3 = CheckBoxRapor3VarlikEsas.Value

If CheckBoxRapor1VarlikEsas.Value = True Then
    If CheckBoxGenelVarlikEsas.Value = True And CheckBoxRaporVarlikEsas.Value = True And _
        CheckBoxRapor2_2VarlikEsas.Value = True And CheckBoxRapor3VarlikEsas.Value = True Then
        If CheckBoxTumVarliklarEsas.Value = False Then
            CheckBoxTumVarliklarEsas.Value = True
        End If
    End If
ElseIf CheckBoxRapor1VarlikEsas.Value = False Then
    If CheckBoxTumVarliklarEsas.Value = True Then
        CheckBoxTumVarliklarEsas.Value = False
    End If
End If

CheckBoxGenelVarlikEsas.Value = Genel
CheckBoxRaporVarlikEsas.Value = Rapor
CheckBoxRapor2_2VarlikEsas.Value = Rapor2_2
CheckBoxRapor3VarlikEsas.Value = Rapor3

End Sub

''''''''''''''''''''Yardımcı Varliklar

Private Sub CheckBoxTumVarliklar_Click()
If CheckBoxTumVarliklar.Value = True Then
    CheckBoxGenelVarlik.Value = True
    CheckBoxRaporVarlik.Value = True
    CheckBoxRapor2_2Varlik.Value = True
    CheckBoxRapor3Varlik.Value = True
    CheckBoxRapor1Varlik.Value = True
ElseIf CheckBoxTumVarliklar.Value = False Then
    CheckBoxGenelVarlik.Value = False
    CheckBoxRaporVarlik.Value = False
    CheckBoxRapor2_2Varlik.Value = False
    CheckBoxRapor3Varlik.Value = False
    CheckBoxRapor1Varlik.Value = False
End If
End Sub

Private Sub CheckBoxGenelVarlik_Click()
Dim Genel As Boolean, Rapor As Boolean, Rapor2_2 As Boolean, Rapor3 As Boolean, Rapor1 As Boolean

Rapor = CheckBoxRaporVarlik.Value
Rapor2_2 = CheckBoxRapor2_2Varlik.Value
Rapor3 = CheckBoxRapor3Varlik.Value
Rapor1 = CheckBoxRapor1Varlik.Value

If CheckBoxGenelVarlik.Value = True Then
    If CheckBoxRaporVarlik.Value = True And CheckBoxRapor2_2Varlik.Value = True And _
        CheckBoxRapor3Varlik.Value = True And CheckBoxRapor1Varlik.Value = True Then
        If CheckBoxTumVarliklar.Value = False Then
            CheckBoxTumVarliklar.Value = True
        End If
    End If
ElseIf CheckBoxGenelVarlik.Value = False Then
    If CheckBoxTumVarliklar.Value = True Then
        CheckBoxTumVarliklar.Value = False
    End If
End If

CheckBoxRaporVarlik.Value = Rapor
CheckBoxRapor2_2Varlik.Value = Rapor2_2
CheckBoxRapor3Varlik.Value = Rapor3
CheckBoxRapor1Varlik.Value = Rapor1
    
End Sub

Private Sub CheckBoxRaporVarlik_Click()
Dim Genel As Boolean, Rapor As Boolean, Rapor2_2 As Boolean, Rapor3 As Boolean, Rapor1 As Boolean

Genel = CheckBoxGenelVarlik.Value
Rapor2_2 = CheckBoxRapor2_2Varlik.Value
Rapor3 = CheckBoxRapor3Varlik.Value
Rapor1 = CheckBoxRapor1Varlik.Value

If CheckBoxRaporVarlik.Value = True Then
    If CheckBoxGenelVarlik.Value = True And CheckBoxRapor2_2Varlik.Value = True And _
        CheckBoxRapor3Varlik.Value = True And CheckBoxRapor1Varlik.Value = True Then
        If CheckBoxTumVarliklar.Value = False Then
            CheckBoxTumVarliklar.Value = True
        End If
    End If
ElseIf CheckBoxRaporVarlik.Value = False Then
    If CheckBoxTumVarliklar.Value = True Then
        CheckBoxTumVarliklar.Value = False
    End If
End If

CheckBoxGenelVarlik.Value = Genel
CheckBoxRapor2_2Varlik.Value = Rapor2_2
CheckBoxRapor3Varlik.Value = Rapor3
CheckBoxRapor1Varlik.Value = Rapor1
    
End Sub

Private Sub CheckBoxRapor2_2Varlik_Click()
Dim Genel As Boolean, Rapor As Boolean, Rapor2_2 As Boolean, Rapor3 As Boolean, Rapor1 As Boolean

Genel = CheckBoxGenelVarlik.Value
Rapor = CheckBoxRaporVarlik.Value
Rapor3 = CheckBoxRapor3Varlik.Value
Rapor1 = CheckBoxRapor1Varlik.Value

If CheckBoxRapor2_2Varlik.Value = True Then
    If CheckBoxGenelVarlik.Value = True And CheckBoxRaporVarlik.Value = True And _
        CheckBoxRapor3Varlik.Value = True And CheckBoxRapor1Varlik.Value = True Then
        If CheckBoxTumVarliklar.Value = False Then
            CheckBoxTumVarliklar.Value = True
        End If
    End If
ElseIf CheckBoxRapor2_2Varlik.Value = False Then
    If CheckBoxTumVarliklar.Value = True Then
        CheckBoxTumVarliklar.Value = False
    End If
End If

CheckBoxGenelVarlik.Value = Genel
CheckBoxRaporVarlik.Value = Rapor
CheckBoxRapor3Varlik.Value = Rapor3
CheckBoxRapor1Varlik.Value = Rapor1

End Sub

Private Sub CheckBoxRapor3Varlik_Click()
Dim Genel As Boolean, Rapor As Boolean, Rapor2_2 As Boolean, Rapor3 As Boolean, Rapor1 As Boolean

Genel = CheckBoxGenelVarlik.Value
Rapor = CheckBoxRaporVarlik.Value
Rapor2_2 = CheckBoxRapor2_2Varlik.Value
Rapor1 = CheckBoxRapor1Varlik.Value

If CheckBoxRapor3Varlik.Value = True Then
    If CheckBoxGenelVarlik.Value = True And CheckBoxRaporVarlik.Value = True And _
        CheckBoxRapor2_2Varlik.Value = True And CheckBoxRapor1Varlik.Value = True Then
        If CheckBoxTumVarliklar.Value = False Then
            CheckBoxTumVarliklar.Value = True
        End If
    End If
ElseIf CheckBoxRapor3Varlik.Value = False Then
    If CheckBoxTumVarliklar.Value = True Then
        CheckBoxTumVarliklar.Value = False
    End If
End If

CheckBoxGenelVarlik.Value = Genel
CheckBoxRaporVarlik.Value = Rapor
CheckBoxRapor2_2Varlik.Value = Rapor2_2
CheckBoxRapor1Varlik.Value = Rapor1

End Sub

Private Sub CheckBoxRapor1Varlik_Click()
Dim Genel As Boolean, Rapor As Boolean, Rapor2_2 As Boolean, Rapor3 As Boolean, Rapor1 As Boolean

Genel = CheckBoxGenelVarlik.Value
Rapor = CheckBoxRaporVarlik.Value
Rapor2_2 = CheckBoxRapor2_2Varlik.Value
Rapor3 = CheckBoxRapor3Varlik.Value

If CheckBoxRapor1Varlik.Value = True Then
    If CheckBoxGenelVarlik.Value = True And CheckBoxRaporVarlik.Value = True And _
        CheckBoxRapor2_2Varlik.Value = True And CheckBoxRapor3Varlik.Value = True Then
        If CheckBoxTumVarliklar.Value = False Then
            CheckBoxTumVarliklar.Value = True
        End If
    End If
ElseIf CheckBoxRapor1Varlik.Value = False Then
    If CheckBoxTumVarliklar.Value = True Then
        CheckBoxTumVarliklar.Value = False
    End If
End If

CheckBoxGenelVarlik.Value = Genel
CheckBoxRaporVarlik.Value = Rapor
CheckBoxRapor2_2Varlik.Value = Rapor2_2
CheckBoxRapor3Varlik.Value = Rapor3

End Sub

Private Sub VarlikTarihiText_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'suppport_calendar_UI
suppport_calendar_UI.Show

If CalTarih = "" Then
    GoTo Son
Else
    VarlikTarihiText.Value = CalTarih
    VarlikTarihiText.Value = Format(VarlikTarihiText.Value, "dd.mm.yyyy")
End If

Son:
CalTarih = ""
Cancel = True

End Sub

Private Sub VarlikTarihiText_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    
    On Error Resume Next
    'Delete ve Backspace tuşları textboxu sil.
    If KeyCode = vbKeyDelete Then
        VarlikTarihiText.Value = ""
    End If
    If KeyCode = vbKeyBack Then
        VarlikTarihiText.Value = ""
    End If

End Sub

Private Sub VarlikTarihiLabel_Click()

'suppport_calendar_UI
suppport_calendar_UI.Show

If CalTarih = "" Then
    GoTo Son
Else
    VarlikTarihiText.Value = CalTarih
    VarlikTarihiText.Value = Format(VarlikTarihiText.Value, "dd.mm.yyyy")
End If

Call VerileriGetir

Son:
CalTarih = ""

End Sub

Private Sub VarlikImza1EkleKaldirLabel_Click()
support_signatures_UI.Show
End Sub
Private Sub VarlikImza2EkleKaldirLabel_Click()
support_signatures_UI.Show
End Sub
Private Sub VarlikImza3EkleKaldirLabel_Click()
support_signatures_UI.Show
End Sub

Private Sub VarlikImza1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next

   Select Case KeyCode
        Case 38  'Yukarı
            If VarlikImza1.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                VarlikImza1.ListIndex = VarlikImza1.ListIndex - 1
            End If
            Me.VarlikImza1.DropDown
            
        Case 40 'Aşağı
            If VarlikImza1.ListIndex = VarlikImza1.ListCount - 1 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                VarlikImza1.ListIndex = VarlikImza1.ListIndex + 1
            End If
            Me.VarlikImza1.DropDown
    End Select
    Abort = False
    
End Sub

Private Sub VarlikImza1_Change()

If VarlikImza1.ListIndex = -1 And VarlikImza1.Value <> "" Then
   VarlikImza1.Value = ""
   GoTo Son
End If

If VarlikImza1.Value <> "" Then
    VarlikImza1.SelStart = 0
    VarlikImza1.SelLength = Len(VarlikImza1.Value)
End If


Son:

VarlikImza1.DropDown

End Sub

Private Sub VarlikImza2_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next

   Select Case KeyCode
        Case 38  'Yukarı
            If VarlikImza2.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                VarlikImza2.ListIndex = VarlikImza2.ListIndex - 1
            End If
            Me.VarlikImza2.DropDown
            
        Case 40 'Aşağı
            If VarlikImza2.ListIndex = VarlikImza2.ListCount - 1 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                VarlikImza2.ListIndex = VarlikImza2.ListIndex + 1
            End If
            Me.VarlikImza2.DropDown
    End Select
    Abort = False
    
End Sub

Private Sub VarlikImza2_Change()

If VarlikImza2.ListIndex = -1 And VarlikImza2.Value <> "" Then
   VarlikImza2.Value = ""
   GoTo Son
End If

If VarlikImza2.Value <> "" Then
    VarlikImza2.SelStart = 0
    VarlikImza2.SelLength = Len(VarlikImza2.Value)
End If


Son:

VarlikImza2.DropDown

End Sub

Private Sub VarlikImza3_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next

   Select Case KeyCode
        Case 38  'Yukarı
            If VarlikImza3.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                VarlikImza3.ListIndex = VarlikImza3.ListIndex - 1
            End If
            Me.VarlikImza3.DropDown
            
        Case 40 'Aşağı
            If VarlikImza3.ListIndex = VarlikImza3.ListCount - 1 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                VarlikImza3.ListIndex = VarlikImza3.ListIndex + 1
            End If
            Me.VarlikImza3.DropDown
    End Select
    Abort = False
    
End Sub

Private Sub VarlikImza3_Change()

If VarlikImza3.ListIndex = -1 And VarlikImza3.Value <> "" Then
   VarlikImza3.Value = ""
   GoTo Son
End If

If VarlikImza3.Value <> "" Then
    VarlikImza3.SelStart = 0
    VarlikImza3.SelLength = Len(VarlikImza3.Value)
End If


Son:

VarlikImza3.DropDown

End Sub

Private Sub VarlikImza1EkleKaldirLabel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
VarlikImza1EkleKaldirLabel.BackColor = RGB(60, 100, 180)
VarlikImza1EkleKaldirLabel.ForeColor = RGB(255, 255, 255)
End Sub
Private Sub VarlikImza1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
Call SetComboBoxHook(VarlikImza1) 'Open scrollable with mouse
End Sub

Private Sub LblVarlikImza1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub

Private Sub VarlikImza2EkleKaldirLabel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
VarlikImza2EkleKaldirLabel.BackColor = RGB(60, 100, 180)
VarlikImza2EkleKaldirLabel.ForeColor = RGB(255, 255, 255)
End Sub
Private Sub VarlikImza2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
Call SetComboBoxHook(VarlikImza2) 'Open scrollable with mouse
End Sub
Private Sub LblVarlikImza2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub

Private Sub VarlikImza3EkleKaldirLabel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
VarlikImza3EkleKaldirLabel.BackColor = RGB(60, 100, 180)
VarlikImza3EkleKaldirLabel.ForeColor = RGB(255, 255, 255)
End Sub
Private Sub VarlikImza3_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
Call SetComboBoxHook(VarlikImza3) 'Open scrollable with mouse
End Sub
Private Sub LblVarlikImza3_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub

Private Sub VarlikTarihiLabel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
VarlikTarihiLabel.BackColor = RGB(60, 100, 180) 'RGB(60, 100, 180) 'RGB(60, 100, 180)
VarlikTarihiLabel.ForeColor = RGB(255, 255, 255)
End Sub
Private Sub VarlikTarihiText_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub

Private Sub LblVarlikTarihi_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub

Private Sub CheckBox1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
CheckBox1.BackColor = RGB(60, 100, 180) 'RGB(60, 100, 180) 'RGB(60, 100, 180)
CheckBox1.ForeColor = RGB(255, 255, 255)
End Sub
Private Sub CheckBox2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
CheckBox2.BackColor = RGB(60, 100, 180) 'RGB(60, 100, 180) 'RGB(60, 100, 180)
CheckBox2.ForeColor = RGB(255, 255, 255)
End Sub
Private Sub CheckBox3_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
CheckBox3.BackColor = RGB(60, 100, 180) 'RGB(60, 100, 180) 'RGB(60, 100, 180)
CheckBox3.ForeColor = RGB(255, 255, 255)
End Sub

Private Sub LabelEkleGiris_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
LabelEkleGiris.BackColor = RGB(60, 100, 180) 'RGB(60, 100, 180) 'RGB(60, 100, 180)
LabelEkleGiris.ForeColor = RGB(255, 255, 255)
End Sub
Private Sub LabelKaldirGiris_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
LabelKaldirGiris.BackColor = RGB(60, 100, 180) 'RGB(60, 100, 180) 'RGB(60, 100, 180)
LabelKaldirGiris.ForeColor = RGB(255, 255, 255)
End Sub

Private Sub LabelEkleCikis_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
LabelEkleCikis.BackColor = RGB(60, 100, 180) 'RGB(60, 100, 180) 'RGB(60, 100, 180)
LabelEkleCikis.ForeColor = RGB(255, 255, 255)
End Sub

Private Sub LabelKaldirCikis_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
LabelKaldirCikis.BackColor = RGB(60, 100, 180) 'RGB(60, 100, 180) 'RGB(60, 100, 180)
LabelKaldirCikis.ForeColor = RGB(255, 255, 255)
End Sub

Private Sub FrameGiris_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)

Call ColorChangerGenel
Call RemoveScrollHook
If ScrollTakip1 > 180 Then
    Call SetScrollHook(Me.FrameGiris, ScrollTakip1)
Else
    FrameGiris.ScrollTop = 0
    RemoveScrollHook
    FrameGiris.ScrollBars = fmScrollBarsNone
End If

End Sub
Private Sub FrameMevcut_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)

Call ColorChangerGenel
Call RemoveScrollHook
If ScrollTakip2 > 180 Then
    Call SetScrollHook(Me.FrameMevcut, ScrollTakip2)
Else
    FrameMevcut.ScrollTop = 0
    RemoveScrollHook
    FrameMevcut.ScrollBars = fmScrollBarsNone
End If

End Sub
Private Sub FrameCikis_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)

Call ColorChangerGenel
Call RemoveScrollHook
If ScrollTakip3 > 180 Then
    Call SetScrollHook(Me.FrameCikis, ScrollTakip3)
Else
    FrameCikis.ScrollTop = 0
    RemoveScrollHook
    FrameCikis.ScrollBars = fmScrollBarsNone
End If

End Sub

Private Sub DirektYazdir_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
DirektYazdir.BackColor = RGB(60, 100, 180) 'RGB(60, 100, 180) 'RGB(60, 100, 180)
DirektYazdir.ForeColor = RGB(255, 255, 255)
End Sub

Private Sub Yardim_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
Yardim.BackColor = RGB(60, 100, 180) 'RGB(60, 100, 180) 'RGB(60, 100, 180)
Yardim.ForeColor = RGB(255, 255, 255)
End Sub

Private Sub Tutanak_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
Tutanak.BackColor = RGB(60, 100, 180) 'RGB(60, 100, 180) 'RGB(60, 100, 180)
Tutanak.ForeColor = RGB(255, 255, 255)
End Sub
Private Sub Kapat_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
Kapat.BackColor = RGB(60, 100, 180) 'RGB(60, 100, 180) 'RGB(60, 100, 180)
Kapat.ForeColor = RGB(255, 255, 255)
End Sub

Private Sub FrameEsasVarlikRaporlari_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub
Private Sub FrameYardimciVarlikRaporlari_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub
Private Sub FrameVarlikIslemleri_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub


Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
Call RemoveScrollHook
End Sub
Private Sub TasiyiciFrame_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
Call RemoveScrollHook
End Sub
Private Sub BaslikFrame_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub
Private Sub UstMenuFrame_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub
Private Sub AltMenuFrame_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub

Private Sub FrameGirisx_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
Call RemoveScrollHook
End Sub
Private Sub FrameCikisx_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
Call RemoveScrollHook
End Sub
Private Sub FrameMevcutx_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
Call RemoveScrollHook
End Sub
Private Sub FrameAlt1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
Call RemoveScrollHook
End Sub
Private Sub FrameAlt2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
Call RemoveScrollHook
End Sub
Private Sub FrameAlt3_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
Call RemoveScrollHook
End Sub


Sub ColorChangerGenel()

If DirektYazdir.BackColor <> RGB(225, 235, 245) Then
DirektYazdir.BackColor = RGB(225, 235, 245)
DirektYazdir.ForeColor = RGB(30, 30, 30)
End If
If Tutanak.BackColor <> RGB(225, 235, 245) Then
Tutanak.BackColor = RGB(225, 235, 245)
Tutanak.ForeColor = RGB(30, 30, 30)
End If
If Yardim.BackColor <> RGB(225, 235, 245) Then
Yardim.BackColor = RGB(225, 235, 245)
Yardim.ForeColor = RGB(30, 30, 30)
End If
If Kapat.BackColor <> RGB(225, 235, 245) Then
Kapat.BackColor = RGB(225, 235, 245)
Kapat.ForeColor = RGB(30, 30, 30)
End If


If VarlikImza1EkleKaldirLabel.BackColor <> RGB(254, 254, 254) Then
VarlikImza1EkleKaldirLabel.BackColor = RGB(254, 254, 254)
VarlikImza1EkleKaldirLabel.ForeColor = RGB(70, 70, 70)
End If
If VarlikImza2EkleKaldirLabel.BackColor <> RGB(254, 254, 254) Then
VarlikImza2EkleKaldirLabel.BackColor = RGB(254, 254, 254)
VarlikImza2EkleKaldirLabel.ForeColor = RGB(70, 70, 70)
End If
If VarlikImza3EkleKaldirLabel.BackColor <> RGB(254, 254, 254) Then
VarlikImza3EkleKaldirLabel.BackColor = RGB(254, 254, 254)
VarlikImza3EkleKaldirLabel.ForeColor = RGB(70, 70, 70)
End If

If CheckBox1.BackColor <> RGB(254, 254, 254) Then
CheckBox1.BackColor = RGB(254, 254, 254)
CheckBox1.ForeColor = RGB(70, 70, 70)
End If
If CheckBox2.BackColor <> RGB(254, 254, 254) Then
CheckBox2.BackColor = RGB(254, 254, 254) 'RGB(254, 254, 254)
CheckBox2.ForeColor = RGB(70, 70, 70)
End If
If CheckBox3.BackColor <> RGB(254, 254, 254) Then
CheckBox3.BackColor = RGB(254, 254, 254) 'RGB(254, 254, 254)
CheckBox3.ForeColor = RGB(70, 70, 70)
End If

If VarlikTarihiLabel.BackColor <> RGB(254, 254, 254) Then
VarlikTarihiLabel.BackColor = RGB(254, 254, 254)
VarlikTarihiLabel.ForeColor = RGB(70, 70, 70)
End If

If LabelEkleGiris.BackColor <> RGB(254, 254, 254) Then
LabelEkleGiris.BackColor = RGB(254, 254, 254)
LabelEkleGiris.ForeColor = RGB(70, 70, 70)
End If
If LabelEkleCikis.BackColor <> RGB(254, 254, 254) Then
LabelEkleCikis.BackColor = RGB(254, 254, 254)
LabelEkleCikis.ForeColor = RGB(70, 70, 70)
End If
If LabelKaldirGiris.BackColor <> RGB(254, 254, 254) Then
LabelKaldirGiris.BackColor = RGB(254, 254, 254)
LabelKaldirGiris.ForeColor = RGB(70, 70, 70)
End If
If LabelKaldirCikis.BackColor <> RGB(254, 254, 254) Then
LabelKaldirCikis.BackColor = RGB(254, 254, 254)
LabelKaldirCikis.ForeColor = RGB(70, 70, 70)
End If

If CheckBoxTumVarliklarEsas.BackColor <> RGB(254, 254, 254) Then
CheckBoxTumVarliklarEsas.BackColor = RGB(254, 254, 254)
CheckBoxTumVarliklarEsas.ForeColor = RGB(70, 70, 70)
End If
If CheckBoxGenelVarlikEsas.BackColor <> RGB(254, 254, 254) Then
CheckBoxGenelVarlikEsas.BackColor = RGB(254, 254, 254)
CheckBoxGenelVarlikEsas.ForeColor = RGB(70, 70, 70)
End If
If CheckBoxRaporVarlikEsas.BackColor <> RGB(254, 254, 254) Then
CheckBoxRaporVarlikEsas.BackColor = RGB(254, 254, 254)
CheckBoxRaporVarlikEsas.ForeColor = RGB(70, 70, 70)
End If
If CheckBoxRapor2_2VarlikEsas.BackColor <> RGB(254, 254, 254) Then
CheckBoxRapor2_2VarlikEsas.BackColor = RGB(254, 254, 254)
CheckBoxRapor2_2VarlikEsas.ForeColor = RGB(70, 70, 70)
End If
If CheckBoxRapor3VarlikEsas.BackColor <> RGB(254, 254, 254) Then
CheckBoxRapor3VarlikEsas.BackColor = RGB(254, 254, 254)
CheckBoxRapor3VarlikEsas.ForeColor = RGB(70, 70, 70)
End If
If CheckBoxRapor1VarlikEsas.BackColor <> RGB(254, 254, 254) Then
CheckBoxRapor1VarlikEsas.BackColor = RGB(254, 254, 254)
CheckBoxRapor1VarlikEsas.ForeColor = RGB(70, 70, 70)
End If


If CheckBoxTumVarliklar.BackColor <> RGB(254, 254, 254) Then
CheckBoxTumVarliklar.BackColor = RGB(254, 254, 254)
CheckBoxTumVarliklar.ForeColor = RGB(70, 70, 70)
End If
If CheckBoxGenelVarlik.BackColor <> RGB(254, 254, 254) Then
CheckBoxGenelVarlik.BackColor = RGB(254, 254, 254)
CheckBoxGenelVarlik.ForeColor = RGB(70, 70, 70)
End If
If CheckBoxRaporVarlik.BackColor <> RGB(254, 254, 254) Then
CheckBoxRaporVarlik.BackColor = RGB(254, 254, 254)
CheckBoxRaporVarlik.ForeColor = RGB(70, 70, 70)
End If
If CheckBoxRapor2_2Varlik.BackColor <> RGB(254, 254, 254) Then
CheckBoxRapor2_2Varlik.BackColor = RGB(254, 254, 254)
CheckBoxRapor2_2Varlik.ForeColor = RGB(70, 70, 70)
End If
If CheckBoxRapor3Varlik.BackColor <> RGB(254, 254, 254) Then
CheckBoxRapor3Varlik.BackColor = RGB(254, 254, 254)
CheckBoxRapor3Varlik.ForeColor = RGB(70, 70, 70)
End If
If CheckBoxRapor1Varlik.BackColor <> RGB(254, 254, 254) Then
CheckBoxRapor1Varlik.BackColor = RGB(254, 254, 254)
CheckBoxRapor1Varlik.ForeColor = RGB(70, 70, 70)
End If

End Sub

Private Sub CheckBox1_Click()
Dim i As Long
Dim ctl As MSForms.Control
Dim Say1 As Integer, Say2 As Integer, Cont1 As Integer, Cont2 As Integer
Dim LstBx As MSForms.ListBox

'Tümünü seç
If CheckBox1.Value = True Then
    For Each ctl In core_asset_manager_UI.FrameGiris.Controls
        If TypeName(ctl) = "ListBox" Then
            ctl.Selected(0) = True
        End If
    Next ctl
Else 'Tümünü iptal et
    For Each ctl In core_asset_manager_UI.FrameGiris.Controls
        If TypeName(ctl) = "ListBox" Then
            ctl.Selected(0) = False
        End If
    Next ctl
End If

End Sub
Private Sub CheckBox2_Click()
Dim i As Long
Dim ctl As MSForms.Control
Dim Say1 As Integer, Say2 As Integer, Cont1 As Integer, Cont2 As Integer
Dim LstBx As MSForms.ListBox

'Tümünü seç
If CheckBox2.Value = True Then
    For Each ctl In core_asset_manager_UI.FrameMevcut.Controls
        If TypeName(ctl) = "ListBox" Then
            ctl.Selected(0) = True
        End If
    Next ctl
Else 'Tümünü iptal et
    For Each ctl In core_asset_manager_UI.FrameMevcut.Controls
        If TypeName(ctl) = "ListBox" Then
            ctl.Selected(0) = False
        End If
    Next ctl
End If

End Sub
Private Sub CheckBox3_Click()
Dim i As Long
Dim ctl As MSForms.Control
Dim Say1 As Integer, Say2 As Integer, Cont1 As Integer, Cont2 As Integer
Dim LstBx As MSForms.ListBox

'Tümünü seç
If CheckBox3.Value = True Then
    For Each ctl In core_asset_manager_UI.FrameCikis.Controls
        If TypeName(ctl) = "ListBox" Then
            ctl.Selected(0) = True
        End If
    Next ctl
Else 'Tümünü iptal et
    For Each ctl In core_asset_manager_UI.FrameCikis.Controls
        If TypeName(ctl) = "ListBox" Then
            ctl.Selected(0) = False
        End If
    Next ctl
End If

End Sub

Private Sub LabelEkleGiris_Click()
Dim i As Long
Dim ctl As MSForms.Control
Dim Say1 As Integer, Say2 As Integer, Cont1 As Integer, Cont2 As Integer
Dim LstBx As MSForms.ListBox
Dim LblSira1 As MSForms.Label
Dim LblSira2 As MSForms.Label
Dim IlkSira As Long, SiraBul As Range

Application.ScreenUpdating = False

ThisWorkbook.Unprotect "123"
Worksheets(4).Unprotect Password:="123"

IlkSira = 0
For Each ctl In core_asset_manager_UI.FrameGiris.Controls
    If TypeName(ctl) = "ListBox" Then
        If ctl.Selected(0) = True Then
            'MsgBox Mid(Ctl.List(0), InStr(Ctl.List(0), "|"), 9) & " : " & Len(Mid(Ctl.List(0), InStr(Ctl.List(0), "|"), 9))
            If Mid(ctl.List(0), InStr(ctl.List(0), "|"), 9) = "| B | <R>" Then
                Set SiraBul = ThisWorkbook.Worksheets(4).Range("E7:E100000").Find(What:=Left(ctl.List(0), InStr(ctl.List(0), "|") - 2), SearchDirection:=xlNext, _
                                SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
                If Not SiraBul Is Nothing Then
                    IlkSira = SiraBul.Row
                End If
            End If
        End If
    End If
Next ctl
Worksheets(4).Protect Password:="123" ', DrawingObjects:=False
ThisWorkbook.Protect "123"

'Aktarım için en az bir veri seçili olmalıdır.
Say1 = 0
For Each ctl In core_asset_manager_UI.FrameGiris.Controls
    If TypeName(ctl) = "ListBox" Then
        If ctl.Selected(0) = True Then
            Say1 = Say1 + 1
        End If
    End If
Next ctl
If Say1 = 0 Then
    GoTo Son
End If

Say2 = 0
For Each ctl In core_asset_manager_UI.FrameCikis.Controls
    If TypeName(ctl) = "ListBox" Then
        Say2 = Say2 + 1
    End If
    If TypeName(ctl) = "Label" Then
        FrameCikis.Controls.Remove ctl.name
    End If
Next ctl

Say1 = 0
For Each ctl In core_asset_manager_UI.FrameGiris.Controls
    If TypeName(ctl) = "ListBox" Then
        If ctl.Selected(0) = True Then

            If IlkSira <> 0 And Mid(ctl.List(0), InStr(ctl.List(0), "|"), 9) = "| B | <R>" Then
                If ThisWorkbook.Worksheets(4).Cells(IlkSira, 191).Value <> "" Then
                    Say1 = Say1 + 1
                    Set LstBx = FrameCikis.Controls.Add("Forms.ListBox.1")
                    With LstBx
                        .Top = (Say2 + Say1 - 1) * 12
                        .Left = 18
                        .Height = 12
                        .Width = 300
                        .BackColor = &H80000000
                        .SpecialEffect = fmSpecialEffectFlat
                        .MultiSelect = fmMultiSelectMulti 'fmMultiSelectSingle
                        .TextAlign = fmTextAlignLeft '02092021
                        .AddItem ctl.List(0)
                    End With
                    FrameGiris.Controls.Remove ctl.name
                    GoTo CTLDonguDevam1
                Else
                    Say1 = Say1 + 1
                    MsgBox """" & ctl.List(0) & """" & " cannot be completed because the item(s) sent to XXXMud due to a Technique A suspicion as part of this operation have not been detected as returned to the unit.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
                    GoTo CTLDonguDevam1
                End If
            End If

            Say1 = Say1 + 1
            Set LstBx = FrameCikis.Controls.Add("Forms.ListBox.1")
            With LstBx
                .Top = (Say2 + Say1 - 1) * 12
                .Left = 18
                .Height = 12
                .Width = 300
                .BackColor = &H80000000
                .SpecialEffect = fmSpecialEffectFlat
                .MultiSelect = fmMultiSelectMulti 'fmMultiSelectSingle
                .AddItem ctl.List(0)
            End With
            FrameGiris.Controls.Remove ctl.name
        End If
    End If
CTLDonguDevam1:
    'Sıra noları temizle
    If TypeName(ctl) = "Label" Then
        FrameGiris.Controls.Remove ctl.name
    End If
Next ctl
If Say1 = 0 Then
    GoTo Son
End If

ScrollTakip3 = ScrollTakip3 + Say1 * 12
ScrollTakip1 = ScrollTakip1 - Say1 * 12

If ScrollTakip3 > 180 Then
    Call SetScrollHook(Me.FrameCikis, ScrollTakip3)
    FrameCikis.ScrollTop = ScrollTakip3
Else
    FrameCikis.ScrollTop = 0
    RemoveScrollHook
    FrameCikis.ScrollBars = fmScrollBarsNone
End If

If ScrollTakip1 > 180 Then
    Call SetScrollHook(Me.FrameGiris, ScrollTakip1)
    FrameGiris.ScrollTop = FrameGiris.ScrollTop - Say1 * 12
Else
    FrameGiris.ScrollTop = 0
    RemoveScrollHook
    FrameGiris.ScrollBars = fmScrollBarsNone
End If

'FrameGiris tekrar sırala
Say1 = 0
For Each ctl In core_asset_manager_UI.FrameGiris.Controls
    If TypeName(ctl) = "ListBox" Then
        Say1 = Say1 + 1
        ctl.Top = (Say1 - 1) * 12
    End If

    Set LblSira1 = FrameGiris.Controls.Add("Forms.Label.1", "Lbl" & Say1)
    With LblSira1
        .Top = (Say1 - 1) * 12
        .Left = 0
        .Height = 12
        .Width = 18
        .SpecialEffect = fmSpecialEffectEtched
        .TextAlign = fmTextAlignCenter
        .Caption = Say1
    End With
Next ctl

'FrameCikis sira no ver.
Say2 = 0
For Each ctl In core_asset_manager_UI.FrameCikis.Controls
    If TypeName(ctl) = "ListBox" Then
        Say2 = Say2 + 1
        ctl.Top = (Say2 - 1) * 12
    End If

    Set LblSira2 = FrameCikis.Controls.Add("Forms.Label.1", "Lbl" & Say2)
    With LblSira2
        .Top = (Say2 - 1) * 12
        .Left = 0
        .Height = 12
        .Width = 18
        .SpecialEffect = fmSpecialEffectEtched
        .TextAlign = fmTextAlignCenter
        .Caption = Say2
    End With
Next ctl

Son:

CheckBox1.Value = False
CheckBox3.Value = False

Application.ScreenUpdating = True

End Sub

Private Sub LabelKaldirGiris_Click()
Dim ctl As MSForms.Control
Dim Say1 As Integer, Say2 As Integer, Cont1 As Integer, Cont2 As Integer
Dim LstBx As MSForms.ListBox, i As Integer
Dim LblSira1 As MSForms.Label
Dim LblSira2 As MSForms.Label

Application.ScreenUpdating = False

'Aktarım için en az bir veri seçili olmalıdır.
Say2 = 0
For Each ctl In core_asset_manager_UI.FrameCikis.Controls
    If TypeName(ctl) = "ListBox" Then
        If ctl.Selected(0) = True And ctl.BackColor = &H80000000 Then
            Say2 = Say2 + 1
        End If
    End If
Next ctl
If Say2 = 0 Then
    GoTo Son
End If

Say1 = 0
For Each ctl In core_asset_manager_UI.FrameGiris.Controls
    If TypeName(ctl) = "ListBox" Then
        Say1 = Say1 + 1
    End If
    If TypeName(ctl) = "Label" Then
        FrameGiris.Controls.Remove ctl.name
    End If
Next ctl

Say2 = 0
For Each ctl In core_asset_manager_UI.FrameCikis.Controls
    If TypeName(ctl) = "ListBox" Then
        If ctl.Selected(0) = True And ctl.BackColor = &H80000000 Then
            Say2 = Say2 + 1
            Set LstBx = FrameGiris.Controls.Add("Forms.ListBox.1")
            With LstBx
                .Top = (Say1 + Say2 - 1) * 12
                .Left = 18
                .Height = 12
                .Width = 300
                .BackColor = &H80000000
                .SpecialEffect = fmSpecialEffectFlat
                .MultiSelect = fmMultiSelectMulti 'fmMultiSelectSingle
                .TextAlign = fmTextAlignLeft '02092021
                .AddItem ctl.List(0)
            End With
            FrameCikis.Controls.Remove ctl.name
        End If
    End If
    'Sıra noları temizle
    If TypeName(ctl) = "Label" Then
        FrameCikis.Controls.Remove ctl.name
    End If
Next ctl
If Say2 = 0 Then
    GoTo Son
End If

ScrollTakip1 = ScrollTakip1 + Say2 * 12
ScrollTakip3 = ScrollTakip3 - Say2 * 12

If ScrollTakip1 > 180 Then
    Call SetScrollHook(Me.FrameGiris, ScrollTakip1)
    FrameGiris.ScrollTop = ScrollTakip1
Else
    FrameGiris.ScrollTop = 0
    RemoveScrollHook
    FrameGiris.ScrollBars = fmScrollBarsNone
End If

If ScrollTakip3 > 180 Then
    Call SetScrollHook(Me.FrameCikis, ScrollTakip3)
    FrameCikis.ScrollTop = FrameCikis.ScrollTop - Say2 * 12
Else
    FrameCikis.ScrollTop = 0
    RemoveScrollHook
    FrameCikis.ScrollBars = fmScrollBarsNone
End If

'FrameCikis tekrar sırala
Say2 = 0
For Each ctl In core_asset_manager_UI.FrameCikis.Controls
    If TypeName(ctl) = "ListBox" Then
        Say2 = Say2 + 1
        ctl.Top = (Say2 - 1) * 12
    End If
    
    Set LblSira2 = FrameCikis.Controls.Add("Forms.Label.1", "Lbl" & Say2)
    With LblSira2
        .Top = (Say2 - 1) * 12
        .Left = 0
        .Height = 12
        .Width = 18
        .SpecialEffect = fmSpecialEffectEtched
        .TextAlign = fmTextAlignCenter
        .Caption = Say2
    End With
Next ctl

'FrameGiris sira no ver.
Say1 = 0
For Each ctl In core_asset_manager_UI.FrameGiris.Controls
    If TypeName(ctl) = "ListBox" Then
        Say1 = Say1 + 1
        ctl.Top = (Say1 - 1) * 12
    End If
    
    Set LblSira1 = FrameGiris.Controls.Add("Forms.Label.1", "Lbl" & Say1)
    With LblSira1
        .Top = (Say1 - 1) * 12
        .Left = 0
        .Height = 12
        .Width = 18
        .SpecialEffect = fmSpecialEffectEtched
        .TextAlign = fmTextAlignCenter
        .Caption = Say1
    End With
Next ctl

Son:

Application.ScreenUpdating = True

End Sub

Private Sub LabelEkleCikis_Click()
Dim i As Long
Dim ctl As MSForms.Control
Dim Say1 As Integer, Say2 As Integer, Cont1 As Integer, Cont2 As Integer
Dim LstBx As MSForms.ListBox
Dim LblSira1 As MSForms.Label
Dim LblSira2 As MSForms.Label
Dim IlkSira As Long, SiraBul As Range

Application.ScreenUpdating = False

ThisWorkbook.Unprotect "123"
Worksheets(4).Unprotect Password:="123"

IlkSira = 0
For Each ctl In core_asset_manager_UI.FrameMevcut.Controls
    If TypeName(ctl) = "ListBox" Then
        If ctl.Selected(0) = True Then
            'MsgBox Mid(Ctl.List(0), InStr(Ctl.List(0), "|"), 9) & " : " & Len(Mid(Ctl.List(0), InStr(Ctl.List(0), "|"), 9))
            If Mid(ctl.List(0), InStr(ctl.List(0), "|"), 9) = "| B | <R>" Then
                Set SiraBul = ThisWorkbook.Worksheets(4).Range("E7:E100000").Find(What:=Left(ctl.List(0), InStr(ctl.List(0), "|") - 2), SearchDirection:=xlNext, _
                                SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
                If Not SiraBul Is Nothing Then
                    IlkSira = SiraBul.Row
                End If
            End If
        End If
    End If
Next ctl
Worksheets(4).Protect Password:="123" ', DrawingObjects:=False
ThisWorkbook.Protect "123"

'Aktarım için en az bir veri seçili olmalıdır.
Say1 = 0
For Each ctl In core_asset_manager_UI.FrameMevcut.Controls
    If TypeName(ctl) = "ListBox" Then
        If ctl.Selected(0) = True Then
            Say1 = Say1 + 1
        End If
    End If
Next ctl
If Say1 = 0 Then
    GoTo Son
End If

Say2 = 0
For Each ctl In core_asset_manager_UI.FrameCikis.Controls
    If TypeName(ctl) = "ListBox" Then
        Say2 = Say2 + 1
    End If
    If TypeName(ctl) = "Label" Then
        FrameCikis.Controls.Remove ctl.name
    End If
Next ctl

Say1 = 0
For Each ctl In core_asset_manager_UI.FrameMevcut.Controls
    If TypeName(ctl) = "ListBox" Then
        If ctl.Selected(0) = True Then

            If IlkSira <> 0 And Mid(ctl.List(0), InStr(ctl.List(0), "|"), 9) = "| B | <R>" Then
                If ThisWorkbook.Worksheets(4).Cells(IlkSira, 191).Value <> "" Then
                    Say1 = Say1 + 1
                    Set LstBx = FrameCikis.Controls.Add("Forms.ListBox.1")
                    With LstBx
                        .Top = (Say2 + Say1 - 1) * 12
                        .Left = 18
                        .Height = 12
                        .Width = 300
                        .BackColor = &H80000003
                        .SpecialEffect = fmSpecialEffectFlat
                        .MultiSelect = fmMultiSelectMulti 'fmMultiSelectSingle
                        .TextAlign = fmTextAlignLeft '02092021
                        .AddItem ctl.List(0)
                    End With
                    FrameMevcut.Controls.Remove ctl.name
                    GoTo CTLDonguDevam1
                Else
                    Say1 = Say1 + 1
                    MsgBox """" & ctl.List(0) & """" & " cannot be completed because the item(s) sent to XXXMud due to a Technique A suspicion as part of this operation have not been detected as returned to the unit.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
                    GoTo CTLDonguDevam1
                End If
            End If
            
            Say1 = Say1 + 1
            Set LstBx = FrameCikis.Controls.Add("Forms.ListBox.1")
            With LstBx
                .Top = (Say2 + Say1 - 1) * 12
                .Left = 18
                .Height = 12
                .Width = 300
                .BackColor = &H80000003
                .SpecialEffect = fmSpecialEffectFlat
                .MultiSelect = fmMultiSelectMulti 'fmMultiSelectSingle
                .TextAlign = fmTextAlignLeft '02092021
                .AddItem ctl.List(0)
            End With
            FrameMevcut.Controls.Remove ctl.name
        End If
    End If
CTLDonguDevam1:
    'Sıra noları temizle
    If TypeName(ctl) = "Label" Then
        FrameMevcut.Controls.Remove ctl.name
    End If
Next ctl
If Say1 = 0 Then
    GoTo Son
End If


ScrollTakip3 = ScrollTakip3 + Say1 * 12
ScrollTakip2 = ScrollTakip2 - Say1 * 12

If ScrollTakip3 > 180 Then
    Call SetScrollHook(Me.FrameCikis, ScrollTakip3)
    FrameCikis.ScrollTop = ScrollTakip3
Else
    FrameCikis.ScrollTop = 0
    RemoveScrollHook
    FrameCikis.ScrollBars = fmScrollBarsNone
End If

If ScrollTakip2 > 180 Then
    Call SetScrollHook(Me.FrameMevcut, ScrollTakip2)
    FrameMevcut.ScrollTop = FrameMevcut.ScrollTop - Say1 * 12
Else
    FrameMevcut.ScrollTop = 0
    RemoveScrollHook
    FrameMevcut.ScrollBars = fmScrollBarsNone
End If

'FrameMevcut tekrar sırala
Say1 = 0
For Each ctl In core_asset_manager_UI.FrameMevcut.Controls
    If TypeName(ctl) = "ListBox" Then
        Say1 = Say1 + 1
        ctl.Top = (Say1 - 1) * 12
    End If

    Set LblSira1 = FrameMevcut.Controls.Add("Forms.Label.1", "Lbl" & Say1)
    With LblSira1
        .Top = (Say1 - 1) * 12
        .Left = 0
        .Height = 12
        .Width = 18
        .SpecialEffect = fmSpecialEffectEtched
        .TextAlign = fmTextAlignCenter
        .Caption = Say1
    End With
Next ctl

'FrameCikis sira no ver.
Say2 = 0
For Each ctl In core_asset_manager_UI.FrameCikis.Controls
    If TypeName(ctl) = "ListBox" Then
        Say2 = Say2 + 1
        ctl.Top = (Say2 - 1) * 12
    End If

    Set LblSira2 = FrameCikis.Controls.Add("Forms.Label.1", "Lbl" & Say2)
    With LblSira2
        .Top = (Say2 - 1) * 12
        .Left = 0
        .Height = 12
        .Width = 18
        .SpecialEffect = fmSpecialEffectEtched
        .TextAlign = fmTextAlignCenter
        .Caption = Say2
    End With
Next ctl

Son:

CheckBox2.Value = False
CheckBox3.Value = False

Application.ScreenUpdating = True

End Sub

Private Sub LabelKaldirCikis_Click()
Dim ctl As MSForms.Control
Dim Say1 As Integer, Say2 As Integer, Cont1 As Integer, Cont2 As Integer
Dim LstBx As MSForms.ListBox, i As Integer
Dim LblSira1 As MSForms.Label
Dim LblSira2 As MSForms.Label

Application.ScreenUpdating = False

'Aktarım için en az bir veri seçili olmalıdır.
Say2 = 0
For Each ctl In core_asset_manager_UI.FrameCikis.Controls
    If TypeName(ctl) = "ListBox" Then
        If ctl.Selected(0) = True And ctl.BackColor = &H80000003 Then
            Say2 = Say2 + 1
        End If
    End If
Next ctl
If Say2 = 0 Then
    GoTo Son
End If

Say1 = 0
For Each ctl In core_asset_manager_UI.FrameMevcut.Controls
    If TypeName(ctl) = "ListBox" Then
        Say1 = Say1 + 1
    End If
    If TypeName(ctl) = "Label" Then
        FrameMevcut.Controls.Remove ctl.name
    End If
Next ctl

Say2 = 0
For Each ctl In core_asset_manager_UI.FrameCikis.Controls
    If TypeName(ctl) = "ListBox" Then
        If ctl.Selected(0) = True And ctl.BackColor = &H80000003 Then
            Say2 = Say2 + 1
            Set LstBx = FrameMevcut.Controls.Add("Forms.ListBox.1")
            With LstBx
                .Top = (Say1 + Say2 - 1) * 12
                .Left = 18
                .Height = 12
                .Width = 300
                .BackColor = &H80000003
                .SpecialEffect = fmSpecialEffectFlat
                .MultiSelect = fmMultiSelectMulti 'fmMultiSelectSingle
                .TextAlign = fmTextAlignLeft '02092021
                .AddItem ctl.List(0)
            End With
            FrameCikis.Controls.Remove ctl.name
        End If
    End If
    'Sıra noları temizle
    If TypeName(ctl) = "Label" Then
        FrameCikis.Controls.Remove ctl.name
    End If
Next ctl
If Say2 = 0 Then
    GoTo Son
End If

ScrollTakip2 = ScrollTakip2 + Say2 * 12
ScrollTakip3 = ScrollTakip3 - Say2 * 12

If ScrollTakip2 > 180 Then
    Call SetScrollHook(Me.FrameMevcut, ScrollTakip2)
    FrameMevcut.ScrollTop = ScrollTakip2
Else
    FrameMevcut.ScrollTop = 0
    RemoveScrollHook
    FrameMevcut.ScrollBars = fmScrollBarsNone
End If

If ScrollTakip3 > 180 Then
    Call SetScrollHook(Me.FrameCikis, ScrollTakip3)
    FrameCikis.ScrollTop = FrameCikis.ScrollTop - Say2 * 12
Else
    FrameCikis.ScrollTop = 0
    RemoveScrollHook
    FrameCikis.ScrollBars = fmScrollBarsNone
End If

'FrameCikis tekrar sırala
Say2 = 0
For Each ctl In core_asset_manager_UI.FrameCikis.Controls
    If TypeName(ctl) = "ListBox" Then
        Say2 = Say2 + 1
        ctl.Top = (Say2 - 1) * 12
    End If
    
    Set LblSira2 = FrameCikis.Controls.Add("Forms.Label.1", "Lbl" & Say2)
    With LblSira2
        .Top = (Say2 - 1) * 12
        .Left = 0
        .Height = 12
        .Width = 18
        .SpecialEffect = fmSpecialEffectEtched
        .TextAlign = fmTextAlignCenter
        .Caption = Say2
    End With
Next ctl

'FrameMevcut sira no ver.
Say1 = 0
For Each ctl In core_asset_manager_UI.FrameMevcut.Controls
    If TypeName(ctl) = "ListBox" Then
        Say1 = Say1 + 1
        ctl.Top = (Say1 - 1) * 12
    End If
    
    Set LblSira1 = FrameMevcut.Controls.Add("Forms.Label.1", "Lbl" & Say1)
    With LblSira1
        .Top = (Say1 - 1) * 12
        .Left = 0
        .Height = 12
        .Width = 18
        .SpecialEffect = fmSpecialEffectEtched
        .TextAlign = fmTextAlignCenter
        .Caption = Say1
    End With
Next ctl

Son:

Application.ScreenUpdating = True

End Sub

Private Sub Kapat_Click()
    Unload Me
End Sub

Private Sub VerileriGetir()
Dim i As Integer, TarihBul As Range
Dim Say As Long, j As Long, Cont As Long, Sno As Variant
Dim LstBx As MSForms.ListBox, ctl As MSForms.Control
Dim LblSira1 As MSForms.Label
Dim LblSira2 As MSForms.Label

ScrollTakip1 = 0
ScrollTakip2 = 0
ScrollTakip3 = 0

Application.ScreenUpdating = False

''ThisWorkbook.Worksheets(7).Visible = True
''ThisWorkbook.Worksheets(7).Activate


'suppport_calendar_UI
If VarlikTarihiText.Value = "" Then
    MsgBox "The operation cannot be initiated because the asset date has not been selected.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If
CalTarih = VarlikTarihiText.Value

'Verileri çağır başlatıldığında Frameleri boşalt
For Each ctl In core_asset_manager_UI.FrameGiris.Controls
    If TypeName(ctl) = "ListBox" Then
       FrameGiris.Controls.Remove ctl.name
    End If
    If TypeName(ctl) = "Label" Then
        FrameGiris.Controls.Remove ctl.name
    End If
Next ctl
For Each ctl In core_asset_manager_UI.FrameMevcut.Controls
    If TypeName(ctl) = "ListBox" Then
       FrameMevcut.Controls.Remove ctl.name
    End If
    If TypeName(ctl) = "Label" Then
        FrameMevcut.Controls.Remove ctl.name
    End If
Next ctl
For Each ctl In core_asset_manager_UI.FrameCikis.Controls
    If TypeName(ctl) = "ListBox" Then
       FrameCikis.Controls.Remove ctl.name
    End If
    If TypeName(ctl) = "Label" Then
        FrameCikis.Controls.Remove ctl.name
    End If
Next ctl


CalTarihTakip = CalTarih
ContTakipGiris = 0
ContTakipMevcut = 0
ContTakipCikis = 0

Call ModuleReport1.Rapor1VarlikHareketleriGiris
Call ModuleReport1.Rapor1VarlikHareketleriMevcut
Call ModuleReport1.Rapor1VarlikHareketleriCikis


Call ModuleReport2.Rapor2_1VarlikHareketleriGiris
Call ModuleReport2.Rapor2_1VarlikHareketleriMevcut
Call ModuleReport2.Rapor2_1VarlikHareketleriCikis

Call ModuleReport2.Rapor2_2VarlikHareketleriGiris
Call ModuleReport2.Rapor2_2VarlikHareketleriMevcut
Call ModuleReport2.Rapor2_2VarlikHareketleriCikis

Call ModuleReport2.Rapor2_2VarlikHareketleriGirisXXXMud
Call ModuleReport2.Rapor2_2VarlikHareketleriMevcutXXXMud
Call ModuleReport2.Rapor2_2VarlikHareketleriCikisXXXMud

Call ModuleReport3.Rapor3_2VarlikHareketleriGiris
Call ModuleReport3.Rapor3_2VarlikHareketleriMevcut
Call ModuleReport3.Rapor3_2VarlikHareketleriCikis

Call ModuleReport3.Rapor3_1VarlikHareketleriGiris
Call ModuleReport3.Rapor3_1VarlikHareketleriMevcut
Call ModuleReport3.Rapor3_1VarlikHareketleriCikis


If ScrollTakip1 > 180 Then
    Call SetScrollHook(Me.FrameGiris, ScrollTakip1)
    FrameGiris.ScrollTop = 0 'ScrollTakip1
End If
If ScrollTakip2 > 180 Then
    Call SetScrollHook(Me.FrameMevcut, ScrollTakip2)
    FrameMevcut.ScrollTop = 0 'ScrollTakip2
End If
If ScrollTakip3 > 180 Then
    Call SetScrollHook(Me.FrameCikis, ScrollTakip3)
    FrameCikis.ScrollTop = 0 'ScrollTakip3
End If

Son:

''ThisWorkbook.Worksheets(7).Visible = False
''Call ModuleRibbon.Rapor1 'Bu kod geçici daha sonra sheet1 olacak.

Application.ScreenUpdating = True


End Sub

Private Sub DirektYazdir_Click()
    
    core_asset_manager_UIDirektYazdir = True

    On Error GoTo Son
    core_asset_manager_UISayPrt = InputBox(Prompt:="Yazdırmak istediğiniz çıktı sayısını giriniz.", Title:="Enterprise Document Automation System")
    If core_asset_manager_UISayPrt = "" Or core_asset_manager_UISayPrt = 0 Then
        GoTo Son
    End If
    
    If IsNumeric(core_asset_manager_UISayPrt) = False Then
        MsgBox "The print operation could not be initiated due to non-numeric data entry.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    If core_asset_manager_UISayPrt > 3 Then
        MsgBox "The number of copies to be printed at once cannot exceed 3.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If

    
    Call Tutanak_Click
    
    core_asset_manager_UIDirektYazdir = False

Son:

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
SourceTaslak = AutoPath & "\System Files\Help Documents\Asset Manager Panel – Help.docm"
'Birden fazla kullanıcının operasyon klasöründe çakışmasını önlemek için operasyon klasörünün içinde yeni klasör aç.
DestOpUserFolderName = "Operation-" & Left(ThisWorkbook.name, InStr(ThisWorkbook.name, ".") - 1) '& " " & Format(Now(), "ddmmyyyyhhmmss")
DestOpUserFolder = DestOperasyon & DestOpUserFolderName & "\"

'Check the name of the "System Files" folder.
If Not Dir(AutoPath & "\System Files\", vbDirectory) <> vbNullString Then
    MsgBox AutoPath & "\System Files\" & " directory could not be accessed. The folder named 'System Files' may have been renamed or deleted.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

'Check the name of the "Operation" folder.
If Not Dir(DestOperasyon, vbDirectory) <> vbNullString Then
    MsgBox DestOperasyon & " directory could not be accessed. The folder named 'Operation' may have been renamed or the 'System Files' folder may have been deleted.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

'Check folder names.
If Not Dir(SourceTaslak, vbDirectory) <> vbNullString Then
    MsgBox SourceTaslak & " directory could not be accessed. The names of the folders and/or files in this directory may have been changed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
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
Call ModuleReport1.OpenWordControl

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
Private Sub Tutanak_Click()
Dim AutoPath As String, DestOperasyon As String, SourceTeslim As String, DestOpUserFolderName As String, DestOpUserFolder As String
Dim ReNameTeslim As String, OpenKontrolName As String, ContSay As Long, KontrolFile As String
Dim fso As Object, objWord As Object, objDoc As Object, Birimx As String, OpenControl As String
Dim Cont As Long, SiraBul As Range, i As Long, j As Long, Say As Long
'Dim Ctl As MSForms.Control
Dim Ctlx As MSForms.Control, Say2 As Integer, IlkSira As Long
Dim MyRange As Object, AdetSay As Integer, VeriKontrol As Integer
Dim EkFormu As String, SourceTip5 As String, ReNameTip5 As String, EkFormuTip5 As String
Dim a() As Variant, b As Variant
Dim GidenPaketAdet As Integer
Dim MyFile As String, VarlikSecimKontrol As Integer

ThisWorkbook.Activate

Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.EnableEvents = False

ThisWorkbook.Unprotect "123"
Worksheets(3).Unprotect Password:="123"
Worksheets(4).Unprotect Password:="123"
Worksheets(5).Unprotect Password:="123"

'Tarih kontrolü
If VarlikTarihiText.Value = "" Then
    MsgBox "The process cannot be started because the asset date has not been selected.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If


For Each ctl In core_asset_manager_UI.FrameCikis.Controls '________________RAPOR1 BÖLÜMÜ
    If TypeName(ctl) = "ListBox" Then
        'MsgBox Mid(Ctl.List(0), InStr(Ctl.List(0), "|") + 1, 2)
        If Mid(ctl.List(0), InStr(ctl.List(0), "|") + 1, 2) = " Ö" Then
            Call Rapor1SiraNo
            If IlkSiraGlobal = 0 Then
                GoTo Rapor1Devam1
            End If
            IlkSira = IlkSiraGlobal
            ThisWorkbook.Worksheets(3).Cells(IlkSira, 96).Value = VarlikTarihiText.Value
        End If
    End If
Rapor1Devam1:
Next ctl '________________RAPOR1 BÖLÜMÜ

For Each ctl In core_asset_manager_UI.FrameGiris.Controls '________________RAPOR1 BÖLÜMÜ
    If TypeName(ctl) = "ListBox" Then
        'MsgBox Mid(Ctl.List(0), InStr(Ctl.List(0), "|") + 1, 2)
        If Mid(ctl.List(0), InStr(ctl.List(0), "|") + 1, 2) = " Ö" Then
            Call Rapor1SiraNo
            If IlkSiraGlobal = 0 Then
                GoTo Rapor1Devam2
            End If
            IlkSira = IlkSiraGlobal
            If CDate(ThisWorkbook.Worksheets(3).Cells(IlkSira, 96).Value) = CDate(VarlikTarihiText.Value) And _
                ThisWorkbook.Worksheets(3).Cells(IlkSira, 96).Value <> "" Then
                ThisWorkbook.Worksheets(3).Cells(IlkSira, 96).ClearContents
            End If
        End If
    End If
Rapor1Devam2:
Next ctl '________________RAPOR1 BÖLÜMÜ

For Each ctl In core_asset_manager_UI.FrameMevcut.Controls '________________RAPOR1 BÖLÜMÜ
    If TypeName(ctl) = "ListBox" Then
        'MsgBox Mid(Ctl.List(0), InStr(Ctl.List(0), "|") + 1, 2)
        If Mid(ctl.List(0), InStr(ctl.List(0), "|") + 1, 2) = " Ö" Then
            Call Rapor1SiraNo
            If IlkSiraGlobal = 0 Then
                GoTo Rapor1Devam3
            End If
            IlkSira = IlkSiraGlobal
            If CDate(ThisWorkbook.Worksheets(3).Cells(IlkSira, 96).Value) = CDate(VarlikTarihiText.Value) And _
                ThisWorkbook.Worksheets(3).Cells(IlkSira, 96).Value <> "" Then
                ThisWorkbook.Worksheets(3).Cells(IlkSira, 96).ClearContents
            End If
        End If
    End If
Rapor1Devam3:
Next ctl '________________RAPOR1 BÖLÜMÜ



For Each ctl In core_asset_manager_UI.FrameCikis.Controls '________________RAPOR BÖLÜMÜ
    If TypeName(ctl) = "ListBox" Then
        'MsgBox Mid(Ctl.List(0), InStr(Ctl.List(0), "|") + 1, 2)
        If Mid(ctl.List(0), InStr(ctl.List(0), "|") + 1, 2) = " R" Then
            Call RaporSiraNo
            If IlkSiraGlobal = 0 Then
                GoTo RaporDevam1
            End If
            IlkSira = IlkSiraGlobal
            ThisWorkbook.Worksheets(4).Cells(IlkSira, 104).Value = VarlikTarihiText.Value
        End If
    End If
RaporDevam1:
Next ctl '________________RAPOR BÖLÜMÜ

For Each ctl In core_asset_manager_UI.FrameGiris.Controls '________________RAPOR BÖLÜMÜ
    If TypeName(ctl) = "ListBox" Then
        'MsgBox Mid(Ctl.List(0), InStr(Ctl.List(0), "|") + 1, 2)
        If Mid(ctl.List(0), InStr(ctl.List(0), "|") + 1, 2) = " R" Then
            Call RaporSiraNo
            If IlkSiraGlobal = 0 Then
                GoTo RaporDevam2
            End If
            IlkSira = IlkSiraGlobal
            If CDate(ThisWorkbook.Worksheets(4).Cells(IlkSira, 104).Value) = CDate(VarlikTarihiText.Value) And _
                ThisWorkbook.Worksheets(4).Cells(IlkSira, 104).Value <> "" Then
                ThisWorkbook.Worksheets(4).Cells(IlkSira, 104).ClearContents
            End If
        End If
    End If
RaporDevam2:
Next ctl '________________RAPOR BÖLÜMÜ

For Each ctl In core_asset_manager_UI.FrameMevcut.Controls '________________RAPOR BÖLÜMÜ
    If TypeName(ctl) = "ListBox" Then
        'MsgBox Mid(Ctl.List(0), InStr(Ctl.List(0), "|") + 1, 2)
        If Mid(ctl.List(0), InStr(ctl.List(0), "|") + 1, 2) = " R" Then
            Call RaporSiraNo
            If IlkSiraGlobal = 0 Then
                GoTo RaporDevam3
            End If
            IlkSira = IlkSiraGlobal
            If CDate(ThisWorkbook.Worksheets(4).Cells(IlkSira, 104).Value) = CDate(VarlikTarihiText.Value) And _
                ThisWorkbook.Worksheets(4).Cells(IlkSira, 104).Value <> "" Then
                ThisWorkbook.Worksheets(4).Cells(IlkSira, 104).ClearContents
            End If
        End If
    End If
RaporDevam3:
Next ctl '________________RAPOR BÖLÜMÜ



For Each ctl In core_asset_manager_UI.FrameCikis.Controls '________________RAPOR2_2 BÖLÜMÜ
    If TypeName(ctl) = "ListBox" Then
        'MsgBox Mid(Ctl.List(0), InStr(Ctl.List(0), "|") + 1, 2)
        If Mid(ctl.List(0), InStr(ctl.List(0), "|") + 1, 2) = " B" Then
            Call RaporSiraNo
            If IlkSiraGlobal = 0 Then
                GoTo Rapor2_2Devam1
            End If
            IlkSira = IlkSiraGlobal
            'If ThisWorkbook.Worksheets(4).Cells(IlkSira, 191).Value <> "" Then 'XXXMuddan geliş
            If InStr(ctl.List(0), "(Gelen)") <> 0 Then
                If InStr(ctl.List(0), "<T>") <> 0 Then
                    ThisWorkbook.Worksheets(4).Cells(IlkSira, 106).Value = VarlikTarihiText.Value
                End If
                If InStr(ctl.List(0), "<L>") <> 0 Then
                    ThisWorkbook.Worksheets(4).Cells(IlkSira, 106).Value = VarlikTarihiText.Value
                End If
            Else 'XXXMudya gidiş
                If InStr(ctl.List(0), "<T>") <> 0 Then
                    ThisWorkbook.Worksheets(4).Cells(IlkSira, 105).Value = VarlikTarihiText.Value
                End If
                If InStr(ctl.List(0), "<L>") <> 0 Then
                    ThisWorkbook.Worksheets(4).Cells(IlkSira, 105).Value = VarlikTarihiText.Value
                End If
            End If
            If InStr(ctl.List(0), "<R>") <> 0 Then
                ThisWorkbook.Worksheets(4).Cells(IlkSira, 104).Value = VarlikTarihiText.Value
            End If

        End If
    End If
Rapor2_2Devam1:
Next ctl '________________RAPOR2_2 BÖLÜMÜ

For Each ctl In core_asset_manager_UI.FrameGiris.Controls '________________RAPOR2_2 BÖLÜMÜ
    If TypeName(ctl) = "ListBox" Then
        'MsgBox Mid(Ctl.List(0), InStr(Ctl.List(0), "|") + 1, 2)
        If Mid(ctl.List(0), InStr(ctl.List(0), "|") + 1, 2) = " B" Then
            Call RaporSiraNo
            If IlkSiraGlobal = 0 Then
                GoTo Rapor2_2Devam2
            End If
            IlkSira = IlkSiraGlobal
            'If ThisWorkbook.Worksheets(4).Cells(IlkSira, 191).Value <> "" Then
            If InStr(ctl.List(0), "(Gelen)") <> 0 Then
                If InStr(ctl.List(0), "<T>") <> 0 Then
                    If CDate(ThisWorkbook.Worksheets(4).Cells(IlkSira, 106).Value) = CDate(VarlikTarihiText.Value) And _
                        ThisWorkbook.Worksheets(4).Cells(IlkSira, 106).Value <> "" Then
                        ThisWorkbook.Worksheets(4).Cells(IlkSira, 106).ClearContents
                    End If
                End If
                If InStr(ctl.List(0), "<L>") <> 0 Then
                    If CDate(ThisWorkbook.Worksheets(4).Cells(IlkSira, 106).Value) = CDate(VarlikTarihiText.Value) And _
                        ThisWorkbook.Worksheets(4).Cells(IlkSira, 106).Value <> "" Then
                        ThisWorkbook.Worksheets(4).Cells(IlkSira, 106).ClearContents
                    End If
                End If
            Else
                If InStr(ctl.List(0), "<T>") <> 0 Then
                    If CDate(ThisWorkbook.Worksheets(4).Cells(IlkSira, 105).Value) = CDate(VarlikTarihiText.Value) And _
                        ThisWorkbook.Worksheets(4).Cells(IlkSira, 105).Value <> "" Then
                        ThisWorkbook.Worksheets(4).Cells(IlkSira, 105).ClearContents
                    End If
                End If
                If InStr(ctl.List(0), "<L>") <> 0 Then
                    If CDate(ThisWorkbook.Worksheets(4).Cells(IlkSira, 105).Value) = CDate(VarlikTarihiText.Value) And _
                        ThisWorkbook.Worksheets(4).Cells(IlkSira, 105).Value <> "" Then
                        ThisWorkbook.Worksheets(4).Cells(IlkSira, 105).ClearContents
                    End If
                End If
            End If
            If InStr(ctl.List(0), "<R>") <> 0 Then
                If CDate(ThisWorkbook.Worksheets(4).Cells(IlkSira, 104).Value) = CDate(VarlikTarihiText.Value) And _
                    ThisWorkbook.Worksheets(4).Cells(IlkSira, 104).Value <> "" Then
                    ThisWorkbook.Worksheets(4).Cells(IlkSira, 104).ClearContents
                End If
            End If
            
        End If
    End If
Rapor2_2Devam2:
Next ctl '________________RAPOR2_2 BÖLÜMÜ

For Each ctl In core_asset_manager_UI.FrameMevcut.Controls '________________RAPOR2_2 BÖLÜMÜ
    If TypeName(ctl) = "ListBox" Then
        'MsgBox Mid(Ctl.List(0), InStr(Ctl.List(0), "|") + 1, 2)
        If Mid(ctl.List(0), InStr(ctl.List(0), "|") + 1, 2) = " B" Then
            Call RaporSiraNo
            If IlkSiraGlobal = 0 Then
                GoTo Rapor2_2Devam3
            End If
            IlkSira = IlkSiraGlobal
            'If ThisWorkbook.Worksheets(4).Cells(IlkSira, 191).Value <> "" Then
            If InStr(ctl.List(0), "(Gelen)") <> 0 Then
                If InStr(ctl.List(0), "<T>") <> 0 Then
                    If CDate(ThisWorkbook.Worksheets(4).Cells(IlkSira, 106).Value) = CDate(VarlikTarihiText.Value) And _
                        ThisWorkbook.Worksheets(4).Cells(IlkSira, 106).Value <> "" Then
                        ThisWorkbook.Worksheets(4).Cells(IlkSira, 106).ClearContents
                    End If
                End If
                If InStr(ctl.List(0), "<L>") <> 0 Then
                    If CDate(ThisWorkbook.Worksheets(4).Cells(IlkSira, 106).Value) = CDate(VarlikTarihiText.Value) And _
                        ThisWorkbook.Worksheets(4).Cells(IlkSira, 106).Value <> "" Then
                        ThisWorkbook.Worksheets(4).Cells(IlkSira, 106).ClearContents
                    End If
                End If
            Else
                If InStr(ctl.List(0), "<T>") <> 0 Then
                    If CDate(ThisWorkbook.Worksheets(4).Cells(IlkSira, 105).Value) = CDate(VarlikTarihiText.Value) And _
                        ThisWorkbook.Worksheets(4).Cells(IlkSira, 105).Value <> "" Then
                        ThisWorkbook.Worksheets(4).Cells(IlkSira, 105).ClearContents
                    End If
                End If
                If InStr(ctl.List(0), "<L>") <> 0 Then
                    If CDate(ThisWorkbook.Worksheets(4).Cells(IlkSira, 105).Value) = CDate(VarlikTarihiText.Value) And _
                        ThisWorkbook.Worksheets(4).Cells(IlkSira, 105).Value <> "" Then
                        ThisWorkbook.Worksheets(4).Cells(IlkSira, 105).ClearContents
                    End If
                End If
            End If
            If InStr(ctl.List(0), "<R>") <> 0 Then
                If CDate(ThisWorkbook.Worksheets(4).Cells(IlkSira, 104).Value) = CDate(VarlikTarihiText.Value) And _
                    ThisWorkbook.Worksheets(4).Cells(IlkSira, 104).Value <> "" Then
                    ThisWorkbook.Worksheets(4).Cells(IlkSira, 104).ClearContents
                End If
            End If
            
        End If
    End If
Rapor2_2Devam3:
Next ctl '________________RAPOR2_2 BÖLÜMÜ



For Each ctl In core_asset_manager_UI.FrameCikis.Controls '________________RAPOR3_2 BÖLÜMÜ
    If TypeName(ctl) = "ListBox" Then
        'MsgBox Mid(Ctl.List(0), InStr(Ctl.List(0), "|") + 1, 2)
        If Mid(ctl.List(0), InStr(ctl.List(0), "|") + 1, 2) = " M" Then
            Call Rapor3Rapor3_2SiraNo
            If IlkSiraGlobal = 0 Then
                GoTo Rapor3Rapor3_2Devam1
            End If
            IlkSira = IlkSiraGlobal
            ThisWorkbook.Worksheets(5).Cells(IlkSira, 176).Value = VarlikTarihiText.Value
        End If
    End If
Rapor3Rapor3_2Devam1:
Next ctl '________________RAPOR3_2 BÖLÜMÜ

For Each ctl In core_asset_manager_UI.FrameGiris.Controls '________________RAPOR3_2 BÖLÜMÜ
    If TypeName(ctl) = "ListBox" Then
        'MsgBox Mid(Ctl.List(0), InStr(Ctl.List(0), "|") + 1, 2)
        If Mid(ctl.List(0), InStr(ctl.List(0), "|") + 1, 2) = " M" Then
            Call Rapor3Rapor3_2SiraNo
            If IlkSiraGlobal = 0 Then
                GoTo Rapor3Rapor3_2Devam2
            End If
            IlkSira = IlkSiraGlobal
            If CDate(ThisWorkbook.Worksheets(5).Cells(IlkSira, 176).Value) = CDate(VarlikTarihiText.Value) And _
                ThisWorkbook.Worksheets(5).Cells(IlkSira, 176).Value <> "" Then
                ThisWorkbook.Worksheets(5).Cells(IlkSira, 176).ClearContents
            End If
        End If
    End If
Rapor3Rapor3_2Devam2:
Next ctl '________________RAPOR3_2 BÖLÜMÜ

For Each ctl In core_asset_manager_UI.FrameMevcut.Controls '________________RAPOR3_2 BÖLÜMÜ
    If TypeName(ctl) = "ListBox" Then
        'MsgBox Mid(Ctl.List(0), InStr(Ctl.List(0), "|") + 1, 2)
        If Mid(ctl.List(0), InStr(ctl.List(0), "|") + 1, 2) = " M" Then
            Call Rapor3Rapor3_2SiraNo
            If IlkSiraGlobal = 0 Then
                GoTo Rapor3Rapor3_2Devam3
            End If
            IlkSira = IlkSiraGlobal
            If CDate(ThisWorkbook.Worksheets(5).Cells(IlkSira, 176).Value) = CDate(VarlikTarihiText.Value) And _
                ThisWorkbook.Worksheets(5).Cells(IlkSira, 176).Value <> "" Then
                ThisWorkbook.Worksheets(5).Cells(IlkSira, 176).ClearContents
            End If
        End If
    End If
Rapor3Rapor3_2Devam3:
Next ctl '________________RAPOR3_2 BÖLÜMÜ



For Each ctl In core_asset_manager_UI.FrameCikis.Controls '________________RAPOR3_1 BÖLÜMÜ
    If TypeName(ctl) = "ListBox" Then
        'MsgBox Mid(Ctl.List(0), InStr(Ctl.List(0), "|") + 1, 2)
        If Mid(ctl.List(0), InStr(ctl.List(0), "|") + 1, 2) = " G" Then
            Call Rapor3Rapor3_1SiraNo
            If IlkSiraGlobal = 0 Then
                GoTo Rapor3Rapor3_1Devam1
            End If
            IlkSira = IlkSiraGlobal
            ThisWorkbook.Worksheets(5).Cells(IlkSira, 176).Value = VarlikTarihiText.Value
        End If
    End If
Rapor3Rapor3_1Devam1:
Next ctl '________________RAPOR3_1 BÖLÜMÜ

For Each ctl In core_asset_manager_UI.FrameGiris.Controls '________________RAPOR3_1 BÖLÜMÜ
    If TypeName(ctl) = "ListBox" Then
        'MsgBox Mid(Ctl.List(0), InStr(Ctl.List(0), "|") + 1, 2)
        If Mid(ctl.List(0), InStr(ctl.List(0), "|") + 1, 2) = " G" Then
            Call Rapor3Rapor3_1SiraNo
            If IlkSiraGlobal = 0 Then
                GoTo Rapor3Rapor3_1Devam2
            End If
            IlkSira = IlkSiraGlobal
            If CDate(ThisWorkbook.Worksheets(5).Cells(IlkSira, 176).Value) = CDate(VarlikTarihiText.Value) And _
                ThisWorkbook.Worksheets(5).Cells(IlkSira, 176).Value <> "" Then
                ThisWorkbook.Worksheets(5).Cells(IlkSira, 176).ClearContents
            End If
        End If
    End If
Rapor3Rapor3_1Devam2:
Next ctl '________________RAPOR3_1 BÖLÜMÜ

For Each ctl In core_asset_manager_UI.FrameMevcut.Controls '________________RAPOR3_1 BÖLÜMÜ
    If TypeName(ctl) = "ListBox" Then
        'MsgBox Mid(Ctl.List(0), InStr(Ctl.List(0), "|") + 1, 2)
        If Mid(ctl.List(0), InStr(ctl.List(0), "|") + 1, 2) = " G" Then
            Call Rapor3Rapor3_1SiraNo
            If IlkSiraGlobal = 0 Then
                GoTo Rapor3Rapor3_1Devam3
            End If
            IlkSira = IlkSiraGlobal
            If CDate(ThisWorkbook.Worksheets(5).Cells(IlkSira, 176).Value) = CDate(VarlikTarihiText.Value) And _
                ThisWorkbook.Worksheets(5).Cells(IlkSira, 176).Value <> "" Then
                ThisWorkbook.Worksheets(5).Cells(IlkSira, 176).ClearContents
            End If
        End If
    End If
Rapor3Rapor3_1Devam3:
Next ctl '________________RAPOR3_1 BÖLÜMÜ


''Açık tüm excel varlık dosyalarını kapat.
'Pathfinder...
AutoPath = ThisWorkbook.Path
DestOperasyon = AutoPath & "\System Files\Operation\"
'Birden fazla kullanıcının operasyon klasöründe çakışmasını önlemek için operasyon klasörünün içinde yeni klasör aç.
DestOpUserFolderName = "Operation-" & Left(ThisWorkbook.name, InStr(ThisWorkbook.name, ".") - 1) '& " " & Format(Now(), "ddmmyyyyhhmmss")
DestOpUserFolder = DestOperasyon & DestOpUserFolderName & "\"

On Error Resume Next
    MyFile = Dir(DestOpUserFolder & "*.xl??")
    Do While MyFile <> ""
        DoEvents
        Workbooks(MyFile).Close SaveChanges:=False
        MyFile = Dir
    Loop
On Error GoTo 0


'GoTo Son

'Varlik tutanakları
VarlikSecimKontrol = 0
'Esas Varliklar
If CheckBoxRapor1VarlikEsas.Value = True Then
    Call ModuleAsset.Rapor1VarlikEsas
    VarlikSecimKontrol = 1
End If
If CheckBoxRaporVarlikEsas.Value = True Then
    Call ModuleAsset.RaporVarlikEsas
    VarlikSecimKontrol = 1
End If
If CheckBoxRapor2_2VarlikEsas.Value = True Then
    Call ModuleAsset.Rapor2_2VarlikEsas
    VarlikSecimKontrol = 1
End If
If CheckBoxRapor3VarlikEsas.Value = True Then
    Call ModuleAsset.Rapor3VarlikEsas
    VarlikSecimKontrol = 1
End If
If CheckBoxGenelVarlikEsas.Value = True Then
    Call ModuleAsset.GenelVarlikEsas
    VarlikSecimKontrol = 1
End If
'Yardımcı Varliklar
If CheckBoxRapor1Varlik.Value = True Then
    Call ModuleAsset.Rapor1Varlik
    VarlikSecimKontrol = 1
End If
If CheckBoxRaporVarlik.Value = True Then
    Call ModuleAsset.RaporVarlik
    VarlikSecimKontrol = 1
End If
If CheckBoxRapor2_2Varlik.Value = True Then
    Call ModuleAsset.Rapor2_2Varlik
    VarlikSecimKontrol = 1
End If
If CheckBoxRapor3Varlik.Value = True Then
    Call ModuleAsset.Rapor3Varlik
    VarlikSecimKontrol = 1
End If
If CheckBoxGenelVarlik.Value = True Then
    Call ModuleAsset.GenelVarlik
    VarlikSecimKontrol = 1
End If
If VarlikSecimKontrol = 0 Then
    MsgBox "It has been detected that the report type has not been specified. Please select at least one report type from either the Main Asset Reports or the Supporting Asset Reports section.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

Call ModuleSystemSettings.DropDownKapat

Son:

Worksheets(3).Protect Password:="123"
Worksheets(4).Protect Password:="123"
Worksheets(5).Protect Password:="123"

ThisWorkbook.Protect "123"

Application.ScreenUpdating = True
Application.DisplayAlerts = True
Application.EnableEvents = True

End Sub

Sub Rapor1SiraNo()
Dim SiraBul As Range

IlkSiraGlobal = 0
Set SiraBul = ThisWorkbook.Worksheets(3).Range("E7:E100000").Find(What:=Left(ctl.List(0), InStr(ctl.List(0), "|") - 2), SearchDirection:=xlNext, _
                SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
If Not SiraBul Is Nothing Then
    IlkSiraGlobal = SiraBul.Row
Else
    GoTo Rapor1Devam1
End If

Rapor1Devam1:

End Sub

Sub RaporSiraNo()
Dim SiraBul As Range

IlkSiraGlobal = 0
Set SiraBul = ThisWorkbook.Worksheets(4).Range("E7:E100000").Find(What:=Left(ctl.List(0), InStr(ctl.List(0), "|") - 2), SearchDirection:=xlNext, _
                SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
If Not SiraBul Is Nothing Then
    IlkSiraGlobal = SiraBul.Row
Else
    GoTo RaporDevam1
End If

RaporDevam1:

End Sub

Sub Rapor3Rapor3_2SiraNo()
Dim SiraBul As Range

IlkSiraGlobal = 0
Set SiraBul = ThisWorkbook.Worksheets(5).Range("E7:E100000").Find(What:=Left(ctl.List(0), InStr(ctl.List(0), "|") - 2), SearchDirection:=xlNext, _
                SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
If Not SiraBul Is Nothing Then
    IlkSiraGlobal = SiraBul.Row
Else
    GoTo Rapor3Rapor3_2Devam1
End If
        
Rapor3Rapor3_2Devam1:

End Sub

Sub Rapor3Rapor3_1SiraNo()
Dim SiraBul As Range

IlkSiraGlobal = 0
Set SiraBul = ThisWorkbook.Worksheets(5).Range("E7:E100000").Find(What:=Left(ctl.List(0), InStr(ctl.List(0), "|") - 2), SearchDirection:=xlNext, _
                SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
If Not SiraBul Is Nothing Then
    IlkSiraGlobal = SiraBul.Row
Else
    GoTo Rapor3Rapor3_1Devam1
End If

Rapor3Rapor3_1Devam1:

End Sub

Private Sub UserForm_Initialize()
Dim i As Long
Dim Say As Long, j As Long, Cont As Long, Sno As Variant
Dim LstBx As MSForms.ListBox
Dim LblSira1 As MSForms.Label
Dim LblSira2 As MSForms.Label
Dim ClrLab As MSForms.Control

ThisWorkbook.Activate

ScrollTakip1 = 0
ScrollTakip2 = 0
ScrollTakip3 = 0

For Each ClrLab In core_asset_manager_UI.Controls

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

DirektYazdir.BackColor = RGB(225, 235, 245)
DirektYazdir.ForeColor = RGB(30, 30, 30)
Tutanak.BackColor = RGB(225, 235, 245)
Tutanak.ForeColor = RGB(30, 30, 30)
Yardim.BackColor = RGB(225, 235, 245)
Yardim.ForeColor = RGB(30, 30, 30)
Kapat.BackColor = RGB(225, 235, 245)
Kapat.ForeColor = RGB(30, 30, 30)

core_asset_manager_UI.BackColor = RGB(230, 230, 230) 'YENİ

LabelSiraGiris.BorderColor = RGB(180, 180, 180)
LabelKayitNoGiris.BorderColor = RGB(180, 180, 180)
LabelGonderenGiris.BorderColor = RGB(180, 180, 180)
LabelSiraCikis.BorderColor = RGB(180, 180, 180)
LabelKayitNoCikis.BorderColor = RGB(180, 180, 180)
LabelGonderenCikis.BorderColor = RGB(180, 180, 180)
LabelSiraMevcut.BorderColor = RGB(180, 180, 180)
LabelKayitNoMevcut.BorderColor = RGB(180, 180, 180)
LabelGonderenMevcut.BorderColor = RGB(180, 180, 180)


TasiyiciFrame.Height = 514
FrameGirisx.Height = 264
FrameCikisx.Height = 264
FrameMevcutx.Height = 264

FrameGirisx.ZOrder msoBringToFront
FrameCikisx.ZOrder msoBringToFront
FrameMevcutx.ZOrder msoBringToFront
FrameGiris.ZOrder msoBringToFront
FrameCikis.ZOrder msoBringToFront
FrameMevcut.ZOrder msoBringToFront

FrameAlt1.ZOrder msoBringToFront
FrameAlt2.ZOrder msoBringToFront
FrameAlt3.ZOrder msoBringToFront

core_asset_manager_UIDirektYazdir = False

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
Dim yukseklik As Variant, genislik As Variant
Dim Rep As Variant

RemoveScrollHook

yukseklik = Me.Height
Rep = Me.Width
Do
DoEvents
Rep = Rep - 90
Call timeout(0.01)
    If Rep > 90 Then
        core_asset_manager_UI.Width = Rep
        yukseklik = yukseklik - 90
        core_asset_manager_UI.Height = yukseklik
        If yukseklik <= 90 Then
            yukseklik = 90
            core_asset_manager_UI.Height = yukseklik
        End If
    ElseIf Rep <= 50 Then
        Rep = 50
        core_asset_manager_UI.Width = Rep
        yukseklik = yukseklik - 50
        core_asset_manager_UI.Height = yukseklik
        If yukseklik <= 50 Then
            yukseklik = 50
            core_asset_manager_UI.Height = yukseklik
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




