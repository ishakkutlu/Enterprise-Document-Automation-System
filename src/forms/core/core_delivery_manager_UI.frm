VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} core_delivery_manager_UI 
   Caption         =   "Enterprise Document Automation System"
   ClientHeight    =   11130
   ClientLeft      =   30
   ClientTop       =   375
   ClientWidth     =   16320
   OleObjectBlob   =   "core_delivery_manager_UI.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "core_delivery_manager_UI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim TeslimTutanaklariFormuDirektYazdir As Boolean
Dim TeslimTutanaklariFormuSayPrt As Variant
Dim GidenTemaGlobal As String, IlkSiraGlobal As Long
Dim ctl As MSForms.Control
Dim Abort As Boolean

Private Sub Tip1Imza1EkleKaldirLabel_Click()
support_signatures_UI.Show
End Sub
Private Sub Tip1Imza2EkleKaldirLabel_Click()
support_signatures_UI.Show
End Sub
Private Sub Tip1Imza3EkleKaldirLabel_Click()
support_signatures_UI.Show
End Sub

Private Sub Tip1Imza1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next

   Select Case KeyCode
        Case 38  'Yukarı
            If Tip1Imza1.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                Tip1Imza1.ListIndex = Tip1Imza1.ListIndex - 1
            End If
            Me.Tip1Imza1.DropDown
            
        Case 40 'Aşağı
            If Tip1Imza1.ListIndex = Tip1Imza1.ListCount - 1 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                Tip1Imza1.ListIndex = Tip1Imza1.ListIndex + 1
            End If
            Me.Tip1Imza1.DropDown
    End Select
    Abort = False
    
End Sub

Private Sub Tip1Imza1_Change()

If Tip1Imza1.ListIndex = -1 And Tip1Imza1.Value <> "" Then
   Tip1Imza1.Value = ""
   GoTo Son
End If

If Tip1Imza1.Value <> "" Then
    Tip1Imza1.SelStart = 0
    Tip1Imza1.SelLength = Len(Tip1Imza1.Value)
End If


Son:

Tip1Imza1.DropDown

End Sub

Private Sub Tip1Imza2_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next

   Select Case KeyCode
        Case 38  'Yukarı
            If Tip1Imza2.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                Tip1Imza2.ListIndex = Tip1Imza2.ListIndex - 1
            End If
            Me.Tip1Imza2.DropDown
            
        Case 40 'Aşağı
            If Tip1Imza2.ListIndex = Tip1Imza2.ListCount - 1 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                Tip1Imza2.ListIndex = Tip1Imza2.ListIndex + 1
            End If
            Me.Tip1Imza2.DropDown
    End Select
    Abort = False
    
End Sub

Private Sub Tip1Imza2_Change()

If Tip1Imza2.ListIndex = -1 And Tip1Imza2.Value <> "" Then
   Tip1Imza2.Value = ""
   GoTo Son
End If

If Tip1Imza2.Value <> "" Then
    Tip1Imza2.SelStart = 0
    Tip1Imza2.SelLength = Len(Tip1Imza2.Value)
End If


Son:

Tip1Imza2.DropDown

End Sub

Private Sub Tip1Imza3_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next

   Select Case KeyCode
        Case 38  'Yukarı
            If Tip1Imza3.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                Tip1Imza3.ListIndex = Tip1Imza3.ListIndex - 1
            End If
            Me.Tip1Imza3.DropDown
            
        Case 40 'Aşağı
            If Tip1Imza3.ListIndex = Tip1Imza3.ListCount - 1 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                Tip1Imza3.ListIndex = Tip1Imza3.ListIndex + 1
            End If
            Me.Tip1Imza3.DropDown
    End Select
    Abort = False
    
End Sub

Private Sub Tip1Imza3_Change()

If Tip1Imza3.ListIndex = -1 And Tip1Imza3.Value <> "" Then
   Tip1Imza3.Value = ""
   GoTo Son
End If

If Tip1Imza3.Value <> "" Then
    Tip1Imza3.SelStart = 0
    Tip1Imza3.SelLength = Len(Tip1Imza3.Value)
End If


Son:

Tip1Imza3.DropDown

End Sub

Private Sub Tip1Option_AfterUpdate()
Me.Frame3.SetFocus
End Sub

Private Sub Tip1Option_Click()
    TutanakNoText.Value = ""
    TutanakNoText.Enabled = False
    TutanakNoFrame.Visible = False
    Tip1Imza1Frame.Visible = True
    Tip1Imza2Frame.Visible = True
    Tip1Imza3Frame.Visible = True
End Sub

Private Sub Tip2Option_AfterUpdate()
Me.Frame3.SetFocus
End Sub

Private Sub Tip2Option_Click()
    TutanakNoText.Enabled = True
    TutanakNoFrame.Visible = True
    Tip1Imza1Frame.Visible = False
    Tip1Imza2Frame.Visible = False
    Tip1Imza3Frame.Visible = False
    Call SonTutanakNoTeslim
End Sub
Private Sub Tip3Option_Click()
    TutanakNoText.Enabled = True
    TutanakNoFrame.Visible = True
    Tip1Imza1Frame.Visible = False
    Tip1Imza2Frame.Visible = False
    Tip1Imza3Frame.Visible = False
    Call SonTutanakNoTeslim
End Sub
Private Sub Tip4Option_Click()
    TutanakNoText.Enabled = True
    TutanakNoFrame.Visible = True
    Tip1Imza1Frame.Visible = False
    Tip1Imza2Frame.Visible = False
    Tip1Imza3Frame.Visible = False
    Call SonTutanakNoTeslim
End Sub

Private Sub Tip3Option_AfterUpdate()
Me.Frame3.SetFocus
End Sub
Private Sub Tip4Option_AfterUpdate()
Me.Frame3.SetFocus
End Sub

'Private Sub CheckBox1_AfterUpdate()
'Me.Frame1.SetFocus
'End Sub
'
'Private Sub CheckBox2_AfterUpdate()
'Me.Frame2.SetFocus
'End Sub


Private Sub Tip1Option_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
Tip1Option.BackColor = RGB(60, 100, 180) 'RGB(60, 100, 180) 'RGB(60, 100, 180)
Tip1Option.ForeColor = RGB(255, 255, 255)
End Sub
Private Sub Tip1Imza1EkleKaldirLabel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
Tip1Imza1EkleKaldirLabel.BackColor = RGB(60, 100, 180)
Tip1Imza1EkleKaldirLabel.ForeColor = RGB(255, 255, 255)
End Sub
Private Sub Tip1Imza1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
Call SetComboBoxHook(Tip1Imza1) 'Open scrollable with mouse
End Sub
Private Sub LblTip1Imza1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub
Private Sub Tip1Imza2EkleKaldirLabel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
Tip1Imza2EkleKaldirLabel.BackColor = RGB(60, 100, 180)
Tip1Imza2EkleKaldirLabel.ForeColor = RGB(255, 255, 255)
End Sub
Private Sub Tip1Imza2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
Call SetComboBoxHook(Tip1Imza2) 'Open scrollable with mouse
End Sub
Private Sub LblTip1Imza2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub
Private Sub Tip1Imza3EkleKaldirLabel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
Tip1Imza3EkleKaldirLabel.BackColor = RGB(60, 100, 180)
Tip1Imza3EkleKaldirLabel.ForeColor = RGB(255, 255, 255)
End Sub
Private Sub Tip1Imza3_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
Call SetComboBoxHook(Tip1Imza3) 'Open scrollable with mouse
End Sub
Private Sub LblTip1Imza3_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub

Private Sub Tip2Option_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
Tip2Option.BackColor = RGB(60, 100, 180) 'RGB(60, 100, 180) 'RGB(60, 100, 180)
Tip2Option.ForeColor = RGB(255, 255, 255)
End Sub
Private Sub Tip3Option_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
Tip3Option.BackColor = RGB(60, 100, 180) 'RGB(60, 100, 180) 'RGB(60, 100, 180)
Tip3Option.ForeColor = RGB(255, 255, 255)
End Sub
Private Sub Tip4Option_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
Tip4Option.BackColor = RGB(60, 100, 180) 'RGB(60, 100, 180) 'RGB(60, 100, 180)
Tip4Option.ForeColor = RGB(255, 255, 255)
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

Private Sub TutanakNoText_Change()
Dim a() As Variant, i As Variant
Dim j As Integer

''Comboda tanımlı değer girilemez.
'a() = TutanakNoText.List
'For i = LBound(a) To UBound(a)
'    If a(i, 0) = TutanakNoText.Value Then
'        TutanakNoText.Value = ""
'    End If
'Next i

End Sub


Sub ColorChangerGenel()

If TakvimBtn.BackColor <> RGB(225, 235, 245) Then
TakvimBtn.BackColor = RGB(225, 235, 245) 'RGB(254, 254, 254)
TakvimBtn.ForeColor = RGB(30, 30, 30)
End If
If Tutanak.BackColor <> RGB(225, 235, 245) Then
Tutanak.BackColor = RGB(225, 235, 245) 'RGB(254, 254, 254)
Tutanak.ForeColor = RGB(30, 30, 30)
End If
If DirektYazdir.BackColor <> RGB(225, 235, 245) Then
DirektYazdir.BackColor = RGB(225, 235, 245) 'RGB(254, 254, 254)
DirektYazdir.ForeColor = RGB(30, 30, 30)
End If
If Kapat.BackColor <> RGB(225, 235, 245) Then
Kapat.BackColor = RGB(225, 235, 245) 'RGB(254, 254, 254)
Kapat.ForeColor = RGB(30, 30, 30)
End If
If Yardim.BackColor <> RGB(225, 235, 245) Then
Yardim.BackColor = RGB(225, 235, 245) 'RGB(254, 254, 254)
Yardim.ForeColor = RGB(30, 30, 30)
End If


If Tip1Imza1EkleKaldirLabel.BackColor <> RGB(254, 254, 254) Then
Tip1Imza1EkleKaldirLabel.BackColor = RGB(254, 254, 254)
Tip1Imza1EkleKaldirLabel.ForeColor = RGB(70, 70, 70)
End If
If Tip1Imza2EkleKaldirLabel.BackColor <> RGB(254, 254, 254) Then
Tip1Imza2EkleKaldirLabel.BackColor = RGB(254, 254, 254)
Tip1Imza2EkleKaldirLabel.ForeColor = RGB(70, 70, 70)
End If
If Tip1Imza3EkleKaldirLabel.BackColor <> RGB(254, 254, 254) Then
Tip1Imza3EkleKaldirLabel.BackColor = RGB(254, 254, 254)
Tip1Imza3EkleKaldirLabel.ForeColor = RGB(70, 70, 70)
End If

If Ekle.BackColor <> RGB(254, 254, 254) Then
Ekle.BackColor = RGB(254, 254, 254)
Ekle.ForeColor = RGB(70, 70, 70)
End If
If Kaldir.BackColor <> RGB(254, 254, 254) Then
Kaldir.BackColor = RGB(254, 254, 254)
Kaldir.ForeColor = RGB(70, 70, 70)
End If

If CheckBox1.BackColor <> RGB(254, 254, 254) Then
CheckBox1.BackColor = RGB(254, 254, 254) 'RGB(254, 254, 254)
CheckBox1.ForeColor = RGB(70, 70, 70)
End If
If CheckBox2.BackColor <> RGB(254, 254, 254) Then
CheckBox2.BackColor = RGB(254, 254, 254) 'RGB(254, 254, 254)
CheckBox2.ForeColor = RGB(70, 70, 70)
End If

If TutanakTarihiLabel.BackColor <> RGB(254, 254, 254) Then
TutanakTarihiLabel.BackColor = RGB(254, 254, 254) 'RGB(254, 254, 254)
TutanakTarihiLabel.ForeColor = RGB(70, 70, 70)
End If

If Tip1Option.BackColor <> RGB(254, 254, 254) Then
Tip1Option.BackColor = RGB(254, 254, 254) 'RGB(254, 254, 254)
Tip1Option.ForeColor = RGB(70, 70, 70)
End If
If Tip2Option.BackColor <> RGB(254, 254, 254) Then
Tip2Option.BackColor = RGB(254, 254, 254) 'RGB(254, 254, 254)
Tip2Option.ForeColor = RGB(70, 70, 70)
End If
If Tip3Option.BackColor <> RGB(254, 254, 254) Then
Tip3Option.BackColor = RGB(254, 254, 254) 'RGB(254, 254, 254)
Tip3Option.ForeColor = RGB(70, 70, 70)
End If
If Tip4Option.BackColor <> RGB(254, 254, 254) Then
Tip4Option.BackColor = RGB(254, 254, 254) 'RGB(254, 254, 254)
Tip4Option.ForeColor = RGB(70, 70, 70)
End If

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
Private Sub BilgilendirmeFrame_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub
Private Sub LblBilgilendirme_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub
Private Sub AltMenuFrame_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub

Private Sub Frame1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)

Call ColorChangerGenel
Call RemoveScrollHook
If ScrollTakip1 > 180 Then
    Call SetScrollHook(Me.Frame1, ScrollTakip1)
Else
    Frame1.ScrollTop = 0
    RemoveScrollHook
    Frame1.ScrollBars = fmScrollBarsNone
End If

End Sub
Private Sub Frame2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)

Call ColorChangerGenel
Call RemoveScrollHook
If ScrollTakip2 > 180 Then
    Call SetScrollHook(Me.Frame2, ScrollTakip2)
Else
    Frame2.ScrollTop = 0
    RemoveScrollHook
    Frame2.ScrollBars = fmScrollBarsNone
End If

End Sub

Private Sub Frame1x_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
Call RemoveScrollHook
End Sub
Private Sub Frame2x_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
Call RemoveScrollHook
End Sub
Private Sub Frame3_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub
Private Sub Frame5_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub
Private Sub FrameAlt1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
Call RemoveScrollHook
End Sub
Private Sub FrameAlt2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
Call RemoveScrollHook
End Sub


Private Sub DirektYazdir_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
DirektYazdir.BackColor = RGB(60, 100, 180) 'RGB(60, 100, 180) 'RGB(60, 100, 180)
DirektYazdir.ForeColor = RGB(255, 255, 255)
End Sub

Private Sub TakvimBtn_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
TakvimBtn.BackColor = RGB(60, 100, 180) 'RGB(60, 100, 180) 'RGB(60, 100, 180)
TakvimBtn.ForeColor = RGB(255, 255, 255)
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
Private Sub Yardim_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
Yardim.BackColor = RGB(60, 100, 180) 'RGB(60, 100, 180) 'RGB(60, 100, 180)
Yardim.ForeColor = RGB(255, 255, 255)
End Sub
Private Sub TutanakNoText_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub


Private Sub TutanakTarihiLabel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
TutanakTarihiLabel.BackColor = RGB(60, 100, 180) 'RGB(60, 100, 180) 'RGB(60, 100, 180)
TutanakTarihiLabel.ForeColor = RGB(255, 255, 255)
End Sub
Private Sub TutanakTarihiText_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub
Private Sub LblTutanakTarihi_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub
Private Sub Ekle_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
Ekle.BackColor = RGB(60, 100, 180) 'RGB(60, 100, 180) 'RGB(60, 100, 180)
Ekle.ForeColor = RGB(255, 255, 255)
End Sub
Private Sub Kaldir_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
Kaldir.BackColor = RGB(60, 100, 180) 'RGB(60, 100, 180) 'RGB(60, 100, 180)
Kaldir.ForeColor = RGB(255, 255, 255)
End Sub


Private Sub LblSiraNo1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub
Private Sub LblBelgeNo1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub
Private Sub LblGonderen1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub

Private Sub LblSiraNo2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub
Private Sub LblBelgeNo2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub
Private Sub LblGonderen2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub


Private Sub CheckBox1_Click()
Dim i As Long
Dim ctl As MSForms.Control
Dim Say1 As Integer, Say2 As Integer, Cont1 As Integer, Cont2 As Integer
Dim LstBx As MSForms.ListBox

'Tümünü seç
If CheckBox1.Value = True Then
    For Each ctl In core_delivery_manager_UI.Frame1.Controls
        If TypeName(ctl) = "ListBox" Then
            ctl.Selected(0) = True
        End If
    Next ctl
Else 'Tümünü iptal et
    For Each ctl In core_delivery_manager_UI.Frame1.Controls
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
    For Each ctl In core_delivery_manager_UI.Frame2.Controls
        If TypeName(ctl) = "ListBox" Then
            ctl.Selected(0) = True
        End If
    Next ctl
Else 'Tümünü iptal et
    For Each ctl In core_delivery_manager_UI.Frame2.Controls
        If TypeName(ctl) = "ListBox" Then
            ctl.Selected(0) = False
        End If
    Next ctl
End If

End Sub

Private Sub Ekle_Click()
Dim i As Long
Dim ctl As MSForms.Control
Dim Say1 As Integer, Say2 As Integer, Cont1 As Integer, Cont2 As Integer
Dim LstBx As MSForms.ListBox
Dim LblSira1 As MSForms.Label
Dim LblSira2 As MSForms.Label

Application.ScreenUpdating = False

'Aktarım için en az bir veri seçili olmalıdır.
Say1 = 0
For Each ctl In core_delivery_manager_UI.Frame1.Controls
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
For Each ctl In core_delivery_manager_UI.Frame2.Controls
    If TypeName(ctl) = "ListBox" Then
        Say2 = Say2 + 1
    End If
    If TypeName(ctl) = "Label" Then
        Frame2.Controls.Remove ctl.name
    End If
Next ctl

Say1 = 0
For Each ctl In core_delivery_manager_UI.Frame1.Controls
    If TypeName(ctl) = "ListBox" Then
        If ctl.Selected(0) = True Then
            Say1 = Say1 + 1
            Set LstBx = Frame2.Controls.Add("Forms.ListBox.1")
            With LstBx
                .Top = (Say2 + Say1 - 1) * 12
                .Left = 18
                .Height = 12
                .Width = 300
                .SpecialEffect = fmSpecialEffectFlat
                .MultiSelect = fmMultiSelectMulti 'fmMultiSelectSingle
                .TextAlign = fmTextAlignLeft '02092021
                .AddItem ctl.List(0)
            End With
            Frame1.Controls.Remove ctl.name
        End If
    End If
    'Sıra noları temizle
    If TypeName(ctl) = "Label" Then
        Frame1.Controls.Remove ctl.name
    End If
Next ctl
If Say1 = 0 Then
    GoTo Son
End If

ScrollTakip2 = ScrollTakip2 + Say1 * 12
ScrollTakip1 = ScrollTakip1 - Say1 * 12

If ScrollTakip2 > 180 Then
    Call SetScrollHook(Me.Frame2, ScrollTakip2)
    Frame2.ScrollTop = ScrollTakip2
Else
    Frame2.ScrollTop = 0
    RemoveScrollHook
    Frame2.ScrollBars = fmScrollBarsNone
End If

If ScrollTakip1 > 180 Then
    Call SetScrollHook(Me.Frame1, ScrollTakip1)
    Frame1.ScrollTop = Frame1.ScrollTop - Say1 * 12
Else
    Frame1.ScrollTop = 0
    RemoveScrollHook
    Frame1.ScrollBars = fmScrollBarsNone
End If

'Frame1 tekrar sırala
Say1 = 0
For Each ctl In core_delivery_manager_UI.Frame1.Controls
    If TypeName(ctl) = "ListBox" Then
        Say1 = Say1 + 1
        ctl.Top = (Say1 - 1) * 12
    End If
    
    Set LblSira1 = Frame1.Controls.Add("Forms.Label.1", "Lbl" & Say1)
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

'Frame2 sira no ver.
Say2 = 0
For Each ctl In core_delivery_manager_UI.Frame2.Controls
    If TypeName(ctl) = "ListBox" Then
        Say2 = Say2 + 1
        ctl.Top = (Say2 - 1) * 12
    End If
    
    Set LblSira2 = Frame2.Controls.Add("Forms.Label.1", "Lbl" & Say2)
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

CheckBox1.Value = False

Son:

Application.ScreenUpdating = True

End Sub

Private Sub Kaldir_Click()
Dim ctl As MSForms.Control
Dim Say1 As Integer, Say2 As Integer, Cont1 As Integer, Cont2 As Integer
Dim LstBx As MSForms.ListBox, i As Integer
Dim LblSira1 As MSForms.Label
Dim LblSira2 As MSForms.Label

Application.ScreenUpdating = False

'Aktarım için en az bir veri seçili olmalıdır.
Say2 = 0
For Each ctl In core_delivery_manager_UI.Frame2.Controls
    If TypeName(ctl) = "ListBox" Then
        If ctl.Selected(0) = True Then
            Say2 = Say2 + 1
        End If
    End If
Next ctl
If Say2 = 0 Then
    GoTo Son
End If

Say1 = 0
For Each ctl In core_delivery_manager_UI.Frame1.Controls
    If TypeName(ctl) = "ListBox" Then
        Say1 = Say1 + 1
    End If
    If TypeName(ctl) = "Label" Then
        Frame1.Controls.Remove ctl.name
    End If
Next ctl

Say2 = 0
For Each ctl In core_delivery_manager_UI.Frame2.Controls
    If TypeName(ctl) = "ListBox" Then
        If ctl.Selected(0) = True Then
            Say2 = Say2 + 1
            Set LstBx = Frame1.Controls.Add("Forms.ListBox.1")
            With LstBx
                .Top = (Say1 + Say2 - 1) * 12
                .Left = 18
                .Height = 12
                .Width = 300
                .SpecialEffect = fmSpecialEffectFlat
                .MultiSelect = fmMultiSelectMulti 'fmMultiSelectSingle
                .TextAlign = fmTextAlignLeft '02092021
                .AddItem ctl.List(0)
            End With
            Frame2.Controls.Remove ctl.name
        End If
    End If
    'Sıra noları temizle
    If TypeName(ctl) = "Label" Then
        Frame2.Controls.Remove ctl.name
    End If
Next ctl
If Say2 = 0 Then
    GoTo Son
End If

ScrollTakip1 = ScrollTakip1 + Say2 * 12
ScrollTakip2 = ScrollTakip2 - Say2 * 12

If ScrollTakip1 > 180 Then
    Call SetScrollHook(Me.Frame1, ScrollTakip1)
    Frame1.ScrollTop = ScrollTakip1
Else
    Frame1.ScrollTop = 0
    RemoveScrollHook
    Frame1.ScrollBars = fmScrollBarsNone
End If

If ScrollTakip2 > 180 Then
    Call SetScrollHook(Me.Frame2, ScrollTakip2)
    Frame2.ScrollTop = Frame2.ScrollTop - Say2 * 12
Else
    Frame2.ScrollTop = 0
    RemoveScrollHook
    Frame2.ScrollBars = fmScrollBarsNone
End If

'Frame2 tekrar sırala
Say2 = 0
For Each ctl In core_delivery_manager_UI.Frame2.Controls
    If TypeName(ctl) = "ListBox" Then
        Say2 = Say2 + 1
        ctl.Top = (Say2 - 1) * 12
    End If
    
    Set LblSira2 = Frame2.Controls.Add("Forms.Label.1", "Lbl" & Say2)
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

'Frame1 sira no ver.
Say1 = 0
For Each ctl In core_delivery_manager_UI.Frame1.Controls
    If TypeName(ctl) = "ListBox" Then
        Say1 = Say1 + 1
        ctl.Top = (Say1 - 1) * 12
    End If
    
    Set LblSira1 = Frame1.Controls.Add("Forms.Label.1", "Lbl" & Say1)
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

CheckBox2.Value = False

Son:

Application.ScreenUpdating = True

End Sub

Private Sub Kapat_Click()
    Unload Me
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
SourceTaslak = AutoPath & "\System Files\Help Documents\Delivery Manager Panel – Help.docm"
'Birden fazla kullanıcının operasyon klasöründe çakışmasını önlemek için operasyon klasörünün içinde yeni klasör aç.
DestOpUserFolderName = "Operation-" & Left(ThisWorkbook.name, InStr(ThisWorkbook.name, ".") - 1) '& " " & Format(Now(), "ddmmyyyyhhmmss")
DestOpUserFolder = DestOperasyon & DestOpUserFolderName & "\"

'System Files klasör adını kontrol et.
If Not Dir(AutoPath & "\System Files\", vbDirectory) <> vbNullString Then
    MsgBox AutoPath & "\System Files\" & " directory could not be accessed. The folder named 'System Files' might have been renamed or deleted.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

'Check the name of the Operation folder.
If Not Dir(DestOperasyon, vbDirectory) <> vbNullString Then
    MsgBox DestOperasyon & " directory could not be accessed. The folder named 'Operation' might have been renamed or deleted.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If


'RmDir DestOpUserFolder 'Sistem kapanırken DestOpUserFolder klasörünü temizle EKLENECEK!
'_______________

'Klasör isimlerini kontrol et.
If Not Dir(SourceTaslak, vbDirectory) <> vbNullString Then
    MsgBox SourceTaslak & " directory could not be accessed. The names of the folders and/or files under this directory might have been changed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
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

Private Sub TakvimBtn_Click()
Dim i As Integer, TarihBul As Range
Dim Say As Long, j As Long, Cont As Long, Sno As Variant
Dim LstBx As MSForms.ListBox, ctl As MSForms.Control
Dim LblSira1 As MSForms.Label
Dim LblSira2 As MSForms.Label

ThisWorkbook.Activate

ScrollTakip1 = 0
ScrollTakip2 = 0

Application.ScreenUpdating = False

'Columns("CE:CF").EntireColumn.Hidden = False

'suppport_calendar_UI
suppport_calendar_UI.Show

If CalTarih = "" Then
    GoTo Son
End If

'Verileri çağır başlatıldığında Frame 1 ve 2 yi boşalt
For Each ctl In core_delivery_manager_UI.Frame1.Controls
    If TypeName(ctl) = "ListBox" Then
       Frame1.Controls.Remove ctl.name
    End If
    If TypeName(ctl) = "Label" Then
        Frame1.Controls.Remove ctl.name
    End If
Next ctl
For Each ctl In core_delivery_manager_UI.Frame2.Controls
    If TypeName(ctl) = "ListBox" Then
       Frame2.Controls.Remove ctl.name
    End If
    If TypeName(ctl) = "Label" Then
        Frame2.Controls.Remove ctl.name
    End If
Next ctl


CalTarihTakip = CalTarih
ContTakip = 0

Call ModuleReport1.Rapor1TeslimTutanaklari
Call ModuleReport2.Rapor2_1TeslimTutanaklari
Call ModuleReport2.Rapor2_2BilgilendirmeTeslimTutanaklari
Call ModuleReport2.Rapor2_2XXXMudTeslimTutanaklari
Call ModuleReport2.Rapor2_2SonucTeslimTutanaklari
Call ModuleReport3.Rapor3_1TeslimTutanaklari
Call ModuleReport3.Rapor3_2TeslimTutanaklari
Call ModuleReport3.Rapor3_2TeslimTutanaklariFinansalBirim


If ScrollTakip1 > 180 Then
    Call SetScrollHook(Me.Frame1, ScrollTakip1)
    Frame1.ScrollTop = 0 'ScrollTakip1
End If
'If ScrollTakip2 > 180 Then
'    Call SetScrollHook(Me.Frame2, ScrollTakip2)
'    Frame2.ScrollTop = 0 'ScrollTakip2
'End If

Son:
CalTarih = ""

Application.ScreenUpdating = True

'Columns("CE:CF").EntireColumn.Hidden = True

End Sub

Private Sub DirektYazdir_Click()
    
    TeslimTutanaklariFormuDirektYazdir = True

    On Error GoTo Son
    TeslimTutanaklariFormuSayPrt = InputBox(Prompt:="Enter the number of copies you want to print.", Title:="Enterprise Document Automation System")
    If TeslimTutanaklariFormuSayPrt = "" Or TeslimTutanaklariFormuSayPrt = 0 Then
        GoTo Son
    End If
    If IsNumeric(TeslimTutanaklariFormuSayPrt) = False Then
        MsgBox "The print process could not be started because a non-numeric input was detected.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    If TeslimTutanaklariFormuSayPrt > 3 Then
        MsgBox "You cannot print more than 3 copies at once.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If

    
    Call Tutanak_Click
    
    TeslimTutanaklariFormuDirektYazdir = False

Son:

End Sub

Sub Tip1Proseduru()
Dim AutoPath As String, DestOperasyon As String, SourceTeslim As String, DestOpUserFolderName As String, DestOpUserFolder As String
Dim ReNameTeslim As String, OpenKontrolName As String, ContSay As Long, KontrolFile As String
Dim fso As Object, objWord As Object, objDoc As Object, Birimx As String, OpenControl As String
Dim Cont As Long, SiraBul As Range, i As Long, j As Long, Say As Long
'Dim Ctl As MSForms.Control
Dim Ctlx As MSForms.Control, Say2 As Integer, GidenTema As String, IlkSira As Long
Dim MyRange As Object, AdetSay As Integer, VeriKontrol As Integer
Dim EkFormu As String, SourceTip5 As String, ReNameTip5 As String, EkFormuTip5 As String
Dim a() As Variant, b As Variant

Dim SonSiraBul As Range, SonSira As Long
Dim GidenPaketAdet As Integer, Tutanak1Sayfa As Integer, DokumSayfa As Integer, RaporSayfa As Integer, Tutanak2Sayfa As Integer
Dim KimlikSayfa As Integer, DesteBandiFotoSayfa As Integer, TesDekFotoSayfa As Integer, IlgiYaziFotoSayfa As Integer
Dim ItemBul As Range, Unvan1 As String, Unvan2 As String, Unvan3 As String, Tip1Ekler As String
Dim StrContent, BstrContent As String, x As Integer
Dim Rapor2_2RaporSayfa, DigitalIcerikAdedi, AnalizSayfa As Integer
Dim TesDuzDekSayfa, TutanakSayfa, DesteBandiAdet As Integer
Dim FinansalBirimSayfa As Integer

'Worksheets(3).Activate
Application.ScreenUpdating = False
Application.DisplayAlerts = False
'Application.EnableEvents = False

ThisWorkbook.Unprotect "123"
Worksheets(3).Unprotect Password:="123"
Worksheets(4).Unprotect Password:="123"
Worksheets(5).Unprotect Password:="123"

    
'Columns("CE:CF").EntireColumn.Hidden = False
'Columns("CK:CN").EntireColumn.Hidden = False

If TutanakTarihiText.Value = "" Then
    MsgBox "The process cannot be completed because the statement date has not been specified.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

VeriKontrol = 0
For Each ctl In core_delivery_manager_UI.Frame2.Controls
    If TypeName(ctl) = "ListBox" Then
        VeriKontrol = VeriKontrol + 1
        GoTo Devam2
    End If
Next ctl
Devam2:
If VeriKontrol = 0 Then
    MsgBox "The process cannot be completed because no transfer has been made to the 'To Be Sent' section, which is required for preparing Statement Type 1.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

'Pathfinder...
AutoPath = ThisWorkbook.Path
DestOperasyon = AutoPath & "\System Files\Operation\"
'EK TANIMLARI
If Tip1Option.Value = True Then
    EkFormu = "Type 1 Delivery Statement"
Else
    MsgBox "The process cannot be completed because the form type has not been selected from the upper-right corner.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

SourceTeslim = AutoPath & "\System Files\System Templates\Acceptance and Delivery Statements\" & EkFormu & ".docm"

'Birden fazla kullanıcının operasyon klasöründe çakışmasını önlemek için operasyon klasörünün içinde yeni klasör aç.
DestOpUserFolderName = "Operation-" & Left(ThisWorkbook.name, InStr(ThisWorkbook.name, ".") - 1) '& " " & Format(Now(), "ddmmyyyyhhmmss")
DestOpUserFolder = DestOperasyon & DestOpUserFolderName & "\"

'Check the folder name "System Files".
If Not Dir(AutoPath & "\System Files\", vbDirectory) <> vbNullString Then
    MsgBox AutoPath & "\System Files\" & " directory cannot be accessed. The folder named 'System Files' may have been renamed or deleted.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

'Check the "Operation" folder name.
If Not Dir(DestOperasyon, vbDirectory) <> vbNullString Then
    MsgBox DestOperasyon & " directory cannot be accessed. The folder named 'Operation' may have been renamed or the 'System Files' folder may have been deleted.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

'Check the folder names.
If Not Dir(SourceTeslim, vbDirectory) <> vbNullString Then
    MsgBox SourceTeslim & " directory cannot be accessed. The folder and/or file names under this directory may have been renamed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

'Operation klasörü içinde kullanıcı modülü klasörü yoksa oluştur.
If Not Dir(DestOpUserFolder, vbDirectory) <> vbNullString Then
    MkDir DestOpUserFolder
End If

ReNameTeslim = EkFormu

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
fso.CopyFile (SourceTeslim), DestOpUserFolder & ReNameTeslim & ".docm", True
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
objWord.Documents.Open FileName:=DestOpUserFolder & ReNameTeslim & ".docm"
objWord.Visible = True
objWord.Activate 'Ekrana getirir.
'objDoc.Activate 'Ekrana getirmez.
'objWord.Application.WindowState = wdWindowStateMaximize
Set objDoc = GetObject(DestOpUserFolder & ReNameTeslim & ".docm")
'________________________________________

'Birim
Birimx = UCase(Replace(Replace(Worksheets(2).Cells(6, 99).Value, "i", "I"), "ı", "I")) & " UNIT"
objDoc.Tables(1).Cell(Row:=1, Column:=1).Range.Text = Birimx
''Tutanak tarihi
'objDoc.Tables(1).Cell(Row:=4, Column:=3).Range.Text = TutanakTarihiText.Value

'________________AKTARIMLAR 1 (Başlangıç)

Say2 = 0
AdetSay = 0
For Each ctl In core_delivery_manager_UI.Frame2.Controls '________________RAPOR1 BÖLÜMÜ (Başlangıç, Tüm Raporların Ortak Bölümü)
    If TypeName(ctl) = "ListBox" Then
        'MsgBox Mid(Ctl.List(0), InStr(Ctl.List(0), "|") + 1, 2)
        If Mid(ctl.List(0), InStr(ctl.List(0), "|") + 1, 2) = " Ö" Then
            
            Set SonSiraBul = ThisWorkbook.Worksheets(3).Range("CF7:CF100000").Find(What:=Left(ctl.List(0), InStr(ctl.List(0), "|") - 2), SearchDirection:=xlNext, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not SonSiraBul Is Nothing Then
                SonSira = SonSiraBul.Row
            Else
                GoTo Tip1Rapor1Devam1
            End If

            Call Rapor1GidenTema
            If IlkSiraGlobal = 0 Then
                GoTo Tip1Rapor1Devam1
            End If
            IlkSira = IlkSiraGlobal
            GidenTema = GidenTemaGlobal
            
            'Tabloya satır ekle
            Say2 = Say2 + 1
            'MsgBox Left(Ctl.List(0), InStr(Ctl.List(0), "|") - 2)
            If Say2 > 1 Then
                objDoc.Tables(2).Rows.Add BeforeRow:=objDoc.Tables(2).Rows(Say2 + 1)
            End If
            'Sıra no
            objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=1).Range.Text = Say2
            'Gönderilen birim
            objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=4).Range.Text = GidenTema
            'kayıt no
            Tip1Ekler = ""
            For i = IlkSira To SonSira
                If Tip1Ekler = "" Then
                    If ThisWorkbook.Worksheets(3).Cells(i, 59).Value <> "" Then
                        Tip1Ekler = "ÖR/" & ThisWorkbook.Worksheets(3).Cells(i, 59).Value
                    End If
                Else
                    If ThisWorkbook.Worksheets(3).Cells(i, 59).Value <> "" Then
                        Tip1Ekler = Tip1Ekler & ", ÖR/" & ThisWorkbook.Worksheets(3).Cells(i, 59).Value
                    End If
                End If
            Next i
            objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=3).Range.Text = Tip1Ekler 'ThisWorkbook.Worksheets(3).Cells(IlkSira, 59).Value
            'Tutanak2 tarihi
            objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=2).Range.Text = TutanakTarihiText.Value
            'Ek sayısı (üst yazının)
            IlgiYaziFotoSayfa = Application.Sum(ThisWorkbook.Worksheets(3).Range(ThisWorkbook.Worksheets(3).Cells(IlkSira, 74), ThisWorkbook.Worksheets(3).Cells(SonSira, 74)))
            GidenPaketAdet = Application.Sum(ThisWorkbook.Worksheets(3).Range(ThisWorkbook.Worksheets(3).Cells(IlkSira, 68), ThisWorkbook.Worksheets(3).Cells(SonSira, 68)))
            Tutanak1Sayfa = Application.Sum(ThisWorkbook.Worksheets(3).Range(ThisWorkbook.Worksheets(3).Cells(IlkSira, 89), ThisWorkbook.Worksheets(3).Cells(SonSira, 89)))
            DokumSayfa = Application.Sum(ThisWorkbook.Worksheets(3).Range(ThisWorkbook.Worksheets(3).Cells(IlkSira, 90), ThisWorkbook.Worksheets(3).Cells(SonSira, 90)))
            RaporSayfa = Application.Sum(ThisWorkbook.Worksheets(3).Range(ThisWorkbook.Worksheets(3).Cells(IlkSira, 91), ThisWorkbook.Worksheets(3).Cells(SonSira, 91)))
            Tutanak2Sayfa = Application.Sum(ThisWorkbook.Worksheets(3).Range(ThisWorkbook.Worksheets(3).Cells(IlkSira, 92), ThisWorkbook.Worksheets(3).Cells(SonSira, 92)))
            objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=5).Range.Text = IlgiYaziFotoSayfa + GidenPaketAdet + Tutanak1Sayfa + DokumSayfa + RaporSayfa + Tutanak2Sayfa
            
            'Kilitlenme bilgisi
            'ThisWorkbook.Worksheets(3).Cells(IlkSira, 86).Value = "*"
        End If
    End If
Tip1Rapor1Devam1:
Next ctl '________________RAPOR1 BÖLÜMÜ (Bitiş, Tüm Raporların Ortak Bölümü)


For Each ctl In core_delivery_manager_UI.Frame2.Controls '________________RAPOR BÖLÜMÜ (Başlangıç, Tüm Raporların Ortak Bölümü)'Rapor ve Bilgilendirme
    If TypeName(ctl) = "ListBox" Then
        'MsgBox Mid(Ctl.List(0), InStr(Ctl.List(0), "|") + 1, 2)
        If Mid(ctl.List(0), InStr(ctl.List(0), "|") + 1, 2) = " R" Then

            Set SonSiraBul = ThisWorkbook.Worksheets(4).Range("CN7:CN100000").Find(What:=Left(ctl.List(0), InStr(ctl.List(0), "|") - 2), SearchDirection:=xlNext, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not SonSiraBul Is Nothing Then
                SonSira = SonSiraBul.Row
            Else
                GoTo Tip1RaporDevam1
            End If
            
            Call RaporGidenTema
            If IlkSiraGlobal = 0 Then
                GoTo Tip1RaporDevam1
            End If
            IlkSira = IlkSiraGlobal
            GidenTema = GidenTemaGlobal
            
            'XXXMud, bilgilendirme ve sonuç yazıları için süreç
            If ThisWorkbook.Worksheets(4).Cells(IlkSira, 174).Value = "Yes" Then
                GoTo Bilgilendirme
            End If
            
            'Tabloya satır ekle
            Say2 = Say2 + 1
            'MsgBox Left(Ctl.List(0), InStr(Ctl.List(0), "|") - 2)
            If Say2 > 1 Then
                objDoc.Tables(2).Rows.Add BeforeRow:=objDoc.Tables(2).Rows(Say2 + 1)
            End If
            'Sıra no
            objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=1).Range.Text = Say2
            'Gönderilen birim
            objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=4).Range.Text = GidenTema
            'Çıkış kayıt no
            Tip1Ekler = ""
            For i = IlkSira To SonSira
                If Tip1Ekler = "" Then
                    If ThisWorkbook.Worksheets(4).Cells(i, 67).Value <> "" Then
                        Tip1Ekler = "R/" & ThisWorkbook.Worksheets(4).Cells(i, 67).Value
                    End If
                Else
                    If ThisWorkbook.Worksheets(4).Cells(i, 67).Value <> "" Then
                        Tip1Ekler = Tip1Ekler & ", R/" & ThisWorkbook.Worksheets(4).Cells(i, 67).Value
                    End If
                End If
            Next i
            objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=3).Range.Text = Tip1Ekler 'ThisWorkbook.Worksheets(3).Cells(IlkSira, 59).Value
            'Tutanak2 tarihi
            objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=2).Range.Text = TutanakTarihiText.Value
            'Ek sayısı (üst yazının)
            IlgiYaziFotoSayfa = Application.Sum(ThisWorkbook.Worksheets(4).Range(ThisWorkbook.Worksheets(4).Cells(IlkSira, 82), ThisWorkbook.Worksheets(4).Cells(SonSira, 82)))
            GidenPaketAdet = Application.Sum(ThisWorkbook.Worksheets(4).Range(ThisWorkbook.Worksheets(4).Cells(IlkSira, 76), ThisWorkbook.Worksheets(4).Cells(SonSira, 76)))
            Tutanak1Sayfa = Application.Sum(ThisWorkbook.Worksheets(4).Range(ThisWorkbook.Worksheets(4).Cells(IlkSira, 97), ThisWorkbook.Worksheets(4).Cells(SonSira, 97)))
            DokumSayfa = Application.Sum(ThisWorkbook.Worksheets(4).Range(ThisWorkbook.Worksheets(4).Cells(IlkSira, 98), ThisWorkbook.Worksheets(4).Cells(SonSira, 98)))
            RaporSayfa = Application.Sum(ThisWorkbook.Worksheets(4).Range(ThisWorkbook.Worksheets(4).Cells(IlkSira, 99), ThisWorkbook.Worksheets(4).Cells(SonSira, 99)))
            Tutanak2Sayfa = Application.Sum(ThisWorkbook.Worksheets(4).Range(ThisWorkbook.Worksheets(4).Cells(IlkSira, 100), ThisWorkbook.Worksheets(4).Cells(SonSira, 100)))
            objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=5).Range.Text = IlgiYaziFotoSayfa + GidenPaketAdet + Tutanak1Sayfa + DokumSayfa + RaporSayfa + Tutanak2Sayfa
            'Kilitlenme bilgisi
            'ThisWorkbook.Worksheets(4).Cells(IlkSira, 86).Value = "*"
            GoTo Tip1RaporDevam1
            
Bilgilendirme:

            'BİLGİLENDİRME
            'Tabloya satır ekle
            Say2 = Say2 + 1
            'MsgBox Left(Ctl.List(0), InStr(Ctl.List(0), "|") - 2)
            If Say2 > 1 Then
                objDoc.Tables(2).Rows.Add BeforeRow:=objDoc.Tables(2).Rows(Say2 + 1)
            End If
            'Sıra no
            objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=1).Range.Text = Say2
            'Gönderilen birim
            objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=4).Range.Text = GidenTema
            'Çıkış kayıt no
            Tip1Ekler = ""
            For i = IlkSira To SonSira
                If Tip1Ekler = "" Then
                    If ThisWorkbook.Worksheets(4).Cells(i, 67).Value <> "" Then
                        Tip1Ekler = "R/" & ThisWorkbook.Worksheets(4).Cells(i, 67).Value
                    End If
                Else
                    If ThisWorkbook.Worksheets(4).Cells(i, 67).Value <> "" Then
                        Tip1Ekler = Tip1Ekler & ", R/" & ThisWorkbook.Worksheets(4).Cells(i, 67).Value
                    End If
                End If
            Next i
            Tip1Ekler = Tip1Ekler & " (B/" & ThisWorkbook.Worksheets(4).Cells(IlkSira, 65).Value & ")"
            objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=3).Range.Text = Tip1Ekler 'ThisWorkbook.Worksheets(3).Cells(IlkSira, 59).Value
            'Tutanak2 tarihi
            objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=2).Range.Text = TutanakTarihiText.Value
            'Ek sayısı (üst yazının)
            IlgiYaziFotoSayfa = Application.Sum(ThisWorkbook.Worksheets(4).Range(ThisWorkbook.Worksheets(4).Cells(IlkSira, 206), ThisWorkbook.Worksheets(4).Cells(SonSira, 206)))
            Tutanak1Sayfa = Application.Sum(ThisWorkbook.Worksheets(4).Range(ThisWorkbook.Worksheets(4).Cells(IlkSira, 221), ThisWorkbook.Worksheets(4).Cells(SonSira, 221)))
            'Yukarıdaki Tutanak1Sayfa,esasında ilgi b (XXXMud'ye yazılan üst yazının) yazı fotokopisidir.
            RaporSayfa = Application.Sum(ThisWorkbook.Worksheets(4).Range(ThisWorkbook.Worksheets(4).Cells(IlkSira, 219), ThisWorkbook.Worksheets(4).Cells(SonSira, 219)))
            objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=5).Range.Text = IlgiYaziFotoSayfa + Tutanak1Sayfa + RaporSayfa
            'Kilitlenme bilgisi
            'ThisWorkbook.Worksheets(4).Cells(IlkSira, 86).Value = "*"
            
        End If
    End If
Tip1RaporDevam1:
Next ctl '________________RAPOR BÖLÜMÜ (Bitiş, Tüm Raporların Ortak Bölümü)


For Each ctl In core_delivery_manager_UI.Frame2.Controls '________________RAPOR BÖLÜMÜ (Başlangıç, Tüm Raporların Ortak Bölümü)'XXXMud
    If TypeName(ctl) = "ListBox" Then
        'MsgBox Mid(Ctl.List(0), InStr(Ctl.List(0), "|") + 1, 2)
        If Mid(ctl.List(0), InStr(ctl.List(0), "|") + 1, 2) = " B" And InStr(ctl.List(0), "XXXMud") <> 0 Then
            Set SonSiraBul = ThisWorkbook.Worksheets(4).Range("CN7:CN100000").Find(What:=Left(ctl.List(0), InStr(ctl.List(0), "|") - 2), SearchDirection:=xlNext, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not SonSiraBul Is Nothing Then
                SonSira = SonSiraBul.Row
            Else
                GoTo Tip1RaporDevam2
            End If
            
            Call RaporGidenTema
            If IlkSiraGlobal = 0 Then
                GoTo Tip1RaporDevam2
            End If
            IlkSira = IlkSiraGlobal
            GidenTema = GidenTemaGlobal

            'XXXMud
            'Tabloya satır ekle
            Say2 = Say2 + 1
            'MsgBox Left(Ctl.List(0), InStr(Ctl.List(0), "|") - 2)
            If Say2 > 1 Then
                objDoc.Tables(2).Rows.Add BeforeRow:=objDoc.Tables(2).Rows(Say2 + 1)
            End If
            'Sıra no
            objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=1).Range.Text = Say2
            'Gönderilen birim
            GidenTema = "ORGANIZATION A XXX Directorate"
            objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=4).Range.Text = GidenTema
            'Çıkış kayıt no
            Tip1Ekler = ""

            If ThisWorkbook.Worksheets(4).Cells(IlkSira, 187).Value = "All" Then
                For i = IlkSira To SonSira
                    If Tip1Ekler = "" Then
                        If ThisWorkbook.Worksheets(4).Cells(i, 67).Value <> "" Then
                            Tip1Ekler = "R/" & ThisWorkbook.Worksheets(4).Cells(i, 67).Value
                        End If
                    Else
                        If ThisWorkbook.Worksheets(4).Cells(i, 67).Value <> "" Then
                            Tip1Ekler = Tip1Ekler & ", R/" & ThisWorkbook.Worksheets(4).Cells(i, 67).Value
                        End If
                    End If
                Next i
            ElseIf ThisWorkbook.Worksheets(4).Cells(IlkSira, 187).Value = "Technique A" Then
                For i = IlkSira To SonSira
                    If Left(ThisWorkbook.Worksheets(4).Cells(i, 63).Value, 11) = "Technique A" Then
                        If Tip1Ekler = "" Then
                            If ThisWorkbook.Worksheets(4).Cells(i, 67).Value <> "" Then
                                Tip1Ekler = "R/" & ThisWorkbook.Worksheets(4).Cells(i, 67).Value
                            End If
                        Else
                            If ThisWorkbook.Worksheets(4).Cells(i, 67).Value <> "" Then
                                Tip1Ekler = Tip1Ekler & ", R/" & ThisWorkbook.Worksheets(4).Cells(i, 67).Value
                            End If
                        End If
                    End If
                Next i
            End If
            

            Tip1Ekler = "B/" & ThisWorkbook.Worksheets(4).Cells(IlkSira, 65).Value & " (" & Tip1Ekler & ")"
            objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=3).Range.Text = Tip1Ekler 'ThisWorkbook.Worksheets(3).Cells(IlkSira, 59).Value
            'Tutanak2 tarihi
            objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=2).Range.Text = TutanakTarihiText.Value
            'Ek sayısı (üst yazının)
            IlgiYaziFotoSayfa = Application.Sum(ThisWorkbook.Worksheets(4).Range(ThisWorkbook.Worksheets(4).Cells(IlkSira, 206), ThisWorkbook.Worksheets(4).Cells(SonSira, 206)))
            Tutanak2Sayfa = Application.Sum(ThisWorkbook.Worksheets(4).Range(ThisWorkbook.Worksheets(4).Cells(IlkSira, 220), ThisWorkbook.Worksheets(4).Cells(SonSira, 220)))
            GidenPaketAdet = Application.Sum(ThisWorkbook.Worksheets(4).Range(ThisWorkbook.Worksheets(4).Cells(IlkSira, 185), ThisWorkbook.Worksheets(4).Cells(SonSira, 185)))
            objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=5).Range.Text = IlgiYaziFotoSayfa + GidenPaketAdet + Tutanak2Sayfa
            'MsgBox "1:" & IlgiYaziFotoSayfa & " 2:" & Tutanak2Sayfa & " 3:" & GidenPaketAdet
            'Kilitlenme bilgisi
            'ThisWorkbook.Worksheets(4).Cells(IlkSira, 86).Value = "*"
        End If
    End If
Tip1RaporDevam2:
Next ctl '________________RAPOR BÖLÜMÜ (Bitiş, Tüm Raporların Ortak Bölümü)


For Each ctl In core_delivery_manager_UI.Frame2.Controls '________________RAPOR BÖLÜMÜ (Başlangıç, Tüm Raporların Ortak Bölümü)'Sonuç
    If TypeName(ctl) = "ListBox" Then
        'MsgBox Mid(Ctl.List(0), InStr(Ctl.List(0), "|") + 1, 2)
        If Mid(ctl.List(0), InStr(ctl.List(0), "|") + 1, 2) = " B" And InStr(ctl.List(0), "XXXMud") = 0 Then
'InStr(ItemName, "-") <> 0
            Set SonSiraBul = ThisWorkbook.Worksheets(4).Range("CN7:CN100000").Find(What:=Left(ctl.List(0), InStr(ctl.List(0), "|") - 2), SearchDirection:=xlNext, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not SonSiraBul Is Nothing Then
                SonSira = SonSiraBul.Row
            Else
                GoTo Tip1RaporDevam3
            End If
            
            Call Rapor2_2GidenTema
            If IlkSiraGlobal = 0 Then
                GoTo Tip1RaporDevam3
            End If
            IlkSira = IlkSiraGlobal
            GidenTema = GidenTemaGlobal

            'SONUÇ
            'Tabloya satır ekle
            Say2 = Say2 + 1
            'MsgBox Left(Ctl.List(0), InStr(Ctl.List(0), "|") - 2)
            If Say2 > 1 Then
                objDoc.Tables(2).Rows.Add BeforeRow:=objDoc.Tables(2).Rows(Say2 + 1)
            End If
            'Sıra no
            objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=1).Range.Text = Say2
            'Gönderilen birim
            objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=4).Range.Text = GidenTema
            'Çıkış kayıt no
            Tip1Ekler = ""
            For i = IlkSira To SonSira
                If Tip1Ekler = "" Then
                    If ThisWorkbook.Worksheets(4).Cells(i, 67).Value <> "" Then
                        Tip1Ekler = "R/" & ThisWorkbook.Worksheets(4).Cells(i, 67).Value
                    End If
                Else
                    If ThisWorkbook.Worksheets(4).Cells(i, 67).Value <> "" Then
                        Tip1Ekler = Tip1Ekler & ", R/" & ThisWorkbook.Worksheets(4).Cells(i, 67).Value
                    End If
                End If
            Next i

            'Rapor2_2 rapor numaralarının önüne B harfi gelecek
            'BstrContent = ""
            StrContent = "B/" & ThisWorkbook.Worksheets(4).Cells(IlkSira, 65).Value
'            For i = 1 To Len(StrContent)
'                x = x + 1
'                BstrContent = BstrContent & Mid(StrContent, i, 1)
'                If Mid(StrContent, i, 1) = " " Then
'                    BstrContent = Left(BstrContent, x) & "B/"
'                    x = x + 2
'                    'MsgBox BstrContent
'                End If
'            Next i
            'Tip1Ekler = BstrContent & " (" & Tip1Ekler & ")"
            Tip1Ekler = Tip1Ekler & " (" & StrContent & ")"
            
            
            objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=3).Range.Text = Tip1Ekler 'ThisWorkbook.Worksheets(3).Cells(IlkSira, 59).Value
            'Tutanak2 tarihi
            objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=2).Range.Text = TutanakTarihiText.Value
            'Ek sayısı (üst yazının)
            Rapor2_2RaporSayfa = Application.Sum(ThisWorkbook.Worksheets(4).Range(ThisWorkbook.Worksheets(4).Cells(IlkSira, 173), ThisWorkbook.Worksheets(4).Cells(SonSira, 173)))
            DigitalIcerikAdedi = Application.Sum(ThisWorkbook.Worksheets(4).Range(ThisWorkbook.Worksheets(4).Cells(IlkSira, 181), ThisWorkbook.Worksheets(4).Cells(SonSira, 181)))
            AnalizSayfa = Application.Sum(ThisWorkbook.Worksheets(4).Range(ThisWorkbook.Worksheets(4).Cells(IlkSira, 182), ThisWorkbook.Worksheets(4).Cells(SonSira, 182)))
            Tutanak1Sayfa = Application.Sum(ThisWorkbook.Worksheets(4).Range(ThisWorkbook.Worksheets(4).Cells(IlkSira, 217), ThisWorkbook.Worksheets(4).Cells(SonSira, 217)))
            Tutanak2Sayfa = Application.Sum(ThisWorkbook.Worksheets(4).Range(ThisWorkbook.Worksheets(4).Cells(IlkSira, 224), ThisWorkbook.Worksheets(4).Cells(SonSira, 225)))
            GidenPaketAdet = Application.Sum(ThisWorkbook.Worksheets(4).Range(ThisWorkbook.Worksheets(4).Cells(IlkSira, 76), ThisWorkbook.Worksheets(4).Cells(SonSira, 76)))
            GidenPaketAdet = GidenPaketAdet + Application.Sum(ThisWorkbook.Worksheets(4).Range(ThisWorkbook.Worksheets(4).Cells(IlkSira, 202), ThisWorkbook.Worksheets(4).Cells(SonSira, 202)))
            
            objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=5).Range.Text = Rapor2_2RaporSayfa + DigitalIcerikAdedi + AnalizSayfa + Tutanak1Sayfa + Tutanak2Sayfa + GidenPaketAdet
            'Kilitlenme bilgisi
            'ThisWorkbook.Worksheets(4).Cells(IlkSira, 86).Value = "*"
        End If
    End If
Tip1RaporDevam3:
Next ctl '________________RAPOR BÖLÜMÜ (Bitiş, Tüm Raporların Ortak Bölümü)

For Each ctl In core_delivery_manager_UI.Frame2.Controls '________________RAPOR3_1 BÖLÜMÜ (Başlangıç, Tüm Raporların Ortak Bölümü)
    If TypeName(ctl) = "ListBox" Then
        'MsgBox Mid(Ctl.List(0), InStr(Ctl.List(0), "|") + 1, 2)
        If Mid(ctl.List(0), InStr(ctl.List(0), "|") + 1, 2) = " G" Then

            Set SonSiraBul = ThisWorkbook.Worksheets(5).Range("FH7:FH100000").Find(What:=Left(ctl.List(0), InStr(ctl.List(0), "|") - 2), SearchDirection:=xlNext, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not SonSiraBul Is Nothing Then
                SonSira = SonSiraBul.Row
            Else
                GoTo Tip1Rapor3_1Devam1
            End If

            Call Rapor3_1GidenTema
            If IlkSiraGlobal = 0 Then
                GoTo Tip1Rapor3_1Devam1
            End If
            IlkSira = IlkSiraGlobal
            GidenTema = GidenTemaGlobal
            
            'Tabloya satır ekle
            Say2 = Say2 + 1
            'MsgBox Left(Ctl.List(0), InStr(Ctl.List(0), "|") - 2)
            If Say2 > 1 Then
                objDoc.Tables(2).Rows.Add BeforeRow:=objDoc.Tables(2).Rows(Say2 + 1)
            End If
            'Sıra no
            objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=1).Range.Text = Say2
            'Gönderilen birim
            objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=4).Range.Text = GidenTema
            
            'Çıkış kayıt no 'Tema/Temax no
            Tip1Ekler = ""
            For i = IlkSira To SonSira
                If Tip1Ekler = "" Then
                    If ThisWorkbook.Worksheets(5).Cells(i, 13).Value <> "" Then
                        Tip1Ekler = "R/" & ThisWorkbook.Worksheets(5).Cells(i, 13).Value
                    End If
                Else
                    If ThisWorkbook.Worksheets(5).Cells(i, 13).Value <> "" Then
                        Tip1Ekler = Tip1Ekler & ", R/" & ThisWorkbook.Worksheets(5).Cells(i, 13).Value
                    End If
                End If
            Next i
            If Tip1Ekler = "" Then
                objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=3).Range.Text = ThisWorkbook.Worksheets(5).Cells(IlkSira, 98).Value
            Else
                objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=3).Range.Text = Tip1Ekler 'ThisWorkbook.Worksheets(5).Cells(IlkSira, 98).Value
            End If
            
            'Tutanak2 tarihi
            objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=2).Range.Text = TutanakTarihiText.Value
            'Ek sayısı (üst yazının)
            KimlikSayfa = Application.Sum(ThisWorkbook.Worksheets(5).Range(ThisWorkbook.Worksheets(5).Cells(IlkSira, 126), ThisWorkbook.Worksheets(5).Cells(SonSira, 126)))
            GidenPaketAdet = Application.Sum(ThisWorkbook.Worksheets(5).Range(ThisWorkbook.Worksheets(5).Cells(IlkSira, 150), ThisWorkbook.Worksheets(5).Cells(SonSira, 150)))
            Tutanak1Sayfa = Application.Sum(ThisWorkbook.Worksheets(5).Range(ThisWorkbook.Worksheets(5).Cells(IlkSira, 169), ThisWorkbook.Worksheets(5).Cells(SonSira, 169)))
            DokumSayfa = Application.Sum(ThisWorkbook.Worksheets(5).Range(ThisWorkbook.Worksheets(5).Cells(IlkSira, 170), ThisWorkbook.Worksheets(5).Cells(SonSira, 170)))
            RaporSayfa = Application.Sum(ThisWorkbook.Worksheets(5).Range(ThisWorkbook.Worksheets(5).Cells(IlkSira, 174), ThisWorkbook.Worksheets(5).Cells(SonSira, 174)))
            Tutanak2Sayfa = Application.Sum(ThisWorkbook.Worksheets(5).Range(ThisWorkbook.Worksheets(5).Cells(IlkSira, 171), ThisWorkbook.Worksheets(5).Cells(SonSira, 171)))
            objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=5).Range.Text = KimlikSayfa + GidenPaketAdet + Tutanak1Sayfa + DokumSayfa + RaporSayfa + Tutanak2Sayfa
            'Kilitlenme bilgisi
            'ThisWorkbook.Worksheets(5).Cells(IlkSira, 166).Value = "*"
        End If
    End If
Tip1Rapor3_1Devam1:
Next ctl '________________RAPOR3_1 BÖLÜMÜ (Bitiş, Tüm Raporların Ortak Bölümü)


For Each ctl In core_delivery_manager_UI.Frame2.Controls '________________RAPOR3_2 BÖLÜMÜ (Başlangıç, Tüm Raporların Ortak Bölümü)
    If TypeName(ctl) = "ListBox" Then
        'MsgBox Mid(Ctl.List(0), InStr(Ctl.List(0), "|") + 1, 2)
        If Mid(ctl.List(0), InStr(ctl.List(0), "|") + 1, 2) = " M" Then 'And Mid(ctl.List(0), InStrRev(ctl.List(0), "-") - 5, 5) <> "FinansalBirim" Then 'And InStr(ctl.List(0), "(FinansalBirim-TipA)") = 0 Then

            'If Mid(ctl.List(0), InStrRev(ctl.List(0), "-") - 5, 5) = "FinansalBirim" Then 'FinansalBirim
            If InStr(ctl.List(0), "(FinansalBirim-TipA)") <> 0 Or InStr(ctl.List(0), "(FinansalBirim-TipB)") <> 0 Then
                
                Set SonSiraBul = ThisWorkbook.Worksheets(5).Range("FH7:FH100000").Find(What:=Left(ctl.List(0), InStr(ctl.List(0), "|") - 2), SearchDirection:=xlNext, _
                                SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
                If Not SonSiraBul Is Nothing Then
                    SonSira = SonSiraBul.Row
                Else
                    GoTo Tip1Rapor3_2Devam1
                End If
                
                Call Rapor3_2GidenTema
                If IlkSiraGlobal = 0 Then
                    GoTo Tip1Rapor3_2Devam1
                End If
                IlkSira = IlkSiraGlobal
                GidenTema = ThisWorkbook.Worksheets(5).Cells(IlkSira, 30).Value & " " & ThisWorkbook.Worksheets(5).Cells(IlkSira, 31).Value 'GidenTemaGlobal
                
                'Tabloya satır ekle
                Say2 = Say2 + 1
                'MsgBox Left(Ctl.List(0), InStr(Ctl.List(0), "|") - 2)
                If Say2 > 1 Then
                    objDoc.Tables(2).Rows.Add BeforeRow:=objDoc.Tables(2).Rows(Say2 + 1)
                End If
                'Sıra no
                objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=1).Range.Text = Say2
                'Gönderilen birim
               ' GidenTema = ThisWorkbook.Worksheets(5).Cells(IlkSira, 30).Value
                objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=4).Range.Text = GidenTema
                
                'Çıkış kayıt no
                Tip1Ekler = ""
                For i = IlkSira To SonSira
                    If Tip1Ekler = "" Then
                        If ThisWorkbook.Worksheets(5).Cells(i, 13).Value <> "" Then
                            Tip1Ekler = "R/" & ThisWorkbook.Worksheets(5).Cells(i, 13).Value
                        End If
                    Else
                        If ThisWorkbook.Worksheets(5).Cells(i, 13).Value <> "" Then
                            Tip1Ekler = Tip1Ekler & ", R/" & ThisWorkbook.Worksheets(5).Cells(i, 13).Value
                        End If
                    End If
                Next i
                If Tip1Ekler = "" Then
                    objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=3).Range.Text = ThisWorkbook.Worksheets(5).Cells(IlkSira, 26).Value
                Else
                    objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=3).Range.Text = Tip1Ekler
                End If
                
                'Tutanak2 tarihi
                objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=2).Range.Text = TutanakTarihiText.Value
                'Ek sayısı (üst yazının)
                TesDuzDekSayfa = Application.Sum(ThisWorkbook.Worksheets(5).Range(ThisWorkbook.Worksheets(5).Cells(IlkSira, 80), ThisWorkbook.Worksheets(5).Cells(SonSira, 80)))
                TutanakSayfa = Application.Sum(ThisWorkbook.Worksheets(5).Range(ThisWorkbook.Worksheets(5).Cells(IlkSira, 169), ThisWorkbook.Worksheets(5).Cells(SonSira, 169)))
                DesteBandiAdet = Application.Sum(ThisWorkbook.Worksheets(5).Range(ThisWorkbook.Worksheets(5).Cells(IlkSira, 81), ThisWorkbook.Worksheets(5).Cells(SonSira, 81)))
                DokumSayfa = Application.Sum(ThisWorkbook.Worksheets(5).Range(ThisWorkbook.Worksheets(5).Cells(IlkSira, 50), ThisWorkbook.Worksheets(5).Cells(SonSira, 50)))
                objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=5).Range.Text = TesDuzDekSayfa + TutanakSayfa + DesteBandiAdet + DokumSayfa
                'Kilitlenme bilgisi
                'ThisWorkbook.Worksheets(5).Cells(IlkSira, 166).Value = "*"
            
            Else 'DIRECTORATELIK
            
                Set SonSiraBul = ThisWorkbook.Worksheets(5).Range("FH7:FH100000").Find(What:=Left(ctl.List(0), InStr(ctl.List(0), "|") - 2), SearchDirection:=xlNext, _
                                SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
                If Not SonSiraBul Is Nothing Then
                    SonSira = SonSiraBul.Row
                Else
                    GoTo Tip1Rapor3_2Devam1
                End If
    
                Call Rapor3_2GidenTema
                If IlkSiraGlobal = 0 Then
                    GoTo Tip1Rapor3_2Devam1
                End If
                IlkSira = IlkSiraGlobal
                GidenTema = GidenTemaGlobal
                
                'Tabloya satır ekle
                Say2 = Say2 + 1
                'MsgBox Left(Ctl.List(0), InStr(Ctl.List(0), "|") - 2)
                If Say2 > 1 Then
                    objDoc.Tables(2).Rows.Add BeforeRow:=objDoc.Tables(2).Rows(Say2 + 1)
                End If
                'Sıra no
                objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=1).Range.Text = Say2
                'Gönderilen birim
                objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=4).Range.Text = GidenTema
                
                'Çıkış kayıt no
                Tip1Ekler = ""
                For i = IlkSira To SonSira
                    If Tip1Ekler = "" Then
                        If ThisWorkbook.Worksheets(5).Cells(i, 13).Value <> "" Then
                            Tip1Ekler = "R/" & ThisWorkbook.Worksheets(5).Cells(i, 13).Value
                        End If
                    Else
                        If ThisWorkbook.Worksheets(5).Cells(i, 13).Value <> "" Then
                            Tip1Ekler = Tip1Ekler & ", R/" & ThisWorkbook.Worksheets(5).Cells(i, 13).Value
                        End If
                    End If
                Next i
                If Tip1Ekler = "" Then
                    objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=3).Range.Text = ThisWorkbook.Worksheets(5).Cells(IlkSira, 26).Value
                Else
                    objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=3).Range.Text = Tip1Ekler
                End If
                
                'Tutanak2 tarihi
                objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=2).Range.Text = TutanakTarihiText.Value
                'Ek sayısı (üst yazının)
                DesteBandiFotoSayfa = Application.Sum(ThisWorkbook.Worksheets(5).Range(ThisWorkbook.Worksheets(5).Cells(IlkSira, 49), ThisWorkbook.Worksheets(5).Cells(SonSira, 49)))
                TesDekFotoSayfa = Application.Sum(ThisWorkbook.Worksheets(5).Range(ThisWorkbook.Worksheets(5).Cells(IlkSira, 50), ThisWorkbook.Worksheets(5).Cells(SonSira, 50)))
                GidenPaketAdet = Application.Sum(ThisWorkbook.Worksheets(5).Range(ThisWorkbook.Worksheets(5).Cells(IlkSira, 72), ThisWorkbook.Worksheets(5).Cells(SonSira, 72)))
                Tutanak1Sayfa = Application.Sum(ThisWorkbook.Worksheets(5).Range(ThisWorkbook.Worksheets(5).Cells(IlkSira, 169), ThisWorkbook.Worksheets(5).Cells(SonSira, 169)))
                DokumSayfa = Application.Sum(ThisWorkbook.Worksheets(5).Range(ThisWorkbook.Worksheets(5).Cells(IlkSira, 170), ThisWorkbook.Worksheets(5).Cells(SonSira, 170)))
                RaporSayfa = Application.Sum(ThisWorkbook.Worksheets(5).Range(ThisWorkbook.Worksheets(5).Cells(IlkSira, 174), ThisWorkbook.Worksheets(5).Cells(SonSira, 174)))
                Tutanak2Sayfa = Application.Sum(ThisWorkbook.Worksheets(5).Range(ThisWorkbook.Worksheets(5).Cells(IlkSira, 171), ThisWorkbook.Worksheets(5).Cells(SonSira, 171)))
                FinansalBirimSayfa = Application.Sum(ThisWorkbook.Worksheets(5).Range(ThisWorkbook.Worksheets(5).Cells(IlkSira, 172), ThisWorkbook.Worksheets(5).Cells(SonSira, 172)))
                objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=5).Range.Text = DesteBandiFotoSayfa + TesDekFotoSayfa + GidenPaketAdet + Tutanak1Sayfa + DokumSayfa + RaporSayfa + Tutanak2Sayfa + FinansalBirimSayfa
                'Kilitlenme bilgisi
                'ThisWorkbook.Worksheets(5).Cells(IlkSira, 166).Value = "*"
            End If
        End If
    End If
Tip1Rapor3_2Devam1:
Next ctl '________________RAPOR3_2 BÖLÜMÜ (Bitiş, Tüm Raporların Ortak Bölümü)



'________________AKTARIMLAR 1 (Bitiş)

'Artık satırı sil
objDoc.Tables(2).Rows(Say2 + 2).Delete


'Tip1 imzaları
Set ItemBul = Worksheets(2).Range("DY6:DY1000").Find(What:=Tip1Imza1.Value, SearchDirection:=xlNext, _
                SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
If Not ItemBul Is Nothing Then
    '
Else
    GoTo Tip1Imza1DuzeltmeIslemAtla
End If
Unvan1 = Worksheets(2).Range("DZ" & ItemBul.Row)
Tip1Imza1DuzeltmeIslemAtla:

Set ItemBul = Worksheets(2).Range("DY6:DY1000").Find(What:=Tip1Imza2.Value, SearchDirection:=xlNext, _
                SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
If Not ItemBul Is Nothing Then
    '
Else
    GoTo Tip1Imza2DuzeltmeIslemAtla
End If
Unvan2 = Worksheets(2).Range("DZ" & ItemBul.Row)
Tip1Imza2DuzeltmeIslemAtla:

Set ItemBul = Worksheets(2).Range("DY6:DY1000").Find(What:=Tip1Imza3.Value, SearchDirection:=xlNext, _
                SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
If Not ItemBul Is Nothing Then
    '
Else
    GoTo Tip1Imza3DuzeltmeIslemAtla
End If
Unvan3 = Worksheets(2).Range("DZ" & ItemBul.Row)
Tip1Imza3DuzeltmeIslemAtla:


'imzalar
If Tip1Imza3.Value <> "" Then
    objDoc.Tables(4).Cell(Row:=5, Column:=1).Range.Text = Tip1Imza1.Value 'Ad Soyad1
    objDoc.Tables(4).Cell(Row:=6, Column:=1).Range.Text = Unvan1 'Unvan1
    objDoc.Tables(4).Cell(Row:=5, Column:=2).Range.Text = Tip1Imza2.Value  'Ad Soyad2
    objDoc.Tables(4).Cell(Row:=6, Column:=2).Range.Text = Unvan2 'Unvan2
    objDoc.Tables(4).Cell(Row:=5, Column:=3).Range.Text = Tip1Imza3.Value  'Ad Soyad3
    objDoc.Tables(4).Cell(Row:=6, Column:=3).Range.Text = Unvan3 'Unvan3
Else
    objDoc.Tables(4).Cell(Row:=4, Column:=1).Range.Text = ""
    objDoc.Tables(4).Cell(Row:=5, Column:=2).Range.Text = Tip1Imza1.Value 'Ad Soyad1
    objDoc.Tables(4).Cell(Row:=6, Column:=2).Range.Text = Unvan1 'Unvan1
    objDoc.Tables(4).Cell(Row:=5, Column:=3).Range.Text = Tip1Imza2.Value  'Ad Soyad2
    objDoc.Tables(4).Cell(Row:=6, Column:=3).Range.Text = Unvan2 'Unvan2
End If

'Alt bilgi ekle
objDoc.Sections(1).Footers(wdHeaderFooterPrimary).Range.Tables(1).Cell(Row:=1, Column:=1).Range.Text = ReNameTeslim

'Direkt yazdır
If TeslimTutanaklariFormuDirektYazdir = True Then
    objWord.Activate
'    For i = 1 To TeslimTutanaklariFormuSayPrt
'        objDoc.PrintOut
'    Next i
    objDoc.PrintOut Background:=False, Copies:=TeslimTutanaklariFormuSayPrt
    'objWord.Documents.Save
    objDoc.Close SaveChanges:=False
    objWord.Visible = False
End If

Son:

'Nesneleri temizle
Set fso = Nothing
Set objWord = Nothing
Set objDoc = Nothing

TeslimTutanaklariFormuDirektYazdir = False

'Columns("CK:CN").EntireColumn.Hidden = True
'Columns("CE:CF").EntireColumn.Hidden = True

Worksheets(3).Protect Password:="123" ', DrawingObjects:=False
Worksheets(4).Protect Password:="123" ', DrawingObjects:=False
Worksheets(5).Protect Password:="123" ', DrawingObjects:=False
ThisWorkbook.Protect "123"

'Worksheets(3).Activate
Application.ScreenUpdating = True
Application.DisplayAlerts = True
'Application.EnableEvents = True
    
End Sub
Sub Rapor1GidenTema()
Dim SiraBul As Range

IlkSiraGlobal = 0
Set SiraBul = ThisWorkbook.Worksheets(3).Range("E7:E100000").Find(What:=Left(ctl.List(0), InStr(ctl.List(0), "|") - 2), SearchDirection:=xlNext, _
                SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
If Not SiraBul Is Nothing Then
    'Cells(SiraBul.Row, 31).Value
    IlkSiraGlobal = SiraBul.Row
Else
    'MsgBox "Belirtilen tarihte herhangi bir üst yazı hazırlanmadığından işleminiz gerçekleştirilemiyor.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Rapor1Devam1
End If

'Giden tema
GidenTemaGlobal = ""
If ThisWorkbook.Worksheets(3).Cells(IlkSiraGlobal, 64).Value = "Provincial Directorate B" Then
    If ThisWorkbook.Worksheets(3).Cells(IlkSiraGlobal, 65).Value <> "" Then
        GidenTemaGlobal = ThisWorkbook.Worksheets(3).Cells(IlkSiraGlobal, 69).Value & " Provincial Governorship Provincial Directorate B " & ThisWorkbook.Worksheets(3).Cells(IlkSiraGlobal, 65).Value
    Else
        GidenTemaGlobal = ThisWorkbook.Worksheets(3).Cells(IlkSiraGlobal, 69).Value & " Provincial Governorship Provincial Directorate B"
    End If
ElseIf ThisWorkbook.Worksheets(3).Cells(IlkSiraGlobal, 64).Value = "District Directorate B" Then
    If ThisWorkbook.Worksheets(3).Cells(IlkSiraGlobal, 65).Value <> "" Then
        GidenTemaGlobal = ThisWorkbook.Worksheets(3).Cells(IlkSiraGlobal, 70).Value & " District Governorship District Directorate B " & ThisWorkbook.Worksheets(3).Cells(IlkSiraGlobal, 65).Value
    Else
        GidenTemaGlobal = ThisWorkbook.Worksheets(3).Cells(IlkSiraGlobal, 70).Value & " District Governorship District Directorate B"
    End If
ElseIf ThisWorkbook.Worksheets(3).Cells(IlkSiraGlobal, 64).Value = "Provincial Directorate C" Then
    If ThisWorkbook.Worksheets(3).Cells(IlkSiraGlobal, 65).Value <> "" Then
        GidenTemaGlobal = ThisWorkbook.Worksheets(3).Cells(IlkSiraGlobal, 69).Value & " Provincial Governorship Provincial Directorate C " & ThisWorkbook.Worksheets(3).Cells(IlkSiraGlobal, 65).Value
    Else
        GidenTemaGlobal = ThisWorkbook.Worksheets(3).Cells(IlkSiraGlobal, 69).Value & " Provincial Governorship Provincial Directorate C"
    End If
ElseIf ThisWorkbook.Worksheets(3).Cells(IlkSiraGlobal, 64).Value = "District Directorate C" Then
    If ThisWorkbook.Worksheets(3).Cells(IlkSiraGlobal, 65).Value <> "" Then
        GidenTemaGlobal = ThisWorkbook.Worksheets(3).Cells(IlkSiraGlobal, 70).Value & " District Governorship District Directorate C " & ThisWorkbook.Worksheets(3).Cells(IlkSiraGlobal, 65).Value
    Else
        GidenTemaGlobal = ThisWorkbook.Worksheets(3).Cells(IlkSiraGlobal, 70).Value & " District Governorship District Directorate C"
    End If
ElseIf ThisWorkbook.Worksheets(3).Cells(IlkSiraGlobal, 64).Value = "Provincial Directorate D" Then
    If ThisWorkbook.Worksheets(3).Cells(IlkSiraGlobal, 65).Value <> "" Then
        GidenTemaGlobal = ThisWorkbook.Worksheets(3).Cells(IlkSiraGlobal, 69).Value & " Provincial Governorship Provincial Directorate D " & ThisWorkbook.Worksheets(3).Cells(IlkSiraGlobal, 65).Value
    Else
        GidenTemaGlobal = ThisWorkbook.Worksheets(3).Cells(IlkSiraGlobal, 69).Value & " Provincial Governorship Provincial Directorate D"
    End If
ElseIf ThisWorkbook.Worksheets(3).Cells(IlkSiraGlobal, 64).Value = "District Directorate D" Then
    If ThisWorkbook.Worksheets(3).Cells(IlkSiraGlobal, 65).Value <> "" Then
        GidenTemaGlobal = ThisWorkbook.Worksheets(3).Cells(IlkSiraGlobal, 70).Value & " District Governorship District Directorate D " & ThisWorkbook.Worksheets(3).Cells(IlkSiraGlobal, 65).Value
    Else
        GidenTemaGlobal = ThisWorkbook.Worksheets(3).Cells(IlkSiraGlobal, 70).Value & " District Governorship District Directorate D"
    End If
ElseIf ThisWorkbook.Worksheets(3).Cells(IlkSiraGlobal, 64).Value = "Provincial Directorate E" Then
    If ThisWorkbook.Worksheets(3).Cells(IlkSiraGlobal, 65).Value <> "" Then
        GidenTemaGlobal = ThisWorkbook.Worksheets(3).Cells(IlkSiraGlobal, 69).Value & " Provincial Governorship Provincial Directorate E " & ThisWorkbook.Worksheets(3).Cells(IlkSiraGlobal, 65).Value
    Else
        GidenTemaGlobal = ThisWorkbook.Worksheets(3).Cells(IlkSiraGlobal, 69).Value & " Provincial Governorship Provincial Directorate E"
    End If
ElseIf ThisWorkbook.Worksheets(3).Cells(IlkSiraGlobal, 64).Value = "District Directorate E" Then
    If ThisWorkbook.Worksheets(3).Cells(IlkSiraGlobal, 65).Value <> "" Then
        GidenTemaGlobal = ThisWorkbook.Worksheets(3).Cells(IlkSiraGlobal, 70).Value & " District Governorship District Directorate E " & ThisWorkbook.Worksheets(3).Cells(IlkSiraGlobal, 65).Value
    Else
        GidenTemaGlobal = ThisWorkbook.Worksheets(3).Cells(IlkSiraGlobal, 70).Value & " District Governorship District Directorate E"
    End If
ElseIf InStr(ThisWorkbook.Worksheets(3).Cells(IlkSiraGlobal, 64).Value, "General Directorate") <> 0 Or InStr(ThisWorkbook.Worksheets(3).Cells(IlkSiraGlobal, 64).Value, "Regional Directorate") <> 0 Then
    If ThisWorkbook.Worksheets(3).Cells(IlkSiraGlobal, 65).Value <> "" Then
        GidenTemaGlobal = WorksheetFunction.Proper(ThisWorkbook.Worksheets(2).Cells(6, 111).Value) & " " & ThisWorkbook.Worksheets(3).Cells(IlkSiraGlobal, 64).Value & " " & ThisWorkbook.Worksheets(3).Cells(IlkSiraGlobal, 65).Value
    Else
        GidenTemaGlobal = WorksheetFunction.Proper(ThisWorkbook.Worksheets(2).Cells(6, 111).Value) & " " & ThisWorkbook.Worksheets(3).Cells(IlkSiraGlobal, 64).Value
    End If
Else
    If ThisWorkbook.Worksheets(3).Cells(IlkSiraGlobal, 65).Value <> "" Then
        If InStr(ThisWorkbook.Worksheets(3).Cells(IlkSiraGlobal, 64).Value, "X.X. ") > 0 Then
            GidenTemaGlobal = Mid(ThisWorkbook.Worksheets(3).Cells(IlkSiraGlobal, 64).Value, 6, Len(ThisWorkbook.Worksheets(3).Cells(IlkSiraGlobal, 64).Value)) & " " & ThisWorkbook.Worksheets(3).Cells(IlkSiraGlobal, 65).Value
        Else
            GidenTemaGlobal = ThisWorkbook.Worksheets(3).Cells(IlkSiraGlobal, 64).Value & " " & ThisWorkbook.Worksheets(3).Cells(IlkSiraGlobal, 65).Value
        End If
    Else
        If InStr(ThisWorkbook.Worksheets(3).Cells(IlkSiraGlobal, 64).Value, "X.X. ") > 0 Then
            GidenTemaGlobal = Mid(ThisWorkbook.Worksheets(3).Cells(IlkSiraGlobal, 64).Value, 6, Len(ThisWorkbook.Worksheets(3).Cells(IlkSiraGlobal, 64).Value))
        Else
            GidenTemaGlobal = ThisWorkbook.Worksheets(3).Cells(IlkSiraGlobal, 64).Value
        End If
    End If
End If

Rapor1Devam1:

End Sub
Sub RaporGidenTema()
Dim SiraBul As Range
Dim KoorGonderilen As Long, KoorGidenTema As Long, KoorIl As Long, KoorIlce As Long


'Rapor ve Bilgilendirme üst yazısını ayrıştır

IlkSiraGlobal = 0
Set SiraBul = ThisWorkbook.Worksheets(4).Range("E7:E100000").Find(What:=Left(ctl.List(0), InStr(ctl.List(0), "|") - 2), SearchDirection:=xlNext, _
                SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
If Not SiraBul Is Nothing Then
    'Cells(SiraBul.Row, 31).Value
    IlkSiraGlobal = SiraBul.Row
Else
    'MsgBox "Belirtilen tarihte herhangi bir üst yazı hazırlanmadığından işleminiz gerçekleştirilemiyor.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo RaporDevam1
End If

'Giden tema
GidenTemaGlobal = ""


'Değişken Koordinat Düzenleyicisi
If ThisWorkbook.Worksheets(4).Cells(IlkSiraGlobal, 174).Value = "Yes" Then 'Bilgilendirme (Rapor2_2 evet)
    'Bilgilendirme
    KoorGidenTema = 199
    KoorGonderilen = 200
    KoorIl = 203
    KoorIlce = 204
Else 'Rapor (Rapor2_2 hayır)
    'Rapor
    KoorGidenTema = 72
    KoorGonderilen = 73
    KoorIl = 77
    KoorIlce = 78
End If

'Giden tema
If ThisWorkbook.Worksheets(4).Cells(IlkSiraGlobal, KoorGidenTema).Value = "Provincial Directorate B" Then
    If ThisWorkbook.Worksheets(4).Cells(IlkSiraGlobal, KoorGonderilen).Value <> "" Then
        GidenTemaGlobal = ThisWorkbook.Worksheets(4).Cells(IlkSiraGlobal, KoorIl).Value & " Provincial Governorship Provincial Directorate B " & ThisWorkbook.Worksheets(4).Cells(IlkSiraGlobal, KoorGonderilen).Value
    Else
        GidenTemaGlobal = ThisWorkbook.Worksheets(4).Cells(IlkSiraGlobal, KoorIl).Value & " Provincial Governorship Provincial Directorate B"
    End If
ElseIf ThisWorkbook.Worksheets(4).Cells(IlkSiraGlobal, KoorGidenTema).Value = "District Directorate B" Then
    If ThisWorkbook.Worksheets(4).Cells(IlkSiraGlobal, KoorGonderilen).Value <> "" Then
        GidenTemaGlobal = ThisWorkbook.Worksheets(4).Cells(IlkSiraGlobal, KoorIlce).Value & " District Governorship District Directorate B " & ThisWorkbook.Worksheets(4).Cells(IlkSiraGlobal, KoorGonderilen).Value
    Else
        GidenTemaGlobal = ThisWorkbook.Worksheets(4).Cells(IlkSiraGlobal, KoorIlce).Value & " District Governorship District Directorate B"
    End If
ElseIf ThisWorkbook.Worksheets(4).Cells(IlkSiraGlobal, KoorGidenTema).Value = "Provincial Directorate C" Then
    If ThisWorkbook.Worksheets(4).Cells(IlkSiraGlobal, KoorGonderilen).Value <> "" Then
        GidenTemaGlobal = ThisWorkbook.Worksheets(4).Cells(IlkSiraGlobal, KoorIl).Value & " Provincial Governorship Provincial Directorate C " & ThisWorkbook.Worksheets(4).Cells(IlkSiraGlobal, KoorGonderilen).Value
    Else
        GidenTemaGlobal = ThisWorkbook.Worksheets(4).Cells(IlkSiraGlobal, KoorIl).Value & " Provincial Governorship Provincial Directorate C"
    End If
ElseIf ThisWorkbook.Worksheets(4).Cells(IlkSiraGlobal, KoorGidenTema).Value = "District Directorate C" Then
    If ThisWorkbook.Worksheets(4).Cells(IlkSiraGlobal, KoorGonderilen).Value <> "" Then
        GidenTemaGlobal = ThisWorkbook.Worksheets(4).Cells(IlkSiraGlobal, KoorIlce).Value & " District Governorship District Directorate C " & ThisWorkbook.Worksheets(4).Cells(IlkSiraGlobal, KoorGonderilen).Value
    Else
        GidenTemaGlobal = ThisWorkbook.Worksheets(4).Cells(IlkSiraGlobal, KoorIlce).Value & " District Governorship District Directorate C"
    End If
ElseIf ThisWorkbook.Worksheets(4).Cells(IlkSiraGlobal, KoorGidenTema).Value = "Provincial Directorate D" Then
    If ThisWorkbook.Worksheets(4).Cells(IlkSiraGlobal, KoorGonderilen).Value <> "" Then
        GidenTemaGlobal = ThisWorkbook.Worksheets(4).Cells(IlkSiraGlobal, KoorIl).Value & " Provincial Governorship Provincial Directorate D " & ThisWorkbook.Worksheets(4).Cells(IlkSiraGlobal, KoorGonderilen).Value
    Else
        GidenTemaGlobal = ThisWorkbook.Worksheets(4).Cells(IlkSiraGlobal, KoorIl).Value & " Provincial Governorship Provincial Directorate D"
    End If
ElseIf ThisWorkbook.Worksheets(4).Cells(IlkSiraGlobal, KoorGidenTema).Value = "District Directorate D" Then
    If ThisWorkbook.Worksheets(4).Cells(IlkSiraGlobal, KoorGonderilen).Value <> "" Then
        GidenTemaGlobal = ThisWorkbook.Worksheets(4).Cells(IlkSiraGlobal, KoorIlce).Value & " District Governorship District Directorate D " & ThisWorkbook.Worksheets(4).Cells(IlkSiraGlobal, KoorGonderilen).Value
    Else
        GidenTemaGlobal = ThisWorkbook.Worksheets(4).Cells(IlkSiraGlobal, KoorIlce).Value & " District Governorship District Directorate D"
    End If
ElseIf ThisWorkbook.Worksheets(4).Cells(IlkSiraGlobal, KoorGidenTema).Value = "Provincial Directorate E" Then
    If ThisWorkbook.Worksheets(4).Cells(IlkSiraGlobal, KoorGonderilen).Value <> "" Then
        GidenTemaGlobal = ThisWorkbook.Worksheets(4).Cells(IlkSiraGlobal, KoorIl).Value & " Provincial Governorship Provincial Directorate E " & ThisWorkbook.Worksheets(4).Cells(IlkSiraGlobal, KoorGonderilen).Value
    Else
        GidenTemaGlobal = ThisWorkbook.Worksheets(4).Cells(IlkSiraGlobal, KoorIl).Value & " Provincial Governorship Provincial Directorate E"
    End If
ElseIf ThisWorkbook.Worksheets(4).Cells(IlkSiraGlobal, KoorGidenTema).Value = "District Directorate E" Then
    If ThisWorkbook.Worksheets(4).Cells(IlkSiraGlobal, KoorGonderilen).Value <> "" Then
        GidenTemaGlobal = ThisWorkbook.Worksheets(4).Cells(IlkSiraGlobal, KoorIlce).Value & " District Governorship District Directorate E " & ThisWorkbook.Worksheets(4).Cells(IlkSiraGlobal, KoorGonderilen).Value
    Else
        GidenTemaGlobal = ThisWorkbook.Worksheets(4).Cells(IlkSiraGlobal, KoorIlce).Value & " District Governorship District Directorate E"
    End If
ElseIf InStr(ThisWorkbook.Worksheets(4).Cells(IlkSiraGlobal, KoorGidenTema).Value, "General Directorate") <> 0 Or InStr(ThisWorkbook.Worksheets(4).Cells(IlkSiraGlobal, KoorGidenTema).Value, "Regional Directorate") <> 0 Then
    If ThisWorkbook.Worksheets(4).Cells(IlkSiraGlobal, KoorGonderilen).Value <> "" Then
        GidenTemaGlobal = WorksheetFunction.Proper(ThisWorkbook.Worksheets(2).Cells(6, 111).Value) & " " & ThisWorkbook.Worksheets(4).Cells(IlkSiraGlobal, KoorGidenTema).Value & " " & ThisWorkbook.Worksheets(4).Cells(IlkSiraGlobal, KoorGonderilen).Value
    Else
        GidenTemaGlobal = WorksheetFunction.Proper(ThisWorkbook.Worksheets(2).Cells(6, 111).Value) & " " & ThisWorkbook.Worksheets(4).Cells(IlkSiraGlobal, KoorGidenTema).Value
    End If
Else
    If ThisWorkbook.Worksheets(4).Cells(IlkSiraGlobal, KoorGonderilen).Value <> "" Then
        If InStr(ThisWorkbook.Worksheets(4).Cells(IlkSiraGlobal, KoorGidenTema).Value, "X.X. ") > 0 Then
            GidenTemaGlobal = Mid(ThisWorkbook.Worksheets(4).Cells(IlkSiraGlobal, KoorGidenTema).Value, 6, Len(ThisWorkbook.Worksheets(4).Cells(IlkSiraGlobal, KoorGidenTema).Value)) & " " & ThisWorkbook.Worksheets(4).Cells(IlkSiraGlobal, KoorGonderilen).Value
        Else
            GidenTemaGlobal = ThisWorkbook.Worksheets(4).Cells(IlkSiraGlobal, KoorGidenTema).Value & " " & ThisWorkbook.Worksheets(4).Cells(IlkSiraGlobal, KoorGonderilen).Value
        End If
    Else
        If InStr(ThisWorkbook.Worksheets(4).Cells(IlkSiraGlobal, KoorGidenTema).Value, "X.X. ") > 0 Then
            GidenTemaGlobal = Mid(ThisWorkbook.Worksheets(4).Cells(IlkSiraGlobal, KoorGidenTema).Value, 6, Len(ThisWorkbook.Worksheets(4).Cells(IlkSiraGlobal, KoorGidenTema).Value))
        Else
            GidenTemaGlobal = ThisWorkbook.Worksheets(4).Cells(IlkSiraGlobal, KoorGidenTema).Value
        End If
    End If
End If

RaporDevam1:

End Sub

Sub Rapor2_2GidenTema()
Dim SiraBul As Range
Dim KoorGonderilen As Long, KoorGidenTema As Long, KoorIl As Long, KoorIlce As Long

'Sonuç üst yazısı için tema

IlkSiraGlobal = 0
Set SiraBul = ThisWorkbook.Worksheets(4).Range("E7:E100000").Find(What:=Left(ctl.List(0), InStr(ctl.List(0), "|") - 2), SearchDirection:=xlNext, _
                SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
If Not SiraBul Is Nothing Then
    'Cells(SiraBul.Row, 31).Value
    IlkSiraGlobal = SiraBul.Row
Else
    'MsgBox "Belirtilen tarihte herhangi bir üst yazı hazırlanmadığından işleminiz gerçekleştirilemiyor.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Rapor2_2Devam1
End If
            
'Giden tema
GidenTemaGlobal = ""

'Değişken Koordinat Düzenleyicisi
'Bilgilendirme
KoorGidenTema = 199
KoorGonderilen = 200
KoorIl = 203
KoorIlce = 204
            
'Giden tema
If ThisWorkbook.Worksheets(4).Cells(IlkSiraGlobal, KoorGidenTema).Value = "Provincial Directorate B" Then
    If ThisWorkbook.Worksheets(4).Cells(IlkSiraGlobal, KoorGonderilen).Value <> "" Then
        GidenTemaGlobal = ThisWorkbook.Worksheets(4).Cells(IlkSiraGlobal, KoorIl).Value & " Provincial Governorship Provincial Directorate B " & ThisWorkbook.Worksheets(4).Cells(IlkSiraGlobal, KoorGonderilen).Value
    Else
        GidenTemaGlobal = ThisWorkbook.Worksheets(4).Cells(IlkSiraGlobal, KoorIl).Value & " Provincial Governorship Provincial Directorate B"
    End If
ElseIf ThisWorkbook.Worksheets(4).Cells(IlkSiraGlobal, KoorGidenTema).Value = "District Directorate B" Then
    If ThisWorkbook.Worksheets(4).Cells(IlkSiraGlobal, KoorGonderilen).Value <> "" Then
        GidenTemaGlobal = ThisWorkbook.Worksheets(4).Cells(IlkSiraGlobal, KoorIlce).Value & " District Governorship District Directorate B " & ThisWorkbook.Worksheets(4).Cells(IlkSiraGlobal, KoorGonderilen).Value
    Else
        GidenTemaGlobal = ThisWorkbook.Worksheets(4).Cells(IlkSiraGlobal, KoorIlce).Value & " District Governorship District Directorate B"
    End If
ElseIf ThisWorkbook.Worksheets(4).Cells(IlkSiraGlobal, KoorGidenTema).Value = "Provincial Directorate C" Then
    If ThisWorkbook.Worksheets(4).Cells(IlkSiraGlobal, KoorGonderilen).Value <> "" Then
        GidenTemaGlobal = ThisWorkbook.Worksheets(4).Cells(IlkSiraGlobal, KoorIl).Value & " Provincial Governorship Provincial Directorate C " & ThisWorkbook.Worksheets(4).Cells(IlkSiraGlobal, KoorGonderilen).Value
    Else
        GidenTemaGlobal = ThisWorkbook.Worksheets(4).Cells(IlkSiraGlobal, KoorIl).Value & " Provincial Governorship Provincial Directorate C"
    End If
ElseIf ThisWorkbook.Worksheets(4).Cells(IlkSiraGlobal, KoorGidenTema).Value = "District Directorate C" Then
    If ThisWorkbook.Worksheets(4).Cells(IlkSiraGlobal, KoorGonderilen).Value <> "" Then
        GidenTemaGlobal = ThisWorkbook.Worksheets(4).Cells(IlkSiraGlobal, KoorIlce).Value & " District Governorship District Directorate C " & ThisWorkbook.Worksheets(4).Cells(IlkSiraGlobal, KoorGonderilen).Value
    Else
        GidenTemaGlobal = ThisWorkbook.Worksheets(4).Cells(IlkSiraGlobal, KoorIlce).Value & " District Governorship District Directorate C"
    End If
ElseIf ThisWorkbook.Worksheets(4).Cells(IlkSiraGlobal, KoorGidenTema).Value = "Provincial Directorate D" Then
    If ThisWorkbook.Worksheets(4).Cells(IlkSiraGlobal, KoorGonderilen).Value <> "" Then
        GidenTemaGlobal = ThisWorkbook.Worksheets(4).Cells(IlkSiraGlobal, KoorIl).Value & " Provincial Governorship Provincial Directorate D " & ThisWorkbook.Worksheets(4).Cells(IlkSiraGlobal, KoorGonderilen).Value
    Else
        GidenTemaGlobal = ThisWorkbook.Worksheets(4).Cells(IlkSiraGlobal, KoorIl).Value & " Provincial Governorship Provincial Directorate D"
    End If
ElseIf ThisWorkbook.Worksheets(4).Cells(IlkSiraGlobal, KoorGidenTema).Value = "District Directorate D" Then
    If ThisWorkbook.Worksheets(4).Cells(IlkSiraGlobal, KoorGonderilen).Value <> "" Then
        GidenTemaGlobal = ThisWorkbook.Worksheets(4).Cells(IlkSiraGlobal, KoorIlce).Value & " District Governorship District Directorate D " & ThisWorkbook.Worksheets(4).Cells(IlkSiraGlobal, KoorGonderilen).Value
    Else
        GidenTemaGlobal = ThisWorkbook.Worksheets(4).Cells(IlkSiraGlobal, KoorIlce).Value & " District Governorship District Directorate D"
    End If
ElseIf ThisWorkbook.Worksheets(4).Cells(IlkSiraGlobal, KoorGidenTema).Value = "Provincial Directorate E" Then
    If ThisWorkbook.Worksheets(4).Cells(IlkSiraGlobal, KoorGonderilen).Value <> "" Then
        GidenTemaGlobal = ThisWorkbook.Worksheets(4).Cells(IlkSiraGlobal, KoorIl).Value & " Provincial Governorship Provincial Directorate E " & ThisWorkbook.Worksheets(4).Cells(IlkSiraGlobal, KoorGonderilen).Value
    Else
        GidenTemaGlobal = ThisWorkbook.Worksheets(4).Cells(IlkSiraGlobal, KoorIl).Value & " Provincial Governorship Provincial Directorate E"
    End If
ElseIf ThisWorkbook.Worksheets(4).Cells(IlkSiraGlobal, KoorGidenTema).Value = "District Directorate E" Then
    If ThisWorkbook.Worksheets(4).Cells(IlkSiraGlobal, KoorGonderilen).Value <> "" Then
        GidenTemaGlobal = ThisWorkbook.Worksheets(4).Cells(IlkSiraGlobal, KoorIlce).Value & " District Governorship District Directorate E " & ThisWorkbook.Worksheets(4).Cells(IlkSiraGlobal, KoorGonderilen).Value
    Else
        GidenTemaGlobal = ThisWorkbook.Worksheets(4).Cells(IlkSiraGlobal, KoorIlce).Value & " District Governorship District Directorate E"
    End If
ElseIf InStr(ThisWorkbook.Worksheets(4).Cells(IlkSiraGlobal, KoorGidenTema).Value, "General Directorate") <> 0 Or InStr(ThisWorkbook.Worksheets(4).Cells(IlkSiraGlobal, KoorGidenTema).Value, "Regional Directorate") <> 0 Then
    If ThisWorkbook.Worksheets(4).Cells(IlkSiraGlobal, KoorGonderilen).Value <> "" Then
        GidenTemaGlobal = WorksheetFunction.Proper(ThisWorkbook.Worksheets(2).Cells(6, 111).Value) & " " & ThisWorkbook.Worksheets(4).Cells(IlkSiraGlobal, KoorGidenTema).Value & " " & ThisWorkbook.Worksheets(4).Cells(IlkSiraGlobal, KoorGonderilen).Value
    Else
        GidenTemaGlobal = WorksheetFunction.Proper(ThisWorkbook.Worksheets(2).Cells(6, 111).Value) & " " & ThisWorkbook.Worksheets(4).Cells(IlkSiraGlobal, KoorGidenTema).Value
    End If
Else
    If ThisWorkbook.Worksheets(4).Cells(IlkSiraGlobal, KoorGonderilen).Value <> "" Then
        If InStr(ThisWorkbook.Worksheets(4).Cells(IlkSiraGlobal, KoorGidenTema).Value, "X.X. ") > 0 Then
            GidenTemaGlobal = Mid(ThisWorkbook.Worksheets(4).Cells(IlkSiraGlobal, KoorGidenTema).Value, 6, Len(ThisWorkbook.Worksheets(4).Cells(IlkSiraGlobal, KoorGidenTema).Value)) & " " & ThisWorkbook.Worksheets(4).Cells(IlkSiraGlobal, KoorGonderilen).Value
        Else
            GidenTemaGlobal = ThisWorkbook.Worksheets(4).Cells(IlkSiraGlobal, KoorGidenTema).Value & " " & ThisWorkbook.Worksheets(4).Cells(IlkSiraGlobal, KoorGonderilen).Value
        End If
    Else
        If InStr(ThisWorkbook.Worksheets(4).Cells(IlkSiraGlobal, KoorGidenTema).Value, "X.X. ") > 0 Then
            GidenTemaGlobal = Mid(ThisWorkbook.Worksheets(4).Cells(IlkSiraGlobal, KoorGidenTema).Value, 6, Len(ThisWorkbook.Worksheets(4).Cells(IlkSiraGlobal, KoorGidenTema).Value))
        Else
            GidenTemaGlobal = ThisWorkbook.Worksheets(4).Cells(IlkSiraGlobal, KoorGidenTema).Value
        End If
    End If
End If

Rapor2_2Devam1:

End Sub

Sub Rapor3_1GidenTema()
Dim SiraBul As Range

IlkSiraGlobal = 0
Set SiraBul = ThisWorkbook.Worksheets(5).Range("E7:E100000").Find(What:=Left(ctl.List(0), InStr(ctl.List(0), "|") - 2), SearchDirection:=xlNext, _
                SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
If Not SiraBul Is Nothing Then
    'Cells(SiraBul.Row, 31).Value
    IlkSiraGlobal = SiraBul.Row
Else
    'MsgBox "Belirtilen tarihte herhangi bir üst yazı hazırlanmadığından işleminiz gerçekleştirilemiyor.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Rapor3_1Devam1
End If

'Giden tema
GidenTemaGlobal = ""
If ThisWorkbook.Worksheets(5).Cells(IlkSiraGlobal, 102).Value = "Provincial Directorate B" Then
    If ThisWorkbook.Worksheets(5).Cells(IlkSiraGlobal, 103).Value <> "" Then
        GidenTemaGlobal = ThisWorkbook.Worksheets(5).Cells(IlkSiraGlobal, 91).Value & " Provincial Governorship Provincial Directorate B " & ThisWorkbook.Worksheets(5).Cells(IlkSiraGlobal, 103).Value
    Else
        GidenTemaGlobal = ThisWorkbook.Worksheets(5).Cells(IlkSiraGlobal, 91).Value & " Provincial Governorship Provincial Directorate B"
    End If
ElseIf ThisWorkbook.Worksheets(5).Cells(IlkSiraGlobal, 102).Value = "District Directorate B" Then
    If ThisWorkbook.Worksheets(5).Cells(IlkSiraGlobal, 103).Value <> "" Then
        GidenTemaGlobal = ThisWorkbook.Worksheets(5).Cells(IlkSiraGlobal, 92).Value & " District Governorship District Directorate B " & ThisWorkbook.Worksheets(5).Cells(IlkSiraGlobal, 103).Value
    Else
        GidenTemaGlobal = ThisWorkbook.Worksheets(5).Cells(IlkSiraGlobal, 92).Value & " District Governorship District Directorate B"
    End If
ElseIf ThisWorkbook.Worksheets(5).Cells(IlkSiraGlobal, 102).Value = "Provincial Directorate C" Then
    If ThisWorkbook.Worksheets(5).Cells(IlkSiraGlobal, 103).Value <> "" Then
        GidenTemaGlobal = ThisWorkbook.Worksheets(5).Cells(IlkSiraGlobal, 91).Value & " Provincial Governorship Provincial Directorate C " & ThisWorkbook.Worksheets(5).Cells(IlkSiraGlobal, 103).Value
    Else
        GidenTemaGlobal = ThisWorkbook.Worksheets(5).Cells(IlkSiraGlobal, 91).Value & " Provincial Governorship Provincial Directorate C"
    End If
ElseIf ThisWorkbook.Worksheets(5).Cells(IlkSiraGlobal, 102).Value = "District Directorate C" Then
    If ThisWorkbook.Worksheets(5).Cells(IlkSiraGlobal, 103).Value <> "" Then
        GidenTemaGlobal = ThisWorkbook.Worksheets(5).Cells(IlkSiraGlobal, 92).Value & " District Governorship District Directorate C " & ThisWorkbook.Worksheets(5).Cells(IlkSiraGlobal, 103).Value
    Else
        GidenTemaGlobal = ThisWorkbook.Worksheets(5).Cells(IlkSiraGlobal, 92).Value & " District Governorship District Directorate C"
    End If
ElseIf ThisWorkbook.Worksheets(5).Cells(IlkSiraGlobal, 102).Value = "Provincial Directorate D" Then
    If ThisWorkbook.Worksheets(5).Cells(IlkSiraGlobal, 103).Value <> "" Then
        GidenTemaGlobal = ThisWorkbook.Worksheets(5).Cells(IlkSiraGlobal, 91).Value & " Provincial Governorship Provincial Directorate D " & ThisWorkbook.Worksheets(5).Cells(IlkSiraGlobal, 103).Value
    Else
        GidenTemaGlobal = ThisWorkbook.Worksheets(5).Cells(IlkSiraGlobal, 91).Value & " Provincial Governorship Provincial Directorate D"
    End If
ElseIf ThisWorkbook.Worksheets(5).Cells(IlkSiraGlobal, 102).Value = "District Directorate D" Then
    If ThisWorkbook.Worksheets(5).Cells(IlkSiraGlobal, 103).Value <> "" Then
        GidenTemaGlobal = ThisWorkbook.Worksheets(5).Cells(IlkSiraGlobal, 92).Value & " District Governorship District Directorate D " & ThisWorkbook.Worksheets(5).Cells(IlkSiraGlobal, 103).Value
    Else
        GidenTemaGlobal = ThisWorkbook.Worksheets(5).Cells(IlkSiraGlobal, 92).Value & " District Governorship District Directorate D"
    End If
ElseIf ThisWorkbook.Worksheets(5).Cells(IlkSiraGlobal, 102).Value = "Provincial Directorate E" Then
    If ThisWorkbook.Worksheets(5).Cells(IlkSiraGlobal, 103).Value <> "" Then
        GidenTemaGlobal = ThisWorkbook.Worksheets(5).Cells(IlkSiraGlobal, 91).Value & " Provincial Governorship Provincial Directorate E " & ThisWorkbook.Worksheets(5).Cells(IlkSiraGlobal, 103).Value
    Else
        GidenTemaGlobal = ThisWorkbook.Worksheets(5).Cells(IlkSiraGlobal, 91).Value & " Provincial Governorship Provincial Directorate E"
    End If
ElseIf ThisWorkbook.Worksheets(5).Cells(IlkSiraGlobal, 102).Value = "District Directorate E" Then
    If ThisWorkbook.Worksheets(5).Cells(IlkSiraGlobal, 103).Value <> "" Then
        GidenTemaGlobal = ThisWorkbook.Worksheets(5).Cells(IlkSiraGlobal, 92).Value & " District Governorship District Directorate E " & ThisWorkbook.Worksheets(5).Cells(IlkSiraGlobal, 103).Value
    Else
        GidenTemaGlobal = ThisWorkbook.Worksheets(5).Cells(IlkSiraGlobal, 92).Value & " District Governorship District Directorate E"
    End If
ElseIf InStr(ThisWorkbook.Worksheets(5).Cells(IlkSiraGlobal, 102).Value, "General Directorate") <> 0 Or InStr(ThisWorkbook.Worksheets(5).Cells(IlkSiraGlobal, 102).Value, "Regional Directorate") <> 0 Then
    If ThisWorkbook.Worksheets(5).Cells(IlkSiraGlobal, 103).Value <> "" Then
        GidenTemaGlobal = WorksheetFunction.Proper(ThisWorkbook.Worksheets(2).Cells(6, 111).Value) & " " & ThisWorkbook.Worksheets(5).Cells(IlkSiraGlobal, 102).Value & " " & ThisWorkbook.Worksheets(5).Cells(IlkSiraGlobal, 103).Value
    Else
        GidenTemaGlobal = WorksheetFunction.Proper(ThisWorkbook.Worksheets(2).Cells(6, 111).Value) & " " & ThisWorkbook.Worksheets(5).Cells(IlkSiraGlobal, 102).Value
    End If
Else
    If ThisWorkbook.Worksheets(5).Cells(IlkSiraGlobal, 103).Value <> "" Then
        If InStr(ThisWorkbook.Worksheets(5).Cells(IlkSiraGlobal, 102).Value, "X.X. ") > 0 Then
            GidenTemaGlobal = Mid(ThisWorkbook.Worksheets(5).Cells(IlkSiraGlobal, 102).Value, 6, Len(ThisWorkbook.Worksheets(5).Cells(IlkSiraGlobal, 102).Value)) & " " & ThisWorkbook.Worksheets(5).Cells(IlkSiraGlobal, 103).Value
        Else
            GidenTemaGlobal = ThisWorkbook.Worksheets(5).Cells(IlkSiraGlobal, 102).Value & " " & ThisWorkbook.Worksheets(5).Cells(IlkSiraGlobal, 103).Value
        End If
    Else
        If InStr(ThisWorkbook.Worksheets(5).Cells(IlkSiraGlobal, 102).Value, "X.X. ") > 0 Then
            GidenTemaGlobal = Mid(ThisWorkbook.Worksheets(5).Cells(IlkSiraGlobal, 102).Value, 6, Len(ThisWorkbook.Worksheets(5).Cells(IlkSiraGlobal, 102).Value))
        Else
            GidenTemaGlobal = ThisWorkbook.Worksheets(5).Cells(IlkSiraGlobal, 102).Value
        End If
    End If
End If

Rapor3_1Devam1:

End Sub

Sub Rapor3_2GidenTema()
Dim SiraBul As Range

IlkSiraGlobal = 0
Set SiraBul = ThisWorkbook.Worksheets(5).Range("E7:E100000").Find(What:=Left(ctl.List(0), InStr(ctl.List(0), "|") - 2), SearchDirection:=xlNext, _
                SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
If Not SiraBul Is Nothing Then
    'Cells(SiraBul.Row, 31).Value
    IlkSiraGlobal = SiraBul.Row
Else
    'MsgBox "Belirtilen tarihte herhangi bir üst yazı hazırlanmadığından işleminiz gerçekleştirilemiyor.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Rapor3_2Devam1
End If

'Giden tema
GidenTemaGlobal = ""
If ThisWorkbook.Worksheets(5).Cells(IlkSiraGlobal, 47).Value = "Provincial Directorate B" Then
    If ThisWorkbook.Worksheets(5).Cells(IlkSiraGlobal, 48).Value <> "" Then
        GidenTemaGlobal = ThisWorkbook.Worksheets(5).Cells(IlkSiraGlobal, 19).Value & " Provincial Governorship Provincial Directorate B " & ThisWorkbook.Worksheets(5).Cells(IlkSiraGlobal, 48).Value
    Else
        GidenTemaGlobal = ThisWorkbook.Worksheets(5).Cells(IlkSiraGlobal, 19).Value & " Provincial Governorship Provincial Directorate B"
    End If
ElseIf ThisWorkbook.Worksheets(5).Cells(IlkSiraGlobal, 47).Value = "District Directorate B" Then
    If ThisWorkbook.Worksheets(5).Cells(IlkSiraGlobal, 48).Value <> "" Then
        GidenTemaGlobal = ThisWorkbook.Worksheets(5).Cells(IlkSiraGlobal, 20).Value & " District Governorship District Directorate B " & ThisWorkbook.Worksheets(5).Cells(IlkSiraGlobal, 48).Value
    Else
        GidenTemaGlobal = ThisWorkbook.Worksheets(5).Cells(IlkSiraGlobal, 20).Value & " District Governorship District Directorate B"
    End If
ElseIf ThisWorkbook.Worksheets(5).Cells(IlkSiraGlobal, 47).Value = "Provincial Directorate C" Then
    If ThisWorkbook.Worksheets(5).Cells(IlkSiraGlobal, 48).Value <> "" Then
        GidenTemaGlobal = ThisWorkbook.Worksheets(5).Cells(IlkSiraGlobal, 19).Value & " Provincial Governorship Provincial Directorate C " & ThisWorkbook.Worksheets(5).Cells(IlkSiraGlobal, 48).Value
    Else
        GidenTemaGlobal = ThisWorkbook.Worksheets(5).Cells(IlkSiraGlobal, 19).Value & " Provincial Governorship Provincial Directorate C"
    End If
ElseIf ThisWorkbook.Worksheets(5).Cells(IlkSiraGlobal, 47).Value = "District Directorate C" Then
    If ThisWorkbook.Worksheets(5).Cells(IlkSiraGlobal, 48).Value <> "" Then
        GidenTemaGlobal = ThisWorkbook.Worksheets(5).Cells(IlkSiraGlobal, 20).Value & " District Governorship District Directorate C " & ThisWorkbook.Worksheets(5).Cells(IlkSiraGlobal, 48).Value
    Else
        GidenTemaGlobal = ThisWorkbook.Worksheets(5).Cells(IlkSiraGlobal, 20).Value & " District Governorship District Directorate C"
    End If
ElseIf ThisWorkbook.Worksheets(5).Cells(IlkSiraGlobal, 47).Value = "Provincial Directorate D" Then
    If ThisWorkbook.Worksheets(5).Cells(IlkSiraGlobal, 48).Value <> "" Then
        GidenTemaGlobal = ThisWorkbook.Worksheets(5).Cells(IlkSiraGlobal, 19).Value & " Provincial Governorship Provincial Directorate D " & ThisWorkbook.Worksheets(5).Cells(IlkSiraGlobal, 48).Value
    Else
        GidenTemaGlobal = ThisWorkbook.Worksheets(5).Cells(IlkSiraGlobal, 19).Value & " Provincial Governorship Provincial Directorate D"
    End If
ElseIf ThisWorkbook.Worksheets(5).Cells(IlkSiraGlobal, 47).Value = "District Directorate D" Then
    If ThisWorkbook.Worksheets(5).Cells(IlkSiraGlobal, 48).Value <> "" Then
        GidenTemaGlobal = ThisWorkbook.Worksheets(5).Cells(IlkSiraGlobal, 20).Value & " District Governorship District Directorate D " & ThisWorkbook.Worksheets(5).Cells(IlkSiraGlobal, 48).Value
    Else
        GidenTemaGlobal = ThisWorkbook.Worksheets(5).Cells(IlkSiraGlobal, 20).Value & " District Governorship District Directorate D"
    End If
ElseIf ThisWorkbook.Worksheets(5).Cells(IlkSiraGlobal, 47).Value = "Provincial Directorate E" Then
    If ThisWorkbook.Worksheets(5).Cells(IlkSiraGlobal, 48).Value <> "" Then
        GidenTemaGlobal = ThisWorkbook.Worksheets(5).Cells(IlkSiraGlobal, 19).Value & " Provincial Governorship Provincial Directorate E " & ThisWorkbook.Worksheets(5).Cells(IlkSiraGlobal, 48).Value
    Else
        GidenTemaGlobal = ThisWorkbook.Worksheets(5).Cells(IlkSiraGlobal, 19).Value & " Provincial Governorship Provincial Directorate E"
    End If
ElseIf ThisWorkbook.Worksheets(5).Cells(IlkSiraGlobal, 47).Value = "District Directorate E" Then
    If ThisWorkbook.Worksheets(5).Cells(IlkSiraGlobal, 48).Value <> "" Then
        GidenTemaGlobal = ThisWorkbook.Worksheets(5).Cells(IlkSiraGlobal, 20).Value & " District Governorship District Directorate E " & ThisWorkbook.Worksheets(5).Cells(IlkSiraGlobal, 48).Value
    Else
        GidenTemaGlobal = ThisWorkbook.Worksheets(5).Cells(IlkSiraGlobal, 20).Value & " District Governorship District Directorate E"
    End If
ElseIf InStr(ThisWorkbook.Worksheets(5).Cells(IlkSiraGlobal, 47).Value, "General Directorate") <> 0 Or InStr(ThisWorkbook.Worksheets(5).Cells(IlkSiraGlobal, 47).Value, "Regional Directorate") <> 0 Then
    If ThisWorkbook.Worksheets(5).Cells(IlkSiraGlobal, 48).Value <> "" Then
        GidenTemaGlobal = WorksheetFunction.Proper(ThisWorkbook.Worksheets(2).Cells(6, 111).Value) & " " & ThisWorkbook.Worksheets(5).Cells(IlkSiraGlobal, 47).Value & " " & ThisWorkbook.Worksheets(5).Cells(IlkSiraGlobal, 48).Value
    Else
        GidenTemaGlobal = WorksheetFunction.Proper(ThisWorkbook.Worksheets(2).Cells(6, 111).Value) & " " & ThisWorkbook.Worksheets(5).Cells(IlkSiraGlobal, 47).Value
        
    End If
Else
    If ThisWorkbook.Worksheets(5).Cells(IlkSiraGlobal, 48).Value <> "" Then
        If InStr(ThisWorkbook.Worksheets(5).Cells(IlkSiraGlobal, 47).Value, "X.X. ") > 0 Then
            GidenTemaGlobal = Mid(ThisWorkbook.Worksheets(5).Cells(IlkSiraGlobal, 47).Value, 6, Len(ThisWorkbook.Worksheets(5).Cells(IlkSiraGlobal, 47).Value)) & " " & ThisWorkbook.Worksheets(5).Cells(IlkSiraGlobal, 48).Value
        Else
            GidenTemaGlobal = ThisWorkbook.Worksheets(5).Cells(IlkSiraGlobal, 47).Value & " " & ThisWorkbook.Worksheets(5).Cells(IlkSiraGlobal, 48).Value
        End If
    Else
        If InStr(ThisWorkbook.Worksheets(5).Cells(IlkSiraGlobal, 47).Value, "X.X. ") > 0 Then
            GidenTemaGlobal = Mid(ThisWorkbook.Worksheets(5).Cells(IlkSiraGlobal, 47).Value, 6, Len(ThisWorkbook.Worksheets(5).Cells(IlkSiraGlobal, 47).Value))
        Else
            GidenTemaGlobal = ThisWorkbook.Worksheets(5).Cells(IlkSiraGlobal, 47).Value
        End If
    End If
End If
        
Rapor3_2Devam1:

End Sub
Private Sub Tutanak_Click()
Dim AutoPath As String, DestOperasyon As String, SourceTeslim As String, DestOpUserFolderName As String, DestOpUserFolder As String
Dim ReNameTeslim As String, OpenKontrolName As String, ContSay As Long, KontrolFile As String
Dim fso As Object, objWord As Object, objDoc As Object, Birimx As String, OpenControl As String
Dim Cont As Long, SiraBul As Range, i As Long, j As Long, Say As Long
'Dim Ctl As MSForms.Control
Dim Ctlx As MSForms.Control, Say2 As Integer, GidenTema As String, IlkSira As Long
Dim MyRange As Object, AdetSay As Integer, VeriKontrol As Integer
Dim EkFormu As String, SourceTip5 As String, ReNameTip5 As String, EkFormuTip5 As String
Dim a() As Variant, b As Variant
Dim GidenPaketAdet As Integer
Dim Ek1, Ek2, Ek3, Ek4, Ek5 As String

ThisWorkbook.Activate

If Tip1Option.Value = True Then
    Call Tip1Proseduru
    GoTo Tip1Son
End If

'Worksheets(3).Activate
Application.ScreenUpdating = False
Application.DisplayAlerts = False
'Application.EnableEvents = False

ThisWorkbook.Unprotect "123"
Worksheets(3).Unprotect Password:="123"
Worksheets(4).Unprotect Password:="123"
Worksheets(5).Unprotect Password:="123"


'Columns("CE:CF").EntireColumn.Hidden = False
'Columns("CK:CN").EntireColumn.Hidden = False

If TutanakTarihiText.Value = "" Then
    MsgBox "The statement date has not been specified, so the operation cannot be completed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

If Tip1Option.Value = False And Tip2Option.Value = False And Tip3Option.Value = False And Tip4Option.Value = False Then
    MsgBox "The statement type (e.g., Type 1, Type 2, etc.) has not been specified, so the operation cannot be completed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

If TutanakNoText.Value = "" Then
    MsgBox "The statement number has not been specified, so the operation cannot be completed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If


VeriKontrol = 0
For Each ctl In core_delivery_manager_UI.Frame2.Controls
    If TypeName(ctl) = "ListBox" Then
        VeriKontrol = VeriKontrol + 1
        GoTo Devam2
    End If
Next ctl
Devam2:
If VeriKontrol = 0 Then
    MsgBox "The operation cannot be completed because no transfer has been made to the 'To Be Sent' section for the deliveries to be executed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

'Pathfinder...
AutoPath = ThisWorkbook.Path
DestOperasyon = AutoPath & "\System Files\Operation\"
'EK TANIMLARI
If Tip2Option.Value = True Then
    EkFormu = "Type 2 Delivery Statement"
ElseIf Tip3Option.Value = True Then
    EkFormu = "Type 3 Delivery Statement"
ElseIf Tip4Option.Value = True Then
    EkFormu = "Type 4 Delivery Statement"
    EkFormuTip5 = "Type 5 Delivery Statement"
ElseIf Tip1Option.Value = True Then
    EkFormu = "Type 1 Delivery Statement"
Else
    MsgBox "The operation cannot be completed because the form type has not been selected from the top-right corner.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

SourceTeslim = AutoPath & "\System Files\System Templates\Acceptance and Delivery Statements\" & EkFormu & ".docm"
SourceTip5 = AutoPath & "\System Files\System Templates\Acceptance and Delivery Statements\" & EkFormuTip5 & ".docm"

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
If Not Dir(SourceTeslim, vbDirectory) <> vbNullString Then
    MsgBox SourceTeslim & " directory could not be accessed. Folder and/or file names under this directory may have been renamed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

'Operation klasörü içinde kullanıcı modülü klasörü yoksa oluştur.
If Not Dir(DestOpUserFolder, vbDirectory) <> vbNullString Then
    MkDir DestOpUserFolder
End If

ReNameTeslim = EkFormu
ReNameTip5 = EkFormuTip5

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
fso.CopyFile (SourceTeslim), DestOpUserFolder & ReNameTeslim & ".docm", True
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
objWord.Documents.Open FileName:=DestOpUserFolder & ReNameTeslim & ".docm"
objWord.Visible = True
objWord.Activate 'Ekrana getirir.
'objDoc.Activate 'Ekrana getirmez.
'objWord.Application.WindowState = wdWindowStateMaximize
Set objDoc = GetObject(DestOpUserFolder & ReNameTeslim & ".docm")
'________________________________________


'Birim
Birimx = UCase(Replace(Replace(Worksheets(2).Cells(6, 99).Value, "i", "I"), "ı", "I")) & " UNIT"
objDoc.Tables(1).Cell(Row:=1, Column:=1).Range.Text = Birimx
If Tip2Option.Value = True Or Tip3Option.Value = True Or Tip4Option.Value = True Then
    'Tutanak tarihi
    objDoc.Tables(1).Cell(Row:=5, Column:=3).Range.Text = TutanakTarihiText.Value
    'Tutanak no
    objDoc.Tables(1).Cell(Row:=4, Column:=3).Range.Text = TutanakNoText.Value
End If


'________________AKTARIMLAR 1 (Başlangıç)

Say2 = 0
AdetSay = 0
For Each ctl In core_delivery_manager_UI.Frame2.Controls '________________RAPOR1 BÖLÜMÜ (Başlangıç, Tüm Raporların Ortak Bölümü)
    If TypeName(ctl) = "ListBox" Then
        'MsgBox Mid(Ctl.List(0), InStr(Ctl.List(0), "|") + 1, 2)
        If Mid(ctl.List(0), InStr(ctl.List(0), "|") + 1, 2) = " Ö" Then

            Call Rapor1GidenTema
            If IlkSiraGlobal = 0 Then
                GoTo Rapor1Devam1
            End If
            
            IlkSira = IlkSiraGlobal
            GidenTema = GidenTemaGlobal
            'Tabloya satır ekle
            Say2 = Say2 + 1
            'MsgBox Left(Ctl.List(0), InStr(Ctl.List(0), "|") - 2)
            If Say2 > 1 Then
                objDoc.Tables(2).Rows.Add BeforeRow:=objDoc.Tables(2).Rows(Say2 + 1)
            End If
            'Sıra no
            objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=1).Range.Text = Say2
            'Gönderilen birim
            objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=2).Range.Text = GidenTema
            'Çıkış kayıt no
            objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=3).Range.Text = ThisWorkbook.Worksheets(3).Cells(IlkSira, 76).Value
            'Package A/Package B/Package C
            objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=4).Range.Text = 1 * ThisWorkbook.Worksheets(3).Cells(IlkSira, 68).Value & " unit(s) of " & Left(ThisWorkbook.Worksheets(3).Cells(IlkSira, 67).Value, InStr(ThisWorkbook.Worksheets(3).Cells(IlkSira, 67).Value, "/") - 1)
            AdetSay = AdetSay + ThisWorkbook.Worksheets(3).Cells(IlkSira, 68).Value
            If Tip2Option.Value = True Then
                'Gönderim şekli
                objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=5).Range.Text = WorksheetFunction.Proper(Mid(ThisWorkbook.Worksheets(3).Cells(IlkSira, 67).Value, InStr(ThisWorkbook.Worksheets(3).Cells(IlkSira, 67).Value, "/") + 1, Len(ThisWorkbook.Worksheets(3).Cells(IlkSira, 67).Value) - InStr(ThisWorkbook.Worksheets(3).Cells(IlkSira, 67).Value, "/") + 1))
            End If
            'Kilitlenme bilgisi
'            ThisWorkbook.Worksheets(3).Cells(IlkSira, 86).Value = "*"
        End If
    End If
Rapor1Devam1:
Next ctl '________________RAPOR1 BÖLÜMÜ (Bitiş, Tüm Raporların Ortak Bölümü)

For Each ctl In core_delivery_manager_UI.Frame2.Controls '________________RAPOR BÖLÜMÜ (Başlangıç, Tüm Raporların Ortak Bölümü)'Rapor ve bilgilendirme için ortak
    If TypeName(ctl) = "ListBox" Then
        'MsgBox Mid(Ctl.List(0), InStr(Ctl.List(0), "|") + 1, 2)
        If Mid(ctl.List(0), InStr(ctl.List(0), "|") + 1, 2) = " R" Then
            Call RaporGidenTema
            If IlkSiraGlobal = 0 Then
                GoTo RaporDevam1
            End If
            
            IlkSira = IlkSiraGlobal
            GidenTema = GidenTemaGlobal

            'Tabloya satır ekle
            Say2 = Say2 + 1
            'MsgBox Left(Ctl.List(0), InStr(Ctl.List(0), "|") - 2)
            If Say2 > 1 Then
                objDoc.Tables(2).Rows.Add BeforeRow:=objDoc.Tables(2).Rows(Say2 + 1)
            End If
            'Sıra no
            objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=1).Range.Text = Say2
            'Gönderilen birim
            objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=2).Range.Text = GidenTema
            'Çıkış kayıt no
            objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=3).Range.Text = ThisWorkbook.Worksheets(4).Cells(IlkSira, 84).Value

            If ThisWorkbook.Worksheets(4).Cells(IlkSiraGlobal, 174).Value = "Yes" Then 'Bilgilendirme (Rapor2_2 evet)
                objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=2).Range.Text = GidenTema & " (Bilgilendirme)"
                'Package A/Package B/Package C
                objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=4).Range.Text = 1 * 1 & " unit(s) of " & Left(ThisWorkbook.Worksheets(4).Cells(IlkSira, 81).Value, InStr(ThisWorkbook.Worksheets(4).Cells(IlkSira, 81).Value, "/") - 1) 'ThisWorkbook.Worksheets(4).Cells(IlkSira, 76).Value & " unit(s) of " & Left(ThisWorkbook.Worksheets(4).Cells(IlkSira, 75).Value, InStr(ThisWorkbook.Worksheets(4).Cells(IlkSira, 75).Value, "/") - 1)
                AdetSay = AdetSay + 1 'ThisWorkbook.Worksheets(4).Cells(IlkSira, 76).Value
                If Tip2Option.Value = True Then
                    'Gönderim şekli
                    objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=5).Range.Text = WorksheetFunction.Proper(Mid(ThisWorkbook.Worksheets(4).Cells(IlkSira, 81).Value, InStr(ThisWorkbook.Worksheets(4).Cells(IlkSira, 81).Value, "/") + 1, Len(ThisWorkbook.Worksheets(4).Cells(IlkSira, 81).Value) - InStr(ThisWorkbook.Worksheets(4).Cells(IlkSira, 81).Value, "/") + 1))
                End If
            Else
                'Package A/Package B/Package C
                objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=4).Range.Text = 1 * ThisWorkbook.Worksheets(4).Cells(IlkSira, 76).Value & " unit(s) of " & Left(ThisWorkbook.Worksheets(4).Cells(IlkSira, 75).Value, InStr(ThisWorkbook.Worksheets(4).Cells(IlkSira, 75).Value, "/") - 1)
                AdetSay = AdetSay + ThisWorkbook.Worksheets(4).Cells(IlkSira, 76).Value
                If Tip2Option.Value = True Then
                    'Gönderim şekli
                    objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=5).Range.Text = WorksheetFunction.Proper(Mid(ThisWorkbook.Worksheets(4).Cells(IlkSira, 75).Value, InStr(ThisWorkbook.Worksheets(4).Cells(IlkSira, 75).Value, "/") + 1, Len(ThisWorkbook.Worksheets(4).Cells(IlkSira, 75).Value) - InStr(ThisWorkbook.Worksheets(4).Cells(IlkSira, 75).Value, "/") + 1))
                End If
            End If
            'Kilitlenme bilgisi
            'ThisWorkbook.Worksheets(4).Cells(IlkSira, 94).Value = "*"
        End If
    End If
RaporDevam1:
Next ctl '________________RAPOR BÖLÜMÜ (Bitiş, Tüm Raporların Ortak Bölümü)


For Each ctl In core_delivery_manager_UI.Frame2.Controls '________________RAPOR BÖLÜMÜ (Başlangıç, Tüm Raporların Ortak Bölümü)'XXXMud
    If TypeName(ctl) = "ListBox" Then
        'MsgBox Mid(Ctl.List(0), InStr(Ctl.List(0), "|") + 1, 2)
        If Mid(ctl.List(0), InStr(ctl.List(0), "|") + 1, 2) = " B" And InStr(ctl.List(0), "XXXMud") <> 0 Then
            Call RaporGidenTema
            If IlkSiraGlobal = 0 Then
                GoTo RaporDevam1XXXMud
            End If
            
            IlkSira = IlkSiraGlobal
            GidenTema = GidenTemaGlobal

            'Tabloya satır ekle
            Say2 = Say2 + 1
            'MsgBox Left(Ctl.List(0), InStr(Ctl.List(0), "|") - 2)
            If Say2 > 1 Then
                objDoc.Tables(2).Rows.Add BeforeRow:=objDoc.Tables(2).Rows(Say2 + 1)
            End If
            'Sıra no
            objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=1).Range.Text = Say2
            'Gönderilen birim
            GidenTema = "ORGANIZATION A XXX Directorate"
            objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=2).Range.Text = GidenTema
            'Çıkış kayıt no
            objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=3).Range.Text = ThisWorkbook.Worksheets(4).Cells(IlkSira, 176).Value
            'Package A/Package B/Package C
            objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=4).Range.Text = 1 * ThisWorkbook.Worksheets(4).Cells(IlkSira, 185).Value & " unit(s) of " & Left(ThisWorkbook.Worksheets(4).Cells(IlkSira, 184).Value, InStr(ThisWorkbook.Worksheets(4).Cells(IlkSira, 184).Value, "/") - 1)
            AdetSay = AdetSay + ThisWorkbook.Worksheets(4).Cells(IlkSira, 185).Value
            If Tip2Option.Value = True Then
                'Gönderim şekli
                objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=5).Range.Text = WorksheetFunction.Proper(Mid(ThisWorkbook.Worksheets(4).Cells(IlkSira, 184).Value, InStr(ThisWorkbook.Worksheets(4).Cells(IlkSira, 184).Value, "/") + 1, Len(ThisWorkbook.Worksheets(4).Cells(IlkSira, 184).Value) - InStr(ThisWorkbook.Worksheets(4).Cells(IlkSira, 184).Value, "/") + 1))
            End If
            'Kilitlenme bilgisi
            'ThisWorkbook.Worksheets(4).Cells(IlkSira, 94).Value = "*"
        End If
    End If
RaporDevam1XXXMud:
Next ctl '________________RAPOR BÖLÜMÜ (Bitiş, Tüm Raporların Ortak Bölümü)


For Each ctl In core_delivery_manager_UI.Frame2.Controls '________________RAPOR BÖLÜMÜ (Başlangıç, Tüm Raporların Ortak Bölümü)'Sonuç
    If TypeName(ctl) = "ListBox" Then
        'MsgBox Mid(Ctl.List(0), InStr(Ctl.List(0), "|") + 1, 2)
        If Mid(ctl.List(0), InStr(ctl.List(0), "|") + 1, 2) = " B" And InStr(ctl.List(0), "XXXMud") = 0 Then

            Call Rapor2_2GidenTema
            If IlkSiraGlobal = 0 Then
                GoTo RaporDevam1Sonuc
            End If
            
            IlkSira = IlkSiraGlobal
            GidenTema = GidenTemaGlobal

            'Tabloya satır ekle
            Say2 = Say2 + 1
            'MsgBox Left(Ctl.List(0), InStr(Ctl.List(0), "|") - 2)
            If Say2 > 1 Then
                objDoc.Tables(2).Rows.Add BeforeRow:=objDoc.Tables(2).Rows(Say2 + 1)
            End If
            'Sıra no
            objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=1).Range.Text = Say2
            'Gönderilen birim
            objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=2).Range.Text = GidenTema & " (Sonuç)"
            'Çıkış kayıt no
            objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=3).Range.Text = ThisWorkbook.Worksheets(4).Cells(IlkSira, 178).Value


            'Senaryo düzenleyicisi (veriler  kurum tutanak2sından mı, XXXMud tutanak2sından mı yok sa her ikisinden mi alınacak)
            If ThisWorkbook.Worksheets(4).Cells(IlkSira, 187).Value = "All" Then 'XXXMudya giden
                If ThisWorkbook.Worksheets(4).Cells(IlkSira, 188).Value = "Yes" Then 'XXXMud'den gelen paketin açılma durumu
                    'Kurum
                    GidenPaketAdet = Application.Sum(ThisWorkbook.Worksheets(4).Range(ThisWorkbook.Worksheets(4).Cells(IlkSira, 76), ThisWorkbook.Worksheets(4).Cells(IlkSira, 76)))
                    'Package A/Package B/Package C
                    objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=4).Range.Text = 1 * GidenPaketAdet & " unit(s) of " & Left(ThisWorkbook.Worksheets(4).Cells(IlkSira, 75).Value, InStr(ThisWorkbook.Worksheets(4).Cells(IlkSira, 75).Value, "/") - 1)
                    AdetSay = AdetSay + GidenPaketAdet
                    If Tip2Option.Value = True Then
                        'Gönderim şekli
                        objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=5).Range.Text = WorksheetFunction.Proper(Mid(ThisWorkbook.Worksheets(4).Cells(IlkSira, 75).Value, InStr(ThisWorkbook.Worksheets(4).Cells(IlkSira, 75).Value, "/") + 1, Len(ThisWorkbook.Worksheets(4).Cells(IlkSira, 75).Value) - InStr(ThisWorkbook.Worksheets(4).Cells(IlkSira, 75).Value, "/") + 1))
                    End If
                ElseIf ThisWorkbook.Worksheets(4).Cells(IlkSira, 188).Value = "No" Then 'XXXMud'den gelen paketin açılma durumu
                    'XXXMud
                    GidenPaketAdet = Application.Sum(ThisWorkbook.Worksheets(4).Range(ThisWorkbook.Worksheets(4).Cells(IlkSira, 202), ThisWorkbook.Worksheets(4).Cells(IlkSira, 202)))
                    'Package A/Package B/Package C
                    objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=4).Range.Text = 1 * GidenPaketAdet & " unit(s) of " & Left(ThisWorkbook.Worksheets(4).Cells(IlkSira, 201).Value, InStr(ThisWorkbook.Worksheets(4).Cells(IlkSira, 201).Value, "/") - 1)
                    AdetSay = AdetSay + GidenPaketAdet
                    If Tip2Option.Value = True Then
                        'Gönderim şekli
                        objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=5).Range.Text = WorksheetFunction.Proper(Mid(ThisWorkbook.Worksheets(4).Cells(IlkSira, 201).Value, InStr(ThisWorkbook.Worksheets(4).Cells(IlkSira, 201).Value, "/") + 1, Len(ThisWorkbook.Worksheets(4).Cells(IlkSira, 201).Value) - InStr(ThisWorkbook.Worksheets(4).Cells(IlkSira, 201).Value, "/") + 1))
                    End If
                End If
            ElseIf ThisWorkbook.Worksheets(4).Cells(IlkSira, 187).Value = "Technique A" Then 'XXXMudya giden
                If ThisWorkbook.Worksheets(4).Cells(IlkSira, 188).Value = "Yes" Then 'XXXMud'den gelen paketin açılma durumu
                    If ThisWorkbook.Worksheets(4).Cells(IlkSira, 189).Value = "Yes" Then 'Tutanak2 tut. birleşme durumu
                        'Kurum
                        GidenPaketAdet = Application.Sum(ThisWorkbook.Worksheets(4).Range(ThisWorkbook.Worksheets(4).Cells(IlkSira, 76), ThisWorkbook.Worksheets(4).Cells(IlkSira, 76)))
                        'Package A/Package B/Package C
                        objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=4).Range.Text = 1 * GidenPaketAdet & " unit(s) of " & Left(ThisWorkbook.Worksheets(4).Cells(IlkSira, 75).Value, InStr(ThisWorkbook.Worksheets(4).Cells(IlkSira, 75).Value, "/") - 1)
                        AdetSay = AdetSay + GidenPaketAdet
                        If Tip2Option.Value = True Then
                            'Gönderim şekli
                            objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=5).Range.Text = WorksheetFunction.Proper(Mid(ThisWorkbook.Worksheets(4).Cells(IlkSira, 75).Value, InStr(ThisWorkbook.Worksheets(4).Cells(IlkSira, 75).Value, "/") + 1, Len(ThisWorkbook.Worksheets(4).Cells(IlkSira, 75).Value) - InStr(ThisWorkbook.Worksheets(4).Cells(IlkSira, 75).Value, "/") + 1))
                        End If
                    ElseIf ThisWorkbook.Worksheets(4).Cells(IlkSira, 189).Value = "No" Then 'Tutanak2 tut. birleşme durumu
                        'Kurum ve XXXMud
                        GidenPaketAdet = Application.Sum(ThisWorkbook.Worksheets(4).Range(ThisWorkbook.Worksheets(4).Cells(IlkSira, 76), ThisWorkbook.Worksheets(4).Cells(IlkSira, 76)))
                        GidenPaketAdet = GidenPaketAdet + Application.Sum(ThisWorkbook.Worksheets(4).Range(ThisWorkbook.Worksheets(4).Cells(IlkSira, 202), ThisWorkbook.Worksheets(4).Cells(IlkSira, 202)))
                        'Package A/Package B/Package C
                        objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=4).Range.Text = 1 * GidenPaketAdet & " unit(s) of " & Left(ThisWorkbook.Worksheets(4).Cells(IlkSira, 75).Value, InStr(ThisWorkbook.Worksheets(4).Cells(IlkSira, 75).Value, "/") - 1)
                        AdetSay = AdetSay + GidenPaketAdet
                        If Tip2Option.Value = True Then
                            'Gönderim şekli
                            objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=5).Range.Text = WorksheetFunction.Proper(Mid(ThisWorkbook.Worksheets(4).Cells(IlkSira, 75).Value, InStr(ThisWorkbook.Worksheets(4).Cells(IlkSira, 75).Value, "/") + 1, Len(ThisWorkbook.Worksheets(4).Cells(IlkSira, 75).Value) - InStr(ThisWorkbook.Worksheets(4).Cells(IlkSira, 75).Value, "/") + 1))
                        End If
                    End If
                ElseIf ThisWorkbook.Worksheets(4).Cells(IlkSira, 188).Value = "No" Then 'XXXMud'den gelen paketin açılma durumu
                    'Kurum
                    GidenPaketAdet = Application.Sum(ThisWorkbook.Worksheets(4).Range(ThisWorkbook.Worksheets(4).Cells(IlkSira, 76), ThisWorkbook.Worksheets(4).Cells(IlkSira, 76)))
                    'Package A/Package B/Package C
                    objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=4).Range.Text = 1 * GidenPaketAdet & " unit(s) of " & Left(ThisWorkbook.Worksheets(4).Cells(IlkSira, 75).Value, InStr(ThisWorkbook.Worksheets(4).Cells(IlkSira, 75).Value, "/") - 1)
                    AdetSay = AdetSay + GidenPaketAdet
                    If Tip2Option.Value = True Then
                        'Gönderim şekli
                        objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=5).Range.Text = WorksheetFunction.Proper(Mid(ThisWorkbook.Worksheets(4).Cells(IlkSira, 75).Value, InStr(ThisWorkbook.Worksheets(4).Cells(IlkSira, 75).Value, "/") + 1, Len(ThisWorkbook.Worksheets(4).Cells(IlkSira, 75).Value) - InStr(ThisWorkbook.Worksheets(4).Cells(IlkSira, 75).Value, "/") + 1))
                    End If
                End If
            End If

            'Kilitlenme bilgisi
            'ThisWorkbook.Worksheets(4).Cells(IlkSira, 94).Value = "*"
        End If
    End If
RaporDevam1Sonuc:
Next ctl '________________RAPOR BÖLÜMÜ (Bitiş, Tüm Raporların Ortak Bölümü)


For Each ctl In core_delivery_manager_UI.Frame2.Controls '________________RAPOR3_1 BÖLÜMÜ (Başlangıç, Tüm Raporların Ortak Bölümü)
    If TypeName(ctl) = "ListBox" Then
        'MsgBox Mid(Ctl.List(0), InStr(Ctl.List(0), "|") + 1, 2)
        If Mid(ctl.List(0), InStr(ctl.List(0), "|") + 1, 2) = " G" Then

            Call Rapor3_1GidenTema
            If IlkSiraGlobal = 0 Then
                GoTo Rapor3_1Devam1
            End If
            
            IlkSira = IlkSiraGlobal
            GidenTema = GidenTemaGlobal
            'Tabloya satır ekle
            Say2 = Say2 + 1
            'MsgBox Left(Ctl.List(0), InStr(Ctl.List(0), "|") - 2)
            If Say2 > 1 Then
                objDoc.Tables(2).Rows.Add BeforeRow:=objDoc.Tables(2).Rows(Say2 + 1)
            End If
            'Sıra no
            objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=1).Range.Text = Say2
            'Gönderilen birim
            objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=2).Range.Text = GidenTema
            'Çıkış kayıt no
            objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=3).Range.Text = ThisWorkbook.Worksheets(5).Cells(IlkSira, 156).Value
            'Package A/Package B/Package C
            objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=4).Range.Text = 1 * ThisWorkbook.Worksheets(5).Cells(IlkSira, 150).Value & " unit(s) of " & Left(ThisWorkbook.Worksheets(5).Cells(IlkSira, 149).Value, InStr(ThisWorkbook.Worksheets(5).Cells(IlkSira, 149).Value, "/") - 1)
            AdetSay = AdetSay + ThisWorkbook.Worksheets(5).Cells(IlkSira, 150).Value
            If Tip2Option.Value = True Then
                'Gönderim şekli
                objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=5).Range.Text = WorksheetFunction.Proper(Mid(ThisWorkbook.Worksheets(5).Cells(IlkSira, 149).Value, InStr(ThisWorkbook.Worksheets(5).Cells(IlkSira, 149).Value, "/") + 1, Len(ThisWorkbook.Worksheets(5).Cells(IlkSira, 149).Value) - InStr(ThisWorkbook.Worksheets(5).Cells(IlkSira, 149).Value, "/") + 1))
            End If
            'Kilitlenme bilgisi
'            ThisWorkbook.Worksheets(5).Cells(IlkSira, 166).Value = "*"
        End If
    End If
Rapor3_1Devam1:
Next ctl '________________RAPOR3_1 BÖLÜMÜ (Bitiş, Tüm Raporların Ortak Bölümü)

'MsgBox "Burada mısın?"
For Each ctl In core_delivery_manager_UI.Frame2.Controls '________________RAPOR3_2 BÖLÜMÜ (Başlangıç, Tüm Raporların Ortak Bölümü)
    If TypeName(ctl) = "ListBox" Then
        'MsgBox Mid(Ctl.List(0), InStr(Ctl.List(0), "|") + 1, 2)
        If Mid(ctl.List(0), InStr(ctl.List(0), "|") + 1, 2) = " M" Then 'And Mid(ctl.List(0), InStrRev(ctl.List(0), "-") - 5, 5) <> "FinansalBirim" Then 'BU YÖNTEMDE HATA VAR. TUTANAK İÇİN RAPOR1 VEYA RAPOR3_1 SEÇİLİRSE BURDAKİ KOŞULUN İÇİNE GİREBİLİYOR
            
            'MsgBox "Test1: " & Mid(ctl.List(0), InStrRev(ctl.List(0), "-") - 5, 5)
            
            'If Mid(ctl.List(0), InStrRev(ctl.List(0), "-") - 5, 5) = "FinansalBirim" Then 'FinansalBirim
            If InStr(ctl.List(0), "(FinansalBirim-TipA)") <> 0 Or InStr(ctl.List(0), "(FinansalBirim-TipB)") <> 0 Then
    
                Call Rapor3_2GidenTema
                If IlkSiraGlobal = 0 Then
                    GoTo Rapor3_2Devam1
                End If
    
                IlkSira = IlkSiraGlobal
                GidenTema = ThisWorkbook.Worksheets(5).Cells(IlkSira, 30).Value & " " & ThisWorkbook.Worksheets(5).Cells(IlkSira, 31).Value 'GidenTemaGlobal
                'Tabloya satır ekle
                Say2 = Say2 + 1
                'MsgBox Left(Ctl.List(0), InStr(Ctl.List(0), "|") - 2)
                If Say2 > 1 Then
                    objDoc.Tables(2).Rows.Add BeforeRow:=objDoc.Tables(2).Rows(Say2 + 1)
                End If
                'Sıra no
                objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=1).Range.Text = Say2
    
                'Gönderilen birim
                objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=2).Range.Text = GidenTema
                'Çıkış kayıt no
                objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=3).Range.Text = ThisWorkbook.Worksheets(5).Cells(IlkSira, 76).Value
    
                'Package A/Package B/Package C
                objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=4).Range.Text = "1 Adet Package A"
                AdetSay = AdetSay + 1
                If Tip2Option.Value = True Then
                    'Gönderim şekli
                    objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=5).Range.Text = ThisWorkbook.Worksheets(5).Cells(IlkSira, 85).Value
                End If
                
            Else 'DIRECTORATELIK
                
                'MsgBox "Test2: " & Mid(ctl.List(0), InStrRev(ctl.List(0), "-") - 5, 5)
                
                'MsgBox "Kontrol"
                Call Rapor3_2GidenTema
                If IlkSiraGlobal = 0 Then
                    GoTo Rapor3_2Devam1
                End If
                
                IlkSira = IlkSiraGlobal
                GidenTema = GidenTemaGlobal
                'Tabloya satır ekle
                Say2 = Say2 + 1
                'MsgBox Left(Ctl.List(0), InStr(Ctl.List(0), "|") - 2)
                If Say2 > 1 Then
                    objDoc.Tables(2).Rows.Add BeforeRow:=objDoc.Tables(2).Rows(Say2 + 1)
                End If
                'Sıra no
                objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=1).Range.Text = Say2
                'Gönderilen birim
                objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=2).Range.Text = GidenTema
                'Çıkış kayıt no
                objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=3).Range.Text = ThisWorkbook.Worksheets(5).Cells(IlkSira, 84).Value
                'Package A/Package B/Package C
                objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=4).Range.Text = 1 * ThisWorkbook.Worksheets(5).Cells(IlkSira, 72).Value & " unit(s) of " & Left(ThisWorkbook.Worksheets(5).Cells(IlkSira, 71).Value, InStr(ThisWorkbook.Worksheets(5).Cells(IlkSira, 71).Value, "/") - 1)
                AdetSay = AdetSay + ThisWorkbook.Worksheets(5).Cells(IlkSira, 72).Value
                If Tip2Option.Value = True Then
                    'Gönderim şekli
                    objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=5).Range.Text = WorksheetFunction.Proper(Mid(ThisWorkbook.Worksheets(5).Cells(IlkSira, 71).Value, InStr(ThisWorkbook.Worksheets(5).Cells(IlkSira, 71).Value, "/") + 1, Len(ThisWorkbook.Worksheets(5).Cells(IlkSira, 71).Value) - InStr(ThisWorkbook.Worksheets(5).Cells(IlkSira, 71).Value, "/") + 1))
                End If
                'Kilitlenme bilgisi
    '            ThisWorkbook.Worksheets(5).Cells(IlkSira, 166).Value = "*"
            End If
        End If
    End If
Rapor3_2Devam1:
Next ctl '________________RAPOR3_2 BÖLÜMÜ (Bitiş, Tüm Raporların Ortak Bölümü)



'________________AKTARIMLAR 1 (Bitiş)


If Tip2Option.Value = True Or Tip3Option.Value = True Or Tip4Option.Value = True Then
    AdetSay = AdetSay * 1
    If AdetSay > 1 Then
        Ek1 = "items"
    Else
        Ek1 = "item"
    End If
    If Tip2Option.Value = True Then
        'İlk metin
        objDoc.Tables(3).Cell(Row:=1, Column:=1).Range.Text = "Delivery of the " & AdetSay & " " & Ek1 & " mentioned above (Package A/B/C) has been completed."
    ElseIf Tip3Option.Value = True Or Tip4Option.Value = True Then
        'İlk metin
        objDoc.Tables(3).Cell(Row:=1, Column:=1).Range.Text = "The " & AdetSay & " " & Ek1 & " mentioned above (Package A/B/C) have been handed over to the authorized personnel."
    End If
    'İkinci metin
    objDoc.Tables(3).Cell(Row:=11, Column:=1).Range.Text = "The " & AdetSay & " " & Ek1 & " (Package A/B/C) have been received."
    'Adedi bold yap.
    Set MyRange = objDoc.Tables(3).Cell(Row:=1, Column:=1).Range
    With MyRange.Find
        .Execute FindText:=AdetSay
    End With
    MyRange.Font.Bold = True
    'Adedi bold yap.
    Set MyRange = objDoc.Tables(3).Cell(Row:=11, Column:=1).Range
    With MyRange.Find
        .Execute FindText:=AdetSay
        '.Execute Forward:=True
    End With
    MyRange.Font.Bold = True

    'Alt bilgi ekle
    objDoc.Sections(1).Footers(wdHeaderFooterPrimary).Range.Tables(1).Cell(Row:=1, Column:=1).Range.Text = ReNameTeslim
    
End If

'Artık satırı sil
objDoc.Tables(2).Rows(Say2 + 2).Delete
If Tip2Option.Value = True Then
'    'İmza boşluğunu sayfaya sığdırmak için düzenle
    If Say2 = 13 Then
        objDoc.Tables(3).Rows(2).Delete
        objDoc.Tables(3).Rows(11).Delete
    ElseIf Say2 = 14 Then
        objDoc.Tables(3).Rows(2).Delete
        objDoc.Tables(3).Rows(11).Delete
        For i = 1 To 2
            objDoc.Tables(3).Rows(8).Delete
        Next i
    ElseIf Say2 = 15 Then
        objDoc.Tables(3).Rows(2).Delete
        objDoc.Tables(3).Rows(11).Delete
        For i = 1 To 2
            objDoc.Tables(3).Rows(8).Delete
        Next i
        objDoc.Tables(3).Rows(4).Delete
        objDoc.Tables(3).Rows(9).Delete
    ElseIf Say2 = 16 Then
        objDoc.Tables(3).Rows(2).Delete
        objDoc.Tables(3).Rows(11).Delete
        For i = 1 To 2
            objDoc.Tables(3).Rows(8).Delete
        Next i
        For i = 1 To 2
            objDoc.Tables(3).Rows(4).Delete
            objDoc.Tables(3).Rows(9).Delete
        Next i
    ElseIf Say2 = 17 Then
        objDoc.Tables(3).Rows(2).Delete
        objDoc.Tables(3).Rows(11).Delete
        For i = 1 To 2
            objDoc.Tables(3).Rows(8).Delete
        Next i
        For i = 1 To 2
            objDoc.Tables(3).Rows(4).Delete
            objDoc.Tables(3).Rows(9).Delete
        Next i
        objDoc.Tables(3).Rows(4).Delete
        objDoc.Tables(3).Rows(8).Delete
    ElseIf Say2 = 18 Then
        objDoc.Tables(3).Rows(2).Delete
        objDoc.Tables(3).Rows(11).Delete
        For i = 1 To 2
            objDoc.Tables(3).Rows(8).Delete
        Next i
        For i = 1 To 2
            objDoc.Tables(3).Rows(4).Delete
            objDoc.Tables(3).Rows(9).Delete
        Next i
        objDoc.Tables(3).Rows(4).Delete
        objDoc.Tables(3).Rows(8).Delete
        objDoc.Tables(3).Rows(4).Delete
        objDoc.Tables(3).Rows(7).Delete
    ElseIf Say2 = 19 Then
        objDoc.Tables(3).Rows(2).Delete
        objDoc.Tables(3).Rows(11).Delete
        For i = 1 To 2
            objDoc.Tables(3).Rows(8).Delete
        Next i
        For i = 1 To 2
            objDoc.Tables(3).Rows(4).Delete
            objDoc.Tables(3).Rows(9).Delete
        Next i
        objDoc.Tables(3).Rows(4).Delete
        objDoc.Tables(3).Rows(8).Delete
        objDoc.Tables(3).Rows(4).Delete
        objDoc.Tables(3).Rows(7).Delete
        objDoc.Tables(3).Rows(3).Delete
        objDoc.Tables(3).Rows(5).Delete
    ElseIf Say2 = 20 Then
        For i = 2 To 21
            objDoc.Tables(2).Rows(i).Height = 27
        Next i
    ElseIf Say2 = 39 Then
        objDoc.Tables(3).Rows(2).Delete
        objDoc.Tables(3).Rows(11).Delete
    ElseIf Say2 = 40 Then
        objDoc.Tables(3).Rows(2).Delete
        objDoc.Tables(3).Rows(11).Delete
        For i = 1 To 2
            objDoc.Tables(3).Rows(8).Delete
        Next i
    ElseIf Say2 = 41 Then
        objDoc.Tables(3).Rows(2).Delete
        objDoc.Tables(3).Rows(11).Delete
        For i = 1 To 2
            objDoc.Tables(3).Rows(8).Delete
        Next i
        objDoc.Tables(3).Rows(4).Delete
        objDoc.Tables(3).Rows(9).Delete
    ElseIf Say2 = 42 Then
        objDoc.Tables(3).Rows(2).Delete
        objDoc.Tables(3).Rows(11).Delete
        For i = 1 To 2
            objDoc.Tables(3).Rows(8).Delete
        Next i
        For i = 1 To 2
            objDoc.Tables(3).Rows(4).Delete
            objDoc.Tables(3).Rows(9).Delete
        Next i
    ElseIf Say2 = 43 Then
        objDoc.Tables(3).Rows(2).Delete
        objDoc.Tables(3).Rows(11).Delete
        For i = 1 To 2
            objDoc.Tables(3).Rows(8).Delete
        Next i
        For i = 1 To 2
            objDoc.Tables(3).Rows(4).Delete
            objDoc.Tables(3).Rows(9).Delete
        Next i
        objDoc.Tables(3).Rows(4).Delete
        objDoc.Tables(3).Rows(8).Delete
    ElseIf Say2 = 44 Then
        objDoc.Tables(3).Rows(2).Delete
        objDoc.Tables(3).Rows(11).Delete
        For i = 1 To 2
            objDoc.Tables(3).Rows(8).Delete
        Next i
        For i = 1 To 2
            objDoc.Tables(3).Rows(4).Delete
            objDoc.Tables(3).Rows(9).Delete
        Next i
        objDoc.Tables(3).Rows(4).Delete
        objDoc.Tables(3).Rows(8).Delete
        objDoc.Tables(3).Rows(4).Delete
        objDoc.Tables(3).Rows(7).Delete
    ElseIf Say2 = 45 Then
        objDoc.Tables(3).Rows(2).Delete
        objDoc.Tables(3).Rows(11).Delete
        For i = 1 To 2
            objDoc.Tables(3).Rows(8).Delete
        Next i
        For i = 1 To 2
            objDoc.Tables(3).Rows(4).Delete
            objDoc.Tables(3).Rows(9).Delete
        Next i
        objDoc.Tables(3).Rows(4).Delete
        objDoc.Tables(3).Rows(8).Delete
        objDoc.Tables(3).Rows(4).Delete
        objDoc.Tables(3).Rows(7).Delete
        objDoc.Tables(3).Rows(3).Delete
        objDoc.Tables(3).Rows(5).Delete
    ElseIf Say2 = 46 Then
        For i = 23 To 47
            objDoc.Tables(2).Rows(i).Height = 27
        Next i
    End If
ElseIf Tip3Option.Value = True Then
    
'    'İmza boşluğunu sayfaya sığdırmak için düzenle
    If Say2 = 10 Then
        objDoc.Tables(3).Rows(2).Delete
        objDoc.Tables(3).Rows(11).Delete
    ElseIf Say2 = 11 Then
        objDoc.Tables(3).Rows(2).Delete
        objDoc.Tables(3).Rows(11).Delete
        For i = 1 To 2
            objDoc.Tables(3).Rows(8).Delete
        Next i
    ElseIf Say2 = 12 Then
        objDoc.Tables(3).Rows(2).Delete
        objDoc.Tables(3).Rows(11).Delete
        For i = 1 To 2
            objDoc.Tables(3).Rows(8).Delete
        Next i
        objDoc.Tables(3).Rows(4).Delete
        objDoc.Tables(3).Rows(9).Delete
    ElseIf Say2 = 13 Then
        objDoc.Tables(3).Rows(2).Delete
        objDoc.Tables(3).Rows(11).Delete
        For i = 1 To 2
            objDoc.Tables(3).Rows(8).Delete
        Next i
        For i = 1 To 2
            objDoc.Tables(3).Rows(4).Delete
            objDoc.Tables(3).Rows(9).Delete
        Next i
    ElseIf Say2 = 14 Then
        objDoc.Tables(3).Rows(2).Delete
        objDoc.Tables(3).Rows(11).Delete
        For i = 1 To 2
            objDoc.Tables(3).Rows(8).Delete
        Next i
        For i = 1 To 2
            objDoc.Tables(3).Rows(4).Delete
            objDoc.Tables(3).Rows(9).Delete
        Next i
        objDoc.Tables(3).Rows(4).Delete
        objDoc.Tables(3).Rows(8).Delete
    ElseIf Say2 = 15 Then
        objDoc.Tables(3).Rows(2).Delete
        objDoc.Tables(3).Rows(11).Delete
        For i = 1 To 2
            objDoc.Tables(3).Rows(8).Delete
        Next i
        For i = 1 To 2
            objDoc.Tables(3).Rows(4).Delete
            objDoc.Tables(3).Rows(9).Delete
        Next i
        objDoc.Tables(3).Rows(4).Delete
        objDoc.Tables(3).Rows(8).Delete
        objDoc.Tables(3).Rows(4).Delete
        objDoc.Tables(3).Rows(7).Delete
    ElseIf Say2 = 16 Then
        objDoc.Tables(3).Rows(2).Delete
        objDoc.Tables(3).Rows(11).Delete
        For i = 1 To 2
            objDoc.Tables(3).Rows(8).Delete
        Next i
        For i = 1 To 2
            objDoc.Tables(3).Rows(4).Delete
            objDoc.Tables(3).Rows(9).Delete
        Next i
        objDoc.Tables(3).Rows(4).Delete
        objDoc.Tables(3).Rows(8).Delete
        objDoc.Tables(3).Rows(4).Delete
        objDoc.Tables(3).Rows(7).Delete
        objDoc.Tables(3).Rows(11).Delete
        objDoc.Tables(3).Rows(11).Delete
    ElseIf Say2 = 17 Then
        objDoc.Tables(3).Rows(2).Delete
        objDoc.Tables(3).Rows(11).Delete
        For i = 1 To 2
            objDoc.Tables(3).Rows(8).Delete
        Next i
        For i = 1 To 2
            objDoc.Tables(3).Rows(4).Delete
            objDoc.Tables(3).Rows(9).Delete
        Next i
        objDoc.Tables(3).Rows(4).Delete
        objDoc.Tables(3).Rows(8).Delete
        objDoc.Tables(3).Rows(4).Delete
        objDoc.Tables(3).Rows(7).Delete
        objDoc.Tables(3).Rows(11).Delete
        objDoc.Tables(3).Rows(11).Delete
        objDoc.Tables(3).Rows(3).Delete
        objDoc.Tables(3).Rows(5).Delete
    ElseIf Say2 = 18 Then
        objDoc.Tables(3).Rows(2).Delete
        objDoc.Tables(3).Rows(11).Delete
        For i = 1 To 2
            objDoc.Tables(3).Rows(8).Delete
        Next i
        For i = 1 To 2
            objDoc.Tables(3).Rows(4).Delete
            objDoc.Tables(3).Rows(9).Delete
        Next i
        objDoc.Tables(3).Rows(4).Delete
        objDoc.Tables(3).Rows(8).Delete
        objDoc.Tables(3).Rows(4).Delete
        objDoc.Tables(3).Rows(7).Delete
        objDoc.Tables(3).Rows(11).Delete
        objDoc.Tables(3).Rows(11).Delete
        objDoc.Tables(3).Rows(3).Delete
        objDoc.Tables(3).Rows(5).Delete
        objDoc.Tables(3).Rows(7).Delete
        objDoc.Tables(3).Rows(7).Delete
    ElseIf Say2 = 19 Then
        For i = 2 To 20
            objDoc.Tables(2).Rows(i).Height = 28 '27
        Next i
    ElseIf Say2 = 20 Then
        For i = 2 To 21
            objDoc.Tables(2).Rows(i).Height = 28 '27
        Next i
    ElseIf Say2 = 36 Then
        objDoc.Tables(3).Rows(2).Delete
        objDoc.Tables(3).Rows(11).Delete
    ElseIf Say2 = 37 Then
        objDoc.Tables(3).Rows(2).Delete
        objDoc.Tables(3).Rows(11).Delete
        For i = 1 To 2
            objDoc.Tables(3).Rows(8).Delete
        Next i
    ElseIf Say2 = 38 Then
        objDoc.Tables(3).Rows(2).Delete
        objDoc.Tables(3).Rows(11).Delete
        For i = 1 To 2
            objDoc.Tables(3).Rows(8).Delete
        Next i
        objDoc.Tables(3).Rows(4).Delete
        objDoc.Tables(3).Rows(9).Delete
    ElseIf Say2 = 39 Then
        objDoc.Tables(3).Rows(2).Delete
        objDoc.Tables(3).Rows(11).Delete
        For i = 1 To 2
            objDoc.Tables(3).Rows(8).Delete
        Next i
        For i = 1 To 2
            objDoc.Tables(3).Rows(4).Delete
            objDoc.Tables(3).Rows(9).Delete
        Next i
    ElseIf Say2 = 40 Then
        objDoc.Tables(3).Rows(2).Delete
        objDoc.Tables(3).Rows(11).Delete
        For i = 1 To 2
            objDoc.Tables(3).Rows(8).Delete
        Next i
        For i = 1 To 2
            objDoc.Tables(3).Rows(4).Delete
            objDoc.Tables(3).Rows(9).Delete
        Next i
        objDoc.Tables(3).Rows(4).Delete
        objDoc.Tables(3).Rows(8).Delete
    ElseIf Say2 = 41 Then
        objDoc.Tables(3).Rows(2).Delete
        objDoc.Tables(3).Rows(11).Delete
        For i = 1 To 2
            objDoc.Tables(3).Rows(8).Delete
        Next i
        For i = 1 To 2
            objDoc.Tables(3).Rows(4).Delete
            objDoc.Tables(3).Rows(9).Delete
        Next i
        objDoc.Tables(3).Rows(4).Delete
        objDoc.Tables(3).Rows(8).Delete
        objDoc.Tables(3).Rows(4).Delete
        objDoc.Tables(3).Rows(7).Delete
    ElseIf Say2 = 42 Then
        objDoc.Tables(3).Rows(2).Delete
        objDoc.Tables(3).Rows(11).Delete
        For i = 1 To 2
            objDoc.Tables(3).Rows(8).Delete
        Next i
        For i = 1 To 2
            objDoc.Tables(3).Rows(4).Delete
            objDoc.Tables(3).Rows(9).Delete
        Next i
        objDoc.Tables(3).Rows(4).Delete
        objDoc.Tables(3).Rows(8).Delete
        objDoc.Tables(3).Rows(4).Delete
        objDoc.Tables(3).Rows(7).Delete
        objDoc.Tables(3).Rows(11).Delete
        objDoc.Tables(3).Rows(11).Delete
    ElseIf Say2 = 43 Then
        objDoc.Tables(3).Rows(2).Delete
        objDoc.Tables(3).Rows(11).Delete
        For i = 1 To 2
            objDoc.Tables(3).Rows(8).Delete
        Next i
        For i = 1 To 2
            objDoc.Tables(3).Rows(4).Delete
            objDoc.Tables(3).Rows(9).Delete
        Next i
        objDoc.Tables(3).Rows(4).Delete
        objDoc.Tables(3).Rows(8).Delete
        objDoc.Tables(3).Rows(4).Delete
        objDoc.Tables(3).Rows(7).Delete
        objDoc.Tables(3).Rows(11).Delete
        objDoc.Tables(3).Rows(11).Delete
        objDoc.Tables(3).Rows(3).Delete
        objDoc.Tables(3).Rows(5).Delete
    ElseIf Say2 = 44 Then
        objDoc.Tables(3).Rows(2).Delete
        objDoc.Tables(3).Rows(11).Delete
        For i = 1 To 2
            objDoc.Tables(3).Rows(8).Delete
        Next i
        For i = 1 To 2
            objDoc.Tables(3).Rows(4).Delete
            objDoc.Tables(3).Rows(9).Delete
        Next i
        objDoc.Tables(3).Rows(4).Delete
        objDoc.Tables(3).Rows(8).Delete
        objDoc.Tables(3).Rows(4).Delete
        objDoc.Tables(3).Rows(7).Delete
        objDoc.Tables(3).Rows(11).Delete
        objDoc.Tables(3).Rows(11).Delete
        objDoc.Tables(3).Rows(3).Delete
        objDoc.Tables(3).Rows(5).Delete
        objDoc.Tables(3).Rows(7).Delete
        objDoc.Tables(3).Rows(7).Delete
    ElseIf Say2 = 45 Then
        For i = 23 To 46
            objDoc.Tables(2).Rows(i).Height = 28 '27
        Next i
    ElseIf Say2 = 46 Then
        For i = 36 To 47
            objDoc.Tables(2).Rows(i).Height = 28 '27
        Next i
    ElseIf Say2 = 47 Then
        For i = 43 To 48
            objDoc.Tables(2).Rows(i).Height = 28 '27
        Next i
    End If
ElseIf Tip4Option.Value = True Then

'    'İmza boşluğunu sayfaya sığdırmak için düzenle
    If Say2 = 13 Then
        objDoc.Tables(3).Rows(2).Delete
        objDoc.Tables(3).Rows(11).Delete
    ElseIf Say2 = 14 Then
        objDoc.Tables(3).Rows(2).Delete
        objDoc.Tables(3).Rows(11).Delete
        For i = 1 To 2
            objDoc.Tables(3).Rows(8).Delete
        Next i
    ElseIf Say2 = 15 Then
        objDoc.Tables(3).Rows(2).Delete
        objDoc.Tables(3).Rows(11).Delete
        For i = 1 To 2
            objDoc.Tables(3).Rows(8).Delete
        Next i
        objDoc.Tables(3).Rows(4).Delete
        objDoc.Tables(3).Rows(9).Delete
    ElseIf Say2 = 16 Then
        objDoc.Tables(3).Rows(2).Delete
        objDoc.Tables(3).Rows(11).Delete
        For i = 1 To 2
            objDoc.Tables(3).Rows(8).Delete
        Next i
        For i = 1 To 2
            objDoc.Tables(3).Rows(4).Delete
            objDoc.Tables(3).Rows(9).Delete
        Next i
    ElseIf Say2 = 17 Then
        objDoc.Tables(3).Rows(2).Delete
        objDoc.Tables(3).Rows(11).Delete
        For i = 1 To 2
            objDoc.Tables(3).Rows(8).Delete
        Next i
        For i = 1 To 2
            objDoc.Tables(3).Rows(4).Delete
            objDoc.Tables(3).Rows(9).Delete
        Next i
        objDoc.Tables(3).Rows(4).Delete
        objDoc.Tables(3).Rows(8).Delete
    ElseIf Say2 = 18 Then
        objDoc.Tables(3).Rows(2).Delete
        objDoc.Tables(3).Rows(11).Delete
        For i = 1 To 2
            objDoc.Tables(3).Rows(8).Delete
        Next i
        For i = 1 To 2
            objDoc.Tables(3).Rows(4).Delete
            objDoc.Tables(3).Rows(9).Delete
        Next i
        objDoc.Tables(3).Rows(4).Delete
        objDoc.Tables(3).Rows(8).Delete
        objDoc.Tables(3).Rows(4).Delete
        objDoc.Tables(3).Rows(7).Delete
    ElseIf Say2 = 19 Then
        objDoc.Tables(3).Rows(2).Delete
        objDoc.Tables(3).Rows(11).Delete
        For i = 1 To 2
            objDoc.Tables(3).Rows(8).Delete
        Next i
        For i = 1 To 2
            objDoc.Tables(3).Rows(4).Delete
            objDoc.Tables(3).Rows(9).Delete
        Next i
        objDoc.Tables(3).Rows(4).Delete
        objDoc.Tables(3).Rows(8).Delete
        objDoc.Tables(3).Rows(4).Delete
        objDoc.Tables(3).Rows(7).Delete
        objDoc.Tables(3).Rows(3).Delete
        objDoc.Tables(3).Rows(5).Delete
    ElseIf Say2 = 20 Then
        For i = 2 To 21
            objDoc.Tables(2).Rows(i).Height = 27
        Next i
    ElseIf Say2 = 39 Then
        objDoc.Tables(3).Rows(2).Delete
        objDoc.Tables(3).Rows(11).Delete
    ElseIf Say2 = 40 Then
        objDoc.Tables(3).Rows(2).Delete
        objDoc.Tables(3).Rows(11).Delete
        For i = 1 To 2
            objDoc.Tables(3).Rows(8).Delete
        Next i
    ElseIf Say2 = 41 Then
        objDoc.Tables(3).Rows(2).Delete
        objDoc.Tables(3).Rows(11).Delete
        For i = 1 To 2
            objDoc.Tables(3).Rows(8).Delete
        Next i
        objDoc.Tables(3).Rows(4).Delete
        objDoc.Tables(3).Rows(9).Delete
    ElseIf Say2 = 42 Then
        objDoc.Tables(3).Rows(2).Delete
        objDoc.Tables(3).Rows(11).Delete
        For i = 1 To 2
            objDoc.Tables(3).Rows(8).Delete
        Next i
        For i = 1 To 2
            objDoc.Tables(3).Rows(4).Delete
            objDoc.Tables(3).Rows(9).Delete
        Next i
    ElseIf Say2 = 43 Then
        objDoc.Tables(3).Rows(2).Delete
        objDoc.Tables(3).Rows(11).Delete
        For i = 1 To 2
            objDoc.Tables(3).Rows(8).Delete
        Next i
        For i = 1 To 2
            objDoc.Tables(3).Rows(4).Delete
            objDoc.Tables(3).Rows(9).Delete
        Next i
        objDoc.Tables(3).Rows(4).Delete
        objDoc.Tables(3).Rows(8).Delete
    ElseIf Say2 = 44 Then
        objDoc.Tables(3).Rows(2).Delete
        objDoc.Tables(3).Rows(11).Delete
        For i = 1 To 2
            objDoc.Tables(3).Rows(8).Delete
        Next i
        For i = 1 To 2
            objDoc.Tables(3).Rows(4).Delete
            objDoc.Tables(3).Rows(9).Delete
        Next i
        objDoc.Tables(3).Rows(4).Delete
        objDoc.Tables(3).Rows(8).Delete
        objDoc.Tables(3).Rows(4).Delete
        objDoc.Tables(3).Rows(7).Delete
    ElseIf Say2 = 45 Then
        objDoc.Tables(3).Rows(2).Delete
        objDoc.Tables(3).Rows(11).Delete
        For i = 1 To 2
            objDoc.Tables(3).Rows(8).Delete
        Next i
        For i = 1 To 2
            objDoc.Tables(3).Rows(4).Delete
            objDoc.Tables(3).Rows(9).Delete
        Next i
        objDoc.Tables(3).Rows(4).Delete
        objDoc.Tables(3).Rows(8).Delete
        objDoc.Tables(3).Rows(4).Delete
        objDoc.Tables(3).Rows(7).Delete
        objDoc.Tables(3).Rows(3).Delete
        objDoc.Tables(3).Rows(5).Delete
    ElseIf Say2 = 46 Then
        For i = 23 To 47
            objDoc.Tables(2).Rows(i).Height = 27 '27
        Next i
    ElseIf Say2 = 47 Then
        For i = 33 To 48
            objDoc.Tables(2).Rows(i).Height = 27 '27
        Next i
    End If
End If


'Direkt yazdır
If TeslimTutanaklariFormuDirektYazdir = True Then
    objWord.Activate
'    For i = 1 To TeslimTutanaklariFormuSayPrt
'        objDoc.PrintOut
'    Next i
    objDoc.PrintOut Background:=False, Copies:=TeslimTutanaklariFormuSayPrt
    'objWord.Documents.Save
    objDoc.Close SaveChanges:=False
    objWord.Visible = False
End If

'Tip4 oluşturulduğunda Tip5 da oluşturulsun.
If Tip4Option.Value = True Then
    'MsgBox "Test"
    'Dosyayı şablondan operasyon klasörüne kopyala ve adını değiştir.
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.CopyFile (SourceTip5), DestOpUserFolder & ReNameTip5 & ".docm", True
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
    objWord.Documents.Open FileName:=DestOpUserFolder & ReNameTip5 & ".docm"
    objWord.Visible = True
    objWord.Activate 'Ekrana getirir.
    'objDoc.Activate 'Ekrana getirmez.
    'objWord.Application.WindowState = wdWindowStateMaximize
    Set objDoc = GetObject(DestOpUserFolder & ReNameTip5 & ".docm")
    '________________________________________
    
    'Birim
    Birimx = UCase(Replace(Replace(Worksheets(2).Cells(6, 99).Value, "i", "I"), "ı", "I")) & " UNIT"
    objDoc.Tables(1).Cell(Row:=1, Column:=1).Range.Text = Birimx
    'Tutanak tarihi
    objDoc.Tables(1).Cell(Row:=4, Column:=3).Range.Text = TutanakTarihiText.Value

'________________AKTARIMLAR 2 (Başlangıç)

    Say2 = 0
    AdetSay = 0
    For Each ctl In core_delivery_manager_UI.Frame2.Controls '________________RAPOR1 BÖLÜMÜ (Başlangıç, Tip5)
        If TypeName(ctl) = "ListBox" Then
            'MsgBox Mid(Ctl.List(0), InStr(Ctl.List(0), "|") + 1, 2)
            If Mid(ctl.List(0), InStr(ctl.List(0), "|") + 1, 2) = " Ö" Then
                
                Call Rapor1GidenTema
                If IlkSiraGlobal = 0 Then
                    GoTo Rapor1Devam2
                End If
                
                IlkSira = IlkSiraGlobal
                GidenTema = GidenTemaGlobal
    
                'Tabloya satır ekle
                Say2 = Say2 + 1
                'MsgBox Left(Ctl.List(0), InStr(Ctl.List(0), "|") - 2)
                If Say2 > 1 Then
                    objDoc.Tables(2).Rows.Add BeforeRow:=objDoc.Tables(2).Rows(Say2 + 1)
                End If
                'Sıra no
                objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=1).Range.Text = Say2
                'Gönderilen birim
                objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=2).Range.Text = GidenTema
                'Çıkış kayıt no
                objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=3).Range.Text = ThisWorkbook.Worksheets(3).Cells(IlkSira, 76).Value
                'Package A/Package B/Package C
                objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=4).Range.Text = ThisWorkbook.Worksheets(3).Cells(IlkSira, 68).Value & " unit(s) of " & Left(ThisWorkbook.Worksheets(3).Cells(IlkSira, 67).Value, InStr(ThisWorkbook.Worksheets(3).Cells(IlkSira, 67).Value, "/") - 1)
                AdetSay = AdetSay + ThisWorkbook.Worksheets(3).Cells(IlkSira, 68).Value
            End If
        End If
Rapor1Devam2:
    Next ctl '________________RAPOR1 BÖLÜMÜ (Bitiş, Tip5)
    
    For Each ctl In core_delivery_manager_UI.Frame2.Controls '________________RAPOR BÖLÜMÜ (Başlangıç, Tip5) 'Rapor ve bilgilendirme için ortak
        If TypeName(ctl) = "ListBox" Then
            'MsgBox Mid(Ctl.List(0), InStr(Ctl.List(0), "|") + 1, 2)
            If Mid(ctl.List(0), InStr(ctl.List(0), "|") + 1, 2) = " R" Then

                Call RaporGidenTema
                If IlkSiraGlobal = 0 Then
                    GoTo RaporDevam2
                End If
                
                IlkSira = IlkSiraGlobal
                GidenTema = GidenTemaGlobal
    
                'Tabloya satır ekle
                Say2 = Say2 + 1
                'MsgBox Left(Ctl.List(0), InStr(Ctl.List(0), "|") - 2)
                If Say2 > 1 Then
                    objDoc.Tables(2).Rows.Add BeforeRow:=objDoc.Tables(2).Rows(Say2 + 1)
                End If
                'Sıra no
                objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=1).Range.Text = Say2
                'Gönderilen birim
                objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=2).Range.Text = GidenTema
                'Çıkış kayıt no
                objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=3).Range.Text = ThisWorkbook.Worksheets(4).Cells(IlkSira, 84).Value

                If ThisWorkbook.Worksheets(4).Cells(IlkSiraGlobal, 174).Value = "Yes" Then 'Bilgilendirme (Rapor2_2 evet)
                    'Package A/Package B/Package C
                    objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=4).Range.Text = 1 & " unit(s) of " & Left(ThisWorkbook.Worksheets(4).Cells(IlkSira, 81).Value, InStr(ThisWorkbook.Worksheets(4).Cells(IlkSira, 81).Value, "/") - 1)
                    AdetSay = AdetSay + 1
                Else
                    'Package A/Package B/Package C
                    objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=4).Range.Text = ThisWorkbook.Worksheets(4).Cells(IlkSira, 76).Value & " unit(s) of " & Left(ThisWorkbook.Worksheets(4).Cells(IlkSira, 75).Value, InStr(ThisWorkbook.Worksheets(4).Cells(IlkSira, 75).Value, "/") - 1)
                    AdetSay = AdetSay + ThisWorkbook.Worksheets(4).Cells(IlkSira, 76).Value
                End If
            
            End If
        End If
RaporDevam2:
    Next ctl '________________RAPOR BÖLÜMÜ (Bitiş, Tip5)

    For Each ctl In core_delivery_manager_UI.Frame2.Controls '________________RAPOR BÖLÜMÜ (Başlangıç, Tip5) 'XXXMud
        If TypeName(ctl) = "ListBox" Then
            'MsgBox Mid(Ctl.List(0), InStr(Ctl.List(0), "|") + 1, 2)
            If Mid(ctl.List(0), InStr(ctl.List(0), "|") + 1, 2) = " B" And InStr(ctl.List(0), "XXXMud") <> 0 Then

                Call RaporGidenTema
                If IlkSiraGlobal = 0 Then
                    GoTo RaporDevam2XXXMud
                End If
                
                IlkSira = IlkSiraGlobal
                GidenTema = GidenTemaGlobal
    
                'Tabloya satır ekle
                Say2 = Say2 + 1
                'MsgBox Left(Ctl.List(0), InStr(Ctl.List(0), "|") - 2)
                If Say2 > 1 Then
                    objDoc.Tables(2).Rows.Add BeforeRow:=objDoc.Tables(2).Rows(Say2 + 1)
                End If
                'Sıra no
                objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=1).Range.Text = Say2
                'Gönderilen birim
                GidenTema = "ORGANIZATION A XXX Directorate"
                objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=2).Range.Text = GidenTema
                'Çıkış kayıt no
                objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=3).Range.Text = ThisWorkbook.Worksheets(4).Cells(IlkSira, 176).Value
                'Package A/Package B/Package C
                objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=4).Range.Text = ThisWorkbook.Worksheets(4).Cells(IlkSira, 185).Value & " unit(s) of " & Left(ThisWorkbook.Worksheets(4).Cells(IlkSira, 184).Value, InStr(ThisWorkbook.Worksheets(4).Cells(IlkSira, 184).Value, "/") - 1)
                AdetSay = AdetSay + ThisWorkbook.Worksheets(4).Cells(IlkSira, 185).Value
            End If
        End If
RaporDevam2XXXMud:
    Next ctl '________________RAPOR BÖLÜMÜ (Bitiş, Tip5)
    
    For Each ctl In core_delivery_manager_UI.Frame2.Controls '________________RAPOR BÖLÜMÜ (Başlangıç, Tip5) 'Sonuç
        If TypeName(ctl) = "ListBox" Then
            'MsgBox Mid(Ctl.List(0), InStr(Ctl.List(0), "|") + 1, 2)
            If Mid(ctl.List(0), InStr(ctl.List(0), "|") + 1, 2) = " B" And InStr(ctl.List(0), "XXXMud") = 0 Then

                Call Rapor2_2GidenTema
                If IlkSiraGlobal = 0 Then
                    GoTo RaporDevam2Sonuc
                End If
                
                IlkSira = IlkSiraGlobal
                GidenTema = GidenTemaGlobal
    
                'Tabloya satır ekle
                Say2 = Say2 + 1
                'MsgBox Left(Ctl.List(0), InStr(Ctl.List(0), "|") - 2)
                If Say2 > 1 Then
                    objDoc.Tables(2).Rows.Add BeforeRow:=objDoc.Tables(2).Rows(Say2 + 1)
                End If
                'Sıra no
                objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=1).Range.Text = Say2
                'Gönderilen birim
                objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=2).Range.Text = GidenTema
                'Çıkış kayıt no
                objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=3).Range.Text = ThisWorkbook.Worksheets(4).Cells(IlkSira, 178).Value
                
                'Senaryo düzenleyicisi (veriler  kurum tutanak2sından mı, XXXMud tutanak2sından mı yok sa her ikisinden mi alınacak)
                If ThisWorkbook.Worksheets(4).Cells(IlkSira, 187).Value = "All" Then 'XXXMudya giden
                    If ThisWorkbook.Worksheets(4).Cells(IlkSira, 188).Value = "Yes" Then 'XXXMud'den gelen paketin açılma durumu
                        'Kurum
                        GidenPaketAdet = Application.Sum(ThisWorkbook.Worksheets(4).Range(ThisWorkbook.Worksheets(4).Cells(IlkSira, 76), ThisWorkbook.Worksheets(4).Cells(IlkSira, 76)))
                        'Package A/Package B/Package C
                        objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=4).Range.Text = GidenPaketAdet & " unit(s) of " & Left(ThisWorkbook.Worksheets(4).Cells(IlkSira, 75).Value, InStr(ThisWorkbook.Worksheets(4).Cells(IlkSira, 75).Value, "/") - 1)
                        AdetSay = AdetSay + GidenPaketAdet
                    ElseIf ThisWorkbook.Worksheets(4).Cells(IlkSira, 188).Value = "No" Then 'XXXMud'den gelen paketin açılma durumu
                        'XXXMud
                        GidenPaketAdet = Application.Sum(ThisWorkbook.Worksheets(4).Range(ThisWorkbook.Worksheets(4).Cells(IlkSira, 202), ThisWorkbook.Worksheets(4).Cells(IlkSira, 202)))
                        'Package A/Package B/Package C
                        objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=4).Range.Text = GidenPaketAdet & " unit(s) of " & Left(ThisWorkbook.Worksheets(4).Cells(IlkSira, 201).Value, InStr(ThisWorkbook.Worksheets(4).Cells(IlkSira, 201).Value, "/") - 1)
                        AdetSay = AdetSay + GidenPaketAdet
                    End If
                ElseIf ThisWorkbook.Worksheets(4).Cells(IlkSira, 187).Value = "Technique A" Then 'XXXMudya giden
                    If ThisWorkbook.Worksheets(4).Cells(IlkSira, 188).Value = "Yes" Then 'XXXMud'den gelen paketin açılma durumu
                        If ThisWorkbook.Worksheets(4).Cells(IlkSira, 189).Value = "Yes" Then 'Tutanak2 tut. birleşme durumu
                            'Kurum
                            GidenPaketAdet = Application.Sum(ThisWorkbook.Worksheets(4).Range(ThisWorkbook.Worksheets(4).Cells(IlkSira, 76), ThisWorkbook.Worksheets(4).Cells(IlkSira, 76)))
                            'Package A/Package B/Paket C
                            objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=4).Range.Text = GidenPaketAdet & " unit(s) of " & Left(ThisWorkbook.Worksheets(4).Cells(IlkSira, 75).Value, InStr(ThisWorkbook.Worksheets(4).Cells(IlkSira, 75).Value, "/") - 1)
                            AdetSay = AdetSay + GidenPaketAdet
                        ElseIf ThisWorkbook.Worksheets(4).Cells(IlkSira, 189).Value = "No" Then 'Tutanak2 tut. birleşme durumu
                            'Kurum ve XXXMud
                            GidenPaketAdet = Application.Sum(ThisWorkbook.Worksheets(4).Range(ThisWorkbook.Worksheets(4).Cells(IlkSira, 76), ThisWorkbook.Worksheets(4).Cells(IlkSira, 76)))
                            GidenPaketAdet = GidenPaketAdet + Application.Sum(ThisWorkbook.Worksheets(4).Range(ThisWorkbook.Worksheets(4).Cells(IlkSira, 202), ThisWorkbook.Worksheets(4).Cells(IlkSira, 202)))
                            'Package A/Package B/Package C
                            objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=4).Range.Text = GidenPaketAdet & " unit(s) of " & Left(ThisWorkbook.Worksheets(4).Cells(IlkSira, 75).Value, InStr(ThisWorkbook.Worksheets(4).Cells(IlkSira, 75).Value, "/") - 1)
                            AdetSay = AdetSay + GidenPaketAdet
                        End If
                    ElseIf ThisWorkbook.Worksheets(4).Cells(IlkSira, 188).Value = "No" Then 'XXXMud'den gelen paketin açılma durumu
                        'Kurum
                        GidenPaketAdet = Application.Sum(ThisWorkbook.Worksheets(4).Range(ThisWorkbook.Worksheets(4).Cells(IlkSira, 76), ThisWorkbook.Worksheets(4).Cells(IlkSira, 76)))
                        'Package A/Package B/Package C
                        objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=4).Range.Text = GidenPaketAdet & " unit(s) of " & Left(ThisWorkbook.Worksheets(4).Cells(IlkSira, 75).Value, InStr(ThisWorkbook.Worksheets(4).Cells(IlkSira, 75).Value, "/") - 1)
                        AdetSay = AdetSay + GidenPaketAdet
                    End If
                End If

            End If
        End If
RaporDevam2Sonuc:
    Next ctl '________________RAPOR BÖLÜMÜ (Bitiş, Tip5)


    For Each ctl In core_delivery_manager_UI.Frame2.Controls '________________RAPOR3_1 BÖLÜMÜ (Başlangıç, Tip5)
        If TypeName(ctl) = "ListBox" Then
            'MsgBox Mid(Ctl.List(0), InStr(Ctl.List(0), "|") + 1, 2)
            If Mid(ctl.List(0), InStr(ctl.List(0), "|") + 1, 2) = " G" Then

                Call Rapor3_1GidenTema
                If IlkSiraGlobal = 0 Then
                    GoTo Rapor3_1Devam2
                End If
                
                IlkSira = IlkSiraGlobal
                GidenTema = GidenTemaGlobal
    
                'Tabloya satır ekle
                Say2 = Say2 + 1
                'MsgBox Left(Ctl.List(0), InStr(Ctl.List(0), "|") - 2)
                If Say2 > 1 Then
                    objDoc.Tables(2).Rows.Add BeforeRow:=objDoc.Tables(2).Rows(Say2 + 1)
                End If
                'Sıra no
                objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=1).Range.Text = Say2
                'Gönderilen birim
                objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=2).Range.Text = GidenTema
                'Çıkış kayıt no
                objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=3).Range.Text = ThisWorkbook.Worksheets(5).Cells(IlkSira, 156).Value
                'Package A/Package B/Package C
                objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=4).Range.Text = ThisWorkbook.Worksheets(5).Cells(IlkSira, 150).Value & " unit(s) of " & Left(ThisWorkbook.Worksheets(5).Cells(IlkSira, 149).Value, InStr(ThisWorkbook.Worksheets(5).Cells(IlkSira, 149).Value, "/") - 1)
                AdetSay = AdetSay + ThisWorkbook.Worksheets(5).Cells(IlkSira, 150).Value
            End If
        End If
Rapor3_1Devam2:
    Next ctl '________________RAPOR3_1 BÖLÜMÜ (Bitiş, Tip5)


    For Each ctl In core_delivery_manager_UI.Frame2.Controls '________________RAPOR3_2 BÖLÜMÜ (Başlangıç, Tip5)
        If TypeName(ctl) = "ListBox" Then
            'MsgBox Mid(Ctl.List(0), InStr(Ctl.List(0), "|") + 1, 2)
            If Mid(ctl.List(0), InStr(ctl.List(0), "|") + 1, 2) = " M" Then  'Mid(ctl.List(0), InStrRev(ctl.List(0), "-") - 5, 5) <> "FinansalBirim" Then
            
                'If Mid(ctl.List(0), InStrRev(ctl.List(0), "-") - 5, 5) = "FinansalBirim" Then 'FinansalBirim
                If InStr(ctl.List(0), "(FinansalBirim-TipA)") <> 0 Or InStr(ctl.List(0), "(FinansalBirim-TipB)") <> 0 Then

                    Call Rapor3_2GidenTema
                    If IlkSiraGlobal = 0 Then
                        GoTo Rapor3_2Devam2
                    End If
        
                    IlkSira = IlkSiraGlobal
                    GidenTema = ThisWorkbook.Worksheets(5).Cells(IlkSira, 30).Value & " " & ThisWorkbook.Worksheets(5).Cells(IlkSira, 31).Value 'GidenTemaGlobal
                    'Tabloya satır ekle
                    Say2 = Say2 + 1
                    'MsgBox Left(Ctl.List(0), InStr(Ctl.List(0), "|") - 2)
                    If Say2 > 1 Then
                        objDoc.Tables(2).Rows.Add BeforeRow:=objDoc.Tables(2).Rows(Say2 + 1)
                    End If
                    'Sıra no
                    objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=1).Range.Text = Say2
        
                    'Gönderilen birim
                    objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=2).Range.Text = GidenTema
                    'Çıkış kayıt no
                    objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=3).Range.Text = ThisWorkbook.Worksheets(5).Cells(IlkSira, 76).Value
        
                    'Package A/Package B/Package C
                    objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=4).Range.Text = "1 Adet Package A"
                    AdetSay = AdetSay + 1
                    If Tip2Option.Value = True Then
                        'Gönderim şekli
                        objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=5).Range.Text = ThisWorkbook.Worksheets(5).Cells(IlkSira, 85).Value
                    End If
                    
                Else 'DIRECTORATELIK
                
                    'InStr(Mid(ctl.List(0), InStrRev(ctl.List(0), "-") - 5, 5), "FinansalBirim") = 0
                    Call Rapor3_2GidenTema
                    If IlkSiraGlobal = 0 Then
                        GoTo Rapor3_2Devam2
                    End If
                    
                    IlkSira = IlkSiraGlobal
                    GidenTema = GidenTemaGlobal
        
                    'Tabloya satır ekle
                    Say2 = Say2 + 1
                    'MsgBox Left(Ctl.List(0), InStr(Ctl.List(0), "|") - 2)
                    If Say2 > 1 Then
                        objDoc.Tables(2).Rows.Add BeforeRow:=objDoc.Tables(2).Rows(Say2 + 1)
                    End If
                    'Sıra no
                    objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=1).Range.Text = Say2
                    'Gönderilen birim
                    objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=2).Range.Text = GidenTema
                    'Çıkış kayıt no
                    objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=3).Range.Text = ThisWorkbook.Worksheets(5).Cells(IlkSira, 84).Value
                    'Package A/Package B/Package C
                    objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=4).Range.Text = ThisWorkbook.Worksheets(5).Cells(IlkSira, 72).Value & " unit(s) of " & Left(ThisWorkbook.Worksheets(5).Cells(IlkSira, 71).Value, InStr(ThisWorkbook.Worksheets(5).Cells(IlkSira, 71).Value, "/") - 1)
                    AdetSay = AdetSay + ThisWorkbook.Worksheets(5).Cells(IlkSira, 72).Value
                
                End If
            End If
        End If
Rapor3_2Devam2:
    Next ctl '________________RAPOR3_2 BÖLÜMÜ (Bitiş, Tip5)
    
    
'________________AKTARIMLAR 2 (Bitiş)


    'Tip5 metin
    objDoc.Tables(3).Cell(Row:=1, Column:=1).Range.Text = "Delivery of the " & AdetSay & " " & Ek1 & " mentioned above (Package A/B/C) has been completed."
    'Adedi bold yap.
    Set MyRange = objDoc.Tables(3).Cell(Row:=1, Column:=1).Range
    With MyRange.Find
        .Execute FindText:=AdetSay
    End With
    MyRange.Font.Bold = True

    'Alt bilgi ekle
    objDoc.Sections(1).Footers(wdHeaderFooterPrimary).Range.Tables(1).Cell(Row:=1, Column:=1).Range.Text = ReNameTip5
    
    'Artık satırı sil
    objDoc.Tables(2).Rows(Say2 + 2).Delete
    'İmza boşluğunu sayfaya sığdırmak için düzenle
    If Say2 = 19 Then
        objDoc.Tables(4).Rows(1).Delete
        objDoc.Tables(4).Rows(2).Delete
    ElseIf Say2 = 20 Then
        objDoc.Tables(4).Rows(1).Delete
        objDoc.Tables(4).Rows(2).Delete
        For i = 1 To 2
            objDoc.Tables(4).Rows(3).Delete
        Next i
    ElseIf Say2 = 21 Then
        For i = 2 To 22
            objDoc.Tables(2).Rows(i).Height = 27
        Next i
    ElseIf Say2 = 45 Then
        objDoc.Tables(4).Rows(1).Delete
        objDoc.Tables(4).Rows(2).Delete
    ElseIf Say2 = 46 Then
        objDoc.Tables(4).Rows(1).Delete
        objDoc.Tables(4).Rows(2).Delete
        For i = 1 To 2
            objDoc.Tables(4).Rows(3).Delete
        Next i
    ElseIf Say2 = 47 Then
        For i = 24 To 48
            objDoc.Tables(2).Rows(i).Height = 27
        Next i
    ElseIf Say2 = 48 Then
        For i = 44 To 49
            objDoc.Tables(2).Rows(i).Height = 27
        Next i
    End If

    'Direkt yazdır
    If TeslimTutanaklariFormuDirektYazdir = True Then
        objWord.Activate
'        For i = 1 To TeslimTutanaklariFormuSayPrt
'            objDoc.PrintOut
'        Next i
        objDoc.PrintOut Background:=False, Copies:=TeslimTutanaklariFormuSayPrt
        'objWord.Documents.Save
        objDoc.Close SaveChanges:=False
        objWord.Visible = False
    End If

End If

    'Tutanak numarasını kaydet
    ThisWorkbook.Worksheets(6).Unprotect Password:="123"
    'Tip2
    If Tip2Option.Value = True Then
        'Comboda tanımlı değer ise
        a() = TutanakNoText.List
        For b = LBound(a) To UBound(a)
            If a(b, 0) = TutanakNoText.Value Then
                GoTo Tip2NoEklemedenAtla
            End If
        Next b
        Say = ThisWorkbook.Worksheets(6).Range("B100000").End(xlUp).Row + 1
        If Say < 3 Then
            Say = 3
        End If
        ThisWorkbook.Worksheets(6).Cells(Say, 2) = TutanakNoText.Value
    End If
Tip2NoEklemedenAtla:

    'Ek 181
    If Tip3Option.Value = True Then
        'Comboda tanımlı değer ise
        a() = TutanakNoText.List
        For b = LBound(a) To UBound(a)
            If a(b, 0) = TutanakNoText.Value Then
                GoTo Tip3NoEklemedenAtla
            End If
        Next b
        Say = ThisWorkbook.Worksheets(6).Range("C100000").End(xlUp).Row + 1
        If Say < 3 Then
            Say = 3
        End If
        ThisWorkbook.Worksheets(6).Cells(Say, 3) = TutanakNoText.Value
    End If
Tip3NoEklemedenAtla:

    'Ek 182
    If Tip4Option.Value = True Then
        'Comboda tanımlı değer ise
        a() = TutanakNoText.List
        For b = LBound(a) To UBound(a)
            If a(b, 0) = TutanakNoText.Value Then
                GoTo Tip4NoEklemedenAtla
            End If
        Next b
        Say = ThisWorkbook.Worksheets(6).Range("D100000").End(xlUp).Row + 1
        If Say < 3 Then
            Say = 3
        End If
        ThisWorkbook.Worksheets(6).Cells(Say, 4) = TutanakNoText.Value
    End If
Tip4NoEklemedenAtla:

    ThisWorkbook.Worksheets(6).Protect Password:="123" ', DrawingObjects:=False
    
    'TutanakNoText.Value = ""

Call ModuleSystemSettings.DropDownKapat

Son:

    Call SonTutanakNoTeslim

    'Nesneleri temizle
    Set fso = Nothing
    Set objWord = Nothing
    Set objDoc = Nothing

    TeslimTutanaklariFormuDirektYazdir = False

    'Columns("CK:CN").EntireColumn.Hidden = True
    'Columns("CE:CF").EntireColumn.Hidden = True
    
    Worksheets(3).Protect Password:="123" ', DrawingObjects:=False
    Worksheets(4).Protect Password:="123" ', DrawingObjects:=False
    Worksheets(5).Protect Password:="123" ', DrawingObjects:=False
    ThisWorkbook.Protect "123"
    
    'Worksheets(3).Activate
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    'Application.EnableEvents = True

Tip1Son:

End Sub

Private Sub TutanakTarihiText_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

'suppport_calendar_UI
suppport_calendar_UI.Show

If CalTarih = "" Then
    GoTo Son
Else
    TutanakTarihiText.Value = CalTarih
    TutanakTarihiText.Value = Format(TutanakTarihiText.Value, "dd.mm.yyyy")
End If

Son:
CalTarih = ""
Cancel = True

End Sub

Private Sub TutanakTarihiText_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    
    On Error Resume Next
    'Delete ve Backspace tuşları textboxu sil.
    If KeyCode = vbKeyDelete Then
        TutanakTarihiText.Value = ""
    End If
    If KeyCode = vbKeyBack Then
        TutanakTarihiText.Value = ""
    End If

End Sub

Private Sub TutanakTarihiLabel_Click()

'suppport_calendar_UI
suppport_calendar_UI.Show

If CalTarih = "" Then
    GoTo Son
Else
    TutanakTarihiText.Value = CalTarih
    TutanakTarihiText.Value = Format(TutanakTarihiText.Value, "dd.mm.yyyy")
End If

Son:
CalTarih = ""

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

For Each ClrLab In core_delivery_manager_UI.Controls
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

TakvimBtn.BackColor = RGB(225, 235, 245)
TakvimBtn.ForeColor = RGB(30, 30, 30)
Kapat.BackColor = RGB(225, 235, 245)
Kapat.ForeColor = RGB(30, 30, 30)
Yardim.BackColor = RGB(225, 235, 245)
Yardim.ForeColor = RGB(30, 30, 30)
Tutanak.BackColor = RGB(225, 235, 245)
Tutanak.ForeColor = RGB(30, 30, 30)
DirektYazdir.BackColor = RGB(225, 235, 245)
DirektYazdir.ForeColor = RGB(30, 30, 30)
core_delivery_manager_UI.BackColor = RGB(230, 230, 230) 'YENİ

LblSiraNo1.BorderColor = RGB(180, 180, 180)
LblBelgeNo1.BorderColor = RGB(180, 180, 180)
LblGonderen1.BorderColor = RGB(180, 180, 180)
LblSiraNo2.BorderColor = RGB(180, 180, 180)
LblBelgeNo2.BorderColor = RGB(180, 180, 180)
LblGonderen2.BorderColor = RGB(180, 180, 180)

TutanakNoFrame.Visible = False 'Açılışta tutanak no olmasın

TasiyiciFrame.Height = 514
Frame1x.Height = 264
Frame2x.Height = 264

Frame1x.ZOrder msoBringToFront
Frame2x.ZOrder msoBringToFront
Frame1.ZOrder msoBringToFront
Frame2.ZOrder msoBringToFront
FrameAlt1.ZOrder msoBringToFront
FrameAlt2.ZOrder msoBringToFront

TeslimTutanaklariFormuDirektYazdir = False


End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
Dim yukseklik As Variant, genislik As Variant
Dim Rep As Variant


RemoveScrollHook


yukseklik = Me.Height
Rep = Me.Width
Do
DoEvents
Rep = Rep - 70
Call timeout(0.01)
    If Rep > 70 Then
        core_delivery_manager_UI.Width = Rep
        yukseklik = yukseklik - 70
        core_delivery_manager_UI.Height = yukseklik
        If yukseklik <= 70 Then
            yukseklik = 70
            core_delivery_manager_UI.Height = yukseklik
        End If
    ElseIf Rep <= 50 Then
        Rep = 50
        core_delivery_manager_UI.Width = Rep
        yukseklik = yukseklik - 50
        core_delivery_manager_UI.Height = yukseklik
        If yukseklik <= 50 Then
            yukseklik = 50
            core_delivery_manager_UI.Height = yukseklik
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

Private Sub SonTutanakNoTeslim()
Dim i As Integer
Dim Say As Long, j As Long, Cont As Long, Tno As Variant

Application.ScreenUpdating = False

'Tutanak numarasını getir
ThisWorkbook.Unprotect "123"
ThisWorkbook.Worksheets(6).Unprotect Password:="123"

'Tip2
If Tip2Option.Value = True Then
    Say = ThisWorkbook.Worksheets(6).Range("B100000").End(xlUp).Row
    If Say < 3 Then
        Say = 3
    End If
    Cont = 0
    TutanakNoText.Clear
    For j = Say To 3 Step -1
        If Cont < 5 Then
            If ThisWorkbook.Worksheets(6).Cells(j, 2).Value <> "" Then
                Cont = Cont + 1
                Tno = ThisWorkbook.Worksheets(6).Cells(j, 2).Value
                With TutanakNoText
                    .AddItem (Tno)
                End With
            End If
        Else
            GoTo DonguSon1
        End If
    Next j
DonguSon1:

'    If ThisWorkbook.Worksheets(6).Cells(Say, 2).Value <> "" And IsNumeric(ThisWorkbook.Worksheets(6).Cells(Say, 2).Value) = True Then
'        TutanakNoText.Value = ThisWorkbook.Worksheets(6).Cells(Say, 2).Value + 1
'    ElseIf ThisWorkbook.Worksheets(6).Cells(Say, 2).Value = "" Then
'        TutanakNoText.Value = 1
'    End If
End If

'Tip3
If Tip3Option.Value = True Then
    Say = ThisWorkbook.Worksheets(6).Range("C100000").End(xlUp).Row
    If Say < 3 Then
        Say = 3
    End If
    Cont = 0
    TutanakNoText.Clear
    For j = Say To 3 Step -1
        If Cont < 5 Then
            If ThisWorkbook.Worksheets(6).Cells(j, 3).Value <> "" Then
                Cont = Cont + 1
                Tno = ThisWorkbook.Worksheets(6).Cells(j, 3).Value
                With TutanakNoText
                    .AddItem (Tno)
                End With
            End If
        Else
            GoTo DonguSon2
        End If
    Next j
DonguSon2:
'    If ThisWorkbook.Worksheets(6).Cells(Say, 3).Value <> "" And IsNumeric(ThisWorkbook.Worksheets(6).Cells(Say, 3).Value) = True Then
'        TutanakNoText.Value = ThisWorkbook.Worksheets(6).Cells(Say, 3).Value + 1
'    ElseIf ThisWorkbook.Worksheets(6).Cells(Say, 3).Value = "" Then
'        TutanakNoText.Value = 1
'    End If
End If

'Tip4
If Tip4Option.Value = True Then
    Say = ThisWorkbook.Worksheets(6).Range("D100000").End(xlUp).Row
    If Say < 3 Then
        Say = 3
    End If
    Cont = 0
    TutanakNoText.Clear
    For j = Say To 3 Step -1
        If Cont < 5 Then
            If ThisWorkbook.Worksheets(6).Cells(j, 4).Value <> "" Then
                Cont = Cont + 1
                Tno = ThisWorkbook.Worksheets(6).Cells(j, 4).Value
                With TutanakNoText
                    .AddItem (Tno)
                End With
            End If
        Else
            GoTo DonguSon3
        End If
    Next j
DonguSon3:
'    If ThisWorkbook.Worksheets(6).Cells(Say, 4).Value <> "" And IsNumeric(ThisWorkbook.Worksheets(6).Cells(Say, 4).Value) = True Then
'        TutanakNoText.Value = ThisWorkbook.Worksheets(6).Cells(Say, 4).Value + 1
'    ElseIf ThisWorkbook.Worksheets(6).Cells(Say, 4).Value = "" Then
'        TutanakNoText.Value = 1
'    End If
End If

ThisWorkbook.Worksheets(6).Protect Password:="123" ', DrawingObjects:=False
ThisWorkbook.Protect "123"

Application.ScreenUpdating = True

End Sub


