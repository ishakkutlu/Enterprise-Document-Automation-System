VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} core_acceptance_manager_UI 
   Caption         =   "Enterprise Document Automation System"
   ClientHeight    =   11130
   ClientLeft      =   30
   ClientTop       =   375
   ClientWidth     =   16320
   OleObjectBlob   =   "core_acceptance_manager_UI.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "core_acceptance_manager_UI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim GirisDirektYazdir As Boolean
Dim GirisSayPrt As Variant

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

Private Sub CheckBoxDosyaNo_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
CheckBoxDosyaNo.BackColor = RGB(60, 100, 180) 'RGB(60, 100, 180) 'RGB(60, 100, 180)
CheckBoxDosyaNo.ForeColor = RGB(255, 255, 255)
End Sub

Private Sub TutanakTarihiLabel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
TutanakTarihiLabel.BackColor = RGB(60, 100, 180) 'RGB(60, 100, 180) 'RGB(60, 100, 180)
TutanakTarihiLabel.ForeColor = RGB(255, 255, 255)
End Sub
Private Sub TutanakTarihiText_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
TutanakTarihiLabel.BackColor = RGB(254, 254, 254)
TutanakTarihiLabel.ForeColor = RGB(30, 30, 30)
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

Sub ColorChangerGenel()

If Tutanak.BackColor <> RGB(225, 235, 245) Then
Tutanak.BackColor = RGB(225, 235, 245)
Tutanak.ForeColor = RGB(30, 30, 30)
End If
If DirektYazdir.BackColor <> RGB(225, 235, 245) Then
DirektYazdir.BackColor = RGB(225, 235, 245)
DirektYazdir.ForeColor = RGB(30, 30, 30)
End If
If TakvimBtn.BackColor <> RGB(225, 235, 245) Then
TakvimBtn.BackColor = RGB(225, 235, 245)
TakvimBtn.ForeColor = RGB(30, 30, 30)
End If
If Yardim.BackColor <> RGB(225, 235, 245) Then
Yardim.BackColor = RGB(225, 235, 245)
Yardim.ForeColor = RGB(30, 30, 30)
End If
If Kapat.BackColor <> RGB(225, 235, 245) Then
Kapat.BackColor = RGB(225, 235, 245)
Kapat.ForeColor = RGB(30, 30, 30)
End If

If CheckBox1.BackColor <> RGB(254, 254, 254) Then
CheckBox1.BackColor = RGB(254, 254, 254) 'RGB(254, 254, 254)
CheckBox1.ForeColor = RGB(70, 70, 70)
End If
If CheckBox2.BackColor <> RGB(254, 254, 254) Then
CheckBox2.BackColor = RGB(254, 254, 254) 'RGB(254, 254, 254)
CheckBox2.ForeColor = RGB(70, 70, 70)
End If
If Ekle.BackColor <> RGB(254, 254, 254) Then
Ekle.BackColor = RGB(254, 254, 254)
Ekle.ForeColor = RGB(70, 70, 70)
End If
If Kaldir.BackColor <> RGB(254, 254, 254) Then
Kaldir.BackColor = RGB(254, 254, 254)
Kaldir.ForeColor = RGB(70, 70, 70)
End If

If CheckBoxDosyaNo.BackColor <> RGB(254, 254, 254) Then
CheckBoxDosyaNo.BackColor = RGB(254, 254, 254)
CheckBoxDosyaNo.ForeColor = RGB(70, 70, 70)
End If
If TutanakTarihiLabel.BackColor <> RGB(254, 254, 254) Then
TutanakTarihiLabel.BackColor = RGB(254, 254, 254)
TutanakTarihiLabel.ForeColor = RGB(70, 70, 70)
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
Private Sub Yardim_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
Yardim.BackColor = RGB(60, 100, 180) 'RGB(60, 100, 180) 'RGB(60, 100, 180)
Yardim.ForeColor = RGB(255, 255, 255)
End Sub
Private Sub Kapat_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
Kapat.BackColor = RGB(60, 100, 180) 'RGB(60, 100, 180) 'RGB(60, 100, 180)
Kapat.ForeColor = RGB(255, 255, 255)
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
    For Each ctl In core_acceptance_manager_UI.Frame1.Controls
        If TypeName(ctl) = "ListBox" Then
            ctl.Selected(0) = True
        End If
    Next ctl
Else 'Tümünü iptal et
    For Each ctl In core_acceptance_manager_UI.Frame1.Controls
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
    For Each ctl In core_acceptance_manager_UI.Frame2.Controls
        If TypeName(ctl) = "ListBox" Then
            ctl.Selected(0) = True
        End If
    Next ctl
Else 'Tümünü iptal et
    For Each ctl In core_acceptance_manager_UI.Frame2.Controls
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
For Each ctl In core_acceptance_manager_UI.Frame1.Controls
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
For Each ctl In core_acceptance_manager_UI.Frame2.Controls
    If TypeName(ctl) = "ListBox" Then
        Say2 = Say2 + 1
    End If
    If TypeName(ctl) = "Label" Then
        Frame2.Controls.Remove ctl.name
    End If
Next ctl

Say1 = 0
For Each ctl In core_acceptance_manager_UI.Frame1.Controls
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
For Each ctl In core_acceptance_manager_UI.Frame1.Controls
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
For Each ctl In core_acceptance_manager_UI.Frame2.Controls
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
For Each ctl In core_acceptance_manager_UI.Frame2.Controls
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
For Each ctl In core_acceptance_manager_UI.Frame1.Controls
    If TypeName(ctl) = "ListBox" Then
        Say1 = Say1 + 1
    End If
    If TypeName(ctl) = "Label" Then
        Frame1.Controls.Remove ctl.name
    End If
Next ctl

Say2 = 0
For Each ctl In core_acceptance_manager_UI.Frame2.Controls
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
For Each ctl In core_acceptance_manager_UI.Frame2.Controls
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
For Each ctl In core_acceptance_manager_UI.Frame1.Controls
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
For Each ctl In core_acceptance_manager_UI.Frame1.Controls
    If TypeName(ctl) = "ListBox" Then
       core_acceptance_manager_UI.Frame1.Controls.Remove ctl.name
    End If
    If TypeName(ctl) = "Label" Then
        core_acceptance_manager_UI.Frame1.Controls.Remove ctl.name
    End If
Next ctl
For Each ctl In core_acceptance_manager_UI.Frame2.Controls
    If TypeName(ctl) = "ListBox" Then
       core_acceptance_manager_UI.Frame2.Controls.Remove ctl.name
    End If
    If TypeName(ctl) = "Label" Then
        core_acceptance_manager_UI.Frame2.Controls.Remove ctl.name
    End If
Next ctl



CalTarihTakip = CalTarih
ContTakip = 0

Call ModuleReport1.Rapor1GelenBelgeGiris
Call ModuleReport2.Rapor2GelenBelgeGiris


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
    
    GirisDirektYazdir = True

    On Error GoTo Son
    GirisSayPrt = InputBox(Prompt:="Please enter the number of copies you wish to print.", Title:="Enterprise Document Automation System")
    If GirisSayPrt = "" Or GirisSayPrt = 0 Then
        GoTo Son
    End If
    If Not IsNumeric(GirisSayPrt) Then
        MsgBox "Invalid input detected: Please enter a numeric value to start the printing process.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    If GirisSayPrt > 3 Then
        MsgBox "Printing more than 3 copies at once is not allowed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If

    Call Tutanak_Click
    
    GirisDirektYazdir = False

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
SourceTaslak = AutoPath & "\System Files\Help Documents\Acceptance Manager Panel – Help.docm"
'Birden fazla kullanıcının operasyon klasöründe çakışmasını önlemek için operasyon klasörünün içinde yeni klasör aç.
DestOpUserFolderName = "Operation-" & Left(ThisWorkbook.name, InStr(ThisWorkbook.name, ".") - 1) '& " " & Format(Now(), "ddmmyyyyhhmmss")
DestOpUserFolder = DestOperasyon & DestOpUserFolderName & "\"

'System Files folder name check.
If Not Dir(AutoPath & "\System Files\", vbDirectory) <> vbNullString Then
    MsgBox "Unable to access directory: " & AutoPath & "\System Files\" & ". The folder named 'System Files' may have been renamed or deleted.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

'Operation folder name check.
If Not Dir(DestOperasyon, vbDirectory) <> vbNullString Then
    MsgBox "Unable to access directory: " & DestOperasyon & ". The folder named 'Operation' may have been renamed or deleted.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

'RmDir DestOpUserFolder 'To be added: Clean DestOpUserFolder upon system shutdown
'_______________

'Folder names check.
If Not Dir(SourceTaslak, vbDirectory) <> vbNullString Then
    MsgBox "Unable to access directory: " & SourceTaslak & ". Folder or file names within this directory may have been changed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
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
Dim AutoPath As String, DestOperasyon As String, SourceXXS As String, DestOpUserFolderName As String, DestOpUserFolder As String
Dim ReNameXXS As String, OpenKontrolName As String, ContSay As Long, KontrolFile As String
Dim fso As Object, objWord As Object, objDoc As Object, Birimx As String, OpenControl As String
Dim Cont As Long, SiraBul As Range, i As Long, j As Long, Say As Long
Dim ctl As MSForms.Control, Ctlx As MSForms.Control, Say2 As Integer, GelenTema As String, IlkSira As Long
Dim MyRange As Object, SayfaSay As Integer, VeriKontrol As Integer
Dim a() As Variant, b As Variant, DosyaNoKolon As Integer
Dim Ek1, Ek2, Ek3, Ek4, Ek5 As String

ThisWorkbook.Activate

'Worksheets(3).Activate
Application.ScreenUpdating = False
Application.DisplayAlerts = False

'Worksheets(3).Unprotect Password:="123"

'Columns("CE:CF").EntireColumn.Hidden = False
'Columns("CK:CN").EntireColumn.Hidden = False

If TutanakTarihiText.Value = "" Then
    MsgBox "The statement date has not been specified; therefore, your operation cannot be performed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If
If TutanakNoText.Value = "" Then
    MsgBox "The statement number has not been specified; therefore, your operation cannot be performed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If


VeriKontrol = 0
For Each ctl In core_acceptance_manager_UI.Frame2.Controls
    If TypeName(ctl) = "ListBox" Then
        VeriKontrol = VeriKontrol + 1
        GoTo Devam2
    End If
Next ctl
Devam2:
If VeriKontrol = 0 Then
    MsgBox "Your operation cannot be completed because no entries have been made in the 'To Be Sent' section for document acceptance.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If


'Pathfinder...
AutoPath = ThisWorkbook.Path
DestOperasyon = AutoPath & "\System Files\Operation\"
'XXS TANIMLARI
DosyaNoKolon = 0
If CheckBoxDosyaNo.Value = True Then
    DosyaNoKolon = 1
    SourceXXS = AutoPath & "\System Files\System Templates\Acceptance and Delivery Statements\Acceptance Statement (File No).docm"
Else
    SourceXXS = AutoPath & "\System Files\System Templates\Acceptance and Delivery Statements\Acceptance Statement.docm"
End If

'Birden fazla kullanıcının operasyon klasöründe çakışmasını önlemek için operasyon klasörünün içinde yeni klasör aç.
DestOpUserFolderName = "Operation-" & Left(ThisWorkbook.name, InStr(ThisWorkbook.name, ".") - 1) '& " " & Format(Now(), "ddmmyyyyhhmmss")
DestOpUserFolder = DestOperasyon & DestOpUserFolderName & "\"

'Sistem Files folder name check.
If Not Dir(AutoPath & "\System Files\", vbDirectory) <> vbNullString Then
    MsgBox "Unable to access directory: " & AutoPath & "\System Files\" & ". The folder named 'System Files' may have been renamed or deleted.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

'Operation folder name check.
If Not Dir(DestOperasyon, vbDirectory) <> vbNullString Then
    MsgBox "Unable to access directory: " & DestOperasyon & ". The folder named 'Operation' may have been renamed or deleted.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

'Folder names check.
If Not Dir(SourceXXS, vbDirectory) <> vbNullString Then
    MsgBox "Unable to access directory: " & SourceXXS & ". Folder or file names within this directory may have been changed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

'Operation klasörü içinde kullanıcı modülü klasörü yoksa oluştur.
If Not Dir(DestOpUserFolder, vbDirectory) <> vbNullString Then
    MkDir DestOpUserFolder
End If

ReNameXXS = "Acceptance Statement"

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
fso.CopyFile (SourceXXS), DestOpUserFolder & ReNameXXS & ".docm", True
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
objWord.Documents.Open FileName:=DestOpUserFolder & ReNameXXS & ".docm"
objWord.Visible = True
objWord.Activate 'Ekrana getirir.
'objDoc.Activate 'Ekrana getirmez.
'objWord.Application.WindowState = wdWindowStateMaximize
Set objDoc = GetObject(DestOpUserFolder & ReNameXXS & ".docm")
'________________________________________

'Birim
Birimx = UCase(Replace(Replace(ThisWorkbook.Worksheets(2).Cells(6, 99).Value, "i", "I"), "ı", "I")) & " UNIT"
objDoc.Tables(1).Cell(Row:=1, Column:=1).Range.Text = Birimx
'Tutanak tarihi
objDoc.Tables(1).Cell(Row:=4, Column:=3).Range.Text = TutanakTarihiText.Value
'Tutanak no
objDoc.Tables(1).Cell(Row:=5, Column:=3).Range.Text = TutanakNoText.Value



'________________AKTARIMLAR (Başlangıç)

Say2 = 0
SayfaSay = 0
For Each ctl In core_acceptance_manager_UI.Frame2.Controls '________________RAPOR1 BÖLÜMÜ (Başlangıç)
    If TypeName(ctl) = "ListBox" Then
        'MsgBox Mid(Ctl.List(0), InStr(Ctl.List(0), "|") + 1, 2)
        If Mid(ctl.List(0), InStr(ctl.List(0), "|") + 1, 2) = " Ö" Then
            Say2 = Say2 + 1
            'MsgBox Left(Ctl.List(0), InStr(Ctl.List(0), "|") - 2)
            Set SiraBul = ThisWorkbook.Worksheets(3).Range("E7:E100000").Find(What:=Left(ctl.List(0), InStr(ctl.List(0), "|") - 2), SearchDirection:=xlNext, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not SiraBul Is Nothing Then
                'Cells(SiraBul.Row, 31).Value
                IlkSira = SiraBul.Row
            Else
                'MsgBox "Belirtilen tarihte herhangi bir tutanak1 işlemi yapılmadığından işleminiz gerçekleştirilemiyor.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
                GoTo Rapor1Devam
            End If
            
            'Tabloya satır ekle
            If Say2 > 1 Then
                objDoc.Tables(2).Rows.Add BeforeRow:=objDoc.Tables(2).Rows(Say2 + 1)
            End If
            objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=1).Range.Text = Say2
            If CheckBoxDosyaNo.Value = True Then
                objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=2).Range.Text = ThisWorkbook.Worksheets(3).Cells(SiraBul.Row, 34).Value
            End If
            objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=2 + DosyaNoKolon).Range.Text = ThisWorkbook.Worksheets(3).Cells(SiraBul.Row, 28).Value
            objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=3 + DosyaNoKolon).Range.Text = ThisWorkbook.Worksheets(3).Cells(SiraBul.Row, 20).Value
            objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=4 + DosyaNoKolon).Range.Text = ThisWorkbook.Worksheets(3).Cells(SiraBul.Row, 21).Value
            'Gelen tema
            If ThisWorkbook.Worksheets(3).Cells(IlkSira, 25).Value = "Provincial Directorate B" Then
                If ThisWorkbook.Worksheets(3).Cells(IlkSira, 26).Value <> "" Then
                    GelenTema = ThisWorkbook.Worksheets(3).Cells(IlkSira, 17).Value & " Provincial Governorship Provincial Directorate B " & ThisWorkbook.Worksheets(3).Cells(IlkSira, 26).Value
                Else
                    GelenTema = ThisWorkbook.Worksheets(3).Cells(IlkSira, 17).Value & " Provincial Governorship Provincial Directorate B"
                End If
            ElseIf ThisWorkbook.Worksheets(3).Cells(IlkSira, 25).Value = "District Directorate B" Then
                If ThisWorkbook.Worksheets(3).Cells(IlkSira, 26).Value <> "" Then
                    GelenTema = ThisWorkbook.Worksheets(3).Cells(IlkSira, 18).Value & " District Governorship District Directorate B " & ThisWorkbook.Worksheets(3).Cells(IlkSira, 26).Value
                Else
                    GelenTema = ThisWorkbook.Worksheets(3).Cells(IlkSira, 18).Value & " District Governorship District Directorate B"
                End If
            ElseIf ThisWorkbook.Worksheets(3).Cells(IlkSira, 25).Value = "Provincial Directorate C" Then
                If ThisWorkbook.Worksheets(3).Cells(IlkSira, 26).Value <> "" Then
                    GelenTema = ThisWorkbook.Worksheets(3).Cells(IlkSira, 17).Value & " Provincial Governorship Provincial Directorate C " & ThisWorkbook.Worksheets(3).Cells(IlkSira, 26).Value
                Else
                    GelenTema = ThisWorkbook.Worksheets(3).Cells(IlkSira, 17).Value & " Provincial Governorship Provincial Directorate C"
                End If
            ElseIf ThisWorkbook.Worksheets(3).Cells(IlkSira, 25).Value = "District Directorate C" Then
                If ThisWorkbook.Worksheets(3).Cells(IlkSira, 26).Value <> "" Then
                    GelenTema = ThisWorkbook.Worksheets(3).Cells(IlkSira, 18).Value & " District Governorship District Directorate C " & ThisWorkbook.Worksheets(3).Cells(IlkSira, 26).Value
                Else
                    GelenTema = ThisWorkbook.Worksheets(3).Cells(IlkSira, 18).Value & " District Governorship District Directorate C"
                End If
            ElseIf ThisWorkbook.Worksheets(3).Cells(IlkSira, 25).Value = "Provincial Directorate D" Then
                If ThisWorkbook.Worksheets(3).Cells(IlkSira, 26).Value <> "" Then
                    GelenTema = ThisWorkbook.Worksheets(3).Cells(IlkSira, 17).Value & " Provincial Governorship Provincial Directorate D " & ThisWorkbook.Worksheets(3).Cells(IlkSira, 26).Value
                Else
                    GelenTema = ThisWorkbook.Worksheets(3).Cells(IlkSira, 17).Value & " Provincial Governorship Provincial Directorate D"
                End If
            ElseIf ThisWorkbook.Worksheets(3).Cells(IlkSira, 25).Value = "District Directorate D" Then
                If ThisWorkbook.Worksheets(3).Cells(IlkSira, 26).Value <> "" Then
                    GelenTema = ThisWorkbook.Worksheets(3).Cells(IlkSira, 18).Value & " District Governorship District Directorate D " & ThisWorkbook.Worksheets(3).Cells(IlkSira, 26).Value
                Else
                    GelenTema = ThisWorkbook.Worksheets(3).Cells(IlkSira, 18).Value & " District Governorship District Directorate D"
                End If
            ElseIf ThisWorkbook.Worksheets(3).Cells(IlkSira, 25).Value = "Provincial Directorate E" Then
                If ThisWorkbook.Worksheets(3).Cells(IlkSira, 26).Value <> "" Then
                    GelenTema = ThisWorkbook.Worksheets(3).Cells(IlkSira, 17).Value & " Provincial Governorship Provincial Directorate E " & ThisWorkbook.Worksheets(3).Cells(IlkSira, 26).Value
                Else
                    GelenTema = ThisWorkbook.Worksheets(3).Cells(IlkSira, 17).Value & " Provincial Governorship Provincial Directorate E"
                End If
            ElseIf ThisWorkbook.Worksheets(3).Cells(IlkSira, 25).Value = "District Directorate E" Then
                If ThisWorkbook.Worksheets(3).Cells(IlkSira, 26).Value <> "" Then
                    GelenTema = ThisWorkbook.Worksheets(3).Cells(IlkSira, 18).Value & " District Governorship District Directorate E " & ThisWorkbook.Worksheets(3).Cells(IlkSira, 26).Value
                Else
                    GelenTema = ThisWorkbook.Worksheets(3).Cells(IlkSira, 18).Value & " District Governorship District Directorate E"
                End If
            ElseIf InStr(ThisWorkbook.Worksheets(3).Cells(IlkSira, 25).Value, "General Directorate") <> 0 Or InStr(ThisWorkbook.Worksheets(3).Cells(IlkSira, 25).Value, "Regional Directorate") <> 0 Then
                If ThisWorkbook.Worksheets(3).Cells(IlkSira, 26).Value <> "" Then
                    GelenTema = WorksheetFunction.Proper(ThisWorkbook.Worksheets(2).Cells(6, 111).Value) & " " & ThisWorkbook.Worksheets(3).Cells(IlkSira, 25).Value & " " & ThisWorkbook.Worksheets(3).Cells(IlkSira, 26).Value
                Else
                    GelenTema = WorksheetFunction.Proper(ThisWorkbook.Worksheets(2).Cells(6, 111).Value) & " " & ThisWorkbook.Worksheets(3).Cells(IlkSira, 25).Value
                End If
            Else
                If ThisWorkbook.Worksheets(3).Cells(IlkSira, 26).Value <> "" Then
                    If InStr(ThisWorkbook.Worksheets(3).Cells(IlkSira, 25).Value, "X.X. ") > 0 Then
                        GelenTema = Mid(ThisWorkbook.Worksheets(3).Cells(IlkSira, 25).Value, 6, Len(ThisWorkbook.Worksheets(3).Cells(IlkSira, 25).Value)) & " " & ThisWorkbook.Worksheets(3).Cells(IlkSira, 26).Value
                    Else
                        GelenTema = ThisWorkbook.Worksheets(3).Cells(IlkSira, 25).Value & " " & ThisWorkbook.Worksheets(3).Cells(IlkSira, 26).Value
                    End If
                Else
                    If InStr(ThisWorkbook.Worksheets(3).Cells(IlkSira, 25).Value, "X.X. ") > 0 Then
                        GelenTema = Mid(ThisWorkbook.Worksheets(3).Cells(IlkSira, 25).Value, 6, Len(ThisWorkbook.Worksheets(3).Cells(IlkSira, 25).Value))
                    Else
                        GelenTema = ThisWorkbook.Worksheets(3).Cells(IlkSira, 25).Value
                    End If
                End If
            End If
            objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=5 + DosyaNoKolon).Range.Text = GelenTema
            objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=6 + DosyaNoKolon).Range.Text = ThisWorkbook.Worksheets(3).Cells(SiraBul.Row, 33).Value
            SayfaSay = SayfaSay + ThisWorkbook.Worksheets(3).Cells(SiraBul.Row, 33).Value
        End If
    End If
Rapor1Devam:
Next ctl '________________RAPOR1 BÖLÜMÜ (Bitiş)

For Each ctl In core_acceptance_manager_UI.Frame2.Controls '________________RAPOR BÖLÜMÜ (Başlangıç)
    If TypeName(ctl) = "ListBox" Then
        'MsgBox Mid(Ctl.List(0), InStr(Ctl.List(0), "|") + 1, 2)
        If Mid(ctl.List(0), InStr(ctl.List(0), "|") + 1, 2) = " R" Then
            Say2 = Say2 + 1
            'MsgBox Left(Ctl.List(0), InStr(Ctl.List(0), "|") - 2)
            Set SiraBul = ThisWorkbook.Worksheets(4).Range("E7:E100000").Find(What:=Left(ctl.List(0), InStr(ctl.List(0), "|") - 2), SearchDirection:=xlNext, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not SiraBul Is Nothing Then
                'Cells(SiraBul.Row, 31).Value
                IlkSira = SiraBul.Row
            Else
                'MsgBox "Belirtilen tarihte herhangi bir tutanak1 işlemi yapılmadığından işleminiz gerçekleştirilemiyor.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
                GoTo RaporDevam
            End If
            
            'Tabloya satır ekle
            If Say2 > 1 Then
                objDoc.Tables(2).Rows.Add BeforeRow:=objDoc.Tables(2).Rows(Say2 + 1)
            End If
            objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=1).Range.Text = Say2
            If CheckBoxDosyaNo.Value = True Then
                objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=2).Range.Text = ThisWorkbook.Worksheets(4).Cells(SiraBul.Row, 42).Value
            End If
            objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=2 + DosyaNoKolon).Range.Text = ThisWorkbook.Worksheets(4).Cells(SiraBul.Row, 36).Value
            objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=3 + DosyaNoKolon).Range.Text = ThisWorkbook.Worksheets(4).Cells(SiraBul.Row, 28).Value
            objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=4 + DosyaNoKolon).Range.Text = ThisWorkbook.Worksheets(4).Cells(SiraBul.Row, 29).Value
            'Gelen tema
            If ThisWorkbook.Worksheets(4).Cells(IlkSira, 33).Value = "Provincial Directorate B" Then
                If ThisWorkbook.Worksheets(4).Cells(IlkSira, 34).Value <> "" Then
                    GelenTema = ThisWorkbook.Worksheets(4).Cells(IlkSira, 25).Value & " Provincial Governorship Provincial Directorate B " & ThisWorkbook.Worksheets(4).Cells(IlkSira, 34).Value
                Else
                    GelenTema = ThisWorkbook.Worksheets(4).Cells(IlkSira, 25).Value & " Provincial Governorship Provincial Directorate B"
                End If
            ElseIf ThisWorkbook.Worksheets(4).Cells(IlkSira, 33).Value = "District Directorate B" Then
                If ThisWorkbook.Worksheets(4).Cells(IlkSira, 34).Value <> "" Then
                    GelenTema = ThisWorkbook.Worksheets(4).Cells(IlkSira, 26).Value & " District Governorship District Directorate B " & ThisWorkbook.Worksheets(4).Cells(IlkSira, 34).Value
                Else
                    GelenTema = ThisWorkbook.Worksheets(4).Cells(IlkSira, 26).Value & " District Governorship District Directorate B"
                End If
            ElseIf ThisWorkbook.Worksheets(4).Cells(IlkSira, 33).Value = "Provincial Directorate C" Then
                If ThisWorkbook.Worksheets(4).Cells(IlkSira, 34).Value <> "" Then
                    GelenTema = ThisWorkbook.Worksheets(4).Cells(IlkSira, 25).Value & " Provincial Governorship Provincial Directorate C " & ThisWorkbook.Worksheets(4).Cells(IlkSira, 34).Value
                Else
                    GelenTema = ThisWorkbook.Worksheets(4).Cells(IlkSira, 25).Value & " Provincial Governorship Provincial Directorate C"
                End If
            ElseIf ThisWorkbook.Worksheets(4).Cells(IlkSira, 33).Value = "District Directorate C" Then
                If ThisWorkbook.Worksheets(4).Cells(IlkSira, 34).Value <> "" Then
                    GelenTema = ThisWorkbook.Worksheets(4).Cells(IlkSira, 26).Value & " District Governorship District Directorate C " & ThisWorkbook.Worksheets(4).Cells(IlkSira, 34).Value
                Else
                    GelenTema = ThisWorkbook.Worksheets(4).Cells(IlkSira, 26).Value & " District Governorship District Directorate C"
                End If
            ElseIf ThisWorkbook.Worksheets(4).Cells(IlkSira, 33).Value = "Provincial Directorate D" Then
                If ThisWorkbook.Worksheets(4).Cells(IlkSira, 34).Value <> "" Then
                    GelenTema = ThisWorkbook.Worksheets(4).Cells(IlkSira, 25).Value & " Provincial Governorship Provincial Directorate D " & ThisWorkbook.Worksheets(4).Cells(IlkSira, 34).Value
                Else
                    GelenTema = ThisWorkbook.Worksheets(4).Cells(IlkSira, 25).Value & " Provincial Governorship Provincial Directorate D"
                End If
            ElseIf ThisWorkbook.Worksheets(4).Cells(IlkSira, 33).Value = "District Directorate D" Then
                If ThisWorkbook.Worksheets(4).Cells(IlkSira, 34).Value <> "" Then
                    GelenTema = ThisWorkbook.Worksheets(4).Cells(IlkSira, 26).Value & " District Governorship District Directorate D " & ThisWorkbook.Worksheets(4).Cells(IlkSira, 34).Value
                Else
                    GelenTema = ThisWorkbook.Worksheets(4).Cells(IlkSira, 26).Value & " District Governorship District Directorate D"
                End If
            ElseIf ThisWorkbook.Worksheets(4).Cells(IlkSira, 33).Value = "Provincial Directorate E" Then
                If ThisWorkbook.Worksheets(4).Cells(IlkSira, 34).Value <> "" Then
                    GelenTema = ThisWorkbook.Worksheets(4).Cells(IlkSira, 25).Value & " Provincial Governorship Provincial Directorate E " & ThisWorkbook.Worksheets(4).Cells(IlkSira, 34).Value
                Else
                    GelenTema = ThisWorkbook.Worksheets(4).Cells(IlkSira, 25).Value & " Provincial Governorship Provincial Directorate E"
                End If
            ElseIf ThisWorkbook.Worksheets(4).Cells(IlkSira, 33).Value = "District Directorate E" Then
                If ThisWorkbook.Worksheets(4).Cells(IlkSira, 34).Value <> "" Then
                    GelenTema = ThisWorkbook.Worksheets(4).Cells(IlkSira, 26).Value & " District Governorship District Directorate E " & ThisWorkbook.Worksheets(4).Cells(IlkSira, 34).Value
                Else
                    GelenTema = ThisWorkbook.Worksheets(4).Cells(IlkSira, 26).Value & " District Governorship District Directorate E"
                End If
            ElseIf InStr(ThisWorkbook.Worksheets(4).Cells(IlkSira, 33).Value, "General Directorate") <> 0 Or InStr(ThisWorkbook.Worksheets(4).Cells(IlkSira, 33).Value, "Regional Directorate") <> 0 Then
                If ThisWorkbook.Worksheets(4).Cells(IlkSira, 34).Value <> "" Then
                    GelenTema = WorksheetFunction.Proper(ThisWorkbook.Worksheets(2).Cells(6, 111).Value) & " " & ThisWorkbook.Worksheets(4).Cells(IlkSira, 33).Value & " " & ThisWorkbook.Worksheets(4).Cells(IlkSira, 34).Value
                Else
                    GelenTema = WorksheetFunction.Proper(ThisWorkbook.Worksheets(2).Cells(6, 111).Value) & " " & ThisWorkbook.Worksheets(4).Cells(IlkSira, 33).Value
                End If
            Else
                If ThisWorkbook.Worksheets(4).Cells(IlkSira, 34).Value <> "" Then
                    If InStr(ThisWorkbook.Worksheets(4).Cells(IlkSira, 33).Value, "X.X. ") > 0 Then
                        GelenTema = Mid(ThisWorkbook.Worksheets(4).Cells(IlkSira, 33).Value, 6, Len(ThisWorkbook.Worksheets(4).Cells(IlkSira, 33).Value)) & " " & ThisWorkbook.Worksheets(4).Cells(IlkSira, 34).Value
                    Else
                        GelenTema = ThisWorkbook.Worksheets(4).Cells(IlkSira, 33).Value & " " & ThisWorkbook.Worksheets(4).Cells(IlkSira, 34).Value
                    End If
                Else
                    If InStr(ThisWorkbook.Worksheets(4).Cells(IlkSira, 33).Value, "X.X. ") > 0 Then
                        GelenTema = Mid(ThisWorkbook.Worksheets(4).Cells(IlkSira, 33).Value, 6, Len(ThisWorkbook.Worksheets(4).Cells(IlkSira, 33).Value))
                    Else
                        GelenTema = ThisWorkbook.Worksheets(4).Cells(IlkSira, 33).Value
                    End If
                End If
            End If
            objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=5 + DosyaNoKolon).Range.Text = GelenTema
            objDoc.Tables(2).Cell(Row:=Say2 + 1, Column:=6 + DosyaNoKolon).Range.Text = ThisWorkbook.Worksheets(4).Cells(SiraBul.Row, 41).Value
            SayfaSay = SayfaSay + ThisWorkbook.Worksheets(4).Cells(SiraBul.Row, 41).Value
        End If
    End If
RaporDevam:
Next ctl '________________RAPOR BÖLÜMÜ (Bitiş)



'________________AKTARIMLAR (Bitiş)


'Artık satırı sil
objDoc.Tables(2).Rows(Say2 + 2).Delete

If Say2 > 1 Then
    Ek1 = "documents"
Else
    Ek1 = "document"
End If
If SayfaSay > 1 Then
    Ek2 = "pages"
Else
    Ek2 = "page"
End If
objDoc.Tables(3).Cell(Row:=1, Column:=1).Range.Text = "The " & Say2 & " " & Ek1 & " mentioned above (a total of " & SayfaSay & " " & Ek2 & ") have been delivered to the XXX Service for entry into the XX System."

'Belge sayısını bold yap.
Set MyRange = objDoc.Tables(3).Cell(Row:=1, Column:=1).Range
With MyRange.Find
    .Execute FindText:=Say2
End With
MyRange.Font.Bold = True
'Sonrakinde yer alan sayfa sayısını bold yap
Set MyRange = objDoc.Tables(3).Cell(Row:=1, Column:=1).Range
With MyRange.Find
    .Execute FindText:=SayfaSay
    .Execute Forward:=True
End With
MyRange.Font.Bold = True

'Çapraz teslim güncellemesi
objDoc.Tables(5).Cell(Row:=1, Column:=1).Range.Text = "The " & Say2 & " " & Ek1 & " mentioned above (a total of " & SayfaSay & " " & Ek2 & ") have been received from the XXX Service."

'Belge sayısını bold yap.
Set MyRange = objDoc.Tables(5).Cell(Row:=1, Column:=1).Range
With MyRange.Find
    .Execute FindText:=Say2
End With
MyRange.Font.Bold = True
'Sonrakinde yer alan sayfa sayısını bold yap
Set MyRange = objDoc.Tables(5).Cell(Row:=1, Column:=1).Range
With MyRange.Find
    .Execute FindText:=SayfaSay
    .Execute Forward:=True
End With
MyRange.Font.Bold = True


'Alt bilgi ekle
objDoc.Sections(1).Footers(wdHeaderFooterPrimary).Range.Tables(1).Cell(Row:=1, Column:=1).Range.Text = ReNameXXS


'Direkt yazdır
If GirisDirektYazdir = True Then
    objWord.Activate
'    For i = 1 To GirisSayPrt
        'On Error Resume Next 'Aktif yazıcı bulamazsa hata veridiği için eklendi. (sadece burada oluyor.)
        'objDoc.PrintOut
        objDoc.PrintOut Background:=False, Copies:=GirisSayPrt
'    Next i
    'objWord.Documents.Save
    objDoc.Close SaveChanges:=False
    objWord.Visible = False
End If

'Tutanak numarasını kaydet
ThisWorkbook.Unprotect "123"
ThisWorkbook.Worksheets(6).Unprotect Password:="123"
'Comboda tanımlı değer ise
a() = TutanakNoText.List
For b = LBound(a) To UBound(a)
    If a(b, 0) = TutanakNoText.Value Then
        GoTo GirisNoEklemedenAtla
    End If
Next b
Say = ThisWorkbook.Worksheets(6).Range("A100000").End(xlUp).Row + 1
If Say < 3 Then
    Say = 3
End If
ThisWorkbook.Worksheets(6).Cells(Say, 1) = TutanakNoText.Value
GirisNoEklemedenAtla:

ThisWorkbook.Worksheets(6).Protect Password:="123" ', DrawingObjects:=False
ThisWorkbook.Protect "123"

'TutanakNoText.Value = ""

Son:
    
    Call SonTutanakNoGiris

    'Nesneleri temizle
    Set fso = Nothing
    Set objWord = Nothing
    Set objDoc = Nothing

    GirisDirektYazdir = False

    'Columns("CK:CN").EntireColumn.Hidden = True
    'Columns("CE:CF").EntireColumn.Hidden = True
    
    'Worksheets(3).Protect Password:="123"', DrawingObjects:=False
    
    'Worksheets(3).Activate
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    
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

For Each ClrLab In core_acceptance_manager_UI.Controls

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

core_acceptance_manager_UI.BackColor = RGB(230, 230, 230) 'YENİ

LblSiraNo1.BorderColor = RGB(180, 180, 180)
LblBelgeNo1.BorderColor = RGB(180, 180, 180)
LblGonderen1.BorderColor = RGB(180, 180, 180)
LblSiraNo2.BorderColor = RGB(180, 180, 180)
LblBelgeNo2.BorderColor = RGB(180, 180, 180)
LblGonderen2.BorderColor = RGB(180, 180, 180)

'Tutanak numarasını çağır
Call SonTutanakNoGiris

TasiyiciFrame.Height = 514
Frame1x.Height = 264
Frame2x.Height = 264

Frame1x.ZOrder msoBringToFront
Frame2x.ZOrder msoBringToFront
Frame1.ZOrder msoBringToFront
Frame2.ZOrder msoBringToFront
FrameAlt1.ZOrder msoBringToFront
FrameAlt2.ZOrder msoBringToFront

GirisDirektYazdir = False

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
        core_acceptance_manager_UI.Width = Rep
        yukseklik = yukseklik - 70
        core_acceptance_manager_UI.Height = yukseklik
        If yukseklik <= 70 Then
            yukseklik = 70
            core_acceptance_manager_UI.Height = yukseklik
        End If
    ElseIf Rep <= 50 Then
        Rep = 50
        core_acceptance_manager_UI.Width = Rep
        yukseklik = yukseklik - 50
        core_acceptance_manager_UI.Height = yukseklik
        If yukseklik <= 50 Then
            yukseklik = 50
            core_acceptance_manager_UI.Height = yukseklik
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

Private Sub SonTutanakNoGiris()
Dim i As Integer
Dim Say As Long, j As Long, Cont As Long, Tno As Variant

Application.ScreenUpdating = False

'Tutanak numarasını getir
ThisWorkbook.Unprotect "123"
ThisWorkbook.Worksheets(6).Unprotect Password:="123"
Say = ThisWorkbook.Worksheets(6).Range("A100000").End(xlUp).Row
If Say < 3 Then
    Say = 3
End If
Cont = 0
TutanakNoText.Clear
For j = Say To 3 Step -1
    If Cont < 5 Then
        If ThisWorkbook.Worksheets(6).Cells(j, 1).Value <> "" Then
            Cont = Cont + 1
            Tno = ThisWorkbook.Worksheets(6).Cells(j, 1).Value
            With TutanakNoText
                .AddItem (Tno)
            End With
        End If
    Else
        GoTo DonguSon
    End If
Next j
DonguSon:

ThisWorkbook.Worksheets(6).Protect Password:="123" ', DrawingObjects:=False
ThisWorkbook.Protect "123"

Application.ScreenUpdating = True

End Sub




