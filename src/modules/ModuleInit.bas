Attribute VB_Name = "ModuleInit"
Option Explicit
Public TumDoc As Boolean
Public SayPrt As Variant

Sub Rapor1Baslat()
Dim ItemBul As Range, UserName As String

TumDoc = False
If ActiveCell.Column = 6 Then
    Call ModuleReport1.Rapor1Tutanak1
ElseIf ActiveCell.Column = 7 Then
    Call ModuleReport1.Rapor1Rapor
ElseIf ActiveCell.Column = 8 Then
    Call ModuleReport1.Rapor1Tutanak2
ElseIf ActiveCell.Column = 9 Then
    Call ModuleReport1.Rapor1UstYazi
ElseIf ActiveCell.Column = 10 Then
    TumDoc = True
    
    'PARAF HAZIRLIK İŞLEMİ
    UserName = Environ("UserProfile")
    UserName = UCase(Right(UserName, 7))
    Set ItemBul = Worksheets(2).Range("DR6:DR1000").Find(What:=UserName, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    If Not ItemBul Is Nothing Then
        '
    Else
        MsgBox UserName & " session is not registered in the system, so your operation cannot be started. Please register your session in the system using the Initials Interface located in the Settings Group of the Enterprise Document Automation System menu and try the operation again.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo ModulSonu
    End If
    
    On Error GoTo ModulSonu
    SayPrt = InputBox(Prompt:="Enter the number of copies you want to print.", Title:="Enterprise Document Automation System")
    If SayPrt = "" Or SayPrt = 0 Then
        GoTo ModulSonu
    End If
    If IsNumeric(SayPrt) = False Then
        MsgBox "Non-numeric input detected, so the print operation could not be started.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo ModulSonu
    End If
    If SayPrt > 3 Then
        MsgBox "The number of copies to be printed at one time cannot exceed 3.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo ModulSonu
    End If

    Call ModuleReport1.Rapor1Tutanak1
    Call ModuleReport1.Rapor1Rapor
    Call ModuleReport1.Rapor1Tutanak2
    Call ModuleReport1.Rapor1UstYazi
End If

ModulSonu:
ThisWorkbook.Worksheets(3).Activate

End Sub

Sub Rapor2Baslat()
Dim ItemBul As Range, UserName As String

TumDoc = False
'Rapor2_1
If ActiveCell.Column = 6 Then
    Call ModuleReport2.Rapor2_1Tutanak1
ElseIf ActiveCell.Column = 7 Then
    Call ModuleReport2.Rapor2_1Rapor
ElseIf ActiveCell.Column = 8 Then
    Call ModuleReport2.Rapor2_1Tutanak2
ElseIf ActiveCell.Column = 9 Then
    Call ModuleReport2.Rapor2_1UstYazi
ElseIf ActiveCell.Column = 10 Then
    TumDoc = True

    'PARAF HAZIRLIK İŞLEMİ
    UserName = Environ("UserProfile")
    UserName = UCase(Right(UserName, 7))
    Set ItemBul = Worksheets(2).Range("DR6:DR1000").Find(What:=UserName, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)

    If Not ItemBul Is Nothing Then
        '
    Else
        MsgBox UserName & " session is not registered in the system, so your operation cannot be started. Please register your session in the system using the Initials Interface located in the Settings Group of the Enterprise Document Automation System menu and try the operation again.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo ModulSonu
    End If
    
    On Error GoTo ModulSonu
    SayPrt = InputBox(Prompt:="Enter the number of copies you want to print.", Title:="Enterprise Document Automation System")
    If SayPrt = "" Or SayPrt = 0 Then
        GoTo ModulSonu
    End If
    If IsNumeric(SayPrt) = False Then
        MsgBox "Non-numeric input detected, so the print operation could not be started.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo ModulSonu
    End If
    If SayPrt > 3 Then
        MsgBox "The number of copies to be printed at one time cannot exceed 3.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo ModulSonu
    End If
    
    Call ModuleReport2.Rapor2_1Tutanak1
    Call ModuleReport2.Rapor2_1Rapor
    Call ModuleReport2.Rapor2_1Tutanak2
    Call ModuleReport2.Rapor2_1UstYazi
End If

'Rapor2_2
If ActiveCell.Column = 13 Then
    Call ModuleReport2.Rapor2_2Tutanak1
ElseIf ActiveCell.Column = 14 Then
    Call ModuleReport2.Rapor2_2Rapor
ElseIf ActiveCell.Column = 15 Then
    Call ModuleReport2.Rapor2_2Tutanak2XXXMudGiden
ElseIf ActiveCell.Column = 16 Then
    Call ModuleReport2.Rapor2_2XXXMudUstYazi
ElseIf ActiveCell.Column = 17 Then
    Call ModuleReport2.Rapor2_2BilgilendirmeUstYazi
ElseIf ActiveCell.Column = 18 Then
    TumDoc = True

    'PARAF HAZIRLIK İŞLEMİ
    UserName = Environ("UserProfile")
    UserName = UCase(Right(UserName, 7))
    Set ItemBul = Worksheets(2).Range("DR6:DR1000").Find(What:=UserName, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    If Not ItemBul Is Nothing Then
        '
    Else
        MsgBox UserName & " session is not registered in the system, so your operation cannot be started. Please register your session in the system using the Initials Interface located in the Settings Group of the Enterprise Document Automation System menu and try the operation again.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo ModulSonu
    End If
    
    On Error GoTo ModulSonu
    SayPrt = InputBox(Prompt:="Enter the number of copies you want to print.", Title:="Enterprise Document Automation System")
    If SayPrt = "" Or SayPrt = 0 Then
        GoTo ModulSonu
    End If
    If IsNumeric(SayPrt) = False Then
        MsgBox "Non-numeric input detected, so the print operation could not be started.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo ModulSonu
    End If
    If SayPrt > 3 Then
        MsgBox "The number of copies to be printed at one time cannot exceed 3.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo ModulSonu
    End If

    Call ModuleReport2.Rapor2_2Tutanak1
    Call ModuleReport2.Rapor2_2Rapor
    Call ModuleReport2.Rapor2_2Tutanak2XXXMudGiden
    Call ModuleReport2.Rapor2_2XXXMudUstYazi
    Call ModuleReport2.Rapor2_2BilgilendirmeUstYazi
    
End If

If ActiveCell.Column = 19 Then 'XXXMud'den gelen tutanak1
    Call ModuleReport2.Rapor2_2XXXMudTutanak1
    GoTo ModulSonu
ElseIf ActiveCell.Column = 20 Then 'XXXMud gelen tutanak2
    Call ModuleReport2.Rapor2_2Tutanak2XXXMudGelen
    GoTo ModulSonu
ElseIf ActiveCell.Column = 21 Then 'İlgili birim tutanak2
    Call ModuleReport2.Rapor2_2Tutanak2IlgiliBirim
    GoTo ModulSonu
ElseIf ActiveCell.Column = 22 Then 'Sonuç
    Call ModuleReport2.Rapor2_2SonucUstYazi
    GoTo ModulSonu
End If

ModulSonu:
ThisWorkbook.Worksheets(4).Activate

End Sub

Sub Rapor3_1Baslat()
Dim ItemBul As Range, UserName As String

TumDoc = False
If ActiveCell.Column = 6 Then
    If Cells(ActiveCell.Row, 100) = "Type A" Then
        Call ModuleReport3.Rapor3_1Tutanak
    ElseIf Cells(ActiveCell.Row, 100) = "Type B" Then
        Call ModuleReport3.Rapor3_1TutanakTipB
    End If
ElseIf ActiveCell.Column = 7 Then
    Call ModuleReport3.Rapor3_1Rapor
ElseIf ActiveCell.Column = 8 Then
    If Cells(ActiveCell.Row, 100) = "Type A" Then
        Call ModuleReport3.Rapor3_1Tutanak2
    ElseIf Cells(ActiveCell.Row, 100) = "Type B" Then
        Call ModuleReport3.Rapor3_1Tutanak2TipB
    End If
ElseIf ActiveCell.Column = 9 Then
    'Rapor3_1de finansal birim üst yazısı yok
ElseIf ActiveCell.Column = 10 Then
    If Cells(ActiveCell.Row, 100) = "Type A" Then
        Call ModuleReport3.Rapor3_1UstYazi
    ElseIf Cells(ActiveCell.Row, 100) = "Type B" Then
        Call ModuleReport3.Rapor3_1UstYaziTipB
    End If
ElseIf ActiveCell.Column = 11 Then
    TumDoc = True

    'PARAF HAZIRLIK İŞLEMİ
    UserName = Environ("UserProfile")
    UserName = UCase(Right(UserName, 7))
    Set ItemBul = Worksheets(2).Range("DR6:DR1000").Find(What:=UserName, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    If Not ItemBul Is Nothing Then
        '
    Else
        MsgBox UserName & " session is not registered in the system, so your operation cannot be started. Please register your session in the system using the Initials Interface located in the Settings Group of the Enterprise Document Automation System menu and try the operation again.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo ModulSonu
    End If
    
    On Error GoTo ModulSonu
    SayPrt = InputBox(Prompt:="Enter the number of copies you want to print.", Title:="Enterprise Document Automation System")
    If SayPrt = "" Or SayPrt = 0 Then
        GoTo ModulSonu
    End If
    If IsNumeric(SayPrt) = False Then
        MsgBox "Non-numeric input detected, so the print operation could not be started.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo ModulSonu
    End If
    If SayPrt > 3 Then
        MsgBox "The number of copies to be printed at one time cannot exceed 3.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo ModulSonu
    End If
    
    'All için TipB/TipA ayrımı
    If Cells(ActiveCell.Row, 100) = "Type A" Then
        Call ModuleReport3.Rapor3_1Tutanak
        Call ModuleReport3.Rapor3_1Rapor
        Call ModuleReport3.Rapor3_1Tutanak2
        Call ModuleReport3.Rapor3_1UstYazi
    ElseIf Cells(ActiveCell.Row, 100) = "Type B" Then
        Call ModuleReport3.Rapor3_1TutanakTipB
        Call ModuleReport3.Rapor3_1Tutanak2TipB
        Call ModuleReport3.Rapor3_1UstYaziTipB
    End If
End If

ModulSonu:
ThisWorkbook.Worksheets(5).Activate

End Sub
Sub FinansalBirimBaslat()
Dim ItemBul As Range, UserName As String

TumDoc = False
If ActiveCell.Column = 6 Then
    If Cells(ActiveCell.Row, 28) = "Type A" Then
        Call ModuleReport3.Rapor3_2Tutanak
    ElseIf Cells(ActiveCell.Row, 28) = "Type B" Then
        Call ModuleReport3.Rapor3_2TutanakTipB
    End If
ElseIf ActiveCell.Column = 7 Then
    'If Cells(ActiveCell.Row, 28) = "Type A" Then
        Call ModuleReport3.Rapor3_2Rapor
    'End If
ElseIf ActiveCell.Column = 8 Then
    If Cells(ActiveCell.Row, 28) = "Type A" Then
        Call ModuleReport3.Rapor3_2Tutanak2
    ElseIf Cells(ActiveCell.Row, 28) = "Type B" Then
        Call ModuleReport3.Rapor3_2Tutanak2TipB
    End If
ElseIf ActiveCell.Column = 9 Then
    If Cells(ActiveCell.Row, 28) = "Type A" Then
        Call ModuleReport3.Rapor3_2FinansalBirimUstYazi
    ElseIf Cells(ActiveCell.Row, 28) = "Type B" Then
        Call ModuleReport3.Rapor3_2FinansalBirimUstYaziTipB
    End If
ElseIf ActiveCell.Column = 10 Then
    If Cells(ActiveCell.Row, 28) = "Type A" Then
        Call ModuleReport3.Rapor3_2UstYazi
    ElseIf Cells(ActiveCell.Row, 28) = "Type B" Then
        Call ModuleReport3.Rapor3_2UstYaziTipB
    End If
ElseIf ActiveCell.Column = 11 Then
    TumDoc = True

    'PARAF HAZIRLIK İŞLEMİ
    UserName = Environ("UserProfile")
    UserName = UCase(Right(UserName, 7))
    Set ItemBul = Worksheets(2).Range("DR6:DR1000").Find(What:=UserName, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    If Not ItemBul Is Nothing Then
        '
    Else
        MsgBox UserName & " session is not registered in the system, so your operation cannot be started. Please register your session in the system using the Initials Interface located in the Settings Group of the Enterprise Document Automation System menu and try the operation again.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo ModulSonu
    End If
    
    On Error GoTo ModulSonu
    SayPrt = InputBox(Prompt:="Enter the number of copies you want to print.", Title:="Enterprise Document Automation System")
    If SayPrt = "" Or SayPrt = 0 Then
        GoTo ModulSonu
    End If
    If IsNumeric(SayPrt) = False Then
        MsgBox "Non-numeric input detected, so the print operation could not be started.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo ModulSonu
    End If
    If SayPrt > 3 Then
        MsgBox "The number of copies to be printed at one time cannot exceed 3.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo ModulSonu
    End If
    
    'All için TipB/TipA ayrımı
    If Cells(ActiveCell.Row, 28) = "Type A" Then
        Call ModuleReport3.Rapor3_2Tutanak
        Call ModuleReport3.Rapor3_2Rapor
        Call ModuleReport3.Rapor3_2Tutanak2
        Call ModuleReport3.Rapor3_2FinansalBirimUstYazi
        Call ModuleReport3.Rapor3_2UstYazi
    ElseIf Cells(ActiveCell.Row, 28) = "Type B" Then
        Call ModuleReport3.Rapor3_2TutanakTipB
        Call ModuleReport3.Rapor3_2Tutanak2TipB
        Call ModuleReport3.Rapor3_2FinansalBirimUstYaziTipB
        Call ModuleReport3.Rapor3_2UstYaziTipB
    End If
End If

ModulSonu:
ThisWorkbook.Worksheets(5).Activate


End Sub

Public Function IsLoaded(formName As String) As Boolean
Dim frm As Object
For Each frm In VBA.UserForms
    If frm.name = formName Then
        IsLoaded = True
        Exit Function
    End If
Next frm
IsLoaded = False
End Function

Sub UserFormlariKapat()

If IsLoaded("core_report1_entry_UI") = True Then
    Unload core_report1_entry_UI
End If
If IsLoaded("core_report2_entry_UI") = True Then
    Unload core_report2_entry_UI
End If
If IsLoaded("core_report3_1_entry_UI") = True Then
    Unload core_report3_1_entry_UI
End If
If IsLoaded("core_report3_2_entry_UI") = True Then
    Unload core_report3_2_entry_UI
End If
If IsLoaded("core_asset_manager_UI") = True Then
    Unload core_asset_manager_UI
End If
If IsLoaded("core_acceptance_manager_UI") = True Then
    Unload core_acceptance_manager_UI
End If
If IsLoaded("core_delivery_manager_UI") = True Then
    Unload core_delivery_manager_UI
End If
If IsLoaded("core_performance_report_UI") = True Then
    Unload core_performance_report_UI
End If
If IsLoaded("core_unit_settings_UI") = True Then
    Unload core_unit_settings_UI
End If
If IsLoaded("core_auto_close_settings_UI") = True Then
    Unload core_auto_close_settings_UI
End If
If IsLoaded("core_initials_UI") = True Then
    Unload core_initials_UI
End If
If IsLoaded("core_system_reset_wizard2_UI") = True Then
    Unload core_system_reset_wizard2_UI
End If
If IsLoaded("core_system_reset_wizard1_UI") = True Then
    Unload core_system_reset_wizard1_UI
End If


End Sub


