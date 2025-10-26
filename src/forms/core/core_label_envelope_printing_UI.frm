VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} core_label_envelope_printing_UI 
   Caption         =   "Enterprise Document Automation System"
   ClientHeight    =   8895
   ClientLeft      =   30
   ClientTop       =   375
   ClientWidth     =   13920
   OleObjectBlob   =   "core_label_envelope_printing_UI.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "core_label_envelope_printing_UI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Public OpenWordTakip As Boolean

Dim Abort As Boolean

Dim Bolum1 As String, Bolum3 As String
Dim IlBuyukHarf As String, IlKucukHarf As String
Dim IlceBuyukHarf As String, IlceKucukHarf As String
Dim M2 As Boolean, M3 As Boolean, M4 As Boolean
Dim IlceSakla As String
Dim BoolKucukZarf As Boolean, BoolBuyukZarf As Boolean
Dim BoolFinansalBirimKucukZarf As Boolean, BoolFinansalBirimBuyukZarf As Boolean
Dim WsEtiket As Worksheet, PrtNo As Integer



Private Sub Yazdir_Click()
Dim i As Long, j As Long


Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.EnableEvents = False

ThisWorkbook.Unprotect "123"
ThisWorkbook.Worksheets(12).Unprotect Password:="123"
ThisWorkbook.Worksheets(13).Unprotect Password:="123"
ThisWorkbook.Worksheets(14).Unprotect Password:="123"
ThisWorkbook.Worksheets(15).Unprotect Password:="123"

ThisWorkbook.Worksheets(12).Visible = True
ThisWorkbook.Worksheets(13).Visible = True
ThisWorkbook.Worksheets(14).Visible = True
ThisWorkbook.Worksheets(15).Visible = True

If EtiketPrinterOption.Value = False And PikurPrinterOption.Value = False Then
    MsgBox "Please select the printer from the options on the side.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

Set WsEtiket = Nothing

PrtNo = CInt(ThisWorkbook.Worksheets(2).Cells(8, 115).Value)

'ETİKET YAZICI
If EtiketPrinterOption.Value = True Then
    Set WsEtiket = ThisWorkbook.Worksheets(12)
    WsEtiket.Rows("1:30").EntireRow.Delete
    'WsEtiket.UsedRange.ClearContents

    j = 1
    If TextKURUM_A.Value <> "" Then
        WsEtiket.Cells(j, 1).Value = TextKURUM_A.Value
        j = j + 1
    End If
    If TextSube.Value <> "" Then
        WsEtiket.Cells(j, 1).Value = TextSube.Value
        j = j + 1
    End If
    If TextTeskilatDosyaNoBelgeNo.Value <> "" Then
        WsEtiket.Cells(j, 1).Value = TextTeskilatDosyaNoBelgeNo.Value
        j = j + 1
    End If

    If TextGonderiSekli.Value <> "" Then
        WsEtiket.Rows(j).RowHeight = 40
        WsEtiket.Range("A" & j).Font.Size = 20
        WsEtiket.Cells(j, 1).Value = TextGonderiSekli.Value
        j = j + 1
    End If

    'Gönderi tipi boş ise 1 satır boşluk bırak
    If TextGonderiSekli.Value = "" Then
        j = j + 1
    End If

    For i = 1 To 7
        If Controls("TextMuhatapTema" & i).Value <> "" Then
            WsEtiket.Cells(j, 1).Value = Controls("TextMuhatapTema" & i).Value
            j = j + 1
        End If
    Next i

    If j <= 2 Then
        GoTo Son
    End If

    GoTo Yazdir

End If

'PİKÜR YAZICI
If PikurPrinterOption.Value = True Then

    If LblAltMenuButton1.BackColor = RGB(180, 210, 240) Then 'tutanak2 zarfı
        GoTo Tutanak2ZarfinaGit
    ElseIf LblAltMenuButton2.BackColor = RGB(180, 210, 240) Or LblAltMenuButton6.BackColor = RGB(180, 210, 240) Then 'Küçük Zarf
        GoTo KucukZarfaGit
    ElseIf LblAltMenuButton3.BackColor = RGB(180, 210, 240) Or LblAltMenuButton7.BackColor = RGB(180, 210, 240) Then 'Büyük Zarf
        GoTo BuyukZarfaGit
    End If

    If Rapor3_2.BackColor = RGB(180, 210, 240) Then
        If LblAltMenuButton4.BackColor = RGB(180, 210, 240) Then 'Küçük Zarf
            GoTo KucukZarfaGit
        ElseIf LblAltMenuButton5.BackColor = RGB(180, 210, 240) Then 'Büyük Zarf
            GoTo BuyukZarfaGit
        End If
    End If
    
    If Rapor2_2.BackColor = RGB(180, 210, 240) Then
        If LblAltMenuButton4.BackColor = RGB(180, 210, 240) Or LblAltMenuButton5.BackColor = RGB(180, 210, 240) Then 'Tutanak2 zarfı
            GoTo Tutanak2ZarfinaGit
        End If
    End If
    
End If

GoTo Son

'________________________

Tutanak2ZarfinaGit:
'Tutanak2 zarfı-ayar
Set WsEtiket = ThisWorkbook.Worksheets(13)
WsEtiket.Rows("1:100").EntireRow.Delete
'WsEtiket.UsedRange.ClearContents
j = 3
If TextKURUM_A.Value <> "" Then
    WsEtiket.Cells(j, 1).Value = TextKURUM_A.Value
    j = j + 1
End If
If TextSube.Value <> "" Then
    WsEtiket.Cells(j, 1).Value = TextSube.Value
    j = j + 1
End If
If TextTeskilatDosyaNoBelgeNo.Value <> "" Then
    WsEtiket.Cells(j, 1).Value = TextTeskilatDosyaNoBelgeNo.Value
    j = j + 1
End If

If TextGonderiSekli.Value <> "" Then
    WsEtiket.Rows(j).RowHeight = 40
    WsEtiket.Range("A" & j).Font.Size = 20
    WsEtiket.Cells(j, 1).Value = TextGonderiSekli.Value
    j = j + 1
End If

'Gönderi tipi boş ise 1 satır boşluk bırak
If TextGonderiSekli.Value = "" Then
    j = j + 1
End If

'j = j + 3


For i = 1 To 7
    If Controls("TextMuhatapTema" & i).Value <> "" Then
        WsEtiket.Cells(j, 1).Value = Controls("TextMuhatapTema" & i).Value
        j = j + 1
    End If
Next i
 
If j <= 4 Then
    GoTo Son
End If

GoTo Yazdir

'________________________


KucukZarfaGit:
'Küçük Zarf-ayar
Set WsEtiket = ThisWorkbook.Worksheets(14)
WsEtiket.Rows("1:100").EntireRow.Delete
'WsEtiket.UsedRange.ClearContents
j = 7
If TextKURUM_A.Value <> "" Then
    WsEtiket.Cells(j, 1).Value = TextKURUM_A.Value
    j = j + 1
End If
If TextSube.Value <> "" Then
    WsEtiket.Cells(j, 1).Value = TextSube.Value
    j = j + 1
End If
If TextTeskilatDosyaNoBelgeNo.Value <> "" Then
    WsEtiket.Cells(j, 1).Value = TextTeskilatDosyaNoBelgeNo.Value
    j = j + 1
End If

If TextGonderiSekli.Value <> "" Then
    WsEtiket.Rows(j).RowHeight = 40
    WsEtiket.Range("E" & j).Font.Size = 20
    WsEtiket.Cells(j, 5).Value = TextGonderiSekli.Value
    j = j + 1
End If

''Gönderi tipi boş ise 1 satır boşluk bırak
'If TextGonderiSekli.Value = "" Then
'    j = j + 1
'End If

j = 22
For i = 1 To 7
    If Controls("TextMuhatapTema" & i).Value <> "" Then
        WsEtiket.Cells(j, 5).Value = Controls("TextMuhatapTema" & i).Value
        j = j + 1
    End If
Next i
 
If j <= 2 Then
    GoTo Son
End If

GoTo Yazdir



'________________________


BuyukZarfaGit:
'Büyük Zarf-ayar
Set WsEtiket = ThisWorkbook.Worksheets(15)
WsEtiket.Rows("1:100").EntireRow.Delete
'WsEtiket.UsedRange.ClearContents
j = 1
If TextKURUM_A.Value <> "" Then
    WsEtiket.Cells(j, 2).Value = TextKURUM_A.Value
    j = j + 1
End If
If TextSube.Value <> "" Then
    WsEtiket.Cells(j, 2).Value = TextSube.Value
    j = j + 1
End If
If TextTeskilatDosyaNoBelgeNo.Value <> "" Then
    WsEtiket.Cells(j, 2).Value = TextTeskilatDosyaNoBelgeNo.Value
    j = j + 1
End If

If TextGonderiSekli.Value <> "" Then
    WsEtiket.Rows(j).RowHeight = 40
    WsEtiket.Range("E" & j).Font.Size = 20
    WsEtiket.Cells(j, 5).Value = TextGonderiSekli.Value
    j = j + 1
End If

''Gönderi tipi boş ise 1 satır boşluk bırak
'If TextGonderiSekli.Value = "" Then
'    j = j + 1
'End If

j = 22
For i = 1 To 7
    If Controls("TextMuhatapTema" & i).Value <> "" Then
        WsEtiket.Cells(j, 5).Value = Controls("TextMuhatapTema" & i).Value
        j = j + 1
    End If
Next i
 
If j <= 2 Then
    GoTo Son
End If

GoTo Yazdir

'________________________


Yazdir:
Call MyOptionPrinter


Son:

'ThisWorkbook.Worksheets(12).Activate


ThisWorkbook.Worksheets(12).Visible = False
ThisWorkbook.Worksheets(13).Visible = False
ThisWorkbook.Worksheets(14).Visible = False
ThisWorkbook.Worksheets(15).Visible = False

ThisWorkbook.Worksheets(12).Protect Password:="123"
ThisWorkbook.Worksheets(13).Protect Password:="123"
ThisWorkbook.Worksheets(14).Protect Password:="123"
ThisWorkbook.Worksheets(15).Protect Password:="123"
ThisWorkbook.Protect "123"


ActiveSheet.DisplayPageBreaks = False

Application.ScreenUpdating = True
Application.DisplayAlerts = True
Application.EnableEvents = True

End Sub

Sub MyOptionPrinter()
Dim MyCurrentPrinter As String
Dim WindowsPrinterName As String
Dim MyPrinterName As String
Dim WsEtiketx As Worksheet

Set WsEtiketx = WsEtiket 'ThisWorkbook.Worksheets(12)
MyCurrentPrinter = Application.ActivePrinter

'WindowsPrinterName = "MP 4000"
WindowsPrinterName = ThisWorkbook.Worksheets(2).Cells(PrtNo, 115).Value
'Get full name of specified Windows printer, including its network port
MyPrinterName = FindPrinter(WindowsPrinterName)
Application.ActivePrinter = MyPrinterName
WsEtiketx.PrintOut ActivePrinter:=MyPrinterName, Copies:=1 ', Collate:=True, From:=1, To:=1, IgnorePrintAreas:=False

Application.ActivePrinter = MyCurrentPrinter

ActiveSheet.DisplayPageBreaks = False

End Sub


Private Sub EtiketPrinterOption_Click()
Dim StrPrinters As Variant, i As Long
Dim KullanilanYazici As String, prtDogrula As Boolean

Dim AutoPath As String, DestTarget As String, OpenControl As String
Dim FileName As String


Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.EnableEvents = False

AutoPath = ThisWorkbook.Path
DestTarget = AutoPath & "\System Files\System Definitions\"
FileName = "Definitions.xlsx"

OpenControl = IsFileOpen(DestTarget & FileName)
If OpenControl = True Then 'Açıksa
    Workbooks(FileName).Close SaveChanges:=True
ElseIf OpenControl = False Then
    '
End If

Workbooks.Open (DestTarget & FileName)
Workbooks(FileName).Worksheets(1).Activate

'______________________

prtDogrula = False

If EtiketPrinterOption.Value = True Then

    KullanilanYazici = ThisWorkbook.Worksheets(2).Cells(6, 115).Value
    If KullanilanYazici = "" Then
        EtiketPrinterOption.Value = False
        Workbooks(FileName).Worksheets(1).Cells(8, 115).Value = ""
        ThisWorkbook.Worksheets(2).Cells(8, 115).Value = ""
        MsgBox "No LABEL PRINTER has been detected in the Enterprise Document Automation System. If a label printer is assigned on your computer, you can define it in the system by clicking the 'Assign Label Printer' button.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    
    '________________________
    
    StrPrinters = ListPrinters
    'Fist check whether the array is filled with anything, by calling another function, IsBounded.
    If IsBounded(StrPrinters) Then
        For i = LBound(StrPrinters) To UBound(StrPrinters)
            If KullanilanYazici = StrPrinters(i) Then
                prtDogrula = True
            End If
        Next i
    Else
        EtiketPrinterOption.Value = False
        Workbooks(FileName).Worksheets(1).Cells(8, 115).Value = ""
        ThisWorkbook.Worksheets(2).Cells(8, 115).Value = ""
        MsgBox "No printer has been detected on your computer.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    End If
    
    '________________________
    
    
    If prtDogrula = True Then
        '
        Workbooks(FileName).Worksheets(1).Cells(8, 115).Value = 6
        ThisWorkbook.Worksheets(2).Cells(8, 115).Value = 6
    Else
        EtiketPrinterOption.Value = False
        Workbooks(FileName).Worksheets(1).Cells(8, 115).Value = ""
        ThisWorkbook.Worksheets(2).Cells(8, 115).Value = ""
        MsgBox "No LABEL PRINTER has been detected in the Enterprise Document Automation System. If a label printer is assigned on your computer, you can define it in the system by clicking the 'Assign Label Printer' button.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    End If
    
End If

Son:

OpenControl = IsFileOpen(DestTarget & FileName)
If OpenControl = True Then 'Açıksa
    Workbooks(FileName).Close SaveChanges:=True
ElseIf OpenControl = False Then
    '
End If

ThisWorkbook.Activate

ActiveSheet.DisplayPageBreaks = False

Application.ScreenUpdating = True
Application.DisplayAlerts = True
Application.EnableEvents = True


End Sub

Private Sub PikurPrinterOption_Click()

Dim StrPrinters As Variant, i As Long
Dim KullanilanYazici As String, prtDogrula As Boolean

Dim AutoPath As String, DestTarget As String, OpenControl As String
Dim FileName As String


Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.EnableEvents = False

AutoPath = ThisWorkbook.Path
DestTarget = AutoPath & "\System Files\System Definitions\"
FileName = "Definitions.xlsx"

OpenControl = IsFileOpen(DestTarget & FileName)
If OpenControl = True Then 'Açıksa
    Workbooks(FileName).Close SaveChanges:=True
ElseIf OpenControl = False Then
    '
End If

Workbooks.Open (DestTarget & FileName)
Workbooks(FileName).Worksheets(1).Activate

'______________________

prtDogrula = False

If PikurPrinterOption.Value = True Then

    KullanilanYazici = ThisWorkbook.Worksheets(2).Cells(7, 115).Value
    If KullanilanYazici = "" Then
        PikurPrinterOption.Value = False
        Workbooks(FileName).Worksheets(1).Cells(8, 115).Value = ""
        ThisWorkbook.Worksheets(2).Cells(8, 115).Value = ""
        MsgBox "No RECEIPT PRINTER has been detected in the Enterprise Document Automation System. If a receipt printer is assigned on your computer, you can define it in the system by clicking the 'Assign Receipt Printer' button.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    
    StrPrinters = ListPrinters
    'Fist check whether the array is filled with anything, by calling another function, IsBounded.
    If IsBounded(StrPrinters) Then
        For i = LBound(StrPrinters) To UBound(StrPrinters)
            If KullanilanYazici = StrPrinters(i) Then
                prtDogrula = True
            End If
        Next i
    Else
        PikurPrinterOption.Value = False
        Workbooks(FileName).Worksheets(1).Cells(8, 115).Value = ""
        ThisWorkbook.Worksheets(2).Cells(8, 115).Value = ""
        MsgBox "No printer has been detected on your computer.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    End If
    
    If prtDogrula = True Then
        '
        Workbooks(FileName).Worksheets(1).Cells(8, 115).Value = 7
        ThisWorkbook.Worksheets(2).Cells(8, 115).Value = 7
    Else
        PikurPrinterOption.Value = False
        Workbooks(FileName).Worksheets(1).Cells(8, 115).Value = ""
        ThisWorkbook.Worksheets(2).Cells(8, 115).Value = ""
        MsgBox "No RECEIPT PRINTER has been detected in the Enterprise Document Automation System. If a receipt printer is assigned on your computer, you can define it in the system by clicking the 'Assign Receipt Printer' button.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    End If
    
End If

Son:

OpenControl = IsFileOpen(DestTarget & FileName)
If OpenControl = True Then 'Açıksa
    Workbooks(FileName).Close SaveChanges:=True
ElseIf OpenControl = False Then
    '
End If

ThisWorkbook.Activate

ActiveSheet.DisplayPageBreaks = False

Application.ScreenUpdating = True
Application.DisplayAlerts = True
Application.EnableEvents = True


End Sub

Private Sub EtiketTanimla_Click()
Dim Bilgi As Variant
Dim MyCurrentPrinter As String
Dim onlyPrinterName As String

Dim AutoPath As String, DestTarget As String, OpenControl As String
Dim FileName As String


Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.EnableEvents = False

AutoPath = ThisWorkbook.Path
DestTarget = AutoPath & "\System Files\System Definitions\"
FileName = "Definitions.xlsx"

OpenControl = IsFileOpen(DestTarget & FileName)
If OpenControl = True Then 'Açıksa
    Workbooks(FileName).Close SaveChanges:=True
ElseIf OpenControl = False Then
    '
End If

'_______________

MyCurrentPrinter = Application.ActivePrinter

'Application.Dialogs(xlDialogPrinterSetup).Show
If Application.Dialogs(xlDialogPrinterSetup).Show = False Then GoTo Son

onlyPrinterName = Application.ActivePrinter
onlyPrinterName = Replace(Application.ActivePrinter, " üzerindeki ", " on ")
onlyPrinterName = Left(onlyPrinterName, InStr(onlyPrinterName, " on ") - 1)


Bilgi = MsgBox("Click '" & "Yes" & "' to assign the printer named '" & onlyPrinterName & "' as the LABEL PRINTER, or click '" & "No" & "' to cancel the operation.", vbYesNo + vbInformation, "Enterprise Document Automation System")
If Bilgi = vbNo Then
    GoTo Son
End If

'_________________

Workbooks.Open (DestTarget & FileName)
Workbooks(FileName).Worksheets(1).Activate
Workbooks(FileName).Worksheets(1).Cells(6, 115).Value = onlyPrinterName
Workbooks(FileName).Save
OpenControl = IsFileOpen(DestTarget & FileName)
If OpenControl = True Then 'Açıksa
    Workbooks(FileName).Close SaveChanges:=False
ElseIf OpenControl = False Then
    '
End If


ThisWorkbook.Activate
ThisWorkbook.Worksheets(2).Cells(6, 115).Value = onlyPrinterName
ActiveSheet.DisplayPageBreaks = False
MsgBox "The printer named """ & onlyPrinterName & """ has been successfully assigned as the LABEL PRINTER.", vbOKOnly + vbInformation, "Enterprise Document Automation System"

Son:

OpenControl = IsFileOpen(DestTarget & FileName)
If OpenControl = True Then 'Açıksa
    Workbooks(FileName).Close SaveChanges:=True
ElseIf OpenControl = False Then
    '
End If

ThisWorkbook.Activate

Application.ActivePrinter = MyCurrentPrinter 'varsayılan aktif yazıcıya geri dön

ActiveSheet.DisplayPageBreaks = False

Application.ScreenUpdating = True
Application.DisplayAlerts = True
Application.EnableEvents = True

End Sub


Private Sub PikurTanimla_Click()
Dim Bilgi As Variant
Dim MyCurrentPrinter As String
Dim onlyPrinterName As String

Dim AutoPath As String, DestTarget As String, OpenControl As String
Dim FileName As String


Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.EnableEvents = False

AutoPath = ThisWorkbook.Path
DestTarget = AutoPath & "\System Files\System Definitions\"
FileName = "Definitions.xlsx"

OpenControl = IsFileOpen(DestTarget & FileName)
If OpenControl = True Then 'Açıksa
    Workbooks(FileName).Close SaveChanges:=True
ElseIf OpenControl = False Then
    '
End If

'_______________

MyCurrentPrinter = Application.ActivePrinter

'Application.Dialogs(xlDialogPrinterSetup).Show
If Application.Dialogs(xlDialogPrinterSetup).Show = False Then GoTo Son

onlyPrinterName = Application.ActivePrinter
onlyPrinterName = Replace(Application.ActivePrinter, " üzerindeki ", " on ")
onlyPrinterName = Left(onlyPrinterName, InStr(onlyPrinterName, " on ") - 1)

Bilgi = MsgBox("Click '" & "Yes" & "' to assign the printer named '" & onlyPrinterName & "' as the RECEIPT PRINTER, or click '" & "No" & "' to cancel the operation.", vbYesNo + vbInformation, "Enterprise Document Automation System")
If Bilgi = vbNo Then
    GoTo Son
End If

'_________________


Workbooks.Open (DestTarget & FileName)
Workbooks(FileName).Worksheets(1).Activate
Workbooks(FileName).Worksheets(1).Cells(7, 115).Value = onlyPrinterName
Workbooks(FileName).Save
OpenControl = IsFileOpen(DestTarget & FileName)
If OpenControl = True Then 'Açıksa
    Workbooks(FileName).Close SaveChanges:=False
ElseIf OpenControl = False Then
    '
End If


ThisWorkbook.Activate
ThisWorkbook.Worksheets(2).Cells(7, 115).Value = onlyPrinterName
ActiveSheet.DisplayPageBreaks = False
MsgBox "The printer named """ & onlyPrinterName & """ has been successfully assigned as the RECEIPT PRINTER.", vbOKOnly + vbInformation, "Enterprise Document Automation System"

Son:

OpenControl = IsFileOpen(DestTarget & FileName)
If OpenControl = True Then 'Açıksa
    Workbooks(FileName).Close SaveChanges:=True
ElseIf OpenControl = False Then
    '
End If

ThisWorkbook.Activate

Application.ActivePrinter = MyCurrentPrinter 'varsayılan aktif yazıcıya geri dön

ActiveSheet.DisplayPageBreaks = False

Application.ScreenUpdating = True
Application.DisplayAlerts = True
Application.EnableEvents = True

End Sub


Sub ColorChangerGenel()

If Yazdir.BackColor <> RGB(225, 235, 245) Then
Yazdir.BackColor = RGB(225, 235, 245)
Yazdir.ForeColor = RGB(30, 30, 30)
End If

If Kapat.BackColor <> RGB(225, 235, 245) Then
Kapat.BackColor = RGB(225, 235, 245)
Kapat.ForeColor = RGB(30, 30, 30)
End If

'Rapor1
If Rapor1.BackColor <> RGB(180, 210, 240) Then
    If Rapor1.BackColor <> RGB(225, 235, 245) Then
        Rapor1.BackColor = RGB(225, 235, 245)
        Rapor1.ForeColor = RGB(30, 30, 30)
    End If
End If

'Rapor
If Rapor.BackColor <> RGB(180, 210, 240) Then
    If Rapor.BackColor <> RGB(225, 235, 245) Then
        Rapor.BackColor = RGB(225, 235, 245)
        Rapor.ForeColor = RGB(30, 30, 30)
    End If
End If

'Rapor2_2
If Rapor2_2.BackColor <> RGB(180, 210, 240) Then
    If Rapor2_2.BackColor <> RGB(225, 235, 245) Then
        Rapor2_2.BackColor = RGB(225, 235, 245)
        Rapor2_2.ForeColor = RGB(30, 30, 30)
    End If
End If


'Rapor3_2
If Rapor3_2.BackColor <> RGB(180, 210, 240) Then
    If Rapor3_2.BackColor <> RGB(225, 235, 245) Then
        Rapor3_2.BackColor = RGB(225, 235, 245)
        Rapor3_2.ForeColor = RGB(30, 30, 30)
    End If
End If

'Rapor3_1
If Rapor3_1.BackColor <> RGB(180, 210, 240) Then
    If Rapor3_1.BackColor <> RGB(225, 235, 245) Then
        Rapor3_1.BackColor = RGB(225, 235, 245)
        Rapor3_1.ForeColor = RGB(30, 30, 30)
    End If
End If

'LblAltMenuButton1
If LblAltMenuButton1.BackColor <> RGB(180, 210, 240) Then
    If LblAltMenuButton1.BackColor <> RGB(225, 235, 245) Then
        LblAltMenuButton1.BackColor = RGB(225, 235, 245)
        LblAltMenuButton1.ForeColor = RGB(30, 30, 30)
    End If
End If

'LblAltMenuButton2
If LblAltMenuButton2.BackColor <> RGB(180, 210, 240) Then
    If LblAltMenuButton2.BackColor <> RGB(225, 235, 245) Then
        LblAltMenuButton2.BackColor = RGB(225, 235, 245)
        LblAltMenuButton2.ForeColor = RGB(30, 30, 30)
    End If
End If

'LblAltMenuButton3
If LblAltMenuButton3.BackColor <> RGB(180, 210, 240) Then
    If LblAltMenuButton3.BackColor <> RGB(225, 235, 245) Then
        LblAltMenuButton3.BackColor = RGB(225, 235, 245)
        LblAltMenuButton3.ForeColor = RGB(30, 30, 30)
    End If
End If

'LblAltMenuButton4
If LblAltMenuButton4.BackColor <> RGB(180, 210, 240) Then
    If LblAltMenuButton4.BackColor <> RGB(225, 235, 245) Then
        LblAltMenuButton4.BackColor = RGB(225, 235, 245)
        LblAltMenuButton4.ForeColor = RGB(30, 30, 30)
    End If
End If

'LblAltMenuButton5
If LblAltMenuButton5.BackColor <> RGB(180, 210, 240) Then
    If LblAltMenuButton5.BackColor <> RGB(225, 235, 245) Then
        LblAltMenuButton5.BackColor = RGB(225, 235, 245)
        LblAltMenuButton5.ForeColor = RGB(30, 30, 30)
    End If
End If

'LblAltMenuButton6
If LblAltMenuButton6.BackColor <> RGB(180, 210, 240) Then
    If LblAltMenuButton6.BackColor <> RGB(225, 235, 245) Then
        LblAltMenuButton6.BackColor = RGB(225, 235, 245)
        LblAltMenuButton6.ForeColor = RGB(30, 30, 30)
    End If
End If

'LblAltMenuButton7
If LblAltMenuButton7.BackColor <> RGB(180, 210, 240) Then
    If LblAltMenuButton7.BackColor <> RGB(225, 235, 245) Then
        LblAltMenuButton7.BackColor = RGB(225, 235, 245)
        LblAltMenuButton7.ForeColor = RGB(30, 30, 30)
    End If
End If

'PikurPrinterOption
If PikurPrinterOption.BackColor <> RGB(225, 235, 245) Then
PikurPrinterOption.BackColor = RGB(225, 235, 245)
PikurPrinterOption.ForeColor = RGB(30, 30, 30)
End If

'EtiketPrinterOption
If EtiketPrinterOption.BackColor <> RGB(225, 235, 245) Then
EtiketPrinterOption.BackColor = RGB(225, 235, 245)
EtiketPrinterOption.ForeColor = RGB(30, 30, 30)
End If

If PikurTanimla.BackColor <> RGB(225, 235, 245) Then
PikurTanimla.BackColor = RGB(225, 235, 245)
PikurTanimla.ForeColor = RGB(30, 30, 30)
End If

If EtiketTanimla.BackColor <> RGB(225, 235, 245) Then
EtiketTanimla.BackColor = RGB(225, 235, 245)
EtiketTanimla.ForeColor = RGB(30, 30, 30)
End If


End Sub

Private Sub Yazdir_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
Yazdir.BackColor = RGB(60, 100, 180) 'RGB(60, 100, 180)
Yazdir.ForeColor = RGB(255, 255, 255)
End Sub

Private Sub Kapat_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
Kapat.BackColor = RGB(60, 100, 180) 'RGB(60, 100, 180)
Kapat.ForeColor = RGB(255, 255, 255)
End Sub

Private Sub PikurTanimla_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
PikurTanimla.BackColor = RGB(60, 100, 180) 'RGB(60, 100, 180)
PikurTanimla.ForeColor = RGB(255, 255, 255)
End Sub

Private Sub EtiketTanimla_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
EtiketTanimla.BackColor = RGB(60, 100, 180) 'RGB(60, 100, 180)
EtiketTanimla.ForeColor = RGB(255, 255, 255)
End Sub


Private Sub EtiketPrinterOption_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
EtiketPrinterOption.BackColor = RGB(60, 100, 180) 'RGB(60, 100, 180)
EtiketPrinterOption.ForeColor = RGB(255, 255, 255)
End Sub

Private Sub PikurPrinterOption_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
PikurPrinterOption.BackColor = RGB(60, 100, 180) 'RGB(60, 100, 180)
PikurPrinterOption.ForeColor = RGB(255, 255, 255)
End Sub

Private Sub Rapor1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
If Rapor1.BackColor <> RGB(180, 210, 240) Then
    Rapor1.BackColor = RGB(60, 100, 180) 'RGB(60, 100, 180)
    Rapor1.ForeColor = RGB(255, 255, 255)
End If
End Sub

Private Sub Rapor_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
If Rapor.BackColor <> RGB(180, 210, 240) Then
    Rapor.BackColor = RGB(60, 100, 180) 'RGB(60, 100, 180)
    Rapor.ForeColor = RGB(255, 255, 255)
End If
End Sub


Private Sub Rapor2_2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
If Rapor2_2.BackColor <> RGB(180, 210, 240) Then
    Rapor2_2.BackColor = RGB(60, 100, 180) 'RGB(60, 100, 180)
    Rapor2_2.ForeColor = RGB(255, 255, 255)
End If
End Sub


Private Sub Rapor3_2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
If Rapor3_2.BackColor <> RGB(180, 210, 240) Then
    Rapor3_2.BackColor = RGB(60, 100, 180) 'RGB(60, 100, 180)
    Rapor3_2.ForeColor = RGB(255, 255, 255)
End If
End Sub

Private Sub Rapor3_1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
If Rapor3_1.BackColor <> RGB(180, 210, 240) Then
    Rapor3_1.BackColor = RGB(60, 100, 180) 'RGB(60, 100, 180)
    Rapor3_1.ForeColor = RGB(255, 255, 255)
End If
End Sub

Private Sub LblAltMenuButton1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
If LblAltMenuButton1.BackColor <> RGB(180, 210, 240) And LblAltMenuButton1.Caption <> "" Then
    LblAltMenuButton1.BackColor = RGB(60, 100, 180) 'RGB(60, 100, 180)
    LblAltMenuButton1.ForeColor = RGB(255, 255, 255)
End If
End Sub

Private Sub LblAltMenuButton2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
If LblAltMenuButton2.BackColor <> RGB(180, 210, 240) And LblAltMenuButton2.Caption <> "" Then
    LblAltMenuButton2.BackColor = RGB(60, 100, 180) 'RGB(60, 100, 180)
    LblAltMenuButton2.ForeColor = RGB(255, 255, 255)
End If
End Sub

Private Sub LblAltMenuButton3_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
If LblAltMenuButton3.BackColor <> RGB(180, 210, 240) And LblAltMenuButton3.Caption <> "" Then
    LblAltMenuButton3.BackColor = RGB(60, 100, 180) 'RGB(60, 100, 180)
    LblAltMenuButton3.ForeColor = RGB(255, 255, 255)
End If
End Sub

Private Sub LblAltMenuButton4_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
If LblAltMenuButton4.BackColor <> RGB(180, 210, 240) And LblAltMenuButton4.Caption <> "" Then
    LblAltMenuButton4.BackColor = RGB(60, 100, 180) 'RGB(60, 100, 180)
    LblAltMenuButton4.ForeColor = RGB(255, 255, 255)
End If
End Sub

Private Sub LblAltMenuButton5_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
If LblAltMenuButton5.BackColor <> RGB(180, 210, 240) And LblAltMenuButton5.Caption <> "" Then
    LblAltMenuButton5.BackColor = RGB(60, 100, 180) 'RGB(60, 100, 180)
    LblAltMenuButton5.ForeColor = RGB(255, 255, 255)
End If
End Sub

Private Sub LblAltMenuButton6_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
If LblAltMenuButton6.BackColor <> RGB(180, 210, 240) And LblAltMenuButton6.Caption <> "" Then
    LblAltMenuButton6.BackColor = RGB(60, 100, 180) 'RGB(60, 100, 180)
    LblAltMenuButton6.ForeColor = RGB(255, 255, 255)
End If
End Sub

Private Sub LblAltMenuButton7_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
If LblAltMenuButton7.BackColor <> RGB(180, 210, 240) And LblAltMenuButton7.Caption <> "" Then
    LblAltMenuButton7.BackColor = RGB(60, 100, 180) 'RGB(60, 100, 180)
    LblAltMenuButton7.ForeColor = RGB(255, 255, 255)
End If
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
Private Sub UstMenuFrameAlt_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub
Private Sub BilgilendirmeFrame_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub
Private Sub AltMenuFrame_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub

'ComboGetir
Private Sub ComboGetir_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
Call SetComboBoxHook(ComboGetir) 'Open scrollable with mouse
End Sub
Private Sub ComboGetir_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Me.ComboGetir.DropDown
End Sub
Private Sub ComboGetir_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    If KeyCode = vbKeyTab Then
        TextKURUM_A.SetFocus
    End If
    If KeyCode = vbKeyDown Then
        TextKURUM_A.SetFocus
    End If
    
    Select Case KeyCode
        Case 38  'Up
            If ComboGetir.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                ComboGetir.ListIndex = ComboGetir.ListIndex
            End If
        Case 40 'Down
            If ComboGetir.ListIndex >= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                ComboGetir.ListIndex = ComboGetir.ListIndex
            End If
    End Select
    Abort = False
    
End Sub


Private Sub Kapat_Click()
    Unload Me
End Sub

Private Sub Rapor1_Click()
Dim ws As Object, i As Integer
Dim ctl As MSForms.Control

Application.DisplayAlerts = False
Application.EnableEvents = False
Application.ScreenUpdating = False

ThisWorkbook.Unprotect "123"

i = 0
ThisWorkbook.Worksheets(3).Visible = True
For Each ws In ThisWorkbook.Worksheets
    i = i + 1
    If i <> 3 Then
        ws.Visible = False
    End If
Next ws
ThisWorkbook.Worksheets(3).Activate

With ActiveWindow
    .DisplayHorizontalScrollBar = True
    .DisplayVerticalScrollBar = True
End With

Call ComboGetirResetRapor1
Call EtiketAlanlariReset


For Each ctl In UstMenuFrame.Controls
    If TypeName(ctl) = "Label" Then
        ctl.BackColor = RGB(225, 235, 245)  'RGB(254, 254, 254)
        ctl.ForeColor = RGB(30, 30, 30)
    End If
Next ctl
Rapor1.BackColor = RGB(180, 210, 240)
Rapor1.ForeColor = RGB(30, 30, 30)

'_________________________

For Each ctl In UstMenuFrameAlt.Controls
    If TypeName(ctl) = "Label" Then
        ctl.BackColor = RGB(225, 235, 245)  'RGB(254, 254, 254)
        ctl.ForeColor = RGB(30, 30, 30)
        ctl.Caption = ""
    End If
Next ctl

LblAltMenuButton1.Caption = "R1 - Env. 1" ' Statement 2 Env.
LblAltMenuButton1.ForeColor = RGB(30, 30, 30)

LblAltMenuButton2.Caption = "R1 - Env. 2" 'Small Env.
LblAltMenuButton2.ForeColor = RGB(30, 30, 30)

LblAltMenuButton3.Caption = "R1 - Env. 3" 'Large Env.
LblAltMenuButton3.ForeColor = RGB(30, 30, 30)



ThisWorkbook.Protect "123"

Application.DisplayAlerts = True
Application.EnableEvents = True
Application.ScreenUpdating = True


End Sub

Private Sub Rapor_Click()
Dim ws As Object, i As Integer
Dim ctl As MSForms.Control

Application.DisplayAlerts = False
Application.EnableEvents = False
Application.ScreenUpdating = False

ThisWorkbook.Unprotect "123"

i = 0
ThisWorkbook.Worksheets(4).Visible = True
For Each ws In ThisWorkbook.Worksheets
    i = i + 1
    If i <> 4 Then
        ws.Visible = False
    End If
Next ws
ThisWorkbook.Worksheets(4).Activate

With ActiveWindow
    .DisplayHorizontalScrollBar = True
    .DisplayVerticalScrollBar = True
End With

Call ComboGetirResetRapor
Call EtiketAlanlariReset

For Each ctl In UstMenuFrame.Controls
    If TypeName(ctl) = "Label" Then
        ctl.BackColor = RGB(225, 235, 245)  'RGB(254, 254, 254)
        ctl.ForeColor = RGB(30, 30, 30)
    End If
Next ctl
Rapor.BackColor = RGB(180, 210, 240)
Rapor.ForeColor = RGB(30, 30, 30)

'_________________________

For Each ctl In UstMenuFrameAlt.Controls
    If TypeName(ctl) = "Label" Then
        ctl.BackColor = RGB(225, 235, 245)  'RGB(254, 254, 254)
        ctl.ForeColor = RGB(30, 30, 30)
        ctl.Caption = ""
    End If
Next ctl



LblAltMenuButton1.Caption = "R2.1 - Env. 1" ' Statement 2 Env.
LblAltMenuButton1.ForeColor = RGB(30, 30, 30)

LblAltMenuButton2.Caption = "R2.1 - Env. 2" 'Small Env.
LblAltMenuButton2.ForeColor = RGB(30, 30, 30)

LblAltMenuButton3.Caption = "R2.1 - Env. 3" 'Large Env.
LblAltMenuButton3.ForeColor = RGB(30, 30, 30)


ThisWorkbook.Protect "123"

Application.DisplayAlerts = True
Application.EnableEvents = True
Application.ScreenUpdating = True

End Sub

Private Sub Rapor2_2_Click()
Dim ws As Object, i As Integer
Dim ctl As MSForms.Control

Application.DisplayAlerts = False
Application.EnableEvents = False
Application.ScreenUpdating = False

ThisWorkbook.Unprotect "123"

i = 0
ThisWorkbook.Worksheets(4).Visible = True
For Each ws In ThisWorkbook.Worksheets
    i = i + 1
    If i <> 4 Then
        ws.Visible = False
    End If
Next ws
ThisWorkbook.Worksheets(4).Activate

With ActiveWindow
    .DisplayHorizontalScrollBar = True
    .DisplayVerticalScrollBar = True
End With

Call ComboGetirResetRapor2_2
Call EtiketAlanlariReset

For Each ctl In UstMenuFrame.Controls
    If TypeName(ctl) = "Label" Then
        ctl.BackColor = RGB(225, 235, 245)  'RGB(254, 254, 254)
        ctl.ForeColor = RGB(30, 30, 30)
    End If
Next ctl
Rapor2_2.BackColor = RGB(180, 210, 240)
Rapor2_2.ForeColor = RGB(30, 30, 30)

'_________________________

For Each ctl In UstMenuFrameAlt.Controls
    If TypeName(ctl) = "Label" Then
        ctl.BackColor = RGB(225, 235, 245)  'RGB(254, 254, 254)
        ctl.ForeColor = RGB(30, 30, 30)
        ctl.Caption = ""
    End If
Next ctl


LblAltMenuButton1.Caption = "R2.2 - Env. 1" '"XXXMud-G1-T2-Zf."
LblAltMenuButton1.ForeColor = RGB(30, 30, 30)

LblAltMenuButton2.Caption = "R2.2 - Env. 2" '"Bilgi. K. Zf."
LblAltMenuButton2.ForeColor = RGB(30, 30, 30)

LblAltMenuButton3.Caption = "R2.2 - Env. 3" '"Bilgi. B. Zf."
LblAltMenuButton3.ForeColor = RGB(30, 30, 30)

LblAltMenuButton4.Caption = "R2.2 - Env. 4" '"XXXMud-G2-T2-Zf."
LblAltMenuButton4.ForeColor = RGB(30, 30, 30)

LblAltMenuButton5.Caption = "R2.2 - Env. 5" '"Kurum T2. Zf."
LblAltMenuButton5.ForeColor = RGB(30, 30, 30)

LblAltMenuButton6.Caption = "R2.2 - Env. 6" '"Sonuç K. Zf."
LblAltMenuButton6.ForeColor = RGB(30, 30, 30)

LblAltMenuButton7.Caption = "R2.2 - Env. 7" '"Sonuç B. Zf."
LblAltMenuButton7.ForeColor = RGB(30, 30, 30)


ThisWorkbook.Protect "123"

Application.DisplayAlerts = True
Application.EnableEvents = True
Application.ScreenUpdating = True

End Sub



Private Sub Rapor3_1_Click()
Dim ws As Object, i As Integer
Dim ctl As MSForms.Control

Application.DisplayAlerts = False
Application.EnableEvents = False
Application.ScreenUpdating = False

ThisWorkbook.Unprotect "123"

i = 0
ThisWorkbook.Worksheets(5).Visible = True
For Each ws In ThisWorkbook.Worksheets
    i = i + 1
    If i <> 5 Then
        ws.Visible = False
    End If
Next ws
ThisWorkbook.Worksheets(5).Activate

With ActiveWindow
    .DisplayHorizontalScrollBar = True
    .DisplayVerticalScrollBar = True
End With

Call ComboGetirResetRapor3_1
Call EtiketAlanlariReset

For Each ctl In UstMenuFrame.Controls
    If TypeName(ctl) = "Label" Then
        ctl.BackColor = RGB(225, 235, 245)  'RGB(254, 254, 254)
        ctl.ForeColor = RGB(30, 30, 30)
    End If
Next ctl
Rapor3_1.BackColor = RGB(180, 210, 240)
Rapor3_1.ForeColor = RGB(30, 30, 30)

'_________________________

For Each ctl In UstMenuFrameAlt.Controls
    If TypeName(ctl) = "Label" Then
        ctl.BackColor = RGB(225, 235, 245)  'RGB(254, 254, 254)
        ctl.ForeColor = RGB(30, 30, 30)
        ctl.Caption = ""
    End If
Next ctl

LblAltMenuButton1.Caption = "R3.1 - Env. 1" '"Statement 2 Envelope"
LblAltMenuButton1.ForeColor = RGB(30, 30, 30)

LblAltMenuButton2.Caption = "R3.1 - Env. 2" '"Small Envelope"
LblAltMenuButton2.ForeColor = RGB(30, 30, 30)

LblAltMenuButton3.Caption = "R3.1 - Env. 3" '"Large Envelope"
LblAltMenuButton3.ForeColor = RGB(30, 30, 30)




ThisWorkbook.Protect "123"

Application.DisplayAlerts = True
Application.EnableEvents = True
Application.ScreenUpdating = True

End Sub

Private Sub Rapor3_2_Click()
Dim ws As Object, i As Integer
Dim ctl As MSForms.Control

Application.DisplayAlerts = False
Application.EnableEvents = False
Application.ScreenUpdating = False

ThisWorkbook.Unprotect "123"

i = 0
ThisWorkbook.Worksheets(5).Visible = True
For Each ws In ThisWorkbook.Worksheets
    i = i + 1
    If i <> 5 Then
        ws.Visible = False
    End If
Next ws
ThisWorkbook.Worksheets(5).Activate

With ActiveWindow
    .DisplayHorizontalScrollBar = True
    .DisplayVerticalScrollBar = True
End With

Call ComboGetirResetRapor3_2
Call EtiketAlanlariReset

For Each ctl In UstMenuFrame.Controls
    If TypeName(ctl) = "Label" Then
        ctl.BackColor = RGB(225, 235, 245)  'RGB(254, 254, 254)
        ctl.ForeColor = RGB(30, 30, 30)
    End If
Next ctl
Rapor3_2.BackColor = RGB(180, 210, 240)
Rapor3_2.ForeColor = RGB(30, 30, 30)

'_________________________

For Each ctl In UstMenuFrameAlt.Controls
    If TypeName(ctl) = "Label" Then
        ctl.BackColor = RGB(225, 235, 245)  'RGB(254, 254, 254)
        ctl.ForeColor = RGB(30, 30, 30)
        ctl.Caption = ""
    End If
Next ctl

LblAltMenuButton1.Caption = "R3.2 - Env. 1" '"Statement 2 Envelope"
LblAltMenuButton1.ForeColor = RGB(30, 30, 30)

LblAltMenuButton2.Caption = "R3.2 - Env. 2" '"Small Envelope"
LblAltMenuButton2.ForeColor = RGB(30, 30, 30)

LblAltMenuButton3.Caption = "R3.2 - Env. 3" '"Large Envelope"
LblAltMenuButton3.ForeColor = RGB(30, 30, 30)

LblAltMenuButton4.Caption = "R3.2 - Env. 4" '"Finansal-K. Zf."
LblAltMenuButton4.ForeColor = RGB(30, 30, 30)

LblAltMenuButton5.Caption = "R3.2 - Env. 5" '"Finansal-B. Zf."
LblAltMenuButton5.ForeColor = RGB(30, 30, 30)




ThisWorkbook.Protect "123"

Application.DisplayAlerts = True
Application.EnableEvents = True
Application.ScreenUpdating = True

End Sub

Private Sub LblAltMenuButton1_Click()
Dim ws As Object, i As Integer
Dim ctl As MSForms.Control

If LblAltMenuButton1.Caption = "" Then
    Exit Sub
End If

Application.DisplayAlerts = False
Application.EnableEvents = False
Application.ScreenUpdating = False

'ThisWorkbook.Unprotect "123"

For Each ctl In UstMenuFrameAlt.Controls
    If TypeName(ctl) = "Label" Then
        ctl.BackColor = RGB(225, 235, 245)  'RGB(254, 254, 254)
        ctl.ForeColor = RGB(30, 30, 30)
    End If
Next ctl
If LblAltMenuButton1.Caption <> "" Then
    LblAltMenuButton1.BackColor = RGB(180, 210, 240)
    LblAltMenuButton1.ForeColor = RGB(30, 30, 30)
End If

Call EtiketAlanlariReset

If Rapor1.BackColor = RGB(180, 210, 240) Then
    'MsgBox "Rapor1"
    'Tutanak2 zarfı
    Call Rapor1Tutanak2Zarfi
ElseIf Rapor.BackColor = RGB(180, 210, 240) Then
    'MsgBox "Rapor"
    'Tutanak2 zarfı
    Call RaporTutanak2Zarfi
ElseIf Rapor2_2.BackColor = RGB(180, 210, 240) Then
    'MsgBox "Rapor2_2"
    'Tutanak2 zarfı
    Call XXXMudGidenTutanak2Zarfi
ElseIf Rapor3_1.BackColor = RGB(180, 210, 240) Then
    'MsgBox "Report 3.1"
    'Tutanak2 zarfı
    Call Rapor3_1Tutanak2Zarfi
ElseIf Rapor3_2.BackColor = RGB(180, 210, 240) Then
    'MsgBox "Report 3.2"
    'Tutanak2 zarfı
    Call Rapor3_2Tutanak2Zarfi
End If


'ThisWorkbook.Protect "123"

Application.DisplayAlerts = True
Application.EnableEvents = True
Application.ScreenUpdating = True

End Sub

Sub Rapor1Tutanak2Zarfi()
Dim ctl As MSForms.Control
Dim i As Long, j As Long

Call EtiketAlanlariReset
  
'__________

If ComboGetir.Value <> "" Then
    Set IlkSiraBulGlobal = Range("CE7:CE100000").Find(What:=ComboGetir.Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    Set SonSiraBulGlobal = Range("CF7:CF100000").Find(What:=ComboGetir.Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    If Not IlkSiraBulGlobal Is Nothing Then
        IlkSiraGlobal = IlkSiraBulGlobal.Row
    Else
        For Each ctl In UstMenuFrameAlt.Controls
            If TypeName(ctl) = "Label" Then
                ctl.BackColor = RGB(225, 235, 245)  'RGB(254, 254, 254)
                ctl.ForeColor = RGB(30, 30, 30)
            End If
        Next ctl
        MsgBox "The operation cannot be completed since the entered serial number was not found.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    If Not SonSiraBulGlobal Is Nothing Then
        SonSiraGlobal = SonSiraBulGlobal.Row
    Else
        For Each ctl In UstMenuFrameAlt.Controls
            If TypeName(ctl) = "Label" Then
                ctl.BackColor = RGB(225, 235, 245)  'RGB(254, 254, 254)
                ctl.ForeColor = RGB(30, 30, 30)
            End If
        Next ctl
        MsgBox "The operation cannot be completed since the entered serial number was not found.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
Else
    For Each ctl In UstMenuFrameAlt.Controls
        If TypeName(ctl) = "Label" Then
            ctl.BackColor = RGB(225, 235, 245)  'RGB(254, 254, 254)
            ctl.ForeColor = RGB(30, 30, 30)
        End If
    Next ctl
    MsgBox "Please select the serial number for the module you wish to operate on from the drop-down list located in the upper-left corner.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

'__________


'Etiket bilgileri
TextKURUM_A.Value = "ORGANIZATION A"
TextSube.Value = ThisWorkbook.Worksheets(2).Cells(6, 99).Value & " Unit"

If SonSiraGlobal - IlkSiraGlobal + 1 < 7 Then
    j = 0
    For i = IlkSiraGlobal To SonSiraGlobal
        j = j + 1
        '1 Adet 10 Öğe Türü-X1 gibi yazacak
        Controls("TextMuhatapTema" & j).Value = Cells(i, 44).Value & " unit(s) of " & Cells(i, 41).Value & " " & Cells(i, 38).Value
        If i = SonSiraGlobal Then 'Tema no en lat satıra
            j = j + 1
            Controls("TextMuhatapTema" & j).Value = "Theme 1 No: " & Cells(IlkSiraGlobal, 23).Value
        End If
    Next i
End If


Son:

End Sub

Sub RaporTutanak2Zarfi()
Dim ctl As MSForms.Control
Dim i As Long, j As Long

Call EtiketAlanlariReset
  
'__________

If ComboGetir.Value <> "" Then
    Set IlkSiraBulGlobal = Range("CM7:CM100000").Find(What:=ComboGetir.Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    Set SonSiraBulGlobal = Range("CN7:CN100000").Find(What:=ComboGetir.Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    If Not IlkSiraBulGlobal Is Nothing Then
        IlkSiraGlobal = IlkSiraBulGlobal.Row
    Else
        For Each ctl In UstMenuFrameAlt.Controls
            If TypeName(ctl) = "Label" Then
                ctl.BackColor = RGB(225, 235, 245)  'RGB(254, 254, 254)
                ctl.ForeColor = RGB(30, 30, 30)
            End If
        Next ctl
        MsgBox "The operation cannot be completed since the entered serial number was not found.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    If Not SonSiraBulGlobal Is Nothing Then
        SonSiraGlobal = SonSiraBulGlobal.Row
    Else
        For Each ctl In UstMenuFrameAlt.Controls
            If TypeName(ctl) = "Label" Then
                ctl.BackColor = RGB(225, 235, 245)  'RGB(254, 254, 254)
                ctl.ForeColor = RGB(30, 30, 30)
            End If
        Next ctl
        MsgBox "The operation cannot be completed since the entered serial number was not found.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
Else
    For Each ctl In UstMenuFrameAlt.Controls
        If TypeName(ctl) = "Label" Then
            ctl.BackColor = RGB(225, 235, 245)  'RGB(254, 254, 254)
            ctl.ForeColor = RGB(30, 30, 30)
        End If
    Next ctl
    MsgBox "Please select the serial number for the module you wish to operate on from the drop-down list located in the upper-left corner.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

'__________


'Etiket bilgileri
TextKURUM_A.Value = "ORGANIZATION A"
TextSube.Value = ThisWorkbook.Worksheets(2).Cells(6, 99).Value & " Unit"

If SonSiraGlobal - IlkSiraGlobal + 1 < 7 Then
    j = 0
    For i = IlkSiraGlobal To SonSiraGlobal
        j = j + 1
        '1 Adet 10 Öğe Türü-X1 gibi yazacak
        Controls("TextMuhatapTema" & j).Value = Cells(i, 52).Value & " unit(s) of " & Cells(i, 49).Value & " " & Cells(i, 46).Value
        If i = SonSiraGlobal Then 'Tema no en lat satıra
            j = j + 1
            Controls("TextMuhatapTema" & j).Value = "Theme 1 No: " & Cells(IlkSiraGlobal, 31).Value
        End If
    Next i
End If


Son:

End Sub

Sub XXXMudGidenTutanak2Zarfi()
Dim ctl As MSForms.Control
Dim i As Long, j As Long

Call EtiketAlanlariReset
  
'__________

If ComboGetir.Value <> "" Then
    Set IlkSiraBulGlobal = Range("CM7:CM100000").Find(What:=ComboGetir.Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    Set SonSiraBulGlobal = Range("CN7:CN100000").Find(What:=ComboGetir.Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    If Not IlkSiraBulGlobal Is Nothing Then
        IlkSiraGlobal = IlkSiraBulGlobal.Row
    Else
        For Each ctl In UstMenuFrameAlt.Controls
            If TypeName(ctl) = "Label" Then
                ctl.BackColor = RGB(225, 235, 245)  'RGB(254, 254, 254)
                ctl.ForeColor = RGB(30, 30, 30)
            End If
        Next ctl
        MsgBox "The operation cannot be completed since the entered serial number was not found.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    If Not SonSiraBulGlobal Is Nothing Then
        SonSiraGlobal = SonSiraBulGlobal.Row
    Else
        For Each ctl In UstMenuFrameAlt.Controls
            If TypeName(ctl) = "Label" Then
                ctl.BackColor = RGB(225, 235, 245)  'RGB(254, 254, 254)
                ctl.ForeColor = RGB(30, 30, 30)
            End If
        Next ctl
        MsgBox "The operation cannot be completed since the entered serial number was not found.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
Else
    For Each ctl In UstMenuFrameAlt.Controls
        If TypeName(ctl) = "Label" Then
            ctl.BackColor = RGB(225, 235, 245)  'RGB(254, 254, 254)
            ctl.ForeColor = RGB(30, 30, 30)
        End If
    Next ctl
    MsgBox "Please select the serial number for the module you wish to operate on from the drop-down list located in the upper-left corner.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

'__________


'Etiket bilgileri
TextKURUM_A.Value = "ORGANIZATION A"
TextSube.Value = ThisWorkbook.Worksheets(2).Cells(6, 99).Value & " Unit"

If Cells(IlkSiraGlobal, 187).Value = "All" Then
    If SonSiraGlobal - IlkSiraGlobal + 1 < 7 Then
        j = 0
        For i = IlkSiraGlobal To SonSiraGlobal
            j = j + 1
            '1 Adet 10 Öğe Türü-X1 gibi yazacak
            Controls("TextMuhatapTema" & j).Value = Cells(i, 52).Value & " unit(s) of " & Cells(i, 49).Value & " " & Cells(i, 46).Value
            If i = SonSiraGlobal Then 'Tema no en lat satıra
                j = j + 1
                Controls("TextMuhatapTema" & j).Value = "Theme 1 No: " & Cells(IlkSiraGlobal, 31).Value
            End If
        Next i
    End If
ElseIf Cells(IlkSiraGlobal, 187).Value = "Technique A" Then 'Sadece Technique A
    If SonSiraGlobal - IlkSiraGlobal + 1 < 7 Then
        j = 0
        For i = IlkSiraGlobal To SonSiraGlobal
            If Left(Cells(i, 63).Value, 11) = "Technique A" Then
                j = j + 1
                '1 Adet 10 Öğe Türü-X1 gibi yazacak
                Controls("TextMuhatapTema" & j).Value = Cells(i, 52).Value & " unit(s) of " & Cells(i, 49).Value & " " & Cells(i, 46).Value
            End If
            If i = SonSiraGlobal Then 'Tema no en lat satıra
                j = j + 1
                Controls("TextMuhatapTema" & j).Value = "Theme 1 No: " & Cells(IlkSiraGlobal, 31).Value
            End If
        Next i
    End If
End If

Son:

End Sub

Sub XXXMudGelenTutanak2Zarfi()
Dim ctl As MSForms.Control
Dim i As Long, j As Long

Call EtiketAlanlariReset
  
'__________

If ComboGetir.Value <> "" Then
    Set IlkSiraBulGlobal = Range("CM7:CM100000").Find(What:=ComboGetir.Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    Set SonSiraBulGlobal = Range("CN7:CN100000").Find(What:=ComboGetir.Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    If Not IlkSiraBulGlobal Is Nothing Then
        IlkSiraGlobal = IlkSiraBulGlobal.Row
    Else
        For Each ctl In UstMenuFrameAlt.Controls
            If TypeName(ctl) = "Label" Then
                ctl.BackColor = RGB(225, 235, 245)  'RGB(254, 254, 254)
                ctl.ForeColor = RGB(30, 30, 30)
            End If
        Next ctl
        MsgBox "The operation cannot be completed since the entered serial number was not found.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    If Not SonSiraBulGlobal Is Nothing Then
        SonSiraGlobal = SonSiraBulGlobal.Row
    Else
        For Each ctl In UstMenuFrameAlt.Controls
            If TypeName(ctl) = "Label" Then
                ctl.BackColor = RGB(225, 235, 245)  'RGB(254, 254, 254)
                ctl.ForeColor = RGB(30, 30, 30)
            End If
        Next ctl
        MsgBox "The operation cannot be completed since the entered serial number was not found.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
Else
    For Each ctl In UstMenuFrameAlt.Controls
        If TypeName(ctl) = "Label" Then
            ctl.BackColor = RGB(225, 235, 245)  'RGB(254, 254, 254)
            ctl.ForeColor = RGB(30, 30, 30)
        End If
    Next ctl
    MsgBox "Please select the serial number for the module you wish to operate on from the drop-down list located in the upper-left corner.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

'__________

If Cells(IlkSiraGlobal, 187).Value = "All" Then 'TÜM TipALAR

    If Cells(IlkSiraGlobal, 188).Value = "No" Then 'XXXMud'den gelen paket açılmayacak
        For Each ctl In UstMenuFrameAlt.Controls
            If TypeName(ctl) = "Label" Then
                ctl.BackColor = RGB(225, 235, 245)  'RGB(254, 254, 254)
                ctl.ForeColor = RGB(30, 30, 30)
            End If
        Next ctl
        MsgBox "Since it is indicated in the Report 2.2 entry section that all Type A items will be sent to XXXMud and the package received from XXXMud will not be opened, it is not possible to print a label/envelope for the Statement 2 envelope.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    ElseIf Cells(IlkSiraGlobal, 188).Value = "Yes" Then 'XXXMud'den gelen paket işlenecek, dolayısıyla tüm tipAlar tek zarfa girecek
        For Each ctl In UstMenuFrameAlt.Controls
            If TypeName(ctl) = "Label" Then
                ctl.BackColor = RGB(225, 235, 245)  'RGB(254, 254, 254)
                ctl.ForeColor = RGB(30, 30, 30)
            End If
        Next ctl
        MsgBox "Since it is indicated in the Report 2.2 entry section that all Type A items will be sent to XXXMud and the package received from XXXMud will be opened, only one Statement 2 operation can be performed. To print the label/envelope for the Statement 2 envelope, please click the 'Institution Statement 2 Envelope' button.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    Else 'XXXMud'den gelen paket kararı yok; henüz açılmadığı varsayıldı.
        For Each ctl In UstMenuFrameAlt.Controls
            If TypeName(ctl) = "Label" Then
                ctl.BackColor = RGB(225, 235, 245)  'RGB(254, 254, 254)
                ctl.ForeColor = RGB(30, 30, 30)
            End If
        Next ctl
        MsgBox "Since Report 2.2 entry has not been specified via the Report/Report 2.2 interface, the label/envelope for the Statement 2 envelope cannot be printed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If

ElseIf Cells(IlkSiraGlobal, 187).Value = "Technique A" Then 'SADECE Technique A TipALAR

    If Cells(IlkSiraGlobal, 188).Value = "No" Then 'XXXMud'den gelen paket açılmayacak
        For Each ctl In UstMenuFrameAlt.Controls
            If TypeName(ctl) = "Label" Then
                ctl.BackColor = RGB(225, 235, 245)  'RGB(254, 254, 254)
                ctl.ForeColor = RGB(30, 30, 30)
            End If
        Next ctl
        MsgBox "Since it is specified in the Report 2.2 entry section that the package received from XXXMud will not be opened, the label/envelope for the Statement 2 envelope received from XXXMud cannot be printed. Please click the 'Institution Statement 2 Envelope' button to print the envelope/label for the Statement 2 of type A items that are not sent to XXXMud (i.e., those retained within the unit).", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    ElseIf Cells(IlkSiraGlobal, 188).Value = "Yes" Then 'XXXMud'den gelen paket işlenecek, dolayısıyla tüm tipAlar tek zarfa girecek
        
        If Cells(IlkSiraGlobal, 189).Value = "Yes" Then 'tutanak2 birleşecek(Tüm tipAlar bu zarfta)
            For Each ctl In UstMenuFrameAlt.Controls
                If TypeName(ctl) = "Label" Then
                    ctl.BackColor = RGB(225, 235, 245)  'RGB(254, 254, 254)
                    ctl.ForeColor = RGB(30, 30, 30)
                End If
            Next ctl
            MsgBox "Since it is specified in the Report 2.2 entry section that only Technique A type A items will be sent to XXXMud, the package received from XXXMud will be opened, and the Statement 2 processes will be merged, only one Statement 2 operation can be performed. Please click the 'Institution Statement 2 Envelope' button to print the envelope/label for the Statement 2.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
            GoTo Son
        ElseIf Cells(IlkSiraGlobal, 189).Value = "No" Then 'tutanak2 birleşmeyecek, sadece Technique Ai tekrar kapat
            
            'Etiket bilgileri
            TextKURUM_A.Value = "ORGANIZATION A"
            TextSube.Value = ThisWorkbook.Worksheets(2).Cells(6, 99).Value & " Unit"

            If SonSiraGlobal - IlkSiraGlobal + 1 < 7 Then
                j = 0
                For i = IlkSiraGlobal To SonSiraGlobal
                    If Left(Cells(i, 63).Value, 11) = "Technique A" Then
                        j = j + 1
                        '1 Adet 10 Öğe Türü-X1 gibi yazacak
                        Controls("TextMuhatapTema" & j).Value = Cells(i, 52).Value & " unit(s) of " & Cells(i, 49).Value & " " & Cells(i, 46).Value
                    End If
                    If i = SonSiraGlobal Then 'Tema no en lat satıra
                        j = j + 1
                        Controls("TextMuhatapTema" & j).Value = "Theme 1 No: " & Cells(IlkSiraGlobal, 31).Value
                    End If
                Next i
            End If
            
        Else
            For Each ctl In UstMenuFrameAlt.Controls
                If TypeName(ctl) = "Label" Then
                    ctl.BackColor = RGB(225, 235, 245)  'RGB(254, 254, 254)
                    ctl.ForeColor = RGB(30, 30, 30)
                End If
            Next ctl
            MsgBox "Since the Report 2.2 entry has not been specified in the Report2.1 / Report 2.2 interface, the envelope/label for the Statement 2 envelope cannot be printed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
            GoTo Son
        End If

    Else 'XXXMud'den gelen paket kararı yok; henüz açılmadığı varsayıldı.
        For Each ctl In UstMenuFrameAlt.Controls
            If TypeName(ctl) = "Label" Then
                ctl.BackColor = RGB(225, 235, 245)  'RGB(254, 254, 254)
                ctl.ForeColor = RGB(30, 30, 30)
            End If
        Next ctl
        MsgBox "Since the Report 2.2 entry has not been specified in the Report 2.1 / Report 2.2 interface, the label/envelope for the Statement 2 envelope cannot be printed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If

Else
    For Each ctl In UstMenuFrameAlt.Controls
        If TypeName(ctl) = "Label" Then
            ctl.BackColor = RGB(225, 235, 245)  'RGB(254, 254, 254)
            ctl.ForeColor = RGB(30, 30, 30)
        End If
    Next ctl
    MsgBox "Since the Report 2.2 entry has not been specified in the Report 2.1 / Report 2.2 interface, the label/envelope for the Statement 2 envelope cannot be printed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If
         

Son:

End Sub

Sub KurumTutanak2Zarfi()
Dim ctl As MSForms.Control
Dim i As Long, j As Long

Call EtiketAlanlariReset
  
'__________

If ComboGetir.Value <> "" Then
    Set IlkSiraBulGlobal = Range("CM7:CM100000").Find(What:=ComboGetir.Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    Set SonSiraBulGlobal = Range("CN7:CN100000").Find(What:=ComboGetir.Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    If Not IlkSiraBulGlobal Is Nothing Then
        IlkSiraGlobal = IlkSiraBulGlobal.Row
    Else
        For Each ctl In UstMenuFrameAlt.Controls
            If TypeName(ctl) = "Label" Then
                ctl.BackColor = RGB(225, 235, 245)  'RGB(254, 254, 254)
                ctl.ForeColor = RGB(30, 30, 30)
            End If
        Next ctl
        MsgBox "The operation cannot be completed since the entered serial number was not found.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    If Not SonSiraBulGlobal Is Nothing Then
        SonSiraGlobal = SonSiraBulGlobal.Row
    Else
        For Each ctl In UstMenuFrameAlt.Controls
            If TypeName(ctl) = "Label" Then
                ctl.BackColor = RGB(225, 235, 245)  'RGB(254, 254, 254)
                ctl.ForeColor = RGB(30, 30, 30)
            End If
        Next ctl
        MsgBox "The operation cannot be completed since the entered serial number was not found.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
Else
    For Each ctl In UstMenuFrameAlt.Controls
        If TypeName(ctl) = "Label" Then
            ctl.BackColor = RGB(225, 235, 245)  'RGB(254, 254, 254)
            ctl.ForeColor = RGB(30, 30, 30)
        End If
    Next ctl
    MsgBox "Please select the serial number for the module you wish to operate on from the drop-down list located in the upper-left corner.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

'__________

If Cells(IlkSiraGlobal, 187).Value = "All" Then 'TÜM TipALAR

    If Cells(IlkSiraGlobal, 188).Value = "No" Then 'XXXMud'den gelen paket açılmayacak
        For Each ctl In UstMenuFrameAlt.Controls
            If TypeName(ctl) = "Label" Then
                ctl.BackColor = RGB(225, 235, 245)  'RGB(254, 254, 254)
                ctl.ForeColor = RGB(30, 30, 30)
            End If
        Next ctl
        MsgBox "Since the Report 2.2 entry section indicates that all Technique A items will be sent to XXXMud and the package received from XXXMud will not be opened, the label/envelope for the Statement 2 envelope cannot be printed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    ElseIf Cells(IlkSiraGlobal, 188).Value = "Yes" Then 'XXXMud'den gelen paket işlenecek, dolayısıyla tüm tipAlar tek zarfa girecek

        'Etiket bilgileri
        TextKURUM_A.Value = "ORGANIZATION A"
        TextSube.Value = ThisWorkbook.Worksheets(2).Cells(6, 99).Value & " Unit"

        If SonSiraGlobal - IlkSiraGlobal + 1 < 7 Then
            j = 0
            For i = IlkSiraGlobal To SonSiraGlobal
                j = j + 1
                '1 Adet 10 Öğe Türü-X1 gibi yazacak
                Controls("TextMuhatapTema" & j).Value = Cells(i, 52).Value & " unit(s) of " & Cells(i, 49).Value & " " & Cells(i, 46).Value
                If i = SonSiraGlobal Then 'Tema no en lat satıra
                    j = j + 1
                    Controls("TextMuhatapTema" & j).Value = "Theme 1 No: " & Cells(IlkSiraGlobal, 31).Value
                End If
            Next i
        End If
        
    Else 'XXXMud'den gelen paket kararı yok; henüz açılmadığı varsayıldı.
        For Each ctl In UstMenuFrameAlt.Controls
            If TypeName(ctl) = "Label" Then
                ctl.BackColor = RGB(225, 235, 245)  'RGB(254, 254, 254)
                ctl.ForeColor = RGB(30, 30, 30)
            End If
        Next ctl
        MsgBox "Since the Report / Report 2.2 interface does not include a Report 2.2 entry, the label/envelope for the Statement 2 envelope cannot be printed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If

ElseIf Cells(IlkSiraGlobal, 187).Value = "Technique A" Then 'SADECE Technique A TipALAR

    If Cells(IlkSiraGlobal, 188).Value = "No" Then 'XXXMud'den gelen paket açılmayacak, sadece Technique A olmayanlar kapatılacak

        'Etiket bilgileri
        TextKURUM_A.Value = "ORGANIZATION A"
        TextSube.Value = ThisWorkbook.Worksheets(2).Cells(6, 99).Value & " Unit"
        
        If SonSiraGlobal - IlkSiraGlobal + 1 < 7 Then
            j = 0
            For i = IlkSiraGlobal To SonSiraGlobal
                If Left(Cells(i, 63).Value, 8) <> "Technique A" Then
                    j = j + 1
                    '1 Adet 10 Öğe Türü-X1 gibi yazacak
                    Controls("TextMuhatapTema" & j).Value = Cells(i, 52).Value & " unit(s) of " & Cells(i, 49).Value & " " & Cells(i, 46).Value
                End If
                If i = SonSiraGlobal Then 'Tema no en lat satıra
                    j = j + 1
                    Controls("TextMuhatapTema" & j).Value = "Theme 1 No: " & Cells(IlkSiraGlobal, 31).Value
                End If
            Next i
        End If
        
    ElseIf Cells(IlkSiraGlobal, 188).Value = "Yes" Then 'XXXMud'den gelen paket işlenecek, dolayısıyla tüm tipAlar tek zarfa girecek
        
        If Cells(IlkSiraGlobal, 189).Value = "Yes" Then 'tutanak2 birleşecek(Tüm tipAlar bu zarfta)

            'Etiket bilgileri
            TextKURUM_A.Value = "ORGANIZATION A"
            TextSube.Value = ThisWorkbook.Worksheets(2).Cells(6, 99).Value & " Unit"
        
            If SonSiraGlobal - IlkSiraGlobal + 1 < 7 Then
                j = 0
                For i = IlkSiraGlobal To SonSiraGlobal
                    j = j + 1
                    '1 Adet 10 Öğe Türü-X1 gibi yazacak
                    Controls("TextMuhatapTema" & j).Value = Cells(i, 52).Value & " unit(s) of " & Cells(i, 49).Value & " " & Cells(i, 46).Value
                    If i = SonSiraGlobal Then 'Tema no en lat satıra
                        j = j + 1
                        Controls("TextMuhatapTema" & j).Value = "Theme 1 No: " & Cells(IlkSiraGlobal, 31).Value
                    End If
                Next i
            End If
        
        ElseIf Cells(IlkSiraGlobal, 189).Value = "No" Then 'tutanak2 birleşmeyecek, sadece Technique A olmayanları kapat
   
            'Etiket bilgileri
            TextKURUM_A.Value = "ORGANIZATION A"
            TextSube.Value = ThisWorkbook.Worksheets(2).Cells(6, 99).Value & " Unit"
        
            If SonSiraGlobal - IlkSiraGlobal + 1 < 7 Then
                j = 0
                For i = IlkSiraGlobal To SonSiraGlobal
                    If Left(Cells(i, 63).Value, 8) <> "Technique A" Then
                        j = j + 1
                        '1 Adet 10 Öğe Türü-X1 gibi yazacak
                        Controls("TextMuhatapTema" & j).Value = Cells(i, 52).Value & " unit(s) of " & Cells(i, 49).Value & " " & Cells(i, 46).Value
                    End If
                    If i = SonSiraGlobal Then 'Tema no en lat satıra
                        j = j + 1
                        Controls("TextMuhatapTema" & j).Value = "Theme 1 No: " & Cells(IlkSiraGlobal, 31).Value
                    End If
                Next i
            End If
            
        Else
            For Each ctl In UstMenuFrameAlt.Controls
                If TypeName(ctl) = "Label" Then
                    ctl.BackColor = RGB(225, 235, 245)  'RGB(254, 254, 254)
                    ctl.ForeColor = RGB(30, 30, 30)
                End If
            Next ctl
            MsgBox "Since no Report 2.2 entry has been specified via the Report 2.1 / Report 2.2 interface, the label/envelope for the Statement 2 envelope cannot be printed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
            GoTo Son
        End If

    Else 'XXXMud'den gelen paket kararı yok; henüz açılmadığı varsayıldı.
        For Each ctl In UstMenuFrameAlt.Controls
            If TypeName(ctl) = "Label" Then
                ctl.BackColor = RGB(225, 235, 245)  'RGB(254, 254, 254)
                ctl.ForeColor = RGB(30, 30, 30)
            End If
        Next ctl
        MsgBox "Since no Report 2.2 entry has been specified via the Report 2.1 / Report 2.2 interface, the label/envelope for the Statement 2 envelope cannot be printed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If

Else
    For Each ctl In UstMenuFrameAlt.Controls
        If TypeName(ctl) = "Label" Then
            ctl.BackColor = RGB(225, 235, 245)  'RGB(254, 254, 254)
            ctl.ForeColor = RGB(30, 30, 30)
        End If
    Next ctl
    MsgBox "Since no Report 2.2 entry has been specified via the Report 2.1 / Report 2.2 interface, the label/envelope for the Statement 2 envelope cannot be printed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

        

Son:

End Sub



Sub Rapor3_1Tutanak2Zarfi()
Dim ctl As MSForms.Control
Dim i As Long, j As Long

Call EtiketAlanlariReset
  
'__________

If ComboGetir.Value <> "" Then
    Set IlkSiraBulGlobal = Range("FG7:FG100000").Find(What:=ComboGetir.Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    Set SonSiraBulGlobal = Range("FH7:FH100000").Find(What:=ComboGetir.Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    If Not IlkSiraBulGlobal Is Nothing Then
        IlkSiraGlobal = IlkSiraBulGlobal.Row
    Else
        For Each ctl In UstMenuFrameAlt.Controls
            If TypeName(ctl) = "Label" Then
                ctl.BackColor = RGB(225, 235, 245)  'RGB(254, 254, 254)
                ctl.ForeColor = RGB(30, 30, 30)
            End If
        Next ctl
        MsgBox "The operation cannot be completed since the entered serial number was not found.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    If Not SonSiraBulGlobal Is Nothing Then
        SonSiraGlobal = SonSiraBulGlobal.Row
    Else
        For Each ctl In UstMenuFrameAlt.Controls
            If TypeName(ctl) = "Label" Then
                ctl.BackColor = RGB(225, 235, 245)  'RGB(254, 254, 254)
                ctl.ForeColor = RGB(30, 30, 30)
            End If
        Next ctl
        MsgBox "The operation cannot be completed since the entered serial number was not found.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
Else
    For Each ctl In UstMenuFrameAlt.Controls
        If TypeName(ctl) = "Label" Then
            ctl.BackColor = RGB(225, 235, 245)  'RGB(254, 254, 254)
            ctl.ForeColor = RGB(30, 30, 30)
        End If
    Next ctl
    MsgBox "Please select the serial number for the module you wish to operate on from the drop-down list located in the upper-left corner.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

'__________


'Etiket bilgileri
TextKURUM_A.Value = "ORGANIZATION A"
TextSube.Value = ThisWorkbook.Worksheets(2).Cells(6, 99).Value & " Unit"

If SonSiraGlobal - IlkSiraGlobal + 1 < 7 Then
    j = 0
    For i = IlkSiraGlobal To SonSiraGlobal
        j = j + 1
        '1 Adet 10 Öğe Türü-X1 gibi yazacak
        Controls("TextMuhatapTema" & j).Value = Cells(i, 136).Value & " unit(s) of " & Cells(i, 133).Value & " " & Cells(i, 130).Value
        If i = SonSiraGlobal Then 'Tema no en lat satıra
            j = j + 1
            If Cells(IlkSiraGlobal, 100).Value = "Type A" Then
                Controls("TextMuhatapTema" & j).Value = "Theme 1 No: " & Cells(IlkSiraGlobal, 98).Value
            Else
                Controls("TextMuhatapTema" & j).Value = "Theme 2 No: " & Cells(IlkSiraGlobal, 98).Value
            End If
        End If
    Next i
End If


Son:

End Sub

Sub Rapor3_2Tutanak2Zarfi()
Dim ctl As MSForms.Control
Dim i As Long, j As Long

Call EtiketAlanlariReset
  
'__________

If ComboGetir.Value <> "" Then
    Set IlkSiraBulGlobal = Range("FG7:FG100000").Find(What:=ComboGetir.Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    Set SonSiraBulGlobal = Range("FH7:FH100000").Find(What:=ComboGetir.Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    If Not IlkSiraBulGlobal Is Nothing Then
        IlkSiraGlobal = IlkSiraBulGlobal.Row
    Else
        For Each ctl In UstMenuFrameAlt.Controls
            If TypeName(ctl) = "Label" Then
                ctl.BackColor = RGB(225, 235, 245)  'RGB(254, 254, 254)
                ctl.ForeColor = RGB(30, 30, 30)
            End If
        Next ctl
        MsgBox "The operation cannot be completed since the entered serial number was not found.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    If Not SonSiraBulGlobal Is Nothing Then
        SonSiraGlobal = SonSiraBulGlobal.Row
    Else
        For Each ctl In UstMenuFrameAlt.Controls
            If TypeName(ctl) = "Label" Then
                ctl.BackColor = RGB(225, 235, 245)  'RGB(254, 254, 254)
                ctl.ForeColor = RGB(30, 30, 30)
            End If
        Next ctl
        MsgBox "The operation cannot be completed since the entered serial number was not found.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
Else
    For Each ctl In UstMenuFrameAlt.Controls
        If TypeName(ctl) = "Label" Then
            ctl.BackColor = RGB(225, 235, 245)  'RGB(254, 254, 254)
            ctl.ForeColor = RGB(30, 30, 30)
        End If
    Next ctl
    MsgBox "Please select the serial number for the module you wish to operate on from the drop-down list located in the upper-left corner.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

'__________


'Etiket bilgileri
TextKURUM_A.Value = "ORGANIZATION A"
TextSube.Value = ThisWorkbook.Worksheets(2).Cells(6, 99).Value & " Unit"

If SonSiraGlobal - IlkSiraGlobal + 1 < 7 Then
    j = 0
    For i = IlkSiraGlobal To SonSiraGlobal
        j = j + 1
        '1 Adet 10 Öğe Türü-X1 gibi yazacak
        Controls("TextMuhatapTema" & j).Value = Cells(i, 58).Value & " unit(s) of " & Cells(i, 55).Value & " " & Cells(i, 52).Value
        If i = SonSiraGlobal Then 'Tema no en lat satıra
            j = j + 1
            If Cells(IlkSiraGlobal, 28).Value = "Type A" Then
                Controls("TextMuhatapTema" & j).Value = "Theme 1 No: " & Cells(IlkSiraGlobal, 26).Value
            Else
                Controls("TextMuhatapTema" & j).Value = "Theme 2 No: " & Cells(IlkSiraGlobal, 26).Value
            End If
        End If
    Next i
End If


Son:

End Sub


Private Sub LblAltMenuButton2_Click()
Dim ws As Object, i As Integer
Dim ctl As MSForms.Control

If LblAltMenuButton2.Caption = "" Then
    Exit Sub
End If

Application.DisplayAlerts = False
Application.EnableEvents = False
Application.ScreenUpdating = False

'ThisWorkbook.Unprotect "123"

BoolKucukZarf = True

For Each ctl In UstMenuFrameAlt.Controls
    If TypeName(ctl) = "Label" Then
        ctl.BackColor = RGB(225, 235, 245)  'RGB(254, 254, 254)
        ctl.ForeColor = RGB(30, 30, 30)
    End If
Next ctl
If LblAltMenuButton2.Caption <> "" Then
    LblAltMenuButton2.BackColor = RGB(180, 210, 240)
    LblAltMenuButton2.ForeColor = RGB(30, 30, 30)
End If

Call EtiketAlanlariReset

If Rapor1.BackColor = RGB(180, 210, 240) Then
    'MsgBox "Rapor1"
    'Küçük Zarf
    Call Rapor1KucukZarf
ElseIf Rapor.BackColor = RGB(180, 210, 240) Then
    'MsgBox "Rapor"
    'Küçük Zarf
    Call RaporKucukZarf
ElseIf Rapor2_2.BackColor = RGB(180, 210, 240) Then
    'MsgBox "Rapor2_2"
    'Küçük Zarf
    Call BilgilendirmeKucukZarf
ElseIf Rapor3_1.BackColor = RGB(180, 210, 240) Then
    'MsgBox "Report 3.1"
    'Küçük Zarf
    Call Rapor3_1KucukZarf
ElseIf Rapor3_2.BackColor = RGB(180, 210, 240) Then
    'MsgBox "Report 3.2"
    'Küçük Zarf
    Call Rapor3_2KucukZarf
End If


BoolKucukZarf = False

'ThisWorkbook.Protect "123"

Application.DisplayAlerts = True
Application.EnableEvents = True
Application.ScreenUpdating = True

End Sub

Sub Rapor1KucukZarf()
Dim ctl As MSForms.Control
Dim i As Long, j As Long

Dim AutoPath As String, DestTaslak As String, OpenControl As String
Dim TaslakFile As String, objWord As Object, objDoc As Object, Kurum_ANoStr As String
Dim StrDosyaNo As String


Call EtiketAlanlariReset
  
'__________

If ComboGetir.Value <> "" Then
    Set IlkSiraBulGlobal = Range("CE7:CE100000").Find(What:=ComboGetir.Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    Set SonSiraBulGlobal = Range("CF7:CF100000").Find(What:=ComboGetir.Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    If Not IlkSiraBulGlobal Is Nothing Then
        IlkSiraGlobal = IlkSiraBulGlobal.Row
    Else
        For Each ctl In UstMenuFrameAlt.Controls
            If TypeName(ctl) = "Label" Then
                ctl.BackColor = RGB(225, 235, 245)  'RGB(254, 254, 254)
                ctl.ForeColor = RGB(30, 30, 30)
            End If
        Next ctl
        MsgBox "The operation cannot be completed since the entered serial number was not found.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    If Not SonSiraBulGlobal Is Nothing Then
        SonSiraGlobal = SonSiraBulGlobal.Row
    Else
        For Each ctl In UstMenuFrameAlt.Controls
            If TypeName(ctl) = "Label" Then
                ctl.BackColor = RGB(225, 235, 245)  'RGB(254, 254, 254)
                ctl.ForeColor = RGB(30, 30, 30)
            End If
        Next ctl
        MsgBox "The operation cannot be completed since the entered serial number was not found.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
Else
    For Each ctl In UstMenuFrameAlt.Controls
        If TypeName(ctl) = "Label" Then
            ctl.BackColor = RGB(225, 235, 245)  'RGB(254, 254, 254)
            ctl.ForeColor = RGB(30, 30, 30)
        End If
    Next ctl
    MsgBox "Please select the serial number for the module you wish to operate on from the drop-down list located in the upper-left corner.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

'__________


AutoPath = ThisWorkbook.Path
DestTaslak = AutoPath & "\System Files\System Templates\Report 1 Cover Letter Template\"
TaslakFile = "Report 1 Cover Letter.docm"

'System Files klasör adını kontrol et.
If Not Dir(AutoPath & "\System Files\", vbDirectory) <> vbNullString Then
    For Each ctl In UstMenuFrameAlt.Controls
        If TypeName(ctl) = "Label" Then
            ctl.BackColor = RGB(225, 235, 245)  'RGB(254, 254, 254)
            ctl.ForeColor = RGB(30, 30, 30)
        End If
    Next ctl
    MsgBox AutoPath & "\System Files\" & " directory cannot be accessed. The folder named 'System Files' might have been renamed or deleted."
    GoTo Son
End If
If Not Dir(DestTaslak & TaslakFile, vbDirectory) <> vbNullString Then
    For Each ctl In UstMenuFrameAlt.Controls
        If TypeName(ctl) = "Label" Then
            ctl.BackColor = RGB(225, 235, 245)  'RGB(254, 254, 254)
            ctl.ForeColor = RGB(30, 30, 30)
        End If
    Next ctl
    MsgBox DestTaslak & TaslakFile & " directory cannot be accessed. The folder(s) and/or file(s) associated with this path might have been renamed."
    GoTo Son
End If

'Close the all Word application
Call OpenWordControl

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
'On Error GoTo 0

objWord.Documents.Open FileName:=DestTaslak & TaslakFile
Set objDoc = GetObject(DestTaslak & TaslakFile)
'objDoc.ActiveWindow.Visible = False

'DosyaNo getir getir
Kurum_ANoStr = objDoc.Tables(1).Cell(Row:=4, Column:=1).Range.Text

objDoc.Close SaveChanges:=False
objWord.Visible = False
'Nesneleri temizle
Set objWord = Nothing
Set objDoc = Nothing

StrDosyaNo = Mid(Kurum_ANoStr, InStrRev(Kurum_ANoStr, "-") + 1, Len(Kurum_ANoStr))
'MsgBox Left(StrDosyaNo, InStr(StrDosyaNo, "/") - 1)
StrDosyaNo = Left(StrDosyaNo, InStr(StrDosyaNo, "/") - 1)
'MsgBox StrDosyaNo
'Kurum_ANoStr = Mid(Kurum_ANoStr, InStr(Kurum_ANoStr, "KURUM_A") + 5, Len(Kurum_ANoStr))
'Kurum_ANoStr = Left(Kurum_ANoStr, 8)
Kurum_ANoStr = Left(Kurum_ANoStr, 15)
'MsgBox Kurum_ANoStr

'_____________________


'Etiket bilgileri
TextKURUM_A.Value = "ORGANIZATION A"
TextSube.Value = ThisWorkbook.Worksheets(2).Cells(6, 99).Value & " Unit"
TextTeskilatDosyaNoBelgeNo.Value = Kurum_ANoStr & "-" & StrDosyaNo & "/" & Cells(IlkSiraGlobal, 76).Value

Call Rapor1MuhatapTema

If BoolKucukZarf = True Then
    TextGonderiSekli.Value = "A N O N Y M O U S"
End If

If BoolBuyukZarf = True Then
    TextGonderiSekli.Value = Right(Cells(IlkSiraGlobal, 67).Value, Len(Cells(IlkSiraGlobal, 67).Value) - InStr(Cells(IlkSiraGlobal, 67).Value, "/"))
End If


Son:

End Sub


Sub RaporKucukZarf()
Dim ctl As MSForms.Control
Dim i As Long, j As Long

Dim AutoPath As String, DestTaslak As String, OpenControl As String
Dim TaslakFile As String, objWord As Object, objDoc As Object, Kurum_ANoStr As String
Dim StrDosyaNo As String


Call EtiketAlanlariReset
  
'__________

If ComboGetir.Value <> "" Then
    Set IlkSiraBulGlobal = Range("CM7:CM100000").Find(What:=ComboGetir.Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    Set SonSiraBulGlobal = Range("CN7:CN100000").Find(What:=ComboGetir.Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    If Not IlkSiraBulGlobal Is Nothing Then
        IlkSiraGlobal = IlkSiraBulGlobal.Row
    Else
        For Each ctl In UstMenuFrameAlt.Controls
            If TypeName(ctl) = "Label" Then
                ctl.BackColor = RGB(225, 235, 245)  'RGB(254, 254, 254)
                ctl.ForeColor = RGB(30, 30, 30)
            End If
        Next ctl
        MsgBox "The operation cannot be completed since the entered serial number was not found.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    If Not SonSiraBulGlobal Is Nothing Then
        SonSiraGlobal = SonSiraBulGlobal.Row
    Else
        For Each ctl In UstMenuFrameAlt.Controls
            If TypeName(ctl) = "Label" Then
                ctl.BackColor = RGB(225, 235, 245)  'RGB(254, 254, 254)
                ctl.ForeColor = RGB(30, 30, 30)
            End If
        Next ctl
        MsgBox "The operation cannot be completed since the entered serial number was not found.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
Else
    For Each ctl In UstMenuFrameAlt.Controls
        If TypeName(ctl) = "Label" Then
            ctl.BackColor = RGB(225, 235, 245)  'RGB(254, 254, 254)
            ctl.ForeColor = RGB(30, 30, 30)
        End If
    Next ctl
    MsgBox "Please select the serial number for the module you wish to operate on from the drop-down list located in the upper-left corner.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

'__________


AutoPath = ThisWorkbook.Path
DestTaslak = AutoPath & "\System Files\System Templates\Report 2 Cover Letter Templates\"
TaslakFile = "Report 2 Cover Letter.docm"

'System Files klasör adını kontrol et.
If Not Dir(AutoPath & "\System Files\", vbDirectory) <> vbNullString Then
    For Each ctl In UstMenuFrameAlt.Controls
        If TypeName(ctl) = "Label" Then
            ctl.BackColor = RGB(225, 235, 245)  'RGB(254, 254, 254)
            ctl.ForeColor = RGB(30, 30, 30)
        End If
    Next ctl
    MsgBox AutoPath & "\System Files\" & " directory cannot be accessed. The folder named 'System Files' might have been renamed or deleted."
    GoTo Son
End If
If Not Dir(DestTaslak & TaslakFile, vbDirectory) <> vbNullString Then
    For Each ctl In UstMenuFrameAlt.Controls
        If TypeName(ctl) = "Label" Then
            ctl.BackColor = RGB(225, 235, 245)  'RGB(254, 254, 254)
            ctl.ForeColor = RGB(30, 30, 30)
        End If
    Next ctl
    MsgBox DestTaslak & TaslakFile & " directory cannot be accessed. The folder(s) and/or file(s) associated with this path might have been renamed."
    GoTo Son
End If

'Close the all Word application
Call OpenWordControl

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
'On Error GoTo 0

objWord.Documents.Open FileName:=DestTaslak & TaslakFile
Set objDoc = GetObject(DestTaslak & TaslakFile)
'objDoc.ActiveWindow.Visible = False

'DosyaNo getir
Kurum_ANoStr = objDoc.Tables(1).Cell(Row:=4, Column:=1).Range.Text

objDoc.Close SaveChanges:=False
objWord.Visible = False
'Nesneleri temizle
Set objWord = Nothing
Set objDoc = Nothing

StrDosyaNo = Mid(Kurum_ANoStr, InStrRev(Kurum_ANoStr, "-") + 1, Len(Kurum_ANoStr))
'MsgBox Left(StrDosyaNo, InStr(StrDosyaNo, "/") - 1)
StrDosyaNo = Left(StrDosyaNo, InStr(StrDosyaNo, "/") - 1)
'MsgBox StrDosyaNo
'Kurum_ANoStr = Mid(Kurum_ANoStr, InStr(Kurum_ANoStr, "KURUM_A") + 5, Len(Kurum_ANoStr))
'Kurum_ANoStr = Left(Kurum_ANoStr, 8)
Kurum_ANoStr = Left(Kurum_ANoStr, 15)
'MsgBox Kurum_ANoStr

'_____________________


'Etiket bilgileri
TextKURUM_A.Value = "ORGANIZATION A"
TextSube.Value = ThisWorkbook.Worksheets(2).Cells(6, 99).Value & " Unit"
TextTeskilatDosyaNoBelgeNo.Value = Kurum_ANoStr & "-" & StrDosyaNo & "/" & Cells(IlkSiraGlobal, 84).Value

Call RaporMuhatapTema

If BoolKucukZarf = True Then
    TextGonderiSekli.Value = "A N O N Y M O U S"
End If

If BoolBuyukZarf = True Then
    TextGonderiSekli.Value = Right(Cells(IlkSiraGlobal, 75).Value, Len(Cells(IlkSiraGlobal, 75).Value) - InStr(Cells(IlkSiraGlobal, 75).Value, "/"))
End If

Son:

End Sub

Sub BilgilendirmeKucukZarf()
Dim ctl As MSForms.Control
Dim i As Long, j As Long

Dim AutoPath As String, DestTaslak As String, OpenControl As String
Dim TaslakFile As String, objWord As Object, objDoc As Object, Kurum_ANoStr As String
Dim StrDosyaNo As String


Call EtiketAlanlariReset
  
'__________

If ComboGetir.Value <> "" Then
    Set IlkSiraBulGlobal = Range("CM7:CM100000").Find(What:=ComboGetir.Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    Set SonSiraBulGlobal = Range("CN7:CN100000").Find(What:=ComboGetir.Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    If Not IlkSiraBulGlobal Is Nothing Then
        IlkSiraGlobal = IlkSiraBulGlobal.Row
    Else
        For Each ctl In UstMenuFrameAlt.Controls
            If TypeName(ctl) = "Label" Then
                ctl.BackColor = RGB(225, 235, 245)  'RGB(254, 254, 254)
                ctl.ForeColor = RGB(30, 30, 30)
            End If
        Next ctl
        MsgBox "The operation cannot be completed since the entered serial number was not found.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    If Not SonSiraBulGlobal Is Nothing Then
        SonSiraGlobal = SonSiraBulGlobal.Row
    Else
        For Each ctl In UstMenuFrameAlt.Controls
            If TypeName(ctl) = "Label" Then
                ctl.BackColor = RGB(225, 235, 245)  'RGB(254, 254, 254)
                ctl.ForeColor = RGB(30, 30, 30)
            End If
        Next ctl
        MsgBox "The operation cannot be completed since the entered serial number was not found.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
Else
    For Each ctl In UstMenuFrameAlt.Controls
        If TypeName(ctl) = "Label" Then
            ctl.BackColor = RGB(225, 235, 245)  'RGB(254, 254, 254)
            ctl.ForeColor = RGB(30, 30, 30)
        End If
    Next ctl
    MsgBox "Please select the serial number for the module you wish to operate on from the drop-down list located in the upper-left corner.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

'__________


AutoPath = ThisWorkbook.Path
DestTaslak = AutoPath & "\System Files\System Templates\Report 2 Cover Letter Templates\"
TaslakFile = "Informative Cover Letter.docm"

'System Files klasör adını kontrol et.
If Not Dir(AutoPath & "\System Files\", vbDirectory) <> vbNullString Then
    For Each ctl In UstMenuFrameAlt.Controls
        If TypeName(ctl) = "Label" Then
            ctl.BackColor = RGB(225, 235, 245)  'RGB(254, 254, 254)
            ctl.ForeColor = RGB(30, 30, 30)
        End If
    Next ctl
    MsgBox AutoPath & "\System Files\" & " directory cannot be accessed. The folder named 'System Files' might have been renamed or deleted."
    GoTo Son
End If
If Not Dir(DestTaslak & TaslakFile, vbDirectory) <> vbNullString Then
    For Each ctl In UstMenuFrameAlt.Controls
        If TypeName(ctl) = "Label" Then
            ctl.BackColor = RGB(225, 235, 245)  'RGB(254, 254, 254)
            ctl.ForeColor = RGB(30, 30, 30)
        End If
    Next ctl
    MsgBox DestTaslak & TaslakFile & " directory cannot be accessed. The folder(s) and/or file(s) associated with this path might have been renamed."
    GoTo Son
End If

'Close the all Word application
Call OpenWordControl

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
'On Error GoTo 0

objWord.Documents.Open FileName:=DestTaslak & TaslakFile
Set objDoc = GetObject(DestTaslak & TaslakFile)
'objDoc.ActiveWindow.Visible = False

'DosyaNo getir
Kurum_ANoStr = objDoc.Tables(1).Cell(Row:=4, Column:=1).Range.Text

objDoc.Close SaveChanges:=False
objWord.Visible = False
'Nesneleri temizle
Set objWord = Nothing
Set objDoc = Nothing

StrDosyaNo = Mid(Kurum_ANoStr, InStrRev(Kurum_ANoStr, "-") + 1, Len(Kurum_ANoStr))
'MsgBox Left(StrDosyaNo, InStr(StrDosyaNo, "/") - 1)
StrDosyaNo = Left(StrDosyaNo, InStr(StrDosyaNo, "/") - 1)
'MsgBox StrDosyaNo
'Kurum_ANoStr = Mid(Kurum_ANoStr, InStr(Kurum_ANoStr, "KURUM_A") + 5, Len(Kurum_ANoStr))
'Kurum_ANoStr = Left(Kurum_ANoStr, 8)
Kurum_ANoStr = Left(Kurum_ANoStr, 15)
'MsgBox Kurum_ANoStr

'_____________________


'Etiket bilgileri
TextKURUM_A.Value = "ORGANIZATION A"
TextSube.Value = ThisWorkbook.Worksheets(2).Cells(6, 99).Value & " Unit"
TextTeskilatDosyaNoBelgeNo.Value = Kurum_ANoStr & "-" & StrDosyaNo & "/" & Cells(IlkSiraGlobal, 84).Value

Call BilgilendirmeMuhatapTema

If BoolKucukZarf = True Then
    TextGonderiSekli.Value = "A N O N Y M O U S"
End If

If BoolBuyukZarf = True Then
    TextGonderiSekli.Value = Right(Cells(IlkSiraGlobal, 81).Value, Len(Cells(IlkSiraGlobal, 81).Value) - InStr(Cells(IlkSiraGlobal, 81).Value, "/"))
End If

Son:

End Sub

Sub Rapor3_1KucukZarf()
Dim ctl As MSForms.Control
Dim i As Long, j As Long

Dim AutoPath As String, DestTaslak As String, OpenControl As String
Dim TaslakFile As String, objWord As Object, objDoc As Object, Kurum_ANoStr As String
Dim StrDosyaNo As String


Call EtiketAlanlariReset
  
'__________

If ComboGetir.Value <> "" Then
    Set IlkSiraBulGlobal = Range("FG7:FG100000").Find(What:=ComboGetir.Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    Set SonSiraBulGlobal = Range("FH7:FH100000").Find(What:=ComboGetir.Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    If Not IlkSiraBulGlobal Is Nothing Then
        IlkSiraGlobal = IlkSiraBulGlobal.Row
    Else
        For Each ctl In UstMenuFrameAlt.Controls
            If TypeName(ctl) = "Label" Then
                ctl.BackColor = RGB(225, 235, 245)  'RGB(254, 254, 254)
                ctl.ForeColor = RGB(30, 30, 30)
            End If
        Next ctl
        MsgBox "The operation cannot be completed since the entered serial number was not found.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    If Not SonSiraBulGlobal Is Nothing Then
        SonSiraGlobal = SonSiraBulGlobal.Row
    Else
        For Each ctl In UstMenuFrameAlt.Controls
            If TypeName(ctl) = "Label" Then
                ctl.BackColor = RGB(225, 235, 245)  'RGB(254, 254, 254)
                ctl.ForeColor = RGB(30, 30, 30)
            End If
        Next ctl
        MsgBox "The operation cannot be completed since the entered serial number was not found.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
Else
    For Each ctl In UstMenuFrameAlt.Controls
        If TypeName(ctl) = "Label" Then
            ctl.BackColor = RGB(225, 235, 245)  'RGB(254, 254, 254)
            ctl.ForeColor = RGB(30, 30, 30)
        End If
    Next ctl
    MsgBox "Please select the serial number for the module you wish to operate on from the drop-down list located in the upper-left corner.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

'__________


AutoPath = ThisWorkbook.Path
DestTaslak = AutoPath & "\System Files\System Templates\Report 3 Cover Letter Templates\"
If Cells(IlkSiraGlobal, 100).Value = "Type A" Then
    TaslakFile = "Report 3.1 – Type A Cover Letter.docm"
Else
    TaslakFile = "Report 3.1 – Type B Cover Letter.docm"
End If

'System Files klasör adını kontrol et.
If Not Dir(AutoPath & "\System Files\", vbDirectory) <> vbNullString Then
    For Each ctl In UstMenuFrameAlt.Controls
        If TypeName(ctl) = "Label" Then
            ctl.BackColor = RGB(225, 235, 245)  'RGB(254, 254, 254)
            ctl.ForeColor = RGB(30, 30, 30)
        End If
    Next ctl
    MsgBox AutoPath & "\System Files\" & " directory cannot be accessed. The folder named 'System Files' might have been renamed or deleted."
    GoTo Son
End If
If Not Dir(DestTaslak & TaslakFile, vbDirectory) <> vbNullString Then
    For Each ctl In UstMenuFrameAlt.Controls
        If TypeName(ctl) = "Label" Then
            ctl.BackColor = RGB(225, 235, 245)  'RGB(254, 254, 254)
            ctl.ForeColor = RGB(30, 30, 30)
        End If
    Next ctl
    MsgBox DestTaslak & TaslakFile & " directory cannot be accessed. The folder(s) and/or file(s) associated with this path might have been renamed."
    GoTo Son
End If

'Close the all Word application
Call OpenWordControl

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
'On Error GoTo 0

objWord.Documents.Open FileName:=DestTaslak & TaslakFile
Set objDoc = GetObject(DestTaslak & TaslakFile)
'objDoc.ActiveWindow.Visible = False

'DosyaNo getir
Kurum_ANoStr = objDoc.Tables(1).Cell(Row:=4, Column:=1).Range.Text

objDoc.Close SaveChanges:=False
objWord.Visible = False
'Nesneleri temizle
Set objWord = Nothing
Set objDoc = Nothing

StrDosyaNo = Mid(Kurum_ANoStr, InStrRev(Kurum_ANoStr, "-") + 1, Len(Kurum_ANoStr))
'MsgBox Left(StrDosyaNo, InStr(StrDosyaNo, "/") - 1)
StrDosyaNo = Left(StrDosyaNo, InStr(StrDosyaNo, "/") - 1)
'MsgBox StrDosyaNo
'Kurum_ANoStr = Mid(Kurum_ANoStr, InStr(Kurum_ANoStr, "KURUM_A") + 5, Len(Kurum_ANoStr))
'Kurum_ANoStr = Left(Kurum_ANoStr, 8)
Kurum_ANoStr = Left(Kurum_ANoStr, 15)
'MsgBox Kurum_ANoStr

'_____________________


'Etiket bilgileri
TextKURUM_A.Value = "ORGANIZATION A"
TextSube.Value = ThisWorkbook.Worksheets(2).Cells(6, 99).Value & " Unit"
TextTeskilatDosyaNoBelgeNo.Value = Kurum_ANoStr & "-" & StrDosyaNo & "/" & Cells(IlkSiraGlobal, 156).Value

Call Rapor3_1MuhatapTema

If BoolKucukZarf = True Then
    TextGonderiSekli.Value = "A N O N Y M O U S"
End If

If BoolBuyukZarf = True Then
    TextGonderiSekli.Value = Right(Cells(IlkSiraGlobal, 149).Value, Len(Cells(IlkSiraGlobal, 149).Value) - InStr(Cells(IlkSiraGlobal, 149).Value, "/"))
End If

Son:

End Sub

Sub Rapor3_2KucukZarf()
Dim ctl As MSForms.Control
Dim i As Long, j As Long

Dim AutoPath As String, DestTaslak As String, OpenControl As String
Dim TaslakFile As String, objWord As Object, objDoc As Object, Kurum_ANoStr As String
Dim StrDosyaNo As String


Call EtiketAlanlariReset
  
'__________

If ComboGetir.Value <> "" Then
    Set IlkSiraBulGlobal = Range("FG7:FG100000").Find(What:=ComboGetir.Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    Set SonSiraBulGlobal = Range("FH7:FH100000").Find(What:=ComboGetir.Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    If Not IlkSiraBulGlobal Is Nothing Then
        IlkSiraGlobal = IlkSiraBulGlobal.Row
    Else
        For Each ctl In UstMenuFrameAlt.Controls
            If TypeName(ctl) = "Label" Then
                ctl.BackColor = RGB(225, 235, 245)  'RGB(254, 254, 254)
                ctl.ForeColor = RGB(30, 30, 30)
            End If
        Next ctl
        MsgBox "The operation cannot be completed since the entered serial number was not found.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    If Not SonSiraBulGlobal Is Nothing Then
        SonSiraGlobal = SonSiraBulGlobal.Row
    Else
        For Each ctl In UstMenuFrameAlt.Controls
            If TypeName(ctl) = "Label" Then
                ctl.BackColor = RGB(225, 235, 245)  'RGB(254, 254, 254)
                ctl.ForeColor = RGB(30, 30, 30)
            End If
        Next ctl
        MsgBox "The operation cannot be completed since the entered serial number was not found.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
Else
    For Each ctl In UstMenuFrameAlt.Controls
        If TypeName(ctl) = "Label" Then
            ctl.BackColor = RGB(225, 235, 245)  'RGB(254, 254, 254)
            ctl.ForeColor = RGB(30, 30, 30)
        End If
    Next ctl
    MsgBox "Please select the serial number for the module you wish to operate on from the drop-down list located in the upper-left corner.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

'__________

AutoPath = ThisWorkbook.Path
DestTaslak = AutoPath & "\System Files\System Templates\Report 3 Cover Letter Templates\"
If Cells(IlkSiraGlobal, 28).Value = "Type A" Then
    TaslakFile = "Report 3.2 – Type A Cover Letter.docm"
Else
    TaslakFile = "Report 3.2 – Type B Cover Letter.docm"
End If

'System Files klasör adını kontrol et.
If Not Dir(AutoPath & "\System Files\", vbDirectory) <> vbNullString Then
    For Each ctl In UstMenuFrameAlt.Controls
        If TypeName(ctl) = "Label" Then
            ctl.BackColor = RGB(225, 235, 245)  'RGB(254, 254, 254)
            ctl.ForeColor = RGB(30, 30, 30)
        End If
    Next ctl
    MsgBox AutoPath & "\System Files\" & " directory cannot be accessed. The folder named 'System Files' might have been renamed or deleted."
    GoTo Son
End If
If Not Dir(DestTaslak & TaslakFile, vbDirectory) <> vbNullString Then
    For Each ctl In UstMenuFrameAlt.Controls
        If TypeName(ctl) = "Label" Then
            ctl.BackColor = RGB(225, 235, 245)  'RGB(254, 254, 254)
            ctl.ForeColor = RGB(30, 30, 30)
        End If
    Next ctl
    MsgBox DestTaslak & TaslakFile & " directory cannot be accessed. The folder(s) and/or file(s) associated with this path might have been renamed."
    GoTo Son
End If

'Close the all Word application
Call OpenWordControl

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
'On Error GoTo 0

objWord.Documents.Open FileName:=DestTaslak & TaslakFile
Set objDoc = GetObject(DestTaslak & TaslakFile)
'objDoc.ActiveWindow.Visible = False

'DosyaNo getir
Kurum_ANoStr = objDoc.Tables(1).Cell(Row:=4, Column:=1).Range.Text

objDoc.Close SaveChanges:=False
objWord.Visible = False
'Nesneleri temizle
Set objWord = Nothing
Set objDoc = Nothing

StrDosyaNo = Mid(Kurum_ANoStr, InStrRev(Kurum_ANoStr, "-") + 1, Len(Kurum_ANoStr))
'MsgBox Left(StrDosyaNo, InStr(StrDosyaNo, "/") - 1)
StrDosyaNo = Left(StrDosyaNo, InStr(StrDosyaNo, "/") - 1)
'MsgBox StrDosyaNo
'Kurum_ANoStr = Mid(Kurum_ANoStr, InStr(Kurum_ANoStr, "KURUM_A") + 5, Len(Kurum_ANoStr))
'Kurum_ANoStr = Left(Kurum_ANoStr, 8)
Kurum_ANoStr = Left(Kurum_ANoStr, 15)
'MsgBox Kurum_ANoStr

'_____________________


'Etiket bilgileri
TextKURUM_A.Value = "ORGANIZATION A"
TextSube.Value = ThisWorkbook.Worksheets(2).Cells(6, 99).Value & " Unit"
TextTeskilatDosyaNoBelgeNo.Value = Kurum_ANoStr & "-" & StrDosyaNo & "/" & Cells(IlkSiraGlobal, 84).Value

Call Rapor3_2MuhatapTema

If BoolKucukZarf = True Then
    TextGonderiSekli.Value = "A N O N Y M O U S"
End If

If BoolBuyukZarf = True Then
    TextGonderiSekli.Value = Right(Cells(IlkSiraGlobal, 71).Value, Len(Cells(IlkSiraGlobal, 71).Value) - InStr(Cells(IlkSiraGlobal, 71).Value, "/"))
End If
    
Son:

End Sub


Private Sub LblAltMenuButton3_Click()
Dim ws As Object, i As Integer
Dim ctl As MSForms.Control

If LblAltMenuButton3.Caption = "" Then
    Exit Sub
End If

Application.DisplayAlerts = False
Application.EnableEvents = False
Application.ScreenUpdating = False

'ThisWorkbook.Unprotect "123"

BoolBuyukZarf = True

For Each ctl In UstMenuFrameAlt.Controls
    If TypeName(ctl) = "Label" Then
        ctl.BackColor = RGB(225, 235, 245)  'RGB(254, 254, 254)
        ctl.ForeColor = RGB(30, 30, 30)
    End If
Next ctl
If LblAltMenuButton3.Caption <> "" Then
    LblAltMenuButton3.BackColor = RGB(180, 210, 240)
    LblAltMenuButton3.ForeColor = RGB(30, 30, 30)
End If

Call EtiketAlanlariReset

'NOT: Küçük Zarf ile tek farkı gizli yerine göndri şeklinin yer alması olduğundan
'BoolBuyukZarf ve BoolKucukZarf mantıksal değişkenleri yardımıyla sadece KucukZarf prosedürleri kullanıldı.

If Rapor1.BackColor = RGB(180, 210, 240) Then
    'MsgBox "Rapor1"
    'Büyük Zarf
    Call Rapor1KucukZarf
ElseIf Rapor.BackColor = RGB(180, 210, 240) Then
    'MsgBox "Rapor"
    'Büyük Zarf
    Call RaporKucukZarf
ElseIf Rapor2_2.BackColor = RGB(180, 210, 240) Then
    'MsgBox "Rapor2_2"
    'Büyük Zarf
    Call BilgilendirmeKucukZarf
ElseIf Rapor3_1.BackColor = RGB(180, 210, 240) Then
    'MsgBox "Report 3.1"
    'Büyük Zarf
    Call Rapor3_1KucukZarf
ElseIf Rapor3_2.BackColor = RGB(180, 210, 240) Then
    'MsgBox "Report 3.2"
    'Büyük Zarf
    Call Rapor3_2KucukZarf
End If


BoolBuyukZarf = False

'ThisWorkbook.Protect "123"

Application.DisplayAlerts = True
Application.EnableEvents = True
Application.ScreenUpdating = True

End Sub

Private Sub LblAltMenuButton4_Click()
Dim ws As Object, i As Integer
Dim ctl As MSForms.Control

If LblAltMenuButton4.Caption = "" Then
    Exit Sub
End If

Application.DisplayAlerts = False
Application.EnableEvents = False
Application.ScreenUpdating = False

'ThisWorkbook.Unprotect "123"

BoolFinansalBirimKucukZarf = True

For Each ctl In UstMenuFrameAlt.Controls
    If TypeName(ctl) = "Label" Then
        ctl.BackColor = RGB(225, 235, 245)  'RGB(254, 254, 254)
        ctl.ForeColor = RGB(30, 30, 30)
    End If
Next ctl
If LblAltMenuButton4.Caption <> "" Then
    LblAltMenuButton4.BackColor = RGB(180, 210, 240)
    LblAltMenuButton4.ForeColor = RGB(30, 30, 30)
End If

Call EtiketAlanlariReset


If Rapor3_2.BackColor = RGB(180, 210, 240) Then
    'MsgBox "Report 3.2"
    'FinansalBirim küçük Zarf
    Call FinansalBirimKucukZarf
ElseIf Rapor2_2.BackColor = RGB(180, 210, 240) Then
    'MsgBox "Rapor2_2"
    'XXXMud Gelen Tutanak2 Zarfı
    Call XXXMudGelenTutanak2Zarfi
End If


BoolFinansalBirimKucukZarf = False

'ThisWorkbook.Protect "123"

Application.DisplayAlerts = True
Application.EnableEvents = True
Application.ScreenUpdating = True

End Sub

Sub FinansalBirimKucukZarf()
Dim ctl As MSForms.Control
Dim i As Long, j As Long

Dim AutoPath As String, DestTaslak As String, OpenControl As String
Dim TaslakFile As String, objWord As Object, objDoc As Object, Kurum_ANoStr As String
Dim StrDosyaNo As String


Call EtiketAlanlariReset
  
'__________

If ComboGetir.Value <> "" Then
    Set IlkSiraBulGlobal = Range("FG7:FG100000").Find(What:=ComboGetir.Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    Set SonSiraBulGlobal = Range("FH7:FH100000").Find(What:=ComboGetir.Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    If Not IlkSiraBulGlobal Is Nothing Then
        IlkSiraGlobal = IlkSiraBulGlobal.Row
    Else
        For Each ctl In UstMenuFrameAlt.Controls
            If TypeName(ctl) = "Label" Then
                ctl.BackColor = RGB(225, 235, 245)  'RGB(254, 254, 254)
                ctl.ForeColor = RGB(30, 30, 30)
            End If
        Next ctl
        MsgBox "The operation cannot be completed since the entered serial number was not found.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    If Not SonSiraBulGlobal Is Nothing Then
        SonSiraGlobal = SonSiraBulGlobal.Row
    Else
        For Each ctl In UstMenuFrameAlt.Controls
            If TypeName(ctl) = "Label" Then
                ctl.BackColor = RGB(225, 235, 245)  'RGB(254, 254, 254)
                ctl.ForeColor = RGB(30, 30, 30)
            End If
        Next ctl
        MsgBox "The operation cannot be completed since the entered serial number was not found.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
Else
    For Each ctl In UstMenuFrameAlt.Controls
        If TypeName(ctl) = "Label" Then
            ctl.BackColor = RGB(225, 235, 245)  'RGB(254, 254, 254)
            ctl.ForeColor = RGB(30, 30, 30)
        End If
    Next ctl
    MsgBox "Please select the serial number for the module you wish to operate on from the drop-down list located in the upper-left corner.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

'__________

AutoPath = ThisWorkbook.Path
DestTaslak = AutoPath & "\System Files\System Templates\Report 3 Cover Letter Templates\"
If Cells(IlkSiraGlobal, 28).Value = "Type A" Then
    TaslakFile = "Report 3.2 – Type A Cover Letter – Financial Unit.docm"
Else
    TaslakFile = "Report 3.2 – Type B Cover Letter – Financial Unit.docm"
End If

'System Files klasör adını kontrol et.
If Not Dir(AutoPath & "\System Files\", vbDirectory) <> vbNullString Then
    For Each ctl In UstMenuFrameAlt.Controls
        If TypeName(ctl) = "Label" Then
            ctl.BackColor = RGB(225, 235, 245)  'RGB(254, 254, 254)
            ctl.ForeColor = RGB(30, 30, 30)
        End If
    Next ctl
    MsgBox AutoPath & "\System Files\" & " directory cannot be accessed. The folder named 'System Files' might have been renamed or deleted."
    GoTo Son
End If
If Not Dir(DestTaslak & TaslakFile, vbDirectory) <> vbNullString Then
    For Each ctl In UstMenuFrameAlt.Controls
        If TypeName(ctl) = "Label" Then
            ctl.BackColor = RGB(225, 235, 245)  'RGB(254, 254, 254)
            ctl.ForeColor = RGB(30, 30, 30)
        End If
    Next ctl
    MsgBox DestTaslak & TaslakFile & " directory cannot be accessed. The folder(s) and/or file(s) associated with this path might have been renamed."
    GoTo Son
End If


'Close the all Word application
Call OpenWordControl

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
'On Error GoTo 0

objWord.Documents.Open FileName:=DestTaslak & TaslakFile
Set objDoc = GetObject(DestTaslak & TaslakFile)
'objDoc.ActiveWindow.Visible = False

'DosyaNo getir
Kurum_ANoStr = objDoc.Tables(1).Cell(Row:=4, Column:=1).Range.Text

objDoc.Close SaveChanges:=False
objWord.Visible = False
'Nesneleri temizle
Set objWord = Nothing
Set objDoc = Nothing

StrDosyaNo = Mid(Kurum_ANoStr, InStrRev(Kurum_ANoStr, "-") + 1, Len(Kurum_ANoStr))
'MsgBox Left(StrDosyaNo, InStr(StrDosyaNo, "/") - 1)
StrDosyaNo = Left(StrDosyaNo, InStr(StrDosyaNo, "/") - 1)
'MsgBox StrDosyaNo
'Kurum_ANoStr = Mid(Kurum_ANoStr, InStr(Kurum_ANoStr, "KURUM_A") + 5, Len(Kurum_ANoStr))
'Kurum_ANoStr = Left(Kurum_ANoStr, 8)
Kurum_ANoStr = Left(Kurum_ANoStr, 15)
'MsgBox Kurum_ANoStr

'_____________________


'Etiket bilgileri
TextKURUM_A.Value = "ORGANIZATION A"
TextSube.Value = ThisWorkbook.Worksheets(2).Cells(6, 99).Value & " Unit"
TextTeskilatDosyaNoBelgeNo.Value = Kurum_ANoStr & "-" & StrDosyaNo & "/" & Cells(IlkSiraGlobal, 76).Value

Call FinansalBirimMuhatapTema

If BoolFinansalBirimKucukZarf = True Then
    TextGonderiSekli.Value = "A N O N Y M O U S"
End If

If BoolFinansalBirimBuyukZarf = True Then
    TextGonderiSekli.Value = UCase(Replace(Replace(Cells(IlkSiraGlobal, 85).Value, "i", "I"), "ı", "I"))
End If
    
Son:

End Sub


Private Sub LblAltMenuButton5_Click()
Dim ws As Object, i As Integer
Dim ctl As MSForms.Control

If LblAltMenuButton5.Caption = "" Then
    Exit Sub
End If

Application.DisplayAlerts = False
Application.EnableEvents = False
Application.ScreenUpdating = False

'ThisWorkbook.Unprotect "123"

BoolFinansalBirimBuyukZarf = True

For Each ctl In UstMenuFrameAlt.Controls
    If TypeName(ctl) = "Label" Then
        ctl.BackColor = RGB(225, 235, 245)  'RGB(254, 254, 254)
        ctl.ForeColor = RGB(30, 30, 30)
    End If
Next ctl
If LblAltMenuButton5.Caption <> "" Then
    LblAltMenuButton5.BackColor = RGB(180, 210, 240)
    LblAltMenuButton5.ForeColor = RGB(30, 30, 30)
End If

Call EtiketAlanlariReset

'NOT: Küçük Zarf ile tek farkı gizli yerine gönderi şeklinin yer alması olduğundan
'BoolFinansalBirimBuyukZarf ve BoolFinansalBirimKucukZarf mantıksal değişkenleri yardımıyla sadece KucukZarf prosedürleri kullanıldı.

If Rapor3_2.BackColor = RGB(180, 210, 240) Then
    'MsgBox "Report 3.2"
    'FinansalBirim büyük Zarf
    Call FinansalBirimKucukZarf
ElseIf Rapor2_2.BackColor = RGB(180, 210, 240) Then
    'MsgBox "Rapor2_2"
    'Kurum Tutanak2 Zarfı
    Call KurumTutanak2Zarfi
End If


BoolFinansalBirimBuyukZarf = False



'ThisWorkbook.Protect "123"

Application.DisplayAlerts = True
Application.EnableEvents = True
Application.ScreenUpdating = True

End Sub

Private Sub LblAltMenuButton6_Click()
Dim ws As Object, i As Integer
Dim ctl As MSForms.Control

If LblAltMenuButton6.Caption = "" Then
    Exit Sub
End If

Application.DisplayAlerts = False
Application.EnableEvents = False
Application.ScreenUpdating = False

'ThisWorkbook.Unprotect "123"

BoolKucukZarf = True

For Each ctl In UstMenuFrameAlt.Controls
    If TypeName(ctl) = "Label" Then
        ctl.BackColor = RGB(225, 235, 245)  'RGB(254, 254, 254)
        ctl.ForeColor = RGB(30, 30, 30)
    End If
Next ctl
If LblAltMenuButton6.Caption <> "" Then
    LblAltMenuButton6.BackColor = RGB(180, 210, 240)
    LblAltMenuButton6.ForeColor = RGB(30, 30, 30)
End If

Call EtiketAlanlariReset


If Rapor2_2.BackColor = RGB(180, 210, 240) Then
    'MsgBox "Rapor2_2"
    'Sonuç Küçük Zarfı
    Call SonucKucukZarf
End If


BoolKucukZarf = False

'ThisWorkbook.Protect "123"

Application.DisplayAlerts = True
Application.EnableEvents = True
Application.ScreenUpdating = True

End Sub

Sub SonucKucukZarf()
Dim ctl As MSForms.Control
Dim i As Long, j As Long

Dim AutoPath As String, DestTaslak As String, OpenControl As String
Dim TaslakFile As String, objWord As Object, objDoc As Object, Kurum_ANoStr As String
Dim StrDosyaNo As String


Call EtiketAlanlariReset
  
'__________

If ComboGetir.Value <> "" Then
    Set IlkSiraBulGlobal = Range("CM7:CM100000").Find(What:=ComboGetir.Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    Set SonSiraBulGlobal = Range("CN7:CN100000").Find(What:=ComboGetir.Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    If Not IlkSiraBulGlobal Is Nothing Then
        IlkSiraGlobal = IlkSiraBulGlobal.Row
    Else
        For Each ctl In UstMenuFrameAlt.Controls
            If TypeName(ctl) = "Label" Then
                ctl.BackColor = RGB(225, 235, 245)  'RGB(254, 254, 254)
                ctl.ForeColor = RGB(30, 30, 30)
            End If
        Next ctl
        MsgBox "The operation cannot be completed since the entered serial number was not found.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    If Not SonSiraBulGlobal Is Nothing Then
        SonSiraGlobal = SonSiraBulGlobal.Row
    Else
        For Each ctl In UstMenuFrameAlt.Controls
            If TypeName(ctl) = "Label" Then
                ctl.BackColor = RGB(225, 235, 245)  'RGB(254, 254, 254)
                ctl.ForeColor = RGB(30, 30, 30)
            End If
        Next ctl
        MsgBox "The operation cannot be completed since the entered serial number was not found.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
Else
    For Each ctl In UstMenuFrameAlt.Controls
        If TypeName(ctl) = "Label" Then
            ctl.BackColor = RGB(225, 235, 245)  'RGB(254, 254, 254)
            ctl.ForeColor = RGB(30, 30, 30)
        End If
    Next ctl
    MsgBox "Please select the serial number for the module you wish to operate on from the drop-down list located in the upper-left corner.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

'__________


AutoPath = ThisWorkbook.Path
DestTaslak = AutoPath & "\System Files\System Templates\Report 2 Cover Letter Templates\"
TaslakFile = "Final Cover Letter.docm"

'System Files klasör adını kontrol et.
If Not Dir(AutoPath & "\System Files\", vbDirectory) <> vbNullString Then
    For Each ctl In UstMenuFrameAlt.Controls
        If TypeName(ctl) = "Label" Then
            ctl.BackColor = RGB(225, 235, 245)  'RGB(254, 254, 254)
            ctl.ForeColor = RGB(30, 30, 30)
        End If
    Next ctl
    MsgBox AutoPath & "\System Files\" & " directory cannot be accessed. The folder named 'System Files' might have been renamed or deleted."
    GoTo Son
End If
If Not Dir(DestTaslak & TaslakFile, vbDirectory) <> vbNullString Then
    For Each ctl In UstMenuFrameAlt.Controls
        If TypeName(ctl) = "Label" Then
            ctl.BackColor = RGB(225, 235, 245)  'RGB(254, 254, 254)
            ctl.ForeColor = RGB(30, 30, 30)
        End If
    Next ctl
    MsgBox DestTaslak & TaslakFile & " directory cannot be accessed. The folder(s) and/or file(s) associated with this path might have been renamed."
    GoTo Son
End If

'Close the all Word application
Call OpenWordControl

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
'On Error GoTo 0

objWord.Documents.Open FileName:=DestTaslak & TaslakFile
Set objDoc = GetObject(DestTaslak & TaslakFile)
'objDoc.ActiveWindow.Visible = False

'DosyaNo getir
Kurum_ANoStr = objDoc.Tables(1).Cell(Row:=4, Column:=1).Range.Text

objDoc.Close SaveChanges:=False
objWord.Visible = False
'Nesneleri temizle
Set objWord = Nothing
Set objDoc = Nothing

StrDosyaNo = Mid(Kurum_ANoStr, InStrRev(Kurum_ANoStr, "-") + 1, Len(Kurum_ANoStr))
'MsgBox Left(StrDosyaNo, InStr(StrDosyaNo, "/") - 1)
StrDosyaNo = Left(StrDosyaNo, InStr(StrDosyaNo, "/") - 1)
'MsgBox StrDosyaNo
'Kurum_ANoStr = Mid(Kurum_ANoStr, InStr(Kurum_ANoStr, "KURUM_A") + 5, Len(Kurum_ANoStr))
'Kurum_ANoStr = Left(Kurum_ANoStr, 8)
Kurum_ANoStr = Left(Kurum_ANoStr, 15)
'MsgBox Kurum_ANoStr

'_____________________


'Etiket bilgileri
TextKURUM_A.Value = "ORGANIZATION A"
TextSube.Value = ThisWorkbook.Worksheets(2).Cells(6, 99).Value & " Unit"
TextTeskilatDosyaNoBelgeNo.Value = Kurum_ANoStr & "-" & StrDosyaNo & "/" & Cells(IlkSiraGlobal, 178).Value

Call BilgilendirmeMuhatapTema

If BoolKucukZarf = True Then
    TextGonderiSekli.Value = "A N O N Y M O U S"
End If

If BoolBuyukZarf = True Then
    If Cells(IlkSiraGlobal, 75).Value <> "" And Cells(IlkSiraGlobal, 201).Value = "" Then 'Sadece 75 esas al
        TextGonderiSekli.Value = Right(Cells(IlkSiraGlobal, 75).Value, Len(Cells(IlkSiraGlobal, 75).Value) - InStr(Cells(IlkSiraGlobal, 75).Value, "/")) 'Giden Paket Tipi
    ElseIf Cells(IlkSiraGlobal, 75).Value = "" And Cells(IlkSiraGlobal, 201).Value <> "" Then 'Sadece 201 esas al
        TextGonderiSekli.Value = Right(Cells(IlkSiraGlobal, 201).Value, Len(Cells(IlkSiraGlobal, 201).Value) - InStr(Cells(IlkSiraGlobal, 201).Value, "/")) 'Giden Paket Tipi
    ElseIf Cells(IlkSiraGlobal, 75).Value <> "" And Cells(IlkSiraGlobal, 201).Value <> "" Then '75 ve 201, ancak 75 esas al
        TextGonderiSekli.Value = Right(Cells(IlkSiraGlobal, 75).Value, Len(Cells(IlkSiraGlobal, 75).Value) - InStr(Cells(IlkSiraGlobal, 75).Value, "/")) 'Giden Paket Tipi
    End If
End If

Son:

End Sub

Private Sub LblAltMenuButton7_Click()
Dim ws As Object, i As Integer
Dim ctl As MSForms.Control

If LblAltMenuButton7.Caption = "" Then
    Exit Sub
End If

Application.DisplayAlerts = False
Application.EnableEvents = False
Application.ScreenUpdating = False

'ThisWorkbook.Unprotect "123"

BoolBuyukZarf = True

For Each ctl In UstMenuFrameAlt.Controls
    If TypeName(ctl) = "Label" Then
        ctl.BackColor = RGB(225, 235, 245)  'RGB(254, 254, 254)
        ctl.ForeColor = RGB(30, 30, 30)
    End If
Next ctl
If LblAltMenuButton7.Caption <> "" Then
    LblAltMenuButton7.BackColor = RGB(180, 210, 240)
    LblAltMenuButton7.ForeColor = RGB(30, 30, 30)
End If

Call EtiketAlanlariReset


If Rapor2_2.BackColor = RGB(180, 210, 240) Then
    'MsgBox "Rapor2_2"
    'Sonuç Küçük Zarfı
    Call SonucKucukZarf
End If


BoolBuyukZarf = False

'ThisWorkbook.Protect "123"

Application.DisplayAlerts = True
Application.EnableEvents = True
Application.ScreenUpdating = True

End Sub


Sub ComboGetirResetRapor1()
Dim Say As Long, i As Long

On Error Resume Next ' Son Sıra numarası sayısal olmayan karakter içeriyorsa userform açılmıyor.
ComboGetir.Value = ""
ComboGetir.Clear
Say = ThisWorkbook.Worksheets(3).Range("E100000").End(xlUp).Row
If Say < 7 Then
    GoTo GetirBos
End If

'Getir liste değerleri
For i = Say To 7 Step -1
    If ThisWorkbook.Worksheets(3).Range("E" & i).Value <> "" Then
        With ComboGetir
            .AddItem (ThisWorkbook.Worksheets(3).Range("E" & i).Value)
        End With
    End If
Next i
GetirBos:

End Sub

Sub ComboGetirResetRapor()
Dim Say As Long, i As Long

On Error Resume Next ' Son Sıra numarası sayısal olmayan karakter içeriyorsa userform açılmıyor.
ComboGetir.Value = ""
ComboGetir.Clear
Say = ThisWorkbook.Worksheets(4).Range("E100000").End(xlUp).Row
If Say < 7 Then
    GoTo GetirBos
End If

'Getir liste değerleri
For i = Say To 7 Step -1
    If ThisWorkbook.Worksheets(4).Range("E" & i).Value <> "" Then
        If ThisWorkbook.Worksheets(4).Range("FR" & i) <> "Yes" Then
            With ComboGetir
                .AddItem (ThisWorkbook.Worksheets(4).Range("E" & i).Value)
            End With
        End If
    End If
Next i
GetirBos:

End Sub

Sub ComboGetirResetRapor2_2()
Dim Say As Long, i As Long

On Error Resume Next ' Son Sıra numarası sayısal olmayan karakter içeriyorsa userform açılmıyor.
ComboGetir.Value = ""
ComboGetir.Clear
Say = ThisWorkbook.Worksheets(4).Range("E100000").End(xlUp).Row
If Say < 7 Then
    GoTo GetirBos
End If

'Getir liste değerleri
For i = Say To 7 Step -1
    If ThisWorkbook.Worksheets(4).Range("E" & i).Value <> "" Then
        If ThisWorkbook.Worksheets(4).Range("FR" & i) = "Yes" Then
            With ComboGetir
                .AddItem (ThisWorkbook.Worksheets(4).Range("E" & i).Value)
            End With
        End If
    End If
Next i
GetirBos:

End Sub


Sub ComboGetirResetRapor3_2()
Dim Say As Long, i As Long

On Error Resume Next ' Son Sıra numarası sayısal olmayan karakter içeriyorsa userform açılmıyor.
ComboGetir.Clear
Say = ThisWorkbook.Worksheets(5).Range("E100000").End(xlUp).Row
If Say < 7 Then
    GoTo GetirBos
End If
'Getir liste değerleri
For i = Say To 7 Step -1
    If ThisWorkbook.Worksheets(5).Range("L" & i) = "Point2" Or ThisWorkbook.Worksheets(5).Range("L" & i) = "Point3" Then
        With ComboGetir
            .AddItem (ThisWorkbook.Worksheets(5).Range("E" & i).Value)
        End With
    End If
Next i
GetirBos:


End Sub

Sub ComboGetirResetRapor3_1()
Dim Say As Long, i As Long

On Error Resume Next ' Son Sıra numarası sayısal olmayan karakter içeriyorsa userform açılmıyor.
ComboGetir.Clear
Say = ThisWorkbook.Worksheets(5).Range("E100000").End(xlUp).Row
If Say < 7 Then
    GoTo GetirBos
End If
'Getir liste değerleri
For i = Say To 7 Step -1
    If ThisWorkbook.Worksheets(5).Range("L" & i) = "Point1" Then
        With ComboGetir
            .AddItem (ThisWorkbook.Worksheets(5).Range("E" & i).Value)
        End With
    End If
Next i
GetirBos:

End Sub

Private Sub UserForm_Initialize()
Dim i As Long, WsSKP As Object
Dim ClrLab As MSForms.Control

Dim StrPrinters As Variant
Dim KullanilanYazici As String, prtDogrula As Boolean


ThisWorkbook.Activate


'__________________Kullanılan printeri tespit et ve hazırla

prtDogrula = False
If ThisWorkbook.Worksheets(2).Cells(8, 115).Value <> "" And IsNumeric(ThisWorkbook.Worksheets(2).Cells(8, 115).Value) = True Then
    PrtNo = CInt(ThisWorkbook.Worksheets(2).Cells(8, 115).Value)
    If ThisWorkbook.Worksheets(2).Cells(PrtNo, 115).Value <> "" Then
        KullanilanYazici = ThisWorkbook.Worksheets(2).Cells(PrtNo, 115).Value
        
        '________________________
        
        StrPrinters = ListPrinters
        'Fist check whether the array is filled with anything, by calling another function, IsBounded.
        If IsBounded(StrPrinters) Then
            For i = LBound(StrPrinters) To UBound(StrPrinters)
                If KullanilanYazici = StrPrinters(i) Then
                    prtDogrula = True
                End If
            Next i
        Else
            EtiketPrinterOption.Value = False
            PikurPrinterOption.Value = False
        End If
        
        '________________________

        If prtDogrula = True Then
            If PrtNo = 6 Then
                EtiketPrinterOption.Value = True
            ElseIf PrtNo = 7 Then
                PikurPrinterOption.Value = True
            End If
        Else
            EtiketPrinterOption.Value = False
            PikurPrinterOption.Value = False
        End If
    Else
        EtiketPrinterOption.Value = False
        PikurPrinterOption.Value = False
    End If
Else
    EtiketPrinterOption.Value = False
    PikurPrinterOption.Value = False
End If

'__________________Kullanılan printeri tespit et ve hazırla



For Each ClrLab In core_label_envelope_printing_UI.Controls

    If TypeName(ClrLab) = "Label" Then
        ClrLab.BackColor = RGB(254, 254, 254)
        ClrLab.ForeColor = RGB(30, 30, 30)
    End If
    If TypeName(ClrLab) = "CheckBox" Then
        ClrLab.BackColor = RGB(254, 254, 254) 'RGB(254, 254, 254)
        ClrLab.ForeColor = RGB(30, 30, 30)
    End If
    If TypeName(ClrLab) = "OptionButton" Then
        ClrLab.BackColor = RGB(254, 254, 254)
        ClrLab.ForeColor = RGB(30, 30, 30)
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
UstMenuFrameAlt.BackColor = RGB(225, 235, 245) 'YENİ
AltMenuFrame.BackColor = RGB(225, 235, 245) 'YENİ


Yazdir.BackColor = RGB(225, 235, 245)
Yazdir.ForeColor = RGB(30, 30, 30)
Kapat.BackColor = RGB(225, 235, 245)
Kapat.ForeColor = RGB(30, 30, 30)

PikurTanimla.BackColor = RGB(225, 235, 245)
PikurTanimla.ForeColor = RGB(30, 30, 30)
EtiketTanimla.BackColor = RGB(225, 235, 245)
EtiketTanimla.ForeColor = RGB(30, 30, 30)

EtiketPrinterOption.BackColor = RGB(225, 235, 245)
EtiketPrinterOption.ForeColor = RGB(30, 30, 30)
PikurPrinterOption.BackColor = RGB(225, 235, 245)
PikurPrinterOption.ForeColor = RGB(30, 30, 30)

Rapor1.BackColor = RGB(225, 235, 245)
Rapor1.ForeColor = RGB(30, 30, 30)
Rapor.BackColor = RGB(225, 235, 245)
Rapor.ForeColor = RGB(30, 30, 30)
Rapor2_2.BackColor = RGB(225, 235, 245)
Rapor2_2.ForeColor = RGB(30, 30, 30)
Rapor3_2.BackColor = RGB(225, 235, 245)
Rapor3_2.ForeColor = RGB(30, 30, 30)
Rapor3_1.BackColor = RGB(225, 235, 245)
Rapor3_1.ForeColor = RGB(30, 30, 30)

LblAltMenuButton1.BackColor = RGB(225, 235, 245)
LblAltMenuButton1.ForeColor = RGB(30, 30, 30)
LblAltMenuButton2.BackColor = RGB(225, 235, 245)
LblAltMenuButton2.ForeColor = RGB(30, 30, 30)
LblAltMenuButton3.BackColor = RGB(225, 235, 245)
LblAltMenuButton3.ForeColor = RGB(30, 30, 30)
LblAltMenuButton4.BackColor = RGB(225, 235, 245)
LblAltMenuButton4.ForeColor = RGB(30, 30, 30)
LblAltMenuButton5.BackColor = RGB(225, 235, 245)
LblAltMenuButton5.ForeColor = RGB(30, 30, 30)
LblAltMenuButton6.BackColor = RGB(225, 235, 245)
LblAltMenuButton6.ForeColor = RGB(30, 30, 30)
LblAltMenuButton7.BackColor = RGB(225, 235, 245)
LblAltMenuButton7.ForeColor = RGB(30, 30, 30)


core_label_envelope_printing_UI.BackColor = RGB(230, 230, 230) 'YENİ
ComboGetir.BackColor = RGB(225, 235, 245)

On Error Resume Next ' Son Sıra numarası sayısal olmayan karakter içeriyorsa userform açılmıyor.
ComboGetir.Value = ""
ComboGetir.Clear
On Error GoTo 0


Call EtiketAlanlariReset

BoolKucukZarf = False
BoolBuyukZarf = False
BoolFinansalBirimKucukZarf = False
BoolFinansalBirimBuyukZarf = False


'Açık dropdown kapat
Call ModuleSystemSettings.DropDownKapat


End Sub

Sub EtiketAlanlariReset()

'Etiket alanlarını boşalt
TextKURUM_A.Value = ""
TextSube.Value = ""
TextTeskilatDosyaNoBelgeNo.Value = ""
TextGonderiSekli.Value = ""

TextMuhatapTema1.Value = ""
TextMuhatapTema2.Value = ""
TextMuhatapTema3.Value = ""
TextMuhatapTema4.Value = ""
TextMuhatapTema5.Value = ""
TextMuhatapTema6.Value = ""
TextMuhatapTema7.Value = ""

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
        core_label_envelope_printing_UI.Width = Rep
        yukseklik = yukseklik - 60
        core_label_envelope_printing_UI.Height = yukseklik
        If yukseklik <= 60 Then
            yukseklik = 60
            core_label_envelope_printing_UI.Height = yukseklik
        End If
    ElseIf Rep <= 50 Then
        Rep = 50
        core_label_envelope_printing_UI.Width = Rep
        yukseklik = yukseklik - 50
        core_label_envelope_printing_UI.Height = yukseklik
        If yukseklik <= 50 Then
            yukseklik = 50
            core_label_envelope_printing_UI.Height = yukseklik
        End If
    End If
Loop Until yukseklik = 50

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


Sub Rapor1MuhatapTema()

IlceSakla = Cells(IlkSiraGlobal, 70).Value
If InStr(Cells(IlkSiraGlobal, 70).Value, " Organization A") <> 0 Then
    IlceSakla = ""
End If


'Muhatap
IlBuyukHarf = UCase(Replace(Replace(Cells(IlkSiraGlobal, 69).Value, "i", "I"), "ı", "I"))
IlKucukHarf = Cells(IlkSiraGlobal, 69).Value
If IlceSakla <> "" Then
    IlceBuyukHarf = UCase(Replace(Replace(IlceSakla, "i", "I"), "ı", "I"))
    IlceKucukHarf = IlceSakla
Else
    IlceBuyukHarf = ""
    IlceKucukHarf = ""
End If
If IlceSakla <> "" Then
    Bolum3 = IlceKucukHarf & "/" & IlBuyukHarf
Else
    Bolum3 = IlBuyukHarf
End If

'YENİ MUHATAP TEMASI

M2 = False
M3 = False
M4 = False

'4'lük
If Cells(IlkSiraGlobal, 65).Value <> "" Then
    If Cells(IlkSiraGlobal, 64).Value = "Provincial Directorate B" Or Cells(IlkSiraGlobal, 64).Value = "Provincial Directorate C" Or _
    Cells(IlkSiraGlobal, 64).Value = "Provincial Directorate D" Or Cells(IlkSiraGlobal, 64).Value = "Provincial Directorate E" Then 'VALİLİK
        TextMuhatapTema1.Value = IlBuyukHarf & " PROVINCIAL GOVERNORSHIP"
        TextMuhatapTema2.Value = Cells(IlkSiraGlobal, 64).Value  'UCase(Replace(Replace(Cells(IlkSiraGlobal, 64).Value, "i", "I"), "ı", "I"))
        TextMuhatapTema3.Value = "(" & Cells(IlkSiraGlobal, 65).Value & ")"
        TextMuhatapTema4.Value = Bolum3
        M4 = True
    ElseIf Cells(IlkSiraGlobal, 64).Value = "District Directorate B" Or Cells(IlkSiraGlobal, 64).Value = "District Directorate C" Or _
    Cells(IlkSiraGlobal, 64).Value = "District Directorate D" Or Cells(IlkSiraGlobal, 64).Value = "District Directorate E" Then 'KAYMAKAMLIK
        TextMuhatapTema1.Value = IlceBuyukHarf & " DISTRICT GOVERNORSHIP"
        TextMuhatapTema2.Value = Cells(IlkSiraGlobal, 64).Value 'UCase(Replace(Replace(Cells(IlkSiraGlobal, 64).Value, "i", "I"), "ı", "I"))
        TextMuhatapTema3.Value = "(" & Cells(IlkSiraGlobal, 65).Value & ")"
        TextMuhatapTema4.Value = Bolum3
        M4 = True
    ElseIf InStr(Cells(IlkSiraGlobal, 64).Value, "General Directorate") <> 0 Or InStr(Cells(IlkSiraGlobal, 64).Value, "Regional Directorate") <> 0 Then
        TextMuhatapTema1.Value = UCase(Replace(Replace(ThisWorkbook.Worksheets(2).Cells(6, 111).Value, "i", "I"), "ı", "I"))
        TextMuhatapTema2.Value = Cells(IlkSiraGlobal, 64).Value 'UCase(Replace(Replace(Cells(IlkSiraGlobal, 64).Value, "i", "I"), "ı", "I"))
        TextMuhatapTema3.Value = "(" & Cells(IlkSiraGlobal, 65).Value & ")"
        TextMuhatapTema4.Value = Bolum3
        M4 = True
    Else 'YARGI 3'lük
        Bolum1 = UCase(Replace(Replace(Cells(IlkSiraGlobal, 64).Value, "i", "I"), "ı", "I"))
        If InStr(Bolum1, "X.X. ") > 0 Then
            TextMuhatapTema1.Value = Mid(Bolum1, 6, Len(Bolum1))
            TextMuhatapTema2.Value = "(" & Cells(IlkSiraGlobal, 65).Value & ")"
            TextMuhatapTema3.Value = Bolum3
            M3 = True
        Else
            TextMuhatapTema1.Value = Bolum1
            TextMuhatapTema2.Value = "(" & Cells(IlkSiraGlobal, 65).Value & ")"
            TextMuhatapTema3.Value = Bolum3
            M3 = True
        End If
    End If
End If
'3'lük
If Cells(IlkSiraGlobal, 65).Value = "" Then
    If Cells(IlkSiraGlobal, 64).Value = "Provincial Directorate B" Or Cells(IlkSiraGlobal, 64).Value = "Provincial Directorate C" Or _
    Cells(IlkSiraGlobal, 64).Value = "Provincial Directorate D" Or Cells(IlkSiraGlobal, 64).Value = "Provincial Directorate E" Then 'VALİLİK
        TextMuhatapTema1.Value = IlBuyukHarf & " PROVINCIAL GOVERNORSHIP"
        TextMuhatapTema2.Value = "(" & Cells(IlkSiraGlobal, 64).Value & ")" 'UCase(Replace(Replace(Cells(IlkSiraGlobal, 64).Value, "i", "I"), "ı", "I"))
        TextMuhatapTema3.Value = Bolum3
        M3 = True
    ElseIf Cells(IlkSiraGlobal, 64).Value = "District Directorate B" Or Cells(IlkSiraGlobal, 64).Value = "District Directorate C" Or _
    Cells(IlkSiraGlobal, 64).Value = "District Directorate D" Or Cells(IlkSiraGlobal, 64).Value = "District Directorate E" Then
        TextMuhatapTema1.Value = IlceBuyukHarf & " DISTRICT GOVERNORSHIP"
        TextMuhatapTema2.Value = "(" & Cells(IlkSiraGlobal, 64).Value & ")" 'UCase(Replace(Replace(Cells(IlkSiraGlobal, 64).Value, "i", "I"), "ı", "I"))
        TextMuhatapTema3.Value = Bolum3
        M3 = True
    ElseIf InStr(Cells(IlkSiraGlobal, 64).Value, "General Directorate") <> 0 Or InStr(Cells(IlkSiraGlobal, 64).Value, "Regional Directorate") <> 0 Then
        TextMuhatapTema1.Value = UCase(Replace(Replace(ThisWorkbook.Worksheets(2).Cells(6, 111).Value, "i", "I"), "ı", "I"))
        TextMuhatapTema2.Value = "(" & Cells(IlkSiraGlobal, 64).Value & ")"  'UCase(Replace(Replace(Cells(IlkSiraGlobal, 64).Value, "i", "I"), "ı", "I"))
        TextMuhatapTema3.Value = Bolum3
        M3 = True
    Else 'YARGI 2'lik
        Bolum1 = UCase(Replace(Replace(Cells(IlkSiraGlobal, 64).Value, "i", "I"), "ı", "I"))
        If InStr(Bolum1, "X.X. ") > 0 Then
            TextMuhatapTema1.Value = Mid(Bolum1, 6, Len(Bolum1))
            TextMuhatapTema2.Value = Bolum3
            M2 = True
        Else
            TextMuhatapTema1.Value = Bolum1
            TextMuhatapTema2.Value = Bolum3
            M2 = True
        End If
    End If
End If

End Sub

Sub RaporMuhatapTema()

IlceSakla = Cells(IlkSiraGlobal, 78).Value
If InStr(Cells(IlkSiraGlobal, 78).Value, " Organization A") <> 0 Then
    IlceSakla = ""
End If


'Muhatap
IlBuyukHarf = UCase(Replace(Replace(Cells(IlkSiraGlobal, 77).Value, "i", "I"), "ı", "I"))
IlKucukHarf = Cells(IlkSiraGlobal, 77).Value
If IlceSakla <> "" Then
    IlceBuyukHarf = UCase(Replace(Replace(IlceSakla, "i", "I"), "ı", "I"))
    IlceKucukHarf = IlceSakla
Else
    IlceBuyukHarf = ""
    IlceKucukHarf = ""
End If
If IlceSakla <> "" Then
    Bolum3 = IlceKucukHarf & "/" & IlBuyukHarf
Else
    Bolum3 = IlBuyukHarf
End If

'YENİ MUHATAP TEMASI

M2 = False
M3 = False
M4 = False

'4'lük
If Cells(IlkSiraGlobal, 73).Value <> "" Then
    If Cells(IlkSiraGlobal, 72).Value = "Provincial Directorate B" Or Cells(IlkSiraGlobal, 72).Value = "Provincial Directorate C" Or _
    Cells(IlkSiraGlobal, 72).Value = "Provincial Directorate D" Or Cells(IlkSiraGlobal, 72).Value = "Provincial Directorate E" Then 'VALİLİK
        TextMuhatapTema1.Value = IlBuyukHarf & " PROVINCIAL GOVERNORSHIP"
        TextMuhatapTema2.Value = Cells(IlkSiraGlobal, 72).Value  'UCase(Replace(Replace(Cells(IlkSiraGlobal, 72).Value, "i", "I"), "ı", "I"))
        TextMuhatapTema3.Value = "(" & Cells(IlkSiraGlobal, 73).Value & ")"
        TextMuhatapTema4.Value = Bolum3
        M4 = True
    ElseIf Cells(IlkSiraGlobal, 72).Value = "District Directorate B" Or Cells(IlkSiraGlobal, 72).Value = "District Directorate C" Or _
    Cells(IlkSiraGlobal, 72).Value = "District Directorate D" Or Cells(IlkSiraGlobal, 72).Value = "District Directorate E" Then 'KAYMAKAMLIK
        TextMuhatapTema1.Value = IlceBuyukHarf & " DISTRICT GOVERNORSHIP"
        TextMuhatapTema2.Value = Cells(IlkSiraGlobal, 72).Value 'UCase(Replace(Replace(Cells(IlkSiraGlobal, 72).Value, "i", "I"), "ı", "I"))
        TextMuhatapTema3.Value = "(" & Cells(IlkSiraGlobal, 73).Value & ")"
        TextMuhatapTema4.Value = Bolum3
        M4 = True
    ElseIf InStr(Cells(IlkSiraGlobal, 72).Value, "General Directorate") <> 0 Or InStr(Cells(IlkSiraGlobal, 72).Value, "Regional Directorate") <> 0 Then
        TextMuhatapTema1.Value = UCase(Replace(Replace(ThisWorkbook.Worksheets(2).Cells(6, 111).Value, "i", "I"), "ı", "I"))
        TextMuhatapTema2.Value = Cells(IlkSiraGlobal, 72).Value 'UCase(Replace(Replace(Cells(IlkSiraGlobal, 72).Value, "i", "I"), "ı", "I"))
        TextMuhatapTema3.Value = "(" & Cells(IlkSiraGlobal, 73).Value & ")"
        TextMuhatapTema4.Value = Bolum3
        M4 = True
    Else 'YARGI 3'lük
        Bolum1 = UCase(Replace(Replace(Cells(IlkSiraGlobal, 72).Value, "i", "I"), "ı", "I"))
        If InStr(Bolum1, "X.X. ") > 0 Then
            TextMuhatapTema1.Value = Mid(Bolum1, 6, Len(Bolum1))
            TextMuhatapTema2.Value = "(" & Cells(IlkSiraGlobal, 73).Value & ")"
            TextMuhatapTema3.Value = Bolum3
            M3 = True
        Else
            TextMuhatapTema1.Value = Bolum1
            TextMuhatapTema2.Value = "(" & Cells(IlkSiraGlobal, 73).Value & ")"
            TextMuhatapTema3.Value = Bolum3
            M3 = True
        End If
    End If
End If
'3'lük
If Cells(IlkSiraGlobal, 73).Value = "" Then
    If Cells(IlkSiraGlobal, 72).Value = "Provincial Directorate B" Or Cells(IlkSiraGlobal, 72).Value = "Provincial Directorate C" Or _
    Cells(IlkSiraGlobal, 72).Value = "Provincial Directorate D" Or Cells(IlkSiraGlobal, 72).Value = "Provincial Directorate E" Then 'VALİLİK
        TextMuhatapTema1.Value = IlBuyukHarf & " PROVINCIAL GOVERNORSHIP"
        TextMuhatapTema2.Value = "(" & Cells(IlkSiraGlobal, 72).Value & ")" 'UCase(Replace(Replace(Cells(IlkSiraGlobal, 72).Value, "i", "I"), "ı", "I"))
        TextMuhatapTema3.Value = Bolum3
        M3 = True
    ElseIf Cells(IlkSiraGlobal, 72).Value = "District Directorate B" Or Cells(IlkSiraGlobal, 72).Value = "District Directorate C" Or _
    Cells(IlkSiraGlobal, 72).Value = "District Directorate D" Or Cells(IlkSiraGlobal, 72).Value = "District Directorate E" Then
        TextMuhatapTema1.Value = IlceBuyukHarf & " DISTRICT GOVERNORSHIP"
        TextMuhatapTema2.Value = "(" & Cells(IlkSiraGlobal, 72).Value & ")" 'UCase(Replace(Replace(Cells(IlkSiraGlobal, 72).Value, "i", "I"), "ı", "I"))
        TextMuhatapTema3.Value = Bolum3
        M3 = True
    ElseIf InStr(Cells(IlkSiraGlobal, 72).Value, "General Directorate") <> 0 Or InStr(Cells(IlkSiraGlobal, 72).Value, "Regional Directorate") <> 0 Then
        TextMuhatapTema1.Value = UCase(Replace(Replace(ThisWorkbook.Worksheets(2).Cells(6, 111).Value, "i", "I"), "ı", "I"))
        TextMuhatapTema2.Value = "(" & Cells(IlkSiraGlobal, 72).Value & ")"  'UCase(Replace(Replace(Cells(IlkSiraGlobal, 72).Value, "i", "I"), "ı", "I"))
        TextMuhatapTema3.Value = Bolum3
        M3 = True
    Else 'YARGI 2'lik
        Bolum1 = UCase(Replace(Replace(Cells(IlkSiraGlobal, 72).Value, "i", "I"), "ı", "I"))
        If InStr(Bolum1, "X.X. ") > 0 Then
            TextMuhatapTema1.Value = Mid(Bolum1, 6, Len(Bolum1))
            TextMuhatapTema2.Value = Bolum3
            M2 = True
        Else
            TextMuhatapTema1.Value = Bolum1
            TextMuhatapTema2.Value = Bolum3
            M2 = True
        End If
    End If
End If

End Sub

Sub BilgilendirmeMuhatapTema()

IlceSakla = Cells(IlkSiraGlobal, 204).Value
If InStr(Cells(IlkSiraGlobal, 204).Value, " Organization A") <> 0 Then
    IlceSakla = ""
End If


'Muhatap
IlBuyukHarf = UCase(Replace(Replace(Cells(IlkSiraGlobal, 203).Value, "i", "I"), "ı", "I"))
IlKucukHarf = Cells(IlkSiraGlobal, 203).Value
If IlceSakla <> "" Then
    IlceBuyukHarf = UCase(Replace(Replace(IlceSakla, "i", "I"), "ı", "I"))
    IlceKucukHarf = IlceSakla
Else
    IlceBuyukHarf = ""
    IlceKucukHarf = ""
End If
If IlceSakla <> "" Then
    Bolum3 = IlceKucukHarf & "/" & IlBuyukHarf
Else
    Bolum3 = IlBuyukHarf
End If

'YENİ MUHATAP TEMASI

M2 = False
M3 = False
M4 = False

'4'lük
If Cells(IlkSiraGlobal, 200).Value <> "" Then
    If Cells(IlkSiraGlobal, 199).Value = "Provincial Directorate B" Or Cells(IlkSiraGlobal, 199).Value = "Provincial Directorate C" Or _
    Cells(IlkSiraGlobal, 199).Value = "Provincial Directorate D" Or Cells(IlkSiraGlobal, 199).Value = "Provincial Directorate E" Then 'VALİLİK
        TextMuhatapTema1.Value = IlBuyukHarf & " PROVINCIAL GOVERNORSHIP"
        TextMuhatapTema2.Value = Cells(IlkSiraGlobal, 199).Value  'UCase(Replace(Replace(Cells(IlkSiraGlobal, 199).Value, "i", "I"), "ı", "I"))
        TextMuhatapTema3.Value = "(" & Cells(IlkSiraGlobal, 200).Value & ")"
        TextMuhatapTema4.Value = Bolum3
        M4 = True
    ElseIf Cells(IlkSiraGlobal, 199).Value = "District Directorate B" Or Cells(IlkSiraGlobal, 199).Value = "District Directorate C" Or _
    Cells(IlkSiraGlobal, 199).Value = "District Directorate D" Or Cells(IlkSiraGlobal, 199).Value = "District Directorate E" Then 'KAYMAKAMLIK
        TextMuhatapTema1.Value = IlceBuyukHarf & " DISTRICT GOVERNORSHIP"
        TextMuhatapTema2.Value = Cells(IlkSiraGlobal, 199).Value 'UCase(Replace(Replace(Cells(IlkSiraGlobal, 199).Value, "i", "I"), "ı", "I"))
        TextMuhatapTema3.Value = "(" & Cells(IlkSiraGlobal, 200).Value & ")"
        TextMuhatapTema4.Value = Bolum3
        M4 = True
    ElseIf InStr(Cells(IlkSiraGlobal, 199).Value, "General Directorate") <> 0 Or InStr(Cells(IlkSiraGlobal, 199).Value, "Regional Directorate") <> 0 Then
        TextMuhatapTema1.Value = UCase(Replace(Replace(ThisWorkbook.Worksheets(2).Cells(6, 111).Value, "i", "I"), "ı", "I"))
        TextMuhatapTema2.Value = Cells(IlkSiraGlobal, 199).Value 'UCase(Replace(Replace(Cells(IlkSiraGlobal, 199).Value, "i", "I"), "ı", "I"))
        TextMuhatapTema3.Value = "(" & Cells(IlkSiraGlobal, 200).Value & ")"
        TextMuhatapTema4.Value = Bolum3
        M4 = True
    Else 'YARGI 3'lük
        Bolum1 = UCase(Replace(Replace(Cells(IlkSiraGlobal, 199).Value, "i", "I"), "ı", "I"))
        If InStr(Bolum1, "X.X. ") > 0 Then
            TextMuhatapTema1.Value = Mid(Bolum1, 6, Len(Bolum1))
            TextMuhatapTema2.Value = "(" & Cells(IlkSiraGlobal, 200).Value & ")"
            TextMuhatapTema3.Value = Bolum3
            M3 = True
        Else
            TextMuhatapTema1.Value = Bolum1
            TextMuhatapTema2.Value = "(" & Cells(IlkSiraGlobal, 200).Value & ")"
            TextMuhatapTema3.Value = Bolum3
            M3 = True
        End If
    End If
End If
'3'lük
If Cells(IlkSiraGlobal, 200).Value = "" Then
    If Cells(IlkSiraGlobal, 199).Value = "Provincial Directorate B" Or Cells(IlkSiraGlobal, 199).Value = "Provincial Directorate C" Or _
    Cells(IlkSiraGlobal, 199).Value = "Provincial Directorate D" Or Cells(IlkSiraGlobal, 199).Value = "Provincial Directorate E" Then 'VALİLİK
        TextMuhatapTema1.Value = IlBuyukHarf & " PROVINCIAL GOVERNORSHIP"
        TextMuhatapTema2.Value = "(" & Cells(IlkSiraGlobal, 199).Value & ")" 'UCase(Replace(Replace(Cells(IlkSiraGlobal, 199).Value, "i", "I"), "ı", "I"))
        TextMuhatapTema3.Value = Bolum3
        M3 = True
    ElseIf Cells(IlkSiraGlobal, 199).Value = "District Directorate B" Or Cells(IlkSiraGlobal, 199).Value = "District Directorate C" Or _
    Cells(IlkSiraGlobal, 199).Value = "District Directorate D" Or Cells(IlkSiraGlobal, 199).Value = "District Directorate E" Then
        TextMuhatapTema1.Value = IlceBuyukHarf & " DISTRICT GOVERNORSHIP"
        TextMuhatapTema2.Value = "(" & Cells(IlkSiraGlobal, 199).Value & ")" 'UCase(Replace(Replace(Cells(IlkSiraGlobal, 199).Value, "i", "I"), "ı", "I"))
        TextMuhatapTema3.Value = Bolum3
        M3 = True
    ElseIf InStr(Cells(IlkSiraGlobal, 199).Value, "General Directorate") <> 0 Or InStr(Cells(IlkSiraGlobal, 199).Value, "Regional Directorate") <> 0 Then
        TextMuhatapTema1.Value = UCase(Replace(Replace(ThisWorkbook.Worksheets(2).Cells(6, 111).Value, "i", "I"), "ı", "I"))
        TextMuhatapTema2.Value = "(" & Cells(IlkSiraGlobal, 199).Value & ")"  'UCase(Replace(Replace(Cells(IlkSiraGlobal, 199).Value, "i", "I"), "ı", "I"))
        TextMuhatapTema3.Value = Bolum3
        M3 = True
    Else 'YARGI 2'lik
        Bolum1 = UCase(Replace(Replace(Cells(IlkSiraGlobal, 199).Value, "i", "I"), "ı", "I"))
        If InStr(Bolum1, "X.X. ") > 0 Then
            TextMuhatapTema1.Value = Mid(Bolum1, 6, Len(Bolum1))
            TextMuhatapTema2.Value = Bolum3
            M2 = True
        Else
            TextMuhatapTema1.Value = Bolum1
            TextMuhatapTema2.Value = Bolum3
            M2 = True
        End If
    End If
End If

End Sub



Sub Rapor3_1MuhatapTema()

IlceSakla = Cells(IlkSiraGlobal, 92).Value
If InStr(Cells(IlkSiraGlobal, 92).Value, " Organization A") <> 0 Then
    IlceSakla = ""
End If


'Muhatap
IlBuyukHarf = UCase(Replace(Replace(Cells(IlkSiraGlobal, 91).Value, "i", "I"), "ı", "I"))
IlKucukHarf = Cells(IlkSiraGlobal, 91).Value
If IlceSakla <> "" Then
    IlceBuyukHarf = UCase(Replace(Replace(IlceSakla, "i", "I"), "ı", "I"))
    IlceKucukHarf = IlceSakla
Else
    IlceBuyukHarf = ""
    IlceKucukHarf = ""
End If
If IlceSakla <> "" Then
    Bolum3 = IlceKucukHarf & "/" & IlBuyukHarf
Else
    Bolum3 = IlBuyukHarf
End If

'YENİ MUHATAP TEMASI

M2 = False
M3 = False
M4 = False

'4'lük
If Cells(IlkSiraGlobal, 103).Value <> "" Then
    If Cells(IlkSiraGlobal, 102).Value = "Provincial Directorate B" Or Cells(IlkSiraGlobal, 102).Value = "Provincial Directorate C" Or _
    Cells(IlkSiraGlobal, 102).Value = "Provincial Directorate D" Or Cells(IlkSiraGlobal, 102).Value = "Provincial Directorate E" Then 'VALİLİK
        TextMuhatapTema1.Value = IlBuyukHarf & " PROVINCIAL GOVERNORSHIP"
        TextMuhatapTema2.Value = Cells(IlkSiraGlobal, 102).Value  'UCase(Replace(Replace(Cells(IlkSiraGlobal, 102).Value, "i", "I"), "ı", "I"))
        TextMuhatapTema3.Value = "(" & Cells(IlkSiraGlobal, 103).Value & ")"
        TextMuhatapTema4.Value = Bolum3
        M4 = True
    ElseIf Cells(IlkSiraGlobal, 102).Value = "District Directorate B" Or Cells(IlkSiraGlobal, 102).Value = "District Directorate C" Or _
    Cells(IlkSiraGlobal, 102).Value = "District Directorate D" Or Cells(IlkSiraGlobal, 102).Value = "District Directorate E" Then 'KAYMAKAMLIK
        TextMuhatapTema1.Value = IlceBuyukHarf & " DISTRICT GOVERNORSHIP"
        TextMuhatapTema2.Value = Cells(IlkSiraGlobal, 102).Value 'UCase(Replace(Replace(Cells(IlkSiraGlobal, 102).Value, "i", "I"), "ı", "I"))
        TextMuhatapTema3.Value = "(" & Cells(IlkSiraGlobal, 103).Value & ")"
        TextMuhatapTema4.Value = Bolum3
        M4 = True
    ElseIf InStr(Cells(IlkSiraGlobal, 102).Value, "General Directorate") <> 0 Or InStr(Cells(IlkSiraGlobal, 102).Value, "Regional Directorate") <> 0 Then
        TextMuhatapTema1.Value = UCase(Replace(Replace(ThisWorkbook.Worksheets(2).Cells(6, 111).Value, "i", "I"), "ı", "I"))
        TextMuhatapTema2.Value = Cells(IlkSiraGlobal, 102).Value 'UCase(Replace(Replace(Cells(IlkSiraGlobal, 102).Value, "i", "I"), "ı", "I"))
        TextMuhatapTema3.Value = "(" & Cells(IlkSiraGlobal, 103).Value & ")"
        TextMuhatapTema4.Value = Bolum3
        M4 = True
    Else 'YARGI 3'lük
        Bolum1 = UCase(Replace(Replace(Cells(IlkSiraGlobal, 102).Value, "i", "I"), "ı", "I"))
        If InStr(Bolum1, "X.X. ") > 0 Then
            TextMuhatapTema1.Value = Mid(Bolum1, 6, Len(Bolum1))
            TextMuhatapTema2.Value = "(" & Cells(IlkSiraGlobal, 103).Value & ")"
            TextMuhatapTema3.Value = Bolum3
            M3 = True
        Else
            TextMuhatapTema1.Value = Bolum1
            TextMuhatapTema2.Value = "(" & Cells(IlkSiraGlobal, 103).Value & ")"
            TextMuhatapTema3.Value = Bolum3
            M3 = True
        End If
    End If
End If
'3'lük
If Cells(IlkSiraGlobal, 103).Value = "" Then
    If Cells(IlkSiraGlobal, 102).Value = "Provincial Directorate B" Or Cells(IlkSiraGlobal, 102).Value = "Provincial Directorate C" Or _
    Cells(IlkSiraGlobal, 102).Value = "Provincial Directorate D" Or Cells(IlkSiraGlobal, 102).Value = "Provincial Directorate E" Then 'VALİLİK
        TextMuhatapTema1.Value = IlBuyukHarf & " PROVINCIAL GOVERNORSHIP"
        TextMuhatapTema2.Value = "(" & Cells(IlkSiraGlobal, 102).Value & ")" 'UCase(Replace(Replace(Cells(IlkSiraGlobal, 102).Value, "i", "I"), "ı", "I"))
        TextMuhatapTema3.Value = Bolum3
        M3 = True
    ElseIf Cells(IlkSiraGlobal, 102).Value = "District Directorate B" Or Cells(IlkSiraGlobal, 102).Value = "District Directorate C" Or _
    Cells(IlkSiraGlobal, 102).Value = "District Directorate D" Or Cells(IlkSiraGlobal, 102).Value = "District Directorate E" Then
        TextMuhatapTema1.Value = IlceBuyukHarf & " DISTRICT GOVERNORSHIP"
        TextMuhatapTema2.Value = "(" & Cells(IlkSiraGlobal, 102).Value & ")" 'UCase(Replace(Replace(Cells(IlkSiraGlobal, 102).Value, "i", "I"), "ı", "I"))
        TextMuhatapTema3.Value = Bolum3
        M3 = True
    ElseIf InStr(Cells(IlkSiraGlobal, 102).Value, "General Directorate") <> 0 Or InStr(Cells(IlkSiraGlobal, 102).Value, "Regional Directorate") <> 0 Then
        TextMuhatapTema1.Value = UCase(Replace(Replace(ThisWorkbook.Worksheets(2).Cells(6, 111).Value, "i", "I"), "ı", "I"))
        TextMuhatapTema2.Value = "(" & Cells(IlkSiraGlobal, 102).Value & ")"  'UCase(Replace(Replace(Cells(IlkSiraGlobal, 102).Value, "i", "I"), "ı", "I"))
        TextMuhatapTema3.Value = Bolum3
        M3 = True
    Else 'YARGI 2'lik
        Bolum1 = UCase(Replace(Replace(Cells(IlkSiraGlobal, 102).Value, "i", "I"), "ı", "I"))
        If InStr(Bolum1, "X.X. ") > 0 Then
            TextMuhatapTema1.Value = Mid(Bolum1, 6, Len(Bolum1))
            TextMuhatapTema2.Value = Bolum3
            M2 = True
        Else
            TextMuhatapTema1.Value = Bolum1
            TextMuhatapTema2.Value = Bolum3
            M2 = True
        End If
    End If
End If

End Sub

Sub Rapor3_2MuhatapTema()

IlceSakla = Cells(IlkSiraGlobal, 20).Value
If InStr(Cells(IlkSiraGlobal, 20).Value, " Organization A") <> 0 Then
    IlceSakla = ""
End If


'Muhatap
IlBuyukHarf = UCase(Replace(Replace(Cells(IlkSiraGlobal, 19).Value, "i", "I"), "ı", "I"))
IlKucukHarf = Cells(IlkSiraGlobal, 19).Value
If IlceSakla <> "" Then
    IlceBuyukHarf = UCase(Replace(Replace(IlceSakla, "i", "I"), "ı", "I"))
    IlceKucukHarf = IlceSakla
Else
    IlceBuyukHarf = ""
    IlceKucukHarf = ""
End If
If IlceSakla <> "" Then
    Bolum3 = IlceKucukHarf & "/" & IlBuyukHarf
Else
    Bolum3 = IlBuyukHarf
End If

'YENİ MUHATAP TEMASI

M2 = False
M3 = False
M4 = False

'4'lük
If Cells(IlkSiraGlobal, 48).Value <> "" Then
    If Cells(IlkSiraGlobal, 47).Value = "Provincial Directorate B" Or Cells(IlkSiraGlobal, 47).Value = "Provincial Directorate C" Or _
    Cells(IlkSiraGlobal, 47).Value = "Provincial Directorate D" Or Cells(IlkSiraGlobal, 47).Value = "Provincial Directorate E" Then 'VALİLİK
        TextMuhatapTema1.Value = IlBuyukHarf & " PROVINCIAL GOVERNORSHIP"
        TextMuhatapTema2.Value = Cells(IlkSiraGlobal, 47).Value  'UCase(Replace(Replace(Cells(IlkSiraGlobal, 47).Value, "i", "I"), "ı", "I"))
        TextMuhatapTema3.Value = "(" & Cells(IlkSiraGlobal, 48).Value & ")"
        TextMuhatapTema4.Value = Bolum3
        M4 = True
    ElseIf Cells(IlkSiraGlobal, 47).Value = "District Directorate B" Or Cells(IlkSiraGlobal, 47).Value = "District Directorate C" Or _
    Cells(IlkSiraGlobal, 47).Value = "District Directorate D" Or Cells(IlkSiraGlobal, 47).Value = "District Directorate E" Then 'KAYMAKAMLIK
        TextMuhatapTema1.Value = IlceBuyukHarf & " DISTRICT GOVERNORSHIP"
        TextMuhatapTema2.Value = Cells(IlkSiraGlobal, 47).Value 'UCase(Replace(Replace(Cells(IlkSiraGlobal, 47).Value, "i", "I"), "ı", "I"))
        TextMuhatapTema3.Value = "(" & Cells(IlkSiraGlobal, 48).Value & ")"
        TextMuhatapTema4.Value = Bolum3
        M4 = True
    ElseIf InStr(Cells(IlkSiraGlobal, 47).Value, "General Directorate") <> 0 Or InStr(Cells(IlkSiraGlobal, 47).Value, "Regional Directorate") <> 0 Then
        TextMuhatapTema1.Value = UCase(Replace(Replace(ThisWorkbook.Worksheets(2).Cells(6, 111).Value, "i", "I"), "ı", "I"))
        TextMuhatapTema2.Value = Cells(IlkSiraGlobal, 47).Value 'UCase(Replace(Replace(Cells(IlkSiraGlobal, 47).Value, "i", "I"), "ı", "I"))
        TextMuhatapTema3.Value = "(" & Cells(IlkSiraGlobal, 48).Value & ")"
        TextMuhatapTema4.Value = Bolum3
        M4 = True
    Else 'YARGI 3'lük
        Bolum1 = UCase(Replace(Replace(Cells(IlkSiraGlobal, 47).Value, "i", "I"), "ı", "I"))
        If InStr(Bolum1, "X.X. ") > 0 Then
            TextMuhatapTema1.Value = Mid(Bolum1, 6, Len(Bolum1))
            TextMuhatapTema2.Value = "(" & Cells(IlkSiraGlobal, 48).Value & ")"
            TextMuhatapTema3.Value = Bolum3
            M3 = True
        Else
            TextMuhatapTema1.Value = Bolum1
            TextMuhatapTema2.Value = "(" & Cells(IlkSiraGlobal, 48).Value & ")"
            TextMuhatapTema3.Value = Bolum3
            M3 = True
        End If
    End If
End If
'3'lük
If Cells(IlkSiraGlobal, 48).Value = "" Then
    If Cells(IlkSiraGlobal, 47).Value = "Provincial Directorate B" Or Cells(IlkSiraGlobal, 47).Value = "Provincial Directorate C" Or _
    Cells(IlkSiraGlobal, 47).Value = "Provincial Directorate D" Or Cells(IlkSiraGlobal, 47).Value = "Provincial Directorate E" Then 'VALİLİK
        TextMuhatapTema1.Value = IlBuyukHarf & " PROVINCIAL GOVERNORSHIP"
        TextMuhatapTema2.Value = "(" & Cells(IlkSiraGlobal, 47).Value & ")" 'UCase(Replace(Replace(Cells(IlkSiraGlobal, 47).Value, "i", "I"), "ı", "I"))
        TextMuhatapTema3.Value = Bolum3
        M3 = True
    ElseIf Cells(IlkSiraGlobal, 47).Value = "District Directorate B" Or Cells(IlkSiraGlobal, 47).Value = "District Directorate C" Or _
    Cells(IlkSiraGlobal, 47).Value = "District Directorate D" Or Cells(IlkSiraGlobal, 47).Value = "District Directorate E" Then
        TextMuhatapTema1.Value = IlceBuyukHarf & " DISTRICT GOVERNORSHIP"
        TextMuhatapTema2.Value = "(" & Cells(IlkSiraGlobal, 47).Value & ")" 'UCase(Replace(Replace(Cells(IlkSiraGlobal, 47).Value, "i", "I"), "ı", "I"))
        TextMuhatapTema3.Value = Bolum3
        M3 = True
    ElseIf InStr(Cells(IlkSiraGlobal, 47).Value, "General Directorate") <> 0 Or InStr(Cells(IlkSiraGlobal, 47).Value, "Regional Directorate") <> 0 Then
        TextMuhatapTema1.Value = UCase(Replace(Replace(ThisWorkbook.Worksheets(2).Cells(6, 111).Value, "i", "I"), "ı", "I"))
        TextMuhatapTema2.Value = "(" & Cells(IlkSiraGlobal, 47).Value & ")"  'UCase(Replace(Replace(Cells(IlkSiraGlobal, 47).Value, "i", "I"), "ı", "I"))
        TextMuhatapTema3.Value = Bolum3
        M3 = True
    Else 'YARGI 2'lik
        Bolum1 = UCase(Replace(Replace(Cells(IlkSiraGlobal, 47).Value, "i", "I"), "ı", "I"))
        If InStr(Bolum1, "X.X. ") > 0 Then
            TextMuhatapTema1.Value = Mid(Bolum1, 6, Len(Bolum1))
            TextMuhatapTema2.Value = Bolum3
            M2 = True
        Else
            TextMuhatapTema1.Value = Bolum1
            TextMuhatapTema2.Value = Bolum3
            M2 = True
        End If
    End If
End If

End Sub

Sub FinansalBirimMuhatapTema()


IlceSakla = Cells(IlkSiraGlobal, 78).Value
If InStr(Cells(IlkSiraGlobal, 78).Value, " Organization A") <> 0 Then
    IlceSakla = ""
End If


'YENİ MUHATAP TEMASI
M2 = False
M3 = False
M4 = False

'Muhatap
If Cells(IlkSiraGlobal, 82).Value <> "" Then
    TextMuhatapTema1.Value = UCase(Replace(Replace(Cells(IlkSiraGlobal, 30).Value, "i", "I"), "ı", "I")) 'FinansalBirim
    TextMuhatapTema2.Value = "(" & Cells(IlkSiraGlobal, 82).Value & ")" 'Birim
    TextMuhatapTema3.Value = Cells(IlkSiraGlobal, 79).Value 'Adres
    If IlceSakla <> "" Then
        TextMuhatapTema4.Value = IlceSakla & "/" & UCase(Replace(Replace(Cells(IlkSiraGlobal, 77).Value, "i", "I"), "ı", "I"))
    Else
        TextMuhatapTema4.Value = UCase(Replace(Replace(Cells(IlkSiraGlobal, 77).Value, "i", "I"), "ı", "I"))
    End If
    M4 = True
Else
    TextMuhatapTema1.Value = UCase(Replace(Replace(Cells(IlkSiraGlobal, 30).Value, "i", "I"), "ı", "I")) 'FinansalBirim
    TextMuhatapTema2.Value = Cells(IlkSiraGlobal, 79).Value 'Adres
    If IlceSakla <> "" Then
        TextMuhatapTema3.Value = IlceSakla & "/" & UCase(Replace(Replace(Cells(IlkSiraGlobal, 77).Value, "i", "I"), "ı", "I"))
    Else
        TextMuhatapTema3.Value = UCase(Replace(Replace(Cells(IlkSiraGlobal, 77).Value, "i", "I"), "ı", "I"))
    End If
    M3 = True
End If


End Sub

