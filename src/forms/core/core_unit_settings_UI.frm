VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} core_unit_settings_UI 
   Caption         =   "Enterprise Document Automation System"
   ClientHeight    =   10890
   ClientLeft      =   30
   ClientTop       =   375
   ClientWidth     =   13920
   OleObjectBlob   =   "core_unit_settings_UI.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "core_unit_settings_UI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim Abort As Boolean
Public OpenWordTakip As Boolean


Private Sub ComboBirim_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

Me.ComboBirim.DropDown

End Sub

Private Sub ComboBirim_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next
    
    Select Case KeyCode
        Case 38  'Up
            If ComboBirim.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                ComboBirim.ListIndex = ComboBirim.ListIndex
            End If
        Case 40 'Down
            If ComboBirim.ListIndex >= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                ComboBirim.ListIndex = ComboBirim.ListIndex
            End If
    End Select
    Abort = False
    
End Sub

Private Sub ComboBirim_Change()
Dim a() As Variant, i As Variant
Dim AutoPath As String, DestTaslak As String, OpenControl As String
Dim TaslakFile As String, SayHedef As Long, ItemName As String, j As Integer
Dim fso As Object, objWord As Object, objDoc As Object, Kurum_ANoStr As String

On Error GoTo Atla
'Birimi çağır
If Left(ComboBirim.Value, Len(ComboBirim.Value) - 5) = Worksheets(2).Cells(6, 99).Value And ComboBirim.Value <> "" Then
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    AutoPath = ThisWorkbook.Path
    DestTaslak = AutoPath & "\System Files\System Templates\Footer Field\"
    TaslakFile = "External Footer.docm"

    'System Files klasör adını kontrol et.
    If Not Dir(AutoPath & "\System Files\", vbDirectory) <> vbNullString Then
        MsgBox AutoPath & "\System Files\" & " directory cannot be accessed. The folder named 'System Files' might have been renamed or deleted.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo SonCombo
    End If
    If Not Dir(DestTaslak & TaslakFile, vbDirectory) <> vbNullString Then
        MsgBox DestTaslak & TaslakFile & " directory cannot be accessed. Folder and/or file names in this directory might have been changed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo SonCombo
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
    objWord.Documents.Open FileName:=DestTaslak & TaslakFile
    Set objDoc = GetObject(DestTaslak & TaslakFile)
    'objDoc.ActiveWindow.Visible = False
    
    'Kurum_ANo getir
    Kurum_ANoStr = objDoc.Tables(1).Cell(Row:=4, Column:=1).Range.Text
    Kurum_ANo.Value = Mid(Kurum_ANoStr, 8, 8)
    
    Aktar.Visible = True
    
    Aktar.Value = objDoc.Sections(1).Footers(wdHeaderFooterFirstPage).Range.Tables(1).Cell(Row:=1, Column:=1).Range.Text
    
    objDoc.Close SaveChanges:=False
    objWord.Visible = False
    'Nesneleri temizle
    Set fso = Nothing
    Set objWord = Nothing
    Set objDoc = Nothing
    
    Aktar.SetFocus
    For i = 0 To Aktar.LineCount - 1
        If i = 0 Then
            Adres.Value = Split(Aktar.Value, Chr(13))(i)
        ElseIf i = 1 Then
            Tel.Value = Split(Aktar.Value, Chr(13))(i)
        ElseIf i = 2 Then
            Eposta.Value = Split(Aktar.Value, Chr(13))(i)
        ElseIf i = 3 Then
            ElektronikAg.Value = Split(Aktar.Value, Chr(13))(i)
        End If
    Next i
    
    Aktar.Visible = False
    
    'Birden fazla boşluk varsa kaldır
    'Sağdaki ve soldaki tek boşluğu kaldır
    For i = 1 To 50
        Adres.Value = Replace(Adres.Value, "  ", " ")
    Next i
    Do While Left(Adres.Value, 1) = " "
        Adres.Value = Right(Adres.Value, Len(Adres.Value) - 1)
    Loop
    Do While Right(Adres.Value, 1) = " "
        Adres.Value = Left(Adres.Value, Len(Adres.Value) - 1)
    Loop
    For i = 1 To 50
        Tel.Value = Replace(Tel.Value, "  ", " ")
    Next i
    Do While Left(Tel.Value, 1) = " "
        Tel.Value = Right(Tel.Value, Len(Tel.Value) - 1)
    Loop
    Do While Right(Tel.Value, 1) = " "
        Tel.Value = Left(Tel.Value, Len(Tel.Value) - 1)
    Loop
    For i = 1 To 50
        Eposta.Value = Replace(Eposta.Value, "  ", " ")
    Next i
    Do While Left(Eposta.Value, 1) = " "
        Eposta.Value = Right(Eposta.Value, Len(Eposta.Value) - 1)
    Loop
    Do While Right(Eposta.Value, 1) = " "
        Eposta.Value = Left(Eposta.Value, Len(Eposta.Value) - 1)
    Loop
    For i = 1 To 50
        ElektronikAg.Value = Replace(ElektronikAg.Value, "  ", " ")
    Next i
    Do While Left(ElektronikAg.Value, 1) = " "
        ElektronikAg.Value = Right(ElektronikAg.Value, Len(ElektronikAg.Value) - 1)
    Loop
    Do While Right(ElektronikAg.Value, 1) = " "
        ElektronikAg.Value = Left(ElektronikAg.Value, Len(ElektronikAg.Value) - 1)
    Loop
    
    'Paragraf işaretlerini kaldır.
    If InStr(Adres.Value, Chr(13)) > 0 Then
        Adres.Value = Right(Adres.Value, Len(Adres.Value) - 2)
    End If
    If InStr(Tel.Value, Chr(13)) > 0 Then
        Tel.Value = Right(Tel.Value, Len(Tel.Value) - 2)
    End If
    If InStr(Eposta.Value, Chr(13)) > 0 Then
        Eposta.Value = Right(Eposta.Value, Len(Eposta.Value) - 2)
    End If
    If InStr(ElektronikAg.Value, Chr(13)) > 0 Then
        ElektronikAg.Value = Right(ElektronikAg.Value, Len(ElektronikAg.Value) - 2)
    End If
    
SonCombo:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

End If
Atla:

ComboBirim.SetFocus
If ComboBirim.ListIndex = -1 And ComboBirim.Value <> "" Then
   ComboBirim.Value = ""
   GoTo Son
End If

If ComboBirim.Value <> "" Then
    ComboBirim.SelStart = 0
    ComboBirim.SelLength = Len(ComboBirim.Value)
End If
    
'ComboBirim.DropDown

Son:

End Sub


Sub ColorChangerGenel()

'Ekle
If Ekle.BackColor <> RGB(225, 235, 245) Then
    Ekle.BackColor = RGB(225, 235, 245)
    Ekle.ForeColor = RGB(30, 30, 30)
End If
'Getir
If Getir.BackColor <> RGB(225, 235, 245) Then
    Getir.BackColor = RGB(225, 235, 245)
    Getir.ForeColor = RGB(30, 30, 30)
End If
'Yardim
If Yardim.BackColor <> RGB(225, 235, 245) Then
    Yardim.BackColor = RGB(225, 235, 245)
    Yardim.ForeColor = RGB(30, 30, 30)
End If
'Kapat
If Kapat.BackColor <> RGB(225, 235, 245) Then
    Kapat.BackColor = RGB(225, 235, 245)
    Kapat.ForeColor = RGB(30, 30, 30)
End If


End Sub


Private Sub Ekle_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
Ekle.BackColor = RGB(60, 100, 180) 'RGB(60, 100, 180) 'RGB(60, 100, 180)
Ekle.ForeColor = RGB(255, 255, 255)
End Sub
Private Sub Kapat_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
Kapat.BackColor = RGB(60, 100, 180) 'RGB(60, 100, 180) 'RGB(60, 100, 180)
Kapat.ForeColor = RGB(255, 255, 255)
End Sub
Private Sub Getir_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
Getir.BackColor = RGB(60, 100, 180) 'RGB(60, 100, 180) 'RGB(60, 100, 180)
Getir.ForeColor = RGB(255, 255, 255)
End Sub

Private Sub Yardim_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
Yardim.BackColor = RGB(60, 100, 180) 'RGB(60, 100, 180) 'RGB(60, 100, 180)
Yardim.ForeColor = RGB(255, 255, 255)
End Sub

Private Sub Frame1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub
Private Sub Frame2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub
Private Sub Frame3_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
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

Private Sub Kurum_ANo_Change()

'Sadece numerik karakter
If Kurum_ANo.Value <> "" And IsNumeric(Kurum_ANo.Value) = False Then
    MsgBox "Organization label/code cannot contain non-numeric characters.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    Kurum_ANo.Value = ""
End If

End Sub

Private Sub Getir_Click()
Dim a() As Variant, i As Variant
Dim AutoPath As String, DestTaslak As String, OpenControl As String
Dim TaslakFile As String, SayHedef As Long, ItemName As String, j As Integer

Dim fso As Object, objWord As Object, objDoc As Object, Kurum_ANoStr As String


ThisWorkbook.Activate

Application.ScreenUpdating = False
Application.DisplayAlerts = False

AutoPath = ThisWorkbook.Path
DestTaslak = AutoPath & "\System Files\System Templates\Footer Field\"
TaslakFile = "Draft File.docm"

'System Files klasör adını kontrol et.
If Not Dir(AutoPath & "\System Files\", vbDirectory) <> vbNullString Then
    MsgBox AutoPath & "\System Files\" & " directory cannot be accessed. The folder named 'System Files' might have been renamed or deleted.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If
If Not Dir(DestTaslak & TaslakFile, vbDirectory) <> vbNullString Then
    MsgBox DestTaslak & TaslakFile & " directory cannot be accessed. Folder and/or file names in this directory might have been changed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
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
objWord.Documents.Open FileName:=DestTaslak & TaslakFile
Set objDoc = GetObject(DestTaslak & TaslakFile)
'objDoc.ActiveWindow.Visible = False

'Kurum_ANo getir
Kurum_ANoStr = objDoc.Tables(1).Cell(Row:=4, Column:=1).Range.Text
Kurum_ANo.Value = Mid(Kurum_ANoStr, 8, 8)

Aktar.Visible = True

Aktar.Value = objDoc.Sections(1).Footers(wdHeaderFooterFirstPage).Range.Tables(1).Cell(Row:=1, Column:=1).Range.Text

objDoc.Close SaveChanges:=False
objWord.Visible = False
'Nesneleri temizle
Set fso = Nothing
Set objWord = Nothing
Set objDoc = Nothing

ComboBirim.Value = "Demo-A Unit"
Aktar.SetFocus
For i = 0 To Aktar.LineCount - 1
    If i = 0 Then
        Adres.Value = Split(Aktar.Value, Chr(13))(i)
    ElseIf i = 1 Then
        Tel.Value = Split(Aktar.Value, Chr(13))(i)
    ElseIf i = 2 Then
        Eposta.Value = Split(Aktar.Value, Chr(13))(i)
    ElseIf i = 3 Then
        ElektronikAg.Value = Split(Aktar.Value, Chr(13))(i)
    End If
Next i

Aktar.Visible = False

'Birden fazla boşluk varsa kaldır
'Sağdaki ve soldaki tek boşluğu kaldır
For i = 1 To 50
    Adres.Value = Replace(Adres.Value, "  ", " ")
Next i
Do While Left(Adres.Value, 1) = " "
    Adres.Value = Right(Adres.Value, Len(Adres.Value) - 1)
Loop
Do While Right(Adres.Value, 1) = " "
    Adres.Value = Left(Adres.Value, Len(Adres.Value) - 1)
Loop
For i = 1 To 50
    Tel.Value = Replace(Tel.Value, "  ", " ")
Next i
Do While Left(Tel.Value, 1) = " "
    Tel.Value = Right(Tel.Value, Len(Tel.Value) - 1)
Loop
Do While Right(Tel.Value, 1) = " "
    Tel.Value = Left(Tel.Value, Len(Tel.Value) - 1)
Loop
For i = 1 To 50
    Eposta.Value = Replace(Eposta.Value, "  ", " ")
Next i
Do While Left(Eposta.Value, 1) = " "
    Eposta.Value = Right(Eposta.Value, Len(Eposta.Value) - 1)
Loop
Do While Right(Eposta.Value, 1) = " "
    Eposta.Value = Left(Eposta.Value, Len(Eposta.Value) - 1)
Loop
For i = 1 To 50
    ElektronikAg.Value = Replace(ElektronikAg.Value, "  ", " ")
Next i
Do While Left(ElektronikAg.Value, 1) = " "
    ElektronikAg.Value = Right(ElektronikAg.Value, Len(ElektronikAg.Value) - 1)
Loop
Do While Right(ElektronikAg.Value, 1) = " "
    ElektronikAg.Value = Left(ElektronikAg.Value, Len(ElektronikAg.Value) - 1)
Loop

'Paragraf işaretlerini kaldır.
If InStr(Adres.Value, Chr(13)) > 0 Then
    Adres.Value = Right(Adres.Value, Len(Adres.Value) - 2)
End If
If InStr(Tel.Value, Chr(13)) > 0 Then
    Tel.Value = Right(Tel.Value, Len(Tel.Value) - 2)
End If
If InStr(Eposta.Value, Chr(13)) > 0 Then
    Eposta.Value = Right(Eposta.Value, Len(Eposta.Value) - 2)
End If
If InStr(ElektronikAg.Value, Chr(13)) > 0 Then
    ElektronikAg.Value = Right(ElektronikAg.Value, Len(ElektronikAg.Value) - 2)
End If


Son:

'ThisWorkbook.Activate

Application.ScreenUpdating = True
Application.DisplayAlerts = True

End Sub

Private Sub Ekle_Click()
Dim a() As Variant, i As Variant, Tanimlar As String, DestTanimlar As String
Dim AutoPath As String, DestRapor3lar As String, OpenControl As String
Dim Rapor3_2File As String, SayHedef As Long, ItemName As String, j As Integer
Dim fso As Object, objWord As Object, objDoc As Object, Birimx As String, Kurum_A As String, Kurum_ANoStr As String
Dim Rapor3_1File As String, Rapor3_1TipBFile As String, FinansalBirimFile As String, RaporFile As String, Rapor1File As String
Dim DestRapor2 As String, DestRapor1 As String, SubeBosluk As String, DestPerformans As String, Performans As String
Dim BilgilendirmeFile, XXXMudGidenUstYaziFile, SonucUstYaziFile As String
Dim Rapor3_2FileTipB, FinansalBirimFileTipB As String
Dim MyFile As String, VarlikSablonlar As String, WsSablon As Object, DestAltBilgi As String
Dim AltBilgiFileKDisi As String, AltBilgiFileKIci As String


ThisWorkbook.Activate

Application.ScreenUpdating = False
Application.DisplayAlerts = False

AutoPath = ThisWorkbook.Path
DestTanimlar = AutoPath & "\System Files\System Definitions\"
DestPerformans = AutoPath & "\System Files\System Templates\Performance Report\"
DestRapor3lar = AutoPath & "\System Files\System Templates\Report 3 Cover Letter Templates\"
DestRapor1 = AutoPath & "\System Files\System Templates\Report 1 Cover Letter Template\"
DestRapor2 = AutoPath & "\System Files\System Templates\Report 2 Cover Letter Templates\"


DestAltBilgi = AutoPath & "\System Files\System Templates\Footer Field\"

Tanimlar = "Definitions.xlsx"
Performans = "Performance Report.xlsx"

Rapor3_2File = "Report 3.2 – Type A Cover Letter.docm"
Rapor3_2FileTipB = "Report 3.2 – Type B Cover Letter.docm"
Rapor3_1File = "Report 3.1 – Type A Cover Letter.docm"
Rapor3_1TipBFile = "Report 3.1 – Type B Cover Letter.docm"

FinansalBirimFile = "Report 3.2 – Type A Cover Letter – Financial Unit.docm"
FinansalBirimFileTipB = "Report 3.2 – Type B Cover Letter – Financial Unit.docm"
Rapor1File = "Report 1 Cover Letter.docm"
RaporFile = "Report 2 Cover Letter.docm"

BilgilendirmeFile = "Informative Cover Letter.docm"
XXXMudGidenUstYaziFile = "XXX Directorate Cover Letter.docm"
SonucUstYaziFile = "Final Cover Letter.docm"

AltBilgiFileKDisi = "External Footer.docm"
AltBilgiFileKIci = "Internal Footer.docm"




VarlikSablonlar = AutoPath & "\System Files\System Templates\Asset Templates\"


'Sistem Files folder check.
If Not Dir(AutoPath & "\System Files\", vbDirectory) <> vbNullString Then
    MsgBox "The directory " & AutoPath & "\System Files\" & " cannot be accessed. The folder named 'System Files' might have been renamed or deleted.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

If Not Dir(DestTanimlar & Tanimlar, vbDirectory) <> vbNullString Then
    MsgBox "The directory " & DestTanimlar & Tanimlar & " cannot be accessed. Folder or file names under this directory might have been changed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

If Not Dir(DestPerformans & Performans, vbDirectory) <> vbNullString Then
    MsgBox "The directory " & DestPerformans & Performans & " cannot be accessed. Folder or file names under this directory might have been changed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

If Not Dir(DestRapor3lar & Rapor3_2File, vbDirectory) <> vbNullString Then
    MsgBox "The directory " & DestRapor3lar & Rapor3_2File & " cannot be accessed. Folder or file names under this directory might have been changed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

If Not Dir(DestRapor3lar & Rapor3_1File, vbDirectory) <> vbNullString Then
    MsgBox "The directory " & DestRapor3lar & Rapor3_1File & " cannot be accessed. Folder or file names under this directory might have been changed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

If Not Dir(DestRapor3lar & Rapor3_1TipBFile, vbDirectory) <> vbNullString Then
    MsgBox "The directory " & DestRapor3lar & Rapor3_1TipBFile & " cannot be accessed. Folder or file names under this directory might have been changed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

If Not Dir(DestRapor3lar & FinansalBirimFile, vbDirectory) <> vbNullString Then
    MsgBox "The directory " & DestRapor3lar & FinansalBirimFile & " cannot be accessed. Folder or file names under this directory might have been changed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

If Not Dir(DestRapor1 & Rapor1File, vbDirectory) <> vbNullString Then
    MsgBox "The directory " & DestRapor1 & Rapor1File & " cannot be accessed. Folder or file names under this directory might have been changed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

If Not Dir(DestRapor2 & RaporFile, vbDirectory) <> vbNullString Then
    MsgBox "The directory " & DestRapor2 & RaporFile & " cannot be accessed. Folder or file names under this directory might have been changed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

If Not Dir(DestRapor2 & BilgilendirmeFile, vbDirectory) <> vbNullString Then
    MsgBox "The directory " & DestRapor2 & BilgilendirmeFile & " cannot be accessed. Folder or file names under this directory might have been changed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

If Not Dir(DestRapor2 & XXXMudGidenUstYaziFile, vbDirectory) <> vbNullString Then
    MsgBox "The directory " & DestRapor2 & XXXMudGidenUstYaziFile & " cannot be accessed. Folder or file names under this directory might have been changed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

If Not Dir(DestRapor2 & SonucUstYaziFile, vbDirectory) <> vbNullString Then
    MsgBox "The directory " & DestRapor2 & SonucUstYaziFile & " cannot be accessed. Folder or file names under this directory might have been changed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

If Not Dir(DestRapor3lar & Rapor3_2FileTipB, vbDirectory) <> vbNullString Then
    MsgBox "The directory " & DestRapor3lar & Rapor3_2FileTipB & " cannot be accessed. Folder or file names under this directory might have been changed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

If Not Dir(DestRapor3lar & FinansalBirimFileTipB, vbDirectory) <> vbNullString Then
    MsgBox "The directory " & DestRapor3lar & FinansalBirimFileTipB & " cannot be accessed. Folder or file names under this directory might have been changed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

If Not Dir(VarlikSablonlar, vbDirectory) <> vbNullString Then
    MsgBox "The directory " & VarlikSablonlar & " cannot be accessed. The folder name might have been changed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

If Not Dir(DestAltBilgi & AltBilgiFileKDisi, vbDirectory) <> vbNullString Then
    MsgBox "The directory " & DestAltBilgi & AltBilgiFileKDisi & " cannot be accessed. Folder or file names under this directory might have been changed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

If Not Dir(DestAltBilgi & AltBilgiFileKIci, vbDirectory) <> vbNullString Then
    MsgBox "The directory " & DestAltBilgi & AltBilgiFileKIci & " cannot be accessed. Folder or file names under this directory might have been changed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If



'Birden fazla boşluk varsa kaldır
'Sağdaki ve soldaki tek boşluğu kaldır
For i = 1 To 50
    Kurum_ANo.Value = Replace(Kurum_ANo.Value, "  ", " ")
Next i
Do While Left(Kurum_ANo.Value, 1) = " "
    Kurum_ANo.Value = Right(Kurum_ANo.Value, Len(Kurum_ANo.Value) - 1)
Loop
Do While Right(Kurum_ANo.Value, 1) = " "
    Kurum_ANo.Value = Left(Kurum_ANo.Value, Len(Kurum_ANo.Value) - 1)
Loop
If Len(Kurum_ANo.Value) <> 8 Then
    MsgBox "The Organization Tag / Code must consist of 8 numeric characters. To make changes, please enter an 8-digit organization code.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

If ComboBirim.Value <> "" Then
    ThisWorkbook.Unprotect "123"
    Worksheets(2).Unprotect Password:="123"
    Worksheets(2).Cells(6, 99).Value = Left(ComboBirim.Value, Len(ComboBirim.Value) - 5)
    Worksheets(2).Protect Password:="123" ', DrawingObjects:=False
    ThisWorkbook.Protect "123"
    'MsgBox Left(ComboBirim.Value, Len(ComboBirim.Value) - 5)

    '____________Güncelleme 14112019
    OpenControl = IsFileOpen(DestTanimlar & Tanimlar)
    If OpenControl = True Then 'Açıksa
        Workbooks(Tanimlar).Close SaveChanges:=True
    End If
    Workbooks.Open (DestTanimlar & Tanimlar)
    Workbooks(Tanimlar).Worksheets(1).Activate
    
    Workbooks(Tanimlar).Worksheets(1).Unprotect Password:="123"
    Workbooks(Tanimlar).Worksheets(1).Cells(6, 99).Value = Left(ComboBirim.Value, Len(ComboBirim.Value) - 5)
    Workbooks(Tanimlar).Worksheets(1).Protect Password:="123" ', DrawingObjects:=False
    
    Workbooks(Tanimlar).Save
    OpenControl = IsFileOpen(DestTanimlar & Tanimlar)
    If OpenControl = True Then 'Açıksa
        Workbooks(Tanimlar).Close SaveChanges:=True
    End If
    
    '____________Güncelleme 14112019
    
Else
    MsgBox "Please select your unit.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

'Birimx = UCase(Replace(Replace(Worksheets(2).Cells(6, 99).Value, "i", "I"), "ı", "I")) & " UNIT"
Birimx = Worksheets(2).Cells(6, 99).Value & " Unit"
Kurum_A = "ORGANIZATION A" & vbNewLine & Birimx

'Birden fazla boşluk varsa kaldır
'Sağdaki ve soldaki tek boşluğu kaldır
For i = 1 To 50
    Adres.Value = Replace(Adres.Value, "  ", " ")
Next i
Do While Left(Adres.Value, 1) = " "
    Adres.Value = Right(Adres.Value, Len(Adres.Value) - 1)
Loop
Do While Right(Adres.Value, 1) = " "
    Adres.Value = Left(Adres.Value, Len(Adres.Value) - 1)
Loop
For i = 1 To 50
    Tel.Value = Replace(Tel.Value, "  ", " ")
Next i
Do While Left(Tel.Value, 1) = " "
    Tel.Value = Right(Tel.Value, Len(Tel.Value) - 1)
Loop
Do While Right(Tel.Value, 1) = " "
    Tel.Value = Left(Tel.Value, Len(Tel.Value) - 1)
Loop
For i = 1 To 50
    Eposta.Value = Replace(Eposta.Value, "  ", " ")
Next i
Do While Left(Eposta.Value, 1) = " "
    Eposta.Value = Right(Eposta.Value, Len(Eposta.Value) - 1)
Loop
Do While Right(Eposta.Value, 1) = " "
    Eposta.Value = Left(Eposta.Value, Len(Eposta.Value) - 1)
Loop
For i = 1 To 50
    ElektronikAg.Value = Replace(ElektronikAg.Value, "  ", " ")
Next i
Do While Left(ElektronikAg.Value, 1) = " "
    ElektronikAg.Value = Right(ElektronikAg.Value, Len(ElektronikAg.Value) - 1)
Loop
Do While Right(ElektronikAg.Value, 1) = " "
    ElektronikAg.Value = Left(ElektronikAg.Value, Len(ElektronikAg.Value) - 1)
Loop

'Paragraf işaretlerini kaldır.
If InStr(Adres.Value, Chr(13)) > 0 Then
    Adres.Value = Right(Adres.Value, Len(Adres.Value) - 2)
End If
If InStr(Tel.Value, Chr(13)) > 0 Then
    Tel.Value = Right(Tel.Value, Len(Tel.Value) - 2)
End If
If InStr(Eposta.Value, Chr(13)) > 0 Then
    Eposta.Value = Right(Eposta.Value, Len(Eposta.Value) - 2)
End If
If InStr(ElektronikAg.Value, Chr(13)) > 0 Then
    ElektronikAg.Value = Right(ElektronikAg.Value, Len(ElektronikAg.Value) - 2)
End If

'Tanımları güncelle
OpenControl = IsFileOpen(DestTanimlar & Tanimlar)
If OpenControl = True Then 'Açıksa
    Workbooks(Tanimlar).Close SaveChanges:=True
ElseIf OpenControl = False Then
    '
End If
Workbooks.Open (DestTanimlar & Tanimlar)
Workbooks(Tanimlar).Worksheets(1).Activate
Workbooks(Tanimlar).Worksheets(1).Unprotect Password:="123"
Workbooks(Tanimlar).Worksheets(1).Cells(6, 99).Value = Left(ComboBirim.Value, Len(ComboBirim.Value) - 5)
Workbooks(Tanimlar).Worksheets(1).Protect Password:="123" ', DrawingObjects:=False
OpenControl = IsFileOpen(DestTanimlar & Tanimlar)
If OpenControl = True Then 'Açıksa
    Workbooks(Tanimlar).Close SaveChanges:=True
ElseIf OpenControl = False Then
    '
End If
ThisWorkbook.Activate

'Performans güncelle
OpenControl = IsFileOpen(DestPerformans & Performans)
If OpenControl = True Then 'Açıksa
    Workbooks(Performans).Close SaveChanges:=True
ElseIf OpenControl = False Then
    '
End If
Workbooks.Open (DestPerformans & Performans)
Workbooks(Performans).Worksheets(1).Activate
Workbooks(Performans).Worksheets(1).Unprotect Password:="123"
Workbooks(Performans).Worksheets(1).Cells(2, 2).Value = Kurum_A 'Left(ComboBirim.Value, Len(ComboBirim.Value) - 5)
Workbooks(Performans).Worksheets(1).Protect Password:="123", DrawingObjects:=False

Workbooks(Performans).Save
OpenControl = IsFileOpen(DestPerformans & Performans)
If OpenControl = True Then 'Açıksa
    Workbooks(Performans).Close SaveChanges:=True
ElseIf OpenControl = False Then
    
End If


ThisWorkbook.Activate

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
    objWord.Visible = False 'True 'False
End If
objWord.Visible = False

'Report 3.2 – Type A Cover Letter
objWord.Documents.Open FileName:=DestRapor3lar & Rapor3_2File
Set objDoc = GetObject(DestRapor3lar & Rapor3_2File)
'objDoc.ActiveWindow.Visible = False
objDoc.Sections(1).Footers(wdHeaderFooterFirstPage).Range.Tables(1).Cell(Row:=1, Column:=1).Range.Text = Adres.Value & vbNewLine & _
                                                                                   Tel.Value & vbNewLine & _
                                                                                   Eposta.Value & vbNewLine & _
                                                                                   ElektronikAg.Value
objDoc.Sections(1).Footers(wdHeaderFooterPrimary).Range.Tables(1).Cell(Row:=1, Column:=1).Range.Text = Adres.Value & vbNewLine & _
                                                                                   Tel.Value & vbNewLine & _
                                                                                   Eposta.Value & vbNewLine & _
                                                                                   ElektronikAg.Value
'Kurum_ANo
Kurum_ANoStr = objDoc.Tables(1).Cell(Row:=4, Column:=1).Range.Text
Kurum_ANoStr = Left(Kurum_ANoStr, 7) & Kurum_ANo.Value & Mid(Kurum_ANoStr, 16, Len(Kurum_ANoStr) - 17)
'MsgBox Kurum_ANoStr
objDoc.Tables(1).Cell(Row:=4, Column:=1).Range.Text = Kurum_ANoStr
'KURUM_A
objDoc.Tables(2).Cell(Row:=4, Column:=1).Range.Text = Kurum_A
objDoc.Close SaveChanges:=True
objWord.Visible = False

'Report 3.2 – Type B Cover Letter
objWord.Documents.Open FileName:=DestRapor3lar & Rapor3_2FileTipB
Set objDoc = GetObject(DestRapor3lar & Rapor3_2FileTipB)
'objDoc.ActiveWindow.Visible = False
objDoc.Sections(1).Footers(wdHeaderFooterFirstPage).Range.Tables(1).Cell(Row:=1, Column:=1).Range.Text = Adres.Value & vbNewLine & _
                                                                                   Tel.Value & vbNewLine & _
                                                                                   Eposta.Value & vbNewLine & _
                                                                                   ElektronikAg.Value
objDoc.Sections(1).Footers(wdHeaderFooterPrimary).Range.Tables(1).Cell(Row:=1, Column:=1).Range.Text = Adres.Value & vbNewLine & _
                                                                                   Tel.Value & vbNewLine & _
                                                                                   Eposta.Value & vbNewLine & _
                                                                                   ElektronikAg.Value
'Kurum_ANo
Kurum_ANoStr = objDoc.Tables(1).Cell(Row:=4, Column:=1).Range.Text
Kurum_ANoStr = Left(Kurum_ANoStr, 7) & Kurum_ANo.Value & Mid(Kurum_ANoStr, 16, Len(Kurum_ANoStr) - 17)
'MsgBox Kurum_ANoStr
objDoc.Tables(1).Cell(Row:=4, Column:=1).Range.Text = Kurum_ANoStr
'KURUM_A
objDoc.Tables(2).Cell(Row:=4, Column:=1).Range.Text = Kurum_A
objDoc.Close SaveChanges:=True
objWord.Visible = False

'Report 3.1 – Type A Cover Letter
objWord.Documents.Open FileName:=DestRapor3lar & Rapor3_1File
Set objDoc = GetObject(DestRapor3lar & Rapor3_1File)
objDoc.Sections(1).Footers(wdHeaderFooterFirstPage).Range.Tables(1).Cell(Row:=1, Column:=1).Range.Text = Adres.Value & vbNewLine & _
                                                                                   Tel.Value & vbNewLine & _
                                                                                   Eposta.Value & vbNewLine & _
                                                                                   ElektronikAg.Value
objDoc.Sections(1).Footers(wdHeaderFooterPrimary).Range.Tables(1).Cell(Row:=1, Column:=1).Range.Text = Adres.Value & vbNewLine & _
                                                                                   Tel.Value & vbNewLine & _
                                                                                   Eposta.Value & vbNewLine & _
                                                                                   ElektronikAg.Value
'Kurum_ANo
Kurum_ANoStr = objDoc.Tables(1).Cell(Row:=4, Column:=1).Range.Text
Kurum_ANoStr = Left(Kurum_ANoStr, 7) & Kurum_ANo.Value & Mid(Kurum_ANoStr, 16, Len(Kurum_ANoStr) - 17)
'MsgBox Kurum_ANoStr
objDoc.Tables(1).Cell(Row:=4, Column:=1).Range.Text = Kurum_ANoStr
'KURUM_A
objDoc.Tables(2).Cell(Row:=4, Column:=1).Range.Text = Kurum_A
objDoc.Close SaveChanges:=True
objWord.Visible = False

'Report 3.1 – Type B Üst Yazı
objWord.Documents.Open FileName:=DestRapor3lar & Rapor3_1TipBFile
Set objDoc = GetObject(DestRapor3lar & Rapor3_1TipBFile)
objDoc.Sections(1).Footers(wdHeaderFooterFirstPage).Range.Tables(1).Cell(Row:=1, Column:=1).Range.Text = Adres.Value & vbNewLine & _
                                                                                   Tel.Value & vbNewLine & _
                                                                                   Eposta.Value & vbNewLine & _
                                                                                   ElektronikAg.Value
objDoc.Sections(1).Footers(wdHeaderFooterPrimary).Range.Tables(1).Cell(Row:=1, Column:=1).Range.Text = Adres.Value & vbNewLine & _
                                                                                   Tel.Value & vbNewLine & _
                                                                                   Eposta.Value & vbNewLine & _
                                                                                   ElektronikAg.Value
'Kurum_ANo
Kurum_ANoStr = objDoc.Tables(1).Cell(Row:=4, Column:=1).Range.Text
Kurum_ANoStr = Left(Kurum_ANoStr, 7) & Kurum_ANo.Value & Mid(Kurum_ANoStr, 16, Len(Kurum_ANoStr) - 17)
'MsgBox Kurum_ANoStr
objDoc.Tables(1).Cell(Row:=4, Column:=1).Range.Text = Kurum_ANoStr
'KURUM_A
objDoc.Tables(2).Cell(Row:=4, Column:=1).Range.Text = Kurum_A
objDoc.Close SaveChanges:=True
objWord.Visible = False

'FinansalBirim Üst Yazı
objWord.Documents.Open FileName:=DestRapor3lar & FinansalBirimFile
Set objDoc = GetObject(DestRapor3lar & FinansalBirimFile)
objDoc.Sections(1).Footers(wdHeaderFooterFirstPage).Range.Tables(1).Cell(Row:=1, Column:=1).Range.Text = Adres.Value & vbNewLine & _
                                                                                   Tel.Value & vbNewLine & _
                                                                                   Eposta.Value & vbNewLine & _
                                                                                   ElektronikAg.Value
objDoc.Sections(1).Footers(wdHeaderFooterPrimary).Range.Tables(1).Cell(Row:=1, Column:=1).Range.Text = Adres.Value & vbNewLine & _
                                                                                   Tel.Value & vbNewLine & _
                                                                                   Eposta.Value & vbNewLine & _
                                                                                   ElektronikAg.Value
'Kurum_ANo
Kurum_ANoStr = objDoc.Tables(1).Cell(Row:=4, Column:=1).Range.Text
Kurum_ANoStr = Left(Kurum_ANoStr, 7) & Kurum_ANo.Value & Mid(Kurum_ANoStr, 16, Len(Kurum_ANoStr) - 17)
'MsgBox Kurum_ANoStr
objDoc.Tables(1).Cell(Row:=4, Column:=1).Range.Text = Kurum_ANoStr
'KURUM_A
objDoc.Tables(2).Cell(Row:=4, Column:=1).Range.Text = Kurum_A
objDoc.Close SaveChanges:=True
objWord.Visible = False

'Report 3.2 – Type B Cover Letter – Financial Unit
objWord.Documents.Open FileName:=DestRapor3lar & FinansalBirimFileTipB
Set objDoc = GetObject(DestRapor3lar & FinansalBirimFileTipB)
objDoc.Sections(1).Footers(wdHeaderFooterFirstPage).Range.Tables(1).Cell(Row:=1, Column:=1).Range.Text = Adres.Value & vbNewLine & _
                                                                                   Tel.Value & vbNewLine & _
                                                                                   Eposta.Value & vbNewLine & _
                                                                                   ElektronikAg.Value
objDoc.Sections(1).Footers(wdHeaderFooterPrimary).Range.Tables(1).Cell(Row:=1, Column:=1).Range.Text = Adres.Value & vbNewLine & _
                                                                                   Tel.Value & vbNewLine & _
                                                                                   Eposta.Value & vbNewLine & _
                                                                                   ElektronikAg.Value
'Kurum_ANo
Kurum_ANoStr = objDoc.Tables(1).Cell(Row:=4, Column:=1).Range.Text
Kurum_ANoStr = Left(Kurum_ANoStr, 7) & Kurum_ANo.Value & Mid(Kurum_ANoStr, 16, Len(Kurum_ANoStr) - 17)
'MsgBox Kurum_ANoStr
objDoc.Tables(1).Cell(Row:=4, Column:=1).Range.Text = Kurum_ANoStr
'KURUM_A
objDoc.Tables(2).Cell(Row:=4, Column:=1).Range.Text = Kurum_A
objDoc.Close SaveChanges:=True
objWord.Visible = False


'Report 1 Cover Letter
objWord.Documents.Open FileName:=DestRapor1 & Rapor1File
Set objDoc = GetObject(DestRapor1 & Rapor1File)
objDoc.Sections(1).Footers(wdHeaderFooterFirstPage).Range.Tables(1).Cell(Row:=1, Column:=1).Range.Text = Adres.Value & vbNewLine & _
                                                                                   Tel.Value & vbNewLine & _
                                                                                   Eposta.Value & vbNewLine & _
                                                                                   ElektronikAg.Value
objDoc.Sections(1).Footers(wdHeaderFooterPrimary).Range.Tables(1).Cell(Row:=1, Column:=1).Range.Text = Adres.Value & vbNewLine & _
                                                                                   Tel.Value & vbNewLine & _
                                                                                   Eposta.Value & vbNewLine & _
                                                                                   ElektronikAg.Value
'Kurum_ANo
Kurum_ANoStr = objDoc.Tables(1).Cell(Row:=4, Column:=1).Range.Text
Kurum_ANoStr = Left(Kurum_ANoStr, 7) & Kurum_ANo.Value & Mid(Kurum_ANoStr, 16, Len(Kurum_ANoStr) - 17)
'MsgBox Kurum_ANoStr
objDoc.Tables(1).Cell(Row:=4, Column:=1).Range.Text = Kurum_ANoStr
'KURUM_A
objDoc.Tables(2).Cell(Row:=4, Column:=1).Range.Text = Kurum_A
objDoc.Close SaveChanges:=True
objWord.Visible = False

'Report 2 Cover Letter
objWord.Documents.Open FileName:=DestRapor2 & RaporFile
Set objDoc = GetObject(DestRapor2 & RaporFile)
objDoc.Sections(1).Footers(wdHeaderFooterFirstPage).Range.Tables(1).Cell(Row:=1, Column:=1).Range.Text = Adres.Value & vbNewLine & _
                                                                                   Tel.Value & vbNewLine & _
                                                                                   Eposta.Value & vbNewLine & _
                                                                                   ElektronikAg.Value
objDoc.Sections(1).Footers(wdHeaderFooterPrimary).Range.Tables(1).Cell(Row:=1, Column:=1).Range.Text = Adres.Value & vbNewLine & _
                                                                                   Tel.Value & vbNewLine & _
                                                                                   Eposta.Value & vbNewLine & _
                                                                                   ElektronikAg.Value
'Kurum_ANo
Kurum_ANoStr = objDoc.Tables(1).Cell(Row:=4, Column:=1).Range.Text
Kurum_ANoStr = Left(Kurum_ANoStr, 7) & Kurum_ANo.Value & Mid(Kurum_ANoStr, 16, Len(Kurum_ANoStr) - 17)
'MsgBox Kurum_ANoStr
objDoc.Tables(1).Cell(Row:=4, Column:=1).Range.Text = Kurum_ANoStr
'KURUM_A
objDoc.Tables(2).Cell(Row:=4, Column:=1).Range.Text = Kurum_A
objDoc.Close SaveChanges:=True
objWord.Visible = False

'Informative Cover Letter
objWord.Documents.Open FileName:=DestRapor2 & BilgilendirmeFile
Set objDoc = GetObject(DestRapor2 & BilgilendirmeFile)
objDoc.Sections(1).Footers(wdHeaderFooterFirstPage).Range.Tables(1).Cell(Row:=1, Column:=1).Range.Text = Adres.Value & vbNewLine & _
                                                                                   Tel.Value & vbNewLine & _
                                                                                   Eposta.Value & vbNewLine & _
                                                                                   ElektronikAg.Value
objDoc.Sections(1).Footers(wdHeaderFooterPrimary).Range.Tables(1).Cell(Row:=1, Column:=1).Range.Text = Adres.Value & vbNewLine & _
                                                                                   Tel.Value & vbNewLine & _
                                                                                   Eposta.Value & vbNewLine & _
                                                                                   ElektronikAg.Value
'Kurum_ANo
Kurum_ANoStr = objDoc.Tables(1).Cell(Row:=4, Column:=1).Range.Text
Kurum_ANoStr = Left(Kurum_ANoStr, 7) & Kurum_ANo.Value & Mid(Kurum_ANoStr, 16, Len(Kurum_ANoStr) - 17)
'MsgBox Kurum_ANoStr
objDoc.Tables(1).Cell(Row:=4, Column:=1).Range.Text = Kurum_ANoStr
'KURUM_A
objDoc.Tables(2).Cell(Row:=4, Column:=1).Range.Text = Kurum_A
objDoc.Close SaveChanges:=True
objWord.Visible = False

'XXXMud Giden Üst Yazı
objWord.Documents.Open FileName:=DestRapor2 & XXXMudGidenUstYaziFile
Set objDoc = GetObject(DestRapor2 & XXXMudGidenUstYaziFile)
objDoc.Sections(1).Footers(wdHeaderFooterFirstPage).Range.Tables(1).Cell(Row:=1, Column:=2).Range.Text = Adres.Value & vbNewLine & _
                                                                                   Tel.Value & vbNewLine & _
                                                                                   Eposta.Value & vbNewLine & _
                                                                                   ElektronikAg.Value
objDoc.Sections(1).Footers(wdHeaderFooterPrimary).Range.Tables(1).Cell(Row:=1, Column:=2).Range.Text = Adres.Value & vbNewLine & _
                                                                                   Tel.Value & vbNewLine & _
                                                                                   Eposta.Value & vbNewLine & _
                                                                                   ElektronikAg.Value
'Kurum_ANo
Kurum_ANoStr = objDoc.Tables(1).Cell(Row:=6, Column:=1).Range.Text
Kurum_ANoStr = Left(Kurum_ANoStr, 7) & Kurum_ANo.Value & Mid(Kurum_ANoStr, 16, Len(Kurum_ANoStr) - 17)
'MsgBox Kurum_ANoStr
objDoc.Tables(1).Cell(Row:=6, Column:=1).Range.Text = Kurum_ANoStr
''KURUM_A 'KURUM_A Birim adı olmayacak; bu kısmı kurum içi imza prosedürü düzenleyecek.
'objDoc.Tables(2).Cell(Row:=4, Column:=1).Range.Text = Kurum_A
objDoc.Close SaveChanges:=True
objWord.Visible = False

'Final Cover Letter
objWord.Documents.Open FileName:=DestRapor2 & SonucUstYaziFile
Set objDoc = GetObject(DestRapor2 & SonucUstYaziFile)
objDoc.Sections(1).Footers(wdHeaderFooterFirstPage).Range.Tables(1).Cell(Row:=1, Column:=1).Range.Text = Adres.Value & vbNewLine & _
                                                                                   Tel.Value & vbNewLine & _
                                                                                   Eposta.Value & vbNewLine & _
                                                                                   ElektronikAg.Value
objDoc.Sections(1).Footers(wdHeaderFooterPrimary).Range.Tables(1).Cell(Row:=1, Column:=1).Range.Text = Adres.Value & vbNewLine & _
                                                                                   Tel.Value & vbNewLine & _
                                                                                   Eposta.Value & vbNewLine & _
                                                                                   ElektronikAg.Value
'Kurum_ANo
Kurum_ANoStr = objDoc.Tables(1).Cell(Row:=4, Column:=1).Range.Text
Kurum_ANoStr = Left(Kurum_ANoStr, 7) & Kurum_ANo.Value & Mid(Kurum_ANoStr, 16, Len(Kurum_ANoStr) - 17)
'MsgBox Kurum_ANoStr
objDoc.Tables(1).Cell(Row:=4, Column:=1).Range.Text = Kurum_ANoStr
'KURUM_A
objDoc.Tables(2).Cell(Row:=4, Column:=1).Range.Text = Kurum_A
objDoc.Close SaveChanges:=True
objWord.Visible = False

'External Footer
objWord.Documents.Open FileName:=DestAltBilgi & AltBilgiFileKDisi
Set objDoc = GetObject(DestAltBilgi & AltBilgiFileKDisi)
objDoc.Sections(1).Footers(wdHeaderFooterFirstPage).Range.Tables(1).Cell(Row:=1, Column:=1).Range.Text = Adres.Value & vbNewLine & _
                                                                                   Tel.Value & vbNewLine & _
                                                                                   Eposta.Value & vbNewLine & _
                                                                                   ElektronikAg.Value
objDoc.Sections(1).Footers(wdHeaderFooterPrimary).Range.Tables(1).Cell(Row:=1, Column:=1).Range.Text = Adres.Value & vbNewLine & _
                                                                                   Tel.Value & vbNewLine & _
                                                                                   Eposta.Value & vbNewLine & _
                                                                                   ElektronikAg.Value
'Kurum_ANo
Kurum_ANoStr = objDoc.Tables(1).Cell(Row:=4, Column:=1).Range.Text
Kurum_ANoStr = Left(Kurum_ANoStr, 7) & Kurum_ANo.Value & Mid(Kurum_ANoStr, 16, Len(Kurum_ANoStr) - 17)
'MsgBox Kurum_ANoStr
objDoc.Tables(1).Cell(Row:=4, Column:=1).Range.Text = Kurum_ANoStr
'KURUM_A
objDoc.Tables(2).Cell(Row:=4, Column:=1).Range.Text = Kurum_A
objDoc.Close SaveChanges:=True
objWord.Visible = False

'Internal Footer
objWord.Documents.Open FileName:=DestAltBilgi & AltBilgiFileKIci
Set objDoc = GetObject(DestAltBilgi & AltBilgiFileKIci)
objDoc.Sections(1).Footers(wdHeaderFooterFirstPage).Range.Tables(1).Cell(Row:=1, Column:=2).Range.Text = Adres.Value & vbNewLine & _
                                                                                   Tel.Value & vbNewLine & _
                                                                                   Eposta.Value & vbNewLine & _
                                                                                   ElektronikAg.Value
objDoc.Sections(1).Footers(wdHeaderFooterPrimary).Range.Tables(1).Cell(Row:=1, Column:=2).Range.Text = Adres.Value & vbNewLine & _
                                                                                   Tel.Value & vbNewLine & _
                                                                                   Eposta.Value & vbNewLine & _
                                                                                   ElektronikAg.Value
'Kurum_ANo
Kurum_ANoStr = objDoc.Tables(1).Cell(Row:=6, Column:=1).Range.Text
Kurum_ANoStr = Left(Kurum_ANoStr, 7) & Kurum_ANo.Value & Mid(Kurum_ANoStr, 16, Len(Kurum_ANoStr) - 17)
'MsgBox Kurum_ANoStr
objDoc.Tables(1).Cell(Row:=6, Column:=1).Range.Text = Kurum_ANoStr
''KURUM_A 'KURUM_A Birim adı olmayacak; bu kısmı kurum içi imza prosedürü düzenleyecek.
'objDoc.Tables(2).Cell(Row:=4, Column:=1).Range.Text = Kurum_A
objDoc.Close SaveChanges:=True
objWord.Visible = False



ThisWorkbook.Activate
For i = 1 To Len(Birimx)
    If i < Len(Birimx) Then 'Karakterlerden sonra boşluk ekle
        SubeBosluk = SubeBosluk & Mid(Birimx, i, 1) & " "
        'MsgBox SubeBosluk
    Else 'Son karakterden sonra boşluk ekleme
        SubeBosluk = SubeBosluk & Mid(Birimx, i, 1)
        'MsgBox SubeBosluk
    End If
Next i
SubeBosluk = UCase(Replace(Replace(SubeBosluk, "i", "I"), "ı", "I"))

'Ana Sayfa
ThisWorkbook.Unprotect "123"
Worksheets(1).Unprotect Password:="123"
Worksheets(1).Cells(7, 3).Value = UCase(Replace(Replace(Birimx, "i", "I"), "ı", "I"))
Worksheets(1).Protect Password:="123" ', DrawingObjects:=False
'Rapor1
Worksheets(3).Unprotect Password:="123"
Worksheets(3).Cells(2, 11).Value = UCase(Replace(Replace(Birimx, "i", "I"), "ı", "I"))
Worksheets(3).Protect Password:="123" ', DrawingObjects:=False
'Rapor
Worksheets(4).Unprotect Password:="123"
Worksheets(4).Cells(2, 21).Value = UCase(Replace(Replace(Birimx, "i", "I"), "ı", "I"))
Worksheets(4).Protect Password:="123" ', DrawingObjects:=False
'Rapor3 İşlemleri
Worksheets(5).Unprotect Password:="123"
Worksheets(5).Cells(2, 13).Value = UCase(Replace(Replace(Birimx, "i", "I"), "ı", "I"))
Worksheets(5).Protect Password:="123" ', DrawingObjects:=False
'MsgBox UCase(Replace(Replace(SubeBosluk, "i", "I"), "ı", "I"))
ThisWorkbook.Protect "123"

'VARLIKLAR
'Açık dosyaları kapat
On Error Resume Next
MyFile = Dir(VarlikSablonlar & "*.xl??")
Do While MyFile <> ""
    DoEvents
    Workbooks(MyFile).Close SaveChanges:=True
    MyFile = Dir
Loop
On Error GoTo 0

'Dosyaların birim adını değiştir.
'On Error Resume Next
MyFile = Dir(VarlikSablonlar & "*.xl??")
Do While MyFile <> ""
    DoEvents
    Workbooks.Open (VarlikSablonlar & MyFile)
    Set WsSablon = Workbooks(MyFile).Worksheets(1)
    WsSablon.Unprotect Password:="123"

    WsSablon.Range("C2:E2").UnMerge
    WsSablon.Range("C2") = UCase(Replace(Replace(Birimx, "i", "I"), "ı", "I"))
    WsSablon.Range("C2:E2").Merge
    WsSablon.Range("C2:E2").HorizontalAlignment = xlLeft
    
    WsSablon.Activate
    WsSablon.Protect Password:="123"
    ActiveWorkbook.Save
    Workbooks(MyFile).Close SaveChanges:=False
    MyFile = Dir
Loop
On Error GoTo 0

ThisWorkbook.Activate

'Açık dropdown kapat
Call ModuleSystemSettings.DropDownKapat

MsgBox "The system has been successfully customized according to the information you provided.", vbOKOnly + vbInformation, "Enterprise Document Automation System"


'GoTo Atla
'Descrp:
'If Err.Number <> 0 Then
'MsgBox "Error # " & Str(Err.Number)
'End If
'Atla:


Son:

'ThisWorkbook.Activate

Application.ScreenUpdating = True
Application.DisplayAlerts = True

End Sub

'Private Sub LabelGuncelle_Click()
'
''Call Ekle_Click
'
'End Sub

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
'Draft File
SourceTaslak = AutoPath & "\System Files\Help Documents\Unit Settings Panel – Help.docm"
'Birden fazla kullanıcının operasyon klasöründe çakışmasını önlemek için operasyon klasörünün içinde yeni klasör aç.
DestOpUserFolderName = "Operation-" & Left(ThisWorkbook.name, InStr(ThisWorkbook.name, ".") - 1) '& " " & Format(Now(), "ddmmyyyyhhmmss")
DestOpUserFolder = DestOperasyon & DestOpUserFolderName & "\"

'Sistem Files folder name check.
If Not Dir(AutoPath & "\System Files\", vbDirectory) <> vbNullString Then
    MsgBox "Cannot access directory: " & AutoPath & "\System Files\" & ". The folder named 'System Files' may have been renamed or deleted.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

'Operation folder name check.
If Not Dir(DestOperasyon, vbDirectory) <> vbNullString Then
    MsgBox "Cannot access directory: " & DestOperasyon & ". The folder named 'Operation' may have been renamed or deleted.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

'RmDir DestOpUserFolder 'Clean DestOpUserFolder on system shutdown - TO BE ADDED!
'_______________

'Folder names check.
If Not Dir(SourceTaslak, vbDirectory) <> vbNullString Then
    MsgBox "Cannot access directory: " & SourceTaslak & ". Folder or file names within this directory may have been changed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
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

Private Sub Kapat_Click()
    Unload Me
End Sub

Private Sub UserForm_Initialize()
Dim i As Long, WsSKP As Object
Dim ClrLab As MSForms.Control


ThisWorkbook.Activate


For Each ClrLab In core_unit_settings_UI.Controls
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

Ekle.BackColor = RGB(225, 235, 245)
Ekle.ForeColor = RGB(30, 30, 30)
Kapat.BackColor = RGB(225, 235, 245)
Kapat.ForeColor = RGB(30, 30, 30)
Getir.BackColor = RGB(225, 235, 245)
Getir.ForeColor = RGB(30, 30, 30)
Yardim.BackColor = RGB(225, 235, 245)
Yardim.ForeColor = RGB(30, 30, 30)
core_unit_settings_UI.BackColor = RGB(230, 230, 230) 'YENİ

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
Dim yukseklik As Variant, genislik As Variant
Dim Rep As Variant


yukseklik = Me.Height
Rep = Me.Width
Do
DoEvents
Rep = Rep - 70
Call timeout(0.01)
    If Rep > 70 Then
        core_unit_settings_UI.Width = Rep
        yukseklik = yukseklik - 70
        core_unit_settings_UI.Height = yukseklik
        If yukseklik <= 70 Then
            yukseklik = 70
            core_unit_settings_UI.Height = yukseklik
        End If
    ElseIf Rep <= 60 Then
        Rep = 60
        core_unit_settings_UI.Width = Rep
        yukseklik = yukseklik - 60
        core_unit_settings_UI.Height = yukseklik
        If yukseklik <= 60 Then
            yukseklik = 60
            core_unit_settings_UI.Height = yukseklik
        End If
    End If
Loop Until yukseklik = 60

Unload Me

End Sub

Sub timeout(duration_ms As Double)
Dim Start_Time As Variant

Start_Time = Timer
Do
DoEvents
Loop Until (Timer - Start_Time) >= duration_ms

End Sub


