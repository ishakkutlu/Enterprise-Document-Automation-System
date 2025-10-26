VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} support_report_templates_UI 
   Caption         =   "Enterprise Document Automation System"
   ClientHeight    =   10140
   ClientLeft      =   30
   ClientTop       =   375
   ClientWidth     =   15915
   OleObjectBlob   =   "support_report_templates_UI.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "support_report_templates_UI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Sub ColorChangerGenel()

'LabelEkle
If LabelEkle.BackColor <> RGB(225, 235, 245) Then
    LabelEkle.BackColor = RGB(225, 235, 245)
    LabelEkle.ForeColor = RGB(30, 30, 30)
End If
'LabelKaldir
If LabelKaldir.BackColor <> RGB(225, 235, 245) Then
    LabelKaldir.BackColor = RGB(225, 235, 245)
    LabelKaldir.ForeColor = RGB(30, 30, 30)
End If
'LabelGuncelle
If LabelGuncelle.BackColor <> RGB(225, 235, 245) Then
    LabelGuncelle.BackColor = RGB(225, 235, 245)
    LabelGuncelle.ForeColor = RGB(30, 30, 30)
End If
'LabelKapat
If LabelKapat.BackColor <> RGB(225, 235, 245) Then
    LabelKapat.BackColor = RGB(225, 235, 245)
    LabelKapat.ForeColor = RGB(30, 30, 30)
End If
''LabelYardim
If LabelYardim.BackColor <> RGB(225, 235, 245) Then
    LabelYardim.BackColor = RGB(225, 235, 245)
    LabelYardim.ForeColor = RGB(30, 30, 30)
End If

'LabelGetir
If LabelGetir.BackColor <> RGB(254, 254, 254) Then
LabelGetir.BackColor = RGB(254, 254, 254)
LabelGetir.ForeColor = RGB(30, 30, 30)
End If

End Sub

Private Sub ComboRaporTipi_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Me.ComboRaporTipi.DropDown
End Sub

Private Sub LabelEkle_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
LabelEkle.BackColor = RGB(60, 100, 180) 'RGB(28, 49, 68) 'RGB(60, 100, 180)
LabelEkle.ForeColor = RGB(255, 255, 255)
End Sub
Private Sub LabelKaldir_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
LabelKaldir.BackColor = RGB(60, 100, 180) 'RGB(60, 100, 180) 'RGB(60, 100, 180)
LabelKaldir.ForeColor = RGB(255, 255, 255)
End Sub
Private Sub LabelGuncelle_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
LabelGuncelle.BackColor = RGB(60, 100, 180) 'RGB(60, 100, 180) 'RGB(60, 100, 180)
LabelGuncelle.ForeColor = RGB(255, 255, 255)
End Sub
Private Sub LabelKapat_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
LabelKapat.BackColor = RGB(60, 100, 180) 'RGB(60, 100, 180) 'RGB(60, 100, 180)
LabelKapat.ForeColor = RGB(255, 255, 255)
End Sub
Private Sub LabelYardim_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
LabelYardim.BackColor = RGB(60, 100, 180) 'RGB(60, 100, 180) 'RGB(60, 100, 180)
LabelYardim.ForeColor = RGB(255, 255, 255)
End Sub


Private Sub LabelGetir_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
LabelGetir.BackColor = RGB(60, 100, 180) 'RGB(60, 100, 180) 'RGB(60, 100, 180)
LabelGetir.ForeColor = RGB(255, 255, 255)
End Sub


Private Sub ComboRaporTipi_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
Call SetComboBoxHook(ComboRaporTipi) 'Open scrollable with mouse
End Sub

Private Sub LblRaporTipi_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub
Private Sub TextIlk_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub
Private Sub Text1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub
Private Sub Text2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub
Private Sub Text3_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub
Private Sub Text4_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub
Private Sub Text5_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub
Private Sub Text6_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub
Private Sub TextSon_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
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

Private Sub LabelGetir_Click()
Dim AutoPath As String, DestTarget As String, OpenControl As String
Dim FileName As String, SayHedef As Long, ItemName As String, j As Integer
Dim HedefFile As String, fso As Object, TextLine As String
Dim NotIcerigi As String, MyRange As Object

Dim objWord As Object, objDoc As Object

Application.ScreenUpdating = False
Application.DisplayAlerts = False

     
'Öncekini temizle
TextIlk.Value = ""
For j = 1 To 6
    Controls("Text" & j).Value = ""
Next j
TextSon.Value = ""

AutoPath = ThisWorkbook.Path
DestTarget = AutoPath & "\System Files\System Templates\Report 2 Templates\"

If ComboRaporTipi.Value = "" Then
    MsgBox "Content cannot be retrieved because no report type was selected.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If


FileName = ComboRaporTipi.Value
HedefFile = DestTarget & FileName & ".docm"

'Dosyanın olup olmadığını kontrol et.
If Not Dir(HedefFile, vbDirectory) <> vbNullString Then
    MsgBox "The report type named " & FileName & " cannot be retrieved because it has not been created before.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If


'dosyayı aç
On Error Resume Next
Set objWord = GetObject(, "Word.Application")
Set objWord = GetObject(, "Word.Application")
Set objWord = GetObject(, "Word.Application")
Set objWord = GetObject(, "Word.Application")
Set objWord = GetObject(, "Word.Application")
If objWord Is Nothing Then
    Set objWord = CreateObject("Word.Application")
    objWord.Visible = False
End If
objWord.Documents.Open FileName:=HedefFile
'objWord.Visible = True
'objWord.Activate 'Ekrana getirir.
Set objDoc = GetObject(HedefFile)

If objDoc.Tables(3).Rows.Count > 0 Then
    TextIlk.Value = objDoc.Tables(3).Cell(Row:=2, Column:=1).Range.Text
    'Boşlukları ve boş satırları kaldır
    TextIlk.Value = Replace(Replace(TextIlk.Value, Chr(10), ""), Chr(13), "")
    TextLine = TextIlk.Value
    Do While Left(TextLine, 1) = " " ' Delete any excess spaces
        TextLine = Right(TextLine, Len(TextLine) - 1)
    Loop
    Do While Right(TextLine, 1) = " " ' Delete any excess spaces
        TextLine = Left(TextLine, Len(TextLine) - 1)
    Loop
    TextIlk.Value = TextLine
End If

If objDoc.Tables(4).Rows.Count > 0 Then
    For j = 1 To objDoc.Tables(4).Rows.Count
        Controls("Text" & j).Value = objDoc.Tables(4).Cell(Row:=j, Column:=2).Range.Text
        'Boşlukları ve boş satırları kaldır
        Controls("Text" & j).Value = Replace(Replace(Controls("Text" & j).Value, Chr(10), ""), Chr(13), "")
        TextLine = Controls("Text" & j).Value
        Do While Left(TextLine, 1) = " " ' Delete any excess spaces
            TextLine = Right(TextLine, Len(TextLine) - 1)
        Loop
        Do While Right(TextLine, 1) = " " ' Delete any excess spaces
            TextLine = Left(TextLine, Len(TextLine) - 1)
        Loop
        Controls("Text" & j).Value = TextLine
    Next j
End If

If objDoc.Tables(5).Rows.Count > 0 Then
    TextSon.Value = objDoc.Tables(5).Cell(Row:=1, Column:=1).Range.Text
    'Boşlukları ve boş satırları kaldır
    TextSon.Value = Replace(Replace(TextSon.Value, Chr(10), ""), Chr(13), "")
    TextLine = TextSon.Value
    Do While Left(TextLine, 1) = " " ' Delete any excess spaces
        TextLine = Right(TextLine, Len(TextLine) - 1)
    Loop
    Do While Right(TextLine, 1) = " " ' Delete any excess spaces
        TextLine = Left(TextLine, Len(TextLine) - 1)
    Loop
    TextSon.Value = TextLine
End If


objDoc.Close SaveChanges:=False

Son:

ThisWorkbook.Activate

Application.ScreenUpdating = True
Application.DisplayAlerts = True

End Sub

Private Sub LabelEkle_Click()
Dim AutoPath As String, DestTarget As String, OpenControl As String
Dim FileName As String, SayHedef As Long, ItemName As String, j As Integer
Dim HedefFile As String, fso As Object
Dim NotIcerigi As String, MyRange As Object
Dim a() As Variant, i As Variant, SourceFile As String, Cont As Integer
Dim objWord As Object, objDoc As Object

Application.ScreenUpdating = False
Application.DisplayAlerts = False

AutoPath = ThisWorkbook.Path
DestTarget = AutoPath & "\System Files\System Templates\Report 2 Templates\"

'Report type cannot be empty.
If ComboRaporTipi.Value = "" Then
    MsgBox "Template cannot be created because the report type is not named.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

''Invalid characters...
If InStr(ComboRaporTipi.Value, "/") > 0 Or InStr(ComboRaporTipi.Value, "\") > 0 Or InStr(ComboRaporTipi.Value, "<") > 0 Or InStr(ComboRaporTipi.Value, ">") > 0 Or _
   InStr(ComboRaporTipi.Value, ":") > 0 Or InStr(ComboRaporTipi.Value, "*") > 0 Or InStr(ComboRaporTipi.Value, "?") > 0 Or InStr(ComboRaporTipi.Value, "|") > 0 Or _
   InStr(ComboRaporTipi.Value, """") > 0 Or InStr(ComboRaporTipi.Value, "[") > 0 Or InStr(ComboRaporTipi.Value, "]") > 0 Or InStr(ComboRaporTipi.Value, "_") > 0 Or _
   InStr(ComboRaporTipi.Value, "(") > 0 Or InStr(ComboRaporTipi.Value, ")") > 0 Or InStr(ComboRaporTipi.Value, ".") > 0 Or InStr(ComboRaporTipi.Value, ",") > 0 Then
    MsgBox """ /, \, <, >, ], [, :, "", *, |, ?, _, (, ), ., , characters are reserved by the system, so your template cannot be created. Please do not use any of these characters when naming your report type template. You can use the hyphen (-) instead.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        ComboRaporTipi.Value = ""
    GoTo Son
End If

'Value already defined in ComboBox cannot be entered.
a() = ComboRaporTipi.List
For i = LBound(a) To UBound(a)
    If a(i, 0) = ComboRaporTipi.Value Then
        MsgBox "The name specified for the report type is already registered, so your template cannot be created.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
Next i

If InStr(TextIlk, "<item>") = 0 Then
    MsgBox "Template cannot be created because the <item> tag was not detected in the first line.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If


'ilk harfler büyük
ComboRaporTipi.Value = WorksheetFunction.Proper(ComboRaporTipi.Value)

FileName = ComboRaporTipi.Value
SourceFile = DestTarget & "Standard Invalid.docm"
'Check if the source file exists.
If Not Dir(SourceFile, vbDirectory) <> vbNullString Then
    MsgBox "The source file could not be found, so the operation cannot be completed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

'Dosyayı şablondan operasyon klasörüne kopyala ve adını değiştir.
Set fso = CreateObject("Scripting.FileSystemObject")
fso.CopyFile (SourceFile), DestTarget & FileName & ".docm", True

HedefFile = DestTarget & FileName & ".docm"

'dosyayı aç
On Error Resume Next
Set objWord = GetObject(, "Word.Application")
Set objWord = GetObject(, "Word.Application")
Set objWord = GetObject(, "Word.Application")
Set objWord = GetObject(, "Word.Application")
Set objWord = GetObject(, "Word.Application")
If objWord Is Nothing Then
    Set objWord = CreateObject("Word.Application")
    objWord.Visible = False
End If
objWord.Documents.Open FileName:=HedefFile
'objWord.Visible = True
'objWord.Activate 'Ekrana getirir.
Set objDoc = GetObject(HedefFile)

'İlk bölümü doldur
If TextIlk.Value <> "" Then
    objDoc.Tables(3).Cell(Row:=2, Column:=1).Range.Text = TextIlk.Value
End If

'Önce satırları sil ve teke düşür
For j = 1 To objDoc.Tables(4).Rows.Count - 1
    objDoc.Tables(4).Rows(2).Delete
Next j
'Satır bold olmasın
objDoc.Tables(4).Cell(Row:=1, Column:=2).Range.Font.Bold = False
'Sonra tabloya satır ekle
Cont = 0
For j = 1 To 6
    If Controls("Text" & j).Value <> "" Then
        Cont = Cont + 1
        If Cont > 1 Then
            With objDoc.Tables(4)
                .Rows.Add BeforeRow:=objDoc.Tables(4).Rows(1)
                .Rows(1).Height = 10
            End With
        End If
    End If
Next j
'Arayı doldur
If Cont > 0 Then
    For j = 1 To Cont
        objDoc.Tables(4).Cell(Row:=j, Column:=2).Range.Text = Controls("Text" & j).Value
    Next j
End If
'Son bölümü doldur
If TextSon.Value <> "" Then
    objDoc.Tables(5).Cell(Row:=1, Column:=1).Range.Text = TextSon.Value
End If

'ANAHTAR KELİMELERİ BOLD YAP
'invalid (ilk satır)
Set MyRange = objDoc.Tables(3).Cell(Row:=2, Column:=1).Range
MyRange.Find.Execute FindText:="invalid"
If InStr(MyRange.Text, "invalid") <> 0 Or InStr(MyRange.Text, "invalid") <> 0 Then
    MyRange.Font.Bold = True
End If
'valid (ilk satır)
Set MyRange = objDoc.Tables(3).Cell(Row:=2, Column:=1).Range
MyRange.Find.Execute FindText:="valid"
If InStr(MyRange.Text, "valid") <> 0 Or InStr(MyRange.Text, "valid") <> 0 Then
    MyRange.Font.Bold = True
End If
'invalid (1. satır)
Set MyRange = objDoc.Tables(4).Cell(Row:=1, Column:=2).Range
MyRange.Find.Execute FindText:="invalid"
If InStr(MyRange.Text, "invalid") <> 0 Or InStr(MyRange.Text, "invalid") <> 0 Then
    MyRange.Font.Bold = True
End If
'valid (1. satır)
Set MyRange = objDoc.Tables(4).Cell(Row:=1, Column:=2).Range
MyRange.Find.Execute FindText:="valid"
If InStr(MyRange.Text, "valid") <> 0 Or InStr(MyRange.Text, "valid") <> 0 Then
    MyRange.Font.Bold = True
End If

'Kullanılmayan satırların içeriğini boşalt.
If Text1.Value = "" Then
    objDoc.Tables(4).Cell(Row:=1, Column:=1).Range.Text = ""
    objDoc.Tables(4).Cell(Row:=1, Column:=1).Range.ListFormat.RemoveNumbers
    objDoc.Tables(4).Cell(Row:=1, Column:=2).Range.Text = ""
End If
If TextSon.Value = "" Then
    objDoc.Tables(5).Cell(Row:=1, Column:=1).Range.Text = ""
End If

'Öğeyi kaydet
objDoc.Close SaveChanges:=True

'Yeni öğeyi comboya tanıt
SayHedef = ThisWorkbook.Worksheets(2).Range("DL1000").End(xlUp).Row
If SayHedef < 6 Then
    SayHedef = 6
End If
If SayHedef > 304 Then
    Kill (HedefFile)
    MsgBox "The dropdown list for report types is full, so your template named " & FileName & " could not be defined.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If
'Arada boş satır varsa onu bul ve öğeyi boş satıra yaz.
If SayHedef > 6 Then
    For j = 6 To SayHedef
        If ThisWorkbook.Worksheets(2).Cells(j, 116).Value = "" Then
            SayHedef = j - 1
            GoTo DonguSon
        End If
    Next j
End If
DonguSon:
ThisWorkbook.Worksheets(2).Cells(SayHedef + 1, 116).Value = FileName

MsgBox "Your report template named " & FileName & " has been successfully created.", vbOKOnly + vbInformation, "Enterprise Document Automation System"

Son:

ThisWorkbook.Activate

Application.ScreenUpdating = True
Application.DisplayAlerts = True

End Sub

Private Sub LabelKaldir_Click()
Dim AutoPath As String, DestTarget As String, OpenControl As String
Dim FileName As String, SayHedef As Long, ItemName As String, j As Integer
Dim HedefFile As String, fso As Object
Dim NotIcerigi As String, MyRange As Object
Dim a() As Variant, i As Variant, SourceFile As String, Bilgi As Variant
Dim objWord As Object, objDoc As Object

Application.ScreenUpdating = False
Application.DisplayAlerts = False

AutoPath = ThisWorkbook.Path
DestTarget = AutoPath & "\System Files\System Templates\Report 2 Templates\"
    
'Rapor tipi boş olamaz.
If ComboRaporTipi.Value = "" Then
    MsgBox "Delete operation cannot be started because no report type was selected.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

FileName = ComboRaporTipi.Value
HedefFile = DestTarget & FileName & ".docm"
If FileName = "Standard Valid" Or FileName = "Standard Invalid" Then
    MsgBox "The file named " & FileName & " is a system file and therefore cannot be deleted.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

Bilgi = MsgBox("Your report template named " & FileName & " will be deleted. Click " & """" & "Yes" & """" & " to proceed or " & """" & "No" & """" & " to cancel.", vbYesNo + vbExclamation, "Enterprise Document Automation System")
If Bilgi = vbNo Then
    GoTo Son
End If

'Check if the target file exists.
If Not Dir(HedefFile, vbDirectory) <> vbNullString Then
    MsgBox "The target file could not be found, so the operation cannot be completed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
Else
    Kill (HedefFile)
End If


SayHedef = ThisWorkbook.Worksheets(2).Range("DL1000").End(xlUp).Row
If SayHedef < 6 Then
    SayHedef = 6
End If
If SayHedef > 304 Then
    Kill (HedefFile)
    MsgBox "The dropdown list for report types is full, so your template named " & FileName & " could not be defined.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

'Öğeyi skp'de bul ve sil.
If SayHedef > 6 Then
    For j = 6 To SayHedef
        If ThisWorkbook.Worksheets(2).Cells(j, 116).Value = FileName Then
            ThisWorkbook.Worksheets(2).Cells(j, 116).Value = ""
        End If
    Next j
End If
DonguSon:

TextIlk.Value = ""
Text1.Value = ""
Text2.Value = ""
Text3.Value = ""
Text4.Value = ""
Text5.Value = ""
Text6.Value = ""
TextSon.Value = ""


MsgBox "Your report template named " & FileName & " has been successfully removed.", vbOKOnly + vbInformation, "Enterprise Document Automation System"

Son:

ThisWorkbook.Activate

Application.ScreenUpdating = True
Application.DisplayAlerts = True

End Sub

Private Sub LabelGuncelle_Click()
Dim AutoPath As String, DestTarget As String, OpenControl As String
Dim FileName As String, SayHedef As Long, ItemName As String, j As Integer
Dim HedefFile As String, fso As Object
Dim NotIcerigi As String, MyRange As Object
Dim objWord As Object, objDoc As Object
Dim a() As Variant, i As Variant, SourceFile As String, Cont As Integer


Application.ScreenUpdating = False
Application.DisplayAlerts = False

AutoPath = ThisWorkbook.Path
DestTarget = AutoPath & "\System Files\System Templates\Report 2 Templates\"

'Report type cannot be empty.
If ComboRaporTipi.Value = "" Then
    MsgBox "Update cannot be performed because no report type was selected.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

''Invalid characters...
If InStr(ComboRaporTipi.Value, "/") > 0 Or InStr(ComboRaporTipi.Value, "\") > 0 Or InStr(ComboRaporTipi.Value, "<") > 0 Or InStr(ComboRaporTipi.Value, ">") > 0 Or _
   InStr(ComboRaporTipi.Value, ":") > 0 Or InStr(ComboRaporTipi.Value, "*") > 0 Or InStr(ComboRaporTipi.Value, "?") > 0 Or InStr(ComboRaporTipi.Value, "|") > 0 Or _
   InStr(ComboRaporTipi.Value, """") > 0 Or InStr(ComboRaporTipi.Value, "[") > 0 Or InStr(ComboRaporTipi.Value, "]") > 0 Or InStr(ComboRaporTipi.Value, "_") > 0 Or _
   InStr(ComboRaporTipi.Value, "(") > 0 Or InStr(ComboRaporTipi.Value, ")") > 0 Or InStr(ComboRaporTipi.Value, ".") > 0 Or InStr(ComboRaporTipi.Value, ",") > 0 Then
    MsgBox """ /, \, <, >, ], [, :, "", *, |, ?, _, (, ), ., , characters are reserved by the system, so your template cannot be created. Please do not use any of these characters when naming your report type template. You can use the hyphen (-) instead.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    ComboRaporTipi.Value = ""
    GoTo Son
End If

'Comboda tanımlı değer girilmeli.(DÜZELT!!)
'a() = ComboRaporTipi.List
'For i = LBound(a) To UBound(a)
'    If a(i, 0) = ComboRaporTipi.Value Then
'        MsgBox "Rapor tipinde belirtilen isimde bir şablon bulunamadığından güncellene gerçekleştirilemiyor.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
'        GoTo Son
'    End If
'Next i

'<item> kontrolü
If InStr(TextIlk, "<item>") = 0 Then
    MsgBox "The <item> tag was not detected in the first line, so your template cannot be updated.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If


'ilk harfler büyük
ComboRaporTipi.Value = WorksheetFunction.Proper(ComboRaporTipi.Value)

FileName = ComboRaporTipi.Value
HedefFile = DestTarget & FileName & ".docm"
If FileName = "Standard Valid" Or FileName = "Standard Invalid" Then
    MsgBox "The file named " & FileName & " is a system file and therefore the update operation cannot be completed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

'Check if the file exists.
If Not Dir(HedefFile, vbDirectory) <> vbNullString Then
    MsgBox "The report type named " & FileName & " has not been created before, so the update cannot be performed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If


'Önceki dosyayı sil
Kill (HedefFile)

SourceFile = DestTarget & "Standard Invalid.docm"
'Kaynak dosyanın olup olmadığını kontrol et.
If Not Dir(SourceFile, vbDirectory) <> vbNullString Then
    MsgBox "The source file could not be found, so the operation cannot be completed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

'Dosyayı şablondan operasyon klasörüne kopyala ve adını değiştir.
Set fso = CreateObject("Scripting.FileSystemObject")
fso.CopyFile (SourceFile), DestTarget & FileName & ".docm", True

HedefFile = DestTarget & FileName & ".docm"

'dosyayı aç
On Error Resume Next
Set objWord = GetObject(, "Word.Application")
Set objWord = GetObject(, "Word.Application")
Set objWord = GetObject(, "Word.Application")
Set objWord = GetObject(, "Word.Application")
Set objWord = GetObject(, "Word.Application")
If objWord Is Nothing Then
    Set objWord = CreateObject("Word.Application")
    objWord.Visible = False
End If
objWord.Documents.Open FileName:=HedefFile
'objWord.Visible = True
'objWord.Activate 'Ekrana getirir.
Set objDoc = GetObject(HedefFile)

'İlk bölümü doldur
If TextIlk.Value <> "" Then
    objDoc.Tables(3).Cell(Row:=2, Column:=1).Range.Text = TextIlk.Value
End If

'Önce satırları sil ve teke düşür
For j = 1 To objDoc.Tables(4).Rows.Count - 1
    objDoc.Tables(4).Rows(2).Delete
Next j
'Satır bold olmasın
objDoc.Tables(4).Cell(Row:=1, Column:=2).Range.Font.Bold = False
'Sonra tabloya satır ekle
Cont = 0
For j = 1 To 6
    If Controls("Text" & j).Value <> "" Then
        Cont = Cont + 1
        If Cont > 1 Then
            With objDoc.Tables(4)
                .Rows.Add BeforeRow:=objDoc.Tables(4).Rows(1)
                .Rows(1).Height = 10
            End With
        End If
    End If
Next j
'Arayı doldur
If Cont > 0 Then
    For j = 1 To Cont
        objDoc.Tables(4).Cell(Row:=j, Column:=2).Range.Text = Controls("Text" & j).Value
    Next j
End If
'Son bölümü doldur
If TextSon.Value <> "" Then
    objDoc.Tables(5).Cell(Row:=1, Column:=1).Range.Text = TextSon.Value
End If

'ANAHTAR KELİMELERİ BOLD YAP
'invalid (ilk satır)
Set MyRange = objDoc.Tables(3).Cell(Row:=2, Column:=1).Range
MyRange.Find.Execute FindText:="invalid"
If InStr(MyRange.Text, "invalid") <> 0 Or InStr(MyRange.Text, "invalid") <> 0 Then
    MyRange.Font.Bold = True
End If
'valid (ilk satır)
Set MyRange = objDoc.Tables(3).Cell(Row:=2, Column:=1).Range
MyRange.Find.Execute FindText:="valid"
If InStr(MyRange.Text, "valid") <> 0 Or InStr(MyRange.Text, "valid") <> 0 Then
    MyRange.Font.Bold = True
End If
'invalid (1. satır)
Set MyRange = objDoc.Tables(4).Cell(Row:=1, Column:=2).Range
MyRange.Find.Execute FindText:="invalid"
If InStr(MyRange.Text, "invalid") <> 0 Or InStr(MyRange.Text, "invalid") <> 0 Then
    MyRange.Font.Bold = True
End If
'valid (1. satır)
Set MyRange = objDoc.Tables(4).Cell(Row:=1, Column:=2).Range
MyRange.Find.Execute FindText:="valid"
If InStr(MyRange.Text, "valid") <> 0 Or InStr(MyRange.Text, "valid") <> 0 Then
    MyRange.Font.Bold = True
End If

'Kullanılmayan satırların içeriğini boşalt.
If Text1.Value = "" Then
    objDoc.Tables(4).Cell(Row:=1, Column:=1).Range.Text = ""
    objDoc.Tables(4).Cell(Row:=1, Column:=1).Range.ListFormat.RemoveNumbers
    objDoc.Tables(4).Cell(Row:=1, Column:=2).Range.Text = ""
End If
If TextSon.Value = "" Then
    objDoc.Tables(5).Cell(Row:=1, Column:=1).Range.Text = ""
End If

objDoc.Close SaveChanges:=True

MsgBox "Your report template named " & FileName & " has been successfully updated.", vbOKOnly + vbInformation, "Enterprise Document Automation System"


Son:

ThisWorkbook.Activate

Application.ScreenUpdating = True
Application.DisplayAlerts = True

End Sub
Private Sub LabelYardim_Click()
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
SourceTaslak = AutoPath & "\System Files\Help Documents\Report Templates Manager – Help.docm"
'Birden fazla kullanıcının operasyon klasöründe çakışmasını önlemek için operasyon klasörünün içinde yeni klasör aç.
DestOpUserFolderName = "Operation-" & Left(ThisWorkbook.name, InStr(ThisWorkbook.name, ".") - 1) '& " " & Format(Now(), "ddmmyyyyhhmmss")
DestOpUserFolder = DestOperasyon & DestOpUserFolderName & "\"

'System Files klasör adını kontrol et.
If Not Dir(AutoPath & "\System Files\", vbDirectory) <> vbNullString Then
    MsgBox "Cannot access the directory " & AutoPath & "\System Files\" & ". The folder named 'System Files' may have been renamed or deleted.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

'Check the operation folder name.
If Not Dir(DestOperasyon, vbDirectory) <> vbNullString Then
    MsgBox "Cannot access the directory " & DestOperasyon & ". The folder named 'Operation' may have been renamed or deleted.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If


'RmDir DestOpUserFolder 'Sistem kapanırken DestOpUserFolder klasörünü temizle EKLENECEK!
'_______________

'Klasör isimlerini kontrol et.
If Not Dir(SourceTaslak, vbDirectory) <> vbNullString Then
    MsgBox "Cannot access the directory " & SourceTaslak & ". The names of folders and/or files in this directory may have been changed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
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


Private Sub LabelKapat_Click()
    Unload Me
End Sub


Private Sub UserForm_Initialize()
Dim i As Long, WsSKP As Object
Dim ClrLab As MSForms.Control

ThisWorkbook.Activate

For Each ClrLab In support_report_templates_UI.Controls
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

LabelEkle.BackColor = RGB(225, 235, 245)
LabelEkle.ForeColor = RGB(30, 30, 30)
LabelGuncelle.BackColor = RGB(225, 235, 245)
LabelGuncelle.ForeColor = RGB(30, 30, 30)
LabelKapat.BackColor = RGB(225, 235, 245)
LabelKapat.ForeColor = RGB(30, 30, 30)
LabelYardim.BackColor = RGB(225, 235, 245)
LabelYardim.ForeColor = RGB(30, 30, 30)
LabelKaldir.BackColor = RGB(225, 235, 245)
LabelKaldir.ForeColor = RGB(30, 30, 30)

support_report_templates_UI.BackColor = RGB(230, 230, 230) 'YENİ

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
Dim yukseklik As Variant, genislik As Variant
Dim Rep As Variant

yukseklik = Me.Height
Rep = Me.Width
Do
DoEvents
Rep = Rep - 90
Call timeout(0.01)
    If Rep > 90 Then
        support_report_templates_UI.Width = Rep
        yukseklik = yukseklik - 90
        support_report_templates_UI.Height = yukseklik
        If yukseklik <= 90 Then
            yukseklik = 90
            support_report_templates_UI.Height = yukseklik
        End If
    ElseIf Rep <= 50 Then
        Rep = 50
        support_report_templates_UI.Width = Rep
        yukseklik = yukseklik - 50
        support_report_templates_UI.Height = yukseklik
        If yukseklik <= 50 Then
            yukseklik = 50
            support_report_templates_UI.Height = yukseklik
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



