VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} support_contact_themes_UI 
   Caption         =   "Enterprise Document Automation System"
   ClientHeight    =   8895
   ClientLeft      =   30
   ClientTop       =   375
   ClientWidth     =   13920
   OleObjectBlob   =   "support_contact_themes_UI.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "support_contact_themes_UI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text


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
'LabelKapat
If LabelKapat.BackColor <> RGB(225, 235, 245) Then
    LabelKapat.BackColor = RGB(225, 235, 245)
    LabelKapat.ForeColor = RGB(30, 30, 30)
End If

'CheckBoxDuzelt
If CheckBoxDuzelt.BackColor <> RGB(254, 254, 254) Then
    CheckBoxDuzelt.BackColor = RGB(254, 254, 254)
    CheckBoxDuzelt.ForeColor = RGB(70, 70, 70)
End If


End Sub

Private Sub ComboGelenGidenMuhatapTemasi_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Me.ComboGelenGidenMuhatapTemasi.DropDown
End Sub

Private Sub LabelEkle_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
LabelEkle.BackColor = RGB(60, 100, 180) 'RGB(60, 100, 180) 'RGB(60, 100, 180)
LabelEkle.ForeColor = RGB(255, 255, 255)
End Sub

Private Sub LabelKaldir_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
LabelKaldir.BackColor = RGB(60, 100, 180) 'RGB(60, 100, 180) 'RGB(60, 100, 180)
LabelKaldir.ForeColor = RGB(255, 255, 255)
End Sub

Private Sub LabelKapat_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
LabelKapat.BackColor = RGB(60, 100, 180) 'RGB(60, 100, 180) 'RGB(60, 100, 180)
LabelKapat.ForeColor = RGB(255, 255, 255)
End Sub

Private Sub ComboGelenGidenMuhatapTemasi_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
Call SetComboBoxHook(ComboGelenGidenMuhatapTemasi) 'Open scrollable with mouse
End Sub

Private Sub CheckBoxDuzelt_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
CheckBoxDuzelt.BackColor = RGB(60, 100, 180)
CheckBoxDuzelt.ForeColor = RGB(255, 255, 255)
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


Private Sub LabelEkle_Click()
Dim a() As Variant, i As Variant
Dim AutoPath As String, DestTarget As String, OpenControl As String
Dim FileName As String, SayHedef As Long, ItemName As String, j As Integer
Dim TC As String, ItemNameBuyuk As String

Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.EnableEvents = False

AutoPath = ThisWorkbook.Path
DestTarget = AutoPath & "\System Files\System Definitions\"
FileName = "Definitions.xlsx"
ItemName = ComboGelenGidenMuhatapTemasi.Value

If ItemName <> "" Then
    'Comboya tanımlı değer girilemez.(Rezerv tanımları için)
    a() = ComboGelenGidenMuhatapTemasi.List
    For i = LBound(a) To UBound(a)
        If a(i, 0) = ItemName Then
            MsgBox "The incoming/outgoing contact theme named " & ItemName & " has already been defined for the related dropdown lists, therefore the operation could not be completed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
            GoTo Son
        End If
    Next i
    
    'Birden fazla boşluk varsa kaldır
    For i = 1 To 50
        ItemName = Replace(ItemName, "  ", " ")
    Next i
    'Sağdaki ve soldaki tek boşluğu kaldır
    Do While Left(ItemName, 1) = " "
        ItemName = Right(ItemName, Len(ItemName) - 1)
    Loop
    Do While Right(ItemName, 1) = " "
        ItemName = Left(ItemName, Len(ItemName) - 1)
    Loop
    
    'Harfleri büyüt
    'ItemNameBuyuk = UCase(ItemName)
    
    'Directorate, Arbitration, Decision Board kelimelerini kontrol et
    If InStr(ItemName, "Directorate") > 0 Or InStr(ItemName, "Arbitration") > 0 Or InStr(ItemName, "Decision Board") > 0 Or _
        InStr(ItemName, "DIRECTORATE") > 0 Or InStr(ItemName, "ARBITRATION") > 0 Or InStr(ItemName, "DECISION BOARD") > 0 Then
        '
    ElseIf InStr(ItemName, "General Directorate") > 0 Or InStr(ItemName, "Regional Directorate") > 0 Or _
        InStr(ItemName, "GENERAL DIRECTORATE") > 0 Or InStr(ItemName, "REGIONAL DIRECTORATE") > 0 Then
        '
    Else
        MsgBox "The contact theme you are trying to add is not one of the following: Directorate, Decision Board, or Arbitration. Therefore, the operation could not be completed. Only Directorate, Decision Board, or Arbitration units can be defined in this field.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    
    
    'X.X.'den sonraki ilk harfler büyük
    'ItemName = WorksheetFunction.Proper(Mid(ItemName, 5, Len(ItemName) - 4))
    'ItemName = TC & ItemName
    If CheckBoxDuzelt.Value = True Then

        If Left(ItemName, 4) = "X.X." Then
            ItemName = WorksheetFunction.Proper(Mid(ItemName, 5, Len(ItemName) - 4))
            ItemName = "X.X. " & ItemName
        ElseIf Left(ItemName, 3) = "X.X" Or Left(ItemName, 3) = "TC." Then
            ItemName = WorksheetFunction.Proper(Mid(ItemName, 4, Len(ItemName) - 3))
            ItemName = "X.X. " & ItemName
        ElseIf Left(ItemName, 2) = "XX" Then
            ItemName = WorksheetFunction.Proper(Mid(ItemName, 3, Len(ItemName) - 2))
            ItemName = "X.X. " & ItemName
        Else
            ItemName = WorksheetFunction.Proper(ItemName)
        End If
        'Yukarıdaki düzeltme sonrası boşluk oluşursa kaldır.
        For i = 1 To 50
            ItemName = Replace(ItemName, "  ", " ")
        Next i
        'Sağdaki ve soldaki tek boşluğu kaldır
        Do While Left(ItemName, 1) = " "
            ItemName = Right(ItemName, Len(ItemName) - 1)
        Loop
        Do While Right(ItemName, 1) = " "
            ItemName = Left(ItemName, Len(ItemName) - 1)
        Loop
        
        'Ve kelimelerini düzelt
        If InStr(ItemName, " And ") <> 0 Then
            ItemName = Replace(ItemName, " And ", " and ")
        End If
    End If

'    'XX ile devamı arasında birden fazla boşluk varsa teke düşür
'    For i = 1 To 5
'        ItemName = Replace(ItemName, "  ", " ")
'    Next i
    'Comboya tanımlı değer girilemez.
    a() = ComboGelenGidenMuhatapTemasi.List
    For i = LBound(a) To UBound(a)
        If a(i, 0) = ItemName Then
            MsgBox "The contact theme information named " & ItemName & " has already been defined in the relevant drop-down lists, so the operation could not be completed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
            GoTo Son
        End If
    Next i
Else
    GoTo Son
End If

OpenControl = IsFileOpen(DestTarget & FileName)
If OpenControl = True Then 'Açıksa
    Workbooks(FileName).Close SaveChanges:=True
ElseIf OpenControl = False Then
    '
End If
Workbooks.Open (DestTarget & FileName)
Workbooks(FileName).Worksheets(1).Activate

SayHedef = Workbooks(FileName).Worksheets(1).Range("CV1000").End(xlUp).Row
If SayHedef < 6 Then
    SayHedef = 6
End If
If SayHedef > 304 Then
    MsgBox "The drop-down definition area for the incoming/outgoing contact theme selection is full, so the sender/recipient unit information named " & ItemName & " could not be defined.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    Workbooks(FileName).Close SaveChanges:=True
    GoTo Son
End If

'Arada boş satır varsa onu bul ve öğeyi boş satıra yaz.
If SayHedef > 6 Then
    For j = 6 To SayHedef
        If Workbooks(FileName).Worksheets(1).Cells(j, 100).Value = "" Then
            SayHedef = j - 1
            GoTo DonguSon
        End If
    Next j
End If
DonguSon:
''Ve kelimelerini düzelt
'If InStr(ItemName, " And ") <> 0 Then
'    ItemName = Replace(ItemName, " And ", " and ")
'End If

Workbooks(FileName).Worksheets(1).Cells(SayHedef + 1, 100).Value = ItemName
ThisWorkbook.Worksheets(2).Cells(SayHedef + 1, 100).Value = ItemName

'Sıralama ekle komutuna da uygulanacak.
SayHedef = Workbooks(FileName).Worksheets(1).Range("CV1000").End(xlUp).Row
If SayHedef > 14 Then
    'A'dan Z'ye sırala ve böylece arada bulunan boş satırları da kaldır.
    Workbooks(FileName).Worksheets(1).Unprotect Password:="123"
    ThisWorkbook.Unprotect "123"
    ThisWorkbook.Worksheets(2).Unprotect Password:="123"
    ThisWorkbook.Worksheets(2).Visible = True
    'Sort A to Z
    Workbooks(FileName).Worksheets(1).Range("CV" & 14 & ":CV" & SayHedef).Sort key1:=Workbooks(FileName).Worksheets(1).Range("CV" & 14), order1:=xlAscending, Header:=xlNo
    ThisWorkbook.Worksheets(2).Range("CV" & 14 & ":CV" & SayHedef).Sort key1:=ThisWorkbook.Worksheets(2).Range("CV" & 14), order1:=xlAscending, Header:=xlNo
    ThisWorkbook.Worksheets(2).Visible = False
    Workbooks(FileName).Worksheets(1).Protect Password:="123" ', DrawingObjects:=False
    ThisWorkbook.Worksheets(2).Protect Password:="123" ', DrawingObjects:=False
    ThisWorkbook.Protect "123"
End If

ComboGelenGidenMuhatapTemasi.Value = ""
Workbooks(FileName).Save
'ThisWorkbook.Save
MsgBox "The incoming/outgoing contact theme information named " & ItemName & " has been successfully added to the relevant drop-down lists.", vbOKOnly + vbInformation, "Enterprise Document Automation System"

Son:

OpenControl = IsFileOpen(DestTarget & FileName)
If OpenControl = True Then 'Açıksa
    Workbooks(FileName).Close SaveChanges:=True
ElseIf OpenControl = False Then
    '
End If

ThisWorkbook.Activate

'Sadece RAPOR FORMU için
'İlgili birim combobox list (GelenMuhatapTemasi'nin sadece Directoratelık ve Karar Kurulları)
core_report2_entry_UI.IlgiliBirim.Clear
a() = core_report2_entry_UI.GelenMuhatapTemasi.List
j = 0
For i = LBound(a) To UBound(a)
    If j > 7 Then
        core_report2_entry_UI.IlgiliBirim.AddItem (a(i, 0))
    End If
    j = j + 1
Next i

Application.ScreenUpdating = True
Application.DisplayAlerts = True
Application.EnableEvents = True

End Sub

Private Sub LabelKaldir_Click()
Dim a() As Variant, i As Variant
Dim AutoPath As String, DestTarget As String, OpenControl As String, ItemName As String
Dim FileName As String, ListControl As Integer, ItemBul As Range, SayHedef As Long

Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.EnableEvents = False

AutoPath = ThisWorkbook.Path
DestTarget = AutoPath & "\System Files\System Definitions\"
FileName = "Definitions.xlsx"
ItemName = ComboGelenGidenMuhatapTemasi.Value

If ItemName <> "" Then
    'Comboya tanımlı değer girilmelidir.
    ListControl = 0
    a() = ComboGelenGidenMuhatapTemasi.List
    For i = LBound(a) To UBound(a)
        If a(i, 0) = ItemName Then
            ListControl = 1
        End If
    Next i
Else
    GoTo Son
End If

If ListControl = 0 Then
    MsgBox "The incoming/outgoing contact theme information named " & ItemName & " could not be removed because it has not been previously defined in the relevant drop-down lists.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

OpenControl = IsFileOpen(DestTarget & FileName)
If OpenControl = True Then 'Açıksa
    Workbooks(FileName).Close SaveChanges:=True
ElseIf OpenControl = False Then
    '
End If
Workbooks.Open (DestTarget & FileName)
Workbooks(FileName).Worksheets(1).Activate

Set ItemBul = Workbooks(FileName).Worksheets(1).Range("CV6:CV13").Find(What:=ItemName, SearchDirection:=xlNext, _
                SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
If Not ItemBul Is Nothing Then
    MsgBox "The incoming/outgoing contact theme named " & ItemName & " could not be removed because it is used by the system. The system does not allow modifications to the first 8 themes in the drop-down list.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If


Set ItemBul = Workbooks(FileName).Worksheets(1).Range("CV14:CV1000").Find(What:=ItemName, SearchDirection:=xlNext, _
                SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
If Not ItemBul Is Nothing Then
    '
Else
    Workbooks(FileName).Close SaveChanges:=True
    GoTo Son
End If

Workbooks(FileName).Worksheets(1).Cells(ItemBul.Row, 100).Value = ""
ThisWorkbook.Worksheets(2).Cells(ItemBul.Row, 100).Value = ""

'Sıralama ekle komutuna da uygulanacak.
SayHedef = Workbooks(FileName).Worksheets(1).Range("CV1000").End(xlUp).Row
If SayHedef > 14 Then
    'A'dan Z'ye sırala ve böylece arada bulunan boş satırları da kaldır.
    Workbooks(FileName).Worksheets(1).Unprotect Password:="123"
    ThisWorkbook.Unprotect "123"
    ThisWorkbook.Worksheets(2).Unprotect Password:="123"
    ThisWorkbook.Worksheets(2).Visible = True
    'Sort A to Z
    Workbooks(FileName).Worksheets(1).Range("CV" & 14 & ":CV" & SayHedef).Sort key1:=Workbooks(FileName).Worksheets(1).Range("CV" & 14), order1:=xlAscending, Header:=xlNo
    ThisWorkbook.Worksheets(2).Range("CV" & 14 & ":CV" & SayHedef).Sort key1:=ThisWorkbook.Worksheets(2).Range("CV" & 14), order1:=xlAscending, Header:=xlNo
    ThisWorkbook.Worksheets(2).Visible = False
    Workbooks(FileName).Worksheets(1).Protect Password:="123" ', DrawingObjects:=False
    ThisWorkbook.Worksheets(2).Protect Password:="123" ', DrawingObjects:=False
    ThisWorkbook.Protect "123"
End If

ComboGelenGidenMuhatapTemasi.Value = ""
Workbooks(FileName).Save
'ThisWorkbook.Save
MsgBox "The incoming/outgoing contact theme named " & ItemName & " has been successfully removed from the relevant drop-down lists.", vbOKOnly + vbInformation, "Enterprise Document Automation System"

Son:

OpenControl = IsFileOpen(DestTarget & FileName)
If OpenControl = True Then 'Açıksa
    Workbooks(FileName).Close SaveChanges:=True
ElseIf OpenControl = False Then
    '
End If

ThisWorkbook.Activate

Application.ScreenUpdating = True
Application.DisplayAlerts = True
Application.EnableEvents = True

End Sub

Private Sub LabelKapat_Click()
    Unload Me
End Sub


Private Sub UserForm_Initialize()
Dim i As Long, WsSKP As Object
Dim ClrLab As MSForms.Control

ThisWorkbook.Activate

For Each ClrLab In support_contact_themes_UI.Controls
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
LabelKapat.BackColor = RGB(225, 235, 245)
LabelKapat.ForeColor = RGB(30, 30, 30)
LabelKaldir.BackColor = RGB(225, 235, 245)
LabelKaldir.ForeColor = RGB(30, 30, 30)

support_contact_themes_UI.BackColor = RGB(230, 230, 230) 'YENİ

CheckBoxDuzelt.Value = True

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
        support_contact_themes_UI.Width = Rep
        yukseklik = yukseklik - 60
        support_contact_themes_UI.Height = yukseklik
        If yukseklik <= 60 Then
            yukseklik = 60
            support_contact_themes_UI.Height = yukseklik
        End If
    ElseIf Rep <= 50 Then
        Rep = 50
        support_contact_themes_UI.Width = Rep
        yukseklik = yukseklik - 50
        support_contact_themes_UI.Height = yukseklik
        If yukseklik <= 50 Then
            yukseklik = 50
            support_contact_themes_UI.Height = yukseklik
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


