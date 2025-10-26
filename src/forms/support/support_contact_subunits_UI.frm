VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} support_contact_subunits_UI 
   Caption         =   "Enterprise Document Automation System"
   ClientHeight    =   8895
   ClientLeft      =   30
   ClientTop       =   375
   ClientWidth     =   13920
   OleObjectBlob   =   "support_contact_subunits_UI.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "support_contact_subunits_UI"
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

Private Sub ComboGonderenGonderilen_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Me.ComboGonderenGonderilen.DropDown
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

Private Sub ComboGonderenGonderilen_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
Call SetComboBoxHook(ComboGonderenGonderilen) 'Open scrollable with mouse
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

Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.EnableEvents = False

AutoPath = ThisWorkbook.Path
DestTarget = AutoPath & "\System Files\System Definitions\"
FileName = "Definitions.xlsx"
ItemName = ComboGonderenGonderilen.Value

If ItemName <> "" Then
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
    
    If CheckBoxDuzelt.Value = True Then
        'İlk harfler büyük
        ItemName = WorksheetFunction.Proper(ItemName)
        'Ve kelimelerini düzelt
        If InStr(ItemName, " And ") <> 0 Then
            ItemName = Replace(ItemName, " And ", " and ")
        End If
    End If
    
    'Comboya tanımlı değer girilemez.
    a() = ComboGonderenGonderilen.List
    For i = LBound(a) To UBound(a)
        If a(i, 0) = ItemName Then
            MsgBox "The sender/recipient subunit information named " & ItemName & " has already been defined in the relevant dropdown lists, so the operation could not be completed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
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

SayHedef = Workbooks(FileName).Worksheets(1).Range("CW1000").End(xlUp).Row
If SayHedef < 8 Then
    SayHedef = 8
End If
If SayHedef > 304 Then
    MsgBox "The dropdown definition area for selecting the sender/recipient subunit is full, so the sender/recipient subunit information named " & ItemName & " could not be defined.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    Workbooks(FileName).Close SaveChanges:=True
    GoTo Son
End If

'Arada boş satır varsa onu bul ve öğeyi boş satıra yaz.
If SayHedef > 8 Then
    For j = 8 To SayHedef
        If Workbooks(FileName).Worksheets(1).Cells(j, 101).Value = "" Then
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

Workbooks(FileName).Worksheets(1).Cells(SayHedef + 1, 101).Value = ItemName
ThisWorkbook.Worksheets(2).Cells(SayHedef + 1, 101).Value = ItemName


'Sıralama ekle komutuna da uygulanacak.
SayHedef = Workbooks(FileName).Worksheets(1).Range("CW1000").End(xlUp).Row
If SayHedef > 8 Then
    'A'dan Z'ye sırala ve böylece arada bulunan boş satırları da kaldır.
    Workbooks(FileName).Worksheets(1).Unprotect Password:="123"
    ThisWorkbook.Unprotect "123"
    ThisWorkbook.Worksheets(2).Unprotect Password:="123"
    ThisWorkbook.Worksheets(2).Visible = True
    'Sort A to Z
    Workbooks(FileName).Worksheets(1).Range("CW" & 8 & ":CW" & SayHedef).Sort key1:=Workbooks(FileName).Worksheets(1).Range("CW" & 8), order1:=xlAscending, Header:=xlNo
    ThisWorkbook.Worksheets(2).Range("CW" & 8 & ":CW" & SayHedef).Sort key1:=ThisWorkbook.Worksheets(2).Range("CW" & 8), order1:=xlAscending, Header:=xlNo
    ThisWorkbook.Worksheets(2).Visible = False
    Workbooks(FileName).Worksheets(1).Protect Password:="123" ', DrawingObjects:=False
    ThisWorkbook.Worksheets(2).Protect Password:="123" ', DrawingObjects:=False
    ThisWorkbook.Protect "123"
End If


ComboGonderenGonderilen.Value = ""
Workbooks(FileName).Save
'ThisWorkbook.Save
MsgBox "The sender/recipient subunit information named " & ItemName & " has been successfully assigned to the relevant dropdown lists.", vbOKOnly + vbInformation, "Enterprise Document Automation System"

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
ItemName = ComboGonderenGonderilen.Value

If ItemName <> "" Then
    'Comboya tanımlı değer girilmelidir.
    ListControl = 0
    a() = ComboGonderenGonderilen.List
    For i = LBound(a) To UBound(a)
        If a(i, 0) = ItemName Then
            ListControl = 1
        End If
    Next i
Else
    GoTo Son
End If

If ListControl = 0 Then
    MsgBox "The subunit '" & ItemName & "' cannot be removed as it was not previously assigned in the dropdown lists.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
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

Set ItemBul = Workbooks(FileName).Worksheets(1).Range("CW8:CW1000").Find(What:=ItemName, SearchDirection:=xlNext, _
                SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
If Not ItemBul Is Nothing Then
    '
Else
    Workbooks(FileName).Close SaveChanges:=True
    GoTo Son
End If

Workbooks(FileName).Worksheets(1).Cells(ItemBul.Row, 101).Value = ""
ThisWorkbook.Worksheets(2).Cells(ItemBul.Row, 101).Value = ""

'Sıralama ekle komutuna da uygulanacak.
SayHedef = Workbooks(FileName).Worksheets(1).Range("CW1000").End(xlUp).Row
If SayHedef > 8 Then
    'A'dan Z'ye sırala ve böylece arada bulunan boş satırları da kaldır.
    Workbooks(FileName).Worksheets(1).Unprotect Password:="123"
    ThisWorkbook.Unprotect "123"
    ThisWorkbook.Worksheets(2).Unprotect Password:="123"
    ThisWorkbook.Worksheets(2).Visible = True
    'Sort A to Z
    Workbooks(FileName).Worksheets(1).Range("CW" & 8 & ":CW" & SayHedef).Sort key1:=Workbooks(FileName).Worksheets(1).Range("CW" & 8), order1:=xlAscending, Header:=xlNo
    ThisWorkbook.Worksheets(2).Range("CW" & 8 & ":CW" & SayHedef).Sort key1:=ThisWorkbook.Worksheets(2).Range("CW" & 8), order1:=xlAscending, Header:=xlNo
    ThisWorkbook.Worksheets(2).Visible = False
    Workbooks(FileName).Worksheets(1).Protect Password:="123" ', DrawingObjects:=False
    ThisWorkbook.Worksheets(2).Protect Password:="123" ', DrawingObjects:=False
    ThisWorkbook.Protect "123"
End If


ComboGonderenGonderilen.Value = ""
Workbooks(FileName).Save
'ThisWorkbook.Save
MsgBox "The subunit '" & ItemName & "' was successfully removed from the dropdown lists.", vbOKOnly + vbInformation, "Enterprise Document Automation System"

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

For Each ClrLab In support_contact_subunits_UI.Controls
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

support_contact_subunits_UI.BackColor = RGB(230, 230, 230) 'YENİ

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
        support_contact_subunits_UI.Width = Rep
        yukseklik = yukseklik - 60
        support_contact_subunits_UI.Height = yukseklik
        If yukseklik <= 60 Then
            yukseklik = 60
            support_contact_subunits_UI.Height = yukseklik
        End If
    ElseIf Rep <= 50 Then
        Rep = 50
        support_contact_subunits_UI.Width = Rep
        yukseklik = yukseklik - 50
        support_contact_subunits_UI.Height = yukseklik
        If yukseklik <= 50 Then
            yukseklik = 50
            support_contact_subunits_UI.Height = yukseklik
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

