VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} support_item_types_UI 
   Caption         =   "Enterprise Document Automation System"
   ClientHeight    =   8895
   ClientLeft      =   30
   ClientTop       =   375
   ClientWidth     =   13920
   OleObjectBlob   =   "support_item_types_UI.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "support_item_types_UI"
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
'LabelKapat
If LabelKapat.BackColor <> RGB(225, 235, 245) Then
    LabelKapat.BackColor = RGB(225, 235, 245)
    LabelKapat.ForeColor = RGB(30, 30, 30)
End If

End Sub

Private Sub ComboOgeTuru_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Me.ComboOgeTuru.DropDown
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

Private Sub ComboOgeTuru_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
Call SetComboBoxHook(ComboOgeTuru) 'Open scrollable with mouse
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
Dim Kisaltma As String

Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.EnableEvents = False

AutoPath = ThisWorkbook.Path
DestTarget = AutoPath & "\System Files\System Definitions\"
FileName = "Definitions.xlsx"
ItemName = ComboOgeTuru.Value

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
    'İlk harfler büyük
    ItemName = WorksheetFunction.Proper(ItemName)
    'Tireden sonraki harfler büyük. Tireden önce ve sonra boşluk ekle.
    If InStr(ItemName, "-") <> 0 Then
        If Mid(ItemName, InStr(ItemName, "-") - 1, 1) = " " Then
            '
        Else
            ItemName = Replace(ItemName, "-", " -")
        End If
        If Mid(ItemName, InStr(ItemName, "-") + 1, 1) = " " Then
            '
        Else
            ItemName = Replace(ItemName, "-", "- ")
        End If
        Kisaltma = Mid(ItemName, InStr(ItemName, "-") + 1, Len(ItemName) - InStr(ItemName, "-"))
        'MsgBox Kisaltma
        If Kisaltma = " X2" Then
            Kisaltma = " X2"
        Else
            Kisaltma = UCase(Replace(Replace(Kisaltma, "ı", "I"), "i", "I"))
        End If
        ItemName = Left(ItemName, InStr(ItemName, "-") - 1) & "-" & Kisaltma
    Else
    MsgBox "Since a hyphen (-) was not detected before the item type abbreviation, the operation could not be completed. Please enter the item type as shown in the example above.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    'Comboya tanımlı değer girilemez.
    a() = ComboOgeTuru.List
    For i = LBound(a) To UBound(a)
        If a(i, 0) = ItemName Then
            MsgBox "The item type named " & ItemName & " has already been defined for the related dropdown lists, and therefore the operation could not be completed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
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

SayHedef = Workbooks(FileName).Worksheets(1).Range("CY1000").End(xlUp).Row
If SayHedef < 6 Then
    SayHedef = 6
End If
If SayHedef > 104 Then
    MsgBox "The dropdown list for item type selection is full, so the item type named " & ItemName & " could not be defined.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    Workbooks(FileName).Close SaveChanges:=True
    GoTo Son
End If

'Arada boş satır varsa onu bul ve öğeyi boş satıra yaz.
If SayHedef > 6 Then
    For j = 6 To SayHedef
        If Workbooks(FileName).Worksheets(1).Cells(j, 103).Value = "" Then
            SayHedef = j - 1
            GoTo DonguSon
        End If
    Next j
End If
DonguSon:
'Ve kelimelerini düzelt
If InStr(ItemName, " And ") <> 0 Then
    ItemName = Replace(ItemName, " And ", " and ")
End If

Workbooks(FileName).Worksheets(1).Cells(SayHedef + 1, 103).Value = ItemName
ThisWorkbook.Worksheets(2).Cells(SayHedef + 1, 103).Value = ItemName

'Sıralama ekle komutuna da uygulanacak.
SayHedef = Workbooks(FileName).Worksheets(1).Range("CY1000").End(xlUp).Row
If SayHedef > 7 Then
    'A'dan Z'ye sırala ve böylece arada bulunan boş satırları da kaldır.
    Workbooks(FileName).Worksheets(1).Unprotect Password:="123"
    ThisWorkbook.Unprotect "123"
    ThisWorkbook.Worksheets(2).Unprotect Password:="123"
    ThisWorkbook.Worksheets(2).Visible = True
    'Sort A to Z
    Workbooks(FileName).Worksheets(1).Range("CY" & 7 & ":CY" & SayHedef).Sort key1:=Workbooks(FileName).Worksheets(1).Range("CY" & 7), order1:=xlAscending, Header:=xlNo
    ThisWorkbook.Worksheets(2).Range("CY" & 7 & ":CY" & SayHedef).Sort key1:=ThisWorkbook.Worksheets(2).Range("CY" & 7), order1:=xlAscending, Header:=xlNo
    ThisWorkbook.Worksheets(2).Visible = False
    Workbooks(FileName).Worksheets(1).Protect Password:="123" ', DrawingObjects:=False
    ThisWorkbook.Worksheets(2).Protect Password:="123" ', DrawingObjects:=False
    ThisWorkbook.Protect "123"
End If

ComboOgeTuru.Value = ""
Workbooks(FileName).Save
'ThisWorkbook.Save
MsgBox "The item type named " & ItemName & " has been successfully defined for the related dropdown lists.", vbOKOnly + vbInformation, "Enterprise Document Automation System"

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
Dim Kisaltma As String

Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.EnableEvents = False

AutoPath = ThisWorkbook.Path
DestTarget = AutoPath & "\System Files\System Definitions\"
FileName = "Definitions.xlsx"
ItemName = ComboOgeTuru.Value

If ItemName <> "" Then
    'Comboya tanımlı değer girilmelidir.
    ListControl = 0
    a() = ComboOgeTuru.List
    For i = LBound(a) To UBound(a)
        If a(i, 0) = ItemName Then
            ListControl = 1
        End If
    Next i
Else
    GoTo Son
End If

If ListControl = 0 Then
    MsgBox "The item type named " & ItemName & " has not been previously defined for the related dropdown lists, and therefore the removal operation could not be completed.", vbOKOnly + vbInformation, "Enterprise Document Automation System"

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

Set ItemBul = Workbooks(FileName).Worksheets(1).Range("CY6:CY1000").Find(What:=ItemName, SearchDirection:=xlNext, _
                SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
If Not ItemBul Is Nothing Then
    '
Else
    Workbooks(FileName).Close SaveChanges:=True
    GoTo Son
End If

Workbooks(FileName).Worksheets(1).Cells(ItemBul.Row, 103).Value = ""
ThisWorkbook.Worksheets(2).Cells(ItemBul.Row, 103).Value = ""

'Sıralama ekle komutuna da uygulanacak.
SayHedef = Workbooks(FileName).Worksheets(1).Range("CY1000").End(xlUp).Row
If SayHedef > 7 Then
    'A'dan Z'ye sırala ve böylece arada bulunan boş satırları da kaldır.
    Workbooks(FileName).Worksheets(1).Unprotect Password:="123"
    ThisWorkbook.Unprotect "123"
    ThisWorkbook.Worksheets(2).Unprotect Password:="123"
    ThisWorkbook.Worksheets(2).Visible = True
    'Sort A to Z
    Workbooks(FileName).Worksheets(1).Range("CY" & 7 & ":CY" & SayHedef).Sort key1:=Workbooks(FileName).Worksheets(1).Range("CY" & 7), order1:=xlAscending, Header:=xlNo
    ThisWorkbook.Worksheets(2).Range("CY" & 7 & ":CY" & SayHedef).Sort key1:=ThisWorkbook.Worksheets(2).Range("CY" & 7), order1:=xlAscending, Header:=xlNo
    ThisWorkbook.Worksheets(2).Visible = False
    Workbooks(FileName).Worksheets(1).Protect Password:="123" ', DrawingObjects:=False
    ThisWorkbook.Worksheets(2).Protect Password:="123" ', DrawingObjects:=False
    ThisWorkbook.Protect "123"
End If

ComboOgeTuru.Value = ""
Workbooks(FileName).Save
'ThisWorkbook.Save
MsgBox "The item type named " & ItemName & " has been successfully removed from the related dropdown lists.", vbOKOnly + vbInformation, "Enterprise Document Automation System"


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

For Each ClrLab In support_item_types_UI.Controls
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

support_item_types_UI.BackColor = RGB(230, 230, 230) 'YENİ

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
        support_item_types_UI.Width = Rep
        yukseklik = yukseklik - 60
        support_item_types_UI.Height = yukseklik
        If yukseklik <= 60 Then
            yukseklik = 60
            support_item_types_UI.Height = yukseklik
        End If
    ElseIf Rep <= 50 Then
        Rep = 50
        support_item_types_UI.Width = Rep
        yukseklik = yukseklik - 50
        support_item_types_UI.Height = yukseklik
        If yukseklik <= 50 Then
            yukseklik = 50
            support_item_types_UI.Height = yukseklik
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


