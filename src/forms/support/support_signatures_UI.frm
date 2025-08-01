VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} support_signatures_UI 
   Caption         =   "Enterprise Document Automation System"
   ClientHeight    =   8895
   ClientLeft      =   30
   ClientTop       =   375
   ClientWidth     =   13920
   OleObjectBlob   =   "support_signatures_UI.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "support_signatures_UI"
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
'CheckBoxDuzelt
If CheckBoxDuzelt.BackColor <> RGB(254, 254, 254) Then
    CheckBoxDuzelt.BackColor = RGB(254, 254, 254)
    CheckBoxDuzelt.ForeColor = RGB(70, 70, 70)
End If

End Sub

Private Sub ComboAdSoyad_Change()
Dim ItemBul As Range


    On Error Resume Next
    Set ItemBul = Worksheets(2).Range("DY6:DY1000").Find(What:=ComboAdSoyad.Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    If Not ItemBul Is Nothing Then
        '
    Else
        GoTo Son
    End If
    
    ComboUnvan.Value = Worksheets(2).Cells(ItemBul.Row, 130).Value
    ComboSicil.Value = Worksheets(2).Cells(ItemBul.Row, 131).Value
    GoTo Out
    
Son:
    ComboUnvan.Value = ""
    ComboSicil.Value = ""
    
Out:


End Sub

Private Sub ComboAdSoyad_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Me.ComboAdSoyad.DropDown
End Sub
Sub TekilUnvanlar()
Dim toAdd As Boolean, UniqueUnvan As Integer, i As Integer, j As Integer
Dim SayHedef As Integer

SayHedef = ThisWorkbook.Worksheets(2).Range("DY1000").End(xlUp).Row
If SayHedef < 6 Then
    SayHedef = 6
End If
If SayHedef > 304 Then
    GoTo Son
End If

ThisWorkbook.Worksheets(2).Cells(6, 132).Value = ThisWorkbook.Worksheets(2).Cells(6, 130).Value
UniqueUnvan = 6
toAdd = True
For i = 6 To SayHedef
    For j = 6 To UniqueUnvan
        If ThisWorkbook.Worksheets(2).Cells(i, 130).Value = ThisWorkbook.Worksheets(2).Cells(j, 132).Value Then
            toAdd = False
        End If
    Next j
    If toAdd = True Then
        ThisWorkbook.Worksheets(2).Cells(UniqueUnvan + 1, 132).Value = ThisWorkbook.Worksheets(2).Cells(i, 130).Value
        UniqueUnvan = UniqueUnvan + 1
    End If
    toAdd = True
Next i

Son:
End Sub
Private Sub ComboUnvan_DropButtonClick()
Call TekilUnvanlar
End Sub

Private Sub ComboUnvan_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Me.ComboUnvan.DropDown
End Sub
Private Sub ComboSicil_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Me.ComboSicil.DropDown
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
Private Sub CheckBoxDuzelt_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
CheckBoxDuzelt.BackColor = RGB(60, 100, 180)
CheckBoxDuzelt.ForeColor = RGB(255, 255, 255)
End Sub
Private Sub ComboAdSoyad_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
Call SetComboBoxHook(ComboAdSoyad) 'Open scrollable with mouse
End Sub
Private Sub ComboUnvan_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
Call SetComboBoxHook(ComboUnvan) 'Open scrollable with mouse
End Sub
Private Sub ComboSicil_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
Call SetComboBoxHook(ComboSicil) 'Open scrollable with mouse
End Sub

Private Sub LblKisi_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub
Private Sub LblUnvan_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub
Private Sub LblSicil_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
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


Private Sub LabelEkle_Click()
Dim a() As Variant, i As Variant
Dim AutoPath As String, DestTarget As String, OpenControl As String
Dim FileName As String, SayHedef As Long, ItemName As String, j As Integer, x As Integer
Dim TC As String, ItemNameBuyuk As String, ItemName1 As String, ItemName2 As String, Soyad As String, Ad As String

Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.EnableEvents = False


AutoPath = ThisWorkbook.Path
DestTarget = AutoPath & "\System Files\System Definitions\"
FileName = "Definitions.xlsx"
ItemName = ComboAdSoyad.Value
ItemName1 = ComboUnvan.Value
ItemName2 = ComboSicil.Value
'İsim soyisim
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

    'Karakterleri otomatik düzelt
    If CheckBoxDuzelt.Value = True Then
        x = 0
        ItemName = WorksheetFunction.Proper(ItemName)
        For j = Len(ItemName) To 1 Step -1
            If Mid(ItemName, j, 1) <> " " Then
                Soyad = Mid(ItemName, j, 1) & Soyad
                x = x + 1
            Else
                x = x
                GoTo SoyadBulSon
            End If
        Next j
SoyadBulSon:
        x = Len(ItemName) - x 'Soldan ad karakter sayısı
        Ad = Left(ItemName, x)
        Soyad = UCase(Replace(Replace(Soyad, "i", "I"), "ı", "I"))
        ItemName = Ad & Soyad
    End If
    
    'Comboya tanımlı değer girilemez.
    a() = ComboAdSoyad.List
    For i = LBound(a) To UBound(a)
        If a(i, 0) = ItemName Then
            MsgBox "The person named '" & ItemName & "' has already been defined for the corresponding dropdown list, so your operation could not be completed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
            GoTo Son
        End If
    Next i
Else
    MsgBox "The person field cannot be left blank.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If
'Unvan
If ItemName1 <> "" Then
    'Birden fazla boşluk varsa kaldır
    For i = 1 To 50
        ItemName1 = Replace(ItemName1, "  ", " ")
    Next i
    'Sağdaki ve soldaki tek boşluğu kaldır
    Do While Left(ItemName1, 1) = " "
        ItemName1 = Right(ItemName1, Len(ItemName1) - 1)
    Loop
    Do While Right(ItemName1, 1) = " "
        ItemName1 = Left(ItemName1, Len(ItemName1) - 1)
    Loop
    
    'Karakterleri otomatik düzelt
    If CheckBoxDuzelt.Value = True Then
        ItemName1 = WorksheetFunction.Proper(ItemName1)
    End If
    
'    'Comboya tanımlı değer girilemez.
'    a() = ComboUnvan.List
'    For i = LBound(a) To UBound(a)
'        If a(i, 0) = ItemName1 Then
'            MsgBox ItemName1 & " isimli unvan bilgisi ilgili açılır listeler için daha önce tanımlanmış olduğundan işleminiz gerçekleştirilemedi.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
'            GoTo Son
'        End If
'    Next i
Else
    MsgBox "The title field cannot be left empty.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If
'Sicil
If ItemName2 <> "" Then
    'Birden fazla boşluk varsa kaldır
    For i = 1 To 50
        ItemName2 = Replace(ItemName2, "  ", " ")
    Next i
    'Sağdaki ve soldaki tek boşluğu kaldır
    Do While Left(ItemName2, 1) = " "
        ItemName2 = Right(ItemName2, Len(ItemName2) - 1)
    Loop
    Do While Right(ItemName2, 1) = " "
        ItemName2 = Left(ItemName2, Len(ItemName2) - 1)
    Loop
    
    'Karakterleri otomatik düzelt
    If CheckBoxDuzelt.Value = True Then
        ItemName2 = UCase(Replace(Replace(ItemName2, "i", "I"), "ı", "I"))
    End If
    
    'Comboya tanımlı değer girilemez.
    a() = ComboSicil.List
    For i = LBound(a) To UBound(a)
        If a(i, 0) = ItemName2 Then
            MsgBox "The registry number " & ItemName2 & " has already been assigned to another person, so the operation could not be completed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
            GoTo Son
        End If
    Next i
Else
    MsgBox "The registry number field cannot be left empty.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
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

SayHedef = Workbooks(FileName).Worksheets(1).Range("DY1000").End(xlUp).Row
If SayHedef < 6 Then
    SayHedef = 6
End If
If SayHedef > 304 Then
    MsgBox "The dropdown list definition area for person selection is full, so the person information named " & ItemName & " could not be defined.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    Workbooks(FileName).Close SaveChanges:=True
    GoTo Son
End If

'Arada boş satır varsa onu bul ve öğeyi boş satıra yaz.
'If SayHedef > 6 Then
    For j = 6 To SayHedef
        If Workbooks(FileName).Worksheets(1).Cells(j, 129).Value = "" Then
            SayHedef = j - 1
            GoTo DonguSon
        End If
    Next j
'End If
DonguSon:
'Ve kelimelerini düzelt
If InStr(ItemName, " And ") <> 0 Then
    ItemName = Replace(ItemName, " And ", " and ")
End If
If InStr(ItemName1, " And ") <> 0 Then
    ItemName1 = Replace(ItemName1, " And ", " and ")
End If
If InStr(ItemName2, " And ") <> 0 Then
    ItemName2 = Replace(ItemName2, " And ", " and ")
End If

Workbooks(FileName).Worksheets(1).Cells(SayHedef + 1, 129).Value = ItemName
Workbooks(FileName).Worksheets(1).Cells(SayHedef + 1, 130).Value = ItemName1
Workbooks(FileName).Worksheets(1).Cells(SayHedef + 1, 131).Value = ItemName2
ThisWorkbook.Worksheets(2).Cells(SayHedef + 1, 129).Value = ItemName
ThisWorkbook.Worksheets(2).Cells(SayHedef + 1, 130).Value = ItemName1
ThisWorkbook.Worksheets(2).Cells(SayHedef + 1, 131).Value = ItemName2
ComboAdSoyad.Value = ""
ComboUnvan.Value = ""
ComboSicil.Value = ""

'Sıralama ekle komutuna da uygulanacak.
SayHedef = Workbooks(FileName).Worksheets(1).Range("DY1000").End(xlUp).Row
If SayHedef > 6 Then
    'A'dan Z'ye sırala ve böylece arada bulunan boş satırları da kaldır.
    Workbooks(FileName).Worksheets(1).Unprotect Password:="123"
    ThisWorkbook.Unprotect "123"
    ThisWorkbook.Worksheets(2).Unprotect Password:="123"
    ThisWorkbook.Worksheets(2).Visible = True
    'Sort A to Z
    Workbooks(FileName).Worksheets(1).Range("DY" & 6 & ":EA" & SayHedef).Sort key1:=Workbooks(FileName).Worksheets(1).Range("DY" & 6 & ":DY" & SayHedef), order1:=xlAscending, Header:=xlNo
    ThisWorkbook.Worksheets(2).Range("DY" & 6 & ":EA" & SayHedef).Sort key1:=ThisWorkbook.Worksheets(2).Range("DY" & 6 & ":DY" & SayHedef), order1:=xlAscending, Header:=xlNo
    ThisWorkbook.Worksheets(2).Visible = False
    Workbooks(FileName).Worksheets(1).Protect Password:="123" ', DrawingObjects:=False
    ThisWorkbook.Worksheets(2).Protect Password:="123" ', DrawingObjects:=False
    ThisWorkbook.Protect "123"
End If


ThisWorkbook.Worksheets(2).Range("EB6:EB305").ClearContents
Call TekilUnvanlar

'Açık dropdown kapat
Call ModuleSystemSettings.DropDownKapat

Workbooks(FileName).Save
'ThisWorkbook.Save
MsgBox "The person named " & ItemName & " has been successfully defined in the system with the title " & ItemName1 & " and the registration number " & ItemName2 & ".", vbOKOnly + vbInformation, "Enterprise Document Automation System"

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
Dim a() As Variant, b() As Variant, c() As Variant, i As Variant
Dim AutoPath As String, DestTarget As String, OpenControl As String, ItemName As String
Dim FileName As String, ListControl As Integer, ItemBul As Range, ItemName1 As String, ItemName2 As String
Dim counter As Integer, SayHedef As Integer, ListControl1 As Integer, ListControl2 As Integer

Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.EnableEvents = False

AutoPath = ThisWorkbook.Path
DestTarget = AutoPath & "\System Files\System Definitions\"
FileName = "Definitions.xlsx"
ItemName = ComboAdSoyad.Value
ItemName1 = ComboUnvan.Value
ItemName2 = ComboSicil.Value

If ItemName <> "" Then
    'Comboya tanımlı değer girilmelidir.
    ListControl = 0
    a() = ComboAdSoyad.List
    For i = LBound(a) To UBound(a)
        If a(i, 0) = ItemName Then
            ListControl = 1
        End If
    Next i
Else
    MsgBox "The person field cannot be left empty.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If
'Unvan combosunun içeriğini oluşturduktan sonra kontrol et. Yoksa combo içeriği bopş kalıyor.
ThisWorkbook.Worksheets(2).Range("EB6:EB305").ClearContents
Call TekilUnvanlar
If ItemName1 <> "" Then
    'Comboya tanımlı değer girilmelidir.
    ListControl1 = 0
    b() = ComboUnvan.List
    For i = LBound(b) To UBound(b)
        If b(i, 0) = ItemName1 Then
            ListControl1 = 1
        End If
    Next i
Else
    MsgBox "The title field cannot be left empty.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If
If ItemName2 <> "" Then
    'Comboya tanımlı değer girilmelidir.
    ListControl2 = 0
    c() = ComboSicil.List
    For i = LBound(c) To UBound(c)
        If c(i, 0) = ItemName2 Then
            ListControl2 = 1
        End If
    Next i
Else
    MsgBox "The registration number field cannot be left empty.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

If ListControl = 0 Then
    MsgBox "The person named " & ItemName & " has not been previously defined in the relevant dropdown list, so the removal operation could not be completed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

If ListControl1 = 0 Then
    MsgBox "The title '" & ItemName1 & "' for the person named " & ItemName & " has not been previously defined in the relevant dropdown list, so the removal operation could not be completed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

If ListControl2 = 0 Then
    MsgBox "The registration number '" & ItemName2 & "' for the person named " & ItemName & " has not been previously defined in the relevant dropdown list, so the removal operation could not be completed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
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

Set ItemBul = Workbooks(FileName).Worksheets(1).Range("DY6:DY1000").Find(What:=ItemName, SearchDirection:=xlNext, _
                SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
If Not ItemBul Is Nothing Then
    '
Else
    Workbooks(FileName).Close SaveChanges:=True
    GoTo Son
End If

Workbooks(FileName).Worksheets(1).Cells(ItemBul.Row, 129).Value = ""
Workbooks(FileName).Worksheets(1).Cells(ItemBul.Row, 130).Value = ""
Workbooks(FileName).Worksheets(1).Cells(ItemBul.Row, 131).Value = ""
ThisWorkbook.Worksheets(2).Cells(ItemBul.Row, 129).Value = ""
ThisWorkbook.Worksheets(2).Cells(ItemBul.Row, 130).Value = ""
ThisWorkbook.Worksheets(2).Cells(ItemBul.Row, 131).Value = ""
ComboAdSoyad.Value = ""
ComboUnvan.Value = ""
ComboSicil.Value = ""

'Boşlukları kaldır
SayHedef = ThisWorkbook.Worksheets(2).Range("DY1000").End(xlUp).Row
If SayHedef < 6 Then
    SayHedef = 6
End If
If SayHedef > 304 Then
    SayHedef = 304
End If
counter = 5
For i = 6 To SayHedef
    If ThisWorkbook.Worksheets(2).Cells(i, 129).Value <> "" Then
        ThisWorkbook.Worksheets(2).Cells(counter + 1, 133).Value = ThisWorkbook.Worksheets(2).Cells(i, 129).Value
        ThisWorkbook.Worksheets(2).Cells(counter + 1, 134).Value = ThisWorkbook.Worksheets(2).Cells(i, 130).Value
        ThisWorkbook.Worksheets(2).Cells(counter + 1, 135).Value = ThisWorkbook.Worksheets(2).Cells(i, 131).Value
        Workbooks(FileName).Worksheets(1).Cells(counter + 1, 133).Value = Workbooks(FileName).Worksheets(1).Cells(i, 129).Value
        Workbooks(FileName).Worksheets(1).Cells(counter + 1, 134).Value = Workbooks(FileName).Worksheets(1).Cells(i, 130).Value
        Workbooks(FileName).Worksheets(1).Cells(counter + 1, 135).Value = Workbooks(FileName).Worksheets(1).Cells(i, 131).Value
        counter = counter + 1
    End If
Next i
Workbooks(FileName).Worksheets(1).Range("DY6:EA305").Value = ""
Workbooks(FileName).Worksheets(1).Range("DY6:EA305").Value = ThisWorkbook.Worksheets(2).Range("EC6:EE305").Value
Workbooks(FileName).Worksheets(1).Range("EC6:EE305").Value = ""
ThisWorkbook.Worksheets(2).Range("DY6:EA305").Value = ""
ThisWorkbook.Worksheets(2).Range("DY6:EA305").Value = ThisWorkbook.Worksheets(2).Range("EC6:EE305").Value
ThisWorkbook.Worksheets(2).Range("EC6:EE305").Value = ""

ThisWorkbook.Worksheets(2).Range("EB6:EB305").ClearContents
Call TekilUnvanlar

'Sıralama ekle komutuna da uygulanacak.
SayHedef = Workbooks(FileName).Worksheets(1).Range("DY1000").End(xlUp).Row
If SayHedef > 6 Then
    'A'dan Z'ye sırala ve böylece arada bulunan boş satırları da kaldır.
    Workbooks(FileName).Worksheets(1).Unprotect Password:="123"
    ThisWorkbook.Unprotect "123"
    ThisWorkbook.Worksheets(2).Unprotect Password:="123"
    ThisWorkbook.Worksheets(2).Visible = True
    'Sort A to Z
    Workbooks(FileName).Worksheets(1).Range("DY" & 6 & ":EA" & SayHedef).Sort key1:=Workbooks(FileName).Worksheets(1).Range("DY" & 6 & ":DY" & SayHedef), order1:=xlAscending, Header:=xlNo
    ThisWorkbook.Worksheets(2).Range("DY" & 6 & ":EA" & SayHedef).Sort key1:=ThisWorkbook.Worksheets(2).Range("DY" & 6 & ":DY" & SayHedef), order1:=xlAscending, Header:=xlNo
    ThisWorkbook.Worksheets(2).Visible = False
    Workbooks(FileName).Worksheets(1).Protect Password:="123" ', DrawingObjects:=False
    ThisWorkbook.Worksheets(2).Protect Password:="123" ', DrawingObjects:=False
    ThisWorkbook.Protect "123"
End If

'Açık dropdown kapat
Call ModuleSystemSettings.DropDownKapat

Workbooks(FileName).Save
'ThisWorkbook.Save
MsgBox "The person named " & ItemName & ", with the title '" & ItemName1 & "' and registration number '" & ItemName2 & "', has been successfully removed from the system.", vbOKOnly + vbInformation, "Enterprise Document Automation System"

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
Dim SayHedef As Integer

SayHedef = Worksheets(2).Range("DY1000").End(xlUp).Row
If SayHedef < 6 Then
    SayHedef = 6
End If
Worksheets(2).Range("EB6:EB305").ClearContents
    
Unload Me

End Sub


Private Sub UserForm_Initialize()
Dim i As Long, WsSKP As Object
Dim ClrLab As MSForms.Control
Dim SayHedef As Integer


ThisWorkbook.Activate

For Each ClrLab In support_signatures_UI.Controls
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

support_signatures_UI.BackColor = RGB(230, 230, 230) 'YENİ



SayHedef = Worksheets(2).Range("DY1000").End(xlUp).Row
If SayHedef < 6 Then
    SayHedef = 6
End If
Worksheets(2).Range("EB6:EB305").ClearContents

CheckBoxDuzelt.Value = True


End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
Dim yukseklik As Variant, genislik As Variant
Dim Rep As Variant
Dim SayHedef As Integer

SayHedef = Worksheets(2).Range("DY1000").End(xlUp).Row
If SayHedef < 6 Then
    SayHedef = 6
End If
Worksheets(2).Range("EB6:EB305").ClearContents

yukseklik = Me.Height
Rep = Me.Width
Do
DoEvents
Rep = Rep - 60
Call timeout(0.01)
    If Rep > 60 Then
        support_signatures_UI.Width = Rep
        yukseklik = yukseklik - 60
        support_signatures_UI.Height = yukseklik
        If yukseklik <= 60 Then
            yukseklik = 60
            support_signatures_UI.Height = yukseklik
        End If
    ElseIf Rep <= 50 Then
        Rep = 50
        support_signatures_UI.Width = Rep
        yukseklik = yukseklik - 50
        support_signatures_UI.Height = yukseklik
        If yukseklik <= 50 Then
            yukseklik = 50
            support_signatures_UI.Height = yukseklik
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
