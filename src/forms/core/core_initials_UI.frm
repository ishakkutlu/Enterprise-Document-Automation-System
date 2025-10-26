VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} core_initials_UI 
   Caption         =   "Enterprise Document Automation System"
   ClientHeight    =   8895
   ClientLeft      =   30
   ClientTop       =   375
   ClientWidth     =   13920
   OleObjectBlob   =   "core_initials_UI.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "core_initials_UI"
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

Private Sub ComboUserName_Change()
Dim ItemBul As Range

    On Error Resume Next
    Set ItemBul = Worksheets(2).Range("DR6:DR1000").Find(What:=ComboUserName.Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    If Not ItemBul Is Nothing Then
        '
    Else
        GoTo Son
    End If
    
    ComboAdSoyad.Value = Worksheets(2).Cells(ItemBul.Row, 123).Value
    ComboUnvan.Value = Worksheets(2).Cells(ItemBul.Row, 124).Value
    ComboSicil.Value = Worksheets(2).Cells(ItemBul.Row, 125).Value
    ComboTel.Value = Worksheets(2).Cells(ItemBul.Row, 126).Value
    GoTo Out
    
Son:
    ComboAdSoyad.Value = ""
    ComboUnvan.Value = ""
    ComboSicil.Value = ""
    ComboTel.Value = ""
    
Out:


End Sub

Private Sub ComboUserName_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Me.ComboUserName.DropDown
End Sub

Private Sub ComboAdSoyad_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Me.ComboAdSoyad.DropDown
End Sub
Sub TekilUnvanlar()
Dim toAdd As Boolean, UniqueUnvan As Integer, i As Integer, j As Integer
Dim SayHedef As Integer

SayHedef = ThisWorkbook.Worksheets(2).Range("DR1000").End(xlUp).Row
If SayHedef < 6 Then
    SayHedef = 6
End If
If SayHedef > 104 Then
    GoTo Son
End If

ThisWorkbook.Worksheets(2).Cells(6, 132).Value = ThisWorkbook.Worksheets(2).Cells(6, 124).Value
UniqueUnvan = 6
toAdd = True
For i = 6 To SayHedef
    For j = 6 To UniqueUnvan
        If ThisWorkbook.Worksheets(2).Cells(i, 124).Value = ThisWorkbook.Worksheets(2).Cells(j, 132).Value Then
            toAdd = False
        End If
    Next j
    If toAdd = True Then
        ThisWorkbook.Worksheets(2).Cells(UniqueUnvan + 1, 132).Value = ThisWorkbook.Worksheets(2).Cells(i, 124).Value
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
Private Sub ComboTel_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Me.ComboTel.DropDown
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
Private Sub ComboUserName_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
Call SetComboBoxHook(ComboUserName) 'Open scrollable with mouse
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
Private Sub ComboTel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
Call SetComboBoxHook(ComboTel) 'Open scrollable with mouse
End Sub

Private Sub LblUserNameBilgi_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub
Private Sub LblOturumSahibi_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub
Private Sub LblUserName_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
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
Private Sub LblTel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
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
Dim ItemName3 As String, ItemUserName As String, ItemDuzenle As String, UserName As String
Dim ItemBul As Range

Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.EnableEvents = False


AutoPath = ThisWorkbook.Path
DestTarget = AutoPath & "\System Files\System Definitions\"
FileName = "Definitions.xlsx"
ItemName = ComboAdSoyad.Value
ItemName1 = ComboUnvan.Value
ItemName2 = ComboSicil.Value
ItemName3 = ComboTel.Value
ItemUserName = ComboUserName.Value


'Oturum adı
If ItemUserName <> "" Then

    'Boşlukları kaldır
    For i = 1 To 50
        ItemUserName = Replace(ItemUserName, " ", "")
    Next i
    
    If CheckBoxDuzelt.Value = True Then
        'Alfabetik karakter büyük harf olsun
        'ItemUserName = UCase(Replace(Replace(ItemUserName, "i", "I"), "ı", "I"))
        ItemUserName = UCase(ItemUserName)
    End If
    
'    'Comboya tanımlı değer girilemez.
'    a() = ComboUserName.List
'    For i = LBound(a) To UBound(a)
'        If a(i, 0) = ItemUserName Then
'        MsgBox "The session name '" & ItemUserName & "' has already been defined for the relevant dropdown list, so your operation could not be completed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
'            GoTo Son
'        End If
'    Next i
Else
    MsgBox "The session name field cannot be left empty.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

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
Else
    MsgBox "The person field cannot be left empty.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
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
    
Else
    MsgBox "The title field cannot be left empty.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If
'Sicil
If ItemName2 <> "" Then
    'Tüm boşlukları kaldır
    For i = 1 To 50
        ItemName2 = Replace(ItemName2, " ", "")
    Next i
    
    'Karakterleri otomatik düzelt
    If CheckBoxDuzelt.Value = True Then
        ItemName2 = UCase(Replace(Replace(ItemName2, "i", "I"), "ı", "I"))
    End If
    
Else
    MsgBox "The registration number field cannot be left empty.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

'Tel
If ItemName3 <> "" Then
    If CheckBoxDuzelt.Value = True Then
        'Tüm boşlukları kaldır
        For i = 1 To 50
            ItemName3 = Replace(ItemName3, " ", "")
        Next i
        If Len(ItemName3) > 4 Then
            If Left(ItemName3, 1) <> "(" Then '(0xxx) Solda parantez yoksa ekle
                ItemName3 = "(" & ItemName3
            End If
    
            If Mid(ItemName3, 2, 1) <> 0 Then '(0xxx) Parantez içinde 0 yoksa ekle
                ItemName3 = Replace(ItemName3, "(", "(0")
            End If
    
            If Mid(ItemName3, 6, 1) <> ")" Then '(0xxx) Sağda parantez yoksa ekle
                ItemDuzenle = Left(ItemName3, 5) & ")"
                ItemName3 = ItemDuzenle & Mid(ItemName3, 6, Len(ItemName3))
            End If
            
            If Len(ItemName3) <> 13 Then 'Parantezler ve 0 dahil 13 karakter olmalı
                MsgBox "The phone number for " & ItemName3 & " is invalid because, excluding spaces and parentheses, it does not consist of 11 digits or an internal extension number with 4 digits. Please enter your phone number in the format 0xxx xxx xx xx for an 11-digit number, or enter your internal extension as a 4-digit number (xxxx) in the Tel field.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
                GoTo Son
            End If
            
            'Numaralar arasına boşluk ekle
            ItemDuzenle = Left(ItemName3, 6) & " " 'Birinci boşluk
            ItemDuzenle = ItemDuzenle & Mid(ItemName3, 7, 3) & " " 'İkinci boşluk
            ItemDuzenle = ItemDuzenle & Mid(ItemName3, 10, 2) & " " 'Üçüncü boşluk
            ItemDuzenle = ItemDuzenle & Mid(ItemName3, 12, 2)
            ItemName3 = ItemDuzenle
        Else
            If Len(ItemName3) <> 4 Then 'Parantezler ve 0 dahil 13 karakter olmalı
                MsgBox "The internal phone number for " & ItemName3 & " is invalid because it does not consist of 4 digits. Please enter your internal phone number as a 4-digit number in the Tel field.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
                GoTo Son
            End If
        End If
        
    End If
    
    'MsgBox ItemName3
    
Else
    MsgBox "The phone field cannot be left empty. Please enter your phone number as an 11-digit number (starting with 0) or your internal extension as a 4-digit number in the phone field.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

'MsgBox ItemUserName & " : " & ItemName & " : " & ItemName1 & " : " & ItemName2 & " : " & ItemName3


OpenControl = IsFileOpen(DestTarget & FileName)
If OpenControl = True Then 'Açıksa
    Workbooks(FileName).Close SaveChanges:=True
ElseIf OpenControl = False Then
    '
End If
Workbooks.Open (DestTarget & FileName)
Workbooks(FileName).Worksheets(1).Activate

'____________


'UPDATE

Dim rng As Range
Dim firstAddress As String
Dim foundCell As Range

Set foundCell = ThisWorkbook.Worksheets(2).Range("DS6:DS104").Find(What:=ComboAdSoyad.Value, LookIn:=xlValues, LookAt:=xlWhole)
If Not foundCell Is Nothing Then
    firstAddress = foundCell.Address
    Do
        ' Aynı satırdaki ComboAdSoyad değerini kontrol et (DR sütunu)
        If ThisWorkbook.Worksheets(2).Cells(foundCell.Row, "DR").Value = ComboUserName.Value Then
            SayHedef = foundCell.Row
            GoTo Updade
        End If
        Set foundCell = rng.FindNext(foundCell)
    Loop While Not foundCell Is Nothing And foundCell.Address <> firstAddress
End If

'____________


SayHedef = Workbooks(FileName).Worksheets(1).Range("DR1000").End(xlUp).Row
If SayHedef < 6 Then
    SayHedef = 6
ElseIf SayHedef > 104 Then
    MsgBox "The dropdown list for session names is full, so the session named '" & ItemUserName & "' could not be added.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    Workbooks(FileName).Close SaveChanges:=True
    GoTo Son
Else
    SayHedef = SayHedef + 1
End If

'____________

Updade:

'Kayıt işlemi: dış dosya
With Workbooks(FileName).Worksheets(1)
    .Cells(SayHedef, 122).Value = ItemUserName
    .Cells(SayHedef, 123).Value = ItemName
    .Cells(SayHedef, 124).Value = ItemName1
    .Cells(SayHedef, 125).Value = ItemName2
    .Cells(SayHedef, 126).Value = ItemName3
End With

'Kayıt işlemi: iç dosya
With ThisWorkbook.Worksheets(2)
    .Cells(SayHedef, 122).Value = ItemUserName
    .Cells(SayHedef, 123).Value = ItemName
    .Cells(SayHedef, 124).Value = ItemName1
    .Cells(SayHedef, 125).Value = ItemName2
    .Cells(SayHedef, 126).Value = ItemName3
End With


ComboUserName.Value = ""
ComboAdSoyad.Value = ""
ComboUnvan.Value = ""
ComboSicil.Value = ""
ComboTel.Value = ""

'Sıralama ekle komutuna da uygulanacak.
SayHedef = Workbooks(FileName).Worksheets(1).Range("DR1000").End(xlUp).Row
If SayHedef >= 6 Then
    'A'dan Z'ye sırala ve böylece arada bulunan boş satırları da kaldır.
    Workbooks(FileName).Worksheets(1).Unprotect Password:="123"
    ThisWorkbook.Unprotect "123"
    ThisWorkbook.Worksheets(2).Unprotect Password:="123"
    ThisWorkbook.Worksheets(2).Visible = True
    'Sort A to Z
    Workbooks(FileName).Worksheets(1).Range("DR" & 6 & ":DV" & SayHedef).Sort key1:=Workbooks(FileName).Worksheets(1).Range("DR" & 6 & ":DR" & SayHedef), order1:=xlAscending, Header:=xlNo
    ThisWorkbook.Worksheets(2).Range("DR" & 6 & ":DV" & SayHedef).Sort key1:=ThisWorkbook.Worksheets(2).Range("DR" & 6 & ":DR" & SayHedef), order1:=xlAscending, Header:=xlNo
    ThisWorkbook.Worksheets(2).Visible = False
    Workbooks(FileName).Worksheets(1).Protect Password:="123" ', DrawingObjects:=False
    ThisWorkbook.Worksheets(2).Protect Password:="123" ', DrawingObjects:=False
    ThisWorkbook.Protect "123"
End If


ThisWorkbook.Worksheets(2).Range("EB6:EB105").ClearContents
Call TekilUnvanlar

'Etiketi düzeltmen gerekiyorsa düzelt
UserName = Environ("UserProfile")
UserName = UCase(Right(UserName, 7))
If ItemUserName = UserName Then
    LblUserNameBilgi.Caption = " User " & UserName & " has been successfully registered in the system."
    LblUserNameBilgi.ForeColor = RGB(19, 117, 71)
End If

'Açık dropdown kapat
Call ModuleSystemSettings.DropDownKapat

Workbooks(FileName).Save
'ThisWorkbook.Save

MsgBox "The session named '" & ItemUserName & "' has been successfully registered in the system with the following details:" & vbNewLine & vbNewLine & _
"Session Name: " & ItemUserName & vbNewLine & _
"Full Name: " & ItemName & vbNewLine & _
"Title: " & ItemName1 & vbNewLine & _
"Registration No.: " & ItemName2 & vbNewLine & _
"Phone: " & ItemName3, vbOKOnly + vbInformation, "Enterprise Document Automation System"


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
Dim ItemName3 As String, ItemUserName As String, ItemDuzenle As String, UserName As String

Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.EnableEvents = False

AutoPath = ThisWorkbook.Path
DestTarget = AutoPath & "\System Files\System Definitions\"
FileName = "Definitions.xlsx"
ItemName = ComboAdSoyad.Value
ItemName1 = ComboUnvan.Value
ItemName2 = ComboSicil.Value
ItemName3 = ComboTel.Value
ItemUserName = ComboUserName.Value

'Oturum adı
If ItemUserName <> "" Then
    'Boşlukları kaldır
    For i = 1 To 50
        ItemUserName = Replace(ItemUserName, " ", "")
    Next i
    
    If CheckBoxDuzelt.Value = True Then
        'Alfabetik karakter büyük harf olsun
        'ItemUserName = UCase(Replace(Replace(ItemUserName, "i", "I"), "ı", "I"))
        ItemUserName = UCase(ItemUserName)
    End If
    
    'Comboya tanımlı değer girilmelidir.
    ListControl = 0
    a() = ComboUserName.List
    For i = LBound(a) To UBound(a)
        If a(i, 0) = ItemUserName Then
            ListControl = 1
        End If
    Next i
Else
    MsgBox "The session name field cannot be left empty.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

If ListControl = 0 Then
    MsgBox "The session named '" & ItemUserName & "' was not previously defined, so the removal operation could not be performed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
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

Set ItemBul = Workbooks(FileName).Worksheets(1).Range("DR6:DR1000").Find(What:=ItemUserName, SearchDirection:=xlNext, _
                SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
If Not ItemBul Is Nothing Then
    '
Else
    Workbooks(FileName).Close SaveChanges:=True
    GoTo Son
End If

Workbooks(FileName).Worksheets(1).Cells(ItemBul.Row, 122).Value = ""
Workbooks(FileName).Worksheets(1).Cells(ItemBul.Row, 123).Value = ""
Workbooks(FileName).Worksheets(1).Cells(ItemBul.Row, 124).Value = ""
Workbooks(FileName).Worksheets(1).Cells(ItemBul.Row, 125).Value = ""
Workbooks(FileName).Worksheets(1).Cells(ItemBul.Row, 126).Value = ""

ThisWorkbook.Worksheets(2).Cells(ItemBul.Row, 122).Value = ""
ThisWorkbook.Worksheets(2).Cells(ItemBul.Row, 123).Value = ""
ThisWorkbook.Worksheets(2).Cells(ItemBul.Row, 124).Value = ""
ThisWorkbook.Worksheets(2).Cells(ItemBul.Row, 125).Value = ""
ThisWorkbook.Worksheets(2).Cells(ItemBul.Row, 126).Value = ""

ComboUserName.Value = ""
ComboAdSoyad.Value = ""
ComboUnvan.Value = ""
ComboSicil.Value = ""
ComboTel.Value = ""

'Boşlukları kaldır
SayHedef = ThisWorkbook.Worksheets(2).Range("DR1000").End(xlUp).Row
If SayHedef < 6 Then
    SayHedef = 6
End If
If SayHedef > 104 Then
    SayHedef = 104
End If
counter = 5
For i = 6 To SayHedef
    If ThisWorkbook.Worksheets(2).Cells(i, 122).Value <> "" Then
        ThisWorkbook.Worksheets(2).Cells(counter + 1, 133).Value = ThisWorkbook.Worksheets(2).Cells(i, 122).Value
        ThisWorkbook.Worksheets(2).Cells(counter + 1, 134).Value = ThisWorkbook.Worksheets(2).Cells(i, 123).Value
        ThisWorkbook.Worksheets(2).Cells(counter + 1, 135).Value = ThisWorkbook.Worksheets(2).Cells(i, 124).Value
        ThisWorkbook.Worksheets(2).Cells(counter + 1, 136).Value = ThisWorkbook.Worksheets(2).Cells(i, 125).Value
        ThisWorkbook.Worksheets(2).Cells(counter + 1, 137).Value = ThisWorkbook.Worksheets(2).Cells(i, 126).Value
        
        Workbooks(FileName).Worksheets(1).Cells(counter + 1, 133).Value = Workbooks(FileName).Worksheets(1).Cells(i, 122).Value
        Workbooks(FileName).Worksheets(1).Cells(counter + 1, 134).Value = Workbooks(FileName).Worksheets(1).Cells(i, 123).Value
        Workbooks(FileName).Worksheets(1).Cells(counter + 1, 135).Value = Workbooks(FileName).Worksheets(1).Cells(i, 124).Value
        Workbooks(FileName).Worksheets(1).Cells(counter + 1, 136).Value = Workbooks(FileName).Worksheets(1).Cells(i, 125).Value
        Workbooks(FileName).Worksheets(1).Cells(counter + 1, 137).Value = Workbooks(FileName).Worksheets(1).Cells(i, 126).Value
        counter = counter + 1
    End If
Next i
Workbooks(FileName).Worksheets(1).Range("DR6:DV105").Value = ""
Workbooks(FileName).Worksheets(1).Range("DR6:DV105").Value = ThisWorkbook.Worksheets(2).Range("EC6:EG105").Value
Workbooks(FileName).Worksheets(1).Range("EC6:EG105").Value = ""
ThisWorkbook.Worksheets(2).Range("DR6:DV105").Value = ""
ThisWorkbook.Worksheets(2).Range("DR6:DV105").Value = ThisWorkbook.Worksheets(2).Range("EC6:EG105").Value
ThisWorkbook.Worksheets(2).Range("EC6:EG105").Value = ""

ThisWorkbook.Worksheets(2).Range("EB6:EB105").ClearContents
Call TekilUnvanlar

'Sıralama ekle komutuna da uygulanacak.
SayHedef = Workbooks(FileName).Worksheets(1).Range("DR1000").End(xlUp).Row
If SayHedef > 6 Then
    'A'dan Z'ye sırala ve böylece arada bulunan boş satırları da kaldır.
    Workbooks(FileName).Worksheets(1).Unprotect Password:="123"
    ThisWorkbook.Unprotect "123"
    ThisWorkbook.Worksheets(2).Unprotect Password:="123"
    ThisWorkbook.Worksheets(2).Visible = True
    'Sort A to Z
    Workbooks(FileName).Worksheets(1).Range("DR" & 6 & ":DV" & SayHedef).Sort key1:=Workbooks(FileName).Worksheets(1).Range("DR" & 6 & ":DR" & SayHedef), order1:=xlAscending, Header:=xlNo
    ThisWorkbook.Worksheets(2).Range("DR" & 6 & ":DV" & SayHedef).Sort key1:=ThisWorkbook.Worksheets(2).Range("DR" & 6 & ":DR" & SayHedef), order1:=xlAscending, Header:=xlNo
    ThisWorkbook.Worksheets(2).Visible = False
    Workbooks(FileName).Worksheets(1).Protect Password:="123" ', DrawingObjects:=False
    ThisWorkbook.Worksheets(2).Protect Password:="123" ', DrawingObjects:=False
    ThisWorkbook.Protect "123"
End If

'Etiketi düzeltmen gerekiyorsa düzelt
UserName = Environ("UserProfile")
UserName = UCase(Right(UserName, 7))
If ItemUserName = UserName Then
    LblUserNameBilgi.Caption = " User " & UserName & " has been successfully removed from the system."
    LblUserNameBilgi.ForeColor = RGB(60, 100, 180)
End If

'Açık dropdown kapat
Call ModuleSystemSettings.DropDownKapat

Workbooks(FileName).Save
'ThisWorkbook.Save

MsgBox "The session named '" & ItemUserName & "' has been successfully removed from the system with the following details:" & vbNewLine & vbNewLine & _
"Removed Session Name: " & ItemUserName & vbNewLine & _
"Full Name: " & ItemName & vbNewLine & _
"Title: " & ItemName1 & vbNewLine & _
"Registration No.: " & ItemName2 & vbNewLine & _
"Phone: " & ItemName3, vbOKOnly + vbInformation, "Enterprise Document Automation System"


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

SayHedef = Worksheets(2).Range("DR1000").End(xlUp).Row
If SayHedef < 6 Then
    SayHedef = 6
End If
Worksheets(2).Range("EB6:EB105").ClearContents
    
Unload Me

End Sub


Private Sub UserForm_Initialize()
Dim i As Long, WsSKP As Object
Dim ClrLab As MSForms.Control
Dim SayHedef As Integer, UserName As String
Dim ItemBul As Range

ThisWorkbook.Activate

For Each ClrLab In core_initials_UI.Controls
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

core_initials_UI.BackColor = RGB(230, 230, 230) 'YENİ



SayHedef = Worksheets(2).Range("DR1000").End(xlUp).Row
If SayHedef < 6 Then
    SayHedef = 6
End If
Worksheets(2).Range("EB6:EB105").ClearContents

CheckBoxDuzelt.Value = True

'Username kontrolü
UserName = Environ("UserProfile")
'UserName = UCase(Replace(Replace(Right(UserName, 7), "i", "I"), "ı", "I"))
UserName = UCase(Right(UserName, 7))
Set ItemBul = Worksheets(2).Range("DR6:DR1000").Find(What:=UserName, SearchDirection:=xlNext, _
                SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
If Not ItemBul Is Nothing Then
    LblUserNameBilgi.Caption = " User " & UserName & " is already registered in the system."
    LblUserNameBilgi.ForeColor = RGB(19, 117, 71)

    ComboUserName.Value = Worksheets(2).Cells(ItemBul.Row, 122).Value
    ComboAdSoyad.Value = Worksheets(2).Cells(ItemBul.Row, 123).Value
    ComboUnvan.Value = Worksheets(2).Cells(ItemBul.Row, 124).Value
    ComboSicil.Value = Worksheets(2).Cells(ItemBul.Row, 125).Value
    ComboTel.Value = Worksheets(2).Cells(ItemBul.Row, 126).Value
    
Else
    LblUserNameBilgi.Caption = " User " & UserName & " is not registered in the system."
    LblUserNameBilgi.ForeColor = RGB(60, 100, 180)
End If
    

'MsgBox UserName

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
Dim yukseklik As Variant, genislik As Variant
Dim Rep As Variant
Dim SayHedef As Integer

SayHedef = Worksheets(2).Range("DR1000").End(xlUp).Row
If SayHedef < 6 Then
    SayHedef = 6
End If
Worksheets(2).Range("EB6:EB105").ClearContents

yukseklik = Me.Height
Rep = Me.Width
Do
DoEvents
Rep = Rep - 60
Call timeout(0.01)
    If Rep > 60 Then
        core_initials_UI.Width = Rep
        yukseklik = yukseklik - 60
        core_initials_UI.Height = yukseklik
        If yukseklik <= 60 Then
            yukseklik = 60
            core_initials_UI.Height = yukseklik
        End If
    ElseIf Rep <= 50 Then
        Rep = 50
        core_initials_UI.Width = Rep
        yukseklik = yukseklik - 50
        core_initials_UI.Height = yukseklik
        If yukseklik <= 50 Then
            yukseklik = 50
            core_initials_UI.Height = yukseklik
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
