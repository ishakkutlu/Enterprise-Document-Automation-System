VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} support_item_type_notes_UI 
   Caption         =   "Enterprise Document Automation System"
   ClientHeight    =   8895
   ClientLeft      =   30
   ClientTop       =   375
   ClientWidth     =   13920
   OleObjectBlob   =   "support_item_type_notes_UI.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "support_item_type_notes_UI"
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
'LabelGetir
If LabelGetir.BackColor <> RGB(254, 254, 254) Then
LabelGetir.BackColor = RGB(254, 254, 254)
LabelGetir.ForeColor = RGB(30, 30, 30)
End If

End Sub

Private Sub ComboOgeTuru_Change()

    If ComboOgeTuru.ListIndex = -1 And ComboOgeTuru.Value <> "" Then
       ComboOgeTuru.Value = ""
       GoTo Son
    End If
    If ComboOgeTuru.Value <> "" Then
        ComboOgeTuru.SelStart = 0
        ComboOgeTuru.SelLength = Len(ComboOgeTuru.Value)
    End If

Son:
    NotText.Value = ""
    
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

Private Sub ComboOgeTuru_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
Call SetComboBoxHook(ComboOgeTuru) 'Open scrollable with mouse
End Sub

Private Sub LabelGetir_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
LabelGetir.BackColor = RGB(60, 100, 180) 'RGB(60, 100, 180) 'RGB(60, 100, 180)
LabelGetir.ForeColor = RGB(255, 255, 255)
End Sub


Private Sub NotText_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
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
Dim FileName As String, SayHedef As Long, ItemName As String, j As Integer
Dim fso As Object, NotTxtFile As Object, HedefFile As String

Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.EnableEvents = False

AutoPath = ThisWorkbook.Path
DestTarget = AutoPath & "\System Files\System Templates\Item Notes\"

'Item type must be selected.
If ComboOgeTuru.Value = "" Then
    MsgBox "Note could not be added because no item type was selected.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

'Note field must not be empty.
If NotText.Value = "" Then
    MsgBox "Note could not be added because the note field is empty.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

FileName = ComboOgeTuru.Value
HedefFile = DestTarget & FileName & ".txt"

'Dosyanın olup olmadığını kontrol et.
If Dir(HedefFile, vbDirectory) <> vbNullString Then
    MsgBox "A note named " & FileName & " has already been created, so the add operation cannot be completed. If you want to update a note related to an item type, please use the update button.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If


'Not için text dosyası oluştur.
Set fso = CreateObject("Scripting.FileSystemObject")
FileName = ComboOgeTuru.Value
Set NotTxtFile = fso.CreateTextFile(DestTarget & FileName & ".txt", True, False) 'ANSI için False olmalı
NotTxtFile.Write NotText.Value
NotTxtFile.Close

MsgBox "Your note for the " & FileName & " has been successfully defined.", vbOKOnly + vbInformation, "Enterprise Document Automation System"
NotText.Value = ""
ComboOgeTuru.Value = ""

Son:

ThisWorkbook.Activate

Application.ScreenUpdating = True
Application.DisplayAlerts = True
Application.EnableEvents = True

End Sub

Private Sub LabelKaldir_Click()
Dim a() As Variant, i As Variant
Dim AutoPath As String, DestTarget As String, OpenControl As String
Dim FileName As String, SayHedef As Long, ItemName As String, j As Integer
Dim HedefFile As String, fso As Object, NotTxtFile As Object

Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.EnableEvents = False

AutoPath = ThisWorkbook.Path
DestTarget = AutoPath & "\System Files\System Templates\Item Notes\"

'Öğe türü seçimi yapılmalıdır.
If ComboOgeTuru.Value = "" Then
    MsgBox "Note removal could not be completed because no item type was selected.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    NotText.Value = ""
    GoTo Son
End If

FileName = ComboOgeTuru.Value
HedefFile = DestTarget & FileName & ".txt"
'Dosyanın olup olmadığını kontrol et.
If Not Dir(HedefFile, vbDirectory) <> vbNullString Then
    MsgBox "Since no note has been created for the " & FileName & ", the removal operation cannot be completed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If
    
Kill HedefFile

MsgBox "Your note for the " & FileName & " has been successfully removed.", vbOKOnly + vbInformation, "Enterprise Document Automation System"
NotText.Value = ""
ComboOgeTuru.Value = ""

Son:

ThisWorkbook.Activate

Application.ScreenUpdating = True
Application.DisplayAlerts = True
Application.EnableEvents = True

End Sub

Private Sub LabelGetir_Click()
Dim a() As Variant, i As Variant
Dim AutoPath As String, DestTarget As String, OpenControl As String
Dim FileName As String, SayHedef As Long, ItemName As String, j As Integer
Dim HedefFile As String, fso As Object, NotTxtFile As Object, TextLine As String
Dim NotIcerigi As String

Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.EnableEvents = False

AutoPath = ThisWorkbook.Path
DestTarget = AutoPath & "\System Files\System Templates\Item Notes\"

'Öğe türü seçimi yapılmalıdır.
If ComboOgeTuru.Value = "" Then
    MsgBox "Note retrieval could not be completed because no item type was selected.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

If NotText.Value = "" Then
    'Notu çağır
    FileName = ComboOgeTuru.Value
    HedefFile = DestTarget & FileName & ".txt"
    'Dosyanın olup olmadığını kontrol et.
    If Not Dir(HedefFile, vbDirectory) <> vbNullString Then
    MsgBox "Since no note has been created for the " & FileName & ", the data retrieval operation cannot be completed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    Open HedefFile For Input As #1
        Do Until EOF(1)
            Line Input #1, TextLine
            NotIcerigi = NotIcerigi & TextLine
        Loop
    Close #1
    
    NotText.Value = NotIcerigi
End If

Son:

ThisWorkbook.Activate

Application.ScreenUpdating = True
Application.DisplayAlerts = True
Application.EnableEvents = True


End Sub

Private Sub LabelGuncelle_Click()
Dim a() As Variant, i As Variant
Dim AutoPath As String, DestTarget As String, OpenControl As String
Dim FileName As String, SayHedef As Long, ItemName As String, j As Integer
Dim HedefFile As String, fso As Object, NotTxtFile As Object, TextLine As String
Dim NotIcerigi As String

Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.EnableEvents = False

AutoPath = ThisWorkbook.Path
DestTarget = AutoPath & "\System Files\System Templates\Item Notes\"

'Öğe türü seçimi yapılmalıdır.
If ComboOgeTuru.Value = "" Then
    MsgBox "Note update could not be completed because no item type was selected.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

FileName = ComboOgeTuru.Value
HedefFile = DestTarget & FileName & ".txt"

'Dosyanın olup olmadığını kontrol et.
If Not Dir(HedefFile, vbDirectory) <> vbNullString Then
    MsgBox "Since a note named " & FileName & " has not been created before, the update operation cannot be completed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

If NotText.Value = "" Then
    FileName = ComboOgeTuru.Value
MsgBox "Since no note could be found for the " & FileName & ", the update operation cannot be completed." & vbNewLine & vbNewLine & _
        "First, select an item type and then click the Load Data button to load the defined note (if any) into the text box. After editing the note in the text box, click the Update button.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
Else
    'Notu güncelle.
    Set fso = CreateObject("Scripting.FileSystemObject")
    FileName = ComboOgeTuru.Value
    Set NotTxtFile = fso.CreateTextFile(DestTarget & FileName & ".txt", True, False) 'ANSI için False olmalı
    NotTxtFile.Write NotText.Value
    NotTxtFile.Close
    MsgBox "Your note for the " & FileName & " has been successfully updated.", vbOKOnly + vbInformation, "Enterprise Document Automation System"
    NotText.Value = ""
    ComboOgeTuru.Value = ""
End If

Son:

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

For Each ClrLab In support_item_type_notes_UI.Controls
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
LabelKaldir.BackColor = RGB(225, 235, 245)
LabelKaldir.ForeColor = RGB(30, 30, 30)

support_item_type_notes_UI.BackColor = RGB(230, 230, 230) 'YENİ

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
        support_item_type_notes_UI.Width = Rep
        yukseklik = yukseklik - 60
        support_item_type_notes_UI.Height = yukseklik
        If yukseklik <= 60 Then
            yukseklik = 60
            support_item_type_notes_UI.Height = yukseklik
        End If
    ElseIf Rep <= 50 Then
        Rep = 50
        support_item_type_notes_UI.Width = Rep
        yukseklik = yukseklik - 50
        support_item_type_notes_UI.Height = yukseklik
        If yukseklik <= 50 Then
            yukseklik = 50
            support_item_type_notes_UI.Height = yukseklik
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


