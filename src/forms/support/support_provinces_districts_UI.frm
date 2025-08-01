VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} support_provinces_districts_UI 
   Caption         =   "Enterprise Document Automation System"
   ClientHeight    =   8895
   ClientLeft      =   30
   ClientTop       =   375
   ClientWidth     =   13920
   OleObjectBlob   =   "support_provinces_districts_UI.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "support_provinces_districts_UI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Abort As Boolean, ComboChangeControl As Boolean


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

Private Sub ComboIl_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

Me.ComboIl.DropDown

End Sub

Private Sub ComboIl_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    
    On Error Resume Next

   Select Case KeyCode
        Case 38  'Yukarı
            If ComboIl.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                ComboIl.ListIndex = ComboIl.ListIndex - 1
            End If
            Me.ComboIl.DropDown
            
        Case 40 'Aşağı
            If ComboIl.ListIndex = ComboIl.ListCount - 1 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                ComboIl.ListIndex = ComboIl.ListIndex + 1
            End If
            Me.ComboIl.DropDown
    End Select
    Abort = False

End Sub

Private Sub ComboIl_Change()
Dim ItemBul As Range

If ComboChangeControl = False Then

    ComboIlce.Value = "" 'Bunu diğer il combolarına da uyarla !!!! İl degişince ilçe değerini kaldır.
    ComboIlKodu.Value = ""
    ComboIlceKodu.Value = ""
    
    'İl kodunu bul
    On Error Resume Next
    Set ItemBul = Worksheets(2).Range("F6:F95").Find(What:=ComboIl.Value, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    If Not ItemBul Is Nothing Then
        '
    Else
        GoTo Bos
    End If
    ComboIlKodu.Value = Worksheets(2).Cells(ItemBul.Row, 5).Value
    
    
    'İlçeleri bul
    On Error GoTo 0
    'ComboIlçe seçimlerini İl seçimine göre göster.
    On Error GoTo Bos
    ComboIlce.RowSource = Replace(ComboIl.Value, " ", "_")
    'ComboIl.DropDown
    GoTo Son
    
Bos:
    ComboIlce.RowSource = ""
    ComboIlKodu.Value = ""
    ComboIlceKodu.Value = ""
    
Son:
    
    ComboIl.DropDown
End If


End Sub


Private Sub ComboIlce_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

Me.ComboIlce.DropDown

End Sub

Private Sub ComboIlce_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    On Error Resume Next

   Select Case KeyCode
        Case 38  'Yukarı
            If ComboIlce.ListIndex <= 0 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                ComboIlce.ListIndex = ComboIlce.ListIndex - 1
            End If
            Me.ComboIlce.DropDown
            
        Case 40 'Aşağı
            If ComboIlce.ListIndex = ComboIlce.ListCount - 1 Then KeyCode = 0
            Abort = True
            If Not KeyCode = 0 Then
                KeyCode = 0
                ComboIlce.ListIndex = ComboIlce.ListIndex + 1
            End If
            Me.ComboIlce.DropDown
    End Select
    Abort = False
    
End Sub

Private Sub ComboIlce_Change()
Dim IlBul As Range, IlceBul As Range, IlEsleyicisi As Integer

If ComboChangeControl = False Then

    ComboIlceKodu.Value = ""
    
    'Il
    Set IlBul = ThisWorkbook.Worksheets(2).Columns("F").Find(What:=ComboIl.Value, SearchDirection:=xlNext, _
                SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    If Not IlBul Is Nothing Then
        IlEsleyicisi = ThisWorkbook.Worksheets(2).Range("C" & IlBul.Row).Value
    Else
        GoTo Out
    End If
    'Ilce
    Set IlceBul = ThisWorkbook.Worksheets(2).Columns(IlEsleyicisi + 6).Find(What:=ComboIlce.Value, SearchDirection:=xlNext, _
                SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
    If Not IlceBul Is Nothing Then
        ComboIlceKodu.Value = Worksheets(2).Cells(IlceBul.Row, 4).Value
    Else
        GoTo Out
    End If
    
    ComboIlce.DropDown

Out:

End If

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
Private Sub ComboIl_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
Call SetComboBoxHook(ComboIl) 'Open scrollable with mouse
End Sub
Private Sub ComboIlce_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
Call SetComboBoxHook(ComboIlce) 'Open scrollable with mouse
End Sub
Private Sub ComboIlKodu_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
Call SetComboBoxHook(ComboIlKodu) 'Open scrollable with mouse
End Sub

Private Sub LblIl_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub
Private Sub LblIlKodu_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub
Private Sub LblIlce_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Call ColorChangerGenel
End Sub
Private Sub LblIlceKodu_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
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
Dim AutoPath As String, DestTarget As String, OpenControl As String
Dim FileName As String, i As Integer, j As Integer, IlSayHedef As Long, IlceSayHedef As Long
Dim IlBul As Range, IlceBul As Range, IlKoduBul As Range, IlceKoduBul As Range, IlEsleyicisi As Integer
Dim IlKodu As String, IlceKodu As String, IlStr As String, IlceStr As String
Dim WsTanimlar As Object, Bilgi As Variant, WsIlKodu As Variant, WsIlceKodu As Variant
Dim WsIlStr As String, WsIlceStr As String
Dim NameDuzenleyici As String


Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.EnableEvents = False


ComboChangeControl = True

AutoPath = ThisWorkbook.Path
DestTarget = AutoPath & "\System Files\System Definitions\"
FileName = "Definitions.xlsx"

IlStr = ComboIl.Value
IlKodu = ComboIlKodu.Value
IlceStr = ComboIlce.Value
IlceKodu = ComboIlceKodu.Value

IlSayHedef = ThisWorkbook.Worksheets(2).Range("F1000").End(xlUp).Row + 1
If IlSayHedef < 5 Then
    IlSayHedef = 6
End If
If IlSayHedef > 95 Then
    MsgBox "The definition area is full, therefore your operation cannot be completed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Out
End If


'İl kontrol
If IlStr = "" Then
    MsgBox "Province field cannot be left empty.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Out
End If
If IlKodu = "" Then
    MsgBox "Province code field cannot be left empty.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Out
End If


'''''''''''''''''''''''
ThisWorkbook.Unprotect "123"
ThisWorkbook.Worksheets(2).Unprotect Password:="123"

OpenControl = IsFileOpen(DestTarget & FileName)
If OpenControl = True Then 'Açıksa
    Workbooks(FileName).Close SaveChanges:=True
ElseIf OpenControl = False Then
    '
End If
Workbooks.Open (DestTarget & FileName)
Workbooks(FileName).Worksheets(1).Activate
Set WsTanimlar = Workbooks(FileName)

WsTanimlar.Worksheets(1).Unprotect Password:="123"
''''''''''''''''''''''''

IlStr = WorksheetFunction.Proper(IlStr)
For i = 1 To 50
    IlKodu = Replace(IlKodu, " ", "")
Next i
'MsgBox IlKodu
If Left(IlKodu, 1) = "0" Then
    IlKodu = Mid(IlKodu, 2, Len(IlKodu) - 1)
    'MsgBox IlKodu
End If

'On Error Resume Next
Set IlBul = ThisWorkbook.Worksheets(2).Range("F6:F95").Find(What:=IlStr, SearchDirection:=xlNext, _
            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
If Not IlBul Is Nothing Then
'    MsgBox "İl eşleşmesi sağlandı.", vbOKOnly + vbInformation, "Enterprise Document Automation System"
    IlEsleyicisi = ThisWorkbook.Worksheets(2).Cells(IlBul.Row, 3).Value
    
    'İl kodu kontrol
    If IlKodu = ThisWorkbook.Worksheets(2).Cells(IlBul.Row, 5).Value Then 'İl kodu eşleşti
'        MsgBox "İl kodu eşleşmesi sağlandı.", vbOKOnly + vbInformation, "Enterprise Document Automation System"
    Else 'İl kodu eşleşmedi
'        MsgBox "İl kodu eşleşmesi sağlanamadı.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        If Len(IlKodu) = 1 Then
            IlKodu = "0" & IlKodu
            'MsgBox IlKodu
        End If
        Set IlKoduBul = ThisWorkbook.Worksheets(2).Range("E6:E95").Find(What:=IlKodu, SearchDirection:=xlNext, _
                        SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
        If Not IlKoduBul Is Nothing Then 'İl kodu kullanımda
            WsIlStr = ThisWorkbook.Worksheets(2).Cells(IlKoduBul.Row, 6).Value
            MsgBox "Province code " & IlKodu & " is already used by " & WsIlStr & " and is therefore not available.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
            GoTo Son
        Else 'İl kodu boşta
            WsIlKodu = ThisWorkbook.Worksheets(2).Cells(IlBul.Row, 5).Value
            If Len(WsIlKodu) = 1 Then
                WsIlKodu = "0" & WsIlKodu
                'MsgBox WsIlKodu
            End If
            Bilgi = MsgBox("Province code " & IlKodu & " is available. The province code of " & IlStr & " will be changed from " & WsIlKodu & " to " & IlKodu & "." & vbNewLine & _
                    "Click """ & "Yes" & """ to confirm the change, or """ & "No" & """ to cancel.", vbYesNo + vbInformation, "Enterprise Document Automation System")
            If Bilgi = vbYes Then
            
                'İl kodunu değiştir.
                ThisWorkbook.Worksheets(2).Cells(IlBul.Row, 5).Value = IlKodu
                WsTanimlar.Worksheets(1).Cells(IlBul.Row, 5).Value = IlKodu
                
                MsgBox "The province code of " & IlStr & " has been successfully changed from " & WsIlKodu & " to " & IlKodu & ".", vbOKOnly + vbInformation, "Enterprise Document Automation System"

            ElseIf Bilgi = vbNo Then
                GoTo Son
            End If
        End If
    End If
    
    
    'İlçe kontrol
    If IlceStr <> "" Then 'İlçe alanı dolu
        IlceStr = WorksheetFunction.Proper(IlceStr)
        For i = 1 To 50
            IlceKodu = Replace(IlceKodu, " ", "")
        Next i
        'MsgBox IlceKodu
        If Left(IlceKodu, 1) = "0" Then
            IlceKodu = Mid(IlceKodu, 2, Len(IlceKodu) - 1)
            'MsgBox IlceKodu
        End If
        Set IlceBul = ThisWorkbook.Worksheets(2).Columns(IlEsleyicisi + 6).Find(What:=IlceStr, SearchDirection:=xlNext, _
                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
        If Not IlceBul Is Nothing Then 'İlçe bulundu
'            MsgBox "İlçe eşleşmesi sağlandı.", vbOKOnly + vbInformation, "Enterprise Document Automation System"
            'İlçe kodu kontrol
            If IlceKodu <> "" Then 'İlçe eşleşmesi sağlandı ve ilçe kodu alanı dolu
                If IlceKodu = ThisWorkbook.Worksheets(2).Cells(IlceBul.Row, 4).Value Then 'İlçe ve ilçe kodu eşleşmesi sağlandı
'                    MsgBox "İlçe kodu eşleşmesi sağlandı.", vbOKOnly + vbInformation, "Enterprise Document Automation System"
                Else 'İlçe eşleşmesi sağlandı, ama ilçe kodu eşleşmedi (İlçe kodu değiştirilmek isteniyor.)
'                    MsgBox "İlçe kodu eşleşmesi sağlanamadı.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
                    IlEsleyicisi = ThisWorkbook.Worksheets(2).Cells(IlBul.Row, 3).Value

                    If Len(IlceKodu) = 1 Then
                        IlceKodu = "0" & IlceKodu
                        'MsgBox IlceKodu
                    End If
                    Set IlceKoduBul = ThisWorkbook.Worksheets(2).Range("D6:D55").Find(What:=IlceKodu, SearchDirection:=xlNext, _
                                    SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
                    If Not IlceKoduBul Is Nothing Then 'İlçe kodu tanımlama alanı içinde
                        WsIlceStr = ThisWorkbook.Worksheets(2).Cells(IlceKoduBul.Row, IlEsleyicisi + 6).Value
                        If WsIlceStr <> "" Then 'İlçe kodu kulanımda
                            MsgBox "The district code " & IlceKodu & " is already in use by " & WsIlceStr & " and is therefore not available.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
                            GoTo Son
                        Else 'İlçe kodu boşta
                            WsIlceKodu = ThisWorkbook.Worksheets(2).Cells(IlceBul.Row, 4).Value
                            If Len(WsIlceKodu) = 1 Then
                                WsIlceKodu = "0" & WsIlceKodu
                                'MsgBox WsIlceKodu
                            End If
                            Bilgi = MsgBox("The district code " & IlceKodu & " is available. The district code for " & IlceStr & " will be changed from " & WsIlceKodu & " to " & IlceKodu & ". Click " & """" & "Yes" & """" & " to confirm the change, or " & """" & "No" & """" & " to cancel.", vbYesNo + vbInformation, "Enterprise Document Automation System")
                            If Bilgi = vbYes Then
                            
                                'İlçe kodunu değiştir.
                                ThisWorkbook.Worksheets(2).Cells(IlceKoduBul.Row, IlEsleyicisi + 6).Value = IlceStr
                                WsTanimlar.Worksheets(1).Cells(IlceKoduBul.Row, IlEsleyicisi + 6).Value = IlceStr
                                ThisWorkbook.Worksheets(2).Cells(IlceBul.Row, IlEsleyicisi + 6).Value = ""
                                WsTanimlar.Worksheets(1).Cells(IlceBul.Row, IlEsleyicisi + 6).Value = ""

                                MsgBox "The district code for " & IlceStr & " has been successfully changed from " & WsIlceKodu & " to " & IlceKodu & ".", vbOKOnly + vbInformation, "Enterprise Document Automation System"

                            ElseIf Bilgi = vbNo Then
                                GoTo Son
                            End If
                        End If
                    Else 'İlçe kodu tanımlama alanı dışında
                        MsgBox "The district code " & IlceKodu & " is outside the allowed range and therefore not available. Please enter a district code between 1 and 50.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
                        GoTo Son
                    End If

                End If
            Else 'İlçe eşleşmesi sağlandı, ama ilçe kodu boş
                MsgBox "The district code field cannot be left empty while the district name field is filled.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
                GoTo Son
            End If
        Else 'İlçe bulunamadı (Yeni ilçe eklenmek isteniyor.)
            If IlceKodu <> "" Then
'                MsgBox "İlçe eşleşmesi sağlanamadı.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
                IlEsleyicisi = ThisWorkbook.Worksheets(2).Cells(IlBul.Row, 3).Value
                
                If Len(IlceKodu) = 1 Then
                    IlceKodu = "0" & IlceKodu
                    'MsgBox IlceKodu
                End If
                Set IlceKoduBul = ThisWorkbook.Worksheets(2).Range("D6:D55").Find(What:=IlceKodu, SearchDirection:=xlNext, _
                                SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
                If Not IlceKoduBul Is Nothing Then 'İlçe kodu tanımlama alanı içinde
                    WsIlceStr = ThisWorkbook.Worksheets(2).Cells(IlceKoduBul.Row, IlEsleyicisi + 6).Value
                    If WsIlceStr <> "" Then 'İlçe kodu kulanımda
                        MsgBox "The district code " & IlceKodu & " is already in use by " & WsIlceStr & " and therefore cannot be used.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
                        GoTo Son
                    Else 'İlçe kodu boşta
                        WsIlceKodu = ThisWorkbook.Worksheets(2).Cells(IlceKoduBul.Row, 4).Value
                        If Len(WsIlceKodu) = 1 Then
                            WsIlceKodu = "0" & WsIlceKodu
                            'MsgBox WsIlceKodu
                        End If
                        Bilgi = MsgBox("A new district named " & IlceStr & " with the district code " & IlceKodu & " will be added to the province " & IlStr & ". To confirm the operation, click " & """" & "Yes" & """" & "; to cancel, click " & """" & "No" & """" & ".", vbYesNo + vbInformation, "Enterprise Document Automation System")
                        If Bilgi = vbYes Then
                        
                            'Yeni ilçe ekle
                            ThisWorkbook.Worksheets(2).Cells(IlceKoduBul.Row, IlEsleyicisi + 6).Value = IlceStr
                            WsTanimlar.Worksheets(1).Cells(IlceKoduBul.Row, IlEsleyicisi + 6).Value = IlceStr

                            MsgBox "A new district named " & IlceStr & " with the district code " & IlceKodu & " has been successfully added to the province " & IlStr & ".", vbOKOnly + vbInformation, "Enterprise Document Automation System"
    
                        ElseIf Bilgi = vbNo Then
                            GoTo Son
                        End If
                    End If
                Else 'İlçe kodu tanımlama alanı dışında
                    MsgBox "The district code " & IlceKodu & " is outside the allowed definition range and therefore unavailable. Please enter a district code between 1 and 50.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
                    GoTo Son
                End If
            Else
                MsgBox "The district code field cannot be left blank while the district name field is filled.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
                GoTo Son
            End If
        End If
    Else 'İlçe alanı boş
        '
    End If
    
Else
'    MsgBox "İl eşleşmesi sağlanamadı.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"

    'İlçe kontrol
    If IlceStr <> "" Then 'İlçe alanı dolu

        If Len(IlceKodu) = 1 Then
            IlceKodu = "0" & IlceKodu
            'MsgBox IlceKodu
        End If
        Set IlceKoduBul = ThisWorkbook.Worksheets(2).Range("D6:D55").Find(What:=IlceKodu, SearchDirection:=xlNext, _
                        SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
        If Not IlceKoduBul Is Nothing Then 'İlçe kodu tanımlama alanı içinde

            If Len(IlKodu) = 1 Then
                IlKodu = "0" & IlKodu
                'MsgBox IlKodu
            End If
            Set IlKoduBul = ThisWorkbook.Worksheets(2).Range("E6:E95").Find(What:=IlKodu, SearchDirection:=xlNext, _
                            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not IlKoduBul Is Nothing Then 'İl kodu kullanımda
                WsIlStr = ThisWorkbook.Worksheets(2).Cells(IlKoduBul.Row, 6).Value
                MsgBox "The province code " & IlKodu & " is already in use by " & WsIlStr & " and therefore is not available.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
                GoTo Son
            Else 'İl kodu boşta
                Bilgi = MsgBox("A new province named " & IlStr & " with province code " & IlKodu & " and a new district named " & IlceStr & " with district code " & IlceKodu & " will be added. To confirm the operation, click " & """" & "Yes" & """" & ", or click " & """" & "No" & """" & " to cancel.", vbYesNo + vbInformation, "Enterprise Document Automation System")
                If Bilgi = vbYes Then
                    'İl kodu ekle
                    ThisWorkbook.Worksheets(2).Cells(IlSayHedef, 5).Value = IlKodu
                    WsTanimlar.Worksheets(1).Cells(IlSayHedef, 5).Value = IlKodu
                    'İli dikeyde ekle
                    ThisWorkbook.Worksheets(2).Cells(IlSayHedef, 6).Value = IlStr
                    WsTanimlar.Worksheets(1).Cells(IlSayHedef, 6).Value = IlStr
                    'İli yatayda ekle
                    IlEsleyicisi = ThisWorkbook.Worksheets(2).Cells(IlSayHedef, 3).Value
                    ThisWorkbook.Worksheets(2).Cells(5, IlEsleyicisi + 6).Value = IlStr
                    WsTanimlar.Worksheets(1).Cells(5, IlEsleyicisi + 6).Value = IlStr
                    'Name manager ekle
                    NameDuzenleyici = Replace(ThisWorkbook.Worksheets(2).Cells(5, IlEsleyicisi + 6).Value, " ", "_")
                    ThisWorkbook.Names.Add name:=NameDuzenleyici, _
                    RefersTo:=ThisWorkbook.Worksheets(2).Range(ThisWorkbook.Worksheets(2).Cells(6, IlEsleyicisi + 6), _
                    ThisWorkbook.Worksheets(2).Cells(55, IlEsleyicisi + 6))
                    'Yeni ilçe ekle
                    ThisWorkbook.Worksheets(2).Cells(IlceKoduBul.Row, IlEsleyicisi + 6).Value = IlceStr
                    WsTanimlar.Worksheets(1).Cells(IlceKoduBul.Row, IlEsleyicisi + 6).Value = IlceStr

                    MsgBox "A new province named " & IlStr & " with province code " & IlKodu & " and a new district named " & IlceStr & " with district code " & IlceKodu & " have been successfully added to the system.", vbOKOnly + vbInformation, "Enterprise Document Automation System"

                ElseIf Bilgi = vbNo Then
                    GoTo Son
                End If
            End If
            
        Else 'İlçe kodu tanımlama alanı dışında
            MsgBox "The district code " & IlceKodu & " is outside the allowed definition range and is therefore not available. Please enter a district code between 1 and 50.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
            GoTo Son
        End If

    Else 'İlçe alanı boş

        If Len(IlKodu) = 1 Then
            IlKodu = "0" & IlKodu
            'MsgBox IlKodu
        End If
        Set IlKoduBul = ThisWorkbook.Worksheets(2).Range("E6:E95").Find(What:=IlKodu, SearchDirection:=xlNext, _
                        SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
        If Not IlKoduBul Is Nothing Then 'İl kodu kullanımda
            WsIlStr = ThisWorkbook.Worksheets(2).Cells(IlKoduBul.Row, 6).Value
            MsgBox "The province code " & IlKodu & " is already used by " & WsIlStr & " and is therefore not available.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
            GoTo Son
        Else 'İl kodu boşta
            Bilgi = MsgBox("A new province named " & IlStr & " with the code " & IlKodu & " will be added. To confirm the operation, click " & """" & "Yes" & """" & "; to cancel, click " & """" & "No" & """" & ".", vbYesNo + vbInformation, "Enterprise Document Automation System")
            If Bilgi = vbYes Then
                'İl kodu ekle
                ThisWorkbook.Worksheets(2).Cells(IlSayHedef, 5).Value = IlKodu
                WsTanimlar.Worksheets(1).Cells(IlSayHedef, 5).Value = IlKodu
                'İli dikeyde ekle
                ThisWorkbook.Worksheets(2).Cells(IlSayHedef, 6).Value = IlStr
                WsTanimlar.Worksheets(1).Cells(IlSayHedef, 6).Value = IlStr
                'İli yatayda ekle
                IlEsleyicisi = ThisWorkbook.Worksheets(2).Cells(IlSayHedef, 3).Value
                ThisWorkbook.Worksheets(2).Cells(5, IlEsleyicisi + 6).Value = IlStr
                WsTanimlar.Worksheets(1).Cells(5, IlEsleyicisi + 6).Value = IlStr
                'Name manager ekle
                NameDuzenleyici = Replace(ThisWorkbook.Worksheets(2).Cells(5, IlEsleyicisi + 6).Value, " ", "_")
                ThisWorkbook.Names.Add name:=NameDuzenleyici, _
                RefersTo:=ThisWorkbook.Worksheets(2).Range(ThisWorkbook.Worksheets(2).Cells(6, IlEsleyicisi + 6), _
                ThisWorkbook.Worksheets(2).Cells(55, IlEsleyicisi + 6))
                
                MsgBox "A new province named " & IlStr & " with the code " & IlKodu & " has been successfully added to the system.", vbOKOnly + vbInformation, "Enterprise Document Automation System"
            
            ElseIf Bilgi = vbNo Then
                GoTo Son
            End If
        End If

    End If

End If


Son:

ThisWorkbook.Worksheets(2).Protect Password:="123"
ThisWorkbook.Protect "123"

'ThisWorkbook.Save

OpenControl = IsFileOpen(DestTarget & FileName)
If OpenControl = True Then 'Açıksa
    WsTanimlar.Worksheets(1).Protect Password:="123"
    WsTanimlar.Save
    Workbooks(FileName).Close SaveChanges:=True
ElseIf OpenControl = False Then
    WsTanimlar.Save
End If

Out:

ComboChangeControl = False

ThisWorkbook.Activate

Application.ScreenUpdating = True
Application.DisplayAlerts = True
Application.EnableEvents = True

End Sub

Private Sub LabelKaldir_Click()
Dim AutoPath As String, DestTarget As String, OpenControl As String
Dim FileName As String, i As Integer, j As Integer, IlSayHedef As Long, IlceSayHedef As Long
Dim IlBul As Range, IlceBul As Range, IlKoduBul As Range, IlceKoduBul As Range, IlEsleyicisi As Integer
Dim IlKodu As String, IlceKodu As String, IlStr As String, IlceStr As String
Dim WsTanimlar As Object, Bilgi As Variant, WsIlKodu As Variant, WsIlceKodu As Variant
Dim WsIlStr As String, WsIlceStr As String, Sifre As Variant
Dim NameDuzenleyici As String


Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.EnableEvents = False


ComboChangeControl = True

AutoPath = ThisWorkbook.Path
DestTarget = AutoPath & "\System Files\System Definitions\"
FileName = "Definitions.xlsx"

IlStr = ComboIl.Value
IlKodu = ComboIlKodu.Value
IlceStr = ComboIlce.Value
IlceKodu = ComboIlceKodu.Value

'Province check
If IlStr = "" Then
    MsgBox "The province field cannot be left empty.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Out
End If
If IlKodu = "" Then
    MsgBox "The province code field cannot be left empty.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Out
End If


''''''''''''''''''''''''''
ThisWorkbook.Unprotect "123"
ThisWorkbook.Worksheets(2).Unprotect Password:="123"

OpenControl = IsFileOpen(DestTarget & FileName)
If OpenControl = True Then 'Açıksa
    Workbooks(FileName).Close SaveChanges:=True
ElseIf OpenControl = False Then
    '
End If
Workbooks.Open (DestTarget & FileName)
Workbooks(FileName).Worksheets(1).Activate
Set WsTanimlar = Workbooks(FileName)

WsTanimlar.Worksheets(1).Unprotect Password:="123"
''''''''''''''''''''''''''

IlStr = WorksheetFunction.Proper(IlStr)
For i = 1 To 50
    IlKodu = Replace(IlKodu, " ", "")
Next i
'MsgBox IlKodu
If Left(IlKodu, 1) = "0" Then
    IlKodu = Mid(IlKodu, 2, Len(IlKodu) - 1)
    'MsgBox IlKodu
End If

'On Error Resume Next
Set IlBul = ThisWorkbook.Worksheets(2).Range("F6:F95").Find(What:=IlStr, SearchDirection:=xlNext, _
            SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
If Not IlBul Is Nothing Then
'    MsgBox "İl eşleşmesi sağlandı.", vbOKOnly + vbInformation, "Enterprise Document Automation System"
    IlEsleyicisi = ThisWorkbook.Worksheets(2).Cells(IlBul.Row, 3).Value
    
    'İl kodu kontrol
    If IlKodu = ThisWorkbook.Worksheets(2).Cells(IlBul.Row, 5).Value Then 'İl kodu eşleşti
'        MsgBox "İl kodu eşleşmesi sağlandı.", vbOKOnly + vbInformation, "Enterprise Document Automation System"
    
         'İlçe kontrol
        If IlceStr <> "" Then 'İlçe alanı dolu
            IlceStr = WorksheetFunction.Proper(IlceStr)
            For i = 1 To 50
                IlceKodu = Replace(IlceKodu, " ", "")
            Next i
            'MsgBox IlceKodu
            If Left(IlceKodu, 1) = "0" Then
                IlceKodu = Mid(IlceKodu, 2, Len(IlceKodu) - 1)
                'MsgBox IlceKodu
            End If
            Set IlceBul = ThisWorkbook.Worksheets(2).Columns(IlEsleyicisi + 6).Find(What:=IlceStr, SearchDirection:=xlNext, _
                        SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
            If Not IlceBul Is Nothing Then 'İlçe bulundu
'                MsgBox "İlçe eşleşmesi sağlandı.", vbOKOnly + vbInformation, "Enterprise Document Automation System"
                
                'İlçe kodu kontrol
                If IlceKodu <> "" Then 'İlçe eşleşmesi sağlandı ve ilçe kodu alanı dolu
                    If IlceKodu = ThisWorkbook.Worksheets(2).Cells(IlceBul.Row, 4).Value Then 'İlçe ve ilçe kodu eşleşmesi sağlandı
'                        MsgBox "İlçe kodu eşleşmesi sağlandı.", vbOKOnly + vbInformation, "Enterprise Document Automation System"
                        
                        'MsgBox "İlçe silme prosedürü buraya gelecek"
                        
                        Sifre = InputBox(Prompt:="The district named " & IlceStr & " with district code " & IlceKodu & " under the province named " & IlStr & " with province code " & IlKodu & " will be deleted." & vbNewLine & vbNewLine & "To initiate the deletion process, please enter the password: 123", Title:="Enterprise Document Automation System")
                        If Sifre = "123" Then
                            
                            ThisWorkbook.Worksheets(2).Cells(IlceBul.Row, IlEsleyicisi + 6).Value = ""
                            WsTanimlar.Worksheets(1).Cells(IlceBul.Row, IlEsleyicisi + 6).Value = ""
                            ComboIlceKodu.Value = ""
                            
                            MsgBox "The district '" & IlceStr & "' (Code: " & IlceKodu & ") under the province '" & IlStr & "' (Code: " & IlKodu & ") has been successfully removed from the system.", vbOKOnly + vbInformation, "Enterprise Document Automation System"
                        
                        ElseIf Sifre = vbCancel Then
                            GoTo Son
                        ElseIf Sifre <> "" And Sifre <> "123" Then
                            MsgBox "Incorrect password. The province/district deletion process could not be started.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
                            GoTo Son
                        End If
                        
                    Else 'İlçe eşleşmesi sağlandı, ama ilçe kodu eşleşmedi
                        MsgBox "District code mismatch. The deletion process could not be completed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
                        GoTo Son
                    End If
                Else 'İlçe eşleşmesi sağlandı, ama ilçe kodu boş
                    MsgBox "District code cannot be left blank when the district name is filled.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
                    GoTo Son
                End If
            Else 'İlçe bulunamadı
                If IlceKodu <> "" Then
                    MsgBox "District match could not be found. Deletion cannot proceed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
                    GoTo Son
                Else
                    MsgBox "District code is required when the district name is provided.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
                    GoTo Son
                End If
            End If
        Else 'İlçe alanı boş (İl ve tüm ilçeler silinecek)
            
            'MsgBox "İl ve tüm ilçeler silinecek"

            Sifre = InputBox(Prompt:="The province '" & IlStr & "' (Code: " & IlKodu & ") and all its associated districts will be permanently deleted." & vbNewLine & vbNewLine & "To proceed, please enter the password '123'.", Title:="Enterprise Document Automation System")
            If Sifre = "123" Then
                IlSayHedef = ThisWorkbook.Worksheets(2).Range(ThisWorkbook.Worksheets(2).Cells(1000, IlEsleyicisi + 6), _
                            ThisWorkbook.Worksheets(2).Cells(1000, IlEsleyicisi + 6)).End(xlUp).Row
                If IlSayHedef < 6 Then
                    IlSayHedef = 6
                End If
                
                'Name manager sil
                On Error Resume Next
                NameDuzenleyici = Replace(ThisWorkbook.Worksheets(2).Cells(5, IlEsleyicisi + 6).Value, " ", "_")
                ThisWorkbook.Names(NameDuzenleyici).Delete
                On Error GoTo 0
                'Verileri sil
                ThisWorkbook.Worksheets(2).Range(ThisWorkbook.Worksheets(2).Cells(5, IlEsleyicisi + 6), _
                            ThisWorkbook.Worksheets(2).Cells(IlSayHedef, IlEsleyicisi + 6)).Value = ""
                ThisWorkbook.Worksheets(2).Cells(IlBul.Row, 6).Value = ""
                ThisWorkbook.Worksheets(2).Cells(IlBul.Row, 5).Value = ""
                
                WsTanimlar.Worksheets(1).Range(WsTanimlar.Worksheets(1).Cells(5, IlEsleyicisi + 6), _
                            WsTanimlar.Worksheets(1).Cells(IlSayHedef, IlEsleyicisi + 6)).Value = ""
                WsTanimlar.Worksheets(1).Cells(IlBul.Row, 6).Value = ""
                WsTanimlar.Worksheets(1).Cells(IlBul.Row, 5).Value = ""
                
                ComboIl.Value = ""
                ComboIlKodu.Value = ""
                ComboIlce.Value = ""
                ComboIlceKodu.Value = ""
                
                MsgBox "The province '" & IlStr & "' (Code: " & IlKodu & ") and all associated districts have been successfully removed from the system.", vbOKOnly + vbInformation, "Enterprise Document Automation System"
            
            ElseIf Sifre = vbCancel Then
                GoTo Son
            ElseIf Sifre <> "" And Sifre <> "123" Then
                MsgBox "Incorrect password. The province/district deletion process was not initiated.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
                GoTo Son
            End If
 
        End If
    
    Else 'İl kodu eşleşmedi
        MsgBox "Deletion failed: Province code mismatch.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
        GoTo Son
    End If
    
Else 'İl eşleşmedi
    MsgBox "Deletion failed: Province match not found.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If


Son:

ThisWorkbook.Worksheets(2).Protect Password:="123"
ThisWorkbook.Protect "123"

'ThisWorkbook.Save

OpenControl = IsFileOpen(DestTarget & FileName)
If OpenControl = True Then 'Açıksa
    WsTanimlar.Worksheets(1).Protect Password:="123"
    WsTanimlar.Save
    Workbooks(FileName).Close SaveChanges:=True
ElseIf OpenControl = False Then
    WsTanimlar.Save
End If

Out:

ComboChangeControl = False

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

For Each ClrLab In support_provinces_districts_UI.Controls
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

support_provinces_districts_UI.BackColor = RGB(230, 230, 230) 'YENİ

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
        support_provinces_districts_UI.Width = Rep
        yukseklik = yukseklik - 60
        support_provinces_districts_UI.Height = yukseklik
        If yukseklik <= 60 Then
            yukseklik = 60
            support_provinces_districts_UI.Height = yukseklik
        End If
    ElseIf Rep <= 50 Then
        Rep = 50
        support_provinces_districts_UI.Width = Rep
        yukseklik = yukseklik - 50
        support_provinces_districts_UI.Height = yukseklik
        If yukseklik <= 50 Then
            yukseklik = 50
            support_provinces_districts_UI.Height = yukseklik
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

