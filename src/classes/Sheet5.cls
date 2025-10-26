VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Dim BtnRapor As Object, Abort As Boolean
Public OpenWordTakip As Boolean

Private Sub RaporOlustur_Click()
Dim IlkSiraBul As Range, SonSiraBul As Range, IlkSira As Long, SonSira As Long
Dim j As Long, i As Long

Application.ScreenUpdating = False
Application.EnableEvents = False

ThisWorkbook.Unprotect "123"
Worksheets(5).Unprotect Password:="123"

j = ActiveCell.Row
For i = ActiveCell.Row To 7 Step -1
    If Cells(i, 5).Value <> "" Then
        GoTo SiraNoBulundu
    Else
        j = j - 1
    End If
Next i
SiraNoBulundu:
If j < 7 Then
    MsgBox "The serial number could not be found, so the operation cannot be completed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

Set IlkSiraBul = Range("FG7:FG100000").Find(What:=Cells(j, 5).Value, SearchDirection:=xlNext, _
                SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
Set SonSiraBul = Range("FH7:FH100000").Find(What:=Cells(j, 5).Value, SearchDirection:=xlNext, _
                SearchOrder:=xlByRows, LookIn:=xlValues, LookAt:=xlWhole)
If Not IlkSiraBul Is Nothing Then
    IlkSira = IlkSiraBul.Row
Else
    MsgBox "The serial number could not be found, so the operation cannot be completed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If
If Not SonSiraBul Is Nothing Then
    SonSira = SonSiraBul.Row
Else
    MsgBox "The serial number could not be found, so the operation cannot be completed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If
    

If Cells(IlkSira, 12).Value = "Point1" Then
    Call ModuleInit.Rapor3_1Baslat
ElseIf Cells(IlkSira, 12).Value = "Point2" Or Cells(IlkSira, 12).Value = "Point3" Then
    Call ModuleInit.FinansalBirimBaslat
End If


Son:

Worksheets(5).Protect Password:="123" ', DrawingObjects:=False
ThisWorkbook.Protect "123"


Application.EnableEvents = True
Application.ScreenUpdating = True

End Sub

Private Sub Worksheet_SelectionChange(ByVal target As Range)
Dim i As Long

Worksheets(5).Unprotect Password:="123"

'Rapor butonlarını oluştur...
On Error Resume Next
Set BtnRapor = ActiveSheet.OLEObjects("RaporOlustur")
If (ActiveCell.Value = "ü" And ActiveCell.Row > 6 And Cells(ActiveCell.Row, 217).Value <> "" And ActiveCell.Column = 7) Or _
   (ActiveCell.Value = "ü" And ActiveCell.Row > 6 And ActiveCell.Value <> "" And _
   (ActiveCell.Column = 6 Or ActiveCell.Column = 8 Or ActiveCell.Column = 9 Or ActiveCell.Column = 10 Or _
   ActiveCell.Column = 11)) Then
    If BtnRapor Is Nothing Then
        With ActiveSheet.OLEObjects.Add(ClassType:="Forms.CommandButton.1", _
                                      Width:=Intersect(target, Cells(ActiveCell.Row, ActiveCell.Column)).Width, _
                                      Height:=Intersect(target, Cells(ActiveCell.Row, ActiveCell.Column)).Height)
        .name = "RaporOlustur"
        End With
        Set BtnRapor = ActiveSheet.OLEObjects("RaporOlustur")
    End If
    'Butonları isimlendir
    If ActiveCell.Column = 7 Then
        If Cells(ActiveCell.Row, 5).Value = "" Then
            For i = ActiveCell.Row To 7 Step -1
                If Cells(i, 5).Value <> "" Then
                    ActiveSheet.OLEObjects("RaporOlustur").Object.Caption = Cells(i, 5).Value & "-" & Cells(6, ActiveCell.Column).Value & " " & Cells(ActiveCell.Row, 13).Value
                    GoTo RaporNoDonguSon
                End If
            Next i
        Else
            ActiveSheet.OLEObjects("RaporOlustur").Object.Caption = Cells(ActiveCell.Row, 5).Value & "-" & Cells(6, ActiveCell.Column).Value & " " & Cells(ActiveCell.Row, 13).Value
        End If
    Else
        ActiveSheet.OLEObjects("RaporOlustur").Object.Caption = Cells(ActiveCell.Row, 5).Value & "-" & Cells(6, ActiveCell.Column).Value
    End If
RaporNoDonguSon:
    If target Is Nothing Then
        BtnRapor.Visible = False
        BtnRapor.Enabled = False
        GoTo AtlaBtnRapor
    End If
   
    With BtnRapor
        .Top = Intersect(target, ActiveCell).Top
        .Left = Intersect(target, ActiveCell).Left
        .Width = Intersect(target, Cells(ActiveCell.Row, ActiveCell.Column)).Width
        .Height = Intersect(target, Cells(ActiveCell.Row, ActiveCell.Column)).Height
        .Enabled = True
        .Visible = True
        .Activate
        '.TakeFocusOnClick = True
    End With
Else
    BtnRapor.Visible = False
    BtnRapor.Enabled = False
    GoTo AtlaBtnRapor
End If

AtlaBtnRapor:


Worksheets(5).Protect Password:="123"

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
