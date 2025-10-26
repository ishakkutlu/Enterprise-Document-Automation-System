VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Dim BtnRapor As Object, Abort As Boolean
Public OpenWordTakip As Boolean

Private Sub RaporOlustur_Click()

Application.ScreenUpdating = False
Application.EnableEvents = False

ThisWorkbook.Unprotect "123"
Worksheets(4).Unprotect Password:="123"

Call ModuleInit.Rapor2Baslat

Worksheets(4).Protect Password:="123" ', DrawingObjects:=False
ThisWorkbook.Protect "123"


Application.EnableEvents = True
Application.ScreenUpdating = True

End Sub

Private Sub Worksheet_SelectionChange(ByVal target As Range)
Dim i As Long

Worksheets(4).Unprotect Password:="123"

'Rapor butonlarını oluştur...
On Error Resume Next
Set BtnRapor = ActiveSheet.OLEObjects("RaporOlustur")
If (ActiveCell.Value = "ü" And ActiveCell.Row > 6 And Cells(ActiveCell.Row, 67).Value <> "" And (ActiveCell.Column = 7 Or ActiveCell.Column = 14)) Or _
   (ActiveCell.Value = "ü" And ActiveCell.Row > 6 And ActiveCell.Value <> "" And _
   (ActiveCell.Column = 6 Or ActiveCell.Column = 8 Or ActiveCell.Column = 9 Or ActiveCell.Column = 10 Or _
   ActiveCell.Column = 13 Or ActiveCell.Column = 15 Or ActiveCell.Column = 16 Or ActiveCell.Column = 17 Or ActiveCell.Column = 18 Or ActiveCell.Column = 19 Or ActiveCell.Column = 20 Or ActiveCell.Column = 21 Or ActiveCell.Column = 22)) Then
    If BtnRapor Is Nothing Then
        With ActiveSheet.OLEObjects.Add(ClassType:="Forms.CommandButton.1", _
                                      Width:=Intersect(target, Cells(ActiveCell.Row, ActiveCell.Column)).Width, _
                                      Height:=Intersect(target, Cells(ActiveCell.Row, ActiveCell.Column)).Height)
        .name = "RaporOlustur"
        End With
        Set BtnRapor = ActiveSheet.OLEObjects("RaporOlustur")
    End If
    'Butonları isimlendir
    If ActiveCell.Column = 7 Or ActiveCell.Column = 14 Then
        If Cells(ActiveCell.Row, 5).Value = "" Then
            For i = ActiveCell.Row To 7 Step -1
                If Cells(i, 5).Value <> "" Then
                    ActiveSheet.OLEObjects("RaporOlustur").Object.Caption = Cells(i, 5).Value & "-" & Cells(6, ActiveCell.Column).Value & " " & Cells(ActiveCell.Row, 11).Value
                    GoTo RaporNoDonguSon
                End If
            Next i
        Else
            ActiveSheet.OLEObjects("RaporOlustur").Object.Caption = Cells(ActiveCell.Row, 5).Value & "-" & Cells(6, ActiveCell.Column).Value & " " & Cells(ActiveCell.Row, 11).Value
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

Worksheets(4).Protect Password:="123"

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

