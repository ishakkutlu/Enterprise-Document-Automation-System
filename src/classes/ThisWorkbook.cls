VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Dim BtnRapor As Object

Private Sub Workbook_Activate()
Application.DisplayFormulaBar = False
ActiveSheet.DisplayPageBreaks = False
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
Dim AutoPath As String, DestOperasyon As String, DestOpUserFolderName As String, DestOpUserFolder As String
Dim fso As Object

Call ModuleSystemSettings.TimeStop

Application.EnableEvents = False
Application.EnableEvents = True

'Gizli Word uygulamasını kapat
Call CloseWordApp

'___________________Operation klasöründeki kalıntıları temizle.

'Pathfinder...
AutoPath = ThisWorkbook.Path
DestOperasyon = AutoPath & "\System Files\Operation\"
DestOpUserFolderName = "Operation-" & Left(ThisWorkbook.name, InStr(ThisWorkbook.name, ".") - 1)
DestOpUserFolder = DestOperasyon & DestOpUserFolderName

'System Files klasör adını kontrol et.
If Not Dir(AutoPath & "\System Files\", vbDirectory) <> vbNullString Then
    MsgBox AutoPath & "\System Files\" & " directory cannot be accessed. The folder named 'System Files' may have been renamed or deleted.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

'Operation klasörü adını kontrol et.
If Not Dir(DestOperasyon, vbDirectory) <> vbNullString Then
    MsgBox DestOperasyon & " directory cannot be accessed. The folder named 'Operation' may have been renamed or the 'System Files' folder may have been deleted.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

'Operation kalıntılarını temizle
On Error Resume Next
Set fso = CreateObject("Scripting.FileSystemObject")
fso.DeleteFolder DestOpUserFolder


''''''''''''''''''''''

'Button ve picture artıklarını temizle
ThisWorkbook.Unprotect "123"

On Error Resume Next

ThisWorkbook.Worksheets(3).Visible = True
ThisWorkbook.Worksheets(4).Visible = True
ThisWorkbook.Worksheets(5).Visible = True

ThisWorkbook.Worksheets(3).Activate
ThisWorkbook.Worksheets(3).Unprotect Password:="123"
Call ModuleSystemSettings.DeleteAllMyObjectsKapat
ThisWorkbook.Worksheets(3).Protect Password:="123"

ThisWorkbook.Worksheets(4).Activate
ThisWorkbook.Worksheets(4).Unprotect Password:="123"
Call ModuleSystemSettings.DeleteAllMyObjectsKapat
ThisWorkbook.Worksheets(4).Protect Password:="123"

ThisWorkbook.Worksheets(5).Activate
ThisWorkbook.Worksheets(5).Unprotect Password:="123"
Call ModuleSystemSettings.DeleteAllMyObjectsKapat

ThisWorkbook.Worksheets(1).Visible = True

ThisWorkbook.Worksheets(3).Visible = False
ThisWorkbook.Worksheets(4).Visible = False
ThisWorkbook.Worksheets(5).Visible = False

ThisWorkbook.Worksheets(1).Activate

ThisWorkbook.Worksheets(5).Protect Password:="123"

On Error GoTo 0

ThisWorkbook.Protect "123"

Son:


ThisWorkbook.Save

End Sub

Private Sub CloseWordApp()
Dim ObjWordx As Object

    On Error Resume Next
    Set ObjWordx = GetObject(, "Word.Application")
    Set ObjWordx = GetObject(, "Word.Application")
    Set ObjWordx = GetObject(, "Word.Application")
    Set ObjWordx = GetObject(, "Word.Application")
    Set ObjWordx = GetObject(, "Word.Application")
    Set ObjWordx = GetObject(, "Word.Application")
    Set ObjWordx = GetObject(, "Word.Application")
    Set ObjWordx = GetObject(, "Word.Application")
    Set ObjWordx = GetObject(, "Word.Application")
    Set ObjWordx = GetObject(, "Word.Application")

    ObjWordx.Visible = True
    ObjWordx.Quit SaveChanges:=True
    Set ObjWordx = Nothing
    
End Sub

Private Sub Workbook_Deactivate()
Application.DisplayFormulaBar = True
End Sub

Private Sub Workbook_Open()
Dim ws As Worksheet, Wb As Workbook
Dim SourceFile As String, AutoPath As String, KontrolFile As String
Dim j As Integer, i As Integer

Dim DestOperasyon As String, DestOpUserFolderName As String, DestOpUserFolder As String
Dim fso As Object

Dim DestTarget As String, OpenControl As String, Tanimlar As String
Dim FileName As String, SayHedef As Long


Application.DisplayAlerts = False
Application.EnableEvents = False
Application.ScreenUpdating = False

Set Wb = ThisWorkbook
If Wb.ReadOnly = True Then
    MsgBox "The file is open on another computer, so the module cannot be accessed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    ThisWorkbook.Close SaveChanges:=False
    GoTo Son
End If

AutoPath = ThisWorkbook.Path
SourceFile = AutoPath & "\System Files\System Templates\Report 2 Templates\"

DestTarget = AutoPath & "\System Files\System Definitions\"
FileName = "Definitions.xlsx"
Tanimlar = DestTarget & FileName

'___________________________________________Bu dosyanın aynı kullanıcıda açık olup olmadığını kontrol et.

'Pathfinder...
DestOperasyon = AutoPath & "\System Files\Operation\"
DestOpUserFolderName = "Operation-" & Left(ThisWorkbook.name, InStr(ThisWorkbook.name, ".") - 1) '& " " & Format(Now(), "ddmmyyyyhhmmss")
DestOpUserFolder = DestOperasyon & DestOpUserFolderName

'Check System Files folder name
If Not Dir(AutoPath & "\System Files\", vbDirectory) <> vbNullString Then
    MsgBox AutoPath & "\System Files\" & " directory cannot be accessed, and the application will automatically close. The folder name 'System Files' may have been changed or the application may have been opened in recovery mode." & vbNewLine & vbNewLine & _
    "After the application closes automatically, please close any remaining windows and reopen the application.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    ThisWorkbook.Close SaveChanges:=False
    GoTo Son
End If

'Check Operation folder name
If Not Dir(DestOperasyon, vbDirectory) <> vbNullString Then
    MsgBox DestOperasyon & " directory cannot be accessed. The folder named 'Operation' may have been renamed or the 'System Files' folder may have been deleted.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

'Check System Definitions folder name
If Not Dir(DestTarget, vbDirectory) <> vbNullString Then
    MsgBox DestTarget & " directory cannot be accessed. The folder named 'System Definitions' may have been renamed or deleted.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

'Check System Definitions file
If Not Dir(Tanimlar, vbDirectory) <> vbNullString Then
    MsgBox Tanimlar & " directory cannot be accessed. The file named 'System Definitions' may have been renamed or deleted.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

'Operation klasörü içinde kullanıcı modülü klasörü yoksa oluştur.
If Dir(DestOpUserFolder, vbDirectory) <> vbNullString Then
    MsgBox "Enterprise Document Automation System is already running and cannot be opened again. For data security, the application will be closed automatically. If you think the application is not open on any computer, please try reopening it.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"

    'Op kalıntılarını temizle
    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.DeleteFolder DestOpUserFolder
    
    ThisWorkbook.Close SaveChanges:=True
    GoTo Son
End If
MkDir DestOpUserFolder


'___________________________________________


'Klasör isimlerini kontrol et.
If Not Dir(SourceFile, vbDirectory) <> vbNullString Then
    MsgBox SourceFile & " directory cannot be accessed. The names of folders and/or files in this directory might have been changed.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    GoTo Son
End If

'______________________________________

'SİSTEM SAYFASI OLMAYANLARI SİL (Başlangıç)____________________________

ThisWorkbook.Unprotect "123"
Call ModuleSystemSettings.DelSheet
ThisWorkbook.Protect "123"

'SİSTEM SAYFASI OLMAYANLARI SİL (Bitiş) ____________________________


'System Definitions dosyasını içe aktar.

OpenControl = IsFileOpen(DestTarget & FileName)
If OpenControl = True Then 'Açıksa kapat
    Workbooks(FileName).Close SaveChanges:=True
ElseIf OpenControl = False Then
    '
End If
Workbooks.Open (DestTarget & FileName)
Workbooks(FileName).Worksheets(1).Activate

Workbooks(FileName).Worksheets(1).Unprotect Password:="123"
ThisWorkbook.Unprotect "123"
ThisWorkbook.Worksheets(2).Unprotect Password:="123"
ThisWorkbook.Worksheets(2).Visible = True

ThisWorkbook.Worksheets(2).Range("C" & 1 & ":EE" & 305).Value = Workbooks(FileName).Worksheets(1).Range("C" & 1 & ":EE" & 305).Value

ThisWorkbook.Worksheets(2).Visible = False
Workbooks(FileName).Worksheets(1).Protect Password:="123"
ThisWorkbook.Worksheets(2).Protect Password:="123"
ThisWorkbook.Protect "123"

'Workbooks(FileName).Save
OpenControl = IsFileOpen(DestTarget & FileName)
If OpenControl = True Then 'Açıksa kapat
    Workbooks(FileName).Close SaveChanges:=False
ElseIf OpenControl = False Then
    '
End If

ThisWorkbook.Activate

'________________________________________

'Rapor çeşitleri için rapor System Templatesından verileri çek.
'Öncekini temizle
ThisWorkbook.Worksheets(2).Range("DL6:DL305").Value = ""
'Açılır liste için verileri çek
j = 6
KontrolFile = Dir(SourceFile & "*.docm")
Do While KontrolFile <> ""
    If j < 306 Then
        ThisWorkbook.Worksheets(2).Range("DL" & j) = Mid(KontrolFile, 1, Len(KontrolFile) - 5)
        j = j + 1
    End If
    'MsgBox Mid(KontrolFile, 1, Len(KontrolFile) - 5)
    KontrolFile = Dir()
Loop

ThisWorkbook.Unprotect "123"


''''''''''''''''''''''
On Error Resume Next

ThisWorkbook.Worksheets(3).Visible = True
ThisWorkbook.Worksheets(4).Visible = True
ThisWorkbook.Worksheets(5).Visible = True

ThisWorkbook.Worksheets(3).Activate
ThisWorkbook.Worksheets(3).Unprotect Password:="123"
'Call ModuleSystemSettings.DeleteAllMyObjects
Set BtnRapor = Nothing
Set BtnRapor = ActiveSheet.OLEObjects("RaporOlustur")
If BtnRapor Is Nothing Then
    With ActiveSheet.OLEObjects.Add(ClassType:="Forms.CommandButton.1", _
                                      Width:=Intersect(Cells(7, 6), Cells(7, 6)).Width, _
                                      Height:=Intersect(Cells(7, 6), Cells(7, 6)).Height)
    .name = "RaporOlustur"
    End With
    Set BtnRapor = ActiveSheet.OLEObjects("RaporOlustur")
    ActiveSheet.OLEObjects("RaporOlustur").Object.Caption = "NewButton"
    BtnRapor.Visible = False
    BtnRapor.Enabled = False
End If
ThisWorkbook.Worksheets(3).Protect Password:="123"


ThisWorkbook.Worksheets(4).Activate
ThisWorkbook.Worksheets(4).Unprotect Password:="123"
'Call ModuleSystemSettings.DeleteAllMyObjects
Set BtnRapor = Nothing
Set BtnRapor = ActiveSheet.OLEObjects("RaporOlustur")
If BtnRapor Is Nothing Then
    With ActiveSheet.OLEObjects.Add(ClassType:="Forms.CommandButton.1", _
                                      Width:=Intersect(Cells(7, 6), Cells(7, 6)).Width, _
                                      Height:=Intersect(Cells(7, 6), Cells(7, 6)).Height)
    .name = "RaporOlustur"
    End With
    Set BtnRapor = ActiveSheet.OLEObjects("RaporOlustur")
    ActiveSheet.OLEObjects("RaporOlustur").Object.Caption = "NewButton"
    BtnRapor.Visible = False
    BtnRapor.Enabled = False
End If
ThisWorkbook.Worksheets(4).Protect Password:="123"


ThisWorkbook.Worksheets(5).Activate
ThisWorkbook.Worksheets(5).Unprotect Password:="123"
'Call ModuleSystemSettings.DeleteAllMyObjects
Set BtnRapor = Nothing
Set BtnRapor = ActiveSheet.OLEObjects("RaporOlustur")
If BtnRapor Is Nothing Then
    With ActiveSheet.OLEObjects.Add(ClassType:="Forms.CommandButton.1", _
                                      Width:=Intersect(Cells(7, 6), Cells(7, 6)).Width, _
                                      Height:=Intersect(Cells(7, 6), Cells(7, 6)).Height)
    .name = "RaporOlustur"
    End With
    Set BtnRapor = ActiveSheet.OLEObjects("RaporOlustur")
    ActiveSheet.OLEObjects("RaporOlustur").Object.Caption = "NewButton"
    BtnRapor.Visible = False
    BtnRapor.Enabled = False
End If
ThisWorkbook.Worksheets(5).Protect Password:="123"


ThisWorkbook.Worksheets(1).Visible = True
ThisWorkbook.Worksheets(1).Activate

On Error GoTo 0
''''''''''''''''''''''



For Each ws In ThisWorkbook.Worksheets
    On Error Resume Next
    ws.Visible = False
    Worksheets(1).Visible = True
Next

Call CloseWordApp
Call OpenWordApp

Worksheets(1).Activate
With ActiveWindow
    .DisplayHorizontalScrollBar = False
    .DisplayVerticalScrollBar = False
End With

Call ModuleAsset.hideheadings

ThisWorkbook.Protect "123"


Call ModuleSystemSettings.TimeSetting


Son:

ThisWorkbook.Protect "123"

Application.DisplayAlerts = True
Application.EnableEvents = True
Application.ScreenUpdating = True

End Sub

Private Sub OpenWordApp()
Dim objWord As Object

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
    
End Sub
