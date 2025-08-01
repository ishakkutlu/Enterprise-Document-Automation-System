Attribute VB_Name = "ModuleSystemSettings"
Option Explicit
Public colLabelEvent As Collection
Public colLabels As Collection
Public bSecondDate As Boolean
Public sActiveDay As String
Public lDays As Long
Public lFirstDay As Long
Public lStartPos As Long
Public lSelMonth As Long
Public lSelYear As Long
Public lSelMonth1 As Long
Public lSelYear1 As Long
Public sayac As Integer
Public CalTarih As String
Public CalTarihTakip As String
Public EnglishMonths As Variant
Public sMonth As String

Private Declare PtrSafe Function EnumPrinters Lib "winspool.drv" Alias "EnumPrintersA" _
    (ByVal flags As Long, ByVal name As String, ByVal Level As Long, _
    pPrinterEnum As LongPtr, ByVal cdBuf As Long, pcbNeeded As LongPtr, _
    pcReturned As LongPtr) As Long
Private Declare PtrSafe Function PtrToStr Lib "kernel32" Alias "lstrcpyA" _
    (ByVal RetVal As String, ByVal Ptr As LongPtr) As Long
Private Declare PtrSafe Function StrLen Lib "kernel32" Alias "lstrlenA" _
    (ByVal Ptr As LongPtr) As Long

Const PRINTER_ENUM_CONNECTIONS = &H4
Const PRINTER_ENUM_LOCAL = &H2

'''''''''''''''''''''''''''''''''''''''''''''

Public RepYukseklik As Variant
Dim CloseTime As Date
Public xSaat As Variant, xDakika As Variant, xSaniye As Variant, xDurum As String
Public GlobalResetKapsami As Integer, ComboGetirAktar As String

Public Declare PtrSafe Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Declare PtrSafe Function MsgBoxTimeout _
    Lib "user32" _
    Alias "MessageBoxTimeoutA" ( _
        ByVal hwnd As LongPtr, _
        ByVal lpText As String, _
        ByVal lpCaption As String, _
        ByVal wType As VbMsgBoxStyle, _
        ByVal wlange As Long, _
        ByVal dwTimeout As Long) _
As Long


Function ListPrinters() As Variant

Dim bSuccess As Boolean
Dim iBufferRequired As Long
Dim iBufferSize As Long
Dim iBuffer() As Long
Dim iEntries As Long
Dim iIndex As Long
Dim strPrinterName As String
Dim iDummy As Long
Dim iDriverBuffer() As Long
Dim StrPrinters() As String

iBufferSize = 3072

ReDim iBuffer((iBufferSize \ 4) - 1) As Long

'EnumPrinters will return a value False if the buffer is not big enough
bSuccess = EnumPrinters(PRINTER_ENUM_CONNECTIONS Or _
        PRINTER_ENUM_LOCAL, vbNullString, _
        1, iBuffer(0), iBufferSize, iBufferRequired, iEntries)

If Not bSuccess Then
    If iBufferRequired > iBufferSize Then
        iBufferSize = iBufferRequired
        Debug.Print "iBuffer too small. Trying again with "; _
        iBufferSize & " bytes."
        ReDim iBuffer(iBufferSize \ 4) As Long
    End If
    'Try again with new buffer
    bSuccess = EnumPrinters(PRINTER_ENUM_CONNECTIONS Or _
            PRINTER_ENUM_LOCAL, vbNullString, _
            1, iBuffer(0), iBufferSize, iBufferRequired, iEntries)
End If

If Not bSuccess Then
    'Enumprinters returned False
    MsgBox "Error enumerating printers."
    Exit Function
Else
    'Enumprinters returned True, use found printers to fill the array
    ReDim StrPrinters(iEntries - 1)
    For iIndex = 0 To iEntries - 1
        'Get the printername
        strPrinterName = Space$(StrLen(iBuffer(iIndex * 4 + 2)))
        iDummy = PtrToStr(strPrinterName, iBuffer(iIndex * 4 + 2))
        StrPrinters(iIndex) = strPrinterName
    Next iIndex
End If

ListPrinters = StrPrinters

End Function

Function IsBounded(vArray As Variant) As Boolean

    'If the variant passed to this function is an array, the function will return True;
    'otherwise it will return False
    On Error Resume Next
    IsBounded = IsNumeric(UBound(vArray))

End Function

Function FindPrinter(ByVal PrinterName As String) As String
    Dim Arr As Variant
    Dim Device As Variant
    Dim Devices As Variant
    Dim printer As String
    Dim RegObj As Object
    Dim RegValue As String
    Const HKEY_CURRENT_USER = &H80000001

    Set RegObj = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
    RegObj.enumValues HKEY_CURRENT_USER, "Software\Microsoft\Windows NT\CurrentVersion\Devices", Devices, Arr

    For Each Device In Devices
        RegObj.getstringvalue HKEY_CURRENT_USER, "Software\Microsoft\Windows NT\CurrentVersion\Devices", Device, RegValue
        printer = Device & " on " & Split(RegValue, ",")(1)
        If StrComp(Device, PrinterName, vbTextCompare) = 0 Then
            FindPrinter = printer
            Exit Function
        End If
    Next Device
End Function


Function IsWorkBookOpen(FileName As String)
    Dim ff As Long, ErrNo As Long

    On Error Resume Next
    ff = FreeFile()
    Open FileName For Input Lock Read As #ff
    Close ff
    ErrNo = Err
    On Error GoTo 0

    Select Case ErrNo
    Case 0:    IsWorkBookOpen = False
    Case 70:   IsWorkBookOpen = True
    Case Else: Error ErrNo
    End Select
End Function

Function IsFileOpen(FileName As String)
    Dim ff As Long, ErrNo As Long

    On Error Resume Next
    ff = FreeFile()
    Open FileName For Input Lock Read As #ff
    Close ff
    ErrNo = Err
    On Error GoTo 0

    Select Case ErrNo
    Case 0:    IsFileOpen = False
    Case 70:   IsFileOpen = True
    Case 53:   GoTo Son
    Case Else: Error ErrNo
    End Select
Son:
End Function

Public Function FormatEnglishDate(inputDate As Date) As String
    Dim d As Integer, m As Integer, y As Integer
    d = Day(inputDate)
    m = Month(inputDate)
    y = Year(inputDate)
    
    If IsEmpty(EnglishMonths) Then Init_EnglishMonths
    
    FormatEnglishDate = d & " " & EnglishMonths(m) & " " & y
End Function

Public Sub Init_EnglishMonths()
    EnglishMonths = Array("", "January", "February", "March", "April", "May", "June", _
                          "July", "August", "September", "October", "November", "December")
End Sub

Public Sub SignToday()
Dim i As Integer, Gun As Control, StrTdy As String, ObjTdy As Object
Dim MonthNm As Integer, YearNm As Integer
Dim ClrLabDay As Control

YearNm = suppport_calendar_UI.cmbYear.Value
' MonthNm = Month(DateValue("01 " & suppport_calendar_UI.cmbMonth.Value & " 2015")) 'Turkish Months

'English Months
Call Init_EnglishMonths
For i = 1 To 12
    If suppport_calendar_UI.cmbMonth.Value = EnglishMonths(i) Then
        MonthNm = i
        Exit For
    End If
Next i

If Format(Date, "yyyy") = YearNm And Format(Date, "mm") = MonthNm Then
    For Each Gun In suppport_calendar_UI.Frame1.Controls
        If TypeName(Gun) = "Label" And Not Gun.ForeColor = &HE0E0E0 Then
         If Gun < 10 Then
            If Gun = Format(Date, "d") Then
                GoTo Bulundu
            End If
         Else
            If Gun = Format(Date, "dd") Then
                GoTo Bulundu
            End If
         End If
        End If
    Next
Else
    For Each ClrLabDay In suppport_calendar_UI.Controls
        If TypeName(ClrLabDay) = "Label" Then
            ClrLabDay.BackColor = RGB(180, 180, 180)
        End If
    Next
End If

GoTo Atla
Bulundu:
StrTdy = Gun.name
Set ObjTdy = Gun
ObjTdy.BackColor = RGB(94, 140, 199) 'RGB(249, 194, 19)

Atla:

End Sub

Sub DeleteAllMyObjectsKapat()
Dim shp As Shape

'ThisWorkbook.Worksheets(5).Unprotect Password:="123"

On Error Resume Next
For Each shp In ActiveSheet.Shapes
If Not shp.name = "MyPicture" Then shp.Delete
Next shp
On Error GoTo 0

End Sub

Sub DelSheet()
Dim WsKont As Worksheet, shtInt As Integer, Str1 As String


ThisWorkbook.Worksheets(1).Visible = True
On Error Resume Next
ThisWorkbook.Worksheets(1).Activate
On Error GoTo 0

shtInt = 0
For Each WsKont In ThisWorkbook.Worksheets
    shtInt = shtInt + 1
    If WsKont.name = "Document Automation System" Or WsKont.name = "Definitions" Or WsKont.name = "Report 1 Workflow" Or WsKont.name = "Report 2 Workflow" Or WsKont.name = "Report 3 Workflow" _
        Or WsKont.name = "Statement Index" Or WsKont.name = "Temp Discrepancies" Or WsKont.name = "Discrepancies – Report 1" Or WsKont.name = "Discrepancies – Report 2" _
        Or WsKont.name = "Report 2 Numbers" Or WsKont.name = "Report 1 Numbers" Or WsKont.name = "Thermal Label" Or WsKont.name = "Processing Envelope" _
        Or WsKont.name = "Small Envelope" Or WsKont.name = "Large Envelope" Then
        'MsgBox shtInt & " : Ok"
    Else
        Str1 = Left(WsKont.name, 5)
        If StrComp(Str1, "Sheet", vbTextCompare) Or StrComp(Str1, "Sayfa", vbTextCompare) Then
            'MsgBox shtInt & " : Sheet Name: " & WsKont.name
            Worksheets(WsKont.name).Delete
        End If
    End If
Next WsKont

End Sub


'''''''''''''''''''''''''''''''''''''''''''''

Sub TimeSetting()
Dim WsSKP As Object

Set WsSKP = ThisWorkbook.Worksheets(2)
xDurum = WsSKP.Cells(6, 121).Value
xSaat = WsSKP.Cells(7, 121).Value
xDakika = WsSKP.Cells(8, 121).Value
xSaniye = WsSKP.Cells(9, 121).Value

If xDurum = "Enable" Then
    If xSaat = "" Then
        xSaat = "00"
    End If
    If xDakika = "" Then
        xDakika = "00"
    End If
    If xSaniye = "" Then
        xSaniye = "00"
    End If

    On Error GoTo PasifeGetir
    CloseTime = Now + TimeValue(xSaat & ":" & xDakika & ":" & xSaniye)
    'CloseTime = Now + TimeValue("00:30:00")
    On Error Resume Next
    Application.OnTime EarliestTime:=CloseTime, _
      Procedure:="OtomatikKapat", Schedule:=True
    
    On Error GoTo 0
    GoTo PasifAtla

PasifeGetir:
Call ModuleSystemSettings.TimeStop

PasifAtla:
End If

End Sub
Sub TimeStop()
    On Error Resume Next
    Application.OnTime EarliestTime:=CloseTime, _
      Procedure:="OtomatikKapat", Schedule:=False
 End Sub
''Sub SavedAndClose()
''    ThisWorkbook.Close Savechanges:=True
''End Sub

Sub OtomatikKapat()
Dim ReturnValue As Variant
Dim Wbk As Workbook

Application.ScreenUpdating = False
Application.DisplayAlerts = False

    ReturnValue = MsgBoxTimeout(0, "Enterprise Document Automation System will automatically save and close in 1 minute. To postpone automatic closure again for the waiting time you previously set, click No or Cancel. To close now, click Yes.", "Enterprise Document Automation System", vbQuestion + vbYesNoCancel + vbDefaultButton3, 0, 60000)
    Select Case ReturnValue
        Case vbYes
            'Açık userformlar varsa kapatılsın.
            Call ModuleInit.UserFormlariKapat
            For Each Wbk In Workbooks
              Wbk.Save
            Next Wbk
            Application.Quit
            GoTo Son
        Case vbNo
            'Debug.Print "You picked No."
            Call ModuleSystemSettings.TimeStop
            Call ModuleSystemSettings.TimeSetting
            GoTo Son
        Case vbCancel
            'Debug.Print "You picked Cancel."
            Call ModuleSystemSettings.TimeStop
            Call ModuleSystemSettings.TimeSetting
            GoTo Son
    End Select
 
'Açık userformlar varsa kapatılsın.
Call ModuleInit.UserFormlariKapat
For Each Wbk In Workbooks
  Wbk.Save
Next Wbk
Application.Quit
            
Son:

Application.ScreenUpdating = True
Application.DisplayAlerts = True

End Sub


Sub ScrollActive()

With ActiveWindow
    .DisplayHorizontalScrollBar = True
    .DisplayVerticalScrollBar = True
End With

End Sub


Sub NumLockAc()
Dim NumLockState As Boolean

Let NumLockState = CBool(GetKeyState(vbKeyNumlock) And 1)
'MsgBox NumLockState
If NumLockState = False Then
    Application.SendKeys "{NUMLOCK}{NUMLOCK}", True
End If

End Sub

Sub DropDownKapat()
Dim NumLockState As Boolean

Let NumLockState = CBool(GetKeyState(vbKeyNumlock) And 1)
'MsgBox NumLockState
If NumLockState = False Then
    Application.SendKeys "{NUMLOCK}", True
End If

'Açık dropdown kapat
Application.SendKeys "{F10}", True
'Tek sendkeys F10 komutu mousemove prosedürülerini kapattığı için, ikincisinde f10 görevi iptal ediliyor.
Application.SendKeys "{F10}", True


Let NumLockState = CBool(GetKeyState(vbKeyNumlock) And 1)
'MsgBox NumLockState
If NumLockState = False Then
    Application.SendKeys "{NUMLOCK}", True
End If

'ÖNEMLİ: userform numlock true/false değerini vbModeless iken alıyor; vbModal iken numlock boolean değeri alınamıyor.
'Bu yüzden userform vbmodeless modunda açılırken, userform vbModal olmadan önce NumLockAc prosedürü çağrılıp numlock'un boolean değeri elde edilmiş olunuyor.

'Numlock open olması için numlocku etkileyen kodlardan önce ve sonra yukarıdaki komutlar kullanılır.
'Numlock açıksa açık, kapalı ise kapalı kalıyor. Ancak userform aktifken numlock açılır veya kapatılır ise açılıp kapanma oluyor.
'Userformu sağ üst köşeden küçültüp numlock açılır. Tekrar aynı buton ile userform genişletilir ve sorun çözülmüş olur.

End Sub


