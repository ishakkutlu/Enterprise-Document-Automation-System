VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} suppport_calendar_UI 
   Caption         =   "Enterprise Document Automation System"
   ClientHeight    =   4860
   ClientLeft      =   30
   ClientTop       =   375
   ClientWidth     =   4755
   OleObjectBlob   =   "suppport_calendar_UI.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "suppport_calendar_UI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bCmbSel As Boolean
Dim lFirstDayInMonth As Long
Dim lDayPos As Long
Dim lMonthPos As Long
'Dim sMonth As String
Dim sDateFormat As String
Dim datFirstDay As Date
Dim datLastDay As Date

Private Sub Kapat_Click()
    CloseUserForm
End Sub

Private Sub Kapat_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
FrameKapat.BackColor = RGB(180, 180, 180)
End Sub

Private Sub Kapat_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
FrameKapat.BackColor = RGB(200, 200, 200)
End Sub

Private Sub UserForm_Initialize()
Dim ctl As MSForms.Control
Dim lCount As Long
Dim InputLblEvt As clLabelClassCalendar
Dim ClrLab As Control

On Error GoTo ErrorHandle

For Each ClrLab In suppport_calendar_UI.Controls
    If TypeName(ClrLab) = "Label" Then
        ClrLab.BackColor = RGB(180, 180, 180) 'RGB(254, 254, 254)
        ClrLab.ForeColor = RGB(30, 30, 30)
    End If

    'YENİ
    If TypeName(ClrLab) = "Frame" Then
        ClrLab.BackColor = RGB(254, 254, 254)
        ClrLab.ForeColor = RGB(30, 30, 30)
        ClrLab.BorderColor = RGB(180, 180, 180)
    End If
Next ClrLab
FrameKapat.BackColor = RGB(254, 254, 254)


suppport_calendar_UI.BackColor = RGB(230, 230, 230) 'YENİ'RGB(225, 235, 245)
cmbMonth.BackColor = RGB(255, 255, 255)
cmbYear.BackColor = RGB(255, 255, 255)

Set colLabelEvent = New Collection
Set colLabels = New Collection

For Each ctl In Frame1.Controls
   If TypeOf ctl Is MSForms.Label Then
      Set InputLblEvt = New clLabelClassCalendar
      Set InputLblEvt.InputLabel1 = ctl
      colLabelEvent.Add InputLblEvt
      colLabels.Add ctl, ctl.name
   End If
Next

Set InputLblEvt = Nothing

Call Init_EnglishMonths ' English
For lCount = 1 To 12
   cmbMonth.AddItem EnglishMonths(lCount)
Next

'Turkish
'For lCount = 1 To 12
'   With cmbMonth
'      .AddItem MonthName(lCount)
'   End With
'Next

For lCount = 1900 To Year(Now) + 100
   With cmbYear
      .AddItem lCount
   End With
Next

' Turkish
'lblDay1.Caption = StrConv(Left(WeekdayName(1, , vbUseSystemDayOfWeek), 3), 1)
'lblDay2.Caption = StrConv(Left(WeekdayName(2, , vbUseSystemDayOfWeek), 3), 1)
'lblDay3.Caption = StrConv(Left(WeekdayName(3, , vbUseSystemDayOfWeek), 3), 1)
'lblDay4.Caption = StrConv(Left(WeekdayName(4, , vbUseSystemDayOfWeek), 3), 1)
'lblDay5.Caption = StrConv(Left(WeekdayName(5, , vbUseSystemDayOfWeek), 3), 1)
'lblDay6.Caption = StrConv(Left(WeekdayName(6, , vbUseSystemDayOfWeek), 3), 1)
'lblDay7.Caption = StrConv(Left(WeekdayName(7, , vbUseSystemDayOfWeek), 3), 1)

lblDay1.Caption = "Mon"
lblDay2.Caption = "Tue"
lblDay3.Caption = "Wed"
lblDay4.Caption = "Thu"
lblDay5.Caption = "Fri"
lblDay6.Caption = "Sat"
lblDay7.Caption = "Sun"

With colLabels
   For lCount = 1 To .Count
      .Item(lCount).Tag = lCount
   Next
End With

LabelCaptions Month(Now), Year(Now)

lDayPos = Day("01-02-03")
lMonthPos = Month("01-02-03")

cmbMonthFrame.BackColor = RGB(254, 254, 254)
cmbYearFrame.BackColor = RGB(254, 254, 254)
lblBackFrame.BackColor = RGB(254, 254, 254)
lblForwardFrame.BackColor = RGB(254, 254, 254)

Call SignToday

Exit Sub
ErrorHandle:
MsgBox Err.Description

End Sub
Sub LabelCaptions(lMonth As Long, lYear As Long)
Dim lCount As Long
Dim lNumber As Long
Dim lMonthPrev As Long
Dim lDaysPrev As Long
Dim lYearPrev As Long


sMonth = EnglishMonths(lMonth) ' English
' sMonth = MonthName(lMonth) ' Turkish
lSelMonth = lMonth
lSelYear = lYear

If bSecondDate = False Then
    lSelMonth1 = lSelMonth
    lSelYear1 = lSelYear
End If

Select Case lMonth
   Case 2 To 11
      lMonthPrev = lMonth - 1
      lYearPrev = lYear
   Case 1
      lMonthPrev = 12
      lYearPrev = lYear - 1
   Case 12
      lMonthPrev = 11
      lYearPrev = lYear
End Select
   
lDays = DaysInMonth(lMonth, lYear)
lDaysPrev = DaysInMonth(lMonthPrev, lYearPrev)

If lSelYear >= 1900 And lSelMonth > 1 Then
   lblBack.Enabled = True
ElseIf lSelYear = 1900 And lSelMonth = 1 Then
   lblBack.Enabled = False
End If

If bCmbSel = False Then
   cmbMonth.Text = sMonth
   cmbYear.Text = lYear
End If

lFirstDayInMonth = DateSerial(lSelYear, lSelMonth, 1)
lFirstDayInMonth = Weekday(lFirstDayInMonth, vbUseSystemDayOfWeek)

If lFirstDayInMonth = 1 Then
   lStartPos = 8
Else
   lStartPos = lFirstDayInMonth
End If

lNumber = lDaysPrev + 1
For lCount = lStartPos - 1 To 1 Step -1
   lNumber = lNumber - 1
   With colLabels.Item(lCount)
      .Caption = lNumber
      .ForeColor = &HE0E0E0
   End With
Next

lNumber = 0
For lCount = lStartPos To lDays + lStartPos - 1
   lNumber = lNumber + 1
   With colLabels.Item(lCount)
      .Caption = lNumber
      .ForeColor = RGB(30, 30, 30)
   End With
Next

lNumber = 0
For lCount = lDays + lStartPos To 42
   lNumber = lNumber + 1
   With colLabels.Item(lCount)
      .Caption = lNumber
      .ForeColor = &HE0E0E0
   End With
Next

End Sub
Function DaysInMonth(lMonth As Long, lYear As Long) As Long

Select Case lMonth
   Case 1, 3, 5, 7, 8, 10, 12
      DaysInMonth = 31
   Case 2
      If IsDate("29/2/" & lYear) = False Then
         DaysInMonth = 28
      Else
         DaysInMonth = 29
      End If
   Case Else
      DaysInMonth = 30
End Select

End Function

Private Sub lblBack_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
   'lblBack.SpecialEffect = fmSpecialEffectSunken
   lblBackFrame.BackColor = RGB(180, 180, 180)
End Sub
Private Sub lblBack_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
   'lblBack.SpecialEffect = fmSpecialEffectRaised
   lblBackFrame.BackColor = RGB(200, 200, 200)
End Sub
Private Sub lblForward_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
   'lblForward.SpecialEffect = fmSpecialEffectSunken
   lblForwardFrame.BackColor = RGB(180, 180, 180)
End Sub
Private Sub lblForward_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
   'lblForward.SpecialEffect = fmSpecialEffectRaised
   lblForwardFrame.BackColor = RGB(200, 200, 200)
End Sub


Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    lblBackFrame.BackColor = RGB(254, 254, 254)
    cmbMonthFrame.BackColor = RGB(254, 254, 254)
    cmbYearFrame.BackColor = RGB(254, 254, 254)
    lblForwardFrame.BackColor = RGB(254, 254, 254)
    FrameKapat.BackColor = RGB(254, 254, 254)
End Sub

Private Sub Frame2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    lblBackFrame.BackColor = RGB(254, 254, 254)
    cmbMonthFrame.BackColor = RGB(254, 254, 254)
    cmbYearFrame.BackColor = RGB(254, 254, 254)
    lblForwardFrame.BackColor = RGB(254, 254, 254)
End Sub
Private Sub Frame3_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    lblBackFrame.BackColor = RGB(254, 254, 254)
    cmbMonthFrame.BackColor = RGB(254, 254, 254)
    cmbYearFrame.BackColor = RGB(254, 254, 254)
    lblForwardFrame.BackColor = RGB(254, 254, 254)
End Sub
Private Sub Frame4_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    lblBackFrame.BackColor = RGB(254, 254, 254)
    cmbMonthFrame.BackColor = RGB(254, 254, 254)
    cmbYearFrame.BackColor = RGB(254, 254, 254)
    lblForwardFrame.BackColor = RGB(254, 254, 254)
    FrameKapat.BackColor = RGB(254, 254, 254)
End Sub


Private Sub Frame1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
Dim MonthNm As Integer, YearNm As Integer
Dim ClrLab As Control

YearNm = suppport_calendar_UI.cmbYear.Value
' MonthNm = Month(DateValue("01 " & suppport_calendar_UI.cmbMonth.Value & " 2015")) ' Turkish

'English Months
Call Init_EnglishMonths
Dim i As Long
For i = 1 To 12
    If suppport_calendar_UI.cmbMonth.Value = EnglishMonths(i) Then
        MonthNm = i
        Exit For
    End If
Next i

For Each ClrLab In suppport_calendar_UI.Frame1.Controls
    If ClrLab.BackColor <> RGB(180, 180, 180) Then
        ClrLab.BackColor = RGB(180, 180, 180)
    End If
    If ClrLab < 10 Then
     If Format(Date, "yyyy") = YearNm And Format(Date, "mm") = MonthNm And ClrLab = Format(Date, "d") And Not ClrLab.ForeColor = &HE0E0E0 Then
        If ClrLab.BackColor = RGB(249, 194, 19) Then
            '
        Else
            ClrLab.BackColor = RGB(249, 194, 19)
        End If
     End If
    Else
     If Format(Date, "yyyy") = YearNm And Format(Date, "mm") = MonthNm And ClrLab = Format(Date, "dd") And Not ClrLab.ForeColor = &HE0E0E0 Then
        If ClrLab.BackColor = RGB(249, 194, 19) Then
            '
        Else
            ClrLab.BackColor = RGB(249, 194, 19)
        End If
     End If
    End If
    
Next

lblBackFrame.BackColor = RGB(254, 254, 254)
cmbMonthFrame.BackColor = RGB(254, 254, 254)
cmbYearFrame.BackColor = RGB(254, 254, 254)
lblForwardFrame.BackColor = RGB(254, 254, 254)
FrameKapat.BackColor = RGB(254, 254, 254)

Call SignToday

End Sub

Private Sub Kapat_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    FrameKapat.BackColor = RGB(225, 235, 245)
End Sub
Private Sub FrameKapat_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    FrameKapat.BackColor = RGB(225, 235, 245)
End Sub

Private Sub lblBack_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    lblBackFrame.BackColor = RGB(225, 235, 245) 'RGB(200, 200, 200)
    cmbMonthFrame.BackColor = RGB(254, 254, 254)
End Sub
Private Sub lblBackFrame_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    lblBackFrame.BackColor = RGB(254, 254, 254)
    cmbMonthFrame.BackColor = RGB(254, 254, 254)
End Sub

Private Sub cmbMonth_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    cmbMonthFrame.BackColor = RGB(225, 235, 245) 'RGB(200, 200, 200)
    lblBackFrame.BackColor = RGB(254, 254, 254)
    cmbYearFrame.BackColor = RGB(254, 254, 254)
End Sub
Private Sub cmbMonthFrame_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    lblBackFrame.BackColor = RGB(254, 254, 254)
    cmbMonthFrame.BackColor = RGB(254, 254, 254)
    cmbYearFrame.BackColor = RGB(254, 254, 254)
End Sub

Private Sub cmbYear_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    cmbYearFrame.BackColor = RGB(225, 235, 245) 'RGB(200, 200, 200)
    cmbMonthFrame.BackColor = RGB(254, 254, 254)
    lblForwardFrame.BackColor = RGB(254, 254, 254)
    Call SetComboBoxHook(cmbYear) 'Open scrollable with mouse
End Sub

Private Sub cmbYearFrame_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    cmbMonthFrame.BackColor = RGB(254, 254, 254)
    cmbYearFrame.BackColor = RGB(254, 254, 254)
    lblForwardFrame.BackColor = RGB(254, 254, 254)
End Sub


Private Sub lblForward_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    lblForwardFrame.BackColor = RGB(225, 235, 245) 'RGB(200, 200, 200)
    cmbYearFrame.BackColor = RGB(254, 254, 254)
End Sub
Private Sub lblForwardFrame_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    cmbYearFrame.BackColor = RGB(254, 254, 254)
    lblForwardFrame.BackColor = RGB(254, 254, 254)
End Sub


Private Sub Label1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label1.BackColor = RGB(249, 194, 21)
End Sub
Private Sub Label2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label2.BackColor = RGB(249, 194, 21)
End Sub
Private Sub Label3_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label3.BackColor = RGB(249, 194, 21)
End Sub
Private Sub Label4_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label4.BackColor = RGB(249, 194, 21)
End Sub
Private Sub Label5_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label5.BackColor = RGB(249, 194, 21)
End Sub
Private Sub Label6_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label6.BackColor = RGB(249, 194, 21)
End Sub
Private Sub Label7_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label7.BackColor = RGB(249, 194, 21)
End Sub
Private Sub Label8_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label8.BackColor = RGB(249, 194, 21)
End Sub
Private Sub Label9_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label9.BackColor = RGB(249, 194, 21)
End Sub
Private Sub Label10_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label10.BackColor = RGB(249, 194, 21)
End Sub
Private Sub Label11_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label11.BackColor = RGB(249, 194, 21)
End Sub
Private Sub Label12_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label12.BackColor = RGB(249, 194, 21)
End Sub
Private Sub Label13_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label13.BackColor = RGB(249, 194, 21)
End Sub
Private Sub Label14_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label14.BackColor = RGB(249, 194, 21)
End Sub
Private Sub Label15_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label15.BackColor = RGB(249, 194, 21)
End Sub
Private Sub Label16_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label16.BackColor = RGB(249, 194, 21)
End Sub
Private Sub Label17_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label17.BackColor = RGB(249, 194, 21)
End Sub
Private Sub Label18_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label18.BackColor = RGB(249, 194, 21)
End Sub
Private Sub Label19_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label19.BackColor = RGB(249, 194, 21)
End Sub
Private Sub Label20_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label20.BackColor = RGB(249, 194, 21)
End Sub
Private Sub Label21_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label21.BackColor = RGB(249, 194, 21)
End Sub
Private Sub Label22_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label22.BackColor = RGB(249, 194, 21)
End Sub
Private Sub Label23_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label23.BackColor = RGB(249, 194, 21)
End Sub
Private Sub Label24_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label24.BackColor = RGB(249, 194, 21)
End Sub
Private Sub Label25_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label25.BackColor = RGB(249, 194, 21)
End Sub
Private Sub Label26_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label26.BackColor = RGB(249, 194, 21)
End Sub
Private Sub Label27_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label27.BackColor = RGB(249, 194, 21)
End Sub
Private Sub Label28_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label28.BackColor = RGB(249, 194, 21)
End Sub
Private Sub Label29_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label29.BackColor = RGB(249, 194, 21)
End Sub
Private Sub Label30_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label30.BackColor = RGB(249, 194, 21)
End Sub
Private Sub Label31_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label31.BackColor = RGB(249, 194, 21)
End Sub
Private Sub Label32_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label32.BackColor = RGB(249, 194, 21)
End Sub
Private Sub Label33_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label33.BackColor = RGB(249, 194, 21)
End Sub
Private Sub Label34_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label34.BackColor = RGB(249, 194, 21)
End Sub
Private Sub Label35_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label35.BackColor = RGB(249, 194, 21)
End Sub
Private Sub Label43_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label43.BackColor = RGB(249, 194, 21)
End Sub
Private Sub Label44_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label44.BackColor = RGB(249, 194, 21)
End Sub
Private Sub Label45_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label45.BackColor = RGB(249, 194, 21)
End Sub
Private Sub Label46_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label46.BackColor = RGB(249, 194, 21)
End Sub
Private Sub Label47_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label47.BackColor = RGB(249, 194, 21)
End Sub
Private Sub Label48_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label48.BackColor = RGB(249, 194, 21)
End Sub
Private Sub Label49_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    Label49.BackColor = RGB(249, 194, 21)
End Sub

Sub lblForward_Click()
Dim MonthNm As Integer, YearNm As Integer

YearNm = suppport_calendar_UI.cmbYear.Value
' MonthNm = Month(DateValue("01 " & suppport_calendar_UI.cmbMonth.Value & " 2015")) 'Turkish

'English Months
Call Init_EnglishMonths
Dim i As Long
For i = 1 To 12
    If suppport_calendar_UI.cmbMonth.Value = EnglishMonths(i) Then
        MonthNm = i
        Exit For
    End If
Next i

lSelMonth = MonthNm
lSelYear = YearNm

If lSelMonth < 12 Then
   lSelMonth = lSelMonth + 1
Else
   lSelMonth = 1
   lSelYear = lSelYear + 1
End If

If Len(sActiveDay) > 0 Then
   With colLabels.Item(sActiveDay)
      .BorderColor = &H8000000E
      .BorderStyle = fmBorderStyleNone
   End With
End If

Call cmbMonth_Change
Call cmbYear_Change
LabelCaptions lSelMonth, lSelYear

Call SignToday

End Sub
Sub lblBack_Click()

If lSelMonth > 1 Then
   lSelMonth = lSelMonth - 1
Else
   lSelMonth = 12
   lSelYear = lSelYear - 1
End If

If Len(sActiveDay) > 0 Then
   With colLabels.Item(sActiveDay)
      .BorderColor = &H8000000E
      .BorderStyle = fmBorderStyleNone
   End With
End If

Call cmbMonth_Change
Call cmbYear_Change
LabelCaptions lSelMonth, lSelYear

Call SignToday

End Sub
Private Sub cmbMonth_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
bCmbSel = True
End Sub
Private Sub cmbMonth_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
bCmbSel = True
End Sub
Private Sub cmbMonth_Change()
Dim lOldMonth As Long

If cmbMonth.matchFound = False Then
    MsgBox "An unrecognized month was entered by the system. Please select a valid month.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    
    'English Months
    Call Init_EnglishMonths
    Dim i As Long
    For i = 1 To 12
        If suppport_calendar_UI.cmbMonth.Value = EnglishMonths(i) Then
            lSelMonth = i
            Exit For
        End If
    Next i
    cmbMonth.Text = EnglishMonths(lSelMonth)
    
    ' cmbMonth.Text = MonthName(lSelMonth) ' Turkish
    GoTo Atla
End If

If bCmbSel Then
   If cmbMonth.matchFound = False Then Exit Sub
   lOldMonth = lSelMonth
   ' lSelMonth = Month(DateValue("01 " & cmbMonth.Text & " 2015")) 'Turkish

    'English Months
    Call Init_EnglishMonths
    For i = 1 To 12
        If suppport_calendar_UI.cmbMonth.Value = EnglishMonths(i) Then
            lSelMonth = i
            Exit For
        End If
    Next i

   If lSelMonth <> lOldMonth Then
      LabelCaptions lSelMonth, lSelYear
   End If
   bCmbSel = False
   If Len(sActiveDay) > 0 Then
      colLabels.Item(sActiveDay).SpecialEffect = fmSpecialEffectFlat
   End If
   Call SignToday
End If

Atla:

End Sub
Private Sub cmbMonth_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Me.cmbMonth.DropDown
End Sub

Private Sub cmbYear_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
bCmbSel = True
End Sub
Private Sub cmbYear_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
bCmbSel = True
End Sub

Private Sub cmbYear_Change()
Dim lOldYear As Long

If cmbYear.matchFound = False Then
    MsgBox "An unrecognized year was entered by the system. Please select a valid year.", vbOKOnly + vbExclamation, "Enterprise Document Automation System"
    cmbYear.Text = Year(Now)
    GoTo Atla
End If

If bCmbSel Then
   lOldYear = lSelYear
   If Val(cmbYear.Text) < 1900 Then
      cmbYear.Text = lSelYear
      bCmbSel = False
      Exit Sub
   End If
   lSelYear = Year("01 " & MonthName(lSelMonth) & " " & cmbYear.Text)
   If lSelYear <> lOldYear Then
      LabelCaptions lSelMonth, lSelYear
   End If
   bCmbSel = False
   If Len(sActiveDay) > 0 Then
      colLabels.Item(sActiveDay).SpecialEffect = fmSpecialEffectFlat
   End If
   Call SignToday
End If

Atla:

End Sub

Private Sub cmbYear_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Me.cmbYear.DropDown
End Sub

Sub FillFirstDay()

datFirstDay = ReturnDate(lFirstDay, lSelMonth, lSelYear)
sDateFormat = Format("dd/mm/yyyy")

CalTarih = Format(datFirstDay, sDateFormat)

Unload Me

End Sub

Sub CloseUserForm()

Unload Me

End Sub

Function ReturnDate(ByVal lDay As Long, ByVal lMonth As Long, ByVal lYear As Long) As Date

If lDayPos = 1 And lMonthPos = 2 Then
   ReturnDate = lDay & "/" & lMonth & "/" & lYear
   Exit Function
ElseIf lDayPos = 2 And lMonthPos = 1 Then
   ReturnDate = lMonth & "/" & lDay & "/" & lYear
   Exit Function
ElseIf lDayPos = 3 And lMonthPos = 2 Then
   ReturnDate = lYear & "/" & lMonth & "/" & lDay
   Exit Function
ElseIf lDayPos = 2 And lMonthPos = 3 Then
   ReturnDate = lYear & "/" & lDay & "/" & lMonth
   Exit Function
ElseIf lDayPos = 1 And lMonthPos = 3 Then
   ReturnDate = lDay & "/" & lYear & "/" & lMonth
   Exit Function
ElseIf lMonthPos = 1 And lDayPos = 3 Then
   ReturnDate = lMonth & "/" & lYear & "/" & lDay
End If

End Function




