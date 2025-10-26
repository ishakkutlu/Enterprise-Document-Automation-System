VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clLabelClassCalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public WithEvents InputLabel1 As MSForms.Label
Attribute InputLabel1.VB_VarHelpID = -1

Private Sub InputLabel1_click()

With InputLabel1
   If .Tag < lStartPos Then
      If suppport_calendar_UI.lblBack.Enabled = True Then
         suppport_calendar_UI.lblBack_Click
      End If
      Exit Sub
   End If
   If .Tag > lDays + lStartPos - 1 Then
      suppport_calendar_UI.lblForward_Click
      Exit Sub
   End If
   If .BorderColor = vbBlue And lSelMonth1 = lSelMonth And lSelYear1 = lSelYear Then Exit Sub
   
   .BorderColor = vbBlue
   .BorderStyle = fmBorderStyleNone

   If Len(sActiveDay) > 0 Then
      If sActiveDay <> InputLabel1.name Then
         With colLabels.Item(sActiveDay)
            .BorderColor = &H8000000E
            .BorderStyle = fmBorderStyleNone
         End With
      End If
   End If
   sActiveDay = InputLabel1.name
   lFirstDay = Val(InputLabel1.Caption)

    On Error GoTo Bilgi
    suppport_calendar_UI.FillFirstDay
    GoTo Son
Bilgi:
    suppport_calendar_UI.CloseUserForm
Son:
   
   
End With

End Sub
