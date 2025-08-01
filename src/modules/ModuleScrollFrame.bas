Attribute VB_Name = "ModuleScrollFrame"
Option Explicit
Public ScrollTakip As Long
Public ScrollTakip1 As Long
Public ScrollTakip2 As Long
Public ScrollTakip3 As Long

Private Declare PtrSafe Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As LongPtr, ByVal hmod As LongPtr, ByVal dwThreadId As Long) As LongPtr
Private Declare PtrSafe Function CallNextHookEx Lib "user32" (ByVal hHook As LongPtr, ByVal nCode As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
Private Declare PtrSafe Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As LongPtr) As Long
Private Declare PtrSafe Function GetActiveWindow Lib "user32" () As LongPtr
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)

Private Type MSLLHOOKSTRUCT
    pt As POINTAPI
    mouseData As Long
    flags As Long
    time As Long
    dwExtraInfo As LongPtr
End Type

Private Const WH_MOUSE_LL As Long = 14
Private Const WM_MOUSEWHEEL As Long = &H20A
Private Const SCROLL_STEP As Long = 20

Private hHook As LongPtr
Private FormHwnd As LongPtr
Private TargetScrollFrame As Object

Public Sub SetScrollHook(ScrollableFrame As Object, ScrollTrack As Long, Optional ScrollMargin As Long = 0)
    If Not TypeName(ScrollableFrame) = "Frame" Then Exit Sub
    Set TargetScrollFrame = ScrollableFrame
    FormHwnd = GetActiveWindow

    With ScrollableFrame
        .ScrollBars = fmScrollBarsVertical
        .KeepScrollBarsVisible = fmScrollBarsVertical
        .PictureAlignment = fmPictureAlignmentTopLeft
        .ScrollWidth = .InsideWidth * 3
        .ScrollHeight = ScrollTrack + ScrollMargin
    End With

    If hHook = 0 Then
        hHook = SetWindowsHookEx(WH_MOUSE_LL, AddressOf MouseProc, 0, 0)
    End If
End Sub

Public Sub Move_SetScrollHook(ScrollableFrame As Object, ThresholdVal As Long, ScrollTrack As Long)
    If Not TypeName(ScrollableFrame) = "Frame" Then Exit Sub
    Set TargetScrollFrame = ScrollableFrame
    FormHwnd = GetActiveWindow
    If ScrollTrack > ThresholdVal Then
        If hHook = 0 Then
            hHook = SetWindowsHookEx(WH_MOUSE_LL, AddressOf MouseProc, 0, 0)
        End If
    End If
End Sub

Public Sub RemoveScrollHook()
    If hHook <> 0 Then
        UnhookWindowsHookEx hHook
        hHook = 0
    End If
End Sub

Private Function MouseProc(ByVal nCode As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
Dim delta As Long
Dim mouseStruct As MSLLHOOKSTRUCT

    On Error GoTo errH

    If TargetScrollFrame Is Nothing Then
        MouseProc = CallNextHookEx(hHook, nCode, wParam, lParam)
        Exit Function
    End If
    If nCode = 0 And wParam = WM_MOUSEWHEEL Then
        'CopyMemory mouseStruct, ByVal lParam, LenB(mouseStruct)
        'CopyMemory ByVal VarPtr(mouseStruct), ByVal lParam, LenB(mouseStruct)
        CopyMemory ByVal VarPtr(mouseStruct), ByVal lParam, Len(mouseStruct)
        delta = (mouseStruct.mouseData And &HFFFF0000) \ &H10000
        With TargetScrollFrame
            If delta > 0 Then
                .ScrollTop = Application.Max(0, .ScrollTop - SCROLL_STEP)
            Else
                .ScrollTop = Application.Min(.ScrollHeight - .InsideHeight, .ScrollTop + SCROLL_STEP)
            End If
        End With
    End If

    MouseProc = CallNextHookEx(hHook, nCode, wParam, lParam)
    Exit Function

errH:
    RemoveScrollHook
End Function






