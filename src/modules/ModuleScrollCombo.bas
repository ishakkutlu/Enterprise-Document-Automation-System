Attribute VB_Name = "ModuleScrollCombo"
Option Explicit

#If Win64 Then
    Declare PtrSafe Function WindowFromPoint Lib "user32" (ByVal Point As LongPtr) As LongPtr
#Else
    Declare PtrSafe Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As LongPtr
#End If
Declare PtrSafe Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As LongPtr, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Declare PtrSafe Function GetParent Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
Declare PtrSafe Function GetActiveWindow Lib "user32" () As LongPtr
Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
Declare PtrSafe Function GetCursorPos Lib "user32" (ByRef lpPoint As POINTAPI) As LongPtr
Declare PtrSafe Function SetFocus Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
Declare PtrSafe Function IsWindow Lib "user32" (ByVal hwnd As LongPtr) As Long
Declare PtrSafe Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As LongPtr, ByVal hmod As LongPtr, ByVal dwThreadId As Long) As LongPtr
Declare PtrSafe Function CallNextHookEx Lib "user32" (ByVal hHook As LongPtr, ByVal nCode As Long, ByVal wParam As LongPtr, lParam As Any) As LongPtr
Declare PtrSafe Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As LongPtr) As LongPtr
Declare PtrSafe Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As Long
Declare PtrSafe Function GetClientRect Lib "user32" (ByVal hwnd As LongPtr, lpRect As RECT) As Long
Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
    
Type POINTAPI
    x As Long
    y As Long
End Type

Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Type MSLLHOOKSTRUCT
    pt As POINTAPI
    mouseData As Long
    flags As Long
    time As Long
    dwExtraInfo As Long
End Type

Const WH_MOUSE_LL = 14
Const WM_MOUSEWHEEL = &H20A
Const HC_ACTION = 0
Const WM_LBUTTONDOWN = &H201
Const WM_LBUTTONUP = &H202
Const MK_LBUTTON = &H1
Const SM_CXVSCROLL = 2

Dim hwnd As LongPtr, lMouseHook As LongPtr


Sub SetComboBoxHook(ByVal Control As Object)
    Dim tPT As POINTAPI
    Dim sBuffer As String
    Dim lRet As Long
    
    If lMouseHook = 0 Then
        GetCursorPos tPT
        #If VBA7 And Win64 Then
            Dim lPt As LongPtr
            CopyMemory lPt, tPT, LenB(tPT)
            hwnd = WindowFromPoint(lPt)
        #Else
            hwnd = WindowFromPoint(tPT.x, tPT.y)
        #End If
        sBuffer = Space(256)
        lRet = GetClassName(GetParent(hwnd), sBuffer, 256)
        If InStr(Left(sBuffer, lRet), "MdcPopup") Then
            SetFocus hwnd
            #If VBA7 Then
                lMouseHook = SetWindowsHookEx(WH_MOUSE_LL, AddressOf MouseProc, Application.HinstancePtr, 0)
            #Else
                lMouseHook = SetWindowsHookEx(WH_MOUSE_LL, AddressOf MouseProc, Application.hInstance, 0)
            #End If
        End If
    End If
End Sub

Sub RemoveComboBoxHook()
    UnhookWindowsHookEx lMouseHook: lMouseHook = 0
End Sub


#If VBA7 Then
    Function MouseProc(ByVal nCode As Long, ByVal wParam As LongPtr, lParam As MSLLHOOKSTRUCT) As LongPtr
#Else
    Function MouseProc(ByVal nCode As Long, ByVal wParam As Long, lParam As MSLLHOOKSTRUCT) As Long
#End If

    Dim sBuffer As String
    Dim lRet As Long
    Dim tRect As RECT
        
    sBuffer = Space(256)
    lRet = GetClassName(GetActiveWindow, sBuffer, 256)
    If Left(sBuffer, lRet) = "wndclass_desked_gsk" Then Call RemoveComboBoxHook
    If IsWindow(hwnd) = 0 Then Call RemoveComboBoxHook
    
    If (nCode = HC_ACTION) Then
        If wParam = WM_MOUSEWHEEL Then
            #If VBA7 And Win64 Then
                Dim lPt As LongPtr
                Dim Low As Integer, High As Integer
                Dim lParm As LongPtr
                CopyMemory lPt, lParam.pt, LenB(lPt)
                If WindowFromPoint(lPt) = hwnd Then
            #Else
                Dim Low As Integer, High As Integer
                Dim lParm As Long
                If WindowFromPoint(lParam.pt.x, lParam.pt.y) = hwnd Then
            #End If
                    GetClientRect hwnd, tRect
                    If lParam.mouseData > 0 Then
                        Low = tRect.Right - (GetSystemMetrics(SM_CXVSCROLL) / 2)
                        High = tRect.Top + ((GetSystemMetrics(SM_CXVSCROLL) / 2) + 1)
                        lParm = MakeDWord(Low, High)
                    Else
                        Low = tRect.Right - (GetSystemMetrics(SM_CXVSCROLL) / 2)
                        High = tRect.Bottom - ((GetSystemMetrics(SM_CXVSCROLL) / 2) + 1)
                        lParm = MakeDWord(Low, High)
                    End If
                    PostMessage hwnd, WM_LBUTTONDOWN, MK_LBUTTON, lParm
                    PostMessage hwnd, WM_LBUTTONUP, MK_LBUTTON, lParm
            End If
        End If
    End If
    
    MouseProc = CallNextHookEx(lMouseHook, nCode, wParam, ByVal lParam)
End Function

Private Function MakeDWord(ByVal LoWord As Integer, ByVal HiWord As Integer) As Long
    MakeDWord = (HiWord * &H10000) Or (LoWord And &HFFFF&)
End Function



