Attribute VB_Name = "modMsgBox"
Option Explicit

Private Declare Function MessageBoxEx Lib "user32" Alias "MessageBoxExA" (ByVal hWnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal uType As Long, ByVal wLanguageId As Long) As Long
Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function GetCurrentThreadId Lib "KERNEL32" () As Long
Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long

Private Const GWL_HINSTANCE = (-6)
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOZORDER = &H4
Private Const SWP_NOACTIVATE = &H10
Private Const HCBT_ACTIVATE = 5
Private Const WH_CBT = 5

Private Type RECT
    left As Long
    top As Long
    Right As Long
    Bottom As Long
End Type
    
Private hHook As Long
Private parenthWnd As Long


Public Function MsgBox(ByVal Prompt As String, Optional ByVal Buttons As VbMsgBoxStyle = vbOKOnly, Optional ByVal Title As String = "", Optional ByVal hWnd As Long = 0) As VbMsgBoxResult
 Dim hInst As Long
 Dim Thread As Long
    
    '/* Set up the hook
    parenthWnd = hWnd
    hInst = GetWindowLong(hWnd, GWL_HINSTANCE)
    Thread = GetCurrentThreadId()

    If hWnd > 0 Then
        hHook = SetWindowsHookEx(WH_CBT, AddressOf WinProcCenterForm, hInst, Thread)
    Else
        hHook = SetWindowsHookEx(WH_CBT, AddressOf WinProcCenterScreen, hInst, Thread)
    End If
    MsgBox = MessageBoxEx(hWnd, Prompt, Title, Buttons, 0)

End Function


Private Sub WinProcCenterScreen(ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
  Dim rForm As RECT, rMsg As RECT
  Dim x As Long, y As Long

    If lMsg = HCBT_ACTIVATE Then
        
        '/* Get the sizes and coordinates of the screen and the message box
        GetWindowRect wParam, rMsg
        '/* Center MsgBox on screen
        x = Screen.Width / Screen.TwipsPerPixelX / 2 - (rMsg.Right - rMsg.left) / 2
        y = Screen.Height / Screen.TwipsPerPixelY / 2 - (rMsg.Bottom - rMsg.top) / 2
        
        SetWindowPos wParam, -1, x, y, 0, 0, SWP_NOSIZE Or SWP_NOACTIVATE
        '/* Release the hook
        UnhookWindowsHookEx hHook
    End If

End Sub


Private Sub WinProcCenterForm(ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
  Dim rForm As RECT, rMsg As RECT
  Dim x As Long, y As Long
    
    '/* On Form Active, show the MsgBox centered over Form
    If lMsg = HCBT_ACTIVATE Then
        
        '/* Get the sizes and coordinates of the form and the message box
        GetWindowRect parenthWnd, rForm
        GetWindowRect wParam, rMsg
        x = (rForm.left + (rForm.Right - rForm.left) / 2) - ((rMsg.Right - rMsg.left) / 2)
        y = (rForm.top + (rForm.Bottom - rForm.top) / 2) - ((rMsg.Bottom - rMsg.top) / 2)
        
        '/* Center MsgBox over Form
        SetWindowPos wParam, 0, x, y, 0, 0, SWP_NOSIZE Or SWP_NOZORDER Or SWP_NOACTIVATE
        
        '/* Release the hook
        UnhookWindowsHookEx hHook
    End If

End Sub

