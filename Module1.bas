Attribute VB_Name = "Module1"
' Bas module for implementing system - wide shell hook.
' Using undocumented Shell32 function RegisterShellHook.
' Thanks to James Holderness for his help on using this function.
' You can find many othar undoc shell32 functions at
' http://www.geocities.com/SiliconValley/4942/contents.html

Public Declare Function RegisterWindowMessage Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String) As Long
Public Declare Function RegisterShellHook Lib "Shell32" Alias "#181" (ByVal hwnd As Long, ByVal nAction As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetKeyboardLayout Lib "user32" (ByVal dwLayout As Long) As Long
Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long

Public Const GWL_WNDPROC = (-4)

Public Const RSH_DEREGISTER = 0
Public Const RSH_REGISTER = 1
Public Const RSH_REGISTER_PROGMAN = 2
Public Const RSH_REGISTER_TASKMAN = 3

Const HSHELL_ACTIVATESHELLWINDOW = 3
Const HSHELL_WINDOWCREATED = 1
Const HSHELL_WINDOWDESTROYED = 2
Const HSHELL_WINDOWACTIVATED = 4
Const HSHELL_GETMINRECT = 5
Const HSHELL_REDRAW = 6
Const HSHELL_TASKMAN = 7
Const HSHELL_LANGUAGE = 8
Const HSHELL_ACCESSIBILITYSTATE = 11
Const LOCALE_SENGLANGUAGE As Long = &H1001
Public OldProc As Long, uRegMsg As Long

Public Function WndProc(ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  If wMsg = uRegMsg Then
     Dim sText As String
     Select Case wParam
            Case HSHELL_WINDOWCREATED
                 sText = "Window created. Caption = " & GetWndText(lParam) & " Handle = " & lParam
            Case HSHELL_WINDOWDESTROYED
                 sText = "Window destroyed. Caption = " & GetWndText(lParam) & " Handle = " & lParam
            Case HSHELL_WINDOWACTIVATED
                 sText = "Window activated. Caption = " & GetWndText(lParam) & " Handle = " & lParam
            Case HSHELL_LANGUAGE
                 Dim LocId As Long
                 LocId = LoWord(GetKeyboardLayout(0&))
                 sText = "Language changed to " & GetLanguageInfo(LocId, LOCALE_SENGLANGUAGE)
            Case HSHELL_GETMINRECT
                 sText = "Get Window RECT"
            Case HSHELL_REDRAW
                 sText = "Title in taskbar has been redrawn. Caption = " & GetWndText(lParam) & " Handle = " & lParam
            Case HSHELL_TASKMAN
                 sText = "Task Manager activated"
            Case HSHELL_ACTIVATESHELLWINDOW
                 sText = "Shell window activated"
     End Select
     Form1.Text1 = Form1.Text1 & sText & vbCrLf
  Else
     WndProc = CallWindowProc(OldProc, hwnd, wMsg, wParam, lParam)
  End If
End Function

Private Function GetWndText(hwnd As Long) As String
  Dim k As Long, sName As String
  sName = Space$(128)
  k = GetWindowText(hwnd, sName, 128)
  If k > 0 Then sName = Left$(sName, k) Else sName = "No caption"
  GetWndText = sName
End Function

Private Function LoWord(DWORD As Long) As Integer
   If DWORD And &H8000& Then
      LoWord = &H8000 Or (DWORD And &H7FFF&)
   Else
      LoWord = DWORD And &HFFFF&
   End If
End Function

Private Function GetLanguageInfo(ByVal dwLocaleID As Long, ByVal dwLCType As Long) As String
   Dim sReturn As String, nRet As Long
   sReturn = String$(128, 0)
   nRet = GetLocaleInfo(dwLocaleID, dwLCType, sReturn, Len(sReturn))
   If nRet > 0 Then GetLanguageInfo = Left$(sReturn, nRet - 1)
End Function

