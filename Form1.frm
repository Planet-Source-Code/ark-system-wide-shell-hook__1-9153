VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   3525
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5880
   LinkTopic       =   "Form1"
   ScaleHeight     =   3525
   ScaleWidth      =   5880
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   2895
      Left            =   360
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "Form1.frx":0000
      Top             =   240
      Width           =   5415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
  Text1 = ""
  Caption = "Shell log"
  uRegMsg = RegisterWindowMessage(ByVal "SHELLHOOK")
  Call RegisterShellHook(hwnd, RSH_REGISTER) ' Or RSH_REGISTER_TASKMAN Or RSH_REGISTER_PROGMAN)
  OldProc = GetWindowLong(hwnd, GWL_WNDPROC)
  SetWindowLong hwnd, GWL_WNDPROC, AddressOf WndProc
End Sub

Private Sub Form_Resize()
  Text1.Move 0, 0, ScaleWidth, ScaleHeight
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call RegisterShellHook(hwnd, RSH_DEREGISTER)
    SetWindowLong hwnd, GWL_WNDPROC, OldProc
End Sub

