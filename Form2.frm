VERSION 5.00
Begin VB.Form Cover_R 
   BackColor       =   &H00000000&
   BorderStyle     =   1  '單線固定
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   1905
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   1905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '系統預設值
End
Attribute VB_Name = "Cover_R"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const SWP_NOMOVE = &H2
Const SWP_NOSIZE = &H1
Const FLAG = SWP_NOMOVE Or SWP_NOSIZE
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const HWND_TOP = 0
Const HWND_BOTTOM = 1
Private Declare Function SetWindowPos Lib "user32" _
  (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, _
   ByVal y As Long, ByVal cx As Long, ByVal cy As Long, _
   ByVal wFlags As Long) As Long
Private Sub Form_Load()
Call SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAG)
Cover_R.Left = Screen.Width - Cover_R.Width
Cover_R.Top = Screen.Height - Cover_L.Height - 1000
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
ct_form.showright.Enabled = True
End Sub
