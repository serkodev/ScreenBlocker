VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form ct_form 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  '單線固定
   Caption         =   "檔屏工具v1.5"
   ClientHeight    =   4695
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4095
   Icon            =   "ct_form.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   4095
   Begin VB.CommandButton Command10 
      Caption         =   "確定"
      Height          =   255
      Left            =   3360
      TabIndex        =   29
      Top             =   4320
      Width           =   615
   End
   Begin VB.CommandButton Command9 
      Caption         =   "確定"
      Height          =   255
      Left            =   120
      TabIndex        =   28
      Top             =   4320
      Width           =   615
   End
   Begin VB.CommandButton Command8 
      Caption         =   "確定"
      Height          =   255
      Left            =   1800
      TabIndex        =   27
      Top             =   1440
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   1800
      MaxLength       =   5
      TabIndex        =   26
      Top             =   1080
      Width           =   615
   End
   Begin VB.TextBox Text3 
      Height          =   270
      Left            =   3360
      MaxLength       =   5
      TabIndex        =   25
      Top             =   3960
      Width           =   615
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   120
      MaxLength       =   5
      TabIndex        =   24
      Top             =   3960
      Width           =   615
   End
   Begin VB.CommandButton Command7 
      Caption         =   "顯示右檔屏"
      Height          =   375
      Left            =   2880
      TabIndex        =   23
      Top             =   120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      Caption         =   "顯示左右檔屏"
      Height          =   375
      Left            =   1320
      TabIndex        =   22
      Top             =   120
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton Command5 
      Caption         =   "顯示左檔屏"
      Height          =   375
      Left            =   120
      TabIndex        =   21
      Top             =   120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "右寬等於左寬"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1080
      TabIndex        =   20
      Top             =   3120
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "左寬等於右寬"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1080
      TabIndex        =   19
      Top             =   2520
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "解除鎖定兩則檔屏"
      Height          =   375
      Left            =   1080
      TabIndex        =   18
      Top             =   1920
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "隱藏左右檔屏"
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton showright 
      Caption         =   "顯示右檔屏"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton showleft 
      Caption         =   "顯示左檔屏"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
   Begin MSComctlLib.Slider covervalue 
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   840
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   661
      _Version        =   393216
      MousePointer    =   9
      Min             =   2000
      Max             =   10000
      SelStart        =   3000
      TickStyle       =   3
      Value           =   3000
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   2295
      Left            =   480
      TabIndex        =   4
      Top             =   1680
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   4048
      _Version        =   393216
      MousePointer    =   7
      Orientation     =   1
      Min             =   3500
      Max             =   15000
      SelStart        =   3500
      TickStyle       =   3
      Value           =   3500
   End
   Begin MSComctlLib.Slider Slider2 
      Height          =   2295
      Left            =   3240
      TabIndex        =   5
      Top             =   1680
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   4048
      _Version        =   393216
      MousePointer    =   5
      Orientation     =   1
      Max             =   10000
      SelStart        =   3500
      TickStyle       =   3
      Value           =   3500
   End
   Begin VB.Label Label5 
      Appearance      =   0  '平面
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  '單線固定
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   7
      Left            =   2520
      TabIndex        =   17
      Top             =   3960
      Width           =   255
   End
   Begin VB.Label Label5 
      Appearance      =   0  '平面
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  '單線固定
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   6
      Left            =   2520
      TabIndex        =   16
      Top             =   4320
      Width           =   255
   End
   Begin VB.Label Label5 
      Appearance      =   0  '平面
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  '單線固定
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   5
      Left            =   2115
      TabIndex        =   15
      Top             =   4320
      Width           =   255
   End
   Begin VB.Label Label5 
      Appearance      =   0  '平面
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  '單線固定
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   1725
      TabIndex        =   14
      Top             =   4320
      Width           =   255
   End
   Begin VB.Label Label5 
      Appearance      =   0  '平面
      BackColor       =   &H000000FF&
      BorderStyle     =   1  '單線固定
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   1320
      TabIndex        =   13
      Top             =   4320
      Width           =   255
   End
   Begin VB.Label Label5 
      Appearance      =   0  '平面
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  '單線固定
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   2115
      TabIndex        =   12
      Top             =   3960
      Width           =   255
   End
   Begin VB.Label Label5 
      Appearance      =   0  '平面
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  '單線固定
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   1725
      TabIndex        =   11
      Top             =   3960
      Width           =   255
   End
   Begin VB.Label Label5 
      Appearance      =   0  '平面
      BackColor       =   &H00000000&
      BorderStyle     =   1  '單線固定
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   1320
      TabIndex        =   10
      Top             =   3960
      Width           =   255
   End
   Begin VB.Label Label4 
      Alignment       =   2  '置中對齊
      Caption         =   "檔屏垂直"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3000
      TabIndex        =   9
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label Label3 
      Alignment       =   2  '置中對齊
      Caption         =   "檔屏高度"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   240
      TabIndex        =   8
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   2  '置中對齊
      Caption         =   "檔屏顏色"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1320
      TabIndex        =   7
      Top             =   3720
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  '置中對齊
      Caption         =   "檔屏寬度"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   120
      TabIndex        =   6
      Top             =   600
      Width           =   3855
   End
End
Attribute VB_Name = "ct_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Command1.Caption = "隱藏左右檔屏" Then
Unload Cover_L
Unload Cover_R
Command1.Caption = "顯示左右檔屏"
Else
Cover_L.Width = covervalue.Value
Cover_R.Width = covervalue.Value
Cover_R.Left = Screen.Width - Cover_R.Width

Cover_L.Height = Slider1.Value
Cover_R.Height = Slider1.Value

Cover_L.Top = Slider2.Value
Cover_R.Top = Slider2.Value

Cover_L.Show
Cover_R.Show
Command1.Caption = "隱藏左右檔屏"
End If
End Sub

Private Sub Command10_Click()
Slider2.Value = Text3
Cover_L.Top = Slider2.Value
Cover_R.Top = Slider2.Value
End Sub

Private Sub Command2_Click()
If Command2.Caption = "解除鎖定兩則檔屏" Then
Command2.Caption = "鎖定兩則檔屏"
Free_L.Top = Cover_L.Top
Free_L.Left = Cover_L.Left
Free_L.Height = Cover_L.Height
Free_L.Width = Cover_L.Width
Free_R.Top = Cover_R.Top
Free_R.Left = Cover_R.Left
Free_R.Height = Cover_R.Height
Free_R.Width = Cover_R.Width
Unload Cover_L
Unload Cover_R
Command1.Caption = "顯示左右檔屏"

Command5.Visible = True
Command6.Visible = True
Command7.Visible = True
Command1.Visible = False
showleft.Visible = False
showright.Visible = False

covervalue.Enabled = False
Slider1.Enabled = False
Slider2.Enabled = False
Free_L.Show
Free_R.Show
Command3.Enabled = True
Command4.Enabled = True
Else
Command2.Caption = "解除鎖定兩則檔屏"
Unload Free_L
Unload Free_R

Command1.Visible = True
showleft.Visible = True
showright.Visible = True
Command5.Visible = False
Command6.Visible = False
Command7.Visible = False

covervalue.Enabled = True
Slider1.Enabled = True
Slider2.Enabled = True
Command3.Enabled = False
Command4.Enabled = False
Command1_Click
End If
End Sub

Private Sub Command3_Click()
Free_L.Width = Free_R.Width
End Sub

Private Sub Command4_Click()
Free_R.Width = Free_L.Width
End Sub

Private Sub Command5_Click()
Free_L.Show
End Sub

Private Sub Command6_Click()
Free_L.Show
Free_R.Show
End Sub

Private Sub Command7_Click()
Free_R.Show
End Sub

Private Sub Command8_Click()
covervalue.Value = Text1
Cover_L.Width = covervalue.Value
Cover_R.Width = covervalue.Value
Cover_R.Left = Screen.Width - Cover_R.Width
End Sub

Private Sub Command9_Click()
Slider1.Value = Text2
Cover_L.Height = Slider1.Value
Cover_R.Height = Slider1.Value
End Sub

Private Sub Form_Load()
ct_form.Caption = "激舞小緣檔屏工具v1.5"
ct_form.Left = Screen.Width / 2 - (ct_form.Width / 2)
ct_form.Top = 380
Cover_L.Left = 0
Cover_R.Left = Screen.Width - Cover_R.Width
Cover_L.Top = Screen.Height - Cover_L.Height - 1800
Cover_R.Top = Cover_L.Top
Slider2.Max = Screen.Height
covervalue.Value = Cover_L.Width
Slider1.Value = Cover_L.Height
Slider2.Value = Cover_L.Top
Slider2.Max = Screen.Height
Slider1.Max = Screen.Height
Text3.Text = Slider2.Value
Text2.Text = Slider1.Value
Text1.Text = covervalue.Value
covervalue.Max = Screen.Width / 2
Cover_L.Show
Cover_R.Show
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
End
End Sub
Private Sub covervalue_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Cover_L.Width = covervalue.Value
Cover_R.Width = covervalue.Value
Cover_R.Left = Screen.Width - Cover_R.Width

End Sub
Private Sub covervalue_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Cover_L.Width = covervalue.Value
Cover_R.Width = covervalue.Value
Cover_R.Left = Screen.Width - Cover_R.Width
Text1.Text = covervalue.Value
End Sub




Private Sub Label5_Click(Index As Integer)
Select Case Index
Case 0
Cover_L.BackColor = &H0&
Cover_R.BackColor = Cover_L.BackColor
Free_L.BackColor = &H0&
Free_R.BackColor = Free_L.BackColor

Case 1
Cover_L.BackColor = &HFFFFFF
Cover_R.BackColor = Cover_L.BackColor
Free_L.BackColor = &HFFFFFF
Free_R.BackColor = Free_L.BackColor

Case 2
Cover_L.BackColor = &H808080
Cover_R.BackColor = Cover_L.BackColor
Free_L.BackColor = &H808080
Free_R.BackColor = Free_L.BackColor

Case 3
Cover_L.BackColor = &HFF&
Cover_R.BackColor = Cover_L.BackColor
Free_L.BackColor = &HFF&
Free_R.BackColor = Free_L.BackColor


Case 4
Cover_L.BackColor = &HFF00&
Cover_R.BackColor = Cover_L.BackColor
Free_L.BackColor = &HFF00&
Free_R.BackColor = Free_L.BackColor


Case 5
Cover_L.BackColor = &HFF0000
Cover_R.BackColor = Cover_L.BackColor
Free_L.BackColor = &HFF0000
Free_R.BackColor = Free_L.BackColor


Case 6
Cover_L.BackColor = &HFFFF&
Cover_R.BackColor = Cover_L.BackColor
Free_L.BackColor = &HFFFF&
Free_R.BackColor = Free_L.BackColor



Case 7
Cover_L.BackColor = &HFFFF00
Cover_R.BackColor = Cover_L.BackColor
Free_L.BackColor = &HFFFF00
Free_R.BackColor = Free_L.BackColor



End Select
End Sub


Private Sub showleft_Click()
Cover_L.Width = covervalue.Value
Cover_R.Width = covervalue.Value
Cover_R.Left = Screen.Width - Cover_R.Width

Cover_L.Height = Slider1.Value
Cover_R.Height = Slider1.Value

Cover_L.Top = Slider2.Value
Cover_R.Top = Slider2.Value
Cover_L.Show
showleft.Enabled = False
End Sub
Private Sub showright_Click()
Cover_L.Width = covervalue.Value
Cover_R.Width = covervalue.Value
Cover_R.Left = Screen.Width - Cover_R.Width

Cover_L.Height = Slider1.Value
Cover_R.Height = Slider1.Value

Cover_L.Top = Slider2.Value
Cover_R.Top = Slider2.Value
Cover_R.Show
showright.Enabled = False
End Sub

Private Sub Slider1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Cover_L.Height = Slider1.Value
Cover_R.Height = Slider1.Value

End Sub
Private Sub Slider1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Cover_L.Height = Slider1.Value
Cover_R.Height = Slider1.Value
Text2.Text = Slider1.Value
End Sub

Private Sub Slider2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Cover_L.Top = Slider2.Value
Cover_R.Top = Slider2.Value

End Sub
Private Sub Slider2_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Cover_L.Top = Slider2.Value
Cover_R.Top = Slider2.Value
Text3.Text = Slider2.Value
End Sub
