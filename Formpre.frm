VERSION 5.00
Begin VB.Form Formpre 
   Appearance      =   0  'Flat
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "准备"
   ClientHeight    =   2070
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3165
   DrawMode        =   1  'Blackness
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2070
   ScaleWidth      =   3165
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Textlength 
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   840
      TabIndex        =   7
      Text            =   "4"
      Top             =   600
      Width           =   2175
   End
   Begin VB.CommandButton Commandquit 
      Caption         =   "退出"
      Height          =   495
      Left            =   2400
      TabIndex        =   6
      Top             =   1440
      Width           =   615
   End
   Begin VB.CommandButton Commandreset 
      Caption         =   "重置"
      Height          =   495
      Left            =   1680
      TabIndex        =   5
      Top             =   1440
      Width           =   615
   End
   Begin VB.CommandButton Commandstart 
      Caption         =   "就绪"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   1455
   End
   Begin VB.HScrollBar HScrollspeed 
      Height          =   255
      Left            =   840
      Max             =   10
      TabIndex        =   3
      Top             =   960
      Width           =   2175
   End
   Begin VB.TextBox Textgrid 
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   840
      TabIndex        =   1
      Text            =   "15"
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label Labellength 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "长度:"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   300
      TabIndex        =   8
      Top             =   640
      Width           =   450
   End
   Begin VB.Label Labelspeed 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "速度:"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   360
      TabIndex        =   2
      Top             =   1000
      Width           =   450
   End
   Begin VB.Label Labelgrid 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "格网数:"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   225
      TabIndex        =   0
      Top             =   240
      Width           =   630
   End
End
Attribute VB_Name = "Formpre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Commandquit_Click()
End
End Sub

Private Sub Commandreset_Click()
Call reset
End Sub

Private Sub Commandstart_Click()
Dim speed As Integer
Dim lengthini As Integer
Dim row As Integer

row = Int(Textgrid.Text)
lengthini = Int(Textlength.Text)
speed = HScrollspeed.Value

If row < 3 Then
    MsgBox "需要行列数>3"
    Call reset
ElseIf lengthini >= row Then
    MsgBox "初始长度过大"
    Call reset
Else
    Form1.Show
End If

End Sub

Private Sub Form_Activate()
Formpre.Top = 0.5 * (Screen.Height - Formpre.Height)
Formpre.Left = 0.5 * (Screen.Width - Formpre.Width)
End Sub

Sub reset()
Textgrid.Text = "15"
Textlength.Text = "4"
HScrollspeed.Value = 0
Textgrid.SetFocus
End Sub
