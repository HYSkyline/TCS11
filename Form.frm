VERSION 5.00
Begin VB.Form Formmain 
   BackColor       =   &H80000016&
   Caption         =   "MainForm"
   ClientHeight    =   660
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   1920
   LinkTopic       =   "Form1"
   ScaleHeight     =   660
   ScaleWidth      =   1920
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Commandpause 
      Appearance      =   0  'Flat
      Caption         =   "空格键暂停"
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   600
      Top             =   120
   End
   Begin VB.Label Label 
      BackColor       =   &H80000000&
      Height          =   375
      Index           =   0
      Left            =   100
      TabIndex        =   0
      Top             =   100
      Width           =   375
   End
End
Attribute VB_Name = "Formmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i, j As Integer
Dim pausei As Integer
Dim row As Integer
Dim forward As String, forwardlast As String
Dim forwardnum As Integer, forwardlastnum As Integer
Dim pos() As Integer
Dim posinix, posiniy As Integer
Dim lengthini As Integer, lengthrec As Integer
Dim goalposnum As Integer

Private Sub Commandpause_Click()
    pausei = 1 - pausei
    If pausei = 0 Then
        forward = ""
    Else
        forward = forwardlast
    End If
    'Debug.Print forward; forwardlast; pausei
End Sub

Private Sub Commandpause_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 87 Then forward = "up"
    If KeyCode = 83 Then forward = "down"
    If KeyCode = 65 Then forward = "left"
    If KeyCode = 68 Then forward = "right"
End Sub

Private Sub Form_Activate()
    Call initial
End Sub

Sub initial()
    Dim speed As Integer
    Randomize
    row = InputBox("输入行数/列数", "Row/Column Input", 15)
    'row = 15
    
    Form1.Width = Label(0).Left + (Label(0).Width + 15) * (row + 1)
    Form1.Height = Label(0).Top + (Label(0).Height + 15) * (row + 1) + 1300
    Form1.Top = 0.5 * Screen.Height - 0.5 * Form1.Height
    Form1.Left = 0.5 * Screen.Width - 0.5 * Form1.Width
    Commandpause.Left = 100
    Commandpause.Width = (Label(0).Width + 15) * row - 15
    Commandpause.Top = Label(0).Top + (Label(0).Height + 15) * row + 20
    Commandpause.Height = 1000
    Call caboccur
    lengthini = InputBox("输入初始长度格数", "Initial Length Input", 4)
    lengthrec = lengthini
    'lengthini = 5

    speed = InputBox("输入速度值(0-100,100最快,0最慢" & vbCrLf & "千万不要尝试100,不要问为什么", "", 80)
    Timer1.Interval = Int(-9.9 * speed + 1000)
    'Timer1.Interval = 500
    Call ini
    Call goalgeneral
    
    Commandpause.SetFocus
    pausei = 1
End Sub

Sub ini()
    Form1.Cls
    For j = 0 To row - 1
        For i = 0 To row - 1
            Label(10 * j + i).BackColor = &H80000000
        Next i
    Next j
    
    i = Int(Rnd() * 2)
    If i = 0 Then
        posinix = Int(Rnd() * (row - 3 - lengthini)) + 2
        posiniy = Int(Rnd() * row)
        For i = 1 To lengthini
            ReDim Preserve pos(i)
            pos(i) = posiniy * row + posinix + i - 1
            Label(pos(i)).BackColor = RGB(255, 255, 255)
        Next i
        forward = "left"
    Else
        posinix = Int(Rnd() * row)
        posiniy = Int(Rnd() * (row - 3 - lengthini)) + 2
        For i = 1 To lengthini
            ReDim Preserve pos(i)
            pos(i) = posiniy * row + posinix + (i - 1) * row
            Label(pos(i)).BackColor = RGB(255, 255, 255)
        Next i
        forward = "up"
    End If
    
    'Debug.Print pos(1); pos(2); pos(3); pos(4)
End Sub

Sub goalgeneral()
    goalposnum = Int(Rnd() * row * row)
    Do Until goalgeneraljudge(goalposnum) = False
        goalposnum = Int(Rnd() * row * row)
    Loop
    Label(goalposnum).BackColor = RGB(0, 0, 0)
End Sub

Function goalgeneraljudge(goalposnum As Integer) As Boolean
    For j = 1 To lengthini
        If goalposnum = pos(j) Then
            goalgeneraljudge = True
        End If
    Next j
End Function

Sub caboccur()
    'Label(0).Caption = 0
    j = 0
    For i = 1 To row - 1
        Load Label(i + j * row)
        Label(i + j * row).Top = Label(i + j * row - 1).Top
        Label(i + j * row).Left = Label(i + j * row - 1).Left + Label(i + j * row - 1).Width + 15
        Label(i + j * row).BackColor = Label(i + j * row - 1).BackColor
        Label(i + j * row).Visible = True
        'Label(i + j * row).Caption = i + j * row
    Next i
    For j = 1 To row - 1
        Load Label(j * row)
        Label(j * row).Left = Label((j - 1) * row).Left
        Label(j * row).Top = Label((j - 1) * row).Top + Label((j - 1) * row).Height + 15
        Label(j * row).BackColor = Label((j - 1) * row).BackColor
        Label(j * row).Visible = True
        'Label(j * row).Caption = j * row
        For i = 1 To row - 1
            Load Label(i + j * row)
            Label(i + j * row).Top = Label(i + j * row - 1).Top
            Label(i + j * row).Left = Label(i + j * row - 1).Left + Label(i + j * row - 1).Width + 15
            Label(i + j * row).BackColor = Label(i + j * row - 1).BackColor
            Label(i + j * row).Visible = True
            'Label(i + j * row).Caption = i + j * row
        Next i
    Next j
End Sub

Private Sub Timer1_Timer()
    forwardnum = forwardval(forward)
    'Debug.Print forwardlast; forward
    If forwardnum + forwardlastnum = 0 Then forward = forwardlast
    Call forwardmove
    If forward <> "" Then forwardlast = forward
    forwardlastnum = forwardval(forwardlast)
    'Debug.Print pos(1); pos(2); pos(3); pos(4)
End Sub

Sub boundjudge()
    If forward = "left" And pos(1) - (pos(1) \ row) * row = 0 Then
        Label(pos(1)).BackColor = RGB(255, 0, 0)
        MsgBox ("超出界限,卒")
        Call refreshcode
    ElseIf forward = "up" And pos(1) < row Then
        Label(pos(1)).BackColor = RGB(255, 0, 0)
        MsgBox ("超出界限,卒")
        Call refreshcode
    ElseIf forward = "right" And pos(1) - (pos(1) \ row) * row = row - 1 Then
        Label(pos(1)).BackColor = RGB(255, 0, 0)
        MsgBox ("超出界限,卒")
        Call refreshcode
    ElseIf forward = "down" And pos(1) >= row * (row - 1) Then
        Label(pos(1)).BackColor = RGB(255, 0, 0)
        MsgBox ("超出界限,卒")
        Call refreshcode
    End If
End Sub

Function forwardval(forward As String) As Integer
    If forward = "left" Then
        forwardval = -1
    ElseIf forward = "up" Then
        forwardval = -2
    ElseIf forward = "right" Then
        forwardval = 1
    ElseIf forward = "down" Then
        forwardval = 2
    End If
End Function

Sub forwardmove()
    If forward = "left" Then
        Call boundjudge
        If pos(1) - 1 <> goalposnum Then
            Label(pos(lengthini)).BackColor = &H80000000
            For i = lengthini To 2 Step -1
                pos(i) = pos(i - 1)
            Next i
            pos(1) = pos(1) - 1
            Call selfeat
            Label(pos(1)).BackColor = RGB(255, 255, 255)
        Else
            lengthini = lengthini + 1
            ReDim Preserve pos(lengthini)
            For i = lengthini To 2 Step -1
                pos(i) = pos(i - 1)
            Next i
            pos(1) = pos(1) - 1
            Call selfeat
            Label(pos(1)).BackColor = RGB(255, 255, 255)
            Call goalgeneral
        End If
    ElseIf forward = "right" Then
        Call boundjudge
        If pos(1) + 1 <> goalposnum Then
            Label(pos(lengthini)).BackColor = &H80000000
            For i = lengthini To 2 Step -1
                pos(i) = pos(i - 1)
            Next i
            pos(1) = pos(1) + 1
            Call selfeat
            Label(pos(1)).BackColor = RGB(255, 255, 255)
        Else
            lengthini = lengthini + 1
            ReDim Preserve pos(lengthini)
            For i = lengthini To 2 Step -1
                pos(i) = pos(i - 1)
            Next i
            pos(1) = pos(1) + 1
            Call selfeat
            Label(pos(1)).BackColor = RGB(255, 255, 255)
            Call goalgeneral
        End If
    ElseIf forward = "up" Then
        Call boundjudge
        If pos(1) - row <> goalposnum Then
            Label(pos(lengthini)).BackColor = &H80000000
            For i = lengthini To 2 Step -1
                pos(i) = pos(i - 1)
            Next i
            pos(1) = pos(1) - row
            Call selfeat
            Label(pos(1)).BackColor = RGB(255, 255, 255)
        Else
            lengthini = lengthini + 1
            ReDim Preserve pos(lengthini)
            For i = lengthini To 2 Step -1
                pos(i) = pos(i - 1)
            Next i
            pos(1) = pos(1) - row
            Call selfeat
            Label(pos(1)).BackColor = RGB(255, 255, 255)
            Call goalgeneral
        End If
    ElseIf forward = "down" Then
        Call boundjudge
        If pos(1) + row <> goalposnum Then
            Label(pos(lengthini)).BackColor = &H80000000
            For i = lengthini To 2 Step -1
                pos(i) = pos(i - 1)
            Next i
            pos(1) = pos(1) + row
            Call selfeat
            Label(pos(1)).BackColor = RGB(255, 255, 255)
        Else
            lengthini = lengthini + 1
            ReDim Preserve pos(lengthini)
            For i = lengthini To 2 Step -1
                pos(i) = pos(i - 1)
            Next i
            pos(1) = pos(1) + row
            Call selfeat
            Label(pos(1)).BackColor = RGB(255, 255, 255)
            Call goalgeneral
        End If
    End If
End Sub

Sub selfeat()
    For i = 2 To lengthini
        If pos(1) = pos(i) Then
            Label(pos(1)).BackColor = RGB(255, 0, 0)
            MsgBox ("咬到自己,卒" & vbCrLf & "这都可以……SA11")
            Call refreshcode
            Exit For
        End If
    Next i
End Sub

Sub refreshcode()
    Dim ii As Integer, jj As Integer
    For ii = 0 To row - 1
        For jj = 0 To row - 1
            Label(ii * row + jj).BackColor = &H80000000
        Next jj
    Next ii
    lengthini = lengthrec
    Call ini
    Call goalgeneral
End Sub
