VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5295
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5130
   LinkTopic       =   "Form1"
   ScaleHeight     =   5295
   ScaleWidth      =   5130
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command3 
      Caption         =   "結束"
      Height          =   375
      Left            =   3960
      TabIndex        =   8
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "重新輸入"
      Height          =   375
      Left            =   3960
      TabIndex        =   7
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "輸入成績"
      Height          =   495
      Left            =   3960
      TabIndex        =   6
      Top             =   600
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Caption         =   "計算結果"
      Height          =   2655
      Left            =   120
      TabIndex        =   4
      Top             =   2520
      Width           =   4815
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2175
         Left            =   360
         MultiLine       =   -1  'True
         ScrollBars      =   1  '水平捲軸
         TabIndex        =   5
         Top             =   240
         Width           =   4095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "請先勾選項目"
      Height          =   1935
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   3255
      Begin VB.CheckBox Check3 
         Caption         =   "計算及格與不及格科數"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   1200
         Width           =   2535
      End
      Begin VB.CheckBox Check2 
         Caption         =   "找出最高分與最低分"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   2535
      End
      Begin VB.CheckBox Check1 
         Caption         =   "計算總分與平均"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2535
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As Integer
Private Sub Check1_Click()

If Check1 = 1 Or Check2 = 1 Or Check3 = 1 Then
Command1.Enabled = True
Else
Command1.Enabled = False
End If
End Sub
Private Sub Check2_Click()

If Check1 = 1 Or Check2 = 1 Or Check3 = 1 Then
Command1.Enabled = True
Else
Command1.Enabled = False
End If
End Sub
Private Sub Check3_Click()

If Check1 = 1 Or Check2 = 1 Or Check3 = 1 Then
Command1.Enabled = True
Else
Command1.Enabled = False
End If
End Sub
Private Sub Command1_Click()
Dim su, max, m, sco, n As Integer

sun = 0
max = 0
m = 100
n = 0
Text1 = "段考成績  "

For i = 1 To a
sco = Val(InputBox("輸入第" & i & "段考成績", "成績統計"))
su = su + sco
Text1 = Text1 & sco & "分" & Space(2)
If sco > max Then max = sco
If sco < m Then m = sco
If sco >= 60 Then n = n + 1
Next i


If Check1.Value = 1 Then Text1 = Text1 & vbCrLf & a & "科總分為 " & su & "分" & vbCrLf & a & "科平均為 " & Format(su / a, "0.0") & "分"
If Check2.Value = 1 Then Text1 = Text1 & vbCrLf & "最高分為 " & max & "分" & vbCrLf & "最低分為 " & m & "分"
If Check3.Value = 1 Then Text1 = Text1 & vbCrLf & "及格科數為 " & n & "科" & vbCrLf & "不及格科數為 " & a - n & "科"
End Sub

Private Sub Command2_Click()
Text1 = ""
Check1.Value = 0
Check2.Value = 0
Check3.Value = 0
a = InputBox("輸入科數", "成績統計")
Command1.Enabled = False


End Sub



Private Sub Command3_Click()
End

End Sub

Private Sub Form_Activate()

a = InputBox("輸入科數", "成績統計")
Command1.Enabled = False


End Sub

