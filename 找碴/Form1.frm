VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "眼力大考驗"
   ClientHeight    =   6915
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14010
   FillColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   6915
   ScaleWidth      =   14010
   StartUpPosition =   3  '系統預設值
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   3840
      Top             =   5640
   End
   Begin VB.CommandButton Command1 
      Caption         =   "結束"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   20.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4935
      Left            =   10560
      TabIndex        =   0
      Top             =   600
      Width           =   3135
   End
   Begin VB.Label Label8 
      Alignment       =   2  '置中對齊
      BackStyle       =   0  '透明
      Height          =   375
      Left            =   9960
      TabIndex        =   8
      Top             =   4680
      Width           =   375
   End
   Begin VB.Label Label7 
      Alignment       =   2  '置中對齊
      BackStyle       =   0  '透明
      Height          =   495
      Left            =   9840
      TabIndex        =   7
      Top             =   2880
      Width           =   495
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '透明
      Height          =   495
      Left            =   8400
      TabIndex        =   6
      Top             =   3840
      Width           =   375
   End
   Begin VB.Label Label5 
      Alignment       =   2  '置中對齊
      BackStyle       =   0  '透明
      Height          =   255
      Left            =   7320
      TabIndex        =   5
      Top             =   2280
      Width           =   495
   End
   Begin VB.Label Label4 
      Alignment       =   2  '置中對齊
      BackStyle       =   0  '透明
      Height          =   255
      Left            =   7440
      TabIndex        =   4
      Top             =   1920
      Width           =   255
   End
   Begin VB.Label Label3 
      Alignment       =   2  '置中對齊
      BackStyle       =   0  '透明
      Height          =   375
      Left            =   5760
      TabIndex        =   3
      Top             =   4320
      Width           =   615
   End
   Begin VB.Label Label2 
      Alignment       =   2  '置中對齊
      BackStyle       =   0  '透明
      Height          =   255
      Left            =   5880
      TabIndex        =   2
      Top             =   3600
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "請在右圖中，找出與左圖差異處(共7處)並單按滑鼠左鍵"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   18
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   480
      TabIndex        =   1
      Top             =   360
      Width           =   9975
   End
   Begin VB.Image Image2 
      Height          =   3435
      Left            =   5640
      Picture         =   "Form1.frx":0000
      Top             =   1680
      Width           =   4770
   End
   Begin VB.Image Image1 
      Height          =   3480
      Left            =   360
      Picture         =   "Form1.frx":628B
      Top             =   1680
      Width           =   4800
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a, d As Integer
Private Sub Command1_Click()

End

End Sub

Private Sub Label2_Click()
a = a + 1
MsgBox "找到" & a & "個", , "好棒歐!"
Label2.Enabled = False
Label2.BackStyle = 1
Label2.ForeColor = vbWhite
Label2.BackColor = vbBlack
Label2 = a & "個"
If a = 7 Then
MsgBox "花了" & d & "秒", , "好棒歐!"
End If


End Sub
Private Sub Label3_Click()
a = a + 1
MsgBox "找到" & a & "個", , "好棒歐!"
Label3.Enabled = False
Label3.BackStyle = 1
Label3.ForeColor = vbWhite
Label3.BackColor = vbBlack
Label3 = a & "個"
If a = 7 Then
MsgBox "花了" & d & "秒", , "好棒歐!"
End If

End Sub
Private Sub Label4_Click()
a = a + 1
MsgBox "找到" & a & "個", , "好棒歐!"
Label4.Enabled = False
Label4.BackStyle = 1
Label4.ForeColor = vbWhite
Label4.BackColor = vbBlack
Label4 = a & "個"
If a = 7 Then
MsgBox "花了" & d & "秒", , "好棒歐!"
End If
End Sub
Private Sub Label5_Click()
a = a + 1
MsgBox "找到" & a & "個", , "好棒歐!"
Label5.Enabled = False
Label5.BackStyle = 1
Label5.ForeColor = vbWhite
Label5.BackColor = vbBlack
Label5 = a & "個"
If a = 7 Then
MsgBox "花了" & d & "秒", , "好棒歐!"
End If
End Sub
Private Sub Label6_Click()
a = a + 1
MsgBox "找到" & a & "個", , "好棒歐!"
Label6.Enabled = False
Label6.BackStyle = 1
Label6.ForeColor = vbWhite
Label6.BackColor = vbBlack
Label6 = a & "個"
If a = 7 Then
MsgBox "花了" & d & "秒", , "好棒歐!"
End If
End Sub
Private Sub Label7_Click()
a = a + 1
MsgBox "找到" & a & "個", , "好棒歐!"
Label7.Enabled = False
Label7.BackStyle = 1
Label7.ForeColor = vbWhite
Label7.BackColor = vbBlack
Label7 = a & "個"
If a = 7 Then
MsgBox "花了" & d & "秒", , "好棒歐!"
End If
End Sub
Private Sub Label8_Click()
a = a + 1
MsgBox "找到" & a & "個", , "好棒歐!"
Label8.Enabled = False
Label8.BackStyle = 1
Label8.ForeColor = vbWhite
Label8.BackColor = vbBlack
Label8 = a & "個"
If a = 7 Then
MsgBox "花了" & d & "秒", , "好棒歐!"
End If
End Sub

Private Sub Timer1_Timer()
d = d + 1
End Sub
