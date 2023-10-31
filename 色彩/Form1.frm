VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5325
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   ScaleHeight     =   5325
   ScaleWidth      =   6030
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command2 
      Caption         =   "白色"
      Height          =   495
      Left            =   4200
      TabIndex        =   8
      Top             =   4200
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "黑色"
      Height          =   495
      Left            =   3120
      TabIndex        =   7
      Top             =   4200
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  '平面
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   3135
      Left            =   2880
      ScaleHeight     =   3105
      ScaleWidth      =   2505
      TabIndex        =   3
      Top             =   720
      Width           =   2535
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   375
      Index           =   2
      Left            =   240
      Max             =   255
      TabIndex        =   2
      Top             =   3120
      Width           =   1935
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   375
      Index           =   1
      Left            =   240
      Max             =   255
      TabIndex        =   1
      Top             =   2160
      Width           =   1935
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   375
      Index           =   0
      Left            =   240
      Max             =   255
      TabIndex        =   0
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  '置中對齊
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   15.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   2280
      TabIndex        =   6
      Top             =   3120
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  '置中對齊
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   15.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   2280
      TabIndex        =   5
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  '置中對齊
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   15.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   2280
      TabIndex        =   4
      Top             =   1200
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Picture1.BackColor = RGB(0, 0, 0)
For i = 0 To 2
HScroll1(i).Value = 0
Label1(i) = 0
Next i
End Sub

Private Sub Command2_Click()
Picture1.BackColor = RGB(255, 255, 255)
For i = 0 To 2
HScroll1(i).Value = 255
Label1(i) = 255
Next i

End Sub




Private Sub HScroll1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Label1(Index) = HScroll1(Index).Value
Picture1.BackColor = RGB(Label1(0), Label1(1), Label1(2))
End Sub

Private Sub HScroll1_Scroll(Index As Integer)
Label1(Index) = HScroll1(Index).Value
Picture1.BackColor = RGB(Label1(0), Label1(1), Label1(2))
End Sub
