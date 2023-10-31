VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7215
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4125
   LinkTopic       =   "Form1"
   ScaleHeight     =   7215
   ScaleWidth      =   4125
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command7 
      Caption         =   "存N次資料到表格"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   12
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2760
      TabIndex        =   13
      Top             =   4920
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  '置中對齊
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   14.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   12
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "清除資料"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   12
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   11
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "儲存資料"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   12
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   10
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "匯入資料"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   12
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   9
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  '置中對齊
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   14.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4665
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  '垂直捲軸
      TabIndex        =   8
      Top             =   2280
      Width           =   2535
   End
   Begin VB.CommandButton Command3 
      Caption         =   "存入表格"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   12
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   7
      Top             =   1680
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  '置中對齊
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   14.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      Caption         =   "清除"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   14.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   5
      Top             =   960
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "隨機"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   14.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   4
      Top             =   480
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  '置中對齊
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   14.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Text            =   "9"
      Top             =   1080
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  '置中對齊
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   14.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   0
      Text            =   "1"
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "數字"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   18
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "英文"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   18
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   480
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
txt = ""
For i = 1 To Val(Text1)
txt = txt & Chr(Int(Rnd() * 26 + 97))
Next i

For i = 1 To Val(Text2)
txt = txt & Int(Rnd() * 10)
Next i

Text3 = txt

End Sub

Private Sub Command2_Click()
Text3 = ""
End Sub

Private Sub Command3_Click()
Text4 = Text4 & Text3 & vbCrLf
End Sub

Private Sub Command4_Click()
Open App.Path & "/1.txt" For Input As #1
Text4 = ""
Do Until EOF(1)

Line Input #1, a

Text4 = Text4 & a & vbCrLf

Loop

Close #1
End Sub

Private Sub Command5_Click()
Open App.Path & "/1.txt" For Output As #2


a = Text4

Print #2, a



Close #2
End Sub

Private Sub Command6_Click()
Text4 = ""
End Sub

Private Sub Command7_Click()
For i = 1 To Val(Text5)

Command1.Value = True
Command3.Value = True
Next i
End Sub

Private Sub Form_Load()


End Sub
