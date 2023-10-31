VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "通訊錄"
   ClientHeight    =   3945
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6375
   LinkTopic       =   "Form1"
   ScaleHeight     =   3945
   ScaleWidth      =   6375
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command3 
      Caption         =   "結束"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   15.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      TabIndex        =   15
      Top             =   3240
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "顯示"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   15.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   14
      Top             =   3240
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "儲存"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   15.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      TabIndex        =   13
      Top             =   3240
      Width           =   855
   End
   Begin VB.ComboBox Combo3 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   15.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4320
      TabIndex        =   12
      Text            =   "1"
      Top             =   960
      Width           =   1095
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   15.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2760
      TabIndex        =   11
      Text            =   "1"
      Top             =   960
      Width           =   1095
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   15.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1200
      TabIndex        =   10
      Text            =   "80"
      Top             =   960
      Width           =   975
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   15.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   6
      Top             =   2520
      Width           =   3615
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   15.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      TabIndex        =   5
      Top             =   1680
      Width           =   3855
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   15.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   4
      Top             =   240
      Width           =   2775
   End
   Begin VB.Label Label7 
      Alignment       =   2  '置中對齊
      Caption         =   "日"
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
      Left            =   5520
      TabIndex        =   9
      Top             =   960
      Width           =   495
   End
   Begin VB.Label Label6 
      Alignment       =   2  '置中對齊
      Caption         =   "月"
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
      Left            =   3840
      TabIndex        =   8
      Top             =   960
      Width           =   495
   End
   Begin VB.Label Label5 
      Alignment       =   2  '置中對齊
      Caption         =   "年"
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
      Left            =   2280
      TabIndex        =   7
      Top             =   960
      Width           =   495
   End
   Begin VB.Label Label4 
      Alignment       =   2  '置中對齊
      Caption         =   "手機(電話):"
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
      Left            =   240
      TabIndex        =   3
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Label Label3 
      Alignment       =   2  '置中對齊
      Caption         =   "E-Mail:"
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
      Left            =   240
      TabIndex        =   2
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  '置中對齊
      Caption         =   "生日:"
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
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  '置中對齊
      Caption         =   "姓名:"
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
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a(1 To 50, 1 To 4) As String
Public n
Private Sub Command1_Click()
n = n + 1
a(n, 1) = Text1
a(n, 2) = Combo1 & " / " & Combo2 & " / " & Combo3
a(n, 3) = Text2
a(n, 4) = Text3
MsgBox "第" & n & "筆資料輸入完成", , "儲存聯絡人資料"

Text1 = ""
Text2 = ""
Text3 = ""
Combo1.ListIndex = 0
Combo2.ListIndex = 0
Combo3.ListIndex = 0

For i = 1 To n
Write #1, a(i, 1), a(i, 2), a(i, 3), a(i, 4)
Next i


End Sub
Private Sub Command2_Click()
Close #1
Form1.Hide
Form2.Show
Form2.Cls
Form2.Print " 姓 名 ", "生 日", "電子信箱", "手機(電話)"
Form2.Print

For i = 1 To n
Form2.Print a(i, 1), a(i, 2), a(i, 3), a(i, 4)
Next i

n = 0


End Sub

Private Sub Command3_Click()
Close #1
End
End Sub

Private Sub Form_Activate()
Form2.FontSize = 14
n = 0
For i = 80 To 90
   
   Combo1.AddItem i
   
Next i
For i = 1 To 12
   
   Combo2.AddItem i
   
Next i
For i = 1 To 31
   
   Combo3.AddItem i
   
Next i

Open "資ㄧ1.txt" For Append As #1
End Sub

