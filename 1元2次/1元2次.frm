VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "一元二次"
   ClientHeight    =   3540
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7140
   BeginProperty Font 
      Name            =   "新細明體"
      Size            =   9
      Charset         =   136
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3540
   ScaleWidth      =   7140
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command3 
      Caption         =   "結束"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4560
      TabIndex        =   8
      Top             =   1440
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "清除"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2640
      TabIndex        =   7
      Top             =   1440
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "計算"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   720
      TabIndex        =   6
      Top             =   1440
      Width           =   1455
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  '置中對齊
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   20.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4560
      TabIndex        =   2
      Top             =   600
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  '置中對齊
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   20.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   1
      Top             =   600
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  '置中對齊
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   20.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label5 
      Alignment       =   2  '置中對齊
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   18
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   10
      Top             =   2880
      Width           =   6015
   End
   Begin VB.Label Label4 
      Alignment       =   2  '置中對齊
      Caption         =   " = 0"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   21.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6000
      TabIndex        =   9
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label3 
      Alignment       =   2  '置中對齊
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   18
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   5
      Top             =   2280
      Width           =   6015
   End
   Begin VB.Label Label2 
      Alignment       =   2  '置中對齊
      Caption         =   "X+"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   21.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3960
      TabIndex        =   4
      Top             =   600
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "X^2+"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   20.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      TabIndex        =   3
      Top             =   600
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

a = Text1
b = Text2
c = Text3

If a = "" Or b = "" Or c = "" Then Label3 = "請輸入完整"

If a <> "" And b <> "" And c <> "" Then f = b ^ 2 - 4 * a * c

If f <> "" Then If a = 0 And b = 0 And c = 0 Then Label3 = "X為任意數"
If f <> "" Then If a = 0 And b = 0 And c <> 0 Then Label3 = "無解"
If f <> "" Then If a = 0 And b <> 0 Then Label3 = "X = " & -c / b
If f <> "" Then If f > 0 And a <> 0 Then Label3 = "X = " & (-b + (b ^ 2 - 4 * a * c) ^ 0.5) / (2 * a)
If f <> "" Then If f > 0 And a <> 0 Then Label5 = "X = " & (-b - (b ^ 2 - 4 * a * c) ^ 0.5) / (2 * a)
If f <> "" Then If f = 0 And a <> 0 Then Label3 = "X = " & (-b - (b ^ 2 - 4 * a * c) ^ 0.5) / (2 * a)
If f <> "" Then If f < 0 And a <> 0 Then Label3 = "無實數解"
If f <> "" Then If f <= 0 Or a = 0 Then Label5 = ""

End Sub

Private Sub Command2_Click()

Text1 = ""
Text2 = ""
Text3 = ""
Label3 = ""
Label5 = ""

End Sub

Private Sub Command3_Click()

End

End Sub

