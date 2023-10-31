VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2625
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4335
   LinkTopic       =   "Form1"
   ScaleHeight     =   2625
   ScaleWidth      =   4335
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command3 
      Caption         =   "結束"
      Height          =   375
      Left            =   3000
      TabIndex        =   6
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "清除"
      Height          =   375
      Left            =   1680
      TabIndex        =   5
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "計算"
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   1920
      Width           =   975
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
      Left            =   1680
      TabIndex        =   3
      Top             =   1080
      Width           =   2535
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
      Left            =   1680
      TabIndex        =   2
      Top             =   360
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  '置中對齊
      AutoSize        =   -1  'True
      Caption         =   "圓面積："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   18
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  '置中對齊
      AutoSize        =   -1  'True
      Caption         =   "半徑："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   18
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

Dim a As Double

a = Val(Text1) * Val(Text1) * 3.1415926

Text2 = a

End Sub

Private Sub Command2_Click()

Text1 = ""
Text2 = ""

End Sub

Private Sub Command3_Click()

End

End Sub
