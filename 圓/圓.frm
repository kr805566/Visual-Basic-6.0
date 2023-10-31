VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5175
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6885
   BeginProperty Font 
      Name            =   "新細明體"
      Size            =   18
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5175
   ScaleWidth      =   6885
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command3 
      Caption         =   "清除"
      Height          =   735
      Left            =   2760
      TabIndex        =   12
      Top             =   4080
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "結束"
      Height          =   735
      Left            =   5040
      TabIndex        =   11
      Top             =   4080
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "計算"
      Height          =   735
      Left            =   480
      TabIndex        =   10
      Top             =   4080
      Width           =   1335
   End
   Begin VB.TextBox Text5 
      Height          =   495
      Left            =   2520
      TabIndex        =   9
      Top             =   3240
      Width           =   4095
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   2520
      TabIndex        =   8
      Top             =   2520
      Width           =   4095
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   2520
      TabIndex        =   7
      Top             =   1800
      Width           =   4095
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   2520
      TabIndex        =   6
      Top             =   1080
      Width           =   4095
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   2520
      TabIndex        =   5
      Top             =   360
      Width           =   1695
   End
   Begin VB.Label Label5 
      Alignment       =   2  '置中對齊
      AutoSize        =   -1  'True
      Caption         =   "圓面積:"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   20.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   960
      TabIndex        =   4
      Top             =   1800
      Width           =   1395
   End
   Begin VB.Label Label4 
      Alignment       =   2  '置中對齊
      AutoSize        =   -1  'True
      Caption         =   "球體表面積:"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   20.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   120
      TabIndex        =   3
      Top             =   3240
      Width           =   2235
   End
   Begin VB.Label Label3 
      Alignment       =   2  '置中對齊
      AutoSize        =   -1  'True
      Caption         =   "球體體積:"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   20.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   480
      TabIndex        =   2
      Top             =   2520
      Width           =   1815
   End
   Begin VB.Label Label2 
      Alignment       =   2  '置中對齊
      AutoSize        =   -1  'True
      Caption         =   "圓周長:"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   20.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   960
      TabIndex        =   1
      Top             =   1080
      Width           =   1395
   End
   Begin VB.Label Label1 
      Alignment       =   2  '置中對齊
      AutoSize        =   -1  'True
      Caption         =   "半徑:"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   20.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1320
      TabIndex        =   0
      Top             =   360
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

Dim a As Double
Dim b As Double
Dim c As Double
Dim d As Double

a = Val(Text1) * 2 * 3.1415926
b = Val(Text1) ^ 2 * 3.1415926
c = Val(Text1) ^ 3 * 4 * 3.1415926 / 3
d = Val(Text1) ^ 2 * 4 * 3.1415926

Text2 = a
Text3 = b
Text4 = c
Text5 = d

End Sub

Private Sub Command2_Click()

End

End Sub

Private Sub Command3_Click()

Text1 = ""
Text2 = ""
Text3 = ""
Text4 = ""
Text5 = ""

End Sub
