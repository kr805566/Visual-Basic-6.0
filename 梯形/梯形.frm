VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "梯形"
   ClientHeight    =   4260
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5670
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
   ScaleHeight     =   4260
   ScaleWidth      =   5670
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command3 
      Caption         =   "清除"
      Height          =   735
      Left            =   2160
      TabIndex        =   10
      Top             =   3240
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "結束"
      Height          =   735
      Left            =   3720
      TabIndex        =   9
      Top             =   3240
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "計算"
      Height          =   735
      Left            =   600
      TabIndex        =   8
      Top             =   3240
      Width           =   1335
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   2400
      TabIndex        =   7
      Top             =   2520
      Width           =   2655
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   2400
      TabIndex        =   6
      Top             =   1800
      Width           =   2655
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   2400
      TabIndex        =   5
      Top             =   1080
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   2400
      TabIndex        =   4
      Top             =   360
      Width           =   2655
   End
   Begin VB.Label Label5 
      Alignment       =   2  '置中對齊
      AutoSize        =   -1  'True
      Caption         =   "高度:"
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
      Left            =   840
      TabIndex        =   3
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label3 
      Alignment       =   2  '置中對齊
      AutoSize        =   -1  'True
      Caption         =   "梯形面積:"
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
      Left            =   360
      TabIndex        =   2
      Top             =   2520
      Width           =   1815
   End
   Begin VB.Label Label2 
      Alignment       =   2  '置中對齊
      AutoSize        =   -1  'True
      Caption         =   "下底:"
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
      Left            =   840
      TabIndex        =   1
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  '置中對齊
      AutoSize        =   -1  'True
      Caption         =   "上底:"
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
      Left            =   840
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

a = (Val(Text1) + Val(Text2)) * Val(Text3) / 2

If Val(Text1) > 0 And Val(Text2) > 0 And Val(Text3) > 0 Then Text4 = a

If Val(Text1) = 0 Then Text1 = "上底不能為零"
If Val(Text2) = 0 Then Text2 = "下底不能為零"
If Val(Text3) = 0 Then Text3 = "高度不能為零"

If Val(Text1) < 0 Then Text1 = "上底不能為負數"
If Val(Text2) < 0 Then Text2 = "下底不能為負數"
If Val(Text3) < 0 Then Text3 = "高度不能為負數"


End Sub

Private Sub Command2_Click()

End

End Sub

Private Sub Command3_Click()

Text1 = ""
Text2 = ""
Text3 = ""
Text4 = ""

End Sub

