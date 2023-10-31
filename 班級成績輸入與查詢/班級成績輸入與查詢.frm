VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "班級成績輸入與查詢"
   ClientHeight    =   1965
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4650
   LinkTopic       =   "Form1"
   ScaleHeight     =   1965
   ScaleWidth      =   4650
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command2 
      Caption         =   "查詢"
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
      Left            =   3720
      TabIndex        =   5
      Top             =   1200
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "輸入"
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
      Left            =   3720
      TabIndex        =   4
      Top             =   480
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  '置中對齊
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
      Left            =   2160
      TabIndex        =   3
      Top             =   1200
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  '置中對齊
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
      Left            =   2160
      TabIndex        =   2
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   2  '置中對齊
      Caption         =   "程式成績:"
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
      Left            =   360
      TabIndex        =   1
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "座號(1~50):"
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
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim c As Integer
Dim a(1 To 50), b(1 To 50) As Single

Private Sub Command1_Click()

If Val(Text1) > 50 Then MsgBox "請輸入1~50", 32, "錯誤": Text1 = "": Exit Sub

c = c + 1

a(c) = Text1
b(c) = Text2

Text1 = ""
Text2 = ""

End Sub

Private Sub Command2_Click()
d = Text1
For i = 1 To 50
If a(i) = d Then Exit For
Next i
If i = 51 Then MsgBox "尚未輸入成績", 32, "錯誤": Exit Sub

Text2 = b(i)
End Sub

Private Sub Form_Activate()
c = 0
End Sub

