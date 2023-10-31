VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "最大公因數與最小公倍數"
   ClientHeight    =   2190
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5490
   LinkTopic       =   "Form1"
   ScaleHeight     =   2190
   ScaleWidth      =   5490
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command3 
      Caption         =   "結束"
      Height          =   495
      Left            =   3960
      TabIndex        =   6
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "清除"
      Height          =   615
      Left            =   3960
      TabIndex        =   5
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "輸入"
      Height          =   495
      Left            =   3960
      TabIndex        =   3
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  '置中對齊
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   1
      Top             =   720
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  '置中對齊
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   735
      Left            =   240
      TabIndex        =   4
      Top             =   1320
      Width           =   3375
   End
   Begin VB.Label Label1 
      Caption         =   "請輸入任意兩個正整數"
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
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim a, b, c As Integer

a = Val(Text1)
b = Val(Text2)

If a > b Then c = a: a = b: b = c

For i = a To 1 Step -1
   If a Mod i = 0 And b Mod i = 0 Then Exit For
Next i
       
Label2 = "最大公因數為 " & i & vbCrLf & "最小公倍數為 " & a * b / i

End Sub

Private Sub Command2_Click()
Text1 = ""
Text2 = ""
Label2 = ""
Text1.SetFocus
End Sub
Private Sub Command3_Click()
End
End Sub

