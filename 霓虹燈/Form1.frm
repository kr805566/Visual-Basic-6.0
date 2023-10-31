VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "霓虹燈"
   ClientHeight    =   5775
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6480
   LinkTopic       =   "Form1"
   ScaleHeight     =   5775
   ScaleWidth      =   6480
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command1 
      Caption         =   "END"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   20.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   2400
      TabIndex        =   9
      Top             =   4440
      Width           =   975
   End
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   960
      Top             =   4560
   End
   Begin VB.Label Label1 
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
      Height          =   495
      Index           =   8
      Left            =   5040
      TabIndex        =   8
      Top             =   480
      Width           =   495
   End
   Begin VB.Label Label1 
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
      Height          =   495
      Index           =   7
      Left            =   4440
      TabIndex        =   7
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label Label1 
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
      Height          =   495
      Index           =   6
      Left            =   3840
      TabIndex        =   6
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label Label1 
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
      Height          =   495
      Index           =   5
      Left            =   3240
      TabIndex        =   5
      Top             =   2280
      Width           =   495
   End
   Begin VB.Label Label1 
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
      Height          =   495
      Index           =   4
      Left            =   2640
      TabIndex        =   4
      Top             =   2880
      Width           =   495
   End
   Begin VB.Label Label1 
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
      Height          =   495
      Index           =   3
      Left            =   2040
      TabIndex        =   3
      Top             =   2280
      Width           =   495
   End
   Begin VB.Label Label1 
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
      Height          =   495
      Index           =   2
      Left            =   1440
      TabIndex        =   2
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label Label1 
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
      Height          =   495
      Index           =   1
      Left            =   840
      TabIndex        =   1
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label Label1 
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
      Height          =   495
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim B As Integer, 開關

Private Sub Command1_Click()
End
End Sub

Private Sub Form_Activate()

A = "大明高中資訊科張三"

For I = 0 To 8

Label1(I) = Mid(A, I + 1, 1)

Next I

開關 = True

End Sub

Private Sub Timer1_Timer()
If 開關 Then
    Label1(B).ForeColor = vbBlue
    Label1(B).FontSize = 12
    B = B + 1
    If B = 9 Then 開關 = False: B = 8
    Label1(B).ForeColor = vbRed
    Label1(B).FontSize = 20
Else
    Label1(B).ForeColor = vbBlue
    Label1(B).FontSize = 12
    B = B - 1
    If B = -1 Then 開關 = True: B = 0
    Label1(B).ForeColor = vbRed
    Label1(B).FontSize = 20
End If
End Sub

