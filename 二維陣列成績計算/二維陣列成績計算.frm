VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4425
   ClientLeft      =   4065
   ClientTop       =   4560
   ClientWidth     =   12870
   LinkTopic       =   "Form1"
   ScaleHeight     =   4425
   ScaleWidth      =   12870
   Begin VB.TextBox Text6 
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
      Left            =   10560
      TabIndex        =   20
      Top             =   2280
      Width           =   855
   End
   Begin VB.TextBox Text5 
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
      Left            =   10560
      TabIndex        =   19
      Top             =   1560
      Width           =   855
   End
   Begin VB.TextBox Text5 
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
      Left            =   9480
      TabIndex        =   18
      Top             =   1560
      Width           =   855
   End
   Begin VB.TextBox Text4 
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
      Left            =   6480
      TabIndex        =   17
      Top             =   2280
      Width           =   855
   End
   Begin VB.TextBox Text3 
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
      Left            =   6480
      TabIndex        =   16
      Top             =   1560
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
      Height          =   495
      Index           =   1
      Left            =   2280
      TabIndex        =   15
      Top             =   2280
      Width           =   855
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
      Height          =   495
      Index           =   1
      Left            =   2280
      TabIndex        =   14
      Top             =   1560
      Width           =   855
   End
   Begin VB.TextBox Text6 
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
      Left            =   9480
      TabIndex        =   13
      Top             =   2280
      Width           =   855
   End
   Begin VB.TextBox Text4 
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
      Left            =   5400
      TabIndex        =   12
      Top             =   2280
      Width           =   855
   End
   Begin VB.TextBox Text3 
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
      Left            =   5400
      TabIndex        =   11
      Top             =   1560
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
      Height          =   495
      Index           =   0
      Left            =   1200
      TabIndex        =   10
      Top             =   2280
      Width           =   855
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
      Height          =   495
      Index           =   0
      Left            =   1200
      TabIndex        =   9
      Top             =   1560
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "＝"
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
      Left            =   8040
      TabIndex        =   8
      Top             =   1920
      Width           =   495
   End
   Begin VB.Label Label4 
      Alignment       =   2  '置中對齊
      Caption         =   "二維陣列成績計算"
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
      Left            =   3480
      TabIndex        =   7
      Top             =   360
      Width           =   5415
   End
   Begin VB.Label Label2 
      Caption         =   "╳"
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
      Left            =   4080
      TabIndex        =   6
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   27.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Index           =   5
      Left            =   11640
      TabIndex        =   5
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   27.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Index           =   4
      Left            =   8760
      TabIndex        =   4
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   27.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Index           =   3
      Left            =   7560
      TabIndex        =   3
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   27.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Index           =   2
      Left            =   4680
      TabIndex        =   2
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   27.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Index           =   1
      Left            =   3360
      TabIndex        =   1
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   27.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Index           =   0
      Left            =   480
      TabIndex        =   0
      Top             =   1080
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim A(1, 1), B(1, 1) As Integer

For I = 0 To 1
    A(0, I) = Val(Text1(I))
    A(1, I) = Val(Text2(I))
    B(0, I) = Val(Text3(I))
    B(1, I) = Val(Text4(I))
Next I

Text5(0) = A(0, 0) * B(0, 0) + A(0, 1) * B(1, 0)
Text5(1) = A(0, 0) * B(0, 1) + A(0, 1) * B(1, 1)
Text6(0) = A(1, 0) * B(0, 0) + A(1, 1) * B(1, 0)
Text6(1) = A(1, 0) * B(0, 1) + A(1, 1) * B(1, 1)

End Sub

Private Sub Form_Activate()
For I = 0 To 2
Label1(2 * I) = "┌" & vbCrLf & "│" & vbCrLf & "│" & vbCrLf & "└"
Label1(2 * I + 1) = "┐" & vbCrLf & "│" & vbCrLf & "│" & vbCrLf & "┘"
Next I

Text1(0).SetFocus

End Sub


