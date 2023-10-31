VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4470
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6165
   LinkTopic       =   "Form1"
   ScaleHeight     =   4470
   ScaleWidth      =   6165
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command7 
      Caption         =   "/"
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
      Left            =   3480
      TabIndex        =   10
      Top             =   3480
      Width           =   495
   End
   Begin VB.CommandButton Command6 
      Caption         =   "*"
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
      Left            =   2760
      TabIndex        =   9
      Top             =   3480
      Width           =   495
   End
   Begin VB.CommandButton Command5 
      Caption         =   "-"
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
      Left            =   2040
      TabIndex        =   8
      Top             =   3480
      Width           =   495
   End
   Begin VB.CommandButton Command4 
      Caption         =   "+"
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
      Left            =   1320
      TabIndex        =   7
      Top             =   3480
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "清除"
      Height          =   615
      Left            =   4440
      TabIndex        =   6
      Top             =   600
      Width           =   1455
   End
   Begin VB.TextBox Text2 
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
      Left            =   2160
      TabIndex        =   5
      Top             =   1440
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "結束"
      Height          =   615
      Left            =   4440
      TabIndex        =   3
      Top             =   1440
      Width           =   1455
   End
   Begin VB.TextBox Text1 
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
      Left            =   2160
      TabIndex        =   2
      Top             =   600
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Caption         =   "算數結果"
      Height          =   1095
      Left            =   840
      TabIndex        =   0
      Top             =   2280
      Width           =   3615
      Begin VB.Label Label3 
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
         Left            =   360
         TabIndex        =   11
         Top             =   360
         Width           =   3015
      End
   End
   Begin VB.Label Label2 
      Alignment       =   2  '置中對齊
      Caption         =   "Y="
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
      Left            =   600
      TabIndex        =   4
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  '置中對齊
      Caption         =   "X="
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
      Left            =   600
      TabIndex        =   1
      Top             =   600
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Command1_Click()
End
End Sub

Private Sub Command2_Click()

Text1 = ""
Text2 = ""

End Sub

Private Sub Command4_Click()
Label3 = Val(Text1) + Val(Text2)
End Sub

Private Sub Command5_Click()
Label3 = Val(Text1) - Val(Text2)
End Sub

Private Sub Command6_Click()
Label3 = Val(Text1) * Val(Text2)
End Sub

Private Sub Command7_Click()
Label3 = Val(Text1) / Val(Text2)
End Sub
