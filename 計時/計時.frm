VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "計時"
   ClientHeight    =   1875
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6960
   LinkTopic       =   "Form1"
   ScaleHeight     =   1875
   ScaleWidth      =   6960
   StartUpPosition =   3  '系統預設值
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   720
      Top             =   1440
   End
   Begin VB.CommandButton Command3 
      Caption         =   "歸零"
      Height          =   495
      Left            =   6000
      TabIndex        =   7
      Top             =   1320
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "暫停"
      Height          =   495
      Left            =   6000
      TabIndex        =   6
      Top             =   720
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "開始"
      Height          =   495
      Left            =   6000
      TabIndex        =   5
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label6 
      Alignment       =   2  '置中對齊
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   36
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5160
      TabIndex        =   10
      Top             =   480
      Width           =   615
   End
   Begin VB.Label Label5 
      Alignment       =   2  '置中對齊
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   36
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4440
      TabIndex        =   9
      Top             =   480
      Width           =   615
   End
   Begin VB.Label LabelB 
      Alignment       =   2  '置中對齊
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   36
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3960
      TabIndex        =   8
      Top             =   360
      Width           =   255
   End
   Begin VB.Label LabelA 
      Alignment       =   2  '置中對齊
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   36
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1920
      TabIndex        =   4
      Top             =   360
      Width           =   255
   End
   Begin VB.Label Label4 
      Alignment       =   2  '置中對齊
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   36
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3120
      TabIndex        =   3
      Top             =   480
      Width           =   615
   End
   Begin VB.Label Label3 
      Alignment       =   2  '置中對齊
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   36
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2400
      TabIndex        =   2
      Top             =   480
      Width           =   615
   End
   Begin VB.Label Label2 
      Alignment       =   2  '置中對齊
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   36
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1080
      TabIndex        =   1
      Top             =   480
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   2  '置中對齊
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   36
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim T As Integer

Private Sub Command1_Click()
Timer1 = True
End Sub

Private Sub Command2_Click()
Timer1 = False

End Sub

Private Sub Command3_Click()
Label1 = 0
Label2 = 0
Label3 = 0
Label4 = 0
Label5 = 0
Label6 = 0
T = 0
Timer1 = False

End Sub

Private Sub Timer1_Timer()

T = T + 1
A = T Mod 10
Label6 = A



If Val(Label6) = 0 Then
Label5 = Val(Label5) + 1
End If

If Val(Label5) = 6 Then
Label5 = 0
Label4 = Val(Label4) + 1
End If

If Val(Label4) = 10 Then
Label4 = 0
Label3 = Val(Label3) + 1
End If

If Val(Label3) = 6 Then
Label3 = 0
Label2 = Val(Label2) + 1
End If

If Val(Label2) = 10 Then
Label2 = 0
Label1 = Val(Label3) + 1
End If

End Sub
