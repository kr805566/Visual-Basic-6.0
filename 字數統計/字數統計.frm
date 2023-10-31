VERSION 5.00
Begin VB.Form 计参璸 
   Caption         =   "Form1"
   ClientHeight    =   4425
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7920
   LinkTopic       =   "Form1"
   ScaleHeight     =   4425
   ScaleWidth      =   7920
   StartUpPosition =   3  '╰参箇砞
   Begin VB.TextBox Text1 
      Height          =   2055
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   3  'ㄢΤ
      TabIndex        =   5
      Top             =   960
      Width           =   6735
   End
   Begin VB.CommandButton Command3 
      Caption         =   "挡"
      Height          =   615
      Left            =   5640
      TabIndex        =   4
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "睲埃"
      Height          =   615
      Left            =   4080
      TabIndex        =   3
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton Form1 
      Caption         =   "计参璸"
      Height          =   615
      Left            =   2520
      TabIndex        =   2
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label Label2 
      Height          =   615
      Left            =   360
      TabIndex        =   1
      Top             =   3240
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "叫块ゅ"
      BeginProperty Font 
         Name            =   "穝灿砰"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "计参璸"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

c = Len(Text1)
Label2 = "计" & c & ""

End Sub

Private Sub Command2_Click()
Label2 = ""
Text1 = ""
Text1.SetFocus

End Sub

Private Sub Command3_Click()

End

End Sub

Private Sub Form_Activate()
Text1.SetFocus
End Sub
