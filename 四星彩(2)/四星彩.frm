VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "四星彩"
   ClientHeight    =   3150
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4950
   LinkTopic       =   "Form1"
   ScaleHeight     =   3150
   ScaleWidth      =   4950
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command1 
      Caption         =   "開始"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   15.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   3960
      TabIndex        =   16
      Top             =   840
      Width           =   855
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   60
      Left            =   4440
      Top             =   2520
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   60
      Left            =   3960
      Top             =   2520
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   72
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1455
      Index           =   3
      Left            =   3120
      TabIndex        =   11
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   72
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1455
      Index           =   2
      Left            =   2160
      TabIndex        =   10
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   72
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1455
      Index           =   1
      Left            =   1200
      TabIndex        =   9
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   72
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1455
      Index           =   3
      Left            =   3120
      TabIndex        =   8
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   72
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1455
      Index           =   2
      Left            =   2160
      TabIndex        =   7
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   72
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1455
      Index           =   1
      Left            =   1200
      TabIndex        =   6
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   72
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1455
      Index           =   3
      Left            =   3120
      TabIndex        =   5
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   72
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1455
      Index           =   2
      Left            =   2160
      TabIndex        =   4
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   72
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1455
      Index           =   1
      Left            =   1200
      TabIndex        =   3
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   72
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1455
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   72
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1455
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   72
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1455
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label4 
      BackColor       =   &H00E0E0E0&
      Height          =   2895
      Index           =   0
      Left            =   240
      TabIndex        =   12
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label4 
      BackColor       =   &H00E0E0E0&
      Height          =   2895
      Index           =   1
      Left            =   1200
      TabIndex        =   13
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label4 
      BackColor       =   &H00E0E0E0&
      Height          =   2895
      Index           =   2
      Left            =   2160
      TabIndex        =   14
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label4 
      BackColor       =   &H00E0E0E0&
      Height          =   2895
      Index           =   3
      Left            =   3120
      TabIndex        =   15
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim d As Integer
Private Sub Command1_Click()
d = 0
Timer1.Interval = 60
Timer2.Interval = 60


For i = 0 To 3
Randomize Time




Label1(i) = Int(Rnd() * 10)

Label2(i) = Int(Rnd() * 10)

Next i

Timer1.Enabled = True

End Sub

Private Sub Form_Activate()


For i = 0 To 3

Label1(i).Visible = False
Label3(i).Visible = False


Next i



End Sub


Private Sub Timer1_Timer()




For i = 0 To 3
Label1(i).Visible = True
Label3(i).Visible = True
Label2(i).Visible = False



Label3(i) = Label2(i)
Label2(i) = ""
Randomize Time


Label1(i) = Int(Rnd() * 10)

Next i


Timer2.Enabled = True
Timer1.Enabled = False



End Sub

Private Sub Timer2_Timer()
For i = 0 To 3
Label1(i).Visible = False
Label3(i).Visible = False
Label2(i).Visible = True

Label2(i) = Label1(i)

Next i

If d = 10 Then
Timer1.Enabled = False
Else
Timer1.Enabled = True
End If
Timer2.Enabled = False



d = d + 1

If d = 5 Then Timer1.Interval = 100: Timer2.Interval = 100


End Sub

