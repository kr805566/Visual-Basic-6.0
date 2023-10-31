VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "記憶力大考驗"
   ClientHeight    =   4560
   ClientLeft      =   9030
   ClientTop       =   750
   ClientWidth     =   5370
   Icon            =   "Form1(記憶力大考驗).frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4560
   ScaleWidth      =   5370
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1320
      Top             =   3480
   End
   Begin VB.CommandButton end 
      Caption         =   "結束"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4320
      TabIndex        =   5
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton 看答案 
      Caption         =   "看答案"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4320
      TabIndex        =   4
      Top             =   2040
      Width           =   855
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   840
      Top             =   3480
   End
   Begin VB.CommandButton GO 
      Caption         =   "GO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4320
      TabIndex        =   3
      Top             =   840
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  '置中對齊
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   2040
      Width           =   4095
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  '置中對齊
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   4095
   End
   Begin VB.Label Label4 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   18
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   8
      Top             =   3000
      Width           =   3375
   End
   Begin VB.Label Label3 
      Caption         =   "請輸入印象中的5個數字"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   1680
      Width           =   2775
   End
   Begin VB.Label Label2 
      Caption         =   "按「看答案」鍵會顯示上述數字，並在5秒後消失"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   3960
      Width           =   5175
   End
   Begin VB.Label Label1 
      Caption         =   "按GO鍵會顯示5個數字，並在３秒後消失"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   4335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim d As Integer
Private Sub end_Click()
End
End Sub

Private Sub Form_Activate()



Text1.Enabled = False
Text2.Enabled = False

看答案.Enabled = False



End Sub

Private Sub GO_Click()
d = 0

Dim N As Integer

Randomize Time

While i < 5

N = Int(Rnd() * 99) + 1
Text1 = Text1 & N & "  "
i = i + 1
Wend
Timer1.Enabled = True
GO.Enabled = False

Label4 = "數字 3 秒後消失"


End Sub
Private Sub Timer1_Timer()

d = d + 1

Label4 = "數字 " & 3 - d & " 秒後消失"
If d = 3 Then

Text1.Visible = False

看答案.Enabled = True

Timer1.Enabled = False

Text2.Enabled = True
Text2.SetFocus

Label4 = ""


End If


End Sub

Private Sub Timer2_Timer()


d = d + 1


Label4 = 5 - d & " 秒後初始化"

If d = 5 Then
Text1 = ""
Text2 = ""
GO.Enabled = True

看答案.Enabled = False

Timer2.Enabled = False
Text2.Enabled = False
Label4 = ""
GO.SetFocus

End If
End Sub

Private Sub 看答案_Click()
d = 0
Text1.Visible = True
Timer2.Enabled = True
Label4 = " 5 秒後初始化"
End Sub
