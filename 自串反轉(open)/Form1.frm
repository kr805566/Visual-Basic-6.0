VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "字串反轉"
   ClientHeight    =   1710
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6435
   BeginProperty Font 
      Name            =   "新細明體"
      Size            =   15.75
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   1710
   ScaleWidth      =   6435
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command1 
      Caption         =   "反轉"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2640
      TabIndex        =   1
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "新細明體"
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
      Width           =   5775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()



Open App.Path & "\text.txt" For Input As #1


Do Until EOF(1)

Line Input #1, a

i = Len(a)

While i >= 1
Text2 = Text2 & Mid(a, i, 1)
i = i - 1
Wend


Loop

Close

End Sub




