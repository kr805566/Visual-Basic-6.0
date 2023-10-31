VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6315
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7665
   LinkTopic       =   "Form1"
   ScaleHeight     =   6315
   ScaleWidth      =   7665
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command4 
      Caption         =   "結束"
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
      Left            =   6120
      TabIndex        =   3
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "清除"
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
      Left            =   6120
      TabIndex        =   2
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "縮小"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   8.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6120
      TabIndex        =   1
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "放大"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   20.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6120
      TabIndex        =   0
      Top             =   360
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

Form1.FontSize = Form1.FontSize + 4
Form1.ForeColor = &HFF00FF
Form1.FontItalic = True
Form1.FontBold = False
Print "資ㄧ1 "
End Sub

Private Sub Command2_Click()

Form1.FontSize = Form1.FontSize - 4
Form1.ForeColor = &HFFFF00
Form1.FontItalic = False
Form1.FontBold = True
Print "資ㄧ1 "

End Sub

Private Sub Command3_Click()
Cls
Form1.FontSize = 12
Form1.ForeColor = vbRed
Form1.FontItalic = False
Form1.FontBold = False
Print "資ㄧ1 "
End Sub

Private Sub Command4_Click()
End
End Sub

Private Sub Form_Activate()
Form1.FontSize = 12
Form1.ForeColor = vbRed
Print "資ㄧ1 "

End Sub
