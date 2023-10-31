VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0C0C0&
   Caption         =   "CURVE"
   ClientHeight    =   4605
   ClientLeft      =   1110
   ClientTop       =   1545
   ClientWidth     =   8445
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   PaletteMode     =   1  '最上層控制項的調色盤
   ScaleHeight     =   4605
   ScaleWidth      =   8445
   Begin VB.CommandButton Command4 
      Caption         =   "csc"
      Height          =   255
      Left            =   7800
      TabIndex        =   10
      Top             =   4080
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      Caption         =   "sec"
      Height          =   255
      Left            =   7800
      TabIndex        =   9
      Top             =   3720
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "cot"
      Height          =   255
      Left            =   7080
      TabIndex        =   8
      Top             =   4080
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "tan"
      Height          =   255
      Left            =   7080
      TabIndex        =   7
      Top             =   3720
      Width           =   615
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "清除"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      TabIndex        =   6
      Top             =   3720
      Width           =   960
   End
   Begin VB.Frame Frame1 
      Caption         =   "函數圖形"
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   3600
      Width           =   3255
      Begin VB.OptionButton optCos 
         Caption         =   "Cos函數圖形"
         Height          =   255
         Left            =   1680
         TabIndex        =   5
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton optSin 
         Caption         =   "Sin函數圖形"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Value           =   -1  'True
         Width           =   1455
      End
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   3420
      Left            =   120
      ScaleHeight     =   3360
      ScaleWidth      =   8085
      TabIndex        =   2
      Top             =   120
      Width           =   8145
   End
   Begin VB.CommandButton cmdEnd 
      Caption         =   "結束"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5925
      TabIndex        =   1
      Top             =   3720
      Width           =   960
   End
   Begin VB.CommandButton cmdDraw 
      Caption         =   "畫出"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   0
      Top             =   3720
      Width           =   960
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClear_Click()
Picture1.Cls
End Sub

Private Sub cmdDraw_Click()
Picture1.Line (0, 0)-(0, Picture1.ScaleHeight), vbRed
Picture1.Line (0, Picture1.ScaleHeight / 2)-(Picture1.ScaleWidth, Picture1.ScaleHeight / 2), vbRed


If optSin = True Then

For i = 0 To 3.14159 * 2 Step 0.01

Picture1.PSet (i * 1000, -Sin(i) * 1000 + Picture1.ScaleHeight / 2), vbGreen

Next i

Else

For i = 0 To 3.14159 * 2 Step 0.01

Picture1.PSet (i * 1000, -Cos(i) * 1000 + Picture1.ScaleHeight / 2), vbBlue

Next i

End If

End Sub

Private Sub Command1_Click()

Picture1.Line (0, 0)-(0, Picture1.ScaleHeight), vbRed
Picture1.Line (0, Picture1.ScaleHeight / 2)-(Picture1.ScaleWidth, Picture1.ScaleHeight / 2), vbRed

For i = 0 To 3.14159 * 2 Step 0.001

Picture1.PSet (i * 1000, -Tan(i) * 1000 + Picture1.ScaleHeight / 2), vbBlack

Next i
End Sub

Private Sub Command2_Click()
Picture1.Line (0, 0)-(0, Picture1.ScaleHeight), vbRed
Picture1.Line (0, Picture1.ScaleHeight / 2)-(Picture1.ScaleWidth, Picture1.ScaleHeight / 2), vbRed

For i = 0 To 3.14159 * 2 Step 0.001

If -Tan(i) <> 0 Then
Picture1.PSet (i * 1000, (1 / -Tan(i)) * 1000 + Picture1.ScaleHeight / 2), vbBlack
End If
Next i
End Sub

Private Sub Command3_Click()
Picture1.Line (0, 0)-(0, Picture1.ScaleHeight), vbRed
Picture1.Line (0, Picture1.ScaleHeight / 2)-(Picture1.ScaleWidth, Picture1.ScaleHeight / 2), vbRed
For i = 0 To 3.14159 * 2 Step 0.01

Picture1.PSet (i * 1000, (1 / -Cos(i)) * 1000 + Picture1.ScaleHeight / 2), vbBlue

Next i

End Sub

Private Sub Command4_Click()
Picture1.Line (0, 0)-(0, Picture1.ScaleHeight), vbRed
Picture1.Line (0, Picture1.ScaleHeight / 2)-(Picture1.ScaleWidth, Picture1.ScaleHeight / 2), vbRed

For i = 0 To 3.14159 * 2 Step 0.01
If -Sin(i) <> 0 Then
Picture1.PSet (i * 1000, (1 / -Sin(i)) * 1000 + Picture1.ScaleHeight / 2), vbBlue
End If
Next i

End Sub
