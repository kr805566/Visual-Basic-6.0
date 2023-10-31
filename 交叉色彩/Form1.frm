VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5310
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   5310
   ScaleWidth      =   6585
   StartUpPosition =   3  '系統預設值
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   2040
      Top             =   2880
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   5400
      TabIndex        =   0
      Top             =   3600
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim R(1), g(1), B(1)
Form1.ScaleMode = 3

Randomize Time

For i = 0 To 1
R(i) = Rnd() * 256
g(i) = Rnd() * 256
B(i) = Rnd() * 256
Next i

For i = 0 To Form1.ScaleWidth

Line (i, Form1.ScaleHeight / 2 - 50)-(i, Form1.ScaleHeight / 2 + 50), RGB(R(1), g(1), B(1)), BF

Next i



For i = 0 To Form1.ScaleHeight

Line (Form1.ScaleWidth / 2 - 50, i)-(Form1.ScaleWidth / 2 + 50, i), RGB(R(0), g(0), B(0)), BF

Next i




End Sub

Private Sub Timer1_Timer()
Command1 = True
End Sub
