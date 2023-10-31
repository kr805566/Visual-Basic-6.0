VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "數學程式"
   ClientHeight    =   2190
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3585
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2190
   ScaleWidth      =   3585
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command3 
      Caption         =   "判斷是否互質"
      Height          =   495
      Left            =   840
      TabIndex        =   2
      Top             =   1440
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "判斷是否為質數"
      Height          =   495
      Left            =   840
      TabIndex        =   1
      Top             =   840
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "找出範圍內的質數"
      Height          =   495
      Left            =   840
      TabIndex        =   0
      Top             =   240
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form1.Hide
Form2.Show

End Sub

Private Sub Command2_Click()

Form1.Hide
Form3.Show


End Sub

Private Sub Command3_Click()

Form1.Hide
Form4.Show


End Sub
Private Sub Form_Load()
Call lineGradient(240, 60, 50, 30, 40, 60)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
End
End Sub

Sub lineGradient(R1 As Single, G1 As Single, B1 As Single, R2 As Single, G2 As Single, B2 As Single)


Form1.AutoRedraw = True
Form1.DrawWidth = 1
Form1.ScaleMode = 1


RD = (R2 - R1) / Form1.ScaleHeight * Form1.DrawWidth
GD = (G2 - G1) / Form1.ScaleHeight * Form1.DrawWidth
BD = (B2 - B1) / Form1.ScaleHeight * Form1.DrawWidth



For I = 0 To Form1.ScaleHeight
If R1 > 255 Then R1 = 255
If G1 > 255 Then G1 = 255
If B1 > 255 Then B1 = 255
If R1 < 0 Then R1 = 0
If G1 < 0 Then G1 = 0
If B1 < 0 Then B1 = 0


Form1.ForeColor = RGB(R1, G1, B1)
Line (0, I)-(Form1.ScaleWidth, I)
R1 = R1 + RD
G1 = G1 + GD
B1 = B1 + BD
Next I

End Sub
