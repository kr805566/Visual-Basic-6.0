VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6885
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   5850
   LinkTopic       =   "Form1"
   ScaleHeight     =   6885
   ScaleWidth      =   5850
   Begin VB.CommandButton Command2 
      Caption         =   "清除"
      Height          =   495
      Left            =   3720
      TabIndex        =   3
      Top             =   6120
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "輸入"
      Height          =   495
      Left            =   3720
      TabIndex        =   2
      Top             =   5520
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   2040
      TabIndex        =   0
      Top             =   5640
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  '平面
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5055
      Left            =   600
      ScaleHeight     =   5025
      ScaleWidth      =   5025
      TabIndex        =   1
      Top             =   240
      Width           =   5055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()


Dim a, b As Double


a = Val(Text1)
b = 6.28 / a

For i = 1 To a
   For j = i + 1 To a
Picture1.Line (Sin(b * i) * 2500 + 2500, Cos(b * i) * 2500 + 2500)-(Sin(b * j) * 2500 + 2500, Cos(b * j) * 2500 + 2500)

Next j, i


End Sub

Private Sub Command2_Click()

Picture1.Cls
 

End Sub

