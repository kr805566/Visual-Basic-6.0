VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "�q�T��"
   ClientHeight    =   4080
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7920
   LinkTopic       =   "Form2"
   ScaleHeight     =   4080
   ScaleWidth      =   7920
   StartUpPosition =   3  '�t�ιw�]��
   Begin VB.CommandButton Command1 
      Caption         =   "����"
      Height          =   495
      Left            =   6480
      TabIndex        =   0
      Top             =   3240
      Width           =   855
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Open "�꣸1.txt" For Append As #1
Form2.FontSize = 16

Form1.Show
Form2.Hide

End Sub

