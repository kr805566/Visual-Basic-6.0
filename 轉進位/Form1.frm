VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  '�t�ιw�]��
   Begin VB.CommandButton Command2 
      Caption         =   "8��10"
      Height          =   735
      Left            =   960
      TabIndex        =   1
      Top             =   1320
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "16��10"
      Height          =   735
      Left            =   960
      TabIndex        =   0
      Top             =   360
      Width           =   2655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

a = InputBox("�п�J16�i���", "��10�i��")
MsgBox "�o��16�i��Ʀr�ন10�i��ƬO " & Val("&h" & a), , "16��10"


End Sub

Private Sub Command2_Click()

a = InputBox("�п�J8�i���", "��10�i��")
MsgBox "�o��8�i��Ʀr�ন10�i��ƬO " & Val("&O" & a), , "8��10"

End Sub
