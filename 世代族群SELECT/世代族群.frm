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
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()

Form1.Hide
Dim y As Integer
y = InputBox("�п�J�X�ͦ~��", "��J�~��")
Select Case y
Case Is <= 59
MsgBox "�ݼֱ�", , "�@�N�ڸs"
Case 60 To 69
MsgBox "�����", , "�@�N�ڸs"
Case Else
MsgBox "���e���", , "�@�N�ڸs"
End Select

End






End Sub
