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

Dim i As Integer


word = InputBox("�ƤJ�r��", "����")

i = Len(word)

While i >= 1
drow = drow & Mid(word, i, 1)

i = i - 1
Wend
MsgBox word & "����ᬰ" & drow, , "�r�����"
End Sub

