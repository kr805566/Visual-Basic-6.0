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

c = MsgBox("�����{����?", vbYesNo + vbQuestion, "ĵ�i")
If c = 6 Then
MsgBox "�Цs�ɦb���}", , "�����{��"

End

Else

MsgBox "�^��D�e��", , "��^�{��"

Shell ("C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE")

End If

End Sub
