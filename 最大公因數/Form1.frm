VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3105
   ClientLeft      =   60
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3105
   ScaleWidth      =   4680
   StartUpPosition =   3  '�t�ιw�]��
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
��1 = Val(InputBox("�п�Ja"))

��2 = Val(InputBox("�п�Jb"))
Call ��(��1, ��2)

End Sub
Function ��(a, b)
If b > a Then t = a: a = b: b = t
If a Mod b = 0 Then
MsgBox b
Else
Call ��(b, a Mod b)
End If




End Function
