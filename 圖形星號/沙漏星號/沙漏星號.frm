VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "�F�|�P��"
   ClientHeight    =   9315
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4605
   LinkTopic       =   "Form1"
   ScaleHeight     =   9315
   ScaleWidth      =   4605
   StartUpPosition =   3  '�t�ιw�]��
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
k = InputBox("��J�Ӽ�", "�F�|")

If k Mod 2 = 1 Then

 q = (k - 3) / 2

 Print String(k, "*")

 For i = q To 1 Step -1
 Print Spc(q + 1 - i); String(1, "*"); Spc(2 * i - 1); String(1, "*")
 Next i

 Print Spc(q + 1); String(1, "*")

 For i = 1 To q
 Print Spc(q + 1 - i); String(1, "*"); Spc(2 * i - 1); String(1, "*")
 Next i

 Print String(k, "*")

Else

 q = (k - 2) / 2

Print String(k, "*")

For i = q To 2 Step -1
Print Spc(q + 1 - i); String(1, "*"); Spc(2 * i - 2); String(1, "*")
Next i

Print Spc(q); String(2, "*")

For i = 2 To q
Print Spc(q + 1 - i); String(1, "*"); Spc(2 * i - 2); String(1, "*")
Next i

Print String(k, "*")


End If
End Sub

