VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "�q�Ʀr"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   Icon            =   "�q�Ʀr.frx":0000
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

Dim A, b, c, x, y As Integer

Randomize Time


A = Int(Rnd() * 900) + 100

b = Val(InputBox("�п�J 99 < N < 1000 �����", "�q�Ʀr"))

c = 0
x = 99
y = 1000





Do While A <> b

c = c + 1

If A > b Then

x = b

b = Val(InputBox("�Ʀr�Ӥp�F,�d��" & x & " < N < " & y & "," & "�q�F" & c & "��", "�q�Ʀr"))

Else

y = b

b = Val(InputBox("�Ʀr�Ӥj�F,�d��" & x & " < N < " & y & "," & "�q�F" & c & "��", "�q�Ʀr"))

End If

Loop

MsgBox "����!�зǵ��׬O" & A & ",�z�@�q�F" & c + 1 & "��", , "�q�Ʀr"

End

End Sub



