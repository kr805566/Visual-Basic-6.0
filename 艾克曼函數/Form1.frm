VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  '系統預設值
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()

n = 8   ' n<9
m = 3  ' M<4

Print x(m, n)


End Sub

Function x(m, n)

If m = 0 Then
x = n + 1
ElseIf n = 0 Then
x = x(m - 1, 1)
Else
x = x(m - 1, x(m, n - 1))
End If
End Function
