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
Dim A(1 To 3)

Open App.Path & "\1.txt" For Input As #1


Input #1, A(1), A(2), A(3)


Do
     
   For I = 1 To 2
       For J = I To 3
       
       If A(I) > A(J) Then
       Temp = A(J)
       A(J) = A(I)
       A(I) = Temp
       End If
       
   Next J, I

Print A(1); A(2); A(3)

Input #1, A(1), A(2), A(3)

Loop Until A(1) = 0 And A(2) = 0 And A(3) = 0

Close





End Sub




