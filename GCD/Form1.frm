VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4620
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7650
   LinkTopic       =   "Form1"
   ScaleHeight     =   4620
   ScaleWidth      =   7650
   StartUpPosition =   3  '系統預設值
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()

Open App.Path & "\1.txt" For Input As #1
Open App.Path & "\2.txt" For Output As #2



Input #1, A, B


Do
  
  For I = A To 1

  If A Mod I = 0 And B Mod I = 0 Then Exit For

  Next I

  Ans = Ans & "GCD(" & A & "," & B & ")=" & I & vbCrLf
  
  Input #1, A, B

Loop Until A = 0 And B = 0

Print #2, Ans

Close

End

End Sub
