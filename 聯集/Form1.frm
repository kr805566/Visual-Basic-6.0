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

Open App.Path & "\1.TXT" For Input As #1

Input #1, A, B

Do Until A = 0 And B = 0

  For I = A To B


  If I Mod 2 = 0 Or I Mod 3 = 0 Or I Mod 5 = 0 Then K = K + 1


  Next I
Print K
K = 0
Input #1, A, B
Loop
Close

End Sub

