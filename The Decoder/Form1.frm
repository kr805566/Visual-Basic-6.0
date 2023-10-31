VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7485
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   7485
   StartUpPosition =   3  '系統預設值
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()

Open App.Path & "\1.txt" For Input As #1

Do Until EOF(1)
Line Input #1, A

For I = 1 To Len(A)

B = B & Chr(Asc(Mid(A, I, 1)) - 7)
Next I

Print B
B = ""

Loop


Close

End Sub
