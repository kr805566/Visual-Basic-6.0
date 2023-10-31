VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   10020
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   10020
   ScaleWidth      =   4560
   StartUpPosition =   3  '系統預設值
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
Call hh(4, "A", "B", "C")
End Sub

Function hh(n, s, t, e)

If n > 0 Then

hh = hh(n - 1, s, e, t)
Print s & " 移動到 " & e & "  " & n
hh = hh(n - 1, t, s, e)

End If



End Function
