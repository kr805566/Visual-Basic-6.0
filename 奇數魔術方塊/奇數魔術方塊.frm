VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "奇數魔術方塊"
   ClientHeight    =   4665
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6285
   LinkTopic       =   "Form1"
   ScaleHeight     =   4665
   ScaleWidth      =   6285
   StartUpPosition =   3  '系統預設值
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a()

Private Sub Form_Activate()
n = 5
ReDim a(1 To n, 1 To n)
Randomize
b = Int(Rnd() * n + 1)
c = Int(Rnd() * n + 1)
a(b, c) = 1
k = 1
For i = 2 To n * n

If k = n Then
  If b = n Then
  b = 1
  Else
  b = b + 1
  End If
  k = 1
Else
  If b = 1 Then
  b = n
  Else
  b = b - 1
  End If

  If c = n Then
  c = 1
  Else
  c = c + 1
  End If
  
  k = k + 1
End If
a(b, c) = i
Next i






For i = 1 To n
   For j = 1 To n
Print a(i, j);
Next j
Print
Next i
End Sub

