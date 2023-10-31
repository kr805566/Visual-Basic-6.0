VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "瞦计"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   Icon            =   "瞦计.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  '╰参箇砞
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

b = Val(InputBox("叫块 99 < N < 1000 俱计", "瞦计"))

c = 0
x = 99
y = 1000





Do While A <> b

c = c + 1

If A > b Then

x = b

b = Val(InputBox("计び,絛瞅" & x & " < N < " & y & "," & "瞦" & c & "Ω", "瞦计"))

Else

y = b

b = Val(InputBox("计び,絛瞅" & x & " < N < " & y & "," & "瞦" & c & "Ω", "瞦计"))

End If

Loop

MsgBox "尺!夹非氮琌" & A & ",眤瞦" & c + 1 & "Ω", , "瞦计"

End

End Sub



