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

Dim D(1 To 1000000)

Open App.Path & "\1.TXT" For Input As #1





Do Until EOF(1)
Input #1, A, B

X = A
Y = B

If A > B Then
Temp = A
A = B
B = Temp
End If

For I = A To B
   N = I
   Do
   If N Mod 2 = 1 Then
   N = 3 * N + 1
   Else
   N = N / 2
   End If
   K = K + 1
   Loop Until N = 1
   D(I) = K + 1
   K = 0
Next I

For I = A To B
If D(I) > Max Then Max = D(I)
Next I


Print X; Y; Max

Max = 0

Loop
Close

End Sub
