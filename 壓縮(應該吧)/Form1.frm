VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3255
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5040
   LinkTopic       =   "Form1"
   ScaleHeight     =   3255
   ScaleWidth      =   5040
   StartUpPosition =   3  '系統預設值
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a()
Private Sub Form_Activate()


B = 0
z = 11110001
For I = 1 To Len(z)

If Mid(z, I, 1) = 1 Then
c = c + 1
End If
Next

ReDim a(1 To c + 1)


For I = 1 To Len(z)

 If Mid(z, I, 1) = 0 Then
 B = B + 1
 Else
 D = D + 1
 a(D) = B
 B = 0

 End If

Next
For I = 1 To D + 1

F = F & 十轉2(Val(a(I)))

Next
Print F

End Sub




Function 十轉2(X)


Do

Y = X \ 2
r = X Mod 2
A2 = A2 & r
X = Y

Loop While X >= 2

If X > 0 Then A2 = A2 & X

For I = Len(A2) To 1 Step -1
B2 = B2 & Mid(A2, I, 1)
Next I

十轉2 = B2

End Function











