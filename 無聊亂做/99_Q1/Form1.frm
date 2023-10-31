VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4935
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6780
   LinkTopic       =   "Form1"
   ScaleHeight     =   4935
   ScaleWidth      =   6780
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command3 
      Caption         =   "exit"
      Height          =   375
      Left            =   5160
      TabIndex        =   5
      Top             =   2640
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Random Set"
      Height          =   495
      Left            =   5160
      TabIndex        =   4
      Top             =   360
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Encoding"
      Height          =   495
      Left            =   5160
      TabIndex        =   3
      Top             =   1320
      Width           =   975
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   2520
      Width           =   4695
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   4695
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   4695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a(), s(39)

Private Sub Command1_Click()



b = 0
e = 0

z = Text1

For i = 1 To Len(z)

If Mid(z, i, 1) = 1 Then
c = c + 1
End If
Next

ReDim a(1 To c + 1)


For i = 1 To Len(z)

 If Mid(z, i, 1) = 0 Then
 b = b + 1
 Else
 d = d + 1
 a(d) = b
 b = 0

 End If

Next


If Mid(z, 40, 1) = 0 Then
d = d + 1
a(d) = b

End If


For i = 1 To d



f = f & 十轉2(Val(a(i))) & " "

Next
Text2 = f

Text3 = (Len(f) - 1) / Len(z) * 100 & "%"
End Sub

Function 十轉2(X)



Do

Y = X \ 2
r = X Mod 2
A2 = A2 & r
X = Y

Loop While X >= 2

If X > 0 Then A2 = A2 & X

For i = Len(A2) To 1 Step -1
B2 = B2 & Mid(A2, i, 1)
Next i

十轉2 = B2

End Function

Private Sub Command2_Click()


For i = 0 To 39
s(i) = 0

Next i
Text1 = ""

Randomize Time
For i = 1 To 4
c = Int((Rnd * 40))

If s(c) = 1 Then
i = i - 1
Else
s(c) = 1
End If


Next i
For i = 0 To 39

Text1 = Text1 & s(i)

Next i

End Sub

Private Sub Command3_Click()
End
End Sub




