VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "��i��"
   ClientHeight    =   4905
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   4905
   ScaleWidth      =   6585
   StartUpPosition =   3  '�t�ιw�]��
   Begin VB.CommandButton Command1 
      Caption         =   "�Q���i����"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   3
      Left            =   4920
      TabIndex        =   5
      Top             =   1200
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�Q�i����"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   2
      Left            =   3360
      TabIndex        =   4
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�K�i����"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   1800
      TabIndex        =   3
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�G�i����"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   1920
      TabIndex        =   0
      Top             =   360
      Width           =   2175
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   360
      TabIndex        =   2
      Top             =   2160
      Width           =   5775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public a2, a8, a10, a16, b2, b8, b10, b16
Dim a, c As Integer




Private Sub Command1_Click(Index As Integer)
z2 = "�G�i�쬰 "
z8 = "�K�i�쬰 "
z10 = "�Q�i�쬰 "
z16 = "�Q���i�쬰 "

Label1 = ""
a16 = ""
a10 = ""
a8 = ""
a2 = ""
b16 = ""
b10 = ""
b8 = ""
b2 = ""

Select Case Index

Case 0

Call �G��10
c = Val(b10)
Call �Q��8
Call �Q��16
Label1 = z10 & b10 & vbCrLf & z8 & b8 & vbCrLf & z16 & b16
Case 1
Call �K��10
c = Val(b10)
Call �Q��2
Call �Q��16
Label1 = z2 & b2 & vbCrLf & z10 & b10 & vbCrLf & z16 & b16
Case 2
c = Val(Text1)
Call �Q��2
Call �Q��8
Call �Q��16
Label1 = z2 & b2 & vbCrLf & z8 & b8 & vbCrLf & z16 & b16
Case 3
Call �Q����10
c = Val(b10)
Call �Q��2
Call �Q��8
Label1 = z2 & b2 & vbCrLf & z10 & b10 & vbCrLf & z8 & b8

End Select
End Sub

Sub �Q��2()

a = c
Do

b = a \ 2
r = a Mod 2
a2 = a2 & r
a = b

Loop While a >= 2

If a > 0 Then a2 = a2 & a

For i = Len(a2) To 1 Step -1
b2 = b2 & Mid(a2, i, 1)
Next i


End Sub
Sub �Q��8()

a = c
Do

b = a \ 8
r = a Mod 8
a8 = a8 & r
a = b

Loop While a >= 8

If a > 0 Then a8 = a8 & a

For i = Len(a8) To 1 Step -1
b8 = b8 & Mid(a8, i, 1)
Next i



End Sub

Sub �Q��16()

a = c
Do

b = a \ 16
r = a Mod 16
If r > 9 Then r = Chr(55 + r)

a16 = a16 & r
a = b

Loop While a >= 16
If a > 9 Then a = Chr(55 + a)
If a > 0 Then a16 = a16 & a

For i = Len(a16) To 1 Step -1
b16 = b16 & Mid(a16, i, 1)
Next i

End Sub
Sub �G��10()
a = Val(Text1)
b = Len(Text1)
For i = b To 1 Step -1
b10 = Val(b10) + 2 ^ (b - i) * Val(Mid(a, i, 1))
Next
End Sub
Sub �K��10()
a = Val(Text1)
b = Len(Text1)
For i = b To 1 Step -1
b10 = Val(b10) + 8 ^ (b - i) * Val(Mid(a, i, 1))
Next
End Sub

Sub �Q����10()
a = Text1
b = Len(Text1)
For i = b To 1 Step -1
d = Mid(a, i, 1)
Select Case d


Case "A"
d = 10
Case "B"
d = 11
Case "C"
d = 12
Case "D"
d = 13
Case "E"
d = 14
Case "F"
d = 15
End Select

b10 = Val(b10) + 16 ^ (b - i) * d
Next i

End Sub
