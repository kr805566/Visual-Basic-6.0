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
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   1320
      Width           =   4215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   255
      Left            =   600
      TabIndex        =   2
      Top             =   2280
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Text            =   "22222"
      Top             =   720
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Text            =   "22222"
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim a, b, c
Private Sub Command1_Click()
n = Len(Text1)
m = Len(Text2)
ReDim a(n), b(m), c(n + m)

For i = 1 To n
    a(i) = Mid(Text1, n - i + 1, 1)
Next i

For i = 1 To m
    b(i) = Mid(Text2, m - i + 1, 1)
Next i

For i = 1 To m
    For j = 1 To n

    c(j + i - 1) = c(j + i - 1) + a(j) * b(i)

    Next j
Next i

For i = 1 To m + n - 1
 c(i + 1) = c(i + 1) + c(i) \ 10
c(i) = c(i) Mod 10
Next i

Text3 = ""

For i = 1 To m + n
If i = m + n And c(m + n) = 0 Then Exit For
Text3 = c(i) & Text3

Next i
End Sub

