VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  '系統預設值
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   480
      TabIndex        =   3
      Top             =   2040
      Width           =   3615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "+"
      Height          =   300
      Left            =   1680
      TabIndex        =   2
      Top             =   840
      Width           =   615
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   1200
      Width           =   3615
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   3615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()

Dim A(), B()

x = Text1
y = Text2

z = Len(Text1)

If Len(Text1) > Len(Text2) Then
y = String(Len(Text1) - Len(Text2), "0") & Text2
ElseIf Len(Text1) > Len(Text2) Then
z = Len(Text2)
x = String(Len(Text2) - Len(Text1), "0") & Text1
End If


ReDim A(z), B(z)



For I = 0 To z - 1
A(I) = Val(Mid(x, z - I, 1))
B(I) = Val(Mid(y, z - I, 1))
Next I

For I = 0 To z


N = A(I) + B(I) + W / 10
M = N Mod 10
W = N - M

C = C & M

Next I

g = z

If A(z - 1) + B(z - 1) > 9 Then g = g + 1


For I = g To 1 Step -1

d = d & Mid(C, I, 1)

Next I
Text3 = d


For I = 0 To z
A(I) = ""
B(I) = ""
Next I

C = ""
d = ""
N = ""
M = ""
W = ""
End Sub

