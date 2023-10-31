VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8370
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10560
   LinkTopic       =   "Form1"
   ScaleHeight     =   8370
   ScaleWidth      =   10560
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   735
      Left            =   5160
      TabIndex        =   1
      Top             =   7320
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   735
      Left            =   2880
      TabIndex        =   0
      Top             =   7200
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

Form1.Cls
n = Val(Text1) + 2


ReDim a(n, n)

a(1, 1) = 1


For i = 2 To n

    For j = 1 To i - 1
    
    a(i, j) = a(i - 1, j) + a(i - 1, j - 1)

    Print a(i, j);

    Next j
    
Print "  N = " & i - 2

Next i






End Sub
