VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3105
   ClientLeft      =   60
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3105
   ScaleWidth      =   4680
   StartUpPosition =   3  '�t�ιw�]��
   Begin VB.CommandButton Command1 
      Caption         =   "����"
      Height          =   495
      Left            =   3360
      TabIndex        =   1
      Top             =   600
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   2655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
b = ""
a = Text1

For i = Len(a) To 1 Step -1

b = b & Mid(a, i, 1)


Next i

If a = b Then

MsgBox "�O�j��", , "�P�_"

Else
MsgBox "���O�j��", , "�P�_"

End If


End Sub
