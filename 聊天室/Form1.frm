VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "��ѫ�"
   ClientHeight    =   4920
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7410
   LinkTopic       =   "Form1"
   ScaleHeight     =   4920
   ScaleWidth      =   7410
   StartUpPosition =   3  '�t�ιw�]��
   Begin VB.CommandButton Command4 
      Caption         =   "     ���}        ��ѫ�"
      Height          =   975
      Left            =   6120
      TabIndex        =   8
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "�M�����e"
      Height          =   855
      Left            =   6120
      TabIndex        =   7
      Top             =   1320
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "��J"
      Height          =   495
      Left            =   6120
      TabIndex        =   6
      Top             =   4080
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "��J"
      Height          =   495
      Left            =   6120
      TabIndex        =   5
      Top             =   360
      Width           =   975
   End
   Begin VB.TextBox Text3 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   1320
      MultiLine       =   -1  'True
      ScrollBars      =   3  '��̬Ҧ�
      TabIndex        =   4
      Top             =   960
      Width           =   4575
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   15.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      TabIndex        =   3
      Top             =   4080
      Width           =   4695
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   15.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      HideSelection   =   0   'False
      Left            =   1320
      TabIndex        =   1
      Top             =   360
      Width           =   4695
   End
   Begin VB.Label Label2 
      Caption         =   "�A��:"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   15.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   4200
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "�Ҥ�:"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   15.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

Text3 = Text3 & Label1 & Text1 & vbCrLf
Text3.FontSize = 12
Text1.SetFocus
Text1 = ""

End Sub

Private Sub Command2_Click()

Text3 = Text3 & Label2 & Text2 & vbCrLf
Text3.FontSize = 12
Text2.SetFocus
Text2 = ""

End Sub

Private Sub Command3_Click()

Text1 = ""
Text2 = ""
Text3 = ""

End Sub

Private Sub Command4_Click()
c = MsgBox("�T�w���}��ѫ�?", vbYesNo + vbQuestion, "ĵ�i")
If c = 6 Then

End

Else

MsgBox "�^���ѫ�", , "��^�{��"

End If

End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then Command1 = True
If KeyAscii = 13 Then KeyAscii = 0

End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then Command2 = True
If KeyAscii = 13 Then KeyAscii = 0

End Sub

Private Sub Text3_Change()
Text3.SelLength = Len(Text3)
End Sub
