VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "�����T�{"
   ClientHeight    =   2670
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4905
   BeginProperty Font 
      Name            =   "�s�ө���"
      Size            =   18
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   2670
   ScaleWidth      =   4905
   StartUpPosition =   3  '�t�ιw�]��
   Begin VB.CommandButton ���� 
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      TabIndex        =   6
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton ���s��J 
      Caption         =   "���s��J"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   5
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton �n�J 
      Caption         =   "�n�J"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   4
      Top             =   2040
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   495
      IMEMode         =   3  '�Ȥ�
      Left            =   1440
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1320
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   1440
      TabIndex        =   0
      Top             =   600
      Width           =   3135
   End
   Begin VB.Label Label2 
      Alignment       =   2  '�m�����
      Caption         =   "�K�X"
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
      Left            =   360
      TabIndex        =   2
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  '�m�����
      Caption         =   "�b��"
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
      Left            =   360
      TabIndex        =   1
      Top             =   600
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub �n�J_Click()

A = Text1
b = Text2
Select Case A
 Case Is = "123456"
  Select Case b
   Case Is = "A654321"
   MsgBox "�w��n�J", , "�^���p��"
   Case Else
   MsgBox "�K�X���~!", vbCritical, "�����T�{"
   
   End Select
 Case Else
  MsgBox "�L���b��!", vbCritical, "�����T�{"
 
End Select
End Sub

Private Sub ���s��J_Click()
Text1 = ""
Text2 = ""
End Sub

Private Sub ����_Click()

End

End Sub
Private Sub Text1_LostFocus()
A = Text1
Select Case A
 Case Is = "123456"
Case Else
  MsgBox "�L���b��!", vbCritical, "�����T�{"
  Text1 = ""
  Text1.SetFocus
  
End Select
End Sub
