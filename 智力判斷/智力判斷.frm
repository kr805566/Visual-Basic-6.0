VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "���O�P�_"
   ClientHeight    =   3225
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   3975
   LinkTopic       =   "Form1"
   ScaleHeight     =   3225
   ScaleWidth      =   3975
   StartUpPosition =   3  '�t�ιw�]��
   Begin VB.Frame Frame1 
      Caption         =   "���O����"
      BeginProperty Font 
         Name            =   "�رd�s���۸O"
         Size            =   12
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   1815
      Begin VB.OptionButton Option5 
         Caption         =   "IEP"
         BeginProperty Font 
            Name            =   "�رd�ײʶ�"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton Option4 
         Caption         =   "�Ѥ~"
         BeginProperty Font 
            Name            =   "�رd�ײʶ�"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   1800
         Width           =   1095
      End
      Begin VB.OptionButton Option3 
         Caption         =   "�����O"
         BeginProperty Font 
            Name            =   "�رd�ײʶ�"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   1440
         Width           =   1095
      End
      Begin VB.OptionButton Option2 
         Caption         =   "�����O"
         BeginProperty Font 
            Name            =   "�رd�ײʶ�"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   1080
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "���q���O"
         BeginProperty Font 
            Name            =   "�رd�ײʶ�"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   720
         Width           =   1095
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "��J"
      Height          =   495
      Left            =   2280
      TabIndex        =   1
      Top             =   1320
      Width           =   1530
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  '�m�����
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
      Left            =   2280
      TabIndex        =   0
      Top             =   600
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Option1.FontBold = False
Option2.FontBold = False
Option3.FontBold = False
Option4.FontBold = False
Option5.FontBold = False

Dim y As Integer
y = Val(Text1)
Select Case y
Case Is <= 79
MsgBox "�ڰJ�ߪ���ĳ�z!!�z�i��ݭn�ˬd�@�U�F!!", , "���O����"
Option5.FontBold = True
Option5.Value = True

Case 80 To 110
MsgBox "���q�H���O", , "���O����"
Option1.FontBold = True
Option1.Value = True
Case 111 To 120
MsgBox "�����O", , "���O����"
Option2.FontBold = True
Option2.Value = True
Case 121 To 140
MsgBox "�����O", , "���O����"
Option3.FontBold = True
Option3.Value = True
Case Else
MsgBox "�Ѥ~", , "���O����"
Option4.FontBold = True
Option4.Value = True

End Select

End Sub

