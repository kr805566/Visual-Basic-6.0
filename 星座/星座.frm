VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "�P�y�d��"
   ClientHeight    =   3960
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5085
   LinkTopic       =   "Form1"
   ScaleHeight     =   3960
   ScaleWidth      =   5085
   StartUpPosition =   3  '�t�ιw�]��
   Begin VB.CommandButton �M�� 
      Caption         =   "�M��"
      Height          =   495
      Left            =   2040
      TabIndex        =   20
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton ���� 
      Caption         =   "����"
      Height          =   495
      Left            =   3360
      TabIndex        =   19
      Top             =   720
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   3480
      TabIndex        =   18
      Top             =   240
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   2160
      TabIndex        =   17
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton �d�� 
      Caption         =   "�d��"
      Height          =   495
      Left            =   720
      TabIndex        =   13
      Top             =   720
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "�Q�G�P�y"
      Height          =   2415
      Left            =   480
      TabIndex        =   0
      Top             =   1440
      Width           =   4335
      Begin VB.OptionButton Option12 
         Caption         =   "�g��y"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2880
         TabIndex        =   12
         Top             =   1800
         Width           =   1095
      End
      Begin VB.OptionButton Option11 
         Caption         =   "���Ȯy"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2880
         TabIndex        =   11
         Top             =   1320
         Width           =   1095
      End
      Begin VB.OptionButton Option10 
         Caption         =   "�ѯ��y"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2880
         TabIndex        =   10
         Top             =   840
         Width           =   1095
      End
      Begin VB.OptionButton Option9 
         Caption         =   "�B�k�y"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2880
         TabIndex        =   9
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton Option8 
         Caption         =   "��l�y"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1560
         TabIndex        =   8
         Top             =   1800
         Width           =   1095
      End
      Begin VB.OptionButton Option7 
         Caption         =   "���ɮy"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1560
         TabIndex        =   7
         Top             =   1320
         Width           =   1095
      End
      Begin VB.OptionButton Option6 
         Caption         =   "���l�y"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1560
         TabIndex        =   6
         Top             =   840
         Width           =   1095
      End
      Begin VB.OptionButton Option5 
         Caption         =   "�����y"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1560
         TabIndex        =   5
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton Option4 
         Caption         =   "�d�Ϯy"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   1800
         Width           =   1095
      End
      Begin VB.OptionButton Option3 
         Caption         =   "�����y"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   1320
         Width           =   1095
      End
      Begin VB.OptionButton Option2 
         Caption         =   "���~�y"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   840
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "�]�~�y"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Label Label3 
      Caption         =   "��"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4560
      TabIndex        =   16
      Top             =   240
      Width           =   255
   End
   Begin VB.Label Label2 
      Caption         =   "��"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3240
      TabIndex        =   15
      Top             =   240
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "�п�J�X�ͤ��:"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   240
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub �d��_Click()
Dim m, d As Integer
m = Val(Text1)
d = Val(Text2)
Option1.FontBold = False
Option2.FontBold = False
Option3.FontBold = False
Option4.FontBold = False
Option5.FontBold = False
Option6.FontBold = False
Option7.FontBold = False
Option8.FontBold = False
Option9.FontBold = False
Option10.FontBold = False
Option11.FontBold = False
Option12.FontBold = False

If d <= 31 And d >= 1 Then
Select Case m

Case 1
    If d >= 1 And d <= 19 Then
    Option1.Value = True
    Option1.FontBold = True
    Else
    Option2.Value = True
    Option2.FontBold = True
    End If
Case 2
    If d >= 1 And d <= 29 Then
    If d >= 1 And d <= 19 Then
    Option2.Value = True
    Option2.FontBold = True
    Else
    Option3.Value = True
    Option3.FontBold = True
    End If
    Else
    GoTo a
    End If
Case 3
    If d >= 1 And d <= 20 Then
    Option3.Value = True
    Option3.FontBold = True
    Else
    Option4.Value = True
    Option4.FontBold = True
    End If
Case 4
    If d >= 1 And d <= 30 Then
    If d >= 1 And d <= 20 Then
    Option4.Value = True
    Option4.FontBold = True
    Else
    Option5.Value = True
    Option5.FontBold = True
    End If
    Else
    GoTo a
    End If
Case 5
    If d >= 1 And d <= 20 Then
    Option5.Value = True
    Option5.FontBold = True
    Else
    Option6.Value = True
    Option6.FontBold = True
    End If
Case 6
    If d >= 1 And d <= 30 Then
    If d >= 1 And d <= 21 Then
    Option6.Value = True
    Option6.FontBold = True
    Else
    Option7.Value = True
    Option7.FontBold = True
    End If
    Else
    GoTo a
    End If
Case 7
    If d >= 1 And d <= 22 Then
    Option7.Value = True
    Option7.FontBold = True
    Else
    Option8.Value = True
    Option8.FontBold = True
    End If
Case 8
    If d >= 1 And d <= 22 Then
    Option8.Value = True
    Option8.FontBold = True
    Else
    Option9.Value = True
    Option9.FontBold = True
    End If
Case 9
    If d >= 1 And d <= 30 Then
    If d >= 1 And d <= 22 Then
    Option9.Value = True
    Option9.FontBold = True
    Else
    Option10.Value = True
    Option10.FontBold = True
    End If
   Else
    GoTo a
    End If
Case 10
    If d >= 1 And d <= 22 Then
    Option10.Value = True
    Option10.FontBold = True
    Else
    Option11.Value = True
    Option11.FontBold = True
    End If
Case 11
    If d >= 1 And d <= 30 Then
    If d >= 1 And d <= 21 Then
    Option11.Value = True
    Option11.FontBold = True
    Else
    Option12.Value = True
    Option12.FontBold = True
    End If
    Else
    GoTo a
    End If
Case 12
    If d >= 1 And d <= 21 Then
    Option12.Value = True
    Option12.FontBold = True
    Else
    Option1.Value = True
    Option1.FontBold = True
    End If
Case Else
    MsgBox "������~", 16, "�P�y�d��"
    Text1 = ""
    Text1.SetFocus
End Select

Else
GoTo a
End If

Exit Sub

a: If m <= 12 And m >= 1 Then
   MsgBox m & "��S��" & d & "��", 16, "�P�y�d��"
   Text2 = ""
   Text2.SetFocus
   Else
   MsgBox "������~", 16, "�P�y�d��"
   Text1 = ""
   Text2 = ""
   Text1.SetFocus
   
   
   End If

End Sub
Private Sub �M��_Click()

Text1 = ""
Text2 = ""
Text1.SetFocus
Option1.FontBold = False
Option2.FontBold = False
Option3.FontBold = False
Option4.FontBold = False
Option5.FontBold = False
Option6.FontBold = False
Option7.FontBold = False
Option8.FontBold = False
Option9.FontBold = False
Option10.FontBold = False
Option11.FontBold = False
Option12.FontBold = False
Option1.Value = False
Option2.Value = False
Option3.Value = False
Option4.Value = False
Option5.Value = False
Option6.Value = False
Option7.Value = False
Option8.Value = False
Option9.Value = False
Option10.Value = False
Option11.Value = False
Option12.Value = False
End Sub
Private Sub ����_Click()

X = MsgBox("�O�_�n�����{��", 32 + vbYesNo, "���}")
If X = 6 Then End

End Sub

Private Sub Form_Activate()

Text1.SetFocus
Option1.Value = False
Option2.Value = False
Option3.Value = False
Option4.Value = False
Option5.Value = False
Option6.Value = False
Option7.Value = False
Option8.Value = False
Option9.Value = False
Option10.Value = False
Option11.Value = False
Option12.Value = False

End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If MsgBox("�O�_�n�����{��", 32 + vbYesNo, "���}") = vbNo Then Cancel = True
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then Text2.SetFocus

If KeyAscii = 13 Then KeyAscii = 0

End Sub


Private Sub Text2_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then �d��.SetFocus

If KeyAscii = 13 Then KeyAscii = 0

End Sub

