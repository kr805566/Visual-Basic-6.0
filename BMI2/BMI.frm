VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "BMI�� �p��"
   ClientHeight    =   3675
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6825
   LinkTopic       =   "Form1"
   ScaleHeight     =   3675
   ScaleWidth      =   6825
   StartUpPosition =   3  '�t�ιw�]��
   Begin VB.CommandButton Command3 
      Caption         =   "���}"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   15.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5400
      TabIndex        =   14
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "�M��"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   15.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4200
      TabIndex        =   13
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�p��"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   15.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3000
      TabIndex        =   10
      Top             =   2520
      Width           =   975
   End
   Begin VB.Frame Frame3 
      Caption         =   "��J���� �魫"
      Height          =   1695
      Left            =   3000
      TabIndex        =   5
      Top             =   240
      Width           =   3495
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   840
         TabIndex        =   7
         Top             =   960
         Width           =   1935
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   840
         TabIndex        =   6
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label4 
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2880
         TabIndex        =   12
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2880
         TabIndex        =   11
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "�魫:"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   9
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "����:"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "��ܩʧO"
      Height          =   1575
      Left            =   360
      TabIndex        =   2
      Top             =   1800
      Width           =   1575
      Begin VB.OptionButton Option2 
         Caption         =   "�k"
         Height          =   495
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Caption         =   "�k"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "��J�m�W"
      Height          =   1215
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   2055
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
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1575
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bmi As Single
Private Sub Command1_Click()
If Text1 <> "" And Val(Text2) > 0 And Val(Text3) > 0 And (Option1.Value = True Or Option2.Value = True) Then

   bmi = Format(Val(Text3) / (Val(Text2) / 100) ^ 2, "##.#")
   If Option1.Value = True Then
      
         If bmi <= 27.8 Then
         MsgBox "�z��BMI�Ȭ�" & bmi & "�A�ݩ�@�먭��", , Text1 & "����"
         Else
         MsgBox "�z��BMI�Ȭ�" & bmi & "�A�ݩ�έD����", , Text1 & "����"
         End If
   End If

      
   If Option2.Value = True Then
         
         If bmi <= 27.3 Then
         MsgBox "�z��BMI�Ȭ�" & bmi & "�A�ݩ�@�먭��", , Text1 & "�k�h"
         Else
         MsgBox "�z��BMI�Ȭ�" & bmi & "�A�ݩ�έD����", , Text1 & "�k�h"
         End If
  End If
  
  A = MsgBox("�~��ϥΥ��{����?", 32 + 4, "���}")
  If A = 7 Then End
  If A = 6 Then B = MsgBox("���s��J��?", 32 + 4, "���s��J")
  If B = 6 Then Command2 = True
  
Else
 If Text1 = "" Then
 MsgBox "�п�J�m�W", 16, "���~"
 Text1.SetFocus
 Else
    If Text2 = "" Then
    MsgBox "�п�J����", 16, "���~"
    Text2.SetFocus
    Else
       If Val(Text2) < 0 Then
       MsgBox "�����S���t��", 16, "���~"
       Text2 = ""
       Text2.SetFocus
       Else
          If Text3 = "" Then
          MsgBox "�п�J�魫", 16, "���~"
          Text3.SetFocus
          Else
            If Val(Text3) < 0 Then
            MsgBox "�魫�S���t��", 16, "���~"
            Text3 = ""
            Text3.SetFocus
            Else
            If Option1.Value = False And Option2.Value = False Then MsgBox "�п�ܩʧO", 16, "���~"
            Option1.SetFocus
            End If
          End If
        End If
    End If
 End If
End If
End Sub

Private Sub Command2_Click()
Text1.SetFocus
Text1 = ""
Text2 = ""
Text3 = ""
Option1 = True
Option2 = False
End Sub

Private Sub Command3_Click()

End

End Sub

Private Sub Form_Activate()

Text1.SetFocus

End Sub


Private Sub Text1_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then Text2.SetFocus
If KeyAscii = 13 Then KeyAscii = 0

End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then Text3.SetFocus
If KeyAscii = 13 Then KeyAscii = 0

End Sub
Private Sub Text3_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then Option1.SetFocus
If KeyAscii = 13 Then KeyAscii = 0

End Sub
Private Sub Option1_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then Command1.SetFocus
If KeyAscii = 13 Then KeyAscii = 0

End Sub
Private Sub Option2_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then Command1.SetFocus
If KeyAscii = 13 Then KeyAscii = 0

End Sub
