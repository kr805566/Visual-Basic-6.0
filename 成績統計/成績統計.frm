VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5295
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5130
   LinkTopic       =   "Form1"
   ScaleHeight     =   5295
   ScaleWidth      =   5130
   StartUpPosition =   3  '�t�ιw�]��
   Begin VB.CommandButton Command3 
      Caption         =   "����"
      Height          =   375
      Left            =   3960
      TabIndex        =   8
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "���s��J"
      Height          =   375
      Left            =   3960
      TabIndex        =   7
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "��J���Z"
      Height          =   495
      Left            =   3960
      TabIndex        =   6
      Top             =   600
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Caption         =   "�p�⵲�G"
      Height          =   2655
      Left            =   120
      TabIndex        =   4
      Top             =   2520
      Width           =   4815
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2175
         Left            =   360
         MultiLine       =   -1  'True
         ScrollBars      =   1  '�������b
         TabIndex        =   5
         Top             =   240
         Width           =   4095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "�Х��Ŀﶵ��"
      Height          =   1935
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   3255
      Begin VB.CheckBox Check3 
         Caption         =   "�p��ή�P���ή���"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   1200
         Width           =   2535
      End
      Begin VB.CheckBox Check2 
         Caption         =   "��X�̰����P�̧C��"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   2535
      End
      Begin VB.CheckBox Check1 
         Caption         =   "�p���`���P����"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2535
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As Integer
Private Sub Check1_Click()

If Check1 = 1 Or Check2 = 1 Or Check3 = 1 Then
Command1.Enabled = True
Else
Command1.Enabled = False
End If
End Sub
Private Sub Check2_Click()

If Check1 = 1 Or Check2 = 1 Or Check3 = 1 Then
Command1.Enabled = True
Else
Command1.Enabled = False
End If
End Sub
Private Sub Check3_Click()

If Check1 = 1 Or Check2 = 1 Or Check3 = 1 Then
Command1.Enabled = True
Else
Command1.Enabled = False
End If
End Sub
Private Sub Command1_Click()
Dim su, max, m, sco, n As Integer

sun = 0
max = 0
m = 100
n = 0
Text1 = "�q�Ҧ��Z  "

For i = 1 To a
sco = Val(InputBox("��J��" & i & "�q�Ҧ��Z", "���Z�έp"))
su = su + sco
Text1 = Text1 & sco & "��" & Space(2)
If sco > max Then max = sco
If sco < m Then m = sco
If sco >= 60 Then n = n + 1
Next i


If Check1.Value = 1 Then Text1 = Text1 & vbCrLf & a & "���`���� " & su & "��" & vbCrLf & a & "�쥭���� " & Format(su / a, "0.0") & "��"
If Check2.Value = 1 Then Text1 = Text1 & vbCrLf & "�̰����� " & max & "��" & vbCrLf & "�̧C���� " & m & "��"
If Check3.Value = 1 Then Text1 = Text1 & vbCrLf & "�ή��Ƭ� " & n & "��" & vbCrLf & "���ή��Ƭ� " & a - n & "��"
End Sub

Private Sub Command2_Click()
Text1 = ""
Check1.Value = 0
Check2.Value = 0
Check3.Value = 0
a = InputBox("��J���", "���Z�έp")
Command1.Enabled = False


End Sub



Private Sub Command3_Click()
End

End Sub

Private Sub Form_Activate()

a = InputBox("��J���", "���Z�έp")
Command1.Enabled = False


End Sub

