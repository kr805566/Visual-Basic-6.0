VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "�|�P�m"
   ClientHeight    =   2355
   ClientLeft      =   6570
   ClientTop       =   6255
   ClientWidth     =   9390
   BeginProperty Font 
      Name            =   "�s�ө���"
      Size            =   9
      Charset         =   136
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   2355
   ScaleWidth      =   9390
   Begin VB.CommandButton Command4 
      Caption         =   "���s�}�l"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6960
      TabIndex        =   13
      Top             =   1800
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "�}�l�U�`"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5040
      TabIndex        =   12
      Top             =   1800
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "-10"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8640
      TabIndex        =   9
      Top             =   1320
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "+10"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8640
      TabIndex        =   8
      Top             =   840
      Width           =   615
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   120
      Top             =   2040
   End
   Begin VB.Label Label7 
      Caption         =   "��"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8040
      TabIndex        =   11
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label Label6 
      Caption         =   "��"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8040
      TabIndex        =   10
      Top             =   360
      Width           =   375
   End
   Begin VB.Label Label5 
      Caption         =   "��`���B:"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5040
      TabIndex        =   7
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label4 
      Alignment       =   2  '�m�����
      BackColor       =   &H00FFFFFF&
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6480
      TabIndex        =   6
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label3 
      Alignment       =   2  '�m�����
      BackColor       =   &H00FFFFFF&
      Caption         =   "200"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6480
      TabIndex        =   5
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "�Ѿl���B:"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5040
      TabIndex        =   4
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  '�m�����
      BackColor       =   &H00C0C0FF&
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   72
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1575
      Index           =   3
      Left            =   3600
      TabIndex        =   3
      Top             =   360
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  '�m�����
      BackColor       =   &H00C0C0FF&
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   72
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1575
      Index           =   2
      Left            =   2520
      TabIndex        =   2
      Top             =   360
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  '�m�����
      BackColor       =   &H00C0C0FF&
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   72
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1575
      Index           =   1
      Left            =   1440
      TabIndex        =   1
      Top             =   360
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  '�m�����
      BackColor       =   &H00C0C0FF&
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   72
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1575
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim c(0 To 3), d, e, f As Integer
Private Sub Command1_Click()
If Val(Label4) + 10 > Val(Label3) Then MsgBox "���B����", , "�|�P�m": Exit Sub
If Val(Label4) = 50 Then MsgBox "��`�W����50��", , "�|�P�m": Exit Sub
Label4 = Val(Label4) + 10

End Sub
Private Sub Command2_Click()

If Val(Label4) = 10 Then MsgBox "��`�U����10��", , "�|�P�m": Exit Sub
Label4 = Val(Label4) - 10
End Sub

Private Sub Command3_Click()
If Val(Label3) < Val(Label4) Then MsgBox "���B����", , "�|�P�m": Label4 = Label3: Exit Sub
d = 0
e = 0
f = 0
For i = 0 To 3
A:
c(i) = InputBox("��J��" & i + 1 & "�ӼƦr", "�|�P�m")
If c(i) < 0 Or c(i) > 9 Then MsgBox "�п�J�Ӧ��", , "�|�P�m": GoTo A
Next i
Timer1.Enabled = True
Command1.Enabled = False
Command2.Enabled = False
Command3.Enabled = False
Command4.Enabled = False
End Sub

Private Sub Command4_Click()
Command1.Enabled = True
Command2.Enabled = True
Command3.Enabled = True
Label3 = 200

End Sub

Private Sub Timer1_Timer()
For i = 0 To 3
z = z & " " & c(i)
Next i
For i = 0 To 3

'Randomize Time


Label1(i) = Int(Rnd() * 10)


Next i

d = d + 1

If d = 30 Then
Timer1.Enabled = False

For i = 0 To 2
   For j = i + 1 To 3

   If Label1(i) > Label1(j) Then e = Label1(j): Label1(j) = Label1(i): Label1(i) = e
  Next j
Next i

For i = 0 To 2
   For j = i + 1 To 3

   If c(i) > c(j) Then e = c(j): c(j) = c(i): c(i) = e
  Next j
Next i

For i = 0 To 3
If c(i) = Label1(i) Then f = f + 1
Next i

If f = 4 Then
Label3 = Val(Label3) + Val(Label4) * 10
MsgBox "�AĹ�F" & Val(Label4) * 10 & "��" & ",�A��J���O " & z, , "�|�P�m"
Else
Label3 = Val(Label3) - Val(Label4)
MsgBox "�A��F" & Val(Label4) & "��" & ",�A��J���O " & z, , "�|�P�m"
End If
If Label3 = 0 Then
Command4.Enabled = True
Else
Command1.Enabled = True
Command2.Enabled = True
Command3.Enabled = True
Command4.Enabled = True
End If
End If



End Sub
