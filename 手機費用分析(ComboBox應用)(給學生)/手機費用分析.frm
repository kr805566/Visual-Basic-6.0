VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "101�~�W�b�~�פ���O�θ��"
   ClientHeight    =   7530
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10785
   Icon            =   "����O�Τ��R.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7530
   ScaleWidth      =   10785
   StartUpPosition =   3  '�t�ιw�]��
   Begin VB.CommandButton Command2 
      Caption         =   "�M��"
      Height          =   495
      Left            =   6120
      TabIndex        =   21
      Top             =   960
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "����B���B"
      Height          =   1575
      Left            =   480
      TabIndex        =   15
      Top             =   240
      Width           =   5175
      Begin VB.CommandButton Command1 
         Caption         =   "�x�s"
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
         Left            =   4080
         TabIndex        =   20
         Top             =   480
         Width           =   975
      End
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2040
         TabIndex        =   17
         Text            =   "�п��"
         Top             =   600
         Width           =   1215
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         TabIndex        =   16
         Text            =   "�п��"
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   2  '�m�����
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
         Left            =   3360
         TabIndex        =   19
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   2  '�m�����
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
         Left            =   1320
         TabIndex        =   18
         Top             =   600
         Width           =   615
      End
   End
   Begin VB.CommandButton ���� 
      Caption         =   "����"
      Height          =   495
      Left            =   6120
      TabIndex        =   12
      Top             =   1560
      Width           =   1335
   End
   Begin VB.CommandButton ��Ƥ��R 
      Caption         =   "��Ƥ��R"
      Height          =   495
      Left            =   6120
      TabIndex        =   11
      Top             =   360
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Caption         =   "�έp���G"
      Height          =   4815
      Left            =   120
      TabIndex        =   0
      Top             =   2160
      Width           =   7335
      Begin VB.ListBox List2 
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2460
         ItemData        =   "����O�Τ��R.frx":058A
         Left            =   2640
         List            =   "����O�Τ��R.frx":058C
         TabIndex        =   4
         Top             =   960
         Width           =   2295
      End
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2460
         ItemData        =   "����O�Τ��R.frx":058E
         Left            =   120
         List            =   "����O�Τ��R.frx":0590
         TabIndex        =   1
         Top             =   960
         Width           =   2295
      End
      Begin VB.Label Label12 
         Alignment       =   2  '�m�����
         BackColor       =   &H00C0FFC0&
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
         TabIndex        =   14
         Top             =   2880
         Width           =   1455
      End
      Begin VB.Label Label11 
         Alignment       =   2  '�m�����
         AutoSize        =   -1  'True
         Caption         =   "��ú���B"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   14.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5160
         TabIndex        =   13
         Top             =   2400
         Width           =   1155
      End
      Begin VB.Label Label10 
         Alignment       =   2  '�m�����
         BackColor       =   &H0080C0FF&
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
         TabIndex        =   10
         Top             =   3960
         Width           =   1455
      End
      Begin VB.Label Label9 
         Alignment       =   2  '�m�����
         AutoSize        =   -1  'True
         Caption         =   "�������B"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   14.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5160
         TabIndex        =   9
         Top             =   3600
         Width           =   1155
      End
      Begin VB.Label Label8 
         Alignment       =   2  '�m�����
         BackColor       =   &H00FFC0FF&
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
         TabIndex        =   8
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label Label7 
         Alignment       =   2  '�m�����
         BackColor       =   &H00C0FFFF&
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
         TabIndex        =   7
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label6 
         Alignment       =   2  '�m�����
         AutoSize        =   -1  'True
         Caption         =   "�̧C���B"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   14.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5160
         TabIndex        =   6
         Top             =   1440
         Width           =   1155
      End
      Begin VB.Label Label5 
         Alignment       =   2  '�m�����
         AutoSize        =   -1  'True
         Caption         =   "�̰����B"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   14.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5145
         TabIndex        =   5
         Top             =   360
         Width           =   1155
      End
      Begin VB.Label Label4 
         Alignment       =   2  '�m�����
         AutoSize        =   -1  'True
         Caption         =   "�Ƨǫ�"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   14.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3480
         TabIndex        =   3
         Top             =   360
         Width           =   885
      End
      Begin VB.Label Label3 
         Alignment       =   2  '�m�����
         AutoSize        =   -1  'True
         Caption         =   "�Ƨǫe"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   14.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   840
         TabIndex        =   2
         Top             =   360
         Width           =   885
      End
   End
   Begin VB.Image Image1 
      Height          =   1695
      Left            =   8280
      Picture         =   "����O�Τ��R.frx":0592
      Stretch         =   -1  'True
      Top             =   4200
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a(1 To 6) As String
Dim b(1 To 6) As Integer



Private Sub Command1_Click()
If a(Val(Combo1)) = "" Then

a(Val(Combo1)) = Combo1
b(Val(Combo1)) = Combo2

List1.AddItem a(Val(Combo1)) & "�����O�ά�" & b(Val(Combo1)) & "��"


End If

Combo1 = "�п��"
Combo2 = "�п��"

End Sub

Private Sub Command2_Click()

List1.Clear
List2.Clear
Label7 = ""
Label8 = ""
Label12 = ""
Label10 = ""

For i = 1 To 6

a(i) = ""
 
Next i

End Sub

Private Sub Form_Activate()

For i = 1 To 6

Combo1.AddItem i

Next i
For i = 300 To 700 Step 50

Combo2.AddItem i

Next i



End Sub




Private Sub Image1_Click()

MsgBox "�ڬO��@1�L�F�o"
End Sub

Private Sub ����_Click()
End
End Sub

Private Sub ��Ƥ��R_Click()

Max = 300
Min = 700

For i = 1 To 6
If Max < b(i) Then Max = b(i)
If Min > b(i) Then Min = b(i)
s = s + b(i)
Next i
Label7 = Max
Label8 = Min
Label12 = s
Label10 = Format(s / 6, "0.0")

For i = 1 To 5
  For j = i + 1 To 6

  If b(i) < b(j) Then
  
  D = a(i): a(i) = a(j): a(j) = D
  e = b(i): b(i) = b(j): b(j) = e
  End If

Next j, i

For i = 1 To 6

List2.AddItem a(i) & "�����O�ά�" & b(i) & "��"

Next i


End Sub

