VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2625
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5070
   LinkTopic       =   "Form1"
   ScaleHeight     =   2625
   ScaleWidth      =   5070
   StartUpPosition =   3  '�t�ιw�]��
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   1320
      Width           =   4455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "��J"
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   2160
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Text            =   "70"
      Top             =   720
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Text            =   "2"
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label Label4 
      Caption         =   "m= 1~99"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   24
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      TabIndex        =   7
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "n= 1~9"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   24
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      TabIndex        =   6
      Top             =   0
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "m"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   27.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2160
      TabIndex        =   5
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "n"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   48
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1800
      TabIndex        =   4
      Top             =   1680
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim A, B, C  'A2A1 * B2B1 = C4C3C2C1
Private Sub Command1_Click()
 Call ������(Text1, Val(Text2), 1) 'text1����  text2����  1��l��
End Sub

Function ������(����, ����, ���G)
If ���� = 0 Then Text3 = ���G: Exit Function
Call ������(����, ���� - 1, �ۭ�(����, ���G)) '���Ƴv����1 ��0�ɥN��w��M���F

End Function

Function �ۭ�(A��, B��)
A�Ʀ�� = Len(A��) '��X�Ʀr�X���
B�Ʀ�� = Len(B��)
ReDim A(A�Ʀ��), B(B�Ʀ��), C(A�Ʀ�� + B�Ʀ��) 'A B �ۭ�  C����Ƴ̦h����A�����+B�����

For I = 1 To A�Ʀ��
    A(I) = Mid(A��, A�Ʀ�� - I + 1, 1) ' �C�@�쪺�Ʀr��i�h �Ҧp:314 A(3)=3 A(2)=1 A(1)=4
Next I

For I = 1 To B�Ʀ��
    B(I) = Mid(B��, B�Ʀ�� - I + 1, 1)
Next I

For I = 1 To A�Ʀ��
    For j = 1 To B�Ʀ��                       '   A2   A1           �Ҧp   1   2
                                               '   B2   B1                  2   9
    C(j + I - 1) = C(j + I - 1) + A(I) * B(j)  'X--------                X------
                                               '   C2(1) C1                 9  18
    Next j                                     'C3 C2(2)                 2  4
Next I                                         '----------              ------------
                                               'C3 C2    C1              2 13  18
For I = 1 To A�Ʀ�� + B�Ʀ�� - 1             '                        ------------
 C(I + 1) = C(I + 1) + C(I) \ 10               '�b�P�_�O�_�W�L9          2 13  8
C(I) = C(I) Mod 10                             '�b�i��                     +1
Next I                                         '                        ------------
                                               '                         2 14 8
                                               '                        ------------
                                               '                         2  4  8
                                               '                        +1
                                               '                        -------------
                                               '                          3  4  8
                                               
                                               
If C(A�Ʀ�� + B�Ʀ��) = 0 Then C(A�Ʀ�� + B�Ʀ��) = ""   ' �o�ӧP�_�O���F����e���h�X�{0
                                                             '�Ҧp: 11 * 12 = 0132

For I = 1 To A�Ʀ�� + B�Ʀ��

�ۭ� = C(I) & �ۭ�                                            '�N��զ^�h
                                                              '�Ҧp C(3)=2  C(2)=5  C(1)=6   �ۭ�=256

Next I




End Function


