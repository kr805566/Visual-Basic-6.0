VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7530
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   15375
   LinkTopic       =   "Form1"
   ScaleHeight     =   7530
   ScaleWidth      =   15375
   StartUpPosition =   3  '�t�ιw�]��
   Begin VB.Frame Frame2 
      Caption         =   "B�ȤH"
      Height          =   6135
      Left            =   7440
      TabIndex        =   1
      Top             =   240
      Width           =   7095
      Begin VB.Image Image4 
         Height          =   2535
         Index           =   1
         Left            =   3720
         Stretch         =   -1  'True
         Top             =   3240
         Width           =   3015
      End
      Begin VB.Image Image3 
         Height          =   2535
         Index           =   1
         Left            =   240
         Stretch         =   -1  'True
         Top             =   3240
         Width           =   3015
      End
      Begin VB.Image Image2 
         Height          =   2655
         Index           =   1
         Left            =   3720
         Stretch         =   -1  'True
         Top             =   360
         Width           =   3015
      End
      Begin VB.Image Image1 
         Height          =   2655
         Index           =   1
         Left            =   240
         Stretch         =   -1  'True
         Top             =   360
         Width           =   3015
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "A�ȤH"
      Height          =   6135
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   7095
      Begin VB.Image Image4 
         Height          =   2535
         Index           =   0
         Left            =   3720
         Stretch         =   -1  'True
         Top             =   3240
         Width           =   3015
      End
      Begin VB.Image Image3 
         Height          =   2535
         Index           =   0
         Left            =   240
         Stretch         =   -1  'True
         Top             =   3240
         Width           =   3015
      End
      Begin VB.Image Image2 
         Height          =   2655
         Index           =   0
         Left            =   3720
         Stretch         =   -1  'True
         Top             =   360
         Width           =   3015
      End
      Begin VB.Image Image1 
         Height          =   2655
         Index           =   0
         Left            =   240
         Stretch         =   -1  'True
         Top             =   360
         Width           =   3015
      End
   End
   Begin VB.Menu �D�\ 
      Caption         =   "�D�\"
      Begin VB.Menu ���X���׸q�j�Q�� 
         Caption         =   "���X���׸q�j�Q��"
      End
      Begin VB.Menu ù�Ǯ��A�q�j�Q�� 
         Caption         =   "ù�Ǯ��A�q�j�Q��"
      End
      Begin VB.Menu ���o�t���q�j�Q�� 
         Caption         =   "���o�t���q�j�Q��"
      End
      Begin VB.Menu �հs�����q�j�Q�� 
         Caption         =   "�հs�����q�j�Q��"
      End
   End
   Begin VB.Menu �F�� 
      Caption         =   "�F��"
      Begin VB.Menu ���G�u��F�� 
         Caption         =   "���G�u��F��"
      End
      Begin VB.Menu �ͬK���F�� 
         Caption         =   "�ͬK���F��"
      End
      Begin VB.Menu �ޥյ��C���F�� 
         Caption         =   "�ޥյ��C���F��"
      End
      Begin VB.Menu �����X�F�� 
         Caption         =   "�����X�F��"
      End
   End
   Begin VB.Menu �@�� 
      Caption         =   "�@��"
      Begin VB.Menu �j�[��Ĵ� 
         Caption         =   "�j�[��Ĵ�"
      End
      Begin VB.Menu ���X���A�� 
         Caption         =   "���X���A��"
      End
      Begin VB.Menu �u�C������ 
         Caption         =   "�u�C������"
      End
      Begin VB.Menu ������״� 
         Caption         =   "������״�"
      End
   End
   Begin VB.Menu ���I 
      Caption         =   "���I"
      Begin VB.Menu ��N�B�N�O 
         Caption         =   "��N�B�N�O"
      End
      Begin VB.Menu �۪������ 
         Caption         =   "�۪������"
      End
      Begin VB.Menu �եɵ��� 
         Caption         =   "�եɵ���"
      End
      Begin VB.Menu �¿}���T 
         Caption         =   "�¿}���T"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim A, B, C, D, S(1)





Private Sub �j�[��Ĵ�_Click()
Image3(C).Picture = LoadPicture(App.Path & "/9.bmp")
C = C + 1
S(C) = S(C) + 45
End Sub

Private Sub ������״�_Click()
Image3(C).Picture = LoadPicture(App.Path & "/12.bmp")
C = C + 1
S(C) = S(C) + 50
End Sub

Private Sub ���G�u��F��_Click()
Image2(B).Picture = LoadPicture(App.Path & "/5.bmp")
B = B + 1
S(B) = S(B) + 45
End Sub

Private Sub �����X�F��_Click()
Image2(B).Picture = LoadPicture(App.Path & "/8.bmp")
B = B + 1
S(B) = S(B) + 40
End Sub



Private Sub ���o�t���q�j�Q��_Click()
Image1(A).Picture = LoadPicture(App.Path & "/3.bmp")
A = A + 1
S(A) = S(A) + 65
End Sub

Private Sub �ͬK���F��_Click()
Image2(B).Picture = LoadPicture(App.Path & "/6.bmp")
B = B + 1
S(B) = S(B) + 40
End Sub

Private Sub �եɵ���_Click()
Image4(D).Picture = LoadPicture(App.Path & "/15.bmp")
D = D + 1
S(D) = S(D) + 40
End Sub

Private Sub �հs�����q�j�Q��_Click()
Image2(B).Picture = LoadPicture(App.Path & "/4.bmp")
A = A + 1
S(A) = S(A) + 80
End Sub

Private Sub �u�C������_Click()
Image3(C).Picture = LoadPicture(App.Path & "/11.bmp")
C = C + 1
S(C) = S(C) + 40
End Sub

Private Sub ��N�B�N�O_Click()
Image4(D).Picture = LoadPicture(App.Path & "/13.bmp")
D = D + 1
S(D) = S(D) + 40
End Sub

Private Sub �۪������_Click()
Image4(D).Picture = LoadPicture(App.Path & "/14.bmp")
D = D + 1
S(D) = S(D) + 45
End Sub

Private Sub �¿}���T_Click()
Image4(D).Picture = LoadPicture(App.Path & "/16.bmp")
D = D + 1
S(D) = S(D) + 35
End Sub

Private Sub �ޥյ��C���F��_Click()
Image1(A).Picture = LoadPicture(App.Path & "/7.bmp")
B = B + 1
S(B) = S(B) + 50
End Sub

Private Sub ���X���A��_Click()
Image3(C).Picture = LoadPicture(App.Path & "/10.bmp")
C = C + 1

S(C) = S(C) + 40
End Sub

Private Sub ���X���׸q�j�Q��_Click()
Image1(A).Picture = LoadPicture(App.Path & "/1.bmp")
A = A + 1
S(A) = S(A) + 75
End Sub

Private Sub ù�Ǯ��A�q�j�Q��_Click()
Image1(A).Picture = LoadPicture(App.Path & "/2.bmp")
A = A + 1
S(A) = S(A) + 70
End Sub
