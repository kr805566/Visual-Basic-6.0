VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "�ϧβ��;�"
   ClientHeight    =   10095
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9705
   LinkTopic       =   "Form1"
   ScaleHeight     =   10095
   ScaleWidth      =   9705
   StartUpPosition =   3  '�t�ιw�]��
   Begin VB.Frame �C���ܮج[ 
      Caption         =   "�C����"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   11.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   840
      TabIndex        =   12
      Top             =   8880
      Width           =   8775
      Begin VB.PictureBox �C���� 
         Appearance      =   0  '����
         BackColor       =   &H00400000&
         ForeColor       =   &H80000008&
         Height          =   615
         Index           =   11
         Left            =   8040
         ScaleHeight     =   585
         ScaleWidth      =   585
         TabIndex        =   24
         Top             =   360
         Width           =   615
      End
      Begin VB.PictureBox �C���� 
         Appearance      =   0  '����
         BackColor       =   &H00C0C000&
         ForeColor       =   &H80000008&
         Height          =   615
         Index           =   10
         Left            =   7320
         ScaleHeight     =   585
         ScaleWidth      =   585
         TabIndex        =   23
         Top             =   360
         Width           =   615
      End
      Begin VB.PictureBox �C���� 
         Appearance      =   0  '����
         BackColor       =   &H0000C000&
         ForeColor       =   &H80000008&
         Height          =   615
         Index           =   9
         Left            =   6600
         ScaleHeight     =   585
         ScaleWidth      =   585
         TabIndex        =   22
         Top             =   360
         Width           =   615
      End
      Begin VB.PictureBox �C���� 
         Appearance      =   0  '����
         BackColor       =   &H0000C0C0&
         ForeColor       =   &H80000008&
         Height          =   615
         Index           =   8
         Left            =   5880
         ScaleHeight     =   585
         ScaleWidth      =   585
         TabIndex        =   21
         Top             =   360
         Width           =   615
      End
      Begin VB.PictureBox �C���� 
         Appearance      =   0  '����
         BackColor       =   &H000040C0&
         ForeColor       =   &H80000008&
         Height          =   615
         Index           =   7
         Left            =   5160
         ScaleHeight     =   585
         ScaleWidth      =   585
         TabIndex        =   20
         Top             =   360
         Width           =   615
      End
      Begin VB.PictureBox �C���� 
         Appearance      =   0  '����
         BackColor       =   &H000000FF&
         ForeColor       =   &H80000008&
         Height          =   615
         Index           =   6
         Left            =   4440
         ScaleHeight     =   585
         ScaleWidth      =   585
         TabIndex        =   19
         Top             =   360
         Width           =   615
      End
      Begin VB.PictureBox �C���� 
         Appearance      =   0  '����
         BackColor       =   &H00FF00FF&
         ForeColor       =   &H80000008&
         Height          =   615
         Index           =   5
         Left            =   3720
         ScaleHeight     =   585
         ScaleWidth      =   585
         TabIndex        =   18
         Top             =   360
         Width           =   615
      End
      Begin VB.PictureBox �C���� 
         Appearance      =   0  '����
         BackColor       =   &H00FF0000&
         ForeColor       =   &H80000008&
         Height          =   615
         Index           =   4
         Left            =   3000
         ScaleHeight     =   585
         ScaleWidth      =   585
         TabIndex        =   17
         Top             =   360
         Width           =   615
      End
      Begin VB.PictureBox �C���� 
         Appearance      =   0  '����
         BackColor       =   &H00FFFF00&
         ForeColor       =   &H80000008&
         Height          =   615
         Index           =   3
         Left            =   2280
         ScaleHeight     =   585
         ScaleWidth      =   585
         TabIndex        =   16
         Top             =   360
         Width           =   615
      End
      Begin VB.PictureBox �C���� 
         Appearance      =   0  '����
         BackColor       =   &H0000FF00&
         ForeColor       =   &H80000008&
         Height          =   615
         Index           =   2
         Left            =   1560
         ScaleHeight     =   585
         ScaleWidth      =   585
         TabIndex        =   15
         Top             =   360
         Width           =   615
      End
      Begin VB.PictureBox �C���� 
         Appearance      =   0  '����
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H80000008&
         Height          =   615
         Index           =   1
         Left            =   840
         ScaleHeight     =   585
         ScaleWidth      =   585
         TabIndex        =   14
         Top             =   360
         Width           =   615
      End
      Begin VB.PictureBox �C���� 
         Appearance      =   0  '����
         BackColor       =   &H000080FF&
         ForeColor       =   &H80000008&
         Height          =   615
         Index           =   0
         Left            =   120
         ScaleHeight     =   585
         ScaleWidth      =   585
         TabIndex        =   13
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.PictureBox �C�� 
      Appearance      =   0  '����
      BackColor       =   &H000000FF&
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   120
      ScaleHeight     =   585
      ScaleWidth      =   585
      TabIndex        =   11
      Top             =   9240
      Width           =   615
   End
   Begin VB.Frame �ϧβ��ͱ���ج[ 
      Caption         =   "���ͱ���"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      TabIndex        =   2
      Top             =   7200
      Width           =   9495
      Begin VB.Frame �I���C��ج[ 
         Caption         =   "�I���C��"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   11.25
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   7440
         TabIndex        =   8
         Top             =   360
         Width           =   1815
         Begin VB.PictureBox �I���C�� 
            Appearance      =   0  '����
            BackColor       =   &H00000000&
            ForeColor       =   &H80000008&
            Height          =   615
            Index           =   1
            Left            =   960
            ScaleHeight     =   585
            ScaleWidth      =   585
            TabIndex        =   10
            Top             =   360
            Width           =   615
         End
         Begin VB.PictureBox �I���C�� 
            Appearance      =   0  '����
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   615
            Index           =   0
            Left            =   240
            ScaleHeight     =   585
            ScaleWidth      =   585
            TabIndex        =   9
            Top             =   360
            Width           =   615
         End
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         Left            =   3840
         Max             =   2000
         Min             =   10
         TabIndex        =   7
         Top             =   1200
         Value           =   10
         Width           =   2895
      End
      Begin VB.TextBox ��� 
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   14.25
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3840
         TabIndex        =   6
         Text            =   "10"
         Top             =   720
         Width           =   2895
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   14.25
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         ItemData        =   "�ϧβ��;�.frx":0000
         Left            =   120
         List            =   "�ϧβ��;�.frx":000D
         TabIndex        =   4
         Text            =   "���"
         Top             =   720
         Width           =   2895
      End
      Begin VB.Label ���Lab 
         Caption         =   "�b�|�G"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   11.25
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3840
         TabIndex        =   5
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label �Ϊ�Lab 
         Caption         =   "�Ϊ��G"
         BeginProperty Font 
            Name            =   "�s�ө���"
            Size            =   11.25
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame �ϧβ��ͮج[ 
      Caption         =   "�ϧβ���"
      BeginProperty Font 
         Name            =   "�s�ө���"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9495
      Begin VB.PictureBox Pictures 
         Appearance      =   0  '����
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   6615
         Left            =   120
         ScaleHeight     =   6585
         ScaleWidth      =   9225
         TabIndex        =   1
         Top             =   240
         Width           =   9255
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub HScroll1_Change() ' HScroll1 �O [���b] ���� �b�ݩʪ� Max (�̤j��) Min (�̤p��) �]�w�̤j�Ȩ�̤p�� �ӽd��N�O Min ~ Max
���.Text = HScroll1.Value ' ��� ( TextBox ��r������� ���e = HScroll1 (���b����) �� Value (�ƭ�) )
End Sub

Private Sub �I���C��_Click(Index As Integer) ' ��ܭI��
Pictures.BackColor = �I���C��(Index).BackColor ' Pictures ( Picture �Ϥ����� �ΨӦb�W���e�ϧ� ) �� BackColor (�I���ϼ�) = �I���Ϥ� ( �]�O�� Picture "�}�C" ���� (0) �O�զ� (1) �O�¦� )
End Sub

Private Sub �C����_Click(Index As Integer) ' ��ܤ�r�C��
�C��.BackColor = �C����(Index).BackColor
End Sub



Private Sub Pictures_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

Pictures.ForeColor = �C��.BackColor


Select Case Combo1.Text
    Case "���"
        Pictures.Circle (X, Y), ��� ' .Circle ( X,Y ) , �b�|
    Case "�����"
        Pictures.Line (X, Y)-(X + ��� * Sqr(2), Y + ��� * 2 ^ 0.5), , B
    Case "���T����"
        Pictures.Line (X, Y)-(X + 0.5 * ���, Y + Sqr(1 ^ 2 - 0.5 ^ 2) * ���)
        Pictures.Line (X, Y)-(X - 0.5 * ���, Y + Sqr(1 ^ 2 - 0.5 ^ 2) * ���)
        Pictures.Line (X + 0.5 * ���, Y + Sqr(1 ^ 2 - 0.5 ^ 2) * ���)-(X - 0.5 * ���, Y + Sqr(1 ^ 2 - 0.5 ^ 2) * ���)
End Select

End Sub



Private Sub Combo1_Click()

Select Case Combo1.Text
    Case "���"
        ���Lab.Caption = "�b�|�G"
    Case Else
        ���Lab.Caption = "����G"
End Select

End Sub

