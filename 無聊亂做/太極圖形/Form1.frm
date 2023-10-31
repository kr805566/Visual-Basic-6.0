VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7575
   ClientLeft      =   2685
   ClientTop       =   1875
   ClientWidth     =   9495
   LinkTopic       =   "Form1"
   ScaleHeight     =   7575
   ScaleWidth      =   9495
   Begin VB.Frame Frame1 
      Caption         =   "樣式"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   240
      TabIndex        =   1
      Top             =   6000
      Width           =   6255
      Begin VB.CommandButton Command3 
         Caption         =   "結束"
         Height          =   375
         Left            =   4920
         TabIndex        =   13
         Top             =   960
         Width           =   1215
      End
      Begin VB.VScrollBar VScroll2 
         Height          =   375
         Left            =   4320
         Max             =   1
         Min             =   30
         TabIndex        =   11
         Top             =   840
         Value           =   1
         Width           =   255
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  '置中對齊
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3360
         TabIndex        =   10
         Text            =   "100"
         Top             =   840
         Width           =   975
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   375
         Left            =   4320
         Max             =   1
         Min             =   30
         TabIndex        =   8
         Top             =   360
         Value           =   1
         Width           =   255
      End
      Begin VB.CommandButton Command2 
         Caption         =   "清除"
         Height          =   375
         Left            =   4920
         TabIndex        =   7
         Top             =   600
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "繪圖"
         Height          =   375
         Left            =   4920
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  '置中對齊
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3360
         TabIndex        =   5
         Text            =   "100"
         Top             =   360
         Width           =   975
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "Form1.frx":0000
         Left            =   960
         List            =   "Form1.frx":000D
         TabIndex        =   2
         Text            =   "Combo1"
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label4 
         Alignment       =   2  '置中對齊
         Caption         =   " "
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   14.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label3 
         Alignment       =   2  '置中對齊
         Caption         =   "寬度 :"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   14.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         TabIndex        =   9
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label2 
         Alignment       =   2  '置中對齊
         Caption         =   "長度 :"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   14.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         TabIndex        =   4
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   2  '置中對齊
         Caption         =   "圖型 :"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   14.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   5715
      Left            =   120
      ScaleHeight     =   5655
      ScaleWidth      =   9015
      TabIndex        =   0
      Top             =   120
      Width           =   9075
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
x = Picture1.Width / 2
y = Picture1.Height / 2

Select Case Combo1.ListIndex

Case 0
    a = Val(Text1)
    pi = 3.14159265358979
    Picture1.Cls
    Picture1.FillStyle = 0
    Picture1.FillColor = RGB(255, 255, 255)
    Picture1.Circle (x, y), a, RGB(255, 255, 255), -pi * 3 / 2, -pi / 2
    Picture1.FillColor = RGB(0, 0, 0)
    Picture1.Circle (x, y), a, RGB(0, 0, 0), -pi / 2, -pi * 3 / 2
    Picture1.FillColor = RGB(255, 255, 255)
    Picture1.Circle (x, y - a / 2), a / 2, RGB(255, 255, 255), -pi / 2, -pi * 3 / 2
    Picture1.FillColor = RGB(0, 0, 0)
    Picture1.Circle (x, y + a / 2), a / 2, RGB(0, 0, 0), -pi * 3 / 2, -pi / 2
    Picture1.FillColor = RGB(255, 255, 255)
    Picture1.Circle (x, y + a / 2), a / 8, RGB(255, 255, 255)
    Picture1.FillColor = RGB(0, 0, 0)
    Picture1.Circle (x, y - a / 2), a / 8, RGB(0, 0, 0)
Case 1
    Picture1.Cls
    Picture1.FillStyle = 1
    For i = 1 To 20
    Picture1.Line (x - 100 * i, y - 100 * i)-(x + 100 * i, y + 100 * i), , B
    Next i
Case 2
    Picture1.Cls
    Picture1.FillStyle = 1
    Picture1.Line (x, y - 1000)-(x, y + 1000)
    Picture1.Line (x - 1000, y)-(x + 1000, y)
    Picture1.Circle (x, y), 1000
    Picture1.Line (x, y - 500 * Sqr(2) * 2)-(x - 500 * Sqr(2) * 2, y)
    Picture1.Line (x, y + 500 * Sqr(2) * 2)-(x + 500 * Sqr(2) * 2, y)
    Picture1.Line (x, y + 500 * Sqr(2) * 2)-(x - 500 * Sqr(2) * 2, y)
    Picture1.Line (x, y - 500 * Sqr(2) * 2)-(x + 500 * Sqr(2) * 2, y)
    Picture1.Line (x - 500 * Sqr(2) * 2, y - 500 * Sqr(2) * 2)-(x + 500 * Sqr(2) * 2, y + 500 * Sqr(2) * 2), , B
End Select




End Sub



Private Sub Command3_Click()
End
End Sub

Private Sub Form_Load()
Combo1.ListIndex = 0
End Sub

Private Sub VScroll1_Change()
Select Case Combo1.ListIndex

Case 0
Text1 = VScroll1.Value * 100
Text2 = VScroll1.Value * 100
End Select

End Sub

Private Sub VScroll2_Change()

Select Case Combo1.ListIndex

Case 0
Text1 = VScroll1.Value * 100
Text2 = VScroll1.Value * 100
End Select
End Sub
