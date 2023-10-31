VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  '平面
   BackColor       =   &H80000005&
   Caption         =   "Form1"
   ClientHeight    =   12660
   ClientLeft      =   2130
   ClientTop       =   735
   ClientWidth     =   13185
   LinkTopic       =   "Form1"
   ScaleHeight     =   12660
   ScaleWidth      =   13185
   Begin VB.CommandButton Command1 
      Caption         =   "比你們都白"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   21.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   22
      Left            =   8040
      TabIndex        =   30
      Top             =   11520
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "←比你還白"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   21.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   21
      Left            =   5400
      TabIndex        =   29
      Top             =   11520
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "真的好白喔"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   21.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   20
      Left            =   2760
      TabIndex        =   28
      Top             =   11520
      Width           =   2415
   End
   Begin VB.CommandButton 停止計時 
      Caption         =   "停止計時"
      Height          =   615
      Left            =   3000
      TabIndex        =   27
      Top             =   120
      Width           =   495
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   11640
      Top             =   480
   End
   Begin VB.CommandButton Command1 
      Caption         =   "RGB黑白"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   21.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   19
      Left            =   10680
      TabIndex        =   25
      Top             =   10320
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "淺藍"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   21.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   18
      Left            =   8040
      TabIndex        =   24
      Top             =   10320
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "淺綠"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   21.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   17
      Left            =   5400
      TabIndex        =   23
      Top             =   10320
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "淺紅"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   21.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   16
      Left            =   2760
      TabIndex        =   22
      Top             =   10320
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "快速複製"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   21.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   15
      Left            =   120
      TabIndex        =   21
      Top             =   10320
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "自內外相反"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   21.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   14
      Left            =   10680
      TabIndex        =   20
      Top             =   9120
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "深藍"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   21.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   13
      Left            =   8040
      TabIndex        =   19
      Top             =   9120
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "深綠"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   21.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   12
      Left            =   5400
      TabIndex        =   18
      Top             =   9120
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "深紅"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   21.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   11
      Left            =   2760
      TabIndex        =   17
      Top             =   9120
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "自外內相反"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   21.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   10
      Left            =   120
      TabIndex        =   16
      Top             =   9120
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "左右反"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   21.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   3
      Left            =   8040
      TabIndex        =   15
      Top             =   6720
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "複製"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   21.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   2
      Left            =   5400
      TabIndex        =   14
      Top             =   6720
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "全色轉黑白"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   21.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   1
      Left            =   2760
      TabIndex        =   13
      Top             =   6720
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "兩色轉一色"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   21.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   9
      Left            =   10680
      TabIndex        =   11
      Top             =   7920
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "外暗內亮"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   21.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   8
      Left            =   8040
      TabIndex        =   10
      Top             =   7920
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "暗到亮到暗"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   21.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   7
      Left            =   5400
      TabIndex        =   9
      Top             =   7920
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "暗到亮"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   21.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   6
      Left            =   2760
      TabIndex        =   8
      Top             =   7920
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "三分段"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   21.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   5
      Left            =   120
      TabIndex        =   7
      Top             =   7920
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "相反色"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   21.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   4
      Left            =   10680
      TabIndex        =   6
      Top             =   6720
      Width           =   2415
   End
   Begin VB.CommandButton 二化 
      Caption         =   "二化"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   36
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   8880
      TabIndex        =   5
      Top             =   120
      Width           =   2415
   End
   Begin VB.CommandButton 灰階 
      Caption         =   "灰階"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   36
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   6240
      TabIndex        =   4
      Top             =   120
      Width           =   2415
   End
   Begin VB.CommandButton 載入 
      Caption         =   "載入"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   36
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3600
      TabIndex        =   3
      Top             =   120
      Width           =   2415
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  '平面
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4455
      Left            =   6600
      ScaleHeight     =   4425
      ScaleWidth      =   6180
      TabIndex        =   2
      Top             =   1320
      Width           =   6210
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  '平面
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4455
      Left            =   120
      ScaleHeight     =   4425
      ScaleWidth      =   6180
      TabIndex        =   1
      Top             =   1320
      Width           =   6210
   End
   Begin VB.CommandButton Command1 
      Caption         =   "單色轉黑白"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   21.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   6720
      Width           =   2415
   End
   Begin VB.Label Label2 
      Alignment       =   2  '置中對齊
      Caption         =   " 10 秒內載入"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   26.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   26
      Top             =   120
      Width           =   2895
   End
   Begin VB.Label Label1 
      Alignment       =   2  '置中對齊
      Appearance      =   0  '平面
      BackColor       =   &H80000005&
      Caption         =   "---------------------↓亂做區↓--------------------"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   36
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   120
      TabIndex        =   12
      Top             =   5880
      Width           =   12855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim time, s22

Private Sub Combo1_Change()

End Sub

Private Sub Command2_Click()

End Sub

Private Sub 二化_Click()
For i = 0 To Picture1.Width Step 15
    For j = 0 To Picture1.Height Step 15
    s = Picture1.Point(i, j)
    If s22 < s Then
    Picture2.PSet (i, j), RGB(255, 255, 255)
    Else
    Picture2.PSet (i, j), RGB(0, 0, 0)
    End If
Next j, i


End Sub

Private Sub 載入_Click()
Picture1.Picture = LoadPicture(App.Path & "/2.jpg")
End Sub
Private Sub 灰階_Click()

For i = 0 To Picture1.Width Step 15
    For j = 0 To Picture1.Height Step 15
    s = Picture1.Point(i, j)
   t = Abs(s Mod 256) * 0.299 + Abs(s \ 256 Mod 256) * 0.587 + Abs(s \ 256 \ 256 Mod 256) * 0.114
        Picture2.PSet (i, j), RGB(t, t, t)
    ''''''''''''''''''''''''''''''''''''''''''
                   '
    ''''''''''''''''''''''''''''''''''''''''''
Next j, i

'''''''''''''''''''''

'''''''''''''''''''''

End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub Timer1_Timer()
time = time + 1
Label2 = " " & 10 - time & " 秒內載入"
If time = 10 Then 載入 = True: Label2.Visible = False: Timer1.Enabled = False: 停止計時.Visible = False
End Sub

Private Sub 停止計時_Click()
Label2.Visible = False
Timer1.Enabled = False
停止計時.Visible = False
載入.SetFocus
End Sub
Private Sub Command1_Click(Index As Integer)
Select Case Index
Case 0

For i = 0 To Picture1.Width Step 15
    For j = 0 To Picture1.Height Step 15
    s = Picture1.Point(i, j)
    Picture2.PSet (i, j), RGB(Abs(s Mod 256), Abs(s Mod 256), Abs(s Mod 256))
Next j, i

Case 1
For i = 0 To Picture1.Width Step 15
    For j = 0 To Picture1.Height Step 15
    s = Picture1.Point(i, j)
    Picture2.PSet (i, j), RGB((Abs(s Mod 256) + Abs(s \ 256 Mod 256) + Abs(s \ 256 \ 256 Mod 256)) / 3, (Abs(s Mod 256) + Abs(s \ 256 Mod 256) + Abs(s \ 256 \ 256 Mod 256)) / 3, (Abs(s Mod 256) + Abs(s \ 256 Mod 256) + Abs(s \ 256 \ 256 Mod 256)) / 3)
Next j, i
Case 2
For i = 0 To Picture1.Width - 1 Step 15
    For j = 0 To Picture1.Height - 1 Step 15
    s = Picture1.Point(i, j)
    Picture2.PSet (i, j), RGB(Abs(s Mod 256), Abs(s \ 256 Mod 256), Abs(s \ 256 \ 256 Mod 256))
Next j, i
Case 3
For j = 0 To Picture1.Height Step 15
    For i = 0 To Picture1.Width Step 15
    s = Picture1.Point(i - 45, j)
    Picture2.PSet (Picture1.Width - i, j), RGB(Abs(s Mod 256), Abs(s \ 256 Mod 256), Abs(s \ 256 \ 256 Mod 256))
Next i, j
Case 4
For i = 0 To Picture1.Width Step 15
    For j = 0 To Picture1.Height Step 15
    s = Picture1.Point(i, j)
    Picture2.PSet (i, j), RGB(Abs(255 - s Mod 256), Abs(255 - s \ 256 Mod 256), Abs(255 - s \ 256 \ 256 Mod 256))
Next j, i
Case 5
For i = 0 To Picture1.Width Step 15
    For j = 0 To Picture1.Height Step 15
    s = Picture1.Point(i, j)

    If i <= Picture1.Width \ 3 Then
    Picture2.PSet (i, j), RGB(Abs(s Mod 256), Abs(s \ 256 Mod 256), Abs(s \ 256 \ 256 Mod 256))
    ElseIf i >= Picture1.Width * 2 \ 3 Then
    Picture2.PSet (i, j), RGB(Abs(255 - s Mod 256), Abs(255 - s Mod 256), Abs(255 - s Mod 256))
    ElseIf i < Picture1.Width * 2 \ 3 And i > Picture1.Width \ 3 Then
    Picture2.PSet (i, j), RGB(Abs(s Mod 256) * 0.299, Abs(s \ 256 Mod 256) * 0.587, Abs(s \ 256 \ 256 Mod 256) * 0.114)
    End If
Next j, i
Case 6
c = 0
For i = 0 To Picture1.Width Step 15
    For j = 0 To Picture1.Height Step 15
    s = Picture1.Point(i, j)
    Picture2.PSet (i, j), RGB(Abs(s Mod 256) * c, Abs(s \ 256 Mod 256) * c, Abs(s \ 256 \ 256 Mod 256) * c)
Next j
c = c + 0.0024
Next i
Case 7
c = 0
a = 0.0048
For i = 0 To Picture1.Width Step 15
    For j = 0 To Picture1.Height Step 15
    s = Picture1.Point(i, j)
    Picture2.PSet (i, j), RGB(Abs(s Mod 256) * c, Abs(s \ 256 Mod 256) * c, Abs(s \ 256 \ 256 Mod 256) * c)
Next j
c = c + a
If c > 1 Then c = c - a: a = -a
Next i
Case 8
a = 0
x = 0.0066

c = 0
z = 0.0048

For i = 0 To Picture1.Width Step 15
    For j = 0 To Picture1.Height Step 15
    a = a + x
    If a > 1 Then a = a - x: x = -x
    s = Picture1.Point(i, j)
    Picture2.PSet (i, j), RGB(Abs(s Mod 256) * (c + a) / 2, Abs(s \ 256 Mod 256) * (c + a) / 2, Abs(s \ 256 \ 256 Mod 256) * (c + a) / 2)
Next j
x = -x
a = 0.0066
c = c + z
If c > 1 Then c = c - z: z = -z
Next i
Case 9
For i = 0 To Picture1.Width Step 15
    For j = 0 To Picture1.Height Step 15
    s = Picture1.Point(i, j)
    Picture2.PSet (i, j), RGB((Abs(s \ 256 Mod 256) + Abs(s \ 256 \ 256 Mod 256)) / 2, (Abs(s Mod 256) + Abs(s \ 256 \ 256 Mod 256)) / 2, (Abs(s Mod 256) + Abs(s \ 256 Mod 256)) / 2)
Next j, i
Case 10
For i = 0 To Picture1.Width / 2 Step 15
    For j = 0 To Picture1.Height Step 15
    s1 = Picture2.Point(i, j)
    s2 = Picture2.Point(Picture1.Width - i, j)
    Picture2.PSet (i, j), RGB(Abs(s2 Mod 256), Abs(s2 \ 256 Mod 256), Abs(s2 \ 256 \ 256 Mod 256))
    Picture2.PSet (Picture1.Width - i, j), RGB(Abs(s1 Mod 256), Abs(s1 \ 256 Mod 256), Abs(s1 \ 256 \ 256 Mod 256))
    

Next j, i
Case 11
For i = 0 To Picture1.Width Step 15
    For j = 0 To Picture1.Height Step 15
    s = Picture1.Point(i, j)
    Picture2.PSet (i, j), RGB(Abs(s Mod 256), Abs(s \ 256 Mod 256) * 0.3, Abs(s \ 256 \ 256 Mod 256) * 0.3)
Next j, i
Case 12
For i = 0 To Picture1.Width Step 15
    For j = 0 To Picture1.Height Step 15
    s = Picture1.Point(i, j)
    Picture2.PSet (i, j), RGB(Abs(s Mod 256) * 0.3, Abs(s \ 256 Mod 256), Abs(s \ 256 \ 256 Mod 256) * 0.3)
Next j, i
Case 13
For i = 0 To Picture1.Width Step 15
    For j = 0 To Picture1.Height Step 15
    s = Picture1.Point(i, j)
    Picture2.PSet (i, j), RGB(Abs(s Mod 256) * 0.3, Abs(s \ 256 Mod 256) * 0.3, Abs(s \ 256 \ 256 Mod 256))
Next j, i
Case 14
For i = 0 To Picture1.Width / 2 Step 15
    For j = 0 To Picture1.Height Step 15
    s1 = Picture2.Point(Picture1.Width / 2 - i, j)
    s2 = Picture2.Point(Picture1.Width / 2 + i, j)
    Picture2.PSet (Picture1.Width / 2 + i, j), RGB(Abs(s1 Mod 256), Abs(s1 \ 256 Mod 256), Abs(s1 \ 256 \ 256 Mod 256))
    Picture2.PSet (Picture1.Width / 2 - i, j), RGB(Abs(s2 Mod 256), Abs(s2 \ 256 Mod 256), Abs(s2 \ 256 \ 256 Mod 256))
    

Next j, i
Case 15
Picture2.Picture = LoadPicture(App.Path & "/2.jpg")
Case 16
For i = 0 To Picture1.Width Step 15
    For j = 0 To Picture1.Height Step 15
    s = Picture1.Point(i, j)
    Picture2.PSet (i, j), RGB(Abs(s Mod 256), Abs(s \ 256 Mod 256) * 0.7, Abs(s \ 256 \ 256 Mod 256) * 0.7)
Next j, i
Case 17
For i = 0 To Picture1.Width Step 15
    For j = 0 To Picture1.Height Step 15
    s = Picture1.Point(i, j)
    Picture2.PSet (i, j), RGB(Abs(s Mod 256) * 0.7, Abs(s \ 256 Mod 256), Abs(s \ 256 \ 256 Mod 256) * 0.7)
Next j, i
Case 18
For i = 0 To Picture1.Width Step 15
    For j = 0 To Picture1.Height Step 15
    s = Picture1.Point(i, j)
    Picture2.PSet (i, j), RGB(Abs(s Mod 256) * 0.7, Abs(s \ 256 Mod 256) * 0.7, Abs(s \ 256 \ 256 Mod 256))
Next j, i
Case 19
s = 0
g = 0
For i = 0 To Picture1.Width Step 15
    For j = 0 To Picture1.Height Step 15
    s = s + Picture1.Point(i, j)
    g = g + 1
Next j, i
s2 = s / g * 0.7
For i = 0 To Picture1.Width Step 15
    For j = 0 To Picture1.Height Step 15
    s = Picture1.Point(i, j)
    If s > s2 Then
     Picture2.PSet (i, j), RGB(255, 255, 255)
   Else
    Picture2.PSet (i, j), RGB(0, 0, 0)
   End If
Next j, i
    


Case 20
Picture2.Picture = LoadPicture
Case 21
Picture1.Picture = LoadPicture
Picture2.Picture = LoadPicture
Case 22

For i = 0 To 22
Command1(i).Visible = False
Next i
載入.Visible = False
灰階.Visible = False
二化.Visible = False
Label1.Visible = False
Label2.Visible = False
停止計時.Value = False
Picture1.Visible = False
Picture2.Visible = False
End Select

End Sub


