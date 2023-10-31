VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "圖形產生器"
   ClientHeight    =   10095
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9705
   LinkTopic       =   "Form1"
   ScaleHeight     =   10095
   ScaleWidth      =   9705
   StartUpPosition =   3  '系統預設值
   Begin VB.Frame 顏色選擇框架 
      Caption         =   "顏色選擇"
      BeginProperty Font 
         Name            =   "新細明體"
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
      Begin VB.PictureBox 顏色選擇 
         Appearance      =   0  '平面
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
      Begin VB.PictureBox 顏色選擇 
         Appearance      =   0  '平面
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
      Begin VB.PictureBox 顏色選擇 
         Appearance      =   0  '平面
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
      Begin VB.PictureBox 顏色選擇 
         Appearance      =   0  '平面
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
      Begin VB.PictureBox 顏色選擇 
         Appearance      =   0  '平面
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
      Begin VB.PictureBox 顏色選擇 
         Appearance      =   0  '平面
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
      Begin VB.PictureBox 顏色選擇 
         Appearance      =   0  '平面
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
      Begin VB.PictureBox 顏色選擇 
         Appearance      =   0  '平面
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
      Begin VB.PictureBox 顏色選擇 
         Appearance      =   0  '平面
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
      Begin VB.PictureBox 顏色選擇 
         Appearance      =   0  '平面
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
      Begin VB.PictureBox 顏色選擇 
         Appearance      =   0  '平面
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
      Begin VB.PictureBox 顏色選擇 
         Appearance      =   0  '平面
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
   Begin VB.PictureBox 顏色 
      Appearance      =   0  '平面
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
   Begin VB.Frame 圖形產生控制框架 
      Caption         =   "產生控制"
      BeginProperty Font 
         Name            =   "新細明體"
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
      Begin VB.Frame 背景顏色框架 
         Caption         =   "背景顏色"
         BeginProperty Font 
            Name            =   "新細明體"
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
         Begin VB.PictureBox 背景顏色 
            Appearance      =   0  '平面
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
         Begin VB.PictureBox 背景顏色 
            Appearance      =   0  '平面
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
      Begin VB.TextBox 邊長 
         BeginProperty Font 
            Name            =   "新細明體"
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
            Name            =   "新細明體"
            Size            =   14.25
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         ItemData        =   "圖形產生器.frx":0000
         Left            =   120
         List            =   "圖形產生器.frx":000D
         TabIndex        =   4
         Text            =   "圓形"
         Top             =   720
         Width           =   2895
      End
      Begin VB.Label 顯示Lab 
         Caption         =   "半徑："
         BeginProperty Font 
            Name            =   "新細明體"
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
      Begin VB.Label 形狀Lab 
         Caption         =   "形狀："
         BeginProperty Font 
            Name            =   "新細明體"
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
   Begin VB.Frame 圖形產生框架 
      Caption         =   "圖形產生"
      BeginProperty Font 
         Name            =   "新細明體"
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
         Appearance      =   0  '平面
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
Private Sub HScroll1_Change() ' HScroll1 是 [捲軸] 物件 在屬性的 Max (最大值) Min (最小值) 設定最大值到最小值 而範圍就是 Min ~ Max
邊長.Text = HScroll1.Value ' 邊長 ( TextBox 文字方塊物件 內容 = HScroll1 (捲軸物件) 的 Value (數值) )
End Sub

Private Sub 背景顏色_Click(Index As Integer) ' 選擇背景
Pictures.BackColor = 背景顏色(Index).BackColor ' Pictures ( Picture 圖片物件 用來在上面畫圖形 ) 的 BackColor (背景圖樣) = 背景圖片 ( 也是個 Picture "陣列" 物件 (0) 是白色 (1) 是黑色 )
End Sub

Private Sub 顏色選擇_Click(Index As Integer) ' 選擇文字顏色
顏色.BackColor = 顏色選擇(Index).BackColor
End Sub



Private Sub Pictures_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

Pictures.ForeColor = 顏色.BackColor


Select Case Combo1.Text
    Case "圓形"
        Pictures.Circle (X, Y), 邊長 ' .Circle ( X,Y ) , 半徑
    Case "正方形"
        Pictures.Line (X, Y)-(X + 邊長 * Sqr(2), Y + 邊長 * 2 ^ 0.5), , B
    Case "正三角形"
        Pictures.Line (X, Y)-(X + 0.5 * 邊長, Y + Sqr(1 ^ 2 - 0.5 ^ 2) * 邊長)
        Pictures.Line (X, Y)-(X - 0.5 * 邊長, Y + Sqr(1 ^ 2 - 0.5 ^ 2) * 邊長)
        Pictures.Line (X + 0.5 * 邊長, Y + Sqr(1 ^ 2 - 0.5 ^ 2) * 邊長)-(X - 0.5 * 邊長, Y + Sqr(1 ^ 2 - 0.5 ^ 2) * 邊長)
End Select

End Sub



Private Sub Combo1_Click()

Select Case Combo1.Text
    Case "圓形"
        顯示Lab.Caption = "半徑："
    Case Else
        顯示Lab.Caption = "邊長："
End Select

End Sub

