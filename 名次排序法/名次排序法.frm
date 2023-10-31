VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "排列名次(固定5人)"
   ClientHeight    =   4470
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7185
   LinkTopic       =   "Form1"
   ScaleHeight     =   4470
   ScaleWidth      =   7185
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton 輸入分數 
      Caption         =   "1.輸入分數(5科)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4920
      TabIndex        =   6
      Top             =   1080
      Width           =   1575
   End
   Begin VB.CommandButton END 
      Caption         =   "結束"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   15.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4920
      TabIndex        =   5
      Top             =   3000
      Width           =   1455
   End
   Begin VB.CommandButton 排名次 
      Caption         =   "2.排列名次"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4920
      TabIndex        =   4
      Top             =   2040
      Width           =   1575
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   15.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1635
      ItemData        =   "名次排序法.frx":0000
      Left            =   480
      List            =   "名次排序法.frx":0002
      TabIndex        =   3
      Top             =   1200
      Width           =   1095
   End
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   15.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1635
      ItemData        =   "名次排序法.frx":0004
      Left            =   3000
      List            =   "名次排序法.frx":0006
      TabIndex        =   2
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "名次陣列"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   15.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      TabIndex        =   1
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "分數陣列"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   15.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a(1 To 5), b(1 To 5), C(1 To 5) As Integer

Private Sub END_Click()
End
End Sub

Private Sub 排名次_Click()

For I = 1 To 4
    For J = I + 1 To 5
   If C(I) < C(J) Then
  D = C(I): C(I) = C(J): C(J) = D
   End If

Next J, I


For I = 1 To 5
    For J = 1 To 5
If C(I) = a(J) Then
   b(J) = I

End If

Next J, I


For I = 1 To 5
List2.AddItem b(I)

Next I


End Sub

Private Sub 輸入分數_Click()
For I = 1 To 5
b(I) = 1
a(I) = Val(InputBox("輸入第" & I & "成績", "輸入成績"))
C(I) = a(I)
List1.AddItem a(I)

Next I

End Sub
