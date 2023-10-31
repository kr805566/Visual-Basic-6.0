VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "排序"
   ClientHeight    =   5865
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8625
   LinkTopic       =   "Form1"
   ScaleHeight     =   5865
   ScaleWidth      =   8625
   StartUpPosition =   3  '系統預設值
   Begin VB.TextBox Text1 
      Alignment       =   2  '置中對齊
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   18
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3000
      TabIndex        =   9
      Top             =   240
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "結束"
      BeginProperty Font 
         Name            =   "超研澤中行書"
         Size            =   18
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3240
      TabIndex        =   5
      Top             =   4200
      Width           =   1575
   End
   Begin VB.Frame Frame2 
      Caption         =   "排序後"
      Height          =   3975
      Left            =   5160
      TabIndex        =   4
      Top             =   1440
      Width           =   2535
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
         Height          =   2895
         ItemData        =   "Form1.frx":0000
         Left            =   240
         List            =   "Form1.frx":0002
         TabIndex        =   7
         Top             =   480
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "排序前"
      Height          =   3975
      Left            =   240
      TabIndex        =   3
      Top             =   1440
      Width           =   2415
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
         Height          =   2895
         ItemData        =   "Form1.frx":0004
         Left            =   240
         List            =   "Form1.frx":0006
         TabIndex        =   6
         Top             =   480
         Width           =   1935
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "輸入分數"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   18
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5280
      TabIndex        =   2
      Top             =   120
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "遞減"
      BeginProperty Font 
         Name            =   "超研澤中行書"
         Size            =   18
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3240
      TabIndex        =   1
      Top             =   2640
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "遞增"
      BeginProperty Font 
         Name            =   "超研澤中行書"
         Size            =   18
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3240
      TabIndex        =   0
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  '置中對齊
      Caption         =   "輸入人數:"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   18
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   8
      Top             =   480
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a() As Integer
Dim n As Integer

Private Sub Command1_Click()

For i = 1 To n - 1
  For j = 0 To n - 1 - i

If a(j) > a(j + 1) Then

b = a(j)
a(j) = a(j + 1)
a(j + 1) = b
End If
Next j, i



List2.Clear

For i = 0 To n - 1

List2.AddItem a(i)
Next i


End Sub

Private Sub Command2_Click()

For i = 0 To n - 2
  For j = i + 1 To n - 1

If a(i) < a(j) Then

b = a(i)
a(i) = a(j)
a(j) = b
End If
Next j, i



List2.Clear

For i = 0 To n - 1

List2.AddItem a(i)
Next i




End Sub


Private Sub Command3_Click()
n = Val(Text1)
ReDim a(n)
For i = 0 To n - 1
a(i) = Val(InputBox("請輸入第" & i + 1 & "號分數", "排序"))
List1.AddItem a(i)
Next i
Command2.Enabled = True
Command1.Enabled = True
Command3.Enabled = False
End Sub

Private Sub Command4_Click()
End
End Sub

Private Sub Form_Activate()

Text1.SetFocus

Command1.Enabled = False

Command2.Enabled = False
End Sub


