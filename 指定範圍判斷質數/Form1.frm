VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "指定範圍判斷質數"
   ClientHeight    =   6585
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7590
   LinkTopic       =   "Form1"
   ScaleHeight     =   6585
   ScaleWidth      =   7590
   StartUpPosition =   3  '系統預設值
   Begin VB.ListBox List1 
      Height          =   5820
      ItemData        =   "Form1.frx":0000
      Left            =   4080
      List            =   "Form1.frx":0002
      TabIndex        =   7
      Top             =   360
      Width           =   3255
   End
   Begin VB.CommandButton Command3 
      Caption         =   "清除"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   18
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2280
      TabIndex        =   6
      Top             =   4320
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "結束"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   18
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   5
      Top             =   4320
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "開始"
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
      Left            =   1080
      TabIndex        =   4
      Top             =   3120
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   1920
      TabIndex        =   3
      Top             =   1320
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   1920
      TabIndex        =   2
      Top             =   360
      Width           =   1935
   End
   Begin VB.Label Label2 
      Alignment       =   2  '置中對齊
      Caption         =   "終止值"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   18
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  '置中對齊
      Caption         =   "起始值"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   18
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim s, o, c, k As Single
k = 1
c = 0

s = Val(Text1)
o = Val(Text2)
If s = 1 Then s = s + 1
For i = s To o
   For j = 2 To i / 2
   If i Mod j = 0 Then k = 0: Exit For
   Next j
If k = 1 Then
List1.AddItem i
c = c + 1
End If
k = 1
Next i
x = "共有" & c & "個質數"
List1.AddItem x
End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Command3_Click()
Text1 = ""
Text2 = ""
y = ""
List1.AddItem y
End Sub
