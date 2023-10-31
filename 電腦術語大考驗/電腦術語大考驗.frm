VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "電腦術語大考驗"
   ClientHeight    =   2880
   ClientLeft      =   5385
   ClientTop       =   4035
   ClientWidth     =   6975
   LinkTopic       =   "Form1"
   ScaleHeight     =   2880
   ScaleWidth      =   6975
   Begin VB.CommandButton Command2 
      Caption         =   "重新"
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
      Left            =   5760
      TabIndex        =   5
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "解答"
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
      Left            =   4440
      TabIndex        =   4
      Top             =   2040
      Width           =   1095
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
      ItemData        =   "電腦術語大考驗.frx":0000
      Left            =   4560
      List            =   "電腦術語大考驗.frx":0002
      TabIndex        =   3
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label Label3 
      Alignment       =   2  '置中對齊
      BackColor       =   &H00C0C0FF&
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
      Left            =   240
      TabIndex        =   2
      Top             =   1920
      Width           =   3615
   End
   Begin VB.Label Label2 
      Caption         =   "題目:"
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
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   3975
   End
   Begin VB.Label Label1 
      Caption         =   "請點選相對應的電腦術語:"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   15.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   3615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim n As Integer
Dim a(1 To 5) As String
Dim b(1 To 5) As String
Private Sub Command1_Click()


If List1 = b(n) Then
Label3 = "答對了"
Else
Label3 = "答錯了,答案是" & b(n)
End If

End Sub

Private Sub Command2_Click()

Label2 = "題目:"
Label3 = ""
List1.Clear
Call Form_Activate

End Sub

Private Sub Form_Activate()

Dim i As Integer


a(1) = "電腦輔助教學軟體"
a(2) = "中央處理單元"
a(3) = "隨機存取記憶體"
a(4) = "全球資訊網"
a(5) = "區域網路"

b(1) = "CAI"
b(2) = "CPU"
b(3) = "RAM"
b(4) = "WWW"
b(5) = "LAN"

Randomize

n = Int(Rnd() * 5) + 1

Label2 = Label2 & a(n) & "=>"

For i = 1 To 5

List1.AddItem b(i)

Next i

End Sub


