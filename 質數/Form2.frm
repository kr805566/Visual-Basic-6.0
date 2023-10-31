VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "找出範圍內的質數"
   ClientHeight    =   3600
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5565
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   3600
   ScaleWidth      =   5565
   StartUpPosition =   3  '系統預設值
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  '垂直捲軸
      TabIndex        =   6
      Top             =   1560
      Width           =   3615
   End
   Begin VB.CommandButton Command3 
      Caption         =   "離開程式"
      Height          =   615
      Left            =   4080
      TabIndex        =   5
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "清除"
      Height          =   615
      Left            =   4080
      TabIndex        =   4
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "執行"
      Enabled         =   0   'False
      Height          =   615
      Left            =   4080
      TabIndex        =   3
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "回到主選單"
      Height          =   615
      Left            =   4080
      TabIndex        =   2
      Top             =   1920
      Width           =   1215
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
      Height          =   495
      Left            =   2640
      TabIndex        =   1
      Top             =   720
      Width           =   1335
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
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   2  '置中對齊
      Caption         =   "到"
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
      Left            =   1920
      TabIndex        =   8
      Top             =   720
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      Caption         =   "輸入任意兩數 "
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
      Left            =   1320
      TabIndex        =   7
      Top             =   240
      Width           =   2055
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim a, b, c As Long
Text3 = ""
If Val(Text1) > 2147483647 Then Text3 = "數字太大了 >_< ": Text1 = "": Text1.SetFocus: Exit Sub
If Val(Text2) > 2147483647 Then Text3 = "數字太大了 >_< ": Text2 = "": Text2.SetFocus: Exit Sub
a = Val(Text1)
b = Val(Text2)
c = 0
d = 0
If a > b Then c = a: a = b: b = c

e = a
f = b

If a < 2 Then a = 2
If b < 2 Then b = 2


For i = a To b
    k = 1
    For j = 2 To i ^ 0.5
    If i Mod j = 0 Then k = 0: Exit For
    
    Next j
   
   If k = 1 Then Text3 = Text3 & i & vbCrLf: d = d + 1

Next i


If Text3 = "" Then Text3 = e & "到" & f & "之間沒有質數": Exit Sub

Text3 = Text3 & e & "到" & f & "之間有" & d & "個質數"

End Sub

Private Sub Command2_Click()
Text1 = ""
Text2 = ""
Text3 = ""
Text1.SetFocus

End Sub

Private Sub Command3_Click()
End
End Sub

Private Sub Command4_Click()
Form2.Hide
Form1.Show

End Sub

Private Sub Form_Activate()
Text1.SetFocus

End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
End
End Sub



Private Sub Text1_Change()

If Text1 <> "" And Text2 <> "" Then Command1.Enabled = True
If Text1 = "" Then Command1.Enabled = False


End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)


If KeyAscii = 8 Then Exit Sub
If KeyAscii = 13 Then Text2.SetFocus
If KeyAscii < 48 Or KeyAscii > 58 Then KeyAscii = 0

End Sub
Private Sub Text2_Change()

If Text1 <> "" And Text2 <> "" Then Command1.Enabled = True
If Text1 = "" Then Command1.Enabled = False


End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)


If KeyAscii = 8 Then Exit Sub
If KeyAscii = 13 Then Command1 = True
If KeyAscii < 48 Or KeyAscii > 58 Then KeyAscii = 0

End Sub


Private Sub Text3_Change()

Text3.SelLength = Len(Text3)
End Sub
