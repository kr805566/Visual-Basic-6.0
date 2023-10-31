VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "判斷是否互質"
   ClientHeight    =   2490
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5910
   ForeColor       =   &H00000000&
   Icon            =   "Form4.frx":0000
   LinkTopic       =   "Form4"
   ScaleHeight     =   2490
   ScaleWidth      =   5910
   StartUpPosition =   3  '系統預設值
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
      Left            =   2400
      TabIndex        =   5
      Top             =   240
      Width           =   1335
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
      Left            =   4320
      TabIndex        =   4
      Top             =   240
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "回到主選單"
      Height          =   615
      Left            =   3120
      TabIndex        =   3
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "執行"
      Enabled         =   0   'False
      Height          =   615
      Left            =   240
      TabIndex        =   2
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "清除"
      Height          =   615
      Left            =   1680
      TabIndex        =   1
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "離開程式"
      Height          =   615
      Left            =   4560
      TabIndex        =   0
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   2  '置中對齊
      Caption         =   "m ="
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
      Left            =   3720
      TabIndex        =   8
      Top             =   240
      Width           =   495
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   15.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   360
      TabIndex        =   7
      Top             =   960
      Width           =   5055
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      Caption         =   "輸入任意兩數n ="
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
      Left            =   0
      TabIndex        =   6
      Top             =   240
      Width           =   2295
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()


Dim a, b As Long

If Val(Text1) > 2147483647 Or Val(Text1) < 1 Then Label2 = "請輸入0<n<=2147483647": Text1 = "": Text1.SetFocus: Exit Sub
If Val(Text2) > 2147483647 Or Val(Text2) < 1 Then Label2 = "請輸入0<m<=2147483647": Text2 = "": Text2.SetFocus: Exit Sub

a = Val(Text1)
b = Val(Text2)
c = 0

If a > b Then c = a: a = b: b = c

For i = 2 To b \ 2
   If a Mod i = 0 And b Mod i = 0 Then GoTo a
Next i


Label2 = a & "和" & b & "互質"

Exit Sub

a: Label2 = a & "和" & b & "公因數為" & i



End Sub


Private Sub Command2_Click()
Text1 = ""
Text2 = ""
Label2 = ""
Text1.SetFocus

End Sub

Private Sub Command3_Click()
End
End Sub

Private Sub Command4_Click()
Form4.Hide
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


