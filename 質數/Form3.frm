VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "判斷是否為質數"
   ClientHeight    =   3270
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   3495
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   3270
   ScaleWidth      =   3495
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
      Left            =   1920
      TabIndex        =   4
      Top             =   240
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "執行"
      Enabled         =   0   'False
      Height          =   615
      Left            =   240
      TabIndex        =   3
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "清除"
      Height          =   615
      Left            =   240
      TabIndex        =   2
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "離開程式"
      Height          =   615
      Left            =   1800
      TabIndex        =   1
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "回到主選單"
      Height          =   615
      Left            =   1800
      TabIndex        =   0
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "輸入任意數 n  ="
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
      Left            =   120
      TabIndex        =   6
      Top             =   360
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
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
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   3015
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim a As Long

If Val(Text1) > 2147483647 Then Label2 = "數字太大了 >_< 請輸入n<=2147483647": Text1 = "": Text1.SetFocus: Exit Sub
a = Val(Text1)

If a < 2 Then Label2 = a & " 不為質數": Exit Sub

For i = 2 To a ^ 0.5
If a Mod i = 0 Then GoTo a
Next i
Label2 = a & "為質數"
Exit Sub
a: Label2 = a & " 不為質數可以被 " & i & " 整除"



End Sub

Private Sub Command2_Click()
Text1 = ""
Label2 = ""
Text1.SetFocus

End Sub

Private Sub Command3_Click()
End
End Sub

Private Sub Command4_Click()

Form3.Hide
Form1.Show


End Sub

Private Sub Form_Activate()
Text1.SetFocus

End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
End
End Sub

Private Sub Text1_Change()

If Text1 <> "" Then Command1.Enabled = True
If Text1 = "" Then Command1.Enabled = False


End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)


If KeyAscii = 8 Then Exit Sub
If KeyAscii = 13 Then Command1 = True
If KeyAscii < 48 Or KeyAscii > 58 Then KeyAscii = 0

End Sub

