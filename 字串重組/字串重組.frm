VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "字串重組"
   ClientHeight    =   2280
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   3795
   LinkTopic       =   "Form1"
   ScaleHeight     =   2280
   ScaleWidth      =   3795
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command1 
      Caption         =   "重組"
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
      Left            =   2760
      TabIndex        =   2
      Top             =   600
      Width           =   855
   End
   Begin VB.TextBox Text2 
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
      Left            =   480
      TabIndex        =   1
      Top             =   960
      Width           =   1815
   End
   Begin VB.TextBox Text1 
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
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label Label1 
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
      Left            =   600
      TabIndex        =   3
      Top             =   1560
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim A(1 To 8, 1 To 2)

Private Sub Command1_Click()

For i = 1 To 8
   For j = 1 To 8
  
   If A(i, 1) = A(j, 2) Then
   Label1 = Label1 & A(i, 1)
   N = N + 1
   End If
   Next j
Next i

For i = 1 To 8 - N

Label1 = Label1 & A(i, 1)

Next i
End Sub

Private Sub Form_Load()

Text1.MaxLength = 8
Text2.MaxLength = 8
End Sub




Private Sub Text1_LostFocus()

If Text1 <> "" Then
For i = 1 To 8
A(i, 1) = Mid(Text1, i, 1)
Next i



For i = 1 To 7
   For j = i + 1 To 8
  
   If A(i, 1) = A(j, 1) Then Text1 = "": MsgBox "字不能重複", , "字串重組": Text1.SetFocus: Exit Sub
   
   Next j
Next i

End If

End Sub
Private Sub Text2_LostFocus()

If Text2 <> "" Then
For i = 1 To 8
A(i, 2) = Mid(Text2, i, 1)
Next i



For i = 1 To 7
   For j = i + 1 To 8
  
   If A(i, 2) = A(j, 2) Then Text2 = "": MsgBox "字不能重複", , "字串重組": Text2.SetFocus: Exit Sub
   
   Next j
Next i

End If

End Sub
