VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "真靈運勢占卜"
   ClientHeight    =   2010
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4785
   LinkTopic       =   "Form1"
   ScaleHeight     =   2010
   ScaleWidth      =   4785
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command2 
      Caption         =   "結束"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3840
      TabIndex        =   3
      Top             =   1080
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "抽籤"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3840
      TabIndex        =   2
      Top             =   240
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Caption         =   "占卜結果"
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   3495
      Begin VB.TextBox Text1 
         CausesValidation=   0   'False
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   14.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   240
         Width           =   3015
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Randomize

a = 1 + Int(Rnd * 6) * 1

Select Case a
Case 1
Text1 = "下下籤:" & vbCrLf & "諸事不宜"
Case 2
Text1 = "中上籤:" & vbCrLf & "會有貴人相助"
Case 3
Text1 = "中下籤:" & vbCrLf & "口舌之爭"
Case 4
Text1 = "平籤:" & vbCrLf & "保持平常心"
Case 5
Text1 = "平籤:" & vbCrLf & "沒事就是好事"
Case 6
Text1 = "上上籤:" & vbCrLf & "意外之財降臨"
End Select
       
       
End Sub


Private Sub Command2_Click()

End

End Sub
