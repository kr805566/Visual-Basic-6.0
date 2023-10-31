VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "判斷質數總匯"
   ClientHeight    =   7830
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7515
   LinkTopic       =   "Form1"
   ScaleHeight     =   7830
   ScaleWidth      =   7515
   StartUpPosition =   3  '系統預設值
   Begin VB.Frame Frame5 
      Caption         =   "DO     LOOP  UNTIL(後測試) "
      Height          =   2055
      Left            =   3960
      TabIndex        =   15
      Top             =   3240
      Width           =   3375
      Begin VB.CommandButton DO_LOOP_UNTIL 
         Caption         =   "DO              LOOP UNTIL"
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
         Left            =   600
         TabIndex        =   16
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label6 
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
         Left            =   720
         TabIndex        =   17
         Top             =   1320
         Width           =   1935
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "DO     LOOP  WHILE (後測試) "
      Height          =   2055
      Left            =   360
      TabIndex        =   12
      Top             =   3240
      Width           =   3375
      Begin VB.CommandButton DO_LOOP_WHILE 
         Caption         =   "DO               LOOP WHILE"
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
         Left            =   600
         TabIndex        =   13
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label5 
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
         Left            =   720
         TabIndex        =   14
         Top             =   1320
         Width           =   1815
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "DO  UNTIL      LOOP (前測試) "
      Height          =   2055
      Left            =   3960
      TabIndex        =   9
      Top             =   960
      Width           =   3375
      Begin VB.CommandButton DO_UNTIL_LOOP 
         Caption         =   "DO   UNTIL"
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
         Left            =   600
         TabIndex        =   10
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label4 
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
         Left            =   720
         TabIndex        =   11
         Top             =   1320
         Width           =   1815
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "DO  WHILE --  LOOP (前測試) "
      Height          =   2055
      Left            =   360
      TabIndex        =   6
      Top             =   960
      Width           =   3375
      Begin VB.CommandButton DO_WHILE_LOOP 
         Caption         =   "DO WHILE"
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
         Left            =   600
         TabIndex        =   7
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label3 
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
         Left            =   720
         TabIndex        =   8
         Top             =   1320
         Width           =   1695
      End
   End
   Begin VB.CommandButton END 
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
      Height          =   495
      Left            =   4440
      TabIndex        =   3
      Top             =   120
      Width           =   735
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
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   3000
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "WHILE  WEND"
      Height          =   2055
      Left            =   2280
      TabIndex        =   0
      Top             =   5520
      Width           =   3375
      Begin VB.CommandButton WHILE_WEND 
         Caption         =   "WHILE   WEND"
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
         Left            =   600
         TabIndex        =   4
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label2 
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
         Left            =   720
         TabIndex        =   5
         Top             =   1320
         Width           =   2055
      End
   End
   Begin VB.Label Label1 
      Caption         =   "輸入欲判斷的數值"
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
      Left            =   600
      TabIndex        =   1
      Top             =   240
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As Integer


Private Sub END_Click()
End
End Sub

Private Sub WHILE_WEND_Click()

a = Text1
m = "是質數"
i = 2
While i <= a ^ 0.5

If a Mod i = 0 Then
m = "不是質數"
GoTo k
End If
i = i + 1
Wend
k:
Label2 = a & m




End Sub
Private Sub DO_WHILE_LOOP_Click()

a = Text1
m = "是質數"
i = 2
Do While i <= a ^ 0.5

If a Mod i = 0 Then
m = "不是質數"
Exit Do
End If
i = i + 1
Loop

Label3 = a & m


End Sub
Private Sub DO_UNTIL_LOOP_Click()
a = Text1
m = "是質數"
i = 2
Do Until i > a ^ 0.5

If a Mod i = 0 Then
m = "不是質數"
Exit Do
End If
i = i + 1
Loop

Label4 = a & m
End Sub
Private Sub DO_LOOP_WHILE_Click()


a = Text1
m = "是質數"
i = 2
Do

If a Mod i = 0 Then
m = "不是質數"
Exit Do
End If
i = i + 1
Loop While i <= a ^ 0.5

Label5 = a & m


End Sub


Private Sub DO_LOOP_UNTIL_Click()

a = Text1
m = "是質數"
i = 2
Do

If a Mod i = 0 Then
m = "不是質數"
Exit Do
End If
i = i + 1
Loop Until i > a ^ 0.5

Label6 = a & m

End Sub
