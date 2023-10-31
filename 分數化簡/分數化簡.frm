VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "分數化簡"
   ClientHeight    =   2775
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   ScaleHeight     =   2775
   ScaleWidth      =   6030
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command2 
      Caption         =   "清除"
      Height          =   495
      Left            =   4800
      TabIndex        =   10
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "輸入"
      Height          =   495
      Left            =   4800
      TabIndex        =   8
      Top             =   480
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  '置中對齊
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  '置中對齊
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label7 
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
      Left            =   960
      TabIndex        =   9
      Top             =   1920
      Width           =   3255
   End
   Begin VB.Label Label6 
      Caption         =   "____________"
      Height          =   375
      Left            =   3360
      TabIndex        =   7
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label5 
      Alignment       =   2  '置中對齊
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      TabIndex        =   6
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label4 
      Alignment       =   2  '置中對齊
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      TabIndex        =   5
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   2  '置中對齊
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   4
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label2 
      Alignment       =   2  '置中對齊
      Caption         =   "="
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
      Left            =   1560
      TabIndex        =   3
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "____________"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()



Dim a, b, c, d, e, f As Integer

Label7 = ""

a = Val(Text1)
b = Val(Text2)



For i = a To 1 Step -1
   If a Mod i = 0 And b Mod i = 0 Then Exit For
Next i
       
       
c = a / i
d = b / i


If c > d Then
 e = c \ d
f = c Mod d
 Label3 = e
Label4 = f
Label5 = d

Else

 Label3 = ""
 Label4 = c
 Label5 = d


End If

If i = 1 Then Label7 = a & "和" & b & "互質"




End Sub
Private Sub Command2_Click()

Text1 = ""
Text2 = ""




Label3 = ""
Label4 = ""
Label5 = ""
Label7 = ""
End Sub


