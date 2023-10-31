VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5025
   ClientLeft      =   60
   ClientTop       =   435
   ClientWidth     =   7185
   LinkTopic       =   "Form1"
   ScaleHeight     =   5025
   ScaleWidth      =   7185
   StartUpPosition =   3  '系統預設值
   Begin MSFlexGridLib.MSFlexGrid a1 
      Height          =   1815
      Left            =   840
      TabIndex        =   9
      Top             =   2760
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   3201
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "應檢人資料"
      Height          =   1455
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   5415
      Begin VB.TextBox Text4 
         Height          =   270
         Left            =   3600
         TabIndex        =   8
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox Text3 
         Height          =   270
         Left            =   3600
         TabIndex        =   6
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         Height          =   270
         Left            =   960
         TabIndex        =   4
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Left            =   960
         TabIndex        =   2
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "考試日期"
         Height          =   255
         Left            =   2520
         TabIndex        =   7
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "准考證號碼"
         Height          =   255
         Left            =   2400
         TabIndex        =   5
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "座　號"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "姓　名"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   615
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub a_Click()

End Sub

Private Sub Form_Activate()

a1.Cols = 5
a1.Row = 0
a1.Col = 1: a1.Text = "VALUE1"
a1.Col = 2: a1.Text = " OP "
a1.Col = 3: a1.Text = "VALUE2"
a1.Col = 4: a1.Text = "ANSWER"

Open App.Path & "\940308.sm" For Input As #1
Do Until EOF(1)
Input #1, b, A, p, y, x

If p = "+" Then ans1 = (b * x + A * y): ans2 = (A * x)
If p = "-" Then ans1 = (b * x - A * y): ans2 = (A * x)
If p = "*" Then ans1 = (b * y): ans2 = (A * x)
If p = "/" Then ans1 = (b * x): ans2 = (A * y)


ans11 = ans1
ans22 = ans2

If ans22 > ans11 Then t = ans22: ans22 = ans11: ans11 = t
If ans22 = 0 Then ans22 = 1

Do Until ans11 Mod ans22 = 0

c = ans22
t = ans11
ans11 = ans22
ans22 = t Mod ans22




Loop

a1.Row = a1.Rows - 1
a1.Col = 1: a1.Text = " " & b & "/" & A
a1.Col = 2: a1.Text = " " & p
a1.Col = 3: a1.Text = " " & y & "/" & x
If ans1 Mod ans2 = 0 Then
a1.Col = 4: a1.Text = " " & ans1 / ans2
Else
a1.Col = 4: a1.Text = " " & ans1 / ans22 & "/" & ans2 / ans22
End If
a1.Rows = a1.Rows + 1
Loop

a1.Rows = a1.Rows - 1
Close #1

End Sub

