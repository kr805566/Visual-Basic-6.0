VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4185
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4815
   LinkTopic       =   "Form1"
   ScaleHeight     =   4185
   ScaleWidth      =   4815
   StartUpPosition =   3  '系統預設值
   Begin VB.Frame Frame1 
      Caption         =   "吉祥物"
      Height          =   1455
      Left            =   600
      TabIndex        =   3
      Top             =   2280
      Width           =   3615
      Begin VB.OptionButton Option6 
         Caption         =   "蛇"
         Height          =   375
         Left            =   2400
         TabIndex        =   9
         Top             =   960
         Width           =   1095
      End
      Begin VB.OptionButton Option5 
         Caption         =   "獅"
         Height          =   375
         Left            =   1320
         TabIndex        =   8
         Top             =   960
         Width           =   1095
      End
      Begin VB.OptionButton Option4 
         Caption         =   "象"
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   960
         Width           =   1095
      End
      Begin VB.OptionButton Option3 
         Caption         =   "熊"
         Height          =   375
         Left            =   2400
         TabIndex        =   6
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton Option2 
         Caption         =   "鯨"
         Height          =   375
         Left            =   1320
         TabIndex        =   5
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "牛"
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "查詢"
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   1080
      Width           =   1095
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
      Height          =   495
      Left            =   3120
      TabIndex        =   1
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "請輸入職棒名稱:"
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
      Top             =   240
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Select Case Text1.Text
Case "興農"
Option1.Value = True
Case "中信"
Option2.Value = True
Case "LaNew"
Option3.Value = True
Case "兄弟"
Option4.Value = True
Case "統一"
Option5.Value = True
Case "誠泰"
Option6.Value = True
Case Else
MsgBox "無此球隊", , "重新輸入"
End Select
End Sub
