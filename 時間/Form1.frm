VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1920
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   1920
   ScaleWidth      =   4560
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command1 
      Caption         =   "顯示時間"
      Height          =   615
      Left            =   2880
      TabIndex        =   0
      Top             =   1200
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

Cls

Print
Print
Print

Print "　　　　　今日日期: " & Date

Print

Print "　　　　　現在時間: " & Time

End Sub
