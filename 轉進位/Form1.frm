VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command2 
      Caption         =   "8轉10"
      Height          =   735
      Left            =   960
      TabIndex        =   1
      Top             =   1320
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "16轉10"
      Height          =   735
      Left            =   960
      TabIndex        =   0
      Top             =   360
      Width           =   2655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

a = InputBox("請輸入16進位數", "轉10進位")
MsgBox "這個16進位數字轉成10進位數是 " & Val("&h" & a), , "16轉10"


End Sub

Private Sub Command2_Click()

a = InputBox("請輸入8進位數", "轉10進位")
MsgBox "這個8進位數字轉成10進位數是 " & Val("&O" & a), , "8轉10"

End Sub
