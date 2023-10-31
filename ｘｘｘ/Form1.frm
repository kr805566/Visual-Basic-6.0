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
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()

Form1.Hide

c = MsgBox("結束程式嗎?", vbYesNo + vbQuestion, "警告")
If c = 6 Then
MsgBox "請存檔在離開", , "結束程式"

End

Else

MsgBox "回到主畫面", , "返回程式"

Shell ("C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE")

End If

End Sub
