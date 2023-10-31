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
Dim y As Integer
y = InputBox("請輸入出生年次", "輸入年次")
If y <= 59 Then
MsgBox "芭樂族", , "世代族群"
ElseIf y >= 60 And y <= 69 Then
MsgBox "草莓族", , "世代族群"
Else
MsgBox "水蜜桃族", , "世代族群"
End If

End







End Sub
