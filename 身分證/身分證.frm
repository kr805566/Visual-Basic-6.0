VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3375
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4920
   LinkTopic       =   "Form1"
   ScaleHeight     =   3375
   ScaleWidth      =   4920
   StartUpPosition =   3  '系統預設值
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
Dim b(1 To 11) As Integer
Form1.Hide
A = InputBox("請輸入身分字號", "身分字號")

Select Case Mid(A, 1, 1)
Case "B"
b(1) = 1
b(11) = 1
Case "L"
b(1) = 2
b(11) = 0
Case "M"
b(1) = 2
b(11) = 1
Case "N"
b(1) = 22
b(11) = 2


End Select
For i = 2 To 10
b(i) = Mid(A, i, 1)
Next i

If 10 - ((b(1) + b(11) * 9 + b(2) * 8 + b(3) * 7 + b(4) * 6 + b(5) * 5 + b(6) * 4 + b(7) * 3 + b(8) * 2 + b(9)) Mod 10) - b(10) = 0 Then
MsgBox "身分字號 " & A & " 為正確", , "身分字號"
Else
MsgBox "身分字號 " & A & " 為錯誤", , "身分字號"
End If





End Sub

