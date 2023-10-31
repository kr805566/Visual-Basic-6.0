VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "凌型星號"
   ClientHeight    =   4095
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   4095
   ScaleWidth      =   4680
   StartUpPosition =   3  '系統預設值
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Activate()
K = 12
For i = 1 To K
Print Spc(K - i); String(i * 2 - 1, "*")
Next i

For i = (K - 1) To 1 Step -1
Print Spc(K - i); String(i * 2 - 1, "*")
Next i

End Sub


