VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5805
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6975
   LinkTopic       =   "Form1"
   ScaleHeight     =   5805
   ScaleWidth      =   6975
   StartUpPosition =   3  '系統預設值
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()

Dim a(10000)

For i = 6 To 10000
k = 0
s = 0
    For j = 1 To i - 1

    If i Mod j = 0 Then
    k = k + 1
    a(k) = j
    s = s + a(k)
    End If
    
Next j
If i = s Then
    
    Print i & "=";
     For l = 1 To k
      Print a(l);
      
     Next l
     Print
End If
Next i



End Sub

