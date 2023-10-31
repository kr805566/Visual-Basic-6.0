VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5805
   ClientLeft      =   60
   ClientTop       =   435
   ClientWidth     =   5445
   LinkTopic       =   "Form1"
   ScaleHeight     =   5805
   ScaleWidth      =   5445
   StartUpPosition =   3  '系統預設值
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Open App.Path & "/In.txt" For Input As #1
'Open App.Path & "/Out.txt" For Output As #2

Dim m()
Dim ss(99), xx

Function a(l, n)
If Len(l) = Len(n) Then
For k = 1 To Len(n)
su1 = su1 & m(Val(Mid(n, k, 1)))
Next k
'Print su1
st = True
For i = 0 To xx
    If ss(i) = su1 Then st = False
Next i

If st Then
    ss(xx) = su1
    xx = xx + 1
End If

Else
    For i = 1 To Len(l)
        s = True
        For j = 1 To Len(n)
        If Mid(l, i, 1) = Mid(n, j, 1) Then s = False
        
        Next j
    If s = True Then Call a(l, n & Mid(l, i, 1))
    Next i
End If


End Function


Private Sub Form_Activate()
  '  Do While Not EOF(1)
   '     Input #1, Inp
'
  '  Loop
    '    Print #2, Out
    '    Close
      '  End
b = "123321"
      
      
ReDim m(Len(b))

For i = 1 To Len(b)
    m(i) = Mid(b, i, 1)
    su = su & i
Next i


Call a(su, "")

For i = 0 To xx
If Len(ss(i)) Mod 2 = 0 Then
    s = True
    For j = 1 To Len(ss(i)) / 2
    If Mid(ss(i), j, 1) <> Mid(ss(i), Len(ss(i)) - j + 1, 1) Then s = False
    Next j
If s Then Print ss(i)
Else

   s = True
    For j = 1 To (Len(ss(i)) - 1) / 2
    If Mid(ss(i), j, 1) <> Mid(ss(i), Len(ss(i)) - j + 1, 1) Then s = False
    Next j
If s Then Print ss(i)





End If
Next i

End Sub

