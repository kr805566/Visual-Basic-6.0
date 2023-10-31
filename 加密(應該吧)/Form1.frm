VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4050
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4710
   LinkTopic       =   "Form1"
   ScaleHeight     =   4050
   ScaleWidth      =   4710
   StartUpPosition =   3  '系統預設值
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   3615
   End
   Begin VB.Label Label1 
      Appearance      =   0  '平面
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   480
      TabIndex        =   1
      Top             =   1320
      Width           =   3615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a(1 To 100)
Dim F

Private Sub Text1_KeyPress(KeyAscii As Integer)
Label1 = ""


If KeyAscii = 13 Then

For F = 1 To Len(Text1)
If Asc(Mid(Text1, F, 1)) < 89 Then
 a(F) = Chr(Asc(Mid(Text1, F, 1)) + 2)
 Else
 a(F) = Mid(Text1, F, 1)
 Select Case a(F)
 
       Case "Y"
       a(F) = "A"
       Case "Z"
       a(F) = "B"
 End Select
End If
 
 Call b
 
 Call C
 
Next F
For I = 1 To Len(Text1)

   Label1 = Label1 & a(I)
    
Next I



End If

End Sub


Sub b()
Select Case a(F)

       Case "A"
       a(F) = "K"
       Case "Z"
       a(F) = "E"
       Case "C"
       a(F) = "H"
       Case "S"
       a(F) = "U"
       Case "R"
       a(F) = "V"
       Case "K"
       a(F) = "N"
       Case "P"
       a(F) = "T"
       Case "B"
       a(F) = "C"
      
       
       End Select
End Sub
Sub C()


Select Case a(F)



Case "A"
a(F) = "a"
Case "E"
a(F) = "e"
Case "I"
a(F) = "i"
Case "O"
a(F) = "o"
Case "U"
a(F) = "u"
Case "J"
a(F) = 1
Case "Q"
a(F) = 2
Case "K"
a(F) = 3
Case "X"
a(F) = "?"
Case "Y"
a(F) = "?"
Case "Z"
a(F) = "?"




End Select




End Sub
