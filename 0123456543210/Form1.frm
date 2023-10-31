VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7035
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   7035
   ScaleWidth      =   6585
   StartUpPosition =   3  '系統預設值
   Begin VB.TextBox Text1 
      Height          =   5295
      Left            =   480
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "Form1.frx":0000
      Top             =   360
      Width           =   4815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
n = 9
Text1 = ""
For i = 0 To n
    Text1 = Text1 & Space((n - i) * 2)
    For j = 0 To i
    Text1 = Text1 & j
    Next j
    
    For j = i - 1 To 0 Step -1
    Text1 = Text1 & j
    Next j
    Text1 = Text1 & vbCrLf
Next i




For i = n - 1 To 0 Step -1
    Text1 = Text1 & Space((n - i) * 2)
    For j = 0 To i
    Text1 = Text1 & j
    Next j
    
    For j = i - 1 To 0 Step -1
    Text1 = Text1 & j
    Next j
    Text1 = Text1 & vbCrLf
Next i



End Sub

