VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Plan"
   ClientHeight    =   6720
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10545
   BeginProperty Font 
      Name            =   "新細明體"
      Size            =   14.25
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   6720
   ScaleWidth      =   10545
   StartUpPosition =   3  '系統預設值
   Begin VB.TextBox Text2 
      Height          =   4095
      Left            =   5400
      MultiLine       =   -1  'True
      ScrollBars      =   2  '垂直捲軸
      TabIndex        =   7
      Text            =   "plan.frx":0000
      Top             =   1200
      Width           =   4815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "計畫讀取"
      Height          =   525
      Left            =   2280
      TabIndex        =   6
      Top             =   5880
      Width           =   1815
   End
   Begin VB.ComboBox Combo3 
      Height          =   405
      Left            =   3480
      TabIndex        =   5
      Text            =   "Combo1"
      Top             =   5400
      Width           =   1095
   End
   Begin VB.ComboBox Combo2 
      Height          =   405
      Left            =   1920
      TabIndex        =   4
      Text            =   "Combo1"
      Top             =   5400
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Height          =   405
      Left            =   360
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   5400
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   4095
      Left            =   360
      MultiLine       =   -1  'True
      ScrollBars      =   2  '垂直捲軸
      TabIndex        =   2
      Top             =   1200
      Width           =   4815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "計畫輸入"
      Height          =   525
      Left            =   360
      TabIndex        =   1
      Top             =   5880
      Width           =   1815
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   6840
      Top             =   720
   End
   Begin VB.Label Label1 
      Caption         =   "時間：2017/01/01"
      Height          =   615
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   4815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim R(100), RX As Integer
Private Sub Command1_Click()
 NewText = ""
RX = 0
Open App.Path & "\plan.jpg" For Output As #1
    Print #1, Format(Now, "yyyy/mm/dd")

        For i = 1 To Len(Combo1.Text)
    Y = Y & DateToWord(Mid(Combo1.Text, i, 1))
    Next i
    For i = 1 To Len(Combo2.Text)
    M = M & DateToWord(Mid(Combo2.Text, i, 1))
    Next i
    For i = 1 To Len(Combo3.Text)
    d = d & DateToWord(Mid(Combo3.Text, i, 1))
    Next i
    
    Print #1, Y & "," & M & "," & d
    
    For i = 1 To Len(Text1.Text)
        If Asc(Mid(Text1, i, 1)) = 13 Then
             NewText = NewText & vbCrLf
        Else
            If i > 1 Then
                If Asc(Mid(Text1, i - 1, 1)) <> 13 Then
                NewText = NewText & WordToNewWord(Mid(Text1, i, 1))
                RX = (RX + 1) Mod 100
                End If
            Else
                NewText = NewText & WordToNewWord(Mid(Text1, i, 1))
                RX = (RX + 1) Mod 100
            End If
        End If
    
    
    Next i
    Print #1, NewText
    
Close #1


End Sub

Private Sub Command2_Click()
Dim S, str, Y, M, d As String
str = ""
RX = 0

    Open App.Path & "\plan.jpg" For Input As #1
    Line Input #1, X
    str = "計畫開始日期:" & X & vbCrLf
        Input #1, Y1, M1, D1
    For i = 1 To Len(Y1)
    Y = Y & WordToDate(Mid(Y1, i, 1))
    Next i
    For i = 1 To Len(M1)
    M = M & WordToDate(Mid(M1, i, 1))
    Next i
    For i = 1 To Len(D1)
    d = d & WordToDate(Mid(D1, i, 1))
    Next i
    
    str = str & "計畫結算日期:" & Y & "/" & M & "/" & d & vbCrLf & vbCrLf
    
    Do Until EOF(1)
        Line Input #1, S
        For i = 1 To Len(S)
        
        str = str & NewWordToWord(Mid(S, i, 1))
        RX = (RX + 1) Mod 100
        Next
        
        str = str & vbCrLf
    Loop

    Close #1

    
    If val(Year(Now)) > val(Y) Or (val(Year(Now)) = val(Y) And val(Month(Now)) > val(M)) Or (val(Year(Now)) = val(Y) And val(Month(Now)) = val(M) And val(Day(Now)) >= val(d)) Then
        Text2.Text = str
    Else
        MsgBox "時間尚未到達" & Y & "/" & Format(M, "00") & "/" & Format(d, "00")
    End If
    
    
End Sub

Private Sub Form_Load()
Label1.Caption = "日期：" & Format(Now, "yyyy/mm/dd")

For i = 0 To 100
    Combo1.AddItem (i + Year(Now))
Next i

For i = 1 To 12
    Combo2.AddItem (Format(i, "00"))
Next i



Combo1.ListIndex = 0
Combo2.ListIndex = (Month(Now)) Mod 12
LeapYear
Combo3.ListIndex = 0
Dim Y, M, d As String


Open App.Path & "\plan.jpg" For Input As #1
        Line Input #1, X
        Input #1, Y1, M1, D1
       
        
    Close #1
    
     Z = Replace(X, "/", "")
 Z1 = (val(Mid(Z, 1, 2)) * 0.5 + val(Mid(Z, 3, 2)) + val(Mid(Z, 5, 2)) * 1.5 + val(Mid(Z, 7, 2)) * 2)
Randomize (val(Z1))

For i = 0 To 100
 R(i) = Int(Rnd() * 10 + 1)
Next i
    
    
    For i = 1 To Len(Y1)
    Y = Y & WordToDate(Mid(Y1, i, 1))
    Next i
    For i = 1 To Len(M1)
    M = M & WordToDate(Mid(M1, i, 1))
    Next i
    For i = 1 To Len(D1)
    d = d & WordToDate(Mid(D1, i, 1))
    Next i

 If val(Year(Now)) < val(Y) Or (val(Year(Now)) = val(Y) And val(Month(Now)) < val(M)) Or (val(Year(Now)) = val(Y) And val(Month(Now)) = val(M) And val(Day(Now)) < val(d)) Then
 Command1.Enabled = False
 End If

End Sub

Private Sub Timer1_Timer()
Label1.Caption = "日期：" & Format(Now, "yyyy/mm/dd")




End Sub





Function LeapYear()

Combo3.Clear
Select Case Combo2.ListIndex
    
    Case 0, 2, 4, 6, 7, 9, 11
        For i = 1 To 31
            Combo3.AddItem (Format(i, "00"))
        Next i

    Case 3, 5, 8, 10
        For i = 1 To 30
            Combo3.AddItem (Format(i, "00"))
        Next i
    Case 1
        
        For i = 1 To 28
            Combo3.AddItem (Format(i, "00"))
        Next i
        
        Y = Combo1.Text
        If (Y Mod 400 = 0) Or ((Y Mod 100 <> 0) And (Y Mod 4 = 0)) Then
            Combo3.AddItem (29)
        
        End If
           
        
End Select
Combo3.ListIndex = 0


End Function


Function WordToNewWord(word As String)

WordToNewWord = Chr(Asc(word) + R(RX))
End Function
Function NewWordToWord(word As String)
NewWordToWord = Chr(Asc(word) - R(RX))
End Function

Function DateToWord(word As String)
DateToWord = Chr(Asc(word) + 17 + R(0))
End Function

Function WordToDate(ByVal word As String)
WordToDate = Chr(Asc(word) - 17 - R(0))
End Function





