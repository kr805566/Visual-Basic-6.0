VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2550
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5010
   LinkTopic       =   "Form1"
   ScaleHeight     =   2550
   ScaleWidth      =   5010
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command1 
      Caption         =   "加密"
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   1080
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   1680
      Width           =   4695
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   4695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim Z(1 To 9)
For L = 1 To Len(Text1)
K = Mid(Text1, L, 1)



Select Case K
       
       Case "a"
          
          Select Case Z(1)
          
                 Case 0
                 K = " 09"
                 Case 1
                 K = " 12"
                 Case 2
                 K = " 33"
                 Case 3
                 K = " 47"
                 Case 4
                 K = " 53"
                 Case 5
                 K = " 67"
                 Case 6
                 K = " 78"
                 Case 7
                 K = " 92"
                 Z(1) = -1
          End Select
          Z(1) = Z(1) + 1
       
       Case "b"
          Select Case Z(2)
          
                 Case 0
                 K = " 48"
                 Case 1
                 K = " 81"
                 Z(2) = -1
          End Select
          Z(2) = Z(2) + 1
       
       Case "c"
        Select Case Z(3)
          
                 Case 0
                 K = " 13"
                 Case 1
                 K = " 41"
                 Case 2
                 K = " 62"
                 Z(3) = -1
          End Select
          Z(3) = Z(3) + 1
       
       Case "d"
            Select Case Z(4)
          
                 Case 0
                 K = " 01"
                 Case 1
                 K = " 03"
                 Case 2
                 K = " 45"
                 Case 3
                 K = " 79"
                 Z(4) = -1
          End Select
          Z(4) = Z(4) + 1
       
       
       Case "e"
        Select Case Z(5)
          
                 Case 0
                 K = " 14"
                 Case 1
                 K = " 16"
                 Case 2
                 K = " 24"
                 Case 3
                 K = " 44"
                 Case 4
                 K = " 46"
                 Case 5
                 K = " 55"
                 Case 6
                 K = " 57"
                 Case 7
                 K = " 64"
                 Case 7
                 K = " 74"
                 Case 7
                 K = " 82"
                 Case 7
                 K = " 87"
                 Case 7
                 K = " 98"
                 Z(5) = -1
          End Select
          Z(5) = Z(5) + 1
       Case "f"
         Select Case Z(6)
          
                 Case 0
                 K = " 10"
                 Case 1
                 K = " 31"
                 Z(6) = -1
          End Select
          Z(6) = Z(6) + 1
       
       Case "g"
         Select Case Z(7)
          
                 Case 0
                 K = " 06"
                 Case 1
                 K = " 25"
                 Z(7) = -1
          End Select
          Z(7) = Z(7) + 1

       Case "h"
          Select Case Z(8)
          
                 Case 0
                 K = " 23"
                 Case 1
                 K = " 39"
                 Case 2
                 K = " 50"
                 Case 3
                 K = " 56"
                 Case 4
                 K = " 65"
                 Case 5
                 K = " 68"
                
                 Z(8) = -1
          End Select
          Z(8) = Z(8) + 1
       
       
       Case "i"
               Select Case Z(9)
          
                 Case 0
                 K = " 32"
                 Case 1
                 K = " 70"
                 Case 2
                 K = " 73"
                 Case 3
                 K = " 83"
                 Case 4
                 K = " 88"
                 Case 5
                 K = " 93"
                
                 Z(9) = -1
          End Select
          Z(9) = Z(9) + 1
       
       Case "j"
        K = 15



End Select


X = X & K

Next L

Text2 = X
End Sub
