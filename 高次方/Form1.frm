VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2625
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5070
   LinkTopic       =   "Form1"
   ScaleHeight     =   2625
   ScaleWidth      =   5070
   StartUpPosition =   3  '╰参箇砞
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   1320
      Width           =   4455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "块"
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   2160
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Text            =   "70"
      Top             =   720
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Text            =   "2"
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label Label4 
      Caption         =   "m= 1~99"
      BeginProperty Font 
         Name            =   "穝灿砰"
         Size            =   24
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      TabIndex        =   7
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "n= 1~9"
      BeginProperty Font 
         Name            =   "穝灿砰"
         Size            =   24
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      TabIndex        =   6
      Top             =   0
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "m"
      BeginProperty Font 
         Name            =   "穝灿砰"
         Size            =   27.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2160
      TabIndex        =   5
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "n"
      BeginProperty Font 
         Name            =   "穝灿砰"
         Size            =   48
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1800
      TabIndex        =   4
      Top             =   1680
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim A, B, C  'A2A1 * B2B1 = C4C3C2C1
Private Sub Command1_Click()
 Call 蔼Ωよ(Text1, Val(Text2), 1) 'text1┏计  text2计  1﹍
End Sub

Function 蔼Ωよ(┏计, 计, 挡狦)
If 计 = 0 Then Text3 = 挡狦: Exit Function
Call 蔼Ωよ(┏计, 计 - 1, (┏计, 挡狦)) '计硋Ω搭1 0MΩ

End Function

Function (A计, B计)
A计计 = Len(A计) '衡计碭计
B计计 = Len(B计)
ReDim A(A计计), B(B计计), C(A计计 + B计计) 'A B   C计程单A计+B计

For I = 1 To A计计
    A(I) = Mid(A计, A计计 - I + 1, 1) ' –计秈 ㄒ:314 A(3)=3 A(2)=1 A(1)=4
Next I

For I = 1 To B计计
    B(I) = Mid(B计, B计计 - I + 1, 1)
Next I

For I = 1 To A计计
    For j = 1 To B计计                       '   A2   A1           ㄒ   1   2
                                               '   B2   B1                  2   9
    C(j + I - 1) = C(j + I - 1) + A(I) * B(j)  'X--------                X------
                                               '   C2(1) C1                 9  18
    Next j                                     'C3 C2(2)                 2  4
Next I                                         '----------              ------------
                                               'C3 C2    C1              2 13  18
For I = 1 To A计计 + B计计 - 1             '                        ------------
 C(I + 1) = C(I + 1) + C(I) \ 10               '耞琌禬筁9          2 13  8
C(I) = C(I) Mod 10                             '秈                     +1
Next I                                         '                        ------------
                                               '                         2 14 8
                                               '                        ------------
                                               '                         2  4  8
                                               '                        +1
                                               '                        -------------
                                               '                          3  4  8
                                               
                                               
If C(A计计 + B计计) = 0 Then C(A计计 + B计计) = ""   ' 硂耞琌ňゎ玡瞷0
                                                             'ㄒ: 11 * 12 = 0132

For I = 1 To A计计 + B计计

 = C(I) &                                             '盢舱
                                                              'ㄒ C(3)=2  C(2)=5  C(1)=6   =256

Next I




End Function


