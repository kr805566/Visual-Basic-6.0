VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "巴斯卡(Pascal)三角形"
   ClientHeight    =   4725
   ClientLeft      =   3780
   ClientTop       =   3210
   ClientWidth     =   7275
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   4725
   ScaleWidth      =   7275
   Begin VB.CommandButton cmdCal 
      Caption         =   "計算"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6000
      TabIndex        =   2
      Top             =   4080
      Width           =   1095
   End
   Begin VB.TextBox txtN 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6000
      TabIndex        =   1
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Input N=? (N<=10)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   0
      Top             =   3600
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCal_Click()
Dim a()

Form1.Cls
n = Val(txtN) + 1

ReDim a(n, n)

For i = 1 To n

a(i, 1) = 1
a(1, i) = 1
Next i


For i = 2 To n
  For j = 2 To n - i

a(i, j) = a(i - 1, j) + a(i, j - 1)
Next j, i



For i = 2 To n
Print Spc(n - i + 25);

  For j = 1 To i - 1


Print a(j, i - j);

Next j
Print Spc(5);
Print "N=" & i - 1;
Print
Next i
End Sub
