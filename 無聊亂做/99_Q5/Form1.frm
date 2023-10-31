VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7530
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   15375
   LinkTopic       =   "Form1"
   ScaleHeight     =   7530
   ScaleWidth      =   15375
   StartUpPosition =   3  '¨t²Î¹w³]­È
   Begin VB.Frame Frame2 
      Caption         =   "B«È¤H"
      Height          =   6135
      Left            =   7440
      TabIndex        =   1
      Top             =   240
      Width           =   7095
      Begin VB.Image Image4 
         Height          =   2535
         Index           =   1
         Left            =   3720
         Stretch         =   -1  'True
         Top             =   3240
         Width           =   3015
      End
      Begin VB.Image Image3 
         Height          =   2535
         Index           =   1
         Left            =   240
         Stretch         =   -1  'True
         Top             =   3240
         Width           =   3015
      End
      Begin VB.Image Image2 
         Height          =   2655
         Index           =   1
         Left            =   3720
         Stretch         =   -1  'True
         Top             =   360
         Width           =   3015
      End
      Begin VB.Image Image1 
         Height          =   2655
         Index           =   1
         Left            =   240
         Stretch         =   -1  'True
         Top             =   360
         Width           =   3015
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "A«È¤H"
      Height          =   6135
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   7095
      Begin VB.Image Image4 
         Height          =   2535
         Index           =   0
         Left            =   3720
         Stretch         =   -1  'True
         Top             =   3240
         Width           =   3015
      End
      Begin VB.Image Image3 
         Height          =   2535
         Index           =   0
         Left            =   240
         Stretch         =   -1  'True
         Top             =   3240
         Width           =   3015
      End
      Begin VB.Image Image2 
         Height          =   2655
         Index           =   0
         Left            =   3720
         Stretch         =   -1  'True
         Top             =   360
         Width           =   3015
      End
      Begin VB.Image Image1 
         Height          =   2655
         Index           =   0
         Left            =   240
         Stretch         =   -1  'True
         Top             =   360
         Width           =   3015
      End
   End
   Begin VB.Menu ¥DÀ\ 
      Caption         =   "¥DÀ\"
      Begin VB.Menu ¿»­XÂû¦×¸q¤j§QÄÑ 
         Caption         =   "¿»­XÂû¦×¸q¤j§QÄÑ"
      End
      Begin VB.Menu Ã¹°Ç®üÂA¸q¤j§QÄÑ 
         Caption         =   "Ã¹°Ç®üÂA¸q¤j§QÄÑ"
      End
      Begin VB.Menu ¥¤ªoÂtÂû¸q¤j§QÄÑ 
         Caption         =   "¥¤ªoÂtÂû¸q¤j§QÄÑ"
      End
      Begin VB.Menu ¥Õ°sµðÄ÷¸q¤j§QÄÑ 
         Caption         =   "¥Õ°sµðÄ÷¸q¤j§QÄÑ"
      End
   End
   Begin VB.Menu ¨F©Ô 
      Caption         =   "¨F©Ô"
      Begin VB.Menu ¤ôªGÀu®æ¨F©Ô 
         Caption         =   "¤ôªGÀu®æ¨F©Ô"
      End
      Begin VB.Menu ¥Í¬K±²¨F©Ô 
         Caption         =   "¥Í¬K±²¨F©Ô"
      End
      Begin VB.Menu ÚÞ¥Õµ««C½­¨F©Ô 
         Caption         =   "ÚÞ¥Õµ««C½­¨F©Ô"
      End
      Begin VB.Menu ¤û¿»­X¨F©Ô 
         Caption         =   "¤û¿»­X¨F©Ô"
      End
   End
   Begin VB.Menu ¿@´ö 
      Caption         =   "¿@´ö"
      Begin VB.Menu ¤j»[µð¸Ä´ö 
         Caption         =   "¤j»[µð¸Ä´ö"
      End
      Begin VB.Menu ¿»­X®üÂA´ö 
         Caption         =   "¿»­X®üÂA´ö"
      End
      Begin VB.Menu §uÆC¨ýäø´ö 
         Caption         =   "§uÆC¨ýäø´ö"
      End
      Begin VB.Menu ¤¸®ð¤û¦×´ö 
         Caption         =   "¤¸®ð¤û¦×´ö"
      End
   End
   Begin VB.Menu ²¢ÂI 
      Caption         =   "²¢ÂI"
      Begin VB.Menu ­ì¿N¦B²N²O 
         Caption         =   "­ì¿N¦B²N²O"
      End
      Begin VB.Menu ®Ûªá¬õ¨§´ö 
         Caption         =   "®Ûªá¬õ¨§´ö"
      End
      Begin VB.Menu ¥Õ¥Éµµ¦Ì 
         Caption         =   "¥Õ¥Éµµ¦Ì"
      End
      Begin VB.Menu ¶Â¿}¥¤¹T 
         Caption         =   "¶Â¿}¥¤¹T"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim A, B, C, D, S(1)





Private Sub ¤j»[µð¸Ä´ö_Click()
Image3(C).Picture = LoadPicture(App.Path & "/9.bmp")
C = C + 1
S(C) = S(C) + 45
End Sub

Private Sub ¤¸®ð¤û¦×´ö_Click()
Image3(C).Picture = LoadPicture(App.Path & "/12.bmp")
C = C + 1
S(C) = S(C) + 50
End Sub

Private Sub ¤ôªGÀu®æ¨F©Ô_Click()
Image2(B).Picture = LoadPicture(App.Path & "/5.bmp")
B = B + 1
S(B) = S(B) + 45
End Sub

Private Sub ¤û¿»­X¨F©Ô_Click()
Image2(B).Picture = LoadPicture(App.Path & "/8.bmp")
B = B + 1
S(B) = S(B) + 40
End Sub



Private Sub ¥¤ªoÂtÂû¸q¤j§QÄÑ_Click()
Image1(A).Picture = LoadPicture(App.Path & "/3.bmp")
A = A + 1
S(A) = S(A) + 65
End Sub

Private Sub ¥Í¬K±²¨F©Ô_Click()
Image2(B).Picture = LoadPicture(App.Path & "/6.bmp")
B = B + 1
S(B) = S(B) + 40
End Sub

Private Sub ¥Õ¥Éµµ¦Ì_Click()
Image4(D).Picture = LoadPicture(App.Path & "/15.bmp")
D = D + 1
S(D) = S(D) + 40
End Sub

Private Sub ¥Õ°sµðÄ÷¸q¤j§QÄÑ_Click()
Image2(B).Picture = LoadPicture(App.Path & "/4.bmp")
A = A + 1
S(A) = S(A) + 80
End Sub

Private Sub §uÆC¨ýäø´ö_Click()
Image3(C).Picture = LoadPicture(App.Path & "/11.bmp")
C = C + 1
S(C) = S(C) + 40
End Sub

Private Sub ­ì¿N¦B²N²O_Click()
Image4(D).Picture = LoadPicture(App.Path & "/13.bmp")
D = D + 1
S(D) = S(D) + 40
End Sub

Private Sub ®Ûªá¬õ¨§´ö_Click()
Image4(D).Picture = LoadPicture(App.Path & "/14.bmp")
D = D + 1
S(D) = S(D) + 45
End Sub

Private Sub ¶Â¿}¥¤¹T_Click()
Image4(D).Picture = LoadPicture(App.Path & "/16.bmp")
D = D + 1
S(D) = S(D) + 35
End Sub

Private Sub ÚÞ¥Õµ««C½­¨F©Ô_Click()
Image1(A).Picture = LoadPicture(App.Path & "/7.bmp")
B = B + 1
S(B) = S(B) + 50
End Sub

Private Sub ¿»­X®üÂA´ö_Click()
Image3(C).Picture = LoadPicture(App.Path & "/10.bmp")
C = C + 1

S(C) = S(C) + 40
End Sub

Private Sub ¿»­XÂû¦×¸q¤j§QÄÑ_Click()
Image1(A).Picture = LoadPicture(App.Path & "/1.bmp")
A = A + 1
S(A) = S(A) + 75
End Sub

Private Sub Ã¹°Ç®üÂA¸q¤j§QÄÑ_Click()
Image1(A).Picture = LoadPicture(App.Path & "/2.bmp")
A = A + 1
S(A) = S(A) + 70
End Sub
