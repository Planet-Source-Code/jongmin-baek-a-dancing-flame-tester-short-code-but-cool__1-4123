VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Flame Works"
   ClientHeight    =   3195
   ClientLeft      =   5595
   ClientTop       =   3135
   ClientWidth     =   5565
   ClipControls    =   0   'False
   Icon            =   "FORM1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   5565
   Begin VB.VScrollBar VScroll1 
      Height          =   3015
      Left            =   4680
      Max             =   80
      TabIndex        =   2
      Top             =   120
      Width           =   255
   End
   Begin VB.Timer Timer1 
      Interval        =   30
      Left            =   1920
      Top             =   960
   End
   Begin VB.PictureBox Main 
      BackColor       =   &H80000007&
      Height          =   3015
      Left            =   120
      ScaleHeight     =   3000
      ScaleMode       =   0  'User
      ScaleWidth      =   4400
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
   Begin VB.PictureBox Original 
      BackColor       =   &H80000007&
      Height          =   3015
      Left            =   120
      Picture         =   "FORM1.frx":12FA
      ScaleHeight     =   3000
      ScaleMode       =   0  'User
      ScaleWidth      =   4400
      TabIndex        =   1
      Top             =   120
      Width           =   4455
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000007&
      BackStyle       =   0  'Transparent
      Caption         =   "Height"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   5040
      TabIndex        =   3
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim P As Integer
Dim HeightValue As Integer          'It's the height of the Flame.
Private Sub Form_Load()
Randomize Timer
P = 1
End Sub
Private Sub Timer1_Timer()
g = 20                              'It's the resolution. If it's too small, the program won't run.
                                    'The smaller it is, the slower the program work.
                                    
                                    'If you change the speed of the program, change the variable G or the inteval of Timer1.
                                    
PlusMinus = 0.3                     'It decides how intense the flame dances.
                                    'The bigger it is, the intenser the flame dances.
                                    'Probably 0.3 is the best value. Test many values.
                                   
For I = 0 To Main.ScaleWidth Step g

10 R = Rnd * PlusMinus * 2 + 1 - PlusMinus  'It decides the degree the flame transforms. It's a random value.
If Abs(P - R) > PlusMinus Then GoTo 10 'If the degree varys too much, Make a new random value.

H = HeightValue                     'I use "H" because "HeightValue" is too long

Main.PaintPicture Original.Picture, I, H + Main.ScaleHeight * (1 - R) / 3000 * (3000 - H), g + 10, 3000 - (H + Main.ScaleHeight * (1 - R) / 3000 * (3000 - H)), I, 0, g + 1, Original.ScaleHeight, vbSrcCopy
'PaintPicture. If you move the Main picturebox, you can see "Original" picturebox.

If 1 > R Then Main.Line (I, 0)-(I + g + 10, 1500 + Main.ScaleHeight * (1 - R) / 2), 0, BF
'Fill the empty area with black
P = R
'It saves the degree and next time it will compare P with the new random number.
Next I
End Sub
Private Sub VScroll1_Change()
HeightValue = VScroll1.Value * 25
'Update HeightValue as the scroll is changed.
End Sub
Private Sub VScroll1_Scroll()
HeightValue = VScroll1.Value * 25
'Update HeightValue as the scroll is scrolled.
End Sub
