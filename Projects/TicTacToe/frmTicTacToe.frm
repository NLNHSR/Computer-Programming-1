VERSION 5.00
Begin VB.Form frmTicTacToe 
   BackColor       =   &H8000000E&
   Caption         =   "Form1"
   ClientHeight    =   10260
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11790
   LinkTopic       =   "Form1"
   ScaleHeight     =   10260
   ScaleWidth      =   11790
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton cmdCheckWinnerCount 
      Height          =   375
      Left            =   9420
      TabIndex        =   18
      Top             =   7800
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Timer Timer 
      Interval        =   100
      Left            =   1380
      Top             =   4140
   End
   Begin VB.CommandButton cmdCheckWinner 
      Height          =   435
      Left            =   9480
      TabIndex        =   17
      Top             =   7080
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   675
      Left            =   9000
      TabIndex        =   7
      Top             =   5040
      Width           =   1815
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset"
      Height          =   795
      Left            =   8940
      TabIndex        =   6
      Top             =   3960
      Width           =   1935
   End
   Begin VB.CommandButton cmdNewGame 
      Caption         =   "New Game"
      Height          =   795
      Left            =   8940
      TabIndex        =   5
      Top             =   2880
      Width           =   1995
   End
   Begin VB.Image Image 
      Height          =   2295
      Left            =   2280
      Picture         =   "frmTicTacToe.frx":0000
      Stretch         =   -1  'True
      Top             =   240
      Width           =   8535
   End
   Begin VB.Label lblTie 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Its A Tie!"
      Height          =   1275
      Left            =   4800
      TabIndex        =   4
      Top             =   4440
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label lblXWins 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "X Wins!"
      Height          =   1395
      Left            =   4620
      TabIndex        =   2
      Top             =   4320
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label lblOWins 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "O Wins!"
      Height          =   1275
      Left            =   4680
      TabIndex        =   3
      Top             =   4380
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Caption         =   "Total Ties"
      Height          =   255
      Index           =   7
      Left            =   600
      TabIndex        =   16
      Top             =   7500
      Width           =   1335
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Caption         =   "Total O Wins"
      Height          =   255
      Index           =   6
      Left            =   540
      TabIndex        =   15
      Top             =   6480
      Width           =   1515
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Caption         =   "Total X Wins"
      Height          =   255
      Index           =   5
      Left            =   360
      TabIndex        =   14
      Top             =   5220
      Width           =   1635
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Caption         =   "Turn"
      Height          =   255
      Index           =   4
      Left            =   480
      TabIndex        =   13
      Top             =   2700
      Width           =   1215
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Caption         =   "Player Name"
      Height          =   255
      Index           =   3
      Left            =   480
      TabIndex        =   12
      Top             =   1680
      Width           =   1275
   End
   Begin VB.Line Line 
      BorderWidth     =   5
      Index           =   3
      X1              =   3240
      X2              =   7800
      Y1              =   5700
      Y2              =   5700
   End
   Begin VB.Line Line 
      BorderWidth     =   5
      Index           =   2
      X1              =   3240
      X2              =   7920
      Y1              =   4260
      Y2              =   4260
   End
   Begin VB.Line Line 
      BorderWidth     =   5
      Index           =   1
      X1              =   6300
      X2              =   6300
      Y1              =   2820
      Y2              =   7440
   End
   Begin VB.Line Line 
      BorderWidth     =   5
      Index           =   0
      X1              =   4860
      X2              =   4860
      Y1              =   2760
      Y2              =   7380
   End
   Begin VB.Line Line8 
      BorderColor     =   &H000000FF&
      BorderWidth     =   10
      Visible         =   0   'False
      X1              =   3180
      X2              =   7860
      Y1              =   2820
      Y2              =   7140
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000000FF&
      BorderWidth     =   10
      Visible         =   0   'False
      X1              =   2940
      X2              =   8340
      Y1              =   6480
      Y2              =   6480
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000FF&
      BorderWidth     =   10
      Visible         =   0   'False
      X1              =   3000
      X2              =   7860
      Y1              =   5040
      Y2              =   5040
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   10
      Visible         =   0   'False
      X1              =   3060
      X2              =   7980
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Line Line6 
      BorderColor     =   &H000000FF&
      BorderWidth     =   10
      Visible         =   0   'False
      X1              =   6960
      X2              =   6960
      Y1              =   2760
      Y2              =   7200
   End
   Begin VB.Line Line5 
      BorderColor     =   &H000000FF&
      BorderWidth     =   10
      Visible         =   0   'False
      X1              =   5580
      X2              =   5580
      Y1              =   2580
      Y2              =   7380
   End
   Begin VB.Line Line4 
      BorderColor     =   &H000000FF&
      BorderWidth     =   10
      Visible         =   0   'False
      X1              =   4140
      X2              =   4140
      Y1              =   2640
      Y2              =   7260
   End
   Begin VB.Line Line7 
      BorderColor     =   &H000000FF&
      BorderWidth     =   10
      Visible         =   0   'False
      X1              =   3120
      X2              =   7980
      Y1              =   7140
      Y2              =   3000
   End
   Begin VB.Label lblPlayer 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   2160
      TabIndex        =   11
      Top             =   3120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image img2 
      BorderStyle     =   1  'Fixed Single
      Height          =   1215
      Left            =   4980
      Stretch         =   -1  'True
      Top             =   2940
      Width           =   1275
   End
   Begin VB.Image img1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1275
      Left            =   3480
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   1275
   End
   Begin VB.Image img9 
      BorderStyle     =   1  'Fixed Single
      Height          =   1155
      Left            =   6420
      Stretch         =   -1  'True
      Top             =   5760
      Width           =   1155
   End
   Begin VB.Image img8 
      BorderStyle     =   1  'Fixed Single
      Height          =   1035
      Left            =   4920
      Stretch         =   -1  'True
      Top             =   5820
      Width           =   1335
   End
   Begin VB.Image img7 
      BorderStyle     =   1  'Fixed Single
      Height          =   1155
      Left            =   3540
      Stretch         =   -1  'True
      Top             =   5820
      Width           =   1275
   End
   Begin VB.Image img6 
      BorderStyle     =   1  'Fixed Single
      Height          =   1155
      Left            =   6420
      Stretch         =   -1  'True
      Top             =   4440
      Width           =   1155
   End
   Begin VB.Image img5 
      BorderStyle     =   1  'Fixed Single
      Height          =   1155
      Left            =   4980
      Stretch         =   -1  'True
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Image img4 
      BorderStyle     =   1  'Fixed Single
      Height          =   1095
      Left            =   3480
      Stretch         =   -1  'True
      Top             =   4440
      Width           =   1335
   End
   Begin VB.Image img3 
      BorderStyle     =   1  'Fixed Single
      Height          =   1215
      Left            =   6420
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label lblNumTies 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Left            =   480
      TabIndex        =   10
      Top             =   7740
      Width           =   1755
   End
   Begin VB.Label lblNumOWins 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Left            =   420
      TabIndex        =   9
      Top             =   6840
      Width           =   1755
   End
   Begin VB.Label lblNumXWins 
      BorderStyle     =   1  'Fixed Single
      Height          =   675
      Left            =   360
      TabIndex        =   8
      Top             =   5640
      Width           =   1755
   End
   Begin VB.Label lblTurn 
      BorderStyle     =   1  'Fixed Single
      Height          =   675
      Left            =   360
      TabIndex        =   1
      Top             =   3000
      Width           =   1515
   End
   Begin VB.Label lblPlayerName 
      BorderStyle     =   1  'Fixed Single
      Height          =   555
      Left            =   360
      TabIndex        =   0
      Top             =   2040
      Width           =   1815
   End
End
Attribute VB_Name = "frmTicTacToe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim XImg, OImg, P1Name, P2Name As String
Dim Player, One, Two, Three, Four, Five, Six, Seven, Eight, Nine As Integer
Dim OneClick, TwoClick, ThreeClick, FourClick, FiveClick, SixClick, SevenClick, EightClick, NineClick, XWin, OWin As Boolean

Private Sub cmdCheckWinner_Click()

If One = 1 And Two = 1 And Three = 1 Then 'Checks if three images in a row are set to X(1 = x, 2 = 0)
    XWin = True
    line1.Visible = True 'Makes a line that goes through all images, visible
    lblXWins.Visible = True 'Makes the X Wins label visible
ElseIf Four = 1 And Five = 1 And Six = 1 Then
    XWin = True
    Line2.Visible = True
    lblXWins.Visible = True
ElseIf Seven = 1 And Eight = 1 And Nine = 1 Then
    XWin = True
    Line3.Visible = True
    lblXWins.Visible = True
ElseIf One = 1 And Four = 1 And Seven = 1 Then
    XWin = True
    Line4.Visible = True
    lblXWins.Visible = True
ElseIf Two = 1 And Five = 1 And Eight = 1 Then
    XWin = True
    Line5.Visible = True
    lblXWins.Visible = True
ElseIf Three = 1 And Six = 1 And Nine = 1 Then
    XWin = True
    Line6.Visible = True
    lblXWins.Visible = True
ElseIf Three = 1 And Five = 1 And Seven = 1 Then
    XWin = True
    Line7.Visible = True
    lblXWins.Visible = True
ElseIf One = 1 And Five = 1 And Nine = 1 Then
    XWin = True
    Line8.Visible = True
    lblXWins.Visible = True
ElseIf One = 2 And Two = 2 And Three = 2 Then
    OWin = True
    line1.Visible = True
    lblOWins.Visible = True
ElseIf Four = 2 And Five = 2 And Six = 2 Then
    OWin = True
    Line2.Visible = True
    lblOWins.Visible = True
ElseIf Seven = 2 And Eight = 2 And Nine = 2 Then
    OWin = True
    Line3.Visible = True
    lblOWins.Visible = True
ElseIf One = 2 And Four = 2 And Seven = 2 Then
    OWin = True
    Line4.Visible = True
    lblOWins.Visible = True
ElseIf Two = 2 And Five = 2 And Eight = 2 Then
    OWin = True
    Line5.Visible = True
    lblOWins.Visible = True
ElseIf Three = 2 And Six = 2 And Nine = 2 Then
    OWin = True
    Line6.Visible = True
    lblOWins.Visible = True
ElseIf Three = 2 And Five = 2 And Seven = 2 Then
    OWin = True
    Line7.Visible = True
    lblOWins.Visible = True
ElseIf One = 2 And Five = 2 And Nine = 2 Then
    OWin = True
    Line8.Visible = True
    lblOWins.Visible = True
Else
    If OneClick = True And TwoClick = True And ThreeClick = True And FourClick = True And FiveClick = True And SixClick = True And SevenClick = True And EightClick = True And NineClick = True Then
        lblTie.Visible = True
    End If
End If

End Sub

Private Sub cmdNewGame_Click()

Call cmdReset_Click 'Resets the game so the new game is blank
frmStart.Visible = True 'Opens the start form
frmTicTacToe.Visible = False 'Closes the Game form
lblNumXWins = 0 'Resets the number of wins
lblNumOWins = 0
lblNumTies = 0

End Sub

Private Sub cmdQuit_Click()

End 'Ends the program

End Sub

Private Sub cmdReset_Click()

Set img1.Picture = Nothing 'Makes the images blank
Set img2.Picture = Nothing
Set img3.Picture = Nothing
Set img4.Picture = Nothing
Set img5.Picture = Nothing
Set img6.Picture = Nothing
Set img7.Picture = Nothing
Set img8.Picture = Nothing
Set img9.Picture = Nothing
line1.Visible = False 'Makes the lines invisible
Line2.Visible = False
Line3.Visible = False
Line4.Visible = False
Line5.Visible = False
Line6.Visible = False
Line7.Visible = False
Line8.Visible = False
OneClick = False 'Makes the varibles that control whether or not an image has been clicked false
TwoClick = False
ThreeClick = False
FourClick = False
FiveClick = False
SixClick = False
SevenClick = False
EightClick = False
NineClick = False
One = 0 'Makes the varibles that control whether an image is x or o equal to neither
Two = 0
Three = 0
Four = 0
Five = 0
Six = 0
Seven = 0
Eight = 0
Nine = 0
XWin = False 'X Wins or O Wins are false
OWin = False
lblXWins.Visible = False
lblOWins.Visible = False
lblTie.Visible = False

End Sub

Private Sub cmdCheckWinnerCount_Click()

If lblXWins.Visible = True Then 'Code for keeping a running total of the wins and ties
    lblNumXWins = lblNumXWins + 1
ElseIf lblOWins.Visible = True Then
    lblNumOWins = lblNumOWins + 1
ElseIf lblTie.Visible = True Then
    lblNumTies = lblNumTies + 1
End If

End Sub

Private Sub Form_Load()

XImg = "C:\Users\neel.shettigar\Desktop\19 S2 NS CP1\TicTacToe\ximage.jpeg" 'Sets path for images
OImg = "C:\Users\neel.shettigar\Desktop\19 S2 NS CP1\TicTacToe\oimage.jpg"
P1Name = frmStart.txtPlayer1Name 'Assigns the names from the start form
P2Name = frmStart.txtPlayer2Name
lblPlayerName = P1Name
lblTurn.Caption = "Player 1 Turn (X)"
OneClick = False
TwoClick = False
ThreeClick = False
FourClick = False
FiveClick = False
SixClick = False
SevenClick = False
EightClick = False
NineClick = False
XWin = False
OWin = False
lblXWins.Visible = False
lblOWins.Visible = False
lblTie.Visible = False
lblNumXWins = 0
lblNumOWins = 0
lblNumTies = 0

End Sub

Private Sub img1_Click()

If XWin = False And OWin = False Then 'Only does code if a game isn't won
    If lblPlayerName = P1Name Then 'Checks if its player ones turn
        If OneClick = False Then 'Checks if it hasn't already been clicked
                Set img1.Picture = LoadPicture(XImg) 'Makes the image an x
                lblPlayer = 1
                OneClick = True
                One = 1
            Call cmdCheckWinner_Click 'Checks if it won
            Call cmdCheckWinnerCount_Click
        End If
    ElseIf lblPlayerName = P2Name Then
        If OneClick = False Then
                Set img1.Picture = LoadPicture(OImg)
                lblPlayer = 2
                OneClick = True
                One = 2
            Call cmdCheckWinner_Click
            Call cmdCheckWinnerCount_Click
        End If
    End If
End If

End Sub

Private Sub img2_Click()

If XWin = False And OWin = False Then
    If lblPlayerName = P1Name Then
        If TwoClick = False Then
                Set img2.Picture = LoadPicture(XImg)
                lblPlayer = 3
                TwoClick = True
                Two = 1
            Call cmdCheckWinner_Click
            Call cmdCheckWinnerCount_Click
        End If
    ElseIf lblPlayerName = P2Name Then
        If TwoClick = False Then
                Set img2.Picture = LoadPicture(OImg)
                lblPlayer = 4
                TwoClick = True
                Two = 2
            Call cmdCheckWinner_Click
            Call cmdCheckWinnerCount_Click
        End If
    End If
End If

End Sub

Private Sub img3_Click()

If XWin = False And OWin = False Then
    If lblPlayerName = P1Name Then
        If ThreeClick = False Then
                Set img3.Picture = LoadPicture(XImg)
                lblPlayer = 5
                ThreeClick = True
                Three = 1
            Call cmdCheckWinner_Click
            Call cmdCheckWinnerCount_Click
        End If
    ElseIf lblPlayerName = P2Name Then
        If ThreeClick = False Then
                Set img3.Picture = LoadPicture(OImg)
                lblPlayer = 6
                ThreeClick = True
                Three = 2
            Call cmdCheckWinner_Click
            Call cmdCheckWinnerCount_Click
        End If
    End If
End If

End Sub

Private Sub img4_Click()

If XWin = False And OWin = False Then
    If lblPlayerName = P1Name Then
        If FourClick = False Then
                Set img4.Picture = LoadPicture(XImg)
                lblPlayer = 7
                FourClick = True
                Four = 1
            Call cmdCheckWinner_Click
            Call cmdCheckWinnerCount_Click
        End If
    ElseIf lblPlayerName = P2Name Then
        If FourClick = False Then
                Set img4.Picture = LoadPicture(OImg)
                lblPlayer = 8
                FourClick = True
                Four = 2
            Call cmdCheckWinner_Click
            Call cmdCheckWinnerCount_Click
        End If
    End If
End If

End Sub

Private Sub img5_Click()

If XWin = False And OWin = False Then
    If lblPlayerName = P1Name Then
        If FiveClick = False Then
                Set img5.Picture = LoadPicture(XImg)
                lblPlayer = 9
                FiveClick = True
                Five = 1
            Call cmdCheckWinner_Click
            Call cmdCheckWinnerCount_Click
        End If
    ElseIf lblPlayerName = P2Name Then
        If FiveClick = False Then
                Set img5.Picture = LoadPicture(OImg)
                lblPlayer = 10
                FiveClick = True
                Five = 2
            Call cmdCheckWinner_Click
            Call cmdCheckWinnerCount_Click
        End If
    End If
End If

End Sub

Private Sub img6_Click()

If XWin = False And OWin = False Then
    If lblPlayerName = P1Name Then
        If SixClick = False Then
                Set img6.Picture = LoadPicture(XImg)
                lblPlayer = 11
                SixClick = True
                Six = 1
            Call cmdCheckWinner_Click
            Call cmdCheckWinnerCount_Click
        End If
    ElseIf lblPlayerName = P2Name Then
        If SixClick = False Then
                Set img6.Picture = LoadPicture(OImg)
                lblPlayer = 12
                SixClick = True
                Six = 2
            Call cmdCheckWinner_Click
            Call cmdCheckWinnerCount_Click
        End If
    End If
End If

End Sub

Private Sub img7_Click()

If XWin = False And OWin = False Then
    If lblPlayerName = P1Name Then
        If SevenClick = False Then
                Set img7.Picture = LoadPicture(XImg)
                lblPlayer = 13
                SevenClick = True
                Seven = 1
            Call cmdCheckWinner_Click
            Call cmdCheckWinnerCount_Click
        End If
    ElseIf lblPlayerName = P2Name Then
        If SevenClick = False Then
                Set img7.Picture = LoadPicture(OImg)
                lblPlayer = 14
                SevenClick = True
                Seven = 2
            Call cmdCheckWinner_Click
            Call cmdCheckWinnerCount_Click
        End If
    End If
End If

End Sub

Private Sub img8_Click()

If XWin = False And OWin = False Then
    If lblPlayerName = P1Name Then
        If EightClick = False Then
                Set img8.Picture = LoadPicture(XImg)
                lblPlayer = 15
                EightClick = True
                Eight = 1
            Call cmdCheckWinner_Click
            Call cmdCheckWinnerCount_Click
        End If
    ElseIf lblPlayerName = P2Name Then
        If EightClick = False Then
                Set img8.Picture = LoadPicture(OImg)
                lblPlayer = 16
                EightClick = True
                Eight = 2
            Call cmdCheckWinner_Click
            Call cmdCheckWinnerCount_Click
        End If
    End If
End If

End Sub

Private Sub img9_Click()

If XWin = False And OWin = False Then
    If lblPlayerName = P1Name Then
        If NineClick = False Then
                Set img9.Picture = LoadPicture(XImg)
                lblPlayer = 17
                NineClick = True
                Nine = 1
            Call cmdCheckWinner_Click
            Call cmdCheckWinnerCount_Click
        End If
    ElseIf lblPlayerName = P2Name Then
        If NineClick = False Then
                Set img9.Picture = LoadPicture(OImg)
                lblPlayer = 18
                NineClick = True
                Nine = 2
            Call cmdCheckWinner_Click
            Call cmdCheckWinnerCount_Click
        End If
    End If
End If

End Sub

Private Sub lblPlayer_Change()

If lblPlayerName = P1Name Then 'Code for alternating between player 1 and 2
    lblPlayerName = P2Name
ElseIf lblPlayerName = P2Name Then
    lblPlayerName = P1Name
End If

If lblPlayerName = P1Name Then
    lblTurn.Caption = "Player 1 Turn (X)"
ElseIf lblPlayerName = P2Name Then
    lblTurn.Caption = "Player 2 Turn (O)"
End If

If lblXWins.Visible = True Then
    lblNumXWins = lblNumXWins + 1
ElseIf lblOWins.Visible = True Then
    lblNumOWins = lblNumOWins + 1
ElseIf lblTie.Visible = True Then
    lblNumTies = lblNumTies + 1
End If

End Sub

Private Sub Timer_Timer()

If lblXWins.BackColor = &HFFFF& Then 'Makes the labels blink
    lblXWins.BackColor = &H80FF&
ElseIf lblXWins.BackColor = &H80FF& Then
    lblXWins.BackColor = &HFFFF&
End If

If lblOWins.BackColor = &HFFFF& Then
    lblOWins.BackColor = &H80FF&
ElseIf lblOWins.BackColor = &H80FF& Then
    lblOWins.BackColor = &HFFFF&
End If

If lblTie.BackColor = &HFFFF& Then
    lblTie.BackColor = &H80FF&
ElseIf lblTie.BackColor = &H80FF& Then
    lblTie.BackColor = &HFFFF&
End If

End Sub
