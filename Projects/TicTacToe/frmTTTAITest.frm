VERSION 5.00
Begin VB.Form frmTTTAITest 
   Caption         =   "Form1"
   ClientHeight    =   6465
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10020
   LinkTopic       =   "Form1"
   ScaleHeight     =   6465
   ScaleWidth      =   10020
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton cmdNewGame 
      BackColor       =   &H8000000B&
      Caption         =   "New"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7500
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   5640
      Width           =   915
   End
   Begin VB.CommandButton cmdReset 
      BackColor       =   &H8000000B&
      Caption         =   "Reset"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6420
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5640
      Width           =   997
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2220
      Left            =   8880
      TabIndex        =   0
      Top             =   3840
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblPlayerName 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6420
      TabIndex        =   21
      Top             =   960
      Width           =   1995
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Player Name:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   6420
      TabIndex        =   20
      Top             =   360
      Width           =   1995
   End
   Begin VB.Label lblNumTies 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6420
      TabIndex        =   19
      Top             =   4920
      Width           =   1995
   End
   Begin VB.Label lblNumOWins 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6420
      TabIndex        =   18
      Top             =   3720
      Width           =   1995
   End
   Begin VB.Label lblNumXWins 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6420
      TabIndex        =   17
      Top             =   2520
      Width           =   1995
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "# Ties"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   6420
      TabIndex        =   16
      Top             =   4320
      Width           =   1995
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Computer(O) # Wins"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   6420
      TabIndex        =   15
      Top             =   3180
      Width           =   1995
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Player(X) # Wins"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   6420
      TabIndex        =   14
      Top             =   1920
      Width           =   1995
   End
   Begin VB.Label lblTie 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Its a Tie!"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   1980
      TabIndex        =   13
      Top             =   2820
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label lblOWins 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "O Wins!"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   2040
      TabIndex        =   12
      Top             =   2820
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label lblXWins 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "X Wins!"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   700
      Left            =   1980
      TabIndex        =   11
      Top             =   2820
      Visible         =   0   'False
      Width           =   2300
   End
   Begin VB.Line Line8 
      BorderColor     =   &H000000FF&
      BorderWidth     =   5
      Visible         =   0   'False
      X1              =   540
      X2              =   5760
      Y1              =   780
      Y2              =   5700
   End
   Begin VB.Line Line7 
      BorderColor     =   &H000000FF&
      BorderWidth     =   5
      Visible         =   0   'False
      X1              =   660
      X2              =   5640
      Y1              =   5520
      Y2              =   960
   End
   Begin VB.Line Line6 
      BorderColor     =   &H000000FF&
      BorderWidth     =   5
      Visible         =   0   'False
      X1              =   4920
      X2              =   4920
      Y1              =   660
      Y2              =   5760
   End
   Begin VB.Line Line5 
      BorderColor     =   &H000000FF&
      BorderWidth     =   5
      Visible         =   0   'False
      X1              =   3120
      X2              =   3120
      Y1              =   600
      Y2              =   5940
   End
   Begin VB.Line Line4 
      BorderColor     =   &H000000FF&
      BorderWidth     =   5
      Visible         =   0   'False
      X1              =   1320
      X2              =   1320
      Y1              =   540
      Y2              =   5820
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000000FF&
      BorderWidth     =   5
      Visible         =   0   'False
      X1              =   720
      X2              =   6000
      Y1              =   5040
      Y2              =   5040
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000FF&
      BorderWidth     =   5
      Visible         =   0   'False
      X1              =   540
      X2              =   5640
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Line line1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   5
      Visible         =   0   'False
      X1              =   480
      X2              =   5580
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Label lbl9 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   60
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   4200
      TabIndex        =   10
      Top             =   4140
      Width           =   1455
   End
   Begin VB.Label lbl8 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   60
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1395
      Left            =   2400
      TabIndex        =   9
      Top             =   4140
      Width           =   1575
   End
   Begin VB.Label lbl7 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   60
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   720
      TabIndex        =   8
      Top             =   4140
      Width           =   1455
   End
   Begin VB.Label lbl6 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   60
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   4200
      TabIndex        =   7
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Label lbl5 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   60
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1395
      Left            =   2400
      TabIndex        =   6
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Label lbl4 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   60
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1395
      Left            =   720
      TabIndex        =   5
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Label lbl3 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   60
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   4200
      TabIndex        =   4
      Top             =   900
      Width           =   1455
   End
   Begin VB.Label lbl2 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   60
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1395
      Left            =   2400
      TabIndex        =   3
      Top             =   960
      Width           =   1515
   End
   Begin VB.Label lbl1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   60
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1395
      Left            =   780
      TabIndex        =   2
      Top             =   900
      Width           =   1395
   End
   Begin VB.Image Image 
      Height          =   5775
      Left            =   180
      Picture         =   "frmTTTAITest.frx":0000
      Stretch         =   -1  'True
      Top             =   360
      Width           =   6135
   End
End
Attribute VB_Name = "frmTTTAITest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lIndex As Long
Dim Text1, NumXWins, NumOWins, NumTies As Integer
Dim OneClick, TwoClick, ThreeClick, FourClick, FiveClick, SixClick, SevenClick, EightClick, NineClick, Win As Boolean
Dim P1Name As String

Private Sub cmdNewGame_Click()

Call cmdReset_Click 'Resets the game so the new game is blank
frmStart.Visible = True 'Opens Start form
frmTTTAITest.Visible = False 'Closes Game Form
NumXWins = 0 'Resets number of wins
lblNumXWins.Caption = NumXWins
NumOWins = 0
lblNumOWins.Caption = NumOWins
NumTies = 0
lblNumTies.Caption = lblNumTies

End Sub

Private Sub cmdReset_Click()

'Resets list
Call RemoveNumber(List1, "1")
Call RemoveNumber(List1, "2")
Call RemoveNumber(List1, "3")
Call RemoveNumber(List1, "4")
Call RemoveNumber(List1, "5")
Call RemoveNumber(List1, "6")
Call RemoveNumber(List1, "7")
Call RemoveNumber(List1, "8")
Call RemoveNumber(List1, "9")

List1.AddItem "1"
List1.AddItem "2"
List1.AddItem "3"
List1.AddItem "4"
List1.AddItem "5"
List1.AddItem "6"
List1.AddItem "7"
List1.AddItem "8"
List1.AddItem "9"
Text1 = 0

'Resets game boxes to nothing
lbl1.Caption = ""
lbl2.Caption = ""
lbl3.Caption = ""
lbl4.Caption = ""
lbl5.Caption = ""
lbl6.Caption = ""
lbl7.Caption = ""
lbl8.Caption = ""
lbl9.Caption = ""

'Clears all winning lines
line1.Visible = False
Line2.Visible = False
Line3.Visible = False
Line4.Visible = False
Line5.Visible = False
Line6.Visible = False
Line7.Visible = False
Line8.Visible = False

'Clears variable that determines whether a box has been clicked
OneClick = False
TwoClick = False
ThreeClick = False
FourClick = False
FiveClick = False
SixClick = False
SevenClick = False
EightClick = False
NineClick = False

'Hides any win or tie box
lblXWins.Visible = False
lblOWins.Visible = False
lblTie.Visible = False

Win = False

End Sub

Private Sub Form_Load()

P1Name = frmStart.txtPlayer1Name 'Sets player name to one entered on start
lblPlayerName = P1Name

'Creates the list
List1.AddItem "1"
List1.AddItem "2"
List1.AddItem "3"
List1.AddItem "4"
List1.AddItem "5"
List1.AddItem "6"
List1.AddItem "7"
List1.AddItem "8"
List1.AddItem "9"

'Resets variable that determines whether or not a box has been clicked
OneClick = False
TwoClick = False
ThreeClick = False
FourClick = False
FiveClick = False
SixClick = False
SevenClick = False
EightClick = False
NineClick = False

'Resets number of wins
NumXWins = 0
NumOWins = 0
NumTies = 0

Win = False

End Sub

Public Function RemoveNumber(GObject As Object, ByVal Item As Variant)

Dim i As Integer

'Loop that finds the user entered "Item" in the user entered "GObject" and removes it
For i = GObject.ListCount - 1 To 0 Step -1 'Goes through each item in the list
If GObject.List(i) = Item Then 'Checks if the index of the item in list is equal to "Item"
GObject.RemoveItem i 'Removes Item
End If
Next

End Function

Private Sub lbl1_Click()

If OneClick = False And Win = False Then 'Checks if it hasn't been clicked, and nobody has won
    lbl1.Caption = "x" 'Makes the box an x
    Call RemoveNumber(List1, "1") 'Removes the option for this box from the list, so it cannot be clicked on again
    Call CheckWin 'Checks the win
    Call ComputerTurn 'Tells the computer to go
    OneClick = True 'Makes a variable true to indicate the box has been clicked
End If

End Sub

Private Sub lbl2_Click()

If TwoClick = False And Win = False Then
    lbl2.Caption = "x"
    Call RemoveNumber(List1, "2")
    Call CheckWin
    Call ComputerTurn
    TwoClick = True
End If

End Sub

Private Sub lbl3_Click()

If ThreeClick = False And Win = False Then
    lbl3.Caption = "x"
    Call RemoveNumber(List1, "3")
    Call CheckWin
    Call ComputerTurn
    ThreeClick = True
End If

End Sub

Private Sub lbl4_Click()

If FourClick = False And Win = False Then
    lbl4.Caption = "x"
    Call RemoveNumber(List1, "4")
    Call CheckWin
    Call ComputerTurn
    FourClick = True
End If

End Sub

Private Sub lbl5_Click()

If FiveClick = False And Win = False Then
    lbl5.Caption = "x"
    Call RemoveNumber(List1, "5")
    Call CheckWin
    Call ComputerTurn
    FiveClick = True
End If

End Sub

Private Sub lbl6_Click()

If SixClick = False And Win = False Then
    lbl6.Caption = "x"
    Call RemoveNumber(List1, "6")
    Call CheckWin
    Call ComputerTurn
    SixClick = True
End If

End Sub

Private Sub lbl7_Click()

If SevenClick = False And Win = False Then
    lbl7.Caption = "x"
    Call RemoveNumber(List1, "7")
    Call CheckWin
    Call ComputerTurn
    SevenClick = True
End If

End Sub

Private Sub lbl8_Click()

If EightClick = False And Win = False Then
    lbl8.Caption = "x"
    Call RemoveNumber(List1, "8")
    Call CheckWin
    Call ComputerTurn
    EightClick = True
End If

End Sub

Private Sub lbl9_Click()

If NineClick = False And Win = False Then
    lbl9.Caption = "x"
    Call RemoveNumber(List1, "9")
    Call CheckWin
    Call ComputerTurn
    NineClick = True
End If

End Sub

Public Function CheckWin()

If lbl1.Caption = "x" And lbl2.Caption = "x" And lbl3.Caption = "x" Then 'Checks if three images in a row are set to X(1 = x, 2 = 0)
    line1.Visible = True 'Makes a line that goes through all images, visible
    lblXWins.Visible = True 'Makes the X Wins label visible
    Win = True 'Tells the program that the game has been won
    NumXWins = NumXWins + 1 'Adds to the running total of wins throughout the game
    lblNumXWins.Caption = NumXWins
ElseIf lbl4.Caption = "x" And lbl5.Caption = "x" And lbl6.Caption = "x" Then
    Line2.Visible = True
    lblXWins.Visible = True
    Win = True
    NumXWins = NumXWins + 1
    lblNumXWins.Caption = NumXWins
ElseIf lbl7.Caption = "x" And lbl8.Caption = "x" And lbl9.Caption = "x" Then
    Line3.Visible = True
    lblXWins.Visible = True
    Win = True
    NumXWins = NumXWins + 1
    lblNumXWins.Caption = NumXWins
ElseIf lbl1.Caption = "x" And lbl4.Caption = "x" And lbl7.Caption = "x" Then
    Line4.Visible = True
    lblXWins.Visible = True
    Win = True
    NumXWins = NumXWins + 1
    lblNumXWins.Caption = NumXWins
ElseIf lbl2.Caption = "x" And lbl5.Caption = "x" And lbl8.Caption = "x" Then
    Line5.Visible = True
    lblXWins.Visible = True
    Win = True
    NumXWins = NumXWins + 1
    lblNumXWins.Caption = NumXWins
ElseIf lbl3.Caption = "x" And lbl6.Caption = "x" And lbl9.Caption = "x" Then
    Line6.Visible = True
    lblXWins.Visible = True
    Win = True
    NumXWins = NumXWins + 1
    lblNumXWins.Caption = NumXWins
ElseIf lbl3.Caption = "x" And lbl5.Caption = "x" And lbl7.Caption = "x" Then
    Line7.Visible = True
    lblXWins.Visible = True
    Win = True
    NumXWins = NumXWins + 1
    lblNumXWins.Caption = NumXWins
ElseIf lbl1.Caption = "x" And lbl5.Caption = "x" And lbl9.Caption = "x" Then
    Line8.Visible = True
    lblXWins.Visible = True
    Win = True
    NumXWins = NumXWins + 1
    lblNumXWins.Caption = NumXWins
End If

If lbl1.Caption = "o" And lbl2.Caption = "o" And lbl3.Caption = "o" Then
    line1.Visible = True
    lblOWins.Visible = True
    Win = True
    NumOWins = NumOWins + 1
    lblNumOWins.Caption = NumOWins
ElseIf lbl4.Caption = "o" And lbl5.Caption = "o" And lbl6.Caption = "o" Then
    Line2.Visible = True
    lblOWins.Visible = True
    Win = True
    NumOWins = NumOWins + 1
    lblNumOWins.Caption = NumOWins
ElseIf lbl7.Caption = "o" And lbl8.Caption = "o" And lbl9.Caption = "o" Then
    Line3.Visible = True
    lblOWins.Visible = True
    Win = True
    NumOWins = NumOWins + 1
    lblNumOWins.Caption = NumOWins
ElseIf lbl1.Caption = "o" And lbl4.Caption = "o" And lbl7.Caption = "o" Then
    Line4.Visible = True
    lblOWins.Visible = True
    Win = True
    NumOWins = NumOWins + 1
    lblNumOWins.Caption = NumOWins
ElseIf lbl2.Caption = "o" And lbl5.Caption = "o" And lbl8.Caption = "o" Then
    Line5.Visible = True
    lblOWins.Visible = True
    Win = True
    NumOWins = NumOWins + 1
    lblNumOWins.Caption = NumOWins
ElseIf lbl3.Caption = "o" And lbl6.Caption = "o" And lbl9.Caption = "o" Then
    Line6.Visible = True
    lblOWins.Visible = True
    Win = True
    NumOWins = NumOWins + 1
    lblNumOWins.Caption = NumOWins
ElseIf lbl3.Caption = "o" And lbl5.Caption = "o" And lbl7.Caption = "o" Then
    Line7.Visible = True
    lblOWins.Visible = True
    Win = True
    NumOWins = NumOWins + 1
    lblNumOWins.Caption = NumOWins
ElseIf lbl1.Caption = "o" And lbl5.Caption = "o" And lbl9.Caption = "o" Then
    Line8.Visible = True
    lblOWins.Visible = True
    Win = True
    NumOWins = NumOWins + 1
    lblNumOWins.Caption = NumOWins
End If

If lbl1.Caption <> "" And lbl2.Caption <> "" And lbl3.Caption <> "" And lbl4.Caption <> "" And lbl5.Caption <> "" And lbl6.Caption <> "" And lbl7.Caption <> "" And lbl8.Caption <> "" And lbl9.Caption <> "" And lblXWins.Visible = False And lblOWins.Visible = False Then
    lblTie.Visible = True
    NumTies = NumTies + 1
    lblNumTies = NumTies - 1
End If

End Function

Public Function ComputerTurn()

If Win = False Then
    Randomize
    If List1.ListCount Then
        lIndex = Int(Rnd * List1.ListCount) 'Picks a random item from the list by using a random function on the index
        Text1 = List1.List(lIndex) 'Sets a variable equal to the randomly picked item
        List1.RemoveItem lIndex 'Removes the randomly picked item from the list so that it can't be used again, until reset
    End If
    
    If Text1 = "1" Then 'Checks if the variable is equal to an item from the list
        lbl1.Caption = "o" 'Sets a corresponding box equal to o
        Call CheckWin 'Checks for a win
        OneClick = True 'Sets variable that determines if a box has been clicked to true
    ElseIf Text1 = "2" Then
        lbl2.Caption = "o"
        Call CheckWin
        TwoClick = True
    ElseIf Text1 = "3" Then
        lbl3.Caption = "o"
        Call CheckWin
        ThreeClick = True
    ElseIf Text1 = "4" Then
        lbl4.Caption = "o"
        Call CheckWin
        FourClick = True
    ElseIf Text1 = "5" Then
        lbl5.Caption = "o"
        Call CheckWin
        FiveClick = True
    ElseIf Text1 = "6" Then
        lbl6.Caption = "o"
        Call CheckWin
        SixClick = True
    ElseIf Text1 = "7" Then
        lbl7.Caption = "o"
        Call CheckWin
        SevenClick = True
    ElseIf Text1 = "8" Then
        lbl8.Caption = "o"
        Call CheckWin
        EightClick = True
    ElseIf Text1 = "9" Then
        lbl9.Caption = "o"
        Call CheckWin
        NineClick = True
    End If
End If

End Function
