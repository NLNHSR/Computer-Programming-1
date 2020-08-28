VERSION 5.00
Begin VB.Form frmCerealProgram 
   Caption         =   "Cereal Projects"
   ClientHeight    =   4275
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7335
   LinkTopic       =   "Form1"
   ScaleHeight     =   4275
   ScaleWidth      =   7335
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer 
      Interval        =   100
      Left            =   4620
      Top             =   900
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   495
      Left            =   3000
      TabIndex        =   4
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   495
      Left            =   3000
      TabIndex        =   3
      Top             =   2820
      Width           =   1215
   End
   Begin VB.CommandButton cmdCheck 
      Caption         =   "Check  Supply"
      Height          =   495
      Left            =   3000
      TabIndex        =   2
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox txtBowls 
      Height          =   700
      Left            =   2640
      TabIndex        =   1
      Top             =   1320
      Width           =   1680
   End
   Begin VB.TextBox txtBoxes 
      Height          =   700
      Left            =   840
      TabIndex        =   0
      Top             =   1320
      Width           =   1500
   End
   Begin VB.Image img5 
      Height          =   1920
      Left            =   4800
      Picture         =   "frmCerealProgram.frx":0000
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   1440
   End
   Begin VB.Image img3 
      Height          =   1920
      Left            =   4800
      Picture         =   "frmCerealProgram.frx":073D
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   1440
   End
   Begin VB.Image img2 
      Height          =   1920
      Left            =   4800
      Picture         =   "frmCerealProgram.frx":0E7A
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   1440
   End
   Begin VB.Image img4 
      Height          =   1920
      Left            =   4800
      Picture         =   "frmCerealProgram.frx":14AC
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   1440
   End
   Begin VB.Image img1 
      Height          =   1920
      Left            =   4800
      Picture         =   "frmCerealProgram.frx":1C95
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   1440
   End
   Begin VB.Label lblOK 
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cereal supply is OK"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   840
      TabIndex        =   8
      Top             =   3120
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Label lblBuy 
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Buy more cereal"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   840
      TabIndex        =   7
      Top             =   2280
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Label Label 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Number of bowls eaten per week."
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Index           =   1
      Left            =   2640
      TabIndex        =   6
      Top             =   420
      Width           =   1680
   End
   Begin VB.Label Label 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Number of boxes on hand."
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Index           =   0
      Left            =   840
      TabIndex        =   5
      Top             =   420
      Width           =   1500
   End
End
Attribute VB_Name = "frmCerealProgram"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Boxes As Integer, Bowls As Integer
Dim Servings As Integer
    
Private Sub cmdCheck_Click()

    Boxes = Val(txtBoxes)
    Bowls = Val(txtBowls)
    Servings = Boxes * 12
    If Servings >= Bowls * 2 Then
        lblOK.Visible = True
        lblBuy.Visible = False
    Else
        lblBuy.Visible = True
        lblOK.Visible = False
    End If
    cmdClear.SetFocus
End Sub

Private Sub cmdClear_Click()
    txtBoxes = ""
    txtBowls = ""
    lblBuy.Visible = False
    lblOK.Visible = False
    txtBoxes.SetFocus
End Sub

Private Sub cmdQuit_Click()

End

End Sub

Private Sub Timer_Timer()

If img1.Visible = True Then
    img1.Visible = False
    img2.Visible = True
    img3.Visible = False
    img4.Visible = False
    img5.Visible = False
ElseIf img2.Visible = True Then
    img1.Visible = False
    img2.Visible = False
    img3.Visible = True
    img4.Visible = False
    img5.Visible = False
ElseIf img3.Visible = True Then
    img1.Visible = False
    img2.Visible = False
    img3.Visible = False
    img4.Visible = True
    img5.Visible = False
ElseIf img4.Visible = True Then
    img1.Visible = False
    img2.Visible = False
    img3.Visible = False
    img4.Visible = False
    img5.Visible = True
ElseIf img5.Visible = True Then
    img1.Visible = True
    img2.Visible = False
    img3.Visible = False
    img4.Visible = False
    img5.Visible = False
End If

End Sub

Private Sub txtBowls_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtBoxes.SetFocus
End If
End Sub


Private Sub txtBoxes_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtBowls.SetFocus
End If
End Sub
