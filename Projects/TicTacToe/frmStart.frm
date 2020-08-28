VERSION 5.00
Begin VB.Form frmStart 
   Caption         =   "Form1"
   ClientHeight    =   3810
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5655
   LinkTopic       =   "Form1"
   ScaleHeight     =   3810
   ScaleWidth      =   5655
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAi 
      Caption         =   "A.I."
      Height          =   735
      Left            =   3240
      TabIndex        =   5
      Top             =   2400
      Width           =   1455
   End
   Begin VB.CommandButton cmdPvp 
      Caption         =   "PvP"
      Height          =   795
      Left            =   660
      TabIndex        =   4
      Top             =   2340
      Width           =   1395
   End
   Begin VB.TextBox txtPlayer2Name 
      Height          =   735
      Left            =   3300
      TabIndex        =   3
      Top             =   1260
      Width           =   1455
   End
   Begin VB.TextBox txtPlayer1Name 
      Height          =   675
      Left            =   720
      TabIndex        =   2
      Top             =   1260
      Width           =   1395
   End
   Begin VB.Label Label 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Player 2 Name"
      Height          =   555
      Index           =   1
      Left            =   3240
      TabIndex        =   1
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label Label 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Player 1 Name"
      Height          =   555
      Index           =   0
      Left            =   720
      TabIndex        =   0
      Top             =   600
      Width           =   1395
   End
End
Attribute VB_Name = "frmStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAi_Click()

If txtPlayer1Name.Text <> "" Then 'Only opens form if name isn't blank
    frmTTTAITest.Visible = True
End If

End Sub

Private Sub cmdPvp_Click()

If txtPlayer1Name.Text <> "" And txtPlayer2Name.Text <> "" And txtPlayer1Name.Text <> txtPlayer2Name.Text Then 'Only opens form if names 1 and 2 aren't blank and equal to eachother
    frmTicTacToe.Visible = True
End If

End Sub
