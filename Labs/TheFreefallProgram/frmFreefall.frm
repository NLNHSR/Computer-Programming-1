VERSION 5.00
Begin VB.Form frmFreefall 
   Caption         =   "Freefall"
   ClientHeight    =   2910
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   2385
   LinkTopic       =   "Form1"
   ScaleHeight     =   2910
   ScaleWidth      =   2385
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtTime 
      Height          =   495
      Left            =   420
      TabIndex        =   3
      Text            =   " "
      Top             =   780
      Width           =   1215
   End
   Begin VB.Label lblDistance 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   420
      TabIndex        =   2
      Top             =   1980
      Width           =   1215
   End
   Begin VB.Label lblDistanceCaption 
      Caption         =   "Distance in Feet"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   420
      TabIndex        =   1
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label lblTime 
      Caption         =   "Time in Seconds"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   420
      TabIndex        =   0
      Top             =   180
      Width           =   1335
   End
   Begin VB.Menu mnuCalculate 
      Caption         =   "Calculate"
   End
   Begin VB.Menu mnuClear 
      Caption         =   "Clear"
   End
   Begin VB.Menu mnuQuit 
      Caption         =   "Quit"
   End
End
Attribute VB_Name = "frmFreefall"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub mnuCalculate_Click()

Dim FallTime As Single, Distance As Single
Dim Grav As Single
Grav = 32.2
FallTime = Val(txtTime)
Distance = 0.5 * Grav * FallTime ^ 2
lblDistance.Caption = Distance

End Sub

Private Sub mnuClear_Click()

txtTime = ""
lblDistance = ""
txtTime.SetFocus

End Sub

Private Sub mnuQuit_Click()

End

End Sub
