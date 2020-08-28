VERSION 5.00
Begin VB.Form frmFile 
   Caption         =   "Chose a Filename"
   ClientHeight    =   4905
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5805
   LinkTopic       =   "Form1"
   ScaleHeight     =   4905
   ScaleWidth      =   5805
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return"
      Height          =   435
      Left            =   4080
      TabIndex        =   8
      Top             =   4020
      Width           =   1515
   End
   Begin VB.TextBox txtFName 
      Height          =   375
      Left            =   300
      TabIndex        =   3
      Text            =   "Enter the filename"
      Top             =   3480
      Width           =   4215
   End
   Begin VB.FileListBox File 
      Height          =   2430
      Left            =   240
      TabIndex        =   2
      Top             =   900
      Width           =   2655
   End
   Begin VB.DirListBox Dir 
      Height          =   2340
      Left            =   3120
      TabIndex        =   1
      Top             =   1020
      Width           =   2475
   End
   Begin VB.DriveListBox Drive 
      Height          =   315
      Left            =   3120
      TabIndex        =   0
      Top             =   600
      Width           =   2475
   End
   Begin VB.Label lblPName 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   300
      TabIndex        =   7
      Top             =   4380
      Width           =   105
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "choose file from list"
      Height          =   375
      Index           =   1
      Left            =   240
      TabIndex        =   6
      Top             =   540
      Width           =   2655
   End
   Begin VB.Label Label 
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Your Selection:"
      Height          =   315
      Index           =   0
      Left            =   300
      TabIndex        =   5
      Top             =   4020
      Width           =   1275
   End
   Begin VB.Label lblPath 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   240
      Width           =   105
   End
End
Attribute VB_Name = "frmFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim FName As String

Sub Form_Load()
    lblPath = Dir.path
End Sub

Sub Drive_Change()
    Dir.path = Drive.Drive
    Dir.SetFocus
End Sub

Sub Dir_Change()
    File.path = Dir.path
    lblPath = "Path: " & Dir.path
    File.SetFocus
End Sub

Sub File_DBLClick()
    '-Display the selected file name when DBLClicked.
    lblPName.Caption = Dir.path & "\" & UCase(File.FileName)
    txtFName.Text = Dir.path & "\" & UCase(File.FileName)
    cmdReturn.SetFocus
End Sub

Sub file_change()
    lblPath.Caption = "Path: " & Dir.path
End Sub

Sub txtFName_Keypress(keyascii As Integer)
    If keyascii = 13 Then
        lblPName = Dir.path & "\" & UCase$(Trim$((txtFName.Text))) & ".dat"
        cmdReturn.SetFocus
    End If
End Sub

Sub cmdReturn_Click()
    frmPaint.lblFileName = lblPName.Caption
    frmFile.Hide
    frmPaint.Show
    frmPaint.lblopentest = 1
End Sub
