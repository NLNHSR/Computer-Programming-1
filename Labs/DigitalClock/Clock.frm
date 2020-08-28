VERSION 5.00
Begin VB.Form Clock 
   BackColor       =   &H000000FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Digital Clock"
   ClientHeight    =   2820
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4560
   FillColor       =   &H80000001&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2820
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer 
      Interval        =   1000
      Left            =   1200
      Top             =   2040
   End
   Begin VB.CommandButton Ok 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "OK"
      Height          =   375
      Left            =   3120
      MaskColor       =   &H80000008&
      TabIndex        =   1
      Top             =   1920
      UseMaskColor    =   -1  'True
      Width           =   675
   End
   Begin VB.Label lblTime 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   3195
   End
End
Attribute VB_Name = "Clock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim nHeight As Integer
Dim n, x As Integer

Private Sub Ok_Click()

End

End Sub

Private Sub Timer_Timer()

lblTime = Time

n = Second(Now) Mod 10
If n = 0 Then
    nHeight = 10
ElseIf n = 1 Or n = 9 Then
    nHeight = 15
ElseIf n = 2 Or n = 8 Then
    nHeight = 20
ElseIf n = 3 Or n = 7 Then
    nHeight = 25
ElseIf n = 4 Or n = 6 Then
    nHeight = 30
Else: n = 5
    nHeight = 35
End If
lblTime.FontSize = nHeight
lblTime.Caption = Time$

x = Second(Now) Mod 2
If x = 0 Then
    lblTime.BackColor = &H0&
    lblTime.ForeColor = &H80000005
ElseIf x = 1 Then
    lblTime.BackColor = &H80000005
    lblTime.ForeColor = &H0&
End If

End Sub
