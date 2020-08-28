VERSION 5.00
Begin VB.Form frmMyNameLoop 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.Label lblName 
      Height          =   1455
      Left            =   360
      TabIndex        =   0
      Top             =   300
      Width           =   3375
   End
End
Attribute VB_Name = "frmMyNameLoop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

Dim x As Integer
For x = 1 To 10
    Print x; lblName.Caption = "My name is Neel"
Next x

End Sub
