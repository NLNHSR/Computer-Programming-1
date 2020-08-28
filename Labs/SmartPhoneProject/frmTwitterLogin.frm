VERSION 5.00
Begin VB.Form frmTwitterLogin 
   Caption         =   "Form1"
   ClientHeight    =   9045
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14145
   LinkTopic       =   "Form1"
   ScaleHeight     =   9045
   ScaleWidth      =   14145
   StartUpPosition =   3  'Windows Default
   Begin VB.Image imgTwitterLogin 
      Height          =   8415
      Left            =   4500
      Picture         =   "frmTwitterLogin.frx":0000
      Stretch         =   -1  'True
      Top             =   360
      Width           =   4995
   End
End
Attribute VB_Name = "frmTwitterLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub imgTwitterLogin_Click()

Unload Me
frmSmartPhoneBackground.Show

End Sub
