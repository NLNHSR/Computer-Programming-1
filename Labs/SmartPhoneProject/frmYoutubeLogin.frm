VERSION 5.00
Begin VB.Form frmYoutubeLogin 
   Caption         =   "Form1"
   ClientHeight    =   9150
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14040
   LinkTopic       =   "Form1"
   ScaleHeight     =   9150
   ScaleWidth      =   14040
   StartUpPosition =   3  'Windows Default
   Begin VB.Image imgYoutubeLogin 
      Height          =   8415
      Left            =   3600
      Picture         =   "frmYoutubeLogin.frx":0000
      Stretch         =   -1  'True
      Top             =   300
      Width           =   5775
   End
End
Attribute VB_Name = "frmYoutubeLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub imgYoutubeLogin_Click()

Unload Me
frmSmartPhoneBackground.Show

End Sub
