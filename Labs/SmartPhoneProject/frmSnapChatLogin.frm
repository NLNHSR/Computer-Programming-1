VERSION 5.00
Begin VB.Form frmSnapChatLogin 
   Caption         =   "Form1"
   ClientHeight    =   9105
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14145
   LinkTopic       =   "Form1"
   ScaleHeight     =   9105
   ScaleWidth      =   14145
   StartUpPosition =   3  'Windows Default
   Begin VB.Image imgSnapChatLogin 
      Height          =   8235
      Left            =   660
      Picture         =   "frmSnapChatLogin.frx":0000
      Stretch         =   -1  'True
      Top             =   420
      Width           =   12915
   End
End
Attribute VB_Name = "frmSnapChatLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub imgSnapChatLogin_Click()

Unload Me
frmSmartPhoneBackground.Show

End Sub
