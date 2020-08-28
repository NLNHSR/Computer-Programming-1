VERSION 5.00
Begin VB.Form frmFacebookLogin 
   Caption         =   "Form1"
   ClientHeight    =   9045
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14235
   LinkTopic       =   "Form1"
   ScaleHeight     =   9045
   ScaleWidth      =   14235
   StartUpPosition =   3  'Windows Default
   Begin VB.Image imgFacebookLogin 
      Height          =   8355
      Left            =   4560
      Picture         =   "frmFacebookLogin.frx":0000
      Stretch         =   -1  'True
      Top             =   300
      Width           =   5955
   End
End
Attribute VB_Name = "frmFacebookLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub imgFacebookLogin_Click()

Unload Me
frmSmartPhoneBackground.Show

End Sub
