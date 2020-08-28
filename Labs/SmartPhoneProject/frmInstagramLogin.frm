VERSION 5.00
Begin VB.Form frmInstagramLogin 
   Caption         =   "Form1"
   ClientHeight    =   9060
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14295
   LinkTopic       =   "Form1"
   ScaleHeight     =   9060
   ScaleWidth      =   14295
   StartUpPosition =   3  'Windows Default
   Begin VB.Image imgInstagramLogin 
      Height          =   8535
      Left            =   4380
      Picture         =   "frmInstagramLogin.frx":0000
      Stretch         =   -1  'True
      Top             =   240
      Width           =   5895
   End
End
Attribute VB_Name = "frmInstagramLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub imgInstagramLogin_Click()

Unload Me
frmSmartPhoneBackground.Show

End Sub
