VERSION 5.00
Begin VB.Form frmWhatsappLogin 
   Caption         =   "Form1"
   ClientHeight    =   9135
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14235
   LinkTopic       =   "Form1"
   ScaleHeight     =   9135
   ScaleWidth      =   14235
   StartUpPosition =   3  'Windows Default
   Begin VB.Image imgWhatsappLogin 
      Height          =   8115
      Left            =   4440
      Picture         =   "frmWhatsappLogin.frx":0000
      Stretch         =   -1  'True
      Top             =   600
      Width           =   5235
   End
End
Attribute VB_Name = "frmWhatsappLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub imgWhatsappLogin_Click()

Unload Me
frmSmartPhoneBackground.Show

End Sub
