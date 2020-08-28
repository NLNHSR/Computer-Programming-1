VERSION 5.00
Begin VB.Form frmName 
   BackColor       =   &H00008000&
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtName 
      Height          =   675
      Left            =   1500
      TabIndex        =   0
      Top             =   840
      Width           =   1515
   End
   Begin VB.CommandButton cmdEnter 
      Caption         =   "Enter"
      Height          =   735
      Left            =   1500
      TabIndex        =   1
      Top             =   1560
      Width           =   1515
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Enter Name"
      Height          =   375
      Left            =   1500
      TabIndex        =   2
      Top             =   420
      Width           =   1515
   End
End
Attribute VB_Name = "frmName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdEnter_Click()

If txtName.Text = "" Then
    MsgBox "Please Enter A Name", vbOKOnly, "Name Error"
Else
    frmClicker.Show 'this will show frmClicker when clicked.
End If

End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    Call cmdEnter_Click
End If

End Sub
