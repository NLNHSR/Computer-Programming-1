VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form frmSmartPhoneBackground 
   Caption         =   "Form1"
   ClientHeight    =   9135
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14130
   LinkTopic       =   "Form1"
   ScaleHeight     =   9135
   ScaleWidth      =   14130
   StartUpPosition =   3  'Windows Default
   Begin WMPLibCtl.WindowsMediaPlayer wmp1 
      Height          =   495
      Left            =   6540
      TabIndex        =   6
      Top             =   7380
      Visible         =   0   'False
      Width           =   555
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   979
      _cy             =   873
   End
   Begin VB.Label lblWhatsApp 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "Whatsapp"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   315
      Left            =   7260
      TabIndex        =   5
      Top             =   6780
      Width           =   1215
   End
   Begin VB.Label lblYoutube 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "Youtube"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   315
      Left            =   5040
      TabIndex        =   4
      Top             =   6780
      Width           =   1095
   End
   Begin VB.Label lblFacebook 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "Facebook"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   7140
      TabIndex        =   3
      Top             =   4800
      Width           =   1155
   End
   Begin VB.Label lblTwitter 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "Twitter"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   5040
      TabIndex        =   2
      Top             =   4740
      Width           =   1095
   End
   Begin VB.Label lblInstagram 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000012&
      Caption         =   "Instagram"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   315
      Left            =   7200
      TabIndex        =   1
      Top             =   2940
      Width           =   1155
   End
   Begin VB.Label lblSnapChat 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "SnapChat"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   5040
      TabIndex        =   0
      Top             =   2940
      Width           =   1155
   End
   Begin VB.Image imgWhatsapp 
      Height          =   1395
      Left            =   7200
      Picture         =   "frmSmartPhoneBackground.frx":0000
      Stretch         =   -1  'True
      Top             =   5220
      Width           =   1395
   End
   Begin VB.Image imgYoutube 
      Height          =   1515
      Left            =   4920
      Picture         =   "frmSmartPhoneBackground.frx":4C02
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   1515
   End
   Begin VB.Image imgFacebook 
      Height          =   1335
      Left            =   7200
      Picture         =   "frmSmartPhoneBackground.frx":88CB
      Stretch         =   -1  'True
      Top             =   3300
      Width           =   1335
   End
   Begin VB.Image imgTwitter 
      Height          =   1395
      Left            =   4920
      Picture         =   "frmSmartPhoneBackground.frx":A188
      Stretch         =   -1  'True
      Top             =   3300
      Width           =   1455
   End
   Begin VB.Image imgInstagram 
      Height          =   1395
      Left            =   7080
      Picture         =   "frmSmartPhoneBackground.frx":B4FC
      Stretch         =   -1  'True
      Top             =   1500
      Width           =   1515
   End
   Begin VB.Image imgSnapChat 
      Height          =   1335
      Left            =   4920
      Picture         =   "frmSmartPhoneBackground.frx":23B27
      Stretch         =   -1  'True
      Top             =   1500
      Width           =   1335
   End
   Begin VB.Image imgPhoneBackground 
      Height          =   8475
      Left            =   2940
      Picture         =   "frmSmartPhoneBackground.frx":27743
      Stretch         =   -1  'True
      Top             =   60
      Width           =   7815
   End
End
Attribute VB_Name = "frmSmartPhoneBackground"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub imgFacebook_Click()

Unload Me
wmp1.URL = "C:\Users\neel.shettigar\19 S2 NS CP1\SmartPhoneProject\BellSound.wav"
frmFacebookLogin.Show

End Sub

Private Sub imgInstagram_Click()

Unload Me
wmp1.URL = "C:\Users\neel.shettigar\19 S2 NS CP1\SmartPhoneProject\BellSound.wav"
frmInstagramLogin.Show

End Sub

Private Sub imgSnapChat_Click()

Unload Me
wmp1.URL = "C:\Users\neel.shettigar\19 S2 NS CP1\SmartPhoneProject\BellSound.wav"
frmSnapChatLogin.Show

End Sub

Private Sub imgTwitter_Click()

Unload Me
wmp1.URL = "C:\Users\neel.shettigar\19 S2 NS CP1\SmartPhoneProject\BellSound.wav"
frmTwitterLogin.Show

End Sub

Private Sub imgWhatsapp_Click()

Unload Me
wmp1.URL = "C:\Users\neel.shettigar\19 S2 NS CP1\SmartPhoneProject\BellSound.wav"
frmWhatsappLogin.Show

End Sub

Private Sub imgYoutube_Click()

Unload Me
wmp1.URL = "C:\Users\neel.shettigar\19 S2 NS CP1\SmartPhoneProject\BellSound.wav"
frmYoutubeLogin.Show

End Sub

Private Sub SnapChat_Change()

End Sub
