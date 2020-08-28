VERSION 5.00
Begin VB.Form frmGrossAndNetPay 
   Caption         =   "Form1"
   ClientHeight    =   7575
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14115
   LinkTopic       =   "Form1"
   ScaleHeight     =   7575
   ScaleWidth      =   14115
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuarterly 
      Caption         =   "Quarterly "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   10800
      TabIndex        =   15
      Top             =   1980
      Width           =   1500
   End
   Begin VB.CommandButton cmdSemimonthly 
      Caption         =   "Semi-Monthly"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   10800
      TabIndex        =   14
      Top             =   1320
      Width           =   1500
   End
   Begin VB.CommandButton cmdMonthly 
      Caption         =   "Monthly"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   9000
      TabIndex        =   13
      Top             =   3300
      Width           =   1500
   End
   Begin VB.CommandButton cmdBiweekly 
      Caption         =   "Bi-weekly"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   9000
      TabIndex        =   12
      Top             =   2640
      Width           =   1500
   End
   Begin VB.CommandButton cmdAnnual 
      BackColor       =   &H80000005&
      Caption         =   "Annual"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   9000
      MaskColor       =   &H8000000D&
      TabIndex        =   11
      Top             =   1320
      Width           =   1500
   End
   Begin VB.TextBox txtHourlyPay 
      Alignment       =   2  'Center
      Height          =   500
      Left            =   7260
      TabIndex        =   5
      Top             =   1980
      Width           =   1500
   End
   Begin VB.TextBox txtHoursWorked 
      Alignment       =   2  'Center
      Height          =   500
      Left            =   5460
      TabIndex        =   3
      Top             =   1980
      Width           =   1500
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   10800
      TabIndex        =   2
      Top             =   3300
      Width           =   1500
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   10800
      TabIndex        =   1
      Top             =   2640
      Width           =   1500
   End
   Begin VB.CommandButton cmdWeekly 
      Caption         =   "Weekly"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   9000
      MaskColor       =   &H80000006&
      TabIndex        =   0
      Top             =   1980
      Width           =   1500
   End
   Begin VB.Image Image 
      Height          =   3195
      Left            =   1620
      Picture         =   "frmGrossAndNetPay.frx":0000
      Stretch         =   -1  'True
      Top             =   600
      Width           =   3495
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Gross And Net Pay Calculator "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   4
      Left            =   5460
      TabIndex        =   16
      Top             =   600
      Width           =   6795
   End
   Begin VB.Label lblNetPay 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   7260
      TabIndex        =   10
      Top             =   3300
      Width           =   1500
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Net Pay"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   7260
      TabIndex        =   9
      Top             =   2640
      Width           =   1500
   End
   Begin VB.Label lblGrossPay 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   5460
      TabIndex        =   8
      Top             =   3300
      Width           =   1500
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Gross Pay"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   5460
      TabIndex        =   7
      Top             =   2640
      Width           =   1500
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Hourly Pay"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   7260
      TabIndex        =   6
      Top             =   1320
      Width           =   1500
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Hours Worked Per Week"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   5460
      TabIndex        =   4
      Top             =   1320
      Width           =   1500
   End
End
Attribute VB_Name = "frmGrossAndNetPay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim HoursWorked As Single
Dim HourlyPay, GrossPay, NetPay As Currency


Private Sub cmdAnnual_Click()

Call cmdWeekly_Click

lblGrossPay.Caption = GrossPay * 52
lblNetPay.Caption = NetPay * 52

End Sub

Private Sub cmdBiweekly_Click()

Call cmdWeekly_Click

lblGrossPay.Caption = GrossPay * 2
lblNetPay.Caption = NetPay * 2

End Sub

Private Sub cmdMonthly_Click()

Call cmdWeekly_Click

lblGrossPay.Caption = (GrossPay * 52) / 12
lblNetPay.Caption = (NetPay * 52) / 12

End Sub

Private Sub cmdQuarterly_Click()

Call cmdWeekly_Click

lblGrossPay.Caption = (GrossPay * 52) / 4
lblNetPay.Caption = (NetPay * 52) / 4

End Sub

Private Sub cmdSemiMonthly_Click()

Call cmdWeekly_Click

lblGrossPay.Caption = (GrossPay * 52) / 24
lblNetPay.Caption = (NetPay * 52) / 24


End Sub

Private Sub cmdWeekly_Click()
HoursWorked = Val(txtHoursWorked.Text)
HourlyPay = Val(txtHourlyPay.Text)

If HoursWorked > 40 Then
    GrossPay = ((HoursWorked - 40) * (1.5 * HourlyPay)) + 40 * HourlyPay
Else:
    GrossPay = HoursWorked * HourlyPay
End If

lblGrossPay.Caption = GrossPay

NetPay = GrossPay * 0.7
lblNetPay.Caption = NetPay

End Sub

Private Sub cmdClear_Click()

lblGrossPay = ""
lblNetPay = ""
txtHoursWorked = ""
txtHourlyPay = ""
txtHoursWorked.SetFocus


End Sub

Private Sub cmdQuit_Click()

End

End Sub

