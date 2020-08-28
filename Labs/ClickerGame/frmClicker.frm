VERSION 5.00
Begin VB.Form frmClicker 
   BackColor       =   &H0000FF00&
   Caption         =   "Form1"
   ClientHeight    =   6240
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6825
   LinkTopic       =   "Form1"
   ScaleHeight     =   6240
   ScaleWidth      =   6825
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAutoclickMultiplier 
      BackColor       =   &H80000005&
      Caption         =   "Cost: 200"
      Height          =   375
      Left            =   5100
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   5280
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Timer Background 
      Interval        =   100
      Left            =   240
      Top             =   2220
   End
   Begin VB.CommandButton cmdDoubleMultiplier 
      BackColor       =   &H80000005&
      Caption         =   "Cost: 100"
      Height          =   375
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   5220
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CommandButton cmdAutoClickQuadrupler 
      BackColor       =   &H80000005&
      Caption         =   "Cost: 30"
      Height          =   375
      Left            =   5100
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   4200
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdMultiplier 
      BackColor       =   &H80000005&
      Caption         =   "Cost: 50"
      Height          =   375
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4080
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CommandButton cmdClickTripler 
      BackColor       =   &H80000005&
      Caption         =   "Cost: 25"
      Height          =   375
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2880
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CommandButton cmdAutoClickDoubler 
      BackColor       =   &H80000005&
      Caption         =   "Cost: 20"
      Height          =   435
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2940
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CommandButton cmdClickDoubler 
      BackColor       =   &H80000005&
      Caption         =   "Cost: 10"
      Height          =   375
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1740
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CommandButton cmdAutoClick 
      BackColor       =   &H80000005&
      Caption         =   "Cost: 15"
      Height          =   435
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1680
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Timer Timer 
      Interval        =   1000
      Left            =   240
      Top             =   1560
   End
   Begin VB.CommandButton cmdClick 
      Caption         =   "Click Me"
      Height          =   2235
      Left            =   840
      Picture         =   "frmClicker.frx":0000
      TabIndex        =   2
      Top             =   1920
      Width           =   2295
   End
   Begin VB.Label lblCPS 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   1800
      TabIndex        =   25
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Clicks Per Second"
      Height          =   315
      Index           =   9
      Left            =   1800
      TabIndex        =   24
      Top             =   420
      Width           =   1695
   End
   Begin VB.Label lblCountdown3 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   6120
      TabIndex        =   23
      Top             =   4680
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "10 sec X5 Auto Click"
      Height          =   495
      Index           =   8
      Left            =   5100
      TabIndex        =   21
      Top             =   4680
      Width           =   975
   End
   Begin VB.Label lblCountdown2 
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   4500
      TabIndex        =   20
      Top             =   4620
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "10 sec. +15 Multiplier "
      Height          =   495
      Index           =   7
      Left            =   3360
      TabIndex        =   18
      Top             =   4620
      Width           =   1035
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Auto Click Quadrupler"
      Height          =   495
      Index           =   6
      Left            =   5100
      TabIndex        =   16
      Top             =   3600
      Width           =   975
   End
   Begin VB.Label lblCountdown 
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Height          =   435
      Left            =   4500
      TabIndex        =   15
      Top             =   3540
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "5 sec. +10 Multiplier"
      Height          =   435
      Index           =   5
      Left            =   3360
      TabIndex        =   13
      Top             =   3540
      Width           =   1035
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Click Tripler"
      Height          =   315
      Index           =   4
      Left            =   3360
      TabIndex        =   11
      Top             =   2400
      Width           =   1035
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Auto Click Doubler"
      Height          =   435
      Index           =   3
      Left            =   5040
      TabIndex        =   9
      Top             =   2400
      Width           =   1035
   End
   Begin VB.Label Label 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Click Doubler"
      Height          =   315
      Index           =   2
      Left            =   3360
      TabIndex        =   7
      Top             =   1320
      Width           =   1035
   End
   Begin VB.Label lblAutoClick 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Auto Click"
      Height          =   255
      Index           =   2
      Left            =   5040
      TabIndex        =   6
      Top             =   1320
      Width           =   1035
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Number of times clicked"
      Height          =   435
      Index           =   1
      Left            =   4020
      TabIndex        =   4
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label lblNumClicks 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   4020
      TabIndex        =   3
      Top             =   780
      Width           =   1335
   End
   Begin VB.Label lblPlayerName 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Player Name"
      Height          =   315
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   420
      Width           =   1335
   End
End
Attribute VB_Name = "frmClicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Clicks, AutoClickCost, ClickAmount, ClickDoublerCost, AutoClickDoublerCost, ClickTriplerCost, MultiplierCost, DoubleMultiplierCost, AutoClickQuadruplerCost, AutoclickMultiplierCost As Integer
Dim PerSecond As Single
Dim AutoClickCheck, AutoClickPurchase, ClickDoublerPurchase, AutoClickDoublerPurchase, ClickTriplerPurchase, MultiplierCheck, MultiplierRunningCheck, DoubleMultiplierCheck, DoubleMultiplierRunningCheck, AutoClickQuadruplerPurchase, AutoclickMultiplierRunningCheck, AutoclickMultiplierCheck As Boolean

Private Sub Background_Timer()

'Cycles through colors in the background, making a flashing appearance

If frmClicker.BackColor = &HFF00& Then
    frmClicker.BackColor = &HC000&
ElseIf frmClicker.BackColor = &HC000& Then
    frmClicker.BackColor = &H8000&
ElseIf frmClicker.BackColor = &H8000& Then
    frmClicker.BackColor = &H4000&
ElseIf frmClicker.BackColor = &H4000& Then
    frmClicker.BackColor = &HFF00&
End If

End Sub

Private Sub cmdAutoClick_Click()

'Codes for AutoClick- automatically clicks the button every one second

If AutoClickPurchase = False Then
    If Clicks >= AutoClickCost Then
        AutoClickCheck = True
        Clicks = Clicks - AutoClickCost
        lblNumClicks = Clicks
        AutoClickPurchase = True
    End If
End If

End Sub

Private Sub cmdAutoClickDoubler_Click()

'Doubles the time it takes autoclick to click- automatically clicking the button twice every second

If AutoClickDoublerPurchase = False Then
    If Clicks >= AutoClickDoublerCost Then
        Timer.Interval = 500
        Clicks = Clicks - AutoClickDoublerCost
        lblNumClicks = Clicks
        AutoClickDoublerPurchase = True
    End If
End If

End Sub

Private Sub cmdAutoclickMultiplier_Click()

'Temporarly multiplies the autoclick interval by ten for ten seconds

If AutoClickPurchase = True Then
    If Clicks >= AutoclickMultiplierCost Then
        If AutoclickMultiplierRunningCheck = False Then
            Clicks = Clicks - AutoclickMultiplierCost
            lblNumClicks = Clicks
            AutoclickMultiplierCheck = True
            lblCountdown3.Caption = 10
        End If
    End If
End If

End Sub

Private Sub cmdAutoClickQuadrupler_Click()

'Doubles the autoclick interval again, automatically clicking 4 times in a second

If AutoClickQuadruplerPurchase = False Then
    If Clicks >= AutoClickQuadruplerCost Then
        Timer.Interval = 250
        Clicks = Clicks - AutoClickQuadruplerCost
        lblNumClicks = Clicks
        AutoClickQuadruplerPurchase = True
    End If
End If

End Sub

Private Sub cmdClick_Click()

Clicks = Val(lblNumClicks)  'Tells the computer where it can find the value of the variable
Clicks = Clicks + ClickAmount  'Looks at the value of the variable and adds +1 to the value
lblNumClicks = Clicks    'Displays the value of Clicks in the lblNumClicks

End Sub

Private Sub cmdClickDoubler_Click()

'Increases the click amount to two

If ClickDoublerPurchase = False Then
    If Clicks >= ClickDoublerCost Then
        ClickAmount = 2
        Clicks = Clicks - ClickDoublerCost
        lblNumClicks = Clicks
        ClickDoublerPurchase = True
    End If
End If

End Sub

Private Sub cmdClickTripler_Click()

'Increases the click amount to three

If ClickTriplerPurchase = False Then
    If Clicks >= ClickTriplerCost Then
        ClickAmount = 3
        Clicks = Clicks - ClickTriplerCost
        lblNumClicks = Clicks
        ClickTriplerPurchase = True
    End If
End If

End Sub

Private Sub cmdMultiplier_Click()

'Temporarily increases the click amount to ten for five seconds

If Clicks >= MultiplierCost And DoubleMultiplierRunningCheck = False Then
    If MultiplierRunningCheck = False Then
        Clicks = Clicks - MultiplierCost
        lblNumClicks = Clicks
        MultiplierCheck = True
        lblCountdown.Caption = 5
    End If
End If

End Sub

Private Sub cmdDoubleMultiplier_Click()

'Temporarily increases the click amount to fifteen for ten seconds

If Clicks >= DoubleMultiplierCost And MultiplierRunningCheck = False Then
    If DoubleMultiplierRunningCheck = False Then
        Clicks = Clicks - DoubleMultiplierCost
        lblNumClicks = Clicks
        DoubleMultiplierCheck = True
        lblCountdown2.Caption = 10
    End If
End If

End Sub

Private Sub Form_Load()

lblPlayerName = frmName.txtName 'This gets the entered name from frmName and puts it in the lblPlayerName on frmClicker
ClickAmount = 1
ClickDoublerCost = 10
ClickDoublerPurchase = False
AutoClickCost = 15
AutoClickPurchase = False
AutoClickDoublerCost = 20
AutoClickDoublerPurchase = False
ClickTriplerCost = 25
ClickTriplerPurchase = False
MultiplierCost = 50
MultiplierCheck = False
MultiplierRunningCheck = False
DoubleMultiplierCost = 100
DoubleMultiplierCheck = False
DoubleMultiplierRunningCheck = False
AutoClickQuadruplerCost = 30
AutoClickQuadruplerPurchase = False
AutoclickMultiplierCost = 200
AutoclickMultiplierRunningCheck = False
AutoclickMultiplierCheck = False

End Sub

Private Sub lblNumClicks_Change()


'All code below is to make sure the powerups only show up when they can be purchased

If Clicks >= MultiplierCost Then
    cmdMultiplier.Visible = True
    lblCountdown.Visible = True
ElseIf MultiplierRunningCheck = True Then
    cmdMultiplier.Visible = True
    lblCountdown.Visible = True
ElseIf Clicks < MultiplierCost Then
    cmdMultiplier.Visible = False
    lblCountdown.Visible = False
End If

If Clicks >= AutoclickMultiplierCost And AutoClickPurchase = True Then
    cmdAutoclickMultiplier.Visible = True
    lblCountdown3.Visible = True
ElseIf AutoclickMultiplierRunningCheck = True Then
    cmdAutoclickMultiplier.Visible = True
    lblCountdown3.Visible = True
ElseIf Clicks < MultiplierCost Then
    cmdAutoclickMultiplier.Visible = False
    lblCountdown3.Visible = False
End If

If Clicks >= DoubleMultiplierCost Then
    cmdDoubleMultiplier.Visible = True
    lblCountdown2.Visible = True
ElseIf DoubleMultiplierRunningCheck = True Then
    cmdDoubleMultiplier.Visible = True
    lblCountdown2.Visible = True
ElseIf Clicks < DoubleMultiplierCost Then
    cmdDoubleMultiplier.Visible = False
    lblCountdown2.Visible = False
End If

If Clicks >= ClickDoublerCost And ClickDoublerPurchase = False Then
    cmdClickDoubler.Visible = True
ElseIf Clicks < ClickDoublerCost And ClickDoublerPurchase = False Then
    cmdClickDoubler.Visible = False
ElseIf ClickDoublerPurchase = True Then
    cmdClickDoubler.Visible = True
    cmdClickDoubler.BackColor = &H808080
End If

If Clicks >= AutoClickCost And AutoClickPurchase = False Then
    cmdAutoClick.Visible = True
ElseIf Clicks < AutoClickCost And AutoClickPurchase = False Then
    cmdAutoClick.Visible = False
ElseIf AutoClickPurchase = True Then
    cmdAutoClick.Visible = True
    cmdAutoClick.BackColor = &H808080
End If

If Clicks >= AutoClickDoublerCost And AutoClickDoublerPurchase = False And AutoClickPurchase = True Then
    cmdAutoClickDoubler.Visible = True
ElseIf Clicks < AutoClickDoublerCost And AutoClickDoublerPurchase = False Then
    cmdAutoClickDoubler.Visible = False
ElseIf AutoClickDoublerPurchase = True Then
    cmdAutoClickDoubler.Visible = True
    cmdAutoClickDoubler.BackColor = &H808080
End If

If Clicks >= ClickTriplerCost And ClickTriplerPurchase = False And ClickDoublerPurchase = True Then
    cmdClickTripler.Visible = True
ElseIf Clicks < ClickTriplerCost And ClickTriplerPurchase = False Then
    cmdClickTripler.Visible = False
ElseIf ClickTriplerPurchase = True Then
    cmdClickTripler.Visible = True
    cmdClickTripler.BackColor = &H808080
End If

If Clicks >= AutoClickQuadruplerCost And AutoClickQuadruplerPurchase = False And AutoClickDoublerPurchase = True Then
    cmdAutoClickQuadrupler.Visible = True
ElseIf Clicks < AutoClickQuadruplerCost And AutoClickQuadruplerPurchase = False Then
    cmdAutoClickQuadrupler.Visible = False
ElseIf AutoClickQuadruplerPurchase = True Then
    cmdAutoClickQuadrupler.Visible = True
    cmdAutoClickQuadrupler.BackColor = &H808080
End If

'Calculates the clicks per second

If Timer.Interval = 1000 And AutoClickPurchase = True Then
    PerSecond = 1
ElseIf Timer.Interval = 500 Then
    PerSecond = 2
ElseIf Timer.Interval = 250 Then
    PerSecond = 4
ElseIf Timer.Interval = 200 Then
    PerSecond = 5
ElseIf Timer.Interval = 100 Then
    PerSecond = 10
ElseIf Timer.Interval = 50 Then
    PerSecond = 20
End If

lblCPS.Caption = ClickAmount * PerSecond

End Sub



Private Sub Timer_Timer()

If MultiplierCheck = True Then
    MultiplierRunningCheck = True
    If AutoClickDoublerPurchase = True And AutoClickQuadruplerPurchase = False And AutoclickMultiplierRunningCheck = False Then
        lblCountdown.Caption = lblCountdown.Caption - 0.5
    ElseIf AutoClickQuadruplerPurchase = True And AutoclickMultiplierRunningCheck = False Then
        lblCountdown.Caption = lblCountdown.Caption - 0.25
    ElseIf AutoclickMultiplierRunningCheck = False Then
        lblCountdown.Caption = lblCountdown.Caption - 1
    ElseIf AutoClickDoublerPurchase = True And AutoClickQuadruplerPurchase = False And AutoclickMultiplierRunningCheck = True Then
        lblCountdown.Caption = lblCountdown.Caption - 0.1
    ElseIf AutoClickQuadruplerPurchase = True And AutoclickMultiplierRunningCheck = True Then
        lblCountdown.Caption = lblCountdown.Caption - 0.05
    ElseIf AutoclickMultiplierRunningCheck = True Then
        lblCountdown.Caption = lblCountdown.Caption - 0.2
    End If
    ClickAmount = 10
    cmdMultiplier.Visible = True
    lblCountdown.Visible = True
    If lblCountdown.Caption <= 0 Then
        MultiplierCheck = False
        MultiplierRunningCheck = False
        If ClickDoublerPurchase = True Then
            ClickAmount = 2
        ElseIf ClickTriplerPurchase = True Then
            ClickAmount = 3
        Else
            ClickAmount = 1
        End If
        cmdMultiplier.Visible = False
        lblCountdown.Visible = False
        lblCountdown.Caption = ""
    End If
End If

If DoubleMultiplierCheck = True Then
    DoubleMultiplierRunningCheck = True
    If AutoClickDoublerPurchase = True And AutoClickQuadruplerPurchase = False And AutoclickMultiplierRunningCheck = False Then
        lblCountdown2.Caption = lblCountdown2.Caption - 0.5
    ElseIf AutoClickQuadruplerPurchase = True And AutoclickMultiplierRunningCheck = False Then
        lblCountdown2.Caption = lblCountdown2.Caption - 0.25
    ElseIf AutoclickMultiplierRunningCheck = False Then
        lblCountdown2.Caption = lblCountdown2.Caption - 1
    ElseIf AutoClickDoublerPurchase = True And AutoClickQuadruplerPurchase = False And AutoclickMultiplierRunningCheck = True Then
        lblCountdown2.Caption = lblCountdown2.Caption - 0.1
    ElseIf AutoClickQuadruplerPurchase = True And AutoclickMultiplierRunningCheck = True Then
        lblCountdown2.Caption = lblCountdown2.Caption - 0.05
    ElseIf AutoclickMultiplierRunningCheck = True Then
        lblCountdown2.Caption = lblCountdown2.Caption - 0.2
    End If
    ClickAmount = 15
    cmdDoubleMultiplier.Visible = True
    lblCountdown2.Visible = True
    If lblCountdown2.Caption <= 0 Then
        DoubleMultiplierCheck = False
        DoubleMultiplierRunningCheck = False
        If ClickDoublerPurchase = True Then
            ClickAmount = 2
        ElseIf ClickTriplerPurchase = True Then
            ClickAmount = 3
        Else
            ClickAmount = 1
        End If
        cmdDoubleMultiplier.Visible = False
        lblCountdown2.Visible = False
        lblCountdown2.Caption = ""
    End If
End If

If AutoClickCheck = True Then
    Clicks = Clicks + ClickAmount
    lblNumClicks = Clicks
End If

If AutoclickMultiplierCheck = True Then
    AutoclickMultiplierRunningCheck = True
    If AutoClickDoublerPurchase = True And AutoClickQuadruplerPurchase = False Then
        lblCountdown3.Caption = lblCountdown3.Caption - 0.1
    ElseIf AutoClickQuadruplerPurchase = True Then
        lblCountdown3.Caption = lblCountdown3.Caption - 0.05
    Else
        lblCountdown3.Caption = lblCountdown3.Caption - 0.2
    End If
    If AutoClickPurchase = True And AutoClickDoublerPurchase = False And AutoClickQuadruplerPurchase = False Then
        Timer.Interval = 200
    ElseIf AutoClickDoublerPurchase = True And AutoClickQuadruplerPurchase = False Then
        Timer.Interval = 100
    ElseIf AutoClickQuadruplerPurchase = True Then
        Timer.Interval = 50
    End If
    cmdAutoclickMultiplier.Visible = True
    lblCountdown3.Visible = True
    If lblCountdown3.Caption = 0 Then
        AutoclickMultiplierCheck = False
        AutoclickMultiplierRunningCheck = False
        If AutoClickPurchase = True And AutoClickDoublerPurchase = False And AutoClickQuadruplerPurchase = False Then
            Timer.Interval = 1000
        ElseIf AutoClickDoublerPurchase = True And AutoClickQuadruplerPurchase = False Then
            Timer.Interval = 500
        ElseIf AutoClickQuadruplerPurchase = True Then
            Timer.Interval = 250
        End If
        cmdAutoclickMultiplier.Visible = False
        lblCountdown3.Visible = False
        lblCountdown3.Caption = ""
    End If
End If


End Sub
