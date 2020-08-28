VERSION 5.00
Begin VB.Form frmAmortizationTableProject 
   Caption         =   "Amortization Table Project"
   ClientHeight    =   5715
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6825
   LinkTopic       =   "Form1"
   ScaleHeight     =   5715
   ScaleWidth      =   6825
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtYearlyExtra 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   5460
      TabIndex        =   23
      Top             =   2400
      Width           =   1155
   End
   Begin VB.TextBox txtMonthlyExtra 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   4140
      TabIndex        =   22
      Top             =   2400
      Width           =   1155
   End
   Begin VB.TextBox txtOneTimeExtra 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   2700
      TabIndex        =   21
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   5460
      TabIndex        =   14
      Top             =   2940
      Width           =   1155
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   495
      Left            =   4140
      TabIndex        =   13
      Top             =   2940
      Width           =   1155
   End
   Begin VB.CommandButton cmdCalculate 
      Caption         =   "Calculate"
      Height          =   495
      Left            =   2700
      TabIndex        =   12
      Top             =   2940
      Width           =   1215
   End
   Begin VB.Frame Frame 
      Caption         =   "Amortization Table"
      Height          =   1935
      Index           =   2
      Left            =   120
      TabIndex        =   9
      Top             =   3720
      Width           =   6495
      Begin VB.HScrollBar hsbPayment 
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1440
         Width           =   6255
      End
      Begin VB.TextBox txtAmortTable 
         BorderStyle     =   0  'None
         Height          =   555
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   10
         Top             =   720
         Width           =   6255
      End
      Begin VB.Label Label 
         Caption         =   "Monthly Principle"
         Height          =   315
         Index           =   8
         Left            =   5160
         TabIndex        =   20
         Top             =   420
         Width           =   1335
      End
      Begin VB.Label Label 
         Caption         =   "Monthly Interest"
         Height          =   255
         Index           =   7
         Left            =   3900
         TabIndex        =   19
         Top             =   420
         Width           =   1275
      End
      Begin VB.Label Label 
         Caption         =   "Total Interest "
         Height          =   255
         Index           =   6
         Left            =   2820
         TabIndex        =   18
         Top             =   420
         Width           =   975
      End
      Begin VB.Label Label 
         Caption         =   "Current Balance"
         Height          =   255
         Index           =   5
         Left            =   1560
         TabIndex        =   17
         Top             =   420
         Width           =   1215
      End
      Begin VB.Label Label 
         Caption         =   "Year"
         Height          =   255
         Index           =   4
         Left            =   1080
         TabIndex        =   16
         Top             =   420
         Width           =   975
      End
      Begin VB.Label Label 
         Caption         =   "Payment #"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   15
         Top             =   420
         Width           =   1095
      End
   End
   Begin VB.Frame Frame 
      Caption         =   "Monthly Payment"
      Height          =   1155
      Index           =   1
      Left            =   120
      TabIndex        =   7
      Top             =   2280
      Width           =   2055
      Begin VB.Label lblMonthlyPayment 
         Height          =   495
         Left            =   240
         TabIndex        =   8
         Top             =   420
         Width           =   1455
      End
   End
   Begin VB.Frame Frame 
      Caption         =   "Enter Values"
      Height          =   1755
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   300
      Width           =   6495
      Begin VB.TextBox txtYears 
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   4620
         TabIndex        =   2
         Top             =   780
         Width           =   1395
      End
      Begin VB.TextBox txtYrlyRate 
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   2280
         TabIndex        =   1
         Top             =   780
         Width           =   1635
      End
      Begin VB.TextBox txtLoanAmount 
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   180
         TabIndex        =   0
         Top             =   780
         Width           =   1455
      End
      Begin VB.Label Label 
         Alignment       =   2  'Center
         Caption         =   "Years"
         Height          =   375
         Index           =   2
         Left            =   4620
         TabIndex        =   6
         Top             =   480
         Width           =   1395
      End
      Begin VB.Label Label 
         Alignment       =   2  'Center
         Caption         =   "Yearly Interest Rate "
         Height          =   315
         Index           =   1
         Left            =   2280
         TabIndex        =   5
         Top             =   480
         Width           =   1635
      End
      Begin VB.Label Label 
         Alignment       =   2  'Center
         Caption         =   "Loan Amount"
         Height          =   315
         Index           =   0
         Left            =   180
         TabIndex        =   4
         Top             =   480
         Width           =   1335
      End
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Caption         =   "Yearly Extra"
      Height          =   255
      Index           =   11
      Left            =   5460
      TabIndex        =   26
      Top             =   2160
      Width           =   1155
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Caption         =   "Monthly Extra"
      Height          =   255
      Index           =   10
      Left            =   4140
      TabIndex        =   25
      Top             =   2160
      Width           =   1155
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Caption         =   "One Time Extra"
      Height          =   255
      Index           =   9
      Left            =   2700
      TabIndex        =   24
      Top             =   2160
      Width           =   1215
   End
End
Attribute VB_Name = "frmAmortizationTableProject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim AmortTable(360) As String

Private Sub cmdCalculate_Click()
'-Variable declarations
Dim LoanAmount As Currency, MonthlyPayment As Currency
Dim YrlyRate As Single, MonthlyRate As Single
Dim Years As Integer, Payments As Integer
Dim OneTimeExtra, MonthlyExtra, YearlyExtra As Double
'-Reading values from the form
YearlyExtra = Val(txtYearlyExtra)
MonthlyExtra = Val(txtMonthlyExtra)
OneTimeExtra = Val(txtOneTimeExtra)
LoanAmount = Val(txtLoanAmount)
YrlyRate = Val(txtYrlyRate)
Years = Val(txtYears)
'-Intermediate Calculations
MonthlyRate = YrlyRate / 1200
Payments = Years * 12
'-Monthly Payment
MonthlyPayment = (LoanAmount * MonthlyRate / (1 - (1 + MonthlyRate) ^ (-Payments))) + MonthlyExtra
'-Display results
lblMonthlyPayment = Format$(MonthlyPayment, "Currency")
'-Additional local declarations
Dim PaymentNumber As Integer
Dim MonthlyInt, MonthlyPrinciple As Currency
Dim TotalInt As Currency, CurrentAmt As Currency
Dim YearNumber As Integer, Displine As String
'-Initialize the TotalInterest And CurrentBalance
TotalInt = 0
CurrentAmt = LoanAmount
'-Set up loop
For PaymentNumber = 1 To Payments
    If CurrentAmt = 0 Or CurrentAmt < 0 Then
        Exit For
    End If
    '-Make Calculations
    If PaymentNumber = 1 Then
        CurrentAmt = CurrentAmt - OneTimeExtra
    End If
    MonthlyInt = CurrentAmt * MonthlyRate
    TotalInt = TotalInt + MonthlyInt
    CurrentAmt = CurrentAmt + MonthlyInt - MonthlyPayment
    If PaymentNumber Mod 12 = 0 Then
        YearNumber = PaymentNumber \ 12
        CurrentAmt = CurrentAmt - YearlyExtra
    Else
        YearNumber = (PaymentNumber \ 12) + 1
    End If
    MonthlyPrinciple = MonthlyPayment - MonthlyInt
    '-Build display line
    Displine = vbTab & Format$(PaymentNumber, "####")
    Displine = Displine & vbTab & Format$(YearNumber, "#0")
    Displine = Displine & vbTab & Format$(CurrentAmt, "Currency")
    Displine = Displine & vbTab & Format$(TotalInt, "Currency")
    Displine = Displine & vbTab & Format$(MonthlyInt, "Currency")
    Displine = Displine & vbTab & Format$(MonthlyPrinciple, "Currency")
    '-Transfer display line to array element
    AmortTable(PaymentNumber) = Displine
Next PaymentNumber  'end of loop
'-Set up scroll bar
hsbPayment.Min = 1
hsbPayment.Max = PaymentNumber - 1
hsbPayment.LargeChange = 12 ' one year = 12 payments
hsbPayment.Value = 1
'-Put first line of table into the textbox
txtAmortTable = AmortTable(1)

    
End Sub

Private Sub cmdClear_Click()
'-Clearing the Display labels
lblMonthlyPayment = ""
txtLoanAmount.Text = ""
txtYrlyRate.Text = ""
txtYears.Text = ""
txtOneTimeExtra.Text = ""
txtMonthlyExtra.Text = ""
txtYearlyExtra.Text = ""
txtAmortTable.Text = ""
'-Set focus back to the Loan Amount
txtLoanAmount.SetFocus
End Sub

Private Sub cmdExit_Click()
End
End Sub

Private Sub hsbPayment_Change()
txtAmortTable = AmortTable(hsbPayment.Value)
End Sub
