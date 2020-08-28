VERSION 5.00
Begin VB.Form frmPrimeNumbers 
   Caption         =   "The Prime Number Project"
   ClientHeight    =   6375
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11385
   LinkTopic       =   "Form1"
   ScaleHeight     =   6375
   ScaleWidth      =   11385
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtPFactors2 
      Height          =   795
      Left            =   6420
      MultiLine       =   -1  'True
      TabIndex        =   39
      Top             =   3960
      Width           =   3135
   End
   Begin VB.TextBox txtPFactors 
      Height          =   855
      Left            =   6480
      MultiLine       =   -1  'True
      TabIndex        =   37
      Top             =   1080
      Width           =   3075
   End
   Begin VB.ListBox lstPrimeList 
      Height          =   3765
      Left            =   9840
      TabIndex        =   30
      Top             =   1080
      Width           =   855
   End
   Begin VB.ListBox lstDivisorList2 
      Height          =   840
      Left            =   5160
      TabIndex        =   18
      Top             =   3960
      Width           =   1095
   End
   Begin VB.TextBox txtDivisorList2 
      Height          =   1335
      Left            =   2280
      MultiLine       =   -1  'True
      TabIndex        =   16
      Top             =   3960
      Width           =   2595
   End
   Begin VB.TextBox txtNumber2 
      Height          =   795
      Left            =   480
      TabIndex        =   14
      Top             =   3960
      Width           =   1695
   End
   Begin VB.ListBox lstDivisorList 
      Height          =   840
      Left            =   5100
      TabIndex        =   8
      Top             =   1080
      Width           =   1155
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   8400
      TabIndex        =   5
      Top             =   5460
      Width           =   1035
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   675
      Left            =   6720
      TabIndex        =   4
      Top             =   5220
      Width           =   1035
   End
   Begin VB.TextBox txtDivisorList 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Left            =   2460
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   1080
      Width           =   2475
   End
   Begin VB.TextBox txtNumber 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   420
      TabIndex        =   0
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Prime Factorization"
      Height          =   375
      Index           =   13
      Left            =   6420
      TabIndex        =   38
      Top             =   3540
      Width           =   1395
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Prime Factorization"
      Height          =   375
      Index           =   12
      Left            =   6480
      TabIndex        =   36
      Top             =   600
      Width           =   1395
   End
   Begin VB.Label lblPerfectSquare2 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Perfect Square!"
      Height          =   375
      Left            =   480
      TabIndex        =   35
      Top             =   5280
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblPerfectNumber2 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Perfect Number!"
      Height          =   375
      Left            =   480
      TabIndex        =   34
      Top             =   4860
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblPerfectSquare 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Perfect Square!"
      Height          =   375
      Left            =   420
      TabIndex        =   33
      Top             =   2460
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label lblPerfectNumber 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Perfect Number!"
      Height          =   375
      Left            =   420
      TabIndex        =   32
      Top             =   2040
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Primes"
      Height          =   375
      Left            =   9840
      TabIndex        =   31
      Top             =   600
      Width           =   855
   End
   Begin VB.Label lblLCM 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   9120
      TabIndex        =   29
      Top             =   2700
      Width           =   435
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "LCM:"
      Height          =   315
      Index           =   11
      Left            =   8460
      TabIndex        =   28
      Top             =   2700
      Width           =   615
   End
   Begin VB.Label lblGCF 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   9120
      TabIndex        =   27
      Top             =   2280
      Width           =   435
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "GCF:"
      Height          =   315
      Index           =   10
      Left            =   8460
      TabIndex        =   26
      Top             =   2280
      Width           =   615
   End
   Begin VB.Label lblNotPrime2 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "It is not prime!"
      Height          =   375
      Left            =   2400
      TabIndex        =   25
      Top             =   5460
      Width           =   1095
   End
   Begin VB.Label lblNotPrime 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "It is not prime!"
      Height          =   375
      Left            =   2400
      TabIndex        =   24
      Top             =   2580
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.Label lblSum2 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   5280
      TabIndex        =   23
      Top             =   5700
      Width           =   495
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Sum:"
      Height          =   315
      Index           =   9
      Left            =   4740
      TabIndex        =   22
      Top             =   5700
      Width           =   495
   End
   Begin VB.Label lblNumDivisors2 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   5280
      TabIndex        =   21
      Top             =   5340
      Width           =   495
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "# Divisors:"
      Height          =   315
      Index           =   8
      Left            =   4380
      TabIndex        =   20
      Top             =   5340
      Width           =   855
   End
   Begin VB.Label lblPrime2 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "It is prime!"
      Height          =   315
      Left            =   2400
      TabIndex        =   19
      Top             =   5460
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Divisors"
      Height          =   375
      Index           =   7
      Left            =   5160
      TabIndex        =   17
      Top             =   3540
      Width           =   1095
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Divisors"
      Height          =   375
      Index           =   6
      Left            =   2280
      TabIndex        =   15
      Top             =   3540
      Width           =   1935
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Enter a Number"
      Height          =   375
      Index           =   5
      Left            =   480
      TabIndex        =   13
      Top             =   3540
      Width           =   1695
   End
   Begin VB.Label lblSum 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   5280
      TabIndex        =   12
      Top             =   2820
      Width           =   495
   End
   Begin VB.Label lblNumDivisors 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   5280
      TabIndex        =   11
      Top             =   2460
      Width           =   495
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Sum:"
      Height          =   315
      Index           =   4
      Left            =   4680
      TabIndex        =   10
      Top             =   2820
      Width           =   555
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "# Divisors:"
      Height          =   315
      Index           =   3
      Left            =   4380
      TabIndex        =   9
      Top             =   2460
      Width           =   855
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Divisors"
      Height          =   375
      Index           =   2
      Left            =   5100
      TabIndex        =   7
      Top             =   600
      Width           =   1155
   End
   Begin VB.Label lblPrime 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "It is prime!"
      Height          =   375
      Left            =   2400
      TabIndex        =   6
      Top             =   2580
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Divisors"
      Height          =   375
      Index           =   1
      Left            =   2400
      TabIndex        =   3
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Enter a Number"
      Height          =   375
      Index           =   0
      Left            =   420
      TabIndex        =   1
      Top             =   600
      Width           =   1815
   End
End
Attribute VB_Name = "frmPrimeNumbers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim NumDivisors, DivisorSum, NumDivisors2, DivisorSum2 As Integer
Dim Number, Number2 As Long
Dim Sum, Sum2 As Long
Dim Divisor, Divisor2, Divisor3 As Long
Dim GCF, LargerNumber As Long
Dim DivisorList, DivisorList2 As String


Private Sub cmdClear_Click()

txtNumber.Text = "" 'Clear textbox
txtDivisorList.Text = "" 'Clear Divisor lisst
lblPrime.Visible = False 'Makes labels invisible
lblNotPrime.Visible = False
txtNumber.SetFocus 'Puts focus on textbox

txtNumber2.Text = ""
txtDivisorList2.Text = ""
lblPrime2.Visible = False
lblNotPrime2.Visible = False

lstDivisorList.Clear
NumDivisors = 0
lblNumDivisors = NumDivisors
DivisorSum = 0
lblSum = DivisorSum
Number = 0

lstDivisorList2.Clear
NumDivisors2 = 0
lblNumDivisors2 = NumDivisors2
DivisorSum2 = 0
lblSum2 = DivisorSum2
Number2 = 0

lstPrimeList.Clear

'Resets everything

GCF = 0
lblGCF = GCF
lblLCM = "0"

txtPFactors.Text = ""
txtPFactors2.Text = ""

lblPerfectNumber.Visible = False
lblPerfectSquare.Visible = False
lblPerfectNumber2.Visible = False
lblPerfectSquare2.Visible = False

End Sub

Private Sub cmdExit_Click()

End

End Sub

Private Sub txtNumber_KeyPress(KeyAscii As Integer)

Number = Val(txtNumber) 'read number

If KeyAscii = 13 And Number > 0 Then

    Sum = 0 'initialize list
    DivisorList = ""
    For Divisor = 1 To Number 'loop of divisors
    If Number Mod Divisor = 0 Then 'test
        Sum = Sum + Divisor 'running total
        DivisorList = DivisorList & Str$(Divisor) & ", "
        lstDivisorList.AddItem (Str$(Divisor))
        NumDivisors = NumDivisors + 1
        lblNumDivisors = NumDivisors
        DivisorSum = DivisorSum + Divisor
    End If
    Next Divisor
    lblSum = DivisorSum
    txtDivisorList.Text = DivisorList 'display
    If Sum = Number + 1 Then 'test if prime
        lblPrime.Visible = True
    Else
        lblNotPrime.Visible = True
    End If
    
    Call GCFLCM
    
    If DivisorSum = Number * 2 Then 'Perfect Number test
        lblPerfectNumber.Visible = True
    End If
    
    If Sqr(Number) = Int(Sqr(Number)) Then 'Perfect Square test
        lblPerfectSquare.Visible = True
    End If
    
    Dim i, Divisor4, Sum3 As Integer
    Dim Check As Boolean
    Dim Number3, PFactors As String
    
    Number3 = Val(txtNumber)
    
    PFactors = ""
    
    If lblPrime.Visible = False Then
        For i = 1 To Number3 'Goes from 1 to number
            If Number3 Mod i = 0 Then 'Checks if number inputed(Number3) divides by divisor(i), evenly
                'Checks if divisor if a prime number
                Sum3 = 0
                For Divisor4 = 2 To i
                    If i Mod Divisor4 = 0 Then
                        Sum3 = Sum3 + Divisor4
                    End If
                Next Divisor4
    
                If Sum3 <> i + 1 Then
                'If the divisor is prime, it does the following
                    Number3 = Number3 / i 'Sets number equal the number divided by the divisor
                    PFactors = PFactors & Str$(i) & " * " 'Adds Divisor to text box
                    'resets variables so loop can run again
                    i = 1
                    Divisor4 = 2
                Else
                    i = i + 1 'Adds one to i so that the loop goes to the next divisor
                End If
            End If
        Next i
    End If
    txtPFactors.Text = PFactors 'Adds divisors to text Box
    txtNumber2.SetFocus 'shift focus to second box
End If

End Sub

Private Sub txtNumber2_KeyPress(KeyAscii As Integer)

Number2 = Val(txtNumber2)
Number = Val(txtNumber)
    
If KeyAscii = 13 And Number2 > 0 And Number > 0 Then
    Sum2 = 0 'initialize list
    DivisorList2 = ""
    For Divisor2 = 1 To Number2 'loop of divisors
    If Number2 Mod Divisor2 = 0 Then 'test
        Sum2 = Sum2 + Divisor2 'running total
        DivisorList2 = DivisorList2 & Str$(Divisor2) & ", "
        lstDivisorList2.AddItem (Str$(Divisor2))
        NumDivisors2 = NumDivisors2 + 1
        lblNumDivisors2 = NumDivisors2
        DivisorSum2 = DivisorSum2 + Divisor2
    End If
    Next Divisor2
    lblSum2 = DivisorSum2
    txtDivisorList2.Text = DivisorList2 'display
    If Sum2 = Number2 + 1 Then 'test if prime
        lblPrime2.Visible = True
    Else
        lblNotPrime2.Visible = True
    End If
    
    Call GCFLCM
    Call PrimeList
        
    If DivisorSum2 = Number2 * 2 Then
        lblPerfectNumber2.Visible = True
    End If
    
    If Sqr(Number2) = Int(Sqr(Number2)) Then
        lblPerfectSquare2.Visible = True
    End If
    
    Dim i2, Divisor5, Sum4 As Integer
    Dim Check2 As Boolean
    Dim Number4, PFactors2 As String
    
    Number4 = Val(txtNumber2)
    
    PFactors2 = ""
    
    If lblPrime2.Visible = False Then
        For i2 = 1 To Number4
            If Number4 Mod i2 = 0 Then
                Sum4 = 0
                For Divisor5 = 2 To i2
                    If i2 Mod Divisor5 = 0 Then
                        Sum4 = Sum4 + Divisor5
                    End If
                Next Divisor5
                
                If Sum4 <> i2 + 12 Then
                    Number4 = Number4 / i2
                    PFactors2 = PFactors2 & Str$(i2) & " * "
                    i2 = 1
                    Divisor5 = 2
                Else
                    i2 = i2 + 1
                End If
            End If
        Next i2
    End If
    txtPFactors2.Text = PFactors2
    cmdClear.SetFocus 'shift focus to clear
End If

End Sub

Public Function GCFLCM()

If Number <> 0 And Number2 <> 0 Then
For Divisor3 = 1 To Number2
    If (Number Mod Divisor3) = 0 And (Number2 Mod Divisor3) = 0 Then
        GCF = Divisor3
    End If
Next Divisor3
lblGCF = GCF
lblLCM = (Number * Number2) / GCF
End If

End Function

Public Function PrimeList()

Dim Count, CountPrime, PrimeDivisor As Integer

If Number > Number2 Then
    LargerNumber = Number
ElseIf Number2 > Number Then
    LargerNumber = Number2
Else
    LargerNumber = Number
End If

For Count = 1 To LargerNumber 'Counts up to largest number
    CountPrime = 0 'initializes CountPrime = 0
        For PrimeDivisor = 1 To Count 'loop of divisors
            If Count Mod PrimeDivisor = 0 Then 'Divides all #s up to count
                CountPrime = CountPrime + 1 'Adds +1 if Mod 0
            End If
        Next PrimeDivisor

If CountPrime = 2 Then 'if there are 2 divisors then
    lstPrimeList.AddItem (Str$(Count))
End If

Next Count

End Function
