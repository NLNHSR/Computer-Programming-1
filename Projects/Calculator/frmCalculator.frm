VERSION 5.00
Begin VB.Form frmCalculator 
   Caption         =   "Calculator"
   ClientHeight    =   8460
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   17775
   LinkTopic       =   "Form1"
   Picture         =   "frmCalculator.frx":0000
   ScaleHeight     =   8460
   ScaleWidth      =   17775
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDecimal 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7620
      TabIndex        =   35
      Top             =   5700
      Width           =   600
   End
   Begin VB.CommandButton cmdNegative 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   16.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9000
      TabIndex        =   34
      Top             =   5700
      Width           =   600
   End
   Begin VB.CommandButton cmd0 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8340
      TabIndex        =   33
      Top             =   5700
      Width           =   600
   End
   Begin VB.CommandButton cmd9 
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9000
      TabIndex        =   32
      Top             =   5100
      Width           =   600
   End
   Begin VB.CommandButton cmd8 
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8340
      TabIndex        =   31
      Top             =   5100
      Width           =   600
   End
   Begin VB.CommandButton cmd7 
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7620
      TabIndex        =   30
      Top             =   5100
      Width           =   600
   End
   Begin VB.CommandButton cmd6 
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9000
      TabIndex        =   29
      Top             =   4440
      Width           =   600
   End
   Begin VB.CommandButton cmd5 
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8340
      TabIndex        =   28
      Top             =   4440
      Width           =   600
   End
   Begin VB.CommandButton cmd4 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7620
      TabIndex        =   27
      Top             =   4440
      Width           =   600
   End
   Begin VB.CommandButton cmd3 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9000
      TabIndex        =   26
      Top             =   3780
      Width           =   600
   End
   Begin VB.CommandButton cmd2 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8340
      TabIndex        =   25
      Top             =   3780
      Width           =   600
   End
   Begin VB.CommandButton cmd1 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7620
      TabIndex        =   24
      Top             =   3780
      Width           =   600
   End
   Begin VB.CommandButton cmdAverage 
      Caption         =   "Avg"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6540
      TabIndex        =   22
      Top             =   3780
      Width           =   608
   End
   Begin VB.CommandButton cmdLogarithm 
      Caption         =   "Log(X)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5940
      TabIndex        =   21
      Top             =   5700
      Width           =   1215
   End
   Begin VB.CommandButton cmdCubeRoot 
      Caption         =   "Cbrt(X)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5940
      TabIndex        =   20
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton cmdFactorial 
      Caption         =   "X!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5820
      TabIndex        =   19
      Top             =   3780
      Width           =   608
   End
   Begin VB.CommandButton cmdTan 
      Caption         =   "Tan(X)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      TabIndex        =   18
      Top             =   5700
      Width           =   1215
   End
   Begin VB.CommandButton cmdCos 
      Caption         =   "Cos(X)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5940
      TabIndex        =   17
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton cmdSin 
      Caption         =   "Sin(X)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      TabIndex        =   16
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton cmdAbs 
      Caption         =   "Abs(X)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5160
      TabIndex        =   15
      Top             =   6360
      Width           =   1215
   End
   Begin VB.CommandButton cmdExp 
      Caption         =   "^"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   16.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      TabIndex        =   14
      Top             =   3780
      Width           =   608
   End
   Begin VB.CommandButton cmdSqrt 
      BackColor       =   &H80000005&
      Caption         =   "Sqrt(X)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      TabIndex        =   13
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton cmdSq 
      Caption         =   "X^2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5040
      TabIndex        =   12
      Top             =   3780
      Width           =   608
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   16.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11640
      TabIndex        =   8
      Top             =   7080
      Width           =   1215
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "CE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8040
      TabIndex        =   7
      Top             =   6360
      Width           =   1215
   End
   Begin VB.CommandButton cmdMultiply 
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5820
      TabIndex        =   6
      Top             =   3120
      Width           =   608
   End
   Begin VB.CommandButton cmdDivide 
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6540
      TabIndex        =   5
      Top             =   3120
      Width           =   608
   End
   Begin VB.CommandButton cmdSubtract 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   16.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5040
      TabIndex        =   4
      Top             =   3120
      Width           =   608
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      TabIndex        =   3
      Top             =   3120
      Width           =   608
   End
   Begin VB.TextBox txtNum2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5940
      TabIndex        =   1
      Top             =   2340
      Width           =   1215
   End
   Begin VB.TextBox txtNum1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      TabIndex        =   0
      Top             =   2340
      Width           =   1215
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Caption         =   "*Y Left Blank Will Count As Zero"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Index           =   4
      Left            =   10020
      TabIndex        =   36
      Top             =   6960
      Width           =   1515
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Calculator"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   675
      Index           =   3
      Left            =   4320
      TabIndex        =   23
      Top             =   720
      Width           =   8535
   End
   Begin VB.Image Image 
      BorderStyle     =   1  'Fixed Single
      Height          =   2775
      Index           =   1
      Left            =   10020
      Picture         =   "frmCalculator.frx":14B366
      Stretch         =   -1  'True
      Top             =   1620
      Width           =   2835
   End
   Begin VB.Image Image 
      BorderStyle     =   1  'Fixed Single
      Height          =   2295
      Index           =   0
      Left            =   10020
      Picture         =   "frmCalculator.frx":1677CE
      Stretch         =   -1  'True
      Top             =   4560
      Width           =   2835
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Answer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   16.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   7620
      TabIndex        =   11
      Top             =   1620
      Width           =   2000
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Y"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   5940
      TabIndex        =   10
      Top             =   1620
      Width           =   1215
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   4320
      TabIndex        =   9
      Top             =   1620
      Width           =   1215
   End
   Begin VB.Label lblAns 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1275
      Left            =   7620
      TabIndex        =   2
      Top             =   2340
      Width           =   1995
   End
End
Attribute VB_Name = "frmCalculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Num1, Num2, Ans As Double
Dim Num1Focus, Num2Focus As Boolean

Private Sub cmd0_Click()

If Num2Focus = True Then
    txtNum2.Text = txtNum2.Text & 0
ElseIf Num1Focus = True Then
    txtNum1.Text = txtNum1.Text & 0
End If

End Sub

Private Sub cmd1_Click()

If Num2Focus = True Then
    txtNum2.Text = txtNum2.Text & 1
ElseIf Num1Focus = True Then
    txtNum1.Text = txtNum1.Text & 1
End If

End Sub

Private Sub cmd2_Click()

If Num2Focus = True Then
    txtNum2.Text = txtNum2.Text & 2
ElseIf Num1Focus = True Then
    txtNum1.Text = txtNum1.Text & 2
End If

End Sub

Private Sub cmd3_Click()

If Num2Focus = True Then
    txtNum2.Text = txtNum2.Text & 3
ElseIf Num1Focus = True Then
    txtNum1.Text = txtNum1.Text & 3
End If

End Sub

Private Sub cmd4_Click()

If Num2Focus = True Then
    txtNum2.Text = txtNum2.Text & 4
ElseIf Num1Focus = True Then
    txtNum1.Text = txtNum1.Text & 4
End If

End Sub

Private Sub cmd5_Click()

If Num2Focus = True Then
    txtNum2.Text = txtNum2.Text & 5
ElseIf Num1Focus = True Then
    txtNum1.Text = txtNum1.Text & 5
End If

End Sub

Private Sub cmd6_Click()

If Num2Focus = True Then
    txtNum2.Text = txtNum2.Text & 6
ElseIf Num1Focus = True Then
    txtNum1.Text = txtNum1.Text & 6
End If

End Sub

Private Sub cmd7_Click()

If Num2Focus = True Then
    txtNum2.Text = txtNum2.Text & 7
ElseIf Num1Focus = True Then
    txtNum1.Text = txtNum1.Text & 7
End If

End Sub

Private Sub cmd8_Click()

If Num2Focus = True Then
    txtNum2.Text = txtNum2.Text & 8
ElseIf Num1Focus = True Then
    txtNum1.Text = txtNum1.Text & 8
End If

End Sub

Private Sub cmd9_Click()

If Num2Focus = True Then
    txtNum2.Text = txtNum2.Text & 9
ElseIf Num1Focus = True Then
    txtNum1.Text = txtNum1.Text & 9
End If

End Sub

Private Sub cmdAbs_Click()

Num1 = Val(txtNum1)

If txtNum1 = "" Then
    lblAns = "Please Enter A Value For X"
Else
    Ans = Math.Abs(Num1)
    lblAns = Math.Round(Ans, 5)
End If

Num1Focus = False
Num2Focus = False

End Sub

Private Sub cmdAdd_Click()

Num1 = Val(txtNum1)
Num2 = Val(txtNum2)

If txtNum1 = "" Then
    lblAns = "Please Enter A Value For X"
Else
    Ans = Num1 + Num2
    lblAns = Math.Round(Ans, 5)
End If

Num1Focus = False
Num2Focus = False

End Sub

Private Sub cmdAverage_Click()

Num1 = Val(txtNum1)
Num2 = Val(txtNum2)

If txtNum1 = "" Then
    lblAns = "Please Enter A Value For X"
Else
    Ans = (Num1 + Num2) / 2
    lblAns = Math.Round(Ans, 5)
End If

Num1Focus = False
Num2Focus = False

End Sub

Private Sub cmdClear_Click()

txtNum1 = ""
txtNum2 = ""

lblAns = ""
Num1Focus = False
Num2Focus = False

End Sub

Private Sub cmdCos_Click()

Num1 = Val(txtNum1)


If txtNum1 = "" Then
    lblAns = "Please Enter A Value For X"
Else
    Ans = Math.Cos(Num1)
    lblAns = Math.Round(Ans, 5)
End If

Num1Focus = False
Num2Focus = False

End Sub

Private Sub cmdCubeRoot_Click()

Num1 = Val(txtNum1)

If txtNum1 = "" Then
    lblAns = "Please Enter A Value For X"
Else
    If Num1 > 0 Then
        Ans = Num1 ^ (1 / 3)
        lblAns = Math.Round(Ans, 5)
    ElseIf Num1 < 0 Then
        Num1 = Math.Abs(Num1)
        Ans = Num1 ^ (1 / 3)
        lblAns = "-" & (Math.Round(Ans, 5))
    End If
End If

Num1Focus = False
Num2Focus = False

End Sub

Private Sub cmdDecimal_Click()

If Num2Focus = True Then
    txtNum2.Text = txtNum2.Text & "."
ElseIf Num1Focus = True Then
    txtNum1.Text = txtNum1.Text & "."
End If

End Sub

Private Sub cmdDivide_Click()

Num1 = Val(txtNum1)
Num2 = Val(txtNum2)

If txtNum1 = "" Then
    lblAns = "Please Enter A Value For X"
Else
    If Num2 = 0 Then
        lblAns = "Error: Undefined"
    Else
        Ans = Num1 / Num2
        lblAns = Math.Round(Ans, 5)
    End If
End If

Num1Focus = False
Num2Focus = False

End Sub

Private Sub cmdExp_Click()

Num1 = Val(txtNum1)
Num2 = Val(txtNum2)

If txtNum1 = "" Then
    lblAns = "Please Enter A Value For X"
Else
    Ans = Num1 ^ Num2
    lblAns = Math.Round(Ans, 5)
End If

Num1Focus = False
Num2Focus = False

End Sub

Private Sub cmdFactorial_Click()

Num1 = Val(txtNum1)

If txtNum1 = "" Then
    lblAns = "Please Enter A Value For X"
Else
    If Num1 > 0 Then
        Ans = Factorial(Val(txtNum1))
        lblAns = Math.Round(Ans, 5)
    ElseIf Num1 < 0 Then
        Ans = Factorial(Math.Abs(Val(txtNum1)))
        lblAns = "-" & (Math.Round(Ans, 5))
    End If
End If

Num1Focus = False
Num2Focus = False

End Sub

Private Sub cmdLogarithm_Click()

Num1 = Val(txtNum1)

If txtNum1 = "" Then
    lblAns = "Please Enter A Value For X"
Else
    If Num1 < 0 Then
        lblAns = "Error: Negative Base"
    ElseIf Num1 = 0 Then
        lblAns = "Error: Undefined"
    Else
        Ans = Math.Log(Num1)
        lblAns = Math.Round(Ans, 5)
    End If
End If

Num1Focus = False
Num2Focus = False

End Sub

Private Sub cmdMultiply_Click()

Num1 = Val(txtNum1)
Num2 = Val(txtNum2)

If txtNum1 = "" Then
    lblAns = "Please Enter A Value For X"
Else
    Ans = Num1 * Num2
    lblAns = Math.Round(Ans, 5)
End If

Num1Focus = False
Num2Focus = False

End Sub

Private Sub cmdNegative_Click()

If Num2Focus = True Then
    txtNum2.Text = txtNum2.Text & "-"
ElseIf Num1Focus = True Then
    txtNum1.Text = txtNum1.Text & "-"
End If

End Sub

Private Sub cmdQuit_Click()

End

End Sub

Private Sub cmdSin_Click()

Num1 = Val(txtNum1)

If txtNum1 = "" Then
    lblAns = "Please Enter A Value For X"
Else
    Ans = Math.Sin(Num1)
    lblAns = Math.Round(Ans, 5)
End If

Num1Focus = False
Num2Focus = False

End Sub

Private Sub cmdSq_Click()

Num1 = Val(txtNum1)

If txtNum1 = "" Then
    lblAns = "Please Enter A Value For X"
Else
    Ans = Num1 ^ 2
    lblAns = Math.Round(Ans, 5)
End If

Num1Focus = False
Num2Focus = False

End Sub

Private Sub cmdSqrt_Click()

Num1 = Val(txtNum1)

If txtNum1 = "" Then
    lblAns = "Please Enter A Value For X"
Else
    If Num1 < 0 Then
        lblAns = "Error: Negative Root"
    Else
        Ans = Sqr(Num1)
        lblAns = Math.Round(Ans, 5)
    End If
End If

Num1Focus = False
Num2Focus = False

End Sub

Private Sub cmdSubtract_Click()

Num1 = Val(txtNum1)
Num2 = Val(txtNum2)

If txtNum1 = "" Then
    lblAns = "Please Enter A Value For X"
Else
    Ans = Num1 - Num2
    lblAns = Math.Round(Ans, 5)
End If

Num1Focus = False
Num2Focus = False

End Sub

Private Sub cmdTan_Click()

Num1 = Val(txtNum1)


If txtNum1 = "" Then
    lblAns = "Please Enter A Value For X"
Else
    Ans = Math.Tan(Num1)
    lblAns = Math.Round(Ans, 5)
End If

Num1Focus = False
Num2Focus = False

End Sub

Public Function Factorial(a As Integer) As Long

Dim i As Integer
Factorial = 1
For i = 1 To a
Factorial = Factorial * i
Next

End Function

Private Sub txtNum1_GotFocus()

Num1Focus = True
Num2Focus = False

End Sub

Private Sub txtNum2_GotFocus()

Num2Focus = True
Num1Focus = False

End Sub
