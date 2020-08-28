VERSION 5.00
Begin VB.Form frmTwoPointProblem 
   BackColor       =   &H00000005&
   Caption         =   "Two Point "
   ClientHeight    =   7515
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9285
   LinkTopic       =   "Form1"
   ScaleHeight     =   7515
   ScaleWidth      =   9285
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   375
      Left            =   7980
      TabIndex        =   49
      Top             =   7020
      Width           =   1095
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   375
      Left            =   6720
      TabIndex        =   48
      Top             =   7020
      Width           =   1095
   End
   Begin VB.CommandButton cmdCalculate 
      Caption         =   "Calculate"
      Height          =   375
      Left            =   5460
      TabIndex        =   47
      Top             =   7020
      Width           =   1095
   End
   Begin VB.PictureBox pic1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000E&
      Height          =   6735
      Left            =   2340
      ScaleHeight     =   6675
      ScaleWidth      =   6675
      TabIndex        =   46
      Top             =   180
      Width           =   6735
      Begin VB.Label lblP4 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Height          =   195
         Left            =   1740
         TabIndex        =   55
         Top             =   4140
         Visible         =   0   'False
         Width           =   45
      End
      Begin VB.Label lblP3 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Height          =   195
         Left            =   540
         TabIndex        =   54
         Top             =   2940
         Visible         =   0   'False
         Width           =   45
      End
      Begin VB.Label lblP2 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Height          =   195
         Left            =   2220
         TabIndex        =   53
         Top             =   360
         Visible         =   0   'False
         Width           =   45
      End
      Begin VB.Label lblP1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   50
         Top             =   180
         Visible         =   0   'False
         Width           =   45
      End
   End
   Begin VB.TextBox txtY4 
      Height          =   315
      Left            =   1680
      TabIndex        =   19
      Top             =   2940
      Width           =   555
   End
   Begin VB.TextBox txtX4 
      Height          =   315
      Left            =   1680
      TabIndex        =   17
      Top             =   2580
      Width           =   555
   End
   Begin VB.TextBox txtY3 
      Height          =   315
      Left            =   540
      TabIndex        =   14
      Top             =   2940
      Width           =   555
   End
   Begin VB.TextBox txtX3 
      Height          =   315
      Left            =   540
      TabIndex        =   12
      Top             =   2580
      Width           =   555
   End
   Begin VB.TextBox txtY2 
      Height          =   315
      Left            =   1680
      TabIndex        =   9
      Top             =   1320
      Width           =   555
   End
   Begin VB.TextBox txtX2 
      Height          =   315
      Left            =   1680
      TabIndex        =   7
      Top             =   960
      Width           =   555
   End
   Begin VB.TextBox txtY1 
      Height          =   315
      Left            =   540
      TabIndex        =   4
      Top             =   1320
      Width           =   555
   End
   Begin VB.TextBox txtX1 
      Height          =   315
      Left            =   540
      TabIndex        =   2
      Top             =   960
      Width           =   555
   End
   Begin VB.Label lblPosition 
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   3180
      TabIndex        =   52
      Top             =   7020
      Width           =   855
   End
   Begin VB.Label Label 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Position"
      Height          =   375
      Index           =   27
      Left            =   2340
      TabIndex        =   51
      Top             =   7020
      Width           =   735
   End
   Begin VB.Label lblIntersectionPoint 
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1320
      TabIndex        =   45
      Top             =   7140
      Width           =   855
   End
   Begin VB.Label Label 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Intersection Point"
      Height          =   435
      Index           =   26
      Left            =   240
      TabIndex        =   44
      Top             =   6960
      Width           =   1035
   End
   Begin VB.Label lblEquation2 
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   1320
      TabIndex        =   43
      Top             =   6360
      Width           =   855
   End
   Begin VB.Label Label 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Equation"
      Height          =   255
      Index           =   25
      Left            =   1320
      TabIndex        =   42
      Top             =   6060
      Width           =   855
   End
   Begin VB.Label lblDistance2 
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1320
      TabIndex        =   41
      Top             =   5160
      Width           =   855
   End
   Begin VB.Label Label 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Distance"
      Height          =   255
      Index           =   24
      Left            =   1320
      TabIndex        =   40
      Top             =   4860
      Width           =   855
   End
   Begin VB.Label lblMidpoint2 
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1320
      TabIndex        =   39
      Top             =   4560
      Width           =   855
   End
   Begin VB.Label Label 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Midpoint"
      Height          =   255
      Index           =   23
      Left            =   1320
      TabIndex        =   38
      Top             =   4260
      Width           =   855
   End
   Begin VB.Label lblYIntercept2 
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1320
      TabIndex        =   37
      Top             =   3960
      Width           =   855
   End
   Begin VB.Label Label 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Y-Intercept"
      Height          =   255
      Index           =   22
      Left            =   1320
      TabIndex        =   36
      Top             =   3660
      Width           =   855
   End
   Begin VB.Label lblSlope2 
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1320
      TabIndex        =   35
      Top             =   5760
      Width           =   855
   End
   Begin VB.Label Label 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Slope"
      Height          =   255
      Index           =   21
      Left            =   1320
      TabIndex        =   34
      Top             =   5460
      Width           =   855
   End
   Begin VB.Label lblEquation 
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   240
      TabIndex        =   33
      Top             =   6360
      Width           =   855
   End
   Begin VB.Label Label 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Equation"
      Height          =   255
      Index           =   20
      Left            =   240
      TabIndex        =   32
      Top             =   6060
      Width           =   855
   End
   Begin VB.Label lblDistance 
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   240
      TabIndex        =   31
      Top             =   5160
      Width           =   855
   End
   Begin VB.Label Label 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Distance"
      Height          =   255
      Index           =   19
      Left            =   240
      TabIndex        =   30
      Top             =   4860
      Width           =   855
   End
   Begin VB.Label lblMidpoint 
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   240
      TabIndex        =   29
      Top             =   4560
      Width           =   855
   End
   Begin VB.Label Label 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Midpoint"
      Height          =   255
      Index           =   18
      Left            =   240
      TabIndex        =   28
      Top             =   4260
      Width           =   855
   End
   Begin VB.Label lblYIntercept 
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   240
      TabIndex        =   27
      Top             =   3960
      Width           =   855
   End
   Begin VB.Label Label 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Y-Intercept"
      Height          =   255
      Index           =   17
      Left            =   240
      TabIndex        =   26
      Top             =   3660
      Width           =   855
   End
   Begin VB.Label lblSlope 
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   240
      TabIndex        =   25
      Top             =   5760
      Width           =   855
   End
   Begin VB.Label Label 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Slope"
      Height          =   255
      Index           =   8
      Left            =   240
      TabIndex        =   24
      Top             =   5460
      Width           =   855
   End
   Begin VB.Label Label 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Line 2"
      Height          =   255
      Index           =   7
      Left            =   1320
      TabIndex        =   23
      Top             =   3360
      Width           =   915
   End
   Begin VB.Label Label 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Line 1"
      Height          =   255
      Index           =   6
      Left            =   240
      TabIndex        =   22
      Top             =   3360
      Width           =   855
   End
   Begin VB.Label Label 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Line 1"
      Height          =   255
      Index           =   16
      Left            =   240
      TabIndex        =   21
      Top             =   180
      Width           =   855
   End
   Begin VB.Label Label 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Line 2"
      Height          =   255
      Index           =   15
      Left            =   240
      TabIndex        =   20
      Top             =   1740
      Width           =   855
   End
   Begin VB.Label Label 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Y:"
      Height          =   315
      Index           =   14
      Left            =   1320
      TabIndex        =   18
      Top             =   2940
      Width           =   255
   End
   Begin VB.Label Label 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "X:"
      Height          =   315
      Index           =   13
      Left            =   1320
      TabIndex        =   16
      Top             =   2580
      Width           =   255
   End
   Begin VB.Label Label 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Point D"
      Height          =   315
      Index           =   12
      Left            =   1320
      TabIndex        =   15
      Top             =   2160
      Width           =   675
   End
   Begin VB.Label Label 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Y:"
      Height          =   315
      Index           =   11
      Left            =   240
      TabIndex        =   13
      Top             =   2940
      Width           =   255
   End
   Begin VB.Label Label 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "X:"
      Height          =   315
      Index           =   10
      Left            =   240
      TabIndex        =   11
      Top             =   2580
      Width           =   255
   End
   Begin VB.Label Label 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Point C"
      Height          =   315
      Index           =   9
      Left            =   240
      TabIndex        =   10
      Top             =   2160
      Width           =   675
   End
   Begin VB.Label Label 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Y:"
      Height          =   315
      Index           =   5
      Left            =   1320
      TabIndex        =   8
      Top             =   1320
      Width           =   255
   End
   Begin VB.Label Label 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "X:"
      Height          =   315
      Index           =   4
      Left            =   1320
      TabIndex        =   6
      Top             =   960
      Width           =   255
   End
   Begin VB.Label Label 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Point B"
      Height          =   315
      Index           =   3
      Left            =   1320
      TabIndex        =   5
      Top             =   540
      Width           =   675
   End
   Begin VB.Label Label 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Y:"
      Height          =   315
      Index           =   2
      Left            =   240
      TabIndex        =   3
      Top             =   1320
      Width           =   255
   End
   Begin VB.Label Label 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "X:"
      Height          =   315
      Index           =   1
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   255
   End
   Begin VB.Label Label 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Point A"
      Height          =   315
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   540
      Width           =   675
   End
End
Attribute VB_Name = "frmTwoPointProblem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim x1, x2, y1, y2, xcont, ycont, xcont2, ycont2 As Double
Dim Slope, YIntercept, Distance, MidX, MidY As Double
Dim Equation As String
Dim x12, x22, y12, y22, x2cont, y2cont, x2cont2, y2cont2 As Double
Dim Slope2, YIntercept2, Distance2, MidX2, MidY2 As Double
Dim Equation2 As String
Dim p As Integer
Dim i As Integer
Dim xintp, yintp As Double

Private Sub cmdCalculate_Click()
'-transfer text and convert to value
x1 = Val(txtX1)
y1 = Val(txtY1)
x2 = Val(txtX2)
y2 = Val(txtY2)
'-check to see if slope exists
If x2 - x1 <> 0 Then 'slope exists
    Slope = (y2 - y1) / (x2 - x1)
    lblSlope = Format$(Slope, "Fixed")
'-y-intercept
    YIntercept = y1 - Slope * x1
    lblYIntercept = Format$(YIntercept, "Fixed")
'-equation
    Equation = "y = " & Format$(Slope, "Fixed") & "x + " & Format$(YIntercept, "Fixed")
    lblEquation = Equation
Else
    lblSlope = "Undefined"
    lblYIntercept = "None"
    lblEquation = "x = " & Str$(x1)
End If
'-distance
Distance = Sqr((x2 - x1) ^ 2 + (y2 - y1) ^ 2)
lblDistance = Format$(Distance, "Fixed")
'-midpoint
MidX = (x1 + x2) / 2
MidY = (y1 + y2) / 2
lblMidpoint = "(" & Str$(MidX) & ", " & Str$(MidY) & ")"
'-plot points
pic1.Circle (x1, y1), 0.25, vbGreen
pic1.Circle (x2, y2), 0.25, vbGreen

If Abs(Slope) = Slope Then 'Checks if slope if positive
    xcont = x1 + ((x2 - x1) * 100)
    ycont = y1 + ((y2 - y1) * 100)
    x2cont = x2 - ((x2 - x1) * 100)
    y2cont = y2 - ((y2 - y1) * 100)
    pic1.Line (xcont, ycont)-(x2cont, y2cont)
ElseIf Abs(Slope) <> Slope Then 'Checks if slope is negative
    xcont = x1 - ((x2 - x1) * 100)
    ycont = y1 - ((y2 - y1) * 100)
    x2cont = x2 + ((x2 - x1) * 100)
    y2cont = y2 + ((y2 - y1) * 100)
    pic1.Line (xcont, ycont)-(x2cont, y2cont)
End If

lblP1.Left = x1 + 0.5
lblP1.Top = y1 - 0.5
lblP1 = Format(x1, "Fixed") + "," + Format(y1, "Fixed")
lblP1.Visible = True
lblP2.Left = x2 + 0.5
lblP2.Top = y2 - 0.5
lblP2 = Format(x2, "Fixed") + "," + Format(y2, "Fixed")
lblP2.Visible = True
'-transfer text and convert to value
x12 = Val(txtX3)
y12 = Val(txtY3)
x22 = Val(txtX4)
y22 = Val(txtY4)
'-check to see if slope exists
If x22 - x12 <> 0 Then 'slope exists
    Slope2 = (y22 - y12) / (x22 - x12)
    lblSlope2 = Format$(Slope2, "Fixed")
'-y-intercept
    YIntercept2 = y12 - Slope2 * x12
    lblYIntercept2 = Format$(YIntercept2, "Fixed")
'-equation
    Equation2 = "y = " & Format$(Slope2, "Fixed") & "x + " & Format$(YIntercept2, "Fixed")
    lblEquation2 = Equation2
Else
    lblSlope2 = "Undefined"
    lblYIntercept2 = "None"
    lblEquation2 = "x = " & Str$(x12)
End If
'-distance
Distance2 = Sqr((x22 - x12) ^ 2 + (y22 - y12) ^ 2)
lblDistance2 = Format$(Distance2, "Fixed")
'-midpoint
MidX2 = (x12 + x22) / 2
MidY2 = (y12 + y22) / 2
lblMidpoint2 = "(" & Str$(MidX2) & ", " & Str$(MidY2) & ")"
'-plot points
pic1.Circle (x12, y12), 0.25, vbGreen
pic1.Circle (x22, y22), 0.25, vbGreen
If Abs(Slope2) = Slope2 Then 'Checks if slope if positive
    xcont2 = x12 + ((x22 - x12) * 100)
    ycont2 = y12 + ((y22 - y12) * 100)
    x2cont2 = x22 - ((x22 - x12) * 100)
    y2cont2 = y22 - ((y22 - y12) * 100)
    pic1.Line (xcont2, ycont2)-(x2cont2, y2cont2)
ElseIf Abs(Slope2) <> Slope2 Then 'Checks if slope is negative
    xcont2 = x12 - ((x22 - x12) * 100)
    ycont2 = y12 - ((y22 - y12) * 100)
    x2cont2 = x22 + ((x22 - x12) * 100)
    y2cont2 = y22 + ((y22 - y12) * 100)
    pic1.Line (xcont2, ycont2)-(x2cont2, y2cont2)
End If
pic1.Line (x12, y12)-(x22, y22)
lblP3.Left = x12 + 0.5
lblP3.Top = y12 - 0.5
lblP3 = Format(x12, "Fixed") + "," + Format(y12, "Fixed")
lblP3.Visible = True
lblP4.Left = x22 + 0.5
lblP4.Top = y22 - 0.5
lblP4 = Format(x22, "Fixed") + "," + Format(y22, "Fixed")
lblP4.Visible = True
'-intersection point
If Equation = Equation2 Then
    lblIntersectionPoint.Caption = "Same Line"
ElseIf Slope = Slope2 Then
    lblIntersectionPoint.Caption = "Parallel"
Else
    xintp = (YIntercept2 - YIntercept) / (Slope - Slope2)
    yintp = (Slope * xintp) + YIntercept
    lblIntersectionPoint = "(" + Format(xintp, "Fixed") + "," + Format(yintp, "Fixed") + ")"
End If
End Sub

Private Sub cmdClear_Click()
Slope = 0
Slope2 = 0
YIntercept = 0
YIntercept2 = 0
Distance = 0
Distance2 = 0
MidX = 0
MidX2 = 0
MidY = 0
MidY2 = 0
Equation = ""
Equation2 = ""
txtX1 = ""
txtX2 = ""
txtX3 = ""
txtX4 = ""
txtY1 = ""
txtY2 = ""
txtY3 = ""
txtY4 = ""
lblYIntercept = ""
lblYIntercept2 = ""
lblMidpoint = ""
lblMidpoint2 = ""
lblDistance = ""
lblDistance2 = ""
lblSlope = ""
lblSlope2 = ""
lblEquation = ""
lblEquation2 = ""
lblIntersectionPoint = ""
pic1.Cls
Form_Activate
Form_Load
lblP1.Visible = False
lblP2.Visible = False
lblP3.Visible = False
lblP4.Visible = False

End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub Form_Activate()
pic1.Scale (-10, 10)-(10, -10) 'scale for x and y axis
pic1.Line (-10, 0)-(10, 0), vbRed 'line for x axis(horizontal line)
pic1.Line (0, -10)-(0, 10), vbBlue 'line for y axis(vertical line)
For i = -10 To 10
    pic1.Line (i, 0.5)-(i, -0.5), vbBlue 'Could change the 0.5 to 10 if you want lines to run to the edge
    pic1.Line (0.5, i)-(-0.5, i), vbRed
Next i
End Sub

Private Sub Form_Load()
p = 0
End Sub

Private Sub pic1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Pic1 - This is the picture box
p = p + 1 'This will keep a running total of your mousedown events for each point
If p = 1 Then
    x1 = X
    y1 = Y
    txtX1.Text = X
    txtY1.Text = Y
    pic1.Circle (x1, y1), 0.25, vbGreen
    lblP1.Left = x1 + 0.5
    lblP1.Top = y1 - 0.5
    lblP1 = Format(x1, "Fixed") + "," + Format(y1, "Fixed")
    lblP1.Visible = True
End If
If p = 2 Then
    x2 = X
    y2 = Y
    txtX2.Text = X
    txtY2.Text = Y
    pic1.Circle (x2, y2), 0.25, vbGreen
    pic1.Line (x1, y1)-(x2, y2)
    lblP2.Left = x2 + 0.5
    lblP2.Top = y2 - 0.5
    lblP2 = Format(x2, "Fixed") + "," + Format(y2, "Fixed")
    lblP2.Visible = True
    If Abs(Slope) = Slope Then 'Checks if slope if positive
        xcont = x1 + ((x2 - x1) * 10)
        ycont = y1 + ((y2 - y1) * 10)
        x2cont = x2 - ((x2 - x1) * 10)
        y2cont = y2 - ((y2 - y1) * 10)
    pic1.Line (xcont, ycont)-(x2cont, y2cont)
    ElseIf Abs(Slope) <> Slope Then 'Checks if slope is negative
        xcont = x1 - ((x2 - x1) * 10)
        ycont = y1 - ((y2 - y1) * 10)
        x2cont = x2 + ((x2 - x1) * 10)
        y2cont = y2 + ((y2 - y1) * 10)
        pic1.Line (xcont, ycont)-(x2cont, y2cont)
    End If
    Call cmdCalculate_Click
End If
If p = 3 Then
    x12 = X
    y12 = Y
    txtX3.Text = X
    txtY3.Text = Y
    pic1.Circle (x12, y12), 0.25, vbGreen
    lblP3.Left = x12 + 0.5
    lblP3.Top = y12 - 0.5
    lblP3 = Format(x12, "Fixed") + "," + Format(y12, "Fixed")
    lblP3.Visible = True
End If
If p = 4 Then
    x22 = X
    y22 = Y
    txtX4.Text = X
    txtY4.Text = Y
    pic1.Circle (x22, y22), 0.25, vbGreen
    pic1.Line (x12, y12)-(x22, y22)
    lblP4.Left = x22 + 0.5
    lblP4.Top = y22 - 0.5
    lblP4 = Format(x22, "Fixed") + "," + Format(y22, "Fixed")
    lblP4.Visible = True
    If Abs(Slope2) = Slope2 Then 'Checks if slope if positive
        xcont2 = x12 + ((x22 - x12) * 10)
        ycont2 = y12 + ((y22 - y12) * 10)
        x2cont2 = x22 - ((x22 - x12) * 10)
        y2cont2 = y22 - ((y22 - y12) * 10)
        pic1.Line (xcont2, ycont2)-(x2cont2, y2cont2)
    ElseIf Abs(Slope2) <> Slope2 Then 'Checks if slope is negative
        xcont2 = x12 - ((x22 - x12) * 10)
        ycont2 = y12 - ((y22 - y12) * 10)
        x2cont2 = x22 + ((x22 - x12) * 10)
        y2cont2 = y22 + ((y22 - y12) * 10)
        pic1.Line (xcont2, ycont2)-(x2cont2, y2cont2)
    End If
    Call cmdCalculate_Click
End If
End Sub

Private Sub pic1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblPosition = Format(X, "fixed") + "," + Format(Y, "fixed")
End Sub

Private Sub txtX1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtY1.SetFocus
End If
End Sub

Private Sub txtX2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtY2.SetFocus
End If
End Sub

Private Sub txtX3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtY3.SetFocus
End If
End Sub

Private Sub txtX4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtY4.SetFocus
End If
End Sub

Private Sub txtY1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtX2.SetFocus
End If
End Sub

Private Sub txtY2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtX3.SetFocus
End If
End Sub

Private Sub txtY3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtX4.SetFocus
End If
End Sub

Private Sub txtY4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdCalculate.SetFocus
End If
End Sub
