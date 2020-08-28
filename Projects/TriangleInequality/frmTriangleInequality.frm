VERSION 5.00
Begin VB.Form frmTriangleInequality 
   BackColor       =   &H00000080&
   Caption         =   "Form1"
   ClientHeight    =   8985
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14610
   LinkTopic       =   "Form1"
   ScaleHeight     =   8985
   ScaleWidth      =   14610
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pic1 
      AutoRedraw      =   -1  'True
      Height          =   4695
      Left            =   8340
      ScaleHeight     =   4635
      ScaleWidth      =   4995
      TabIndex        =   34
      Top             =   2940
      Width           =   5055
      Begin VB.Label lblC 
         AutoSize        =   -1  'True
         Caption         =   "C"
         Height          =   195
         Left            =   420
         TabIndex        =   37
         Top             =   60
         Width           =   105
      End
      Begin VB.Label lblB 
         AutoSize        =   -1  'True
         Caption         =   "B"
         Height          =   195
         Left            =   240
         TabIndex        =   36
         Top             =   60
         Width           =   105
      End
      Begin VB.Label lblA 
         AutoSize        =   -1  'True
         Caption         =   "A"
         Height          =   195
         Left            =   60
         TabIndex        =   35
         Top             =   60
         Width           =   105
      End
   End
   Begin VB.Timer Timer 
      Interval        =   100
      Left            =   3780
      Top             =   2100
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H000000C0&
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   3660
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   8280
      Width           =   800
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H000000C0&
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   8280
      Width           =   1600
   End
   Begin VB.CommandButton cmdCalculate 
      Appearance      =   0  'Flat
      BackColor       =   &H000000C0&
      Caption         =   "Calc"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   180
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   8280
      UseMaskColor    =   -1  'True
      Width           =   1600
   End
   Begin VB.TextBox txtC 
      BackColor       =   &H0080C0FF&
      Height          =   800
      Left            =   4320
      TabIndex        =   2
      Top             =   780
      Width           =   1800
   End
   Begin VB.TextBox txtB 
      BackColor       =   &H0080C0FF&
      Height          =   800
      Left            =   2280
      TabIndex        =   1
      Top             =   780
      Width           =   1800
   End
   Begin VB.TextBox txtA 
      BackColor       =   &H0080C0FF&
      Height          =   800
      Left            =   180
      TabIndex        =   0
      Top             =   780
      Width           =   1800
   End
   Begin VB.Label lblAngleExtC 
      BackColor       =   &H0080C0FF&
      BorderStyle     =   1  'Fixed Single
      Height          =   600
      Left            =   6360
      TabIndex        =   33
      Top             =   4680
      Width           =   1800
   End
   Begin VB.Label lblAngleExtB 
      BackColor       =   &H0080C0FF&
      BorderStyle     =   1  'Fixed Single
      Height          =   600
      Left            =   6360
      TabIndex        =   32
      Top             =   3540
      Width           =   1800
   End
   Begin VB.Label lblAngleExtA 
      BackColor       =   &H0080C0FF&
      BorderStyle     =   1  'Fixed Single
      Height          =   600
      Left            =   6360
      TabIndex        =   31
      Top             =   2400
      Width           =   1800
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Angle Ext.C"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   405
      Index           =   11
      Left            =   6360
      TabIndex        =   30
      Top             =   4200
      Width           =   1800
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Angle Ext.B"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   405
      Index           =   10
      Left            =   6360
      TabIndex        =   29
      Top             =   3060
      Width           =   1800
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Angle Ext.A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   405
      Index           =   9
      Left            =   6360
      TabIndex        =   28
      Top             =   1860
      Width           =   1800
   End
   Begin VB.Image img5 
      Height          =   2115
      Left            =   5640
      Picture         =   "frmTriangleInequality.frx":0000
      Stretch         =   -1  'True
      Top             =   5520
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Image img4 
      Height          =   2115
      Left            =   5640
      Picture         =   "frmTriangleInequality.frx":073D
      Stretch         =   -1  'True
      Top             =   5520
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Image img3 
      Height          =   2115
      Left            =   5640
      Picture         =   "frmTriangleInequality.frx":0F26
      Stretch         =   -1  'True
      Top             =   5520
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Image img2 
      Height          =   2115
      Left            =   5640
      Picture         =   "frmTriangleInequality.frx":1663
      Stretch         =   -1  'True
      Top             =   5520
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Image img1 
      Height          =   2115
      Left            =   5640
      Picture         =   "frmTriangleInequality.frx":1C95
      Stretch         =   -1  'True
      Top             =   5520
      Width           =   1575
   End
   Begin VB.Label Label 
      BackColor       =   &H00000080&
      Caption         =   "*c is longest side (Unless triangle is Equilateral)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   315
      Index           =   8
      Left            =   180
      TabIndex        =   27
      Top             =   7860
      Width           =   4275
   End
   Begin VB.Image Image 
      Height          =   2520
      Left            =   8580
      Picture         =   "frmTriangleInequality.frx":23D2
      Stretch         =   -1  'True
      Top             =   120
      Width           =   2745
   End
   Begin VB.Label lblAngleC 
      BackColor       =   &H0080C0FF&
      BorderStyle     =   1  'Fixed Single
      Height          =   600
      Left            =   4320
      TabIndex        =   26
      Top             =   4680
      Width           =   1800
   End
   Begin VB.Label lblAngleB 
      BackColor       =   &H0080C0FF&
      BorderStyle     =   1  'Fixed Single
      Height          =   600
      Left            =   4320
      TabIndex        =   25
      Top             =   3540
      Width           =   1800
   End
   Begin VB.Label lblAngleA 
      BackColor       =   &H0080C0FF&
      BorderStyle     =   1  'Fixed Single
      Height          =   600
      Left            =   4320
      TabIndex        =   24
      Top             =   2400
      Width           =   1800
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Angle C"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   405
      Index           =   7
      Left            =   4320
      TabIndex        =   23
      Top             =   4200
      Width           =   1800
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Angle B"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   405
      Index           =   6
      Left            =   4320
      TabIndex        =   22
      Top             =   3060
      Width           =   1800
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Angle A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   405
      Index           =   5
      Left            =   4320
      TabIndex        =   21
      Top             =   1860
      Width           =   1800
   End
   Begin VB.Image imgEquilateral 
      Height          =   2160
      Left            =   2880
      Picture         =   "frmTriangleInequality.frx":7C33
      Stretch         =   -1  'True
      Top             =   5520
      Visible         =   0   'False
      Width           =   2580
   End
   Begin VB.Image imgIsosceles 
      Height          =   2160
      Left            =   2820
      Picture         =   "frmTriangleInequality.frx":CD7F
      Stretch         =   -1  'True
      Top             =   5520
      Visible         =   0   'False
      Width           =   2580
   End
   Begin VB.Image imgScalene 
      Height          =   2160
      Left            =   2820
      Picture         =   "frmTriangleInequality.frx":10E30
      Stretch         =   -1  'True
      Top             =   5520
      Visible         =   0   'False
      Width           =   2580
   End
   Begin VB.Image imgRight 
      Height          =   2160
      Left            =   180
      Picture         =   "frmTriangleInequality.frx":1AB6B
      Stretch         =   -1  'True
      Top             =   5520
      Visible         =   0   'False
      Width           =   2580
   End
   Begin VB.Image imgAcute 
      Height          =   2160
      Left            =   180
      Picture         =   "frmTriangleInequality.frx":1E925
      Stretch         =   -1  'True
      Top             =   5520
      Visible         =   0   'False
      Width           =   2580
   End
   Begin VB.Image imgObtuse 
      Height          =   2160
      Left            =   180
      Picture         =   "frmTriangleInequality.frx":24FDF
      Stretch         =   -1  'True
      Top             =   5520
      Visible         =   0   'False
      Width           =   2580
   End
   Begin VB.Label lblScalene 
      Alignment       =   2  'Center
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Scalene"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   405
      Left            =   2280
      TabIndex        =   20
      Top             =   4680
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.Label lblEquilateral 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Equilateral"
      Height          =   405
      Left            =   2280
      TabIndex        =   19
      Top             =   4680
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.Label lblIsosceles 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Isosceles"
      Height          =   405
      Left            =   2280
      TabIndex        =   18
      Top             =   4680
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.Label lblObtuse 
      Alignment       =   2  'Center
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Obtuse"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   405
      Left            =   180
      TabIndex        =   17
      Top             =   4680
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.Label lblAcute 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Acute"
      Height          =   405
      Left            =   180
      TabIndex        =   16
      Top             =   4680
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.Label lblRight 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Right "
      Height          =   405
      Left            =   180
      TabIndex        =   15
      Top             =   4680
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.Label lblArea 
      BackColor       =   &H0080C0FF&
      BorderStyle     =   1  'Fixed Single
      Height          =   795
      Left            =   2280
      TabIndex        =   14
      Top             =   3540
      Width           =   1800
   End
   Begin VB.Label lblPerimeter 
      BackColor       =   &H0080C0FF&
      BorderStyle     =   1  'Fixed Single
      Height          =   795
      Left            =   180
      TabIndex        =   13
      Top             =   3540
      Width           =   1800
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Area"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   405
      Index           =   4
      Left            =   2280
      TabIndex        =   12
      Top             =   2760
      Width           =   1800
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Perimeter"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   405
      Index           =   3
      Left            =   180
      TabIndex        =   11
      Top             =   2820
      Width           =   1800
   End
   Begin VB.Label lblIsNot 
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "It is not a Triangle!"
      Height          =   495
      Left            =   180
      TabIndex        =   10
      Top             =   2040
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Label lblIs 
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "It is a Triangle!"
      Height          =   495
      Left            =   180
      TabIndex        =   9
      Top             =   2040
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Side c"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   405
      Index           =   2
      Left            =   4320
      TabIndex        =   6
      Top             =   240
      Width           =   1800
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Side b"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   405
      Index           =   1
      Left            =   2280
      TabIndex        =   5
      Top             =   240
      Width           =   1800
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Side a"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   405
      Index           =   0
      Left            =   180
      TabIndex        =   4
      Top             =   240
      Width           =   1800
   End
End
Attribute VB_Name = "frmTriangleInequality"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim a, b, c, s, x, y, z As Double
Dim PointAx, PointAy, PointBx, PointBy, PointCx, PointCy As Double
Dim HeightOfSideB, HeightOfSideBGraph, WidthOfSideBGraph, GraphSideC, Area As Double
Dim i As Integer

Private Sub cmdCalculate_Click()

a = Val(txtA)
b = Val(txtB)
c = Val(txtC)

'Determine longest side

If a > c Then
    c = Val(txtA)
    a = Val(txtC)
ElseIf b > c Then
    c = Val(txtB)
    b = Val(txtC)
ElseIf a > c And b > c Then
    If a > b Then
        c = Val(txtA)
        a = Val(txtC)
    ElseIf b > a Then
        c = Val(txtB)
        b = Val(txtA)
    End If
End If

'Test if sides form triangle

If a + b > c And b + c > a And a + c > b Then
    lblIs.Visible = True
    lblIsNot.Visible = False
    
    'Perimeter
  
    lblPerimeter = a + b + c
   
    'Area
    
    s = (a + b + c) / 2
    Area = Math.Sqr(s * (s - a) * (s - b) * (s - c))
    lblArea = Area
    'Test what type of Triangle
   
    If (a ^ 2) + (b ^ 2) = (c ^ 2) Then
            lblRight.Visible = True
    ElseIf (a ^ 2) + (b ^ 2) > (c ^ 2) Then
            lblAcute.Visible = True
    ElseIf (a ^ 2) + (b ^ 2) < (c ^ 2) Then
            lblObtuse.Visible = True
    End If

    
    If a = b And b = c And a = c Then
            lblEquilateral.Visible = True
    ElseIf a = b Or b = c Or a = c Then
            lblIsosceles.Visible = True
    ElseIf a <> b And b <> c And a <> c Then
            lblScalene.Visible = True
    End If
   
    'Sample triangle images
    
    If lblRight.Visible = True Then
        imgRight.Visible = True
    ElseIf lblAcute.Visible = True Then
        imgAcute.Visible = True
    ElseIf lblObtuse.Visible = True Then
        imgObtuse.Visible = True
    End If
    
    If lblEquilateral.Visible = True Then
        imgEquilateral.Visible = True
    ElseIf lblIsosceles.Visible = True Then
        imgIsosceles.Visible = True
    ElseIf lblScalene.Visible = True Then
        imgScalene.Visible = True
    End If
    
    'Angles
    
    x = ((a ^ 2) + (b ^ 2) - (c ^ 2)) / (2 * a * b)
    lblAngleC = (180 / 3.14159265359) * (Atn(-x / Sqr(-x * x + 1)) + 2 * Atn(1))
    y = ((b ^ 2) + (c ^ 2) - (a ^ 2)) / (2 * b * c)
    lblAngleA = (180 / 3.14159265359) * (Atn(-y / Sqr(-y * y + 1)) + 2 * Atn(1))
    z = ((c ^ 2) + (a ^ 2) - (b ^ 2)) / (2 * c * a)
    lblAngleB = (180 / 3.14159265359) * (Atn(-z / Sqr(-z * z + 1)) + 2 * Atn(1))
    
    lblAngleExtA = 180 - lblAngleA
    lblAngleExtB = 180 - lblAngleB
    lblAngleExtC = 180 - lblAngleC
    

Else
    
    lblIsNot.Visible = True
    lblIs.Visible = False
    lblPerimeter = ""
    lblArea = ""
    lblAngleA = ""
    lblAngleB = ""
    lblAngleC = ""
    lblAngleExtA = ""
    lblAngleExtB = ""
    lblAngleExtC = ""
    lblRight.Visible = False 'make all labels invisible
    lblAcute.Visible = False
    lblObtuse.Visible = False
    lblEquilateral.Visible = False
    lblIsosceles.Visible = False
    lblScalene.Visible = False
    imgRight.Visible = False    'make all images invisible
    imgAcute.Visible = False
    imgObtuse.Visible = False
    imgEquilateral.Visible = False
    imgIsosceles.Visible = False
    imgScalene.Visible = False
    
End If

cmdClear.SetFocus

End Sub

Private Sub cmdClear_Click()

txtA = ""              'clear the textboxes
txtB = ""
txtC = ""
lblPerimeter = ""
lblArea = ""
lblAngleA = ""
lblAngleB = ""
lblAngleC = ""
lblAngleExtA = ""
lblAngleExtB = ""
lblAngleExtC = ""
lblIs.Visible = False   'make all labels invisible
lblIsNot.Visible = False
lblRight.Visible = False
lblAcute.Visible = False
lblObtuse.Visible = False
lblEquilateral.Visible = False
lblIsosceles.Visible = False
lblScalene.Visible = False
imgRight.Visible = False    'make all images invisible
imgAcute.Visible = False
imgObtuse.Visible = False
imgEquilateral.Visible = False
imgIsosceles.Visible = False
imgScalene.Visible = False
txtA.SetFocus           'shift focus to txtA
pic1.Cls
Call Form_Activate

End Sub

Private Sub cmdExit_Click()

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
Private Sub pic1_Click()
Dim cx, cy As Double

pic1.Cls
Call Form_Activate
pic1.Line (0, 0)-(b, 0)
s = (a + b + c) / 2
cx = ((a ^ 2) - (b ^ 2) - (c ^ 2)) / -(2 * b)
cy = (2 * Sqr(s * (s - a) * (s - b) * (s - c))) / b
pic1.Line (0, 0)-(cx, cy)
pic1.Line (cx, cy)-(b, 0)
lblA.Top = -0.25
lblA.Left = -0.25
lblB.Top = -0.25
lblB.Left = b + 0.25
lblC.Top = cy + 1
lblC.Left = cx - 0.25

End Sub

Private Sub Timer_Timer()

If img1.Visible = True Then
    img1.Visible = False
    img2.Visible = True
    img3.Visible = False
    img4.Visible = False
    img5.Visible = False
ElseIf img2.Visible = True Then
    img1.Visible = False
    img2.Visible = False
    img3.Visible = True
    img4.Visible = False
    img5.Visible = False
ElseIf img3.Visible = True Then
    img1.Visible = False
    img2.Visible = False
    img3.Visible = False
    img4.Visible = True
    img5.Visible = False
ElseIf img4.Visible = True Then
    img1.Visible = False
    img2.Visible = False
    img3.Visible = False
    img4.Visible = False
    img5.Visible = True
ElseIf img5.Visible = True Then
    img1.Visible = True
    img2.Visible = False
    img3.Visible = False
    img4.Visible = False
    img5.Visible = False
End If

End Sub

Private Sub txtA_Change()

Call cmdCalculate_Click
txtA.SetFocus

End Sub

Private Sub txtA_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    txtB.SetFocus
End If

End Sub

Private Sub txtB_Change()

Call cmdCalculate_Click
txtB.SetFocus

End Sub

Private Sub txtB_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    txtC.SetFocus
End If

End Sub

Private Sub txtC_Change()

Call cmdCalculate_Click
txtC.SetFocus

End Sub

Private Sub txtC_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    cmdCalculate.SetFocus
End If

End Sub
