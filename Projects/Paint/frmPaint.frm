VERSION 5.00
Begin VB.Form frmPaint 
   AutoRedraw      =   -1  'True
   BackColor       =   &H8000000E&
   Caption         =   "Paint"
   ClientHeight    =   6990
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   10905
   LinkTopic       =   "Form1"
   ScaleHeight     =   6990
   ScaleWidth      =   10905
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdFiletest 
      Height          =   615
      Left            =   11040
      TabIndex        =   22
      Top             =   1140
      Width           =   675
   End
   Begin VB.ListBox lstArrayData 
      Height          =   1815
      ItemData        =   "frmPaint.frx":0000
      Left            =   7800
      List            =   "frmPaint.frx":0002
      TabIndex        =   20
      Top             =   60
      Width           =   3075
   End
   Begin VB.Frame Frame 
      Height          =   1875
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7755
      Begin VB.HScrollBar hsbB 
         Height          =   255
         Left            =   1155
         Max             =   255
         TabIndex        =   4
         Top             =   1440
         Width           =   4215
      End
      Begin VB.HScrollBar hsbG 
         Height          =   255
         Left            =   1155
         Max             =   255
         TabIndex        =   3
         Top             =   1080
         Width           =   4215
      End
      Begin VB.HScrollBar hsbR 
         Height          =   255
         Left            =   1155
         Max             =   255
         TabIndex        =   2
         Top             =   720
         Width           =   4215
      End
      Begin VB.HScrollBar hsbWidth 
         Height          =   255
         Left            =   1140
         Max             =   100
         Min             =   1
         TabIndex        =   1
         Top             =   240
         Value           =   5
         Width           =   4215
      End
      Begin VB.Image imgRainbow 
         Height          =   495
         Left            =   5460
         Picture         =   "frmPaint.frx":0004
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   555
      End
      Begin VB.Label lblBlack 
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   7260
         TabIndex        =   19
         Top             =   600
         Width           =   315
      End
      Begin VB.Label lblGrey 
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   6825
         TabIndex        =   18
         Top             =   600
         Width           =   315
      End
      Begin VB.Label lblWhite 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   6375
         TabIndex        =   17
         Top             =   600
         Width           =   315
      End
      Begin VB.Label lblBrown 
         BackColor       =   &H00004080&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   5925
         TabIndex        =   16
         Top             =   600
         Width           =   315
      End
      Begin VB.Label lblPurple 
         BackColor       =   &H00C000C0&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   5460
         TabIndex        =   15
         Top             =   600
         Width           =   315
      End
      Begin VB.Label lblBlue 
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   7260
         TabIndex        =   14
         Top             =   240
         Width           =   315
      End
      Begin VB.Label lblGreen 
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   6825
         TabIndex        =   13
         Top             =   240
         Width           =   315
      End
      Begin VB.Label lblYellow 
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   6375
         TabIndex        =   12
         Top             =   240
         Width           =   315
      End
      Begin VB.Label lblOrange 
         BackColor       =   &H000080FF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   5925
         TabIndex        =   11
         Top             =   240
         Width           =   315
      End
      Begin VB.Label lblRed 
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   5460
         TabIndex        =   10
         Top             =   240
         Width           =   315
      End
      Begin VB.Label lblColor 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   675
         Left            =   6240
         TabIndex        =   9
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Blue"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   8
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Green"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   7
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Red"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Pen Width"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Timer Timer 
      Interval        =   10
      Left            =   -20
      Top             =   -20
   End
   Begin VB.Label lblopentest 
      Height          =   255
      Left            =   11760
      TabIndex        =   23
      Top             =   1260
      Width           =   495
   End
   Begin VB.Label lblFileName 
      BorderStyle     =   1  'Fixed Single
      Height          =   795
      Left            =   11100
      TabIndex        =   21
      Top             =   180
      Width           =   1815
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuOpen 
         Caption         =   "Open"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save"
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Save As"
      End
      Begin VB.Menu mnuNew 
         Caption         =   "New"
      End
      Begin VB.Menu mnuSeperator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmPaint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim R, G, B As Integer
Dim Drawing As Integer
Const MaxPoints = 3000
Dim SingleX(0 To MaxPoints) As Single
Dim SingleY(0 To MaxPoints) As Single
Dim arrayR(5000) As Double
Dim arrayG(5000) As Double
Dim arrayB(5000) As Double
Dim arrayW(5000) As Double
Dim NumPoints As Integer
Dim i As Long
Dim w As Long
Dim strfn, path, ans As String
Dim Rainbow, FTest, OpenTest As Boolean
Dim v, z As Integer
Dim FName As String

Sub DrawCircle(X As Single, Y As Single)
Circle (X, Y), 1, RGB(R, G, B)
If NumPoints < MaxPoints Then
    NumPoints = NumPoints + 1
    SingleX(NumPoints) = X
    SingleY(NumPoints) = Y
    arrayR(NumPoints) = R
    arrayG(NumPoints) = G
    arrayB(NumPoints) = B
    arrayW(NumPoints) = w
End If
End Sub
Sub Drawlines()
CurrentX = SingleX(1)
CurrentY = SingleY(1)
Circle (SingleX(1), SingleY(1)), 1
For i = 1 To NumPoints
    If SingleX(i) <> -20 And SingleY(i) <> -20 Then
        frmPaint.DrawWidth = (arrayW(i)) + 1
        Line -(SingleX(i), SingleY(i)), RGB(arrayR(i), arrayG(i), arrayB(i))
        Circle (SingleX(i), SingleY(i)), 1, RGB(arrayR(i), arrayG(i), arrayB(i))
    Else
        frmPaint.CurrentX = SingleX(i + 1)
        frmPaint.CurrentY = SingleY(i + 1)
    End If
    lstArrayData.AddItem (SingleX(i) & " " & SingleY(i) & " " & arrayR(i) & " " & arrayG(i) & " " & arrayB(i) & " " & arrayW(i))
Next i
End Sub

Private Sub Command_Click()
Rainbow = True
End Sub

Private Sub Form_Load()
NumPoints = 0
w = 1
R = 0
G = 0
B = 0
frmPaint.AutoRedraw = True
Rainbow = False
z = 0
v = 1
FTest = False
FName = ""
OpenTest = False
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Drawing = 1
w = hsbWidth.Value
DrawCircle X, Y
lstArrayData.AddItem (Str(SingleX(NumPoints)) & " " & Str(SingleY(NumPoints)) & " " & R & " " & G & " " & B & " " & w)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
w = hsbWidth.Value
If Drawing = 1 Then
    Line -(X, Y), RGB(R, G, B)
    DrawCircle X, Y
    lstArrayData.AddItem (Str(SingleX(NumPoints)) & " " & Str(SingleY(NumPoints)) & " " & R & " " & G & " " & B & " " & w)
End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Drawing = 0

If NumPoints < 5000 Then
    NumPoints = NumPoints + 1
    SingleX(NumPoints) = -20
    SingleY(NumPoints) = -20
    arrayR(NumPoints) = hsbR.Value
    arrayG(NumPoints) = hsbG.Value
    arrayB(NumPoints) = hsbB.Value
    arrayW(NumPoints) = hsbWidth.Value
    lstArrayData.AddItem (Str(SingleX(NumPoints)) & " " & Str(SingleY(NumPoints)) & " " & R & " " & G & " " & B & " " & w)
End If
End Sub

Private Sub hsbWidth_Change()
w = hsbWidth.Value
frmPaint.DrawWidth = w
End Sub

Private Sub imgRainbow_Click()
If v = 1 Then
    Rainbow = True
    v = 2
ElseIf v = 2 Then
    Rainbow = False
    v = 1
End If
End Sub

Private Sub mnuExit_Click()
End
End Sub

Private Sub mnuNew_Click()
Rainbow = False
NumPoints = 0
frmPaint.Cls
R = 0
G = 0
B = 0
w = 1
frmPaint.DrawWidth = w
lblColor.BackColor = RGB(R, G, B)
lstArrayData.Clear
FName = ""
FTest = False
lblFileName = ""
frmFile.txtFName = ""
frmFile.lblPath = ""
frmFile.lblPName = ""
End Sub

Private Sub mnuSave_Click()
Open path For Binary Access Write As #1
Put #1, , NumPoints
Dim i As Integer
For i = 0 To NumPoints
    Put #1, , SingleX(i)
    Put #1, , SingleY(i)
    Put #1, , arrayR(i)
    Put #1, , arrayG(i)
    Put #1, , arrayB(i)
    Put #1, , arrayW(i)
Next i
Close #1
End Sub

Private Sub mnuOpen_Click()
ans = vbNo
Do While ans = vbNo
    Dim strfn As String
    strfn = UCase(InputBox("Filename", "filename", "Bob"))
'    path = "F:\Vignesh Files\" + strfn + ".Dat"
    path = "C:\Users\Neel Shettigar\Desktop\19 S2 NS CP1\Paint\Paintings" + strfn + ".Dat"
    ans = MsgBox(path, vbYesNo, "Is this the path?")
Loop
If ans = vbYes Then
    Open path For Binary Access Read As #1
    Get #1, , NumPoints
    Dim i As Integer
    For i = 0 To NumPoints
        Get #1, , SingleX(i)
        Get #1, , SingleY(i)
        Get #1, , arrayR(i)
        Get #1, , arrayG(i)
        Get #1, , arrayB(i)
        Get #1, , arrayW(i)
    Next i
    Close #1
End If
frmPaint.Cls
Drawlines
End Sub

Private Sub mnuSaveAs_Click()
ans = vbNo
Do While ans = vbNo
    Dim strfn As String
    strfn = UCase(InputBox("Filename", "filename", "Bob"))
'    path = "F:\Vignesh Files\" + strfn + ".Dat"
    path = "C:\Users\Neel Shettigar\Desktop\19 S2 NS CP1\Paint\Paintings" + strfn + ".Dat"
    ans = MsgBox(path, vbYesNo, "Is this the path?")
Loop
If ans = vbYes Then
    Open path For Binary Access Write As #1
    Put #1, , NumPoints
    Dim i As Integer
    For i = 0 To NumPoints
        Put #1, , SingleX(i)
        Put #1, , SingleY(i)
        Put #1, , arrayR(i)
        Put #1, , arrayG(i)
        Put #1, , arrayB(i)
        Put #1, , arrayW(i)
    Next i
    Close #1
End If
End Sub
Private Sub hsbR_Change()
R = hsbR.Value
lblColor.BackColor = RGB(R, G, B)
End Sub

Private Sub hsbG_Change()
G = hsbG.Value
lblColor.BackColor = RGB(R, G, B)
End Sub

Private Sub hsbB_Change()
B = hsbB.Value
lblColor.BackColor = RGB(R, G, B)
End Sub

Private Sub lblBlack_Click()
hsbR.Value = 0
hsbG.Value = 0
hsbB.Value = 0
lblColor.BackColor = RGB(R, G, B)
End Sub

Private Sub lblBlue_Click()
hsbR.Value = 0
hsbG.Value = 0
hsbB.Value = 255
lblColor.BackColor = RGB(R, G, B)
End Sub

Private Sub lblBrown_Click()
hsbR.Value = 154
hsbG.Value = 76
hsbB.Value = 0
lblColor.BackColor = RGB(R, G, B)
End Sub

Private Sub lblGreen_Click()
hsbR.Value = 0
hsbG.Value = 255
hsbB.Value = 0
lblColor.BackColor = RGB(R, G, B)
End Sub

Private Sub lblGrey_Click()
hsbR.Value = 128
hsbG.Value = 128
hsbB.Value = 128
lblColor.BackColor = RGB(R, G, B)
End Sub

Private Sub lblOrange_Click()
hsbR.Value = 255
hsbG.Value = 128
hsbB.Value = 0
lblColor.BackColor = RGB(R, G, B)
End Sub

Private Sub lblPurple_Click()
hsbR.Value = 204
hsbG.Value = 0
hsbB.Value = 204
lblColor.BackColor = RGB(R, G, B)
End Sub

Private Sub lblRed_Click()
hsbR.Value = 255
hsbG.Value = 0
hsbB.Value = 0
lblColor.BackColor = RGB(R, G, B)
End Sub

Private Sub lblWhite_Click()
hsbR.Value = 255
hsbG.Value = 255
hsbB.Value = 255
lblColor.BackColor = RGB(R, G, B)
End Sub

Private Sub lblYellow_Click()
hsbR.Value = 255
hsbG.Value = 255
hsbB.Value = 0
lblColor.BackColor = RGB(R, G, B)
End Sub

Private Sub Timer_Timer()
If Rainbow = True Then
    z = z + 1
    If z = 7 Then
        z = 1
    End If
    If z = 1 Then
        lblRed_Click
    ElseIf z = 2 Then
        lblOrange_Click
    ElseIf z = 3 Then
        lblYellow_Click
    ElseIf z = 4 Then
        lblGreen_Click
    ElseIf z = 5 Then
        lblBlue_Click
    ElseIf z = 6 Then
        lblPurple_Click
    End If
End If
End Sub
