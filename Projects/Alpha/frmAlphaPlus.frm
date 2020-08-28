VERSION 5.00
Begin VB.Form frmAlphaPlus 
   Caption         =   "Form1"
   ClientHeight    =   5055
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4065
   LinkTopic       =   "Form1"
   ScaleHeight     =   5055
   ScaleWidth      =   4065
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   500
      Left            =   2160
      TabIndex        =   11
      Top             =   4260
      Width           =   1500
   End
   Begin VB.CommandButton Command 
      Caption         =   "Alpha"
      Height          =   500
      Left            =   360
      TabIndex        =   8
      Top             =   4320
      Width           =   1500
   End
   Begin VB.TextBox txt5 
      Height          =   500
      Left            =   360
      TabIndex        =   4
      Top             =   3540
      Width           =   1500
   End
   Begin VB.TextBox txt4 
      Height          =   500
      Left            =   360
      TabIndex        =   3
      Top             =   2760
      Width           =   1500
   End
   Begin VB.TextBox txt3 
      Height          =   500
      Left            =   360
      TabIndex        =   2
      Top             =   1980
      Width           =   1500
   End
   Begin VB.TextBox txt2 
      Height          =   500
      Left            =   360
      TabIndex        =   1
      Top             =   1140
      Width           =   1500
   End
   Begin VB.TextBox txt1 
      Height          =   500
      Left            =   360
      TabIndex        =   0
      Top             =   300
      Width           =   1500
   End
   Begin VB.Label lbl2 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   2160
      TabIndex        =   10
      Top             =   1140
      Width           =   1500
   End
   Begin VB.Label lbl1 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   2160
      TabIndex        =   9
      Top             =   300
      Width           =   1500
   End
   Begin VB.Label lbl5 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   2160
      TabIndex        =   7
      Top             =   3540
      Width           =   1500
   End
   Begin VB.Label lbl4 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   2160
      TabIndex        =   6
      Top             =   2760
      Width           =   1500
   End
   Begin VB.Label lbl3 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   2160
      TabIndex        =   5
      Top             =   1980
      Width           =   1500
   End
End
Attribute VB_Name = "frmAlphaPlus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Word1, Word2 As String
Dim Word1Number, Word2Number, Word3Number, Word4Number, Word5Number As Integer
Dim X, Y As Integer

Private Sub cmdClear_Click()

'clears everything
Word1Number = 0
Word2Number = 0
Word3Number = 0
Word4Number = 0
Word5Number = 0
Word1 = ""
Word2 = ""
txt1.Text = ""
txt2.Text = ""
txt3.Text = ""
txt4.Text = ""
txt5.Text = ""
lbl1.Caption = ""
lbl2.Caption = ""
lbl3.Caption = ""
lbl4.Caption = ""
lbl5.Caption = ""

End Sub

Private Sub Command_Click()

'resets everything
Word1Number = 0
Word2Number = 0
Word3Number = 0
Word4Number = 0
Word5Number = 0
Word1 = ""
Word2 = ""
For X = 1 To 5 'goes from 1 to 5
    'changes word1's string based on x value
    If X = 1 Then
        Word1 = txt1.Text
    ElseIf X = 2 Then
        Word1 = txt2.Text
    ElseIf X = 3 Then
        Word1 = txt3.Text
    ElseIf X = 4 Then
        Word1 = txt4.Text
    ElseIf X = 5 Then
        Word1 = txt5.Text
    End If
    For Y = 1 To 5 'goes from 1 to 5
        'changes word2's string based on y value
        If Y = 1 Then
            Word2 = txt1.Text
        ElseIf Y = 2 Then
            Word2 = txt2.Text
        ElseIf Y = 3 Then
            Word2 = txt3.Text
        ElseIf Y = 4 Then
            Word2 = txt4.Text
        ElseIf Y = 5 Then
            Word2 = txt5.Text
        End If
        If StrComp(Word1, Word2, 1) = -1 Then 'checks whether word1's string is alphabetically higher than word2's string
            'adds to a counter variable based on what x is
            If X = 1 Then
                Word1Number = Word1Number + 1
            ElseIf X = 2 Then
                Word2Number = Word2Number + 1
            ElseIf X = 3 Then
                Word3Number = Word3Number + 1
            ElseIf X = 4 Then
                Word4Number = Word4Number + 1
            ElseIf X = 5 Then
                Word5Number = Word5Number + 1
            End If
        End If
    Next Y
Next X
'checks the amount of times a string was larger than another, and asigns its alphabetical ranking based off of that, then outputs the strings to the captions
If Word1Number = 4 Then
    lbl1.Caption = txt1.Text
ElseIf Word1Number = 3 Then
    lbl2.Caption = txt1.Text
ElseIf Word1Number = 2 Then
    lbl3.Caption = txt1.Text
ElseIf Word1Number = 1 Then
    lbl4.Caption = txt1.Text
Else
    lbl5.Caption = txt1.Text
End If
If Word2Number = 4 Then
    lbl1.Caption = txt2.Text
ElseIf Word2Number = 3 Then
    lbl2.Caption = txt2.Text
ElseIf Word2Number = 2 Then
    lbl3.Caption = txt2.Text
ElseIf Word2Number = 1 Then
    lbl4.Caption = txt2.Text
Else
    lbl5.Caption = txt2.Text
End If
If Word3Number = 4 Then
    lbl1.Caption = txt3.Text
ElseIf Word3Number = 3 Then
    lbl2.Caption = txt3.Text
ElseIf Word3Number = 2 Then
    lbl3.Caption = txt3.Text
ElseIf Word3Number = 1 Then
    lbl4.Caption = txt3.Text
Else
    lbl5.Caption = txt3.Text
End If
If Word4Number = 4 Then
    lbl1.Caption = txt4.Text
ElseIf Word4Number = 3 Then
    lbl2.Caption = txt4.Text
ElseIf Word4Number = 2 Then
    lbl3.Caption = txt4.Text
ElseIf Word4Number = 1 Then
    lbl4.Caption = txt4.Text
Else
    lbl5.Caption = txt4.Text
End If
If Word5Number = 4 Then
    lbl1.Caption = txt5.Text
ElseIf Word5Number = 3 Then
    lbl2.Caption = txt5.Text
ElseIf Word5Number = 2 Then
    lbl3.Caption = txt5.Text
ElseIf Word5Number = 1 Then
    lbl4.Caption = txt5.Text
Else
    lbl5.Caption = txt5.Text
End If

End Sub
