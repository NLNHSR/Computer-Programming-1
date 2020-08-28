VERSION 5.00
Begin VB.Form frmAlpha 
   Caption         =   "Form1"
   ClientHeight    =   4995
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4140
   LinkTopic       =   "Form1"
   ScaleHeight     =   4995
   ScaleWidth      =   4140
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   500
      Left            =   2040
      TabIndex        =   5
      Top             =   4020
      Width           =   1500
   End
   Begin VB.TextBox txt4 
      Alignment       =   2  'Center
      Height          =   500
      Left            =   240
      TabIndex        =   3
      Top             =   3120
      Width           =   1500
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Alpha"
      Height          =   500
      Left            =   240
      TabIndex        =   4
      Top             =   4020
      Width           =   1500
   End
   Begin VB.TextBox txt3 
      Alignment       =   2  'Center
      Height          =   500
      Left            =   240
      TabIndex        =   2
      Top             =   2280
      Width           =   1500
   End
   Begin VB.TextBox txt2 
      Alignment       =   2  'Center
      Height          =   500
      Left            =   240
      TabIndex        =   1
      Top             =   1380
      Width           =   1500
   End
   Begin VB.TextBox txt1 
      Alignment       =   2  'Center
      Height          =   500
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   1500
   End
   Begin VB.Label lbl4 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   2040
      TabIndex        =   9
      Top             =   3120
      Width           =   1500
   End
   Begin VB.Label lbl3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   2040
      TabIndex        =   8
      Top             =   2280
      Width           =   1500
   End
   Begin VB.Label lbl2 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   2040
      TabIndex        =   7
      Top             =   1380
      Width           =   1500
   End
   Begin VB.Label lbl1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   2040
      TabIndex        =   6
      Top             =   480
      Width           =   1500
   End
End
Attribute VB_Name = "frmAlpha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Word1, Word2, Word3, Word4 As String

Private Sub cmd_Click()

Word1 = txt1.Text
Word2 = txt2.Text
Word3 = txt3.Text
Word4 = txt4.Text

If LCase(Word1) < LCase(Word2) And LCase(Word1) < LCase(Word3) And LCase(Word1) < LCase(Word4) Then
    
    lbl1.Caption = Word1
    If LCase(Word2) < LCase(Word3) And LCase(Word2) < LCase(Word4) Then
        
        lbl2.Caption = Word2
            If LCase(Word3) < LCase(Word4) Then
                lbl3.Caption = Word3
                lbl4.Caption = Word4
            ElseIf LCase(Word4) < LCase(Word3) Then
                lbl3.Caption = Word4
                lbl4.Caption = Word3
            End If
            
    ElseIf LCase(Word3) < LCase(Word2) And LCase(Word3) < LCase(Word4) Then
        
        lbl2.Caption = Word3
            If LCase(Word2) < LCase(Word4) Then
                lbl3.Caption = Word2
                lbl4.Caption = Word4
            ElseIf LCase(Word4) < LCase(Word2) Then
                lbl3.Caption = Word4
                lbl4.Caption = Word2
            End If
            
    ElseIf LCase(Word4) < LCase(Word3) And LCase(Word4) < LCase(Word2) Then
        
        lbl2.Caption = Word4
            If LCase(Word2) < LCase(Word3) Then
                lbl3.Caption = Word2
                lbl4.Caption = Word3
            ElseIf LCase(Word3) < LCase(Word2) Then
                lbl3.Caption = Word3
                lbl4.Caption = Word2
            End If
            
    End If
    
ElseIf LCase(Word2) < LCase(Word1) And LCase(Word2) < LCase(Word3) And LCase(Word2) < LCase(Word4) Then
    
    lbl1.Caption = Word2
    If LCase(Word1) < LCase(Word3) And LCase(Word1) < LCase(Word4) Then
        
        lbl2.Caption = Word1
            If LCase(Word3) < LCase(Word4) Then
                lbl3.Caption = Word3
                lbl4.Caption = Word4
            ElseIf LCase(Word4) < LCase(Word3) Then
                lbl3.Caption = Word4
                lbl4.Caption = Word3
            End If
   
    ElseIf LCase(Word3) < LCase(Word1) And LCase(Word3) < LCase(Word4) Then
        
        lbl2.Caption = Word3
            If LCase(Word1) < LCase(Word4) Then
                lbl3.Caption = Word1
                lbl4.Caption = Word4
            ElseIf LCase(Word4) < LCase(Word1) Then
                lbl3.Caption = Word4
                lbl4.Caption = Word1
            End If
            
    ElseIf LCase(Word4) < LCase(Word3) And LCase(Word4) < LCase(Word1) Then
        
        lbl2.Caption = Word4
            If LCase(Word1) < LCase(Word3) Then
                lbl3.Caption = Word1
                lbl4.Caption = Word3
            ElseIf LCase(Word3) < LCase(Word1) Then
                lbl3.Caption = Word3
                lbl4.Caption = Word1
            End If
            
    End If
   
ElseIf LCase(Word3) < LCase(Word1) And LCase(Word3) < LCase(Word2) And LCase(Word3) < LCase(Word4) Then
    
    lbl1.Caption = Word3
    If LCase(Word1) < LCase(Word2) And LCase(Word1) < LCase(Word4) Then
        
        lbl2.Caption = Word1
            If LCase(Word2) < LCase(Word4) Then
                lbl3.Caption = Word2
                lbl4.Caption = Word4
            ElseIf LCase(Word4) < LCase(Word2) Then
                lbl3.Caption = Word4
                lbl4.Caption = Word2
            End If
   
    ElseIf LCase(Word2) < LCase(Word1) And LCase(Word2) < LCase(Word4) Then
        
        lbl2.Caption = Word2
            If LCase(Word1) < LCase(Word4) Then
                lbl3.Caption = Word1
                lbl4.Caption = Word4
            ElseIf LCase(Word4) < LCase(Word1) Then
                lbl3.Caption = Word4
                lbl4.Caption = Word1
            End If
            
    ElseIf LCase(Word4) < LCase(Word2) And LCase(Word4) < LCase(Word1) Then
        
        lbl2.Caption = Word4
            If LCase(Word1) < LCase(Word2) Then
                lbl3.Caption = Word1
                lbl4.Caption = Word2
            ElseIf LCase(Word2) < LCase(Word1) Then
                lbl3.Caption = Word2
                lbl4.Caption = Word1
            End If
            
    End If
  
ElseIf LCase(Word4) < LCase(Word1) And LCase(Word4) < LCase(Word2) And LCase(Word4) < LCase(Word3) Then
    
    lbl1.Caption = Word4
    If LCase(Word1) < LCase(Word2) And LCase(Word1) < LCase(Word3) Then
        
        lbl2.Caption = Word1
            If LCase(Word2) < LCase(Word3) Then
                lbl3.Caption = Word2
                lbl4.Caption = Word3
            ElseIf LCase(Word3) < LCase(Word2) Then
                lbl3.Caption = Word3
                lbl4.Caption = Word2
            End If
   
    ElseIf LCase(Word2) < LCase(Word1) And LCase(Word2) < LCase(Word3) Then
        
        lbl2.Caption = Word2
            If LCase(Word1) < LCase(Word3) Then
                lbl3.Caption = Word1
                lbl4.Caption = Word3
            ElseIf LCase(Word3) < LCase(Word1) Then
                lbl3.Caption = Word3
                lbl4.Caption = Word1
            End If
            
    ElseIf LCase(Word3) < LCase(Word2) And LCase(Word3) < LCase(Word1) Then
        
        lbl2.Caption = Word3
            If LCase(Word1) < LCase(Word2) Then
                lbl3.Caption = Word1
                lbl4.Caption = Word2
            ElseIf LCase(Word2) < LCase(Word1) Then
                lbl3.Caption = Word2
                lbl4.Caption = Word1
            End If
            
    End If
  
End If

End Sub

Private Sub cmdClear_Click()

txt1.Text = ""
txt2.Text = ""
txt3.Text = ""
txt4.Text = ""

lbl1.Caption = ""
lbl2.Caption = ""
lbl3.Caption = ""
lbl4.Caption = ""

End Sub

Private Sub txt1_Change()

Call cmd_Click

End Sub

Private Sub txt2_Change()

Call cmd_Click

End Sub

Private Sub txt3_Change()

Call cmd_Click

End Sub

Private Sub txt4_Change()

Call cmd_Click

End Sub
