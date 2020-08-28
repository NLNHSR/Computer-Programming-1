VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmEdit 
   Caption         =   "MiniEdit"
   ClientHeight    =   6180
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   7425
   LinkTopic       =   "Form1"
   ScaleHeight     =   6180
   ScaleWidth      =   7425
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer 
      Interval        =   500
      Left            =   6180
      Top             =   1320
   End
   Begin RichTextLib.RichTextBox txtEdit 
      Height          =   4695
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   8281
      _Version        =   393217
      ScrollBars      =   2
      TextRTF         =   $"frmEdit.frx":0000
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuNew 
         Caption         =   "New "
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "Open"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save"
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "SaveAs"
      End
      Begin VB.Menu mnuSeperator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Edit"
      Begin VB.Menu mnuCopy 
         Caption         =   "Copy"
      End
      Begin VB.Menu mnuCut 
         Caption         =   "Cut"
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "Paste"
      End
      Begin VB.Menu mnuSelectAll 
         Caption         =   "Select All"
      End
      Begin VB.Menu mnuFindReplace 
         Caption         =   "Find and Replace"
      End
   End
   Begin VB.Menu mnuFont 
      Caption         =   "Font"
      Begin VB.Menu mnuSize 
         Caption         =   "Size"
      End
      Begin VB.Menu mnuType 
         Caption         =   "Type"
         Begin VB.Menu mnuBold 
            Caption         =   "Bold"
         End
         Begin VB.Menu mnuItalicize 
            Caption         =   "Italicize"
         End
         Begin VB.Menu mnuUnderline 
            Caption         =   "Underline"
         End
      End
      Begin VB.Menu mnuColor 
         Caption         =   "Color"
         Begin VB.Menu mnuRed 
            Caption         =   "Red"
         End
         Begin VB.Menu mnuBlue 
            Caption         =   "Blue"
         End
         Begin VB.Menu mnuGreen 
            Caption         =   "Green"
         End
         Begin VB.Menu mnuBlack 
            Caption         =   "Black"
         End
      End
   End
End
Attribute VB_Name = "frmEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Filename As String
Dim Change As Boolean
Dim Path As String

Private Sub Form_Load()

ChDrive "D:\19 S2 NS CP1\MiniEdit\MEdit"
Filename = ""
Change = False

End Sub

Private Sub Form_Resize()

txtEdit.Width = frmEdit.ScaleWidth
txtEdit.Height = frmEdit.ScaleHeight

End Sub

Private Sub mnuBlack_Click()

txtEdit.SelColor = &H0&

End Sub

Private Sub mnuBlue_Click()

txtEdit.SelColor = &HFF0000

End Sub

Private Sub mnuBold_Click()

If txtEdit.SelBold = False Then
    txtEdit.SelBold = True
ElseIf txtEdit.SelBold = True Then
    txtEdit.SelBold = False
End If

End Sub

Private Sub mnuCopy_Click()

Clipboard.SetText (txtEdit.SelText)

End Sub

Private Sub mnuCut_Click()

Clipboard.SetText (txtEdit.SelText)
txtEdit.SelText = ""

End Sub

Private Sub mnuExit_Click()

If Change = True Then
    If Filename = "" Then
        Call mnuSaveAs_Click
        End
    ElseIf Filename <> "" Then
        Call mnuSave_Click
        End
    Else
        End
    End If
Else
    End
End If

End Sub

Private Sub mnuFindReplace_Click()

Dim Find As String, Replace As String
Dim Start As Long

Find = InputBox("Find what?", "Find and Replace")
Replace = InputBox("Replace with:", "Find and Replace")

Do Until Start = -1
    Start = InStr(1, txtEdit.Text, Find, vbTextCompare)
    If Start = -1 Or Start = 0 Then
        Exit Do
    Else
        txtEdit.SelStart = Start - 1
        txtEdit.SelLength = Len(Find)
        txtEdit.SelText = Replace
    End If
Loop
        
End Sub

Private Sub mnuGreen_Click()

txtEdit.SelColor = &HC000&

End Sub

Private Sub mnuItalicize_Click()

If txtEdit.SelItalic = False Then
    txtEdit.SelItalic = True
ElseIf txtEdit.SelItalic = True Then
    txtEdit.SelItalic = False
End If

End Sub

Private Sub mnuNew_Click()

If Change = True Then
    mnuSaveAs_Click
End If
txtEdit.Text = ""
Filename = ""
Change = False

End Sub

Private Sub mnuOpen_Click()

txtEdit.Text = ""
Dim strFn, Path As String
strFn = UCase$(Trim$(InputBox("Filename", "Open File")))
Path = "D:\19 S2 NS CP1\MiniEdit\MEdit\" + strFn + ".txt"
Filename = strFn
Open Path For Input As #1
Dim FileSize As Integer
FileSize = LOF(1)
txtEdit = Input$(FileSize, #1)
Close #1

End Sub

Private Sub mnuPaste_Click()

If Clipboard.GetFormat(vbCFText) And Not Clipboard.GetFormat(vbCFRTF) Then
    txtEdit.SelText = Clipboard.GetText(vbCFText)
    txtEdit.SetFocus
End If

End Sub

Private Sub mnuRed_Click()

txtEdit.SelColor = &HFF&

End Sub

Private Sub mnuSave_Click()

Dim Path As String
If Filename <> "" Then
    Path = "D:\19 S2 NS CP1\MiniEdit\MEdit\" + Filename + ".txt"
    Open Path For Output As #1
    Print #1, txtEdit
    Close #1
Else
    Call mnuSaveAs_Click
End If

End Sub

Private Sub mnuSaveAs_Click()

Dim strFn, Path As String
strFn = UCase$(Trim$(InputBox("Filename", "Save As...")))
Path = "D:\19 S2 NS CP1\MiniEdit\MEdit\" + strFn + ".txt"
Filename = strFn
Open Path For Output As #1
Print #1, txtEdit
Close #1

End Sub

Private Sub mnuSelectAll_Click()

With txtEdit
   .SelStart = 0
   .SelLength = Len(.Text)
End With

End Sub

Private Sub mnuSize_Click()

Dim size As String
size = UCase$(Trim$(InputBox("Enter a number", "Text Size")))
txtEdit.SelFontSize = size

End Sub

Private Sub mnuUnderline_Click()

If txtEdit.SelUnderline = False Then
    txtEdit.SelUnderline = True
ElseIf txtEdit.SelUnderline = True Then
    txtEdit.SelUnderline = False
End If

End Sub

Private Sub txtEdit_Change()

Change = True

End Sub
