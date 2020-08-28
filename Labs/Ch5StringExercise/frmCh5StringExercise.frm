VERSION 5.00
Begin VB.Form frmCh5StringExercise 
   Caption         =   "Form1"
   ClientHeight    =   7725
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13170
   LinkTopic       =   "Form1"
   ScaleHeight     =   7725
   ScaleWidth      =   13170
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmCh5StringExercise"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

Dim Fruit As String, Word As String
Dim Word1 As String, Word2 As String, Msg As String

Fruit = "apple"
Word = "aple"
If Fruit = Word Then
    Msg = "The strings are equal."
Else
    Msg = Fruit & " and " & Word & " are not equal "
End If
MsgBox Msg

Word1 = "play"
Word2 = "ball"
If Word1 > Word2 Then
    Msg = Word1 & " is greater than " & Word2
Else
    Msg = Word1 & " is less than " & Word2
End If
MsgBox Msg

Word1 = "Play"
If Word1 > Word2 Then
    Msg = Word1 & " is greater than " & Word2
Else
    Msg = Word1 & " is less than " & Word2
End If
MsgBox Msg

If UCase(Word1) > UCase(Word2) Then
    Msg = UCase(Word1) & " is greater than " & UCase(Word2)
Else
    Msg = UCase(Word1) & " is less than " & UCase(Word2)
End If
MsgBox Msg

Dim Result As Integer
Word1 = "play"
Result = StrComp(Word1, Word2, 1)
If Result = 0 Then
    Msg = Word1 & " is equal to " & Word2
ElseIf Result = -1 Then
    Msg = Word1 & "is less than " & Word2
ElseIf Result = 1 Then
    Msg = Word1 & " is greater than " & Word2
End If
MsgBox Msg

Msg = "The length of " & Word1 & " is " & Str$(Len(Word1))
MsgBox Msg

End Sub
