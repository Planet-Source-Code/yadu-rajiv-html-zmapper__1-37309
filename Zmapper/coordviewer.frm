VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form zcoordfrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "mapper "
   ClientHeight    =   3930
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8190
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000A&
   Icon            =   "coordviewer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3930
   ScaleWidth      =   8190
   Begin VB.CommandButton Command4 
      Caption         =   "Remap"
      Height          =   375
      Left            =   3000
      TabIndex        =   7
      Top             =   3480
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "change map name"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   3480
      Width           =   1335
   End
   Begin VB.ListBox poly 
      Height          =   2205
      Left            =   2880
      TabIndex        =   5
      Top             =   5760
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.ListBox recta 
      Height          =   2205
      Left            =   120
      TabIndex        =   4
      Top             =   5760
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.ListBox circ 
      Height          =   2205
      Left            =   5640
      TabIndex        =   3
      Top             =   5760
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      Caption         =   "copy"
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   3480
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Hide"
      Height          =   375
      Left            =   6840
      TabIndex        =   1
      Top             =   3480
      Width           =   1215
   End
   Begin RichTextLib.RichTextBox coordsrtb 
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   5530
      _Version        =   393217
      ScrollBars      =   3
      Appearance      =   0
      TextRTF         =   $"coordviewer.frx":030A
   End
End
Attribute VB_Name = "zcoordfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim coordslstx As String

Private Sub Command1_Click()
Me.Visible = False

End Sub

Private Sub Command2_Click()
On Error GoTo copyer
Clipboard.SetText coordsrtb.Text
MsgBox "Succesfully copied to clipboard."
Exit Sub
copyer:  MsgBox "Could not be copied"
End Sub

Private Sub Command3_Click()
remap:
mapname = InputBox("Enter Mapname", "mapper 1.0", mapname)
If Trim(mapname) = "" Then
    GoTo remap
End If

makecode

End Sub

Private Sub Command4_Click()
makecode
End Sub


Private Sub Form_Load()
AlwaysOnTop Me
makecode
Me.Caption = Me.Caption & App.Major & "." & App.Minor & " Built: " & App.Revision
End Sub


Sub makecode()
On Error Resume Next
Dim i As Integer
coordsrtb.Text = "<map name = " & Chr(34) & mapname & Chr(34) & ">" & vbCrLf

i = 0
If recta.ListCount <> 0 Then
    Do
        coordsrtb.Text = coordsrtb.Text & recta.List(i) & vbCrLf
        i = i + 1
    Loop Until i = recta.ListCount
End If

i = 0
If circ.ListCount <> 0 Then
    Do
        coordsrtb.Text = coordsrtb.Text & circ.List(i) & vbCrLf
        i = i + 1
    Loop Until i = circ.ListCount
End If


i = 0
If poly.ListCount <> 0 Then
        coordsrtb.Text = coordsrtb.Text & polytext
End If

i = 0
coordsrtb.Text = coordsrtb.Text & "</map>"

End Sub

Private Sub Form_LostFocus()
Me.Visible = False
End Sub

