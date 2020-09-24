VERSION 5.00
Begin VB.Form zchild 
   BackColor       =   &H8000000C&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6345
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7440
   DrawStyle       =   2  'Dot
   FillColor       =   &H80000002&
   ForeColor       =   &H80000001&
   Icon            =   "zchild.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   423
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   496
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox measure 
      AutoSize        =   -1  'True
      Height          =   855
      Left            =   4680
      ScaleHeight     =   53
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   61
      TabIndex        =   0
      Top             =   3240
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Shape cir 
      BorderColor     =   &H00FF0000&
      Height          =   255
      Left            =   1320
      Shape           =   3  'Circle
      Top             =   1680
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Line xline 
      BorderColor     =   &H000000FF&
      BorderStyle     =   3  'Dot
      Index           =   0
      Visible         =   0   'False
      X1              =   96
      X2              =   232
      Y1              =   120
      Y2              =   120
   End
End
Attribute VB_Name = "zchild"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'mouse coords on mousedown
Dim xx, yy, oldx, oldy
Dim rad, linec As Integer




Private Sub Form_GotFocus()
zcoordfrm.Visible = False
End Sub

Private Sub Form_Load()
tool = 1
linec = 0
    Me.Left = (zmapper.Width - (Me.Width / 2)) / 2
End Sub

Private Sub Form_MouseDown(button As Integer, Shift As Integer, X As Single, Y As Single)
'stores x and y coords in xx and yy
If button = 1 Then

  
    If tool = 3 And linec = 0 Then
        cir.Top = Y - (cir.Height / 2)
        cir.Left = X - (cir.Width / 2)
        cir.Visible = True
        oldx = X
        oldy = Y
        xx = X
        yy = Y
        linec = 0
        zcoordfrm.poly.AddItem "start"
        zcoordfrm.poly.AddItem xx & "," & yy & ","
        Exit Sub
    ElseIf tool = 2 Then
          oldx = X: oldy = Y
          xx = oldx: yy = oldy
          Exit Sub
    ElseIf tool = 3 And linec <> 0 Then
        Exit Sub
    End If
    xx = X
    yy = Y
    rad = 1
        

Else
    Select Case tool
        Case "1"
            PopupMenu zmapper.tools, 8, , , zmapper.rect
        Case "2"
            PopupMenu zmapper.tools, 8, , , zmapper.cir
        Case "3"
            PopupMenu zmapper.tools, 8, , , zmapper.poly
    End Select
End If
End Sub

Private Sub Form_MouseMove(button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next


Select Case tool
Case "1"
    zmapper.StatusBar1.Panels(2).Text = "X1: " & xx & " Y1: " & yy & " X2: " & X & " Y2: " & Y
Case "2"
    zmapper.StatusBar1.Panels(2).Text = "X1: " & xx & " Y1: " & yy & " Radius: " & Y
Case "3"
    zmapper.StatusBar1.Panels(2).Text = "X: " & X & " Y: " & Y & " (statrting positions X: " & xx & " Y: " & yy & ")"
End Select

If button = 1 Then

Select Case tool
Case "1"
'clear form
    Me.Cls
'draw box
    Me.Line (xx, yy)-(X, Y), RGB(100, 100, 255), B
    'displaying coords
    zmapper.StatusBar1.Panels(2).Text = "X1: " & xx & " Y1: " & yy & " X2: " & X & " Y2: " & Y

    
Case "2"
'displaying coords
 '   Label1.Caption = "x1 = " & xx & " y1 = " & yy & " X = " & X & " Y = " & Y
'clear form
    Me.Cls
'draw circle
    'Me.Circle (xx, yy), Y, RGB(100, 100, 255)
    Circle (oldx, oldy), Sqr((xx - oldx) ^ 2 + (yy - oldy) ^ 2), vbWhite
    Line (oldx, oldy)-(xx, yy), vbWhite

    Caption = "Radius= " & CStr(Sqr((X - oldx) ^ 2 + (Y - oldy) ^ 2))
    xx = X: yy = Y
    
        'displaying coords
    zmapper.StatusBar1.Panels(2).Text = "X1: " & xx & " Y1: " & yy & " Radius: " & Y

Case "3"

If linec <> 0 Then
    Load xline(linec)
End If
    xline(linec).Visible = True
    xline(linec).X1 = oldx
    xline(linec).Y1 = oldy
    xline(linec).X2 = X
    xline(linec).Y2 = Y
    zmapper.StatusBar1.Panels(2).Text = "X: " & X & " Y: " & Y & " (statrting positions X: " & xx & " Y: " & yy
    
End Select



End If
End Sub

Private Sub Form_MouseUp(button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo mder
Dim yn As String
If button = 1 Then
Select Case tool
Case "1"
'draw final box
    Me.Line (xx, yy)-(X, Y), RGB(255, 0, 0), B
'calls function to find out what was selected
    'Label2.Caption = was_anything_selected(xx, yy, X, Y)
    'Label3.Caption = was_anything_selected_m2(xx, yy, X, Y)
    
    yn = InputBox("The Coordinates are " & zmapper.StatusBar1.Panels(2).Text & ". Please enter a URL.", "Enter URL", "http://www.somesite.com")
    If yn = "" Then yn = "http://www.somesite.com"
    zcoordfrm.poly.AddItem "#" & yn
    
    GoTo addq
Case "2"
    Cls
    Circle (oldx, oldy), Sqr((xx - oldx) ^ 2 + (yy - oldy) ^ 2), vbRed
    Line (oldx, oldy)-(xx, yy), vbBlue
    yn = InputBox("The Coordinates are " & zmapper.StatusBar1.Panels(2).Text & ". Please enter a URL.", "Enter URL", "http://www.somesite.com")
    If yn = "" Then yn = "http://www.somesite.com"
    zcoordfrm.poly.AddItem "#" & yn
    GoTo addq
Case "3"

   If (xline.Count - 1) <> linec Then
        Load xline(linec)
        xline(linec).Visible = True
    End If
    xline(linec).Y1 = oldy
    xline(linec).X1 = oldx
    xline(linec).X2 = X
    xline(linec).Y2 = Y
    oldx = X
    oldy = Y
    
    zcoordfrm.poly.AddItem X & "," & Y & ","
    
    If X <= xx + (cir.Width / 2) And Y <= yy + (cir.Height / 2) And X >= xx - (cir.Width / 2) And Y >= yy - (cir.Height / 2) And xline.Count <> 1 Then
        xline(linec).X2 = xx
        xline(linec).Y2 = yy
        zmapper.StatusBar1.Panels(2).Text = "X: " & xx & " Y: " & yy & " (statrting positions X: " & xx & " Y: " & yy & ")"
        yn = InputBox("The Coordinates are " & zmapper.StatusBar1.Panels(2).Text & ". Please enter a URL.", "Enter URL", "http://www.somesite.com")
        If yn = "" Then yn = "http://www.somesite.com"
        zcoordfrm.poly.AddItem "#" & yn
        zcoordfrm.poly.AddItem "end"
        GoTo addq
    End If
    
    linec = linec + 1
End Select
End If
Exit Sub

addq:



    Select Case tool
        Case "1"
            zcoordfrm.recta.AddItem "<area href =" & Chr(34) & yn & Chr(34) & " shape =" & Chr(34) & "rect" & Chr(34) & " coords =" & Chr(34) & xx & ", " & yy & ", " & X & ", " & Y & Chr(34) & ">"
            Me.Cls
        Case "2"
            zcoordfrm.circ.AddItem "<area href =" & Chr(34) & yn & Chr(34) & " shape =" & Chr(34) & "circle" & Chr(34) & " coords =" & Chr(34) & xx & ", " & yy & ", " & Y & Chr(34) & ">"
            Me.Cls
        Case "3"
            polytext = polycode()
            unloadpoly
End Select
    
    Me.Cls
    If tool = 3 Then unloadpoly

Exit Sub
mder:
MsgBox Error
End Sub


Function polycode() As String
Dim i As Integer
Dim url As String

i = 0
Do
If zcoordfrm.poly.List(i) = "start" Then
    GoTo skip
End If

If Left(zcoordfrm.poly.List(i), 1) = "#" Then url = Right(zcoordfrm.poly.List(i), Len(zcoordfrm.poly.List(i)) - 1): GoTo skip

If zcoordfrm.poly.List(i) = "end" Then
    xpolycode(polycount) = "<area href =" & Chr(34) & url & Chr(34) & " shape =" & Chr(34) & "polygon" & Chr(34) & " coords =" & Chr(34) & Left(pcode(polycount), Len(pcode(polycount)) - 1) & Chr(34) & ">" & vbCrLf
    polycode = polycode & xpolycode(polycount)
    polycount = polycount + 1
    GoTo skip
End If

pcode(polycount) = pcode(polycount) & zcoordfrm.poly.List(i)
skip:
i = i + 1
Loop Until i = zcoordfrm.poly.ListCount




End Function


Sub unloadpoly()
On Error Resume Next
Dim i As Integer
i = 0
Do
    Unload xline(i)
    i = i + 1
Loop Until i = linec + 1

xline(0).Visible = False
cir.Visible = False
linec = 0

End Sub
