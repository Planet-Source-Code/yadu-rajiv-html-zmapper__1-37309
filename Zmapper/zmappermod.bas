Attribute VB_Name = "Module1"
Option Explicit

Declare Function SetWindowPos Lib "user32" (ByVal h%, ByVal hb%, ByVal X%, ByVal Y%, ByVal cx%, ByVal cy%, ByVal F%) As Integer

Public tool As Integer

Public mapname As String

Public polycount As Integer

Public pcode(1000) As String

Public polytext As String

Public xpolycode(1000) As String






'makes the form stay always on top
Sub AlwaysOnTop(frmID As Form)
Dim ontop
   Const SWP_NOMOVE = 2
   Const SWP_NOSIZE = 1
   Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
   Const HWND_TOPMOST = -1
   Const HWND_NOTOPMOST = -2
   ontop = SetWindowPos(frmID.hWnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
End Sub


'declaration for mapper
