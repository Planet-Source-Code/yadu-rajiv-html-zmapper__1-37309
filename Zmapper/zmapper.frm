VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.MDIForm zmapper 
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000C&
   Caption         =   "Mapper"
   ClientHeight    =   6450
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   7755
   Icon            =   "zmapper.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   6075
      Width           =   7755
      _ExtentX        =   13679
      _ExtentY        =   661
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   2
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Text            =   "Ready"
            TextSave        =   "Ready"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Bevel           =   0
            TextSave        =   ""
            Key             =   "xx"
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComDlg.CommonDialog mapperdlg 
      Left            =   1560
      Top             =   2880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Menu mapper 
      Caption         =   "&Mapper"
      Begin VB.Menu opn 
         Caption         =   "&Open image"
         Shortcut        =   ^O
      End
      Begin VB.Menu about 
         Caption         =   "&About"
      End
      Begin VB.Menu div1 
         Caption         =   "-"
      End
      Begin VB.Menu exit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu tools 
      Caption         =   "&Tools"
      Begin VB.Menu cir 
         Caption         =   "&Circle"
         Shortcut        =   ^E
      End
      Begin VB.Menu rect 
         Caption         =   "&Rectangle"
         Shortcut        =   ^R
      End
      Begin VB.Menu poly 
         Caption         =   "&Polygon"
         Shortcut        =   ^P
      End
      Begin VB.Menu div 
         Caption         =   "-"
      End
      Begin VB.Menu dispcoords 
         Caption         =   "Create coor&ds"
         Shortcut        =   ^D
      End
   End
End
Attribute VB_Name = "zmapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' deja_vu
' feedback : deja_vu555@yahoo.com
' sorry for not commenting this.. done over night because i was angry..
' i couldnt find a decent image mapper made in visual basic.. :p
'

Dim openflag As Boolean

Private Sub about_Click()
MsgBox "iMage mapper version: " & App.Major & "." & App.Minor & " Built: " & App.Revision & vbCrLf & "by deja_vu..", vbInformation
End Sub

Private Sub cir_Click()
tool = 2
End Sub

Private Sub dispcoords_Click()
zcoordfrm.Show
zcoordfrm.makecode
End Sub

Private Sub exit_Click()
Unload Me
End Sub

Private Sub MDIForm_Load()
On Error GoTo loader


openflag = False

mapperdlg.Filter = "image files(*.jpg,*.gif)|*.jpg;*.gif|"
mapperdlg.FLAGS = cdlOFNPathMustExist Or cdlOFNHideReadOnly


zcoordfrm.poly.Clear
zcoordfrm.circ.Clear
zcoordfrm.recta.Clear



Me.Caption = "Mage mapper version: " & App.Major & "." & App.Minor & " Built: " & App.Revision

mapperdlg.ShowOpen

zchild.measure.Picture = LoadPicture(mapperdlg.FileName)
zchild.Picture = zchild.measure.Picture
zchild.Width = zchild.measure.Width * 15
zchild.Height = zchild.measure.Height * 15

mapname = InputBox("Enter Mapname", "mapper 1.0", "mymap" & Now)
zchild.Show

Exit Sub
loader: If Error = "Cancel was selected." Then
            Exit Sub
        Else
            MsgBox Error, vbCritical
            Unload Me
        End If
End Sub


Private Sub opn_Click()
On Error GoTo opner

Dim yn As Integer

If openflag = True Then
    yn = MsgBox("All mapping information will be lost.. Continue?", vbYesNo)
    If yn = 7 Then
        Exit Sub
    End If
End If

openflag = True

mapperdlg.ShowOpen

zchild.measure.Picture = LoadPicture(mapperdlg.FileName)
zchild.Picture = zchild.measure.Picture
zchild.Width = zchild.measure.Width * 15
zchild.Height = zchild.measure.Height * 15

zchild.Show

mapname = InputBox("Enter Mapname", "mapper 1.0", Now)

zcoordfrm.poly.Clear
zcoordfrm.circ.Clear
zcoordfrm.recta.Clear

polycount = 0

Exit Sub
opner: If Error = "Cancel was selected." Then
            Exit Sub
        Else
            MsgBox Error, vbCritical
            Unload Me
        End If
End Sub

Private Sub poly_Click()
tool = 3
End Sub

Private Sub rect_Click()
tool = 1
End Sub
