VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   "Map Editor"
   ClientHeight    =   5790
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   7965
   LinkTopic       =   "Form1"
   ScaleHeight     =   5790
   ScaleWidth      =   7965
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picTop 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      ScaleHeight     =   225
      ScaleWidth      =   7935
      TabIndex        =   8
      Top             =   0
      Width           =   7960
      Begin VB.Label lblTile 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tile: 0"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   0
         TabIndex        =   15
         Top             =   0
         Width           =   345
      End
      Begin VB.Label lblSelTile 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Selected Tile: 0"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   720
         TabIndex        =   14
         Top             =   0
         Width           =   900
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name: Noname"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   2760
         TabIndex        =   13
         Top             =   0
         Width           =   960
      End
      Begin VB.Label lblFXType 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FXType: 5"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   1920
         TabIndex        =   12
         Top             =   0
         Width           =   630
      End
      Begin VB.Label lblWalkable 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Walkable: No"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   3960
         TabIndex        =   11
         Top             =   0
         Width           =   780
      End
      Begin VB.Label lblLayer 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Layer: 1/9"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   5040
         TabIndex        =   10
         Top             =   0
         Width           =   600
      End
      Begin VB.Label lblEvent 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Event: None"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   5880
         TabIndex        =   9
         Top             =   0
         Width           =   750
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Properties"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   6360
      TabIndex        =   6
      Top             =   1800
      Width           =   1455
      Begin VB.TextBox txtLayer 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   720
         MaxLength       =   1
         TabIndex        =   17
         Text            =   "1"
         Top             =   540
         Width           =   615
      End
      Begin VB.CheckBox chkWalkable 
         Caption         =   "Walkable"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lblLayerr 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Layer:"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   120
         TabIndex        =   16
         Top             =   600
         Width           =   375
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Flood"
      Height          =   255
      Left            =   6360
      TabIndex        =   5
      Top             =   1440
      Width           =   855
   End
   Begin VB.PictureBox Tile 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   3
      Left            =   6840
      Picture         =   "frmMain.frx":0000
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   25
      TabIndex        =   4
      Top             =   960
      Width           =   375
   End
   Begin VB.PictureBox Tile 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   2
      Left            =   6360
      Picture         =   "frmMain.frx":06BA
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   25
      TabIndex        =   3
      Top             =   960
      Width           =   375
   End
   Begin VB.PictureBox Tile 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   1
      Left            =   6840
      Picture         =   "frmMain.frx":0D74
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   25
      TabIndex        =   2
      Top             =   480
      Width           =   375
   End
   Begin MSComDlg.CommonDialog cm1 
      Left            =   7440
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Tile 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   0
      Left            =   6360
      Picture         =   "frmMain.frx":142E
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   1
      Top             =   480
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5550
      Left            =   0
      ScaleHeight     =   368
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   414
      TabIndex        =   0
      Top             =   240
      Width           =   6240
      Begin VB.Shape Shape2 
         BorderColor     =   &H000000FF&
         Height          =   360
         Left            =   0
         Top             =   0
         Width           =   360
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H000000FF&
         Height          =   360
         Left            =   0
         Top             =   0
         Width           =   360
      End
   End
   Begin VB.Menu mnuMain 
      Caption         =   "Main"
      Begin VB.Menu mnuNew 
         Caption         =   "New Map"
      End
      Begin VB.Menu mnuLine 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "Open Map"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save Map"
      End
      Begin VB.Menu mnuLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Edit"
      Begin VB.Menu mnuRndPattern 
         Caption         =   "Random Pattern"
      End
      Begin VB.Menu mnuAdvProp 
         Caption         =   "Advanced Properties"
      End
      Begin VB.Menu mnuSetName 
         Caption         =   "Set Name"
      End
      Begin VB.Menu mnuEvent 
         Caption         =   "Add Event"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuHlp 
         Caption         =   "Help"
      End
      Begin VB.Menu mnuLine3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sx As Integer, sy As Integer, Sel_tile As Integer
Dim tx As Integer, ty As Integer

Private Sub Command1_Click()
Flood Sel_tile
End Sub

Private Sub Form_Load()
NewMap
DrawGrid
End Sub

Private Sub lblName_Click()
ThisMap.sname = InputBox("Enter a name for the map.", "Enter name")
If ThisMap.sname <> "" Then
lblName.Caption = "Name: " & ThisMap.sname
Else
ThisMap.sname = "Noname"
End If
End Sub

Private Sub mnuAbout_Click()
MsgBox "Tile Map Editor by Hans Bjerndell" & vbCrLf & "For Tilebased games" & vbCrLf & "Copyright © 2002 Hans Bjerndell", vbInformation + vbOKOnly, "About"
End Sub

Private Sub mnuAdvProp_Click()
frmAdvanced.Show
End Sub

Private Sub mnuEvent_Click()
Dim ii As Integer
ii = Mid(lblSelTile.Caption, InStr(lblSelTile.Caption, ":") + 1)
If ThisMap.Tiles(ii).Event <> "" Then
Select Case Left(ThisMap.Tiles(ii).Event, 3)
Case "MSG"
frmEvent.lblTile.Caption = "Tile: " & ii
frmEvent.iIndex = ii
frmEvent.Option1.Value = True
frmEvent.txtMsg.Text = Mid(ThisMap.Tiles(ii).Event, InStr(ThisMap.Tiles(ii).Event, "=") + 1)
frmEvent.Show
Case "WAR"
frmEvent.lblTile.Caption = "Tile: " & ii
frmEvent.iIndex = ii
frmEvent.Option2.Value = True
frmEvent.txtWarp.Text = Mid(ThisMap.Tiles(ii).Event, InStr(ThisMap.Tiles(ii).Event, "=") + 1)
frmEvent.Show
Case "DAM"
frmEvent.lblTile.Caption = "Tile: " & ii
frmEvent.iIndex = ii
frmEvent.Option3.Value = True
frmEvent.txtDamage.Text = Mid(ThisMap.Tiles(ii).Event, InStr(ThisMap.Tiles(ii).Event, "=") + 1)
frmEvent.Show
End Select
Else
frmEvent.lblTile.Caption = "Tile: " & ii
frmEvent.iIndex = ii
frmEvent.Show
End If
End Sub

Private Sub mnuExit_Click()
End
End Sub

Private Sub mnuHlp_Click()
frmHelp.Show
End Sub

Private Sub mnuNew_Click()
NewMap
End Sub

Private Sub mnuOpen_Click()
cm1.Filter = "Supported types |*.cms|"
cm1.ShowOpen
If cm1.Filename <> "" Then
NewMap
LoadMap cm1.Filename
Else
GoTo error:
End If
error:
End Sub

Private Sub mnuRndPattern_Click()
frmRandom.Show
End Sub

Private Sub mnuSave_Click()
On Error GoTo error:
cm1.Filter = "Supported types |*.cms|"
cm1.ShowSave
If cm1.Filename <> "" Then
SaveMap cm1.Filename
Else
GoTo error:
End If
error:
End Sub

Function SaveMap(Filename As String)
Dim i
Open Filename For Output As #1
Print #1, ThisMap.sname
For i = 0 To 287
If ThisMap.Tiles(i).Event <> "" Then
Print #1, ThisMap.Tiles(i).FXType & ":" & ThisMap.Tiles(i).Walkable & ":" & ThisMap.Tiles(i).Layer & "," & Replace(ThisMap.Tiles(i).Event, vbCrLf, "")
Else
Print #1, ThisMap.Tiles(i).FXType & ":" & ThisMap.Tiles(i).Walkable & ":" & ThisMap.Tiles(i).Layer
End If
Next i
Close #1
End Function

Function LoadMap(Filename As String)
Dim x As Integer, y As Integer, rx As Integer, ry As Integer
Dim i, temp As String, arr() As String
Open Filename For Input As #1
Input #1, ThisMap.sname
For i = 0 To 287
Line Input #1, temp
arr = Split(temp, ":")
If UBound(arr()) < 2 Then
MsgBox "This seemes to be an old version, or an unsupported filetype. Unable to open file.", vbCritical + vbOKOnly, "Error"
Close #1
Exit Function
Else
End If
ThisMap.Tiles(i).FXType = CInt(arr(0))
ThisMap.Tiles(i).Walkable = arr(1)
If FindPart(arr(2), ",") = 1 Then
ThisMap.Tiles(i).Layer = Mid(arr(2), 1, 1)
ThisMap.Tiles(i).Event = Mid(arr(2), InStr(arr(2), ",") + 1)
Else
ThisMap.Tiles(i).Layer = arr(2)
End If
If arr(0) = "5" Then
BitBlt Picture1.hDC, rx, ry, 23, 23, Tile(0).hDC, 0, 0, WHITENESS
Else
BitBlt Picture1.hDC, rx, ry, 23, 23, Tile(ThisMap.Tiles(i).FXType).hDC, 0, 0, SRCCOPY
End If
x = x + 1
rx = rx + 23
If x >= 18 Then
y = y + 1
x = 0
ry = ry + 23
rx = 0
End If
Next i
Close #1
Me.Caption = "Map Editor - " & ThisMap.sname
lblName.Caption = "Name: " & ThisMap.sname
Exit Function
End Function

Private Sub mnuSetName_Click()
lblName_Click
End Sub

Private Sub Picture1_DblClick()
If ThisMap.Tiles(GetTile(tx, ty)).Event <> "" Then
Select Case Left(ThisMap.Tiles(GetTile(tx, ty)).Event, 3)
Case "MSG"
frmEvent.lblTile.Caption = "Tile: " & GetTile(tx, ty)
frmEvent.iIndex = GetTile(tx, ty)
frmEvent.Option1.Value = True
frmEvent.txtMsg.Text = Mid(ThisMap.Tiles(GetTile(tx, ty)).Event, InStr(ThisMap.Tiles(GetTile(tx, ty)).Event, "=") + 1)
frmEvent.Show
Case "WAR"
frmEvent.lblTile.Caption = "Tile: " & GetTile(tx, ty)
frmEvent.iIndex = GetTile(tx, ty)
frmEvent.Option2.Value = True
frmEvent.txtWarp.Text = Mid(ThisMap.Tiles(GetTile(tx, ty)).Event, InStr(ThisMap.Tiles(GetTile(tx, ty)).Event, "=") + 1)
frmEvent.Show
Case "DAM"
frmEvent.lblTile.Caption = "Tile: " & GetTile(tx, ty)
frmEvent.iIndex = GetTile(tx, ty)
frmEvent.Option3.Value = True
frmEvent.txtDamage.Text = Mid(ThisMap.Tiles(GetTile(tx, ty)).Event, InStr(ThisMap.Tiles(GetTile(tx, ty)).Event, "=") + 1)
frmEvent.Show
End Select
Else
frmEvent.lblTile.Caption = "Tile: " & GetTile(tx, ty)
frmEvent.iIndex = GetTile(tx, ty)
frmEvent.Show
End If
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 1 Then
BitBlt Picture1.hDC, Shape1.Left, Shape1.Top, 23, 23, Tile(Sel_tile).hDC, 0, 0, SRCCOPY
ThisMap.Tiles(GetTile(tx, ty)).FXType = Sel_tile
ThisMap.Tiles(GetTile(tx, ty)).Walkable = chkWalkable.Value
ThisMap.Tiles(GetTile(tx, ty)).Layer = txtLayer.Text
ElseIf Button = 2 Then
BitBlt Picture1.hDC, Shape1.Left, Shape1.Top, 23, 23, Tile(Sel_tile).hDC, 0, 0, WHITENESS
ThisMap.Tiles(GetTile(tx, ty)).FXType = 5
ThisMap.Tiles(GetTile(tx, ty)).Walkable = False
ThisMap.Tiles(GetTile(tx, ty)).Layer = 1
DrawGrid
End If
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
sx = Replace(Mid(x / 23, 1, 2), ".", "")
sy = Replace(Mid(y / 23, 1, 2), ".", "")
If sx > 17 Then Exit Sub
If sy > 15 Then Exit Sub
If sx < 0 Then Exit Sub
If sy < 0 Then Exit Sub
tx = sx
ty = sy
'Me.Caption = sx & " " & sy
lblTile.Caption = "Tile: " & GetTile(sx, sy)

GetFXType GetTile(sx, sy)
GetWalkable GetTile(sx, sy)
GetLayer GetTile(sx, sy)
GetEvent GetTile(sx, sy)

sy = sy * 23
sx = sx * 23
Shape1.Left = sx
Shape1.Top = sy



If Button = 1 Then
BitBlt Picture1.hDC, Shape1.Left, Shape1.Top, 23, 23, Tile(Sel_tile).hDC, 0, 0, SRCCOPY
ThisMap.Tiles(GetTile(tx, ty)).FXType = Sel_tile
ThisMap.Tiles(GetTile(tx, ty)).Walkable = chkWalkable.Value
ThisMap.Tiles(GetTile(tx, ty)).Layer = txtLayer.Text
ElseIf Button = 2 Then
BitBlt Picture1.hDC, Shape1.Left, Shape1.Top, 23, 23, Tile(Sel_tile).hDC, 0, 0, WHITENESS
ThisMap.Tiles(GetTile(tx, ty)).FXType = 5
ThisMap.Tiles(GetTile(tx, ty)).Walkable = False
ThisMap.Tiles(GetTile(tx, ty)).Layer = 1
DrawGrid
End If
End Sub

Function GetTile(x As Integer, y As Integer) As Integer
GetTile = y * 18 + x
End Function

Function GetEvent(iTile As Integer)
If ThisMap.Tiles(iTile).Event <> "" Then

    Select Case Left(ThisMap.Tiles(iTile).Event, 3)
    Case "MSG"
    lblEvent.Caption = "Event: MSG"
    Case "WAR"
    lblEvent.Caption = "Event: WARP"
    Case "DAM"
    lblEvent.Caption = "Event: Damage Point"
    End Select

Else
lblEvent.Caption = "Event: None"
End If
End Function

Function GetLayer(iTile As Integer)
lblLayer.Caption = "Layer: " & ThisMap.Tiles(iTile).Layer & "/9"
End Function

Function GetFXType(iTile As Integer)
lblFXType.Caption = "FXType: " & ThisMap.Tiles(iTile).FXType
End Function

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
lblSelTile.Caption = "Selceted Tile: " & GetTile(tx, ty)
Shape2.Left = tx * 23
Shape2.Top = ty * 23
End Sub

Private Sub Tile_Click(Index As Integer)
Sel_tile = Index
Release
Tile(Index).BorderStyle = 1
End Sub

Function Release()
Dim i
For i = 0 To 3
Tile(i).BorderStyle = 0
Next i
End Function

Function NewMap()
Picture1.Cls
Dim i
For i = 0 To 287
ThisMap.Tiles(i).FXType = 5
ThisMap.Tiles(i).Walkable = 0
ThisMap.Tiles(i).Layer = 1
ThisMap.Tiles(i).Event = ""
Next i
DrawGrid
End Function

Function GetWalkable(iTile As Integer)
If ThisMap.Tiles(iTile).Walkable = 1 Then
lblWalkable.Caption = "Walkable: Yes"
Else
lblWalkable.Caption = "Walkable: No"
End If
End Function
Function Flood(iTile As Integer)
Dim i, x As Integer, y As Integer, rx As Integer, ry As Integer
Do Until i = 288
DoEvents
If x >= 18 Then
ry = ry + 23
y = y + 1
x = 0
rx = 0
End If
BitBlt Picture1.hDC, rx, ry, 23, 23, Tile(iTile).hDC, 0, 0, SRCCOPY
ThisMap.Tiles(i).Walkable = chkWalkable.Value
ThisMap.Tiles(i).Layer = txtLayer.Text
ThisMap.Tiles(i).FXType = iTile
Picture1.Refresh
i = i + 1
x = x + 1
rx = rx + 23
Loop
End Function

Function DrawGrid()
Dim x As Integer, y As Integer
Do Until x > Picture1.ScaleWidth And y > Picture1.ScaleHeight
DoEvents

    Picture1.Line (x, 0)-(x, Picture1.Height)
    Picture1.Line (0, y)-(Picture1.Width, y)
    
y = y + 23
x = x + 23
Loop
End Function
Private Sub txtLayer_KeyPress(KeyAscii As Integer)
KeyAscii = IIf(Not KeyAscii = 8 And Not Val((Chr(KeyAscii))) > 0, 0, KeyAscii)
End Sub
