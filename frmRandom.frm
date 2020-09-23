VERSION 5.00
Begin VB.Form frmRandom 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Random Pattern"
   ClientHeight    =   2955
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   1650
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2955
   ScaleWidth      =   1650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox SelTile 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   3
      Left            =   120
      Picture         =   "frmRandom.frx":0000
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   12
      Top             =   2040
      Width           =   375
   End
   Begin VB.PictureBox SelTile 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   2
      Left            =   120
      Picture         =   "frmRandom.frx":06BA
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   11
      Top             =   1440
      Width           =   375
   End
   Begin VB.PictureBox SelTile 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   1
      Left            =   120
      Picture         =   "frmRandom.frx":0D74
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   10
      Top             =   120
      Width           =   375
   End
   Begin VB.PictureBox SelTile 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   0
      Left            =   120
      Picture         =   "frmRandom.frx":142E
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   9
      Top             =   840
      Width           =   375
   End
   Begin VB.CheckBox Options 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Include"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   195
      Index           =   0
      Left            =   600
      MaskColor       =   &H00000000&
      TabIndex        =   8
      Top             =   120
      Value           =   1  'Checked
      Width           =   855
   End
   Begin VB.CheckBox Options 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Walkable"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   195
      Index           =   1
      Left            =   600
      MaskColor       =   &H00000000&
      TabIndex        =   7
      Top             =   360
      Value           =   1  'Checked
      Width           =   855
   End
   Begin VB.CheckBox Options 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Include"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   195
      Index           =   2
      Left            =   600
      MaskColor       =   &H00000000&
      TabIndex        =   6
      Top             =   840
      Value           =   1  'Checked
      Width           =   855
   End
   Begin VB.CheckBox Options 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Walkable"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   195
      Index           =   3
      Left            =   600
      MaskColor       =   &H00000000&
      TabIndex        =   5
      Top             =   1080
      Value           =   1  'Checked
      Width           =   855
   End
   Begin VB.CheckBox Options 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Include"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   195
      Index           =   4
      Left            =   600
      MaskColor       =   &H00000000&
      TabIndex        =   4
      Top             =   1440
      Value           =   1  'Checked
      Width           =   855
   End
   Begin VB.CheckBox Options 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Walkable"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   195
      Index           =   5
      Left            =   600
      MaskColor       =   &H00000000&
      TabIndex        =   3
      Top             =   1680
      Value           =   1  'Checked
      Width           =   855
   End
   Begin VB.CheckBox Options 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Include"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   195
      Index           =   6
      Left            =   600
      MaskColor       =   &H00000000&
      TabIndex        =   2
      Top             =   2040
      Value           =   1  'Checked
      Width           =   855
   End
   Begin VB.CheckBox Options 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Walkable"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   195
      Index           =   7
      Left            =   600
      MaskColor       =   &H00000000&
      TabIndex        =   1
      Top             =   2280
      Width           =   855
   End
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "Generate"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   2640
      Width           =   1455
   End
End
Attribute VB_Name = "frmRandom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdGenerate_Click()
Generate
End Sub

Function Generate()
Dim x As Integer, y As Integer, rx As Integer, ry As Integer
Dim i, temp As Integer
For i = 0 To 287
    
    temp = Int(Rnd * 4)
    
    Select Case temp
    Case 0
    If Options(0).Value = 1 Then
    BitBlt frmMain.Picture1.hDC, rx, ry, 23, 23, frmMain.Tile(0).hDC, 0, 0, SRCCOPY
    ThisMap.Tiles(i).FXType = 0
    If Options(1).Value = 1 Then
    ThisMap.Tiles(i).Walkable = 1
    Else
    ThisMap.Tiles(i).Walkable = 0
    End If
    Else
    End If
    Case 1
    If Options(2).Value = 1 Then
    BitBlt frmMain.Picture1.hDC, rx, ry, 23, 23, frmMain.Tile(1).hDC, 0, 0, SRCCOPY
    ThisMap.Tiles(i).FXType = 1
    If Options(3).Value = 1 Then
    ThisMap.Tiles(i).Walkable = 1
    Else
    ThisMap.Tiles(i).Walkable = 0
    End If
    Else
    End If
    Case 2
    If Options(4).Value = 1 Then
    BitBlt frmMain.Picture1.hDC, rx, ry, 23, 23, frmMain.Tile(2).hDC, 0, 0, SRCCOPY
    ThisMap.Tiles(i).FXType = 2
    If Options(5).Value = 1 Then
    ThisMap.Tiles(i).Walkable = 1
    Else
    ThisMap.Tiles(i).Walkable = 0
    End If
    Else
    End If
    Case 3
    If Options(6).Value = 1 Then
    BitBlt frmMain.Picture1.hDC, rx, ry, 23, 23, frmMain.Tile(3).hDC, 0, 0, SRCCOPY
    ThisMap.Tiles(i).FXType = 3
    If Options(7).Value = 1 Then
    ThisMap.Tiles(i).Walkable = 1
    Else
    ThisMap.Tiles(i).Walkable = 0
    End If
    Else
    End If
    End Select
    frmMain.Picture1.Refresh
    x = x + 1
    rx = rx + 23
    If x >= 18 Then
    y = y + 1
    ry = ry + 23
    x = 0
    rx = 0
    End If
Next i
End Function

Private Sub Options_Click(Index As Integer)
Select Case Index
Case 0
If Options(0).Value = 1 Then
Options(1).Enabled = True
Else
Options(1).Enabled = False
End If
Case 2
If Options(2).Value = 1 Then
Options(3).Enabled = True
Else
Options(3).Enabled = False
End If
Case 4
If Options(4).Value = 1 Then
Options(5).Enabled = True
Else
Options(5).Enabled = False
End If
Case 6
If Options(6).Value = 1 Then
Options(7).Enabled = True
Else
Options(7).Enabled = False
End If
End Select
End Sub
