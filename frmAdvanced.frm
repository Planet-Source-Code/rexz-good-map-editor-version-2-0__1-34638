VERSION 5.00
Begin VB.Form frmAdvanced 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Advanced Properties"
   ClientHeight    =   3315
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   1650
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3315
   ScaleWidth      =   1650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox SelTile 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   0
      Left            =   120
      Picture         =   "frmAdvanced.frx":0000
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   12
      Top             =   600
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
      Picture         =   "frmAdvanced.frx":06BA
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   11
      Top             =   120
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
      Picture         =   "frmAdvanced.frx":0D74
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   10
      Top             =   1080
      Width           =   375
   End
   Begin VB.PictureBox SelTile 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   3
      Left            =   120
      Picture         =   "frmAdvanced.frx":142E
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   9
      Top             =   1560
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
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   0
      Left            =   600
      TabIndex        =   8
      Top             =   120
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
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   600
      TabIndex        =   7
      Top             =   600
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
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   2
      Left            =   600
      TabIndex        =   6
      Top             =   1080
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
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   3
      Left            =   600
      TabIndex        =   5
      Top             =   1560
      Width           =   855
   End
   Begin VB.OptionButton Option1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Make UnWalkable"
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
      Left            =   120
      TabIndex        =   4
      Top             =   2040
      Value           =   -1  'True
      Width           =   1455
   End
   Begin VB.OptionButton Option2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Make Walkable"
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
      Left            =   120
      TabIndex        =   3
      Top             =   2280
      Width           =   1455
   End
   Begin VB.TextBox txtLayer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      MaxLength       =   1
      TabIndex        =   2
      Text            =   "1"
      Top             =   2760
      Width           =   1455
   End
   Begin VB.CommandButton cmdExecute 
      Caption         =   "Execute"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   3000
      Width           =   1455
   End
   Begin VB.OptionButton Option3 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Change Layer to"
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
      Left            =   120
      TabIndex        =   0
      Top             =   2520
      Width           =   1455
   End
End
Attribute VB_Name = "frmAdvanced"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Function Execute()
Dim i
If Option1.Value = True Then

    For i = 0 To 287
    If ThisMap.Tiles(i).FXType = 5 Then
    Else
    If Options(ThisMap.Tiles(i).FXType).Value = 1 Then
    ThisMap.Tiles(i).Walkable = 0
    Else
    End If
    End If
    Next i
    
ElseIf Option2.Value = True Then

    For i = 0 To 287
    If ThisMap.Tiles(i).FXType = 5 Then
    Else
    If Options(ThisMap.Tiles(i).FXType).Value = 1 Then
    ThisMap.Tiles(i).Walkable = 1
    Else
    End If
    End If
    Next i
    
ElseIf Option3.Value = True Then

    For i = 0 To 287
    If ThisMap.Tiles(i).FXType = 5 Then
    Else
    If Options(ThisMap.Tiles(i).FXType).Value = 1 Then
    ThisMap.Tiles(i).Layer = CInt(txtLayer.Text)
    Else
    End If
    End If
    Next i

End If
End Function

Private Sub cmdExecute_Click()
Execute
End Sub

Private Sub txtLayer_KeyPress(KeyAscii As Integer)
KeyAscii = IIf(Not KeyAscii = 8 And Not Val((Chr(KeyAscii))) > 0, 0, KeyAscii)
End Sub

