VERSION 5.00
Begin VB.Form frmEvent 
   BackColor       =   &H80000007&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Add Event"
   ClientHeight    =   3420
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   1545
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3420
   ScaleWidth      =   1545
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton Option1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      Caption         =   "Message Event"
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
      TabIndex        =   6
      Top             =   360
      Value           =   -1  'True
      Width           =   1455
   End
   Begin VB.TextBox txtMsg 
      Appearance      =   0  'Flat
      Height          =   735
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Text            =   "frmEvent.frx":0000
      Top             =   600
      Width           =   1335
   End
   Begin VB.OptionButton Option2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      Caption         =   "Warp Event"
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
      Top             =   1440
      Width           =   1455
   End
   Begin VB.TextBox txtWarp 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   495
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Text            =   "frmEvent.frx":0016
      Top             =   1680
      Width           =   1335
   End
   Begin VB.OptionButton Option3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      Caption         =   "Damage Event"
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
      TabIndex        =   2
      Top             =   2280
      Width           =   1455
   End
   Begin VB.TextBox txtDamage 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   495
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "frmEvent.frx":002B
      Top             =   2520
      Width           =   1335
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add Event"
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
      TabIndex        =   0
      Top             =   3120
      Width           =   1335
   End
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
      ForeColor       =   &H00FFFFFF&
      Height          =   165
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   345
   End
End
Attribute VB_Name = "frmEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public iIndex As Integer
Private Sub cmdAdd_Click()
If Option1.Value = True Then
ThisMap.Tiles(iIndex).Event = "MSG=" & txtMsg.Text
ElseIf Option2.Value = True Then
ThisMap.Tiles(iIndex).Event = "WARP=" & txtWarp.Text
ElseIf Option3.Value = True Then
ThisMap.Tiles(iIndex).Event = "DAMAGE=" & txtDamage.Text
End If
Unload Me
End Sub

Private Sub Option1_Click()
txtMsg.Enabled = True
txtWarp.Enabled = False
txtDamage.Enabled = False
End Sub

Private Sub Option2_Click()
txtWarp.Enabled = True
txtMsg.Enabled = False
txtDamage.Enabled = False
End Sub

Private Sub Option3_Click()
txtDamage.Enabled = True
txtWarp.Enabled = False
txtMsg.Enabled = False
End Sub
