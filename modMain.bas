Attribute VB_Name = "modMain"
Option Explicit
Declare Function BitBlt Lib "gdi32" ( _
        ByVal hDestDC As Long, _
        ByVal x As Long, _
        ByVal y As Long, _
        ByVal nWidth As Long, _
        ByVal nHeight As Long, _
        ByVal hSrcDC As Long, _
        ByVal xSrc As Long, _
        ByVal ySrc As Long, _
        ByVal dwRop As Long) As Long

Public Const SRCCOPY = &HCC0020
Public Const SRCAND = &H8800C6
Public Const SRCINVERT = &H660046
Public Const SRCPAINT = &HEE0086
Public Const SRCERASE = &H4400328
Public Const WHITENESS = &HFF0062
Public Const BLACKNESS = &H42

Private Type FXTile
    Walkable As Integer
    FXType As Integer
    Layer As Integer
    Event As String
End Type

Public Type Map
    Tiles(287) As FXTile
    sname As String
End Type

Public ThisMap As Map

Function FindPart(lzStr As String, mPart As String) As Integer
Dim TPos As Integer
    TPos = InStr(lzStr, mPart)
    If TPos Then
        FindPart = 1
    Else
        FindPart = 0
    End If
    
End Function
