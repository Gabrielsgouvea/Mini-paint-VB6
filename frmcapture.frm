VERSION 5.00
Begin VB.Form frmcapture 
   BackColor       =   &H00C0C0C0&
   Caption         =   "capturar cor"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmr 
      Left            =   1350
      Top             =   690
   End
End
Attribute VB_Name = "frmcapture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rgbvalue As Long
Dim pt As POINTAPI
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetPixel Lib "gdi32.dll" (ByVal hdc As Long, ByVal nXPos As Long, ByVal nYPos As Long) As Long
Private Declare Function CreateDC Lib "gdi32.dll" Alias "CreateDCA" (ByVal lpszDriver As String, ByVal lpszDevice As String, ByVal lpszOutput As Long, lpInitData As Any) As Long
Private Declare Function GetDC Lib "user32.dll" (ByVal hWnd As Long) As Long

'codigogos de captura de cor
Private Sub tmr_Timer()
GetCursorPos pt
rgbvalue = GetPixel(GetDC(DefaultMonitor), pt.X, pt.Y)
Me.BackColor = rgbvalue
HexColor = Hex(rgbvalue)
Me.Caption = Str(rgbvalue)
End Sub
