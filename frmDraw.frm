VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmDraw 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "font"
   ClientHeight    =   8340
   ClientLeft      =   150
   ClientTop       =   480
   ClientWidth     =   13380
   BeginProperty Font 
      Name            =   "Arial Black"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDraw.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8340
   ScaleWidth      =   13380
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picmenu 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   405
      Left            =   -75
      ScaleHeight     =   405
      ScaleWidth      =   13545
      TabIndex        =   36
      Top             =   -30
      Width           =   13545
      Begin VB.Label lblArq 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         BackStyle       =   0  'Transparent
         Caption         =   "Sair"
         Height          =   375
         Index           =   5
         Left            =   4290
         TabIndex        =   66
         Top             =   30
         Width           =   825
      End
      Begin VB.Label lblArq 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         BackStyle       =   0  'Transparent
         Caption         =   "Texto"
         Height          =   375
         Index           =   4
         Left            =   3465
         TabIndex        =   48
         Top             =   30
         Width           =   825
      End
      Begin VB.Label lblArq 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         BackStyle       =   0  'Transparent
         Caption         =   "Novo"
         Height          =   375
         Index           =   0
         Left            =   60
         TabIndex        =   40
         Top             =   30
         Width           =   825
      End
      Begin VB.Label lblArq 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         BackStyle       =   0  'Transparent
         Caption         =   "Salvar"
         Height          =   375
         Index           =   1
         Left            =   885
         TabIndex        =   39
         Top             =   30
         Width           =   825
      End
      Begin VB.Label lblArq 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         BackStyle       =   0  'Transparent
         Caption         =   "Abrir"
         Height          =   375
         Index           =   2
         Left            =   1710
         TabIndex        =   38
         Top             =   30
         Width           =   825
      End
      Begin VB.Label lblArq 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         BackStyle       =   0  'Transparent
         Caption         =   "Imprimir"
         Height          =   375
         Index           =   3
         Left            =   2535
         TabIndex        =   37
         Top             =   30
         Width           =   930
      End
   End
   Begin VB.PictureBox picj 
      AutoRedraw      =   -1  'True
      Height          =   510
      Index           =   6
      Left            =   9270
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   35
      Top             =   480
      Width           =   510
   End
   Begin VB.PictureBox picj 
      AutoRedraw      =   -1  'True
      Height          =   510
      Index           =   5
      Left            =   8670
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   34
      Top             =   480
      Width           =   510
   End
   Begin VB.PictureBox picj 
      AutoRedraw      =   -1  'True
      Height          =   510
      Index           =   4
      Left            =   8055
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   33
      Top             =   480
      Width           =   510
   End
   Begin VB.PictureBox picj 
      AutoRedraw      =   -1  'True
      Height          =   510
      Index           =   3
      Left            =   7440
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   32
      Top             =   480
      Width           =   510
   End
   Begin VB.PictureBox picj 
      AutoRedraw      =   -1  'True
      Height          =   510
      Index           =   2
      Left            =   6780
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   31
      Top             =   480
      Width           =   510
   End
   Begin VB.PictureBox picj 
      AutoRedraw      =   -1  'True
      Height          =   510
      Index           =   1
      Left            =   6120
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   30
      Top             =   480
      Width           =   510
   End
   Begin VB.PictureBox picj 
      AutoRedraw      =   -1  'True
      Height          =   510
      Index           =   0
      Left            =   5490
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   29
      Top             =   480
      Width           =   510
   End
   Begin VB.TextBox txtr 
      Alignment       =   2  'Center
      Height          =   390
      Left            =   285
      TabIndex        =   23
      Text            =   "0"
      Top             =   7665
      Visible         =   0   'False
      Width           =   1005
   End
   Begin MSComDlg.CommonDialog cdlsave 
      Left            =   12930
      Top             =   3225
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "Bitmaps (*.bmp)|*.bmp|"
   End
   Begin VB.PictureBox picQuadro 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6375
      Left            =   1995
      ScaleHeight     =   6315
      ScaleWidth      =   11085
      TabIndex        =   0
      Top             =   1755
      Width           =   11145
      Begin VB.PictureBox pictexto 
         BackColor       =   &H00808080&
         Height          =   3900
         Left            =   2850
         ScaleHeight     =   3840
         ScaleWidth      =   5535
         TabIndex        =   61
         Top             =   1125
         Visible         =   0   'False
         Width           =   5595
         Begin VB.TextBox txtsize 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   2625
            TabIndex        =   69
            Text            =   "10"
            Top             =   1665
            Width           =   1215
         End
         Begin VB.FileListBox filfont 
            Height          =   1980
            Left            =   150
            TabIndex        =   67
            Top             =   1650
            Width           =   2070
         End
         Begin VB.PictureBox picok 
            AutoRedraw      =   -1  'True
            BorderStyle     =   0  'None
            Height          =   645
            Left            =   3780
            ScaleHeight     =   645
            ScaleWidth      =   1590
            TabIndex        =   64
            Top             =   3060
            Width           =   1590
            Begin VB.Label lblok 
               Alignment       =   2  'Center
               BackColor       =   &H000000FF&
               BackStyle       =   0  'Transparent
               Caption         =   "OK"
               BeginProperty Font 
                  Name            =   "Arial Black"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   465
               Left            =   90
               TabIndex        =   65
               Top             =   90
               Width           =   1410
            End
            Begin VB.Shape Shape24 
               BorderColor     =   &H00808080&
               BorderWidth     =   3
               Height          =   525
               Left            =   45
               Shape           =   4  'Rounded Rectangle
               Top             =   60
               Width           =   1485
            End
            Begin VB.Shape Shape23 
               BorderColor     =   &H00808080&
               BorderWidth     =   5
               Height          =   585
               Left            =   30
               Top             =   30
               Width           =   1530
            End
         End
         Begin VB.TextBox txttexto 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   1260
            TabIndex        =   62
            Top             =   420
            Width           =   2925
         End
         Begin VB.Shape shpcortext 
            BorderWidth     =   3
            Height          =   315
            Left            =   3015
            Shape           =   4  'Rounded Rectangle
            Top             =   2595
            Width           =   495
         End
         Begin VB.Label lblcortext 
            BackColor       =   &H00000000&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3045
            TabIndex        =   74
            Top             =   2580
            Width           =   450
         End
         Begin VB.Label Label16 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "cor"
            Height          =   300
            Left            =   2745
            TabIndex        =   73
            Top             =   2235
            Width           =   1005
         End
         Begin VB.Label Label15 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Texto"
            Height          =   300
            Left            =   2310
            TabIndex        =   72
            Top             =   0
            Width           =   1005
         End
         Begin VB.Label Label14 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Tamanho"
            Height          =   300
            Left            =   2715
            TabIndex        =   71
            Top             =   1245
            Width           =   1005
         End
         Begin VB.Shape Shape26 
            BorderWidth     =   3
            Height          =   525
            Left            =   2520
            Shape           =   4  'Rounded Rectangle
            Top             =   1605
            Width           =   1425
         End
         Begin VB.Label Label13 
            BackColor       =   &H00FFFFFF&
            Height          =   495
            Left            =   2550
            TabIndex        =   70
            Top             =   1620
            Width           =   1365
         End
         Begin VB.Label Label12 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Fontes"
            Height          =   300
            Left            =   720
            TabIndex        =   68
            Top             =   1290
            Width           =   870
         End
         Begin VB.Shape Shape25 
            Height          =   2595
            Left            =   120
            Shape           =   4  'Rounded Rectangle
            Top             =   1200
            Width           =   2145
         End
         Begin VB.Shape Shape22 
            BorderWidth     =   3
            Height          =   705
            Left            =   1155
            Shape           =   4  'Rounded Rectangle
            Top             =   360
            Width           =   3135
         End
         Begin VB.Label Label11 
            BackColor       =   &H00FFFFFF&
            Height          =   675
            Left            =   1185
            TabIndex        =   63
            Top             =   375
            Width           =   3075
         End
      End
      Begin VB.Line linr 
         Visible         =   0   'False
         X1              =   -2850
         X2              =   -2850
         Y1              =   1560
         Y2              =   2985
      End
      Begin VB.Line linl 
         Visible         =   0   'False
         X1              =   -3945
         X2              =   -3945
         Y1              =   1320
         Y2              =   2445
      End
      Begin VB.Line shpb 
         Visible         =   0   'False
         X1              =   -3735
         X2              =   -2025
         Y1              =   525
         Y2              =   525
      End
      Begin VB.Line shpt 
         Visible         =   0   'False
         X1              =   -3855
         X2              =   -2145
         Y1              =   855
         Y2              =   855
      End
      Begin VB.Line linquadro 
         Visible         =   0   'False
         X1              =   -1890
         X2              =   -180
         Y1              =   1290
         Y2              =   2580
      End
      Begin VB.Shape shpcircle 
         Height          =   555
         Left            =   -1000
         Shape           =   3  'Circle
         Top             =   885
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.Label lblborracha 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   600
         Left            =   -1000
         TabIndex        =   41
         Top             =   1590
         Width           =   600
      End
   End
   Begin VB.Label lblinvisible 
      BackStyle       =   0  'Transparent
      Height          =   480
      Left            =   2010
      TabIndex        =   60
      Top             =   1200
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.Label lblfundo 
      BackColor       =   &H00FF00FF&
      DragIcon        =   "frmDraw.frx":0ECA
      DragMode        =   1  'Automatic
      Height          =   480
      Index           =   7
      Left            =   5385
      TabIndex        =   59
      Top             =   1200
      Width           =   480
   End
   Begin VB.Label lblfundo 
      BackColor       =   &H00FFFF00&
      DragIcon        =   "frmDraw.frx":1794
      DragMode        =   1  'Automatic
      Height          =   480
      Index           =   6
      Left            =   4905
      TabIndex        =   58
      Top             =   1200
      Width           =   480
   End
   Begin VB.Label lblfundo 
      BackColor       =   &H0000FFFF&
      DragIcon        =   "frmDraw.frx":205E
      DragMode        =   1  'Automatic
      Height          =   480
      Index           =   5
      Left            =   4425
      TabIndex        =   57
      Top             =   1200
      Width           =   480
   End
   Begin VB.Label lblfundo 
      BackColor       =   &H00FFFFFF&
      DragIcon        =   "frmDraw.frx":2928
      DragMode        =   1  'Automatic
      Height          =   480
      Index           =   4
      Left            =   3945
      TabIndex        =   56
      Top             =   1200
      Width           =   480
   End
   Begin VB.Label lblfundo 
      BackColor       =   &H00000000&
      DragIcon        =   "frmDraw.frx":31F2
      DragMode        =   1  'Automatic
      Height          =   480
      Index           =   3
      Left            =   3465
      TabIndex        =   55
      Top             =   1200
      Width           =   480
   End
   Begin VB.Label lblfundo 
      BackColor       =   &H0000FF00&
      DragIcon        =   "frmDraw.frx":3ABC
      DragMode        =   1  'Automatic
      Height          =   480
      Index           =   2
      Left            =   2985
      TabIndex        =   54
      Top             =   1200
      Width           =   480
   End
   Begin VB.Label lblfundo 
      BackColor       =   &H00FF0000&
      DragIcon        =   "frmDraw.frx":4386
      DragMode        =   1  'Automatic
      Height          =   480
      Index           =   1
      Left            =   2505
      TabIndex        =   53
      Top             =   1200
      Width           =   480
   End
   Begin VB.Label lblfundo 
      BackColor       =   &H000000FF&
      DragIcon        =   "frmDraw.frx":4C50
      DragMode        =   1  'Automatic
      Height          =   480
      Index           =   0
      Left            =   2025
      TabIndex        =   52
      Top             =   1200
      Width           =   480
   End
   Begin VB.Label lbllápis 
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   285
      TabIndex        =   51
      Top             =   5160
      Width           =   1035
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   240
      TabIndex        =   50
      Top             =   5175
      Width           =   1170
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   615
      Picture         =   "frmDraw.frx":551A
      Top             =   5220
      Width           =   480
   End
   Begin VB.Line linf 
      BorderWidth     =   2
      Visible         =   0   'False
      X1              =   12405
      X2              =   13005
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label lbllin 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   330
      TabIndex        =   49
      Top             =   6645
      Width           =   1005
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   420
      X2              =   1230
      Y1              =   6780
      Y2              =   6780
   End
   Begin VB.Line lin 
      X1              =   8865
      X2              =   9420
      Y1              =   1365
      Y2              =   1365
   End
   Begin VB.Shape Shape21 
      Height          =   255
      Left            =   8235
      Shape           =   3  'Circle
      Top             =   1215
      Width           =   375
   End
   Begin VB.Label lblmenos 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      Height          =   285
      Left            =   8220
      TabIndex        =   46
      Top             =   1185
      Width           =   390
   End
   Begin VB.Label lblmais 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "+"
      Height          =   285
      Left            =   9795
      TabIndex        =   45
      Top             =   1200
      Width           =   390
   End
   Begin VB.Shape Shape20 
      Height          =   255
      Left            =   9795
      Shape           =   3  'Circle
      Top             =   1215
      Width           =   375
   End
   Begin VB.Label lblcorb 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   12435
      TabIndex        =   43
      Top             =   1230
      Width           =   450
   End
   Begin VB.Shape shpcorb 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      FillColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   12405
      Shape           =   4  'Rounded Rectangle
      Top             =   1230
      Width           =   495
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ferramentas"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   105
      TabIndex        =   42
      Top             =   4725
      Width           =   1440
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "  Janela"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   4485
      TabIndex        =   28
      Top             =   435
      Width           =   5415
   End
   Begin VB.Shape shpf 
      BackColor       =   &H00000000&
      Height          =   300
      Left            =   12390
      Top             =   555
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.Image imgf 
      Height          =   480
      Left            =   12510
      Picture         =   "frmDraw.frx":5DE4
      Top             =   480
      Width           =   480
   End
   Begin VB.Shape shpa 
      BackColor       =   &H00000000&
      BorderColor     =   &H00000000&
      BorderWidth     =   3
      FillColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   3645
      Shape           =   4  'Rounded Rectangle
      Top             =   570
      Width           =   495
   End
   Begin VB.Label lblr 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Raio"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   330
      TabIndex        =   24
      Top             =   7335
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Caixa de cores"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   60
      TabIndex        =   22
      Top             =   450
      Width           =   1545
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   1650
      X2              =   1650
      Y1              =   390
      Y2              =   8340
   End
   Begin VB.Label lblcirculo 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   495
      TabIndex        =   19
      Top             =   6885
      Width           =   600
   End
   Begin VB.Shape Shape19 
      BackColor       =   &H00000000&
      Height          =   420
      Left            =   300
      Shape           =   3  'Circle
      Top             =   6945
      Width           =   1005
   End
   Begin VB.Label lblquadrado 
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   300
      TabIndex        =   17
      Top             =   5685
      Width           =   1035
   End
   Begin VB.Shape Shape17 
      Height          =   420
      Left            =   315
      Top             =   5745
      Width           =   1005
   End
   Begin VB.Shape Shape16 
      BorderColor     =   &H00000080&
      BorderWidth     =   3
      Height          =   315
      Left            =   195
      Shape           =   4  'Rounded Rectangle
      Top             =   2955
      Width           =   495
   End
   Begin VB.Shape Shape15 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      Height          =   315
      Left            =   915
      Shape           =   4  'Rounded Rectangle
      Top             =   2955
      Width           =   495
   End
   Begin VB.Shape Shape14 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   3
      Height          =   315
      Left            =   900
      Shape           =   4  'Rounded Rectangle
      Top             =   3390
      Width           =   495
   End
   Begin VB.Shape Shape13 
      BorderColor     =   &H00800080&
      BorderWidth     =   3
      Height          =   315
      Left            =   180
      Shape           =   4  'Rounded Rectangle
      Top             =   3390
      Width           =   495
   End
   Begin VB.Shape Shape12 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   3
      Height          =   315
      Left            =   180
      Shape           =   4  'Rounded Rectangle
      Top             =   4275
      Width           =   495
   End
   Begin VB.Shape Shape11 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      FillColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   900
      Shape           =   4  'Rounded Rectangle
      Top             =   4275
      Width           =   495
   End
   Begin VB.Shape Shape10 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   3
      Height          =   315
      Left            =   915
      Shape           =   4  'Rounded Rectangle
      Top             =   3825
      Width           =   495
   End
   Begin VB.Shape Shape9 
      BorderColor     =   &H00008080&
      BorderWidth     =   3
      Height          =   315
      Left            =   195
      Shape           =   4  'Rounded Rectangle
      Top             =   3825
      Width           =   495
   End
   Begin VB.Shape Shape8 
      BorderColor     =   &H00008000&
      BorderWidth     =   3
      Height          =   315
      Left            =   195
      Shape           =   4  'Rounded Rectangle
      Top             =   2085
      Width           =   495
   End
   Begin VB.Shape Shape7 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   3
      Height          =   315
      Left            =   915
      Shape           =   4  'Rounded Rectangle
      Top             =   2085
      Width           =   495
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H00FFFF00&
      BorderWidth     =   3
      Height          =   315
      Left            =   900
      Shape           =   4  'Rounded Rectangle
      Top             =   2520
      Width           =   495
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H00808000&
      BorderWidth     =   3
      Height          =   315
      Left            =   180
      Shape           =   4  'Rounded Rectangle
      Top             =   2520
      Width           =   495
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00800000&
      BorderWidth     =   3
      Height          =   315
      Left            =   180
      Shape           =   4  'Rounded Rectangle
      Top             =   1650
      Width           =   495
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      Height          =   315
      Left            =   900
      Shape           =   4  'Rounded Rectangle
      Top             =   1650
      Width           =   495
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00808080&
      BorderWidth     =   3
      Height          =   315
      Left            =   915
      Shape           =   4  'Rounded Rectangle
      Top             =   1200
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   3
      Height          =   315
      Left            =   195
      Shape           =   4  'Rounded Rectangle
      Top             =   1200
      Width           =   495
   End
   Begin VB.Label lblCor 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   15
      Left            =   930
      TabIndex        =   16
      Top             =   4275
      Width           =   450
   End
   Begin VB.Label lblCor 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   14
      Left            =   930
      TabIndex        =   15
      Top             =   3825
      Width           =   450
   End
   Begin VB.Label lblCor 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   13
      Left            =   930
      TabIndex        =   14
      Top             =   3390
      Width           =   450
   End
   Begin VB.Label lblCor 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   12
      Left            =   930
      TabIndex        =   13
      Top             =   2955
      Width           =   450
   End
   Begin VB.Label lblCor 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   11
      Left            =   930
      TabIndex        =   12
      Top             =   2520
      Width           =   450
   End
   Begin VB.Label lblCor 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   10
      Left            =   930
      TabIndex        =   11
      Top             =   2085
      Width           =   450
   End
   Begin VB.Label lblCor 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   9
      Left            =   930
      TabIndex        =   10
      Top             =   1650
      Width           =   450
   End
   Begin VB.Label lblCor 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   8
      Left            =   930
      TabIndex        =   9
      Top             =   1200
      Width           =   450
   End
   Begin VB.Label lblCor 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   7
      Left            =   210
      TabIndex        =   8
      Top             =   4275
      Width           =   450
   End
   Begin VB.Label lblCor 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   6
      Left            =   210
      TabIndex        =   7
      Top             =   3825
      Width           =   450
   End
   Begin VB.Label lblCor 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   5
      Left            =   210
      TabIndex        =   6
      Top             =   3390
      Width           =   450
   End
   Begin VB.Label lblCor 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   4
      Left            =   210
      TabIndex        =   5
      Top             =   2955
      Width           =   450
   End
   Begin VB.Label lblCor 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   3
      Left            =   210
      TabIndex        =   4
      Top             =   2520
      Width           =   450
   End
   Begin VB.Label lblCor 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   2
      Left            =   210
      TabIndex        =   3
      Top             =   2085
      Width           =   450
   End
   Begin VB.Label lblCor 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   210
      TabIndex        =   2
      Top             =   1650
      Width           =   450
   End
   Begin VB.Label lblCor 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   210
      TabIndex        =   1
      Top             =   1200
      Width           =   450
   End
   Begin VB.Label lblquadradof 
      BackColor       =   &H00FF0000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Left            =   270
      TabIndex        =   18
      Top             =   6180
      Width           =   1080
   End
   Begin VB.Shape Shape18 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   420
      Left            =   315
      Top             =   6195
      Width           =   1005
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   3075
      Left            =   225
      TabIndex        =   20
      Top             =   5160
      Width           =   1170
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   3600
      Left            =   105
      TabIndex        =   21
      Top             =   1080
      Width           =   1410
   End
   Begin VB.Label lbla 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3675
      TabIndex        =   26
      Top             =   570
      Width           =   450
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "  Cor atual"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   2235
      TabIndex        =   25
      Top             =   495
      Width           =   2205
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Ferramenta atual"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   10080
      TabIndex        =   27
      Top             =   495
      Width           =   3015
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "  Cor da borracha"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   10485
      TabIndex        =   44
      Top             =   1155
      Width           =   2565
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "  Espessura da linha"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   6030
      TabIndex        =   47
      Top             =   1155
      Width           =   4230
   End
End
Attribute VB_Name = "frmDraw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim desenhando As Boolean
Dim alt As Integer ' pega a autura do botão
Dim larg As Integer ' pega a largura do botão
Dim cor As Variant ' guarda a cor do botão
Dim step As Variant ' o passo para saber quantos retangulos vão desenhar
Dim L As Integer 'usado  no loop
Dim quadrado As Byte 'utiliza para controle de desanhar quadrado
Dim qx As Integer ' usado para guarda x para desenhar retangulo
Dim qy As Integer ' usado para guarda y para desenhar retangulo
Dim corborracha As Variant ' corda borracha
Dim indfundo As Byte ' cor do fundo usando dragdrop
Dim n As Byte 'loop
Dim fonte As String 'guarda a fonte
Private Sub filfont_Click()
'mostra as fontes na pasta font do windows
fonte = CStr(filfont.FileName)
fonte = Mid(fonte, 1, Len(fonte) - 4)
txttexto.Font = fonte
picQuadro.Font = fonte
End Sub

Private Sub Form_Load()
'carrega as cores em cada label da array
Dim I As Integer
For I = 0 To 15
lblCor(I).BackColor = QBColor(I)
Next
' faz efeito 3d no form
cor = 255
larg = Me.ScaleWidth
alt = Me.ScaleHeight
step = 255 / alt
For L = 0 To alt
Me.Line (0, L)-(larg, L), RGB(cor, cor, cor), BF
cor = cor - step
Next

' faz efeito 3d no "botão" ok
cor = 255
larg = picok.ScaleWidth
alt = picok.ScaleHeight
step = 200 / alt
For L = 0 To alt
picok.Line (0, L)-(larg, L), RGB(cor, cor, cor), BF
cor = cor - step
Next

'para mudar o ponteiro
Me.MousePointer = vbCustom

' faz efeito no menu
cor = 255
larg = picmenu.ScaleWidth
alt = picmenu.ScaleHeight
step = 255 / alt
For L = 0 To alt
picmenu.Line (0, L)-(larg, L), RGB(cor, cor, cor), BF
cor = cor - step
Next

'cores para a janela
'pictur0
cor = 255
larg = picj(0).ScaleWidth
alt = picj(0).ScaleHeight
step = 255 / alt
For L = 0 To alt
picj(0).Line (0, L)-(larg, L), RGB(cor, cor, cor), BF
cor = cor - step
Next
'pictur1
cor = 255
larg = picj(1).ScaleWidth
alt = picj(1).ScaleHeight
step = 255 / alt
For L = 0 To alt
picj(1).Line (0, L)-(larg, L), RGB(255, cor, cor), BF
cor = cor - step
Next
'pictur2
cor = 255
larg = picj(2).ScaleWidth
alt = picj(2).ScaleHeight
step = 255 / alt
For L = 0 To alt
picj(2).Line (0, L)-(larg, L), RGB(cor, 255, cor), BF
cor = cor - step
Next
'pictur3
cor = 255
larg = picj(3).ScaleWidth
alt = picj(3).ScaleHeight
step = 255 / alt
For L = 0 To alt
picj(3).Line (0, L)-(larg, L), RGB(cor, cor, 255), BF
cor = cor - step
Next
'pictur4
cor = 255
larg = picj(4).ScaleWidth
alt = picj(4).ScaleHeight
step = 255 / alt
For L = 0 To alt
picj(4).Line (0, L)-(larg, L), RGB(cor, cor, 0), BF
cor = cor - step
Next
'pictur5
cor = 255
larg = picj(5).ScaleWidth
alt = picj(5).ScaleHeight
step = 255 / alt
For L = 0 To alt
picj(5).Line (0, L)-(larg, L), RGB(0, cor, cor), BF
cor = cor - step
Next
'pictur6
cor = 255
larg = picj(6).ScaleWidth
alt = picj(6).ScaleHeight
step = 255 / alt
For L = 0 To alt
picj(6).Line (0, L)-(larg, L), RGB(cor, 0, cor), BF
cor = cor - step
Next
corborracha = RGB(255, 255, 255)
frmcapture.Refresh
'exibe os fontes
filfont.Path = ("c:\windows\fonts")
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim n As Byte
'valta o ponteiroao normal
Me.MouseIcon = LoadPicture()
'desfaz os "selecionado" dos menus
For n = 0 To 5
     lblArq(n).BackStyle = 0
Next
End Sub
Private Sub Form_Unload(Cancel As Integer)
'para fechaar o form de captura de cor
Unload frmcapture
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'desfaz os "selecionado" dos menus
Dim n As Byte
For n = 0 To 5
     lblArq(n).BackStyle = 0
Next

End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'desfaz os "selecionado" dos menus
Dim n As Byte
For n = 0 To 5
     lblArq(n).BackStyle = 0
Next
End Sub


Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'desfaz os "selecionado" dos menus
Dim n As Byte
For n = 0 To 5
     lblArq(n).BackStyle = 0
Next
End Sub

Private Sub Label6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'desfaz os "selecionado" dos menus
Dim n As Byte
For n = 0 To 5
     lblArq(n).BackStyle = 0
Next
End Sub

Private Sub lblArq_Click(Index As Integer)
Dim Response As Integer 'usada para capturar a resposta da textbox
Select Case Index
    Case 0
        ' certifica-se que o utilizador quer realmente recomeçar
        Beep
        Response = MsgBox("Tem certeza que quer iniciar novo desenho?", vbYesNo + vbQuestion, "Novo Desenho")
        If Response = vbYes Then
            picQuadro.Cls
            picQuadro.Picture = LoadPicture()
        End If
    Case 1
        'istruções para salvar um desenho
        cdlsave.Filter = "Files (*.bmp)|*.bmp"
        cdlsave.DefaultExt = "bmp"
        cdlsave.DialogTitle = "Salvar Arquivo"
        cdlsave.Flags = cdlOFNFileMustExist + cdlOFNPathMustExist
        cdlsave.ShowSave
        'previne erro caso alguem desista de salvar
        On Error GoTo fim
        SavePicture picQuadro.Image, cdlsave.InitDir + cdlsave.FileName
fim:
    Case 2
        'istruções para abrir um desenho ou imagem
        cdlsave.Filter = "Files (*.bmp)|*.bmp"
        cdlsave.DefaultExt = "bmp"
        cdlsave.DialogTitle = "Abrir Arquivo"
        cdlsave.Flags = cdlOFNFileMustExist + cdlOFNPathMustExist
        cdlsave.ShowOpen
        picQuadro.Picture = LoadPicture(cdlsave.InitDir + cdlsave.FileName)
    Case 3
        Beep
        Response = MsgBox("Tem certeza que quer imprimir", vbYesNo + vbQuestion, "Imprimir")
        If Response = vbYes Then
            cdlsave.Flags = cdlOFNFileMustExist + cdlOFNPathMustExist
            frmprint.picprint.Picture = picQuadro.Image
            cdlsave.ShowPrinter
            frmprint.PrintForm
        End If
    Case 4
        pictexto.Visible = True
        quadrado = 5
    Case 5
        ' certifica-se se o utilizador realmente quer sair
        Response = MsgBox("Tem certeza que quer sair do programa?", vbYesNo + vbCritical + vbDefaultButton2, "Sair do Quadro Mágico")
        If Response = vbYes Then End
End Select
End Sub

Private Sub lblArq_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim n As Integer
For n = 0 To 5
     lblArq(n).BackStyle = 0
Next
lblArq(Index).BackStyle = 1
lblArq(Index).BackColor = RGB(51, 153, 255)

If X <= 0 Or Y <= 0 Then
    For n = 0 To 5
        lblArq(n).BackStyle = 0
    Next
End If
End Sub

Private Sub lblcirculo_Click()
'utilizaado para desenhar um circulo
quadrado = 3
lblr.Visible = True
txtr.Visible = True
imgf.Visible = False
linf.Visible = False
shpf.Visible = True
shpf.Shape = 3
shpf.BackStyle = 0
shpcircle.Visible = True
picQuadro.ForeColor = lbla.BackColor
End Sub

Private Sub lblCor_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
' troca para a cor escolhida
If Button = 1 Then
    picQuadro.ForeColor = lblCor(Index).BackColor
    shpa.BorderColor = lblCor(Index).BackColor
    lbla.BackColor = lblCor(Index).BackColor
End If
If Button = 2 Then
    corborracha = lblCor(Index).BackColor
    lblborracha.BackColor = lblCor(Index).BackColor
    lblcorb.BackColor = lblCor(Index).BackColor
    shpcorb.BorderColor = lblCor(Index).BackColor
End If
End Sub

Private Sub lblcorb_Click()
'opçõe avançadas para cor da borracha
cdlsave.ShowColor
lblcorb.BackColor = cdlsave.Color
shpcorb.BorderColor = cdlsave.Color
lblborracha.BackColor = cdlsave.Color
corborracha = cdlsave.Color
End Sub

Private Sub lblcortext_Click()
'muda a cor da fonte
cdlsave.ShowColor
picQuadro.ForeColor = cdlsave.Color
lblcortext.BackColor = cdlsave.Color
shpcortext.BorderColor = cdlsave.Color
txttexto.ForeColor = cdlsave.Color
End Sub

Private Sub lblfundo_DragOver(Index As Integer, Source As Control, X As Single, Y As Single, State As Integer)
'captura o index para mudar o fundo
Select Case Index
    Case 0
        indfundo = 0
    Case 1
        indfundo = 1
    Case 2
        indfundo = 2
    Case 3
        indfundo = 3
    Case 4
        indfundo = 4
    Case 5
        indfundo = 5
    Case 6
        indfundo = 6
    Case 7
        indfundo = 7
End Select
lblinvisible.Visible = True
End Sub

Private Sub lblinvisible_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
For n = 0 To 7
    lblfundo(n).Enabled = True
Next
lblinvisible.Visible = False
End Sub

Private Sub lbllápis_Click()
quadrado = 0
imgf.Visible = True
shpf.Visible = False
linf.Visible = False
lblr.Visible = False
txtr.Visible = False
shpcircle.Visible = False
shpcircle.Left = -1000
picQuadro.ForeColor = lbla.BackColor
End Sub

Private Sub lbllin_Click()
quadrado = 4
imgf.Visible = False
shpf.Visible = False
linf.Visible = True
lblr.Visible = False
txtr.Visible = False
shpcircle.Visible = False
shpcircle.Left = -1000
picQuadro.ForeColor = lbla.BackColor
End Sub

Private Sub lblmais_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    If picQuadro.DrawWidth < 30 Then
        picQuadro.DrawWidth = picQuadro.DrawWidth + 1
        lin.BorderWidth = picQuadro.DrawWidth
    End If
End If
If Button = 2 Then
    picQuadro.DrawWidth = 30
    lin.BorderWidth = picQuadro.DrawWidth
End If
End Sub

Private Sub lblmenos_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    If picQuadro.DrawWidth > 1 Then
        picQuadro.DrawWidth = picQuadro.DrawWidth - 1
        lin.BorderWidth = picQuadro.DrawWidth
    End If
End If
If Button = 2 Then
    picQuadro.DrawWidth = 1
    lin.BorderWidth = picQuadro.DrawWidth
End If
End Sub

Private Sub lblok_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
picok.BackColor = RGB(180, 180, 180)
If txtsize.Text = "" Then txtsize.Text = 10
picQuadro.FontSize = txtsize.Text
picQuadro.ForeColor = lblcortext.BackColor
End Sub

Private Sub lblok_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
' faz efeito 3d no "botão" ok
cor = 255
larg = picok.ScaleWidth
alt = picok.ScaleHeight
step = 200 / alt
For L = 0 To alt
picok.Line (0, L)-(larg, L), RGB(cor, cor, cor), BF
cor = cor - step
Next

pictexto.Visible = False
quadrado = 5
End Sub

Private Sub lblquadrado_Click()
'utilizaado para desenhar um quadrado
quadrado = 1
imgf.Visible = False
shpf.Visible = True
linf.Visible = False
shpf.Shape = 0
shpf.BackStyle = 0
lblr.Visible = False
txtr.Visible = False
shpcircle.Visible = False
shpcircle.Left = -1000
picQuadro.ForeColor = lbla.BackColor
End Sub
Private Sub lblquadradof_Click()
'utilizaado para desenhar um quadrado fechado
quadrado = 2
imgf.Visible = False
shpf.Visible = True
linf.Visible = False
shpf.Shape = 0
shpf.BackStyle = 1
lblr.Visible = False
txtr.Visible = False
shpcircle.Visible = False
shpcircle.Left = -1000
picQuadro.ForeColor = lbla.BackColor
End Sub
Private Sub picj_Click(Index As Integer)
Select Case Index
    Case 0
        ' faz efeito 3d nas picture box
        cor = 255
        larg = Me.ScaleWidth
        alt = Me.ScaleHeight
        step = 255 / alt
        For L = 0 To alt
        Me.Line (0, L)-(larg, L), RGB(cor, cor, cor), BF
        cor = cor - step
        Next
    Case 1
        ' faz efeito 3d nas picture box
        cor = 255
        larg = Me.ScaleWidth
        alt = Me.ScaleHeight
        step = 255 / alt
        For L = 0 To alt
        Me.Line (0, L)-(larg, L), RGB(255, cor, cor), BF
        cor = cor - step
        Next
    Case 2
        ' faz efeito 3d nas picture box
        cor = 255
        larg = Me.ScaleWidth
        alt = Me.ScaleHeight
        step = 255 / alt
        For L = 0 To alt
        Me.Line (0, L)-(larg, L), RGB(cor, 255, cor), BF
        cor = cor - step
        Next
    Case 3
        ' faz efeito 3d nas picture box
        cor = 255
        larg = Me.ScaleWidth
        alt = Me.ScaleHeight
        step = 255 / alt
        For L = 0 To alt
        Me.Line (0, L)-(larg, L), RGB(cor, cor, 255), BF
        cor = cor - step
        Next
    Case 4
        ' faz efeito 3d nas picture box
        cor = 255
        larg = Me.ScaleWidth
        alt = Me.ScaleHeight
        step = 255 / alt
        For L = 0 To alt
        Me.Line (0, L)-(larg, L), RGB(cor, cor, 0), BF
        cor = cor - step
        Next
    Case 5
        ' faz efeito 3d nas picture box
        cor = 255
        larg = Me.ScaleWidth
        alt = Me.ScaleHeight
        step = 255 / alt
        For L = 0 To alt
        Me.Line (0, L)-(larg, L), RGB(0, cor, cor), BF
        cor = cor - step
        Next
    Case 6
        ' faz efeito 3d nas picture box
        cor = 255
        larg = Me.ScaleWidth
        alt = Me.ScaleHeight
        step = 255 / alt
        For L = 0 To alt
        Me.Line (0, L)-(larg, L), RGB(cor, 0, cor), BF
        cor = cor - step
        Next
End Select
End Sub

Private Sub picmenu_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'desfaz os "selecionado" dos menus
Dim n As Byte
For n = 0 To 5
     lblArq(n).BackStyle = 0
Next
End Sub

Private Sub picQuadro_DragDrop(Source As Control, X As Single, Y As Single)
'colore o form e a borracha
If indfundo = 0 Then picQuadro.BackColor = vbRed: corborracha = vbRed: lblborracha.BackColor = vbRed: lblcorb.BackColor = vbRed: shpcorb.BorderColor = vbRed
If indfundo = 1 Then picQuadro.BackColor = vbBlue: corborracha = vbBlue: lblborracha.BackColor = vbBlue: lblcorb.BackColor = vbBlue: shpcorb.BorderColor = vbBlue
If indfundo = 2 Then picQuadro.BackColor = vbGreen: corborracha = vbGreen: lblborracha.BackColor = vbGreen: lblcorb.BackColor = vbGreen: shpcorb.BorderColor = vbGreen
If indfundo = 3 Then picQuadro.BackColor = vbBlack: corborracha = vbBlack: lblborracha.BackColor = vbBlack: lblcorb.BackColor = vbBlack: shpcorb.BorderColor = vbBlack
If indfundo = 4 Then picQuadro.BackColor = vbWhite: corborracha = vbWhite: lblborracha.BackColor = vbWhite: lblcorb.BackColor = vbWhite: shpcorb.BorderColor = vbWhite
If indfundo = 5 Then picQuadro.BackColor = vbYellow: corborracha = vbYellow: lblborracha.BackColor = vbYellow: lblcorb.BackColor = vbYellow: shpcorb.BorderColor = vbYellow
If indfundo = 6 Then picQuadro.BackColor = vbCyan: corborracha = vbCyan: lblborracha.BackColor = vbCyan: lblcorb.BackColor = vbCyan: shpcorb.BorderColor = vbCyan
If indfundo = 7 Then picQuadro.BackColor = vbMagenta: corborracha = vbMagenta: lblborracha.BackColor = vbMagenta: lblcorb.BackColor = vbMagenta: shpcorb.BorderColor = vbMagenta
End Sub

Private Sub picQuadro_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'captura cor
If Button = vbMiddleButton Then
    frmcapture.tmr.Interval = 1
    Me.MouseIcon = LoadPicture(App.Path + "/gotas.cur")
    picQuadro.ForeColor = frmcapture.BackColor
    shpa.BorderColor = frmcapture.BackColor
    lbla.BackColor = frmcapture.BackColor
    GoTo fim
End If
' começamos a desenhar
If Button = vbRightButton Then
    Me.MousePointer = 0
    lblborracha.Top = Y - (lblborracha.Height / 2)
    lblborracha.Left = X - (lblborracha.Width / 2)
    picQuadro.DrawWidth = 1
    picQuadro.Line ((X - 300), (Y - 300))-((X + 300), (Y + 300)), corborracha, BF
    GoTo fim
End If
If quadrado >= 1 Then
    qx = X
    qy = Y
    If quadrado = 1 Then shpt.Visible = True: shpb.Visible = True: linl.Visible = True: linr.Visible = True
    If quadrado = 2 Then shpt.Visible = True: shpb.Visible = True: linl.Visible = True: linr.Visible = True
    If quadrado = 4 Then linquadro.Visible = True
    linquadro.X1 = qx
    linquadro.Y1 = qy
    shpb.X1 = qx
    shpt.Y2 = qy
    shpt.Y1 = qy
    shpt.X1 = qx
    linl.X1 = qx
    linl.Y1 = qy
    linl.X2 = qx
    linr.Y1 = qy
    GoTo fim
End If
If Button = vbLeftButton Then
    desenhando = True
    picQuadro.CurrentX = X
    picQuadro.CurrentY = Y
    picQuadro.PSet (X, Y)
End If
fim:
End Sub

Private Sub picQuadro_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
frmcapture.tmr.Interval = 0
'desenha um quadrado
linquadro.Visible = False
shpt.Visible = False
shpb.Visible = False
linl.Visible = False
linr.Visible = False
'captura cor
If Button = vbMiddleButton Then
    GoTo fim
End If
If Button = vbRightButton Then
    Me.MousePointer = vbCustom
    lblborracha.Left = -1000
    picQuadro.DrawWidth = lin.BorderWidth
    GoTo fim
End If
If quadrado = 1 Then
    picQuadro.Line (qx, qy)-(X, Y), picQuadro.ForeColor, B
    quadrado = 1
    GoTo fim
End If
'desenha um quadrado fechado
If quadrado = 2 Then
    picQuadro.Line (qx, qy)-(X, Y), picQuadro.ForeColor, BF
    quadrado = 2
    GoTo fim
End If
'desenha um circulo
If quadrado = 3 Then
    If txtr.Text = "" Then txtr.Text = 0
    picQuadro.Circle (qx, qy), txtr.Text, picQuadro.ForeColor
    quadrado = 3
    GoTo fim
End If
If quadrado = 4 Then
    picQuadro.Line (qx, qy)-(X, Y)
    quadrado = 4
    GoTo fim
End If
If quadrado = 5 Then
    picQuadro.CurrentX = X
    picQuadro.CurrentY = Y
    picQuadro.Print (txttexto.Text)
    GoTo fim
End If
' interrompemos o desenho
If Button = vbLeftButton Then desenhando = False
fim:
End Sub

Private Sub picQuadro_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If txtr.Text = "" Then txtr.Text = 0
Me.MouseIcon = LoadPicture(App.Path + "/lápis.cur")
linquadro.BorderColor = lbla.BackColor
shpcircle.BorderColor = lbla.BackColor

shpt.BorderColor = lbla.BackColor
shpb.BorderColor = lbla.BackColor
linl.BorderColor = lbla.BackColor
linr.BorderColor = lbla.BackColor

shpcircle.BorderWidth = lin.BorderWidth
linquadro.BorderWidth = lin.BorderWidth
shpt.BorderWidth = lin.BorderWidth
shpb.BorderWidth = lin.BorderWidth
linl.BorderWidth = lin.BorderWidth
linr.BorderWidth = lin.BorderWidth
'captura cor
If Button = vbMiddleButton Then
    Me.MouseIcon = LoadPicture(App.Path + "/gotas.cur")
    picQuadro.ForeColor = frmcapture.BackColor
    shpa.BorderColor = frmcapture.BackColor
    lbla.BackColor = frmcapture.BackColor
    GoTo fim
End If
If Button = vbRightButton Then
    lblborracha.Visible = True
    lblborracha.Top = Y - (lblborracha.Height / 2)
    lblborracha.Left = X - (lblborracha.Width / 2)
    picQuadro.Line ((X - 300), (Y - 300))-((X + 300), (Y + 300)), corborracha, BF
    GoTo fim
End If
If quadrado >= 1 And quadrado <= 4 Then
    linquadro.X2 = X
    linquadro.Y2 = Y
    picQuadro.MousePointer = 2
    
    shpt.X2 = X
    shpb.X2 = X
    shpb.Y2 = Y
    shpb.Y1 = Y
    linl.Y2 = Y
    linr.X1 = X
    linr.X2 = X
    linr.Y2 = Y
    
    shpcircle.Height = txtr.Text * 2
    shpcircle.Width = txtr.Text * 2
    shpcircle.Top = Y - (shpcircle.Height / 2)
    shpcircle.Left = X - (shpcircle.Width / 2)
    GoTo fim
End If
If quadrado = 5 Then
    Me.MouseIcon = LoadPicture()
End If
' prosseguimos desenhando
picQuadro.MousePointer = 0
If desenhando = True Then picQuadro.Line -(X, Y), picQuadro.ForeColor
fim:
'desfaz os "selecionado" dos menus
Dim n As Byte
For n = 0 To 5
     lblArq(n).BackStyle = 0
Next
End Sub
Private Sub pictexto_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
picok.MouseIcon = LoadPicture()
End Sub

Private Sub txtr_KeyPress(KeyAscii As Integer)
'para egitar erros, uma validação de números
If KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Or KeyAscii = vbKeyBack Then
    Exit Sub
Else
    Beep
    KeyAscii = 0
End If
End Sub
Private Sub txtsize_KeyPress(KeyAscii As Integer)
'permite apenas número e backspace
If KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Or KeyAscii = vbKeyBack Then
    Exit Sub
Else
    Beep
    KeyAscii = 0
End If
End Sub
