VERSION 5.00
Object = "{BB35AEF3-E525-4F8B-81F2-511FF805ABB1}#2.1#0"; "ScrollerII.ocx"
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Asistente para instalación"
   ClientHeight    =   6315
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10185
   DrawMode        =   1  'Blackness
   FillColor       =   &H000000D5&
   Icon            =   "selecionnar programas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   421
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   679
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command4 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      DisabledPicture =   "selecionnar programas.frx":187DE
      DownPicture     =   "selecionnar programas.frx":18D78
      Height          =   450
      Left            =   360
      MaskColor       =   &H00C0FFC0&
      Picture         =   "selecionnar programas.frx":19313
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Salir"
      Top             =   9840
      UseMaskColor    =   -1  'True
      Width           =   720
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Contraseña"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   13560
      TabIndex        =   58
      Top             =   6120
      Visible         =   0   'False
      Width           =   2775
      Begin VB.TextBox Text11 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   59
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   360
         TabIndex        =   62
         Top             =   1080
         Width           =   2295
      End
   End
   Begin VB.ComboBox Combo2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "selecionnar programas.frx":198C8
      Left            =   120
      List            =   "selecionnar programas.frx":198D5
      Sorted          =   -1  'True
      TabIndex        =   19
      Top             =   3180
      Width           =   2745
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      DownPicture     =   "selecionnar programas.frx":1991E
      Height          =   500
      Left            =   120
      Picture         =   "selecionnar programas.frx":19F70
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   1800
      Width           =   500
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4695
      Left            =   3600
      TabIndex        =   42
      Top             =   3240
      Visible         =   0   'False
      Width           =   3135
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         Left            =   240
         Min             =   50
         TabIndex        =   92
         Top             =   4320
         Value           =   50
         Width           =   2535
      End
      Begin VB.Timer Timer13 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   2400
         Top             =   3720
      End
      Begin VB.TextBox Text17 
         Height          =   375
         Left            =   1680
         TabIndex        =   91
         Text            =   "Text17"
         Top             =   3720
         Width           =   615
      End
      Begin VB.Timer Timer12 
         Interval        =   50
         Left            =   2520
         Top             =   2280
      End
      Begin VB.Timer Timer11 
         Interval        =   1
         Left            =   2520
         Top             =   1800
      End
      Begin VB.Timer Timer10 
         Interval        =   1
         Left            =   2520
         Top             =   1320
      End
      Begin VB.TextBox poi1 
         Height          =   375
         Left            =   240
         TabIndex        =   90
         Text            =   "Text17"
         Top             =   3720
         Width           =   615
      End
      Begin VB.TextBox poi 
         Height          =   375
         Left            =   960
         TabIndex        =   89
         Text            =   "Text17"
         Top             =   3720
         Width           =   615
      End
      Begin VB.Timer Timer9 
         Interval        =   1
         Left            =   2520
         Top             =   840
      End
      Begin VB.Timer Timer8 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   2040
         Top             =   2280
      End
      Begin VB.CheckBox Check22 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Negrita"
         Height          =   255
         Left            =   1920
         TabIndex        =   79
         Top             =   3240
         Width           =   1095
      End
      Begin VB.TextBox Text16 
         Height          =   375
         Left            =   1320
         TabIndex        =   78
         Text            =   "10"
         Top             =   3240
         Width           =   495
      End
      Begin VB.CheckBox Check21 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Negrita"
         Height          =   375
         Left            =   1920
         TabIndex        =   77
         Top             =   2760
         Width           =   855
      End
      Begin VB.TextBox Text15 
         Height          =   375
         Left            =   1320
         TabIndex        =   76
         Text            =   "16"
         Top             =   2760
         Width           =   495
      End
      Begin VB.TextBox Text14 
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   75
         Text            =   "Arial"
         Top             =   3240
         Width           =   1095
      End
      Begin VB.TextBox Text13 
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   74
         Text            =   "Elephant"
         Top             =   2760
         Width           =   1095
      End
      Begin VB.Timer Timer7 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   2040
         Top             =   1800
      End
      Begin VB.Timer Timer6 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   2040
         Top             =   1320
      End
      Begin VB.CommandButton Command7 
         Caption         =   "configuracion"
         Height          =   255
         Left            =   1080
         TabIndex        =   57
         Top             =   480
         Width           =   1215
      End
      Begin VB.Timer Timer5 
         Interval        =   1
         Left            =   2040
         Top             =   840
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   120
         TabIndex        =   56
         Text            =   "0"
         Top             =   840
         Width           =   375
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   120
         TabIndex        =   55
         Text            =   "0"
         Top             =   1320
         Width           =   375
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   120
         TabIndex        =   54
         Text            =   "0"
         Top             =   1800
         Width           =   375
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   120
         TabIndex        =   53
         Text            =   "0"
         Top             =   2280
         Width           =   375
      End
      Begin VB.TextBox Text5 
         Height          =   375
         Left            =   600
         TabIndex        =   52
         Text            =   "0"
         Top             =   840
         Width           =   375
      End
      Begin VB.TextBox Text6 
         Height          =   375
         Left            =   600
         TabIndex        =   51
         Text            =   "0"
         Top             =   1320
         Width           =   375
      End
      Begin VB.TextBox Text7 
         Height          =   375
         Left            =   1080
         TabIndex        =   50
         Text            =   "0"
         Top             =   840
         Width           =   375
      End
      Begin VB.TextBox Text8 
         Height          =   375
         Left            =   1080
         TabIndex        =   49
         Text            =   "0"
         Top             =   1320
         Width           =   375
      End
      Begin VB.TextBox Text9 
         Height          =   375
         Left            =   1080
         TabIndex        =   48
         Text            =   "0"
         Top             =   1800
         Width           =   375
      End
      Begin VB.TextBox Text10 
         Height          =   375
         Left            =   1080
         TabIndex        =   47
         Text            =   "0"
         Top             =   2280
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Caption         =   "0"
         Height          =   255
         Left            =   840
         TabIndex        =   46
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton Command3 
         Caption         =   "fondo"
         Height          =   255
         Left            =   120
         TabIndex        =   45
         Top             =   240
         Width           =   735
      End
      Begin VB.Timer Timer1 
         Interval        =   500
         Left            =   1560
         Top             =   840
      End
      Begin VB.Timer Timer2 
         Interval        =   1
         Left            =   1560
         Top             =   1320
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Cargar"
         Height          =   255
         Left            =   1560
         TabIndex        =   44
         Top             =   240
         Width           =   855
      End
      Begin VB.Timer Timer3 
         Interval        =   1
         Left            =   1560
         Top             =   1800
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Instalador"
         Height          =   315
         Left            =   120
         TabIndex        =   43
         Top             =   480
         Width           =   975
      End
      Begin VB.Timer Timer4 
         Interval        =   1
         Left            =   1560
         Top             =   2280
      End
      Begin VB.Label anterior 
         Caption         =   "0"
         Height          =   255
         Left            =   2520
         TabIndex        =   93
         Top             =   480
         Width           =   255
      End
      Begin VB.Label ok 
         Caption         =   "0"
         Height          =   255
         Left            =   2400
         TabIndex        =   60
         Top             =   480
         Width           =   495
      End
   End
   Begin VB.TextBox Text12 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   435
      Left            =   120
      TabIndex        =   61
      Top             =   3600
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.CheckBox Check11 
      Caption         =   "Check11"
      Height          =   195
      Left            =   10680
      TabIndex        =   32
      Top             =   1440
      Width           =   190
   End
   Begin VB.CheckBox Check12 
      Caption         =   "Check11"
      Height          =   195
      Left            =   10920
      TabIndex        =   33
      Top             =   1440
      Width           =   190
   End
   Begin VB.CheckBox Check13 
      Caption         =   "Check11"
      Height          =   195
      Left            =   11160
      TabIndex        =   34
      Top             =   1440
      Width           =   190
   End
   Begin VB.CheckBox Check14 
      Caption         =   "Check11"
      Height          =   195
      Left            =   11400
      TabIndex        =   35
      Top             =   1440
      Width           =   190
   End
   Begin VB.CheckBox Check15 
      Caption         =   "Check11"
      Height          =   195
      Left            =   11640
      TabIndex        =   36
      Top             =   1440
      Width           =   190
   End
   Begin VB.CheckBox Check16 
      Caption         =   "Check11"
      Height          =   195
      Left            =   11880
      TabIndex        =   37
      Top             =   1440
      Width           =   190
   End
   Begin VB.CheckBox Check17 
      Caption         =   "Check11"
      Height          =   195
      Left            =   12120
      TabIndex        =   38
      Top             =   1440
      Width           =   190
   End
   Begin VB.CheckBox Check18 
      Caption         =   "Check11"
      Height          =   195
      Left            =   12360
      TabIndex        =   39
      Top             =   1440
      Width           =   190
   End
   Begin VB.CheckBox Check19 
      Caption         =   "Check11"
      Height          =   195
      Left            =   12600
      TabIndex        =   40
      Top             =   1440
      Width           =   190
   End
   Begin VB.CheckBox Check20 
      Caption         =   "Check11"
      Height          =   195
      Left            =   12840
      TabIndex        =   41
      Top             =   1440
      Width           =   190
   End
   Begin VB.CheckBox Check10 
      Caption         =   "Check1"
      Height          =   200
      Index           =   0
      Left            =   12840
      TabIndex        =   31
      Top             =   1080
      Width           =   190
   End
   Begin VB.CheckBox Check9 
      Caption         =   "Check1"
      Height          =   200
      Index           =   0
      Left            =   12600
      TabIndex        =   30
      Top             =   1080
      Width           =   190
   End
   Begin VB.CheckBox Check8 
      Caption         =   "Check1"
      Height          =   200
      Index           =   0
      Left            =   12360
      TabIndex        =   29
      Top             =   1080
      Width           =   190
   End
   Begin VB.CheckBox Check7 
      Caption         =   "Check1"
      Height          =   200
      Index           =   0
      Left            =   12120
      TabIndex        =   28
      Top             =   1080
      Width           =   190
   End
   Begin VB.CheckBox Check6 
      Caption         =   "Check1"
      Height          =   200
      Index           =   0
      Left            =   11880
      TabIndex        =   27
      Top             =   1080
      Width           =   190
   End
   Begin VB.CheckBox Check5 
      Caption         =   "Check1"
      Height          =   200
      Index           =   0
      Left            =   11640
      TabIndex        =   26
      Top             =   1080
      Width           =   190
   End
   Begin VB.CheckBox Check4 
      Caption         =   "Check1"
      Height          =   200
      Index           =   0
      Left            =   11400
      TabIndex        =   25
      Top             =   1080
      Width           =   190
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Check1"
      Height          =   200
      Index           =   0
      Left            =   11160
      TabIndex        =   24
      Top             =   1080
      Width           =   190
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Check1"
      Height          =   200
      Index           =   0
      Left            =   10920
      TabIndex        =   23
      Top             =   1080
      Width           =   190
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      CausesValidation=   0   'False
      Height          =   200
      Index           =   0
      Left            =   10680
      TabIndex        =   2
      Top             =   1080
      Width           =   190
   End
   Begin ScrollerII.FormScroller FormScroller1 
      Left            =   1440
      Top             =   4320
      _ExtentX        =   2170
      _ExtentY        =   1085
      SmallChange     =   100
      LargeChange     =   720
      BackColor       =   -2147483632
      ScrollBarWidth  =   13.333
      ScaleMode       =   0
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   615
      Left            =   15
      TabIndex        =   94
      Top             =   0
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      BorderStyle     =   0  'Transparent
      DrawMode        =   15  'Merge Pen Not
      FillColor       =   &H00000040&
      FillStyle       =   0  'Solid
      Height          =   765
      Index           =   11
      Left            =   0
      Top             =   1680
      Width           =   2835
   End
   Begin VB.Label inst 
      BackStyle       =   0  'Transparent
      Caption         =   "Comenzar la Instalación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   840
      TabIndex        =   0
      Top             =   1680
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      ToolTipText     =   "Salir"
      Top             =   9840
      Width           =   615
   End
   Begin VB.Shape ProgressBar1 
      BackColor       =   &H00FFFFC0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   2  'Dash
      FillStyle       =   4  'Upward Diagonal
      Height          =   255
      Left            =   150
      Top             =   9135
      Width           =   2610
   End
   Begin VB.Label Label16 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "50"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   1305
      TabIndex        =   18
      ToolTipText     =   $"selecionnar programas.frx":1A57F
      Top             =   8280
      Width           =   855
   End
   Begin VB.Label Label15 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "0:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   720
      TabIndex        =   17
      Top             =   8280
      Width           =   1455
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "La Instalación comenzará en"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   855
      Left            =   600
      TabIndex        =   16
      Top             =   7680
      Width           =   1815
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0080FFFF&
      BorderStyle     =   3  'Dot
      DrawMode        =   6  'Mask Pen Not
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   285
      Left            =   120
      Top             =   9120
      Width           =   2655
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      DrawMode        =   3  'Not Merge Pen
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   22
      Left            =   2040
      Top             =   840
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Label porcentaje 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   13080
      TabIndex        =   88
      Top             =   1080
      Width           =   750
   End
   Begin VB.Label apli_des 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "Corbel"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   0
      Left            =   4920
      TabIndex        =   87
      Top             =   8400
      Width           =   90
   End
   Begin VB.Label ext_des 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "Corbel"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   0
      Left            =   7080
      TabIndex        =   86
      Top             =   8400
      Width           =   90
   End
   Begin VB.Label nav_des 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "Corbel"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   0
      Left            =   6840
      TabIndex        =   85
      Top             =   8400
      Width           =   90
   End
   Begin VB.Label ofim_des 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "Corbel"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   0
      Left            =   6360
      TabIndex        =   84
      Top             =   8400
      Width           =   90
   End
   Begin VB.Label con_des 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "Corbel"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   0
      Left            =   6120
      TabIndex        =   83
      Top             =   8400
      Width           =   90
   End
   Begin VB.Label men_des 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "Corbel"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   0
      Left            =   5880
      TabIndex        =   82
      Top             =   8400
      Width           =   90
   End
   Begin VB.Label twek_des 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "Corbel"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   0
      Left            =   6600
      TabIndex        =   81
      Top             =   8400
      Width           =   90
   End
   Begin VB.Label seg_des 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "Corbel"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   0
      Left            =   5640
      TabIndex        =   80
      Top             =   8400
      Width           =   90
   End
   Begin VB.Label nav_nom 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   405
      Index           =   0
      Left            =   9360
      TabIndex        =   73
      Top             =   8280
      Width           =   120
   End
   Begin VB.Label ext_nom 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   405
      Index           =   0
      Left            =   9600
      TabIndex        =   72
      Top             =   8280
      Width           =   120
   End
   Begin VB.Label twek_nom 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   405
      Index           =   0
      Left            =   9120
      TabIndex        =   71
      Top             =   8280
      Width           =   120
   End
   Begin VB.Label grab_des 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "Corbel"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   0
      Left            =   5400
      TabIndex        =   70
      Top             =   8400
      Width           =   90
   End
   Begin VB.Label mul_des 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "Corbel"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   0
      Left            =   5160
      TabIndex        =   69
      Top             =   8400
      Width           =   90
   End
   Begin VB.Label men_nom 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   405
      Index           =   0
      Left            =   8400
      TabIndex        =   68
      Top             =   8280
      Width           =   120
   End
   Begin VB.Label ofim_nom 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   405
      Index           =   0
      Left            =   8880
      TabIndex        =   67
      Top             =   8280
      Width           =   120
   End
   Begin VB.Label con_nom 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   405
      Index           =   0
      Left            =   8640
      TabIndex        =   66
      Top             =   8280
      Width           =   120
   End
   Begin VB.Label seg_nom 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   405
      Index           =   0
      Left            =   8160
      TabIndex        =   65
      Top             =   8280
      Width           =   120
   End
   Begin VB.Label grab_nom 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   405
      Index           =   0
      Left            =   7920
      TabIndex        =   64
      Top             =   8280
      Width           =   120
   End
   Begin VB.Label mul_nom 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   405
      Index           =   0
      Left            =   7680
      TabIndex        =   63
      Top             =   8280
      Width           =   120
   End
   Begin VB.Image Image9 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   0
      Left            =   8640
      Stretch         =   -1  'True
      Top             =   7680
      Width           =   495
   End
   Begin VB.Image Image8 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   0
      Left            =   8160
      Stretch         =   -1  'True
      Top             =   7680
      Width           =   495
   End
   Begin VB.Image Image10 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   0
      Left            =   9120
      Stretch         =   -1  'True
      Top             =   7680
      Width           =   495
   End
   Begin VB.Image Image7 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   0
      Left            =   7680
      Stretch         =   -1  'True
      Top             =   7680
      Width           =   495
   End
   Begin VB.Image Image2 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   480
      Index           =   0
      Left            =   5280
      Stretch         =   -1  'True
      Top             =   7680
      Width           =   480
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   480
      Index           =   0
      Left            =   4800
      Stretch         =   -1  'True
      Top             =   7680
      Width           =   480
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Tweaks y otros"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   5040
      TabIndex        =   22
      Top             =   5880
      Width           =   1935
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Acerca de.."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Index           =   0
      Left            =   360
      TabIndex        =   21
      Top             =   720
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label apli_nom 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Index           =   0
      Left            =   7320
      TabIndex        =   20
      Top             =   8280
      Width           =   165
   End
   Begin VB.Image Image6 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   0
      Left            =   7200
      Stretch         =   -1  'True
      Top             =   7680
      Width           =   495
   End
   Begin VB.Image Image5 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   480
      Index           =   0
      Left            =   6720
      Stretch         =   -1  'True
      Top             =   7680
      Width           =   480
   End
   Begin VB.Image Image4 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   0
      Left            =   6240
      Stretch         =   -1  'True
      Top             =   7680
      Width           =   495
   End
   Begin VB.Image Image3 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   480
      Index           =   0
      Left            =   5760
      Stretch         =   -1  'True
      Top             =   7680
      Width           =   480
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Navegación"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   5040
      TabIndex        =   12
      Top             =   6480
      Width           =   1335
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Ofimática"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   5040
      TabIndex        =   11
      Top             =   5280
      Width           =   1335
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Extra"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   5040
      TabIndex        =   10
      Top             =   7080
      Width           =   1335
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Complemento"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   5040
      TabIndex        =   9
      Top             =   4560
      Width           =   1575
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Mensajería"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   5040
      TabIndex        =   8
      Top             =   3960
      Width           =   1335
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Seguridad"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   5160
      TabIndex        =   7
      Top             =   3360
      Width           =   1335
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Grabación y Backups"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   5160
      TabIndex        =   6
      Top             =   2760
      Width           =   2895
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Multimedia"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   5160
      TabIndex        =   5
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Aplicaciones"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   0
      Left            =   3600
      TabIndex        =   4
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   704
      X2              =   728
      Y1              =   128
      Y2              =   128
   End
   Begin VB.Label selec 
      BackStyle       =   0  'Transparent
      Caption         =   "Selecciones:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   2760
      Width           =   2175
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      BorderStyle     =   0  'Transparent
      DrawMode        =   15  'Merge Pen Not
      FillColor       =   &H00800080&
      FillStyle       =   0  'Solid
      Height          =   525
      Index           =   1
      Left            =   4800
      Top             =   3840
      Visible         =   0   'False
      Width           =   2220
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      BorderStyle     =   0  'Transparent
      DrawMode        =   15  'Merge Pen Not
      FillColor       =   &H00000040&
      FillStyle       =   0  'Solid
      Height          =   525
      Index           =   16
      Left            =   8280
      Top             =   1080
      Visible         =   0   'False
      Width           =   2220
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      BorderStyle     =   0  'Transparent
      DrawMode        =   15  'Merge Pen Not
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   525
      Index           =   0
      Left            =   3180
      Top             =   1080
      Width           =   5115
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      BorderStyle     =   0  'Transparent
      DrawMode        =   15  'Merge Pen Not
      FillColor       =   &H00400040&
      FillStyle       =   0  'Solid
      Height          =   525
      Index           =   2
      Left            =   4800
      Top             =   3840
      Width           =   5115
   End
   Begin VB.Label Label13_ñ 
      BackStyle       =   0  'Transparent
      Caption         =   "Asistente de Instalación Rapida  CIBER CITY "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   585
      Left            =   4800
      TabIndex        =   13
      Top             =   0
      Width           =   11115
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      BorderStyle     =   0  'Transparent
      DrawMode        =   15  'Merge Pen Not
      FillColor       =   &H00008000&
      FillStyle       =   0  'Solid
      Height          =   525
      Index           =   12
      Left            =   4800
      Top             =   2040
      Visible         =   0   'False
      Width           =   2220
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      BorderStyle     =   0  'Transparent
      DrawMode        =   15  'Merge Pen Not
      FillColor       =   &H00000040&
      FillStyle       =   0  'Solid
      Height          =   525
      Index           =   14
      Left            =   4800
      Top             =   3240
      Visible         =   0   'False
      Width           =   2220
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      BorderStyle     =   0  'Transparent
      DrawMode        =   15  'Merge Pen Not
      FillColor       =   &H00808000&
      FillStyle       =   0  'Solid
      Height          =   525
      Index           =   13
      Left            =   4800
      Top             =   2640
      Visible         =   0   'False
      Width           =   2820
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      BorderStyle     =   0  'Transparent
      DrawMode        =   15  'Merge Pen Not
      FillColor       =   &H00800000&
      FillStyle       =   0  'Solid
      Height          =   525
      Index           =   17
      Left            =   4800
      Top             =   4440
      Visible         =   0   'False
      Width           =   2220
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      BorderStyle     =   0  'Transparent
      DrawMode        =   15  'Merge Pen Not
      FillColor       =   &H00400000&
      FillStyle       =   0  'Solid
      Height          =   525
      Index           =   3
      Left            =   4800
      Top             =   4440
      Width           =   5115
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      BorderStyle     =   0  'Transparent
      DrawMode        =   15  'Merge Pen Not
      FillColor       =   &H00000040&
      FillStyle       =   0  'Solid
      Height          =   525
      Index           =   15
      Left            =   4800
      Top             =   5160
      Visible         =   0   'False
      Width           =   2220
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      BorderStyle     =   0  'Transparent
      DrawMode        =   15  'Merge Pen Not
      FillColor       =   &H00404080&
      FillStyle       =   0  'Solid
      Height          =   525
      Index           =   7
      Left            =   4800
      Top             =   5160
      Width           =   5115
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      BorderStyle     =   0  'Transparent
      DrawMode        =   15  'Merge Pen Not
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   525
      Index           =   18
      Left            =   4800
      Top             =   5760
      Visible         =   0   'False
      Width           =   2220
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      BorderStyle     =   0  'Transparent
      DrawMode        =   15  'Merge Pen Not
      FillColor       =   &H00404040&
      FillStyle       =   0  'Solid
      Height          =   525
      Index           =   8
      Left            =   4800
      Top             =   5760
      Width           =   5115
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      BorderStyle     =   0  'Transparent
      DrawMode        =   15  'Merge Pen Not
      FillColor       =   &H00004000&
      FillStyle       =   0  'Solid
      Height          =   525
      Index           =   6
      Left            =   4800
      Top             =   2040
      Width           =   5115
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      BorderStyle     =   0  'Transparent
      DrawMode        =   15  'Merge Pen Not
      FillColor       =   &H00404000&
      FillStyle       =   0  'Solid
      Height          =   525
      Index           =   4
      Left            =   4800
      Top             =   2640
      Width           =   5115
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      BorderStyle     =   0  'Transparent
      DrawMode        =   15  'Merge Pen Not
      FillColor       =   &H00004080&
      FillStyle       =   0  'Solid
      Height          =   525
      Index           =   5
      Left            =   4800
      Top             =   3240
      Width           =   5115
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      BorderStyle     =   0  'Transparent
      DrawMode        =   15  'Merge Pen Not
      FillColor       =   &H00008080&
      FillStyle       =   0  'Solid
      Height          =   525
      Index           =   19
      Left            =   4800
      Top             =   6360
      Visible         =   0   'False
      Width           =   2220
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      BorderStyle     =   0  'Transparent
      DrawMode        =   15  'Merge Pen Not
      FillColor       =   &H00004040&
      FillStyle       =   0  'Solid
      Height          =   525
      Index           =   9
      Left            =   4800
      Top             =   6360
      Width           =   5115
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      BorderStyle     =   0  'Transparent
      DrawMode        =   15  'Merge Pen Not
      FillColor       =   &H00000040&
      FillStyle       =   0  'Solid
      Height          =   525
      Index           =   20
      Left            =   4800
      Top             =   6960
      Visible         =   0   'False
      Width           =   2220
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      BorderStyle     =   0  'Transparent
      DrawMode        =   15  'Merge Pen Not
      FillColor       =   &H00800080&
      FillStyle       =   0  'Solid
      Height          =   525
      Index           =   10
      Left            =   4800
      Top             =   6960
      Width           =   5115
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'------------- ini
 Const APPLICATION As String = "Seleccionador"
      
    Dim Pantalla_completa As Single
    Dim ancho_Seleccionador As Single
    Dim alto_Seleccionador As Single
    Dim m_nombre As Single
    Dim m_descripcion As Single
    Dim m_path As Single
    Dim m_estado As Single
      
    Dim Path_Archivo_Ini As String
      
    'Función api que recupera un valor-dato de un archivo Ini
    Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" ( _
        ByVal lpApplicationName As String, _
        ByVal lpKeyName As String, _
        ByVal lpDefault As String, _
        ByVal lpReturnedString As String, _
        ByVal nSize As Long, _
        ByVal lpFileName As String) As Long
      
    'Función api que Escribe un valor - dato en un archivo Ini
    Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" ( _
        ByVal lpApplicationName As String, _
        ByVal lpKeyName As String, _
        ByVal lpString As String, _
        ByVal lpFileName As String) As Long
      
'----------------------------
    'Función Api GetShortPathName para obtener _
    los paths de los archivos en formato corto
    Private Declare Function GetShortPathName _
        Lib "kernel32" _
        Alias "GetShortPathNameA" ( _
            ByVal lpszLongPath As String, _
            ByVal lpszShortPath As String, _
            ByVal lBuffer As Long) As Long
      
    'Función Api mciExecute para reproducir los archivos de música
    Private Declare Function mciExecute _
        Lib "winmm.dll" ( _
            ByVal lpstrCommand As String) As Long
    Dim ret As Long, path As String
    
    '----------------- declaracion para todos lados
Dim espacio As Integer
Dim fso As Object
Dim obj_FSO As Object
Dim k As Integer
Dim i As Integer
Dim apli As Integer
Dim mul As Integer
Dim grab As Integer
Dim seg As Integer
Dim men As Integer
Dim com As Integer
Dim ofim As Integer
Dim twek As Integer
Dim nav As Integer
Dim ext As Integer
Dim v As Integer
Dim val As Integer
Dim fin As Integer
Dim ruta As String
'Lee un dato _
    -----------------------------
    'Recibe la ruta del archivo, la clave a leer y _
     el valor por defecto en caso de que la Key no exista
    Private Function Leer_Ini(Path_INI As String, Key As String, Default As Variant) As String
      
    Dim bufer As String * 256
    Dim Len_Value As Long
      
            Len_Value = GetPrivateProfileString(APPLICATION, _
                                             Key, _
                                             Default, _
                                             bufer, _
                                             Len(bufer), _
                                             Path_INI)
              
            Leer_Ini = Left$(bufer, Len_Value)
      
    End Function
      
    'Escribe un dato en el INI _
    -----------------------------
    'Recibe la ruta del archivo, La clave a escribir y el valor a añadir en dicha clave
      
    Private Function Grabar_Ini(Path_INI As String, Key As String, Valor As Variant) As String
      
        WritePrivateProfileString APPLICATION, _
                                             Key, _
                                             Valor, _
                                             Path_INI
      
    End Function







'-------------------------

Private Sub Combo2_Change()
Combo2_Click
End Sub

Private Sub Combo2_Click()
If Form1.Combo2.Text = "Selecionar Todos" Then
Call todos
Form1.Combo2.Text = ""
End If

If Form1.Combo2.Text = "No selecionar Ninguno" Then
Call ninguno
Form1.Combo2.Text = ""
End If

If Form1.Combo2.Text = "Selecionar Predeterminados" Then
Call ninguno




Call Desconectar
Call IniciarConexion
i = 1
Set rs = cnn.Execute("SELECT * from aplicaciones")
While rs.EOF = False
If rs!predeterninado = 0 Then
Check1(i).Value = 0
Else
Check1(i).Value = 1
End If
rs.MoveNext
i = i + 1
Wend
'-----------------------
i = 1
Set rs = cnn.Execute("SELECT * from multimedia")
While rs.EOF = False
If rs!predeterninado = 0 Then
Check2(i).Value = 0
Else
Check2(i).Value = 1
End If
rs.MoveNext
i = i + 1
Wend
'-----------------------
i = 1
Set rs = cnn.Execute("SELECT * from Grabación_y_Backups")
While rs.EOF = False
If rs!predeterninado = 0 Then
Check3(i).Value = 0
Else
Check3(i).Value = 1
End If
rs.MoveNext
i = i + 1
Wend
'-----------------------
i = 1
Set rs = cnn.Execute("SELECT * from Seguridad")
While rs.EOF = False
If rs!predeterninado = 0 Then
Check4(i).Value = 0
Else
Check4(i).Value = 1
End If
rs.MoveNext
i = i + 1
Wend
'-----------------------
i = 1
Set rs = cnn.Execute("SELECT * from Mensajería")
While rs.EOF = False
If rs!predeterninado = 0 Then
Check5(i).Value = 0
Else
Check5(i).Value = 1
End If
rs.MoveNext
i = i + 1
Wend
'-----------------------
i = 1
Set rs = cnn.Execute("SELECT * from Complemento")
While rs.EOF = False
If rs!predeterninado = 0 Then
Check6(i).Value = 0
Else
Check6(i).Value = 1
End If
rs.MoveNext
i = i + 1
Wend
'-----------------------
i = 1
Set rs = cnn.Execute("SELECT * from Ofimática")
While rs.EOF = False
If rs!predeterninado = 0 Then
Check7(i).Value = 0
Else
Check7(i).Value = 1
End If
rs.MoveNext
i = i + 1
Wend
'-----------------------
i = 1
Set rs = cnn.Execute("SELECT * from Tweaks_y_otros")
While rs.EOF = False
If rs!predeterninado = 0 Then
Check8(i).Value = 0
Else
Check8(i).Value = 1
End If
rs.MoveNext
i = i + 1
Wend
'-----------------------
i = 1
Set rs = cnn.Execute("SELECT * from Navegación")
While rs.EOF = False
If rs!predeterninado = 0 Then
Check9(i).Value = 0
Else
Check9(i).Value = 1
End If
rs.MoveNext
i = i + 1
Wend
'-----------------------
i = 1
Set rs = cnn.Execute("SELECT * from Extra")
While rs.EOF = False
If rs!predeterninado = 0 Then
Check10(i).Value = 0
Else
Check10(i).Value = 1
End If
rs.MoveNext
i = i + 1
Wend
End If

If Combo2.Text = "comandos" Then
Frame2.Visible = True
Combo2.Text = ""
Text11.SetFocus
End If
Form1.Combo2.Text = ""


End Sub

Private Sub Command1_Click()
If Command1.Caption = 0 Then
Command1.Caption = 1
Else
If Command1.Caption = 1 Then
Command1.Caption = 0
End If
End If
End Sub



Private Sub Command2_Click()
Dim ui As Integer
Dim Resph As Integer
Dim Resphi As Integer

If Timer1.Enabled = False Then
'Resph = MsgBox("¿Desea Iniciar la instalación?" & vbCrLf, vbQuestion + vbYesNo + vbDefaultButton2, "Advertencia")
'If Resph = 6 Then
'End If

'If Combo2.Text = "No selecionar Ninguno" Then
'Resph = MsgBox("No ha selecionado ninguna opción", vbInformation + vbOKOnly, "Información")
'Resphi = MsgBox("¿Desea Salir?", vbQuestion + vbYesNo + vbDefaultButton2, "Advertencia")
'If Resphi = 6 Then
'End
'End If

'Else
Timer1.Enabled = False
Form2.Visible = True
'---------------------------
'Call seleccionar_programas
Form2.SetFocus
'Unload Form1
'End

End If
'End If





End Sub


Private Sub Command3_Click()
Set fso = CreateObject("Scripting.FileSystemObject")
If fso.FileExists(App.path & "\recursos\Wallpaper.jpg") = True Then
Set Me.Picture = LoadPicture(App.path & "\recursos\Wallpaper.jpg")
Else
Form1.Label2.BorderStyle = 0
Form1.Label2.BackStyle = 0

Form1.Label2.Visible = True
Form1.Label2.Caption = "No hay imagen de fondo de pantalla se aplicara el color negro"
' "No hay imagen de fondo de pantalla se aplicara el color negro", vbInformation, "Información"
Form1.backColor = &H0&
End If

End Sub


Private Sub Command4_Click()
End
End Sub

Private Sub Command6_Click()
'Form2.Show
End Sub






Private Sub Form_Activate()
'Label13.Left = 1280
poi.Text = Form1.Width
poi1.Text = Form1.Height

Frame2.Left = 8
Frame2.Top = 296
Shape1(22).Width = Form1.Width
Shape1(22).Height = Form1.Height
Shape1(22).Left = 0
Shape1(22).Top = 0


End Sub

Private Sub Form_Click()
Timer1.Enabled = False
Combo2.Enabled = True
Shape1(11).FillColor = &H80&
End Sub

Private Sub Form_Load()
Dim distancia As String
Dim tyui As Integer
Me.ScaleMode = vbPixels

'-------------------intento de instacion anterior
Dim fso As Object
Dim Rett
Dim Retti


'---------------------
distancia = 360
espacio = 65
apli = Text1.Text
mul = Text2.Text
grab = Text3.Text
seg = Text4.Text
men = Text5.Text
com = Text6.Text
ofim = Text7.Text
twek = Text8.Text
nav = Text9.Text
ext = Text10.Text

'---------------------------labels coordenadas
Label5.Left = Label1(0).Left
Label6.Left = Label1(0).Left
Label7.Left = Label1(0).Left
Label8.Top = Label1(0).Top
Label8.Left = Label1(0).Left + distancia
Label9.Left = Label8.Left
Label11.Left = Label8.Left + distancia
Label11.Top = Label1(0).Top
Label18.Left = Label11.Left
Label12.Left = Label11.Left
Label10.Left = Label11.Left



'------------------
' generar separadores

'------------- multimedia
Load Line1(1)
Load Line1(2)
'------------- seguridad
Load Line1(3)
'------------- mensajeria
Load Line1(4)
'------------- complemento
Load Line1(5)
'------------- ofimatica
Load Line1(6)

'------------- Tweaks
Load Line1(7)

'------------- navegacion
Load Line1(8)
'------------- Extra
Load Line1(9)
'---------- checks
       
   ' --------------- aplicaciones
       Check1(0).Top = Label1(0).Top + 3
       Check1(0).Left = Label1(0).Left - 24
       Call Desconectar
       Call IniciarConexion
       
     i = 1
     Set rs = cnn.Execute("SELECT * from aplicaciones")
     
     While rs.EOF = False
            ' Crea un nuevo control
            Load Check1(i)
            Load Image1(i)
            Load apli_nom(i)
            Load apli_des(i)
            '--------Le establecemos algunas propiedades
            Check1(i).Visible = True
            Check1(i).Top = Label1(0).Top + i * 51

            
            
           '--------------------------------------
            With Image1(i)
            .Visible = True
            .Top = Check1(i).Top - 6
            .Left = Check1(i).Left + 20
            .BorderStyle = 0
            End With
            '-----------------------------------
                'Instanciar el objeto FSO para poder usar las funciones FileExists y FolderExists
                Set fso = CreateObject("Scripting.FileSystemObject")
                ' Comprobar archivo
                If fso.FileExists(App.path & "\recursos\logos\" & rs!nombre) = True Then
                    Image1(i).Picture = LoadPicture(App.path & "\recursos\logos\" & rs!nombre)
                Else
            '-------------- si no existe -----------
                  Image1(i).BorderStyle = 1
                 
                End If
                    Set fso = Nothing
            '---------------- nombre
            With apli_nom(i)
            .Caption = rs!nombre
            .Visible = True
            .Top = Image1(i).Top - 5
            .Left = Image1(i).Left + 40
            End With
            '---------------- descripcion
            With apli_des(i)
            .Caption = rs!descripción
            .Visible = True
            .ForeColor = &H808000
            .Top = apli_nom(i).Top + 22 '+ 20
            .Left = apli_nom(i).Left
            End With
    
            Label5.Top = Check1(i).Top + espacio
            rs.MoveNext
        i = i + 1
        Text1.Text = i - 1
        Wend
        
' --------------- Multimedia
    With Check2(0)
       .Top = Label5.Top + 3
       .Left = Label5.Left - 24
    End With
        
     i = 1
     Set rs = cnn.Execute("SELECT * from multimedia")
     While rs.EOF = False
            
            ' Crea un nuevo control
            Load Check2(i)
            Load Image2(i)
            Load mul_nom(i)
            Load mul_des(i)
            
            'Le establecemos algunas propiedades
           With Check2(i)
           .Visible = True
           .Top = Label5.Top + i * 51
           End With
            '-------------------
            With Image2(i)
            .Visible = True
            .Top = Check2(i).Top - 6
            .Left = Check2(i).Left + 20
            End With
            '-----------------------------------
                'Instanciar el objeto FSO para poder usar las funciones FileExists y FolderExists
                Set fso = CreateObject("Scripting.FileSystemObject")
                ' Comprobar archivo
                If fso.FileExists(App.path & "\recursos\logos\" & rs!nombre) = True Then
                    Image2(i).Picture = LoadPicture(App.path & "\recursos\logos\" & rs!nombre)
                Else
            '-------------- si no existe -----------
                  Image2(i).BorderStyle = 1
                 
                End If
                    Set fso = Nothing
            '---------------- nombre
            With mul_nom(i)
            .Caption = rs!nombre
            .Visible = True
            .Top = Image2(i).Top - 5
            .Left = Image2(i).Left + 40
            End With
            '---------------- descripcion
            With mul_des(i)
            .Caption = rs!descripción
            .Visible = True
            .ForeColor = &H808000
            .Top = mul_nom(i).Top + 24
            .Left = mul_nom(i).Left
            End With

        Label6.Top = Check2(i).Top + espacio
        rs.MoveNext
        i = i + 1
        Text2.Text = i - 1
        Wend



' --------------- Grabacion
        With Check3(0)
        .Top = Label6.Top + 3
        .Left = Label6.Left - 24
        End With
  
     i = 1
     Set rs = cnn.Execute("SELECT * from Grabación_y_Backups")
     While rs.EOF = False
        
        
            ' Crea un nuevo control
            Load Check3(i)
            Load Image3(i)
            Load grab_nom(i)
            Load grab_des(i)
            
            'Le establecemos algunas propiedades
            With Check3(i)
            .Visible = True
            .Top = Label6.Top + i * 51
            End With
           '-------------------
            With Image3(i)
            .Visible = True
            .Top = Check3(i).Top - 6
            .Left = Check3(i).Left + 20
            End With
            '-----------------------------------
                'Instanciar el objeto FSO para poder usar las funciones FileExists y FolderExists
                Set fso = CreateObject("Scripting.FileSystemObject")
                ' Comprobar archivo
                If fso.FileExists(App.path & "\recursos\logos\" & rs!nombre) = True Then
                    Image3(i).Picture = LoadPicture(App.path & "\recursos\logos\" & rs!nombre)
                Else
            '-------------- si no existe -----------
                  Image3(i).BorderStyle = 1
                 
                End If
                    Set fso = Nothing
             '---------------- nombre
            With grab_nom(i)
            .Caption = rs!nombre
            .Visible = True
            .Top = Image3(i).Top - 5
            .Left = Image3(i).Left + 40
            End With
            '---------------- descripcion
            With grab_des(i)
            .Caption = rs!descripción
            .Visible = True
            .ForeColor = &H808000
            .Top = grab_nom(i).Top + 24
            .Left = grab_nom(i).Left
            End With
        Label7.Top = Check3(i).Top + espacio
        rs.MoveNext
        i = i + 1
        Text3.Text = i - 1
        Wend

' --------------- Seguridad
       With Check4(0)
       .Top = Label7.Top + 3
       .Left = Label7.Left - 24
       End With
  
     i = 1
     Set rs = cnn.Execute("SELECT * from seguridad")
     While rs.EOF = False

            ' Crea un nuevo control
            Load Check4(i)
            Load Image4(i)
            Load seg_nom(i)
            Load seg_des(i)
            
            'Le establecemos algunas propiedades
          With Check4(i)
          .Visible = True
          .Top = Label7.Top + i * 51
          End With
            '-------------------
          With Image4(i)
          .Visible = True
          .Top = Check4(i).Top - 6
          .Left = Check4(i).Left + 20
          End With

          '-----------------------------------
                'Instanciar el objeto FSO para poder usar las funciones FileExists y FolderExists
                Set fso = CreateObject("Scripting.FileSystemObject")
                ' Comprobar archivo
                If fso.FileExists(App.path & "\recursos\logos\" & rs!nombre) = True Then
                    Image4(i).Picture = LoadPicture(App.path & "\recursos\logos\" & rs!nombre)
                Else
            '-------------- si no existe -----------
                  Image4(i).BorderStyle = 1
                 
                End If
                    Set fso = Nothing
           '---------------- nombre
            With seg_nom(i)
            .Caption = rs!nombre
            .Visible = True
            .Top = Image4(i).Top - 5
            .Left = Image4(i).Left + 40
            End With
            '---------------- descripcion
            With seg_des(i)
            .Caption = rs!descripción
            .Visible = True
            .ForeColor = &H808000
            .Top = seg_nom(i).Top + 24
            .Left = seg_nom(i).Left
            End With
        rs.MoveNext
        i = i + 1
        Text4.Text = i - 1
        Wend
' --------------- mensajeria
       
       With Check5(0)
       .Top = Label8.Top + 3
       .Left = Label8.Left - 24
       End With
        
         i = 1
     Set rs = cnn.Execute("SELECT * from Mensajería")
     While rs.EOF = False
            ' Crea un nuevo control
            Load Check5(i)
            Load Image5(i)
            Load men_nom(i)
            Load men_des(i)
            
            'Le establecemos algunas propiedades
            With Check5(i)
            .Visible = True
            .Top = Label8.Top + i * 51
            End With
            '-------------------
            With Image5(i)
            .Visible = True
            .Top = Check5(i).Top - 6
            .Left = Check5(i).Left + 20
            End With
            
                     '-----------------------------------
                'Instanciar el objeto FSO para poder usar las funciones FileExists y FolderExists
                Set fso = CreateObject("Scripting.FileSystemObject")
                ' Comprobar archivo
                If fso.FileExists(App.path & "\recursos\logos\" & rs!nombre) = True Then
                    Image5(i).Picture = LoadPicture(App.path & "\recursos\logos\" & rs!nombre)
                Else
            '-------------- si no existe -----------
                  Image5(i).BorderStyle = 1
                 
                End If
                    Set fso = Nothing
            '---------------- nombre
            With men_nom(i)
            .Caption = rs!nombre
            .Visible = True
            .Top = Image5(i).Top - 5
            .Left = Image5(i).Left + 40
            End With
            '---------------- descripcion
            With men_des(i)
            .Caption = rs!descripción
            .Visible = True
            .ForeColor = &H808000
            .Top = men_nom(i).Top + 24
            .Left = men_nom(i).Left
            End With
Label9.Top = Check5(i).Top + espacio
        rs.MoveNext
        i = i + 1
        Text5.Text = i - 1
        Wend


' --------------- complementos
       With Check6(0)
       .Top = Label9.Top + 3
       .Left = Label9.Left - 24
       End With
        
        i = 1
     Set rs = cnn.Execute("SELECT * from complemento")
     While rs.EOF = False
  
            ' Crea un nuevo control
            Load Check6(i)
            Load Image6(i)
            Load con_nom(i)
            Load con_des(i)
            
            'Le establecemos algunas propiedades
            With Check6(i)
            .Visible = True
            .Top = Label9.Top + i * 51
            End With
            '-------------------
            With Image6(i)
            .Visible = True
            .Top = Check6(i).Top - 6
            .Left = Check6(i).Left + 20
            End With
         '-----------------------------------
                'Instanciar el objeto FSO para poder usar las funciones FileExists y FolderExists
                Set fso = CreateObject("Scripting.FileSystemObject")
                ' Comprobar archivo
                If fso.FileExists(App.path & "\recursos\logos\" & rs!nombre) = True Then
                    Image6(i).Picture = LoadPicture(App.path & "\recursos\logos\" & rs!nombre)
                Else
            '-------------- si no existe -----------
                  Image6(i).BorderStyle = 1
                 
                End If
                    Set fso = Nothing
           '---------------- nombre
 
        With con_nom(i)
            .Caption = rs!nombre
            .Visible = True
            .Top = Image6(i).Top - 5
            .Left = Image6(i).Left + 40
            End With
            '---------------- descripcion
            With con_des(i)
            .Caption = rs!descripción
            .Visible = True
            .ForeColor = &H808000
            .Top = con_nom(i).Top + 24
            .Left = con_nom(i).Left
            End With
                  
            
 rs.MoveNext
        i = i + 1
        Text6.Text = i - 1
        Wend

' --------------- ofimatica
       
       With Check7(0)
       .Top = Label11.Top + 3
       .Left = Label11.Left - 24
       End With
  
        i = 1
     Set rs = cnn.Execute("SELECT * from ofimática")
     While rs.EOF = False
            ' Crea un nuevo control
            Load Check7(i)
            Load Image7(i)
            Load ofim_nom(i)
            Load ofim_des(i)
            
            'Le establecemos algunas propiedades
            With Check7(i)
            .Visible = True
            .Top = Label11.Top + i * 51
            End With
            '-------------------
            With Image7(i)
            .Visible = True
            .Top = Check7(i).Top - 6
            .Left = Check7(i).Left + 20
            End With
                 '-----------------------------------
                'Instanciar el objeto FSO para poder usar las funciones FileExists y FolderExists
                Set fso = CreateObject("Scripting.FileSystemObject")
                ' Comprobar archivo
                If fso.FileExists(App.path & "\recursos\logos\" & rs!nombre) = True Then
                    Image7(i).Picture = LoadPicture(App.path & "\recursos\logos\" & rs!nombre)
                Else
            '-------------- si no existe -----------
                  Image7(i).BorderStyle = 1
                 
                End If
                    Set fso = Nothing
            '---------------- nombre
            With ofim_nom(i)
            .Caption = rs!nombre
            .Visible = True
            .Top = Image7(i).Top - 5
            .Left = Image7(i).Left + 40
            End With
            '---------------- descripcion
            With ofim_des(i)
            .Caption = rs!descripción
            .Visible = True
            .ForeColor = &H808000
            .Top = ofim_nom(i).Top + 24
            .Left = ofim_nom(i).Left
            End With
Label18.Top = Check7(i).Top + espacio
        
        rs.MoveNext
        i = i + 1
        Text7.Text = i - 1

        Wend

' --------------- tweaks
       With Check8(0)
       .Top = Label18.Top + 3
       .Left = Label18.Left - 24
       End With
  
         i = 1
     Set rs = cnn.Execute("SELECT * from Tweaks_y_otros")
     While rs.EOF = False
     
            ' Crea un nuevo control
            Load Check8(i)
            Load Image8(i)
            Load twek_nom(i)
            Load twek_des(i)
            
            'Le establecemos algunas propiedades
            With Check8(i)
            .Visible = True
            .Top = Label18.Top + i * 51
            End With
            '-------------------
            With Image8(i)
            .Visible = True
            .Top = Check8(i).Top - 6
            .Left = Check8(i).Left + 20
            End With
                    '-----------------------------------
                'Instanciar el objeto FSO para poder usar las funciones FileExists y FolderExists
                Set fso = CreateObject("Scripting.FileSystemObject")
                ' Comprobar archivo
                If fso.FileExists(App.path & "\recursos\logos\" & rs!nombre) = True Then
                    Image8(i).Picture = LoadPicture(App.path & "\recursos\logos\" & rs!nombre)
                Else
            '-------------- si no existe -----------
                  Image8(i).BorderStyle = 1
                 
                End If
                    Set fso = Nothing
           '---------------- nombre
 With twek_nom(i)
            .Caption = rs!nombre
            .Visible = True
            .Top = Image8(i).Top - 5
            .Left = Image8(i).Left + 40
            End With
            '---------------- descripcion
            With twek_des(i)
            .Caption = rs!descripción
            .Visible = True
            .ForeColor = &H808000
            .Top = twek_nom(i).Top + 24
            .Left = twek_nom(i).Left
            End With
        
        Label12.Top = Check8(i).Top + espacio
        rs.MoveNext
        i = i + 1
        Text8.Text = i - 1

        Wend

' --------------- navegacion
       With Check9(0)
       .Top = Label12.Top + 3
       .Left = Label12.Left - 24
       End With

         i = 1
     Set rs = cnn.Execute("SELECT * from Navegación")
     While rs.EOF = False
            
            ' Crea un nuevo control
            Load Check9(i)
            Load Image9(i)
            Load nav_nom(i)
            Load nav_des(i)
            
            'Le establecemos algunas propiedades
            With Check9(i)
            .Visible = True
            .Top = Label12.Top + i * 51
            End With
            '-------------------
            With Image9(i)
            .Visible = True
            .Top = Check9(i).Top - 6
            .Left = Check9(i).Left + 20
            End With
                     '-----------------------------------
                'Instanciar el objeto FSO para poder usar las funciones FileExists y FolderExists
                Set fso = CreateObject("Scripting.FileSystemObject")
                ' Comprobar archivo
                If fso.FileExists(App.path & "\recursos\logos\" & rs!nombre) = True Then
                    Image9(i).Picture = LoadPicture(App.path & "\recursos\logos\" & rs!nombre)
                Else
            '-------------- si no existe -----------
                  Image9(i).BorderStyle = 1
                 
                End If
                    Set fso = Nothing
           '---------------- nombre

            With nav_nom(i)
            .Caption = rs!nombre
            .Visible = True
            .Top = Image9(i).Top - 5
            .Left = Image9(i).Left + 40
            End With
            '---------------- descripcion
            With nav_des(i)
            .Caption = rs!descripción
            .Visible = True
            .ForeColor = &H808000
            .Top = nav_nom(i).Top + 24
            .Left = nav_nom(i).Left
            End With
            
Label10.Top = Check9(i).Top + espacio
rs.MoveNext
        i = i + 1
        Text9.Text = i - 1

        Wend
        
        ' --------------- extra
       With Check10(0)
       .Top = Label10.Top + 3
       .Left = Label10.Left - 24
       End With
   
         i = 1
     Set rs = cnn.Execute("SELECT * from Extra")
     While rs.EOF = False
     
            ' Crea un nuevo control
            Load Check10(i)
            Load Image10(i)
            Load ext_nom(i)
            Load ext_des(i)
            
            'Le establecemos algunas propiedades
            With Check10(i)
            .Visible = True
            .Top = Label10.Top + i * 51
            End With
            '-------------------
            With Image10(i)
            .Visible = True
            .Top = Check10(i).Top - 6
            .Left = Check10(i).Left + 20
            End With
                     '-----------------------------------
                'Instanciar el objeto FSO para poder usar las funciones FileExists y FolderExists
                Set fso = CreateObject("Scripting.FileSystemObject")
                ' Comprobar archivo
                If fso.FileExists(App.path & "\recursos\logos\" & rs!nombre) = True Then
                    Image10(i).Picture = LoadPicture(App.path & "\recursos\logos\" & rs!nombre)
                Else
            '-------------- si no existe -----------
                  Image10(i).BorderStyle = 1
                 
                End If
                    Set fso = Nothing
           '---------------- nombre
 
            With ext_nom(i)
            .Caption = rs!nombre
            .Visible = True
            .Top = Image10(i).Top - 5
            .Left = Image10(i).Left + 40
            End With
            '---------------- descripcion
            With ext_des(i)
            .Caption = rs!descripción
            .Visible = True
            .ForeColor = &H808000
            .Top = ext_nom(i).Top + 24
            .Left = ext_nom(i).Left
            End With
rs.MoveNext
        i = i + 1
        Text10.Text = i - 1
        Wend


'----------jeje
Command3_Click



Timer8.Enabled = True
       




'---------------------- shapes de colores
Shape1(0).Left = Label1(0).Left - 28
Shape1(0).Top = Label1(0).Top - 8
'------------------------
Shape1(6).Left = Label5.Left - 28
Shape1(6).Top = Label5.Top - 8
'------------------------
Shape1(4).Left = Label6.Left - 28
Shape1(4).Top = Label6.Top - 8
'------------------------
Shape1(5).Left = Label7.Left - 28
Shape1(5).Top = Label7.Top - 8
'-------------------------
Shape1(2).Left = Label8.Left - 28
Shape1(2).Top = Label8.Top - 8
'------------------------
Shape1(3).Left = Label9.Left - 28
Shape1(3).Top = Label9.Top - 8
'------------------------
Shape1(7).Left = Label11.Left - 28
Shape1(7).Top = Label11.Top - 8
'------------------------
Shape1(8).Left = Label18.Left - 28
Shape1(8).Top = Label18.Top - 8
'------------------------
Shape1(9).Left = Label12.Left - 28
Shape1(9).Top = Label12.Top - 8
'------------------------
Shape1(10).Left = Label10.Left - 28
Shape1(10).Top = Label10.Top - 8
'------------------------
Shape1(16).Left = Shape1(0).Left
Shape1(16).Top = Shape1(0).Top

Shape1(12).Left = Shape1(6).Left
Shape1(12).Top = Shape1(6).Top

Shape1(13).Left = Shape1(4).Left
Shape1(13).Top = Shape1(4).Top

Shape1(14).Left = Shape1(5).Left
Shape1(14).Top = Shape1(5).Top

Shape1(1).Left = Shape1(2).Left
Shape1(1).Top = Shape1(2).Top

Shape1(17).Left = Shape1(3).Left
Shape1(17).Top = Shape1(3).Top

Shape1(15).Left = Shape1(7).Left
Shape1(15).Top = Shape1(7).Top

Shape1(18).Left = Shape1(8).Left
Shape1(18).Top = Shape1(8).Top

Shape1(19).Left = Shape1(9).Left
Shape1(19).Top = Shape1(9).Top

Shape1(20).Left = Shape1(10).Left
Shape1(20).Top = Shape1(10).Top
'-----------------------------intento de ini
Set fso = CreateObject("Scripting.FileSystemObject")
If fso.FileExists(Environ("TEMP") & "\Instalador.ini") = True Then
'Form4.Width = 10
'Form4.Height = 10
Form1.Timer1.Enabled = False
Unload Form4
Rett = MsgBox("Se ha detectado una instalación incompleta", vbInformation + vbOKOnly, "Información")
Retti = MsgBox("¿Desea continuar con la instalación anterior?", vbQuestion + vbYesNo + vbDefaultButton2, "Advertencia")
If Retti = 6 Then
Form1.anterior.Caption = 1
Call seleanterior
Form1.Shape1(22).Visible = False
Form1.Timer1.Enabled = False
If Form1.Timer1.Enabled = False Then
Form1.Timer1.Enabled = False
Form2.Visible = True
'Call seleccionar_programas
End If




Else
Kill Environ("TEMP") & "\Instalador.ini"
Form1.anterior.Caption = 0
End If
End If
Set fso = Nothing



End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Form3.Hide
   
End Sub




Private Sub Form_Resize()
If Timer1.Enabled = True Then
Shape1(22).Width = Form1.Width
Shape1(22).Height = Form1.Height
Shape1(22).Left = 0
Shape1(22).Top = 0
Shape1(22).Visible = True
End If
End Sub

Private Sub FormScroller1_AfterScroll(ScrollType As ScrollerII.fsScrollTypes)
Shape1(22).Top = Form1.Top
Shape1(22).Left = Form1.Left
End Sub

Private Sub FormScroller1_BeforeScroll(ScrollType As ScrollerII.fsScrollTypes, Cancel As Boolean)
Shape1(22).Top = Form1.Top
Shape1(22).Left = Form1.Left

End Sub

Private Sub inst_Click()
Command2_Click
End Sub



Private Sub Text11_Change()
If Text11.Text = "C1o2D3i4A5%" Then
Text11.Text = ""
Label19.Caption = ""
Frame2.Visible = False
Text12.Visible = True
Text12.SetFocus
Else
Label19.Caption = "Incorrecto"
End If
End Sub

Private Sub Text12_Change()
If Text12.Text = "ocultar" Then
Text12.Visible = False
Text12.Text = ""

Else


If Text12.Text = "root" Then
Timer5.Enabled = False
Frame1.Visible = True
Text12.Text = ""
Else
If Text12.Text = "normal" Then
Timer5.Enabled = True
Frame1.Visible = False
Text12.Text = ""

Else
If Text12.Text = "cancelar" Then
Timer1.Enabled = False
Text12.Text = ""

Else
If Text12.Text = "reanudar" Then
Timer1.Enabled = True
Label14.Visible = True
Label15.Visible = True
Label16.Visible = True
progressbar1.Visible = True
Text12.Text = ""

Else
If Text12.Text = "instalador" Then
Form2.Visible = True
Text12.Text = ""

Else
If Text12.Text = "salir" Then
End
Text12.Text = ""

Else
If Text12.Text = "cambiar00" Then
Timer6.Enabled = True

Frame1.Caption = 1
Command1.Caption = 0
Form1.Refresh
Text12.Text = ""

Else
If Text12.Text = "cambiar01" Then
Timer6.Enabled = True

Frame1.Caption = 1
Command1.Caption = 1
Form1.Refresh
Text12.Text = ""

Else
If Text12.Text = "fondo" Then
Timer6.Enabled = False
Command3_Click
Text12.Text = ""

Else
If Text12.Text = "rojo" Then
Timer6.Enabled = False
For i = 0 To 9
Line1(i).BorderColor = &HFF&
Text12.Text = ""
Next

Else
If Text12.Text = "verde" Then
Timer6.Enabled = False
For i = 0 To 9
Line1(i).BorderColor = &HFF00&
Next
Text12.Text = ""

Else
If Text12.Text = "azul" Then
Timer6.Enabled = False
For i = 0 To 9
Line1(i).BorderColor = &HFFFF00
Next
Text12.Text = ""

Else
If Text12.Text = "amarillo" Then
Timer6.Enabled = False
For i = 0 To 9
Line1(i).BorderColor = &HFFFF&
Next
Text12.Text = ""


Else
If Text12.Text = "todos" Then
Timer6.Enabled = False
Check11.Value = 1
Check12.Value = 1
Check13.Value = 1
Check14.Value = 1
Check15.Value = 1
Check16.Value = 1
Check17.Value = 1
Check18.Value = 1
Check19.Value = 1
Check20.Value = 1
Text12.Text = ""

Else
If Text12.Text = "ninguno" Then
Timer6.Enabled = False
Check11.Value = 0
Check12.Value = 0
Check13.Value = 0
Check14.Value = 0
Check15.Value = 0
Check16.Value = 0
Check17.Value = 0
Check18.Value = 0
Check19.Value = 0
Check20.Value = 0
Text12.Text = ""




End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If




End Sub

Private Sub Timer1_Timer()
Dim con As String
Dim con2 As String
Dim con3 As String
'Form4.Show vbModeless, Form1
'Form4.Show

If Form4.Visible = True Then
Form4.SetFocus
End If




ruta = App.path & "\recursos\Windows Balloon (2).wav"
'Call sndPlaySound(App.Path & "\recurosos\TimerSound.wav", SND_SYNC + SND_NODEFAULT)
Dim so As String
con = progressbar1.Width
con2 = con



'If ProgressBar1.Width = 1 Then
If porcentaje.Width = 1 Then


'Timer1.Enabled = False
'Unload Form1
'Form2.Show

Combo2.Enabled = True
Timer1.Enabled = False
Form2.Visible = True
'---------------------------
'Call seleccionar_programas
Form2.SetFocus
Else

Form4.Visible = True
Form4.SetFocus
'If ProgressBar1.Width = 1 Then

'Else
progressbar1.Width = progressbar1.Width - 3
'--------------contador
If con2 Mod 3 <> 1 Then
porcentaje.Width = porcentaje.Width - 1
Label16.Caption = porcentaje.Width

If porcentaje.Width = 35 Then
progressbar1.backColor = &HFFC0C0
progressbar1.FillColor = &HFFFFFF
End If



'--------------estilo
If porcentaje.Width = 30 Then
progressbar1.backColor = &H80FF&
progressbar1.FillColor = &HFFFFFF
End If




If porcentaje.Width = 25 Then
progressbar1.backColor = &H80C0FF
End If




If porcentaje.Width = 20 Then
progressbar1.backColor = &H8080FF
Label15.ForeColor = &HC0C0C0
End If

If porcentaje.Width = 15 Then
progressbar1.backColor = &HFF&
Label15.ForeColor = &H808080
End If




If porcentaje.Width = 10 Then
Label16.ForeColor = &HFF&
'----------
Label15.ForeColor = &H404040
Call sndPlaySound(ruta, ASND_SYNC Or SND_NODEFAULT)
End If

If porcentaje.Width = 9 Then
Label16.ForeColor = &HFF&
Label16.Caption = "09"
Form1.SetFocus
Call sndPlaySound(ruta, ASND_SYNC)
Beep
End If

If porcentaje.Width = 8 Then
Label16.ForeColor = &HFF&
Label16.Caption = "08"
Call sndPlaySound(ruta, ASND_SYNC)
Beep
End If

If porcentaje.Width = 7 Then
Label16.ForeColor = &HFF&
Label16.Caption = "07"
Call sndPlaySound(ruta, ASND_SYNC)
Beep
End If

If porcentaje.Width = 6 Then
Label16.ForeColor = &HFF&
Label16.Caption = "06"
Call sndPlaySound(ruta, ASND_SYNC)
Beep
End If

If porcentaje.Width = 5 Then
Label16.ForeColor = &HFF&
Label16.Caption = "05"
Shape1(11).FillColor = &HC0&
'Beep
Call sndPlaySound(ruta, ASND_SYNC)
End If


If porcentaje.Width = 4 Then
Label16.ForeColor = &HFF&
Label16.Caption = "04"
Shape1(11).FillColor = &H80&
'Beep
Call sndPlaySound(ruta, ASND_SYNC)
End If

If porcentaje.Width = 3 Then
Label16.ForeColor = &HFF&
Label16.Caption = "03"
Shape1(11).FillColor = &HC0&
'Beep
Call sndPlaySound(ruta, ASND_SYNC)
End If

If porcentaje.Width = 2 Then
Label16.ForeColor = &HFF&
Label16.Caption = "02"
Shape1(11).FillColor = &H80&
Beep
Call sndPlaySound(ruta, ASND_SYNC)
End If


If porcentaje.Width = 1 Then
Label16.ForeColor = &HFF&
Label16.Caption = "01"
Shape1(11).FillColor = &HC0&
Beep
Call sndPlaySound(ruta, ASND_SYNC)


End If


If porcentaje.Width = 0 Then
Label16.ForeColor = &HFF&
Label16.Caption = "00"

Call sndPlaySound(ruta, ASND_SYNC)
'mciExecute ("Play C:\Users\user\Desktop\instalador_peko\Copia\TimerSound.wav")
End If
End If
'------------------
End If
'End If





End Sub

Private Sub Timer10_Timer()
Path_Archivo_Ini = App.path & "\recursos\config.ini"
Pantalla_completa = Leer_Ini(Path_Archivo_Ini, "Pantalla_completa", 0)
ancho_Seleccionador = Leer_Ini(Path_Archivo_Ini, "ancho_Seleccionador", 0)
alto_Seleccionador = Leer_Ini(Path_Archivo_Ini, "alto_Seleccionador", 0)
    

Set fso = CreateObject("Scripting.FileSystemObject")
If fso.FileExists(Path_Archivo_Ini) = True Then
If Pantalla_completa = 1 Then
Form1.WindowState = 0
Form1.Width = Screen.Width + 200
Form1.Height = Screen.Height + 200
End If

If Pantalla_completa = 0 Then
Form1.WindowState = 0
'MsgBox "With: " & Form1.Width
'MsgBox "Height: " & Form1.Height

Form1.Width = ancho_Seleccionador
Form1.Height = alto_Seleccionador
'Form1.Width = 7560
'Form1.Height = 5340

'Form1.Width = 9000
'Form1.Height = 6870
End If
Else
Form1.Width = Screen.Width + 200
Form1.Height = Screen.Height + 200
End If
Set fso = Nothing
End Sub






Private Sub Timer13_Timer()
'Dim tex As String
'Dim tex1 As String

'tex = "Todas las opciones estan deshabilitadas...  para instalar predetermiandos. Si desea cancelar esta instalacion de clic en el formulario "
'tex1 = "Asistente de Instalación Rapida  CIBER CITY "
'Label13.Left = Label13.Left - 5
'Label13.AutoSize = True



'--------------tex1

'If Timer1.Enabled = True Then
'Text17.Text = Label13.Left - Label13.Width
'Text17.Text = Label13.Left
 '           If Label13.Left = -800 And Label13.Caption = "Asistente de Instalación Rapida  CIBER CITY " Then
  '             Label13.Caption = tex
   '            Label13.ForeColor = &H80FFFF
    '            Label13.Left = 1290
     '       Else
      '
        '    If Label13.Left = -2165 And Label13.Caption = tex Then
       '        Label13.ForeColor = &HFFFFFF
         '      Label13.Caption = tex1
          '      Label13.Left = 1290



'End If
'End If
'End If
End Sub







Private Sub Timer2_Timer()
If Timer1.Enabled = False Then
Form1.Shape1(11).FillColor = &H40& '&H80&
inst.ForeColor = &HFFFFFF
'------------------
Command2.Enabled = True
Label14.Visible = False
Label15.Visible = False
Label16.Visible = False
progressbar1.Visible = False
If progressbar1.Visible = False Then
Shape2.Visible = False
Else
Shape2.Visible = True
End If




'-----------------mensage de todos las opciones deshabilitadas
'Label2.Visible = False
Shape1(22).Visible = False
Unload Form4
'-----------------
'Shape1(21).Visible = False


'Label13.Left = Label13.Left - 5

'Shape1(21).Visible = False
'Label13.Caption = "Asistente de Instalación Rapida  CIBER CITY "
'Label13.ForeColor = &HFFFFFF
'Timer13.Enabled = False
'Label13.Left = 1290




Else
'Form1.SetFocus
End If
End Sub

'--------------

Private Sub Check1_Click(Index As Integer)
'Timer1.Enabled = False
End Sub

Private Sub Check11_Click()
apli = Text1.Text
If Check11.Value = 1 Then
Shape1(16).Visible = True
Label1(0).ForeColor = &H80&
For i = 0 To apli
Check1(i).Value = 1
Next
Check1(1).Value = 1
Else

If Check11.Value = 0 Then
Shape1(16).Visible = False
Label1(0).ForeColor = &HC0C0C0
For i = 0 To apli
If Check1.Item(i).Value = 1 Then
Check1.Item(i).Value = 0
End If
Next
End If
End If


End Sub

Private Sub Check12_Click()
mul = Text2.Text
If Check12.Value = 1 Then
Shape1(12).Visible = True
Label5.ForeColor = &H8000&

For i = 0 To mul
Check2(i).Value = 1
Next
Check2(1).Value = 1
Else
If Check12.Value = 0 Then
Shape1(12).Visible = False
Label5.ForeColor = &HC0C0C0

For i = 0 To mul
If Check2.Item(i).Value = 1 Then
Check2.Item(i).Value = 0
End If
Next
End If
End If
On Error GoTo 0

End Sub

Private Sub Check13_Click()

If Check13.Value = 1 Then
Shape1(13).Visible = True
Label6.ForeColor = &H808000


For i = 0 To Text3.Text
Check3(i).Value = 1
Next
Check3(1).Value = 1
Else
If Check13.Value = 0 Then
Shape1(13).Visible = False
Label6.ForeColor = &HC0C0C0


For i = 0 To Text3.Text
If Check3.Item(i).Value = 1 Then
Check3.Item(i).Value = 0
End If
Next
End If
End If
On Error GoTo 0

End Sub

Private Sub Check14_Click()

If Check14.Value = 1 Then
Shape1(14).Visible = True
Label7.ForeColor = &H4080&

For i = 0 To Text4.Text
Check4(i).Value = 1
Next
Check4(1).Value = 1
Else
If Check14.Value = 0 Then
Shape1(14).Visible = False
Label7.ForeColor = &H8000000F


For i = 0 To Text4.Text
If Check4.Item(i).Value = 1 Then
Check4.Item(i).Value = 0
End If
Next
End If
End If
On Error GoTo 0

End Sub

Private Sub Check15_Click()
If Check15.Value = 1 Then
Shape1(1).Visible = True
Label8.ForeColor = &H800080

For i = 0 To Text5.Text
Check5(i).Value = 1
Next
Check5(1).Value = 1
Else
If Check15.Value = 0 Then
Shape1(1).Visible = False
Label8.ForeColor = &H8000000F

For i = 0 To Text5.Text
If Check5.Item(i).Value = 1 Then
Check5.Item(i).Value = 0
End If
Next
End If
End If
On Error GoTo 0
End Sub

Private Sub Check16_Click()
If Check16.Value = 1 Then
Shape1(17).Visible = True
Label9.ForeColor = &H800000

For i = 0 To Text6.Text
Check6(i).Value = 1
Next
Check6(1).Value = 1
Else
If Check16.Value = 0 Then
Shape1(17).Visible = False
Label9.ForeColor = &H8000000F

For i = 0 To Text6.Text
If Check6.Item(i).Value = 1 Then
Check6.Item(i).Value = 0
End If
Next
End If
End If
On Error GoTo 0
End Sub

Private Sub Check17_Click()
If Check17.Value = 1 Then
Shape1(15).Visible = True
Label11.ForeColor = &H80&

For i = 0 To Text7.Text
Check7(i).Value = 1
Next
Check7(1).Value = 1
Else
If Check17.Value = 0 Then
Shape1(15).Visible = False
Label11.ForeColor = &H8000000F

For i = 0 To Text7.Text
If Check7.Item(i).Value = 1 Then
Check7.Item(i).Value = 0
End If
Next
End If
End If
On Error GoTo 0
End Sub

Private Sub Check18_Click()
If Check18.Value = 1 Then
Shape1(18).Visible = True
Label18.ForeColor = &H404040

For i = 0 To Text8.Text
Check8(i).Value = 1
Next
Check8(1).Value = 1
Else
Shape1(18).Visible = False
Label18.ForeColor = &H8000000F

If Check18.Value = 0 Then
For i = 0 To Text8.Text
If Check8.Item(i).Value = 1 Then
Check8.Item(i).Value = 0
End If
Next
End If
End If
On Error GoTo 0
End Sub



Private Sub Check19_Click()
If Check19.Value = 1 Then
Shape1(19).Visible = True
Label12.ForeColor = &H8080&

For i = 0 To Text9.Text
Check9(i).Value = 1
Next
Check9(1).Value = 1
Else
If Check19.Value = 0 Then
Shape1(19).Visible = False
Label12.ForeColor = &H404040

For i = 0 To Text9.Text
If Check9.Item(i).Value = 1 Then
Check9.Item(i).Value = 0
End If
Next
End If
End If
On Error GoTo 0
End Sub


Private Sub Check20_Click()
If Check20.Value = 1 Then
Shape1(20).Visible = True
Label10.ForeColor = &H400040

For i = 0 To Text10.Text
Check10(i).Value = 1
Next
Check10(1).Value = 1
Else
If Check20.Value = 0 Then
Shape1(20).Visible = False
Label10.ForeColor = &HC0C0C0

For i = 0 To Text10.Text
If Check10.Item(i).Value = 1 Then
Check10.Item(i).Value = 0
End If
Next
End If
End If
On Error GoTo 0
End Sub

Private Sub Timer3_Timer()
apli_des(0).Visible = False
mul_des(0).Visible = False
grab_des(0).Visible = False
seg_des(0).Visible = False
men_des(0).Visible = False
con_des(0).Visible = False
ofim_des(0).Visible = False
twek_des(0).Visible = False
nav_des(0).Visible = False
ext_des(0).Visible = False
'-------------------------
apli_nom(0).Visible = False
mul_nom(0).Visible = False
grab_nom(0).Visible = False
seg_nom(0).Visible = False
men_nom(0).Visible = False
con_nom(0).Visible = False
ofim_nom(0).Visible = False
twek_nom(0).Visible = False
nav_nom(0).Visible = False
ext_nom(0).Visible = False
'---------------------------
Image1(0).Visible = False
Image2(0).Visible = False
Image3(0).Visible = False
Image4(0).Visible = False
Image5(0).Visible = False
Image6(0).Visible = False
Image7(0).Visible = False
Image8(0).Visible = False
Image9(0).Visible = False
Image10(0).Visible = False


End Sub

Private Sub Timer4_Timer()
Check11.Left = Check1(0).Left
Check11.Top = Check1(0).Top

Check12.Left = Check2(0).Left
Check12.Top = Check2(0).Top

Check13.Left = Check3(0).Left
Check13.Top = Check3(0).Top

Check14.Left = Check4(0).Left
Check14.Top = Check4(0).Top

Check15.Left = Check5(0).Left
Check15.Top = Check5(0).Top

Check16.Left = Check6(0).Left
Check16.Top = Check6(0).Top

Check17.Left = Check7(0).Left
Check17.Top = Check7(0).Top

Check18.Left = Check8(0).Left
Check18.Top = Check8(0).Top

Check19.Left = Check9(0).Left
Check19.Top = Check9(0).Top

Check20.Left = Check10(0).Left
Check20.Top = Check10(0).Top


End Sub

Private Sub Timer5_Timer()
Frame1.Visible = False
End Sub

Private Sub Timer6_Timer()
If Frame1.Caption = 1 Then

If Command1.Caption = 0 Then
Set Me.Picture = Nothing

Else
If Command1.Caption = 1 Then
Set Me.Picture = LoadPicture("C:\Users\user\Pictures\Sin título.jpg")

End If
End If
End If
Timer6.Enabled = False
End Sub

Private Sub Timer7_Timer()
'For i = 1 To Text1.Text
'apli_nom(i).Font = Text13.Text
'apli_nom(i).FontSize = Text15.Text
'----------------------------------
'apli_des(i).Font = Text14.Text
'apli_des(i).FontSize = Text16.Text
'Next
End Sub

Private Sub Timer8_Timer()
Dim a As Integer
Dim b As Integer
Dim c As Integer
Dim d As Integer
Dim e As Integer
Dim f As Integer
Dim g As Integer
Dim h As Integer
Dim k As Integer
Dim l As Integer

Dim fuente_nombre As String
Dim fuente_descripcion As String
fuente_nombre = "Elephant"
fuente_descripcion = "Trebuchet MS"

'---------------------------apli
For a = 1 To Text1.Text
If Check1(a).Value = 1 Then


With apli_nom(a)
.Font = fuente_nombre
.ForeColor = &HFFFFFF
.FontBold = False
End With

With apli_des(a)
.ForeColor = &H404040
.FontBold = True
End With
Else

If Check1(a).Value = 0 Then
With apli_nom(a)
.FontBold = True
.Font = fuente_descripcion
.ForeColor = &HC0C0C0
End With

With apli_des(a)
.ForeColor = &H808000
.FontBold = False
End With
End If
End If
Next

'---------------------------mul
For b = 1 To Text2.Text
If Check2(b).Value = 1 Then
With mul_nom(b)
.Font = fuente_nombre
.ForeColor = &HFFFFFF
.FontBold = False
End With
With mul_des(b)
.ForeColor = &H404040
.FontBold = True
End With

Else
If Check2(b).Value = 0 Then
With mul_nom(b)
.FontBold = True
.Font = fuente_descripcion
.ForeColor = &HC0C0C0
End With

With mul_des(b)
.ForeColor = &H808000
.FontBold = False
End With
End If
End If
Next
'---------------------------grab
For c = 1 To Text3.Text
If Check3(c).Value = 1 Then
With grab_nom(c)
.Font = fuente_nombre
.ForeColor = &HFFFFFF
.FontBold = False
End With
With grab_des(c)
.ForeColor = &H404040
.FontBold = True
End With

Else
If Check3(c).Value = 0 Then
With grab_nom(c)
.FontBold = True
.Font = fuente_descripcion
.ForeColor = &HC0C0C0
End With

With grab_des(c)
.ForeColor = &H808000
.FontBold = False
End With
End If
End If
Next

'---------------------------seg
For d = 1 To Text4.Text
If Check4(d).Value = 1 Then
With seg_nom(d)
.Font = fuente_nombre
.ForeColor = &HFFFFFF
.FontBold = False
End With
With seg_des(d)
.ForeColor = &H404040
.FontBold = True
End With

Else
If Check4(d).Value = 0 Then
With seg_nom(d)
.FontBold = True
.Font = fuente_descripcion
.ForeColor = &HC0C0C0
End With

With seg_des(d)
.ForeColor = &H808000
.FontBold = False
End With
End If
End If
Next

'---------------------------men
For e = 1 To Text5.Text
If Check5(e).Value = 1 Then
With men_nom(e)
.Font = fuente_nombre
.ForeColor = &HFFFFFF
.FontBold = False
End With
With men_des(e)
.ForeColor = &H404040
.FontBold = True
End With

Else
If Check5(e).Value = 0 Then
With men_nom(e)
.FontBold = True
.Font = fuente_descripcion
.ForeColor = &HC0C0C0
End With

With men_des(e)
.ForeColor = &H808000
.FontBold = False
End With
End If
End If
Next

'---------------------------com
For f = 1 To Text6.Text
If Check6(f).Value = 1 Then
With con_nom(f)
.Font = fuente_nombre
.ForeColor = &HFFFFFF
.FontBold = False
End With
With con_des(f)
.ForeColor = &H404040
.FontBold = True
End With

Else
If Check6(f).Value = 0 Then
With con_nom(f)
.FontBold = True
.Font = fuente_descripcion
.ForeColor = &HC0C0C0
End With

With con_des(f)
.ForeColor = &H808000
.FontBold = False
End With
End If
End If
Next

'---------------------------ofim
For g = 1 To Text7.Text
If Check7(g).Value = 1 Then
With ofim_nom(g)
.Font = fuente_nombre
.ForeColor = &HFFFFFF
.FontBold = False
End With
With ofim_des(g)
.ForeColor = &H404040
.FontBold = True
End With

Else
If Check7(g).Value = 0 Then
With ofim_nom(g)
.FontBold = True
.Font = fuente_descripcion
.ForeColor = &HC0C0C0
End With

With ofim_des(g)
.ForeColor = &H808000
.FontBold = False
End With
End If
End If
Next

'---------------------------twek
For h = 1 To Text8.Text
If Check8(h).Value = 1 Then
With twek_nom(h)
.Font = fuente_nombre
.ForeColor = &HFFFFFF
.FontBold = False
End With
With twek_des(h)
.ForeColor = &H404040
.FontBold = True
End With

Else
If Check8(h).Value = 0 Then
With twek_nom(h)
.FontBold = True
.Font = fuente_descripcion
.ForeColor = &HC0C0C0
End With

With twek_des(h)
.ForeColor = &H808000
.FontBold = False
End With
End If
End If
Next

'---------------------------nav
For k = 1 To Text9.Text
If Check9(k).Value = 1 Then
With nav_nom(k)
.Font = fuente_nombre
.ForeColor = &HFFFFFF
.FontBold = False
End With
With nav_des(k)
.ForeColor = &H404040
.FontBold = True
End With

Else
If Check9(k).Value = 0 Then
With nav_nom(k)
.FontBold = True
.Font = fuente_descripcion
.ForeColor = &HC0C0C0
End With

With nav_des(k)
.ForeColor = &H808000
.FontBold = False
End With
End If
End If
Next

'---------------------------ext
For l = 1 To Text10.Text
If Check10(l).Value = 1 Then
With ext_nom(l)
.Font = fuente_nombre
.ForeColor = &HFFFFFF
.FontBold = False
End With
With ext_des(l)
.ForeColor = &H404040
.FontBold = True
End With

Else
If Check10(l).Value = 0 Then
With ext_nom(l)
.FontBold = True
.Font = fuente_descripcion
.ForeColor = &HC0C0C0
End With

With ext_des(l)
.ForeColor = &H808000
.FontBold = False
End With
End If
End If
Next



End Sub

Private Sub Timer9_Timer()
'----------- acomodar lineas
With Line1(0)
.BorderColor = &HFFFF00
.Y1 = Label1(0).Top + 26
.Y2 = Label1(0).Top + 26
.X1 = Label1(0).Left - 28
.X2 = Label1(0).Left + 312
End With

With Line1(1)
.Visible = True
.Y1 = Label5.Top + 26
.Y2 = Label5.Top + 26
.X1 = Label5.Left - 28
.X2 = Label5.Left + 312
.BorderColor = Line1(0).BorderColor
End With

With Line1(2)
.Visible = True
.Y1 = Label6.Top + 26
.Y2 = Label6.Top + 26
.X1 = Label6.Left - 28
.X2 = Label6.Left + 312
.BorderColor = Line1(0).BorderColor
End With

With Line1(3)
.Visible = True
.Y1 = Label7.Top + 26
.Y2 = Label7.Top + 26
.X1 = Label7.Left - 28
.X2 = Label7.Left + 312
.BorderColor = Line1(0).BorderColor
End With

With Line1(4)
.Visible = True
.Y1 = Label8.Top + 26
.Y2 = Label8.Top + 26
.X1 = Label8.Left - 28
.X2 = Label8.Left + 312
.BorderColor = Line1(0).BorderColor
End With

With Line1(5)
.Visible = True
.Y1 = Label9.Top + 26
.Y2 = Label9.Top + 26
.X1 = Label9.Left - 28
.X2 = Label9.Left + 312
.BorderColor = Line1(0).BorderColor
End With

With Line1(6)
.Visible = True
.Y1 = Label11.Top + 26
.Y2 = Label11.Top + 26
.X1 = Label11.Left - 28
.X2 = Label11.Left + 312
.BorderColor = Line1(0).BorderColor
End With

With Line1(7)
.Visible = True
.Y1 = Label18.Top + 26
.Y2 = Label18.Top + 26
.X1 = Label18.Left - 28
.X2 = Label18.Left + 312
.BorderColor = Line1(0).BorderColor
End With

With Line1(8)
.Visible = True
.Y1 = Label12.Top + 26
.Y2 = Label12.Top + 26
.X1 = Label12.Left - 28
.X2 = Label12.Left + 312
.BorderColor = Line1(0).BorderColor
End With

With Line1(9)
.Visible = True
.Y1 = Label10.Top + 26
.Y2 = Label10.Top + 26
.X1 = Label10.Left - 28
.X2 = Label10.Left + 312
.BorderColor = Line1(0).BorderColor
End With
End Sub
