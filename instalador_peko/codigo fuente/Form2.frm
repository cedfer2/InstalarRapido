VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Form2 
   BackColor       =   &H8000000C&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Asistente de Instalación"
   ClientHeight    =   7125
   ClientLeft      =   210
   ClientTop       =   1410
   ClientWidth     =   14205
   DrawStyle       =   5  'Transparent
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   475
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   947
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   255
      Left            =   11880
      TabIndex        =   40
      Top             =   5040
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "Form2.frx":12F4B
      Left            =   4680
      List            =   "Form2.frx":12F4D
      Style           =   2  'Dropdown List
      TabIndex        =   38
      Top             =   6360
      Width           =   1875
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   495
      Left            =   8640
      TabIndex        =   36
      Top             =   3000
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   495
      Left            =   7320
      TabIndex        =   17
      Top             =   3480
      Width           =   1335
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   7560
      TabIndex        =   16
      Text            =   "Text5"
      Top             =   5520
      Width           =   3255
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   8520
      TabIndex        =   7
      Top             =   2400
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   7320
      TabIndex        =   6
      Top             =   3000
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   7080
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   1200
      Width           =   7215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   1575
      Left            =   8640
      TabIndex        =   0
      Top             =   3960
      Visible         =   0   'False
      Width           =   2175
      Begin VB.Timer Timer5 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   1680
         Top             =   360
      End
      Begin VB.Timer Timer4 
         Interval        =   1
         Left            =   1440
         Top             =   840
      End
      Begin VB.Timer Timer3 
         Interval        =   1
         Left            =   1080
         Top             =   240
      End
      Begin VB.Timer Timer2 
         Interval        =   1
         Left            =   600
         Top             =   240
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   375
         Left            =   720
         TabIndex        =   4
         Top             =   1200
         Width           =   975
      End
      Begin VB.Timer Timer1 
         Interval        =   1
         Left            =   120
         Top             =   240
      End
      Begin VB.Label salida 
         Caption         =   "0"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label estado 
         Caption         =   "0"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   1200
         Width           =   255
      End
      Begin VB.Label iniciado 
         Caption         =   "0"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   840
         Width           =   375
      End
   End
   Begin VB.Label Label25 
      Caption         =   "---"
      Height          =   615
      Left            =   10680
      TabIndex        =   39
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label24 
      BackStyle       =   0  'Transparent
      Caption         =   "La instalción a terminado. Cerrando en: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1995
      Left            =   840
      TabIndex        =   37
      Top             =   2280
      Visible         =   0   'False
      Width           =   5130
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H00FFFFFF&
      DrawMode        =   3  'Not Merge Pen
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   975
      Left            =   9360
      Top             =   -120
      Width           =   855
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H008080FF&
      BorderColor     =   &H00000000&
      BorderStyle     =   0  'Transparent
      DrawMode        =   14  'Copy Pen
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   8280
      Top             =   600
      Width           =   300
   End
   Begin VB.Label Label23 
      Caption         =   "Label23"
      Height          =   255
      Left            =   9840
      TabIndex        =   35
      Top             =   6720
      Width           =   975
   End
   Begin VB.Label Label22 
      Caption         =   "Label22"
      Height          =   255
      Left            =   8760
      TabIndex        =   34
      Top             =   6720
      Width           =   975
   End
   Begin VB.Label Label21 
      Caption         =   "Label21"
      Height          =   255
      Left            =   7560
      TabIndex        =   33
      Top             =   6720
      Width           =   975
   End
   Begin VB.Label Label20 
      Caption         =   "Label20"
      Height          =   255
      Left            =   8760
      TabIndex        =   32
      Top             =   6360
      Width           =   975
   End
   Begin VB.Label Label19 
      Caption         =   "Label19"
      Height          =   255
      Left            =   7560
      TabIndex        =   31
      Top             =   6360
      Width           =   975
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "/"
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   330
      TabIndex        =   30
      Top             =   6900
      Width           =   75
   End
   Begin VB.Label Text4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   75
      TabIndex        =   29
      Top             =   6900
      Width           =   90
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   480
      TabIndex        =   28
      Top             =   6900
      Width           =   90
   End
   Begin VB.Label Label16 
      Caption         =   "0"
      Height          =   495
      Left            =   12600
      TabIndex        =   27
      Top             =   3840
      Width           =   615
   End
   Begin VB.Label Label15 
      Caption         =   "0"
      Height          =   495
      Left            =   12600
      TabIndex        =   26
      Top             =   3240
      Width           =   615
   End
   Begin VB.Label Label14 
      Caption         =   "0"
      Height          =   495
      Left            =   12600
      TabIndex        =   25
      Top             =   2640
      Width           =   615
   End
   Begin VB.Label Label13 
      Caption         =   "0"
      Height          =   495
      Left            =   12600
      TabIndex        =   24
      Top             =   2040
      Width           =   615
   End
   Begin VB.Label Label12 
      Caption         =   "0"
      Height          =   495
      Left            =   11760
      TabIndex        =   23
      Top             =   2640
      Width           =   615
   End
   Begin VB.Label Label11 
      Caption         =   "0"
      Height          =   495
      Left            =   11760
      TabIndex        =   22
      Top             =   2040
      Width           =   615
   End
   Begin VB.Label Label10 
      Caption         =   "0"
      Height          =   495
      Left            =   11040
      TabIndex        =   21
      Top             =   3840
      Width           =   615
   End
   Begin VB.Label Label9 
      Caption         =   "0"
      Height          =   495
      Left            =   11040
      TabIndex        =   20
      Top             =   3240
      Width           =   615
   End
   Begin VB.Label Label8 
      Caption         =   "0"
      Height          =   495
      Left            =   11040
      TabIndex        =   19
      Top             =   2640
      Width           =   615
   End
   Begin VB.Label Label7 
      Caption         =   "0"
      Height          =   495
      Left            =   11040
      TabIndex        =   18
      Top             =   2040
      Width           =   615
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "?????????????????"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   150
      TabIndex        =   15
      Top             =   480
      Width           =   11295
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100 %"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   6840
      TabIndex        =   14
      Top             =   600
      Width           =   765
   End
   Begin MSForms.CommandButton CommandButton1 
      Height          =   330
      Left            =   5400
      TabIndex        =   13
      Top             =   120
      Visible         =   0   'False
      Width           =   1155
      ForeColor       =   16777215
      BackColor       =   192
      Caption         =   "CANCELAR"
      PicturePosition =   131072
      Size            =   "2037;582"
      FontName        =   "Arial Rounded MT Bold"
      FontHeight      =   135
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00E0E0E0&
      DrawMode        =   5  'Not Copy Pen
      X1              =   176
      X2              =   428
      Y1              =   383
      Y2              =   383
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Asistente de Instalación Rapida"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   240
      TabIndex        =   12
      Top             =   5640
      Width           =   2235
   End
   Begin MSForms.Label Label4 
      Height          =   285
      Left            =   4680
      TabIndex        =   11
      Top             =   6840
      Width           =   2100
      ForeColor       =   16777215
      BackColor       =   12632256
      PicturePosition =   6
      Size            =   "3704;503"
      BorderStyle     =   1
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.Label Label6 
      Height          =   285
      Left            =   0
      TabIndex        =   10
      Top             =   6840
      Width           =   810
      BackColor       =   12632256
      Size            =   "1429;503"
      BorderStyle     =   1
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.TextBox Text1 
      Height          =   3870
      Left            =   250
      TabIndex        =   9
      Top             =   1350
      Width           =   6045
      VariousPropertyBits=   -1935652845
      BackColor       =   33023
      ForeColor       =   4210752
      BorderStyle     =   1
      ScrollBars      =   3
      Size            =   "10663;6826"
      BorderColor     =   16777215
      SpecialEffect   =   0
      FontName        =   "Lucida Sans Unicode"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Shape progressbar1 
      BackColor       =   &H008080FF&
      BorderColor     =   &H00000000&
      BorderStyle     =   0  'Transparent
      DrawMode        =   14  'Copy Pen
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   135
      Top             =   735
      Width           =   75
   End
   Begin VB.Label Label1 
      BackColor       =   &H00400000&
      BackStyle       =   0  'Transparent
      Caption         =   "tamaño max por categoria: 721"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   7680
      TabIndex        =   8
      Top             =   1800
      Width           =   6615
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      DrawMode        =   15  'Merge Pen Not
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   3930
      Left            =   240
      Top             =   1320
      Width           =   6075
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   0
      Left            =   -120
      Top             =   6840
      Width           =   7335
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00C0C0C0&
      BorderColor     =   &H0080FFFF&
      BorderStyle     =   0  'Transparent
      DrawMode        =   2  'Blackness
      FillColor       =   &H0080C0FF&
      FillStyle       =   0  'Solid
      Height          =   405
      Left            =   0
      Top             =   5520
      Width           =   6795
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0080FFFF&
      BorderStyle     =   3  'Dot
      DrawMode        =   5  'Not Copy Pen
      FillColor       =   &H00FFFFFF&
      FillStyle       =   5  'Downward Diagonal
      Height          =   405
      Left            =   120
      Top             =   720
      Width           =   6495
   End
   Begin VB.Image Image1 
      Height          =   6840
      Left            =   0
      Picture         =   "Form2.frx":12F4F
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6780
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit
    'Estructura NOTIFYICONDATA para usar con Shell_NotifyIcon
    Private Type NOTIFYICONDATA
       cbSize As Long
       hwnd As Long
       uID As Long
       uFlags As Long
       uCallbackMessage As Long
       hIcon As Long
       szTip As String * 128
       dwState As Long
       dwStateMask As Long
       szInfo As String * 256
       uTimeout As Long
       szInfoTitle As String * 64
       dwInfoFlags As Long
    End Type
      
    'Variable para la estructura anterior
    Private sysTray As NOTIFYICONDATA
      
      
    'Constantes
    Private Const NOTIFYICON_VERSION = 3
    Private Const NOTIFYICON_OLDVERSION = 0
      
    Private Const NIM_ADD = &H0
    Private Const NIM_MODIFY = &H1
    Private Const NIM_DELETE = &H2
      
    Private Const NIM_SETFOCUS = &H3
    Private Const NIM_SETVERSION = &H4
      
    Private Const NIF_MESSAGE = &H1
    Private Const NIF_ICON = &H2
    Private Const NIF_TIP = &H4
      
    Private Const NIF_STATE = &H8
    Private Const NIF_INFO = &H10
      
    Private Const NIS_HIDDEN = &H1
    Private Const NIS_SHAREDICON = &H2
      
    Private Const NIIF_NONE = &H0
    Private Const NIIF_WARNING = &H2
    Private Const NIIF_ERROR = &H3
    Private Const NIIF_INFO = &H1
    Private Const NIIF_GUID = &H4
      
    Private Const WM_MOUSEMOVE = &H200
    Private Const WM_LBUTTONDOWN = &H201
    Private Const WM_LBUTTONUP = &H202
    Private Const WM_LBUTTONDBLCLK = &H203
    Private Const WM_RBUTTONDOWN = &H204
    Private Const WM_RBUTTONUP = &H205
    Private Const WM_RBUTTONDBLCLK = &H206
       
    ' Declaración Api
    Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" _
    (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
    
'----------------------------------------------
Private Declare Function ShellExecute Lib "shell32.dll" Alias _
    "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, _
    ByVal lpFile As String, ByVal lpParameters As String, _
    ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
'----------------------------------------------
    Dim ruta1 As String
    Dim ruta2 As String
    Dim ruta3 As String
    Dim ruta4 As String
    Dim ruta5 As String
    

Private Sub Combo1_Change()
Combo1_Click
End Sub

'
'

Private Sub Combo1_Click()
 If Form2.Combo1.Text = "Apagar" Then
    frmToolTip.Caption = "Apagar"
 '   Unload Form1
'    Unload Form2
    Load frmToolTip
    frmToolTip.Visible = True
    frmToolTip.Label1.Caption = "El equipo se apagará en:"
    frmToolTip.Shape1.Width = 20
    frmToolTip.Label16.Visible = False
    frmToolTip.Label3.Visible = False
    frmToolTip.Timer1.Enabled = True
 End If
 
 If Form2.Combo1.Text = "Reiniciar" Then
 frmToolTip.Caption = "Reiniciar"
 '   Unload Form1
'    Unload Form2
    Load frmToolTip
    frmToolTip.Visible = True
    frmToolTip.Label1.Caption = "El equipo se reiniciará en:"
    frmToolTip.Shape1.Width = 20
    frmToolTip.Label16.Visible = False
    frmToolTip.Label3.Visible = False
    frmToolTip.Timer1.Enabled = True
 End If
 
 
End Sub

Private Sub Command2_Click()
Dim fer As String
'Form2.progressbar1.Width = Form1.Text1.Text / 9
'fer = Form2.progressbar1.Width = 721 / Form2.Text4.Text


End Sub

Private Sub Command3_Click()
Form2.Text1.Text = Form2.Text1.Text & "........................................"
Call seleccionar_programas
End Sub


Private Sub Command5_Click()
Kill Environ("TEMP") & "\Instalador.ini"
End Sub

Private Sub CommandButton1_Click()
Dim Resp As Integer
Resp = MsgBox("¿Desea Cancelar?" & vbCrLf, vbQuestion + vbYesNo + vbDefaultButton2, "Alerta")
If Resp = 6 Then
End
End If


End Sub




Private Sub Form_Activate()
Form1.Icon = Nothing
'Form1.ShowInTaskbar = False

Form1.Timer10.Enabled = False
Form1.BorderStyle = 5
Form1.WindowState = 0
Form1.Width = 1980
Form1.Height = 1320
Form2.Top = Form1.Top + Form1.Height
End Sub

Private Sub Form_Load()
Form2.Combo1.AddItem ("Apagar")
Form2.Combo1.AddItem ("Reiniciar")

Form2.Width = 6870 '6840
Form2.Height = 7455 '8310

' Call seleccionar_programas
Form1.Timer1.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
ruta3 = App.path & "\recursos\Exit.WAV"
Call sndPlaySound(ruta3, ASND_SYNC)
End
End Sub


Private Sub Label24_Click()
Kill Environ("TEMP") & "\Instalador.ini"
Timer5.Enabled = False
Label24.Visible = False
Label5.Caption = "Cerrado cancelado"
Shape6.Visible = False
Form2.Caption = "Asistente de Instalación  -Cerrado Cancelado"
Form2.Combo1.Enabled = True
End Sub

Private Sub Text5_Change()
Form2.Text1.Text = Form2.Text1.Text & Form2.Text5.Text
End Sub

Private Sub Timer1_Timer()
ruta1 = App.path & "\recursos\AtBeginning.wav"
ruta2 = App.path & "\recursos\AtEnd.wav"

If iniciado.Caption = 1 Then
Call sndPlaySound(ruta1, ASND_SYNC)
iniciado.Caption = 0
Else
If iniciado.Caption = 2 Then
Call sndPlaySound(ruta2, ASND_SYNC)
End
End If
End If




End Sub

Private Sub Timer2_Timer()
Form2.Label4.Caption = Date & "  " & Time
End Sub

Private Sub Timer3_Timer()
Call seleccionar_programas
Form2.Timer3.Enabled = False
'Form2.iniciado.Caption = 2
End Sub

Private Sub Timer5_Timer()
If Shape5.Width = 1 Then
Command5_Click
End

Else
CommandButton1.Visible = False
Shape5.Width = Shape5.Width - 1
Form2.Caption = "Asistente de Instalación  -Cerrando en: " & Shape5.Width
Label24.Caption = "La instalción a terminado." & vbCrLf & "Cerrando en: " & Shape5.Width & vbCrLf & "Si desea Cancelar cerrado..." & vbCrLf & "De clic en este mensaje"
End If
End Sub
