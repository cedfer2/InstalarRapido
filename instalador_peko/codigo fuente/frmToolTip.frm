VERSION 5.00
Begin VB.Form frmToolTip 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   1605
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3540
   DrawMode        =   5  'Not Copy Pen
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   107
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   236
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   800
      Left            =   3000
      Top             =   120
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   2400
      TabIndex        =   4
      ToolTipText     =   $"frmToolTip.frx":0000
      Top             =   -120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label2 
      Height          =   255
      Left            =   1800
      TabIndex        =   3
      Top             =   0
      Width           =   735
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   255
      Left            =   120
      Top             =   1320
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Label Label16 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   855
      Left            =   1200
      TabIndex        =   1
      ToolTipText     =   $"frmToolTip.frx":0088
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "????????????"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   2655
   End
   Begin VB.Label Label15 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "0:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   735
      Left            =   600
      TabIndex        =   2
      Top             =   600
      Width           =   1455
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   1125
      Left            =   2400
      Picture         =   "frmToolTip.frx":0110
      Top             =   480
      Width           =   1125
   End
End
Attribute VB_Name = "frmToolTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
With Shape1
.BackStyle = 0
.FillStyle = 1
.BackStyle = 0
End With
Dim purgs As Integer
If Shape1.Width = 1 Then
Timer1.Enabled = False
If Form2.Combo1.Text = "Apagar" Then
Label16.Caption = 0
Call ShutDownNT(True)


End If
If Form2.Combo1.Text = "Reiniciar" Then
Label16.Caption = 0
Call RebootNT(True)
End If


Else
Label16.Visible = True
Label16.Top = Label15.Top
Label16.Left = Label15.Left + 50

Shape1.Width = Shape1.Width - 1
Label16.Caption = Shape1.Width
frmToolTip.Caption = Form2.Combo1.Text & " en: " & Shape1.Width
If Shape1.Width <= 9 Then
Label3.Visible = True
Label3.Top = Label15.Top
Label3.Left = Label15.Left + 50
Label16.Left = Label3.Left + 30
End If
End If
End Sub
