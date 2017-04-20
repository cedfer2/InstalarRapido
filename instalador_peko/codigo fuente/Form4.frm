VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Form4"
   ClientHeight    =   2790
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6825
   LinkTopic       =   "Form4"
   Picture         =   "Form4.frx":0000
   ScaleHeight     =   2790
   ScaleWidth      =   6825
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Command1"
      Height          =   255
      Left            =   5880
      TabIndex        =   2
      Top             =   2880
      Width           =   495
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   135
      Left            =   2880
      TabIndex        =   0
      Top             =   3000
      Width           =   2775
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   $"Form4.frx":5A993
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
      Height          =   2625
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   6795
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
End
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Activate()
If Form1.anterior.Caption = 0 Then
Form1.Combo2.Text = "Selecionar Predeterminados"
End If
End Sub

Private Sub Form_Click()
Unload Form4
Form1.Shape1(22).Visible = False
Form1.Timer1.Enabled = False
End Sub

Private Sub Form_Load()
HScroll1.Max = 255
HScroll1.Min = 50
HScroll1.Value = 150

Form1.Combo2.Enabled = False
Form1.Check11.Enabled = False
Form1.Check12.Enabled = False
Form1.Check13.Enabled = False
Form1.Check14.Enabled = False
Form1.Check15.Enabled = False
Form1.Check16.Enabled = False
Form1.Check17.Enabled = False
Form1.Check18.Enabled = False
Form1.Check19.Enabled = False
Form1.Check20.Enabled = False

For f = 1 To Form1.Text1.Text
Form1.Check1(f).Enabled = False
Next

For f = 1 To Form1.Text2.Text
Form1.Check2(f).Enabled = False
Next

For f = 1 To Form1.Text3.Text
Form1.Check3(f).Enabled = False
Next

For f = 1 To Form1.Text4.Text
Form1.Check4(f).Enabled = False
Next

For f = 1 To Form1.Text5.Text
Form1.Check5(f).Enabled = False
Next

For f = 1 To Form1.Text6.Text
Form1.Check6(f).Enabled = False
Next

For f = 1 To Form1.Text7.Text
Form1.Check7(f).Enabled = False
Next

For f = 1 To Form1.Text8.Text
Form1.Check8(f).Enabled = False
Next

For f = 1 To Form1.Text9.Text
Form1.Check9(f).Enabled = False
Next

For f = 1 To Form1.Text10.Text
Form1.Check10(f).Enabled = False
Next




End Sub
      
Private Sub Form_Unload(Cancel As Integer)



Form1.Combo2.Enabled = True
Form1.Check11.Enabled = True
Form1.Check12.Enabled = True
Form1.Check13.Enabled = True
Form1.Check14.Enabled = True
Form1.Check15.Enabled = True
Form1.Check16.Enabled = True
Form1.Check17.Enabled = True
Form1.Check18.Enabled = True
Form1.Check19.Enabled = True
Form1.Check20.Enabled = True

For f = 1 To Form1.Text1.Text
Form1.Check1(f).Enabled = True
Next

For f = 1 To Form1.Text2.Text
Form1.Check2(f).Enabled = True
Next

For f = 1 To Form1.Text3.Text
Form1.Check3(f).Enabled = True
Next

For f = 1 To Form1.Text4.Text
Form1.Check4(f).Enabled = True
Next

For f = 1 To Form1.Text5.Text
Form1.Check5(f).Enabled = True
Next

For f = 1 To Form1.Text6.Text
Form1.Check6(f).Enabled = True
Next

For f = 1 To Form1.Text7.Text
Form1.Check7(f).Enabled = True
Next

For f = 1 To Form1.Text8.Text
Form1.Check8(f).Enabled = True
Next

For f = 1 To Form1.Text9.Text
Form1.Check9(f).Enabled = True
Next

For f = 1 To Form1.Text10.Text
Form1.Check10(f).Enabled = True
Next

If Form1.Visible = True Then
Form1.SetFocus
Else
Unload Form4
End If


End Sub

Private Sub HScroll1_Change()
 'Call Aplicar_Transparencia(Me.hWnd, CByte(HScroll1.Value))
  
    End Sub

Private Sub Label2_Click()
Form_Click
End Sub
