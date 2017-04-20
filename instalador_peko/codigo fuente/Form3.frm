VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Form3"
   ClientHeight    =   1710
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4410
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1710
   ScaleWidth      =   4410
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Shape Shape2 
      BackColor       =   &H000000C0&
      BackStyle       =   1  'Opaque
      DrawMode        =   9  'Not Mask Pen
      FillStyle       =   4  'Upward Diagonal
      Height          =   300
      Left            =   120
      Top             =   1200
      Width           =   4170
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H000000C0&
      BackStyle       =   1  'Opaque
      DrawMode        =   3  'Not Merge Pen
      Height          =   300
      Left            =   120
      Top             =   360
      Width           =   4170
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H000000C0&
      BackStyle       =   1  'Opaque
      DrawMode        =   5  'Not Copy Pen
      Height          =   300
      Left            =   120
      Top             =   720
      Width           =   4170
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000C0&
      BorderWidth     =   2
      Height          =   1455
      Left            =   720
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
Shape1.Left = 10
Shape1.Top = 10
Shape1.Width = Form3.Width - 10
Shape1.Height = Form3.Height - 10
Shape1.BorderWidth = 2
End Sub
