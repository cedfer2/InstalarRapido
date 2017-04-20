VERSION 5.00
Begin VB.Form proceso 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Instalar Complementos"
   ClientHeight    =   675
   ClientLeft      =   -15
   ClientTop       =   210
   ClientWidth     =   690
   Icon            =   "procesos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   675
   ScaleWidth      =   690
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   120
      Top             =   120
   End
End
Attribute VB_Name = "proceso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fso As Object

Private Sub Form_Load()
Dim mi_dir As String
Dim pg As Long

mi_dir = "\recursos\registrar_componentes\registrar_componentes.exe"
Set fso = CreateObject("Scripting.FileSystemObject")
If fso.FileExists(App.Path & mi_dir) = True Then
Shell (App.Path & mi_dir)
Timer1.Enabled = True
Else
MsgBox "No se puede registrar componentes", vbInformation + vbOKOnly, "Información"
Timer1.Interval = 1
Timer1.Enabled = True
End If




End Sub

Private Sub Timer1_Timer()
On Err GoTo g

Dim mi_dirh As String

mi_dirh = "\instalador.exe"
Set fso = CreateObject("Scripting.FileSystemObject")
If fso.FileExists(App.Path & mi_dirh) = True Then
Shell (App.Path & mi_dirh)
End
Else
MsgBox "No se puede ejecutar aplicación", vbInformation + vbOKOnly, "Información"
End
End If
Exit Sub

g:
End
End Sub
