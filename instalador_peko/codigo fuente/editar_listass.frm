VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form detilar_lit 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Editar Programas"
   ClientHeight    =   5685
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9225
   Icon            =   "editar_listass.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   379
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   615
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   5400
      TabIndex        =   27
      Text            =   "Text1"
      Top             =   120
      Width           =   3495
   End
   Begin MSComDlg.CommonDialog CommonDialog2 
      Left            =   4440
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   3480
      Top             =   120
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   26
      Top             =   120
      Width           =   2655
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Programa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2415
      Left            =   120
      TabIndex        =   7
      Top             =   3120
      Width           =   9015
      Begin VB.CommandButton cmdNav 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Caption         =   ">>"
         Height          =   375
         Index           =   3
         Left            =   5160
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   1950
         Width           =   615
      End
      Begin VB.CommandButton cmdNav 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Caption         =   ">"
         Height          =   375
         Index           =   2
         Left            =   4560
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   1950
         Width           =   615
      End
      Begin VB.CommandButton cmdNav 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Caption         =   "<"
         Height          =   375
         Index           =   1
         Left            =   3960
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   1950
         Width           =   615
      End
      Begin VB.CommandButton cmdNav 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Caption         =   "<<"
         Height          =   375
         Index           =   0
         Left            =   3360
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   1950
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Nuevo"
         Height          =   375
         Index           =   0
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   1560
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Eliminar"
         Height          =   375
         Index           =   1
         Left            =   3120
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   1560
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Actualizar"
         Height          =   375
         Index           =   2
         Left            =   4080
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   1560
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Editar"
         Height          =   375
         Index           =   3
         Left            =   5040
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   1560
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Cancelar"
         Height          =   375
         Index           =   4
         Left            =   6000
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   1560
         Width           =   975
      End
      Begin VB.CommandButton Command3 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   375
         Left            =   7560
         TabIndex        =   13
         Top             =   600
         Width           =   495
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Limpiar"
         Height          =   375
         Left            =   7440
         TabIndex        =   12
         Top             =   1080
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txt_Field 
         Height          =   375
         Index           =   3
         Left            =   1440
         TabIndex        =   10
         Top             =   1080
         Width           =   5895
      End
      Begin VB.TextBox txt_Field 
         Height          =   375
         Index           =   2
         Left            =   1440
         TabIndex        =   8
         Top             =   600
         Width           =   5895
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Comandos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   4
         Left            =   120
         TabIndex        =   11
         Top             =   1080
         Width           =   1155
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Direccion"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   2
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   990
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Datos de Aplicación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   9015
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   7440
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   16
         Top             =   960
         Width           =   495
      End
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Caption         =   "Selecionar Imagen"
         Height          =   495
         Index           =   5
         Left            =   6360
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   1920
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Caption         =   "Eliminar Imagen"
         Height          =   495
         Index           =   6
         Left            =   7680
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   1920
         Width           =   1215
      End
      Begin VB.TextBox txt_Field 
         Height          =   375
         Index           =   0
         Left            =   1800
         TabIndex        =   2
         Top             =   960
         Width           =   4335
      End
      Begin VB.TextBox txt_Field 
         Height          =   375
         Index           =   1
         Left            =   1800
         TabIndex        =   1
         Top             =   1440
         Width           =   4335
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Id"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   3
         Left            =   360
         TabIndex        =   6
         Top             =   600
         Width           =   210
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descripción"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   360
         TabIndex        =   5
         Top             =   1440
         Width           =   1245
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   360
         TabIndex        =   4
         Top             =   960
         Width           =   840
      End
      Begin VB.Label lblID 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1800
         TabIndex        =   3
         Top             =   480
         Width           =   75
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   11400
      Top             =   3600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "detilar_lit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim cn As Connection
Dim rst As Recordset

' Primer registro, siguiente, etc...
Private Sub cmdNav_Click(Index As Integer)

    ' Si hay registro activo sale
    If rst.BOF And rst.EOF Then Exit Sub

    Select Case Index

    Case 0
        rst.MoveFirst
    Case 1
        rst.MovePrevious
        If rst.BOF Then rst.MoveFirst
    Case 2
        rst.MoveNext
        If rst.EOF Then rst.MoveLast
    Case 3
        rst.MoveLast

    End Select

    ' Carga la imagen en el Picture
    Mostrar_Imagen

End Sub

Private Sub Command1_Click(Index As Integer)

    Select Case Index
 
        'Agrega un nuevo registro
        Case 0
            rst.AddNew
            Picture1.Cls
            'Elimina el registro activo
            
            CmdNuevo
            Command3.Enabled = True
            
        Case 1
            If rst.EOF Or rst.BOF Then Exit Sub
            If MsgBox("Eliminar Registro", vbQuestion + vbYesNo) = vbNo Then Exit Sub
            
            Picture1.Cls
    
            'Elimina el archivo de la carpeta de imagenes
            If rst(Field_Img) <> "" Then
                Call Kill(Carpeta_IMG & rst(Field_Img))
            End If
        
            rst.Delete
            
            If rst.RecordCount > 0 Then
               cmdNormal
            Else
               cmdSinRegistros
            End If
            
            If rst.EOF Or rst.BOF Then
                Exit Sub
            End If
            rst.MoveNext
            
            If rst.EOF Then
               On Error Resume Next
               rst.MoveLast
            End If
            'Carga la imagen del registro activo
            Mostrar_Imagen
            Exit Sub
             
        ' Botón Actualizar los cambios en la base de datos
        Case 2
            If Not rst.EOF And Not rst.BOF Then
                rst.Update
                Guardar_Imagen
                cmdNormal
            End If

        ' Cancela la atualización o edición del registro que se editando o añadiendo
        Case 3
        Command3.Enabled = True
            cmdEditar
            Setear_TextBox
            Exit Sub
  
        'Botón Editar el registro activo
        Case 4
            
            If rst.EOF And rst.BOF Then Exit Sub
            rst.CancelUpdate
  
            If Not rst.BOF And Not rst.EOF Then
                If rst(Field_Img) <> "" Then
                    Call Dibujar_Imagen(Picture1, Carpeta_IMG & rst(Field_Img))
                End If
                
            End If
            
            If rst.RecordCount > 0 Then
                cmdNormal
            Else
                cmdSinRegistros
            End If
        'Carga una imagen en el control Picture1
        Case 5
  
            With CommonDialog1
                .DialogTitle = " Seleccionar imagen"
                .Filter = "BMP|*.bmp|JPEG|*.jpeg|GIF|*.gif|JPG|*.jpg|Todos|*.*"
     
                .ShowOpen
     
                If .FileName = "" Then
                    Exit Sub
                Else
         
         
         
         
         '-------------------------------------------
         
                    ' Graba el nombre en el campo, el id de imagen _
                    que es el mismo que el campo Id
         
                    rst(Field_Img) = rst!nombre '
         
        
                    ' se dibuja la imagen en el Picture
                    Call Dibujar_Imagen(Picture1, .FileName)
         
                End If
            End With
            
            Exit Sub

        Case 6

            ' Limpia la imagen del Picture y Elimina el id de _
            imagen del registro actual de la base
            
            If MsgBox("Desea eliminar la imagen ?", vbYesNo + vbQuestion) = vbYes Then
               Picture1.Cls
               rst(Field_Img) = ""
               Exit Sub
            End If

    End Select

    
    Setear_TextBox

    ' Muestra la imagen
    Mostrar_Imagen

End Sub

Sub Guardar_Imagen()

Dim hNew2 As Long
    ' Si el campo Id_Imagen no está vacio ...
    If rst(Field_Img) <> "" And CommonDialog1.FileName <> "" Then
        
'-------------cambiar tamaño a 32*32
Picture1.Picture = LoadPicture(CommonDialog1.FileName)

hNew2 = CopyImage(Picture1.Picture, IMAGE_BITMAP, Val(32), Val(32), LR_COPYRETURNORG)
OpenClipboard Me.hwnd
EmptyClipboard
SetClipboardData CF_BITMAP, hNew2
CloseClipboard

Picture1.Picture = Clipboard.GetData(2)
'SavePicture Picture1.Picture, "C:\Users\user\Desktop\foto2.bmp"
SavePicture Picture1.Picture, Carpeta_IMG & "\" & rst!nombre

 '-----------------------------------
 
 ' Copia el archivo a la carpeta de imagen
       ' Call FileCopy(CommonDialog1.FileName, _
                      Carpeta_IMG & "\" & rst!nombre)

        
        
        '... si no, si el archivo está en la carpeta lo  elimina
    ElseIf Dir(Carpeta_IMG & "\" & rst!nombre) <> "" And rst(Field_Img) = "" Then
       Call Kill(Carpeta_IMG & rst!nombre)








    End If
End Sub


Private Sub Mostrar_Imagen()
    With rst
        ' Si no hay ningún registro activo sale
        If .EOF Or .BOF Then
            Exit Sub
        End If
        
        ' Si el registro no tiene una imagen asociada Limpia el Picture
        If .Fields(Field_Img) = "" Or .Fields(Field_Img) = 0 Then
           Picture1.Cls
        Else
           ' Lee el archivo de imagen y lo dibuja en el Picture
            Call Dibujar_Imagen(Picture1, Carpeta_IMG & .Fields(Field_Img))
        End If

    End With

Exit Sub





End Sub

Private Sub Setear_TextBox()
    'Bloquea y desbloquea los textbox
    Dim T As TextBox
    For Each T In Me.txt_Field
        T.Locked = Not T.Locked
    Next
End Sub

' Habilita y deshabilita los CommandButton

Private Sub Setear_botones()

    Dim i As Integer

    For i = 0 To Command1.Count - 1
        Command1(i).Enabled = Not Command1(i).Enabled
    Next

    For i = 0 To cmdNav.Count - 1
        cmdNav(i).Enabled = Not cmdNav(i).Enabled
    Next

End Sub


Private Sub Command3_Click()
Dim apli_dir As String
apli_dir = App.Path & "\recursos\aplicaciones\"


 If Dir(App.Path & "\recursos\aplicaciones\", vbDirectory) = "" Then
        MkDir App.Path & "\recursos\aplicaciones\"
    End If
    

 With CommonDialog2
                .DialogTitle = " Seleccionar Aplicación"
                .Filter = "Archivos ejecutables    *.exe|*.exe|Microsoft Windows Installer  *.msi|*.msi|Archivo batch     *.bat|*.bat|Archivo por lotes     *.cmd|*.cmd"
     
                .ShowOpen
     
                If .FileName = "" Then
                    Exit Sub
                Else
                
                ' Copia el archivo a la carpeta de imagen
        Call FileCopy(CommonDialog2.FileName, App.Path & "\recursos\aplicaciones\" & CommonDialog2.FileTitle)
       ' rst!nombre & ".exe")
                Text1.Text = apli_dir
        txt_Field(2).Text = CommonDialog2.FileTitle

                End If
End With




End Sub



Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next

    If Not rst Is Nothing Then
        If rst.State = adStateOpen Then rst.Close
        Set rst = Nothing
    End If
    If Not cn Is Nothing Then
        If cn.State = adStateOpen Then cn.Close
        Set cn = Nothing
    End If
    
End Sub

Private Sub Form_Load()

    Dim Pathbd As String, cadena As String
    Dim T As TextBox
    
    Set cn = New Connection

    Pathbd = App.Path & "\recursos\datos.mdb"

    cadena = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Pathbd & _
                                     ";Persist Security Info=False"

    
    cn.Open cadena

    Set rst = New Recordset

Set rst = cn.Execute("SELECT * from categorias")
    
    While rst.EOF = False
            Combo1.AddItem (rst.Fields("categorias"))
            rst.MoveNext

        Wend
rst.MoveFirst
Combo1.Text = rst!categorias
    

End Sub


Sub cmdNormal()

    DeshabilitarTodosCmd

    Command1(0).Enabled = True
    Command1(1).Enabled = True
    Command1(3).Enabled = True
    
End Sub

Sub cmdSinRegistros()

    DeshabilitarTodosCmd
    Command1(0).Enabled = True

End Sub

Sub cmdEditar()
        
    DeshabilitarTodosCmd
    Command1(2).Enabled = True
    Command1(4).Enabled = True
    Command1(5).Enabled = True
    Command1(6).Enabled = True
    
End Sub

Sub CmdNuevo()
    DeshabilitarTodosCmd
    Command1(2).Enabled = True
    Command1(4).Enabled = True
    
    Command1(5).Enabled = True
    Command1(6).Enabled = True
End Sub

Sub DeshabilitarTodosCmd()
    Command1(0).Enabled = False
    Command1(1).Enabled = False
    Command1(2).Enabled = False
    Command1(3).Enabled = False
    Command1(4).Enabled = False
    Command1(5).Enabled = False
    Command1(6).Enabled = False
    
End Sub

Private Sub combo1_click()
rst.Close
    
    rst.Open "Select * FROM " & Combo1.Text, cn, adOpenStatic, adLockOptimistic

    ' Nombre del campo  que tiene el ID de imagen
    Field_Img = "nombre"
    ' Path de la carpeta donde están las imagenes
    Carpeta_IMG = App.Path & "\recursos\logos\"

    ' Si no existe la carpeta para guardar las imagen la crea
    If Dir(App.Path & "\recursos\logos\", vbDirectory) = "" Then
        MkDir App.Path & "\recursos\logos\"
    End If
    
    If rst.RecordCount > 0 Then
        Call cmdNormal
    Else
        Call cmdSinRegistros
    End If
    
    Set txt_Field(0).DataSource = rst
    Set txt_Field(1).DataSource = rst
    Set txt_Field(2).DataSource = rst
    Set txt_Field(3).DataSource = rst
    
    txt_Field(0).DataField = "Nombre"
    txt_Field(1).DataField = "Descripción"
    txt_Field(2).DataField = "path"
    txt_Field(3).DataField = "comandos"

    'Opcional: esto visualiza el Id del registro en un label
    Set lblID.DataSource = rst
    lblID.DataField = "Id"

    Call Setear_TextBox

    ' carga la imagen en el registro si es que tiene
    Call Mostrar_Imagen

End Sub
