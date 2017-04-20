Attribute VB_Name = "imagensd"
 Option Explicit
'Variable para la imagen
Private Pic As IPictureDisp
'Variables para la carpeta de imagenes _
y para el campo que tiene el Id de imagen
Public Carpeta_IMG As String
Public Field_Img As String

'Subrutina que dibuja el gráfico en el control Picture _
 en forma centrada y a escala
 '*******************************************************
Dim hNew2 As Long
Public Const IMAGE_BITMAP = 0
Public Const LR_COPYRETURNORG = &H4
Public Const CF_BITMAP = 2
Public Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Public Declare Function CopyImage Lib "user32" (ByVal handle As Long, ByVal imageType As Long, ByVal newWidth As Long, ByVal newHeight As Long, ByVal lFlags As Long) As Long
Public Declare Function EmptyClipboard Lib "user32" () As Long
Public Declare Function CloseClipboard Lib "user32" () As Long
Public Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long

'--------------------------





Sub Dibujar_Imagen(Objeto As Object, Path_Imagen As String)

On Error GoTo ErrSub

Dim Pos_x As Single
Dim Pos_y As Single
Dim Ancho_IMG As Single
Dim Alto_IMG As Single
Dim Ancho_Obj As Single
Dim Alto_Obj As Single
Dim Old_Scale As Single


    Set Pic = LoadPicture(Path_Imagen)

    With Objeto
    
    .AutoRedraw = True
    .Cls
    
    Old_Scale = .ScaleMode
    
    .ScaleMode = vbPixels
    Ancho_IMG = .ScaleX(Pic.Width, vbHimetric, vbPixels)
    Alto_IMG = .ScaleY(Pic.Height, vbHimetric, vbPixels)
    
    Ancho_Obj = .ScaleWidth
    Alto_Obj = .ScaleHeight
    
    If Ancho_IMG > Ancho_Obj Then
        Alto_IMG = Alto_IMG * Ancho_Obj / Ancho_IMG
        Ancho_IMG = Ancho_Obj
    End If
    If Alto_IMG > Alto_Obj Then
        Ancho_IMG = Ancho_IMG * Alto_Obj / Alto_IMG
        Alto_IMG = Alto_Obj
    End If
    Pos_x = (Ancho_Obj - Ancho_IMG) / 2
    Pos_y = (Alto_Obj - Alto_IMG) / 2
    
    End With
    

    Objeto.PaintPicture Pic, Pos_x, Pos_y, Ancho_IMG, Alto_IMG
    
    Objeto.ScaleMode = Old_Scale
    
    Exit Sub
    
'Error
ErrSub:
    
    If Err.Number = 76 Then
       Objeto.Cls
       Exit Sub
    Else
       MsgBox Err.Description, vbCritical
    End If
End Sub





