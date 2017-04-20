Attribute VB_Name = "seleccionar_anterior"
 Option Explicit
      
    Const APPLICATION_d As String = "Instalador"
      
    Dim aplicacii As Single
    Dim multimedii As Single
    Dim grabacii As Single
    Dim segii As Single
    Dim menii As Single
    Dim conii As Single
    Dim ofimii As Single
    Dim twekii As Single
    Dim navii As Single
    Dim extii As Single
    
     
 Dim Path_Archivo_Ini_d As String
      
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
      
      
    'Lee un dato _
    -----------------------------
    'Recibe la ruta del archivo, la clave a leer y _
     el valor por defecto en caso de que la Key no exista
    Private Function Leer_Ini_d(Path_INI_d As String, Key_d As String, Default_d As Variant) As String
      
    Dim bufer_d As String * 256
    Dim Len_Value_d As Long
      
            Len_Value_d = GetPrivateProfileString(APPLICATION_d, _
                                             Key_d, _
                                             Default_d, _
                                             bufer_d, _
                                             Len(bufer_d), _
                                             Path_INI_d)
              
            Leer_Ini_d = Left$(bufer_d, Len_Value_d)
      
    End Function
      
    'Escribe un dato en el INI _
    -----------------------------
    'Recibe la ruta del archivo, La clave a escribir y el valor a añadir en dicha clave
      
    Private Function Grabar_Ini_d(Path_INI_d As String, Key_d As String, Valor_d As Variant) As String
      
        WritePrivateProfileString APPLICATION_d, _
                                             Key_d, _
                                             Valor_d, _
                                             Path_INI_d
      
    End Function


Public Sub seleanterior()
Dim aa As Integer
Dim bb As Integer
Dim cc As Integer
Dim dd As Integer
Dim ee As Integer
Dim ff As Integer
Dim gg As Integer
Dim hh As Integer
Dim ii As Integer
Dim jj As Integer
'------------------------
Path_Archivo_Ini_d = Environ("TEMP") & "\Instalador.ini"
'---------------------------
For aa = 1 To Form1.Text1.Text
aplicacii = Leer_Ini_d(Path_Archivo_Ini_d, Form1.apli_nom(aa).Caption, 0)

'--------0= no seleccionado
If aplicacii = 0 Then
Form1.Check1(aa).Value = 0
End If
'--------2= Se ejecuto pero se quedo en espera
If aplicacii = 2 Then
Form1.Check1(aa).Value = 0
End If
'--------3= se instalo correctamente o cerro la aplicacion
If aplicacii = 3 Then
Form1.Check1(aa).Value = 0
End If
'--------0= se selecciono pero no instalo
If aplicacii = 1 Then
Form1.Check1(aa).Value = 1
'Call Grabar_Ini_d(Path_Archivo_Ini_d, "seleccionado", 1)
End If
'--------4= Error de algun tipo como archivo no encontrado
If aplicacii = 4 Then
Form1.Check1(aa).Value = 1
End If
Next
'+++++++++++++++++++++++++++++++++++++++++++++++++++++
For bb = 1 To Form1.Text2.Text
multimedii = Leer_Ini_d(Path_Archivo_Ini_d, Form1.mul_nom(bb).Caption, 0)

If multimedii = 0 Then
Form1.Check2(bb).Value = 0
End If

If multimedii = 3 Then
Form1.Check2(bb).Value = 0
End If

If multimedii = 2 Then
Form1.Check2(bb).Value = 0
End If

If multimedii = 1 Then
Form1.Check2(bb).Value = 1
End If

If multimedii = 4 Then
Form1.Check2(bb).Value = 1
End If
Next
'+++++++++++++++++++++++++++++++++++++++++++++++++++
For cc = 1 To Form1.Text3.Text
grabacii = Leer_Ini_d(Path_Archivo_Ini_d, Form1.grab_nom(cc).Caption, 0)

If grabacii = 0 Then
Form1.Check3(cc).Value = 0
End If

If grabacii = 2 Then
Form1.Check3(cc).Value = 0
End If

If grabacii = 3 Then
Form1.Check3(cc).Value = 0
End If


If grabacii = 1 Then
Form1.Check3(cc).Value = 1
End If

If grabacii = 4 Then
Form1.Check3(cc).Value = 1
End If
Next
'++++++++++++++++++++++++++++++++++++++++++++++
For dd = 1 To Form1.Text4.Text
segii = Leer_Ini_d(Path_Archivo_Ini_d, Form1.seg_nom(dd).Caption, 0)

If segii = 0 Then
Form1.Check4(dd).Value = 0
End If

If segii = 2 Then
Form1.Check4(dd).Value = 0
End If

If segii = 3 Then
Form1.Check4(dd).Value = 0
End If

If segii = 1 Then
Form1.Check4(dd).Value = 1
End If

If segii = 4 Then
Form1.Check4(dd).Value = 1
End If
Next
'++++++++++++++++++++++++++++++++++++++++++++
For ee = 1 To Form1.Text5.Text
menii = Leer_Ini_d(Path_Archivo_Ini_d, Form1.men_nom(ee).Caption, 0)

If menii = 0 Then
Form1.Check5(ee).Value = 0
End If

If menii = 2 Then
Form1.Check5(ee).Value = 0
End If

If menii = 3 Then
Form1.Check5(ee).Value = 0
End If

If menii = 1 Then
Form1.Check5(ee).Value = 1
End If

If menii = 4 Then
Form1.Check5(ee).Value = 1
End If
Next
'+++++++++++++++++++++++++++++++++++++++++++++
For ff = 1 To Form1.Text6.Text
conii = Leer_Ini_d(Path_Archivo_Ini_d, Form1.con_nom(ff).Caption, 0)

If conii = 0 Then
Form1.Check6(ff).Value = 0
End If

If conii = 2 Then
Form1.Check6(ff).Value = 0
End If

If conii = 3 Then
Form1.Check6(ff).Value = 0
End If

If conii = 1 Then
Form1.Check6(ff).Value = 1
End If

If conii = 4 Then
Form1.Check6(ff).Value = 1
End If
Next
'++++++++++++++++++++++++++++++++++++++++++++++
For gg = 1 To Form1.Text7.Text
ofimii = Leer_Ini_d(Path_Archivo_Ini_d, Form1.ofim_nom(gg).Caption, 0)

If ofimii = 0 Then
Form1.Check7(gg).Value = 0
End If

If ofimii = 2 Then
Form1.Check7(gg).Value = 0
End If

If ofimii = 3 Then
Form1.Check7(gg).Value = 0
End If

If ofimii = 1 Then
Form1.Check7(gg).Value = 1
End If

If ofimii = 4 Then
Form1.Check7(gg).Value = 1
End If
Next
'++++++++++++++++++++++++++++++++++++++++++++++++
For hh = 1 To Form1.Text8.Text
twekii = Leer_Ini_d(Path_Archivo_Ini_d, Form1.twek_nom(hh).Caption, 0)

If twekii = 0 Then
Form1.Check8(hh).Value = 0
End If

If twekii = 2 Then
Form1.Check8(hh).Value = 0
End If

If twekii = 3 Then
Form1.Check8(hh).Value = 0
End If

If twekii = 1 Then
Form1.Check8(hh).Value = 1
End If

If twekii = 4 Then
Form1.Check8(hh).Value = 1
End If
Next
'+++++++++++++++++++++++++++++++++++++++++++++++++++
For ii = 1 To Form1.Text9.Text
navii = Leer_Ini_d(Path_Archivo_Ini_d, Form1.nav_nom(ii).Caption, 0)

If navii = 0 Then
Form1.Check9(ii).Value = 0
End If

If navii = 2 Then
Form1.Check9(ii).Value = 0
End If

If navii = 3 Then
Form1.Check9(ii).Value = 0
End If

If navii = 1 Then
Form1.Check9(ii).Value = 1
End If

If navii = 4 Then
Form1.Check9(ii).Value = 1
End If
Next
'+++++++++++++++++++++++++++++++++++++++++++++++++++++
For jj = 1 To Form1.Text10.Text
extii = Leer_Ini_d(Path_Archivo_Ini_d, Form1.ext_nom(jj).Caption, 0)

If extii = 0 Then
Form1.Check10(jj).Value = 0
End If

If extii = 2 Then
Form1.Check10(jj).Value = 0
End If

If extii = 3 Then
Form1.Check10(jj).Value = 0
End If

If extii = 1 Then
Form1.Check10(jj).Value = 1
End If

If extii = 4 Then
Form1.Check10(jj).Value = 1
End If
Next
End Sub
