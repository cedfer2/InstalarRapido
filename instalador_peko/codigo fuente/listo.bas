Attribute VB_Name = "listo"
     Option Explicit
      
    Const APPLICATION_j As String = "Instalador"
      
    Dim seleccionado As Single
    Dim mul_nombre As Single
    Dim mul_descripcion As Single
    Dim mul_path As Single
    Dim mul_estado As Single
      
    Dim Path_Archivo_Ini_j As String
      
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
    Private Function Leer_Ini_j(Path_INI_j As String, Key_j As String, Default_j As Variant) As String
      
    Dim bufer_j As String * 256
    Dim Len_Value_j As Long
      
            Len_Value_j = GetPrivateProfileString(APPLICATION_j, _
                                             Key_j, _
                                             Default_j, _
                                             bufer_j, _
                                             Len(bufer_j), _
                                             Path_INI_j)
              
            Leer_Ini_j = Left$(bufer_j, Len_Value_j)
      
    End Function
      
    'Escribe un dato en el INI _
    -----------------------------
    'Recibe la ruta del archivo, La clave a escribir y el valor a añadir en dicha clave
      
    Private Function Grabar_Ini_j(Path_INI_j As String, Key_j As String, Valor_j As Variant) As String
      
        WritePrivateProfileString APPLICATION_j, _
                                             Key_j, _
                                             Valor_j, _
                                             Path_INI_j
      
    End Function



'------------------
Public Sub seleccionar_programas()
Dim ruta_def As String
ruta_def = App.path & "\recursos\aplicaciones\"
Call Desconectar
Call IniciarConexion
Dim total_de_aplis1 As Integer
Dim total_de_aplis2 As Integer
Dim total_de_aplis3 As Integer
Dim total_de_aplis4 As Integer
Dim total_de_aplis5 As Integer
Dim total_de_aplis6 As Integer
Dim total_de_aplis7 As Integer
Dim total_de_aplis8 As Integer
Dim total_de_aplis9 As Integer
Dim total_de_aplis10 As Integer
Dim contador_de_ejecut As Integer

Dim aaa As Integer
Dim bbb As Integer
Dim ccc As Integer
Dim ddd As Integer
Dim eee As Integer
Dim fff As Integer
Dim ggg As Integer
Dim hhh As Integer
Dim iii As Integer
Dim jjj As Integer



Dim ui As Integer
Dim ruta2 As String
Path_Archivo_Ini_j = Environ("TEMP") & "\Instalador.ini"
Dim i As Integer




total_de_aplis1 = 0
total_de_aplis2 = 0
total_de_aplis3 = 0
total_de_aplis4 = 0
total_de_aplis5 = 0
total_de_aplis6 = 0
total_de_aplis7 = 0
total_de_aplis8 = 0
total_de_aplis9 = 0
total_de_aplis10 = 0

On Error GoTo nananna
'------contador de selecionados
'Form2.CommandButton1.Visible = True
Call Grabar_Ini_j(Path_Archivo_Ini_j, "seleccionado", 0)
For i = 1 To Form1.Text1.Text
Call Grabar_Ini_j(Path_Archivo_Ini_j, Form1.apli_nom(i).Caption, 0)
If Form1.Check1(i).Value = 1 Then
Call Grabar_Ini_j(Path_Archivo_Ini_j, Form1.apli_nom(i).Caption, 1)
Call Grabar_Ini_j(Path_Archivo_Ini_j, "seleccionado", 1)
total_de_aplis1 = total_de_aplis1 + 1
End If
Next
Form2.Label7.Caption = total_de_aplis1
'---------------------------------------

For i = 1 To Form1.Text2.Text
Call Grabar_Ini_j(Path_Archivo_Ini_j, Form1.mul_nom(i).Caption, 0)
If Form1.Check2(i).Value = 1 Then
Call Grabar_Ini_j(Path_Archivo_Ini_j, Form1.mul_nom(i).Caption, 1)
Call Grabar_Ini_j(Path_Archivo_Ini_j, "seleccionado", 1)
total_de_aplis2 = total_de_aplis2 + 1
End If
Next
Form2.Label8.Caption = total_de_aplis2
'---------------------------------------

For i = 1 To Form1.Text3.Text
Call Grabar_Ini_j(Path_Archivo_Ini_j, Form1.grab_nom(i).Caption, 0)
If Form1.Check3(i).Value = 1 Then
Call Grabar_Ini_j(Path_Archivo_Ini_j, Form1.grab_nom(i).Caption, 1)
Call Grabar_Ini_j(Path_Archivo_Ini_j, "seleccionado", 1)
total_de_aplis3 = total_de_aplis3 + 1
End If
Next
Form2.Label9.Caption = total_de_aplis3
'---------------------------------------

For i = 1 To Form1.Text4.Text
Call Grabar_Ini_j(Path_Archivo_Ini_j, Form1.seg_nom(i).Caption, 0)
If Form1.Check4(ddd).Value = 1 Then
Call Grabar_Ini_j(Path_Archivo_Ini_j, Form1.seg_nom(i).Caption, 1)
Call Grabar_Ini_j(Path_Archivo_Ini_j, "seleccionado", 1)
total_de_aplis4 = total_de_aplis4 + 1
End If
Next
Form2.Label10.Caption = total_de_aplis4
'---------------------------------------

For i = 1 To Form1.Text5.Text
Call Grabar_Ini_j(Path_Archivo_Ini_j, Form1.men_nom(i).Caption, 0)
If Form1.Check5(i).Value = 1 Then
Call Grabar_Ini_j(Path_Archivo_Ini_j, Form1.men_nom(i).Caption, 1)
Call Grabar_Ini_j(Path_Archivo_Ini_j, "seleccionado", 1)
total_de_aplis5 = total_de_aplis5 + 1
End If
Next
Form2.Label11.Caption = total_de_aplis5
'---------------------------------------

For i = 1 To Form1.Text6.Text
Call Grabar_Ini_j(Path_Archivo_Ini_j, Form1.con_nom(i).Caption, 0)
If Form1.Check6(i).Value = 1 Then
Call Grabar_Ini_j(Path_Archivo_Ini_j, Form1.con_nom(i).Caption, 1)
Call Grabar_Ini_j(Path_Archivo_Ini_j, "seleccionado", 1)
total_de_aplis6 = total_de_aplis6 + 1
End If
Next
Form2.Label12.Caption = total_de_aplis6
'---------------------------------------

For i = 1 To Form1.Text7.Text
Call Grabar_Ini_j(Path_Archivo_Ini_j, Form1.ofim_nom(i).Caption, 0)
If Form1.Check7(i).Value = 1 Then
Call Grabar_Ini_j(Path_Archivo_Ini_j, Form1.ofim_nom(i).Caption, 1)
Call Grabar_Ini_j(Path_Archivo_Ini_j, "seleccionado", 1)
total_de_aplis7 = total_de_aplis7 + 1
End If
Next
Form2.Label13.Caption = total_de_aplis7
'---------------------------------------

For i = 1 To Form1.Text8.Text
Call Grabar_Ini_j(Path_Archivo_Ini_j, Form1.twek_nom(i).Caption, 0)
If Form1.Check8(hhh).Value = 1 Then
Call Grabar_Ini_j(Path_Archivo_Ini_j, Form1.twek_nom(i).Caption, 1)
Call Grabar_Ini_j(Path_Archivo_Ini_j, "seleccionado", 1)
total_de_aplis8 = total_de_aplis8 + 1
End If
Next
Form2.Label14.Caption = total_de_aplis8
'---------------------------------------

For i = 1 To Form1.Text9.Text
Call Grabar_Ini_j(Path_Archivo_Ini_j, Form1.nav_nom(i).Caption, 0)
If Form1.Check9(i).Value = 1 Then
Call Grabar_Ini_j(Path_Archivo_Ini_j, Form1.nav_nom(i).Caption, 1)
Call Grabar_Ini_j(Path_Archivo_Ini_j, "seleccionado", 1)
total_de_aplis9 = total_de_aplis9 + 1
End If
Next
Form2.Label15.Caption = total_de_aplis9
'---------------------------------------

For i = 1 To Form1.Text10.Text
Call Grabar_Ini_j(Path_Archivo_Ini_j, Form1.ext_nom(i).Caption, 0)
If Form1.Check10(i).Value = 1 Then
Call Grabar_Ini_j(Path_Archivo_Ini_j, Form1.ext_nom(i).Caption, 1)
Call Grabar_Ini_j(Path_Archivo_Ini_j, "seleccionado", 1)
total_de_aplis10 = total_de_aplis10 + 1
End If
Next
Form2.Label16.Caption = total_de_aplis10
'--------------
Dim Resph As Integer
Dim Resphi As Integer

seleccionado = Leer_Ini_j(Path_Archivo_Ini_j, "seleccionado", 1)

If seleccionado = 0 Then
Form1.Visible = False
Kill Environ("TEMP") & "\Instalador.ini"
Form2.Label5.Caption = "Ninguna aplicación fué seleccionada"
Resph = MsgBox("No ha selecionado ninguna opción", vbInformation + vbOKOnly, "Información")
Resphi = MsgBox("¿Desea Salir?", vbQuestion + vbYesNo + vbDefaultButton2, "Advertencia")
If Resphi = 6 Then
End
End If

End If
'--------------










Form2.Label17.Caption = val(Form2.Label7) + val(Form2.Label8) + val(Form2.Label9) + val(Form2.Label10) + val(Form2.Label11) + val(Form2.Label12) + val(Form2.Label13) + val(Form2.Label14) + val(Form2.Label15) + val(Form2.Label16)
'form2.Label19.Caption=

Form2.Label19.Caption = 425 '426


Form2.Label20.Caption = Form2.Label19.Caption / Form2.Label17.Caption
Form2.Label21.Caption = Form2.progressbar1.Width
Form2.Label22.Caption = Form2.Shape2.Width
Form2.Label23.Caption = Form2.Shape2.Width - Form2.progressbar1.Width
Form2.iniciado.Caption = 1
'---------------------------
contador_de_ejecut = 0

For i = 1 To Form1.Text1.Text
If Form1.Check1(i).Value = 1 Then
Form2.Text1.Text = Form2.Text1.Text & "----------------- [ Aplicaciones ] ----------------" + vbCrLf
ui = i
Set rs = cnn.Execute("SELECT * from aplicaciones where id=" & ui)
Form2.Caption = "Asistente de Instalación  -Instalando " & rs("nombre")
Form2.Text1.Text = Form2.Text1.Text & "Instalar " & rs("nombre") & vbCrLf & "    Comando... " & vbCrLf
Form2.Text2.Text = rs("path")
Call Ejecutar_shell(ruta_def & Form2.Text2.Text)
Call Grabar_Ini_j(Path_Archivo_Ini_j, rs("nombre"), Form2.Label25.Caption)
Form1.Check1(i).Value = 0
contador_de_ejecut = contador_de_ejecut + 1
Form2.Text4.Caption = contador_de_ejecut
Form2.progressbar1.Width = Form2.progressbar1.Width + Form2.Label20.Caption
End If
Next
'---------------------------------------

For i = 1 To Form1.Text2.Text
If Form1.Check2(i).Value = 1 Then
Form2.Text1.Text = Form2.Text1.Text & vbCrLf & "------------------ [ Multimedia ] ----------------" & vbCrLf & vbCrLf
ui = i
Set rs = cnn.Execute("SELECT * from multimedia where id=" & ui)
Form2.Caption = "Asistente de Instalación  -Instalando " & rs("nombre")
Form2.Text1.Text = Form2.Text1.Text & "Instalar " & rs("nombre") & vbCrLf & "    Comando... " & vbCrLf
Form2.Text2.Text = rs("path")
Call Ejecutar_shell(ruta_def & Form2.Text2.Text)
Call Grabar_Ini_j(Path_Archivo_Ini_j, rs("nombre"), Form2.Label25.Caption)

Form1.Check2(i).Value = 0
contador_de_ejecut = contador_de_ejecut + 1
Form2.Text4.Caption = contador_de_ejecut '- 1
Form2.progressbar1.Width = Form2.progressbar1.Width + Form2.Label20.Caption
End If
Next
'---------------------------------------

For i = 1 To Form1.Text3.Text
If Form1.Check3(i).Value = 1 Then
Form2.Text1.Text = Form2.Text1.Text & "------------------ [ Grabacion ] -----------------" & vbCrLf
ui = i
Set rs = cnn.Execute("SELECT * from Grabación_y_Backups where id=" & ui)
Form2.Text1.Text = Form2.Text1.Text & "Instalar " & rs("nombre") & vbCrLf & "    Comando... " & vbCrLf
Form2.Caption = "Asistente de Instalación  -Instalando " & rs("nombre")
Form2.Text2.Text = rs("path")
Call Ejecutar_shell(ruta_def & Form2.Text2.Text)
Call Grabar_Ini_j(Path_Archivo_Ini_j, rs("nombre"), Form2.Label25.Caption)
Form1.Check3(i).Value = 0
contador_de_ejecut = contador_de_ejecut + 1
Form2.Text4.Caption = contador_de_ejecut
Form2.progressbar1.Width = Form2.progressbar1.Width + Form2.Label20.Caption
End If
Next
'---------------------------------------



For i = 1 To Form1.Text4.Text
If Form1.Check4(i).Value = 1 Then
Form2.Text1.Text = Form2.Text1.Text & "------------------ [ Seguridad ] -----------------" & vbCrLf
ui = i
Set rs = cnn.Execute("SELECT * from Seguridad where id=" & ui)
Form2.Text1.Text = Form2.Text1.Text & "Instalar " & rs("nombre") & vbCrLf & "    Comando... " & vbCrLf
Form2.Caption = "Asistente de Instalación  -Instalando " & rs("nombre")
Form2.Text2.Text = rs("path")
Call Ejecutar_shell(ruta_def & Form2.Text2.Text)
Call Grabar_Ini_j(Path_Archivo_Ini_j, rs("nombre"), Form2.Label25.Caption)
Form1.Check4(i).Value = 0
contador_de_ejecut = contador_de_ejecut + 1
Form2.Text4.Caption = contador_de_ejecut
Form2.progressbar1.Width = Form2.progressbar1.Width + Form2.Label20.Caption
End If
Next
'---------------------------------------

For i = 1 To Form1.Text5.Text
If Form1.Check5(i).Value = 1 Then
Form2.Text1.Text = Form2.Text1.Text & "----------------- [ Mensajería ] ------------------" & vbCrLf
ui = i
Set rs = cnn.Execute("SELECT * from Mensajería where id=" & ui)
Form2.Text1.Text = Form2.Text1.Text & "Instalar " & rs("nombre") & vbCrLf & "    Comando... " & vbCrLf
Form2.Caption = "Asistente de Instalación  -Instalando " & rs("nombre")
Form2.Text2.Text = rs("path")
Call Ejecutar_shell(ruta_def & Form2.Text2.Text)
Call Grabar_Ini_j(Path_Archivo_Ini_j, rs("nombre"), Form2.Label25.Caption)
Form1.Check5(i).Value = 0
contador_de_ejecut = contador_de_ejecut + 1
Form2.Text4.Caption = contador_de_ejecut
Form2.progressbar1.Width = Form2.progressbar1.Width + Form2.Label20.Caption
End If
Next

'---------------------------------------
For i = 1 To Form1.Text6.Text
If Form1.Check6(i).Value = 1 Then
Form2.Text1.Text = Form2.Text1.Text & "----------------- [ Complemento ] ---------------" & vbCrLf
ui = i
Set rs = cnn.Execute("SELECT * from Complemento where id=" & ui)
Form2.Text1.Text = Form2.Text1.Text & "Instalar " & rs("nombre") & vbCrLf & "    Comando... " & vbCrLf
Form2.Caption = "Asistente de Instalación  -Instalando " & rs("nombre")
Form2.Text2.Text = rs("path")
Call Ejecutar_shell(ruta_def & Form2.Text2.Text)
Call Grabar_Ini_j(Path_Archivo_Ini_j, rs("nombre"), Form2.Label25.Caption)
Form1.Check6(i).Value = 0
contador_de_ejecut = contador_de_ejecut + 1
Form2.Text4.Caption = contador_de_ejecut
Form2.progressbar1.Width = Form2.progressbar1.Width + Form2.Label20.Caption
End If
Next
'---------------------------------------


For i = 1 To Form1.Text7.Text
If Form1.Check7(i).Value = 1 Then
Form2.Text1.Text = Form2.Text1.Text & "------------------ [ Ofimática ] -----------------" & vbCrLf
ui = i
Set rs = cnn.Execute("SELECT * from Ofimática where id=" & ui)
Form2.Text1.Text = Form2.Text1.Text & "Instalar " & rs("nombre") & vbCrLf & "    Comando... " & vbCrLf
Form2.Caption = "Asistente de Instalación  -Instalando " & rs("nombre")
Form2.Text2.Text = rs("path")
Call Ejecutar_shell(ruta_def & Form2.Text2.Text)
Call Grabar_Ini_j(Path_Archivo_Ini_j, rs("nombre"), Form2.Label25.Caption)
Form1.Check7(i).Value = 0
contador_de_ejecut = contador_de_ejecut + 1
Form2.Text4.Caption = contador_de_ejecut
Form2.progressbar1.Width = Form2.progressbar1.Width + Form2.Label20.Caption
End If
Next
'---------------------------------------

For i = 1 To Form1.Text8.Text
If Form1.Check8(i).Value = 1 Then
Form2.Text1.Text = Form2.Text1.Text & "---------------- [ Tweaks_y_otros ] --------------" & vbCrLf
ui = i
Set rs = cnn.Execute("SELECT * from Tweaks_y_otros where id=" & ui)
Form2.Text1.Text = Form2.Text1.Text & rs("nombre") & vbCrLf
Form2.Text1.Text = Form2.Text1.Text & "Instalar " & rs("nombre") & vbCrLf & "    Comando... " & vbCrLf
Form2.Caption = "Asistente de Instalación  -Instalando " & rs("nombre")
Form2.Text2.Text = rs("path")
Call Ejecutar_shell(ruta_def & Form2.Text2.Text)
Call Grabar_Ini_j(Path_Archivo_Ini_j, rs("nombre"), Form2.Label25.Caption)
Form1.Check8(i).Value = 0
contador_de_ejecut = contador_de_ejecut + 1
Form2.Text4.Caption = contador_de_ejecut
Form2.progressbar1.Width = Form2.progressbar1.Width + Form2.Label20.Caption
End If
Next
'---------------------------------------

For i = 1 To Form1.Text9.Text
If Form1.Check9(i).Value = 1 Then
Form2.Text1.Text = Form2.Text1.Text & "------------------ [ Navegación ] ----------------" & vbCrLf
ui = i
Set rs = cnn.Execute("SELECT * from Navegación where id=" & ui)
Form2.Text1.Text = Form2.Text1.Text & "Instalar " & rs("nombre") & vbCrLf & "    Comando... " & vbCrLf
Form2.Caption = "Asistente de Instalación  -Instalando " & rs("nombre")
Form2.Text2.Text = rs("path")
Call Ejecutar_shell(ruta_def & Form2.Text2.Text)
Call Grabar_Ini_j(Path_Archivo_Ini_j, rs("nombre"), Form2.Label25.Caption)
Form1.Check9(i).Value = 0
contador_de_ejecut = contador_de_ejecut + 1
Form2.Text4.Caption = contador_de_ejecut
Form2.progressbar1.Width = Form2.progressbar1.Width + Form2.Label20.Caption
End If
Next
'---------------------------------------

For i = 1 To Form1.Text10.Text
If Form1.Check10(i).Value = 1 Then
Form2.Text1.Text = Form2.Text1.Text & "--------------------- [ Extra ] ------------------" & vbCrLf
ui = i
Set rs = cnn.Execute("SELECT * from Extra where id=" & ui)
Form2.Text1.Text = Form2.Text1.Text & "Instalar " & rs("nombre") & vbCrLf & "    Comando... " & vbCrLf
Form2.Caption = "Asistente de Instalación  -Instalando " & rs("nombre")
Form2.Text2.Text = rs("path")
Call Ejecutar_shell(ruta_def & Form2.Text2.Text)
Call Grabar_Ini_j(Path_Archivo_Ini_j, rs("nombre"), Form2.Label25.Caption)
Form1.Check10(i).Value = 0
contador_de_ejecut = contador_de_ejecut + 1
Form2.Text4.Caption = contador_de_ejecut
Form2.progressbar1.Width = Form2.progressbar1.Width + Form2.Label20.Caption
End If
Next
'---------------------------------------

ruta2 = App.path & "\recursos\AtEnd.wav"
Call sndPlaySound(ruta2, ASND_SYNC)
Form2.Caption = "Asistente de Instalación  -Instalación Finalizada"
Form1.Visible = False
Form2.Timer5.Enabled = True
Form2.Label24.Caption = ""
Form2.Label24.Visible = True
Form2.Shape6.Left = 0
Form2.Shape6.Top = 0
Form2.Shape6.Width = Form2.Width
Form2.Shape6.Height = Form2.Height
Exit Sub





nananna:
'MsgBox "Error"
End Sub
