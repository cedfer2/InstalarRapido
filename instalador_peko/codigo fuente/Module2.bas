Attribute VB_Name = "Module2"
Option Explicit

Public Declare Sub InitCommonControls Lib "comctl32" ()

Public cnn As New ADODB.Connection
Public rs As New ADODB.Recordset

Public Sub IniciarConexion()
On Error GoTo fr
Dim fso As Object
Dim mia_dir As String
Dim mie_dir As String

mia_dir = App.path & "\recursos\datos.mdb"
mie_dir = Environ("TEMP") & "\datos.mdb"
Set fso = CreateObject("Scripting.FileSystemObject")

If fso.FileExists(mie_dir) = True Then
Kill mie_dir
End If
'Set fso = Nothing


If fso.FileExists(mia_dir) = True Then
FileCopy mia_dir, mie_dir
    'With cnn
     '   .CursorLocation = adUseClient
      '  .Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & _
       '       App.path & "\recursos\datos.mdb" & ";Persist Security Info=False"
    'End With

With cnn
        .CursorLocation = adUseClient
        .Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & _
         mie_dir & ";Persist Security Info=False"
End With
Set fso = Nothing
Else
Form1.Timer1.Enabled = False
Form4.Hide
MsgBox "No se puede inicar operación porque no se encuentra la base de datos", vbCritical, "Advertencia"
MsgBox "Asegurese que la base de datos existe, e intente correr el instalador", vbInformation, "Información"
End
End If
Exit Sub
fr:
Form1.Timer1.Enabled = False
Form4.Hide
MsgBox "Error en la base de datos" & vbCrLf & ":Descripción" & vbCrLf & Err.Description & vbCrLf & vbCrLf & " De clic en aceptar para cerrar aplicacion", vbCritical + vbOKOnly + vbMsgBoxRtlReading, "Error"
End
End Sub


Sub Desconectar()
    On Local Error Resume Next
    rs.Close
    Set rs = Nothing
    cnn.Close
    Set cnn = Nothing
End Sub

Public Sub todos()

Form1.Check11.Value = 1
Form1.Check12.Value = 1
Form1.Check13.Value = 1
Form1.Check14.Value = 1
Form1.Check15.Value = 1
Form1.Check16.Value = 1
Form1.Check17.Value = 1
Form1.Check18.Value = 1
Form1.Check19.Value = 1
Form1.Check20.Value = 1

End Sub

Public Sub ninguno()
Dim i As Integer
Form1.Check11.Value = 0
Form1.Check12.Value = 0
Form1.Check13.Value = 0
Form1.Check14.Value = 0
Form1.Check15.Value = 0
Form1.Check16.Value = 0
Form1.Check17.Value = 0
Form1.Check18.Value = 0
Form1.Check19.Value = 0
Form1.Check20.Value = 0

For i = 1 To Form1.Text1.Text
Form1.Check1(i).Value = 0
Next

For i = 1 To Form1.Text2.Text
Form1.Check2(i).Value = 0
Next

For i = 1 To Form1.Text3.Text
Form1.Check3(i).Value = 0
Next

For i = 1 To Form1.Text4.Text
Form1.Check4(i).Value = 0
Next

For i = 1 To Form1.Text5.Text
Form1.Check5(i).Value = 0
Next

For i = 1 To Form1.Text6.Text
Form1.Check6(i).Value = 0
Next


For i = 1 To Form1.Text7.Text
Form1.Check7(i).Value = 0
Next

For i = 1 To Form1.Text8.Text
Form1.Check8(i).Value = 0
Next

For i = 1 To Form1.Text9.Text
Form1.Check9(i).Value = 0
Next

For i = 1 To Form1.Text10.Text
Form1.Check10(i).Value = 0
Next

End Sub
