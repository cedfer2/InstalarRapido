Attribute VB_Name = "ejec"
    Option Explicit
      
      
    'Funciones del api
    '''''''''''''''''''''''''''''''''''''''''''''''''''
    Private Declare Function OpenProcess Lib "kernel32" _
      (ByVal dwDesiredAccess As Long, _
       ByVal bInheritHandle As Long, _
       ByVal dwProcessId As Long) As Long
      
    Private Declare Function GetExitCodeProcess Lib "kernel32" _
      (ByVal hProcess As Long, lpExitCode As Long) As Long
      
    Private Declare Function CloseHandle Lib "kernel32" _
      (ByVal hObject As Long) As Long
      
    'Constantes
    '''''''''''''''''''''''''''''''''''''''''''''''''''
    Private Const PROCESS_QUERY_INFORMATION = &H400
    Private Const STATUS_PENDING = &H103&
    
    Dim r As Long
    Dim msg As String
   
    ' Recibe el argumento de la línea de comandos para pasarle al Shell
    Public Sub Ejecutar_shell(programa As String)
      On Error GoTo errSub:
        Dim ruta_def As String
        Dim handle_Process As Long
        Dim id_process As Long
        Dim lp_ExitCode As Long
        ruta_def = App.path & "\recursos\aplicaciones\"
        ' Abre el proceso con el shell
        id_process = Shell(programa, 1)
          
          
          
        ' handle del proceso
        handle_Process = OpenProcess(PROCESS_QUERY_INFORMATION, False, id_process)
          
        ' Mientras lp_ExitCode = STATUS_PENDING, se ejecuta el do
        Do
      
            Call GetExitCodeProcess(handle_Process, lp_ExitCode)
              
            DoEvents
        'Form2.Text1.Text = Form2.Text1.Text
        'Form2.Label5.Caption = " Se esta ejecutando " & programa & " y esperaro a que cierre "
        Form2.Label25.Caption = 2
        Form2.Label5.Caption = programa
        Form2.Text5.Text = "         -Se esta ejecutando proceso con ID:" & handle_Process & vbCrLf & "         -Esperaro a que cierre Proceso ID: " & handle_Process
        
        
        Loop While lp_ExitCode = STATUS_PENDING
        
        ' fin
        ' Cierra
        Call CloseHandle(handle_Process)
      
      'Form2.Label1.Caption = "Se cerro... " & programa
      '  MsgBox "Se cerró el " & programa, vbInformation
      
    '  Form2.Text1.Text = Form2.Text1.Text & "Se cerro... " & programa & vbCrLf & " --------------------------------------------------------------------------------------------------"

     Dim rutaju As String
      rutaju = App.path & "\recursos\Yes.wav"
Call sndPlaySound(rutaju, ASND_SYNC)
        Form2.SetFocus
    Form2.Label25.Caption = 3
      Form2.Text1.Text = Form2.Text1.Text & vbCrLf & "         -Finalizo proceso: " & handle_Process & vbCrLf
    Exit Sub
    
    
errSub:
 
     Dim rutajuju As String
      rutajuju = App.path & "\recursos\nNo.wav"
Call sndPlaySound(rutajuju, ASND_SYNC)
        Form2.SetFocus
 
 Form2.Label25.Caption = 4
 Form2.Text1.Text = Form2.Text1.Text & "        -Error No. " & Err & vbCrLf & "        -" & Error(Err) & vbCrLf
    Form2.Label5.Caption = "Error!!!"
  'Form2.Label1.Caption = "Error No. " & Err & ": " & Error(Err)
  Form2.SetFocus
    End Sub


