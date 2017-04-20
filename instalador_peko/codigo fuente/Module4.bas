Attribute VB_Name = "Module4"

Option Explicit


Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2


Private Type POINTAPI
        X As Long
        Y As Long
End Type
Private Declare Function SetWindowPos Lib "user32" ( _
    ByVal hwnd As Long, _
    ByVal hWndInsertAfter As Long, _
    ByVal X As Long, _
    ByVal Y As Long, _
    ByVal cx As Long, _
    ByVal cy As Long, _
    ByVal wFlags As Long) As Long
    


Private Declare Function GetCursorPos _
    Lib "user32" ( _
    lpPoint As POINTAPI) As Long
Declare Function SetTimer Lib "user32" ( _
    ByVal hwnd As Long, _
    ByVal nIDEvent As Long, _
    ByVal uElapse As Long, _
    ByVal lpTimerFunc As Long) As Long

Declare Function KillTimer Lib "user32" ( _
    ByVal hwnd As Long, _
    ByVal nIDEvent As Long) As Long

Dim m_Form              As Form
Dim label_text          As Label
Dim label_title         As Label
Dim Control_Image       As Image


 Sub show_ToolTipText( _
    the_form As Form, _
    the_title As String, _
    the_text As String, _
    Optional backColor As Long = &HC0FFFF, _
    Optional ForeColor As Long = 0, _
    Optional ForeColorTitle As Long = 0, _
    Optional image_path As String)
    
    Call destroy_tooltiptext
      
    Set m_Form = the_form
    Set label_text = m_Form.Controls.Add("vb.Label", "label_text") ' the title
    Set label_title = m_Form.Controls.Add("vb.Label", "label_title") ' the text
    
    If Len(image_path) Then
        Set Control_Image = m_Form.Controls.Add("vb.image", "Img1")
        With Control_Image
            .Picture = LoadPicture(image_path) ' load picture
            .Move 25, 25
            .Visible = True
        End With
    End If
    With label_title
        .Caption = the_title
        .BackStyle = 0
        .AutoSize = True
        .FontBold = True
        .FontSize = 9
        .ForeColor = ForeColorTitle
        If Len(image_path) Then
            .Left = 100 + Control_Image.Width
        Else
            .Left = 100
        End If
        .Top = 100
        .Visible = True
    End With
    With label_text
        .Caption = the_text
        .BackStyle = 0
        .AutoSize = True
        .ForeColor = ForeColor
        If Len(image_path) Then
            .Left = 100 + Control_Image.Width
        Else
            .Left = 100
        End If
        .Top = 100 + label_title.Top + label_title.Height
        .Visible = True
    End With
    With m_Form
        
        .backColor = backColor
        
        
        Dim h As Long
        Dim w As Long
        
        h = label_text.Height + label_title.Height + 350
        w = label_text.Width + 250
        If Len(image_path) Then
            .Width = w + Control_Image.Width
            .Height = h + Control_Image.Height
        Else
            .Width = w
            .Height = h
        End If
        .AutoRedraw = True
    End With
    
    With m_Form
        m_Form.Line (0, 0)-(m_Form.ScaleWidth, 0), &H80000009, B
        m_Form.Line (0, 0)-(0, m_Form.ScaleHeight), &H80000009, B
        m_Form.Line (0, .ScaleHeight - 10)-(.ScaleWidth, .ScaleHeight - 10), vbBlack, B
        m_Form.Line (.ScaleWidth - 10, 0)-(.ScaleWidth - 10, .ScaleHeight), vbBlack, B
    End With

    SetTimer m_Form.hwnd, 0, 1, AddressOf TimerProc
    
End Sub

Sub TimerProc( _
    ByVal hwnd As Long, _
    ByVal nIDEvent As Long, _
    ByVal uElapse As Long, _
    ByVal lpTimerFunc As Long)
    
    Dim mouse As POINTAPI
    GetCursorPos mouse
    
    With m_Form
        .Left = (mouse.X * Screen.TwipsPerPixelX) + 100
        .Top = (mouse.Y * Screen.TwipsPerPixelY) + 100
        
        If Not .Visible Then
            SetWindowPos .hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
            .Visible = True
        End If
        
    End With

End Sub

 Sub destroy_tooltiptext()
 
     If m_Form Is Nothing Then Exit Sub
     
     Call KillTimer(m_Form.hwnd, 0)
     
     Unload m_Form
     Set m_Form = Nothing
 
 End Sub
 


