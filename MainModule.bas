Attribute VB_Name = "MainModule"
Option Explicit




Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" ( _
    ByVal hwnd As Long, _
    ByVal nIndex As Long) As Long
    
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" ( _
    ByVal hwnd As Long, _
    ByVal nIndex As Long, _
    ByVal dwNewLong As Long) As Long
    
Private Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" ( _
    ByVal hwnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    ByVal lParam As Long) As Long
    
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" ( _
    ByVal lpPrevWndFunc As Long, _
    ByVal hwnd As Long, _
    ByVal Msg As Long, _
    ByVal wParam As Long, _
    ByVal lParam As Long) As Long
    
Public Declare Function GetProp Lib "user32" Alias "GetPropA" ( _
    ByVal hwnd As Long, _
    ByVal lpString As String) As Long
    
Public Declare Function SetProp Lib "user32" Alias "SetPropA" ( _
    ByVal hwnd As Long, _
    ByVal lpString As String, _
    ByVal hData As Long) As Long
    
Public Declare Function RemoveProp Lib "user32" Alias "RemovePropA" ( _
    ByVal hwnd As Long, _
    ByVal lpString As String) As Long

Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" ( _
    ByVal uAction As Long, _
    ByVal uParam As Long, _
    ByVal lpvParam As Any, _
    ByVal fuWinIni As Long) As Long

Public Declare Function GetSystemMetrics Lib "user32" ( _
    ByVal nIndex As Long) As Long

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)

Public Const GWL_WNDPROC = -4

Private Const WM_MOUSEWHEEL = &H20A

Private Const WHEEL_DELTA = 120
Private Const WHEEL_PAGESCROLL = &HFFFFFFFF

Public Const SPI_GETWHEELSCROLLLINES = 104

Public Const SM_MOUSEWHEELPRESENT = 75

Private Const MK_CONTROL = &H8          'Control key
Private Const MK_SHIFT = &H4            'Shift key
Private Const MK_LBUTTON = &H202        'Left mouse button
Private Const MK_MBUTTON = &H10         'Middle mouse button
Private Const MK_RBUTTON = &H2          'Right mouse button
Private Const MK_XBUTTON1 = &H20        'First X button; Windows 2000/XP only
Private Const MK_XBUTTON2 = &H40        'Second X button; Windows 2000/XP only

Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Const MOUSEEVENTF_LEFTDOWN = &H2
Const MOUSEEVENTF_LEFTUP = &H4
Const MOUSEEVENTF_MIDDLEDOWN = &H20
Const MOUSEEVENTF_MIDDLEUP = &H40
Const MOUSEEVENTF_MOVE = &H1
Const MOUSEEVENTF_ABSOLUTE = &H8000
Const MOUSEEVENTF_RIGHTDOWN = &H8
Const MOUSEEVENTF_RIGHTUP = &H10


    ' store a pointer to the form object
    ' which is set via ObjPtr
    Public lpFormObj As Long


Public Function WndProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim objForm As Form1
   On Error GoTo errorHandler
    If uMsg = WM_MOUSEWHEEL Then
      
    
        ' ##### Button/key pressed #####
        Select Case LoWord(wParam)
            
            Case MK_XBUTTON1
            Case MK_LBUTTON
            Case MK_MBUTTON
            Case MK_RBUTTON
            Case MK_XBUTTON2
        End Select
        

    'If the flexGrid is the active control then
     If TypeOf Form1.ActiveControl Is MSFlexGrid Then
        ' ##### Scroll direction #####
          If (HiWord(wParam) / WHEEL_DELTA) < 0 Then
            'Scrolling down
            Debug.Print "Down"
            

                ' instantiate the pointer we have to the form
                Set objForm = PtrToForm(lpFormObj)
                ' call the method
                objForm.ScrollDown
                ' destroy the reference
                Set objForm = Nothing

            
           Else
            'Scrolling up
            Debug.Print "UP"
            

                ' instantiate the pointer we have to the form
                Set objForm = PtrToForm(lpFormObj)
                ' call the method
                objForm.ScrollUp
                ' destroy the reference
                Set objForm = Nothing
            
            
        End If
    End If
         
        ' ##### Paging = suggested number of lines to scroll (e.g. in a textbox) #####
        ' Windows 95: Not supported
        Dim r As Long

        SystemParametersInfo SPI_GETWHEELSCROLLLINES, 0, r, 0
        
        If r = WHEEL_PAGESCROLL Then
            'Wheel roll should be interpreted as clicking
            'once in the page down or page up regions of
            'the scroll bar
        Else
            'Scroll 3 lines (3 is the default value)
        End If
    
        'Pass the message to default window procedure and then onto the parent
        DefWindowProc hwnd, uMsg, wParam, lParam
    Else
        'No messages handled, call original window procedure
        WndProc = CallWindowProc(GetProp(Form1.hwnd, "PrevWndProc"), hwnd, uMsg, wParam, lParam)
    End If
    Exit Function
errorHandler:
Debug.Print Err.Number & " " & Err.Description
End Function

Public Function HiWord(dw As Long) As Integer
    If dw And &H80000000 Then
        HiWord = (dw \ 65535) - 1
    Else
        HiWord = dw \ 65535
    End If
End Function

Public Function LoWord(dw As Long) As Integer
    If dw And &H8000& Then
        LoWord = &H8000 Or (dw And &H7FFF&)
    Else
        LoWord = dw And &HFFFF&
    End If
End Function

'//--[PtrToForm]--------------------------------//
'
'  Creates a dummy object from an ObjPtr
'
Public Function PtrToForm(ByVal lPtr As Long) As Form1
Dim obj As Form1
    ' instantiate the illegal referece
    CopyMemory obj, lPtr, 4
    Set PtrToForm = obj
    CopyMemory obj, 0&, 4
End Function
