VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4845
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4095
   LinkTopic       =   "Form1"
   ScaleHeight     =   4845
   ScaleWidth      =   4095
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   5530
      _Version        =   393216
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   2
      ScrollBars      =   2
      SelectionMode   =   1
   End
   Begin VB.Label Label1 
      Caption         =   "Click on the grid.            Then with the mouse wheel scroll either up or down."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   360
      TabIndex        =   1
      Top             =   0
      Width           =   3495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Dim X As Integer
Dim topRow As Integer
Dim ctl As Control
Dim lngResult As Long

Private Sub Form_Load()
formatGrid MSFlexGrid1
ColorForm

For X = 0 To 49
 'vbTab tabs to the next column.
 MSFlexGrid1.AddItem Chr$(Int(Rnd() * 26) + 65) & vbTab & Int(1000 * Rnd) & vbTab & Int(100 * Rnd) & vbTab & Int(1000 * Rnd)
Next X

    ' set the module level callback pointer
    lpFormObj = ObjPtr(Me)


SetProp Form1.hwnd, "PrevWndProc", SetWindowLong(Form1.hwnd, GWL_WNDPROC, AddressOf WndProc)
If GetSystemMetrics(SM_MOUSEWHEELPRESENT) Then
   ' MsgBox "A simple call to GetSystemMetrics tells you whether or not the mouse has a wheel. The mouse connected to this computer does have a wheel.", vbInformation + vbOKOnly, App.Title
   Debug.Print "Yes Wheel"
Else
  ' MsgBox "A simple call to GetSystemMetrics tells you whether or not the mouse has a wheel. The mouse connected to this computer doesn't have a wheel.", vbInformation + vbOKOnly, App.Title
    Debug.Print "No Wheel"
End If

topRow = 1
End Sub



'//--[ScrollUp]---------------------------//
'
'  called from the MainModule WndProc sub
'  when a up-scrolling mouse message is
'  received
'
Public Sub ScrollUp()
    ' scroll up..
    If topRow > 1 Then
        topRow = topRow - 1
        MSFlexGrid1.topRow = topRow
    End If
End Sub

'//--[ScrollDown]---------------------------//
'
'  called from the MainModule WndProc sub
'  when a down-scrolling mouse message is
'  received
'
Public Sub ScrollDown()
    ' scroll down..
    If topRow < MSFlexGrid1.Rows - 1 Then
        topRow = topRow + 1
        MSFlexGrid1.topRow = topRow
    End If
End Sub

Private Sub formatGrid(gRid As MSFlexGrid)

'Format the color of the grid
With gRid
    .Rows = 2
    'BackColorFixed is the fixed row at the top and to the left
    'Ones here are named Field1, Field2 and Field3 Field4 at the top,
    'random letters down the left
    .BackColorFixed = vbWhite 'RGB(185, 212, 233)
    'Furthest back background color
    .BackColorBkg = vbBlack 'RGB(48, 111, 160)
    'Background of cellwhen it is selected
    .BackColorSel = RGB(95, 95, 143)
    'Forecolor of cell when it is selected
    .ForeColorSel = RGB(192, 189, 215)
    'Backcolor of cell when it is unselected
    .BackColor = RGB(192, 189, 215)
    'Forecolor of cell when it is unselected
    .ForeColor = RGB(95, 95, 143)
    .GridColor = vbBlack
    Dim S$
    'Put headers in a string
    S$ = "<Field1|<Field2|<Field3|<Field4 "
    
    'Print the header string
    .FormatString = S$
    
    'Set the width of the columns
    .ColWidth(0) = 1000
    .ColWidth(1) = 750
    .ColWidth(2) = 1000
    
End With
End Sub
Private Sub ColorForm()
'Loop through all controls in form and
'set the forcolor and backcolor for all textboxs
For Each ctl In Me
 If TypeOf ctl Is TextBox Then
  With ctl
  .BackColor = RGB(95, 95, 143)
  .ForeColor = vbWhite
  End With
 End If
Next

'Loop through all controls in form and
'set the forcolor and backcolor for all labels
For Each ctl In Me
 If TypeOf ctl Is Label Then
  With ctl
  .BackColor = RGB(48, 111, 160)
  .ForeColor = vbWhite
  End With
 End If
Next

'Set the forms backcolor
Me.BackColor = RGB(48, 111, 160)

End Sub

