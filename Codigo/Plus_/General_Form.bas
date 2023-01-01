Attribute VB_Name = "General_Form"

' ***** IMPIDE CERRAR L VENTANA ******
'Definición de la API bloquear las teclas
Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Integer, ByVal bRevert As Integer) As Integer
Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Integer, ByVal nPosition As Integer, ByVal wFlags As Integer) As Integer
'Definicion de las constantes
Global Const MF_BYPOSITION = &H400
' ***** IMPIDE CERRAR L VENTANA ******



' ***** SYSTRAY ******
Option Explicit
Public Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type
Public Enum enm_NIM_Shell
    NIM_ADD = &H0
    NIM_MODIFY = &H1
    NIM_DELETE = &H2
    NIF_MESSAGE = &H1
    NIF_ICON = &H2
    NIF_TIP = &H4
    WM_MOUSEMOVE = &H200
End Enum
Public nidProgramData As NOTIFYICONDATA


Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function GetForegroundWindow Lib "user32" () As Long
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As enm_NIM_Shell, pnid As NOTIFYICONDATA) As Boolean
Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

Public Const VK_O = 79
Public Const HWND_BOTTOM = 1
Public Const SWP_HIDEWINDOW = &H80
Public Const VK_ESCAPE = &H1B
Public Const HWND_TOPMOST = -1
Public Const SWP_SHOWWINDOW = &H40
Public Const HWND_TOP = 0


Global Wnd(10) As Long
Global num_win As Long
Global xx(10) As Long
Global yy(10) As Long
Global ww(10) As Long
Global hh(10) As Long
Global show_in_menu As Boolean
Global letter As Long
Global restore_win As Boolean
' ***** SYSTRAY ******
