Attribute VB_Name = "Module1"

Public Const GWL_EXSTYLE As Long = (-20)
Public Const WS_EX_LAYERED As Long = &H80000
Public Const WS_EX_TRANSPARENT As Long = &H20&
Public Const LWA_ALPHA As Long = &H2&
Public Const LWA_COLORKEY As Integer = &H1

Public Declare Function FindWindow Lib "User32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function GetWindowLong Lib "User32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "User32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetLayeredWindowAttributes Lib "User32" (ByVal hwnd As Long, ByVal crey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

Public sldFrames_Value(1 To 2) As Long

Public Dice(1 To 2) As New Form3
Public RND_Frame(1 To 6) As Integer
Public RND_Dice(1 To 2) As Integer
Public capture As Integer
Public Type POINTAPI
    X As Long
    Y As Long
End Type
Public Declare Function SetCapture Lib "User32" (ByVal hwnd As Long) As Long
Public Declare Function ReleaseCapture Lib "User32" () As Long
Public Declare Function GetCursorPos Lib "User32" (lpPoint As POINTAPI) As Long
Public Pt As POINTAPI
Public XX As Long
Public YY As Long
Public RXDice1  As Long
Public RYDice1  As Long
Public RXDice2  As Long
Public RYDice2  As Long


