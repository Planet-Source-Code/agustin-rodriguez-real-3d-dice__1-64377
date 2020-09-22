VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H000000FF&
   BorderStyle     =   0  'None
   ClientHeight    =   3090
   ClientLeft      =   -10005
   ClientTop       =   0
   ClientWidth     =   4680
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   206
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picView 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   780
      Left            =   0
      ScaleHeight     =   52
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   81
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Timer tmrMoveNext 
      Enabled         =   0   'False
      Left            =   315
      Top             =   1755
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'                   ANIMATED AND TRANSPARENT SPLASH WINDOW USING GIF FILES

'This Project use the Excelent Animated GIF Class By Vlad Vissoultchev

' Run on Windows XP

Option Explicit

Private Declare Function StretchBlt Lib "GDI32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function SetStretchBltMode Lib "GDI32" (ByVal hDC As Long, ByVal hStretchMode As Long) As Long
Private Const STRETCHMODE As Long = vbPaletteModeNone
Private Declare Function apiSetWindowPos Lib "User32" Alias "SetWindowPos" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetWindowLong Lib "User32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "User32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "User32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

Private Const HWND_TOPMOST As Integer = -1
Private Const HWND_NOTOPMOST As Integer = -2
Private Const SWP_NOMOVE As Integer = &H2
Private Const SWP_NOSIZE As Integer = &H1
Private Const LWA_COLORKEY As Integer = &H1
Private Const LWA_ALPHA As Integer = &H2
Private Const GWL_EXSTYLE As Integer = (-20)
Private Const WS_EX_LAYERED As Long = &H80000
Private Const MODULE_NAME As String = "frmAnimation"

Private sldFrames_Max As Long
Private GIF As cGifReader
Private m_oRenderer             As cBmpRenderer
Private WithEvents m_oReader    As cGifReader
Attribute m_oReader.VB_VarHelpID = -1
Private m_lFrameCount           As Long
Private m_aFrames()             As UcsFrameInfo

Private Type UcsFrameInfo
    oPic        As StdPicture
    nDelay      As Long
End Type

Private Sub Form_Activate()

  Const FUNC_NAME     As String = "Form_Activate"
  Dim lIdx            As Long
  Dim sInfo           As String
  Static vez As Integer
  
    On Error GoTo EH
    If vez = False Then
        vez = True
    
        If UBound(m_aFrames) < 0 And m_lFrameCount > 0 Then
            ReDim m_aFrames(1 To m_lFrameCount)
            If m_oRenderer.MoveFirst() Then
                lIdx = 0
                Do While True
                    If Not m_oRenderer.MoveNext Then
                        Exit Do '>---> Loop
                    End If
                    lIdx = lIdx + 1
                    With m_aFrames(lIdx)
                        Set .oPic = m_oRenderer.Image
                        .nDelay = m_oRenderer.Reader.DelayTime
                        sldFrames_Value(Tag) = lIdx
                        If lIdx = 1 Then
                            sldFrames_Change
                        End If
                        'DoEvents
                    End With 'M_AFRAMES(LIDX)
                Loop
            End If
        End If
    
        If Tag = 2 Then
            Form1.MousePointer = 0
            Form1.Enabled = True
        End If
   
    End If

Exit Sub

EH:
    Resume Next

End Sub

Private Sub Form_Initialize()

    Set m_oRenderer = New cBmpRenderer
    ReDim m_aFrames(-1 To -1)
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode

      Case 107
        Dice(Tag).Width = Dice(Tag).Width + Dice(Tag).Width * 10 / 100
        Dice(Tag).Height = Dice(Tag).Height + Dice(Tag).Height * 10 / 100
 
      Case 109
        Dice(Tag).Width = Dice(Tag).Width - Dice(Tag).Width * 10 / 100
        Dice(Tag).Height = Dice(Tag).Height - Dice(Tag).Height * 10 / 100

    End Select

    sldFrames_Value(Tag) = 0
    Dice(Tag).ZOrder 0
    Dice(Tag).tmrMoveNext_Timer
    Dice(Tag).tmrMoveNext.Interval = 1
    Dice(Tag).tmrMoveNext.Enabled = True
   
End Sub

Private Sub Form_Load()

  Dim ret As Long
  Dim Arquivo As String
  Dim filenumber As Integer

  Dim sFilename   As String * 260
  Dim lRetval     As Long
  Dim Path_wallpaper As String

    ret = GetWindowLong(Me.hwnd, GWL_EXSTYLE)
    ret = ret Or WS_EX_LAYERED
    SetWindowLong Me.hwnd, GWL_EXSTYLE, ret
    
    SetLayeredWindowAttributes Me.hwnd, 255, 0, LWA_COLORKEY Or LWA_ALPHA
    
    Set GIF = New cGifReader
    
    Arquivo = App.Path & "\dado pq2.gif"
    
    If GIF.Init(Arquivo) Then
        GIF.MoveFirst
    End If
    ReDim m_aFrames(-1 To -1)
    Init GIF
      
    Width = (GIF.ScreenWidth) * Screen.TwipsPerPixelX  '+ 1000            'HERE YOU SET THE WIDTH
    Height = (GIF.ScreenHeight) * Screen.TwipsPerPixelY  '+ 1000          'AND THE HEIGHT
    
    picView.Width = Width
    picView.Height = Height
    'picView.BackColor = GIF.BackgroundColor
    
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
          
    XX = X * Screen.TwipsPerPixelX
    YY = Y * Screen.TwipsPerPixelY
        
    capture = True
    ReleaseCapture
    SetCapture Me.hwnd

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

  Dim i As Integer
  
    If capture Then
        GetCursorPos Pt
        
        Move Pt.X * Screen.TwipsPerPixelX - XX, Pt.Y * Screen.TwipsPerPixelY - YY
        
    End If

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    capture = False

End Sub

Private Sub Form_Resize()
     
    StretchBlt hDC, 0, 0, ScaleWidth, ScaleHeight, picView.hDC, 0, 0, picView.ScaleWidth, picView.ScaleHeight, vbSrcCopy
   
End Sub

Private Sub sldFrames_Change()

  Const FUNC_NAME     As String = "sldFrames_Change"
  Dim lDelay          As Long
    
    On Error GoTo EH
    With m_aFrames(sldFrames_Value(Tag))
        lDelay = IIf(.nDelay < 8, 10, .nDelay * 10)
        Set picView.Picture = .oPic
        picView.Refresh
        StretchBlt hDC, 0, 0, ScaleWidth, ScaleHeight, picView.hDC, 0, 0, picView.ScaleWidth, picView.ScaleHeight, vbSrcCopy
       
        If tmrMoveNext.Enabled Then
            tmrMoveNext.Interval = lDelay
            tmrMoveNext.Enabled = False
            tmrMoveNext.Enabled = True
        End If
    End With 'M_AFRAMES(SLDFRAMES_VALUE(TAG))
        
Exit Sub

EH:
    Resume Next

End Sub

Public Sub tmrMoveNext_Timer()
    
    If sldFrames_Value(Tag) + 1 > sldFrames_Max + RND_Frame(RND_Dice(Tag)) Then
        tmrMoveNext.Enabled = False
        Exit Sub '>---> Bottom
    End If
        
    sldFrames_Value(Tag) = sldFrames_Value(Tag) + 1
    sldFrames_Change
    DoEvents

End Sub

Public Function Init(oRdr As cGifReader) As Boolean

  Const FUNC_NAME     As String = "Init"
    
    On Error GoTo EH
    
    Set m_oReader = oRdr
    If m_oRenderer.Init(oRdr) Then
        Set picView.Picture = Nothing
        If oRdr.MoveLast() Then
            m_lFrameCount = oRdr.FrameIndex + 1
            If m_lFrameCount > 1 Then
                sldFrames_Max = m_lFrameCount
            End If
        End If
    End If

Exit Function

EH:
    Resume Next

End Function


