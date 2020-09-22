VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H000000FF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5805
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8115
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   387
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   541
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "virtual_guitar_1@hotmail.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   6
      Left            =   2550
      TabIndex        =   6
      Top             =   5385
      Width           =   3150
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   5
      Left            =   6660
      TabIndex        =   5
      Top             =   105
      Width           =   255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "By Agustin Rodriguez"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   555
      Index           =   3
      Left            =   1635
      TabIndex        =   3
      Top             =   4890
      Width           =   4920
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Click a Dice and press + - to change the size"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   300
      Index           =   2
      Left            =   1290
      TabIndex        =   2
      Top             =   4440
      Width           =   5370
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Drag the Table or the Dice"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   300
      Index           =   1
      Left            =   4590
      TabIndex        =   1
      Top             =   4125
      Width           =   3210
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Double click to run the Dice"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   300
      Index           =   0
      Left            =   465
      TabIndex        =   0
      Top             =   4095
      Width           =   3330
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "By Agustin Rodriguez"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   555
      Index           =   4
      Left            =   1590
      TabIndex        =   4
      Top             =   4830
      Width           =   4920
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This is a bit of my "REAL 3D BACKGAMMON" project that I will publish soon.
'I believe that this is the ultimate procedure to make Games that uses Dice.
'Requires Windows XP

Option Explicit

Private Sub Form_DblClick()
    Dim i As Integer
    
    RND_Dice(1) = Int(Rnd * 6) + 1
   
    sldFrames_Value(1) = 0
    Dice(1).ZOrder 0
    Dice(1).tmrMoveNext_Timer
    Dice(1).tmrMoveNext.Interval = 1
    Dice(1).tmrMoveNext.Enabled = True
    
    SetLayeredWindowAttributes Dice(1).hwnd, 255, 255, LWA_COLORKEY Or LWA_ALPHA
    
    For i = 0 To 10000
        DoEvents
    Next i
    
    RND_Dice(2) = Int(Rnd * 6) + 1
    
    sldFrames_Value(2) = 0
    
    Dice(2).ZOrder 0
    Dice(2).tmrMoveNext_Timer
    Dice(2).tmrMoveNext.Interval = 1
    Dice(2).tmrMoveNext.Enabled = True
    SetLayeredWindowAttributes Dice(2).hwnd, 255, 255, LWA_COLORKEY Or LWA_ALPHA
    
End Sub

Private Sub Form_Load()
    Dim ret As Long
    
    Randomize
    
    ret = GetWindowLong(Me.hwnd, GWL_EXSTYLE)
    ret = ret Or WS_EX_LAYERED
    SetWindowLong Me.hwnd, GWL_EXSTYLE, ret
    SetLayeredWindowAttributes Me.hwnd, 255, 255, LWA_COLORKEY Or LWA_ALPHA
    
    RND_Frame(1) = -4
    RND_Frame(2) = 0
    RND_Frame(3) = -7
    RND_Frame(4) = -2
    RND_Frame(5) = -17
    RND_Frame(6) = -37
    
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    RXDice1 = Dice(1).Left - Left
    RYDice1 = Dice(1).Top - Top
    RXDice2 = Dice(2).Left - Left
    RYDice2 = Dice(2).Top - Top
    
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
    
        Dice(1).Move Left + RXDice1, Top + RYDice1
        Dice(2).Move Left + RXDice2, Top + RYDice2
        
    End If

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    capture = False

End Sub

Private Sub Form_Resize()

    DoEvents
    
    Dice(1).Tag = 1
    Dice(2).Tag = 2
 
    DoEvents
    Dice(1).Show , Me
    Dice(2).Show , Me
  
    Dice(1).Move Left + 2500, Top + 400 ', 4010, 1895 '5010, 1995
    Dice(2).Move Left + 3000, Top - 200 ', 4010, 1895 '5010, 1995

End Sub

Private Sub Label1_Click(Index As Integer)

    Unload Dice(1)
    Unload Dice(2)
    Unload Me

End Sub


