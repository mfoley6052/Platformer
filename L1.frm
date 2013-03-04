VERSION 5.00
Begin VB.Form Form1 
   ClientHeight    =   5610
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   12930
   LinkTopic       =   "Form1"
   ScaleHeight     =   5610
   ScaleWidth      =   12930
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrJcount 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2040
      Top             =   2280
   End
   Begin VB.TextBox txtDbg 
      Enabled         =   0   'False
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   360
      Width           =   9015
   End
   Begin VB.Timer tmrJump 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1680
      Top             =   2640
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   12960
      Y1              =   4800
      Y2              =   4800
   End
   Begin VB.Shape Player 
      BackStyle       =   1  'Opaque
      Height          =   255
      Left            =   360
      Shape           =   3  'Circle
      Top             =   4560
      Width           =   255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim speed As Single
Dim vspeed As Single
Dim gravity As Single
Dim jumping As Boolean
Dim plLeft As Boolean
Dim plRight As Boolean
Dim jPressed As Boolean
Dim jTime As Integer
Dim initDist As Single


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'MsgBox (KeyCode)
If KeyCode = 37 Then
    Player.left = Player.left - speed
    plLleft = True
    plRight = False
ElseIf KeyCode = 39 Then
    Player.left = Player.left + speed
    plRight = True
    plLeft = False
ElseIf KeyCode = 38 Then
    If jumping = False Then
        initDist = Player.Top
        vspeed = 100
        jumping = True
        tmrJump.Enabled = True
'    Else
'        jPressed = True
'        tmrJcount.Enabled = False
    End If
End If
'If jumping = False Then
'vspeed = jTime
'End If
End Sub

Private Sub Form_Load()
speed = 50
vspeed = 100
gravity = 9.8
End Sub

'Private Sub tmrJcount_Timer()
'If jPressed = True Then
'jTime = jTime + 1
'Else
'    tmrJcount.Enabled = False
'End If
'End Sub

Private Sub tmrJump_Timer()
Dim dist As Single
Dim curDist As Single

'If jPressed = True Then
    Player.Top = Player.Top - vspeed
'End If
If plLeft = True Then
    Player.left = Player.left - speed
ElseIf plRight = True Then
    Player.left = Player.left + speed
End If
If vspeed >= 0 Then
    vspeed = vspeed - (0.3 * gravity)
Else
    vspeed = vspeed - (10 * gravity)
End If
txtDbg.Text = vspeed & ", " & jumping
'If vspeed <= -jTime Then
curDist = Player.Top
dist = initDist - curDist
If dist <= 0 Then
    Player.Top = initDist
    jumping = False
    vspeed = 100
    tmrJump.Enabled = False
Else
    jumping = True
End If
End Sub
