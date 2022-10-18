VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00808000&
   Caption         =   "Form1"
   ClientHeight    =   6120
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8130
   LinkTopic       =   "Form1"
   ScaleHeight     =   6120
   ScaleWidth      =   8130
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer3 
      Left            =   1920
      Top             =   5640
   End
   Begin VB.Timer Timer2 
      Left            =   1320
      Top             =   5640
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   720
      Top             =   5640
   End
   Begin VB.Label pointwon 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3240
      TabIndex        =   2
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Shape powerup1 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   3840
      Top             =   4560
      Width           =   255
   End
   Begin VB.Label player2 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3840
      TabIndex        =   1
      Top             =   120
      Width           =   735
   End
   Begin VB.Label player1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2520
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
   Begin VB.Shape ball 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   375
      Left            =   4560
      Shape           =   3  'Circle
      Top             =   2880
      Width           =   255
   End
   Begin VB.Shape paddle2 
      FillColor       =   &H0000C0C0&
      FillStyle       =   0  'Solid
      Height          =   1335
      Left            =   7200
      Top             =   1680
      Width           =   255
   End
   Begin VB.Shape paddle1 
      FillColor       =   &H0000C0C0&
      FillStyle       =   0  'Solid
      Height          =   1335
      Left            =   240
      Top             =   1680
      Width           =   255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim x, x2, y, p1, p2, total As Integer
Dim hit, normal1, normal2 As Boolean
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Function GetActiveWindow Lib "user32" () As Long
Private Sub Form_Load()
total = InputBox("How many points too win?")
Randomize
p1 = 0 'player 1 score
p2 = 0  'player 2 score
Me.KeyPreview = True
powerup1.Visible = True
'scorelabels
player1.Caption = p1
player2.Caption = p2

x2 = 4
normal1 = True
normal2 = True
powerup1.Left = 5000 * Rnd
powerup1.Top = 5000 * Rnd
'speed of ball
x = -50
y = -50

'speed of game loop
Timer1.Interval = 1
Timer2.Interval = 10
Timer3.Interval = 1000

'sets the forms dimensions
Form1.Height = 5900
Form1.Width = 7755

'starts the timer
Timer1.Enabled = True
Timer2.Enabled = True
Timer3.Enabled = False
End Sub

Private Sub Timer1_Timer()
'moves ball by x and y speed
ball.Move ball.Left + y, ball.Top + x
pointwon.Caption = " "
x2 = 4

    'if ball hits the top reverse ball
    If ball.Top <= 0 Then
    x = x * -1.05
        If crazy = True Then
            ball.BackColor = QBColor(Rnd * 15)
            BackColor = QBColor(Rnd * 15)
            paddle1.FillColor = QBColor(Rnd * 15)
            paddle2.FillColor = QBColor(Rnd * 15)
            powerup1.FillColor = QBColor(Rnd * 15)
        End If
    End If
    
    'if ball collides with paddle1 reverse ball
    If paddle1.Left + paddle1.Width >= ball.Left And paddle1.Left <= ball.Left + ball.Width And paddle1.Top + paddle1.Height >= ball.Top And paddle1.Top <= ball.Top + ball.Height Then
        y = y * -1.05
        hit = True
        If crazy = True Then
            ball.BackColor = QBColor(Rnd * 15)
            BackColor = QBColor(Rnd * 15)
            paddle1.FillColor = QBColor(Rnd * 15)
            paddle2.FillColor = QBColor(Rnd * 15)
            powerup1.FillColor = QBColor(Rnd * 15)
        End If
    End If
    
    'a powerup that halfs your paddles length
    If powerup1.Left + powerup1.Width >= ball.Left And powerup1.Left <= ball.Left + ball.Width And powerup1.Top + powerup1.Height >= ball.Top And powerup1.Top <= ball.Top + ball.Height Then
        If hit = True And powerup1.Visible = True Then
            paddle1.Height = paddle1.Height / 2
            normal1 = False
        End If
        If hit = False And powerup1.Visible = True Then
            paddle2.Height = paddle2.Height / 2
            normal2 = False
        End If
        powerup1.Visible = False
        
        
        
        
    End If
    
    'if ball goes past paddle1 then add score and reset ball position
    If ball.Left <= 200 Then
        p2 = p2 + 1
        ball.Top = 5000
        ball.Left = 5000
        player2.Caption = p2
        x = -50
        y = -50
        powerup1.Left = 5000 * Rnd
        powerup1.Top = 5000 * Rnd
        powerup1.Visible = True
        Timer3.Enabled = True
        Timer1.Enabled = False
        If normal1 = False Then
            paddle1.Height = paddle1.Height * 2
            normal1 = True
        ElseIf normal2 = False Then
            paddle2.Height = paddle2.Height * 2
            normal2 = True
           
        End If
    End If
    
    'if ball goes past paddle2 then add score and reset ball position
    If ball.Left >= 7400 Then
        p1 = p1 + 1
        ball.Top = 1000
        ball.Left = 1000
        player1.Caption = p1
        x = 50
        y = 50
        powerup1.Left = 5000 * Rnd
        powerup1.Top = 5000 * Rnd
        powerup1.Visible = True
        Timer3.Enabled = True
        Timer1.Enabled = False
        If normal1 = False Then
            paddle1.Height = paddle1.Height * 2
            normal1 = True
        ElseIf normal2 = False Then
            paddle2.Height = paddle2.Height * 2
            normal2 = True
        End If
    End If
    'if ball hits bottom of form reverse ball
    If ball.Top >= 5100 Then
        x = x * -1.05
        If crazy = True Then
            ball.BackColor = QBColor(Rnd * 15)
            BackColor = QBColor(Rnd * 15)
            paddle1.FillColor = QBColor(Rnd * 15)
            paddle2.FillColor = QBColor(Rnd * 15)
            powerup1.FillColor = QBColor(Rnd * 15)
        End If
    End If
    
    'if ball collides with paddle2 reverse ball
    If paddle2.Left + paddle2.Width >= ball.Left And paddle2.Left <= ball.Left + ball.Width And paddle2.Top + paddle2.Height >= ball.Top And paddle2.Top <= ball.Top + ball.Height Then
        y = y * -1.05
        hit = False
        If crazy = True Then
            ball.BackColor = QBColor(Rnd * 15)
            BackColor = QBColor(Rnd * 15)
            paddle1.FillColor = QBColor(Rnd * 15)
            paddle2.FillColor = QBColor(Rnd * 15)
            powerup1.FillColor = QBColor(Rnd * 15)
        End If
    End If
    
    
    If p1 = total Then
        MsgBox "Player 1 has won!!"
        Unload Me
    ElseIf p2 = total Then
        MsgBox "Player 2 has won!!"
        Unload Me
    End If

    
End Sub




Private Sub Timer2_Timer()
    'check if window has focus
    'otherwise window will move even when another program has focus
    If GetActiveWindow() <> Me.hWnd Then Exit Sub
    
    If GetAsyncKeyState(vbKeyW) And paddle1.Top >= 0 Then paddle1.Top = paddle1.Top - 120
        
    If GetAsyncKeyState(vbKeyS) And paddle1.Top <= 3900 Then paddle1.Top = paddle1.Top + 120
    
    If GetAsyncKeyState(vbKeyUp) And paddle2.Top >= 0 Then paddle2.Top = paddle2.Top - 120
    If GetAsyncKeyState(vbKeyDown) And paddle2.Top <= 3900 Then paddle2.Top = paddle2.Top + 120
End Sub

Private Sub Timer3_Timer()
x2 = x2 - 1
pointwon.Caption = x2
If x2 = 0 Then
    Timer1.Enabled = True
    Timer3.Enabled = False
End If

End Sub
