VERSION 5.00
Begin VB.Form Game 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "NIBBLES!"
   ClientHeight    =   3375
   ClientLeft      =   5535
   ClientTop       =   4005
   ClientWidth     =   4320
   Icon            =   "Game.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   4320
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Bomb 
      Left            =   3600
      Top             =   360
   End
   Begin VB.Timer tmrGetReady 
      Left            =   3480
      Top             =   2880
   End
   Begin VB.Timer tmrWorm 
      Left            =   3600
      Top             =   2880
   End
   Begin VB.Timer tmrQuit 
      Left            =   3720
      Top             =   2880
   End
   Begin VB.Shape Itzabomb 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   135
      Left            =   2760
      Shape           =   3  'Circle
      Top             =   840
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label lblHighScore 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFC0FF&
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   240
      Width           =   4335
   End
   Begin VB.Label Life 
      BackStyle       =   0  'Transparent
      Caption         =   "Lives:"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0FF&
      Height          =   255
      Left            =   1920
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H0000FF00&
      Height          =   255
      Left            =   0
      Shape           =   5  'Rounded Square
      Top             =   0
      Width           =   255
   End
   Begin VB.Label Quit 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   255
   End
   Begin VB.Shape TheWall 
      BackColor       =   &H00000080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFF00&
      Height          =   135
      Index           =   0
      Left            =   2280
      Top             =   480
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label lblReady 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "GET READY"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   1320
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Shape Wall 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      Height          =   135
      Left            =   2040
      Top             =   840
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape Special 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000FFFF&
      Height          =   135
      Left            =   2280
      Top             =   840
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Options 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Options"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   1560
      TabIndex        =   1
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0000FF00&
      Height          =   495
      Left            =   1560
      Shape           =   4  'Rounded Rectangle
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Shape Target 
      BackColor       =   &H0000FF00&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000FF00&
      Height          =   135
      Left            =   2520
      Top             =   840
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape Nibblepart 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000005&
      Height          =   135
      Index           =   0
      Left            =   4200
      Top             =   3240
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Start 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Start"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   1560
      TabIndex        =   0
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000FF00&
      Height          =   495
      Left            =   1560
      Shape           =   4  'Rounded Rectangle
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label Stuff 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright (c) 2003 by Jason Liang"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   3120
      Width           =   4335
   End
   Begin VB.Label lblScore 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0FF&
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   1455
   End
End
Attribute VB_Name = "Game"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'dimming the file system object
Dim fso As New FileSystemObject
Dim strm As TextStream

Dim SndPlayed As Boolean
Dim spdSlow As Boolean
Dim spdMedium As Boolean
Dim spdFast As Boolean
Dim spdUltra As Boolean
Dim optWrap As Boolean
Dim optSelf As Boolean
Dim optSuicide As Boolean
Dim Score As Integer
Dim HiScore As Integer
Dim Lives As Integer
'algebra class all over again
Dim X As Integer
Dim Y As Integer
Dim i As Integer
Dim z As Integer
Dim l As Integer
Dim p As Integer
Dim t As Integer
Dim XXX As Integer
Dim Direction As Integer
Dim CanMove As Boolean
Dim WormNum As Integer
Dim WallNum As Integer
Dim GameOver As Boolean
Dim Bonus As Integer
Dim HiName As String
Dim pAUSEd As Boolean
Dim OrigionAL As Integer

Private Sub Bomb_Timer()
    t = t + 1
    If t = 9 Then
        Itzabomb.Visible = False
        Bomb.Interval = 0
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'Can you move yet?
    If CanMove = False Then Exit Sub
    
    'now if you put suicide on... then you can run into
    'yourself by pressing the opposite direction you are
    'going... but if you do turn it on, you get a bonus
    'in score every time you get a powerup!!
    
    'pauses when you press spacebar
    If pAUSEd = False Then
        If KeyCode = 32 Then tmrWorm.Interval = 0
        pAUSEd = True
    Else
        If KeyCode = 32 Then tmrWorm.Interval = OrigionAL
        pAUSEd = False
    End If
     
    If optSuicide = True Then
        If KeyCode = 38 Then Direction = 0
        If KeyCode = 40 Then Direction = 1
        If KeyCode = 37 Then Direction = 2
        If KeyCode = 39 Then Direction = 3
        Exit Sub
    End If
    
    'This prevents you going the opposite direction you are going (Suicide)
    If optSuicide = False Then
        If Direction <> 1 Then
            If KeyCode = 38 Then Direction = 0
            CanMove = False
        End If
        If Direction <> 0 Then
            If KeyCode = 40 Then Direction = 1
            CanMove = False
        End If
        If Direction <> 3 Then
            If KeyCode = 37 Then Direction = 2
            CanMove = False
        End If
        If Direction <> 2 Then
            If KeyCode = 39 Then Direction = 3
            CanMove = False
        End If
    End If
End Sub

Private Sub Form_Load()
    'file system object, load-save stuff... you have to go
    'to project (the menu bar), click references and check
    'Microsoft Scripting Runtime to make this work.
    
2 'the places it goes to
    
    'if there is no nibbles.jas then it
    'will create one with Jason Liang (me) at 5000
    On Error GoTo 1
    
    'loads the highscore from nibbles.jas
    Set strm = fso.OpenTextFile(App.Path & "\nibbles.jas", ForReading)
    With strm
       HiName = .ReadLine
       HiScore = .ReadLine
       .Close
    End With
    'writes the highscores ands tuff
    lblHighScore.Caption = "Highscore: " & HiName & " " & HiScore
    Score = 0 'sets the scores and lives
    Lives = 3
    tmrQuit.Interval = 1 'moves the X
    WormNum = 5 'there are 6 nibbleparts (0, 1, 2, 3, 4, 5)
    Call FormOnTop(Me.hWnd, True) 'makes the form on top
    Direction = 0 'sets the direction to Up
    Exit Sub
1 'the place that creates a nibbles.jas
    Set strm = fso.CreateTextFile(App.Path & "\nibbles.jas", True)
    With strm
        .WriteLine "Jason Liang"
        .WriteLine "5000"
        .Close
    End With
    GoTo 2 'goes to 2
End Sub

'if your mouse is on the form then reset the buttons
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ResetButtons
End Sub

'loads the options screen
Private Sub Options_Click()
    PlaySnd (theOptions) 'plays the sound
    Load frmOptions
    Call FormOnTop(Me.hWnd, False)
End Sub

'makes the Options button turn read when your mouse is over it
Private Sub Options_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If SndPlayed = False Then
        PlaySnd (TheButton) 'plays the sound
        SndPlayed = True
    End If
    Options.ForeColor = vbRed
    Shape3.BorderColor = vbRed
End Sub

Private Sub Quit_Click()
    End 'quits the program
End Sub

'makes the X turn red when your mouse it over it
Private Sub Quit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If SndPlayed = False Then
        PlaySnd (TheButton) 'plays the sound
        SndPlayed = True
    End If
    Quit.ForeColor = vbRed
    Shape2.BorderColor = vbRed
End Sub

'starts the game
Private Sub Start_Click()
    'makes all the buttons go away
    PlaySnd (goStart) 'plays the sound
    Start.Visible = False
    Shape1.Visible = False
    Shape3.Visible = False
    Options.Visible = False
    Stuff.Visible = False
    'if you didn't select a speed then it's automatically medium!
    If spdSlow = False And spdMedium = False And spdFast = False And spdUltra = False Then spdMedium = True
    'loads the nibbleparts for action
    For i = 1 To WormNum
        Load Nibblepart(i)
        l = i - 1
        Nibblepart(i).Top = Nibblepart(l).Top + 135
    Next i
    GetReady
End Sub

'makes the start button turn red when your mouse is over it
Private Sub Start_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If SndPlayed = False Then
        PlaySnd (TheButton) 'plays the sound
        SndPlayed = True
    End If
    Start.ForeColor = vbRed
    Shape1.BorderColor = vbRed
End Sub

'flashes the getready thing 3 times
Private Sub tmrGetReady_Timer()
    'now you see me... now you don't
    If lblReady.Visible = False Then
        lblReady.Visible = True
    Else
        lblReady.Visible = False
    End If
    z = z + 1
    'after 3 times it says get ready
    '(z=6 because everytime get ready disappears,
    'it counts as one too!)
    If z = 6 Then
        tHeGAmE 'calls tHeGAmE sub
        'if the game is over then the worm stops!
        If GameOver = True Then tmrWorm.Interval = 0
        tmrGetReady.Interval = 0
    End If
End Sub

'this moves the little X quit thing
Private Sub tmrQuit_Timer()
    Quit.Left = Quit.Left + 10
    Shape2.Left = Quit.Left
    If Quit.Left = 3960 Then tmrQuit.Interval = 0
End Sub

'this turns all the buttons green again!
Private Sub ResetButtons()
    Shape1.BorderColor = vbGreen
    Shape2.BorderColor = vbGreen
    Shape3.BorderColor = vbGreen
    Quit.ForeColor = vbGreen
    Start.ForeColor = vbGreen
    Options.ForeColor = vbGreen
    SndPlayed = False
End Sub

'gets the options from the options menu!
Public Sub GetOptions(Slow As Boolean, Medium As Boolean, Fast As Boolean, Ultra As Boolean, Wrap As Boolean, Self As Boolean, Suicide As Boolean)
    If Slow = True Then spdSlow = True
    If Medium = True Then spdMedium = True
    If Fast = True Then spdFast = True
    If Ultra = True Then spdUltra = True
    If Wrap = True Then optWrap = True
    If Self = True Then optSelf = True
    If Suicide = True Then optSuicide = True
End Sub

Private Sub GetReady()
    If GameOver = False Then lblHighScore.Visible = False
    If Lives = 0 Then GameBeOver
    For i = 0 To WormNum
        Nibblepart(i).Left = 2160
        Nibblepart(i).Top = 3240
    Next i
    tmrGetReady.Interval = 500
End Sub

'this sets everything and makes sure everything is right
'before the game actually starts.
Private Sub tHeGAmE()
    'make the labels fore lives and score visible...
    Life.Visible = True
    Life.Caption = "Lives: " & Lives
    lblScore.Visible = True
    lblScore.Caption = Score
    'randomly places the first wall
    TheWall(0).Visible = True
    Randomize
    X = Rnd * 31
    Y = Rnd * 24
    Y = Y * 135
    X = X * 135
    TheWall(0).Left = X
    TheWall(0).Top = Y
    'now you can move!
    CanMove = True
    'your going up
    Direction = 0
    'how fast and a bonus for how fast you are going
    If spdSlow = True Then
        tmrWorm.Interval = 250
        Bonus = 0
        OrigionAL = 250
    End If
    If spdMedium = True Then
        tmrWorm.Interval = 100
        Bonus = 25
        OrigionAL = 100
    End If
    If spdFast = True Then
        tmrWorm.Interval = 60
        Bonus = 50
        OrigionAL = 60
    End If
    If spdUltra = True Then
        tmrWorm.Interval = 30
        Bonus = 100
        OrigionAL = 30
    End If
    For i = 0 To WormNum
        Nibblepart(i).Visible = True
    Next i
    'bonuses for using wrap or self or suicide.
    'notice that wrap and self lowers your bonus
    If optWrap = True Then Bonus = Bonus - 25
    If optSelf = True Then Bonus = Bonus - 50
    If optSuicide = True Then Bonus = Bonus + 30
    'stops the game when the game is over
    If GameOver = True Then tmrWorm.Interval = 0
    CreatePowerup
End Sub

Private Sub CreatePowerup()
1
    'random numbers that determine
    'where the powerup will be
    Randomize
    X = Rnd * 31
    Y = Rnd * 24
    Y = Y * 135
    X = X * 135
    Randomize
    
    'makes it so that it doesn't overlap anything
    For i = 0 To WallNum
        If X = TheWall(i).Left And Y = TheWall(i).Top Then GoTo 1
    Next i
    
    For i = 0 To WormNum
        If X = Nibblepart(i).Left And Y = Nibblepart(i).Top Then GoTo 1
    Next i
    
    'random type of powerup
    XXX = Rnd * 2
    If XXX = 0 Then
        Target.Left = X
        Target.Top = Y
        Target.Visible = True
        Target.ZOrder (front)
    End If
    If XXX = 1 Then
        Wall.Left = X
        Wall.Top = Y
        Wall.Visible = True
        Wall.ZOrder (front)
    End If
    If XXX = 2 Then
        Special.Left = X
        Special.Top = Y
        Special.Visible = True
        Special.ZOrder (front)
    End If
End Sub

'the timer that runs the game
Private Sub tmrWorm_Timer()
    CanMove = True
    Y = Nibblepart(0).Top
    X = Nibblepart(0).Left
    
    'set the lives and score captions...
    Life.Caption = "Lives: " & Lives
    lblScore.Caption = Score
    
    'which way are you going?
    Select Case Direction
        Case 0:
            Y = Y - 135
        Case 1:
            Y = Y + 135
        Case 2:
            X = X - 135
        Case 3:
            X = X + 135
    End Select

    'hey look at that! you're moving!
    Nibblepart(0).Top = Y
    Nibblepart(0).Left = X
    
    'randomly creates a bomb
    Randomize
    p = Rnd * 499
    If p = 71 Then
        CreateBomb
        Bomb.Interval = 1000
        t = 0
    End If
    
    'if you run into the bomb then...
    If Nibblepart(0).Top = Itzabomb.Top And Nibblepart(0).Left = Itzabomb.Left And Itzabomb.Visible = True Then
        Itzabomb.Visible = False
        Score = Score + 200
        PlaySnd (Explosion)
        For i = 0 To 4
            If WormNum = 1 Then
                WormNum = WormNum + 1
                Load Nibblepart(WormNum)
            End If
            Unload Nibblepart(WormNum)
            WormNum = WormNum - 1
        Next i
    End If
    
    'don't hit the walls...
    For i = 0 To WallNum
        If X = TheWall(i).Left And Y = TheWall(i).Top Then
            PlaySnd (Die) 'plays the sound
            Lives = Lives - 1
            If Lives = 0 Then GameBeOver
            GetReady 'the getready flash thing
            tmrWorm.Interval = 0 'stops the timer
            Target.Visible = False 'removes these powerups
            Wall.Visible = False
            Special.Visible = False
            z = 0
        End If
    Next i
    
    'so you can't go thru your self!
    If optSelf = False Then
        For i = 1 To WormNum
            If X = Nibblepart(i).Left And Y = Nibblepart(i).Top Then
                PlaySnd (Die) 'plays the sound
                Lives = Lives - 1
                If Lives = 0 Then GameBeOver
                GetReady 'the getready flash thing
                tmrWorm.Interval = 0 'stops the timer
                Target.Visible = False 'removes these powerups
                Wall.Visible = False
                Special.Visible = False
                z = 0
            End If
        Next i
    End If
    
    'this is to make it NOT wrap... so you DIE when you hit the sides
    If optWrap = False Then
        If Nibblepart(0).Left < 0 Then
            PlaySnd (Die) 'plays the sound
            Lives = Lives - 1
            If Lives = 0 Then GameBeOver
            GetReady 'the getready flash thing
            tmrWorm.Interval = 0 'stops the timer
            Target.Visible = False 'removes these powerups
            Wall.Visible = False
            Special.Visible = False
            z = 0
        End If
        If Nibblepart(0).Top < 0 Then
            PlaySnd (Die) 'plays the sound
            Lives = Lives - 1
            If Lives = 0 Then GameBeOver
            GetReady 'the getready flash thing
            tmrWorm.Interval = 0 'stops the timer
            Target.Visible = False 'removes these powerups
            Wall.Visible = False
            Special.Visible = False
            z = 0
        End If
        If Nibblepart(0).Left > 4200 Then
            PlaySnd (Die) 'plays the sound
            Lives = Lives - 1
            If Lives = 0 Then GameBeOver
            GetReady 'the getready flash thing
            tmrWorm.Interval = 0 'stops the timer
            Target.Visible = False 'removes these powerups
            Wall.Visible = False
            Special.Visible = False
            z = 0
        End If
        If Nibblepart(0).Top > 3240 Then
            PlaySnd (Die) 'plays the sound
            Lives = Lives - 1
            If Lives = 0 Then GameBeOver
            GetReady 'the getready flash thing
            tmrWorm.Interval = 0 'stops the timer
            Target.Visible = False 'removes these powerups
            Wall.Visible = False
            Special.Visible = False
            z = 0
        End If
    End If
    
    Dim q As Integer
    'wrap around the screen and makes sure it is to grid
    If optWrap = True Then
        If Nibblepart(0).Top > 3240 Then
            Nibblepart(0).Top = 0
            q = Nibblepart(0).Left
            q = q / 135
            q = q * 135
            Nibblepart(0).Left = q
        End If
        If Nibblepart(0).Top < 0 Then
            Nibblepart(0).Top = 3240
            q = Nibblepart(0).Left
            q = q / 135
            q = q * 135
            Nibblepart(0).Left = q
        End If
        If Nibblepart(0).Left > 4200 Then
            Nibblepart(0).Left = 0
            q = Nibblepart(0).Top
            q = q / 135
            q = q * 135
            Nibblepart(0).Top = q
        End If
        If Nibblepart(0).Left < 0 Then
            Nibblepart(0).Left = 4200
            q = Nibblepart(0).Top
            q = q / 135
            q = q * 135
            Nibblepart(0).Top = q
            q = Nibblepart(0).Left
            q = q / 135
            q = q * 135
            Nibblepart(0).Left = q
        End If
    End If
    
    'determine the kind of power up you're lookin for
    If XXX = 0 Then
        If Nibblepart(0).Top = Target.Top And Nibblepart(0).Left = Target.Left Then
            PlaySnd (Food) 'plays the sound
            WormNum = WormNum + 1 'the worm grows!
            Load Nibblepart(WormNum)
            Nibblepart(WormNum).Visible = True
            Target.Visible = False
            CreatePowerup 'creates a new powerup
            Score = Score + 50 + Bonus 'adds the score...
        End If
    End If
    
    If XXX = 1 Then
        If Nibblepart(0).Top = Wall.Top And Nibblepart(0).Left = Wall.Left Then
            PlaySnd (Food) 'plays the sound
            WormNum = WormNum + 1 'the worm grows!
            Load Nibblepart(WormNum)
            Nibblepart(WormNum).Visible = True
            Wall.Visible = False
            CreatePowerup 'creates a new powerup
            CreateWall 'creates a wall (yes the red one create walls)
            Score = Score + 75 + Bonus 'adds the score...
        End If
    End If
    If XXX = 2 Then
        If Nibblepart(0).Top = Special.Top And Nibblepart(0).Left = Special.Left Then
            PlaySnd (Food) 'plays the sound
            WormNum = WormNum + 1 'the worm grows!
            Load Nibblepart(WormNum)
            Nibblepart(WormNum).Visible = True
            Special.Visible = False
            Score = Score + 100 + Bonus 'adds the score...
            'makes you faster but unfortunately, vb slows
            'down when there are more nibbleparts so...
            'this kind of keeps it at the same speed
            If tmrWorm.Interval > 0 Then
                tmrWorm.Interval = tmrWorm.Interval - 1
            End If
            CreatePowerup 'creates a new powerup
        End If
    End If
    
    'this makes the worm move!
    For i = 1 To WormNum
        l = i + 1
        If l > WormNum Then l = 0
        Nibblepart(i).Left = Nibblepart(l).Left
        Nibblepart(i).Top = Nibblepart(l).Top
    'this makes the nibblepart on top
        Nibblepart(i).ZOrder (front)
    Next i
End Sub

Private Sub CreateBomb()
1
    'random numbers for the position... blah blah blah
    Randomize
    X = Rnd * 31
    Y = Rnd * 24
    Y = Y * 135
    X = X * 135

    'makes sure the bomb doesn't overlap anything...
    For i = 0 To WallNum
        If X = TheWall(i).Left And Y = TheWall(i).Top Then GoTo 1
    Next i
    
    For i = 0 To WormNum
        If X = Nibblepart(i).Left And Y = Nibblepart(i).Top Then GoTo 1
    Next i
    
    If X = Target.Left And Y = Target.Top Then GoTo 1
    If X = Special.Left And Y = Special.Top Then GoTo 1
    If X = Wall.Left And Y = Wall.Top Then GoTo 1
    
    Itzabomb.Left = X
    Itzabomb.Top = Y
    Itzabomb.Visible = True
End Sub

'creating the walls!
Private Sub CreateWall()
1
    'random numbers for the position... blah blah blah
    Randomize
    X = Rnd * 31
    Y = Rnd * 24
    Y = Y * 135
    X = X * 135
    
    'makes sure the wall doesn't overlap anything...
    For i = 0 To WallNum
        If X = TheWall(i).Left And Y = TheWall(i).Top Then GoTo 1
    Next i
    
    For i = 0 To WormNum
        If X = Nibblepart(i).Left And Y = Nibblepart(i).Top Then GoTo 1
    Next i
    
    If X = Target.Left And Y = Target.Top Then GoTo 1
    If X = Special.Left And Y = Special.Top Then GoTo 1
    If X = Wall.Left And Y = Wall.Top Then GoTo 1
    
    'hey! it's a new wall!
    WallNum = WallNum + 1
    Load TheWall(WallNum)
    TheWall(WallNum).Left = X
    TheWall(WallNum).Top = Y
    TheWall(WallNum).Visible = True
    
    'brings the wall to the front
    TheWall(WallNum).ZOrder (front)
End Sub

Private Sub GameBeOver()
    'GAME OVA!
    GameOver = True
    tmrWorm.Interval = 0
    'flashes the game over sign
    lblReady.Caption = "Game Over"
    lblReady.Visible = True
    'determines high score
    If Score > HiScore Then
        'makes the form not on top
        Call FormOnTop(Me.hWnd, False)
        InputBox "What is your name?", HiName 'your name?
        'writes the name and the highscore in nibbles.jas... clever eh?
        Set strm = fso.CreateTextFile(App.Path & "\nibbles.jas", True)
        With strm
            .WriteLine HiName 'writes the name
            .WriteLine Score 'writes the score
            .Close 'closes the textstream
        End With
        Call FormOnTop(Me.hWnd, True)
        'makes the form on top again
    End If
    'reads the high score...
    Set strm = fso.OpenTextFile(App.Path & "\nibbles.jas", ForReading)
    With strm
       HiName = .ReadLine 'read HiName
       HiScore = .ReadLine 'read HiScore
       .Close 'closes the textstream
    End With
    'makes the highscore visible and displays the name and score!
    lblHighScore.Visible = True
    lblHighScore.Caption = "Highscore: " & HiName & " " & HiScore
End Sub
