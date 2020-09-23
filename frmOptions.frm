VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   0  'None
   Caption         =   "Options"
   ClientHeight    =   2505
   ClientLeft      =   6540
   ClientTop       =   4515
   ClientWidth     =   2175
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2505
   ScaleWidth      =   2175
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox Suicide 
      Caption         =   "Suicide"
      Height          =   195
      Left            =   0
      TabIndex        =   9
      Top             =   2280
      Width           =   1575
   End
   Begin VB.CheckBox Self 
      Caption         =   "Run thru self?"
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   2040
      Width           =   1455
   End
   Begin VB.CommandButton OK 
      Caption         =   "OK"
      Height          =   255
      Left            =   1200
      TabIndex        =   7
      Top             =   720
      Width           =   975
   End
   Begin VB.CheckBox Wrap 
      Caption         =   "Wrap"
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   1800
      Width           =   1455
   End
   Begin VB.OptionButton Ultra 
      Caption         =   "Ultra"
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   1080
      Width           =   1095
   End
   Begin VB.OptionButton Fast 
      Caption         =   "Fast"
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   840
      Width           =   1095
   End
   Begin VB.OptionButton Medium 
      Caption         =   "Medium"
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   600
      Value           =   -1  'True
      Width           =   1095
   End
   Begin VB.OptionButton Slow 
      Caption         =   "Slow"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Misc 
      Alignment       =   2  'Center
      Caption         =   "Misc"
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   1440
      Width           =   2175
   End
   Begin VB.Label Speeds 
      Alignment       =   2  'Center
      Caption         =   "Speeds"
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   2175
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'look familiar?
Dim spdFast As Boolean
Dim spdSlow As Boolean
Dim spdMedium As Boolean
Dim spdUltra As Boolean
Dim optWrap As Boolean
Dim optSelf As Boolean
Dim optSuicide As Boolean

'if you click this, spdFast will be true
Private Sub Fast_Click()
    PlaySnd (Check) 'plays the sound
    MakeFalse 'this makes all the others false
    spdFast = True
End Sub

Private Sub Form_Load()
    Call FormOnTop(Me.hWnd, True) 'makes the form on top
    'makes the menu the way yes left it
    If spdSlow = True Then Slow.Value = True
    If spdMedium = True Then Medium.Value = True
    If spdFast = True Then Fast.Value = True
    If spdUltra = True Then Ultra.Value = True
    If optWrap = True Then Wrap.Value = 1
    If optSelf = True Then Self.Value = 1
    If optSuicide = True Then Suicide.Value = 1
    
    'makes the little thing that tells you what it does if u put
    'your mouse over it
    Slow.ToolTipText = "Slow Speed, no bonus"
    Medium.ToolTipText = "Medium Speed, 25 bonus"
    Fast.ToolTipText = "Fast Speed, 50 bonus"
    Ultra.ToolTipText = "Ultra Fast! ??? bonus!"
    Wrap.ToolTipText = "Wraps around screen, -25 bonus"
    Self.ToolTipText = "You can go through yourself, -50 bonus"
    Suicide.ToolTipText = "If you accidentally press the opposite direction you are going, you are dead! 30 bonus"
    OK.ToolTipText = "Save settings and go back to the game screen"
End Sub

'if you click this, spdMedium will be true
Private Sub Medium_Click()
    PlaySnd (Check) 'plays the sound
    MakeFalse 'makes everything else false
    spdMedium = True
End Sub

'sends all the info to GetOptions in the form Game
Private Sub OK_Click()
    Call Game.GetOptions(spdSlow, spdMedium, spdFast, spdUltra, optWrap, optSelf, optSuicide)
    Call FormOnTop(Game.hWnd, True)
    Unload Me
End Sub

'if you click this, optSelf will be true
Private Sub Self_Click()
    PlaySnd (Check) 'plays the sound
    If optSelf = False Then
        optSelf = True
    Else
        optSelf = False
    End If
End Sub

'if you click this, spdSlow will be true
Private Sub Slow_Click()
    PlaySnd (Check) 'plays the sound
    MakeFalse
    spdSlow = True
End Sub

'makes every speed false
Private Sub MakeFalse()
    spdSlow = False
    spdMedium = False
    spdFast = False
    spdUltra = False
End Sub

'if you click this, optSuicide will be true
Private Sub Suicide_Click()
    PlaySnd (Check) 'plays the sound
    If optSuicide = False Then
        optSuicide = True
    Else
        optSuicide = False
    End If
End Sub

'if you click this, spdUltra will be true
Private Sub Ultra_Click()
    PlaySnd (Check) 'plays the sound
    MakeFalse 'makes every other speed false
    spdUltra = True
End Sub

'if you click this, optWrap will be true
Private Sub Wrap_Click()
    PlaySnd (Check) 'plays the sound
    If optWrap = False Then
        optWrap = True
    Else
        optWrap = False
    End If
End Sub
