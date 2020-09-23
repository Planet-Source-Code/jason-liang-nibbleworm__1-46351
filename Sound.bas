Attribute VB_Name = "Sound"
'Call API
Public Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Public Const SND_ASYNC = &H1

'wHaTSoUnD
Public Enum wHaTSoUnD
    TheButton = 1
    Check = 2
    Food = 3
    goStart = 4
    theOptions = 5
    Die = 6
    Explosion = 7
End Enum

'The sub that plays the sounds
Public Sub PlaySnd(Soundtype As wHaTSoUnD)
    Select Case Soundtype
        Case TheButton
            Call PlaySound(App.Path + "\btnover.wav", 0, SND_ASYNC)
        Case Check
            Call PlaySound(App.Path + "\check.wav", 0, SND_ASYNC)
        Case Food
            Call PlaySound(App.Path + "\food.wav", 0, SND_ASYNC)
        Case goStart
            Call PlaySound(App.Path + "\start.wav", 0, SND_ASYNC)
        Case theOptions
            Call PlaySound(App.Path + "\options.wav", 0, SND_ASYNC)
        Case Die
            Call PlaySound(App.Path + "\die!.wav", 0, SND_ASYNC)
        Case Explosion
            Call PlaySound(App.Path + "\explosion.wav", 0, SND_ASYNC)
    End Select
End Sub

'What a useful little module!
