Attribute VB_Name = "Sound"

'That is the module Play the different sounds
'when showing the messages, because that messages
'are based on dialog forms, and can't Play a messages
'sound when they show. So, that is the manual way to
'Play the different message sounds, that uses windows
'Media folder to get the WAV file, to Play.

Public Sub Critical(FRM As Form)

'Play file URL (Path)
FRM.WMP1.URL = "C:\WINDOWS\Media\Windows XP Critical Stop.wav"

'Using Windows Media Control based on FRM
FRM.WMP1.Controls.play

End Sub

Public Sub Exclamation(FRM As Form)

'Play file URL (Path)
FRM.WMP1.URL = "C:\WINDOWS\Media\Windows XP Exclamation.wav"

'Using Windows Media Control based on FRM
FRM.WMP1.Controls.play

End Sub

Public Sub Information(FRM As Form)

'Play file URL (Path)
FRM.WMP1.URL = "C:\WINDOWS\Media\Windows XP Error.wav"

'Using Windows Media Control based on FRM
FRM.WMP1.Controls.play

End Sub

Public Sub Ding(FRM As Form)

'Play file URL (Path)
FRM.WMP1.URL = "C:\WINDOWS\Media\Windows XP Ding.wav"

'Using Windows Media Control based on FRM
FRM.WMP1.Controls.play

End Sub
Public Sub Recycle(FRM As Form)

'Play file URL (Path)
FRM.WMP1.URL = "C:\WINDOWS\Media\Recycle.wav"

'Using Windows Media Control based on FRM
FRM.WMP1.Controls.play

End Sub
