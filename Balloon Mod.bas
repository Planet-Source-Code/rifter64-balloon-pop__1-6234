Attribute VB_Name = "Module1"
Declare Function SystemParametersInfo Lib _
"user32" Alias "SystemParametersInfoA" (ByVal uAction _
As Long, ByVal uParam As Long, ByVal lpvParam As Any, _
ByVal fuWinIni As Long) As Long

Declare Function sndPlaySound Lib "winmm" Alias "sndPlaySoundA" _
    (ByVal lpszSoundName As String, ByVal uflags As Long) As Long

Public Popup, Win, Time, High As Integer

