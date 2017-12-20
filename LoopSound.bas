Attribute VB_Name = "LoopSound"
Option Compare Database

Public Declare Function sndPlaySound32 _
    Lib "winmm.dll" _
    Alias "sndPlaySoundA" ( _
        ByVal lpszSoundName As String, _
        ByVal uFlags As Long) As Long

Sub playSound(ByVal whatSound As String, Optional Flags As Long = 0)
    If Dir(whatSound, vbNormal) = "" Then
        ' WhatSound is not a file. Get the file named by
        ' WhatSound from the Windows\Media directory.
        whatSound = Environ("SystemRoot") & "\Media\" & whatSound
        If InStr(1, whatSound, ".") = 0 Then
            ' if WhatSound does not have a .wav extension,
            ' add one.
            whatSound = whatSound & ".wav"
        End If
        If Dir(whatSound, vbNormal) = vbNullString Then
            ' Can't find the file. Do a simple Beep.
            Beep
            Exit Sub
        End If
    Else
        ' WhatSound is a file. Use it.
    End If
    ' Finally, play the sound.
    sndPlaySound32 whatSound, Flags
End Sub

Sub loopTimer()

Dim startTime, waitTime, loopTime

Dim n As Integer '  set n to limit the maximum number of loops

n = 30
waitTime = 10   ' waittime in seconds between code execution
startTime = Round(Timer, 0)
loopTime = Round(Timer, 0)
Do While Round(Timer, 0) <= loopTime + waitTime
    DoEvents
    If Round(Timer, 0) = loopTime + waitTime Then
        DoEvents    ' to allow breaks in the code if accidental infinite loop is enabled
        Debug.Print Timer
        playSound "Critical Error", &H1
           loopTime = Round(Timer, 0)
           'If Timer > startTime + n * waitTime Then Exit Do '     loop limiter, adjust variable n to set number of loops
    End If
Loop
    
End Sub

