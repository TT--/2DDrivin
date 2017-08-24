Attribute VB_Name = "Module1"
Option Explicit

Global Const numofcars As Integer = 2    '1 is LEFT 2 is RIGHT

Public Type CarStructure
    number As Integer   'car number
    stopped As Boolean 'done racing?
    Dir As Integer    'direction - 1 to 36
    tlxpos As Single  'player top left x position
    tlypos As Single 'player y position
    cnxpos As Single   'player centre x position
    cnypos As Single 'player centre y position
    speed As Single  'cars speed
    bounce As Single  'cars bounce speed
    txpos As Single  'temp x pos for collisions
    typos As Single  'temp y
    righton As Byte
    lefton As Byte
    upon As Byte
    downon As Byte
    waypoint(1 To 4) As Integer
    score As Integer
    time As String
End Type

Global car(1 To numofcars) As CarStructure

Global speedup, speedmax, speedmin, laps As Integer
Global firstdone, sound As Boolean
Global Const spritewidth = 45
Global Const spriteheight = 42

Global Trackname, Midiname, winnertime, losertime As String
Global winnernum, losernum, randomtrack, randommidi, counter As Integer
Global TotalTenthSeconds, TotalSeconds, TenthSeconds, Seconds, Minutes As Integer
Global PixCol, PixCar, PixTrk As Long

'API DECLARATIONS
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function HideCaret Lib "user32" (ByVal hwnd As Long) As Long
Declare Function mciSendString Lib "winmm.dll" Alias _
"mciSendStringA" (ByVal lpstrCommand As String, _
ByVal lpstrReturnString As String, ByVal uReturnLength _
As Long, ByVal hwndCallback As Long) As Long

'FOR BITBLT
Public Const SRCAND = &H8800C6  ' (DWORD) dest = source AND dest
Public Const SRCCOPY = &HCC0020 ' (DWORD) dest = source
Public Const SRCPAINT = &HEE0086        ' (DWORD) dest = source OR dest

Public Sub StopMidi(MidiFileName As String)
Call mciSendString("stop " + MidiFileName, 0&, 0, 0)
Call mciSendString("close " + MidiFileName, 0&, 0, 0)
End Sub

Public Function PlayMIDI(MidiFileName As String)
On Error Resume Next
Call mciSendString("open " + MidiFileName + " type sequencer", 0&, 0, 0)
If mciSendString("play " + MidiFileName + "", 0&, 0, 0) = 0 Then
PlayMIDI = 0
Else
PlayMIDI = 1
End If
End Function

Public Sub clearall()
firstdone = False

'initialize waypoints and speed
Dim z As Integer
Dim Y As Integer
For z = 1 To numofcars
For Y = 1 To 4
car(z).waypoint(Y) = 0
car(z).speed = 0
car(z).lefton = 0
car(z).righton = 0
car(z).upon = 0
car(z).downon = 0
Next Y
Next z

'specific track settings
If Trackname = "Track1" Then
car(1).Dir = 10
car(2).Dir = 10
car(1).tlxpos = 500
car(1).tlypos = 30
car(2).tlxpos = 500
car(2).tlypos = 60
ElseIf Trackname = "Track2" Then
car(1).Dir = 10
car(2).Dir = 10
car(1).tlxpos = 380
car(1).tlypos = 495
car(2).tlxpos = 380
car(2).tlypos = 525
ElseIf Trackname = "Track3" Then
car(1).Dir = 28
car(2).Dir = 28
car(1).tlxpos = 140
car(1).tlypos = 490
car(2).tlxpos = 140
car(2).tlypos = 520
ElseIf Trackname = "Track4" Then
car(1).Dir = 1
car(2).Dir = 1
car(1).tlxpos = 20
car(1).tlypos = 405
car(2).tlxpos = 45
car(2).tlypos = 405
End If
End Sub

Public Sub movecars()
Dim a As Integer
  For a = 1 To numofcars
    If car(a).Dir = 19 Then  'down
      car(a).tlypos = car(a).tlypos - car(a).speed / 20 - (car(a).bounce / 20)
    End If
    If car(a).Dir = 20 Then
      car(a).tlypos = car(a).tlypos - car(a).speed / 22 - (car(a).bounce / 22)
      car(a).tlxpos = car(a).tlxpos + car(a).speed / 100 + (car(a).bounce / 100)
    End If
    If car(a).Dir = 21 Then
      car(a).tlypos = car(a).tlypos - car(a).speed / 25 - (car(a).bounce / 25)
      car(a).tlxpos = car(a).tlxpos + car(a).speed / 60 + (car(a).bounce / 60)
    End If
     
     If car(a).Dir = 22 Then
      car(a).tlypos = car(a).tlypos - car(a).speed / 27 - (car(a).bounce / 27)
      car(a).tlxpos = car(a).tlxpos + car(a).speed / 50 + (car(a).bounce / 50)
    End If
    If car(a).Dir = 23 Then
      car(a).tlypos = car(a).tlypos - car(a).speed / 29 - (car(a).bounce / 29)
      car(a).tlxpos = car(a).tlxpos + car(a).speed / 35 + (car(a).bounce / 35)
    End If
    If car(a).Dir = 24 Then
      car(a).tlypos = car(a).tlypos - car(a).speed / 35 - (car(a).bounce / 35)
      car(a).tlxpos = car(a).tlxpos + car(a).speed / 29 + (car(a).bounce / 29)
    End If
    If car(a).Dir = 25 Then
      car(a).tlypos = car(a).tlypos - car(a).speed / 50 - (car(a).bounce / 50)
      car(a).tlxpos = car(a).tlxpos + car(a).speed / 27 + (car(a).bounce / 27)
    End If
    If car(a).Dir = 26 Then
      car(a).tlypos = car(a).tlypos - car(a).speed / 60 - (car(a).bounce / 60)
      car(a).tlxpos = car(a).tlxpos + car(a).speed / 25 + (car(a).bounce / 25)
    End If
    If car(a).Dir = 27 Then
      car(a).tlypos = car(a).tlypos - car(a).speed / 100 - (car(a).bounce / 100)
      car(a).tlxpos = car(a).tlxpos + car(a).speed / 22 + (car(a).bounce / 22)
    End If
    If car(a).Dir = 28 Then
      car(a).tlxpos = car(a).tlxpos + car(a).speed / 20 + (car(a).bounce / 20)
    End If
    If car(a).Dir = 29 Then
      car(a).tlypos = car(a).tlypos + car(a).speed / 100 + (car(a).bounce / 100)
      car(a).tlxpos = car(a).tlxpos + car(a).speed / 22 + (car(a).bounce / 22)
    End If
    If car(a).Dir = 30 Then
      car(a).tlypos = car(a).tlypos + car(a).speed / 60 + (car(a).bounce / 60)
      car(a).tlxpos = car(a).tlxpos + car(a).speed / 25 + (car(a).bounce / 25)
    End If
    If car(a).Dir = 31 Then
      car(a).tlypos = car(a).tlypos + car(a).speed / 50 + (car(a).bounce / 50)
      car(a).tlxpos = car(a).tlxpos + car(a).speed / 27 + (car(a).bounce / 27)
    End If
    If car(a).Dir = 32 Then
      car(a).tlypos = car(a).tlypos + car(a).speed / 35 + (car(a).bounce / 35)
      car(a).tlxpos = car(a).tlxpos + car(a).speed / 29 + (car(a).bounce / 29)
    End If
    If car(a).Dir = 33 Then
      car(a).tlypos = car(a).tlypos + car(a).speed / 29 + (car(a).bounce / 29)
      car(a).tlxpos = car(a).tlxpos + car(a).speed / 35 + (car(a).bounce / 35)
    End If
    If car(a).Dir = 34 Then
      car(a).tlypos = car(a).tlypos + car(a).speed / 27 + (car(a).bounce / 27)
      car(a).tlxpos = car(a).tlxpos + car(a).speed / 50 + (car(a).bounce / 50)
    End If
    If car(a).Dir = 35 Then
      car(a).tlypos = car(a).tlypos + car(a).speed / 25 + (car(a).bounce / 25)
      car(a).tlxpos = car(a).tlxpos + car(a).speed / 60 + (car(a).bounce / 60)
    End If
    If car(a).Dir = 36 Then
      car(a).tlypos = car(a).tlypos + car(a).speed / 22 + (car(a).bounce / 22)
      car(a).tlxpos = car(a).tlxpos + car(a).speed / 100 + (car(a).bounce / 100)
    End If
    
    If car(a).Dir = 1 Then  'straight up
      car(a).tlypos = car(a).tlypos + car(a).speed / 20 + (car(a).bounce / 20)
    End If
      If car(a).Dir = 2 Then
      car(a).tlypos = car(a).tlypos + car(a).speed / 22 + (car(a).bounce / 22)
      car(a).tlxpos = car(a).tlxpos - car(a).speed / 100 - (car(a).bounce / 100)
    End If
    If car(a).Dir = 3 Then
      car(a).tlypos = car(a).tlypos + car(a).speed / 25 + (car(a).bounce / 25)
      car(a).tlxpos = car(a).tlxpos - car(a).speed / 60 - (car(a).bounce / 60)
    End If
    If car(a).Dir = 4 Then
      car(a).tlypos = car(a).tlypos + car(a).speed / 27 + (car(a).bounce / 27)
      car(a).tlxpos = car(a).tlxpos - car(a).speed / 50 - (car(a).bounce / 50)
    End If
    If car(a).Dir = 5 Then
      car(a).tlypos = car(a).tlypos + car(a).speed / 29 + (car(a).bounce / 29)
      car(a).tlxpos = car(a).tlxpos - car(a).speed / 35 - (car(a).bounce / 35)
    End If
    If car(a).Dir = 6 Then
      car(a).tlypos = car(a).tlypos + car(a).speed / 35 + (car(a).bounce / 35)
      car(a).tlxpos = car(a).tlxpos - car(a).speed / 29 - (car(a).bounce / 29)
    End If
    If car(a).Dir = 7 Then
      car(a).tlypos = car(a).tlypos + car(a).speed / 50 + (car(a).bounce / 50)
      car(a).tlxpos = car(a).tlxpos - car(a).speed / 27 - (car(a).bounce / 27)
    End If
    If car(a).Dir = 8 Then
      car(a).tlypos = car(a).tlypos + car(a).speed / 60 + (car(a).bounce / 60)
      car(a).tlxpos = car(a).tlxpos - car(a).speed / 25 - (car(a).bounce / 25)
    End If
    If car(a).Dir = 9 Then
      car(a).tlypos = car(a).tlypos + car(a).speed / 100 + (car(a).bounce / 100)
      car(a).tlxpos = car(a).tlxpos - car(a).speed / 22 - (car(a).bounce / 22)
    End If
    If car(a).Dir = 10 Then       'LEFT
      car(a).tlxpos = car(a).tlxpos - car(a).speed / 20 - (car(a).bounce / 20)
    End If
    If car(a).Dir = 11 Then
      car(a).tlypos = car(a).tlypos - car(a).speed / 100 - (car(a).bounce / 100)
      car(a).tlxpos = car(a).tlxpos - car(a).speed / 22 - (car(a).bounce / 22)
    End If
    If car(a).Dir = 12 Then
      car(a).tlypos = car(a).tlypos - car(a).speed / 60 - (car(a).bounce / 60)
      car(a).tlxpos = car(a).tlxpos - car(a).speed / 25 - (car(a).bounce / 25)
    End If
    If car(a).Dir = 13 Then
      car(a).tlypos = car(a).tlypos - car(a).speed / 50 - (car(a).bounce / 50)
      car(a).tlxpos = car(a).tlxpos - car(a).speed / 27 - (car(a).bounce / 27)
    End If
    If car(a).Dir = 14 Then
      car(a).tlypos = car(a).tlypos - car(a).speed / 35 - (car(a).bounce / 35)
      car(a).tlxpos = car(a).tlxpos - car(a).speed / 29 - (car(a).bounce / 29)
    End If
    If car(a).Dir = 15 Then
      car(a).tlypos = car(a).tlypos - car(a).speed / 29 - (car(a).bounce / 29)
      car(a).tlxpos = car(a).tlxpos - car(a).speed / 35 - (car(a).bounce / 35)
    End If
    If car(a).Dir = 16 Then
      car(a).tlypos = car(a).tlypos - car(a).speed / 27 - (car(a).bounce / 27)
      car(a).tlxpos = car(a).tlxpos - car(a).speed / 50 - (car(a).bounce / 50)
    End If
    If car(a).Dir = 17 Then
      car(a).tlypos = car(a).tlypos - car(a).speed / 25 - (car(a).bounce / 25)
      car(a).tlxpos = car(a).tlxpos - car(a).speed / 60 - (car(a).bounce / 60)
    End If
    If car(a).Dir = 18 Then
      car(a).tlypos = car(a).tlypos - car(a).speed / 22 - (car(a).bounce / 22)
      car(a).tlxpos = car(a).tlxpos - car(a).speed / 100 - (car(a).bounce / 100)
    End If

'Centres
  car(a).cnxpos = car(a).tlxpos + (spritewidth / 2)
  car(a).cnypos = car(a).tlypos + (spriteheight / 2)
  Next a
End Sub
