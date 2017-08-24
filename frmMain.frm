VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   Caption         =   "2D Drivin"
   ClientHeight    =   8490
   ClientLeft      =   1155
   ClientTop       =   1860
   ClientWidth     =   11880
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text8 
      Enabled         =   0   'False
      Height          =   495
      Left            =   7320
      TabIndex        =   18
      Text            =   "Text8"
      Top             =   1920
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text7 
      Enabled         =   0   'False
      Height          =   495
      Left            =   7320
      TabIndex        =   17
      Text            =   "Text7"
      Top             =   1320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text6 
      Enabled         =   0   'False
      Height          =   495
      Left            =   7320
      TabIndex        =   16
      Text            =   "Text6"
      Top             =   720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      Enabled         =   0   'False
      Height          =   495
      Left            =   7320
      TabIndex        =   15
      Text            =   "Text5"
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      Enabled         =   0   'False
      Height          =   495
      Left            =   6000
      TabIndex        =   14
      Text            =   "Text4"
      Top             =   1920
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Enabled         =   0   'False
      Height          =   495
      Left            =   6000
      TabIndex        =   13
      Text            =   "Text3"
      Top             =   1320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   495
      Left            =   6000
      TabIndex        =   12
      Text            =   "Text2"
      Top             =   720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   495
      Left            =   6000
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame fraWinner 
      BackColor       =   &H00C00000&
      BorderStyle     =   0  'None
      Height          =   3375
      Left            =   2093
      TabIndex        =   6
      Top             =   2558
      Visible         =   0   'False
      Width           =   7695
      Begin VB.CommandButton cmdDig 
         BackColor       =   &H00008000&
         Caption         =   "Continue"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2760
         MaskColor       =   &H00008000&
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   2760
         Width           =   2055
      End
      Begin VB.Image imgPlayerLose 
         Height          =   735
         Left            =   2760
         Top             =   1080
         Width           =   495
      End
      Begin VB.Image imgPlayerWin 
         Height          =   735
         Left            =   2760
         Top             =   240
         Width           =   495
      End
      Begin VB.Image imgWinner 
         Height          =   735
         Left            =   840
         Top             =   2040
         Width           =   495
      End
      Begin VB.Label lblLoseTime 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3840
         TabIndex        =   9
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label lblWinTime 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3840
         TabIndex        =   8
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label lblWin 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "is the Champion of the Universe"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   495
         Left            =   1560
         TabIndex        =   7
         Top             =   2040
         Width           =   5895
      End
   End
   Begin VB.Frame fraStarter 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   2895
      Left            =   4613
      TabIndex        =   5
      Top             =   2798
      Visible         =   0   'False
      Width           =   2655
      Begin VB.Shape shGr2 
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   1440
         Shape           =   3  'Circle
         Top             =   1800
         Width           =   615
      End
      Begin VB.Shape shGr1 
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   600
         Shape           =   3  'Circle
         Top             =   1800
         Width           =   615
      End
      Begin VB.Shape shYel2 
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   1440
         Shape           =   3  'Circle
         Top             =   1080
         Width           =   615
      End
      Begin VB.Shape shYel1 
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   600
         Shape           =   3  'Circle
         Top             =   1080
         Width           =   615
      End
      Begin VB.Shape shRed2 
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   1440
         Shape           =   3  'Circle
         Top             =   360
         Width           =   615
      End
      Begin VB.Shape shRed1 
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   600
         Shape           =   3  'Circle
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Timer tmrStarter 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   240
      Top             =   1320
   End
   Begin VB.TextBox txtStopWatch 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   0
      Width           =   1095
   End
   Begin VB.Timer tmrStopWatch 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   840
      Top             =   1320
   End
   Begin VB.Timer tmrRefresh 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   240
      Top             =   4200
   End
   Begin VB.Timer tmrTurn 
      Interval        =   35
      Left            =   240
      Top             =   4920
   End
   Begin VB.PictureBox picCarsw 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      Enabled         =   0   'False
      Height          =   1215
      Left            =   4200
      ScaleHeight     =   1155
      ScaleWidth      =   1635
      TabIndex        =   2
      Top             =   960
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.PictureBox picCarsb 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      Enabled         =   0   'False
      Height          =   1215
      Left            =   2400
      ScaleHeight     =   1155
      ScaleWidth      =   1635
      TabIndex        =   1
      Top             =   960
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.PictureBox picCar 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   2055
      Left            =   5280
      ScaleHeight     =   2055
      ScaleWidth      =   3015
      TabIndex        =   0
      Top             =   3360
      Width           =   3015
   End
   Begin VB.PictureBox picTrack 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   2055
      Left            =   2040
      ScaleHeight     =   2055
      ScaleWidth      =   2295
      TabIndex        =   3
      Top             =   3360
      Visible         =   0   'False
      Width           =   2295
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Sub drawcars()
'BitBlt(where to, tlxpos(b), tlypos(b), width on source, height on source, picCarsb.hdc, (CarsDirs(b) - 1) * 45, 45 * (b - 1), SRCPAINT)
Call BitBlt(picCar.hdc, car(1).tlxpos, car(1).tlypos, 45, 42, picCarsb.hdc, (car(1).Dir - 1) * 45, 42 * (car(1).number - 1), SRCPAINT)
Call BitBlt(picCar.hdc, car(1).tlxpos, car(1).tlypos, 45, 42, picCarsw.hdc, (car(1).Dir - 1) * 45, 42 * (car(1).number - 1), SRCAND)

Call BitBlt(picCar.hdc, car(2).tlxpos, car(2).tlypos, 45, 42, picCarsb.hdc, (car(2).Dir - 1) * 45, 42 * (car(2).number - 1), SRCPAINT)
Call BitBlt(picCar.hdc, car(2).tlxpos, car(2).tlypos, 45, 42, picCarsw.hdc, (car(2).Dir - 1) * 45, 42 * (car(2).number - 1), SRCAND)
End Sub

Private Sub cmdDig_Click()
If sound = True Then
randommidi = Int((5 * Rnd + 1))
Midiname = "Midi" + Format(randommidi)
PlayMIDI (App.Path & "\sounds\" & Midiname & ".mid")
End If
Call clearall

Load frmTracks
frmTracks.Visible = True
frmTracks.Enabled = True
Unload frmMain
End Sub

Private Sub Form_Load()

Call clearall
'start the countdown to the race
counter = 4  'should be 4
tmrStarter.Enabled = True
fraStarter.Visible = True
Dim u As Integer
For u = 1 To numofcars
car(u).stopped = False
Next u

frmMain.Width = 800 * 15
frmMain.Height = 600 * 15
frmMain.Top = 0
frmMain.Left = 0

picCar.Width = 800 * 15
picCar.Height = 600 * 15
picCar.Left = 0
picCar.Top = 0

'Set up the screen
picTrack.Picture = LoadPicture(App.Path & "\images\" & Trackname & ".gif")
picCar.Picture = LoadPicture(App.Path & "\images\" & Trackname & ".gif")
picCarsw.Picture = LoadPicture(App.Path & "\images\carsw.bmp")
picCarsb.Picture = LoadPicture(App.Path & "\images\carsb.bmp")

clearall

Call drawcars

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

Select Case KeyCode
 Case 37
    car(2).lefton = 1
 Case 38
    car(2).upon = 1
 Case 39
    car(2).righton = 1
 Case 40
    car(2).downon = 1
 Case 65
    car(1).lefton = 1
 Case 87
    car(1).upon = 1
 Case 68
    car(1).righton = 1
 Case 83
    car(1).downon = 1
 End Select

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
 Case 37
    car(2).lefton = 0
 Case 38
    car(2).upon = 0
 Case 39
    car(2).righton = 0
 Case 40
    car(2).downon = 0
 Case 65
    car(1).lefton = 0
 Case 87
    car(1).upon = 0
 Case 68
    car(1).righton = 0
 Case 83
    car(1).downon = 0
 End Select

End Sub

Private Sub tmrRefresh_Timer()

'SPEED
Dim a, b As Integer

For a = 1 To numofcars
If car(a).stopped = False Then
    If car(a).upon = 1 And car(a).downon = 0 Then
      If car(a).speed < speedmax Then car(a).speed = car(a).speed + speedup
    ElseIf car(a).downon = 1 And car(a).upon = 0 Then
      If car(a).speed > speedmin Then car(a).speed = car(a).speed - speedup
    ElseIf car(a).upon = 0 And car(a).downon = 0 Then
        If car(a).speed > 0 Then
        car(a).speed = car(a).speed - (speedup / 3)
        ElseIf car(a).speed < 0 Then
        car(a).speed = car(a).speed + (speedup / 3)
     End If
    End If
End If
Next a

'WAYPOINT system - prevents cheating, used in scoring
Dim n As Integer
For n = 1 To numofcars
PixCol = GetPixel(picTrack.hdc, car(n).cnxpos, car(n).cnypos)

    If PixCol = 33792 Then  '  <> 8685188 not grey  school:8684676
       'MsgBox (PixCol)
         car(n).speed = car(n).speed / 1.25
    ElseIf PixCol = 16711680 Then         'blue
        If car(n).waypoint(1) = car(n).waypoint(4) Then
        car(n).waypoint(1) = car(n).waypoint(1) + 1
        End If
    ElseIf PixCol = 65280 Then        'green
        If car(n).waypoint(2) <> car(n).waypoint(1) Then
        car(n).waypoint(2) = car(n).waypoint(2) + 1
        End If
    ElseIf PixCol = 16711935 Then      'pink
        If car(n).waypoint(1) = car(n).waypoint(2) And car(n).waypoint(2) <> car(n).waypoint(3) Then
        car(n).waypoint(3) = car(n).waypoint(3) + 1
        End If
    ElseIf PixCol = 65535 Then         'yellow
        If car(n).waypoint(2) = car(n).waypoint(3) And car(n).waypoint(3) <> car(n).waypoint(4) Then
        car(n).waypoint(4) = car(n).waypoint(4) + 1
        End If
        End If

Text1 = PixCol
Text2 = car(1).waypoint(2)
Text3 = car(1).waypoint(3)
Text4 = car(1).waypoint(4)
Text5 = car(2).waypoint(1)
Text6 = car(2).waypoint(2)
Text7 = car(2).waypoint(3)
Text8 = car(2).waypoint(4)

Next n

For n = 1 To numofcars
If car(n).waypoint(1) = laps + 1 And car(n).stopped = False Then
car(n).time = txtStopWatch.Text
car(n).stopped = True
car(n).speed = 0
car(n).bounce = 0
If firstdone = False Then
    winnernum = car(n).number
    car(n).score = car(n).score + 1
    winnertime = car(n).time
    firstdone = True
Else
    losertime = car(n).time
    losernum = car(n).number
End If
End If
Next n

Dim p, s As Integer

If car(1).stopped And car(2).stopped Then
    tmrStopWatch.Enabled = False
    fraWinner.Visible = True
       
    If winnernum = 1 Then
        imgWinner.Picture = LoadPicture(App.Path & "\images\yellows.gif")
        imgPlayerWin.Picture = LoadPicture(App.Path & "\images\yellows.gif")
    ElseIf winnernum = 2 Then
        imgWinner.Picture = LoadPicture(App.Path & "\images\silvers.gif")
        imgPlayerWin.Picture = LoadPicture(App.Path & "\images\silvers.gif")
    ElseIf winnernum = 3 Then
        imgWinner.Picture = LoadPicture(App.Path & "\images\blues.gif")
        imgPlayerWin.Picture = LoadPicture(App.Path & "\images\blues.gif")
    ElseIf winnernum = 4 Then
        imgWinner.Picture = LoadPicture(App.Path & "\images\reds.gif")
        imgPlayerWin.Picture = LoadPicture(App.Path & "\images\reds.gif")
    End If
    
    If losernum = 1 Then
        imgPlayerLose.Picture = LoadPicture(App.Path & "\images\yellows.gif")
    ElseIf losernum = 2 Then
        imgPlayerLose.Picture = LoadPicture(App.Path & "\images\silvers.gif")
    ElseIf losernum = 3 Then
        imgPlayerLose.Picture = LoadPicture(App.Path & "\images\blues.gif")
    ElseIf losernum = 4 Then
        imgPlayerLose.Picture = LoadPicture(App.Path & "\images\reds.gif")
    End If
    
lblWinTime.Caption = winnertime
lblLoseTime.Caption = losertime

End If

Call movecars

'Collisions between cars

For b = 1 To numofcars
   For a = 1 To numofcars
      If a <> b Then
        If (car(a).Dir >= 6 And car(a).Dir <= 15) Or (car(a).Dir >= 24 And car(a).Dir <= 33) Then
            If car(b).tlxpos > car(a).tlxpos - 35 And car(b).tlxpos < car(a).tlxpos + 35 And car(b).tlypos > car(a).tlypos - 20 And car(b).tlypos < car(a).tlypos + 20 Then
                car(b).tlxpos = car(b).txpos
                car(b).tlypos = car(b).typos
                car(b).bounce = -car(b).speed
                car(b).speed = 0
            Else
                car(b).txpos = car(b).tlxpos
                car(b).typos = car(b).tlypos
            End If
        Else
            If car(b).tlxpos > car(a).tlxpos - 20 And car(b).tlxpos < car(a).tlxpos + 20 And car(b).tlypos > car(a).tlypos - 35 And car(b).tlypos < car(a).tlypos + 35 Then
                car(b).tlxpos = car(b).txpos
                car(b).tlypos = car(b).typos
                car(b).bounce = -car(b).speed
                car(b).speed = 0
            Else
                car(b).txpos = car(b).tlxpos
                car(b).typos = car(b).tlypos
            End If
        End If
        End If
   Next a
Next b

'Adjust bounce speed
For a = 1 To numofcars
    If car(a).bounce > 0 Then
        car(a).bounce = car(a).bounce - speedup
    ElseIf car(a).bounce < 0 Then
     car(a).bounce = car(a).bounce + speedup
     ElseIf Abs(car(a).bounce) < 10 Then
     car(a).bounce = 0
    End If
Next a

picCar.Cls    'to redraw
Call drawcars
End Sub

Private Sub tmrStarter_Timer()
counter = counter - 1

If counter = 3 Then
shRed1.FillColor = RGB(255, 0, 0)
shRed2.FillColor = RGB(255, 0, 0)
ElseIf counter = 2 Then
shYel1.FillColor = RGB(255, 255, 0)
shYel2.FillColor = RGB(255, 255, 0)
ElseIf counter = 1 Then
shGr1.FillColor = RGB(0, 255, 0)
shGr2.FillColor = RGB(0, 255, 0)
ElseIf counter = 0 Then
fraStarter.Visible = False
tmrStarter.Enabled = False
tmrRefresh.Enabled = True
tmrStopWatch.Enabled = True
TotalTenthSeconds = -1
End If
clearall
End Sub

Private Sub tmrStopWatch_Timer()
HideCaret (txtStopWatch.hwnd)

' increase the total amount of Tenth Seconds.
' we set the timer interval to 100, so every tenth second
' this sub will be executed
TotalTenthSeconds = TotalTenthSeconds + 1
' if the TotalTenthSeconds is equal to 10, set it to 0.
TenthSeconds = TotalTenthSeconds Mod 10
' 10 tenth seconds are equal to 1 second
' int - will give us the integer part of the number:
' int(0.9) = 0
TotalSeconds = Int(TotalTenthSeconds / 10)
' if the Seconds is equal to 60, set it to 0
Seconds = TotalSeconds Mod 60
Minutes = Int(TotalSeconds / 60) Mod 60
' update the textbox
If Seconds < 10 Then
txtStopWatch.Text = Minutes & " : 0" & Seconds & "." & TenthSeconds
Else
txtStopWatch.Text = Minutes & " : " & Seconds & "." & TenthSeconds
End If
End Sub

'Handles Rotation
Private Sub tmrTurn_Timer()
Dim a As Integer
a = 1
For a = 1 To numofcars
    If car(a).stopped = False Then
    If car(a).lefton = 1 And car(a).righton = 0 Then
      car(a).Dir = car(a).Dir - 1
    ElseIf car(a).righton = 1 And car(a).lefton = 0 Then
      car(a).Dir = car(a).Dir + 1
    End If
    If car(a).Dir < 1 Then
        car(a).Dir = 36
    End If
    If car(a).Dir > 36 Then
        car(a).Dir = 1
    End If
    End If
Next a
End Sub
