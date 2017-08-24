VERSION 5.00
Begin VB.Form frmOptions 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Game Settings"
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5595
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   5595
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSoundoff 
      BackColor       =   &H00008000&
      Caption         =   "Sound Off"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   3960
      Width           =   1575
   End
   Begin VB.CommandButton cmdSoundOn 
      BackColor       =   &H00008000&
      Caption         =   "Sound On"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   3960
      Width           =   1575
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00008000&
      Caption         =   "Dec"
      Height          =   255
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   3240
      Width           =   855
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00008000&
      Caption         =   "Inc"
      Height          =   255
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   2880
      Width           =   855
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00008000&
      Caption         =   "Inc"
      Height          =   255
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2040
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00008000&
      Caption         =   "Dec"
      Height          =   255
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2400
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00008000&
      Caption         =   "Inc"
      Height          =   255
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1200
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00008000&
      Caption         =   "Dec"
      Height          =   255
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1560
      Width           =   855
   End
   Begin VB.CommandButton cmdSUdown 
      BackColor       =   &H00008000&
      Caption         =   "Dec"
      Height          =   255
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   720
      Width           =   855
   End
   Begin VB.CommandButton cmdSUup 
      BackColor       =   &H00008000&
      Caption         =   "Inc"
      Height          =   255
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   360
      Width           =   855
   End
   Begin VB.CommandButton cmcContinue 
      BackColor       =   &H00008000&
      Caption         =   "Continue"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2070
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5040
      Width           =   1455
   End
   Begin VB.Label lblLaps1 
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   330
      Left            =   3030
      TabIndex        =   16
      Top             =   3000
      Width           =   75
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      Caption         =   "Laps per Race"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   495
      TabIndex        =   13
      Top             =   3000
      Width           =   1995
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      Caption         =   "Maximum Speed"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   240
      TabIndex        =   8
      Top             =   1320
      Width           =   2250
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      Caption         =   "Minimum Speed"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   315
      TabIndex        =   7
      Top             =   2160
      Width           =   2175
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      Caption         =   "Speed Up"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   1080
      TabIndex        =   6
      Top             =   480
      Width           =   1410
   End
   Begin VB.Label lblSpeedMax1 
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   330
      Left            =   3030
      TabIndex        =   3
      Top             =   1320
      Width           =   75
   End
   Begin VB.Label lblSpeedMin1 
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   330
      Left            =   3030
      TabIndex        =   2
      Top             =   2160
      Width           =   75
   End
   Begin VB.Label lblSpeedUp1 
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   330
      Left            =   3030
      TabIndex        =   1
      Top             =   480
      Width           =   75
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmcContinue_Click()
Unload frmOptions
frmTracks.Enabled = True
frmTracks.Visible = True
End Sub

Private Sub cmdSoundoff_Click()
sound = False
StopMidi (App.Path & "\sounds\" & Midiname & ".mid")
End Sub

Private Sub cmdSoundOn_Click()
sound = True
randommidi = Int((5 * Rnd + 1))
Midiname = "Midi" + Format(randommidi)
PlayMIDI (App.Path & "\sounds\" & Midiname & ".mid")
End Sub

Private Sub cmdSUdown_Click()
If speedup > 1 Then
speedup = speedup - 1
End If
lblSpeedUp1 = speedup
End Sub

Private Sub cmdSUup_Click()
speedup = speedup + 1
lblSpeedUp1 = speedup
End Sub

Private Sub Command1_Click()
If speedmax > 10 Then
speedmax = speedmax - 5
End If
lblSpeedMax1 = speedmax
End Sub

Private Sub Command2_Click()
speedmax = speedmax + 5
lblSpeedMax1 = speedmax
End Sub

Private Sub Command3_Click()
speedmin = speedmin - 5
lblSpeedMin1 = speedmin
End Sub

Private Sub Command4_Click()
If speedmin < -10 Then
speedmin = speedmin + 5
End If
lblSpeedMin1 = speedmin
End Sub

Private Sub Command5_Click()
laps = laps + 1
lblLaps1 = laps
End Sub

Private Sub Command6_Click()
If laps > 1 Then
laps = laps - 1
End If
lblLaps1 = laps
End Sub

Private Sub Form_Load()
lblLaps1 = laps
lblSpeedUp1 = speedup
lblSpeedMax1 = speedmax
lblSpeedMin1 = speedmin
End Sub
