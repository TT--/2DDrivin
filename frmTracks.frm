VERSION 5.00
Begin VB.Form frmTracks 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "2D Drivin' by Nick and Tyler"
   ClientHeight    =   8520
   ClientLeft      =   315
   ClientTop       =   435
   ClientWidth     =   11910
   Icon            =   "frmTracks.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8520
   ScaleWidth      =   11910
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picCredits 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   4695
      Left            =   3008
      ScaleHeight     =   4665
      ScaleWidth      =   5865
      TabIndex        =   10
      Top             =   1913
      Visible         =   0   'False
      Width           =   5895
      Begin VB.PictureBox Picture2 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   4455
         Left            =   120
         ScaleHeight     =   4455
         ScaleWidth      =   5655
         TabIndex        =   11
         Top             =   120
         Width           =   5655
         Begin VB.Label Label5 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Click here to continue"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFC0C0&
            Height          =   270
            Left            =   1620
            TabIndex        =   16
            Top             =   3840
            Width           =   2235
         End
         Begin VB.Label Label4 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Graphics:  Tyler"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   330
            Left            =   1275
            TabIndex        =   15
            Top             =   2280
            Width           =   1995
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Coding:  Nick and Tyler"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   330
            Left            =   1275
            TabIndex        =   14
            Top             =   1440
            Width           =   2925
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "2D Drivin' (c) 2002 by Nick and Tyler"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   330
            Left            =   240
            TabIndex        =   13
            Top             =   240
            Width           =   4995
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "ttrezise@hotmail.com"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   330
            Left            =   1372
            TabIndex        =   12
            Top             =   3240
            Width           =   2730
         End
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   600
      Top             =   3240
   End
   Begin VB.CommandButton cmdCredits 
      BackColor       =   &H00008000&
      Caption         =   "About"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "The Credits"
      Top             =   6720
      Width           =   2175
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00008000&
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Race outta here"
      Top             =   6720
      Width           =   2175
   End
   Begin VB.CommandButton cmdOptions 
      BackColor       =   &H00008000&
      Caption         =   "Game Options"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Adjust Speed and Laps"
      Top             =   6720
      Width           =   2175
   End
   Begin VB.CommandButton cmdTrkr 
      BackColor       =   &H00008000&
      Caption         =   "Random Track"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   4928
      Picture         =   "frmTracks.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Click here for a Random track"
      Top             =   4080
      Width           =   2055
   End
   Begin VB.CommandButton cmdScores 
      BackColor       =   &H00008000&
      Caption         =   "Current Score"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "See the Running Total"
      Top             =   6720
      Width           =   2175
   End
   Begin VB.CommandButton cmdTrk1 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Caption         =   "Track 1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   360
      Picture         =   "frmTracks.frx":0779
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Click to choose this track"
      Top             =   360
      Width           =   2655
   End
   Begin VB.CommandButton cmdTrk2 
      BackColor       =   &H00008000&
      Caption         =   "Track 2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   3240
      Picture         =   "frmTracks.frx":D103
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Click to choose this track"
      Top             =   360
      Width           =   2655
   End
   Begin VB.CommandButton cmdTrk3 
      BackColor       =   &H00008000&
      Caption         =   "Track 3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   6120
      Picture         =   "frmTracks.frx":19C55
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Click to choose this track"
      Top             =   360
      Width           =   2655
   End
   Begin VB.CommandButton cmdTrk4 
      BackColor       =   &H00008000&
      Caption         =   "Track 4"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   9000
      Picture         =   "frmTracks.frx":265DF
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Click to choose this track"
      Top             =   360
      Width           =   2655
   End
   Begin VB.Label lblpicktrk 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Click to Select a Track"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   615
      Left            =   2948
      TabIndex        =   4
      Top             =   2880
      Width           =   6015
   End
End
Attribute VB_Name = "frmTracks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCredits_Click()
picCredits.Visible = True
Timer1.Enabled = True
picCredits.ScaleMode = 3
Picture2.BorderStyle = 0
Picture2.Top = picCredits.ScaleHeight
Picture2.Left = (picCredits.Width \ Screen.TwipsPerPixelX - Picture2.Width) / 2
End Sub

Private Sub cmdExit_Click()
StopMidi (App.Path & "\sounds\" & Midiname & ".mid")
End
End Sub

Private Sub cmdOptions_Click()
Load frmOptions
frmOptions.Enabled = True
frmOptions.Visible = True
End Sub

Private Sub cmdScores_Click()
Load frmScores
frmScores.Enabled = True
frmScores.Visible = True
End Sub

Private Sub cmdTrk1_Click()
Trackname = "Track1"
Load frmMain
frmMain.Enabled = True
frmMain.Visible = True
'Unload frmTracks
End Sub
Private Sub cmdTrk2_Click()
Trackname = "Track2"
Load frmMain
frmMain.Enabled = True
frmMain.Visible = True
'Unload frmTracks
End Sub
Private Sub cmdTrk3_Click()
Trackname = "Track3"
Load frmMain
frmMain.Enabled = True
frmMain.Visible = True
'Unload frmTracks
End Sub
Private Sub cmdTrk4_Click()
Trackname = "Track4"
Load frmMain
frmMain.Enabled = True
frmMain.Visible = True
'Unload frmTracks
End Sub

Private Sub cmdTrkr_Click()
Randomize
randomtrack = Int((4 * Rnd + 1))
Trackname = "Track" + Format(randomtrack)
Load frmMain
frmMain.Enabled = True
frmMain.Visible = True
'Unload frmTracks
End Sub

Private Sub picCredits_Click()
Timer1.Enabled = False
picCredits.Visible = False
End Sub

Private Sub picture2_Click()
Timer1.Enabled = False
picCredits.Visible = False
End Sub

Private Sub Timer1_Timer()
Picture2.Top = Picture2.Top - 1
If Picture2.Top <= -Picture2.Height Then Picture2.Top = picCredits.ScaleHeight
End Sub

