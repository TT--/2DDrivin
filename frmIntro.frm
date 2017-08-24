VERSION 5.00
Begin VB.Form frmIntro 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "2D Drivin' by Nick and Tyler"
   ClientHeight    =   8520
   ClientLeft      =   870
   ClientTop       =   720
   ClientWidth     =   11910
   ControlBox      =   0   'False
   Icon            =   "frmIntro.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8520
   ScaleWidth      =   11910
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   4500
      Left            =   4035
      Picture         =   "frmIntro.frx":030A
      Top             =   3480
      Width           =   3855
   End
   Begin VB.Image Image2 
      Height          =   2475
      Left            =   3428
      Picture         =   "frmIntro.frx":ADD6
      Top             =   120
      Width           =   5055
   End
End
Attribute VB_Name = "frmIntro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Click()
Load frmTracks
frmCarSelect.Enabled = True
frmCarSelect.Visible = True
Unload frmIntro
End Sub

Private Sub Form_Load()
Randomize
randommidi = Int((5 * Rnd + 1))
Midiname = "Midi" + Format(randommidi)
PlayMIDI (App.Path & "\sounds\" & Midiname & ".mid")
sound = True

laps = 2            'should be 2
speedup = 1    'home 1  school 10
speedmax = 71    'home 61 school 151
speedmin = -31    'home -31 school -51

car(1).score = 0
car(2).score = 0
End Sub

Private Sub Image1_Click()
Load frmTracks
frmCarSelect.Enabled = True
frmCarSelect.Visible = True
Unload frmIntro
End Sub

Private Sub Image2_Click()
Load frmTracks
frmCarSelect.Enabled = True
frmCarSelect.Visible = True
Unload frmIntro
End Sub

