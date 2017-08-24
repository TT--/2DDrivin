VERSION 5.00
Begin VB.Form frmScores 
   BackColor       =   &H80000007&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Current Score"
   ClientHeight    =   3210
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4110
   ControlBox      =   0   'False
   Enabled         =   0   'False
   Icon            =   "frmScores.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   4110
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
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
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label lblLeftScore 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   420
      TabIndex        =   2
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label lblRightScore 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2520
      TabIndex        =   1
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Image imgPlayerL 
      Height          =   735
      Left            =   720
      Top             =   240
      Width           =   495
   End
   Begin VB.Image imgPlayerR 
      Height          =   735
      Left            =   2880
      Top             =   240
      Width           =   495
   End
End
Attribute VB_Name = "frmScores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmcContinue_Click()
Unload frmScores
frmTracks.Visible = True
frmTracks.Enabled = True
End Sub

Private Sub Form_Load()

lblLeftScore.Caption = car(1).score
lblRightScore.Caption = car(2).score


Select Case car(1).number
 Case 1
   imgPlayerL.Picture = LoadPicture(App.Path & "\images\yellows.gif")
 Case 2
   imgPlayerL.Picture = LoadPicture(App.Path & "\images\silvers.gif")
 Case 3
   imgPlayerL.Picture = LoadPicture(App.Path & "\images\blues.gif")
 Case 4
    imgPlayerL.Picture = LoadPicture(App.Path & "\images\reds.gif")
 End Select

Select Case car(2).number
 Case 1
   imgPlayerR.Picture = LoadPicture(App.Path & "\images\yellows.gif")
 Case 2
   imgPlayerR.Picture = LoadPicture(App.Path & "\images\silvers.gif")
 Case 3
   imgPlayerR.Picture = LoadPicture(App.Path & "\images\blues.gif")
 Case 4
    imgPlayerR.Picture = LoadPicture(App.Path & "\images\reds.gif")
 End Select

End Sub
