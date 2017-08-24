VERSION 5.00
Begin VB.Form frmCarSelect 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select a Car"
   ClientHeight    =   8520
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11910
   ControlBox      =   0   'False
   Icon            =   "frmCarSelect.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8520
   ScaleWidth      =   11910
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   480
      Top             =   3480
   End
   Begin VB.Image Image4 
      Height          =   630
      Left            =   11400
      Picture         =   "frmCarSelect.frx":030A
      Top             =   5040
      Width           =   405
   End
   Begin VB.Image Image3 
      Height          =   600
      Left            =   120
      Picture         =   "frmCarSelect.frx":098B
      Top             =   4920
      Width           =   375
   End
   Begin VB.Image Image2 
      Height          =   600
      Left            =   11520
      Picture         =   "frmCarSelect.frx":0FDA
      Top             =   2520
      Width           =   360
   End
   Begin VB.Image Image1 
      Height          =   630
      Left            =   120
      Picture         =   "frmCarSelect.frx":1659
      Top             =   2520
      Width           =   375
   End
   Begin VB.Image imgRight 
      DragMode        =   1  'Automatic
      Height          =   480
      Left            =   6120
      Picture         =   "frmCarSelect.frx":1CCB
      Tag             =   "Right"
      Top             =   4080
      Width           =   480
   End
   Begin VB.Image imgLeft 
      DragMode        =   1  'Automatic
      Height          =   480
      Left            =   5280
      Picture         =   "frmCarSelect.frx":1FD5
      Tag             =   "Left"
      Top             =   4080
      Width           =   480
   End
   Begin VB.Image imgCarYellow 
      Height          =   2400
      Left            =   6600
      OLEDragMode     =   1  'Automatic
      Picture         =   "frmCarSelect.frx":22DF
      Top             =   120
      Width           =   5250
   End
   Begin VB.Image imgCarSilver 
      Height          =   2400
      Left            =   120
      Picture         =   "frmCarSelect.frx":9FAA
      Top             =   120
      Width           =   5250
   End
   Begin VB.Image imgCarBlue 
      Height          =   2925
      Left            =   120
      Picture         =   "frmCarSelect.frx":13F3D
      Top             =   5520
      Width           =   5250
   End
   Begin VB.Image imgEmpty 
      Height          =   1815
      Left            =   3315
      Top             =   3120
      Width           =   5295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Drag to Select a Car"
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
      Left            =   3240
      TabIndex        =   0
      Top             =   3360
      Width           =   5415
   End
   Begin VB.Image imgCarRed 
      Height          =   2745
      Left            =   6600
      Picture         =   "frmCarSelect.frx":21501
      Top             =   5640
      Width           =   5190
   End
End
Attribute VB_Name = "frmCarSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub imgCarBlue_DragDrop(Source As Control, X As Single, Y As Single)
    If Source.Tag = "Left" Then
    imgLeft.Left = 2400
    imgLeft.Top = 6720
    car(1).number = 3
    ElseIf Source.Tag = "Right" Then
    imgRight.Left = 2400
    imgRight.Top = 6720
    car(2).number = 3
    End If
End Sub
Private Sub imgCarRed_DragDrop(Source As Control, X As Single, Y As Single)
    If Source.Tag = "Left" Then
    imgLeft.Left = 9120
    imgLeft.Top = 6840
    car(1).number = 4
    ElseIf Source.Tag = "Right" Then
    imgRight.Left = 9120
    imgRight.Top = 6840
    car(2).number = 4
    End If
End Sub
Private Sub imgCarSilver_DragDrop(Source As Control, X As Single, Y As Single)
    If Source.Tag = "Left" Then
    imgLeft.Left = 2280
    imgLeft.Top = 1080
    car(1).number = 2
    ElseIf Source.Tag = "Right" Then
    imgRight.Left = 2280
    imgRight.Top = 1080
    car(2).number = 2
    End If
End Sub
Private Sub imgCarYellow_DragDrop(Source As Control, X As Single, Y As Single)
    If Source.Tag = "Left" Then
    imgLeft.Left = 9120
    imgLeft.Top = 1080
    car(1).number = 1
    ElseIf Source.Tag = "Right" Then
    imgRight.Left = 9120
    imgRight.Top = 1080
    car(2).number = 1
    End If
End Sub
Private Sub imgEmpty_DragDrop(Source As Control, X As Single, Y As Single)
    If Source.Tag = "Left" Then
    imgLeft.Left = 5280
    imgLeft.Top = 4080
    car(1).number = 0
    ElseIf Source.Tag = "Right" Then
    imgRight.Left = 6120
    imgRight.Top = 4080
    car(2).number = 0
    End If
End Sub

Private Sub Timer1_Timer()
If car(1).number <> 0 And car(2).number <> 0 Then
If car(1).number <> car(2).number Then
Load frmTracks
frmTracks.Visible = True
frmTracks.Enabled = True
frmCarSelect.Visible = False
frmCarSelect.Enabled = False
End If
End If
End Sub
