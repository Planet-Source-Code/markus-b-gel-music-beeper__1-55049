VERSION 5.00
Begin VB.Form frmMusicBeeper 
   BorderStyle     =   0  'Kein
   ClientHeight    =   6135
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5625
   Icon            =   "frmMusicBeeper.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   5625
   Begin VB.Frame fraRecording 
      Caption         =   "Aufnahme und Wiedergabe"
      Height          =   1455
      Left            =   120
      TabIndex        =   13
      Top             =   1800
      Width           =   5385
      Begin VB.Label lblRecording 
         Caption         =   "Wiedergabe starten / beenden"
         Height          =   255
         Index           =   2
         Left            =   480
         TabIndex        =   19
         Top             =   1080
         Width           =   2280
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblRecordingValue 
         Caption         =   "F10"
         Height          =   255
         Index           =   2
         Left            =   3840
         TabIndex        =   18
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label lblRecordingValue 
         Caption         =   "F9"
         Height          =   255
         Index           =   1
         Left            =   3840
         TabIndex        =   17
         Top             =   720
         Width           =   615
      End
      Begin VB.Label lblRecording 
         Caption         =   "Aufnahme starten / beenden"
         Height          =   255
         Index           =   1
         Left            =   480
         TabIndex        =   16
         Top             =   720
         Width           =   2160
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblRecordingValue 
         Caption         =   "Pause"
         Height          =   255
         Index           =   0
         Left            =   3840
         TabIndex        =   15
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lblRecording 
         Caption         =   "unterbrechen / fortsetzen"
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   14
         Top             =   360
         Width           =   2040
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame fraForms 
      Caption         =   "Anzeige"
      Height          =   1815
      Left            =   120
      TabIndex        =   4
      Top             =   3360
      Width           =   5385
      Begin VB.Label lblFormsValue 
         Caption         =   "F8"
         Height          =   255
         Index           =   3
         Left            =   3840
         TabIndex        =   12
         Top             =   1440
         Width           =   870
      End
      Begin VB.Label lblForms 
         Caption         =   "Anweisungen ein- / ausblenden"
         Height          =   255
         Index           =   3
         Left            =   480
         TabIndex        =   11
         Top             =   1440
         Width           =   2400
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblFormsValue 
         Caption         =   "F7"
         Height          =   255
         Index           =   2
         Left            =   3840
         TabIndex        =   10
         Top             =   1080
         Width           =   870
      End
      Begin VB.Label lblForms 
         Caption         =   "Hörtest ein- / ausblenden"
         Height          =   255
         Index           =   2
         Left            =   480
         TabIndex        =   9
         Top             =   1080
         Width           =   2160
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblFormsValue 
         Caption         =   "F6"
         Height          =   255
         Index           =   1
         Left            =   3840
         TabIndex        =   8
         Top             =   720
         Width           =   870
      End
      Begin VB.Label lblForms 
         Caption         =   "Tastatur ein- / ausblenden"
         Height          =   255
         Index           =   1
         Left            =   480
         TabIndex        =   7
         Top             =   720
         Width           =   2160
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblFormsValue 
         Caption         =   "F5"
         Height          =   255
         Index           =   0
         Left            =   3840
         TabIndex        =   6
         Top             =   360
         Width           =   870
      End
      Begin VB.Label lblForms 
         Caption         =   "Recorder ein- / ausblenden"
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   5
         Top             =   360
         Width           =   2160
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame fraSound 
      Caption         =   "Klang"
      Height          =   1095
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   5385
      Begin VB.OptionButton optSound 
         Caption         =   "variable Dauer"
         Height          =   195
         Index           =   1
         Left            =   480
         TabIndex        =   2
         Top             =   720
         Width           =   2055
      End
      Begin VB.OptionButton optSound 
         Caption         =   "gleiche Dauer"
         Height          =   195
         Index           =   0
         Left            =   480
         TabIndex        =   1
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label lblSound 
         Caption         =   "F2"
         Height          =   255
         Index           =   1
         Left            =   3840
         TabIndex        =   21
         Top             =   720
         Width           =   855
      End
      Begin VB.Label lblSound 
         Caption         =   "F1"
         Height          =   255
         Index           =   0
         Left            =   3840
         TabIndex        =   20
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Timer tmrScan 
      Interval        =   10
      Left            =   5040
      Top             =   240
   End
   Begin VB.Label lblForms 
      Caption         =   "Music-Beeper beenden"
      Height          =   255
      Index           =   5
      Left            =   600
      TabIndex        =   25
      Top             =   5760
      Width           =   2160
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblFormsValue 
      Caption         =   "ESC"
      Height          =   255
      Index           =   5
      Left            =   3960
      TabIndex        =   24
      Top             =   5760
      Width           =   870
   End
   Begin VB.Label lblForms 
      Caption         =   "Einstellungen ändern"
      Height          =   255
      Index           =   4
      Left            =   600
      TabIndex        =   23
      Top             =   5400
      Width           =   2160
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblFormsValue 
      Caption         =   "F12"
      Height          =   255
      Index           =   4
      Left            =   3960
      TabIndex        =   22
      Top             =   5400
      Width           =   870
   End
   Begin VB.Line lin 
      BorderColor     =   &H80000015&
      Index           =   15
      X1              =   0
      X2              =   0
      Y1              =   -120
      Y2              =   600
   End
   Begin VB.Line lin 
      BorderColor     =   &H80000010&
      Index           =   14
      X1              =   0
      X2              =   0
      Y1              =   240
      Y2              =   720
   End
   Begin VB.Line lin 
      BorderColor     =   &H80000016&
      Index           =   13
      X1              =   0
      X2              =   0
      Y1              =   120
      Y2              =   840
   End
   Begin VB.Line lin 
      BorderColor     =   &H80000014&
      Index           =   12
      X1              =   0
      X2              =   0
      Y1              =   240
      Y2              =   720
   End
   Begin VB.Line lin 
      BorderColor     =   &H80000015&
      Index           =   11
      X1              =   0
      X2              =   0
      Y1              =   120
      Y2              =   720
   End
   Begin VB.Line lin 
      BorderColor     =   &H80000010&
      Index           =   10
      X1              =   0
      X2              =   0
      Y1              =   120
      Y2              =   720
   End
   Begin VB.Line lin 
      BorderColor     =   &H80000016&
      Index           =   9
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   720
   End
   Begin VB.Line lin 
      BorderColor     =   &H80000014&
      Index           =   8
      X1              =   0
      X2              =   0
      Y1              =   120
      Y2              =   720
   End
   Begin VB.Line lin 
      BorderColor     =   &H80000015&
      Index           =   3
      X1              =   15
      X2              =   0
      Y1              =   15
      Y2              =   720
   End
   Begin VB.Line lin 
      BorderColor     =   &H80000016&
      Index           =   1
      X1              =   0
      X2              =   0
      Y1              =   120
      Y2              =   720
   End
   Begin VB.Line lin 
      BorderColor     =   &H80000010&
      Index           =   6
      X1              =   0
      X2              =   0
      Y1              =   720
      Y2              =   0
   End
   Begin VB.Line lin 
      BorderColor     =   &H80000015&
      Index           =   7
      X1              =   0
      X2              =   0
      Y1              =   720
      Y2              =   120
   End
   Begin VB.Line lin 
      BorderColor     =   &H80000014&
      Index           =   4
      X1              =   0
      X2              =   0
      Y1              =   240
      Y2              =   840
   End
   Begin VB.Line lin 
      BorderColor     =   &H80000016&
      Index           =   5
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   840
   End
   Begin VB.Line lin 
      BorderColor     =   &H80000010&
      Index           =   2
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   720
   End
   Begin VB.Line lin 
      BorderColor     =   &H80000014&
      Index           =   0
      X1              =   0
      X2              =   0
      Y1              =   120
      Y2              =   840
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Zentriert
      Caption         =   "Music-Beeper"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5175
   End
End
Attribute VB_Name = "frmMusicBeeper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_bDragging As Boolean 'flag: window is moved
Private m_lDownX As Long 'point on which the mouse was
Private m_lDownY As Long 'put down over the form

' ##################################################################################
'
' form events
'
' ##################################################################################

Private Sub Form_Load()
  
  Width = 5625
  Height = 6135
  Top = (Screen.Height - Height) / 2
  Left = (Screen.Width - Width) / 2
  SetLanguage
  SetBorder Me, EB_SIMPLE

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  
  If KeyCode = 112 Then 'F1
    optSound(0).Enabled = True
  ElseIf KeyCode = 113 Then 'F2
    optSound(1).Enabled = True
  End If
  g_clsMusicBeeper.GetKey KeyCode, Shift
  
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, _
                           x As Single, Y As Single)
  
  If (x < 106 Or x > Width - 106 _
  Or Y < 106 Or Y > Height - 106) Then  'border
    m_lDownX = x
    m_lDownY = Y
    m_bDragging = True
  End If

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, _
                           x As Single, Y As Single)
  If m_bDragging Then
    'pull window
    Move Left + (x - m_lDownX), Top + (Y - m_lDownY)
  
    If g_eRecDocked > ED_NO Then
      'pull record window
      g_frmRecord.Move g_frmRecord.Left + (x - m_lDownX), g_frmRecord.Top + (Y - m_lDownY)
    End If
    
    If g_eKeyDocked > ED_NO Then
      'pull key window
      g_frmKeys.Move g_frmKeys.Left + (x - m_lDownX), g_frmKeys.Top + (Y - m_lDownY)
    End If
    
    If g_eTestDocked > ED_NO Then
      'pull test window
      g_frmTest.Move g_frmTest.Left + (x - m_lDownX), g_frmTest.Top + (Y - m_lDownY)
    End If
    
  End If
  
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, _
                         x As Single, Y As Single)
  m_bDragging = False

End Sub

Private Sub optSound_Click(Index As Integer)

  Select Case Index
  Case 0
    g_clsMusicBeeper.GetKey 112, 0
  Case 1
    g_clsMusicBeeper.GetKey 113, 0
  End Select

End Sub

' ##################################################################################
'
' control events
'
' ##################################################################################

Private Sub tmrScan_Timer()

  If g_iMode = 1 Or g_iRunMode = RUNMODE_RECORD Then
    ScanCurrentKeyState
  End If

End Sub

' ##################################################################################
'
' public procedures
'
' ##################################################################################

Public Sub SetLanguage()
  
  Select Case g_eLanguage
  Case EL_GERMAN
    Caption = " Music-Beeper"
    lblTitle.Caption = "Music-Beeper"
    fraSound.Caption = "Klang"
    fraRecording.Caption = "Aufnahme und Wiedergabe"
    fraForms.Caption = "Anzeige"
    optSound(0).Caption = "feste Dauer"
    optSound(1).Caption = "variable Dauer"
    lblForms(0).Caption = "Recorder ein- / ausblenden"
    lblForms(1).Caption = "Tastatur ein- / ausblenden"
    lblForms(2).Caption = "Hörtest ein- / ausblenden"
    lblForms(3).Caption = "Anweisungen ein- / ausblenden"
    lblForms(4).Caption = "Einstellungen ändern"
    lblForms(5).Caption = "Music-Beeper beenden"
    lblRecording(0).Caption = "unterbrechen / fortsetzen"
    lblRecordingValue(0).Caption = "Pause"
    lblRecording(1).Caption = "Aufnahme starten / beenden"
    lblRecording(2).Caption = "Wiedergabe starten / beenden"
  
  Case EL_ENGLISH
    Caption = " Music Beeper"
    lblTitle.Caption = "Music Beeper"
    fraSound.Caption = "Sound"
    fraRecording.Caption = "Record and play"
    fraForms.Caption = "View"
    optSound(0).Caption = "fixed duration"
    optSound(1).Caption = "variant duration"
    lblForms(0).Caption = "show / hide recorder"
    lblForms(1).Caption = "show / hide keyboard"
    lblForms(2).Caption = "show / hide audio test"
    lblForms(3).Caption = "hide intructions"
    lblForms(4).Caption = "change settings"
    lblForms(5).Caption = "close Music Beeper"
    lblRecording(0).Caption = "stop / continue"
    lblRecordingValue(0).Caption = "BREAK"
    lblRecording(1).Caption = "start / finish recording"
    lblRecording(2).Caption = "start / finish play"
    
  End Select
 
End Sub

