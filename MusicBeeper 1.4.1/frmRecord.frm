VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmRecord 
   BorderStyle     =   0  'Kein
   Caption         =   "1095"
   ClientHeight    =   1935
   ClientLeft      =   105
   ClientTop       =   -180
   ClientWidth     =   2535
   Icon            =   "frmRecord.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   2535
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraRecord 
      BorderStyle     =   0  'Kein
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   960
      Width           =   2175
      Begin VB.OptionButton optRecorder 
         Height          =   195
         Index           =   0
         Left            =   1200
         TabIndex        =   11
         Top             =   120
         Width           =   1095
      End
      Begin VB.OptionButton optRecorder 
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   10
         Top             =   120
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdRecord 
      Height          =   345
      Index           =   6
      Left            =   2100
      Picture         =   "frmRecord.frx":044A
      Style           =   1  'Grafisch
      TabIndex        =   7
      Top             =   0
      Width           =   345
   End
   Begin VB.CommandButton cmdRecord 
      Height          =   345
      Index           =   5
      Left            =   360
      Picture         =   "frmRecord.frx":0594
      Style           =   1  'Grafisch
      TabIndex        =   6
      Top             =   0
      Width           =   345
   End
   Begin VB.CommandButton cmdRecord 
      Height          =   345
      Index           =   4
      Left            =   15
      Picture         =   "frmRecord.frx":06DE
      Style           =   1  'Grafisch
      TabIndex        =   5
      Top             =   0
      Width           =   345
   End
   Begin MSComDlg.CommonDialog cdlRecord 
      Left            =   1800
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdRecord 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   3
      Left            =   1800
      Picture         =   "frmRecord.frx":0828
      Style           =   1  'Grafisch
      TabIndex        =   3
      Top             =   360
      Width           =   600
   End
   Begin VB.CommandButton cmdRecord 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   2
      Left            =   1200
      Picture         =   "frmRecord.frx":097A
      Style           =   1  'Grafisch
      TabIndex        =   2
      Top             =   360
      Width           =   600
   End
   Begin VB.CommandButton cmdRecord 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   1
      Left            =   600
      Picture         =   "frmRecord.frx":0ACC
      Style           =   1  'Grafisch
      TabIndex        =   1
      Top             =   360
      Width           =   600
   End
   Begin VB.CommandButton cmdRecord 
      Cancel          =   -1  'True
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   0
      Left            =   0
      Picture         =   "frmRecord.frx":0C1E
      Style           =   1  'Grafisch
      TabIndex        =   0
      Top             =   360
      Width           =   600
   End
   Begin VB.Line lin 
      BorderColor     =   &H80000015&
      Index           =   23
      X1              =   0
      X2              =   0
      Y1              =   720
      Y2              =   0
   End
   Begin VB.Line lin 
      BorderColor     =   &H80000010&
      Index           =   22
      X1              =   0
      X2              =   0
      Y1              =   720
      Y2              =   0
   End
   Begin VB.Line lin 
      BorderColor     =   &H80000016&
      Index           =   21
      X1              =   0
      X2              =   0
      Y1              =   720
      Y2              =   0
   End
   Begin VB.Line lin 
      BorderColor     =   &H80000014&
      Index           =   20
      X1              =   0
      X2              =   0
      Y1              =   720
      Y2              =   0
   End
   Begin VB.Line lin 
      BorderColor     =   &H80000015&
      Index           =   19
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   720
   End
   Begin VB.Line lin 
      BorderColor     =   &H80000010&
      Index           =   18
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line lin 
      BorderColor     =   &H80000016&
      Index           =   17
      X1              =   0
      X2              =   0
      Y1              =   600
      Y2              =   0
   End
   Begin VB.Line lin 
      BorderColor     =   &H80000014&
      Index           =   16
      X1              =   0
      X2              =   0
      Y1              =   720
      Y2              =   0
   End
   Begin VB.Line lin 
      BorderColor     =   &H80000014&
      Index           =   0
      X1              =   0
      X2              =   0
      Y1              =   120
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
      BorderColor     =   &H80000016&
      Index           =   5
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   840
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
      BorderColor     =   &H80000015&
      Index           =   7
      X1              =   0
      X2              =   0
      Y1              =   720
      Y2              =   120
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
      BorderColor     =   &H80000016&
      Index           =   1
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
      Y1              =   0
      Y2              =   705
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
      BorderColor     =   &H80000016&
      Index           =   9
      X1              =   0
      X2              =   0
      Y1              =   0
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
      BorderColor     =   &H80000015&
      Index           =   11
      X1              =   0
      X2              =   0
      Y1              =   120
      Y2              =   720
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
      BorderColor     =   &H80000016&
      Index           =   13
      X1              =   0
      X2              =   0
      Y1              =   120
      Y2              =   840
   End
   Begin VB.Line lin 
      BorderColor     =   &H80000010&
      Index           =   14
      X1              =   0
      X2              =   0
      Y1              =   120
      Y2              =   600
   End
   Begin VB.Line lin 
      BorderColor     =   &H80000015&
      Index           =   15
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   720
   End
   Begin VB.Label lblRecord 
      Alignment       =   2  'Zentriert
      Caption         =   "Recorder"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   720
      TabIndex        =   8
      Top             =   0
      Width           =   1335
   End
   Begin VB.Label lblRecord 
      Height          =   480
      Index           =   1
      Left            =   0
      TabIndex        =   4
      Top             =   1440
      Width           =   2475
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_iKeyCodesRec() As Integer
Private m_lDurationsRec() As Long
Private m_iCounter As Integer
Private m_iPosition As Integer 'position in a recording
Private m_iStart As Integer 'where it is to be continued after a break
Private m_bStored As Boolean 'actually played sound has been stored in a file

'properties for moving the form
Private m_bDragging As Boolean 'flag: window is moved
Private m_lDownX As Long  'point on which the mouse was
Private m_lDownY As Long  '  put down on the form

Public m_iTop As Integer  'top border of g_frmMusicBeeper
Public m_iBottom As Integer  'bottom border of g_frmMusicBeeper
Public m_iLeft As Integer  'left border of g_frmMusicBeeper
Public m_iRight As Integer  'right border of g_frmMusicBeeper

Private m_sCurrent As String 'currently loaded file

'for labeling:
Private m_sPlay As String
Private m_sPlayLast As String
Private m_sStarted As String
Private m_sStopped As String
Private m_sStop As String
Private m_sContinue As String

' ##################################################################################
'
' form events
'
' ##################################################################################

Private Sub Form_Load()
  
  cmdRecord(1).Enabled = False
  cmdRecord(2).Enabled = False
  cmdRecord(3).Enabled = False
  SetLayout
  SetLanguage
  
  Top = (g_frmMusicBeeper.Top + g_frmMusicBeeper.Height / 2) - (Me.Height / 2)
  Left = g_frmMusicBeeper.Left + g_frmMusicBeeper.Width
  g_eRecDocked = ED_LEFT
  
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  
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
  End If
  
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, _
                         x As Single, Y As Single)
  m_bDragging = False
  FormMove Me, g_eRecDocked
  
  Select Case g_eRecDocked
  Case ED_TOP
    Top = m_iTop - Height
  Case ED_LEFT
    Left = m_iLeft - Width
  Case ED_BOTTOM
    Top = m_iBottom
  Case ED_RIGHT
    Left = m_iRight
  Case ED_NO
    'no action
  End Select

End Sub

' ##################################################################################
'
' control events
'
' ##################################################################################

Public Sub cmdRecord_Click(Index As Integer)

  Select Case Index
  
  Case 0 'start recording
    g_iRun = RUN_STARTED
    g_iRunMode = RUNMODE_RECORD
    ReDim m_iKeyCodesRec(0)
    ReDim m_lDurationsRec(0)
    m_iCounter = 0
    g_frmMusicBeeper.SetFocus
    
  Case 1 'break/continue recording/playing
    g_iRun = IIf(g_iRun = RUN_STARTED, RUN_STOPPED, RUN_STARTED)
    cmdRecord(1).ToolTipText = IIf(cmdRecord(1).ToolTipText = "Anhalten", "Fortsetzen", "Anhalten")
    m_iStart = m_iPosition
    If g_iRunMode = RUNMODE_RECORD Then
      g_frmMusicBeeper.SetFocus
    ElseIf g_iRunMode = RUNMODE_PLAY Then
      If g_iRun = RUN_STARTED Then
        Play m_iStart
      Else
        cmdRecord(1).SetFocus
      End If
    End If
    
  Case 2 'finish recording/playing
    g_iRun = RUN_FINISHED
    If g_iRunMode = RUNMODE_RECORD Then
      If MsgBox("Aufzeichnung speichern?", vbQuestion + vbYesNo, "Speichern?") = vbYes Then
        SaveFile
        m_bStored = True
      Else
        m_bStored = False
      End If
    ElseIf g_iRunMode = RUNMODE_PLAY Then
      m_iStart = 0
    End If
    g_iRunMode = RUNMODE_NOTHING
  
  Case 3 'play recorded sound or file
    g_iRunMode = RUNMODE_PLAY
    g_iRun = RUN_STARTED
    m_iStart = 0 'value is set back
    EnDisableKeys
    Play 0
    EnDisableKeys
  
  Case 4 'open recorded file
    OpenFile
    
  Case 5 'save recorded sound to file
    SaveFile
    
  Case 6 'close window
    Unload Me
    
  End Select
  
  EnDisableKeys
  
End Sub

' ##################################################################################
'
' public procedures
'
' ##################################################################################

Public Function PatchLabel()
Dim sString As String
  
  Select Case g_iRun
  Case RUN_NOTHING
  Case RUN_STARTED
    sString = sString & m_sStarted & vbCrLf
  Case RUN_STOPPED
    sString = sString & m_sStopped & vbCrLf
  Case RUN_FINISHED
  End Select
  
  Select Case g_iRunMode
  Case RUNMODE_NOTHING
  Case RUNMODE_PLAY
    If m_bStored Then
      sString = sString & m_sPlay & cdlRecord.FileTitle
    Else
      sString = sString & m_sPlayLast
    End If
  Case RUNMODE_RECORD
  End Select
  
  lblRecord(1).Caption = sString

End Function

Public Sub RecordIntervals(iKeycode As Integer)

  ReDim Preserve m_iKeyCodesRec(UBound(m_iKeyCodesRec) + 1) 'um ein Element erweitern
  ReDim Preserve m_lDurationsRec(UBound(m_lDurationsRec) + 1) 'um ein Element erweitern
  
  'it is checked if the tone is similar to the tone before
  If iKeycode = m_iKeyCodesRec(m_iCounter) _
  And g_iRecordMode = RECMODE_CLEAR Then  'if the clear mode is set
    'the duration of the tone before is extended
    m_lDurationsRec(m_iCounter) = m_lDurationsRec(m_iCounter) + IIf(iKeycode = 32, g_lInterval, 2 * g_lInterval)
  Else 'the tone and its duration are added to the arrays
    m_iCounter = m_iCounter + 1
    m_iKeyCodesRec(m_iCounter) = iKeycode
    m_lDurationsRec(m_iCounter) = g_lInterval
  End If

End Sub

Public Sub RecordKeys(iKeycode As Integer)

  ReDim Preserve m_iKeyCodesRec(UBound(m_iKeyCodesRec) + 1) 'add one element
  ReDim Preserve m_lDurationsRec(UBound(m_lDurationsRec) + 1) 'add one element
    
  m_iCounter = m_iCounter + 1
  m_iKeyCodesRec(m_iCounter) = iKeycode
  m_lDurationsRec(m_iCounter) = g_lDuration

End Sub

Public Sub SetLanguage()

  Select Case g_eLanguage
  
  Case EL_GERMAN
    m_sStarted = "Gestartet -"
    m_sStopped = "Unterbrochen -"
    m_sPlay = "Abspielen von "
    m_sPlayLast = "Abspielen der letzen Eingabe"
    m_sStop = "Anhalten"
    m_sContinue = "Fortsetzen"
      
    cmdRecord(0).ToolTipText = "Aufnahme starten"
    cmdRecord(2).ToolTipText = "Beenden"
    cmdRecord(3).ToolTipText = "Abspielen"
    cmdRecord(4).ToolTipText = "Öffnen"
    cmdRecord(5).ToolTipText = "Speichern"
    cmdRecord(6).ToolTipText = "Beenden"
  
    optRecorder(0).Caption = "klar"
    optRecorder(1).Caption = "rau"
  
  Case EL_ENGLISH
    m_sStarted = "Started -"
    m_sStopped = "Stopped -"
    m_sPlay = "Play "
    m_sPlayLast = "Play the last input"
    m_sStop = "Stop"
    m_sContinue = "Continue"
    
    cmdRecord(0).ToolTipText = "Start recording"
    cmdRecord(2).ToolTipText = "Close"
    cmdRecord(3).ToolTipText = "Play"
    cmdRecord(4).ToolTipText = "Open"
    cmdRecord(5).ToolTipText = "Save"
    cmdRecord(6).ToolTipText = "Close"
  
    optRecorder(0).Caption = "clear"
    optRecorder(1).Caption = "rough"
  
  End Select
  
  cmdRecord(1).ToolTipText = m_sStop
  
  PatchLabel
 
End Sub

' ##################################################################################
'
' private procedures
'
' ##################################################################################

Private Sub EnDisableKeys()

  cmdRecord(0).Enabled = (g_iRun = RUN_FINISHED Or RUN_NOTHING)
  cmdRecord(1).Enabled = (g_iRun = RUN_STARTED Or g_iRun = RUN_STOPPED)
  cmdRecord(2).Enabled = (g_iRun = RUN_STARTED)
  cmdRecord(3).Enabled = (g_iRun = RUN_FINISHED)
  cmdRecord(4).Enabled = (g_iRun = RUN_FINISHED Or RUN_NOTHING)
  cmdRecord(5).Enabled = (g_iRun = RUN_FINISHED And g_iRunMode = RUNMODE_RECORD)
  cmdRecord(6).Enabled = (g_iRun = RUN_FINISHED)
  
  PatchLabel
  
End Sub

Private Sub OpenFile()
Dim sLine As String
Dim i As Integer

  'read the file and fill the arrays m_iKeyCodesRec and m_lDurationsRec
  cdlRecord.InitDir = g_sPathSave
  cdlRecord.Filter = "Textdateien (*.txt)|*.txt"
  cdlRecord.ShowOpen
  If cdlRecord.FileName <> "" Then
    Open cdlRecord.FileName For Input As 1#
      ReDim m_iKeyCodesRec(0)
      ReDim m_lDurationsRec(0)
      i = 0
      Do Until EOF(1)
        ReDim Preserve m_iKeyCodesRec(UBound(m_iKeyCodesRec) + 1)
        ReDim Preserve m_lDurationsRec(UBound(m_lDurationsRec) + 1)
        Line Input #1, sLine
        m_iKeyCodesRec(i) = Mid(sLine, 1, InStr(sLine, " ") - 1)
        m_lDurationsRec(i) = Mid(sLine, InStr(sLine, " ") + 1)
        i = i + 1
      Loop
    Close #1
  End If
  m_iStart = 0
  m_iPosition = 0
  m_bStored = True
  EnDisableKeys

End Sub

Private Sub Play(iStart As Integer)
Dim i As Long

  EnDisableKeys
  m_iPosition = 0
  For i = iStart To UBound(m_iKeyCodesRec)
    DoEvents
    If g_iRun = RUN_STOPPED Or g_iRun = RUN_FINISHED Then 'to create a break
      Exit For
    End If
    m_iPosition = i
    g_clsMusicBeeper.GenerateBeep m_iKeyCodesRec(i), m_lDurationsRec(i)
  Next i
  If i = UBound(m_iKeyCodesRec) + 1 Then
    g_iRun = RUN_FINISHED
    g_iRunMode = RUNMODE_NOTHING
  End If

End Sub

Private Sub SaveFile()
Dim sLine As String
Dim i As Integer

  cdlRecord.InitDir = g_sPathSave
  cdlRecord.Filter = "Textdateien (*.txt)|*.txt"
  cdlRecord.ShowSave
  If cdlRecord.FileName <> "" Then
    'read the file and fill the arrays m_iKeyCodesRec and m_lDurationsRec
    Open cdlRecord.FileName For Output As 1#
      For i = 0 To UBound(m_iKeyCodesRec)
        Print #1, m_iKeyCodesRec(i) & " " & m_lDurationsRec(i)
      Next i
    Close #1
  End If
  g_iRun = RUN_FINISHED
  g_iRunMode = RUNMODE_NOTHING

End Sub

Private Sub SetLayout()
Dim iMidBorder As Integer

  Height = 2100
  Width = 2685
  
  iMidBorder = 1440 '1095
  
  SetBorder Me, EB_HORIZONTAL_BREAK, iMidBorder
  
  cmdRecord(0).Top = 480
  cmdRecord(1).Top = 480
  cmdRecord(2).Top = 480
  cmdRecord(3).Top = 480
  cmdRecord(4).Top = 120
  cmdRecord(5).Top = 120
  cmdRecord(6).Top = 120

  cmdRecord(0).Left = 120
  cmdRecord(1).Left = 735
  cmdRecord(2).Left = 1350
  cmdRecord(3).Left = 1965
  cmdRecord(4).Left = 120
  cmdRecord(5).Left = 480
  cmdRecord(6).Left = Width - 465

  lblRecord(0).Top = 150
  lblRecord(0).Left = 840
  lblRecord(0).Width = Width - 1320
  
  lblRecord(1).Top = iMidBorder + 120
  lblRecord(1).Left = 120
  lblRecord(1).Width = Width - 240
  lblRecord(1).Height = 420
  
  fraRecord.Top = 1140
  fraRecord.Left = 120
  fraRecord.Width = Width - 240
  fraRecord.Height = 300
  optRecorder(0).Top = 30
  optRecorder(1).Top = 30
  optRecorder(0).Left = 30
  optRecorder(1).Left = fraRecord.Width / 2
  optRecorder(0).Value = True
  
End Sub

Private Sub optRecorder_Click(Index As Integer)

  g_iRecordMode = Index
  'RECMODE_CLEAR = 0, RECMODE_ROUGH = 1
  
End Sub
