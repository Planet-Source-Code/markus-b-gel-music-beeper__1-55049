VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMusicBeeper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum eCommand
  EC_KEYCODE
  EC_OPEN_SETTINGS
  EC_OPEN_RECORDER
  EC_OPEN_KEYBOARD
  EC_OPEN_TEST
  EC_HIDE_MAIN
  EC_FINISH_ALL
  EC_BREAK
  EC_START_RECORDING
  EC_START_PLAY
  EC_FIXED_FREQUENCY
  EC_VARIANT_FREQUENCY
End Enum

Private m_lFrequencies() As Long
Private m_iKeyCodes() As Integer

Property Get CountFrequencies() As Integer
  CountFrequencies = UBound(m_lFrequencies)
End Property

Property Get Frequency(iIndex As Integer) As Long
  Frequency = m_lFrequencies(iIndex)
End Property

Property Get KeyCodes(iIndex As Integer) As Integer
  KeyCodes = m_iKeyCodes(iIndex)
End Property

' ##################################################################################
'
' class procedures
'
' ##################################################################################

Private Sub Class_Initialize()
  
  SetFrequencies
  SetKeyCodes

End Sub

' ##################################################################################
'
' public procedures
'
' ##################################################################################
 
Public Sub GenerateBeep(iKeycode As Integer, lDuration As Long)
Dim i As Integer
  
  If iKeycode = 32 Then 'simulate break
    Sleep lDuration
  End If
  
  For i = 0 To UBound(m_iKeyCodes) 'array is searched for the right index
    If iKeycode = m_iKeyCodes(i) Then
      Beep m_lFrequencies(i), lDuration
      If g_iRunMode = RUNMODE_RECORD And (g_iMode = MODE_FIXED) _
      And (g_iRun <> RUN_STOPPED) Then 'to create a break
        g_frmRecord.RecordKeys iKeycode 'recording is started
      End If
    End If
  Next i

End Sub

Public Function GetFrequency(iIndex As Integer) As Long

  GetFrequency = m_lFrequencies(iIndex)
  
End Function

'Form_KeyDown-event of all forms
Public Sub GetKey(KeyCode As Integer, Shift As Integer)

  Select Case KeyCode
  Case 19 'BREAK
    ExecuteCommand EC_BREAK
  
  Case 27  'ESC
    ExecuteCommand EC_FINISH_ALL
  
  Case 112 'F1
    ExecuteCommand EC_FIXED_FREQUENCY
  
  Case 113 'F2
    ExecuteCommand EC_VARIANT_FREQUENCY
  
  Case 114 'F3
    'nothing happens
  
  Case 115 'F4
    'nothing happens
  
  Case 116 'F5
    ExecuteCommand EC_OPEN_RECORDER
  
  Case 117 'F6
    ExecuteCommand EC_OPEN_KEYBOARD
  
  Case 118 'F7
    ExecuteCommand EC_OPEN_TEST
  
  Case 119 'F8
    ExecuteCommand EC_HIDE_MAIN
  
  Case 120 'F9
    ExecuteCommand EC_START_RECORDING
  
  Case 121  'F10
    ExecuteCommand EC_START_PLAY
  
  Case 122  'F11
    'nothing happens
  
  Case 123 'F12
    ExecuteCommand EC_OPEN_SETTINGS
    
  Case Else
    ExecuteCommand EC_KEYCODE, KeyCode
    
  End Select

End Sub

Public Sub SetKeyCodes()
Dim i As Integer

  ReDim m_iKeyCodes(48)
  i = 0
  
  m_iKeyCodes(i + 1) = Asc("1") '49
  m_iKeyCodes(i + 2) = Asc("2") '50
  m_iKeyCodes(i + 3) = Asc("3") '51
  m_iKeyCodes(i + 4) = Asc("4") '52
  m_iKeyCodes(i + 5) = Asc("5") '53
  m_iKeyCodes(i + 6) = Asc("6") '54
  m_iKeyCodes(i + 7) = Asc("7") '55
  m_iKeyCodes(i + 8) = Asc("8") '56
  m_iKeyCodes(i + 9) = Asc("9") '57
  m_iKeyCodes(i + 10) = Asc("0") '48
  m_iKeyCodes(i + 13) = Asc("Q") '81
  m_iKeyCodes(i + 14) = Asc("W") '87
  m_iKeyCodes(i + 15) = Asc("E") '69
  m_iKeyCodes(i + 16) = Asc("R") '82
  m_iKeyCodes(i + 17) = Asc("T") '84
  m_iKeyCodes(i + 19) = Asc("U") '85
  m_iKeyCodes(i + 20) = Asc("I") '73
  m_iKeyCodes(i + 21) = Asc("O") '79
  m_iKeyCodes(i + 22) = Asc("P") '80
  m_iKeyCodes(i + 25) = Asc("A") '65
  m_iKeyCodes(i + 26) = Asc("S") '83
  m_iKeyCodes(i + 27) = Asc("D") '68
  m_iKeyCodes(i + 28) = Asc("F") '70
  m_iKeyCodes(i + 29) = Asc("G") '71
  m_iKeyCodes(i + 30) = Asc("H") '72
  m_iKeyCodes(i + 31) = Asc("J") '74
  m_iKeyCodes(i + 32) = Asc("K") '75
  m_iKeyCodes(i + 33) = Asc("L") '76
  m_iKeyCodes(i + 38) = Asc("X") '88
  m_iKeyCodes(i + 39) = Asc("C") '67
  m_iKeyCodes(i + 40) = Asc("V") '86
  m_iKeyCodes(i + 41) = Asc("B") '66
  m_iKeyCodes(i + 42) = Asc("N") '78
  m_iKeyCodes(i + 43) = Asc("M") '77
  m_iKeyCodes(i + 44) = 188 ',
  m_iKeyCodes(i + 45) = 190 '.

  Select Case g_eLanguage
  Case EL_ENGLISH
    m_iKeyCodes(i) = Asc("`") '96
    m_iKeyCodes(i + 11) = Asc("-") '45
    m_iKeyCodes(i + 12) = Asc("=") '61
    m_iKeyCodes(i + 18) = Asc("Y") '89
    m_iKeyCodes(i + 23) = Asc("[") '91
    m_iKeyCodes(i + 24) = Asc("]") '93
    m_iKeyCodes(i + 34) = Asc(";") '59
    m_iKeyCodes(i + 35) = Asc("'") '39
    m_iKeyCodes(i + 36) = Asc("\") '92 'in German two different characters#<
    m_iKeyCodes(i + 37) = Asc("Z") '90
    m_iKeyCodes(i + 46) = Asc("/") '47
  
  Case EL_GERMAN
    m_iKeyCodes(i) = 220  '^
    m_iKeyCodes(i + 11) = 219 'ß
    m_iKeyCodes(i + 12) = 221 '´
    m_iKeyCodes(i + 18) = Asc("Z") '90
    m_iKeyCodes(i + 23) = 186 'Ü
    m_iKeyCodes(i + 24) = 187 '+
    m_iKeyCodes(i + 34) = 192 'Ö
    m_iKeyCodes(i + 35) = 222 'Ä
    m_iKeyCodes(i + 36) = 191 '#
'    m_iKeyCodes(i + 37) = 226 '<
    m_iKeyCodes(i + 37) = Asc("Y") '89
    m_iKeyCodes(i + 46) = 189 '-
  End Select

End Sub

' ##################################################################################
'
' private procedures
'
' ##################################################################################

Private Sub ExecuteCommand(eCom As eCommand, Optional iKeycode As Integer)
Dim sCaption As String
Dim sCloseText As String 'finish-dialogue

  Select Case eCom
  Case EC_OPEN_SETTINGS
    g_frmSettings.Show vbModal
  
  Case EC_OPEN_RECORDER
    g_frmRecord.Visible = Not g_frmRecord.Visible
  
  Case EC_OPEN_KEYBOARD
    g_frmKeys.Visible = Not g_frmKeys.Visible
    g_frmKeys.m_iFocus = 0
  
  Case EC_OPEN_TEST
    g_frmTest.Visible = Not g_frmTest.Visible
  
  Case EC_HIDE_MAIN
    If g_frmMusicBeeper.Width = 1 Then
      g_frmMusicBeeper.Width = 5625
      g_frmMusicBeeper.Height = 6135
    Else 'minimize
      g_frmMusicBeeper.Width = 1
      g_frmMusicBeeper.Height = 1
    End If
  
  Case EC_FINISH_ALL
    Select Case g_eLanguage
    Case EL_GERMAN
      sCaption = "MusikBeeper schließen?"
      sCloseText = "Programm beenden?"
    Case EL_ENGLISH
      sCaption = "Close MusicBeeper?"
      sCloseText = "Finish application?"
    End Select
    If MsgBox(sCloseText, vbYesNo + vbQuestion, sCaption) = vbYes Then
      End
    End If
  
  Case EC_BREAK
    g_frmRecord.Show
    g_frmRecord.cmdRecord_Click 1
  
  Case EC_START_RECORDING
    If g_iRun = RUN_FINISHED Or g_iRun = RUN_NOTHING Then
      g_frmRecord.Show
      g_frmRecord.cmdRecord_Click 0
    ElseIf g_iRun = RUN_STARTED And g_iRunMode = RUNMODE_RECORD Then
      g_frmRecord.Show
      g_frmRecord.cmdRecord_Click 2
    End If

  Case EC_START_PLAY
    If g_iRun = RUN_FINISHED Then
      g_frmRecord.Show
      g_frmRecord.cmdRecord_Click 3
    ElseIf g_iRun = RUN_STARTED And g_iRunMode = RUNMODE_PLAY Then
      g_frmRecord.Show
      g_frmRecord.cmdRecord_Click 2
    End If

  Case EC_FIXED_FREQUENCY
    g_iMode = MODE_FIXED
    g_frmMusicBeeper.optSound(0).Value = True
  
  Case EC_VARIANT_FREQUENCY
    g_iMode = MODE_VARIANT
    g_frmMusicBeeper.optSound(1).Value = True
  
  Case Else
    If g_iMode = MODE_FIXED Then 'duration fixed
      GenerateBeep iKeycode, g_lDuration 'generate beep
    ElseIf g_iMode = MODE_VARIANT Then 'scan-Mode / duration variant
      'nothing to do
    End If
    If g_frmKeys.Visible And iKeycode > 47 Then
      g_frmKeys.m_iFocus = 0
      g_frmKeys.cmdKey(iKeycode).SetFocus
    End If
    
  End Select

  If g_frmRecord.Visible Then
    g_frmRecord.PatchLabel
  End If
   
End Sub

Private Sub SetFrequencies()
Dim i As Integer
  
  i = 0
  ReDim m_lFrequencies(i + 72)
  
  'the first 25 are set
  '37 is the lowest frequency generated by the PC
  m_lFrequencies(i) = 65  '64 'first one with assigned key
  m_lFrequencies(i + 1) = 73 '72
  m_lFrequencies(i + 2) = 82 '81
  m_lFrequencies(i + 3) = 87 '86
  m_lFrequencies(i + 4) = 98
  m_lFrequencies(i + 5) = 110
  m_lFrequencies(i + 6) = 123 '122
  m_lFrequencies(i + 7) = 131 '128
  m_lFrequencies(i + 8) = 147 '144
  m_lFrequencies(i + 9) = 165 '161
  m_lFrequencies(i + 10) = 175 '171
  m_lFrequencies(i + 11) = 196 '195
  m_lFrequencies(i + 12) = 220
  m_lFrequencies(i + 13) = 247
  m_lFrequencies(i + 14) = 262
  m_lFrequencies(i + 15) = 294
  m_lFrequencies(i + 16) = 330
  m_lFrequencies(i + 17) = 349
  m_lFrequencies(i + 18) = 392
  m_lFrequencies(i + 19) = 440
  m_lFrequencies(i + 20) = 494
  m_lFrequencies(i + 21) = 523
  m_lFrequencies(i + 22) = 587
  m_lFrequencies(i + 23) = 659
  m_lFrequencies(i + 24) = 698
  m_lFrequencies(i + 25) = 783
  'the others are generated by doubling
  For i = 26 To 72
    m_lFrequencies(i) = 2 * (m_lFrequencies(i - 7))
  Next i

End Sub

