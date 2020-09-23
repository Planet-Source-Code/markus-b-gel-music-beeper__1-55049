Attribute VB_Name = "modMain"
Option Explicit

'##############################################################
'
' The beep function (which is basic to this program)
' requires Windows NT 3.1 or later.
' With Windows 95 / 98 this program will not run effectively.
'
'##############################################################

Global g_iMode As Integer 'beep-mode (fixed or variable duration)
Global g_iRun As Integer 'differs between recording and playing mode
Global g_iRunMode As Integer 'flag to store the play/recording state
Global g_iRecordMode As Integer 'flag to store the recording sound mode
Global g_lProgId As Long 'ID of the started instance of waveplayer
Global g_eRecDocked As eDocking 'flag to store the docking state of g_frmRecord
Global g_eKeyDocked As eDocking 'flag to store the docking state of g_frmKeys
Global g_eTestDocked As eDocking 'flag to store the docking state of g_frmTest

Global g_lDuration As Long 'fixed duration
Global g_lInterval As Long 'interval for checking the keystate
Global g_eLanguage As eLanguage 'used language for labels and messages
Global g_eKeylanguage As eLanguage 'used language for keyboard
Global g_sPathSave As String 'path of the stored (saved / opened) files

Global g_frmKeys As frmKeys
Global g_frmMusicBeeper As frmMusicBeeper
Global g_frmRecord As frmRecord
Global g_frmSettings As frmSettings
Global g_frmTest As frmTest
Global g_clsMusicBeeper As clsMusicBeeper

Global Const MODE_FIXED = 0
Global Const MODE_VARIANT = 1

Global Const RUN_NOTHING = 0
Global Const RUN_STARTED = 1
Global Const RUN_STOPPED = 2
Global Const RUN_FINISHED = 3

Global Const RUNMODE_NOTHING = 0
Global Const RUNMODE_RECORD = 1
Global Const RUNMODE_PLAY = 2

Global Const RECMODE_CLEAR = 0
Global Const RECMODE_ROUGH = 1

Enum eLanguage
  EL_GERMAN
  EL_ENGLISH
End Enum

Enum eDocking
  ED_NO
  ED_TOP
  ED_LEFT
  ED_BOTTOM
  ED_RIGHT
End Enum

Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Declare Function Beep Lib "kernel32" _
  (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long

'start procedure
Public Sub Main()
  
  Set g_frmKeys = New frmKeys
  Set g_frmMusicBeeper = New frmMusicBeeper
  Set g_frmRecord = New frmRecord
  Set g_frmSettings = New frmSettings
  Set g_frmTest = New frmTest
  Set g_clsMusicBeeper = New clsMusicBeeper 'Class_Initialize must happen after getting the settings

  If App.PrevInstance = True Then
    'if another instance has already been started
    AppActivate "MusicBeeper"
    End
  End If
  
  'get settings from registry
  g_lDuration = GetSetting("MusicBeeper", "Settings", "Duration", "200")
  g_lInterval = GetSetting("MusicBeeper", "Settings", "Interval", "10")
  g_eLanguage = GetSetting("MusicBeeper", "Settings", "Language", EL_ENGLISH)
  g_eKeylanguage = GetSetting("MusicBeeper", "Settings", "Keylanguage", EL_ENGLISH)
  g_sPathSave = GetSetting("MusicBeeper", "Settings", "PathSave", App.Path & "\Aufnahmen")
  
  g_lInterval = 10
  g_frmMusicBeeper.tmrScan.Interval = 10
  
  g_frmMusicBeeper.Show
  
  g_iRun = RUN_NOTHING
  g_iRunMode = RUNMODE_NOTHING

End Sub

