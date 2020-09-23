Attribute VB_Name = "modScanKeyboard"
Option Explicit

Private Type KeyboardBytes
  kbByte(0 To 255) As Byte
End Type

Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Private kbArray As KeyboardBytes

' ##################################################################################
'
' public procedures
'
' ##################################################################################

'check if a key is pressed or has been pressed at least before
Public Sub ScanCurrentKeyState()
Dim b As Byte
Dim l As Long
Dim bFound As Boolean

  bFound = False
  For b = 48 To 226
    l = GetAsyncKeyState(b)
    If l <> 0 Then
      BeepTone b
      bFound = True
      Exit For
    End If
  Next b
  
  b = IIf(bFound And b <> 227, b, 32) 'generate pause
  
  If Not (b = 32 And g_iMode = MODE_FIXED) Then
    If g_iRunMode = RUNMODE_RECORD Then
      'stores the currently pressed key
      g_frmRecord.RecordIntervals CInt(b)
    End If
  End If

End Sub

' ##################################################################################
'
' private procedures
'
' ##################################################################################

Private Sub BeepTone(Tone As Byte)
  
  If g_iMode = 1 Then
    'generates a beep
    g_clsMusicBeeper.GenerateBeep CInt(Tone), g_lInterval
  End If
  
End Sub

