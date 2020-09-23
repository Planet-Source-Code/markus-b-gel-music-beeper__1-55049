VERSION 5.00
Begin VB.Form frmTest 
   BorderStyle     =   0  'Kein
   Caption         =   " Hörtest"
   ClientHeight    =   2655
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2055
   Icon            =   "frmTest.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   2055
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdTest 
      Caption         =   "Schließen"
      Height          =   375
      Index           =   2
      Left            =   240
      TabIndex        =   4
      Top             =   2040
      Width           =   1575
   End
   Begin VB.CommandButton cmdTest 
      Height          =   375
      Index           =   1
      Left            =   240
      TabIndex        =   1
      Top             =   1560
      Width           =   1575
   End
   Begin VB.TextBox txtTest 
      Enabled         =   0   'False
      Height          =   405
      Left            =   240
      TabIndex        =   2
      Top             =   480
      Width           =   1575
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "Start"
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Line lin 
      BorderColor     =   &H80000015&
      Index           =   15
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   720
   End
   Begin VB.Line lin 
      BorderColor     =   &H80000010&
      Index           =   14
      X1              =   0
      X2              =   0
      Y1              =   360
      Y2              =   840
   End
   Begin VB.Line lin 
      BorderColor     =   &H80000016&
      Index           =   13
      X1              =   0
      X2              =   0
      Y1              =   240
      Y2              =   960
   End
   Begin VB.Line lin 
      BorderColor     =   &H80000014&
      Index           =   12
      X1              =   0
      X2              =   0
      Y1              =   360
      Y2              =   840
   End
   Begin VB.Line lin 
      BorderColor     =   &H80000015&
      Index           =   11
      X1              =   0
      X2              =   0
      Y1              =   240
      Y2              =   840
   End
   Begin VB.Line lin 
      BorderColor     =   &H80000010&
      Index           =   10
      X1              =   0
      X2              =   0
      Y1              =   240
      Y2              =   840
   End
   Begin VB.Line lin 
      BorderColor     =   &H80000016&
      Index           =   9
      X1              =   0
      X2              =   0
      Y1              =   120
      Y2              =   840
   End
   Begin VB.Line lin 
      BorderColor     =   &H80000014&
      Index           =   8
      X1              =   0
      X2              =   0
      Y1              =   240
      Y2              =   840
   End
   Begin VB.Line lin 
      BorderColor     =   &H80000015&
      Index           =   3
      X1              =   15
      X2              =   0
      Y1              =   135
      Y2              =   840
   End
   Begin VB.Line lin 
      BorderColor     =   &H80000016&
      Index           =   1
      X1              =   0
      X2              =   0
      Y1              =   240
      Y2              =   840
   End
   Begin VB.Line lin 
      BorderColor     =   &H80000010&
      Index           =   6
      X1              =   0
      X2              =   0
      Y1              =   840
      Y2              =   120
   End
   Begin VB.Line lin 
      BorderColor     =   &H80000015&
      Index           =   7
      X1              =   0
      X2              =   0
      Y1              =   840
      Y2              =   240
   End
   Begin VB.Line lin 
      BorderColor     =   &H80000014&
      Index           =   4
      X1              =   0
      X2              =   0
      Y1              =   360
      Y2              =   960
   End
   Begin VB.Line lin 
      BorderColor     =   &H80000016&
      Index           =   5
      X1              =   0
      X2              =   0
      Y1              =   120
      Y2              =   960
   End
   Begin VB.Line lin 
      BorderColor     =   &H80000010&
      Index           =   2
      X1              =   0
      X2              =   0
      Y1              =   120
      Y2              =   840
   End
   Begin VB.Line lin 
      BorderColor     =   &H80000014&
      Index           =   0
      X1              =   0
      X2              =   0
      Y1              =   240
      Y2              =   960
   End
   Begin VB.Label lblTest 
      Caption         =   "Frequenz:"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum eSequence
  ES_FIRST
  ES_SECOND
  ES_THIRD
  ES_FOURTH
  ES_FINISHED
End Enum

Private m_lLastToneHeared As Long
Private m_lFirstToneNotHeared As Long
Private m_lEnd As Long
Private m_i As Integer
Private m_eIncreasingSequence As Integer

'properties for moving the form
Private m_bDragging As Boolean 'flag: window is moved
Private m_lDownX As Long 'point on which the mouse was
Private m_lDownY As Long 'put down over the form

Public m_iTop As Integer 'top border of g_frmMusicBeeper
Public m_iBottom As Integer  'bottom border of g_frmMusicBeeper
Public m_iLeft As Integer  'left border of g_frmMusicBeeper
Public m_iRight As Integer  'right border of g_frmMusicBeeper

Private m_sCaption As String
Private m_sHeard As String
Private m_sNotHeared As String
Private m_sGoOn As String
Private m_sClose As String
Private m_sFinished As String
Private m_sBat As String
Private m_sError As String

' ##################################################################################
'
' form events
'
' ##################################################################################

Private Sub Form_Load()
  
  SetLanguage
  SetBorder Me, EB_SIMPLE
  
  m_i = 40
  m_eIncreasingSequence = ES_FIRST
  cmdTest(1).Visible = False
  
  Top = g_frmMusicBeeper.Top + (g_frmMusicBeeper.Height - Height) / 2
  Left = g_frmMusicBeeper.Left - Width
  g_eTestDocked = ED_RIGHT
  
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
  FormMove Me, g_eTestDocked
  
  Select Case g_eTestDocked
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

Private Sub cmdTest_Click(Index As Integer)
  
  Select Case Index
  
  Case 0
    cmdTest(0).Caption = m_sHeard
    cmdTest(1).Visible = True
    cmdTest(1).Caption = m_sNotHeared
    
    Select Case m_eIncreasingSequence
    
    Case ES_FIRST
      IncreaseFrequency1
    
    Case ES_SECOND
      IncreaseFrequency2 100
    
    Case ES_THIRD
      IncreaseFrequency2 10
    
    Case ES_FOURTH
      IncreaseFrequency2 1
    
    Case ES_FINISHED
      MsgBox m_sFinished & m_lLastToneHeared & " Hertz."
      cmdTest(1).Visible = False
  
    End Select
  
  Case 1
    m_eIncreasingSequence = m_eIncreasingSequence + 1
    m_i = m_lLastToneHeared
    m_lEnd = m_lFirstToneNotHeared
    cmdTest(0).Caption = m_sGoOn
  
  Case 2
    Unload Me
  
  End Select
  
End Sub

' ##################################################################################
'
' public procedures
'
' ##################################################################################

Public Sub SetLanguage()

  Select Case g_eLanguage
  
  Case EL_ENGLISH
    m_sCaption = " Ear Test"
    m_sHeard = "Heared"
    m_sNotHeared = "Not heared"
    m_sGoOn = "Go on"
    m_sClose = "Close"
    m_sCaption = "Frequency"
    m_sFinished = "The last heared frequency was "
    m_sBat = "Incredible! You got bat's ears!"
    m_sError = "Error ..."
    
  Case EL_GERMAN
    m_sCaption = " Hörtest"
    m_sHeard = "Gehört"
    m_sNotHeared = "Nicht gehört"
    m_sGoOn = "Weiter"
    m_sClose = "Schließen"
    m_sCaption = "Frequenz"
    m_sFinished = "Die höchste gehörte Frequenz war "
    m_sBat = "Phantastisch! Ein Gehör wie eine Fledermaus!"
    m_sError = "Fehler ..."
  End Select
  
  cmdTest(2).Caption = m_sClose
  Caption = m_sCaption
  lblTest.Caption = m_sCaption

End Sub

' ##################################################################################
'
' private procedures
'
' ##################################################################################

Private Sub IncreaseFrequency1()

  If m_i >= g_clsMusicBeeper.CountFrequencies Then
    MsgBox m_sBat
    m_eIncreasingSequence = ES_FINISHED
    Exit Sub
  End If
  
  m_lLastToneHeared = g_clsMusicBeeper.Frequency(m_i)
  m_i = m_i + 1
  txtTest.Text = g_clsMusicBeeper.Frequency(m_i) & " Hz"
  Beep g_clsMusicBeeper.Frequency(m_i), 5 * g_lDuration
  m_lFirstToneNotHeared = g_clsMusicBeeper.Frequency(m_i)
  
End Sub

Private Sub IncreaseFrequency2(iInterval As Integer)
  
  If m_i >= m_lEnd Then
    MsgBox m_sError
    Exit Sub
  End If
  
  m_lLastToneHeared = m_i
  txtTest.Text = m_i & " Hz"
  Beep m_i, 5 * g_lDuration
  m_i = m_i + iInterval
  m_lFirstToneNotHeared = m_i
  
End Sub

