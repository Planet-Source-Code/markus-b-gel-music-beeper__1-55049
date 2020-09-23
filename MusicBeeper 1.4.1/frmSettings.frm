VERSION 5.00
Begin VB.Form frmSettings 
   BorderStyle     =   0  'Kein
   ClientHeight    =   4005
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6735
   Icon            =   "frmSettings.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4005
   ScaleWidth      =   6735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.Frame fraKeylanguage 
      BorderStyle     =   0  'Kein
      Height          =   375
      Left            =   360
      TabIndex        =   12
      Top             =   600
      Width           =   6015
      Begin VB.OptionButton optKeylanguage 
         Height          =   195
         Index           =   1
         Left            =   4080
         TabIndex        =   14
         Top             =   120
         Width           =   1815
      End
      Begin VB.OptionButton optKeylanguage 
         Height          =   195
         Index           =   0
         Left            =   2160
         TabIndex        =   13
         Top             =   120
         Width           =   1815
      End
      Begin VB.Label lblSettings 
         Caption         =   "Pfad"
         Height          =   255
         Index           =   4
         Left            =   0
         TabIndex        =   15
         Top             =   120
         Width           =   1620
      End
   End
   Begin VB.TextBox txtSettings 
      Height          =   285
      Index           =   2
      Left            =   360
      TabIndex        =   3
      Top             =   2550
      Width           =   1080
   End
   Begin VB.TextBox txtSettings 
      Height          =   285
      Index           =   3
      Left            =   3600
      TabIndex        =   4
      Top             =   2550
      Width           =   1080
   End
   Begin VB.Frame fraLanguage 
      BorderStyle     =   0  'Kein
      Height          =   375
      Left            =   360
      TabIndex        =   8
      Top             =   240
      Width           =   6015
      Begin VB.OptionButton optLanguage 
         Height          =   195
         Index           =   0
         Left            =   2160
         TabIndex        =   0
         Top             =   120
         Width           =   1815
      End
      Begin VB.OptionButton optLanguage 
         Height          =   195
         Index           =   1
         Left            =   4080
         TabIndex        =   1
         Top             =   120
         Width           =   1815
      End
      Begin VB.Label lblSettings 
         Caption         =   "Pfad"
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   11
         Top             =   120
         Width           =   1620
      End
   End
   Begin VB.CommandButton cmdSettings 
      Cancel          =   -1  'True
      Height          =   375
      Index           =   1
      Left            =   3720
      TabIndex        =   6
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CommandButton cmdSettings 
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   1440
      TabIndex        =   5
      Top             =   3360
      Width           =   1335
   End
   Begin VB.TextBox txtSettings 
      Height          =   285
      Index           =   1
      Left            =   360
      TabIndex        =   2
      Top             =   1560
      Width           =   6000
   End
   Begin VB.Line linSettings 
      BorderColor     =   &H80000010&
      Index           =   2
      X1              =   240
      X2              =   6480
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line linSettings 
      BorderColor     =   &H80000014&
      Index           =   3
      X1              =   240
      X2              =   6480
      Y1              =   2055
      Y2              =   2055
   End
   Begin VB.Line linSettings 
      BorderColor     =   &H80000010&
      Index           =   0
      X1              =   240
      X2              =   6480
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Line linSettings 
      BorderColor     =   &H80000014&
      Index           =   1
      X1              =   240
      X2              =   6480
      Y1              =   1095
      Y2              =   1095
   End
   Begin VB.Line lin 
      BorderColor     =   &H80000014&
      Index           =   0
      X1              =   0
      X2              =   0
      Y1              =   240
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
      BorderColor     =   &H80000016&
      Index           =   5
      X1              =   0
      X2              =   0
      Y1              =   120
      Y2              =   960
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
      BorderColor     =   &H80000015&
      Index           =   7
      X1              =   0
      X2              =   0
      Y1              =   840
      Y2              =   240
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
      BorderColor     =   &H80000016&
      Index           =   1
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
      BorderColor     =   &H80000014&
      Index           =   8
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
      BorderColor     =   &H80000010&
      Index           =   10
      X1              =   0
      X2              =   0
      Y1              =   240
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
      BorderColor     =   &H80000014&
      Index           =   12
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
      BorderColor     =   &H80000010&
      Index           =   14
      X1              =   0
      X2              =   0
      Y1              =   360
      Y2              =   840
   End
   Begin VB.Line lin 
      BorderColor     =   &H80000015&
      Index           =   15
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   720
   End
   Begin VB.Line linSettings 
      BorderColor     =   &H80000014&
      Index           =   7
      X1              =   240
      X2              =   6480
      Y1              =   3045
      Y2              =   3045
   End
   Begin VB.Line linSettings 
      BorderColor     =   &H80000010&
      Index           =   6
      X1              =   240
      X2              =   6480
      Y1              =   3030
      Y2              =   3030
   End
   Begin VB.Label lblSettings 
      Caption         =   "Pfad"
      Height          =   255
      Index           =   3
      Left            =   3600
      TabIndex        =   10
      Top             =   2310
      Width           =   2820
   End
   Begin VB.Label lblSettings 
      Caption         =   "Pfad"
      Height          =   255
      Index           =   2
      Left            =   360
      TabIndex        =   9
      Top             =   2310
      Width           =   2940
   End
   Begin VB.Label lblSettings 
      Caption         =   "Pfad"
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   7
      Top             =   1320
      Width           =   3300
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ##################################################################################
'
' form events
'
' ##################################################################################

Private Sub Form_Load()

  SetBorder Me, EB_SIMPLE
  SetLanguage
  SetValues
  cmdSettings(0).Enabled = False
  
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  
  If KeyCode = 120 Then 'F9
    g_clsMusicBeeper.GetKey KeyCode, Shift
  End If
  
End Sub

' ##################################################################################
'
' control events
'
' ##################################################################################

Private Sub cmdSettings_Click(Index As Integer)
  
  Select Case Index
  Case 0
    If Save = 0 Then
      Me.Hide
      g_frmMusicBeeper.Show
    End If
  Case 1
    Me.Hide
    g_frmMusicBeeper.Show
  End Select

End Sub

Private Sub optKeylanguage_Click(Index As Integer)

  cmdSettings(0).Enabled = True

End Sub

Private Sub optLanguage_Click(Index As Integer)

  cmdSettings(0).Enabled = True

End Sub

Private Sub txtSettings_Change(Index As Integer)

  cmdSettings(0).Enabled = True

End Sub

' ##################################################################################
'
' private procedures
'
' ##################################################################################

Private Function Save() As Integer
  
  g_eLanguage = IIf(optLanguage(0), EL_GERMAN, EL_ENGLISH)
  g_eKeylanguage = IIf(optKeylanguage(0), EL_GERMAN, EL_ENGLISH)
  g_sPathSave = txtSettings(1).Text
  g_lDuration = txtSettings(2).Text
  g_lInterval = txtSettings(3).Text
    
  Save = 0
  
  SaveSetting "MusicBeeper", "Settings", "Language", g_eLanguage
  SaveSetting "MusicBeeper", "Settings", "Keylanguage", g_eKeylanguage
  SaveSetting "MusicBeeper", "Settings", "PathSave", g_sPathSave
  SaveSetting "MusicBeeper", "Settings", "Duration", g_lDuration
  SaveSetting "MusicBeeper", "Settings", "Interval", g_lInterval
  
  g_frmRecord.SetLanguage
  g_frmMusicBeeper.SetLanguage
  g_frmTest.SetLanguage
  g_frmKeys.SetCaptions
  SetLanguage
  g_clsMusicBeeper.SetKeyCodes

End Function

Private Sub SetLanguage()
  
  Select Case g_eLanguage
  Case EL_GERMAN
    Caption = "Einstellungen"
    lblSettings(0).Caption = "Sprache:"
    optLanguage(0).Caption = "Deutsch"
    optLanguage(1).Caption = "Englisch"
    lblSettings(4).Caption = "Tastaturbelegung:"
    optKeylanguage(0).Caption = "Deutsch"
    optKeylanguage(1).Caption = "Englisch"
    lblSettings(1).Caption = "Speicherort für Dateien (Voreinstellung):"
    lblSettings(2).Caption = "Feste Dauer der Töne:"
    lblSettings(3).Caption = "Intervall des Tastatur-Scans:"
    cmdSettings(0).Caption = "Speichern"
    cmdSettings(1).Caption = "Abbrechen"
  Case EL_ENGLISH
    Caption = "Settings"
    lblSettings(0).Caption = "Language:"
    optLanguage(0).Caption = "German"
    optLanguage(1).Caption = "English"
    lblSettings(4).Caption = "Keylanguage:"
    optKeylanguage(0).Caption = "German"
    optKeylanguage(1).Caption = "English"
    lblSettings(1).Caption = "Default path of stored files:"
    lblSettings(2).Caption = "Fixed duration of beep:"
    lblSettings(3).Caption = "Interval for key scan:"
    cmdSettings(0).Caption = "Save"
    cmdSettings(1).Caption = "Cancel"
  End Select
 
End Sub

Private Sub SetValues()

  optLanguage(0).Value = (g_eLanguage = EL_GERMAN)
  optLanguage(1).Value = (g_eLanguage = EL_ENGLISH)
  optKeylanguage(0).Value = (g_eKeylanguage = EL_GERMAN)
  optKeylanguage(1).Value = (g_eKeylanguage = EL_ENGLISH)
  txtSettings(1).Text = g_sPathSave
  txtSettings(2).Text = g_lDuration
  txtSettings(3).Text = g_lInterval

End Sub

