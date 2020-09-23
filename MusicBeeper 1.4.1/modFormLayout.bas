Attribute VB_Name = "modFormLayout"
Option Explicit

Enum eBorderType
  EB_SIMPLE
  EB_HORIZONTAL_BREAK
  EB_VERTICAL_BREAK
End Enum

' ##################################################################################
'
' public procedures
'
' ##################################################################################

'sets properties of a form when it is moved
Public Sub FormMove(frmForm As Form, eDocked As eDocking)

  With frmForm
    
    .m_iTop = g_frmMusicBeeper.Top
    .m_iBottom = g_frmMusicBeeper.Top + g_frmMusicBeeper.Height
    .m_iLeft = g_frmMusicBeeper.Left
    .m_iRight = g_frmMusicBeeper.Left + g_frmMusicBeeper.Width
    
    eDocked = ED_NO
    
    'check if form can be docked
    If .Top + .Height - 90 < .m_iTop And .m_iTop < .Top + .Height + 90 Then
      eDocked = ED_TOP
    End If

    If .Left + .Width - 90 < .m_iLeft And .m_iLeft < .Left + .Width + 90 Then
      eDocked = ED_LEFT
    End If

    If .Top - 90 < .m_iBottom And .m_iBottom < .Top + 90 Then
      eDocked = ED_BOTTOM
    End If

    If .Left - 90 < .m_iRight And .m_iRight < .Left + 90 Then
      eDocked = ED_RIGHT
    End If
  
  End With

End Sub

'paints the border of a form
Public Sub SetBorder(frmForm As Form, eBorder As eBorderType, Optional iBreak As Integer)

  frmForm.lin(0).X1 = 0
  frmForm.lin(0).X2 = frmForm.Width
  frmForm.lin(0).Y1 = 0
  frmForm.lin(0).Y2 = 0
  
  frmForm.lin(1).X1 = 15
  frmForm.lin(1).X2 = frmForm.Width - 15
  frmForm.lin(1).Y1 = 15
  frmForm.lin(1).Y2 = 15
  
  frmForm.lin(2).X1 = 60
  frmForm.lin(2).X2 = frmForm.Width - 60
  frmForm.lin(2).Y1 = 60
  frmForm.lin(2).Y2 = 60
  
  frmForm.lin(3).X1 = 75
  frmForm.lin(3).X2 = frmForm.Width - 75
  frmForm.lin(3).Y1 = 75
  frmForm.lin(3).Y2 = 75
  
  frmForm.lin(4).X1 = 0
  frmForm.lin(4).X2 = 0
  frmForm.lin(4).Y1 = 0
  frmForm.lin(4).Y2 = frmForm.Height

  frmForm.lin(5).X1 = 15
  frmForm.lin(5).X2 = 15
  frmForm.lin(5).Y1 = 15
  frmForm.lin(5).Y2 = frmForm.Height - 15

  frmForm.lin(6).X1 = 60
  frmForm.lin(6).X2 = 60
  frmForm.lin(6).Y1 = 60
  
  frmForm.lin(7).X1 = 75
  frmForm.lin(7).X2 = 75
  frmForm.lin(7).Y1 = 75

  frmForm.lin(8).X1 = frmForm.Width - 90
  frmForm.lin(8).X2 = frmForm.Width - 90
  frmForm.lin(8).Y1 = 75
  
  frmForm.lin(9).X1 = frmForm.Width - 75
  frmForm.lin(9).X2 = frmForm.Width - 75
  frmForm.lin(9).Y1 = 60

  Select Case eBorder
  Case EB_SIMPLE
    frmForm.lin(6).Y2 = frmForm.Height - 60
    frmForm.lin(7).Y2 = frmForm.Height - 75
    frmForm.lin(8).Y2 = frmForm.Height - 75
    frmForm.lin(9).Y2 = frmForm.Height - 60
  Case EB_HORIZONTAL_BREAK
    frmForm.lin(6).Y2 = iBreak + 15
    frmForm.lin(7).Y2 = iBreak
    frmForm.lin(8).Y2 = iBreak
    frmForm.lin(9).Y2 = iBreak + 15
  End Select

  frmForm.lin(10).X1 = frmForm.Width - 30
  frmForm.lin(10).X2 = frmForm.Width - 30
  frmForm.lin(10).Y1 = 15
  frmForm.lin(10).Y2 = frmForm.Height - 15

  frmForm.lin(11).X1 = frmForm.Width - 15
  frmForm.lin(11).X2 = frmForm.Width - 15
  frmForm.lin(11).Y1 = 0
  frmForm.lin(11).Y2 = frmForm.Height - 0
  
  frmForm.lin(12).X1 = 75
  frmForm.lin(12).X2 = frmForm.Width - 75
  frmForm.lin(12).Y1 = frmForm.Height - 90
  frmForm.lin(12).Y2 = frmForm.Height - 90
  
  frmForm.lin(13).X1 = 60
  frmForm.lin(13).X2 = frmForm.Width - 60
  frmForm.lin(13).Y1 = frmForm.Height - 75
  frmForm.lin(13).Y2 = frmForm.Height - 75
  
  frmForm.lin(14).X1 = 15
  frmForm.lin(14).X2 = frmForm.Width - 15
  frmForm.lin(14).Y1 = frmForm.Height - 30
  frmForm.lin(14).Y2 = frmForm.Height - 30

  frmForm.lin(15).X1 = 0
  frmForm.lin(15).X2 = frmForm.Width
  frmForm.lin(15).Y1 = frmForm.Height - 15
  frmForm.lin(15).Y2 = frmForm.Height - 15

  Select Case eBorder
  Case EB_SIMPLE
  
  Case EB_HORIZONTAL_BREAK
    frmForm.lin(16).X1 = 75
    frmForm.lin(16).X2 = frmForm.Width - 75
    frmForm.lin(16).Y1 = iBreak
    frmForm.lin(16).Y2 = iBreak
  
    frmForm.lin(17).X1 = 60
    frmForm.lin(17).X2 = frmForm.Width - 60
    frmForm.lin(17).Y1 = iBreak + 15
    frmForm.lin(17).Y2 = iBreak + 15
  
    frmForm.lin(18).X1 = 60
    frmForm.lin(18).X2 = frmForm.Width - 75
    frmForm.lin(18).Y1 = iBreak + 45
    frmForm.lin(18).Y2 = iBreak + 45
    
    frmForm.lin(19).X1 = 75
    frmForm.lin(19).X2 = frmForm.Width - 90
    frmForm.lin(19).Y1 = iBreak + 60
    frmForm.lin(19).Y2 = iBreak + 60
    
    
    frmForm.lin(20).X1 = frmForm.Width - 90
    frmForm.lin(20).X2 = frmForm.Width - 90
    frmForm.lin(20).Y1 = iBreak + 75
    frmForm.lin(20).Y2 = frmForm.Height - 90
  
    frmForm.lin(21).X1 = frmForm.Width - 75
    frmForm.lin(21).X2 = frmForm.Width - 75
    frmForm.lin(21).Y1 = iBreak + 60
    frmForm.lin(21).Y2 = frmForm.Height - 75
    
    frmForm.lin(22).X1 = 60
    frmForm.lin(22).X2 = 60
    frmForm.lin(22).Y1 = iBreak + 60
    frmForm.lin(22).Y2 = frmForm.Height - 75
  
    frmForm.lin(23).X1 = 75
    frmForm.lin(23).X2 = 75
    frmForm.lin(23).Y1 = iBreak + 75
    frmForm.lin(23).Y2 = frmForm.Height - 90
    
  End Select

End Sub

