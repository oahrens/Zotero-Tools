VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ZtProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Class ZtProgress.
' It manages progress bars in ZtStartForm.
'
' Zotero Tools.
' This software is under Revised ('New') BSD license.
' Copyright © 2019, Olaf Ahrens. All rights reserved.
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Private constants.
Private Const PVT_SUBSTEPS_RUNNING_WIDTH_DIVISOR As Integer = 10
Private Const PVT_SUBSTEPS_RUNNING_SHIFT_DIVISOR As Integer = 7
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Private variables.
Private pvtSub0Control As MSForms.Control
Private pvtSub1Control As MSForms.Control
Private pvtControl As MSForms.Control
Private pvtSubMax As Single
Private pvtMax As Single
Private pvtSubProgressWidth As Single
Private pvtProgressWidth As Single
Private pvtSub0CurrentWidth As Single
Private pvtSub1CurrentWidth As Single
Private pvtCurrentWidth As Single
Private pvtSubRunningWidth As Single
Private pvtSubRunningShift As Single
Private pvtSub0Left As Single
Private pvtIsInRunningMode As Boolean
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Friend procedures and properties.
Friend Sub Initialize(ByVal valSubstepsControl0 As MSForms.Control, ByVal valSubstepsControl1 As MSForms.Control, ByVal valStepsControl As MSForms.Control, _
  ByVal valSubstepsMaxWidth As Single, ByVal valStepsMaxWidth As Single)
  
  Set pvtSub0Control = valSubstepsControl0
  Set pvtSub1Control = valSubstepsControl1
  Set pvtControl = valStepsControl
  pvtSubMax = valSubstepsMaxWidth
  pvtMax = valStepsMaxWidth
  pvtSubRunningWidth = pvtSubMax / PVT_SUBSTEPS_RUNNING_WIDTH_DIVISOR
  pvtSubRunningShift = pvtSubRunningWidth / PVT_SUBSTEPS_RUNNING_SHIFT_DIVISOR
  Me.Reset

End Sub

Friend Sub Reset()

  pvtSubProgressWidth = 0
  pvtProgressWidth = 0
  pvtSub0CurrentWidth = 0
  pvtSub1CurrentWidth = 0
  pvtCurrentWidth = 0
  pvtSub0Left = 0
  With pvtSub0Control
    .Left = pvtSub0Left
    .Width = pvtSub0CurrentWidth
  End With
  With pvtSub1Control
    .Left = 0
    .Width = pvtSub1CurrentWidth
  End With
  With pvtControl
    .Left = 0
    .Width = pvtCurrentWidth
  End With
  pvtIsInRunningMode = False
  DoEvents

End Sub

Friend Sub SetCompleted()

  pvtIsInRunningMode = False
  pvtSubProgressWidth = pvtSubMax
  pvtSub0Control.Width = pvtSubProgressWidth
  pvtSub0Left = 0
  pvtSub0Control.Left = pvtSub0Left
  pvtSub1CurrentWidth = 0
  pvtSub1Control.Width = pvtSub1CurrentWidth
  pvtCurrentWidth = pvtMax
  pvtControl.Width = pvtCurrentWidth
  
End Sub

Friend Sub SetStepsCt(ByVal valCt As Integer)

  pvtProgressWidth = pvtMax / CSng(valCt + 2) ' For initial step and last step.
  Me.SetCompleteStep

End Sub

Friend Sub SetStep(Optional ByVal valSubstepsCt As Integer = 0)

  If pvtProgressWidth > 0 Then
    If pvtIsInRunningMode Then
      If valSubstepsCt > 0 Then
        Me.SetToNormalMode
      End If
    Else
      If valSubstepsCt = 0 Then
        Me.SetToRunningMode
      End If
    End If
    pvtCurrentWidth = pvtCurrentWidth + pvtProgressWidth
    pvtControl.Width = pvtCurrentWidth
    If Not pvtIsInRunningMode Then
      pvtSubProgressWidth = pvtSubMax / valSubstepsCt
      pvtSub0CurrentWidth = 0
      pvtSub0Control.Width = pvtSub0CurrentWidth
      pvtSub1CurrentWidth = 0
      pvtSub1Control.Width = pvtSub1CurrentWidth
    End If
    DoEvents
  End If

End Sub

Friend Sub SetCompleteStep()

  Me.SetStep 1
  Me.SetSubstep
  
End Sub

Friend Sub SetSubstep()

  If pvtProgressWidth > 0 Then
    If pvtIsInRunningMode Then
      pvtSub0Left = (pvtSub0Left + pvtSubRunningShift) Mod pvtSubMax
      pvtSub0Control.Left = pvtSub0Left
      If pvtSub0Left + pvtSubRunningWidth > pvtSubMax Then
        pvtSub1Control.Width = pvtSub0Left + pvtSubRunningWidth - pvtSubMax
      Else
        pvtSub1Control.Width = 0
      End If
    Else
      pvtSub0CurrentWidth = pvtSub0CurrentWidth + pvtSubProgressWidth
      pvtSub0Control.Width = pvtSub0CurrentWidth
    End If
    DoEvents
  End If

End Sub

Friend Sub SetToRunningMode()

  If Not pvtIsInRunningMode Then
    pvtIsInRunningMode = True
    pvtSub0Control.Width = pvtSubRunningWidth
    DoEvents
  End If
  
End Sub

Friend Sub SetToNormalMode()
  
  If pvtIsInRunningMode Then
    pvtIsInRunningMode = False
    pvtSub0Left = 0
    pvtSub0Control.Left = pvtSub0Left
    pvtSub0Control.Width = pvtSub0CurrentWidth
    pvtSub1CurrentWidth = 0
    pvtSub1Control.Width = pvtSub1CurrentWidth
    DoEvents
  End If
  
End Sub
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

