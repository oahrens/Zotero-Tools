VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ZtMessageDisplay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Class ZtMessageDisplay.
' Shows messages and manages user reactions to these.
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Private constants.
Private Const PVT_SOUND_CRITICAL As String = "SystemHand"
Private Const PVT_SOUND_EXCLAMATION As String = "SystemExclamation"
Private Const PVT_SOUND_QUESTION As String = "SystemQuestion"
Private Const PVT_SOUND_INFORMATION As String = ".Default"
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Private variables.
Private pvtPreparable As ZtIPreparable
Private pvtTextBox As MSForms.TextBox
Private WithEvents pvtProcedeButton As MSForms.CommandButton
Private WithEvents pvtCancelButton As MSForms.CommandButton
Private WithEvents pvtDisableButton As MSForms.CommandButton
Private pvtResult As ZtFMessageType
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Friend procedures and properties.
Friend Sub Initialize(ByVal valPreparable As ZtIPreparable, ByVal valTextBox As MSForms.TextBox, ByVal valProcedeButton As MSForms.CommandButton, _
  ByVal valCancelButton As MSForms.CommandButton, ByVal valDiableButton As MSForms.CommandButton)

  Set pvtPreparable = valPreparable
  Set pvtTextBox = valTextBox
  Set pvtProcedeButton = valProcedeButton
  Set pvtCancelButton = valCancelButton
  Set pvtDisableButton = valDiableButton
  
End Sub

Friend Function Show(ByVal valMessage As String, ByVal valType As ZtFMessageType, Optional ByVal valRePrepare As Boolean = True, _
  Optional ByVal valSelectionRange As Word.Range = Nothing) As ZtFMessageType
  
  Dim locStart As Date
  
  pvtPreparable.Unprepare
  If Not valSelectionRange Is Nothing Then
    valSelectionRange.Select
    DoEvents
  End If
  With pvtTextBox
    If (valType And MessageCritical) = MessageCritical Then
      ZtApiProcedures.PlaySound PVT_SOUND_CRITICAL, vbNullString, CLng(SndAlias + SndAsync)
      .Font.Bold = True
      .ForeColor = vbRed
    ElseIf (valType And MessageExclamation) = MessageExclamation Then
      ZtApiProcedures.PlaySound PVT_SOUND_EXCLAMATION, vbNullString, CLng(SndAlias + SndAsync)
      .Font.Bold = False
      .ForeColor = vbBlue
    ElseIf (valType And MessageQuestion) = MessageQuestion Then
      ZtApiProcedures.PlaySound PVT_SOUND_QUESTION, vbNullString, CLng(SndAlias + SndAsync)
      .Font.Bold = False
      .ForeColor = vbBlack
    Else
      ZtApiProcedures.PlaySound PVT_SOUND_INFORMATION, vbNullString, CLng(SndAlias + SndAsync)
      .Font.Bold = False
      .ForeColor = vbBlack
    End If
    .Text = valMessage
  End With
  
  ' Enable buttons.
  pvtProcedeButton.Enabled = ((valType And MessageOk) = MessageOk)
  pvtCancelButton.Enabled = ((valType And MessageCancel) = MessageCancel) Or ((valType And Messageno) = Messageno)
  pvtDisableButton.Enabled = ((valType And MessageDisable) = MessageDisable)
  
  ' Wait for user reaction.
  pvtResult = MessageNone
  locStart = Now
  Do While pvtResult = MessageNone
    Do
      DoEvents
    Loop Until Now >= DateAdd("s", 1, locStart)
  Loop
  
  ' Resume after user reaction.
  If valRePrepare And pvtResult <> MessageCancel Then
    pvtPreparable.Prepare
  End If
  pvtTextBox.Text = vbNullString
  
  Show = pvtResult
  
End Function

Friend Sub Clear()
  
  pvtTextBox.Text = vbNullString
  pvtEnableProcedureControls False
  
End Sub
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Event procedures.
Private Sub pvtProcedeButton_Click()

  pvtResult = MessageOk
  pvtEnableProcedureControls False
  
End Sub

Private Sub pvtCancelButton_Click()

  pvtResult = MessageCancel
  pvtEnableProcedureControls False
  
End Sub

Private Sub pvtDisableButton_Click()

  pvtResult = MessageDisable
  pvtEnableProcedureControls False
  
End Sub
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Private procedures.
Private Sub pvtEnableProcedureControls(ByVal valEnable As Boolean)
  
  pvtProcedeButton.Enabled = valEnable
  pvtCancelButton.Enabled = valEnable
  pvtDisableButton.Enabled = valEnable
  DoEvents
  
End Sub
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

