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
Implements ZtIMessageDisplayable


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Class ZtMessageDisplay.
' Shows messages and manages user reactions to these.
'
' Zotero Tools.
' This software is under Revised ('New') BSD license.
' Copyright © 2019, Olaf Ahrens. All rights reserved.
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
Private pvtTextBox As MSForms.TextBox
Private WithEvents pvtProcedeButton As MSForms.CommandButton
Private WithEvents pvtCancelButton As MSForms.CommandButton
Private WithEvents pvtSuppressButton As MSForms.CommandButton
Private pvtResult As ZtFMessageType
Private pvtAppPrepare As ZtIAppPreparable
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Private interface procedures and properties directing to Friend procedures and properties below.
Private Function ZtIMessageDisplayable_Show(ByVal valMessage As String, ByVal valType As ZtFMessageType, Optional ByVal valRePrepare As Boolean = True, _
  Optional ByVal valSelectionRange As Word.Range = Nothing) As ZtFMessageType

  ZtIMessageDisplayable_Show = Me.Show(valMessage, valType, valRePrepare, valSelectionRange)
  
End Function
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Friend procedures and properties.
Friend Sub Initialize(ByVal valTextBox As MSForms.TextBox, ByVal valProcedeButton As MSForms.CommandButton, _
  Optional ByVal valCancelButton As MSForms.CommandButton = Nothing, Optional ByVal valSuppressButton As MSForms.CommandButton = Nothing, _
  Optional ByVal valAppPrepare As ZtIAppPreparable = Nothing)

  Set pvtAppPrepare = valAppPrepare
  Set pvtTextBox = valTextBox
  Set pvtProcedeButton = valProcedeButton
  Set pvtCancelButton = valCancelButton
  Set pvtSuppressButton = valSuppressButton
  
End Sub

Friend Function Show(ByVal valMessage As String, ByVal valType As ZtFMessageType, Optional ByVal valRePrepare As Boolean = True, _
  Optional ByVal valSelectionRange As Word.Range = Nothing) As ZtFMessageType
  
  Dim locStart As Date
  
  pvtAppPrepare.Unprepare
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
  If Not pvtProcedeButton Is Nothing Then
    pvtProcedeButton.Enabled = ((valType And MessageOk) = MessageOk) Or ((valType And MessageNo) = MessageNo)
  End If
  If Not pvtCancelButton Is Nothing Then
    pvtCancelButton.Enabled = ((valType And MessageCancel) = MessageCancel)
  End If
  If Not pvtSuppressButton Is Nothing Then
    pvtSuppressButton.Enabled = ((valType And MessageSuppress) = MessageSuppress)
  End If
  
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
    pvtAppPrepare.Prepare
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

Private Sub pvtSuppressButton_Click()

  pvtResult = MessageSuppress
  pvtEnableProcedureControls False
  
End Sub
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Private procedures.
Private Sub pvtEnableProcedureControls(ByVal valEnable As Boolean)
  
  If Not pvtProcedeButton Is Nothing Then
    pvtProcedeButton.Enabled = valEnable
  End If
  If Not pvtCancelButton Is Nothing Then
    pvtCancelButton.Enabled = valEnable
  End If
  If Not pvtSuppressButton Is Nothing Then
    pvtSuppressButton.Enabled = valEnable
  End If
  DoEvents
  
End Sub
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

