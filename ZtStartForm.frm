VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ZtStartForm 
   Caption         =   "Zotero Tools (c) 2019 Olaf Ahrens"
   ClientHeight    =   9795
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9690
   OleObjectBlob   =   "ZtStartForm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "ZtStartForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Form ZtStartForm.
' Start this form to run the macro (press F5).
'
' Zotero Tools.
' This software is under Revised ('New') BSD license.
' Copyright © 2019, Olaf Ahrens. All rights reserved.
'
'    Redistribution and use in source and binary forms, with or without
'    modification, are permitted provided that the following conditions are met:
'     * Redistributions of source code must retain the above copyright
'       notice, this list of conditions and the following disclaimer.
'     * Redistributions in binary form must reproduce the above copyright
'       notice, this list of conditions and the following disclaimer in the
'       documentation and/or other materials provided with the distribution.
'     * Neither the name of the copyright holder nor the
'       names of its contributors may be used to endorse or promote products
'       derived from this software without specific prior written permission.
'
'    THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND
'    ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
'    WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
'    DISCLAIMED. IN NO EVENT SHALL <COPYRIGHT HOLDER> BE LIABLE FOR ANY
'    DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
'    (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
'    LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
'    ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
'    (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
'    SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Private constants.
Private Const PVT_VERSION As String = "1"
Private Const PVT_CAPTION As String = "Zotero Tools version " & PVT_VERSION & " - copyright © 2019 Olaf Ahrens"
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Private variables.
Private pvtConfig As ZtConfig
Private pvtProcedures As Collection
Private pvtCurrentProcedure As ZtIProcedure
Private pvtProgress As ZtProgress
Private pvtMessageDisplay As ZtMessageDisplay
Private pvtWordInvisibleCheck As ZtCheckBox
Private pvtCitationZeroWidthSpaceCheck As ZtCheckBox
Private pvtBackwardLinkingCheck As ZtCheckBox
Private pvtDebuggingCheck As ZtCheckBox
Private pvtLicenseShown As Boolean
Private pvtAppPrepare As ZtAppPreparer
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Constructor.
Private Sub UserForm_Initialize()

  Dim locProcedure As ZtIProcedure
  Dim locStylePresetObj As Object
  Dim locStylePreset As ZtConfigUserStylePreset
  
  Me.Caption = PVT_CAPTION
  
  ' Initialize outsourced objects.
  Set pvtProgress = New ZtProgress
  pvtProgress.Initialize SubstepsLabel0, SubstepsLabel1, StepsLabel, SubstepsFrame.Width, StepsFrame.Width
  pvtProgress.Reset
  
  Set pvtAppPrepare = New ZtAppPreparer
  Set pvtMessageDisplay = New ZtMessageDisplay
  pvtMessageDisplay.Initialize InformationBox, ProcedureProcedeButton, ProcedureCancelButton, ProcedureDisableButton, pvtAppPrepare
  pvtMessageDisplay.Clear
  
  Set pvtWordInvisibleCheck = New ZtCheckBox
  Set pvtCitationZeroWidthSpaceCheck = New ZtCheckBox
  Set pvtBackwardLinkingCheck = New ZtCheckBox
  Set pvtDebuggingCheck = New ZtCheckBox
  pvtWordInvisibleCheck.Initialize WordInvisibleCheck, "WordInvisible", pvtDebuggingCheck
  pvtCitationZeroWidthSpaceCheck.Initialize CitationZeroWidthSpaceCheck, "CitationZeroWidthSpace"
  pvtBackwardLinkingCheck.Initialize BackwardLinkingCheck, "BackwardLinking"
  pvtDebuggingCheck.Initialize DebuggingCheck, "Debugging", pvtWordInvisibleCheck
  
  ' Keep all style presets (sorted).
  Set pvtConfig = New ZtConfig
  With pvtConfig
    .Initialize pvtMessageDisplay
    .KeepUserStylePresets
  End With
  
  ' Insert all style presets into listbox.
  For Each locStylePresetObj In pvtConfig.UserStylePresets
    Set locStylePreset = locStylePresetObj
    StylePresetList.AddItem locStylePreset.Name
  Next
  
  ' Keep user macro presets.
  pvtWordInvisibleCheck.Value = pvtConfig.User.Macro.WordIsInvisibleWhileOperation
  pvtCitationZeroWidthSpaceCheck.Value = pvtConfig.User.Macro.CitationInsertZeroWidthSpace
  pvtBackwardLinkingCheck.Value = pvtConfig.User.Macro.WithBackwardLinking
  pvtDebuggingCheck.Value = pvtConfig.User.Macro.Debugging
  StylePresetList.Value = pvtConfig.User.Macro.StylePresetName
  
  ' Initialize all procedures.
  Set pvtProcedures = New Collection
  With pvtProcedures
    Set locProcedure = New ZtSetWebLinks
    .Add locProcedure, locProcedure.Name
    
    Set locProcedure = New ZtSetInternalLinking
    .Add locProcedure, locProcedure.Name
    
    Set locProcedure = New ZtRemoveInternalLinking
    .Add locProcedure, locProcedure.Name
    
    Set locProcedure = New ZtJoinCitationGroupsSelection
    .Add locProcedure, locProcedure.Name
    
    Set locProcedure = New ZtJoinCitationGroupsAll
    .Add locProcedure, locProcedure.Name
    
    Set locProcedure = New ZtAdjustPunctuation
    .Add locProcedure, locProcedure.Name
    
    Set locProcedure = New ZtResolveUnreachableCitations
    .Add locProcedure, locProcedure.Name
  End With
  
  ' Insert all procedure names into listbox.
  For Each locProcedure In pvtProcedures
    ProcedureList.AddItem locProcedure.Name
  Next
  ProcedureList.ListIndex = 0
  
  ' List all open documents.
  DocumentListRefreshButton_Click
  
End Sub
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Event procedures.
Private Sub DocumentListRefreshButton_Click()

  Dim locDocument As Word.Document
  
  DocumentList.Clear
  For Each locDocument In Application.Documents
    DocumentList.AddItem locDocument.Name
  Next
  DocumentList.ListIndex = 0

End Sub

Private Sub StylePresetList_Change()

  On Error GoTo OnError
  
  pvtConfig.User.KeepUserStyles StylePresetList.Value
  pvtConfig.KeepFinal
  
  Exit Sub
  
OnError:
  MsgBox "Sorry, an error has happened: " & vbNewLine & _
           Err.Description & vbNewLine & _
           Err.Number & vbNewLine & _
           "in: " & Err.Source, _
         vbOKOnly + vbExclamation, _
         "Zotero Tools"
  MacroEndButton_Click
  
End Sub

Private Sub ProcedureList_Change()

  Dim locIsSetInternalLinking As Boolean
  Dim locIsJoinCitationGroupsSelection As Boolean

  Set pvtCurrentProcedure = pvtProcedures.Item(ProcedureList.Value)
  
  If TypeOf pvtCurrentProcedure Is ZtSetInternalLinking Then
    locIsSetInternalLinking = True
  ElseIf TypeOf pvtCurrentProcedure Is ZtJoinCitationGroupsSelection Then
    locIsJoinCitationGroupsSelection = True
  End If
  
  pvtCitationZeroWidthSpaceCheck.EnabledCacheValue = locIsSetInternalLinking
  pvtBackwardLinkingCheck.EnabledCacheValue = locIsSetInternalLinking
  
  ' From https://www.mrexcel.com/forum/excel-questions/408356-how-can-i-change-modal-userform-modal-modeless-run-time.html.
  ZtApiProcedures.EnableWindow ZtApiProcedures.GetParent(ZtApiProcedures.FindWindow(vbNullString, Me.Caption)), IIf(locIsJoinCitationGroupsSelection, CLng(CTrue), CLng(CFalse))
  SelectRangeLabel.Visible = locIsJoinCitationGroupsSelection

  ProcedureDescriptionBox.Text = pvtCurrentProcedure.Description

End Sub

Private Sub ProcedureStartButton_Click()

  Dim locDocument As ZtDocument
  Dim locSetInternalLinkingProcedure As ZtSetInternalLinking
  Dim locProcedureInitializer As ZtProcedureInitializer
  
  On Error GoTo OnError
  pvtEnableMacroControls False
  Set locDocument = New ZtDocument
  locDocument.Initialize pvtConfig, Application.Documents.Item(DocumentList.Value), pvtMessageDisplay, pvtProgress
  pvtAppPrepare.Initialize pvtConfig, locDocument
  Set locProcedureInitializer = New ZtProcedureInitializer
  If TypeOf pvtCurrentProcedure Is ZtSetInternalLinking Then
    Set locSetInternalLinkingProcedure = pvtCurrentProcedure
    Set locSetInternalLinkingProcedure.RemoveInternaleLinkingProcedure = pvtRemoveInternalLinkingProcedure
  End If
  pvtCurrentProcedure.Start pvtConfig, pvtMessageDisplay, pvtProgress, pvtAppPrepare, locProcedureInitializer, locDocument, pvtLicenseShown
  pvtProgress.Reset
  pvtEnableMacroControls True
  
  Exit Sub

OnError:
  MsgBox "Sorry, an error has happened: " & vbNewLine & _
           Err.Description & vbNewLine & _
           Err.Number & vbNewLine & _
           "in: " & Err.Source, _
         vbOKOnly + vbExclamation, _
         "Zotero Tools"
  MacroEndButton_Click

End Sub

Private Sub MacroEndButton_Click()

  End

End Sub

Private Sub UserForm_Terminate()

  MacroEndButton_Click
  
End Sub
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Friend/Public procedures and properties.
' All members that should be callable by CallByName procedure must be public.
Public Property Let WordInvisible(ByVal valInvisible As Boolean)

  pvtConfig.User.Macro.WordIsInvisibleWhileOperation = valInvisible

End Property

Public Property Get WordInvisible() As Boolean

  WordInvisible = pvtConfig.User.Macro.WordIsInvisibleWhileOperation

End Property

Public Property Let CitationZeroWidthSpace(ByVal valWithSpace As Boolean)

  pvtConfig.User.Macro.CitationInsertZeroWidthSpace = valWithSpace
  
End Property

Public Property Get CitationZeroWidthSpace() As Boolean

  CitationZeroWidthSpace = pvtConfig.User.Macro.CitationInsertZeroWidthSpace
  
End Property

Public Property Let BackwardLinking(ByVal valWithLinking As Boolean)

  pvtConfig.User.Macro.WithBackwardLinking = valWithLinking
  
End Property

Public Property Get BackwardLinking() As Boolean

  BackwardLinking = pvtConfig.User.Macro.WithBackwardLinking
  
End Property

Public Property Let Debugging(ByVal valDebugging As Boolean)

  pvtConfig.User.Macro.Debugging = valDebugging
  
End Property

Public Property Get Debugging() As Boolean

  Debugging = pvtConfig.User.Macro.Debugging
  
End Property
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Private procedures and properties.
Private Sub pvtEnableMacroControls(ByVal valEnable As Boolean)

  Dim locControl As MSForms.Control
  
  For Each locControl In Me.Controls
    With locControl
      If .Tag = "Macro" Then
        .Enabled = valEnable
      End If
    End With
  Next
  pvtWordInvisibleCheck.EnabledCacheEnabled = valEnable
  pvtCitationZeroWidthSpaceCheck.EnabledCacheEnabled = valEnable
  pvtBackwardLinkingCheck.EnabledCacheEnabled = valEnable
  pvtDebuggingCheck.EnabledCacheEnabled = valEnable
  
  DoEvents
  
End Sub

Private Property Get pvtRemoveInternalLinkingProcedure() As ZtRemoveInternalLinking

  Dim locProcedure As ZtIProcedure
  
  For Each locProcedure In pvtProcedures
    If TypeOf locProcedure Is ZtRemoveInternalLinking Then
      Set pvtRemoveInternalLinkingProcedure = locProcedure
      Exit Property
    End If
  Next
  
End Property
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
