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
Implements ZtIPreparable

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Form ZtStartForm.
' Start this form to run the macro (press F5).
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
Private pvtLicenceShown As Boolean
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
  
  Set pvtMessageDisplay = New ZtMessageDisplay
  pvtMessageDisplay.Initialize Me, InformationBox, ProcedureProcedeButton, ProcedureCancelButton, ProcedureDisableButton
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
  pvtConfig.Initialize Me
  pvtConfig.KeepUserStylePresets
  
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
  
  Set locProcedure = New ZtSetWebLinks
  pvtProcedures.Add locProcedure, locProcedure.Name
  
  Set locProcedure = New ZtSetInternalLinking
  pvtProcedures.Add locProcedure, locProcedure.Name
  
  Set locProcedure = New ZtRemoveInternalLinking
  pvtProcedures.Add locProcedure, locProcedure.Name
  
  Set locProcedure = New ZtJoinCitationGroupsSelection
  pvtProcedures.Add locProcedure, locProcedure.Name
  
  Set locProcedure = New ZtJoinCitationGroupsAll
  pvtProcedures.Add locProcedure, locProcedure.Name
  
  Set locProcedure = New ZtAdjustPunctuation
  pvtProcedures.Add locProcedure, locProcedure.Name
  
  Set locProcedure = New ZtResolveUnreachableCitations
  pvtProcedures.Add locProcedure, locProcedure.Name
  
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
  
  pvtConfig.KeepUserStyles StylePresetList.Value
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
  
  On Error GoTo OnError
  pvtEnableMacroControls False
  Set locDocument = New ZtDocument
  locDocument.Initialize pvtConfig, Me, Application.Documents.Item(DocumentList.Value)
  pvtCurrentProcedure.Start pvtConfig, Me, locDocument, pvtLicenceShown
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

Private Sub ZtIPreparable_Prepare()

  Me.Prepare
  
End Sub

Private Sub ZtIPreparable_Unprepare()

  Me.Unprepare
  
End Sub
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Friend/Public procedures and properties.
' All members that should be callable by CallByName procedure must be public.
Friend Sub Prepare()

  pvtCurrentProcedure.Prepare
  
End Sub

Friend Sub Unprepare()

  pvtCurrentProcedure.Unprepare
  
End Sub

Friend Property Get RemoveInternalLinkingProcedure() As ZtRemoveInternalLinking

  Dim locProcedure As ZtIProcedure
  
  For Each locProcedure In pvtProcedures
    If TypeOf locProcedure Is ZtRemoveInternalLinking Then
      Set RemoveInternalLinkingProcedure = locProcedure
      Exit Property
    End If
  Next
  
End Property

Friend Property Get Progress() As ZtProgress

  Set Progress = pvtProgress
  
End Property

Friend Property Get MessageDisplay() As ZtMessageDisplay

  Set MessageDisplay = pvtMessageDisplay
  
End Property

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
    If locControl.Tag = "Macro" Then
      locControl.Enabled = valEnable
    End If
  Next
  pvtWordInvisibleCheck.EnabledCacheEnabled = valEnable
  pvtCitationZeroWidthSpaceCheck.EnabledCacheEnabled = valEnable
  pvtBackwardLinkingCheck.EnabledCacheEnabled = valEnable
  pvtDebuggingCheck.EnabledCacheEnabled = valEnable
  
  DoEvents
  
End Sub
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
