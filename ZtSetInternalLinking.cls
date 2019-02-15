VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ZtSetInternalLinking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements ZtIProcedure

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Class ZtSetInternalLinking.
' This is the 'Set internal linking between citations and references' procedure.
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Private variables.
Private pvtName As String
Private pvtDescription As String
Private pvtFunctions As ZtProcedureFunctions
Private pvtStepsCt As Integer
Private pvtDocument As ZtDocument
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Constructor.
Private Sub Class_Initialize()

  pvtName = "Set internal linking between citations and references"
  pvtDescription = "This procedure sets internal links between citations and bibliography." & vbNewLine & vbNewLine & _
                     "This procedure is only working for numeric style citations. " & _
                       "By switch 'user.macro.withBackwardLinking' in ZtConfig.xml you can set whether this linking is " & _
                       "unidirectional (citation -> reference) or bidirectional (citation <-> reference)." & vbNewLine & vbNewLine & _
                     "Thanks pavelninin for different ideas from https://github.com/pavelninin/zotero_two_way_hyperlinking/blob/master/zotero_two_way_hyperlinking.vbs."
End Sub
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Private interface procedures and properties directing to Friend procedures and properties below.
Private Property Get ZtIProcedure_Name() As String

  ZtIProcedure_Name = Me.Name
  
End Property

Private Property Get ZtIProcedure_Description() As String

  ZtIProcedure_Description = Me.Description
  
End Property

Private Property Get ZtIProcedure_ProcedureFunctions() As ZtProcedureFunctions

  Set ZtIProcedure_ProcedureFunctions = Me.ProcedureFunctions
  
End Property

Private Function ZtIProcedure_Start(ByVal valConfig As ZtConfig, ByVal valStartForm As ZtStartForm, ByVal valDocument As ZtDocument, ByRef refLicenceShown As Boolean) As ZtFMessageType

  ZtIProcedure_Start = Me.Start(valConfig, valStartForm, valDocument, refLicenceShown)
  
End Function

Private Sub ZtIProcedure_Prepare()

  Me.Prepare
  
End Sub

Private Sub ZtIProcedure_Unprepare()

  Me.Unprepare
  
End Sub
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *



' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Friend procedures and properties.
Friend Property Get Name() As String

  Name = pvtName
  
End Property

Friend Property Get Description() As String

  Description = pvtDescription
  
End Property

Friend Property Get ProcedureFunctions() As ZtProcedureFunctions

  Set ProcedureFunctions = pvtFunctions
  
End Property

Friend Function Start(ByVal valConfig As ZtConfig, ByVal valStartForm As ZtStartForm, ByVal valDocument As ZtDocument, ByRef refLicenceShown As Boolean) As ZtFMessageType

  Dim locResult As ZtFMessageType
  
  Set pvtDocument = valDocument
  With valStartForm.RemoveInternalLinkingProcedure
    Set pvtFunctions = .ProcedureFunctions
    .ToResume = True
    locResult = .Start(valConfig, valStartForm, pvtDocument, refLicenceShown)
    .ToResume = False
  End With
  
  If locResult = MessageOk Then
    valStartForm.Progress.Reset
    locResult = valStartForm.MessageDisplay.Show("Any previous internal linking has been removed and in the next step new one will be done." & vbNewLine & vbNewLine & _
                                                   "Your citations are up to date?" & vbNewLine & vbNewLine & _
                                                   "You may procede, otherwise cancel and run the Zotero 'Refresh' command before starting this procedure again.", _
                                                 MessageQuestion + MessageOkCancel)
    If locResult = MessageOk Then
      If Len(valConfig.User.CitationGroupStyle.Prefix) = 0 And Len(valConfig.User.CitationStyle.Prefix) = 0 And Len(valConfig.User.CitationStyle.Suffix) = 0 Then
        If Not valConfig.User.Macro.CitationInsertZeroWidthSpace Then
          locResult = valStartForm.MessageDisplay.Show("You should perhaps set 'user.macro.citationInsertZeroWidthSpace' in ZtConfig.xml to 'True', to let all internal hyperlinks work correctly. " & _
                                                         "For further explanations, please have a look at the configuration file." & vbNewLine & vbNewLine & _
                                                         "You may procede, otherwise cancel.", _
                                                       MessageQuestion + MessageOkCancel)
        Else
          locResult = valStartForm.MessageDisplay.Show("You have set 'user.macro.citationInsertZeroWidthSpace' in ZtConfig.xml to 'True', to let all hyperlinks work correctly. " & _
                                                         "Please, note that Zotero will ask you on some citations, whether it should discard them for " & _
                                                         "retaining some changes on them. ALWAYS ANSWER THIS QUESTION WITH 'NO'!" & vbNewLine & _
                                                         "To prevent these questions you can do one of the following:" & vbNewLine & _
                                                         vbTab & "- run Word's undo command before Zotero's refresh" & vbNewLine & _
                                                         vbTab & "- run '" & valStartForm.RemoveInternalLinkingProcedure.Name & "' procedure before Zotero's refresh" & vbNewLine & vbNewLine & _
                                                         "You may procede, otherwise cancel.", _
                                                       MessageQuestion + MessageOkCancel)
        End If
      End If
      
      If locResult = MessageOk Then
        valStartForm.Progress.SetStepsCt pvtGetStepsCt
        pvtDocument.Repaginate
        locResult = pvtFunctions.Initialize(valConfig, valStartForm, pvtDocument, pvtName, refLicenceShown, True, True)
        If locResult = MessageOk Then
          locResult = pvtDocument.SetInternalLinking
          valStartForm.Progress.SetCompleted
        End If
      End If
    End If
  End If
  pvtFunctions.Finish locResult, _
                      IIf(locResult = MessageOk, _
                          "You can also use the '" & valStartForm.RemoveInternalLinkingProcedure.Name & "' procedure for this purpose.", _
                          vbNullString)
  
  Start = locResult
  
End Function

Friend Sub Prepare()

  pvtFunctions.Prepare
  
End Sub

Friend Sub Unprepare()

  pvtFunctions.Unprepare
  
End Sub
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Private procedures and properties.
Private Function pvtGetStepsCt() As Integer

  If pvtStepsCt = 0 Then
    pvtStepsCt = pvtDocument.GetSetInternalLinkingStepsCt
  End If
  
  pvtGetStepsCt = pvtStepsCt
  
End Function
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

