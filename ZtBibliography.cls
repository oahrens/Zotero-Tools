VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ZtBibliography"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Class ZtBibliography.
' It preserves properties of the bibliography inside a document and methods to edit it.
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Private variables.
Private pvtConfig As ZtConfig
Private pvtStartForm As ZtStartForm
Private pvtMainStory As ZtMainStory
Private pvtRange As Word.Range
Private pvtReferences() As ZtReference
Private pvtReferencesCt As Integer
Private pvtShowNoCitationError As Boolean
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Constructor.
Private Sub Class_Initialize()

  pvtReferencesCt = -1
  
End Sub
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Friend procedures and properties.
Friend Function Initialize(ByVal valConfig As ZtConfig, ByVal valStartForm As ZtStartForm, ByVal valMainStory As ZtMainStory) As ZtFMessageType

  Dim locField As Word.Field
  Dim locResult As ZtFMessageType

  If pvtRange Is Nothing Then
    Set pvtConfig = valConfig
    Set pvtStartForm = valStartForm
    Set pvtMainStory = valMainStory
    
    For Each locField In pvtMainStory.Document.Fields
      If InStr(locField.Code.Text, pvtConfig.Basic.Zotero.BibliographyIdentifier) > 0 Then
        Set pvtRange = locField.Result
        locResult = MessageOk
        Exit For
      End If
    Next
    If locResult <> MessageOk Then
      locResult = pvtStartForm.MessageDisplay.Show("I've not found any bibliography!" & vbNewLine & vbNewLine & _
                                                     "Are you shure that this a Zotero document?", _
                                                   MessageCritical + MessageCancel, _
                                                   False)
    End If
  Else
    locResult = MessageOk
  End If
  
  Initialize = locResult
  
End Function

Friend Function SetBookmarks() As ZtFMessageType

  Dim locReferenceCtr As Integer
  Dim locResult As ZtFMessageType
  
  locResult = pvtKeepReferences
  If locResult = MessageOk Then
    pvtStartForm.Progress.SetStep pvtReferencesCt
    For locReferenceCtr = 0 To pvtReferencesCt - 1
      pvtReferences(locReferenceCtr).SetBookmark
      pvtStartForm.Progress.SetSubstep
    Next
  End If
  
  SetBookmarks = locResult
  
End Function

Friend Function SetInternalLinking() As ZtFMessageType

  Dim locReferenceCtr As Integer
  Dim locResult As ZtFMessageType

  locResult = pvtKeepReferences
  If locResult = MessageOk Then
    pvtStartForm.Progress.SetStep pvtReferencesCt
    For locReferenceCtr = 0 To pvtReferencesCt - 1
      locResult = pvtReferences(locReferenceCtr).SetInternalLinking
      pvtStartForm.Progress.SetSubstep
      If locResult = MessageCancel Then
        Exit For
      End If
    Next
  End If
  Erase pvtReferences
  pvtReferencesCt = -1
  
  SetInternalLinking = locResult

End Function

Friend Function RemoveCitedOnPage() As ZtFMessageType

  Dim locResult As ZtFMessageType
  Dim locReferenceCtr As Integer
  
  locResult = pvtKeepReferences
  If locResult = MessageOk Then
    pvtStartForm.Progress.SetStep pvtReferencesCt
    For locReferenceCtr = 0 To pvtReferencesCt - 1
      pvtReferences(locReferenceCtr).RemoveCitedOnPage
      pvtStartForm.Progress.SetSubstep
    Next
  End If
  Erase pvtReferences
  pvtReferencesCt = -1
  
  RemoveCitedOnPage = locResult

End Function

Friend Function SetWebLinks() As ZtFMessageType

  Dim locReferenceCtr As Integer
  Dim locResult As ZtFMessageType
  Dim locAffectedReferencesCt As Integer
  
  locResult = pvtKeepReferences
  If locResult = MessageOk Then
    pvtStartForm.Progress.SetStep pvtReferencesCt
    For locReferenceCtr = 0 To pvtReferencesCt - 1
      locResult = pvtReferences(locReferenceCtr).SetWebLink(locAffectedReferencesCt)
      If locResult = MessageCancel Then
        Exit For
      End If
      pvtStartForm.Progress.SetSubstep
    Next
    If locResult = MessageOk And locAffectedReferencesCt = 0 Then
      locResult = pvtStartForm.MessageDisplay.Show("I couldn't find any web or doi addresses in the bibliography or all addresses already have been linked." & vbNewLine & vbNewLine & _
                                                     "If you're shure that unlinked addresses do exist, please check the settings for reference style in ZtConfig.xml.", _
                                                   MessageExclamation + Messageno, _
                                                   False)
    End If
  End If
  
  SetWebLinks = locResult
  
End Function

Friend Property Get ShowNoCitationError() As Boolean

  ShowNoCitationError = pvtShowNoCitationError
  
End Property

Friend Property Let ShowNoCitationError(ByVal valShow As Boolean)

  pvtShowNoCitationError = valShow
  
End Property


Friend Property Get ReferencesCt() As Integer

  ReferencesCt = pvtReferencesCt
  
End Property

Friend Property Get References(ByVal valNr As Integer) As ZtReference

  Set References = pvtReferences(valNr)
  
End Property

Friend Property Get MainStory() As ZtMainStory

  Set MainStory = pvtMainStory

End Property
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Private procedures and properties.
Private Function pvtKeepReferences() As ZtFMessageType

  Dim locRange As Word.Range
  Dim locResult As ZtFMessageType
  Dim locParagraph As Word.Paragraph

  If pvtReferencesCt = -1 Then
    pvtReferencesCt = 0
    pvtStartForm.Progress.SetStep pvtRange.Paragraphs.Count
    For Each locParagraph In pvtRange.Paragraphs
      ' The first paragraph in the paragraphs collection incomporates the whole bibliography plus next paragraph!?
      ' Otherwise the MoveEndUntil vbNewLine mechanism is very, very slow, when backreferences occure in a reference.
      If pvtReferencesCt = 0 Then
        Set locRange = pvtRange.Duplicate
        locRange.Collapse wdCollapseStart
        locRange.MoveEndUntil vbNewLine, wdForward
      Else
        Set locRange = locParagraph.Range.Duplicate
        With locRange
          .SetRange .Start, .End - 1
        End With
      End If
      locResult = pvtReferenceFactory(locRange)
      If locResult = MessageCancel Then
        Exit For
      ElseIf pvtReferences(pvtReferencesCt - 1).Nr <> pvtReferencesCt - 1 Then
        locResult = pvtStartForm.MessageDisplay.Show("There are irregularities in the sequence of your reference numbering, have a look at the selected reference.", _
                                                     MessageCritical + MessageCancel, _
                                                     False, _
                                                     locRange)
        Exit For
      Else
        locResult = MessageOk
      End If
      pvtStartForm.Progress.SetSubstep
    Next
    
    If locResult <> MessageCancel And pvtReferencesCt = 0 Then
    
      locResult = pvtStartForm.MessageDisplay.Show("No references found!" & vbNewLine & vbNewLine & _
                                                     "Are you shure this is a Zotero document? Are you also shure you've set the correct reference style in ZtConfig.xml?", _
                                                   MessageCritical + MessageCancel, _
                                                   False)
    End If
  Else
    locResult = MessageOk
  End If
  
  pvtKeepReferences = locResult

End Function

Private Function pvtReferenceFactory(ByVal valRange As Word.Range) As ZtFMessageType

  Dim locResult As ZtFMessageType
  
  ReDim Preserve pvtReferences(pvtReferencesCt)
  Set pvtReferences(pvtReferencesCt) = New ZtReference
  locResult = pvtReferences(pvtReferencesCt).Initialize(pvtConfig, pvtStartForm, Me, valRange)
  If locResult = MessageOk Then
    pvtReferencesCt = pvtReferencesCt + 1
  ElseIf pvtReferencesCt > 0 Then
    ReDim Preserve pvtReferences(pvtReferencesCt - 1)
  End If
  
  pvtReferenceFactory = locResult

End Function
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *





