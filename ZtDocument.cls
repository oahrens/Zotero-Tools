VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ZtDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Class ZtDocument.
' It preserves properties of the document and methods to edit it.
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Private variables.
Private pvtConfig As ZtConfig
Private pvtStartForm As ZtStartForm
Private pvtDocument As Word.Document
Private pvtStories() As ZtIStory
Private pvtStoriesCt As Integer
Private pvtShowFieldCode As Boolean
Private pvtShowAll As Boolean
Private pvtShowHiddenText As Boolean
Private pvtJunk As Long
Private pvtCitationGroupsCt As Integer
Private pvtIsPrepared As Boolean
Private pvtShowManualAdjustation As Boolean
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Constructor.
Private Sub Class_Initialize()

  pvtStoriesCt = -1
  pvtCitationGroupsCt = -1
  pvtShowManualAdjustation = True
  
End Sub
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Friend procedures and variables.
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Intenal functionality.
Friend Sub Initialize(ByVal valConfig As ZtConfig, ByVal valStartForm As ZtStartForm, ByVal valDocument As Word.Document)
  
  Set pvtConfig = valConfig
  Set pvtStartForm = valStartForm
  Set pvtDocument = valDocument
  
  ' Fix the skipped blank header/footer problem.
  ' From: https://wordmvp.com/FAQs/Customization/ReplaceAnywhere.htm.
  pvtJunk = pvtDocument.Sections(1).Headers(1).Range.StoryType
  
End Sub

Friend Sub SaveBackupCopy()

  Dim locFileNameCache As String

  locFileNameCache = pvtDocument.FullName
  pvtDocument.SaveAs ZtSubprocedures.GetPath(pvtDocument.FullName) & _
                     Application.PathSeparator & _
                     ZtSubprocedures.GetFileNameWithoutExtension(pvtDocument.Name) & "_ZtBackup_" & Format(Now, "YYYYMMDD_HHmmss") & ZtSubprocedures.GetExtension(pvtDocument.Name)
  DoEvents
  pvtDocument.SaveAs locFileNameCache
  DoEvents
  
End Sub

Friend Sub Prepare()
  
  If Not pvtIsPrepared Then
    With pvtDocument.ActiveWindow.View
      pvtShowFieldCode = .ShowFieldCodes
      .ShowFieldCodes = False
      pvtShowAll = .ShowAll
      .ShowAll = False
      pvtShowHiddenText = .ShowHiddenText
      .ShowHiddenText = False
    End With
    pvtIsPrepared = True
  End If

End Sub

Friend Sub Unprepare()
  
  If pvtIsPrepared Then
    With pvtDocument.ActiveWindow.View
      .ShowFieldCodes = pvtShowFieldCode
      .ShowAll = pvtShowAll
      .ShowHiddenText = pvtShowHiddenText
    End With
    pvtIsPrepared = False
  End If

End Sub

Friend Function GetAdjustPunctuationStepsCt() As Integer
  
  Dim locResult As Integer
  Dim locStoryCtr As Integer
  Dim locStory As ZtIStory
  
  locResult = 2 ' For pvtKeepStories and pvtKeepCitationGroups.
  pvtKeepStories
  For locStoryCtr = 0 To pvtStoriesCt - 1
    Set locStory = pvtStories(locStoryCtr)
    If locStory.CanHoldCitations Then
      locResult = locResult + 1
    End If
  Next
  
  GetAdjustPunctuationStepsCt = locResult
  
End Function

Friend Function GetJoinCitationGroupsAllStepsCt() As Integer
  
  Dim locResult As Integer
  Dim locStoryCtr As Integer
  Dim locStory As ZtIStory
  
  locResult = 2 ' For pvtKeepStories and pvtKeepCitationGroups.
  pvtKeepStories
  For locStoryCtr = 0 To pvtStoriesCt - 1
    Set locStory = pvtStories(locStoryCtr)
    If locStory.CanHoldCitations Then
      locResult = locResult + 1
    End If
  Next
  
  GetJoinCitationGroupsAllStepsCt = locResult
  
End Function

Friend Function GetJoinCitationGroupsSelectionStepsCt() As Integer
  
  Dim locResult As Integer
  
  locResult = 3 ' For pvtKeepCitationGroups, keep citation groups of one story, and one story.
  
  GetJoinCitationGroupsSelectionStepsCt = locResult
  
End Function

Friend Function GetRemoveInternalLinkingStepsCt() As Integer
  
  Dim locResult As Integer
  Dim locStoryCtr As Integer
  Dim locStory As ZtIStory
  
  locResult = 1 ' For pvtKeepStories.
  pvtKeepStories
  For locStoryCtr = 0 To pvtStoriesCt - 1
    Set locStory = pvtStories(locStoryCtr)
    If locStory.CanHoldCitations Then
      locResult = locResult + 2 ' For hyperlinks and bookmarks.
    End If
  Next
  locResult = locResult + 2 ' For pvtKeepReferences and RemoveCitedOnPage in bibliography.
  
  GetRemoveInternalLinkingStepsCt = locResult
  
End Function

Friend Function GetSetInternalLinkingStepsCt() As Integer
  
  Dim locResult As Integer
  Dim locStoryCtr As Integer
  Dim locStory As ZtIStory
  
  locResult = 5 ' For pvtKeepStories, pvtKeepCitationGroups, SetBibliographyBookmarks, pvtKeepReferences, SetBookmarks.
  If pvtConfig.User.Macro.WithBackwardLinking Then
    locResult = locResult + 2 ' For KeepCaptions and SetBibliographyLinking.
  End If
  pvtKeepStories
  For locStoryCtr = 0 To pvtStoriesCt - 1
    Set locStory = pvtStories(locStoryCtr)
    If locStory.CanHoldCitations Then
      locResult = locResult + 1
    End If
  Next
  locResult = locResult + 1 ' For pvtUpdateAllFields
  
  GetSetInternalLinkingStepsCt = locResult
  
End Function

Friend Function GetResolveUnreachableCitationsStepsCt() As Integer
  
  Dim locResult As Integer
  Dim locStoryCtr As Integer
  Dim locStory As ZtIStory
  
  locResult = 2 ' For pvtKeepReferences, pvtKeepCitationGroups.
  pvtKeepStories
  For locStoryCtr = 0 To pvtStoriesCt - 1
    Set locStory = pvtStories(locStoryCtr)
    If Not locStory.CanHoldCitations Then
      locResult = locResult + 1
    End If
  Next
  
  GetResolveUnreachableCitationsStepsCt = locResult
  
End Function

Friend Function GetSetWebLinksStepsCt() As Integer
  
  Dim locResult As Integer
  
  locResult = 3 ' For pvtKeepCitationGroups, pvtKeepReferences, and SetWebLinks.
  
  GetSetWebLinksStepsCt = locResult
  
End Function

Friend Function AdjustPunctuation() As ZtFMessageType

  Dim locStoryCtr As Integer
  Dim locStory As ZtIStory
  Dim locResult As ZtFMessageType
  Dim locAffectedCt As Integer
  Dim locNotAffectedCt As Integer
  Dim locNotCorrectedCt As Integer
  
  Err.Raise 1234, " Test"
  
  pvtKeepStories
  locResult = pvtKeepCitationGroups
  If locResult = MessageOk Then
    For locStoryCtr = 0 To pvtStoriesCt - 1
      Set locStory = pvtStories(locStoryCtr)
      If locStory.CanHoldCitations Then
        locResult = locStory.AdjustPunctuation(locAffectedCt, locNotAffectedCt, locNotCorrectedCt)
        If locResult = MessageCancel Then
          Exit For
        End If
      End If
    Next
    If locResult = MessageOk Then
      pvtStartForm.Progress.SetCompleted
      locResult = pvtCheckAffectedCitationGroupsCt(locAffectedCt, locNotAffectedCt, locNotCorrectedCt, "punctation adjustment of")
    End If
  End If
  
  AdjustPunctuation = locResult

End Function

Friend Function RemoveInternalLinking() As ZtFMessageType

  Dim locStoryCtr As Integer
  Dim locStory As ZtIStory
  Dim locResult As ZtFMessageType
  
  pvtKeepStories
  ' delete hyperlinks to references and citations, bookmarks of references and citations, and hidden spaces
  For locStoryCtr = 0 To pvtStoriesCt - 1
    Set locStory = pvtStories(locStoryCtr)
    If locStory.CanHoldCitations Then
      locResult = locStory.RemoveInternalLinking()
      If locResult = MessageCancel Then
        Exit For
      End If
    End If
  Next
  pvtStartForm.Progress.SetCompleted
  
  RemoveInternalLinking = locResult

End Function

Friend Function SetInternalLinking() As ZtFMessageType

  Dim locStoryCtr As Integer
  Dim locStory As ZtIStory
  Dim locResult As ZtFMessageType
  
  pvtKeepStories
  locResult = pvtKeepCitationGroups
  If locResult = MessageOk Then
    ' set bookmarks on references
    locResult = MainStory.SetBibliographyBookmarks
    If locResult = MessageOk Then
    
      ' Pick captions.
      If pvtConfig.User.Macro.WithBackwardLinking Then
        MainStory.KeepCaptions
      End If
        
      ' link citations to references
      For locStoryCtr = 0 To pvtStoriesCt - 1
        Set locStory = pvtStories(locStoryCtr)
        If locStory.CanHoldCitations Then
          locResult = locStory.SetInternalLinking
          If locResult = MessageCancel Then
            Exit For
          End If
        End If
      Next
      
      ' if necessary link reference to citation (backward linking)
      If locResult = MessageOk And pvtConfig.User.Macro.WithBackwardLinking Then
        locResult = MainStory.SetBibliographyLinking
      End If
      pvtUpdateAllFields
    End If
  End If
  
  SetInternalLinking = locResult

End Function

Friend Function JoinCitationGroupsAll() As ZtFMessageType

  Dim locStory As ZtIStory
  Dim locStoryCtr As Integer
  Dim locResult As ZtFMessageType
  Dim locAffectedCt As Integer
  Dim locNotAffectedCt As Integer
  
  pvtKeepStories
  locResult = pvtKeepCitationGroups
  If locResult = MessageOk Then
    pvtStartForm.Progress.SetToRunningMode
    For locStoryCtr = 0 To pvtStoriesCt - 1
      Set locStory = pvtStories(locStoryCtr)
      If locStory.CanHoldCitations Then
        locStory.JoinCitationGroups locAffectedCt, locNotAffectedCt
        pvtStartForm.Progress.SetSubstep
      End If
    Next
    pvtStartForm.Progress.SetCompleted
    pvtCitationGroupsCt = -1
    locResult = pvtCheckAffectedCitationGroupsCt(locAffectedCt, locNotAffectedCt, 0, "joining of")
  End If
  
  JoinCitationGroupsAll = locResult

End Function

Friend Function JoinCitationGroupsSelection() As ZtFMessageType

  Dim locStory As ZtIStory
  Dim locStoryCtr As Integer
  Dim locAffectedCt As Integer
  Dim locResult As ZtFMessageType
  Dim locRange As Word.Range
  
  Set locRange = Selection.Range
  pvtKeepStories
  For locStoryCtr = 0 To pvtStoriesCt - 1
    Set locStory = pvtStories(locStoryCtr)
    If locRange.InRange(locStory.Range) Then
      locStory.KeepCitationGroups pvtCitationGroupsCt
      pvtStartForm.Progress.SetStep
      pvtStartForm.Progress.SetToRunningMode
      locStory.JoinCitationGroups locAffectedCt, 0, locRange
      Exit For
    End If
  Next
  pvtStartForm.Progress.SetCompleted
  pvtCitationGroupsCt = -1
  locResult = pvtCheckAffectedCitationGroupsCt(locAffectedCt, 0, 0, "joining of")
  
  JoinCitationGroupsSelection = locResult

End Function

Friend Function SetWebLinks() As ZtFMessageType

  Dim locResult As ZtFMessageType
  
  pvtKeepStories
  locResult = MainStory.SetBibliographyWebLinks
  
  SetWebLinks = locResult
  
End Function

Friend Function ResolveUnreachableCitationGroups() As ZtFMessageType

  Dim locStoryCtr As Integer
  Dim locStory As ZtIStory
  Dim locAffectedCt As Integer
  Dim locResult As ZtFMessageType
  
  pvtKeepStories
  locResult = pvtKeepCitationGroups
  If locResult = MessageOk Then
    For locStoryCtr = 0 To pvtStoriesCt - 1
      Set locStory = pvtStories(locStoryCtr)
      If Not locStory.CanHoldCitations Then
        locStory.ResolveCitationGroups locAffectedCt
      End If
    Next
    locResult = pvtCheckAffectedCitationGroupsCt(locAffectedCt, 0, 0, "resolving unreachable")
  End If
  pvtCitationGroupsCt = -1
  
  ResolveUnreachableCitationGroups = locResult
  
End Function

Friend Function MainStory() As ZtMainStory

  Set MainStory = pvtStories(0)
  
End Function

Friend Property Get CitationGroupsCt() As Integer

  CitationGroupsCt = pvtCitationGroupsCt
  
End Property

Friend Property Get ShowManualAdjustation() As Boolean

  ShowManualAdjustation = pvtShowManualAdjustation
  
End Property

Friend Property Let ShowManualAdjustation(ByVal valShow As Boolean)

  pvtShowManualAdjustation = valShow
  
End Property
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Word functionality.
Friend Property Get Name() As String

  Name = pvtDocument.Name
  
End Property

Friend Property Get Styles(ByVal valStyle As Variant) As Word.Style

  Set Styles = pvtDocument.Styles(valStyle)
  
End Property

Friend Property Get Fields() As Word.Fields

  Set Fields = pvtDocument.Fields
  
End Property

Friend Property Get Footnotes() As Word.Footnotes

  Set Footnotes = pvtDocument.Footnotes
  
End Property

Friend Property Get Endnotes() As Word.Endnotes

  Set Endnotes = pvtDocument.Endnotes
  
End Property

Friend Property Get Comments() As Word.Comments

  Set Comments = pvtDocument.Comments
  
End Property

Friend Sub Repaginate()

  pvtDocument.Repaginate
  
End Sub

Friend Function AddBookmark(ByVal valName As String, ByVal valRange As Word.Range) As Word.Bookmark

  Set AddBookmark = pvtDocument.Bookmarks.Add(valName, valRange)
  
End Function

Friend Function AddInternalLink(ByVal valAnchor As Word.Range, ByVal valSubAddress As String, ByVal valScreenTip As String) As Word.Hyperlink

  Set AddInternalLink = pvtDocument.Hyperlinks.Add(Anchor:=valAnchor, SubAddress:=valSubAddress, ScreenTip:=valScreenTip)
  
End Function

Friend Function AddWebLink(ByVal valAnchor As Word.Range, ByVal valAddress As String) As Word.Hyperlink

  Set AddWebLink = pvtDocument.Hyperlinks.Add(Anchor:=valAnchor, Address:=valAddress)
  
End Function

Friend Function GoingTo(Optional ByVal valWhat As Variant, Optional ByVal valWhich As Variant, Optional ByVal valCount As Variant, _
  Optional ByVal valName As Variant) As Word.Range

  Set GoingTo = pvtDocument.GoTo(valWhat, valWhich, valCount, valName)
  
End Function

Friend Property Get TablesOfContents() As Word.TablesOfContents

  Set TablesOfContents = pvtDocument.TablesOfContents
  
End Property

Friend Property Get TablesOfFigures() As Word.TablesOfFigures

  Set TablesOfFigures = pvtDocument.TablesOfFigures
  
End Property
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Private procedures and variables.

' thanks to https://superuser.com/questions/196703/how-do-i-update-all-fields-in-a-word-document
Private Sub pvtUpdateAllFields()

  Dim locStoryCtr As Integer
  
  '' Update fields everywhere. This includes updates of page numbers in
  '' tables (but would not add or remove entries). This also takes care of
  '' all index updates.
  pvtStartForm.Progress.SetStep pvtStoriesCt
  For locStoryCtr = 0 To pvtStoriesCt - 1
    pvtStories(locStoryCtr).UpdateAllFields
    pvtStartForm.Progress.SetSubstep
  Next
    
End Sub

Private Sub pvtKeepStories()

  Dim locStoryRange As Word.Range
  Dim locShapeRange As Word.ShapeRange

  If pvtStoriesCt = -1 Then
    pvtStoriesCt = 0
    pvtStartForm.Progress.SetStep pvtDocument.StoryRanges.Count
    For Each locStoryRange In pvtDocument.StoryRanges
      If Not ZtSubprocedures.ArrayContains(locStoryRange.StoryType, pvtConfig.Basic.Word.CantHoldFieldStoryTypes) Then
        ' Word can't insert fields into 'normal' text frames (called 'textboxes', also we are in a text frame story), so we can skip this story.
        ' Word can insert fields into frames stored in form fields. But these are part of and queued in the Word main story.
        ' Further informations about frames and textboxes: http://www.addbalance.com/word/frames_textboxes.htm.
        Do
          pvtStoryFactory locStoryRange
          
          ' From: https://wordmvp.com/FAQs/Customization/ReplaceAnywhere.htm.
          If ZtSubprocedures.ArrayContains(locStoryRange.StoryType, pvtConfig.Basic.Word.MultipleStoriesStoryTypes) Then
            If locStoryRange.ShapeRange.Count > 0 Then
              For Each locShapeRange In locStoryRange.ShapeRange
                If locShapeRange.TextFrame.HasText Then
                  pvtStoryFactory locShapeRange.TextFrame.TextRange
                End If
              Next
            End If
            
          End If
          
          Set locStoryRange = locStoryRange.NextStoryRange
          
        Loop Until locStoryRange Is Nothing
      End If
      pvtStartForm.Progress.SetSubstep
    Next
  Else
    pvtStartForm.Progress.SetCompleteStep
  End If
  
End Sub

Private Function pvtKeepCitationGroups() As ZtFMessageType

  Dim locResult As ZtFMessageType
  Dim locStoryCtr As Integer
  
  If pvtCitationGroupsCt = -1 Then
    pvtCitationGroupsCt = 0
    pvtStartForm.Progress.SetStep pvtStoriesCt
    For locStoryCtr = 0 To pvtStoriesCt - 1
      pvtStories(locStoryCtr).KeepCitationGroups pvtCitationGroupsCt
      pvtStartForm.Progress.SetSubstep
    Next
    
    If pvtCitationGroupsCt = 0 Then
      locResult = pvtStartForm.MessageDisplay.Show("No citation groups found!" & vbNewLine & vbNewLine & _
                                                     "Are you shure this is a Zotero document?", _
                                                   MessageCritical + MessageCancel, _
                                                   False)
    Else
      locResult = MessageOk
    End If
  Else
    locResult = MessageOk
  End If
  
  pvtKeepCitationGroups = locResult
  
End Function

Private Sub pvtStoryFactory(ByVal valStoryRange As Word.Range)

  ReDim Preserve pvtStories(pvtStoriesCt)
  If valStoryRange.StoryType = wdMainTextStory Then
    Set pvtStories(pvtStoriesCt) = New ZtMainStory
  Else
    Set pvtStories(pvtStoriesCt) = New ZtStory
  End If
  pvtStories(pvtStoriesCt).Initialize pvtConfig, pvtStartForm, Me, valStoryRange
  pvtStoriesCt = pvtStoriesCt + 1

End Sub

Private Function pvtCheckAffectedCitationGroupsCt(ByVal valAffectedCt As Integer, ByVal valNotAffectedCt As Integer, _
  ByVal valNotCorrectedCt As Integer, ByVal valDescription As String) As ZtFMessageType

  Dim locResult As ZtFMessageType
  
  locResult = pvtStartForm.MessageDisplay.Show("Result for " & valDescription & " citation groups:" & vbNewLine & _
                                                 vbTab & "affected: " & CStr(valAffectedCt) & vbNewLine & _
                                                 vbTab & "not affected: " & CStr(valNotAffectedCt) & vbNewLine & _
                                                 vbTab & "not corrected: " & CStr(valNotCorrectedCt), _
                                               IIf(valNotCorrectedCt > 0, MessageCancel, MessageOk) + IIf(valNotCorrectedCt > 0, MessageExclamation, MessageInformation), _
                                               False)
  
  pvtCheckAffectedCitationGroupsCt = locResult

End Function
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *



