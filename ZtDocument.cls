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
'
' Zotero Tools.
' This software is under Revised ('New') BSD license.
' Copyright © 2019, Olaf Ahrens. All rights reserved.
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Private variables.
Private pvtConfig As ZtConfig
Private pvtMessageDisplay As ZtIMessageDisplayable
Private pvtProgress As ZtProgress
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
Friend Sub Initialize(ByVal valConfig As ZtConfig, ByVal valDocument As Word.Document, ByVal valMessageDisplay As ZtIMessageDisplayable, ByVal valProgress As ZtProgress)
  
  Set pvtConfig = valConfig
  Set pvtMessageDisplay = valMessageDisplay
  Set pvtProgress = valProgress
  Set pvtDocument = valDocument
  
  ' Fix the skipped blank header/footer problem.
  ' From: https://wordmvp.com/FAQs/Customization/ReplaceAnywhere.htm.
  pvtJunk = pvtDocument.Sections(1).Headers(1).Range.StoryType
  
End Sub

Friend Sub SaveBackupCopy()

  Dim locFileNameCache As String

  With pvtDocument
    locFileNameCache = .FullName
    .SaveAs ZtSubprocedures.GetPath(pvtDocument.FullName) & _
            Application.PathSeparator & _
            ZtSubprocedures.GetFileNameWithoutExtension(pvtDocument.Name) & "_ZtBackup_" & Format$(Now, "YYYYMMDD_HHmmss") & ZtSubprocedures.GetExtension(pvtDocument.Name)
    pvtProgress.SetCompleteStep
    .SaveAs locFileNameCache
    pvtProgress.SetCompleteStep
  End With
  
End Sub

Friend Sub Prepare()
  
  Dim locView As Word.View
  
  If Not pvtIsPrepared Then
    Set locView = pvtDocument.ActiveWindow.View
    With locView
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
  
  Dim locView As Word.View
  
  If pvtIsPrepared Then
    Set locView = pvtDocument.ActiveWindow.View
    With locView
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
  
  locResult = 4 ' For pvtKeepStories, pvtKeepCitationGroups, and 2 x save the document.
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
  
  locResult = 4 ' For pvtKeepStories, pvtKeepCitationGroups, and 2 x save the document.
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
  
  locResult = 5 ' For pvtKeepCitationGroups, keep citation groups of one story, one story, and 2 x save the document.
  
  GetJoinCitationGroupsSelectionStepsCt = locResult
  
End Function

Friend Function GetRemoveInternalLinkingStepsCt() As Integer
  
  Dim locResult As Integer
  Dim locStoryCtr As Integer
  Dim locStory As ZtIStory
  
  locResult = 3 ' For pvtKeepStories and 2 x save the document.
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
  
  locResult = 7 ' For pvtKeepStories, pvtKeepCitationGroups, SetBibliographyBookmarks, pvtKeepReferences, SetBookmarks, and 2 x save the document.
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
  
  locResult = 4 ' For pvtKeepReferences, pvtKeepCitationGroups, and 2 x save the document.
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
  
  locResult = 5 ' For pvtKeepCitationGroups, pvtKeepReferences, SetWebLinks, and 2 x save the document.
  
  GetSetWebLinksStepsCt = locResult
  
End Function

Friend Function AdjustPunctuation() As ZtFMessageType

  Dim locStoryCtr As Integer
  Dim locStory As ZtIStory
  Dim locResult As ZtFMessageType
  Dim locAffectedCt As Integer
  Dim locNotAffectedCt As Integer
  Dim locNotCorrectedCt As Integer
  
  pvtKeepStories
  locResult = pvtKeepCitationGroups
  If locResult = MessageOk Then
    For locStoryCtr = 0 To pvtStoriesCt - 1
      Set locStory = pvtStories(locStoryCtr)
      With locStory
        If .CanHoldCitations Then
          locResult = .AdjustPunctuation(locAffectedCt, locNotAffectedCt, locNotCorrectedCt)
          If locResult = MessageCancel Then
            Exit For
          End If
        End If
      End With
    Next
    If locResult = MessageOk Then
      pvtProgress.SetCompleted
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
  ' Delete hyperlinks to references and citations, bookmarks of references and citations, and hidden spaces.
  For locStoryCtr = 0 To pvtStoriesCt - 1
    Set locStory = pvtStories(locStoryCtr)
    If locStory.CanHoldCitations Then
      locResult = locStory.RemoveInternalLinking()
      If locResult = MessageCancel Then
        Exit For
      End If
    End If
  Next
  pvtProgress.SetCompleted
  
  RemoveInternalLinking = locResult

End Function

Friend Function SetInternalLinking() As ZtFMessageType

  Dim locStoryCtr As Integer
  Dim locStory As ZtIStory
  Dim locResult As ZtFMessageType
  
  pvtKeepStories
  locResult = pvtKeepCitationGroups
  If locResult = MessageOk Then
    ' Set bookmarks on references.
    locResult = MainStory.SetBibliographyBookmarks
    If locResult = MessageOk Then
    
      ' Pick captions.
      If pvtConfig.User.Macro.WithBackwardLinking Then
        MainStory.KeepCaptions
      End If
        
      ' Link citations to references.
      For locStoryCtr = 0 To pvtStoriesCt - 1
        Set locStory = pvtStories(locStoryCtr)
        If locStory.CanHoldCitations Then
          locResult = locStory.SetInternalLinking
          If locResult = MessageCancel Then
            Exit For
          End If
        End If
      Next
      
      ' If necessary link reference to citation (backward linking).
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
    pvtProgress.SetToRunningMode
    For locStoryCtr = 0 To pvtStoriesCt - 1
      Set locStory = pvtStories(locStoryCtr)
      If locStory.CanHoldCitations Then
        locStory.JoinCitationGroups locAffectedCt, locNotAffectedCt
        pvtProgress.SetSubstep
      End If
    Next
    pvtProgress.SetCompleted
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
    With locStory
      If locRange.InRange(.Range) Then
        .KeepCitationGroups pvtCitationGroupsCt
        pvtProgress.SetStep
        pvtProgress.SetToRunningMode
        .JoinCitationGroups locAffectedCt, 0, locRange
        Exit For
      End If
    End With
  Next
  pvtProgress.SetCompleted
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
      With locStory
        If Not .CanHoldCitations Then
          .ResolveCitationGroups locAffectedCt
        End If
      End With
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

' Thanks to https://superuser.com/questions/196703/how-do-i-update-all-fields-in-a-word-document.
Private Sub pvtUpdateAllFields()

  Dim locStoryCtr As Integer
  
  ' Update fields everywhere. This includes updates of page numbers in tables (but would not add or remove entries).
  ' This also takes care of all index updates.
  pvtProgress.SetStep pvtStoriesCt
  For locStoryCtr = 0 To pvtStoriesCt - 1
    pvtStories(locStoryCtr).UpdateAllFields
    pvtProgress.SetSubstep
  Next
    
End Sub

Private Sub pvtKeepStories()

  Dim locStoryRange As Word.Range
  Dim locShapeRange As Word.ShapeRange

  If pvtStoriesCt = -1 Then
    pvtStoriesCt = 0
    pvtProgress.SetStep pvtDocument.StoryRanges.Count
    For Each locStoryRange In pvtDocument.StoryRanges
      With locStoryRange
        If Not ZtSubprocedures.ArrayContains(.StoryType, pvtConfig.Basic.Word.CantHoldFieldStoryTypes) Then
          ' Word can't insert fields into 'normal' text frames (called 'textboxes', also we are in a text frame story), so we can skip this story.
          ' Word can insert fields into frames stored in form fields. But these are part of and queued in the Word main story.
          ' Further informations about frames and textboxes: http://www.addbalance.com/word/frames_textboxes.htm.
          Do
            pvtStoryFactory locStoryRange
            
            ' From: https://wordmvp.com/FAQs/Customization/ReplaceAnywhere.htm.
            If ZtSubprocedures.ArrayContains(.StoryType, pvtConfig.Basic.Word.MultipleStoriesStoryTypes) Then
              If .ShapeRange.Count > 0 Then
                For Each locShapeRange In .ShapeRange
                  With locShapeRange
                    If .TextFrame.HasText Then
                      pvtStoryFactory .TextFrame.TextRange
                    End If
                  End With
                Next
              End If
              
            End If
            
            Set locStoryRange = locStoryRange.NextStoryRange
            
          Loop Until locStoryRange Is Nothing
        End If
      End With
      pvtProgress.SetSubstep
    Next
  Else
    pvtProgress.SetCompleteStep
  End If
  
End Sub

Private Function pvtKeepCitationGroups() As ZtFMessageType

  Dim locResult As ZtFMessageType
  Dim locStoryCtr As Integer
  
  If pvtCitationGroupsCt = -1 Then
    pvtCitationGroupsCt = 0
    pvtProgress.SetStep pvtStoriesCt
    For locStoryCtr = 0 To pvtStoriesCt - 1
      pvtStories(locStoryCtr).KeepCitationGroups pvtCitationGroupsCt
      pvtProgress.SetSubstep
    Next
    
    If pvtCitationGroupsCt = 0 Then
      locResult = pvtMessageDisplay.Show("No citation groups found!" & vbNewLine & vbNewLine & _
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
  pvtStories(pvtStoriesCt).Initialize pvtConfig, pvtMessageDisplay, pvtProgress, Me, valStoryRange
  pvtStoriesCt = pvtStoriesCt + 1

End Sub

Private Function pvtCheckAffectedCitationGroupsCt(ByVal valAffectedCt As Integer, ByVal valNotAffectedCt As Integer, _
  ByVal valNotCorrectedCt As Integer, ByVal valDescription As String) As ZtFMessageType

  Dim locResult As ZtFMessageType
  
  locResult = pvtMessageDisplay.Show("Result for " & valDescription & " citation groups:" & vbNewLine & _
                                       vbTab & "affected: " & CStr(valAffectedCt) & vbNewLine & _
                                       vbTab & "not affected: " & CStr(valNotAffectedCt) & vbNewLine & _
                                       vbTab & "not corrected: " & CStr(valNotCorrectedCt), _
                                     IIf(valNotCorrectedCt > 0, MessageCancel, MessageOk) + IIf(valNotCorrectedCt > 0, MessageExclamation, MessageInformation), _
                                     False)
  
  pvtCheckAffectedCitationGroupsCt = locResult

End Function
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *




