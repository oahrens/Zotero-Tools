VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ZtStory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements ZtIStory


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Class ZtStory.
' It preserves informations and procedures of a Word story inside a document.
' E.g. the main text, footnotes, endnotes, and comments are different Word stories.
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
Private pvtDocument As ZtDocument
Private pvtRange As Word.Range
Private pvtCitationGroups() As ZtCitationGroup
Private pvtCitationGroupsCt As Integer
Private pvtCitationGroupsOffset As Integer
Private pvtCorrelations() As ZtStoryCorrelation
Private pvtCorrelationsCt As Integer
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Constructor.
Private Sub Class_Initialize()

  pvtCitationGroupsCt = -1
  pvtCorrelationsCt = -1
  
End Sub
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Private interface procedures and properties directing to Friend procedures and properties below.
Private Sub ZtIStory_Initialize(ByVal valConfig As ZtConfig, ByVal valMessageDisplay As ZtIMessageDisplayable, ByVal valProgress As ZtProgress, _
  ByVal valDocument As ZtDocument, ByVal valRange As Word.Range)
  
  Me.Initialize valConfig, valMessageDisplay, valProgress, valDocument, valRange
  
End Sub

Private Property Get ZtIStory_CanHoldCitations() As Boolean

  ZtIStory_CanHoldCitations = Me.CanHoldCitations
  
End Property

Private Sub ZtIStory_KeepCitationGroups(ByRef refCitationGroupsOffset As Integer)
  
  Me.KeepCitationGroups refCitationGroupsOffset

End Sub

Private Function ZtIStory_RemoveInternalLinking() As ZtFMessageType

  ZtIStory_RemoveInternalLinking = Me.RemoveInternalLinking

End Function

Private Function ZtIStory_SetInternalLinking() As ZtFMessageType

  ZtIStory_SetInternalLinking = Me.SetInternalLinking

End Function

Private Function ZtIStory_AdjustPunctuation(ByRef refAffectedCt As Integer, ByRef refNotAffectedCt As Integer, ByRef refNotCorrectedCt As Integer) As ZtFMessageType

  ZtIStory_AdjustPunctuation = Me.AdjustPunctuation(refAffectedCt, refNotAffectedCt, refNotCorrectedCt)
  
End Function

Private Sub ZtIStory_ResolveCitationGroups(ByRef refAffectedCt As Integer)

  Me.ResolveCitationGroups refAffectedCt

End Sub

Private Sub ZtIStory_UpdateAllFields()

  Me.UpdateAllFields

End Sub

Private Property Get ZtIStory_Range() As Word.Range

  Set ZtIStory_Range = Me.Range
  
End Property

Private Property Get ZtIStory_Document() As ZtDocument

  Set ZtIStory_Document = Me.Document
  
End Property

Private Sub ZtIStory_JoinCitationGroups(ByRef refAffectedCt As Integer, ByRef refNotAffectedCt As Integer, Optional ByVal valRange As Word.Range = Nothing)

  Me.JoinCitationGroups refAffectedCt, refNotAffectedCt, valRange
  
End Sub
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *



' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Friend procedures and properties.
Friend Sub Initialize(ByVal valConfig As ZtConfig, ByVal valMessageDisplay As ZtIMessageDisplayable, ByVal valProgress As ZtProgress, _
  ByVal valDocument As ZtDocument, ByVal valRange As Word.Range)
  
  Set pvtConfig = valConfig
  Set pvtMessageDisplay = valMessageDisplay
  Set pvtProgress = valProgress
  Set pvtDocument = valDocument
  Set pvtRange = valRange.Duplicate
  
End Sub

Friend Property Get CanHoldCitations() As Boolean

  CanHoldCitations = ZtSubprocedures.ArrayContains(pvtRange.StoryType, pvtConfig.Basic.Zotero.CanHoldCitationStoryTypes)
  
End Property

Friend Sub KeepCitationGroups(ByRef refCitationGroupsOffset As Integer)

  Dim locField As Word.Field
  
  If pvtCitationGroupsCt = -1 Then
    pvtCitationGroupsCt = 0
    pvtCitationGroupsOffset = refCitationGroupsOffset
    For Each locField In pvtRange.Fields
      If InStr(locField.Code.Text, pvtConfig.Basic.Zotero.CitationGroupIdentifier) > 0 Then
        pvtCitationGroupFactory locField
      End If
    Next
  End If
  refCitationGroupsOffset = refCitationGroupsOffset + pvtCitationGroupsCt
  
End Sub

Friend Function RemoveInternalLinking() As ZtFMessageType

  Dim locCtr As Integer
  Dim locBookmarks() As Word.Bookmark
  Dim locBookmark As Word.Bookmark
  Dim locFields() As Word.Field
  Dim locField As Word.Field
  Dim locInsertRange As Word.Range
  Dim locInsertText As String
  Dim locFormat As ZtCharacterFormat
  
  With pvtRange
    
    ' Delete internal hyperlinks in citation groups.
    ' It's not possible to delete Word hyperlinks if they're nested inside another field and border directly on the end of the nesting field.
    ' So we have to delete the Word fields of type wdFieldHyperlink and re-insert the result text of the deleted field.
    ' Another aproache would be to insert a zero width space between the ends of the hyperlink and its nesting field when setting the hyperlink.
    ' But choosing this aproach Zotero's refresh will alarm on each citation because of the citation text has been changed.
    
    ' We can't delete an item of a collection while looping in a For-Each-Next loop.
    ' For this purpose we have to loop backwards through the collection, using a For-To-Next loop.
    ' This approach is EXTREM slow! So the fastest procedure is first to collect each item in an array
    ' and to loop backwards through this array.
    
    If .Fields.Count > 0 Then
      pvtKeepFields locFields
      pvtProgress.SetStep UBound(locFields)
      For locCtr = UBound(locFields) To 0 Step -1
        Set locField = locFields(locCtr)
        With locField
          If .Type = wdFieldHyperlink Then
            If InStr(.Code.Text, pvtConfig.Basic.Macro.ReferenceBookmarkPrefix) > 0 Then
              Set locInsertRange = .Result
              Set locFormat = New ZtCharacterFormat
              locFormat.SetFormat locInsertRange
              locInsertText = locInsertRange.Text
              Do While Left$(locInsertText, 1) = pvtConfig.Basic.Characters.ZeroWidthSpace
                locInsertText = Right$(locInsertText, Len(locInsertText) - 1)
              Loop
              locInsertRange.Collapse wdCollapseStart
              .Delete
              locInsertRange.Text = locInsertText
              locFormat.GetFormat locInsertRange
            End If
          End If
        End With
        pvtProgress.SetSubstep
      Next
    Else
      pvtProgress.SetCompleteStep
    End If
    
    ' Delete bookmarks that have been referenced by internal hyperlinks.
    If .Bookmarks.Count > 0 Then
      pvtKeepBookmarks locBookmarks
      pvtProgress.SetStep UBound(locBookmarks)
      For locCtr = UBound(locBookmarks) To 0 Step -1
        Set locBookmark = locBookmarks(locCtr)
        With locBookmark
          If InStr(.Name, pvtConfig.Basic.Macro.CitationBookmarkPrefix) > 0 Then
            .Delete
          ElseIf InStr(.Name, pvtConfig.Basic.Macro.ReferenceBookmarkPrefix) > 0 Then
            .Delete
          End If
        End With
      Next
      pvtProgress.SetSubstep
    Else
      pvtProgress.SetCompleteStep
    End If
        
  End With
  
  RemoveInternalLinking = MessageOk
  
End Function

Friend Function SetInternalLinking() As ZtFMessageType

  Dim locCitationGroup As ZtCitationGroup
  Dim locCitationGroupCtr As Integer
  Dim locCaptions() As ZtCaption
  Dim locResult As ZtFMessageType
  
  If pvtConfig.User.Macro.WithBackwardLinking Then
    pvtProgress.SetStep pvtCitationGroupsCt * 2
  Else
    pvtProgress.SetStep pvtCitationGroupsCt
  End If
  For locCitationGroupCtr = 0 To pvtCitationGroupsCt - 1
    Set locCitationGroup = pvtCitationGroups(locCitationGroupCtr)

    ' Save bookmarks for backward linking if necessary.
    If pvtConfig.User.Macro.WithBackwardLinking Then
      pvtDocument.MainStory.GetCaptions locCaptions
      pvtKeepStoryCorrelations
      locCitationGroup.SetBookmark locCaptions, pvtCorrelations, pvtCorrelationsCt
      pvtProgress.SetSubstep
    End If
    
    ' Set hyperlink on each citation.
    locResult = locCitationGroup.SetInternalLinking
    pvtProgress.SetSubstep
    If locResult = MessageCancel Then
      Exit For
    End If
  Next
  
  SetInternalLinking = locResult

End Function

Friend Function AdjustPunctuation(ByRef refAffectedCt As Integer, ByRef refNotAffectedCt As Integer, ByRef refNotCorrectedCt As Integer) As ZtFMessageType

  Dim locCitationGroupCtr As Integer
  Dim locResult As ZtFMessageType
  
  pvtProgress.SetStep pvtCitationGroupsCt
  For locCitationGroupCtr = 0 To pvtCitationGroupsCt - 1
    locResult = pvtCitationGroups(locCitationGroupCtr).AdjustPunctuation(refAffectedCt, refNotAffectedCt, refNotCorrectedCt)
    pvtProgress.SetSubstep
    If locResult = MessageCancel Then
      Exit For
    End If
  Next
  
  AdjustPunctuation = locResult
  
End Function

Friend Sub ResolveCitationGroups(ByRef refAffectedCt As Integer)

  Dim locCitationGroupCtr As Integer
  
  pvtProgress.SetStep pvtCitationGroupsCt
  For locCitationGroupCtr = 0 To pvtCitationGroupsCt - 1
    pvtCitationGroups(locCitationGroupCtr).Resolve
    refAffectedCt = refAffectedCt + 1
    pvtProgress.SetSubstep
  Next
  pvtCitationGroupsCt = -1

End Sub

Friend Sub UpdateAllFields()

  Dim locField As Word.Field
      
  ' To prevent pop-up 'Word cannot undo this action. Do you want to continue?'.
  If ZtSubprocedures.ArrayContains(pvtRange.StoryType, pvtConfig.Basic.Word.IndividuallyUpdateFieldStoryTypes) Then
    
    For Each locField In pvtRange.Fields
      locField.Update
    Next
    
  Else
    pvtRange.Fields.Update
  End If

End Sub

Friend Property Get Range() As Word.Range

  Set Range = pvtRange
  
End Property

Friend Property Get Document() As ZtDocument

  Set Document = pvtDocument
  
End Property

Friend Sub JoinCitationGroups(ByRef refAffectedCt As Integer, ByRef refNotAffectedCt As Integer, Optional ByVal valRange As Word.Range = Nothing)
  
  Dim locBetweenText As String
  Dim locCitationGroupCtr As Integer
  Dim locCitationGroupIsFound As Boolean
  Dim locInsertRemnantsRange As Word.Range
  Dim locInsertRemnantsProcedures As ZtRangeProcedures
  
  If valRange Is Nothing Then
    For locCitationGroupCtr = pvtCitationGroupsCt - 1 To 1 Step -1
      Set locInsertRemnantsRange = pvtCitationGroups(locCitationGroupCtr).ResultRange
      With locInsertRemnantsRange
        .Move wdCharacter, 1
        locBetweenText = pvtCitationGroups(locCitationGroupCtr).JoinWithPrevious(True, refAffectedCt, refNotAffectedCt)
        If Len(locBetweenText) > 0 Then
          .MoveStart wdCharacter, -Len(locBetweenText)
          .Delete
        End If
      End With
      pvtProgress.SetSubstep
    Next
  Else
    For locCitationGroupCtr = pvtCitationGroupsCt - 1 To 1 Step -1
      If pvtCitationGroups(locCitationGroupCtr).ResultRange.InRange(valRange) Then
        If pvtCitationGroups(locCitationGroupCtr - 1).ResultRange.InRange(valRange) Then
          If Not locCitationGroupIsFound Then
            locCitationGroupIsFound = True
            Set locInsertRemnantsRange = pvtCitationGroups(locCitationGroupCtr).ResultRange
            locInsertRemnantsRange.SetRange pvtCitationGroups(locCitationGroupCtr - 1).ResultRange.End + 1, pvtCitationGroups(locCitationGroupCtr).CodeRange.Start - 1
          End If
          locBetweenText = locBetweenText & pvtCitationGroups(locCitationGroupCtr).JoinWithPrevious(False, refAffectedCt, 0)
        ElseIf pvtCitationGroups(locCitationGroupCtr - 1).ResultRange.Start < valRange.Start Then
          Exit For
        End If
      ElseIf pvtCitationGroups(locCitationGroupCtr).ResultRange.Start < valRange.Start Then
        Exit For
      End If
      pvtProgress.SetSubstep
    Next
    
    If locCitationGroupIsFound Then
      Set locInsertRemnantsProcedures = New ZtRangeProcedures
      With locInsertRemnantsProcedures
        .SetRange locInsertRemnantsRange
        .MoveStartWhile pvtConfig.Basic.Characters.SpaceWOZero, wdBackward
        .MoveEndWhile pvtConfig.Basic.Characters.SpaceWOZero, wdForward
      End With
      
      locBetweenText = pvtConfig.Final.CitationGroup.BetweenDeleteAffixedSpaceRegExp.Replace( _
                         pvtConfig.Final.CitationGroup.BetweenDeleteMultipleSpaceRegExp.Replace( _
                           pvtConfig.Final.CitationGroup.BetweenDeleteRegExp.Replace(locBetweenText, _
                                                                                    vbNullString), _
                           " "), _
                         "${text}")
      If Len(locBetweenText) > 0 Then
        locInsertRemnantsRange.Text = " " & locBetweenText & " "
      Else
        locInsertRemnantsRange.Text = " "
      End If
    End If
  End If
  pvtCitationGroupsCt = -1

End Sub
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Private procedures and properties.
Private Sub pvtKeepFields(ByRef refFields() As Word.Field)
    
  Dim locCtr As Integer
  Dim locField As Word.Field
  
  With pvtRange
    ReDim refFields(.Fields.Count - 1)
    For Each locField In .Fields
      Set refFields(locCtr) = locField
      locCtr = locCtr + 1
    Next
  End With
  
End Sub

Private Sub pvtKeepBookmarks(ByRef refBookmarks() As Word.Bookmark)
    
  Dim locCtr As Integer
  Dim locBookmark As Word.Bookmark
  
  With pvtRange
    ReDim Preserve refBookmarks(.Bookmarks.Count - 1)
    locCtr = 0
    For Each locBookmark In .Bookmarks
      Set refBookmarks(locCtr) = locBookmark
      locCtr = locCtr + 1
    Next
  End With
  
End Sub

Private Sub pvtKeepStoryCorrelations()

  Dim locFootnote As Word.Footnote
  Dim locEndnote As Word.Endnote
  
  If pvtCorrelationsCt = -1 Then
    pvtCorrelationsCt = 0
    If pvtRange.StoryType = wdFootnotesStory Then
      For Each locFootnote In pvtDocument.Footnotes
        With locFootnote
          pvtCorrelationFactory .Reference.Start, .Range.Start
        End With
      Next
    ElseIf pvtRange.StoryType = wdEndnotesStory Then
      For Each locEndnote In pvtDocument.Endnotes
        With locEndnote
          pvtCorrelationFactory .Reference.Start, .Range.Start
        End With
      Next
    End If
  End If
  
End Sub

Private Sub pvtCorrelationFactory(ByVal valStartInMainStory As Long, ByVal valStartInSubStory As Long)

  ReDim Preserve pvtCorrelations(pvtCorrelationsCt)
  Set pvtCorrelations(pvtCorrelationsCt) = New ZtStoryCorrelation
  pvtCorrelations(pvtCorrelationsCt).Initialize valStartInMainStory, valStartInSubStory
  pvtCorrelationsCt = pvtCorrelationsCt + 1

End Sub

Private Sub pvtCitationGroupFactory(ByVal valField As Word.Field)

  ReDim Preserve pvtCitationGroups(pvtCitationGroupsCt)
  Set pvtCitationGroups(pvtCitationGroupsCt) = New ZtCitationGroup
  If pvtCitationGroupsCt = 0 Then
    pvtCitationGroups(pvtCitationGroupsCt).Initialize pvtConfig, pvtMessageDisplay, Me, valField, pvtCitationGroupsOffset, pvtCitationGroupsCt
  Else
    pvtCitationGroups(pvtCitationGroupsCt).Initialize pvtConfig, pvtMessageDisplay, Me, valField, pvtCitationGroupsOffset, pvtCitationGroupsCt, pvtCitationGroups(pvtCitationGroupsCt - 1)
    Set pvtCitationGroups(pvtCitationGroupsCt - 1).NextCitationGroup = pvtCitationGroups(pvtCitationGroupsCt)
  End If
  pvtCitationGroupsCt = pvtCitationGroupsCt + 1

End Sub
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


