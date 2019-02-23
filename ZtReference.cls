VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ZtReference"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Class ZtReference.
' It preserves information of a reference entry in the bibliography, i.e. text, Word bookmark, and corresponding citations.
'
' Zotero Tools.
' This software is under Revised ('New') BSD license.
' Copyright © 2019, Olaf Ahrens. All rights reserved.
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Private variables.
Private pvtConfig As ZtConfig
Private pvtMessageDisplay As ZtIMessageDisplayable
Private pvtBibliography As ZtBibliography
Private pvtRange As Word.Range
Private pvtMatch As ZtRegMatch
Private pvtBookmark As Word.Bookmark
Private pvtCitationGroups() As ZtCitationGroup
Private pvtCitationGroupsCt As Integer
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Friend procedures and properties.
Friend Function Initialize(ByVal valConfig As ZtConfig, ByVal valMessageDisplay As ZtIMessageDisplayable, ByVal valBibliography As ZtBibliography, _
  ByVal valRange As Word.Range) As ZtFMessageType
  
  Dim locResult As ZtFMessageType
  
  Set pvtConfig = valConfig
  Set pvtMessageDisplay = valMessageDisplay
  Set pvtBibliography = valBibliography
  Set pvtRange = valRange.Duplicate
  Set pvtMatch = pvtConfig.Final.Reference.WholeRegExp.FirstMatch(pvtRange.Text)
  If Not pvtMatch Is Nothing Then
    If pvtConfig.User.Macro.Debugging Then
      locResult = pvtConfig.Final.Reference.WholeRegExp.Debugging(pvtMatch, pvtRange)
    Else
      locResult = MessageOk
    End If
  ElseIf Len(Trim$(pvtRange.Text)) > 0 Then
    locResult = pvtMessageDisplay.Show("The reference style you set in ZtConfig.xml seems not to apply to the selected reference.", _
                                       MessageCritical + MessageCancel, _
                                       False, _
                                       pvtRange)
  Else
    locResult = MessageNo
  End If
  
  Initialize = locResult
  
End Function

Friend Property Get Nr() As Integer

  Nr = CInt(pvtMatch.Groups("number")) - 1
  
End Property

Friend Property Get ScreenTipText() As String

  ScreenTipText = Left$(Trim$(pvtMatch.Groups("text")), pvtConfig.Basic.Word.MaxScreenTipTextLength)
  
End Property

Friend Sub RemoveCitedOnPage()

  Dim locToCorrect As Boolean
  Dim locRange As Word.Range
        
  With pvtMatch
    ' Determine whether reference text has to be changed.
    If Len(.Groups("backwardLinking")) > 0 Then
    
      locToCorrect = True
      
    ElseIf Len(.Groups("accessedAtEnd")) > 0 And _
      .Groups("delimiterBeforeBackwardLinking") <> pvtConfig.User.ReferenceStyle.ReferenceAccessedAtEndSuffix Then
      
      locToCorrect = True
    
    ElseIf Len(.Groups("hyperlinkAtEnd")) > 0 And _
      .Groups("delimiterBeforeBackwardLinking") <> pvtConfig.User.ReferenceStyle.ReferenceHyperlinkAtEndSuffix Then
      
      locToCorrect = True
    
    ElseIf .Groups("delimiterBeforeBackwardLinking") <> pvtConfig.User.ReferenceStyle.ReferenceNormalSuffix Then
    
        locToCorrect = True
    
    End If
  
    ' If necessary change reference text.
    If locToCorrect Then
      Set locRange = pvtRange.Duplicate
      locRange.MoveStart wdCharacter, _
                         Len(.Groups("numbering")) + _
                           Len(.Groups("text")) + _
                           Len(.Groups("hyperlinkAtEnd")) + _
                           Len(.Groups("accessedAtEnd"))
      If Len(.Groups("accessedAtEnd")) > 0 Then
        locRange.Text = pvtConfig.User.ReferenceStyle.ReferenceAccessedAtEndSuffix
      ElseIf Len(.Groups("hyperlinkAtEnd")) > 0 Then
        locRange.Text = pvtConfig.User.ReferenceStyle.ReferenceHyperlinkAtEndSuffix
      Else
        locRange.Text = pvtConfig.User.ReferenceStyle.ReferenceNormalSuffix
      End If
    End If
  End With

End Sub

Friend Sub SetBookmark()

  Set pvtBookmark = pvtBibliography.MainStory.Document.AddBookmark(pvtConfig.Basic.Macro.ReferenceBookmarkPrefix & _
                                                                     ZtSubprocedures.AddLeadingZeros(Me.Nr + 1, _
                                                                                                     ZtSubprocedures.DigitsCt(pvtBibliography.ReferencesCt)), _
                                                                   pvtRange)
  
End Sub

Friend Sub AddCitationGroup(ByVal valCitationGroup As ZtCitationGroup)

  ReDim Preserve pvtCitationGroups(pvtCitationGroupsCt)
  Set pvtCitationGroups(pvtCitationGroupsCt) = valCitationGroup
  pvtCitationGroupsCt = pvtCitationGroupsCt + 1
  
End Sub

Friend Function SetWebLink(ByRef refAffectedReferencesCt As Integer) As ZtFMessageType

  Dim locMatch As ZtRegMatch
  Dim locWebLinkRange As Word.Range
  Dim locResult As ZtFMessageType
  
  Set locMatch = pvtConfig.Final.Reference.WebLinkRegExp.FirstMatch(pvtRange.Text)
  If Not locMatch Is Nothing Then
    If pvtConfig.User.Macro.Debugging Then
      locResult = pvtConfig.Final.Reference.WebLinkRegExp.Debugging(locMatch, pvtRange)
    Else
      locResult = MessageOk
    End If
    If locResult = MessageOk Then
      Set locWebLinkRange = pvtRange.Duplicate
      With locWebLinkRange
        .SetRange .Start + locMatch.FirstIndex, _
                  .Start + locMatch.FirstIndex + Len(locMatch.Groups("link"))
        If .Hyperlinks.Count = 0 Then
          If Len(locMatch.Groups("httpLink")) > 0 Then
            pvtBibliography.MainStory.Document.AddWebLink locWebLinkRange, _
                                                          .Text
          ElseIf Len(locMatch.Groups("doiLink")) > 0 Then
            pvtBibliography.MainStory.Document.AddWebLink locWebLinkRange, _
                                                          pvtConfig.Basic.Macro.DoiResolver & locMatch.Groups("doiLink")
          Else
            pvtBibliography.MainStory.Document.AddWebLink locWebLinkRange, _
                                                          pvtConfig.Basic.Macro.DoiResolver & locMatch.Groups("shortDoiLink")
          End If
          refAffectedReferencesCt = refAffectedReferencesCt + 1
        End If
      End With
    End If
  Else
    locResult = MessageOk
  End If
  
  SetWebLink = locResult

End Function

Friend Function SetInternalLinking() As ZtFMessageType

  Dim locRange As Word.Range
  Dim locCitationCtr As Integer
  Dim locCitationGroup As ZtCitationGroup
  Dim locDeleteRange As Word.Range
  Dim locDeleteProcedures As ZtRangeProcedures
  Dim locResult As ZtFMessageType
  Dim locDeleteCharacters As String
    
  Set locRange = pvtBookmark.Range.Duplicate
  Set locDeleteProcedures = New ZtRangeProcedures
  locDeleteCharacters = pvtConfig.Basic.Characters.Space & pvtConfig.Final.Reference.ReferenceUniversalSuffixCharacters
  
  ' Sort citations by their position in the document. Until now they are sorted by the Word stories they belong to.
  pvtSortCitationGroups
  
  With locRange
    If pvtCitationGroupsCt > 0 Then
      ' Correct delimiter and spaces at the end of the range if necessary.
      Set locDeleteRange = .Duplicate
      With locDeleteRange
        .Collapse wdCollapseEnd
        locDeleteProcedures.SetRange locDeleteRange
        locDeleteProcedures.MoveStartWhile locDeleteCharacters, wdBackward
        If Len(.Text) > 0 Then
          .Delete
        End If
      End With
      
      ' Insert introducting text.
      .InsertAfter pvtConfig.User.BackwardLinkingStyle.DelimiterBefore
      If pvtCitationGroupsCt = 1 Then
        .InsertAfter pvtConfig.User.BackwardLinkingStyle.PrefixSingular
      Else
        .InsertAfter pvtConfig.User.BackwardLinkingStyle.PrefixPlural
      End If
      
      ' Insert backward linking as crossreference which will update automatically, not fixed page numbers.
      ' Insert extra hyperlinks instead of crossreferences with hyperlinks, for screen tip texts.
      .Collapse wdCollapseEnd
      .InsertAfter " " ' For protection.
      .Collapse wdCollapseStart
      For locCitationCtr = 0 To pvtCitationGroupsCt - 1
        Set locCitationGroup = pvtCitationGroups(locCitationCtr)
        .InsertCrossReference wdRefTypeBookmark, wdPageNumber, locCitationGroup.Bookmark.Name, False, False, False
        .MoveEndUntil " ", wdForward
        pvtBibliography.MainStory.Document.AddInternalLink locRange, _
                                                           locCitationGroup.Bookmark.Name, _
                                                           locCitationGroup.Caption.ScreenTipText(.Information(wdActiveEndAdjustedPageNumber))
        .Style = pvtBibliography.MainStory.Document.Styles(pvtConfig.User.BackwardLinkingStyle.HyperlinkStyleName)
        If locCitationCtr < pvtCitationGroupsCt - 1 Then
          .Move wdCharacter, 1
          .InsertAfter pvtConfig.User.BackwardLinkingStyle.Delimiter
        End If
        .Collapse wdCollapseEnd
      Next
      .Move wdCharacter, 1
      .InsertAfter pvtConfig.User.BackwardLinkingStyle.Suffix
      .Collapse wdCollapseEnd
      .MoveEnd wdCharacter, 1
      .Delete ' Delete protection.
      locResult = MessageOk
    ElseIf pvtBibliography.ShowNoCitationError Then
      locResult = pvtMessageDisplay.Show("I couldn't find any citations of the selected reference." & vbNewLine & vbNewLine & _
                                           "You may procede, otherwise cancel.", _
                                         MessageCritical + MessageOkSuppressCancel, _
                                         True, _
                                         locRange)
      If locResult = MessageSuppress Then
        pvtBibliography.ShowNoCitationError = False
        locResult = MessageOk
      End If
    End If
  End With
  
  SetInternalLinking = locResult

End Function
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Private procedures and properties.
' Optimized bubblesort from https://de.wikipedia.org/wiki/Bubblesort
Private Sub pvtSortCitationGroups()

  Dim locCtr As Integer
  Dim locBubbleBorder As Integer
  Dim locNewBubbleBorder As Integer
  Dim locPageNr0 As Integer
  Dim locPageNr1 As Integer
  
  locBubbleBorder = pvtCitationGroupsCt - 1
  Do
    locNewBubbleBorder = 1
    For locCtr = 0 To locBubbleBorder - 1
      locPageNr0 = pvtCitationGroups(locCtr).Bookmark.Range.Information(wdActiveEndPageNumber)
      locPageNr1 = pvtCitationGroups(locCtr + 1).Bookmark.Range.Information(wdActiveEndPageNumber)
      If locPageNr0 >= locPageNr1 Then
        If locPageNr0 > locPageNr1 Then
          pvtSwapCitationGroups locCtr
          locNewBubbleBorder = locCtr
        ElseIf pvtCitationGroups(locCtr).CorrectedResultStart > pvtCitationGroups(locCtr + 1).CorrectedResultStart Then
          pvtSwapCitationGroups locCtr
          locNewBubbleBorder = locCtr
        End If
      End If
    Next
    locBubbleBorder = locNewBubbleBorder
  Loop While locBubbleBorder > 1
  
End Sub

Private Sub pvtSwapCitationGroups(ByVal valFirstCitationGroup As Integer)

  Dim locTempCitationGroup As ZtCitationGroup

  Set locTempCitationGroup = pvtCitationGroups(valFirstCitationGroup)
  Set pvtCitationGroups(valFirstCitationGroup) = pvtCitationGroups(valFirstCitationGroup + 1)
  Set pvtCitationGroups(valFirstCitationGroup + 1) = locTempCitationGroup

End Sub
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *



