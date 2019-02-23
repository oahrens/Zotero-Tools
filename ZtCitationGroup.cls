VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ZtCitationGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Class ZtCitationGroup.
' It preserves information about a citation group, i.e. '[1-3, 7]'.
'
' Zotero Tools.
' This software is under Revised ('New') BSD license.
' Copyright © 2019, Olaf Ahrens. All rights reserved.
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Private variables.
Private pvtConfig As ZtConfig
Private pvtMessageDisplay As ZtIMessageDisplayable
Private pvtStory As ZtIStory
Private pvtField As Word.Field
Private pvtFieldProcedures As ZtFieldProcedures
Private pvtBookmark As Word.Bookmark
Private pvtCaption As ZtCaption
Private pvtCorrectedStart As Double
Private pvtOffset As Integer
Private pvtNr As Integer
Private pvtCitationInfos() As ZtCitationInfo
Private pvtCitationInfosCt As Integer
Private pvtPreviousCitationGroup As ZtCitationGroup
Private pvtNextCitationGroup As ZtCitationGroup
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Constructor.
Private Sub Class_Initialize()

  Set pvtFieldProcedures = New ZtFieldProcedures

End Sub
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Friend procedures and properties.
Friend Sub Initialize(ByVal valConfig As ZtConfig, ByVal valMessageDisplay As ZtIMessageDisplayable, ByVal valStory As ZtIStory, ByVal valField As Word.Field, _
  ByVal valOffset As Integer, ByVal valNr As Integer, Optional ByVal valPreviousCitationGroup As ZtCitationGroup = Nothing)
  
  Set pvtConfig = valConfig
  Set pvtMessageDisplay = valMessageDisplay
  Set pvtStory = valStory
  Set pvtField = valField
  pvtFieldProcedures.Initialize pvtField
  pvtOffset = valOffset
  pvtNr = valNr
  Set pvtPreviousCitationGroup = valPreviousCitationGroup
  
End Sub

Friend Property Set NextCitationGroup(ByVal valNextCitationGroup As ZtCitationGroup)

  Set pvtNextCitationGroup = valNextCitationGroup
  
End Property

Friend Property Get Bookmark() As Word.Bookmark

  Set Bookmark = pvtBookmark
  
End Property

Friend Property Get CodeRange() As Word.Range

  Set CodeRange = pvtField.Code
  
End Property

Friend Property Get ResultRange() As Word.Range

  Set ResultRange = pvtField.Result
  
End Property

Friend Property Get CorrectedResultStart() As Double

  CorrectedResultStart = pvtCorrectedStart
  
End Property

Friend Property Get Caption() As ZtCaption

  Set Caption = pvtCaption
  
End Property

Friend Function AdjustPunctuation(ByRef refAffectedCt As Integer, ByRef refNotAffectedCt As Integer, ByRef refNotCorrectedCt As Integer) As ZtFMessageType
  
  Dim locCompleteRange As Word.Range
  Dim locCompleteProcedures As ZtRangeProcedures
  Dim locBeforeRange As Word.Range
  Dim locBeforeProcedures As ZtRangeProcedures
  Dim locBeforeBackup As Word.Range
  Dim locAfterRange As Word.Range
  Dim locAfterProcedures As ZtRangeProcedures
  Dim locAfterBackup As Word.Range
  Dim locMatch As ZtRegMatch
  Dim locToChange As Boolean
  Dim locHasBeenChanged As Boolean
  Dim locNewBefore As ZtStringBuilder
  Dim locNewAfter As ZtStringBuilder
  Dim locResult As ZtFMessageType
  Dim locLengthBackup As Long
  Dim locSurroundingSigns As ZtFSurroundingSigns
  Dim locFindText As String
  
  ' Delete hidden character formats.
  pvtFieldProcedures.RemoveCodeFormat

  ' Set ranges.
  Set locCompleteRange = pvtFieldProcedures.GetCompleteRange
  Set locBeforeRange = locCompleteRange.Duplicate
  locBeforeRange.Collapse wdCollapseStart
  Set locAfterRange = locCompleteRange.Duplicate
  locAfterRange.Collapse wdCollapseEnd
  Set locNewBefore = New ZtStringBuilder
  Set locNewAfter = New ZtStringBuilder
  Set locCompleteProcedures = New ZtRangeProcedures
  Set locBeforeProcedures = New ZtRangeProcedures
  Set locAfterProcedures = New ZtRangeProcedures
  
  With locCompleteProcedures
    .SetRange locCompleteRange
    .MoveStartWhile pvtConfig.Basic.Characters.SpaceWOZeroPunctuation, wdBackward
    locLengthBackup = Len(locCompleteRange.Text)
    .MoveStartUntil pvtConfig.Basic.Characters.PunctuationBreakField, -1
    If locLengthBackup < Len(locCompleteRange.Text) Then
      locSurroundingSigns = locSurroundingSigns Or SignBefore
    End If
    .MoveEndWhile pvtConfig.Basic.Characters.Punctuation, wdForward
    locLengthBackup = Len(locCompleteRange.Text)
    .MoveEndUntil pvtConfig.Basic.Characters.PunctuationBreakField, 1
    If locLengthBackup < Len(locCompleteRange.Text) Then
      locSurroundingSigns = locSurroundingSigns Or SignAfter
    End If
  End With
  
  With locBeforeProcedures
    .SetRange locBeforeRange
    .MoveStartWhile pvtConfig.Basic.Characters.SpaceWOZeroPunctuation, wdBackward
    .MoveStartUntil pvtConfig.Basic.Characters.PunctuationBreakField, -1
    Set locBeforeBackup = locBeforeRange.Duplicate
    locBeforeRange.Collapse wdCollapseStart
  End With
  
  With locAfterProcedures
    .SetRange locAfterRange
    .MoveEndWhile pvtConfig.Basic.Characters.Punctuation, wdForward
    .MoveEndUntil pvtConfig.Basic.Characters.PunctuationBreakField, 1
    Set locAfterBackup = locAfterRange.Duplicate
    locAfterRange.Collapse wdCollapseEnd
  End With
  
  ' Adjust punctuation.
  Set locMatch = pvtConfig.Final.CitationGroup.PunctuationRegExp(locSurroundingSigns).FirstMatch(locCompleteRange.Text)
  If Not locMatch Is Nothing Then
    If pvtConfig.User.Macro.Debugging Then
      locResult = pvtConfig.Final.CitationGroup.PunctuationRegExp(locSurroundingSigns).Debugging(locMatch, locCompleteRange)
    Else
      locResult = MessageOk
    End If
    If locResult = MessageOk Then
      With locMatch
      
        If pvtConfig.User.PunctuationStyle.PunctuationIsBefore And Len(.Groups("punctuationAfter")) > 0 Then
          locToChange = True
        ElseIf Not pvtConfig.User.PunctuationStyle.PunctuationIsBefore And Len(.Groups("punctuationBefore")) > 0 Then
          locToChange = True
        ElseIf pvtConfig.User.PunctuationStyle.SpaceIsBefore <> (Len(.Groups("spaceBefore")) > 0) Then
          locToChange = True
        End If
        
        If locToChange Then
          locNewBefore.Append .Groups("signBefore")
          If pvtConfig.User.PunctuationStyle.PunctuationIsBefore Then
            locNewBefore.Append .Groups("punctuationBefore")
            locNewBefore.Append .Groups("punctuationAfter")
            If pvtConfig.User.PunctuationStyle.SpaceIsBefore Then
              locNewBefore.Append " "
            End If
          Else
            If pvtConfig.User.PunctuationStyle.SpaceIsBefore Then
              locNewBefore.Append " "
            End If
            locNewAfter.Append .Groups("punctuationBefore")
            locNewAfter.Append .Groups("punctuationAfter")
          End If
          locNewAfter.Append .Groups("signAfter")
    
          If Len(locBeforeBackup.Text) = locBeforeBackup.End - locBeforeBackup.Start And _
            Len(locAfterBackup.Text) = locAfterBackup.End - locAfterBackup.Start Then
            
            locFindText = .Groups("signBefore") & .Groups("punctuationBefore") & .Groups("spaceBefore")
            With locBeforeRange
              If Len(locFindText) > 0 Then
                ZtSubprocedures.ResetFind .Find
                With .Find
                  .Text = locFindText
                  With .Replacement
                    .Text = locNewBefore.ToString
                  End With
          
                  On Error GoTo OnErrorBeforeRange
                  locHasBeenChanged = .Execute(Replace:=wdReplaceOne)
                  On Error GoTo 0
                End With
              Else
                .InsertAfter locNewBefore.ToString
                locHasBeenChanged = True
              End If
            End With
    
            If locHasBeenChanged Then
              locFindText = .Groups("punctuationAfter") & .Groups("signAfter")
              With locAfterRange
                If Len(locFindText) > 0 Then
                  ZtSubprocedures.ResetFind .Find
                  With .Find
                    .Forward = False
                    .Text = locFindText
                    With .Replacement
                      .Text = locNewAfter.ToString
                    End With
            
                    On Error GoTo OnErrorAfterRange
                    locHasBeenChanged = .Execute(Replace:=wdReplaceOne)
                    On Error GoTo 0
                  End With
                Else
                  .InsertBefore locNewAfter.ToString
                  locHasBeenChanged = True
                End If
              End With
            End If
            If locHasBeenChanged Then
              refAffectedCt = refAffectedCt + 1
            Else
              refNotCorrectedCt = refNotCorrectedCt + 1
            End If
          Else
            refNotCorrectedCt = refNotCorrectedCt + 1
            If pvtStory.Document.ShowManualAdjustation Then
              ' Disturbing comments or fields.
              locResult = pvtMessageDisplay.Show("The selected citation group must be changed manually!" & vbNewLine & vbNewLine & _
                                                   "With 'Disable' this error won't be shown anymore, otherwise procede or cancel.", _
                                                 MessageExclamation + MessageOkDisableCancel, _
                                                 True, _
                                                 locCompleteRange)
              If locResult = MessageDisable Then
                pvtStory.Document.ShowManualAdjustation = False
                locResult = MessageOk
              End If
            End If
          End If
        Else
          refNotAffectedCt = refNotAffectedCt + 1
        End If
      End With
    End If
  Else
    locResult = pvtMessageDisplay.Show("The citation group style you set in ZtConfig.xml seems not to apply to the selected citation group." & vbNewLine & vbNewLine & _
                                         "You may procede, otherwise cancel.", _
                                       MessageCritical + MessageOkCancel, _
                                       True, _
                                       locCompleteRange)
  End If

  AdjustPunctuation = locResult

  Exit Function
  
OnErrorBeforeRange:
  Set locBeforeRange = locBeforeBackup
  With locBeforeRange
    If Err.Number = 4198 Then
      .MoveStart wdCharacter, 1
      If Len(.Text) > 0 Then ' For a collapsed range Range.Characters.Count results in 1!
        .Delete
      End If
      .InsertBefore Right$(locNewBefore.ToString, locNewBefore.Length - 1)
      Resume Next
    Else
      Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpContext
    End If
  End With
  
OnErrorAfterRange:
  Set locAfterRange = locAfterBackup
  With locAfterRange
    If Err.Number = 4198 Then
      .MoveEnd wdCharacter, -1
      If Len(.Text) > 0 Then ' For a collapsed range Range.Characters.Count results in 1!
        .Delete
      End If
      .InsertAfter Left$(locNewAfter.ToString, locNewAfter.Length - 1)
      Resume Next
    Else
      Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpContext
    End If
  End With

End Function

Friend Sub SetBookmark(ByRef refCaptions() As ZtCaption, ByRef refCorrelations() As ZtStoryCorrelation, ByVal valCorrelationsCt As Integer)
  
  Dim locRange As Word.Range
  
  Set locRange = pvtField.Result
  Set pvtBookmark = pvtStory.Document.AddBookmark(pvtConfig.Basic.Macro.CitationBookmarkPrefix & _
                                                    ZtSubprocedures.AddLeadingZeros(pvtNr + 1 + pvtOffset, _
                                                                                    ZtSubprocedures.DigitsCt(pvtStory.Document.CitationGroupsCt)), _
                                                  pvtFieldProcedures.GetCompleteRange)
  pvtSetCaption refCaptions, locRange.Start, refCorrelations, valCorrelationsCt, locRange.StoryLength

End Sub

Friend Function SetInternalLinking() As ZtFMessageType

  Dim locRange As Word.Range
  Dim locMatches As Collection
  Dim locMatch As ZtRegMatch
  Dim locCitationStart As Long
  Dim locCitationEnd As Long
  Dim locCitationRange As Word.Range
  Dim locReferenceNr As Integer
  Dim locReferenceAfterLastNr As Integer
  Dim locReferenceCtr As Integer
  Dim locCitationTexts() As ZtCitationText
  Dim locCitationTextsCt As Integer
  Dim locCitationCtr As Integer
  Dim locResult As ZtFMessageType
  
  Set locRange = pvtField.Result
  Set locCitationRange = locRange.Duplicate
  Set locMatches = pvtConfig.Final.Citation.ResultRegExp.Execute(locRange.Text)
  
  ' Delete hidden character formats.
  pvtFieldProcedures.RemoveCodeFormat
  
  If locMatches.Count > 0 Then
    For Each locMatch In locMatches
      With locMatch
        If pvtConfig.User.Macro.Debugging Then
          locResult = pvtConfig.Final.Citation.ResultRegExp.Debugging(locMatch, locRange)
          If locResult = MessageCancel Then
            Exit For
          End If
        Else
          locResult = MessageOk
        End If
        locCitationStart = locRange.Start + .FirstIndex
        locCitationEnd = locCitationStart + Len(.Groups("numbering"))
        locReferenceNr = CInt(.Groups("number")) - 1
        If locReferenceAfterLastNr = 0 Then
          locReferenceAfterLastNr = locReferenceNr
        End If
        
        locCitationRange.SetRange locCitationStart, locCitationEnd
        
        ' Keep citations.
        For locReferenceCtr = locReferenceAfterLastNr To locReferenceNr
          If locReferenceCtr = locReferenceNr Then
            pvtCitationTextFactory locCitationTexts, locCitationTextsCt, locReferenceCtr, locCitationCtr = 0, locCitationRange
          Else
            pvtCitationTextFactory locCitationTexts, locCitationTextsCt, locReferenceCtr, False
          End If
        Next
        
        ' Prepare for range citation.
        If Len(.Groups("rangeDelimiter")) > 0 Then
          locReferenceAfterLastNr = locReferenceNr + 1
        Else
          locReferenceAfterLastNr = 0
        End If
        locCitationCtr = locCitationCtr + 1
      End With
    Next
  
    ' Set linking.
    If locResult = MessageOk Then
      For locCitationCtr = 0 To locCitationTextsCt - 1
        locCitationTexts(locCitationCtr).SetInternalLinking
      Next
    End If
    
  Else
    locResult = pvtMessageDisplay.Show("The citation group and citation styles you set in ZtConfig.xml seem not to apply to the selected citation group." & vbNewLine & vbNewLine & _
                                         "You may procede, otherwise cancel.", _
                                       MessageCritical + MessageOkCancel, _
                                       True, _
                                       locRange)
  End If
  
  SetInternalLinking = locResult

End Function

Friend Function JoinWithPrevious(ByVal valJoinConditionally As Boolean, ByRef refAffectedCt As Integer, ByRef refNotAffectedCt As Integer) As String

  Dim locRange As Word.Range
  Dim locToJoin As Boolean
  
  Set locRange = pvtField.Result
  locRange.SetRange pvtPreviousCitationGroup.ResultRange.End + 1, pvtField.Code.Start - 1
  
  If Not valJoinConditionally Then
    locToJoin = True
  ElseIf Len(pvtConfig.Final.CitationGroup.BetweenDeleteRegExp.Replace(locRange.Text, vbNullString)) = 0 Then
    locToJoin = True
  End If
  
  If locToJoin Then
    locRange.SetRange pvtPreviousCitationGroup.CodeRange.Start - 1, pvtField.Result.End + 1
    JoinWithPrevious = pvtJoinWithPrevious
    refAffectedCt = refAffectedCt + 1
  Else
    refNotAffectedCt = refNotAffectedCt + 1
  End If
  
End Function

Friend Sub Resolve()

  Dim locCode As String
  Dim locCitationGroupJson As String
  Dim locCitationGroup As Scripting.Dictionary
  Dim locCitationDicts As Collection
  Dim locCitationDict As Scripting.Dictionary
  Dim locConverter As ZtJsonConverter
  Dim locString As ZtStringBuilder
  Dim locCitationInfoCtr As Integer
  Dim locFieldRange As Word.Range
  
  Set locConverter = New ZtJsonConverter
  locCode = Trim$(pvtField.Code)
  locCitationGroupJson = Trim$(Right$(locCode, Len(locCode) - Len(pvtConfig.Basic.Zotero.CitationGroupIdentifier)))
  
  Set locCitationGroup = locConverter.Parse(locCitationGroupJson)
  Set locCitationDicts = locCitationGroup.Item("citationItems")
  For Each locCitationDict In locCitationDicts
    pvtCitationInfoFactory locCitationDict.Item("itemData")
  Next
  
  Set locString = New ZtStringBuilder
  With locString
    .Append pvtConfig.User.ResolveCitationStyle.GroupPrefix
    For locCitationInfoCtr = 0 To pvtCitationInfosCt - 1
      .Append pvtConfig.User.ResolveCitationStyle.Prefix
      .Append pvtCitationInfos(locCitationInfoCtr).ToString
      .Append pvtConfig.User.ResolveCitationStyle.Suffix
      .Append pvtConfig.User.ResolveCitationStyle.Delimiter
    Next
    .RemoveFromEnd Len(pvtConfig.User.ResolveCitationStyle.Delimiter)
    .Append pvtConfig.User.ResolveCitationStyle.GroupSuffix
    Set locFieldRange = pvtField.Result
    pvtField.Delete
    locFieldRange.Text = .ToString
  End With
  
End Sub

Friend Property Get Story() As ZtIStory

  Set Story = pvtStory
  
End Property
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Private procedures and properties.
Private Sub pvtSetCaption(ByRef refCaptions() As ZtCaption, ByVal valStart As Double, _
  Optional ByRef refCorrelations As Variant, Optional ByVal valCorrelationsCt As Integer = 0, Optional ByVal valStoryLength As Long = 0)
  
  Dim locCaptionCtr As Integer
  Dim locCorrelation As ZtStoryCorrelation
  Dim locCorrelationCtr As Integer
  
  If IsMissing(refCorrelations) Or valCorrelationsCt = 0 Then
    pvtCorrectedStart = valStart
    For locCaptionCtr = UBound(refCaptions) To 0 Step -1
      If CLng(valStart) >= refCaptions(locCaptionCtr).Start Then
        Set pvtCaption = refCaptions(locCaptionCtr)
        
        Exit Sub
        
      End If
    Next
  Else
    For locCorrelationCtr = valCorrelationsCt - 1 To 0 Step -1
      Set locCorrelation = refCorrelations(locCorrelationCtr)
      If valStart >= locCorrelation.StartInSubstory Then
        pvtSetCaption refCaptions, locCorrelation.StartInMainStory + valStart / valStoryLength
        
        Exit Sub
        
      End If
    Next
  End If

End Sub

Private Sub pvtCitationInfoFactory(ByVal valData As Scripting.Dictionary)

  ReDim Preserve pvtCitationInfos(pvtCitationInfosCt)
  Set pvtCitationInfos(pvtCitationInfosCt) = New ZtCitationInfo
  pvtCitationInfos(pvtCitationInfosCt).Initialize pvtConfig, valData
  pvtCitationInfosCt = pvtCitationInfosCt + 1

End Sub

Private Sub pvtCitationTextFactory(ByRef refCitationTexts() As ZtCitationText, ByRef refCitationTextsCt As Integer, _
   ByVal valReferenceNr As Integer, ByVal valIsFirstInGroup As Boolean, Optional ByVal valCitationRange As Word.Range = Nothing)

  ReDim Preserve refCitationTexts(refCitationTextsCt)
  Set refCitationTexts(refCitationTextsCt) = New ZtCitationText
  refCitationTexts(refCitationTextsCt).Initialize pvtConfig, Me, valReferenceNr, valIsFirstInGroup, valCitationRange
  refCitationTextsCt = refCitationTextsCt + 1
  
End Sub

Private Function pvtJoinWithPrevious() As String

  Dim locBetweenRange As Word.Range
  Dim locNewCode As String
  Dim locThisCodeStart As Long
  Dim locPreviouseCodeEnd As Long
  
  ' Determine the cut text between two citation groups which will be the result of this function.
  Set locBetweenRange = pvtPreviousCitationGroup.ResultRange
  locBetweenRange.SetRange pvtPreviousCitationGroup.ResultRange.End + 1, pvtField.Code.Start - 1
  
  ' Determine the code for the combination of the two citation groups.
  locPreviouseCodeEnd = InStr(pvtPreviousCitationGroup.CodeRange.Text, pvtConfig.Basic.Zotero.CitationGroupSuffixIdentifier)
  locNewCode = Left$(pvtPreviousCitationGroup.CodeRange.Text, locPreviouseCodeEnd - 1) & pvtConfig.Basic.Zotero.CitationGroupDelimiter
  locThisCodeStart = InStr(pvtField.Code.Text, _
                           pvtConfig.Basic.Zotero.CitationGroupPrefixIdentifier) + _
                     Len(pvtConfig.Basic.Zotero.CitationGroupPrefixIdentifier)
  locNewCode = locNewCode & Right$(pvtField.Code, Len(pvtField.Code.Text) - locThisCodeStart + 1)
  
  ' Insert the new code into previouse field and delete the field of this citation group (Me).
  pvtPreviousCitationGroup.CodeRange.Text = locNewCode
  pvtField.Delete
  
  ' Return.
  pvtJoinWithPrevious = locBetweenRange.Text
  
End Function
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
