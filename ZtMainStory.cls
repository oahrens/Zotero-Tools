VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ZtMainStory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements ZtIStory

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Class ZtMainStory.
' It extends ZtStory for captions and bibliography.
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Private variables.
Private pvtConfig As ZtConfig
Private pvtStartForm As ZtStartForm
Private pvtDocument As ZtDocument
Private pvtRange As Word.Range
Private pvtStory As ZtStory
Private pvtBibliography As ZtBibliography
Private pvtCaptions() As ZtCaption
Private pvtCaptionsCt As Integer
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Constructor.
Private Sub Class_Initialize()

  Set pvtStory = New ZtStory
  Set pvtBibliography = New ZtBibliography
  pvtCaptionsCt = -1
  
End Sub
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Private interface procedures and properties directing to Friend procedures and properties below.
Private Sub ZtIStory_Initialize(ByVal valConfig As ZtConfig, ByVal valStartForm As ZtStartForm, ByVal valDocument As ZtDocument, ByVal valRange As Word.Range)
  
  Me.Initialize valConfig, valStartForm, valDocument, valRange
  
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
Friend Sub Initialize(ByVal valConfig As ZtConfig, ByVal valStartForm As ZtStartForm, ByVal valDocument As ZtDocument, ByVal valRange As Word.Range)
  
  Set pvtConfig = valConfig
  Set pvtStartForm = valStartForm
  Set pvtDocument = valDocument
  Set pvtRange = valRange
  pvtStory.Initialize valConfig, valStartForm, valDocument, valRange
  
End Sub

Friend Property Get CanHoldCitations() As Boolean

  CanHoldCitations = pvtStory.CanHoldCitations
  
End Property

Friend Sub KeepCitationGroups(ByRef refCitationGroupsOffset As Integer)

  pvtStory.KeepCitationGroups refCitationGroupsOffset
  
End Sub

Friend Function RemoveInternalLinking() As ZtFMessageType

  Dim locResult As ZtFMessageType
  
  locResult = pvtStory.RemoveInternalLinking()
  If locResult = MessageOk Then
    locResult = pvtBibliography.Initialize(pvtConfig, pvtStartForm, Me)
    If locResult = MessageOk Then
      ' delete 'Cited on page n' text
      locResult = pvtBibliography.RemoveCitedOnPage
    End If
  End If
  
  RemoveInternalLinking = locResult

End Function

Friend Function SetInternalLinking() As ZtFMessageType

  SetInternalLinking = pvtStory.SetInternalLinking()

End Function

Friend Function AdjustPunctuation(ByRef refAffectedCt As Integer, ByRef refNotAffectedCt As Integer, ByRef refNotCorrectedCt As Integer) As ZtFMessageType

  AdjustPunctuation = pvtStory.AdjustPunctuation(refAffectedCt, refNotAffectedCt, refNotCorrectedCt)
  
End Function

Friend Sub ResolveCitationGroups(ByRef refAffectedCt As Integer)

  pvtStory.ResolveCitationGroups refAffectedCt

End Sub

Friend Sub UpdateAllFields()

  Dim locToc As Word.TableOfContents
  Dim locTof As Word.TableOfFigures
  
  '' Update tables. We do this first so that they contain all necessary
  '' entries and so extend to their final number of pages.
  For Each locToc In pvtDocument.TablesOfContents
    locToc.Update
  Next
  
  For Each locTof In pvtDocument.TablesOfFigures
    locTof.Update
  Next

  pvtStory.UpdateAllFields

End Sub

Friend Property Get Range() As Word.Range

  Set Range = pvtRange
  
End Property

Friend Property Get Document() As ZtDocument

  Set Document = pvtDocument
  
End Property

Friend Sub JoinCitationGroups(ByRef refAffectedCt As Integer, ByRef refNotAffectedCt As Integer, Optional ByVal valRange As Word.Range = Nothing)

  pvtStory.JoinCitationGroups refAffectedCt, refNotAffectedCt, valRange
  
End Sub
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Extending ZtStory procedures and properties.
Friend Property Get Bibliography() As ZtBibliography
  
  Set Bibliography = pvtBibliography
  
End Property

Friend Function SetBibliographyBookmarks() As ZtFMessageType

  Dim locResult As ZtFMessageType
  
  pvtStartForm.Progress.SetStep 1
  locResult = pvtBibliography.Initialize(pvtConfig, pvtStartForm, Me)
  pvtStartForm.Progress.SetSubstep
  If locResult = MessageOk Then
    locResult = pvtBibliography.SetBookmarks
  End If
  
  SetBibliographyBookmarks = locResult
  
End Function

Friend Function SetBibliographyLinking() As ZtFMessageType

  SetBibliographyLinking = pvtBibliography.SetInternalLinking
  
End Function

Friend Function SetBibliographyWebLinks() As ZtFMessageType

  Dim locResult As ZtFMessageType
  
  locResult = pvtBibliography.Initialize(pvtConfig, pvtStartForm, Me)
  If locResult = MessageOk Then
    locResult = pvtBibliography.SetWebLinks
  End If
  
  SetBibliographyWebLinks = locResult
  
End Function

Friend Sub GetCaptions(ByRef refCaptions() As ZtCaption)

  refCaptions = pvtCaptions
  
End Sub

Friend Sub KeepCaptions()

  Dim locCaptionRange As Word.Range
  Dim locPreviousCaptionStart As Long
  
  If pvtCaptionsCt = -1 Then
    pvtStartForm.Progress.SetToRunningMode
    pvtStartForm.Progress.SetStep
    pvtCaptionsCt = 0
    ' Null-caption for all citations before first caption (in the introduction etc).
    pvtCaptionFactory "?", "?", 0
  
    ' Real captions.
    Set locCaptionRange = pvtDocument.GoingTo(wdGoToHeading, wdGoToFirst)
    Do Until locPreviousCaptionStart = locCaptionRange.Start
      With locCaptionRange
        If .Paragraphs(1).Range.ListParagraphs.Count = 1 Then
          If .ListFormat.ListLevelNumber <= pvtConfig.User.BackwardLinkingStyle.MaxCaptionLevel Then
            pvtCaptionFactory .ListFormat.ListString, .Paragraphs(1).Range.Text, .Start
          End If
        End If
        locPreviousCaptionStart = .Start
        pvtStartForm.Progress.SetSubstep
        Set locCaptionRange = .GoToNext(wdGoToHeading)
      End With
    Loop
    pvtStartForm.Progress.SetToNormalMode
    If pvtCaptionsCt = 1 Then
      pvtStartForm.MessageDisplay.Show "I couldn't find any caption." & vbNewLine & vbNewLine & _
                                         "Please, ensure that you have used Word's build-in formatting for heading levels 1 to 9 in your document.", _
                                       MessageExclamation + MessageCancel
    End If
  End If
  
End Sub
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Private procedures and properties.
Private Sub pvtCaptionFactory(ByVal valNr As String, ByVal valText As String, ByVal valStart As Long)

  ReDim Preserve pvtCaptions(pvtCaptionsCt)
  Set pvtCaptions(pvtCaptionsCt) = New ZtCaption
  pvtCaptions(pvtCaptionsCt).Initialize pvtConfig, valNr, valText, valStart
  pvtCaptionsCt = pvtCaptionsCt + 1

End Sub
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *



