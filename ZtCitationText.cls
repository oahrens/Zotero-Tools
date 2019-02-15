VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ZtCitationText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Class ZtCitationText.
' It preserves range information of a single citation, i.e. '1', '2', '3', '7' in '[1-3, 7]'.
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Private variables.
Private pvtConfig As ZtConfig
Private pvtCitationGroup As ZtCitationGroup
Private pvtRange As Word.Range
Private pvtReferenceNr As Integer
Private pvtIsFirstInGroup As Boolean
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Friend procedures and properties.
Friend Sub Initialize(ByVal valConfig As ZtConfig, ByVal valCitationGroup As ZtCitationGroup, ByVal valReferenceNr As Integer, _
  ByVal valIsFirstInGroup As Boolean, Optional ByVal valRange As Word.Range = Nothing)

  Set pvtConfig = valConfig
  Set pvtCitationGroup = valCitationGroup
  pvtReferenceNr = valReferenceNr
  pvtIsFirstInGroup = valIsFirstInGroup
  If Not valRange Is Nothing Then
    Set pvtRange = valRange.Duplicate
  End If
  
End Sub

Friend Sub SetInternalLinking()

  Dim locReference As ZtReference
  Dim locFormat As ZtCharacterFormat
  
  Set locReference = pvtCitationGroup.Story.Document.MainStory.Bibliography.References(pvtReferenceNr)

  ' Set hyperlink citation -> reference only if the citation isn't part of a citation range.
  If Not pvtRange Is Nothing Then
    If pvtIsFirstInGroup And pvtConfig.User.Macro.CitationInsertZeroWidthSpace And Len(pvtConfig.User.CitationGroupStyle.Prefix) = 0 And Len(pvtRange.Text) = 1 Then
      Set locFormat = New ZtCharacterFormat
      locFormat.SetFormat pvtRange
      pvtRange.InsertBefore pvtConfig.Basic.Characters.ZeroWidthSpace
      locFormat.GetFormat pvtRange
    End If
    pvtCitationGroup.Story.Document.AddInternalLink pvtRange, _
                                                    pvtConfig.Basic.Macro.ReferenceBookmarkPrefix & _
                                                      ZtSubprocedures.AddLeadingZeros(pvtReferenceNr + 1, _
                                                                                      ZtSubprocedures.DigitsCt(pvtCitationGroup.Story.Document.CitationGroupsCt)), _
                                                    locReference.ScreenTipText
  End If
  
  ' Save the link for backward linking, if necessary.
  If pvtConfig.User.Macro.WithBackwardLinking Then
    locReference.AddCitationGroup pvtCitationGroup
  End If

End Sub
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *




