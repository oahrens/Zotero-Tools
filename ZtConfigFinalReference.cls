VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ZtConfigFinalReference"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Class ZtConfigFinalReference.
' It preserves a part of the final configuration informations and methods to read them from ZtConfig.xml.
'
' Zotero Tools.
' This software is under Revised ('New') BSD license.
' Copyright © 2019, Olaf Ahrens. All rights reserved.
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Private variables.
Private pvtReferenceUniversalSuffixCharacters As String
Private pvtReferenceUniversalSuffixChoiceRegString As String
Private pvtWholeRegExp As ZtRegExp
Private pvtWebLinkRegExp As ZtRegExp
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Friend/Public procedures and properties.
' All members that should be callable by CallByName procedure must be public.
Friend Sub Initialize(ByVal valConfig As ZtConfig, ByVal valXml As MSXML2.DOMDocument60)

  Dim locNode As MSXML2.IXMLDOMNode
  Dim locReferenceUniversalSuffixChoiceRegString As ZtStringBuilder
  Dim locReferenceUniversalSuffixCharacters As ZtStringBuilder

  Set locReferenceUniversalSuffixCharacters = New ZtStringBuilder
  Set locReferenceUniversalSuffixChoiceRegString = New ZtStringBuilder
  If Len(Trim$(valConfig.User.ReferenceStyle.ReferenceNormalSuffix)) > 0 Then
    locReferenceUniversalSuffixCharacters.Append Trim$(valConfig.User.ReferenceStyle.ReferenceNormalSuffix)
    locReferenceUniversalSuffixChoiceRegString.Append Trim$(valConfig.User.ReferenceStyle.ReferenceNormalSuffixRegPiece.Value)
    locReferenceUniversalSuffixChoiceRegString.Append "|"
  End If
  If Len(Trim$(valConfig.User.ReferenceStyle.ReferenceHyperlinkAtEndSuffix)) > 0 Then
    If Trim$(valConfig.User.ReferenceStyle.ReferenceHyperlinkAtEndSuffix) <> Trim$(valConfig.User.ReferenceStyle.ReferenceNormalSuffix) Then
      locReferenceUniversalSuffixCharacters.Append Trim$(valConfig.User.ReferenceStyle.ReferenceHyperlinkAtEndSuffix)
      locReferenceUniversalSuffixChoiceRegString.Append Trim$(valConfig.User.ReferenceStyle.ReferenceHyperlinkAtEndSuffixRegPiece.Value)
      locReferenceUniversalSuffixChoiceRegString.Append "|"
    End If
  End If
  If Len(Trim$(valConfig.User.ReferenceStyle.ReferenceAccessedAtEndSuffix)) > 0 Then
    If Trim$(valConfig.User.ReferenceStyle.ReferenceAccessedAtEndSuffix) <> Trim$(valConfig.User.ReferenceStyle.ReferenceNormalSuffix) Then
      If Trim$(valConfig.User.ReferenceStyle.ReferenceAccessedAtEndSuffix) <> Trim$(valConfig.User.ReferenceStyle.ReferenceHyperlinkAtEndSuffix) Then
        locReferenceUniversalSuffixCharacters.Append Trim$(valConfig.User.ReferenceStyle.ReferenceAccessedAtEndSuffix)
        locReferenceUniversalSuffixChoiceRegString.Append Trim$(valConfig.User.ReferenceStyle.ReferenceAccessedAtEndSuffixRegPiece.Value)
        locReferenceUniversalSuffixChoiceRegString.Append "|"
      End If
    End If
  End If
  If Len(Trim$(valConfig.User.BackwardLinkingStyle.DelimiterBefore)) > 0 Then
    If Trim$(valConfig.User.BackwardLinkingStyle.DelimiterBefore) <> Trim$(valConfig.User.ReferenceStyle.ReferenceNormalSuffix) Then
      If Trim$(valConfig.User.BackwardLinkingStyle.DelimiterBefore) <> Trim$(valConfig.User.ReferenceStyle.ReferenceHyperlinkAtEndSuffix) Then
        If Trim$(valConfig.User.BackwardLinkingStyle.DelimiterBefore) <> Trim$(valConfig.User.ReferenceStyle.ReferenceAccessedAtEndSuffix) Then
          locReferenceUniversalSuffixCharacters.Append Trim$(valConfig.User.BackwardLinkingStyle.DelimiterBefore)
          locReferenceUniversalSuffixChoiceRegString.Append Trim$(valConfig.User.BackwardLinkingStyle.DelimiterBeforeRegPiece.Value)
          locReferenceUniversalSuffixChoiceRegString.Append "|"
        End If
      End If
    End If
  End If
  pvtReferenceUniversalSuffixCharacters = locReferenceUniversalSuffixCharacters.ToString
  locReferenceUniversalSuffixChoiceRegString.RemoveFromEnd 1
  pvtReferenceUniversalSuffixChoiceRegString = locReferenceUniversalSuffixChoiceRegString.ToString
  
  Set locNode = valXml.SelectSingleNode("/zoteroTools/final/reference")
  With locNode
    valConfig.RegExpFactory pvtWholeRegExp, _
                            "whole reference", _
                            valConfig.ResolveTextElement(.SelectSingleNode("wholeRegString"), True), _
                            True
    valConfig.RegExpFactory pvtWebLinkRegExp, _
                            "reference web link", _
                            valConfig.ResolveTextElement(.SelectSingleNode("webLinkRegString"), True), _
                            True
  End With

End Sub

Friend Property Get ReferenceUniversalSuffixCharacters() As String

  ReferenceUniversalSuffixCharacters = pvtReferenceUniversalSuffixCharacters
  
End Property

Public Property Get ReferenceUniversalSuffixChoiceRegString() As String

  ReferenceUniversalSuffixChoiceRegString = pvtReferenceUniversalSuffixChoiceRegString
  
End Property

Friend Property Get WholeRegExp() As ZtRegExp

  Set WholeRegExp = pvtWholeRegExp
  
End Property

Friend Property Get WebLinkRegExp() As ZtRegExp

  Set WebLinkRegExp = pvtWebLinkRegExp
  
End Property
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

