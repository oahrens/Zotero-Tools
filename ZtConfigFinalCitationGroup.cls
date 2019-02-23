VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ZtConfigFinalCitationGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Class ZtConfigFinalCitationGroup.
' It preserves a part of the final configuration informations and methods to read them from ZtConfig.xml.
'
' Zotero Tools.
' This software is under Revised ('New') BSD license.
' Copyright © 2019, Olaf Ahrens. All rights reserved.
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Private variables.
Private pvtBetweenDeleteRegExp As ZtRegExp
Private pvtBetweenDeleteMultipleSpaceRegExp As ZtRegExp
Private pvtBetweenDeleteAffixedSpaceRegExp As ZtRegExp
Private pvtPunctuationRegExp() As ZtRegExp
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Friend procedures and properties.
Friend Sub Initialize(ByVal valConfig As ZtConfig, ByVal valXml As MSXML2.DOMDocument60)

  Dim locNode As MSXML2.IXMLDOMNode
  
  ReDim pvtPunctuationRegExp(3)
  Set locNode = valXml.SelectSingleNode("/zoteroTools/final/citationGroup")
  With locNode
    valConfig.RegExpFactory pvtBetweenDeleteRegExp, _
                            "delete between citation groups", _
                            valConfig.ResolveTextElement(.SelectSingleNode("betweenDeleteRegString"), True), _
                            True
    valConfig.RegExpFactory pvtBetweenDeleteMultipleSpaceRegExp, _
                            "delete multiple spaces between citation groups", _
                            valConfig.ResolveTextElement(.SelectSingleNode("betweenDeleteMultipleSpaceRegString"), True), _
                            True
    valConfig.RegExpFactory pvtBetweenDeleteAffixedSpaceRegExp, _
                            "delete affixed spaces between citation groups", _
                            valConfig.ResolveTextElement(.SelectSingleNode("betweenDeleteAffixedSpaceRegString"), True), _
                            True
    valConfig.RegExpFactory pvtPunctuationRegExp(SignsBeforeAndAfter), _
                            "punctuation", _
                            valConfig.ResolveTextElement(.SelectSingleNode("punctuationRegString[@surroundingSigns='" & CStr(SignsBeforeAndAfter) & "']"), True), _
                            True
    valConfig.RegExpFactory pvtPunctuationRegExp(SignBefore), _
                            "punctuation", _
                            valConfig.ResolveTextElement(.SelectSingleNode("punctuationRegString[@surroundingSigns='" & CStr(SignBefore) & "']"), True), _
                            True
    valConfig.RegExpFactory pvtPunctuationRegExp(SignAfter), _
                            "punctuation", _
                            valConfig.ResolveTextElement(.SelectSingleNode("punctuationRegString[@surroundingSigns='" & CStr(SignAfter) & "']"), True), _
                            True
    valConfig.RegExpFactory pvtPunctuationRegExp(NoSigns), _
                            "punctuation", _
                            valConfig.ResolveTextElement(.SelectSingleNode("punctuationRegString[@surroundingSigns='" & CStr(NoSigns) & "']"), True), _
                            True
  End With

End Sub

Friend Property Get BetweenDeleteRegExp() As ZtRegExp

  Set BetweenDeleteRegExp = pvtBetweenDeleteRegExp
  
End Property

Friend Property Get BetweenDeleteMultipleSpaceRegExp() As ZtRegExp

  Set BetweenDeleteMultipleSpaceRegExp = pvtBetweenDeleteMultipleSpaceRegExp
  
End Property

Friend Property Get BetweenDeleteAffixedSpaceRegExp() As ZtRegExp

  Set BetweenDeleteAffixedSpaceRegExp = pvtBetweenDeleteAffixedSpaceRegExp
  
End Property

Friend Function PunctuationRegExp(ByVal valSurroundingSigns As ZtFSurroundingSigns) As ZtRegExp

  Set PunctuationRegExp = pvtPunctuationRegExp(valSurroundingSigns)
  
End Function
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

