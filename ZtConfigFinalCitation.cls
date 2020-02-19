VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ZtConfigFinalCitation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Class ZtConfigFinalCitation.
' It preserves a part of the final configuration informations and methods to read them from ZtConfig.xml.
'
' Zotero Tools.
' This software is under Revised ('New') BSD license.
' Copyright © 2019, Olaf Ahrens. All rights reserved.
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Private variables.
Private pvtResultRegExp As ZtRegExp
Private pvtHtmlTagsRegExp As ZtRegExp
Private pvtGivenNameRegExp As ZtRegExp
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Friend procedures and properties.
Friend Sub Initialize(ByVal valConfig As ZtConfig, ByVal valXml As MSXML2.DOMDocument60)

  Dim locNode As MSXML2.IXMLDOMNode
  
  Set locNode = valXml.SelectSingleNode("/zoteroTools/final/citation")
  With locNode
    If valConfig.User.CitationStyle.SuffixIsBeforeLocalizer Then
      valConfig.RegExpFactory pvtResultRegExp, _
                              "citation", _
                              valConfig.ResolveTextElement(.SelectSingleNode("resultRegString[@suffixIsBeforeLocalizer='true']"), True), _
                              True
    Else
      valConfig.RegExpFactory pvtResultRegExp, _
                              "citation", _
                              valConfig.ResolveTextElement(.SelectSingleNode("resultRegString[@suffixIsBeforeLocalizer='false']"), True), _
                              True
    End If
    valConfig.RegExpFactory pvtHtmlTagsRegExp, _
                            "html tags", _
                            valConfig.ResolveTextElement(.SelectSingleNode("htmlTagsRegString"), True), _
                            True
    valConfig.RegExpFactory pvtGivenNameRegExp, _
                            "given name", _
                            valConfig.ResolveTextElement(.SelectSingleNode("givenNameRegString"), True), _
                            True
  End With

End Sub

Friend Property Get ResultRegExp() As ZtRegExp

  Set ResultRegExp = pvtResultRegExp
  
End Property

Friend Property Get HtmlTagsRegExp() As ZtRegExp

  Set HtmlTagsRegExp = pvtHtmlTagsRegExp
  
End Property

Friend Property Get GivenNameRegExp() As ZtRegExp

  Set GivenNameRegExp = pvtGivenNameRegExp
  
End Property
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

