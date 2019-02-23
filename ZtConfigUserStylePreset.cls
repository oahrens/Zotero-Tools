VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ZtConfigUserStylePreset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Class ZtConfigUserBackwardLinking.
' It preserves a part of the user configuration informations and methods to read them from ZtConfig.xml.
' It is not part of ZtConfigUser because its informations are read only once.
'
' Zotero Tools.
' This software is under Revised ('New') BSD license.
' Copyright © 2019, Olaf Ahrens. All rights reserved.
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Private variables.
Private pvtName As String
Private pvtCitationGroupStyle As String
Private pvtPunctuationStyle As String
Private pvtCitationStyle As String
Private pvtResolveCitationStyle As String
Private pvtReferenceStyle As String
Private pvtBackwardLinkingStyle As String
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Friend/Public procedures and properties.
Friend Sub Initialize(ByVal valConfig As ZtConfig, ByVal valNode As MSXML2.IXMLDOMNode)
      
  pvtCitationGroupStyle = valConfig.ResolveTextElement(valNode.SelectSingleNode("citationGroupStyle"))
  pvtPunctuationStyle = valConfig.ResolveTextElement(valNode.SelectSingleNode("punctuationStyle"))
  pvtCitationStyle = valConfig.ResolveTextElement(valNode.SelectSingleNode("citationStyle"))
  pvtResolveCitationStyle = valConfig.ResolveTextElement(valNode.SelectSingleNode("resolveCitationStyle"))
  pvtReferenceStyle = valConfig.ResolveTextElement(valNode.SelectSingleNode("referenceStyle"))
  pvtBackwardLinkingStyle = valConfig.ResolveTextElement(valNode.SelectSingleNode("backwardLinkingStyle"))
  pvtName = valConfig.ResolveTextElement(valNode.Attributes.getNamedItem("name"))

End Sub

Friend Property Get Name() As String

  Name = pvtName
  
End Property

Friend Property Get CitationGroupStyle() As String

  CitationGroupStyle = pvtCitationGroupStyle
  
End Property

Friend Property Get PunctuationStyle() As String

  PunctuationStyle = pvtPunctuationStyle
  
End Property

Friend Property Get CitationStyle() As String

  CitationStyle = pvtCitationStyle
  
End Property

Friend Property Get ResolveCitationStyle() As String

  ResolveCitationStyle = pvtResolveCitationStyle
  
End Property

Friend Property Get ReferenceStyle() As String

  ReferenceStyle = pvtReferenceStyle
  
End Property

Friend Property Get BackwardLinkingStyle() As String

  BackwardLinkingStyle = pvtBackwardLinkingStyle
  
End Property
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

