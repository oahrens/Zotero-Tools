VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ZtConfigBasicMacro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Class ZtConfigBasicMacro.
' It preserves a part of the basic configuration informations and methods to read them from ZtConfig.xml.
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Private variables.
Private pvtCitationBookmarkPrefix As String
Private pvtReferenceBookmarkPrefix As String
Private pvtBackwardLinkingPageNrMarker As String
Private pvtBackwardLinkingCaptionNrMarker As String
Private pvtBackwardLinkingCaptionTextMarker As String
Private pvtDoiResolver As String
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Friend/Public procedures and properties.
' All members that should be callable by CallByName procedure must be public.
Friend Sub Initialize(ByVal valConfig As ZtConfig, ByVal valXml As MSXML2.DOMDocument60)

  Dim locNode As MSXML2.IXMLDOMNode
  
  Set locNode = valXml.SelectSingleNode("/zoteroTools/basic/macro")
  With locNode
    pvtCitationBookmarkPrefix = valConfig.ResolveTextElement(.SelectSingleNode("citationBookmarkPrefix"))
    pvtReferenceBookmarkPrefix = valConfig.ResolveTextElement(.SelectSingleNode("referenceBookmarkPrefix"))
    pvtBackwardLinkingPageNrMarker = valConfig.ResolveTextElement(.SelectSingleNode("backwardLinkingPageNrMarker"))
    pvtBackwardLinkingCaptionNrMarker = valConfig.ResolveTextElement(.SelectSingleNode("backwardLinkingCaptionNrMarker"))
    pvtBackwardLinkingCaptionTextMarker = valConfig.ResolveTextElement(.SelectSingleNode("backwardLinkingCaptionTextMarker"))
    pvtDoiResolver = valConfig.ResolveTextElement(.SelectSingleNode("doiResolver"))
  End With

End Sub

Public Property Get CitationBookmarkPrefix() As String

  CitationBookmarkPrefix = pvtCitationBookmarkPrefix
  
End Property

Public Property Get ReferenceBookmarkPrefix() As String

  ReferenceBookmarkPrefix = pvtReferenceBookmarkPrefix
  
End Property

Public Property Get BackwardLinkingPageNrMarker() As String

  BackwardLinkingPageNrMarker = pvtBackwardLinkingPageNrMarker
  
End Property

Public Property Get BackwardLinkingCaptionNrMarker() As String

  BackwardLinkingCaptionNrMarker = pvtBackwardLinkingCaptionNrMarker
  
End Property

Public Property Get BackwardLinkingCaptionTextMarker() As String

  BackwardLinkingCaptionTextMarker = pvtBackwardLinkingCaptionTextMarker
  
End Property

Public Property Get DoiResolver() As String

  DoiResolver = pvtDoiResolver
  
End Property
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

