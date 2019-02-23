VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ZtConfigBasicZotero"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Class ZtConfigBasicZotero.
' It preserves a part of the basic configuration informations and methods to read them from ZtConfig.xml.
'
' Zotero Tools.
' This software is under Revised ('New') BSD license.
' Copyright © 2019, Olaf Ahrens. All rights reserved.
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Private variables.
Private pvtCitationGroupIdentifier As String
Private pvtCitationGroupPrefixIdentifier As String
Private pvtCitationGroupSuffixIdentifier As String
Private pvtCitationGroupDelimiter As String
Private pvtBibliographyIdentifier As String
Private pvtCanHoldCitationStoryTypes() As WdStoryType
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Friend/Public procedures and properties.
' All members that should be callable by CallByName procedure must be public.
Friend Sub Initialize(ByVal valConfig As ZtConfig, ByVal valXml As MSXML2.DOMDocument60)

  Dim locNode As MSXML2.IXMLDOMNode
  
  Set locNode = valXml.SelectSingleNode("/zoteroTools/basic/zotero")
  With locNode
    pvtCitationGroupIdentifier = valConfig.ResolveTextElement(.SelectSingleNode("citationGroupIdentifier"))
    pvtCitationGroupPrefixIdentifier = valConfig.ResolveTextElement(.SelectSingleNode("citationGroupPrefixIdentifier"))
    pvtCitationGroupSuffixIdentifier = valConfig.ResolveTextElement(.SelectSingleNode("citationGroupSuffixIdentifier"))
    pvtCitationGroupDelimiter = valConfig.ResolveTextElement(.SelectSingleNode("citationGroupDelimiter"))
    pvtBibliographyIdentifier = valConfig.ResolveTextElement(.SelectSingleNode("bibliographyIdentifier"))
    pvtCanHoldCitationStoryTypes = valConfig.GetWordStoryTypes(Split(valConfig.ResolveTextElement(.SelectSingleNode("canHoldCitationStoryTypes")), ","))
  End With

End Sub

Public Property Get CitationGroupIdentifier() As String

  CitationGroupIdentifier = pvtCitationGroupIdentifier
  
End Property

Public Property Get CitationGroupPrefixIdentifier() As String

  CitationGroupPrefixIdentifier = pvtCitationGroupPrefixIdentifier
  
End Property

Public Property Get CitationGroupSuffixIdentifier() As String

  CitationGroupSuffixIdentifier = pvtCitationGroupSuffixIdentifier
  
End Property

Public Property Get CitationGroupDelimiter() As String

  CitationGroupDelimiter = pvtCitationGroupDelimiter
  
End Property

Public Property Get BibliographyIdentifier() As String

  BibliographyIdentifier = pvtBibliographyIdentifier
  
End Property

Public Property Get CanHoldCitationStoryTypes() As WdStoryType()

  CanHoldCitationStoryTypes = pvtCanHoldCitationStoryTypes
  
End Property
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

