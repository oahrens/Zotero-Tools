VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ZtConfigUserReference"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Class ZtConfigUserReference.
' It preserves a part of the user configuration informations and methods to read them from ZtConfig.xml.
'
' Zotero Tools.
' This software is under Revised ('New') BSD license.
' Copyright © 2019, Olaf Ahrens. All rights reserved.
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Private variables.
Private pvtNumberPrefix As String
Private pvtNumberPrefixRegPiece As ZtRegPiece
Private pvtNumberSuffix As String
Private pvtNumberSuffixRegPiece As ZtRegPiece
Private pvtReferenceNormalSuffix As String
Private pvtReferenceNormalSuffixRegPiece As ZtRegPiece
Private pvtReferenceHyperlinkAtEndSuffix As String
Private pvtReferenceHyperlinkAtEndSuffixRegPiece As ZtRegPiece
Private pvtReferenceAccessedAtEndSuffix As String
Private pvtReferenceAccessedAtEndSuffixRegPiece As ZtRegPiece
Private pvtAccessedAtEndPrefix As String
Private pvtAccessedAtEndPrefixRegPiece As ZtRegPiece
Private pvtAccessedAtEndSuffix As String
Private pvtAccessedAtEndSuffixRegPiece As ZtRegPiece
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Friend/Public procedures and properties.
' All members that should be callable by CallByName procedure must be public.
Friend Sub Initialize(ByVal valConfig As ZtConfig, ByVal valXml As MSXML2.DOMDocument60, ByVal valStyleName As String)

  Dim locNode As MSXML2.IXMLDOMNode

  Set locNode = valXml.SelectSingleNode("/zoteroTools/user/referenceStyle[@name='" & valStyleName & "']")
  With locNode
    pvtNumberPrefix = valConfig.ResolveTextElement(.SelectSingleNode("numberPrefix"))
    valConfig.RegPieceFactoryFromString pvtNumberPrefix, _
                                        "ZtConfigXml.zoteroTools.user.referenceStyle(" & valStyleName & ").numberPrefix", _
                                        pvtNumberPrefixRegPiece
    pvtNumberSuffix = valConfig.ResolveTextElement(.SelectSingleNode("numberSuffix"))
    valConfig.RegPieceFactoryFromString pvtNumberSuffix, _
                                        "ZtConfigXml.zoteroTools.user.referenceStyle(" & valStyleName & ").numberSuffix", _
                                        pvtNumberSuffixRegPiece
    pvtReferenceNormalSuffix = valConfig.ResolveTextElement(.SelectSingleNode("referenceNormalSuffix"))
    valConfig.RegPieceFactoryFromString pvtReferenceNormalSuffix, _
                                        "ZtConfigXml.zoteroTools.user.referenceStyle(" & valStyleName & ").referenceNormalSuffix", _
                                        pvtReferenceNormalSuffixRegPiece
    pvtReferenceHyperlinkAtEndSuffix = valConfig.ResolveTextElement(.SelectSingleNode("referenceHyperlinkAtEndSuffix"))
    valConfig.RegPieceFactoryFromString pvtReferenceHyperlinkAtEndSuffix, _
                                        "ZtConfigXml.zoteroTools.user.referenceStyle(" & valStyleName & ").referenceHyperlinkAtEndSuffix", _
                                        pvtReferenceHyperlinkAtEndSuffixRegPiece
    pvtReferenceAccessedAtEndSuffix = valConfig.ResolveTextElement(.SelectSingleNode("referenceAccessedAtEndSuffix"))
    valConfig.RegPieceFactoryFromString pvtReferenceAccessedAtEndSuffix, _
                                        "ZtConfigXml.zoteroTools.user.referenceStyle(" & valStyleName & ").referenceAccessedAtEndSuffix", _
                                        pvtReferenceAccessedAtEndSuffixRegPiece
    pvtAccessedAtEndPrefix = valConfig.ResolveTextElement(.SelectSingleNode("accessedAtEndPrefix"))
    valConfig.RegPieceFactoryFromString pvtAccessedAtEndPrefix, _
                                        "ZtConfigXml.zoteroTools.user.referenceStyle(" & valStyleName & ").accessedAtEndPrefix", _
                                        pvtAccessedAtEndPrefixRegPiece
    pvtAccessedAtEndSuffix = valConfig.ResolveTextElement(.SelectSingleNode("accessedAtEndSuffix"))
    valConfig.RegPieceFactoryFromString pvtAccessedAtEndSuffix, _
                                        "ZtConfigXml.zoteroTools.user.referenceStyle(" & valStyleName & ").accessedAtEndSuffix", _
                                        pvtAccessedAtEndSuffixRegPiece

  End With

End Sub

Friend Property Get NumberPrefix() As String

  NumberPrefix = pvtNumberPrefix
  
End Property

Public Property Get NumberPrefixRegPiece() As ZtRegPiece

  Set NumberPrefixRegPiece = pvtNumberPrefixRegPiece
  
End Property

Friend Property Get NumberSuffix() As String

  NumberSuffix = pvtNumberSuffix
  
End Property

Public Property Get NumberSuffixRegPiece() As ZtRegPiece

  Set NumberSuffixRegPiece = pvtNumberSuffixRegPiece
  
End Property

Friend Property Get ReferenceNormalSuffix() As String

  ReferenceNormalSuffix = pvtReferenceNormalSuffix
  
End Property

Public Property Get ReferenceNormalSuffixRegPiece() As ZtRegPiece

  Set ReferenceNormalSuffixRegPiece = pvtReferenceNormalSuffixRegPiece
  
End Property

Friend Property Get ReferenceHyperlinkAtEndSuffix() As String

  ReferenceHyperlinkAtEndSuffix = pvtReferenceHyperlinkAtEndSuffix
  
End Property

Public Property Get ReferenceHyperlinkAtEndSuffixRegPiece() As ZtRegPiece

  Set ReferenceHyperlinkAtEndSuffixRegPiece = pvtReferenceHyperlinkAtEndSuffixRegPiece
  
End Property

Friend Property Get ReferenceAccessedAtEndSuffix() As String

  ReferenceAccessedAtEndSuffix = pvtReferenceAccessedAtEndSuffix
  
End Property

Public Property Get ReferenceAccessedAtEndSuffixRegPiece() As ZtRegPiece

  Set ReferenceAccessedAtEndSuffixRegPiece = pvtReferenceAccessedAtEndSuffixRegPiece
  
End Property

Friend Property Get AccessedAtEndPrefix() As String

  AccessedAtEndPrefix = pvtAccessedAtEndPrefix
  
End Property

Public Property Get AccessedAtEndPrefixRegPiece() As ZtRegPiece

  Set AccessedAtEndPrefixRegPiece = pvtAccessedAtEndPrefixRegPiece
  
End Property

Friend Property Get AccessedAtEndSuffix() As String

  AccessedAtEndSuffix = pvtAccessedAtEndSuffix
  
End Property

Public Property Get AccessedAtEndSuffixRegPiece() As ZtRegPiece

  Set AccessedAtEndSuffixRegPiece = pvtAccessedAtEndSuffixRegPiece
  
End Property
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

