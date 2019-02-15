VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ZtConfigUserBackwardLinking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Class ZtConfigUserBackwardLinking.
' It preserves a part of the user configuration informations and methods to read them from ZtConfig.xml.
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Private variables.
Private pvtMaxCaptionLevel As Integer
Private pvtHyperlinkStyleName As String
Private pvtDelimiterBefore As String
Private pvtDelimiterBeforeRegPiece As ZtRegPiece
Private pvtPrefixSingular As String
Private pvtPrefixSingularRegPiece As ZtRegPiece
Private pvtPrefixPlural As String
Private pvtPrefixPluralRegPiece As ZtRegPiece
Private pvtSuffix As String
Private pvtSuffixRegPiece As ZtRegPiece
Private pvtDelimiter As String
Private pvtScreenTipPattern As String
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Friend/Public procedures and properties.
' All members that should be callable by CallByName procedure must be public.
Friend Sub Initialize(ByVal valConfig As ZtConfig, ByVal valXml As MSXML2.DOMDocument60, ByVal valStyleName As String)

  Dim locNode As MSXML2.IXMLDOMNode
  
  Set locNode = valXml.SelectSingleNode("/zoteroTools/user/backwardLinkingStyle[@name='" & valStyleName & "']")
  With locNode
    pvtMaxCaptionLevel = CInt(.SelectSingleNode("maxCaptionLevel").Text)
    pvtHyperlinkStyleName = valConfig.ResolveTextElement(.SelectSingleNode("hyperlinkStyleName"))
    pvtDelimiterBefore = valConfig.ResolveTextElement(.SelectSingleNode("delimiterBefore"))
    valConfig.RegPieceFactoryFromString pvtDelimiterBefore, _
                                        "ZtConfigXml.zoteroTools.user.backwardLinkingStyle(" & valStyleName & ").delimiterBefore", _
                                        pvtDelimiterBeforeRegPiece
    pvtPrefixSingular = valConfig.ResolveTextElement(.SelectSingleNode("prefixSingular"))
    valConfig.RegPieceFactoryFromString pvtPrefixSingular, _
                                        "ZtConfigXml.zoteroTools.user.backwardLinkingStyle(" & valStyleName & ").prefixSingular", _
                                        pvtPrefixSingularRegPiece
    pvtPrefixPlural = valConfig.ResolveTextElement(.SelectSingleNode("prefixPlural"))
    valConfig.RegPieceFactoryFromString pvtPrefixPlural, _
                                        "ZtConfigXml.zoteroTools.user.backwardLinkingStyle(" & valStyleName & ").prefixPlural", _
                                        pvtPrefixPluralRegPiece
    pvtDelimiter = valConfig.ResolveTextElement(.SelectSingleNode("delimiter"))
    pvtSuffix = valConfig.ResolveTextElement(.SelectSingleNode("suffix"))
    valConfig.RegPieceFactoryFromString pvtSuffix, _
                                        "ZtConfigXml.zoteroTools.user.backwardLinkingStyle(" & valStyleName & ").suffix", _
                                        pvtSuffixRegPiece
    pvtScreenTipPattern = valConfig.ResolveTextElement(.SelectSingleNode("screenTipPattern"))
  End With

End Sub

Public Property Get MaxCaptionLevel() As Integer

  MaxCaptionLevel = pvtMaxCaptionLevel
  
End Property

Public Property Get HyperlinkStyleName() As String

  HyperlinkStyleName = pvtHyperlinkStyleName
  
End Property

Friend Property Get DelimiterBefore() As String

  DelimiterBefore = pvtDelimiterBefore
  
End Property

Public Property Get DelimiterBeforeRegPiece() As ZtRegPiece

  Set DelimiterBeforeRegPiece = pvtDelimiterBeforeRegPiece
  
End Property

Friend Property Get PrefixSingular() As String

  PrefixSingular = pvtPrefixSingular
  
End Property

Public Property Get PrefixSingularRegPiece() As ZtRegPiece

  Set PrefixSingularRegPiece = pvtPrefixSingularRegPiece
  
End Property

Friend Property Get PrefixPlural() As String

  PrefixPlural = pvtPrefixPlural
  
End Property

Public Property Get PrefixPluralRegPiece() As ZtRegPiece

  Set PrefixPluralRegPiece = pvtPrefixPluralRegPiece
  
End Property

Friend Property Get Suffix() As String

  Suffix = pvtSuffix
  
End Property

Public Property Get SuffixRegPiece() As ZtRegPiece

  Set SuffixRegPiece = pvtSuffixRegPiece
  
End Property

Friend Property Get Delimiter() As String

  Delimiter = pvtDelimiter
  
End Property

Friend Property Get ScreenTipPattern() As String

  ScreenTipPattern = pvtScreenTipPattern
  
End Property
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

