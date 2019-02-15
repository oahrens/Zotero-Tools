VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ZtConfigUserCitation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Class ZtConfigUserCitation.
' It preserves a part of the user configuration informations and methods to read them from ZtConfig.xml.
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Private variables.
Private pvtPrefix As String
Private pvtPrefixRegPiece As ZtRegPiece
Private pvtSuffix As String
Private pvtSuffixRegPiece As ZtRegPiece
Private pvtLocalizerRegString As String
Private pvtLocalizerRegPiece As ZtRegPiece
Private pvtLocalizerIsBeforeSuffix As Boolean
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Friend/Public procedures and properties.
' All members that should be callable by CallByName procedure must be public.
Friend Sub Initialize(ByVal valConfig As ZtConfig, ByVal valXml As MSXML2.DOMDocument60, ByVal valStyleName As String)

  Dim locNode As MSXML2.IXMLDOMNode
  
  Set locNode = valXml.SelectSingleNode("/zoteroTools/user/citationStyle[@name='" & valStyleName & "']")
  With locNode
    pvtPrefix = valConfig.ResolveTextElement(.SelectSingleNode("prefix"))
    valConfig.RegPieceFactoryFromString pvtPrefix, _
                                        "ZtConfigXml.zoteroTools.user.citationStyle(" & valStyleName & ").prefix", _
                                        pvtPrefixRegPiece
    pvtSuffix = valConfig.ResolveTextElement(.SelectSingleNode("suffix"))
    valConfig.RegPieceFactoryFromString pvtSuffix, _
                                        "ZtConfigXml.zoteroTools.user.citationStyle(" & valStyleName & ").suffix", _
                                        pvtSuffixRegPiece
    pvtLocalizerRegString = valConfig.ResolveTextElement(.SelectSingleNode("localizerRegString"))
    valConfig.RegPieceFactory pvtLocalizerRegString, _
                              "ZtConfigXml.zoteroTools.user.citationStyle(" & valStyleName & ").localizerRegString", _
                              pvtLocalizerRegPiece
    pvtLocalizerIsBeforeSuffix = CBool(.SelectSingleNode("localizerIsBeforeSuffix").Text)
  End With

End Sub

Friend Property Get Prefix() As String

  Prefix = pvtPrefix
  
End Property

Public Property Get PrefixRegPiece() As ZtRegPiece

  Set PrefixRegPiece = pvtPrefixRegPiece
  
End Property

Friend Property Get Suffix() As String

  Suffix = pvtSuffix
  
End Property

Public Property Get SuffixRegPiece() As ZtRegPiece

  Set SuffixRegPiece = pvtSuffixRegPiece
  
End Property

Public Property Get LocalizerRegString() As String

  LocalizerRegString = pvtLocalizerRegString
  
End Property

Public Property Get LocalizerRegPiece() As ZtRegPiece

  Set LocalizerRegPiece = pvtLocalizerRegPiece
  
End Property

Friend Property Get LocalizerIsBeforeSuffix() As Boolean

  LocalizerIsBeforeSuffix = pvtLocalizerIsBeforeSuffix
  
End Property
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

