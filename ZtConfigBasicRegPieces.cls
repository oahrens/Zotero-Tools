VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ZtConfigBasicRegPieces"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Class ZtConfigBasicRegPieces.
' It preserves a part of the basic configuration informations and methods to read them from ZtConfig.xml.
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Private variables.
Private pvtPunctuationPattern As ZtRegPiece
Private pvtNotPunctuationPattern As ZtRegPiece
Private pvtNotSpaceWOZeroPunctuationPattern As ZtRegPiece
Private pvtNotSpaceWOZeroPunctuationBreakFieldPattern As ZtRegPiece
Private pvtNotPunctuationBreakFieldPattern As ZtRegPiece
Private pvtZeroWidthSpace As ZtRegPiece
Private pvtSpacePattern As ZtRegPiece
Private pvtSpaceWOZeroPattern As ZtRegPiece
Private pvtSpaceWOZeroChoice As ZtRegPiece
Private pvtWhiteSpaceWOZeroPattern As ZtRegPiece
Private pvtSpaceWOZeroCommaSemicolonPattern As ZtRegPiece
Private pvtSpaceDashPattern As ZtRegPiece
Private pvtNotSpaceDashPattern As ZtRegPiece
Private pvtLineBreakChoice As ZtRegPiece
Private pvtHttpDoiPrefixChoice As ZtRegPiece
Private pvtCommaSeparatedIntegerList As ZtRegPiece
Private pvtHttpElementNameChoice As ZtRegPiece
Private pvtFamilyNameSuffixChoice As ZtRegPiece
Private pvtFamilyNameTheSuffixChoice As ZtRegPiece
Private pvtNotUrlAllowedCharacters As ZtRegPiece
Private pvtNotUrlAllowedPattern As ZtRegPiece
Private pvtNotUrlEndAllowedPattern As ZtRegPiece
Private pvtNotDoiAllowedCharacters As ZtRegPiece
Private pvtNotDoiAllowedPattern As ZtRegPiece
Private pvtNotDoiEndAllowedPattern As ZtRegPiece
Private pvtShortDoiAllowedPattern As ZtRegPiece
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Friend/Public procedures and properties.
' All members that should be callable by CallByName procedure must be public.
Friend Sub Initialize(ByVal valConfig As ZtConfig, ByVal valXml As MSXML2.DOMDocument60)

  Dim locNode As MSXML2.IXMLDOMNode
  
  Set locNode = valXml.SelectSingleNode("/zoteroTools/basic/regPieces")
  With locNode
    valConfig.RegPieceFactory valConfig.ResolveTextElement(.SelectSingleNode("punctuationPatternRegPiece")), _
                              "basic.regPieces.punctuationPatternRegPiece", _
                              pvtPunctuationPattern
    valConfig.RegPieceFactory valConfig.ResolveTextElement(.SelectSingleNode("notPunctuationPatternRegPiece")), _
                              "regPieces.notPunctuationPatternRegPiece", _
                              pvtNotPunctuationPattern
    valConfig.RegPieceFactory valConfig.ResolveTextElement(.SelectSingleNode("notSpaceWOZeroPunctuationPatternRegPiece")), _
                              "basic.regPieces.notSpaceWOZeroPunctuationPatternRegPiece", _
                              pvtNotSpaceWOZeroPunctuationPattern
    valConfig.RegPieceFactory valConfig.ResolveTextElement(.SelectSingleNode("notSpaceWOZeroPunctuationBreakFieldPatternRegPiece")), _
                              "basic.regPieces.notSpaceWOZeroPunctuationBreakFieldPatternRegPiece", _
                              pvtNotSpaceWOZeroPunctuationBreakFieldPattern
    valConfig.RegPieceFactory valConfig.ResolveTextElement(.SelectSingleNode("notPunctuationBreakFieldPatternRegPiece")), _
                              "basic.regPieces.notPunctuationBreakFieldPatternRegPiece", _
                              pvtNotPunctuationBreakFieldPattern
    valConfig.RegPieceFactory valConfig.ResolveTextElement(.SelectSingleNode("zeroWidthSpaceRegPiece")), _
                              "basic.regPieces.zeroWidthSpaceRegPiece", _
                              pvtZeroWidthSpace
    valConfig.RegPieceFactory valConfig.ResolveTextElement(.SelectSingleNode("spacePatternRegPiece")), _
                              "basic.regPieces.spacePattern", _
                              pvtSpacePattern
    valConfig.RegPieceFactory valConfig.ResolveTextElement(.SelectSingleNode("spaceWOZeroPatternRegPiece")), _
                              "basic.regPieces.spaceWOZeroPatternRegPiece", _
                              pvtSpaceWOZeroPattern
    valConfig.RegPieceFactory valConfig.ResolveTextElement(.SelectSingleNode("spaceWOZeroChoiceRegPiece")), _
                              "basic.regPieces.spaceWOZeroChoiceRegPiece", _
                              pvtSpaceWOZeroChoice
    valConfig.RegPieceFactory valConfig.ResolveTextElement(.SelectSingleNode("whiteSpaceWOZeroPatternRegPiece")), _
                              "basic.regPieces.whiteSpaceWOZeroPatternRegPiece", _
                              pvtWhiteSpaceWOZeroPattern
    valConfig.RegPieceFactory valConfig.ResolveTextElement(.SelectSingleNode("spaceWOZeroCommaSemicolonPatternRegPiece")), _
                              "basic.regPieces.spaceWOZeroCommaSemicolonPatternRegPiece", _
                              pvtSpaceWOZeroCommaSemicolonPattern
    valConfig.RegPieceFactory valConfig.ResolveTextElement(.SelectSingleNode("spaceDashPatternRegPiece")), _
                              "basic.regPieces.spaceDashPatternRegPiece", _
                              pvtSpaceDashPattern
    valConfig.RegPieceFactory valConfig.ResolveTextElement(.SelectSingleNode("notSpaceDashPatternRegPiece")), _
                              "basic.regPieces.notSpaceDashPatternRegPiece", _
                              pvtNotSpaceDashPattern
    valConfig.RegPieceFactory valConfig.ResolveTextElement(.SelectSingleNode("lineBreakChoiceRegPiece")), _
                              "basic.regPieces.lineBreakChoiceRegPiece", _
                              pvtLineBreakChoice
    valConfig.RegPieceFactory valConfig.ResolveTextElement(.SelectSingleNode("httpDoiPrefixChoiceRegPiece")), _
                              "basic.regPieces.httpDoiPrefixChoiceRegPiece", _
                              pvtHttpDoiPrefixChoice
    valConfig.RegPieceFactory valConfig.ResolveTextElement(.SelectSingleNode("commaSeparatedIntegerListRegPiece")), _
                              "basic.regPieces.commaSeparatedIntegerListRegPiece", _
                              pvtCommaSeparatedIntegerList
    valConfig.RegPieceFactory valConfig.ResolveTextElement(.SelectSingleNode("httpElementNameChoiceRegPiece")), _
                              "basic.regPieces.httpElementNameChoiceRegPiece", _
                              pvtHttpElementNameChoice
    valConfig.RegPieceFactory valConfig.ResolveTextElement(.SelectSingleNode("familyNameSuffixChoiceRegPiece")), _
                              "basic.regPieces.familyNameSuffixChoiceRegPiece", _
                              pvtFamilyNameSuffixChoice
    valConfig.RegPieceFactory valConfig.ResolveTextElement(.SelectSingleNode("familyNameTheSuffixChoiceRegPiece")), _
                              "basic.regPieces.familyNameTheSuffixChoiceRegPiece", _
                              pvtFamilyNameTheSuffixChoice
    valConfig.RegPieceFactory valConfig.ResolveTextElement(.SelectSingleNode("notUrlAllowedCharactersRegPiece")), _
                              "basic.regPieces.notUrlAllowedCharactersRegPiece", _
                              pvtNotUrlAllowedCharacters
    valConfig.RegPieceFactory valConfig.ResolveTextElement(.SelectSingleNode("notUrlAllowedPatternRegPiece")), _
                              "basic.regPieces.notUrlAllowedPatternRegPiece", _
                              pvtNotUrlAllowedPattern
    valConfig.RegPieceFactory valConfig.ResolveTextElement(.SelectSingleNode("notUrlEndAllowedPatternRegPiece")), _
                              "basic.regPieces.notUrlEndAllowedPatternRegPiece", _
                              pvtNotUrlEndAllowedPattern
    valConfig.RegPieceFactory valConfig.ResolveTextElement(.SelectSingleNode("notDoiAllowedCharactersRegPiece")), _
                              "basic.regPieces.notDoiAllowedCharactersRegPiece", _
                              pvtNotDoiAllowedCharacters
    valConfig.RegPieceFactory valConfig.ResolveTextElement(.SelectSingleNode("notDoiAllowedPatternRegPiece")), _
                              "basic.regPieces.notDoiAllowedPatternRegPiece", _
                              pvtNotDoiAllowedPattern
    valConfig.RegPieceFactory valConfig.ResolveTextElement(.SelectSingleNode("notDoiEndAllowedPatternRegPiece")), _
                              "basic.regPieces.notDoiEndAllowedPatternRegPiece", _
                              pvtNotDoiEndAllowedPattern
    valConfig.RegPieceFactory valConfig.ResolveTextElement(.SelectSingleNode("shortDoiAllowedPatternRegPiece")), _
                              "basic.regPieces.shortDoiAllowedPatternRegPiece", _
                              pvtShortDoiAllowedPattern
                       
  End With

End Sub

Public Property Get PunctuationPatternRegPiece() As ZtRegPiece

  Set PunctuationPatternRegPiece = pvtPunctuationPattern
  
End Property

Public Property Get NotPunctuationPatternRegPiece() As ZtRegPiece

  Set NotPunctuationPatternRegPiece = pvtNotPunctuationPattern
  
End Property

Public Property Get NotSpaceWOZeroPunctuationPatternRegPiece() As ZtRegPiece

  Set NotSpaceWOZeroPunctuationPatternRegPiece = pvtNotSpaceWOZeroPunctuationPattern
  
End Property

Public Property Get NotSpaceWOZeroPunctuationBreakFieldPatternRegPiece() As ZtRegPiece

  Set NotSpaceWOZeroPunctuationBreakFieldPatternRegPiece = pvtNotSpaceWOZeroPunctuationBreakFieldPattern
  
End Property

Public Property Get NotPunctuationBreakFieldPatternRegPiece() As ZtRegPiece

  Set NotPunctuationBreakFieldPatternRegPiece = pvtNotPunctuationBreakFieldPattern
  
End Property

Public Property Get ZeroWidthSpaceRegPiece() As ZtRegPiece

  Set ZeroWidthSpaceRegPiece = pvtZeroWidthSpace
  
End Property

Public Property Get SpacePatternRegPiece() As ZtRegPiece

  Set SpacePatternRegPiece = pvtSpacePattern
  
End Property

Public Property Get SpaceWOZeroPatternRegPiece() As ZtRegPiece

  Set SpaceWOZeroPatternRegPiece = pvtSpaceWOZeroPattern
  
End Property

Public Property Get SpaceWOZeroChoiceRegPiece() As ZtRegPiece

  Set SpaceWOZeroChoiceRegPiece = pvtSpaceWOZeroChoice
  
End Property

Public Property Get WhiteSpaceWOZeroPatternRegPiece() As ZtRegPiece

  Set WhiteSpaceWOZeroPatternRegPiece = pvtWhiteSpaceWOZeroPattern
  
End Property

Public Property Get SpaceWOZeroCommaSemicolonPatternRegPiece() As ZtRegPiece

  Set SpaceWOZeroCommaSemicolonPatternRegPiece = pvtSpaceWOZeroCommaSemicolonPattern
  
End Property

Public Property Get SpaceDashPatternRegPiece() As ZtRegPiece

  Set SpaceDashPatternRegPiece = pvtSpaceDashPattern
  
End Property

Public Property Get NotSpaceDashPatternRegPiece() As ZtRegPiece

  Set NotSpaceDashPatternRegPiece = pvtNotSpaceDashPattern
  
End Property

Public Property Get LineBreakChoiceRegPiece() As ZtRegPiece

  Set LineBreakChoiceRegPiece = pvtLineBreakChoice
  
End Property

Public Property Get HttpDoiPrefixChoiceRegPiece() As ZtRegPiece

  Set HttpDoiPrefixChoiceRegPiece = pvtHttpDoiPrefixChoice
  
End Property

Public Property Get CommaSeparatedIntegerListRegPiece() As ZtRegPiece

  Set CommaSeparatedIntegerListRegPiece = pvtCommaSeparatedIntegerList
  
End Property

Public Property Get HttpElementNameChoiceRegPiece() As ZtRegPiece

  Set HttpElementNameChoiceRegPiece = pvtHttpElementNameChoice
  
End Property

Public Property Get FamilyNameSuffixChoiceRegPiece() As ZtRegPiece

  Set FamilyNameSuffixChoiceRegPiece = pvtFamilyNameSuffixChoice
  
End Property

Public Property Get FamilyNameTheSuffixChoiceRegPiece() As ZtRegPiece

  Set FamilyNameTheSuffixChoiceRegPiece = pvtFamilyNameTheSuffixChoice
  
End Property

Public Property Get NotUrlAllowedCharactersRegPiece() As ZtRegPiece

  Set NotUrlAllowedCharactersRegPiece = pvtNotUrlAllowedCharacters
  
End Property

Public Property Get NotUrlAllowedPatternRegPiece() As ZtRegPiece

  Set NotUrlAllowedPatternRegPiece = pvtNotUrlAllowedPattern
  
End Property

Public Property Get NotUrlEndAllowedPatternRegPiece() As ZtRegPiece

  Set NotUrlEndAllowedPatternRegPiece = pvtNotUrlEndAllowedPattern
  
End Property

Public Property Get NotDoiAllowedCharactersRegPiece() As ZtRegPiece

  Set NotDoiAllowedCharactersRegPiece = pvtNotDoiAllowedCharacters
  
End Property

Public Property Get NotDoiAllowedPatternRegPiece() As ZtRegPiece

  Set NotDoiAllowedPatternRegPiece = pvtNotDoiAllowedPattern
  
End Property

Public Property Get NotDoiEndAllowedPatternRegPiece() As ZtRegPiece

  Set NotDoiEndAllowedPatternRegPiece = pvtNotDoiEndAllowedPattern
  
End Property

Public Property Get ShortDoiAllowedPatternRegPiece() As ZtRegPiece

  Set ShortDoiAllowedPatternRegPiece = pvtShortDoiAllowedPattern
  
End Property
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

