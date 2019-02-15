VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ZtConfigUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Class ZtConfigUser.
' It preserves user configuration informations and methods to read them from ZtConfig.xml.
' It distributes these informations and work to the other ZtConfigUser... classes.
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Private variables.
Private pvtConfig As ZtConfig
Private pvtXml As MSXML2.DOMDocument60
Private pvtBackwardLinking As ZtConfigUserBackwardLinking
Private pvtCitation As ZtConfigUserCitation
Private pvtCitationGroup As ZtConfigUserCitationGroup
Private pvtMacro As ZtConfigUserMacro
Private pvtPunctuation As ZtConfigUserPunctuation
Private pvtReference As ZtConfigUserReference
Private pvtResolveCitation As ZtConfigUserResolveCitation
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Constructor.
Private Sub Class_Initialize()

  Set pvtMacro = New ZtConfigUserMacro

End Sub
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Friend/Public procedures and properties.
' All members that should be callable by CallByName procedure must be public.
Friend Sub Initialize(ByVal valConfig As ZtConfig, ByVal valXml As MSXML2.DOMDocument60)

  Set pvtConfig = valConfig
  Set pvtXml = valXml
  pvtMacro.Initialize pvtConfig, pvtXml

End Sub

Friend Sub KeepUserStyles(ByVal valName As String)

  Dim locPreset As ZtConfigUserStylePreset
  
  Set locPreset = pvtConfig.UserStylePresets.Item(valName)
  
  ' Read citation group style.
  Set pvtCitationGroup = New ZtConfigUserCitationGroup
  pvtCitationGroup.Initialize pvtConfig, pvtXml, locPreset.CitationGroupStyle
  
  ' Read citation group punctuation style.
  Set pvtPunctuation = New ZtConfigUserPunctuation
  pvtPunctuation.Initialize pvtConfig, pvtXml, locPreset.PunctuationStyle
  
  ' Read citation style.
  Set pvtCitation = New ZtConfigUserCitation
  pvtCitation.Initialize pvtConfig, pvtXml, locPreset.CitationStyle

  ' Read resolve citation style.
  Set pvtResolveCitation = New ZtConfigUserResolveCitation
  pvtResolveCitation.Initialize pvtConfig, pvtXml, locPreset.ResolveCitationStyle
  
  ' Read reference style.
  Set pvtReference = New ZtConfigUserReference
  pvtReference.Initialize pvtConfig, pvtXml, locPreset.ReferenceStyle

  ' Read backward linking style.
  Set pvtBackwardLinking = New ZtConfigUserBackwardLinking
  pvtBackwardLinking.Initialize pvtConfig, pvtXml, locPreset.BackwardLinkingStyle

End Sub

Public Property Get BackwardLinkingStyle() As ZtConfigUserBackwardLinking

  Set BackwardLinkingStyle = pvtBackwardLinking
  
End Property

Public Property Get CitationStyle() As ZtConfigUserCitation

  Set CitationStyle = pvtCitation
  
End Property

Public Property Get CitationGroupStyle() As ZtConfigUserCitationGroup

  Set CitationGroupStyle = pvtCitationGroup
  
End Property

Public Property Get Macro() As ZtConfigUserMacro

  Set Macro = pvtMacro
  
End Property

Public Property Get PunctuationStyle() As ZtConfigUserPunctuation

  Set PunctuationStyle = pvtPunctuation
  
End Property

Public Property Get ResolveCitationStyle() As ZtConfigUserResolveCitation

  Set ResolveCitationStyle = pvtResolveCitation
  
End Property

Public Property Get ReferenceStyle() As ZtConfigUserReference

  Set ReferenceStyle = pvtReference
  
End Property
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
