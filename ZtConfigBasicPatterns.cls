VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ZtConfigBasicPatterns"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Class ZtConfigBasicPatterns.
' It preserves a part of the basic configuration informations and methods to read them from ZtConfig.xml.
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Private variables.
Private pvtPunctuation As String
Private pvtSpaceWOZero As String
Private pvtSpaceWOZeroPunctuation As String
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Friend/Public procedures and properties.
' All members that should be callable by CallByName procedure must be public.
Friend Sub Initialize(ByVal valConfig As ZtConfig, ByVal valXml As MSXML2.DOMDocument60)

  Dim locNode As MSXML2.IXMLDOMNode
  
  Set locNode = valXml.SelectSingleNode("/zoteroTools/basic/patterns")
  With locNode
    pvtPunctuation = valConfig.ResolveTextElement(.SelectSingleNode("punctuation"))
    pvtSpaceWOZero = valConfig.ResolveTextElement(.SelectSingleNode("spaceWOZero"))
    pvtSpaceWOZeroPunctuation = valConfig.ResolveTextElement(.SelectSingleNode("spaceWOZeroPunctuation"))
  End With

End Sub

Public Property Get Punctuation() As String

  Punctuation = pvtPunctuation
  
End Property

Public Property Get SpaceWOZero() As String

  SpaceWOZero = pvtSpaceWOZero
  
End Property

Public Property Get SpaceWOZeroPunctuation() As String

  SpaceWOZeroPunctuation = pvtSpaceWOZeroPunctuation
  
End Property
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

