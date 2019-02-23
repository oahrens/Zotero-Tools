VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ZtConfigUserMacro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Class ZtConfigUserMacro.
' It preserves a part of the user configuration informations and methods to read them from ZtConfig.xml.
'
' Zotero Tools.
' This software is under Revised ('New') BSD license.
' Copyright © 2019, Olaf Ahrens. All rights reserved.
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Private variables.
Private pvtWordIsInvisibleWhileOperation As Boolean
Private pvtDebugging As Boolean
Private pvtCitationInsertZeroWidthSpace As Boolean
Private pvtWithBackwardLinking As Boolean
Private pvtStylePresetName As String
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Friend/Public procedures and properties.
' All members that should be callable by CallByName procedure must be public.
Friend Sub Initialize(ByVal valConfig As ZtConfig, ByVal valXml As MSXML2.DOMDocument60)

  Dim locNode As MSXML2.IXMLDOMNode

  Set locNode = valXml.SelectSingleNode("/zoteroTools/user/macro")
  pvtStylePresetName = valConfig.ResolveTextElement(locNode.SelectSingleNode("stylePresetName"))
  pvtWordIsInvisibleWhileOperation = CBool(locNode.SelectSingleNode("wordIsInvisibleWhileOperation").Text)
  pvtCitationInsertZeroWidthSpace = CBool(locNode.SelectSingleNode("citationInsertZeroWidthSpace").Text)
  pvtWithBackwardLinking = CBool(locNode.SelectSingleNode("withBackwardLinking").Text)
  pvtDebugging = CBool(locNode.SelectSingleNode("debugging").Text)
  
End Sub

Public Property Let WordIsInvisibleWhileOperation(ByVal valVisible As Boolean)

  pvtWordIsInvisibleWhileOperation = valVisible
  
End Property

Public Property Get WordIsInvisibleWhileOperation() As Boolean

  WordIsInvisibleWhileOperation = pvtWordIsInvisibleWhileOperation
  
End Property

Public Property Let Debugging(ByVal valDebugging As Boolean)

  pvtDebugging = valDebugging
  
End Property

Public Property Get Debugging() As Boolean

  Debugging = pvtDebugging
  
End Property

Public Property Let CitationInsertZeroWidthSpace(ByVal valInsert As Boolean)

  pvtCitationInsertZeroWidthSpace = valInsert
  
End Property

Public Property Get CitationInsertZeroWidthSpace() As Boolean

  CitationInsertZeroWidthSpace = pvtCitationInsertZeroWidthSpace
  
End Property

Public Property Let WithBackwardLinking(ByVal valWithLinking As Boolean)

  pvtWithBackwardLinking = valWithLinking
  
End Property

Public Property Get WithBackwardLinking() As Boolean

  WithBackwardLinking = pvtWithBackwardLinking
  
End Property

Public Property Let StylePresetName(ByVal valName As String)

  pvtStylePresetName = valName
  
End Property

Public Property Get StylePresetName() As String

  StylePresetName = pvtStylePresetName
  
End Property
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

