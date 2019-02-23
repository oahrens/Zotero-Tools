VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ZtConfigBasicWord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Class ZtConfigBasicWord.
' It preserves a part of the basic configuration informations and methods to read them from ZtConfig.xml.
'
' Zotero Tools.
' This software is under Revised ('New') BSD license.
' Copyright © 2019, Olaf Ahrens. All rights reserved.
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Private variables.
Private pvtMaxScreenTipTextLength As Integer
Private pvtCantHoldFieldStoryTypes() As WdStoryType
Private pvtMultipleStoriesStoryTypes() As WdStoryType
Private pvtIndividuallyUpdateFieldStoryTypes() As WdStoryType
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Friend/Public procedures and properties.
' All members that should be callable by CallByName procedure must be public.
Friend Sub Initialize(ByVal valConfig As ZtConfig, ByVal valXml As MSXML2.DOMDocument60)

  Dim locNode As MSXML2.IXMLDOMNode
  Dim locEnums() As String
  
  Set locNode = valXml.SelectSingleNode("/zoteroTools/basic/word")
  With locNode
    pvtMaxScreenTipTextLength = CInt(.SelectSingleNode("maxScreenTipTextLength").Text)
    locEnums = Split(valConfig.ResolveTextElement(.SelectSingleNode("cantHoldFieldStoryTypes")), ",")
    pvtCantHoldFieldStoryTypes = valConfig.GetWordStoryTypes(Split(valConfig.ResolveTextElement(.SelectSingleNode("cantHoldFieldStoryTypes")), ","))
    pvtMultipleStoriesStoryTypes = valConfig.GetWordStoryTypes(Split(valConfig.ResolveTextElement(.SelectSingleNode("multipleStoriesStoryTypes")), ","))
    pvtIndividuallyUpdateFieldStoryTypes = valConfig.GetWordStoryTypes(Split(valConfig.ResolveTextElement(.SelectSingleNode("individuallyUpdateFieldStoryTypes")), ","))
  End With
  

End Sub

Public Property Get MaxScreenTipTextLength() As Integer

  MaxScreenTipTextLength = pvtMaxScreenTipTextLength
  
End Property

Public Property Get CantHoldFieldStoryTypes() As WdStoryType()

  CantHoldFieldStoryTypes = pvtCantHoldFieldStoryTypes
  
End Property

Public Property Get MultipleStoriesStoryTypes() As WdStoryType()

  MultipleStoriesStoryTypes = pvtMultipleStoriesStoryTypes
  
End Property

Public Property Get IndividuallyUpdateFieldStoryTypes() As WdStoryType()

  IndividuallyUpdateFieldStoryTypes = pvtIndividuallyUpdateFieldStoryTypes
  
End Property
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

