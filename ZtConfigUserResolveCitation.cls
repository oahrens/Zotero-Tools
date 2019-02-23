VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ZtConfigUserResolveCitation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Class ZtConfigUserResolveCitation.
' It preserves a part of the user configuration informations and methods to read them from ZtConfig.xml.
'
' Zotero Tools.
' This software is under Revised ('New') BSD license.
' Copyright © 2019, Olaf Ahrens. All rights reserved.
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Private variables.
Private pvtMaxNamesCt As Integer
Private pvtGivenNameStyle As ZtEGivenNameStyle
Private pvtGroupPrefix As String
Private pvtGroupSuffix As String
Private pvtPrefix As String
Private pvtSuffix As String
Private pvtDelimiter As String
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Friend/Public procedures and properties.
' All members that should be callable by CallByName procedure must be public.
Friend Sub Initialize(ByVal valConfig As ZtConfig, ByVal valXml As MSXML2.DOMDocument60, ByVal valStyleName As String)

  Dim locNode As MSXML2.IXMLDOMNode

  Set locNode = valXml.SelectSingleNode("/zoteroTools/user/resolveCitationStyle[@name='" & valStyleName & "']")
  With locNode
    pvtMaxNamesCt = CInt(.SelectSingleNode("maxNamesCt").Text)
    pvtGivenNameStyle = valConfig.GetGivenNameStyle(valConfig.ResolveTextElement(.SelectSingleNode("givenNameStyle")))
    pvtGroupPrefix = valConfig.ResolveTextElement(.SelectSingleNode("groupPrefix"))
    pvtGroupSuffix = valConfig.ResolveTextElement(.SelectSingleNode("groupSuffix"))
    pvtPrefix = valConfig.ResolveTextElement(.SelectSingleNode("prefix"))
    pvtSuffix = valConfig.ResolveTextElement(.SelectSingleNode("suffix"))
    pvtDelimiter = valConfig.ResolveTextElement(.SelectSingleNode("delimiter"))
  End With

End Sub

Public Property Get MaxNamesCt() As Integer

  MaxNamesCt = pvtMaxNamesCt
  
End Property

Public Property Get GivenNameStyle() As ZtEGivenNameStyle

  GivenNameStyle = pvtGivenNameStyle
  
End Property

Public Property Get GroupPrefix() As String

  GroupPrefix = pvtGroupPrefix
  
End Property

Public Property Get GroupSuffix() As String

  GroupSuffix = pvtGroupSuffix
  
End Property

Public Property Get Prefix() As String

  Prefix = pvtPrefix
  
End Property

Public Property Get Suffix() As String

  Suffix = pvtSuffix
  
End Property

Public Property Get Delimiter() As String

  Delimiter = pvtDelimiter
  
End Property
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

