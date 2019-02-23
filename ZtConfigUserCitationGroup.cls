VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ZtConfigUserCitationGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Class ZtConfigUserCitationGroup.
' It preserves a part of the user configuration informations and methods to read them from ZtConfig.xml.
'
' Zotero Tools.
' This software is under Revised ('New') BSD license.
' Copyright © 2019, Olaf Ahrens. All rights reserved.
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Private variables.
Private pvtRangeDelimiterRegString As String
Private pvtRangeDelimiterRegPiece As ZtRegPiece
Private pvtPrefix As String
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Friend/Public procedures and properties.
' All members that should be callable by CallByName procedure must be public.
Friend Sub Initialize(ByVal valConfig As ZtConfig, ByVal valXml As MSXML2.DOMDocument60, ByVal valStyleName As String)

  Dim locNode As MSXML2.IXMLDOMNode
  
  Set locNode = valXml.SelectSingleNode("/zoteroTools/user/citationGroupStyle[@name='" & valStyleName & "']")
  With locNode
    pvtRangeDelimiterRegString = valConfig.ResolveTextElement(.SelectSingleNode("rangeDelimiterRegString"))
    valConfig.RegPieceFactory pvtRangeDelimiterRegString, _
                              "ZtConfigXml.zoteroTools.user.citationGroupStyle(" & valStyleName & ").rangeDelimiterRegString", _
                              pvtRangeDelimiterRegPiece
    pvtPrefix = valConfig.ResolveTextElement(.SelectSingleNode("prefix"))
  End With

End Sub

Public Property Get RangeDelimiterRegString() As String

  RangeDelimiterRegString = pvtRangeDelimiterRegString
  
End Property

Public Property Get RangeDelimiterRegPiece() As ZtRegPiece

  Set RangeDelimiterRegPiece = pvtRangeDelimiterRegPiece
  
End Property

Friend Property Get Prefix() As String

  Prefix = pvtPrefix
  
End Property
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

