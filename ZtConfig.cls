VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ZtConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Class ZtConfig.
' It preserves all configuration informations and methods to read them from ZtConfig.xml.
' It distributes these informations and work to ZtConfigBasic, ZtConfigUser, and ZtConfigFinal.
'
' Zotero Tools.
' This software is under Revised ('New') BSD license.
' Copyright © 2019, Olaf Ahrens. All rights reserved.
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Private constants.
Private Const PVT_CONFIG_FILE_NAME As String = "ZtConfig.xml"
Private Const PVT_CONTROL_CHARACTER_REGEX_PATTERN As String = "(?<signsBefore>^|[^\\](?:\\{2})*)(?<whole>\{Chr\((?:(?:x(?<hexNumber>[\dabcdefABCDEF]+))|(?<decNumber>\d+)\)\}))"
Private Const PVT_RESOLVE_PROPERTY_REGEX_PATTERN As String = "(?<signsBefore>^|[^\\](?:\\{2})*)(?<whole>\{(?<propertyString>[a-zA-Z0-9_]+(?:\.[a-zA-Z0-9_]+){2,3})\})"
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Private variables.
Private pvtMessageDisplay As ZtIMessageDisplayable
Private pvtBasic As ZtConfigBasic
Private pvtUserStylePresets As Collection
Private pvtUser As ZtConfigUser
Private pvtXml As MSXML2.DOMDocument60
Private pvtControlCharacterRegExp As ZtRegExp
Private pvtResolvePropertyRegExp As ZtRegExp
Private pvtFinal As ZtConfigFinal
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Constructor.
Private Sub Class_Initialize()

  Set pvtXml = New MSXML2.DOMDocument60
  Set pvtUserStylePresets = New Collection
  Set pvtControlCharacterRegExp = New ZtRegExp
  Set pvtResolvePropertyRegExp = New ZtRegExp
  Set pvtBasic = New ZtConfigBasic
  Set pvtUser = New ZtConfigUser

End Sub
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Friend/Public procedures and properties.
' All members that should be callable by CallByName procedure must be public.
Friend Sub Initialize(ByVal valMessageDisplay As ZtIMessageDisplayable)

  Dim locXmlPathFileName As String
  
  Set pvtMessageDisplay = valMessageDisplay
  
  locXmlPathFileName = ThisDocument.Path & "\" & PVT_CONFIG_FILE_NAME
  pvtXml.Load locXmlPathFileName
  pvtXml.preserveWhiteSpace = True
  pvtControlCharacterRegExp.Initialize "XML control characters", _
                                       PVT_CONTROL_CHARACTER_REGEX_PATTERN, _
                                       pvtMessageDisplay, _
                                       False, _
                                       True, _
                                       True
  pvtResolvePropertyRegExp.Initialize "XML resolve property", _
                                      PVT_RESOLVE_PROPERTY_REGEX_PATTERN, _
                                      pvtMessageDisplay, _
                                      False, _
                                      True, _
                                      True
                                       
  ' Set basic config.
  pvtBasic.Initialize Me, pvtXml
  
  ' Set user config.
  pvtUser.Initialize Me, pvtXml
  
End Sub

Public Property Get Basic() As ZtConfigBasic

  Set Basic = pvtBasic

End Property

Public Property Get User() As ZtConfigUser

  Set User = pvtUser
  
End Property

Public Property Get Final() As ZtConfigFinal

  Set Final = pvtFinal
  
End Property

Friend Property Get UserStylePresets() As Collection

  Set UserStylePresets = pvtUserStylePresets

End Property

Friend Sub KeepUserStylePresets()

  Dim locNode As MSXML2.IXMLDOMNode
  Dim locPreset0 As ZtConfigUserStylePreset
  Dim locPreset1 As ZtConfigUserStylePreset
  Dim locCtr As Integer
  
  For Each locNode In pvtXml.SelectNodes("/zoteroTools/user/stylePreset")
    Set locPreset0 = New ZtConfigUserStylePreset
    locPreset0.Initialize Me, locNode
    If pvtUserStylePresets.Count = 0 Then
      pvtUserStylePresets.Add locPreset0, locNode.Attributes.getNamedItem("name").Text
    Else
      locCtr = 1
      For Each locPreset1 In pvtUserStylePresets
        If locPreset0.Name < locPreset1.Name Then
          Exit For
        End If
        locCtr = locCtr + 1
      Next
      If locCtr <= pvtUserStylePresets.Count Then
        pvtUserStylePresets.Add locPreset0, locNode.Attributes.getNamedItem("name").Text, locCtr
      Else
        pvtUserStylePresets.Add locPreset0, locNode.Attributes.getNamedItem("name").Text
      End If
    End If
  Next

End Sub

Friend Sub KeepFinal()
  
  Set pvtFinal = New ZtConfigFinal
  pvtFinal.Initialize Me, pvtXml

End Sub

Friend Sub RegExpFactory(ByRef refRegExp As ZtRegExp, ByVal valName As String, ByVal valPattern As String, _
  Optional ByVal valGlobal As Boolean = False, Optional ByVal valIgnoreCase As Boolean = False, Optional ByVal valMultiLine As Boolean = False)

  Set refRegExp = New ZtRegExp
  refRegExp.Initialize valName, valPattern, pvtMessageDisplay, valGlobal, valIgnoreCase, valMultiLine

End Sub

Friend Function ResolveTextElement(ByVal valElement As MSXML2.IXMLDOMNode, Optional ByVal valClearTabAndLineBreak As Boolean = False) As String

  Dim locResult As String
  Dim locMatches As Collection
  Dim locMatch As ZtRegMatch
 
  locResult = valElement.Text
  
  If valClearTabAndLineBreak Then
    locResult = Replace(Replace(Replace(locResult, Chr$(9), vbNullString), Chr$(10), vbNullString), Chr$(13), vbNullString)
  End If
  
  Set locMatches = pvtControlCharacterRegExp.Execute(locResult)
  Do While locMatches.Count > 0
    For Each locMatch In locMatches
      With locMatch
        If Len(.Groups("hexNumber")) > 0 Then
          locResult = Left$(locResult, .FirstIndex + Len(.Groups("signsBefore"))) & _
                      ChrW$(CLng("&H" & .Groups("hexNumber"))) & _
                      Right$(locResult, Len(locResult) - .FirstIndex - .Length)
        Else
          locResult = Left$(locResult, .FirstIndex + Len(.Groups("signsBefore"))) & _
                      ChrW$(CLng(.Groups("decNumber"))) & _
                      Right$(locResult, Len(locResult) - .FirstIndex - .Length)
        End If
      End With
      Exit For
    Next
    Set locMatches = pvtControlCharacterRegExp.Execute(locResult)
  Loop
 
  Set locMatches = pvtResolvePropertyRegExp.Execute(locResult)
  Do While locMatches.Count > 0
    For Each locMatch In locMatches
      With locMatch
        locResult = Left$(locResult, .FirstIndex + Len(.Groups("signsBefore"))) & _
                    CStr(pvtResolvePropertyScript(Me, .Groups("propertyString"))) & _
                    Right$(locResult, Len(locResult) - .FirstIndex - .Length)
      End With
      Exit For
    Next
    Set locMatches = pvtResolvePropertyRegExp.Execute(locResult)
  Loop

  locResult = Replace(locResult, "\{", "{")

  ResolveTextElement = locResult

End Function

Friend Sub RegPieceFactoryFromString(ByVal valText As String, ByVal valName As String, ByRef refRegPiece As ZtRegPiece)

  Dim locSpecialChars As String
  Dim locCharCtr As Integer
  Dim locChar As String
  Dim locText As String
  
  locText = valText
  locSpecialChars = "\?+^$()[]{}.|"
  For locCharCtr = 1 To Len(locSpecialChars)
    locChar = Mid$(locSpecialChars, locCharCtr, 1)
    locText = Replace(locText, locChar, "\" & locChar)
  Next
  
  Me.RegPieceFactory locText, valName, refRegPiece
  
End Sub

Friend Sub RegPieceFactory(ByVal valTextWithRegExpSyntax As String, ByVal valName As String, ByRef refRegPiece As ZtRegPiece)
  
  If refRegPiece Is Nothing Then
    Set refRegPiece = New ZtRegPiece
    refRegPiece.Initialize valTextWithRegExpSyntax, valName, pvtMessageDisplay
  End If
  
End Sub

Friend Function GetGivenNameStyle(ByVal valEnumName As String) As ZtEGivenNameStyle

  Select Case valEnumName
    Case "GivenNameNone": GetGivenNameStyle = GivenNameNone
    Case "GivenNameAbbreviated": GetGivenNameStyle = GivenNameAbbreviated
    Case "GivenNameFull": GetGivenNameStyle = GivenNameFull
  End Select

End Function

Friend Function GetWordStoryTypes(ByRef refEnumNames As Variant) As WdStoryType()

  Dim locCtr As Integer
  Dim locTypes() As WdStoryType
  
  ReDim locTypes(UBound(refEnumNames))
  For locCtr = 0 To UBound(refEnumNames)
    Select Case refEnumNames(locCtr)
      Case "wdCommentsStory": locTypes(locCtr) = wdCommentsStory
      Case "wdEndnoteContinuationNoticeStory": locTypes(locCtr) = wdEndnoteContinuationNoticeStory
      Case "wdEndnoteContinuationSeparatorStory": locTypes(locCtr) = wdEndnoteContinuationSeparatorStory
      Case "wdEndnoteSeparatorStory": locTypes(locCtr) = wdEndnoteSeparatorStory
      Case "wdEndnotesStory": locTypes(locCtr) = wdEndnotesStory
      Case "wdEvenPagesFooterStory": locTypes(locCtr) = wdEvenPagesFooterStory
      Case "wdEvenPagesHeaderStory": locTypes(locCtr) = wdEvenPagesHeaderStory
      Case "wdFirstPageFooterStory": locTypes(locCtr) = wdFirstPageFooterStory
      Case "wdFirstPageHeaderStory": locTypes(locCtr) = wdFirstPageHeaderStory
      Case "wdFootnoteContinuationNoticeStory": locTypes(locCtr) = wdFootnoteContinuationNoticeStory
      Case "wdFootnoteContinuationSeparatorStory": locTypes(locCtr) = wdFootnoteContinuationSeparatorStory
      Case "wdFootnoteSeparatorStory": locTypes(locCtr) = wdFootnoteSeparatorStory
      Case "wdFootnotesStory": locTypes(locCtr) = wdFootnotesStory
      Case "wdMainTextStory": locTypes(locCtr) = wdMainTextStory
      Case "wdPrimaryFooterStory": locTypes(locCtr) = wdPrimaryFooterStory
      Case "wdPrimaryHeaderStory": locTypes(locCtr) = wdPrimaryHeaderStory
      Case "wdTextFrameStory": locTypes(locCtr) = wdTextFrameStory
    End Select
  Next

  GetWordStoryTypes = locTypes
  
End Function
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Private procedures and properties.
Private Function pvtResolvePropertyScript(ByVal valObject As Object, ByVal valPropertyList As String) As Variant

  Dim locResult As Variant
  Dim locFirstElementEndPosition As Integer
  Dim locPropertyList As String
  
  locPropertyList = valPropertyList
  locFirstElementEndPosition = InStr(locPropertyList, ".")
  If locFirstElementEndPosition = 0 Then
    locResult = CallByName(valObject, locPropertyList, VbGet)
    pvtResolvePropertyScript = locResult
  Else
    Set locResult = CallByName(valObject, Left$(locPropertyList, locFirstElementEndPosition - 1), VbGet)
    pvtResolvePropertyScript = pvtResolvePropertyScript(locResult, Right$(locPropertyList, Len(locPropertyList) - locFirstElementEndPosition))
  End If
  
End Function
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

