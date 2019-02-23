VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ZtRegExp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Class ZtRegExp.
' It capsules VBScript regexp, overcomes some shortcomings of it and automatically checks validity of a regex pattern.
'
' Zotero Tools.
' This software is under Revised ('New') BSD license.
' Copyright © 2019, Olaf Ahrens. All rights reserved.
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Private constants.
Private Const PVT_CLASS_NAME As String = "Zotero Tools Regular Expressions"
Private Const PVT_ERROR_NR As Long = 10002
Private Const PVT_GROUP_REGEXP_PATTERN As String = "\((?!(?:\?:|\?=|\?!|[^\)\[]*?\]))(?:\?<([a-zA-Z0-9_]+?)>)?"
Private Const PVT_GROUP_REGEXP_REPLACEMENTS As String = "\\,\(,\[,\]"
Private Const PVT_REPLACE_REGEXP_PATTERN As String = "(?:[^$]|(?:[^$]|^)(?:\${2})+|^)(\$\{([a-zA-Z0-9_]+)\})"
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Private variables.
Private pvtDebugging As Boolean
Private pvtMessageDisplay As ZtIMessageDisplayable
Private pvtGroupRegExp As VBScript_RegExp_55.RegExp
Private pvtGroupRegExpReplacements() As String
Private pvtReplaceRegExp As VBScript_RegExp_55.RegExp
Private pvtRealRegExp As VBScript_RegExp_55.RegExp
Private pvtName As String
Private pvtGroups As Collection
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *



' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Constructor.
Private Sub Class_Initialize()

  pvtGroupRegExpReplacements = Split(PVT_GROUP_REGEXP_REPLACEMENTS, ",")
  Set pvtGroupRegExp = New VBScript_RegExp_55.RegExp
  With pvtGroupRegExp
    .Pattern = PVT_GROUP_REGEXP_PATTERN
    .Global = True
  End With
  Set pvtReplaceRegExp = New VBScript_RegExp_55.RegExp
  With pvtReplaceRegExp
    .Pattern = PVT_REPLACE_REGEXP_PATTERN
    .Global = True
  End With
  Set pvtRealRegExp = New VBScript_RegExp_55.RegExp
  Set pvtGroups = New Collection
  
End Sub
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Friend procedures and properties.
Friend Sub Initialize(ByVal valName As String, ByVal valPattern As String, ByVal valMessageDisplay As ZtIMessageDisplayable, _
  Optional ByVal valGlobal As Boolean = False, Optional ByVal valIgnoreCase As Boolean = False, Optional ByVal valMultiLine As Boolean = False)
  
  Dim locGroupPattern As String
  Dim locRealPattern As String
  Dim locCtr As Integer
  Dim locVBGroups As VBScript_RegExp_55.MatchCollection
  Dim locVBGroup As VBScript_RegExp_55.Match
  Dim locGroup As ZtRegGroup
  
  pvtName = valName
  
  ' Find unnamed and named capturing groups.
  locGroupPattern = valPattern
  For locCtr = 0 To UBound(pvtGroupRegExpReplacements)
    locGroupPattern = VBA.Strings.Replace(locGroupPattern, pvtGroupRegExpReplacements(locCtr), "X")
  Next
  Set pvtMessageDisplay = valMessageDisplay
     
  Set locVBGroups = pvtGroupRegExp.Execute(locGroupPattern)
  locCtr = 0
  Set locGroup = New ZtRegGroup
  locGroup.Initialize locCtr, "wholeMatch"
  pvtGroups.Add locGroup, locGroup.Name
  locCtr = locCtr + 1
  For Each locVBGroup In locVBGroups
    Set locGroup = New ZtRegGroup
    If IsEmpty(locVBGroup.SubMatches(0)) Then
      locGroup.Initialize locCtr, CStr(locCtr)
    Else
      locGroup.Initialize locCtr, locVBGroup.SubMatches(0)
    End If
    pvtGroups.Add locGroup, locGroup.Name
    locCtr = locCtr + 1
  Next
    
  ' Create new VBScript regex valid pattern and set regex for this pattern.
  locRealPattern = pvtGroupRegExp.Replace(valPattern, "(")
  With pvtRealRegExp
    .Pattern = locRealPattern
    .Global = valGlobal
    .IgnoreCase = valIgnoreCase
    .MultiLine = valMultiLine
  End With
  
  On Error GoTo OnError
  pvtRealRegExp.Test vbNullString
  
  Exit Sub
  
OnError:
  VBA.Interaction.MsgBox "The pattern" & vbNewLine & vbNewLine & _
                            Space$(5) & valPattern & vbNewLine & vbNewLine & _
                            "from the regex" & vbNewLine & vbNewLine & _
                            Space$(5) & valName & vbNewLine & vbNewLine & _
                            "isn't well formed. Please have a look at file 'ZtReadMe.txt' for the regex introduction.", _
                          vbCritical + vbOKOnly, _
                          PVT_CLASS_NAME & ": procedure Initialize"
                        
  Err.Raise PVT_ERROR_NR, PVT_CLASS_NAME & ": procedure Initialize", "Regex isn't well formed."

End Sub

Friend Function Test(ByVal valSourceString As String) As Boolean
  
  Test = pvtRealRegExp.Test(valSourceString)
  
End Function

Friend Function Execute(ByVal valSourceString As String) As Collection

  Dim locVBMatches As VBScript_RegExp_55.MatchCollection
  Dim locVBMatch As VBScript_RegExp_55.Match
  Dim locMatches As Collection
  Dim locMatch As ZtRegMatch
  
  Set locVBMatches = pvtRealRegExp.Execute(valSourceString)
  Set locMatches = New Collection
  
  For Each locVBMatch In locVBMatches
    Set locMatch = New ZtRegMatch
    locMatch.Initialize pvtGroups, locVBMatch
    locMatches.Add locMatch
  Next
  
  Set Execute = locMatches
  
End Function

Friend Function FirstMatch(ByVal valSourceString As String) As ZtRegMatch

  Dim locVBMatches As VBScript_RegExp_55.MatchCollection
  Dim locVBMatch As VBScript_RegExp_55.Match
  Dim locMatch As ZtRegMatch
  
  Set locVBMatches = pvtRealRegExp.Execute(valSourceString)
  
  For Each locVBMatch In locVBMatches
    Set locMatch = New ZtRegMatch
    locMatch.Initialize pvtGroups, locVBMatch
    Exit For
  Next
  
  Set FirstMatch = locMatch
  
End Function

' @valReplacement may be:
'  - a substituion:
'     - '${group name}'
'     - '$group number', group number is 1-based
'     - '$&' for the whole match
'  - a simple string
'     - the $-sign must be escaped by doubling
'  - combination of both
Friend Function Replace(ByVal valSourceString As String, ByVal valReplacement As String) As String

  Dim locMatch As VBScript_RegExp_55.Match
  Dim locMatches As VBScript_RegExp_55.MatchCollection
  Dim locGroup As ZtRegGroup
  Dim locReplacement As String
  
  locReplacement = valReplacement
  Set locMatches = pvtReplaceRegExp.Execute(locReplacement)
  For Each locMatch In locMatches
    Set locGroup = pvtGroups.Item(locMatch.SubMatches(1))
    locReplacement = VBA.Strings.Replace(locReplacement, locMatch.SubMatches(0), locGroup.InternalNr)
  Next

  Replace = pvtRealRegExp.Replace(valSourceString, locReplacement)
  
End Function

Friend Function Debugging(ByVal valMatch As ZtRegMatch, ByVal valRange As Word.Range) As ZtFMessageType

  Dim locResult As ZtFMessageType
  Dim locMessageText As ZtStringBuilder
  Dim locGroup As ZtRegGroup

  If pvtDebugging Then
    Set locMessageText = New ZtStringBuilder
    With locMessageText
      .Append "The '"
      .Append pvtName
      .Append "' regex has found the following groups in the selected range:"
      .Append vbNewLine
      .Append vbNewLine
      For Each locGroup In pvtGroups
        .Append locGroup.Name
        .Append ": "
        .Append "'"
        .Append valMatch.Groups(locGroup.Nr)
        .Append "'"
        .Append vbNewLine
        .Append vbNewLine
        .Append "With 'Disable' this debug information won't be shown anymore, otherwise procede or cancel."
      Next
    End With
    locResult = pvtMessageDisplay.Show(locMessageText.ToString, MessageOkDisableCancel + MessageInformation, True, valRange)
    If locResult = MessageDisable Then
      pvtDebugging = False
      locResult = MessageOk
    End If
  Else
    locResult = MessageOk
  End If
  
  If locResult = vbOK Then
    Debugging = MessageOk
  ElseIf locResult = vbCancel Then
    Debugging = MessageCancel
  End If

End Function
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
