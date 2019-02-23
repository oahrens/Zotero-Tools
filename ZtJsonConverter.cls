VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ZtJsonConverter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Class ZtJsonConverter.
' It converts a JSON string to:
'  - a Collection in case that the string represents an array
'  - a Scripting.Dictionary in case that the string represents an object
' This is a basic and slightly changed version from Tim Hall's code.
' It converts every JSON string and acts not CSL specific.
'
' Zotero Tools.
' This software is under Revised ('New') BSD license.
' Copyright © 2019, Olaf Ahrens. All rights reserved.
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' From the original code:
'    Part of VBA-JSON v2.3.0
'    (c) Tim Hall - https://github.com/VBA-tools/VBA-JSON
'
'    Part of JSON Converter for VBA
'
'    Errors:
'    10001 - JSON parse error
'
'    @class ZtJsonConverter
'    @author tim.hall.engr@gmail.com
'    @license MIT (http://www.opensource.org/licenses/mit-license.php)
'
'    Based originally on vba-json (with extensive changes)
'    BSD license included below
'
'    JSONLib, http://code.google.com/p/vba-json/
'
'    Copyright (c) 2013, Ryo Yokoyama
'    All rights reserved.
'
'    Redistribution and use in source and binary forms, with or without
'    modification, are permitted provided that the following conditions are met:
'     * Redistributions of source code must retain the above copyright
'       notice, this list of conditions and the following disclaimer.
'     * Redistributions in binary form must reproduce the above copyright
'       notice, this list of conditions and the following disclaimer in the
'       documentation and/or other materials provided with the distribution.
'     * Neither the name of the <organization> nor the
'       names of its contributors may be used to endorse or promote products
'       derived from this software without specific prior written permission.
'
'    THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND
'    ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
'    WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
'    DISCLAIMED. IN NO EVENT SHALL <COPYRIGHT HOLDER> BE LIABLE FOR ANY
'    DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
'    (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
'    LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
'    ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
'    (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
'    SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Private constants.
Private Const PVT_CLASS_NAME As String = "Zotero Tools JSON converter"
Private Const PVT_ERROR_NR As Long = 10001
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Private variables.
Private pvtJsonString As String
Private pvtJsonLength As Long
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Friend procedures and properties.

' Convert JSON string to object (Dictionary/Collection)
' @method Parse
' @param {String} valJsonString
' @return {Object} (Dictionary or Collection)
' @throws 10001 - JSON parse error
Friend Function Parse(ByVal valJsonString As String) As Object

  Dim locIndex As Long
  
  locIndex = 1

  ' Remove vbCr, vbLf, and vbTab from valJsonString.
  pvtJsonString = Replace(Replace(Replace(valJsonString, vbCr, vbNullString), vbLf, vbNullString), vbTab, vbNullString)
  pvtJsonLength = Len(pvtJsonString)

  pvtSkipSpaces locIndex
  Select Case pvtJsonSubStr(locIndex, 1)
    Case "{"
      Set Parse = pvtParseObject(locIndex)
    Case "["
      Set Parse = pvtParseArray(locIndex)
    Case Else
      ' Error: Invalid JSON string
      Err.Raise PVT_ERROR_NR, _
                PVT_CLASS_NAME & ": procedure Parse", _
                pvtGetErrorMessage(locIndex, "Expecting '{' or '['")
  End Select
  
End Function
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Private procedures and properties.
Private Function pvtParseObject(ByRef refIndex As Long) As Scripting.Dictionary

  Dim locKey As String
  Dim locChar As String
  Dim locNextChar As String

  Set pvtParseObject = New Scripting.Dictionary
  pvtSkipSpaces refIndex
  refIndex = refIndex + 1

  Do
    pvtSkipSpaces refIndex
    locChar = pvtJsonSubStr(refIndex, 1)
    If locChar = "}" Then
      refIndex = refIndex + 1
      Exit Function
    ElseIf locChar = "," Then
      refIndex = refIndex + 1
      pvtSkipSpaces refIndex
    End If

    locKey = pvtParseKey(refIndex)
    locNextChar = pvtPeek(refIndex)
    If locNextChar = "[" Or locNextChar = "{" Then
      Set pvtParseObject.Item(locKey) = pvtParseValue(refIndex)
    Else
      pvtParseObject.Item(locKey) = pvtParseValue(refIndex)
    End If
  Loop
  
End Function

Private Function pvtParseArray(ByRef refIndex As Long) As Collection

  Dim locChar As String
  
  Set pvtParseArray = New Collection

  pvtSkipSpaces refIndex
  refIndex = refIndex + 1

  Do
    pvtSkipSpaces refIndex
    locChar = pvtJsonSubStr(refIndex, 1)
    If locChar = "]" Then
      refIndex = refIndex + 1
      Exit Function
    ElseIf locChar = "," Then
      refIndex = refIndex + 1
      pvtSkipSpaces refIndex
    End If

    pvtParseArray.Add pvtParseValue(refIndex)
  Loop
  
End Function

Private Function pvtParseValue(ByRef refIndex As Long) As Variant

  pvtSkipSpaces refIndex
  
  Select Case pvtJsonSubStr(refIndex, 1)
    Case "{"
      Set pvtParseValue = pvtParseObject(refIndex)
    Case "["
      Set pvtParseValue = pvtParseArray(refIndex)
    Case """", "'"
      pvtParseValue = pvtParseString(refIndex)
    Case Else
      If pvtJsonSubStr(refIndex, 4) = "true" Then
        pvtParseValue = True
        refIndex = refIndex + 4
      ElseIf pvtJsonSubStr(refIndex, 5) = "false" Then
        pvtParseValue = False
        refIndex = refIndex + 5
      ElseIf pvtJsonSubStr(refIndex, 4) = "null" Then
        pvtParseValue = Null
        refIndex = refIndex + 4
      ElseIf InStr("+-0123456789", pvtJsonSubStr(refIndex, 1)) > 0 Then
        pvtParseValue = pvtParseDecimal(refIndex)
      Else
        Err.Raise PVT_ERROR_NR, _
                  PVT_CLASS_NAME & ": procedure pvtPaseValue", _
                  pvtGetErrorMessage(refIndex, "Expecting 'STRING', 'NUMBER', null, true, false, '{', or '['")
      End If
  End Select
  
End Function

Private Function pvtParseString(ByRef refIndex As Long) As String

  Dim locQuote As String
  Dim locChar As String
  Dim locCode As String
  Dim locString As ZtStringBuilder

  Set locString = New ZtStringBuilder
  pvtSkipSpaces refIndex

  ' Store opening quote to look for matching closing quote
  locQuote = pvtJsonSubStr(refIndex, 1)
  refIndex = refIndex + 1

  Do While refIndex > 0 And refIndex <= pvtJsonLength
    locChar = pvtJsonSubStr(refIndex, 1)

    Select Case locChar
      Case "\"
        ' Escaped string, \\, or \/
        refIndex = refIndex + 1
        locChar = pvtJsonSubStr(refIndex, 1)

        Select Case locChar
          Case """", "\", "/", "'"
            locString.Append locChar
            refIndex = refIndex + 1
          Case "b"
            locString.Append vbBack
            refIndex = refIndex + 1
          Case "f"
            locString.Append vbFormFeed
            refIndex = refIndex + 1
          Case "n"
            locString.Append vbNewLine
            refIndex = refIndex + 1
          Case "r"
            locString.Append vbCr
            refIndex = refIndex + 1
          Case "t"
            locString.Append vbTab
            refIndex = refIndex + 1
          Case "u"
            ' Unicode character escape (e.g. \u00a9 = Copyright)
            refIndex = refIndex + 1
            locCode = pvtJsonSubStr(refIndex, 4)
            locString.Append ChrW$(val("&h" + locCode))
            refIndex = refIndex + 4
        End Select
      Case locQuote
        pvtParseString = locString.ToString
        refIndex = refIndex + 1
        Exit Function
      Case Else
        locString.Append locChar
        refIndex = refIndex + 1
    End Select
  Loop
  
End Function

Private Function pvtParseDecimal(ByRef refIndex As Long) As Variant

  Dim locChar As String
  Dim locString As String
  Dim locIsLargeDecimal As Boolean

  pvtSkipSpaces refIndex

  Do While refIndex > 0 And refIndex <= pvtJsonLength
    locChar = pvtJsonSubStr(refIndex, 1)

    If InStr("+-0123456789.eE", locChar) > 0 Then
      ' Unlikely to have massive number, so use simple append rather than buffer here
      locString = locString & locChar
      refIndex = refIndex + 1
    Else
      ' Excel only stores 15 significant digits, so any numbers larger than that are truncated
      ' This can lead to issues when BIGINT's are used (e.g. for Ids or Credit Cards), as they will be invalid above 15 digits
      ' See: http://support.microsoft.com/kb/269370
      '
      ' Fix: Parse -> String, Convert -> String longer than 15/16 characters containing only numbers and decimal points -> Number
      ' (decimal doesn't factor into significant digit count, so if present check for 15 digits + decimal = 16)
      If Len(locString) >= 16 Then
        If InStr(locString, ".") > 0 Then
          locIsLargeDecimal = True
        ElseIf Len(locString) >= 17 Then
          locIsLargeDecimal = True
        End If
      End If
      ' Val does not use regional settings, so guard for comma is not needed
      pvtParseDecimal = val(locString)
      Exit Function
    End If
  Loop
  
End Function

Private Function pvtParseKey(ByRef refIndex As Long) As String

  ' Parse key with single or double quotes
  If InStr("""'", pvtJsonSubStr(refIndex, 1)) > 0 Then
    pvtParseKey = pvtParseString(refIndex)
  Else
    Err.Raise PVT_ERROR_NR, _
              PVT_CLASS_NAME & ": procedure pvtParseKey", _
              pvtGetErrorMessage(refIndex, "Expecting '""' or '''")
  End If

  ' Check for colon and skip if present or throw if not present
  pvtSkipSpaces refIndex
  If pvtJsonSubStr(refIndex, 1) <> ":" Then
    Err.Raise PVT_ERROR_NR, _
              PVT_CLASS_NAME & ": procedure pvtParseKey", _
              pvtGetErrorMessage(refIndex, "Expecting ':'")
  Else
    refIndex = refIndex + 1
  End If
  
End Function

Private Function pvtPeek(ByVal valIndex As Long, Optional ByVal valCharactersCt As Long = 1) As String
  ' "Peek" at the next number of characters without incrementing valIndex (ByVal instead of ByRef)
  
  pvtSkipSpaces valIndex
  pvtPeek = pvtJsonSubStr(valIndex, valCharactersCt)
  
End Function

Private Sub pvtSkipSpaces(ByRef refIndex As Long)
  ' Increment index to skip over spaces

  Do While refIndex > 0 And refIndex <= pvtJsonLength And pvtJsonSubStr(refIndex, 1) = " "
    refIndex = refIndex + 1
  Loop
  
End Sub

Private Function pvtGetErrorMessage(ByVal valIndex As Long, ByVal valMessage As String) As String
  ' Provide detailed parse error message, including details of where and what occurred
  '
  ' Example:
  ' Error parsing JSON:
  ' {"abcde":True}
  '          ^
  ' Expecting 'STRING', 'NUMBER', null, true, false, '{', or '['

  Dim locStartIndex As Long
  Dim locEndIndex As Long

  ' Include 10 characters before and after error (if possible)
  locStartIndex = valIndex - 10
  locEndIndex = valIndex + 10
  If locStartIndex <= 0 Then
    locStartIndex = 1
  End If
  If locEndIndex > pvtJsonLength Then
    locEndIndex = pvtJsonLength
  End If

  pvtGetErrorMessage = "Error parsing JSON:" & vbNewLine & _
                       pvtJsonSubStr(locStartIndex, locEndIndex - locStartIndex + 1) & vbNewLine & _
                       Space$(valIndex - locStartIndex) & "^" & vbNewLine & _
                       valMessage
End Function

Private Function pvtJsonSubStr(ByVal valStart As Long, ByVal valLength As Long) As String

  pvtJsonSubStr = Mid$(pvtJsonString, valStart, valLength)
  
End Function
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
