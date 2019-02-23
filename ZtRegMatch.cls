VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ZtRegMatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Class ZtRegMatch.
' It capsules VBScript match and enables named groups in VBScript regexp.
'
' Zotero Tools.
' This software is under Revised ('New') BSD license.
' Copyright © 2019, Olaf Ahrens. All rights reserved.
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Private variables.
Private pvtGroups As Collection
Private pvtMatch As VBScript_RegExp_55.Match
Private pvtInternalNrPattern As String
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Constructor.
Private Sub Class_Initialize()
  
  pvtInternalNrPattern = "$#"
  
End Sub
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Friend procedures and properties.
Friend Sub Initialize(ByVal valGroups As Collection, ByVal valMatch As VBScript_RegExp_55.Match)

  Set pvtGroups = valGroups
  Set pvtMatch = valMatch

End Sub

' @valGroupNameOrNr:
'  - 0 = the whole match
'  - 1, 2, ... = the first, second, ... named or unnamed group of the match
'  - 'group name' = the specified group of the match
Friend Property Get Groups(ByVal valGroupNameOrNr As Variant) As String

  Dim locRegGroup As ZtRegGroup
  Dim locGroupNameOrNr As Variant

  locGroupNameOrNr = valGroupNameOrNr
  If locGroupNameOrNr Like pvtInternalNrPattern Then
    locGroupNameOrNr = Right$(locGroupNameOrNr, 1)
  End If
  If IsNumeric(locGroupNameOrNr) Then
    ' VBA collections are 1-based.
    locGroupNameOrNr = locGroupNameOrNr + 1
  End If
  Set locRegGroup = pvtGroups.Item(locGroupNameOrNr)

  If locRegGroup.Nr = 0 Then
    Groups = pvtMatch.Value
  Else
    Groups = pvtMatch.SubMatches(locRegGroup.Nr - 1)
  End If

End Property

Friend Property Get FirstIndex() As Long

  FirstIndex = pvtMatch.FirstIndex
  
End Property

Friend Property Get Length() As Long

  Length = pvtMatch.Length
  
End Property

Friend Property Get Value() As String

  Value = pvtMatch.Value

End Property
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
