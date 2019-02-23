VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ZtRangeProcedures"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Class ZtRangeProcedures.
' It contains Word.Range specific procedures.
'
' Zotero Tools.
' This software is under Revised ('New') BSD license.
' Copyright © 2019, Olaf Ahrens. All rights reserved.
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Private variables.
Private pvtRange As Word.Range
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Friend procedures and properties.
Friend Sub SetRange(ByVal valRange As Word.Range)

  Set pvtRange = valRange ' Don't Duplicate!

End Sub

' Word.Range.MoveStartWhile incorporates a whole hyperlink, when comes upon it it.
' Word.Range.MoveStartWhile hangs if move mechanism comes upon a comment reference start or end.
Friend Function MoveStartWhile(ByVal valCharacters As String, Optional ByVal valCt As Long = wdForward) As Long

  Dim locCtr As Long
  Dim locIncludeHiddenText As Boolean
  
  With pvtRange
    locIncludeHiddenText = .TextRetrievalMode.IncludeHiddenText
    .TextRetrievalMode.IncludeHiddenText = True
    If valCt > 0 Then
      Do
        .MoveStart wdCharacter, 1
        locCtr = locCtr + 1
      Loop While InStr(valCharacters, ChrW$(AscW(.Characters(1)))) > 0 And locCtr <= valCt
      .MoveStart wdCharacter, -1
    ElseIf valCt < 0 Then
      Do
        .MoveStart wdCharacter, -1
        locCtr = locCtr - 1
      Loop While InStr(valCharacters, ChrW$(AscW(.Characters(1)))) > 0 And locCtr >= valCt
      .MoveStart wdCharacter, 1
      End If
    .TextRetrievalMode.IncludeHiddenText = locIncludeHiddenText
  End With
  
  MoveStartWhile = locCtr
    
End Function

' Word.Range.MoveStartWhile hangs if move mechanism comes upon a comment reference start or end.
Friend Function MoveEndWhile(ByVal valCharacters As String, Optional ByVal valCt As Long = wdForward) As Long

  Dim locCtr As Long
  Dim locIncludeHiddenText As Boolean
  
  With pvtRange
    locIncludeHiddenText = .TextRetrievalMode.IncludeHiddenText
    .TextRetrievalMode.IncludeHiddenText = True
    If valCt > 0 Then
      Do
        .MoveEnd wdCharacter, 1
        locCtr = locCtr + 1
      Loop While InStr(valCharacters, ChrW$(AscW(.Characters(.Characters.Count)))) > 0 And locCtr <= valCt
      .MoveEnd wdCharacter, -1
    ElseIf valCt < 0 Then
      Do
        .MoveEnd wdCharacter, -1
        locCtr = locCtr - 1
      Loop While InStr(valCharacters, ChrW$(AscW(.Characters(.Characters.Count)))) > 0 And locCtr >= valCt
      .MoveEnd wdCharacter, 1
    End If
    .TextRetrievalMode.IncludeHiddenText = locIncludeHiddenText
  End With
  
  MoveEndWhile = locCtr

End Function

Friend Function MoveStartUntil(ByVal valCharacters As String, Optional ByVal valCt As Long = wdForward) As Long

  Dim locCtr As Long
  Dim locIncludeHiddenText As Boolean
  
  With pvtRange
    locIncludeHiddenText = .TextRetrievalMode.IncludeHiddenText
    .TextRetrievalMode.IncludeHiddenText = True
    If valCt > 0 Then
      Do
        .MoveStart wdCharacter, 1
        locCtr = locCtr + 1
      Loop Until InStr(valCharacters, ChrW$(AscW(.Characters(1)))) > 0 Or locCtr = valCt + 1
      .MoveStart wdCharacter, -1
    ElseIf valCt < 0 Then
      Do
        .MoveStart wdCharacter, -1
        locCtr = locCtr - 1
      Loop Until InStr(valCharacters, ChrW$(AscW(.Characters(1)))) > 0 Or locCtr = valCt - 1
      .MoveStart wdCharacter, 1
    End If
    .TextRetrievalMode.IncludeHiddenText = locIncludeHiddenText
  End With
  
  MoveStartUntil = locCtr
    
End Function

Friend Function MoveEndUntil(ByVal valCharacters As String, Optional ByVal valCt As Long = wdForward) As Long

  Dim locCtr As Long
  Dim locIncludeHiddenText As Boolean
  
  With pvtRange
    locIncludeHiddenText = .TextRetrievalMode.IncludeHiddenText
    .TextRetrievalMode.IncludeHiddenText = True
    If valCt > 0 Then
      Do
        .MoveEnd wdCharacter, 1
        locCtr = locCtr + 1
      Loop Until InStr(valCharacters, ChrW$(AscW(.Characters(.Characters.Count)))) > 0 Or locCtr = valCt + 1
      .MoveEnd wdCharacter, -1
    ElseIf valCt < 0 Then
      Do
        .MoveEnd wdCharacter, -1
        locCtr = locCtr - 1
      Loop Until InStr(valCharacters, ChrW$(AscW(.Characters(.Characters.Count)))) > 0 Or locCtr = valCt - 1
      .MoveEnd wdCharacter, 1
    End If
    .TextRetrievalMode.IncludeHiddenText = locIncludeHiddenText
  End With
  
  MoveEndUntil = locCtr

End Function
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

