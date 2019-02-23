Attribute VB_Name = "ZtSubprocedures"
Option Explicit


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Helping subprocedures.
'
' Zotero Tools.
' This software is under Revised ('New') BSD license.
' Copyright © 2019, Olaf Ahrens. All rights reserved.
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Public procedures.
Public Function AddLeadingZeros(ByVal valNr As Integer, ByVal valDigitsCt As Integer) As String

  AddLeadingZeros = String$(valDigitsCt - Len(CStr(valNr)), "0") & CStr(valNr)
  
End Function

Public Function Min(ByVal valNr0 As Double, ByVal valNr1 As Double) As Double

  If valNr0 < valNr1 Then
    Min = valNr0
  Else
    Min = valNr1
  End If

End Function

Public Function Max(ByVal valNr0 As Double, ByVal valNr1 As Double) As Double

  If valNr0 > valNr1 Then
    Max = valNr0
  Else
    Max = valNr1
  End If

End Function

Public Function Ceiling(ByVal valDouble As Double) As Long

  If Int(valDouble) = valDouble Then
    Ceiling = Int(valDouble)
  Else
    Ceiling = Int(valDouble) + 1
  End If
  
End Function

Public Sub ResetFind(ByVal valFind As Word.Find)

  With valFind
    .Text = vbNullString
    .Replacement.Text = vbNullString
    .Forward = True
    .Wrap = wdFindStop
    .Format = False
    .IgnorePunct = False
    .IgnoreSpace = False
    .MatchCase = False
    .MatchWholeWord = False
    .MatchWildcards = False
    .MatchSoundsLike = False
    .MatchAllWordForms = False
    With .Replacement
      .ClearFormatting
      .Text = vbNullString
    End With
  End With

End Sub

Public Function DigitsCt(ByVal valNr As Integer) As Integer

  If valNr = 0 Then
    DigitsCt = 0
  Else
    DigitsCt = ZtSubprocedures.Ceiling(Log(valNr) / Log(10))
  End If
  
End Function

Public Function ArrayContains(ByVal valMember As Variant, ByRef refArray As Variant) As Boolean

  Dim locCtr As Integer
  Dim locResult As Boolean
  
  For locCtr = LBound(refArray) To UBound(refArray)
    If valMember = refArray(locCtr) Then
      locResult = True
      Exit For
    End If
  Next
  
  ArrayContains = locResult
  
End Function

Public Function GetPath(ByVal valPathAndFileName As String) As String

  Dim locSeparatorPosition As Integer
  
  locSeparatorPosition = InStrRev(valPathAndFileName, Word.Application.PathSeparator)
  
  GetPath = Left$(valPathAndFileName, locSeparatorPosition - 1)
  
End Function

Public Function GetFileNameWithoutExtension(ByVal valPathAndFileName As String) As String

  Dim locSeparatorPosition As Integer
  Dim locDotPosition As Integer
  
  locSeparatorPosition = InStrRev(valPathAndFileName, Word.Application.PathSeparator)
  locDotPosition = InStrRev(valPathAndFileName, ".")
  
  GetFileNameWithoutExtension = Mid$(valPathAndFileName, locSeparatorPosition + 1, locDotPosition - locSeparatorPosition - 1)
  
End Function

Public Function GetExtension(ByVal valPathAndFileName As String) As String

  Dim locDotPosition As String
  
  locDotPosition = InStrRev(valPathAndFileName, ".")
  GetExtension = Right$(valPathAndFileName, Len(valPathAndFileName) - locDotPosition + 1)
  
End Function
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

