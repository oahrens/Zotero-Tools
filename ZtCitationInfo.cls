VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ZtCitationInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Class ZtCitationInfo.
' It preserves reference (bibliographic) information of a single citation inside a citation group for resolving it.
'
' Zotero Tools.
' This software is under Revised ('New') BSD license.
' Copyright © 2019, Olaf Ahrens. All rights reserved.
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Private variables.
Private pvtConfig As ZtConfig
Private pvtType As String
Private pvtAuthors As String
Private pvtEditors As String
Private pvtTitle As String
Private pvtContainerTitle As String
Private pvtJournalAbbreviation As String
Private pvtVolume As String
Private pvtIssue As String
Private pvtPage As String
Private pvtYear As Integer
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Friend procedures and properties.
Friend Sub Initialize(ByVal valConfig As ZtConfig, ByVal valData As Scripting.Dictionary)
  
  Set pvtConfig = valConfig
  pvtType = pvtGetString(valData, "type")
  pvtAuthors = pvtGetNames(valData, "author")
  pvtEditors = pvtGetNames(valData, "editor")
  pvtJournalAbbreviation = pvtGetString(valData, "journalAbbreviation")
  pvtContainerTitle = pvtGetString(valData, "container-title")
  pvtTitle = pvtGetString(valData, "title")
  pvtVolume = pvtGetString(valData, "volume")
  pvtIssue = pvtGetString(valData, "issue")
  pvtPage = pvtGetString(valData, "page")
  pvtYear = pvtGetYear(valData, "issued")

End Sub

Friend Function ToString() As String

  Dim locString As ZtStringBuilder
  
  Set locString = New ZtStringBuilder
  
  With locString
  
    If pvtType = "article-journal" Then
    
      If Len(pvtAuthors) > 0 Then
        .Append pvtAuthors
        .Append pvtAppendDot(pvtAuthors)
      End If
      If Len(pvtTitle) > 0 Then
        .Append pvtTitle
        .Append pvtAppendDot(pvtTitle)
      End If
      If Len(pvtJournalAbbreviation) > 0 Then
        .Append pvtRemoveDots(pvtJournalAbbreviation)
        .Append ". "
      ElseIf Len(pvtContainerTitle) > 0 Then
        .Append pvtRemoveDots(pvtContainerTitle)
        .Append ". "
      End If
      If pvtYear > 0 Then
        .Append CStr(pvtYear)
        If Len(pvtVolume) > 0 Then
          .Append ";"
          .Append pvtVolume
          If Len(pvtIssue) > 0 Then
            .Append "("
            .Append pvtIssue
            .Append ")"
          End If
          If Len(pvtPage) > 0 Then
            .Append ":"
            .Append pvtPage
          End If
        ElseIf Len(pvtIssue) > 0 Then
          .Append ";"
          .Append "("
          .Append pvtIssue
          .Append ")"
          If Len(pvtPage) > 0 Then
            .Append ":"
            .Append pvtPage
          End If
        ElseIf Len(pvtPage) > 0 Then
          .Append ";"
          .Append pvtPage
        End If
        .Append ". "
      End If
      
    ElseIf pvtType = "chapter" Then
      
      If Len(pvtAuthors) > 0 Then
        .Append pvtAuthors
        .Append pvtAppendDot(pvtAuthors)
      End If
      If Len(pvtTitle) > 0 Then
        .Append pvtTitle
        .Append pvtAppendDot(pvtTitle)
      End If
      If Len(pvtEditors) > 0 Or Len(pvtContainerTitle) > 0 Then
        .Append "In: "
        If Len(pvtEditors) > 0 Then
          .Append pvtEditors
          .Append " (eds.). "
        End If
        If Len(pvtContainerTitle) > 0 Then
          .Append pvtContainerTitle
          .Append pvtAppendDot(pvtContainerTitle)
        End If
        If pvtYear > 0 Then
          .Append CStr(pvtYear)
          .Append ". "
        End If
        If Len(pvtPage) > 0 Then
          .Append "p. "
          .Append pvtPage
          .Append ". "
        End If
      End If
  
    Else
      
      If Len(pvtAuthors) > 0 Then
        .Append pvtAuthors
        .Append ". "
      End If
      If Len(pvtTitle) > 0 Then
        .Append pvtTitle
        .Append ". "
      End If
      If pvtYear > 0 Then
        .Append CStr(pvtYear)
        .Append ". "
      End If
      
    End If
    
    ToString = Trim$(.ToString)
    
  End With
  
End Function
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Private procedures and properties.
Private Function pvtGetString(ByVal valData As Scripting.Dictionary, ByVal valItemName As String) As String
  
  If Not IsEmpty(valData.Item(valItemName)) Then
    pvtGetString = valData.Item(valItemName)
    ' Looping for nested tags.
    Do While pvtConfig.Final.Citation.HtmlTagsRegExp.Test(pvtGetString)
      pvtGetString = pvtConfig.Final.Citation.HtmlTagsRegExp.Replace(pvtGetString, "${replacement}")
    Loop
  End If

End Function

Private Function pvtGetYear(ByVal valData As Scripting.Dictionary, ByVal valItemName As String) As Integer

  Dim locDate As Scripting.Dictionary
  Dim locDateParts As Collection
  
  If Not IsEmpty(valData.Item(valItemName)) Then
    Set locDate = valData.Item(valItemName)
    If Not IsEmpty(locDate.Item("date-parts")) Then
      Set locDateParts = locDate.Item("date-parts")
      pvtGetYear = CInt(locDateParts.Item(1).Item(1))
    End If
  End If
  
End Function

Private Function pvtGetNames(ByVal valData As Scripting.Dictionary, ByVal valItemName As String) As String

  Dim locNames As Collection
  Dim locName As Scripting.Dictionary
  Dim locGivenName As String
  Dim locString As ZtStringBuilder
  Dim locNameCtr As Integer
  
  If Not IsEmpty(valData.Item(valItemName)) Then
  
    Set locString = New ZtStringBuilder
    Set locNames = valData.Item(valItemName)
    
    With locString
      For Each locName In locNames
        If Not IsEmpty(locName.Item("family")) Then
          .Append locName.Item("family")
          If pvtConfig.User.ResolveCitationStyle.GivenNameStyle <> GivenNameNone Then
            If Not IsEmpty(locName.Item("given")) Then
              If pvtConfig.User.ResolveCitationStyle.GivenNameStyle = GivenNameFull Then
                locGivenName = Trim$(locName.Item("given"))
              Else
                locGivenName = pvtConfig.Final.Citation.GivenNameRegExp.Replace(Trim$(locName.Item("given")), "${initial}")
              End If
              If Len(locGivenName) > 0 Then
                .Append " "
                .Append locGivenName
              End If
            End If
          End If
          .Append ", "
          locNameCtr = locNameCtr + 1
        ElseIf Not IsEmpty(locName.Item("literal")) Then
          .Append locName.Item("literal")
          .Append ", "
          locNameCtr = locNameCtr + 1
        End If
        If pvtConfig.User.ResolveCitationStyle.MaxNamesCt > 0 And locNameCtr = pvtConfig.User.ResolveCitationStyle.MaxNamesCt Then Exit For
      Next
      .RemoveFromEnd 2
      If pvtConfig.User.ResolveCitationStyle.MaxNamesCt > 0 And locNameCtr < locNames.Count Then
        .Append " et al"
      End If
      
      pvtGetNames = .ToString
      
    End With
    
  End If
  
End Function

Private Function pvtAppendDot(ByVal valString As String) As String

  If Right$(valString, 1) Like pvtConfig.Basic.Patterns.Punctuation Then
    pvtAppendDot = " "
  Else
    pvtAppendDot = ". "
  End If
  
End Function

Private Function pvtRemoveDots(ByVal valString As String) As String

  pvtRemoveDots = Replace(valString, ".", vbNullString)
  
End Function
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


