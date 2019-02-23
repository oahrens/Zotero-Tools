VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ZtIStory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Interface ZtIStory.
' Interface for ZtMainStory and ZtStory.
'
' Zotero Tools.
' This software is under Revised ('New') BSD license.
' Copyright © 2019, Olaf Ahrens. All rights reserved.
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Interface procedures and properties.
' All interface definitions must be Public.
Public Sub Initialize(ByVal valConfig As ZtConfig, ByVal valMessageDisplay As ZtIMessageDisplayable, ByVal valProgress As ZtProgress, _
  ByVal valDocument As ZtDocument, ByVal valRange As Word.Range)
End Sub

Public Property Get CanHoldCitations() As Boolean
End Property

Public Sub KeepCitationGroups(ByRef refCitationGroupsOffset As Integer)
End Sub

Public Function RemoveInternalLinking() As ZtFMessageType
End Function

Public Function SetInternalLinking() As ZtFMessageType
End Function

Public Function AdjustPunctuation(ByRef refAffectedCt As Integer, ByRef refNotAffectedCt As Integer, ByRef refNotCorrectedCt As Integer) As ZtFMessageType
End Function

Public Sub ResolveCitationGroups(ByRef refAffectedCt As Integer)
End Sub

Public Sub UpdateAllFields()
End Sub

Public Property Get Range() As Word.Range
End Property

Public Property Get Document() As ZtDocument
End Property

Public Sub JoinCitationGroups(ByRef refAffectedCt As Integer, ByRef refNotAffectedCt As Integer, Optional ByVal valRange As Word.Range = Nothing)
End Sub
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


