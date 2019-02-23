VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ZtJoinCitationGroupsSelection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements ZtIProcedure


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Class ZtJoinCitationGroupsSelection.
' This is the 'Join citation groups in selection' procedure.
'
' Zotero Tools.
' This software is under Revised ('New') BSD license.
' Copyright © 2019, Olaf Ahrens. All rights reserved.
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Private variables.
Private pvtName As String
Private pvtDescription As String
Private pvtStepsCt As Integer
Private pvtDocument As ZtDocument
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Constructor.
Private Sub Class_Initialize()

  pvtName = "Join citation groups in selection"
  pvtDescription = "This procedure joins ALL citation groups in the CURRENTLY SELECTED TEXT, e.g. from " & vbNewLine & _
                     vbTab & "'The first sentence.[1,2], [4-6] and [7] The next sentence.'" & vbNewLine & _
                       vbTab & "                _______________________________                  (selected text)" & vbNewLine & _
                       "to" & vbNewLine & _
                       vbTab & "'The first sentence.[1,2,4-7] and The next sentence.'." & vbNewLine & vbNewLine & _
                     "This procedure removes spaces, commas, and semicolons standing in between the citation groups automatically. " & _
                       "All other text will be shifted behind the new produced citation group. " & _
                       "This procedure is working for numeric style and author-year style citations. " & _
                       "You have to run Zoteros 'Refresh' command afterwards." & vbNewLine & vbNewLine & _
                     "Extended version of wh_fzj_ste's macro at https://forums.zotero.org/discussion/3262/single-to-multiple-citations-in-word-plugin/p3."
  
End Sub
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Private interface procedures and properties directing to Friend procedures and properties below.
Private Property Get ZtIProcedure_Name() As String

  ZtIProcedure_Name = Me.Name
  
End Property

Private Property Get ZtIProcedure_Description() As String

  ZtIProcedure_Description = Me.Description
  
End Property

Private Function ZtIProcedure_Start(ByVal valConfig As ZtConfig, ByVal valMessageDisplay As ZtIMessageDisplayable, ByVal valProgress As ZtProgress, _
  ByVal valAppPrepare As ZtIAppPreparable, ByVal valProcedureInitializer As ZtIProcedureInitializable, _
  ByVal valDocument As ZtDocument, ByRef refLicenseShown As Boolean) As ZtFMessageType

  ZtIProcedure_Start = Me.Start(valConfig, valMessageDisplay, valProgress, valAppPrepare, valProcedureInitializer, valDocument, refLicenseShown)
  
End Function
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *



' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Friend procedures and properties.
Friend Property Get Name() As String

  Name = pvtName
  
End Property

Friend Property Get Description() As String

  Description = pvtDescription
  
End Property

Friend Function Start(ByVal valConfig As ZtConfig, ByVal valMessageDisplay As ZtIMessageDisplayable, ByVal valProgress As ZtProgress, _
  ByVal valAppPrepare As ZtIAppPreparable, ByVal valProcedureInitializer As ZtIProcedureInitializable, _
  ByVal valDocument As ZtDocument, ByRef refLicenseShown As Boolean) As ZtFMessageType

  Dim locResult As ZtFMessageType

  Set pvtDocument = valDocument
  locResult = valProcedureInitializer.Start(valConfig, valMessageDisplay, valAppPrepare, pvtDocument, pvtName, refLicenseShown)
  If locResult = MessageOk Then
    valProgress.SetStepsCt pvtGetStepsCt
    locResult = pvtDocument.JoinCitationGroupsSelection
  End If
  valProcedureInitializer.Finish locResult, IIf(locResult = MessageOk, _
                                                "You have to run Zoteros 'Refresh' command now to take full effect. " & _
                                                  "Or, after clicking 'Procede', you can change the text selection " & _
                                                  "of the document and run this procedure again!", _
                                                vbNullString)
  
  Start = locResult
  
End Function
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Private procedures and properties.
Private Function pvtGetStepsCt() As Integer

  If pvtStepsCt = 0 Then
    pvtStepsCt = pvtDocument.GetJoinCitationGroupsSelectionStepsCt
  End If
  
  pvtGetStepsCt = pvtStepsCt
  
End Function
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

