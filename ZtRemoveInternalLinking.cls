VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ZtRemoveInternalLinking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements ZtIProcedure


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Class ZtRemoveInternalLinking.
' This is the 'Remove internal linking between citations and references' procedure.
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
Private pvtToResume As Boolean
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Constructor.
Private Sub Class_Initialize()

  pvtName = "Remove internal linking between citations and references"
  pvtDescription = "This procedure removes all changes done by the procedure 'Set internal linking between citations and references'." & vbNewLine & vbNewLine & _
                     "This procedure is only working for numeric style citations. " & vbNewLine & vbNewLine & _
                     "Thanks pavelninin for different ideas from https://github.com/pavelninin/zotero_two_way_hyperlinking/blob/master/zotero_two_way_hyperlinking.vbs."
                     
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
    locResult = pvtDocument.RemoveInternalLinking
  End If
  valProcedureInitializer.Finish locResult, , pvtToResume
  
  Start = locResult
  
End Function

Friend Property Let ToResume(ByVal valToResume As Boolean)

  pvtToResume = valToResume
  
End Property
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Private procedures and properties.
Private Function pvtGetStepsCt() As Integer

  If pvtStepsCt = 0 Then
    pvtStepsCt = pvtDocument.GetRemoveInternalLinkingStepsCt
  End If
  
  pvtGetStepsCt = pvtStepsCt
  
End Function
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

