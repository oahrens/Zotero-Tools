VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ZtAdjustPunctuation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements ZtIProcedure


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Class ZtAdjustPunctuation.
' This is the 'Adjust punctuation surrounding citation groups' procedure.
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

  pvtName = "Adjust punctuation surrounding citation groups"
  pvtDescription = "This procedure is working for numeric style and author-year style citations." & vbNewLine & _
                     "It switches citations between" & vbNewLine & _
                       vbTab & "- for author-year style citations:" & vbNewLine & _
                         vbTab & vbTab & "'This is a phenomenal, (Shakespeare 1565) and unbelivable sentence. (Raleigh 1554) This is not.'" & vbNewLine & _
                         vbTab & vbTab & "'This is a phenomenal,(Shakespeare 1565) and unbelivable sentence.(Raleigh 1554) This is not.'" & vbNewLine & _
                         vbTab & vbTab & "'This is a phenomenal (Shakespeare 1565), and unbelivable sentence (Raleigh 1554). This is not.'" & vbNewLine & _
                       vbTab & "- for superscript numeric style citations:" & vbNewLine & _
                         vbTab & vbTab & "'This is a phenomenal, (1) and unbelivable sentence. (2) This is not.'" & vbNewLine & _
                         vbTab & vbTab & "'This is a phenomenal,(1) and unbelivable sentence.(2) This is not.'" & vbNewLine & _
                         vbTab & vbTab & "'This is a phenomenal(1), and unbelivable sentence(2). This is not.'" & vbNewLine & _
                       vbTab & "- for non-superscript numeric style citations:" & vbNewLine & _
                         vbTab & vbTab & "'This is a phenomenal, (1) and unbelivable sentence. (2) This is not.'" & vbNewLine & _
                         vbTab & vbTab & "'This is a phenomenal (1), and unbelivable sentence (2). This is not.'" & vbNewLine & vbNewLine & _
                     "It depends on the settings of 'citationGroup.punctuationIsBefore' and 'citationGroup.spaceIsBefore' in ZtUserConfig.xml." & vbNewLine & vbNewLine & _
                     "Thanks ailintom for different ideas from https://github.com/ailintom/Cleanup-after-style-change-in-Zotero-for-Word."
  
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
  locResult = valProcedureInitializer.Start(valConfig, valMessageDisplay, valAppPrepare, pvtDocument, pvtName, refLicenseShown, True)
  If locResult = MessageOk Then
    valProgress.SetStepsCt pvtGetStepsCt
    locResult = pvtDocument.AdjustPunctuation
  End If
  valProcedureInitializer.Finish locResult, IIf(locResult = MessageOk, "You can also run this procedure with different settings again, for this purpose.", vbNullString)
  
  Start = locResult
  
End Function
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Private procedures and properties.
Private Function pvtGetStepsCt() As Integer

  If pvtStepsCt = 0 Then
    pvtStepsCt = pvtDocument.GetAdjustPunctuationStepsCt
  End If
  
  pvtGetStepsCt = pvtStepsCt
  
End Function
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

