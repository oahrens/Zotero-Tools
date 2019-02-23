VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ZtProcedureInitializer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements ZtIProcedureInitializable


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Class ZtProcedureInitializer.
' Manages Word application and document preparation and shows starting and finishing messages of the procedures.
'
' Zotero Tools.
' This software is under Revised ('New') BSD license.
' Copyright © 2019, Olaf Ahrens. All rights reserved.
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Private variables.
Private pvtConfig As ZtConfig
Private pvtMessageDisplay As ZtIMessageDisplayable
Private pvtUndo As Word.UndoRecord
Private pvtAppPrepare As ZtIAppPreparable
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Private interface procedures and properties directing to Friend procedures and properties below.
Private Function ZtIProcedureInitializable_Start(ByVal valConfig As ZtConfig, ByVal valMessageDisplay As ZtIMessageDisplayable, ByVal valAppPrepare As ZtIAppPreparable, _
  ByVal valDocument As ZtDocument, ByVal valProcedureName As String, ByRef refLicenseShown As Boolean, Optional ByVal valTakesTime As Boolean = False, _
  Optional ByVal valResumed As Boolean = False) As ZtFMessageType
  
  ZtIProcedureInitializable_Start = Me.Start(valConfig, valMessageDisplay, valAppPrepare, valDocument, valProcedureName, refLicenseShown, valTakesTime, valResumed)

End Function

Private Sub ZtIProcedureInitializable_Finish(ByVal valResult As ZtFMessageType, Optional ByVal valExtraClosingText As String = vbNullString, Optional ByVal valResume As Boolean = False)
  
  Me.Finish valResult, valExtraClosingText, valResume

End Sub
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Friend procedures and properties.
Friend Function Start(ByVal valConfig As ZtConfig, ByVal valMessageDisplay As ZtIMessageDisplayable, ByVal valAppPrepare As ZtIAppPreparable, _
  ByVal valDocument As ZtDocument, ByVal valProcedureName As String, ByRef refLicenseShown As Boolean, Optional ByVal valTakesTime As Boolean = False, _
  Optional ByVal valResumed As Boolean = False) As ZtFMessageType
  
  Dim locResult As ZtFMessageType
  
  Set pvtMessageDisplay = valMessageDisplay
  Set pvtAppPrepare = valAppPrepare
  
  If valResumed Then
    locResult = MessageOk
  Else
    Set pvtConfig = valConfig
  
    If refLicenseShown Then
      locResult = MessageOk
    Else
      locResult = pvtMessageDisplay.Show("This software is under Revised ('New') BSD license. Copyright © 2019, Olaf Ahrens. All rights reserved. " & _
                                             "Please, read the license text and introduction in 'ZtReadMe.txt' carefully." & vbNewLine & vbNewLine & _
                                           "This software may destroy your document irrevocable, or Zotero won't be able to run on your document any more. " & _
                                             "But, a backup copy will be saved with every procedure been run! " & _
                                           "Please, ensure that you have adapted the setting of the macro for the chosen citation and bibliography style!" & vbNewLine & vbNewLine & _
                                           "Do you want to proceed?", _
                                         MessageQuestion + MessageOkCancel, _
                                         False)
    End If
    If locResult = MessageOk Then
      refLicenseShown = True
      locResult = pvtMessageDisplay.Show("This procedure will edit document" & vbNewLine & _
                                           vbTab & valDocument.Name & vbNewLine & _
                                           "with style" & vbNewLine & _
                                           vbTab & pvtConfig.User.Macro.StylePresetName & vbNewLine & _
                                           IIf(valTakesTime, "Be patient, it may take a few minutes!" & vbNewLine, vbNullString) & _
                                           "At first, this document and an additional backup copy will be saved." & vbNewLine & vbNewLine & _
                                           "Do you want to proceed?", _
                                         MessageQuestion + MessageOkCancel)
                       
    End If
  End If
  
  If locResult = MessageOk Then
    If Not valResumed Then
      valDocument.SaveBackupCopy
    End If
    Set pvtUndo = Word.Application.UndoRecord
    pvtUndo.StartCustomRecord valProcedureName
  
    If Not valResumed Then
      pvtAppPrepare.Prepare
    End If
  Else
    locResult = MessageNo
  End If
  
  Start = locResult
  
End Function

Friend Sub Finish(ByVal valResult As ZtFMessageType, Optional ByVal valExtraClosingText As String = vbNullString, Optional ByVal valResume As Boolean = False)
  
  If Not valResume Then
    pvtAppPrepare.Unprepare
    
    If valResult = MessageOk Then
      pvtMessageDisplay.Show "Done." & vbNewLine & vbNewLine & _
                               "You can undo all actions of the procedure by Word's undo function. " & valExtraClosingText, _
                             MessageInformation + MessageOk, _
                             False
    ElseIf valResult = MessageNo Then
      pvtMessageDisplay.Show "The document hasn't been changed.", _
                             MessageInformation + MessageNo, _
                             False
    Else
      pvtMessageDisplay.Show "Sorry, it seems there are some dicrepancies to correct before running this procedure again to take full effect.", _
                             MessageExclamation + MessageCancel, _
                             False
    End If
  End If
  If Not pvtUndo Is Nothing Then
    pvtUndo.EndCustomRecord
  End If
  
End Sub
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
