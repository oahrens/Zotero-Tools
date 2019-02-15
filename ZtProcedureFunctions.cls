VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ZtProcedureFunctions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Class ZtProcedureFunctions.
' Manages Word application and document preparation and shows starting and finishing messages of the procedures.
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Private variables.
Private pvtConfig As ZtConfig
Private pvtStartForm As ZtStartForm
Private pvtDocument As ZtDocument
Private pvtUndo As Word.UndoRecord
Private pvtIsPrepared As Boolean
Private pvtShouldBePrepared As Boolean
Private pvtStatusBar As Boolean
Private pvtPagination As Boolean
Private pvtCheckGrammarType As Boolean
Private pvtCheckGrammarSpelling  As Boolean
Private pvtCheckSpellingType As Boolean
Private pvtAutoWordSelection As Boolean
Private pvtPasteWordSpacing As Boolean
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Friend procedures and properties.
Friend Function Initialize(ByVal valConfig As ZtConfig, ByVal valStartForm As ZtStartForm, ByVal valDocument As ZtDocument, _
  ByVal valProcedureName As String, ByRef refLicenceShown As Boolean, Optional ByVal valTakesTime As Boolean = False, Optional ByVal valResumed As Boolean = False) As ZtFMessageType
  
  Dim locResult As ZtFMessageType
  
  Set pvtStartForm = valStartForm
  Set pvtDocument = valDocument
  
  If valResumed Then
    locResult = MessageOk
  Else
    Set pvtConfig = valConfig
  
    If refLicenceShown Then
      locResult = MessageOk
    Else
      locResult = pvtStartForm.MessageDisplay.Show("This software is under FreeBSD licence. Copyright (c) 2019, Olaf Ahrens. " & _
                                                       "Please, read licence text and introduction in 'ZtReadMe.txt' carefully." & vbNewLine & vbNewLine & _
                                                     "This software may destroy your document irrevocable, or Zotero won't be able to run on your document any more. " & _
                                                       "But, a backup copy will be saved with every procedure been run! " & _
                                                     "Please, ensure that you have adapted the setting of the macro for the chosen citation and bibliography style!" & vbNewLine & vbNewLine & _
                                                     "Do you want to proceed?", _
                                                   MessageQuestion + MessageOkCancel)
    End If
    If locResult = MessageOk Then
      refLicenceShown = True
      locResult = pvtStartForm.MessageDisplay.Show("This procedure will edit document" & vbNewLine & _
                                                     vbTab & pvtDocument.Name & vbNewLine & _
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
      pvtDocument.SaveBackupCopy
    End If
    Set pvtUndo = Word.Application.UndoRecord
    pvtUndo.StartCustomRecord valProcedureName
  
    If Not valResumed Then
      pvtShouldBePrepared = Not pvtConfig.User.Macro.Debugging
      pvtDocument.Prepare
      With Word.Application.Options
        pvtPasteWordSpacing = .PasteAdjustWordSpacing
        .PasteAdjustWordSpacing = False
      End With
      Me.Prepare
    End If
  Else
    locResult = Messageno
  End If
  
  Initialize = locResult
  
End Function

Friend Sub Finish(ByVal valResult As ZtFMessageType, Optional ByVal valExtraClosingText As String = vbNullString, Optional ByVal valResume As Boolean = False)
  
  If Not valResume Then
    Me.Unprepare
    Word.Application.Options.PasteAdjustWordSpacing = pvtPasteWordSpacing
    pvtDocument.Unprepare
    pvtShouldBePrepared = False
    
    If valResult = MessageOk Then
      pvtStartForm.MessageDisplay.Show "Done." & vbNewLine & vbNewLine & _
                                         "You can undo all actions of the procedure by Word's undo function. " & valExtraClosingText, _
                                       MessageInformation + MessageOk, _
                                       False
    ElseIf valResult = Messageno Then
      pvtStartForm.MessageDisplay.Show "The document hasn't been changed.", _
                                       MessageInformation + Messageno, _
                                       False
    Else
      pvtStartForm.MessageDisplay.Show "Sorry, it seems there are some dicrepancies to correct before running this procedure again to take full effect.", _
                                       MessageExclamation + MessageCancel, _
                                       False
    End If
  End If
  If Not pvtUndo Is Nothing Then
    pvtUndo.EndCustomRecord
  End If
  
End Sub

Friend Sub Prepare()

  If Not pvtIsPrepared And pvtShouldBePrepared Then
    With Word.Application
      If pvtConfig.User.Macro.WordIsInvisibleWhileOperation Then
        .Visible = False
      End If
      .ScreenUpdating = False
      pvtStatusBar = .DisplayStatusBar
      .DisplayStatusBar = False
      With .Options
        pvtPagination = .Pagination
        .Pagination = False
        pvtCheckGrammarType = .CheckGrammarAsYouType
        .CheckGrammarAsYouType = False
        pvtCheckGrammarSpelling = .CheckGrammarWithSpelling
        .CheckGrammarWithSpelling = False
        pvtCheckSpellingType = .CheckSpellingAsYouType
        .CheckSpellingAsYouType = False
        pvtAutoWordSelection = .AutoWordSelection
        .AutoWordSelection = False
      End With
    End With
    pvtIsPrepared = True
  End If

End Sub

Friend Sub Unprepare()

  If pvtIsPrepared Then
    With Word.Application
      .ScreenUpdating = True
      .DisplayStatusBar = pvtStatusBar
      With .Options
        .Pagination = pvtPagination
        .CheckGrammarAsYouType = pvtCheckGrammarType
        .CheckGrammarWithSpelling = pvtCheckGrammarSpelling
        .CheckSpellingAsYouType = pvtCheckSpellingType
        .AutoWordSelection = pvtAutoWordSelection
      End With
      If pvtConfig.User.Macro.WordIsInvisibleWhileOperation Then
        .Visible = True
      End If
    End With
    pvtIsPrepared = False
  End If

End Sub
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
