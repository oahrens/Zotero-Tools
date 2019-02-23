VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ZtAppPreparer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements ZtIAppPreparable


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Class ZtAppPreparer.
' It prepares and unprepares the Word.Application object and calls ZtDocument.Prepare or ZtDocument.Unprepare which prepares and unprepares the chosen Word document.
'
' Zotero Tools.
' This software is under Revised ('New') BSD license.
' Copyright © 2019, Olaf Ahrens. All rights reserved.
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Private variables.
Private pvtIsPrepared As Boolean
Private pvtStatusBar As Boolean
Private pvtPagination As Boolean
Private pvtCheckGrammarType As Boolean
Private pvtCheckGrammarSpelling  As Boolean
Private pvtCheckSpellingType As Boolean
Private pvtAutoWordSelection As Boolean
Private pvtPasteWordSpacing As Boolean
Private pvtConfig As ZtConfig
Private pvtDocument As ZtDocument
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Private interface procedures and properties directing to Friend procedures and properties below.
Private Sub ZTIAppPreparable_Prepare()

  Me.Prepare
  
End Sub

Private Sub ZTIAppPreparable_Unprepare()

  Me.Unprepare
  
End Sub
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Friend procedures and properties.
Friend Sub Initialize(ByVal valConfig As ZtConfig, ByVal valDocument As ZtDocument)

  Set pvtConfig = valConfig
  Set pvtDocument = valDocument

End Sub

Friend Sub Prepare()

  If Not pvtIsPrepared And Not pvtConfig.User.Macro.Debugging Then
    pvtDocument.Prepare
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
        pvtPasteWordSpacing = .PasteAdjustWordSpacing
        .PasteAdjustWordSpacing = False
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
        .PasteAdjustWordSpacing = pvtPasteWordSpacing
      End With
      If pvtConfig.User.Macro.WordIsInvisibleWhileOperation Then
        .Visible = True
      End If
    End With
    pvtDocument.Unprepare
    pvtIsPrepared = False
  End If

End Sub
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

