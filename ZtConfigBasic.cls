VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ZtConfigBasic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Class ZtConfigBasic.
' It preserves basic configuration informations and methods to read them from ZtConfig.xml.
' It distributes these informations and work to the other ZtConfigBasic... classes.
'
' Zotero Tools.
' This software is under Revised ('New') BSD license.
' Copyright © 2019, Olaf Ahrens. All rights reserved.
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Private variables.
Private pvtCharacters As ZtConfigBasicCharacters
Private pvtMacro As ZtConfigBasicMacro
Private pvtPatterns As ZtConfigBasicPatterns
Private pvtRegPieces As ZtConfigBasicRegPieces
Private pvtWord As ZtConfigBasicWord
Private pvtZotero As ZtConfigBasicZotero
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Constructor.
Private Sub Class_Initialize()

  Set pvtMacro = New ZtConfigBasicMacro
  Set pvtWord = New ZtConfigBasicWord
  Set pvtZotero = New ZtConfigBasicZotero
  Set pvtCharacters = New ZtConfigBasicCharacters
  Set pvtPatterns = New ZtConfigBasicPatterns
  Set pvtRegPieces = New ZtConfigBasicRegPieces

End Sub
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Friend/Public procedures and properties.
' All members that should be callable by CallByName procedure must be public.
Friend Sub Initialize(ByVal valConfig As ZtConfig, ByVal valXml As MSXML2.DOMDocument60)

  pvtMacro.Initialize valConfig, valXml
  pvtWord.Initialize valConfig, valXml
  pvtZotero.Initialize valConfig, valXml
  pvtCharacters.Initialize valConfig, valXml
  pvtPatterns.Initialize valConfig, valXml
  pvtRegPieces.Initialize valConfig, valXml
  
End Sub

Public Property Get Characters() As ZtConfigBasicCharacters

  Set Characters = pvtCharacters
  
End Property

Public Property Get Macro() As ZtConfigBasicMacro

  Set Macro = pvtMacro
  
End Property

Public Property Get Patterns() As ZtConfigBasicPatterns

  Set Patterns = pvtPatterns
  
End Property

Public Property Get RegPieces() As ZtConfigBasicRegPieces

  Set RegPieces = pvtRegPieces
  
End Property

Public Property Get Word() As ZtConfigBasicWord

  Set Word = pvtWord
  
End Property

Public Property Get Zotero() As ZtConfigBasicZotero

  Set Zotero = pvtZotero
  
End Property
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
