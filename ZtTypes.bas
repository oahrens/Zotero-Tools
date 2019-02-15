Attribute VB_Name = "ZTTypes"
Option Explicit

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Enum definitions.
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

Public Enum ZtEGivenNameStyle
  GivenNameNone = 0
  GivenNameAbbreviated = 1
  GivenNameFull = 2
End Enum

Public Enum ZtFSound
  SndSync = 0
  SndAsync = 1
  SndNoDefault = 2
  SndMemory = 4
  SndLoop = 8
  SndNoStop = 16
  SndPurge = 64
  SndApplication = 128
  SndNoWait = 8192
  SndAlias = 65536
  SndFileName = 131072
  SndResource = 262148
  SndSentry = 524288
  SndAliasId = 1114112
  SndSystem = 2097152
End Enum

Public Enum ZtFMessageType
  MessageNone = 0
  MessageOk = 1
  MessageCancel = 2
  MessageOkCancel = 3
  MessageDisable = 4
  MessageOkDisableCancel = 7
  Messageno = 8
  MessageInformation = 16
  MessageQuestion = 32
  MessageExclamation = 64
  MessageCritical = 128
End Enum

Public Enum ZtFSurroundingSigns
  NoSigns = 0
  SignBefore = 1
  SignAfter = 2
  SignsBeforeAndAfter = 3
End Enum

Public Enum ZtCBoolean
  CFalse = 0
  CTrue = 1
End Enum
