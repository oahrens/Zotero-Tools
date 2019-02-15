VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ZtRegDebuggingWrapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Class ZtRegDebuggingWrapper.
' It capsules ZtRegExp to distribute its debugging informations to ZtMessageDisplay.
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Private variables.
Private pvtRegExp As ZtRegExp
Private pvtConfig As ZtConfig
Private pvtStartForm As ZtStartForm
Private pvtDebugging As Boolean
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Constructor.
Private Sub Class_Initialize()

  pvtDebugging = True
  Set pvtRegExp = New ZtRegExp
  
End Sub
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Friend procedures and properties.
Friend Sub Initialize(ByVal valConfg As ZtConfig, ByVal valStartForm As ZtStartForm, ByVal valName As String, ByVal valPattern As String, _
  Optional ByVal valGlobal As Boolean = False, Optional ByVal valIgnoreCase As Boolean = False, Optional ByVal valMultiLine As Boolean = False)

  Set pvtConfig = valConfg
  Set pvtStartForm = valStartForm
  pvtRegExp.Initialize valName, valPattern, valGlobal, valIgnoreCase, valMultiLine
  
End Sub

Friend Function Debugging(ByVal valMatch As ZtRegMatch, ByVal valRange As Word.Range) As ZtFMessageType

  Dim locResult As ZtFMessageType
  Dim locText As String

  If pvtConfig.User.Macro.Debugging And pvtDebugging Then
    locText = pvtRegExp.DebugText(valMatch) & _
              vbNewLine & vbNewLine & _
              "With 'Disable' this debug information won't be shown anymore, otherwise procede or cancel."
    locResult = pvtStartForm.MessageDisplay.Show(locText, MessageOkDisableCancel + MessageInformation, True, valRange)
    If locResult = MessageDisable Then
      pvtDebugging = False
      locResult = MessageOk
    End If
  Else
    locResult = MessageOk
  End If
  
  If locResult = vbOK Then
    Debugging = MessageOk
  ElseIf locResult = vbCancel Then
    Debugging = MessageCancel
  End If

End Function

Friend Function Test(ByVal valSourceString As String) As Boolean
  
  Test = pvtRegExp.Test(valSourceString)
  
End Function

Friend Function Execute(ByVal valSourceString As String) As Collection

  Set Execute = pvtRegExp.Execute(valSourceString)

End Function

Friend Function FirstMatch(ByVal valSourceString As String) As ZtRegMatch

  Set FirstMatch = pvtRegExp.FirstMatch(valSourceString)
  
End Function

Friend Function Replace(ByVal valSourceString As String, ByVal valReplacement As String) As String

  Replace = pvtRegExp.Replace(valSourceString, valReplacement)
  
End Function
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

