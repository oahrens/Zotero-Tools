VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ZtStringBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Class ZtStringBuilder.
' It overcomes bad performance when concatenating strings with '&' repeatedly .
' This is a slightly revised and extended version of Philip Swannell's code (https://github.com/VBA-tools/VBA-JSON/pull/82).
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Private variables.
Private pvtString As String
Private pvtLength As Long
Private pvtCapacity As Long
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Friend procedures and properties.
Friend Sub Append(ByVal valAppend As String)

  Dim locAppendLength As Long
  Dim locAddLength As Long

  locAppendLength = Len(valAppend)

  ' The Mid$ function doen't work with length = 0.
  If locAppendLength > 0 Then
    If pvtLength + locAppendLength > pvtCapacity Then
      locAddLength = ZtSubprocedures.Max(locAppendLength, pvtCapacity)
      pvtString = pvtString & Space$(locAddLength)
      pvtCapacity = pvtCapacity + locAddLength
    End If
    Mid$(pvtString, pvtLength + 1, locAppendLength) = valAppend
    pvtLength = pvtLength + locAppendLength
  End If
  
End Sub

Friend Property Get Length() As Long

  Length = pvtLength
  
End Property

Friend Sub RemoveFromEnd(ByVal valLength As Long)

  If valLength >= pvtLength Then
    pvtLength = 0
  Else
    pvtLength = pvtLength - valLength
  End If
  
End Sub

Friend Sub Clear()

  pvtLength = 0
  
End Sub

Friend Function ToString() As String

  If pvtLength > 0 Then
    ToString = Left$(pvtString, pvtLength)
  End If
  
End Function
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
