VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ZtIProcedure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Interface ZtIProcedure.
' Interface for all main procedures of the macro, i.e. the procedures started by ZtStartForm.
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Interface procedures and properties.
' All interface definitions must be Public.
Public Property Get Name() As String
End Property

Public Property Get Description() As String
End Property

Public Property Get ProcedureFunctions() As ZtProcedureFunctions
End Property

Public Function Start(ByVal valConfig As ZtConfig, ByVal valStartForm As ZtStartForm, ByVal valDocument As ZtDocument, ByRef refLicenceShown As Boolean) As ZtFMessageType
End Function

Public Sub Prepare()
End Sub

Public Sub Unprepare()
End Sub
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

