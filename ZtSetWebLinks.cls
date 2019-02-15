VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ZtSetWebLinks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements ZtIProcedure

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Class ZtSetWebLinks.
' This is the 'Set web links for references' procedure.
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Private variables.
Private pvtName As String
Private pvtDescription As String
Private pvtFunctions As ZtProcedureFunctions
Private pvtStepsCt As Integer
Private pvtDocument As ZtDocument
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Constructor.
Private Sub Class_Initialize()

  Set pvtFunctions = New ZtProcedureFunctions
  pvtName = "Set web links for references"
  pvtDescription = "This procedure sets hyperlinks in Zotero-generated bibliographies." & vbNewLine & vbNewLine & _
                     "It works for following link formats:" & vbNewLine & _
                     vbTab & "http[s]://en.wikipedia.org" & vbNewLine & _
                     vbTab & "[short[-]]doi:[space|no-break space|tab]10.4315/0362-028X-45.6.533" & vbNewLine & _
                     "This procedure is working for numeric style and author-year style citations. " & _
                       "It only affects the bibliography. Its actions will be overwritten by each Zotero refresh and you have to run it again afterwards."
  
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

Private Property Get ZtIProcedure_ProcedureFunctions() As ZtProcedureFunctions

  Set ZtIProcedure_ProcedureFunctions = Me.ProcedureFunctions
  
End Property

Private Function ZtIProcedure_Start(ByVal valConfig As ZtConfig, ByVal valStartForm As ZtStartForm, ByVal valDocument As ZtDocument, ByRef refLicenceShown As Boolean) As ZtFMessageType

  ZtIProcedure_Start = Me.Start(valConfig, valStartForm, valDocument, refLicenceShown)
  
End Function

Private Sub ZtIProcedure_Prepare()

  Me.Prepare
  
End Sub

Private Sub ZtIProcedure_Unprepare()

  Me.Unprepare
  
End Sub
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *



' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Friend procedures and properties.
Friend Property Get Name() As String

  Name = pvtName
  
End Property

Friend Property Get Description() As String

  Description = pvtDescription
  
End Property

Friend Property Get ProcedureFunctions() As ZtProcedureFunctions

  Set ProcedureFunctions = pvtFunctions
  
End Property

Friend Function Start(ByVal valConfig As ZtConfig, ByVal valStartForm As ZtStartForm, ByVal valDocument As ZtDocument, ByRef refLicenceShown As Boolean) As ZtFMessageType

  Dim locResult As ZtFMessageType

  Set pvtDocument = valDocument
  locResult = pvtFunctions.Initialize(valConfig, valStartForm, pvtDocument, pvtName, refLicenceShown)
  If locResult = MessageOk Then
    valStartForm.Progress.SetStepsCt pvtGetStepsCt
    locResult = pvtDocument.SetWebLinks
    valStartForm.Progress.SetCompleted
  End If
  pvtFunctions.Finish locResult
  
  Start = locResult
  
End Function

Friend Sub Prepare()

  pvtFunctions.Prepare
  
End Sub

Friend Sub Unprepare()

  pvtFunctions.Unprepare
  
End Sub
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Private procedures and properties.
Private Function pvtGetStepsCt() As Integer

  If pvtStepsCt = 0 Then
    pvtStepsCt = pvtDocument.GetSetWebLinksStepsCt
  End If
  
  pvtGetStepsCt = pvtStepsCt
  
End Function
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

