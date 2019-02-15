Attribute VB_Name = "ZtStart"
Option Explicit

Public Sub Start()

  Dim locStartForm As ZtStartForm
  
  Set locStartForm = New ZtStartForm
  locStartForm.Show vbModal
  
End Sub
