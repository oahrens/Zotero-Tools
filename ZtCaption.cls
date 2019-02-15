VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ZtCaption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Class ZtCaption.
' It preserves informations about a caption, i.e number (e.g. '1.2.2'), text, and position in the document.
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Private variables.
Private pvtConfig As ZtConfig
Private pvtNr As String
Private pvtText As String
Private pvtStart As Long
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Friend procedures and properties.
Friend Sub Initialize(ByVal valConfig As ZtConfig, ByVal valNr As String, ByVal valText As String, ByVal valStart As Long)

  Set pvtConfig = valConfig
  pvtNr = valNr
  pvtText = valText
  Do While Right$(pvtText, 1) = vbCr Or Right$(pvtText, 1) = vbLf
    pvtText = Left$(pvtText, Len(pvtText) - 1)
  Loop
  pvtText = Trim$(pvtText)
  pvtStart = valStart
  
End Sub

Friend Property Get Start() As Long

  Start = pvtStart
  
End Property

Friend Function ScreenTipText(ByVal valPageNr As Integer) As String

  ScreenTipText = Replace(Replace(Replace(pvtConfig.User.BackwardLinkingStyle.ScreenTipPattern, _
                                          pvtConfig.Basic.Macro.BackwardLinkingCaptionNrMarker, _
                                          pvtNr), _
                                  pvtConfig.Basic.Macro.BackwardLinkingCaptionTextMarker, _
                                  pvtText), _
                          pvtConfig.Basic.Macro.BackwardLinkingPageNrMarker, _
                          CStr(valPageNr))
                     
End Function
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *



