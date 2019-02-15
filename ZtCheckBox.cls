VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ZtCheckBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Class ZtCheckBox.
' It capsules MSForms.CheckBox.
' It enables the check box to cache its enabled state before disabled and its value before disabled and unchecked simultaneously.
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Private variables.
Private WithEvents pvtCheckBox As MSForms.CheckBox
Attribute pvtCheckBox.VB_VarHelpID = -1
Private pvtControl As MSForms.Control
Private pvtCallbackProperty As String
Private pvtDependentCheckBoxes() As ZtCheckBox
Private pvtHasDependentCheckBoxes As Boolean
Private pvtIsProtected As Boolean
Private pvtEnabledCache As Boolean
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Friend procedures and properties.
Friend Sub Initialize(ByVal valCheckBox As MSForms.CheckBox, ByVal valCallbackProperty As String, ParamArray refDependentCheckBoxes() As Variant)
  
  Dim locCtr As Integer
  
  Set pvtCheckBox = valCheckBox
  Set pvtControl = valCheckBox
  pvtCallbackProperty = valCallbackProperty
  Debug.Print
  If UBound(refDependentCheckBoxes) >= LBound(refDependentCheckBoxes) Then ' Not IsMissing(valDependentCheckBoxes) works too, unrespected what MS says.
    ReDim pvtDependentCheckBoxes(UBound(refDependentCheckBoxes))
    For locCtr = 0 To UBound(refDependentCheckBoxes)
      Set pvtDependentCheckBoxes(locCtr) = refDependentCheckBoxes(locCtr)
    Next
    pvtHasDependentCheckBoxes = True
  End If
  
End Sub

Friend Property Get Enabled() As Boolean

  Enabled = pvtCheckBox.Enabled
  
End Property

Friend Property Let Enabled(ByVal valEnable As Boolean)
  
  If valEnable <> pvtCheckBox.Enabled Then
    pvtEnabledCache = valEnable
    pvtCheckBox.Enabled = valEnable
  End If
  
End Property

Friend Property Let EnabledCacheValue(ByVal valEnable As Boolean)

  If valEnable <> pvtCheckBox.Enabled Then
    pvtCheckBox.Enabled = valEnable
    pvtEnabledCache = valEnable
    pvtIsProtected = True
    If valEnable Then
      pvtCheckBox.Value = CallByName(pvtControl.Parent, pvtCallbackProperty, VbGet)
    Else
      CallByName pvtControl.Parent, pvtCallbackProperty, VbLet, pvtCheckBox.Value
      pvtCheckBox.Value = False
    End If
    pvtIsProtected = False
  End If
  
End Property

Friend Property Let EnabledCacheEnabled(ByVal valEnable As Boolean)

  If valEnable <> pvtCheckBox.Enabled Then
    If valEnable Then
      pvtCheckBox.Enabled = pvtEnabledCache
    Else
      pvtEnabledCache = pvtCheckBox.Enabled
      pvtCheckBox.Enabled = valEnable
    End If
  End If
  
End Property

Friend Property Get Value() As Boolean
  
  Value = pvtCheckBox.Value
  
End Property

Friend Property Let Value(ByVal valValue As Boolean)

  If Not pvtIsProtected Then
    pvtCheckBox.Value = valValue
  End If
  
End Property
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Event procedures.
Private Sub pvtCheckBox_Change()

  Dim locCtr As Integer
  
  If Not pvtIsProtected Then
    If pvtHasDependentCheckBoxes Then
      pvtIsProtected = True
      For locCtr = 0 To UBound(pvtDependentCheckBoxes)
        pvtDependentCheckBoxes(locCtr).Value = False
        pvtDependentCheckBoxes(locCtr).Enabled = Not pvtCheckBox.Value
      Next
      pvtIsProtected = False
    End If
    CallByName pvtControl.Parent, pvtCallbackProperty, VbLet, pvtCheckBox.Value
  End If

End Sub
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

