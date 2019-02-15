Attribute VB_Name = "ZtApiProcedures"
Option Explicit

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Class ZtApiFunctions.
' It contains all calls to the Windows API (application programming interface).
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

#If VBA7 Then
  Public Declare PtrSafe Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszSoundName As String, ByVal hMod As Any, ByVal fdwSound As Long) As Long
  Public Declare PtrSafe Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
  Public Declare PtrSafe Function GetParent Lib "user32.dll" (ByVal hWnd As LongPtr) As LongPtr
  Public Declare PtrSafe Function EnableWindow Lib "user32.dll" (ByVal hWnd As LongPtr, ByVal fEnable As Long) As Long
#Else
  Public Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszSoundName As String, ByVal hMod As Long, ByVal fdwSound As Long) As Long
  Public Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
  Public Declare Function GetParent Lib "user32.dll" (ByVal hWnd As Long) As Long
  Public Declare Function EnableWindow Lib "user32.dll" (ByVal hWnd As Long, ByVal fEnable As Long) As Long
#End If
