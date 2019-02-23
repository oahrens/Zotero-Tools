VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ZtIMessageDisplayable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Interface ZtIMessageDisplayable.
' Interface for ZtMessageDisplay to separate specification from implementation.
'
' Zotero Tools.
' This software is under Revised ('New') BSD license.
' Copyright © 2019, Olaf Ahrens. All rights reserved.


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Public procedures and properties.
' All interface definitions must be Public.
Public Function Show(ByVal valMessage As String, ByVal valType As ZtFMessageType, Optional ByVal valRePrepare As Boolean = True, _
  Optional ByVal valSelectionRange As Word.Range = Nothing) As ZtFMessageType
End Function
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


