VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ZtIProcedureInitializable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Interface ZtIProcedureInitializable.
' Interface for ZtProceduteInitializer to separate specification from implementation.
'
' Zotero Tools.
' This software is under Revised ('New') BSD license.
' Copyright © 2019, Olaf Ahrens. All rights reserved.


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Public procedures and properties.
' All interface definitions must be Public.
Public Function Start(ByVal valConfig As ZtConfig, ByVal valMessageDisplay As ZtIMessageDisplayable, ByVal valAppPrepare As ZtIAppPreparable, _
  ByVal valDocument As ZtDocument, ByVal valProcedureName As String, ByRef refLicenseShown As Boolean, Optional ByVal valTakesTime As Boolean = False, _
  Optional ByVal valResumed As Boolean = False) As ZtFMessageType
End Function

Public Sub Finish(ByVal valResult As ZtFMessageType, Optional ByVal valExtraClosingText As String = vbNullString, Optional ByVal valResume As Boolean = False)
End Sub
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


