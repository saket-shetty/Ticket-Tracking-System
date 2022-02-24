VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DBConfiguration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Function getConfigFrom(ByVal fileContent As String, configName As String) As String
    
    Dim charIndex As Integer
    Dim returnValue As String
    
    On Error Resume Next
    
    charIndex = InStr(1, fileContent, configName, vbTextCompare)
    returnValue = Mid$(fileContent, charIndex)
    
    charIndex = InStr(1, returnValue, vbCrLf, vbTextCompare)
    returnValue = Mid$(returnValue, 1, charIndex - 1)
    
    charIndex = InStr(1, returnValue, "=", vbTextCompare)
    returnValue = Mid$(returnValue, charIndex + 1)
    
    getConfigFrom = returnValue
    
End Function
