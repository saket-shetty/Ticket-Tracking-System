VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DBConnectionClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Public Function ConnectDatabase()
    Dim serverName As String
    Dim dbName As String
    Dim configObject As New DBConfiguration
    Dim configString As String
    
    Set DBConnection = New ADODB.Connection
    
    configString = readConfig(App.Path & "\" & "Config.ini")
    
    serverName = Trim$(configObject.getConfigFrom(configString, "Db_Server"))
    dbName = Trim$(configObject.getConfigFrom(configString, "Db_name"))
    
    DBConnection.ConnectionString = "Driver=SQL Server;Server=" & serverName & ";DATABASE=" & dbName & ";"

    Call DBConnection.Open
    
    ConnectDatabase = DBConnection
End Function

Public Function readConfig(ByRef configPath As String) As String
    Dim fsObject As New FileSystemObject
    On Error Resume Next
    
    readConfig = fsObject.OpenTextFile(configPath, ForReading, False).ReadAll
End Function
