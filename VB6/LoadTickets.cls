VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LoadTickets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public ticketId As Integer
Public loggedId As String
Public raisedDate As String
Public severity As String
Public ticketDesc As String
Public resolvedBy As String
Public resolution As String
Public resolvedDate As String
Public status As String
Public tickCreatorName As String


Public Function Load_All_Tickets()
    
    If ticketCollection.Count = 0 Then
        
        Dim lticket As New LoadTickets
    
        Dim sql As String
        Dim recSet As New ADODB.Recordset
        
        sql = "Select * From TICKET As T Inner join EMPLOYEE As E on E.EID = T.LOGGED_BY"
        
        Set recSet = DBConnection.Execute(sql)
        
        Do
            Set lticket = New LoadTickets
            
            With lticket
                .ticketId = recSet.Fields(0)
                .loggedId = recSet.Fields(1)
                .raisedDate = recSet.Fields(2)
                .severity = recSet.Fields(3)
                .ticketDesc = recSet.Fields(4)
                .resolvedBy = IIf(IsNull(recSet.Fields(5)), "NULL", recSet.Fields(5))
                .resolution = IIf(IsNull(recSet.Fields(6)), "NULL", recSet.Fields(6))
                .resolvedDate = IIf(IsNull(recSet.Fields(7)), "NULL", recSet.Fields(7))
                .status = recSet.Fields(8)
                .tickCreatorName = recSet.Fields(10)
            End With
            
            ticketCollection.Add lticket
            
            recSet.MoveNext
        Loop Until recSet.EOF
    
    End If
    
End Function

Public Function Create_New_Ticket(loggedId As String, tdate As Date, severity As String, desc As String, status As String, ticketCreateName As String) As Boolean
    On Error GoTo Error_Label
    
        Dim sql As String
        Dim rc As New ADODB.Recordset
        
        sql = "INSERT INTO TICKET (LOGGED_BY, RAISED_DATE, SEVERITY, TICKET_DESC, STATUS) VALUES ("
        sql = sql & "'" & loggedId & "','" & tdate & "','" & severity & "','" & desc & "','" & status & "')"
    
        DBConnection.Execute (sql)
        
        sql = "Select TOP(1) TICKET_ID From Ticket Order By TICKET_ID DESC"
        
        Set rc = DBConnection.Execute(sql)
        
        MsgBox "Ticket " & rc.Fields(0) & " is submitted Successfully"
        
        Dim newTicket As New LoadTickets
        
        newTicket.ticketId = rc.Fields(0)
        newTicket.loggedId = loggedId
        newTicket.raisedDate = tdate
        newTicket.severity = severity
        newTicket.status = "OPEN"
        newTicket.ticketDesc = desc
        newTicket.tickCreatorName = ticketCreateName
        
        ticketCollection.Add newTicket
        
        Create_New_Ticket = True
    Exit Function
Error_Label:
    MsgBox "Error! " & Err.Number & " : " & Err.Description
    Create_New_Ticket = False
End Function


Public Function Close_Ticket(ticketId As String, resolverId As String, res As String) As Boolean
    On Error GoTo Error_Label
    
        Dim sql As String
        
        sql = "UPDATE TICKET SET RESOLVED_BY  = '" & resolverId & "', RESOLUTION = '" & res & "', RESOLVED_DATE = '" & DateTime.Date & " " & DateTime.Time & "', STATUS = 'CLOSED' Where TICKET_ID = " & ticketId
        
        DBConnection.Execute (sql)
        
        Close_Ticket = True
    
    Exit Function
Error_Label:
    MsgBox "Error! " & Err.Number & " : " & Err.Description
    Close_Ticket = False
End Function
