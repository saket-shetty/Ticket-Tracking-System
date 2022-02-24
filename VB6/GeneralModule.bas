Attribute VB_Name = "GeneralModule"
Public DBConnection As New ADODB.Connection

Public departmentCollection As New Collection

Public userDetailds As New loginClass

Public ticketCollection As New Collection

Public allEmployeeList As New Collection


Public Sub isLogin(mnuVisible As Boolean)
    With MDIHomepage
        .mnuLogin.Visible = Not mnuVisible
        .mnuCloseTicket.Visible = mnuVisible
        .mnuCreateTicket.Visible = mnuVisible
        .mnuLogout.Visible = mnuVisible
        .mnuReport.Visible = mnuVisible
        .mnuViewTicket.Visible = mnuVisible
    End With
End Sub
