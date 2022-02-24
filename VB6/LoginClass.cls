VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LoginClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private userIdLocal As String
Private passwordLocal As String
Private departmentLocal As String
Private userNameLocal As String

Public Property Let Set_userId(id As String)
    userIdLocal = id
End Property

Public Property Let Set_password(pass As String)
    passwordLocal = pass
End Property

Public Property Let Set_department(dept As String)
    departmentLocal = dept
End Property

Public Property Let Set_UserName(name As String)
    userNameLocal = name
End Property

Public Property Get Get_UserName() As String
    Get_UserName = userNameLocal
End Property

Public Property Get Get_UserDepartment() As String
    Get_UserDepartment = departmentLocal
End Property

Public Property Get Get_UserId() As String
    Get_UserId = userIdLocal
End Property

Public Function Check_Login() As Boolean
    On Error GoTo Error_Label
    
        Dim sql As String
        Dim recSet As New ADODB.Recordset
        
        sql = "Select EA.User_ID From EMPLOYEE As E " & _
                "inner join EMPLOYEEAUTHENTICATION As EA " & _
                "on E.EID = EA.EID " & _
                "Where E.EID = '" & userIdLocal & "' And EA.Password = '" & passwordLocal & "' And E.Dept = '" & departmentLocal & "'"
        
        Set recSet = DBConnection.Execute(sql)
        
        If Not recSet.EOF Then
            MsgBox "Login successfully"
            Set_UserName = recSet.Fields(0)
            Check_Login = True
        Else
            MsgBox "Login Failed: employee id or password was incorrect"
            Check_Login = False
        End If
    
    Exit Function
Error_Label:
    MsgBox "Error! " & Err.Number & " : " & Err.Description
End Function


Public Function Get_All_Employee()
    On Error GoTo Error_Label
    
        If allEmployeeList.Count = 0 Then
            Dim sql As String
            Dim rc As New ADODB.Recordset
            
            sql = "Select * From EMPLOYEE"
            
            Set rc = DBConnection.Execute(sql)
            
            Dim emp As New loginClass
            
            Do
                Set emp = New loginClass
                
                With emp
                    .Set_userId = rc.Fields(0)
                    .Set_UserName = rc.Fields(1)
                    .Set_department = rc.Fields(3)
                End With
                
                allEmployeeList.Add emp
                
                rc.MoveNext
            Loop Until rc.EOF
        End If
    
    Exit Function
Error_Label:
    MsgBox "Error! " & Err.Number & " : " & Err.Description
End Function


