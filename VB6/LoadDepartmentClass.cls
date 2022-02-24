VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LoadDepartmentClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private departmentName As String

Public Property Let Set_DeptName(dept As String)
    departmentName = dept
End Property

Public Property Get Get_DeptName() As String
    Get_DeptName = departmentName
End Property

Public Function Load_Department_Combo_Box()
    On Error GoTo Error_Label
        If departmentCollection.Count = 0 Then
            Dim sql As String
            Dim recSet As New ADODB.Recordset
            
            sql = "Select Dept from Dept_Value"
            
            Set recSet = DBConnection.Execute(sql)
            
            Do
                Set_DeptName = recSet.Fields(0)
                departmentCollection.Add Get_DeptName
                recSet.MoveNext
            Loop Until recSet.EOF
        End If
    Exit Function
Error_Label:
    MsgBox "Error! " & Err.Number & " : " & Err.Description
End Function
