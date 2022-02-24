VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PasswordAuth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Function Check_Password(pass As String) As Boolean
    
    Dim upper As Boolean, special As Boolean, length As Boolean
    
    If Len(pass) >= 8 And Len(pass) < 13 Then
        length = True
    End If
    
    For x = 1 To Len(pass)
        Dim KeyAscii As Integer
        KeyAscii = Asc(Mid(pass, x, 1))
        If KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Then
            upper = True
        End If
        If (KeyAscii >= 33 And KeyAscii <= 47) Or (KeyAscii >= 58 And KeyAscii <= 64) Or (KeyAscii >= 123 And KeyAscii <= 126) Then
            special = True
        End If
    Next
    
    If upper And special And length Then
        Check_Password = True
    Else
        Call MsgBox("Invalid password. Password should be combination of " & vbCrLf & _
                "Upper character, Special character and length should be at least 8 " & vbCrLf & _
                "Character", vbOKOnly, "Password Validation Error")
        Check_Password = False
    End If
End Function
