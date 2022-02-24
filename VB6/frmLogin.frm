VERSION 5.00
Begin VB.Form frmLogin 
   Caption         =   "Login"
   ClientHeight    =   7170
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8880
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7170
   ScaleWidth      =   8880
   ShowInTaskbar   =   0   'False
   Begin VB.Frame FraEnterDetails 
      Caption         =   "Enter Details to Login"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5655
      Left            =   840
      TabIndex        =   0
      Top             =   960
      Width           =   7095
      Begin VB.ComboBox ComDepartment 
         Height          =   405
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   2880
         Width           =   3735
      End
      Begin VB.TextBox txtPassword 
         Height          =   495
         IMEMode         =   3  'DISABLE
         Left            =   2280
         PasswordChar    =   "*"
         TabIndex        =   7
         Top             =   1800
         Width           =   3735
      End
      Begin VB.TextBox txtEmployeeId 
         Height          =   495
         Left            =   2280
         TabIndex        =   6
         Top             =   840
         Width           =   3735
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   480
         Left            =   2280
         TabIndex        =   5
         Top             =   4680
         Width           =   1710
      End
      Begin VB.CommandButton cmdSubmit 
         Caption         =   "Submit"
         Height          =   480
         Left            =   2280
         TabIndex        =   4
         Top             =   3840
         Width           =   1710
      End
      Begin VB.Label lblDepartment 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Department: "
         Height          =   285
         Left            =   240
         TabIndex        =   3
         Top             =   2880
         Width           =   1635
      End
      Begin VB.Label lblPassword 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password: "
         Height          =   285
         Left            =   240
         TabIndex        =   2
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label lblEmployeeID 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Employee Id: "
         Height          =   285
         Left            =   240
         TabIndex        =   1
         Top             =   960
         Width           =   1680
      End
   End
   Begin VB.Label lblHelloWelcome 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hello Welcome to ticket tracking portal, login to check you tickets"
      Height          =   285
      Left            =   360
      TabIndex        =   9
      Top             =   360
      Width           =   8070
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private loginClass As New loginClass

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSubmit_Click()
    Dim success As Boolean
    loginClass.Set_userId = txtEmployeeId
    loginClass.Set_password = txtPassword
    loginClass.Set_department = ComDepartment.Text
    
    Dim x As New PasswordAuth
    
    Dim validFields As Boolean
    
    validFields = True
    
    If txtEmployeeId.Text = Empty Or txtEmployeeId.Text = "" Then
        Call MsgBox("Employee Id cannot be empty.", vbOKOnly, "Employee Id")
        validFields = False
    End If
    
    If ComDepartment.Text = "" Then
        Call MsgBox("Please select a department from dropdown.", vbOKOnly, "Department")
        validFields = False
    End If
    
    
    If validFields And x.Check_Password(txtPassword) Then
        success = loginClass.Check_Login
    
        If success Then
            Set userDetailds = loginClass
            isLogin (True)
            Unload Me
            
            If UCase(userDetailds.Get_UserDepartment) = "DEVOPS" Then
                MDIHomepage.mnuCreateTicket.Visible = False
            Else
                MDIHomepage.mnuViewTicket.Visible = False
                MDIHomepage.mnuCloseTicket.Visible = False
                MDIHomepage.mnuReport.Visible = False
            End If
        End If
    End If
    
End Sub

Private Sub Form_Load()
    
    Dim dbCon As New DBConnectionClass
    Dim loadCol As New LoadDepartmentClass
    Dim loadTick As New LoadTickets
    
    Call dbCon.ConnectDatabase
    Call loadCol.Load_Department_Combo_Box
    
    Call loginClass.Get_All_Employee
    
    Call loadTick.Load_All_Tickets
    
    For Each x In departmentCollection
        ComDepartment.AddItem x
    Next
End Sub
