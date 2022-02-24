VERSION 5.00
Begin VB.Form frmCreateTicket 
   Caption         =   "Create / Log a ticket"
   ClientHeight    =   6825
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   10710
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6825
   ScaleWidth      =   10710
   ShowInTaskbar   =   0   'False
   Begin VB.Frame FraCreateTicket 
      Caption         =   "Create Ticket"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6015
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   10335
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   525
         Left            =   4080
         TabIndex        =   11
         Top             =   5160
         Width           =   1590
      End
      Begin VB.CommandButton cmdSubmit 
         Caption         =   "Submit"
         Height          =   525
         Left            =   1920
         TabIndex        =   10
         Top             =   5160
         Width           =   1830
      End
      Begin VB.TextBox txtDescription 
         Height          =   1815
         Left            =   1920
         MultiLine       =   -1  'True
         TabIndex        =   9
         Top             =   3000
         Width           =   6375
      End
      Begin VB.ComboBox ComSeverity 
         Height          =   405
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   2160
         Width           =   3975
      End
      Begin VB.TextBox txtDate 
         Height          =   405
         Left            =   1920
         TabIndex        =   5
         Top             =   1440
         Width           =   3975
      End
      Begin VB.ComboBox ComEmployeeDetails 
         Height          =   405
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   720
         Width           =   3975
      End
      Begin VB.Label lblDesccription 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Description: "
         Height          =   285
         Left            =   360
         TabIndex        =   8
         Top             =   3000
         Width           =   1350
      End
      Begin VB.Label lblSeverity 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Severity: "
         Height          =   285
         Left            =   360
         TabIndex        =   6
         Top             =   2160
         Width           =   1005
      End
      Begin VB.Label lblTicketDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ticket Date: "
         Height          =   285
         Left            =   360
         TabIndex        =   4
         Top             =   1440
         Width           =   1350
      End
      Begin VB.Label lblEmployee 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Employee: "
         Height          =   285
         Left            =   360
         TabIndex        =   2
         Top             =   720
         Width           =   1200
      End
   End
   Begin VB.Label lblGreetings 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Greetings"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   1200
   End
End
Attribute VB_Name = "frmCreateTicket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSubmit_Click()

    If Form_Validate Then
    
        Dim ticket As New LoadTickets
        
        Dim success As Boolean
        
        Dim id As String
        Dim creatorName As String
        
        id = Split(ComEmployeeDetails.Text, ": ")(0)
        creatorName = Split(ComEmployeeDetails.Text, ": ")(1)
        
        success = ticket.Create_New_Ticket(id, CDate(txtDate), ComSeverity.Text, txtDescription, "OPEN", creatorName)
        
        If success Then
            Unload Me
        End If
            
    End If
End Sub

Private Sub Form_Load()
    txtDate.Enabled = False
    lblGreetings = "Hello " & userDetailds.Get_UserName & " here you can create a ticket about you issue."
    
    txtDate = DateTime.Date & " " & DateTime.Time
    
    Dim x As New loginClass
    
    For Each x In allEmployeeList
        ComEmployeeDetails.AddItem x.Get_UserId & ": " & x.Get_UserName
    Next
    
    ComSeverity.AddItem "Major"
    ComSeverity.AddItem "Critical"

End Sub

Private Function Form_Validate() As Boolean
    Dim valid As Boolean
    valid = True
    
    If ComEmployeeDetails.Text = "" Then
        MsgBox "Please select one employee from dropdown."
        valid = False
    End If
    
    If txtDate.Text <> "" Or txtDate.Text <> Empty Then
        If CDate(txtDate) > CDate(DateTime.Date & " " & DateTime.Time) Then
            MsgBox "Date cannot be greater than current date."
            valid = False
        End If
    Else
        MsgBox "Date cannot be empty."
    End If

    
    If ComSeverity.Text = "" Then
        MsgBox "Please select severity of the ticket from the dropdown."
        valid = False
    End If
    
    If txtDescription = Empty Then
        MsgBox "Description cannot be empty."
        valid = False
    End If
    
    Form_Validate = valid
End Function

