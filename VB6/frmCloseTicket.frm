VERSION 5.00
Begin VB.Form frmCloseTicket 
   Caption         =   "Close Ticket"
   ClientHeight    =   6795
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8925
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
   ScaleHeight     =   6795
   ScaleWidth      =   8925
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Detail for closing ticket"
      Height          =   5295
      Left            =   360
      TabIndex        =   0
      Top             =   1080
      Width           =   7695
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   405
         Left            =   4440
         TabIndex        =   8
         Top             =   4440
         Width           =   1590
      End
      Begin VB.CommandButton cmdSubmit 
         Caption         =   "Submit"
         Height          =   405
         Left            =   2280
         TabIndex        =   7
         Top             =   4440
         Width           =   1590
      End
      Begin VB.TextBox txtResolution 
         Height          =   2055
         Left            =   2280
         MultiLine       =   -1  'True
         TabIndex        =   6
         Top             =   1920
         Width           =   4695
      End
      Begin VB.ComboBox ComDevopsEmployee 
         Height          =   405
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1200
         Width           =   4695
      End
      Begin VB.ComboBox comOpenTicket 
         Height          =   405
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   480
         Width           =   2415
      End
      Begin VB.Label lblResolution 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Resolution: "
         Height          =   285
         Left            =   240
         TabIndex        =   5
         Top             =   2040
         Width           =   1470
      End
      Begin VB.Label lblResolvedBy 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Resolved By:"
         Height          =   285
         Left            =   240
         TabIndex        =   3
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label lblTicketId 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ticket Id: "
         Height          =   285
         Left            =   240
         TabIndex        =   1
         Top             =   600
         Width           =   1260
      End
   End
   Begin VB.Label lblGreetings 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Greetings"
      Height          =   285
      Left            =   480
      TabIndex        =   9
      Top             =   360
      Width           =   1185
   End
End
Attribute VB_Name = "frmCloseTicket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim loadTicket As New LoadTickets

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSubmit_Click()

    If Validate_Fields Then
        Dim resId As String
        Dim validResult As Boolean
        
        resId = Split(ComDevopsEmployee.Text, ":")(0)
    
        validResult = loadTicket.Close_Ticket(comOpenTicket.Text, resId, txtResolution)
        
        If validResult Then
            MsgBox "Ticket " & comOpenTicket.Text & " is closed."
            
            For Each loadTicket In ticketCollection
                If loadTicket.ticketId = comOpenTicket.Text Then
                    loadTicket.status = "CLOSED"
                    Exit For
                End If
            Next
                    
            Unload Me
        Else
            MsgBox "Something went wrong."
        End If
        
    End If

End Sub

Private Sub Form_Load()
    
    lblGreetings = "Hello " & userDetailds.Get_UserName & " here you can close any open ticket."
    
    Call loadTicket.Load_All_Tickets
    
    For Each loadTicket In ticketCollection
        If loadTicket.status = "OPEN" Then
            comOpenTicket.AddItem loadTicket.ticketId
        End If
    Next
    
    Dim emp As New loginClass
    
    Call emp.Get_All_Employee
    
    For Each emp In allEmployeeList
        If emp.Get_UserDepartment = "DEVOPS" Then
            ComDevopsEmployee.AddItem emp.Get_UserId & ": " & emp.Get_UserName
        End If
    Next
    
End Sub


Private Function Validate_Fields() As Boolean
    Dim valid As Boolean
    valid = True
    
    If comOpenTicket.Text = "" Then
        MsgBox "Please select a ticket from dropdown."
        valid = False
    End If
    
    If ComDevopsEmployee.Text = "" Then
        MsgBox "Please select one employee from dropdown."
        valid = False
    End If
    
    If txtResolution = Empty Or txtResolution = "" Then
        MsgBox "Resolution of the ticket cannot be empty."
        valid = False
    End If
    
    Validate_Fields = valid
End Function
