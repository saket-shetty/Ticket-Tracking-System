VERSION 5.00
Begin VB.Form frmViewTicket 
   Caption         =   "View Ticket"
   ClientHeight    =   6120
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   12495
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
   ScaleHeight     =   6120
   ScaleWidth      =   12495
   ShowInTaskbar   =   0   'False
   Begin VB.Frame FraTicketDetails 
      Caption         =   "Ticket Details"
      Height          =   4935
      Left            =   5160
      TabIndex        =   2
      Top             =   840
      Width           =   7095
      Begin VB.Label lblTicketId 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ticket Id"
         Height          =   285
         Left            =   360
         TabIndex        =   8
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
         Height          =   285
         Left            =   360
         TabIndex        =   7
         Top             =   4080
         Width           =   780
      End
      Begin VB.Label lblRaisedDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Raised Date:"
         Height          =   285
         Left            =   360
         TabIndex        =   6
         Top             =   3360
         Width           =   1575
      End
      Begin VB.Label lblSeverity 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Severity"
         Height          =   285
         Left            =   360
         TabIndex        =   5
         Top             =   2640
         Width           =   1005
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         Height          =   285
         Left            =   360
         TabIndex        =   4
         Top             =   1920
         Width           =   705
      End
      Begin VB.Label lblDescription 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         Height          =   285
         Left            =   360
         TabIndex        =   3
         Top             =   1200
         Width           =   1380
      End
   End
   Begin VB.ListBox lstTicketList 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4890
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   4455
   End
   Begin VB.Label lblGreetings 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Greetings"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   1380
   End
End
Attribute VB_Name = "frmViewTicket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    FraTicketDetails.Visible = False
    Dim loadTick As New LoadTickets
    lblGreetings = "Hello " & userDetailds.Get_UserName & " here you can see all the open tickets in the portal"
    
    loadTick.Load_All_Tickets
    
    For Each loadTick In ticketCollection
        lstTicketList.AddItem loadTick.ticketId & ": " & loadTick.ticketDesc
    Next

End Sub

Private Sub lstTicketList_Click()
    FraTicketDetails.Visible = True
    Dim loadTick As New LoadTickets
    For Each loadTick In ticketCollection
        If Split(lstTicketList.Text, ":")(0) = loadTick.ticketId Then
            lblDescription = "Description: " & loadTick.ticketDesc
            lblRaisedDate = "Raised Date: " & loadTick.raisedDate
            lblSeverity = "Severity: " & loadTick.severity
            lblStatus = "Status: " & loadTick.status
            lblTicketId = "Ticket Id: " & loadTick.ticketId
            lblName = "Name: " & loadTick.tickCreatorName
        End If
    Next
End Sub
