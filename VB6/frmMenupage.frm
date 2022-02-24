VERSION 5.00
Begin VB.Form frmMenupage 
   Caption         =   "Home Page"
   ClientHeight    =   6030
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   12165
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   14.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   6030
   ScaleWidth      =   12165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Label lblGreetings 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Greetings"
      Height          =   345
      Left            =   720
      TabIndex        =   0
      Top             =   240
      Width           =   1380
   End
   Begin VB.Menu mnuViewTicket 
      Caption         =   "View Tickets"
   End
   Begin VB.Menu mnuCreateTicket 
      Caption         =   "Create Ticket"
   End
   Begin VB.Menu mnuCloseTicket 
      Caption         =   "Close a Ticket"
   End
   Begin VB.Menu mnuReport 
      Caption         =   "Report"
   End
   Begin VB.Menu mnuLogout 
      Caption         =   "Logout"
   End
End
Attribute VB_Name = "frmMenupage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    lblGreetings = "Hello " & userDetailds.Get_UserName & " Welcome to Ticket Tracking Portal"
    
    If UCase(userDetailds.Get_UserDepartment) = "DEVOPS" Then
        mnuCreateTicket.Visible = False
    Else
        mnuViewTicket.Visible = False
    End If
End Sub

Private Sub mnuCloseTicket_Click()
    frmCloseTicket.Show
End Sub

Private Sub mnuCreateTicket_Click()
    frmCreateTicket.Show
End Sub

Private Sub mnuLogout_Click()
    Unload Me
    frmLogin.Show
End Sub

Private Sub mnuReport_Click()
    Dim crApp As New CRAXDRT.Application
    Dim crRpt As New CRAXDRT.Report
    
    Dim filePath As String
    
    filePath = App.Path & "\TICKET_TRACKING_REPORT.rpt"
    
    Set crRpt = crApp.OpenReport(filePath)
    
    frmReport.crTicketReport.ReportSource = crRpt
    frmReport.crTicketReport.ViewReport
    frmReport.Show
End Sub

Private Sub mnuViewTicket_Click()
    frmViewTicket.Show
End Sub
