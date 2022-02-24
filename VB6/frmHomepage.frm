VERSION 5.00
Begin VB.Form frmHomepage 
   Caption         =   "Landing Page"
   ClientHeight    =   6405
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7965
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
   MDIChild        =   -1  'True
   ScaleHeight     =   6405
   ScaleWidth      =   7965
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCreateNew 
      Caption         =   "Create New User"
      Height          =   480
      Left            =   2160
      TabIndex        =   2
      Top             =   4200
      Width           =   3390
   End
   Begin VB.CommandButton cmdLogin 
      Caption         =   "Login"
      Height          =   480
      Left            =   2160
      TabIndex        =   1
      Top             =   3240
      Width           =   3390
   End
   Begin VB.Label lblWelcomeTo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome to Ticket Tracking System"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   0
      Top             =   840
      Width           =   5610
   End
End
Attribute VB_Name = "frmHomepage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdLogin_Click()
    frmLogin.Show
End Sub

Private Sub Form_Load()

    Dim dbCon As New DBConnectionClass
    Dim loadCol As New LoadDepartmentClass
    
    Call dbCon.ConnectDatabase
    Call loadCol.Load_Department_Combo_Box
End Sub
