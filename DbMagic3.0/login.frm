VERSION 5.00
Begin VB.Form Form5 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Security Window"
   ClientHeight    =   1590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4440
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1590
   ScaleWidth      =   4440
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Security Window"
      Height          =   1335
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   4215
      Begin VB.CommandButton Command1 
         Caption         =   "Connect"
         Height          =   375
         Left            =   3000
         TabIndex        =   0
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox Text2 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1320
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1320
         TabIndex        =   1
         Text            =   "sa"
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Password :"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   840
         Width           =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "User Id :"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   600
      End
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
uname = Text1.Text
pass = Text2.Text
Me.Hide
Unload Me
db = ""
DoEvents
LogonServer (Provider)
DoEvents
End Sub

Private Sub Form_Load()
If Provider = "SQL Server" Then
    Text1.Text = "sa"
    Text2.Text = ""
End If

If Provider = "Oracle" Then
    Text1.Text = "scott"
    Text2.Text = "tiger"
End If
End Sub
