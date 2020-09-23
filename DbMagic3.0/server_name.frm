VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Server Name"
   ClientHeight    =   1425
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3960
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1425
   ScaleWidth      =   3960
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cmd 
      Left            =   3360
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "Server Name"
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   3855
      Begin VB.ComboBox Text1 
         Height          =   315
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   2535
      End
      Begin VB.CommandButton Command1 
         Caption         =   "OK"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2760
         TabIndex        =   1
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Database Name"
      Height          =   1215
      Left            =   0
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   3855
      Begin VB.CommandButton Command3 
         Caption         =   "OK"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2520
         TabIndex        =   5
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   240
         Width           =   2295
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Browse"
         Height          =   375
         Left            =   2520
         TabIndex        =   3
         Top             =   240
         Width           =   1215
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
ServerName = Text1.Text
If cn.State = 1 Then
    cn.Close
End If
Unload Me
Form5.Show 1
End Sub

Private Sub Command2_Click()
If Provider = "Ms Access 2000" Or Provider = "Ms Access 97" Then
    cmd.Filter = "*.mdb"
    cmd.FileName = "*.mdb"
End If
If Provider = "Foxpro" Then
    cmd.Filter = "*.dbf"
    cmd.FileName = "*.dbf"
End If

cmd.DialogTitle = "Select Database"
cmd.ShowOpen

If Provider = "Foxpro" Then
a = cmd.FileTitle
b = cmd.FileName
Text2.Text = Mid(b, 1, Len(b) - Len(a) - 1)
Else
Text2.Text = cmd.FileName
End If

Command3.SetFocus
End Sub

Private Sub Command3_Click()
db = Text2.Text
Unload Me
If cn.State = 1 Then
    cn.Close
End If
LogonServer (Provider)
End Sub

Private Sub Form_Load()
If Provider = "SQL Server" Or Provider = "Oracle" Then
    Text1.Text = Text1.List(0)
    Frame2.Visible = False
    Frame1.Visible = True
End If
If Provider = "Ms Access 2000" Or Provider = "Ms Access 97" Or Provider = "Foxpro" Then
    Frame1.Visible = False
    Frame2.Visible = True
End If
End Sub

Private Sub Text1_Change()
If Len(Text1.Text) = 0 Then
    Command1.Enabled = False
Else
    Command1.Enabled = True
End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    If Text1.Text = "" Then
        MsgBox "Please Enter Server Name", vbCritical, "Help!"
        Exit Sub
    End If
    Call Text1_Change
    
    ServerName = Text1.Text
    If cn.State = 1 Then
        cn.Close
    End If
    Unload Me
    Form5.Show 1
End If

End Sub

Private Sub Text2_Change()
If Len(Text2.Text) = 0 Then
    Command3.Enabled = False
Else
    Command3.Enabled = True
End If
End Sub
