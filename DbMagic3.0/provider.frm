VERSION 5.00
Begin VB.Form Form4 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Provider"
   ClientHeight    =   930
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3240
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   930
   ScaleWidth      =   3240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Provider"
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   3135
      Begin VB.CommandButton Command1 
         Caption         =   "OK!"
         Height          =   255
         Left            =   2280
         TabIndex        =   2
         Top             =   300
         Width           =   735
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   2055
      End
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Provider = Combo1.Text
    Me.Hide
    Form2.Show 1
End If
End Sub

Private Sub Command1_Click()
Provider = Combo1.Text
Me.Hide
Form2.Show 1
End Sub

Private Sub Form_Load()
Combo1.AddItem "SQL Server"
Combo1.AddItem "Ms Access 2000"
Combo1.AddItem "Ms Access 97"
Combo1.AddItem "Oracle"
Combo1.AddItem "Foxpro"
Combo1.Text = "SQL Server"
End Sub
