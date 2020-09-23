VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Database Selection"
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3420
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   3420
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Selection Of Database"
      Height          =   3495
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   3375
      Begin VB.CommandButton Command1 
         Caption         =   "Ok!"
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   3000
         Width           =   2895
      End
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   2580
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   3135
      End
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
db = List1.Text
GetTables
Unload Form3
End Sub

Private Sub Form_Activate()
Form3.List1.Selected(0) = True
End Sub
    
Private Sub List1_DblClick()
db = List1.Text
GetTables
Unload Form3
End Sub
