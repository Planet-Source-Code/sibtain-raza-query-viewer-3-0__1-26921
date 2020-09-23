VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Form1 
   Caption         =   "Query Viewer - Ver 3.0 - www.cispl.com"
   ClientHeight    =   8340
   ClientLeft      =   60
   ClientTop       =   1695
   ClientWidth     =   11880
   Icon            =   "main.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8340
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame4 
      Caption         =   "Table Information"
      Height          =   2295
      Left            =   9000
      TabIndex        =   11
      Top             =   4560
      Width           =   2775
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1560
         TabIndex        =   16
         Top             =   1440
         Width           =   60
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1560
         TabIndex        =   15
         Top             =   960
         Width           =   60
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Total Fields :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Total Records :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1275
         TabIndex        =   12
         Top             =   360
         Width           =   105
      End
   End
   Begin VB.ComboBox Text1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   0
      TabIndex        =   10
      Top             =   6960
      Width           =   8895
   End
   Begin VB.Frame Frame3 
      Caption         =   "Information"
      Height          =   975
      Left            =   8880
      TabIndex        =   5
      Top             =   7320
      Width           =   3015
      Begin VB.Label Label1 
         Caption         =   "This Query Viewer Is Jointly Designed By Sibtain, Zaheer Abbas, Aamir Saeed And Imran"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   2775
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Advance Features"
      Height          =   855
      Left            =   0
      TabIndex        =   4
      Top             =   7320
      Width           =   8775
      Begin VB.CommandButton Command4 
         Caption         =   "Change Provider"
         Height          =   375
         Left            =   3240
         TabIndex        =   9
         Top             =   360
         Width           =   1455
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Change Database"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1680
         TabIndex        =   7
         Top             =   360
         Width           =   1455
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Change Server"
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   1455
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid 
      Height          =   6855
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   12091
      _Version        =   393216
      BackColor       =   15724527
      ForeColor       =   0
      FixedCols       =   0
      GridColor       =   -2147483627
      AllowUserResizing=   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
      _Band(0).GridLineWidthBand=   1
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Run!"
      Enabled         =   0   'False
      Height          =   375
      Left            =   9000
      TabIndex        =   0
      Top             =   6960
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tables List"
      Height          =   4575
      Left            =   9000
      TabIndex        =   1
      Top             =   0
      Width           =   2775
      Begin VB.ListBox List1 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   4050
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   2535
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Msg As String
Private Sub Command1_Click()
Text1.AddItem Text1.Text
Msg = "True"
RunQuery
End Sub

Private Sub Command2_Click()
Form2.Show 1
End Sub

Private Sub Command3_Click()
If Provider = "SQL Server" Or Provider = "Oracle" Then
    GetTables
End If
If Provider = "Ms Access 2000" Or Provider = "Ms Access 97" Or Provider = "Foxpro" Then
    Form2.Show 1
End If
End Sub

Private Sub Command4_Click()
Form4.Show 1
End Sub

Private Sub Form_Load()
Form1.Show
Form4.Show 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
If cn.State = 1 Then
    cn.Close
End If
End
End Sub

Private Sub RunQuery()
On Error Resume Next
Grid.Clear
If Text1.Text = "" Then
    MsgBox "Please Enter Query", vbCritical, "Help!"
    Exit Sub
End If

Dim rs As New Recordset
rs.Open Text1.Text, cn, adOpenKeyset, adLockOptimistic, 1

If Err.Number <> 0 Then
    MsgBox Err.Description
    Exit Sub
End If

Set Grid.DataSource = rs
Label5.Caption = rs.RecordCount
Label6.Caption = rs.Fields.Count

rs.Close
Text1.SelStart = 0
Text1.SelLength = Len(Text1.Text)

If Msg = "True" Then
    If UCase(Left(Text1.Text, 6)) <> "SELECT" Then
        If Provider = "Oracle" Then
            rs.Open "select * from " & List1.Text, cn, adOpenForwardOnly, adLockOptimistic, 1
        Else
            rs.Open "select * from " & "[" & List1.Text & "]", cn, adOpenForwardOnly, adLockOptimistic, 1
        End If
        Set Grid.DataSource = rs
        Msg = "False"
    End If
End If
End Sub

Private Sub List1_Click()
If Provider = "Oracle" Then
    Text1.Text = "Select * from " & List1.Text
Else
    Text1.Text = "Select * from " & "[" & List1.Text & "]"
End If
Label2.Caption = List1.Text
RunQuery
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text1.AddItem Text1.Text
    Msg = "True"
    RunQuery
End If
End Sub
