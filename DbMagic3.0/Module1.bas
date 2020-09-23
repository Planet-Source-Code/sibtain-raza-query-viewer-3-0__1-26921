Attribute VB_Name = "Module1"
'''''''''''''''''''''''''''''''''''''''''''''''''''''
'          Query Viewer - Ver - 3.0                 '
'                                                   '
'    Includes SQL Server, Oracle, Ms Access2000     '
'                 Ms Access 97 And Foxpro           '
'                                                   '
'                                                   '
' If Your Are Intrested In Making Softwares Please  '
'        Contact razasibtain@hotmail.com            '
'                                                   '
'       Now Supports Foxpro And Oracle also         '
'''''''''''''''''''''''''''''''''''''''''''''''''''''
Public db As String
Public cn As New Connection
Public ServerName As String
Public Provider As String
Public uname As String
Public pass As String

Public Sub LogonServer(Provider As String)
On Error Resume Next
If Provider = "SQL Server" Then
    cn.ConnectionString = ""
    cn.Provider = "MSDASQL;Driver={SQL Server};SERVER=" & ServerName & ";user id=" & uname & ";Password=" & pass & ";Database=" & db & ""
    cn.Open
    If Err.Number <> 0 Then
        MsgBox Err.Description
        Exit Sub
    End If
    GetDatabase (Provider)
End If

If Provider = "Ms Access 2000" Or Provider = "Ms Access 97" Then

    If Provider = "Ms Access 2000" Then
        cn.Provider = "Microsoft.Jet.Oledb.4.0.Provider"
    End If
    If Provider = "Ms Access 97" Then
        cn.Provider = "Microsoft.Jet.Oledb.3.51.Provider"
    End If
        cn.ConnectionString = db
        cn.Open
        If Err.Number <> 0 Then
            MsgBox Err.Description
            Exit Sub
        End If
        GetDatabase (Provider)
End If

If Provider = "Oracle" Then
    cn.ConnectionString = ""
    cn.Provider = "MSDAORA.1;Data Source=" & ServerName & ";user id=" & uname & ";Password=" & pass & ";Database=" & db & ""
    cn.Open
    If Err.Number <> 0 Then
        MsgBox Err.Description
        Exit Sub
   End If
    GetDatabase (Provider)
End If

If Provider = "Foxpro" Then
    cn.ConnectionString = ""
    cn.Provider = "MSDASQL.1;Data Source=dBASE Files;Initial Catalog=" & db
    cn.Open
        If Err.Number <> 0 Then
        MsgBox Err.Description
        Exit Sub
   End If
   GetDatabase (Provider)
End If
End Sub

Public Sub GetDatabase(Provider As String)
On Error Resume Next
If Provider = "SQL Server" Then
    Set rs = cn.Execute("sp_databases")
    Unload Form2
    Do While Not rs.EOF
        Form3.List1.AddItem rs.Fields(0)
        rs.MoveNext
    Loop
    Form3.Show 1
'    rs.Close
End If

If Provider = "Oracle" Then
    Set rs = cn.Execute("select * from cat")
    Form1.List1.Clear
    Do While Not rs.EOF
        Form1.List1.AddItem rs("table_name")
        rs.MoveNext
    Loop
 'rs.close
    Form1.Show
    Form1.List1.Enabled = True
    Form1.Command1.Enabled = True
    Form1.Command2.Enabled = True
    Form1.Command3.Enabled = True
    Form1.Command4.Enabled = True
    Form1.Text1.Enabled = True
End If

If Provider = "Ms Access 2000" Or Provider = "Ms Access 97" Then
    Set rs = cn.OpenSchema(adSchemaTables)
    Form1.List1.Clear
    Do While Not rs.EOF
        If Left(rs("table_name"), 4) <> "MSys" Then
            Form1.List1.AddItem rs("table_name")
        End If
        rs.MoveNext
    Loop
    Form1.Show
'    rs.Close
    Form1.List1.Enabled = True
    Form1.Command1.Enabled = True
    Form1.Command2.Enabled = True
    Form1.Command3.Enabled = False
    Form1.Command4.Enabled = True
    Form1.Text1.Enabled = True
End If

If Provider = "Foxpro" Then
    Set rs = cn.OpenSchema(adSchemaTables)
    Form1.List1.Clear
    Do While Not rs.EOF
        Form1.List1.AddItem rs("table_name")
        rs.MoveNext
    Loop
    Form1.Show
    Form1.List1.Enabled = True
    Form1.Command1.Enabled = True
    Form1.Command2.Enabled = False
    Form1.Command3.Enabled = True
    Form1.Command4.Enabled = True
    Form1.Text1.Enabled = True
End If
End Sub

Public Sub GetTables()
cn.Close
LogonServer (Provider)
Form1.List1.Clear
Dim rs As New Recordset
rs.Open "Select * from sysobjects where xtype='U'", cn, adOpenForwardOnly, adLockOptimistic
Do While Not rs.EOF
    Form1.List1.AddItem rs!Name
    rs.MoveNext
Loop
rs.Close
Form1.List1.Enabled = True
Form1.Command1.Enabled = True
Form1.Command2.Enabled = True
Form1.Command3.Enabled = True
Form1.Command4.Enabled = True
Form1.Text1.Enabled = True
End Sub

