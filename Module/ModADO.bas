Attribute VB_Name = "ModADO"
Option Explicit

'-// SCripting Runtime for Backup, Compute Bday, Call CHM files
Public script           As New cls_RJSoft_Script

'-// For Username & Password
Public Uname            As String
Public Pword            As String

'-// Global Connection
Global CN               As New Connection
Global Rs               As New ADODB.Recordset

'-// For Backup and Restore
Public EnableBackup     As Boolean
'-// For Login
Public rsLogin          As New ADODB.Recordset
'-// For NewUser/ChangePass/attendance
Public Adors            As New ADODB.Recordset
'-// For Employee file Form
Public adoEmp           As New ADODB.Recordset
Public adorsEmp         As New ADODB.Recordset
Public adoEmpSearch     As New ADODB.Recordset

'-// For Rank select
Public adoSelectRank          As New ADODB.Recordset

'-// For attendance
Public adoAttend        As New ADODB.Recordset


'-// For department Form
Public rsDepart         As New ADODB.Recordset
Public rsRec            As New ADODB.Recordset '//For deleting department/position

'-// For Position
Public rsPosition       As New ADODB.Recordset
Public rsdelrankpos     As New ADODB.Recordset

'-// For Ranking
Public rsRank           As New ADODB.Recordset
Public rsRankSave       As New ADODB.Recordset

'-// For WithTax
Public RsTax            As New ADODB.Recordset

'-// For Payroll Report Grid
Public rsDGridPayroll   As New ADODB.Recordset

'-// For Payroll Report Printing
Public rsgridprint      As New ADODB.Recordset

Public RsPayroll        As New ADODB.Recordset

Public Sub getconnect()

 On Error GoTo ErrHandler
 
    CN.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data source=" & App.Path & "\Database\Payroll.mdb;"
    CN.Properties("Jet OLEDB:Database Password") = "smspayroll"
    CN.Open
 
    Exit Sub
    
ErrHandler:
    MsgBox Err.Number & " " & Err.Description, vbCritical, "Connection Error..."
    
End Sub

