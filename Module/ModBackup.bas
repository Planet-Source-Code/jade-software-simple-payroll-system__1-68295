Attribute VB_Name = "ModBackup"
Option Explicit

Sub backup_rec()

Dim datengayon As String
datengayon = Trim(Format(Date, "mm/dd/yy"))
script.CopyFile App.Path & "\Database\payroll.mdb", frmBackup.Dir1.Path & "\backup" & Left(datengayon, 2) & Mid(datengayon, 4, 2) & Right(datengayon, 2) & ".sms"

End Sub

Sub restore_rec()

script.CopyFile frmBackup.Dir1.Path & "\" & frmBackup.Text1, App.Path & "\Database\payroll.mdb"

End Sub


