VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm frmMain 
   BackColor       =   &H000000C0&
   Caption         =   "Payroll System "
   ClientHeight    =   8145
   ClientLeft      =   735
   ClientTop       =   630
   ClientWidth     =   10845
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   240
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   27
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":225C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2F38
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":48CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":625C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7BEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9580
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A25A
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":AF34
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":BC0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":C8EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":D5C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":DEA2
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":EB7E
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":F85A
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":10536
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":10E1A
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":11AF6
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":123D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":130AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":14A42
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":163D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":16CB2
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1758E
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":183E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":18CBE
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1943A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList i16x16 
      Left            =   120
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":19BB6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10845
      _ExtentX        =   19129
      _ExtentY        =   1058
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      DisabledImageList=   "ImageList1"
      HotImageList    =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   14
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "employee"
            Object.ToolTipText     =   "Employee File"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "position"
            Object.ToolTipText     =   "Position"
            ImageIndex      =   18
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "department"
            Object.ToolTipText     =   "Department"
            ImageIndex      =   21
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "table"
            Object.ToolTipText     =   "Tables"
            ImageIndex      =   26
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   3
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "sss"
                  Text            =   "SSS Table        "
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "phil"
                  Text            =   "PhilHealth Table"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "with"
                  Text            =   "WithHolding Tax       "
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "transaction"
            Object.ToolTipText     =   "transaction"
            ImageIndex      =   14
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   6
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "mnuattendance"
                  Text            =   "&Attendance       "
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "pay"
                  Text            =   "PaySlip"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "rank"
                  Text            =   "&Rank"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "mnusep"
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "living"
                  Text            =   "Living Allowance       "
               EndProperty
               BeginProperty ButtonMenu6 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "rice"
                  Text            =   "Rice Allowance      "
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "backup"
            Object.ToolTipText     =   "Backup"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "changepass"
            Object.ToolTipText     =   "Change Password"
            ImageIndex      =   23
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "newuser"
            Object.ToolTipText     =   "New User"
            ImageIndex      =   17
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Reports"
            Object.ToolTipText     =   "Reports"
            ImageIndex      =   13
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "payslip"
                  Text            =   "PaySlip"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "rpt"
                  Text            =   "Payroll Report        "
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Help"
            Object.ToolTipText     =   "Help"
            ImageIndex      =   25
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   1
      Top             =   7860
      Width           =   10845
      _ExtentX        =   19129
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   14
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   441
            MinWidth        =   441
            Picture         =   "frmMain.frx":19F50
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   1587
            MinWidth        =   1587
            Text            =   "User Name:"
            TextSave        =   "User Name:"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3881
            MinWidth        =   3881
            Text            =   "Waiting..."
            TextSave        =   "Waiting..."
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   441
            MinWidth        =   441
            Picture         =   "frmMain.frx":1A4EA
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   1764
            MinWidth        =   1764
            Text            =   "Time Log-in:"
            TextSave        =   "Time Log-in:"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2893
            MinWidth        =   2893
            Text            =   "Waiting..."
            TextSave        =   "Waiting..."
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   529
            MinWidth        =   529
            Picture         =   "frmMain.frx":1A884
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "Date:"
            TextSave        =   "Date:"
         EndProperty
         BeginProperty Panel9 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3422
            MinWidth        =   3422
            Text            =   "[ DATE ]"
            TextSave        =   "[ DATE ]"
         EndProperty
         BeginProperty Panel10 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            Enabled         =   0   'False
            Object.Width           =   1235
            MinWidth        =   1235
            TextSave        =   "INS"
         EndProperty
         BeginProperty Panel11 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Object.Width           =   1235
            MinWidth        =   1235
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel12 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            Object.Width           =   1235
            MinWidth        =   1235
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel13 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   4
            Enabled         =   0   'False
            Object.Width           =   1235
            MinWidth        =   1235
            TextSave        =   "SCRL"
         EndProperty
         BeginProperty Panel14 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   6174
            MinWidth        =   6174
            Text            =   " Copyright (c) 2006"
            TextSave        =   " Copyright (c) 2006"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuemp 
         Caption         =   "&Employee File             "
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuPosition 
         Caption         =   "&Position            "
      End
      Begin VB.Menu mnudepart 
         Caption         =   "&Department           "
      End
      Begin VB.Menu mnusep31 
         Caption         =   "-"
      End
      Begin VB.Menu mnulock 
         Caption         =   "&Lock Application      "
         Shortcut        =   ^L
      End
      Begin VB.Menu mnusep43 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLogOff 
         Caption         =   ""
      End
      Begin VB.Menu mnusep12 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit Application    "
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuTransaction 
      Caption         =   "&Transaction"
      Begin VB.Menu mnuPayroll1 
         Caption         =   "&Payroll                  "
         Begin VB.Menu mnupayslip 
            Caption         =   "&PaySlip       "
         End
         Begin VB.Menu mnusep56 
            Caption         =   "-"
         End
         Begin VB.Menu mnuattendance 
            Caption         =   "&Attendance       "
         End
         Begin VB.Menu mnurank 
            Caption         =   "&Rank      "
            Shortcut        =   ^R
         End
         Begin VB.Menu mnusep545 
            Caption         =   "-"
         End
         Begin VB.Menu mnuLiving 
            Caption         =   "&Living Allowance       "
            Shortcut        =   {F3}
         End
         Begin VB.Menu mnurice 
            Caption         =   "&Rice Allowance"
            Shortcut        =   {F4}
         End
      End
      Begin VB.Menu mnutables 
         Caption         =   "&Tables"
         Begin VB.Menu mnussstable 
            Caption         =   "&SSS Table           "
            Shortcut        =   ^S
         End
         Begin VB.Menu mnuphilhealth 
            Caption         =   "&PhilHealth Table"
         End
         Begin VB.Menu mnutax 
            Caption         =   "&WithHolding Tax       "
            Shortcut        =   ^W
         End
      End
   End
   Begin VB.Menu mnuUtility 
      Caption         =   "&Utility"
      Begin VB.Menu mnuBackup 
         Caption         =   "&Backup..."
         Shortcut        =   ^B
      End
      Begin VB.Menu mnusep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuChangePassword 
         Caption         =   "&Change Password           "
      End
      Begin VB.Menu mnuUser 
         Caption         =   "&New User"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu mnureport 
      Caption         =   "&Reports   "
      Begin VB.Menu mnuslip 
         Caption         =   "Payslip"
      End
      Begin VB.Menu mnupayrollReport 
         Caption         =   "Payroll Report          "
         Shortcut        =   {F7}
      End
   End
   Begin VB.Menu mnuhelp 
      Caption         =   "&Help"
      Begin VB.Menu mnucontents 
         Caption         =   "&Contents...            "
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuathour 
         Caption         =   "The &Author                  "
         Shortcut        =   {F2}
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub MDIForm_Load()

    Me.Caption = "Payroll System ver. " & App.Major & "." & App.Minor
    
End Sub

Private Sub mnuathour_Click()
frmabout.Show
frmabout.SetFocus
End Sub

Private Sub mnuattendance_Click()
frmAttendance.Show
frmAttendance.SetFocus
End Sub

Private Sub mnuBackup_Click()
frmBackup.Show
frmBackup.SetFocus
End Sub

Private Sub mnuChangePassword_Click()

frmChangePass.Show
frmChangePass.SetFocus

End Sub

Private Sub mnucontents_Click()
    script.help_filepath ("\helpsms.chm")
End Sub

Private Sub mnudepart_Click()
    frmdepartment.Show
    frmdepartment.SetFocus
End Sub

Private Sub mnuemp_Click()
frmEmployeeFile.Show
frmEmployeeFile.SetFocus
End Sub

Private Sub mnuFileExit_Click()
  Unload Me
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    Unload Me
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Dim msg As Integer

    msg = MsgBox("This will terminate the application. Do you want to proceed?", vbExclamation + vbYesNo, "Confirm Exiting...")
    
    If msg = vbNo Then
    
        Cancel = 1
        
    Else
        
        '-// Close all Global connection (to database)
        CN.Close
        
        '-// Delete all *.tmp files from the computer
        Shell App.Path + "\deltemp.bat", vbHide

    End If
        
End Sub

Private Sub mnuLiving_Click()
frmliving.Show
frmliving.SetFocus
End Sub

Private Sub mnulock_Click()
frmlock.Show vbModal
End Sub

Private Sub mnuLogOff_Click()
 Dim msg As Integer

    msg = MsgBox("Are you sure you want to Log-off . Proceed?", vbExclamation + vbYesNo, "Confirm Log-Off...")
    
    If msg = vbYes Then
    
    '-// Unload all Forms
    Unload frmabout
    Unload frmAttendance
    Unload frmBackup
    Unload frmChangePass
    Unload frmdepartment
    Unload frmEmployeeFile
    Unload frmliving
    Unload frmNewUser
    Unload frmPayroll
    Unload frmpayrollReport
    Unload frmPayslip
    Unload frmPhilHealth
    Unload frmposition
    Unload frmrank
    Unload frmriceallowance
    Unload frmsearchEmp
    Unload frmselectrank
    Unload frmSplash
    Unload frmssstable
    Unload frmUpdate
    Unload frmWithTax
       
    
    FrmLogOff.Show vbModal
  
    End If
    
End Sub


Private Sub mnupayrollReport_Click()
frmpayrollReport.Show
frmpayrollReport.SetFocus
End Sub

Private Sub mnupayslip_Click()
frmPayroll.Show
frmPayroll.SetFocus
End Sub

Private Sub mnuphilhealth_Click()
frmPhilHealth.Show
frmPhilHealth.SetFocus
End Sub

Private Sub mnuPosition_Click()
frmposition.Show
frmposition.SetFocus
End Sub

Private Sub mnurank_Click()
frmrank.Show
frmrank.SetFocus
End Sub

Private Sub mnurice_Click()
frmriceallowance.Show
frmriceallowance.SetFocus
End Sub

Private Sub mnuslip_Click()
frmPayslip.Show
frmPayslip.SetFocus
End Sub

Private Sub mnussstable_Click()
frmssstable.Show
frmssstable.SetFocus
End Sub

Private Sub mnutax_Click()
frmWithTax.Show
frmWithTax.SetFocus
End Sub

Private Sub mnuUser_Click()
    frmNewUser.Show
    frmNewUser.SetFocus
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case Button.Key
    
    Case "employee"
        mnuemp_Click
        
    Case "position"
        mnuPosition_Click
        
    Case "department"
        mnudepart_Click

    Case "backup"
        mnuBackup_Click
    
    Case "changepass"
        mnuChangePassword_Click
        
    Case "newuser"
        mnuUser_Click
    
    Case "Help"
        mnucontents_Click
        
End Select

End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    
   Select Case ButtonMenu.Key
   
    Case "pay"
        mnupayslip_Click
        
    Case "mnuattendance"
        mnuattendance_Click
        
    Case "rank"
        mnurank_Click
    
    Case "living"
        mnuLiving_Click
        
    Case "rice"
        mnurice_Click
    
    Case "sss"
        mnussstable_Click
    
    Case "phil"
        mnuphilhealth_Click
    
    Case "with"
        mnutax_Click
    
    End Select
    
End Sub
