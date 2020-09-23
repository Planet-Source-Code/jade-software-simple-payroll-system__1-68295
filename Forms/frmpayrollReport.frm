VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmpayrollReport 
   BackColor       =   &H000000C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Payroll Report"
   ClientHeight    =   4185
   ClientLeft      =   30
   ClientTop       =   420
   ClientWidth     =   8880
   Icon            =   "frmpayrollReport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   8880
   Begin VB.CommandButton CmdClose 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7920
      TabIndex        =   4
      Top             =   3720
      Width           =   855
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H000000C0&
      Caption         =   "13 Month ?"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   255
      Left            =   4920
      TabIndex        =   8
      Top             =   120
      Width           =   1455
   End
   Begin MSDataGridLib.DataGrid Dgrid 
      Height          =   2772
      Left            =   240
      TabIndex        =   7
      Top             =   720
      Width           =   8412
      _ExtentX        =   14843
      _ExtentY        =   4895
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      HeadLines       =   1
      RowHeight       =   19
      WrapCellPointer =   -1  'True
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Payroll Report"
      ColumnCount     =   15
      BeginProperty Column00 
         DataField       =   "Ecode"
         Caption         =   "Employee Code"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "fullname"
         Caption         =   "Employee Name"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "com_basic"
         Caption         =   "Semi-Monthly"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "living_allowance"
         Caption         =   "Living Allowance"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "rice_allowance"
         Caption         =   "Rice Allowance"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "honorarium"
         Caption         =   "Honorarium"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "sss_premium"
         Caption         =   "SSS Premium"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "phil_health"
         Caption         =   "PhilHealth"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "sss_loan"
         Caption         =   "SSS Loan"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column09 
         DataField       =   "withholding_tax"
         Caption         =   "WithHolding Tax"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column10 
         DataField       =   "pagibig"
         Caption         =   "Pag-ibig"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column11 
         DataField       =   "other_deduction"
         Caption         =   "Other Deduction"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column12 
         DataField       =   "gross_pay"
         Caption         =   "Gross Pay"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   """$""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column13 
         DataField       =   "total_deduction"
         Caption         =   "Total Deduction"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   """$""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column14 
         DataField       =   "net_pay"
         Caption         =   "Net Pay"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   """$""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            Locked          =   -1  'True
            ColumnWidth     =   1305.071
         EndProperty
         BeginProperty Column01 
         EndProperty
         BeginProperty Column02 
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1409.953
         EndProperty
         BeginProperty Column04 
         EndProperty
         BeginProperty Column05 
         EndProperty
         BeginProperty Column06 
         EndProperty
         BeginProperty Column07 
         EndProperty
         BeginProperty Column08 
         EndProperty
         BeginProperty Column09 
         EndProperty
         BeginProperty Column10 
         EndProperty
         BeginProperty Column11 
         EndProperty
         BeginProperty Column12 
         EndProperty
         BeginProperty Column13 
         EndProperty
         BeginProperty Column14 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton CmdPrint 
      Caption         =   "&Print"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7080
      TabIndex        =   5
      Top             =   3720
      Width           =   855
   End
   Begin VB.CommandButton cmdCompute 
      Caption         =   "&Compute"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6240
      TabIndex        =   6
      Top             =   3720
      Width           =   855
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   288
      Left            =   1560
      TabIndex        =   0
      Top             =   120
      Width           =   1572
      _ExtentX        =   2778
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   24444929
      CurrentDate     =   38695
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   288
      Left            =   3240
      TabIndex        =   1
      Top             =   120
      Width           =   1572
      _ExtentX        =   2778
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   24444929
      CurrentDate     =   38695
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "&Select"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7920
      TabIndex        =   9
      Top             =   3720
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   3
      Height          =   3012
      Left            =   120
      Top             =   600
      Width           =   8652
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0E0FF&
      Height          =   3012
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   8652
   End
   Begin VB.Label Label1 
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Payroll Period :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   252
      Left            =   120
      TabIndex        =   2
      Top             =   132
      Width           =   1572
   End
End
Attribute VB_Name = "frmpayrollReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Check1_Click()

If Check1.Value = 1 Then


Set rsDGridPayroll = New ADODB.Recordset
   rsDGridPayroll.CursorLocation = adUseClient
rsDGridPayroll.Open "SELECT *, [gross_pay]-[total_deduction]+[basic_pay] As net_pay FROM [qryrpt] WHERE payroll_period Between #" & DTPicker1.Value & "# AND #" & DTPicker2.Value & "# ORDER BY ecode", CN, adOpenStatic, adLockOptimistic
Set Dgrid.DataSource = rsDGridPayroll
Dgrid.Refresh

Else

Set rsDGridPayroll = New ADODB.Recordset
   rsDGridPayroll.CursorLocation = adUseClient
rsDGridPayroll.Open "SELECT *, [gross_pay]-[total_deduction] As net_pay FROM [qryrpt] WHERE payroll_period Between #" & DTPicker1.Value & "# AND #" & DTPicker2.Value & "# ORDER BY ecode", CN, adOpenStatic, adLockOptimistic
Set Dgrid.DataSource = rsDGridPayroll
Dgrid.Refresh

 
End If


End Sub

Private Sub cmdCompute_Click()

If Check1.Value = 1 Then

Set rsDGridPayroll = New ADODB.Recordset
   rsDGridPayroll.CursorLocation = adUseClient
rsDGridPayroll.Open "SELECT *, [gross_pay]-[total_deduction]+[basic_pay] As net_pay FROM [qryrpt] WHERE payroll_period Between #" & DTPicker1.Value & "# AND #" & DTPicker2.Value & "# ORDER BY ecode", CN, adOpenStatic, adLockOptimistic
Set Dgrid.DataSource = rsDGridPayroll
Dgrid.ReBind
Dgrid.Refresh


Else

'-// Without 13 Month
Set rsDGridPayroll = New ADODB.Recordset
   rsDGridPayroll.CursorLocation = adUseClient
rsDGridPayroll.Open "SELECT *, [gross_pay]-[total_deduction] As net_pay FROM [qryrpt] WHERE payroll_period Between #" & DTPicker1.Value & "# AND #" & DTPicker2.Value & "# ORDER BY ecode", CN, adOpenStatic, adLockOptimistic
Set Dgrid.DataSource = rsDGridPayroll
Dgrid.Refresh
End If



End Sub

Private Sub cmdPrint_Click()

 Set RptPayroll.DataSource = rsDGridPayroll
 RptPayroll.Sections("Section4").Controls.Item("label27").Caption = "Payroll Period : FROM : " & Format(DTPicker1.Value, "MM-DD-YY") & " TO : " & Format(DTPicker2.Value, "MM-DD-YY")
 RptPayroll.Show
 RptPayroll.SetFocus

End Sub

Private Sub Form_Load()

Set rsDGridPayroll = New ADODB.Recordset
rsDGridPayroll.CursorLocation = adUseClient
rsDGridPayroll.Open "SELECT * FROM [computation]", CN, adOpenStatic, adLockPessimistic
Set Dgrid.DataSource = rsDGridPayroll
Dgrid.Refresh

End Sub

Private Sub Form_Unload(Cancel As Integer)

'-// Clear variables in computer memory
Set rsDGridPayroll = Nothing

End Sub

