VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPayroll 
   BackColor       =   &H000000C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Payroll Information"
   ClientHeight    =   5256
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   6084
   Icon            =   "frmPayroll.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5256
   ScaleWidth      =   6084
   Begin VB.TextBox TxtODeduction 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4320
      TabIndex        =   43
      Top             =   3840
      Width           =   1335
   End
   Begin VB.TextBox TxtPagibig 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4320
      Locked          =   -1  'True
      TabIndex        =   42
      Text            =   "50.00"
      Top             =   3480
      Width           =   1335
   End
   Begin VB.TextBox TxtTax 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   41
      Top             =   3840
      Width           =   1332
   End
   Begin VB.TextBox TxtLoan 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1560
      TabIndex        =   40
      Top             =   3480
      Width           =   1335
   End
   Begin VB.TextBox TxtSSS 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   39
      Top             =   3120
      Width           =   1332
   End
   Begin VB.TextBox TxtPhilHealth 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4320
      Locked          =   -1  'True
      TabIndex        =   38
      Top             =   3120
      Width           =   1320
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000C&
      Height          =   285
      Left            =   4440
      Locked          =   -1  'True
      TabIndex        =   34
      Top             =   1560
      Width           =   1452
   End
   Begin VB.TextBox TxtHonorarium 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4440
      TabIndex        =   28
      Top             =   2256
      Width           =   1452
   End
   Begin VB.TextBox TxtLivingAllow 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4440
      Locked          =   -1  'True
      TabIndex        =   27
      Top             =   1920
      Width           =   1452
   End
   Begin VB.TextBox TxtRiceAllow 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   26
      Top             =   2280
      Width           =   1335
   End
   Begin VB.TextBox TxtBasicPay 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   25
      Top             =   1920
      Width           =   1335
   End
   Begin VB.TextBox Txtmonthly 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000C&
      Height          =   285
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   24
      Top             =   1560
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   5280
      TabIndex        =   22
      Text            =   "Text2"
      Top             =   6240
      Width           =   735
   End
   Begin VB.TextBox txtresult 
      Height          =   285
      Left            =   4680
      TabIndex        =   21
      Text            =   "13+gpay"
      Top             =   6600
      Width           =   975
   End
   Begin VB.TextBox txtcomputed_gpay 
      Height          =   285
      Left            =   1800
      TabIndex        =   20
      Text            =   "txtcomputed_gpay"
      Top             =   6600
      Width           =   1575
   End
   Begin VB.TextBox txttempmonth 
      Height          =   285
      Left            =   3480
      TabIndex        =   19
      Text            =   "13month"
      Top             =   6600
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   18
      Text            =   "txttemphonorarium"
      Top             =   6840
      Width           =   1455
   End
   Begin VB.TextBox TxtYears 
      Height          =   285
      Left            =   240
      TabIndex        =   16
      Top             =   6240
      Width           =   495
   End
   Begin VB.CommandButton CmdClose 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      TabIndex        =   2
      Top             =   4800
      Width           =   855
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4320
      TabIndex        =   3
      Top             =   4800
      Width           =   855
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      TabIndex        =   15
      Top             =   4800
      Width           =   855
   End
   Begin VB.CommandButton CmdUpdate 
      Caption         =   "&Update"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   4
      Top             =   4800
      Width           =   855
   End
   Begin VB.CommandButton CmdEdit 
      Caption         =   "&Edit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   5
      Top             =   4800
      Width           =   855
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   14
      Top             =   4800
      Width           =   855
   End
   Begin VB.ComboBox cmbecode 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1560
      TabIndex        =   13
      Top             =   480
      Width           =   2535
   End
   Begin VB.CommandButton CmdLast 
      Caption         =   ">>"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   265
      Left            =   4560
      TabIndex        =   9
      ToolTipText     =   "Last"
      Top             =   4440
      Width           =   1455
   End
   Begin VB.CommandButton CmdNext 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   265
      Left            =   3120
      TabIndex        =   10
      ToolTipText     =   "Next"
      Top             =   4440
      Width           =   1455
   End
   Begin VB.CommandButton CmdPrevious 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   265
      Left            =   1680
      TabIndex        =   11
      ToolTipText     =   "Previous"
      Top             =   4440
      Width           =   1455
   End
   Begin VB.CommandButton CmdFirst 
      Caption         =   "<<"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   265
      Left            =   120
      TabIndex        =   12
      ToolTipText     =   "First"
      Top             =   4440
      Width           =   1575
   End
   Begin VB.TextBox Txtename 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1560
      TabIndex        =   1
      Top             =   880
      Width           =   4455
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Left            =   1560
      TabIndex        =   0
      Top             =   90
      Width           =   2535
      _ExtentX        =   4466
      _ExtentY        =   550
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   43515905
      CurrentDate     =   38673
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Deductions"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   252
      Left            =   120
      TabIndex        =   50
      Top             =   2760
      Width           =   1692
   End
   Begin VB.Label LblODeduction 
      BackColor       =   &H000000C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Other Deduction"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   252
      Left            =   3000
      TabIndex        =   49
      Top             =   3876
      Width           =   1332
   End
   Begin VB.Label LblPagibig 
      BackColor       =   &H000000C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Pag-ibig"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   252
      Left            =   3000
      TabIndex        =   48
      Top             =   3516
      Width           =   1212
   End
   Begin VB.Label LblPhilHealth 
      BackColor       =   &H000000C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Phil Health"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   252
      Left            =   3000
      TabIndex        =   47
      Top             =   3156
      Width           =   1332
   End
   Begin VB.Label Label9 
      BackColor       =   &H000000C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Withholding Tax"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   252
      Left            =   240
      TabIndex        =   46
      Top             =   3876
      Width           =   1212
   End
   Begin VB.Label Label3 
      BackColor       =   &H000000C0&
      BackStyle       =   0  'Transparent
      Caption         =   "SSS Loan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   252
      Left            =   240
      TabIndex        =   45
      Top             =   3516
      Width           =   1212
   End
   Begin VB.Label LblSSS 
      BackColor       =   &H000000C0&
      BackStyle       =   0  'Transparent
      Caption         =   "SSS Premium"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   252
      Left            =   240
      TabIndex        =   44
      Top             =   3156
      Width           =   1212
   End
   Begin VB.Shape Shape2 
      BorderWidth     =   3
      Height          =   1332
      Left            =   120
      Top             =   3000
      Width           =   5892
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0E0FF&
      Height          =   1332
      Left            =   120
      TabIndex        =   37
      Top             =   3000
      Width           =   5892
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Semi-Monthly rate"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   252
      Left            =   3120
      TabIndex        =   35
      Top             =   1560
      Width           =   1332
   End
   Begin VB.Label LblHonorarium 
      BackColor       =   &H008080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Honorarium"
      ForeColor       =   &H00000000&
      Height          =   252
      Left            =   3120
      TabIndex        =   33
      Top             =   2280
      Width           =   1212
   End
   Begin VB.Label LblLivingAllow 
      BackColor       =   &H008080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Living Allowance"
      ForeColor       =   &H00000000&
      Height          =   252
      Left            =   3120
      TabIndex        =   32
      Top             =   1920
      Width           =   1212
   End
   Begin VB.Label LblRiceAllow 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Rice Allowance"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   252
      Left            =   360
      TabIndex        =   31
      Top             =   2280
      Width           =   1212
   End
   Begin VB.Label LblBasicPay 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Basic Pay"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   252
      Left            =   360
      TabIndex        =   30
      Top             =   1920
      Width           =   1212
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Monthly rate"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   252
      Left            =   360
      TabIndex        =   29
      Top             =   1560
      Width           =   1212
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   3
      Height          =   1212
      Left            =   120
      Top             =   1440
      Width           =   5892
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0E0FF&
      Height          =   1212
      Left            =   120
      TabIndex        =   23
      Top             =   1440
      Width           =   5892
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Years of Service"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   6600
      Width           =   1335
   End
   Begin VB.Label LblPayrollPeriod 
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Payroll Period"
      ForeColor       =   &H00C0E0FF&
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label LblEname 
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Name"
      ForeColor       =   &H00C0E0FF&
      Height          =   252
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Width           =   1332
   End
   Begin VB.Label Label1 
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Code"
      ForeColor       =   &H00C0E0FF&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Salary"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   252
      Left            =   120
      TabIndex        =   36
      Top             =   1200
      Width           =   1692
   End
End
Attribute VB_Name = "frmPayroll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim addflag As Boolean
Dim living, rice

Private Sub cmbecode_Click()

 Set rsRec = New ADODB.Recordset
    
    rsRec.Open "SELECT * FROM employeefile where ecode='" & cmbecode.Text & "'", CN, adOpenStatic, adLockPessimistic
    
        If rsRec.RecordCount >= 1 Then
        
        Txtename.Text = rsRec!sname & ", " & rsRec!fname & " " & rsRec!mi
        Txtmonthly.Text = Val(rsRec!basicpay) & " "
        Text3.Text = Val(rsRec!semi_monthly) & " "
        TxtLivingAllow.Text = Val(rsRec!living_value) & " "
        TxtRiceAllow.Text = Val(rsRec!rice_value) & " "
        TxtSSS.Text = Val(rsRec!SSSPremium) & " "
        TxtPhilHealth.Text = Val(rsRec!PHHealthValue) & " "
        TxtTax.Text = Val(rsRec!withTaxvalue) & " "
        
  End If
  
    Set rsRec = Nothing
    

Set rsRec = New ADODB.Recordset
    
    rsRec.Open "SELECT * FROM attendance where ecode='" & cmbecode.Text & "'", CN, adOpenStatic, adLockPessimistic
    
     If rsRec.RecordCount >= 1 Then TxtBasicPay.Text = rsRec!basic_pay
 
 Set rsRec = Nothing
  
End Sub

Private Sub Form_Load()

Me.Show
'cmbecode.Locked = True
Call ecode_load
Set RsPayroll = New ADODB.Recordset
RsPayroll.Open "SELECT * FROM computation", CN, adOpenKeyset, adLockOptimistic

If RsPayroll.BOF And RsPayroll.EOF Then MsgBox "No record(s) to display!", vbCritical, "ERROR": disable_nav_button: Exit Sub
Display_rec_payroll

ToggleButtonUpdateCancel "OFF"
End Sub

Sub Display_rec_payroll()

On Error Resume Next

With RsPayroll

    cmbecode = !ECODE & " "
    Txtename = !fullname & " "
    TxtBasicPay = !basic_pay & " "
    TxtLivingAllow = !living_allowance & " "
    TxtRiceAllow = !rice_allowance & " "
    TxtHonorarium = !honorarium & " "
    TxtPhilHealth = !Phil_Health & " "
    TxtSSS = !SSS_premium & " "
    TxtLoan = !SSS_loan & " "
    TxtPagibig = !Pagibig & " "
    TxtTax = !withholding_tax & " "
    TxtODeduction = !other_deduction & " "
    DTPicker1.Value = !payroll_period & " "
    
End With

End Sub

Sub ecode_load()


   Set Rs = New ADODB.Recordset
   Rs.Open "SELECT * FROM [employeefile]", CN, adOpenKeyset, adLockOptimistic
  
   Do Until Rs.EOF
   cmbecode.AddItem Rs!ECODE
   Rs.MoveNext
   Loop
   Set Rs = Nothing

End Sub

Private Sub Form_Unload(Cancel As Integer)

Set RsPayroll = Nothing

End Sub

Private Sub CmdFirst_Click()

On Error Resume Next
    RsPayroll.MoveFirst
    Display_rec_payroll
    If RsPayroll.BOF Then Exit Sub
    
End Sub

Private Sub CmdLast_Click()

On Error Resume Next
    RsPayroll.MoveLast
    Display_rec_payroll
    If RsPayroll.EOF Then Exit Sub
       
End Sub

Private Sub CmdNext_Click()

On Error Resume Next
 RsPayroll.MoveNext
 Display_rec_payroll
        If Adors.EOF Then
             MsgBox "You have reached the End Of File!", vbCritical + vbOKOnly
             RsPayroll.MoveLast
        End If
End Sub


Private Sub CmdPrevious_Click()
 On Error Resume Next
    RsPayroll.MovePrevious
    Display_rec_payroll

    If RsPayroll.BOF Then
             MsgBox "You have reached the Beginning Of File!", vbCritical + vbOKOnly
             RsPayroll.MoveFirst
        End If
End Sub

Private Sub CmdAdd_Click()

    clear_txt
    cmbecode.Locked = False
    TxtPagibig.Text = "50.00"
    addflag = True
    ToggleButtonUpdateCancel "ON"
 
End Sub

Private Sub CmdClose_Click()
Unload Me
End Sub

Private Sub cmdcancel_Click()
 
 cmbecode.Locked = True
 ToggleButtonUpdateCancel "OFF"
 If RsPayroll.BOF And RsPayroll.EOF Then disable_nav_button: Exit Sub
 Display_rec_payroll
 
End Sub

Private Sub CmdEdit_Click()

If RsPayroll.BOF And RsPayroll.EOF Then MsgBox "No record(s) to edit", vbCritical, "ERROR":  Exit Sub
ToggleButtonUpdateCancel "ON"
addflag = False

End Sub

Private Sub ToggleButtonUpdateCancel(OnorOff As String)
 If UCase(OnorOff) = "ON" Then 'Click Add or Edit
    
    cmdAdd.Enabled = False
    CmdUpdate.Enabled = True
    CmdCancel.Enabled = True
    
    CmdEdit.Enabled = False
    CmdClose.Enabled = False
    
    CmdFirst.Enabled = False
    CmdPrevious.Enabled = False
    CmdNext.Enabled = False
    CmdLast.Enabled = False
    
    TxtHonorarium.Locked = False
    TxtLoan.Locked = False
    TxtODeduction.Locked = False
    
    
Else
    
     cmdAdd.Enabled = True
    CmdUpdate.Enabled = False
    CmdCancel.Enabled = False
       
    CmdEdit.Enabled = True
    CmdClose.Enabled = True
    
    CmdFirst.Enabled = True
    CmdPrevious.Enabled = True
    CmdNext.Enabled = True
    CmdLast.Enabled = True
     
  
    TxtHonorarium.Locked = True
    TxtLoan.Locked = True
    TxtODeduction.Locked = True
    
          
End If
End Sub

Sub clear_txt()

TxtPhilHealth.Text = ""
Txtmonthly.Text = ""
Text3.Text = ""
TxtBasicPay.Text = ""
TxtLivingAllow.Text = ""
TxtRiceAllow.Text = ""
TxtSSS.Text = ""
TxtTax.Text = ""
TxtODeduction.Text = ""
TxtLoan.Text = ""
TxtHonorarium.Text = 0

End Sub


Sub disable_nav_button()

CmdFirst.Enabled = False
CmdNext.Enabled = False
CmdLast.Enabled = False
CmdPrevious.Enabled = False

End Sub

Private Sub CmdUpdate_Click()

If TxtHonorarium.Text = "" Then TxtHonorarium = 0
If IsNumeric(TxtLoan) = False Then MsgBox "Invalid input. Please check the value!", vbCritical, "ERROR": TxtLoan.SetFocus: Exit Sub
If IsNumeric(TxtODeduction) = False Then MsgBox "Invalid input. Please check the value!", vbCritical, "ERROR": TxtODeduction.SetFocus: Exit Sub

If Trim(TxtHonorarium) = "" Then MsgBox "Sorry cannot continue!! Honorarium is empty", vbCritical + vbOKOnly: Exit Sub
If Trim(TxtLoan) = "" Then MsgBox "Sorry cannot continue!! SSS Loan is empty", vbCritical + vbOKOnly: Exit Sub
If Trim(TxtODeduction) = "" Then MsgBox "Sorry cannot continue!! Other Deduction is empty", vbCritical + vbOKOnly: Exit Sub
TxtHonorarium.Text = Val(TxtHonorarium.Text) / 2

'Dim TempRs As New ADODB.Recordset
Me.MousePointer = vbHourglass

If addflag = True Then
    'TempRs.Open "SELECT * FROM  computation where Ecode ='" & TxtEcode & "'", CN, adOpenForwardOnly, adLockReadOnly
    'If Not TempRs.EOF Then 'not empty
    'MsgBox TxtEcode & " Employee code already exist! Duplication is not allowed", vbCritical
    
    'TxtEcode.SetFocus
    'Exit Sub
    'End If
    'Set TempRs = Nothing
    
RsPayroll.AddNew 'blank record will be created

End If

With RsPayroll

    !payroll_period = DTPicker1
    !ECODE = cmbecode.Text
    !fullname = Txtename.Text
    !basic_pay = Txtmonthly.Text
    !com_basic = TxtBasicPay
    !semi_monthly = Text3.Text
    !living_allowance = TxtLivingAllow
    !rice_allowance = TxtRiceAllow
    !honorarium = TxtHonorarium
    !Phil_Health = TxtPhilHealth
    !SSS_premium = TxtSSS
    !SSS_loan = TxtLoan
    !Pagibig = TxtPagibig
    !withholding_tax = TxtTax
    !other_deduction = TxtODeduction
    
    .Update
    .Requery
    
End With

MsgBox "Record was saved successfully!!", vbInformation + vbOKOnly
Me.MousePointer = vbDefault
ToggleButtonUpdateCancel "OFF"

End Sub

Private Sub TxtHonorarium_Change()
If IsNumeric(TxtHonorarium) = False Then MsgBox "Invalid input. Please check the value!", vbCritical, "ERROR": TxtHonorarium.SetFocus: Exit Sub
End Sub
