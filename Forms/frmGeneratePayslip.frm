VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPayslip 
   BackColor       =   &H000000C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Generate Payslip"
   ClientHeight    =   2085
   ClientLeft      =   30
   ClientTop       =   420
   ClientWidth     =   3960
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2085
   ScaleWidth      =   3960
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1080
      Width           =   2172
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Print"
      Height          =   325
      Left            =   2760
      TabIndex        =   3
      Top             =   1680
      Width           =   1092
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print  &All"
      Height          =   325
      Left            =   1560
      TabIndex        =   0
      Top             =   1680
      Width           =   1092
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   288
      Left            =   240
      TabIndex        =   6
      Top             =   600
      Width           =   1692
      _ExtentX        =   2990
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
      Format          =   24379393
      CurrentDate     =   38695
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   288
      Left            =   2040
      TabIndex        =   7
      Top             =   600
      Width           =   1692
      _ExtentX        =   2990
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
      Format          =   24379393
      CurrentDate     =   38695
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Generate PaySlip"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   240
      TabIndex        =   4
      Top             =   240
      Width           =   3492
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Employee Name"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   1332
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   3
      Height          =   1452
      Left            =   120
      Top             =   120
      Width           =   3732
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0E0FF&
      Height          =   1452
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3732
   End
End
Attribute VB_Name = "frmPayslip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdPrint_Click()

Set Rs = New ADODB.Recordset
Rs.Open "SELECT *, [gross_pay]-[total_deduction] As net_pay FROM [qryrpt] WHERE payroll_period Between #" & DTPicker1.Value & "# AND #" & DTPicker2.Value & "# ORDER BY ecode", CN, adOpenStatic, adLockOptimistic
 
If Rs.RecordCount >= 1 Then
 
Set RptPayslipALl.DataSource = Rs
RptPayslipALl.Sections("Section1").Controls.Item("label3").Caption = "Payroll Period : FROM : " & Format(DTPicker1.Value, "MM-DD-YY") & " TO : " & Format(DTPicker2.Value, "MM-DD-YY")
RptPayslipALl.Sections("Section1").Controls.Item("Line1").Visible = True
RptPayslipALl.Show
RptPayslipALl.SetFocus

Else

    MsgBox "No Employee record found with this Payroll Period!", vbCritical, "ERROR"
    Exit Sub


End If

Set Rs = Nothing

End Sub

Private Sub Command1_Click()

Set rsRec = New ADODB.Recordset
rsRec.Open "SELECT *, [gross_pay]-[total_deduction] As net_pay FROM [qryrpt] WHERE fullname='" & Combo1.Text & "' and payroll_period Between #" & DTPicker1.Value & "# AND #" & DTPicker2.Value & "# ORDER BY ecode", CN, adOpenStatic, adLockOptimistic
 

If Combo1.Text = "" Then MsgBox "Please Select Employee Name!", vbCritical, "ERROR": Exit Sub
    
If rsRec.RecordCount >= 1 Then

 Set RptPayslip.DataSource = rsRec
 RptPayslip.Sections("Section4").Controls.Item("label3").Caption = "Payroll Period : FROM : " & Format(DTPicker1.Value, "MM-DD-YY") & " TO : " & Format(DTPicker2.Value, "MM-DD-YY")

 RptPayslip.Show
  
Else
    
    MsgBox "No Employee record found with this Payroll Period!", vbCritical, "ERROR"
    Exit Sub

End If

Set rsRec = Nothing

End Sub

Private Sub Form_Load()

Set Rs = New ADODB.Recordset
Rs.Open "SELECT * FROM QRyrpt", CN, adOpenStatic, adLockOptimistic
Do Until Rs.EOF
Combo1.AddItem Rs!fullname
Rs.MoveNext
Loop

End Sub

Private Sub Form_Unload(Cancel As Integer)

Set Rs = Nothing

End Sub
