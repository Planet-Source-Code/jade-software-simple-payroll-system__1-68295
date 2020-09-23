VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAttendance 
   BackColor       =   &H000000C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Attendance"
   ClientHeight    =   4200
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   5388
   Icon            =   "frmAttendance.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   5388
   Begin VB.ComboBox cmbEcode 
      Height          =   315
      Left            =   1680
      TabIndex        =   41
      Top             =   680
      Width           =   1935
   End
   Begin VB.TextBox txttemp 
      Height          =   285
      Left            =   1080
      TabIndex        =   39
      Top             =   9720
      Width           =   1335
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
      Left            =   1680
      TabIndex        =   27
      Top             =   1080
      Width           =   3375
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
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   26
      Top             =   2520
      Width           =   1095
   End
   Begin VB.TextBox TxtWorkingDays 
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
      Left            =   1680
      TabIndex        =   25
      Top             =   1440
      Width           =   1095
   End
   Begin VB.TextBox TxtDaysWorked 
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
      Height          =   270
      Left            =   1680
      TabIndex        =   24
      Top             =   1800
      Width           =   1095
   End
   Begin VB.TextBox TxtLates 
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
      Height          =   300
      Left            =   1680
      TabIndex        =   23
      Top             =   2160
      Width           =   1095
   End
   Begin VB.TextBox txthalfday 
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
      Left            =   3600
      TabIndex        =   22
      Top             =   1440
      Width           =   1455
   End
   Begin VB.TextBox txtabsent 
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
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   21
      Top             =   1800
      Width           =   1455
   End
   Begin VB.CommandButton cmdcompute 
      Caption         =   "Com&pute"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   310
      Left            =   3600
      TabIndex        =   20
      Top             =   2160
      Width           =   1455
   End
   Begin VB.TextBox txttempabsent 
      Height          =   285
      Left            =   2760
      TabIndex        =   14
      Text            =   "txttempabsent"
      Top             =   10320
      Width           =   1215
   End
   Begin VB.TextBox txttemptdeduc 
      Height          =   285
      Left            =   2640
      TabIndex        =   13
      Text            =   "txttemptdeduc"
      Top             =   9690
      Width           =   1215
   End
   Begin VB.TextBox txttemphalfday 
      Height          =   285
      Left            =   1080
      TabIndex        =   12
      Text            =   "temphalfday"
      Top             =   10920
      Width           =   1335
   End
   Begin VB.TextBox txttemplate 
      Height          =   285
      Left            =   2520
      TabIndex        =   11
      Text            =   "txtlate"
      Top             =   10920
      Width           =   1215
   End
   Begin VB.TextBox txtdaily 
      Height          =   285
      Left            =   1080
      TabIndex        =   10
      Text            =   "txttempdaily"
      Top             =   10320
      Width           =   1215
   End
   Begin VB.CommandButton cmdClose 
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
      Left            =   4440
      TabIndex        =   9
      Top             =   3720
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
      Left            =   3480
      TabIndex        =   2
      Top             =   3720
      Width           =   975
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "&Save"
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
      TabIndex        =   3
      Top             =   3720
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
      TabIndex        =   4
      Top             =   3720
      Width           =   855
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
      Left            =   4080
      TabIndex        =   0
      ToolTipText     =   "Last"
      Top             =   3360
      Width           =   1215
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
      Left            =   2760
      TabIndex        =   1
      ToolTipText     =   "Next"
      Top             =   3360
      Width           =   1335
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
      Left            =   1440
      TabIndex        =   5
      ToolTipText     =   "Previous"
      Top             =   3360
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Height          =   30
      HelpContextID   =   10
      Left            =   120
      TabIndex        =   8
      Top             =   3240
      Width           =   5175
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
      TabIndex        =   7
      ToolTipText     =   "First"
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CommandButton CmdAdd 
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
      TabIndex        =   6
      Top             =   3720
      Width           =   855
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   330
      Left            =   1680
      TabIndex        =   28
      Top             =   240
      Width           =   1935
      _ExtentX        =   3408
      _ExtentY        =   593
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
      Format          =   43646977
      CurrentDate     =   38686
   End
   Begin VB.Label Label13 
      Caption         =   "Semi-Temp"
      Height          =   255
      Left            =   1080
      TabIndex        =   40
      Top             =   9480
      Width           =   1335
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   3
      Height          =   3015
      Left            =   120
      Top             =   120
      Width           =   5175
   End
   Begin VB.Label Label1 
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Code"
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
      Height          =   255
      Left            =   240
      TabIndex        =   37
      Top             =   675
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Name"
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
      Height          =   255
      Left            =   240
      TabIndex        =   36
      Top             =   1020
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Payroll Period"
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
      Height          =   255
      Left            =   240
      TabIndex        =   35
      Top             =   285
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Required Working Days"
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
      Height          =   495
      Left            =   240
      TabIndex        =   34
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label Label5 
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Days Worked"
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
      Height          =   375
      Left            =   240
      TabIndex        =   33
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label Label6 
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "No. of Lates"
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
      Height          =   375
      Left            =   240
      TabIndex        =   32
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label Label7 
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Half Day"
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
      Height          =   255
      Left            =   2880
      TabIndex        =   31
      Top             =   1470
      Width           =   855
   End
   Begin VB.Label Label8 
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Semi-monthly rate"
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
      Height          =   375
      Left            =   240
      TabIndex        =   30
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Label Label9 
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Absent"
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
      Height          =   255
      Left            =   2880
      TabIndex        =   29
      Top             =   1830
      Width           =   855
   End
   Begin VB.Label Label15 
      Caption         =   "txtlate"
      Height          =   255
      Left            =   2520
      TabIndex        =   19
      Top             =   10680
      Width           =   975
   End
   Begin VB.Label Label14 
      Caption         =   "tempdaily"
      Height          =   255
      Left            =   1080
      TabIndex        =   18
      Top             =   10080
      Width           =   975
   End
   Begin VB.Label Label12 
      Caption         =   "absent"
      Height          =   255
      Left            =   2760
      TabIndex        =   17
      Top             =   10080
      Width           =   975
   End
   Begin VB.Label Label11 
      Caption         =   "tempdeduc"
      Height          =   255
      Left            =   2640
      TabIndex        =   16
      Top             =   9480
      Width           =   975
   End
   Begin VB.Label Label10 
      Caption         =   "halfday"
      Height          =   255
      Left            =   1080
      TabIndex        =   15
      Top             =   10680
      Width           =   975
   End
   Begin VB.Label Label16 
      BackColor       =   &H00C0E0FF&
      Height          =   3015
      Left            =   120
      TabIndex        =   38
      Top             =   120
      Width           =   5175
   End
End
Attribute VB_Name = "frmAttendance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim late, halfday, tempbas, absent, tmpabsent
Dim tempdeduc, basepay
Dim addflag As Boolean
Const lates = 50#

Sub Displayfields()

   On Error Resume Next
    
    cmbecode = adoAttend.Fields("ecode") & " "
    Txtename.Text = adoAttend.Fields("fullname") & " "
    TxtWorkingDays = adoAttend.Fields("workingdays") & " "
    TxtDaysWorked = adoAttend.Fields("daysworked") & " "
    TxtLates = adoAttend.Fields("lates") & " "
    TxtBasicPay = adoAttend.Fields("basic_pay") & " "
    DTPicker1 = adoAttend.Fields("payroll_period") & " "
    txtabsent.Text = adoAttend.Fields("absent") & " "
    txthalfday.Text = adoAttend.Fields("halfday") & " "

End Sub


Sub disable_nav_button()

CmdFirst.Enabled = False
CmdNext.Enabled = False
CmdLast.Enabled = False
CmdPrevious.Enabled = False

End Sub

Private Sub ToggleButtonSaveCancel(OnorOff As String)
 If UCase(OnorOff) = "ON" Then 'Click Add or Edit
      
   CmdSave.Enabled = True
    CmdCancel.Enabled = True
    
    CmdEdit.Enabled = False
    CmdClose.Enabled = False
    
    CmdFirst.Enabled = False
    CmdPrevious.Enabled = False
    CmdNext.Enabled = False
    CmdLast.Enabled = False
    
   
    TxtWorkingDays.Locked = False
    TxtDaysWorked.Locked = False
    TxtLates.Locked = False
    TxtBasicPay.Locked = False
    
Else

  
    CmdSave.Enabled = False
    CmdCancel.Enabled = False
       
    CmdEdit.Enabled = True
    CmdClose.Enabled = True
    
    CmdFirst.Enabled = True
    CmdPrevious.Enabled = True
    CmdNext.Enabled = True
    CmdLast.Enabled = True
     
  
    TxtWorkingDays.Locked = True
    TxtDaysWorked.Locked = True
    TxtLates.Locked = True
    TxtBasicPay.Locked = True
          
End If
End Sub


Private Sub cmbecode_Click()

Set rsRec = New ADODB.Recordset
rsRec.Open "SELECT * FROM employeefile where ecode='" & cmbecode.Text & "'", CN, adOpenStatic, adLockPessimistic

If rsRec.RecordCount >= 1 Then

    Txtename.Text = rsRec!sname & ", " & rsRec!fname & " " & rsRec!mi
    txttemp.Text = rsRec!semi_monthly & " "
    
End If

Set rsRec = Nothing

End Sub

Private Sub CmdAdd_Click()

addflag = True
ToggleButtonSaveCancel "ON"

cmbecode.Text = ""
cmbecode.Locked = False
Txtename.Text = ""
TxtWorkingDays.Text = ""
TxtDaysWorked.Text = ""
TxtLates.Text = ""
TxtBasicPay.Text = ""
txthalfday.Text = ""
txtabsent.Text = ""
TxtBasicPay.Locked = True
TxtWorkingDays.SetFocus

End Sub

Private Sub cmdcancel_Click()
  
ToggleButtonSaveCancel "OFF"

If adoAttend.BOF And adoAttend.EOF Then disable_nav_button: Exit Sub
Displayfields
 
End Sub

Private Sub CmdClose_Click()
Unload Me
End Sub

Private Sub cmdCompute_Click()

If IsNumeric(TxtWorkingDays) = False Then MsgBox "Invalid input. Please check value!", vbCritical + vbOKOnly: Exit Sub
If IsNumeric(TxtDaysWorked) = False Then MsgBox "Invalid input. Please check value!", vbCritical + vbOKOnly: Exit Sub
If IsNumeric(TxtLates) = False Then MsgBox "Invalid input. Please check value!", vbCritical + vbOKOnly: Exit Sub
If IsNumeric(txthalfday) = False Then MsgBox "Invalid input. Please check value!", vbCritical + vbOKOnly: Exit Sub

If Val(TxtWorkingDays.Text) <= 0 Then MsgBox "Invalid input. Please check value!", vbCritical + vbOKOnly: Exit Sub
If Val(TxtDaysWorked.Text) <= 0 Then MsgBox "Invalid input. Please check value!", vbCritical + vbOKOnly: Exit Sub
If Val(TxtLates.Text) < 0 Then MsgBox "Invalid input. Please check value!", vbCritical + vbOKOnly: Exit Sub
If Val(txthalfday.Text) < 0 Then MsgBox "Invalid input. Please check value!", vbCritical + vbOKOnly: Exit Sub
If TxtWorkingDays.Text < TxtDaysWorked.Text Then MsgBox "Invalid input. Please check value!", vbCritical + vbOKOnly: Exit Sub
If TxtWorkingDays = "" Or TxtDaysWorked = "" Or TxtLates = "" Or txthalfday = "" Or txtabsent = "" Then MsgBox "Missing Fields!, Please check it!", vbCritical: Exit Sub

'-// Computation for Daily Rate
tempbas = Val(txttemp.Text) / Val(TxtWorkingDays.Text)
txtdaily.Text = tempbas

'-// COmputation for halfday
halfday = tempbas / 2 * Val(txthalfday.Text)
txttemphalfday.Text = halfday

'-// Computation for lates
late = Val(TxtLates.Text) * lates
txttemplate.Text = late

'-// Computation for absent
tmpabsent = Val(txtabsent.Text) * Val(txtdaily.Text)
txttempabsent = tmpabsent

'-// COmputation for Deduction
tempdeduc = halfday + tmpabsent + late
txttemptdeduc.Text = tempdeduc

'-// Computation for Semi-monthly rate
basepay = Val(txttemp.Text) - Val(txttemptdeduc.Text)
TxtBasicPay.Text = Format(basepay, "##,###.#0")

End Sub



Private Sub CmdSave_Click()

'Dim TempRs As New ADODB.Recordset

If Trim(TxtWorkingDays) = "" Then MsgBox "Sorry cannot continue!! Number of Working Days is empty", vbCritical + vbOKOnly: Exit Sub
If Trim(TxtDaysWorked) = "" Then MsgBox "Sorry cannot continue!! Numberof Days Worked is empty", vbCritical + vbOKOnly:  Exit Sub
If Trim(TxtLates) = "" Then MsgBox "Sorry cannot continue!! Number of Lates is empty", vbCritical + vbOKOnly: Exit Sub
If Trim(TxtBasicPay) = "" Then MsgBox "Sorry cannot continue!! Basic Pay is empty", vbCritical + vbOKOnly: Exit Sub

Me.MousePointer = vbHourglass

If addflag = True Then
    
    'TempRs.Open "SELECT * from attendance where Ecode ='" & TxtEcode & "'", CN, adOpenForwardOnly, adLockReadOnly
    'If Not TempRs.EOF Then 'not empty
    'MsgBox TxtEcode & " employee code already exist! Duplication is not allowed", vbCritical
    'TxtEcode.SetFocus
    'Exit Sub
     'End If

    'Set TempRs = Nothing
    adoAttend.AddNew 'blank record will be created
End If

   
        adoAttend.Fields("ecode") = cmbecode
        adoAttend.Fields("fullname") = Txtename.Text
        adoAttend.Fields("workingdays") = Val(TxtWorkingDays.Text)
        adoAttend.Fields("daysworked") = Val(TxtDaysWorked.Text)
        adoAttend.Fields("lates") = Val(TxtLates.Text)
        adoAttend.Fields("basic_pay") = Format(TxtBasicPay.Text, "##,###,###.#0")
        adoAttend.Fields("payroll_period") = DTPicker1
        adoAttend.Fields("halfday") = Val(txthalfday.Text)
        adoAttend.Fields("absent") = Val(txtabsent)

        adoAttend.Update
        adoAttend.Requery
 
MsgBox "Record was saved successfully!!", vbInformation + vbOKOnly
Me.MousePointer = vbDefault
ToggleButtonSaveCancel "OFF"

End Sub

Private Sub Form_Load()

Me.Show
'cmbecode.Locked = True
combo_rec
Set adoAttend = New ADODB.Recordset
adoAttend.Open "SELECT * from attendance", CN, adOpenKeyset, adLockOptimistic
If adoAttend.BOF And adoAttend.EOF Then MsgBox "No record(s) to display!", vbCritical, "ERROR":  disable_nav_button: Exit Sub
ToggleButtonSaveCancel "OFF"
Displayfields
DTPicker1 = Date
  
End Sub

Sub combo_rec()

Set Adors = New ADODB.Recordset
Adors.Open "SELECT * FROM employeefile", CN, adOpenStatic, adLockOptimistic
Do Until Adors.EOF
cmbecode.AddItem Adors!Ecode
Adors.MoveNext
Loop

Set Adors = Nothing

End Sub

Private Sub Form_Unload(Cancel As Integer)
Set adoAttend = Nothing
End Sub

Private Sub TxtWorkingDays_Change()

absent = Val(TxtWorkingDays.Text) - Val(TxtDaysWorked)
txtabsent.Text = absent

End Sub

Private Sub TxtDaysWorked_Change()

absent = Val(TxtWorkingDays.Text) - Val(TxtDaysWorked)
txtabsent.Text = absent

End Sub

Private Sub CmdFirst_Click()
On Error Resume Next
    adoAttend.MoveFirst
    Displayfields
    If adoAttend.BOF Then Exit Sub
    
End Sub

Private Sub CmdLast_Click()
On Error Resume Next
    adoAttend.MoveLast
    Displayfields
    If adoAttend.EOF Then Exit Sub
       
End Sub

Private Sub CmdNext_Click()
On Error Resume Next
 adoAttend.MoveNext
 Displayfields
        If adoAttend.EOF Then
             MsgBox "Last record has been reached!", vbCritical + vbOKOnly
             adoAttend.MoveLast
        End If
End Sub

Private Sub CmdPrevious_Click()
 On Error Resume Next
    adoAttend.MovePrevious
   Displayfields

    If adoAttend.BOF Then
             MsgBox "First record has been reached!", vbCritical + vbOKOnly
             Rs.MoveFirst
        End If
End Sub

Private Sub CmdEdit_Click()

cmbecode.Locked = False
If adoAttend.RecordCount = 0 Then MsgBox "No record(s) to edit", vbCritical, "ERROR": Exit Sub
ToggleButtonSaveCancel "ON"
addflag = False
TxtWorkingDays.SetFocus
End Sub
