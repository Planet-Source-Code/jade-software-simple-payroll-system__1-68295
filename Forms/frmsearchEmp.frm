VERSION 5.00
Begin VB.Form frmsearchEmp 
   BackColor       =   &H000000C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search..."
   ClientHeight    =   3000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4440
   Icon            =   "frmsearchEmp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   4440
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Text            =   "Enter some text here!"
      Top             =   1215
      Width           =   4215
   End
   Begin VB.CommandButton cmdcancel 
      Caption         =   "Cancel"
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
      Left            =   3240
      TabIndex        =   2
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search"
      Default         =   -1  'True
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
      Left            =   2040
      TabIndex        =   1
      Top             =   2520
      Width           =   1095
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmsearchEmp.frx":08CA
      Left            =   120
      List            =   "frmsearchEmp.frx":08D4
      TabIndex        =   0
      Top             =   1935
      Width           =   4215
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000014&
      X1              =   120
      X2              =   4320
      Y1              =   2415
      Y2              =   2415
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Search for:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   975
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmsearchEmp.frx":08F7
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   720
      TabIndex        =   5
      Top             =   165
      Width           =   3495
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      X1              =   120
      X2              =   4320
      Y1              =   855
      Y2              =   855
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      BorderWidth     =   2
      X1              =   120
      X2              =   4320
      Y1              =   855
      Y2              =   855
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmsearchEmp.frx":0981
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Look in:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1695
      Width           =   1815
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000010&
      BorderWidth     =   2
      X1              =   120
      X2              =   4320
      Y1              =   2415
      Y2              =   2415
   End
End
Attribute VB_Name = "frmsearchEmp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdcancel_Click()
Unload Me
End Sub

Private Sub cmdSearch_Click()

If Combo1.Text = "Employee No" Then

    Dim x As String
    x = Text1.Text
    
    Set adoEmpSearch = New ADODB.Recordset
    adoEmpSearch.Open "SELECT * FROM employeefile where ecode='" & x & "'", CN, adOpenStatic, adLockOptimistic
    
   
    
    If adoEmpSearch.RecordCount >= 1 Then
   
    Me.MousePointer = vbHourglass
    Display_EmpRec
    Me.MousePointer = vbDefault
     
    Else
            Me.MousePointer = vbHourglass
            MsgBox "No record found on database!", vbCritical
            Me.MousePointer = vbDefault
            Exit Sub
            
    End If
        
      Set adoEmpSearch = Nothing
    
   
       
ElseIf Combo1.Text = "Employee Surname" Then

    Dim xx As String
    xx = Text1.Text
    
    Set adoEmpSearch = New ADODB.Recordset
    adoEmpSearch.Open "SELECT * FROM employeefile where sname='" & xx & "'", CN, adOpenStatic, adLockOptimistic
    

      
    If adoEmpSearch.RecordCount >= 1 Then
    
    Me.MousePointer = vbHourglass
    Display_EmpRec
    Me.MousePointer = vbDefault
    
    Else
            
            Me.MousePointer = vbHourglass
            MsgBox "No record found on database!", vbCritical
            Me.MousePointer = vbDefault
            Exit Sub
            
    End If
    
      Set adoEmpSearch = Nothing


Else

    Me.MousePointer = vbHourglass
    MsgBox "Please Select category to search !", vbInformation, "Info."
    Me.MousePointer = vbDefault
    Exit Sub
    
End If

   Unload Me
   
End Sub

Private Sub Text1_Click()
Text1.Text = ""
End Sub

Sub Display_EmpRec()

On Error Resume Next

With adoEmpSearch

    frmEmployeeFile.TxtEcode.Text = !Ecode & " " '-// To avoid Null value
    frmEmployeeFile.ttfname.Text = !fname & " "
    frmEmployeeFile.TxtSname.Text = !sname & " "
    frmEmployeeFile.txtmi.Text = !mi & " "
    frmEmployeeFile.TxtEage.Text = !eage & " "
   
    '-// retrieve info. (gender)

    If !gender = "m" Then
      frmEmployeeFile.OptionMale.Value = True
       
    Else
      frmEmployeeFile.OptionFemale.Value = True
    End If
    
    frmEmployeeFile.CmbCivilStat = !civil_status & " "
    frmEmployeeFile.DTbirthdate.Value = !birthdate
    frmEmployeeFile.TxtEaddress.Text = !eaddress & " "
    frmEmployeeFile.TxtPhoneNo.Text = !phone_no & " "
    frmEmployeeFile.TxtEadd.Text = !email_add & " "
    frmEmployeeFile.TxtSemi.Text = !semi_monthly & " "
    frmEmployeeFile.TxtBasicPay.Text = !basicpay & " "
    frmEmployeeFile.CmbPosition.Text = !posdesc & " "
    frmEmployeeFile.TxtSyears.Text = !syears & " "
    frmEmployeeFile.CmbDepartment.Text = !depdesc & " "
    frmEmployeeFile.CmbTaxStat.Text = !taxheadercode & " "
    frmEmployeeFile.cmbrice.Text = !rice_all_code & " "
    frmEmployeeFile.cmbliving.Text = !living_all_code & " "
    frmEmployeeFile.Text5.Text = !rank & " "
    frmEmployeeFile.Text3.Text = !living_value & " "
    frmEmployeeFile.Text4.Text = !rice_value & " "
    frmEmployeeFile.TxtSSS.Text = !SSSPremium & " "
    frmEmployeeFile.TxtPhilHealth = !PHHealthValue & " "
    frmEmployeeFile.Text2.Text = !Pagibig & " "
    frmEmployeeFile.Text1.Text = !withTaxvalue & " "
    
End With

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then cmdSearch_Click

End Sub
