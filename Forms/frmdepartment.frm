VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmdepartment 
   BackColor       =   &H000000C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Department"
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4290
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   4290
   Begin VB.TextBox TxtDept 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1680
      TabIndex        =   11
      Top             =   720
      Width           =   2385
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "&Cancel"
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
      TabIndex        =   3
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "&Save"
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
      Left            =   2400
      TabIndex        =   4
      Top             =   1800
      Width           =   855
   End
   Begin VB.CommandButton CmdDelete 
      Caption         =   "&Delete"
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
      Left            =   1680
      TabIndex        =   5
      Top             =   1800
      Width           =   735
   End
   Begin VB.CommandButton CmdEdit 
      Caption         =   "&Edit"
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
      Left            =   960
      TabIndex        =   6
      Top             =   1800
      Width           =   735
   End
   Begin VB.CommandButton CmdLast 
      Caption         =   ">>"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   265
      Left            =   3000
      TabIndex        =   1
      ToolTipText     =   "Last"
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton CmdNext 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   265
      Left            =   2040
      TabIndex        =   2
      ToolTipText     =   "Next"
      Top             =   1440
      Width           =   975
   End
   Begin VB.CommandButton CmdPrevious 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   265
      Left            =   1080
      TabIndex        =   7
      ToolTipText     =   "Previous"
      Top             =   1440
      Width           =   975
   End
   Begin VB.TextBox TxtDeptCode 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1680
      MaxLength       =   2
      TabIndex        =   12
      Top             =   360
      Width           =   1335
   End
   Begin VB.CommandButton CmdFirst 
      Caption         =   "<<"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   265
      Left            =   120
      TabIndex        =   9
      ToolTipText     =   "First"
      Top             =   1440
      Width           =   975
   End
   Begin VB.CommandButton CmdAdd 
      Caption         =   "&Add"
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
      Left            =   120
      TabIndex        =   8
      Top             =   1800
      Width           =   855
   End
   Begin VB.Frame Frame2 
      Height          =   15
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   4095
      Begin MSComctlLib.ImageList Img 
         Left            =   0
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmdepartment.frx":0000
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Label Label1 
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Department Code"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Department"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   720
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   3
      Height          =   1095
      Left            =   120
      Top             =   120
      Width           =   4095
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0E0FF&
      Height          =   1095
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   4095
   End
End
Attribute VB_Name = "frmdepartment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim addflag As Boolean

Sub Display_record()
On Error Resume Next
    TxtDeptCode = rsDepart.Fields("depcode")
    TxtDept = rsDepart.Fields("depdesc")
End Sub
Sub lock_txt()

TxtDeptCode.Locked = True
TxtDept.Locked = True
End Sub

Sub unlock_txt()
TxtDeptCode.Locked = False
TxtDept.Locked = False
End Sub

Private Sub CmdAdd_Click()

addflag = True
ToggleButtonSaveCancel "ON"
unlock_txt
TxtDeptCode = ""
TxtDept = ""
TxtDeptCode.SetFocus
End Sub

Sub cleartext()
TxtDeptCode = ""
TxtDept = ""
End Sub

Private Sub CmdCancel_Click()

    ToggleButtonSaveCancel "OFF"
    If rsDepart.BOF And rsDepart.EOF Then disable_nav_button: Exit Sub
    Display_record
    
End Sub

Private Sub ToggleButtonSaveCancel(OnorOff As String)
 If UCase(OnorOff) = "ON" Then 'Click Add or Edit
    cmdAdd.Enabled = False
    CmdEdit.Enabled = False
    
    CmdSave.Enabled = True
    CmdCancel.Enabled = True
    
    CmdDelete.Enabled = False
  
    
    CmdFirst.Enabled = False
    CmdPrevious.Enabled = False
    CmdNext.Enabled = False
    CmdLast.Enabled = False
    
    TxtDeptCode.Locked = False
    TxtDept.Locked = False
    
Else
    cmdAdd.Enabled = True
    CmdEdit.Enabled = True
    
    CmdSave.Enabled = False
    CmdCancel.Enabled = False
    
    CmdDelete.Enabled = True
  
    
    CmdFirst.Enabled = True
    CmdPrevious.Enabled = True
    CmdNext.Enabled = True
    CmdLast.Enabled = True
    
    TxtDeptCode.Locked = True
    TxtDept.Locked = True
   
End If
End Sub


Private Sub CmdDelete_Click()
On Error Resume Next

   If Trim(TxtDept) = "" Then MsgBox "No record found to be deleted!", vbInformation: disable_nav_button: Exit Sub
 
    Set rsRec = New ADODB.Recordset
        rsRec.Open "SELECT * FROM [employeefile] where depdesc='" & TxtDept & "'", CN, adOpenKeyset, adLockOptimistic

       
    If rsRec.RecordCount >= 1 Then MsgBox "Access denied! Cannot delete this record." & vbCrLf & "There is a record bound with this position." & vbCrLf & "Please delete first Employee Record before continuing...", vbCritical: Exit Sub
     
    
    
If rsDepart.BOF And rsDepart.EOF = True Then
      
            MsgBox "Empty Database", vbInformation
            Exit Sub
            
 Else
 
        
 Dim msg As Integer

    msg = MsgBox("This will delete " & TxtDept & ". Proceed? ", vbQuestion + vbYesNo, "Confirm Deleting...")
    
    If msg = vbYes Then
        
     
        Me.MousePointer = vbHourglass
         rsDepart.Delete
         rsDepart.Requery
        Display_record
         rsDepart.MoveFirst
        If rsDepart.EOF Then rsDepart.Requery: cleartext: disable_nav_button
        MsgBox "Record deleted successfully...", vbInformation
        Me.MousePointer = vbDefault
    End If
 End If
 

End Sub

Private Sub CmdEdit_Click()

If rsDepart.BOF And rsDepart.EOF Then MsgBox "No record(s) to edit", vbCritical, "ERROR":  Exit Sub
ToggleButtonSaveCancel "ON"
addflag = False
unlock_txt
TxtDeptCode.SetFocus
End Sub

Private Sub CmdSave_Click()

If IsNumeric(TxtDeptCode) = True Then MsgBox "Invalid input. Please check it!", vbCritical, "ERROR": TxtDeptCode.SetFocus: Exit Sub
If IsNumeric(TxtDept) = True Then MsgBox "Invalid input. Please check it!", vbCritical, "ERROR": TxtDept.SetFocus: Exit Sub

'-//check if empty
If Trim(TxtDeptCode) = "" Then MsgBox "Sorry cannot continue!! Department Code is empty", vbCritical + vbOKOnly: Exit Sub
If Trim(TxtDept) = "" Then MsgBox "Sorry cannot continue!! Department is empty", vbCritical + vbOKOnly: Exit Sub

Dim TempRs As New ADODB.Recordset

Me.MousePointer = vbHourglass

If addflag = True Then

    TempRs.Open "SELECT * from [department] where DepCode ='" & TxtDeptCode & "'", CN, adOpenForwardOnly, adLockReadOnly
    
    If Not TempRs.EOF Then 'not empty
        
        MsgBox TxtDeptCode & vbCrLf & vbCrLf & " already exist! Duplication of record is not allowed.", vbCritical
        TxtDeptCode.SetFocus
        Me.MousePointer = vbDefault
        Exit Sub
              
    End If
    
      

 
    Set TempRs = Nothing

     rsDepart.AddNew 'blank record will be created

End If

rsDepart.Fields("depcode") = TxtDeptCode
rsDepart.Fields("depdesc") = TxtDept
rsDepart.Update
rsDepart.Requery '-// Refresh

MsgBox "Record was saved successfully!!", vbInformation + vbOKOnly
Me.MousePointer = vbDefault
Call lock_txt
ToggleButtonSaveCancel "OFF"


End Sub

Private Sub Form_Load()
Me.Icon = Img.ListImages(1).ExtractIcon
Me.Show
Call lock_txt
Set rsDepart = New ADODB.Recordset
rsDepart.Open "SELECT * from department", CN, adOpenKeyset, adLockOptimistic
If rsDepart.BOF And rsDepart.EOF Then MsgBox "No record(s) to display!", vbCritical, "ERROR": disable_nav_button: Exit Sub
ToggleButtonSaveCancel "OFF"
Display_record
  
End Sub

Sub disable_nav_button()

CmdFirst.Enabled = False
CmdNext.Enabled = False
CmdLast.Enabled = False
CmdPrevious.Enabled = False

End Sub

Private Sub Form_Unload(Cancel As Integer)
Set rsDepart = Nothing
End Sub

Private Sub CmdFirst_Click()
On Error Resume Next
     rsDepart.MoveFirst
      Display_record
    If rsDepart.BOF Then Exit Sub
    
End Sub

Private Sub CmdLast_Click()
On Error Resume Next
     rsDepart.MoveLast
      Display_record
    If rsDepart.EOF Then Exit Sub
       
End Sub

Private Sub CmdNext_Click()
On Error Resume Next
  rsDepart.MoveNext
    Display_record
        If rsDepart.EOF Then
             MsgBox "You have reached the End of File!", vbCritical + vbOKOnly
              rsDepart.MoveLast
        End If
End Sub

Private Sub CmdPrevious_Click()
    
On Error Resume Next
 rsDepart.MovePrevious
    Display_record

    If rsDepart.BOF Then
             MsgBox "You have reached the Beginning of File!", vbCritical + vbOKOnly
            Display_record
        End If
End Sub


Private Sub TxtDept_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then CmdSave_Click
End Sub

Private Sub TxtDeptCode_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then TxtDept.SetFocus
End Sub
