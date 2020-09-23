VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmposition 
   BackColor       =   &H000000C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Position"
   ClientHeight    =   2280
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4290
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2280
   ScaleWidth      =   4290
   Begin VB.TextBox TxtPosCode 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Left            =   1320
      MaxLength       =   2
      TabIndex        =   12
      Top             =   360
      Width           =   1095
   End
   Begin VB.TextBox TxtPosition 
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
      Left            =   1320
      TabIndex        =   11
      Top             =   720
      Width           =   2775
   End
   Begin VB.Frame Frame2 
      Height          =   15
      Left            =   120
      TabIndex        =   9
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
               Picture         =   "frmposition.frx":0000
               Key             =   ""
            EndProperty
         EndProperty
      End
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
      TabIndex        =   0
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
      TabIndex        =   1
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
      TabIndex        =   2
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
      TabIndex        =   3
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
      TabIndex        =   4
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
      TabIndex        =   5
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
      TabIndex        =   6
      ToolTipText     =   "Previous"
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
      TabIndex        =   7
      ToolTipText     =   "First"
      Top             =   1440
      Width           =   975
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3840
      Top             =   2040
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
            Picture         =   "frmposition.frx":1994
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Position Code"
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
      Top             =   390
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Position"
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
      Top             =   750
      Width           =   975
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
Attribute VB_Name = "frmposition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim addflag As Boolean

Sub disable_nav_button()

CmdFirst.Enabled = False
CmdNext.Enabled = False
CmdLast.Enabled = False
CmdPrevious.Enabled = False

End Sub

Private Sub CmdAdd_Click()
addflag = True
ToggleButtonSaveCancel "ON"
Call unlock_txt
clear_ALL_ctrl
TxtPosCode.SetFocus
End Sub

Sub clear_ALL_ctrl()
TxtPosCode = ""
TxtPosition = ""

End Sub

Private Sub CmdCancel_Click()

    ToggleButtonSaveCancel "OFF"
    If rsPosition.BOF And rsPosition.EOF Then disable_nav_button: Exit Sub
    display_rec
    
End Sub

Private Sub CmdDelete_Click()
'-// Relations with Employee File

On Error Resume Next

   If Trim(TxtPosition) = "" Then MsgBox "No record found to be deleted!", vbInformation: disable_nav_button: Exit Sub
 
    Set rsRec = New ADODB.Recordset
        rsRec.Open "SELECT * FROM [employeefile] where posdesc='" & TxtPosition & "'", CN, adOpenKeyset, adLockOptimistic

       
    If rsRec.RecordCount >= 1 Then MsgBox "Access denied! Cannot delete this record." & vbCrLf & "There is a record bound with this position." & vbCrLf & "Please delete first Employee Record before continuing...", vbCritical: Exit Sub
     
    
    
If rsPosition.BOF And rsPosition.EOF = True Then
      
            MsgBox "Empty Database", vbInformation
            Exit Sub
            
 Else
 
        
 Dim msg As Integer

    msg = MsgBox("This will delete " & TxtPosition & ". Proceed? ", vbQuestion + vbYesNo, "Confirm Deleting...")
    
    If msg = vbYes Then
        
     
        Me.MousePointer = vbHourglass
        Call delposdesc_rankTable
        rsPosition.Delete
        rsPosition.Requery
        display_rec
        rsPosition.MoveFirst
        If rsPosition.EOF Then rsPosition.Requery: clear_ALL_ctrl: disable_nav_button
        MsgBox "Record deleted successfully...", vbInformation
        Me.MousePointer = vbDefault
    End If
 End If
 
 
End Sub

Sub delposdesc_rankTable()
Set rsdelrankpos = Nothing
    
    Set rsdelrankpos = New ADODB.Recordset
     rsdelrankpos.Open "SELECT * FROM [rank] where posdesc='" & TxtPosition & "'", CN, adOpenKeyset, adLockOptimistic
      rsdelrankpos.Delete
      rsdelrankpos.Requery
Set rsdelrankpos = Nothing
 End Sub

Private Sub CmdEdit_Click()
ToggleButtonSaveCancel "ON"
addflag = False
Call unlock_txt
TxtPosCode.SetFocus
End Sub

Private Sub CmdFirst_Click()
On Error Resume Next
     rsPosition.MoveFirst
      display_rec
    If rsPosition.BOF Then Exit Sub
    
End Sub

Private Sub CmdLast_Click()
On Error Resume Next
     rsPosition.MoveLast
      display_rec
    If rsPosition.EOF Then Exit Sub
       
End Sub

Private Sub CmdNext_Click()
On Error Resume Next
 rsPosition.MoveNext
    display_rec
        If rsPosition.EOF Then
             MsgBox "You have reached the End of File!", vbCritical + vbOKOnly
              rsPosition.MoveLast
        End If
End Sub

Private Sub CmdPrevious_Click()
    
On Error Resume Next
 rsPosition.MovePrevious
    display_rec

    If rsPosition.BOF Then
             MsgBox "You have reached the Beginning of File!", vbCritical + vbOKOnly
            display_rec
        End If
End Sub


Private Sub ToggleButtonSaveCancel(OnorOff As String)
 If UCase(OnorOff) = "ON" Then 'Click Add or Edit
    CmdAdd.Enabled = False
    CmdEdit.Enabled = False
    
    CmdSave.Enabled = True
    CmdCancel.Enabled = True
    
    CmdDelete.Enabled = False
  
    
    CmdFirst.Enabled = False
    CmdPrevious.Enabled = False
    CmdNext.Enabled = False
    CmdLast.Enabled = False
    
    TxtPosCode.Locked = False
    TxtPosition.Locked = False
    
Else
    CmdAdd.Enabled = True
    CmdEdit.Enabled = True
    
    CmdSave.Enabled = False
    CmdCancel.Enabled = False
    
    CmdDelete.Enabled = True

    
    CmdFirst.Enabled = True
    CmdPrevious.Enabled = True
    CmdNext.Enabled = True
    CmdLast.Enabled = True
    
    TxtPosCode.Locked = True
    TxtPosition.Locked = True
   
End If

End Sub

Private Sub CmdSave_Click()
Dim TempRs As New ADODB.Recordset

If IsNumeric(TxtPosCode) = True Then MsgBox "Invalid input. Please check it!", vbCritical, "ERROR": TxtPosCode.SetFocus: Exit Sub
If IsNumeric(TxtPosition) = True Then MsgBox "Invalid input. Please check it!", vbCritical, "ERROR": TxtPosition.SetFocus: Exit Sub

If Trim(TxtPosCode) = "" Then MsgBox "Sorry cannot continue!! Position Code is empty", vbCritical + vbOKOnly: Exit Sub
If Trim(TxtPosition) = "" Then MsgBox "Sorry cannot continue!! Position is empty", vbCritical + vbOKOnly:     Exit Sub

Me.MousePointer = vbHourglass

If addflag = True Then

    TempRs.Open "SELECT * from [position] where Poscode ='" & TxtPosCode & "'", CN, adOpenForwardOnly, adLockReadOnly
    
    If Not TempRs.EOF Then 'not empty
        
        MsgBox TxtPosCode & vbCrLf & vbCrLf & " already exist! Duplication of record is not allowed.", vbCritical
        TxtPosCode.SetFocus
        Me.MousePointer = vbDefault
        Exit Sub
        
    End If

    Set TempRs = Nothing

   rsPosition.AddNew 'blank record will be created

End If

Call saveIn_rankTable '-// Add Position Record on Rank Table

rsPosition.Fields("poscode") = TxtPosCode
rsPosition.Fields("posdesc") = TxtPosition
rsPosition.Update
rsPosition.Requery

MsgBox "Record was saved successfully!!", vbInformation + vbOKOnly
Me.MousePointer = vbDefault
Call lock_txt
ToggleButtonSaveCancel "OFF"

End Sub

Sub saveIn_rankTable()

Set rsRank = New ADODB.Recordset
rsRank.Open "SELECT * FROM rank", CN, adOpenStatic, adLockOptimistic

rsRank.AddNew
rsRank.Fields("posdesc") = TxtPosition
rsRank.Update
rsRank.Requery
Set rsRank = Nothing

End Sub

Private Sub Form_Load()
 Me.Icon = img.ListImages(1).ExtractIcon
 Me.Show
 Call lock_txt
 Set rsPosition = New ADODB.Recordset
 
 rsPosition.Open "SELECT poscode, posdesc from [position]", CN, adOpenKeyset, adLockOptimistic
    

    If rsPosition.BOF And rsPosition.EOF Then MsgBox "No record(s) to display!", vbCritical, "ERROR": disable_nav_button: Exit Sub

    ToggleButtonSaveCancel "OFF"
    display_rec
    
End Sub

Sub lock_txt()
TxtPosCode.Locked = True
TxtPosition.Locked = True
End Sub

Sub unlock_txt()
TxtPosCode.Locked = False
TxtPosition.Locked = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set rsPosition = Nothing
End Sub

Private Sub TxtPosCode_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then TxtPosition.SetFocus
End Sub

Private Sub TxtPosition_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then CmdSave_Click
End Sub

Private Sub display_rec()

On Error Resume Next
TxtPosCode = rsPosition.Fields("poscode") & " "
TxtPosition = rsPosition.Fields("posdesc") & " "

End Sub
