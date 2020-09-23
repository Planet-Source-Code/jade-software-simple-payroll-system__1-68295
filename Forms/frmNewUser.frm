VERSION 5.00
Begin VB.Form frmNewUser 
   BackColor       =   &H000000C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New User"
   ClientHeight    =   1965
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3840
   Icon            =   "frmNewUser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1965
   ScaleWidth      =   3840
   Begin VB.TextBox TxtUsername 
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
      TabIndex        =   3
      Top             =   240
      Width           =   1935
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
      Left            =   2640
      TabIndex        =   1
      Top             =   1350
      Width           =   975
   End
   Begin VB.TextBox TxtConfirm 
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
      IMEMode         =   3  'DISABLE
      Left            =   1680
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   960
      Width           =   1935
   End
   Begin VB.TextBox TxtPass 
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
      IMEMode         =   3  'DISABLE
      Left            =   1680
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   600
      Width           =   1935
   End
   Begin VB.CommandButton CmdOk 
      Caption         =   "&Ok"
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
      Top             =   1350
      Width           =   975
   End
   Begin VB.Label Label4 
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Confirm Password"
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
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
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
      TabIndex        =   7
      Top             =   630
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Username"
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
      TabIndex        =   6
      Top             =   270
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   1720
      Left            =   120
      Top             =   120
      Width           =   3615
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0E0FF&
      Height          =   1720
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3615
   End
End
Attribute VB_Name = "frmNewUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()

If Trim(TxtUsername) = "" Then MsgBox "Empty fields. Please check it!", vbCritical: Exit Sub
If Trim(TxtPass) = "" Then MsgBox "Empty fields! Please check it!", vbCritical: Exit Sub
If Trim(TxtConfirm) = "" Then MsgBox "Empty fields! Please check it!", vbCritical: Exit Sub
 
    
Set Adors = New ADODB.Recordset
Adors.Open "select * from login where username = '" & TxtUsername & "'", CN, adOpenForwardOnly, adLockOptimistic

If Not Adors.EOF Then

   MsgBox "User already exist. Duplication is not allowed.", vbCritical
   Exit Sub
    
Else

 If TxtPass.Text = TxtConfirm.Text Then
 
        Adors.AddNew
        Adors.Fields!username = TxtUsername.Text
        Adors.Fields!password = TxtPass.Text
        Adors.Update
        Adors.Requery '-// refresh
        
        MsgBox "New User added succesfully!", vbInformation
        Unload Me
        
    Else
    
        MsgBox "Your password didn't match with your confirm password", vbInformation
        
    End If
End If

End Sub

Private Sub Form_Activate()
    TxtUsername.SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Set Adors = Nothing
    
End Sub
