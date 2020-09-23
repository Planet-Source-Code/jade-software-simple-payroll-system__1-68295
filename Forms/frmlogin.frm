VERSION 5.00
Begin VB.Form frmlogin 
   BackColor       =   &H000000FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login"
   ClientHeight    =   2385
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4065
   Icon            =   "frmlogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   4065
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Height          =   1455
      Left            =   0
      TabIndex        =   6
      Top             =   960
      Width           =   4095
      Begin VB.TextBox txtpassword 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1080
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   600
         Width           =   2775
      End
      Begin VB.TextBox txtusername 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   1080
         TabIndex        =   1
         Top             =   240
         Width           =   2775
      End
      Begin VB.CommandButton cmdCancel 
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
         Height          =   315
         Left            =   2760
         TabIndex        =   4
         Top             =   975
         Width           =   1110
      End
      Begin VB.CommandButton cmdok 
         Caption         =   "&Log-In"
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
         Left            =   1560
         TabIndex        =   3
         Top             =   960
         Width           =   1110
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Password:"
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
         Left            =   0
         TabIndex        =   8
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "User Name:"
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
         Left            =   0
         TabIndex        =   7
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Southernside Montessori School"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   735
      Left            =   1200
      TabIndex        =   9
      Top             =   120
      Width           =   2895
   End
   Begin VB.Image Image1 
      Height          =   945
      Left            =   0
      Picture         =   "frmlogin.frx":08CA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   945
      Left            =   0
      Top             =   0
      Width           =   4095
   End
   Begin VB.Label Label7 
      BackColor       =   &H000000FF&
      Caption         =   "0"
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
      Left            =   1680
      TabIndex        =   0
      Top             =   3405
      Width           =   855
   End
   Begin VB.Label Label4 
      BackColor       =   &H000000FF&
      Caption         =   "Attempt:"
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
      Left            =   840
      TabIndex        =   5
      Top             =   3405
      Width           =   735
   End
End
Attribute VB_Name = "frmlogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim attempt

Private Sub cmdcancel_Click()
End
End Sub

Private Sub cmdOK_Click()

'-// assing value to variables
Uname = txtusername.Text
Pword = txtpassword.Text

Set rsLogin = New ADODB.Recordset  'create recorset
 rsLogin.Open "Select * from login where username='" & txtusername.Text & "' and password='" & txtpassword.Text & "'", CN, adOpenKeyset, adLockOptimistic
  
 With rsLogin
 
    If .RecordCount = 1 Then
        
         Unload Me
         Load frmMain
         frmMain.Show
         
         frmMain.StatusBar1.Panels(3).Text = Uname
         frmMain.StatusBar1.Panels(6).Text = Format(Time, "HH:MM:SS AM/PM")
         frmMain.StatusBar1.Panels(9).Text = Date
         frmMain.mnuLogOff.Caption = "Log-Off..." & " " & Uname
        
    Else
    
        MsgBox "Invalid Login... Please check it!", vbCritical, "Login-Error"
        attempt = attempt + 1
        If attempt = 3 Then MsgBox "Sorry, You have reached the maximum allowable login.", vbCritical: End
        
         
  End If

End With

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

If UnloadMode = 0 Then End

End Sub

Private Sub Form_Activate()

        txtusername.SetFocus
        attempt = 0
        
End Sub

Private Sub Form_Unload(Cancel As Integer)

'-//Clear variable from computer memory
Set rsLogin = Nothing

End Sub

Private Sub txtpassword_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then cmdOK_Click
   
End Sub

Private Sub txtusername_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then txtpassword.SetFocus

End Sub

