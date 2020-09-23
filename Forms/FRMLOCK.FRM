VERSION 5.00
Begin VB.Form frmlock 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   660
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   6000
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   660
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   -80
      Width           =   6005
      Begin VB.TextBox Text1 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1320
         PasswordChar    =   "*"
         TabIndex        =   1
         ToolTipText     =   "Type your password and press <Enter> in your keyboard."
         Top             =   240
         Width           =   4575
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Enter Password:"
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
         Left            =   120
         TabIndex        =   2
         Top             =   280
         Width           =   1815
      End
   End
End
Attribute VB_Name = "frmlock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_Activate()
Text1.SetFocus
End Sub

Private Sub Form_Load()

'------------------
'-Center Form
'------------------
Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)


If KeyAscii = 13 Then

    If Text1.Text <> Pword Then
      
        MsgBox "Invalid password. Please try again!", vbCritical, "Error"
        Text1.Text = ""
        Text1.SetFocus
        
    Else
        Unload Me
        
    End If
   
  
End If

End Sub

