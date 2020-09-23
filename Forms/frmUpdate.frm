VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmUpdate 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   804
   ClientLeft      =   48
   ClientTop       =   48
   ClientWidth     =   4908
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   11  'Hourglass
   ScaleHeight     =   804
   ScaleWidth      =   4908
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Left            =   4800
      Top             =   480
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   475
      Left            =   120
      TabIndex        =   0
      Top             =   250
      Width           =   4695
      _ExtentX        =   8276
      _ExtentY        =   826
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label Label1 
      Caption         =   "Updating Database . Please wait..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   30
      Width           =   3615
   End
End
Attribute VB_Name = "frmUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Timer1.Enabled = True
Timer1.Interval = 80
End Sub

Private Sub Timer1_Timer()

If EnableBackup = True Then

    frmMain.MousePointer = vbHourglass
    ProgressBar1.Value = ProgressBar1.Value + 1
    
    If ProgressBar1.Value = 98 Then backup_rec
    If ProgressBar1.Value = 100 Then
        frmMain.MousePointer = vbDefault
        Unload Me
        MsgBox "Database successfully back-up.", vbInformation, "Backup"
    End If

Else

    frmMain.MousePointer = vbHourglass
    ProgressBar1.Value = ProgressBar1.Value + 1
    
    If ProgressBar1.Value = 98 Then restore_rec
    If ProgressBar1.Value = 100 Then
        frmMain.MousePointer = vbDefault
        Unload Me
        MsgBox "Database was restored successfully.", vbInformation, "Restore..."
    End If

End If

End Sub

