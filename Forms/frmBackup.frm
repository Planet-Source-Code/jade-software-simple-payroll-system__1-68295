VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBackup 
   BackColor       =   &H000000C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Backup & Restore Database"
   ClientHeight    =   3150
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4335
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   4335
   Begin VB.FileListBox File1 
      Height          =   1650
      Left            =   2280
      TabIndex        =   8
      Top             =   600
      Width           =   1815
   End
   Begin VB.DirListBox Dir1 
      Height          =   1665
      Left            =   240
      TabIndex        =   7
      Top             =   600
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2280
      TabIndex        =   6
      Top             =   240
      Width           =   1815
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   240
      TabIndex        =   5
      Top             =   240
      Width           =   1935
   End
   Begin VB.Frame Frame2 
      Height          =   30
      HelpContextID   =   10
      Left            =   120
      TabIndex        =   3
      Top             =   2520
      Width           =   4095
   End
   Begin VB.CommandButton CmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   2760
      TabIndex        =   0
      Top             =   2640
      Width           =   1455
   End
   Begin VB.CommandButton CmdRestore 
      Caption         =   "&Restore"
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   2640
      Width           =   1335
   End
   Begin VB.CommandButton CmdBackUp 
      Caption         =   "&BackUp"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   2640
      Width           =   1335
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4680
      Top             =   3960
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
            Picture         =   "frmBackup.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Height          =   2295
      Left            =   120
      Top             =   120
      Width           =   4095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0E0FF&
      Height          =   2295
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   4095
   End
End
Attribute VB_Name = "frmBackup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdBackUp_Click()
    EnableBackup = True
    frmUpdate.Show vbModal
End Sub

Private Sub CmdClose_Click()
Unload Me
End Sub

Private Sub CmdRestore_Click()

EnableBackup = False
If Trim(frmBackup.Text1) = True Then MsgBox "No File was Selected.", vbCritical, "WARNING!": Exit Sub
If MsgBox("WARNING! Are you sure you want to  re-store?" & vbCrLf & "This will overwrite your current database. Proceed ? ", vbYesNo + vbExclamation, "WARNING!") = vbYes Then frmUpdate.Show vbModal


End Sub

Private Sub Form_Load()
  Me.Icon = ImageList1.ListImages(1).ExtractIcon
  File1.Pattern = "*.mdb;*.sms"
End Sub

Private Sub Dir1_Change()
 File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
On Error GoTo Mali:

  Dir1.Path = Drive1.Drive
  File1.Path = Drive1.Drive
  
  Exit Sub
Mali:

  MsgBox Err.Description, vbCritical + vbOKOnly
  Resume Next
 
End Sub

Private Sub File1_Click()
  Text1.Text = File1.FileName
End Sub
