VERSION 5.00
Begin VB.Form frmliving 
   BackColor       =   &H000000C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Living Allowance"
   ClientHeight    =   2040
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2640
   Icon            =   "frmliving.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   2640
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   960
      TabIndex        =   10
      Top             =   3600
      Width           =   495
   End
   Begin VB.CommandButton CmdClose 
      Caption         =   "&Close"
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
      Left            =   1800
      TabIndex        =   6
      Top             =   1560
      Width           =   735
   End
   Begin VB.CommandButton CmdUpdate 
      Caption         =   "&Update"
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
      TabIndex        =   5
      Top             =   1560
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
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   735
   End
   Begin VB.TextBox Text4 
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
      Left            =   960
      TabIndex        =   3
      Text            =   "2000"
      Top             =   960
      Width           =   1455
   End
   Begin VB.TextBox Text3 
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
      Left            =   960
      TabIndex        =   2
      Text            =   "1500"
      Top             =   600
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
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
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   "2"
      Top             =   960
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
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
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   0
      Text            =   "1"
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Living Allowance"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   8
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Code"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   240
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   3
      Height          =   1335
      Left            =   120
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0E0FF&
      Height          =   1335
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "frmliving"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdClose_Click()
Unload Me
End Sub

Sub lock_ctrl()

Text3.Locked = True
Text4.Locked = True

End Sub

Sub unlock_ctrl()

Text3.Locked = False
Text4.Locked = False

End Sub

Private Sub CmdEdit_Click()
unlock_ctrl
End Sub

Private Sub CmdUpdate_Click()

If IsNumeric(Text1) = False Then MsgBox "Invalid input! Please check it!", vbInformation: Exit Sub
If IsNumeric(Text2) = False Then MsgBox "Invalid input! Please check it!", vbInformation: Exit Sub
If IsNumeric(Text3) = False Then MsgBox "Invalid input! Please check it!", vbInformation: Exit Sub
If IsNumeric(Text4) = False Then MsgBox "Invalid input! Please check it!", vbInformation: Exit Sub

Call Modregistry.saving_living_rec
Call lock_ctrl
MsgBox "Record Succesfully Updated!", vbInformation, "Success..."


End Sub


Private Sub Form_Load()
 
 Call lock_ctrl
    
 Text5.Text = GetSetting("IsInstallLiving", "setting", "value")
 
 If Text5.Text = "Yes" Then
     
    Text1.Text = GetSetting("living_allowance", "cost", "living Code")
    Text2.Text = GetSetting("living_allowance", "cost", "living Code1")
    Text3.Text = GetSetting("living_allowance", "cost", "living allowance")
    Text4.Text = GetSetting("living_allowance", "cost", "living allowance2")
    
Else
    Call SaveSetting("IsInstallLiving", "setting", "value", "Yes")
    Call Modregistry.saving_living_rec
    
End If


End Sub




