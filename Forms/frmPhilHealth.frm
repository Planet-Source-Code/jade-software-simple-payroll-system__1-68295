VERSION 5.00
Begin VB.Form frmPhilHealth 
   BackColor       =   &H000000C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PhilHealth Table"
   ClientHeight    =   5490
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4305
   Icon            =   "frmPhilHealth.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5490
   ScaleWidth      =   4305
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   4920
      TabIndex        =   61
      Top             =   4080
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Height          =   25
      Left            =   120
      TabIndex        =   60
      Top             =   4920
      Width           =   4095
   End
   Begin VB.CommandButton CmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   2760
      TabIndex        =   59
      Top             =   5040
      Width           =   1455
   End
   Begin VB.CommandButton CmdUpdate 
      Caption         =   "&Update"
      Height          =   375
      Left            =   1320
      TabIndex        =   58
      Top             =   5040
      Width           =   1335
   End
   Begin VB.CommandButton CmdEdit 
      Caption         =   "&Edit"
      Height          =   375
      Left            =   120
      TabIndex        =   57
      Top             =   5040
      Width           =   1095
   End
   Begin VB.TextBox Text3 
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
      TabIndex        =   50
      Text            =   "1"
      Top             =   615
      Width           =   735
   End
   Begin VB.TextBox Text4 
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
      TabIndex        =   49
      Text            =   "2"
      Top             =   855
      Width           =   735
   End
   Begin VB.TextBox Text5 
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
      TabIndex        =   48
      Text            =   "3"
      Top             =   1095
      Width           =   735
   End
   Begin VB.TextBox Text6 
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
      TabIndex        =   47
      Text            =   "4"
      Top             =   1335
      Width           =   735
   End
   Begin VB.TextBox Text7 
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
      TabIndex        =   46
      Text            =   "5"
      Top             =   1575
      Width           =   735
   End
   Begin VB.TextBox Text8 
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
      TabIndex        =   45
      Text            =   "6"
      Top             =   1815
      Width           =   735
   End
   Begin VB.TextBox Text9 
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
      TabIndex        =   44
      Text            =   "7"
      Top             =   2055
      Width           =   735
   End
   Begin VB.TextBox Text10 
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
      TabIndex        =   43
      Text            =   "8"
      Top             =   2295
      Width           =   735
   End
   Begin VB.TextBox Text11 
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
      TabIndex        =   42
      Text            =   "9"
      Top             =   2535
      Width           =   735
   End
   Begin VB.TextBox Text12 
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
      TabIndex        =   41
      Text            =   "10"
      Top             =   2775
      Width           =   735
   End
   Begin VB.TextBox Text13 
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
      TabIndex        =   40
      Text            =   "11"
      Top             =   3015
      Width           =   735
   End
   Begin VB.TextBox Text14 
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
      TabIndex        =   39
      Text            =   "12"
      Top             =   3255
      Width           =   735
   End
   Begin VB.TextBox Text15 
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
      TabIndex        =   38
      Text            =   "13"
      Top             =   3495
      Width           =   735
   End
   Begin VB.TextBox Text16 
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
      TabIndex        =   37
      Text            =   "14"
      Top             =   3735
      Width           =   735
   End
   Begin VB.TextBox Text17 
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
      TabIndex        =   36
      Text            =   "15"
      Top             =   3975
      Width           =   735
   End
   Begin VB.TextBox Text18 
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
      TabIndex        =   35
      Text            =   "16"
      Top             =   4215
      Width           =   735
   End
   Begin VB.TextBox Text19 
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
      TabIndex        =   34
      Text            =   "17"
      Top             =   4455
      Width           =   735
   End
   Begin VB.TextBox txt2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   33
      Text            =   "5000"
      Top             =   855
      Width           =   1215
   End
   Begin VB.TextBox txt3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   32
      Text            =   "6000"
      Top             =   1095
      Width           =   1215
   End
   Begin VB.TextBox txt4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   31
      Text            =   "7000"
      Top             =   1335
      Width           =   1215
   End
   Begin VB.TextBox txt5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   30
      Text            =   "8000"
      Top             =   1575
      Width           =   1215
   End
   Begin VB.TextBox txt6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   29
      Text            =   "9000"
      Top             =   1815
      Width           =   1215
   End
   Begin VB.TextBox txt7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   28
      Text            =   "10000"
      Top             =   2055
      Width           =   1215
   End
   Begin VB.TextBox txt8 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   27
      Text            =   "11000"
      Top             =   2295
      Width           =   1215
   End
   Begin VB.TextBox txt9 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   26
      Text            =   "12000"
      Top             =   2535
      Width           =   1215
   End
   Begin VB.TextBox txt10 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   25
      Text            =   "13000"
      Top             =   2775
      Width           =   1215
   End
   Begin VB.TextBox txt11 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   24
      Text            =   "14000"
      Top             =   3015
      Width           =   1215
   End
   Begin VB.TextBox txt12 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   23
      Text            =   "15000"
      Top             =   3255
      Width           =   1215
   End
   Begin VB.TextBox txt13 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   22
      Text            =   "16000"
      Top             =   3495
      Width           =   1215
   End
   Begin VB.TextBox txt14 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   21
      Text            =   "17000"
      Top             =   3735
      Width           =   1215
   End
   Begin VB.TextBox txt15 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   20
      Text            =   "18000"
      Top             =   3975
      Width           =   1215
   End
   Begin VB.TextBox txt16 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   19
      Text            =   "19000"
      Top             =   4215
      Width           =   1215
   End
   Begin VB.TextBox txt17 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   18
      Text            =   "20000"
      Top             =   4455
      Width           =   1215
   End
   Begin VB.TextBox txtE1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2880
      TabIndex        =   17
      Text            =   "50"
      Top             =   615
      Width           =   1215
   End
   Begin VB.TextBox txtE2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2880
      TabIndex        =   16
      Text            =   "62.5"
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox txtE3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2880
      TabIndex        =   15
      Text            =   "75"
      Top             =   1095
      Width           =   1215
   End
   Begin VB.TextBox txtE4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2880
      TabIndex        =   14
      Text            =   "87.5"
      Top             =   1335
      Width           =   1215
   End
   Begin VB.TextBox txtE5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2880
      TabIndex        =   13
      Text            =   "100"
      Top             =   1575
      Width           =   1215
   End
   Begin VB.TextBox txtE6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2880
      TabIndex        =   12
      Text            =   "112.5"
      Top             =   1815
      Width           =   1215
   End
   Begin VB.TextBox txtE7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2880
      TabIndex        =   11
      Text            =   "125"
      Top             =   2055
      Width           =   1215
   End
   Begin VB.TextBox txtE8 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2880
      TabIndex        =   10
      Text            =   "137.5"
      Top             =   2300
      Width           =   1215
   End
   Begin VB.TextBox txtE9 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2880
      TabIndex        =   9
      Text            =   "150"
      Top             =   2535
      Width           =   1215
   End
   Begin VB.TextBox txtE10 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2880
      TabIndex        =   8
      Text            =   "162.5"
      Top             =   2775
      Width           =   1215
   End
   Begin VB.TextBox txtE11 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2880
      TabIndex        =   7
      Text            =   "175"
      Top             =   3015
      Width           =   1215
   End
   Begin VB.TextBox txtE12 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2880
      TabIndex        =   6
      Text            =   "187.5"
      Top             =   3255
      Width           =   1215
   End
   Begin VB.TextBox txtE13 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2880
      TabIndex        =   5
      Text            =   "200"
      Top             =   3495
      Width           =   1215
   End
   Begin VB.TextBox txtE14 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2880
      TabIndex        =   4
      Text            =   "212.5"
      Top             =   3735
      Width           =   1215
   End
   Begin VB.TextBox txtE15 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2880
      TabIndex        =   3
      Text            =   "225"
      Top             =   3975
      Width           =   1215
   End
   Begin VB.TextBox txtE16 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2880
      TabIndex        =   2
      Text            =   "237.5"
      Top             =   4215
      Width           =   1215
   End
   Begin VB.TextBox txtE17 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2880
      TabIndex        =   1
      Text            =   "250"
      Top             =   4455
      Width           =   1215
   End
   Begin VB.TextBox txt1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   0
      Text            =   "4000"
      Top             =   615
      Width           =   1215
   End
   Begin VB.Shape Shape2 
      BorderWidth     =   3
      Height          =   4320
      Left            =   1200
      Top             =   495
      Width           =   1455
   End
   Begin VB.Shape Shape3 
      BorderWidth     =   3
      Height          =   4320
      Left            =   2760
      Top             =   495
      Width           =   1455
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0E0FF&
      Height          =   4320
      Left            =   2760
      TabIndex        =   52
      Top             =   495
      Width           =   1455
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   3
      Height          =   4320
      Left            =   120
      Top             =   495
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "PhilHealth Code"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   0
      TabIndex        =   55
      Top             =   80
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0E0FF&
      Height          =   4320
      Left            =   1200
      TabIndex        =   54
      Top             =   495
      Width           =   1455
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Salary Base"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   1320
      TabIndex        =   53
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Share"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   2760
      TabIndex        =   51
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0E0FF&
      Height          =   4320
      Left            =   120
      TabIndex        =   56
      Top             =   495
      Width           =   975
   End
End
Attribute VB_Name = "frmPhilHealth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdClose_Click()
Unload Me
End Sub

Sub lock_ctrl()

txtE1.Locked = True
txtE2.Locked = True
txtE3.Locked = True
txtE4.Locked = True
txtE5.Locked = True
txtE6.Locked = True
txtE7.Locked = True
txtE8.Locked = True
txtE9.Locked = True
txtE10.Locked = True
txtE11.Locked = True
txtE12.Locked = True
txtE13.Locked = True
txtE14.Locked = True
txtE15.Locked = True
txtE16.Locked = True
txtE17.Locked = True

End Sub

Sub unlock_ctrl()

txtE1.Locked = False
txtE2.Locked = False
txtE3.Locked = False
txtE4.Locked = False
txtE5.Locked = False
txtE6.Locked = False
txtE7.Locked = False
txtE8.Locked = False
txtE9.Locked = False
txtE10.Locked = False
txtE11.Locked = False
txtE12.Locked = False
txtE13.Locked = False
txtE14.Locked = False
txtE15.Locked = False
txtE16.Locked = False
txtE17.Locked = False

End Sub

Private Sub CmdEdit_Click()
unlock_ctrl
txtE1.SetFocus
End Sub

Private Sub CmdUpdate_Click()

Call lock_ctrl
Modregistry.saving_PH_share
MsgBox "Record Succesfully Updated!", vbInformation, "Success..."

End Sub

Private Sub Form_Load()

Call lock_ctrl

 Text1.Text = GetSetting("IsEnable", "setting", "value")
 
 If Text1.Text = "Yes" Then
    
      
    txtE1 = GetSetting("PhilHealth Table", "EShare", "PH Share1")
    txtE2 = GetSetting("PhilHealth Table", "EShare", "PH Share2")
    txtE3 = GetSetting("PhilHealth Table", "EShare", "PH Share3")
    txtE4 = GetSetting("PhilHealth Table", "EShare", "PH Share4")
    txtE5 = GetSetting("PhilHealth Table", "EShare", "PH Share5")
    txtE6 = GetSetting("PhilHealth Table", "EShare", "PH Share6")
    txtE7 = GetSetting("PhilHealth Table", "EShare", "PH Share7")
    txtE8 = GetSetting("PhilHealth Table", "EShare", "PH Share8")
    txtE9 = GetSetting("PhilHealth Table", "EShare", "PH Share9")
    txtE10 = GetSetting("PhilHealth Table", "EShare", "PH Share10")
    txtE11 = GetSetting("PhilHealth Table", "EShare", "PH Share11")
    txtE12 = GetSetting("PhilHealth Table", "EShare", "PH Share12")
    txtE13 = GetSetting("PhilHealth Table", "EShare", "PH Share13")
    txtE14 = GetSetting("PhilHealth Table", "EShare", "PH Share14")
    txtE15 = GetSetting("PhilHealth Table", "EShare", "PH Share15")
    txtE16 = GetSetting("PhilHealth Table", "EShare", "PH Share16")
    txtE17 = GetSetting("PhilHealth Table", "EShare", "PH Share17")

 
     
Else

    Call SaveSetting("IsEnable", "setting", "value", "Yes")
    Call Modregistry.saving_PH_share
    
End If



End Sub
