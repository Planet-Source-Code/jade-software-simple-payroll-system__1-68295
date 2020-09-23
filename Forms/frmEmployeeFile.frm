VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmEmployeeFile 
   BackColor       =   &H000000C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Employee Information  "
   ClientHeight    =   9096
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   6588
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9096
   ScaleWidth      =   6588
   Begin VB.CommandButton cmdRank 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3050
      Style           =   1  'Graphical
      TabIndex        =   74
      ToolTipText     =   "Browse record."
      Top             =   5640
      Width           =   315
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   20
      Top             =   5640
      Width           =   1335
   End
   Begin VB.TextBox txtmi 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5880
      MaxLength       =   1
      TabIndex        =   3
      Top             =   840
      Width           =   495
   End
   Begin VB.TextBox ttfname 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3840
      TabIndex        =   2
      Top             =   840
      Width           =   1575
   End
   Begin VB.CommandButton CmdClose 
      Caption         =   "C&lose"
      Height          =   375
      Left            =   5400
      TabIndex        =   27
      Top             =   8625
      Width           =   1095
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4440
      TabIndex        =   26
      Top             =   8625
      Width           =   975
   End
   Begin VB.CommandButton CmdSearch 
      Caption         =   "&Search"
      Height          =   375
      Left            =   3480
      TabIndex        =   25
      Top             =   8625
      Width           =   975
   End
   Begin VB.CommandButton CmdDelete 
      Caption         =   "&Delete"
      Height          =   375
      Left            =   2640
      TabIndex        =   24
      Top             =   8625
      Width           =   855
   End
   Begin VB.CommandButton CmdUpdate 
      Caption         =   "&Update"
      Height          =   375
      Left            =   1800
      TabIndex        =   23
      Top             =   8625
      Width           =   855
   End
   Begin VB.CommandButton CmdEdit 
      Caption         =   "&Edit"
      Height          =   375
      Left            =   960
      TabIndex        =   22
      Top             =   8625
      Width           =   855
   End
   Begin VB.CommandButton CmdLast 
      Caption         =   ">>"
      Height          =   265
      Left            =   4800
      TabIndex        =   68
      ToolTipText     =   "Last"
      Top             =   8280
      Width           =   1695
   End
   Begin VB.CommandButton CmdNext 
      Caption         =   ">"
      Height          =   265
      Left            =   3120
      TabIndex        =   69
      ToolTipText     =   "Next"
      Top             =   8280
      Width           =   1695
   End
   Begin VB.CommandButton CmdPrevious 
      Caption         =   "<"
      Height          =   265
      Left            =   1560
      TabIndex        =   70
      ToolTipText     =   "Previous"
      Top             =   8280
      Width           =   1575
   End
   Begin VB.CommandButton CmdAdd 
      Caption         =   "&Add"
      Height          =   375
      Left            =   120
      TabIndex        =   21
      Top             =   8625
      Width           =   855
   End
   Begin VB.CommandButton CmdFirst 
      Caption         =   "<<"
      Height          =   265
      Left            =   120
      TabIndex        =   71
      ToolTipText     =   "First"
      Top             =   8280
      Width           =   1455
   End
   Begin VB.TextBox TxtSSS 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   63
      Top             =   7400
      Width           =   1575
   End
   Begin VB.TextBox TxtPhilhealth 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   62
      Top             =   7750
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4560
      Locked          =   -1  'True
      TabIndex        =   61
      Top             =   7400
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4560
      Locked          =   -1  'True
      TabIndex        =   60
      Text            =   "50.00"
      Top             =   7750
      Width           =   1815
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   54
      Top             =   6570
      Width           =   1575
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4560
      Locked          =   -1  'True
      TabIndex        =   53
      Top             =   6570
      Width           =   1815
   End
   Begin VB.ComboBox CmbTaxStat 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "frmEmployeeFile.frx":0000
      Left            =   4800
      List            =   "frmEmployeeFile.frx":0002
      TabIndex        =   17
      Top             =   4800
      Width           =   1575
   End
   Begin VB.TextBox TxtSyears 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4800
      TabIndex        =   15
      Top             =   4440
      Width           =   1575
   End
   Begin VB.ComboBox CmbPosition 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "frmEmployeeFile.frx":0004
      Left            =   1680
      List            =   "frmEmployeeFile.frx":0006
      TabIndex        =   14
      Top             =   4440
      Width           =   1695
   End
   Begin VB.ComboBox CmbDepartment 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "frmEmployeeFile.frx":0008
      Left            =   1680
      List            =   "frmEmployeeFile.frx":000A
      TabIndex        =   16
      Top             =   4840
      Width           =   1695
   End
   Begin VB.TextBox txtBasicPay 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4800
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   4080
      Width           =   1575
   End
   Begin VB.TextBox TxtSemi 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   4080
      Width           =   1695
   End
   Begin VB.ComboBox cmbrice 
      Height          =   315
      ItemData        =   "frmEmployeeFile.frx":000C
      Left            =   1680
      List            =   "frmEmployeeFile.frx":0016
      TabIndex        =   18
      Top             =   5250
      Width           =   1695
   End
   Begin VB.ComboBox cmbliving 
      Height          =   315
      ItemData        =   "frmEmployeeFile.frx":0020
      Left            =   4800
      List            =   "frmEmployeeFile.frx":002A
      TabIndex        =   19
      Top             =   5200
      Width           =   1575
   End
   Begin VB.TextBox TxtEaddress 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1560
      TabIndex        =   9
      Top             =   2520
      Width           =   4815
   End
   Begin VB.TextBox TxtPhoneNo 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1560
      MaxLength       =   11
      TabIndex        =   10
      Top             =   2880
      Width           =   1935
   End
   Begin VB.TextBox TxtEadd 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1560
      TabIndex        =   11
      Top             =   3240
      Width           =   3495
   End
   Begin VB.TextBox TxtEcode 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   480
      Width           =   1575
   End
   Begin VB.TextBox TxtSname 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1200
      TabIndex        =   1
      Top             =   840
      Width           =   1815
   End
   Begin VB.TextBox TxtEage 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   1200
      Width           =   735
   End
   Begin VB.OptionButton OptionMale 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Male"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   2760
      TabIndex        =   5
      Top             =   1200
      Width           =   735
   End
   Begin VB.OptionButton OptionFemale 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Female"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   3600
      TabIndex        =   6
      Top             =   1200
      Width           =   855
   End
   Begin VB.ComboBox CmbCivilStat 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmEmployeeFile.frx":0034
      Left            =   1200
      List            =   "frmEmployeeFile.frx":0044
      TabIndex        =   7
      Text            =   "CmbCivilStat"
      Top             =   1580
      Width           =   1935
   End
   Begin MSComCtl2.DTPicker DTbirthdate 
      Height          =   375
      Left            =   4080
      TabIndex        =   8
      Top             =   1580
      Width           =   1455
      _ExtentX        =   2561
      _ExtentY        =   656
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   43646977
      CurrentDate     =   38728
   End
   Begin MSComctlLib.ImageList img 
      Left            =   6600
      Top             =   9120
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmployeeFile.frx":006D
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList i16x16 
      Left            =   6600
      Top             =   9120
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmployeeFile.frx":19FF
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label31 
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "M.I."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   5520
      TabIndex        =   73
      Top             =   840
      Width           =   375
   End
   Begin VB.Label Label30 
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "FirstName"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3070
      TabIndex        =   72
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label LblSSS 
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "SSS Premium"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   67
      Top             =   7400
      Width           =   1095
   End
   Begin VB.Label Label29 
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "PhilHealth"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   66
      Top             =   7750
      Width           =   1215
   End
   Begin VB.Label Label28 
      BackStyle       =   0  'Transparent
      Caption         =   "Withholding Tax"
      Height          =   255
      Left            =   3360
      TabIndex        =   65
      Top             =   7400
      Width           =   1335
   End
   Begin VB.Label Label27 
      BackStyle       =   0  'Transparent
      Caption         =   "Pag-ibig"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3360
      TabIndex        =   64
      Top             =   7750
      Width           =   1215
   End
   Begin VB.Shape Shape5 
      BorderWidth     =   3
      Height          =   885
      Left            =   120
      Top             =   7260
      Width           =   6375
   End
   Begin VB.Label Label26 
      BackStyle       =   0  'Transparent
      Caption         =   "Standard Deduction"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   59
      Top             =   7000
      Width           =   2175
   End
   Begin VB.Label Label25 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   120
      TabIndex        =   58
      Top             =   7260
      Width           =   6375
   End
   Begin VB.Label Label24 
      BackStyle       =   0  'Transparent
      Caption         =   "Standard Allowance"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   57
      Top             =   6195
      Width           =   2175
   End
   Begin VB.Shape Shape4 
      BorderWidth     =   3
      Height          =   495
      Left            =   120
      Top             =   6450
      Width           =   6375
   End
   Begin VB.Label Label23 
      BackStyle       =   0  'Transparent
      Caption         =   "Living Allowance"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   56
      Top             =   6600
      Width           =   1215
   End
   Begin VB.Label Label22 
      BackStyle       =   0  'Transparent
      Caption         =   "Rice Allowance"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3360
      TabIndex        =   55
      Top             =   6600
      Width           =   1215
   End
   Begin VB.Label Label21 
      BackColor       =   &H00C0E0FF&
      Height          =   495
      Left            =   120
      TabIndex        =   52
      Top             =   6450
      Width           =   6375
   End
   Begin VB.Label Label20 
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Tax Status Code"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3480
      TabIndex        =   51
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Label LblPosition 
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Position"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   50
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Label LblDepartment 
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Department"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   49
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Label LblSyears 
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Years of Service"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3480
      TabIndex        =   48
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Label Label19 
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Monthly rate"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3480
      TabIndex        =   47
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Semi-monthly rate"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   46
      Top             =   4080
      Width           =   1335
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Rice Allowance Code"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   45
      Top             =   5160
      Width           =   1335
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Living Allowance Code"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3480
      TabIndex        =   44
      Top             =   5160
      Width           =   1215
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Rank"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   43
      Top             =   5700
      Width           =   975
   End
   Begin VB.Shape Shape3 
      BorderWidth     =   3
      Height          =   2175
      Left            =   120
      Top             =   3960
      Width           =   6375
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Account Information"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   42
      Top             =   3720
      Width           =   2175
   End
   Begin VB.Label Label13 
      BackColor       =   &H00C0E0FF&
      Height          =   2175
      Left            =   120
      TabIndex        =   41
      Top             =   3960
      Width           =   6375
   End
   Begin VB.Label Label12 
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   40
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label11 
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Cell./Phone No."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   39
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label Label10 
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "E-mail Add"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   38
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Contact Information"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   37
      Top             =   2160
      Width           =   2175
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00000000&
      BorderWidth     =   3
      Height          =   1245
      Left            =   120
      Top             =   2400
      Width           =   6375
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0E0FF&
      Height          =   1215
      Left            =   120
      TabIndex        =   36
      Top             =   2400
      Width           =   6375
   End
   Begin VB.Label LblBirthdate 
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Birthdate"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3240
      TabIndex        =   34
      Top             =   1635
      Width           =   855
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Gender"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   2040
      TabIndex        =   33
      Top             =   1215
      Width           =   735
   End
   Begin VB.Label Label6 
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Code"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   32
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label5 
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Surname"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   31
      Top             =   885
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Age"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   30
      Top             =   1260
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Civil Status"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   29
      Top             =   1635
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000000&
      BorderWidth     =   3
      Height          =   1695
      Left            =   120
      Top             =   360
      Width           =   6375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0E0FF&
      Height          =   1695
      Left            =   120
      TabIndex        =   28
      Top             =   360
      Width           =   6375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Personal Information"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   35
      Top             =   105
      Width           =   2175
   End
End
Attribute VB_Name = "frmEmployeeFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SemiMonthtlyRate        As Double
Dim addflag                 As Boolean

Sub increment_EmpNo()
Dim Curr_Num, Newnum

Set Rs = New ADODB.Recordset
Rs.Open "SELECT * FROM [employeefile]", CN, adOpenStatic, adLockOptimistic
Curr_Num = Rs.RecordCount

Newnum = "E00000" & Curr_Num + 1
TxtEcode.Text = Newnum

Set Rs = Nothing

End Sub

Private Sub ToggleButtonUpdateCancel(OnorOff As String)

 If UCase(OnorOff) = "ON" Then
 
    cmdAdd.Enabled = False
    CmdEdit.Enabled = False
    CmdUpdate.Enabled = True
    CmdCancel.Enabled = True
    cmdSearch.Enabled = False
    CmdDelete.Enabled = False
    CmdClose.Enabled = False
    CmdFirst.Enabled = False
    CmdPrevious.Enabled = False
    CmdNext.Enabled = False
    CmdLast.Enabled = False
    
    
Else

    cmdAdd.Enabled = True
    CmdEdit.Enabled = True
    CmdUpdate.Enabled = False
    CmdCancel.Enabled = False
    cmdSearch.Enabled = True
    CmdDelete.Enabled = True
    CmdClose.Enabled = True
    CmdFirst.Enabled = True
    CmdPrevious.Enabled = True
    CmdNext.Enabled = True
    CmdLast.Enabled = True
    
           
End If

End Sub

Sub clear_all()
   
    ttfname.Text = ""
    TxtEage.Text = ""
    TxtSname.Text = ""
    txtmi.Text = ""
    CmbCivilStat.Text = ""
    TxtEaddress.Text = ""
    TxtPhoneNo.Text = ""
    TxtEadd.Text = ""
    CmbPosition.Text = ""
    TxtSyears.Text = ""
    CmbDepartment.Text = ""
    CmbTaxStat.Text = ""
    cmbrice.Text = ""
    cmbliving.Text = ""
    Text5.Text = ""
    TxtBasicPay.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    Text1.Text = ""
   TxtSemi = ""
   TxtPhilHealth = ""
   TxtSSS.Text = ""
   
End Sub

Sub lock_txt()

    TxtEcode.Locked = True
    ttfname.Locked = True
    TxtSname.Locked = True
    txtmi.Locked = True
    CmbCivilStat.Locked = True
    TxtEaddress.Locked = True
    TxtPhoneNo.Locked = True
    TxtEadd.Locked = True
    CmbPosition.Locked = True
    TxtSyears.Locked = True
    CmbDepartment.Locked = True
    CmbTaxStat.Locked = True
    cmbrice.Locked = True
    cmbliving.Locked = True
    Text5.Locked = True
    
End Sub

Sub unlock_txt()
     
    TxtEcode.Locked = False
    ttfname.Locked = False
    TxtSname.Locked = False
    txtmi.Locked = False
    CmbCivilStat.Locked = False
    TxtEaddress.Locked = False
    TxtPhoneNo.Locked = False
    TxtEadd.Locked = False
    CmbPosition.Locked = False
    TxtSyears.Locked = False
    CmbDepartment.Locked = False
    CmbTaxStat.Locked = False
    cmbrice.Locked = False
    cmbliving.Locked = False
    Text5.Locked = False
  
End Sub

Private Sub cmbliving_Click()

frmliving.Hide
If cmbliving = "1" Then Text3.Text = Val(frmliving.Text3.Text) / 2: Exit Sub
If cmbliving = "2" Then Text3.Text = Val(frmliving.Text4.Text) / 2: Exit Sub

End Sub

Private Sub cmbrice_Click()

frmriceallowance.Hide

If cmbrice.Text = "1" Then Text4.Text = Val(frmriceallowance.Text3.Text) / 2: Exit Sub
If cmbrice.Text = "2" Then Text4.Text = Val(frmriceallowance.Text4.Text) / 2: Exit Sub

End Sub

Private Sub CmbTaxStat_Click()

frmWithTax.Hide

If TxtSemi.Text = "" Or TxtBasicPay.Text = "" Then MsgBox "Empty Field. Please check it!", vbCritical, "ERROR": TxtBasicPay.SetFocus: Exit Sub

Dim difference, percentage As Double

'// For Single
If CmbTaxStat = "S" Then

Select Case Val(TxtSemi)

Case 1 To 832.99

    difference = Val(TxtSemi) - Val(frmWithTax.txts1.Text)
    percentage = difference * Val(frmWithTax.txtop1.Text)
    Text1 = Val(frmWithTax.txtex1.Text) + Val(percentage)
              
 Case 833 To 1249.99

    difference = Val(TxtSemi) - Val(frmWithTax.txts2.Text)
    percentage = difference * Val(frmWithTax.txtop2.Text)
    Text1 = Val(frmWithTax.txtex2.Text) + Val(percentage)
    
Case 1250 To 2082.99
    
    difference = Val(TxtSemi) - Val(frmWithTax.txts3.Text)
    percentage = difference * Val(frmWithTax.txtop3.Text)
    Text1 = Val(frmWithTax.txtex3.Text) + Val(percentage)
             
Case 2083 To 3749.99
    
    difference = Val(TxtSemi) - Val(frmWithTax.txts4.Text)
    percentage = difference * Val(frmWithTax.txtop4.Text)
    Text1 = Val(frmWithTax.txtex4.Text) + Val(percentage)
        
Case 3750 To 6666.99
     
     difference = Val(TxtSemi) - Val(frmWithTax.txts5.Text)
     percentage = difference * Val(frmWithTax.txtop5.Text)
     Text1 = Val(frmWithTax.txtex5.Text) + Val(percentage)
     
Case 6667 To 11249.99
     
     difference = Val(TxtSemi) - Val(frmWithTax.txts6.Text)
     percentage = difference * Val(frmWithTax.txtop6.Text)
     Text1 = Val(frmWithTax.txtex6.Text) + Val(percentage)

Case 11250 To 21666.99
    
     difference = Val(TxtSemi) - Val(frmWithTax.txts7.Text)
     percentage = difference * Val(frmWithTax.txtop7.Text)
     Text1 = Val(frmWithTax.txtex7.Text) + Val(percentage)

Case Is > 21667
                    
     difference = Val(TxtSemi) - Val(frmWithTax.txts8.Text)
     percentage = difference * Val(frmWithTax.txtop8.Text)
     Text1 = Val(frmWithTax.txtex8.Text) + Val(percentage)
          
End Select

 '// For Head of the Family without dependent
ElseIf CmbTaxStat = "HF" Then


  Select Case Val(TxtSemi)
  
     Case 1 To 1041.99
    
      difference = Val(TxtSemi) - Val(frmWithTax.txthf1.Text)
      percentage = difference * Val(frmWithTax.txtop1.Text)
      Text1 = Val(frmWithTax.txtex1.Text) + Val(percentage)
      
     Case 1042 To 1457.99
                 
      difference = Val(TxtSemi) - Val(frmWithTax.txthf2.Text)
      percentage = difference * Val(frmWithTax.txtop2.Text)
      Text1 = Val(frmWithTax.txtex2.Text) + Val(percentage)
                
     Case 1458 To 2291.99
       
       difference = Val(TxtSemi) - Val(frmWithTax.txthf3.Text)
       percentage = difference * Val(frmWithTax.txtop3.Text)
       Text1 = Val(frmWithTax.txtex3.Text) + Val(percentage)
       
     Case 2292 To 3957.99
                    
       difference = Val(TxtSemi) - Val(frmWithTax.txthf4.Text)
       percentage = difference * Val(frmWithTax.txtop4.Text)
       Text1 = Val(frmWithTax.txtex4.Text) + percentage
                
     Case 3958 To 6874.99
        
       difference = Val(TxtSemi) - Val(frmWithTax.txthf5.Text)
       percentage = difference * Val(frmWithTax.txtop5.Text)
       Text1 = Val(frmWithTax.txtex5.Text) + percentage
                        
    Case 6875 To 11457.99
       
       difference = Val(TxtSemi) - Val(frmWithTax.txthf6.Text)
       percentage = difference * Val(frmWithTax.txtop6.Text)
       Text1 = Val(frmWithTax.txtex6.Text) + percentage
    
    Case 11458 To 21874.99
                    
       difference = Val(TxtSemi) - Val(frmWithTax.txthf7.Text)
       percentage = difference * Val(frmWithTax.txtop7.Text)
       Text1 = Val(frmWithTax.txtex7.Text) + percentage
    
    Case Is > 21875
       difference = Val(TxtSemi) - Val(frmWithTax.txthf8.Text)
       percentage = difference * Val(frmWithTax.txtop8.Text)
       Text1 = Val(frmWithTax.txtex8.Text) + percentage
    
    End Select
            
 '// For Married without dependent
ElseIf CmbTaxStat = "ME" Then
    
    Select Case Val(TxtSemi)
    
      Case 1 To 1332.99
        difference = Val(TxtSemi) - Val(frmWithTax.txtme1.Text)
        percentage = difference * Val(frmWithTax.txtop1.Text)
        Text1 = Val(frmWithTax.txtex1.Text) + percentage
                
      Case 1333 To 1749.99
      
        difference = Val(TxtSemi) - Val(frmWithTax.txtme2.Text)
        percentage = difference * Val(frmWithTax.txtop2.Text)
        Text1 = Val(frmWithTax.txtex2.Text) + percentage
        
      Case 1750 To 2582.99
      
         difference = Val(TxtSemi) - Val(frmWithTax.txtme3.Text)
         percentage = difference * Val(frmWithTax.txtop3.Text)
         Text1 = Val(frmWithTax.txtex3.Text) + percentage
         
      Case 2583 To 4249.99
         
         difference = Val(TxtSemi) - Val(frmWithTax.txtme4.Text)
         percentage = difference * Val(frmWithTax.txtop4.Text)
         Text1 = Val(frmWithTax.txtex4.Text) + percentage
                
       Case 4250 To 7166.99
         
         difference = Val(TxtSemi) - Val(frmWithTax.txtme5.Text)
         percentage = difference * Val(frmWithTax.txtop5.Text)
         Text1 = Val(frmWithTax.txtex5.Text) + percentage
                
       Case 7167 To 11749.99
                    
         difference = Val(TxtSemi) - Val(frmWithTax.txtme6.Text)
         percentage = difference * Val(frmWithTax.txtop6.Text)
         Text1 = Val(frmWithTax.txtex6.Text) + percentage
                
        Case 11750 To 22166.99
         
         difference = Val(TxtSemi) - Val(frmWithTax.txtme7.Text)
         percentage = difference * Val(frmWithTax.txtop7.Text)
         Text1 = Val(frmWithTax.txtex7.Text) + percentage
               
       Case Is > 22167
                    
         difference = Val(TxtSemi) - Val(frmWithTax.txtme8.Text)
         percentage = difference * Val(frmWithTax.txtop8.Text)
         Text1 = Val(frmWithTax.txtex8.Text) + percentage
       
       End Select
       
 '// For Head of the Family with 1 Dependent
ElseIf CmbTaxStat = "HF1" Then

     Select Case Val(TxtSemi)
          
          Case 1 To 1374.99
          
               difference = Val(TxtSemi) - Val(frmWithTax.txthf11.Text)
               percentage = difference * Val(frmWithTax.txtop1.Text)
               Text1 = Val(frmWithTax.txtex1.Text) + percentage
               
          Case 1375 To 1791.99
                
                difference = Val(TxtSemi) - Val(frmWithTax.txthf12.Text)
                percentage = difference * Val(frmWithTax.txtop2.Text)
                Text1 = Val(frmWithTax.txtex2.Text) + percentage
                
                Case 1792 To 2324.99
                    
                    difference = Val(TxtSemi) - Val(frmWithTax.txthf13.Text)
                    percentage = difference * Val(frmWithTax.txtop3.Text)
                    Text1 = Val(frmWithTax.txtex3.Text) + percentage
                    
                Case 2625 To 4291.99
                    
                    difference = Val(TxtSemi) - Val(frmWithTax.txthf14.Text)
                    percentage = difference * Val(frmWithTax.txtop4.Text)
                    Text1 = Val(frmWithTax.txtex4.Text) + percentage
                    
                Case 4292 To 7207.99
                    
                    difference = Val(TxtSemi) - Val(frmWithTax.txthf15.Text)
                    percentage = difference * Val(frmWithTax.txtop5.Text)
                    Text1 = Val(frmWithTax.txtex5.Text) + percentage
                    
                Case 7208 To 11791.99
                    
                    difference = Val(TxtSemi) - Val(frmWithTax.txthf16.Text)
                    percentage = difference * Val(frmWithTax.txtop6.Text)
                    Text1 = Val(frmWithTax.txtex6.Text) + percentage
                    
                Case 11792 To 22207.99
                    
                    difference = Val(TxtSemi) - Val(frmWithTax.txthf17.Text)
                    percentage = difference * Val(frmWithTax.txtop7.Text)
                    Text1 = Val(frmWithTax.txtex7.Text) + percentage
                    
               Case Is > 22208
                    difference = Val(TxtSemi) - Val(frmWithTax.txthf18.Text)
                    percentage = difference * Val(frmWithTax.txtop8.Text)
                    Text1 = Val(frmWithTax.txtex8.Text) + percentage
            End Select
       
       
       
  '// For Head of the Family with 2 Dependent
      ElseIf CmbTaxStat = "HF2" Then
      
            Select Case Val(TxtSemi)
            
                Case 1 To 1707.99
                    
                    difference = Val(TxtSemi) - Val(frmWithTax.txthf21.Text)
                    percentage = difference * Val(frmWithTax.txtop1.Text)
                    Text1 = Val(frmWithTax.txtex1.Text) + percentage
                    
                Case 1708 To 2124.99
                
                    difference = Val(TxtSemi) - Val(frmWithTax.txthf22.Text)
                    percentage = difference * Val(frmWithTax.txtop2.Text)
                    Text1 = Val(frmWithTax.txtex2.Text) + percentage
                
                Case 2125 To 2957.99
                    
                    difference = Val(TxtSemi) - Val(frmWithTax.txthf23.Text)
                    percentage = difference * Val(frmWithTax.txtop3.Text)
                    Text1 = Val(frmWithTax.txtex3.Text) + percentage
                
                Case 2958 To 4624.99
                    
                    difference = Val(TxtSemi) - Val(frmWithTax.txthf24.Text)
                    percentage = difference * Val(frmWithTax.txtop4.Text)
                    Text1 = Val(frmWithTax.txtex4.Text) + percentage
                    
                Case 4625 To 7541.99
                
                    difference = Val(TxtSemi) - Val(frmWithTax.txthf25.Text)
                    percentage = difference * Val(frmWithTax.txtop5.Text)
                    Text1 = Val(frmWithTax.txtex5.Text) + percentage
                    
                Case 7542 To 12124.99
                    
                    difference = Val(TxtSemi) - Val(frmWithTax.txthf26.Text)
                    percentage = difference * Val(frmWithTax.txtop6.Text)
                    Text1 = Val(frmWithTax.txtex6.Text) + percentage
                    
                Case 12125 To 22541.99
                    
                    difference = Val(TxtSemi) - Val(frmWithTax.txthf27.Text)
                    percentage = difference * Val(frmWithTax.txtop7.Text)
                    Text1 = Val(frmWithTax.txtex7.Text) + percentage
               
               Case Is > 22542
                    
                    difference = Val(TxtSemi) - Val(frmWithTax.txthf28.Text)
                    percentage = difference * Val(frmWithTax.txtop8.Text)
                    Text1 = Val(frmWithTax.txtex8.Text) + percentage
                    
            End Select
      
      
'// For Head of the Family with 3 Dependent
     
    ElseIf CmbTaxStat = "HF3" Then
    
            Select Case Val(TxtSemi)
            
                Case 1 To 2041.99
                    
                    difference = Val(TxtSemi) - Val(frmWithTax.txthf31.Text)
                    percentage = difference * Val(frmWithTax.txtop1.Text)
                    Text1 = Val(frmWithTax.txtex1.Text) + percentage
                    
                Case 2042 To 2457.99
                    
                    difference = Val(TxtSemi) - Val(frmWithTax.txthf32.Text)
                    percentage = difference * Val(frmWithTax.txtop2.Text)
                    Text1 = Val(frmWithTax.txtex2.Text) + percentage
                    
                Case 2458 To 3291.99
                
                    difference = Val(TxtSemi) - Val(frmWithTax.txthf33.Text)
                    percentage = difference * Val(frmWithTax.txtop3.Text)
                    Text1 = Val(frmWithTax.txtex3.Text) + percentage
                
                Case 3292 To 4957.99
                    
                    difference = Val(TxtSemi) - Val(frmWithTax.txthf34.Text)
                    percentage = difference * Val(frmWithTax.txtop4.Text)
                    Text1 = Val(frmWithTax.txtex4.Text) + percentage
                
                Case 4958 To 7874.99
                    
                    difference = Val(TxtSemi) - Val(frmWithTax.txthf35.Text)
                    percentage = difference * Val(frmWithTax.txtop5.Text)
                    Text1 = Val(frmWithTax.txtex5.Text) + percentage
                    
                Case 7875 To 12457.99
                    
                    difference = Val(TxtSemi) - Val(frmWithTax.txthf36.Text)
                    percentage = difference * Val(frmWithTax.txtop6.Text)
                    Text1 = Val(frmWithTax.txtex6.Text) + percentage
                    
                Case 12458 To 22874.99
                
                    difference = Val(TxtSemi) - Val(frmWithTax.txthf37.Text)
                    percentage = difference * Val(frmWithTax.txtop7.Text)
                    Text1 = Val(frmWithTax.txtex7.Text) + percentage
                    
               Case Is > 22875
                    
                    difference = Val(TxtSemi) - Val(frmWithTax.txthf38.Text)
                    percentage = difference * Val(frmWithTax.txtop8.Text)
                    Text1 = Val(frmWithTax.txtex8.Text) + percentage
                    
            End Select
            
            
            '-// HEAD OF THE FAMILY WITH 4 DEPENDENTS
          ElseIf CmbTaxStat = "HF4" Then
          
            Select Case Val(TxtSemi)
                Case 1 To 2374.99
                    
                    difference = Val(TxtSemi) - Val(frmWithTax.txthf41.Text)
                    percentage = difference * Val(frmWithTax.txtop1.Text)
                    Text1 = Val(frmWithTax.txtex1.Text) + percentage
                    
                Case 2375 To 2791.99
                    
                    difference = Val(TxtSemi) - Val(frmWithTax.txthf42.Text)
                    percentage = difference * Val(frmWithTax.txtop2.Text)
                    Text1 = Val(frmWithTax.txtex2.Text) + percentage
                    
                Case 2792 To 3624.99
                    
                    difference = Val(TxtSemi) - Val(frmWithTax.txthf43.Text)
                    percentage = difference * Val(frmWithTax.txtop3.Text)
                    Text1 = Val(frmWithTax.txtex3.Text) + percentage
                    
                Case 3625 To 5291.99
                    
                    difference = Val(TxtSemi) - Val(frmWithTax.txthf44.Text)
                    percentage = difference * Val(frmWithTax.txtop4.Text)
                    Text1 = Val(frmWithTax.txtex4.Text) + percentage
                    
                Case 5292 To 8207.99
                    
                    difference = Val(TxtSemi) - Val(frmWithTax.txthf45.Text)
                    percentage = difference * Val(frmWithTax.txtop5.Text)
                    Text1 = Val(frmWithTax.txtex5.Text) + percentage
                    
                Case 8208 To 12791.99
                    
                    difference = Val(TxtSemi) - Val(frmWithTax.txthf46.Text)
                    percentage = difference * Val(frmWithTax.txtop6.Text)
                    Text1 = Val(frmWithTax.txtex6.Text) + percentage
                
                Case 12792 To 23207.99
                    
                    difference = Val(TxtSemi) - Val(frmWithTax.txthf47.Text)
                    percentage = difference * Val(frmWithTax.txtop7.Text)
                    Text1 = Val(frmWithTax.txtex7.Text) + percentage
                    
               Case Is > 23208
                    
                    difference = Val(TxtSemi) - Val(frmWithTax.txthf48.Text)
                    percentage = difference * Val(frmWithTax.txtop8.Text)
                    Text1 = Val(frmWithTax.txtex8.Text) + percentage
                    
            End Select
                            
        '-// MARRIED WITH 1 DEPENDENT
          ElseIf CmbTaxStat = "ME1" Then
          
            Select Case Val(TxtSemi)
            
                Case 1 To 1666.99
                    
                    difference = Val(TxtSemi) - Val(frmWithTax.txtme11.Text)
                    percentage = difference * Val(frmWithTax.txtop1.Text)
                    Text1 = Val(frmWithTax.txtex1.Text) + percentage
                    
                Case 1667 To 2082.99
                    
                    difference = Val(TxtSemi) - Val(frmWithTax.txtme12.Text)
                    percentage = difference * Val(frmWithTax.txtop2.Text)
                    Text1 = Val(frmWithTax.txtex2.Text) + percentage
                    
                Case 2083 To 2916.99
                    
                    difference = Val(TxtSemi) - Val(frmWithTax.txtme13.Text)
                    percentage = difference * Val(frmWithTax.txtop3.Text)
                    Text1 = Val(frmWithTax.txtex3.Text) + percentage
                    
                Case 2917 To 4582.99
                
                    difference = Val(TxtSemi) - Val(frmWithTax.txtme14.Text)
                    percentage = difference * Val(frmWithTax.txtop4.Text)
                    Text1 = Val(frmWithTax.txtex4.Text) + percentage
                    
                Case 4583 To 7499.99
                    
                    difference = Val(TxtSemi) - Val(frmWithTax.txtme15.Text)
                    percentage = difference * Val(frmWithTax.txtop5.Text)
                    Text1 = Val(frmWithTax.txtex5.Text) + percentage
                
                Case 7500 To 12082.99
                    
                    difference = Val(TxtSemi) - Val(frmWithTax.txtme16.Text)
                    percentage = difference * Val(frmWithTax.txtop6.Text)
                    Text1 = Val(frmWithTax.txtex6.Text) + percentage
                    
                Case 12083 To 22499.99
                    
                    difference = Val(TxtSemi) - Val(frmWithTax.txtme17.Text)
                    percentage = difference * Val(frmWithTax.txtop7.Text)
                    Text1 = Val(frmWithTax.txtex7.Text) + percentage
               
               Case Is > 22500
                    
                    difference = Val(TxtSemi) - Val(frmWithTax.txtme18.Text)
                    percentage = difference * Val(frmWithTax.txtop8.Text)
                    Text1 = Val(frmWithTax.txtex8.Text) + percentage
                    
            End Select
            
        '-// MARRIED WITH 2
          ElseIf CmbTaxStat = "ME2" Then
          
            Select Case Val(TxtSemi)
                
                Case 1 To 1999.99
                    
                    difference = Val(TxtSemi) - Val(frmWithTax.txtme21.Text)
                    percentage = difference * Val(frmWithTax.txtop1.Text)
                    Text1 = Val(frmWithTax.txtex1.Text) + percentage
                    
                Case 2000 To 2416.99
                    
                    difference = Val(TxtSemi) - Val(frmWithTax.txtme22.Text)
                    percentage = difference * Val(frmWithTax.txtop2.Text)
                    Text1 = Val(frmWithTax.txtex2.Text) + percentage
                    
                    
                Case 2417 To 3249.99
                    
                    difference = Val(TxtSemi) - Val(frmWithTax.txtme23.Text)
                    percentage = difference * Val(frmWithTax.txtop3.Text)
                    Text1 = Val(frmWithTax.txtex3.Text) + percentage
                    
                Case 3250 To 4916.99
                    
                    difference = Val(TxtSemi) - Val(frmWithTax.txtme24.Text)
                    percentage = difference * Val(frmWithTax.txtop4.Text)
                    Text1 = Val(frmWithTax.txtex4.Text) + percentage
                
                Case 4917 To 7832.99
                    
                    difference = Val(TxtSemi) - Val(frmWithTax.txtme25.Text)
                    percentage = difference * Val(frmWithTax.txtop5.Text)
                    Text1 = Val(frmWithTax.txtex5.Text) + percentage
                    
                Case 7833 To 12416.99
                    
                    difference = Val(TxtSemi) - Val(frmWithTax.txtme26.Text)
                    percentage = difference * Val(frmWithTax.txtop6.Text)
                    Text1 = Val(frmWithTax.txtex6.Text) + percentage
                    
                Case 12417 To 22832.99
                    
                    difference = Val(TxtSemi) - Val(frmWithTax.txtme27.Text)
                    percentage = difference * Val(frmWithTax.txtop7.Text)
                    Text1 = Val(frmWithTax.txtex7.Text) + percentage
                    
               Case Is > 22833
               
                    difference = Val(TxtSemi) - Val(frmWithTax.txtme28.Text)
                    percentage = difference * Val(frmWithTax.txtop8.Text)
                    Text1 = Val(frmWithTax.txtex8.Text) + percentage
                    
            End Select
            
            '-// MARIED WITH 3 DEPENDENTS
          ElseIf CmbTaxStat = "ME3" Then
          
            Select Case Val(TxtSemi)
                
                Case 1 To 2332.99
                
                    difference = Val(TxtSemi) - Val(frmWithTax.txtme31.Text)
                    percentage = difference * Val(frmWithTax.txtop1.Text)
                    Text1 = Val(frmWithTax.txtex1.Text) + percentage
                    
                Case 2333 To 2749.99
                
                    difference = Val(TxtSemi) - Val(frmWithTax.txtme32.Text)
                    percentage = difference * Val(frmWithTax.txtop2.Text)
                    Text1 = Val(frmWithTax.txtex2.Text) + percentage
                    
                Case 2750 To 3582.99
                    
                    difference = Val(TxtSemi) - Val(frmWithTax.txtme33.Text)
                    percentage = difference * Val(frmWithTax.txtop3.Text)
                    Text1 = Val(frmWithTax.txtex3.Text) + percentage
                    
                Case 3583 To 5249.99
                
                    difference = Val(TxtSemi) - Val(frmWithTax.txtme34.Text)
                    percentage = difference * Val(frmWithTax.txtop4.Text)
                    Text1 = Val(frmWithTax.txtex4.Text) + percentage
                    
                Case 5250 To 8166.99
                    
                    difference = Val(TxtSemi) - Val(frmWithTax.txtme35.Text)
                    percentage = difference * Val(frmWithTax.txtop5.Text)
                    Text1 = Val(frmWithTax.txtex5.Text) + percentage
                    
                Case 8167 To 12749.99
                    
                    difference = Val(TxtSemi) - Val(frmWithTax.txtme36.Text)
                    percentage = difference * Val(frmWithTax.txtop6.Text)
                    Text1 = Val(frmWithTax.txtex6.Text) + percentage
                
                Case 12750 To 23166.99
                    
                    difference = Val(TxtSemi) - Val(frmWithTax.txtme37.Text)
                    percentage = difference * Val(frmWithTax.txtop7.Text)
                    Text1 = Val(frmWithTax.txtex7.Text) + percentage
                    
               Case Is > 23167
                    
                    difference = Val(TxtSemi) - Val(frmWithTax.txtme38.Text)
                    percentage = difference * Val(frmWithTax.txtop8.Text)
                    Text1 = Val(frmWithTax.txtex8.Text) + percentage
                    
            End Select
            
        '-// MARRIED WITH 4 DEPENDEDENTS
          ElseIf CmbTaxStat = "ME4" Then
          
            Select Case Val(TxtSemi)
                
                Case 1 To 2666.99
                    
                    difference = Val(TxtSemi) - Val(frmWithTax.txtme41.Text)
                    percentage = difference * Val(frmWithTax.txtop1.Text)
                    Text1 = Val(frmWithTax.txtex1.Text) + percentage
                
                Case 2667 To 3082.99
                    
                    difference = Val(TxtSemi) - Val(frmWithTax.txtme42.Text)
                    percentage = difference * Val(frmWithTax.txtop2.Text)
                    Text1 = Val(frmWithTax.txtex2.Text) + percentage
                
                Case 3083 To 3916.99
                    
                    difference = Val(TxtSemi) - Val(frmWithTax.txtme43.Text)
                    percentage = difference * Val(frmWithTax.txtop3.Text)
                    Text1 = Val(frmWithTax.txtex3.Text) + percentage
                    
                Case 3917 To 5582.99
                    
                    difference = Val(TxtSemi) - Val(frmWithTax.txtme44.Text)
                    percentage = difference * Val(frmWithTax.txtop4.Text)
                    Text1 = Val(frmWithTax.txtex4.Text) + percentage
                
                Case 5583 To 8499.99
                    
                    difference = Val(TxtSemi) - Val(frmWithTax.txtme45.Text)
                    percentage = difference * Val(frmWithTax.txtop5.Text)
                    Text1 = Val(frmWithTax.txtex5.Text) + percentage
                    
                Case 8500 To 13082.99
                
                    difference = Val(TxtSemi) - Val(frmWithTax.txtme46.Text)
                    percentage = difference * Val(frmWithTax.txtop6.Text)
                    Text1 = Val(frmWithTax.txtex6.Text) + percentage
                    
                Case 13083 To 23499.99
                    
                    difference = Val(TxtSemi) - Val(frmWithTax.txtme47.Text)
                    percentage = difference * Val(frmWithTax.txtop7.Text)
                    Text1 = Val(frmWithTax.txtex7.Text) + percentage
               
               Case Is > 23500
                    
                    difference = Val(TxtSemi) - Val(frmWithTax.txtme48.Text)
                    percentage = difference * Val(frmWithTax.txtop8.Text)
                    Text1 = Val(frmWithTax.txtex8.Text) + percentage
                    
            End Select
            
End If

End Sub

Private Sub CmdAdd_Click()
increment_EmpNo
addflag = True
cmdRank.Enabled = True
ToggleButtonUpdateCancel "ON"
Call unlock_txt
Call clear_all
TxtEage.Text = "0"
TxtPhoneNo.Text = "0"
TxtSname.SetFocus
TxtBasicPay.Text = "0"
TxtSemi.Text = "0"
Text3.Text = "0"
Text1.Text = "0"
TxtSSS.Text = "0"
TxtPhilHealth.Text = "0"
Text4.Text = "0"
TxtEcode.Locked = True
End Sub

Private Sub cmdcancel_Click()

ToggleButtonUpdateCancel "OFF"
Call lock_txt
If adoEmp.BOF And adoEmp.EOF Then disable_nav_button: Exit Sub
 display_records

End Sub

Private Sub CmdClose_Click()
Unload Me
End Sub

Private Sub cmdPhilHealth_Click()
frmselectrank.Show
frmselectrank.SetFocus
End Sub

Private Sub CmdDelete_Click()
On Error Resume Next
If Trim(TxtEcode) = "" Or Trim(TxtSname) = "" Then MsgBox "No record  found to delete!", vbCritical: Exit Sub

If adoEmp.BOF And adoEmp.EOF = True Then
      
            MsgBox "Empty Database", vbInformation
            Exit Sub
            
 Else
 
        
 Dim msg As Integer

    msg = MsgBox("This will delete " & vbCrLf & vbCrLf & "Employee Code      : " & TxtEcode & vbCrLf & "Employee Surname : " & TxtSname.Text & vbCrLf & vbCrLf & "Proceed? ", vbQuestion + vbYesNo, "Confirm Deleting...")
    
    If msg = vbYes Then
    
        Me.MousePointer = vbHourglass
        Delete_Attend_Comp
         adoEmp.Delete
         adoEmp.Requery
         display_records
         adoEmp.MoveFirst
         If adoEmp.EOF Then adoEmp.Requery: TxtEcode.Text = "": clear_all: disable_nav_button
         MsgBox "Record deleted successfully...", vbInformation
        
        Me.MousePointer = vbDefault
    End If
 End If
   
    
 End Sub

Sub Delete_Attend_Comp()
    
    adoEmp.Delete
    adoEmp.Requery

End Sub

Private Sub CmdEdit_Click()

If adoEmp.BOF And adoEmp.EOF Then MsgBox "No record(s) to edit", vbCritical, "ERROR":  Exit Sub

addflag = False
cmdRank.Enabled = True
ToggleButtonUpdateCancel "ON"
Call unlock_txt

End Sub

Private Sub CmdFirst_Click()
On Error Resume Next
     adoEmp.MoveFirst
      display_records
    If adoEmp.BOF Then Exit Sub
    
End Sub

Private Sub CmdLast_Click()
On Error Resume Next
     adoEmp.MoveLast
     display_records
    If adoEmp.EOF Then Exit Sub
       
End Sub

Private Sub CmdNext_Click()
On Error Resume Next
 adoEmp.MoveNext
    display_records
        If adoEmp.EOF Then
             MsgBox "You have reached the End of File!", vbCritical + vbOKOnly
              adoEmp.MoveLast
        End If
End Sub

Private Sub CmdPrevious_Click()
    
On Error Resume Next
 adoEmp.MovePrevious
   display_records
    If adoEmp.BOF Then
             MsgBox "You have reached the Beginning of File!", vbCritical + vbOKOnly
           display_records
        End If
End Sub

Private Sub cmdRank_Click()
If Trim(CmbPosition) = "" Or CmbPosition.Text = "CmbPosition" Then MsgBox "Invalid input. Please select position!", vbCritical: Exit Sub
frmselectrank.Show
frmselectrank.SetFocus
End Sub

Private Sub cmdSearch_Click()
frmsearchEmp.Show
frmsearchEmp.SetFocus
End Sub

Private Sub CmdUpdate_Click()
Dim TempRs As New ADODB.Recordset

If Trim(TxtPhoneNo) = "" Then TxtPhoneNo.Text = 0
If IsNumeric(TxtPhoneNo) = False Then MsgBox "Invalid input. Please check the value!", vbCritical, "ERROR": TxtPhoneNo.SetFocus: Exit Sub
If IsNumeric(TxtBasicPay) = False Then MsgBox "Invalid input. Please check the value!", vbCritical, "ERROR": TxtBasicPay.Text = "": TxtBasicPay.SetFocus: Exit Sub

If Trim(TxtEcode) = "" Then MsgBox "Sorry cannot continue!! Employee code is empty.", vbCritical + vbOKOnly: Exit Sub
If Trim(TxtSname) = "" Then MsgBox "Sorry cannot continue!! Surname is empty.", vbCritical + vbOKOnly: Exit Sub
If Trim(ttfname) = "" Then MsgBox "Sorry cannot continue!! FirstName is empty.", vbCritical + vbOKOnly: Exit Sub
If Trim(txtmi) = "" Then MsgBox "Sorry cannot continue!! Middle Initial is empty.", vbCritical + vbOKOnly: Exit Sub
If Trim(TxtEage) = "" Then MsgBox "Sorry cannot continue!! Age is empty.", vbCritical + vbOKOnly: Exit Sub
If Trim(CmbCivilStat) = "" Then MsgBox "Sorry cannot continue!! Civil Status is empty.", vbCritical + vbOKOnly: Exit Sub
If Trim(TxtEaddress) = "" Then MsgBox "Sorry cannot continue!! Address is empty.", vbCritical + vbOKOnly: Exit Sub
If Trim(Val(TxtSemi)) = "" Then MsgBox "Sorry cannot continue!! Semi-Monthly Rate is empty.", vbCritical + vbOKOnly: Exit Sub
If Trim(TxtBasicPay) = "" Then MsgBox "Sorry cannot continue!! Monthly Rate is empty.", vbCritical + vbOKOnly: Exit Sub
If Trim(TxtSyears) = "" Then MsgBox "Sorry cannot continue!! Years of Service is empty.", vbCritical + vbOKOnly: Exit Sub
If Trim(CmbDepartment) = "" Then MsgBox "Sorry cannot continue!! Department is empty.", vbCritical + vbOKOnly: Exit Sub
If Trim(cmbliving) = "" Then MsgBox "Sorry cannot Continue!! Living Allowance Code is empty.", vbCritical + vbOKOnly: Exit Sub
If Trim(cmbrice) = "" Then MsgBox "Sorry cannot Continue!! Rice Allowance Code is empty.", vbCritical + vbOKOnly: Exit Sub
If Trim(CmbTaxStat) = "" Then MsgBox "Sorry cannot Continue!! Tax Status Code is empty.", vbCritical + vbOKOnly: Exit Sub

Me.MousePointer = vbHourglass

If addflag = True Then

    TempRs.Open "SELECT * FROM employeefile where Ecode ='" & TxtEcode & "'", CN, adOpenForwardOnly, adLockReadOnly
    If Not TempRs.EOF Then 'not empty
    MsgBox TxtEcode & " employee code already exist! Duplication is not allowed", vbCritical
    TxtEcode.SetFocus
    Exit Sub
    End If

    Set TempRs = Nothing
   
    adoEmp.AddNew
    
End If

With adoEmp

    !ECODE = TxtEcode
    !fname = ttfname.Text
    !sname = TxtSname.Text
    !mi = txtmi.Text
    !eage = TxtEage.Text
    !civil_status = CmbCivilStat
    !birthdate = DTbirthdate.Value
    !eaddress = TxtEaddress.Text
    !phone_no = Val(TxtPhoneNo.Text)
    !email_add = TxtEadd.Text
    !semi_monthly = TxtSemi.Text
    !basicpay = TxtBasicPay.Text
    !posdesc = CmbPosition.Text
    !syears = Val(TxtSyears.Text)
    !depdesc = CmbDepartment.Text
    !taxheadercode = CmbTaxStat.Text
    !rice_all_code = cmbrice.Text
    !living_all_code = cmbliving.Text
    !rank = Text5.Text
    !living_value = Text3.Text
    !rice_value = Text4.Text
    !SSSPremium = TxtSSS.Text
    !PHHealthValue = TxtPhilHealth
    !Pagibig = Text2.Text
    !withTaxvalue = Text1.Text
    
If OptionMale.Value = True Then !gender = "m"
If OptionFemale.Value = True Then !gender = "f"

.Update
.Requery
Call lock_txt

MsgBox "Record was saved successfully!!", vbInformation + vbOKOnly
Me.MousePointer = vbDefault
End With

ToggleButtonUpdateCancel "OFF"

End Sub

Private Sub DTbirthdate_Change()
'-// Compute your Age as of Now()
TxtEage.Text = script.ComputeAccurateAge(DTbirthdate, Now())
End Sub

Private Sub Form_Load()

Me.Show
Call lock_txt
cmdRank.Picture = i16x16.ListImages(1).Picture

display_rec_incombo1
Me.Icon = Img.ListImages(1).ExtractIcon
Set adoEmp = New ADODB.Recordset

adoEmp.Open "SELECT * FROM employeefile", CN, adOpenStatic, adLockPessimistic
If adoEmp.EOF Then MsgBox "No record(s) to display!", vbCritical, "Error":  disable_nav_button: Exit Sub
display_records


End Sub

Sub display_rec_incombo1()
Set Rs = New ADODB.Recordset

Rs.Open "SELECT * FROM [position]", CN, adOpenForwardOnly, adLockPessimistic
    
    Do Until Rs.EOF
    CmbPosition.AddItem Rs!posdesc
    Rs.MoveNext
    Loop
    
Set Rs = Nothing
    
Set Rs = New ADODB.Recordset

Rs.Open "SELECT * FROM [department]", CN, adOpenForwardOnly, adLockPessimistic
    
    Do Until Rs.EOF
    CmbDepartment.AddItem Rs!depdesc
    Rs.MoveNext
    Loop
 Set Rs = Nothing
  
Set Adors = New ADODB.Recordset


Adors.Open "SELECT * FROM [taxheader]", CN, adOpenForwardOnly, adLockPessimistic
    
    Do Until Adors.EOF
    CmbTaxStat.AddItem Adors!taxheadercode
    Adors.MoveNext
    Loop
 Set Adors = Nothing
 
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set adoEmp = Nothing
End Sub


Private Sub txtBasicPay_Change()

frmssstable.Hide
'-// Computation for Semi-Monthly Rate
SemiMonthtlyRate = Val(TxtBasicPay.Text) / 2
TxtSemi.Text = SemiMonthtlyRate

'-// Case for SSS Premium (divide into 2 to get semi-monthly share)
Select Case Val(TxtBasicPay)

'-// No SSS Premium
  Case Is <= 999.99
         
        TxtSSS = 0
            
'-// SSS Code 1
   Case 1000 To 1249.99
       TxtSSS = Val(frmssstable.txtE1) / 2
       
'-// SSS Code 2
   Case 1250 To 1749.99
    TxtSSS = Val(frmssstable.txtE2) / 2
         
'-// SSS Code 3
   Case 1750 To 2249.99
    TxtSSS = Val(frmssstable.txtE3) / 2
         
'-// SSS Code 4
   Case 2250 To 2749.99
    TxtSSS = Val(frmssstable.txtE4) / 2
       
'-// SSS Code 5
   Case 2750 To 3249.99
    TxtSSS = Val(frmssstable.txtE5) / 2
       
'-// SSS Code 6
   Case 3250 To 3749.99
    TxtSSS = Val(frmssstable.txtE6) / 2
  
       
'-// SSS Code 7
   Case 3750 To 4249.99
    TxtSSS = Val(frmssstable.txtE7) / 2
             
'-// SSS Code 8
   Case 4250 To 4749.99
    TxtSSS = Val(frmssstable.txtE8) / 2
    
  '-// SSS Code 9
   Case 4750 To 5249.99
    TxtSSS = Val(frmssstable.txtE9) / 2
    
'-// SSS Code 10
   Case 5250 To 5749.99
    TxtSSS = Val(frmssstable.txtE10) / 2
    
'-// SSS Code 11
   Case 5750 To 6249.99
    TxtSSS = Val(frmssstable.txtE11) / 2
    
'-// SSS Code 12
   Case 6250 To 6749.99
    TxtSSS = Val(frmssstable.txtE12) / 2
    
'-// SSS Code 13
   Case 6750 To 7249.99
    TxtSSS = Val(frmssstable.txtE13) / 2
    
'-// SSS Code 14
   Case 7250 To 7749.99
    TxtSSS = Val(frmssstable.txtE14) / 2
    
'-// SSS Code 15
   Case 7750 To 8249.99
    TxtSSS = Val(frmssstable.txtE15) / 2
    
'-// SSS Code 16
   Case 8250 To 8749.99
    TxtSSS = Val(frmssstable.txtE16) / 2
    
'-// SSS Code 17
   Case 8750 To 9249.99
    TxtSSS = Val(frmssstable.txtE17) / 2
    
'-// SSS Code 18
   Case 9250 To 9749.99
    TxtSSS = Val(frmssstable.txtE18) / 2
    
'-// SSS Code 19
   Case 9750 To 10249.99
    TxtSSS = Val(frmssstable.txtE19) / 2
    
'-// SSS Code 20
   Case 10250 To 10749.99
    TxtSSS = Val(frmssstable.txtE20) / 2
    
'-// SSS Code 21
   Case 10750 To 11249.99
    TxtSSS = Val(frmssstable.txtE21) / 2
    
'-// SSS Code 22
   Case 11250 To 11749.99
    TxtSSS = Val(frmssstable.txtE22) / 2
    
'-// SSS Code 23
   Case 11750 To 12249.99
    TxtSSS = Val(frmssstable.txtE23) / 2
    
'-// SSS Code 24
   Case 12250 To 12749.99
    TxtSSS = Val(frmssstable.txtE24) / 2
    
'-// SSS Code 25
   Case 12750 To 13249.99
    TxtSSS = Val(frmssstable.txtE25) / 2
    
'-// SSS Code 26
   Case 13250 To 13749.99
    TxtSSS = Val(frmssstable.txtE26) / 2
    
'-// SSS Code 27
   Case 13750 To 14249.99
    TxtSSS = Val(frmssstable.txtE27) / 2
    
'-// SSS Code 28
   Case 14250 To 14749.99
    TxtSSS = Val(frmssstable.txtE28) / 2
    
'-// SSS Code 29
   Case Is >= 14750
    TxtSSS = Val(frmssstable.txtE29) / 2
    
End Select

'<<//--------------------------------------------------------------------------///-

'-//Case for PH SHare (divide into 2 to get semi-monthly share)
frmPhilHealth.Hide

Select Case Val(TxtBasicPay)

'-// No PH Code  value
   Case Is <= 3999.99
       TxtPhilHealth = 0
       
'-//PH Code1
   Case 4000 To 4999.99
      TxtPhilHealth = Val(frmPhilHealth.txtE1) / 2
    
   '-//PH Code2
   Case 5000 To 5999.99
      TxtPhilHealth = Val(frmPhilHealth.txtE2) / 2
    
 '-//PH Code3
   Case 6000 To 6999.99
       TxtPhilHealth = Val(frmPhilHealth.txtE3) / 2
    
 '-//PH Code4
   Case 7000 To 7999.99
       TxtPhilHealth = Val(frmPhilHealth.txtE4) / 2
    
  '-//PH Code5
   Case 8000 To 8999.99
      TxtPhilHealth = Val(frmPhilHealth.txtE5) / 2
     
 '-//PH Code6
   Case 9000 To 9999.99
       TxtPhilHealth = Val(frmPhilHealth.txtE6) / 2
     
 '-//PH Code7
   Case 10000 To 10999.99
       TxtPhilHealth = Val(frmPhilHealth.txtE7) / 2
     
 '-//PH Code8
   Case 11000 To 11999.99
      TxtPhilHealth = Val(frmPhilHealth.txtE8) / 2
    
  '-//PH Code9
   Case 12000 To 12999.99
      TxtPhilHealth = Val(frmPhilHealth.txtE9) / 2
    
  '-//PH Code10
   Case 13000 To 13999.99
      TxtPhilHealth = Val(frmPhilHealth.txtE10) / 2
    
  '-//PH Code11
   Case 14000 To 14999.99
      TxtPhilHealth = Val(frmPhilHealth.txtE11) / 2
    
   '-//PH Code12
   Case 15000 To 15999.99
      TxtPhilHealth = Val(frmPhilHealth.txtE12) / 2
      
   '-//PH Code13
   Case 16000 To 16999.99
      TxtPhilHealth = Val(frmPhilHealth.txtE13) / 2
  
  '-//PH Code14
   Case 17000 To 17999.99
      TxtPhilHealth = Val(frmPhilHealth.txtE14) / 2
 
  '-//PH Code15
   Case 18000 To 18999.99
      TxtPhilHealth = Val(frmPhilHealth.txtE15) / 2
 
    '-//PH Code16
   Case 19000 To 19999.99
      TxtPhilHealth = Val(frmPhilHealth.txtE16) / 2
 
    '-//PH Code17
   Case Is >= 20000
      TxtPhilHealth = Val(frmPhilHealth.txtE17) / 2
      
   End Select
   
   
End Sub

Sub disable_nav_button()

CmdFirst.Enabled = False
CmdNext.Enabled = False
CmdLast.Enabled = False
CmdPrevious.Enabled = False

End Sub

Sub display_records()

On Error Resume Next

With adoEmp

    TxtEcode.Text = !ECODE & " " '-// To avoid Null value
    ttfname.Text = !fname & " "
    TxtSname.Text = !sname & " "
    txtmi.Text = !mi & " "
    TxtEage.Text = !eage & " "
   
    '-// retrieve info. (gender)

    If !gender = "m" Then
      OptionMale.Value = True
       
    Else
      OptionFemale.Value = True
    End If
    
    CmbCivilStat = !civil_status & " "
    DTbirthdate.Value = !birthdate
    TxtEaddress.Text = !eaddress & " "
    TxtPhoneNo.Text = !phone_no & " "
    TxtEadd.Text = !email_add & " "
    TxtSemi.Text = !semi_monthly & " "
    TxtBasicPay.Text = !basicpay & " "
    CmbPosition.Text = !posdesc & " "
    TxtSyears.Text = !syears & " "
    CmbDepartment.Text = !depdesc & " "
    CmbTaxStat.Text = !taxheadercode & " "
    cmbrice.Text = !rice_all_code & " "
    cmbliving.Text = !living_all_code & " "
    Text5.Text = !rank & " "
    Text3.Text = !living_value & " "
    Text4.Text = !rice_value & " "
    TxtSSS.Text = !SSSPremium & " "
    TxtPhilHealth = !PHHealthValue & " "
    Text2.Text = !Pagibig & " "
    Text1.Text = !withTaxvalue & " "
    
End With

End Sub

