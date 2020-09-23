VERSION 5.00
Begin VB.Form frmssstable 
   BackColor       =   &H000000C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SSS Table"
   ClientHeight    =   8310
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4335
   ForeColor       =   &H000000C0&
   Icon            =   "frmssstable.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8310
   ScaleWidth      =   4335
   Begin VB.TextBox Text88 
      Height          =   495
      Left            =   4800
      TabIndex        =   97
      Top             =   3600
      Width           =   1215
   End
   Begin VB.TextBox txtE28 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2880
      TabIndex        =   96
      Text            =   "483.3"
      Top             =   6960
      Width           =   1215
   End
   Begin VB.TextBox txtE29 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2880
      TabIndex        =   95
      Text            =   "500"
      Top             =   7200
      Width           =   1215
   End
   Begin VB.TextBox txt29 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   94
      Text            =   "14750"
      Top             =   7200
      Width           =   1215
   End
   Begin VB.TextBox txt1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   93
      Text            =   "1000"
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox Text83 
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
      TabIndex        =   92
      Text            =   "29"
      Top             =   7200
      Width           =   735
   End
   Begin VB.TextBox Text82 
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
      TabIndex        =   91
      Text            =   "28"
      Top             =   6960
      Width           =   735
   End
   Begin VB.TextBox txtE27 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2880
      TabIndex        =   89
      Text            =   "466.7"
      Top             =   6720
      Width           =   1215
   End
   Begin VB.TextBox txtE26 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2880
      TabIndex        =   88
      Text            =   "450"
      Top             =   6480
      Width           =   1215
   End
   Begin VB.TextBox txtE25 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2880
      TabIndex        =   87
      Text            =   "433.3"
      Top             =   6240
      Width           =   1215
   End
   Begin VB.TextBox txtE24 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2880
      TabIndex        =   86
      Text            =   "416.7"
      Top             =   6000
      Width           =   1215
   End
   Begin VB.TextBox txtE23 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2880
      TabIndex        =   85
      Text            =   "400"
      Top             =   5760
      Width           =   1215
   End
   Begin VB.TextBox txtE22 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2880
      TabIndex        =   84
      Text            =   "383.3"
      Top             =   5520
      Width           =   1215
   End
   Begin VB.TextBox txtE21 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2880
      TabIndex        =   83
      Text            =   "366.7"
      Top             =   5280
      Width           =   1215
   End
   Begin VB.TextBox txtE20 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2880
      TabIndex        =   82
      Text            =   "350"
      Top             =   5040
      Width           =   1215
   End
   Begin VB.TextBox txtE19 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2880
      TabIndex        =   81
      Text            =   "333.3"
      Top             =   4800
      Width           =   1215
   End
   Begin VB.TextBox txtE18 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2880
      TabIndex        =   80
      Text            =   "316.7"
      Top             =   4560
      Width           =   1215
   End
   Begin VB.TextBox txtE17 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2880
      TabIndex        =   79
      Text            =   "300"
      Top             =   4320
      Width           =   1215
   End
   Begin VB.TextBox txtE16 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2880
      TabIndex        =   78
      Text            =   "283.3"
      Top             =   4080
      Width           =   1215
   End
   Begin VB.TextBox txtE15 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2880
      TabIndex        =   77
      Text            =   "266.7"
      Top             =   3840
      Width           =   1215
   End
   Begin VB.TextBox txtE14 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2880
      TabIndex        =   76
      Text            =   "250"
      Top             =   3600
      Width           =   1215
   End
   Begin VB.TextBox txtE13 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2880
      TabIndex        =   75
      Text            =   "233.3"
      Top             =   3360
      Width           =   1215
   End
   Begin VB.TextBox txtE12 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2880
      TabIndex        =   74
      Text            =   "216.7"
      Top             =   3120
      Width           =   1215
   End
   Begin VB.TextBox txtE11 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2880
      TabIndex        =   73
      Text            =   "200"
      Top             =   2880
      Width           =   1215
   End
   Begin VB.TextBox txtE10 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2880
      TabIndex        =   72
      Text            =   "163.3"
      Top             =   2640
      Width           =   1215
   End
   Begin VB.TextBox txtE9 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2880
      TabIndex        =   71
      Text            =   "166.7"
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox txtE8 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2880
      TabIndex        =   70
      Text            =   "150"
      Top             =   2160
      Width           =   1215
   End
   Begin VB.TextBox txtE7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2880
      TabIndex        =   69
      Text            =   "133.3"
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox txtE6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2880
      TabIndex        =   68
      Text            =   "116.7"
      Top             =   1680
      Width           =   1215
   End
   Begin VB.TextBox txtE5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2880
      TabIndex        =   67
      Text            =   "100"
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox txtE4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2880
      TabIndex        =   66
      Text            =   "83.30"
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox txtE3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2880
      TabIndex        =   65
      Text            =   "66.7"
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox txtE2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2880
      TabIndex        =   64
      Text            =   "50"
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox txtE1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2880
      TabIndex        =   63
      Text            =   "33.3"
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox txt28 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   60
      Text            =   "14250"
      Top             =   6960
      Width           =   1215
   End
   Begin VB.TextBox txt27 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   59
      Text            =   "13750"
      Top             =   6720
      Width           =   1215
   End
   Begin VB.TextBox txt26 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   58
      Text            =   "13250"
      Top             =   6480
      Width           =   1215
   End
   Begin VB.TextBox txt25 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   57
      Text            =   "12750"
      Top             =   6240
      Width           =   1215
   End
   Begin VB.TextBox txt24 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   56
      Text            =   "12250"
      Top             =   6000
      Width           =   1215
   End
   Begin VB.TextBox txt23 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   55
      Text            =   "11750"
      Top             =   5760
      Width           =   1215
   End
   Begin VB.TextBox txt22 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   54
      Text            =   "11250"
      Top             =   5520
      Width           =   1215
   End
   Begin VB.TextBox txt21 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   53
      Text            =   "10750"
      Top             =   5280
      Width           =   1215
   End
   Begin VB.TextBox txt20 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   52
      Text            =   "10250"
      Top             =   5040
      Width           =   1215
   End
   Begin VB.TextBox txt19 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   51
      Text            =   "9750"
      Top             =   4800
      Width           =   1215
   End
   Begin VB.TextBox txt18 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   50
      Text            =   "9250"
      Top             =   4560
      Width           =   1215
   End
   Begin VB.TextBox txt17 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   49
      Text            =   "8750"
      Top             =   4320
      Width           =   1215
   End
   Begin VB.TextBox txt16 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   48
      Text            =   "8250"
      Top             =   4080
      Width           =   1215
   End
   Begin VB.TextBox txt15 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   47
      Text            =   "7750"
      Top             =   3840
      Width           =   1215
   End
   Begin VB.TextBox txt14 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   46
      Text            =   "7250"
      Top             =   3600
      Width           =   1215
   End
   Begin VB.TextBox txt13 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   45
      Text            =   "6750"
      Top             =   3360
      Width           =   1215
   End
   Begin VB.TextBox txt12 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   44
      Text            =   "6250"
      Top             =   3120
      Width           =   1215
   End
   Begin VB.TextBox txt11 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   43
      Text            =   "5750"
      Top             =   2880
      Width           =   1215
   End
   Begin VB.TextBox txt10 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   42
      Text            =   "5250"
      Top             =   2640
      Width           =   1215
   End
   Begin VB.TextBox txt9 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   41
      Text            =   "4750"
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox txt8 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   40
      Text            =   "4250"
      Top             =   2160
      Width           =   1215
   End
   Begin VB.TextBox txt7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   39
      Text            =   "3750"
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox txt6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   38
      Text            =   "3250"
      Top             =   1680
      Width           =   1215
   End
   Begin VB.TextBox txt5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   37
      Text            =   "2750"
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox txt4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   36
      Text            =   "2250"
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox txt3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   35
      Text            =   "1750"
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox txt2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   34
      Text            =   "1250"
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton CmdEdit 
      Caption         =   "&Edit"
      Height          =   375
      Left            =   120
      TabIndex        =   32
      Top             =   7800
      Width           =   1095
   End
   Begin VB.CommandButton CmdUpdate 
      Caption         =   "&Update"
      Height          =   375
      Left            =   1320
      TabIndex        =   31
      Top             =   7800
      Width           =   1335
   End
   Begin VB.CommandButton CmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   2760
      TabIndex        =   30
      Top             =   7800
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Height          =   25
      Left            =   120
      TabIndex        =   28
      Top             =   7680
      Width           =   4095
   End
   Begin VB.TextBox Text29 
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
      TabIndex        =   27
      Text            =   "27"
      Top             =   6720
      Width           =   735
   End
   Begin VB.TextBox Text28 
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
      TabIndex        =   26
      Text            =   "26"
      Top             =   6480
      Width           =   735
   End
   Begin VB.TextBox Text27 
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
      TabIndex        =   25
      Text            =   "25"
      Top             =   6240
      Width           =   735
   End
   Begin VB.TextBox Text26 
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
      TabIndex        =   24
      Text            =   "24"
      Top             =   6000
      Width           =   735
   End
   Begin VB.TextBox Text25 
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
      TabIndex        =   23
      Text            =   "23"
      Top             =   5760
      Width           =   735
   End
   Begin VB.TextBox Text24 
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
      TabIndex        =   22
      Text            =   "22"
      Top             =   5520
      Width           =   735
   End
   Begin VB.TextBox Text23 
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
      TabIndex        =   21
      Text            =   "21"
      Top             =   5280
      Width           =   735
   End
   Begin VB.TextBox Text22 
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
      TabIndex        =   20
      Text            =   "20"
      Top             =   5040
      Width           =   735
   End
   Begin VB.TextBox Text21 
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
      TabIndex        =   19
      Text            =   "19"
      Top             =   4800
      Width           =   735
   End
   Begin VB.TextBox Text20 
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
      TabIndex        =   18
      Text            =   "18"
      Top             =   4560
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
      TabIndex        =   17
      Text            =   "17"
      Top             =   4320
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
      TabIndex        =   16
      Text            =   "16"
      Top             =   4080
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
      TabIndex        =   15
      Text            =   "15"
      Top             =   3840
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
      TabIndex        =   14
      Text            =   "14"
      Top             =   3600
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
      TabIndex        =   13
      Text            =   "13"
      Top             =   3360
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
      TabIndex        =   12
      Text            =   "12"
      Top             =   3120
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
      TabIndex        =   11
      Text            =   "11"
      Top             =   2880
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
      TabIndex        =   10
      Text            =   "10"
      Top             =   2640
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
      TabIndex        =   9
      Text            =   "9"
      Top             =   2400
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
      TabIndex        =   8
      Text            =   "8"
      Top             =   2160
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
      TabIndex        =   7
      Text            =   "7"
      Top             =   1920
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
      TabIndex        =   6
      Text            =   "6"
      Top             =   1680
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
      TabIndex        =   5
      Text            =   "5"
      Top             =   1440
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
      TabIndex        =   4
      Text            =   "4"
      Top             =   1200
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
      TabIndex        =   3
      Text            =   "3"
      Top             =   960
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
      TabIndex        =   2
      Text            =   "2"
      Top             =   720
      Width           =   735
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
      TabIndex        =   1
      Text            =   "1"
      Top             =   480
      Width           =   735
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
      TabIndex        =   90
      Top             =   120
      Width           =   1455
   End
   Begin VB.Shape Shape3 
      BorderWidth     =   3
      Height          =   7215
      Left            =   2760
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0E0FF&
      Height          =   7215
      Left            =   2760
      TabIndex        =   62
      Top             =   360
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
      TabIndex        =   61
      Top             =   120
      Width           =   1215
   End
   Begin VB.Shape Shape2 
      BorderWidth     =   3
      Height          =   7215
      Left            =   1200
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0E0FF&
      Height          =   7215
      Left            =   1200
      TabIndex        =   33
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "SSS Code"
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
      Left            =   120
      TabIndex        =   29
      Top             =   105
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   3
      Height          =   7215
      Left            =   120
      Top             =   360
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0E0FF&
      Height          =   7215
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   975
   End
End
Attribute VB_Name = "frmssstable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdClose_Click()
Unload Me
End Sub

Private Sub CmdEdit_Click()
unlock_txt
txtE1.SetFocus
End Sub

Sub lock_txt()

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
txtE18.Locked = True
txtE19.Locked = True
txtE20.Locked = True
txtE21.Locked = True
txtE22.Locked = True
txtE23.Locked = True
txtE24.Locked = True
txtE25.Locked = True
txtE26.Locked = True
txtE27.Locked = True
txtE28.Locked = True
txtE29.Locked = True

End Sub

Sub unlock_txt()

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
txtE18.Locked = False
txtE19.Locked = False
txtE20.Locked = False
txtE21.Locked = False
txtE22.Locked = False
txtE23.Locked = False
txtE24.Locked = False
txtE25.Locked = False
txtE26.Locked = False
txtE27.Locked = False
txtE28.Locked = False
txtE29.Locked = False

End Sub

Private Sub CmdUpdate_Click()

Call lock_txt
Modregistry.saving_SSS_Table_rec
MsgBox "Record Succesfully Updated!", vbInformation, "Success..."

End Sub

Private Sub Form_Load()

 Text88.Text = GetSetting("IsStore", "setting", "value")
 
 If Text88.Text = "Yes" Then
    
    '-// Salary Base GetValue
    txt1.Text = GetSetting("SSS Table", "cost", "SSS Code1")
    txt2.Text = GetSetting("SSS Table", "cost", "SSS Code2")
    txt3.Text = GetSetting("SSS Table", "cost", "SSS Code3")
    txt4.Text = GetSetting("SSS Table", "cost", "SSS Code4")
    txt5.Text = GetSetting("SSS Table", "cost", "SSS Code5")
    txt6.Text = GetSetting("SSS Table", "cost", "SSS Code6")
    txt7.Text = GetSetting("SSS Table", "cost", "SSS Code7")
    txt8.Text = GetSetting("SSS Table", "cost", "SSS Code8")
    txt9.Text = GetSetting("SSS Table", "cost", "SSS Code9")
    txt10.Text = GetSetting("SSS Table", "cost", "SSS Code10")
    txt11.Text = GetSetting("SSS Table", "cost", "SSS Code11")
    txt12.Text = GetSetting("SSS Table", "cost", "SSS Code12")
    txt13.Text = GetSetting("SSS Table", "cost", "SSS Code13")
    txt14.Text = GetSetting("SSS Table", "cost", "SSS Code14")
    txt15.Text = GetSetting("SSS Table", "cost", "SSS Code15")
    txt16.Text = GetSetting("SSS Table", "cost", "SSS Code16")
    txt17.Text = GetSetting("SSS Table", "cost", "SSS Code17")
    txt18.Text = GetSetting("SSS Table", "cost", "SSS Code18")
    txt19.Text = GetSetting("SSS Table", "cost", "SSS Code19")
    txt20.Text = GetSetting("SSS Table", "cost", "SSS Code20")
    txt21.Text = GetSetting("SSS Table", "cost", "SSS Code21")
    txt22.Text = GetSetting("SSS Table", "cost", "SSS Code22")
    txt23.Text = GetSetting("SSS Table", "cost", "SSS Code23")
    txt24.Text = GetSetting("SSS Table", "cost", "SSS Code24")
    txt25.Text = GetSetting("SSS Table", "cost", "SSS Code25")
    txt26.Text = GetSetting("SSS Table", "cost", "SSS Code26")
    txt27.Text = GetSetting("SSS Table", "cost", "SSS Code27")
    txt28.Text = GetSetting("SSS Table", "cost", "SSS Code28")
    txt29.Text = GetSetting("SSS Table", "cost", "SSS Code29")
    
    
    txtE1 = GetSetting("SSS Table", "EShare", "SSS Share1")
    txtE2 = GetSetting("SSS Table", "EShare", "SSS Share2")
    txtE3 = GetSetting("SSS Table", "EShare", "SSS Share3")
    txtE4 = GetSetting("SSS Table", "EShare", "SSS Share4")
    txtE5 = GetSetting("SSS Table", "EShare", "SSS Share5")
    txtE6 = GetSetting("SSS Table", "EShare", "SSS Share6")
    txtE7 = GetSetting("SSS Table", "EShare", "SSS Share7")
    txtE8 = GetSetting("SSS Table", "EShare", "SSS Share8")
    txtE9 = GetSetting("SSS Table", "EShare", "SSS Share9")
    txtE10 = GetSetting("SSS Table", "EShare", "SSS Share10")
    txtE11 = GetSetting("SSS Table", "EShare", "SSS Share11")
    txtE12 = GetSetting("SSS Table", "EShare", "SSS Share12")
    txtE13 = GetSetting("SSS Table", "EShare", "SSS Share13")
    txtE14 = GetSetting("SSS Table", "EShare", "SSS Share14")
    txtE15 = GetSetting("SSS Table", "EShare", "SSS Share15")
    txtE16 = GetSetting("SSS Table", "EShare", "SSS Share16")
    txtE17 = GetSetting("SSS Table", "EShare", "SSS Share17")
    txtE18 = GetSetting("SSS Table", "EShare", "SSS Share18")
    txtE19 = GetSetting("SSS Table", "EShare", "SSS Share19")
    txtE20 = GetSetting("SSS Table", "EShare", "SSS Share20")
    txtE21 = GetSetting("SSS Table", "EShare", "SSS Share21")
    txtE22 = GetSetting("SSS Table", "EShare", "SSS Share22")
    txtE23 = GetSetting("SSS Table", "EShare", "SSS Share23")
    txtE24 = GetSetting("SSS Table", "EShare", "SSS Share24")
    txtE25 = GetSetting("SSS Table", "EShare", "SSS Share25")
    txtE26 = GetSetting("SSS Table", "EShare", "SSS Share26")
    txtE27 = GetSetting("SSS Table", "EShare", "SSS Share27")
    txtE28 = GetSetting("SSS Table", "EShare", "SSS Share28")
    txtE29 = GetSetting("SSS Table", "EShare", "SSS Share29")
 
     
Else

    Call SaveSetting("IsStore", "setting", "value", "Yes")
    Call Modregistry.saving_SSS_Table_rec
    
End If


End Sub

