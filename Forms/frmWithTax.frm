VERSION 5.00
Begin VB.Form frmWithTax 
   BackColor       =   &H000000C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "WithHolding Tax Table"
   ClientHeight    =   7080
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8880
   Icon            =   "frmWithTax.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7080
   ScaleWidth      =   8880
   Begin VB.TextBox Text1 
      Height          =   372
      Left            =   2400
      TabIndex        =   157
      Top             =   7800
      Width           =   492
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   5640
      TabIndex        =   156
      Top             =   6600
      Width           =   972
   End
   Begin VB.CommandButton cmdupdate 
      Caption         =   "&Update"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   6720
      TabIndex        =   155
      Top             =   6600
      Width           =   972
   End
   Begin VB.CommandButton cmdclose 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   7800
      TabIndex        =   154
      Top             =   6600
      Width           =   972
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H000000C0&
      Height          =   132
      Left            =   120
      TabIndex        =   153
      Top             =   6360
      Width           =   8652
   End
   Begin VB.TextBox txt00pme2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   148
      Text            =   "48.0"
      Top             =   5400
      Width           =   615
   End
   Begin VB.TextBox txt00pme3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   147
      Text            =   "56.0"
      Top             =   5640
      Width           =   615
   End
   Begin VB.TextBox txt00pme4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   146
      Text            =   "64.0"
      Top             =   5880
      Width           =   615
   End
   Begin VB.TextBox txt00pme1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   145
      Text            =   "40.0"
      Top             =   5160
      Width           =   615
   End
   Begin VB.TextBox txtme28 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7800
      Locked          =   -1  'True
      TabIndex        =   144
      Text            =   "22833"
      Top             =   5400
      Width           =   615
   End
   Begin VB.TextBox txtme38 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7800
      Locked          =   -1  'True
      TabIndex        =   143
      Text            =   "23167"
      Top             =   5640
      Width           =   615
   End
   Begin VB.TextBox txtme48 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7800
      Locked          =   -1  'True
      TabIndex        =   142
      Text            =   "23500"
      Top             =   5880
      Width           =   615
   End
   Begin VB.TextBox txtme18 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7800
      Locked          =   -1  'True
      TabIndex        =   141
      Text            =   "22500"
      Top             =   5160
      Width           =   615
   End
   Begin VB.TextBox txtme27 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6960
      Locked          =   -1  'True
      TabIndex        =   140
      Text            =   "12417"
      Top             =   5400
      Width           =   615
   End
   Begin VB.TextBox txtme37 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6960
      Locked          =   -1  'True
      TabIndex        =   139
      Text            =   "12750"
      Top             =   5640
      Width           =   615
   End
   Begin VB.TextBox txtme47 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6960
      Locked          =   -1  'True
      TabIndex        =   138
      Text            =   "13083"
      Top             =   5880
      Width           =   615
   End
   Begin VB.TextBox txtme17 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6960
      Locked          =   -1  'True
      TabIndex        =   137
      Text            =   "12083"
      Top             =   5160
      Width           =   615
   End
   Begin VB.TextBox txtme26 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6120
      Locked          =   -1  'True
      TabIndex        =   136
      Text            =   "7833"
      Top             =   5400
      Width           =   495
   End
   Begin VB.TextBox txtme36 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6120
      Locked          =   -1  'True
      TabIndex        =   135
      Text            =   "8167"
      Top             =   5640
      Width           =   495
   End
   Begin VB.TextBox txtme46 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6120
      Locked          =   -1  'True
      TabIndex        =   134
      Text            =   "8500"
      Top             =   5880
      Width           =   495
   End
   Begin VB.TextBox txtme16 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6120
      Locked          =   -1  'True
      TabIndex        =   133
      Text            =   "7500"
      Top             =   5160
      Width           =   495
   End
   Begin VB.TextBox txtme25 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   132
      Text            =   "4917"
      Top             =   5400
      Width           =   495
   End
   Begin VB.TextBox txtme35 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   131
      Text            =   "5250"
      Top             =   5640
      Width           =   495
   End
   Begin VB.TextBox txtme45 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   130
      Text            =   "5583"
      Top             =   5880
      Width           =   495
   End
   Begin VB.TextBox txtme15 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   129
      Text            =   "4583"
      Top             =   5160
      Width           =   495
   End
   Begin VB.TextBox txtme24 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4320
      Locked          =   -1  'True
      TabIndex        =   128
      Text            =   "3250"
      Top             =   5400
      Width           =   495
   End
   Begin VB.TextBox txtme34 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4320
      Locked          =   -1  'True
      TabIndex        =   127
      Text            =   "3583"
      Top             =   5640
      Width           =   495
   End
   Begin VB.TextBox txtme44 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4320
      Locked          =   -1  'True
      TabIndex        =   126
      Text            =   "3917"
      Top             =   5880
      Width           =   495
   End
   Begin VB.TextBox txtme14 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4320
      Locked          =   -1  'True
      TabIndex        =   125
      Text            =   "2917"
      Top             =   5160
      Width           =   495
   End
   Begin VB.TextBox txtme23 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3480
      Locked          =   -1  'True
      TabIndex        =   124
      Text            =   "2417"
      Top             =   5400
      Width           =   495
   End
   Begin VB.TextBox txtme33 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3480
      Locked          =   -1  'True
      TabIndex        =   123
      Text            =   "2750"
      Top             =   5640
      Width           =   495
   End
   Begin VB.TextBox txtme43 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3480
      Locked          =   -1  'True
      TabIndex        =   122
      Text            =   "3083"
      Top             =   5880
      Width           =   495
   End
   Begin VB.TextBox txtme13 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3480
      Locked          =   -1  'True
      TabIndex        =   121
      Text            =   "2083"
      Top             =   5160
      Width           =   495
   End
   Begin VB.TextBox txtme22 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   120
      Text            =   "2000"
      Top             =   5400
      Width           =   495
   End
   Begin VB.TextBox txtme32 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   119
      Text            =   "2333"
      Top             =   5640
      Width           =   495
   End
   Begin VB.TextBox txtme42 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   118
      Text            =   "2667"
      Top             =   5880
      Width           =   495
   End
   Begin VB.TextBox txtme12 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   117
      Text            =   "1667"
      Top             =   5160
      Width           =   495
   End
   Begin VB.TextBox txtme21 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   116
      Text            =   "1"
      Top             =   5400
      Width           =   255
   End
   Begin VB.TextBox txtme31 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   115
      Text            =   "1"
      Top             =   5640
      Width           =   255
   End
   Begin VB.TextBox txtme41 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   114
      Text            =   "1"
      Top             =   5880
      Width           =   255
   End
   Begin VB.TextBox txtme11 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   113
      Text            =   "1"
      Top             =   5160
      Width           =   255
   End
   Begin VB.TextBox txthf11 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   105
      Text            =   "1"
      Top             =   3480
      Width           =   255
   End
   Begin VB.TextBox txthf41 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   104
      Text            =   "1"
      Top             =   4200
      Width           =   255
   End
   Begin VB.TextBox txthf31 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   103
      Text            =   "1"
      Top             =   3960
      Width           =   255
   End
   Begin VB.TextBox txthf21 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   102
      Text            =   "1"
      Top             =   3720
      Width           =   255
   End
   Begin VB.TextBox txthf12 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   101
      Text            =   "1375"
      Top             =   3480
      Width           =   495
   End
   Begin VB.TextBox txthf42 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   100
      Text            =   "2375"
      Top             =   4200
      Width           =   495
   End
   Begin VB.TextBox txthf32 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   99
      Text            =   "2042"
      Top             =   3960
      Width           =   495
   End
   Begin VB.TextBox txthf22 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   98
      Text            =   "1708"
      Top             =   3720
      Width           =   495
   End
   Begin VB.TextBox txthf13 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3480
      Locked          =   -1  'True
      TabIndex        =   97
      Text            =   "1792"
      Top             =   3480
      Width           =   495
   End
   Begin VB.TextBox txthf43 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3480
      Locked          =   -1  'True
      TabIndex        =   96
      Text            =   "2792"
      Top             =   4200
      Width           =   495
   End
   Begin VB.TextBox txthf33 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3480
      Locked          =   -1  'True
      TabIndex        =   95
      Text            =   "2458"
      Top             =   3960
      Width           =   495
   End
   Begin VB.TextBox txthf23 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3480
      Locked          =   -1  'True
      TabIndex        =   94
      Text            =   "2125"
      Top             =   3720
      Width           =   495
   End
   Begin VB.TextBox txthf14 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4320
      Locked          =   -1  'True
      TabIndex        =   93
      Text            =   "2625"
      Top             =   3480
      Width           =   495
   End
   Begin VB.TextBox txthf44 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4320
      Locked          =   -1  'True
      TabIndex        =   92
      Text            =   "3625"
      Top             =   4200
      Width           =   495
   End
   Begin VB.TextBox txthf34 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4320
      Locked          =   -1  'True
      TabIndex        =   91
      Text            =   "3292"
      Top             =   3960
      Width           =   495
   End
   Begin VB.TextBox txthf24 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4320
      Locked          =   -1  'True
      TabIndex        =   90
      Text            =   "2958"
      Top             =   3720
      Width           =   495
   End
   Begin VB.TextBox txthf15 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   89
      Text            =   "4292"
      Top             =   3480
      Width           =   495
   End
   Begin VB.TextBox txthf45 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   88
      Text            =   "5292"
      Top             =   4200
      Width           =   495
   End
   Begin VB.TextBox txthf35 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   87
      Text            =   "4958"
      Top             =   3960
      Width           =   495
   End
   Begin VB.TextBox txthf25 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   86
      Text            =   "4625"
      Top             =   3720
      Width           =   495
   End
   Begin VB.TextBox txthf16 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6120
      Locked          =   -1  'True
      TabIndex        =   85
      Text            =   "7208"
      Top             =   3480
      Width           =   495
   End
   Begin VB.TextBox txthf46 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6120
      Locked          =   -1  'True
      TabIndex        =   84
      Text            =   "8208"
      Top             =   4200
      Width           =   495
   End
   Begin VB.TextBox txthf36 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6120
      Locked          =   -1  'True
      TabIndex        =   83
      Text            =   "7875"
      Top             =   3960
      Width           =   495
   End
   Begin VB.TextBox txthf26 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6120
      Locked          =   -1  'True
      TabIndex        =   82
      Text            =   "7542"
      Top             =   3720
      Width           =   495
   End
   Begin VB.TextBox txthf17 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6960
      Locked          =   -1  'True
      TabIndex        =   81
      Text            =   "11792"
      Top             =   3480
      Width           =   615
   End
   Begin VB.TextBox txthf47 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6960
      Locked          =   -1  'True
      TabIndex        =   80
      Text            =   "12792"
      Top             =   4200
      Width           =   615
   End
   Begin VB.TextBox txthf37 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6960
      Locked          =   -1  'True
      TabIndex        =   79
      Text            =   "12458"
      Top             =   3960
      Width           =   615
   End
   Begin VB.TextBox txthf27 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6960
      Locked          =   -1  'True
      TabIndex        =   78
      Text            =   "12125"
      Top             =   3720
      Width           =   615
   End
   Begin VB.TextBox txthf18 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7800
      Locked          =   -1  'True
      TabIndex        =   77
      Text            =   "22208"
      Top             =   3480
      Width           =   615
   End
   Begin VB.TextBox txthf48 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7800
      Locked          =   -1  'True
      TabIndex        =   76
      Text            =   "23208"
      Top             =   4200
      Width           =   615
   End
   Begin VB.TextBox txthf38 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7800
      Locked          =   -1  'True
      TabIndex        =   75
      Text            =   "22875"
      Top             =   3960
      Width           =   615
   End
   Begin VB.TextBox txthf28 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7800
      Locked          =   -1  'True
      TabIndex        =   74
      Text            =   "22542"
      Top             =   3720
      Width           =   615
   End
   Begin VB.TextBox txt00phf1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   73
      Text            =   "33.0"
      Top             =   3480
      Width           =   615
   End
   Begin VB.TextBox txt00phf4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   72
      Text            =   "57.0"
      Top             =   4200
      Width           =   615
   End
   Begin VB.TextBox txt00phf3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   71
      Text            =   "49.0"
      Top             =   3960
      Width           =   615
   End
   Begin VB.TextBox txt00phf2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   70
      Text            =   "41.0"
      Top             =   3720
      Width           =   615
   End
   Begin VB.TextBox txtz1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   63
      Text            =   "0"
      Top             =   1680
      Width           =   255
   End
   Begin VB.TextBox txtme1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   62
      Text            =   "1"
      Top             =   2400
      Width           =   255
   End
   Begin VB.TextBox txthf1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   61
      Text            =   "1"
      Top             =   2160
      Width           =   255
   End
   Begin VB.TextBox txts1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   60
      Text            =   "1"
      Top             =   1920
      Width           =   255
   End
   Begin VB.TextBox txtz2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   59
      Text            =   "0"
      Top             =   1680
      Width           =   495
   End
   Begin VB.TextBox txtme2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   58
      Text            =   "1333"
      Top             =   2400
      Width           =   495
   End
   Begin VB.TextBox txthf2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   57
      Text            =   "1042"
      Top             =   2160
      Width           =   495
   End
   Begin VB.TextBox txts2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   56
      Text            =   "833"
      Top             =   1920
      Width           =   495
   End
   Begin VB.TextBox txtz3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3480
      Locked          =   -1  'True
      TabIndex        =   55
      Text            =   "417"
      Top             =   1680
      Width           =   495
   End
   Begin VB.TextBox txtme3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3480
      Locked          =   -1  'True
      TabIndex        =   54
      Text            =   "1750"
      Top             =   2400
      Width           =   495
   End
   Begin VB.TextBox txthf3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3480
      Locked          =   -1  'True
      TabIndex        =   53
      Text            =   "1458"
      Top             =   2160
      Width           =   495
   End
   Begin VB.TextBox txts3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3480
      Locked          =   -1  'True
      TabIndex        =   52
      Text            =   "1250"
      Top             =   1920
      Width           =   495
   End
   Begin VB.TextBox txtz4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4320
      Locked          =   -1  'True
      TabIndex        =   51
      Text            =   "1250"
      Top             =   1680
      Width           =   495
   End
   Begin VB.TextBox txtme4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4320
      Locked          =   -1  'True
      TabIndex        =   50
      Text            =   "2583"
      Top             =   2400
      Width           =   495
   End
   Begin VB.TextBox txthf4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4320
      Locked          =   -1  'True
      TabIndex        =   49
      Text            =   "2292"
      Top             =   2160
      Width           =   495
   End
   Begin VB.TextBox txts4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4320
      Locked          =   -1  'True
      TabIndex        =   48
      Text            =   "2083"
      Top             =   1920
      Width           =   495
   End
   Begin VB.TextBox txtz5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   47
      Text            =   "2917"
      Top             =   1680
      Width           =   495
   End
   Begin VB.TextBox txtme5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   46
      Text            =   "4250"
      Top             =   2400
      Width           =   495
   End
   Begin VB.TextBox txthf5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   45
      Text            =   "3958"
      Top             =   2160
      Width           =   495
   End
   Begin VB.TextBox txts5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   44
      Text            =   "3750"
      Top             =   1920
      Width           =   495
   End
   Begin VB.TextBox txtz6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6120
      Locked          =   -1  'True
      TabIndex        =   43
      Text            =   "5833"
      Top             =   1680
      Width           =   495
   End
   Begin VB.TextBox txtme6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6120
      Locked          =   -1  'True
      TabIndex        =   42
      Text            =   "7167"
      Top             =   2400
      Width           =   495
   End
   Begin VB.TextBox txthf6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6120
      Locked          =   -1  'True
      TabIndex        =   41
      Text            =   "6875"
      Top             =   2160
      Width           =   495
   End
   Begin VB.TextBox txts6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6120
      Locked          =   -1  'True
      TabIndex        =   40
      Text            =   "6667"
      Top             =   1920
      Width           =   495
   End
   Begin VB.TextBox txtz7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6960
      Locked          =   -1  'True
      TabIndex        =   39
      Text            =   "10417"
      Top             =   1680
      Width           =   615
   End
   Begin VB.TextBox txtme7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6960
      Locked          =   -1  'True
      TabIndex        =   38
      Text            =   "11750"
      Top             =   2400
      Width           =   615
   End
   Begin VB.TextBox txthf7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6960
      Locked          =   -1  'True
      TabIndex        =   37
      Text            =   "11458"
      Top             =   2160
      Width           =   615
   End
   Begin VB.TextBox txts7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6960
      Locked          =   -1  'True
      TabIndex        =   36
      Text            =   "11250"
      Top             =   1920
      Width           =   615
   End
   Begin VB.TextBox txtz8 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7800
      Locked          =   -1  'True
      TabIndex        =   35
      Text            =   "20833"
      Top             =   1680
      Width           =   615
   End
   Begin VB.TextBox txtme8 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7800
      Locked          =   -1  'True
      TabIndex        =   34
      Text            =   "22167"
      Top             =   2400
      Width           =   615
   End
   Begin VB.TextBox txthf8 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7800
      Locked          =   -1  'True
      TabIndex        =   33
      Text            =   "21875"
      Top             =   2160
      Width           =   615
   End
   Begin VB.TextBox txts8 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7800
      Locked          =   -1  'True
      TabIndex        =   32
      Text            =   "21667"
      Top             =   1920
      Width           =   615
   End
   Begin VB.TextBox txt00pz 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   31
      Text            =   "0.0"
      Top             =   1680
      Width           =   615
   End
   Begin VB.TextBox txt00pme 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   30
      Text            =   "32.0"
      Top             =   2400
      Width           =   615
   End
   Begin VB.TextBox txt00phf 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   29
      Text            =   "25.0"
      Top             =   2160
      Width           =   615
   End
   Begin VB.TextBox txt00ps 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   28
      Text            =   "20.0"
      Top             =   1920
      Width           =   615
   End
   Begin VB.TextBox txtex8 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7800
      Locked          =   -1  'True
      TabIndex        =   15
      Text            =   "5208.33"
      Top             =   480
      Width           =   735
   End
   Begin VB.TextBox txtex7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6960
      Locked          =   -1  'True
      TabIndex        =   14
      Text            =   "2083.33"
      Top             =   480
      Width           =   615
   End
   Begin VB.TextBox txtex6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6120
      Locked          =   -1  'True
      TabIndex        =   13
      Text            =   "937.50"
      Top             =   480
      Width           =   615
   End
   Begin VB.TextBox txtex5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   12
      Text            =   "354.17"
      Top             =   480
      Width           =   615
   End
   Begin VB.TextBox txtex4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4440
      Locked          =   -1  'True
      TabIndex        =   11
      Text            =   "104.17"
      Top             =   480
      Width           =   615
   End
   Begin VB.TextBox txtex3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   10
      Text            =   "20.83"
      Top             =   480
      Width           =   495
   End
   Begin VB.TextBox txtex2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   9
      Text            =   "0.00"
      Top             =   480
      Width           =   495
   End
   Begin VB.TextBox txtex1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   8
      Text            =   "0.00"
      Top             =   480
      Width           =   495
   End
   Begin VB.TextBox txtop8 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7800
      Locked          =   -1  'True
      TabIndex        =   7
      Text            =   "0.32"
      Top             =   840
      Width           =   735
   End
   Begin VB.TextBox txtop7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6960
      Locked          =   -1  'True
      TabIndex        =   6
      Text            =   "0.30"
      Top             =   840
      Width           =   615
   End
   Begin VB.TextBox txtop6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6120
      Locked          =   -1  'True
      TabIndex        =   5
      Text            =   "0.25"
      Top             =   840
      Width           =   615
   End
   Begin VB.TextBox txtop5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   4
      Text            =   "0.20"
      Top             =   840
      Width           =   615
   End
   Begin VB.TextBox txtop4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4440
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "0.15"
      Top             =   840
      Width           =   615
   End
   Begin VB.TextBox txtop3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   2
      Text            =   "0.1"
      Top             =   840
      Width           =   495
   End
   Begin VB.TextBox txtop2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   ".05"
      Top             =   840
      Width           =   495
   End
   Begin VB.TextBox txtop1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   0
      Text            =   "0"
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Label147 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "1. ME1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   240
      TabIndex        =   152
      Top             =   5160
      Width           =   612
   End
   Begin VB.Label Label146 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "2. ME2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   240
      TabIndex        =   151
      Top             =   5400
      Width           =   612
   End
   Begin VB.Label Label145 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "3. ME3"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   240
      TabIndex        =   150
      Top             =   5640
      Width           =   612
   End
   Begin VB.Label Label144 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "4. ME4"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   240
      TabIndex        =   149
      Top             =   5880
      Width           =   612
   End
   Begin VB.Shape Shape3 
      BorderWidth     =   3
      Height          =   1332
      Left            =   120
      Top             =   5040
      Width           =   8652
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "C. Table for married employee with qualified dependent child(ren)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   252
      Left            =   120
      TabIndex        =   112
      Top             =   4800
      Width           =   5652
   End
   Begin VB.Label Label16 
      BackColor       =   &H00C0E0FF&
      Height          =   1332
      Left            =   120
      TabIndex        =   111
      Top             =   5040
      Width           =   8652
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "B. Table for heads of the family with dependent child(ren)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   252
      Left            =   120
      TabIndex        =   110
      Top             =   3000
      Width           =   5652
   End
   Begin VB.Label Label107 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "1. HF1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   240
      TabIndex        =   109
      Top             =   3480
      Width           =   492
   End
   Begin VB.Label Label106 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "2. HF2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   240
      TabIndex        =   108
      Top             =   3720
      Width           =   492
   End
   Begin VB.Label Label105 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "3. HF3"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   240
      TabIndex        =   107
      Top             =   3960
      Width           =   612
   End
   Begin VB.Label Label104 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "4. HF4"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   240
      TabIndex        =   106
      Top             =   4200
      Width           =   492
   End
   Begin VB.Shape Shape2 
      BorderWidth     =   3
      Height          =   1452
      Left            =   120
      Top             =   3240
      Width           =   8652
   End
   Begin VB.Label Label14 
      BackColor       =   &H00C0E0FF&
      Height          =   1452
      Left            =   120
      TabIndex        =   69
      Top             =   3240
      Width           =   8652
   End
   Begin VB.Label me 
      BackStyle       =   0  'Transparent
      Caption         =   "4. ME"
      Height          =   252
      Left            =   360
      TabIndex        =   67
      Top             =   2400
      Width           =   492
   End
   Begin VB.Label hf 
      BackStyle       =   0  'Transparent
      Caption         =   "3. HF"
      Height          =   252
      Left            =   360
      TabIndex        =   66
      Top             =   2160
      Width           =   492
   End
   Begin VB.Label a 
      BackStyle       =   0  'Transparent
      Caption         =   "2. S"
      Height          =   252
      Left            =   360
      TabIndex        =   65
      Top             =   1920
      Width           =   492
   End
   Begin VB.Label z 
      BackStyle       =   0  'Transparent
      Caption         =   "1. Z"
      Height          =   252
      Left            =   360
      TabIndex        =   64
      Top             =   1680
      Width           =   492
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   3
      Height          =   1332
      Left            =   120
      Top             =   1560
      Width           =   8652
   End
   Begin VB.Label Label12 
      BackColor       =   &H00C0E0FF&
      Height          =   1332
      Left            =   120
      TabIndex        =   27
      Top             =   1560
      Width           =   8652
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00C0FFFF&
      X1              =   120
      X2              =   8760
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "OOP"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   252
      Left            =   1200
      TabIndex        =   26
      Top             =   840
      Width           =   852
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Status"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   252
      Left            =   240
      TabIndex        =   25
      Top             =   840
      Width           =   852
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Exemption"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   252
      Left            =   720
      TabIndex        =   24
      Top             =   480
      Width           =   852
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0FFFF&
      BorderStyle     =   3  'Dot
      X1              =   120
      X2              =   8880
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "8"
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Left            =   7800
      TabIndex        =   23
      Top             =   120
      Width           =   732
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "7"
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Left            =   6840
      TabIndex        =   22
      Top             =   120
      Width           =   732
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Left            =   6000
      TabIndex        =   21
      Top             =   120
      Width           =   732
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "5"
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Left            =   5160
      TabIndex        =   20
      Top             =   120
      Width           =   732
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "4"
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Left            =   4320
      TabIndex        =   19
      Top             =   120
      Width           =   732
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Left            =   3480
      TabIndex        =   18
      Top             =   120
      Width           =   732
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Left            =   2760
      TabIndex        =   17
      Top             =   120
      Width           =   732
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Left            =   2160
      TabIndex        =   16
      Top             =   120
      Width           =   612
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "A. Table for Employees without dependent child(ren)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   252
      Left            =   120
      TabIndex        =   68
      Top             =   1344
      Width           =   5652
   End
End
Attribute VB_Name = "frmWithTax"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub lock_ctrl()

txtop1.Locked = True
txtop2.Locked = True
txtop3.Locked = True
txtop4.Locked = True
txtop5.Locked = True
txtop6.Locked = True
txtop7.Locked = True
txtop8.Locked = True

txtex1.Locked = True
txtex2.Locked = True
txtex3.Locked = True
txtex4.Locked = True
txtex5.Locked = True
txtex6.Locked = True
txtex7.Locked = True
txtex8.Locked = True

End Sub

Sub unlock_ctrl()
txtop1.Locked = False
txtop2.Locked = False
txtop3.Locked = False
txtop4.Locked = False
txtop6.Locked = False
txtop7.Locked = False
txtop8.Locked = False

txtex1.Locked = False
txtex2.Locked = False
txtex3.Locked = False
txtex4.Locked = False
txtex5.Locked = False
txtex6.Locked = False
txtex7.Locked = False
txtex8.Locked = False
End Sub

Private Sub CmdClose_Click()
Unload Me
End Sub

Private Sub CmdEdit_Click()
unlock_ctrl
txtex1.SetFocus
End Sub

Private Sub CmdUpdate_Click()

lock_ctrl
Modregistry.saving_withTax_rec
MsgBox "Record Successfully Updated!", vbInformation, "Success..."

End Sub

Private Sub Form_Load()

lock_ctrl
Text1.Text = GetSetting("IsWithTaxInstalled", "setting", "value")
 
If Text1.Text = "Yes" Then
    
 txtex1.Text = GetSetting("WithHolding Tax", "Exemption", "ExemptVal1")
 txtex2.Text = GetSetting("WithHolding Tax", "Exemption", "ExemptVal2")
 txtex3.Text = GetSetting("WithHolding Tax", "Exemption", "ExemptVal3")
 txtex4.Text = GetSetting("WithHolding Tax", "Exemption", "ExemptVal4")
 txtex5.Text = GetSetting("WithHolding Tax", "Exemption", "ExemptVal5")
 txtex6.Text = GetSetting("WithHolding Tax", "Exemption", "ExemptVal6")
 txtex7.Text = GetSetting("WithHolding Tax", "Exemption", "ExemptVal7")
 txtex8.Text = GetSetting("WithHolding Tax", "Exemption", "ExemptVal8")


txtop1.Text = GetSetting("WithHolding Tax", "Status", "StatOOPVal1")
txtop2.Text = GetSetting("WithHolding Tax", "Status", "StatOOPVal2")
txtop3.Text = GetSetting("WithHolding Tax", "Status", "StatOOPVal3")
txtop4.Text = GetSetting("WithHolding Tax", "Status", "StatOOPVal4")
txtop5.Text = GetSetting("WithHolding Tax", "Status", "StatOOPVal5")
txtop6.Text = GetSetting("WithHolding Tax", "Status", "StatOOPVal6")
txtop7.Text = GetSetting("WithHolding Tax", "Status", "StatOOPVal7")
txtop8.Text = GetSetting("WithHolding Tax", "Status", "StatOOPVal8")

Else

    Call SaveSetting("IsWithTaxInstalled", "setting", "value", "Yes")
    Call Modregistry.saving_withTax_rec

End If

End Sub

