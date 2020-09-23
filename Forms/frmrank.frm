VERSION 5.00
Begin VB.Form frmrank 
   BackColor       =   &H000000C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rank"
   ClientHeight    =   3960
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3585
   Icon            =   "frmrank.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   3585
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   2280
      TabIndex        =   17
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton CmdUpdate 
      Caption         =   "&Update"
      Height          =   375
      Left            =   1200
      TabIndex        =   18
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton CmdEdit 
      Caption         =   "&Edit"
      Height          =   375
      Left            =   120
      TabIndex        =   19
      Top             =   3480
      Width           =   1095
   End
   Begin VB.ComboBox Combo1 
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
      Left            =   960
      TabIndex        =   8
      Text            =   "[ Select Position ]"
      Top             =   240
      Width           =   2295
   End
   Begin VB.TextBox Text1 
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
      TabIndex        =   7
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox Text2 
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
      TabIndex        =   6
      Top             =   1080
      Width           =   1215
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
      Left            =   1680
      TabIndex        =   5
      Top             =   1440
      Width           =   1215
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
      Left            =   1680
      TabIndex        =   4
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox Text5 
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
      Top             =   2160
      Width           =   1215
   End
   Begin VB.TextBox Text6 
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
      TabIndex        =   2
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox Text7 
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
      TabIndex        =   1
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Position"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   240
      Width           =   735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Rank 1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   15
      Top             =   720
      Width           =   615
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Rank 2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   14
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Rank 3"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   13
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Rank 4"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   12
      Top             =   1800
      Width           =   615
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Rank 5"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   11
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Rank 6"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   10
      Top             =   2520
      Width           =   615
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Rank 7"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   9
      Top             =   2880
      Width           =   615
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   3
      Height          =   3255
      Left            =   120
      Top             =   120
      Width           =   3375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0E0FF&
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "frmrank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub locked_textbox()

Text1.Locked = True
Text2.Locked = True
Text3.Locked = True
Text4.Locked = True
Text5.Locked = True
Text6.Locked = True
Text7.Locked = True

End Sub

Private Sub unlocked_textbox()

Text1.Locked = False
Text2.Locked = False
Text3.Locked = False
Text4.Locked = False
Text5.Locked = False
Text6.Locked = False
Text7.Locked = False

End Sub

Private Sub CmdClose_Click()
Unload Me
End Sub

Private Sub CmdEdit_Click()
If Rs.RecordCount = 0 Then MsgBox "No record(s) to edit", vbCritical, "ERROR":   Exit Sub
unlocked_textbox
Combo1.SetFocus
End Sub

Private Sub CmdUpdate_Click()
   
 If Combo1.ListCount = 0 Or Combo1.Text = "" Then MsgBox "No record(s) to update!", vbCritical, "ERROR": Exit Sub
 
 If Trim(Text1) = "" Then MsgBox "Invalid Input. Please check it.", vbCritical, "ERROR": Text1.SetFocus: Exit Sub
 If Trim(Text2) = "" Then MsgBox "Invalid Input. Please check it.", vbCritical, "ERROR": Text2.SetFocus: Exit Sub
 If Trim(Text3) = "" Then MsgBox "Invalid Input. Please check it.", vbCritical, "ERROR": Text3.SetFocus: Exit Sub
 If Trim(Text4) = "" Then MsgBox "Invalid Input. Please check it.", vbCritical, "ERROR": Text4.SetFocus: Exit Sub
 If Trim(Text5) = "" Then MsgBox "Invalid Input. Please check it.", vbCritical, "ERROR": Text5.SetFocus: Exit Sub
 If Trim(Text6) = "" Then MsgBox "Invalid Input. Please check it.", vbCritical, "ERROR": Text6.SetFocus: Exit Sub
 If Trim(Text7) = "" Then MsgBox "Invalid Input. Please check it.", vbCritical, "ERROR": Text7.SetFocus: Exit Sub
    
 If IsNumeric(Text1) = False Then MsgBox "Invalid Input. Please check value!", vbCritical, "ERROR": Text1.SetFocus: Exit Sub
 If IsNumeric(Text2) = False Then MsgBox "Invalid Input. Please check value!", vbCritical, "ERROR": Text2.SetFocus: Exit Sub
 If IsNumeric(Text3) = False Then MsgBox "Invalid Input. Please check value!", vbCritical, "ERROR": Text3.SetFocus: Exit Sub
 If IsNumeric(Text4) = False Then MsgBox "Invalid Input. Please check value!", vbCritical, "ERROR": Text4.SetFocus: Exit Sub
 If IsNumeric(Text5) = False Then MsgBox "Invalid Input. Please check value!", vbCritical, "ERROR": Text5.SetFocus: Exit Sub
 If IsNumeric(Text6) = False Then MsgBox "Invalid Input. Please check value!", vbCritical, "ERROR": Text6.SetFocus: Exit Sub
 If IsNumeric(Text7) = False Then MsgBox "Invalid Input. Please check value!", vbCritical, "ERROR": Text7.SetFocus: Exit Sub

Set rsRankSave = New ADODB.Recordset
rsRankSave.Open "SELECT * FROM rank WHERE posdesc='" & Combo1 & "'", CN, adOpenStatic, adLockPessimistic
   Me.MousePointer = vbHourglass
   
 If Not rsRankSave.EOF Then
  
  With rsRankSave
    
    !rank1 = Val(Text1)
    !rank2 = Val(Text2)
    !rank3 = Val(Text3)
    !rank4 = Val(Text4)
    !rank5 = Val(Text5)
    !rank6 = Val(Text6)
    !rank7 = Val(Text7)
    
   .Update
   .Requery
   
  End With
  
   
      locked_textbox
    MsgBox "Record Successfully Updated!", vbInformation, "Success..."
    Me.MousePointer = vbDefault
 End If

Set rsRankSave = Nothing
 
End Sub

Private Sub Combo1_Click()

Set rsRank = New ADODB.Recordset
rsRank.Open "SELECT * FROM rank where posdesc ='" & Combo1 & "'", CN, adOpenStatic, adLockOptimistic

If Not rsRank.EOF Then

    Text1 = rsRank!rank1 & " " '-// to avoid null value
    Text2 = rsRank!rank2 & " "
    Text3 = rsRank!rank3 & " "
    Text4 = rsRank!rank4 & " "
    Text5 = rsRank!rank5 & " "
    Text6 = rsRank!rank6 & " "
    Text7 = rsRank!rank7 & " "

Else
  MsgBox "No record(s) to display!", vbCritical + vbOKOnly, "ERROR": Exit Sub
End If

Set rsRank = Nothing

End Sub

Private Sub Form_Load()

Call locked_textbox
Set Rs = New ADODB.Recordset
Rs.Open "SELECT * FROM [rank]", CN, adOpenStatic, adLockOptimistic
Do Until Rs.EOF
Combo1.AddItem Rs!posdesc
Rs.MoveNext
Loop

End Sub

Private Sub Form_Unload(Cancel As Integer)
Set Rs = Nothing
End Sub
