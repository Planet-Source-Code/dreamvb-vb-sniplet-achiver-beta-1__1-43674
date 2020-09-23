VERSION 5.00
Begin VB.Form frmedit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edit Code Tip"
   ClientHeight    =   7230
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6900
   Icon            =   "FRMEDIT.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7230
   ScaleWidth      =   6900
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdcancel 
      Caption         =   "&Cnacel"
      Height          =   375
      Left            =   5625
      TabIndex        =   9
      Top             =   6675
      Width           =   1155
   End
   Begin VB.CommandButton cmdapply 
      Caption         =   "&Apply"
      Height          =   375
      Left            =   2955
      TabIndex        =   7
      Top             =   6675
      Width           =   1155
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "&OK"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4245
      TabIndex        =   8
      Top             =   6675
      Width           =   1155
   End
   Begin VB.Frame Frame1 
      Height          =   6465
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   6705
      Begin VB.TextBox txtver 
         Height          =   300
         Left            =   1305
         TabIndex        =   3
         Top             =   1942
         Width           =   1900
      End
      Begin VB.TextBox txtcode 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1920
         Left            =   285
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   4305
         Width           =   6270
      End
      Begin VB.TextBox txtnotes 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Left            =   285
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   2910
         Width           =   6270
      End
      Begin VB.TextBox txtadded 
         Height          =   300
         Left            =   4410
         TabIndex        =   4
         Top             =   1065
         Width           =   1900
      End
      Begin VB.TextBox txtAuthor 
         Height          =   300
         Left            =   1305
         TabIndex        =   2
         Top             =   1500
         Width           =   1900
      End
      Begin VB.TextBox txtTipName 
         Height          =   300
         Left            =   1305
         TabIndex        =   1
         Top             =   1065
         Width           =   1900
      End
      Begin Project1.Line3D Line3D1 
         Height          =   45
         Left            =   270
         TabIndex        =   11
         Top             =   810
         Width           =   6270
         _ExtentX        =   11060
         _ExtentY        =   79
      End
      Begin Project1.Line3D Line3D2 
         Height          =   45
         Left            =   270
         TabIndex        =   17
         Top             =   2460
         Width           =   6270
         _ExtentX        =   11060
         _ExtentY        =   79
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Insert todays date"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   4425
         TabIndex        =   19
         Top             =   1425
         Width           =   1395
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Main Code Body:"
         Height          =   195
         Left            =   285
         TabIndex        =   18
         Top             =   4035
         Width           =   1215
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Code Notes:"
         Height          =   195
         Left            =   285
         TabIndex        =   16
         Top             =   2640
         Width           =   885
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Tip Added:"
         Height          =   195
         Left            =   3525
         TabIndex        =   15
         Top             =   1095
         Width           =   780
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "VB Version:"
         Height          =   195
         Left            =   285
         TabIndex        =   14
         Top             =   1995
         Width           =   825
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Tip Author:"
         Height          =   195
         Left            =   285
         TabIndex        =   13
         Top             =   1545
         Width           =   780
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tip Name:"
         Height          =   195
         Left            =   285
         TabIndex        =   12
         Top             =   1095
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Modify your code Tip"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   870
         TabIndex        =   10
         Top             =   345
         Width           =   2520
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   210
         Picture         =   "FRMEDIT.frx":0CCA
         Top             =   270
         Width           =   480
      End
   End
End
Attribute VB_Name = "frmedit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private TableName As String
Private TableID As Long


Private Sub cmdapply_Click()
    cmdok.Enabled = True ' Enable the OK Button
    ' Assign the shtip the new data
    ShTip.mTipTitle = Trim$(txtTipName.Text)
    ShTip.mTipAuthor = Trim$(txtAuthor.Text)
    ShTip.mTipVer = Trim$(txtver.Text)
    ShTip.mTipDescription = Trim$(txtnotes.Text)
    ShTip.mTipAdded = Trim$(txtadded.Text)
    ShTip.mTipCode = Trim(txtcode.Text)
    EditTable TableName, TableID
    cmdapply.Enabled = False
    
End Sub

Private Sub cmdcancel_Click()
    Unload frmedit  ' Unload the form
End Sub

Private Sub cmdok_Click()
    frmmain.InitAll
    cmdcancel_Click
End Sub

Private Sub Form_Load()

    FlatTextBox frmedit
    
    txtTipName.Text = ShTip.mTipTitle
    txtAuthor.Text = ShTip.mTipAuthor
    txtver.Text = ShTip.mTipVer
    txtadded.Text = ShTip.mTipAdded
    txtnotes.Text = ShTip.mTipDescription
    txtcode.Text = ShTip.mTipCode
    
    Select Case TabSelect
        Case "TIP_CATS"
            TableName = tName
            TableID = TipID
        Case "ALL_TIPS"
            TableName = tName
            TableID = CLng(tvID)
    End Select
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmedit = Nothing
End Sub

Private Sub Label8_Click()
    txtadded.Text = Date
End Sub
