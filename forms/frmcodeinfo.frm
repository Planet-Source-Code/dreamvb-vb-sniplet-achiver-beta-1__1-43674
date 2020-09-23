VERSION 5.00
Begin VB.Form frmcodeinfo 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Code Information"
   ClientHeight    =   3000
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6210
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   6210
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdcopy 
      Caption         =   "C&opy to Clipboard"
      Height          =   360
      Left            =   3390
      TabIndex        =   4
      Top             =   2505
      Width           =   1635
   End
   Begin VB.CommandButton cmdclose 
      Caption         =   "&Close"
      Height          =   360
      Left            =   5220
      TabIndex        =   3
      Top             =   2505
      Width           =   945
   End
   Begin VB.TextBox txtcodeinfo 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1740
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   660
      Width           =   6195
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   600
      Left            =   0
      ScaleHeight     =   600
      ScaleWidth      =   6210
      TabIndex        =   0
      Top             =   0
      Width           =   6210
      Begin VB.Label lblcodetitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   720
         TabIndex        =   1
         Top             =   135
         Width           =   60
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   120
         Picture         =   "frmcodeinfo.frx":0000
         Top             =   45
         Width           =   480
      End
   End
End
Attribute VB_Name = "frmcodeinfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdclose_Click()
    lblcodetitle.Caption = "" ' Clear the code titlename
    txtcodeinfo.Text = "" ' Clean out the textbox
    Unload frmcodeinfo ' Unload the form
End Sub

Private Sub cmdcopy_Click()
    Clipboard.Clear ' Clear the clipboard
    Clipboard.SetText txtcodeinfo.Text, vbCFText ' Copt the textbox contents to the clipboard
    MsgBox "The text has now been copied to the Clipboard", vbInformation, "Copy Text"
    
End Sub

Private Sub Form_Load()
    FlatBorder txtcodeinfo.hwnd, True ' Add flat affect to the textbox
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmcodeinfo = Nothing ' Release the form from memory
End Sub
