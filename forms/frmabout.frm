VERSION 5.00
Begin VB.Form frmabout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About DM VB-Tips Archiver - Beta 1"
   ClientHeight    =   2790
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5820
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   5820
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdok 
      Caption         =   "&OK"
      Height          =   375
      Left            =   4650
      TabIndex        =   3
      Top             =   2250
      Width           =   1020
   End
   Begin Project1.Line3D Line3D1 
      Height          =   45
      Left            =   0
      TabIndex        =   2
      Top             =   780
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   79
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   765
      Left            =   0
      ScaleHeight     =   765
      ScaleWidth      =   5820
      TabIndex        =   1
      Top             =   0
      Width           =   5820
      Begin VB.Image Image1 
         Height          =   720
         Left            =   105
         Picture         =   "frmabout.frx":0000
         Top             =   30
         Width           =   720
      End
      Begin VB.Image Image2 
         Height          =   405
         Left            =   960
         Picture         =   "frmabout.frx":07BC
         Top             =   165
         Width           =   4365
      End
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright © 2003 Ben Jones"
      Height          =   195
      Left            =   1665
      TabIndex        =   6
      Top             =   1800
      Width           =   2040
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "This program is freeware"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1545
      TabIndex        =   5
      Top             =   1455
      Width           =   2385
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A handy program for programmers to keep all there source code files safe."
      Height          =   195
      Left            =   360
      TabIndex        =   4
      Top             =   1080
      Width           =   5190
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DM VB-Tips Archiver - Beta 1"
      Height          =   195
      Left            =   2145
      TabIndex        =   0
      Top             =   270
      Width           =   2085
   End
End
Attribute VB_Name = "frmabout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdok_Click()
    Unload frmabout ' Unload the form
End Sub

Private Sub Form_Resize()
    Line3D1.Width = frmabout.Width - 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmabout = Nothing ' Release the form form memory
End Sub

