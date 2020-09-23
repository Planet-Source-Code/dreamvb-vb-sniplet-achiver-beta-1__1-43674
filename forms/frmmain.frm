VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmmain 
   Caption         =   "DM VB-Tips Archiver - Beta 1"
   ClientHeight    =   5820
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9180
   Icon            =   "frmmain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5820
   ScaleWidth      =   9180
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CDialog 
      Left            =   7890
      Top             =   1620
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
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
      ForeColor       =   &H00404040&
      Height          =   1335
      Left            =   2760
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   19
      Top             =   1665
      Width           =   1425
   End
   Begin Project1.Line3D Line3D1 
      Height          =   45
      Left            =   0
      TabIndex        =   17
      Top             =   375
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   79
   End
   Begin Project1.Panel Panel2 
      Height          =   315
      Left            =   2760
      TabIndex        =   15
      Top             =   1305
      Width           =   5100
      _ExtentX        =   8996
      _ExtentY        =   556
      PanelStyle      =   1
      Begin VB.Label lblcodeinfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "View Info"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   1500
         TabIndex        =   18
         Top             =   30
         Width           =   915
      End
      Begin VB.Image imginfopic 
         Height          =   240
         Left            =   1200
         Picture         =   "frmmain.frx":08CA
         Top             =   30
         Width           =   240
      End
      Begin VB.Image imgbar 
         Height          =   210
         Left            =   900
         Picture         =   "frmmain.frx":0B38
         Top             =   45
         Width           =   90
      End
      Begin VB.Image imgcodepic 
         Height          =   240
         Left            =   45
         Picture         =   "frmmain.frx":0C92
         Top             =   30
         Width           =   240
      End
      Begin VB.Label lblcodename 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "#107"
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
         Left            =   345
         TabIndex        =   16
         Top             =   30
         Width           =   495
      End
   End
   Begin MSComctlLib.TreeView tv1 
      Height          =   3435
      Left            =   30
      TabIndex        =   4
      Top             =   1725
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   6059
      _Version        =   393217
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "ImageList1"
      Appearance      =   1
   End
   Begin Project1.Panel Panel1 
      Height          =   855
      Left            =   0
      TabIndex        =   3
      Top             =   420
      Width           =   7845
      _ExtentX        =   13838
      _ExtentY        =   1508
      PanelStyle      =   1
      Begin VB.Line Line4 
         BorderColor     =   &H00FFFFFF&
         X1              =   5775
         X2              =   5775
         Y1              =   90
         Y2              =   735
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         X1              =   5760
         X2              =   5760
         Y1              =   90
         Y2              =   735
      End
      Begin VB.Label lbltiptype 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "#106"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   6660
         TabIndex        =   21
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tip Type:"
         Height          =   195
         Left            =   5910
         TabIndex        =   20
         Top             =   480
         Width           =   675
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         X1              =   3270
         X2              =   3270
         Y1              =   105
         Y2              =   750
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         X1              =   3285
         X2              =   3285
         Y1              =   105
         Y2              =   750
      End
      Begin VB.Label lblsize 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "#105"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   6660
         TabIndex        =   14
         Top             =   120
         Width           =   375
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tip Size:"
         Height          =   195
         Left            =   5910
         TabIndex        =   13
         Top             =   120
         Width           =   615
      End
      Begin VB.Label lbladded 
         AutoSize        =   -1  'True
         Caption         =   "#104"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   4335
         TabIndex        =   12
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tip Added:"
         Height          =   195
         Left            =   3435
         TabIndex        =   11
         Top             =   480
         Width           =   780
      End
      Begin VB.Label lblvbver 
         AutoSize        =   -1  'True
         Caption         =   "#103"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   4335
         TabIndex        =   10
         Top             =   120
         Width           =   375
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "VB Version:"
         Height          =   195
         Left            =   3435
         TabIndex        =   9
         Top             =   120
         Width           =   825
      End
      Begin VB.Label lblAuthor 
         AutoSize        =   -1  'True
         Caption         =   "#102"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   1020
         TabIndex        =   8
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tip Author:"
         Height          =   195
         Left            =   165
         TabIndex        =   7
         Top             =   480
         Width           =   780
      End
      Begin VB.Label lbltipname 
         AutoSize        =   -1  'True
         Caption         =   "#101"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   1020
         TabIndex        =   6
         Top             =   120
         Width           =   375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tip Name:"
         Height          =   195
         Left            =   165
         TabIndex        =   5
         Top             =   120
         Width           =   735
      End
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   7890
      Top             =   1005
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483633
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   14
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":0D0D
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":105F
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":13B1
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":1703
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":1A55
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":1DA7
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":20F9
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":244B
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":279D
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":2AEF
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":2E41
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":3193
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":34E5
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":3837
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TabStrip tab1 
      Height          =   3945
      Left            =   -15
      TabIndex        =   2
      Top             =   1275
      Width           =   2715
      _ExtentX        =   4789
      _ExtentY        =   6959
      ShowTips        =   0   'False
      HotTracking     =   -1  'True
      ImageList       =   "ImageList2"
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Tips"
            Key             =   "A"
            ImageVarType    =   2
            ImageIndex      =   1
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "View All"
            Key             =   "B"
            ImageVarType    =   2
            ImageIndex      =   4
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Find"
            Key             =   "C"
            ImageVarType    =   2
            ImageIndex      =   3
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "VB Sites"
            Key             =   "D"
            ImageVarType    =   2
            ImageIndex      =   5
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7890
      Top             =   420
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":3B89
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":3EDB
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":422D
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":457F
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9180
      _ExtentX        =   16193
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   12
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "NEW"
            Object.ToolTipText     =   "Add New Tip Code"
            ImageIndex      =   6
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   4
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "NEW_CAT"
                  Text            =   "New Category"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "NEW_TIP"
                  Text            =   "New Code Tip"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "EDIT_TIP"
                  Text            =   "Edit Code Tip"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "DEL_TIP"
                  Text            =   "Delete Code Tip"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "OPEN"
            Object.ToolTipText     =   "Open Database"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "SAVE"
            Object.ToolTipText     =   "Save Code Tip"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "EDIT"
            Object.ToolTipText     =   "Edit Code Tip"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "DEL"
            Object.ToolTipText     =   "Delete Code Tip"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "COPY"
            Object.ToolTipText     =   "Copy Code Tip"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "FIND"
            Object.ToolTipText     =   "Find Text"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ABOUT"
            Object.ToolTipText     =   "About this program"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "EXIT"
            Object.ToolTipText     =   "Exit....."
            ImageIndex      =   12
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   5490
      Width           =   9180
      _ExtentX        =   16193
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13123
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuopenDB 
         Caption         =   "&On DataBase"
      End
      Begin VB.Menu mnusavetip 
         Caption         =   "&Save Tip"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuCopy 
         Caption         =   "&Copy"
      End
      Begin VB.Menu mnuFileTxt 
         Caption         =   "&Find Text"
      End
   End
   Begin VB.Menu mnuabout 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'DM VB-Tips Archiver - Beta 1
' Writen and designed by Ben Jones
' This is a small tips program I mde for vb prgrammers to keep all thire code examples in
' anyway I still have not finsihed the program and hope to add more cool things into it
' I will also be makeing a website were you can find all the updates to the program

' Ok to all programmers I noticed a small error in this code and can't seem to work it out
' Evey time you add or edit a code example and say you add a load of code the program will add it
' But that it you can edit it agian nor delete it I also went into address to delete the recored and all
' I get back is a message "Serach Key Not Found" this only happends on large recoreds I mean I added one that was 5.16 KB
' have a go and you see what I mean if someone can help me with this it whould be very helpfull
' you can contact me at dreamvb@yahoo.com ' O please take no notice at my poor spelling
' Hope you like the program.

Dim tvKey() As Long ' This is used to hold the treeview IDS
Private Sub DeleteTip()
Dim ans As Integer
    If DelOption = "DEL_CAT" Then MsgBox "This open is not available in this version", vbInformation, frmmain.Caption
    If DelOption = "DEL_TIP" Then
        Select Case TabSelect
            Case "TIP_CATS"
                ans = MsgBox("Are you sure you want to delete " _
                & tv1.SelectedItem.Text, vbYesNo Or vbQuestion, "Delete tip")
                If ans = vbNo Then Exit Sub
                DeleteRecored tName, TipID
                MsgBox "The recored has now been deleted", vbInformation, "Recored Deleted"
                InitAll
            Case "ALL_TIPS"
                ans = MsgBox("Are you sure you want to delete " _
                & tv1.SelectedItem.Text, vbYesNo Or vbQuestion, "Delete tip")
                If ans = vbNo Then Exit Sub
                DeleteRecored tName, CLng(tvID)
                MsgBox "The recored has now been deleted", vbInformation, "Recored Deleted"
                InitAll
            End Select
    End If
End Sub
Sub ShowInfoBar(mHideBar As Boolean)
    If mHideBar Then
        imgcodepic.Visible = False
        lblcodename.Visible = False
        imgbar.Visible = False
        imginfopic.Visible = False
        lblcodeinfo.Visible = False
    Else
        imgcodepic.Visible = True
        lblcodename.Visible = True
        imgbar.Visible = True
        imginfopic.Visible = True
        lblcodeinfo.Visible = True
    End If
    
End Sub
Sub InitAll()
    InitTv tv1 ' Load the treeview control with the table names in the database
    LoadChild tv1 ' Add in child notes to the list view control with tip titles and IDs
    StatusBar1.Panels(1).Text = GetDbRecCount & " Current tips found"  ' Update the statusbar text with total tip count
    tv1.Nodes(2).Selected = True    ' this selects and extends the treeview control
    tab1.Tabs(1).Selected = True    ' This selects the first tab
    EnableDisBut True               ' Disbale the toolbar buttons
    ShowInfoBar True                ' Hide the code info bar
    Toolbar1.Buttons(1).ButtonMenus(2).Enabled = False
    
End Sub
Sub EnableDisBut(mDisable As Boolean)
    If mDisable Then
        Toolbar1.Buttons(1).ButtonMenus(3).Enabled = False
        Toolbar1.Buttons(1).ButtonMenus(4).Enabled = False
        Toolbar1.Buttons(3).Enabled = False
        Toolbar1.Buttons(5).Enabled = False
        Toolbar1.Buttons(6).Enabled = False
        Toolbar1.Buttons(8).Enabled = False
        Toolbar1.Buttons(9).Enabled = False
        mnuCopy.Enabled = False
        mnuFileTxt.Enabled = False
        mnusavetip.Enabled = False
    Else
        Toolbar1.Buttons(1).ButtonMenus(3).Enabled = True
        Toolbar1.Buttons(1).ButtonMenus(4).Enabled = True
        Toolbar1.Buttons(3).Enabled = True
        Toolbar1.Buttons(5).Enabled = True
        Toolbar1.Buttons(6).Enabled = True
        Toolbar1.Buttons(9).Enabled = True
        mnuFileTxt.Enabled = True
        mnusavetip.Enabled = True
    End If

End Sub
Sub CleanVars()
    CleanTipLables
    txtcode.Text = ""
    tv1.Nodes.Clear
    Erase tvKey
    Set Recset = Nothing
    Set T_def = Nothing
    Set frmmain = Nothing
    Db.Close
    StatusBar1.Panels(1).Text = ""
    ShTip.mTipAdded = ""
    ShTip.mTipAuthor = ""
    ShTip.mTipCode = ""
    ShTip.mTipDescription = ""
    ShTip.mTipTitle = ""
    ShTip.mTipType = ""
    ShTip.mTipVer = ""
    dbFilename = ""
    DBTable = ""
    tName = ""
    tvID = ""
    TipID = 0
    TabSelect = ""
    tCodeView = False
    tTipView = False
    TipID = 0
    dbResult = 0
End Sub
Sub CleanTipLables()
    ' This clears all the tip info lables captions
    lbltipname.Caption = ""
    lblAuthor.Caption = ""
    lblvbver.Caption = ""
    lbladded.Caption = ""
    lblsize.Caption = ""
    lbltiptype.Caption = ""
    txtcode.Text = ""
End Sub

Sub UpdateTipInfo()
    lbltipname.Caption = ShTip.mTipTitle
    lblAuthor.Caption = ShTip.mTipAuthor
    lblvbver.Caption = ShTip.mTipVer
    lbladded.Caption = ShTip.mTipAdded
    lblsize.Caption = Format(Len(ShTip.mTipCode), "#,#") & " bytes"
    lbltiptype.Caption = ShTip.mTipType
    lblcodename.Caption = lbltipname.Caption
    txtcode.Text = ShTip.mTipCode
    
    imgbar.Left = lblcodename.Width + lblcodename.Left + 90
    imginfopic.Left = imgbar.Left + 150
    lblcodeinfo.Left = imginfopic.Left + imginfopic.Width + 90

End Sub
Sub ShowAllTips(tv As TreeView)

    tv.Nodes.Clear
    tv.Nodes.Add , tvwFirst, "TOP", "VB Tips", 1, 1
    
    For Each T_def In Db.TableDefs
        If T_def.Attributes = 0 Then
            Set Recset = Db.OpenRecordset(T_def.Name)
            With Recset
                While Not Recset.EOF
                    tv.Nodes.Add 1, tvwChild, T_def.Name & ":" & !id, !Tiptitle, 4, 4
                    .MoveNext
                Wend
            End With
        End If
    Next
    tv.Nodes(2).Selected = True ' Select the seconed node to extend the treeview control
    
End Sub

Sub LoadChild(tv As TreeView)
Dim I As Long, J As Long
On Error Resume Next
Erase tvKey ' Erease all the IDs
For J = 2 To tv.Nodes.Count ' 2 is the index we need to start at
    With Db
        Set Recset = .OpenRecordset(tv1.Nodes(J)) ' Open the table name
        With Recset
             While Not Recset.EOF ' Loop all the way till we reach the end of the table
                I = I + 1   ' Update our counter
                ReDim Preserve tvKey(1 To I) ' Resize the Tip ID Array
                tv1.Nodes.Add J, tvwChild, "a" & I, !Tiptitle, 4, 4
                tvKey(I) = CLng(!id) ' Update the tvkey ID array
                .MoveNext ' Move up one recored
            Wend ' Keep going to the end
        End With
    End With
Next
I = 0
J = 0

End Sub

Sub InitTv(tv As TreeView)
    tv.Nodes.Clear
    tv.Indentation = 160
    tv.Nodes.Add , tvwFirst, "TOP", "VB Tips", 1, 1
    
    For Each T_def In Db.TableDefs
            If T_def.Attributes = 0 Then
                tv1.Nodes.Add 1, tvwChild, T_def.Name, T_def.Name, 3, 2
        End If
    Next
    
End Sub


Private Sub Form_Load()

    CleanTipLables  ' Clean all the captions of the code info lables
    FlatBorder txtcode.hwnd, True ' Add flat effect to the code view textbox
    dbFilename = FixPath(App.Path) & "tips.mdb"
    If FindFile(dbFilename) = False Then
           ' error message needs to go here
        Exit Sub
    Else
        DBTable = "TIPS"
        LoadDB dbFilename, DBTable, isReadOnly(dbFilename)
    End If
    InitAll
    StatusBar1.Panels(2).Text = "DB Version " & Db.Version ' Add the db version to the statsubar
    
End Sub

Private Sub Form_Resize()
On Error Resume Next
    Line3D1.Width = frmmain.ScaleWidth - 1
    Panel1.Width = frmmain.ScaleWidth - 1
    Panel2.Width = Panel1.Width - Panel2.Left - 1
    txtcode.Width = Panel1.Width - txtcode.Left - 1
    tab1.Height = frmmain.ScaleHeight - StatusBar1.Height - tab1.Top
    tv1.Height = frmmain.ScaleHeight - StatusBar1.Height - tv1.Top - 90
    txtcode.Height = tv1.Height + 150
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    CleanVars
End Sub


Private Sub lblcodeinfo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblcodeinfo.ForeColor = vbRed
End Sub

Private Sub lblcodeinfo_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblcodeinfo.ForeColor = vbBlue
    frmcodeinfo.txtcodeinfo.Text = ShTip.mTipDescription
    frmcodeinfo.lblcodetitle.Caption = ShTip.mTipTitle
    frmcodeinfo.Show vbModal, frmmain ' Show the code information form
End Sub

Private Sub mnuabout_Click()
    frmabout.Show vbModal, frmmain ' show the about box
End Sub

Private Sub mnuCopy_Click()
    Clipboard.Clear
    Clipboard.SetText txtcode.SelText
End Sub

Private Sub mnuExit_Click()
Dim ans
    ans = MsgBox("Are you sure you want to quit now", vbQuestion Or vbYesNo, "Quit....")
    If ans = vbNo Then Exit Sub
    Unload frmmain
End Sub

Private Sub mnuFileTxt_Click()
Dim StrFind As String
Dim vPos As Long
    a = Int(frmmain.Left - frmmain.ScaleWidth) / 2
    StrFind = InputBox("Please enter in the string you like to find", "Find Text")
    If Len(Trim(StrFind)) = 0 Then Exit Sub
    vPos = InStr(1, txtcode.Text, StrFind, vbTextCompare)
    If vPos <= 0 Then
        MsgBox "The String " & Chr(34) & StrFind & Chr(34) & " was not found", vbExclamation
    Else
        txtcode.SelStart = vPos - 1
        txtcode.SelLength = Len(StrFind)
        txtcode.SetFocus
    End If
    
End Sub

Private Sub mnuopenDB_Click()
    MsgBox "There is now support for opening databases as yet will be in next version", vbInformation
End Sub

Private Sub mnusavetip_Click()
On Error Resume Next
    CDialog.CancelError = True
    CDialog.DialogTitle = "Save Code Tip"
    CDialog.Filter = "Text File(*.txt)|*.txt|"
    CDialog.FileName = ShTip.mTipTitle
    CDialog.ShowSave
    If Err Then Exit Sub
    SaveText CDialog.FileName, txtcode.Text
    MsgBox "You Code Tip has now been saved", vbInformation, "Save Code Tip"

End Sub

Private Sub Panel1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblcodeinfo.ForeColor = vbBlue
End Sub

Private Sub Panel2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblcodeinfo.ForeColor = vbBlue
End Sub

Private Sub tab1_Click()
    CleanTipLables
    StatusBar1.Panels(1).Text = GetDbRecCount & " Current tips found"
    EnableDisBut True ' Disable the toolbar buttons
    ShowInfoBar True  ' Hide the code info bar
    Select Case tab1.SelectedItem.Key
        Case "A" ' Tips Tab key
            TabSelect = "TIP_CATS" ' First Tab was seleced
            InitTv tv1    ' Load in all the nodes for the treeview control
            LoadChild tv1 ' Load in the child nodes
            tv1.Nodes(2).Selected = True ' select the seconed item in the treeview control to extend the list
        Case "B" ' Code View Tab key
            TabSelect = "ALL_TIPS" ' All Code tab was selected
            ShowAllTips tv1
        Case "C"
            '
 
    End Select
    
    Form_Resize ' call the form resize sub
    
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "OPEN"
            mnuopenDB_Click ' Call the open database menu item
        Case "SAVE"
            mnusavetip_Click ' Call the save tip menu item
        Case "EDIT"
            frmedit.Show vbModal, frmmain
        Case "DEL"
            DeleteTip
        Case "COPY"
            mnuCopy_Click       ' Call the copy menu item
        Case "FIND"
            mnuFileTxt_Click    ' Call the find menu item
        Case "ABOUT"
            mnuabout_Click      ' Show the about box
        Case "EXIT"
            mnuExit_Click       ' Call the exit menu sub
    End Select
    
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Select Case ButtonMenu.Key
        Case "EDIT_TIP"
            frmedit.Show vbModal, frmmain   ' Show the tip edit form
        Case "NEW_TIP"
            frmadd.Show vbModal, frmmain ' show the add tip form
        Case "DEL_TIP"
            DeleteTip
    End Select
    
End Sub

Private Sub tv1_Click()
On Error Resume Next
Dim ipos As Integer
    If TabSelect = "TIP_CATS" Then ' The Tips tab was selected
        tName = ""
        If tv1.SelectedItem.Key = "TOP" Then
            StatusBar1.Panels(1).Text = GetDbRecCount & " Current tips found"
            EnableDisBut True
            Toolbar1.Buttons(1).ButtonMenus(2).Enabled = False
            Exit Sub ' Exit the top item we don't need this
        Else
            tvID = Trim$(tv1.SelectedItem.Text) ' Get the node text title
        End If
        
        If tvID = tv1.SelectedItem.Key Then
            StatusBar1.Panels(1).Text = GetRecoredCount(tvID) _
            & " Records found for " & tvID
            EnableDisBut True
            ShowInfoBar True
            CleanTipLables
            Toolbar1.Buttons(1).ButtonMenus(2).Enabled = True
            Toolbar1.Buttons(6).Enabled = True
            DelOption = "DEL_CAT" ' This tell us that a tip Cat was selected for delete
            Exit Sub
        Else
            TipID = tvKey(CLng(Right(tv1.SelectedItem.Key, Len(tv1.SelectedItem.Key) - 1))) ' This holds and updates the childID of the treeview control
            tName = Trim$(tv1.SelectedItem.Parent.Text) ' This holds the table name
            ShowTip tName, TipID ' Get the tip info from info above
            ' The code below will update all of the lables to show the tip information
            UpdateTipInfo
            EnableDisBut False
        End If
    End If
    
    If TabSelect = "ALL_TIPS" Then ' The Code View Tab was selected
        ipos = InStr(1, tv1.SelectedItem.Key, ":", vbTextCompare) ' Find the position to start at
        tvID = Val(Mid(tv1.SelectedItem.Key, ipos + 1, Len(tv1.SelectedItem.Key))) ' Extract the ID
        tName = Mid(tv1.SelectedItem.Key, 1, ipos - 1) ' extract the table name
        ShowTip tName, CLng(tvID)
        UpdateTipInfo ' Update the tips info lables
        EnableDisBut False
        If tvID = 0 Then EnableDisBut True
    End If
    
    ShowInfoBar False ' Show the info bar
    Toolbar1.Buttons(1).ButtonMenus(2).Enabled = False
    Toolbar1.Buttons(8).Enabled = False
    mnuCopy.Enabled = False
    DelOption = "DEL_TIP" ' This tell us that a tip was selected for delete

End Sub

Private Sub txtcode_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Len(txtcode.SelText) <= 0 Then
        Toolbar1.Buttons(8).Enabled = False
        mnuCopy.Enabled = False
    Else
        Toolbar1.Buttons(8).Enabled = True
        mnuCopy.Enabled = True
    End If
End Sub
