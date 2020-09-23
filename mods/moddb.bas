Attribute VB_Name = "modDB"
Option Explicit

Type T_Tips
    mTipTitle As String
    mTipAuthor As String
    mTipVer As String
    mTipAdded As String
    mTipType As String
    mTipDescription As String
    mTipCode As String
End Type

Public Db As Database           ' Database
Public Recset As Recordset      ' Recored information
Public T_def As TableDef        ' Used for the tables in the database
Public ShTip As T_Tips          ' This is used to hold all the current tip information
Public dbFilename As String     ' Name of the database file to load
Public DBTable As String        ' Name of the table in the databse to open
Public dbResult As Integer      ' Used for returning Results eg Successful or error codes
Public tName As String

' db Codes
Private Const Db_NoTable = 0    ' No table found const
Private Const Db_Ok = 3         ' Database was loaded Successfuly
Private Const Sql_ERR = 4       ' SQL Error
Private Const Sql_OK = 5        ' SQL statment was carryed out ok
Private Const Edit_Err = 6      ' There was an error while updateing the recored
Private Const Edit_ok = 7       ' The table edit was successful

Public Function LoadDB(dbFile As String, tblname As String, Optional TReadOnly As Boolean = False)
On Error Resume Next
    If TReadOnly Then SetAttr dbFilename, vbNormal ' Check to see is the filename is readonly
    Set Db = OpenDatabase(dbFilename, , TReadOnly) ' open the database
    If Db.TableDefs(DBTable).Name = "" Then
        dbResult = Db_Ok ' There was not any tables found in the database
        Exit Function   ' We stop here
    End If
    
End Function

Sub ShowTip(tblname As String, TipID As Long)
Dim StrSql As String
On Error Resume Next
    StrSql = "SELECT Tiptitle,TipBy,VBver,TipType,TipInfo,TipDate,Code " _
    & "FROM " & tblname & " WHERE ID Like '*" & TipID & "*'"
    
    Set Recset = Db.OpenRecordset(StrSql)
    If Recset.RecordCount = 0 Then
        dbResult = Sql_ERR
        Exit Sub
    Else
        With Recset
            ShTip.mTipTitle = Trim$(!Tiptitle)
            ShTip.mTipAuthor = Trim$(!TipBy)
            ShTip.mTipVer = Trim$(!VBver)
            ShTip.mTipType = Trim$(!TipType)
            ShTip.mTipDescription = Trim(!TipInfo)
            ShTip.mTipAdded = Trim$(!TipDate)
            ShTip.mTipCode = Trim(!Code)
            dbResult = Sql_OK
        End With
    End If
    StrSql = "" ' We finished with the SQL statment so we clear it out
    
End Sub
Public Function AddRecored(tblname As String)
On Error Resume Next
    Set Recset = Db.OpenRecordset(tblname)
    With Recset
        .AddNew
            !CAT = UCase(tblname)
            !Tiptitle = ShTip.mTipTitle
            !TipBy = ShTip.mTipAuthor
            !VBver = ShTip.mTipVer
            !TipType = "TEXT"
            !TipInfo = ShTip.mTipDescription
            !TipDate = ShTip.mTipAdded
            !Code = ShTip.mTipCode
            !CodeSize = Len(ShTip.mTipCode)
        .Update
        
    End With
    Set Recset = Nothing
    
End Function
Public Function DeleteRecored(tblname As String, RecID As Long)
On Error Resume Next
Dim StrSql As String
    StrSql = "SELECT ID,CAT,Tiptitle,TipBy,VBver,TipType,TipInfo,TipDate,Code,CodeSize " _
    & "FROM " & tblname & "WHERE ID Like '*" & RecID & "*'"
    
    Set Recset = Db.OpenRecordset(StrSql)
    With Recset
        .Delete
    End With
    
    Set Recset = Nothing
    StrSql = ""
    
    If Err Then Err.Clear
    
End Function
Public Function EditTable(tblname As String, RecID As Long)
On Error Resume Next
Dim StrSql As String

    StrSql = "SELECT ID,Tiptitle,TipBy,VBver,TipType,TipInfo,TipDate,Code,CodeSize FROM " & tblname & " WHERE ID Like '*" & RecID & "*'"
    
    Set Recset = Db.OpenRecordset(StrSql)
    With Recset
        .Edit
            !Tiptitle = ShTip.mTipTitle ' Title of the code
            !TipBy = ShTip.mTipAuthor   ' The person that did the code
            !VBver = ShTip.mTipVer      ' The verision of the code eg VB5, VB6 etc
            !TipType = "TEXT" ' We only dealing with text for the first ver of the program
            !TipInfo = ShTip.mTipDescription ' Description of the code eg what it does
            !TipDate = ShTip.mTipAdded ' The date the code was added or updated
            !Code = ShTip.mTipCode ' The Tip code
            !CodeSize = Len(ShTip.mTipCode) ' The size of the code tip
        .Update ' Update the recored
    End With
    
    Set Recset = Nothing
    StrSql = ""
    
    If Err Then
        dbResult = Edit_Err
        Exit Function
    Else
        dbResult = Edit_ok
    End If
    
End Function
Public Function GetRecoredCount(tblname As String) As Long
    On Error Resume Next
    Set Recset = Db.OpenRecordset(tblname)
    GetRecoredCount = Recset.RecordCount
    Set Recset = Nothing
End Function

Public Function GetDbRecCount() As Long
Dim I As Long
    For Each T_def In Db.TableDefs
        If T_def.Attributes = 0 Then
            Set Recset = Db.OpenRecordset(T_def.Name)
                I = I + Recset.RecordCount
        End If
    Next
    Set Recset = Nothing
    GetDbRecCount = I
    I = 0
    
End Function
