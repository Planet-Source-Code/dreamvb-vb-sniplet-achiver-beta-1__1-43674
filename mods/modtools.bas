Attribute VB_Name = "modtools"
Function FixPath(lzPath As String) As String
    If Right(lzPath, 1) = "\" Then FixPath = lzPath Else FixPath = lzPath & "\"
End Function

Function FindFile(lzFilename As String) As Boolean
    If Len(Trim(lzFilename)) = 0 Then Exit Function
    If Dir(lzFilename) = "" Then FindFile = False Else FindFile = True
End Function

Function GetExtension(lzPathFile As String) As String
Dim ipos As Long, I As Long
    For I = Len(lzPathFile) To 1 Step -1
        If InStr(I, lzPathFile, ".", vbBinaryCompare) Then
            ipos = I
            Exit For
        End If
    Next
    
    If ipos = 0 Then
        GetExtension = ""
    Else
        GetExtension = Mid(lzPathFile, I + 1, Len(lzPathFile))
    End If
    I = 0
End Function

Function isReadOnly(lzFilename As String) As Boolean
    ' This checks to see if a file is read only or not
    If GetAttr(lzFilename) = 33 Then
        isReadOnly = True
        Exit Function
    Else
        isReadOnly = False
    End If
    
End Function

Function SaveText(lzFile As String, StrBuffer As String)
Dim tFile As Long
    tFile = FreeFile
    Open lzFile For Output As #tFile
        Print #1, StrBuffer
    Close #tFile
    
End Function
