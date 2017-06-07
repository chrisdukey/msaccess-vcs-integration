Attribute VB_Name = "VCS_IE_Functions"
Option Compare Database

Option Private Module
Option Explicit

Private Const AggressiveSanitize As Boolean = True
Private Const StripPublishOption As Boolean = True

' Constants for Scripting.FileSystemObject API
Public Const ForReading = 1, ForWriting = 2, ForAppending = 8
Public Const TristateTrue = -1, TristateFalse = 0, TristateUseDefault = -2

Private Const VCS_FWDSLASH As String = "-VCS_fwdslash-"
Private Const VCS_BACKSLASH As String = "-VCS_backslash-"
Private Const VCS_GTRTHAN As String = "-VCS_gtrthan-"
Private Const VCS_LESSTHAN As String = "-VCS_lessthan-"
Private Const VCS_ASTERISK As String = "-VCS_asterisk-"
Private Const VCS_QUESMARK As String = "-VCS_quesmark-"
Private Const VCS_COLON As String = "-VCS_colon-"
Private Const VCS_DBLQUOTE As String = "-VCS_dblquote-"
Private Const VCS_PIPE As String = "-VCS_pipe-"

' Can we export without closing the form?

' Export a database object with optional UCS2-to-UTF-8 conversion.
Public Sub ExportObject(ByVal obj_type_num As Integer, ByVal obj_name As String, _
                    ByVal file_path As String, Optional ByVal Ucs2Convert As Boolean = False)
    
    file_path = SanitizeExportFilePath(file_path)
    
    obj_name = RebuildObjectName(obj_name)
    
    VCS_Dir.MkDirIfNotExist Left$(file_path, InStrRev(file_path, "\"))
    If Ucs2Convert Then
        Dim tempFileName As String
        tempFileName = VCS_File.TempFile()
        Application.SaveAsText obj_type_num, obj_name, tempFileName
        VCS_File.ConvertUcs2Utf8 tempFileName, file_path
    Else
        Application.SaveAsText obj_type_num, obj_name, file_path
    End If
End Sub

' Import a database object with optional UTF-8-to-UCS2 conversion.
Public Sub ImportObject(ByVal obj_type_num As Integer, ByVal obj_name As String, _
                    ByVal file_path As String, Optional ByVal Ucs2Convert As Boolean = False)
    

    If Not VCS_Dir.FileExists(file_path) Then Exit Sub
    
    file_path = SanitizeExportFilePath(file_path)
    
    obj_name = RebuildObjectName(obj_name)
    
    If Ucs2Convert Then
        Dim tempFileName As String
        tempFileName = VCS_File.TempFile()
        VCS_File.ConvertUtf8Ucs2 file_path, tempFileName
        Application.LoadFromText obj_type_num, obj_name, tempFileName
        
        Dim fso As Object
        Set fso = CreateObject("Scripting.FileSystemObject")
        fso.DeleteFile tempFileName
    Else
        Application.LoadFromText obj_type_num, obj_name, file_path
    End If
End Sub

Public Function RebuildObjectName(ByVal FileName As String) As String
Dim RegEx As Object

Set RegEx = CreateObject("vbscript.regexp")
' This line ensures it replaces every occurrance
RegEx.Global = True

' Using regular expressions to replace characters in form and report names that don't translate to
' Windows file names
RegEx.Pattern = VCS_FWDSLASH
FileName = RegEx.Replace(FileName, "/")

RegEx.Pattern = VCS_BACKSLASH
FileName = RegEx.Replace(FileName, "\")

RegEx.Pattern = VCS_GTRTHAN
FileName = RegEx.Replace(FileName, ">")

RegEx.Pattern = VCS_LESSTHAN
FileName = RegEx.Replace(FileName, "<")

RegEx.Pattern = VCS_COLON
FileName = RegEx.Replace(FileName, ":")

RegEx.Pattern = VCS_ASTERISK
FileName = RegEx.Replace(FileName, "*")

RegEx.Pattern = VCS_QUESMARK
FileName = RegEx.Replace(FileName, "?")

RegEx.Pattern = VCS_DBLQUOTE
FileName = RegEx.Replace(FileName, Chr(34))

RegEx.Pattern = VCS_PIPE
FileName = RegEx.Replace(FileName, "¦")

RebuildObjectName = FileName

End Function

Public Function SanitizeExportFilePath(ByVal file_path As String) As String
Dim RegEx As Object
Dim FileName As String
Dim filefolder As String

FileName = getFilenameFromPath(file_path)
filefolder = getFolderFromPath(file_path)

Set RegEx = CreateObject("vbscript.regexp")
' This line ensures it replaces every occurrance
RegEx.Global = True

' Using regular expressions to replace characters in form and report names that don't translate to
' Windows file names
RegEx.Pattern = "[\/]"
FileName = RegEx.Replace(FileName, VCS_FWDSLASH)

RegEx.Pattern = "[\\]"
FileName = RegEx.Replace(FileName, VCS_BACKSLASH)

RegEx.Pattern = "[>]"
FileName = RegEx.Replace(FileName, VCS_GTRTHAN)

RegEx.Pattern = "[<]"
FileName = RegEx.Replace(FileName, VCS_LESSTHAN)

RegEx.Pattern = "[:]"
FileName = RegEx.Replace(FileName, VCS_COLON)

RegEx.Pattern = "[*]"
FileName = RegEx.Replace(FileName, VCS_ASTERISK)

RegEx.Pattern = "[?]"
FileName = RegEx.Replace(FileName, VCS_QUESMARK)

RegEx.Pattern = "[" & Chr(34) & "]"
FileName = RegEx.Replace(FileName, VCS_DBLQUOTE)

RegEx.Pattern = "[¦]"
FileName = RegEx.Replace(FileName, VCS_PIPE)

SanitizeExportFilePath = filefolder & FileName

End Function

Sub testsanitagain()
Debug.Print SanitizeExportFilePath("C:\Users\Chris Duke\OneDrive - BOCS\BOC_Apps\PayBudget devt\in development\source\queries\Next Year Budget Contract Details - Set G/CC To N.bas")
End Sub

Public Function SanitizeImportFilePath(ByVal file_path As String) As String
Dim RegEx As Object
Dim FileName As String
Dim filefolder As String

FileName = getFilenameFromPath(file_path)
filefolder = getFolderFromPath(file_path)

Set RegEx = CreateObject("vbscript.regexp")
' This line ensures it replaces every occurrance
RegEx.Global = True

' Using regular expressions to replace characters in form and report names that don't translate to
' Windows file names
RegEx.Pattern = VCS_FWDSLASH
FileName = RegEx.Replace(FileName, "/")

RegEx.Pattern = VCS_BACKSLASH
FileName = RegEx.Replace(FileName, "\")

RegEx.Pattern = VCS_GTRTHAN
FileName = RegEx.Replace(FileName, ">")

RegEx.Pattern = VCS_LESSTHAN
FileName = RegEx.Replace(FileName, "<")

RegEx.Pattern = VCS_COLON
FileName = RegEx.Replace(FileName, ":")

RegEx.Pattern = VCS_ASTERISK
FileName = RegEx.Replace(FileName, "*")

RegEx.Pattern = VCS_QUESMARK
FileName = RegEx.Replace(FileName, "?")

RegEx.Pattern = VCS_DBLQUOTE
FileName = RegEx.Replace(FileName, Chr(34))

RegEx.Pattern = VCS_PIPE
FileName = RegEx.Replace(FileName, "¦")

SanitizeImportFilePath = filefolder & FileName

End Function

Public Sub testSanit()
Dim str As String

str = SanitizeExportFilePath("c:\chris\blah\hy?>cheese.txt")

Debug.Print str

str = SanitizeImportFilePath(str)

Debug.Print str
End Sub



Public Function getFilenameFromPath(strFullPath As String) As String
    Dim I As Integer

    For I = Len(strFullPath) To 1 Step -1
        If Mid(strFullPath, I, 1) = "\" Then
            getFilenameFromPath = Right(strFullPath, Len(strFullPath) - I)
            Exit For
        End If
    Next
End Function

Public Function getFolderFromPath(strFullPath As String) As String
    Dim I As Integer

    For I = Len(strFullPath) To 1 Step -1
        If Mid(strFullPath, I, 1) = "\" Then
            getFolderFromPath = Left(strFullPath, I)
            Exit For
        End If
    Next
End Function


'shouldn't this be SanitizeTextFile (Singular)?

' For each *.txt in `Path`, find and remove a number of problematic but
' unnecessary lines of VB code that are inserted automatically by the
' Access GUI and change often (we don't want these lines of code in
' version control).
Public Sub SanitizeTextFiles(ByVal Path As String, ByVal Ext As String)

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    '
    '  Setup Block matching Regex.
    Dim rxBlock As Object
    Set rxBlock = CreateObject("VBScript.RegExp")
    rxBlock.ignoreCase = False
    '
    '  Match PrtDevNames / Mode with or  without W
    Dim srchPattern As String
    srchPattern = "PrtDev(?:Names|Mode)[W]?"
    If (AggressiveSanitize = True) Then
      '  Add and group aggressive matches
      srchPattern = "(?:" & srchPattern
      srchPattern = srchPattern & "|GUID|""GUID""|NameMap|dbLongBinary ""DOL"""
      srchPattern = srchPattern & ")"
    End If
    '  Ensure that this is the begining of a block.
    srchPattern = srchPattern & " = Begin"
'Debug.Print srchPattern
    rxBlock.Pattern = srchPattern
    '
    '  Setup Line Matching Regex.
    Dim rxLine As Object
    Set rxLine = CreateObject("VBScript.RegExp")
    srchPattern = "^\s*(?:"
    srchPattern = srchPattern & "Checksum ="
    srchPattern = srchPattern & "|BaseInfo|NoSaveCTIWhenDisabled =1"
    If (StripPublishOption = True) Then
        srchPattern = srchPattern & "|dbByte ""PublishToWeb"" =""1"""
        srchPattern = srchPattern & "|PublishOption =1"
    End If
    srchPattern = srchPattern & ")"
'Debug.Print srchPattern
    rxLine.Pattern = srchPattern
    Dim FileName As String
    FileName = Dir$(Path & "*." & Ext)
    Dim isReport As Boolean
    isReport = False
    
    Do Until Len(FileName) = 0
        DoEvents
        Dim obj_name As String
        obj_name = Mid$(FileName, 1, InStrRev(FileName, ".") - 1)

        Dim InFile As Object
        Set InFile = fso.OpenTextFile(Path & obj_name & "." & Ext, iomode:=ForReading, create:=False, Format:=TristateFalse)
        Dim OutFile As Object
        Set OutFile = fso.CreateTextFile(Path & obj_name & ".sanitize", overwrite:=True, Unicode:=False)
    
        Dim getLine As Boolean
        getLine = True
        
        Do Until InFile.AtEndOfStream
            DoEvents
            Dim txt As String
            '
            ' Check if we need to get a new line of text
            If getLine = True Then
                txt = InFile.ReadLine
            Else
                getLine = True
            End If
            '
            ' Skip lines starting with line pattern
            If rxLine.test(txt) Then
                Dim rxIndent As Object
                Set rxIndent = CreateObject("VBScript.RegExp")
                rxIndent.Pattern = "^(\s+)\S"
                '
                ' Get indentation level.
                Dim matches As Object
                Set matches = rxIndent.Execute(txt)
                '
                ' Setup pattern to match current indent
                Select Case matches.Count
                    Case 0
                        rxIndent.Pattern = "^" & vbNullString
                    Case Else
                        rxIndent.Pattern = "^" & matches(0).SubMatches(0)
                End Select
                rxIndent.Pattern = rxIndent.Pattern + "\S"
                '
                ' Skip lines with deeper indentation
                Do Until InFile.AtEndOfStream
                    txt = InFile.ReadLine
                    If rxIndent.test(txt) Then Exit Do
                Loop
                ' We've moved on at least one line so do get a new one
                ' when starting the loop again.
                getLine = False
            '
            ' skip blocks of code matching block pattern
            ElseIf rxBlock.test(txt) Then
                Do Until InFile.AtEndOfStream
                    txt = InFile.ReadLine
                    If InStr(txt, "End") Then Exit Do
                Loop
            ElseIf InStr(1, txt, "Begin Report") = 1 Then
                isReport = True
                OutFile.WriteLine txt
            ElseIf isReport = True And (InStr(1, txt, "    Right =") Or InStr(1, txt, "    Bottom =")) Then
                'skip line
                If InStr(1, txt, "    Bottom =") Then
                    isReport = False
                End If
            Else
                OutFile.WriteLine txt
            End If
        Loop
        OutFile.Close
        InFile.Close

        fso.DeleteFile (Path & FileName)

        Dim thisFile As Object
        Set thisFile = fso.GetFile(Path & obj_name & ".sanitize")
        thisFile.Move (Path & FileName)
        FileName = Dir$()
    Loop

End Sub



