Attribute VB_Name = "VCS_Loader"
Option Compare Database

Option Explicit

Public Sub loadVCS(Optional ByVal SourceDirectory As String)
    If SourceDirectory = vbNullString Then
      SourceDirectory = CurrentProject.Path & "\MSAccess-VCS\"
    End If

'check if directory exists! - SourceDirectory could be a file or not exist
On Error GoTo Err_DirCheck
    If ((GetAttr(SourceDirectory) And vbDirectory) = vbDirectory) Then
        GoTo Fin_DirCheck
    Else
        'SourceDirectory is not a directory
        Err.Raise 60000, "loadVCS", "Source Directory specified is not a directory"
    End If

Err_DirCheck:
    
    If Err.Number = 53 Then 'SourceDirectory does not exist
        Debug.Print Err.Number & " | " & "File/Directory not found"
    Else
        Debug.Print Err.Number & " | " & Err.Description
    End If
    Exit Sub
Fin_DirCheck:

    'delete if modules already exist + provide warning of deletion?

    On Error GoTo Err_DelHandler

    Dim FileName As String
    'Use the list of files to import as the list to delete
    FileName = Dir$(SourceDirectory & "*.bas")
    Do Until Len(FileName) = 0
        'strip file type from file name
        FileName = Left$(FileName, InStrRev(FileName, ".bas") - 1)
        DoCmd.DeleteObject acModule, FileName
        FileName = Dir$()
    Loop

    GoTo Fin_DelHandler
    
Err_DelHandler:
    If Err.Number <> 7874 Then 'is not - can't find object
        Debug.Print "WARNING (" & Err.Number & ") | " & Err.Description
    End If
    Resume Next
    
Fin_DelHandler:
    FileName = vbNullString

'import files from specific dir? or allow user to input their own dir?
On Error GoTo Err_LoadHandler

    FileName = Dir$(SourceDirectory & "*.bas")
    Do Until Len(FileName) = 0
        'strip file type from file name
        FileName = Left$(FileName, InStrRev(FileName, ".bas") - 1)
        Application.LoadFromText acModule, FileName, SourceDirectory & FileName & ".bas"
        FileName = Dir$()
    Loop

    GoTo Fin_LoadHandler
    
Err_LoadHandler:
    Debug.Print Err.Number & " | " & Err.Description
    Resume Next

Fin_LoadHandler:
    Debug.Print "Done"

End Sub
