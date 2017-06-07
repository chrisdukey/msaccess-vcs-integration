Attribute VB_Name = "VCS_Relation"
Option Compare Database

Option Private Module
Option Explicit


Public Sub ExportRelation(ByVal rel As DAO.Relation, ByVal filePath As String)
    Dim fso As Object
    Dim OutFile As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set OutFile = fso.CreateTextFile(filePath, overwrite:=True, Unicode:=False)

    OutFile.WriteLine rel.Attributes 'RelationAttributeEnum
    OutFile.WriteLine rel.Name
    OutFile.WriteLine rel.table
    OutFile.WriteLine rel.foreignTable
    
    Dim f As DAO.Field
    For Each f In rel.Fields
        OutFile.WriteLine "Field = Begin"
        OutFile.WriteLine f.Name
        OutFile.WriteLine f.ForeignName
        OutFile.WriteLine "End"
    Next
    
    OutFile.Close

End Sub

Public Sub ImportRelation(ByVal filePath As String)
    Dim fso As Object
    Dim InFile As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set InFile = fso.OpenTextFile(filePath, iomode:=ForReading, create:=False, Format:=TristateFalse)
    Dim rel As DAO.Relation
    Set rel = New DAO.Relation
    
    rel.Attributes = InFile.ReadLine
    rel.Name = InFile.ReadLine
    rel.table = InFile.ReadLine
    rel.foreignTable = InFile.ReadLine
    
    Dim f As DAO.Field
    Do Until InFile.AtEndOfStream
        If "Field = Begin" = InFile.ReadLine Then
            Set f = New DAO.Field
            f.Name = InFile.ReadLine
            f.ForeignName = InFile.ReadLine
            If "End" <> InFile.ReadLine Then
                Set f = Nothing
                Err.Raise 40000, "ImportRelation", "Missing 'End' for a 'Begin' in " & filePath
            End If
            rel.Fields.Append f
        End If
    Loop
    
    InFile.Close
    
    CurrentDb.Relations.Append rel

End Sub
