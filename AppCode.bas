Attribute VB_Name = "AppCode"
Option Explicit

Sub FindHyperlinks(xlFilePath As String)
    Dim wkb As Workbook
    Dim wks As Worksheet
    Dim h As Hyperlink
    Dim row As Long: row = Cells(Rows.Count, "A").End(xlUp) + 1
    
    Application.ScreenUpdating = False
    Set wkb = Workbooks.Open(xlFilePath, True, True)
    
    For Each wks In wkb.Sheets
        For Each h In wks.Hyperlinks
            Cells(row, "C").value = h.Address
        Next
    Next
    
    wkb.Close False
    Application.ScreenUpdating = True
End Sub

Sub GetFiles()
    Dim fso As FileSystemObject: Set fso = New FileSystemObject
    Dim oTopFolder As Folder
    
    Set oTopFolder = fso.GetFolder("C:\Users\m294355\Desktop\")
    
    Call GetFilesRecursively(oTopFolder)
        
End Sub

Sub GetFilesRecursively(oFolder As Folder)
    Dim oFile As File
    Dim oSubFolder As Folder
    Dim row As Long
    
    row = Cells(Rows.Count, "A").End(xlUp).row + 1
    
    For Each oFile In oFolder.Files
        ' Do not attempt to analyze this workbook
        If oFile.Name = ThisWorkbook.Name Then GoTo NextIteration
        
        ' Excel files that are hidden are assumed to be temp files that should not be analyzed
        If IsExcelFile(oFile.Path) And Not ((oFile.Attributes And vbHidden) = vbHidden) Then
            Cells(row, "A").value = oFile.Name
            Cells(row, "B").value = oFile.Path
            Cells(row, "C").value = IsExcelFileProtected(oFile.Path)
            row = row + 1
        End If
NextIteration:
    Next
    
    For Each oSubFolder In oFolder.SubFolders
        Call GetFilesRecursively(oSubFolder)
    Next
    
End Sub

Function IsExcelFile(fileName As String) As Boolean
    Dim fileExt As String
    
    fileExt = Right$(fileName, Len(fileName) - InStrRev(fileName, "."))
    
    IsExcelFile = InStr(1, fileExt, "xl") <> 0
End Function

