VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WorkbookInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' File   : WorkbookInfo
' Author : Mike Morris
' Date   : 07/03/2016
' Purpose: Class with basic information about a workbook.
'---------------------------------------------------------------------------------------
Option Explicit

'---------------------------------------------------------------------------------------
' Private Fields
'---------------------------------------------------------------------------------------
'
Private m_Name As String    ' TODO: Refactor as m_WorkbookName
Private m_Path As String
Private m_Worksheets As WorksheetInfoCollection
'

'---------------------------------------------------------------------------------------
' Public Properties
'---------------------------------------------------------------------------------------
'
Public Property Get Name() As String
    Name = m_Name
End Property

Public Property Get Path() As String
    Path = m_Path
End Property

Public Property Get Worksheets() As WorksheetInfoCollection
    Set Worksheets = m_Worksheets
End Property

'---------------------------------------------------------------------------------------
' Public Methods
'---------------------------------------------------------------------------------------
'
'---------------------------------------------------------------------------------------
' Method : IsWorkbookProtected
' Author : Mike Morris
' Date   : 07/03/2016
' Purpose: Determines if Workbook can be opened.  If not, assumed to password protected.
'          Will through error if workbook has not been set yet.
'---------------------------------------------------------------------------------------
Public Function IsWorkbookProtected() As Boolean
    Dim wkb As Workbook
    
    ' Check that workbook has been set
    If m_Name = vbNullString Then
        Err.Raise 17, "xlLinkUpdater.WorkbookInfo", "The workbook has not been set for this object."
        Exit Function
    End If
    
    ' Assume file is not protected
    IsWorkbookProtected = False
    
    Application.ScreenUpdating = False
    On Error Resume Next
        Set wkb = Workbooks.Open(m_Path & m_Name, True, True)
        ' Opening could fail for a number of reasons, but in any event
        ' this workbook could not be opened to make changes
        If Err.Number > 0 Then IsWorkbookProtected = True
    On Error GoTo 0
    wkb.Close False
    Application.ScreenUpdating = True
End Function

'---------------------------------------------------------------------------------------
' Method : SetWorkbook
' Author : Mike Morris
' Date   : 07/03/2016
' Purpose: Sets the workbook for the class.  This functions as something close to a
'          constructor for the class.  fileName must be full path & file name/extension.
'---------------------------------------------------------------------------------------
Public Function SetWorkbook(fileName As String)
    
    ' Make sure file is a valid Excel file
    If Not IsExcelFile(fileName) Then
        Err.Raise 5, "xlLinkUpdater.WorkbookInfo", "The file is not a valid Excel workbook."
        Exit Function
    End If
    
    Call SeparatePathAndFilename(fileName)
    
    'TODO: Populate worksheet info here once WorksheetInfo class is ready.
End Function

'---------------------------------------------------------------------------------------
' Private Methods
'---------------------------------------------------------------------------------------
'
'---------------------------------------------------------------------------------------
' Method : IsExcelFile
' Author : Mike Morris
' Date   : 07/03/2016
' Purpose: Determines if file is a proper Excel file based on file name extension.
'---------------------------------------------------------------------------------------
Function IsExcelFile(fileName As String) As Boolean
    Dim fileExt As String
    
    fileExt = LCase(Right$(fileName, Len(fileName) - InStrRev(fileName, ".")))
    
    ' All Excel file extensions start with xl.  It is possible that some other file types
    ' may start with "xl" as well.  This should result in the file failing to open later.
    ' And it will be assumed the file is password protected.
    IsExcelFile = InStr(1, fileExt, "xl") <> 0
End Function

'---------------------------------------------------------------------------------------
' Method : Class_Initialize
' Author : Mike Morris
' Date   : 07/03/2016
' Purpose: Prepares class for use
'---------------------------------------------------------------------------------------
Private Sub Class_Initialize()
    
End Sub

'---------------------------------------------------------------------------------------
' Method : SeparatePathAndFilename
' Author : Mike Morris
' Date   : 07/03/2016
' Purpose: Helper method to fileName into Path and File components.
'---------------------------------------------------------------------------------------
Private Sub SeparatePathAndFilename(fileName As String)
    Dim lastBackSlash As Integer
    
    lastBackSlash = InStrRev(fileName, "\")
    
    m_Path = Left$(fileName, lastBackSlash)
    m_Name = Right$(fileName, Len(fileName) - lastBackSlash)
End Sub
