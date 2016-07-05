VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WorksheetInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' File   : WorksheetInfo
' Author : Mike Morris
' Date   : 07/03/2016
' Purpose: Class with basic information about a worksheet.
'---------------------------------------------------------------------------------------

Option Explicit

'---------------------------------------------------------------------------------------
' Private Fields
'---------------------------------------------------------------------------------------
'
Private m_WorksheetName As String
Private m_IsProtected As Boolean
Private m_IsPasswordProtected As Boolean
Private m_Hyperlinks As HyperlinkInfoCollection
'

'---------------------------------------------------------------------------------------
' Public Properties
'---------------------------------------------------------------------------------------
'
Public Property Get WorksheetName() As String
    WorksheetName = m_WorksheetName
End Property

Public Property Get IsProtected() As Boolean
    IsProtected = m_IsProtected
End Property

Public Property Get IsPasswordProtected() As Boolean
    IsPasswordProtected = m_IsPasswordProtected
End Property

Public Property Get Hyperlinks() As HyperlinkInfoCollection
    Hyperlinks = m_Hyperlinks
End Property

'---------------------------------------------------------------------------------------
' Public Methods
'---------------------------------------------------------------------------------------
'
'---------------------------------------------------------------------------------------
' Method : SetWorksheet
' Author : Mike Morris
' Date   : 07/03/2016
' Purpose: Sets the worksheet for the class.  This functions as something close to a
'          constructor for the class.
'---------------------------------------------------------------------------------------
Public Sub SetWorksheet(ws As Worksheet)
    m_WorksheetName = ws.Name
    m_IsProtected = HasProtection(ws)
    
    If m_IsProtected Then
        m_IsPasswordProtected = HasPasswordProtection(ws)
    Else
        ' If there is no protection at all, it can't be password protected either
        m_IsPasswordProtected = False
    End If
End Sub


'---------------------------------------------------------------------------------------
' Private Methods
'---------------------------------------------------------------------------------------
'
'---------------------------------------------------------------------------------------
' Method : HasProtection
' Author : Mike Morris
' Date   : 07/03/2016
' Purpose: Determines if Worksheet has some form of Protection turned on. Does not
'          determine if Worksheet is password protected
'---------------------------------------------------------------------------------------
Private Function HasProtection(ws As Worksheet) As Boolean
    HasProtection = ws.ProtectContents Or ws.ProtectDrawingObjects Or ws.ProtectScenarios
End Function

'---------------------------------------------------------------------------------------
' Method : HasPasswordProtection
' Author : Mike Morris
' Date   : 07/03/2016
' Purpose: Determines if Worksheet is password protected
'---------------------------------------------------------------------------------------
Private Function HasPasswordProtection(ws As Worksheet) As Boolean
    On Error Resume Next
        ' Attempt to Unprotect sheet with no password.
        Call ws.Unprotect("")
    On Error GoTo 0
    
    If HasProtection(ws) Then
        ' Unprotect failed and worksheet has a password
        HasPasswordProtection = True
    Else
        ' Make sure protection is turned back on, since the Unprotect method was successful
        ws.Protect
        HasPasswordProtection = False
    End If
End Function
