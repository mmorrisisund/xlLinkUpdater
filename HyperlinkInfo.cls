VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HyperlinkInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Method :
' Author : Mike Morris
' Date   : 07/03/2016
' Purpose:
'---------------------------------------------------------------------------------------
Option Explicit
' TODO: Implement HyperlinkInfo class

'---------------------------------------------------------------------------------------
' Private Fields
'---------------------------------------------------------------------------------------
'
Private m_Address As String
Private m_DisplayText As String
'

'---------------------------------------------------------------------------------------
' Public Properties
'---------------------------------------------------------------------------------------
'
Public Property Let Address(value As String)
    m_Address = value
End Property
Public Property Get Address() As String
    Address = m_Address
End Property

Public Property Let DisplayText(value As String)
    m_DisplayText = value
End Property
Public Property Get DisplayText() As String
    DisplayText = m_DisplayText
End Property
