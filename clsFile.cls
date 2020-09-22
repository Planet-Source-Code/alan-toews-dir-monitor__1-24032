VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Key As String
Public FileName As String
Public path As String
Public Attrib As VbFileAttribute
Public Size As Long
Public LastMod As String
Public Created As String
Public Directory As Boolean


Public Function SetFile(fName As String, Optional fPath As String = "") As Boolean
    On Error GoTo ErrorCode
    
    'get the details about the file.
    If fPath = "" Then fPath = App.path
    fPath = Trim(Replace(fPath & "\ ", "\\ ", "\"))
    
    FileName = fName
    path = fPath
    Key = path & FileName
    Attrib = GetAttr(fPath & fName)
    Directory = CBool(Attrib And vbDirectory)
    Size = FileLen(fPath & fName)
    Created = FileDateTime(fPath & fName)
    SetFile = True
    Exit Function
ErrorCode:
    SetFile = False
End Function

Public Function IsAttrib(TestAttrib As VbFileAttribute) As Boolean
    'Check if a file has the specified attribute
    IsAttrib = CBool(Attrib And TestAttrib)
End Function
