VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private cFile As clsFile
Private mCol As Collection
Private cUpdating As Boolean
Private mPath As String

Public Event OnError(Number As Long, Description As String)
Public Property Get Busy() As Boolean
    Busy = cUpdating
End Property


Public Property Get path() As String
    path = mPath
End Property
Public Property Let path(NewVal As String)
    Dim result
    'load the new path into the file collection
    If Not mPath = NewVal Then
        mPath = NewVal
        result = LoadPath()
        
    End If
End Property

Public Function Count() As Integer
    Count = mCol.Count
End Function
Public Function Item(Index As Variant) As clsFile
    On Error Resume Next
    Set Item = mCol(Index)
End Function
Public Function Refresh()
    LoadPath
End Function
Private Function LoadPath() As Boolean
    On Error GoTo ErrorCode
    cUpdating = True
    Dim fName As String
    'get the first directory listing
    fName = Dir(mPath, vbHidden + vbReadOnly + vbSystem + vbArchive + vbDirectory + vbNormal)
    Clear_mCol
    'loop until all items are listed
    Do While fName <> ""
        Set cFile = New clsFile
        
        'get the file details & add it to our collection
        cFile.SetFile fName, mPath
        mCol.Add cFile, cFile.FileName
        
        'get the next item
        fName = Dir
        DoEvents
    Loop
    LoadPath = True
    cUpdating = False
    Exit Function
ErrorCode:
    LoadPath = False
    cUpdating = False
    RaiseEvent OnError(Err.Number, Err.Description)

End Function

Private Function Clear_mCol()
    'clear the collection.
    Do While mCol.Count > 0
        mCol.Remove 1
    Loop
End Function

Private Sub Class_Initialize()
    'initialize the declared objects
    Set mCol = New Collection
    Set cFile = New clsFile
End Sub

Private Sub Class_Terminate()
    'clean up behind us
    Set mCol = Nothing
    Set cFile = Nothing
End Sub
