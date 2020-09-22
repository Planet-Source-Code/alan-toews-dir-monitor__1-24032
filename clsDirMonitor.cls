VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Monitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Enum cEnumChangeType
    cChangesInAny = 15
    cChangeInFileSize = 1
    cChangeInAttributes = 2
    cChangeFilesAdded = 4
    cChangeFilesRemoved = 8
End Enum

Private WithEvents cFilNew As clsFiles, WithEvents cFilComp As clsFiles
Attribute cFilNew.VB_VarHelpID = -1
Attribute cFilComp.VB_VarHelpID = -1
Private cPath As String
Private cChanges As cEnumChangeType

Dim WithEvents cTimer As Timer
Attribute cTimer.VB_VarHelpID = -1

Public Event OnError(ByVal Number As Integer, ByVal Description As String)
Public Event OnChange(FileName As String, ChangeType As cEnumChangeType)
Public Property Get ChangeType() As cEnumChangeType
    ChangeType = cChanges
End Property

Public Property Let ChangeType(NewVal As cEnumChangeType)
    cChanges = NewVal
End Property

Public Property Get path() As String
    path = cPath
End Property
Public Property Let path(NewVal As String)
    cPath = NewVal
    
    cFilComp.path = cPath
    cFilNew.path = cPath
End Property
Public Property Get Files() As clsFiles
    Set Files = cFilComp
End Property
Public Property Get Interval() As Integer
    Interval = cTimer.Interval
End Property
Public Property Let Interval(NewVal As Integer)
    cTimer.Interval = NewVal
End Property
Public Property Get Enabled() As Boolean
    Enabled = cTimer.Enabled
End Property
Public Property Let Enabled(NewVal As Boolean)
    Do While cFilNew.Busy Or cFilComp.Busy
        DoEvents
        Debug.Print "waiting..."
    Loop
    cTimer.Enabled = NewVal
End Property

Private Sub cFilComp_OnError(Number As Long, Description As String)
    RaiseEvent OnError(Number, Description)
End Sub



Private Sub cFilNew_OnError(Number As Long, Description As String)
    RaiseEvent OnError(Number, Description)
End Sub

Private Sub Class_Initialize()
    Load frmTimer
    frmTimer.Visible = False
    
    'initialize all our called objects.
    Set cFilNew = New clsFiles
    Set cFilComp = New clsFiles
    
    'I tried using an api timer & callback, but it slowed down
    'the whole computer while it was counting.
    'Perhaps I was using it wrong, but I needed to finish this code off
    'quickly, so I linked to a timer object on a form in the dll.
    Set cTimer = frmTimer.Timer1
    
    'initially, the class will look for all possible changes.
    cChanges = cChangesInAny
    
    'make sure the timer is disabled, and the interval is set to ~1 second
    cTimer.Enabled = False
    cTimer.Interval = 1000
End Sub

Private Sub Class_Terminate()

    'clean up behind us.
    Set cFilNew = Nothing
    Set cFilComp = Nothing
    Set cTimer = Nothing
    
End Sub

Private Sub cTimer_Timer()
    cFilNew.Refresh

    On Error GoTo ErrorCode

    Dim fAttribChange As Integer, fnCount As Integer, _
        fcCount As Integer, CheckSize As Boolean, _
        CheckAttrib As Boolean, CheckAdded As Boolean, _
        CheckRemoved As Boolean, CheckingForRemoved As Boolean, _
        CheckingForAdded As Boolean, AddedOrRemoved As Boolean, _
        EventRaised As Boolean, fName As String, _
        x As Integer
    
    'determine which changes we are looking for
    CheckSize = CBool(cChanges And cChangeInFileSize)
    CheckAttrib = CBool(cChanges And cChangeInAttributes)
    CheckAdded = CBool(cChanges And cChangeFilesAdded)
    CheckRemoved = CBool(cChanges And cChangeFilesRemoved)

    
    fnCount = cFilNew.Count
    fcCount = cFilComp.Count

    'we are checking if any files have been removed
    CheckingForRemoved = True
    'test for changes
    For x = 1 To fcCount
        fName = cFilComp.Item(x).FileName
        

        'we want to check for all changes, even if the user hasn't asked for it.
        'If any change ocurs, we want to update out compare list.
        'events will only fire for the requested changes
        
        'check for changes in attributes (also checks for removed files)
        If cFilComp.Item(fName).Attrib <> cFilNew.Item(fName).Attrib Then
            If Not AddedOrRemoved Then
                'raise event, if requested
                If CheckAttrib Then RaiseEvent OnChange(fName, cChangeInAttributes)
                'set flag so we can update the cFilComp later
                EventRaised = True
            End If
        End If

        'check for changes in size
        If cFilComp.Item(fName).Size <> cFilNew.Item(fName).Size Then
            If Not AddedOrRemoved Then
                'raise event, if requested
                If CheckSize Then RaiseEvent OnChange(fName, cChangeInFileSize)
                'set flag so we can update the cFilComp later
                EventRaised = True
            End If
        End If
        
        'reset marker variables.
        'these prevent false events such as
        'cChangeInAttributes, if file is deleted
        'and also prevent cChangeFilesRemoved from occuring twice
        CheckingForRemoved = True
        AddedOrRemoved = False
    Next
    
    'check for new files
    CheckingForRemoved = False
    CheckingForAdded = True
    For x = 1 To fnCount
        fName = cFilNew.Item(x).FileName
        If cFilNew.Item(fName).FileName = cFilComp.Item(fName).FileName Then DoEvents
    Next

    'refresh the comparison list if any changes occured
    If EventRaised Then cFilComp.Refresh
    Exit Sub
ErrorCode:
    If Err.Number = 91 Then
        'if a file is not found, then it has either been added or deleted
        'We can raise the appropriate event depending on our marker variables
        If CheckingForRemoved Then
            If CheckRemoved Then RaiseEvent OnChange(fName, cChangeFilesRemoved)
            CheckingForRemoved = False
        ElseIf CheckingForAdded Then
            If CheckAdded Then RaiseEvent OnChange(fName, cChangeFilesAdded)
        End If
        
        'set flags so we refresh cFilComp later,
        'and so we don't misfire an attribute or size event
        'on the resume command.
        EventRaised = True
        AddedOrRemoved = True
        Resume Next
    End If
    RaiseEvent OnError(Err.Number, Err.Description)
End Sub
