VERSION 5.00
Begin VB.Form frmTest 
   Caption         =   "Directory Monitor"
   ClientHeight    =   3690
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5100
   LinkTopic       =   "Form1"
   ScaleHeight     =   3690
   ScaleWidth      =   5100
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check2 
      Caption         =   "Files Attributes Changed"
      Height          =   195
      Index           =   3
      Left            =   3060
      TabIndex        =   6
      Top             =   3480
      Value           =   1  'Checked
      Width           =   2055
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Files Size Changed"
      Height          =   195
      Index           =   2
      Left            =   3060
      TabIndex        =   5
      Top             =   3240
      Value           =   1  'Checked
      Width           =   1875
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Files Removed"
      Height          =   195
      Index           =   1
      Left            =   1500
      TabIndex        =   4
      Top             =   3480
      Value           =   1  'Checked
      Width           =   1515
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Files Added"
      Height          =   195
      Index           =   0
      Left            =   1500
      TabIndex        =   3
      Top             =   3240
      Value           =   1  'Checked
      Width           =   1515
   End
   Begin VB.ListBox List1 
      Height          =   2400
      ItemData        =   "frmTest.frx":0000
      Left            =   60
      List            =   "frmTest.frx":0002
      TabIndex        =   2
      Top             =   780
      Width           =   4995
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Enabled"
      Height          =   255
      Left            =   60
      TabIndex        =   1
      Top             =   3240
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Text            =   "c:\documents and settings\alan\desktop\"
      Top             =   300
      Width           =   5055
   End
   Begin VB.Label Label2 
      Caption         =   "Directory to watch:"
      Height          =   195
      Left            =   60
      TabIndex        =   8
      Top             =   0
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Changes:"
      Height          =   255
      Left            =   60
      TabIndex        =   7
      Top             =   600
      Width           =   1215
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents dMon As DirMonDll.Monitor
Attribute dMon.VB_VarHelpID = -1

Private Sub Check1_Click()
    dMon.Path = Text1.Text
    dMon.Enabled = CBool(Check1.Value)
End Sub

Private Sub Check2_Click(Index As Integer)
    Dim NewChangeType As cEnumChangeType, x As Integer
    
    NewChangeType = NewChangeType + (cChangeFilesAdded * Check2(0).Value)
    NewChangeType = NewChangeType + (cChangeFilesRemoved * Check2(1).Value)
    NewChangeType = NewChangeType + (cChangeInFileSize * Check2(2).Value)
    NewChangeType = NewChangeType + (cChangeInAttributes * Check2(3).Value)
    
    dMon.ChangeType = NewChangeType

End Sub

Private Sub dMon_OnChange(FileName As String, ChangeType As DirMonDll.cEnumChangeType)
    Select Case ChangeType
        Case cChangeFilesAdded
            List1.AddItem FileName & " Added"
        Case cChangeFilesRemoved
            List1.AddItem FileName & " Removed"
        Case cChangeInAttributes
            List1.AddItem FileName & " Attributes Changed"
        Case cChangeInFileSize
            List1.AddItem FileName & " Size Changed"
    End Select
   
    
    
End Sub

Private Sub dMon_OnTimer()
    Caption = "Working!"
End Sub



Private Sub Form_Load()
    Set dMon = New DirMonDll.Monitor
End Sub
