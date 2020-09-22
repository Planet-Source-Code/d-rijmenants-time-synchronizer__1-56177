VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Private mstrDestination  As String

Public Property Get Destination() As String
    Destination = mstrDestination
End Property

Public Property Let Destination(ByVal sNewValue As String)
    mstrDestination = sNewValue
End Property
    
Public Sub StartLog()
    Dim iFileNumber As Integer
        
    iFileNumber = FreeFile
    Open mstrDestination For Append As #iFileNumber
    Print #iFileNumber, "===================================="
    Print #iFileNumber, "LogFile Started: " & Now()
    Close #iFileNumber
    
End Sub

Public Sub CloseLog()
    Dim iFileNumber As Integer
    
    iFileNumber = FreeFile
    Open mstrDestination For Append As #iFileNumber
    Print #iFileNumber, "LogFile Ended: " & Now()
    Print #iFileNumber, "===================================="
    Close #iFileNumber
    
End Sub

Public Sub LogEntry(strMessage As String)
    
    Dim iFileNumber As Integer
       
    iFileNumber = FreeFile
    Open mstrDestination For Append As #iFileNumber
    Print #iFileNumber, strMessage
    Close #iFileNumber
        
End Sub

Private Sub Class_Terminate()
    Close
End Sub
