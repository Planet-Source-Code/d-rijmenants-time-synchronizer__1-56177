Attribute VB_Name = "modMain"
Option Explicit

'Change to your closest Internet Time Server
'Public Const TIME_SERVER As String = "ntps1-0.cs.tu-berlin.de" ' "Rolex.PeachNet.edu"
Public TIME_SERVER As String '= "ntps1-0.cs.tu-berlin.de" ' "Rolex.PeachNet.edu"
Public ServerIndex As Integer
Public BatchMode As Boolean

Sub Main()

If App.PrevInstance Then End

TIME_SERVER = GetSetting(App.EXEName, "Timeserver", "Name", "ntps1-0.cs.tu-berlin.de")
ServerIndex = Val(GetSetting(App.EXEName, "Timeserver", "Index", "1"))

        If Command() = "/now" Or Command() = "/NOW" Then
            BatchMode = True
            End If
        
        'Open Form and let user use the interface
        Dim oForm As frmMain
        Set oForm = New frmMain
        oForm.WindowState = vbNormal
        oForm.Show 'vbModal
        oForm.Refresh
    
End Sub

Public Function dtString() As String
    'Quick Date and Time string w/full 4 digit year
    dtString = Format(Now, "dd/mm/yyyy hh:mm:ss") 'AM/PM -
End Function

Public Sub SetTime(sTimeServer As String)
    Dim oInetTime As cInetTime
    Set oInetTime = New cInetTime
        
    oInetTime.TimeServer = sTimeServer
    oInetTime.SetTime
    
    Set oInetTime = Nothing

End Sub

