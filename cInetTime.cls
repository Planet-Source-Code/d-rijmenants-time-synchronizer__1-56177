VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cInetTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Private oForm               As frmWinsock
Private WithEvents oWinsock As Winsock
Attribute oWinsock.VB_VarHelpID = -1

Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Private Type TIME_ZONE_INFORMATION
    Bias As Long
    StandardName(31) As Integer
    StandardDate As SYSTEMTIME
    StandardBias As Long
    DaylightName(31) As Integer
    DaylightDate As SYSTEMTIME
    DaylightBias As Long
End Type

Private m_strTimeServer     As String   'Time Server that you are polling
Private m_dtReturnedDate    As Date     'Date & Time returned from Time Server
Private m_lnglAdjustedSecs  As Long     'Seconds adjusted from local PC's time
Private m_lngTimeout        As Long     'Time to give up on getting data back from Time Server
Private m_lngRawReturn      As Long     'Raw Value returned by the Time Server
Private m_blnError          As Boolean  'Did we encounter an error?

Private Declare Function ntohl Lib "WSOCK32.DLL" (ByVal hostlong As Long) As Long
Private Declare Function GetTimeZoneInformation Lib "KERNEL32.DLL" (lpTZI As TIME_ZONE_INFORMATION) As Long

Public Property Get ErrorCheck() As Boolean
    ErrorCheck = m_blnError
End Property

Public Property Get AdjustedSecs() As Long
    AdjustedSecs = m_lnglAdjustedSecs
End Property

Public Property Get ReturnedDate() As Date
Dim strT As String
    ReturnedDate = m_dtReturnedDate
End Property

Public Property Let TimeServer(NewValue As String)
    m_strTimeServer = NewValue
End Property

Private Sub Class_Initialize()

    Set oForm = New frmWinsock
    Set oWinsock = oForm.Winsock1
    m_lngTimeout = 5  'Timeout value
    m_blnError = False
    
End Sub

Private Sub Class_Terminate()
    
    Set oWinsock = Nothing
    Unload oForm
    Set oForm = Nothing

End Sub


Private Sub oWinsock_DataArrival(ByVal bytesTotal As Long)
    Dim dReturn     As Double   ' The returned value as a double
    Dim fok         As Boolean
    Dim sLog        As String

    ' Get the returned vbLong value
    ' and assign it to the m_lngRawReturn variable

    oWinsock.GetData m_lngRawReturn, vbLong

    ' Close the connection

    oWinsock.Close
    
    ' NOTE:  The returned value is a signed long, however,
    '        it is in Network Byte Order.  We need to convert
    '        it to Host Byte Order using the ntohl() function
    '        call provided by the WSOCK32.DLL.  This function
    '        call will return the value as an unsigned long.

    m_lngRawReturn = ntohl(m_lngRawReturn)

    ' Convert the returned value from an unsigned long to a
    ' signed double.

    If m_lngRawReturn < 0 Then
        dReturn = m_lngRawReturn + 4294967296#
    Else
        dReturn = m_lngRawReturn
    End If

    ' Now we need to obtain the GMT Bias as well as the
    ' DaylightBias and adjust the time to represent
    ' the time as configured in Windows.

    Dim vbTZI As TIME_ZONE_INFORMATION

    m_lngRawReturn = GetTimeZoneInformation(vbTZI)

    ' Evaluate the returned value and adjust the time
    ' to compensate for LocalTime offsets.

    Select Case m_lngRawReturn
        Case 0  ' DST Not Recognized
            dReturn = dReturn - (vbTZI.Bias * 60)
        Case 1  ' DST Recognized but not in effect
            dReturn = dReturn - (vbTZI.Bias * 60)
        Case 2  ' DST Recognized and in effect
            dReturn = dReturn - ((vbTZI.Bias * 60) + (vbTZI.DaylightBias * 60))
        Case Else   ' Error, assume DST not in effect
            dReturn = dReturn - (vbTZI.Bias * 60)
    End Select
    ' The following two lines are used to determine how
    ' much your Date/Time needs adjustment.  It will be
    ' used later on in this subroutine.
    'Dim AdjustAmount As Long
    m_lnglAdjustedSecs = dReturn - DateDiff("s", "01 Jan 1900 00:00:00", Now)
    ' Adjust for the Windows Time/Time Protocol time reference
    ' difference and convert the dReturn value so that it
    ' represents the number of days instead of the number
    ' of seconds.  The returned value is a double.
    dReturn = (dReturn + 172800) / 86400
    ' Convert from a double to a date
    m_dtReturnedDate = dReturn
    ' The following lines will adjust your Windows clock
    ' (if necessary) to the Date/Time received from the
    ' Time Protocol Server.
    If m_lnglAdjustedSecs = 0 Then
        'Our PC's time is correct.  Not need to change time.
    Else
        'Set System Date & Time
         Date = Format(m_dtReturnedDate, "dd/mm/yy")
         Time = Format(m_dtReturnedDate, "hh:mm:ss ")
    End If
End Sub

Private Sub oWinsock_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)

    m_blnError = True
    oWinsock.Close

End Sub

Public Sub SetTime()

    With oWinsock
        .RemoteHost = m_strTimeServer
        .Protocol = sckTCPProtocol
        .RemotePort = 37    ' Time Protocol
        .LocalPort = 0
        
        ' Connect to the Time Server
        .Connect
                
        WaitForResponse
    End With

End Sub

Private Sub WaitForResponse()
    'Wait a few seconds so that we have time to get
    'a value back from Time Server
    
    Dim Start     As Single
    
    Start = Timer                       ' Set start time.
    Do While Timer < Start + m_lngTimeout
        DoEvents                        ' Yield to other processes.
        If m_lngRawReturn <> 0 Then     ' Exit loop when server returns any data
            Exit Do
        End If
    Loop
    
End Sub

