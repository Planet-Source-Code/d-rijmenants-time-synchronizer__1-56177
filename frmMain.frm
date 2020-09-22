VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " Time Sync"
   ClientHeight    =   3945
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5595
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   5595
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   4320
      TabIndex        =   1
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Height          =   30
      Left            =   0
      TabIndex        =   8
      Top             =   3360
      Width           =   5535
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   -120
      TabIndex        =   3
      Top             =   -120
      Width           =   6015
      Begin VB.Image Image1 
         Height          =   780
         Left            =   360
         Picture         =   "frmMain.frx":0442
         Top             =   240
         Width           =   690
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Time Server Synchronization"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   1440
         TabIndex        =   4
         Top             =   480
         Width           =   3855
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   250
      Left            =   120
      Top             =   3480
   End
   Begin VB.ComboBox cboServers 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   2880
      Width           =   2535
   End
   Begin VB.CommandButton cmdSynchTime 
      Caption         =   "&Synchronize"
      Height          =   375
      Left            =   2640
      TabIndex        =   0
      Top             =   3480
      Width           =   1575
   End
   Begin VB.Image imgWarn 
      Height          =   275
      Left            =   120
      Picture         =   "frmMain.frx":20F4
      Stretch         =   -1  'True
      Top             =   2225
      Visible         =   0   'False
      Width           =   275
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Computer System Time"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Last Synchronization"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1875
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Time server"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   2640
      Width           =   2415
   End
   Begin VB.Label lblStatus 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Press the Synchronize Button to adjust the Computer System Time."
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   2280
      Width           =   5055
   End
   Begin VB.Label lblDateTime 
      BackStyle       =   0  'Transparent
      Caption         =   "None"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1800
      TabIndex        =   6
      Top             =   1875
      Width           =   3615
   End
   Begin VB.Label lblPcTime 
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1800
      TabIndex        =   5
      Top             =   1440
      Width           =   3615
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboServers_Click()
TIME_SERVER = Me.cboServers.Text
ServerIndex = Me.cboServers.ListIndex
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub Form_Activate()

If BatchMode = True Then Call cmdSynchTime_Click

End Sub

Private Sub Form_Load()

Call LoadServers

TIME_SERVER = Me.cboServers.Text
Me.lblDateTime = GetSetting(App.EXEName, "Timeserver", "Last", "None")

'If BatchMode = True Then Call cmdSynchTime_Click

End Sub

Private Sub cmdSynchTime_Click()
Dim Echeck As Boolean
    Dim oInetTime As cInetTime
    Set oInetTime = New cInetTime
    
    Me.Timer1.Enabled = False
    Me.imgWarn.Visible = False
    lblStatus.Caption = "Connecting to Server..."
    cmdSynchTime.Enabled = False

    With oInetTime
        .TimeServer = TIME_SERVER
        .SetTime
        If .ErrorCheck = True Then
            Echeck = True
            lblStatus.Caption = "          No response from Time Server."
            Me.imgWarn.Visible = True
            
        Else
            lblStatus.Caption = "Adjusted " & .AdjustedSecs & " Sec"
            lblDateTime.Caption = ":  " & Format(.ReturnedDate, "dddd dd mmm yyyy  hh:mm:ss")
            lblDateTime.Visible = True
            Echeck = False
            Me.imgWarn.Visible = False
            
        End If
        
    End With
    
    cmdSynchTime.Enabled = True
    cmdSynchTime.SetFocus
    
    Set oInetTime = Nothing
    
    If BatchMode = True And Echeck = False Then Unload Me
    
    BatchMode = False
    Me.Timer1.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)

    SaveSetting App.EXEName, "Timeserver", "Name", TIME_SERVER
    SaveSetting App.EXEName, "Timeserver", "Index", Trim(Str(ServerIndex))
    SaveSetting App.EXEName, "Timeserver", "Last", Me.lblDateTime.Caption
    End

End Sub

Private Sub Timer1_Timer()
Me.lblPcTime.Caption = ":  " & Format(Now, "dddd dd mmm yyyy  hh:mm:ss")
End Sub

Private Sub LoadServers()
Dim InFile
Dim ServerName As String
Dim remPos As Integer

InFile = FreeFile

On Error GoTo LoadDefaults
Open App.Path & "\SERVERS.txt" For Input As InFile
    While Not EOF(InFile)
        'get next line
        Line Input #InFile, ServerName
        If Left(ServerName, 1) <> ";" And Trim(ServerName) <> "" Then 'skip remarques and empty's
            remPos = InStr(1, ServerName, ";")
            If remPos <> 0 Then ServerName = Trim(Left(ServerName, remPos - 1))
            cboServers.AddItem ServerName
            End If
    Wend
Close InFile
If ServerIndex > cboServers.ListCount - 1 Then ServerIndex = 0
cboServers.ListIndex = ServerIndex
Exit Sub

LoadDefaults:
Close InFile

    With cboServers
        .AddItem "ntps1-0.cs.tu-berlin.de"
        .AddItem "ntps1-0.uni-erlangen.de"
        .AddItem "ntps1-1.uni-erlangen.de"
        .AddItem "ntps1-2.uni-erlangen.de"
        .AddItem "ptbtime1.ptb.de"
        .AddItem "ptbtime2.ptb.de"
        .AddItem "tick.usno.navy.mil"
        .AddItem "timex.cs.columbia.edu"
        .AddItem "nist1.datum.com"
        .AddItem "time.ien.it"
        .AddItem "swisstime.ethz.ch"
        .AddItem "ntp.lth.se"
        .AddItem "Rolex.PeachNet.edu"
        If ServerIndex > .ListCount - 1 Then ServerIndex = 0
        .ListIndex = ServerIndex
    End With

End Sub

