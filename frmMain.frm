VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sanx's Port Scanner"
   ClientHeight    =   4515
   ClientLeft      =   5040
   ClientTop       =   4485
   ClientWidth     =   7470
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4515
   ScaleWidth      =   7470
   Begin MSComctlLib.StatusBar staMain 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   15
      Top             =   4230
      Width           =   7470
      _ExtentX        =   13176
      _ExtentY        =   503
      Style           =   1
      SimpleText      =   "Not Started"
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "Not Started"
            TextSave        =   "Not Started"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton butFTP 
      Caption         =   "Test FTP"
      Height          =   375
      Left            =   2640
      TabIndex        =   22
      Top             =   4440
      Width           =   975
   End
   Begin VB.CommandButton butWeb 
      Caption         =   "Test Web"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1440
      TabIndex        =   21
      Top             =   4440
      Width           =   975
   End
   Begin MSWinsockLib.Winsock sckTest 
      Left            =   6840
      Top             =   4320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton butSMTP 
      Caption         =   "Test SMTP"
      Height          =   375
      Left            =   240
      TabIndex        =   20
      Top             =   4440
      Width           =   975
   End
   Begin VB.TextBox txtDetails 
      BackColor       =   &H00FFFFC0&
      Height          =   2295
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   19
      Top             =   5280
      Width           =   6975
   End
   Begin VB.CommandButton butShow 
      Caption         =   "&Tools"
      Height          =   375
      Left            =   2040
      TabIndex        =   17
      Top             =   3720
      Width           =   735
   End
   Begin VB.Timer tmrMain 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   2880
      Top             =   1320
   End
   Begin MSWinsockLib.Winsock sockets 
      Index           =   0
      Left            =   3360
      Top             =   1320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.ProgressBar barMain 
      Height          =   375
      Left            =   3000
      TabIndex        =   16
      Top             =   3720
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton butExit 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   6000
      TabIndex        =   14
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton butStop 
      Cancel          =   -1  'True
      Caption         =   "S&top"
      Height          =   375
      Left            =   1080
      TabIndex        =   13
      Top             =   3720
      Width           =   735
   End
   Begin VB.CommandButton butStart 
      Caption         =   "&Scan"
      Default         =   -1  'True
      Height          =   375
      Left            =   240
      TabIndex        =   12
      Top             =   3720
      Width           =   855
   End
   Begin VB.ListBox lstPorts 
      BackColor       =   &H00FFFFC0&
      Height          =   1815
      Left            =   240
      TabIndex        =   11
      Top             =   1800
      Width           =   3495
   End
   Begin VB.TextBox txtDisplay 
      BackColor       =   &H00FFFFC0&
      Height          =   3255
      Left            =   3840
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   8
      Top             =   360
      Width           =   3375
   End
   Begin VB.TextBox txtSockets 
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   2640
      TabIndex        =   7
      Text            =   "60"
      Top             =   1080
      Width           =   975
   End
   Begin VB.TextBox txtEnd 
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   1440
      TabIndex        =   5
      Text            =   "135"
      Top             =   1080
      Width           =   975
   End
   Begin VB.TextBox txtStart 
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   240
      TabIndex        =   3
      Text            =   "21"
      Top             =   1080
      Width           =   975
   End
   Begin VB.TextBox txtIP 
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Text            =   "127.0.0.1"
      Top             =   360
      Width           =   3375
   End
   Begin VB.Label Label7 
      Caption         =   "Scan Details:"
      Height          =   255
      Left            =   240
      TabIndex        =   18
      Top             =   4920
      Width           =   1095
   End
   Begin VB.Label Label6 
      Caption         =   "Ports open:"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "Data received from port:"
      Height          =   255
      Left            =   3840
      TabIndex        =   9
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "Sockets to use:"
      Height          =   255
      Left            =   2640
      TabIndex        =   6
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "End Port"
      Height          =   255
      Left            =   1440
      TabIndex        =   4
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Start Port:"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Target IP or Hostname:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This application is copyright Sanx, 2001. This application is
'offered as freeware, and as such, you may copy, modify and
'distribute it without conditions, provided this copyright
'notice remains.
'http://www.sanx.org

Dim isStarted As Boolean
Dim currentPort As Integer
Dim sckUse As Integer
Dim sckCount As Integer
Dim portStart As Integer
Dim portEnd As Integer
Dim scanIP As String
Dim connectData() As String
Dim testString As String
Dim tmrCounts As Integer
Dim inUse As Boolean
Dim twofifty As Integer


Private Sub butExit_Click()

SaveSettings
End

End Sub

Private Sub butFTP_Click()

scanIP = txtIP.Text
butFTP.Enabled = False
txtDetails.Text = ""

With sckTest
    .RemoteHost = scanIP
    .RemotePort = 21
    .Connect
End With

End Sub

Private Sub butShow_Click()

If butShow.Caption = "&Tools" Then
    butShow.Caption = "&Hide"
    frmMain.Height = 8310
Else
    butShow.Caption = "&Tools"
    frmMain.Height = 4890
End If

End Sub

Private Sub butSMTP_Click()

scanIP = txtIP.Text
butSMTP.Enabled = False
txtDetails.Text = ""

With sckTest
    .RemoteHost = scanIP
    .RemotePort = 25
    .Connect
End With
twofifty = 0

End Sub


Private Sub butStart_Click()

isStarted = True
butStart.Enabled = False
lstPorts.Clear
ReDim connectData(1)
ScanPorts

End Sub

Private Sub butStop_Click()

isStarted = False

End Sub

Private Sub ScanPorts()

Dim counter As Integer

sckUse = Val(txtSockets.Text)
portStart = Val(txtStart.Text)
portEnd = Val(txtEnd.Text)
scanIP = txtIP.Text
tmrCounts = 0

If sckUse > portEnd - portStart Then sckUse = portEnd - portStart

For counter = 1 To sckUse
    Load sockets(counter)
Next

For currentPort = portStart To (portStart + sckUse)
    DoEvents
    With sockets(currentPort - portStart)
        .RemoteHost = scanIP
        .RemotePort = (currentPort)
        .Connect
    End With
    StatusUpdate currentPort, "Scanning port:"
    barMain.Value = funPercent()
Next

End Sub

Private Sub Form_Load()

txtIP.Text = GetSetting("Sanx's Port Scanner", "Settings", "LastIP", "127.0.0.1")
txtStart.Text = GetSetting("Sanx's Port Scanner", "Settings", "Start", "21")
txtEnd.Text = GetSetting("Sanx's Port Scanner", "Settings", "End", "135")
txtSockets.Text = GetSetting("Sanx's Port Scanner", "Settings", "Sockets", "60")
frmMain.Top = GetSetting("Sanx's Port Scanner", "Settings", "Top", ((Screen.Width - frmMain.Width) / 2))
frmMain.Left = GetSetting("Sanx's Port Scanner", "Settings", "Left", ((Screen.Height - frmMain.Height) / 2))

testString = "TestData" + vbCrLf + "TestData" + vbCrLf + vbCrLf + vbCrLf

End Sub

Private Sub Form_Unload(Cancel As Integer)

SaveSettings

End Sub

Private Sub SaveSettings()

SaveSetting "Sanx's Port Scanner", "Settings", "LastIP", txtIP.Text
SaveSetting "Sanx's Port Scanner", "Settings", "Start", txtStart.Text
SaveSetting "Sanx's Port Scanner", "Settings", "End", txtEnd.Text
SaveSetting "Sanx's Port Scanner", "Settings", "Sockets", txtSockets.Text
SaveSetting "Sanx's Port Scanner", "Settings", "Top", frmMain.Top
SaveSetting "Sanx's Port Scanner", "Settings", "Left", frmMain.Left

End Sub

Private Sub StatusUpdate(portScan As Integer, actionScan As String)

staMain.SimpleText = actionScan + Str$(portScan)

End Sub

Private Sub lstPorts_Click()

txtDisplay.Text = connectData(lstPorts.ListIndex)

End Sub

Private Sub sckTest_DataArrival(ByVal bytesTotal As Long)

Dim tempVar As String

sckTest.GetData tempVar
txtDetails.Text = txtDetails.Text + tempVar

Select Case sckTest.RemotePort
    Case 21
        Select Case Left$(tempVar, 3)
            Case "220"
                sckTest.SendData "USER anonymous" + vbCrLf
            Case "331"
                sckTest.SendData "PASS test@test.com" + vbCrLf
            Case "530"
                sckTest.SendData "QUIT" + vbCrLf
                txtDetails.Text = txtDetails.Text + "QUIT Sent. Anonymous login disabled." + vbCrLf
            Case "230"
                sckTest.SendData "QUIT" + vbCrLf
                txtDetails.Text = txtDetails.Text + "QUIT Sent. Anonymous login enabled." + vbCrLf
            Case "221"
                sckTest.Close
                butFTP.Enabled = True
        End Select
    Case 25
        Select Case Left$(tempVar, 3)
            Case "220"
                sckTest.SendData "HELO SP test@test.com" + vbCrLf
            Case "250"
                Select Case twofifty
                    Case 0
                        sckTest.SendData "MAIL FROM: test@test.com" + vbCrLf
                    Case 1
                        sckTest.SendData "RCPT TO: test2@test2.com" + vbCrLf
                    Case 2
                        sckTest.SendData "DATA" + vbCrLf
                    Case 3
                        sckTest.SendData vbCrLf + "." + vbCrLf + "QUIT" + vbCrLf
                    Case Else
                End Select
                twofifty = twofifty + 1
            Case "354"
                sckTest.SendData vbCrLf + "." + vbCrLf + "QUIT" + vbCrLf
                txtDetails.Text = txtDetails.Text + "QUIT Sent. Relaying allowed." + vbCrLf
            Case "550"
                sckTest.SendData "QUIT" + vbCrLf
                txtDetails.Text = txtDetails.Text + "QUIT Sent. Relaying prohibited." + vbCrLf
                sckTest.Close
                butSMTP.Enabled = True
            Case "221", "554"
                sckTest.Close
                butSMTP.Enabled = True
            Case "421", "500", "501", "503"
                sckTest.SendData "QUIT" + vbCrLf
                sckTest.Close
                butSMTP.Enabled = True
        End Select
End Select

End Sub

Private Sub sockets_Connect(Index As Integer)

lstPorts.AddItem Str$(sockets(Index).RemotePort)
ReDim Preserve connectData(lstPorts.ListCount)
StatusUpdate sockets(Index).RemotePort, "Connected on and sending data to port:"
DoEvents
Select Case sockets(Index).RemotePort
    Case 25
        testString = ""
    Case 80
        testString = "GET / HTTP/1.0" + vbCrLf + "Accept: image/gif, image/x-xbitmap, image/jpeg, */*" + vbCrLf + "Accept -Language: en" + vbCrLf + "User-Agent: Mozilla/1.22 (compatible; MSIE 2.0d; Windows NT)" + vbCrLf + "Connection: Keep -Alive" + vbCrLf + vbCrLf
    Case Else
        testString = "TestData" + vbCrLf + "TestData" + vbCrLf + vbCrLf + vbCrLf
End Select
If testString <> "" Then sockets(Index).SendData testString

End Sub

Private Sub sockets_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim listEntry As Integer
Dim count As Integer

For count = 0 To (lstPorts.ListCount - 1)
    If lstPorts.List(count) = Str$(sockets(Index).RemotePort) Then
        sockets(Index).GetData connectData(count)
    End If
Next
DoEvents
StatusUpdate sockets(Index).RemotePort, "Data received on port:"
sockets(Index).Close
ContScan Index

End Sub

Private Sub sockets_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)

sockets(Index).Close
ContScan Index

End Sub

Private Sub ContScan(freeSck As Integer)

DoEvents

If currentPort < portEnd And isStarted = True Then
    currentPort = currentPort + 1
    With sockets(freeSck)
        .RemotePort = currentPort
        .Connect
    End With
    StatusUpdate currentPort, "Scanning port:"
    barMain.Value = funPercent()
Else
    staMain.SimpleText = "Stopping Scan..."
    If Not tmrMain.Enabled Then tmrMain.Enabled = True
End If

End Sub

Private Sub tmrMain_Timer()

CheckSockets

End Sub

Private Sub CheckSockets()

On Error GoTo ErrorHandler

Dim count As Integer
Dim notFinished As Boolean

notFinished = False

For count = 0 To sckUse
    If sockets(count).State <> 0 Then notFinished = True
    DoEvents
Next

If notFinished = False Or tmrCounts = 20 Then
    tmrMain.Enabled = False
    For count = 1 To sckUse
        sockets(count).Close
        Unload sockets(count)
    Next
    butStart.Enabled = True
    barMain.Value = 100
    isStarted = False
    staMain.SimpleText = "Finished Scan. All sockets closed"
End If


tmrCounts = tmrCounts + 1

ErrorHandler:
    Resume Next

End Sub
Private Function funPercent()

funPercent = Int((currentPort - portStart) / (portEnd - portStart) * 100)
End Function
