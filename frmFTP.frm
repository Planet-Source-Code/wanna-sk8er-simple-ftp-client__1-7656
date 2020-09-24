VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmFTP 
   AutoRedraw      =   -1  'True
   Caption         =   "ftp Client"
   ClientHeight    =   6015
   ClientLeft      =   2340
   ClientTop       =   390
   ClientWidth     =   6360
   Icon            =   "frmFTP.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   6360
   Begin VB.CommandButton cmdConnect 
      Caption         =   "&Connect"
      Height          =   495
      Left            =   120
      TabIndex        =   12
      Top             =   1020
      Width           =   1095
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "    C&lose Connection"
      Enabled         =   0   'False
      Height          =   495
      Left            =   120
      TabIndex        =   11
      Top             =   3612
      Width           =   1095
   End
   Begin VB.CommandButton cmdDirectory 
      Caption         =   "Get &Directory"
      Enabled         =   0   'False
      Height          =   495
      Left            =   120
      TabIndex        =   10
      Top             =   1668
      Width           =   1095
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   120
      TabIndex        =   9
      Top             =   4260
      Width           =   1095
   End
   Begin VB.CommandButton cmdReadText 
      Caption         =   "  Read Text            File"
      Enabled         =   0   'False
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   2316
      Width           =   1095
   End
   Begin VB.CommandButton cmdDownLoad 
      Caption         =   "  Download           File"
      Enabled         =   0   'False
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   2964
      Width           =   1095
   End
   Begin VB.ListBox lstDir 
      Height          =   3765
      ItemData        =   "frmFTP.frx":000C
      Left            =   1440
      List            =   "frmFTP.frx":000E
      TabIndex        =   6
      Top             =   960
      Width           =   4695
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   1320
      Top             =   4920
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      Protocol        =   2
      RemotePort      =   21
      URL             =   "ftp://anonymous:mesmer@ix.netcom.com@"
      UserName        =   "anonymous"
      Password        =   "mesmer@ix.netcom.com"
   End
   Begin VB.TextBox txtDirectory 
      Height          =   285
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   600
      Width           =   4575
   End
   Begin VB.TextBox txtCommand 
      Height          =   855
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   5040
      Width           =   6075
   End
   Begin VB.TextBox txtServer 
      Height          =   285
      Left            =   1440
      TabIndex        =   0
      Top             =   240
      Width           =   4575
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Host:"
      Height          =   195
      Left            =   0
      TabIndex        =   5
      Top             =   240
      Width           =   375
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Working Directory:"
      Height          =   195
      Left            =   0
      TabIndex        =   4
      Top             =   600
      Width           =   1320
   End
   Begin VB.Label lblWhat 
      AutoSize        =   -1  'True
      Caption         =   "lblWhat.Caption"
      Height          =   195
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   1125
   End
End
Attribute VB_Name = "frmFTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public sCommand As String
Public sServer As String
Public iAction As Integer
' Download Directory Path
Const DownLoadDir As String = "C:\Temp\"
' Activity Constants
Const intntConnect As Integer = 0
Const intntChangeDir As Integer = 1
Const intntGetText As Integer = 2
Const intntGetFile As Integer = 3
Const intntGetDirName As Integer = 4

Sub FillDir(strData As String)
' Parses strData and fills lstDir with
' directory and file names
Dim iCR As Integer
Dim iStart As Integer: iStart = 1
Do
    ' Find vbCrLf
    iCR = InStr(iStart, strData, vbCrLf)
    ' If not found, exit loop
    If iCR = 0 Then Exit Do
    ' Found, add next item to listbox
    lstDir.AddItem Mid(strData, iStart, iCR - iStart)
    ' Skip start position over vbCrLf (2 characters)
    iStart = iCR + 2
Loop
End Sub

Sub subConnected()
cmdDirectory.Enabled = True
cmdDownLoad.Enabled = True
cmdReadText.Enabled = True
cmdClose.Enabled = True
cmdConnect.Enabled = False
End Sub

Sub subDisConnected()
cmdDirectory.Enabled = False
cmdDownLoad.Enabled = False
cmdReadText.Enabled = False
cmdClose.Enabled = False
cmdConnect.Enabled = True
End Sub

Public Sub subGetDir()
'   Reads directory from host after connection
'   is established
'   called from inet control state changed event
Dim vtData As Variant ' Data variable.
Dim strData As String: strData = ""
Dim bDone As Boolean: bDone = False
' Get first chunk.
vtData = Inet1.GetChunk(1024, icString)
'   continue until no data is in the buffer
Do While Not bDone
    lblWhat.Caption = "Reading DIR "
    DoEvents
    strData = strData & vtData
    ' Get next chunk.
    vtData = Inet1.GetChunk(1024, icString)
    lblWhat.Caption = "Reading DIR "
    DoEvents
    If Len(vtData) = 0 Then
        bDone = True
    End If
Loop
If iAction = intntGetDirName Then
    txtDirectory = strData
ElseIf iAction = intntConnect Then
    lstDir.Clear
    FillDir (strData)
ElseIf iAction = intntChangeDir Then
    lstDir.Clear
    lstDir.AddItem ".."
    FillDir (strData)
End If
lblWhat.Caption = "Directory Done"
End Sub

Private Sub cmdClose_Click()
    On Error Resume Next
    Inet1.Execute , "CANCEL"
    Inet1.Execute , "CLOSE"
    Call subDisConnected
End Sub

Private Sub cmdConnect_Click()
On Error GoTo Error_Handler
'   Can't open a non-existant server
If sServer = "" Then
    MsgBox "Must enter server name", vbOKOnly, "Note:"
    Exit Sub
End If
'   OK, open the server and get the directory
With Inet1
    '   Display command being sent
    txtCommand = txtCommand & vbCrLf & "PWD"
    txtCommand.SelStart = Len(txtCommand.Text)
    .URL = sServer
    iAction = intntGetDirName
    .Execute , "PWD"
    '   Wait until done before doing anything else
    Do While Inet1.StillExecuting
        DoEvents
    Loop
    iAction = intntConnect
    '   Display command being sent
    txtCommand = txtCommand & vbCrLf & "LS"
    txtCommand.SelStart = Len(txtCommand.Text)
    .Execute , "LS"
    '   Wait until done before doing anything else
    Do While Inet1.StillExecuting
        DoEvents
    Loop
    Call subConnected
End With
Exit_Normal:
    Exit Sub
Error_Handler:
    MsgBox Err & " " & Error, vbOKOnly, "Connect Error"
    Select Case Err
        Case 35761: Inet1.Cancel    ' Timeout
        Case 35764: Inet1.Cancel    ' Still Executing
    End Select
    Resume Exit_Normal
End Sub

    
Sub cmdDirectory_Click()
On Error GoTo Error_Handler
'   Change directory on remote computer
Dim sDirectory As String
sDirectory = lstDir.List(lstDir.ListIndex)
'   Test for no directory selected
If sDirectory = "" Then
    MsgBox "No directory selected.", vbOKOnly, "Note:"
    Exit Sub
End If
'   Test for file name selected
If sDirectory <> ".." And Right$(sDirectory, 1) <> "/" Then
    MsgBox "File selected.", vbOKOnly, "Note:"
    Exit Sub
End If
'   OK, there is something selected
iAction = intntChangeDir
If sDirectory = ".." Then
    '   Move to root directory
    '   Display activity
    lblWhat = "Moving to the root directory"
    Call subBusy(True)
    DoEvents
    With Inet1
        .URL = sServer
        '   Display command being sent
        txtCommand = txtCommand & vbCrLf & "CD .."
        txtCommand.SelStart = Len(txtCommand.Text)
        sCommand = "CD .."
        .Execute , sCommand
        '   Wait until done before doing anything else
        Do While Inet1.StillExecuting
            DoEvents
        Loop
        '   Display command being sent
        txtCommand = txtCommand & vbCrLf & "PWD"
        txtCommand.SelStart = Len(txtCommand.Text)
        sCommand = "PWD"
        iAction = intntGetDirName
        .Execute , sCommand
        '   Wait until done before doing anything else
        Do While Inet1.StillExecuting
            DoEvents
        Loop
        '   Display command being sent
        txtCommand = txtCommand & vbCrLf & "LS"
        txtCommand.SelStart = Len(txtCommand.Text)
        sCommand = "LS"
        iAction = intntChangeDir
        .Execute , sCommand
        '   Wait until done before doing anything else
        Do While Inet1.StillExecuting
            DoEvents
        Loop
    End With
'   End of CDUP routine
Else
    '   Move to new directory
    '   Display activity
    lblWhat = "Moving to " & sDirectory
    Call subBusy(True)
    DoEvents
    With Inet1
        .URL = sServer
        sCommand = "CD " & txtDirectory
        If Right$(sCommand, 1) <> "/" Then
            sCommand = sCommand & "/" & sDirectory
        Else
            sCommand = sCommand & sDirectory
        End If
        '   Display command being sent
        txtCommand = txtCommand & vbCrLf & sCommand
        txtCommand.SelStart = Len(txtCommand.Text)
        iAction = intntChangeDir
        .Execute , sCommand
        '   Wait until done before doing anything else
        Do While Inet1.StillExecuting
            DoEvents
        Loop
        iAction = intntGetDirName
        '   Display command being sent
        txtCommand = txtCommand & vbCrLf & "PWD"
        txtCommand.SelStart = Len(txtCommand.Text)
        txtCommand = txtCommand & vbCrLf & "PWD"
        txtCommand.SelStart = Len(txtCommand.Text)
        .Execute , "PWD"
        '   Wait until done before doing anything else
        Do While Inet1.StillExecuting
            DoEvents
        Loop
        iAction = intntChangeDir
        '   Display command being sent
        txtCommand = txtCommand & vbCrLf & "LS"
        txtCommand.SelStart = Len(txtCommand.Text)
        .Execute , "LS"
        '   Wait until done before doing anything else
        Do While Inet1.StillExecuting
            DoEvents
        Loop
    End With
'   Directory read
End If
Exit_Normal:
    Call subBusy(False)
    Exit Sub
Error_Handler:
    MsgBox Err & " " & Error, vbOKOnly, "Directory Error"
    Select Case Err
        Case 35761: Inet1.Cancel    ' Timeout
        Case 35764: Inet1.Cancel    ' Still Executing
    End Select
    Resume Exit_Normal
End Sub

Private Sub cmdDownLoad_Click()
'   Download the selected file
Dim sFileName As String, sTemp As String
Dim sDLName As String, iResponse As Integer
On Error GoTo Error_Handler
If lstDir.List(lstDir.ListIndex) = "" Then
    MsgBox "No file selected.", vbOKOnly, "Note:"
    Exit Sub
    '   No file selected, exit this procedure
End If
If Right$(lstDir.List(lstDir.ListIndex), 1) = "/" Then
    MsgBox "Directory selected.", vbOKOnly, "Note:"
    Exit Sub
    '   Not a file -- cannot download
End If
'   build the path and filename string
If txtDirectory = "/" Then
    '   Root directory
    sFileName = txtDirectory & lstDir.List(lstDir.ListIndex)
ElseIf Right$(txtDirectory, 1) = "/" Then
    '   has slash. Build path/filename
    sFileName = txtDirectory & lstDir.List(lstDir.ListIndex)
Else
    '   no slash, add it to build path/filename
    sFileName = txtDirectory & "/" & lstDir.List(lstDir.ListIndex)
End If
sDLName = lstDir.List(lstDir.ListIndex)
If Len(Dir(DownLoadDir & sDLName)) Then
    '   File exists in download directory
    iResponse = MsgBox(sDLName & "already in " & DownLoadDir & _
      vbCrLf & "OverWrite?", vbOKCancel, "Warning!")
    ' No? Then exit sub
    If iResponse = vbCancel Then Exit Sub
    ' Yes, then delete old
    sTemp = DownLoadDir & sDLName
    Kill sTemp
End If
Call subBusy(True)
'   Start the download
With Inet1
    sCommand = "GET " & sFileName & " " & DownLoadDir & _
      sDLName
    '   Display command being sent
    txtCommand = txtCommand & vbCrLf & sCommand
    txtCommand.SelStart = Len(txtCommand.Text)
    .Execute , sCommand
    '   Wait until done before doing anything else
    Do While Inet1.StillExecuting
        lblWhat.Font.Bold = Not lblWhat.Font.Bold
        lblWhat.Caption = "Downloading " & lstDir.List(lstDir.ListIndex)
        DoEvents
    Loop
    lblWhat.Caption = "Done Downloading"
    Call subBusy(False)
    Beep
End With
Exit_Normal:
    Exit Sub
Error_Handler:
    MsgBox Err & " " & Error, vbOKOnly, "Download Error"
    Select Case Err
        Case 35761: Inet1.Cancel    ' Timeout
        Case 35764: Inet1.Cancel    ' Still Executing
    End Select
    Resume Exit_Normal
End Sub

Private Sub cmdQuit_Click()
'   Exit from the program
Unload Me
End
End Sub

Private Sub cmdReadText_Click()
On Error GoTo Error_Handler
'    Read a text file -- display it on frmText
Dim sFileName As String
If lstDir.List(lstDir.ListIndex) = "" Then
    MsgBox "No file selected.", vbOKOnly, "Note:"
    Exit Sub
End If
If Right$(lstDir.List(lstDir.ListIndex), 3) <> "txt" Then
    MsgBox "File is not a text file!", vbOKOnly, "Note:"
    Exit Sub
End If
'   ok, got here so set up to read directory
'   Creat filename of directory
If txtDirectory = "/" Then  '   This is the root
    sFileName = txtServer & txtDirectory & lstDir.List(lstDir.ListIndex)
ElseIf Right$(txtDirectory, 1) = "/" Then   '   / is there
    sFileName = txtServer & "/" & _
      txtDirectory & lstDir.List(lstDir.ListIndex)
Else    '   slash is not there, add one
    sFileName = txtServer & "/" & _
      txtDirectory & "/" & lstDir.List(lstDir.ListIndex)
End If
iAction = intntGetText
'   Get file
'   Display activity
lblWhat = "Reading " & lstDir.List(lstDir.ListIndex)
DoEvents
sCommand = sFileName
'   Display command being sent
txtCommand = txtCommand & vbCrLf & sCommand
txtCommand.SelStart = Len(txtCommand.Text)
Call subBusy(True)
frmText.txtData.Text = Inet1.OpenURL(sCommand)
'   Wait until done before doing anything else
Do While Inet1.StillExecuting
    DoEvents
Loop
'   done reading, display the form
Call subBusy(False)
frmText.Show
lblWhat.Caption = ""
Exit_Normal:
    Exit Sub
Error_Handler:
    MsgBox Err & " " & Error, vbOKOnly, "Read Text"
    Select Case Err
        Case 35761: Inet1.Cancel    ' Timeout
        Case 35764: Inet1.Cancel    ' Still Executing
    End Select
    Resume Exit_Normal
End Sub

Private Sub Form_Load()
Load frmText
lblWhat.Caption = "Enter Server Name"
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim i As Integer
For i = Forms.Count - 1 To 0
    Unload Forms(i)
Next
End
End Sub


Private Sub Inet1_StateChanged(ByVal State As Integer)
'   Fired each time inet control detects a change in state
Dim sMsg As String
On Error GoTo Error_Handler
'   Figure out what to do
Select Case State
    Case icResponseCompleted
    '   All of the preliminaries are done, now get the information
    lblWhat.Caption = "Response Completed "
    '   Display response
    If Inet1.ResponseCode <> 0 Then _
        txtCommand = txtCommand & vbCrLf & Inet1.ResponseCode & _
          ": " & Inet1.ResponseInfo
    txtCommand.SelStart = Len(txtCommand.Text)
    DoEvents
    Select Case iAction
        Case 0, 1, 4: subGetDir '   Make connection, CDUP, or change directory
        Case 2: '   Read text file -- handled in cmdReadText_Click
        Case 3: '   Download a file -- handled in cmdDownLoad_Click
    End Select
    Case icConnecting
        lblWhat.Caption = "Connecting"
        DoEvents
    Case icConnected
        lblWhat.Caption = "Connected"
        DoEvents
    Case icDisconnected
        lblWhat.Caption = "Disconnected"
        DoEvents
    Case icDisconnecting
        lblWhat.Caption = "Disconnecting"
        DoEvents
    Case icHostResolved
        lblWhat.Caption = "Host Resolved"
        DoEvents
    Case icReceivingResponse
        lblWhat.Caption = "Receiving Response "
        DoEvents
    Case icRequesting
        lblWhat.Caption = "Sending Request"
        DoEvents
    Case icRequestSent
        lblWhat.Caption = "Request Sent"
        DoEvents
    Case icResolvingHost
        lblWhat.Caption = "Resolving Host"
        DoEvents
    Case icError
        Select Case Inet1.ResponseCode
            Case 80, 87: Exit Sub
            Case 12007: Exit Sub
        End Select
        sMsg = "Error Code " & Inet1.ResponseCode
        sMsg = sMsg & vbCrLf & Inet1.ResponseInfo
        MsgBox sMsg, vbOKOnly, "Error Response Received"
        DoEvents
        Exit Sub
    Case icResponseReceived
        lblWhat.Caption = "Response Received!"
        '   Display response
        If Inet1.ResponseCode <> 0 Then _
            txtCommand = txtCommand & vbCrLf & Inet1.ResponseCode & _
             ": " & Inet1.ResponseInfo
        txtCommand.SelStart = Len(txtCommand.Text)
        DoEvents
    Case Else   '   Should never get here
        lblWhat.Caption = "Unknown State Received"
        DoEvents
End Select
Ok_Exit:
Exit Sub
Error_Handler:
MsgBox "Error # " & Err & " " & Error, vbOKOnly, "State Changed"
GoTo Ok_Exit
End Sub

Private Sub lstDir_DblClick()
    If Right$(lstDir.List(lstDir.ListIndex), 1) = "/" Then cmdDirectory_Click
End Sub

Private Sub txtServer_Change()
    cmdConnect.Enabled = True
End Sub

Private Sub txtServer_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    sServer = txtServer.Text
    cmdConnect_Click
End If
End Sub

Private Sub txtServer_LostFocus()
'   Records change in server name
sServer = txtServer.Text
End Sub

Public Sub subBusy(bz As Boolean)
If bz Then
    cmdDirectory.Enabled = False
    cmdDownLoad.Enabled = False
    cmdReadText.Enabled = False
    cmdClose.Enabled = True
    cmdConnect.Enabled = False
Else
    cmdDirectory.Enabled = True
    cmdDownLoad.Enabled = True
    cmdReadText.Enabled = True
    cmdClose.Enabled = True
    cmdConnect.Enabled = False
End If
End Sub
