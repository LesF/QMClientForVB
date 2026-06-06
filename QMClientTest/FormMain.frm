VERSION 5.00
Begin VB.Form FormMain 
   Caption         =   "So you want to know if QMClient works huh?"
   ClientHeight    =   4500
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8160
   LinkTopic       =   "Form1"
   ScaleHeight     =   4500
   ScaleWidth      =   8160
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TextLog 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   12
      Text            =   "FormMain.frx":0000
      Top             =   2160
      Width           =   7935
   End
   Begin VB.Frame FrameAPI 
      Caption         =   "API Service"
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7935
      Begin VB.CommandButton CmdConnect 
         Caption         =   "Connect"
         Height          =   375
         Left            =   3840
         TabIndex        =   11
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox TextAccount 
         Height          =   285
         Left            =   1080
         TabIndex        =   10
         Top             =   1440
         Width           =   2415
      End
      Begin VB.TextBox TextPass 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1080
         TabIndex        =   8
         Top             =   1080
         Width           =   2415
      End
      Begin VB.TextBox TextUser 
         Height          =   285
         Left            =   1080
         TabIndex        =   6
         Top             =   720
         Width           =   2415
      End
      Begin VB.TextBox TextPort 
         Height          =   285
         Left            =   4200
         TabIndex        =   4
         Text            =   "4243"
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox TextHost 
         Height          =   285
         Left            =   1080
         TabIndex        =   2
         Text            =   "192.168.1.174"
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "As shown in LIST QM.ACCOUNTS"
         Height          =   195
         Left            =   3720
         TabIndex        =   13
         Top             =   1440
         Width           =   2475
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Account"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   1440
         Width           =   600
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Password"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   1080
         Width           =   690
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "User"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   330
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Port"
         Height          =   195
         Left            =   3720
         TabIndex        =   3
         Top             =   360
         Width           =   285
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Host"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   330
      End
   End
End
Attribute VB_Name = "FormMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    LogThis "Enter API connection properties then click... umm... TODO, Create a Start button.", True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    QMDisconnect
End Sub

Private Sub CmdConnect_Click()
    '-----
    Dim rc As Integer
    Dim hostname As String, Port As String, portNum As Integer
    Dim UserName As String, pwd As String, Account As String
    '-----
    hostname = Trim(TextHost.Text)
    Port = Trim(TextPort.Text)
    UserName = Trim(TextUser.Text)
    pwd = Trim(TextPass.Text)
    Account = Trim(TextAccount.Text)
     
    portNum = 0
    If IsNumeric(Port) Then portNum = CInt(Port)
    If Port = 0 Then
        LogThis "* Invalid port number"
        Exit Sub
    End If

    ' Call QMConnect passing BSTR pointers
    LogThis vbCrLf & "Connecting...", True
    rc = QMConnect(hostname, portNum, UserName, pwd, Account)
    If rc <> 0 Then
        LogThis "QMConnect succeeded", True
        If QMConnected() <> 0 Then
            LogThis " — session OK"
        Else
            LogThis " — QMConnected returned FALSE"
        End If
    Else
        LogThis "QMConnect failed", True
    End If
End Sub

Private Sub Form_Resize()
    '-----
    Dim pad As Integer, logH As Integer, logW As Integer, minW As Integer
    Dim clientW As Integer, clientH As Integer
    '-----
    If FormMain.WindowState = vbMinimized Then Exit Sub
    clientW = FormMain.ScaleWidth
    clientH = FormMain.ScaleHeight
    pad = FrameAPI.Left
    If clientW < (pad * 3) Then Exit Sub
    If clientH < (pad * 10) Then Exit Sub
    '-----
    minW = CmdConnect.Left + CmdConnect.Width + (pad * 2)
    logW = clientW - (pad * 2)
    If logW >= minW Then
        FrameAPI.Width = logW
        TextLog.Width = logW
    End If
    logH = clientH - (TextLog.Top + pad)
    If logH > pad Then TextLog.Height = logH
End Sub

Private Sub LogThis(msg As String, Optional withTimestamp As Boolean = False)
    Dim ts As String
    If withTimestamp Then ts = Format$(Time$, "HH:mm:ss") & ": "
    TextLog.SelStart = Len(TextLog.Text)
    TextLog.SelText = vbCrLf & ts & msg
    TextLog.SelStart = Len(TextLog.Text)
End Sub

