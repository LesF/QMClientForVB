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
      Height          =   2415
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Text            =   "FormMain.frx":0000
      Top             =   1920
      Width           =   7935
   End
   Begin VB.Frame FrameAPI 
      Caption         =   "API Service"
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7935
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "TODO : connection properties"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   360
         TabIndex        =   1
         Top             =   720
         Width           =   4200
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

Private Sub Form_Resize()
    '-----
    Dim pad As Integer, logH As Integer, logW As Integer
    Dim clientW As Integer, clientH As Integer
    '-----
    If FormMain.WindowState = vbMinimized Then Exit Sub
    clientW = FormMain.ScaleWidth
    clientH = FormMain.ScaleHeight
    pad = FrameAPI.Left
    If clientW < (pad * 3) Then Exit Sub
    If clientH < (pad * 10) Then Exit Sub
    '-----
    logW = clientW - (pad * 2)
    If logW > pad Then
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
