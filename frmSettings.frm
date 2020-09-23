VERSION 5.00
Begin VB.Form frmSettings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Program settings"
   ClientHeight    =   3570
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4215
   Icon            =   "frmSettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   4215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Output folder:"
      Height          =   1215
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   3975
      Begin VB.CheckBox chkAutoEncode 
         Caption         =   "Automatically encode WAV to MP3"
         Height          =   255
         Left            =   720
         TabIndex        =   13
         Top             =   720
         Width           =   2895
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "&Browse"
         Height          =   255
         Left            =   3000
         TabIndex        =   11
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox txtOutput 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   720
         TabIndex        =   10
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Path:"
         Height          =   195
         Left            =   240
         TabIndex        =   12
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3240
      TabIndex        =   8
      Top             =   3120
      Width           =   855
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   375
      Left            =   2400
      TabIndex        =   7
      Top             =   3120
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Caption         =   "CDDB:"
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   3975
      Begin VB.TextBox txtServer 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   720
         TabIndex        =   3
         Text            =   "www.freedb.org"
         Top             =   360
         Width           =   3015
      End
      Begin VB.TextBox txtCGI 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   720
         TabIndex        =   2
         Text            =   "/~cddb/cddb.cgi"
         Top             =   720
         Width           =   3015
      End
      Begin VB.TextBox txtUser 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   720
         TabIndex        =   1
         Text            =   "user@akrip.sourceforge.net"
         Top             =   1080
         Width           =   3015
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Server:"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   510
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cgi:"
         Height          =   195
         Left            =   360
         TabIndex        =   5
         Top             =   720
         Width           =   270
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User:"
         Height          =   195
         Left            =   240
         TabIndex        =   4
         Top             =   1080
         Width           =   375
      End
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBrowse_Click()

    Dim lpIDList As Long
    Dim sBuffer As String
    Dim szTitle As String
    Dim tBrowseInfo As BrowseInfo
    
    With tBrowseInfo
       .hWndOwner = Me.hWnd
       .lpszTitle = lstrcat(szTitle, "")
       .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
    End With
    
    lpIDList = SHBrowseForFolder(tBrowseInfo)
    
    If (lpIDList) Then
       sBuffer = Space(MAX_PATH)
       SHGetPathFromIDList lpIDList, sBuffer
       sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
       
        If Right(sBuffer, 1) = "\" Then
            txtOutput.Text = Mid(sBuffer, 1, Len(sBuffer) - 1)
        Else
            txtOutput.Text = sBuffer
        End If
    End If

End Sub

Private Sub cmdCancel_Click()
    
    Unload Me
    
End Sub

Private Sub cmdSave_Click()
    
    writeini "Output", "AutoEncode", chkAutoEncode.Value, App.Path & "\config.ini"
    writeini "CDDB", "Server", txtServer.Text, App.Path & "\config.ini"
    writeini "CDDB", "CGI", txtCGI.Text, App.Path & "\config.ini"
    writeini "CDDB", "User", txtUser.Text, App.Path & "\config.ini"
    writeini "Output", "Path", txtOutput.Text, App.Path & "\config.ini"
    strPath = txtOutput.Text
    
    Unload Me
    
End Sub

Private Sub Form_Load()
    
    EdgeSubClass txtServer.hWnd, sedSunkenOuter
    EdgeSubClass txtCGI.hWnd, sedSunkenOuter
    EdgeSubClass txtUser.hWnd, sedSunkenOuter
    EdgeSubClass txtOutput.hWnd, sedSunkenOuter
    
    chkAutoEncode.Value = ReadINI("Output", "AutoEncode", App.Path & "\config.ini")
    txtServer.Text = ReadINI("CDDB", "Server", App.Path & "\config.ini")
    txtCGI.Text = ReadINI("CDDB", "CGI", App.Path & "\config.ini")
    txtUser.Text = ReadINI("CDDB", "User", App.Path & "\config.ini")
    txtOutput.Text = strPath
    
End Sub
