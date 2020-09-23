VERSION 5.00
Begin VB.Form frmCDDB 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CDDB"
   ClientHeight    =   3105
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4200
   Icon            =   "frmCDDB.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   4200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3240
      TabIndex        =   8
      Top             =   2640
      Width           =   855
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Download CD Info..."
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   2640
      Width           =   3135
   End
   Begin VB.Frame Frame1 
      Caption         =   "Connection settings:"
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3975
      Begin VB.TextBox txtUser 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   240
         TabIndex        =   6
         Text            =   "user@akrip.sourceforge.net"
         Top             =   1920
         Width           =   3495
      End
      Begin VB.TextBox txtCGI 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   240
         TabIndex        =   4
         Text            =   "/~cddb/cddb.cgi"
         Top             =   1320
         Width           =   3495
      End
      Begin VB.TextBox txtServer 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   240
         TabIndex        =   2
         Text            =   "www.freedb.org"
         Top             =   600
         Width           =   3495
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User:"
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   1680
         Width           =   375
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cgi:"
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   1080
         Width           =   270
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Server:"
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   510
      End
   End
End
Attribute VB_Name = "frmCDDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    
    Unload Me
    
End Sub

Private Sub cmdOK_Click()
    
    Screen.MousePointer = 11
    frmMain.LogIT "Getting CD Info..." & vbCrLf, &H0&, True
    DoEvents
    frmMain.CDDBQuery
    Screen.MousePointer = 1
    frmMain.LogIT "Done." & vbCrLf, &H0&, True
    Unload Me
    
End Sub

Private Sub Form_Load()
    
    EdgeSubClass txtServer.hWnd, sedSunkenOuter
    EdgeSubClass txtCGI.hWnd, sedSunkenOuter
    EdgeSubClass txtUser.hWnd, sedSunkenOuter
    
End Sub
