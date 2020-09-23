VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About"
   ClientHeight    =   3345
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6570
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmAbout.frx":038A
   ScaleHeight     =   3345
   ScaleWidth      =   6570
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   5640
      TabIndex        =   5
      Top             =   2880
      Width           =   855
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   2520
      Picture         =   "frmAbout.frx":1E3B
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   1
      Top             =   2880
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   930
      Left            =   240
      Picture         =   "frmAbout.frx":21C5
      ScaleHeight     =   870
      ScaleWidth      =   675
      TabIndex        =   0
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "November 20, 2002"
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   1080
      TabIndex        =   6
      Top             =   1920
      Width           =   1410
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "AKRip CD Extraction"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   165
      Left            =   1080
      TabIndex        =   4
      Top             =   2445
      Width           =   1305
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "L.A.M.E MP3 Encoder Technology"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   165
      Left            =   1080
      TabIndex        =   3
      Top             =   2235
      Width           =   2160
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Developed by: Tom Bishop"
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   1080
      TabIndex        =   2
      Top             =   1680
      Width           =   1920
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
    
    Unload Me
    
End Sub

Private Sub Form_Load()
    
    Label4.Caption = "AKRip CD Extraction v." & AK.GetAKRipDLLVersion
    
End Sub
