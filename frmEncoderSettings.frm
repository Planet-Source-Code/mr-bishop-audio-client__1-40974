VERSION 5.00
Begin VB.Form frmEncoderSettings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Encoder settings"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5880
   Icon            =   "frmEncoderSettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   5880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Caption         =   "Misc"
      Height          =   975
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   5655
      Begin VB.CheckBox chkHighQuality 
         Caption         =   "High quality (slower)"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   600
         Width           =   2055
      End
      Begin VB.CheckBox chkDelFile 
         Caption         =   "Delete source file after encoding."
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   240
         Width           =   2655
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Mode"
      Height          =   1455
      Left            =   3600
      TabIndex        =   4
      Top             =   120
      Width           =   2175
      Begin VB.ListBox lstMode 
         Height          =   1035
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Bitrate"
      Height          =   1455
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3375
      Begin VB.ListBox lstBitrate 
         Height          =   1035
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   3135
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4920
      TabIndex        =   1
      Top             =   2760
      Width           =   855
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   4080
      TabIndex        =   0
      Top             =   2760
      Width           =   855
   End
End
Attribute VB_Name = "frmEncoderSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    
    Unload Me
    
End Sub

Private Sub cmdOK_Click()
    
    writeini "Lame", "Bitrate", lstBitrate.ListIndex, App.Path & "\Config.ini"
    writeini "Lame", "Mode", lstMode.ListIndex, App.Path & "\Config.ini"
    writeini "Lame", "KillFile", chkDelFile.Value, App.Path & "\Config.ini"
    writeini "Lame", "HighQuality", chkHighQuality.Value, App.Path & "\Config.ini"
    
    Unload Me
    
End Sub

Private Sub Form_Load()
    
    With lstBitrate
    
        .AddItem "8 kbits"
        .AddItem "16 kbits"
        .AddItem "24 kbits"
        .AddItem "32 kbits"
        .AddItem "40 kbits"
        .AddItem "48 kbits"
        .AddItem "56 kbits"
        .AddItem "64 kbits"
        .AddItem "80 kbits"
        .AddItem "96 kbits"
        .AddItem "112 kbits"
        .AddItem "128 kbits"
        .AddItem "144 kbits"
        .AddItem "160 kbits"
        .AddItem "192 kbits"
        .AddItem "224 kbits"
        .AddItem "256 kbits"
        .AddItem "320 kbits"
        
        .Selected(ReadINI("Lame", "Bitrate", App.Path & "\Config.ini")) = True
    
    End With
    
    With lstMode
        
        .AddItem "Default"
        .AddItem "Stereo"
        .AddItem "Joint Stereo"
        .AddItem "Forced Joint Stereo"
        .AddItem "Mono"
        
        .Selected(ReadINI("Lame", "Mode", App.Path & "\Config.ini")) = True
        
    End With
       
    chkDelFile.Value = ReadINI("Lame", "KillFile", App.Path & "\Config.ini")
    chkHighQuality.Value = ReadINI("Lame", "HighQuality", App.Path & "\Config.ini")
    
End Sub
