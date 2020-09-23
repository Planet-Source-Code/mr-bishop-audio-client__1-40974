VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{20B402FF-3C7C-4965-9923-85D189470E1D}#1.0#0"; "ProgressBar2.ocx"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "Comct332.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Audio Client"
   ClientHeight    =   5880
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8280
   ForeColor       =   &H00000000&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5880
   ScaleWidth      =   8280
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   5520
      ScaleHeight     =   255
      ScaleWidth      =   2775
      TabIndex        =   12
      Top             =   0
      Width           =   2775
      Begin AudioClient.Hyperlink hAbout 
         Height          =   255
         Left            =   1680
         TabIndex        =   13
         Top             =   30
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   450
         ForeColorIdle   =   16711680
         BackColor       =   -2147483648
         Caption         =   "About"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin AudioClient.Hyperlink Hyperlink1 
         Height          =   255
         Left            =   2280
         TabIndex        =   14
         Top             =   30
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         ForeColorIdle   =   16711680
         Caption         =   "Exit"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin ProgressBar2.cpvProgressBar PB1 
      Height          =   1590
      Left            =   180
      Top             =   4320
      Width           =   180
      _ExtentX        =   318
      _ExtentY        =   2805
      BarPicture      =   "frmMain.frx":038A
      BarPictureBack  =   "frmMain.frx":12C6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Max             =   34
      Orientation     =   1
      Value           =   17
   End
   Begin ProgressBar2.cpvProgressBar PB2 
      Height          =   1590
      Left            =   360
      Top             =   4320
      Width           =   180
      _ExtentX        =   318
      _ExtentY        =   2805
      BarPicture      =   "frmMain.frx":2202
      BarPictureBack  =   "frmMain.frx":313E
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Max             =   34
      Orientation     =   1
      Value           =   17
   End
   Begin ProgressBar2.cpvProgressBar PB3 
      Height          =   1590
      Left            =   0
      Top             =   4320
      Width           =   180
      _ExtentX        =   318
      _ExtentY        =   2805
      BarPicture      =   "frmMain.frx":407A
      BarPictureBack  =   "frmMain.frx":4FB6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Max             =   34
      Orientation     =   1
      Value           =   17
   End
   Begin RichTextLib.RichTextBox Status 
      Height          =   1575
      Left            =   540
      TabIndex        =   0
      Top             =   4320
      Width           =   7755
      _ExtentX        =   13679
      _ExtentY        =   2778
      _Version        =   393217
      BackColor       =   16777215
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmMain.frx":5EF2
   End
   Begin MSComctlLib.ImageList imgMain 
      Left            =   6960
      Top             =   3240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   16
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5F74
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":630E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":66A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6A42
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6DDC
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7176
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7510
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":78AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7C44
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7FDE
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8378
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8712
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8AAC
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8E46
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":91E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":957A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picPlayer 
      BackColor       =   &H00400000&
      BorderStyle     =   0  'None
      Height          =   3935
      Left            =   0
      Picture         =   "frmMain.frx":9914
      ScaleHeight     =   3930
      ScaleWidth      =   8295
      TabIndex        =   15
      Top             =   360
      Visible         =   0   'False
      Width           =   8295
      Begin MSComctlLib.ListView lvPlaylist 
         Height          =   1455
         Left            =   450
         TabIndex        =   26
         Top             =   2020
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   2566
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         SmallIcons      =   "imgMain"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Playlist:"
            Object.Width           =   12418
         EndProperty
      End
      Begin VB.Timer tmrPosition 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   6360
         Top             =   2880
      End
      Begin ProgressBar2.cpvProgressBar pbVolume 
         Height          =   975
         Left            =   7600
         Top             =   400
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   1720
         BarPicture      =   "frmMain.frx":ACCC
         BarPictureBack  =   "frmMain.frx":B02F
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Orientation     =   1
      End
      Begin ProgressBar2.cpvProgressBar PBProgress 
         Height          =   165
         Left            =   360
         Top             =   1440
         Width           =   7500
         _ExtentX        =   13229
         _ExtentY        =   291
         BarPicture      =   "frmMain.frx":B26B
         BarPictureBack  =   "frmMain.frx":B3B6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin AudioClient.Hyperlink Hyperlink2 
         Height          =   200
         Left            =   360
         TabIndex        =   20
         Top             =   1680
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   344
         ForeColorIdle   =   16711680
         ForeColorMouse  =   16777215
         BackColor       =   4194304
         Caption         =   "• Open"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin AudioClient.Hyperlink Hyperlink3 
         Height          =   200
         Left            =   1680
         TabIndex        =   21
         Top             =   1680
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   344
         ForeColorIdle   =   16711680
         ForeColorMouse  =   16777215
         BackColor       =   4194304
         Caption         =   "• Stop"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin AudioClient.Hyperlink Hyperlink4 
         Height          =   200
         Left            =   960
         TabIndex        =   22
         Top             =   1680
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   344
         ForeColorIdle   =   16711680
         ForeColorMouse  =   16777215
         BackColor       =   4194304
         Caption         =   "• Pause"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "100%"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   150
         Left            =   7290
         TabIndex        =   25
         Top             =   1260
         Width           =   270
      End
      Begin VB.Label lblElapsedTime 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "00:00:00"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   195
         Left            =   6360
         TabIndex        =   24
         Top             =   795
         Width           =   1095
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Progress: (click to advance)"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   165
         Left            =   360
         TabIndex        =   23
         Top             =   1200
         Width           =   1740
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Now playing:"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Left            =   480
         TabIndex        =   19
         Top             =   445
         Width           =   765
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Elapsed time:"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   165
         Left            =   5400
         TabIndex        =   18
         Top             =   780
         Width           =   825
      End
      Begin VB.Label lblPosition 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0\0"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFC0&
         Height          =   195
         Left            =   5880
         TabIndex        =   17
         Top             =   1680
         Width           =   1935
      End
      Begin VB.Label lblArtist 
         BackStyle       =   0  'Transparent
         Caption         =   "Click 'Open' to select song..."
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   195
         Left            =   1320
         TabIndex        =   16
         Top             =   445
         Width           =   6135
      End
   End
   Begin VB.PictureBox picExtract 
      BorderStyle     =   0  'None
      Height          =   3975
      Left            =   0
      ScaleHeight     =   3975
      ScaleWidth      =   8295
      TabIndex        =   7
      Top             =   360
      Width           =   8295
      Begin MSComctlLib.ListView lvMain 
         Height          =   3210
         Left            =   0
         TabIndex        =   8
         Top             =   840
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   5662
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         SmallIcons      =   "imgMain"
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Track #"
            Object.Width           =   1482
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Track name"
            Object.Width           =   12594
         EndProperty
         Picture         =   "frmMain.frx":B4FB
      End
      Begin ComCtl3.CoolBar CoolBar1 
         Height          =   735
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   1296
         BandCount       =   2
         _CBWidth        =   8295
         _CBHeight       =   735
         _Version        =   "6.7.8988"
         Child1          =   "tbMain"
         MinHeight1      =   330
         Width1          =   2055
         NewRow1         =   0   'False
         Caption2        =   "CD-Rom Drives:"
         Child2          =   "cmbDrives"
         MinHeight2      =   315
         Width2          =   270
         NewRow2         =   -1  'True
         Begin VB.ComboBox cmbDrives 
            Height          =   315
            Left            =   1365
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   390
            Width           =   6840
         End
         Begin MSComctlLib.Toolbar tbMain 
            Height          =   330
            Left            =   165
            TabIndex        =   10
            Top             =   30
            Width           =   8040
            _ExtentX        =   14182
            _ExtentY        =   582
            ButtonWidth     =   1746
            ButtonHeight    =   582
            Style           =   1
            TextAlignment   =   1
            ImageList       =   "imgMain"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   9
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "Extract"
                  Key             =   "Extract"
                  Object.ToolTipText     =   "Extract audio from cd..."
                  ImageIndex      =   2
                  Style           =   5
                  BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                     NumButtonMenus  =   2
                     BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                        Key             =   "selected"
                        Text            =   "Extract selected tracks"
                     EndProperty
                     BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                        Key             =   "all"
                        Text            =   "Extract all tracks"
                     EndProperty
                  EndProperty
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Caption         =   "Stop!"
                  Key             =   "stop"
                  Object.ToolTipText     =   "Stop extracting..."
                  ImageIndex      =   10
               EndProperty
               BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "CDDB"
                  Key             =   "CDDB"
                  Object.ToolTipText     =   "Get CDDB info..."
                  ImageIndex      =   4
                  Style           =   5
                  BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                     NumButtonMenus  =   3
                     BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                        Key             =   "info"
                        Text            =   "Get CD Info..."
                     EndProperty
                     BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                        Text            =   "-"
                     EndProperty
                     BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                        Key             =   "servers"
                        Text            =   "Get CDDB Servers"
                     EndProperty
                  EndProperty
               EndProperty
               BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Style           =   3
               EndProperty
               BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "Settings"
                  Key             =   "pSettings"
                  Object.ToolTipText     =   "Program settings..."
                  ImageIndex      =   8
               EndProperty
               BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.Visible         =   0   'False
                  Key             =   "eSettings"
                  Object.ToolTipText     =   "Encoder settings..."
                  ImageIndex      =   7
               EndProperty
               BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Style           =   3
               EndProperty
               BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "Wizard"
                  Key             =   "wizard"
                  Object.ToolTipText     =   "Wizard..."
                  ImageIndex      =   11
               EndProperty
               BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.Visible         =   0   'False
                  Caption         =   "About"
                  Key             =   "About"
                  Object.ToolTipText     =   "About"
                  ImageIndex      =   9
               EndProperty
            EndProperty
         End
      End
   End
   Begin VB.PictureBox picEncode 
      BorderStyle     =   0  'None
      Height          =   3975
      Left            =   0
      ScaleHeight     =   3975
      ScaleWidth      =   8295
      TabIndex        =   2
      Top             =   360
      Visible         =   0   'False
      Width           =   8295
      Begin MSComDlg.CommonDialog CD1 
         Left            =   6480
         Top             =   2160
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.FileListBox File1 
         Height          =   285
         Left            =   6000
         TabIndex        =   4
         Top             =   2760
         Visible         =   0   'False
         Width           =   2295
      End
      Begin MSComctlLib.ListView lvEncode 
         Height          =   3555
         Left            =   0
         TabIndex        =   3
         Top             =   400
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   6271
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         SmallIcons      =   "imgMain"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Input file"
            Object.Width           =   6881
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Output file"
            Object.Width           =   7057
         EndProperty
         Picture         =   "frmMain.frx":BF8F
      End
      Begin ComCtl3.CoolBar CoolBar2 
         Height          =   390
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   688
         BandCount       =   1
         _CBWidth        =   8295
         _CBHeight       =   390
         _Version        =   "6.7.8988"
         Child1          =   "tbEncode"
         MinHeight1      =   330
         Width1          =   4335
         NewRow1         =   0   'False
         Begin MSComctlLib.Toolbar tbEncode 
            Height          =   330
            Left            =   30
            TabIndex        =   6
            Top             =   30
            Width           =   8175
            _ExtentX        =   14420
            _ExtentY        =   582
            ButtonWidth     =   1746
            ButtonHeight    =   582
            Style           =   1
            TextAlignment   =   1
            ImageList       =   "imgMain"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   5
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "Encode"
                  ImageIndex      =   3
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Style           =   3
               EndProperty
               BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "Settings"
                  ImageIndex      =   7
               EndProperty
               BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Style           =   3
               EndProperty
               BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "Add"
                  ImageIndex      =   15
                  Style           =   5
                  BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                     NumButtonMenus  =   2
                     BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                        Key             =   "file"
                        Text            =   "File..."
                     EndProperty
                     BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                        Key             =   "folder"
                        Text            =   "Folder..."
                     EndProperty
                  EndProperty
               EndProperty
            EndProperty
         End
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   4455
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   8325
      _ExtentX        =   14684
      _ExtentY        =   7858
      HotTracking     =   -1  'True
      Separators      =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Extractor"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Encoder"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Media Player"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuExtract 
      Caption         =   "Extract"
      Visible         =   0   'False
      Begin VB.Menu mnuExtract2 
         Caption         =   "Extract"
         Begin VB.Menu mnuExtractSelected 
            Caption         =   "Selected tracks"
         End
         Begin VB.Menu mnuExtractAll 
            Caption         =   "All tracks"
         End
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private isPlaying As Boolean
Private sReturnBuffer As String * 30
Private lngLength As Long
Private strStartTime As String

Private Sub cmbDrives_Click()
    
    'Make Sure CD Handle is Closed.
    If Val(intHandle) <> 0 Then
        If AK.CloseCDHandle(Val(intHandle)) = True Then
            intHandle = 0
        Else
            MsgBox "Failure to close CD Handle", vbCritical
            Exit Sub
        End If
    End If

    'Open CD Handle
    If Val(intHandle) = 0 Then
        intHandle = AK.GetCDHandle(1, Val(intCDAdd0), Val(cmbDrives.ListIndex), Val(intCDAdd2), CDR_ANY, False, Val(intOverlap) - 2, Val(intOverlap))
        If Val(intHandle) = 0 Then
            MsgBox "Failed to open a CD Handle at that address", vbCritical
        Else
            ReadTracks
        End If
    Else
        MsgBox "Already have a CD Handle", vbCritical
    End If

End Sub

Private Sub Form_Load()
           
    Dim OSInfo As OSVERSIONINFO, retval As Long
    
    LogIT "Loading..." & vbCrLf, &H80&, False
    
    PB1.Value = 0
    PB2.Value = 0
    PB3.Value = 0
       
    Set AK = New AKRipAX.AKRip
    
    intHandle = 0
    intOverlap = 3
    intCDAdd0 = 1
    intCDAdd1 = 0
    intCDAdd2 = 0
    intCDParam0 = 29
    intCDParam1 = 0
    intStartFrame = 0
    intFrameLen = 2500
    
    GetCDList
    cmbDrives.ListIndex = 0
    
    strPath = ReadINI("Output", "Path", App.Path & "\config.ini")
    If Len(strPath) = 0 Then
        strPath = App.Path
    End If

    OSInfo.dwOSVersionInfoSize = Len(OSInfo)
    
    retval = GetVersionEx(OSInfo)
    If retval <> 0 Then
       Select Case OSInfo.dwPlatformId
           Case 0
               PId = "Win32"
           Case 1
               PId = "Win9x"
           Case 2
               PId = "WinNT"
       End Select
    End If
    
    pbVolume.Min = 0
    pbVolume.Max = 1000
    pbVolume.Value = 500
    
    isPlaying = False
    doWizard = False
    LogIT "  - Client OS: " & PId, &H0&, False
    LogIT "Audio Client v1 | Loaded!" & vbCrLf, &H80&, False
    
              
End Sub

Private Sub GetCDList()

    Dim lngHA As Long
    Dim lngTGT As Long
    Dim lngLUN As Long
    Dim lngPAD As Long
    Dim strID As String
    
    Dim strVendor As String
    Dim strProductID As String
    Dim strRevision As String
    Dim strVendorSpec As String
    Dim lngLoop As Long
    Dim I As Long
     
    cmbDrives.Clear
    AK.GetCDList
    
    If AK.GetCDListCount = 0 Then
    
        cmbDrives.AddItem "No CD-ROM Drives found!"
        
    Else
    
        I = 0
        For lngLoop = 0 To AK.GetCDListCount - 1
          Call AK.GetCDListRecord(lngLoop, lngHA, lngTGT, lngLUN, lngPAD, strID)
          Call AK.GetCDListInfo(lngLoop, strVendor, strProductID, strRevision, strVendorSpec)
          cmbDrives.AddItem Trim(strVendor) & " " & Trim(strProductID)
          I = I + 1
        Next lngLoop
        LogIT "  - Detected: " & I & " drive(s).", &H0&, False
        
    End If

End Sub

Public Sub ReadTracks()

    Dim lngADR As Long
    Dim lngTrackNumber As Long
    Dim lngMins As Long
    Dim lngSecs As Long
    Dim lngFrames As Long
    Dim lngLoop As Long
    Dim listX As ListItem
    
    lvMain.ListItems.Clear
    If Val(intHandle) <> 0 Then
    
        If AK.ModifyCDParms(Val(intHandle), CDP_MSF, True) = False Then
            MsgBox "Error reading tracks!" & vbCrLf & "Make sure that there is a valid CD is the drive.", vbCritical
        End If
        
        If AK.ReadTOC(Val(intHandle)) = SS_COMP Then
        
            For lngLoop = AK.ReadTOCFirstTrack To AK.ReadTOCKLastTrack
                Call AK.ReadTOCTrack(lngLoop - 1, lngADR, lngTrackNumber, lngMins, lngSecs, lngFrames)
                Set listX = lvMain.ListItems.Add(, , , , 6)
                    listX.Text = lngTrackNumber
                    listX.SubItems(1) = "Track " & lngTrackNumber
            Next lngLoop
            Call AK.ReadTOCTrack(AK.ReadTOCKLastTrack, lngADR, lngTrackNumber, lngMins, lngSecs, lngFrames)
        
        Else
        
            MsgBox "Error reading tracks!" & vbCrLf & "Make sure that there is a valid CD is the drive.", vbCritical
        
        End If
        
    Else
        
        MsgBox "Error reading tracks!" & vbCrLf & "No CD Handle.", vbCritical
        
    End If

End Sub

Public Sub CDDBServerList()

    Dim lngNum As Long
    Dim strServer As String
    Dim blnHTTP As Boolean
    Dim lngPort As Long
    Dim strCGI As String
    Dim strNorth As String
    Dim strSouth As String
    Dim strLocation As String
    Dim lngLoop As Long
      
    AK.CDDBSetOption CDDB_OPT_SERVER, ReadINI("CDDB", "Server", App.Path & "\Config.ini"), 0
    AK.CDDBSetOption CDDB_OPT_CGI, ReadINI("CDDB", "CGI", App.Path & "\Config.ini"), 0
    AK.CDDBSetOption CDDB_OPT_USER, ReadINI("CDDB", "User", App.Path & "\Config.ini"), 0
    AK.CDDBSetOption CDDB_OPT_AGENT, "program 1.0", 0
    'Enable lines below and setup with proxy setting if you have a proxy
    'AK.CDDBSetOption CDDB_OPT_USEPROXY, "", True
    'AK.CDDBSetOption CDDB_OPT_PROXYPORT, "", 1080
    'AK.CDDBSetOption CDDB_OPT_PROXY, "192.168.200.1", 0
    
    If AK.CDDBGetServerList(lngNum) = SS_COMP Then
    
        For lngLoop = 0 To lngNum - 1
            AK.CDDBGetServerListRead lngLoop, strServer, blnHTTP, lngPort, strCGI, strNorth, strSouth, strLocation
            LogIT "  Server = " & strServer, &H80&, True
            LogIT "  HTTP = " & blnHTTP, &H0&, False
            LogIT "  Port = " & lngPort, &H0&, False
            LogIT "  CGI = " & strCGI, &H0&, False
            LogIT "  North = " & strNorth, &H0&, False
            LogIT "  South = " & strSouth, &H0&, False
            LogIT "  Location = " & strLocation & vbCrLf, &H0&, False
        Next lngLoop
        
    Else
    
        MsgBox "Error getting servers!", vbCritical
    
    End If

End Sub

Public Sub LogIT(what As String, daColor As String, isBold As Boolean)
    
    On Error Resume Next
    Status.SelStart = Len(Status.Text) - 1
    Status.SelColor = daColor
    Status.SelBold = isBold
    Status.SelText = " " & what & vbCrLf
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    If isPlaying Then
        mciSendString "stop mp3", 0, 0, 0
        mciSendString "close all", 0, 0, 0
    End If
    
End Sub

Private Sub hAbout_Click()
    
    frmAbout.Show 1
    
End Sub

Private Sub Hyperlink1_Click()
    
    Unload Me
    
End Sub

Private Sub Hyperlink2_Click()

    Dim strFile As String
    Dim listX As ListItem
    Dim varTemp As Variant
    Dim intTemp As Long
    Dim ObjMP3 As New clsMP3
    
    CD1.Filter = "MP3 Files (*.mp3)|*.mp3|"
    CD1.ShowOpen
    
    If Len(CD1.FileName) = 0 Then Exit Sub
    strFile = GetShortName(CD1.FileName)
    
    varTemp = Split(CD1.FileName, "\")
    Set listX = lvPlaylist.ListItems.Add(, , , , 3)
        listX.Text = CD1.FileName
        listX.Selected = True
    
    If isPlaying Then
        PBProgress.Value = 0
        lblPosition.Caption = "0"
        mciSendString "stop mp3", 0, 0, 0
        mciSendString "close all", 0, 0, 0
    End If
    
    '1.Open File, 2.Set time format, 3.Play File. 4.Retrieve length of File.
    mciSendString "open " & strFile & " type MPEGVideo alias mp3", 0, 0, 0
    mciSendString "set mp3 time format tmsf", 0, 0, 0
    mciSendString "play mp3", 0, 0, 0
    mciSendString "status mp3 length", sReturnBuffer$, Len(sReturnBuffer$), 0
    mciSendString "setaudio mp3 volume to " & pbVolume.Value, 0, 0, 0
    
    isPlaying = True
    lblPosition.Caption = sReturnBuffer$ & "\0"
    lngLength = sReturnBuffer$
    
    ObjMP3.ReadMP3 CD1.FileName
    If ObjMP3.Artist = "" Then
        intTemp = UBound(varTemp)
        lblArtist.Caption = varTemp(intTemp)
    Else
        lblArtist.Caption = ObjMP3.Artist & " - " & ObjMP3.Songname
    End If
    
    strStartTime = Time
    PBProgress.Value = 0
    PBProgress.Max = sReturnBuffer$
    tmrPosition.Enabled = True

End Sub

Private Sub Hyperlink3_Click()
    
    PBProgress.Value = 0
    lblPosition.Caption = "0\0"
    tmrPosition.Enabled = False
    isPlaying = False
    mciSendString "stop mp3", 0, 0, 0
    mciSendString "close all", 0, 0, 0

End Sub

Private Sub Hyperlink4_Click()
    
    If isPlaying Then
        isPlaying = False
        mciSendString "pause mp3", 0, 0, 0
        tmrPosition.Enabled = False
    Else
    
        isPlaying = True
        mciSendString "play mp3", 0, 0, 0
        mciSendString "status mp3 length", sReturnBuffer$, Len(sReturnBuffer$), 0
        mciSendString "setaudio mp3 volume to " & pbVolume.Value, 0, 0, 0
        tmrPosition.Enabled = True
        
    End If
    
End Sub

Private Sub lvPlaylist_DblClick()

    Dim strFile As String
    Dim varTemp As Variant
    Dim intTemp As Long
    Dim ObjMP3 As New clsMP3
       
    If Len(lvPlaylist.SelectedItem.Text) = 0 Then Exit Sub
    strFile = GetShortName(lvPlaylist.SelectedItem.Text)
    
    varTemp = Split(lvPlaylist.SelectedItem.Text, "\")
    
    If isPlaying Then
        PBProgress.Value = 0
        lblPosition.Caption = "0"
        mciSendString "stop mp3", 0, 0, 0
        mciSendString "close all", 0, 0, 0
    End If
    
    '1.Open File, 2.Set time format, 3.Play File. 4.Retrieve length of File.
    mciSendString "open " & strFile & " type MPEGVideo alias mp3", 0, 0, 0
    mciSendString "set mp3 time format tmsf", 0, 0, 0
    mciSendString "play mp3", 0, 0, 0
    mciSendString "status mp3 length", sReturnBuffer$, Len(sReturnBuffer$), 0
    mciSendString "setaudio mp3 volume to " & pbVolume.Value, 0, 0, 0
    
    isPlaying = True
    lblPosition.Caption = sReturnBuffer$ & "\0"
    lngLength = sReturnBuffer$
    
    ObjMP3.ReadMP3 lvPlaylist.SelectedItem.Text
    If ObjMP3.Artist = "" Then
        intTemp = UBound(varTemp)
        lblArtist.Caption = varTemp(intTemp)
    Else
        lblArtist.Caption = ObjMP3.Artist & " - " & ObjMP3.Songname
    End If
    
    strStartTime = Time
    PBProgress.Value = 0
    PBProgress.Max = sReturnBuffer$
    tmrPosition.Enabled = True



End Sub

Public Sub mnuExtractSelected_Click()
    
    Dim I, X As Long
    Dim intSelected As Long
    
    If lvMain.ListItems.Count = 0 Then
        MsgBox "No tracks to extract!", vbCritical
        Exit Sub
    End If
    
    intSelected = 0
    For I = 1 To lvMain.ListItems.Count
        If lvMain.ListItems(I).Checked = True Then
            intSelected = intSelected + 1
        End If
    Next I
    
    If intSelected = 0 Then
        MsgBox "Please select a track.", vbCritical
        Exit Sub
    End If
    
    X = 1
    PB3.Max = intSelected
    
    tbMain.Buttons(1).Enabled = False
    tbMain.Buttons(3).Enabled = False
    tbMain.Buttons(5).Enabled = False
    tbMain.Buttons(6).Enabled = False
    tbMain.Buttons(8).Enabled = False
    tbMain.Buttons(2).Enabled = True
    
    If doWizard Then
        Unload frmWizard
    End If
    
    LogIT "Starting extraction...", &H0&, True
    LogIT "Output folder: " & strPath & "\" & vbCrLf, &H0&, False
    
    For I = 1 To lvMain.ListItems.Count
        If lvMain.ListItems(I).Checked = True Then
            
            If blStop Then
                PB3.Value = PB3.Max
                LogIT "  Extraction stopped!" & vbCrLf, &H0&, False
                Exit For
            End If
            
            PB3.Value = X
            lvMain.ListItems(I).Selected = True
            lvMain.ListItems(I).EnsureVisible
            LogIT "  Extracting: " & lvMain.ListItems(I).ListSubItems(1).Text, &H0&, False
            
            GrabTrackLBA lvMain.ListItems(I).Text, Replace(lvMain.ListItems(I).ListSubItems(1).Text, " ", "_")
            
            If blStop Then
                PB3.Value = PB3.Max
                LogIT "  Extraction stopped!" & vbCrLf, &H0&, False
                Exit For
            End If
            
            If ReadINI("Output", "AutoEncode", App.Path & "\Config.ini") = 1 Then
                EncodeMP3 Replace(lvMain.ListItems(I).ListSubItems(1).Text, " ", "_") & ".wav", Replace(lvMain.ListItems(I).ListSubItems(1).Text, " ", "_") & ".mp3"
            End If
            
            LogIT "  Extraction complete." & vbCrLf, &H0&, False
            X = X + 1
            
        End If
    Next I
    
    LogIT "Done." & vbCrLf & vbCrLf, &H0&, True
    
    tbMain.Buttons(1).Enabled = True
    tbMain.Buttons(3).Enabled = True
    tbMain.Buttons(5).Enabled = True
    tbMain.Buttons(6).Enabled = True
    tbMain.Buttons(8).Enabled = True
    tbMain.Buttons(2).Enabled = False
    
End Sub

Private Sub PBProgress_Click()
        
    If isPlaying Then
        lblPosition.Caption = lngLength & "\" & PBProgress.Value
        mciSendString "play mp3 from " & PBProgress.Value, 0, 0, 0
    Else
        PBProgress.Value = 0
    End If

End Sub

Private Sub PBProgress_ValueChanged()
    
    If PBProgress.Value = PBProgress.Max Then
        isPlaying = False
        tmrPosition.Enabled = False
        mciSendString "stop mp3", 0, 0, 0
        mciSendString "close all", 0, 0, 0
        
        If Not lvPlaylist.SelectedItem.Index = lvPlaylist.ListItems.Count Then
        
            lvPlaylist.ListItems(lvPlaylist.SelectedItem.Index + 1).Selected = True
            lvPlaylist.ListItems(lvPlaylist.SelectedItem.Index + 1).EnsureVisible
            lvPlaylist_DblClick
        
        End If
        
    End If
    
End Sub

Private Sub pbVolume_ValueChanged()
    
    mciSendString "setaudio mp3 volume to " & pbVolume.Value, 0, 0, 0
    Label3.Caption = pbVolume.Value / 10 & "%"
    
End Sub

Private Sub TabStrip1_Click()

    Select Case TabStrip1.SelectedItem.Caption
        
        Case "Extractor"
            picEncode.Visible = False
            picPlayer.Visible = False
            picExtract.Visible = True
            
        Case "Encoder"
            picEncode.Visible = True
            picPlayer.Visible = False
            picExtract.Visible = False
            
        Case "Media Player"
            picEncode.Visible = False
            picPlayer.Visible = True
            picExtract.Visible = False
            
    End Select
    
End Sub

Private Sub tbEncode_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    Select Case Button.Caption
    
        Case "Encode"
            Dim I As Long
            
            If lvEncode.ListItems.Count = 0 Then Exit Sub
            
            PB3.Value = 0
            PB3.Max = lvEncode.ListItems.Count
            
            LogIT "Encoding started...", &H0&, True
            For I = 1 To lvEncode.ListItems.Count
                
                PB3.Value = I
                EncodeMP32 lvEncode.ListItems(I).Text, lvEncode.ListItems(I).ListSubItems(1).Text
                DoEvents
                
            Next I
            LogIT "Encoding complete.", &H0&, True
        
        Case "Settings"
            frmEncoderSettings.Show 1
            
    End Select
    
End Sub

Private Sub tbEncode_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    
    Dim listX As ListItem
    
    Select Case ButtonMenu.Key
        
        Case "file"
            Dim varTemp As Variant
            Dim intTemp As Integer
            
            CD1.Filter = "WAV Files (*.wav)|*.wav|"
            CD1.DialogTitle = "Select .Wav file"
            CD1.ShowOpen
            
            If Len(CD1.FileName) = 0 Then Exit Sub
            
            varTemp = Split(CD1.FileName, "\")
            intTemp = UBound(varTemp)
            
            Set listX = lvEncode.ListItems.Add(, , , , 14)
                listX.Text = GetShortName(CD1.FileName)
                listX.SubItems(1) = GetShortName(strPath) & "\" & Replace(varTemp(intTemp), "wav", "mp3")
                
        Case "folder"
            Dim lpIDList As Long
            Dim sBuffer As String
            Dim szTitle As String
            Dim tBrowseInfo As BrowseInfo
            Dim strFolder As String
            Dim I As Long
            
            With tBrowseInfo
               .hWndOwner = Me.hWnd
               .lpszTitle = lstrcat("Select folder with .Wav files...", "")
               .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
            End With
            
            lpIDList = SHBrowseForFolder(tBrowseInfo)
            
            If (lpIDList) Then
               sBuffer = Space(MAX_PATH)
               SHGetPathFromIDList lpIDList, sBuffer
               sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
               
                If Right(sBuffer, 1) = "\" Then
                    strFolder = Mid(sBuffer, 1, Len(sBuffer) - 1)
                Else
                    strFolder = sBuffer
                End If
            End If
            
            File1.Path = strFolder
            DoEvents
            
            For I = 0 To File1.ListCount - 1
                If Right(File1.List(I), 3) = "wav" Then
                    Set listX = lvEncode.ListItems.Add(, , , , 14)
                        listX.Text = GetShortName(strFolder) & "\" & File1.List(I)
                        listX.SubItems(1) = GetShortName(strPath) & "\" & Replace(File1.List(I), "wav", "mp3")
                End If
            Next I

        
    End Select
    
End Sub

Private Sub tbMain_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    Select Case Button.Key
        
        Case "About"
            frmAbout.Show 1
        
        Case "pSettings"
            frmSettings.Show 1
                        
        Case "eSettings"
            frmEncoderSettings.Show 1
            
        Case "stop"
            LogIT "  Stopping, Please wait..." & vbCrLf, &H0&, False
            blStop = True
            
        Case "wizard"
            frmWizard.Show , Me
        
    End Select
    
End Sub

Private Sub tbMain_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    
    Select Case ButtonMenu.Key
                
        Case "all"
            Dim I As Long
            For I = 1 To lvMain.ListItems.Count
                lvMain.ListItems(I).Checked = True
            Next I
            
            mnuExtractSelected_Click
        
        Case "selected"
            mnuExtractSelected_Click
        
        Case "info"
            If lvMain.ListItems.Count = 0 Then
                MsgBox "Please insert a CD or choose a drive from the list.", vbCritical
                Exit Sub
            End If
            
            Screen.MousePointer = 11
            frmMain.LogIT "Getting CD Info...", &H0&, True
            DoEvents
            
            CDDBQuery
            
            Screen.MousePointer = 1
            frmMain.LogIT "Done." & vbCrLf, &H0&, True
            
        Case "servers"
            Screen.MousePointer = 11
            LogIT "Getting servers..." & vbCrLf, &H0&, True
            DoEvents
            CDDBServerList
            LogIT "Done." & vbCrLf, &H0&, True
            Screen.MousePointer = 1
        
    End Select
    
End Sub

Public Sub CDDBQuery()

    Dim lngNum As Long
    Dim strCategory As String
    Dim strCDDBID As String
    Dim blnExactMatch As Boolean
    Dim strArtist As String
    Dim strTitle As String
    Dim strCDDBEntry As String
    Dim varTemp, varTitle As Variant
    Dim I As Long
    Dim lngLoop As Long
       
    On Error Resume Next
    
    If Val(intHandle) <> 0 Then
    
        AK.CDDBSetOption CDDB_OPT_SERVER, ReadINI("CDDB", "Server", App.Path & "\Config.ini"), 0
        AK.CDDBSetOption CDDB_OPT_CGI, ReadINI("CDDB", "CGI", App.Path & "\Config.ini"), 0
        AK.CDDBSetOption CDDB_OPT_USER, ReadINI("CDDB", "User", App.Path & "\Config.ini"), 0
        AK.CDDBSetOption CDDB_OPT_AGENT, "program 1.0", 0
        
        If AK.CDDBQuery(Val(intHandle), lngNum) = SS_COMP Then
        
            If lngNum > 0 Then
            
                For lngLoop = 0 To lngNum - 1
                    AK.CDDBQueryRead lngLoop, strCategory, strCDDBID, blnExactMatch, strArtist, strTitle
                    LogIT "  Artist: " & strArtist, &H0&, False
                    LogIT "  Title: " & strTitle, &H0&, False
                    
                    If AK.CDDBGetDiskInfo(lngLoop, strCDDBEntry) = SS_COMP Then

                    varTemp = Split(strCDDBEntry, vbCrLf)
                    For I = 0 To UBound(varTemp)
                        If LCase(Left(varTemp(I), 6)) = "ttitle" Then
                        varTitle = Split(varTemp(I), "=")
                        lvMain.ListItems(Mid(varTitle(0), 7) + 1).ListSubItems(1).Text = varTitle(1)
                        End If
                    Next I
                    
                    Erase varTitle
                    Erase varTemp

                    Else

                        MsgBox "Error Occured", vbCritical

                    End If
                Next lngLoop
                
            Else
            
                MsgBox "No Matches Found", vbCritical
                
            End If
            
        Else
        
            MsgBox "Error Occured", vbCritical
            
        End If
        
    Else
    
        MsgBox "We dont have a CD Handle", vbCritical
        
    End If

End Sub


Private Sub GrabTrackLBAEX(intTrack As Integer, wavFileName As String)

    Dim lngADR As Long
    Dim lngTrackNumber As Long
    Dim lngMins As Long
    Dim lngSecs As Long
    Dim lngFrames As Long
    Dim lngTrack As Long
    Dim lngStartFrame As Long
    Dim lngFrameCount As Long
    
    If Val(intHandle) <> 0 Then
    
        If AK.ModifyCDParms(Val(intHandle), CDP_MSF, True) = False Then
            MsgBox "Error Occured", vbCritical
        End If
        
        If AK.ReadTOC(Val(intHandle)) = SS_COMP Then
        
            lngTrack = Val(intTrack)
            
            If lngTrack <> 0 Then
            
                Call AK.ReadTOCTrack(lngTrack - 1, lngADR, lngTrackNumber, lngMins, lngSecs, lngFrames)
                lngStartFrame = (((lngMins * 60) + lngSecs) * 75) + lngFrames
                
                Call AK.ReadTOCTrack(lngTrack, lngADR, lngTrackNumber, lngMins, lngSecs, lngFrames)
                lngFrameCount = ((((lngMins * 60) + lngSecs) * 75) + lngFrames) - lngStartFrame
                
                Call AK.ReadTOCTrack(0, lngADR, lngTrackNumber, lngMins, lngSecs, lngFrames)
                lngStartFrame = lngStartFrame - ((((lngMins * 60) + lngSecs) * 75) + lngFrames)
                
                Call CreateWavEx(lngStartFrame, lngFrameCount, strPath & "\" & wavFileName & ".wav")
                
            End If
        
        Else
        
            MsgBox "Error Occured", vbCritical
            
        End If
    
    Else
    
        MsgBox "We dont have a CD Handle", vbCritical
        
    End If
  
End Sub

Private Sub CreateWav(lngFrameStart As Long, lngFrames As Long, strFileName As String)

    Dim lngReadFrames As Long
    Dim lngTotalFrames As Long
    Dim blnError As Boolean
    Dim lngPeak As Long
    Dim lngMaxPeak As Long
    
    Const BufFrames = 10
    
    blnError = False
    PB1.Value = 0
    PB1.Max = lngFrames
    PB2.Max = 100
    lngMaxPeak = 0
    lngTotalFrames = 0
    
    AK.ReadCDAudioLBAOpen strFileName
    
    Do
    
        If lngFrames - lngTotalFrames >= BufFrames Then
        
            lngReadFrames = BufFrames
            
        Else
        
            lngReadFrames = lngFrames Mod BufFrames
            
        End If
        
        If AK.ReadCDAudioLBAProcess(Val(intHandle), lngFrameStart + lngTotalFrames, lngReadFrames, lngPeak) <> SS_COMP Then
            
            blnError = True
            
        End If
        
        lngTotalFrames = lngTotalFrames + lngReadFrames
        PB1.Value = lngTotalFrames
        PB2.Value = lngPeak
        If lngPeak > lngMaxPeak Then lngMaxPeak = lngPeak
        DoEvents
        
        If blnError = True Then Exit Do
    
    Loop Until lngFrames = lngTotalFrames
    
    AK.ReadCDAudioLBAClose
    PB2.Value = 0
    
    If blnError = False Then
    
        'MsgBox "(MaxPeak = " & lngMaxPeak & "%) Completed Successfully - " & strFileName
'        If lngMaxPeak < 100 Then
'
'            If MsgBox("Do you wish to normalise this file", vbYesNo) = vbYes Then
'
'                AK.Normalise 100 + (100 - lngMaxPeak), strFileName, App.Path & "\Normalise.wav"
'                MsgBox "Normalise Completed Successfully - " & App.Path & "\Normalise.wav"
'
'            End If
'
'        End If
        
    Else
    
        MsgBox "Error Occured", vbCritical
        
    End If
    
End Sub

Private Sub CreateWavEx(lngFrameStart As Long, lngFrames As Long, strFileName As String)

    Dim lngReadFrames As Long
    Dim lngTotalFrames As Long
    Dim blnError As Boolean
    Dim lngPeak As Long
    Dim lngMaxPeak As Long
    Dim lngErrorCount As Long
    
    Const BufFrames = 10
    
    blnError = False
    PB1.Value = 0
    PB1.Max = lngFrames
    PB2.Max = 100
    lngMaxPeak = 0
    lngTotalFrames = 0
    
    AK.ReadCDAudioLBAExOpen strFileName
    
    Do
    
        If lngFrames - lngTotalFrames >= BufFrames Then
        
            lngReadFrames = BufFrames
            
        Else
        
            lngReadFrames = lngFrames Mod BufFrames
            
        End If
        
    
        lngErrorCount = 0
        Do
            blnError = False
            If AK.ReadCDAudioLBAExProcess(Val(intHandle), lngFrameStart + lngTotalFrames, lngReadFrames, lngPeak) <> SS_COMP Then
                
                lngErrorCount = lngErrorCount + 1
                blnError = True
                
            End If
            
        Loop Until blnError = False Or lngErrorCount = 10
        
        lngTotalFrames = lngTotalFrames + lngReadFrames
        PB1.Value = lngTotalFrames
        PB2.Value = lngPeak
        If lngPeak > lngMaxPeak Then lngMaxPeak = lngPeak
        DoEvents
        
        If blnError = True Then Exit Do
    
    Loop Until lngFrames = lngTotalFrames
    
    AK.ReadCDAudioLBAExClose
    PB2.Value = 0
    
    If blnError = False Then
    
        'MsgBox "(MaxPeak = " & lngMaxPeak & "%) Completed Successfully - " & strFileName
'        If lngMaxPeak < 100 Then
'
'            If MsgBox("Do you wish to normalise this file", vbYesNo) = vbYes Then
'
'                AK.Normalise 100 + (100 - lngMaxPeak), strFileName, App.Path & "\Normalise.wav"
'                MsgBox "Normalise Completed Successfully - " & App.Path & "\Normalise.wav"
'
'            End If
'
'        End If
        
    Else
    
        MsgBox "Error Occured", vbCritical
        
    End If
    
End Sub

Private Sub GrabTrackLBA(intTrack As Integer, wavFileName As String)

    Dim lngADR As Long
    Dim lngTrackNumber As Long
    Dim lngMins As Long
    Dim lngSecs As Long
    Dim lngFrames As Long
    Dim lngTrack As Long
    Dim lngStartFrame As Long
    Dim lngFrameCount As Long
    
    If Val(intHandle) <> 0 Then
    
        If AK.ModifyCDParms(Val(intHandle), CDP_MSF, True) = False Then
            MsgBox "Error Occured"
        End If
        
        If AK.ReadTOC(Val(intHandle)) = SS_COMP Then
        
        lngTrack = intTrack
        
            If lngTrack <> 0 Then
            
                Call AK.ReadTOCTrack(lngTrack - 1, lngADR, lngTrackNumber, lngMins, lngSecs, lngFrames)
                lngStartFrame = (((lngMins * 60) + lngSecs) * 75) + lngFrames
                
                Call AK.ReadTOCTrack(lngTrack, lngADR, lngTrackNumber, lngMins, lngSecs, lngFrames)
                lngFrameCount = ((((lngMins * 60) + lngSecs) * 75) + lngFrames) - lngStartFrame
                
                Call AK.ReadTOCTrack(0, lngADR, lngTrackNumber, lngMins, lngSecs, lngFrames)
                lngStartFrame = lngStartFrame - ((((lngMins * 60) + lngSecs) * 75) + lngFrames)
                
                Call CreateWav(lngStartFrame, lngFrameCount, strPath & "\" & wavFileName & ".wav")
                
            End If
        
        Else
        
            MsgBox "Error Occured", vbCritical
            
        End If
    
    Else
    
        MsgBox "We dont have a CD Handle", vbCritical
    
    End If
  
End Sub



Function Inst(ByVal idProc As Long) As Boolean
    
    Dim f As Long, hModule As Long, c As Long
    
    Inst = False
    If PId = "Win9x" Then
    
        Dim process As PROCESSENTRY32, module As MODULEENTRY32
        Dim hSnap As Long, idModule As Long
        
        hSnap = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
        
        If hSnap = hNull Then Exit Function
        
        ' Loop through to find matching process
        process.dwSize = Len(process)
        f = Process32First(hSnap, process)
        
        Do While f
            If process.th32ProcessID = idProc Then
                ' Save module ID
                Inst = True
                Exit Do
            End If
            f = Process32Next(hSnap, process)
        Loop
        
        CloseHandle hSnap
        
    ElseIf PId = "WinNT" Then
    
        ' First module is the main executable
        f = EnumProcessModules(ProcFromProcID(idProc), hModule, 4, c)
        
        If f = 0 Then Exit Function
        
        Dim modinfo As MODULEINFO
        
        f = GetModuleInformation(ProcFromProcID(idProc), hModule, modinfo, c)
        
        If f Then Inst = True
        
    End If
    
End Function

Function ProcFromProcID(idProc As Long) As Long

    ProcFromProcID = OpenProcess(PROCESS_QUERY_INFORMATION Or _
                                 PROCESS_VM_READ, 0, idProc)
                                 
End Function

Private Sub EncodeMP3(wavFile As String, mp3File As String)

    Dim AppID As Single
    Dim intBitrate As Integer
    Dim intMode As Integer
    Dim strFlags As String
    
    strFlags = ""
    
    If ReadINI("Lame", "Mode", App.Path & "\Config.ini") = 1 Then
        strFlags = strFlags & " -h"
    End If

    intBitrate = ReadINI("Lame", "Bitrate", App.Path & "\Config.ini")
    Select Case intBitrate

        Case 0: strFlags = strFlags & " -b 8"
        Case 1: strFlags = strFlags & " -b 16"
        Case 2: strFlags = strFlags & " -b 24"
        Case 3: strFlags = strFlags & " -b 32"
        Case 4: strFlags = strFlags & " -b 40"
        Case 5: strFlags = strFlags & " -b 48"
        Case 6: strFlags = strFlags & " -b 56"
        Case 7: strFlags = strFlags & " -b 64"
        Case 8: strFlags = strFlags & " -b 80"
        Case 9: strFlags = strFlags & " -b 96"
        Case 10: strFlags = strFlags & " -b 112"
        Case 11: strFlags = strFlags & " -b 128"
        Case 12: strFlags = strFlags & " -b 144"
        Case 13: strFlags = strFlags & " -b 160"
        Case 14: strFlags = strFlags & " -b 192"
        Case 15: strFlags = strFlags & " -b 224"
        Case 16: strFlags = strFlags & " -b 256"
        Case 17: strFlags = strFlags & " -b 320"

    End Select

    intMode = ReadINI("Lame", "Mode", App.Path & "\Config.ini")
    Select Case intMode

        Case 1: strFlags = strFlags & " -m s"
        Case 2: strFlags = strFlags & " -m j"
        Case 3: strFlags = strFlags & " -m f"
        Case 4: strFlags = strFlags & " -m m"

    End Select
    
    LogIT "    Encoding: " & mp3File & "...", &H80FF&, False
    DoEvents
    
    ExecCmd "lame.exe" & strFlags & " " & GetShortName(strPath) & "\" & wavFile & " " & GetShortName(strPath) & "\" & mp3File
    
    If ReadINI("Lame", "KillFile", App.Path & "\Config.ini") = 1 Then
        Kill strPath & "\" & wavFile
    End If
    
    LogIT "    Done.", &H80FF&, False

End Sub

Private Sub EncodeMP32(wavFile As String, mp3File As String)

    Dim AppID As Single
    Dim intBitrate As Integer
    Dim intMode As Integer
    Dim strFlags As String
    
    strFlags = ""
    
    If ReadINI("Lame", "Mode", App.Path & "\Config.ini") = 1 Then
        strFlags = strFlags & " -h"
    End If

    intBitrate = ReadINI("Lame", "Bitrate", App.Path & "\Config.ini")
    Select Case intBitrate

        Case 0: strFlags = strFlags & " -b 8"
        Case 1: strFlags = strFlags & " -b 16"
        Case 2: strFlags = strFlags & " -b 24"
        Case 3: strFlags = strFlags & " -b 32"
        Case 4: strFlags = strFlags & " -b 40"
        Case 5: strFlags = strFlags & " -b 48"
        Case 6: strFlags = strFlags & " -b 56"
        Case 7: strFlags = strFlags & " -b 64"
        Case 8: strFlags = strFlags & " -b 80"
        Case 9: strFlags = strFlags & " -b 96"
        Case 10: strFlags = strFlags & " -b 112"
        Case 11: strFlags = strFlags & " -b 128"
        Case 12: strFlags = strFlags & " -b 144"
        Case 13: strFlags = strFlags & " -b 160"
        Case 14: strFlags = strFlags & " -b 192"
        Case 15: strFlags = strFlags & " -b 224"
        Case 16: strFlags = strFlags & " -b 256"
        Case 17: strFlags = strFlags & " -b 320"

    End Select

    intMode = ReadINI("Lame", "Mode", App.Path & "\Config.ini")
    Select Case intMode

        Case 1: strFlags = strFlags & " -m s"
        Case 2: strFlags = strFlags & " -m j"
        Case 3: strFlags = strFlags & " -m f"
        Case 4: strFlags = strFlags & " -m m"

    End Select
    
    LogIT "    Encoding: " & mp3File & "...", &H80FF&, False
    DoEvents
    
    'LogIT "lame.exe" & strFlags & " " & wavFile & " " & mp3File, &H80FF&, False
    ExecCmd "lame.exe" & strFlags & " " & wavFile & " " & mp3File
    
    If ReadINI("Lame", "KillFile", App.Path & "\Config.ini") = 1 Then
        Kill wavFile
    End If
    
    LogIT "    Done.", &H80FF&, False

End Sub

Private Sub tmrPosition_Timer()
    
    'On Error Resume Next
    mciSendString "status mp3 position", sReturnBuffer$, Len(sReturnBuffer$), 0
    lblPosition.Caption = lngLength & "\" & sReturnBuffer$
    PBProgress.Value = sReturnBuffer$
    lblElapsedTime.Caption = ElapsedTime(strStartTime, Time)

End Sub
