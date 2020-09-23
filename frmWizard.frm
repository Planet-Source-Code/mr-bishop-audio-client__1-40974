VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmWizard 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Wizard"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7620
   Icon            =   "frmWizard.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   7620
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      Height          =   3735
      Left            =   0
      Picture         =   "frmWizard.frx":038A
      ScaleHeight     =   3675
      ScaleWidth      =   2595
      TabIndex        =   0
      Top             =   0
      Width           =   2655
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   1440
         Picture         =   "frmWizard.frx":1E3B
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   13
         Top             =   3240
         Width           =   255
      End
      Begin VB.Line Line10 
         BorderColor     =   &H00808080&
         X1              =   120
         X2              =   1320
         Y1              =   3360
         Y2              =   3360
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         X1              =   1440
         X2              =   2400
         Y1              =   2640
         Y2              =   2640
      End
      Begin VB.Label lblStep 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "One"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   720
         TabIndex        =   3
         Top             =   2400
         Width           =   615
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Step"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   360
         Left            =   120
         TabIndex        =   2
         Top             =   2400
         Width           =   645
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "wizard"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   240
         Left            =   1800
         TabIndex        =   1
         Top             =   3240
         Width           =   675
      End
   End
   Begin VB.PictureBox picOne 
      BorderStyle     =   0  'None
      Height          =   3495
      Left            =   2760
      ScaleHeight     =   3495
      ScaleWidth      =   4815
      TabIndex        =   4
      Top             =   120
      Width           =   4815
      Begin MSComctlLib.ImageList imgMain 
         Left            =   4200
         Top             =   120
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmWizard.frx":21C5
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.ComboBox cmbDrives 
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   360
         Width           =   4455
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   3960
         TabIndex        =   6
         Top             =   3120
         Width           =   855
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "&Next >>"
         Height          =   375
         Left            =   3000
         TabIndex        =   5
         Top             =   3120
         Width           =   855
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "What drive would you like to extract data from?"
         Height          =   195
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Width           =   3330
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FFFFFF&
         X1              =   -120
         X2              =   4800
         Y1              =   3015
         Y2              =   3015
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00404040&
         X1              =   0
         X2              =   4800
         Y1              =   3000
         Y2              =   3000
      End
   End
   Begin VB.PictureBox picThree 
      BorderStyle     =   0  'None
      Height          =   3495
      Left            =   2760
      ScaleHeight     =   3495
      ScaleWidth      =   4815
      TabIndex        =   9
      Top             =   120
      Visible         =   0   'False
      Width           =   4815
      Begin VB.CommandButton cmdNext1 
         Caption         =   "&Next >>"
         Height          =   375
         Left            =   3000
         TabIndex        =   11
         Top             =   3120
         Width           =   855
      End
      Begin VB.CommandButton cmdBack 
         Caption         =   "<< &Back"
         Height          =   375
         Left            =   2160
         TabIndex        =   12
         Top             =   3120
         Width           =   855
      End
      Begin VB.CommandButton cmdCancel1 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   3960
         TabIndex        =   10
         Top             =   3120
         Width           =   855
      End
      Begin MSComctlLib.ListView lvMain 
         Height          =   2520
         Left            =   0
         TabIndex        =   14
         Top             =   360
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   4445
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         SmallIcons      =   "imgMain"
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Track #"
            Object.Width           =   1482
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Track name"
            Object.Width           =   6420
         EndProperty
         Picture         =   "frmWizard.frx":255F
      End
      Begin VB.Label lblSelectAll 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Select All"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   3960
         TabIndex        =   16
         Top             =   0
         Width           =   825
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "What tracks would you like to extract?"
         Height          =   195
         Left            =   0
         TabIndex        =   15
         Top             =   0
         Width           =   2715
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00404040&
         X1              =   0
         X2              =   4800
         Y1              =   3000
         Y2              =   3000
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00FFFFFF&
         X1              =   -120
         X2              =   4800
         Y1              =   3015
         Y2              =   3015
      End
   End
   Begin VB.PictureBox picFour 
      BorderStyle     =   0  'None
      Height          =   3495
      Left            =   2760
      ScaleHeight     =   3495
      ScaleWidth      =   4815
      TabIndex        =   25
      Top             =   120
      Visible         =   0   'False
      Width           =   4815
      Begin VB.CommandButton Command2 
         Caption         =   "Encoder settings"
         Height          =   375
         Left            =   3360
         TabIndex        =   31
         Top             =   600
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Program settings"
         Height          =   375
         Left            =   1920
         TabIndex        =   30
         Top             =   600
         Width           =   1335
      End
      Begin VB.CommandButton cmdNext3 
         Caption         =   "&Start!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3000
         TabIndex        =   26
         Top             =   3120
         Width           =   855
      End
      Begin VB.CommandButton cmdBack2 
         Caption         =   "<< &Back"
         Height          =   375
         Left            =   2160
         TabIndex        =   28
         Top             =   3120
         Width           =   855
      End
      Begin VB.CommandButton cmdCancel3 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   3960
         TabIndex        =   27
         Top             =   3120
         Width           =   855
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmWizard.frx":2FF3
         Height          =   555
         Left            =   0
         TabIndex        =   29
         Top             =   0
         Width           =   4860
      End
      Begin VB.Line Line9 
         BorderColor     =   &H00404040&
         X1              =   0
         X2              =   4800
         Y1              =   3000
         Y2              =   3000
      End
      Begin VB.Line Line8 
         BorderColor     =   &H00FFFFFF&
         X1              =   -120
         X2              =   4800
         Y1              =   3015
         Y2              =   3015
      End
   End
   Begin VB.PictureBox picTwo 
      BorderStyle     =   0  'None
      Height          =   3495
      Left            =   2760
      ScaleHeight     =   3495
      ScaleWidth      =   4815
      TabIndex        =   17
      Top             =   120
      Visible         =   0   'False
      Width           =   4815
      Begin VB.PictureBox picAlert 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   0
         Picture         =   "frmWizard.frx":307D
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   24
         Top             =   2160
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.ComboBox cmbCDDB 
         Height          =   315
         Left            =   3720
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   480
         Width           =   975
      End
      Begin VB.CommandButton cmdNext2 
         Caption         =   "&Next >>"
         Height          =   375
         Left            =   3000
         TabIndex        =   18
         Top             =   3120
         Width           =   855
      End
      Begin VB.CommandButton cmdCancel2 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   3960
         TabIndex        =   20
         Top             =   3120
         Width           =   855
      End
      Begin VB.CommandButton cmdBack1 
         Caption         =   "<< &Back"
         Height          =   375
         Left            =   2160
         TabIndex        =   19
         Top             =   3120
         Width           =   855
      End
      Begin VB.Label lblALert 
         BackStyle       =   0  'Transparent
         Caption         =   "Trying to download CD information from server. This could take a minute or two depending on your connection speed."
         ForeColor       =   &H00000080&
         Height          =   615
         Left            =   360
         TabIndex        =   23
         Top             =   2160
         Visible         =   0   'False
         Width           =   3615
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Would you like to try and download the information for this CD from CDDB?"
         Height          =   435
         Left            =   0
         TabIndex        =   21
         Top             =   0
         Width           =   4695
      End
      Begin VB.Line Line7 
         BorderColor     =   &H00FFFFFF&
         X1              =   -120
         X2              =   4800
         Y1              =   3015
         Y2              =   3015
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00404040&
         X1              =   0
         X2              =   4800
         Y1              =   3000
         Y2              =   3000
      End
   End
End
Attribute VB_Name = "frmWizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbCDDB_Click()
    
    If cmbCDDB.Text = "Yes" Then
    
        picAlert.Visible = True
        lblALert.Visible = True
        DoEvents
        
        Screen.MousePointer = 11
        
        frmMain.CDDBQuery
        picAlert.Visible = False
        lblALert.Visible = False
        cmbCDDB.ListIndex = 1
        cmdNext2_Click
        
        Screen.MousePointer = 1
        
    End If
    
End Sub

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
            frmMain.ReadTracks
        End If
    Else
        MsgBox "Already have a CD Handle", vbCritical
    End If

End Sub

Private Sub cmdBack_Click()
    
    lblStep.Caption = "Two"
    picTwo.Visible = True
    picThree.Visible = False
    
End Sub

Private Sub cmdBack1_Click()

    lblStep.Caption = "One"
    picOne.Visible = True
    picTwo.Visible = False

End Sub

Private Sub cmdBack2_Click()
    
    lblStep.Caption = "Three"
    picThree.Visible = True
    picFour.Visible = False
    
End Sub

Private Sub cmdCancel_Click()
    
    Unload Me
    
End Sub

Private Sub cmdCancel1_Click()
    
    Unload Me
    
End Sub

Private Sub cmdCancel2_Click()
    
    Unload Me
    
End Sub

Private Sub cmdCancel3_Click()
    
    Unload Me
    
End Sub

Private Sub cmdNext_Click()
           
    lblStep.Caption = "Two"
    picOne.Visible = False
    picTwo.Visible = True
    
End Sub

Private Sub cmdNext1_Click()
        
    Dim intSelected As Long
    Dim i As Long
    
    intSelected = 0
    For i = 1 To lvMain.ListItems.Count
        frmMain.lvMain.ListItems(i).Checked = lvMain.ListItems(i).Checked
        If lvMain.ListItems(i).Checked = True Then
            intSelected = intSelected + 1
        End If
    Next i
    
    If intSelected = 0 Then
        MsgBox "Please select at least 1 track.", vbCritical
        Exit Sub
    End If
    
    lblStep.Caption = "Four"
    picThree.Visible = False
    picFour.Visible = True
    
End Sub

Private Sub cmdNext2_Click()

    Dim i As Long
    Dim listX As ListItem
    
    lvMain.ListItems.Clear
    For i = 1 To frmMain.lvMain.ListItems.Count
        
        Set listX = lvMain.ListItems.Add(, , , , 1)
            listX.Text = frmMain.lvMain.ListItems(i).Text
            listX.SubItems(1) = frmMain.lvMain.ListItems(i).ListSubItems(1).Text
        
    Next i
    
    lblStep.Caption = "Three"
    picTwo.Visible = False
    picThree.Visible = True

End Sub

Private Sub cmdNext3_Click()
    
    frmMain.mnuExtractSelected_Click
    
End Sub

Private Sub Command1_Click()
    
    frmSettings.Show 1
    
End Sub

Private Sub Command2_Click()

    frmEncoderSettings.Show 1

End Sub

Private Sub Form_Load()
        
    Dim i As Long
    For i = 0 To frmMain.cmbDrives.ListCount - 1
        cmbDrives.AddItem frmMain.cmbDrives.List(i)
    Next i
    cmbDrives.ListIndex = 0
    
    cmbCDDB.AddItem "Yes"
    cmbCDDB.AddItem "No"
    cmbCDDB.ListIndex = 1
    
    doWizard = True
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    doWizard = False
    
End Sub

Private Sub lblSelectAll_Click()
    
    Dim i As Long
    For i = 1 To lvMain.ListItems.Count
        
        lvMain.ListItems(i).Checked = True
        
    Next i
    
End Sub
