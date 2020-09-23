Attribute VB_Name = "modAKRip32"
Option Explicit

' Aspi errors

Public Const ALERR_NOERROR = 0
Public Const ALERR_NOWNASPI = 1
Public Const ALERR_NOGETASPI32SUPP = 2
Public Const ALERR_NOSENDASPICMD = 3
Public Const ALERR_ASPI = 4
Public Const ALERR_NOCDSELECTED = 5
Public Const ALERR_BUFTOOSMALL = 6
Public Const ALERR_INVHANDLE = 7
Public Const ALERR_NOMOREHAND = 8
Public Const ALERR_BUFPTR = 9
Public Const ALERR_NOTACD = 10
Public Const ALERR_LOCK = 11
Public Const ALERR_DUPHAND = 12
Public Const ALERR_INVPTR = 13
Public Const ALERR_INVPARM = 14
Public Const ALERR_JITTER = 15

' constants used for queryCDParms()
  
Public Const CDP_READCDR = &H1     ' can read CD-R
Public Const CDP_READCDE = &H2     ' can read CD-E
Public Const CDP_METHOD2 = &H3     ' can read CD-R wriiten via method 2
Public Const CDP_WRITECDR = &H4     ' can write CD-R
Public Const CDP_WRITECDE = &H5     ' can write CD-E
Public Const CDP_AUDIOPLAY = &H6     ' can play audio
Public Const CDP_COMPOSITE = &H7     ' composite audio/video stream
Public Const CDP_DIGITAL1 = &H8     ' digital output (IEC958) on port 1
Public Const CDP_DIGITAL2 = &H9     ' digital output (IEC958) on port 2
Public Const CDP_M2FORM1 = &HA     ' reads Mode 2 Form 1 (XA) format
Public Const CDP_M2FORM2 = &HB     ' reads Mode 2 Form 2 format
Public Const CDP_MULTISES = &HC     ' reads multi-session or Photo-CD
Public Const CDP_CDDA = &HD     ' supports cd-da
Public Const CDP_STREAMACC = &HE     ' supports "stream is accurate"
Public Const CDP_RW = &HF     ' can return R-W info
Public Const CDP_RWCORR = &H10    ' returns R-W de-interleaved and err. corrected
Public Const CDP_C2SUPP = &H11   ' C2 error pointers
Public Const CDP_ISRC = &H12    ' can return the ISRC info
Public Const CDP_UPC = &H13   ' can return the Media Catalog Number
Public Const CDP_CANLOCK = &H14    ' prevent/allow cmd. can lock the media
Public Const CDP_LOCKED = &H15    ' current lock state (TRUE = LOCKED)
Public Const CDP_PREVJUMP = &H16    ' prevent/allow jumper state
Public Const CDP_CANEJECT = &H17    ' drive can eject disk
Public Const CDP_MECHTYPE = &H18    ' type of disk loading supported
Public Const CDP_SEPVOL = &H19    ' independent audio level for channels
Public Const CDP_SEPMUTE = &H1A    ' independent mute for channels
Public Const CDP_SDP = &H1B    ' supports disk present (SDP)
Public Const CDP_SSS = &H1C    ' Software Slot Selection
Public Const CDP_MAXSPEED = &H1D    ' maximum supported speed of drive
Public Const CDP_NUMVOL = &H1E    ' number of volume levels
Public Const CDP_BUFSIZE = &H1F    ' size of output buffer
Public Const CDP_CURRSPEED = &H20    ' current speed of drive
Public Const CDP_SPM = &H21    ' "S" units per "M" (MSF format)
Public Const CDP_FPS = &H22    ' "F" units per "S" (MSF format)
Public Const CDP_INACTMULT = &H23    ' inactivity multiplier ( x 125 ms)
Public Const CDP_MSF = &H24    ' use MSF format for READ TOC cmd
Public Const CDP_OVERLAP = &H25    ' number of overlap frames for jitter
Public Const CDP_JITTER = &H26    ' number of frames to check for jitter
Public Const CDP_READMODE = &H27    ' mode to attempt jitter corr.

' defines for GETCDHAND  readType

Public Const CDR_ANY = &H0   ' unknown
Public Const CDR_ATAPI1 = &H1   ' ATAPI per spec
Public Const CDR_ATAPI2 = &H2   ' alternate ATAPI
Public Const CDR_READ6 = &H3   ' using SCSI READ(6)
Public Const CDR_READ10 = &H4   ' using SCSI READ(10)
Public Const CDR_READ_D8 = &H5   ' using command 0xD8 (Plextor?)
Public Const CDR_READ_D4 = &H6   ' using command 0xD4 (NEC?)
Public Const CDR_READ_D4_1 = &H7   ' 0xD4 with a mode select
Public Const CDR_READ10_2 = &H8   ' different mode select w/ READ(10)

' defines for the read mode (CDP_READMODE)
 
Public Const CDRM_NOJITTER = &H0   ' never jitter correct
Public Const CDRM_JITTER = &H1   ' always jitter correct
Public Const CDRM_JITTERONERR = &H2   ' jitter correct only after a read error

' Used by CDDBOptions

Public Const CDDB_NONE = 0
Public Const CDDB_QUERY = 1
Public Const CDDB_ENTRY = 2
Public Const CDDB_OPT_SERVER = 0
Public Const CDDB_OPT_PROXY = 1
Public Const CDDB_OPT_USEPROXY = 2
Public Const CDDB_OPT_AGENT = 3
Public Const CDDB_OPT_USER = 4
Public Const CDDB_OPT_PROXYPORT = 5
Public Const CDDB_OPT_CGI = 6
Public Const CDDB_OPT_HTTPPORT = 7
Public Const CDDB_OPT_USECDPLAYERINI = 8

Public strPath As String

Public Const BIF_RETURNONLYFSDIRS = 1
Public Const BIF_DONTGOBELOWDOMAIN = 2
Public Const MAX_PATH = 260

Public Declare Function SHBrowseForFolder Lib _
"shell32" (lpbi As BrowseInfo) As Long

Public Declare Function SHGetPathFromIDList Lib _
"shell32" (ByVal pidList As Long, ByVal lpBuffer _
As String) As Long

Public Declare Function lstrcat Lib "kernel32" _
Alias "lstrcatA" (ByVal lpString1 As String, ByVal _
lpString2 As String) As Long

Public AK As AKRipAX.AKRip
Public intHandle, intOverlap, intCDAdd0, intCDAdd1, intCDAdd2, intCDParam0, intCDParam1, intStartFrame, intFrameLen As Integer
Public PId As String
Public blStop As Boolean


Public Type BrowseInfo
   hWndOwner As Long
   pIDLRoot As Long
   pszDisplayName As Long
   lpszTitle As Long
   ulFlags As Long
   lpfnCallback As Long
   lParam As Long
   iImage As Long
End Type

