Attribute VB_Name = "modScsiDefs"
Option Explicit

'****************************************************************************
'**
'** Name: scsidefs.h
'**
'** Description: SCSI definitions ('C' Language)
'**
'** Translated for VB by Peter Tribe (e.quinox@virgin.net)
'**
'** History of VB version:
'**
'** 06 July 2001 : First released version
'**
'****************************************************************************
'**
'** This program is free software you can redistribute it and/or modify
'** it under the terms of the GNU General Public License as published by
'** the Free Software Foundation either version 2 of the License, or
'** (at your option) any later version.
'**
'** This program is distributed in the hope that it will be useful,
'** but WITHOUT ANY WARRANTY without even the implied warranty of
'** MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'** GNU General Public License for more details.
'**
'** You should have received a copy of the GNU General Public License
'** along with this program if not, write to the Free Software
'** Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA 02111-1307 USA
'**
'***************************************************************************

'***************************************************************************
'** TARGET STATUS VALUES
'***************************************************************************
Public Const STATUS_GOOD = &H0         ' Status Good
Public Const STATUS_CHKCOND = &H2      ' Check Condition
Public Const STATUS_CONDMET = &H4      ' Condition Met
Public Const STATUS_BUSY = &H8         ' Busy
Public Const STATUS_INTERM = &H10      ' Intermediate
Public Const STATUS_INTCDMET = &H14    ' Intermediate-condition met
Public Const STATUS_RESCONF = &H18     ' Reservation conflict
Public Const STATUS_COMTERM = &H22     ' Command Terminated
Public Const STATUS_QFULL = &H28       ' Queue full

'***************************************************************************
'** SCSI MISCELLANEOUS EQUATES
'***************************************************************************
Public Const MAXLUN = 7                ' Maximum Logical Unit Id
Public Const MAXTARG = 7               ' Maximum Target Id
Public Const MAX_SCSI_LUNS = 64        ' Maximum Number of SCSI LUNs
Public Const MAX_NUM_HA = 8            ' Maximum Number of SCSI HA's

'\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
'
'   SCSI COMMAND OPCODES
'
'/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/

'***************************************************************************
'** Commands for all Device Types
'***************************************************************************
Public Const SCSI_CHANGE_DEF = &H40    ' Change Definition (Optional)
Public Const SCSI_COMPARE = &H39       ' Compare (O)
Public Const SCSI_COPY = &H18          ' Copy (O)
Public Const SCSI_COP_VERIFY = &H3A    ' Copy and Verify (O)
Public Const SCSI_INQUIRY = &H12       ' Inquiry (MANDATORY)
Public Const SCSI_LOG_SELECT = &H4C    ' Log Select (O)
Public Const SCSI_LOG_SENSE = &H4D     ' Log Sense (O)
Public Const SCSI_MODE_SEL6 = &H15     ' Mode Select 6-byte (Device Specific)
Public Const SCSI_MODE_SEL10 = &H55    ' Mode Select 10-byte (Device Specific)
Public Const SCSI_MODE_SEN6 = &H1A     ' Mode Sense 6-byte (Device Specific)
Public Const SCSI_MODE_SEN10 = &H5A    ' Mode Sense 10-byte (Device Specific)
Public Const SCSI_READ_BUFF = &H3C     ' Read Buffer (O)
Public Const SCSI_REQ_SENSE = &H3      ' Request Sense (MANDATORY)
Public Const SCSI_SEND_DIAG = &H1D     ' Send Diagnostic (O)
Public Const SCSI_TST_U_RDY = &H0      ' Test Unit Ready (MANDATORY)
Public Const SCSI_WRITE_BUFF = &H3B    ' Write Buffer (O)

'***************************************************************************
'** Commands Unique to Direct Access Devices
'***************************************************************************
Public Const SCSI_FORMAT = &H4         ' Format Unit (MANDATORY)
Public Const SCSI_LCK_UN_CAC = &H36    ' Lock Unlock Cache (O)
Public Const SCSI_PREFETCH = &H34      ' Prefetch (O)
Public Const SCSI_MED_REMOVL = &H1E    ' Prevent/Allow medium Removal (O)
Public Const SCSI_READ6 = &H8          ' Read 6-byte (MANDATORY)
Public Const SCSI_READ10 = &H28        ' Read 10-byte (MANDATORY)
Public Const SCSI_RD_CAPAC = &H25      ' Read Capacity (MANDATORY)
Public Const SCSI_RD_DEFECT = &H37     ' Read Defect Data (O)
Public Const SCSI_READ_LONG = &H3E     ' Read Long (O)
Public Const SCSI_REASS_BLK = &H7      ' Reassign Blocks (O)
Public Const SCSI_RCV_DIAG = &H1C      ' Receive Diagnostic Results (O)
Public Const SCSI_RELEASE = &H17       ' Release Unit (MANDATORY)
Public Const SCSI_REZERO = &H1         ' Rezero Unit (O)
Public Const SCSI_SRCH_DAT_E = &H31    ' Search Data Equal (O)
Public Const SCSI_SRCH_DAT_H = &H30    ' Search Data High (O)
Public Const SCSI_SRCH_DAT_L = &H32    ' Search Data Low (O)
Public Const SCSI_SEEK6 = &HB          ' Seek 6-Byte (O)
Public Const SCSI_SEEK10 = &H2B        ' Seek 10-Byte (O)
Public Const SCSI_SET_LIMIT = &H33     ' Set Limits (O)
Public Const SCSI_START_STP = &H1B     ' Start/Stop Unit (O)
Public Const SCSI_SYNC_CACHE = &H35    ' Synchronize Cache (O)
Public Const SCSI_VERIFY = &H2F        ' Verify (O)
Public Const SCSI_WRITE6 = &HA         ' Write 6-Byte (MANDATORY)
Public Const SCSI_WRITE10 = &H2A       ' Write 10-Byte (MANDATORY)
Public Const SCSI_WRT_VERIFY = &H2E    ' Write and Verify (O)
Public Const SCSI_WRITE_LONG = &H3F    ' Write Long (O)
Public Const SCSI_WRITE_SAME = &H41    ' Write Same (O)

'***************************************************************************
'** Commands Unique to Sequential Access Devices
'***************************************************************************
Public Const SCSI_ERASE = &H19         ' Erase (MANDATORY)
Public Const SCSI_LOAD_UN = &H1B       ' Load/Unload (O)
Public Const SCSI_LOCATE = &H2B        ' Locate (O)
Public Const SCSI_RD_BLK_LIM = &H5     ' Read Block Limits (MANDATORY)
Public Const SCSI_READ_POS = &H34      ' Read Position (O)
Public Const SCSI_READ_REV = &HF       ' Read Reverse (O)
Public Const SCSI_REC_BF_DAT = &H14    ' Recover Buffer Data (O)
Public Const SCSI_RESERVE = &H16       ' Reserve Unit (MANDATORY)
Public Const SCSI_REWIND = &H1         ' Rewind (MANDATORY)
Public Const SCSI_SPACE = &H11         ' Space (MANDATORY)
Public Const SCSI_VERIFY_T = &H13      ' Verify (Tape) (O)
Public Const SCSI_WRT_FILE = &H10      ' Write Filemarks (MANDATORY)

'***************************************************************************
'** Commands Unique to Printer Devices
'***************************************************************************
Public Const SCSI_PRINT = &HA          ' Print (MANDATORY)
Public Const SCSI_SLEW_PNT = &HB       ' Slew and Print (O)
Public Const SCSI_STOP_PNT = &H1B      ' Stop Print (O)
Public Const SCSI_SYNC_BUFF = &H10     ' Synchronize Buffer (O)

'***************************************************************************
'**Commands Unique to Processor Devices
'***************************************************************************
Public Const SCSI_RECEIVE = &H8        ' Receive (O)
Public Const SCSI_SEND = &HA           ' Send (O)

'***************************************************************************
'** Commands Unique to Write-Once Devices
'***************************************************************************
Public Const SCSI_MEDIUM_SCN = &H38    ' Medium Scan (O)
Public Const SCSI_SRCHDATE10 = &H31    ' Search Data Equal 10-Byte (O)
Public Const SCSI_SRCHDATE12 = &HB1    ' Search Data Equal 12-Byte (O)
Public Const SCSI_SRCHDATH10 = &H30    ' Search Data High 10-Byte (O)
Public Const SCSI_SRCHDATH12 = &HB0    ' Search Data High 12-Byte (O)
Public Const SCSI_SRCHDATL10 = &H32    ' Search Data Low 10-Byte (O)
Public Const SCSI_SRCHDATL12 = &HB2    ' Search Data Low 12-Byte (O)
Public Const SCSI_SET_LIM_10 = &H33    ' Set Limits 10-Byte (O)
Public Const SCSI_SET_LIM_12 = &HB3    ' Set Limits 10-Byte (O)
Public Const SCSI_VERIFY10 = &H2F      ' Verify 10-Byte (O)
Public Const SCSI_VERIFY12 = &HAF      ' Verify 12-Byte (O)
Public Const SCSI_WRITE12 = &HAA       ' Write 12-Byte (O)
Public Const SCSI_WRT_VER10 = &H2E     ' Write and Verify 10-Byte (O)
Public Const SCSI_WRT_VER12 = &HAE     ' Write and Verify 12-Byte (O)

'***************************************************************************
'** Commands Unique to CD-ROM Devices
'***************************************************************************
Public Const SCSI_PLAYAUD_10 = &H45    ' Play Audio 10-Byte (O)
Public Const SCSI_PLAYAUD_12 = &HA5    ' Play Audio 12-Byte 12-Byte (O)
Public Const SCSI_PLAYAUDMSF = &H47    ' Play Audio MSF (O)
Public Const SCSI_PLAYA_TKIN = &H48    ' Play Audio Track/Index (O)
Public Const SCSI_PLYTKREL10 = &H49    ' Play Track Relative 10-Byte (O)
Public Const SCSI_PLYTKREL12 = &HA9    ' Play Track Relative 12-Byte (O)
Public Const SCSI_READCDCAP = &H25     ' Read CD-ROM Capacity (MANDATORY)
Public Const SCSI_READHEADER = &H44    ' Read Header (O)
Public Const SCSI_SUBCHANNEL = &H42    ' Read Subchannel (O)
Public Const SCSI_READ_TOC = &H43      ' Read TOC (O)

'***************************************************************************
'** Commands Unique to Scanner Devices
'***************************************************************************
Public Const SCSI_GETDBSTAT = &H34     ' Get Data Buffer Status (O)
Public Const SCSI_GETWINDOW = &H25     ' Get Window (O)
Public Const SCSI_OBJECTPOS = &H31     ' Object Postion (O)
Public Const SCSI_SCAN = &H1B          ' Scan (O)
Public Const SCSI_SETWINDOW = &H24     ' Set Window (MANDATORY)

'***************************************************************************
'** Commands Unique to Optical Memory Devices
'***************************************************************************
Public Const SCSI_UpdateBlk = &H3D     ' Update Block (O)

'***************************************************************************
'** Commands Unique to Medium Changer Devices
'***************************************************************************
Public Const SCSI_EXCHMEDIUM = &HA6    ' Exchange Medium (O)
Public Const SCSI_INITELSTAT = &H7     ' Initialize Element Status (O)
Public Const SCSI_POSTOELEM = &H2B     ' Position to Element (O)
Public Const SCSI_REQ_VE_ADD = &HB5    ' Request Volume Element Address (O)
Public Const SCSI_SENDVOLTAG = &HB6    ' Send Volume Tag (O)

'***************************************************************************
'** Commands Unique to Communication Devices
'***************************************************************************
Public Const SCSI_GET_MSG_6 = &H8      ' Get Message 6-Byte (MANDATORY)
Public Const SCSI_GET_MSG_10 = &H28    ' Get Message 10-Byte (O)
Public Const SCSI_GET_MSG_12 = &HA8    ' Get Message 12-Byte (O)
Public Const SCSI_SND_MSG_6 = &HA      ' Send Message 6-Byte (MANDATORY)
Public Const SCSI_SND_MSG_10 = &H2A    ' Send Message 10-Byte (O)
Public Const SCSI_SND_MSG_12 = &HAA    ' Send Message 12-Byte (O)

'\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
'
'   END OF SCSI COMMAND OPCODES
'
'/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/

'***************************************************************************
'** REQUEST SENSE ERROR CODE
'***************************************************************************
Public Const SERROR_CURRENT = &H70     ' Current Errors
Public Const SERROR_DEFERED = &H71     ' Deferred Errors

'***************************************************************************
'** REQUEST SENSE BIT DEFINITIONS
'***************************************************************************
Public Const SENSE_VALID = &H80        ' Byte 0 Bit 7
Public Const SENSE_FILEMRK = &H80      ' Byte 2 Bit 7
Public Const SENSE_EOM = &H40          ' Byte 2 Bit 6
Public Const SENSE_ILI = &H20          ' Byte 2 Bit 5

'***************************************************************************
'** REQUEST SENSE SENSE KEY DEFINITIONS
'***************************************************************************
Public Const KEY_NOSENSE = &H0         ' No Sense
Public Const KEY_RECERROR = &H1        ' Recovered Error
Public Const KEY_NOTREADY = &H2        ' Not Ready
Public Const KEY_MEDIUMERR = &H3       ' Medium Error
Public Const KEY_HARDERROR = &H4       ' Hardware Error
Public Const KEY_ILLGLREQ = &H5        ' Illegal Request
Public Const KEY_UNITATT = &H6         ' Unit Attention
Public Const KEY_DATAPROT = &H7        ' Data Protect
Public Const KEY_BLANKCHK = &H8        ' Blank Check
Public Const KEY_VENDSPEC = &H9        ' Vendor Specific
Public Const KEY_COPYABORT = &HA       ' Copy Abort
Public Const KEY_EQUAL = &HC           ' Equal (Search)
Public Const KEY_VOLOVRFLW = &HD       ' Volume Overflow
Public Const KEY_MISCOMP = &HE         ' Miscompare (Search)
Public Const KEY_RESERVED = &HF        ' Reserved

'***************************************************************************
'** PERIPHERAL DEVICE TYPE DEFINITIONS
'***************************************************************************
Public Const DTYPE_DASD = &H0          ' Disk Device
Public Const DTYPE_SEQD = &H1          ' Tape Device
Public Const DTYPE_PRNT = &H2          ' Printer
Public Const DTYPE_PROC = &H3          ' Processor
Public Const DTYPE_WORM = &H4          ' Write-once read-multiple
Public Const DTYPE_CROM = &H5          ' CD-ROM device
Public Const DTYPE_CDROM = &H5         ' CD-ROM device
Public Const DTYPE_SCAN = &H6          ' Scanner device
Public Const DTYPE_OPTI = &H7          ' Optical memory device
Public Const DTYPE_JUKE = &H8          ' Medium Changer device
Public Const DTYPE_COMM = &H9          ' Communications device
Public Const DTYPE_RESL = &HA          ' Reserved (low)
Public Const DTYPE_RESH = &H1E         ' Reserved (high)
Public Const DTYPE_UNKNOWN = &H1F      ' Unknown or no device type

'***************************************************************************
'** ANSI APPROVED VERSION DEFINITIONS
'***************************************************************************
Public Const ANSI_MAYBE = &H0          ' Device may or may not be ANSI approved stand
Public Const ANSI_SCSI1 = &H1          ' Device complies to ANSI X3.131-1986 (SCSI-1)
Public Const ANSI_SCSI2 = &H2          ' Device complies to SCSI-2
Public Const ANSI_RESLO = &H3          ' Reserved (low)
Public Const ANSI_RESHI = &H7          ' Reserved (high)
