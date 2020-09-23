Attribute VB_Name = "modAspi32"
Option Explicit

'****************************************************************************
'**
'** Name: myaspi32.h - Copyright (C) 1999 Jay A. Key
'**
'** API for WNASPI32.DLL
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
'** Module Name:    modAspi32.bas
'**
'** Description:    Header file replacement for wnaspi32.h
'**
'***************************************************************************

'***************************************************************************
'** SCSI MISCELLANEOUS EQUATES
'***************************************************************************
Public Const SENSE_LEN = 14                       ' Default sense buffer length
Public Const SRB_DIR_SCSI = &H0                   ' Direction determined by SCSI
Public Const SRB_POSTING = &H1                    ' Enable ASPI posting
Public Const SRB_ENABLE_RESIDUAL_COUNT = &H4      ' Enable residual byte count reporting
Public Const SRB_DIR_IN = &H8                     ' Transfer from SCSI target to host
Public Const SRB_DIR_OUT = &H10                   ' Transfer from host to SCSI target
Public Const SRB_EVENT_NOTIFY = &H40              ' Enable ASPI event notification
Public Const RESIDUAL_COUNT_SUPPORTED = &H2       ' Extended buffer flag
Public Const MAX_SRB_TIMEOUT = 1080001            ' 30 hour maximum timeout in sec
Public Const DEFAULT_SRB_TIMEOUT = 1080001        ' use max.timeout by default

'***************************************************************************
'** ASPI command definitions
'***************************************************************************
Public Const SC_HA_INQUIRY = &H0                  ' Host adapter inquiry
Public Const SC_GET_DEV_TYPE = &H1                ' Get device type
Public Const SC_EXEC_SCSI_CMD = &H2               ' Execute SCSI command
Public Const SC_ABORT_SRB = &H3                   ' Abort an SRB
Public Const SC_RESET_DEV = &H4                   ' SCSI bus device reset
Public Const SC_SET_HA_PARMS = &H5                ' Set HA parameters
Public Const SC_GET_DISK_INFO = &H6               ' Get Disk
Public Const SC_RESCAN_SCSI_BUS = &H7             ' Rebuild SCSI device map
Public Const SC_GETSET_TIMEOUTS = &H8             ' Get/Set target timeouts


'***************************************************************************
'** SRB Status
'***************************************************************************
Public Const SS_PENDING = &H0                     ' SRB being processed
Public Const SS_COMP = &H1                        ' SRB completed without error
Public Const SS_ABORTED = &H2                     ' SRB aborted
Public Const SS_ABORT_FAIL = &H3                  ' Unable to abort SRB
Public Const SS_ERR = &H4                         ' SRB completed with error
Public Const SS_INVALID_CMD = &H80                ' Invalid ASPI command
Public Const SS_INVALID_HA = &H81                 ' Invalid host adapter number
Public Const SS_NO_DEVICE = &H82                  ' SCSI device not installed
Public Const SS_INVALID_SRB = &HE0                ' Invalid parameter set in SRB
Public Const SS_OLD_MANAGER = &HE1                ' ASPI manager doesn't support windows
Public Const SS_BUFFER_ALIGN = &HE1               ' Buffer not aligned (replaces SS_OLD_MANAGER in Win32)
Public Const SS_ILLEGAL_MODE = &HE2               ' Unsupported Windows mode
Public Const SS_NO_ASPI = &HE3                    ' No ASPI managers
Public Const SS_FAILED_INIT = &HE4                ' ASPI for windows failed init
Public Const SS_ASPI_IS_BUSY = &HE5               ' No resources available to execute command
Public Const SS_BUFFER_TO_BIG = &HE6              ' Buffer size too big to handle
Public Const SS_BUFFER_TOO_BIG = &HE6             ' Correct spelling of 'too'
Public Const SS_MISMATCHED_COMPONENTS = &HE7      ' The DLLs/EXEs of ASPI don't version check
Public Const SS_NO_ADAPTERS = &HE8                ' No host adapters to manager
Public Const SS_INSUFFICIENT_RESOURCES = &HE9     ' Couldn't allocate resources needed to init
Public Const SS_ASPI_IS_SHUTDOWN = &HEA           ' Call came to ASPI after PROCESS_DETACH
Public Const SS_BAD_INSTALL = &HEB                ' The DLL or other components are installed wrong

'***************************************************************************
'** Host Adapter Status
'***************************************************************************
Public Const HASTAT_OK = &H0                      ' No error detected by HA
Public Const HASTAT_SEL_TO = &H11                 ' Selection Timeout
Public Const HASTAT_DO_DU = &H12                  ' Data overrun/data underrun
Public Const HASTAT_BUS_FREE = &H13               ' Unexpected bus free
Public Const HASTAT_PHASE_ERR = &H14              ' Target bus phase sequence
Public Const HASTAT_TIMEOUT = &H9                 ' Timed out while SRB was waiting to be processed
Public Const HASTAT_COMMAND_TIMEOUT = &HB         ' Adapter timed out while processing SRB
Public Const HASTAT_MESSAGE_REJECT = &HD          ' While processing the SRB, the adapter received a MESSAGE
Public Const HASTAT_BUS_RESET = &HE               ' A bus reset was detected
Public Const HASTAT_PARITY_ERROR = &HF            ' A parity error was detected
Public Const HASTAT_REQUEST_SENSE_FAILED = &H10   ' The adapter failed in issuing






