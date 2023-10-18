Option Strict Off
Option Explicit On
Module NIGLOBAL
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	' 32-bit Visual Basic Language Interface
	' Version 1.7
	' Copyright 1998 National Instruments Corporation.
	' All Rights Reserved.
	'
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'   This module contains the variable  declarations,
	'   constant definitions, and type information that
	'   is recognized by the entire application.
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	
	Public ibsta As Short
	Public iberr As Short
	Public ibcnt As Short
	Public ibcntl As Integer
	
	' Needed to register for GPIB Global Thread.
	Public Longibsta As Integer
	Public Longiberr As Integer
	Public Longibcnt As Integer
	Public GPIBglobalsRegistered As Short
	
	Public buf As String
	
	Public bytebuf() As Byte
	
	Public Const UNL As Short = &H3Fs ' GPIB unlisten command
	Public Const UNT As Short = &H5Fs ' GPIB untalk command
	Public Const GTL As Short = &H1s ' GPIB go to local
	Public Const SDC As Short = &H4s ' GPIB selected device clear
	Public Const PPC As Short = &H5s ' GPIB parallel poll configure
	Public Const GGET As Short = &H8s ' GPIB group execute trigger
	Public Const TCT As Short = &H9s ' GPIB take control
	Public Const LLO As Short = &H11s ' GPIB local lock out
	Public Const DCL As Short = &H14s ' GPIB device clear
	Public Const PPU As Short = &H15s ' GPIB parallel poll unconfigure
	Public Const SPE As Short = &H18s ' GPIB serial poll enable
	Public Const SPD As Short = &H19s ' GPIB serial poll disable
	Public Const PPE As Short = &H60s ' GPIB parallel poll enable
	Public Const PPD As Short = &H70s ' GPIB parallel poll disable
	
	' GPIB status bit vector :
	'       status variable ibsta and wait mask
	
	Public Const EERR As Short = &H8000s ' Error detected
	Public Const TIMO As Short = &H4000s ' Timeout
	Public Const EEND As Short = &H2000s ' EOI or EOS detected
	Public Const SRQI As Short = &H1000s ' SRQ detected by CIC
	Public Const RQS As Short = &H800s ' Device requesting service
	Public Const CMPL As Short = &H100s ' I/O completed
	Public Const LOK As Short = &H80s ' Local lockout state
	Public Const RREM As Short = &H40s ' Remote state
	Public Const CIC As Short = &H20s ' Controller-in-Charge
	Public Const AATN As Short = &H10s ' Attention asserted
	Public Const TACS As Short = &H8s ' Talker active
	Public Const LACS As Short = &H4s ' Listener active
	Public Const DTAS As Short = &H2s ' Device trigger state
	Public Const DCAS As Short = &H1s ' Device clear state
	
	' Error messages returned in global variable iberr
	
	Public Const EDVR As Short = 0 ' System error
	Public Const ECIC As Short = 1 ' Function requires GPIB board to be CIC
	Public Const ENOL As Short = 2 ' Write function detected no listeners
	Public Const EADR As Short = 3 ' Interface board not addressed correctly
	Public Const EARG As Short = 4 ' Invalid argument to function call
	Public Const ESAC As Short = 5 ' Function requires GPIB board to be SAC
	Public Const EABO As Short = 6 ' I/O operation aborted
	Public Const ENEB As Short = 7 ' Non-existent interface board
	Public Const EDMA As Short = 8 ' DMA Error
	Public Const EOIP As Short = 10 ' I/O operation started before previous
	' operation completed
	Public Const ECAP As Short = 11 ' No capability for intended operation
	Public Const EFSO As Short = 12 ' File system operation error
	Public Const EBUS As Short = 14 ' Command error during device call
	Public Const ESTB As Short = 15 ' Serial poll status byte lost
	Public Const ESRQ As Short = 16 ' SRQ remains asserted
	Public Const ETAB As Short = 20 ' The return buffer is full
	Public Const ELCK As Short = 21 ' Address or board is locked
	
	
	' EOS mode bits
	
	Public Const BIN As Short = &H1000s ' Eight bit compare
	Public Const XEOS As Short = &H800s ' Send EOI with EOS byte
	Public Const REOS As Short = &H400s ' Terminate read on EOS
	
	' Timeout values and meanings
	
	Public Const TNONE As Short = 0 ' Infinite timeout (disabled)
	Public Const T10us As Short = 1 ' Timeout of 10 us (ideal)
	Public Const T30us As Short = 2 ' Timeout of 30 us (ideal)
	Public Const T100us As Short = 3 ' Timeout of 100 us (ideal)
	Public Const T300us As Short = 4 ' Timeout of 300 us (ideal)
	Public Const T1ms As Short = 5 ' Timeout of 1 ms (ideal)
	Public Const T3ms As Short = 6 ' Timeout of 3 ms (ideal)
	Public Const T10ms As Short = 7 ' Timeout of 10 ms (ideal)
	Public Const T30ms As Short = 8 ' Timeout of 30 ms (ideal)
	Public Const T100ms As Short = 9 ' Timeout of 100 ms (ideal)
	Public Const T300ms As Short = 10 ' Timeout of 300 ms (ideal)
	Public Const T1s As Short = 11 ' Timeout of 1 s (ideal)
	Public Const T3s As Short = 12 ' Timeout of 3 s (ideal)
	Public Const T10s As Short = 13 ' Timeout of 10 s (ideal)
	Public Const T30s As Short = 14 ' Timeout of 30 s (ideal)
	Public Const T100s As Short = 15 ' Timeout of 100 s (ideal)
	Public Const T300s As Short = 16 ' Timeout of 300 s (ideal)
	Public Const T1000s As Short = 17 ' Timeout of 1000 s (maximum)
	
	' IBLN constants
	
	Public Const ALL_SAD As Short = -1
	Public Const NO_SAD As Short = 0
	
	' The following constants are used for the second parameter of the
	' ibconfig function.  They are the "option" selection codes.
	
	Public Const IbcPAD As Short = &H1s ' Primary Address
	Public Const IbcSAD As Short = &H2s ' Secondary Address
	Public Const IbcTMO As Short = &H3s ' Timeout Value
	Public Const IbcEOT As Short = &H4s ' Send EOI with last data byte?
	Public Const IbcPPC As Short = &H5s ' Parallel Poll Configure
	Public Const IbcREADDR As Short = &H6s ' Repeat Addressing
	Public Const IbcAUTOPOLL As Short = &H7s ' Disable Auto Serial Polling
	Public Const IbcCICPROT As Short = &H8s ' Use the CIC Protocol?
	Public Const IbcIRQ As Short = &H9s ' Use PIO for I/O
	Public Const IbcSC As Short = &HAs ' Board is System Controller.
	Public Const IbcSRE As Short = &HBs ' Assert SRE on device calls?
	Public Const IbcEOSrd As Short = &HCs ' Terminate reads on EOS.
	Public Const IbcEOSwrt As Short = &HDs ' Send EOI with EOS character.
	Public Const IbcEOScmp As Short = &HEs ' Use 7 or 8-bit EOS compare.
	Public Const IbcEOSchar As Short = &HFs ' The EOS character.
	Public Const IbcPP2 As Short = &H10s ' Use Parallel Poll Mode 2.
	Public Const IbcTIMING As Short = &H11s ' NORMAL, HIGH, or VERY_HIGH timing.
	Public Const IbcDMA As Short = &H12s ' Use DMA for I/O.
	Public Const IbcReadAdjust As Short = &H13s ' Swap bytes during an ibrd.
	Public Const IbcWriteAdjust As Short = &H14s ' Swap bytes during an ibwrt.
	Public Const IbcSendLLO As Short = &H17s ' Enable/disable the sending of LLO.
	Public Const IbcSPollTime As Short = &H18s ' Set the timeout value for serial polls.
	Public Const IbcPPollTime As Short = &H19s ' Set the parallel poll length period
	Public Const IbcEndBitIsNormal As Short = &H1As ' Remove EOS from END bit of IBSTA.
	Public Const IbcUnAddr As Short = &H1Bs ' Enable/disable device unaddressing.
	Public Const IbcSignalNumber As Short = &H1Cs ' Set UNIX signal number - unsupported
	Public Const IbcHSCableLength As Short = &H1Fs ' Enable/disable high-speed handshaking.
	Public Const IbcIst As Short = &H20s ' Set the IST bit
	Public Const IbcRsv As Short = &H21s ' Set the RSV bit
	Public Const IbcLON As Short = &H22s ' Enable listen only mode.
	
	
	'   Constants that can be used (in addition to the ibconfig constants)
	'   when calling the IBASK function.
	
	Public Const IbaPAD As Short = &H1s ' Primary Address
	Public Const IbaSAD As Short = &H2s ' Secondary Address
	Public Const IbaTMO As Short = &H3s ' Timeout Value
	Public Const IbaEOT As Short = &H4s ' Send EOI with last data byte?
	Public Const IbaPPC As Short = &H5s ' Parallel Poll Configure
	Public Const IbaREADDR As Short = &H6s ' Repeat Addressing
	Public Const IbaAUTOPOLL As Short = &H7s ' Disable Auto Serial Polling
	Public Const IbaCICPROT As Short = &H8s ' Use the CIC Protocol?
	Public Const IbaIRQ As Short = &H9s ' Use PIO for I/O
	Public Const IbaSC As Short = &HAs ' Board is System Controller.
	Public Const IbaSRE As Short = &HBs ' Assert SRE on device calls?
	Public Const IbaEOSrd As Short = &HCs ' Terminate reads on EOS.
	Public Const IbaEOSwrt As Short = &HDs ' Send EOI with EOS character.
	Public Const IbaEOScmp As Short = &HEs ' Use 7 or 8-bit EOS compare.
	Public Const IbaEOSchar As Short = &HFs ' The EOS character.
	Public Const IbaPP2 As Short = &H10s ' Use Parallel Poll Mode 2.
	Public Const IbaTIMING As Short = &H11s ' NORMAL, HIGH, or VERY_HIGH timing.
	Public Const IbaDMA As Short = &H12s ' Use DMA for I/O.
	Public Const IbaReadAdjust As Short = &H13s ' Swap bytes during an ibrd.
	Public Const IbaWriteAdjust As Short = &H14s ' Swap bytes during an ibwrt.
	Public Const IbaSendLLO As Short = &H17s ' Enable/disable the sending of LLO.
	Public Const IbaSPollTime As Short = &H18s ' Set the timeout value for serial polls.
	Public Const IbaPPollTime As Short = &H19s ' Set the parallel poll length period
	Public Const IbaEndBitIsNormal As Short = &H1As ' Remove EOS from END bit of IBSTA.
	Public Const IbaUnAddr As Short = &H1Bs ' Enable/disable device unaddressing.
	Public Const IbaSignalNumber As Short = &H1Cs ' Set UNIX signal number - unsupported
	Public Const IbaHSCableLength As Short = &H1Fs ' Enable/disable high-speed handshaking.
	Public Const IbaIst As Short = &H20s ' Set the IST bit
	Public Const IbaRsv As Short = &H21s ' Set the RSV bit
	Public Const IbaBNA As Short = &H200s ' A device's access board.
	Public Const IbaBaseAddr As Short = &H201s ' A GPIB board's base I/O address.
	Public Const IbaDmaChannel As Short = &H202s ' A GPIB board's DMA channel.
	Public Const IbaIrqLevel As Short = &H203s ' A GPIB board's IRQ level.
	Public Const IbaBaud As Short = &H204s ' Baud rate used to communicate to CT box.
	Public Const IbaParity As Short = &H205s ' Parity setting for CT box.
	Public Const IbaStopBits As Short = &H206s ' Stop bits used for communicating to CT.
	Public Const IbaDataBits As Short = &H207s ' Data bits used for communicating to CT.
	Public Const IbaComPort As Short = &H208s ' System COM port used for CT box.
	Public Const IbaComIrqLevel As Short = &H209s ' System COM port's interrupt level.
	Public Const IbaComPortBase As Short = &H20As ' System COM port's base I/O address.
	Public Const IbaSingleCycleDma As Short = &H20Bs ' Does the board use single cycle DMA?
	Public Const IbaSlotNumber As Short = &H20Cs ' Board's slot number.
	Public Const IbaLPTNumber As Short = &H20Ds ' Parallel port number
	Public Const IbaLPTType As Short = &H20Es ' Parallel port protocol
	
	' These are the values used by the 488.2 Send command
	
	Public Const NULLend As Short = &H0s ' Do nothing at the end of a transfer
	Public Const NLend As Short = &H1s ' Send NL with EOI after a transfer
	Public Const DABend As Short = &H2s ' Send EOI with the last DAB
	
	' This value is useds by the 488.2 Receive command
	
	Public Const STOPend As Short = &H100s ' Stop the read on EOI
	
	' The following values are used by the iblines function.  The integer
	' returned by iblines contains:
	'       The lower byte will contain a "monitor" bit mask.  If a bit
	'               is set (1) in this mask, then the corresponding line
	'               can be monitored by the driver.  If the bit is clear (0),
	'               then the line cannot be monitored.
	'       The upper byte will contain the status of the bus lines.
	'               Each bit corresponds to a certain bus line, and has
	'               a corresponding "monitor" bit in the lower byte.
	
    Public Const ValidEOI As Short = &H80S
    Public Const ValidATN As Short = &H40S
	Public Const ValidSRQ As Short = &H20s
	Public Const ValidREN As Short = &H10s
	Public Const ValidIFC As Short = &H8s
	Public Const ValidNRFD As Short = &H4s
	Public Const ValidNDAC As Short = &H2s
	Public Const ValidDAV As Short = &H1s
	Public Const BusEOI As Short = &H8000s
	Public Const BusATN As Short = &H4000s
	Public Const BusSRQ As Short = &H2000s
	Public Const BusREN As Short = &H1000s
	Public Const BusIFC As Short = &H800s
	Public Const BusNRFD As Short = &H400s
	Public Const BusNDAC As Short = &H200s
	Public Const BusDAV As Short = &H100s
	
	' This value is used to terminate an address list.  It should be
	' assigned to the last entry.  (488.2)
	
    Public Const NOADDR As Short = &HFFFFS

	
	' This value is defined for when the GPIBnotify Callback
	' has failed to rearm.
	
	Public Const IBNOTIFY_REARM_FAILED As Integer = &HE00A003F
	
	' These constants are for use with iblockx/ibunlockx
	' functions for GPIB-ENET
	
	Public Const TIMMEDIATE As Short = -1
	Public Const TINFINITE As Short = -2
	Public Const MAX_LOCKSHARENAME_LENGTH As Short = 64
End Module