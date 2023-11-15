'Option Strict Off
'Option Explicit On

Imports VB = Microsoft.VisualBasic

Public Module Unknown


    '=========================================================
    ' -------------------------------------------------------------------------
    '  Distributed by VXIplug&play Systems Alliance
    '  Do not modify the contents of this file.
    ' -------------------------------------------------------------------------
    '  Title   : VISA32.BAS
    '  Date    : 01-14-2003
    '  Purpose : Include file for the VISA Library 3.0 spec
    ' -------------------------------------------------------------------------

    Public Const VI_SPEC_VERSION As Integer = &H300000

    ' - Resource Template Functions and Operations ----------------------------

    Declare Function viOpenDefaultRM Lib "VISA32.DLL" Alias "#141" (ByRef sesn As Integer) As Integer
    Declare Function viGetDefaultRM Lib "VISA32.DLL" Alias "#128" (ByRef sesn As Integer) As Integer
    Declare Function viFindRsrc Lib "VISA32.DLL" Alias "#129" (ByVal sesn As Integer, ByVal expr As String, ByRef vi As Integer, ByRef retCount As Integer, ByVal desc As String) As Integer
    Declare Function viFindNext Lib "VISA32.DLL" Alias "#130" (ByVal vi As Integer, ByVal desc As String) As Integer
    Declare Function viParseRsrc Lib "VISA32.DLL" Alias "#146" (ByVal sesn As Integer, ByVal desc As String, ByRef intfType As Short, ByRef intfNum As Short) As Integer
    Declare Function viParseRsrcEx Lib "VISA32.DLL" Alias "#147" (ByVal sesn As Integer, ByVal desc As String, ByRef intfType As Short, ByRef intfNum As Short, ByVal rsrcClass As String, ByVal expandedUnaliasedName As String, ByVal aliasIfExists As String) As Integer
    Declare Function viOpen Lib "VISA32.DLL" Alias "#131" (ByVal sesn As Integer, ByVal viDesc As String, ByVal mode As Integer, ByVal timeout As Integer, ByRef vi As Integer) As Integer
    Declare Function viClose Lib "VISA32.DLL" Alias "#132" (ByVal vi As Integer) As Integer
    ' Declare Function viGetAttribute Lib "VISA32.DLL" Alias "#133" (ByVal vi As Integer, ByVal attrName As Integer, ByRef attrValue As System.Delegate) As Integer
    Declare Function viGetAttribute Lib "VISA32.DLL" Alias "#133" (ByVal vi As Integer, ByVal attrName As Integer, ByRef attrValue As String) As Integer
    Declare Function viSetAttribute Lib "VISA32.DLL" Alias "#134" (ByVal vi As Integer, ByVal attrName As Integer, ByVal attrValue As Integer) As Integer
    Declare Function viStatusDesc Lib "VISA32.DLL" Alias "#142" (ByVal vi As Integer, ByVal status As Integer, ByVal desc As String) As Integer
    Declare Function viLock Lib "VISA32.DLL" Alias "#144" (ByVal vi As Integer, ByVal lockType As Integer, ByVal timeout As Integer, ByVal requestedKey As String, ByVal accessKey As String) As Integer
    Declare Function viUnlock Lib "VISA32.DLL" Alias "#145" (ByVal vi As Integer) As Integer
    Declare Function viEnableEvent Lib "VISA32.DLL" Alias "#135" (ByVal vi As Integer, ByVal eventType As Integer, ByVal mechanism As Short, ByVal context As Integer) As Integer
    Declare Function viDisableEvent Lib "VISA32.DLL" Alias "#136" (ByVal vi As Integer, ByVal eventType As Integer, ByVal mechanism As Short) As Integer
    Declare Function viDiscardEvents Lib "VISA32.DLL" Alias "#137" (ByVal vi As Integer, ByVal eventType As Integer, ByVal mechanism As Short) As Integer
    Declare Function viWaitOnEvent Lib "VISA32.DLL" Alias "#138" (ByVal vi As Integer, ByVal inEventType As Integer, ByVal timeout As Integer, ByRef outEventType As Integer, ByRef outEventContext As Integer) As Integer

    ' - Basic I/O Operations --------------------------------------------------

    Declare Function viRead Lib "VISA32.DLL" Alias "#256" (ByVal vi As Integer, ByVal Buffer As String, ByVal count As Integer, ByRef retCount As Long) As Integer
    Declare Function viReadToFile Lib "VISA32.DLL" Alias "#219" (ByVal vi As Integer, ByVal filename As String, ByVal count As Integer, ByRef retCount As Long) As Integer
    Declare Function viWrite Lib "VISA32.DLL" Alias "#257" (ByVal vi As Integer, ByVal Buffer As String, ByVal count As Integer, ByRef retCount As Integer) As Integer
    Declare Function viWriteFromFile Lib "VISA32.DLL" Alias "#218" (ByVal vi As Integer, ByVal filename As String, ByVal count As Integer, ByRef retCount As Integer) As Integer
    Declare Function viAssertTrigger Lib "VISA32.DLL" Alias "#258" (ByVal vi As Integer, ByVal protocol As Short) As Integer
    Declare Function viReadSTB Lib "VISA32.DLL" Alias "#259" (ByVal vi As Integer, ByRef status As Short) As Integer
    Declare Function viClear Lib "VISA32.DLL" Alias "#260" (ByVal vi As Integer) As Integer

    ' - Formatted and Buffered I/O Operations ---------------------------------

    Declare Function viSetBuf Lib "VISA32.DLL" Alias "#267" (ByVal vi As Integer, ByVal mask As Short, ByVal bufSize As Integer) As Integer
    Declare Function viFlush Lib "VISA32.DLL" Alias "#268" (ByVal vi As Integer, ByVal mask As Short) As Integer
    Declare Function viBufWrite Lib "VISA32.DLL" Alias "#202" (ByVal vi As Integer, ByVal Buffer As String, ByVal count As Integer, ByRef retCount As Integer) As Integer
    Declare Function viBufRead Lib "VISA32.DLL" Alias "#203" (ByVal vi As Integer, ByVal Buffer As String, ByVal count As Integer, ByRef retCount As Integer) As Integer
    'Declare Function viVPrintf Lib "VISA32.DLL" Alias "#270" (ByVal vi As Integer, ByVal writeFmt As String, ByRef params As System.Delegate) As Integer
    Declare Function viVPrintf Lib "VISA32.DLL" Alias "#270" (ByVal vi As Long, ByVal writeFmt As String, ByVal params As Object) As Long



    Declare Function viVSPrintf Lib "VISA32.DLL" Alias "#205" (ByVal vi As Integer, ByVal Buffer As String, ByVal writeFmt As String, ByRef params As System.Delegate) As Integer
    Declare Function viVScanf Lib "VISA32.DLL" Alias "#272" (ByVal vi As Integer, ByVal readFmt As String, ByRef params As System.Delegate) As Integer
    Declare Function viVSScanf Lib "VISA32.DLL" Alias "#207" (ByVal vi As Integer, ByVal Buffer As String, ByVal readFmt As String, ByRef params As System.Delegate) As Integer
    'Declare Function viVQueryf Lib "VISA32.DLL" Alias "#280" (ByVal vi As Integer, ByVal writeFmt As String, ByVal readFmt As String, ByRef params As System.Delegate) As Integer
    Declare Function viVQueryf Lib "VISA32.DLL" Alias "#280" (ByVal vi As Integer, ByVal writeFmt As String, ByVal readFmt As String, ByRef params As Object) As Long

    ' - Memory I/O Operations -------------------------------------------------

    Declare Function viIn8 Lib "VISA32.DLL" Alias "#273" (ByVal vi As Integer, ByVal accSpace As Short, ByVal offset As Integer, ByRef val8 As Byte) As Integer
    Declare Function viOut8 Lib "VISA32.DLL" Alias "#274" (ByVal vi As Integer, ByVal accSpace As Short, ByVal offset As Integer, ByVal val8 As Byte) As Integer
    Declare Function viIn16 Lib "VISA32.DLL" Alias "#261" (ByVal vi As Integer, ByVal accSpace As Short, ByVal offset As Integer, ByRef val16 As Short) As Integer
    Declare Function viOut16 Lib "VISA32.DLL" Alias "#262" (ByVal vi As Integer, ByVal accSpace As Short, ByVal offset As Integer, ByVal val16 As Short) As Integer
    Declare Function viIn32 Lib "VISA32.DLL" Alias "#281" (ByVal vi As Integer, ByVal accSpace As Short, ByVal offset As Integer, ByRef val32 As Integer) As Integer
    Declare Function viOut32 Lib "VISA32.DLL" Alias "#282" (ByVal vi As Integer, ByVal accSpace As Short, ByVal offset As Integer, ByVal val32 As Integer) As Integer
    Declare Function viMoveIn8 Lib "VISA32.DLL" Alias "#283" (ByVal vi As Integer, ByVal accSpace As Short, ByVal offset As Integer, ByVal length As Integer, ByRef buf8 As Byte) As Integer
    Declare Function viMoveOut8 Lib "VISA32.DLL" Alias "#284" (ByVal vi As Integer, ByVal accSpace As Short, ByVal offset As Integer, ByVal length As Integer, ByRef buf8 As Byte) As Integer
    Declare Function viMoveIn16 Lib "VISA32.DLL" Alias "#285" (ByVal vi As Integer, ByVal accSpace As Short, ByVal offset As Integer, ByVal length As Integer, ByRef buf16 As Short) As Integer
    Declare Function viMoveOut16 Lib "VISA32.DLL" Alias "#286" (ByVal vi As Integer, ByVal accSpace As Short, ByVal offset As Integer, ByVal length As Integer, ByRef buf16 As Short) As Integer
    Declare Function viMoveIn32 Lib "VISA32.DLL" Alias "#287" (ByVal vi As Integer, ByVal accSpace As Short, ByVal offset As Integer, ByVal length As Integer, ByRef buf32 As Integer) As Integer
    Declare Function viMoveOut32 Lib "VISA32.DLL" Alias "#288" (ByVal vi As Integer, ByVal accSpace As Short, ByVal offset As Integer, ByVal length As Integer, ByRef buf32 As Integer) As Integer
    Declare Function viMove Lib "VISA32.DLL" Alias "#200" (ByVal vi As Integer, ByVal srcSpace As Short, ByVal srcOffset As Integer, ByVal srcWidth As Short, ByVal destSpace As Short, ByVal destOffset As Integer, ByVal destWidth As Short, ByVal srcLength As Integer) As Integer
    Declare Function viMapAddress Lib "VISA32.DLL" Alias "#263" (ByVal vi As Integer, ByVal mapSpace As Short, ByVal mapOffset As Integer, ByVal mapSize As Integer, ByVal accMode As Short, ByVal suggested As Integer, ByRef address As Integer) As Integer
    Declare Function viUnmapAddress Lib "VISA32.DLL" Alias "#264" (ByVal vi As Integer) As Integer
    Declare Sub viPeek8 Lib "VISA32.DLL" Alias "#275" (ByVal vi As Integer, ByVal address As Integer, ByRef val8 As Byte)
    Declare Sub viPoke8 Lib "VISA32.DLL" Alias "#276" (ByVal vi As Integer, ByVal address As Integer, ByVal val8 As Byte)
    Declare Sub viPeek16 Lib "VISA32.DLL" Alias "#265" (ByVal vi As Integer, ByVal address As Integer, ByRef val16 As Short)
    Declare Sub viPoke16 Lib "VISA32.DLL" Alias "#266" (ByVal vi As Integer, ByVal address As Integer, ByVal val16 As Short)
    Declare Sub viPeek32 Lib "VISA32.DLL" Alias "#289" (ByVal vi As Integer, ByVal address As Integer, ByRef val32 As Integer)
    Declare Sub viPoke32 Lib "VISA32.DLL" Alias "#290" (ByVal vi As Integer, ByVal address As Integer, ByVal val32 As Integer)


    ' - Shared Memory Operations ----------------------------------------------

    Declare Function viMemAlloc Lib "VISA32.DLL" Alias "#291" (ByVal vi As Integer, ByVal memSize As Integer, ByRef offset As Integer) As Integer
    Declare Function viMemFree Lib "VISA32.DLL" Alias "#292" (ByVal vi As Integer, ByVal offset As Integer) As Integer

    ' - Interface Specific Operations -----------------------------------------

    Declare Function viGpibControlREN Lib "VISA32.DLL" Alias "#208" (ByVal vi As Integer, ByVal mode As Short) As Integer
    Declare Function viGpibControlATN Lib "VISA32.DLL" Alias "#210" (ByVal vi As Integer, ByVal mode As Short) As Integer
    Declare Function viGpibSendIFC Lib "VISA32.DLL" Alias "#211" (ByVal vi As Integer) As Integer
    Declare Function viGpibCommand Lib "VISA32.DLL" Alias "#212" (ByVal vi As Integer, ByVal Buffer As String, ByVal count As Integer, ByRef retCount As Integer) As Integer
    Declare Function viGpibPassControl Lib "VISA32.DLL" Alias "#213" (ByVal vi As Integer, ByVal primAddr As Short, ByVal secAddr As Short) As Integer
    Declare Function viVxiCommandQuery Lib "VISA32.DLL" Alias "#209" (ByVal vi As Integer, ByVal mode As Short, ByVal devCmd As Integer, ByRef devResponse As Integer) As Integer
    Declare Function viAssertUtilSignal Lib "VISA32.DLL" Alias "#214" (ByVal vi As Integer, ByVal line As Short) As Integer
    Declare Function viAssertIntrSignal Lib "VISA32.DLL" Alias "#215" (ByVal vi As Integer, ByVal mode As Short, ByVal statusID As Integer) As Integer
    Declare Function viMapTrigger Lib "VISA32.DLL" Alias "#216" (ByVal vi As Integer, ByVal trigSrc As Short, ByVal trigDest As Short, ByVal mode As Short) As Integer
    Declare Function viUnmapTrigger Lib "VISA32.DLL" Alias "#217" (ByVal vi As Integer, ByVal trigSrc As Short, ByVal trigDest As Short) As Integer
    Declare Function viUsbControlOut Lib "VISA32.DLL" Alias "#293" (ByVal vi As Integer, ByVal bmRequestType As Short, ByVal bRequest As Short, ByVal wValue As Short, ByVal wIndex As Short, ByVal wLength As Short, ByRef buf As Byte) As Integer
    Declare Function viUsbControlIn Lib "VISA32.DLL" Alias "#294" (ByVal vi As Integer, ByVal bmRequestType As Short, ByVal bRequest As Short, ByVal wValue As Short, ByVal wIndex As Short, ByVal wLength As Short, ByRef buf As Byte, ByRef retCnt As Short) As Integer

    ' - Attributes ------------------------------------------------------------

    Public Const VI_ATTR_RSRC_CLASS As Integer = &HBFFF0001
    Public Const VI_ATTR_RSRC_NAME As Integer = &HBFFF0002
    Public Const VI_ATTR_RSRC_IMPL_VERSION As Integer = &H3FFF0003
    Public Const VI_ATTR_RSRC_LOCK_STATE As Integer = &H3FFF0004
    Public Const VI_ATTR_MAX_QUEUE_LENGTH As Integer = &H3FFF0005
    Public Const VI_ATTR_USER_DATA As Integer = &H3FFF0007
    Public Const VI_ATTR_FDC_CHNL As Integer = &H3FFF000D
    Public Const VI_ATTR_FDC_MODE As Integer = &H3FFF000F
    Public Const VI_ATTR_FDC_GEN_SIGNAL_EN As Integer = &H3FFF0011
    Public Const VI_ATTR_FDC_USE_PAIR As Integer = &H3FFF0013
    Public Const VI_ATTR_SEND_END_EN As Integer = &H3FFF0016
    Public Const VI_ATTR_TERMCHAR As Integer = &H3FFF0018
    Public Const VI_ATTR_TMO_VALUE As Integer = &H3FFF001A
    Public Const VI_ATTR_GPIB_READDR_EN As Integer = &H3FFF001B
    Public Const VI_ATTR_IO_PROT As Integer = &H3FFF001C
    Public Const VI_ATTR_DMA_ALLOW_EN As Integer = &H3FFF001E
    Public Const VI_ATTR_ASRL_BAUD As Integer = &H3FFF0021
    Public Const VI_ATTR_ASRL_DATA_BITS As Integer = &H3FFF0022
    Public Const VI_ATTR_ASRL_PARITY As Integer = &H3FFF0023
    Public Const VI_ATTR_ASRL_STOP_BITS As Integer = &H3FFF0024
    Public Const VI_ATTR_ASRL_FLOW_CNTRL As Integer = &H3FFF0025
    Public Const VI_ATTR_RD_BUF_OPER_MODE As Integer = &H3FFF002A
    Public Const VI_ATTR_RD_BUF_SIZE As Integer = &H3FFF002B
    Public Const VI_ATTR_WR_BUF_OPER_MODE As Integer = &H3FFF002D
    Public Const VI_ATTR_WR_BUF_SIZE As Integer = &H3FFF002E
    Public Const VI_ATTR_SUPPRESS_END_EN As Integer = &H3FFF0036
    Public Const VI_ATTR_TERMCHAR_EN As Integer = &H3FFF0038
    Public Const VI_ATTR_DEST_ACCESS_PRIV As Integer = &H3FFF0039
    Public Const VI_ATTR_DEST_BYTE_ORDER As Integer = &H3FFF003A
    Public Const VI_ATTR_SRC_ACCESS_PRIV As Integer = &H3FFF003C
    Public Const VI_ATTR_SRC_BYTE_ORDER As Integer = &H3FFF003D
    Public Const VI_ATTR_SRC_INCREMENT As Integer = &H3FFF0040
    Public Const VI_ATTR_DEST_INCREMENT As Integer = &H3FFF0041
    Public Const VI_ATTR_WIN_ACCESS_PRIV As Integer = &H3FFF0045
    Public Const VI_ATTR_WIN_BYTE_ORDER As Integer = &H3FFF0047
    Public Const VI_ATTR_GPIB_ATN_STATE As Integer = &H3FFF0057
    Public Const VI_ATTR_GPIB_ADDR_STATE As Integer = &H3FFF005C
    Public Const VI_ATTR_GPIB_CIC_STATE As Integer = &H3FFF005E
    Public Const VI_ATTR_GPIB_NDAC_STATE As Integer = &H3FFF0062
    Public Const VI_ATTR_GPIB_SRQ_STATE As Integer = &H3FFF0067
    Public Const VI_ATTR_GPIB_SYS_CNTRL_STATE As Integer = &H3FFF0068
    Public Const VI_ATTR_GPIB_HS488_CBL_LEN As Integer = &H3FFF0069
    Public Const VI_ATTR_CMDR_LA As Integer = &H3FFF006B
    Public Const VI_ATTR_VXI_DEV_CLASS As Integer = &H3FFF006C
    Public Const VI_ATTR_MAINFRAME_LA As Integer = &H3FFF0070
    Public Const VI_ATTR_MANF_NAME As Integer = &HBFFF0072
    Public Const VI_ATTR_MODEL_NAME As Integer = &HBFFF0077
    Public Const VI_ATTR_VXI_VME_INTR_STATUS As Integer = &H3FFF008B
    Public Const VI_ATTR_VXI_TRIG_STATUS As Integer = &H3FFF008D
    Public Const VI_ATTR_VXI_VME_SYSFAIL_STATE As Integer = &H3FFF0094
    Public Const VI_ATTR_WIN_BASE_ADDR As Integer = &H3FFF0098
    Public Const VI_ATTR_WIN_SIZE As Integer = &H3FFF009A
    Public Const VI_ATTR_ASRL_AVAIL_NUM As Integer = &H3FFF00AC
    Public Const VI_ATTR_MEM_BASE As Integer = &H3FFF00AD
    Public Const VI_ATTR_ASRL_CTS_STATE As Integer = &H3FFF00AE
    Public Const VI_ATTR_ASRL_DCD_STATE As Integer = &H3FFF00AF
    Public Const VI_ATTR_ASRL_DSR_STATE As Integer = &H3FFF00B1
    Public Const VI_ATTR_ASRL_DTR_STATE As Integer = &H3FFF00B2
    Public Const VI_ATTR_ASRL_END_IN As Integer = &H3FFF00B3
    Public Const VI_ATTR_ASRL_END_OUT As Integer = &H3FFF00B4
    Public Const VI_ATTR_ASRL_REPLACE_CHAR As Integer = &H3FFF00BE
    Public Const VI_ATTR_ASRL_RI_STATE As Integer = &H3FFF00BF
    Public Const VI_ATTR_ASRL_RTS_STATE As Integer = &H3FFF00C0
    Public Const VI_ATTR_ASRL_XON_CHAR As Integer = &H3FFF00C1
    Public Const VI_ATTR_ASRL_XOFF_CHAR As Integer = &H3FFF00C2
    Public Const VI_ATTR_WIN_ACCESS As Integer = &H3FFF00C3
    Public Const VI_ATTR_RM_SESSION As Integer = &H3FFF00C4
    Public Const VI_ATTR_VXI_LA As Integer = &H3FFF00D5
    Public Const VI_ATTR_MANF_ID As Integer = &H3FFF00D9
    Public Const VI_ATTR_MEM_SIZE As Integer = &H3FFF00DD
    Public Const VI_ATTR_MEM_SPACE As Integer = &H3FFF00DE
    Public Const VI_ATTR_MODEL_CODE As Integer = &H3FFF00DF
    Public Const VI_ATTR_SLOT As Integer = &H3FFF00E8
    Public Const VI_ATTR_INTF_INST_NAME As Integer = &HBFFF00E9
    Public Const VI_ATTR_IMMEDIATE_SERV As Integer = &H3FFF0100
    Public Const VI_ATTR_INTF_PARENT_NUM As Integer = &H3FFF0101
    Public Const VI_ATTR_RSRC_SPEC_VERSION As Integer = &H3FFF0170
    Public Const VI_ATTR_INTF_TYPE As Integer = &H3FFF0171
    Public Const VI_ATTR_GPIB_PRIMARY_ADDR As Integer = &H3FFF0172
    Public Const VI_ATTR_GPIB_SECONDARY_ADDR As Integer = &H3FFF0173
    Public Const VI_ATTR_RSRC_MANF_NAME As Integer = &HBFFF0174
    Public Const VI_ATTR_RSRC_MANF_ID As Integer = &H3FFF0175
    Public Const VI_ATTR_INTF_NUM As Integer = &H3FFF0176
    Public Const VI_ATTR_TRIG_ID As Integer = &H3FFF0177
    Public Const VI_ATTR_GPIB_REN_STATE As Integer = &H3FFF0181
    Public Const VI_ATTR_GPIB_UNADDR_EN As Integer = &H3FFF0184
    Public Const VI_ATTR_DEV_STATUS_BYTE As Integer = &H3FFF0189
    Public Const VI_ATTR_FILE_APPEND_EN As Integer = &H3FFF0192
    Public Const VI_ATTR_VXI_TRIG_SUPPORT As Integer = &H3FFF0194
    Public Const VI_ATTR_TCPIP_ADDR As Integer = &HBFFF0195
    Public Const VI_ATTR_TCPIP_HOSTNAME As Integer = &HBFFF0196
    Public Const VI_ATTR_TCPIP_PORT As Integer = &H3FFF0197
    Public Const VI_ATTR_TCPIP_DEVICE_NAME As Integer = &HBFFF0199
    Public Const VI_ATTR_TCPIP_NODELAY As Integer = &H3FFF019A
    Public Const VI_ATTR_TCPIP_KEEPALIVE As Integer = &H3FFF019B
    Public Const VI_ATTR_4882_COMPLIANT As Integer = &H3FFF019F
    Public Const VI_ATTR_USB_SERIAL_NUM As Integer = &HBFFF01A0
    Public Const VI_ATTR_USB_INTFC_NUM As Integer = &H3FFF01A1
    Public Const VI_ATTR_USB_PROTOCOL As Integer = &H3FFF01A7
    Public Const VI_ATTR_USB_MAX_INTR_SIZE As Integer = &H3FFF01AF

    Public Const VI_ATTR_JOB_ID As Integer = &H3FFF4006
    Public Const VI_ATTR_EVENT_TYPE As Integer = &H3FFF4010
    Public Const VI_ATTR_SIGP_STATUS_ID As Integer = &H3FFF4011
    Public Const VI_ATTR_RECV_TRIG_ID As Integer = &H3FFF4012
    Public Const VI_ATTR_INTR_STATUS_ID As Integer = &H3FFF4023
    Public Const VI_ATTR_STATUS As Integer = &H3FFF4025
    Public Const VI_ATTR_RET_COUNT As Integer = &H3FFF4026
    Public Const VI_ATTR_BUFFER As Integer = &H3FFF4027
    Public Const VI_ATTR_RECV_INTR_LEVEL As Integer = &H3FFF4041
    Public Const VI_ATTR_OPER_NAME As Integer = &HBFFF4042
    Public Const VI_ATTR_GPIB_RECV_CIC_STATE As Integer = &H3FFF4193
    Public Const VI_ATTR_RECV_TCPIP_ADDR As Integer = &HBFFF4198
    Public Const VI_ATTR_USB_RECV_INTR_SIZE As Integer = &H3FFF41B0
    Public Const VI_ATTR_USB_RECV_INTR_DATA As Integer = &HBFFF41B1

    ' - Event Types -----------------------------------------------------------

    Public Const VI_EVENT_IO_COMPLETION As Integer = &H3FFF2009
    Public Const VI_EVENT_TRIG As Integer = &HBFFF200A
    Public Const VI_EVENT_SERVICE_REQ As Integer = &H3FFF200B
    Public Const VI_EVENT_CLEAR As Integer = &H3FFF200D
    Public Const VI_EVENT_EXCEPTION As Integer = &HBFFF200E
    Public Const VI_EVENT_GPIB_CIC As Integer = &H3FFF2012
    Public Const VI_EVENT_GPIB_TALK As Integer = &H3FFF2013
    Public Const VI_EVENT_GPIB_LISTEN As Integer = &H3FFF2014
    Public Const VI_EVENT_VXI_VME_SYSFAIL As Integer = &H3FFF201D
    Public Const VI_EVENT_VXI_VME_SYSRESET As Integer = &H3FFF201E
    Public Const VI_EVENT_VXI_SIGP As Integer = &H3FFF2020
    Public Const VI_EVENT_VXI_VME_INTR As Integer = &HBFFF2021
    Public Const VI_EVENT_TCPIP_CONNECT As Integer = &H3FFF2036
    Public Const VI_EVENT_USB_INTR As Integer = &H3FFF2037

    Public Const VI_ALL_ENABLED_EVENTS As Integer = &H3FFF7FFF


    ' - Completion and Error Codes --------------------------------------------

    Public Const VI_SUCCESS As Short = &H0
    Public Const VI_SUCCESS_EVENT_EN As Integer = &H3FFF0002
    Public Const VI_SUCCESS_EVENT_DIS As Integer = &H3FFF0003
    Public Const VI_SUCCESS_QUEUE_EMPTY As Integer = &H3FFF0004
    Public Const VI_SUCCESS_TERM_CHAR As Integer = &H3FFF0005
    Public Const VI_SUCCESS_MAX_CNT As Integer = &H3FFF0006
    Public Const VI_SUCCESS_DEV_NPRESENT As Integer = &H3FFF007D
    Public Const VI_SUCCESS_TRIG_MAPPED As Integer = &H3FFF007E
    Public Const VI_SUCCESS_QUEUE_NEMPTY As Integer = &H3FFF0080
    Public Const VI_SUCCESS_NCHAIN As Integer = &H3FFF0098
    Public Const VI_SUCCESS_NESTED_SHARED As Integer = &H3FFF0099
    Public Const VI_SUCCESS_NESTED_EXCLUSIVE As Integer = &H3FFF009A
    Public Const VI_SUCCESS_SYNC As Integer = &H3FFF009B

    Public Const VI_WARN_QUEUE_OVERFLOW As Integer = &H3FFF000C
    Public Const VI_WARN_CONFIG_NLOADED As Integer = &H3FFF0077
    Public Const VI_WARN_NULL_OBJECT As Integer = &H3FFF0082
    Public Const VI_WARN_NSUP_ATTR_STATE As Integer = &H3FFF0084
    Public Const VI_WARN_UNKNOWN_STATUS As Integer = &H3FFF0085
    Public Const VI_WARN_NSUP_BUF As Integer = &H3FFF0088
    Public Const VI_WARN_EXT_FUNC_NIMPL As Integer = &H3FFF00A9

    Public Const VI_ERROR_SYSTEM_ERROR As Integer = &HBFFF0000
    Public Const VI_ERROR_INV_OBJECT As Integer = &HBFFF000E
    Public Const VI_ERROR_RSRC_LOCKED As Integer = &HBFFF000F
    Public Const VI_ERROR_INV_EXPR As Integer = &HBFFF0010
    Public Const VI_ERROR_RSRC_NFOUND As Integer = &HBFFF0011
    Public Const VI_ERROR_INV_RSRC_NAME As Integer = &HBFFF0012
    Public Const VI_ERROR_INV_ACC_MODE As Integer = &HBFFF0013
    Public Const VI_ERROR_TMO As Integer = &HBFFF0015
    Public Const VI_ERROR_CLOSING_FAILED As Integer = &HBFFF0016
    Public Const VI_ERROR_INV_DEGREE As Integer = &HBFFF001B
    Public Const VI_ERROR_INV_JOB_ID As Integer = &HBFFF001C
    Public Const VI_ERROR_NSUP_ATTR As Integer = &HBFFF001D
    Public Const VI_ERROR_NSUP_ATTR_STATE As Integer = &HBFFF001E
    Public Const VI_ERROR_ATTR_READONLY As Integer = &HBFFF001F
    Public Const VI_ERROR_INV_LOCK_TYPE As Integer = &HBFFF0020
    Public Const VI_ERROR_INV_ACCESS_KEY As Integer = &HBFFF0021
    Public Const VI_ERROR_INV_EVENT As Integer = &HBFFF0026
    Public Const VI_ERROR_INV_MECH As Integer = &HBFFF0027
    Public Const VI_ERROR_HNDLR_NINSTALLED As Integer = &HBFFF0028
    Public Const VI_ERROR_INV_HNDLR_REF As Integer = &HBFFF0029
    Public Const VI_ERROR_INV_CONTEXT As Integer = &HBFFF002A
    Public Const VI_ERROR_NENABLED As Integer = &HBFFF002F
    Public Const VI_ERROR_ABORT As Integer = &HBFFF0030
    Public Const VI_ERROR_RAW_WR_PROT_VIOL As Integer = &HBFFF0034
    Public Const VI_ERROR_RAW_RD_PROT_VIOL As Integer = &HBFFF0035
    Public Const VI_ERROR_OUTP_PROT_VIOL As Integer = &HBFFF0036
    Public Const VI_ERROR_INP_PROT_VIOL As Integer = &HBFFF0037
    Public Const VI_ERROR_BERR As Integer = &HBFFF0038
    Public Const VI_ERROR_IN_PROGRESS As Integer = &HBFFF0039
    Public Const VI_ERROR_INV_SETUP As Integer = &HBFFF003A
    Public Const VI_ERROR_QUEUE_ERROR As Integer = &HBFFF003B
    Public Const VI_ERROR_ALLOC As Integer = &HBFFF003C
    Public Const VI_ERROR_INV_MASK As Integer = &HBFFF003D
    Public Const VI_ERROR_IO As Integer = &HBFFF003E
    Public Const VI_ERROR_INV_FMT As Integer = &HBFFF003F
    Public Const VI_ERROR_NSUP_FMT As Integer = &HBFFF0041
    Public Const VI_ERROR_LINE_IN_USE As Integer = &HBFFF0042
    Public Const VI_ERROR_NSUP_MODE As Integer = &HBFFF0046
    Public Const VI_ERROR_SRQ_NOCCURRED As Integer = &HBFFF004A
    Public Const VI_ERROR_INV_SPACE As Integer = &HBFFF004E
    Public Const VI_ERROR_INV_OFFSET As Integer = &HBFFF0051
    Public Const VI_ERROR_INV_WIDTH As Integer = &HBFFF0052
    Public Const VI_ERROR_NSUP_OFFSET As Integer = &HBFFF0054
    Public Const VI_ERROR_NSUP_VAR_WIDTH As Integer = &HBFFF0055
    Public Const VI_ERROR_WINDOW_NMAPPED As Integer = &HBFFF0057
    Public Const VI_ERROR_RESP_PENDING As Integer = &HBFFF0059
    Public Const VI_ERROR_NLISTENERS As Integer = &HBFFF005F
    Public Const VI_ERROR_NCIC As Integer = &HBFFF0060
    Public Const VI_ERROR_NSYS_CNTLR As Integer = &HBFFF0061
    Public Const VI_ERROR_NSUP_OPER As Integer = &HBFFF0067
    Public Const VI_ERROR_INTR_PENDING As Integer = &HBFFF0068
    Public Const VI_ERROR_ASRL_PARITY As Integer = &HBFFF006A
    Public Const VI_ERROR_ASRL_FRAMING As Integer = &HBFFF006B
    Public Const VI_ERROR_ASRL_OVERRUN As Integer = &HBFFF006C
    Public Const VI_ERROR_TRIG_NMAPPED As Integer = &HBFFF006E
    Public Const VI_ERROR_NSUP_ALIGN_OFFSET As Integer = &HBFFF0070
    Public Const VI_ERROR_USER_BUF As Integer = &HBFFF0071
    Public Const VI_ERROR_RSRC_BUSY As Integer = &HBFFF0072
    Public Const VI_ERROR_NSUP_WIDTH As Integer = &HBFFF0076
    Public Const VI_ERROR_INV_PARAMETER As Integer = &HBFFF0078
    Public Const VI_ERROR_INV_PROT As Integer = &HBFFF0079
    Public Const VI_ERROR_INV_SIZE As Integer = &HBFFF007B
    Public Const VI_ERROR_WINDOW_MAPPED As Integer = &HBFFF0080
    Public Const VI_ERROR_NIMPL_OPER As Integer = &HBFFF0081
    Public Const VI_ERROR_INV_LENGTH As Integer = &HBFFF0083
    Public Const VI_ERROR_INV_MODE As Integer = &HBFFF0091
    Public Const VI_ERROR_SESN_NLOCKED As Integer = &HBFFF009C
    Public Const VI_ERROR_MEM_NSHARED As Integer = &HBFFF009D
    Public Const VI_ERROR_LIBRARY_NFOUND As Integer = &HBFFF009E
    Public Const VI_ERROR_NSUP_INTR As Integer = &HBFFF009F
    Public Const VI_ERROR_INV_LINE As Integer = &HBFFF00A0
    Public Const VI_ERROR_FILE_ACCESS As Integer = &HBFFF00A1
    Public Const VI_ERROR_FILE_IO As Integer = &HBFFF00A2
    Public Const VI_ERROR_NSUP_LINE As Integer = &HBFFF00A3
    Public Const VI_ERROR_NSUP_MECH As Integer = &HBFFF00A4
    Public Const VI_ERROR_INTF_NUM_NCONFIG As Integer = &HBFFF00A5
    Public Const VI_ERROR_CONN_LOST As Integer = &HBFFF00A6

    ' - Other VISA Definitions ------------------------------------------------

    Public Const VI_FIND_BUFLEN As Short = 256

    Public Const VI_NULL As Short = 0
    Public Const VI_TRUE As Short = 1
    Public Const VI_FALSE As Short = 0

    Public Const VI_INTF_GPIB As Short = 1
    Public Const VI_INTF_VXI As Short = 2
    Public Const VI_INTF_GPIB_VXI As Short = 3
    Public Const VI_INTF_ASRL As Short = 4
    Public Const VI_INTF_TCPIP As Short = 6
    Public Const VI_INTF_USB As Short = 7

    Public Const VI_PROT_NORMAL As Short = 1
    Public Const VI_PROT_FDC As Short = 2
    Public Const VI_PROT_HS488 As Short = 3
    Public Const VI_PROT_4882_STRS As Short = 4
    Public Const VI_PROT_USBTMC_VENDOR As Short = 5

    Public Const VI_FDC_NORMAL As Short = 1
    Public Const VI_FDC_STREAM As Short = 2

    Public Const VI_LOCAL_SPACE As Short = 0
    Public Const VI_A16_SPACE As Short = 1
    Public Const VI_A24_SPACE As Short = 2
    Public Const VI_A32_SPACE As Short = 3
    Public Const VI_OPAQUE_SPACE As Integer = &HFFFF

    Public Const VI_UNKNOWN_LA As Short = -1
    Public Const VI_UNKNOWN_SLOT As Short = -1
    Public Const VI_UNKNOWN_LEVEL As Short = -1

    Public Const VI_QUEUE As Short = 1
    Public Const VI_ALL_MECH As Integer = &HFFFF

    Public Const VI_TRIG_ALL As Short = -2
    Public Const VI_TRIG_SW As Short = -1
    Public Const VI_TRIG_TTL0 As Short = 0
    Public Const VI_TRIG_TTL1 As Short = 1
    Public Const VI_TRIG_TTL2 As Short = 2
    Public Const VI_TRIG_TTL3 As Short = 3
    Public Const VI_TRIG_TTL4 As Short = 4
    Public Const VI_TRIG_TTL5 As Short = 5
    Public Const VI_TRIG_TTL6 As Short = 6
    Public Const VI_TRIG_TTL7 As Short = 7
    Public Const VI_TRIG_ECL0 As Short = 8
    Public Const VI_TRIG_ECL1 As Short = 9
    Public Const VI_TRIG_PANEL_IN As Short = 27
    Public Const VI_TRIG_PANEL_OUT As Short = 28

    Public Const VI_TRIG_PROT_DEFAULT As Short = 0
    Public Const VI_TRIG_PROT_ON As Short = 1
    Public Const VI_TRIG_PROT_OFF As Short = 2
    Public Const VI_TRIG_PROT_SYNC As Short = 5

    Public Const VI_READ_BUF As Short = 1
    Public Const VI_WRITE_BUF As Short = 2
    Public Const VI_READ_BUF_DISCARD As Short = 4
    Public Const VI_WRITE_BUF_DISCARD As Short = 8
    Public Const VI_IO_IN_BUF As Short = 16
    Public Const VI_IO_OUT_BUF As Short = 32
    Public Const VI_IO_IN_BUF_DISCARD As Short = 64
    Public Const VI_IO_OUT_BUF_DISCARD As Short = 128

    Public Const VI_FLUSH_ON_ACCESS As Short = 1
    Public Const VI_FLUSH_WHEN_FULL As Short = 2
    Public Const VI_FLUSH_DISABLE As Short = 3

    Public Const VI_NMAPPED As Short = 1
    Public Const VI_USE_OPERS As Short = 2
    Public Const VI_DEREF_ADDR As Short = 3

    Public Const VI_TMO_IMMEDIATE As Short = &H0
    Public Const VI_TMO_INFINITE As Short = &HFFFFFFFF

    'accessMode
    Public Const VI_NO_LOCK As Short = 0
    Public Const VI_EXCLUSIVE_LOCK As Short = 1
    Public Const VI_SHARED_LOCK As Short = 2
    Public Const VI_LOAD_CONFIG As Short = 4

    Public Const VI_NO_SEC_ADDR As Integer = &HFFFF

    Public Const VI_ASRL_PAR_NONE As Short = 0
    Public Const VI_ASRL_PAR_ODD As Short = 1
    Public Const VI_ASRL_PAR_EVEN As Short = 2
    Public Const VI_ASRL_PAR_MARK As Short = 3
    Public Const VI_ASRL_PAR_SPACE As Short = 4

    Public Const VI_ASRL_STOP_ONE As Short = 10
    Public Const VI_ASRL_STOP_ONE5 As Short = 15
    Public Const VI_ASRL_STOP_TWO As Short = 20

    Public Const VI_ASRL_FLOW_NONE As Short = 0
    Public Const VI_ASRL_FLOW_XON_XOFF As Short = 1
    Public Const VI_ASRL_FLOW_RTS_CTS As Short = 2
    Public Const VI_ASRL_FLOW_DTR_DSR As Short = 4

    Public Const VI_ASRL_END_NONE As Short = 0
    Public Const VI_ASRL_END_LAST_BIT As Short = 1
    Public Const VI_ASRL_END_TERMCHAR As Short = 2
    Public Const VI_ASRL_END_BREAK As Short = 3

    Public Const VI_STATE_ASSERTED As Short = 1
    Public Const VI_STATE_UNASSERTED As Short = 0
    Public Const VI_STATE_UNKNOWN As Short = -1

    Public Const VI_BIG_ENDIAN As Short = 0
    Public Const VI_LITTLE_ENDIAN As Short = 1

    Public Const VI_DATA_PRIV As Short = 0
    Public Const VI_DATA_NPRIV As Short = 1
    Public Const VI_PROG_PRIV As Short = 2
    Public Const VI_PROG_NPRIV As Short = 3
    Public Const VI_BLCK_PRIV As Short = 4
    Public Const VI_BLCK_NPRIV As Short = 5
    Public Const VI_D64_PRIV As Short = 6
    Public Const VI_D64_NPRIV As Short = 7

    Public Const VI_WIDTH_8 As Short = 1
    Public Const VI_WIDTH_16 As Short = 2
    Public Const VI_WIDTH_32 As Short = 4

    Public Const VI_GPIB_REN_DEASSERT As Short = 0
    Public Const VI_GPIB_REN_ASSERT As Short = 1
    Public Const VI_GPIB_REN_DEASSERT_GTL As Short = 2
    Public Const VI_GPIB_REN_ASSERT_ADDRESS As Short = 3
    Public Const VI_GPIB_REN_ASSERT_LLO As Short = 4
    Public Const VI_GPIB_REN_ASSERT_ADDRESS_LLO As Short = 5
    Public Const VI_GPIB_REN_ADDRESS_GTL As Short = 6

    Public Const VI_GPIB_ATN_DEASSERT As Short = 0
    Public Const VI_GPIB_ATN_ASSERT As Short = 1
    Public Const VI_GPIB_ATN_DEASSERT_HANDSHAKE As Short = 2
    Public Const VI_GPIB_ATN_ASSERT_IMMEDIATE As Short = 3

    Public Const VI_GPIB_HS488_DISABLED As Short = 0
    Public Const VI_GPIB_HS488_NIMPL As Short = -1

    Public Const VI_GPIB_UNADDRESSED As Short = 0
    Public Const VI_GPIB_TALKER As Short = 1
    Public Const VI_GPIB_LISTENER As Short = 2

    Public Const VI_VXI_CMD16 As Short = &H0200
    Public Const VI_VXI_CMD16_RESP16 As Short = &H0202
    Public Const VI_VXI_RESP16 As Short = &H0002
    Public Const VI_VXI_CMD32 As Short = &H0400
    Public Const VI_VXI_CMD32_RESP16 As Short = &H0402
    Public Const VI_VXI_CMD32_RESP32 As Short = &H0404
    Public Const VI_VXI_RESP32 As Short = &H0004

    Public Const VI_ASSERT_SIGNAL As Short = -1
    Public Const VI_ASSERT_USE_ASSIGNED As Short = 0
    Public Const VI_ASSERT_IRQ1 As Short = 1
    Public Const VI_ASSERT_IRQ2 As Short = 2
    Public Const VI_ASSERT_IRQ3 As Short = 3
    Public Const VI_ASSERT_IRQ4 As Short = 4
    Public Const VI_ASSERT_IRQ5 As Short = 5
    Public Const VI_ASSERT_IRQ6 As Short = 6
    Public Const VI_ASSERT_IRQ7 As Short = 7

    Public Const VI_UTIL_ASSERT_SYSRESET As Short = 1
    Public Const VI_UTIL_ASSERT_SYSFAIL As Short = 2
    Public Const VI_UTIL_DEASSERT_SYSFAIL As Short = 3

    Public Const VI_VXI_CLASS_MEMORY As Short = 0
    Public Const VI_VXI_CLASS_EXTENDED As Short = 1
    Public Const VI_VXI_CLASS_MESSAGE As Short = 2
    Public Const VI_VXI_CLASS_REGISTER As Short = 3
    Public Const VI_VXI_CLASS_OTHER As Short = 4

    ' - Backward Compatibility Macros -----------------------------------------

    Public Const VI_ERROR_INV_SESSION As Integer = &HBFFF000E
    Public Const VI_INFINITE As Short = &HFFFFFFFF

    Public Const VI_NORMAL As Short = 1
    Public Const VI_FDC As Short = 2
    Public Const VI_HS488 As Short = 3
    Public Const VI_ASRL488 As Short = 4

    Public Const VI_ASRL_IN_BUF As Short = 16
    Public Const VI_ASRL_OUT_BUF As Short = 32
    Public Const VI_ASRL_IN_BUF_DISCARD As Short = 64
    Public Const VI_ASRL_OUT_BUF_DISCARD As Short = 128


End Module