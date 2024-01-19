Imports System.Runtime.InteropServices
'==== Author: Hungdar Wang ====
'==== Version:  2014.10.28 ====




Module RTBBLibDll

#If DEBUG Then
    Const RTBBDLLNAME As String = "RTBBLib.dll"
#Else
    Const RTBBDLLNAME As String = "RTBBLib.dll"
#End If

    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_EnumBoard",
        CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_EnumBoard() As Integer

    End Function
    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_FreeEnumBoard",
        CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_FreeEnumBoard(ByVal hEnumBoard As Integer) As Boolean

    End Function
    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_GetBoardCount",
        CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_GetBoardCount(ByVal hEnumBoard As Integer) As Integer

    End Function
    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_ConnectToBridgeByIndex",
        CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_ConnectToBridgeByIndex(ByVal nIndex As Integer) As Integer

    End Function
    'Public Declare Function RTBB_ConnectToBridgeByIndex Lib "RTBBLib.dll" (ByVal nIndex As Integer) As Integer

    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_ConnectToBridgeByCapability",
        CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_ConnectToBridgeByCapability(ByVal nIndex As Integer, ByVal nCapability As Integer) As Integer

    End Function
    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_DisconnectBridge",
        CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_DisconnectBridge(ByVal hDevice As Integer) As Boolean

    End Function
    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_FirmwareCheck",
        CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_FirmwareCheck(ByVal hDevice As Integer, ByRef pMinVerBCD As UInteger, ByRef pCurrentBCD As UInteger) As Boolean

    End Function
    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_GetEnumBoardInfo",
        CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_GetEnumBoardInfo(ByVal hEnumBoard As Integer, ByVal nIndex As Integer) As Integer

    End Function
    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_BIGetVenderName",
        CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_BIGetVenderName(ByVal pInfo As Integer) As IntPtr

    End Function
    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_BIGetControllerName",
        CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_BIGetControllerName(ByVal pInfo As Integer) As IntPtr

    End Function
    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_BIGetLibraryName",
        CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_BIGetLibraryName(ByVal pInfo As Integer) As IntPtr

    End Function
    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_BIGetLbraryPath",
        CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_BIGetLbraryPath(ByVal pInfo As Integer) As IntPtr

    End Function
    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_BIGetFirmwareInfo",
        CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_BIGetFirmwareInfo(ByVal pInfo As Integer) As IntPtr

    End Function
    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_BIGetDevicePath",
        CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_BIGetDevicePath(ByVal pInfo As Integer) As IntPtr

    End Function
    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_BIGetBoardName",
        CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_BIGetBoardName(ByVal pInfo As Integer) As IntPtr

    End Function
    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_BIGetIndexOfDevice",
        CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_BIGetIndexOfDevice(ByVal pInfo As Integer) As Integer

    End Function
    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_BIGetVID",
        CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_BIGetVID(ByVal pInfo As Integer) As UInteger

    End Function
    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_BIGetPID",
        CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_BIGetPID(ByVal pInfo As Integer) As UInteger

    End Function
    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_BIGetCapability",
        CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_BIGetCapability(ByVal pInfo As Integer) As Long

    End Function
    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_BIGetGPIOBitsType",
        CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_BIGetGPIOBitsType(ByVal pInfo As Integer) As Long

    End Function
    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_BIGetGPIOPinCount",
        CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_BIGetGPIOPinCount(ByVal pInfo As Integer) As Long

    End Function
    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_BIGetI2CCount",
        CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_BIGetI2CCount(ByVal pInfo As Integer) As Long

    End Function
    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_BIGetSPICount",
        CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_BIGetSPICount(ByVal pInfo As Integer) As Long

    End Function
    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_BIGetUARTCount",
        CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_BIGetUARTCount(ByVal pInfo As Integer) As Long

    End Function
    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_I2CPutByte",
        CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_I2CPutByte(ByVal hDevice As Integer, ByVal nBus As Integer, ByVal nSlaveAddr As Integer, ByVal nData As Byte) As Integer

    End Function
    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_I2CPutByteData",
        CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_I2CPutByteData(ByVal hDevice As Integer, ByVal nBus As Integer, ByVal nSlaveAddr As Integer, ByVal nCmd As Byte, ByVal nData As Byte) As Integer

    End Function
    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_I2CPutWordData",
        CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_I2CPutWordData(ByVal hDevice As Integer, ByVal nBus As Integer, ByVal nSlaveAddr As Integer, ByVal nCmd As Byte, ByVal nData As Integer) As Integer

    End Function
    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_I2CGetByte",
        CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_I2CGetByte(ByVal hDevice As Integer, ByVal nBus As Integer, ByVal nSlaveAddr As Integer) As Integer

    End Function

    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_I2CGetByteData",
        CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_I2CGetByteData(ByVal hDevice As Integer, ByVal nBus As Integer, ByVal nSlaveAddr As Integer, ByVal nCmd As Byte) As Integer

    End Function
    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_I2CGetWordData",
        CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_I2CGetWordData(ByVal hDevice As Integer, ByVal nBus As Integer, ByVal nSlaveAddr As Integer, ByVal nCmd As Byte) As Integer

    End Function
    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_I2CRead",
        CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_I2CRead(ByVal hDevice As Integer, ByVal nBus As Integer, ByVal nSlaveAddr As Integer, ByVal nCmdSize As Integer, ByVal nCmd As Integer, ByVal nBufferSize As Integer, ByRef Buffer As Byte) As Integer

    End Function
    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_I2CWrite",
        CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_I2CWrite(ByVal hDevice As Integer, ByVal nBus As Integer, ByVal nSlaveAddr As Integer, ByVal nCmdSize As Integer, ByVal nCmd As Integer, ByVal nBufferSize As Integer, ByRef Buffer As Byte) As Integer

    End Function
    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_I2CSetFrequency",
        CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_I2CSetFrequency(ByVal hDevice As Integer, ByVal nBus As Integer, ByVal eFrequency As Integer, ByVal nFreqkHz As Integer) As Integer

    End Function
    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_I2CGetCurrentFrequency",
        CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_I2CGetCurrentFrequency(ByVal hDevice As Integer, ByVal nBus As Integer, ByRef eFrequency As Integer, ByRef nFreqkHz As Integer) As Integer

    End Function
    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_I2CGetFrqeuencyCapability",
        CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_I2CGetFrqeuencyCapability(ByVal hDevice As Integer, ByVal nBus As Integer, ByRef pFrequencyCapability As Long, ByRef pMaxFrequencykHz As Integer) As Integer

    End Function
    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_I2CScanSlaveDevice",
        CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_I2CScanSlaveDevice(ByVal hDevice As Integer, ByVal nBus As Integer, ByRef pI2CAvailableAddressVB As Byte) As Integer

    End Function
    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_I2CGetFirstValidSlaveAddr",
        CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_I2CGetFirstValidSlaveAddr(ByRef pI2CAvailableAddress As Byte, ByVal startPos As Integer) As Integer

    End Function
    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_GPIOSetIODirection",
        CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_GPIOSetIODirection(ByVal hDevice As Integer, ByVal nPort As Integer, ByVal nMask As UInteger, ByVal nValue As UInteger) As Integer

    End Function
    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_GPIOGetIODirection",
        CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_GPIOGetIODirection(ByVal hDevice As Integer, ByVal nPort As Integer, ByRef pValue As UInteger) As Integer

    End Function
    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_GPIOWrite",
        CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_GPIOWrite(ByVal hDevice As Integer, ByVal nPort As Integer, ByVal nMask As UInteger, ByVal nValue As UInteger) As Integer

    End Function
    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_GPIORead",
        CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_GPIORead(ByVal hDevice As Integer, ByVal nPort As Integer, ByRef pValue As UInteger) As Integer

    End Function
    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_GPIOSingleSetIODirection",
        CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_GPIOSingleSetIODirection(ByVal hDevice As Integer, ByVal nPinNumber As Integer, ByVal bValue As Boolean) As Integer
    End Function
    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_GPIOSingleWrite",
    CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_GPIOSingleWrite(ByVal hDevice As Integer, ByVal nPinNumber As Integer, ByVal bValue As Boolean) As Integer

    End Function
    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_GPIOSingleRead",
        CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_GPIOSingleRead(ByVal hDevice As Integer, ByVal nPinNumber As Integer, ByRef pValue As Boolean) As Integer

    End Function
    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_GPIOPinNumber2PortNumber",
        CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_GPIOPinNumber2PortNumber(ByVal hDevice As Integer, ByVal nPinNumber As Integer, ByRef pPortNumber As Integer) As Integer

    End Function
    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_GPIOGetPinCount",
        CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_GPIOGetPinCount(ByVal hDevice As Integer, ByRef pCount As Integer) As Integer

    End Function
    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_GPIOGetPinNameAndMode",
        CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_GPIOGetPinNameAndMode(ByVal hDevice As Integer, ByVal nPinNumber As Integer, ByRef pMode As Integer) As IntPtr

    End Function
    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_GPIOExtSetPinMode",
        CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_GPIOExtSetPinMode(ByVal hDevice As Integer, ByVal nPin As Integer, ByVal nMode As UInteger) As Integer

    End Function
    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_GPIOExtGetCurrentPinMode",
        CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_GPIOExtGetCurrentPinMode(ByVal hDevice As Integer, ByVal nPin As Integer, ByRef pMode As UInteger) As Integer

    End Function
    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_GPIOExtGetPinModeCount",
        CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_GPIOExtGetPinModeCount(ByVal hDevice As Integer, ByRef pCount As UInteger) As Integer

    End Function
    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_GPIOExtGetPinSelCount",
        CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_GPIOExtGetPinSelCount(ByVal hDevice As Integer, ByRef pCount As UInteger) As Integer

    End Function
    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_GPIOExtGetPinModeName",
        CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_GPIOExtGetPinModeName(ByVal hDevice As Integer, ByVal nMode As UInteger) As IntPtr

    End Function
    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_GPIOExtGetPinODMode",
        CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_GPIOExtGetPinODMode(ByVal hDevice As Integer, ByVal nPin As Integer, ByRef pOpenDrain As Boolean) As Integer

    End Function
    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_GPIOExtSetPinODMode",
        CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_GPIOExtSetPinODMode(ByVal hDevice As Integer, ByVal nPin As Integer, ByVal pOpenDrain As Boolean) As Integer

    End Function
    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_GPIOExtSetPinODMode",
        CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_GPIOExtSetPinSel(ByVal hDevice As Integer, ByVal nPin As Integer, ByVal nSel As UInteger) As Integer

    End Function
    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_GPIOExtGetPinSelName",
        CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_GPIOExtGetPinSelName(ByVal hDevice As Integer, ByVal nPinNumber As UInteger, ByVal nSel As UInteger) As IntPtr

    End Function
    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_EXTGSMW_SendData",
        CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_EXTGSMW_SendData(ByVal hDevice As Integer, ByVal nLength As UInteger, ByVal nMask As UInteger, ByRef pBuffer As UShort, ByVal bCheckAck As Boolean) As Integer
    End Function

    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_EXTGSMW_GetDataMaxCount",
        CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_EXTGSMW_GetDataMaxCount(ByVal hDevice As Integer) As Integer
    End Function

    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_EXTGSMW_SetBaseClk",
        CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_EXTGSMW_SetBaseClk(ByVal hDevice As Integer, ByVal nClkNs As UInteger, ByVal nClkMode As Byte) As Integer
    End Function
    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_EXTGSMW_GetBaseClk",
        CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_EXTGSMW_GetBaseClk(ByVal hDevice As Integer) As Integer
    End Function
    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_EXTGSMW_GetPinCount",
        CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_EXTGSMW_GetPinCount(ByVal hDevice As Integer) As Integer
    End Function
    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_EXTGSMW_GetAllPinsMask",
        CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_EXTGSMW_GetAllPinsMask(ByVal hDevice As Integer) As Integer
    End Function


    '<DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_SPISetActiveLow",
    '    CallingConvention:=CallingConvention.Cdecl)>
    'Public Function RTBB_SPISetActiveLow(ByVal hDevice As Integer, ByVal bActiveLow As Boolean) As Integer

    'End Function
    '<DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_SPIPutByteData",
    '    CallingConvention:=CallingConvention.Cdecl)>
    'Public Function RTBB_SPIPutByteData(ByVal hDevice As Integer, ByVal nBus As Integer, ByVal nCmd As Byte, ByVal nData As Byte) As Integer

    'End Function
    '<DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_SPIPutWordData",
    '    CallingConvention:=CallingConvention.Cdecl)>
    'Public Function RTBB_SPIPutWordData(ByVal hDevice As Integer, ByVal nBus As Integer, ByVal nCmd As Byte, ByVal nData As Integer) As Integer

    'End Function


    '/*RTBB_EXTBATCHCMD*/

    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Ansi, Pack:=1)>
    Public Structure batchcmd_info
        Public pc As Integer
        Public err As Byte
        Public total_cmd_len As Integer
        Public resvb0 As Byte
        Public resvb1 As Byte
        Public resvb2 As Byte
        Public resv0 As Integer
        Public resv1 As Integer
        Public resv2 As Integer
        Public resv3 As Integer
        Public resv4 As Integer
        Public resv5 As Integer
        Public resv6 As Integer
        Public resv7 As Integer
        Public resv8 As Integer
        Public resv9 As Integer
        Public resv10 As Integer
        Public resv11 As Integer
        Public resv12 As Integer
        
    End Structure
    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_EXTBATCHCMD_Reset",
        CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_EXTBATCHCMD_Reset(ByVal hDevice As Integer) As Integer

    End Function
    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_EXTBATCHCMD_Trigger",
        CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_EXTBATCHCMD_Trigger(ByVal hDevice As Integer) As Integer

    End Function
    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_EXTBATCHCMD_GetStatus",
        CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_EXTBATCHCMD_GetStatus(ByVal hDevice As Integer, ByRef pStatus As Byte) As Integer

    End Function
    'Status: 
    '0:_IDLE,
    '1:_RUN,
    '2:_PAUSE,
    '3: BORTING




    'Public   Function RTBB_EXTBATCHCMD_GetInfo   (ByVal hDevice As Integer,batchcmd_info_t *pInfo) As Integer
    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_EXTBATCHCMD_GetInfo",
        CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_EXTBATCHCMD_GetInfo(ByVal hDevice As Integer, ByRef pInfo As batchcmd_info) As Integer  '

    End Function
    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_EXTBATCHCMD_ReadMemory",
        CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_EXTBATCHCMD_ReadMemory(ByVal hDevice As Integer, ByVal mem_adr As UInteger, ByRef pData As Byte, ByVal len As UInteger) As Integer

    End Function

    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_EXTBATCHCMD_SetPause",
        CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_EXTBATCHCMD_SetPause(ByVal hDevice As Integer, ByVal pause As Boolean) As Integer

    End Function
    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_EXTBATCHCMD_PutTimeoutEnable",
        CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_EXTBATCHCMD_PutTimeoutEnable(ByVal hDevice As Integer, ByVal mili_sec As Integer, ByVal throwable As Boolean) As Integer

    End Function
    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_EXTBATCHCMD_PutTimeoutDisable",
        CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_EXTBATCHCMD_PutTimeoutDisable(ByVal hDevice As Integer) As Integer

    End Function
    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_EXTBATCHCMD_PutI2CWriteFromData",
        CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_EXTBATCHCMD_PutI2CWriteFromData(ByVal hDevice As Integer, ByVal nSlaveAddr As Byte, ByRef pData As Byte, ByVal len As Integer, ByVal throwable As Boolean, ByVal retry As Integer, ByVal retry_us As Integer, ByVal nBus As Byte) As Integer
    End Function
    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_EXTBATCHCMD_PutI2CWriteFromMemory",
        CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_EXTBATCHCMD_PutI2CWriteFromMemory(ByVal hDevice As Integer, ByVal nSlaveAddr As Byte, ByVal mem_addr As UInteger, ByVal len As UInteger, ByVal throwable As Boolean, ByVal retry As UInteger, ByVal retry_us As UInteger, ByVal nBus As Byte) As Integer

    End Function
    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_EXTBATCHCMD_PutI2CRead2Memory",
        CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_EXTBATCHCMD_PutI2CRead2Memory(ByVal hDevice As Integer, ByVal nSlaveAddr As Byte, ByVal mem_addr As UInteger, ByVal len As UInteger, ByVal throwable As Boolean, ByVal retry As UInteger, ByVal retry_us As UInteger, ByVal nBus As Byte) As Integer

    End Function
    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_EXTBATCHCMD_PutWriteMemory",
        CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_EXTBATCHCMD_PutWriteMemory(ByVal hDevice As Integer, ByVal mem_addr As UShort, ByRef pData As Byte, ByVal len As UInteger) As Integer

    End Function
    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_EXTBATCHCMD_PutUpdateMemory",
        CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_EXTBATCHCMD_PutUpdateMemory(ByVal hDevice As Integer, ByVal mem_addr As UShort, ByRef pData As Byte, ByRef pMask As Byte, ByVal len As UInteger) As Integer

    End Function
    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_EXTBATCHCMD_PutSpiRWFromData",
        CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_EXTBATCHCMD_PutSpiRWFromData(ByVal hDevice As Integer, ByRef pData As Byte, ByVal len As UInteger, ByVal rmem_addr As UInteger, ByVal nBus As Byte) As Integer

    End Function
    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_EXTBATCHCMD_PutSpiRWFromMemory",
        CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_EXTBATCHCMD_PutSpiRWFromMemory(ByVal hDevice As Integer, ByVal wmem_addr As UInteger, ByVal rmem_addr As UInteger, ByVal len As UInteger, ByVal nBus As Byte) As Integer

    End Function


    'nPort : Port Ex:P2.5 --> nPort=2
    'nMask : Ex:P2.5 --> 0b0010 0000 , nMask = 0x20, P2.0 -->0b0000 0001  , nMask = 0x01


    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_EXTBATCHCMD_PutGpioSetValue",
        CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_EXTBATCHCMD_PutGpioSetValue(ByVal hDevice As Integer, ByVal nPort As Integer, ByVal nMask As UInteger, ByVal nValue As UInteger) As Integer

    End Function
    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_EXTBATCHCMD_PutGpioSetDir",
        CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_EXTBATCHCMD_PutGpioSetDir(ByVal hDevice As Integer, ByVal nPort As Integer, ByVal nMask As UInteger, ByVal nDir As UInteger) As Integer

    End Function
    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_EXTBATCHCMD_PutGpioReadValue2Memory",
        CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_EXTBATCHCMD_PutGpioReadValue2Memory(ByVal hDevice As Integer, ByVal port As Integer, ByVal mem_addr As UInteger) As Integer

    End Function

    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_EXTBATCHCMD_PutBeq",
        CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_EXTBATCHCMD_PutBeq(ByVal hDevice As Integer, ByVal mem_addr As UShort, ByVal mask As Byte, ByVal val As Byte, ByVal pc As UInteger) As Integer

    End Function
    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_EXTBATCHCMD_PutBneq",
        CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_EXTBATCHCMD_PutBneq(ByVal hDevice As Integer, ByVal mem_addr As UShort, ByVal mask As Byte, ByVal val As Byte, ByVal pc As UInteger) As Integer

    End Function
    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_EXTBATCHCMD_PutSbeq",
        CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_EXTBATCHCMD_PutSbeq(ByVal hDevice As Integer, ByVal mask As Byte, ByVal value As Byte, ByVal pc As UInteger) As Integer

    End Function
    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_EXTBATCHCMD_PutSbneq",
        CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_EXTBATCHCMD_PutSbneq(ByVal hDevice As Integer, ByVal mask As Byte, ByVal value As Byte, ByVal pc As UInteger) As Integer

    End Function
    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_EXTBATCHCMD_PutSMov",
        CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_EXTBATCHCMD_PutSMov(ByVal hDevice As Integer, ByVal mem_addr As UInteger) As Integer

    End Function
    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_EXTBATCHCMD_PutJump",
        CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_EXTBATCHCMD_PutJump(ByVal hDevice As Integer, ByVal pc As UInteger) As Integer

    End Function
    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_EXTBATCHCMD_PutExit",
        CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_EXTBATCHCMD_PutExit(ByVal hDevice As Integer) As Integer

    End Function
    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_EXTBATCHCMD_PutDelayMs",
        CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_EXTBATCHCMD_PutDelayMs(ByVal hDevice As Integer, ByVal ms As Integer) As Integer

    End Function
    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_EXTBATCHCMD_PutNopLoop",
        CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_EXTBATCHCMD_PutNopLoop(ByVal hDevice As Integer, ByVal count As UInteger) As Integer

    End Function
    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_EXTBATCHCMD_PutUnaryOp",
        CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_EXTBATCHCMD_PutUnaryOp(ByVal hDevice As Integer, ByVal mem_addr As UShort, ByVal op As Byte) As Integer

    End Function
    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_EXTBATCHCMD_PutNopDelayUs",
        CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_EXTBATCHCMD_PutNopDelayUs(ByVal hDevice As Integer, ByVal predUs As Double, ByRef realUs As Double) As Integer

    End Function
    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_EXTBATCHCMD_AllocMemory",
        CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_EXTBATCHCMD_AllocMemory(ByVal hDevice As Integer, ByVal len As UInteger, ByRef pMemAdr As UShort) As Integer
    End Function
    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_EXTBATCHCMD_PutHLI2CReadUntilMatch",
        CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_EXTBATCHCMD_PutHLI2CReadUntilMatch(ByVal hDevice As Integer, ByVal nSlaveAddr As Byte, ByVal reg As UInteger, ByVal reg_len As Byte, ByRef pData As Byte, ByRef pMask As Byte, ByVal data_len As UInteger, ByVal timeout_milisec As UInteger, ByVal nBus As Byte) As Integer

    End Function
    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_EXTBATCHCMD_PutHLI2CRead2Memory",
        CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_EXTBATCHCMD_PutHLI2CRead2Memory(ByVal hDevice As Integer, ByVal nSlaveAddr As Byte, ByVal reg As UInteger, ByVal reg_len As Byte, ByRef mem_addr As UInteger, ByVal len As UInteger, ByVal throwable As Boolean, ByVal retry As UInteger, ByVal retry_us As UInteger, ByVal nBus As Byte) As Integer

    End Function


    '判斷GPIO為多少data 就會停，要在timeout_milisec時間內
    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_EXTBATCHCMD_PutHLGpioReadUntilMatch",
        CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_EXTBATCHCMD_PutHLGpioReadUntilMatch(ByVal hDevice As Integer, ByVal nPort As Integer, ByVal data As UInteger, ByVal mask As Byte, ByVal timeout_milisec As UInteger) As Integer

    End Function

    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_SPISetActiveLow", CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_SPISetActiveLow(ByVal hDevice As Integer, ByVal bActiveLow As Boolean) As Integer
    End Function
    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_SPIPutByteData", CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_SPIPutByteData(ByVal hDevice As Integer, ByVal nBus As Integer, ByVal nCmd As Byte, ByVal nData As Byte) As Integer
    End Function
    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_SPIPutWordData", CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_SPIPutWordData(ByVal hDevice As Integer, ByVal nBus As Integer, ByVal nCmd As Byte, ByVal nData As Integer) As Integer
    End Function
    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_SPIGetByteData", CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_SPIGetByteData(ByVal hDevice As Integer, ByVal nBus As Integer, ByVal nCmd As Byte) As Integer
    End Function
    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_SPIGetWordData", CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_SPIGetWordData(ByVal hDevice As Integer, ByVal nBus As Integer, ByVal nCmd As Byte) As Integer
    End Function
    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_SPIPutByte", CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_SPIPutByte(ByVal hDevice As Integer, ByVal nBus As Integer, ByVal nData As Byte) As Integer
    End Function
    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_SPIGetByte", CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_SPIGetByte(ByVal hDevice As Integer, ByVal nBus As Integer) As Integer
    End Function
    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_SPIRead", CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_SPIRead(ByVal hDevice As Integer, ByVal nBus As Integer, ByVal nLength As Integer, ByRef pReadBuffer As Integer) As Integer
    End Function
    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_SPIWrite", CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_SPIWrite(ByVal hDevice As Integer, ByVal nBus As Integer, ByVal nLength As Integer, ByRef pWriteBuffer As Integer) As Integer
    End Function
    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_SPIReadWrite", CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_SPIReadWrite(ByVal hDevice As Integer, ByVal nBus As Integer, ByVal nLength As Integer, ByRef pWriteBuffer As Integer, ByRef pReadBuffer As Integer) As Integer
    End Function
    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_SPIHLRead", CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_SPIHLRead(ByVal hDevice As Integer, ByVal nBus As Integer, ByVal nCmdSize As Byte, ByVal nBufferLength As Integer, ByVal nCmd As UInteger, ByRef pBuffer As Byte) As Integer
    End Function
    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_SPIHLWrite", CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_SPIHLWrite(ByVal hDevice As Integer, ByVal nBus As Integer, ByVal nCmdSize As Byte, ByVal nBufferLength As Integer, ByVal nCmd As UInteger, ByRef pBuffer As Byte) As Integer
    End Function
    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_SPIHLReadWrite", CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_SPIHLReadWrite(ByVal hDevice As Integer, ByVal nBus As Integer, ByVal nCmdSize As Byte, ByVal nBufferLength As Integer, ByVal nCmd As UInteger, ByRef pBuffer As Byte) As Integer
    End Function
    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_SPIPutByteDataCS", CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_SPIPutByteDataCS(ByVal hDevice As Integer, ByVal nBus As Integer, ByVal nCmd As Byte, ByVal nData As Byte, ByVal nPinNumber As Integer) As Integer
    End Function
    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_SPIPutWordDataCS", CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_SPIPutWordDataCS(ByVal hDevice As Integer, ByVal nBus As Integer, ByVal nCmd As Byte, ByVal nData As Integer, ByVal nPinNumber As Integer) As Integer
    End Function
    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_SPIGetByteDataCS", CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_SPIGetByteDataCS(ByVal hDevice As Integer, ByVal nBus As Integer, ByVal nCmd As Byte, ByVal nPinNumber As Integer) As Integer
    End Function
    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_SPIGetWordDataCS", CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_SPIGetWordDataCS(ByVal hDevice As Integer, ByVal nBus As Integer, ByVal nCmd As Byte, ByVal nPinNumber As Integer) As Integer
    End Function
    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_SPIPutByteCS", CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_SPIPutByteCS(ByVal hDevice As Integer, ByVal nBus As Integer, ByVal nData As Byte, ByVal nPinNumber As Integer) As Integer
    End Function
    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_SPIGetByteCS", CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_SPIGetByteCS(ByVal hDevice As Integer, ByVal nBus As Integer, ByVal nPinNumber As Integer) As Integer
    End Function
    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_SPIReadCS", CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_SPIReadCS(ByVal hDevice As Integer, ByVal nBus As Integer, ByVal nLength As Integer, ByRef pReadBuffer As Byte, ByVal nPinNumber As Integer) As Integer
    End Function
    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_SPIWriteCS", CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_SPIWriteCS(ByVal hDevice As Integer, ByVal nBus As Integer, ByVal nLength As Integer, ByRef pWriteBuffer As Byte, ByVal nPinNumber As Integer) As Integer
    End Function
    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_SPIReadWriteCS", CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_SPIReadWriteCS(ByVal hDevice As Integer, ByVal nBus As Integer, ByVal nLength As Integer, ByRef pWriteBuffer As Integer, ByRef pReadBuffer As Integer, ByVal nPinNumber As Integer) As Integer
    End Function
    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_SPIHLReadCS", CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_SPIHLReadCS(ByVal hDevice As Integer, ByVal nBus As Integer, ByVal nPinNumber As Integer, ByVal nCmdSize As Byte, ByVal nBufferLength As Integer, ByVal nCmd As UInteger, ByRef pBuffer As Byte) As Integer
    End Function
    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_SPIHLWriteCS", CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_SPIHLWriteCS(ByVal hDevice As Integer, ByVal nBus As Integer, ByVal nPinNumber As Integer, ByVal nCmdSize As Byte, ByVal nBufferLength As Integer, ByVal nCmd As UInteger, ByRef pBuffer As Byte) As Integer
    End Function
    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_SPIHLReadWriteCS", CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_SPIHLReadWriteCS(ByVal hDevice As Integer, ByVal nBus As Integer, ByVal nPinNumber As Integer, ByVal nCmdSize As Byte, ByVal nBufferLength As Integer, ByVal nCmd As UInteger, ByRef pBuffer As Byte) As Integer
    End Function
    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_SPIChipSelect", CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_SPIChipSelect(ByVal hDevice As Integer, ByVal nPinNumber As Integer) As Integer
    End Function
    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_SPIChipUnselect", CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_SPIChipUnselect(ByVal hDevice As Integer, ByVal nPinNumber As Integer) As Integer
    End Function
    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_SPIGetFrqeuencyCapability", CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_SPIGetFrqeuencyCapability(ByVal hDevice As Integer, ByVal nBus As Integer, ByRef pFrequencyCapability As UInteger, ByRef pMaxFreqkHz As UInteger) As Integer
    End Function
    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_SPISetMode", CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_SPISetMode(ByVal hDevice As Integer, ByVal nBus As Integer, ByVal eMode As UInteger) As Integer
    End Function
    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_SPIGetMode", CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_SPIGetMode(ByVal hDevice As Integer, ByVal nBus As Integer, ByRef pModes As UInteger) As Integer
    End Function
    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_SPISetFrequency", CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_SPISetFrequency(ByVal hDevice As Integer, ByVal nBus As Integer, ByVal EPFrequencyMode As UInteger, ByVal nFreqkHz As UInteger) As Integer
    End Function
    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_SPIGetCurrentFrequency", CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_SPIGetCurrentFrequency(ByVal hDevice As Integer, ByVal nBus As Integer, ByRef pFrequencyMode As UInteger, ByRef pFreqkHz As UInteger) As Integer
    End Function
    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_SPISetCSDelay", CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_SPISetCSDelay(ByVal hDevice As Integer, ByVal nBus As Integer, ByVal nNanoSecond As UInteger) As Integer
    End Function
    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_SPISetFrameDelay", CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_SPISetFrameDelay(ByVal hDevice As Integer, ByVal nBus As Integer, ByVal nNanoSecond As UInteger) As Integer
    End Function
    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_SPIGetCSCount", CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_SPIGetCSCount(ByVal hDevice As Integer, ByVal nBus As Integer) As Integer
    End Function




    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_GetIsoBoardCount", CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_GetIsoBoardCount(ByVal hEnumIsoBoard As Integer) As Integer
    End Function

    '2023/07/11

    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_FreeEnumIsoBoard", CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_FreeEnumIsoBoard(ByVal hEnumIsoBoard As Integer) As Integer
    End Function

    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_DisconnectIso", CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_DisconnectIso(ByVal hIsoDevice As Integer) As Integer
    End Function

    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_GetIsoBoardInfo", CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_GetIsoBoardInfo(ByVal hIsoDevice As Integer) As Integer
    End Function

    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_ConnectToIsoBoardAsBridgeByIndex", CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_ConnectToIsoBoardAsBridgeByIndex(ByVal hEnumIsoBoard As Integer, ByVal nIndex As Integer) As Integer
    End Function

    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_EnumIsolatedBoard", CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_EnumIsolatedBoard(ByVal pBoardNames As String(), ByVal nameCount As Integer) As Integer
    End Function

    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_ConnectToIsoBoardByIndex", CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_ConnectToIsoBoardByIndex(ByVal hEnumIsoBoard As Integer, ByVal nIndex As Integer) As Integer
    End Function

    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_Iso_Transact", CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_Iso_Transact(ByVal hIsoDevice As Integer,
                                                                 ByVal pCmdIn As Integer(), ByVal pDataInCount As Integer(), ByVal pDataIn As Byte(),
                                                                 ByVal pCmdOut As Integer(), ByVal pDataOutCount As Integer(), ByVal pDataOut As Byte()) As Integer
    End Function



    <DllImportAttribute(RTBBDLLNAME, EntryPoint:="RTBB_Iso_Ping", CallingConvention:=CallingConvention.Cdecl)>
    Public Function RTBB_Iso_Ping(ByVal hIsoDevice As Integer) As Integer
    End Function

    ' * -> ByReF
    'uint8_t -> Byte


End Module
