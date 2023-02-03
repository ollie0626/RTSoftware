'==== Author: Hungdar Wang ====
'==== Version:  2014.10.28 ====
Module RTBBLibDll
    Public Declare Function RTBB_EnumBoard Lib "RTBBLib.dll" () As Integer
    Public Declare Function RTBB_FreeEnumBoard Lib "RTBBLib.dll" (ByVal hEnumBoard As Integer) As Boolean
    Public Declare Function RTBB_GetBoardCount Lib "RTBBLib.dll" (ByVal hEnumBoard As Integer) As Integer
    Public Declare Function RTBB_ConnectToBridgeByIndex Lib "RTBBLib.dll" (ByVal nIndex As Integer) As Integer
    Public Declare Function RTBB_ConnectToBridgeByCapability Lib "RTBBLib.dll" (ByVal nIndex As Integer, ByVal nCapability As Integer) As Integer
    Public Declare Function RTBB_DisconnectBridge Lib "RTBBLib.dll" (ByVal hDevice As Integer) As Boolean
    Public Declare Function RTBB_FirmwareCheck Lib "RTBBLib.dll" (ByVal hDevice As Integer, ByRef pMinVerBCD As UInteger, ByRef pCurrentBCD As UInteger) As Boolean
    Public Declare Function RTBB_GetEnumBoardInfo Lib "RTBBLib.dll" (ByVal hEnumBoard As Integer, ByVal nIndex As Integer) As Integer
    Public Declare Function RTBB_BIGetVenderName Lib "RTBBLib.dll" (ByVal pInfo As Integer) As IntPtr
    Public Declare Function RTBB_BIGetControllerName Lib "RTBBLib.dll" (ByVal pInfo As Integer) As IntPtr
    Public Declare Function RTBB_BIGetLibraryName Lib "RTBBLib.dll" (ByVal pInfo As Integer) As IntPtr
    Public Declare Function RTBB_BIGetLbraryPath Lib "RTBBLib.dll" (ByVal pInfo As Integer) As IntPtr
    Public Declare Function RTBB_BIGetFirmwareInfo Lib "RTBBLib.dll" (ByVal pInfo As Integer) As IntPtr
    Public Declare Function RTBB_BIGetDevicePath Lib "RTBBLib.dll" (ByVal pInfo As Integer) As IntPtr
    Public Declare Function RTBB_BIGetBoardName Lib "RTBBLib.dll" (ByVal pInfo As Integer) As IntPtr
    Public Declare Function RTBB_BIGetIndexOfDevice Lib "RTBBLib.dll" (ByVal pInfo As Integer) As Integer
    Public Declare Function RTBB_BIGetVID Lib "RTBBLib.dll" (ByVal pInfo As Integer) As UInteger
    Public Declare Function RTBB_BIGetPID Lib "RTBBLib.dll" (ByVal pInfo As Integer) As UInteger
    Public Declare Function RTBB_BIGetCapability Lib "RTBBLib.dll" (ByVal pInfo As Integer) As Long
    Public Declare Function RTBB_BIGetGPIOBitsType Lib "RTBBLib.dll" (ByVal pInfo As Integer) As Long
    Public Declare Function RTBB_BIGetGPIOPinCount Lib "RTBBLib.dll" (ByVal pInfo As Integer) As Long
    Public Declare Function RTBB_BIGetI2CCount Lib "RTBBLib.dll" (ByVal pInfo As Integer) As Long
    Public Declare Function RTBB_BIGetSPICount Lib "RTBBLib.dll" (ByVal pInfo As Integer) As Long
    Public Declare Function RTBB_BIGetUARTCount Lib "RTBBLib.dll" (ByVal pInfo As Integer) As Long
    Public Declare Function RTBB_I2CPutByte Lib "RTBBLib.dll" (ByVal hDevice As Integer, ByVal nBus As Integer, ByVal nSlaveAddr As Integer, ByVal nData As Byte) As Integer
    Public Declare Function RTBB_I2CPutByteData Lib "RTBBLib.dll" (ByVal hDevice As Integer, ByVal nBus As Integer, ByVal nSlaveAddr As Integer, ByVal nCmd As Byte, ByVal nData As Byte) As Integer
    Public Declare Function RTBB_I2CPutWordData Lib "RTBBLib.dll" (ByVal hDevice As Integer, ByVal nBus As Integer, ByVal nSlaveAddr As Integer, ByVal nCmd As Byte, ByVal nData As Integer) As Integer
    Public Declare Function RTBB_I2CGetByte Lib "RTBBLib.dll" (ByVal hDevice As Integer, ByVal nBus As Integer, ByVal nSlaveAddr As Integer) As Integer
    Public Declare Function RTBB_I2CGetByteData Lib "RTBBLib.dll" (ByVal hDevice As Integer, ByVal nBus As Integer, ByVal nSlaveAddr As Integer, ByVal nCmd As Byte) As Integer
    Public Declare Function RTBB_I2CGetWordData Lib "RTBBLib.dll" (ByVal hDevice As Integer, ByVal nBus As Integer, ByVal nSlaveAddr As Integer, ByVal nCmd As Byte) As Integer
    Public Declare Function RTBB_I2CRead Lib "RTBBLib.dll" (ByVal hDevice As Integer, ByVal nBus As Integer, ByVal nSlaveAddr As Integer, ByVal nCmdSize As Integer, ByVal nCmd As Integer, ByVal nBufferSize As Integer, ByRef Buffer As Byte) As Integer
    Public Declare Function RTBB_I2CWrite Lib "RTBBLib.dll" (ByVal hDevice As Integer, ByVal nBus As Integer, ByVal nSlaveAddr As Integer, ByVal nCmdSize As Integer, ByVal nCmd As Integer, ByVal nBufferSize As Integer, ByRef Buffer As Byte) As Integer
    Public Declare Function RTBB_I2CSetFrequency Lib "RTBBLib.dll" (ByVal hDevice As Integer, ByVal nBus As Integer, ByVal eFrequency As Integer, ByVal nFreqkHz As Integer) As Integer
    Public Declare Function RTBB_I2CGetCurrentFrequency Lib "RTBBLib.dll" (ByVal hDevice As Integer, ByVal nBus As Integer, ByRef eFrequency As Integer, ByRef nFreqkHz As Integer) As Integer
    Public Declare Function RTBB_I2CGetFrqeuencyCapability Lib "RTBBLib.dll" (ByVal hDevice As Integer, ByVal nBus As Integer, ByRef pFrequencyCapability As Long, ByRef pMaxFrequencykHz As Integer) As Integer
    Public Declare Function RTBB_I2CScanSlaveDevice Lib "RTBBLib.dll" (ByVal hDevice As Integer, ByVal nBus As Integer, ByRef pI2CAvailableAddressVB As Byte) As Integer
    Public Declare Function RTBB_I2CGetFirstValidSlaveAddr Lib "RTBBLib.dll" (ByRef pI2CAvailableAddress As Byte, ByVal startPos As Integer) As Integer
    Public Declare Function RTBB_GPIOSetIODirection Lib "RTBBLib.dll" (ByVal hDevice As Integer, ByVal nPort As Integer, ByVal nMask As UInteger, ByVal nValue As UInteger) As Integer
    Public Declare Function RTBB_GPIOGetIODirection Lib "RTBBLib.dll" (ByVal hDevice As Integer, ByVal nPort As Integer, ByRef pValue As UInteger) As Integer
    Public Declare Function RTBB_GPIOWrite Lib "RTBBLib.dll" (ByVal hDevice As Integer, ByVal nPort As Integer, ByVal nMask As UInteger, ByVal nValue As UInteger) As Integer
    Public Declare Function RTBB_GPIORead Lib "RTBBLib.dll" (ByVal hDevice As Integer, ByVal nPort As Integer, ByRef pValue As UInteger) As Integer
    Public Declare Function RTBB_GPIOSingleSetIODirection Lib "RTBBLib.dll" (ByVal hDevice As Integer, ByVal nPinNumber As Integer, ByVal bValue As Boolean) As Integer
    Public Declare Function RTBB_GPIOSingleWrite Lib "RTBBLib.dll" (ByVal hDevice As Integer, ByVal nPinNumber As Integer, ByVal bValue As Boolean) As Integer
    Public Declare Function RTBB_GPIOSingleRead Lib "RTBBLib.dll" (ByVal hDevice As Integer, ByVal nPinNumber As Integer, ByRef pValue As Boolean) As Integer
    Public Declare Function RTBB_GPIOPinNumber2PortNumber Lib "RTBBLib.dll" (ByVal hDevice As Integer, ByVal nPinNumber As Integer, ByRef pPortNumber As Integer) As Integer
    Public Declare Function RTBB_GPIOGetPinCount Lib "RTBBLib.dll" (ByVal hDevice As Integer, ByRef pCount As Integer) As Integer
    Public Declare Function RTBB_GPIOGetPinNameAndMode Lib "RTBBLib.dll" (ByVal hDevice As Integer, ByVal nPinNumber As Integer, ByRef pMode As Integer) As IntPtr
    Public Declare Function RTBB_GPIOExtSetPinMode Lib "RTBBLib.dll" (ByVal hDevice As Integer, ByVal nPin As Integer, ByVal nMode As UInteger) As Integer
    Public Declare Function RTBB_GPIOExtGetCurrentPinMode Lib "RTBBLib.dll" (ByVal hDevice As Integer, ByVal nPin As Integer, ByRef pMode As UInteger) As Integer
    Public Declare Function RTBB_GPIOExtGetPinModeCount Lib "RTBBLib.dll" (ByVal hDevice As Integer, ByRef pCount As UInteger) As Integer
    Public Declare Function RTBB_GPIOExtGetPinSelCount Lib "RTBBLib.dll" (ByVal hDevice As Integer, ByRef pCount As UInteger) As Integer
    Public Declare Function RTBB_GPIOExtGetPinModeName Lib "RTBBLib.dll" (ByVal hDevice As Integer, ByVal nMode As UInteger) As IntPtr
    Public Declare Function RTBB_GPIOExtGetPinODMode Lib "RTBBLib.dll" (ByVal hDevice As Integer, ByVal nPin As Integer, ByRef pOpenDrain As Boolean) As Integer
    Public Declare Function RTBB_GPIOExtSetPinODMode Lib "RTBBLib.dll" (ByVal hDevice As Integer, ByVal nPin As Integer, ByVal pOpenDrain As Boolean) As Integer
    Public Declare Function RTBB_GPIOExtSetPinSel Lib "RTBBLib.dll" (ByVal hDevice As Integer, ByVal nPin As Integer, ByVal nSel As UInteger) As Integer
    Public Declare Function RTBB_GPIOExtGetPinSelName Lib "RTBBLib.dll" (ByVal hDevice As Integer, ByVal nPinNumber As UInteger, ByVal nSel As UInteger) As IntPtr
    Public Declare Function RTBB_SPISetActiveLow Lib "RTBBLib.dll" (ByVal hDevice As Integer, ByVal bActiveLow As Boolean) As Integer
    Public Declare Function RTBB_SPIPutByteData Lib "RTBBLib.dll" (ByVal hDevice As Integer, ByVal nBus As Integer, ByVal nCmd As Byte, ByVal nData As Byte) As Integer
    Public Declare Function RTBB_SPIPutWordData Lib "RTBBLib.dll" (ByVal hDevice As Integer, ByVal nBus As Integer, ByVal nCmd As Byte, ByVal nData As Integer) As Integer
    Public Declare Function RTBB_SPIGetByteData Lib "RTBBLib.dll" (ByVal hDevice As Integer, ByVal nBus As Integer, ByVal nCmd As Byte) As Integer
    Public Declare Function RTBB_SPIGetWordData Lib "RTBBLib.dll" (ByVal hDevice As Integer, ByVal nBus As Integer, ByVal nCmd As Byte) As Integer
    Public Declare Function RTBB_SPIPutByte Lib "RTBBLib.dll" (ByVal hDevice As Integer, ByVal nBus As Integer, ByVal nData As Byte) As Integer
    Public Declare Function RTBB_SPIGetByte Lib "RTBBLib.dll" (ByVal hDevice As Integer, ByVal nBus As Integer) As Integer
    Public Declare Function RTBB_SPIRead Lib "RTBBLib.dll" (ByVal hDevice As Integer, ByVal nBus As Integer, ByVal nLength As Integer, ByRef pReadBuffer As Integer) As Integer
    Public Declare Function RTBB_SPIWrite Lib "RTBBLib.dll" (ByVal hDevice As Integer, ByVal nBus As Integer, ByVal nLength As Integer, ByRef pWriteBuffer As Integer) As Integer
    Public Declare Function RTBB_SPIReadWrite Lib "RTBBLib.dll" (ByVal hDevice As Integer, ByVal nBus As Integer, ByVal nLength As Integer, ByRef pWriteBuffer As Integer, ByRef pReadBuffer As Integer) As Integer
    Public Declare Function RTBB_SPIHLRead Lib "RTBBLib.dll" (ByVal hDevice As Integer, ByVal nBus As Integer, ByVal nCmdSize As Byte, ByVal nBufferLength As Integer, ByVal nCmd As UInteger, ByRef pBuffer As Byte) As Integer
    Public Declare Function RTBB_SPIHLWrite Lib "RTBBLib.dll" (ByVal hDevice As Integer, ByVal nBus As Integer, ByVal nCmdSize As Byte, ByVal nBufferLength As Integer, ByVal nCmd As UInteger, ByRef pBuffer As Byte) As Integer
    Public Declare Function RTBB_SPIHLReadWrite Lib "RTBBLib.dll" (ByVal hDevice As Integer, ByVal nBus As Integer, ByVal nCmdSize As Byte, ByVal nBufferLength As Integer, ByVal nCmd As UInteger, ByRef pBuffer As Byte) As Integer
    Public Declare Function RTBB_SPIPutByteDataCS Lib "RTBBLib.dll" (ByVal hDevice As Integer, ByVal nBus As Integer, ByVal nCmd As Byte, ByVal nData As Byte, ByVal nPinNumber As Integer) As Integer
    Public Declare Function RTBB_SPIPutWordDataCS Lib "RTBBLib.dll" (ByVal hDevice As Integer, ByVal nBus As Integer, ByVal nCmd As Byte, ByVal nData As Integer, ByVal nPinNumber As Integer) As Integer
    Public Declare Function RTBB_SPIGetByteDataCS Lib "RTBBLib.dll" (ByVal hDevice As Integer, ByVal nBus As Integer, ByVal nCmd As Byte, ByVal nPinNumber As Integer) As Integer
    Public Declare Function RTBB_SPIGetWordDataCS Lib "RTBBLib.dll" (ByVal hDevice As Integer, ByVal nBus As Integer, ByVal nCmd As Byte, ByVal nPinNumber As Integer) As Integer
    Public Declare Function RTBB_SPIPutByteCS Lib "RTBBLib.dll" (ByVal hDevice As Integer, ByVal nBus As Integer, ByVal nData As Byte, ByVal nPinNumber As Integer) As Integer
    Public Declare Function RTBB_SPIGetByteCS Lib "RTBBLib.dll" (ByVal hDevice As Integer, ByVal nBus As Integer, ByVal nPinNumber As Integer) As Integer
    Public Declare Function RTBB_SPIReadCS Lib "RTBBLib.dll" (ByVal hDevice As Integer, ByVal nBus As Integer, ByVal nLength As Integer, ByRef pReadBuffer As Byte, ByVal nPinNumber As Integer) As Integer
    Public Declare Function RTBB_SPIWriteCS Lib "RTBBLib.dll" (ByVal hDevice As Integer, ByVal nBus As Integer, ByVal nLength As Integer, ByRef pWriteBuffer As Byte, ByVal nPinNumber As Integer) As Integer
    Public Declare Function RTBB_SPIReadWriteCS Lib "RTBBLib.dll" (ByVal hDevice As Integer, ByVal nBus As Integer, ByVal nLength As Integer, ByRef pWriteBuffer As Integer, ByRef pReadBuffer As Integer, ByVal nPinNumber As Integer) As Integer
    Public Declare Function RTBB_SPIHLReadCS Lib "RTBBLib.dll" (ByVal hDevice As Integer, ByVal nBus As Integer, ByVal nPinNumber As Integer, ByVal nCmdSize As Byte, ByVal nBufferLength As Integer, ByVal nCmd As UInteger, ByRef pBuffer As Byte) As Integer
    Public Declare Function RTBB_SPIHLWriteCS Lib "RTBBLib.dll" (ByVal hDevice As Integer, ByVal nBus As Integer, ByVal nPinNumber As Integer, ByVal nCmdSize As Byte, ByVal nBufferLength As Integer, ByVal nCmd As UInteger, ByRef pBuffer As Byte) As Integer
    Public Declare Function RTBB_SPIHLReadWriteCS Lib "RTBBLib.dll" (ByVal hDevice As Integer, ByVal nBus As Integer, ByVal nPinNumber As Integer, ByVal nCmdSize As Byte, ByVal nBufferLength As Integer, ByVal nCmd As UInteger, ByRef pBuffer As Byte) As Integer
    Public Declare Function RTBB_SPIChipSelect Lib "RTBBLib.dll" (ByVal hDevice As Integer, ByVal nPinNumber As Integer) As Integer
    Public Declare Function RTBB_SPIChipUnselect Lib "RTBBLib.dll" (ByVal hDevice As Integer, ByVal nPinNumber As Integer) As Integer
    Public Declare Function RTBB_SPIGetFrqeuencyCapability Lib "RTBBLib.dll" (ByVal hDevice As Integer, ByVal nBus As Integer, ByRef pFrequencyCapability As UInteger, ByRef pMaxFreqkHz As UInteger) As Integer
    Public Declare Function RTBB_SPISetMode Lib "RTBBLib.dll" (ByVal hDevice As Integer, ByVal nBus As Integer, ByVal eMode As UInteger) As Integer
    Public Declare Function RTBB_SPIGetMode Lib "RTBBLib.dll" (ByVal hDevice As Integer, ByVal nBus As Integer, ByRef pModes As UInteger) As Integer
    Public Declare Function RTBB_SPISetFrequency Lib "RTBBLib.dll" (ByVal hDevice As Integer, ByVal nBus As Integer, ByVal EPFrequencyMode As UInteger, ByVal nFreqkHz As UInteger) As Integer
    Public Declare Function RTBB_SPIGetCurrentFrequency Lib "RTBBLib.dll" (ByVal hDevice As Integer, ByVal nBus As Integer, ByRef pFrequencyMode As UInteger, ByRef pFreqkHz As UInteger) As Integer
    Public Declare Function RTBB_SPISetCSDelay Lib "RTBBLib.dll" (ByVal hDevice As Integer, ByVal nBus As Integer, ByVal nNanoSecond As UInteger) As Integer
    Public Declare Function RTBB_SPISetFrameDelay Lib "RTBBLib.dll" (ByVal hDevice As Integer, ByVal nBus As Integer, ByVal nNanoSecond As UInteger) As Integer
    Public Declare Function RTBB_SPIGetCSCount Lib "RTBBLib.dll" (ByVal hDevice As Integer, ByVal nBus As Integer) As Integer
    Public Declare Function RTBB_EXTGPIOMISC_GetUniversalGPIOCount Lib "RTBBLib.dll" (ByVal hDevice As Integer, ByRef pUniversaPinCount As UInteger) As Integer
    Public Declare Function RTBB_EXTGPIOMISC_GetUniversalGPIOMapping Lib "RTBBLib.dll" (ByVal hDevice As Integer, ByVal nUniversalPinNumber As UInteger, ByRef pGPIOPinNumber As UInteger) As Integer       'get pin mapping : for example, UniversalPin 0 = P0[13],  UinversalPin 1 = P1[16]... etc...
    Public Declare Function RTBB_EXTGPIOMISC_GetSPICSPinCount Lib "RTBBLib.dll" (ByVal hDevice As Integer, ByRef pSPICSPinCount As UInteger) As Integer
    Public Declare Function RTBB_EXTGPIOMISC_GetSPIPinMapping Lib "RTBBLib.dll" (ByVal hDevice As Integer, ByVal nSPICSPinNumber As UInteger, ByRef pGPIOPinNumber As UInteger) As Integer      'get pin mapping : for example, CS0 = P0[16]
    Public Declare Function RTBB_EXTGPIOMISC_GetGPIOPinMode Lib "RTBBLib.dll" (ByVal hDevice As Integer, ByVal nGPIOPinNumber As UInteger, ByRef pPinMode As UInteger) As Integer
    Public Declare Function RTBB_EXTGPIOMISC_SetGPIOPinMode Lib "RTBBLib.dll" (ByVal hDevice As Integer, ByVal nGPIOPinNumber As UInteger, ByVal nPinMode As UInteger) As Integer
    Public Declare Function RTBB_EXTHSI2C_SetHSCode Lib "RTBBLib.dll" (ByVal hDevice As Integer, ByRef pCode As Byte, ByVal nLength As UInteger) As Integer
    Public Declare Function RTBB_EXTHSI2C_I2CWrite Lib "RTBBLib.dll" (ByVal hDevice As Integer, ByVal nBus As Integer, ByVal nSlaveAddr As Integer, ByVal nCmdSize As Integer, ByVal nCmd As Integer, ByVal nBufferSize As Integer, ByRef pBuffer As Byte) As Integer       'for this version, nBus must be zero
    Public Declare Function RTBB_EXTHSI2C_I2CRead Lib "RTBBLib.dll" (ByVal hDevice As Integer, ByVal nBus As Integer, ByVal nSlaveAddr As Integer, ByVal nCmdSize As Integer, ByVal nCmd As Integer, ByVal nBufferSize As Integer, ByRef pBuffer As Byte) As Integer        'for this version, nBus must be zero
    Public Declare Function RTBB_I2CMultiRW Lib "RTBBLib.dll" (ByVal hDevice As Integer, ByVal nBus As Integer, ByVal nCount As Integer, ByRef rw_list As Integer, ByRef rc_list As Integer, ByRef slaveAddr_list As Integer, ByRef cmdSize_list As Integer, ByRef cmd_list As Integer, ByRef bufferSize_list As Integer, ByRef pBuffer As Byte) As Integer

    Public Declare Function RTBB_I2CGetBusCount Lib "RTBBLib.dll" (ByVal hDevice As Integer) As Integer
    Public Declare Function RTBB_SPIGetBusCount Lib "RTBBLib.dll" (ByVal hDevice As Integer) As Integer
    Public Declare Function RTBB_PWMGetPWMPortCount Lib "RTBBLib.dll" (ByVal hDevice As Integer) As Integer
    Public Declare Function RTBB_PWMStart Lib "RTBBLib.dll" (ByVal hDevice As Integer, ByVal nPort As Integer) As Integer
    Public Declare Function RTBB_PWMStop Lib "RTBBLib.dll" (ByVal hDevice As Integer, ByVal nPort As Integer) As Integer
    Public Declare Function RTBB_PWMGetStatus Lib "RTBBLib.dll" (ByVal hDevice As Integer, ByVal nPort As Integer) As Integer
    Public Declare Function RTBB_PWMSetDutyCycle Lib "RTBBLib.dll" (ByVal hDevice As Integer, ByVal nPort As Integer, ByVal fDutyCycle As Double) As Integer
    Public Declare Function RTBB_PWMGetDutyCycle Lib "RTBBLib.dll" (ByVal hDevice As Integer, ByVal nPort As Integer, ByRef pfDutyCycle As Double) As Integer
    Public Declare Function RTBB_PWMGetBaseClock Lib "RTBBLib.dll" (ByVal hDevice As Integer, ByVal nPort As Integer, ByRef pnClock As UInteger) As Integer
    Public Declare Function RTBB_PWMSetPeriod Lib "RTBBLib.dll" (ByVal hDevice As Integer, ByVal nPort As Integer, ByVal nTick As UInteger) As Integer
    Public Declare Function RTBB_PWMSetOnPeriod Lib "RTBBLib.dll" (ByVal hDevice As Integer, ByVal nPort As Integer, ByVal nOnTick As UInteger) As Integer
    Public Declare Function RTBB_PWMGetPeriod Lib "RTBBLib.dll" (ByVal hDevice As Integer, ByVal nPort As Integer, ByRef pnTick As UInteger) As Integer
    Public Declare Function RTBB_PWMGetOnPeriod Lib "RTBBLib.dll" (ByVal hDevice As Integer, ByVal nPort As Integer, ByRef pnOnTick As UInteger) As Integer
    Public Declare Function RTBB_PWMSetPeriodByUs Lib "RTBBLib.dll" (ByVal hDevice As Integer, ByVal nPort As Integer, ByVal nUs As UInteger) As Integer
    Public Declare Function RTBB_PWMSetOnPeriodByUs Lib "RTBBLib.dll" (ByVal hDevice As Integer, ByVal nPort As Integer, ByVal nUs As UInteger) As Integer
    Public Declare Function RTBB_PWMSetPeriodByMs Lib "RTBBLib.dll" (ByVal hDevice As Integer, ByVal nPort As Integer, ByVal nMs As UInteger) As Integer
    Public Declare Function RTBB_PWMSetOnPeriodByMs Lib "RTBBLib.dll" (ByVal hDevice As Integer, ByVal nPort As Integer, ByVal nMs As UInteger) As Integer
    Public Declare Function RTBB_PWMGetPeriodByUs Lib "RTBBLib.dll" (ByVal hDevice As Integer, ByVal nPort As Integer, ByRef pnUs As UInteger) As Integer
    Public Declare Function RTBB_PWMGetOnPeriodByUs Lib "RTBBLib.dll" (ByVal hDevice As Integer, ByVal nPort As Integer, ByRef pnUs As UInteger) As Integer
    Public Declare Function RTBB_EXTGSOW_SendData Lib "RTBBLib.dll" (ByVal hDevice As Integer, ByVal nLength As UInteger, ByVal nMask As UInteger, ByRef pBuffer As UShort, ByVal bCheckAck As Boolean) As Integer  '2014.03.18
    Public Declare Function RTBB_EXTGSOW_GetDataMaxCount Lib "RTBBLib.dll" (ByVal hDevice As Integer) As Integer
    Public Declare Function RTBB_EXTGSOW_SetBaseClk Lib "RTBBLib.dll" (ByVal hDevice As Integer, ByVal nClkNs As UInteger, ByVal nClkMode As Byte) As Integer
    Public Declare Function RTBB_EXTGSOW_GetBaseClk Lib "RTBBLib.dll" (ByVal hDevice As Integer) As Integer
    Public Declare Function RTBB_EXTGSMW_SendData Lib "RTBBLib.dll" (ByVal hDevice As Integer, ByVal nLength As UInteger, ByVal nMask As UInteger, ByRef pBuffer As UShort, ByVal bCheckAck As Boolean) As Integer  '2014.03.18
    Public Declare Function RTBB_EXTGSMW_GetDataMaxCount Lib "RTBBLib.dll" (ByVal hDevice As Integer) As Integer
    Public Declare Function RTBB_EXTGSMW_SetBaseClk Lib "RTBBLib.dll" (ByVal hDevice As Integer, ByVal nClkNs As UInteger, ByVal nClkMode As Byte) As Integer
    Public Declare Function RTBB_EXTGSMW_GetBaseClk Lib "RTBBLib.dll" (ByVal hDevice As Integer) As Integer
    Public Declare Function RTBB_EXTGSMW_GetPinCount Lib "RTBBLib.dll" (ByVal hDevice As Integer) As Integer
    Public Declare Function RTBB_EXTGSMW_GetAllPinsMask Lib "RTBBLib.dll" (ByVal hDevice As Integer) As Integer
    Public Declare Function RTBB_EXTSVI2_SVI2GetCurrentFrequency Lib "RTBBLib.dll" (ByVal hDevice As Integer, ByRef pFrequency As Integer) As Integer
    Public Declare Function RTBB_EXTSVI2_SVI2GetFrqeuencyCapability Lib "RTBBLib.dll" (ByVal hDevice As Integer, ByRef pFrequencyCapability As Long) As Integer
    Public Declare Function RTBB_EXTSVI2_SVI2SetFrequency Lib "RTBBLib.dll" (ByVal hDevice As Integer, ByVal nFrequencyMode As Integer) As Integer
    Public Declare Function RTBB_EXTSVI2_SVI2PowerUp Lib "RTBBLib.dll" (ByVal hDevice As Integer, ByVal eBootVIDCode As Integer) As Integer
    Public Declare Function RTBB_EXTSVI2_SVI2PowerDown Lib "RTBBLib.dll" (ByVal hDevice As Integer) As Integer
    Public Declare Function RTBB_EXTSVI2_SVI2IsPowerUp Lib "RTBBLib.dll" (ByVal hDevice As Integer) As Integer
    Public Declare Function RTBB_EXTSVI2_SVI2SendCmd Lib "RTBBLib.dll" (ByVal hDevice As Integer, ByVal VDD_SEL As UInteger, ByVal VDDNB_SEL As UInteger, ByVal PSI0_L As UInteger, ByVal PSI1_L As UInteger, ByVal VID_CODE As UInteger, ByVal TFN As UInteger, ByVal LoadLineSlopeTrim As UInteger, ByVal OffsetTrim As UInteger) As Integer
    Public Declare Function RTBB_EXTSVI2_SVI2RecvData Lib "RTBBLib.dll" (ByVal hDevice As Integer, ByRef nCount As Byte, ByRef RecvData As UInteger) As Integer

    Public Declare Function RTBB_EXTCFW_GetCFWVersion Lib "RTBBLib.dll" (ByVal hDevice As Integer) As Integer
    Public Declare Function RTBB_EXTCFW_GetCFWVendor Lib "RTBBLib.dll" (ByVal hDevice As Integer) As IntPtr
    Public Declare Function RTBB_EXTCFW_CheckCFWVendor Lib "RTBBLib.dll" (ByVal hDevice As Integer, ByVal strVendor As String) As Integer
    Public Declare Function RTBB_EXTCFW_Transact Lib "RTBBLib.dll" (ByVal hDevice As Integer, ByRef pCmdIn As Integer, ByRef pDataInCount As Integer, ByRef pDataIn As Byte, ByRef pCmdOut As Integer, ByRef pDataOutCount As Integer, ByRef pDataOut As Byte) As Integer
    Public Declare Function RTBB_EXTIOCONF_GetIOConfigurableType Lib "RTBBLib.dll" (ByVal hDevice As Integer) As Integer
    Public Declare Function RTBB_EXTIOCONF_GetIOConfigurableList Lib "RTBBLib.dll" (ByVal hDevice As Integer, ByVal nType As Integer, ByRef pIOPort As IntPtr) As Integer
    Public Declare Function RTBB_EXTIOCONF_GetIOVoltageList Lib "RTBBLib.dll" (ByVal hDevice As Integer, ByVal nType As Integer, ByVal nIOPort As Integer, ByRef pVoltageRange As IntPtr) As Integer
    Public Declare Function RTBB_EXTIOCONF_SetIOVoltageRange Lib "RTBBLib.dll" (ByVal hDevice As Integer, ByVal nType As Integer, ByVal nIOPort As Integer, ByVal min_mV As Integer, ByVal max_mV As Integer) As Integer
    Public Declare Function RTBB_EXTIOCONF_SetIOVoltage Lib "RTBBLib.dll" (ByVal hDevice As Integer, ByVal nType As Integer, ByVal nIOPort As Integer, ByVal mV As Integer) As Integer
    Public Declare Function RTBB_EXTIOCONF_GetIOVoltage Lib "RTBBLib.dll" (ByVal hDevice As Integer, ByVal nType As Integer, ByVal nIOPort As Integer) As Integer
    Public Declare Function RTBB_EXTIOCONF_GetIOVoltageSel Lib "RTBBLib.dll" (ByVal hDevice As Integer, ByVal nType As Integer, ByVal nIOPort As Integer) As Integer

    Public Declare Function RTBB_EXTSTORAGE_GetBankNR Lib "RTBBLib.dll" (ByVal hDevice As Integer) As Integer
    Public Declare Function RTBB_EXTSTORAGE_GetBankSize Lib "RTBBLib.dll" (ByVal hDevice As Integer) As Integer
    Public Declare Function RTBB_EXTSTORAGE_BankRead Lib "RTBBLib.dll" (ByVal hDevice As Integer, ByVal nBank As Integer, ByVal nOffset As Integer, ByVal nSize As Integer, ByRef pDest As Byte) As Integer
    Public Declare Function RTBB_EXTSTORAGE_BankWrite Lib "RTBBLib.dll" (ByVal hDevice As Integer, ByVal nBank As Integer, ByVal nOffset As Integer, ByVal nSize As Integer, ByRef pSrc As Byte) As Integer
    Public Declare Function RTBB_EXTSTORAGE_Flush Lib "RTBBLib.dll" (ByVal hDevice As Integer) As Integer
    Public Declare Function RTBB_EXTSECURITYDATA_GetData Lib "RTBBLib.dll" (ByVal hDevice As Integer, ByRef pSize As Integer, ByRef pData As Byte) As Integer
    Public Declare Function RTBB_EXTSECURITYDATA_GetString Lib "RTBBLib.dll" (ByVal hDevice As Integer) As String
End Module
