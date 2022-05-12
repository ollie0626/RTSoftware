using System;
using System.Runtime.InteropServices;

namespace RTBBLibDotNet
{
    public interface ISPIModule : IBaseModule
    {
        int RTBB_SPIGetCSCount();
        int RTBB_SPIGetFrqeuencyCapability(ref UInt32 pFrequencyCapability, ref uint pMaxFreqkHz);
        int RTBB_SPIGetMode(ref UInt32 pMode);
        int RTBB_SPISetMode(UInt32 eMode);
        int RTBB_SPISetFrequency(UInt32 eFrequencyMode, uint nFreqkHz);
        int RTBB_SPISetCSDelay(uint nNanoSecond);
        int RTBB_SPISetFrameDelay(uint nNanoSecond);
        int RTBB_SPISetActiveLow(bool bActiveLow);
        int RTBB_SPIChipSelect(int nPinNumber);
        int RTBB_SPIChipUnselect(int nPinNumber);
        int RTBB_SPIPutByteDataCS(byte nCmd, byte nData, int nPinNumber);
        int RTBB_SPIPutWordDataCS(byte nCmd, UInt16 nData, int nPinNumber);
        int RTBB_SPIGetByteDataCS(byte nCmd, int nPinNumber);
        int RTBB_SPIGetWordDataCS(byte nCmd, int nPinNumber);
        int RTBB_SPIPutByteCS(byte nData, int nPinNumber);
        int RTBB_SPIGetByteCS(int nPinNumber);
        int RTBB_SPIReadCS(int nLength, byte[] pReadBuffer, int nPinNumber);
        int RTBB_SPIWriteCS(int nLength, byte[] pWriteBuffer, int nPinNumber);
        int RTBB_SPIReadWriteCS(int nLength, byte[] pWriteBuffer, byte[] pReadBuffer, int nPinNumber);
        int RTBB_SPIHLReadCS(int nPinNumber, byte nCmdSize, UInt16 nBufferLength, UInt32 nCmd, byte[] pBuffer);
        int RTBB_SPIHLWriteCS(int nPinNumber, byte nCmdSize, UInt16 nBufferLength, UInt32 nCmd, byte[] pBuffer);
        int RTBB_SPIHLReadWriteCS(int nPinNumber, byte nCmdSize, UInt16 nBufferLength, UInt32 nCmd, byte[] pBuffer);
    }

    public class SPIModule : GlobalVariable, ISPIModule
    {
        private IntPtr hDev = IntPtr.Zero;
        private int mBusIndex = 0;

        public SPIModule(IntPtr hDevice)
        {
            hDev = hDevice;
        }

        public SPIModule(IntPtr hDevice, int busIndex)
        {
            hDev = hDevice;
            mBusIndex = busIndex;
        }

        ///<summary>
        ///Description: return the module name.
        ///If the function succeeds, the return value is the module name
        ///</summary>
        public string getModuleName()
        {
            return "SPI" + mBusIndex.ToString();
        }

        ///<summary>
        ///Description: SPI Get CS Pin Count.
        ///Return 0 if success, other non-zero value.
        ///Get the error description string from RTBB_Result2String().
        ///</summary>
        public int RTBB_SPIGetCSCount()
        {
            return native_RTBB_SPIGetCSCount(hDev, mBusIndex);
        }

        ///<summary>
        ///Description: SPI Get Frequency Capability.
        ///Input Parameters: pFrequencyCapability -> This variable is a combination of the ERTSPIFrequency enumeration.
        ///Input Parameters: pMaxFreqkHz -> Max freqKHz for the specified nBus SPI.
        ///Return 0 if success, other non-zero value.
        ///Get the error description string from RTBB_Result2String().
        ///</summary>
        public int RTBB_SPIGetFrqeuencyCapability(ref UInt32 pFrequencyCapability, ref uint pMaxFreqkHz)
        {
            return native_RTBB_SPIGetFrqeuencyCapability(hDev, mBusIndex, ref pFrequencyCapability, ref pMaxFreqkHz);
        }

        ///<summary>
        ///Description: SPI Get Mode.
        ///Input Parameters: pMode ->  ERTSPIMode which indicates the property mode of SPI bus.
        ///Return 0 if success, other non-zero value.
        ///Get the error description string from RTBB_Result2String().
        ///</summary>
        public int RTBB_SPIGetMode(ref UInt32 pMode)
        {
            return native_RTBB_SPIGetMode(hDev, mBusIndex, ref pMode);
        }

        ///<summary>
        ///Description: SPI Set Mode.
        ///Input Parameters: eMode ->  This variable definition please refer to ERTSPIMode enumeration.
        ///Return 0 if success, other non-zero value.
        ///Get the error description string from RTBB_Result2String().
        ///</summary>
        public int RTBB_SPISetMode(UInt32 eMode)
        {
            return native_RTBB_SPISetMode(hDev, mBusIndex, eMode);
        }

        ///<summary>
        ///Description: SPI Set Frequency.
        ///Input Parameters: eFrequencyMode ->  This variable definition please refer to ERTSPIFrequency enumeration.
        ///Input Parameters: nFreqkHz ->   If eFrequencyMode is not eRTSPIFreqCustom, this value will be ignore.
        ///Return 0 if success, other non-zero value.
        ///Get the error description string from RTBB_Result2String().
        ///</summary>
        public int RTBB_SPISetFrequency(UInt32 eFrequencyMode, uint nFreqkHz)
        {
            return native_RTBB_SPISetFrequency(hDev, mBusIndex, eFrequencyMode, nFreqkHz);
        }

        ///<summary>
        ///Description: SPI Set CS Delay.
        ///Input Parameters: nNanoSecond ->  The new chip select delay of SPI bus, unit is nanoSecond.
        ///Return 0 if success, other non-zero value.
        ///Get the error description string from RTBB_Result2String().
        ///</summary>
        public int RTBB_SPISetCSDelay(uint nNanoSecond)
        {
            return native_RTBB_SPISetCSDelay(hDev, mBusIndex, nNanoSecond);
        }

        ///<summary>
        ///Description: SPI Set Freme Delay.
        ///Input Parameters: nNanoSecond ->  The new frame delay of SPI bus, unit is nanoSecond.
        ///Return 0 if success, other non-zero value.
        ///Get the error description string from RTBB_Result2String().
        ///</summary>
        public int RTBB_SPISetFrameDelay(uint nNanoSecond)
        {
            return native_RTBB_SPISetFrameDelay(hDev, mBusIndex, nNanoSecond);
        }

        ///<summary>
        ///Description: SPI Set Active Low or not.
        ///Input Parameters: bActiveLow -> A Boolean variable which indicates chip select default is active low or not.
        ///Return 0 if success, other non-zero value.
        ///Get the error description string from RTBB_Result2String().
        ///</summary>
        public int RTBB_SPISetActiveLow(bool bActiveLow)
        {
            return native_RTBB_SPISetActiveLow(hDev, bActiveLow);
        }

        ///<summary>
        ///Description: SPI Set ChipSelect Pin.
        ///Input Parameters: nPinNumber -> Chip select pin number.
        ///Return 0 if success, other non-zero value.
        ///Get the error description string from RTBB_Result2String().
        ///</summary>
        public int RTBB_SPIChipSelect(int nPinNumber)
        {
            return native_RTBB_SPIChipSelect(hDev, nPinNumber);
        }

        ///<summary>
        ///Description: SPI Unset ChipSelect Pin.
        ///Input Parameters: nPinNumber -> Chip select pin number.
        ///Return 0 if success, other non-zero value.
        ///Get the error description string from RTBB_Result2String().
        ///</summary>
        public int RTBB_SPIChipUnselect(int nPinNumber)
        {
            return native_RTBB_SPIChipUnselect(hDev, nPinNumber);
        }

        ///<summary>
        ///Description: SPI Put Byte Data with one Command.
        ///Input Parameters: nCmd -> Address of register.
        ///Input Parameters: nData -> A byte of data to be written.
        ///Input Parameters: nPinNumber -> Chip select pin number.
        ///Return 0 if success, other non-zero value.
        ///Get the error description string from RTBB_Result2String().
        ///</summary>
        public int RTBB_SPIPutByteDataCS(byte nCmd, byte nData, int nPinNumber)
        {
            return native_RTBB_SPIPutByteDataCS(hDev, mBusIndex, nCmd, nData, nPinNumber);
        }

        ///<summary>
        ///Description: SPI Put Word Data with one Command.
        ///Input Parameters: nBus -> SPI bus number.
        ///Input Parameters: nCmd -> Address of register.
        ///Input Parameters: nData -> A word of data to be written.
        ///Input Parameters: nPinNumber -> Chip select pin number.
        ///Return 0 if success, other non-zero value.
        ///Get the error description string from RTBB_Result2String().
        ///</summary>
        public int RTBB_SPIPutWordDataCS(byte nCmd, UInt16 nData, int nPinNumber)
        {
            return native_RTBB_SPIPutWordDataCS(hDev, mBusIndex, nCmd, nData, nPinNumber);
        }

        ///<summary>
        ///Description: SPI Get Byte Data with one Command.
        ///Input Parameters: nCmd -> Address of register.
        ///Input Parameters: nData -> A word of data to be written.
        ///Input Parameters: nPinNumber -> Chip select pin number.
        ///If the function succeeds, the return value is the requested data from the SPI bus.
        ///If the function fails, the return value is negative.
        ///To get result description string, call RTBB_Result2String().
        ///</summary>
        public int RTBB_SPIGetByteDataCS(byte nCmd, int nPinNumber)
        {
            return native_RTBB_SPIGetByteDataCS(hDev, mBusIndex, nCmd, nPinNumber);
        }

        ///<summary>
        ///Description: SPI Get Word Data with one Command.
        ///Input Parameters: nCmd -> Address of register.
        ///Input Parameters: nData -> A word of data to be written.
        ///Input Parameters: nPinNumber -> Chip select pin number.
        ///If the function succeeds, the return value is the requested data from the SPI bus.
        ///If the function fails, the return value is negative.
        ///To get result description string, call RTBB_Result2String().
        ///</summary>
        public int RTBB_SPIGetWordDataCS(byte nCmd, int nPinNumber)
        {
            return native_RTBB_SPIGetWordDataCS(hDev, mBusIndex, nCmd, nPinNumber);
        }

        ///<summary>
        ///Description: SPI Put Byte Data.
        ///Input Parameters: nCmd -> Address of register.
        ///Input Parameters: nData -> A word of data to be written.
        ///Input Parameters: nPinNumber -> Chip select pin number.
        ///If the function succeeds, the return value is zero.
        ///Otherwise, the return value is nonzero.
        ///To get result description string, call RTBB_Result2String().
        ///</summary>
        public int RTBB_SPIPutByteCS(byte nData, int nPinNumber)
        {
            return native_RTBB_SPIPutByteCS(hDev, mBusIndex, nData, nPinNumber);
        }

        ///<summary>
        ///Description: SPI Get Byte Data.
        ///Input Parameters: nBus -> SPI bus number.
        ///Input Parameters: nPinNumber -> Chip select pin number.
        ///If the function succeeds, the return value is the requested data from the SPI bus.
        ///If the function fails, the return value is negative.
        ///To get result description string, call RTBB_Result2String().
        ///</summary>
        public int RTBB_SPIGetByteCS(int nPinNumber)
        {
            return native_RTBB_SPIGetByteCS(hDev, mBusIndex, nPinNumber);
        }

        ///<summary>
        ///Description: SPI Read Data.
        ///Input Parameters: nLength -> The number of bytes to be read from SPI bus.
        ///Input Parameters: pReadBuffer -> A pointer to the buffer that receives the data read from SPI bus.
        ///Input Parameters: nPinNumber -> Chip select pin number.
        ///If the function succeeds, the return value is zero.
        ///Otherwise, the return value is nonzero.
        ///To get result description string, call RTBB_Result2String().
        ///</summary>
        public int RTBB_SPIReadCS(int nLength, byte[] pReadBuffer, int nPinNumber)
        {
            return native_RTBB_SPIReadCS(hDev, mBusIndex, nLength, pReadBuffer, nPinNumber);
        }

        ///<summary>
        ///Description: SPI Write Data.
        ///Input Parameters: nLength -> The number of bytes to be written to SPI bus.
        ///Input Parameters: pWriteBuffer -> A pointer to the buffer containing the data to be written to SPI bus.
        ///Input Parameters: nPinNumber -> Chip select pin number.
        ///If the function succeeds, the return value is zero.
        ///Otherwise, the return value is nonzero.
        ///To get result description string, call RTBB_Result2String().
        ///</summary>
        public int RTBB_SPIWriteCS(int nLength, byte[] pWriteBuffer, int nPinNumber)
        {
            return native_RTBB_SPIWriteCS(hDev, mBusIndex, nLength, pWriteBuffer, nPinNumber);
        }

        ///<summary>
        ///Description: SPI Read/Write Data.
        ///Input Parameters: nLength -> The number of bytes to be written to SPI bus.
        ///Input Parameters: pWriteBuffer -> A pointer to the buffer containing the data to be written to SPI bus.
        ///Input Parameters: pReadBuffer -> A pointer to the buffer that receives the data read from SPI bus.
        ///Input Parameters: nPinNumber -> Chip select pin number.
        ///If the function succeeds, the return value is zero.
        ///Otherwise, the return value is nonzero.
        ///To get result description string, call RTBB_Result2String().
        ///</summary>
        public int RTBB_SPIReadWriteCS(int nLength, byte[] pWriteBuffer, byte[] pReadBuffer, int nPinNumber)
        {
            return native_RTBB_SPIReadWriteCS(hDev, mBusIndex, nLength, pWriteBuffer, pReadBuffer, nPinNumber);
        }

        ///<summary>
        ///Description: SPI HL Read Data with commands.
        ///Input Parameters: nPinNumber -> Chip select pin number.
        ///Input Parameters: nCmdSize -> The number of bytes of the nCmd.
        ///Input Parameters: nBufferLength ->  The number of bytes to be read from the SPI bus.
        ///                                ->  The maximum value is SPI_BUF_SIZE. ( 1024 ).            
        ///Input Parameters: nCmd -> SPI command.
        ///Input Parameters: pBuffer -> A pointer to the buffer that containing the data to receives the data from the SPI bus.
        ///If the function succeeds, the return value is zero.
        ///Otherwise, the return value is nonzero.
        ///To get result description string, call RTBB_Result2String().
        ///</summary>
        public int RTBB_SPIHLReadCS(int nPinNumber, byte nCmdSize, UInt16 nBufferLength, UInt32 nCmd, byte[] pBuffer)
        {
            return native_RTBB_SPIHLReadCS(hDev, mBusIndex, nPinNumber, nCmdSize, nBufferLength, nCmd, pBuffer);
        }

        ///<summary>
        ///Description: SPI HL Write Data with commands.
        ///Input Parameters: nPinNumber -> Chip select pin number.
        ///Input Parameters: nCmdSize -> The number of bytes of the nCmd.
        ///Input Parameters: nBufferLength ->  The number of bytes to be read from the SPI bus.
        ///                                ->  The maximum value is SPI_BUF_SIZE. ( 1024 ).            
        ///Input Parameters: nCmd -> SPI command.
        ///Input Parameters: pBuffer -> A pointer to the buffer that containing the data to be written to the SPI bus.
        ///If the function succeeds, the return value is zero.
        ///Otherwise, the return value is nonzero.
        ///To get result description string, call RTBB_Result2String().
        ///</summary>
        public int RTBB_SPIHLWriteCS(int nPinNumber, byte nCmdSize, UInt16 nBufferLength, UInt32 nCmd, byte[] pBuffer)
        {
            return native_RTBB_SPIHLWriteCS(hDev, mBusIndex, nPinNumber, nCmdSize, nBufferLength, nCmd, pBuffer);
        }

        ///<summary>
        ///Description: SPI HL Read/Write Data with commands.
        ///Input Parameters: nPinNumber -> Chip select pin number.
        ///Input Parameters: nCmdSize -> The number of bytes of the nCmd.
        ///Input Parameters: nBufferLength ->  The number of bytes to be read from the SPI bus.
        ///                                ->  The maximum value is SPI_BUF_SIZE. ( 1024 ).            
        ///Input Parameters: nCmd -> SPI command.
        ///Input Parameters: pBuffer -> A pointer to the buffer that containing the data to be written to the SPI bus or receives the data from the SPI bus.
        ///If the function succeeds, the return value is zero.
        ///Otherwise, the return value is nonzero.
        ///To get result description string, call RTBB_Result2String().
        ///</summary>
        public int RTBB_SPIHLReadWriteCS(int nPinNumber, byte nCmdSize, UInt16 nBufferLength, UInt32 nCmd, byte[] pBuffer)
        {
            return native_RTBB_SPIHLReadWriteCS(hDev, mBusIndex, nPinNumber, nCmdSize, nBufferLength, nCmd, pBuffer);
        }

        ///<summary>
        ///Description: SPI Get Bus Count.
        ///If the function succeeds, the return value is the quantities of available SPI bus on Bridgeboard.
        ///If the function fails, the return value is negative.
        ///To get result description string, call RTBB_Result2String().
        ///</summary>
        private int RTBB_SPIGetBusCount()
        {
            return native_RTBB_SPIGetBusCount(hDev);
        }

        /* SPI control Functions */
        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_SPIGetCSCount")]
        private static extern int native_RTBB_SPIGetCSCount(IntPtr hDevice, int nBus);

        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_SPIGetFrqeuencyCapability")]
        private static extern int native_RTBB_SPIGetFrqeuencyCapability(IntPtr hDevice, int nBus, ref UInt32 pFrequencyCapability, ref uint pMaxFreqkHz);

        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_SPIGetMode")]
        private static extern int native_RTBB_SPIGetMode(IntPtr hDevice, int nBus, ref UInt32 pMode);

        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_SPISetMode")]
        private static extern int native_RTBB_SPISetMode(IntPtr hDevice, int nBus, UInt32 eMode);

        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_SPISetFrequency")]
        private static extern int native_RTBB_SPISetFrequency(IntPtr hDevice, int nBus, UInt32 eFrequencyMode, uint nFreqkHz);

        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_SPISetCSDelay")]
        private static extern int native_RTBB_SPISetCSDelay(IntPtr hDevice, int nBus, uint nNanoSecond);

        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_SPISetFrameDelay")]
        private static extern int native_RTBB_SPISetFrameDelay(IntPtr hDevice, int nBus, uint nNanoSecond);

        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_SPISetActiveLow")]
        private static extern int native_RTBB_SPISetActiveLow(IntPtr hDevice, bool bActiveLow);

        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_SPIChipSelect")]
        private static extern int native_RTBB_SPIChipSelect(IntPtr hDevice, int nPinNumber);

        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_SPIChipUnselect")]
        private static extern int native_RTBB_SPIChipUnselect(IntPtr hDevice, int nPinNumber);

        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_SPIPutByteDataCS")]
        private static extern int native_RTBB_SPIPutByteDataCS(IntPtr hDevice, int nBus, byte nCmd, byte nData, int nPinNumber);

        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_SPIPutWordDataCS")]
        private static extern int native_RTBB_SPIPutWordDataCS(IntPtr hDevice, int nBus, byte nCmd, UInt16 nData, int nPinNumber);

        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_SPIGetByteDataCS")]
        private static extern int native_RTBB_SPIGetByteDataCS(IntPtr hDevice, int nBus, byte nCmd, int nPinNumber);

        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_SPIGetWordDataCS")]
        private static extern int native_RTBB_SPIGetWordDataCS(IntPtr hDevice, int nBus, byte nCmd, int nPinNumber);

        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_SPIPutByteCS")]
        private static extern int native_RTBB_SPIPutByteCS(IntPtr hDevice, int nBus, byte nData, int nPinNumber);

        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_SPIGetByteCS")]
        private static extern int native_RTBB_SPIGetByteCS(IntPtr hDevice, int nBus, int nPinNumber);

        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_SPIReadCS")]
        private static extern int native_RTBB_SPIReadCS(IntPtr hDevice, int nBus, int nLength, byte[] pReadBuffer, int nPinNumber);

        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_SPIWriteCS")]
        private static extern int native_RTBB_SPIWriteCS(IntPtr hDevice, int nBus, int nLength, byte[] pWriteBuffer, int nPinNumber);

        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_SPIReadWriteCS")]
        private static extern int native_RTBB_SPIReadWriteCS(IntPtr hDevice, int nBus, int nLength, byte[] pWriteBuffer, byte[] pReadBuffer, int nPinNumber);

        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_SPIHLReadCS")]
        private static extern int native_RTBB_SPIHLReadCS(IntPtr hDevice, int nBus, int nPinNumber, byte nCmdSize, UInt16 nBufferLength, UInt32 nCmd, byte[] pBuffer);

        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_SPIHLWriteCS")]
        private static extern int native_RTBB_SPIHLWriteCS(IntPtr hDevice, int nBus, int nPinNumber, byte nCmdSize, UInt16 nBufferLength, UInt32 nCmd, byte[] pBuffer);

        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_SPIHLReadWriteCS")]
        private static extern int native_RTBB_SPIHLReadWriteCS(IntPtr hDevice, int nBus, int nPinNumber, byte nCmdSize, UInt16 nBufferLength, UInt32 nCmd, byte[] pBuffer);

        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_SPIGetBusCount")]
        private static extern int native_RTBB_SPIGetBusCount(IntPtr hDevice);
    }
}
