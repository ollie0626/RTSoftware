using System;
using System.Runtime.InteropServices;

namespace RTBBLibDotNet
{
    public interface II2CModule : IBaseModule
    {
        int RTBB_I2CWrite(int nSlaveAddr, int nCmdSize, int nCmd, int nBufferSize, byte[] pBuffer);
        int RTBB_I2CRead(int nSlaveAddr, int nCmdSize, int nCmd, int nBufferSize, byte[] pBuffer);
        int RTBB_I2CMultiRW(int nCount, int[] rw_list, int[] rc_list, int[] slaveAddr_list, int[] cmdSize_list, int[] cmd_list, int[] bufferSize_list, byte[,] Buffer);
        int RTBB_I2CPutByte(int nSlaveAddr, byte nData);
        int RTBB_I2CPutByteData(int nSlaveAddr, byte nCmd, byte nData);
        int RTBB_I2CPutWordData(int nSlaveAddr, byte nCmd, UInt16 nData);
        int RTBB_I2CGetByte(int nSlaveAddr);
        int RTBB_I2CGetByteData(int nSlaveAddr, byte nCmd);
        int RTBB_I2CGetWordData(int nSlaveAddr, byte nCmd);
        int RTBB_I2CSetFrequency(GlobalVariable.ERTI2CFrequency nMode, int nFreqkHz);
        int RTBB_I2CGetCurrentFrequency(ref GlobalVariable.ERTI2CFrequency pFrequency, ref uint pFrequencykHz);
        int RTBB_I2CGetFrqeuencyCapability(ref uint pFrequencyCapability, ref uint pMaxFrequencykHz);
        int RTBB_I2CScanSlaveDevice(ref GlobalVariable.I2CSLAVEADDR pI2CAvailableAddress);
        int RTBB_I2CGetFirstValidSlaveAddr(ref GlobalVariable.I2CSLAVEADDR pI2CAvailableAddress, int startPos);
    }

    public class I2CModule : GlobalVariable, II2CModule
    {
        private IntPtr hDev = IntPtr.Zero;
        private int mBusIndex = 0;

        public I2CModule(IntPtr hDevice)
        {
            hDev = hDevice;
        }

        public I2CModule(IntPtr hDevice, int busIndex)
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
            return "I2C" + mBusIndex.ToString();
        }

        ///<summary>
        ///Description: I2C Write.
        ///Input Parameters: nSlaveAddr -> Address of slave device.
        ///Input Parameters: nCmdSize -> The number of bytes of the nCmd.
        ///Input Parameters: nCmd -> I2C command index, or called address of register.
        ///Input Parameters: nBufferSize -> Buffer size that you want to write into the c hip.
        ///Input Parameters: pBuffer -> the buffer that containing the data to be written to the i2c bus.
        ///Return 0 if success, other non-zero value.
        ///Get the error description string from RTBB_Result2String().
        ///</summary>
        public int RTBB_I2CWrite(int nSlaveAddr, int nCmdSize, int nCmd, int nBufferSize, byte[] pBuffer)
        {
            return native_RTBB_I2CWrite(hDev, mBusIndex, nSlaveAddr, nCmdSize, nCmd, nBufferSize, pBuffer);
        }

        ///<summary>
        ///Description: I2C Read.
        ///Input Parameters: nSlaveAddr -> Address of slave device.
        ///Input Parameters: nCmdSize -> The number of bytes of the nCmd.
        ///Input Parameters: nCmd -> I2C command index, or called address of register.
        ///Input Parameters: nBufferSize -> The number of bytes to be read from the i2c bus.
        ///                              -> The maximum value is 256. ( Limited by the RTBridgeBoard library ).
        ///Output Parameters: Buffer -> the buffer that containing the data to be read from the i2c bus.
        ///Return 0 if success, other non-zero value.
        ///Get the error description string from RTBB_Result2String().
        ///</summary>
        public int RTBB_I2CRead(int nSlaveAddr, int nCmdSize, int nCmd, int nBufferSize, byte[] pBuffer)
        {
            return native_RTBB_I2CRead(hDev, mBusIndex, nSlaveAddr, nCmdSize, nCmd, nBufferSize, pBuffer);
        }

        ///<summary>
        ///Description: I2C Multi RW.
        ///Input Parameters: nCount -> The number of transmissions which want to be executed.
        ///Input Parameters: rw_list -> A pointer to a list that contains nCount fields, and each field respectively represents the transmission's direction.
        ///                          -> 0 is mean it is a write transmission,  and 1 is mean it is a read transmission.
        ///Input Parameters: rc_list ->  A pointer to a list that contains nCount fields, and each field respectively represents the transmission's result.
        ///                          -> If the function succeeds, the return value is zero. Otherwise, the return value is nonzero.
        ///                          -> To get result description string, call RTBB_Result2String().
        ///Input Parameters: slaveAddr_list -> A pointer to a list that contains nCount fields, and each field respectively represents a slave's address used in the transmission.
        ///Input Parameters: cmdSize_list ->  A pointer to a list that contains nCount fields, and each field respectively represents a length of the register's address used in the transmission.
        ///Input Parameters: cmd_list -> A pointer to a list that contains nCount fields, and each field respectively represents a register's address used in the transmission.
        ///                               -> Register address is also called i2c command.
        ///Input Parameters: bufferSize_list -> A pointer to a list that contains nCount fields, and each field respectively represents a number of buffers used in the transmission.
        ///Input/Output Parameters: pBuffer -> A pointer to a data buffer, must be size [][256].
        ///Return 0 if success, other non-zero value.
        ///Get the error description string from RTBB_Result2String().
        ///</summary>
        public int RTBB_I2CMultiRW(int nCount, int[] rw_list, int[] rc_list, int[] slaveAddr_list, int[] cmdSize_list, int[] cmd_list, int[] bufferSize_list, byte[,] Buffer)
        {
            return trans_RTBB_I2CMultiRW(hDev, mBusIndex, nCount, rw_list, rc_list, slaveAddr_list, cmdSize_list, cmd_list, bufferSize_list, Buffer);
        }

        ///<summary>
        ///Description: I2C Put Byte.
        ///Input Parameters: nSlaveAddr -> Address of slave device.
        ///Input Parameters: nData -> the byte data to be written.
        ///Return 0 if success, other non-zero value.
        ///Get the error description string from RTBB_Result2String().
        ///</summary>
        public int RTBB_I2CPutByte(int nSlaveAddr, byte nData)
        {
            return native_RTBB_I2CPutByte(hDev, mBusIndex, nSlaveAddr, nData);
        }

        ///<summary>
        ///Description: I2C Put Byte Data with one command.
        ///Input Parameters: nSlaveAddr -> Address of slave device.
        ///Input Parameters: nCmd -> I2C command index, or called address of register.
        ///Input Parameters: nData -> the byte data to be written.
        ///Return 0 if success, other non-zero value.
        ///Get the error description string from RTBB_Result2String().
        ///</summary>
        public int RTBB_I2CPutByteData(int nSlaveAddr, byte nCmd, byte nData)
        {
            return native_RTBB_I2CPutByteData(hDev, mBusIndex, nSlaveAddr, nCmd, nData);
        }

        ///<summary>
        ///Description: I2C Put Word Data with one command.
        ///Input Parameters: nBus ->  Index of i2c bus.
        ///Input Parameters: nSlaveAddr -> Address of slave device.
        ///Input Parameters: nCmd -> I2C command index, or called address of register.
        ///Input Parameters: nData -> the word data to be written.
        ///Return 0 if success, other non-zero value.
        ///Get the error description string from RTBB_Result2String().
        ///</summary>
        public int RTBB_I2CPutWordData(int nSlaveAddr, byte nCmd, UInt16 nData)
        {
            return native_RTBB_I2CPutWordData(hDev, mBusIndex, nSlaveAddr, nCmd, nData);
        }

        ///<summary>
        ///Description: I2C Get Byte Data.
        ///Input Parameters: nSlaveAddr -> Address of slave device.
        ///If the function succeeds, the return value is the return data from the i2c bus.
        ///If the function fails, the return value is negative.
        ///</summary>
        public int RTBB_I2CGetByte(int nSlaveAddr)
        {
            return native_RTBB_I2CGetByte(hDev, mBusIndex, nSlaveAddr);
        }

        ///<summary>
        ///Description: I2C Get Byte Data with one Command.
        ///Input Parameters: nSlaveAddr -> Address of slave device.
        ///Input Parameters: nCmd -> I2C command index, or called address of register.
        ///If the function succeeds, the return value is the return data from the i2c bus.
        ///If the function fails, the return value is negative.
        ///</summary>
        public int RTBB_I2CGetByteData(int nSlaveAddr, byte nCmd)
        {
            return native_RTBB_I2CGetByteData(hDev, mBusIndex, nSlaveAddr, nCmd);
        }

        ///<summary>
        ///Description: I2C Get Word Data with one Command.
        ///Input Parameters: nSlaveAddr -> Address of slave device.
        ///Input Parameters: nCmd -> I2C command index, or called address of register.
        ///If the function succeeds, the return value is the return data from the i2c bus.
        ///If the function fails, the return value is negative.
        ///</summary>
        public int RTBB_I2CGetWordData(int nSlaveAddr, byte nCmd)
        {
            return native_RTBB_I2CGetWordData(hDev, mBusIndex, nSlaveAddr, nCmd);
        }

        ///<summary>
        ///Description: I2C Set Frequency.
        ///Input Parameters: nBus ->  Index of i2c bus.
        ///Input Parameters: ERTI2CFrequency ->  The new frequency mode setting of i2c bus..
        ///Input Parameters: nFreqkHz -> The new frequency value of i2c bus. ( custom mode ).
        ///                           -> If nFrequencyMode is not eRTI2CFreqCustom, this value will be ignore.
        ///If the function succeeds, the return value is zero.
        ///Otherwise, the return value is nonzero.
        ///</summary>
        public int RTBB_I2CSetFrequency(ERTI2CFrequency nMode, int nFreqkHz)
        {
            return native_RTBB_I2CSetFrequency(hDev, mBusIndex, nMode, nFreqkHz);
        }

        ///<summary>
        ///Description: I2C Get Current Frequency.
        ///Input Parameters: ERTI2CFrequency ->  The current frequency mode setting of i2c bus..
        ///Input Parameters: pFrequencykHz -> The current frequency value of i2c bus.
        ///If the function succeeds, the return value is zero.
        ///Otherwise, the return value is nonzero.
        ///</summary>
        public int RTBB_I2CGetCurrentFrequency(ref ERTI2CFrequency pFrequency, ref uint pFrequencykHz)
        {
            return native_RTBB_I2CGetCurrentFrequency(hDev, mBusIndex, ref pFrequency, ref pFrequencykHz);
        }

        ///<summary>
        ///Description: I2C Get Frequency Capability.
        ///Input Parameters: nBus ->  Index of i2c bus.
        ///Input Parameters: pFrequencyCapability -> Frequency modes of i2c bus support.
        ///                                       -> This variable is a combination of the ERTI2CFrequency enumeration.
        ///Input Parameters: pMaxFrequencykHz -> Maximum frequency of i2c bus support,  unit is KHz.
        ///If the function succeeds, the return value is zero.
        ///Otherwise, the return value is nonzero.
        ///</summary>
        public int RTBB_I2CGetFrqeuencyCapability(ref uint pFrequencyCapability, ref uint pMaxFrequencykHz)
        {
            return native_RTBB_I2CGetFrqeuencyCapability(hDev, mBusIndex, ref pFrequencyCapability, ref pMaxFrequencykHz);
        }

        ///<summary>
        ///Description: I2C Scan Slave Device.
        ///Input Parameters: I2CSLAVEADDR -> A pointer to I2CSLAVEADDR structure that receives information about available slave addresses.
        ///                               -> Be careful that is does not directly represent a i2c salve address.
        ///If the function succeeds, the return value is zero.
        ///Otherwise, the return value is nonzero.
        ///</summary>
        public int RTBB_I2CScanSlaveDevice(ref I2CSLAVEADDR pI2CAvailableAddress)
        {
            return native_RTBB_I2CScanSlaveDevice(hDev, mBusIndex, ref pI2CAvailableAddress);
        }

        ///<summary>
        ///Input Parameters: pI2CAvailableAddress -> A pointer to I2CSLAVEADDR structure which represent available slave address on i2c bus, it get from RTBB_I2CScanSlaveDevice().
        ///Input Parameters: startPos -> Starting position of searching available address, max value is 127.
        ///If a available slave address is between [startPos~127], then return value is the first valid slave address.
        ///Otherwise, the return value is negative.
        ///</summary>
        public int RTBB_I2CGetFirstValidSlaveAddr(ref I2CSLAVEADDR pI2CAvailableAddress, int startPos)
        {
            return native_RTBB_I2CGetFirstValidSlaveAddr(ref pI2CAvailableAddress, startPos);
        }

        ///<summary>
        ///Description: I2C Get Bus Count.
        ///If the function succeeds, the return value is the quantities of available I2C bus on Bridgeboard.
        ///If the function fails, the return value is negative.
        ///To get result description string, call RTBB_Result2String().
        ///</summary>
        private int RTBB_I2CGetBusCount()
        {
            return native_RTBB_I2CGetBusCount(hDev);
        }

        ///<summary>
        ///addr is the struct that get from i2c scan device.
        ///x is the slave addr, ex, 0x45 = 69.
        ///Return the x slave addr is valid or not..
        ///</summary>
        public bool I2C_SLAVE_ADDR_IS_VALID(I2CSLAVEADDR addr, int x)
        {
            return (((addr.addr[x / 32]) & (1 << (x % 32))).Equals(0) ? false : true);
        }

        /* I2C Functions */
        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_I2CWrite")]
        private static extern int native_RTBB_I2CWrite(IntPtr hDevice, int nBus, int nSlaveAddr, int nCmdSize, int nCmd, int nBufferSize, byte[] pBuffer);

        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_I2CRead")]
        private static extern int native_RTBB_I2CRead(IntPtr hDevice, int nBus, int nSlaveAddr, int nCmdSize, int nCmd, int nBufferSize, byte[] pBuffer);

        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_I2CMultiRW")]
        private static extern int native_RTBB_I2CMultiRW(IntPtr hDevice, int nBus, int nCount, int[] rw_list, int[] rc_list, int[] slaveAddr_list, int[] cmdSize_list, int[] cmd_list, int[] bufferSize_list, IntPtr pBuffer);

        private int trans_RTBB_I2CMultiRW(IntPtr hDevice, int nBus, int nCount, int[] rw_list, int[] rc_list, int[] slaveAddr_list, int[] cmdSize_list, int[] cmd_list, int[] bufferSize_list, byte[,] Buffer)
        {
            IntPtr pBuffer = Marshal.AllocHGlobal(256 * nCount);
            int ret;
            for (int i = 0; i < nCount; i++)
                for (int j = 0; j < 256; j++)
                    Marshal.WriteByte(pBuffer, (i * 256) + j, Buffer[i, j]);
            ret = native_RTBB_I2CMultiRW(hDevice, nBus, nCount, rw_list, rc_list, slaveAddr_list, cmdSize_list, cmd_list, bufferSize_list, pBuffer);
            for (int i = 0; i < nCount; i++)
                for (int j = 0; j < 256; j++)
                    Buffer[i, j] = Marshal.ReadByte(pBuffer, (i * 256) + j);
            Marshal.FreeHGlobal(pBuffer);
            return ret;
        }

        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_I2CPutByte")]
        private static extern int native_RTBB_I2CPutByte(IntPtr hDevice, int nBus, int nSlaveAddr, byte nData);

        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_I2CPutByteData")]
        private static extern int native_RTBB_I2CPutByteData(IntPtr hDevice, int nBus, int nSlaveAddr, byte nCmd, byte nData);

        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_I2CPutWordData")]
        private static extern int native_RTBB_I2CPutWordData(IntPtr hDevice, int nBus, int nSlaveAddr, byte nCmd, UInt16 nData);

        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_I2CGetByte")]
        private static extern int native_RTBB_I2CGetByte(IntPtr hDevice, int nBus, int nSlaveAddr);

        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_I2CGetByteData")]
        private static extern int native_RTBB_I2CGetByteData(IntPtr hDevice, int nBus, int nSlaveAddr, byte nCmd);

        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_I2CGetWordData")]
        private static extern int native_RTBB_I2CGetWordData(IntPtr hDevice, int nBus, int nSlaveAddr, byte nCmd);

        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_I2CSetFrequency")]
        private static extern int native_RTBB_I2CSetFrequency(IntPtr hDevice, int nBus, ERTI2CFrequency nMode, int nFreqkHz);

        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_I2CGetCurrentFrequency")]
        private static extern int native_RTBB_I2CGetCurrentFrequency(IntPtr hDevice, int nBus, ref ERTI2CFrequency pFrequency, ref uint pFrequencykHz);

        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_I2CGetFrqeuencyCapability")]
        private static extern int native_RTBB_I2CGetFrqeuencyCapability(IntPtr hDevice, int nBus, ref uint pFrequencyCapability, ref uint pMaxFrequencykHz);

        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_I2CScanSlaveDevice")]
        private static extern int native_RTBB_I2CScanSlaveDevice(IntPtr hDevice, int nBus, ref I2CSLAVEADDR pI2CAvailableAddress);

        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_I2CGetFirstValidSlaveAddr")]
        private static extern int native_RTBB_I2CGetFirstValidSlaveAddr(ref I2CSLAVEADDR pI2CAvailableAddress, int startPos);

        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_I2CGetBusCount")]
        private static extern int native_RTBB_I2CGetBusCount(IntPtr hDevice);
    }
}
