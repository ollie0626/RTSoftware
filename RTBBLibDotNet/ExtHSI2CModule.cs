using System;
using System.Runtime.InteropServices;

namespace RTBBLibDotNet
{
    public interface IExtHSI2CModule : IBaseModule
    {
        int RTBB_EXTHSI2C_SetHSCode(byte[] pCode, uint nLength);
        int RTBB_EXTHSI2C_I2CWrite(int nBus, int nSlaveAddr, int nCmdSize, int nCmd, int nBufferSize, byte[] pBuffer);
        int RTBB_EXTHSI2C_I2CRead(int nBus, int nSlaveAddr, int nCmdSize, int nCmd, int nBufferSize, byte[] pBuffer);
    }

    public class ExtHSI2CModule : GlobalVariable, IExtHSI2CModule
    {
        private IntPtr hDev = IntPtr.Zero;

        public ExtHSI2CModule(IntPtr hDevice)
        {
            hDev = hDevice;
        }

        ///<summary>
        ///Description: return the module name.
        ///If the function succeeds, the return value is the module name
        ///</summary>
        public string getModuleName()
        {
            return "ExtHSI2C";
        }

        ///<summary>
        ///Description: HSI2C set HS code.
        ///Input Parameters: pCode -> A pointer to BYTE which indicates the high-speed code.
        ///Input Parameters: nLength -> The number of bytes in the high-speed code.
        ///If the function succeeds, the return value is zero.
        ///If the function fails, the return value is negative.
        ///To get result description string, call RTBB_Result2String().
        ///</summary>
        public int RTBB_EXTHSI2C_SetHSCode(byte[] pCode, uint nLength)
        {
            return native_RTBB_EXTHSI2C_SetHSCode(hDev, pCode, nLength);
        }

        ///<summary>
        ///Description: HSI2C I2C Write.
        ///Input Parameters: nBus -> Index of i2c bus. Start number is 0.
        ///Input Parameters: nSlaveAddr -> Address of slave device.
        ///Input Parameters: nCmdSize -> The number of bytes of the nCmd.
        ///Input Parameters: nCmd -> I2C command index, or called address of register.
        ///Input Parameters: nBufferSize -> The number of bytes to be written to the i2c bus.
        ///                              -> The maximum value is 256. ( Limited by the RTBridgeBoard library ).
        ///Input Parameters: pBuffer -> A pointer to the buffer that containing the data to be written to the i2c bus.
        ///If the function succeeds, the return value is zero.
        ///If the function fails, the return value is negative.
        ///To get result description string, call RTBB_Result2String().
        ///</summary>
        public int RTBB_EXTHSI2C_I2CWrite(int nBus, int nSlaveAddr, int nCmdSize, int nCmd, int nBufferSize, byte[] pBuffer)
        {
            return native_RTBB_EXTHSI2C_I2CWrite(hDev, nBus, nSlaveAddr, nCmdSize, nCmd, nBufferSize, pBuffer);
        }

        ///<summary>
        ///Description: HSI2C I2C Read.
        ///Input Parameters: nBus -> Index of i2c bus. Start number is 0.
        ///Input Parameters: nSlaveAddr -> Address of slave device.
        ///Input Parameters: nCmdSize -> The number of bytes of the nCmd.
        ///Input Parameters: nCmd -> I2C command index, or called address of register.
        ///Input Parameters: nBufferSize -> The number of bytes to be written to the i2c bus.
        ///                              -> The maximum value is 256. ( Limited by the RTBridgeBoard library ).
        ///Output Parameters: pBuffer -> A pointer to the buffer that receives the data from the i2c bus.
        ///If the function succeeds, the return value is zero.
        ///If the function fails, the return value is negative.
        ///To get result description string, call RTBB_Result2String().
        ///</summary>
        public int RTBB_EXTHSI2C_I2CRead(int nBus, int nSlaveAddr, int nCmdSize, int nCmd, int nBufferSize, byte[] pBuffer)
        {
            return native_RTBB_EXTHSI2C_I2CRead(hDev, nBus, nSlaveAddr, nCmdSize, nCmd, nBufferSize, pBuffer);
        }

        /* extGPIOMisc control Functions */
        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_EXTHSI2C_SetHSCode")]
        private static extern int native_RTBB_EXTHSI2C_SetHSCode(IntPtr hDevice, byte[] pCode, uint nLength);

        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_EXTHSI2C_I2CWrite")]
        private static extern int native_RTBB_EXTHSI2C_I2CWrite(IntPtr hDevice, int nBus, int nSlaveAddr, int nCmdSize, int nCmd, int nBufferSize, byte[] pBuffer);

        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_EXTHSI2C_I2CRead")]
        private static extern int native_RTBB_EXTHSI2C_I2CRead(IntPtr hDevice, int nBus, int nSlaveAddr, int nCmdSize, int nCmd, int nBufferSize, byte[] pBuffer);
    }
}
