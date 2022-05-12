using System;
using System.Runtime.InteropServices;

namespace RTBBLibDotNet
{
    public interface IExtQSPIModule : IBaseModule
    {
        int RTBB_EXTQSPI_GetFrequency(ref uint pFreqkHz);

        int RTBB_EXTQSPI_SetFrequency(int nFreqkHz);

        int RTBB_SPIGetBusCount(int busIndex);

        int RTBB_EXTQSPI_SetActiveLow(bool bActiveLow);

        int RTBB_EXTQSPI_HLRead(GlobalVariable.ERTQSPIFiledMode nFieldMode, byte nOpcodeSize, byte nOpcode, byte nAddrSize, int nAddr, byte nIDataSize, int nIData, byte nDummyCyc, int nDataLength, byte[] pData);

        int RTBB_EXTQSPI_HLWrite(GlobalVariable.ERTQSPIFiledMode nFieldMode, byte nOpcodeSize, byte nOpcode, byte nAddrSize, int nAddr, byte nIDataSize, int nIData, byte nDummyCyc, int nDataLength, byte[] pData);

        int RTBB_EXTQSPI_HLReadCS(GlobalVariable.ERTQSPIFiledMode nFieldMode, int nPinNumber, byte nOpcodeSize, byte nOpcode, byte nAddrSize, int nAddr, byte nIDataSize, int nIData, byte nDummyCyc, int nDataLength, byte[] pData);

        int RTBB_EXTQSPI_HLWriteCS(GlobalVariable.ERTQSPIFiledMode nFieldMode, int nPinNumber, byte nOpcodeSize, byte nOpcode, byte nAddrSize, int nAddr, byte nIDataSize, int nIData, byte nDummyCyc, int nDataLength, byte[] pData);
    }


    public class ExtQSPIModule : GlobalVariable, IExtQSPIModule
    {
        private IntPtr hDev = IntPtr.Zero;
        private int mBusIndex = 0;

        public ExtQSPIModule(IntPtr hDevice)
        {
            hDev = hDevice;
        }

        public ExtQSPIModule(IntPtr hDevice, int busIndex)
        {
            hDev = hDevice;
            mBusIndex = busIndex;
        }

        public string getModuleName()
        {
            return "ExtQSPI" + mBusIndex.ToString();
        }

        public int RTBB_EXTQSPI_SetFrequency(int nFreqkHz)
        {
            return native_RTBB_EXTQSPI_SetFrequency(hDev, mBusIndex, nFreqkHz);
        }

        public int RTBB_EXTQSPI_GetFrequency(ref uint pFreqkHz)
        {
            return native_RTBB_EXTQSPI_GetFrequency(hDev, mBusIndex, ref pFreqkHz);
        }

        public int RTBB_SPIGetBusCount(int busIndex)
        {
            return native_RTBB_EXTQSPI_GetBusCount(hDev, busIndex);
        }

        public int RTBB_EXTQSPI_SetActiveLow(bool bActiveLow)
        {
            return native_RTBB_EXTQSPI_SetActiveLow(hDev, bActiveLow);
        }

        public int RTBB_EXTQSPI_HLRead(ERTQSPIFiledMode nFieldMode, byte nOpcodeSize, byte nOpcode, byte nAddrSize, int nAddr, byte nIDataSize, int nIData, byte nDummyCyc, int nDataLength,byte[] pData)
        {
            return native_RTBB_EXTQSPI_HLRead(hDev,
                                                mBusIndex,
                                                nFieldMode,
                                                nOpcodeSize,
                                                nOpcode,
                                                nAddrSize,
                                                nAddr,
                                                nIDataSize,
                                                nIData,
                                                nDummyCyc,
                                                nDataLength,
                                                pData);
        }

        public int RTBB_EXTQSPI_HLWrite(ERTQSPIFiledMode nFieldMode, byte nOpcodeSize, byte nOpcode, byte nAddrSize, int nAddr, byte nIDataSize, int nIData, byte nDummyCyc, int nDataLength, byte[] pData)
        {
            return native_RTBB_EXTQSPI_HLWrite(hDev,
                                                mBusIndex,
                                                nFieldMode,
                                                nOpcodeSize,
                                                nOpcode,
                                                nAddrSize,
                                                nAddr,
                                                nIDataSize,
                                                nIData,
                                                nDummyCyc,
                                                nDataLength,
                                                pData);
        }

        public int RTBB_EXTQSPI_HLReadCS(ERTQSPIFiledMode nFieldMode, int nPinNumber, byte nOpcodeSize, byte nOpcode, byte nAddrSize, int nAddr, byte nIDataSize, int nIData, byte nDummyCyc, int nDataLength, byte[] pData)
        {
            return native_RTBB_EXTQSPI_HLReadCS(hDev,
                                                mBusIndex,
                                                nFieldMode,
                                                nPinNumber,
                                                nOpcodeSize,
                                                nOpcode,
                                                nAddrSize,
                                                nAddr,
                                                nIDataSize,
                                                nIData,
                                                nDummyCyc,
                                                nDataLength,
                                                pData);
        }

        public int RTBB_EXTQSPI_HLWriteCS(ERTQSPIFiledMode nFieldMode, int nPinNumber, byte nOpcodeSize, byte nOpcode, byte nAddrSize, int nAddr, byte nIDataSize, int nIData, byte nDummyCyc, int nDataLength, byte[] pData)
        {
            
            return native_RTBB_EXTQSPI_HLWriteCS(hDev,
                                                mBusIndex,
                                                nFieldMode,
                                                nPinNumber,
                                                nOpcodeSize,
                                                nOpcode,
                                                nAddrSize,
                                                nAddr,
                                                nIDataSize,
                                                nIData,
                                                nDummyCyc,
                                                nDataLength,
                                                pData);
            
        }



        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_EXTQSPI_GetBusCount")]
        private static extern int native_RTBB_EXTQSPI_GetBusCount(IntPtr hDevice, int nBus);

        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_EXTQSPI_SetActiveLow")]
        private static extern int native_RTBB_EXTQSPI_SetActiveLow(IntPtr hDevice, bool bActiveLow);


        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_EXTQSPI_SetFrequency")]
        private static extern int native_RTBB_EXTQSPI_SetFrequency(IntPtr hDevice, int nBus, int nFreqkHz);


        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_EXTQSPI_GetFrequency")]
        private static extern int native_RTBB_EXTQSPI_GetFrequency(IntPtr hDevice, int nBus, ref uint nFreqkHz);

        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_EXTQSPI_HLRead")]
        private static extern int native_RTBB_EXTQSPI_HLRead(   IntPtr hDevice, int nBus,
                                                                ERTQSPIFiledMode nFieldMode, byte nOpcodeSize, byte nOpcode,
                                                                byte nAddrSize, int nAddr,
                                                                byte nIDataSize, int nIData,
                                                                byte nDummyCyc, int nDataLength,
                                                                byte[] pData);

        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_EXTQSPI_HLWrite")]
        private static extern int native_RTBB_EXTQSPI_HLWrite(  IntPtr hDevice, int nBus,
                                                                ERTQSPIFiledMode nFieldMode, byte nOpcodeSize, byte nOpcode,
                                                                byte nAddrSize, int nAddr,
                                                                byte nIDataSize, int nIData,
                                                                byte nDummyCyc, int nDataLength,
                                                                byte[] pData);

        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_EXTQSPI_HLReadCS")]
        private static extern int native_RTBB_EXTQSPI_HLReadCS( IntPtr hDevice, int nBus,
                                                                ERTQSPIFiledMode nFieldMode, int nPinNumber,
                                                                byte nOpcodeSize, byte nOpcode,
                                                                byte nAddrSize, int nAddr,
                                                                byte nIDataSize, int nIData,
                                                                byte nDummyCyc, int nDataLength,
                                                                byte[] pData);

        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_EXTQSPI_HLWriteCS")]
        private static extern int native_RTBB_EXTQSPI_HLWriteCS(    IntPtr hDevice, int nBus,
                                                                    ERTQSPIFiledMode nFieldMode, int nPinNumber,
                                                                    byte nOpcodeSize, byte nOpcode,
                                                                    byte nAddrSize, int nAddr,
                                                                    byte nIDataSize, int nIData,
                                                                    byte nDummyCyc, int nDataLength,
                                                                    byte[] pData);

    }


}
