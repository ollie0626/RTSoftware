using System;
using System.Runtime.InteropServices;

namespace RTBBLibDotNet
{
    public interface IExtStorageModule : IBaseModule
    {
        int RTBB_EXTSTORAGE_GetBankNR();
        int RTBB_EXTSTORAGE_GetBankSize();
        int RTBB_EXTSTORAGE_BankRead(int nBank, int nOffset, int nSize, byte[] pDest);
        int RTBB_EXTSTORAGE_BankWrite(int nBank, int nOffset, int nSize, byte[] pSrc);
        int RTBB_EXTSTORAGE_Flush();
    }

    public class ExtStorageModule : GlobalVariable, IExtStorageModule
    {
        private IntPtr hDev = IntPtr.Zero;

        public ExtStorageModule(IntPtr hDevice)
        {
            hDev = hDevice;
        }

        ///<summary>
        ///Description: return the module name.
        ///If the function succeeds, the return value is the module name
        ///</summary>
        public string getModuleName()
        {
            return "ExtStorage";
        }

        ///<summary>
        ///Description: ExtStorage get bank number.
        ///If the function succeeds, the return value is positive and it means the total amount of data bank.
        ///If the function fails, the return value is negative.
        ///To get result description string, call RTBB_Result2String().
        ///</summary>
        public int RTBB_EXTSTORAGE_GetBankNR()
        {
            return native_RTBB_EXTSTORAGE_GetBankNR(hDev);
        }

        ///<summary>
        ///Description: ExtStorage get bank size.
        ///If the function succeeds, the return value is positive and it means the bank size (bytes).
        ///If the function fails, the return value is negative.
        ///To get result description string, call RTBB_Result2String().
        ///</summary>
        public int RTBB_EXTSTORAGE_GetBankSize()
        {
            return native_RTBB_EXTSTORAGE_GetBankSize(hDev);
        }

        ///<summary>
        ///Description: ExtStorage bank Read.
        ///Input Parameters: nBank -> Storage bank index, start from 0.
        ///Input Parameters: nOffset -> Offset bytes from each nBank start.
        ///Input Parameters: nSize -> Read bytes size from nOffset.
        ///Input Parameters: pDest -> Pointer to the data buffer to be read from storage device.
        ///If the function succeeds, the return value is zero.
        ///If the function fails, the return value is negative.
        ///To get result description string, call RTBB_Result2String().
        ///</summary>
        public int RTBB_EXTSTORAGE_BankRead(int nBank, int nOffset, int nSize, byte[] pDest)
        {
            return native_RTBB_EXTSTORAGE_BankRead(hDev, nBank, nOffset, nSize, pDest);
        }

        ///<summary>
        ///Description: ExtStorage bank Write.
        ///Input Parameters: nBank -> Storage bank index, start from 0.
        ///Input Parameters: nOffset -> Offset bytes from each nBank start.
        ///Input Parameters: nSize -> Read bytes size from nOffset.
        ///Input Parameters: pDest -> Pointer to the data buffer to be read from storage device.
        ///If the function succeeds, the return value is zero.
        ///If the function fails, the return value is negative.
        ///To get result description string, call RTBB_Result2String().
        ///</summary>
        public int RTBB_EXTSTORAGE_BankWrite(int nBank, int nOffset, int nSize, byte[] pSrc)
        {
            return native_RTBB_EXTSTORAGE_BankWrite(hDev, nBank, nOffset, nSize, pSrc);
        }

        ///<summary>
        ///Description: ExtStorage flush function.
        ///If the function succeeds, the return value is zero.
        ///If the function fails, the return value is negative.
        ///To get result description string, call RTBB_Result2String().
        ///</summary>
        public int RTBB_EXTSTORAGE_Flush()
        {
            return native_RTBB_EXTSTORAGE_Flush(hDev);
        }

        /* ExtStorage Control Functions */
        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_EXTSTORAGE_GetBankNR")]
        private static extern int native_RTBB_EXTSTORAGE_GetBankNR(IntPtr hDevice);

        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_EXTSTORAGE_GetBankSize")]
        private static extern int native_RTBB_EXTSTORAGE_GetBankSize(IntPtr hDevice);

        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_EXTSTORAGE_BankRead")]
        private static extern int native_RTBB_EXTSTORAGE_BankRead(IntPtr hDevice, int nBank, int nOffset, int nSize, byte[] pDest);

        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_EXTSTORAGE_BankWrite")]
        private static extern int native_RTBB_EXTSTORAGE_BankWrite(IntPtr hDevice, int nBank, int nOffset, int nSize, byte[] pSrc);

        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_EXTSTORAGE_Flush")]
        private static extern int native_RTBB_EXTSTORAGE_Flush(IntPtr hDevice);
    }
}
