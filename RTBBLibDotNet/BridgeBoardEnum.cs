using System;
using System.Runtime.InteropServices;

namespace RTBBLibDotNet
{
    interface IBoardEnumControl
    {
        int RTBB_GetBoardCount();
        int RTBB_GetEnumBoardInfo(int nIndex, ref GlobalVariable.RTBBInfo BBInfo);
    }

    public class BridgeBoardEnum : GlobalVariable, IBoardEnumControl
    {
        IntPtr hEnum = IntPtr.Zero;

        private BridgeBoardEnum()
        {
            hEnum = native_RTBB_EnumBoard();
        }

        ~BridgeBoardEnum()
        {
            if (hEnum != IntPtr.Zero)
                native_RTBB_FreeEnumBoard(hEnum);
        }

        ///<summary>
        ///Description: This function will return a enum handle.
        ///If the function succeeds, the return value is enum handle.
        ///Else, return IntPtr.Zero.
        ///</summary>
        public IntPtr GetEnumHandle()
        {
            return hEnum;
        }

        ///<summary>
        ///Description: This function will return the bridgeboard count.
        ///If the function succeeds, the return value is the bridgeboard count.
        ///</summary>
        public int RTBB_GetBoardCount()
        {
            return native_RTBB_GetBoardCount(hEnum);
        }

        ///<summary>
        ///Description: This function will return the bridgeboard info.
        ///Input parameter: nIndex -> specified bridgeboard index.
        ///Output parameter: BBInfo -> BridgeBoard info that returned.
        ///If the function succeeds, the return value is zero.
        ///</summary>
        public int RTBB_GetEnumBoardInfo(int nIndex, ref RTBBInfo BBInfo)
        {
            return trans_RTBB_GetEnumBoardInfo(hEnum, nIndex, ref BBInfo);
        }

        ///<summary>
        ///Description: This function will Bridgeboard Enum class instance.
        ///If the function succeeds, the return value is Bridgeboard Enum class instance.
        ///Else, fail and return value is null.
        ///</summary>
        public static BridgeBoardEnum GetBoardEnum()
        {
            BridgeBoardEnum boardEnum = new BridgeBoardEnum();
            if (boardEnum.hEnum == IntPtr.Zero)
                boardEnum = null;
            return boardEnum;
        }

        ///<summary>
        ///Return value will be the enum handle.
        ///</summary>
        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_EnumBoard")]
        private static extern IntPtr native_RTBB_EnumBoard();

        ///<summary>
        ///Input Parameters: hEnumHandle -> Handle to a enumeration of bridge board which get from RTBB_EnumBoard().
        ///Free the enum board handle.
        ///</summary>
        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_FreeEnumBoard")]
        private static extern bool native_RTBB_FreeEnumBoard(IntPtr hEnumHandle);

        ///<summary>
        ///Input Parameters: hEnumHandle -> Handle to a enumeration of bridge board which get from RTBB_EnumBoard().
        ///Return value is the enumerated board count.
        ///</summary>
        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_GetBoardCount")]
        private static extern int native_RTBB_GetBoardCount(IntPtr hEnumHandle);


        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_GetEnumBoardInfo")]
        private static extern IntPtr native_RTBB_GetEnumBoardInfo(IntPtr hEnumHandle, int nIndex);

        ///<summary>
        ///Input Parameters: hEnumHandle -> Handle to a enumeration of bridge board which get from RTBB_EnumBoard().
        ///Input Parameters: nIndex -> Index of bridge board. Start number is 0.
        ///Output Parameters: BBInfo -> if the return value is 0, function success, BBInfo will be filled into the structure.
        ///                          -> else return value is -1, means BBInfo cannot be used.
        ///Return value, if success, return value is 0. Else, -1.
        ///</summary>
        private int trans_RTBB_GetEnumBoardInfo(IntPtr hEnumHandle, int nIndex, ref RTBBInfo BBInfo)
        {
            IntPtr infoPtr = IntPtr.Zero;
            infoPtr = native_RTBB_GetEnumBoardInfo(hEnumHandle, nIndex);
            if (infoPtr == IntPtr.Zero)
                return -1;
            BBInfo = (RTBBInfo)Marshal.PtrToStructure(infoPtr, typeof(RTBBInfo));
            return 0;
        }
    }
}
