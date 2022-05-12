using System;
using System.Runtime.InteropServices;

namespace RTBBLibDotNet
{
    public interface IPWMModule : IBaseModule
    {
        int RTBB_PWMGetPWMPortCount();
        int RTBB_PWMStart(int nPort);
        int RTBB_PWMStop(int nPort);
        int RTBB_PWMGetStatus(int nPort);
        int RTBB_PWMSetDutyCycle(int nPort, double fDutyCycle);
        int RTBB_PWMGetDutyCycle(int nPort, ref double fDutyCycle);
        int RTBB_PWMSetPeriod(int nPort, uint nTick);
        int RTBB_PWMSetOnPeriod(int nPort, uint nOnTick);
        int RTBB_PWMGetPeriod(int nPort, ref uint pnTick);
        int RTBB_PWMGetOnPeriod(int nPort, ref uint pnOnTick);
        int RTBB_PWMSetPeriodByUs(int nPort, uint nUs);
        int RTBB_PWMSetOnPeriodByUs(int nPort, uint nUs);
        int RTBB_PWMSetPeriodByMs(int nPort, uint nMs);
        int RTBB_PWMSetOnPeriodByMs(int nPort, uint nMs);
        int RTBB_PWMGetPeriodByUs(int nPort, ref uint pnUs);
        int RTBB_PWMGetOnPeriodByUs(int nPort, ref uint pnUs);
    }

    public class PWMModule : GlobalVariable, IPWMModule
    {
        private IntPtr hDev = IntPtr.Zero;

        public PWMModule(IntPtr hDevice)
        {
            hDev = hDevice;
        }

        ///<summary>
        ///Description: return the module name.
        ///If the function succeeds, the return value is the module name
        ///</summary>
        public string getModuleName()
        {
            return "PWM";
        }

        ///<summary>
        ///Description: Get PWM Port Count.
        ///If the function succeeds, the return value is the quantities of available PWM port on Bridgeboard.
        ///If the function fails, the return value is negative.
        ///To get result description string, call RTBB_Result2String().
        ///</summary>
        public int RTBB_PWMGetPWMPortCount()
        {
            return native_RTBB_PWMGetPWMPortCount(hDev);
        }

        ///<summary>
        ///Description: PWM Start.
        ///Input Parameters: nPort -> Index of PWM port. Start number is 0.
        ///If the function succeeds, the return value is zero.
        ///If the function fails, the return value is negative.
        ///To get result description string, call RTBB_Result2String().
        ///</summary>
        public int RTBB_PWMStart(int nPort)
        {
            return native_RTBB_PWMStart(hDev, nPort);
        }

        ///<summary>
        ///Description: PWM Stop.
        ///Input Parameters: nPort -> Index of PWM port. Start number is 0.
        ///If the function succeeds, the return value is zero.
        ///If the function fails, the return value is negative.
        ///To get result description string, call RTBB_Result2String().
        ///</summary>
        public int RTBB_PWMStop(int nPort)
        {
            return native_RTBB_PWMStop(hDev, nPort);
        }

        ///<summary>
        ///Description: PWM Get Status.
        ///Input Parameters: nPort -> Index of PWM port. Start number is 0.
        ///If the function succeeds, the return value will be
        ///1: PWM status is Start, or
        ///2: PWM status is Stop
        ///If the function fails, the return value is negative.
        ///To get result description string, call RTBB_Result2String().
        ///</summary>
        public int RTBB_PWMGetStatus(int nPort)
        {
            return native_RTBB_PWMGetStatus(hDev, nPort);
        }

        ///<summary>
        ///Description: PWM Set Duty Cycle.
        ///Input Parameters: nPort -> Index of PWM port. Start number is 0.
        ///Input Parameters: fDutyCycle -> Current duty cycle of PWM port.
        ///                             -> It is double value, 1 means duty cycle = 100%.
        ///If the function succeeds, the return value is zero.
        ///If the function fails, the return value is negative.
        ///To get result description string, call RTBB_Result2String().
        ///</summary>
        public int RTBB_PWMSetDutyCycle(int nPort, double fDutyCycle)
        {
            return native_RTBB_PWMSetDutyCycle(hDev, nPort, fDutyCycle);
        }

        ///<summary>
        ///Description: PWM Get Duty Cycle.
        ///Input Parameters: nPort -> Index of PWM port. Start number is 0.
        ///Output Parameters: fDutyCycle -> Current duty cycle of PWM port.
        ///                             -> It is double value, 1 means duty cycle = 100%.
        ///If the function succeeds, the return value is zero.
        ///If the function fails, the return value is negative.
        ///To get result description string, call RTBB_Result2String().
        ///</summary>
        public int RTBB_PWMGetDutyCycle(int nPort, ref double fDutyCycle)
        {
            return native_RTBB_PWMGetDutyCycle(hDev, nPort, ref fDutyCycle);
        }

        ///<summary>
        ///Description: PWM Set Period.
        ///Input Parameters: nPort -> Index of PWM port. Start number is 0.
        ///Input Parameters: nTick -> Current period of PWM port, unit is the quantities of Base Clock.
        ///If the function succeeds, the return value is zero.
        ///If the function fails, the return value is negative.
        ///To get result description string, call RTBB_Result2String().
        ///</summary>
        public int RTBB_PWMSetPeriod(int nPort, uint nTick)
        {
            return native_RTBB_PWMSetPeriod(hDev, nPort, nTick);
        }

        ///<summary>
        ///Description: PWM Set On Period.
        ///Input Parameters: nPort -> Index of PWM port. Start number is 0.
        ///Input Parameters: nOnTick -> Current ON period of PWM port, unit is the quantities of Base Clock.
        ///If the function succeeds, the return value is zero.
        ///If the function fails, the return value is negative.
        ///To get result description string, call RTBB_Result2String().
        ///</summary>
        public int RTBB_PWMSetOnPeriod(int nPort, uint nOnTick)
        {
            return native_RTBB_PWMSetOnPeriod(hDev, nPort, nOnTick);
        }

        ///<summary>
        ///Description: PWM Get Period.
        ///Input Parameters: nPort -> Index of PWM port. Start number is 0.
        ///Output Parameters: pnTick -> Current period of PWM port, unit is the quantities of Base Clock.
        ///If the function succeeds, the return value is zero.
        ///If the function fails, the return value is negative.
        ///To get result description string, call RTBB_Result2String().
        ///</summary>
        public int RTBB_PWMGetPeriod(int nPort, ref uint pnTick)
        {
            return native_RTBB_PWMGetPeriod(hDev, nPort, ref pnTick);
        }

        ///<summary>
        ///Description: PWM Get On Period.
        ///Input Parameters: nPort -> Index of PWM port. Start number is 0.
        ///Output Parameters: pnOnTick -> Current ON period of PWM port, unit is the quantities of Base Clock.
        ///If the function succeeds, the return value is zero.
        ///If the function fails, the return value is negative.
        ///To get result description string, call RTBB_Result2String().
        ///</summary>
        public int RTBB_PWMGetOnPeriod(int nPort, ref uint pnOnTick)
        {
            return native_RTBB_PWMGetOnPeriod(hDev, nPort, ref pnOnTick);
        }

        ///<summary>
        ///Description: PWM Set Period by uS.
        ///Input Parameters: nPort -> Index of PWM port. Start number is 0.
        ///Input Parameters: nUs -> Current period of PWM port, unit is us. 
        ///If the function succeeds, the return value is zero.
        ///If the function fails, the return value is negative.
        ///To get result description string, call RTBB_Result2String().
        ///</summary>
        public int RTBB_PWMSetPeriodByUs(int nPort, uint nUs)
        {
            return native_RTBB_PWMSetPeriodByUs(hDev, nPort, nUs);
        }

        ///<summary>
        ///Description: PWM Set On Period by uS.
        ///Input Parameters: nPort -> Index of PWM port. Start number is 0.
        ///Input Parameters: nUs -> Current ON period of PWM port, unit is us. 
        ///If the function succeeds, the return value is zero.
        ///If the function fails, the return value is negative.
        ///To get result description string, call RTBB_Result2String().
        ///</summary>
        public int RTBB_PWMSetOnPeriodByUs(int nPort, uint nUs)
        {
            return native_RTBB_PWMSetOnPeriodByUs(hDev, nPort, nUs);
        }

        ///<summary>
        ///Description: PWM Set Period by mS.
        ///Input Parameters: nPort -> Index of PWM port. Start number is 0.
        ///Input Parameters: nMs -> Current period of PWM port, unit is ms. 
        ///If the function succeeds, the return value is zero.
        ///If the function fails, the return value is negative.
        ///To get result description string, call RTBB_Result2String().
        ///</summary>
        public int RTBB_PWMSetPeriodByMs(int nPort, uint nMs)
        {
            return native_RTBB_PWMSetPeriodByMs(hDev, nPort, nMs);
        }

        ///<summary>
        ///Description: PWM Set On Period by mS.
        ///Input Parameters: nPort -> Index of PWM port. Start number is 0.
        ///Input Parameters: nMs -> Current ON period of PWM port, unit is ms. 
        ///If the function succeeds, the return value is zero.
        ///If the function fails, the return value is negative.
        ///To get result description string, call RTBB_Result2String().
        ///</summary>
        public int RTBB_PWMSetOnPeriodByMs(int nPort, uint nMs)
        {
            return native_RTBB_PWMSetOnPeriodByMs(hDev, nPort, nMs);
        }

        ///<summary>
        ///Description: PWM Get Period by uS.
        ///Input Parameters: nPort -> Index of PWM port. Start number is 0.
        ///Output Parameters: pnUs -> Current period of PWM port, unit is us. 
        ///If the function succeeds, the return value is zero.
        ///If the function fails, the return value is negative.
        ///To get result description string, call RTBB_Result2String().
        ///</summary>
        public int RTBB_PWMGetPeriodByUs(int nPort, ref uint pnUs)
        {
            return native_RTBB_PWMGetPeriodByUs(hDev, nPort, ref pnUs);
        }

        ///<summary>
        ///Description: PWM Get On Period by uS.
        ///Input Parameters: nPort -> Index of PWM port. Start number is 0.
        ///Output Parameters: pnUs -> Current ON period of PWM port, unit is us. 
        ///If the function succeeds, the return value is zero.
        ///If the function fails, the return value is negative.
        ///To get result description string, call RTBB_Result2String().
        ///</summary>
        public int RTBB_PWMGetOnPeriodByUs(int nPort, ref uint pnUs)
        {
            return native_RTBB_PWMGetOnPeriodByUs(hDev, nPort, ref pnUs);
        }

        /* PWM control Functions */
        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_PWMGetPWMPortCount")]
        private static extern int native_RTBB_PWMGetPWMPortCount(IntPtr hDevice);

        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_PWMStart")]
        private static extern int native_RTBB_PWMStart(IntPtr hDevice, int nPort);

        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_PWMStop")]
        private static extern int native_RTBB_PWMStop(IntPtr hDevice, int nPort);

        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_PWMGetStatus")]
        private static extern int native_RTBB_PWMGetStatus(IntPtr hDevice, int nPort);

        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_PWMSetDutyCycle")]
        private static extern int native_RTBB_PWMSetDutyCycle(IntPtr hDevice, int nPort, double fDutyCycle);

        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_PWMGetDutyCycle")]
        private static extern int native_RTBB_PWMGetDutyCycle(IntPtr hDevice, int nPort, ref double fDutyCycle);

        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_PWMSetPeriod")]
        private static extern int native_RTBB_PWMSetPeriod(IntPtr hDevice, int nPort, uint nTick);

        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_PWMSetOnPeriod")]
        private static extern int native_RTBB_PWMSetOnPeriod(IntPtr hDevice, int nPort, uint nOnTick);

        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_PWMGetPeriod")]
        private static extern int native_RTBB_PWMGetPeriod(IntPtr hDevice, int nPort, ref uint pnTick);

        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_PWMGetOnPeriod")]
        private static extern int native_RTBB_PWMGetOnPeriod(IntPtr hDevice, int nPort, ref uint pnOnTick);

        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_PWMSetPeriodByUs")]
        private static extern int native_RTBB_PWMSetPeriodByUs(IntPtr hDevice, int nPort, uint nUs);

        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_PWMSetOnPeriodByUs")]
        private static extern int native_RTBB_PWMSetOnPeriodByUs(IntPtr hDevice, int nPort, uint nUs);

        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_PWMSetPeriodByMs")]
        private static extern int native_RTBB_PWMSetPeriodByMs(IntPtr hDevice, int nPort, uint nMs);

        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_PWMSetOnPeriodByMs")]
        private static extern int native_RTBB_PWMSetOnPeriodByMs(IntPtr hDevice, int nPort, uint nMs);

        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_PWMGetPeriodByUs")]
        private static extern int native_RTBB_PWMGetPeriodByUs(IntPtr hDevice, int nPort, ref uint pnUs);

        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_PWMGetOnPeriodByUs")]
        private static extern int native_RTBB_PWMGetOnPeriodByUs(IntPtr hDevice, int nPort, ref uint pnUs);
    }
}
