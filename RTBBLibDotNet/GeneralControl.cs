using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Management;

namespace RTBBLibDotNet
{
    public partial class BridgeBoard
    {
        private ManagementEventWatcher mEventWatcher = null;

        ///<summary>
        ///Description: RTBB firmware check.
        ///Input/Output Parameters: pMinVerBCD -> the minimum bcd version value.
        ///Output Parameters: pCurrentBCD -> the current bcd version value.
        ///Check the firmware version.
        ///if yes, return true, otherwise, false.
        ///</summary>
        public bool RTBB_FirmwareCheck(ref uint pMinVerBCD, ref uint pCurrentBCD)
        {
            return native_RTBB_FirmwareCheck(hDev, ref pMinVerBCD, ref pCurrentBCD);
        }

        private void createModuleByCapability()
        {
            moduleList = new Dictionary<string, IBaseModule>();
            if ((BBInfo.nCapability & RT_BRIDGE_I2C) != 0)
            {
                for (int i = 0; i < RTBB_I2CGetBusCount(this.hDev); i++)
                    moduleList.Add("I2C" + i.ToString(), new I2CModule(this.hDev, i));
            }
            if ((BBInfo.nCapability & RT_BRIDGE_GPIO) != 0)
                moduleList.Add("GPIO", new GPIOModule(this.hDev));
            if ((BBInfo.nCapability & RT_BRIDGE_GPIO_EXT) != 0)
                moduleList.Add("GPIOExt", new GPIOExtModule(this.hDev));
            if ((BBInfo.nCapability & RT_BRIDGE_PWM) != 0)
                moduleList.Add("PWM", new PWMModule(this.hDev));
            if ((BBInfo.nCapability & RT_BRIDGE_SPI) != 0)
            {
                for (int i = 0; i < RTBB_SPIGetBusCount(this.hDev); i++)
                    moduleList.Add("SPI" + i.ToString(), new SPIModule(this.hDev, i));
            }
            if ((BBInfo.nCapability & RT_BRIDGE_EXTENDED) != 0)
            {
                moduleList.Add("ExtCustomizedCommand", new ExtCustomizedCommandModule(this.hDev));
                moduleList.Add("ExtGPIOMisc", new ExtGPIOMiscMdoule(this.hDev));
                moduleList.Add("ExtGSMW", new ExtGSMWModule(this.hDev));
                moduleList.Add("ExtGSOW", new ExtGSOWModule(this.hDev));
                moduleList.Add("ExtHSI2C", new ExtHSI2CModule(this.hDev));
                moduleList.Add("ExtIOConfig", new ExtIOConfigModule(this.hDev));
                moduleList.Add("ExtSecurityData", new ExtSecurityDataModule(this.hDev));
                moduleList.Add("ExtStorage", new ExtStorageModule(this.hDev));
                moduleList.Add("ExtSVI2C", new ExtSVI2CModule(this.hDev));

                moduleList.Add("ExtQSPI", new ExtQSPIModule(this.hDev));
            }
        }

        private void OnDeviceRemovalHandler(object Sender, EventArrivedEventArgs e)
        {
            DeviceRemovalHandler(this, null);
        }

        private void CreateWMIEventWatcher()
        {
            mEventWatcher = new ManagementEventWatcher("root\\cimv2", GetWMIQueryString());
            mEventWatcher.EventArrived += new EventArrivedEventHandler(OnDeviceRemovalHandler);
            mEventWatcher.Start();
        }

        private string GetWMIQueryString()
        {
            string query = "Select * From __InstanceDeletionEvent Within 1 Where TargetInstance Isa \"Win32_PnPEntity\" And ";
            string[] tmp = this.BBInfo.strDevicePath.Split('#');
            query += "TargetInstance.DeviceID Like \"%";
            if (tmp.Length >= 4)
                query += tmp[2].ToUpper();
            query += "%\"";
            return query;
        }

        /* I2C/SPI Bus Count Function */
        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_I2CGetBusCount")]
        private static extern int RTBB_I2CGetBusCount(IntPtr hDevice);

        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_SPIGetBusCount")]
        private static extern int RTBB_SPIGetBusCount(IntPtr hDevice);

        /* BridgeBoard Connect/Discoonnect Functions */
        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true)]
        private static extern IntPtr RTBB_ConnectToBridgeByIndex(int nIndex);

        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true)]
        private static extern IntPtr RTBB_ConnectToBridgeByInfo(IntPtr hEnumHandle, ref RTBBInfo pDeviceInfo);

        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true)]
        private static extern IntPtr RTBB_ConnectToBridgeByCapability(int nindex, UInt32 nCapability);

        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true)]
        private static extern bool RTBB_DisconnectBridge(IntPtr hDevice);

        [DllImport(dll_path + "RTBBLib.dll", SetLastError = true, EntryPoint = "RTBB_FirmwareCheck")]
        private static extern bool native_RTBB_FirmwareCheck(IntPtr hDevice, ref uint pMinVerBCD, ref uint pCurrentBCD);
    }
}
