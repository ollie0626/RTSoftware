using System;
using System.Collections.Generic;

namespace RTBBLibDotNet
{
    public partial class BridgeBoard : GlobalVariable
    {
        private IntPtr hEnum = IntPtr.Zero;
        private IntPtr hDev = IntPtr.Zero;
        private int nIndex = 0;
        private RTBBInfo BBInfo = new RTBBInfo();
        private Dictionary<string, IBaseModule> moduleList;
        public event EventHandler DeviceRemovalHandler;
        
        protected void OnDeviceRemovalHandle(EventArgs e)
        {
            EventHandler handler = DeviceRemovalHandler;
            if (handler != null)
            {
                handler(this, e);
            }
        }

        private BridgeBoard(BridgeBoardEnum hEum)
        {
            this.hEnum = hEum.GetEnumHandle();
            hDev = RTBB_ConnectToBridgeByIndex(this.nIndex);
            if (hDev != IntPtr.Zero)
            {
                hEum.RTBB_GetEnumBoardInfo(this.nIndex, ref BBInfo);
                createModuleByCapability();
                //CreateWMIEventWatcher();
            }
        }

        private BridgeBoard(BridgeBoardEnum hEum, int nIndex)
        {
            this.hEnum = hEum.GetEnumHandle();
            this.nIndex = nIndex;
            hDev = RTBB_ConnectToBridgeByIndex(this.nIndex);
            if (hDev != IntPtr.Zero)
            {
                hEum.RTBB_GetEnumBoardInfo(this.nIndex, ref BBInfo);
                createModuleByCapability();
                //CreateWMIEventWatcher();
            }
        }

        private BridgeBoard(BridgeBoardEnum hEnum, RTBBInfo BBInfo)
        {
            this.hEnum = hEnum.GetEnumHandle();
            hDev = RTBB_ConnectToBridgeByInfo(this.hEnum, ref BBInfo);
            if (hDev != IntPtr.Zero)
            {
                this.BBInfo = BBInfo;
                createModuleByCapability();
                //CreateWMIEventWatcher();
            }
        }

        private BridgeBoard(BridgeBoardEnum hEnum, int nIndex, UInt32 capability)
        {
            this.hEnum = hEnum.GetEnumHandle();
            this.nIndex = nIndex;
            hDev = RTBB_ConnectToBridgeByCapability(nIndex, capability);
            if (hDev != IntPtr.Zero)
            {
                hEnum.RTBB_GetEnumBoardInfo(this.nIndex, ref BBInfo);
                createModuleByCapability();
                //CreateWMIEventWatcher();
            }
        }

        ~BridgeBoard()
        {
            if (this.hDev != IntPtr.Zero)
                RTBB_DisconnectBridge(this.hDev);
        }

        ///<summary>
        ///Description: This function will return a BridgeBoard Instance.
        ///Input Parameters: hEnum -> Please input a BridgeBoardEnum.
        ///If the function succeeds, the return object is BridgeBoard instance.
        ///Else, return null.
        ///</summary>
        public static BridgeBoard ConnectByDefault(BridgeBoardEnum hEnum)
        {
            BridgeBoard b = new BridgeBoard(hEnum);
            if (b.hDev == IntPtr.Zero)
                b = null;
            return b;
        }

        ///<summary>
        ///Description: This function will return a BridgeBoard Instance.
        ///Input Parameters: hEnum -> Please input a BridgeBoardEnum.
        ///Input Parameters: nIndex -> Please specified the index that the bridgeboard you want to connect.
        ///If the function succeeds, the return object is BridgeBoard instance.
        ///Else, return null.
        ///</summary>
        public static BridgeBoard ConnectByIndex(BridgeBoardEnum hEnum, int nIndex)
        {
            BridgeBoard b = new BridgeBoard(hEnum, nIndex);
            if (b.hDev == IntPtr.Zero)
                b = null;
            return b;
        }

        ///<summary>
        ///Description: This function will return a BridgeBoard Instance.
        ///Input Parameters: hEnum -> Please input a BridgeBoardEnum.
        ///Input Parameters: BBInfo -> Connect the bridgeboard by the specified RTBBInfo.
        ///If the function succeeds, the return object is BridgeBoard instance.
        ///Else, return null.
        ///</summary>
        public static BridgeBoard ConnectByBBInfo(BridgeBoardEnum hEnum, RTBBInfo BBInfo)
        {
            BridgeBoard b = new BridgeBoard(hEnum, BBInfo);
            if (b.hDev == IntPtr.Zero)
                b = null;
            return b;
        }

        ///<summary>
        ///Description: This function will return a BridgeBoard Instance.
        ///Input Parameters: hEnum -> Please input a BridgeBoardEnum.
        ///Input Parameters: nIndex -> Please specified the index that the bridgeboard you want to connect.
        ///Input Parameters: capability -> Bridgeboard capability that you to connect.
        ///If the function succeeds, the return object is BridgeBoard instance.
        ///Else, return null.
        ///</summary>
        public static BridgeBoard ConnectByCapability(BridgeBoardEnum hEnum, int nIndex, UInt32 capability)
        {
            BridgeBoard b = new BridgeBoard(hEnum, nIndex, capability);
            if (b.hDev == IntPtr.Zero)
                b = null;
            return b;
        }

        ///<summary>
        ///Description: This function will return the specified module IBaseModule.
        ///Input Parameters: moduleName -> Module name that you want to get.
        ///If the function succeeds, the return object is the specified IBaseModule instance.
        ///Else, return null.
        ///</summary>
        public IBaseModule GetModule(string moduleName)
        {
            if (moduleList.ContainsKey(moduleName))
                return moduleList[moduleName];
            return null;
        }

        ///<summary>
        ///Description: This function will return the supported modules
        ///If the function succeeds, the return object is the supported modules string.
        ///Else, return null.
        ///</summary>
        public string[] ListSupportedModule()
        {
            string[] tmp = null;
            int i = 0;
            if (moduleList.Count >= 1)
            {
                tmp = new string[moduleList.Count];
                foreach (string s in moduleList.Keys)
                    tmp[i++] = s;
            }
            return tmp;
        }

        ///<summary>
        ///Description: This function will return the I2C module.
        ///If the function succeeds, the return object is the specified index 0 I2C module instance.
        ///Else, return null.
        ///</summary>
        public I2CModule GetI2CModule()
        {
            return (I2CModule)GetModule("I2C0");
        }

        public ExtQSPIModule GetExtQSPIModule()
        {
            return (ExtQSPIModule)GetModule("ExtQSPI");
        }

        ///<summary>
        ///Description: This function will return the I2C module.
        ///If the function succeeds, the return object is the specified indexed I2C module instance.
        ///Else, return null.
        ///</summary>
        public I2CModule GetI2CModule(int busIndex)
        {
            return (I2CModule)GetModule("I2C" + busIndex.ToString());
        }

        ///<summary>
        ///Description: This function will return the SPI module.
        ///If the function succeeds, the return object is the specified SPI module instance.
        ///Else, return null.
        ///</summary>
        public SPIModule GetSPIModule()
        {
            return (SPIModule)GetModule("SPI0");
        }

        ///<summary>
        ///Description: This function will return the SPI module.
        ///If the function succeeds, the return object is the specified SPI module instance.
        ///Else, return null.
        ///</summary>
        public SPIModule GetSPIModule(int busIndex)
        {
            return (SPIModule)GetModule("SPI" + busIndex.ToString());
        }

        ///<summary>
        ///Description: This function will return the GPIO module.
        ///If the function succeeds, the return object is the specified GPIO module instance.
        ///Else, return null.
        ///</summary>
        public GPIOModule GetGPIOModule()
        {
            return (GPIOModule)GetModule("GPIO");
        }

        ///<summary>
        ///Description: This function will return the GPIOExt module.
        ///If the function succeeds, the return object is the specified GPIOExt module instance.
        ///Else, return null.
        ///</summary>
        public GPIOExtModule GetGPIOExtModule()
        {
            return (GPIOExtModule)GetModule("GPIOExt");
        }

        ///<summary>
        ///Description: This function will return the PWM module.
        ///If the function succeeds, the return object is the specified PWM module instance.
        ///Else, return null.
        ///</summary>
        public PWMModule GetPWMModule()
        {
            return (PWMModule)GetModule("PWM");
        }

        ///<summary>
        ///Description: This function will return the ExtCustomizedCommand module.
        ///If the function succeeds, the return object is the specified ExtCustomizedCommand module instance.
        ///Else, return null.
        ///</summary>
        public ExtCustomizedCommandModule GetExtCustomizedCommandModule()
        {
            return (ExtCustomizedCommandModule)GetModule("ExtCustomizedCommand");
        }

        ///<summary>
        ///Description: This function will return the ExtGPIOMisc module.
        ///If the function succeeds, the return object is the specified ExtGPIOMisc module instance.
        ///Else, return null.
        ///</summary>
        public ExtGPIOMiscMdoule GetExtGPIOMiscModule()
        {
            return (ExtGPIOMiscMdoule)GetModule("ExtGPIOMisc");
        }

        ///<summary>
        ///Description: This function will return the ExtGSMW module.
        ///If the function succeeds, the return object is the specified ExtGSMW module instance.
        ///Else, return null.
        ///</summary>
        public ExtGSMWModule GetExtGSMWModule()
        {
            return (ExtGSMWModule)GetModule("ExtGSMW");
        }

        ///<summary>
        ///Description: This function will return the ExtGSOW module.
        ///If the function succeeds, the return object is the specified ExtGSOW module instance.
        ///Else, return null.
        ///</summary>
        public ExtGSOWModule GetExtGSOWModule()
        {
            return (ExtGSOWModule)GetModule("ExtGSOW");
        }

        ///<summary>
        ///Description: This function will return the ExtHSI2C module.
        ///If the function succeeds, the return object is the specified ExtHSI2C module instance.
        ///Else, return null.
        ///</summary>
        public ExtHSI2CModule GetExtHSI2CModule()
        {
            return (ExtHSI2CModule)GetModule("ExtHSI2C");
        }

        ///<summary>
        ///Description: This function will return the ExtIOConfig module.
        ///If the function succeeds, the return object is the specified ExtIOConfig module instance.
        ///Else, return null.
        ///</summary>
        public ExtIOConfigModule GetExtIOConfigModule()
        {
            return (ExtIOConfigModule)GetModule("ExtIOConfig");
        }

        ///<summary>
        ///Description: This function will return the ExtSecurityData module.
        ///If the function succeeds, the return object is the specified ExtSecurityData module instance.
        ///Else, return null.
        ///</summary>
        public ExtSecurityDataModule GetExtSecurityDataModule()
        {
            return (ExtSecurityDataModule)GetModule("ExtSecurityData");
        }

        ///<summary>
        ///Description: This function will return the ExtStorage module.
        ///If the function succeeds, the return object is the specified ExtStorage module instance.
        ///Else, return null.
        ///</summary>
        public ExtStorageModule GetExtStorageModule()
        {
            return (ExtStorageModule)GetModule("ExtStorage");
        }


        ///<summary>
        ///Description: This function will only need to be called by VisualBasic application to avoid application exeception.
        ///</summary>
        public void ForceStopEventWatcher()
        {
            mEventWatcher.Stop();
            mEventWatcher.Dispose();
            mEventWatcher = null;
        }
    }
}
