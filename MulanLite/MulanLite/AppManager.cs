using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Sunny.UI;
using System.Windows.Forms;
using System.Runtime.InteropServices;

namespace MulanLite
{
    public partial class main : UIForm
    {
        private IntPtr m_hNotifyDevNode;
        private string win_name = "Mulan Lite tool v3.2";

        private void timer1_Tick(object sender, EventArgs e)
        {
            RTDev.BoardInit();
            timer1.Enabled = false;
            Console.WriteLine("one time shot timer !!!");
        }

        private void main_Load(object sender, EventArgs e)
        {
            //Guid hidGuid = new Guid("745a17a0-74d3-11d0-b6fe-00a0c90f57da");
            //Dbt.HidD_GetHidGuid(ref hidGuid);
            //RegisterNotification(hidGuid);

            // LED packet test
            //int[] buf = new int[] { 0x0E25D, 0x14AFF, 0x0E6F8, 0x1D8C7, 0x12DDE, 0x12DDE, 0x0E6F8, 0x14AFF, 0x12DDE };
            //int[] buf = new int[] { 0x0E25D, 0x14AFF, 0x0E6F8, 0x1D8C7, 0x12DDE };
            //int[] buf = new int[] { 0x12DDE };
            //RTDev.LEDPacket((byte)(buf.Length - 1), 0x0406, buf);


            NumericUpDown[] nu_table = new NumericUpDown[]
            {
                nuopen_ch4, nuopen_ch3, nuopen_ch2, nuopen_ch1, nu_crc, nu_rdo, nu_badlen, nu_badadd, nu_badcmd, nu_badid,
                flag1, flag2, flag3, flag4, flag5, flag6, flag7, flag8, flag9, flag10, flag11, flag12, flag13,
                nu_dont_lower, nu_raise, nu_test_mode, nu_stat_dis, nu_stat_norm, nu_stat_stb, nu_stat_iden, nu_stat_init,
                nu_efuse_load, nu_tsd_mask, nu_tsd, nushort_ch1, nushort_ch2, nushort_ch3, nushort_ch4, numericUpDown1
            };

            for (int i = 0; i < nu_table.Length; i++)
            {
                //nu_table[i].Controls[0].Visible = false;
                nu_table[i].Enabled = false;
            }



            uiTabControl1.TabPages.RemoveAt(4);
            timer1.Interval = 500;
            timer1.Enabled = false;

            Guid hidGuid = new Guid("745a17a0-74d3-11d0-b6fe-00a0c90f57da");
            Dbt.HidD_GetHidGuid(ref hidGuid);
            RegisterNotification(hidGuid);
        }

        private void main_FormClosing(object sender, FormClosingEventArgs e)
        {
            UnregisterNotification();
        }

        private void RegisterNotification(Guid guid)
        {
            Dbt.DEV_BROADCAST_DEVICEINTERFACE devIF = new Dbt.DEV_BROADCAST_DEVICEINTERFACE();
            IntPtr devIFBuffer;

            // Set to HID GUID
            devIF.dbcc_size = Marshal.SizeOf(devIF);
            devIF.dbcc_devicetype = Dbt.DBT_DEVTYP_DEVICEINTERFACE;
            devIF.dbcc_reserved = 0;
            devIF.dbcc_classguid = guid;
            devIF.dbcc_name = (char)0;

            // Allocate a buffer for DLL call
            devIFBuffer = Marshal.AllocHGlobal(devIF.dbcc_size);

            // Copy devIF to buffer
            Marshal.StructureToPtr(devIF, devIFBuffer, true);

            // Register for HID device notifications
            m_hNotifyDevNode = Dbt.RegisterDeviceNotification(this.Handle, devIFBuffer, Dbt.DEVICE_NOTIFY_WINDOW_HANDLE);

            // Copy buffer to devIF
            Marshal.PtrToStructure(devIFBuffer, devIF);

            // Free buffer
            Marshal.FreeHGlobal(devIFBuffer);
        }

        private void UnregisterNotification()
        {
            uint ret = Dbt.UnregisterDeviceNotification(m_hNotifyDevNode);
        }

        /*
            WndProc
        */
        protected override void WndProc(ref Message m)
        {
            // Intercept the WM_DEVICECHANGE message
            if (m.Msg == Dbt.WM_DEVICECHANGE)
            {
                // Get the message event type
                int nEventType = m.WParam.ToInt32();

                // Check for devices being connected or disconnected
                if (nEventType == Dbt.DBT_DEVICEARRIVAL ||
                    nEventType == Dbt.DBT_DEVICEREMOVECOMPLETE)
                {
                    Dbt.DEV_BROADCAST_HDR hdr = new Dbt.DEV_BROADCAST_HDR();

                    // Convert lparam to DEV_BROADCAST_HDR structure
                    Marshal.PtrToStructure(m.LParam, hdr);

                    if (hdr.dbch_devicetype == Dbt.DBT_DEVTYP_DEVICEINTERFACE)
                    {
                        Dbt.DEV_BROADCAST_DEVICEINTERFACE_1 devIF = new Dbt.DEV_BROADCAST_DEVICEINTERFACE_1();

                        // Convert lparam to DEV_BROADCAST_DEVICEINTERFACE structure
                        Marshal.PtrToStructure(m.LParam, devIF);

                        // Get the device path from the broadcast message
                        string devicePath = new string(devIF.dbcc_name);

                        // Remove null-terminated data from the string
                        int pos = devicePath.IndexOf((char)0);
                        if (pos != -1)
                        {
                            devicePath = devicePath.Substring(0, pos);
                        }

                        // An HID device was connected or removed
                        if (nEventType == Dbt.DBT_DEVICEREMOVECOMPLETE)
                        {
                            RTDev.BoardRemove();
                            //RTBBStatus.ForeColor = Color.Red;
                            //RTBBStatus.Text = "Disconnected";
                            //WriteTableReset();
                            this.Text = win_name + " - Disconnected";
                            Console.WriteLine("RTBridge board remove!!!");
                        }
                        else if (nEventType == Dbt.DBT_DEVICEARRIVAL)
                        {

                            Console.WriteLine("RTBridge board arrived!!!");
                            //RTBBStatus.ForeColor = Color.AliceBlue;
                            //RTBBStatus.Text = "Connected";
                            this.Text = win_name + " - Connected";

                            timer1.Enabled = true;
                        }
                    }
                }
            }
            base.WndProc(ref m);
        }

    }
}
