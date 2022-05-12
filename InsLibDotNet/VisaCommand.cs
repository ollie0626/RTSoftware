using System;
using System.Text;
using System.Runtime.InteropServices;


public class VisaCommand
{
    //short intfType = 0;
    //short intfNum = 0;
    int count_out;
    int count_in = 1024;
    public int device = 0;
    protected int Rm = 0;
    byte[] buffer = new byte[1024];
    protected string idn = "";
    public static bool _IsDebug = false;
    short intfType = 0;
    short intfNum = 0;


    public virtual void LinkingIns(string Addr)
    {
#if true
        if (_IsDebug == true || Addr == "")
        {
            device = 0;
            return;
        }
        if (Rm == 0) visa32.viOpenDefaultRM(out Rm);
        visa32.viParseRsrc(Rm, Addr, ref intfType, ref intfNum);
        visa32.viOpen(Rm, Addr, 0, 0, out device);
        visa32.viSetAttribute(device, visa32.VI_ATTR_TMO_VALUE, 1000);
        Console.WriteLine(Addr + "   " + device);
#else
        Rm = 0;
#endif
    }

    protected string doQueryString(string cmd)
    {
#if true
        if (device == 0) return "";

        string str;
        Array.Clear(buffer, 0, buffer.Length);
        visa32.viPrintf(device, cmd + "\r\n");
        visa32.viRead(device, buffer, count_in, out count_out);
        str = Encoding.ASCII.GetString(buffer, 0, count_out);
        return str;
#else
        return "";
#endif
    }

    protected string doReadString()
    {
        if (device == 0) return "";
        string str;
        Array.Clear(buffer, 0, buffer.Length);
        visa32.viRead(device, buffer, count_in, out count_out);
        str = Encoding.ASCII.GetString(buffer, 0, count_out);
        return str;
    }

    public double doQueryNumber(string cmd)
    {
#if true
        if (device == 0) return 0.0;
        double temp = 0.0;
        string[] str = {};
        int visaState;
        int count = 0;
        
        while(count < 3)
        {
            Array.Clear(buffer, 0, buffer.Length);
            visa32.viPrintf(device, cmd + "\r\n");
            visa32.viRead(device, buffer, count_in, out count_out);
            visaState = visa32.viFlush(device, visa32.VI_WRITE_BUF);
            visaState = visa32.viFlush(device, visa32.VI_READ_BUF);
            str = Encoding.ASCII.GetString(buffer, 0, count_out).Split(',');
            if (str[0] != "")
            {
                temp = Convert.ToDouble(str[0]);
                return temp;
            }
            count++;
        }
        
        return 0;
#else
        return 0;
#endif
    }

    protected void doCommand(string cmd)
    {
#if true
        if (device == 0) return;
        visa32.viPrintf(device, cmd + "\r\n");
#endif
    }

    protected void doCommandViWrite(string cmd)
    {
        if (device == 0) return;
        int cout = 0;
        byte[] ASCIIbytes = Encoding.ASCII.GetBytes(cmd);
        visa32.viWrite(device, ASCIIbytes, ASCIIbytes.Length, out cout);
    }


    private bool WaveformCheck(ref byte[] Arr)
    {
        bool state = true;
        for (int i = 0; i < Arr.Length; i++)
        {
            if (Arr[i] != 0)
            {
                state = false;
                break; /* get the picture ok */
            }
            else
            {
                state = true; /* return true keep scan data buf */
            }
        }
        return state; /* return false stop scan data buf */
    }

    protected int IEEEBlock_Bytes(out byte[] Arr)
    {
        if (device == 0)
        {
            Arr = null;
            return 0;
        }

#if true
        //int len = 5000000;
        int len = 300000;
        Arr = new byte[len];
        bool check = true;
        while(check)
        {
            // Argument type: A reference to a data array.
            visa32.viScanf(device, "%#b\n", ref len, Arr);
            System.Threading.Thread.Sleep(500);
            /* all of buffer = 0, keep scan picture data */
            check = WaveformCheck(ref Arr);
        }
        visa32.viFlush(device, visa32.VI_WRITE_BUF);
        visa32.viFlush(device, visa32.VI_READ_BUF);
        return len;
#else
        Arr = new byte[100];
        return 0;
#endif
    }

    protected int IEEEBlock_Bytes2(ref byte[] Arr)
    {
        if (device == 0)
        {
            Arr = null;
            return 0;
        }

#if true
        int len = Arr.Length;
        //Argument type: A location of block binary data.
        int visaState = visa32.viScanf(device, "%#y", ref len, Arr);
        visaState = visa32.viFlush(device, visa32.VI_WRITE_BUF);
        visaState = visa32.viFlush(device, visa32.VI_READ_BUF);
        return len;
#else
        return 0;
#endif
    }

    protected void doQueryNumbers(string cmd, ref double[] arr)
    {
#if true
        if (device == 0) return;
        visa32.viPrintf(device, cmd + "\r\n");
        visa32.viScanf(device, "%,20lfn", arr);
#endif
    }

    public void InsClose()
    {
#if true
        if (device == 0) return;
        visa32.viClose(device);
        device = 0;
        //visa32.viClose(Rm);
#endif
    }

    public string doQueryIDN()
    {
        if (device == 0) return "unlink Ins.";
        idn = doQueryString("*IDN?");
        return idn;
    }
    public virtual bool InsState()
    {
        return (device != 0 && doQueryIDN() != "");
    }


}

public class ViCMD
{
    static int Rm;
    static int vi;
    static StringBuilder Desc = new StringBuilder();

    public static string[] ScanIns()
    {
        if (VisaCommand._IsDebug == true) return null;
        int retCount;
        visa32.viOpenDefaultRM(out Rm);
        visa32.viFindRsrc(Rm, "?*INSTR", out vi, out retCount, Desc);
        string[] InsList = null;
        if (retCount > 0)
        {
            InsList = new string[retCount];
            InsList[0] = Desc.ToString();
            for (int i = 1; i < retCount; i++)
            {
                visa32.viFindNext(vi, Desc);
                InsList[i] = Desc.ToString();
            }
        }
        return InsList;
    }

}


