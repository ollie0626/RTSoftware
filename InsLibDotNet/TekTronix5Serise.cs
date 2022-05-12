using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Drawing;
//using System.Drawing.Imaging;


namespace InsLibDotNet
{
    public class TekTronix5Serise : VisaCommand
    {
        public TekTronix5Serise(string Addr)
        {
            LinkingIns(Addr);
        }

        public TekTronix5Serise()
        {

        }

        ~TekTronix5Serise()
        {
            InsClose();
        }


        public void ConnectOscilloscope(string Addr)
        {
            LinkingIns(Addr);
        }

        public void TekTronixOSC_RST()
        {
            doCommand("*RST");
        }

        public string doQuery(string cmd)
        {
            return doQueryString(cmd);
        }

        public string doRead()
        {
            return "";//doReadString();
        }

        public void RootRun()
        {
            doCommand("ACQuire:STATE 1");
        }

        public void RootSTOP()
        {
            doCommand("ACQuire:STATE 0");
        }

        public void RootClear()
        {
            doCommand("CLEAR");
        }

        public void MesurePercent(int measX)
        {
            doCommand(string.Format("MEASUrement:MEAS{0}:REFLevels1:METHod PERCent", measX));
        }

        public void MesureAbsolute(int measX)
        {
            doCommand(string.Format("MEASUrement:MEAS{0}:REFLevels1:METHod ABSolute", measX));
        }
         
        public void MeasureAbsoluteSetting(int CHx, double Hi, double Mid, double Lo)
        {
            doCommand(string.Format("MEASUrement:CH{0}:REFLevels1:ABSolute:RISEHigh {1}", CHx, Hi));
            doCommand(string.Format("MEASUrement:CH{0}:REFLevels1:ABSolute:RISEMid {1}", CHx, Mid));
            doCommand(string.Format("MEASUrement:CH{0}:REFLevels1:ABSolute:RISELow {1}", CHx, Lo));
        }

        public void MeasureSource(int measX, int ch)
        {
            doCommand(string.Format("MEASUrement:MEAS{0}:SOURCE1 CH{1}", measX, ch));
        }

        public void DisplayChannel(int Ch, bool IsON)
        {
            string tmpStr = IsON ? "ON" : "OFF";
            doCommand(string.Format("DISplay:GLObal:CH{0}:STATE {1}",Ch,tmpStr));
        }

        public void DisplayChannelSel(int ch)
        {
            doCommand(string.Format("DISplay:SELect:SOUrce CH{0}", ch));
        }

        public void MeasDelete(int measx)
        {
            doCommand(string.Format("MEASUrement:DELete \"MEAS{0}\"", measx));
        }

        public void MeasPK2PK(int ch, bool IsMath = false)
        {
            if (IsMath) DiplaySelectMathChannel(ch);
            else DisplayChannelSel(ch);
            doCommand("MEASUrement:ADDMEAS PK2Pk");
        }

        public void MeasAmp(int ch, bool IsMath = false)
        {
            if (IsMath) DiplaySelectMathChannel(ch);
            else DisplayChannelSel(ch);
            doCommand(string.Format("MEASUrement:ADDMEAS AMPlITUDE"));
        }

        public void MeasMax(int ch, bool IsMath = false)
        {
            if (IsMath) DiplaySelectMathChannel(ch);
            else DisplayChannelSel(ch);
            doCommand(string.Format("MEASUrement:ADDMEAS MAXIMUM"));
        }

        public void MeasMean(int ch, bool IsMath = false)
        {
            if (IsMath) DiplaySelectMathChannel(ch);
            else DisplayChannelSel(ch);
            doCommand(string.Format("MEASUrement:ADDMEAS MEAN"));
        }

        public void MeasMin(int ch, bool IsMath = false)
        {
            if (IsMath) DiplaySelectMathChannel(ch);
            else DisplayChannelSel(ch);
            doCommand(string.Format("MEASUrement:ADDMEAS MINIMUM"));
        }

        public void MeasFreq(int ch, bool IsMath = false)
        {
            if (IsMath) DiplaySelectMathChannel(ch);
            else DisplayChannelSel(ch);
            doCommand(string.Format("MEASUrement:ADDMEAS FREQUENCY"));
        }

        public void MeasBase(int ch, bool IsMath = false)
        {
            if (IsMath) DiplaySelectMathChannel(ch);
            else DisplayChannelSel(ch);
            doCommand(string.Format("MEASUrement:ADDMEAS BASE"));
        }

        public void MeasTop(int ch, bool IsMath = false)
        {
            if (IsMath) DiplaySelectMathChannel(ch);
            else DisplayChannelSel(ch);
            doCommand(string.Format("MEASUrement:ADDMEAS TOP"));
        }

        public bool MeasXValue(int Measx, out double Val)
        {
            Val = doQueryNumber(string.Format("MEASUrement:MEAS{0}:VAL?", Measx));
            return doQueryNumber("*ESR?") < 0.1;
        }

        public bool MeasXMax(int Measx, out double Val)
        {
            Val = doQueryNumber(string.Format("MEASUrement:MEAS{0}:MAX?", Measx));
            return doQueryNumber("*ESR?") < 0.1;
        }

        public bool MeasXMin(int Measx, out double Val)
        {
            Val = doQueryNumber(string.Format("MEASUrement:MEAS{0}:MIN?", Measx));
            return doQueryNumber("*ESR?") < 0.1;
        }

        public void SetCHxOffset(int ch, double offset)
        {
            doCommand(string.Format("CH{0}:OFFSet {1:F2}", ch, offset));
        }

        public void SetCHxPosition(int ch, double pos)
        {
            doCommand(string.Format("CH{0}:POSition {1}", ch, pos));
        }

        public void SetCHxScale(int ch, double scale)
        {
            doCommand(string.Format("CH{0}:SCAle {1}", ch, scale));
        }

        public void SetTimeScale(double scale)
        {
            doCommand(string.Format("HORizontal:MODE:SCAle {0}",scale));
        }

        public void TriggerForce()
        {
            doCommand("TRIGger FORCe");
        }

        public void SetTriggerLevel(int ch, double level)
        {
            doCommand(string.Format("TRIGger:A:LEVel:CH{0} {1}",ch , level));
        }

        public void SetTriggerSource(int ch)
        {
            doCommand(string.Format("TRIGger:A:EDGE:SOUrce CH{0}", ch));
        }

        public void SetTriggerRise()
        {
            doCommand(string.Format("TRIGger:A:EDGE:SLOpe RISe"));
        }

        public void SetTriggerFall()
        {
            doCommand(string.Format("TRIGger:A:EDGE:SLOpe FALL"));
        }

        public void SetTriggerEither()
        {
            doCommand(string.Format("TRIGger:A:EDGE:SLOpe Either"));
        }

        public void SetNormalMode()
        {
            doCommand(string.Format("TRIGger:A:MODe NORMal"));
        }

        public void SetAutoMode()
        {
            doCommand(string.Format("TRIGger:A:MODe Auto"));
        }

        public void SetCHxCoupling(int ch, bool IsAC)
        {
            doCommand(string.Format("CH{0}:COUPling {1}", ch, IsAC ?"AC":"DC"));
        }

        public void SetCHxBandWidth(int ch)
        {
            doCommand(string.Format("CH{0}:BANDWIDTH {1}", ch, 20000000));
        }

        public void SetSampleRate(bool IsAuto, double SR)
        {
            if(IsAuto)
            {
                doCommand(string.Format("HORizontal:MODE Auto"));
                doCommand(string.Format("HORizontal:SAMPLERate:ANALYZemode:MINimum:VALue {0}", SR));
            }
            else
            {
                doCommand(string.Format("HORizontal:MODE MANual}"));
                doCommand(string.Format("HORizontal:MODE:SAMPLERate {0}", SR));
            }
        }
        //H Position range 0~100
        public void SetHPosition(int Pos)
        {         
            doCommand(string.Format("HORizontal:POSition {0}", Pos));     
        }

        public void SetArbFilter(int Math, int Ch)
        {
            doCommand(string.Format("MATH:MATH{0}:TYPe Advanced", Math));
            doCommand(string.Format("MATH:MATH{0}:DEFine \"^[CoefFileName=\"\"/media/C:/smooth200.flt\"\"]ArbFlt(Ch{1})\"",Math,Ch));
        }

        public void AddMath(int ch)
        {
            doCommand(string.Format("MATH:ADDNEW \"MATH{0}\"", ch));
        }

        public void DeleteMath(int ch)
        {
            doCommand(string.Format("MATH:DELete \"MATH{0}\"", ch));
        }

        public void DeleteMeasure(int MeaCh)
        {
            doCommand(string.Format("MEASUrement:DELete \"MEAS{0}\"", MeaCh));
        }

        public void SetCHxLabelName(int Ch,string Name,int Y = 99)
        {
            doCommand(string.Format("CH{1}:LABel:NAMe \"{0}\"", Name, Ch));
            if(Y < 99) doCommand(string.Format("CH{1}:LABel:YPOS {0}" , Y, Ch));
        }

        public void DiplayMathChannel(int ch, bool IsOn)
        {
            doCommand(string.Format("DISplay:GLObal:MATH{0}:STATE {1}", ch, IsOn ? 1 : 0));
        }

        public void DiplaySelectMathChannel(int ch)
        {
            doCommand(string.Format("DISplay:SELect:SOUrce MATH{0}", ch));
        }

        public void SetMathScale(int Math, double scale)
        {
            doCommand(string.Format("DISplay:WAVEView1:MATH:MATH{0}:VERTical:SCAle {1}", Math, scale));
        }

        public void SetMathPosition(int Math, double pos)
        {
            doCommand(string.Format("DISplay:WAVEView1:MATH:MATH{0}:VERTical:POSition {1}", Math, pos));
        }

        public void DisplayCursor(bool IsOn)
        {
            doCommand(string.Format("DISPLAY:WAVEVIEW1:CURSOR:CURSOR1:STATE {0}", IsOn? 1:0));
        }

        public void SetMeasXCursor(int MeasCh, bool IsGlobal = false)
        {
            doCommand(string.Format("MEASUrement:MEAS{0}:GATing CURSor", MeasCh));
            doCommand(string.Format("MEASUrement:MEAS{0}:GATing:GLOBal {1}", MeasCh, IsGlobal ? 1 : 0));
        }

        public void SetMeasXScreen(int MeasCh, bool IsGlobal = false)
        {
            doCommand(string.Format("MEASUrement:MEAS{0}:GATing Screen", MeasCh));
            doCommand(string.Format("MEASUrement:MEAS{0}:GATing:GLOBal {1}", MeasCh, IsGlobal ? 1:0));
        }

        public void SetWaveZoomShow(bool IsShow)
        {
            doCommand(string.Format("DISplay:WAVEView1:ZOOM:ZOOM1:STATe {0}", IsShow ? 1 : 0));
        }

        public void SetWaveZoomHScale(double Scale)
        {
            doCommand(string.Format("DISPLAY:WAVEVIEW1:ZOOM:ZOOM1:HORIZONTAL:WINSCALE {0}", Scale));
        }

        public void SetWaveZoomHPosition(int Pos)
        {
            doCommand(string.Format("DISPLAY:WAVEVIEW1:ZOOM:ZOOM1:HORIZONTAL:POSITION {0}", Pos));
        }

        public void SetCursorA(double pos)
        {
            doCommand(string.Format("DISplay:WAVEView1:CURSor:CURSOR1:VBArs:APOSition {0}", pos));
        }

        public void SetCursorB(double pos)
        {
            doCommand(string.Format("DISplay:WAVEView1:CURSor:CURSOR1:VBArs:BPOSition {0}", pos));
        }

        public void SetMathAVEMode(bool IsOn, int Cnt)
        {
            doCommand(string.Format("MATH:MATH1:AVG:MODE {0}", IsOn? "ON":"OFF"));
            doCommand(string.Format("MATH:MATH1:AVG:WEIGht {0}", Cnt));
        }

        public void SaveWaveform(string path, string filename, bool IsComp = false)
        {
            if (device == 0) return;
            string buf = path.Substring(path.Length - 1, 1) == @"/" ? path.Substring(0, path.Length - 1) : path;
            if(IsComp) buf = buf + @"/" + filename + ".jpeg";
            else buf = buf + @"/" + filename + ".png";


            string scope_cwd = "FILESystem:CWD " + (char)34 + @"C:" + (char)34;
            doCommand(scope_cwd);

            string scope_dest = "SAVEON:FILE:DEST " + (char)34 + @"C:" + (char)34;
            doCommand(scope_dest);

            string scope_imag_on = "SAVEON:IMAG ON";
            doCommand(scope_imag_on);

            string scope_save_image = "SAVE:IMAGE " + (char)34 + @"C:\tmp.png" + (char)34;
            doCommand(scope_save_image);
            System.Threading.Thread.Sleep(500);
            string scope_readfile = "FILESYSTEM:READFILE " + (char)34 + @"C:\tmp.png" + (char)34;
            doCommand(scope_readfile);

#if DEBUG
            /*Console.WriteLine("scope save path " + buf);
            Console.WriteLine(scope_cwd);
            Console.WriteLine(scope_dest);
            Console.WriteLine(scope_imag_on);
            Console.WriteLine(scope_save_image);
            Console.WriteLine(scope_readfile);*/
#endif

            System.Threading.Thread.Sleep(1000);
            int count_out = 0;
            int len = 500000;
            byte[] bytRead = new byte[len];
            visa32.viBufRead(device, bytRead, len, out count_out);

            if (!IsComp)
            {
                if (count_out > 0)
                {
                    FileStream fStream = File.Open(buf, FileMode.Create);
                    fStream.Write(bytRead, 0, bytRead.Length);
                    //System.Threading.Thread.Sleep(200);
                    fStream.Close();
                    fStream.Dispose();
                }
            }
            else
            {
                //TODO: Save waveform Fix
                System.IO.MemoryStream ms = new System.IO.MemoryStream(bytRead, 0, count_out);
                //Bitmap bmp = (Bitmap)Bitmap.FromStream(ms);
                //bmp.Save(buf, ImageFormat.Jpeg);
                //bmp.Dispose();
                ms.Dispose();
            }

            visa32.viFlush(device, visa32.VI_READ_BUF);
            visa32.viFlush(device, visa32.VI_WRITE_BUF);
            bytRead = null;
        }


        public void MeasPeriod(int ch)
        {
            string cmd = "DISplay:SELect:SOUrce CH" + ch.ToString();
            doCommand(cmd);
            cmd = "MEASUrement:ADDMEAS PERIOD";
            doCommand(cmd);
        }

        public void MeasHigh(int ch)
        {
            string cmd = "DISplay:SELect:SOUrce CH" + ch.ToString();
            doCommand(cmd);
            cmd = "MEASUrement:ADDMEAS HIGH";
            doCommand(cmd);
        }

        public void MeasLOW(int ch)
        {
            string cmd = "DISplay:SELect:SOUrce CH" + ch.ToString();
            doCommand(cmd);
            cmd = "MEASUrement:ADDMEAS LOW";
            doCommand(cmd);
        }

        public void MeasFallTime(int ch)
        {
            string cmd = "DISplay:SELect:SOUrce CH" + ch.ToString();
            doCommand(cmd);
            cmd = "MEASUrement:ADDMEAS FALLTIME";
            doCommand(cmd);
        }

        public void MeasNDuty(int ch)
        {
            string cmd = "DISplay:SELect:SOUrce CH" + ch.ToString();
            doCommand(cmd);
            cmd = "MEASUrement:ADDMEAS NDUty";
            doCommand(cmd);
        }

        public void MeasPDuty(int ch)
        {
            string cmd = "DISplay:SELect:SOUrce CH" + ch.ToString();
            doCommand(cmd);
            cmd = "MEASUrement:ADDMEAS PDUTY";
            doCommand(cmd);
        }

        public void MeasPOverShoot(int ch)
        {
            string cmd = "DISplay:SELect:SOUrce CH" + ch.ToString();
            doCommand(cmd);
            cmd = "MEASUrement:ADDMEAS POVERSHOOT";
            doCommand(cmd);
        }

        public void MeasNOverShoot(int ch)
        {
            string cmd = "DISplay:SELect:SOUrce CH" + ch.ToString();
            doCommand(cmd);
            cmd = "MEASUrement:ADDMEAS NOVERSHOOT";
            doCommand(cmd);
        }

        public void MeasRiseTime(int ch)
        {
            string cmd = "DISplay:SELect:SOUrce CH" + ch.ToString();
            doCommand(cmd);
            cmd = "MEASUrement:ADDMEAS RISETIME";
            doCommand(cmd);
        }

        public void MeasRiseSlewRate(int ch)
        {
            string cmd = "DISplay:SELect:SOUrce CH" + ch.ToString();
            doCommand(cmd);
            cmd = "MEASUrement:ADDMEAS RISESLEWRATE";
            doCommand(cmd);
        }

        public void MeasFallSlewRate(int ch)
        {
            string cmd = "DISplay:SELect:SOUrce CH" + ch.ToString();
            doCommand(cmd);
            cmd = "MEASUrement:ADDMEAS FALLSLEWRATE";
            doCommand(cmd);
        }

        public void InitWaveform(ref double step)
        {
            doCommand(string.Format(":DATa:ENCdg RIBinary"));
            doCommand(string.Format(":DATa:START 1"));
            //double len = doQueryNumber(string.Format("HORizontal:ACQLENGTH?", ch));
            doCommand(string.Format(":DATa:STOP 1250000"));
            doCommand(string.Format(":WFMOutpre:ENCdg BINary"));
            doCommand(string.Format(":WFMOutpre:BYT_Nr 1"));
            step = doQueryNumber(string.Format(":WFMOutpre:XINcr?"));
        }

        public byte[] GetWaveform(int ch, ref double ymult, ref double  yoff, ref double  yzero)
        {
            string dis = doQueryString(string.Format("DISplay:GLObal:CH{0}:STATE?", ch));
            if (dis.Contains("0")) return null;
            doCommand(string.Format(":DATa:SOUrce CH{0}", ch));
            //doCommand(string.Format(":DATa:ENCdg RIBinary", ch));
            //doCommand(string.Format(":DATa:START 1", ch));
            //double len = doQueryNumber(string.Format("HORizontal:ACQLENGTH?", ch));
            //doCommand(string.Format(":DATa:STOP 1000000", ch));
            //doCommand(string.Format(":WFMOutpre:ENCdg BINary", ch));
            //doCommand(string.Format(":WFMOutpre:BYT_Nr 1", ch));
            ymult = doQueryNumber(string.Format(":WFMOutpre:YMUlt?", ch));
            yoff = doQueryNumber(string.Format(":WFMOutpre:YOFf?", ch));
            yzero = doQueryNumber(string.Format(":WFMOutpre:YZEro?", ch));
            doCommand(string.Format(":Curve?", ch));
            int count_out = 0;
            int len = 1250000;
            byte[] bytRead = new byte[len];
            visa32.viBufRead(device, bytRead, len, out count_out);

            visa32.viFlush(device, visa32.VI_READ_BUF);
            visa32.viFlush(device, visa32.VI_WRITE_BUF);
            return bytRead;
            /*using (StreamWriter sw = new StreamWriter("D:\\TestFile.csv"))   //小寫TXT     
            {
                for(int i = 8; i < count_out; ++i)
                {
                    double tmp = ((double)((sbyte)bytRead[i]) - yoff) * ymult + yzero;
                    sw.WriteLine(tmp.ToString());
                }
               
            }*/
        }

    }
}





//visa32.viOpenDefaultRM(out Rm);
//visa32.viParseRsrc(Rm, "USB0::0x0699::0x0522::C013665::INSTR", ref intfType, ref intfNum);
//visa32.viOpen(Rm, "USB0::0x0699::0x0522::C013665::INSTR", 0, 0, out vi);
//visa32.viPrintf(vi, "*IDN?\r\n");
//visa32.viRead(vi, Buffer, count_in, out count_out);
//str = Encoding.ASCII.GetString(Buffer, 0, count_out);
//Console.Write(str);
//visa32.viPrintf(vi, "*IDN?\r\n");
//visa32.viRead(vi, Buffer, count_in, out count_out);
//str = Encoding.ASCII.GetString(Buffer, 0, count_out);
//Console.Write(str);
//visa32.viPrintf(vi, "FILESystem:CWD " + "\"C:\"\r\n");
//visa32.viPrintf(vi, "FILESystem:CWD?\r\n");
//visa32.viRead(vi, Buffer, count_in, out count_out);
//str = Encoding.ASCII.GetString(Buffer, 0, count_out);
//Console.Write(str);
//str = "SAVEON:FILE:DEST " + "\"C:/temp/\"\r\n";
//visa32.viPrintf(vi, str);

//str = "SAVEON:IMAG ON\r\n";
//visa32.viPrintf(vi, str);
//System.Threading.Thread.Sleep(500);
//str = "SAVE:IMAGE " + "\"C:/temp/Tek_wave.png\"\r\n";
//visa32.viPrintf(vi, str);
//System.Threading.Thread.Sleep(500);
//string cmd = "FILESYSTEM:READFILE " + "\"C:/temp/Tek_wave.png\"\r\n";
//visa32.viPrintf(vi, cmd);
//System.Threading.Thread.Sleep(1000);
//int dataLen = 500000;
//byte[] bytRead = new byte[dataLen];
//visa32.viBufRead(vi, bytRead, dataLen, out count_out);
//FileStream fStream = File.Open(@"D:\\Tek000.png", FileMode.Create);
//fStream.Write(bytRead, 0, bytRead.Length);
//System.Threading.Thread.Sleep(1000);
//fStream.Close();
//fStream.Dispose();

