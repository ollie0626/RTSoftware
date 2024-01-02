﻿using System;
using System.Text;
using System.Text.RegularExpressions;

using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Threading;
using System.Collections;
using System.IO;
using Microsoft.Win32.SafeHandles;


// Richtek dll
//using RichtekStatsInterface;
using System.Security.Principal;
using System.Reflection;

//----------------------------------
using LibUsbDotNet;
using LibUsbDotNet.Info;
using LibUsbDotNet.Main;
using LibUsbDotNet.Descriptors;
using LibUsbDotNet.LibUsb;
using LibUsbDotNet.WinUsb;
using System.Collections.ObjectModel;


namespace Console_test
{

    // GUID 745a17a0-74d3-11d0-b6fe-00a0c90f57da
    // HID\VID_0488&PID_5755\6&25492db5&0&0000
    // PID = 0x5755
    // VID = 0x0488

    class Program
    {
        static void Main()
        {
            int vendorId = 0x0488;
            int productId = 0x5755;
            ErrorCode ec = ErrorCode.None;


            // Find the USB device with the specified VID and PID
            UsbDeviceFinder usbFinder = new UsbDeviceFinder(vendorId, productId);
            UsbDevice usbDevice = UsbDevice.OpenUsbDevice(usbFinder);

            if (usbDevice == null)
            {
                Console.WriteLine("USB device not found.");
                return;
            }
            // Open the USB device
            if (!usbDevice.Open())
            {
                Console.WriteLine("Failed to open USB device.");
                return;
            }

            IUsbDevice wholeUsbDevice = usbDevice as IUsbDevice;
            if (!ReferenceEquals(wholeUsbDevice, null))
            {
                // This is a "whole" USB device. Before it can be used, 
                // the desired configuration and interface must be selected.

                // Select config #1
                wholeUsbDevice.SetConfiguration(1);

                // Claim interface #0.
                wholeUsbDevice.ClaimInterface(0);
            }
            string cmdLine = Regex.Replace(
                    Environment.CommandLine, "^\".+?\"^.*? |^.*? ", "", RegexOptions.Singleline);
            UsbEndpointWriter writer = usbDevice.OpenEndpointWriter(WriteEndpointID.Ep01);
            int bytesWritten;
            ec = writer.Write(Encoding.Default.GetBytes(cmdLine), 2000, out bytesWritten);

            UsbEndpointReader reader = usbDevice.OpenEndpointReader(ReadEndpointID.Ep01);
            byte[] readBuffer = new byte[64];
            int bytesRead;
            ec = reader.Read(readBuffer, 1000, out bytesRead);


            usbDevice.Close();
            Console.ReadKey();
        }



    }
}
