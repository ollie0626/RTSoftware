Imports System.Runtime.InteropServices

Module Module_RTBBLib
    Public pI2CAvailableAddressVBarray(15) As Byte
    'Public hEnum As BridgeBoardEnum
    'Public hDevice As BridgeBoard
    'Public I2C_Module As I2CModule
    'Public GSMW_Module As RTBBLibDotNet.ExtGSMWModule
    'Public GPIO_Module As RTBBLibDotNet.GPIOModule
    ''Public CFW_Module As RTBBLibDotNet.ExtCustomizedCommandModule

    Public RTBB_board As Boolean = False
    Dim Result As Integer



    '---------------------------------------------------------------------
    'I2C
    Public no_slave As String = "No slave device"

    'Public Device_ID As Byte
    Public Device_OK As Boolean
    Dim W_DataBuffer(&HFF) As Byte
    Dim R_DataBuffer(&HFF) As Byte
    '---------------------------------------------------------------------
    'GPIO

    'Dim GPIO_Data(15) As UShort

    'Dim nLength As UInteger = 4

    'Dim nMask As UInteger = CInt("&H00FF")

    'Dim bCheckAck As Boolean = False

    Function Check_Eagleboard() As Boolean
        hEnum = RTBB_EnumBoard()


        BoardCount = RTBB_GetBoardCount(hEnum)

        Main.txt_ID.Text = no_slave
        Main.status_bridgeboad.Text = no_device
        Main.num_ID.Value = 0

        RTBB_board = False
        If BoardCount = 0 Then
            Main.status_bridgeboad.Text = no_device
            Exit Function
        Else


            If BoardCount > 1 Then
                ReDim Device_List(BoardCount - 1)
                ReDim VID_List(BoardCount - 1)

                pEnumBoardInfo = RTBB_GetEnumBoardInfo(hEnum, 0)
                strLibraryName = Marshal.PtrToStringAnsi(RTBB_BIGetLibraryName(pEnumBoardInfo))
                strFirmwareInfo = Marshal.PtrToStringAnsi(RTBB_BIGetFirmwareInfo(pEnumBoardInfo))
                Main.status_bridgeboad.Text = strLibraryName & " (" & strFirmwareInfo & ")"



                For i = 0 To BoardCount - 1
                    Device_List(i) = RTBB_ConnectToBridgeByIndex(i)
                    VID_List(i) = RTBB_BIGetVID(Device_List(i))

                    Console.WriteLine(RTBB_BIGetVID(Device_List(i)))

                    If (IsDBNull(Device_List(i))) Then
                        Main.status_bridgeboad.Text = no_device
                        error_message("Connect Bridge Board Fail")
                        Exit Function
                    End If

                    For a = 0 To 6
                        RTBB_GPIOSingleSetIODirection(Device_List(i), 32 + a, True)
                        RTBB_GPIOSingleWrite(Device_List(i), 32 + a, False) '0
                    Next
                    RTBB_board = True

                    I2CScan()
                    I2CSetFrequency(1024, 1000)
                Next
            Else
                pEnumBoardInfo = RTBB_GetEnumBoardInfo(hEnum, 0)
                strLibraryName = Marshal.PtrToStringAnsi(RTBB_BIGetLibraryName(pEnumBoardInfo))
                strFirmwareInfo = Marshal.PtrToStringAnsi(RTBB_BIGetFirmwareInfo(pEnumBoardInfo))
                Main.status_bridgeboad.Text = strLibraryName & " (" & strFirmwareInfo & ")"
                'Connect Board
                hDevice = RTBB_ConnectToBridgeByIndex(0)
                If (IsDBNull(hDevice)) Then
                    Main.status_bridgeboad.Text = no_device
                    error_message("Connect Bridge Board Fail")

                    Exit Function
                End If
                '----------------------------------------------------
                'I2C_Module = hDevice.GetI2CModule(0)
                '' GSMW_Module = hDevice.GetExtGSMWModule()
                'GPIO_Module = hDevice.GetGPIOModule()
                For i = 0 To 6
                    RTBB_GPIOSingleSetIODirection(hDevice, 32 + i, True)
                    RTBB_GPIOSingleWrite(hDevice, 32 + i, False) '0
                Next
                RTBB_board = True
                I2CScan()
                I2CSetFrequency(1024, 1000, device_sel)
            End If
        End If



    End Function

    Function I2CSetFrequency(ByVal nMode As Integer, ByVal nFrekHz As Integer, Optional ByVal sel As Integer = 0) As Integer
        If BoardCount >= 2 Then
            RTBB_I2CSetFrequency(Device_List(sel), I2CBus, nMode, nFrekHz)
        Else
            RTBB_I2CSetFrequency(hDevice, I2CBus, nMode, nFrekHz)
        End If
    End Function

    Function I2CScan() As Integer
        Dim SlaveAddr As Integer
        Dim i As Integer
        Dim check As Boolean

        Main.txt_ID.Text = no_slave

        Main.data_relay.Rows.Clear()

        Main.data_meas.Rows.Clear()

        For i = 0 To Relay_ID.Length - 1
            Relay_ID_check(i) = False
        Next

        For i = 0 To Meas_ID.Length - 1
            Meas_ID_check(i) = False
        Next

        If BoardCount > 1 Then

        Else
            Result = RTBB_I2CScanSlaveDevice(hDevice, I2CBus, pI2CAvailableAddressVBarray(0))
        End If

        If Result = RT_BB_SUCCESS Then
            SlaveAddr = 0
            While (1)
                SlaveAddr = RTBB_I2CGetFirstValidSlaveAddr(pI2CAvailableAddressVBarray(0), SlaveAddr)
                If SlaveAddr < 0 Then
                    Exit While
                End If

                If Main.txt_ID.Text = no_slave Then
                    Main.txt_ID.Text = "Scan Device= 0x" & SlaveAddr.ToString("X2")
                Else
                    Main.txt_ID.Text = Main.txt_ID.Text & ", 0x" & SlaveAddr.ToString("X2")
                End If
                check = False
                'Check Relay ID
                For i = 0 To Relay_ID.Length - 1
                    If Relay_ID(i) = SlaveAddr Then
                        Relay_ID_check(i) = True
                        Main.data_relay.Rows.Add(SlaveAddr.ToString("X2"), Relay_signal_A(i), Relay_signal_B(i))
                        check = True
                        Exit For
                    End If
                Next
                'Check Meas ID
                For i = 0 To Meas_ID.Length - 1
                    If Meas_ID(i) = SlaveAddr Then
                        Meas_ID_check(i) = True
                        Main.data_meas.Rows.Add(SlaveAddr.ToString("X2"), Meas_signal(i))
                        check = True
                        Exit For
                    End If
                Next
                If check = False Then
                    Main.num_ID.Value = SlaveAddr
                End If
                SlaveAddr += 1
            End While
            'Device_OK = True
        End If

        If Main.txt_ID.Text = no_slave Then
            Device_OK = False
        Else
            Device_OK = True
        End If

        If Main.data_relay.Rows.Count > 0 Then
            relay_init()
        End If


        If Main.data_meas.Rows.Count > 0 Then
            Main.cbox_INA226_b11_9.SelectedIndex = 0
            Main.Panel_INA226.Enabled = True
        Else

            Main.Panel_INA226.Enabled = False
        End If


    End Function

    'Function GPIO_out(ByVal bits As Integer, ByVal bit_value() As Integer) As Integer
    '    Dim i As Integer




    '    nLength = 1
    '    GPIO_Data(0) = 0
    '    For i = 0 To bits - 1
    '        GPIO_Data(0) = GPIO_Data(0) + Val(bit_value(i) * 2 ^ i)
    '    Next

    '    Result = GSMW_Module.RTBB_EXTGSMW_SendData(nLength, nMask, GPIO_Data, bCheckAck)

    '    If Result <> ExtGSMWModule.RT_BB_SUCCESS Then
    '        Check_Eagleboard()
    '        Result = GSMW_Module.RTBB_EXTGSMW_SendData(nLength, nMask, GPIO_Data, bCheckAck)

    '        If Result <> ExtGSMWModule.RT_BB_SUCCESS Then

    '            error_message("Bridgeboard Error: " & ExtGSMWModule.RTBB_Result2String(Result))
    '        End If
    '    End If


    'End Function

    Function GPIO_out(ByVal bits As Integer, ByVal bit_value() As Integer, Optional ByVal sel As Integer = 0) As Integer
        Dim i As Integer


        If BoardCount >= 2 Then
            For i = 0 To bits - 1
                If bit_value(i) = 0 Then
                    RTBB_GPIOSingleWrite(Device_List(sel), 32 + i, False) '0
                Else
                    RTBB_GPIOSingleWrite(Device_List(sel), 32 + i, True) '1
                End If

            Next
        Else
            For i = 0 To bits - 1
                If bit_value(i) = 0 Then
                    RTBB_GPIOSingleWrite(hDevice, 32 + i, False) '0
                Else
                    RTBB_GPIOSingleWrite(hDevice, 32 + i, True) '1
                End If

            Next
        End If




    End Function

    Function GPIO_single_write(ByVal bit As Integer, ByVal value As Integer, Optional ByVal sel As Integer = 0) As Integer


        If BoardCount >= 2 Then
            If value = 0 Then
                RTBB_GPIOSingleWrite(Device_List(sel), 32 + bit, False) '0
            Else
                RTBB_GPIOSingleWrite(Device_List(sel), 32 + bit, True) '1
            End If
        Else
            If value = 0 Then
                RTBB_GPIOSingleWrite(hDevice, 32 + bit, False) '0
            Else
                RTBB_GPIOSingleWrite(hDevice, 32 + bit, True) '1
            End If

        End If



    End Function

    Function GPIO_single_read(ByVal bit As Integer, Optional ByVal sel As Integer = 0) As Integer
        Dim bValue As Boolean

        If BoardCount >= 2 Then
            RTBB_GPIOSingleRead(Device_List(sel), 32 + bit, bValue)
        Else

            RTBB_GPIOSingleRead(hDevice, 32 + bit, bValue)
        End If

        If bValue = True Then
            Return 1
        Else
            Return 0
        End If




    End Function

    Function reg_write_multi(ByVal ID As Byte, ByVal addr As Byte, ByVal data() As Byte, Optional ByVal sel As Integer = 0) As String
        Dim num As Integer
        Dim i As Integer

        num = data.Length

        For i = 0 To num - 1
            W_DataBuffer(i) = data(i)
        Next

        If BoardCount >= 2 Then
            Result = RTBB_I2CWrite(Device_List(sel), I2CBus, ID, 1, addr, num, W_DataBuffer(0))
            If Result <> 0 Then
                Result = RTBB_I2CWrite(Device_List(sel), I2CBus, ID, 1, addr, num, W_DataBuffer(0))
            End If
        Else
            Result = RTBB_I2CWrite(hDevice, I2CBus, ID, 1, addr, num, W_DataBuffer(0))
            If Result <> 0 Then
                Result = RTBB_I2CWrite(hDevice, I2CBus, ID, 1, addr, num, W_DataBuffer(0))
            End If
        End If
        Main.status_error.Text = i2c_status(Result)
        Return Result
    End Function

    Function i2c_status(ByVal num As Integer) As String
        Dim status As String = ""
        Select Case num
            Case 0
                status = "RT_BB_SUCCESS"
            Case -1
                status = "RT_BB_BAD_PARAMETER"
            Case -2
                status = "RT_BB_HARDWARE_NOT_FOUND"
            Case -3
                status = "RT_BB_SLAVE_DEVICE_NOT_FOUND"
            Case -4
                status = "RT_BB_TRANSACTION_FAILED"
            Case -5
                status = "RT_BB_SLAVE_OPENNING_FOR_WRITE_FAILED"
            Case -6
                status = "RT_BB_SLAVE_OPENNING_FOR_READ_FAILED"
            Case -7
                status = "RT_BB_SENDING_MEMORY_ADDRESS_FAILED"
            Case -8
                status = "RT_BB_SENDING_DATA_FAILED"
            Case -9
                status = "RT_BB_NOT_IMPLEMENTED"
            Case -10
                status = "RT_BB_NO_ACK"
            Case -11
                status = "RT_BB_DEVICE_BUSY"
            Case -12
                status = "RT_BB_MEMORY_ERROR"
            Case -13
                status = "RT_BB_UNKNOWN_ERROR"
            Case -14
                status = "RT_BB_I2C_TIMEOUT"
            Case -15
                status = "RT_BB_IDLE"
            Case -16
                status = "RT_BB_NO_DATA"
            Case -17
                status = "RT_BB_BUFFER_OVERFLOW"
        End Select
        Return status
    End Function

    Function reg_write_word(ByVal ID As Byte, ByVal addr As Byte, ByVal data() As Byte, Optional ByVal sel As Integer = 0) As String
        W_DataBuffer(0) = data(0)
        W_DataBuffer(1) = data(1)
        If BoardCount >= 2 Then
            Result = RTBB_I2CWrite(Device_List(sel), I2CBus, ID, 1, addr, 2, W_DataBuffer(0))
        Else
            Result = RTBB_I2CWrite(hDevice, I2CBus, ID, 1, addr, 2, W_DataBuffer(0))
        End If
        Return Result
    End Function

    Function reg_write(ByVal ID As Byte, ByVal addr As Byte, ByVal data As Byte, Optional ByVal sel As Integer = 0) As String
        W_DataBuffer(0) = data
        If BoardCount >= 2 Then
            Result = RTBB_I2CWrite(Device_List(sel), I2CBus, ID, 1, addr, 1, W_DataBuffer(0))
        Else
            Result = RTBB_I2CWrite(hDevice, I2CBus, ID, 1, addr, 1, W_DataBuffer(0))
        End If
        Return Result
    End Function

    Function reg_write_word(ByVal ID As Byte, ByVal addr As Byte, ByVal data As Integer, ByVal H_L As String, Optional ByVal sel As Integer = 0) As Integer
        Dim i2c_error As Integer
        Dim w_data(1) As Byte
        Dim DataBuffer(2) As Byte

        Main.status_error.Text = ""
        w_data = word2byte_data(H_L, data)
        DataBuffer(0) = w_data(0)
        DataBuffer(1) = w_data(1)

        If BoardCount > 1 Then
            i2c_error = RTBB_I2CWrite(Device_List(device_sel), I2CBus, ID, 1, addr, 2, DataBuffer(0))
        Else
            i2c_error = RTBB_I2CWrite(hDevice, I2CBus, ID, 1, addr, 2, DataBuffer(0))
        End If

        '設定 register data 為"data", I2C data to be written to device
        If i2c_error = 0 Then
            Main.status_error.Text = "I2C Write Success!"
        Else
            MsgBox("I2C Write Error:" & i2c_status(i2c_error), MsgBoxStyle.Exclamation, "Error Message")
            Main.status_error.Text = "I2C Write Error:" & i2c_status(i2c_error)
            run = False
        End If
        Return i2c_error
    End Function

    Function reg_read_word(ByVal ID As Byte, ByVal addr As Byte, Optional ByVal sel As Integer = 0) As Integer()
        Dim return_data(1) As Integer
        If BoardCount >= 2 Then
            Result = RTBB_I2CRead(Device_List(sel), I2CBus, ID, 1, addr, 2, R_DataBuffer(0))
        Else
            Result = RTBB_I2CRead(hDevice, I2CBus, ID, 1, addr, 2, R_DataBuffer(0))
        End If
        return_data(0) = Result
        If Result = RT_BB_SUCCESS Then
            return_data(1) = R_DataBuffer(0) * 2 ^ 8 + R_DataBuffer(1)
        End If
        Return return_data
    End Function

    Function reg_read(ByVal ID As Byte, ByVal addr As Byte) As Integer()
        Dim return_data(1) As Integer
        Result = RTBB_I2CRead(hDevice, I2CBus, ID, 1, addr, 1, R_DataBuffer(0))
        return_data(0) = Result
        If Result = RT_BB_SUCCESS Then
            return_data(1) = R_DataBuffer(0)
        End If
        Return return_data
    End Function

    Function reg_read_word(ByVal ID As Byte, ByVal addr As Byte, ByVal H_L As String, Optional ByVal sel As Integer = 0) As Integer()

        Dim i2c_error As Integer
        Dim return_data(1) As Integer
        Dim DataBuffer(2) As Byte
        Main.status_error.Text = ""


        If BoardCount > 1 Then
            i2c_error = RTBB_I2CRead(Device_List(sel), I2CBus, ID, 1, addr, 2, DataBuffer(0))
        Else
            i2c_error = RTBB_I2CRead(hDevice, I2CBus, ID, 1, addr, 2, DataBuffer(0))
        End If

        return_data(0) = i2c_error
        If i2c_error = 0 Then

            If H_L = "H" Then
                'DataBuffer(0):LOW DATA BYTE
                'DataBuffer(1):HIGH DATA BYTE
                return_data(1) = DataBuffer(0) * (2 ^ 8) + DataBuffer(1)
            Else
                return_data(1) = DataBuffer(1) * (2 ^ 8) + DataBuffer(0)
            End If
            Main.status_error.Text = "I2C Read Success!"
        Else
            MsgBox("I2C Read Error:" & i2c_status(i2c_error), MsgBoxStyle.Exclamation, "Error Message")
            Main.status_error.Text = "I2C Read Error:" & i2c_status(i2c_error)
            run = False
        End If
        Return return_data
    End Function

End Module
