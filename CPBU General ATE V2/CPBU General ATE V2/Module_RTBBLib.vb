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
            I2CSetFrequency(1024, 1000)
        End If



    End Function

    Function I2CSetFrequency(ByVal nMode As Integer, ByVal nFrekHz As Integer) As Integer
        RTBB_I2CSetFrequency(hDevice, I2CBus, nMode, nFrekHz)
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


        Result = RTBB_I2CScanSlaveDevice(hDevice, I2CBus, pI2CAvailableAddressVBarray(0))



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

    Function GPIO_out(ByVal bits As Integer, ByVal bit_value() As Integer) As Integer
        Dim i As Integer

        For i = 0 To bits - 1
            If bit_value(i) = 0 Then
                RTBB_GPIOSingleWrite(hDevice, 32 + i, False) '0
            Else
                RTBB_GPIOSingleWrite(hDevice, 32 + i, True) '1
            End If

        Next

      

    End Function

    Function GPIO_single_write(ByVal bit As Integer, ByVal value As Integer) As Integer

        If value = 0 Then
            RTBB_GPIOSingleWrite(hDevice, 32 + bit, False) '0
        Else
            RTBB_GPIOSingleWrite(hDevice, 32 + bit, True) '1
        End If

    End Function

    Function GPIO_single_read(ByVal bit As Integer) As Integer
        Dim bValue As Boolean

        RTBB_GPIOSingleRead(hDevice, 32 + bit, bValue)

        If bValue = True Then
            Return 1
        Else
            Return 0
        End If




    End Function

    Function reg_write_multi(ByVal ID As Byte, ByVal addr As Byte, ByVal data() As Byte) As String
        Dim num As Integer
        Dim i As Integer

        num = data.Length

        For i = 0 To num - 1
            W_DataBuffer(i) = data(i)
        Next

        Result = RTBB_I2CWrite(hDevice, I2CBus, ID, 1, addr, num, W_DataBuffer(0))
        If Result <> 0 Then
            Result = RTBB_I2CWrite(hDevice, I2CBus, ID, 1, addr, num, W_DataBuffer(0))
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


    Function reg_write_word(ByVal ID As Byte, ByVal addr As Byte, ByVal data() As Byte) As String

        W_DataBuffer(0) = data(0)
        W_DataBuffer(1) = data(1)
        Result = RTBB_I2CWrite(hDevice, I2CBus, ID, 1, addr, 2, W_DataBuffer(0))

        Return Result

    End Function
    Function reg_write(ByVal ID As Byte, ByVal addr As Byte, ByVal data As Byte) As String

        W_DataBuffer(0) = data
        Result = RTBB_I2CWrite(hDevice, I2CBus, ID, 1, addr, 1, W_DataBuffer(0))

        Return Result

    End Function

    Function reg_read_word(ByVal ID As Byte, ByVal addr As Byte) As Integer()


        Dim return_data(1) As Integer



        Result = RTBB_I2CRead(hDevice, I2CBus, ID, 1, addr, 2, R_DataBuffer(0))

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

End Module
