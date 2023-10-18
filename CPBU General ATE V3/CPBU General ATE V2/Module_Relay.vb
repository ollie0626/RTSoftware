Module Module_Relay
    '//-----------------------------------------------------------------------------//
    'Relay
    'MCP23017 Bank0
    'IODIRA: Reg Addr=0x00 ,0=Output, 1=Input
    'GPIOA:  Reg Addr=0x12, 0=Low (OFF) , 1=High (ON)
    'IODIRB: Reg Addr=0x01
    'GPIOB:  Reg Addr=0x13

    Public Relay_ID() As Byte = {&H21, &H27}

    Public Relay_signal_A() As String = {"S_Relay1" & vbNewLine & "S_Relay2" & vbNewLine & "S_Relay3" & vbNewLine & "S_Relay4" & vbNewLine & "S_Relay5" & vbNewLine & "S_Relay6" & vbNewLine & "S_Relay7" & vbNewLine & "S_Relay8", _
                                         "S_Relay1" & vbNewLine & "S_Relay2" & vbNewLine & "S_Relay3" & vbNewLine & "S_Relay4" & vbNewLine & "S_Relay5" & vbNewLine & "S_Relay6" & vbNewLine & "S_Relay7" & vbNewLine & "S_Relay8"}
    Public Relay_signal_B() As String = {"S_PSURelay" & vbNewLine & "S_EFFRelay" & vbNewLine & "S_ShuntRelay" & vbNewLine & "S_function1" & vbNewLine & "S_function2" & vbNewLine & "S_Load1" & vbNewLine & "S_Load2" & vbNewLine & "S_Load3", _
                                         "S_PSURelay" & vbNewLine & "S_EFFRelay" & vbNewLine & "S_ShuntRelay" & vbNewLine & "S_function1" & vbNewLine & "S_function2" & vbNewLine & "S_Load1" & vbNewLine & "S_Load2" & vbNewLine & "S_Load3"}
    Public Relay_ID_check(Relay_ID.Length - 1) As Boolean

    Public Relay_dat_A() As Byte = {&H0, &H0}
    Public Relay_dat_B() As Byte = {&H0, &H0}


    Dim Relay_IODIRA As Byte = &H0
    Dim Relay_IODIRB As Byte = &H1
    Dim Relay_GPIOA As Byte = &H12
    Dim Relay_GPIOB As Byte = &H13
    Dim IODIR_OUT As Byte = &H0
    Dim IODIR_IN As Byte = &HF


    Dim S_PSURelay_ON As Boolean
    Dim S_EFFRelay_ON As Boolean
    Dim S_ShuntRelay_ON As Boolean
    Dim S_function1_ON As Boolean
    Dim S_function2_ON As Boolean
    Dim S_Load1_ON As Boolean
    Dim S_Load2_ON As Boolean
    Dim S_Load3_ON As Boolean


    'CH1
    'S_Relay1= Buck1_cap or VIN_C
    'S_Relay2= Islammer or SMBalert
    'S_Relay3= INT/WDT/RST/SUSPEND or SDA



    Public Relay1_BUCK1_VIN As Boolean = False 'bit0
    Public Relay2_Islammer_SMBalert As Boolean = False 'bit1
    Public Relay3_CH1_Other As Boolean = False 'bit2

    'CH2
    'S_Relay4= VCC_EVB or Buck2_cap
    'S_Relay5= VIN_C or SCL
    'S_Relay5= VEN_EVB or Ctrl


    Public Relay4_VCC_BUCK2 As Boolean = False 'bit3
    Public Relay5_VIN_SCL As Boolean = False 'bit4
    Public Relay6_VEN_Ctrl As Boolean = False 'bit5

    'CH4
    'S_Relay7= VSS_EVB or Islammer
    'S_Relay8= PG_EVB or MODE/SYNC


    Public Relay7_Islammer_VSS As Boolean = False
    Public Relay8_PG_MODE As Boolean = False
  

    'INA226
    ',0x40,0x41,0x42,0x43,0x44,0x45,
    Public Meas_VIN_P_ID As Byte = &H40
    Public Meas_VIN_N_ID As Byte = &H41
    Public Meas_BUCK1_CH1_ID As Byte = &H42
    Public Meas_BUCK1_CH2_ID As Byte = &H43
    Public Meas_BUCK2_CH1_ID As Byte = &H44
    Public Meas_BUCK2_CH2_ID As Byte = &H45


    Public Meas_ID() As Integer = {Meas_VIN_P_ID, Meas_VIN_N_ID, Meas_BUCK1_CH1_ID, Meas_BUCK1_CH2_ID, Meas_BUCK2_CH1_ID, Meas_BUCK2_CH2_ID}
    Public Meas_signal() As String = {"Iin1", "Iin2", "Iout1_CH1", "Iout1_CH2", "Iout2_CH1", "Iout2_CH2"}
    Public Meas_ID_check(Meas_ID.Length - 1) As Boolean
    Public Meas_Addr As Byte = &H4
    Public INA226_Iout_max_L As Double = 0.08 '80mA
    Public INA226_Iout_max_H As Double = 4 '4A
    Public INA226_Iin_max_L As Double = 0.08 '80mA
    Public INA226_Iin_max_H As Double = 4 '4A

    Public Iout_board_EN As Double = False

    Dim resolution_Iin_L As Double
    Dim resolution_Iin_H As Double
    Dim resolution_Iout_L As Double
    Dim resolution_Iout_H As Double

    Public INA226_config_data As Integer = &H4127


  

    Function INA226_IIN_set() As Double
        If (iout_now < iin_meter_change) And (Iin_Meter_Max = True) Then

            '切小檔位

            DCLoad_ONOFF("OFF")
            relay_IIN_set(True, True)
            Iin_Meter_Max = False

        ElseIf (iout_now >= iin_meter_change) And (Iin_Meter_Max = False) Then
            '切大檔位

            DCLoad_ONOFF("OFF")
            relay_IIN_set(False, True)
            Iin_Meter_Max = True


        End If


        If DCLoad_ON = False Then
            DCLoad_ONOFF("ON")
            Delay(100)
        End If
    End Function

  

    Function INA226_Iin_initial(ByVal H_range As Boolean) As Integer


        If H_range = True Then
            '保持在高檔位
            relay_IIN_set(False, True)
            Iin_Meter_Max = True
        Else
            '保持在低檔位
            relay_IIN_set(True, True)
            Iin_Meter_Max = False
        End If
 

    End Function

    Function INA226_IOUT_meas(ByVal buck_num As Integer, ByVal average As Integer) As Double
        Dim i, ii As Integer
        Dim total As Double
        Dim temp() As Integer
        Dim error_num As Integer = 5
        Dim Meas_ID As Integer
        Dim resolution As Double
        Dim iout_temp(1) As Double

        If buck_num = 3 Then

            For i = 1 To average

                System.Windows.Forms.Application.DoEvents()

                If run = False Then
                    Exit For
                End If


                For ii = 0 To 1
                    'CH1+CH3
                    If Iout_Meter_Max = True Then

                        If ii = 0 Then
                            Meas_ID = Meas_BUCK1_CH2_ID
                        Else
                            Meas_ID = Meas_BUCK2_CH2_ID
                        End If

                        resolution = resolution_Iout_H

                    Else
                        If ii = 0 Then
                            Meas_ID = Meas_BUCK1_CH1_ID
                        Else
                            Meas_ID = Meas_BUCK2_CH1_ID
                        End If
                        resolution = resolution_Iout_L
                    End If


                    temp = reg_read_word(Meas_ID, Meas_Addr, device_sel)




                    While temp(0) <> 0

                        System.Windows.Forms.Application.DoEvents()

                        If run = False Then
                            Exit While
                        End If

                        temp = reg_read_word(Meas_ID, Meas_Addr, device_sel)
                        iout_temp(ii) = temp(1)
                        Delay(10)

                        If error_num = 0 Then
                            average = i
                            Exit For
                        Else
                            error_num = error_num - 1
                        End If
                    End While

                Next


                If i = 1 Then
                    total = iout_temp(0) + iout_temp(1)
                Else
                    total = total + iout_temp(0) + iout_temp(1)
                End If


            Next


        Else
            If Iout_Meter_Max = True Then

                If buck_num = 1 Then
                    Meas_ID = Meas_BUCK1_CH2_ID
                Else
                    Meas_ID = Meas_BUCK2_CH2_ID
                End If

                resolution = resolution_Iout_H

            Else
                If buck_num = 1 Then
                    Meas_ID = Meas_BUCK1_CH1_ID
                Else
                    Meas_ID = Meas_BUCK2_CH1_ID
                End If
                resolution = resolution_Iout_L
            End If




            For i = 1 To average
                System.Windows.Forms.Application.DoEvents()

                If run = False Then
                    Exit For
                End If

                temp = reg_read_word(Meas_ID, Meas_Addr, device_sel)




                While temp(0) <> 0

                    System.Windows.Forms.Application.DoEvents()

                    If run = False Then
                        Exit While
                    End If

                    temp = reg_read_word(Meas_ID, Meas_Addr, device_sel)

                    Delay(10)

                    If error_num = 0 Then
                        average = i
                        Exit For
                    Else
                        error_num = error_num - 1
                    End If
                End While




                If i = 1 Then
                    total = temp(1)
                Else
                    total = total + temp(1)
                End If

            Next

        End If




        If average > 0 Then
            total = total / average
        Else
            total = 0
        End If

        Return total * resolution 'A

    End Function

    Function INA226_IIN_meas(ByVal average As Integer) As Double
        Dim i As Integer
        Dim total As Double
        Dim temp() As Integer
        Dim error_num As Integer = 5
        Dim Meas_ID As Integer
        Dim resolution As Double

        If Iin_Meter_Max = True Then

            Meas_ID = Meas_VIN_N_ID

            resolution = resolution_Iin_H
        Else

            Meas_ID = Meas_VIN_P_ID
            resolution = resolution_Iin_L
        End If




        For i = 1 To average
            System.Windows.Forms.Application.DoEvents()

            If run = False Then
                Exit For
            End If

            temp = reg_read_word(Meas_ID, Meas_Addr, device_sel)




            While temp(0) <> 0

                System.Windows.Forms.Application.DoEvents()

                If run = False Then
                    Exit While
                End If

                temp = reg_read_word(Meas_ID, Meas_Addr, device_sel)

                Delay(10)

                If error_num = 0 Then
                    average = i
                    Exit For
                Else
                    error_num = error_num - 1
                End If
            End While




            If i = 1 Then
                total = temp(1)
            Else
                total = total + temp(1)
            End If

        Next







        If average > 0 Then
            total = total / average
        Else
            total = 0
        End If

        Return total * resolution  'A

    End Function

    Function INA226_config() As Integer
        Dim config_addr As Byte = &H0
        Dim data() As Byte


        data = word2byte_data("H", INA226_config_data)
        reg_write_word(Meas_VIN_P_ID, config_addr, data, device_sel)
        reg_write_word(Meas_VIN_N_ID, config_addr, data, device_sel)
        reg_write_word(Meas_BUCK1_CH1_ID, config_addr, data, device_sel)
        reg_write_word(Meas_BUCK1_CH2_ID, config_addr, data, device_sel)
        reg_write_word(Meas_BUCK2_CH1_ID, config_addr, data, device_sel)
        reg_write_word(Meas_BUCK2_CH2_ID, config_addr, data, device_sel)

        Main.txt_INA226_00h.Text = INA226_config_data.ToString("X4")
    End Function


    Function current_monitor_init() As Integer
        Dim cal_value() As Byte

        Dim cal_addr As Byte = &H5

        Dim value As Integer

      

        INA226_config()

        'IIN Low
        '      INA226_Iin_max_L = 0.08 / num_IIN_Rshunt_L.Value
        resolution_Iin_L = (INA226_Iin_max_L / (2 ^ 15))  'A

        value = 0.00512 / (Main.num_IIN_Rshunt_L.Value * resolution_Iin_L)

        cal_value = word2byte_data("H", value)

        reg_write_word(Meas_VIN_P_ID, cal_addr, cal_value, device_sel)




        'IIN High
        resolution_Iin_H = (INA226_Iin_max_H / (2 ^ 15))   'A

        value = 0.00512 / (Main.num_IIN_Rshunt_H.Value * resolution_Iin_H)

        cal_value = word2byte_data("H", value)
        reg_write_word(Meas_VIN_N_ID, cal_addr, cal_value, device_sel)


        'IoutLow
        resolution_Iout_L = (INA226_Iout_max_L / (2 ^ 15))  'A

        value = 0.00512 / (Main.num_IOUT_Rshunt_L.Value * resolution_Iout_L)

        cal_value = word2byte_data("H", value)
        reg_write_word(Meas_BUCK1_CH1_ID, cal_addr, cal_value, device_sel)
        reg_write_word(Meas_BUCK2_CH1_ID, cal_addr, cal_value, device_sel)



        'Iout High
        resolution_Iout_H = (INA226_Iout_max_H / (2 ^ 15))   'A

        value = 0.00512 / (Main.num_IOUT_Rshunt_H.Value * resolution_Iout_H)

        cal_value = word2byte_data("H", value)
        reg_write_word(Meas_BUCK1_CH2_ID, cal_addr, cal_value, device_sel)
        reg_write_word(Meas_BUCK2_CH2_ID, cal_addr, cal_value, device_sel)

    End Function

    Function relay_init() As Integer
        'MCP23017 Bank0
        'IODIRA: Reg Addr=0x00 ,0=Output, 1=Input
        'GPIOA:  Reg Addr=0x12, 0=Low (OFF) , 1=High (ON)
        'IODIRB: Reg Addr=0x01
        'GPIOB:  Reg Addr=0x13
        Dim i As Integer


        For i = 0 To Relay_ID.Length - 1
            Relay_dat_A(i) = 0
            Relay_dat_B(i) = 0
            reg_write(Relay_ID(i), Relay_IODIRA, IODIR_OUT, device_sel)
            reg_write(Relay_ID(i), Relay_IODIRB, IODIR_OUT, device_sel)

            reg_write(Relay_ID(i), Relay_GPIOA, Relay_dat_A(i), device_sel)
            reg_write(Relay_ID(i), Relay_GPIOB, Relay_dat_B(i), device_sel)


        Next

        S_PSURelay_ON = False
        S_EFFRelay_ON = False
        S_ShuntRelay_ON = False
        S_function1_ON = False
        S_function2_ON = False
        S_Load1_ON = False
        S_Load2_ON = False
        S_Load3_ON = False


        Relay1_BUCK1_VIN = False
        Relay2_Islammer_SMBalert = False 'bit1
        Relay3_CH1_Other = False 'bit2
        Relay4_VCC_BUCK2 = False 'bit3
        Relay5_VIN_SCL = False 'bit4
        Relay6_VEN_Ctrl = False 'bit5
        Relay7_Islammer_VSS = False
        Relay8_PG_MODE = False



        current_monitor_init()

    End Function



    Function relay_IIN_set(ByVal meter_L_ON As Boolean, ByVal meter_H_ON As Boolean) As Integer

        'default Meter1_ON, Meter2_ON=False


        'VIN
        S_PSURelay_ON = False


        'Meter_L
        S_ShuntRelay_ON = meter_L_ON

        'Meter H
        S_EFFRelay_ON = meter_H_ON

        relay_function_set()

    End Function

    Function relay_function_set() As Integer
        Dim i As Integer
        Dim bit(7) As Byte
        Dim data As Byte = 0

        For i = 0 To 7
            bit(i) = 0
        Next


        If S_PSURelay_ON = True Then
            bit(0) = 1
        End If

        If S_EFFRelay_ON = True Then
            bit(1) = 1
        End If
        If S_ShuntRelay_ON = True Then
            bit(2) = 1
        End If
        If S_function1_ON = True Then
            bit(3) = 1
        End If
        If S_function2_ON = True Then
            bit(4) = 1
        End If
        If S_Load1_ON = True Then
            bit(5) = 1
        End If
        If S_Load2_ON = True Then
            bit(6) = 1
        End If
        If S_Load3_ON = True Then
            bit(7) = 1
        End If

        For i = 0 To 7
            data = data + bit(i) * 2 ^ i
        Next

        Relay_dat_B(0) = data

        relay_function_set = reg_write(Relay_ID(0), Relay_GPIOB, Relay_dat_B(0), device_sel)

        Return relay_function_set



    End Function




    Function relay_Scope_set() As Integer
        Dim i As Integer
        Dim bit(7) As Byte
        Dim data As Byte = 0

        'CH1
        If Relay1_BUCK1_VIN = True Then
            bit(0) = 1
        Else
            bit(0) = 0
        End If

        'CH1
        If Relay2_Islammer_SMBalert = True Then
            bit(1) = 1
        Else
            bit(1) = 0
        End If

        'CH1
        If Relay3_CH1_Other = True Then
            bit(2) = 1
        Else
            bit(2) = 0
        End If

        'CH2
        If Relay4_VCC_BUCK2 = True Then
            bit(3) = 1
        Else
            bit(3) = 0
        End If

        'CH2
        If Relay5_VIN_SCL = True Then
            bit(4) = 1
        Else
            bit(4) = 0
        End If

        'CH2
        If Relay6_VEN_Ctrl = True Then
            bit(5) = 1
        Else
            bit(5) = 0
        End If

        'CH4
        If Relay7_Islammer_VSS = True Then
            bit(6) = 1
        Else
            bit(6) = 0
        End If

        'CH4
        If Relay8_PG_MODE = True Then
            bit(7) = 1
        Else
            bit(7) = 0
        End If



        For i = 0 To 7
            data = data + bit(i) * 2 ^ i
        Next

        Relay_dat_A(0) = data

        relay_Scope_set = reg_write(Relay_ID(0), Relay_GPIOA, Relay_dat_A(0), device_sel)

        Return relay_Scope_set

    End Function

End Module
