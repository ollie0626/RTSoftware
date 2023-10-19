Module Module_Meter
    Public Meter_name() As String
    Public Meter_addr() As String
    Public Meter_num As Integer


    Public Meter_Resolution As String = "DEF"
    Public Meter_iin_range As String
    Public Meter_iout_range As String
    Public Meter_icc_range As String

    'Public Meter_range(1) As String


    Public Meter_range_now As String
    Public Meter_change_range As Boolean = False

    Public Meter_iin_low As String = "4e-1" '400mA
    Public Meter_iout_low As String = "4e-1" '400mA
    Public CURRENT As String = "CURRent"
    Public VOLTAGE As String = "VOLTage"
    Public Meter_function As String = CURRENT


    'Public Meter_fail As Boolean



    Function meter_unit(ByVal cbox As Object, ByVal test_channel As Integer) As Integer
        Dim unit(6) As String

        cbox.Items.Clear()

        If test_channel <> 0 Then
            cbox.Items.Add("CURRent:DC 400mA")
            cbox.Items.Add("CURRent:DC 10A")
            cbox.Items.Add("VOLTage:DC DEF")


        Else
            cbox.Items.Add("N/A")
        End If

        cbox.SelectedIndex = 0


    End Function


    Function meter_config(ByVal meter_name As String, ByVal Meter_Dev As Integer, ByVal range As String) As String
        Dim ts As String = ""

        ' DMM4050: range_set = "MAX" (10A) , range_set = "DEF" (400mA) 'MIN會顯示最小的檔位
        ' 34450A: range_set = "AUTO"

        If meter_name = "DMM6500" Then

            If Meter_function = VOLTAGE Then
                ts = "FUNC ""VOLT"""
                ilwrt(Meter_Dev, ts, CInt(Len(ts)))
                'ts = "SENS:VOLT:RANG " & range
                ts = "SENS:VOLT:RANG:AUTO ON"
            Else
                ts = "FUNC ""CURR"""
                ilwrt(Meter_Dev, ts, CInt(Len(ts)))

                Select Case range
                    Case "MAX"
                        ts = "CURR:RANG 10"
                        ilwrt(Meter_Dev, ts, CInt(Len(ts)))
                        ts = "SENS:CURR:RANG 10"

                    Case "4e-1"
                        ts = "CURR:RANG 3"
                        ilwrt(Meter_Dev, ts, CInt(Len(ts)))
                        ts = "SENS:CURR:RANG " & range
                End Select



                '    ts = "SENS:CURR:RANG:AUTO ON"
            End If


        Else
            ts = "SYST:BEEP:STAT OFF"

            ilwrt(Meter_Dev, ts, CInt(Len(ts)))

            ts = "SYST:ERR:BEEP OFF"

            ilwrt(Meter_Dev, ts, CInt(Len(ts)))


            ts = "CONFigure:" & Meter_function & ":DC " & range


            ilwrt(Meter_Dev, ts, CInt(Len(ts)))

        End If



        Delay(100)

        'Return range

    End Function


    Function meter_scan_init(ByVal Meter_Dev As Integer, ByVal count As Integer, ByVal timer As Double) As Integer
        Dim ts As String = ""

        ts = "SAMP:COUN 1"
        ilwrt(Meter_Dev, ts, CInt(Len(ts)))


        ts = "TRIG:DEL " & timer / 1000

        ilwrt(Meter_Dev, ts, CInt(Len(ts)))


        ts = "TRIG:COUN " & count

        ilwrt(Meter_Dev, ts, CInt(Len(ts)))

        ts = "CALC:FUNC AVERage"
        ilwrt(Meter_Dev, ts, CInt(Len(ts)))


    End Function


    Function meter_scan(ByVal Meter_Dev As Integer, ByVal ONOFF As String) As Integer
        Dim ts As String = ""

        If ONOFF = "ON" Then

            ts = "CALC:STAT ON"

            ilwrt(Meter_Dev, ts, CInt(Len(ts)))


            ts = "INIT"

            ilwrt(Meter_Dev, ts, CInt(Len(ts)))
        Else

            ts = "CALC:STAT OFF"

            ilwrt(Meter_Dev, ts, CInt(Len(ts)))
        End If



    End Function

    Function meter_scan_read(ByVal Meter_Dev As Integer, ByVal calculate As String) As Double
        Dim ts As String = ""
        Dim value As Double 'MIN, AVER, MAX


        Select Case calculate
            Case "MIN"
                ts = "CALCulate:AVERage:MINimum?"

            Case "AVER"

                ts = "CALCulate:AVERage:AVERage?"

            Case "MAX"
                ts = "CALCulate:AVERage:MAXimum?"

        End Select


        ilwrt(Meter_Dev, ts, CInt(Len(ts)))

        ibcntl = 0

        While ibcntl = 0

            ilrd(Meter_Dev, ValueStr, ARRAYSIZE)

            If iberr <> EABO Then
                Exit While
            End If

            System.Windows.Forms.Application.DoEvents()

            If run = False Then
                Exit While
            End If

        End While

        'ilrd(Meter_Dev, ValueStr, ARRAYSIZE)


        If ibcntl > 0 Then
            value = Val(Mid(ValueStr, 1, (ibcntl - 1)))
        Else
            value = 0
        End If

        Return value


    End Function





    Function filter_set(ByVal Meter_Dev As Integer, ByVal digital As String, ByVal analog As String) As Integer
        Dim ts As String = ""

        ts = "FILT:DIG " & digital
        ilwrt(Meter_Dev, ts, CInt(Len(ts)))

        ts = "FILT " & analog   'FILT ON
        ilwrt(Meter_Dev, ts, CInt(Len(ts)))

    End Function

    Function meter_read(ByVal Meter_Dev As Integer) As Double

        Dim temp As Integer
        Dim test As Double = 0
        Dim i As Integer
        ts = "MEAS:CURRent:DC? "

        ilwrt(Meter_Dev, ts, CInt(Len(ts)))
        ilrd(Meter_Dev, ValueStr, ARRAYSIZE)

        For i = 0 To 10
            System.Windows.Forms.Application.DoEvents()
            If iberr = 0 Then

                If (ibcntl > 0) Then
                    temp = ibcntl - 1
                    test = Val(Mid(ValueStr, 1, temp))

                    If test <> 0 Then
                        Exit For
                    End If
                End If

            End If

            If i = 5 Then
                ilwrt(Meter_Dev, ts, CInt(Len(ts)))
                Delay(100)
            End If

            ilrd(Meter_Dev, ValueStr, ARRAYSIZE)

        Next

        Return test
    End Function


    Function meter_meas(ByVal meter_name As String, ByVal Meter_Dev As Integer, ByVal range As String, ByVal mini_range As String) As Double


        Dim temp As Integer
        Dim test As Double = 0
        Dim i As Integer

        Dim down As Boolean = False



        If range = "AUTO" Then
            ts = "MEAS:CURRent:DC? " & range & "," & Meter_Resolution

            ilwrt(Meter_Dev, ts, CInt(Len(ts)))
            ilrd(Meter_Dev, ValueStr, ARRAYSIZE)

            For i = 0 To 10
                System.Windows.Forms.Application.DoEvents()
                If iberr = 0 Then

                    If (ibcntl > 0) Then
                        temp = ibcntl - 1
                        test = Val(Mid(ValueStr, 1, temp))

                        If test <> 0 Then
                            Exit For
                        End If
                    End If

                End If

                If i = 5 Then
                    ilwrt(Meter_Dev, ts, CInt(Len(ts)))
                    Delay(100)
                End If

                ilrd(Meter_Dev, ValueStr, ARRAYSIZE)

            Next
        Else

            If meter_name = "DMM6500" Then
                ts = "MEAS:" & Meter_function & ":DC? "
            Else
                ilwrt(Meter_Dev, "CURR:DC:FILT:DIG ON", CInt(Len(ts)))

                ts = "MEAS:" & Meter_function & ":DC? " & range & "," & Meter_Resolution

            End If

            ilwrt(Meter_Dev, ts, CInt(Len(ts)))
            ilrd(Meter_Dev, ValueStr, ARRAYSIZE)

            For i = 0 To 10

                System.Windows.Forms.Application.DoEvents()

                Meter_change_range = False

                If (iberr = 0) And (ibcntl > 0) Then


                    temp = ibcntl - 1
                    test = Val(Mid(ValueStr, 1, temp))

                    If range = "MAX" Then

                        Meter_range_now = range
                    Else

                        If test > (10 ^ 10) Then
                            'range 過小
                            While test > (10 ^ 10)
                                System.Windows.Forms.Application.DoEvents()

                                If run = False Then
                                    Exit While
                                End If

                                Select Case range

                                    Case "1e-4" '100uA
                                        range = "1e-3"
                                    Case "1e-3" '1mA
                                        range = "1e-2"
                                    Case "1e-2" '10mA
                                        range = "1e-1"
                                    Case "1e-1"  '100mA
                                        range = "4e-1"

                                End Select
                                If meter_name = "DMM6500" Then
                                    ts = "MEAS:" & Meter_function & ":DC? "
                                Else

                                    ts = "MEAS:" & Meter_function & ":DC? " & range & "," & Meter_Resolution

                                End If

                                ilwrt(Meter_Dev, ts, CInt(Len(ts)))
                                Meter_range_now = range
                                Meter_change_range = True
                                Exit Function



                                'ilrd(Meter_Dev, ValueStr, ARRAYSIZE)
                                'If (iberr = 0) And (ibcntl > 0) Then
                                '    temp = ibcntl - 1
                                '    test = Val(Mid(ValueStr, 1, temp))
                                '    Meter_range_now = range
                                'End If
                                'Delay(100)
                            End While

                        Else

                            down = False


                            '過大
                            If (range = "1e-3") And test < (100 * 10 ^ -6) Then
                                If mini_range <> range Then
                                    '100uA
                                    range = "1e-4"
                                    down = True
                                End If


                            ElseIf (range = "1e-2") And test < (1 * 10 ^ -3) Then
                                '1mA
                                If mini_range <> range Then
                                    range = "1e-3"
                                    down = True
                                End If

                            ElseIf (range = "1e-1") And test < (10 * 10 ^ -3) Then
                                '10mA
                                If mini_range <> range Then
                                    range = "1e-2"
                                    down = True
                                End If

                            ElseIf (range = "4e-1") And (test < (100 * 10 ^ -3)) Then
                                '100mA
                                range = "1e-1"
                                down = True
                            End If

                            If down = True Then
                                If meter_name = "DMM6500" Then
                                    ts = "MEAS:" & Meter_function & ":DC? "
                                Else

                                    ts = "MEAS:" & Meter_function & ":DC? " & range & "," & Meter_Resolution

                                End If

                                ilwrt(Meter_Dev, ts, CInt(Len(ts)))
                                Meter_range_now = range
                                Meter_change_range = True
                                Exit Function

                                'ilrd(Meter_Dev, ValueStr, ARRAYSIZE)
                                'If (iberr = 0) And (ibcntl > 0) Then
                                '    temp = ibcntl - 1
                                '    test = Val(Mid(ValueStr, 1, temp))
                                '    Meter_range_now = range
                                'End If

                            End If



                        End If
                    End If




                    If test <> 0 Then
                        Exit For
                    End If

                Else

                    If i = 5 Then

                        GPIB_reset(Meter_Dev)
                        Delay(100)
                        If meter_name = "DMM6500" Then
                            ts = "MEAS:" & Meter_function & ":DC? "
                        Else

                            ts = "MEAS:" & Meter_function & ":DC? " & range & "," & Meter_Resolution

                        End If
                        ilwrt(Meter_Dev, ts, CInt(Len(ts)))
                        Delay(100)
                    End If

                    ilrd(Meter_Dev, ValueStr, ARRAYSIZE)
                End If

            Next

        End If


        Return test

    End Function

    Function meter_average(ByVal meter_name As String, ByVal Meter_Dev As Integer, ByVal average As Integer, ByVal range As String, ByVal mini_range As String) As Double
        Dim i As Integer
        Dim temp, total As Double
        Dim error_num As Integer = 5



        Meter_range_now = range
        For i = 1 To average
            System.Windows.Forms.Application.DoEvents()

            If run = False Then
                Exit For
            End If

            temp = meter_meas(meter_name, Meter_Dev, Meter_range_now, mini_range)

            If Meter_change_range = True Then
                'Select Case sense_vin_test
                '    Case "Test1"
                '        PartI.Sense_vin()
                'End Select
                i = 1
                temp = meter_meas(meter_name, Meter_Dev, Meter_range_now, mini_range)
            End If

            While temp = 0

                System.Windows.Forms.Application.DoEvents()

                If run = False Then
                    Exit While
                End If

                temp = meter_meas(meter_name, Meter_Dev, Meter_range_now, mini_range)
                Delay(10)

                If error_num = 0 Then
                    average = i
                    Exit For
                Else
                    error_num = error_num - 1
                End If
            End While

            If i = 1 Then
                total = temp
            Else
                total = total + temp
            End If
        Next
        If average > 0 Then
            total = total / average
        Else
            total = 0
        End If
        Return total
    End Function

    Function realy_out_meter_initial() As Integer
        reg_write_word(out_io_id, &H3, &H0, "H", device_sel)
        reg_write_word(out_io_id, &H1, &H0, "H", device_sel) '切換最大檔位

        ' write comp value to ic 
        reg_write_word(out_high_id, &H5, out_high_comp, "H", device_sel)
        reg_write_word(out_middle_id, &H5, out_middle_comp, "H", device_sel)
        reg_write_word(out_low_id, &H5, out_low_comp, "H", device_sel)
        Return 0
    End Function

    Function relay_in_meter_intial() As Integer
        reg_write_word(in_io_id, &H3, &H0, "H", device_sel)
        reg_write_word(in_io_id, &H1, &H0, "H", device_sel) '切換最大檔位

        ' write comp value to ic 
        reg_write_word(in_high_id, &H5, in_high_comp, "H", device_sel)
        reg_write_word(in_middle_id, &H5, in_middle_comp, "H", device_sel)
        reg_write_word(in_low_id, &H5, in_low_comp, "H", device_sel)
        Return 0
    End Function

    Function meter_auto(ByVal in_out_sel As Integer, ByVal average As Integer) As Double

        Dim total As Double = 0
        Dim curr_data As Double
        Dim Meas_ID As Integer
        Dim IO_ID As Integer
        Dim data_input As Byte
        Dim read_error As Integer
        Dim resolution As Double

        Select Case in_out_sel
            Case 0 : curr_data = power_read(vin_device, Vin_out, "CURR")
            Case 1 : curr_data = load_read("CURR")
        End Select



        If DUT2_en Then
            If curr_data >= Meter_H Then
                Select Case in_out_sel
                    Case 0 : Meas_ID = in_high_id2 : resolution = in_high_resolution
                    Case 1 : Meas_ID = out_high_id : resolution = out_high_resolution
                End Select

                data_input = &H0
            End If

            If curr_data < Meter_H And curr_data >= Meter_L Then

                Select Case in_out_sel
                    Case 0 : Meas_ID = in_middle_id2 : resolution = in_middle_resolution
                    Case 1 : Meas_ID = out_middle_id : resolution = out_middle_resolution
                End Select

                data_input = &H2
            End If

            If curr_data < Meter_L Then
                Select Case in_out_sel
                    Case 0 : Meas_ID = in_low_id2 : resolution = in_low_resolution
                    Case 1 : Meas_ID = out_low_id : resolution = out_low_resolution
                End Select
                data_input = &H1
            End If

            Select Case in_out_sel
                Case 0 : IO_ID = in_io_id2
                Case 1 : IO_ID = out_io_id
            End Select
        Else
            If curr_data >= Meter_H Then
                Select Case in_out_sel
                    Case 0 : Meas_ID = in_high_id : resolution = in_high_resolution
                    Case 1 : Meas_ID = out_high_id : resolution = out_high_resolution
                End Select

                data_input = &H0
            End If

            If curr_data < Meter_H And curr_data >= Meter_L Then

                Select Case in_out_sel
                    Case 0 : Meas_ID = in_middle_id : resolution = in_middle_resolution
                    Case 1 : Meas_ID = out_middle_id : resolution = out_middle_resolution
                End Select

                data_input = &H2
            End If

            If curr_data < Meter_L Then
                Select Case in_out_sel
                    Case 0 : Meas_ID = in_low_id : resolution = in_low_resolution
                    Case 1 : Meas_ID = out_low_id : resolution = out_low_resolution
                End Select
                data_input = &H1
            End If

            Select Case in_out_sel
                Case 0 : IO_ID = in_io_id
                Case 1 : IO_ID = out_io_id
            End Select
        End If

        ' H: write Hi byte first then low byte
        reg_write_word(IO_ID, &H1, data_input, "H", device_sel)
        Delay(1000)

        Dim array As List(Of Double) = New List(Of Double)()
        Dim remove_data As Integer
        Dim temp() As Integer
        Dim iout_temp As Double
        total = 0
        read_error = 0

        For i = 0 To (average - 1)
            System.Windows.Forms.Application.DoEvents()
            If run = False Then
                Exit For
            End If
            temp = reg_read_word(Meas_ID, &H4, "H", device_sel)
            While temp(0) <> 0 Or temp(1) = 65535
                System.Windows.Forms.Application.DoEvents()
                read_error = read_error + 1
                If (read_error = 5) Or (run = False) Then
                    Return 0
                    Exit Function
                End If
                Delay(10)
                temp = reg_read_word(Meas_ID, &H4, "H", device_sel)
            End While
            Delay(10)
            iout_temp = temp(1) * resolution * 10 ^ -3

            If i >= remove_data Then
                array.Add(iout_temp)
            End If
        Next
        total = array.Sum() / array.Count
        Return total
    End Function

End Module
