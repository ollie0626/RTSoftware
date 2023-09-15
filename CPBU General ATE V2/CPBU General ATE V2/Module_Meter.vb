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
                        ts = "SENS:CURR:RANG:AUTO ON"
                        range = 3
                        'ts = "SENS:CURR:RANG " & range
                End Select

                ilwrt(Meter_Dev, ts, CInt(Len(ts)))

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
        If meter_name = "DMM6500" Then
            ts = "MEAS:" & Meter_function & ":DC? "
            ilwrt(Meter_Dev, ts, CInt(Len(ts)))
            ilrd(Meter_Dev, ValueStr, ARRAYSIZE)

            If (iberr = 0) And (ibcntl > 0) Then
                temp = ibcntl - 1
                test = Val(Mid(ValueStr, 1, temp))
            End If

        Else
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

                ilwrt(Meter_Dev, "CURR:DC:FILT:DIG ON", CInt(Len(ts)))

                ts = "MEAS:" & Meter_function & ":DC? " & range & "," & Meter_Resolution

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


                                    ts = "MEAS:" & Meter_function & ":DC? " & range & "," & Meter_Resolution

                                    ilwrt(Meter_Dev, ts, CInt(Len(ts)))
                                    Meter_range_now = range
                                    Meter_change_range = True
                                    Exit Function

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


                                    ts = "MEAS:" & Meter_function & ":DC? " & range & "," & Meter_Resolution

                                    ilwrt(Meter_Dev, ts, CInt(Len(ts)))
                                    Meter_range_now = range
                                    Meter_change_range = True
                                    Exit Function

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


                            ts = "MEAS:" & Meter_function & ":DC? " & range & "," & Meter_Resolution

                            ilwrt(Meter_Dev, ts, CInt(Len(ts)))
                            Delay(100)
                        End If

                        ilrd(Meter_Dev, ValueStr, ARRAYSIZE)
                    End If

                Next

            End If
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

End Module
