Module Module_Power

    '//-------------------------------------------------------------------------------------------------------------//
    '2230-30-1
    'E3631A
    'E3632A
    '6210-40
    'MODEL 2400
    'MODEL 2410
    'E36312A (2020/07/22 Add)
    '62006P
    '62012P
    'N6705
    '//----------------------------------------------------------------------------------------------------------------//



    'Public N6705_Dev As Integer
    Public Power_2230_Dev As String
    Public Power_Dev As Integer

    'Public Power_2230 As Boolean = False
    Public Power_vi As Integer

    Public Power_name() As String
    Public Power_addr() As String
    Public Power_num As Integer
    Public Power_rurd_ok() As Boolean

    Public E3632_Range_H As Boolean = False
    Public E3632_OCP As Double = 7


    Public N6705_Pulse As String = "PULSe"
    Public N6705_Ramp As String = "RAMP"
    Public N6705_Sine As String = "SINusoid"
    Public N6705_Staircase As String = "STAircase"
    Public N6705_Step As String = "STEP"
    Public N6705_Trapezoid As String = "TRAPezoid"
    Public N6705_UDEFined As String = "UDEFined"


    Public N6705_V0 As String = "V0"
    Public N6705_V1 As String = "V1"
    Public N6705_t0 As String = "t0"
    Public N6705_t1 As String = "t1"
    Public N6705_t2 As String = "t2"
    Public N6705_t3 As String = "t3"
    Public N6705_t4 As String = "t4"
    Public N6705_nst As String = "NST"

    Dim ts As String
    Dim N6705_ARRAYSIZE As Short = &HFFF
    Dim N6705_ValueStr As String = Space(N6705_ARRAYSIZE)

    Function power_channel_set(ByVal cbox_power As Object, ByVal cbox_power_ch As Object) As Integer

        cbox_power_ch.Items.Clear()




        If cbox_power.SelectedItem = no_device Then

            cbox_power_ch.Items.Add(no_device)

        Else


            Select Case cbox_power.SelectedItem
                Case " 2230-30-1"
                    cbox_power_ch.Items.Add("CH1:30V/1.5A")
                    cbox_power_ch.Items.Add("CH2:30V/1.5A")
                    cbox_power_ch.Items.Add("CH3:6V/5A")




                Case "E3631A"
                    cbox_power_ch.Items.Add("6V/5A")
                    cbox_power_ch.Items.Add("±25V/1A")

                Case "E3632A"
                    cbox_power_ch.Items.Add("15V/7A;30V/4A")

                Case "6210-40"
                    cbox_power_ch.Items.Add("40V/25A")

                Case "MODEL 2400"

                    cbox_power_ch.Items.Add("CH1")
                Case "MODEL 2410"

                    cbox_power_ch.Items.Add("CH1")


                Case "E36312A"
                    '2022/07/22 Add
                    cbox_power_ch.Items.Add("CH1:6V/5A")
                    cbox_power_ch.Items.Add("CH2:25V/1A")
                    cbox_power_ch.Items.Add("CH3:25V/1A")


            End Select

            Select Case Mid(cbox_power.SelectedItem, 1, 6)

                Case "62006P"
                    cbox_power_ch.Items.Add("100V/25A")

                Case "62012P"
                    cbox_power_ch.Items.Add("80V/60A")



            End Select

            Select Case Mid(cbox_power.SelectedItem, 1, 5)

                Case "N6705"
                    cbox_power_ch.Items.Add("CH1")
                    cbox_power_ch.Items.Add("CH2")
                    cbox_power_ch.Items.Add("CH3")
                    cbox_power_ch.Items.Add("CH4")


            End Select
        End If






        'cbox_power_ch.SelectedIndex = 0

    End Function


    Function Power_channel(ByVal device_name As String, ByVal channel As Integer) As String
        Dim out As String

        out = ""

        If device_name = "N6705" Or device_name = "2230-30-1" Or device_name = "E36312A" Then
            out = channel + 1
        ElseIf device_name = "E3631A" Then
            Select Case channel
                Case 0
                    out = "P6V"
                Case 1
                    out = "P25V"
            End Select


        Else

            out = ""
        End If


        Return out
    End Function


    Function Power2230_visa(ByVal open As Boolean) As Integer
        If open = True Then
            viOpenDefaultRM(defaultRM)
            viOpen(defaultRM, Power_2230_Dev, VI_NO_LOCK, 2000, Power_vi)
        Else
            viClose(Power_vi)
            viClose(defaultRM)
        End If


    End Function


    Function Power2230_init() As String


        ts = "SYSTem:REMote"
        visa_status = viWrite(Power_vi, ts, Len(ts), retcount)
        ts = "*RST"
        visa_status = viWrite(Power_vi, ts, Len(ts), retcount)



    End Function

    Function Power2230_ONOFF(ByVal on_off As String) As String


        ts = "OUTPUT " & on_off
        visa_status = viWrite(Power_vi, ts, Len(ts), retcount)



    End Function


    Function Power2230_set(ByVal OUT As Integer, ByVal VOLT As Double) As Integer



        ts = "INSTrument:NSELect " & OUT
        visa_status = viWrite(Power_vi, ts, Len(ts), retcount)

        ts = "VOLT " & VOLT
        visa_status = viWrite(Power_vi, ts, Len(ts), retcount)


    End Function

    Function Power2230_read(ByVal OUT As Integer, ByVal volt_curr As String) As Double

        ts = "INSTrument:NSELect " & OUT
        visa_status = viWrite(Power_vi, ts, Len(ts), retcount)


        If volt_curr = "VOLT" Then
            ts = "MEAS:VOLT?"
        Else
            ts = "MEAS:CURR?"
        End If

        visa_status = viWrite(Power_vi, ts, Len(ts), retcount)

        visa_status = viRead(Power_vi, visa_response, Len(visa_response), retcount)

        Return Val(Mid(visa_response, 1, retcount))




    End Function

    Function N6705_OUT_delay(ByVal OUT As String, ByVal delay_value As Double) As Integer
        ts = "OUTP:DELay:RISE " & delay_value & ",(@" & OUT & ")"
        ilwrt(Power_Dev, ts, CInt(Len(ts)))
    End Function

    Function N6705_OUTPUT(ByVal EN_ON As Boolean, ByVal OUT As String) As Integer


        If EN_ON = False Then

            ts = "OUTP:COUP:CHAN "
        Else
            ts = "OUTP:COUP:CHAN " & OUT
            ilwrt(Power_Dev, ts, CInt(Len(ts)))

            'ts = "OUTP:COUP:DOFF 0"
            'ilwrt(Power_Dev, ts, CInt(Len(ts)))

            'ts = "OUTP:COUP:DOFFSet:MODE MANual"


        End If

        'OUTP:COUP:CHAN 1,2,4


        ilwrt(Power_Dev, ts, CInt(Len(ts)))


    End Function



    Function N6705_SEQ_init(ByVal OUT As String, ByVal LAST As Boolean, ByVal INF As Boolean, ByVal count As Integer) As Integer

        'ARB:FUNC:TYPE VOLT,(@1)
        'ARB:FUNC:SHAP SEQ,(@1)
        'ARB:SEQ:RESet (@1)

        ts = "ARB:FUNC:TYPE VOLT" & ",(@" & OUT & ")"
        ilwrt(Power_Dev, ts, CInt(Len(ts)))

        ts = "ARB:FUNC:SHAP SEQ" & ",(@" & OUT & ")"
        ilwrt(Power_Dev, ts, CInt(Len(ts)))


        ts = "ARB:SEQ:RESet " & "(@" & OUT & ")"
        ilwrt(Power_Dev, ts, CInt(Len(ts)))




        If LAST = True Then
            'ON: Last Arb Value
            ts = "ARB:SEQ:TERM:LAST ON,(@" & OUT & ")"
        Else
            'OFF: Return to DC Value
            ts = "ARB:SEQ:TERM:LAST OFF,(@" & OUT & ")"
        End If

        ilwrt(Power_Dev, ts, CInt(Len(ts)))

        If INF = True Then
            ts = "ARB:SEQ:COUN INF " & ",(@" & OUT & ")"
        Else
            ts = "ARB:SEQ:COUN " & count & " ,(@" & OUT & ")"
        End If


        ilwrt(Power_Dev, ts, CInt(Len(ts)))

        ts = "VOLT:MODE ARB" & ",(@" & OUT & ")"

        ilwrt(Power_Dev, ts, CInt(Len(ts)))

        'Trigger source Remote Command
        ts = "TRIG:ARB:SOUR BUS"
        ilwrt(Power_Dev, ts, CInt(Len(ts)))


    End Function

    Function N6705_SEQ_end_set(ByVal OUT As String, ByVal seq_step As Integer) As Integer

        ts = "ARB:SEQ:STEP:PAC TRIG," & seq_step & ",(@" & OUT & ")"

        ilwrt(Power_Dev, ts, CInt(Len(ts)))

        ts = "TRIG:ARB:SOUR BUS," & seq_step & ",(@" & OUT & ")"

        ilwrt(Power_Dev, ts, CInt(Len(ts)))
    End Function

    Function N6705_ARB_mode(ByVal OUT As String, ByVal mode As String) As Integer

        ts = "ARB:FUNC:SHAP " & mode & ",(@" & OUT & ")"
        ilwrt(Power_Dev, ts, CInt(Len(ts)))



    End Function


    Function N6705_ARB_init(ByVal OUT As String, ByVal mode As String, ByVal LAST As Boolean, ByVal INF As Boolean, ByVal count As Integer, ByVal trigger_in As Boolean) As Integer

        ts = "ARB:FUNC:TYPE VOLT" & ",(@" & OUT & ")"
        ilwrt(Power_Dev, ts, CInt(Len(ts)))

        N6705_ARB_mode(OUT, mode)

        If LAST = True Then
            'ON: Last Arb Value
            ts = "ARB:TERM:LAST ON,(@" & OUT & ")"
        Else
            'OFF: Return to DC Value
            ts = "ARB:TERM:LAST OFF,(@" & OUT & ")"
        End If

        ilwrt(Power_Dev, ts, CInt(Len(ts)))

        If INF = True Then
            ts = "ARB:COUN INF " & ",(@" & OUT & ")"
        Else
            ts = "ARB:COUN " & count & " ,(@" & OUT & ")"
        End If


        ilwrt(Power_Dev, ts, CInt(Len(ts)))

        ts = "VOLT:MODE ARB" & ",(@" & OUT & ")"

        ilwrt(Power_Dev, ts, CInt(Len(ts)))

        'Trigger source Remote Command

        N6705_trig_set(trigger_in)


    End Function

    Function N6705_trig_set(ByVal trigger_in As Boolean) As Integer
        If trigger_in = True Then
            'trigger in
            ts = "TRIG:ARB:SOUR EXT"
        Else
            'remote command
            ts = "TRIG:ARB:SOUR BUS"
        End If

        ilwrt(Power_Dev, ts, CInt(Len(ts)))
    End Function

    Function N6705_ARB_parameter(ByVal OUT As String, ByVal mode As String, ByVal parameter As String, ByVal value As String) As Integer

        Select Case parameter

            Case N6705_V0

                If mode = N6705_Sine Then
                    ts = "ARB:VOLT:" & mode & ":AMPL " & value & ",(@" & OUT & ")"
                Else
                    ts = "ARB:VOLT:" & mode & ":STAR " & value & ",(@" & OUT & ")"

                End If

            Case N6705_V1


                If mode = N6705_Pulse Or mode = N6705_Trapezoid Then
                    ts = "ARB:VOLT:" & mode & ":TOP " & value & ",(@" & OUT & ")"

                ElseIf mode = N6705_Sine Then

                    ts = "ARB:VOLT:" & mode & ":OFFS " & value & ",(@" & OUT & ")"
                Else

                    ts = "ARB:VOLT:" & mode & ":END " & value & ",(@" & OUT & ")"

                End If

            Case N6705_t0


                If mode = N6705_Sine Then
                    'f = 1 / t0
                    ts = "ARB:VOLT:" & mode & ":FREQ " & Val(1 / value) & ",(@" & OUT & ")"
                Else
                    ts = "ARB:VOLT:" & mode & ":STAR:TIM " & value & ",(@" & OUT & ")"

                End If



            Case N6705_t1

                If mode = N6705_Pulse Then
                    ts = "ARB:VOLT:" & mode & ":TOP:TIM " & value & ",(@" & OUT & ")"
                ElseIf mode = N6705_Staircase Then
                    ts = "ARB:VOLT:" & mode & ":TIM " & value & ",(@" & OUT & ")"
                ElseIf mode = N6705_Step Then
                    ts = "ARB:VOLT:" & mode & ":END:TIM " & value & ",(@" & OUT & ")"
                Else

                    ts = "ARB:VOLT:" & mode & ":RTIM " & value & ",(@" & OUT & ")"
                End If


            Case N6705_t2

                If mode = N6705_Trapezoid Then

                    ts = "ARB:VOLT:" & mode & ":TOP:TIM " & value & ",(@" & OUT & ")"

                Else
                    ts = "ARB:VOLT:" & mode & ":END:TIM " & value & ",(@" & OUT & ")"
                End If


            Case N6705_t3
                ts = "ARB:VOLT:" & mode & ":FTIM " & value & ",(@" & OUT & ")"
            Case N6705_t4

                ts = "ARB:VOLT:" & mode & ":END:TIM " & value & ",(@" & OUT & ")"

            Case N6705_nst

                ts = "ARB:VOLT:" & mode & ":NST " & value & ",(@" & OUT & ")"
        End Select

        ilwrt(Power_Dev, ts, CInt(Len(ts)))

    End Function

    Function N6705_SEQ_parameter(ByVal OUT As String, ByVal mode As String, ByVal seq_step As Integer, ByVal parameter As String, ByVal value As String) As Integer

        Select Case parameter

            Case N6705_V0

                If mode = N6705_Sine Then
                    ts = "ARB:SEQ:STEP:VOLT:" & mode & ":AMPL " & value & "," & seq_step & ",(@" & OUT & ")"
                Else
                    ts = "ARB:SEQ:STEP:VOLT:" & mode & ":STAR " & value & "," & seq_step & ",(@" & OUT & ")"

                End If

            Case N6705_V1


                If mode = N6705_Pulse Or mode = N6705_Trapezoid Then
                    ts = "ARB:SEQ:STEP:VOLT:" & mode & ":TOP " & value & "," & seq_step & ",(@" & OUT & ")"

                ElseIf mode = N6705_Sine Then

                    ts = "ARB:SEQ:STEP:VOLT:" & mode & ":OFFS " & value & "," & seq_step & ",(@" & OUT & ")"
                Else

                    ts = "ARB:SEQ:STEP:VOLT:" & mode & ":END " & value & "," & seq_step & ",(@" & OUT & ")"

                End If

            Case N6705_t0


                If mode = N6705_Sine Then
                    'f = 1 / t0
                    ts = "ARB:SEQ:STEP:VOLT:" & mode & ":FREQ " & Val(1 / value) & "," & seq_step & ",(@" & OUT & ")"
                Else
                    ts = "ARB:SEQ:STEP:VOLT:" & mode & ":STAR:TIM " & value & "," & seq_step & ",(@" & OUT & ")"

                End If



            Case N6705_t1

                If mode = N6705_Pulse Then
                    ts = "ARB:SEQ:STEP:VOLT:" & mode & ":TOP:TIM " & value & "," & seq_step & ",(@" & OUT & ")"
                ElseIf mode = N6705_Staircase Then
                    ts = "ARB:SEQ:STEP:VOLT:" & mode & ":TIM " & value & "," & seq_step & ",(@" & OUT & ")"
                ElseIf mode = N6705_Step Then
                    ts = "ARB:SEQ:STEP:VOLT:" & mode & ":END:TIM " & value & "," & seq_step & ",(@" & OUT & ")"
                Else

                    ts = "ARB:SEQ:STEP:VOLT:" & mode & ":RTIM " & value & "," & seq_step & ",(@" & OUT & ")"
                End If


            Case N6705_t2

                If mode = N6705_Trapezoid Then

                    ts = "ARB:SEQ:STEP:VOLT:" & mode & ":TOP:TIM " & value & "," & seq_step & ",(@" & OUT & ")"

                Else
                    ts = "ARB:SEQ:STEP:VOLT:" & mode & ":END:TIM " & value & "," & seq_step & ",(@" & OUT & ")"
                End If


            Case N6705_t3
                ts = "ARB:SEQ:STEP:VOLT:" & mode & ":FTIM " & value & "," & seq_step & ",(@" & OUT & ")"
            Case N6705_t4

                ts = "ARB:SEQ:STEP:VOLT:" & mode & ":END:TIM " & value & "," & seq_step & ",(@" & OUT & ")"

            Case N6705_nst

                ts = "ARB:SEQ:STEP:VOLT:" & mode & ":NST " & value & "," & seq_step & ",(@" & OUT & ")"
        End Select

        ilwrt(Power_Dev, ts, CInt(Len(ts)))

    End Function


    Function N6705_ARB_editpoint(ByVal OUT As String) As Integer

        'ARB:VOLT:CONV (@1) 'Edit point
        'ARB:VOLT:UDEF:LEV:POIN? (@1) 'Step
        'ARB:VOLT:UDEF:LEVel? (@1) 'Voltage
        'ARB:VOLT:UDEF:DWELl? (@1)  'Time
        'ARB:VOLT:UDEF:BOSTep? (@1)'Tigger

        ts = "ARB:VOLT:CONV (@" & OUT & ")"
        ilwrt(Power_Dev, ts, CInt(Len(ts)))




    End Function

    Function N6705_UserDF_Volt(ByVal OUT As String) As String()
        Dim temp As String
        Dim level() As String


        'ARB:VOLT:CONV (@1) 'Edit point
        'ARB:VOLT:UDEF:LEV:POIN? (@1) 'Step
        'ARB:VOLT:UDEF:LEVel? (@1) 'Voltage
        'ARB:VOLT:UDEF:DWELl? (@1)  'Time
        'ARB:VOLT:UDEF:BOSTep? (@1)'Tigger

        ts = "ARB:VOLT:UDEF:LEVel? (@" & OUT & ")"
        ilwrt(Power_Dev, ts, CInt(Len(ts)))

        ilrd(Power_Dev, N6705_ValueStr, N6705_ARRAYSIZE)
        If (ibcnt > 0) And (ibsta <> EERR) Then
            temp = ibcntl - 1
            level = Split(Mid(N6705_ValueStr, 1, temp), ",")

        End If

        Return level

    End Function

    Function N6705_UserDF_Time(ByVal OUT As String) As String()
        Dim temp As String
        Dim time() As String


        'ARB:VOLT:CONV (@1) 'Edit point
        'ARB:VOLT:UDEF:LEV:POIN? (@1) 'Step
        'ARB:VOLT:UDEF:LEVel? (@1) 'Voltage
        'ARB:VOLT:UDEF:DWELl? (@1)  'Time
        'ARB:VOLT:UDEF:BOSTep? (@1)'Tigger

        ts = "ARB:VOLT:UDEF:DWELl? (@" & OUT & ")"
        ilwrt(Power_Dev, ts, CInt(Len(ts)))

        ilrd(Power_Dev, N6705_ValueStr, N6705_ARRAYSIZE)
        If (ibcnt > 0) And (ibsta <> EERR) Then
            temp = ibcntl - 1
            time = Split(Mid(N6705_ValueStr, 1, temp), ",")

        End If

        Return time

    End Function

    Function N6705_UserDF_Trig(ByVal OUT As String, ByVal ONOFF() As String) As Integer
        Dim temp As String
        Dim i As Integer

        temp = ""
        For i = 0 To ONOFF.Length - 1
            temp = temp & ONOFF(i) & ","
        Next


        'ARB:VOLT:CONV (@1) 'Edit point
        'ARB:VOLT:UDEF:LEV:POIN? (@1) 'Step
        'ARB:VOLT:UDEF:LEVel? (@1) 'Voltage
        'ARB:VOLT:UDEF:DWELl? (@1)  'Time
        'ARB:VOLT:UDEF:BOSTep? (@1)'Tigger

        ts = "ARB:VOLT:UDEF:BOSTep " & temp & "(@" & OUT & ")"
        ilwrt(Power_Dev, ts, CInt(Len(ts)))



    End Function

    Function N6705_ARB_set(ByVal OUT As String, ByVal mode As String, ByVal V0 As String, ByVal V1 As String, ByVal t0 As String, ByVal t1 As String, ByVal t2 As String, ByVal t3 As String, ByVal t4 As String, ByVal STA_steps As String) As Integer


        N6705_ARB_parameter(OUT, mode, N6705_V0, V0)

        N6705_ARB_parameter(OUT, mode, N6705_V1, V1)

        N6705_ARB_parameter(OUT, mode, N6705_t0, t0)


        If mode = N6705_Staircase Then
            't1
            N6705_ARB_parameter(OUT, mode, N6705_t1, t1)


            't2
            N6705_ARB_parameter(OUT, mode, N6705_t2, t2)

            'step
            N6705_ARB_parameter(OUT, mode, N6705_nst, STA_steps)

        ElseIf mode = N6705_Trapezoid Then

            't1
            N6705_ARB_parameter(OUT, mode, N6705_t1, t1)


            't2
            N6705_ARB_parameter(OUT, mode, N6705_t2, t2)

            't3
            N6705_ARB_parameter(OUT, mode, N6705_t3, t3)


            't4
            N6705_ARB_parameter(OUT, mode, N6705_t4, t4)

        Else
            't1
            N6705_ARB_parameter(OUT, mode, N6705_t1, t1)


            't2
            N6705_ARB_parameter(OUT, mode, N6705_t2, t2)


        End If






    End Function

    Function N6705_SEQ_set(ByVal OUT As String, ByVal mode As String, ByVal seq_step As Integer, ByVal V0 As String, ByVal V1 As String, ByVal t0 As String, ByVal t1 As String, ByVal t2 As String, ByVal t3 As String, ByVal t4 As String, ByVal STA_steps As String, ByVal counts As Integer) As Integer

        ts = "ARB:SEQ:STEP:FUNC:SHAP " & mode & ", " & seq_step & ", (@" & OUT & ")"
        ilwrt(Power_Dev, ts, CInt(Len(ts)))


        N6705_SEQ_parameter(OUT, mode, seq_step, N6705_V0, V0)

        N6705_SEQ_parameter(OUT, mode, seq_step, N6705_V1, V1)

        N6705_SEQ_parameter(OUT, mode, seq_step, N6705_t0, t0)


        If mode = N6705_Staircase Then
            't1
            N6705_SEQ_parameter(OUT, mode, seq_step, N6705_t1, t1)


            't2
            N6705_SEQ_parameter(OUT, mode, seq_step, N6705_t2, t2)

            'step
            N6705_SEQ_parameter(OUT, mode, seq_step, N6705_nst, STA_steps)

        ElseIf mode = N6705_Trapezoid Then

            't1
            N6705_SEQ_parameter(OUT, mode, seq_step, N6705_t1, t1)


            't2
            N6705_SEQ_parameter(OUT, mode, seq_step, N6705_t2, t2)

            't3
            N6705_SEQ_parameter(OUT, mode, seq_step, N6705_t3, t3)


            't4
            N6705_SEQ_parameter(OUT, mode, seq_step, N6705_t4, t4)

        Else
            't1
            N6705_SEQ_parameter(OUT, mode, seq_step, N6705_t1, t1)


            't2
            N6705_SEQ_parameter(OUT, mode, seq_step, N6705_t2, t2)


        End If


        ts = "ARB:SEQ:STEP:COUN " & counts & "," & seq_step & ",(@" & OUT & ")"

        ilwrt(Power_Dev, ts, CInt(Len(ts)))




    End Function




    'Function N6705_OUT_set(ByVal OUT As Integer, ByVal VOLT As Double, ByVal CURR As Double) As Integer


    '    ts = "VOLT " & VOLT & ",(@" & OUT & ")"
    '    ilwrt(Power_Dev, ts, CInt(Len(ts)))

    '    ts = "CURR " & CURR & ",(@" & OUT & ")"
    '    ilwrt(Power_Dev, ts, CInt(Len(ts)))



    'End Function

    'Function N6705_ONOFF(ByVal OUT As Integer, ByVal on_off As String) As Integer


    '    ts = "OUTP " & on_off & ",(@" & OUT & ")"
    '    ilwrt(Power_Dev, ts, CInt(Len(ts)))

    'End Function


    Function N6705_ARB_ONOFF(ByVal OUT As String, ByVal on_off As String) As Integer

        If on_off = "ON" Then

            ts = "INIT:TRAN " & "(@" & OUT & ")"



            ilwrt(Power_Dev, ts, CInt(Len(ts)))

            ts = "*TRG"

        Else

            ts = "ABORt:TRANsient " & "(@" & OUT & ")"
        End If



        ilwrt(Power_Dev, ts, CInt(Len(ts)))


    End Function

    Function N6705_ARB_status(ByVal OUT As String) As Boolean
        Dim temp As String
        Dim bit As Integer


        ts = "STAT:OPER:COND? " & "(@" & OUT & ")"

        ilwrt(Power_Dev, ts, CInt(Len(ts)))

        ilrd(Power_Dev, ValueStr, ARRAYSIZE)
        If (ibcnt > 0) And (ibsta <> EERR) Then
            temp = ibcntl - 1
            bit = bit_check(Val(Mid(ValueStr, 1, temp)), 6, 6)


        End If

        If bit = 1 Then
            Return True
        Else
            Return False
        End If


    End Function



    'Function N6705_ARB_init(ByVal OUT As Integer, ByVal last_Arb As String, ByVal Continuous As Boolean, ByVal Count As Integer) As Integer


    '    'OFF: Return to DC Value
    '    'ON: Last Arb Value


    '    ts = "ARB:FUNCtion TRAPezoid" & ",(@" & OUT & ")"
    '    ilwrt(Power_Dev, ts, CInt(Len(ts)))

    '    ts = "VOLT:MODE ARB" & ",(@" & OUT & ")"
    '    ilwrt(Power_Dev, ts, CInt(Len(ts)))

    '    ts = "ARB:TERM:LAST " & last_Arb & ",(@" & OUT & ")"
    '    ilwrt(Power_Dev, ts, CInt(Len(ts)))

    '    If Continuous = True Then
    '        'Repeat the Arb continuously.:
    '        ts = "ARB:COUN INF " & ",(@" & OUT & ")"

    '    Else
    '        'The number of times the Arb repeats.
    '        ts = "ARB:COUN " & Count & ",(@" & OUT & ")"

    '    End If

    '    ilwrt(Power_Dev, ts, CInt(Len(ts)))


    'End Function




    'Function N6705_ARB_set(ByVal OUT As Integer, ByVal V0 As Double, ByVal V1 As Double, ByVal t0 As Double, ByVal t1 As Double, ByVal t2 As Double, ByVal t3 As Double, ByVal t4 As Double) As Integer
    '    'Start(Setting(I0 Or V0))
    '    'Peak(Setting(I1 Or V1))
    '    'Delay(t0)
    '    'Rise(Time(t1))
    '    'Peak(Width(t2))
    '    'Fall(Time(t3))
    '    'End Time (T4)

    '    'The setting before and after the trapezoid:
    '    ts = "ARB:VOLT:TRAP:STAR " & V0 & ",(@" & OUT & ")"
    '    ilwrt(Power_Dev, ts, CInt(Len(ts)))

    '    'The peak setting:
    '    ts = "ARB:VOLT:TRAP:TOP " & V1 & ",(@" & OUT & ")"
    '    ilwrt(Power_Dev, ts, CInt(Len(ts)))

    '    'The delay after the trigger is received but before the trapezoid starts:
    '    ts = "ARB:VOLT:TRAP:STAR:TIM " & t0 & ",(@" & OUT & ")"
    '    ilwrt(Power_Dev, ts, CInt(Len(ts)))

    '    'The time that the trapezoid ramps up (RTIM) :
    '    ts = "ARB:VOLT:TRAP:RTIM " & t1 & ",(@" & OUT & ")"
    '    ilwrt(Power_Dev, ts, CInt(Len(ts)))

    '    'The width of the peak:
    '    ts = "ARB:VOLT:TRAP:TOP:TIM " & t2 & ",(@" & OUT & ")"
    '    ilwrt(Power_Dev, ts, CInt(Len(ts)))

    '    'The time that the trapezoid ramps down (FTIM):
    '    ts = "ARB:VOLT:TRAP:FTIM " & t3 & ",(@" & OUT & ")"
    '    ilwrt(Power_Dev, ts, CInt(Len(ts)))

    '    'The time the output remains at the end setting after the trapezoid:
    '    ts = "ARB:VOLT:TRAP:END:TIM " & t4 & ",(@" & OUT & ")"
    '    ilwrt(Power_Dev, ts, CInt(Len(ts)))



    'End Function

    Function power_OCP_init(ByVal device As String, ByVal power_out As String, ByVal ocp As Double) As Integer

        If (Mid(device, 1, 6) = "62006P") Or (Mid(device, 1, 6) = "62012P") Then
            ts = "SOURce:CURRent " & Format(ocp, "#0.000")

        ElseIf (device = "N6705") Or (device = "E36312A") Then

            ts = "CURR " & ocp & ",(@" & power_out & ")"

        ElseIf device = "MODEL 2400" Then




        Else
            If power_out <> "" Then
                ts = "INST:SEL " & power_out
                ilwrt(Power_Dev, ts, CInt(Len(ts)))
            End If

            ts = "CURR " & Format(ocp, "#0.0")
        End If



        ilwrt(Power_Dev, ts, CInt(Len(ts)))
        Delay(100)

    End Function

    Function power_OCP_value() As Double
        Dim test As Double = 0
        Dim temp As String
        ts = "CURR?"
        ilwrt(Power_Dev, ts, CInt(Len(ts)))

        ilrd(Power_Dev, ValueStr, ARRAYSIZE)
        If (ibcnt > 0) And (ibsta <> EERR) Then
            temp = ibcntl - 1
            test = Val(Mid(ValueStr, 1, temp))
        End If
        power_OCP_value = Format(Val(test), "#0.0000")
        Return power_OCP_value
    End Function



    Function power_on_off(ByVal device As String, ByVal power_out As String, ByVal on_off As String, Optional ByVal dut2_en As Boolean = False) As Integer

        If dut2_en Then
            Power_Dev = vin_Dev2
        Else
            Power_Dev = vin_Dev
        End If
        If (Mid(device, 1, 6) = "62006P") Or (Mid(device, 1, 6) = "62012P") Then
            ts = "CONFigure:OUTPut " & on_off
        Else
            Select Case device
                Case "N6705"
                    ts = "OUTP " & on_off & ",(@" & power_out & ")"
                Case "E36312A"
                    '2022/07/22
                    ts = "OUTP " & on_off & ",(@" & power_out & ")"
                Case "6210-40"
                    ts = "OUT " & on_off
                Case Else
                    ts = "OUTP " & on_off
            End Select
        End If
        ilwrt(Power_Dev, ts, CInt(Len(ts)))
        'If on_off = "ON" Then
        '    Delay(300)
        'End If
        ts = "*OPC?"
        ilwrt(Power_Dev, ts, CInt(Len(ts)))
        ilrd(Power_Dev, ValueStr, ARRAYSIZE)
    End Function



    Function power_volt(ByVal device As String, ByVal power_out As String, ByVal volt As Double) As Integer


        If (Mid(device, 1, 6) = "62006P") Or (Mid(device, 1, 6) = "62012P") Then
            ts = "SOURce:VOLTage " & Format(volt, "#0.000")
        Else


            Select Case device


                Case "N6705"


                    ts = "VOLT " & volt & ",(@" & power_out & ")"


                Case "E36312A"
                    '2022/07/22 Add
                    ts = "VOLT " & volt & ",(@" & power_out & ")"

                Case "6210-40"

                    ts = "VSET " & Format(volt, "#0.000")


                Case "MODEL 2400"
                    ts = "SOUR:VOLT " & volt.ToString("F4")
                Case "MODEL 2410"
                    ts = "SOUR:VOLT " & volt.ToString("F4")

                Case Else
                    If power_out <> "" Then
                        ts = "INST:SEL " & power_out
                        ilwrt(Power_Dev, ts, CInt(Len(ts)))
                    End If


                    ts = "VOLT " & Format(volt, "#0.000")


            End Select
        End If
        ilwrt(Power_Dev, ts, CInt(Len(ts)))
        Delay(200)
        If device = "E3632A" Then
            If (volt <= 15) Then
                If E3632_Range_H = True Then
                    power_OCP_init(device, power_out, E3632_OCP)
                End If
                E3632_Range_H = False
            Else
                E3632_Range_H = True
            End If
        End If
    End Function

    Function power_read(ByVal device As String, ByVal power_out As String, ByVal volt_curr As String, Optional ByVal dut2_en As Boolean = False) As Double
        Dim test As Double = 0
        Dim temp As String

        If dut2_en Then
            Power_Dev = vin_Dev2
        Else
            Power_Dev = vin_Dev
        End If

        If device = "6210-40" Then

            If volt_curr = "VOLT" Then
                ts = "VOUT?"
            Else
                ts = "IOUT?"
            End If

        ElseIf (Mid(device, 1, 6) = "E36312") Then

            ts = "MEAS:" & volt_curr & "? CH" & power_out



        Else


            If power_out <> "" Then
                ts = "INST:SEL " & power_out
                ilwrt(Power_Dev, ts, CInt(Len(ts)))
            End If

            If volt_curr = "VOLT" Then
                ts = "MEAS:VOLT?"
            Else
                ts = "MEAS:CURR?"
            End If

        End If


        ilwrt(Power_Dev, ts, CInt(Len(ts)))
        ilrd(Power_Dev, ValueStr, ARRAYSIZE)
        If (ibcnt > 0) And (ibsta <> EERR) Then
            temp = ibcntl - 1
            test = Val(Mid(ValueStr, 1, temp))
        End If
        power_read = Format(Val(test), "#0.0000")

        Return power_read

    End Function









End Module
