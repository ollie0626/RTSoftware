Module Module_DCLoad

    'Public Load_name As String
    Public Load_Dev As Integer
    Public Load_addr As Integer


    Public Load_Mode_CC As String = "CC"
    Public Load_Mode_CRL As String = "CRL"
    Public Load_Mode_CRH As String = "CRH"
    Public Load_Mode_CV As String = "CV"
    Public Load_range_L As String = "L"
    Public Load_range_M As String = "M"
    Public Load_range_H As String = "H"
    Public Load_ch As Integer = 1

    Public Load_mode As String = Load_Mode_CC
    Public Load_device As String
    Public Load_range As String = Load_range_L

    Public DCLoad_CCH As Double = 2
    Public DCLoad_CCL As Double = 0.2 'As Double '= 2
    Public LOAD_6312_Model As String = "63103A"

    Public LOAD_63600_CCH() As Double '= 2
    Public LOAD_63600_CCL() As Double '= 0.2
    Public LOAD_63600_Watt_L() As Integer
    Public LOAD_63600_Watt_M() As Integer
    Public LOAD_63600_Watt_H() As Integer
    Public LOAD_63600_Model() As String

    Public DCLOAD_63600 As Boolean = False


    Dim ts As String = ""

    Function load_model_check(ByVal ch As Integer, ByVal num As Integer) As Integer
        Dim temp() As String
        Dim model() As String
        Dim i As Integer
        Dim CCH, CCL As Double
        Dim Watt_L, Watt_M, Watt_H As Integer

        Load_Dev = ildev(BDINDEX, Load_addr, NO_SECONDARY_ADDR, TIMEOUT, EOTMODE, EOSMODE)

        ts = "CHAN " & ch
        ilwrt(Load_Dev, ts, CInt(Len(ts)))
        ts = "CHAN:ID?"
        ilwrt(Load_Dev, ts, CInt(Len(ts)))
        ilrd(Load_Dev, ValueStr, ARRAYSIZE)


        If (ibcnt > 0) And (ibsta <> EERR) Then
            temp = Split(ValueStr, ",")
            model = Split(temp(1), "-")


            Select Case model(0)

                Case "63610"

                    CCL = 0.2
                    CCH = 2
                    Watt_L = 16
                    Watt_M = 30
                    Watt_H = 100

                Case "63630"
                    If model(1) = "600" Then

                        CCL = 0.15
                        CCH = 1.5
                        Watt_L = 90
                        Watt_M = 300
                        Watt_H = 300
                    Else

                        CCL = 0.6
                        CCH = 6
                        Watt_L = 30
                        Watt_M = 60
                        Watt_H = 300
                    End If

                Case "63640"
                    If model(1) = "150" Then

                        CCL = 1
                        CCH = 6
                        Watt_L = 90
                        Watt_M = 400
                        Watt_H = 400
                    Else

                        CCL = 0.8
                        CCH = 8
                        Watt_L = 60
                        Watt_M = 60
                        Watt_H = 400
                    End If

            End Select

            LOAD_63600_Model(num) = temp(1)

            LOAD_63600_CCL(num) = CCL
            LOAD_63600_CCH(num) = CCH * 0.8

            LOAD_63600_Watt_L(num) = Watt_L
            LOAD_63600_Watt_M(num) = Watt_M * 0.8
            LOAD_63600_Watt_H(num) = Watt_H

        End If

        ibonl(Load_Dev, 0)

    End Function

    Function load_init(ByVal range As String) As Double

        ts = "CHAN " & Load_ch

        ilwrt(Load_Dev, ts, CInt(Len(ts)))

        Select Case Load_mode
            Case "CC"
                If Mid(Load_device, 1, 3) = "630" Then
                    ts = "MODE:CURR:A "
                    ilwrt(Load_Dev, ts, CInt(Len(ts)))
                    Select Case range
                        Case "L"
                            ts = "CURR:RANGE MIN"

                        Case "H"
                            ts = "CURR:RANGE MAX"

                    End Select
                Else
                    Select Case range
                        Case "L"
                            ts = "MODE CCL"

                        Case "H"
                            ts = "MODE CCH"


                        Case "M"

                            ts = "MODE CCM"

                    End Select
                End If
            Case "CRL"

                ts = "MODE CRL"


            Case "CRH"

                ts = "MODE CRH"

            Case "CV"
                If Mid(Load_device, 1, 3) = "631" Then

                    ts = "MODE CV"

                Else
                    ts = "MODE CVL"


                End If

        End Select


        ilwrt(Load_Dev, ts, CInt(Len(ts)))

        Delay(1000)

    End Function


    Function load_onoff(ByVal onoff As String) As Integer
        ts = "CHAN " & Load_ch

        ilwrt(Load_Dev, ts, CInt(Len(ts)))



        If onoff = "ON" Then
            ts = "LOAD ON"
            DCLoad_ON = True
        Else
            ts = "LOAD OFF"

            DCLoad_ON = False
        End If

        ilwrt(Load_Dev, ts, CInt(Len(ts)))

    End Function

    Function load_set(ByVal iout As Double) As Double

        ts = "CHAN " & Load_ch

        ilwrt(Load_Dev, ts, CInt(Len(ts)))

        Select Case Load_mode

            Case "CC"

                If Mid(Load_device, 1, 3) = "630" Then

                    ts = "CURR:A " & Format(iout, "#0.000")
                Else

                    ts = "CURR:STAT:L1 " & Format(iout, "#0.0000")
                End If


            Case "CRL"

                If Mid(Load_device, 1, 3) = "631" Then

                    ts = "RES:L1 " & Format(iout, "#0.000")
                Else

                    ts = "RES:STAT:L1 " & Format(iout, "#0.0000")
                End If



            Case "CRH"
                If Mid(Load_device, 1, 3) = "631" Then

                    ts = "RES:L1 " & Format(iout, "#0.000")
                Else

                    ts = "RES:STAT:L1 " & Format(iout, "#0.0000")
                End If



            Case "CV"

                If Mid(Load_device, 1, 3) = "631" Then

                    ts = "VOLT:L1 " & Format(iout, "#0.000")
                Else

                    ts = "VOLT:STAT:L1 " & Format(iout, "#0.0000")
                End If




        End Select




        ilwrt(Load_Dev, ts, CInt(Len(ts)))

        If Mid(Load_device, 1, 3) = "630" Then
            Delay(300)
        End If




    End Function

    Function load_read(ByVal volt_curr As String) As Double
        Dim test As Double = 0
        Dim temp As String

        ts = "CHAN " & Load_ch
        ilwrt(Load_Dev, ts, CInt(Len(ts)))
        Delay(50)

        If volt_curr = "VOLT" Then
            ts = "MEAS:VOLT?"
        Else
            ts = "MEAS:CURR?"
        End If
        ilwrt(Load_Dev, ts, CInt(Len(ts)))
        Delay(50)
        ilrd(Load_Dev, ValueStr, ARRAYSIZE)
        Delay(200)

        If (ibcnt > 0) And (ibsta <> EERR) Then
            temp = ibcntl - 1

            test = Val(Mid(ValueStr, 1, temp))
        End If


        Return Format(Val(test), "#0.000000")


    End Function

    Function Dynamic_init(ByVal range As String) As Integer


        ts = "CHAN " & Load_ch
        ilwrt(Load_Dev, ts, CInt(Len(ts)))

        If Mid(Load_device, 1, 3) = "630" Then

            ts = "MODE:CURR:DYN"

            ilwrt(Load_Dev, ts, CInt(Len(ts)))
            Select Case range
                Case "L"
                    ts = "CURR:RANGE MIN"
                    ilwrt(Load_Dev, ts, CInt(Len(ts)))
                Case "H"
                    ts = "CURR:RANGE MAX"
                    ilwrt(Load_Dev, ts, CInt(Len(ts)))
            End Select
        Else

            Select Case range
                Case "L"
                    ts = "MODE CCDL"
                    ilwrt(Load_Dev, ts, CInt(Len(ts)))

                Case "H"
                    ts = "MODE CCDH"
                    ilwrt(Load_Dev, ts, CInt(Len(ts)))


                Case "M"

                    ts = "MODE CCDM"
                    ilwrt(Load_Dev, ts, CInt(Len(ts)))
                    'Delay(200)

            End Select




        End If





    End Function



    Function Dynamic_set(ByVal Imax As Double, ByVal Imin As Double, ByVal T1 As Double, ByVal T1_unit As String, ByVal T2 As Double, ByVal T2_unit As String, ByVal RISE As Double, ByVal FALL As Double) As Integer



        ts = "CHAN " & Load_ch
        ilwrt(Load_Dev, ts, CInt(Len(ts)))

        If Mid(Load_device, 1, 3) = "630" Then

            ts = "CURR:DYN:H " & Imax

            ilwrt(Load_Dev, ts, CInt(Len(ts)))


            ts = "CURR:DYN:L " & Imin

            ilwrt(Load_Dev, ts, CInt(Len(ts)))


            ts = "CURR:TIME:T1 " & T1 & T1_unit 'S"

            ilwrt(Load_Dev, ts, CInt(Len(ts)))


            ts = "CURR:TIME:T2 " & T2 & T2_unit

            ilwrt(Load_Dev, ts, CInt(Len(ts)))


            ts = "CURRent:SLEW:DYNamic:RISE " & RISE ' "A/uS"

            ilwrt(Load_Dev, ts, CInt(Len(ts)))


            ts = "CURRent:SLEW:DYNamic:FALL " & FALL ' "A/uS"

            ilwrt(Load_Dev, ts, CInt(Len(ts)))


        Else
            ts = "CURR:DYN:L1 " & Imax

            ilwrt(Load_Dev, ts, CInt(Len(ts)))


            ts = "CURR:DYN:L2 " & Imin

            ilwrt(Load_Dev, ts, CInt(Len(ts)))


            ts = "CURR:DYN:T1 " & T1 & T1_unit 'mS, S

            ilwrt(Load_Dev, ts, CInt(Len(ts)))


            ts = "CURR:DYN:T2 " & T2 & T2_unit

            ilwrt(Load_Dev, ts, CInt(Len(ts)))


            ts = "CURRent:DYNamic:RISE " & RISE ' "A/uS"

            ilwrt(Load_Dev, ts, CInt(Len(ts)))


            ts = "CURRent:DYNamic:FALL " & FALL ' "A/uS"

            ilwrt(Load_Dev, ts, CInt(Len(ts)))


        End If



    End Function

    Function Dynamic_slewrate_Max() As Double

        Dim read_val As Double
        Dim temp As String

        ts = "CHAN " & Load_ch
        ilwrt(Load_Dev, ts, CInt(Len(ts)))


        If Mid(Load_device, 1, 3) = "630" Then



            ts = "CURRent:SLEW:DYNamic:RISE MAX"

            ilwrt(Load_Dev, ts, CInt(Len(ts)))

            ts = "CURRent:SLEW:DYNamic:RISE?"

            ilwrt(Load_Dev, ts, CInt(Len(ts)))

        Else

            ts = "CURRent:DYNamic:RISE MAX"

            ilwrt(Load_Dev, ts, CInt(Len(ts)))

            ts = "CURRent:DYNamic:RISE?"

            ilwrt(Load_Dev, ts, CInt(Len(ts)))

        End If


        ilrd(Load_Dev, ValueStr, ARRAYSIZE)
        If ibcntl > 0 Then
            temp = ibcntl - 1
            read_val = Val(Mid(ValueStr, 1, temp))
        End If

        Return read_val


    End Function

    Function Dynamic_Read(ByVal Rise As Boolean) As Double


        Dim read_val As Double
        Dim temp As String

        ts = "CHAN " & Load_ch
        ilwrt(Load_Dev, ts, CInt(Len(ts)))

        If Rise = True Then
            ts = "CURRent:DYNamic:RISE?"
        Else
            ts = "CURRent:DYNamic:FALL?"

        End If



        ilwrt(Load_Dev, ts, CInt(Len(ts)))


        ilrd(Load_Dev, ValueStr, ARRAYSIZE)
        If ibcntl > 0 Then
            temp = ibcntl - 1
            read_val = Val(Mid(ValueStr, 1, temp))
        End If

        Return read_val



    End Function


End Module
