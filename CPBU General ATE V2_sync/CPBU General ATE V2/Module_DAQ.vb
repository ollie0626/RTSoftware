Module Module_DAQ
    Public DAQ_name As String
    Public DAQ_Dev As Integer
    Public DAQ_addr As Integer
    Public DAQ_resolution As String = "DEF" 'DEF=5 1/2; MIN=6 1/2; MAX=4 1/2
    Public DC_AUTO As String = "AUTO"
    Public DC_100 As String = "100"
    Public DC_10 As String = "10"
    Public DC_1 As String = "1"
    Public DAQ_unit As String = DC_AUTO
    Dim ts As String


    Function DAQ_config(ByVal channel As String) As Integer
        Dim ts As String

        ts = "CONFigure:VOLTage:DC " & DAQ_unit & ", " & DAQ_resolution & ",(@" & channel & ")"
        ilwrt(DAQ_Dev, ts, CInt(Len(ts)))

    End Function

    Function DAQ_average(ByVal channel As String, ByVal average As Integer) As Double
        Dim i As Integer
        Dim temp, total As Double



        For i = 1 To average
            System.Windows.Forms.Application.DoEvents()

            If run = False Then
                Exit For
            End If



            temp = DAQ_read(channel)



            While temp > (10 ^ 10)
                System.Windows.Forms.Application.DoEvents()

                If run = False Then
                    Exit While
                End If

                temp = DAQ_read(channel)
                Delay(10)
            End While

            If i = 1 Then
                total = temp
            Else
                total = total + temp
            End If

        Next
        total = total / average
        Return total
    End Function


    Function DAQ_read(ByVal channel As String) As Double

        Dim temp As Integer
        Dim test As Double = 0
        Dim i As Integer

        ts = "MEAS:VOLT:DC? " & DAQ_unit & ", " & DAQ_resolution & ",(@" & channel & ")"

        ilwrt(DAQ_Dev, ts, CInt(Len(ts)))
        ilrd(DAQ_Dev, ValueStr, ARRAYSIZE)

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
                ilwrt(DAQ_Dev, ts, CInt(Len(ts)))
                Delay(100)
            End If

            ilrd(DAQ_Dev, ValueStr, ARRAYSIZE)

        Next

        Return test


    End Function


    Function DAQ_read_Temp(ByVal channel As String, ByVal thermal_type As String) As Double

        ts = "MEAS:TEMP? TC," & thermal_type & ",(@" & channel & ")"

        ilwrt(DAQ_Dev, ts, CInt(Len(ts)))
        ilrd(DAQ_Dev, ValueStr, ARRAYSIZE)



        If ibcntl > 0 Then
            Return Val(Mid(ValueStr, 1, ibcntl - 1))
        Else
            Return 0
        End If


    End Function


    Function DAQ_inputR(ByVal channel As String, ByVal ONOFF As Integer) As Integer
        Dim ts As String = ""

        'AUTO OFF: 10 MΩ (100 mV, 1 V, 10 V, 100 V, 300 V ranges);
        'AUTO ON: > 10 GΩ (100 mV, 1 V, 10 V ranges); 10 MΩ (100 V, 300 V ranges).

        If ONOFF = 0 Then
            ts = "INPUT:IMPEDANCE:AUTO OFF,(@" & channel & ")"

        ElseIf ONOFF = 1 Then
            ts = "INPUT:IMPEDANCE:AUTO ON,(@" & channel & ")"

        End If

        ilwrt(DAQ_Dev, ts, CInt(Len(ts)))


    End Function


    Function DAQ_monitor_set(ByVal channel As String, ByVal on_off As Boolean) As Double

        ts = "ROUTe:MONitor (@" & channel & ")"
        ilwrt(DAQ_Dev, ts, CInt(Len(ts)))


        If on_off = True Then
            ts = "ROUTe:MONitor:STATe ON"
        Else
            ts = "ROUTe:MONitor:STATe OFF"
        End If

        ilwrt(DAQ_Dev, ts, CInt(Len(ts)))
    End Function


    Function DAQ_monitor() As Double


        ts = "ROUTe:MONitor:DATA?"

        ilwrt(DAQ_Dev, ts, CInt(Len(ts)))

        ilrd(DAQ_Dev, ValueStr, ARRAYSIZE)

        If ibcntl > 0 Then
            Return Val(Mid(ValueStr, 1, ibcntl - 1))
        Else
            Return 0
        End If


    End Function


End Module
