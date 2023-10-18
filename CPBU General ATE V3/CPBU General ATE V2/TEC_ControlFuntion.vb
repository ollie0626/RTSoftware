Module TEC_ControlFuntion
    Private cmdOut As Integer
    Private dataOutCount As Integer
    Private dataOut(1023) As Byte
    Public isoboardNames As String() = {"RTIsoSparrowboard"}
    Public TEC_ok As Boolean = False
    Public TEC_MAX As Double = 110
    Public TEC_MIN As Double = -10

    Private Enum TECCmd
        SetTemp = &H1002
        SetOnOff = &H1003
        ReadTemp = &H1008
        ReadTecTempMaxMin = &H1009
    End Enum

    Function Set_TEC_temp(ByVal temp As Double) As Integer
        Dim final_errcode As Integer = 0
        Dim first_cal As Double = 0
        Dim final_cal As Integer = 0
        Dim final_temp As Byte()
        If ((temp * 100 / 25) - (temp * 100 \ 25)) >= 0.5 Then
            first_cal = 0.25 * ((temp * 100 \ 25) + 1)
        Else
            first_cal = 0.25 * (temp * 100 \ 25)
        End If
        final_cal = first_cal * 4
        final_temp = BitConverter.GetBytes(final_cal)


        Dim errCode As Integer = RTBB_Iso_Transact(pIsoDevice, {TECCmd.SetTemp}, {4}, final_temp, {cmdOut}, {dataOutCount}, dataOut)

        If errCode <> 0 Then
            final_errcode = errCode
        ElseIf cmdOut <> 0 Then
            final_errcode = cmdOut
        End If

        Console.WriteLine($"{NameOf(Set_TEC_temp)} final_errcode: {final_errcode}")
        Return final_errcode

    End Function

    Function Read_TEC_temp() As Double()
        Dim final_errcode As Integer = 0
        Dim errCode As Integer = RTBB_Iso_Transact(pIsoDevice, {TECCmd.ReadTemp}, {0}, {}, {cmdOut}, {dataOutCount}, dataOut)
        Dim tempIntegerBy0_25C = BitConverter.ToInt32(dataOut, 0)
        Dim result_data(1) As Double
        Console.WriteLine($"cmdOut: {cmdOut}")
        Console.WriteLine($"dataOutCount: {dataOutCount}")
        Console.WriteLine($"dataOut to Int32: 0x{Hex(tempIntegerBy0_25C)}")
        Console.WriteLine($"errCode: {errCode}")

        result_data(1) = Convert.ToDouble(tempIntegerBy0_25C) / 4

        If errCode <> 0 Then
            final_errcode = errCode
        ElseIf cmdOut <> 0 Then
            final_errcode = cmdOut
        End If

        result_data(0) = final_errcode
        'Console.WriteLine($"{NameOf(Read_TEC_temp)} Now Temp: {temp}")
        'Console.WriteLine($"{NameOf(Read_TEC_temp)} final_errcode: {final_errcode}")

        Return result_data
    End Function

    Function TEC_ONOFF(ByVal status As String) As Integer
        Dim final_errcode As Integer = 0
        Dim ON_OFF As Byte = 0
        If status = "ON" Then
            ON_OFF = 1
        End If

        Dim errCode As Integer = RTBB_Iso_Transact(pIsoDevice, {TECCmd.SetOnOff}, {1}, {ON_OFF}, {cmdOut}, {dataOutCount}, dataOut)

        If errCode <> 0 Then
            final_errcode = errCode
        ElseIf cmdOut <> 0 Then
            final_errcode = cmdOut
        End If

        Console.WriteLine($"{NameOf(TEC_ONOFF)} final_errcode: {final_errcode}")
        Return final_errcode
    End Function


    Function Read_TEC_MAX_MIN() As Double()
        Dim result(2) As Double
        Dim final_errcode As Integer = 0

        Dim errCode As Integer = RTBB_Iso_Transact(pIsoDevice, {TECCmd.ReadTecTempMaxMin}, {0}, {}, {cmdOut}, {dataOutCount}, dataOut)
        Console.WriteLine($"cmdOut: {cmdOut}")
        Console.WriteLine($"dataOutCount: {dataOutCount}")
        Dim TEC_MAX_IntegerBy0_25C = BitConverter.ToInt32(dataOut, 0)
        Console.WriteLine($"dataOut to Int32: 0x{Hex(TEC_MAX_IntegerBy0_25C)}")
        Dim TEC_Min_IntegerBy0_25C = BitConverter.ToInt32(dataOut, 4)
        Console.WriteLine($"dataOut to Int32: 0x{Hex(TEC_Min_IntegerBy0_25C)}")
        Console.WriteLine($"errCode: {errCode}")


        If errCode <> 0 Then
            final_errcode = errCode
        ElseIf cmdOut <> 0 Then
            final_errcode = cmdOut
        End If

        result(0) = final_errcode
        TEC_MAX = Convert.ToDouble(TEC_MAX_IntegerBy0_25C) / 4 'TEC_MAX
        TEC_MIN = Convert.ToDouble(TEC_Min_IntegerBy0_25C) / 4 'TEC_MIN

        result(1) = TEC_MAX
        result(2) = TEC_MIN


        Console.WriteLine($"{NameOf(Read_TEC_MAX_MIN)} final_errcode: {final_errcode}")
        Return result
    End Function




End Module
