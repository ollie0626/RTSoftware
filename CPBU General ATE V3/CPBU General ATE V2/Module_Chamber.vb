Module Module_Chamber


    Public Temp_Dev As Integer
    Public Temp_addr As Integer
    Public ts As String
    Public Temp_Time As Double = 0
    Public Chamber_ok As Boolean = True
    Public Temp_name As String


    Function chamber_init() As Integer
        'TIMEOUT = T10s
        If Chamber_ok = True Then
            Temp_Dev = ildev(BDINDEX, Temp_addr, NO_SECONDARY_ADDR, TIMEOUT, EOTMODE, EOSMODE)
        End If
    End Function

    Function set_temp(ByVal temp As Double, ByVal time As Double) As Integer

        If Chamber_ok = True Then
            '正 = 空白 Convert.ToChar(&HB) , 負 = -
            ts = "T" & temp & "," & time & vbNewLine  'sprintf(ts,"T %d.0,0 \n\r",chamber_settemp); //設定溫箱
            ilwrt(Temp_Dev, ts, CInt(Len(ts)))
        Else
            Set_TEC_temp(temp)
            TEC_ONOFF("ON")
        End If

    End Function


    Function read_set() As Double
        Dim test As Double


        If Chamber_ok = True Then
            'Read Set Temp

            ts = "ST" & vbNewLine  'Convert.ToChar(13) & Convert.ToChar(10)

            ilwrt(Temp_Dev, ts, CInt(Len(ts)))

            ' Delay(100)
            ilrd(Temp_Dev, ValueStr, ARRAYSIZE)

            If ibcnt > 0 Then
                test = Val(Mid(ValueStr, 1, (ibcntl - 1)))
            End If

        End If


        Return test
    End Function

    Function read_temp() As Double
        Dim test As Double
        Dim temp As String
        Dim result_data() As Double


        If Chamber_ok = True Then
            ts = "AT" & vbNewLine  'Convert.ToChar(13) & Convert.ToChar(10)
            ilwrt(Temp_Dev, ts, CInt(Len(ts)))

            ' Delay(100)
            ilrd(Temp_Dev, ValueStr, ARRAYSIZE)

            If ibcnt > 0 Then

                temp = Mid(ValueStr, 1, 1)
                If temp = "-" Or temp = "=" Then
                    test = -Val(Mid(ValueStr, 2, (ibcntl - 1)))
                Else
                    test = Val(Mid(ValueStr, 1, (ibcntl - 1)))
                End If


            End If

        Else

            result_data = Read_TEC_temp()
            If result_data(0) = 0 Then
                test = result_data(1)
            Else
                test = 0
            End If
        End If


        Return test
    End Function

    Function temp_off() As Integer

        If Chamber_ok = True Then

            'Turns temperature off.
            ts = "KT" & vbNewLine 'Convert.ToChar(13)
            ilwrt(Temp_Dev, ts, CInt(Len(ts)))

        Else
            Set_TEC_temp(25)
            Delay(1000)
            TEC_ONOFF("OFF")
        End If

    End Function

    Function Chamber_off() As Integer
        'Turns the chamber off and releases the test wheel.

        If Chamber_ok = True Then

            ts = "O" & vbNewLine 'Convert.ToChar(13)
            ilwrt(Temp_Dev, ts, CInt(Len(ts)))


        Else
            TEC_ONOFF("OFF")
        End If

    End Function


End Module
