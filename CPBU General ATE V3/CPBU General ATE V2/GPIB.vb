Module GPIB

    'GPIB
    Public BDINDEX As Short = 0 ' Board Index
    'Dim PRIMARY_ADDR_OF_DMM As Short = 6 ' Primary address of device
    Public NO_SECONDARY_ADDR As Short = 0 ' Secondary address of device
    Public TIMEOUT As Short = T1s ' Timeout value = 10 seconds
    Public EOTMODE As Short = 1 ' Enable the END message
    Public EOSMODE As Short = 0 ' Disable the EOS mode
    Public ErrMsg As String '  New VB6.FixedLengthString(100)
    Public ARRAYSIZE As Short = 255 ' Size of read buffer
    Public ValueStr As String = Space(ARRAYSIZE)

    Public pad(&H1F) As Short
    Public instrument(&H1F) As Integer
    Public description(&H1F) As String
    Public dev_num As Integer


    Function check_gpib(ByVal status As Object, ByVal device As String) As Boolean
        Dim gpib_error As Boolean
        If (ibsta And EERR) Then
            status.Text = "GPIB connection to the " & device & " is not detected!"
            gpib_error = True
        Else
            gpib_error = False
        End If
        Return gpib_error
    End Function

    Function gpib_rst(ByVal device As String) As Double
        Dim ts As String
        ts = "*RST"
        ilwrt(device, ts, CInt(Len(ts)))

    End Function

    Function check_device(ByVal gpib As String) As String()
        Dim dev_addr(&H1F) As Short

        Dim i As Byte
        Dim dev As Short
        Dim value As Short
        Dim ts As String
        Dim txt(32) As String


        Dim temp() As String

        For i = 0 To &H1D
            dev_addr(i) = i

        Next
        dev_addr(&H1E) = NOADDR
        ' ibonl(0, 0)
        dev = ibfind32(gpib)
        ibconfig(dev, IbcSC, 1)
        ibsic(dev)
        ibconfig(dev, IbcSRE, 1)
        ibconfig(dev, IbcTIMING, 1)
        ibask(dev, IbaPAD, value)
        ibask(dev, IbaSAD, value)

        FindLstn(BDINDEX, dev_addr, pad, &H1E)
        If (ibsta And EERR) Then
            dev_num = 0
            MsgBox("GPIB connection to is not detected!", MsgBoxStyle.Critical, "Error Message")
            If iberr = 28 Then
                Application.Restart()
            End If


        Else


            dev_num = ibcntl - 1

            For i = 1 To dev_num
                instrument(i - 1) = ildev(BDINDEX, pad(i), NO_SECONDARY_ADDR, TIMEOUT, EOTMODE, EOSMODE)
                ibclr(instrument(i - 1))


                'For chamber 2015/05/06 Add
                If pad(i) = 3 Then
                    ts = "AT" & vbNewLine  'Convert.ToChar(13) & Convert.ToChar(10)
                    Temp_addr = 3
                    Temp_Dev = ildev(BDINDEX, Temp_addr, NO_SECONDARY_ADDR, TIMEOUT, EOTMODE, EOSMODE)

                    ilwrt(Temp_Dev, ts, CInt(Len(ts)))

                    ' Delay(100)
                    ilrd(Temp_Dev, ValueStr, ARRAYSIZE)

                    If ibcnt > 0 Then
                        description(i - 1) = "4350B"
                    Else
                        Temp_addr = 0
                        Temp_Dev = 0

                    End If


                Else

                    ts = "*IDN?"
                    ilwrt(instrument(i - 1), ts, CInt(Len(ts)))
                    ilrd(instrument(i - 1), ValueStr, ARRAYSIZE)

                    txt(i - 1) = ""
                    description(i - 1) = ""


                    If ibcntl = 1 Then

                        ts = "ID?"
                        ilwrt(instrument(i - 1), ts, CInt(Len(ts)))
                        ilrd(instrument(i - 1), ValueStr, ARRAYSIZE)
                        If ibcntl > 0 Then
                            txt(i - 1) = Mid(ValueStr, 1, (ibcntl - 1))
                            temp = Split(ValueStr, " ")
                            If temp.Length = 1 Then
                                MsgBox("Have same GPIB address!!", MsgBoxStyle.Critical, "Error Message")
                            Else
                                description(i - 1) = temp(1)
                            End If

                        End If

                    ElseIf ibcntl > 1 Then

                        txt(i - 1) = Mid(ValueStr, 1, (ibcntl - 1))
                        temp = Split(ValueStr, ",")
                        If temp.Length = 1 Then
                            MsgBox("Have same GPIB address!!", MsgBoxStyle.Critical, "Error Message")
                        Else
                            description(i - 1) = temp(1)
                        End If

                    End If
                End If
                ibonl(instrument(i - 1), 0)

            Next
        End If

        ibonl(dev, 0)

        Return txt




    End Function


End Module
