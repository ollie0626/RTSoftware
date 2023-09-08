Module Function_Generator

    Public FG_num As Integer
    Public FG_name() As String
    Public FG_Dev As Integer
    Public FG_Addr() As String
    Dim ts As String = ""


    Function FG_State(ByVal ONOFF As String) As Integer

        'AFG3021C OUTPut[1|2][:STATe] {ON|OFF|<NR1>}

        ' ts = "OUTPut1:STATe " & ONOFF
        ts = "OUTPut " & ONOFF

        ilwrt(FG_Dev, ts, CInt(Len(ts)))


    End Function


    Function FG_Voltage(ByVal HIGH As Double, ByVal LOW As Double) As Integer


        ts = "VOLTage:HIGH " & HIGH
        ilwrt(FG_Dev, ts, CInt(Len(ts)))

        ts = "VOLTage:LOW " & LOW
        ilwrt(FG_Dev, ts, CInt(Len(ts)))


    End Function


    Function FG_OFFset(ByVal OFFset As Double) As Integer

        ts = "VOLTage:OFFSet " & OFFset
        ilwrt(FG_Dev, ts, CInt(Len(ts)))



    End Function

    Function FG_Frequency(ByVal FREQ_Hz As Double) As Integer

        ts = "FREQuency " & FREQ_Hz '& "kHz"
        ilwrt(FG_Dev, ts, CInt(Len(ts)))


    End Function

    Function FG_Function(ByVal mode As String) As Integer
        'FUNCtion {SINusoid|SQUare|RAMP|PULSe|NOISe|DC|USER}

        ts = "FUNCtion " & mode
        ilwrt(FG_Dev, ts, CInt(Len(ts)))

    End Function


    'Function FG_Pulse(ByVal dcycle As Integer, ByVal rise_edge As Integer, ByVal fall_edge As Integer, ByVal unit As String) As Integer
    '    ' [SOURce[1|2]]:PULSe:DCYCle

    '    ts = "PULSe:DCYCle " & dcycle
    '    ilwrt(FG_Dev, ts, CInt(Len(ts)))

    '    'rise_edge -->[SOURce[1|2]]:PULSe:TRANsition[:LEADing]

    '    ts = "PULSe:TRANsition:LEADing " & rise_edge & unit
    '    ilwrt(FG_Dev, ts, CInt(Len(ts)))


    '    ts = "PULSe:TRANsition:TRAiling " & fall_edge & unit
    '    ilwrt(FG_Dev, ts, CInt(Len(ts)))



    'End Function



    Function FG_Pulse(ByVal dcycle As Integer, ByVal R_edge As Integer, ByVal F_edge As Integer, ByVal unit As String) As Integer
        ' [SOURce[1|2]]:PULSe:DCYCle


        'rise_edge -->[SOURce[1|2]]:PULSe:TRANsition[:LEADing]

        ts = "PULSe:TRANsition:LEADing " & R_edge & unit
        ilwrt(FG_Dev, ts, CInt(Len(ts)))


        ts = "PULSe:TRANsition:TRAiling " & F_edge & unit
        ilwrt(FG_Dev, ts, CInt(Len(ts)))


        ts = "PULSe:DCYCle " & dcycle
        ilwrt(FG_Dev, ts, CInt(Len(ts)))

    End Function


End Module
