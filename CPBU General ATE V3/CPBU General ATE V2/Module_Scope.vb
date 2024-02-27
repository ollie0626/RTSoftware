Module Module_Scope

    ' scope sel option
    ' 0: R&S
    ' 1: Tek 7 series
    ' 2: Agilent
    ' 3: Tek MSO series
    Public osc_sel As Integer
    Public cmd As String





    Public Scope_folder As String

    Public Scope_name() As String
    Public Scope_IF() As String

    Public Scope_Dev As Integer
    Public Scope_format As String = "PNG"
    Public Scope_Addr As Integer


    Public Scope_num As Integer

    Public BW_20M As String = "TWEnty"
    Public BW_150M As String = "ONEfifty"
    Public BW_500M As String = "FULl" '"FIVe"
    Public RS_Scope As Boolean = False
    Public RS_Scope_EN As Boolean = False
    Public TDS2024_scope As Boolean = False
    Public TDS5054_scope As Boolean = False


    Public RS_Scope_Dev As String
    Public RS_vi As Integer

    Public Tek_Dev As String
    Public Tek_vi As Integer

    Public MSO_Dev As String
    Public MSO_vi As Integer

    Public Agilent_Dev As String
    Public Agilent_vi As Integer


    Public Scope_vpp As String = "PK2PK"
    Public Scope_Ton As String = "PWIDTH"
    Public Scope_Toff As String = "NWIDTH"
    Public Scope_freq As String = "FREQUENCY"
    Public Scope_vmax As String = "MAXimum"
    Public Scope_vmean As String = "MEAN"
    Public Scope_vmin As String = "MINImum"
    Public Scope_cursor_type As String = "WAVEform"


    Public RS_Scope_vpp As String = "PDELta" 'Peak-to-peak value of the waveform
    Public RS_Scope_Ton As String = "PPULse"
    Public RS_Scope_Toff As String = "NPULse"
    Public RS_Scope_freq As String = "FREQuency"
    Public RS_Scope_vmax As String = "MAXimum"
    Public RS_Scope_vmean As String = "MEAN"
    Public RS_Scope_vmin As String = "MINimum"


    Public Meas_mean As String = "MEAN"
    Public Meas_max As String = "MAX"
    Public Meas_min As String = "MINI"
    Public Scope_Meas As String = "VAL"

    Public RS_Meas_mean As String = "AVG"
    Public RS_Meas_max As String = "PPEak"
    Public RS_Meas_min As String = "NPEak"
    Public RS_Scope_Meas As String = "ACTual"

    Public Meas_max_value As Double = 9 * 10 ^ 30

    Public label_XPOS As Double = 0.2
    Public label_YPOS As Double = 0.5
    Public RS_label_XPOS As Integer = 5  'R&S設定REL %為主
    Public RS_label_YPOS As Integer = 5 'R&S設定REL %為主

    Dim read_error As Integer

    Dim ts As String = ""
    Dim unit_value As String = ""

    Public RS_RES_MES As String = "MEPosition"
    Public RS_RES_CURSOR As String = "CUPosition"
    Public RS_DISP_PREV As String = "PREV"
    Public RS_DISP_FLOA As String = "FLOA"
    Public RS_DISP_DOCK As String = "DOCK"

    Sub Docommand(ByVal cmd As String)

        Dim Dev As String = RS_Scope_Dev
        Dim vi As Integer

        Select Case osc_sel
            Case 0
                Dev = RS_Scope_Dev
                vi = RS_vi
            Case 1
                Dev = Tek_Dev
                vi = Tek_vi
            Case 2
                Dev = Agilent_Dev
                vi = Agilent_vi
            Case 3
                Dev = MSO_Dev
                vi = MSO_vi
        End Select
        visa_write(Dev, vi, cmd)
    End Sub

    Function DoQueryNumber(ByVal cmd As String) As Double
        Dim res As Double
        Dim Dev As String = RS_Scope_Dev
        Dim vi As Integer

        Select Case osc_sel
            Case 0
                Dev = RS_Scope_Dev
                vi = RS_vi
            Case 1
                Dev = Tek_Dev
                vi = Tek_vi
            Case 2
                Dev = Agilent_Dev
                vi = Agilent_vi
            Case 3
                Dev = MSO_Dev
                vi = MSO_vi
        End Select

        visa_write(Dev, vi, cmd)
        visa_status = viRead(vi, visa_response, Len(visa_response), retcount)

        If (retcount > 0) Then
            res = Val(Mid(visa_response, 1, retcount - 1))
        Else
            res = 0
        End If

        Return res
    End Function


    'source_num =1~4
    'Tek Measure error= 99.00000000000E+36\n
    'R&S Measure error 會傳最大的scale值

    Function RS_visa(ByVal open As Boolean) As Integer
        If open = True Then
            viOpenDefaultRM(defaultRM)
            viOpen(defaultRM, RS_Scope_Dev, VI_NO_LOCK, 5000, RS_vi)
            RS_Scope_EN = True
        Else
            viClose(RS_vi)
            viClose(defaultRM)
            RS_Scope_EN = False
        End If


    End Function

    Function RS_Local() As Integer

        ts = "&GTL"
        visa_write(RS_Scope_Dev, RS_vi, ts)
        Delay(100)
    End Function

    Function RS_View(ByVal view_ON As Boolean) As Integer

        If view_ON = True Then
            ts = "SYSTem:DISPlay:UPDate ON"
        Else
            ts = "SYSTem:DISPlay:UPDate OFF"
        End If

        visa_write(RS_Scope_Dev, RS_vi, ts)
        Delay(100)

    End Function


    Function Display_reset() As Integer
        'If RS_Scope = False Then
        '    'clearing of persistence data.
        '    ts = "DISplay:PERSistence:RESET"
        '    ilwrt(Scope_Dev, ts, CInt(Len(ts)))
        'Else
        '    'clearing of persistence data.
        '    '重新累積
        '    ts = "DISPlay:PERSistence:RESet"
        '    visa_write(RS_Scope_Dev, RS_vi, ts)
        'End If

        Select Case osc_sel
            Case 0
                cmd = "DISPlay:PERSistence:RESet"
            Case 1
                cmd = "DISplay:PERSistence:RESET"
            Case 2
                cmd = ":DISPlay:PERSistence MINimum"
            Case 3

        End Select
        Docommand(cmd)
        Delay(10)
    End Function

    'Function RS_Display_reset() As Integer
    '    'clearing of persistence data.
    '    '重新累積
    '    ts = "DISPlay:PERSistence:RESet"
    '     visa_write(RS_Scope_Dev,RS_vi, ts)
    'End Function

    Function FastAcq_ONOFF(ByVal ONOFF As String) As Integer

        If RS_Scope = False Then
            ts = "FASTAcq:STATE " & ONOFF
        End If
        ilwrt(Scope_Dev, ts, CInt(Len(ts)))
        Delay(50)
        ' Display_reset()
    End Function




    Function Display_persistence(ByVal PERSistence_ON As Boolean) As Integer
        'This command sets or queries the persistence aspect of the display.
        'DISplay:PERSistence {OFF|INFPersist|VARpersist}
        'OFF disables the persistence aspect of the display.
        'INFPersist sets a display mode where any pixels, once touched by samples, remain set until cleared by a mode change.
        'VARPersist sets a display mode where set pixels are gradually dimmed.

        'If RS_Scope = False Then
        '    If PERSistence_ON = True Then
        '        '無限持續累積
        '        ts = "DISplay:PERSistence INFPersist"
        '    Else
        '        ts = "DISplay:PERSistence OFF"
        '    End If

        '    ilwrt(Scope_Dev, ts, CInt(Len(ts)))
        '    Display_reset()
        'Else
        '    If PERSistence_ON = True Then
        '        '無限持續累積
        '        ts = "DISplay:PERSistence ON"
        '        visa_write(RS_Scope_Dev, RS_vi, ts)
        '        ts = "DISplay:PERSistence:INFinite ON"
        '    Else
        '        ts = "DISplay:PERSistence:INFinite OFF"
        '        visa_write(RS_Scope_Dev, RS_vi, ts)
        '        ts = "DISplay:PERSistence OFF"
        '    End If

        '    visa_write(RS_Scope_Dev, RS_vi, ts)
        '    'clearing of persistence data.
        '    '重新累積
        '    ts = "DISPlay:PERSistence:RESet"
        '    visa_write(RS_Scope_Dev, RS_vi, ts)
        'End If


        Select Case osc_sel
            Case 0
                If PERSistence_ON Then
                    cmd = "DISplay:PERSistence ON"
                    Docommand(cmd)
                Else
                    cmd = "DISplay:PERSistence:INFinite OFF"
                    Docommand(cmd)
                End If
                Delay(10)
                cmd = "DISplay:PERSistence ON"
                Docommand(cmd)
                Delay(10)
                cmd = "DISPlay:PERSistence:RESet"
                Docommand(cmd)
            Case 1
                If PERSistence_ON Then
                    cmd = "DISplay:PERSistence INFPersist"
                    Docommand(cmd)
                Else
                    cmd = "DISplay:PERSistence OFF"
                    Docommand(cmd)
                End If
            Case 2
                If PERSistence_ON Then
                    cmd = ":DISPlay:PERSistence INFinite"
                    Docommand(cmd)
                Else
                    cmd = ":DISPlay:PERSistence MINimum"
                    Docommand(cmd)
                End If

            Case 3

        End Select




    End Function

    'Function RS_Display_persistence(ByVal PERSistence_ON As Boolean) As Integer
    '    'DISPlay:PERSistence {ON|OFF}
    '    'This command sets or queries the persistence aspect of the display.


    '    If PERSistence_ON = True Then
    '        '無限持續累積
    '        ts = "DISplay:PERSistence:INFinite ON"
    '         visa_write(RS_Scope_Dev,RS_vi, ts)
    '        ts = "DISplay:PERSistence ON"
    '    Else
    '        ts = "DISplay:PERSistence:INFinite OFF"
    '         visa_write(RS_Scope_Dev,RS_vi, ts)
    '        ts = "DISplay:PERSistence OFF"
    '    End If

    '     visa_write(RS_Scope_Dev,RS_vi, ts)

    '    RS_Display_reset()
    'End Function



    Function CHx_display(ByVal source_num As Integer, ByVal ONOFF As String) As Integer
        'This command sets or queries the displayed state of the specified channel waveform. 
        'The x can be channel 1 through 4.
        'SELect:CH<x> {<NR1>|OFF|ON}

        'If RS_Scope = False Then

        '    ts = "SELect:CH" & source_num & " " & ONOFF
        '    ilwrt(Scope_Dev, ts, CInt(Len(ts)))
        'Else
        '    'This command sets or queries the displayed state of the specified channel waveform. 
        '    'The x can be channel 1 through 4.
        '    'CHANnel<x>:STATe {OFF|ON}
        '    '設定CHANnel ON/FF
        '    ts = "CHANnel" & source_num & ":STATe " & ONOFF
        '    visa_write(RS_Scope_Dev, RS_vi, ts)
        'End If

        Select Case osc_sel
            Case 0
                cmd = String.Format("CHANnel{0}:STATe {1}", source_num, ONOFF)
            Case 1
                cmd = String.Format("SELect:CH{0} {1}", source_num, ONOFF)
            Case 2
                cmd = String.Format(":CHANnel{0}:DISPLAY {1}", source_num, ONOFF)
            Case 3
        End Select


        Docommand(cmd)
    End Function

    'Function RS_CHx_display(ByVal source_num As Integer, ByVal ONOFF As String) As Integer
    '    'This command sets or queries the displayed state of the specified channel waveform. 
    '    'The x can be channel 1 through 4.
    '    'CHANnel<x>:STATe {OFF|ON}
    '    '設定CHANnel ON/FF
    '    ts = "CHANnel" & source_num & ":STATe" & " " & ONOFF
    '     visa_write(RS_Scope_Dev,RS_vi, ts)
    'End Function

    Function CouplingSel(ByVal coupling As String) As Integer
        Dim sel As Integer = 0


        Select Case coupling
            Case "DC (1MΩ)"
                sel = 0
            Case "AC"
                sel = 1
            Case "DC (50Ω)"
                sel = 2
        End Select

        Return sel
    End Function


    Function CHx_coupling(ByVal source_num As Integer, ByVal coupling As String) As Integer
        'This command sets or queries the input attenuator coupling setting for the specified channel.
        'CH<x>:COUPling {AC|DC|GND|DCREJect}

        'CH<x>:TERmination
        '50 Ω or 1,000,000 Ω. (50.0E+0, 1.0E+6)
        'DC (1MΩ)
        'AC
        'DC (50Ω)


        'If RS_Scope = False Then
        '    Select Case coupling

        '        Case "DC (1MΩ)"

        '            ts = "CH" & source_num & ":TERmination 1.0E+6"
        '            ilwrt(Scope_Dev, ts, CInt(Len(ts)))

        '            ts = "CH" & source_num & ":COUPling DC"

        '        Case "AC"

        '            ts = "CH" & source_num & ":COUPling AC"

        '        Case "DC (50Ω)"

        '            ts = "CH" & source_num & ":TERmination 50.0E+0"
        '            ilwrt(Scope_Dev, ts, CInt(Len(ts)))

        '            ts = "CH" & source_num & ":COUPling DC"

        '    End Select
        '    ts = "CH" & source_num & ":COUPling " & coupling
        '    ilwrt(Scope_Dev, ts, CInt(Len(ts)))
        'Else
        '    'This command sets or queries the input attenuator coupling setting for the specified channel.
        '    'CHANnel<x>:COUPling {DC|AC|DCLimit}
        '    'DC:50ohm,DCLimit:1M0ohm
        '    '設定碳棒阻抗匹配
        '    Select Case coupling

        '        Case "DC (1MΩ)"
        '            ts = "CHANnel" & source_num & ":COUPling DCLimit"
        '        Case "AC"

        '            ts = "CHANnel" & source_num & ":COUPling AC"

        '        Case "DC (50Ω)"

        '            ts = "CHANnel" & source_num & ":COUPling DC"

        '    End Select
        '    visa_write(RS_Scope_Dev, RS_vi, ts)
        'End If

        ' 0: DC 1M ohm
        ' 1: AC
        ' 2: DC 50 ohm
        Select Case osc_sel
            Case 0 ' R&S Scope
                Select Case CouplingSel(coupling)
                    Case 0
                        cmd = String.Format("CHANnel{0}:COUPling DCLimit", source_num)
                        Docommand(cmd)
                    Case 1
                        cmd = String.Format("CHANnel{0}:COUPling AC", source_num)
                        Docommand(cmd)
                    Case 2
                        cmd = String.Format("CHANnel{0}:COUPling DC", source_num)
                        Docommand(cmd)
                End Select
            Case 1 ' Tek Scope
                Select Case CouplingSel(coupling)
                    Case 0
                        cmd = String.Format("CH{0}:TERmination 1.0E+6", source_num)
                        Docommand(cmd)
                        Delay(10)
                        cmd = String.Format("CH{0}:COUPling DC", source_num)
                        Docommand(cmd)
                    Case 1
                        cmd = String.Format("CH{0}:COUPling AC")
                        Docommand(cmd)
                    Case 2
                        cmd = String.Format("CH{0}:COUPling DC")
                        Docommand(cmd)
                End Select
            Case 2 ' Agilent Scope
                Select Case CouplingSel(coupling)
                    Case 0
                        cmd = String.Format(":CHANnel{0}:INPut DC", source_num)
                        Docommand(cmd)
                    Case 1
                        cmd = String.Format(":CHANnel{0}:INPut AC", source_num)
                        Docommand(cmd)
                    Case 2
                        cmd = String.Format(":CHANnel{0}:INPut DC50", source_num)
                        Docommand(cmd)
                End Select
            Case 3 ' MSO Scope

        End Select




    End Function

    'Function RS_CHx_coupling(ByVal source_num As Integer, ByVal coupling As String) As Integer
    '    'This command sets or queries the input attenuator coupling setting for the specified channel.
    '    'CHANnel<x>:COUPling {DC|AC|DCLimit}
    '    'DC:50ohm,DCLimit:1M0ohm
    '    '設定碳棒阻抗匹配

    '    ts = "CHANnel" & source_num & ":COUPling " & " " & coupling
    '     visa_write(RS_Scope_Dev,RS_vi, ts)

    'End Function

    Function CHx_Bandwidth(ByVal source_num As Integer, ByVal mode As String) As Integer
        'CH<x>:BANdwidth {TWEnty|ONEfifty|FULl|<NR3>}
        ' TWEnty
        'This sets the upper bandwidth limit to 20 MHz.
        '• ONEfifty
        'This sets the upper bandwidth limit to 150 MHz.
        '• FIVe
        'This argument sets the upper bandwidth limit to 500 MHz.
        '• FULl
        'This disables any optional bandwidth limiting. The specified
        'channel operates at its maximum attainable bandwidth.
        '• <NR3>
        'This argument is a double-precision ASCII string. The
        'instrument rounds this value to an available bandwidth using
        'geometric rounding and then uses this value set the upper
        'bandwidth.
        'mode= BW_20M or BW_150M

        'If RS_Scope = False Then
        '    Select Case mode
        '        Case "20MHz"
        '            mode = BW_20M
        '        Case "Full"
        '            mode = BW_500M
        '    End Select
        '    ts = "CH" & source_num & ":BANdwidth " & mode
        '    ilwrt(Scope_Dev, ts, CInt(Len(ts)))
        'Else
        '    'CHANnel<x>:BANDwidth {B20|B200|FULL}
        '    'B20
        '    'This sets the upper bandwidth limit to 20 MHz.
        '    'B200
        '    'This sets the upper bandwidth limit to 200 MHz.
        '    'FULL
        '    'Use full bandwidth
        '    '設定CHANnel頻寬
        '    Select Case mode
        '        Case "20MHz"
        '            mode = "B20"
        '        Case "Full"
        '            mode = "FULL"
        '    End Select
        '    ts = "CHANnel" & source_num & ":BANdwidth " & " " & mode
        '    visa_write(RS_Scope_Dev, RS_vi, ts)
        'End If

        Select Case osc_sel
            Case 0 ' R&S
                If mode = "20MHz" Then
                    mode = "B20"
                ElseIf mode = "Full" Then
                    mode = "FULL"
                End If
                cmd = String.Format("CHANnel{0}:BANdwidth {1}", source_num, mode)

            Case 1 ' Tek
                If mode = "20MHz" Then
                    mode = BW_20M
                ElseIf mode = "Full" Then
                    mode = BW_500M
                End If
                cmd = String.Format("CH{0}:BANdwidth {1}", source_num, mode)
            Case 2 ' Agilent
                If mode = "20MHz" Then
                    mode = "20e6"
                Else
                    mode = "ON"
                End If
                cmd = String.Format(":CHANnel{0}:BWLimit {1}", source_num, mode)
            Case 3

        End Select
        Docommand(cmd)

    End Function

    'Function RS_CHx_Bandwidth(ByVal source_num As Integer, ByVal mode As String) As Integer
    '    'CHANnel<x>:BANDwidth {B20|B200|FULL}
    '    'B20
    '    'This sets the upper bandwidth limit to 20 MHz.
    '    'B200
    '    'This sets the upper bandwidth limit to 200 MHz.
    '    'FULL
    '    'Use full bandwidth
    '    '設定CHANnel頻寬

    '    ts = "CHANnel" & source_num & ":BANdwidth " & " " & mode
    '     visa_write(RS_Scope_Dev,RS_vi, ts)

    'End Function


    Function CHx_scale(ByVal source_num As Integer, ByVal value As Double, ByVal unit As String) As Integer
        Select Case unit
            Case "mV"
                unit_value = value & "E-3"
            Case "V"
                unit_value = value & "E-0"
        End Select

        'If RS_Scope = False Then
        '    'This command sets or queries the vertical scale of the specified channel. 
        '    ts = "CH" & source_num & ":SCAle " & unit_value
        '    ilwrt(Scope_Dev, ts, CInt(Len(ts)))
        'Else
        '    'This command sets or queries the vertical scale of the specified channel. 
        '    'CHANnel<x>:SCALe <scale>
        '    'scale unit:V/div
        '    '設定CHANnel的SCALe為多少V/div
        '    ts = "CHANnel" & source_num & ":SCAle" & " " & unit_value
        '    visa_write(RS_Scope_Dev, RS_vi, ts)
        'End If

        Select Case osc_sel
            Case 0
                cmd = String.Format("CHANnel{0}:SCAle {1}", source_num, unit_value)
            Case 1
                cmd = String.Format("CH{0}:SCAle {1}", source_num, unit_value)
            Case 2
                cmd = String.Format(":CHANNEL{0}:SCALe {1}", unit_value)
        End Select

        Docommand(cmd)

    End Function

    'Function RS_CHx_scale(ByVal source_num As Integer, ByVal value As Double, ByVal unit As String) As Integer
    '    'This command sets or queries the vertical scale of the specified channel. 
    '    'CHANnel<x>:SCALe <scale>
    '    'scale unit:V/div
    '    '設定CHANnel的SCALe為多少V/div

    '    Select Case unit
    '        Case "mV"
    '            unit_value = value & "E-3"
    '        Case "V"
    '            unit_value = value & "E-0"
    '    End Select

    '    ts = "CHANnel" & source_num & ":SCAle" & " " & unit_value
    '     visa_write(RS_Scope_Dev,RS_vi, ts)

    'End Function

    Function CHx_position(ByVal source_num As Integer, ByVal value As Double) As Integer

        'If RS_Scope = False Then
        '    'This command sets or queries the vertical position of the specified channel.
        '    'CH<x>:POSition <NR3>
        '    '<NR3> is the position value, in divisions from the center graticule, ranging from 8.000 to -8.000 divisions.
        '    ts = "CH" & source_num & ":POSition " & value & "E+00"
        '    ilwrt(Scope_Dev, ts, CInt(Len(ts)))

        'Else
        '    'This command sets or queries the vertical position of the specified channel.
        '    'CHANnel<x>:POSition <position>
        '    '<position>, ranging from 5 to -5 divisions,unit:div
        '    '設定CHANnel的水平軸位置
        '    ts = "CHANnel" & source_num & ":POSition " & value
        '    visa_write(RS_Scope_Dev, RS_vi, ts)
        'End If


        Select Case osc_sel
            Case 0
                cmd = String.Format("CHANnel{0}:POSition {1}", source_num, value)
            Case 1
                cmd = String.Format("CH{0}:POSition {1}", source_num, value)
            Case 2
                cmd = String.Format(":CHANnel{0}:SCALe?")
                Dim res As Double = DoQueryNumber(cmd)
                Delay(300)
                cmd = String.Format(":CHANnel{0}:OFFSet {1}", source_num, value * res)
            Case 3

        End Select
        Docommand(cmd)





    End Function

    'Function RS_CHx_position(ByVal source_num As Integer, ByVal value As Double) As Double
    '    'This command sets or queries the vertical position of the specified channel.
    '    'CHANnel<x>:POSition <position>
    '    '<position>, ranging from 5 to -5 divisions,unit:div
    '    '設定CHANnel的水平軸位置

    '    ts = "CHANnel" & source_num & ":POSition " & " " & value
    '     visa_write(RS_Scope_Dev,RS_vi, ts)

    'End Function

    Function CHx_OFFSET(ByVal source_num As Integer, ByVal value As Double) As Integer

        'If RS_Scope = False Then
        '    'This command sets or queries the vertical offset for the specified channel. The
        '    'channel is specified by x. The value of x can range from 1 through 4. This
        '    'command is equivalent to selecting Offset from the Vertical menu.
        '    'CH<x>:OFFSet <NR3>
        '    ts = "CH" & source_num & ":OFFSet " & value
        '    ilwrt(Scope_Dev, ts, CInt(Len(ts)))
        'Else
        '    'This command sets offset voltage of the specified channel.
        '    'CHANnel<x>:OFFSet <offset>
        '    '<offset>   increment:0.01,unit:V
        '    '設定CHANnel的offset

        '    ts = "CHANnel" & source_num & ":OFFSet " & " " & value
        '    visa_write(RS_Scope_Dev, RS_vi, ts)
        'End If

        Select Case osc_sel
            Case 0
                cmd = String.Format("CHANnel{0}:OFFSet {1}", source_num, value)
            Case 1
                cmd = String.Format("CH{0}:OFFSet {1}", source_num, value)
            Case 2
                cmd = String.Format(":CHANnel{0}:OFFSet {1}", source_num, value)
            Case 3
        End Select
        Docommand(cmd)

    End Function

    'Function RS_CHx_offset(ByVal source_num As Integer, ByVal value As Double) As Double
    '    'This command sets offset voltage of the specified channel.
    '    'CHANnel<x>:OFFSet <offset>
    '    '<offset>   increment:0.01,unit:V
    '    '設定CHANnel的offset

    '    ts = "CHANnel" & source_num & ":OFFSet " & " " & value
    '     visa_write(RS_Scope_Dev,RS_vi, ts)

    'End Function

    Function CHx_label(ByVal source_num As Integer, ByVal name As String) As Integer
        Dim temp As Double
        Dim YPOS As Integer
        'Defines or returns the label for the channel(waveform)
        'CH<x>:LABel:NAMe <str> 
        '<str> is an alphanumeric character string, ranging from 1 through 32 characters in length. Ex: CH2:LABEL:NAMe "Pressure"

        'Sets or returns the X display coordinate for the channel waveform label
        'CH<x>:LABel:XPOS <NR1>
        '<NR1>: Arguments should be integers ranging from 0 through 10.

        'Sets or returns the Y display coordinate
        'CH<x>:LABel:YPOS <NR1>
        'If RS_Scope = False Then
        '    ts = "CH" & source_num & ":LABEL:NAMe """ & name & """"
        '    ilwrt(Scope_Dev, ts, CInt(Len(ts)))
        '    Delay(10)
        '    ts = "CH" & source_num & ":LABEL:XPOS " & label_XPOS
        '    ilwrt(Scope_Dev, ts, CInt(Len(ts)))
        '    Delay(10)
        '    ts = "CH" & source_num & ":LABEL:YPOS " & label_YPOS
        '    ilwrt(Scope_Dev, ts, CInt(Len(ts)))
        '    Delay(10)
        'Else
        '    ts = "CHANnel" & source_num & ":POSition?"
        '    visa_write(RS_Scope_Dev, RS_vi, ts)
        '    visa_status = viRead(RS_vi, visa_response, Len(visa_response), retcount)
        '    If retcount > 0 Then
        '        temp = Val(Mid(visa_response, 1, retcount - 1))
        '        YPOS = (5 - Math.Floor(temp)) * 10 - RS_label_YPOS

        '        ts = "DISplay:SIGNal:LABel:ADD 'Label1',C" & source_num & "W1," & "'" & name & "'" & ",REL," & RS_label_XPOS & "," & YPOS
        '        'ts = "DISplay:SIGNal:LABel:ADD 'Label1',C" & source_num & "W1," & "'" & name & "'" & ",ABS," & x & "," & y
        '        visa_write(RS_Scope_Dev, RS_vi, ts)
        '    End If
        '    Delay(10)
        'End If

        Dim value As Double
        Select Case osc_sel
            Case 0
                cmd = "CHANnel" & source_num & ":POSition?"
                value = DoQueryNumber(cmd)
                Delay(10)
                YPOS = (5 - Math.Floor(value)) * 10 - RS_label_YPOS
                cmd = "DISplay:SIGNal:LABel:ADD 'Label1',C" & source_num & "W1," & "'" & name & "'" & ",REL," & RS_label_XPOS & "," & YPOS
                Docommand(cmd)
            Case 1
                cmd = "CH" & source_num & ":LABEL:NAMe """ & name & """"
                Docommand(cmd)
                Delay(10)
                cmd = "CH" & source_num & ":LABEL:XPOS " & label_XPOS
                Docommand(cmd)
                Delay(10)
                cmd = "CH" & source_num & ":LABEL:YPOS " & label_YPOS
                Docommand(cmd)
                Delay(10)
            Case 2

            Case 3
        End Select







    End Function

    'Function RS_CHx_label(ByVal source_num As Integer, ByVal name As String, ByVal x As Double, ByVal y As Double) As Integer
    '    'DISplay:SIGNal:LABel:ADD <LabelID>,<Source>,<LableText>,<PositionMode>,<XPosition>,<YPosition>

    '    'LabelID:Set lable ID

    '    'Source:  C1W1＞set CH1,C2W1＞set CH2,C3W1>set CH3,C4W1＞set CH4

    '    'LableText:String with the label text that is shown on thr display

    '    'PositionMode:ABS>Position in time and voltage values.Absolute positions move with the waveform display when the scales,the vertical posiotion or offest,
    '    'or the reference point are changed.
    '    'REL:Fixed label position in percent.

    '    'XPosition:Horizontal position of the label text.

    '    'YPosition:Vertical position of the label text.

    '    'EX:DISplay:SIGNal:LABel:ADD 'Label1',C3W1,'VOUT',REL,10,30
    '    'LabelID:Label1,Source:CH3,LableText:VOUT,XPosition:10%,YPosition:30%

    '    'EX:DISplay:SIGNal:LABel:ADD 'Label1',C4W1,'VOUT',ABS,0,-5
    '    'LabelID:Label1,Source:CH4,LableText:VOUT,XPosition:0s,YPosition:-5V


    '    ts = "DISplay:SIGNal:LABel:ADD 'Label1',C" & source_num & "W1," & "'" & name & "'" & ",ABS," & x & "," & y
    '     visa_write(RS_Scope_Dev,RS_vi, ts)

    '    Delay(10)


    'End Function

    Function H_reclength(ByVal value As Integer) As Integer
        If RS_Scope = False Then

            'This command sets or queries the record length.
            'HORizontal:MODE:RECOrdlength <NR1>
            ts = "HORizontal:RECOrdlength " & value
            ilwrt(Scope_Dev, ts, CInt(Len(ts)))
            Delay(10)
            'Arguments AUTO selects the automatic horizontal model. Auto mode attempts to keep
            'record length constant as you change the time per division setting. Record length
            'is read only.
            'CONSTANT selects the constant horizontal model. Constant mode attempts to
            'keep sample rate constant as you change the time per division setting. Record
            'length is read only.
            'MANUAL selects the manual horizontal model. Manual mode lets you change
            'sample mode and record length. Time per division or Horizontal scale is read only.

            'NOTE: 只要寫這個command就會變成MANUAL Mode
            ts = "HORizontal:MODE AUTO"
            ilwrt(Scope_Dev, ts, CInt(Len(ts)))
            Delay(10)
        Else
            'This command sets or queries the record length.
            'ACQuire:POINts:VALue <RECOrdlength>
            '<RECOrdlength> Range:1000 to 1000000000
            '要設定"ACQuire:POINts:AUTO RECLength"才能寫入reclength

            ts = "ACQuire:POINts:AUTO RECLength"
            visa_write(RS_Scope_Dev, RS_vi, ts)
            ts = "ACQuire:POINts:VALue" & " " & value
            visa_write(RS_Scope_Dev, RS_vi, ts)
        End If


        Select Case osc_sel
            Case 0
                cmd = "ACQuire:POINts:AUTO RECLength"
                Docommand(10)
                cmd = String.Format("ACQuire:POINts:VALue {0}", value)
                Docommand(cmd)
            Case 1
                cmd = String.Format("HORizontal:RECOrdlength {0}", value)
                Docommand(cmd)
                Docommand(10)
                cmd = String.Format("HORizontal:MODE AUTO")
                Docommand(cmd)
            Case 2
            Case 3

        End Select



    End Function

    'Function RS_H_reclength(ByVal value As Integer) As Integer
    '    'This command sets or queries the record length.
    '    'ACQuire:POINts:VALue <RECOrdlength>
    '    '<RECOrdlength> Range:1000 to 1000000000
    '    '要設定"ACQuire:POINts:AUTO RECLength"才能寫入reclength

    '    ts = "ACQuire:POINts:AUTO RECLength"
    '     visa_write(RS_Scope_Dev,RS_vi, ts)
    '    ts = "ACQuire:POINts:VALue" & " " & value
    '     visa_write(RS_Scope_Dev,RS_vi, ts)
    'End Function

    Function H_Roll(ByVal mode As String) As Integer
        'If RS_Scope = False Then
        '    '            This command sets or queries the Roll Mode status. Use Roll Mode when you
        '    'want to view data at very slow sweep speeds. It is useful for observing data
        '    'samples on the screen as they occur. This command is equivalent to selecting
        '    'Horizontal/Acquisition Setup from the Horiz/Acq menu, selecting the Acquisition
        '    'tab, and setting the Roll Mode to Auto or Off.

        '    'HORizontal:ROLL {AUTO|OFF|ON}
        '    ts = "HORizontal:ROLL " & mode
        '    ilwrt(Scope_Dev, ts, CInt(Len(ts)))
        'Else
        '    '<Mode> AUTO | OFF
        '    ts = "TIMebase:ROLL:ENABle " & mode
        '    visa_write(RS_Scope_Dev, RS_vi, ts)
        'End If

        Select Case osc_sel
            Case 0
                cmd = String.Format("HORizontal:ROLL {0}", mode)
            Case 1
                cmd = String.Format("TIMebase:ROLL:ENABle {0}", mode)
            Case 2
                ':TIMebase:ROLL:ENABLE {{ON | 1} | {OFF | 0}}
                cmd = String.Format(":TIMebase:ROLL:ENABLE {0}", mode)
            Case 3

        End Select
        Docommand(cmd)
    End Function

    'Function RS_H_Roll(ByVal mode As String) As Integer
    '    '<Mode> AUTO | OFF
    '    ts = "TIMebase:ROLL:ENABle " & mode
    '     visa_write(RS_Scope_Dev,RS_vi, ts)
    'End Function


    Function H_scale(ByVal value As Double, ByVal unit As String) As Integer
        Select Case unit
            Case "ns"
                unit_value = value & "E-9"
            Case "us"
                unit_value = value & "E-6"
            Case "ms"
                unit_value = value & "E-3"
            Case "s"
                unit_value = ""
        End Select

        'If RS_Scope = False Then
        '    'This command sets or queries the horizontal scale.
        '    'HORizontal:MODE:SCAle <NR1>
        '    ts = "HORizontal:SCAle " & unit_value
        '    ilwrt(Scope_Dev, ts, CInt(Len(ts)))
        'Else
        '    'This command sets or queries the horizontal scale.
        '    'Timebase:SCALe <Timebase>,Timebase unit:s/div
        '    '設定Timebase的SCALe
        '    ts = "Timebase:SCALe " & " " & unit_value
        '    visa_write(RS_Scope_Dev, RS_vi, ts)
        'End If


        Select Case osc_sel
            Case 0
                cmd = String.Format("Timebase:SCALe {0}", unit_value)
            Case 1
                cmd = String.Format("HORizontal:SCAle {0}", unit_value)
            Case 2
                cmd = String.Format(":TIMebase:SCALe {0}", unit_value)
            Case 3

        End Select

        Docommand(cmd)



    End Function

    Function H_scale_now() As Double
        'If RS_Scope = False Then
        '    'This command sets or queries the horizontal scale.
        '    'HORizontal:MODE:SCAle <NR1>


        '    ts = "HORizontal:SCAle?"
        '    ilwrt(Scope_Dev, ts, CInt(Len(ts)))
        '    ilrd(Scope_Dev, ValueStr, ARRAYSIZE)

        '    If ibcntl > 0 Then
        '        Return Val(Mid(ValueStr, 1, (ibcntl - 1)))
        '    Else
        '        Return 0
        '    End If

        'Else
        '    ts = "Timebase:SCALe?"
        '    visa_write(RS_Scope_Dev, RS_vi, ts)
        '    visa_status = viRead(RS_vi, visa_response, Len(visa_response), retcount)
        '    If visa_status = VI_ERROR_CONN_LOST Then
        '        viOpen(defaultRM, RS_Scope_Dev, VI_NO_LOCK, 2000, RS_vi)
        '    End If

        '    If retcount > 0 Then

        '        Return Val(Mid(visa_response, 1, retcount - 1))
        '    Else

        '        visa_status = viRead(RS_vi, visa_response, Len(visa_response), retcount)
        '        If visa_status = VI_ERROR_CONN_LOST Then
        '            viOpen(defaultRM, RS_Scope_Dev, VI_NO_LOCK, 2000, RS_vi)
        '        End If

        '        While retcount = 0
        '            System.Windows.Forms.Application.DoEvents()
        '            read_error = read_error + 1
        '            If (read_error = 100) Or (run = False) Then
        '                Return 0
        '                Exit Function
        '            End If
        '            Delay(10)
        '            visa_status = viRead(RS_vi, visa_response, Len(visa_response), retcount)
        '        End While

        '        If retcount > 0 Then
        '            Return Val(Mid(visa_response, 1, retcount - 1))
        '        End If
        '    End If
        'End If


        Select Case osc_sel
            Case 0
                cmd = "Timebase:SCALe?"
            Case 1
                cmd = "HORizontal:SCAle?"
            Case 2
                cmd = ":TIMEBASE:SCALE?"
            Case 3
        End Select


        Return DoQueryNumber(cmd)
    End Function


    'Function RS_H_scale(ByVal value As Double, ByVal unit As String) As Integer
    '    'This command sets or queries the horizontal scale.
    '    'Timebase:SCALe <Timebase>,Timebase unit:s/div
    '    '設定Timebase的SCALe


    '    Select Case unit
    '        Case "ns"
    '            unit_value = value & "E-9"
    '        Case "us"
    '            unit_value = value & "E-6"
    '        Case "ms"
    '            unit_value = value & "E-3"
    '    End Select

    '    ts = "Timebase:SCALe " & " " & unit_value
    '     visa_write(RS_Scope_Dev,RS_vi, ts)

    'End Function


    Function H_Samplerate(ByVal value As Double, ByVal unit As String) As Integer

        'MS/s.
        'This command sets or queries the sample rate.
        'HORizontal:MODE:SAMPLERate <NR1>
        Select Case unit

            Case "GS/s"
                unit_value = value & "E9"
            Case "MS/s"
                unit_value = value & "E6"
            Case "kS/s"
                unit_value = value & "E3"
            Case "S/s"
                unit_value = value
        End Select

        'If RS_Scope = False Then
        '    ts = "HORizontal:MODE:SAMPLERate " & unit_value
        '    ilwrt(Scope_Dev, ts, CInt(Len(ts)))
        '    ts = "HORizontal:MODE AUTO"
        '    ilwrt(Scope_Dev, ts, CInt(Len(ts)))
        'Else
        '    ts = "ACQuire:POINts:AUTO RESolution"
        '    visa_write(RS_Scope_Dev, RS_vi, ts)
        '    ts = "ACQuire:SRATe " & unit_value
        '    visa_write(RS_Scope_Dev, RS_vi, ts)
        'End If

        Select Case osc_sel
            Case 0
                cmd = "ACQuire:POINts:AUTO RESolution"
                Docommand(cmd)
                cmd = String.Format("ACQuire:SRATe {0}", unit_value)
                Docommand(cmd)
            Case 1
                cmd = String.Format("HORizontal:MODE:SAMPLERate {0}", unit_value)
                Docommand(cmd)
                cmd = "HORizontal:MODE AUTO"
                Docommand(cmd)
            Case 2
                ' need to test
                cmd = String.Format(":ACQuire:SRATe:ANALog {0}", unit_value)
                Docommand(cmd)

                cmd = ":ACQuire:SRATe:ANALog:AUTO ON"
                Docommand(cmd)

            Case 3
        End Select

    End Function

    'Function RS_H_Samplerate(ByVal value As Double, ByVal unit As String) As Integer
    '    'This command sets or queries the sample rate.
    '    'ACQuire:SRATe <Samplerate>
    '    '<Samplerate> Range:2 to 20E+20
    '    '要設定"ACQuire:POINts:AUTO RESolution"才能寫入Samplerate
    '    Select Case unit
    '        Case "GS/s"
    '            unit_value = value & "E9"
    '        Case "MS/s"
    '            unit_value = value & "E6"
    '        Case "kS/s"
    '            unit_value = value & "E3"
    '        Case "S/s"
    '            unit_value = value
    '    End Select

    '    ts = "ACQuire:POINts:AUTO RESolution"
    '     visa_write(RS_Scope_Dev,RS_vi, ts)
    '    ts = "ACQuire:SRATe " & unit_value
    '     visa_write(RS_Scope_Dev,RS_vi, ts)
    'End Function

    Function H_position(ByVal value As Double) As Integer
        'If RS_Scope = False Then
        '    'HORizontal[:MAIn]:POSition <NR3>
        '    '<NR3> argument can range from 0 to ??00
        '    ts = "HORizontal:POSition " & value
        '    ilwrt(Scope_Dev, ts, CInt(Len(ts)))
        'Else
        '    'Defines the time distance between the reference point and the trigger point.
        '    'TIMebase:HORizontal:POSition <RescaleCenter Time>
        '    '<RescaleCenter Time> argument can range from -100E+24 t0 100E+24
        '    ts = "TIMebase:HORizontal:POSition 0"
        '    visa_write(RS_Scope_Dev, RS_vi, ts)
        '    ts = "TIMebase:REFerence " & value
        '    visa_write(RS_Scope_Dev, RS_vi, ts)
        'End If

        Dim timeScale As Double = 0
        Select Case osc_sel
            Case 0
                cmd = "TIMebase:HORizontal:POSition 0"
                Docommand(cmd)
                cmd = String.Format("TIMebase:REFerence {0}", value)
                Docommand(cmd)
            Case 1
                cmd = String.Format("HORizontal:POSition {0}", value)
                Docommand(cmd)
            Case 2
                timeScale = DoQueryNumber(":TIMEBASE:SCALE?")
                Delay(300)
                cmd = String.Format(":TIMEBASE:POSITION {0}", value * timeScale)
                Docommand(cmd)
            Case 3
        End Select


    End Function

    'Function RS_H_position(ByVal value As Double) As Integer
    '    'Defines the time distance between the reference point and the trigger point.
    '    'TIMebase:HORizontal:POSition <RescaleCenter Time>
    '    '<RescaleCenter Time> argument can range from -100E+24 t0 100E+24

    '    ts = "TIMebase:HORizontal:POSition " & " " & value
    '     visa_write(RS_Scope_Dev,RS_vi, ts)

    'End Function

    Function RS_H_position_reference(ByVal value As Integer) As Double

        '"TIMebase:REFerence 50 ":讓H_Positopn的位置在示波器50%的位置

        ts = "TIMebase:REFerence" & " " & value
        visa_write(RS_Scope_Dev, RS_vi, ts)


    End Function


    Function Cursor_ONOFF(ByVal ONOFF As String) As Integer
        'If RS_Scope = False Then
        '    ts = "CURSor:STATE " & ONOFF
        '    ilwrt(Scope_Dev, ts, CInt(Len(ts)))
        'Else
        '    ts = "CURSor1:STATe " & ONOFF
        '    visa_write(RS_Scope_Dev, RS_vi, ts)
        '    Delay(10)
        'End If

        Select Case osc_sel
            Case 0
                cmd = String.Format("CURSor1:STATe {0}", ONOFF)
            Case 1
                cmd = String.Format("CURSor:STATE {0}", ONOFF)
            Case 2
                If ONOFF = "ON" Then
                    ONOFF = "MANual"
                End If
                cmd = String.Format(":MARKer:MODE {0}", ONOFF)
            Case 3
        End Select

        Docommand(cmd)

    End Function


    Function Cursor_ONOFF_check() As Boolean
        Dim ONOFF As Integer
        If RS_Scope = False Then
            ts = "CURSor:STATE?"
            ilwrt(Scope_Dev, ts, CInt(Len(ts)))
            ilrd(Scope_Dev, ValueStr, ARRAYSIZE)
            If ibcntl > 0 Then
                ONOFF = Val(Mid(ValueStr, 1, (ibcntl - 1)))
            End If
        Else
            ts = "CURSor1:STATe?"
            visa_write(RS_Scope_Dev, RS_vi, ts)
            visa_status = viRead(RS_vi, visa_response, Len(visa_response), retcount)
            ONOFF = Val(Mid(visa_response, 1, (retcount - 1)))
        End If


        If ONOFF = 1 Then
            Return True
        Else
            Return False
        End If


    End Function



    Function Cursor_set(ByVal type As String, ByVal x1 As Integer, ByVal x2 As Integer) As Integer
        If RS_Scope = False Then
            'This command sets or queries the cursor type.
            'CURSor:FUNCtion {OFF|HBArs|VBArs|SCREEN|WAVEform}
            'OFF removes the cursors from the display but does not change the cursor type.
            'HBArs specifies horizontal bar cursors, which measure in vertical units. 
            'VBArs specifies vertical bar cursors, which measure in horizontal units. 


            'This command sets or queries the source(s) for the currently selected cursor type.
            'CURSor:SOUrce<x> {CH<x>|MATH<x>|REF<x>}

            'This command sets or queries the cursor type for Screen mode.
            'CURSor:SCREEN:STYle {LINE_X|LINES|X}
            'LINES specifies the cursor style to be a line.
            'LINE_X specifies the cursor style to be a line with superimposed X.
            'X specifies the cursor style to be an X.

            'type=OFF|HBArs|VBArs|SCREEN|WAVEform



            ts = "CURSor:FUNCtion " & type
            ilwrt(Scope_Dev, ts, CInt(Len(ts)))


            ts = "CURSor:SOUrce1 " & "CH" & x1
            ilwrt(Scope_Dev, ts, CInt(Len(ts)))

            ts = "CURSor:SOUrce2 " & "CH" & x2
            ilwrt(Scope_Dev, ts, CInt(Len(ts)))




        Else
            'Set cursor on or off
            'CURSor<m>:STATe <state>
            'm:1 ,2 ; <state>: ON/OFF

            'Defines the type of the indicated cursor set.
            'CURSor<m>:FUNCtion <Type>
            'm:1 ,2 ; <type>: HORizontal/VERTical/PALRed(both HORizontal and VERTical cursor line pairs)

            'Defines thr source of the cursor measurement.
            'CURSor<m>:SOURce <cursorsource>
            'm:1 ,2 ;<cursorsource>: C1W1＞set CH1,C2W1＞set CH2,C3W1>set CH3,C4W1＞set CH4

            Select Case type

                Case "VBArs"

                    ts = "CURSor1:FUNCtion VERTical"

                Case "HBArs"

                    ts = "CURSor1:FUNCtion HORizontal"


                Case "SCREEN"

                    ts = "CURSor1:FUNCtion PALRed"

            End Select



            visa_write(RS_Scope_Dev, RS_vi, ts)

            Delay(10)

            ts = "CURSor1:SOURce C" & x1 & "W1"
            visa_write(RS_Scope_Dev, RS_vi, ts)

            'ts = "CURSor1:STATe " & " " & ONOFF
            'visa_status = viWrite(RS_vi, ts, Len(ts), retcount)
            Delay(10)
        End If

        ' need to check function application
        Select Case osc_sel
            Case 0
                Select Case type
                    Case "VBArs"
                        cmd = "CURSor1:FUNCtion VERTical"
                    Case "HBArs"
                        cmd = "CURSor1:FUNCtion HORizontal"
                    Case "SCREEN"
                        cmd = "CURSor1:FUNCtion PALRed"
                End Select
                Docommand(cmd)
                Delay(10)
                cmd = "CURSor1:SOURce C" & x1 & "W1"
                Docommand(cmd)

            Case 1
                cmd = "CURSor:FUNCtion " & type
                Docommand(cmd)
                Delay(10)
                cmd = "CURSor:SOUrce1 " & "CH" & x1
                Docommand(cmd)
                Delay(10)
                cmd = "CURSor:SOUrce2 " & "CH" & x2
                Docommand(cmd)
            Case 2
                ':MARKer:MODE {OFF | MANual | WAVeform | MEASurement | XONLy | YONLy}
                cmd = ":MARKer:MODE " & type
                Docommand(cmd)
                Delay(10)
            Case 3

        End Select





    End Function

    'Function RS_Cursor_set(ByVal type As String, ByVal ONOFF As String, ByVal x1 As Integer) As Integer
    '    'Set cursor on or off
    '    'CURSor<m>:STATe <state>
    '    'm:1 ,2 ; <state>: ON/OFF

    '    'Defines the type of the indicated cursor set.
    '    'CURSor<m>:FUNCtion <Type>
    '    'm:1 ,2 ; <type>: HORizontal/VERTical/PALRed(both HORizontal and VERTical cursor line pairs)

    '    'Defines thr source of the cursor measurement.
    '    'CURSor<m>:SOURce <cursorsource>
    '    'm:1 ,2 ;<cursorsource>: C1W1＞set CH1,C2W1＞set CH2,C3W1>set CH3,C4W1＞set CH4

    '    ts = "CURSor1:FUNCtion" & " " & type
    '     visa_write(RS_Scope_Dev,RS_vi, ts)

    '    ts = "CURSor1:SOURce" & " " & "C" & x1 & "W1"
    '     visa_write(RS_Scope_Dev,RS_vi, ts)

    '    ts = "CURSor1:STATe " & " " & ONOFF
    '     visa_write(RS_Scope_Dev,RS_vi, ts)



    'End Function


    Function Cursor_move(ByVal type As String, ByVal position1 As Double, ByVal position2 As Double) As Integer
        Dim value As Double
        If RS_Scope = False Then
            'move time: type="VBArs"
            'move volt: type="HBArs"
            ts = "CURSOR:" & type & ":POSITION1 " & position1 '& "E+00"
            ilwrt(Scope_Dev, ts, CInt(Len(ts)))
            Delay(10)
            ts = "CURSOR:" & type & ":POSITION2 " & position2 '& "E+00"
            ilwrt(Scope_Dev, ts, CInt(Len(ts)))
        Else
            'Defines the position of the vertical cursor line
            'CURSor<m>:X1Position <X1Position>
            'CURSor<m>:X2Position <X2Position>
            '<X2Position>,<X1Position>:Range 0 to 500

            'Defines the position of the horizontal cursor line
            'CURSor<m>:Y1Position <Y1Position>
            'CURSor<m>:Y2Position <Y2Position>
            '<Y2Position>,<Y1Position>:Range -50 to 50

            'If set to ON,the horizontal cursor lines follow the waveform.
            'CURSor<m>:TRACKing <trackcurve>
            '<trackcurve>:ON/OFF

            Select Case type
                Case "VBArs"

                    ts = "CURSor1:X1Position" & " " & position1
                    visa_write(RS_Scope_Dev, RS_vi, ts)
                    Delay(10)

                    ts = "CURSor1:X2Position" & " " & position2
                    visa_write(RS_Scope_Dev, RS_vi, ts)

                Case "HBArs"

                    ts = "CURSor1:Y1Position" & " " & position1
                    visa_write(RS_Scope_Dev, RS_vi, ts)
                    Delay(10)

                    ts = "CURSor1:Y2Position" & " " & position2
                    visa_write(RS_Scope_Dev, RS_vi, ts)
            End Select


        End If

        Delay(10)


        Select Case osc_sel
            Case 0
                Select Case type
                    Case "VBArs"
                        cmd = String.Format("CURSor1:X1Position {0}", position1)
                        Docommand(cmd)
                        Delay(10)
                        cmd = String.Format("CURSor1:X2Position {0}", position2)
                        Docommand(cmd)
                    Case "HBArs"
                        cmd = String.Format("CURSor1:Y1Position {0}", position1)
                        Docommand(cmd)
                        Delay(10)
                        cmd = String.Format("CURSor1:Y2Position {0}", position2)
                        Docommand(cmd)
                End Select
            Case 1
                cmd = "CURSOR:" & type & ":POSITION1 " & position1
                Docommand(cmd)
                Delay(10)
                cmd = "CURSOR:" & type & ":POSITION2 " & position2
                Docommand(cmd)
            Case 2
            Case 3

        End Select




    End Function

    'Function RS_Cursor_move(ByVal position1 As Double, ByVal position2 As Double, ByVal position3 As Double, ByVal position4 As Double) As Integer
    '    'Defines the position of the vertical cursor line
    '    'CURSor<m>:X1Position <X1Position>
    '    'CURSor<m>:X2Position <X2Position>
    '    '<X2Position>,<X1Position>:Range 0 to 500

    '    'Defines the position of the horizontal cursor line
    '    'CURSor<m>:Y1Position <Y1Position>
    '    'CURSor<m>:Y2Position <Y2Position>
    '    '<Y2Position>,<Y1Position>:Range -50 to 50

    '    'If set to ON,the horizontal cursor lines follow the waveform.
    '    'CURSor<m>:TRACKing <trackcurve>
    '    '<trackcurve>:ON/OFF


    '    ts = "CURSor1:X1Position" & " " & position1
    '     visa_write(RS_Scope_Dev,RS_vi, ts)
    '    Delay(10)

    '    ts = "CURSor1:X2Position" & " " & position2
    '     visa_write(RS_Scope_Dev,RS_vi, ts)
    '    Delay(10)

    '    ts = "CURSor1:Y1Position" & " " & position3
    '     visa_write(RS_Scope_Dev,RS_vi, ts)
    '    Delay(10)

    '    ts = "CURSor1:Y2Position" & " " & position4
    '     visa_write(RS_Scope_Dev,RS_vi, ts)



    'End Function
    Function RS_Cursor_track(ByVal ONOFF As String) As Integer


        ts = "CURSor1:TRACKing " & ONOFF
        visa_write(RS_Scope_Dev, RS_vi, ts)


    End Function
    Function Cursor_delta(ByVal type As String) As Double
        Dim delta As String = ""
        If RS_Scope = False Then
            'move time: type="VBArs"
            'move volt: type="HBArs"


            ts = "CURSor:" & type & ":DELTa?"

            ilwrt(Scope_Dev, ts, CInt(Len(ts)))
            ilrd(Scope_Dev, ValueStr, ARRAYSIZE)
            If ibcntl > 0 Then
                Cursor_delta = Val(Mid(ValueStr, 1, (ibcntl - 1)))
            End If


        Else
            Select Case type
                Case "VBArs"
                    delta = ":XDELta?"
                Case "HBArs"
                    delta = ":YDELta?"
            End Select


            ts = "CURSor1" & delta
            visa_write(RS_Scope_Dev, RS_vi, ts)
            visa_status = viRead(RS_vi, visa_response, Len(visa_response), retcount)
            Cursor_delta = Val(Mid(visa_response, 1, (retcount - 1)))
        End If

        Return Cursor_delta

    End Function
    'Function RS_Cursor_delta(ByVal type As String) As Double
    '    'CURSor<x>:XDELta?
    '    'CURSor<x>:YDELta?

    '    Dim delta As String = ""

    '    Select Case type
    '        Case "VBArs"
    '            delta = ":XDELta?"
    '        Case "HBArs"
    '            delta = ":YDELta?"
    '    End Select


    '    ts = "CURSor1" & delta
    '     visa_write(RS_Scope_Dev,RS_vi, ts)
    '    visa_status = viRead(RS_vi, visa_response, Len(visa_response), retcount)
    '    RS_Cursor_delta = Val(Mid(visa_response, 1, (retcount - 1)))
    '    Return RS_Cursor_delta


    'End Function

    Function RS_Cursor_Hvalue(ByVal type As String) As Double
        'CURSor1:Y1Position?
        'CURSor1:Y2Position?
        Dim Cursor_delta As Double
        Dim delta As String = ""

        Select Case type
            Case "Y1"
                delta = ":Y1Position?"
            Case "Y2"
                delta = ":Y2Position?"
        End Select


        ts = "CURSor1" & delta
        visa_write(RS_Scope_Dev, RS_vi, ts)
        visa_status = viRead(RS_vi, visa_response, Len(visa_response), retcount)
        Cursor_delta = Val(Mid(visa_response, 1, (retcount - 1)))
        Return Cursor_delta


    End Function
    Function Trigger_set(ByVal source_num As Integer, ByVal edge As String, ByVal level As Double) As Integer
        If RS_Scope = False Then
            'This command sets or queries the type of coupling for the edge trigger.
            'TRIGger:{A|B}:EDGE:COUPling {AC|DC|HFRej|LFRej|NOISErej|ATRIGger}

            'This command sets or queries the slope for the edge trigger.
            'TRIGger:{A|B}:EDGE:SLOpe {RISe|FALL|EITher}

            'This command sets or queries the source for the edge trigger.
            'TRIGger:{A|B}:EDGE:SOUrce {AUXiliary|CH<x>|LINE}

            'This command sets or queries the level for the trigger.
            'TRIGger:{A|B}:LEVel {ECL|TTL|<NR3>}
            '<NR3> specifies the trigger level in user units (usually volts).

            ts = "TRIGger:A:EDGE:COUPling DC"
            ilwrt(Scope_Dev, ts, CInt(Len(ts)))

            If edge = "R" Then
                ts = "TRIGger:A:EDGE:SLOpe RISe"
            ElseIf edge = "F" Then
                ts = "TRIGger:A:EDGE:SLOpe FALL"
            End If
            ilwrt(Scope_Dev, ts, CInt(Len(ts)))

            ts = "TRIGger:A:EDGE:SOUrce " & "CH" & source_num
            ilwrt(Scope_Dev, ts, CInt(Len(ts)))


            ts = "TRIGger:A:LEVel " & level & "E+00"

            ilwrt(Scope_Dev, ts, CInt(Len(ts)))

        Else
            'Selects the trigger type to trigger on analog channels
            'TRIGger<m>:TYPE <Type>
            '<Type>:EDGE/GLITch/WIDTh

            'Selects the source of the trigger signal
            'TRIGger<m>:SOURce <SOURce>
            '<SOURce>:CHAN1/CHAN2/CHAN3/CHAN4

            'Define the edge for the edge trigger event
            'TRIGger<m>:EDGE:SLOPe <SLOPe>
            '<SLOPe>:POSitive/NEGative/EITHer

            'Sets the trigger level for the specofoed event and source
            'TRIGger<m>:LEVel<n> <Level>
            '<Level>:range -10 to 10 ,default unit:V
            '<n>:1>set CH1 ,2>set CH2 ,3>set CH3,4>set CH4


            ts = "TRIGger1:TYPE EDGE"
            visa_write(RS_Scope_Dev, RS_vi, ts)

            If edge = "R" Then
                ts = "TRIGger1:EDGE:SLOPe POSitive"
            ElseIf edge = "F" Then
                ts = "TRIGger1:EDGE:SLOPe NEGative"
            End If
            visa_write(RS_Scope_Dev, RS_vi, ts)

            ts = "TRIGger1:SOURce" & " " & "CHAN" & source_num
            visa_write(RS_Scope_Dev, RS_vi, ts)


            ts = "TRIGger1:LEVel" & source_num & " " & level
            visa_write(RS_Scope_Dev, RS_vi, ts)
        End If


    End Function

    Function RS_trigger_level(ByVal source_num As Integer) As Double

        ts = "TRIGger1:LEVel" & source_num & " "
        visa_write(RS_Scope_Dev, RS_vi, ts)
    End Function




    'Function RS_Trigger_set(ByVal source_num As Integer, ByVal edge As String, ByVal level As Double) As Integer
    '    'Selects the trigger type to trigger on analog channels
    '    'TRIGger<m>:TYPE <Type>
    '    '<Type>:EDGE/GLITch/WIDTh

    '    'Selects the source of the trigger signal
    '    'TRIGger<m>:SOURce <SOURce>
    '    '<SOURce>:CHAN1/CHAN2/CHAN3/CHAN4

    '    'Define the edge for the edge trigger event
    '    'TRIGger<m>:EDGE:SLOPe <SLOPe>
    '    '<SLOPe>:POSitive/NEGative/EITHer

    '    'Sets the trigger level for the specofoed event and source
    '    'TRIGger<m>:LEVel<n> <Level>
    '    '<Level>:range -10 to 10 ,default unit:V
    '    '<n>:1>set CH1 ,2>set CH2 ,3>set CH3,4>set CH4

    '    ts = "TRIGger1:TYPE EDGE"
    '     visa_write(RS_Scope_Dev,RS_vi, ts)

    '    If edge = "R" Then
    '        ts = "TRIGger1:EDGE:SLOPe POSitive"
    '    ElseIf edge = "F" Then
    '        ts = "TRIGger1:EDGE:SLOPe NEGative"
    '    End If
    '     visa_write(RS_Scope_Dev,RS_vi, ts)

    '    ts = "TRIGger1:SOURce" & " " & "CHAN" & source_num
    '     visa_write(RS_Scope_Dev,RS_vi, ts)


    '    ts = "TRIGger1:LEVel" & source_num & " " & level
    '     visa_write(RS_Scope_Dev,RS_vi, ts)


    'End Function
    Function Trigger_auto_level(ByVal source_num As Integer, ByVal edge As String) As Integer



        If RS_Scope = False Then
            ts = "TRIGger:A:EDGE:SOUrce " & "CH" & source_num
            ilwrt(Scope_Dev, ts, CInt(Len(ts)))

            If edge = "R" Then
                ts = "TRIGger:A:EDGE:SLOpe RISe"
            ElseIf edge = "F" Then
                ts = "TRIGger:A:EDGE:SLOpe FALL"
            End If
            ilwrt(Scope_Dev, ts, CInt(Len(ts)))

            '        This command sets the A trigger level automatically to 50% of the range of the
            'minimum and maximum values of the trigger input signal. 


            ts = "TRIGger:A SETLevel"
            ilwrt(Scope_Dev, ts, CInt(Len(ts)))
        Else

            ts = "TRIGger1:SOURce" & " " & "CHAN" & source_num
            visa_write(RS_Scope_Dev, RS_vi, ts)


            If edge = "R" Then
                ts = "TRIGger1:EDGE:SLOPe POSitive"
            ElseIf edge = "F" Then
                ts = "TRIGger1:EDGE:SLOPe NEGative"
            End If
            visa_write(RS_Scope_Dev, RS_vi, ts)

            'Sets the trigger level automatically  (50%)
            ts = "TRIGger1:FINDlevel"
            visa_write(RS_Scope_Dev, RS_vi, ts)
            Delay(200)

        End If


    End Function

    'Function RS_Trigger_auto_level(ByVal source_num As Integer, ByVal edge As String) As Integer
    '    ts = "TRIGger1:SOURce" & " " & "CHAN" & source_num
    '     visa_write(RS_Scope_Dev,RS_vi, ts)


    '    If edge = "R" Then
    '        ts = "TRIGger1:EDGE:SLOPe POSitive"
    '    ElseIf edge = "F" Then
    '        ts = "TRIGger1:EDGE:SLOPe NEGative"
    '    End If
    '     visa_write(RS_Scope_Dev,RS_vi, ts)

    '    'Sets the trigger level automatically  (50%)
    '    ts = "TRIGger1:FINDlevel"
    '     visa_write(RS_Scope_Dev,RS_vi, ts)

    'End Function

    Function Trigger_timeout_set(ByVal source_num As Integer, ByVal HL As String) As Integer
        If RS_Scope = False Then
            'This command sets or queries the polarity for the A or B pulse timeout trigger for the channel.
            'TRIGger:{A|B}:PULse:TIMEOut:POLarity:CH<x> {STAYSHigh|STAYSLow|EITher}

            If HL = "H" Then
                ts = "TRIGger:A:PULse:TIMEOut:POLarity:" & "CH" & source_num & " STAYSHigh"
            Else

                ts = "TRIGger:A:PULse:TIMEOut:POLarity:" & "CH" & source_num & " STAYSLow"

            End If

            ilwrt(Scope_Dev, ts, CInt(Len(ts)))
        Else

        End If


    End Function



    Function Trigger_timeout_init(ByVal source_num As Integer, ByVal time As Double, ByVal unit As String) As Integer

        Dim temp As String


        temp = Format(time, "#0.000")

        Select Case unit
            Case "ns"
                temp = temp & "E-9"
            Case "us"
                temp = temp & "E-6"
            Case "ms"
                temp = temp & "E-3"
            Case "s"
                temp = temp & "E-0"
        End Select

        If RS_Scope = False Then
            'This command sets or queries the type of trigger.
            'TRIGger:A:TYPe {EDGE|LOGIc|PULse|VIDeo|I2C|CAN|SPI|COMMunication|SERIAL|RS232}}

            'This command sets or queries the source for the pulse trigger.
            'TRIGger:{A|B}:PULse:SOUrce CH<x>

            'This command sets or queries the Timeout Trigger qualification.
            'TRIGger:{A|B}:PULse:TIMEOut:QUAlify {OCCurs|LOGIc}

            'This command sets or queries the pulse timeout trigger time (measured in seconds).
            'TRIGger:{A|B}:PULse:TIMEOut:TIMe <NR3>
            '<NR3> argument specifies the timeout period in seconds.



            ts = "TRIGger:A:TYPe PULse"
            ilwrt(Scope_Dev, ts, CInt(Len(ts)))


            ts = "TRIGger:A:PULse:SOUrce " & "CH" & source_num

            ilwrt(Scope_Dev, ts, CInt(Len(ts)))

            ts = "TRIGger:A:PULse:TIMEOut:QUAlify OCCurs"
            ilwrt(Scope_Dev, ts, CInt(Len(ts)))

            ts = "TRIGger:A:PULse:TIMEOut:TIMe " & temp
            ilwrt(Scope_Dev, ts, CInt(Len(ts)))
        Else
            ts = "TRIGger1:TIMeout:TIME " & temp
            visa_write(RS_Scope_Dev, RS_vi, ts)
        End If


    End Function

    Function Trigger_run(ByVal mode As String) As Integer


        If RS_Scope = False Then
            'This command sets or queries the A trigger mode.
            'TRIGger:A:MODe {AUTO|NORMal}

            If mode = "N" Then
                ts = "TRIGger:A:MODe NORMAL"
            ElseIf mode = "A" Then
                ts = "TRIGger:A:MODe AUTO"
            End If
            ilwrt(Scope_Dev, ts, CInt(Len(ts)))
        Else
            'This command sets or queries the A trigger mode.
            'TRIGger<m>:MODE {AUTO|NORMal}

            If mode = "N" Then
                ts = "TRIGger1:MODE NORMal"
            ElseIf mode = "A" Then
                ts = "TRIGger1:MODE AUTO"
            End If
            visa_write(RS_Scope_Dev, RS_vi, ts)
        End If

    End Function

    'Function RS_Trigger_run(ByVal mode As String) As Integer
    '    'This command sets or queries the A trigger mode.
    '    'TRIGger<m>:MODE {AUTO|NORMal}

    '    If mode = "N" Then
    '        ts = "TRIGger1:MODE NORMal"
    '    ElseIf mode = "A" Then
    '        ts = "TRIGger1:MODE AUTO"
    '    End If
    '     visa_write(RS_Scope_Dev,RS_vi, ts)
    'End Function
    Function RUN_set(ByVal mode As String) As Integer


        If RS_Scope = False Then
            'This command sets or queries the acquisition mode of the instrument.
            'ACQuire:MODe{SAMple|PEAKdetect|HIRes|AVErage|ENVelope}
            'SAMple specifies that the displayed data point value is the first sampled value that is taken during the acquisition interval.

            'This command sets or queries whether the instrument continually acquires acquisitions or acquires a single sequence.
            'ACQuire:STOPAfter {RUNSTop|SEQuence}


            ts = "ACQuire:MODE SAMple"
            ilwrt(Scope_Dev, ts, CInt(Len(ts)))


            ts = "ACQuire:STOPAfter " & mode
            ilwrt(Scope_Dev, ts, CInt(Len(ts)))
        Else
            'This command  sets a single sequence.
            'Pressing "RUN Nx SINGLE" on the front panel button is equivalent to sending this command
            '要先設定"TRIGger1:MODE NORMal",才能設定"SINGle"





            If mode = "SEQuence" Then




                'ts = "SINGle"

                ts = "ACQuire:COUNt 1"

            Else
                'ts = "RUN"
                ts = "ACQuire:COUNt MAX"
            End If


            visa_write(RS_Scope_Dev, RS_vi, ts)
            Delay(10)

        End If

    End Function

    'Function RS_RUN_set() As Integer
    '    'This command  sets a single sequence.
    '    'Pressing "RUN Nx SINGLE" on the front panel button is equivalent to sending this command
    '    '要先設定"TRIGger1:MODE NORMal",才能設定"SINGle"

    '    ts = "SINGle"
    '     visa_write(RS_Scope_Dev,RS_vi, ts)

    'End Function

    Function Scope_RUN(ByVal ONOFF As Boolean) As Integer
        Dim temp As String
        If RS_Scope = False Then
            'This command starts or stops acquisitions.
            'ACQuire:STATE {OFF|ON|RUN|STOP|<NR1>}
            'OFF stops acquisitions.
            'STOP stops acquisitions.
            'ON starts acquisitions.
            'RUN starts acquisitions.

            If ONOFF = True Then
                ts = "ACQuire:STATE RUN"
            Else
                ts = "ACQuire:STATE STOP"
            End If


            ilwrt(Scope_Dev, ts, CInt(Len(ts)))
        Else
            'This command starts or stops acquisitions.
            'ACQuire:STATE {OFF|ON|RUN|STOP|<NR1>}
            'OFF stops acquisitions.
            'STOP stops acquisitions.
            'ON starts acquisitions.
            'RUN starts acquisitions.

            If ONOFF = True Then
                'ts = "RUN"
                ts = "RUNSingle" ';*OPC?"
            Else
                ts = "STOP;*OPC?"
            End If


            visa_write(RS_Scope_Dev, RS_vi, ts)
            '  Delay(10)



        End If

        'Delay(10)


    End Function

    'Function RS_Scope_RUN(ByVal ONOFF As Boolean) As Integer
    '    'This command starts or stops acquisitions.
    '    'ACQuire:STATE {OFF|ON|RUN|STOP|<NR1>}
    '    'OFF stops acquisitions.
    '    'STOP stops acquisitions.
    '    'ON starts acquisitions.
    '    'RUN starts acquisitions.

    '    If ONOFF = True Then
    '        ts = "RUN"
    '    Else
    '        ts = "STOP"
    '    End If


    '     visa_write(RS_Scope_Dev,RS_vi, ts)

    'End Function

    Function Scope_status() As String

        Dim status As String = ""
        Dim temp As Integer


        If RS_Scope = False Then

            ts = "ACQuire:STATE?"
            ilwrt(Scope_Dev, ts, CInt(Len(ts)))
            Delay(10)
            ilrd(Scope_Dev, ValueStr, ARRAYSIZE)

            If ibcntl > 0 Then
                temp = Val(Mid(ValueStr, 1, (ibcntl - 1)))
            End If
            If temp = 0 Or temp = 2 Then
                status = "Stopping"
            ElseIf temp = 1 Or temp = 3 Then
                status = "Running"
            End If
        Else

            ts = "ACQuire:CURRent?"
            visa_write(RS_Scope_Dev, RS_vi, ts)
            visa_status = viRead(RS_vi, visa_response, Len(visa_response), retcount)
            If retcount > 0 Then
                temp = Val(Mid(visa_response, 1, retcount - 1))
                If temp = 1 Then
                    status = "Stopping"
                ElseIf temp = 0 Then
                    status = "Running"
                End If
            End If

        End If



        Return status
    End Function

    Function Acquire_num() As Integer


        If RS_Scope = False Then
            'This query-only command returns the number of waveform acquisitions that have occurred since starting acquisition with the ACQuire:STATE RUN command.

            ts = "ACQuire:NUMACq?"
            ilwrt(Scope_Dev, ts, CInt(Len(ts)))
            ilrd(Scope_Dev, ValueStr, ARRAYSIZE)

            If ibcntl > 0 Then
                Acquire_num = Val(Mid(ValueStr, 1, (ibcntl - 1)))
            End If

        Else

        End If

        Return Acquire_num
    End Function

    Function Scope_measure_clear() As Integer
        Dim i As Integer
        If RS_Scope = True Then
            Scope_RUN(False)
            'RS_View()
        End If
        For i = 1 To 8

            If RS_Scope = False Then
                'This command sets or queries whether the specified measurement slot is computed and displayed.
                'MEASUrement:MEAS<x>:STATE {OFF|ON|<NR1>}
                ts = "MEASUrement:MEAS" & i & ":STATE " & "OFF"
                ilwrt(Scope_Dev, ts, CInt(Len(ts)))
            Else
                'This command sets or queries whether the specified measurement slot is computed and displayed.
                'MEASurement<m> {OFF|ON}
                ts = "MEASurement" & i & " " & "OFF"
                visa_write(RS_Scope_Dev, RS_vi, ts)

            End If
            Delay(10)
        Next



    End Function



    Function RS_Scope_measure_status(ByVal num As Integer, ByVal Status_ON As Boolean) As Integer
        'This command sets or queries whether the specified measurement slot is computed and displayed.
        'MEASurement<m> {OFF|ON}

        If Status_ON = True Then
            ts = "MEASurement" & num & " ON"
        Else
            ts = "MEASurement" & num & " OFF "
        End If


        visa_write(RS_Scope_Dev, RS_vi, ts)


        Delay(10)


    End Function


    Function Scope_measure_set(ByVal x As Integer, ByVal source_num As Integer, ByVal type As String) As Integer

        If RS_Scope = False Then
            'MEASUrement:MEAS<x>:TYPe {AMPlitude|AREa|
            'BURst|CARea|CMEan|CRMs|DELay|DISTDUty|
            'EXTINCTDB|EXTINCTPCT|EXTINCTRATIO|EYEHeight|
            'EYEWIdth|FALL|FREQuency|HIGH|HITs|LOW|
            'MAXimum|MEAN|MEDian|MINImum|NCROss|NDUty|
            'NOVershoot|NWIdth|PBASe|PCROss|PCTCROss|PDUty|
            'PEAKHits|PERIod|PHAse|PK2Pk|PKPKJitter|
            'PKPKNoise|POVershoot|PTOP|PWIdth|QFACtor|
            'RISe|RMS|RMSJitter|RMSNoise|SIGMA1|SIGMA2|
            'SIGMA3|SIXSigmajit|SNRatio|STDdev|UNDEFINED| WAVEFORMS}
            'MEASUrement:MEAS<x>:TYPe?

            '<x>, where <x> can be 1, 2, 3, 4, 5, 6, 7, or 8. There must be an active

            ts = "MEASUrement:MEAS" & x & ":SOUrce CH" & source_num
            ilwrt(Scope_Dev, ts, CInt(Len(ts)))


            ts = "MEASUrement:MEAS" & x & ":TYPe " & type
            '    ts = "MEASUrement:IMMed:TYPE " & type
            ilwrt(Scope_Dev, ts, CInt(Len(ts)))


            ts = "MEASUrement:MEAS" & x & ":STATE " & "ON"
            ilwrt(Scope_Dev, ts, CInt(Len(ts)))


        Else
            'Defines the source of the measurement
            'MEASurement<m>:SOURce <signalsource> 
            '<signalsource>: C1W1＞set CH1,C2W1＞set CH2,C3W1>set CH3,C4W1＞set CH4

            'Defines the measurement type of the selected measurement
            'MEASurement<m>:MAIN <Meastype>
            ' <Meastype>:HIGH/LOW/AMPLitude/MAXimum/MINimum/PDELta/MEAN/RMS/
            'STDDecv/POVershoot/NOVershoot/AREA/RTIMe/FTIMe/PPULse/NPULse/PERiod/
            'FREQuency/PDCYcle/NDCYcle/CYCare/CYCMean/CYCRms/CYCStddev/PULCnt/DELay
            'PHASe/BWIDth/PSWitching/NSWitching/PULSetrain/EDGecount/SHT/SHR/DTOTrigger/PROBemeter

            '先設定MEASurement<x> ON再設定MEASurement<x>:STATistics ON才會開啟STATistics 量測mode

            Select Case type
                Case "PK2Pk"
                    type = "PDELta"

                Case "CRMs"
                    type = "CYCRms"

                Case "CMEan"
                    type = "CYCMean"

                Case "RISe"
                    type = "RTIMe"

                Case "FALL"
                    type = "FTIMe"

                Case "PWIdth"
                    type = "PPULse"

                Case "NWIdth"
                    type = "NPULse"

                Case "PDUty"
                    type = "PDCYcle"
                Case "NDUty"
                    type = "NDCYcle"
            End Select




            ts = "MEASurement" & x & ":SOURce " & " " & "C" & source_num & "W1"
            visa_write(RS_Scope_Dev, RS_vi, ts)

            ts = "MEASurement" & x & ":MAIN " & " " & type
            visa_write(RS_Scope_Dev, RS_vi, ts)

            ts = "MEASurement" & x & " " & "ON"
            visa_write(RS_Scope_Dev, RS_vi, ts)

            ts = "MEASurement" & x & ":STATistics" & " " & "ON"
            visa_write(RS_Scope_Dev, RS_vi, ts)


        End If

    End Function

    'Function RS_Scope_measure_set(ByVal x As Integer, ByVal source_num As Integer, ByVal type As String) As Integer
    '    'Defines the source of the measurement
    '    'MEASurement<m>:SOURce <signalsource> 
    '    '<signalsource>: C1W1＞set CH1,C2W1＞set CH2,C3W1>set CH3,C4W1＞set CH4

    '    'Defines the measurement type of the selected measurement
    '    'MEASurement<m>:MAIN <Meastype>
    '    ' <Meastype>:HIGH/LOW/AMPLitude/MAXimum/MINimum/PDELta/MEAN/RMS/
    '    'STDDecv/POVershoot/NOVershoot/AREA/RTIMe/FTIMe/PPULse/NPULse/PERiod/
    '    'FREQuency/PDCYcle/NDCYcle/CYCare/CYCMean/CYCRms/CYCStddev/PULCnt/DELay
    '    'PHASe/BWIDth/PSWitching/NSWitching/PULSetrain/EDGecount/SHT/SHR/DTOTrigger/PROBemeter

    '    '先設定MEASurement<x> ON再設定MEASurement<x>:STATistics ON才會開啟STATistics 量測mode

    '    ts = "MEASurement" & x & ":SOURce " & " " & "C" & source_num & "W1"
    '     visa_write(RS_Scope_Dev,RS_vi, ts)

    '    ts = "MEASurement" & x & ":MAIN " & " " & type
    '     visa_write(RS_Scope_Dev,RS_vi, ts)

    '    ts = "MEASurement" & x & " " & "ON"
    '     visa_write(RS_Scope_Dev,RS_vi, ts)

    '    ts = "MEASurement" & x & ":STATistics" & " " & "ON"
    '     visa_write(RS_Scope_Dev,RS_vi, ts)


    'End Function


    Function scope_measure_list() As String()
        Dim i As Integer
        Dim measure(8) As String
        Dim temp As String
        Dim note() As String


        If RS_Scope = False Then

            'MEASUrement:MEAS<x>? 
            '0=None, 1=OK

            For i = 1 To 8

                ts = "MEASUrement:MEAS" & i & "?"
                ilwrt(Scope_Dev, ts, CInt(Len(ts)))
                ilrd(Scope_Dev, ValueStr, ARRAYSIZE)
                If ibcntl > 0 Then
                    temp = ibcntl - 1
                    note = Split(Mid(ValueStr, 1, temp), ";")



                    If note(0) = "0" Then
                        measure(i) = "None"
                    Else


                        measure(i) = note(1) & " " & note(2) & " " & note(3) & note(4) & " " & note(5) & note(6) & " " & note(7) & note(8)


                    End If


                End If

            Next

        Else
            'viOpenDefaultRM(defaultRM)
            'viOpen(defaultRM, RS_Scope_Dev, VI_NO_LOCK, 2000, vi)

            'MEAS<x>? 
            '等同 MEASurement<m>?
            '0=None, 1=OK

            For i = 1 To 8

                ts = "MEAS" & i & "?"
                visa_write(RS_Scope_Dev, RS_vi, ts)
                'visa_status = viRead(vi, ts, 1024, retcount)
                visa_status = viRead(RS_vi, visa_response, Len(visa_response), retcount)
                'If Mid(ts, 1, retcount - 1) = "0" Then
                If Mid(visa_response, 1, retcount - 1) = "0" Then
                    measure(i) = "None"
                Else

                    ts = "MEAS" & i & ":MAIN?"
                    visa_write(RS_Scope_Dev, RS_vi, ts)
                    'visa_status = viRead(vi, ts, 1024, retcount)
                    visa_status = viRead(RS_vi, visa_response, Len(visa_response), retcount)
                    'temp = Mid(ts, 1, retcount - 1)
                    temp = Mid(visa_response, 1, retcount - 1)
                    measure(i) = temp

                End If

            Next
            'viClose(vi)
            'viClose(defaultRM)

        End If




        Return measure

    End Function

    Function TDS5054_measure_list() As String()
        Dim i As Integer
        Dim measure(8) As String
        Dim temp As String
        Dim note() As String
        Dim note_temp() As String


        'MEASUrement:MEAS<x>? 
        '0=None, 1=OK

        For i = 1 To 8

            ts = "MEASUrement:MEAS" & i & "?"
            ilwrt(Scope_Dev, ts, CInt(Len(ts)))
            ilrd(Scope_Dev, ValueStr, ARRAYSIZE)
            If ibcntl > 0 Then
                temp = ibcntl - 1
                note = Split(Mid(ValueStr, 1, temp), ";")



                If note(0) = "0" Then
                    measure(i) = "None"
                Else

                    If note(1) = "TYPE UNDEFINED" Then
                        measure(i) = "None"
                    Else

                        note_temp = Split(note(1), " ")

                        measure(i) = note_temp(0) & " "


                        note_temp = Split(note(2), " ")

                        measure(i) = measure(i) & note_temp(0) & " "

                        note_temp = Split(note(3), " ")

                        measure(i) = measure(i) & note_temp(0)

                        note_temp = Split(note(4), " ")

                        measure(i) = measure(i) & note_temp(0) & " "


                        note_temp = Split(note(5), " ")

                        measure(i) = measure(i) & note_temp(0)

                        note_temp = Split(note(6), " ")

                        measure(i) = measure(i) & note_temp(0)

                    End If



                End If


            End If

        Next

        ibonl(Scope_Dev, 0)

        Return measure

    End Function




    'Function RS_measure_list() As String()
    '    Dim i As Integer
    '    Dim measure(8) As String
    '    Dim temp As String
    '    'viOpenDefaultRM(defaultRM)
    '    'viOpen(defaultRM, RS_Scope_Dev, VI_NO_LOCK, 2000, vi)

    '    'MEAS<x>? 
    '    '等同 MEASurement<m>?
    '    '0=None, 1=OK

    '    For i = 1 To 8

    '        ts = "MEAS" & i & "?"
    '         visa_write(RS_Scope_Dev,RS_vi, ts)
    '        'visa_status = viRead(vi, ts, 1024, retcount)
    '        visa_status = viRead(RS_vi, visa_response, Len(visa_response), retcount)
    '        'If Mid(ts, 1, retcount - 1) = "0" Then
    '        If Mid(visa_response, 1, retcount - 1) = "0" Then
    '            measure(i) = "None"
    '        Else

    '            ts = "MEAS" & i & ":MAIN?"
    '             visa_write(RS_Scope_Dev,RS_vi, ts)
    '            'visa_status = viRead(vi, ts, 1024, retcount)
    '            visa_status = viRead(RS_vi, visa_response, Len(visa_response), retcount)
    '            'temp = Mid(ts, 1, retcount - 1)
    '            temp = Mid(visa_response, 1, retcount - 1)
    '            measure(i) = temp

    '        End If

    '    Next
    '    'viClose(vi)
    '    'viClose(defaultRM)

    '    Return measure

    'End Function

    Function Scope_measure_reset() As Integer
        Dim i As Integer
        If RS_Scope = False Then

            'This command (no query form) clears existing measurement statistics from memory.
            ts = "MEASUrement:STATIstics:COUNt RESET"

            ilwrt(Scope_Dev, ts, CInt(Len(ts)))
            Delay(100)

            Scope_RUN(True)
        Else
            'RS_View()



            'ts = "MEASurement:STATistics:RESet"
            'visa_write(RS_Scope_Dev, RS_vi, ts)
            For i = 1 To 8
                ts = "MEASurement" & i & ":STATistics:RESet"
                visa_write(RS_Scope_Dev, RS_vi, ts)
                Delay(10)
            Next
            Scope_RUN(True)
            'RS_View()
            RS_Local()




        End If


        Delay(10)

    End Function



    'Function RS_Scope_measure_reset() As Integer
    '    'This command (no query form) clears existing measurement statistics from memory.
    '    'MEASurement<m>:STATistics:RESet
    '    Dim i As Integer

    '    For i = 1 To 8
    '        ts = "MEASurement" & i & ":STATistics:RESet"
    '         visa_write(RS_Scope_Dev,RS_vi, ts)
    '        Delay(10)
    '    Next
    'End Function



    Function Scope_measure_count(ByVal x As Integer) As Integer
        'This query-only command returns the number of values accumulated for this measurement since the last statistical reset.
        'MEASUrement:MEAS<x>:COUNt?
        Dim measure As Integer

        read_error = 0

        If run = False Then
            Exit Function
        End If

        If RS_Scope = False Then
            ts = "MEASUrement:MEAS" & x & ":COUNt?"


            ilwrt(Scope_Dev, ts, CInt(Len(ts)))
            ilrd(Scope_Dev, ValueStr, ARRAYSIZE)

            If ibcntl > 0 Then

                measure = Val(Mid(ValueStr, 1, (ibcntl - 1)))
            Else
                ilrd(Scope_Dev, ValueStr, ARRAYSIZE)
                While (ibcntl = 0) Or (ibsta = EERR)

                    read_error = read_error + 1
                    If read_error = 10 Then
                        Exit While
                    End If
                    Delay(10)
                    ilrd(Scope_Dev, ValueStr, ARRAYSIZE)
                    'measure = Val(Mid(ValueStr, 1, (ibcntl - 1)))
                End While

                If (ibcnt > 0) And (ibsta <> EERR) Then
                    measure = Val(Mid(ValueStr, 1, (ibcntl - 1)))
                End If


            End If
        Else
            'This query-only command returns the number of values accumulated for this measurement since the last statistical reset.
            'MEASurement<x>:RESult:EVTCount?
            '要設定MEASurement<x>:STATistics ON才讀的到EVTCount的值

            'ts = "MEASurement" & x & ":RESult:EVTCount?"

            ts = "MEASurement" & x & ":RESult:WFMCount?"

            visa_write(RS_Scope_Dev, RS_vi, ts)



            visa_status = viRead(RS_vi, visa_response, Len(visa_response), retcount)
            If visa_status = VI_ERROR_CONN_LOST Then
                viOpen(defaultRM, RS_Scope_Dev, VI_NO_LOCK, 2000, RS_vi)
            End If
            If retcount > 0 Then

                measure = Val(Mid(visa_response, 1, retcount - 1))
            Else


                visa_status = viRead(RS_vi, visa_response, Len(visa_response), retcount)
                If visa_status = VI_ERROR_CONN_LOST Then
                    viOpen(defaultRM, RS_Scope_Dev, VI_NO_LOCK, 2000, RS_vi)
                End If
                While retcount = 0
                    System.Windows.Forms.Application.DoEvents()
                    If run = False Then
                        Exit While
                    End If
                    read_error = read_error + 1
                    If read_error = 10 Then
                        Exit While
                    End If
                    Delay(10)
                    visa_status = viRead(RS_vi, visa_response, Len(visa_response), retcount)
                End While

                If retcount > 0 Then
                    measure = Val(Mid(visa_response, 1, retcount - 1))
                End If

            End If







            'RS_View()


        End If



        Return measure


    End Function

    'Function RS_Scope_measure_count(ByVal x As Integer) As Integer
    '    'This query-only command returns the number of values accumulated for this measurement since the last statistical reset.
    '    'MEASurement<x>:RESult:EVTCount?
    '    '要設定MEASurement<x>:STATistics ON才讀的到EVTCount的值
    '    Dim measure As Integer

    '    ts = "MEASurement" & x & ":RESult:EVTCount?"
    '     visa_write(RS_Scope_Dev,RS_vi, ts)
    '    visa_status = viRead(RS_vi, visa_response, Len(visa_response), retcount)
    '    measure = Val(Mid(visa_response, 1, (retcount - 1)))
    '    Return measure

    'End Function


    Function Scope_measure(ByVal x As Integer, ByVal mode As String) As Double
        Dim measure As Double = 0

        read_error = 0

        If run = False Then
            Exit Function
        End If


        If RS_Scope = False Then
            'x=1~8
            'MEASUrement:MEAS<x>:MEAN?
            'MEASUrement:MEAS<x>:MAXimum?
            'MEASUrement:MEAS<x>:MINImum?
            'MEASUrement:MEAS<x>:VALue?
            'MEASUrement:MEAS<x>:STDdev?


            ts = "MEASUrement:MEAS" & x & ":" & mode & "?"
            ilwrt(Scope_Dev, ts, CInt(Len(ts)))

            ilrd(Scope_Dev, ValueStr, ARRAYSIZE)

            If (ibcnt > 0) And (ibsta <> EERR) Then
                measure = Val(Mid(ValueStr, 1, (ibcntl - 1)))
            Else
                ilrd(Scope_Dev, ValueStr, ARRAYSIZE)
                While (ibcntl = 0) Or (ibsta = EERR)

                    read_error = read_error + 1
                    If read_error = 10 Then
                        Exit While
                    End If
                    Delay(10)
                    ilrd(Scope_Dev, ValueStr, ARRAYSIZE)
                    'measure = Val(Mid(ValueStr, 1, (ibcntl - 1)))
                End While

                If (ibcnt > 0) And (ibsta <> EERR) Then
                    measure = Val(Mid(ValueStr, 1, (ibcntl - 1)))
                End If

            End If

        Else
            Select Case mode
                Case Meas_mean
                    mode = RS_Meas_mean
                Case Meas_max
                    mode = RS_Meas_max

                Case Meas_min
                    mode = RS_Meas_min
                Case Scope_Meas
                    mode = RS_Scope_Meas
            End Select
            'MEASurement<m>:RESult[:ACTual]? [<MeasType>]
            'MEASurement<m>:RESult:AVG? [<MeasType>]
            'MEASurement<m>:RESult:EVTCount? [<MeasType>]
            'MEASurement<m>:RESult:NPEak? [<MeasType>]
            'MEASurement<m>:RESult:PPEak? [<MeasType>]
            'MEASurement<m>:RESult:RMS? [<MeasType>]
            'MEASurement<m>:RESult:WFMCount? [<MeasType>]
            'MEASurement<m>:RESult:STDDev? [<MeasType>]
            ' ● [:ACTual]: current measurement result
            '● AVG: average of the long-term measurement results
            '● EVTCount: number of measurement results in the long-term measurement
            '● NPEak: negative peak value of the long-term measurement results
            '● PPEak: positive peak value of the long-term measurement results
            '● RELiability: reliability of the measurement result
            '● RMS: RMS value of the long-term measurement results
            '● STDDev: standard deviation of the long-term measurement results
            'For a detailed description of the results see "Measurement selection: MEASurement<


            ts = "MEASurement" & x & ":RESult:" & mode & "?"
            visa_write(RS_Scope_Dev, RS_vi, ts)


            visa_status = viRead(RS_vi, visa_response, Len(visa_response), retcount)
            If visa_status = VI_ERROR_CONN_LOST Then
                viOpen(defaultRM, RS_Scope_Dev, VI_NO_LOCK, 2000, RS_vi)
            End If

            If retcount > 0 Then

                measure = Val(Mid(visa_response, 1, retcount - 1))
            Else

                visa_status = viRead(RS_vi, visa_response, Len(visa_response), retcount)
                If visa_status = VI_ERROR_CONN_LOST Then
                    viOpen(defaultRM, RS_Scope_Dev, VI_NO_LOCK, 2000, RS_vi)
                End If

                While retcount = 0
                    System.Windows.Forms.Application.DoEvents()


                    read_error = read_error + 1
                    If (read_error = 20) Or (run = False) Then
                        Return measure
                        Exit Function
                    End If
                    Delay(10)
                    visa_status = viRead(RS_vi, visa_response, Len(visa_response), retcount)
                End While

                If retcount > 0 Then
                    measure = Val(Mid(visa_response, 1, retcount - 1))
                End If

            End If

        End If












        Return measure
    End Function


    'Function RS_Scope_measure(ByVal x As Integer, ByVal mode As String) As Double

    '    'MEASurement<m>:RESult[:ACTual]? [<MeasType>]
    '    'MEASurement<m>:RESult:AVG? [<MeasType>]
    '    'MEASurement<m>:RESult:EVTCount? [<MeasType>]
    '    'MEASurement<m>:RESult:NPEak? [<MeasType>]
    '    'MEASurement<m>:RESult:PPEak? [<MeasType>]
    '    'MEASurement<m>:RESult:RMS? [<MeasType>]
    '    'MEASurement<m>:RESult:WFMCount? [<MeasType>]
    '    'MEASurement<m>:RESult:STDDev? [<MeasType>]
    '    ' ● [:ACTual]: current measurement result
    '    '● AVG: average of the long-term measurement results
    '    '● EVTCount: number of measurement results in the long-term measurement
    '    '● NPEak: negative peak value of the long-term measurement results
    '    '● PPEak: positive peak value of the long-term measurement results
    '    '● RELiability: reliability of the measurement result
    '    '● RMS: RMS value of the long-term measurement results
    '    '● STDDev: standard deviation of the long-term measurement results
    '    'For a detailed description of the results see "Measurement selection: MEASurement<


    '    ts = "MEASurement" & x & ":RESult:" & mode & "?"
    '     visa_write(RS_Scope_Dev,RS_vi, ts)
    '    visa_status = viRead(RS_vi, visa_response, Len(visa_response), retcount)


    '    Return Val(Mid(visa_response, 1, retcount - 1))

    'End Function



    Function TDS5054_measure(ByVal x As Integer, ByVal mode As String) As Double
        Dim measure As Double
        Dim temp() As String

        'x=1~8
        'MEASUrement:MEAS<x>:MEAN?
        'MEASUrement:MEAS<x>:MAXimum?
        'MEASUrement:MEAS<x>:MINImum?
        'MEASUrement:MEAS<x>:VALue?
        'MEASUrement:MEAS<x>:STDdev?

        ts = "MEASUrement:MEAS" & x & ":" & mode & "?"
        ilwrt(Scope_Dev, ts, CInt(Len(ts)))
        ilrd(Scope_Dev, ValueStr, ARRAYSIZE)



        If ibcntl > 0 Then
            temp = Split(ValueStr, " ")
            measure = Val(Mid(temp(1), 1, (ibcntl - 1)))

        End If


        Return measure
    End Function

    Function Waveform_data_init(ByVal data_Stop As Integer) As Integer


        ts = "SAVe:WAVEform:FILEFormat SPREADSHEETCsv"
        ilwrt(Scope_Dev, ts, CInt(Len(ts)))
        ts = "SAVe:WAVEform:DATa:STOP " & data_Stop
        ilwrt(Scope_Dev, ts, CInt(Len(ts)))

        ts = "SAVe:WAVEform:FORCESAMEFilesize ON"
        ilwrt(Scope_Dev, ts, CInt(Len(ts)))


    End Function

    Function RS_Waveform_data_init() As Integer


        ts = "EXPort:WAVeform:SCOPe WFM"

        visa_write(RS_Scope_Dev, RS_vi, ts)


        ts = "EXPort:WAVeform:RAW OFF"

        visa_write(RS_Scope_Dev, RS_vi, ts)



        ts = "EXPort:WAVeform:INCXvalues ON"

        visa_write(RS_Scope_Dev, RS_vi, ts)


        'STB
        'OPC

        ts = "EXPort:WAVeform:DLOGging OFF"
        visa_write(RS_Scope_Dev, RS_vi, ts)


    End Function

    Function Waveform_data(ByVal file_path As String, ByVal save_path As String, ByVal channel As String) As Long
        Dim file_temp As String

        Dim ByteSize As Long
        Dim temp() As String

        read_error = 0

        'If My.Computer.FileSystem.FileExists(save_path) = True Then
        '    My.Computer.FileSystem.DeleteFile(save_path)
        'End If


        'Delay(10)

        If RS_Scope = False Then
            'This command specifies or returns the file format for saved waveforms.
            'SAVe:WAVEform:FILEFormat {INTERNal|MATHCad|MATLab|SPREADSHEETCsv|SPREADSHEETTxt|TIMEStamp}

            'This command (no query form) saves a waveform to one of four reference memory locations or a file.
            'SAVe:WAVEform <wfm>,{<file path>|REF<x>}
            '<wfm> is the waveform that will be saved. Valid waveforms include CH<x>, MATH<y>, and REF<x>.

            'This command (no query form) prints a named file to a named port.
            'FILESystem:READFile <filepath>

            '------------------------------------------------------------

            ts = "SAVe:WAVEform:FILEFormat SPREADSHEETCsv"
            ilwrt(Scope_Dev, ts, CInt(Len(ts)))
            'ts = "SAVe:WAVEform:DATa:STOP 1000000"
            'ilwrt(Scope_Dev, ts, CInt(Len(ts)))


            ts = "SAVe:WAVEform:FORCESAMEFilesize ON"
            ilwrt(Scope_Dev, ts, CInt(Len(ts)))
            ts = "SAVE:WAVEFORM CH" & channel & ", " & """" & file_path & """"
            ilwrt(Scope_Dev, ts, CInt(Len(ts)))
            ts = "FILESystem:READFile " & """" & file_path & """"
            ilwrt(Scope_Dev, ts, CInt(Len(ts)))

            ts = save_path
            ibrdf(Scope_Dev, ts)


        Else


            temp = Split(file_path, ".")
            file_temp = temp(0) & ".Wfm." & temp(1)
            'file format要小寫
            '一次產生兩個file，其中xxx.Wfm.csv才是我們要的
            'FW Ver4.7會將Time與volt分開不同列，之前的是由";"分開


            ts = "EXPort:WAVeform:SOURce C" & channel & "W1"
            visa_write(RS_Scope_Dev, RS_vi, ts)

            'ts = "EXPort:WAVeform:SCOPe WFM"
            'visa_status = viWrite(RS_vi, ts, Len(ts), retcount)

            ts = "EXPort:WAVeform:NAME '" & file_path & "'"
            visa_write(RS_Scope_Dev, RS_vi, ts)

            'export x,y data
            'ts = "EXPort:WAVeform:INCXvalues ON"
            'visa_write(RS_Scope_Dev, RS_vi, ts)

            ts = "EXPort:WAVeform:SAVE"
            visa_write(RS_Scope_Dev, RS_vi, ts)


            'xxx.Wfm.csv儲存的是時間跟電壓資料
            ts = "MMEM:DATA? '" & file_temp & "'"
            visa_write(RS_Scope_Dev, RS_vi, ts)

            visa_status = viRead(RS_vi, ts, 2, retcount)

            While retcount = 0

                System.Windows.Forms.Application.DoEvents()


                read_error = read_error + 1
                If (read_error = 200) Or (run = False) Then
                    Return 0
                    Exit Function
                End If
                Delay(10)
                visa_status = viRead(RS_vi, ts, 2, retcount)
            End While

            If retcount > 0 Then
                visa_status = viRead(RS_vi, ts, Mid(ts, 2, 1), retcount)


                visa_status = viReadToFile(RS_vi, save_path, Val(ts), retcount)
            End If


        End If




        If read_error = 200 Then
            ByteSize = 0
            If My.Computer.FileSystem.FileExists(save_path) = True Then
                My.Computer.FileSystem.DeleteFile(save_path)
            End If
        Else

            ByteSize = FileLen(save_path)
        End If




        Return ByteSize

    End Function

    'Function RS_Waveform_data(ByVal file_path As String, ByVal save_path As String, ByVal x As Integer) As Integer
    '    Dim file_temp As Integer
    '    ts = "EXPort:WAVeform:SOURce C" & x & "W1"
    '     visa_write(RS_Scope_Dev,RS_vi, ts)
    '    ts = "EXPort:WAVeform:SCOPe WFM"
    '     visa_write(RS_Scope_Dev,RS_vi, ts)

    '    ts = "EXPort:WAVeform:NAME 'C:\Temp\" & file_path & "'"
    '     visa_write(RS_Scope_Dev,RS_vi, ts)

    '    ts = "EXPort:WAVeform:SAVE"
    '     visa_write(RS_Scope_Dev,RS_vi, ts)

    '    ts = "MMEM:DATA? 'C:\Temp\" & file_path & "'"
    '     visa_write(RS_Scope_Dev,RS_vi, ts)

    '    visa_status = viRead(RS_vi, ts, 2, retcount)

    '    file_temp = Mid(ts, 2, 1)

    '    visa_status = viRead(RS_vi, ts, file_temp, retcount)

    '    visa_status = viReadToFile(Scope_Dev, save_path, Val(ts), retcount)

    'End Function

    Function RS_Waveform_save(ByVal file_path As String, ByVal save_path As String) As Integer
        Dim file_temp As Integer
        ts = "SYSTem:DISPlay:UPDate ON"
        visa_write(RS_Scope_Dev, RS_vi, ts)

        ts = "HCOP:DEST 'MMEM'"
        visa_write(RS_Scope_Dev, RS_vi, ts)

        ts = "HCOPy:DEV:INV OFF"
        visa_write(RS_Scope_Dev, RS_vi, ts)

        ts = "HCOP:DEV:LANG PNG"
        visa_write(RS_Scope_Dev, RS_vi, ts)

        ts = "MMEM:NAME '" & Scope_folder & "\" & file_path & "'"
        visa_write(RS_Scope_Dev, RS_vi, ts)

        ts = "HCOP:DEV:COL ON"
        visa_write(RS_Scope_Dev, RS_vi, ts)

        ts = "HCOP:IMM;*WAI"
        visa_write(RS_Scope_Dev, RS_vi, ts)

        ts = "MMEM:DATA? '" & Scope_folder & "\" & file_path & "'"

        visa_write(RS_Scope_Dev, RS_vi, ts)

        visa_status = viRead(RS_vi, ts, 2, retcount)

        file_temp = Mid(ts, 2, 1)

        visa_status = viRead(RS_vi, ts, file_temp, retcount)

        visa_status = viReadToFile(Scope_Dev, save_path, Val(ts), retcount)

    End Function



    Function Max_Min(ByVal Num() As Double, ByVal sample As Integer) As Double()
        Dim Jitter() As Double = {32768, -32768}
        'Jitter(0)=Min; Jitter(1)=Max

        For i = 0 To sample
            If Num(i) < Jitter(0) Then
                Jitter(0) = Num(i)
            End If
            If Num(i) > Jitter(1) Then
                Jitter(1) = Num(i)
            End If
        Next

        Return Jitter

    End Function

    Function RS_Hardcopy_init(ByVal scope_format As String) As Integer



        ts = "HCOPy:DEV:INV OFF"
        visa_write(RS_Scope_Dev, RS_vi, ts)

        ts = "HCOP:DEV:COL ON"
        visa_write(RS_Scope_Dev, RS_vi, ts)


        ts = "HCOP:DEV:LANG " & scope_format
        visa_write(RS_Scope_Dev, RS_vi, ts)



    End Function



    Function Hardcopy(ByVal Scope_format As String, ByVal pc_path As String) As Long
        Dim ByteSize As Long
        Dim file_temp As Integer

        read_error = 0



        If RS_Scope = False Then

            ts = "EXP:FORM " & Scope_format
            ilwrt(Scope_Dev, ts, CInt(Len(ts)))
            ts = "HARDCopy:PORT FILE"
            ilwrt(Scope_Dev, ts, CInt(Len(ts)))

            ts = "HARDCopy:FILEName '" & Scope_folder & "\hardcopy." & Scope_format & "'"
            ilwrt(Scope_Dev, ts, CInt(Len(ts)))

            ts = "HARDCopy STARt"
            ilwrt(Scope_Dev, ts, CInt(Len(ts)))

            ts = "FILESystem:READFile '" & Scope_folder & "\hardcopy." & Scope_format & "'"
            ilwrt(Scope_Dev, ts, CInt(Len(ts)))




            '-------------------------------------------------
            'FILESystem:READFile Error

            While (ibcntl = 0) Or (ibsta = EERR)
                System.Windows.Forms.Application.DoEvents()

                read_error = read_error + 1
                If (read_error = 100) Or (run = False) Then
                    Return 0
                    Exit Function
                End If
                Delay(10)
                ilwrt(Scope_Dev, ts, CInt(Len(ts)))

            End While

            '-------------------------------------------------


            'file path of PC
            read_error = 0

            ts = pc_path
            ibrdf32(Scope_Dev, ts)

            '-------------------------------------------------
            ' Error
            While (ibcntl = 0) Or (ibsta = EERR)
                System.Windows.Forms.Application.DoEvents()

                read_error = read_error + 1
                If (read_error = 100) Or (run = False) Then
                    Return 0
                    Exit Function
                End If
                Delay(10)

                ibrdf32(Scope_Dev, ts)
            End While
            '-------------------------------------------------



            ByteSize = FileLen(pc_path)
        Else
            Scope_RUN(False)
            Delay(10)
            ts = "SYSTem:DISPlay:UPDate ON"
            visa_write(RS_Scope_Dev, RS_vi, ts)



            ts = "HCOP:DEST 'MMEM'"
            visa_write(RS_Scope_Dev, RS_vi, ts)

            'ts = "HCOPy:DEV:INV OFF"
            'visa_status = viWrite(RS_vi, ts, Len(ts), retcount)

            'ts = "HCOP:DEV:COL ON"
            'visa_status = viWrite(RS_vi, ts, Len(ts), retcount)


            'ts = "HCOP:DEV:LANG " & Scope_format
            'visa_status = viWrite(RS_vi, ts, Len(ts), retcount)



            ts = "MMEM:NAME '" & Scope_folder & "\hardcopy." & Scope_format & "'"
            visa_write(RS_Scope_Dev, RS_vi, ts)



            'ts = "HCOP:IMM;*WAI"
            ts = "HCOP:IMMediate;*OPC?"
            visa_write(RS_Scope_Dev, RS_vi, ts)
            visa_status = viRead(RS_vi, visa_response, Len(visa_response), retcount)


            While retcount = 0

                System.Windows.Forms.Application.DoEvents()

                read_error = read_error + 1
                If (read_error = 100) Or (run = False) Then
                    Return 0
                    Exit Function
                End If
                Delay(10)

                visa_status = viRead(RS_vi, visa_response, Len(visa_response), retcount)
            End While
            Delay(10)

            ts = "MMEM:DATA? '" & Scope_folder & "\hardcopy." & Scope_format & "'"
            visa_write(RS_Scope_Dev, RS_vi, ts)


            visa_status = viRead(RS_vi, ts, 2, retcount)


            '-------------------------------------------------
            'FILESystem:READFile Error

            While retcount = 0
                System.Windows.Forms.Application.DoEvents()

                read_error = read_error + 1
                If (read_error = 100) Or (run = False) Then
                    Return 0
                    Exit Function
                End If
                Delay(10)
                visa_status = viRead(RS_vi, ts, 2, retcount)
            End While




            '-------------------------------------------------

            file_temp = Mid(ts, 2, 1)

            visa_status = viRead(RS_vi, ts, file_temp, retcount)

            read_error = 0


            '-------------------------------------------------
            'FILESystem:READFile Error

            While retcount = 0
                System.Windows.Forms.Application.DoEvents()

                read_error = read_error + 1
                If (read_error = 100) Or (run = False) Then
                    Return 0
                    Exit Function
                End If
                Delay(10)
                visa_status = viRead(RS_vi, ts, file_temp, retcount)
            End While




            '-------------------------------------------------



            visa_status = viReadToFile(RS_vi, pc_path, Val(ts), retcount)




            '-------------------------------------------------
            'FILESystem:READFile Error
            read_error = 0
            While retcount = 0
                System.Windows.Forms.Application.DoEvents()

                read_error = read_error + 1
                If (read_error = 100) Or (run = False) Then
                    Return 0
                    Exit Function
                End If
                Delay(10)
                visa_status = viReadToFile(RS_vi, pc_path, Val(ts), retcount)
            End While




            '-------------------------------------------------


            ByteSize = 0
            If (System.IO.File.Exists(pc_path)) = True Then
                ByteSize = FileLen(pc_path)

            End If

        End If





        Return ByteSize


    End Function

    'Function RS_Hardcopy(ByVal Scope_format As String, ByVal pc_path As String) As Integer
    '    Dim ByteSize As Long
    '    Dim file_temp As Integer
    '    ts = "SYSTem:DISPlay:UPDate ON"
    '    visa_status = viWrite(vi, ts, Len(ts), retcount)

    '    ts = "HCOP:DEST 'MMEM'"
    '    visa_status = viWrite(vi, ts, Len(ts), retcount)

    '    ts = "HCOPy:DEV:INV OFF"
    '    visa_status = viWrite(vi, ts, Len(ts), retcount)



    '    ts = "HCOP:DEV:LANG PNG"
    '    visa_status = viWrite(vi, ts, Len(ts), retcount)



    '    ts = "MMEM:NAME 'C:\hardcopy." & Scope_format & "'"
    '    visa_status = viWrite(vi, ts, Len(ts), retcount)
    '    ts = "HCOP:DEV:COL ON"
    '    visa_status = viWrite(vi, ts, Len(ts), retcount)
    '    ts = "HCOP:IMM;*WAI"
    '    visa_status = viWrite(vi, ts, Len(ts), retcount)

    '    ts = "MMEM:DATA? 'C:\hardcopy." & Scope_format & "'"
    '    visa_status = viWrite(vi, ts, Len(ts), retcount)


    '    visa_status = viRead(vi, ts, 2, retcount)

    '    file_temp = Mid(ts, 2, 1)

    '    visa_status = viRead(vi, ts, file_temp, retcount)


    '    visa_status = viReadToFile(vi, pc_path, Val(ts), retcount)
    '    ts = "&GTL"
    '    visa_status = viWrite(vi, ts, Len(ts), retcount)

    '    ByteSize = 0
    '    If (System.IO.File.Exists(pc_path)) = True Then
    '        ByteSize = FileLen(pc_path)

    '    End If

    '    Return ByteSize

    'End Function


    Function TDSJIT3() As Integer


        ts = "application:activate ""Jitter Analysis - Advanced"""
        ilwrt(Scope_Dev, ts, CInt(Len(ts)))
    End Function






    Function RS_Display(ByVal result_mode As String, ByVal parameters As String) As Integer
        'PREV | FLOA | DOCK

        ts = "DISPlay:RESultboxes:" & result_mode & " " & parameters

        visa_write(RS_Scope_Dev, RS_vi, ts)

    End Function




End Module
