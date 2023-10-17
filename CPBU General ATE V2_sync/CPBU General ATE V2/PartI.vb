Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Runtime.InteropServices.Marshal
Imports System.Diagnostics
Public Class PartI

    ' Dim vcc_enable As Boolean = False
    '--------------------------------------
    'Excel Setting
    '--------------------------------------
    Dim vin_max, vin_min As Double
    Dim test_temp() As String
    Dim test_vcc() As String
    Dim test_fs() As String
    Dim test_vout() As String
    Dim test_ton() As String
    Dim test_vin() As String
    Dim test_fs0() As String
    Dim test_IOB_start() As String
    Dim test_IOB_stop() As String
    Dim test_Fs_Max() As String
    Dim test_Fs_Min() As String
    Dim test_FCC() As Boolean
    Dim total_iout() As Double
    Dim other_iout() As Double
    Dim total_other_iout As Integer


    Dim PartI_file As String

    Dim PartI_first As Boolean = True

    Dim Stability As String = "Stability"
    Dim Line_Regulation As String = "Line Regulation"
    Dim Load_Regulation As String = "Load Regulation"
    Dim Efficiency As String = "Efficiency"
    Dim Jitter As String = "Jitter"


    'Statbility
    Dim Beta_folder As String
    Dim Error_folder As String
    Dim Jitter_folder As String
    Dim import_now As Boolean = False
    Dim Freq_Chart As Excel.ChartObject
    Dim Ton_Chart As Excel.ChartObject
    Dim Toff_Chart As Excel.ChartObject
    Dim Vpp_Chart As Excel.ChartObject

    Dim error_pic_col, error_pic_row As Integer
    'Dim beta_pic_col, beta_pic_row As Integer
    Dim vout_scale_now As Integer

    '--------------------------------------
    'Jitter
    Dim jitter_pic_col(), jitter_pic_row() As Integer
    Dim jitter_pic_path As String = Environment.CurrentDirectory & "\HARDCOPY.PNG"
    '--------------------------------------
    'Efficiency

    Dim Eff_Chart As Excel.ChartObject
    Dim Iin_change As Boolean = False
    Dim eff_iin_change() As Double

    '--------------------------------------
    'Load Regulation
    Dim LoadR_Chart As Excel.ChartObject
    '--------------------------------------
    'Line Regulation
    Dim LineR_Chart As Excel.ChartObject
    '--------------------------------------
    'Parameter
    Dim ton_value As Double
    ' Dim vin_max, vin_min As Double
    Dim vcc_max, vcc_min As Double
    '--------------------------------------
    'Test
    Dim Full_load As Double = 0
    '--------------------------------------
    'Stability
    Dim jitter_col() As String = {Vout_name, Iout_name, "Ton_mean(ns)", "Toff_min(ns)", "Toff_max(ns)", "Tjitter(ns)", "Dmax", "Dmin", "Dave", "Jitter %", "PASS/FAIL"}
    Dim stability_col() As String = {Vout_name, Iout_name, "Max. Criteria(kHz)", "Min. Criteria(kHz)", "Frequency(kHz)", "Frequency(mean)", "Frequency(min)", "Frequency(max)", "Freq_update(kHz)",
                               "Ton(ns)", "Ton(mean)", "Ton(min)", "Ton(max)", "Ton_update(ns)",
                               "Toff(ns)", "Toff(mean)", "Toff(min)", "Toff(max)", "Toff_update(ns)",
                               "Vpp(mV)", "Vpp(mean)", "Vpp(min)", "Vpp(max)",
                               "Vmax(max)", "Vmin(min)",
                               "PASS/FAIL", "Error"}
    Dim lineR_col() As String = {Vout_name, "Frequency(kHz)", "Frequency(mean)", "Frequency(min)", "Frequency(max)",
                             "Ton(ns)", "Ton(mean)", "Ton(min)", "Ton(max)",
                             "Toff(ns)", "Toff(mean)", "Toff(min)", "Toff(max)"}

    Dim stability_row_start() As Integer '紀錄stability的iout數
    Dim stability_row_stop() As Integer '紀錄stability的iout數
    Dim stability_report_row() As Integer
    Dim data_set_now As Integer
    Dim Fs_Max As Double
    Dim Fs_Min As Double
    Dim Fs_leak_0A As Double
    Dim ton_now As Double
    Dim IOUT_Boundary_Array() As Double
    Dim x As Integer = 1
    Dim H_scale_value As Double
    Dim AutoScalling_EN As Boolean = False
    Dim VoutScalling_CCM As Boolean = False
    Dim ton_pass As Double
    Dim toff_pass As Double
    Dim fs(3) As Double
    Dim fs_update As Double
    Dim ton(3) As Double
    Dim toff(3) As Double
    Dim vpp(5) As Double
    Dim autoscanning_update As Boolean = False
    Dim error_pic_num As Integer
    Dim beta_pic_num As Integer
    Dim Jitter_pic_num As Integer
    Dim Fs_CCM As Boolean = False
    Dim cursor_state As Boolean = False
    Dim IOUT_Boundary_Start As Double
    Dim IOUT_Boundary_Stop As Double
    '--------------------------------------
    'Instrument
    Dim Eff_vout_daq As String
    Dim iin_min_range As String
    Dim iout_min_range As String
    '--------------------------------------
    Dim hyperlink_col, hyperlink_row As Integer
    Dim start_test_time As Date
    Dim LR_Vin_test_num As Integer
    Dim VCC_test_num As Integer
    Dim fs_test_num As Integer
    Dim Vout_test_num As Integer
    Dim Vin_test_num As Integer
    Dim eff_iout_num As Integer
    Dim lineR_iout_num As Integer
    Dim jitter_iout_num As Integer
    Dim stability_iout_num As Integer
    'Dim data_set_num() As Integer
    Dim test_point_num As Integer
    '----------------------------------------------------------------------------------------------
    Dim iin_row As Integer = 10
    '----------------------------------------------------------------------------------------------
    'Efficiency(%)

    Dim eff_title_total As Integer

    Dim scope_meas_col() As String = {"(value)", "(mean)", "(min)", "(max)"}

    Dim total_Eff As Boolean = False


    '------------------------------------------------------------------------------------------------
    'Test
    Function initial() As Integer


        'Chamber
        If Main.check_TA_en.Checked = True Then
            txt_TA.Text = "TA"
        Else
            txt_TA.Text = "START"
        End If

        '-----------------------------------------------------
        'I2C INIT
        If Main.data_i2c.Rows.Count = 0 Then
            pic_i2C_init.Visible = True
            txt_I2C_init.Visible = False
        Else
            pic_i2C_init.Visible = False
            txt_I2C_init.Visible = True
        End If
        '-----------------------------------------------------
        'EN SET
        Main.fs_vout_set()
        If (Main.check_en.Checked = True) Then

            pic_EN.Visible = False
            txt_EN.Visible = True
        Else

            pic_EN.Visible = True
            txt_EN.Visible = False
        End If
        '-----------------------------------------------------
        'Fs SET


        clist_fs.Items.AddRange(fs_value)
        clist_fs.SetItemChecked(0, True)


        If clist_fs.Items.Count = 1 Then

            clist_fs.Enabled = False
        Else
            clist_fs.Enabled = True
        End If



        If Main.cbox_fs_ctr.SelectedItem = no_device Then

            pic_Fs_set.Visible = True
            txt_Fs_set.Visible = False
        Else
            pic_Fs_set.Visible = False
            txt_Fs_set.Visible = True
        End If
        '-----------------------------------------------------
        'Vout SET
        clist_vout.Items.AddRange(vout_value)
        clist_vout.SetItemChecked(0, True)
        If clist_vout.Items.Count = 1 Then
            clist_vout.Enabled = False
        Else
            clist_vout.Enabled = True
        End If

        If Main.cbox_vout_ctr.SelectedItem = no_device Then
            pic_vout_set.Visible = True
            txt_vout_set.Visible = False
        Else
            pic_vout_set.Visible = False
            txt_vout_set.Visible = True
        End If


        Full_load = Main.num_full_load.Value
        num_iout_auto_stop.Maximum = Full_load * 1000
        data_jitter_iout.Rows.Add(Math.Round(Full_load * 0.5, 4).ToString("F2"))
        data_jitter_iout.Rows.Add(Math.Round(Full_load * 0.75, 4).ToString("F2"))
        data_jitter_iout.Rows.Add(Math.Round(Full_load, 4).ToString("F2"))
        data_lineR_iout.Rows.Add("0.0000")
        data_lineR_iout.Rows.Add(Math.Round(Full_load * 0.5, 4).ToString("F4"))
        data_lineR_iout.Rows.Add(Math.Round(Full_load, 4).ToString("F4"))
        '------------------------------------------------------------------
        'Vin
        cbox_vin.Items.Clear()
        cbox_vin2.Items.Clear()
        cbox_VCC.Items.Clear()
        If Power_num > 0 Then
            cbox_vin.Items.AddRange(Power_name)
            cbox_vin2.Items.AddRange(Power_name)
            cbox_VCC.Items.AddRange(Power_name)
        Else
            cbox_vin.Items.Add(no_device)
            cbox_vin2.Items.Add(no_device)
        End If
        cbox_VCC.Items.Add(no_device)
        cbox_vin.SelectedIndex = 0
        cbox_vin2.SelectedIndex = 0
        cbox_VCC.SelectedItem = no_device
        '-----------------------------------------------------
        cbox_IIN_meter.Items.Clear()
        cbox_Iout_meter.Items.Clear()
        cbox_IIN_meter2.Items.Clear()
        cbox_Iout_meter2.Items.Clear()
        cbox_Icc_meter.Items.Clear()
        If Meter_num > 0 Then
            rbtn_meter_iin.Checked = True
            rbtn_meter_iout.Checked = True

            rbtn_meter_iin2.Checked = True
            rbtn_meter_iout2.Checked = True

            cbox_IIN_meter.Items.AddRange(Meter_name)
            cbox_Iout_meter.Items.AddRange(Meter_name)

            cbox_IIN_meter2.Items.AddRange(Meter_name)
            cbox_Iout_meter2.Items.AddRange(Meter_name)

            cbox_Icc_meter.Items.AddRange(Meter_name)
        ElseIf Main.data_meas.Rows.Count > 0 Then
            rbtn_board_iin.Checked = True
            rbtn_board_iout.Checked = True

            rbtn_board_iin2.Checked = True
            rbtn_board_iout2.Checked = True
        Else
            rbtn_Iin_PW.Checked = True
            rbtn_iout_load.Checked = True

            rbtn_Iin_PW2.Checked = True
            rbtn_iout_load2.Checked = True
        End If
        cbox_IIN_meter.Items.Add(no_device)
        cbox_Iout_meter.Items.Add(no_device)
        cbox_IIN_meter2.Items.Add(no_device)
        cbox_Iout_meter2.Items.Add(no_device)
        cbox_Icc_meter.Items.Add(no_device)
        cbox_IIN_meter.SelectedIndex = 0
        cbox_IIN_relay.SelectedIndex = 0
        cbox_IIN_meter2.SelectedIndex = 0
        cbox_IIN_relay2.SelectedIndex = 0
        If Meter_num > 1 Then
            cbox_Iout_meter.SelectedIndex = 1
            cbox_Iout_meter2.SelectedIndex = 1
        Else
            cbox_Iout_meter.SelectedIndex = 0
            cbox_Iout_meter2.SelectedIndex = 0
        End If
        cbox_Iout_relay.SelectedIndex = 1
        cbox_Iout_relay2.SelectedIndex = 1
        If Meter_num > 2 Then
            cbox_Icc_meter.SelectedIndex = 2
        Else
            cbox_Icc_meter.SelectedIndex = cbox_Icc_meter.Items.Count - 1
        End If
        If RTBB_board = False Then
            check_iin.Checked = False
            check_iout.Checked = False

            check_iin2.Checked = False
            check_iout2.Checked = False
        Else
            check_iin.Checked = True
            check_iout.Checked = True

            check_iin2.Checked = True
            check_iout2.Checked = True
        End If
        Panel_model2.Enabled = True
        If DCLOAD_63600 = True Then
            txt_load_model1.Text = LOAD_63600_Model(0)
            txt_load_model12.Text = LOAD_63600_Model(0)
            If LOAD_63600_Model.Length = 2 Then
                txt_load_model2.Text = LOAD_63600_Model(1)
                txt_load_model22.Text = LOAD_63600_Model(1)
            Else
                check_IOUT_ch1.Checked = True
                check_IOUT_ch3.Checked = False
                check_IOUT_ch4.Checked = False
                Panel_model2.Enabled = False


                check_IOUT_ch12.Checked = True
                check_IOUT_ch32.Checked = False
                check_IOUT_ch42.Checked = False
                Panel_model2.Enabled = False
            End If
        ElseIf Mid(Load_device, 1, 4) = "6312" Then
            txt_load_model1.Text = LOAD_6312_Model
            txt_load_model2.Text = LOAD_6312_Model

            txt_load_model12.Text = LOAD_6312_Model
            txt_load_model22.Text = LOAD_6312_Model
        End If
    End Function

    Function scope_init_set() As Integer
        '-------------------------------------------------------------------
        'Scope Set
        'Display_persistence(False)
        'Fs
        cbox_channel_lx.SelectedIndex = 0
        cbox_coupling_lx.SelectedIndex = 0 'DC 1M
        cbox_BW_lx.SelectedItem = "Full"

        cbox_channel_lx2.SelectedIndex = 0
        cbox_coupling_lx2.SelectedIndex = 0 'DC 1M
        cbox_BW_lx2.SelectedItem = "Full"

        'VOUT
        cbox_channel_vout.SelectedIndex = 1
        cbox_coupling_vout.SelectedIndex = 0 'DC 1M
        cbox_BW_vout.SelectedItem = "20MHz"

        cbox_channel_vout2.SelectedIndex = 1
        cbox_coupling_vout2.SelectedIndex = 0 'DC 1M
        cbox_BW_vout2.SelectedItem = "20MHz"


        'VIN
        cbox_channel_vin.SelectedIndex = 2
        cbox_coupling_vin.SelectedIndex = 1  'AC
        cbox_BW_vin.SelectedItem = "20MHz"
        'IOUT
        cbox_channel_iout.SelectedIndex = 3
        cbox_coupling_iout.SelectedIndex = 2 'DC 50Ohm
        cbox_BW_iout.SelectedItem = "20MHz"
        '-------------------------------------------------------------------
        'Test setup
    End Function

    Function reflesh() As Integer
        Dim fs_change As Boolean = False
        Dim vout_change As Boolean = False
        Dim i As Integer
        'Chamber
        If Main.check_TA_en.Checked = True Then
            txt_TA.Text = "TA"
        Else
            txt_TA.Text = "START"
        End If
        '-----------------------------------------------------
        'I2C INIT
        If Main.data_i2c.Rows.Count = 0 Then
            pic_i2C_init.Visible = True
            txt_I2C_init.Visible = False
        Else
            pic_i2C_init.Visible = False
            txt_I2C_init.Visible = True
        End If
        '-----------------------------------------------------
        'EN SET
        If (Main.check_en.Checked = True) Then
            pic_EN.Visible = False
            txt_EN.Visible = True
        Else
            pic_EN.Visible = True
            txt_EN.Visible = False
        End If
        Main.fs_vout_set()
        For i = 0 To clist_fs.Items.Count - 1
            If clist_fs.Items(i) <> fs_value(i) Then
                fs_change = True
                Exit For
            End If
        Next
        If fs_change = True Then
            clist_fs.Items.Clear()
            clist_fs.Items.AddRange(fs_value)
            clist_fs.SetItemChecked(0, True)
            If clist_fs.Items.Count = 1 Then
                clist_fs.Enabled = False
            Else
                clist_fs.Enabled = True
            End If
            If Main.cbox_fs_ctr.SelectedItem = no_device Then
                pic_Fs_set.Visible = True
                txt_Fs_set.Visible = False
            Else
                pic_Fs_set.Visible = False
                txt_Fs_set.Visible = True
            End If
        End If
        For i = 0 To clist_vout.Items.Count - 1
            If clist_vout.Items(i) <> vout_value(i) Then
                vout_change = True
                Exit For
            End If
        Next
        If vout_change = True Then
            clist_vout.Items.Clear()
            clist_vout.Items.AddRange(vout_value)
            clist_vout.SetItemChecked(0, True)
            If clist_vout.Items.Count = 1 Then
                clist_vout.Enabled = False
            Else
                clist_vout.Enabled = True
            End If
            If Main.cbox_vout_ctr.SelectedItem = no_device Then
                pic_vout_set.Visible = True
                txt_vout_set.Visible = False
            Else
                pic_vout_set.Visible = False
                txt_vout_set.Visible = True
            End If

        End If
        '-----------------------------------------------------
        If Full_load <> Main.num_full_load.Value Then
            Full_load = Main.num_full_load.Value
            num_iout_auto_stop.Maximum = Full_load * 1000
            data_jitter_iout.Rows.Clear()
            data_jitter_iout.Rows.Add(Math.Round(Full_load * 0.5, 4))
            data_jitter_iout.Rows.Add(Math.Round(Full_load * 0.75, 4))
            data_jitter_iout.Rows.Add(Math.Round(Full_load, 4))
            data_lineR_iout.Rows.Clear()
            data_lineR_iout.Rows.Add("0")
            data_lineR_iout.Rows.Add(Math.Round(Full_load * 0.5, 4))
            data_lineR_iout.Rows.Add(Math.Round(Full_load, 4))
        End If
        ''-----------------------------------------------------
        ''Vin
        If device_select_same(cbox_vin, txt_vin_addr, Power, False) = False Then
            vin_dev_ch = 0
        End If

        If device_select_same(cbox_vin2, txt_vin_addr2, Power, False) = False Then
            vin_dev_ch2 = 0
        End If
        cbox_vin_ch.SelectedIndex = vin_dev_ch
        cbox_vin_ch2.SelectedIndex = vin_dev_ch2


        If device_select_same(cbox_VCC, txt_vcc_Addr, Power, True) = False Then
            vcc_dev_ch = 0
        End If
        cbox_VCC_ch.SelectedIndex = vcc_dev_ch
        If Meter_num = 0 Then
            If Main.data_meas.Rows.Count > 0 Then
                rbtn_board_iin.Checked = True
                rbtn_board_iout.Checked = True

                rbtn_board_iin2.Checked = True
                rbtn_board_iout2.Checked = True
            Else
                rbtn_Iin_PW.Checked = True
                rbtn_iout_load.Checked = True

                rbtn_Iin_PW2.Checked = True
                rbtn_iout_load2.Checked = True
            End If
        Else
            device_select_same(cbox_IIN_meter, txt_IIN_addr, Meter, True)
            device_select_same(cbox_Iout_meter, txt_Iout_addr, Meter, True)
            device_select_same(cbox_Icc_meter, txt_Icc_addr, Meter, True)


            device_select_same(cbox_IIN_meter2, txt_IIN_addr2, Meter, True)
            device_select_same(cbox_Iout_meter2, txt_Iout_addr2, Meter, True)
        End If


        Panel_model2.Enabled = True

        If DCLOAD_63600 = True Then

            txt_load_model1.Text = LOAD_63600_Model(0)
            txt_load_model12.Text = LOAD_63600_Model(0)
            If LOAD_63600_Model.Length = 2 Then
                txt_load_model2.Text = LOAD_63600_Model(1)
                txt_load_model22.Text = LOAD_63600_Model(1)
            Else
                check_IOUT_ch1.Checked = True
                check_IOUT_ch3.Checked = False
                check_IOUT_ch4.Checked = False
                Panel_model2.Enabled = False

                check_IOUT_ch12.Checked = True
                check_IOUT_ch32.Checked = False
                check_IOUT_ch42.Checked = False

            End If
        ElseIf Mid(Load_device, 1, 4) = "6312" Then
            txt_load_model1.Text = LOAD_6312_Model
            txt_load_model2.Text = LOAD_6312_Model

            txt_load_model12.Text = LOAD_6312_Model
            txt_load_model22.Text = LOAD_6312_Model
        End If

        INA226_Iin_max_L = 0.08 / Main.num_IIN_Rshunt_L.Value

        If rbtn_board_iin.Checked = True Then
            num_iin_change.Maximum = INA226_Iin_max_L * 1000
            num_iin_change2.Maximum = INA226_Iin_max_L * 1000
        End If


        If rbtn_board_iout.Checked = True Then
            num_iout_change.Maximum = INA226_Iout_max_L * 1000
            num_iout_change2.Maximum = INA226_Iout_max_L * 1000
        End If
    End Function


    Function Sense_vin() As String
        vin_power_sense(cbox_vin.SelectedItem, num_vin_sense.Value, num_vin_max.Value, vin_now)
    End Function

    Function scope_time_init() As Integer
        H_Samplerate(Samplerate_num, "MS/s")
        H_position(num_location.Value) '左邊第1格
        H_reclength(RL_value)
        H_Roll("OFF")
        H_scale(H_scale_value, "ns") '1/Fs_Min(Hz)*n/10 
    End Function

    Function scope_measure_init() As Integer
        Dim meas_ch As Integer
        Dim meas_type As String
        Dim meas_freq_type As String = txt_meas1.Text
        Dim meas_pwidth_type As String = txt_meas2.Text
        Dim meas_nwidth_type As String = txt_meas3.Text
        Dim meas_vpp_type As String = txt_meas4.Text
        Dim meas_vmax_type As String = txt_meas5.Text
        Dim meas_vmin_type As String = txt_meas6.Text
        Scope_measure_clear()
        Delay(100)


        If DUT2_en Then
            ' freq
            meas_ch = Val(Mid(txt_meas1_ch.Text, 3))
            Scope_measure_set(meas1, meas_ch, meas_freq_type)
            ' Pwidth
            meas_ch = Val(Mid(txt_meas2_ch.Text, 3))
            Scope_measure_set(meas2, meas_ch, meas_pwidth_type)
            ' NWidth
            meas_ch = Val(Mid(txt_meas3_ch.Text, 3))
            Scope_measure_set(meas3, meas_ch, meas_nwidth_type)
            ' PK2PK
            meas_ch = Val(Mid(txt_meas4_ch.Text, 3))
            Scope_measure_set(meas4, meas_ch, meas_vpp_type)


            ' dut2
            ' freq
            meas_ch = Val(Mid(txt_meas1_ch.Text, 5))
            Scope_measure_set(meas5, meas_ch, meas_freq_type)
            ' Pwidth
            meas_ch = Val(Mid(txt_meas2_ch.Text, 5))
            Scope_measure_set(meas6, meas_ch, meas_pwidth_type)
            ' NWidth
            meas_ch = Val(Mid(txt_meas3_ch.Text, 5))
            Scope_measure_set(meas7, meas_ch, meas_nwidth_type)
            ' PK2PK
            meas_ch = Val(Mid(txt_meas4_ch.Text, 5))
            Scope_measure_set(meas8, meas_ch, meas_vpp_type)
        Else
            meas_ch = Val(Mid(txt_meas1_ch.Text, 3))
            Scope_measure_set(meas1, meas_ch, meas_freq_type)

            meas_ch = Val(Mid(txt_meas2_ch.Text, 3))
            Scope_measure_set(meas2, meas_ch, meas_pwidth_type)

            meas_ch = Val(Mid(txt_meas3_ch.Text, 3))
            Scope_measure_set(meas3, meas_ch, meas_nwidth_type)

            meas_ch = Val(Mid(txt_meas4_ch.Text, 3))
            Scope_measure_set(meas4, meas_ch, meas_vpp_type)

            meas_ch = Val(Mid(txt_meas5_ch.Text, 3))
            Scope_measure_set(meas5, meas_ch, meas_vmax_type)

            meas_ch = Val(Mid(txt_meas6_ch.Text, 3))
            Scope_measure_set(meas6, meas_ch, meas_vmin_type)
        End If
        If RS_Scope = True Then
            RS_Local()
            'If DUT2_en Then : RS_View(True) : End If
        End If
    End Function

    Sub instrument_2nd_init()
        Dim temp As String

        DAQ_resolution2 = cbox_data_resolution2.SelectedItem
        vin_daq2 = Mid(cbox_vin_daq2.SelectedItem, 3)
        vout_daq2 = Mid(cbox_vout_daq2.SelectedItem, 3)
        Eff_vout_daq2 = Mid(cbox_vout1_daq2.SelectedItem, 3)

        DAQ_config(vin_daq2)
        DAQ_config(vout_daq2)
        DAQ_config(Eff_vout_daq2)


        If check_Efficiency.Checked = True Then
            If rbtn_meter_iin2.Checked = True Then
                Meter_iin_addr2 = Val(txt_IIN_addr2.Text)
                Meter_iin_dev2 = ildev(BDINDEX, Meter_iin_addr2, NO_SECONDARY_ADDR, TIMEOUT, EOTMODE, EOSMODE)
                If check_iin2.Checked = True Then
                    ' need check Iin meter function 
                    Iin_Meter_initial(check_iin2, cbox_IIN_meter2, cbox_IIN_relay2)
                End If
            Else
                INA226_Iin_initial(True) 'High Range
            End If
        End If

        If (rbtn_meter_iout2.Checked = True) And (cbox_Iout_meter2.SelectedItem <> no_device) Then
            Meter_iout_addr2 = Val(txt_Iout_addr2.Text)
            Meter_iout_dev2 = ildev(BDINDEX, Meter_iout_addr2, NO_SECONDARY_ADDR, TIMEOUT, EOTMODE, EOSMODE)
            If check_iout2.Checked = True Then
                Iout_Meter_initial(check_iout2, cbox_Iout_meter2, cbox_Iout_relay2)
            End If
        ElseIf rbtn_board_iout2.Checked = True Then
            If iout_now > INA226_Iout_max_L Then
                Iout_Meter_Max = True
            Else
                Iout_Meter_Max = False
            End If
        End If

#Region "----------------- power supply2 setting -----------------"
        temp = data_result.Rows(0).Cells("col_test_vin1").Value
        vin_now = temp
        vin_addr2 = Val(txt_vin_addr2.Text)
        vin_Dev2 = ildev(BDINDEX, vin_addr, NO_SECONDARY_ADDR, TIMEOUT, EOTMODE, EOSMODE)
        Power_Dev2 = vin_Dev2
        vin_device2 = cbox_vin2.SelectedItem
        Vin_out2 = Power_channel(vin_device2, cbox_vin_ch2.SelectedIndex)
        power_volt(vin_device2, Vin_out2, vin_now)
        If num_VIN_OCP2.Value <> 0 Then
            power_OCP_init(vin_device2, Vin_out2, num_VIN_OCP2.Value)
            If vin_device2 = "E3632A" Then
                E3632_OCP = num_VIN_OCP2.Value
            End If
        Else
            If vin_device = "E3632A" Then
                E3632_OCP = 7
            End If
        End If
        power_on_off(vin_device2, Vin_out2, "ON")
        vin_meas2 = DAQ_read(vin_daq2)
        If (check_vin_sense2.Checked = True) And (vin_meas2 < (vin_now * 0.5)) Then
            error_message("Please confirm the VIN DAQ channel setting is correct!")
        End If
#End Region

        monitor_vout2 = check_shutdown2.Checked
        If monitor_vout2 = True Then
            'VOUT
            vout_daq2 = Mid(cbox_vout_daq2.SelectedItem, 3)
            DAQ_config(vout_daq2)
            check_vout2()
        End If

        ' Scope setting
        lx2_ch = Val(Mid(cbox_channel_lx2.SelectedItem, 3))
        vout2_ch = Val(Mid(cbox_channel_vout2.SelectedItem, 3))

        CHx_display(lx2_ch, "ON")
        CHx_coupling(lx2_ch, cbox_coupling_lx2.SelectedItem)
        CHx_position(lx2_ch, num_position_lx2.Value)
        CHx_label(lx2_ch, txt_scope_lx2.Text)
        If cbox_coupling_lx2.SelectedItem <> "AC" Then
            'DC
            CHx_OFFSET(lx2_ch, num_offset_lx2.Value)
        Else
            'AC
            CHx_OFFSET(lx2_ch, 0)
        End If
        CHx_Bandwidth(lx2_ch, cbox_BW_lx2.SelectedItem)
        If rbtn_manual_lx.Checked = True Then
            CHx_scale(lx2_ch, num_scale_lx.Value, "mV") 'Voltage Scale > SW/2
        Else
            CHx_scale(lx2_ch, (vin_now / num_lx_scale.Value), "V") 'Voltage Scale > SW/2
        End If

        CHx_display(vout2_ch, "ON")
        CHx_coupling(vout2_ch, cbox_coupling_vout2.SelectedItem)
        CHx_position(vout2_ch, num_position_vout2.Value)
        CHx_label(vout2_ch, txt_scope_vout2.Text)
        If cbox_coupling_vout2.SelectedItem <> "AC" Then
            'DC
            CHx_OFFSET(vout2_ch, vout_now)
        Else
            CHx_OFFSET(vout2_ch, 0)
        End If
        CHx_Bandwidth(vout2_ch, cbox_BW_vout2.SelectedItem)
    End Sub

    Function instrument_init() As Integer
        Dim temp As String
        Dim i As Integer

        Power_EN(False)
        ''----------------------------------------------------------------------------------
        DAQ_resolution = cbox_data_resolution.SelectedItem

        ''DAQ
        'VIN
        vin_daq = Mid(cbox_vin_daq.SelectedItem, 3)
        DAQ_config(vin_daq)

        'VOUT
        vout_daq = Mid(cbox_vout_daq.SelectedItem, 3)
        DAQ_config(vout_daq)

        'Efficiency
        Eff_vout_daq = Mid(cbox_vout1_daq.SelectedItem, 3)
        DAQ_config(Eff_vout_daq)

        'Vcc
        If cbox_VCC_daq.SelectedItem <> no_device Then
            vcc_daq = Mid(cbox_VCC_daq.SelectedItem, 3)
            DAQ_config(vcc_daq)
        End If
        ''----------------------------------------------------------------------------------
        'DC Load 
        Load_Dev = ildev(BDINDEX, Load_addr, NO_SECONDARY_ADDR, TIMEOUT, EOTMODE, EOSMODE)
        Iout_board_EN = rbtn_board_iout.Checked
        Iout_board_EN2 = rbtn_board_iout2.Checked
        DCload_ch(0) = check_IOUT_ch1.Checked
        DCload_ch(1) = check_IOUT_ch2.Checked
        DCload_ch(2) = check_IOUT_ch3.Checked
        DCload_ch(3) = check_IOUT_ch4.Checked

        DCload_ch2(0) = check_IOUT_ch12.Checked
        DCload_ch2(1) = check_IOUT_ch22.Checked
        DCload_ch2(2) = check_IOUT_ch32.Checked
        DCload_ch2(3) = check_IOUT_ch42.Checked

        iout_now = data_result.Rows(0).Cells("col_test_iout1").Value
        load_num = 0
        For i = 0 To 3
            If DCload_ch(i) = True Then
                ReDim Preserve Load_ch_set(load_num)
                'Load_ch = i + 1
                Load_ch_set(load_num) = i + 1
                Load_ch = Load_ch_set(load_num)
                Load_range = Load_range_L


                load_init(Load_range)
                load_set(0)
                DCLoad_ONOFF("OFF")
                load_num = load_num + 1
            End If
        Next

        If DUT2_en Then
            load_num = 0
            For i = 0 To 3
                If DCload_ch2(i) = True Then
                    ReDim Preserve Load_ch_set2(load_num)
                    'Load_ch = i + 1
                    Load_ch_set2(load_num) = i + 1
                    Load_ch2 = Load_ch_set2(load_num)
                    Load_range = Load_range_L


                    load_init(Load_range, DUT2_en)
                    load_set(0, DUT2_en)
                    DCLoad_ONOFF("OFF", DUT2_en)
                    load_num = load_num + 1
                End If
            Next
        End If




        ''----------------------------------------------------------------------------------
        'Meter
        If check_Efficiency.Checked = True Then
            If rbtn_meter_iin.Checked = True Then
                Meter_iin_addr = Val(txt_IIN_addr.Text)
                Meter_iin_dev = ildev(BDINDEX, Meter_iin_addr, NO_SECONDARY_ADDR, TIMEOUT, EOTMODE, EOSMODE)
                If check_iin.Checked = True Then
                    Iin_Meter_initial(check_iin, cbox_IIN_meter, cbox_IIN_relay)
                End If
            Else
                INA226_Iin_initial(True) 'High Range
            End If

            If DUT2_en Then
                If rbtn_meter_iin2.Checked = True Then
                    Meter_iin_addr2 = Val(txt_IIN_addr2.Text)
                    Meter_iin_dev2 = ildev(BDINDEX, Meter_iin_addr2, NO_SECONDARY_ADDR, TIMEOUT, EOTMODE, EOSMODE)
                    If check_iin2.Checked = True Then
                        Iin_Meter_initial(check_iin2, cbox_IIN_meter2, cbox_IIN_relay2)
                    End If
                Else
                    INA226_Iin_initial(True) 'High Range
                End If
            End If
        End If


        'Meter set High
        If (rbtn_meter_iout.Checked = True) And (cbox_Iout_meter.SelectedItem <> no_device) Then
            Meter_iout_addr = Val(txt_Iout_addr.Text)
            Meter_iout_dev = ildev(BDINDEX, Meter_iout_addr, NO_SECONDARY_ADDR, TIMEOUT, EOTMODE, EOSMODE)
            If check_iout.Checked = True Then
                Iout_Meter_initial(check_iout, cbox_Iout_meter, cbox_Iout_relay)
            End If
        ElseIf rbtn_board_iout.Checked = True Then
            If iout_now > INA226_Iout_max_L Then
                Iout_Meter_Max = True
            Else
                Iout_Meter_Max = False
            End If
        End If

        If DUT2_en Then
            If (rbtn_meter_iout2.Checked = True) And (cbox_Iout_meter2.SelectedItem <> no_device) Then
                Meter_iout_addr2 = Val(txt_Iout_addr2.Text)
                Meter_iout_dev2 = ildev(BDINDEX, Meter_iout_addr2, NO_SECONDARY_ADDR, TIMEOUT, EOTMODE, EOSMODE)
                If check_iout2.Checked = True Then
                    Iout_Meter_initial(check_iout2, cbox_Iout_meter2, cbox_Iout_relay2)
                End If
            ElseIf rbtn_board_iout2.Checked = True Then
                If iout_now > INA226_Iout_max_L Then
                    Iout_Meter_Max = True
                Else
                    Iout_Meter_Max = False
                End If
            End If
        End If

        If cbox_Icc_meter.SelectedItem <> no_device Then
            Meter_icc_addr = Val(txt_Icc_addr.Text)
            Meter_icc_dev = ildev(BDINDEX, Meter_icc_addr, NO_SECONDARY_ADDR, TIMEOUT, EOTMODE, EOSMODE)
        End If

        ''----------------------------------------------------------------------------------
        'Vcc
        temp = data_result.Rows(0).Cells("col_test_vcc1").Value
        If (cbox_VCC.SelectedItem <> no_device) And (temp <> "") Then
            vcc_addr = Val(txt_vcc_Addr.Text)
            VCC_Dev = ildev(BDINDEX, vcc_addr, NO_SECONDARY_ADDR, TIMEOUT, EOTMODE, EOSMODE)
            VCC_device = cbox_VCC.SelectedItem
            VCC_out = Power_channel(VCC_device, cbox_VCC_ch.SelectedIndex)
            vcc_now = temp
            Power_Dev = VCC_Dev
            power_volt(VCC_device, VCC_out, vcc_now)
            power_on_off(VCC_device, VCC_out, "ON")
        End If


        '----------------------------------------------------------------------------------
        'Vin
        temp = data_result.Rows(0).Cells("col_test_vin1").Value
        vin_now = temp
        vin_addr = Val(txt_vin_addr.Text)
        vin_Dev = ildev(BDINDEX, vin_addr, NO_SECONDARY_ADDR, TIMEOUT, EOTMODE, EOSMODE)
        Power_Dev = vin_Dev
        vin_device = cbox_vin.SelectedItem
        Vin_out = Power_channel(vin_device, cbox_vin_ch.SelectedIndex)
        power_volt(vin_device, Vin_out, vin_now)
        If num_VIN_OCP.Value <> 0 Then
            power_OCP_init(vin_device, Vin_out, num_VIN_OCP.Value)
            If vin_device = "E3632A" Then
                E3632_OCP = num_VIN_OCP.Value
            End If
        Else
            If vin_device = "E3632A" Then
                E3632_OCP = 7
            End If
        End If

        power_on_off(vin_device, Vin_out, "ON")
        vin_meas = DAQ_read(vin_daq)
        If (check_vin_sense.Checked = True) And (vin_meas < (vin_now * 0.5)) Then
            error_message("Please confirm the VIN DAQ channel setting is correct!")
        End If

        If DUT2_en Then
            temp = data_result.Rows(0).Cells("col_test_vin1").Value
            vin_now = temp
            vin_addr2 = Val(txt_vin_addr2.Text)
            vin_Dev2 = ildev(BDINDEX, vin_addr2, NO_SECONDARY_ADDR, TIMEOUT, EOTMODE, EOSMODE)
            Power_Dev2 = vin_Dev2
            vin_device2 = cbox_vin2.SelectedItem
            Vin_out2 = Power_channel(vin_device2, cbox_vin_ch2.SelectedIndex)
            power_volt(vin_device2, Vin_out2, vin_now)
            If num_VIN_OCP2.Value <> 0 Then
                power_OCP_init(vin_device2, Vin_out2, num_VIN_OCP2.Value)
                If vin_device2 = "E3632A" Then
                    E3632_OCP = num_VIN_OCP2.Value
                End If
            Else
                If vin_device2 = "E3632A" Then
                    E3632_OCP = 7
                End If
            End If
        End If

        If DUT2_en Then
            power_on_off(vin_device2, Vin_out2, "ON", DUT2_en)
            vin_meas2 = DAQ_read(vin_daq2)
            If (check_vin_sense2.Checked = True) And (vin_meas < (vin_now * 0.5)) Then
                error_message("Please confirm the VIN DAQ channel setting is correct!")
            End If
        End If


        ' ''----------------------------------------------------------------------------------
        'Power Enable
        'Enable
        Power_EN(True)
        ''---------------------------------------------------------------------------------
        'I2C Init
        If Main.data_i2c.Rows.Count > 0 Then
            For i = 0 To Main.data_i2c.Rows.Count - 1
                System.Windows.Forms.Application.DoEvents()
                If run = False Then
                    Exit For
                End If
                reg_write(Val("&H" & Main.data_i2c.Rows(i).Cells(0).Value), Val("&H" & Main.data_i2c.Rows(i).Cells(1).Value), Val("&H" & Main.data_i2c.Rows(i).Cells(2).Value))
            Next
        End If
        ''---------------------------------------------------------------------------------

        If DUT2_en Then
            If Main.data_i2c.Rows.Count > 0 Then
                For i = 0 To Main.data_i2c.Rows.Count - 1
                    System.Windows.Forms.Application.DoEvents()
                    If run = False Then
                        Exit For
                    End If
                    reg_write(Val("&H" & Main.data_i2c.Rows(i).Cells(0).Value), Val("&H" & Main.data_i2c.Rows(i).Cells(1).Value), Val("&H" & Main.data_i2c.Rows(i).Cells(2).Value), DUT2_en)
                Next
            End If
        End If

        'Fs Set
        temp = data_result.Rows(0).Cells("col_test_fsw1").Value
        fs_now = temp * 1000
        If Main.cbox_fs_ctr.SelectedItem <> no_device Then
            Grobal_Control(Fs_control, fs_now)
        End If

        If DUT2_en Then
            If Main.cbox_fs_ctr.SelectedItem <> no_device Then
                Grobal_Control(Fs_control, fs_now, DUT2_en)
            End If
        End If

        'Vout Set
        temp = data_result.Rows(0).Cells("col_test_vout1").Value
        vout_now = temp
        If Main.cbox_vout_ctr.SelectedItem <> no_device Then
            Grobal_Control(Vout_control, vout_now)
        End If

        If DUT2_en Then
            If Main.cbox_vout_ctr.SelectedItem <> no_device Then
                Grobal_Control(Vout_control, vout_now, DUT2_en)
            End If
        End If

        If Main.check_EN_off.Checked = True Then
            Power_EN(False)
            Power_EN(True)
        End If

        ''---------------------------------------------------------------------------------

        'Check Vout
        monitor_vout = check_shutdown.Checked
        If monitor_vout = True Then
            'VOUT
            vout_daq = Mid(cbox_vout_daq.SelectedItem, 3)
            DAQ_config(vout_daq)
            check_vout()
        End If

        If DUT2_en Then
            monitor_vout2 = check_shutdown2.Checked
            If monitor_vout2 = True Then
                'VOUT
                vout_daq2 = Mid(cbox_vout_daq2.SelectedItem, 3)
                DAQ_config(vout_daq2)
                check_vout2()
            End If
        End If


        '---------------------------------------------------------------------------------
        'Scope Init
        Relay1_BUCK1_VIN = False 'CH1
        Relay2_Islammer_SMBalert = False 'CH1
        Relay3_CH1_Other = False 'CH1
        Relay4_VCC_BUCK2 = False 'CH2
        Relay5_VIN_SCL = False 'CH2
        Relay6_VEN_Ctrl = False 'CH2
        Relay7_Islammer_VSS = False 'CH4
        Relay8_PG_MODE = False 'CH4

        If check_stability.Checked = True Or check_jitter.Checked = True Or (check_LineR.Checked = True And check_lineR_scope.Checked = True) Then

            If (Scope_Addr <> 0) And (RS_Scope = False) Then
                Scope_Dev = ibdev32(BDINDEX, Scope_Addr, NO_SECONDARY_ADDR, TIMEOUT, EOTMODE, EOSMODE)
                RUN_set("RUNSTop")
            ElseIf RS_Scope = True Then
                If RS_Scope_EN = False Then
                    RS_visa(True)
                End If
            End If
            Scope_RUN(False)
            '---------------------------------------------------------------------------------
            If check_scope_vin.Checked = True Then
                vin_ch = Val(Mid(cbox_channel_vin.SelectedItem, 3))
                'Check Board
                If txt_board_VOUT.Text = "Buck1" Then
                    Relay5_VIN_SCL = True
                Else
                    Relay1_BUCK1_VIN = True
                End If
            Else
                vin_ch = 0
            End If

            '---------------------------------------------------------------------------------
            lx_ch = Val(Mid(cbox_channel_lx.SelectedItem, 3))
            '---------------------------------------------------------------------------------
            vout_ch = Val(Mid(cbox_channel_vout.SelectedItem, 3))

            If DUT2_en Then
                lx2_ch = Val(Mid(cbox_channel_lx2.SelectedItem, 3))
                vout2_ch = Val(Mid(cbox_channel_vout2.SelectedItem, 3))
            End If

            'Check Board
            If txt_board_VOUT.Text = "Buck1" Then
                Relay1_BUCK1_VIN = True
            ElseIf txt_board_VOUT.Text = "Buck2" Then
                Relay4_VCC_BUCK2 = True
            End If
            '---------------------------------------------------------------------------------

            If (txt_board_VOUT.Text <> "") And (Main.data_relay.Rows.Count > 0) Then
                relay_Scope_set()
            End If

            '---------------------------------------------------------------------------------
            If check_scope_iout.Checked = True Then
                iout_ch = Val(Mid(cbox_channel_iout.SelectedItem, 3))
            Else
                iout_ch = 0
            End If

            '---------------------------------------------------------------------------------
            For i = 1 To 4
                If (vin_ch <> i) And (lx_ch <> i) And (vout_ch <> i) And (iout_ch <> i) Then
                    CHx_display(i, "OFF")
                End If
            Next

            '---------------------------------------------------------------------------------
            RL_value = num_RL.Value * 1000
            Wave_num = num_wave.Value
            Samplerate_num = num_points.Value

            If RS_Scope = True Then
                Display_persistence(False)
                RS_Display(RS_RES_CURSOR, RS_DISP_PREV)
                'RS_Display(RS_RES_MES, RS_DISP_DOCK)
                RS_Display(RS_RES_MES, RS_DISP_PREV)
                RS_Hardcopy_init("PNG")
                RS_Waveform_data_init()
            End If

            '----------------------------------------------------------
            'Cursors

            If check_cursors.Checked = True Then
                Cursor_set("VBArs", lx_ch, lx_ch)
                Cursor_ONOFF("OFF")
            End If

            '----------------------------------------------------------
            'Measurement setup:
            scope_measure_init()
            ''----------------------------------------------------------------------------------
            'Scope Set

            'LX
            CHx_display(lx_ch, "ON")
            CHx_coupling(lx_ch, cbox_coupling_lx.SelectedItem)
            CHx_position(lx_ch, num_position_lx.Value)
            CHx_label(lx_ch, txt_scope_lx.Text)
            If cbox_coupling_lx.SelectedItem <> "AC" Then
                'DC
                CHx_OFFSET(lx_ch, num_offset_lx.Value)
            Else
                'AC
                CHx_OFFSET(lx_ch, 0)
            End If

            CHx_Bandwidth(lx_ch, cbox_BW_lx.SelectedItem)
            If rbtn_manual_lx.Checked = True Then
                CHx_scale(lx_ch, num_scale_lx.Value, "mV") 'Voltage Scale > SW/2
            Else
                CHx_scale(lx_ch, (vin_now / num_lx_scale.Value), "V") 'Voltage Scale > SW/2
            End If

            ''----------------------------------------------------------------------------------
            'VOUT
            CHx_display(vout_ch, "ON")
            CHx_coupling(vout_ch, cbox_coupling_vout.SelectedItem)
            CHx_position(vout_ch, num_position_vout.Value)
            CHx_label(vout_ch, txt_scope_vout.Text)
            If cbox_coupling_vout.SelectedItem <> "AC" Then
                'DC
                CHx_OFFSET(vout_ch, vout_now)
            Else
                CHx_OFFSET(vout_ch, 0)
            End If
            CHx_Bandwidth(vout_ch, cbox_BW_vout.SelectedItem)

            '--------------------------------------------------------------------------------
            'Vin
            If check_scope_vin.Checked = True Then
                'VIN

                CHx_display(vin_ch, "ON")
                CHx_coupling(vin_ch, cbox_coupling_vin.SelectedItem)
                CHx_position(vin_ch, num_position_vin.Value)
                CHx_label(vin_ch, txt_scope_vin.Text)
                If cbox_coupling_vin.SelectedItem <> "AC" Then
                    CHx_OFFSET(vin_ch, num_offset_vin.Value)
                Else
                    CHx_OFFSET(vin_ch, 0)
                End If

                CHx_Bandwidth(vin_ch, cbox_BW_vin.SelectedItem)

                If (num_offset_vin.Value > 10) And (num_vin_scale.Value < 1000) Then

                    CHx_scale(vin_ch, 1, "V") 'Voltage Scale > 200mV
                Else

                    CHx_scale(vin_ch, num_vin_scale.Value, "mV") 'Voltage Scale > 200mV
                End If

            Else
                ' CHx_display(vin_ch, "OFF")

            End If

            ''----------------------------------------------------------------------------------

            If check_scope_iout.Checked = True Then
                'IOUT
                CHx_display(iout_ch, "ON")
                CHx_coupling(iout_ch, cbox_coupling_iout.SelectedItem)
                CHx_position(iout_ch, num_position_iout.Value)
                CHx_label(iout_ch, txt_scope_iout.Text)
                If cbox_coupling_iout.SelectedItem <> "AC" Then
                    CHx_OFFSET(iout_ch, num_offset_iout.Value)
                Else
                    CHx_OFFSET(iout_ch, 0)
                End If
                CHx_Bandwidth(iout_ch, cbox_BW_iout.SelectedItem)
                iout_scale_now = 200
                CHx_scale(iout_ch, iout_scale_now, "mV") 'a. IOUT < 600mA, Scale = 200mA, b. 600mA<IOUT < 3A, Sacle = 1A,c. 3A <IOUT< 6A, Scale = 2A
            Else
                ' CHx_display(iout_ch, "OFF")
            End If

            ''----------------------------------------------------------
            'Timing Scale

            Fs_Min = data_result.Rows(0).Cells("col_Fs_min").Value
            Fs_Min = Fs_Min * (10 ^ 3)
            H_scale_value = ((1 / Fs_Min) * Wave_num / 10) * (10 ^ 9)

            'Timing Scale
            H_scale(H_scale_value, "ns") '1/Fs_Min(Hz)*n/10 

            'See the time scale formula

            scope_time_init()
            '----------------------------------------------------------
            '
            If rbtn_vin_trigger.Checked = True Then
                Trigger_set(lx_ch, "R", vin_now / num_vin_trigger.Value)
            Else
                Trigger_auto_level(lx_ch, "R")
            End If
            'R&S Scope需要先偵測到才能在設定Auto 
            ' Trigger_auto_level(lx_ch, "R")
            Trigger_run("N")
            RUN_set("RUNSTop")
            'Scope_RUN(True)
        End If


        If DUT2_en Then
            instrument_2nd_init()
        End If

    End Function

    Function instrument_closed() As Integer

        'Meter  
        'High Range
        If check_Efficiency.Checked = True Then
            If (rbtn_meter_iin.Checked = True) And (check_iin.Checked = True) Then
                Iin_Meter_initial(check_iin, cbox_IIN_meter, cbox_IIN_relay)
            Else
                INA226_Iin_initial(True) 'High Range
            End If
        End If


        'Meter set High

        If (rbtn_meter_iout.Checked = True) And (check_iout.Checked = True) Then


            Iout_Meter_initial(check_iout, cbox_Iout_meter, cbox_Iout_relay)

        ElseIf rbtn_board_iout.Checked = True Then

            If iout_now > INA226_Iout_max_L Then
                Iout_Meter_Max = True
            Else
                Iout_Meter_Max = False
            End If


        End If



        ''----------------------------------------------------------------------------------

        'DC Load
        If Load_Dev <> 0 Then

            DCLoad_ONOFF("OFF")
            ibonl(Load_Dev, 0)
            Load_Dev = 0
        End If

        'Meter
        If Meter_iin_dev <> 0 Then


            ibonl(Meter_iin_dev, 0)
            Meter_iin_dev = 0
        End If


        If Meter_iout_dev <> 0 Then

            ibonl(Meter_iout_dev, 0)
            Meter_iout_dev = 0
        End If


        If Meter_icc_dev <> 0 Then

            ibonl(Meter_icc_dev, 0)
            Meter_icc_dev = 0
        End If


        '--------------------------------------------------
        'Power 
        If VCC_Dev <> 0 Then

            ibonl(VCC_Dev, 0)
            VCC_Dev = 0

        End If

        If vin_Dev <> 0 Then

            ibonl(vin_Dev, 0)
            vin_Dev = 0

        End If

        If Ven_Dev <> 0 Then

            ibonl(Ven_Dev, 0)
            Ven_Dev = 0

        End If

        '-------------------------------------------------
        'Scope
        If Scope_Dev <> 0 Then

            ibonl(Scope_Dev, 0)
            Scope_Dev = 0
        End If

        If (RS_Scope = True) And (RS_Scope_EN = True) Then


            RS_Local()
            RS_visa(False)
            RS_Scope_EN = False
        End If






    End Function

    Function eff_parameter() As Integer
        Dim i, v As Integer
        Dim iout_temp As Integer
        Dim vout_temp() As String
        Dim vin_temp() As String
        Dim iout_set_temp() As String
        Dim update_ok As Boolean = False
        Dim x As Integer

        If (Test_run = True) Or (import_now = True) Or (PartI_first = True) Then
            Exit Function
        End If

        If (check_Efficiency.Checked = True) And (((rbtn_meter_iin.Checked = True) And (check_iin.Checked = True)) Or (rbtn_board_iin.Checked = True)) Then
            gbox_iin.Enabled = True
        Else
            gbox_iin.Enabled = False

        End If





        If (data_eff.Rows.Count > 0) Then


            ReDim vin_temp(data_eff.Rows.Count - 1)
            ReDim vout_temp(data_eff.Rows.Count - 1)
            ReDim iout_set_temp(data_eff.Rows.Count - 1)

            For i = 0 To data_eff.Rows.Count - 1

                vin_temp(i) = data_eff.Rows(i).Cells(0).Value
                vout_temp(i) = data_eff.Rows(i).Cells(1).Value
                iout_set_temp(i) = data_eff.Rows(i).Cells(2).Value
            Next
            update_ok = True


        End If

        'iout=eff*(Iin*vin/vout)


        ' If check_Efficiency.Checked = True Then



        num_iout_auto_stop.Value = Full_load * 1000

        data_eff.Rows.Clear()


        For i = 0 To clist_vout.Items.Count - 1

            If clist_vout.GetItemChecked(i) = True Then
                vout_now = clist_vout.Items(i)


                For v = 0 To data_vin.Rows.Count - 1

                    vin_now = data_vin.Rows(v).Cells(0).Value

                    data_eff.Rows.Add(vin_now.ToString, vout_now.ToString)

                    If (vin_now = 0 Or vout_now = 0) Then

                        iout_temp = num_iin_change.Value
                    Else
                        If (rbtn_iin_manual.Checked = True) Then
                            iout_temp = ((num_iin_change.Value * vin_now) / vout_now)
                        Else
                            iout_temp = (num_pass_eff.Value / 100) * ((num_iin_change.Value * vin_now) / vout_now)
                        End If

                        If iout_temp > Full_load * 1000 Then
                            iout_temp = num_iin_change.Value
                        End If

                    End If


                    If update_ok = True Then

                        For x = 0 To vin_temp.Length - 1

                            If (vout_now.ToString = vout_temp(x)) And (vin_now.ToString = vin_temp(x)) Then
                                iout_temp = iout_set_temp(x)
                                Exit For
                            End If


                        Next

                    End If

                    data_eff.Rows(data_eff.Rows.Count - 1).Cells(2).Value = iout_temp

                    If data_eff.Rows.Count > 0 Then
                        data_eff.CurrentCell = data_eff.Rows(data_eff.Rows.Count - 1).Cells(0)
                    End If

                Next


            End If


        Next





    End Function

    Function Calculate_IOB(ByVal t As Integer) As Boolean
        Dim IOB As Double
        Dim Iout_Max_set As Double
        Dim temp As String

        Dim i As Integer
        Dim iout As Double



        'Iout _Max
        For i = 0 To data_iout.Rows.Count - 1

            If (i = 0) Or (data_iout.Rows(i).Cells(1).Value > Iout_Max_set) Then

                Iout_Max_set = data_iout.Rows(i).Cells(1).Value
            End If

        Next

        'IOB
        'IOB = 0.5 * (vin_now - vout_now) * (vout_now / vin_now) / L / fs_now



        IOB = 0.5 * (vin_now - vout_now) * (ton_now) / L_Value(t)


        'IOUT_Boundary = Math.Round(IOB, 4)
        '取得IOUT_Boundary_Start
        If IOB - num_IOB_Range.Value <= 0 Then
            IOUT_Boundary_Start = 0
        Else
            IOUT_Boundary_Start = IOB - num_IOB_Range.Value
        End If

        i = 0
        For iout = IOUT_Boundary_Start To (IOB + num_IOB_Range.Value) Step num_IOB_step.Value
            If iout < Iout_Max_set Then

                ReDim Preserve IOUT_Boundary_Array(i)
                IOUT_Boundary_Array(i) = Math.Round(iout, 9)
                IOUT_Boundary_Stop = IOUT_Boundary_Array(i)
                i = i + 1
            End If
        Next

        If i = 0 Then
            IOUT_Boundary_Stop = IOB + num_IOB_Range.Value
            Return False
        Else
            Return True
        End If

    End Function

    Function stability_parameter(ByVal num As Integer) As Double()


        Dim i, ii As Integer
        Dim iout_temp() As Double
        Dim IOB_check As Boolean = False

        If (Test_run = True) Or (import_now = True) Or (PartI_first = True) Then
            Exit Function
        End If



        If data_set.Rows.Count = 0 Then
            btn_iout_add.Enabled = False
            data_eff.Rows.Clear()

        Else
            btn_iout_add.Enabled = True

        End If

        If data_iout.Rows.Count = 0 Or data_set.Rows.Count = 0 Then
            data_test.Rows.Clear()
            Exit Function
        End If


        iout_temp = Calculate_iout(data_iout)
        data_test.Rows.Clear()


        If (check_stability.Checked = True) Then


            If (check_IOB.Checked = True) And (check_Force_CCM.Checked = False) Then
                vin_now = test_vin(num)
                vout_now = test_vout(num)
                ton_now = test_ton(num) / (10 ^ 9)

                IOB_check = Calculate_IOB(test_temp(num))

                test_IOB_start(num) = Math.Round(IOUT_Boundary_Start, 4)
                test_IOB_stop(num) = Math.Round(IOUT_Boundary_Stop, 4)

                data_set.Rows(num).Cells("col_IOB_start").Value = Math.Round(IOUT_Boundary_Start, 4)
                data_set.Rows(num).Cells("col_IBO_Stop").Value = Math.Round(IOUT_Boundary_Stop, 4)

            End If



            For i = 0 To iout_temp.Length - 1

                If (i > 0) And (IOB_check = True) Then
                    For ii = 0 To IOUT_Boundary_Array.Length - 1

                        If IOUT_Boundary_Array(ii) > iout_temp(i - 1) And IOUT_Boundary_Array(ii) < iout_temp(i) Then

                            data_test.Rows.Add(IOUT_Boundary_Array(ii))


                        End If


                    Next

                End If

                data_test.Rows.Add(iout_temp(i))

            Next

            ReDim iout_temp(data_test.Rows.Count - 1)

            For i = 0 To data_test.Rows.Count - 1

                iout_temp(i) = data_test.Rows(i).Cells(0).Value

            Next

        End If



        If check_iout_up.Checked = True Then
            For i = iout_temp.Length - 2 To 0 Step -1

                data_test.Rows.Add(iout_temp(i))

            Next

        End If
        Return iout_temp

    End Function

    Function data_list() As Integer
        Dim vcc_num As Integer = 0
        Dim fs_num As Integer = 0
        Dim vout_num As Integer = 0
        If data_VCC.Rows.Count > 0 Then


            For n = 0 To data_VCC.Rows.Count - 1

                vcc_now = data_VCC.Rows(n).Cells(0).Value
                ReDim Preserve total_vcc(vcc_num)
                total_vcc(vcc_num) = vcc_now
                vcc_num = vcc_num + 1
            Next


        Else
            ReDim total_vcc(0)
            total_vcc(0) = 0


        End If

        For i = 0 To clist_fs.Items.Count - 1
            If clist_fs.GetItemChecked(i) = True Then
                fs_now = clist_fs.Items(i)
                If (check_stability.Checked = True) And (check_Force_CCM.Checked = False) And (fs_now = 0) Then
                    error_message("Fs (kHz) cannot be ""0"".")
                    data_vin.Rows.Clear()
                    Exit Function
                End If

                ReDim Preserve total_fs(fs_num)

                total_fs(fs_num) = fs_now
                fs_num = fs_num + 1

            End If

        Next

        For ii = 0 To clist_vout.Items.Count - 1
            If clist_vout.GetItemChecked(ii) = True Then
                vout_now = clist_vout.Items(ii)
                If (check_stability.Checked = True) And (check_Force_CCM.Checked = False) And (vout_now = 0) Then
                    error_message("VOUT (V) cannot be ""0"".")
                    data_vin.Rows.Clear()
                    Exit Function
                End If

                ReDim Preserve total_vout(vout_num)

                total_vout(vout_num) = vout_now
                vout_num = vout_num + 1
            End If
        Next
    End Function

    Function data_set_list() As Integer

        Dim t, n, i, ii, v As Integer
        Dim vcc_temp As String
        Dim TA_temp() As String
        Dim fs_temp() As String
        Dim vout_temp() As String
        Dim vin_temp() As String
        Dim ton_temp() As String
        Dim fs_0_temp() As String
        Dim IOB_start_temp() As String
        Dim IOB_stop_temp() As String
        Dim update_ok As Boolean = False
        Dim temp As String
        Dim TA_now_temp As String

        Dim num As Integer

        Dim x As Integer



        If (import_now = True) Then
            Exit Function
        End If

        data_list()

        If data_vin.Rows.Count = 0 Then
            data_set.Rows.Clear()

            Exit Function
        End If



        '-------------------------
        For v = 0 To data_vin.Rows.Count - 1

            vin_now = data_vin.Rows(v).Cells(0).Value
            If (check_vin_sense.Checked = True) And (vin_now > num_vin_max.Value) Then
                error_message("The set value is larger than ""VIN MAX""!")
                Exit Function

            End If

        Next


        '-------------------------



        If (data_set.Rows.Count > 0) And (check_Force_CCM.Checked = False) Then
            ReDim TA_temp(data_set.Rows.Count - 1)
            ReDim fs_temp(data_set.Rows.Count - 1)
            ReDim vout_temp(data_set.Rows.Count - 1)
            ReDim vin_temp(data_set.Rows.Count - 1)
            ReDim ton_temp(data_set.Rows.Count - 1)
            ReDim fs_0_temp(data_set.Rows.Count - 1)
            ReDim IOB_start_temp(data_set.Rows.Count - 1)
            ReDim IOB_stop_temp(data_set.Rows.Count - 1)

            For i = 0 To data_set.Rows.Count - 1
                TA_temp(i) = data_set.Rows(i).Cells(0).Value
                fs_temp(i) = data_set.Rows(i).Cells(2).Value
                vout_temp(i) = data_set.Rows(i).Cells(3).Value
                vin_temp(i) = data_set.Rows(i).Cells(4).Value
                ton_temp(i) = data_set.Rows(i).Cells(5).Value
                fs_0_temp(i) = data_set.Rows(i).Cells(6).Value
                IOB_start_temp(i) = data_set.Rows(i).Cells(7).Value
                IOB_stop_temp(i) = data_set.Rows(i).Cells(8).Value
            Next
            update_ok = True

        End If

        data_set.Rows.Clear()


        If (Main.check_TA_en.Checked = True) And (Main.data_Temp.Rows.Count > 0) Then
            TA_num = Main.data_Temp.Rows.Count - 1
        Else
            TA_num = 0
        End If


        vin_max = data_vin.Rows(0).Cells(0).Value
        vin_min = data_vin.Rows(0).Cells(0).Value




        num = 0


        For t = 0 To TA_num

            If (Main.check_TA_en.Checked = True) And (Main.data_Temp.Rows.Count > 0) Then
                TA_now_temp = Main.data_Temp.Rows(t).Cells(0).Value
            Else

                TA_now_temp = "25"
            End If



            For n = 0 To total_vcc.Length - 1

                If data_VCC.Rows.Count = 0 Then
                    vcc_temp = ""

                Else
                    vcc_temp = total_vcc(n)
                End If


                For i = 0 To total_fs.Length - 1

                    fs_now = total_fs(i)

                    For ii = 0 To total_vout.Length - 1

                        vout_now = total_vout(ii)

                        For v = 0 To data_vin.Rows.Count - 1

                            vin_now = data_vin.Rows(v).Cells(0).Value

                            If vin_max < vin_now Then
                                vin_max = vin_now
                            End If

                            If vin_min > vin_now Then
                                vin_min = vin_now
                            End If

                            data_set.Rows.Add(TA_now_temp, vcc_temp, fs_now.ToString, vout_now.ToString, vin_now.ToString)

                            ReDim Preserve test_temp(num)
                            ReDim Preserve test_vcc(num)
                            ReDim Preserve test_fs(num)
                            ReDim Preserve test_vout(num)
                            ReDim Preserve test_vin(num)
                            ReDim Preserve test_ton(num)
                            ReDim Preserve test_fs0(num)
                            ReDim Preserve test_IOB_start(num)
                            ReDim Preserve test_IOB_stop(num)

                            test_temp(num) = t
                            test_vcc(num) = vcc_temp
                            test_fs(num) = fs_now.ToString
                            test_vout(num) = vout_now.ToString
                            test_vin(num) = vin_now.ToString
                            test_ton(num) = ""
                            test_fs0(num) = ""
                            test_IOB_start(num) = ""
                            test_IOB_stop(num) = ""


                            If (check_stability.Checked = True) And (check_Force_CCM.Checked = False) Then


                                ton_now = (vout_now / vin_now) * (1 / fs_now)

                                If update_ok = True Then
                                    For x = 0 To TA_temp.Length - 1

                                        If (TA_now_temp = TA_temp(x)) And (fs_now.ToString = fs_temp(x)) And (vout_now.ToString = vout_temp(x)) And (vin_now.ToString = vin_temp(x)) Then
                                            test_ton(num) = ton_temp(x)
                                            test_fs0(num) = fs_0_temp(x)
                                            test_IOB_start(num) = IOB_start_temp(x)
                                            test_IOB_stop(num) = IOB_stop_temp(x)
                                            Exit For
                                        End If

                                    Next

                                End If

                                If test_ton(num) = "" Then
                                    test_ton(num) = ton_now * (10 ^ 9) 'ns
                                End If


                                If test_fs0(num) = "" Then
                                    test_fs0(num) = num_fs_leak.Value
                                End If

                                data_set.Rows(data_set.Rows.Count - 1).Cells(5).Value = test_ton(num)
                                data_set.Rows(data_set.Rows.Count - 1).Cells(6).Value = test_fs0(num)

                                If test_IOB_start(num) = "" Or test_IOB_stop(num) = "" Then
                                    Calculate_IOB(t)

                                    If test_IOB_start(num) = "" Then
                                        test_IOB_start(num) = Math.Round(IOUT_Boundary_Start, 4)
                                    End If

                                    If test_IOB_stop(num) = "" Then
                                        test_IOB_stop(num) = Math.Round(IOUT_Boundary_Stop, 4)
                                    End If
                                End If



                                data_set.Rows(data_set.Rows.Count - 1).Cells(7).Value = test_IOB_start(num)
                                data_set.Rows(data_set.Rows.Count - 1).Cells(8).Value = test_IOB_stop(num)


                            End If

                            num = num + 1


                        Next

                        If data_set.Rows.Count > 0 Then
                            data_set.CurrentCell = data_set.Rows(data_set.Rows.Count - 1).Cells(0)
                        End If

                    Next


                Next

            Next

        Next

        'If (num_vin_max.Value = 0) Or (num_vin_max.Value < (vin_max + 2)) Then
        '    num_vin_max.Value = vin_max + 2
        'End If

        eff_parameter()
        stability_parameter(data_set.Rows.Count - 1)
    End Function

    Function Inst_check_list() As Integer
        ReDim Preserve Scope_check(data_test_now)
        ReDim Preserve Meter_iin_check(data_test_now)
        ReDim Preserve Relay_iin_check(data_test_now)
        ReDim Preserve Meter_iout_check(data_test_now)
        ReDim Preserve Relay_iout_check(data_test_now)
        ReDim Preserve Meter_icc_check(data_test_now)
        ReDim Preserve Load_check(data_test_now)
        ReDim Preserve Power_vcc_check(data_test_now)


        If check_stability.Checked = True Or check_jitter.Checked = True Then
            Scope_check(data_test_now) = True
        Else
            Scope_check(data_test_now) = False
        End If

        Meter_iin_relay_check = False
        Meter_iout_relay_check = False

        If check_Efficiency.Checked = True Then
            If rbtn_meter_iin.Checked = True Then
                Meter_iin_check(data_test_now) = True
                Relay_iin_check(data_test_now) = False
                If check_iin.Checked = True Then
                    Meter_iin_relay_check = True
                End If
            Else
                Meter_iin_check(data_test_now) = False
                Relay_iin_check(data_test_now) = True
            End If
        Else
            Meter_iin_check(data_test_now) = False
            Relay_iin_check(data_test_now) = False
        End If




        If rbtn_meter_iout.Checked = True Then
            If cbox_Iout_meter.SelectedItem = no_device Then
                Meter_iout_check(data_test_now) = False
            Else
                Meter_iout_check(data_test_now) = True
                If check_iout.Checked = True Then
                    Meter_iout_relay_check = True
                End If
            End If

            Relay_iout_check(data_test_now) = False
        ElseIf rbtn_board_iout.Checked = True Then

            Meter_iout_check(data_test_now) = False
            Relay_iout_check(data_test_now) = True
        Else
            Meter_iout_check(data_test_now) = False
            Relay_iout_check(data_test_now) = False
        End If

        If cbox_VCC.SelectedItem <> no_device Then
            Power_vcc_check(data_test_now) = True
        Else
            Power_vcc_check(data_test_now) = False
        End If

        If cbox_Icc_meter.SelectedItem <> no_device Then
            Meter_icc_check(data_test_now) = True
        Else
            Meter_icc_check(data_test_now) = False
        End If

        Load_check(data_test_now) = True
    End Function

    Function Calculate_pass(ByVal t As Integer) As Integer

        Dim M As Double
        Dim I_Leakage As Double
        Dim R_Leakage As Double
        Dim temp As String
        Dim Fs_op As Double

        If check_Force_CCM.Checked = False Then

            '計算R_leak

            If (fs_now = 0) Or (vout_now = 0) Or (vin_now = 0) Then

                Exit Function
            End If



            M = vout_now / vin_now


            I_Leakage = (Fs_leak_0A * (1 - M) * vout_now * ton_now ^ 2) / (2 * L_Value(t) * M ^ 2)

            R_Leakage = vout_now / I_Leakage



            Fs_op = 2 * L_Value(t) * (M) ^ 2 * (iout_now + vout_now / R_Leakage) / ((1 - M) * vout_now * (ton_now) ^ 2) '取得目前的操作頻率(for constant on time control)


            If Fs_op < fs_now Then
                'Operated in DEM
                Fs_CCM = False

                If iout_now = 0 Then
                    Fs_op = Fs_leak_0A

                End If

                Fs_Max = Fs_op * (1 + num_DEM_pos.Value * 0.01)
                Fs_Min = Fs_op * (1 - num_DEM_neg.Value * 0.01)



            Else
                'Operated in CCM

                If iout_now > IOUT_Boundary_Stop Then
                    Fs_CCM = True
                Else
                    Fs_CCM = False
                End If

                Fs_op = fs_now
                Fs_Max = Fs_op * (1 + num_CCM_pos.Value * 0.01)
                Fs_Min = Fs_op * (1 - num_CCM_neg.Value * 0.01)

            End If
        Else

            Fs_CCM = True
            Fs_op = fs_now
            Fs_Max = Fs_op * (1 + num_CCM_pos.Value * 0.01)
            Fs_Min = Fs_op * (1 - num_CCM_neg.Value * 0.01)

        End If


        Fs_Max = Format(Fs_Max, "#0.0")
        Fs_Min = Format(Fs_Min, "#0.0")
    End Function

    Function result_parameter() As Integer
        Dim i, n, ii, v, x, y As Integer
        Dim t As Integer
        Dim stability_iout() As Double
        Dim TA_temp As Integer
        Dim VCC_num As Integer
        Dim VCC_temp As String
        Dim set_num As Integer = 0
        Dim total_iout_num As Integer


        If (Test_run = True) Or (import_now = True) Or (PartI_first = True) Then
            Exit Function
        End If




        data_result.Rows.Clear()
        data_result.Columns("col_test_vcc1").HeaderText = Vcc_name
        data_result.Columns("col_test_vin1").HeaderText = Vin_name
        data_result.Columns("col_test_iout1").HeaderText = Iout_name
        data_result.Columns("col_test_vout1").HeaderText = Vout_name

        ' data_list()
        data_set_list()

        ReDim stability_row_start(data_set.Rows.Count - 1)
        ReDim stability_row_stop(data_set.Rows.Count - 1)

        If data_set.Rows.Count = 0 Then
            If check_LineR.Checked = True And rbtn_lineR_test1.Checked = True And data_lineR_vin.Rows.Count > 0 Then



            Else
                Exit Function
            End If

        End If


        'Total Iout


        total_other_iout = 0
        If (check_Efficiency.Checked = True Or check_loadR.Checked = True) And (data_eff_iout.Rows.Count > 0) Then
            For i = 0 To data_eff_iout.Rows.Count - 1
                ReDim Preserve other_iout(total_other_iout)
                other_iout(total_other_iout) = data_eff_iout.Rows(i).Cells(0).Value
                total_other_iout = total_other_iout + 1
            Next

        End If

        '---------------------------------------------------------------------------------------------------------
        'Jitter
        If (check_jitter.Checked = True) And (data_jitter_iout.Rows.Count > 0) Then
            For i = 0 To data_jitter_iout.Rows.Count - 1
                ReDim Preserve other_iout(total_other_iout)
                other_iout(total_other_iout) = data_jitter_iout.Rows(i).Cells(0).Value
                total_other_iout = total_other_iout + 1
            Next

        End If

        If (check_LineR.Checked = True) And (rbtn_lineR_test2.Checked = True) And (data_lineR_iout.Rows.Count > 0) Then
            For i = 0 To data_lineR_iout.Rows.Count - 1
                ReDim Preserve other_iout(total_other_iout)
                other_iout(total_other_iout) = data_lineR_iout.Rows(i).Cells(0).Value
                total_other_iout = total_other_iout + 1
            Next

        End If




        If (Main.check_TA_en.Checked = True) And (Main.data_Temp.Rows.Count > 0) Then
            TA_num = Main.data_Temp.Rows.Count - 1
        Else
            TA_num = 0
        End If


        For t = 0 To TA_num

            For n = 0 To total_vcc.Length - 1

                If data_VCC.Rows.Count = 0 Then
                    VCC_temp = ""

                Else
                    VCC_temp = total_vcc(n)
                End If


                For i = 0 To total_fs.Length - 1

                    fs_now = total_fs(i)

                    For ii = 0 To total_vout.Length - 1

                        vout_now = total_vout(ii)

                        'PartI Test

                        For v = 0 To data_vin.Rows.Count - 1

                            vin_now = data_vin.Rows(v).Cells(0).Value


                            If total_other_iout > 0 Then
                                ReDim total_iout(total_other_iout - 1)
                                total_iout = other_iout
                                total_iout_num = total_other_iout
                            Else

                                total_iout_num = 0
                            End If



                            If check_stability.Checked = True Then
                                set_num = t * total_vcc.Length * total_fs.Length * total_vout.Length * data_vin.Rows.Count + n * total_fs.Length * total_vout.Length * data_vin.Rows.Count + i * total_vout.Length * data_vin.Rows.Count + ii * data_vin.Rows.Count + v
                                stability_iout = stability_parameter(set_num)



                                If data_test.Rows.Count > 0 Then
                                    For x = 0 To stability_iout.Length - 1
                                        ReDim Preserve total_iout(total_iout_num)
                                        total_iout(total_iout_num) = stability_iout(x)
                                        total_iout_num = total_iout_num + 1
                                    Next
                                End If

                                If check_Force_CCM.Checked = False Then
                                    Fs_leak_0A = test_fs0(set_num)
                                    ton_now = test_ton(set_num) / (10 ^ 9)
                                    IOUT_Boundary_Start = test_IOB_start(set_num)
                                    IOUT_Boundary_Stop = test_IOB_stop(set_num)

                                End If

                            End If

                            '' 過濾重複的陣列元素

                            If total_iout_num = 0 Then
                                Exit For
                            End If


                            Array.Sort(total_iout)

                            total_iout = total_iout.Distinct.ToArray()






                            'Iout Setting
                            For x = 0 To total_iout.Length - 1
                                iout_now = total_iout(x)


                                data_result.Rows.Add(t, VCC_temp, (fs_now / 1000).ToString, vout_now.ToString, vin_now.ToString, total_iout(x))

                                If check_stability.Checked = True Then
                                    For y = 0 To data_test.Rows.Count - 1

                                        If iout_now = data_test.Rows(y).Cells(0).Value Then


                                            If y = 0 Then
                                                stability_row_start(set_num) = data_result.Rows.Count - 1
                                                If data_test.Rows.Count = 1 Then
                                                    stability_row_stop(set_num) = data_result.Rows.Count - 1
                                                End If
                                            Else
                                                stability_row_stop(set_num) = data_result.Rows.Count - 1
                                            End If

                                            data_result.Rows(data_result.Rows.Count - 1).Cells("col_test_stability").Value = iout_now.ToString
                                            Calculate_pass(t)
                                            data_result.Rows(data_result.Rows.Count - 1).Cells("col_Fs_max").Value = Fs_Max / 10 ^ 3
                                            data_result.Rows(data_result.Rows.Count - 1).Cells("col_Fs_min").Value = Fs_Min / 10 ^ 3
                                            data_result.Rows(data_result.Rows.Count - 1).Cells("col_Fs_CCM").Value = Fs_CCM
                                            data_result.Rows(data_result.Rows.Count - 1).Cells("col_IOB_start1").Value = IOUT_Boundary_Start
                                            data_result.Rows(data_result.Rows.Count - 1).Cells("col_IOB_stop1").Value = IOUT_Boundary_Stop
                                            Exit For
                                        End If

                                    Next
                                End If

                                If check_Efficiency.Checked = True Or check_loadR.Checked = True Then
                                    For y = 0 To data_eff_iout.Rows.Count - 1
                                        If iout_now = data_eff_iout.Rows(y).Cells(0).Value Then
                                            data_result.Rows(data_result.Rows.Count - 1).Cells("col_test_eff").Value = iout_now.ToString
                                            Exit For
                                        End If
                                    Next
                                End If

                                If (check_LineR.Checked = True) And (rbtn_lineR_test2.Checked = True) Then
                                    For y = 0 To data_lineR_iout.Rows.Count - 1
                                        If iout_now = data_lineR_iout.Rows(y).Cells(0).Value Then
                                            data_result.Rows(data_result.Rows.Count - 1).Cells("col_test_line").Value = iout_now.ToString

                                            Exit For
                                        End If
                                    Next
                                End If


                                If check_jitter.Checked = True Then
                                    For y = 0 To data_jitter_iout.Rows.Count - 1

                                        If iout_now = data_jitter_iout.Rows(y).Cells(0).Value Then
                                            data_result.Rows(data_result.Rows.Count - 1).Cells("col_test_jitter").Value = iout_now.ToString

                                            If data_result.Rows(data_result.Rows.Count - 1).Cells("col_test_stability").Value = "" Then
                                                data_result.Rows(data_result.Rows.Count - 1).Cells("col_test_stability").Value = iout_now.ToString
                                                Calculate_pass(t)
                                                data_result.Rows(data_result.Rows.Count - 1).Cells("col_Fs_max").Value = Fs_Max / 10 ^ 3
                                                data_result.Rows(data_result.Rows.Count - 1).Cells("col_Fs_min").Value = Fs_Min / 10 ^ 3
                                                data_result.Rows(data_result.Rows.Count - 1).Cells("col_Fs_CCM").Value = Fs_CCM
                                                data_result.Rows(data_result.Rows.Count - 1).Cells("col_IOB_start1").Value = IOUT_Boundary_Start
                                                data_result.Rows(data_result.Rows.Count - 1).Cells("col_IOB_stop1").Value = IOUT_Boundary_Stop
                                            End If

                                            Exit For
                                        End If
                                    Next
                                End If

                            Next


                            'Stability line up

                            If (check_stability.Checked = True) And (check_iout_up.Checked = True) And (data_test.Rows.Count > 0) Then

                                For y = stability_iout.Length - 2 To 0 Step -1
                                    iout_now = stability_iout(y)
                                    data_result.Rows.Add(t, VCC_temp, (fs_now / 1000).ToString, vout_now.ToString, vin_now.ToString, iout_now)
                                    data_result.Rows(data_result.Rows.Count - 1).Cells("col_test_stability").Value = iout_now.ToString
                                    Calculate_pass(test_temp(n))
                                    data_result.Rows(data_result.Rows.Count - 1).Cells("col_Fs_max").Value = Fs_Max / 10 ^ 3
                                    data_result.Rows(data_result.Rows.Count - 1).Cells("col_Fs_min").Value = Fs_Min / 10 ^ 3
                                    data_result.Rows(data_result.Rows.Count - 1).Cells("col_Fs_CCM").Value = Fs_CCM
                                    data_result.Rows(data_result.Rows.Count - 1).Cells("col_IOB_start1").Value = IOUT_Boundary_Start
                                    data_result.Rows(data_result.Rows.Count - 1).Cells("col_IOB_stop1").Value = IOUT_Boundary_Stop
                                Next


                            End If

                        Next

                        'Line Regulation
                        If (check_LineR.Checked = True) And (rbtn_lineR_test1.Checked = True) Then
                            For x = 0 To data_lineR_iout.Rows.Count - 1

                                iout_now = data_lineR_iout.Rows(x).Cells(0).Value

                                For v = 0 To data_lineR_vin.Rows.Count - 1

                                    vin_now = data_lineR_vin.Rows(v).Cells(0).Value
                                    data_result.Rows.Add(t, VCC_temp, (fs_now / 1000).ToString, vout_now.ToString, vin_now.ToString, iout_now)
                                    data_result.Rows(data_result.Rows.Count - 1).Cells("col_test_line").Value = iout_now.ToString

                                Next


                                If check_lineR_up.Checked = True Then
                                    For v = data_lineR_vin.Rows.Count - 2 To 0 Step -1
                                        vin_now = data_lineR_vin.Rows(v).Cells(0).Value

                                        data_result.Rows.Add(t, VCC_temp, (fs_now / 1000).ToString, vout_now.ToString, vin_now.ToString, iout_now)
                                        data_result.Rows(data_result.Rows.Count - 1).Cells("col_test_line").Value = iout_now.ToString

                                    Next

                                End If


                            Next



                        End If


                        If data_set.Rows.Count > 0 Then
                            data_set.CurrentCell = data_set.Rows(data_set.Rows.Count - 1).Cells(0)
                        End If

                    Next


                Next

            Next

        Next

        txt_points.Text = data_result.Rows.Count

        If data_result.Rows.Count > 0 Then
            btn_ok.Enabled = True
        Else
            btn_ok.Enabled = False
        End If

    End Function

    Function Test_set() As Integer

        row = 1
        col = 1

        '------------------------------------------------------------------------------------
        'Initial Page
        'Main
        xlSheet.Cells(row, col) = Stability
        title_set()
        row = row + 1
        xlSheet.Cells(row, col) = "Enable"
        xlSheet.Cells(row, col + 1) = check_stability.Checked
        row = row + 1

        xlSheet.Cells(row, col) = "Test Item"
        xlSheet.Cells(row, col + 1) = "+ Error(%)"
        xlSheet.Cells(row, col + 2) = "- Error(%)"
        row = row + 1
        xlSheet.Cells(row, col) = "VOUT_DC"
        xlSheet.Cells(row, col + 1) = num_vout_pos.Value
        xlSheet.Cells(row, col + 2) = num_vout_neg.Value
        row = row + 1
        xlSheet.Cells(row, col) = "VOUT_AC"
        xlSheet.Cells(row, col + 1) = num_vout_ac.Value

        row = row + 1
        xlSheet.Cells(row, col) = "Fsw_DEM"
        xlSheet.Cells(row, col + 1) = num_DEM_pos.Value
        xlSheet.Cells(row, col + 2) = num_DEM_neg.Value
        row = row + 1
        xlSheet.Cells(row, col) = "Fsw_CCM"
        xlSheet.Cells(row, col + 1) = num_CCM_pos.Value
        xlSheet.Cells(row, col + 2) = num_CCM_neg.Value
        row = row + 1



        xlSheet.Cells(row, col) = "Chart Type"
        xlSheet.Cells(row, col + 1) = cbox_type_stability.SelectedItem
        row = row + 1


        xlSheet.Cells(row, col) = "Sheet Name"
        xlSheet.Cells(row, col + 1) = txt_stability_sheet.Text
        xlSheet.Cells(row, col + 2) = txt_error_sheet.Text
        xlSheet.Cells(row, col + 3) = txt_beta_sheet.Text
        row = row + 1


        '------------------------------------------------------------------------------------
        'Jitter
        xlSheet.Cells(row, col) = Jitter
        title_set()
        row = row + 1
        xlSheet.Cells(row, col) = "Enable"
        xlSheet.Cells(row, col + 1) = check_jitter.Checked
        row = row + 1

        xlSheet.Cells(row, col) = "PASS Criteria"
        xlSheet.Cells(row, col + 1) = num_pass_jitter.Value
        row = row + 1


        xlSheet.Cells(row, col) = "Sheet Name"
        xlSheet.Cells(row, col + 1) = txt_jitter_sheet.Text

        row = row + 1
        '------------------------------------------------------------------------------------
        xlSheet.Cells(row, col) = Efficiency
        title_set()
        row = row + 1
        xlSheet.Cells(row, col) = "Enable"
        xlSheet.Cells(row, col + 1) = check_Efficiency.Checked
        row = row + 1

        xlSheet.Cells(row, col) = "PASS Criteria"
        xlSheet.Cells(row, col + 1) = num_pass_eff.Value
        row = row + 1


        xlSheet.Cells(row, col) = "Chart Type"
        xlSheet.Cells(row, col + 1) = cbox_type_Eff.SelectedItem
        row = row + 1


        xlSheet.Cells(row, col) = "Sheet Name"
        xlSheet.Cells(row, col + 1) = txt_eff_sheet.Text

        row = row + 1

        '------------------------------------------------------------------------------------

        xlSheet.Cells(row, col) = Load_Regulation
        title_set()
        row = row + 1
        xlSheet.Cells(row, col) = "Enable"
        xlSheet.Cells(row, col + 1) = check_loadR.Checked
        row = row + 1

        xlSheet.Cells(row, col) = "PASS Criteria"
        xlSheet.Cells(row, col + 1) = num_pass_loadR.Value
        row = row + 1



        xlSheet.Cells(row, col) = "Chart Type"
        xlSheet.Cells(row, col + 1) = cbox_type_LoadR.SelectedItem
        row = row + 1


        xlSheet.Cells(row, col) = "Sheet Name"
        xlSheet.Cells(row, col + 1) = txt_LoadR_sheet.Text

        row = row + 1
        '------------------------------------------------------------------------------------

        xlSheet.Cells(row, col) = Line_Regulation
        title_set()
        row = row + 1
        xlSheet.Cells(row, col) = "Enable"
        xlSheet.Cells(row, col + 1) = check_LineR.Checked
        row = row + 1

        xlSheet.Cells(row, col) = "PASS Criteria"
        xlSheet.Cells(row, col + 1) = num_pass_lineR.Value
        row = row + 1


        xlSheet.Cells(row, col) = "Chart Type"
        xlSheet.Cells(row, col + 1) = cbox_type_LineR.SelectedItem
        row = row + 1


        xlSheet.Cells(row, col) = "Sheet Name"
        xlSheet.Cells(row, col + 1) = txt_LineR_sheet.Text
        xlSheet.Cells(row, col + 2) = txt_data_sheet.Text
        row = row + 1

        '------------------------------------------------------------------------------------

        '------------------------------------------------------------------------------------
        xlSheet.Cells(row, col) = "VCC"
        title_set()
        row = row + 1

        xlSheet.Cells(row, col) = txt_vcc_name1.Text
        xlSheet.Cells(row, col + 1) = cbox_VCC.SelectedItem
        xlSheet.Cells(row, col + 2) = cbox_VCC_ch.SelectedItem
        xlSheet.Cells(row, col + 3) = cbox_VCC_daq.SelectedItem
        row = row + 1
        data_test_set(data_VCC)

        xlSheet.Cells(row, col) = "ICC"
        title_set()
        row = row + 1
        xlSheet.Cells(row, col) = txt_ivcc_name1.Text
        xlSheet.Cells(row, col + 1) = cbox_Icc_meter.SelectedItem
        xlSheet.Cells(row, col + 2) = txt_Icc_addr.Text

        row = row + 1
        '------------------------------------------------------------------------------------
        '------------------------------------------------------------------------------------
        'Instrument
        xlSheet.Cells(row, col) = "VIN"
        title_set()
        row = row + 1
        xlSheet.Cells(row, col) = txt_vin_name1.Text
        xlSheet.Cells(row, col + 1) = cbox_vin.SelectedItem
        xlSheet.Cells(row, col + 2) = cbox_vin_ch.SelectedItem
        xlSheet.Cells(row, col + 3) = num_VIN_OCP.Value
        xlSheet.Cells(row, col + 4) = cbox_vin_daq.SelectedItem
        row = row + 1
        xlSheet.Cells(row, col) = "Sense Vin"
        title_set()
        row = row + 1
        xlSheet.Cells(row, col) = check_vin_sense.Checked
        xlSheet.Cells(row, col + 1) = num_vin_sense.Value
        xlSheet.Cells(row, col + 2) = num_vin_max.Value
        row = row + 1

        '------------------------------------------------------------------------------------
        xlSheet.Cells(row, col) = "IIN"
        title_set()
        row = row + 1
        xlSheet.Cells(row, col) = txt_iin_name1.Text
        xlSheet.Cells(row, col + 1) = rbtn_meter_iin.Checked
        xlSheet.Cells(row, col + 2) = cbox_IIN_meter.SelectedItem
        xlSheet.Cells(row, col + 3) = txt_IIN_addr.Text
        xlSheet.Cells(row, col + 4) = check_iin.Checked
        xlSheet.Cells(row, col + 5) = cbox_IIN_relay.SelectedItem
        xlSheet.Cells(row, col + 6) = num_iin_change.Value
        xlSheet.Cells(row, col + 7) = rbtn_Iin_PW.Checked
        xlSheet.Cells(row, col + 8) = rbtn_iin_current_measure.Checked
        row = row + 1


        '------------------------------------------------------------------------------------
        xlSheet.Cells(row, col) = "Vout"
        title_set()
        row = row + 1
        xlSheet.Cells(row, col) = txt_vout_name1.Text
        xlSheet.Cells(row, col + 1) = cbox_vout_daq.SelectedItem
        xlSheet.Cells(row, col + 2) = cbox_vout1_daq.SelectedItem
        row = row + 1
        xlSheet.Cells(row, col) = "Check Vout"
        title_set()
        row = row + 1
        xlSheet.Cells(row, col) = check_shutdown.Checked
        xlSheet.Cells(row, col + 1) = num_Vout_error.Value
        row = row + 1

        '------------------------------------------------------------------------------------
        xlSheet.Cells(row, col) = "IOUT"
        title_set()
        scope_time_init()
        Display_persistence(False)
        ' dut1 data to excel
        xlSheet.Cells(row, col + 2) = cbox_Iout_meter.SelectedItem
        xlSheet.Cells(row, col + 3) = txt_Iout_addr.Text
        xlSheet.Cells(row, col + 4) = check_iout.Checked
        xlSheet.Cells(row, col + 5) = cbox_Iout_relay.SelectedItem
        xlSheet.Cells(row, col + 6) = num_iout_change.Value
        xlSheet.Cells(row, col + 7) = rbtn_board_iout.Checked
        xlSheet.Cells(row, col + 8) = cbox_board_buck.SelectedItem
        xlSheet.Cells(row, col + 9) = rbtn_iout_load.Checked
        xlSheet.Cells(row, col + 10) = rbtn_iout_current_measure.Checked
        row = row + 1

        '------------------------------------------------------------------------------------


        xlSheet.Cells(row, col) = "DC Load"
        title_set()
        row = row + 1
        xlSheet.Cells(row, col) = "Channel"
        xlSheet.Cells(row, col + 1) = check_IOUT_ch1.Checked
        xlSheet.Cells(row, col + 2) = check_IOUT_ch2.Checked
        xlSheet.Cells(row, col + 3) = check_IOUT_ch3.Checked
        xlSheet.Cells(row, col + 4) = check_IOUT_ch4.Checked
        row = row + 1

        xlSheet.Cells(row, col) = "Delay"
        xlSheet.Cells(row, col + 1) = num_delay.Value
        xlSheet.Cells(row, col + 2) = cbox_delay_unit.SelectedItem
        xlSheet.Cells(row, col + 3) = "Iout(A)  >"
        xlSheet.Cells(row, col + 4) = num_iout_delay.Value
        row = row + 1

        '------------------------------------------------------------------------------------
        xlSheet.Cells(row, col) = "DAQ"
        title_set()
        row = row + 1
        xlSheet.Cells(row, col) = "Numbers of Trigger:"
        xlSheet.Cells(row, col + 1) = num_data_count.Value
        row = row + 1
        xlSheet.Cells(row, col) = "Resolution:"
        xlSheet.Cells(row, col + 1) = cbox_data_resolution.SelectedItem
        row = row + 1
        '------------------------------------------------------------------------------------

        'Instrument
        ' Ollie_note: add 2nd page conditions
        xlSheet.Cells(row, col) = "VIN2"
        title_set()
        row = row + 1
        xlSheet.Cells(row, col) = TextBox23.Text
        xlSheet.Cells(row, col + 1) = cbox_vin2.SelectedItem
        xlSheet.Cells(row, col + 2) = cbox_vin_ch2.SelectedItem
        xlSheet.Cells(row, col + 3) = num_VIN_OCP2.Value
        xlSheet.Cells(row, col + 4) = cbox_vin_daq2.SelectedItem
        row = row + 1
        xlSheet.Cells(row, col) = "Sense Vin2"
        title_set()
        row = row + 1
        xlSheet.Cells(row, col) = check_vin_sense2.Checked
        xlSheet.Cells(row, col + 1) = num_vin_sense2.Value
        xlSheet.Cells(row, col + 2) = num_vin_max2.Value
        row = row + 1

        '------------------------------------------------------------------------------------
        xlSheet.Cells(row, col) = "IIN2"
        title_set()
        row = row + 1
        xlSheet.Cells(row, col) = TextBox21.Text
        xlSheet.Cells(row, col + 1) = rbtn_meter_iin2.Checked
        xlSheet.Cells(row, col + 2) = cbox_IIN_meter2.SelectedItem
        xlSheet.Cells(row, col + 3) = txt_IIN_addr2.Text
        xlSheet.Cells(row, col + 4) = check_iin2.Checked
        xlSheet.Cells(row, col + 5) = cbox_IIN_relay2.SelectedItem
        xlSheet.Cells(row, col + 6) = num_iin_change2.Value
        xlSheet.Cells(row, col + 7) = rbtn_Iin_PW2.Checked
        xlSheet.Cells(row, col + 8) = rbtn_iin_current_measure2.Checked
        row = row + 1


        '------------------------------------------------------------------------------------
        xlSheet.Cells(row, col) = "Vout2"
        title_set()
        row = row + 1
        xlSheet.Cells(row, col) = TextBox16.Text
        xlSheet.Cells(row, col + 1) = cbox_vout_daq2.SelectedItem
        xlSheet.Cells(row, col + 2) = cbox_vout1_daq2.SelectedItem
        row = row + 1
        xlSheet.Cells(row, col) = "Check Vout2"
        title_set()
        row = row + 1
        xlSheet.Cells(row, col) = check_shutdown2.Checked
        xlSheet.Cells(row, col + 1) = num_Vout_error2.Value
        row = row + 1

        '------------------------------------------------------------------------------------
        xlSheet.Cells(row, col) = "IOUT2"
        title_set()
        row = row + 1
        xlSheet.Cells(row, col) = TextBox19.Text
        xlSheet.Cells(row, col + 1) = rbtn_meter_iout2.Checked
        xlSheet.Cells(row, col + 2) = cbox_Iout_meter2.SelectedItem
        xlSheet.Cells(row, col + 3) = txt_Iout_addr2.Text
        xlSheet.Cells(row, col + 4) = check_iout2.Checked
        xlSheet.Cells(row, col + 5) = cbox_Iout_relay2.SelectedItem
        xlSheet.Cells(row, col + 6) = num_iout_change2.Value
        xlSheet.Cells(row, col + 7) = rbtn_board_iout2.Checked
        xlSheet.Cells(row, col + 8) = cbox_board_buck2.SelectedItem
        xlSheet.Cells(row, col + 9) = rbtn_iout_load2.Checked
        xlSheet.Cells(row, col + 10) = rbtn_iout_current_measure2.Checked
        row = row + 1

        '------------------------------------------------------------------------------------


        xlSheet.Cells(row, col) = "DC Load2"
        title_set()
        row = row + 1
        xlSheet.Cells(row, col) = "Channel"
        xlSheet.Cells(row, col + 1) = check_IOUT_ch12.Checked
        xlSheet.Cells(row, col + 2) = check_IOUT_ch22.Checked
        xlSheet.Cells(row, col + 3) = check_IOUT_ch32.Checked
        xlSheet.Cells(row, col + 4) = check_IOUT_ch42.Checked
        row = row + 1

        xlSheet.Cells(row, col) = "Delay"
        xlSheet.Cells(row, col + 1) = num_delay2.Value
        xlSheet.Cells(row, col + 2) = cbox_delay_unit2.SelectedItem
        xlSheet.Cells(row, col + 3) = "Iout(A)  >"
        xlSheet.Cells(row, col + 4) = num_iout_delay2.Value
        row = row + 1

        '------------------------------------------------------------------------------------
        xlSheet.Cells(row, col) = "DAQ2"
        title_set()
        row = row + 1
        xlSheet.Cells(row, col) = "Numbers of Trigger:"
        xlSheet.Cells(row, col + 1) = num_data_count2.Value
        row = row + 1
        xlSheet.Cells(row, col) = "Resolution:"
        xlSheet.Cells(row, col + 1) = cbox_data_resolution2.SelectedItem
        row = row + 1
        '------------------------------------------------------------------------------------

        xlSheet.Cells(row, col) = "Meter Board"
        title_set()
        row = row + 1
        xlSheet.Cells(row, col) = "Iin Meter board Low Config:"
        xlSheet.Cells(row, col + 1) = num_slave_in_L.Value
        xlSheet.Cells(row, col + 2) = num_comp_in_L.Value
        xlSheet.Cells(row, col + 3) = num_resolution_in_L.Value
        row = row + 1
        xlSheet.Cells(row, col) = "Iin Meter board Mid Config:"
        xlSheet.Cells(row, col + 1) = num_slave_in_M.Value
        xlSheet.Cells(row, col + 2) = num_comp_in_M.Value
        xlSheet.Cells(row, col + 3) = num_resolution_in_M.Value
        row = row + 1
        xlSheet.Cells(row, col) = "Iin Meter board High Config:"
        xlSheet.Cells(row, col + 1) = num_slave_in_H.Value
        xlSheet.Cells(row, col + 2) = num_comp_in_H.Value
        xlSheet.Cells(row, col + 3) = num_resolution_in_H.Value
        row = row + 1
        xlSheet.Cells(row, col) = "Iin Meter board IO Config:"
        xlSheet.Cells(row, col + 1) = num_slave_in_IO.Value

        row = row + 1
        xlSheet.Cells(row, col) = "Iout Meter board Low Config:"
        xlSheet.Cells(row, col + 1) = num_slave_out_L.Value
        xlSheet.Cells(row, col + 2) = num_comp_out_L.Value
        xlSheet.Cells(row, col + 3) = num_resolution_out_L.Value
        row = row + 1
        xlSheet.Cells(row, col) = "Iout Meter board Mid Config:"
        xlSheet.Cells(row, col + 1) = num_slave_out_M.Value
        xlSheet.Cells(row, col + 2) = num_comp_out_M.Value
        xlSheet.Cells(row, col + 3) = num_resolution_out_M.Value
        row = row + 1
        xlSheet.Cells(row, col) = "Iout Meter board High Config:"
        xlSheet.Cells(row, col + 1) = num_slave_out_H.Value
        xlSheet.Cells(row, col + 2) = num_comp_out_H.Value
        xlSheet.Cells(row, col + 3) = num_resolution_out_H.Value
        row = row + 1
        xlSheet.Cells(row, col) = "Iout Meter board IO Config:"
        xlSheet.Cells(row, col + 1) = num_slave_out_IO.Value

        ' 2nd board setting
        row = row + 1
        xlSheet.Cells(row, col) = "Iin Meter board2 Low Config:"
        xlSheet.Cells(row, col + 1) = num_slave_in_L2.Value
        xlSheet.Cells(row, col + 2) = num_comp_in_L2.Value
        xlSheet.Cells(row, col + 3) = num_resolution_in_L2.Value
        row = row + 1
        xlSheet.Cells(row, col) = "Iin Meter board2 Mid Config:"
        xlSheet.Cells(row, col + 1) = num_slave_in_M2.Value
        xlSheet.Cells(row, col + 2) = num_comp_in_M2.Value
        xlSheet.Cells(row, col + 3) = num_resolution_in_M2.Value
        row = row + 1
        xlSheet.Cells(row, col) = "Iin Meter board2 High Config:"
        xlSheet.Cells(row, col + 1) = num_slave_in_H2.Value
        xlSheet.Cells(row, col + 2) = num_comp_in_H2.Value
        xlSheet.Cells(row, col + 3) = num_resolution_in_H2.Value
        row = row + 1
        xlSheet.Cells(row, col) = "Iin Meter board2 IO Config:"
        xlSheet.Cells(row, col + 1) = num_slave_in_IO2.Value

        row = row + 1
        xlSheet.Cells(row, col) = "Iout Meter board2 Low Config:"
        xlSheet.Cells(row, col + 1) = num_slave_out_L2.Value
        xlSheet.Cells(row, col + 2) = num_comp_out_L2.Value
        xlSheet.Cells(row, col + 3) = num_resolution_out_L2.Value
        row = row + 1
        xlSheet.Cells(row, col) = "Iout Meter board2 Mid Config:"
        xlSheet.Cells(row, col + 1) = num_slave_out_M2.Value
        xlSheet.Cells(row, col + 2) = num_comp_out_M2.Value
        xlSheet.Cells(row, col + 3) = num_resolution_out_M2.Value
        row = row + 1
        xlSheet.Cells(row, col) = "Iout Meter board2 High Config:"
        xlSheet.Cells(row, col + 1) = num_slave_out_H2.Value
        xlSheet.Cells(row, col + 2) = num_comp_out_H2.Value
        xlSheet.Cells(row, col + 3) = num_resolution_out_H2.Value
        row = row + 1
        xlSheet.Cells(row, col) = "Iout Meter board2 IO Config:"
        xlSheet.Cells(row, col + 1) = num_slave_out_IO2.Value
        row = row + 1
        xlSheet.Cells(row, col) = "Meter Count" ' num_meter_count
        xlSheet.Cells(row, col + 1) = num_meter_count.Value
        row = row + 1
        ''------------------------------------------------------------------------------------
        'Scope
        'Step4
        xlSheet.Cells(row, col) = "Scope"
        title_set()
        row = row + 1
        xlSheet.Cells(row, col) = txt_scope_vin.Text
        xlSheet.Cells(row, col + 1) = cbox_channel_vin.SelectedItem
        xlSheet.Cells(row, col + 2) = cbox_coupling_vin.SelectedItem
        xlSheet.Cells(row, col + 3) = num_offset_vin.Value
        xlSheet.Cells(row, col + 4) = num_position_vin.Value
        xlSheet.Cells(row, col + 5) = cbox_BW_vin.SelectedItem
        xlSheet.Cells(row, col + 6) = num_vin_scale.Value
        xlSheet.Cells(row, col + 7) = check_scope_vin.Checked
        row = row + 1

        xlSheet.Cells(row, col) = txt_scope_iout.Text
        xlSheet.Cells(row, col + 1) = cbox_channel_iout.SelectedItem
        xlSheet.Cells(row, col + 2) = cbox_coupling_iout.SelectedItem
        xlSheet.Cells(row, col + 3) = num_offset_iout.Value
        xlSheet.Cells(row, col + 4) = num_position_iout.Value
        xlSheet.Cells(row, col + 5) = cbox_BW_iout.SelectedItem
        xlSheet.Cells(row, col + 6) = check_scope_iout.Checked
        row = row + 1


        xlSheet.Cells(row, col) = txt_scope_lx.Text
        xlSheet.Cells(row, col + 1) = cbox_channel_lx.SelectedItem
        xlSheet.Cells(row, col + 2) = cbox_coupling_lx.SelectedItem
        xlSheet.Cells(row, col + 3) = num_offset_lx.Value
        xlSheet.Cells(row, col + 4) = num_position_lx.Value
        xlSheet.Cells(row, col + 5) = cbox_BW_lx.SelectedItem
        xlSheet.Cells(row, col + 6) = num_lx_scale.Value
        xlSheet.Cells(row, col + 7) = rbtn_manual_lx.Checked
        xlSheet.Cells(row, col + 8) = num_scale_lx.Value
        row = row + 1


        xlSheet.Cells(row, col) = txt_scope_vout.Text
        xlSheet.Cells(row, col + 1) = cbox_channel_vout.SelectedItem
        xlSheet.Cells(row, col + 2) = cbox_coupling_vout.SelectedItem
        xlSheet.Cells(row, col + 3) = check_offset_vout.Checked
        xlSheet.Cells(row, col + 4) = num_position_vout.Value
        xlSheet.Cells(row, col + 5) = cbox_BW_vout.SelectedItem
        xlSheet.Cells(row, col + 6) = rbtn_auto_vout.Checked
        xlSheet.Cells(row, col + 7) = num_vout_auto.Value
        xlSheet.Cells(row, col + 8) = Check_fixed.Checked
        xlSheet.Cells(row, col + 9) = num_vout_DEM.Value
        xlSheet.Cells(row, col + 10) = num_vout_CCM.Value
        row = row + 1

        xlSheet.Cells(row, col) = txt_scope_lx2.Text
        xlSheet.Cells(row, col + 1) = cbox_channel_lx2.SelectedItem
        xlSheet.Cells(row, col + 2) = cbox_coupling_lx2.SelectedItem
        xlSheet.Cells(row, col + 3) = num_offset_lx2.Value
        xlSheet.Cells(row, col + 4) = num_position_lx2.Value
        xlSheet.Cells(row, col + 5) = cbox_BW_lx2.SelectedItem
        row = row + 1

        xlSheet.Cells(row, col) = txt_scope_vout2.Text
        xlSheet.Cells(row, col + 1) = cbox_channel_vout2.SelectedItem
        xlSheet.Cells(row, col + 2) = cbox_coupling_vout2.SelectedItem
        xlSheet.Cells(row, col + 3) = check_offset_vout2.Checked
        xlSheet.Cells(row, col + 4) = num_position_vout2.Value
        xlSheet.Cells(row, col + 5) = cbox_BW_vout2.SelectedItem
        row = row + 1




        xlSheet.Cells(row, col) = "Time Setting"
        title_set()
        row = row + 1

        xlSheet.Cells(row, col) = "RL (K)"
        xlSheet.Cells(row, col + 1) = num_RL.Value
        row = row + 1

        xlSheet.Cells(row, col) = "Samples (MS/s)"
        xlSheet.Cells(row, col + 1) = num_points.Value
        row = row + 1

        xlSheet.Cells(row, col) = "0s Location"
        xlSheet.Cells(row, col + 1) = num_location.Value
        row = row + 1

        xlSheet.Cells(row, col) = "Acquired Counts"
        xlSheet.Cells(row, col + 1) = num_counts_DEM.Value
        xlSheet.Cells(row, col + 2) = num_counts_CCM.Value
        row = row + 1

        xlSheet.Cells(row, col) = "Phase number"
        xlSheet.Cells(row, col + 1) = num_wave.Value
        row = row + 1

        xlSheet.Cells(row, col) = "LX Trigger"
        xlSheet.Cells(row, col + 1) = rbtn_vin_trigger.Checked
        xlSheet.Cells(row, col + 2) = num_vin_trigger.Value
        xlSheet.Cells(row, col + 3) = rbtn_auto_trigger.Checked
        row = row + 1

        xlSheet.Cells(row, col) = "Measurement"
        title_set()
        row = row + 1

        xlSheet.Cells(row, col) = txt_meas1_ch.Text
        xlSheet.Cells(row, col + 1) = txt_type1.Text
        xlSheet.Cells(row, col + 2) = txt_meas1.Text
        row = row + 1

        xlSheet.Cells(row, col) = txt_meas2_ch.Text
        xlSheet.Cells(row, col + 1) = txt_type2.Text
        xlSheet.Cells(row, col + 2) = txt_meas2.Text
        row = row + 1

        xlSheet.Cells(row, col) = txt_meas3_ch.Text
        xlSheet.Cells(row, col + 1) = txt_type3.Text
        xlSheet.Cells(row, col + 2) = txt_meas3.Text
        row = row + 1

        xlSheet.Cells(row, col) = txt_meas4_ch.Text
        xlSheet.Cells(row, col + 1) = txt_type4.Text
        xlSheet.Cells(row, col + 2) = txt_meas4.Text
        row = row + 1

        xlSheet.Cells(row, col) = txt_meas5_ch.Text
        xlSheet.Cells(row, col + 1) = txt_type5.Text
        xlSheet.Cells(row, col + 2) = txt_meas5.Text
        row = row + 1

        xlSheet.Cells(row, col) = txt_meas6_ch.Text
        xlSheet.Cells(row, col + 1) = txt_type6.Text
        xlSheet.Cells(row, col + 2) = txt_meas6.Text
        row = row + 1
        '------------------------------------------------------------------------------------

        'Test Conditions page

        xlSheet.Cells(row, col) = "Test Init"
        title_set()
        row = row + 1

        xlSheet.Cells(row, col) = "Fsw (Hz)"

        For i = 0 To clist_fs.Items.Count - 1
            xlSheet.Cells(row, col + 1 + i) = clist_fs.GetItemChecked(i)
        Next

        row = row + 1

        xlSheet.Cells(row, col) = "VOUT (V)"

        For i = 0 To clist_vout.Items.Count - 1
            xlSheet.Cells(row, col + 1 + i) = clist_vout.GetItemChecked(i)
        Next

        row = row + 1
        data_test_set(data_vin)


        '------------------------------------------------------------------------------------

        xlSheet.Cells(row, col) = "Stability"
        title_set()
        row = row + 1

        xlSheet.Cells(row, col) = check_Force_CCM.Text
        xlSheet.Cells(row, col + 1) = check_Force_CCM.Checked

        row = row + 1

        xlSheet.Cells(row, col) = "Fs_leak_0A"
        xlSheet.Cells(row, col + 1) = num_fs_leak.Value

        row = row + 1

        xlSheet.Cells(row, col) = check_IOB.Text
        xlSheet.Cells(row, col + 1) = check_IOB.Checked
        xlSheet.Cells(row, col + 2) = "Range (A)"
        xlSheet.Cells(row, col + 3) = num_IOB_Range.Value
        xlSheet.Cells(row, col + 4) = "Step (A)"
        xlSheet.Cells(row, col + 5) = num_IOB_step.Value

        row = row + 1

        xlSheet.Cells(row, col) = check_iout_up.Text
        xlSheet.Cells(row, col + 1) = check_iout_up.Checked
        row = row + 1


        xlSheet.Cells(row, col) = "Test Conditions"
        title_set()
        row = row + 1

        data_test_set(data_set)

        data_test_set(data_iout)
        '------------------------------------------------------------------------------------
        xlSheet.Cells(row, col) = "Efficiency and Load Regulation"
        title_set()
        row = row + 1
        data_test_set(data_eff_iout)

        '------------------------------------------------------------------------------------
        xlSheet.Cells(row, col) = "Jitter"
        title_set()
        row = row + 1
        data_test_set(data_jitter_iout)

        '------------------------------------------------------------------------------------

        '------------------------------------------------------------------------------------
        'Test setup Page
        '------------------------------------------------------------------------------------
        'Test setup Page
        xlSheet.Cells(row, col) = "Stability Set"
        title_set()
        row = row + 1

        xlSheet.Cells(row, col) = " Auto Scanning"
        If rbtn_auto_all.Checked = True Then
            xlSheet.Cells(row, col + 1) = rbtn_auto_all.Text
        Else
            xlSheet.Cells(row, col + 1) = rbtn_auto_DEM.Text
        End If
        row = row + 1

        xlSheet.Cells(row, col) = "Add Cursors"
        xlSheet.Cells(row, col + 1) = check_cursors.Checked
        row = row + 1


        xlSheet.Cells(row, col) = "Ton (min)"
        xlSheet.Cells(row, col + 1) = num_ton_vin.Value
        xlSheet.Cells(row, col + 2) = rbtn_ton_cal.Checked
        xlSheet.Cells(row, col + 3) = num_ton_cal.Value
        xlSheet.Cells(row, col + 4) = rbtn_ton_val.Checked
        xlSheet.Cells(row, col + 5) = num_ton_val.Value
        row = row + 1

        xlSheet.Cells(row, col) = "Toff (min)"
        xlSheet.Cells(row, col + 1) = num_toff_vin.Value
        xlSheet.Cells(row, col + 2) = rbtn_toff_cal.Checked
        xlSheet.Cells(row, col + 3) = num_toff_cal.Value
        xlSheet.Cells(row, col + 4) = rbtn_toff_val.Checked
        xlSheet.Cells(row, col + 5) = num_toff_val.Value
        row = row + 1

        xlSheet.Cells(row, col) = " Update: "
        If rbtn_freq_high.Checked = True Then
            xlSheet.Cells(row, col + 1) = rbtn_freq_high.Text
        Else
            xlSheet.Cells(row, col + 1) = rbtn_freq_low.Text
        End If

        row = row + 1

        xlSheet.Cells(row, col) = "Error"
        xlSheet.Cells(row, col + 1) = num_delay_error.Value
        xlSheet.Cells(row, col + 2) = check_error_pic.Checked

        row = row + 1


        xlSheet.Cells(row, col) = "Capture All"
        xlSheet.Cells(row, col + 1) = check_stability_pic.Checked

        row = row + 1


        xlSheet.Cells(row, col) = "Efficiency Set"
        title_set()
        row = row + 1

        xlSheet.Cells(row, col) = "Auto"
        xlSheet.Cells(row, col + 1) = rbtn_iin_auto.Checked
        xlSheet.Cells(row, col + 2) = rbtn_iin_manual.Checked
        row = row + 1



        xlSheet.Cells(row, col) = "IOUT (mA)"
        xlSheet.Cells(row, col + 1) = num_iin_step.Value
        xlSheet.Cells(row, col + 2) = num_iout_auto_stop.Value

        row = row + 1

        data_test_set(data_eff)


        xlSheet.Cells(row, col) = "Minimum range"
        xlSheet.Cells(row, col + 1) = cbox_meter_mini.SelectedItem


        row = row + 1
        'Jitter Page
        xlSheet.Cells(row, col) = "Jitter Set"
        title_set()
        row = row + 1
        xlSheet.Cells(row, col) = "PERSistence"
        xlSheet.Cells(row, col + 1) = check_persistence.Checked
        xlSheet.Cells(row, col + 2) = num_counts_Jitter.Value
        row = row + 1
        xlSheet.Cells(row, col) = "FastAcq"
        xlSheet.Cells(row, col + 1) = check_fastAcq.Checked
        xlSheet.Cells(row, col + 2) = num_FastAcq.Value
        row = row + 1
        '------------------------------------------------------------------------------------
        xlSheet.Cells(row, col) = "Line Regulation"
        title_set()
        row = row + 1
        xlSheet.Cells(row, col) = "RUN Mode"
        xlSheet.Cells(row, col + 1) = rbtn_lineR_test1.Checked
        row = row + 1
        data_test_set(data_lineR_iout)
        data_test_set(data_lineR_vin)

        xlSheet.Cells(row, col) = "LineR Set"
        xlSheet.Cells(row, col + 1) = check_lineR_scope.Checked
        '------------------------------------------------------------------------------------
        row = row + 1
        xlSheet.Cells(row, col) = "Dut2 En"
        title_set()
        row = row + 1
        xlSheet.Cells(row, col) = "Dut2 Check"
        xlSheet.Cells(row, col + 1) = check_DUT2.Checked

        xlSheet.Columns(1).AutoFit()
        FinalReleaseComObject(xlSheet)
        xlSheet = Nothing
        xlBook.Save()
    End Function

    Function Test_import() As Integer
        Dim i, ii As Integer
        Dim import_ok As Boolean = False
        Dim last_col As Integer
        Dim temp As String
        Dim ton_temp As Double
        Dim num As Integer = 0


        row = 1
        col = 1
        Tab_Set.SelectedIndex = 0
        last_col = xlSheet.Range(ConvertToLetter(col) & row).CurrentRegion.Columns.Count
        'xlSheet.Cells(row, col) = Stability
        'title_set()
        row = row + 1
        'xlSheet.Cells(row, col) = "Enable"
        check_stability.Checked = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 1).Value
        row = row + 1

        'xlSheet.Cells(row, col) = "Test Item"
        'xlSheet.Cells(row, col + 1) = "+ Error(%)"
        'xlSheet.Cells(row, col + 2) = "- Error(%)"
        row = row + 1
        'xlSheet.Cells(row, col) = "VOUT_DC"
        num_vout_pos.Value = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 1).Value
        num_vout_neg.Value = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 2).Value
        row = row + 1
        'xlSheet.Cells(row, col) = "VOUT_AC"
        num_vout_ac.Value = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 1).Value

        row = row + 1
        'xlSheet.Cells(row, col) = "Fsw_DEM"
        num_DEM_pos.Value = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 1).Value
        num_DEM_neg.Value = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 2).Value
        row = row + 1
        ' xlSheet.Cells(row, col) = "Fsw_CCM"
        num_CCM_pos.Value = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 1).Value
        num_CCM_neg.Value = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 2).Value
        row = row + 1



        ' xlSheet.Cells(row, col) = "Chart Type"
        cbox_type_stability.SelectedItem = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 1).Value
        row = row + 1

        'xlSheet.Cells(row, col) = "Sheet Name"
        txt_stability_sheet.Text = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 1).Value
        txt_error_sheet.Text = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 2).Value
        txt_beta_sheet.Text = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 3).Value
        row = row + 1


        '------------------------------------------------------------------------------------
        'Jitter

        'xlSheet.Cells(row, col) = "Jitter"
        'title_set()
        row = row + 1
        'xlSheet.Cells(row, col) = "Enable"
        check_jitter.Checked = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 1).Value
        row = row + 1

        'xlSheet.Cells(row, col) = "PASS Criteria"
        num_pass_jitter.Value = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 1).Value
        row = row + 1


        'xlSheet.Cells(row, col) = "Sheet Name"
        txt_jitter_sheet.Text = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 1).Value

        row = row + 1
        '------------------------------------------------------------------------------------
        'xlSheet.Cells(row, col) = "Efficiency"
        'title_set()
        row = row + 1
        ' xlSheet.Cells(row, col) = "Enable"
        check_Efficiency.Checked = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 1).Value
        row = row + 1

        ' xlSheet.Cells(row, col) = "PASS Criteria"
        num_pass_eff.Value = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 1).Value
        row = row + 1


        'xlSheet.Cells(row, col) = "Chart Type"
        cbox_type_Eff.SelectedItem = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 1).Value
        row = row + 1


        'xlSheet.Cells(row, col) = "Sheet Name"
        txt_eff_sheet.Text = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 1).Value

        row = row + 1
        '------------------------------------------------------------------------------------


        ' xlSheet.Cells(row, col) = "Load Regulation"
        ' title_set()
        row = row + 1
        'xlSheet.Cells(row, col) = "Enable"
        check_loadR.Checked = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 1).Value
        row = row + 1

        ' xlSheet.Cells(row, col) = "PASS Criteria"
        num_pass_loadR.Value = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 1).Value
        row = row + 1



        'xlSheet.Cells(row, col) = "Chart Type"
        cbox_type_LoadR.SelectedItem = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 1).Value
        row = row + 1


        'xlSheet.Cells(row, col) = "Sheet Name"
        txt_LoadR_sheet.Text = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 1).Value

        row = row + 1
        '-------------------------------------------------------------------------------------


        'xlSheet.Cells(row, col) = "Line Regulation"
        'title_set()
        row = row + 1
        'xlSheet.Cells(row, col) = "Enable"
        check_LineR.Checked = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 1).Value
        row = row + 1

        'xlSheet.Cells(row, col) = "PASS Criteria"
        num_pass_lineR.Value = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 1).Value
        row = row + 1



        'xlSheet.Cells(row, col) = "Chart Type"
        cbox_type_LineR.SelectedItem = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 1).Value
        row = row + 1


        'xlSheet.Cells(row, col) = "Sheet Name"
        txt_LineR_sheet.Text = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 1).Value
        txt_data_sheet.Text = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 2).Value
        row = row + 1

        '------------------------------------------------------------------------------------

        '------------------------------------------------------------------------------------
        'xlSheet.Cells(row, col) = "VCC"
        'title_set()
        row = row + 1

        txt_vcc_name1.Text = xlSheet.Range(ConvertToLetter(col) & row).Value
        For i = 0 To cbox_VCC.Items.Count - 1
            If cbox_VCC.Items(i) = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 1).Value Then
                cbox_VCC.SelectedIndex = i

                import_ok = True
                Exit For
            End If
        Next

        If import_ok = False Then
            cbox_VCC.SelectedIndex = 0
        End If

        cbox_VCC_ch.SelectedItem = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 2).Value
        cbox_VCC_daq.SelectedItem = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 3).Value
        row = row + 1
        data_test_import(data_VCC, last_col)

        '------------------------------------------------------------------------------------
        'xlSheet.Cells(row, col) = "ICC"
        'title_set()
        row = row + 1
        txt_ivcc_name1.Text = xlSheet.Range(ConvertToLetter(col) & row).Value

        For i = 0 To cbox_Icc_meter.Items.Count - 1
            If cbox_Icc_meter.Items(i) = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 1).Value Then
                cbox_Icc_meter.SelectedIndex = i

                If txt_Icc_addr.Text = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 2).Value Then
                    import_ok = True
                    Exit For
                Else
                    import_ok = False
                End If



            End If
        Next

        If import_ok = False Then
            cbox_Icc_meter.SelectedIndex = 0
        End If



        row = row + 1
        '------------------------------------------------------------------------------------
        '------------------------------------------------------------------------------------
        'Instrument
        'xlSheet.Cells(row, col) = "VIN"
        'title_set()
        Tab_Set.SelectedIndex = 1
        row = row + 1
        txt_vin_name1.Text = xlSheet.Range(ConvertToLetter(col) & row).Value

        For i = 0 To cbox_vin.Items.Count - 1
            If cbox_vin.Items(i) = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 1).Value Then
                cbox_vin.SelectedIndex = i

                import_ok = True
                Exit For
            End If
        Next

        If import_ok = False Then
            cbox_vin.SelectedIndex = 0
        End If
        cbox_vin_ch.SelectedItem = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 2).Value
        num_VIN_OCP.Value = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 3).Value
        cbox_vin_daq.SelectedItem = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 4).Value

        row = row + 1
        'xlSheet.Cells(row, col) = "Sense Vin"
        'title_set()
        row = row + 1
        check_vin_sense.Checked = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 0).Value
        num_vin_sense.Value = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 1).Value
        num_vin_max.Value = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 2).Value
        row = row + 1


        '------------------------------------------------------------------------------------
        'xlSheet.Cells(row, col) = "IIN"
        'title_set()
        row = row + 1
        txt_iin_name1.Text = xlSheet.Range(ConvertToLetter(col) & row).Value
        If xlSheet.Range(ConvertToLetter(col) & row).Offset(, 1).Value = "TRUE" Then
            rbtn_meter_iin.Checked = True
            rbtn_board_iin.Checked = False
            rbtn_Iin_PW.Checked = False
        ElseIf Main.data_meas.Rows.Count > 0 Then

            rbtn_board_iin.Checked = True
            rbtn_meter_iin.Checked = False
            rbtn_Iin_PW.Checked = False
        Else
            rbtn_board_iin.Checked = False
            rbtn_meter_iin.Checked = False
            rbtn_Iin_PW.Checked = True
        End If

        For i = 0 To cbox_IIN_meter.Items.Count - 1
            If cbox_IIN_meter.Items(i) = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 2).Value Then
                cbox_IIN_meter.SelectedIndex = i

                If txt_IIN_addr.Text = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 3).Value Then
                    import_ok = True
                    Exit For
                Else
                    import_ok = False
                End If
            End If
        Next

        If import_ok = False Then
            cbox_IIN_meter.SelectedIndex = 0
        End If

        check_iin.Checked = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 4).Value
        cbox_IIN_relay.SelectedItem = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 5).Value
        num_iin_change.Value = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 6).Value
        rbtn_Iin_PW.Checked = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 7).Value
        rbtn_iin_current_measure.Checked = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 8).Value
        row = row + 1

        '------------------------------------------------------------------------------------
        'xlSheet.Cells(row, col) = "Vout"
        'title_set()
        row = row + 1
        txt_vout_name1.Text = xlSheet.Range(ConvertToLetter(col) & row).Value
        cbox_vout_daq.SelectedItem = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 1).Value
        cbox_vout1_daq.SelectedItem = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 2).Value
        row = row + 1
        'xlSheet.Cells(row, col) = "Check Vout"
        'title_set()
        row = row + 1
        check_shutdown.Checked = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 0).Value
        num_Vout_error.Value = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 1).Value
        row = row + 1
        '------------------------------------------------------------------------------------
        'xlSheet.Cells(row, col) = "IOUT"
        'title_set()
        row = row + 1
        txt_iout_name1.Text = xlSheet.Range(ConvertToLetter(col) & row).Value

        rbtn_meter_iout.Checked = False
        rbtn_board_iout.Checked = False
        rbtn_iout_load.Checked = False


        If xlSheet.Range(ConvertToLetter(col) & row).Offset(, 1).Value = "TRUE" Then
            rbtn_meter_iout.Checked = True
        End If



        For i = 0 To cbox_Iout_meter.Items.Count - 1
            If cbox_Iout_meter.Items(i) = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 2).Value Then
                cbox_Iout_meter.SelectedIndex = i

                If txt_Iout_addr.Text = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 3).Value Then
                    import_ok = True
                    Exit For
                Else
                    import_ok = False
                End If
            End If
        Next

        If import_ok = False Then
            cbox_Iout_meter.SelectedIndex = 0
        End If

        check_iout.Checked = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 4).Value
        cbox_Iout_relay.SelectedItem = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 5).Value
        num_iout_change.Value = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 6).Value

        If xlSheet.Range(ConvertToLetter(col) & row).Offset(, 7).Value = "TRUE" Then
            rbtn_board_iout.Checked = True
        End If
        cbox_board_buck.SelectedItem = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 8).Value
        If rbtn_meter_iout.Checked = False And rbtn_board_iout.Checked = False Then
            rbtn_iout_load.Checked = True
        ElseIf rbtn_meter_iout.Checked = True And cbox_Iout_meter.SelectedItem = no_device Then
            rbtn_iout_load.Checked = True
        End If

        rbtn_iout_load.Checked = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 9).Value
        rbtn_iout_current_measure.Checked = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 10).Value


        row = row + 1
        '------------------------------------------------------------------------------------
        'xlSheet.Cells(row, col) = "DC Load"
        'title_set()
        row = row + 1
        'xlSheet.Cells(row, col) = "Channel"
        check_IOUT_ch1.Checked = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 1).Value
        check_IOUT_ch2.Checked = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 2).Value
        check_IOUT_ch3.Checked = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 3).Value
        check_IOUT_ch4.Checked = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 4).Value
        row = row + 1

        'xlSheet.Cells(row, col) = "Delay"
        num_delay.Value = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 1).Value
        cbox_delay_unit.SelectedItem = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 2).Value
        'xlSheet.Cells(row, col + 3) = "Iout(A)  >"
        num_iout_delay.Value = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 4).Value
        row = row + 1

        '------------------------------------------------------------------------------------
        'xlSheet.Cells(row, col) = "DAQ"
        'title_set()
        row = row + 1
        ' xlSheet.Cells(row, col) = "Numbers of Trigger:"
        num_data_count.Value = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 1).Value
        row = row + 1
        'xlSheet.Cells(row, col) = "Resolution:"
        cbox_data_resolution.SelectedItem = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 1).Value
        row = row + 1


        '------------------------------------------------------------------------------------
        ' instrument2 
        Tab_Set.SelectedIndex = 8
        row = row + 1
        TextBox23.Text = xlSheet.Range(ConvertToLetter(col) & row).Value

        For i = 0 To cbox_vin2.Items.Count - 1
            If cbox_vin2.Items(i) = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 1).Value Then
                cbox_vin2.SelectedIndex = i

                import_ok = True
                Exit For
            End If
        Next

        If import_ok = False Then
            cbox_vin2.SelectedIndex = 0
        End If
        cbox_vin_ch2.SelectedItem = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 2).Value
        num_VIN_OCP2.Value = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 3).Value
        cbox_vin_daq2.SelectedItem = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 4).Value

        row = row + 1
        'xlSheet.Cells(row, col) = "Sense Vin"
        'title_set()
        row = row + 1
        check_vin_sense2.Checked = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 0).Value
        num_vin_sense2.Value = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 1).Value
        num_vin_max2.Value = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 2).Value
        row = row + 1

        '------------------------------------------------------------------------------------
        'xlSheet.Cells(row, col) = "IIN"
        'title_set()
        row = row + 1
        TextBox21.Text = xlSheet.Range(ConvertToLetter(col) & row).Value
        If xlSheet.Range(ConvertToLetter(col) & row).Offset(, 1).Value = "TRUE" Then
            rbtn_meter_iin2.Checked = True
            rbtn_board_iin2.Checked = False
            rbtn_Iin_PW2.Checked = False
        ElseIf Main.data_meas.Rows.Count > 0 Then
            rbtn_board_iin2.Checked = True
            rbtn_meter_iin2.Checked = False
            rbtn_Iin_PW2.Checked = False
        Else
            rbtn_board_iin2.Checked = False
            rbtn_meter_iin2.Checked = False
            rbtn_Iin_PW2.Checked = True
        End If

        For i = 0 To cbox_IIN_meter2.Items.Count - 1
            If cbox_IIN_meter2.Items(i) = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 2).Value Then
                cbox_IIN_meter2.SelectedIndex = i

                If txt_IIN_addr2.Text = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 3).Value Then
                    import_ok = True
                    Exit For
                Else
                    import_ok = False
                End If
            End If
        Next

        If import_ok = False Then
            cbox_IIN_meter2.SelectedIndex = 0
        End If

        check_iin2.Checked = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 4).Value
        cbox_IIN_relay2.SelectedItem = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 5).Value
        num_iin_change2.Value = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 6).Value
        rbtn_Iin_PW2.Checked = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 7).Value
        rbtn_iin_current_measure2.Checked = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 8).Value
        row = row + 1

        '------------------------------------------------------------------------------------
        'xlSheet.Cells(row, col) = "Vout"
        'title_set()
        row = row + 1
        TextBox16.Text = xlSheet.Range(ConvertToLetter(col) & row).Value
        cbox_vout_daq2.SelectedItem = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 1).Value
        cbox_vout1_daq2.SelectedItem = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 2).Value
        row = row + 1
        'xlSheet.Cells(row, col) = "Check Vout"
        'title_set()
        row = row + 1
        check_shutdown2.Checked = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 0).Value
        num_Vout_error2.Value = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 1).Value
        row = row + 1
        '------------------------------------------------------------------------------------
        'xlSheet.Cells(row, col) = "IOUT"
        'title_set()
        row = row + 1
        TextBox19.Text = xlSheet.Range(ConvertToLetter(col) & row).Value

        rbtn_meter_iout2.Checked = False
        rbtn_board_iout2.Checked = False
        rbtn_iout_load2.Checked = False


        If xlSheet.Range(ConvertToLetter(col) & row).Offset(, 1).Value = "TRUE" Then
            rbtn_meter_iout2.Checked = True
        End If

        For i = 0 To cbox_Iout_meter2.Items.Count - 1
            If cbox_Iout_meter2.Items(i) = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 2).Value Then
                cbox_Iout_meter2.SelectedIndex = i

                If txt_Iout_addr2.Text = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 3).Value Then
                    import_ok = True
                    Exit For
                Else
                    import_ok = False
                End If
            End If
        Next

        If import_ok = False Then
            cbox_Iout_meter2.SelectedIndex = 0
        End If

        check_iout2.Checked = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 4).Value
        cbox_Iout_relay2.SelectedItem = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 5).Value
        num_iout_change2.Value = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 6).Value

        If xlSheet.Range(ConvertToLetter(col) & row).Offset(, 7).Value = "TRUE" Then
            rbtn_board_iout2.Checked = True
        End If
        cbox_board_buck2.SelectedItem = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 8).Value
        If rbtn_meter_iout2.Checked = False And rbtn_board_iout2.Checked = False Then
            rbtn_iout_load2.Checked = True
        ElseIf rbtn_meter_iout2.Checked = True And cbox_Iout_meter2.SelectedItem = no_device Then
            rbtn_iout_load2.Checked = True
        End If

        rbtn_iout_load2.Checked = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 9).Value
        rbtn_iout_current_measure2.Checked = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 10).Value
        row = row + 1
        '------------------------------------------------------------------------------------
        'xlSheet.Cells(row, col) = "DC Load"
        'title_set()
        row = row + 1
        'xlSheet.Cells(row, col) = "Channel"
        check_IOUT_ch12.Checked = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 1).Value
        check_IOUT_ch22.Checked = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 2).Value
        check_IOUT_ch32.Checked = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 3).Value
        check_IOUT_ch42.Checked = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 4).Value
        row = row + 1

        'xlSheet.Cells(row, col) = "Delay"
        num_delay2.Value = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 1).Value
        cbox_delay_unit2.SelectedItem = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 2).Value
        'xlSheet.Cells(row, col + 3) = "Iout(A)  >"
        num_iout_delay2.Value = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 4).Value
        row = row + 1

        '------------------------------------------------------------------------------------
        'xlSheet.Cells(row, col) = "DAQ"
        'title_set()
        row = row + 1
        ' xlSheet.Cells(row, col) = "Numbers of Trigger:"
        num_data_count2.Value = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 1).Value
        row = row + 1
        'xlSheet.Cells(row, col) = "Resolution:"
        cbox_data_resolution2.SelectedItem = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 1).Value
        row = row + 1


        ' Meter board parameter
        row = row + 1
        num_slave_in_L.Value = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 1).Value
        num_comp_in_L.Value = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 2).Value
        num_resolution_in_L.Value = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 3).Value
        row = row + 1

        num_slave_in_M.Value = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 1).Value
        num_comp_in_M.Value = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 2).Value
        num_resolution_in_M.Value = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 3).Value
        row = row + 1

        num_slave_in_H.Value = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 1).Value
        num_comp_in_H.Value = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 2).Value
        num_resolution_in_H.Value = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 3).Value
        row = row + 1
        num_slave_in_IO.Value = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 1).Value
        row = row + 1

        num_slave_out_L.Value = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 1).Value
        num_comp_out_L.Value = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 2).Value
        num_resolution_out_L.Value = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 3).Value
        row = row + 1

        num_slave_out_M.Value = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 1).Value
        num_comp_out_M.Value = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 2).Value
        num_resolution_out_M.Value = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 3).Value
        row = row + 1

        num_slave_out_H.Value = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 1).Value
        num_comp_out_H.Value = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 2).Value
        num_resolution_out_H.Value = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 3).Value
        row = row + 1

        num_slave_out_IO.Value = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 1).Value
        row = row + 1

        '' 2nd board setting
        num_slave_in_L2.Value = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 1).Value
        num_comp_in_L2.Value = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 2).Value
        num_resolution_in_L2.Value = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 3).Value
        row = row + 1

        num_slave_in_M2.Value = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 1).Value
        num_comp_in_M2.Value = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 2).Value
        num_resolution_in_M2.Value = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 3).Value
        row = row + 1

        num_slave_in_H2.Value = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 1).Value
        num_comp_in_H2.Value = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 2).Value
        num_resolution_in_H2.Value = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 3).Value
        row = row + 1
        num_slave_in_IO2.Value = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 1).Value
        row = row + 1

        num_slave_out_L2.Value = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 1).Value
        num_comp_out_L2.Value = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 2).Value
        num_resolution_out_L2.Value = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 3).Value
        row = row + 1

        num_slave_out_M2.Value = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 1).Value
        num_comp_out_M2.Value = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 2).Value
        num_resolution_out_M2.Value = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 3).Value
        row = row + 1

        num_slave_out_H2.Value = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 1).Value
        num_comp_out_H2.Value = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 2).Value
        num_resolution_out_H2.Value = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 3).Value
        row = row + 1

        num_slave_out_IO2.Value = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 1).Value
        row = row + 1

        num_meter_count.Value = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 1).Value
        row = row + 1
        '------------------------------------------------------------------------------------
        ''------------------------------------------------------------------------------------
        'Scope
        'Step4

        Tab_Set.SelectedIndex = 2
        row = row + 1
        txt_scope_vin.Text = xlSheet.Range(ConvertToLetter(col) & row).Value
        cbox_channel_vin.SelectedItem = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 1).Value
        cbox_coupling_vin.SelectedItem = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 2).Value
        num_offset_vin.Value = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 3).Value
        num_position_vin.Value = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 4).Value
        cbox_BW_vin.SelectedItem = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 5).Value
        num_vin_scale.Value = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 6).Value
        check_scope_vin.Checked = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 7).Value
        row = row + 1


        txt_scope_iout.Text = xlSheet.Range(ConvertToLetter(col) & row).Value
        cbox_channel_iout.SelectedItem = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 1).Value
        cbox_coupling_iout.SelectedItem = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 2).Value
        num_offset_iout.Value = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 3).Value
        num_position_iout.Value = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 4).Value
        cbox_BW_iout.SelectedItem = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 5).Value
        check_scope_iout.Checked = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 6).Value
        row = row + 1

        txt_scope_lx.Text = xlSheet.Range(ConvertToLetter(col) & row).Value
        cbox_channel_lx.SelectedItem = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 1).Value
        cbox_coupling_lx.SelectedItem = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 2).Value
        num_offset_lx.Value = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 3).Value
        num_position_lx.Value = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 4).Value
        cbox_BW_lx.SelectedItem = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 5).Value
        num_lx_scale.Value = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 6).Value
        rbtn_manual_lx.Checked = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 7).Value
        num_scale_lx.Value = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 8).Value

        row = row + 1

        txt_scope_vout.Text = xlSheet.Range(ConvertToLetter(col) & row).Value
        cbox_channel_vout.SelectedItem = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 1).Value
        cbox_coupling_vout.SelectedItem = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 2).Value
        check_offset_vout.Checked = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 3).Value
        num_position_vout.Value = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 4).Value
        cbox_BW_vout.SelectedItem = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 5).Value

        If xlSheet.Range(ConvertToLetter(col) & row).Offset(, 6).Value = False Then
            rbtn_auto_vout.Checked = False
            rbtn_manual_vout.Checked = True
        Else
            rbtn_auto_vout.Checked = True
            rbtn_manual_vout.Checked = False
        End If
        num_vout_auto.Value = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 7).Value
        Check_fixed.Checked = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 8).Value
        num_vout_DEM.Value = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 9).Value
        num_vout_CCM.Value = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 10).Value
        row = row + 1


        txt_scope_lx2.Text = xlSheet.Range(ConvertToLetter(col) & row).Value
        cbox_channel_lx2.SelectedItem = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 1).Value
        cbox_coupling_lx2.SelectedItem = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 2).Value
        num_offset_lx2.Value = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 3).Value
        num_position_lx2.Value = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 4).Value
        cbox_BW_lx2.SelectedItem = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 5).Value
        row = row + 1


        txt_scope_vout2.Text = xlSheet.Range(ConvertToLetter(col) & row).Value
        cbox_channel_vout2.SelectedItem = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 1).Value
        cbox_coupling_vout2.SelectedItem = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 2).Value
        check_offset_vout2.Checked = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 3).Value
        num_position_vout2.Value = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 4).Value
        cbox_BW_vout2.SelectedItem = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 5).Value
        row = row + 1


        '"Time Setting"   
        row = row + 1

        num_RL.Value = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 1).Value
        row = row + 1

        num_points.Value = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 1).Value
        row = row + 1

        num_location.Value = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 1).Value
        row = row + 1

        num_counts_DEM.Value = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 1).Value
        num_counts_CCM.Value = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 2).Value

        row = row + 1

        num_wave.Value = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 1).Value
        row = row + 1


        rbtn_vin_trigger.Checked = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 1).Value
        num_vin_trigger.Value = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 2).Value
        rbtn_auto_trigger.Checked = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 3).Value
        row = row + 1
        ' "Measurement"
        row = row + 1

        txt_meas1_ch.Text = xlSheet.Range(ConvertToLetter(col) & row).Value
        txt_type1.Text = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 1).Value
        txt_meas1.Text = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 2).Value
        row = row + 1

        txt_meas2_ch.Text = xlSheet.Range(ConvertToLetter(col) & row).Value
        txt_type2.Text = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 1).Value
        txt_meas2.Text = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 2).Value
        row = row + 1

        txt_meas3_ch.Text = xlSheet.Range(ConvertToLetter(col) & row).Value
        txt_type3.Text = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 1).Value
        txt_meas3.Text = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 2).Value
        row = row + 1

        txt_meas4_ch.Text = xlSheet.Range(ConvertToLetter(col) & row).Value
        txt_type4.Text = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 1).Value
        txt_meas4.Text = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 2).Value
        row = row + 1

        txt_meas5_ch.Text = xlSheet.Range(ConvertToLetter(col) & row).Value
        txt_type5.Text = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 1).Value
        txt_meas5.Text = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 2).Value
        row = row + 1

        txt_meas6_ch.Text = xlSheet.Range(ConvertToLetter(col) & row).Value
        txt_type6.Text = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 1).Value
        txt_meas6.Text = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 2).Value
        row = row + 1
        '------------------------------------------------------------------------------------
        Tab_Set.SelectedIndex = 3
        'Test Conditions page

        'xlSheet.Cells(row, col) = "Test Init"
        'title_set()
        row = row + 1

        'xlSheet.Cells(row, col) = "Fsw (Hz)"


        For i = 0 To clist_fs.Items.Count - 1
            clist_fs.SetItemChecked(i, xlSheet.Range(ConvertToLetter(col) & row).Offset(, 1 + i).Value)
        Next

        row = row + 1

        'xlSheet.Cells(row, col) = "VOUT (V)"

        For i = 0 To clist_vout.Items.Count - 1
            clist_vout.SetItemChecked(i, xlSheet.Range(ConvertToLetter(col) & row).Offset(, 1 + i).Value)
        Next

        row = row + 1

        data_test_import(data_vin, last_col)

        data_list()

        '------------------------------------------------------------------------------------

        'xlSheet.Cells(row, col) = "Stability"
        'title_set()
        row = row + 1


        'check_Force_CCM.Text = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 0).Value
        check_Force_CCM.Checked = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 1).Value

        row = row + 1

        ' xlSheet.Cells(row, col) = "Fs_leak_0A"
        num_fs_leak.Value = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 1).Value

        row = row + 1

        'check_IOB.Text = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 0).Value
        check_IOB.Checked = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 1).Value
        'xlSheet.Cells(row, col + 2) = "Range (A)"
        num_IOB_Range.Value = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 3).Value
        'xlSheet.Cells(row, col + 4) = "Step (A)"
        num_IOB_step.Value = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 5).Value

        row = row + 1

        'check_iout_up.Text = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 0).Value
        check_iout_up.Checked = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 1).Value
        row = row + 1


        'xlSheet.Cells(row, col) = "Test Conditions"
        'title_set()
        row = row + 1

        'data_test_set(data_set)
        data_test_import(data_set, last_col)

        data_test_import(data_iout, last_col)

        data_set_list()
        '------------------------------------------------------------------------------------
        'xlSheet.Cells(row, col) = "Efficiency and Load Regulation"
        'title_set()
        row = row + 1
        data_test_import(data_eff_iout, last_col)

        '------------------------------------------------------------------------------------
        'xlSheet.Cells(row, col) = "Jitter"
        'title_set()
        row = row + 1
        data_test_import(data_jitter_iout, last_col)

        '------------------------------------------------------------------------------------

        '------------------------------------------------------------------------------------
        'Test setup Page
        '------------------------------------------------------------------------------------
        'Test setup Page
        'xlSheet.Cells(row, col) = "Stability Set"
        'title_set()
        row = row + 1
        Tab_Set.SelectedIndex = 4
        ' xlSheet.Cells(row, col) = " Auto Scanning"
        If xlSheet.Range(ConvertToLetter(col) & row).Offset(, 1).Value = rbtn_auto_all.Text Then
            rbtn_auto_all.Checked = True
        Else
            rbtn_auto_DEM.Checked = True
        End If

        row = row + 1

        ' xlSheet.Cells(row, col) = "Add Cursors"
        check_cursors.Checked = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 1).Value
        row = row + 1


        num_ton_vin.Value = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 1).Value
        rbtn_ton_cal.Checked = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 2).Value
        num_ton_cal.Value = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 3).Value
        rbtn_ton_val.Checked = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 4).Value
        num_ton_val.Value = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 5).Value
        row = row + 1

        num_toff_vin.Value = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 1).Value
        rbtn_toff_cal.Checked = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 2).Value
        num_toff_cal.Value = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 3).Value
        rbtn_toff_val.Checked = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 4).Value
        num_toff_val.Value = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 5).Value
        row = row + 1

        ' " Update: "
        If xlSheet.Range(ConvertToLetter(col) & row).Offset(, 1).Value = rbtn_freq_high.Text Then
            rbtn_freq_high.Checked = True
        Else
            rbtn_freq_low.Checked = True
        End If
        row = row + 1

        'xlSheet.Cells(row, col) = "Error"
        num_delay_error.Value = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 1).Value
        check_error_pic.Checked = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 2).Value

        row = row + 1


        'xlSheet.Cells(row, col) = "Capture All"
        check_stability_pic.Checked = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 1).Value

        row = row + 1


        'xlSheet.Cells(row, col) = "Efficiency Set"
        'title_set()
        row = row + 1

        ' xlSheet.Cells(row, col) = "Auto"
        rbtn_iin_auto.Checked = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 1).Value
        rbtn_iin_manual.Checked = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 2).Value
        row = row + 1




        ' xlSheet.Cells(row, col) = "IOUT Step"
        num_iin_step.Value = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 1).Value
        num_iout_auto_stop.Value = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 2).Value
        ' xlSheet.Cells(row, col + 2) = "mA"

        row = row + 1

        data_test_import(data_eff, last_col)



        'xlSheet.Cells(row, col) = "Minimum range"
        cbox_meter_mini.SelectedItem = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 1).Value

        row = row + 1

        'Jitter Page
        'xlSheet.Cells(row, col) = "Jitter Set"
        'title_set()
        row = row + 1


        ' xlSheet.Cells(row, col) = "PERSistence"
        check_persistence.Checked = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 1).Value
        num_counts_Jitter.Value = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 2).Value
        row = row + 1
        'xlSheet.Cells(row, col) = "FastAcq"
        check_fastAcq.Checked = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 1).Value
        num_FastAcq.Value = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 2).Value
        row = row + 1



        '------------------------------------------------------------------------------------

        'xlSheet.Cells(row, col) = "Line Regulation"
        'title_set()
        row = row + 1
        'xlSheet.Cells(row, col) = "RUN Mode"
        Tab_Set.SelectedIndex = 5
        If xlSheet.Range(ConvertToLetter(col) & row).Offset(, 1).Value = True Then
            rbtn_lineR_test1.Checked = True
            rbtn_lineR_test2.Checked = False
        Else
            rbtn_lineR_test1.Checked = False
            rbtn_lineR_test2.Checked = True
        End If

        row = row + 1
        data_test_import(data_lineR_iout, last_col)
        data_test_import(data_lineR_vin, last_col)


        'xlSheet.Cells(row, col) = "LineR Set"
        check_lineR_scope.Checked = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 1).Value

        row = row + 2
        check_DUT2.Checked = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 1).Value


        '------------------------------------------------------------------------------------

        import_now = False
        Tab_Set.SelectedIndex = 6
        import_now = True

        FinalReleaseComObject(xlSheet)
        xlSheet = Nothing


    End Function

    Function report_init() As Integer

        '-----------------------------------------------------------------------------------
        'PartI data 
        '-----------------------------------------------------------------------------------
        note_string = ""
        note_display = True
        Information.information_run("Initial Report", note_run)

        excel_open()

        If check_LineR.Checked = True Then
            If TA_Test_num = 0 Then
                If check_lineR_scope.Checked = True Then
                    sheet_init(txt_data_sheet.Text)
                End If
                sheet_init(txt_LineR_sheet.Text)

                If DUT2_en Then : sheet_init(txt_LineR_sheet.Text & add_dut2) : End If
            End If
            test_report_init(Line_Regulation, 0)
            If DUT2_en Then : test_report_init(Line_Regulation, 1) : End If
        End If



        '-----------------------------------------------------------------------------------
        If check_loadR.Checked = True Then
            If TA_Test_num = 0 Then
                sheet_init(txt_LoadR_sheet.Text)
                If DUT2_en Then : sheet_init(txt_LoadR_sheet.Text & add_dut2) : End If
            End If
            'Load Regulation
            test_report_init(Load_Regulation)
            If DUT2_en Then : test_report_init(Load_Regulation, 1) : End If

        End If

        '-----------------------------------------------------------------------------------

        If check_Efficiency.Checked = True Then
            If TA_Test_num = 0 Then
                sheet_init(txt_eff_sheet.Text)
                iin_range_report_info()
                If DUT2_en Then : sheet_init(txt_eff_sheet.Text & add_dut2) : End If
                iin_range_report_info()
            End If
            'Efficiency
            test_report_init(Efficiency)
            If DUT2_en Then : test_report_init(Efficiency, 1) : End If
        End If
        '-----------------------------------------------------------------------------------

        If check_jitter.Checked = True Then
            If TA_Test_num = 0 Then
                sheet_init(txt_jitter_sheet.Text)
                Jitter_folder = folderPath & "\Jitter_" & DateTime.Now.ToString("MMdd") & "_" & DateTime.Now.ToString("HHmmss")
                My.Computer.FileSystem.CreateDirectory(Jitter_folder)

                If DUT2_en Then
                    sheet_init(txt_jitter_sheet.Text & add_dut2)
                    Jitter_folder2 = folderPath & "\Jitter_" & add_dut2 & "_" & DateTime.Now.ToString("MMdd") & "_" & DateTime.Now.ToString("HHmmss")
                    My.Computer.FileSystem.CreateDirectory(Jitter_folder2)
                End If
            End If
            test_report_init(Jitter)
            If DUT2_en Then : test_report_init(Jitter, 1) : End If
        End If

        '-----------------------------------------------------------------------------------

        If check_stability.Checked = True Then

            If TA_Test_num = 0 Then

                'Bata
                If (check_stability_pic.Checked = True) Then
                    'sheet_init(txt_beta_sheet.Text, False)
                    Beta_folder = folderPath & "\Beta_" & DateTime.Now.ToString("MMdd") & "_" & DateTime.Now.ToString("HHmmss")
                    My.Computer.FileSystem.CreateDirectory(Beta_folder)

                    If DUT2_en Then
                        Beta_folder2 = folderPath & "\Beta_" & add_dut2 & "_" & DateTime.Now.ToString("MMdd") & "_" & DateTime.Now.ToString("HHmmss")
                        My.Computer.FileSystem.CreateDirectory(Beta_folder2)
                    End If

                End If
                '----------------------------------------------------------------------------------
                'Error
                sheet_init(txt_error_sheet.Text)
                If DUT2_en Then : sheet_init(txt_error_sheet.Text & add_dut2) : End If
                Error_folder = ""
                '----------------------------------------------------------------------------------
                'Data
                sheet_init(txt_stability_sheet.Text)
                If DUT2_en Then : sheet_init(txt_stability_sheet.Text & add_dut2) : End If
                '----------------------------------------------------------------------------------

            End If
            test_report_init(Stability)
            If DUT2_en Then : test_report_init(Stability, 1) : End If
        End If

        excel_close()

        GC.Collect()
        GC.WaitForPendingFinalizers()

        note_display = False
    End Function

    Function test_report_init(ByVal test_name As String, Optional ByVal dut_sel As Integer = 0) As Integer

        Dim n, f, v, i, ii, nn As Integer

        Dim TA_title, VCC_title, Fs_title, total_title As String
        Dim TA_serial, VCC_serial, Fs_serial, total_serial As String
        Dim first_row As Integer 'title
        Dim last_row As Integer 'last_parameter
        Dim start_col As Integer
        Dim col_num, row_num As Integer

        Dim chart_num As Integer = 0

        Dim stability_num As Integer = 0

        Dim row_num_temp As Integer

        Dim VCC_test As Boolean = False
        Dim iout_col, freq_col, ton_col, toff_col, vpp_col As Integer
        Dim eff_col() As String = {Vin_name, Iin_name, Vout_name, Iout_name, "Efficiency >" & num_pass_eff.Value & "%", "Loss (W)"}


        Dim set_num As Integer
        Dim copy_row As Integer

        note_string = test_name

        '--------------------------------------------------------------------------------------
        'for Efficiency
        jitter_col(0) = Vout_name
        stability_col(0) = Vout_name
        jitter_col(1) = Iout_name
        stability_col(1) = Iout_name
        jitter_col(9) = "Jitter <" & num_pass_jitter.Value & "%"
        col_num = eff_col.Length
        If (cbox_VCC.SelectedItem <> no_device) Or (cbox_VCC_daq.SelectedItem <> no_device) Then
            ReDim Preserve eff_col(col_num)
            eff_col(col_num) = Vcc_name
            col_num = col_num + 1

            If total_vcc.Length > 1 Then
                VCC_test = True
            End If
        End If

        If cbox_Icc_meter.SelectedItem <> no_device Then
            ReDim Preserve eff_col(col_num)
            eff_col(col_num) = Icc_name
            col_num = col_num + 1
        End If

        If ((cbox_VCC.SelectedItem <> no_device) Or (cbox_VCC_daq.SelectedItem <> no_device)) And (txt_Icc_addr.Text <> "") Then
            total_Eff = True
        Else
            total_Eff = False
        End If

        If total_Eff = True Then
            ReDim Preserve eff_col(col_num)
            eff_col(col_num) = "Total Eff (%)"
            col_num = col_num + 1

        End If


        ReDim Preserve eff_col(col_num)
        eff_col(col_num) = "PASS/FAIL"

        eff_title_total = eff_col.Length

        '---------------------------------------------------------------------------------
        'for stability
        For i = 0 To data_set.Rows.Count - 1
            If data_set.Rows(i).Cells(0).Value = TA_now Then
                stability_num = i
                Exit For
            End If
        Next

        '---------------------------------------------------------------------------------
        'For line Regulation

        Dim data_Line_vin As Object

        If rbtn_lineR_test1.Checked = True Then
            data_Line_vin = data_lineR_vin
        Else
            data_Line_vin = data_vin
        End If
        '---------------------------------------------------------------------------------
        'Init
        start_col = test_col
        first_row = test_row
        col_num = 0
        row_num = 0
        last_row = 0
        ''----------------------------------------------------------------------------------
        ''Init
        ''----------------------------------------------------------------------------------
        total_serial = ""
        If TA_Test_num = 0 Then
            Jitter_pic_num = 1
            Jitter_pic_num2 = 1
        Else
            Jitter_pic_num = data_jitter_iout.Rows.Count * total_vcc.Length * total_fs.Length * total_vout.Length * data_vin.Rows.Count * TA_Test_num + 1
            Jitter_pic_num2 = data_jitter_iout.Rows.Count * total_vcc.Length * total_fs.Length * total_vout.Length * data_vin.Rows.Count * TA_Test_num + 1
        End If

        'TA Loop


        If Main.check_TA_en.Checked = False Then
            TA_title = ""
        Else
            TA_now = Main.data_Temp.Rows(TA_Test_num).Cells(0).Value
            TA_title = "TA=" & TA_now & ", "

        End If

        If (Main.data_Temp.Rows.Count > 1) Then
            TA_serial = TA_title
        Else
            TA_serial = ""
        End If

        '---------------------------------------------------------------------------------

        '---------------------------------------------------------------------------------
        For n = 0 To total_vcc.Length - 1
            ' VCC Loop



            vcc_now = total_vcc(n)

            If vcc_now <> 0 Then
                VCC_title = txt_vcc_name1.Text & "=" & vcc_now & "V, "
            Else
                VCC_title = ""
            End If

            If VCC_test = True Then

                VCC_serial = VCC_title
            Else
                VCC_serial = ""
            End If

            '----------------------------------------------------------------------------------

            For f = 0 To total_fs.Length - 1
                ' Fsw Loop


                fs_now = total_fs(f)
                If fs_now <> 0 Then
                    Fs_title = "Fsw=" & fs_now / 1000 & "kHz, "
                Else

                    Fs_title = ""
                End If

                If total_fs.Length > 1 Then

                    Fs_serial = Fs_title
                Else
                    Fs_serial = ""
                End If
                '----------------------------------------------------------------------------------
                For v = 0 To total_vout.Length - 1
                    'Vout Loop
                    System.Windows.Forms.Application.DoEvents()
                    vout_now = total_vout(v)
                    total_title = TA_title & VCC_title & Fs_title & "VOUT=" & vout_now & "V"
                    ''----------------------------------------------------------------------------------
                    'Start Test
                    Select Case test_name
                        Case Stability
                            'Vout, Vin, vcc往下移，fs, temp都往左移
                            Select Case dut_sel
                                Case 0
                                    xlSheet = xlBook.Sheets(txt_stability_sheet.Text)
                                Case 1
                                    xlSheet = xlBook.Sheets(txt_stability_sheet.Text & add_dut2)
                            End Select

                            xlSheet.Activate()
                            '----------------------------------------------------------------------------------
                            'initial
                            'Init col

                            If cbox_coupling_vout.SelectedItem = "AC" Then
                                pass_value_Max = vout_now * (num_vout_ac.Value / 100)
                                stability_col(22) = "Vpp(max) <" & pass_value_Max & "V"


                            Else
                                pass_value_Max = vout_now * (1 + num_vout_pos.Value / 100)
                                pass_value_Min = vout_now * (1 - num_vout_neg.Value / 100)

                                stability_col(23) = "Vmax(max) <" & pass_value_Max & "V"
                                stability_col(24) = "Vmin(min) >" & pass_value_Min & "V"

                            End If


                            col_num = stability_col.Length


                            start_col = test_col + chart_width + col_Space + (TA_Test_num * total_vcc.Length * total_fs.Length + n * total_fs.Length + f) * (col_num + 1)
                            row_num_temp = 0
                            '----------------------------------------------------------------------------------
                            '加了IOB，iout的數量不一定都一樣
                            If (v = 0) Then
                                first_row = test_row
                            Else
                                'row_num = 0
                                For nn = 0 To v - 1

                                    row_num = 0
                                    For i = 0 To data_vin.Rows.Count - 1
                                        set_num = TA_Test_num * total_vcc.Length * total_fs.Length * total_vout.Length * data_vin.Rows.Count + n * total_fs.Length * total_vout.Length * data_vin.Rows.Count + f * total_vout.Length * data_vin.Rows.Count + nn * data_vin.Rows.Count + i

                                        If check_iout_up.Checked = True Then
                                            row_num = row_num + (2 * (stability_row_stop(set_num) - stability_row_start(set_num)) + 1 + 2) + 1
                                        Else
                                            row_num = row_num + (stability_row_stop(set_num) - stability_row_start(set_num) + 1 + 2) + 1
                                        End If
                                    Next



                                    'Init row
                                    If (row_num) < (4 * (chart_height + row_Space) - row_Space) Then
                                        row_num_temp = row_num_temp + (4 * (chart_height + 1 + row_Space))

                                    Else
                                        row_num_temp = row_num_temp + row_num + row_Space
                                    End If


                                Next

                                first_row = test_row + row_num_temp
                                'Init row
                                'If (row_num) < (4 * (chart_height + row_Space) - row_Space) Then
                                '    first_row = test_row + v * (4 * (chart_height + 1 + row_Space))

                                'Else
                                '    first_row = test_row + row_num + v * row_Space
                                'End If



                            End If
                            '----------------------------------------------------------------------------------        




                            col = start_col
                            row = first_row




                            'Total Chart
                            chart_num = v * 4 + 1
                            '----------------------------------------------------------------------------------
                            'Chart
                            If (TA_Test_num = 0) And (n = 0) And (f = 0) Then


                                chart_col = test_col
                                chart_row = first_row

                                chart_init(Freq_Chart, "VOUT=" & vout_now & "V", "Frequency vs Load Current", Iout_title, "Frequency (kHz)", Full_load, 0, "", "", cbox_type_stability.SelectedItem)

                                chart_row = chart_row + chart_height + row_Space

                                chart_init(Ton_Chart, "VOUT=" & vout_now & "V", "Ton vs Load Current", Iout_title, "Ton (ns)", Full_load, 0, "", "", cbox_type_stability.SelectedItem)


                                chart_row = chart_row + chart_height + row_Space

                                chart_init(Toff_Chart, "VOUT=" & vout_now & "V", "Toff vs Load Current", Iout_title, "Toff (ns)", Full_load, 0, "", "", cbox_type_stability.SelectedItem)

                                chart_row = chart_row + chart_height + row_Space

                                chart_init(Vpp_Chart, "VOUT=" & vout_now & "V", "Vpp vs Load Current", Iout_title, "Vpp (mV)", Full_load, 0, "", "", cbox_type_stability.SelectedItem)


                            End If
                            '-------------------------------------------------------------------------------



                            row = first_row

                            For i = 0 To data_vin.Rows.Count - 1





                                xlSheet.Activate()

                                first_row = row



                                ReDim Preserve stability_report_row(stability_num)
                                stability_report_row(stability_num) = first_row


                                row = first_row
                                col = start_col
                                vin_now = data_vin.Rows(i).Cells(0).Value
                                chart_num = v * 4 + 1
                                total_title = TA_title & VCC_title & Fs_title & "VOUT=" & vout_now & "V, VIN=" & vin_now & "V"
                                'Title
                                report_title(total_title, col, row, col_num, 1, data_title_color)
                                row = row + 1


                                If i = 0 Then

                                    copy_row = row
                                    For nn = 0 To stability_col.Length - 1
                                        report_title(stability_col(nn), col, row, 1, 1, data_title_color)
                                        xlSheet.Columns(col).AutoFit()
                                        col = col + 1

                                    Next


                                Else
                                    xlSheet.Range(ConvertToLetter(start_col) & copy_row & ":" & ConvertToLetter(start_col + stability_col.Length - 1) & copy_row).Copy()
                                    xlSheet.Range(ConvertToLetter(start_col) & row).Select()
                                    xlSheet.Paste()
                                End If



                                vin_now = data_vin.Rows(i).Cells(0).Value


                                row = row + 1
                                'stability_parameter(stability_num)

                                'stability_row_num(v * data_vin.Rows.Count + i) = data_test.Rows.Count + 2



                                'X
                                For ii = stability_row_start(stability_num) To stability_row_stop(stability_num)
                                    'Next 'stabitity
                                    last_row = row
                                    row = row + 1
                                Next 'iout

                                If check_iout_up.Checked = True Then

                                    For ii = (stability_row_stop(stability_num) - 1) To stability_row_start(stability_num) Step -1


                                        last_row = row
                                        row = row + 1


                                    Next 'iout




                                End If



                                col = start_col

                                For nn = 0 To stability_col.Length - 1

                                    Select Case stability_col(nn)

                                        Case Vout_name


                                        Case Iout_name


                                            iout_col = col



                                        Case "Frequency(kHz)"
                                            freq_col = col
                                        Case "Ton(ns)"
                                            ton_col = col
                                        Case "Toff(ns)"
                                            toff_col = col
                                        Case "Vpp(mV)"
                                            vpp_col = col
                                    End Select
                                    col = col + 1
                                Next 'stabitity
                                stability_num = stability_num + 1
                                '----------------------------------------------------------------------------------
                                'Add Line
                                report_Group(start_col, first_row, col_num, last_row - first_row + 1)
                                '-------------------------------------------------------------------------------
                                'Add Serial 
                                chart_row_start = first_row + 2
                                chart_row_stop = last_row
                                total_serial = TA_serial & VCC_serial & Fs_serial & "VIN=" & vin_now & "V,Fsw"
                                chart_add_series(txt_stability_sheet.Text, Freq_Chart, chart_num, total_serial, iout_col, freq_col, False)
                                freq_col = freq_col + 1
                                total_serial = TA_serial & VCC_serial & Fs_serial & "VIN=" & vin_now & "V,mean"
                                chart_add_series(txt_stability_sheet.Text, Freq_Chart, chart_num, total_serial, iout_col, freq_col, False)
                                freq_col = freq_col + 1
                                total_serial = TA_serial & VCC_serial & Fs_serial & "VIN=" & vin_now & "V,min"
                                chart_add_series(txt_stability_sheet.Text, Freq_Chart, chart_num, total_serial, iout_col, freq_col, False)
                                freq_col = freq_col + 1
                                total_serial = TA_serial & VCC_serial & Fs_serial & "VIN=" & vin_now & "V,max"
                                chart_add_series(txt_stability_sheet.Text, Freq_Chart, chart_num, total_serial, iout_col, freq_col, False)
                                freq_col = freq_col + 1
                                total_serial = TA_serial & VCC_serial & Fs_serial & "VIN=" & vin_now & "V,update"
                                chart_add_series(txt_stability_sheet.Text, Freq_Chart, chart_num, total_serial, iout_col, freq_col, False)
                                chart_num = chart_num + 1
                                '-------------------------------------------------------------------------------------------------------
                                total_serial = TA_serial & VCC_serial & Fs_serial & "VIN=" & vin_now & "V,Ton"
                                chart_add_series(txt_stability_sheet.Text, Ton_Chart, chart_num, total_serial, iout_col, ton_col, False)
                                ton_col = ton_col + 1
                                total_serial = TA_serial & VCC_serial & Fs_serial & "VIN=" & vin_now & "V,mean"
                                chart_add_series(txt_stability_sheet.Text, Ton_Chart, chart_num, total_serial, iout_col, ton_col, False)
                                ton_col = ton_col + 1
                                total_serial = TA_serial & VCC_serial & Fs_serial & "VIN=" & vin_now & "V,min"
                                chart_add_series(txt_stability_sheet.Text, Ton_Chart, chart_num, total_serial, iout_col, ton_col, False)
                                ton_col = ton_col + 1


                                total_serial = TA_serial & VCC_serial & Fs_serial & "VIN=" & vin_now & "V,max"
                                chart_add_series(txt_stability_sheet.Text, Ton_Chart, chart_num, total_serial, iout_col, ton_col, False)
                                ton_col = ton_col + 1


                                total_serial = TA_serial & VCC_serial & Fs_serial & "VIN=" & vin_now & "V,update"
                                chart_add_series(txt_stability_sheet.Text, Ton_Chart, chart_num, total_serial, iout_col, ton_col, False)

                                chart_num = chart_num + 1
                                '----------------------------------------------------------------------------------
                                total_serial = TA_serial & VCC_serial & Fs_serial & "VIN=" & vin_now & "V,Toff"
                                chart_add_series(txt_stability_sheet.Text, Toff_Chart, chart_num, total_serial, iout_col, toff_col, False)
                                toff_col = toff_col + 1


                                total_serial = TA_serial & VCC_serial & Fs_serial & "VIN=" & vin_now & "V,mean"
                                chart_add_series(txt_stability_sheet.Text, Toff_Chart, chart_num, total_serial, iout_col, toff_col, False)
                                toff_col = toff_col + 1


                                total_serial = TA_serial & VCC_serial & Fs_serial & "VIN=" & vin_now & "V,min"
                                chart_add_series(txt_stability_sheet.Text, Toff_Chart, chart_num, total_serial, iout_col, toff_col, False)
                                toff_col = toff_col + 1


                                total_serial = TA_serial & VCC_serial & Fs_serial & "VIN=" & vin_now & "V,max"
                                chart_add_series(txt_stability_sheet.Text, Toff_Chart, chart_num, total_serial, iout_col, toff_col, False)
                                toff_col = toff_col + 1


                                total_serial = TA_serial & VCC_serial & Fs_serial & "VIN=" & vin_now & "V,update"
                                chart_add_series(txt_stability_sheet.Text, Toff_Chart, chart_num, total_serial, iout_col, toff_col, False)

                                chart_num = chart_num + 1
                                '----------------------------------------------------------------------------------
                                '----------------------------------------------------------------------------------
                                total_serial = TA_serial & VCC_serial & Fs_serial & "VIN=" & vin_now & "V,vpp"
                                chart_add_series(txt_stability_sheet.Text, Vpp_Chart, chart_num, total_serial, iout_col, vpp_col, False)
                                vpp_col = vpp_col + 1


                                total_serial = TA_serial & VCC_serial & Fs_serial & "VIN=" & vin_now & "V,mean"
                                chart_add_series(txt_stability_sheet.Text, Vpp_Chart, chart_num, total_serial, iout_col, vpp_col, False)
                                vpp_col = vpp_col + 1

                                total_serial = TA_serial & VCC_serial & Fs_serial & "VIN=" & vin_now & "V,min"
                                chart_add_series(txt_stability_sheet.Text, Vpp_Chart, chart_num, total_serial, iout_col, vpp_col, False)
                                vpp_col = vpp_col + 1

                                total_serial = TA_serial & VCC_serial & Fs_serial & "VIN=" & vin_now & "V,max"
                                chart_add_series(txt_stability_sheet.Text, Vpp_Chart, chart_num, total_serial, iout_col, vpp_col, False)
                                vpp_col = vpp_col + 1

                                '----------------------------------------------------------------------------------
                                col = col + 1
                                row = row + 1
                            Next  'vin
                            last_row = last_row + 2
                            '----------------------------------------------------------------------------------
                        Case Jitter
                            Select Case dut_sel
                                Case 0
                                    xlSheet = xlBook.Sheets(txt_jitter_sheet.Text)
                                Case 1
                                    xlSheet = xlBook.Sheets(txt_jitter_sheet.Text & add_dut2)
                            End Select
                            xlSheet.Activate()
                            '----------------------------------------------------------------------------------
                            'initial
                            'Init col
                            col_num = jitter_col.Length
                            row_num = data_jitter_iout.Rows.Count + 2
                            start_col = test_col + data_jitter_iout.Rows.Count * (TA_num + 1) * total_vcc.Length * total_fs.Length * (pic_width + 1) + col_Space + (TA_Test_num * total_vcc.Length * total_fs.Length + n * total_fs.Length + f) * (col_num + 1) - 1
                            'Init row
                            If (check_fastAcq.Checked = True) Then

                                If row_num < (2 * pic_height + 1) Then
                                    first_row = test_row + v * (data_vin.Rows.Count * (2 * pic_height + 1 + 1) + row_Space)
                                Else
                                    first_row = test_row + v * (data_vin.Rows.Count * (row_num + 1) + row_Space)
                                End If
                            Else
                                If row_num < (pic_height + 1) Then
                                    first_row = test_row + v * (data_vin.Rows.Count * (pic_height + 1 + 1) + row_Space)
                                Else
                                    first_row = test_row + v * (data_vin.Rows.Count * (row_num + 1) + row_Space)
                                End If
                            End If
                            col = start_col
                            row = first_row
                            'pass_value_Max = num_pass_jitter.Value
                            '-------------------------------------------------------------------------------
                            row = first_row
                            For i = 0 To data_vin.Rows.Count - 1
                                If (check_fastAcq.Checked = True) Then
                                    If row_num < (2 * pic_height + 1) Then
                                        first_row = test_row + v * (data_vin.Rows.Count * (2 * pic_height + 1 + 1) + row_Space) + i * (2 * pic_height + 1 + 1)
                                    Else
                                        first_row = test_row + v * (data_vin.Rows.Count * (row_num + 1) + row_Space) + i * (row_num + 1 + 1)
                                    End If
                                Else
                                    If row_num < (pic_height + 1) Then
                                        first_row = test_row + v * (data_vin.Rows.Count * (pic_height + 1 + 1) + row_Space) + i * (pic_height + 1 + 1)
                                    Else
                                        first_row = test_row + v * (data_vin.Rows.Count * (row_num + 1) + row_Space) + i * (row_num + 1 + 1)
                                    End If
                                End If
                                row = first_row
                                col = start_col
                                vin_now = data_vin.Rows(i).Cells(0).Value
                                total_title = TA_title & VCC_title & Fs_title & "VOUT=" & vout_now & "V, VIN=" & vin_now & "V"
                                'Title
                                report_title(total_title, col, row, col_num, 1, data_title_color)
                                row = row + 1
                                For nn = 0 To jitter_col.Length - 1
                                    report_title(jitter_col(nn), col, row, 1, 1, data_title_color)
                                    xlSheet.Columns(col).AutoFit()
                                    col = col + 1
                                Next
                                vin_now = data_vin.Rows(i).Cells(0).Value
                                row = row + 1
                                'X
                                For ii = 0 To data_jitter_iout.Rows.Count - 1
                                    iout_now = data_jitter_iout.Rows(ii).Cells(0).Value


                                    If dut_sel = 1 Then
                                        ReDim Preserve jitter_pic_col2(Jitter_pic_num2)
                                        ReDim Preserve jitter_pic_row2(Jitter_pic_num2)

                                        If ii = 0 Then
                                            jitter_pic_col2(Jitter_pic_num2) = test_col + (TA_Test_num * total_vcc.Length * total_fs.Length + n * total_fs.Length + f) * data_jitter_iout.Rows.Count * (pic_width + 1)
                                            jitter_pic_row2(Jitter_pic_num2) = first_row
                                        Else
                                            jitter_pic_col2(Jitter_pic_num2) = jitter_pic_col2(Jitter_pic_num2 - 1) + 1 + pic_width
                                            jitter_pic_row2(Jitter_pic_num2) = jitter_pic_row2(Jitter_pic_num2 - 1)
                                        End If
                                        col = start_col
                                        For nn = 0 To jitter_col.Length - 1
                                            col = col + 1
                                        Next 'jitter
                                        last_row = row
                                        row = row + 1
                                        '-------------------------------------------------------------------------------
                                        'Add Picture
                                        If (check_fastAcq.Checked = True) Then
                                            pic_init(total_title & ", Iout=" & iout_now & "A", jitter_pic_col2(Jitter_pic_num2), jitter_pic_row2(Jitter_pic_num2), 2)
                                        Else
                                            pic_init(total_title & ", Iout=" & iout_now & "A", jitter_pic_col2(Jitter_pic_num2), jitter_pic_row2(Jitter_pic_num2), 1)
                                        End If

                                        Jitter_pic_num2 = Jitter_pic_num2 + 1
                                    Else
                                        ReDim Preserve jitter_pic_col(Jitter_pic_num)
                                        ReDim Preserve jitter_pic_row(Jitter_pic_num)
                                        If ii = 0 Then
                                            jitter_pic_col(Jitter_pic_num) = test_col + (TA_Test_num * total_vcc.Length * total_fs.Length + n * total_fs.Length + f) * data_jitter_iout.Rows.Count * (pic_width + 1)
                                            jitter_pic_row(Jitter_pic_num) = first_row
                                        Else
                                            jitter_pic_col(Jitter_pic_num) = jitter_pic_col(Jitter_pic_num - 1) + 1 + pic_width
                                            jitter_pic_row(Jitter_pic_num) = jitter_pic_row(Jitter_pic_num - 1)
                                        End If
                                        col = start_col
                                        For nn = 0 To jitter_col.Length - 1
                                            col = col + 1
                                        Next 'jitter
                                        last_row = row
                                        row = row + 1
                                        '-------------------------------------------------------------------------------
                                        'Add Picture
                                        If (check_fastAcq.Checked = True) Then

                                            pic_init(total_title & ", Iout=" & iout_now & "A", jitter_pic_col(Jitter_pic_num), jitter_pic_row(Jitter_pic_num), 2)
                                        Else
                                            pic_init(total_title & ", Iout=" & iout_now & "A", jitter_pic_col(Jitter_pic_num), jitter_pic_row(Jitter_pic_num), 1)
                                        End If

                                        Jitter_pic_num = Jitter_pic_num + 1
                                    End If
                                Next 'iout
                                '----------------------------------------------------------------------------------
                                'Add Line
                                report_Group(start_col, first_row, col_num, last_row - first_row + 1)
                                '----------------------------------------------------------------------------------
                                col = col + 1
                                row = row + 1
                            Next  'vin
                            last_row = last_row + 2
                            '----------------------------------------------------------------------------------
                        Case Line_Regulation

                            If check_lineR_scope.Checked = True Then

                                Select Case dut_sel
                                    Case 0
                                        xlSheet = xlBook.Sheets(txt_data_sheet.Text)
                                    Case 1
                                        xlSheet = xlBook.Sheets(txt_data_sheet.Text & add_dut2)
                                End Select
                                xlSheet.Activate()

                                '----------------------------------------------------------------------------------
                                'initial
                                'Init col
                                col_num = lineR_col.Length * data_lineR_iout.Rows.Count + 1


                                row_num = data_Line_vin.Rows.Count + 3

                                start_col = test_col + col_Space + (TA_Test_num * total_vcc.Length * total_fs.Length + n * total_fs.Length + f) * (col_num + 1)

                                first_row = test_row + v * (row_num + row_Space)

                                col = start_col
                                row = first_row


                                chart_num = v + 1


                                '-------------------------------------------------------------------------------
                                'Title

                                report_title(total_title, col, row, col_num, 1, data_title_color)
                                '-------------------------------------------------------------------------------
                                row = row + 1
                                '-------------------------------------------------------------------------------
                                '    |VOUT  |
                                'VIN|n* IOUT
                                '-------------------------------------------------------------------------------
                                For i = 0 To data_lineR_iout.Rows.Count
                                    If i = 0 Then
                                        'X
                                        'Vin
                                        report_title(Vin_name, col, row, 1, 2, data_title_color)
                                        row = row + 2

                                        For ii = 0 To data_Line_vin.Rows.Count - 1

                                            vin_now = data_Line_vin.Rows(ii).Cells(0).Value
                                            xlrange = xlSheet.Range(ConvertToLetter(col) & row)
                                            xlrange.Value = vin_now
                                            FinalReleaseComObject(xlrange)

                                            last_row = row
                                            row = row + 1
                                        Next  'vin
                                        If (rbtn_lineR_test1.Checked = True) And (check_lineR_up.Checked = True) Then
                                            For ii = data_Line_vin.Rows.Count - 2 To 0 Step -1

                                                vin_now = data_Line_vin.Rows(ii).Cells(0).Value
                                                xlrange = xlSheet.Range(ConvertToLetter(col) & row)
                                                xlrange.Value = vin_now
                                                FinalReleaseComObject(xlrange)

                                                last_row = row
                                                row = row + 1
                                            Next  'vin
                                        End If
                                        row = row + 1
                                        col = col + 1
                                    Else
                                        'Y
                                        'Iout
                                        '-------------------------------------------------------------------------------
                                        'Iout
                                        row = first_row + 1
                                        iout_now = data_lineR_iout.Rows(i - 1).Cells(0).Value
                                        report_title("IOUT=" & iout_now & "A", col, row, lineR_col.Length, 1, data_title_color)
                                        row = row + 1

                                        For nn = 0 To lineR_col.Length - 1

                                            report_title(lineR_col(nn), col, row, 1, 1, data_title_color)
                                            col = col + 1
                                        Next
                                    End If
                                    xlSheet.Columns(col).AutoFit()
                                Next  'iout
                                report_Group(start_col, first_row, col_num, last_row - first_row + 1)
                                '----------------------------------------------------------------------------------
                                last_row = last_row + 3
                                '----------------------------------------------------------------------------------
                            End If
                            Select Case dut_sel
                                Case 0
                                    xlSheet = xlBook.Sheets(txt_LineR_sheet.Text)
                                Case 1
                                    xlSheet = xlBook.Sheets(txt_LineR_sheet.Text & add_dut2)
                            End Select
                            xlSheet.Activate()

                            '----------------------------------------------------------------------------------
                            'initial
                            'Init col
                            col_num = data_lineR_iout.Rows.Count + 2
                            row_num = data_Line_vin.Rows.Count + 3
                            start_col = test_col + chart_width + col_Space + (TA_Test_num * total_vcc.Length * total_fs.Length + n * total_fs.Length + f) * (col_num + 1)
                            'init row
                            If row_num < (chart_height + 1) Then
                                first_row = test_row + v * ((chart_height + 1) + row_Space)
                            Else
                                first_row = test_row + v * (row_num + row_Space)

                            End If
                            col = start_col
                            row = first_row
                            chart_num = v + 1
                            '----------------------------------------------------------------------------------
                            'Chart
                            If (TA_Test_num = 0) And (n = 0) And (f = 0) Then
                                chart_col = test_col
                                chart_row = first_row
                                pass_value_Max = vout_now * (1 + (num_pass_lineR.Value * 5 / 100))
                                pass_value_Min = vout_now * (1 - (num_pass_lineR.Value * 5 / 100))
                                chart_init(LineR_Chart, "VOUT=" & vout_now & "V", test_name, vin_title, vout_title, vin_max, vin_min, pass_value_Max, pass_value_Min, cbox_type_LineR.SelectedItem)
                            End If

                            pass_value_Max = vout_now * (1 + (num_pass_lineR.Value / 100))
                            pass_value_Min = vout_now * (1 - (num_pass_lineR.Value / 100))
                            '-------------------------------------------------------------------------------
                            'Title
                            If (TA_Test_num = TA_num) And (n = total_vcc.Length - 1) And (f = total_fs.Length - 1) Then
                                'Add Max, Min 
                                report_title(total_title, col, row, col_num + 2, 1, data_title_color)
                            Else
                                report_title(total_title, col, row, col_num, 1, data_title_color)
                            End If

                            '-------------------------------------------------------------------------------
                            row = row + 1
                            '-------------------------------------------------------------------------------
                            '    |VOUT  |
                            'VIN|n* IOUT| PASS
                            '-------------------------------------------------------------------------------
                            For i = 0 To data_lineR_iout.Rows.Count

                                If i = 0 Then
                                    'X
                                    'Vin
                                    report_title(Vin_name, col, row, 1, 2, data_title_color)
                                    row = row + 2

                                    For ii = 0 To data_Line_vin.Rows.Count - 1

                                        vin_now = data_Line_vin.Rows(ii).Cells(0).Value
                                        xlrange = xlSheet.Range(ConvertToLetter(col) & row)
                                        xlrange.Value = vin_now
                                        FinalReleaseComObject(xlrange)

                                        last_row = row
                                        row = row + 1
                                    Next  'vin

                                    If (rbtn_lineR_test1.Checked = True) And (check_lineR_up.Checked = True) Then
                                        For ii = data_Line_vin.Rows.Count - 2 To 0 Step -1

                                            vin_now = data_Line_vin.Rows(ii).Cells(0).Value
                                            xlrange = xlSheet.Range(ConvertToLetter(col) & row)
                                            xlrange.Value = vin_now
                                            FinalReleaseComObject(xlrange)

                                            last_row = row
                                            row = row + 1
                                        Next  'vin
                                    End If




                                    xlrange = xlSheet.Range(ConvertToLetter(col) & row)
                                    xlrange.Value = "max"
                                    FinalReleaseComObject(xlrange)

                                    row = row + 1
                                    xlrange = xlSheet.Range(ConvertToLetter(col) & row)
                                    xlrange.Value = "min"
                                    FinalReleaseComObject(xlrange)

                                    row = row + 1
                                    xlrange = xlSheet.Range(ConvertToLetter(col) & row)
                                    xlrange.Value = "variation"
                                    FinalReleaseComObject(xlrange)

                                    row = row + 1


                                Else

                                    'Y
                                    '-------------------------------------------------------------------------------
                                    'Iout
                                    row = first_row + 1
                                    report_title(Vout_name, col, row, 1, 1, data_title_color)
                                    row = row + 1
                                    iout_now = data_lineR_iout.Rows(i - 1).Cells(0).Value
                                    report_title("IOUT=" & iout_now & "A", col, row, 1, 1, data_title_color)

                                    '-------------------------------------------------------------------------------
                                    'Add Serial 

                                    chart_row_start = first_row + 3
                                    chart_row_stop = last_row
                                    total_serial = TA_serial & VCC_serial & Fs_serial & "IOUT=" & iout_now & "A"
                                    chart_add_series(txt_LineR_sheet.Text, LineR_Chart, chart_num, total_serial, start_col, col, False)

                                    '-------------------------------------------------------------------------------
                                    'Note
                                    row = last_row + 1
                                    xlrange = xlSheet.Range(ConvertToLetter(col) & row)
                                    xlrange.Value = "=MAX(" & ConvertToLetter(col) & chart_row_start & ":" & ConvertToLetter(col) & (last_row) & ")"
                                    FinalReleaseComObject(xlrange)

                                    ' xlSheet.Cells(row, col) = "=MAX(" & ConvertToLetter(col) & chart_row_start & ":" & ConvertToLetter(col) & (last_row) & ")"
                                    row = row + 1

                                    xlrange = xlSheet.Range(ConvertToLetter(col) & row)
                                    xlrange.Value = "=MIN(" & ConvertToLetter(col) & chart_row_start & ":" & ConvertToLetter(col) & (last_row) & ")"
                                    FinalReleaseComObject(xlrange)

                                    ' xlSheet.Cells(row, col) = "=MIN(" & ConvertToLetter(col) & chart_row_start & ":" & ConvertToLetter(col) & (last_row) & ")"
                                    row = row + 1

                                    xlrange = xlSheet.Range(ConvertToLetter(col) & row)
                                    xlrange.Value = "=(" & ConvertToLetter(col) & (last_row + 1) & "-" & ConvertToLetter(col) & (last_row + 2) & ")/" & vout_now


                                    'xlSheet.Cells(row, col) = "=(" & ConvertToLetter(col) & (last_row + 1) & "-" & ConvertToLetter(col) & (last_row + 2) & ")/" & vout_now
                                    ' xlSheet.Range(ConvertToLetter(col) & row & ":" & ConvertToLetter(col) & row).NumberFormatLocal = "0.00%"
                                    xlrange.NumberFormatLocal = "0.00%"
                                    FinalReleaseComObject(xlrange)
                                    row = row + 1
                                    '-------------------------------------------------------------------------------



                                End If
                                xlSheet.Columns(col).AutoFit()
                                col = col + 1


                            Next  'iout

                            '----------------------------------------------------------------------------------

                            ' PASS & Criteria

                            For ii = 0 To data_Line_vin.Rows.Count - 1
                                col = start_col + data_lineR_iout.Rows.Count + 1

                                '----------------------------------------------------------------------------------
                                'Only Last parameter

                                If (TA_Test_num = TA_num) And (n = total_vcc.Length - 1) And (f = total_fs.Length - 1) Then
                                    If ii = 0 Then

                                        row = first_row + 1
                                        report_title("Max. Criteria", col, row, 1, 2, data_title_color)

                                        '----------------------------------------------------------------------------------
                                        'Add Serial 

                                        chart_row_start = first_row + 3
                                        chart_row_stop = chart_row_start + data_Line_vin.Rows.Count - 1
                                        chart_add_series(txt_LineR_sheet.Text, LineR_Chart, chart_num, "Max. Criteria", start_col, col, True)
                                        '----------------------------------------------------------------------------------
                                        row = row + 2
                                    End If

                                    xlrange = xlSheet.Range(ConvertToLetter(col) & row)
                                    xlrange.Value = pass_value_Max
                                    FinalReleaseComObject(xlrange)

                                    '  xlSheet.Cells(row, col) = pass_value_Max


                                    col = col + 1

                                    If ii = 0 Then

                                        row = first_row + 1
                                        report_title("Min. Criteria", col, row, 1, 2, data_title_color)

                                        '----------------------------------------------------------------------------------
                                        'Add Serial 

                                        chart_row_start = first_row + 3
                                        chart_row_stop = chart_row_start + data_Line_vin.Rows.Count - 1
                                        chart_add_series(txt_LineR_sheet.Text, LineR_Chart, chart_num, "Min. Criteria", start_col, col, True)

                                        '----------------------------------------------------------------------------------
                                        row = row + 2
                                    End If
                                    xlrange = xlSheet.Range(ConvertToLetter(col) & row)
                                    xlrange.Value = pass_value_Min
                                    FinalReleaseComObject(xlrange)
                                    'xlSheet.Cells(row, col) = pass_value_Min

                                    col = col + 1
                                End If

                                '----------------------------------------------------------------------------------
                                'PASS
                                If ii = 0 Then

                                    row = first_row + 1
                                    report_title("PASS/FAIL", col, row, 1, 2, data_title_color)
                                    row = row + 2
                                End If

                                col = col + 1

                                row = row + 1
                            Next  'vin

                            '----------------------------------------------------------------------------------
                            'Add Line
                            If (TA_Test_num = TA_num) And (n = total_vcc.Length - 1) And (f = total_fs.Length - 1) Then

                                report_Group(start_col, first_row, col_num + 2, last_row - first_row + 1)

                            Else

                                report_Group(start_col, first_row, col_num, last_row - first_row + 1)
                            End If
                            '----------------------------------------------------------------------------------
                            last_row = last_row + 3


                            '----------------------------------------------------------------------------------








                        Case Load_Regulation

                            Select Case dut_sel
                                Case 0
                                    xlSheet = xlBook.Sheets(txt_LoadR_sheet.Text)
                                Case 1
                                    xlSheet = xlBook.Sheets(txt_LoadR_sheet.Text & add_dut2)
                            End Select

                            xlSheet.Activate()
                            '----------------------------------------------------------------------------------
                            'initial
                            'Init col

                            col_num = data_vin.Rows.Count + 2
                            row_num = data_eff_iout.Rows.Count + 3
                            start_col = test_col + chart_width + col_Space + (TA_Test_num * total_vcc.Length * total_fs.Length + n * total_fs.Length + f) * (col_num + 1)
                            'Init row

                            If row_num < (chart_height + 1) Then
                                first_row = test_row + v * ((chart_height + 1) + row_Space)
                            Else
                                first_row = test_row + v * (row_num + row_Space)

                            End If




                            col = start_col
                            row = first_row



                            chart_num = v + 1
                            '----------------------------------------------------------------------------------
                            'Chart
                            If (TA_Test_num = 0) And (n = 0) And (f = 0) Then

                                chart_col = test_col
                                chart_row = first_row

                                pass_value_Max = vout_now * (1 + (num_pass_loadR.Value * 5 / 100))
                                pass_value_Min = vout_now * (1 - (num_pass_loadR.Value * 5 / 100))
                                iout_now = data_eff_iout.Rows(data_eff_iout.Rows.Count - 1).Cells(0).Value
                                chart_init(LoadR_Chart, "VOUT=" & vout_now & "V", test_name, Iout_title, vout_title, iout_now, 0, pass_value_Max, pass_value_Min, cbox_type_LoadR.SelectedItem)

                            End If

                            pass_value_Max = vout_now * (1 + (num_pass_loadR.Value / 100))
                            pass_value_Min = vout_now * (1 - (num_pass_loadR.Value / 100))
                            '-------------------------------------------------------------------------------
                            'Title
                            If (TA_Test_num = TA_num) And (n = total_vcc.Length - 1) And (f = total_fs.Length - 1) Then
                                'Add Max, Min 
                                report_title(total_title, col, row, col_num + 2, 1, data_title_color)
                            Else
                                report_title(total_title, col, row, col_num, 1, data_title_color)
                            End If

                            '-------------------------------------------------------------------------------
                            row = row + 1
                            '-------------------------------------------------------------------------------
                            '    |VOUT  |
                            'IOUT|n* VIN| PASS
                            '-------------------------------------------------------------------------------
                            For i = 0 To data_vin.Rows.Count

                                If i = 0 Then
                                    'X
                                    'Iout
                                    report_title(Iout_name, col, row, 1, 2, data_title_color)
                                    row = row + 2

                                    For ii = 0 To data_eff_iout.Rows.Count - 1

                                        iout_now = data_eff_iout.Rows(ii).Cells(0).Value
                                        xlrange = xlSheet.Range(ConvertToLetter(col) & row)
                                        xlrange.Value = iout_now
                                        FinalReleaseComObject(xlrange)
                                        'xlSheet.Cells(row, col) = iout_now
                                        last_row = row
                                        row = row + 1
                                    Next  'iout

                                    xlrange = xlSheet.Range(ConvertToLetter(col) & row)
                                    xlrange.Value = "max"
                                    FinalReleaseComObject(xlrange)
                                    'xlSheet.Cells(row, col) = "max"
                                    row = row + 1
                                    xlrange = xlSheet.Range(ConvertToLetter(col) & row)
                                    xlrange.Value = "min"
                                    FinalReleaseComObject(xlrange)
                                    ' xlSheet.Cells(row, col) = "min"
                                    row = row + 1
                                    xlrange = xlSheet.Range(ConvertToLetter(col) & row)
                                    xlrange.Value = "variation"
                                    FinalReleaseComObject(xlrange)
                                    ' xlSheet.Cells(row, col) = "variation"
                                    row = row + 1


                                Else

                                    'Y
                                    '-------------------------------------------------------------------------------
                                    'Vin
                                    row = first_row + 1
                                    report_title(Vout_name, col, row, 1, 1, data_title_color)
                                    row = row + 1
                                    vin_now = data_vin.Rows(i - 1).Cells(0).Value
                                    report_title("VIN=" & vin_now & "V", col, row, 1, 1, data_title_color)

                                    '-------------------------------------------------------------------------------
                                    'Add Serial 

                                    chart_row_start = first_row + 3
                                    chart_row_stop = last_row
                                    total_serial = TA_serial & VCC_serial & Fs_serial & "VIN=" & vin_now & "V"
                                    chart_add_series(txt_LoadR_sheet.Text, LoadR_Chart, chart_num, total_serial, start_col, col, False)

                                    '-------------------------------------------------------------------------------
                                    'Note
                                    row = last_row + 1
                                    xlrange = xlSheet.Range(ConvertToLetter(col) & row)
                                    xlrange.Value = "=MAX(" & ConvertToLetter(col) & chart_row_start & ":" & ConvertToLetter(col) & (last_row) & ")"
                                    FinalReleaseComObject(xlrange)
                                    'xlSheet.Cells(row, col) = "=MAX(" & ConvertToLetter(col) & chart_row_start & ":" & ConvertToLetter(col) & (last_row) & ")"
                                    row = row + 1
                                    xlrange = xlSheet.Range(ConvertToLetter(col) & row)
                                    xlrange.Value = "=MIN(" & ConvertToLetter(col) & chart_row_start & ":" & ConvertToLetter(col) & (last_row) & ")"
                                    FinalReleaseComObject(xlrange)
                                    'xlSheet.Cells(row, col) = "=MIN(" & ConvertToLetter(col) & chart_row_start & ":" & ConvertToLetter(col) & (last_row) & ")"
                                    row = row + 1
                                    xlrange = xlSheet.Range(ConvertToLetter(col) & row)
                                    xlrange.Value = "=(" & ConvertToLetter(col) & (last_row + 1) & "-" & ConvertToLetter(col) & (last_row + 2) & ")/" & vout_now

                                    xlrange.NumberFormatLocal = "0.00%"
                                    FinalReleaseComObject(xlrange)
                                    row = row + 1
                                    '-------------------------------------------------------------------------------



                                End If
                                xlSheet.Columns(col).AutoFit()
                                col = col + 1


                            Next  'iout

                            '----------------------------------------------------------------------------------

                            ' PASS & Criteria

                            For ii = 0 To data_eff_iout.Rows.Count - 1
                                col = start_col + data_vin.Rows.Count + 1

                                '----------------------------------------------------------------------------------
                                'Only Last parameter

                                If (TA_Test_num = TA_num) And (n = total_vcc.Length - 1) And (f = total_fs.Length - 1) Then
                                    If ii = 0 Then

                                        row = first_row + 1
                                        report_title("Max. Criteria", col, row, 1, 2, data_title_color)

                                        '----------------------------------------------------------------------------------
                                        'Add Serial 

                                        chart_row_start = first_row + 3
                                        chart_row_stop = chart_row_start + data_eff_iout.Rows.Count - 1
                                        chart_add_series(txt_LoadR_sheet.Text, LoadR_Chart, chart_num, "Max. Criteria", start_col, col, True)
                                        '----------------------------------------------------------------------------------
                                        row = row + 2
                                    End If

                                    xlrange = xlSheet.Range(ConvertToLetter(col) & row)
                                    xlrange.Value = pass_value_Max
                                    FinalReleaseComObject(xlrange)
                                    ' xlSheet.Cells(row, col) = pass_value_Max


                                    col = col + 1

                                    If ii = 0 Then

                                        row = first_row + 1
                                        report_title("Min. Criteria", col, row, 1, 2, data_title_color)

                                        '----------------------------------------------------------------------------------
                                        'Add Serial 

                                        chart_row_start = first_row + 3
                                        chart_row_stop = chart_row_start + data_eff_iout.Rows.Count - 1
                                        chart_add_series(txt_LoadR_sheet.Text, LoadR_Chart, chart_num, "Min. Criteria", start_col, col, True)

                                        '----------------------------------------------------------------------------------
                                        row = row + 2
                                    End If
                                    xlrange = xlSheet.Range(ConvertToLetter(col) & row)
                                    xlrange.Value = pass_value_Min
                                    FinalReleaseComObject(xlrange)
                                    'xlSheet.Cells(row, col) = pass_value_Min

                                    col = col + 1
                                End If

                                '----------------------------------------------------------------------------------
                                'PASS
                                If ii = 0 Then

                                    row = first_row + 1
                                    report_title("PASS/FAIL", col, row, 1, 2, data_title_color)
                                    row = row + 2
                                End If


                                col = col + 1

                                row = row + 1
                            Next  'iout

                            '----------------------------------------------------------------------------------
                            'Add Line
                            If (TA_Test_num = TA_num) And (n = total_vcc.Length - 1) And (f = total_fs.Length - 1) Then

                                report_Group(start_col, first_row, col_num + 2, last_row - first_row + 1)

                            Else

                                report_Group(start_col, first_row, col_num, last_row - first_row + 1)
                            End If
                            '----------------------------------------------------------------------------------
                            last_row = last_row + 3


                            '----------------------------------------------------------------------------------
                            '----------------------------------------------------------------------------------

                        Case Efficiency
                            Select Case dut_sel
                                Case 0
                                    xlSheet = xlBook.Sheets(txt_eff_sheet.Text)
                                Case 1
                                    xlSheet = xlBook.Sheets(txt_eff_sheet.Text & add_dut2)
                            End Select
                            xlSheet.Activate()
                            '----------------------------------------------------------------------------------
                            'initial
                            'Init col

                            col_num = eff_col.Length
                            row_num = data_eff_iout.Rows.Count + 2
                            start_col = test_col + chart_width + col_Space + (n * total_fs.Length + f) * ((col_num + 1) * (data_vin.Rows.Count))
                            'Init row

                            If row_num < (chart_height + 1) Then

                                first_row = test_row + (v * (TA_num + 1) + TA_Test_num) * ((chart_height + 1) + row_Space)
                            Else
                                first_row = test_row + (v * (TA_num + 1) + TA_Test_num) * (row_num + row_Space)

                            End If



                            col = start_col
                            row = first_row


                            'Total Chart
                            chart_num = TA_Test_num * total_vout.Length + v + 1
                            '----------------------------------------------------------------------------------
                            'Chart
                            If (n = 0) And (f = 0) Then

                                chart_col = test_col
                                chart_row = first_row

                                pass_value_Max = 100
                                pass_value_Min = 0
                                iout_now = data_eff_iout.Rows(data_eff_iout.Rows.Count - 1).Cells(0).Value

                                chart_init(Eff_Chart, TA_title & "VOUT=" & vout_now & "V", test_name, Iout_title, "Efficiency (%)", iout_now, 0, pass_value_Max, pass_value_Min, cbox_type_Eff.SelectedItem)




                            End If
                            '-------------------------------------------------------------------------------

                            'pass_value_Min = num_pass_eff.Value
                            '-------------------------------------------------------------------------------
                            'VIN (V)	IIN(A)	VOUT (V)	IOUT (A)	Efficiency(%)	Loss(W)	PASS/FAIL
                            '-------------------------------------------------------------------------------

                            For i = 0 To data_vin.Rows.Count - 1
                                row = first_row
                                start_col = col
                                col = start_col
                                vin_now = data_vin.Rows(i).Cells(0).Value
                                total_title = TA_title & VCC_title & Fs_title & "VOUT=" & vout_now & "V, VIN=" & vin_now & "V"
                                'Title
                                report_title(total_title, col, row, col_num, 1, data_title_color)
                                row = row + 1
                                For nn = 0 To eff_col.Length - 1
                                    report_title(eff_col(nn), col, row, 1, 1, data_title_color)
                                    xlSheet.Columns(col).AutoFit()
                                    col = col + 1

                                Next



                                row = row + 1
                                'X
                                For ii = 0 To data_eff_iout.Rows.Count - 1

                                    iout_now = data_eff_iout.Rows(ii).Cells(0).Value

                                    col = start_col


                                    For nn = 0 To eff_col.Length - 1


                                        col = col + 1



                                    Next 'eff
                                    last_row = row
                                    row = row + 1


                                Next 'iout



                                '-------------------------------------------------------------------------------
                                'Add Serial 

                                chart_row_start = first_row + 2
                                chart_row_stop = last_row
                                total_serial = TA_serial & VCC_serial & Fs_serial & "VIN=" & vin_now & "V"
                                chart_add_series(txt_eff_sheet.Text, Eff_Chart, chart_num, total_serial, start_col + 3, start_col + 4, False)

                                xlSheet.Columns(col).AutoFit()


                                '----------------------------------------------------------------------------------
                                'Add Line
                                report_Group(start_col, first_row, col_num, last_row - first_row + 1)

                                '----------------------------------------------------------------------------------
                                col = col + 1
                            Next  'vin


                            last_row = last_row + 2


                            '----------------------------------------------------------------------------------








                    End Select




                Next   'vout

                xlBook.Save()


            Next  'fs

            xlBook.Save()

        Next  'vcc

        FinalReleaseComObject(xlSheet)
        xlSheet = Nothing

        xlBook.Save()


        'Next  'TA





    End Function

    Function iin_range_report_info() As Integer

        Dim i, ii As Integer
        Dim report_col As Integer = 2
        Dim note() As String = {"TA", "VIN (V)", "VOUT (V)", "IOUT (mA)"}

        '------------------------------------------------------------------------------------
        'Initial Page

        report_title("IIN Range Change", report_col, iin_row, 2 + (TA_num + 1) * data_vin.Rows.Count, 1, 44)
        xlSheet.Activate()

        For ii = 0 To note.Length - 1
            'xlSheet.Cells(8 + ii, report_col) = note(ii)
            report_title("", report_col, iin_row + 1 + ii, 2, 1, 2)
            xlrange = xlSheet.Range(ConvertToLetter(report_col) & iin_row + 1 + ii)
            xlrange.Value = note(ii)
            FinalReleaseComObject(xlrange)
        Next

        report_Group(report_col, iin_row, 2 + (TA_num + 1) * data_vin.Rows.Count, note.Length + 1)

    End Function

    Function iin_range_update(ByVal num As Integer, ByVal note() As String) As Integer
        Dim i As Integer
        Dim report_col As Integer = 2
        xlrange = xlSheet.Range(ConvertToLetter(report_col + 2 + TA_Test_num * data_vin.Rows.Count + num) & (iin_row + 1))
        For i = 0 To note.Length - 1
            xlrange.Offset(i, 0).Value = note(i)
        Next

        FinalReleaseComObject(xlrange)
    End Function

    Function check_meter_iin_max() As Integer
        Dim iout_start As Double
        Dim iout_stop As Double
        Dim iout_step As Double
        Dim Iin_max_check As Double
        Dim temp As String
        Dim Iin_default As Double
        Dim iout_start_set As Double
        Dim i As Integer
        Dim eff_set_num As Integer = 0

        iout_start_set = 0
        For i = 0 To data_eff.Rows.Count - 1

            If (data_eff.Rows(i).Cells(0).Value = vin_now) And (data_eff.Rows(i).Cells(1).Value = vout_now) Then
                iout_start_set = data_eff.Rows(i).Cells(2).Value
                eff_set_num = i
                Exit For
            End If
        Next
        Iin_change = False
        '先確認0A的時候的電流
        'DCLoad_Iout(0, False)
        'DCLoad_Iout(0, False, DUT2_en)
        Iout_Setting(0, False, DUT2_en)
        Power_Dev = vin_Dev
        Iin_default = power_read(vin_device, Vin_out, "CURR")
        If iin_meter_change > Iin_default Then
            If rbtn_meter_iin.Checked = True Then

                If check_iin.Checked = True Then
                    Iin_Meter_initial(check_iin, cbox_IIN_meter, cbox_IIN_relay) 'High Range
                End If
            Else
                'set low
                INA226_Iin_initial(False)
            End If
            iin_meter_change = num_iin_change.Value / 1000
            If iin_meter_change < Meter_iin_Max Then
                Iin_max_check = iin_meter_change
            Else
                Iin_max_check = Meter_iin_Max * 0.9
            End If
            iout_step = Math.Round(num_iin_step.Value / 1000, 4)
            If rbtn_iin_auto.Checked = True Then
                iout_start = Math.Round(iout_start_set / 1000, 4)
                iout_stop = Math.Round(num_iout_auto_stop.Value / 1000, 4)
                For iin_meter_change = iout_start To iout_stop Step iout_step
                    System.Windows.Forms.Application.DoEvents()
                    If run = False Then
                        Exit For
                    End If
                    'DCLoad_Iout(iin_meter_change, False)
                    'DCLoad_Iout(iin_meter_change, False, DUT2_en)
                    Iout_Setting(iin_meter_change, False, DUT2_en)
                    If (DCLoad_ON = False) Then
                        DCLoad_ONOFF("ON")
                        'If DUT2_en Then
                        '    DCLoad_ONOFF("ON", DUT2_en)
                        'End If
                    End If
                    If rbtn_meter_iin.Checked = True Then
                        iin_meas = Math.Abs(meter_meas(cbox_IIN_meter.SelectedItem, Meter_iin_dev, Meter_iin_range, Meter_iin_low))
                    Else
                        iin_meas = INA226_IIN_meas(1)
                    End If
                    check_vout()
                    If DUT2_en Then
                        check_vout2()
                    End If
                    If iin_meas > Iin_max_check Then
                        iin_meter_change = iin_meter_change - iout_step
                        Exit For
                    End If
                Next
            Else
                iin_meter_change = iout_start_set / 1000
            End If
            If rbtn_board_iin.Checked = True Then
                'set High
                INA226_Iin_initial(True)
            End If
            Iin_change = True
        Else
            iin_meter_change = 0
        End If
        '--------------------------------------------------
        xlSheet = xlBook.Sheets(txt_eff_sheet.Text)
        xlSheet.Activate()
        Dim note(3) As String
        note(0) = TA_now
        note(1) = vin_now
        note(2) = vout_now
        If iin_meter_change = 0 Then
            note(3) = "0 (Iin=" & Iin_default & "A)"
        Else
            note(3) = iin_meter_change * 10 ^ 3
        End If
        iin_range_update(eff_set_num, note)
        ReDim Preserve eff_iin_change(eff_set_num)
        eff_iin_change(eff_set_num) = iin_meter_change
        '--------------------------------------------------
        'DCLoad_Iout(iout_now, monitor_vout)
        'DCLoad_Iout(0, monitor_vout)
        'DCLoad_Iout(0, monitor_vout, DUT2_en)

        Iout_Setting(0, monitor_vout, DUT2_en)
        'DCLoad_ONOFF("OFF")
        Delay(100)
    End Function

    Function LineR_run(Optional ByVal dut2_en As Boolean = False) As Integer

        '-----------------------------------------------------------------------------------------------------------

        Power_Dev = vin_Dev
        power_volt(vin_device, Vin_out, vin_now)
        ''----------------------------------------------------------------------------------
        'Vin Sense
        If check_vin_sense.Checked = True Then
            'Vin Sense
            vin_power_sense(cbox_vin.SelectedItem, num_vin_sense.Value, num_vin_max.Value, vin_now)
        End If

        If dut2_en Then
            Power_Dev = vin_Dev2
            power_volt(vin_device, Vin_out, vin_now)
            ''----------------------------------------------------------------------------------
            'Vin Sense
            If check_vin_sense2.Checked = True Then
                'Vin Sense
                vin_power_sense(cbox_vin2.SelectedItem, num_vin_sense2.Value, num_vin_max2.Value, vin_now)
            End If
        End If


        ''----------------------------------------------------------------------------------
        'Measure

        If (iout_now > num_iout_delay.Value) And (num_delay.Value > 0) Then

            If cbox_delay_unit.SelectedIndex = 1 Then
                Delay_s(num_delay.Value)
            Else
                Delay(num_delay.Value)
            End If
        End If

        ''----------------------------------------------------------------------------------
        'Check Vout
        'vout
        vout_meas = DAQ_average(vout_daq, num_data_count.Value)
        If dut2_en Then
            vout_meas2 = DAQ_average(vout_daq2, num_data_count.Value)
        End If
        If check_lineR_scope.Checked = True Then
            'Time Scale
            Scope_RUN(False)
            Fs_leak_0A = num_fs_leak.Value
            ton_now = (vout_now / vin_now) * (1 / fs_now)
            Calculate_pass(TA_Test_num)
            'Timing Scale

            If (check_Force_CCM.Checked = False) And (iout_now = 0) Then
                H_scale_value = ((1 / Fs_Min) * 10 / 10) * (10 ^ 9) '1/Fs_Min(Hz)*n/10 
            Else
                H_scale_value = ((1 / Fs_Min) * Wave_num / 10) * (10 ^ 9) '1/Fs_Min(Hz)*n/10 
            End If
            'Timing Scale
            H_scale(H_scale_value, "ns") '1/Fs_Min(Hz)*n/10 
            'Scope_RUN(True)
            If RS_Scope = True Then
                RS_View(True)
            End If
            'Lx Scale

            'CHx_scale(lx_ch, (vin_now / num_lx_scale.Value), "V") 'Voltage Scale > SW/2

            If rbtn_manual_lx.Checked = True Then
                CHx_scale(lx_ch, num_scale_lx.Value, "mV") 'Voltage Scale > SW/2
            Else
                CHx_scale(lx_ch, (vin_now / num_lx_scale.Value), "V") 'Voltage Scale > SW/2
            End If

            If rbtn_vin_trigger.Checked = True Then
                Trigger_set(lx_ch, "R", vin_now / num_vin_trigger.Value)
            Else
                Trigger_auto_level(lx_ch, "R")
            End If


            If dut2_en Then
                If rbtn_manual_lx.Checked = True Then
                    CHx_scale(lx2_ch, num_scale_lx.Value, "mV") 'Voltage Scale > SW/2
                Else
                    CHx_scale(lx2_ch, (vin_now / num_lx_scale.Value), "V") 'Voltage Scale > SW/2
                End If

                If rbtn_vin_trigger.Checked = True Then
                    Trigger_set(lx2_ch, "R", vin_now / num_vin_trigger.Value)
                Else
                    Trigger_auto_level(lx2_ch, "R")
                End If
            End If
            'Scope_RUN(True)
            monitor_count(num_counts_CCM.Value, True, "Part I")
            ''----------------------------------------------------------------------------------
            'Measurement
            'Scope

            'KHz
            fs(0) = Scope_measure(meas1, Scope_Meas)
            fs(1) = Scope_measure(meas1, Meas_mean)
            fs(2) = Scope_measure(meas1, Meas_min)
            fs(3) = Scope_measure(meas1, Meas_max)
            'Ton (ns)

            ton(0) = Scope_measure(meas2, Scope_Meas)
            ton(1) = Scope_measure(meas2, Meas_mean)
            ton(2) = Scope_measure(meas2, Meas_min)
            ton(3) = Scope_measure(meas2, Meas_max)

            'Toff
            toff(0) = Scope_measure(meas3, Scope_Meas)
            toff(1) = Scope_measure(meas3, Meas_mean)
            toff(2) = Scope_measure(meas3, Meas_min)
            toff(3) = Scope_measure(meas3, Meas_max)


            If dut2_en Then
                fs(0) = Scope_measure(meas4, Scope_Meas)
                fs(1) = Scope_measure(meas4, Meas_mean)
                fs(2) = Scope_measure(meas4, Meas_min)
                fs(3) = Scope_measure(meas4, Meas_max)
                'Ton (ns)

                ton(0) = Scope_measure(meas5, Scope_Meas)
                ton(1) = Scope_measure(meas5, Meas_mean)
                ton(2) = Scope_measure(meas5, Meas_min)
                ton(3) = Scope_measure(meas5, Meas_max)

                'Toff
                toff(0) = Scope_measure(meas6, Scope_Meas)
                toff(1) = Scope_measure(meas6, Meas_mean)
                toff(2) = Scope_measure(meas6, Meas_min)
                toff(3) = Scope_measure(meas6, Meas_max)
            End If


        End If
    End Function

    Function Efficiency_run(Optional ByVal dut2_en As Boolean = False) As Integer
        'Eff & LoadR
        If check_Efficiency.Checked = True Then

            If vout_daq <> Eff_vout_daq Then
                Eff_vout_meas = DAQ_read(Eff_vout_daq)
            Else
                Eff_vout_meas = vout_meas
            End If


            If dut2_en Then
                If vout_daq2 <> Eff_vout_daq2 Then
                    Eff_vout_meas2 = DAQ_read(Eff_vout_daq2)
                Else
                    Eff_vout_meas2 = vout_meas2
                End If
            End If


            'iin
            If rbtn_meter_iin.Checked = True Then
                iin_meas = meter_average(cbox_IIN_meter.SelectedItem, Meter_iin_dev, num_data_count.Value, Meter_iin_range, Meter_iin_low)
                Meter_iin_range = Meter_range_now
            ElseIf rbtn_board_iin.Checked = True Then
                'relay read
                iin_meas = INA226_IIN_meas(num_data_count.Value)
            ElseIf rbtn_iin_current_measure.Checked Then
                iin_meas = meter_auto(0, num_meter_count.Value)
            Else
                iin_meas = power_read(cbox_vin.SelectedItem, Vin_out, "CURR") ' Format(power_read(cbox_vin.SelectedItem, Vin_out, "CURR"), "#0.000000000")
            End If


            If dut2_en Then
                If rbtn_meter_iin2.Checked = True Then
                    iin_meas2 = meter_average(cbox_IIN_meter2.SelectedItem, Meter_iin_dev2, num_data_count2.Value, Meter_iin_range, Meter_iin_low)
                    Meter_iin_range = Meter_range_now
                ElseIf rbtn_board_iin2.Checked = True Then
                    'relay read
                    iin_meas2 = INA226_IIN_meas(num_data_count.Value)
                ElseIf rbtn_iin_current_measure2.Checked Then
                    iin_meas2 = meter_auto(0, num_meter_count.Value, dut2_en)
                Else
                    iin_meas2 = power_read(cbox_vin2.SelectedItem, Vin_out2, "CURR", dut2_en) ' Format(power_read(cbox_vin.SelectedItem, Vin_out, "CURR"), "#0.000000000")
                End If
            End If

            If cbox_VCC_daq.SelectedItem <> no_device Then
                vcc_meas = DAQ_average(vcc_daq, num_data_count.Value)
            ElseIf cbox_VCC.SelectedItem <> no_device Then
                Power_Dev = VCC_Dev
                If cbox_VCC.SelectedItem = " 2230-30-1" Then
                    vcc_meas = Power2230_read(VCC_out, "VOLT")
                Else
                    vcc_meas = power_read(cbox_VCC.SelectedItem, VCC_out, "VOLT")
                End If
                Power_Dev = vin_Dev
            End If
            If txt_Icc_addr.Text <> "" Then
                icc_meas = meter_average(cbox_Icc_meter.SelectedItem, Meter_icc_dev, num_data_count.Value, Meter_iout_range, "4e-1") ' meter_read(Meter_icc_dev)
            End If
            update_report(Efficiency, dut2_en)
        End If
        If check_loadR.Checked = True Then
            update_report(Load_Regulation, dut2_en)
        End If
    End Function

    Function GetVoutMax_Min(ByVal vout_sel As Integer) As Double()
        Dim res() As Double = {0, 0}
        Dim vout_list As List(Of Double) = New List(Of Double)()
        Dim ByteSize As Long
        Dim max_data As Integer
        Dim list() As String
        Dim scale_value As Double
        note_string = "Check Scope ..."
        scale_value = H_scale_now()
        Dim RL As Integer
        Dim Point As Integer = 7


        'If ton_pass > toff_pass Then
        '    RL = Point / (ton_pass) * 5 * scale_value
        'Else
        '    RL = Point / (toff_pass) * 5 * scale_value
        'End If

        'If RL > 500000 Then
        '    Return res

        '    Exit Function
        'ElseIf (RL > 250000) Then

        '    H_reclength(500000)

        'ElseIf (RL > 100000) Then

        '    H_reclength(250000)

        'ElseIf (RL > 50000) Then

        '    H_reclength(100000)

        'ElseIf (RL > 20000) Then

        '    H_reclength(50000)
        'Else
        '    H_reclength(20000)
        'End If
        ByteSize = Waveform_data(Main.txt_scope_folder.Text & "\wave.csv", wave_pc_path, vout_sel)



        If run = False Then
            Return res
            Exit Function
        End If

        If ByteSize = 0 Then
            check_file_open(wave_pc_path)
            ByteSize = Waveform_data(Main.txt_scope_folder.Text & "\wave.csv", wave_pc_path, vout_sel)
            If run = False Then
                Return res
                Exit Function
            End If
        End If
        If ByteSize > 0 Then
            Dim f As New IO.FileInfo(wave_pc_path)
            Dim sr As IO.StreamReader = f.OpenText '產生StreamReader的sr物件
            note_string = "Get Wave data..."
            list = Split(sr.ReadToEnd, vbNewLine)

            max_data = list.Length
            sr.Close()

            For i = 1 To max_data
                Dim temp = Split(list(i - 1), ",") ' 0 = time, 1 = vout
                If list(i - 1) <> "" Then
                    vout_list.Add(Val(temp(1)))
                End If
            Next

            res(0) = vout_list.Max()
            res(1) = vout_list.Min()

        End If

        Return res ' vmax, vmin
    End Function

    Sub Stability_Measure_Set(ByVal dut_sel As Integer)

        Select Case dut_sel
            Case 0
                ' dut1 measure set
                ' lx info
                Scope_measure_set(meas1, lx_ch, "FREQuency")
                Scope_measure_set(meas2, lx_ch, "PWIdth")
                Scope_measure_set(meas3, lx_ch, "NWIdth")
                Scope_measure_set(meas4, vout_ch, "PK2Pk")
                Scope_measure_set(meas5, vout_ch, "MAXimum")
                Scope_measure_set(meas6, vout_ch, "MINImum")
            Case 1
                Scope_measure_set(meas1, lx2_ch, "FREQuency")
                Scope_measure_set(meas2, lx2_ch, "PWIdth")
                Scope_measure_set(meas3, lx2_ch, "NWIdth")
                Scope_measure_set(meas4, vout2_ch, "PK2Pk")
                Scope_measure_set(meas5, vout2_ch, "MAXimum")
                Scope_measure_set(meas6, vout2_ch, "MINImum")
        End Select

    End Sub

    Sub Stability_Measure_Get(ByVal dut_sel As Integer)
        Select Case dut_sel
            Case 0
                fs(0) = Scope_measure(meas1, Scope_Meas)
                fs(1) = Scope_measure(meas1, Meas_mean)
                fs(2) = Scope_measure(meas1, Meas_min)
                fs(3) = Scope_measure(meas1, Meas_max)
                'Ton (ns)
                ton(0) = Scope_measure(meas2, Scope_Meas)
                ton(1) = Scope_measure(meas2, Meas_mean)
                ton(2) = Scope_measure(meas2, Meas_min)
                ton(3) = Scope_measure(meas2, Meas_max)
                'Toff
                toff(0) = Scope_measure(meas3, Scope_Meas)
                toff(1) = Scope_measure(meas3, Meas_mean)
                toff(2) = Scope_measure(meas3, Meas_min)
                toff(3) = Scope_measure(meas3, Meas_max)
                'vpp
                vpp(0) = Scope_measure(meas4, Scope_Meas)
                vpp(1) = Scope_measure(meas4, Meas_mean)
                vpp(2) = Scope_measure(meas4, Meas_min)
                vpp(3) = Scope_measure(meas4, Meas_max)
                'Vmax
                vpp(4) = Scope_measure(meas5, Meas_max)
                'Vmin
                vpp(5) = Scope_measure(meas6, Meas_min)
            Case 1
                fs(0) = Scope_measure(meas1, Scope_Meas)
                fs(1) = Scope_measure(meas1, Meas_mean)
                fs(2) = Scope_measure(meas1, Meas_min)
                fs(3) = Scope_measure(meas1, Meas_max)
                'Ton (ns)
                ton(0) = Scope_measure(meas2, Scope_Meas)
                ton(1) = Scope_measure(meas2, Meas_mean)
                ton(2) = Scope_measure(meas2, Meas_min)
                ton(3) = Scope_measure(meas2, Meas_max)
                'Toff
                toff(0) = Scope_measure(meas3, Scope_Meas)
                toff(1) = Scope_measure(meas3, Meas_mean)
                toff(2) = Scope_measure(meas3, Meas_min)
                toff(3) = Scope_measure(meas3, Meas_max)
                'vpp
                vpp(0) = Scope_measure(meas4, Scope_Meas)
                vpp(1) = Scope_measure(meas4, Meas_mean)
                vpp(2) = Scope_measure(meas4, Meas_min)
                vpp(3) = Scope_measure(meas4, Meas_max)
                'Vmax
                vpp(4) = Scope_measure(meas5, Meas_max)
                'Vmin
                vpp(5) = Scope_measure(meas6, Meas_min)
        End Select
    End Sub


    Sub Stability_run1(ByVal cnt As Integer)
        Dim vout_temp As Double
        Dim vout_scale_temp As Integer
        Dim iout_temp As Double
        Dim double_check As Boolean = False

        'System.Threading.Thread.Sleep(1000)
        ' all channel enable
        If DUT2_en Then
            CHx_display(1, "ON")
            CHx_display(2, "ON")
            CHx_display(3, "ON")
            CHx_display(4, "ON")
        Else
            CHx_display(lx_ch, "ON")
            CHx_display(vout_ch, "ON")
        End If
        scope_time_init()
        Scope_RUN(False)
        Calculate_pass(TA_Test_num)
        If (check_Force_CCM.Checked = False) And (iout_now = 0) Then
            H_scale_value = ((1 / Fs_Min) * 10 / 10) * (10 ^ 9) '1/Fs_Min(Hz)*n/10 
        Else
            H_scale_value = ((1 / Fs_Min) * Wave_num / 10) * (10 ^ 9) '1/Fs_Min(Hz)*n/10 
        End If
        'Timing Scale
        H_scale(H_scale_value, "ns") ' 1 / Fs_Min(Hz) * n / 10 
        If RS_Scope = True Then
            RS_View(True)
        End If

        ' vout level scale re-size
        If (iout_now = data_test.Rows(0).Cells(0).Value) Then
            If rbtn_auto_vout.Checked = True Then
                vout_temp = Math.Floor(((vout_meas * 1000) * 0.005))
                vout_temp = Math.Floor(vout_temp / 5) * 5
                If vout_temp > 10 Then
                    vout_scale_now = vout_temp
                Else
                    vout_scale_now = 10
                End If
            Else
                If (check_Force_CCM.Checked = True) Then
                    vout_scale_now = num_vout_CCM.Value
                Else
                    vout_scale_now = num_vout_DEM.Value
                End If
            End If
            CHx_scale(vout_ch, vout_scale_now, "mV")
            If DUT2_en Then
                CHx_scale(vout2_ch, vout_scale_now, "mV")
            End If

            monitor_count(10, False, "Part I", meas1)

            If rbtn_auto_vout.Checked = True Then
                '第一次調整vout scale

                vpp(3) = Scope_measure(meas4, Meas_max)
                If DUT2_en Then
                    vpp2(3) = Scope_measure(meas8, Meas_max)
                End If

                vout_temp = vpp(3) * (10 ^ 3) / num_vout_auto.Value   'mV
                If DUT2_en Then
                    vout_temp = vpp2(3) * (10 ^ 3) / num_vout_auto.Value   'mV
                End If
                'Math.Ceiling() 無條件進位, Math.Floor() 捨去小數
                vout_temp = Math.Floor(vout_temp)
                If vout_temp < 5 Then
                    vout_scale_temp = vout_temp
                Else
                    vout_scale_temp = Math.Floor(vout_temp / 5) * 5
                End If
                VoutScalling_CCM = False
            Else
                If (Fs_CCM = True) Then
                    vout_scale_temp = num_vout_CCM.Value
                    VoutScalling_CCM = True
                Else
                    vout_scale_temp = num_vout_DEM.Value
                    VoutScalling_CCM = False
                End If
            End If
            If vout_scale_temp <> vout_scale_now Then
                vout_scale_now = vout_scale_temp
                CHx_scale(vout_ch, vout_scale_now, "mV") 'Voltage Scale > VID * 10% / 4
                System.Threading.Thread.Sleep(500)
                If DUT2_en Then
                    CHx_scale(vout2_ch, vout_scale_now, "mV") 'Voltage Scale > VID * 10% / 4
                End If
            End If
        End If

        ' auto scalling function
        If ((iout_now > 0) And (AutoScalling_EN = True) And (Fs_CCM = False)) Or (rbtn_auto_all.Checked = True) Then
            ' 0: Ton, 1: Toff, 2: Freq
            wave_data = Auto_Scanning(lx_ch)
            If DUT2_en Then
                wave2_data = Auto_Scanning(lx2_ch)
            End If
            'System.Threading.Thread.Sleep(1000)
            If wave_data(0) <> 0 Then
                autoscanning_update = True
            Else
                autoscanning_update = False
            End If
        End If

        Stability_Measure_Set(0)
        If Fs_CCM = True Then
            monitor_count(cnt, True, "Part I")
        Else
            monitor_count(cnt, True, "Part I")
        End If
        Stability_Measure_Get(0)
        update_report(Stability)
        Console.WriteLine("DUT1 Vout Max = {0}, Vout Min = {1}", vpp(4), vpp(5))

        H_reclength(50000)

        If DUT2_en Then
            Stability_Measure_Set(1)
            If Fs_CCM = True Then
                monitor_count(num_counts_CCM.Value, True, "Part I")
            Else
                monitor_count(num_counts_DEM.Value, True, "Part I")
            End If
            Stability_Measure_Get(1)
            update_report(Stability, DUT2_en)
            Console.WriteLine("DUT2 Vout Max = {0}, Vout Min = {1}", vpp(4), vpp(5))
        End If
        iout_temp = iout_now
        Scope_measure_reset()
    End Sub


    Function Stability_run() As Integer
        Dim vout_temp As Double
        Dim vout_scale_temp As Integer
        Dim iout_temp As Double
        Dim double_check As Boolean = False

        'System.Threading.Thread.Sleep(1000)
        ' all channel enable
        If DUT2_en Then
            CHx_display(1, "ON")
            CHx_display(2, "ON")
            CHx_display(3, "ON")
            CHx_display(4, "ON")
        Else
            CHx_display(lx_ch, "ON")
            CHx_display(vout_ch, "ON")
        End If
        scope_time_init()
        Scope_RUN(False)
        Calculate_pass(TA_Test_num)
        If (check_Force_CCM.Checked = False) And (iout_now = 0) Then
            H_scale_value = ((1 / Fs_Min) * 10 / 10) * (10 ^ 9) '1/Fs_Min(Hz)*n/10 
        Else
            H_scale_value = ((1 / Fs_Min) * Wave_num / 10) * (10 ^ 9) '1/Fs_Min(Hz)*n/10 
        End If
        'Timing Scale
        H_scale(H_scale_value, "ns") ' 1 / Fs_Min(Hz) * n / 10 
        If RS_Scope = True Then
            RS_View(True)
        End If

        ' vout level scale re-size
        If (iout_now = data_test.Rows(0).Cells(0).Value) Then
            If rbtn_auto_vout.Checked = True Then
                vout_temp = Math.Floor(((vout_meas * 1000) * 0.005))
                vout_temp = Math.Floor(vout_temp / 5) * 5
                If vout_temp > 10 Then
                    vout_scale_now = vout_temp
                Else
                    vout_scale_now = 10
                End If
            Else
                If (check_Force_CCM.Checked = True) Then
                    vout_scale_now = num_vout_CCM.Value
                Else
                    vout_scale_now = num_vout_DEM.Value
                End If
            End If
            CHx_scale(vout_ch, vout_scale_now, "mV")
            If DUT2_en Then
                CHx_scale(vout2_ch, vout_scale_now, "mV")
            End If

            monitor_count(10, False, "Part I", meas1)

            If rbtn_auto_vout.Checked = True Then
                '第一次調整vout scale

                vpp(3) = Scope_measure(meas4, Meas_max)
                If DUT2_en Then
                    vpp2(3) = Scope_measure(meas8, Meas_max)
                End If

                vout_temp = vpp(3) * (10 ^ 3) / num_vout_auto.Value   'mV
                If DUT2_en Then
                    vout_temp = vpp2(3) * (10 ^ 3) / num_vout_auto.Value   'mV
                End If
                'Math.Ceiling() 無條件進位, Math.Floor() 捨去小數
                vout_temp = Math.Floor(vout_temp)
                If vout_temp < 5 Then
                    vout_scale_temp = vout_temp
                Else
                    vout_scale_temp = Math.Floor(vout_temp / 5) * 5
                End If
                VoutScalling_CCM = False
            Else
                If (Fs_CCM = True) Then
                    vout_scale_temp = num_vout_CCM.Value
                    VoutScalling_CCM = True
                Else
                    vout_scale_temp = num_vout_DEM.Value
                    VoutScalling_CCM = False
                End If
            End If
            If vout_scale_temp <> vout_scale_now Then
                vout_scale_now = vout_scale_temp
                CHx_scale(vout_ch, vout_scale_now, "mV") 'Voltage Scale > VID * 10% / 4
                System.Threading.Thread.Sleep(500)
                If DUT2_en Then
                    CHx_scale(vout2_ch, vout_scale_now, "mV") 'Voltage Scale > VID * 10% / 4
                End If
            End If
        End If

        ' auto scalling function
        If ((iout_now > 0) And (AutoScalling_EN = True) And (Fs_CCM = False)) Or (rbtn_auto_all.Checked = True) Then
            ' 0: Ton, 1: Toff, 2: Freq
            wave_data = Auto_Scanning(lx_ch)
            If DUT2_en Then
                wave2_data = Auto_Scanning(lx2_ch)
            End If
            'System.Threading.Thread.Sleep(1000)
            If wave_data(0) <> 0 Then
                autoscanning_update = True
            Else
                autoscanning_update = False
            End If
        End If

        Stability_Measure_Set(0)
        If Fs_CCM = True Then
            monitor_count(num_counts_CCM.Value, True, "Part I")
        Else
            monitor_count(num_counts_DEM.Value, True, "Part I")
        End If
        Stability_Measure_Get(0)
        update_report(Stability)


        Console.WriteLine("DUT1 Vout Max = {0}, Vout Min = {1}", vpp(4), vpp(5))


        H_reclength(50000)

        If DUT2_en Then
            Stability_Measure_Set(1)
            If Fs_CCM = True Then
                monitor_count(num_counts_CCM.Value, True, "Part I")
            Else
                monitor_count(num_counts_DEM.Value, True, "Part I")
            End If
            Stability_Measure_Get(1)
            update_report(Stability, DUT2_en)

            Console.WriteLine("DUT2 Vout Max = {0}, Vout Min = {1}", vpp(4), vpp(5))
        End If
        iout_temp = iout_now
        Scope_measure_reset()

        ' get measure data
        'If DUT2_en Then
        '    ' freq KHz
        '    fs2(0) = Scope_measure(meas1, Scope_Meas)
        '    fs2(1) = Scope_measure(meas1, Meas_mean)
        '    fs2(2) = Scope_measure(meas1, Meas_min)
        '    fs2(3) = Scope_measure(meas1, Meas_max)
        '    'Ton (ns)
        '    ton2(0) = Scope_measure(meas2, Scope_Meas)
        '    ton2(1) = Scope_measure(meas2, Meas_mean)
        '    ton2(2) = Scope_measure(meas2, Meas_min)
        '    ton2(3) = Scope_measure(meas2, Meas_max)
        '    'Toff
        '    toff2(0) = Scope_measure(meas3, Scope_Meas)
        '    toff2(1) = Scope_measure(meas3, Meas_mean)
        '    toff2(2) = Scope_measure(meas3, Meas_min)
        '    toff2(3) = Scope_measure(meas3, Meas_max)
        '    'vpp
        '    vpp2(0) = Scope_measure(meas4, Scope_Meas)
        '    vpp2(1) = Scope_measure(meas4, Meas_mean)
        '    vpp2(2) = Scope_measure(meas4, Meas_min)
        '    vpp2(3) = Scope_measure(meas4, Meas_max)
        '    ' freq KHz
        '    fs(0) = Scope_measure(meas5, Scope_Meas)
        '    fs(1) = Scope_measure(meas5, Meas_mean)
        '    fs(2) = Scope_measure(meas5, Meas_min)
        '    fs(3) = Scope_measure(meas5, Meas_max)
        '    'Ton (ns)
        '    ton(0) = Scope_measure(meas6, Scope_Meas)
        '    ton(1) = Scope_measure(meas6, Meas_mean)
        '    ton(2) = Scope_measure(meas6, Meas_min)
        '    ton(3) = Scope_measure(meas6, Meas_max)
        '    'Toff
        '    toff(0) = Scope_measure(meas7, Scope_Meas)
        '    toff(1) = Scope_measure(meas7, Meas_mean)
        '    toff(2) = Scope_measure(meas7, Meas_min)
        '    toff(3) = Scope_measure(meas7, Meas_max)
        '    'vpp
        '    vpp(0) = Scope_measure(meas8, Scope_Meas)
        '    vpp(1) = Scope_measure(meas8, Meas_mean)
        '    vpp(2) = Scope_measure(meas8, Meas_min)
        '    vpp(3) = Scope_measure(meas8, Meas_max)
        'Else
        '    ' freq KHz
        '    fs(0) = Scope_measure(meas1, Scope_Meas)
        '    fs(1) = Scope_measure(meas1, Meas_mean)
        '    fs(2) = Scope_measure(meas1, Meas_min)
        '    fs(3) = Scope_measure(meas1, Meas_max)
        '    'Ton (ns)
        '    ton(0) = Scope_measure(meas2, Scope_Meas)
        '    ton(1) = Scope_measure(meas2, Meas_mean)
        '    ton(2) = Scope_measure(meas2, Meas_min)
        '    ton(3) = Scope_measure(meas2, Meas_max)
        '    'Toff
        '    toff(0) = Scope_measure(meas3, Scope_Meas)
        '    toff(1) = Scope_measure(meas3, Meas_mean)
        '    toff(2) = Scope_measure(meas3, Meas_min)
        '    toff(3) = Scope_measure(meas3, Meas_max)
        '    'vpp
        '    vpp(0) = Scope_measure(meas4, Scope_Meas)
        '    vpp(1) = Scope_measure(meas4, Meas_mean)
        '    vpp(2) = Scope_measure(meas4, Meas_min)
        '    vpp(3) = Scope_measure(meas4, Meas_max)

        '    'Vmax
        '    vpp(4) = Scope_measure(meas5, Meas_max)

        '    'Vmin
        '    vpp(5) = Scope_measure(meas6, Meas_min)
        'End If


        'If DUT2_en Then
        '    ' 0: vmax, 1: min
        '    'vout = GetVoutMax_Min(vout_ch)
        '    'vout2 = GetVoutMax_Min(vout2_ch)
        '    'scope_time_init()
        '    'Display_persistence(False)
        '    ' dut data to excel
        '    vpp(4) = vout(0)
        '    vpp(5) = vout(1)
        '    update_report(Stability)
        '    ' dut2 data to excel
        '    Array.Copy(wave2_data, wave_data, wave2_data.Length)
        '    Array.Copy(fs2, fs, fs2.Length)
        '    Array.Copy(toff2, toff, toff2.Length)
        '    Array.Copy(ton2, ton, ton2.Length)
        '    Array.Copy(vpp2, vpp, vpp2.Length)
        '    vpp(4) = vout2(0)
        '    vpp(5) = vout2(1)
        '    update_report(Stability, DUT2_en)
        'Else
        '    update_report(Stability)
        'End If
        'scope_time_init()
        'Scope_measure_reset()

        'Display_persistence(False)


        'Scope_RUN(True)

        'Scope_measure_reset()
        'Dim cnt = 0
        'While (cnt <= 0)
        '    cnt = Scope_measure_count(1)
        'End While

        'For i = 0 To test_cnt
        '    'If i = 0 Then
        '    '    CHx_display(lx_ch, "ON")
        '    '    CHx_display(vout_ch, "ON")
        '    '    Trigger_auto_level(lx_ch, "R")
        '    'Else
        '    '    CHx_display(lx2_ch, "ON")
        '    '    CHx_display(vout2_ch, "ON")
        '    '    Trigger_auto_level(lx2_ch, "R")
        '    'End If

        '    Scope_RUN(False)
        '    Calculate_pass(TA_Test_num)
        '    If (check_Force_CCM.Checked = False) And (iout_now = 0) Then
        '        H_scale_value = ((1 / Fs_Min) * 10 / 10) * (10 ^ 9) '1/Fs_Min(Hz)*n/10 
        '    Else
        '        H_scale_value = ((1 / Fs_Min) * Wave_num / 10) * (10 ^ 9) '1/Fs_Min(Hz)*n/10 
        '    End If
        '    'Timing Scale
        '    H_scale(H_scale_value, "ns") '1/Fs_Min(Hz)*n/10 
        '    If RS_Scope = True Then
        '        RS_View(True)
        '    End If


        '    If (iout_now = data_test.Rows(0).Cells(0).Value) Then
        '        If rbtn_auto_vout.Checked = True Then
        '            vout_temp = Math.Floor(((vout_meas * 1000) * 0.005))
        '            vout_temp = Math.Floor(vout_temp / 5) * 5
        '            If vout_temp > 10 Then
        '                vout_scale_now = vout_temp
        '            Else
        '                vout_scale_now = 10
        '            End If
        '        Else
        '            If (check_Force_CCM.Checked = True) Then
        '                vout_scale_now = num_vout_CCM.Value
        '            Else
        '                vout_scale_now = num_vout_DEM.Value
        '            End If
        '        End If
        '        CHx_scale(vout_list(i), vout_scale_now, "mV") 'Voltage Scale > VID * 10% / 4
        '        monitor_count(num_counts_CCM.Value, False, "Part I")
        '        If rbtn_auto_vout.Checked = True Then
        '            '第一次調整vout scale
        '            If i = 0 Then
        '                vpp(3) = Scope_measure(meas4, Meas_max)
        '            Else
        '                vpp(3) = Scope_measure(meas8, Meas_max)
        '            End If
        '            vout_temp = vpp(3) * (10 ^ 3) / num_vout_auto.Value   'mV
        '            'Math.Ceiling() 無條件進位, Math.Floor() 捨去小數
        '            vout_temp = Math.Floor(vout_temp)
        '            If vout_temp < 5 Then
        '                vout_scale_temp = vout_temp
        '            Else
        '                vout_scale_temp = Math.Floor(vout_temp / 5) * 5
        '            End If
        '            VoutScalling_CCM = False
        '        Else
        '            If (Fs_CCM = True) Then
        '                vout_scale_temp = num_vout_CCM.Value
        '                VoutScalling_CCM = True
        '            Else
        '                vout_scale_temp = num_vout_DEM.Value
        '                VoutScalling_CCM = False
        '            End If
        '        End If
        '        If vout_scale_temp <> vout_scale_now Then
        '            vout_scale_now = vout_scale_temp
        '            CHx_scale(vout_list(i), vout_scale_now, "mV") 'Voltage Scale > VID * 10% / 4
        '        End If
        '    End If

        '    ' first one get data
        '    If Fs_CCM = True Then
        '        monitor_count(num_counts_CCM.Value, True, "Part I")
        '    Else
        '        monitor_count(num_counts_DEM.Value, True, "Part I")
        '    End If


        '    If DUT2_en Then
        '        vpp_max = Scope_measure(meas4, Meas_max)
        '        vpp2_max = Scope_measure(meas8, Meas_max)
        '        pass_value_Max = vout_now * (num_vout_ac.Value / 100)
        '        If vpp_max > pass_value_Max Then
        '            dut1_state = FAIL
        '        End If

        '        If vpp2_max > pass_value_Max Then
        '            dut2_state = FAIL
        '        End If

        '        If dut1_state Then
        '            Scope_measure_set(meas7, lx_ch, "MAXimum")
        '            Scope_measure_set(meas8, lx_ch, "MINImum")
        '            If Fs_CCM = True Then
        '                monitor_count(num_counts_CCM.Value, True, "Part I")
        '            Else
        '                monitor_count(num_counts_DEM.Value, True, "Part I")
        '            End If

        '            vmax = Scope_measure(meas7, Meas_max)
        '            vmin = Scope_measure(meas8, Meas_min)
        '        End If


        '        If dut2_state Then
        '            Scope_measure_set(meas3, lx_ch, "MAXimum")
        '            Scope_measure_set(meas4, lx_ch, "MINImum")
        '            If Fs_CCM = True Then
        '                monitor_count(num_counts_CCM.Value, True, "Part I")
        '            Else
        '                monitor_count(num_counts_DEM.Value, True, "Part I")
        '            End If

        '            vmax2 = Scope_measure(meas3, Meas_max)
        '            vmin2 = Scope_measure(meas4, Meas_min)
        '        End If

        '        ' freq
        '        Scope_measure_set(meas1, lx_ch, "FREQuency")
        '        ' Pwidth
        '        Scope_measure_set(meas2, lx_ch, "PWIdth")
        '        ' NWidth
        '        Scope_measure_set(meas3, lx_ch, "NWIdth")
        '        ' PK2PK
        '        Scope_measure_set(meas4, vout_ch, "PK2Pk")

        '        ' freq
        '        Scope_measure_set(meas5, lx2_ch, "FREQuency")
        '        ' Pwidth
        '        Scope_measure_set(meas6, lx2_ch, "PWIdth")
        '        ' NWidth
        '        Scope_measure_set(meas7, lx2_ch, "NWIdth")
        '        ' PK2PK
        '        Scope_measure_set(meas8, vout2_ch, "PK2Pk")

        '    End If

        '    If rbtn_auto_vout.Checked = True Then
        '        '--------------------------------------------------------------------
        '        '計算Scale
        '        vout_temp = vpp(3) * (10 ^ 3) / num_vout_auto.Value   'mV
        '        'Math.Ceiling() 無條件進位, Math.Floor() 捨去小數
        '        vout_temp = Math.Floor(vout_temp)
        '        If vout_temp < 5 Then
        '            vout_scale_temp = vout_temp
        '        Else
        '            vout_scale_temp = Math.Floor(vout_temp / 5) * 5
        '        End If
        '        '--------------------------------------------------------------------
        '        If Check_fixed.Checked = True Then
        '            '--------------------------------------------------------------------
        '            '不管CCM. DEM都固定調整在同一個隔數內
        '            If vout_scale_temp <> vout_scale_now Then
        '                vout_scale_now = vout_scale_temp
        '                CHx_scale(vout_list(i), vout_scale_now, "mV") 'Voltage Scale > VID * 10% / 4
        '            End If
        '            '----------------------------------------------------------------------
        '        Else
        '            ''只調整一次
        '            If (check_Force_CCM.Checked = False) Then
        '                If (iout_now <= IOUT_Boundary_Stop) And (iout_now >= IOUT_Boundary_Start) Then
        '                    double_check = False
        '                    'Iout上升
        '                    If (iout_now >= iout_temp) And (VoutScalling_CCM = False) And (vout_scale_temp < vout_scale_now) Then
        '                        vout_scale_now = vout_scale_temp
        '                        CHx_scale(vout_list(i), vout_scale_now, "mV") 'Voltage Scale > VID * 10% / 4
        '                        VoutScalling_CCM = True
        '                    End If
        '                    'Iout下降
        '                    If (iout_now < iout_temp) And (VoutScalling_CCM = True) And (vout_scale_temp > vout_scale_now) Then
        '                        vout_scale_now = vout_scale_temp
        '                        CHx_scale(vout_list(i), vout_scale_now, "mV") 'Voltage Scale > VID * 10% / 4
        '                        VoutScalling_CCM = False
        '                    End If
        '                ElseIf (iout_now >= IOUT_Boundary_Stop) And (double_check = False) Then
        '                    vout_scale_now = vout_scale_temp
        '                    CHx_scale(vout_list(i), vout_scale_now, "mV") 'Voltage Scale > VID * 10% / 4
        '                    double_check = True
        '                ElseIf (iout_now <= IOUT_Boundary_Start) And (double_check = False) Then
        '                    vout_scale_now = vout_scale_temp
        '                    CHx_scale(vout_list(i), vout_scale_now, "mV") 'Voltage Scale > VID * 10% / 4
        '                    double_check = True
        '                End If
        '            End If
        '        End If
        '    End If

        '    iout_temp = iout_now
        '    ' ''----------------------------------------------------------------------------------
        '    ''Measurement
        '    ''Scope
        '    If i = 0 Then
        '        update_report(Stability)
        '    Else
        '        update_report(Stability, DUT2_en)
        '    End If

        'Next
    End Function

    Function Jitter_run() As Integer
        Dim vout_temp As Double
        Dim vout_scale_temp As Integer
        Dim lx_list() As Integer = New Integer() {lx_ch, lx2_ch}
        Dim vout_list() As Integer = New Integer() {vout_ch, vout2_ch}
        Dim test_cnt As Integer = 0
        If DUT2_en Then
            test_cnt = 1
        End If

        For i = 0 To test_cnt

            CHx_display(1, "OFF")
            CHx_display(2, "OFF")
            CHx_display(3, "OFF")
            CHx_display(4, "OFF")

            If i = 0 Then
                CHx_display(lx_ch, "ON")
                'CHx_display(vout_ch, "ON")
                Trigger_auto_level(lx_ch, "R")
            Else
                CHx_display(lx2_ch, "ON")
                'CHx_display(vout2_ch, "ON")
                Trigger_auto_level(lx2_ch, "R")
            End If


            Scope_RUN(False)
            If check_cursors.Checked = True Then
                Cursor_ONOFF("OFF")
            End If
            If RS_Scope = True Then
                RS_Display(RS_RES_MES, RS_DISP_DOCK)
                'RS_Display(RS_RES_MES, RS_DISP_PREV)
                Scope_measure_clear()

                If i = 0 Then
                    RS_Scope_measure_status(meas1, True) ' freq
                    RS_Scope_measure_status(meas2, True) ' ton
                    RS_Scope_measure_status(meas3, True) ' toff
                    RS_Scope_measure_status(meas4, True) ' vpp
                Else
                    RS_Scope_measure_status(meas5, True) ' freq
                    RS_Scope_measure_status(meas6, True) ' ton
                    RS_Scope_measure_status(meas7, True) ' toff
                    RS_Scope_measure_status(meas8, True) ' vpp
                End If
                RS_Local()
                'RS_View()
            End If
            '-------------------------------------------------------------------------
            H_scale_value = ((1 / fs_now) * 2 / 8) * (10 ^ 9)
            'Timing Scale
            H_scale(H_scale_value, "ns") '1/Fs_Min(Hz)*n/10 
            If rbtn_vin_trigger.Checked = True Then
                Trigger_set(lx_list(i), "R", vin_now / num_vin_trigger.Value)
            Else
                Trigger_auto_level(lx_list(i), "R")
            End If
            If RS_Scope = True Then
                RS_View(True)
            End If
            RUN_set("RUNSTop")
            If (iout_now = data_jitter_iout.Rows(0).Cells(0).Value) Then
                If rbtn_auto_vout.Checked = True Then
                    '----------------------------------------------------------------------
                    If i = 1 Then
                        monitor_count(num_counts_CCM.Value, True, "Part I", meas5)
                    Else
                        monitor_count(num_counts_CCM.Value, True, "Part I")
                    End If


                    If i = 0 Then
                        vpp(3) = Scope_measure(meas4, Meas_max)
                    Else
                        vpp(3) = Scope_measure(meas8, Meas_max)
                    End If

                    If vpp(3) < (10 ^ 20) Then
                        vout_temp = vpp(3) * (10 ^ 3) / num_vout_auto.Value   'mV
                        'Math.Ceiling() 無條件進位, Math.Floor() 捨去小數
                        vout_temp = Math.Floor(vout_temp)
                        If vout_temp < 5 Then
                            vout_scale_temp = vout_temp
                        Else
                            vout_scale_temp = Math.Floor(vout_temp / 5) * 5
                        End If
                        If vout_scale_temp <> vout_scale_now Then
                            vout_scale_now = vout_scale_temp
                            CHx_scale(vout_list(i), vout_scale_now, "mV") 'Voltage Scale > VID * 10% / 4
                        End If
                    End If
                Else
                    vout_scale_now = num_vout_CCM.Value
                    CHx_scale(vout_list(i), vout_scale_now, "mV") 'Voltage Scale > VID * 10% / 4
                End If
            End If
            '無限持續累積
            If check_persistence.Checked = True Then
                Scope_RUN(False)
                Display_persistence(True)
                Delay(100)
            End If
            Scope_RUN(True)
            Delay(100)
            If i = 1 Then
                monitor_count(num_counts_Jitter.Value, True, "Part I", meas5)
            Else
                monitor_count(num_counts_Jitter.Value, True, "Part I")
            End If
            If i = 0 Then
                fs(2) = Scope_measure(meas1, Meas_min)          ' freq
                ton(1) = Scope_measure(meas2, Meas_mean)        ' Ton (ns)
                toff(2) = Scope_measure(meas3, Meas_min)        ' Toff
                toff(3) = Scope_measure(meas3, Meas_max)        ' Toff
            Else
                fs(2) = Scope_measure(meas5, Meas_min)
                ton(1) = Scope_measure(meas6, Meas_mean)
                toff(2) = Scope_measure(meas7, Meas_min)
                toff(3) = Scope_measure(meas7, Meas_max)
            End If

            If i = 0 Then
                update_report(Jitter)
            Else
                update_report(Jitter, DUT2_en)
            End If
        Next


    End Function

    Function update_report(ByVal test_name As String, Optional ByVal dut2_en As Boolean = False) As Integer

        Dim test_cnt As Integer = 0
        If dut2_en Then
            test_cnt = 1
        End If
        Dim note() As String
        Dim first_row As Integer 'title
        Dim start_col As Integer
        Dim col_num, row_num As Integer
        Dim set_num As Integer
        'Dim wave_data() As Double 'Ton(ns),Toff(ns),Freq(KHz)
        'Dim wave_data2() As Double
        Dim temp As String
        Dim beta_path As String
        Dim row_num_temp As Integer
        'Jitter
        Dim Dave, Dmax, Dmin As Double
        Dim Tjitter, Ton_mean, Toff_max, Toff_min As Double
        Dim Jitter_value As Double
        Dim pass_result As String
        Dim eff As Double
        Dim Hyperlinks_txt As String = ""
        Dim error_pic_path As String
        Dim i, ii As Integer
        col_num = 0
        row_num = 0
        If run = False Then
            Exit Function
        End If
        'Start Test
        Select Case test_name
            Case Stability
                If check_stability_pic.Checked = True Then
                    '將每一張scope的圖都儲存
                    If dut2_en Then
                        beta_path = Beta_folder2 & "\" & beta_pic_num & "_" & "Ta=" & TA_now & "; Fs=" & fs_now & "Hz; Vout=" & vout_now & "V; Vin=" & vin_now & "V; Iout=" & iout_now & "A" & add_dut2 & ".PNG"
                    Else
                        beta_path = Beta_folder & "\" & beta_pic_num & "_" & "Ta=" & TA_now & "; Fs=" & fs_now & "Hz; Vout=" & vout_now & "V; Vin=" & vin_now & "V; Iout=" & iout_now & "A" & ".PNG"
                    End If
                    Hardcopy("PNG", beta_path)
                    beta_pic_num = beta_pic_num + 1
                End If

                If dut2_en Then
                    xlSheet = xlBook.Sheets(txt_stability_sheet.Text & add_dut2)
                Else
                    xlSheet = xlBook.Sheets(txt_stability_sheet.Text)
                End If
                xlSheet.Activate()
                'If Not (dut2_en) Then
                '    If ((iout_now > 0) And (AutoScalling_EN = True) And (Fs_CCM = False)) Or (rbtn_auto_all.Checked = True) Then
                '        '--------------------------------------------------------
                '        wave_data = Auto_Scanning(lx_ch)
                '        If wave_data(0) <> 0 Then
                '            autoscanning_update = True
                '        Else
                '            autoscanning_update = False
                '        End If
                '    End If
                'End If
                'xlSheet.Activate()
                '----------------------------------------------------------------------------------
                'initial
                'Init col
                col_num = stability_col.Length
                start_col = test_col + chart_width + col_Space + (TA_Test_num * total_vcc.Length * total_fs.Length + VCC_test_num * total_fs.Length + fs_test_num) * (col_num + 1)
                first_row = stability_report_row(data_set_now)
                '----------------------------------------------------------------------------------
                'Update Vin
                row = first_row + 2 + stability_iout_num
                col = start_col
                xlrange = xlSheet.Range(ConvertToLetter(col) & row)

                If dut2_en Then
                    xlrange.Value = vout_meas2
                    If vout_meas2 < (vout_now * (vout_err / 100)) Then
                        xlrange.Interior.Color = 255
                    End If
                Else
                    xlrange.Value = vout_meas
                    If vout_meas < (vout_now * (vout_err / 100)) Then
                        xlrange.Interior.Color = 255
                    End If
                End If
                FinalReleaseComObject(xlrange)
                col = col + 1
                xlrange = xlSheet.Range(ConvertToLetter(col) & row)
                xlrange.Value = iout_now
                FinalReleaseComObject(xlrange)
                col = col + 1

                xlrange = xlSheet.Range(ConvertToLetter(col) & row)
                xlrange.Value = Fs_Max / 1000
                FinalReleaseComObject(xlrange)
                col = col + 1

                xlrange = xlSheet.Range(ConvertToLetter(col) & row)
                xlrange.Value = Fs_Min / 1000
                FinalReleaseComObject(xlrange)
                col = col + 1

                '-------------------------------------------------------------------------
                'freq
                For ii = 0 To 3
                    xlrange = xlSheet.Range(ConvertToLetter(col) & row)
                    xlrange.Value = fs(ii) / (10 ^ 3) ' Format(fs(ii) / (10 ^ 3), "#0.000")
                    FinalReleaseComObject(xlrange)
                    col = col + 1
                Next
                xlrange = xlSheet.Range(ConvertToLetter(col) & row)
                'freq_update
                If (AutoScalling_EN = True) Then
                    If autoscanning_update = True Then
                        xlrange.Interior.Color = 49407 '橘色
                        fs_update = wave_data(2)
                        xlrange.Value = fs_update / (10 ^ 3)
                    Else
                        If (Fs_CCM = False) Then
                            xlrange.Interior.Color = 255
                        End If
                        fs_update = fs(1)
                        xlrange.Value = fs_update / (10 ^ 3)
                    End If
                End If
                If (check_Force_CCM.Checked = True) And (rbtn_auto_DEM.Checked = True) Then
                    'fs(2) = Scope_measure(x, Meas_min)
                    'fs(3) = Scope_measure(x, Meas_max)
                    If (fs(2) >= Fs_Min) And (fs(3) <= Fs_Max) Then
                        pass_result = PASS
                    Else
                        pass_result = FAIL
                    End If
                Else
                    If (fs_update >= Fs_Min) And (fs_update <= Fs_Max) Then
                        pass_result = PASS
                    Else
                        pass_result = FAIL
                    End If
                End If
                FinalReleaseComObject(xlrange)
                col = col + 1
                '----------------------------------------------------------
                'Vout Scale (Manual Mode)
                If rbtn_manual_vout.Checked = True Then
                    If (check_Force_CCM.Checked = False) Then
                        'Iout上升
                        If (iout_now >= IOUT_Boundary_Start) And (VoutScalling_CCM = False) Then
                            If ((fs_update >= Fs_Min) And (fs_update <= Fs_Max)) Then
                                vout_scale_now = num_vout_CCM.Value
                                If dut2_en Then
                                    CHx_scale(vout2_ch, vout_scale_now, "mV") 'Voltage Scale > VID * 10% / 4
                                Else
                                    CHx_scale(vout_ch, vout_scale_now, "mV") 'Voltage Scale > VID * 10% / 4
                                End If
                                VoutScalling_CCM = True
                            End If
                        End If
                        'Iout下降
                        If (iout_now < IOUT_Boundary_Stop) And (VoutScalling_CCM = True) Then
                            If (fs_update < Fs_Min) Then
                                vout_scale_now = num_vout_DEM.Value
                                If dut2_en Then
                                    CHx_scale(vout2_ch, vout_scale_now, "mV") 'Voltage Scale > VID * 10% / 4
                                Else
                                    CHx_scale(vout_ch, vout_scale_now, "mV") 'Voltage Scale > VID * 10% / 4
                                End If

                                VoutScalling_CCM = False
                            End If
                        End If
                    End If
                End If
                '-------------------------------------------------------------------------
                'ton
                For ii = 0 To 3
                    xlrange = xlSheet.Range(ConvertToLetter(col) & row)
                    xlrange.Value = ton(ii) * (10 ^ 9) ' Format(ton(ii) * (10 ^ 9), "#0.000")
                    FinalReleaseComObject(xlrange)
                    col = col + 1
                Next
                xlrange = xlSheet.Range(ConvertToLetter(col) & row)
                'ton_update
                If (AutoScalling_EN = True) Then
                    If autoscanning_update = True Then
                        xlrange.Interior.Color = 49407
                        xlrange.Value = wave_data(0) * (10 ^ 9)
                    Else
                        If (Fs_CCM = False) Then
                            xlrange.Interior.Color = 255
                        End If
                        xlrange.Value = ton(1) * (10 ^ 9)
                    End If
                End If
                FinalReleaseComObject(xlrange)
                col = col + 1
                '-------------------------------------------------------------------------
                'toff
                For ii = 0 To 3
                    xlrange = xlSheet.Range(ConvertToLetter(col) & row)
                    xlrange.Value = toff(ii) * (10 ^ 9) ' Format(toff(ii) * (10 ^ 9), "#0.000")
                    FinalReleaseComObject(xlrange)
                    col = col + 1
                Next
                xlrange = xlSheet.Range(ConvertToLetter(col) & row)
                'toff_update
                If (AutoScalling_EN = True) Then
                    If (autoscanning_update = True) Then
                        xlrange.Interior.Color = 49407
                        xlrange.Value = wave_data(1) * (10 ^ 9) ' Format(wave_data(1) * (10 ^ 9), "#0.000")
                    Else
                        If (Fs_CCM = False) Then
                            xlrange.Interior.Color = 255
                        End If
                        xlrange.Value = toff(1) * (10 ^ 9) ' Format(wave_data(1) * (10 ^ 9), "#0.000")
                    End If
                End If
                FinalReleaseComObject(xlrange)
                col = col + 1
                '-------------------------------------------------------------------------
                'vpp
                For ii = 0 To 3
                    xlrange = xlSheet.Range(ConvertToLetter(col) & row)
                    xlrange.Value = vpp(ii) * (10 ^ 3) ' Format(vpp(ii) * (10 ^ 3), "#0.000")
                    FinalReleaseComObject(xlrange)
                    col = col + 1
                Next


                'Vmax
                xlrange = xlSheet.Range(ConvertToLetter(col) & row)
                xlrange.Value = vpp(4) ' Format(vpp(4), "#0.000")
                FinalReleaseComObject(xlrange)
                col = col + 1
                'Vmin
                xlrange = xlSheet.Range(ConvertToLetter(col) & row)
                xlrange.Value = vpp(5) ' Format(vpp(5), "#0.000")
                FinalReleaseComObject(xlrange)

                col = col + 1
                If pass_result = PASS Then
                    If dut2_en Then
                        If cbox_coupling_vout2.SelectedItem = "AC" Then
                            pass_value_Max = vout_now * (num_vout_ac.Value / 100)
                            If (vpp(3) > pass_value_Max) Then
                                pass_result = FAIL
                            End If
                        Else
                            pass_value_Max = vout_now * (1 + num_vout_pos.Value / 100)
                            pass_value_Min = vout_now * (1 - num_vout_neg.Value / 100)
                            If (vpp(5) < pass_value_Min) Or (vpp(4) > pass_value_Max) Then
                                pass_result = FAIL
                            End If
                        End If
                    Else
                        If cbox_coupling_vout.SelectedItem = "AC" Then
                            pass_value_Max = vout_now * (num_vout_ac.Value / 100)
                            If (vpp(3) > pass_value_Max) Then
                                pass_result = FAIL
                            End If
                        Else
                            pass_value_Max = vout_now * (1 + num_vout_pos.Value / 100)
                            pass_value_Min = vout_now * (1 - num_vout_neg.Value / 100)
                            If (vpp(5) < pass_value_Min) Or (vpp(4) > pass_value_Max) Then
                                pass_result = FAIL
                            End If
                        End If
                    End If
                End If

                xlrange = xlSheet.Range(ConvertToLetter(col) & row)
                If pass_result = FAIL Then
                    xlrange.Interior.Color = test_fail_color
                End If
                xlrange.Value = pass_result ' Format(vpp(5), "#0.000")
                FinalReleaseComObject(xlrange)
                col = col + 1
                'temp = xlSheet.Range(ConvertToLetter(col) & row).Value()
                If pass_result = FAIL Then
                    If Error_folder = "" Then
                        Error_folder = folderPath & "\Error_" & DateTime.Now.ToString("MMdd") & "_" & DateTime.Now.ToString("HHmmss")
                        My.Computer.FileSystem.CreateDirectory(Error_folder)
                    End If


                    If Error_folder2 = "" And dut2_en Then
                        Error_folder2 = folderPath & "\Error_" & DateTime.Now.ToString("MMdd") & add_dut2 & DateTime.Now.ToString("HHmmss")
                        My.Computer.FileSystem.CreateDirectory(Error_folder2)
                    End If

                    If dut2_en Then
                        Hyperlinks_txt = "#" & error_pic_num2
                    Else
                        Hyperlinks_txt = "#" & error_pic_num
                    End If

                    If dut2_en Then
                        xlSheet = xlBook.Sheets(txt_error_sheet.Text & add_dut2)
                    Else
                        xlSheet = xlBook.Sheets(txt_error_sheet.Text)
                    End If

                    xlSheet.Activate()
                    xlrange = xlSheet.Range(ConvertToLetter(1) & 1)
                    If dut2_en Then
                        xlrange.Value = error_pic_num2
                    Else
                        xlrange.Value = error_pic_num
                    End If

                    xlBook.Save()
                    '若已經有用autoscanning矯正，就直接抓圖，不再取圖!
                    If autoscanning_update = False Then
                        'vpp(4)=Vmax (max)
                        'vpp(3)=Vpp (max)
                        '改由vout_ch以Vmax (max) - Vpp(max) * (1/10)來trigger，如果沒有偵測到在往下移
                        If dut2_en Then
                            If (cbox_coupling_vout2.SelectedItem <> "AC") And (vpp(5) < (vout_now * (1 - num_vout_neg.Value / 100))) Then
                                error_capture(vout2_ch, "R", vpp(5), True, vpp(2), num_delay_error.Value)
                            Else
                                error_capture(vout2_ch, "R", vpp(4), False, vpp(3), num_delay_error.Value)
                            End If
                        Else
                            If (cbox_coupling_vout.SelectedItem <> "AC") And (vpp(5) < (vout_now * (1 - num_vout_neg.Value / 100))) Then
                                error_capture(vout_ch, "R", vpp(5), True, vpp(2), num_delay_error.Value)
                            Else
                                error_capture(vout_ch, "R", vpp(4), False, vpp(3), num_delay_error.Value)
                            End If
                        End If
                    End If
                    '----------------------------------------------------------------------------------------------------------------
                    If dut2_en Then
                        error_pic_path = Error_folder2 & "\" & error_pic_num2 & "_" & "Ta=" & TA_now & "; Fs=" & fs_now & "Hz; Vout=" & vout_now & "V; Vin=" & vin_now & "V; Iout=" & iout_now & "A" & add_dut2 & ".PNG"
                    Else
                        error_pic_path = Error_folder & "\" & error_pic_num & "_" & "Ta=" & TA_now & "; Fs=" & fs_now & "Hz; Vout=" & vout_now & "V; Vin=" & vin_now & "V; Iout=" & iout_now & "A" & ".PNG"
                    End If
                    Hardcopy("PNG", error_pic_path)

                    If dut2_en Then
                        hyperlink_col2 = error_pic_col2
                        hyperlink_row2 = error_pic_row2
                        If (error_pic_num2 Mod 10 = 0) Then
                            error_pic_col2 = 1
                            error_pic_row2 = error_pic_row2 + pic_height + 2
                        Else
                            error_pic_col2 = error_pic_col2 + pic_width + 1
                        End If
                        error_pic_num2 = error_pic_num2 + 1
                    Else
                        hyperlink_col = error_pic_col
                        hyperlink_row = error_pic_row
                        If (error_pic_num Mod 10 = 0) Then
                            error_pic_col = 1
                            error_pic_row = error_pic_row + pic_height + 2
                        Else
                            error_pic_col = error_pic_col + pic_width + 1
                        End If
                        error_pic_num = error_pic_num + 1
                    End If

                    ''----------------------------------------------------------------------------------------------------------------
                    If check_cursors.Checked = True Then
                        Cursor_ONOFF("OFF")
                    End If
                    If dut2_en Then
                        xlSheet = xlBook.Sheets(txt_stability_sheet.Text & add_dut2)
                    Else
                        xlSheet = xlBook.Sheets(txt_stability_sheet.Text)
                    End If
                    xlSheet.Activate()
                    xlrange = xlSheet.Range(ConvertToLetter(col) & row)


                    If dut2_en Then
                        xlSheet.Hyperlinks.Add(Anchor:=xlrange, Address:="", SubAddress:=txt_error_sheet.Text & add_dut2 & "!" & ConvertToLetter(hyperlink_col) & hyperlink_row, TextToDisplay:=Hyperlinks_txt)
                    Else
                        xlSheet.Hyperlinks.Add(Anchor:=xlrange, Address:="", SubAddress:=txt_error_sheet.Text & "!" & ConvertToLetter(hyperlink_col) & hyperlink_row, TextToDisplay:=Hyperlinks_txt)
                    End If
                End If
                autoscanning_update = False
                If run = False Then
                    Exit Function
                End If

                If dut2_en Then
                    CHx_Bandwidth(lx2_ch, cbox_BW_lx.SelectedItem)
                Else
                    CHx_Bandwidth(lx_ch, cbox_BW_lx.SelectedItem)
                End If
                'H_reclength(RL_value)
                ''Timing Scale
                'H_scale(H_scale_value, "ns") '1/Fs_Min(Hz)*n/10 

                If dut2_en Then
                    If rbtn_vin_trigger.Checked = True Then
                        Trigger_set(lx2_ch, "R", vin_now / num_vin_trigger.Value)
                    Else
                        Trigger_auto_level(lx2_ch, "R")
                    End If
                Else
                    If rbtn_vin_trigger.Checked = True Then
                        Trigger_set(lx_ch, "R", vin_now / num_vin_trigger.Value)
                    Else
                        Trigger_auto_level(lx_ch, "R")
                    End If
                End If
                RUN_set("RUNSTop")
                FinalReleaseComObject(xlrange)
                FinalReleaseComObject(xlSheet)

            Case Jitter
                If dut2_en Then
                    xlSheet = xlBook.Sheets(txt_jitter_sheet.Text & add_dut2)
                Else
                    xlSheet = xlBook.Sheets(txt_jitter_sheet.Text)
                End If
                xlSheet.Activate()

                col_num = jitter_col.Length
                row_num = (data_jitter_iout.Rows.Count + 2)
                pass_value_Max = num_pass_jitter.Value
                start_col = test_col + data_jitter_iout.Rows.Count * (TA_num + 1) * total_vcc.Length * total_fs.Length * (pic_width + 1) + col_Space + (TA_Test_num * total_vcc.Length * total_fs.Length + VCC_test_num * total_fs.Length + fs_test_num) * (col_num + 1) - 1
                If check_fastAcq.Checked = True Then
                    If row_num < (2 * pic_height + 1) Then
                        first_row = test_row + Vout_test_num * (data_vin.Rows.Count * (2 * pic_height + 1 + 1) + row_Space) + Vin_test_num * (2 * pic_height + 1 + 1)
                    Else
                        first_row = test_row + Vout_test_num * (data_vin.Rows.Count * (row_num + 1) + row_Space) + Vin_test_num * (row_num + 1)
                    End If
                Else
                    If row_num < (pic_height + 1) Then
                        first_row = test_row + Vout_test_num * (data_vin.Rows.Count * (pic_height + 1 + 1) + row_Space) + Vin_test_num * (pic_height + 1 + 1)
                    Else
                        first_row = test_row + Vout_test_num * (data_vin.Rows.Count * (row_num + 1) + row_Space) + Vin_test_num * (row_num + 1)
                    End If
                End If
                col = start_col
                row = first_row + 2 + jitter_iout_num
                'Ton_mean(ns)	Toff_min(ns)	Toff_max(ns)
                xlrange = xlSheet.Range(ConvertToLetter(col) & row)
                If i = 0 Then
                    xlrange.Value = vout_meas
                    If vout_meas < (vout_now * (vout_err / 100)) Then
                        xlrange.Interior.Color = 255
                    End If
                Else
                    xlrange.Value = vout_meas2
                    If vout_meas2 < (vout_now * (vout_err / 100)) Then
                        xlrange.Interior.Color = 255
                    End If
                End If

                FinalReleaseComObject(xlrange)
                col = col + 1
                xlrange = xlSheet.Range(ConvertToLetter(col) & row)
                xlrange.Value = iout_now
                FinalReleaseComObject(xlrange)
                col = col + 1
                '"Ton_mean(ns)"
                Ton_mean = ton(1) * (10 ^ 9)
                xlrange = xlSheet.Range(ConvertToLetter(col) & row)
                xlrange.Value = Ton_mean
                FinalReleaseComObject(xlrange)
                col = col + 1
                '"Toff_min(ns)"
                Toff_min = toff(2) * (10 ^ 9)
                xlrange = xlSheet.Range(ConvertToLetter(col) & row)
                xlrange.Value = Toff_min
                FinalReleaseComObject(xlrange)
                col = col + 1
                '"Toff_max(ns)"
                Toff_max = toff(3) * (10 ^ 9)
                xlrange = xlSheet.Range(ConvertToLetter(col) & row)
                xlrange.Value = Toff_max
                FinalReleaseComObject(xlrange)
                col = col + 1
                ' "Tjitter(ns)"
                'Tjitter(ns)=Toff_max-Toff_min
                Tjitter = Toff_max - Toff_min
                xlrange = xlSheet.Range(ConvertToLetter(col) & row)
                xlrange.Value = Tjitter
                FinalReleaseComObject(xlrange)
                col = col + 1
                Dmax = Ton_mean / (Toff_min + Ton_mean)
                xlrange = xlSheet.Range(ConvertToLetter(col) & row)
                xlrange.Value = Dmax
                FinalReleaseComObject(xlrange)
                col = col + 1
                Dmin = Ton_mean / (Toff_max + Ton_mean)
                xlrange = xlSheet.Range(ConvertToLetter(col) & row)
                xlrange.Value = Dmin
                FinalReleaseComObject(xlrange)
                col = col + 1
                Dave = Ton_mean / (Toff_min + Ton_mean + (1 / 2) * Tjitter)
                xlrange = xlSheet.Range(ConvertToLetter(col) & row)
                xlrange.Value = Dave
                FinalReleaseComObject(xlrange)
                col = col + 1
                Jitter_value = 100 * (Dmax - Dmin) / Dave
                xlrange = xlSheet.Range(ConvertToLetter(col) & row)
                xlrange.Value = Jitter_value
                FinalReleaseComObject(xlrange)
                col = col + 1
                If Jitter_value < pass_value_Max Then
                    pass_result = PASS
                Else
                    pass_result = FAIL
                End If
                xlrange = xlSheet.Range(ConvertToLetter(col) & row)
                If pass_result = FAIL Then
                    xlrange.Interior.Color = test_fail_color
                End If
                xlrange.Value = pass_result
                FinalReleaseComObject(xlrange)

                If dut2_en Then
                    jitter_pic_path = Jitter_folder2 & "\" & Jitter_pic_num2 & "_" & "Ta=" & TA_now & "; Fs=" & fs_now & "Hz; Vout=" & vout_now & "V; Vin=" & vin_now & "V; Iout=" & iout_now & "A" & ".PNG"
                Else
                    jitter_pic_path = Jitter_folder & "\" & Jitter_pic_num & "_" & "Ta=" & TA_now & "; Fs=" & fs_now & "Hz; Vout=" & vout_now & "V; Vin=" & vin_now & "V; Iout=" & iout_now & "A" & ".PNG"
                End If




                ' update_pic(jitter_pic_col, jitter_pic_row, jitter_pic_path)
                Hardcopy("PNG", jitter_pic_path)
                'update_pic(jitter_pic_col, jitter_pic_row)
                If (RS_Scope = False) And (check_fastAcq.Checked = True) Then
                    Scope_RUN(False)
                    FastAcq_ONOFF("ON")
                    'Delay(100)
                    Scope_RUN(True)
                    Delay_s(num_FastAcq.Value)

                    If dut2_en Then
                        jitter_pic_path = Jitter_folder2 & "\" & Jitter_pic_num2 & "_Fast_" & "Ta=" & TA_now & "; Fs=" & fs_now & "Hz; Vout=" & vout_now & "V; Vin=" & vin_now & "V; Iout=" & iout_now & "A" & ".PNG"
                    Else
                        jitter_pic_path = Jitter_folder & "\" & Jitter_pic_num & "_Fast_" & "Ta=" & TA_now & "; Fs=" & fs_now & "Hz; Vout=" & vout_now & "V; Vin=" & vin_now & "V; Iout=" & iout_now & "A" & ".PNG"
                    End If
                    ' update_pic(jitter_pic_col, jitter_pic_row, jitter_pic_path)
                    Hardcopy("PNG", jitter_pic_path)
                End If

                If dut2_en Then
                    Jitter_pic_num2 = Jitter_pic_num2 + 1
                Else
                    Jitter_pic_num = Jitter_pic_num + 1
                End If



                ' End If
                If run = False Then
                    Exit Function
                End If
                Scope_RUN(False)
                If RS_Scope = True Then
                    RS_Display(RS_RES_MES, RS_DISP_PREV)
                    Scope_measure_clear()
                    RS_Scope_measure_status(1, True)
                    RS_Scope_measure_status(2, True)
                    RS_Scope_measure_status(3, True)
                    RS_Scope_measure_status(4, True)
                    RS_Scope_measure_status(5, True)
                    RS_Scope_measure_status(6, True)
                    RS_Local()
                    'RS_View()
                End If

                If check_persistence.Checked = True Then
                    Display_persistence(False)
                    'Delay(100)
                End If

                If (RS_Scope = False) And (check_fastAcq.Checked = True) Then
                    FastAcq_ONOFF("OFF")
                    'Delay(100)
                End If


                If rbtn_vin_trigger.Checked = True Then
                    Trigger_set(lx_ch, "R", vin_now / num_vin_trigger.Value)
                Else
                    Trigger_auto_level(lx_ch, "R")
                End If
                RUN_set("RUNSTop")
                ' Scope_RUN(True)
                FinalReleaseComObject(xlrange)
                FinalReleaseComObject(xlSheet)
            Case Line_Regulation
                For i = 0 To test_cnt
                    If i = 0 Then
                        xlSheet = xlBook.Sheets(txt_LineR_sheet.Text)
                    Else
                        xlSheet = xlBook.Sheets(txt_LineR_sheet.Text & add_dut2)
                    End If
                    xlSheet.Activate()
                    If rbtn_lineR_test2.Checked = True Then
                        col_num = data_lineR_iout.Rows.Count + 2
                        row_num = data_vin.Rows.Count + 3
                    Else
                        col_num = data_lineR_iout.Rows.Count + 2
                        If check_lineR_up.Checked = True Then
                            row_num = (2 * data_lineR_vin.Rows.Count - 1) + 3
                        Else
                            row_num = data_lineR_vin.Rows.Count + 3
                        End If
                    End If

                    start_col = test_col + chart_width + col_Space + (TA_Test_num * total_vcc.Length * total_fs.Length + VCC_test_num * total_fs.Length + fs_test_num) * (col_num + 1)
                    'init row
                    If row_num < (chart_height + 1) Then
                        first_row = test_row + Vout_test_num * ((chart_height + 1) + row_Space)
                    Else
                        first_row = test_row + Vout_test_num * (row_num + row_Space)
                    End If

                    col = start_col
                    row = first_row + 3 + LR_Vin_test_num

                    If (TA_Test_num = TA_num) And (VCC_test_num = total_vcc.Length - 1) And (fs_test_num = total_fs.Length - 1) Then
                        xlrange = xlSheet.Range(ConvertToLetter(start_col + data_lineR_iout.Rows.Count + 1 + 2) & row)
                        xlrange.Value = PASS
                    Else
                        xlrange = xlSheet.Range(ConvertToLetter(start_col + data_lineR_iout.Rows.Count + 1) & row)
                        xlrange.Value = PASS
                    End If
                    FinalReleaseComObject(xlrange)


                    col = start_col + (1 + lineR_iout_num)
                    xlrange = xlSheet.Range(ConvertToLetter(col) & row)
                    If i = 0 Then
                        xlrange.Value = vout_meas
                        If vout_meas < (vout_now * (vout_err / 100)) Then
                            xlrange.Interior.Color = 255
                        End If
                    Else
                        xlrange.Value = vout_meas2
                        If vout_meas2 < (vout_now * (vout_err / 100)) Then
                            xlrange.Interior.Color = 255
                        End If
                    End If
                    FinalReleaseComObject(xlrange)

                    pass_value_Max = vout_now * (1 + (num_pass_lineR.Value / 100))
                    pass_value_Min = vout_now * (1 - (num_pass_lineR.Value / 100))

                    Dim judge_vout = 0
                    If i = 0 Then
                        judge_vout = vout_meas
                    Else
                        judge_vout = vout_meas2
                    End If

                    If judge_vout < pass_value_Min Or judge_vout > pass_value_Max Then
                        If (TA_Test_num = TA_num) And (VCC_test_num = total_vcc.Length - 1) And (fs_test_num = total_fs.Length - 1) Then
                            xlrange = xlSheet.Range(ConvertToLetter(start_col + data_lineR_iout.Rows.Count + 1 + 2) & row)
                            xlrange.Value = FAIL
                            xlrange.Interior.Color = test_fail_color
                        Else
                            xlrange = xlSheet.Range(ConvertToLetter(start_col + data_lineR_iout.Rows.Count + 1) & row)
                            xlrange.Value = FAIL
                            xlrange.Interior.Color = test_fail_color
                        End If
                    End If
                    FinalReleaseComObject(xlrange)
                    FinalReleaseComObject(xlSheet)

                    If check_lineR_scope.Checked = True Then

                        If i = 0 Then
                            xlSheet = xlBook.Sheets(txt_data_sheet.Text)
                        Else
                            xlSheet = xlBook.Sheets(txt_data_sheet.Text & add_dut2)
                        End If
                        xlSheet.Activate()
                        'report_test_update(TA_Test_num, start_test_time, txt_test_now.Text)
                        '----------------------------------------------------------------------------------
                        col_num = lineR_col.Length * data_lineR_iout.Rows.Count + 1
                        If check_lineR_up.Checked = True Then
                            row_num = (2 * data_lineR_vin.Rows.Count - 1) + 3
                        Else
                            row_num = data_lineR_vin.Rows.Count + 3
                        End If
                        start_col = test_col + col_Space + (TA_Test_num * total_vcc.Length * total_fs.Length + VCC_test_num * total_fs.Length + fs_test_num) * (col_num + 1)
                        first_row = test_row + Vout_test_num * (row_num + row_Space)
                        col = start_col
                        row = first_row + 3 + LR_Vin_test_num
                        col = start_col + (1 + lineR_iout_num * lineR_col.Length)
                        xlrange = xlSheet.Range(ConvertToLetter(col) & row)
                        If i = 0 Then
                            xlrange.Value = vout_meas
                        Else
                            xlrange.Value = vout_meas2
                        End If

                        FinalReleaseComObject(xlrange)
                        col = col + 1
                        'freq
                        For ii = 0 To 3
                            xlrange = xlSheet.Range(ConvertToLetter(col) & row)
                            xlrange.Value = fs(ii) / (10 ^ 3) ' Format(fs(ii) / (10 ^ 3), "#0.000")
                            FinalReleaseComObject(xlrange)
                            col = col + 1
                        Next
                        'ton
                        For ii = 0 To 3
                            xlrange = xlSheet.Range(ConvertToLetter(col) & row)
                            xlrange.Value = ton(ii) * (10 ^ 9) ' Format(ton(ii) * (10 ^ 9), "#0.000")
                            FinalReleaseComObject(xlrange)
                            col = col + 1
                        Next
                        'toff
                        For ii = 0 To 3
                            xlrange = xlSheet.Range(ConvertToLetter(col) & row)
                            xlrange.Value = toff(ii) * (10 ^ 9) ' Format(toff(ii) * (10 ^ 9), "#0.000")
                            FinalReleaseComObject(xlrange)
                            col = col + 1
                        Next
                        FinalReleaseComObject(xlrange)
                        FinalReleaseComObject(xlSheet)
                    End If
                Next
            Case Load_Regulation

                For i = 0 To test_cnt

                    If i = 0 Then
                        xlSheet = xlBook.Sheets(txt_LoadR_sheet.Text)
                    Else
                        xlSheet = xlBook.Sheets(txt_LoadR_sheet.Text & add_dut2)
                    End If
                    xlSheet.Activate()
                    'report_test_update(TA_Test_num, start_test_time, txt_test_now.Text)
                    '----------------------------------------------------------------------------------
                    'initial
                    'Init col
                    col_num = data_vin.Rows.Count + 2
                    row_num = data_eff_iout.Rows.Count + 3
                    start_col = test_col + chart_width + col_Space + (TA_Test_num * total_vcc.Length * total_fs.Length + VCC_test_num * total_fs.Length + fs_test_num) * (col_num + 1)
                    'Init row

                    If row_num < (chart_height + 1) Then
                        first_row = test_row + Vout_test_num * ((chart_height + 1) + row_Space)
                    Else
                        first_row = test_row + Vout_test_num * (row_num + row_Space)
                    End If
                    '----------------------------------------------------------------------------------
                    'Update Iout
                    col = start_col
                    row = first_row + 3 + eff_iout_num

                    'xlrange = xlSheet.Range(ConvertToLetter(col) & row)
                    'xlrange.Value = iout_now
                    'FinalReleaseComObject(xlrange)
                    If (TA_Test_num = TA_num) And (VCC_test_num = total_vcc.Length - 1) And (fs_test_num = total_fs.Length - 1) Then
                        xlrange = xlSheet.Range(ConvertToLetter(start_col + data_vin.Rows.Count + 1 + 2) & row)
                        xlrange.Value = PASS
                    Else
                        xlrange = xlSheet.Range(ConvertToLetter(start_col + data_vin.Rows.Count + 1) & row)
                        xlrange.Value = PASS
                    End If
                    FinalReleaseComObject(xlrange)

                    'Update Vout
                    pass_value_Max = vout_now * (1 + (num_pass_loadR.Value / 100))
                    pass_value_Min = vout_now * (1 - (num_pass_loadR.Value / 100))

                    col = start_col + (Vin_test_num) + 1
                    xlrange = xlSheet.Range(ConvertToLetter(col) & row)
                    If i = 0 Then
                        xlrange.Value = vout_meas
                    Else
                        xlrange.Value = vout_meas2
                    End If
                    FinalReleaseComObject(xlrange)

                    If i = 0 Then
                        If vout_meas < pass_value_Min Or vout_meas > pass_value_Max Then
                            If (TA_Test_num = TA_num) And (VCC_test_num = total_vcc.Length - 1) And (fs_test_num = total_fs.Length - 1) Then
                                xlrange = xlSheet.Range(ConvertToLetter(start_col + data_vin.Rows.Count + 1 + 2) & row)
                                xlrange.Value = FAIL
                                xlrange.Interior.Color = test_fail_color
                            Else
                                xlrange = xlSheet.Range(ConvertToLetter(start_col + data_vin.Rows.Count + 1) & row)
                                xlrange.Value = FAIL
                                xlrange.Interior.Color = test_fail_color
                            End If
                        End If
                    Else
                        If vout_meas2 < pass_value_Min Or vout_meas2 > pass_value_Max Then
                            If (TA_Test_num = TA_num) And (VCC_test_num = total_vcc.Length - 1) And (fs_test_num = total_fs.Length - 1) Then
                                xlrange = xlSheet.Range(ConvertToLetter(start_col + data_vin.Rows.Count + 1 + 2) & row)
                                xlrange.Value = FAIL
                                xlrange.Interior.Color = test_fail_color
                            Else
                                xlrange = xlSheet.Range(ConvertToLetter(start_col + data_vin.Rows.Count + 1) & row)
                                xlrange.Value = FAIL
                                xlrange.Interior.Color = test_fail_color
                            End If
                        End If

                    End If
                    '----------------------------------------------------------------------------------
                    FinalReleaseComObject(xlrange)
                    FinalReleaseComObject(xlSheet)
                Next

            Case Efficiency


                For i = 0 To test_cnt
                    If i = 0 Then
                        xlSheet = xlBook.Sheets(txt_eff_sheet.Text)
                    Else
                        xlSheet = xlBook.Sheets(txt_eff_sheet.Text & add_dut2)
                    End If
                    xlSheet.Activate()

                    'report_test_update(TA_Test_num, start_test_time, txt_test_now.Text)
                    '----------------------------------------------------------------------------------
                    'initial
                    'Init col
                    col_num = eff_title_total
                    row_num = data_eff_iout.Rows.Count + 2
                    start_col = test_col + chart_width + col_Space + (VCC_test_num * total_fs.Length + fs_test_num) * ((col_num + 1) * (data_vin.Rows.Count)) + Vin_test_num * (col_num + 1)
                    'Init row

                    If row_num < (chart_height + 1) Then

                        first_row = test_row + (Vout_test_num * (TA_num + 1) + TA_Test_num) * ((chart_height + 1) + row_Space)
                    Else
                        first_row = test_row + (Vout_test_num * (TA_num + 1) + TA_Test_num) * (row_num + row_Space)

                    End If


                    '----------------------------------------------------------------------------------
                    pass_value_Min = num_pass_eff.Value
                    'Update Vin
                    row = first_row + 2 + eff_iout_num
                    col = start_col

                    xlrange = xlSheet.Range(ConvertToLetter(col) & row)
                    If i = 0 Then
                        xlrange.Value = vin_meas
                    Else
                        xlrange.Value = vin_meas2
                    End If

                    FinalReleaseComObject(xlrange)
                    col = col + 1

                    xlrange = xlSheet.Range(ConvertToLetter(col) & row)
                    If i = 0 Then
                        xlrange.Value = iin_meas
                    Else
                        xlrange.Value = iin_meas2
                    End If

                    FinalReleaseComObject(xlrange)
                    col = col + 1

                    xlrange = xlSheet.Range(ConvertToLetter(col) & row)

                    If i = 0 Then
                        xlrange.Value = Eff_vout_meas
                        If Eff_vout_meas < (vout_now * (vout_err / 100)) Then
                            xlrange.Interior.Color = 255
                        End If
                    Else
                        xlrange.Value = Eff_vout_meas2
                        If Eff_vout_meas2 < (vout_now * (vout_err / 100)) Then
                            xlrange.Interior.Color = 255
                        End If
                    End If


                    FinalReleaseComObject(xlrange)
                    col = col + 1

                    xlrange = xlSheet.Range(ConvertToLetter(col) & row)
                    If i = 0 Then
                        xlrange.Value = iout_meas
                    Else
                        xlrange.Value = iout_meas2
                    End If

                    FinalReleaseComObject(xlrange)
                    col = col + 1

                    'eff


                    'Efficiency = (VOUT × ILOAD) / (VIN × IIN)
                    If i = 0 Then
                        eff = ((Eff_vout_meas * iout_meas) / (vin_meas * iin_meas)) * 100
                    Else
                        eff = ((Eff_vout_meas2 * iout_meas2) / (vin_meas2 * iin_meas2)) * 100
                    End If
                    xlrange = xlSheet.Range(ConvertToLetter(col) & row)
                    xlrange.Value = eff
                    FinalReleaseComObject(xlrange)
                    col = col + 1
                    If eff > pass_value_Min Then
                        pass_result = PASS
                    Else
                        pass_result = FAIL
                    End If

                    'loss
                    'Loss=(VIN*IIN)-(VOUT*IOUT)
                    xlrange = xlSheet.Range(ConvertToLetter(col) & row)
                    xlrange.Value = ((vin_meas * iin_meas) - (Eff_vout_meas * iout_meas))
                    FinalReleaseComObject(xlrange)
                    col = col + 1


                    If cbox_VCC.SelectedItem <> no_device Then
                        xlrange = xlSheet.Range(ConvertToLetter(col) & row)
                        xlrange.Value = vcc_meas
                        FinalReleaseComObject(xlrange)
                        col = col + 1
                    End If


                    If txt_Icc_addr.Text <> "" Then
                        xlrange = xlSheet.Range(ConvertToLetter(col) & row)
                        xlrange.Value = icc_meas
                        FinalReleaseComObject(xlrange)
                        col = col + 1

                    End If

                    If (cbox_VCC.SelectedItem <> no_device) And (txt_Icc_addr.Text <> "") Then
                        'Total Effieiency = (VOUT × ILOAD) / (VIN × IIN)+(VCC*ICC)

                        xlrange = xlSheet.Range(ConvertToLetter(col) & row)
                        xlrange.Value = ((Eff_vout_meas * iout_meas) / ((vin_meas * iin_meas) + (vcc_meas * icc_meas))) * 100
                        FinalReleaseComObject(xlrange)
                        col = col + 1
                    End If


                    'PASS
                    xlrange = xlSheet.Range(ConvertToLetter(col) & row)

                    If pass_result = FAIL Then
                        xlrange.Interior.Color = test_fail_color
                    End If
                    xlrange.Value = pass_result

                    '----------------------------------------------------------------------------------
                    FinalReleaseComObject(xlrange)
                    FinalReleaseComObject(xlSheet)
                Next ' test cnt loop
        End Select
        xlBook.Save()
    End Function

    Function Auto_Scanning_check() As Integer

        If RS_Scope = False Then
            Waveform_data_init(RL_value * 2)
        End If

        If rbtn_auto_all.Checked = True Then
            AutoScalling_EN = True
        Else
            AutoScalling_EN = False
        End If

        'If check_Force_CCM.Checked = True Then

        '    Exit Function
        'End If

        'Ton_min應該以當時有給定Ton_min為主()
        'toff_min 應該以當時有給定Toff_min為主
        '如果沒有給定的話, 目前先以
        'Ton_min = VOUT / VIN *(1/fs)* 0.7 判定。
        'Toff_min = (1-VOUT/VIN) *(1/fs) *0.7。

        If rbtn_ton_cal.Checked = True Then
            ton_pass = (vout_now / vin_now) * (1 / fs_now) * num_ton_cal.Value
        Else
            ton_pass = num_ton_val.Value * ns
        End If

        If rbtn_toff_cal.Checked = True Then
            toff_pass = (1 - vout_now / vin_now) * (1 / fs_now) * num_toff_cal.Value
        Else
            toff_pass = num_toff_val.Value * ns
        End If

        AutoScalling_EN = True

    End Function



    Function Auto_Scanning(ByVal lx_sel As Integer) As Double()

        Dim i As Integer
        Dim ton_test As Boolean
        Dim ton_start_time() As Double
        Dim ton_stop_time As Double

        Dim toff_test As Boolean
        Dim toff_start_time As Double
        Dim toff_stop_time As Double

        Dim meas_ton_value() As Double
        Dim meas_toff_value() As Double
        Dim meas_freq_value() As Double

        Dim meas_cursor_value() As Double


        Dim meas_num As Integer = 0
        Dim meas_volt_value As Double
        Dim meas_time_value As Double
        Dim max_data As Integer
        Dim meas_ton_update As Double
        Dim meas_toff_update As Double
        Dim meas_freq_update As Double
        Dim wave_data() As Double = {0, 0, 0} 'Ton(ns), Toff(ns), Freq(KHz)


        Dim scale_value As Double
        Dim RL As Integer

        Dim Point As Integer = 7

        Dim error_count As Integer = 0

        Dim error_num As Integer = 1
        Dim ByteSize As Long
        ' Dim time_volt As Object


        Dim start_meas As Boolean = False

        Dim meas_cursor_start, meas_cursor_stop As Double

        Dim temp() As String

        Dim time_volt As Object

        Dim wave_time() As Double
        Dim wave_volt() As Double

        Dim list() As String

        If run = False Then
            Return wave_data
            Exit Function
        End If

        If RS_Scope = True Then
            Scope_RUN(True)
        End If
        note_string = "Start..."
        Information.information_run("Auto Scanning", note_run)

        'Cursors

        If check_cursors.Checked = True Then
            Cursor_ONOFF("ON")
        End If

        'check RL

        note_string = "Check Scope ..."
        scale_value = H_scale_now()





        'Point / min(Ton_min, Toff_min) = RL / (10*A) 

        'A(us/Div) 是個代號，看當時的波形的時間scale而定
        'RL = Point / min(Ton_min, toff_min) * 10 * A

        'Point > N, N = 10



        'If ton_pass > toff_pass Then
        '    RL = Point / (ton_pass) * 10 * scale_value
        'Else
        '    RL = Point / (toff_pass) * 10 * scale_value
        'End If

        If ton_pass > toff_pass Then
            RL = Point / (ton_pass) * 5 * scale_value
        Else
            RL = Point / (toff_pass) * 5 * scale_value
        End If


        '計算出來:
        'If RL > 500K，則忽略掉。
        'If 500K > RL > 100K，則以當時計算的RL取相近(略大)的值為主。
        'If 100K > RL > 20K，則以100K 為主。
        'If 20K > RL，則以20K為主。 



        'If RL > 500000 Then
        '    Return wave_data

        '    Exit Function
        'ElseIf (RL > 250000) Then

        '    H_reclength(500000)

        'ElseIf (RL > 100000) Then

        '    H_reclength(250000)

        'ElseIf (RL > 50000) Then

        '    H_reclength(100000)

        'ElseIf (RL > 20000) Then

        '    H_reclength(50000)
        'Else
        '    H_reclength(50000)
        'End If
        H_reclength(50000)
        'System.Threading.Thread.Sleep(1000)

        CHx_Bandwidth(lx_sel, "20MHz")

        'Timing Scale
        H_scale(H_scale_value, "ns") '1/Fs_Min(Hz)*n/10 

        Scope_measure_reset()
        Dim cnt = 0
        While (cnt <= 0)
            cnt = Scope_measure_count(1)
        End While


        note_string = "Capture Wave ..."
        ' error_capture(vout_ch, "R", vpp(4), False, vpp(3), num_delay_error.Value)


        ' trigger function
        'If (cbox_coupling_vout.SelectedItem <> "AC") And (vpp(5) < (vout_now * (1 - num_vout_neg.Value / 100))) Then
        '    error_capture(vout_ch, "R", vpp(5), True, vpp(2), num_delay_error.Value)
        'Else
        '    error_capture(vout_ch, "R", vpp(4), False, vpp(3), num_delay_error.Value)
        'End If

        ByteSize = Waveform_data(Main.txt_scope_folder.Text & "\wave.csv", wave_pc_path, lx_sel)

        If run = False Then
            Return wave_data
            Exit Function
        End If

        If ByteSize = 0 Then
            check_file_open(wave_pc_path)
            ByteSize = Waveform_data(Main.txt_scope_folder.Text & "\wave.csv", wave_pc_path, lx_sel)
            If run = False Then
                Return wave_data
                Exit Function
            End If
        End If

        If ByteSize > 0 Then
            If RS_Scope = False Then

                'xlApp.DisplayAlerts = False
                'xlApp.Visible = False
                xlBook_wave = xlApp.Workbooks.Open(wave_pc_path)
                '-----------------------------------------------------------------
                xlSheet_wave = xlBook_wave.Sheets("wave")
                xlSheet_wave.Activate()

                max_data = xlSheet_wave.Range("B1").Value
                time_volt = xlSheet_wave.Range(xlApp.Cells(1, 4), xlApp.Cells(max_data, 5)).Value()
                xlBook_wave.Close(True) '關閉工作簿
                ' xlApp.Quit() '結束EXCEL對象
                xlSheet_wave = Nothing
                xlBook_wave = Nothing
                GC.Collect()
            Else
                Dim f As New IO.FileInfo(wave_pc_path)
                Dim sr As IO.StreamReader = f.OpenText '產生StreamReader的sr物件
                note_string = "Get Wave data..."
                list = Split(sr.ReadToEnd, vbNewLine)

                max_data = list.Length

                sr.Close()   '???桀????
            End If
            ReDim ton_start_time(max_data)
            'time_volt(n, 1) =Time  (n=1~max_data)
            'time_volt(n, 2) =Volt  (n=1~max_data)
            note_string = "Analysis Wave data..."
            If max_data = 0 Then
                error_message("Wave Format Error!!!")
            Else
                ReDim wave_time(max_data - 1)
                ReDim wave_volt(max_data - 1)
                For i = 1 To max_data
                    System.Windows.Forms.Application.DoEvents()
                    If run = False Then
                        Exit For
                    End If
                    If RS_Scope = False Then
                        wave_time(i - 1) = time_volt(i, 1)
                        wave_volt(i - 1) = time_volt(i, 2)
                    Else
                        If list(i - 1) <> "" Then
                            temp = Split(list(i - 1), ",")
                            wave_time(i - 1) = Val(temp(0))
                            wave_volt(i - 1) = Val(temp(1))
                        End If

                    End If


                    meas_volt_value = wave_volt(i - 1)


                    If start_meas = False Then
                        '以量測到的電壓低於vin*0.9開始分析，避免誤取到不完整的Ton

                        If meas_volt_value < (vin_now * num_ton_vin.Value) Then
                            start_meas = True
                        End If

                    Else

                        Select Case True

                            Case (ton_test = False) And (toff_test = False)
                                'Ton Start
                                'Step1: 
                                '找尋Lx的電壓大於Vin*0.9就當作Ton的起始點

                                If (meas_volt_value > (vin_now * num_ton_vin.Value)) Then
                                    'ton_start_time(meas_num) = time_volt(i, 1)
                                    ton_start_time(meas_num) = wave_time(i - 1)
                                    ton_test = True

                                End If



                            Case (ton_test = True) And (toff_test = False)
                                'Check Ton
                                'Step2: 
                                '找尋Lx的電壓小於Vin*0.9就當作Ton的終點
                                '算出Ton的時間，若Ton大於Ton_pass的時間，認為是有校的Ton
                                '若小於Ton pass就回到Step1重新找Ton的start
                                '若已經找尋到兩個Ton就算出freq

                                If (meas_volt_value < (vin_now * num_ton_vin.Value)) Then
                                    'ton_stop_time = time_volt(i, 1)
                                    ton_stop_time = wave_time(i - 1)

                                    meas_time_value = (ton_stop_time - ton_start_time(meas_num))
                                    If meas_time_value >= ton_pass Then
                                        If (meas_num > 0) Then

                                            ReDim Preserve meas_freq_value(meas_num - 1)


                                            meas_freq_value(meas_num - 1) = 1 / (ton_start_time(meas_num) - ton_start_time(meas_num - 1)) 'Hz


                                        End If

                                        ReDim Preserve meas_ton_value(meas_num)
                                        meas_ton_value(meas_num) = meas_time_value

                                        ReDim Preserve meas_cursor_value(meas_num)

                                        meas_cursor_value(meas_num) = ton_start_time(meas_num)

                                        toff_test = True
                                    Else
                                        ton_test = False
                                    End If

                                End If

                            Case (ton_test = True) And (toff_test = True)
                                'Step3:
                                '若Ton有大於Ton_pass，在偵測Toff低於Vin*0.2的過程中，有高於Vin*0.9就認為此Ton 無效，重新回到Step1
                                '找尋Lx的電壓小於Vin*0.2就當作Toff的起始點
                                If (meas_volt_value > (vin_now * num_ton_vin.Value)) Then
                                    'Phase ring 
                                    'Reset ton
                                    ton_test = False
                                    toff_test = False
                                Else

                                    'Toff Start
                                    If (meas_volt_value < (vin_now * num_toff_vin.Value)) Then
                                        'toff_start_time = time_volt(i, 1)

                                        toff_start_time = wave_time(i - 1)


                                        ton_test = False
                                    End If

                                End If


                            Case (ton_test = False) And (toff_test = True)
                                'Step4:
                                '找尋Lx的電壓大於Vin*0.2就當作Toff的終點
                                '算出Toff的時間，若Toff大於Toff_pass的時間，認為是有校的Toff
                                '若小於Toff pass就回到Step3重新找Toff的start

                                If (meas_volt_value > (vin_now * num_toff_vin.Value)) Then
                                    'toff_stop_time = time_volt(i, 1)

                                    toff_stop_time = wave_time(i - 1)


                                    meas_time_value = (toff_stop_time - toff_start_time)

                                    If meas_time_value >= toff_pass Then

                                        ReDim Preserve meas_toff_value(meas_num)
                                        meas_toff_value(meas_num) = meas_time_value

                                        meas_num = meas_num + 1
                                        toff_test = False
                                    Else
                                        ton_test = True
                                    End If




                                End If




                        End Select

                    End If










                Next

                If meas_num > 1 Then

                    note_string = "Success!"


                    meas_freq_update = meas_freq_value(0)
                    meas_ton_update = meas_ton_value(0)
                    meas_toff_update = meas_toff_value(0)
                    meas_cursor_start = meas_cursor_value(0)
                    meas_cursor_stop = meas_cursor_value(1)

                    For i = 0 To meas_num - 2


                        If rbtn_freq_low.Checked = True Then
                            'freq 取小
                            If meas_freq_update > meas_freq_value(i) Then
                                meas_freq_update = meas_freq_value(i)
                                meas_ton_update = meas_ton_value(i)
                                meas_toff_update = meas_toff_value(i)
                                meas_cursor_start = meas_cursor_value(i)
                                meas_cursor_stop = meas_cursor_value(i + 1)
                            End If
                        Else
                            'ton 取小
                            'toff取小
                            'freq取大
                            If meas_freq_update < meas_freq_value(i) Then
                                meas_freq_update = meas_freq_value(i)
                                meas_ton_update = meas_ton_value(i)
                                meas_toff_update = meas_toff_value(i)
                                meas_cursor_start = meas_cursor_value(i)
                                meas_cursor_stop = meas_cursor_value(i + 1)
                            End If
                        End If
                    Next
                    wave_data(0) = meas_ton_update
                    wave_data(1) = meas_toff_update
                    wave_data(2) = meas_freq_update

                    Dim cursor_delta_value As Double

                    If (check_cursors.Checked = True) And (wave_data(2) <> 0) Then
                        Cursor_move("VBArs", meas_cursor_start, meas_cursor_stop)
                        cursor_delta_value = Cursor_delta("VBArs")
                    End If


                Else
                    note_string = "Fail!"

                    If check_cursors.Checked = True Then

                        Cursor_ONOFF("OFF")


                    End If

                End If
            End If




        End If



        note_display = False




        Return wave_data

    End Function

    Function error_hyperlink() As Integer
        Dim error_path As String
        'Update Picture

        hyperlink_col = error_pic_col
        hyperlink_row = error_pic_row

        If (error_pic_num Mod 10 = 0) Then
            error_pic_col = 1
            error_pic_row = error_pic_row + pic_height + 2
        Else
            error_pic_col = error_pic_col + pic_width + 1
        End If

        error_pic_num = error_pic_num + 1


    End Function

    Function TestITem_run() As Integer
        Dim i, n, ii, v, x, y As Integer
        Dim t As Integer
        Dim total_iout_num As Integer
        Dim stability_iout() As Double
        Dim TA_temp As Integer
        Dim VCC_num As Integer
        Dim VCC_temp As String
        Dim set_num As Integer = 0
        Dim test_point_temp As Integer
        Dim first_Check As Boolean = True
        Dim vout_temp As Double
        ReDim eff_iin_change(data_eff.Rows.Count - 1)
        Dim num As Integer
        Vout_TA_set = txt_OTP.Text
        Power_recorve = check_OTP.Checked
        TestITem_run_now = True
        ''Init Parameter
        instrument_init()
        If run = False Then
            Exit Function
        End If
        vout_err = num_Vout_error.Value
        Meter_iin_low = Mid(cbox_meter_mini.SelectedItem, 1, 4)
        iin_meter_change = num_iin_change.Value / 1000
        iout_meter_change = num_iout_change.Value / 1000

        iin_meter_change2 = num_iin_change2.Value / 1000
        iout_meter_change2 = num_iout_change2.Value / 1000

        If (check_stability.Checked = True) Or (check_jitter.Checked = True) Then
            '--------------------------------------------------------------------------------
            '2020/01/31
            '初始以最小為10mV來設定, 讀取measure再來微調
            'Initial Vout Scale
            If rbtn_auto_vout.Checked = True Then
                vout_temp = Math.Floor(((vout_now * 1000) * 0.005))
                vout_temp = Math.Floor(vout_temp / 5) * 5
                If vout_temp > 10 Then
                    vout_scale_now = vout_temp
                Else
                    vout_scale_now = 10
                End If
            Else
                If (check_Force_CCM.Checked = True) Then
                    vout_scale_now = num_vout_CCM.Value
                Else
                    vout_scale_now = num_vout_DEM.Value
                End If
            End If
            CHx_scale(vout_ch, vout_scale_now, "mV") 'Voltage Scale > VID * 10% / 4
            If DUT2_en Then : CHx_scale(vout2_ch, vout_scale_now, "mV") : End If
            VoutScalling_CCM = True
            first_Check = True
        End If
        report_init()
        GC.Collect()
        GC.WaitForPendingFinalizers()
        Delay(100)
        If check_jitter.Checked = True Then
            If TA_Test_num = 0 Then

                Jitter_pic_num = 1
                Jitter_pic_num2 = 1
            Else
                Jitter_pic_num = data_jitter_iout.Rows.Count * total_vcc.Length * total_fs.Length * total_vout.Length * data_vin.Rows.Count * TA_Test_num + 1
                Jitter_pic_num2 = data_jitter_iout.Rows.Count * total_vcc.Length * total_fs.Length * total_vout.Length * data_vin.Rows.Count * TA_Test_num + 1
            End If
        End If
        'TA
        Delay(100)
        If check_stability.Checked = True Then
            error_pic_col = 1
            If TA_Test_num = 0 Then
                error_pic_row = 3
                error_pic_num = 1

                error_pic_row2 = 3
                error_pic_num2 = 1
                beta_pic_num = 1
            Else
                excel_open()
                xlSheet = xlBook.Sheets(txt_error_sheet.Text)
                xlSheet.Activate()
                error_pic_num = xlSheet.Range(ConvertToLetter(1) & 1).Value
                error_pic_row = (pic_height + 2) * Int(error_pic_num / 10) + 3
                error_pic_col = (pic_width + 1) * (error_pic_num Mod 10) + 1
                error_pic_num = error_pic_num + 1
                beta_pic_num = data_set.Rows.Count * data_test.Rows.Count * TA_Test_num + 1


                If DUT2_en Then
                    xlSheet = xlBook.Sheets(txt_error_sheet.Text & add_dut2)
                    xlSheet.Activate()
                    error_pic_num2 = xlSheet.Range(ConvertToLetter(1) & 1).Value
                    error_pic_row2 = (pic_height + 2) * Int(error_pic_num2 / 10) + 3
                    error_pic_col2 = (pic_width + 1) * (error_pic_num2 Mod 10) + 1
                    error_pic_num2 = error_pic_num2 + 1
                    beta_pic_num = data_set.Rows.Count * data_test.Rows.Count * TA_Test_num + 1
                End If
                excel_close()
            End If
        End If

        PartI_file = sf_name
        If TA_Test_num = 0 Then
            test_point_num = 0
            txt_test_now.Text = 0
        End If
        test_point_temp = test_point_num
        'Start Test
        start_test_time = Now
        For n = 0 To total_vcc.Length - 1
            System.Windows.Forms.Application.DoEvents()
            While pause = True
                System.Windows.Forms.Application.DoEvents()
                If run = False Then
                    Exit While
                End If
            End While
            If run = False Then
                Exit For
            End If
            VCC_test_num = n
            If data_VCC.Rows.Count > 0 Then
                vcc_now = total_vcc(n)
                DCLoad_ONOFF("OFF")
                If DUT2_en Then : DCLoad_ONOFF("OFF", DUT2_en) : End If
                Power_Dev = VCC_Dev
                power_volt(VCC_device, VCC_out, vcc_now)
            End If
            For i = 0 To total_fs.Length - 1
                System.Windows.Forms.Application.DoEvents()
                While pause = True
                    System.Windows.Forms.Application.DoEvents()
                    If run = False Then
                        Exit While
                    End If
                End While
                If run = False Then
                    Exit For
                End If
                fs_now = total_fs(i)
                fs_test_num = i
                If Main.cbox_fs_ctr.SelectedItem <> no_device Then
                    DCLoad_ONOFF("OFF")
                    Grobal_Control(Fs_control, fs_now)
                    If DUT2_en Then
                        DCLoad_ONOFF("OFF", DUT2_en)
                        Grobal_Control(Fs_control, fs_now, DUT2_en)
                    End If
                    If Main.check_EN_off.Checked = True Then
                        Power_EN(False)
                        Power_EN(True)
                    End If
                End If

                For ii = 0 To total_vout.Length - 1
                    System.Windows.Forms.Application.DoEvents()
                    While pause = True
                        System.Windows.Forms.Application.DoEvents()
                        If run = False Then
                            Exit While
                        End If
                    End While
                    If run = False Then
                        Exit For
                    End If
                    vout_now = total_vout(ii)
                    Vout_test_num = ii
                    DCLoad_check_range()
                    If Main.cbox_vout_ctr.SelectedItem <> no_device Then
                        DCLoad_ONOFF("OFF")
                        Grobal_Control(Vout_control, vout_now)
                        If DUT2_en Then
                            DCLoad_ONOFF("OFF", DUT2_en)
                            Grobal_Control(Vout_control, vout_now, DUT2_en)
                        End If

                        If Main.check_EN_off.Checked = True Then
                            Power_EN(False)
                            Power_EN(True)
                        End If
                        '確認是否要Auto_Scanning
                        first_Check = True
                    End If

                    If check_stability.Checked = True Or check_jitter.Checked = True Or check_Efficiency.Checked = True Or check_loadR.Checked = True Or ((check_LineR.Checked = True) And (rbtn_lineR_test2.Checked = True)) Then
                        For v = 0 To data_vin.Rows.Count - 1
                            System.Windows.Forms.Application.DoEvents()

                            While pause = True
                                System.Windows.Forms.Application.DoEvents()
                                If run = False Then
                                    Exit While
                                End If
                            End While
                            If run = False Then
                                Exit For
                            End If
                            excel_open()
                            vin_now = data_vin.Rows(v).Cells(0).Value
                            Vin_test_num = v
                            '-----------------------------------------------------------------------------------------------------------
                            DCLoad_ONOFF("OFF")
                            If DUT2_en Then
                                DCLoad_ONOFF("OFF", DUT2_en)
                            End If

                            Power_Dev = vin_Dev
                            power_volt(vin_device, Vin_out, vin_now)

                            If DUT2_en Then
                                Power_Dev = vin_Dev2
                                power_volt(vin_device2, Vin_out2, vin_now)
                            End If

                            '-----------------------------------------------------------------------------------------------------------
                            If check_stability.Checked = True Or check_jitter.Checked = True Then
                                'CHx_scale(lx_ch, (vin_now / num_lx_scale.Value), "V") 'Voltage Scale > SW/2
                                If rbtn_manual_lx.Checked = True Then
                                    CHx_scale(lx_ch, num_scale_lx.Value, "mV") 'Voltage Scale > SW/2
                                Else
                                    CHx_scale(lx_ch, (vin_now / num_lx_scale.Value), "V") 'Voltage Scale > SW/2
                                End If

                                If rbtn_vin_trigger.Checked = True Then
                                    Trigger_set(lx_ch, "R", vin_now / num_vin_trigger.Value)
                                Else
                                    Trigger_auto_level(lx_ch, "R")
                                End If

                                If DUT2_en Then
                                    If rbtn_manual_lx.Checked = True Then
                                        CHx_scale(lx2_ch, num_scale_lx.Value, "mV") 'Voltage Scale > SW/2
                                    Else
                                        CHx_scale(lx2_ch, (vin_now / num_lx_scale.Value), "V") 'Voltage Scale > SW/2
                                    End If

                                    'If rbtn_vin_trigger.Checked = True Then
                                    '    Trigger_set(lx2_ch, "R", vin_now / num_vin_trigger.Value)
                                    'Else
                                    '    Trigger_auto_level(lx2_ch, "R")
                                    'End If
                                End If
                            End If



                            '-----------------------------------------------------------------------------------------------------------
                            'Check Iout
                            If total_other_iout > 0 Then
                                ReDim total_iout(total_other_iout - 1)
                                total_iout = other_iout
                                total_iout_num = total_other_iout
                            Else
                                total_iout_num = 0
                            End If

                            '-----------------------------------------------------------------------------------------------------------

                            If check_stability.Checked = True Then
                                set_num = TA_Test_num * total_vcc.Length * total_fs.Length * total_vout.Length * data_vin.Rows.Count + n * total_vcc.Length * total_fs.Length * total_vout.Length * data_vin.Rows.Count + i * total_vout.Length * data_vin.Rows.Count + ii * data_vin.Rows.Count + v
                                num = 0
                                data_set_now = set_num
                                For x = stability_row_start(set_num) To stability_row_stop(set_num)
                                    ReDim Preserve stability_iout(num)
                                    stability_iout(num) = data_result.Rows(x).Cells("col_test_stability").Value
                                    ReDim Preserve total_iout(total_iout_num)
                                    total_iout(total_iout_num) = stability_iout(num)
                                    total_iout_num = total_iout_num + 1
                                    num = num + 1
                                Next

                                If check_Force_CCM.Checked = False Then
                                    Fs_leak_0A = test_fs0(set_num)
                                    ton_now = test_ton(set_num) / (10 ^ 9)
                                    IOUT_Boundary_Start = test_IOB_start(set_num)
                                    IOUT_Boundary_Stop = test_IOB_stop(set_num)
                                End If
                            End If

                            '' 過濾重複的陣列元素

                            If total_iout_num = 0 Then
                                Exit For
                            End If

                            Array.Sort(total_iout)
                            total_iout = total_iout.Distinct.ToArray()
                            '-----------------------------------------------------------------------------------------------------------
                            'check Iin Meter
                            If (check_Efficiency.Checked = True) And ((check_iin.Checked = True) Or (rbtn_board_iin.Checked = True)) Then
                                If (n = 0) And (i = 0) Then
                                    check_meter_iin_max()
                                Else
                                    iin_meter_change = eff_iin_change(ii * data_vin.Rows.Count + v)
                                End If
                            ElseIf rbtn_iin_current_measure.Checked Or
                                   rbtn_iout_current_measure.Checked Or
                                   rbtn_iin_current_measure2.Checked Or
                                   rbtn_iout_current_measure2.Checked Then

                                If rbtn_iin_current_measure.Checked Then
                                    relay_in_meter_intial()
                                End If

                                If rbtn_iout_current_measure.Checked Then
                                    realy_out_meter_initial()
                                End If

                                ' Ollie_note: current meter board add initial code
                                If DUT2_en Then
                                    If rbtn_iin_current_measure2.Checked Then
                                        relay_in_meter_intial(DUT2_en)
                                    End If

                                    If rbtn_iout_current_measure2.Checked Then
                                        realy_out_meter_initial(DUT2_en)
                                    End If
                                End If


                            End If
                            '-----------------------------------------------------------------------------------------------------------
                            'Iout Setting
                            eff_iout_num = 0
                            lineR_iout_num = 0
                            stability_iout_num = 0
                            jitter_iout_num = 0

                            '-----------------------------------------------------------------------------------------------------------
                            'Start RUN

                            For x = 0 To total_iout.Length - 1
                                System.Windows.Forms.Application.DoEvents()
                                While pause = True
                                    System.Windows.Forms.Application.DoEvents()
                                    If run = False Then
                                        Exit While
                                    End If
                                End While

                                If run = False Then
                                    Exit For
                                End If

                                test_point_num = test_point_num + 1
                                txt_test_now.Text = test_point_num
                                iout_now = total_iout(x)
                                '-------------------------------------------------------------------------------------
                                If (check_Efficiency.Checked = True) Then
                                    '如果量測Eff需要調整Meter，其他都保持在Max range

                                    If (Iin_change = True) Then
                                        If rbtn_meter_iin.Checked = True Then
                                            Iin_meter_set(check_iin, cbox_IIN_meter, cbox_IIN_relay)
                                        Else
                                            INA226_IIN_set()
                                        End If

                                        If rbtn_meter_iin2.Checked And DUT2_en Then
                                            Iin_meter_set(check_iin2, cbox_IIN_meter2, cbox_IIN_relay2, DUT2_en)
                                        End If

                                    End If


                                    If rbtn_meter_iout.Checked = True Then
                                        Iout_meter_set(check_iout, cbox_Iout_meter, cbox_Iout_relay)
                                    End If

                                    If rbtn_meter_iout2.Checked And DUT2_en Then
                                        Iout_meter_set(check_iout2, cbox_Iout_meter2, cbox_Iout_relay2, DUT2_en)
                                    End If
                                End If
                                '-------------------------------------------------------------------------------------

                                'Iout
                                'DCLoad_Iout(iout_now, monitor_vout, DUT2_en)
                                Iout_Setting(iout_now, monitor_vout, DUT2_en)
                                'DCLoad_Iout(iout_now, monitor_vout)
                                If (DCLoad_ON = False) Then
                                    DCLoad_ONOFF("ON")
                                    If DUT2_en Then
                                        DCLoad_ONOFF("ON", DUT2_en)
                                    End If
                                End If

                                'Ollie_note: test flow update here
                                '-------------------------------------------------------------------------------------
                                'Measure
                                If (check_Efficiency.Checked = True) And (rbtn_meter_iin.Checked = True) And (Iin_Meter_Max = False) Then
                                    '確認小檔位的單位 100uA, 1mA....
                                    iin_meas = meter_average(cbox_IIN_meter.SelectedItem, Meter_iin_dev, 1, Meter_iin_range, Meter_iin_low)
                                    Meter_iin_range = Meter_range_now
                                End If


                                ''----------------------------------------------------------------------------------
                                'Vin Sense
                                If check_vin_sense.Checked = True Then
                                    'Vin Sense
                                    vin_power_sense(cbox_vin.SelectedItem, num_vin_sense.Value, num_vin_max.Value, vin_now)
                                End If

                                If check_vin_sense2.Checked And DUT2_en Then
                                    vin_power_sense(cbox_vin2.SelectedItem, num_vin_sense2.Value, num_vin_max2.Value, vin_now, DUT2_en)
                                End If


                                ''----------------------------------------------------------------------------------
                                'Measure
                                If (iout_now > num_iout_delay.Value) And (num_delay.Value > 0) Then
                                    If cbox_delay_unit.SelectedIndex = 1 Then
                                        Delay_s(num_delay.Value)
                                    Else
                                        Delay(num_delay.Value)
                                    End If
                                End If


                                'vin
                                vin_meas = DAQ_average(vin_daq, num_data_count.Value)
                                If DUT2_en Then
                                    vin_meas2 = DAQ_average(vin_daq2, num_data_count2.Value)
                                End If
                                ''----------------------------------------------------------------------------------
                                'Check Vout
                                'vout
                                vout_meas = DAQ_average(vout_daq, num_data_count.Value)
                                If DUT2_en Then
                                    vout_meas2 = DAQ_average(vout_daq2, num_data_count2.Value)
                                End If

                                'iout
                                If (rbtn_meter_iout.Checked = True) And (cbox_Iout_meter.SelectedItem <> no_device) Then
                                    iout_meas = meter_average(cbox_Iout_meter.SelectedItem, Meter_iout_dev, num_data_count.Value, Meter_iout_range, Meter_iout_low)

                                    Meter_iout_range = Meter_range_now
                                ElseIf rbtn_iout_current_measure.Checked Then
                                    ' in_out_sel = 0: input current
                                    ' in_out_sel = 1: output current
                                    iout_meas = meter_auto(1, num_meter_count.Value)

                                ElseIf rbtn_board_iout.Checked = True Then
                                    'relay read
                                    iout_meas = INA226_IOUT_meas(cbox_board_buck.SelectedIndex + 1, num_data_count.Value)
                                Else
                                    iout_meas = load_read("CURR") ' Format(load_read("CURR"), "#0.000000000") '"DCLoad_Iout(A)"
                                End If

                                If (rbtn_meter_iout2.Checked = True) And (cbox_Iout_meter2.SelectedItem <> no_device) And DUT2_en Then
                                    iout_meas2 = meter_average(cbox_Iout_meter2.SelectedItem, Meter_iout_dev2, num_data_count2.Value, Meter_iout_range, Meter_iout_low)
                                    Meter_iout_range = Meter_range_now
                                ElseIf rbtn_board_iout2.Checked And DUT2_en Then
                                    iout_meas2 = meter_auto(1, num_meter_count.Value, DUT2_en)
                                ElseIf DUT2_en Then
                                    iout_meas2 = load_read("CURR", DUT2_en)
                                End If


                                '------------------------------------------------------
                                'Update Test Result
                                '------------------------------------------------------
                                'Update Efficiency & Load Regulation
                                If check_Efficiency.Checked = True Or check_loadR.Checked = True Then
                                    For y = 0 To data_eff_iout.Rows.Count - 1
                                        If iout_now = data_eff_iout.Rows(y).Cells(0).Value Then
                                            eff_iout_num = y
                                            Efficiency_run(DUT2_en) ' if DUT2_en = true, efficiency will run 2 group instrument
                                            Exit For
                                        End If
                                    Next
                                End If


                                If (check_LineR.Checked = True) And (rbtn_lineR_test2.Checked = True) Then
                                    LR_Vin_test_num = Vin_test_num
                                    For y = 0 To data_lineR_iout.Rows.Count - 1
                                        If iout_now = data_lineR_iout.Rows(y).Cells(0).Value Then
                                            lineR_iout_num = y
                                            LineR_run(DUT2_en) ' alread add dut2 flow
                                            update_report(Line_Regulation, DUT2_en)

                                            Exit For
                                        End If
                                    Next
                                End If


                                '------------------------------------------------------
                                'Update Stability and Jitter
                                If ((check_stability.Checked = True) Or (check_jitter.Checked = True)) And (first_Check = True) Then
                                    '第一次要啟動
                                    Auto_Scanning_check()
                                    If (check_scope_iout.Checked = True) Then
                                        If (iout_now <= 0.6) And (iout_scale_now <> 200) Then
                                            'IOUT
                                            iout_scale_now = 200
                                            CHx_scale(iout_ch, iout_scale_now, "mV") 'a. IOUT < 600mA, Scale = 200mA, b. 600mA<IOUT < 3A, Sacle = 1A,c. 3A <IOUT< 6A, Scale = 2A
                                        ElseIf (iout_now > 0.6) And (iout_now <= 3) And (iout_scale_now <> 1000) Then
                                            iout_scale_now = 1000
                                            CHx_scale(iout_ch, iout_scale_now, "mV") 'a. IOUT < 600mA, Scale = 200mA, b. 600mA<IOUT < 3A, Sacle = 1A,c. 3A <IOUT< 6A, Scale = 2A
                                        ElseIf (iout_now > 3) And (iout_now <= 6) And (iout_scale_now <> 2000) Then
                                            iout_scale_now = 2000
                                            CHx_scale(iout_ch, iout_scale_now, "mV") 'a. IOUT < 600mA, Scale = 200mA, b. 600mA<IOUT < 3A, Sacle = 1A,c. 3A <IOUT< 6A, Scale = 2A
                                        End If
                                    End If
                                    first_Check = False
                                End If


                                '-------------------------------------------------------------------------
                                '2020/01/22  Vout用DC量測時 Scope的offset要由DAQ校正
                                If ((check_stability.Checked = True) Or (check_jitter.Checked = True)) And (check_offset_vout.Checked = True) And (cbox_coupling_vout.SelectedItem <> "AC") Then
                                    CHx_OFFSET(vout_ch, vout_meas)
                                End If

                                ''----------------------------------------------------------------------------------
                                If check_stability.Checked = True Then
                                    For y = 0 To stability_iout.Length - 1
                                        If iout_now = stability_iout(y) Then
                                            stability_iout_num = y


                                            ' Runtime test
                                            Dim sw As Stopwatch = New Stopwatch()
                                            sw.Reset()
                                            sw.Start()
                                            Stability_run()
                                            sw.Stop()

                                            Dim res_ms As Long = sw.ElapsedMilliseconds
                                            Dim res_s As Long = res_ms / 1000
                                            Dim res_min As Long = Int(res_s / 60)
                                            res_s = res_s Mod 60
                                            Console.WriteLine("count 100 case Spend Time: {0}min_{1}s", res_min, res_s)
                                            Console.WriteLine("-------------------------------------------------------")
                                            '--------------------------------------------------------------------------------

                                            sw.Reset()
                                            sw.Start()
                                            Stability_run1(50)

                                            sw.Stop()
                                            res_ms = sw.ElapsedMilliseconds
                                            res_s = res_ms / 1000
                                            res_min = Int(res_s / 60)
                                            res_s = res_s Mod 60
                                            Console.WriteLine("count 50 case Spend Time: {0}min_{1}s", res_min, res_s)
                                            Console.WriteLine("-------------------------------------------------------")

                                            Exit For
                                        End If
                                    Next
                                End If

                                If check_jitter.Checked = True Then
                                    For y = 0 To data_jitter_iout.Rows.Count - 1

                                        If iout_now = data_jitter_iout.Rows(y).Cells(0).Value Then
                                            jitter_iout_num = y
                                            Jitter_run()
                                            Exit For
                                        End If
                                    Next
                                End If

                            Next


                            '---------------------------------------------------------------------------------------------------------

                            ''----------------------------------------------------------------------------------
                            If ((check_stability.Checked = True) And (check_iout_up.Checked = True) And (data_test.Rows.Count > 0)) Or ((check_LineR.Checked = True) And (rbtn_lineR_test1.Checked = True)) Then
                                'Meter
                                If check_Efficiency.Checked = True Then
                                    If rbtn_meter_iin.Checked = True Then
                                        If check_iin.Checked = True Then
                                            Iin_Meter_initial(check_iin, cbox_IIN_meter, cbox_IIN_relay)
                                        End If
                                    Else
                                        INA226_Iin_initial(True) 'High Range
                                    End If

                                    'Check Iin Max

                                End If


                                'Meter set High
                                If (rbtn_meter_iout.Checked = True) And (cbox_Iout_meter.SelectedItem <> no_device) Then
                                    If check_iout.Checked = True Then
                                        Iout_Meter_initial(check_iout, cbox_Iout_meter, cbox_Iout_relay)
                                    End If
                                ElseIf rbtn_board_iout.Checked = True Then

                                    If iout_now > INA226_Iout_max_L Then
                                        Iout_Meter_Max = True
                                    Else
                                        Iout_Meter_Max = False
                                    End If


                                End If
                            End If




                            '--------------------------------------------------------------------------------------------------------

                            'Stability line up

                            If (check_stability.Checked = True) And (check_iout_up.Checked = True) And (data_test.Rows.Count > 0) Then

                                For y = stability_iout.Length - 2 To 0 Step -1


                                    System.Windows.Forms.Application.DoEvents()

                                    While pause = True
                                        System.Windows.Forms.Application.DoEvents()


                                        If run = False Then
                                            Exit While
                                        End If
                                    End While

                                    If run = False Then
                                        Exit For
                                    End If

                                    test_point_num = test_point_num + 1
                                    txt_test_now.Text = test_point_num



                                    iout_now = stability_iout(y)

                                    'Iout

                                    'DCLoad_Iout(iout_now, monitor_vout, DUT2_en)
                                    Iout_Setting(iout_now, monitor_vout, DUT2_en)
                                    'DCLoad_Iout(iout_now, monitor_vout)
                                    If (DCLoad_ON = False) Then
                                        DCLoad_ONOFF("ON")
                                        DCLoad_ONOFF("ON", DUT2_en)
                                    End If

                                    ''----------------------------------------------------------------------------------
                                    'Vin Sense
                                    If check_vin_sense.Checked = True Then
                                        'Vin Sense
                                        vin_power_sense(cbox_vin.SelectedItem, num_vin_sense.Value, num_vin_max.Value, vin_now)
                                    End If

                                    If check_vin_sense2.Checked = True And DUT2_en Then
                                        'Vin Sense
                                        vin_power_sense(cbox_vin2.SelectedItem, num_vin_sense2.Value, num_vin_max2.Value, vin_now)
                                    End If


                                    ''----------------------------------------------------------------------------------
                                    'Measure

                                    If (iout_now > num_iout_delay.Value) And (num_delay.Value > 0) Then
                                        If cbox_delay_unit.SelectedIndex = 1 Then
                                            Delay_s(num_delay.Value)
                                        Else
                                            Delay(num_delay.Value)
                                        End If
                                    End If

                                    ''----------------------------------------------------------------------------------
                                    'Check Vout
                                    'vout

                                    vout_meas = DAQ_average(vout_daq, num_data_count.Value)
                                    If DUT2_en Then
                                        vout_meas2 = DAQ_average(vout_daq2, num_data_count.Value)
                                    End If

                                    '-------------------------------------------------------------------------
                                    '2020/01/22  Vout用DC量測時 Scope的offset要由DAQ校正
                                    If (check_offset_vout.Checked = True) And (cbox_coupling_vout.SelectedItem <> "AC") Then
                                        CHx_OFFSET(vout_ch, vout_meas)
                                    End If

                                    ''----------------------------------------------------------------------------------
                                    stability_iout_num = stability_iout_num + 1
                                    Stability_run()
                                Next
                            End If
                            excel_close()
                        Next 'vin
                    End If

                    'Line Regulation
                    If (check_LineR.Checked = True) And (rbtn_lineR_test1.Checked = True) Then
                        excel_open()

                        DCLoad_ONOFF("OFF")
                        DCLoad_ONOFF("OFF", DUT2_en)
                        vin_now = data_lineR_vin.Rows(0).Cells(0).Value
                        '-----------------------------------------------------------------------------------------------------------

                        Power_Dev = vin_Dev
                        power_volt(vin_device, Vin_out, vin_now)


                        If DUT2_en Then
                            Power_Dev = vin_Dev2
                            power_volt(vin_device2, Vin_out2, vin_now)
                        End If

                        'DCLoad_Iout(0, monitor_vout, DUT2_en)
                        Iout_Setting(0, monitor_vout, DUT2_en)
                        'DCLoad_Iout(0, monitor_vout)
                        DCLoad_ONOFF("ON")
                        DCLoad_ONOFF("ON", DUT2_en)



                        If check_lineR_scope.Checked = True Then

                            Scope_RUN(False)

                            If RS_Scope = True Then

                                RS_Display(RS_RES_MES, RS_DISP_DOCK)
                                'RS_Display(RS_RES_MES, RS_DISP_PREV)
                                Scope_measure_clear()

                                RS_Scope_measure_status(meas1, True)
                                RS_Scope_measure_status(meas2, True)
                                RS_Scope_measure_status(meas3, True)

                                If DUT2_en Then
                                    RS_Scope_measure_status(meas4, True)
                                    RS_Scope_measure_status(meas5, True)
                                    RS_Scope_measure_status(meas6, True)
                                End If

                                RS_Local()
                                'RS_View()
                            End If

                            '-------------------------------------------------------------------------
                            'Scope
                            'Time Scale
                            '以8格算

                            'If fs_now <> 0 Then
                            '    H_scale_value = ((1 / fs_now) * 2 / 8) * (10 ^ 9)


                            '    'Timing Scale
                            '    H_scale(H_scale_value, "ns") '1/Fs_Min(Hz)*n/10 
                            'End If


                            'CHx_scale(lx_ch, (vin_now / num_lx_scale.Value), "V") 'Voltage Scale > SW/2

                            'If rbtn_vin_trigger.Checked = True Then
                            '    Trigger_set(lx_ch, "R", vin_now / num_vin_trigger.Value)
                            'Else
                            '    Scope_RUN(True)
                            '    Trigger_set(lx_ch, "R", vin_now / 2)

                            '    Trigger_auto_level(lx_ch, "R")
                            'End If

                            If RS_Scope = True Then
                                RS_View(True)
                            Else
                                FastAcq_ONOFF("OFF")
                            End If
                            RUN_set("RUNSTop")
                        End If

                        For x = 0 To data_lineR_iout.Rows.Count - 1
                            System.Windows.Forms.Application.DoEvents()
                            While pause = True
                                System.Windows.Forms.Application.DoEvents()
                                If run = False Then
                                    Exit While
                                End If
                            End While

                            If run = False Then
                                Exit For
                            End If

                            iout_now = data_lineR_iout.Rows(x).Cells(0).Value
                            'DCLoad_Iout(iout_now, monitor_vout, DUT2_en)
                            Iout_Setting(iout_now, monitor_vout, DUT2_en)
                            'DCLoad_Iout(iout_now, monitor_vout)
                            lineR_iout_num = x

                            For v = 0 To data_lineR_vin.Rows.Count - 1

                                System.Windows.Forms.Application.DoEvents()

                                While pause = True
                                    System.Windows.Forms.Application.DoEvents()
                                    If run = False Then
                                        Exit While
                                    End If
                                End While
                                If run = False Then
                                    Exit For
                                End If

                                test_point_num = test_point_num + 1
                                txt_test_now.Text = test_point_num
                                vin_now = data_lineR_vin.Rows(v).Cells(0).Value
                                LineR_run(DUT2_en)
                                LR_Vin_test_num = v
                                update_report(Line_Regulation, DUT2_en)
                            Next

                            If check_lineR_up.Checked = True Then
                                For v = data_lineR_vin.Rows.Count - 2 To 0 Step -1
                                    System.Windows.Forms.Application.DoEvents()
                                    While pause = True
                                        System.Windows.Forms.Application.DoEvents()
                                        If run = False Then
                                            Exit While
                                        End If
                                    End While

                                    If run = False Then
                                        Exit For
                                    End If
                                    test_point_num = test_point_num + 1
                                    txt_test_now.Text = test_point_num
                                    vin_now = data_lineR_vin.Rows(v).Cells(0).Value
                                    LineR_run(DUT2_en)
                                    LR_Vin_test_num = LR_Vin_test_num + 1
                                    update_report(Line_Regulation, DUT2_en)
                                Next
                            End If
                        Next
                        excel_close()
                    End If
                Next
            Next
        Next

        excel_open()
        xlSheet = xlBook.Sheets(1)
        xlSheet.Activate()
        If TA_Test_num = 0 Then

            report_test_info()

        End If

        report_test_update(start_test_time, test_point_num - test_point_temp)


        xlBook.Save()

        excel_close()

        instrument_closed()




    End Function

    Function update_pic2report(ByVal pic_start_col As Integer, ByVal pic_start_row As Integer) As Integer
        Dim pic_format As String = ".PNG"
        Dim num_temp As Integer
        Dim update_row, update_col As Integer
        Dim temp() As String
        Dim height_temp As Double
        Dim width_temp As Double
        If (System.IO.Directory.Exists(Main.txt_folder.Text)) = True Then
            Note.lbl_title.Text = "Paste Pic to Report"
            Note.Show()

            Dim di As New IO.DirectoryInfo(Main.txt_folder.Text)
            Dim diar1 As IO.FileInfo() = di.GetFiles()
            Dim dra As IO.FileInfo

            'list the names of all files in the specified directory

            xlApp = CreateObject("Excel.Application") '?萄遣EXCEL撠情
            xlApp.DisplayAlerts = False

            '開啟或放大檔案會變大
            xlApp.WindowState = Excel.XlWindowState.xlMinimized

            xlApp.Visible = False
            xlBook = xlApp.Workbooks.Open(Main.txt_file.Text)
            xlBook.Activate()
            xlSheet = xlBook.Sheets(Main.txt_sheet.Text)
            xlSheet.Activate()
            'update_row = pic_start_row
            'update_col = pic_start_col
            For Each dra In diar1

                System.Windows.Forms.Application.DoEvents()


                ' If dra.Extension = pic_format Or dra.Extension = UCase(pic_format) Then
                If dra.Extension = pic_format Then
                    temp = Split(dra.Name, "_")
                    num_temp = temp(0)
                    Note.ProgressBar1.Value = num_temp / diar1.Count * 100
                    If (num_temp Mod 10) = 0 Then
                        update_col = pic_start_col + 9 * (pic_width + 1)
                        update_row = pic_start_row + (Int((num_temp / 10)) - 1) * (pic_height + 2)
                    Else
                        update_col = pic_start_col + ((num_temp Mod 10) - 1) * (pic_width + 1)
                        update_row = pic_start_row + Int((num_temp / 10)) * (pic_height + 2)
                    End If
                    'Title
                    ' ''------------------------------------------------------------
                    ' ''Update Picture Title
                    pic_top = ConvertToLetter(update_col) & (update_row - 1)
                    xlrange = xlSheet.Range(pic_top)
                    xlrange.Interior.ColorIndex = 45 'Orange
                    xlrange.Value2 = dra.Name
                    xlrange = xlSheet.Range(pic_top & ":" & ConvertToLetter(update_col + pic_width - 1) & (update_row - 1))
                    xlrange.MergeCells = True
                    FinalReleaseComObject(xlrange)
                    ' ''------------------------------------------------------------
                    ' ''Paste Picture
                    pic_top = ConvertToLetter(update_col) & update_row
                    xlrange = xlSheet.Range(pic_top & ":" & ConvertToLetter(update_col) & (update_row + pic_height - 1))
                    height_temp = xlrange.Height
                    FinalReleaseComObject(xlrange)
                    xlrange = xlSheet.Range(pic_top & ":" & ConvertToLetter(update_col + pic_width - 1) & update_row)
                    width_temp = xlrange.Width
                    FinalReleaseComObject(xlrange)
                    pic_ByteSize = FileLen(Main.txt_folder.Text & "\" & dra.Name)
                    If (pic_ByteSize > 0) Then
                        paste_picture(Main.txt_folder.Text & "\" & dra.Name, pic_top, width_temp, height_temp)
                        Delay(10)
                    End If
                    xlBook.Save()
                End If
            Next


            FinalReleaseComObject(xlSheet)

            excel_close()
            Note.Close()


        End If





    End Function

    Function jitter_update_pic(Optional ByVal dut2_en As Boolean = False) As Integer

        Dim pic_format As String = ".PNG"
        Dim num_temp As Integer
        Dim update_row, update_col As Integer
        Dim temp() As String
        Dim height_temp As Double
        Dim width_temp As Double

        Dim folder As String

        If dut2_en Then
            folder = Jitter_folder2
        Else
            folder = Jitter_folder
        End If


        If (System.IO.Directory.Exists(folder)) = True Then


            Note.lbl_title.Text = "Paste Pic to Report"
            Note.Show()
            Dim path As String = Jitter_folder
            If dut2_en Then
                path = Jitter_folder2
            Else
                path = Jitter_folder
            End If
            Dim di As New IO.DirectoryInfo(path)
            Dim diar1 As IO.FileInfo() = di.GetFiles()
            Dim dra As IO.FileInfo

            'list the names of all files in the specified directory

            xlApp = CreateObject("Excel.Application") '?萄遣EXCEL撠情
            xlApp.DisplayAlerts = False

            '開啟或放大檔案會變大
            xlApp.WindowState = Excel.XlWindowState.xlMinimized

            xlApp.Visible = False


            'xlApp.WindowState = Excel.XlWindowState.xlMaximized

            'xlApp.Visible = True
            xlBook = xlApp.Workbooks.Open(PartI_file)
            xlBook.Activate()
            xlSheet = xlBook.Sheets(txt_jitter_sheet.Text)

            If dut2_en Then
                xlSheet = xlBook.Sheets(txt_jitter_sheet.Text & add_dut2)
            End If
            xlSheet.Activate()
            'update_row = pic_start_row
            'update_col = pic_start_col
            For Each dra In diar1
                System.Windows.Forms.Application.DoEvents()
                ' If dra.Extension = pic_format Or dra.Extension = UCase(pic_format) Then
                If dra.Extension = pic_format Then
                    temp = Split(dra.Name, "_")
                    num_temp = temp(0)

                    If dut2_en Then
                        If (check_fastAcq.Checked = True) And (temp(1) = "Fast") Then
                            update_col = jitter_pic_col2(num_temp)
                            update_row = jitter_pic_row2(num_temp) + pic_height + 1
                        Else
                            update_col = jitter_pic_col2(num_temp)
                            update_row = jitter_pic_row2(num_temp) + 1
                        End If
                    Else
                        If (check_fastAcq.Checked = True) And (temp(1) = "Fast") Then
                            update_col = jitter_pic_col(num_temp)
                            update_row = jitter_pic_row(num_temp) + pic_height + 1
                        Else
                            update_col = jitter_pic_col(num_temp)
                            update_row = jitter_pic_row(num_temp) + 1
                        End If
                    End If

                    ' ''------------------------------------------------------------
                    ' ''Paste Picture
                    pic_top = ConvertToLetter(update_col) & update_row
                    xlrange = xlSheet.Range(pic_top & ":" & ConvertToLetter(update_col) & (update_row + pic_height - 1))
                    height_temp = xlrange.Height
                    FinalReleaseComObject(xlrange)
                    xlrange = xlSheet.Range(pic_top & ":" & ConvertToLetter(update_col + pic_width - 1) & update_row)
                    width_temp = xlrange.Width
                    FinalReleaseComObject(xlrange)
                    If dut2_en Then
                        pic_ByteSize = FileLen(Jitter_folder2 & "\" & dra.Name)
                    Else
                        pic_ByteSize = FileLen(Jitter_folder & "\" & dra.Name)
                    End If

                    If (pic_ByteSize > 0) Then
                        If dut2_en Then
                            paste_picture(Jitter_folder2 & "\" & dra.Name, pic_top, width_temp, height_temp)
                        Else
                            paste_picture(Jitter_folder & "\" & dra.Name, pic_top, width_temp, height_temp)
                        End If
                        Delay(10)
                    End If
                    xlBook.Save()
                End If
            Next
            FinalReleaseComObject(xlSheet)
            excel_close()
            Note.Close()
        End If

    End Function


    Private Sub PartI_Activated(sender As Object, e As EventArgs) Handles Me.Activated
        If Review_set = True Then
            reflesh()

            data_set_list()
            stability_parameter(0)
            'eff_parameter()
            result_parameter()
        End If

        If Save_set = True Then
            Tab_Set.Enabled = True
            Test_set()
        End If

        If (Open_set = True) And (import_now = False) Then
            import_now = True
            Tab_Set.Enabled = True
            ' initial()

            Test_import()
            Inst_check_list()

            'GC.Collect()
            result_parameter()
            Tab_Set.SelectedIndex = 6
            import_now = False

        End If



        If (Test_run = True) And (Me.Enabled = True) And (TestITem_run_now = False) Then

            Tab_Set.Enabled = False
            'data_set_list()
            data_list()
            TestITem_run()


            Tab_Set.Enabled = True
        End If


        If report_run = True And (Me.Enabled = True) And (TestITem_run_now = False) Then
            If (check_jitter.Checked = True) And (Jitter_folder <> "") Then

                'Error
                Tab_Set.Enabled = False

                jitter_update_pic()
                If DUT2_en Then
                    jitter_update_pic(DUT2_en)
                End If
                GC.Collect()
                GC.WaitForPendingFinalizers()
                Tab_Set.Enabled = True
                '----------------------------------------------------------------------------------
            End If
            If (check_stability.Checked = True) And (Error_folder <> "") Then

                'Error
                Tab_Set.Enabled = False
                If check_error_pic.Checked = True Then

                    Main.txt_folder.Text = Error_folder
                    Main.txt_file.Text = PartI_file
                    Main.txt_sheet.Text = txt_error_sheet.Text
                    update_pic2report(1, 4)

                    If DUT2_en Then
                        Main.txt_folder.Text = Error_folder2
                        Main.txt_file.Text = PartI_file
                        Main.txt_sheet.Text = txt_error_sheet.Text & add_dut2
                        update_pic2report(1, 4)
                    End If



                    GC.Collect()
                    GC.WaitForPendingFinalizers()
                End If


                Tab_Set.Enabled = True
                '----------------------------------------------------------------------------------
            End If


        End If

    End Sub

    Private Sub PartI_Click(sender As Object, e As EventArgs) Handles Me.Click

    End Sub

    Private Sub PartI_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        'Auto_Scanning(2)
        'GetVoutMax_Min(1)

        initial()


        scope_init_set()


        cbox_type_stability.SelectedIndex = 0

        cbox_type_Eff.SelectedIndex = 0

        cbox_type_LoadR.SelectedIndex = 0

        cbox_type_LineR.SelectedIndex = 0

        cbox_icc_range.SelectedIndex = 0

        If vin_daq = "" Then
            cbox_vin_daq.SelectedIndex = 0
        Else
            cbox_vin_daq.SelectedItem = vin_daq
        End If

        If vout_daq = "" Then
            cbox_vout_daq.SelectedIndex = 1
        Else
            cbox_vout_daq.SelectedItem = vout_daq
        End If

        If Eff_vout_daq = "" Then
            cbox_vout1_daq.SelectedIndex = 1
        Else
            cbox_vout1_daq.SelectedItem = Eff_vout_daq
        End If

        If vcc_daq = "" Then

            cbox_VCC_daq.SelectedItem = no_device
        Else
            cbox_VCC_daq.SelectedItem = vcc_daq
        End If



        cbox_data_resolution.SelectedIndex = 0
        cbox_data_resolution2.SelectedIndex = 0
        cbox_delay_unit.SelectedIndex = 0
        cbox_meter_mini.SelectedIndex = 0

        cbox_board_buck.SelectedIndex = 0

        PartI_first = False
    End Sub

    Private Sub Tab_Set_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Tab_Set.SelectedIndexChanged
        Dim tpg As TabPage



        tpg = Tab_Set.SelectedTab
        Select Case tpg.Name


            Case "TabPage_Main1"


            Case "TabPage_Instrument1"


            Case "TabPage_Scope1"


            Case "TabPage_Test1"

                data_set_list()


            Case "TabPage_Setup1"

                eff_parameter()

            Case "TabPage_LineR"


            Case "TabPage_Finish1"

                result_parameter()


        End Select



    End Sub

    Private Sub btn_refresh_Click(sender As Object, e As EventArgs) Handles btn_refresh.Click


        reflesh()

        data_set_list()
        stability_parameter(0)
        'eff_parameter()
        result_parameter()


    End Sub


    Private Sub cbox_VCC_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbox_VCC.SelectedIndexChanged
        Dim addr() As String

        power_channel_set(cbox_VCC, cbox_VCC_ch)
        If cbox_VCC.SelectedItem = no_device Then
            txt_vcc_Addr.Text = ""
            vcc_dev_ch = 0
            pic_vcc.Visible = True
            txt_vcc.Visible = False
        Else
            addr = Split(Power_addr(cbox_VCC.SelectedIndex), "::")
            txt_vcc_Addr.Text = addr(1)
            pic_vcc.Visible = False
            txt_vcc.Visible = True
        End If

        cbox_VCC_ch.SelectedIndex = vcc_dev_ch
    End Sub

    Private Sub cbox_vin_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbox_vin.SelectedIndexChanged
        Dim addr() As String

        power_channel_set(cbox_vin, cbox_vin_ch)
        If cbox_vin.SelectedItem = no_device Then
            txt_vin_addr.Text = ""
            vin_dev_ch = 0
        Else
            addr = Split(Power_addr(cbox_vin.SelectedIndex), "::")
            txt_vin_addr.Text = addr(1)
        End If

        cbox_vin_ch.SelectedIndex = vin_dev_ch
    End Sub

    Private Sub cbox_VCC_ch_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbox_VCC_ch.SelectedIndexChanged
        vcc_dev_ch = cbox_VCC_ch.SelectedIndex
    End Sub

    Private Sub cbox_vin_ch_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbox_vin_ch.SelectedIndexChanged
        vin_dev_ch = cbox_vin_ch.SelectedIndex
    End Sub


    Private Sub cbox_IIN_meter_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbox_IIN_meter.SelectedIndexChanged
        Dim addr() As String
        'If PartI_first = True Then
        '    Exit Sub

        'End If

        num_iin_change.Maximum = 0.4 * 1000
        If cbox_IIN_meter.SelectedItem = no_device Then
            txt_IIN_addr.Text = ""
            check_iin.Checked = False

        Else
            check_iin.Checked = True
            addr = Split(Meter_addr(cbox_IIN_meter.SelectedIndex), "::")
            txt_IIN_addr.Text = addr(1)


            'Check Relay
            Select Case Mid(cbox_IIN_meter.SelectedItem, 1, 5)
                Case "34450"
                    '10A
                    Meter_iin_Max = 0.1
                    Meter_iin_relay(0) = 1 'H
                    Meter_iin_relay(1) = 0 'L
                    Meter_iin_range = "AUTO"

                Case "DMM40"
                    '10A
                    Meter_iin_Max = 0.4
                    Meter_iin_relay(0) = 0 'H
                    Meter_iin_relay(1) = 1 'L
                    Meter_iin_range = "MAX"

                Case "DMM65"
                    '10A
                    Meter_iin_Max = 3
                    Meter_iin_relay(0) = 0 'H
                    Meter_iin_relay(1) = 1 'L
                    Meter_iin_range = "MAX"
                    check_iin.Checked = False

            End Select

            num_iin_change.Maximum = Meter_iin_Max * 1000

            If (num_iin_change.Value > Meter_iin_Max * 1000) Or (num_iin_change.Value = 0) Then
                num_iin_change.Value = (Meter_iin_Max * 0.8 * 1000)
            End If


        End If




    End Sub

    Private Sub cbox_Iout_meter_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbox_Iout_meter.SelectedIndexChanged
        Dim addr() As String
        'If PartI_first = True Then
        '    Exit Sub

        'End If
        num_iout_change.Maximum = (0.4 * 1000)
        If cbox_Iout_meter.SelectedItem = no_device Then
            txt_Iout_addr.Text = ""
            check_iout.Checked = False

        Else
            check_iout.Checked = True
            addr = Split(Meter_addr(cbox_Iout_meter.SelectedIndex), "::")
            txt_Iout_addr.Text = addr(1)
            'Check Relay


            Select Case Mid(cbox_Iout_meter.SelectedItem, 1, 5)
                Case "34450"
                    '10A
                    Meter_iout_Max = 0.1
                    Meter_iout_relay(0) = 1 'H
                    Meter_iout_relay(1) = 0 'L
                    Meter_iout_range = "AUTO"

                Case "DMM40"
                    '10A
                    Meter_iout_Max = 0.4
                    Meter_iout_relay(0) = 0 'H
                    Meter_iout_relay(1) = 1 'L
                    Meter_iout_range = "MAX"

                Case "DMM65"
                    '10A
                    Meter_iout_Max = 3
                    Meter_iout_relay(0) = 0 'H
                    Meter_iout_relay(1) = 1 'L
                    Meter_iout_range = "MAX"
                    check_iout.Checked = False

            End Select
            num_iout_change.Maximum = (Meter_iout_Max * 1000)

            If (num_iout_change.Value > (Meter_iout_Max * 1000)) Or (num_iout_change.Value = 0) Then
                num_iout_change.Value = (Meter_iout_Max * 0.9 * 1000)
            End If


        End If

    End Sub



    Private Sub cbox_Icc_meter_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbox_Icc_meter.SelectedIndexChanged
        Dim addr() As String
        If PartI_first = True Then
            Exit Sub

        End If

        If cbox_Icc_meter.SelectedItem = no_device Then
            txt_Icc_addr.Text = ""
        Else
            addr = Split(Meter_addr(cbox_Icc_meter.SelectedIndex), "::")
            txt_Icc_addr.Text = addr(1)
            'Check Relay
            Select Case Mid(cbox_Icc_meter.SelectedItem, 1, 5)
                Case "34450"
                    '10A

                    Meter_icc_range = "AUTO"

                Case "DMM40"

                    Meter_icc_range = cbox_icc_range.SelectedItem

                Case "DMM65"

                    Meter_icc_range = cbox_icc_range.SelectedItem

            End Select

        End If
    End Sub

    Private Sub btn_cancel_Click(sender As Object, e As EventArgs) Handles btn_cancel.Click

        Add_test = False
        If btn_ok.Enabled = False Then
            Me.Close()
        Else
            Me.Hide()
        End If
    End Sub

    Private Sub btn_ok_Click(sender As Object, e As EventArgs) Handles btn_ok.Click

        Dim power_temp As String
        Dim v As Integer

        power_temp = Mid(cbox_vin.SelectedItem, 1, 6)


        If ((power_temp = "62006P") Or (power_temp = "62012P")) And (num_VIN_OCP.Value = 0) Then
            error_message("Please enter the OCP value of VIN!!")
            Exit Sub
        End If


        If check_vin_sense.Checked = True Then


            If num_vin_max.Value = 0 Then
                error_message("Please enter the VIN Max value!!")
                Exit Sub
            Else
                For v = 0 To data_vin.Rows.Count - 1

                    If data_vin.Rows(v).Cells(0).Value > num_vin_max.Value Then
                        error_message("VIN > VIN MAX !!")
                        Exit Sub
                    End If

                Next

                For v = 0 To data_lineR_vin.Rows.Count - 1

                    If data_lineR_vin.Rows(v).Cells(0).Value > num_vin_max.Value Then
                        error_message("VIN > VIN MAX !!")
                        Exit Sub
                    End If

                Next


            End If


        End If

        If check_Efficiency.Checked = True Or check_loadR.Checked = True Then
            If data_eff_iout.Rows.Count = 0 Then
                error_message("Please enter the Iout test value of Efficiency/Load Regulation!!")
                Exit Sub
            End If


        End If



        If check_LineR.Checked = True Then
            If data_lineR_iout.Rows.Count = 0 Then
                error_message("Please enter the Iout test value of Line Regulation!!")
                Exit Sub
            End If

            If rbtn_lineR_test1.Checked = True And data_lineR_vin.Rows.Count = 0 Then
                error_message("Please enter the Vin test value of Line Regulation!!")
                Exit Sub
            End If

        End If


        If check_stability.Checked = True Then
            If data_test.Rows.Count = 0 Then
                error_message("Please enter the Iout test value of Stability!!")
                Exit Sub
            End If
        End If


        If check_jitter.Checked = True Then
            If data_jitter_iout.Rows.Count = 0 Then
                error_message("Please enter the Iout test value of Jitter!!")
                Exit Sub
            End If
        End If

        'If check_vin_sense.Checked = True And rbtn_meter_iin.Checked = True And cbox_vin.SelectedItem = no_device Then
        '    error_message("Please check IIN meter!!")
        '    Exit Sub
        'End If

        'If check_Efficiency.Checked = True And rbtn_Iin_PW.Checked = True Then
        '    error_message("Please Iin choose meter or relay board!!")
        '    Exit Sub
        'End If



        If Add_test = True Then

            Main.data_Test.Rows.Add(True, Me.Name)

            PartI_num = PartI_num + 1
            If PartI_num > 1 Then
                txt_stability_sheet.Text = "Stability_" & PartI_num
                txt_error_sheet.Text = "Stability_Error_" & PartI_num
                txt_beta_sheet.Text = "Stability_Beta_" & PartI_num
                txt_eff_sheet.Text = "Efficiency" & PartI_num
                txt_LoadR_sheet.Text = "Load Regulation" & PartI_num
                txt_LineR_sheet.Text = "Line Regulation" & PartI_num
                txt_jitter_sheet.Text = "Jitter_" & PartI_num
            End If




            data_test_now = Main.data_Test.Rows.Count - 1
            Add_test = False

        End If

        in_high_id = num_slave_in_H.Value
        in_middle_id = num_slave_in_M.Value
        in_low_id = num_slave_in_L.Value
        in_io_id = num_slave_in_IO.Value

        in_high_id2 = num_slave_in_H2.Value
        in_middle_id2 = num_slave_in_M2.Value
        in_low_id2 = num_slave_in_L2.Value
        in_io_id2 = num_slave_in_IO2.Value

        in_high_comp = num_comp_in_H.Value
        in_middle_comp = num_comp_in_M.Value
        in_low_comp = num_comp_in_L.Value

        in_high_resolution = num_resolution_in_H.Value
        in_middle_resolution = num_resolution_in_M.Value
        in_low_resolution = num_resolution_in_L.Value


        out_high_id = num_slave_out_H.Value
        out_middle_id = num_slave_out_M.Value
        out_low_id = num_slave_out_L.Value
        out_io_id = num_slave_out_IO.Value

        out_high_comp = num_comp_out_H.Value
        out_middle_comp = num_comp_out_M.Value
        out_low_comp = num_comp_out_L.Value

        out_high_resolution = num_resolution_out_H.Value
        out_middle_resolution = num_resolution_out_M.Value
        out_low_resolution = num_resolution_out_L.Value

        DUT2_en = check_DUT2.Checked

        Inst_check_list()
        Me.Hide()
    End Sub

    Private Sub cbox_vin_daq_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbox_vin_daq.SelectedIndexChanged
        vin_daq = Mid(cbox_vin_daq.SelectedItem, 3)
    End Sub

    Private Sub cbox_vout_daq_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbox_vout_daq.SelectedIndexChanged
        vout_daq = Mid(cbox_vout_daq.SelectedItem, 3)
    End Sub

    Private Sub cbox_vout1_daq_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbox_vout1_daq.SelectedIndexChanged
        Eff_vout_daq = Mid(cbox_vout1_daq.SelectedItem, 3)
    End Sub

    Private Sub cbox_VCC_daq_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbox_VCC_daq.SelectedIndexChanged
        vcc_daq = Mid(cbox_VCC_daq.SelectedItem, 3)
    End Sub

    Private Sub rbtn_board_iin_CheckedChanged(sender As Object, e As EventArgs) Handles rbtn_board_iin.CheckedChanged
        If rbtn_board_iin.Checked = True Then
            Meter_iin_Max = INA226_Iin_max_L
            'num_iin_change.Maximum = Meter_iin_Max * 1000
            num_iin_change.Value = Meter_iin_Max * 1000 - 5 'mA
            check_iin.Checked = False

            Panel_meter_mini.Visible = False
            num_iin_step.Value = 1


            txt_board_VIN.Text = "VIN"
        End If
    End Sub

    Private Sub rbtn_meter_iin_CheckedChanged(sender As Object, e As EventArgs) Handles rbtn_meter_iin.CheckedChanged
        If rbtn_meter_iin.Checked = True Then


            ' num_iin_change.Maximum = (Meter_iin_Max) * 1000
            num_iin_change.Value = (Meter_iin_Max * 0.8 * 1000) 'mA
            Panel_meter_mini.Visible = True
            txt_board_VIN.Text = ""

            num_iin_step.Value = 10
        End If


    End Sub

    Private Sub rbtn_meter_iout_CheckedChanged(sender As Object, e As EventArgs) Handles rbtn_meter_iout.CheckedChanged
        If rbtn_meter_iout.Checked = True Then
            ' num_iout_change.Maximum = Meter_iout_Max * 1000
            check_IOUT_ch2.Checked = False
            check_IOUT_ch4.Checked = False
            num_iout_change.Value = (Meter_iout_Max * 0.8 * 1000)

            check_IOUT_ch2.Enabled = True
            check_IOUT_ch4.Enabled = True
            txt_board_VOUT.Text = ""

        End If
    End Sub

    Private Sub rbtn_board_iout_CheckedChanged(sender As Object, e As EventArgs) Handles rbtn_board_iout.CheckedChanged

        If rbtn_board_iout.Checked = True Then
            check_IOUT_ch2.Enabled = False
            check_IOUT_ch4.Enabled = False

            If (check_IOUT_ch1.Checked = True) Then
                check_IOUT_ch2.Checked = True
            Else
                check_IOUT_ch3.Checked = True
                check_IOUT_ch4.Checked = True
            End If

            check_iout.Checked = False
            Meter_iout_Max = INA226_Iout_max_L
            ' num_iout_change.Maximum = Meter_iout_Max * 1000
            num_iout_change.Value = Meter_iout_Max * 1000 - 5


            txt_board_VOUT.Text = cbox_board_buck.SelectedItem

            'LX
            cbox_channel_lx.SelectedIndex = 2 'CH3
            cbox_channel_iout.SelectedIndex = 3 'CH4

            If cbox_board_buck.SelectedIndex = 0 Then
                'Buck1
                cbox_channel_vout.SelectedIndex = 0 'CH1
                cbox_channel_vin.SelectedIndex = 1 'CH2

            Else
                'Buck2
                cbox_channel_vout.SelectedIndex = 1 'CH2
                cbox_channel_vin.SelectedIndex = 0 'CH1

            End If

            Iout_board_EN = True
        End If


    End Sub

    Private Sub rbtn_iout_load_CheckedChanged(sender As Object, e As EventArgs) Handles rbtn_iout_load.CheckedChanged
        If rbtn_iout_load.Checked = True Then
            check_IOUT_ch2.Enabled = True
            check_IOUT_ch4.Enabled = True
            txt_board_VIN.Text = ""
            txt_board_VOUT.Text = ""
        End If
    End Sub

    Private Sub txt_vcc_name1_TextChanged(sender As Object, e As EventArgs) Handles txt_vcc_name1.TextChanged
        txt_vcc_name.Text = txt_vcc_name1.Text & " (V)"
    End Sub

    Private Sub txt_ivcc_name1_TextChanged(sender As Object, e As EventArgs) Handles txt_ivcc_name1.TextChanged
        txt_ivcc_name.Text = txt_ivcc_name1.Text & " (A)"
    End Sub

    Private Sub txt_vin_name1_TextChanged(sender As Object, e As EventArgs) Handles txt_vin_name1.TextChanged
        txt_vin_name.Text = txt_vin_name1.Text & " (V)"
    End Sub

    Private Sub txt_iin_name1_TextChanged(sender As Object, e As EventArgs) Handles txt_iin_name1.TextChanged
        txt_iin_name.Text = txt_iin_name1.Text & " (A)"
    End Sub

    Private Sub txt_vout_name1_TextChanged(sender As Object, e As EventArgs) Handles txt_vout_name1.TextChanged
        txt_vout_name.Text = txt_vout_name1.Text & " (V)"
    End Sub

    Private Sub txt_iout_name1_TextChanged(sender As Object, e As EventArgs) Handles txt_iout_name1.TextChanged
        txt_iout_name.Text = txt_iout_name1.Text & " (A)"
    End Sub

    Private Sub txt_vcc_name_TextChanged(sender As Object, e As EventArgs) Handles txt_vcc_name.TextChanged

        Vcc_name = txt_vcc_name.Text
        If data_VCC.Columns.Count > 0 And data_set.Columns.Count > 0 And data_result.Columns.Count > 0 Then
            data_VCC.Columns("col_VCC").HeaderText = Vcc_name
            data_set.Columns("col_VCC1").HeaderText = Vcc_name
            data_result.Columns("col_test_vcc1").HeaderText = Vcc_name
        End If
    End Sub

    Private Sub txt_ivcc_name_TextChanged(sender As Object, e As EventArgs) Handles txt_ivcc_name.TextChanged
        Icc_name = txt_ivcc_name.Text
    End Sub

    Private Sub txt_vin_name_TextChanged(sender As Object, e As EventArgs) Handles txt_vin_name.TextChanged

        Vin_name = txt_vin_name.Text
        If data_vin.Columns.Count > 0 And data_set.Columns.Count > 0 And data_result.Columns.Count > 0 Then
            data_vin.Columns("col_VIN").HeaderText = Vin_name
            data_set.Columns("col_Vin1").HeaderText = Vin_name
            data_result.Columns("col_test_vin1").HeaderText = Vin_name
        End If

    End Sub

    Private Sub txt_iin_name_TextChanged(sender As Object, e As EventArgs) Handles txt_iin_name.TextChanged
        Iin_name = txt_iin_name.Text
    End Sub

    Private Sub txt_vout_name_TextChanged(sender As Object, e As EventArgs) Handles txt_vout_name.TextChanged

        Vout_name = txt_vout_name.Text
        txt_vout_test.Text = Vout_name
        If data_set.Columns.Count > 0 And data_result.Columns.Count > 0 Then
            data_set.Columns("col_Vout1").HeaderText = Vout_name
            data_result.Columns("col_test_vout1").HeaderText = Vout_name
        End If

    End Sub

    Private Sub txt_iout_name_TextChanged(sender As Object, e As EventArgs) Handles txt_iout_name.TextChanged
        Iout_name = txt_iout_name.Text
    End Sub

    Private Sub cbox_channel_lx_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbox_channel_lx.SelectedIndexChanged, cbox_channel_lx2.SelectedIndexChanged
        txt_meas1_ch.Text = cbox_channel_lx.SelectedItem & "/" & Mid(cbox_channel_lx2.SelectedItem, 3)
        txt_meas2_ch.Text = cbox_channel_lx.SelectedItem & "/" & Mid(cbox_channel_lx2.SelectedItem, 3)
        txt_meas3_ch.Text = cbox_channel_lx.SelectedItem & "/" & Mid(cbox_channel_lx2.SelectedItem, 3)
    End Sub

    Private Sub cbox_channel_vout_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbox_channel_vout.SelectedIndexChanged, cbox_channel_vout2.SelectedIndexChanged
        txt_meas4_ch.Text = cbox_channel_vout.Text & "/" & Mid(cbox_channel_vout2.Text, 3)
        txt_meas5_ch.Text = cbox_channel_vout.Text & "/" & Mid(cbox_channel_vout2.Text, 3)
        txt_meas6_ch.Text = cbox_channel_vout.Text & "/" & Mid(cbox_channel_vout2.Text, 3)
    End Sub


    Private Sub cbox_coupling_vout_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbox_coupling_vout.SelectedIndexChanged
        If cbox_coupling_vout.SelectedItem = "AC" Then
            check_offset_vout.Checked = False
        Else
            check_offset_vout.Checked = True
        End If
    End Sub



    Private Sub cbox_coupling_lx_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbox_coupling_lx.SelectedIndexChanged
        If cbox_coupling_lx.SelectedItem = "AC" Then
            num_offset_lx.Value = 0

        End If
    End Sub

    Private Sub cbox_coupling_iout_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbox_coupling_iout.SelectedIndexChanged
        If cbox_coupling_iout.SelectedItem = "AC" Then
            num_offset_iout.Value = 0

        End If
    End Sub

    Private Sub cbox_coupling_vin_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbox_coupling_vin.SelectedIndexChanged
        If cbox_coupling_vin.SelectedItem = "AC" Then
            num_offset_vin.Value = 0

        End If
    End Sub




    Private Sub cbox_board_buck_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbox_board_buck.SelectedIndexChanged
        If rbtn_board_iout.Checked = True Then
            txt_board_VOUT.Text = cbox_board_buck.SelectedItem
            txt_board_VIN.Text = "VIN"
            'LX
            cbox_channel_lx.SelectedIndex = 2 'CH3
            cbox_channel_iout.SelectedIndex = 3 'CH4

            If cbox_board_buck.SelectedIndex = 0 Then
                'Buck1
                cbox_channel_vout.SelectedIndex = 0 'CH1
                cbox_channel_vin.SelectedIndex = 1 'CH2

            Else
                'Buck2
                cbox_channel_vout.SelectedIndex = 1 'CH2
                cbox_channel_vin.SelectedIndex = 0 'CH1

            End If
        Else
            txt_board_VIN.Text = ""
            txt_board_VOUT.Text = ""
        End If
    End Sub

    Private Sub rbtn_iin_auto_CheckedChanged(sender As Object, e As EventArgs) Handles rbtn_iin_auto.CheckedChanged
        If PartI_first = False Then

            If rbtn_iin_auto.Checked = True Then
                Panel_iin_auto.Enabled = True
                data_eff.Columns(2).HeaderText = "IOUT_Start (mA)"
            ElseIf rbtn_iin_manual.Checked = True Then
                Panel_iin_auto.Enabled = False
                data_eff.Columns(2).HeaderText = "IOUT (mA)"
            End If
            data_eff.Rows.Clear()

            eff_parameter()


        End If


    End Sub

    Private Sub rbtn_iin_manual_CheckedChanged(sender As Object, e As EventArgs) Handles rbtn_iin_manual.CheckedChanged
        If PartI_first = False Then

            If rbtn_iin_auto.Checked = True Then
                Panel_iin_auto.Enabled = True
                data_eff.Columns(2).HeaderText = "IOUT_Start (mA)"
            ElseIf rbtn_iin_manual.Checked = True Then
                Panel_iin_auto.Enabled = False
                data_eff.Columns(2).HeaderText = "IOUT (mA)"
            End If

            data_eff.Rows.Clear()
            eff_parameter()
        End If
    End Sub


    Private Sub btn_vcc_add_Click(sender As Object, e As EventArgs) Handles btn_vcc_add.Click

        data_value_add(data_VCC, num_Vcc_volt, 2)

    End Sub

    Private Sub data_vin_RowsRemoved(sender As Object, e As DataGridViewRowsRemovedEventArgs) Handles data_vin.RowsRemoved
        data_set_list()

    End Sub


    Private Sub btn_vin_add_Click(sender As Object, e As EventArgs) Handles btn_vin_add.Click
        If (check_vin_sense.Checked = True) And (num_vin.Value > num_vin_max.Value) Then
        Else
            If (check_stability.Checked = True) And (check_Force_CCM.Checked = False) Then
                If (num_vin.Value = 0) Then
                    error_message("VIN (V) cannot be ""0"".")
                    Exit Sub
                End If

                For i = 0 To clist_fs.Items.Count - 1
                    If clist_fs.GetItemChecked(i) = True Then
                        If clist_fs.Items(i) = 0 Then
                            error_message("Fs (kHz) cannot be ""0"".")
                            Exit Sub
                        End If

                    End If
                Next



                For i = 0 To clist_vout.Items.Count - 1
                    If clist_vout.GetItemChecked(i) = True Then

                        If clist_vout.Items(i) = 0 Then
                            error_message("VOUT (V) cannot be ""0"".")
                            Exit Sub
                        End If

                    End If
                Next



            End If

            data_value_add(data_vin, num_vin, 3)

            data_set_list()
        End If




    End Sub

    Private Sub btn_jitter_add_Click(sender As Object, e As EventArgs) Handles btn_jitter_add.Click
        data_value_add(data_jitter_iout, num_jitter_iout, 2)
    End Sub

    Private Sub btn_lineR_add_Click(sender As Object, e As EventArgs) Handles btn_lineR_add.Click
        data_value_add(data_lineR_iout, num_lineR_iout, 4)
    End Sub




    Private Sub data_set_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles data_set.CellClick



        If e.RowIndex >= 0 Then
            stability_parameter(e.RowIndex)
        End If




    End Sub

    Private Sub data_set_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles data_set.CellEndEdit
        If e.RowIndex >= 0 Then
            test_ton(e.RowIndex) = data_set.Rows(e.RowIndex).Cells(5).Value
            test_fs0(e.RowIndex) = data_set.Rows(e.RowIndex).Cells(6).Value
            test_IOB_start(e.RowIndex) = data_set.Rows(e.RowIndex).Cells(7).Value
            test_IOB_stop(e.RowIndex) = data_set.Rows(e.RowIndex).Cells(8).Value
            stability_parameter(e.RowIndex)
        End If

    End Sub

    Private Sub check_Force_CCM_CheckedChanged(sender As Object, e As EventArgs) Handles check_Force_CCM.CheckedChanged

        data_set_list()

    End Sub

    Private Sub btn_iout_add_Click(sender As Object, e As EventArgs) Handles btn_iout_add.Click
        Dim iout_temp() As Double
        data_iout.Rows.Add(num_iout_start.Value, num_iout_stop.Value, num_iout_step.Value)
        data_iout.CurrentCell = data_iout.Rows(data_iout.Rows.Count - 1).Cells(0)

        If data_set.Rows.Count > 0 Then
            data_set.CurrentCell = data_set.Rows(data_set.Rows.Count - 1).Cells(0)
        End If



        stability_parameter(data_set.Rows.Count - 1)




        data_eff_iout.Rows.Clear()


        iout_temp = Calculate_iout(data_iout)

        For i = 0 To iout_temp.Length - 1
            data_eff_iout.Rows.Add(Math.Round(iout_temp(i), 4))
        Next

        num_iout_start.Value = num_iout_stop.Value

    End Sub

    Private Sub data_iout_RowsRemoved(sender As Object, e As DataGridViewRowsRemovedEventArgs) Handles data_iout.RowsRemoved
        Dim iout_temp() As Double

        If data_set.Rows.Count > 0 Then
            data_set.CurrentCell = data_set.Rows(data_set.Rows.Count - 1).Cells(0)
            stability_parameter(data_set.Rows.Count - 1)
        End If

        If data_iout.Rows.Count > 0 Then
            data_iout.CurrentCell = data_iout.Rows(data_iout.Rows.Count - 1).Cells(0)
            iout_temp = Calculate_iout(data_iout)
            data_eff_iout.Rows.Clear()

            For i = 0 To iout_temp.Length - 1
                data_eff_iout.Rows.Add(Math.Round(iout_temp(i), 4))
            Next
        Else

        End If





        num_iout_start.Value = 0
    End Sub

    Private Sub check_IOB_CheckedChanged(sender As Object, e As EventArgs) Handles check_IOB.CheckedChanged
        If data_set.Rows.Count > 0 Then
            data_set.CurrentCell = data_set.Rows(data_set.Rows.Count - 1).Cells(0)

            stability_parameter(data_set.Rows.Count - 1)
        End If


    End Sub

    Private Sub num_IOB_Range_ValueChanged(sender As Object, e As EventArgs) Handles num_IOB_Range.ValueChanged
        If data_set.Rows.Count > 0 Then
            data_set.CurrentCell = data_set.Rows(data_set.Rows.Count - 1).Cells(0)

            stability_parameter(data_set.Rows.Count - 1)
        End If
    End Sub

    Private Sub num_IOB_step_ValueChanged(sender As Object, e As EventArgs) Handles num_IOB_step.ValueChanged
        If data_set.Rows.Count > 0 Then
            data_set.CurrentCell = data_set.Rows(data_set.Rows.Count - 1).Cells(0)

            stability_parameter(data_set.Rows.Count - 1)
        End If
    End Sub

    Private Sub check_iout_up_CheckedChanged(sender As Object, e As EventArgs) Handles check_iout_up.CheckedChanged
        If data_set.Rows.Count > 0 Then
            data_set.CurrentCell = data_set.Rows(data_set.Rows.Count - 1).Cells(0)

            stability_parameter(data_set.Rows.Count - 1)
        End If
    End Sub


    Private Sub data_vin_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles data_vin.CellEndEdit
        If (check_stability.Checked = True) And (check_Force_CCM.Checked = False) Then
            If (num_vin.Value = 0) Then
                error_message("VIN (V) cannot be ""0"".")
                Exit Sub
            End If
        End If
        data_set_list()
    End Sub

    Private Sub btn_lineR_vin_Click(sender As Object, e As EventArgs) Handles btn_lineR_vin.Click
        Dim VIN_step As Double
        Dim VIN As Double
        Dim VIN_row As Integer

        If num_vin_step.Value = 0 Then
            If num_vin_start.Value > num_vin_stop.Value Then
                VIN_step = -1
            Else
                VIN_step = 1
            End If
        Else

            If num_vin_start.Value > num_vin_stop.Value Then
                VIN_step = -num_vin_step.Value
            Else
                VIN_step = num_vin_step.Value

            End If
        End If


        For VIN = num_vin_start.Value To num_vin_stop.Value Step VIN_step
            If data_lineR_vin.Rows.Count = 0 Then
                VIN_row = 0
            Else
                VIN_row = data_lineR_vin.SelectedCells(0).RowIndex + 1
            End If
            VIN = Format(VIN, "#0.000")

            If vin_max < VIN Then
                vin_max = VIN
            End If

            If vin_min > VIN Then
                vin_min = VIN
            End If

            data_lineR_vin.Rows.Insert(VIN_row, Format(VIN, "#0.000"))
            data_lineR_vin.CurrentCell = data_lineR_vin.Rows(VIN_row).Cells(0)



            If num_vin_step.Value = 0 Or (num_vin_start.Value = num_vin_stop.Value) Then
                Exit For
            End If
        Next

        data_list()

        'If (num_vin_max.Value = 0) Or (num_vin_max.Value < (vin_max + 2)) Then
        '    num_vin_max.Value = vin_max + 2
        'End If



    End Sub





    Private Sub rbtn_lineR_test1_CheckedChanged(sender As Object, e As EventArgs) Handles rbtn_lineR_test1.CheckedChanged
        If rbtn_lineR_test1.Checked = True Then
            check_lineR_scope.Enabled = True
        Else
            check_lineR_scope.Checked = False
            check_lineR_scope.Enabled = False
        End If
    End Sub

    Private Sub rbtn_lineR_test2_CheckedChanged(sender As Object, e As EventArgs) Handles rbtn_lineR_test2.CheckedChanged
        If rbtn_lineR_test1.Checked = True Then
            check_lineR_scope.Enabled = True
        Else
            check_lineR_scope.Checked = False
            check_lineR_scope.Enabled = False
        End If
    End Sub


    Private Sub check_IOUT_ch1_CheckedChanged(sender As Object, e As EventArgs) Handles check_IOUT_ch1.CheckedChanged
        If rbtn_board_iout.Checked = True Then
            check_IOUT_ch2.Enabled = False
            check_IOUT_ch4.Enabled = False

            If (check_IOUT_ch1.Checked = True) Then
                check_IOUT_ch2.Checked = True
            ElseIf (check_IOUT_ch1.Checked = False) Then
                check_IOUT_ch2.Checked = False

            End If

            If (check_IOUT_ch3.Checked = True) Then
                check_IOUT_ch4.Checked = True
            ElseIf (check_IOUT_ch3.Checked = False) Then
                check_IOUT_ch4.Checked = False
            End If

            If check_IOUT_ch1.Checked = True And check_IOUT_ch3.Checked = True Then
                cbox_board_buck.SelectedIndex = 2
            End If

        End If



    End Sub



    Private Sub check_IOUT_ch3_CheckedChanged(sender As Object, e As EventArgs) Handles check_IOUT_ch3.CheckedChanged
        If rbtn_board_iout.Checked = True Then
            check_IOUT_ch2.Enabled = False
            check_IOUT_ch4.Enabled = False

            If (check_IOUT_ch1.Checked = True) Then
                check_IOUT_ch2.Checked = True
            ElseIf (check_IOUT_ch1.Checked = False) Then
                check_IOUT_ch2.Checked = False

            End If

            If (check_IOUT_ch3.Checked = True) Then
                check_IOUT_ch4.Checked = True
            ElseIf (check_IOUT_ch3.Checked = False) Then
                check_IOUT_ch4.Checked = False
            End If

            If check_IOUT_ch1.Checked = True And check_IOUT_ch3.Checked = True Then
                cbox_board_buck.SelectedIndex = 2
            End If
        End If
    End Sub







    Private Sub check_lineR_scope_CheckedChanged(sender As Object, e As EventArgs) Handles check_lineR_scope.CheckedChanged
        Dim i As Integer


        If check_lineR_scope.Checked = True Then
            For i = 0 To clist_fs.Items.Count - 1
                If clist_fs.GetItemChecked(i) = True Then
                    If clist_fs.Items(i) = 0 Then
                        check_lineR_scope.Checked = False
                        error_message("Fs (kHz) cannot be ""0"".")
                        Exit Sub
                    End If

                End If
            Next



            For i = 0 To clist_vout.Items.Count - 1
                If clist_vout.GetItemChecked(i) = True Then

                    If clist_vout.Items(i) = 0 Then
                        check_lineR_scope.Checked = False
                        error_message("VOUT (V) cannot be ""0"".")
                        Exit Sub
                    End If

                End If
            Next
        End If
    End Sub




    Private Sub clist_fs_SelectedIndexChanged(sender As Object, e As EventArgs) Handles clist_fs.SelectedIndexChanged
        data_set_list()
    End Sub

    Private Sub clist_vout_SelectedIndexChanged(sender As Object, e As EventArgs) Handles clist_vout.SelectedIndexChanged
        data_set_list()
    End Sub

    Private Sub num_vin_max_ValueChanged(sender As Object, e As EventArgs) Handles num_vin_max.ValueChanged

    End Sub

    Private Sub cbox_vin2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbox_vin2.SelectedIndexChanged
        Dim addr() As String

        power_channel_set(cbox_vin2, cbox_vin_ch2)
        If cbox_vin2.SelectedItem = no_device Then
            txt_vin_addr2.Text = ""
            vin_dev_ch2 = 0
        Else
            addr = Split(Power_addr(cbox_vin2.SelectedIndex), "::")
            txt_vin_addr2.Text = addr(1)
        End If

        cbox_vin_ch2.SelectedIndex = vin_dev_ch2
    End Sub

    Private Sub cbox_vin_ch2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbox_vin_ch2.SelectedIndexChanged
        vin_dev_ch2 = cbox_vin_ch2.SelectedIndex
    End Sub

    Private Sub cbox_vin_daq2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbox_vin_daq2.SelectedIndexChanged
        vin_daq2 = Mid(cbox_vin_daq2.SelectedItem, 3)
    End Sub

    Private Sub cbox_vout_daq2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbox_vout_daq2.SelectedIndexChanged
        Dim addr() As String
        'If PartI_first = True Then
        '    Exit Sub

        'End If

        num_iin_change2.Maximum = 0.4 * 1000
        If cbox_IIN_meter2.SelectedItem = no_device Then
            txt_IIN_addr2.Text = ""
            check_iin2.Checked = False

        Else
            check_iin2.Checked = True
            addr = Split(Meter_addr(cbox_IIN_meter2.SelectedIndex), "::")
            txt_IIN_addr2.Text = addr(1)


            'Check Relay
            Select Case Mid(cbox_IIN_meter2.SelectedItem, 1, 5)
                Case "34450"
                    '10A
                    Meter_iin_Max = 0.1
                    Meter_iin_relay(0) = 1 'H
                    Meter_iin_relay(1) = 0 'L
                    Meter_iin_range = "AUTO"

                Case "DMM40"
                    '10A
                    Meter_iin_Max = 0.4
                    Meter_iin_relay(0) = 0 'H
                    Meter_iin_relay(1) = 1 'L
                    Meter_iin_range = "MAX"

                Case "DMM65"
                    '10A
                    Meter_iin_Max = 3
                    Meter_iin_relay(0) = 0 'H
                    Meter_iin_relay(1) = 1 'L
                    Meter_iin_range = "MAX"
                    check_iin.Checked = False

            End Select

            num_iin_change2.Maximum = Meter_iin_Max * 1000

            If (num_iin_change2.Value > Meter_iin_Max * 1000) Or (num_iin_change2.Value = 0) Then
                num_iin_change2.Value = (Meter_iin_Max * 0.8 * 1000)
            End If


        End If
    End Sub

    Private Sub cbox_vout1_daq2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbox_vout1_daq2.SelectedIndexChanged
        Eff_vout_daq2 = Mid(cbox_vout1_daq2.SelectedItem, 3)
    End Sub

    Private Sub check_IOUT_ch12_CheckedChanged(sender As Object, e As EventArgs) Handles check_IOUT_ch12.CheckedChanged
        If rbtn_board_iout2.Checked = True Then
            check_IOUT_ch22.Enabled = True
            check_IOUT_ch42.Enabled = True

            If (check_IOUT_ch12.Checked = True) Then
                check_IOUT_ch22.Checked = True
            ElseIf (check_IOUT_ch12.Checked = False) Then
                check_IOUT_ch22.Checked = False

            End If

            If (check_IOUT_ch32.Checked = True) Then
                check_IOUT_ch42.Checked = True
            ElseIf (check_IOUT_ch32.Checked = False) Then
                check_IOUT_ch42.Checked = False
            End If

            If check_IOUT_ch12.Checked = True And check_IOUT_ch32.Checked = True Then
                cbox_board_buck2.SelectedIndex = 2
            End If

        End If
    End Sub

    Private Sub check_IOUT_ch32_CheckedChanged(sender As Object, e As EventArgs) Handles check_IOUT_ch32.CheckedChanged
        If rbtn_board_iout2.Checked = True Then
            check_IOUT_ch22.Enabled = False
            check_IOUT_ch42.Enabled = False

            If (check_IOUT_ch12.Checked = True) Then
                check_IOUT_ch22.Checked = True
            ElseIf (check_IOUT_ch12.Checked = False) Then
                check_IOUT_ch22.Checked = False

            End If

            If (check_IOUT_ch32.Checked = True) Then
                check_IOUT_ch42.Checked = True
            ElseIf (check_IOUT_ch32.Checked = False) Then
                check_IOUT_ch42.Checked = False
            End If

            If check_IOUT_ch12.Checked = True And check_IOUT_ch32.Checked = True Then
                cbox_board_buck2.SelectedIndex = 2
            End If
        End If
    End Sub

    Private Sub cbox_Iout_meter2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbox_Iout_meter2.SelectedIndexChanged
        Dim addr() As String
        'If PartI_first = True Then
        '    Exit Sub

        'End If
        num_iout_change2.Maximum = (0.4 * 1000)
        If cbox_Iout_meter2.SelectedItem = no_device Then
            txt_Iout_addr2.Text = ""
            check_iout2.Checked = False

        Else
            check_iout2.Checked = True
            addr = Split(Meter_addr(cbox_Iout_meter2.SelectedIndex), "::")
            txt_Iout_addr2.Text = addr(1)
            'Check Relay


            Select Case Mid(cbox_Iout_meter2.SelectedItem, 1, 5)
                Case "34450"
                    '10A
                    Meter_iout_Max = 0.1
                    Meter_iout_relay(0) = 1 'H
                    Meter_iout_relay(1) = 0 'L
                    Meter_iout_range = "AUTO"

                Case "DMM40"
                    '10A
                    Meter_iout_Max = 0.4
                    Meter_iout_relay(0) = 0 'H
                    Meter_iout_relay(1) = 1 'L
                    Meter_iout_range = "MAX"

                Case "DMM65"
                    '10A
                    Meter_iout_Max = 3
                    Meter_iout_relay(0) = 0 'H
                    Meter_iout_relay(1) = 1 'L
                    Meter_iout_range = "MAX"
                    check_iout2.Checked = False

            End Select
            num_iout_change2.Maximum = (Meter_iout_Max * 1000)

            If (num_iout_change2.Value > (Meter_iout_Max * 1000)) Or (num_iout_change2.Value = 0) Then
                num_iout_change2.Value = (Meter_iout_Max * 0.9 * 1000)
            End If
        End If
    End Sub

    Private Sub cbox_board_buck2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbox_board_buck2.SelectedIndexChanged
        If rbtn_board_iout2.Checked = True Then
            txt_board_VOUT2.Text = cbox_board_buck2.SelectedItem
            'txt_board_VIN2.Text = "VIN"
            'LX
            cbox_channel_lx2.SelectedIndex = 2 'CH3
            'cbox_channel_iout2.SelectedIndex = 3 'CH4

            If cbox_board_buck2.SelectedIndex = 0 Then
                'Buck1
                cbox_channel_vout2.SelectedIndex = 0 'CH1
                'cbox_channel_vin2.SelectedIndex = 1 'CH2

            Else
                'Buck2
                cbox_channel_vout2.SelectedIndex = 1 'CH2
                'cbox_channel_vin.SelectedIndex = 0 'CH1

            End If
        Else
            txt_board_VIN.Text = ""
            txt_board_VOUT.Text = ""
        End If
    End Sub


    Private Sub num_RL_ValueChanged(sender As Object, e As EventArgs) Handles num_RL.ValueChanged

    End Sub

    Private Sub num_vin_ValueChanged(sender As Object, e As EventArgs) Handles num_vin.ValueChanged
        If (check_vin_sense.Checked = True) And (num_vin.Value > num_vin_max.Value) Then
            ' btn_vin_add.Enabled = False
            error_message("The set value is larger than ""VIN MAX""!")
        Else
            ' btn_vin_add.Enabled = True
        End If
    End Sub


    Private Sub num_vin_start_ValueChanged(sender As Object, e As EventArgs) Handles num_vin_start.ValueChanged
        If (check_vin_sense.Checked = True) And (num_vin_start.Value > num_vin_max.Value) Then
            btn_lineR_vin.Enabled = False
            error_message("The set value is larger than ""VIN MAX""!")
        Else
            btn_lineR_vin.Enabled = True
        End If
    End Sub

    Private Sub num_vin_stop_ValueChanged(sender As Object, e As EventArgs) Handles num_vin_stop.ValueChanged
        If (check_vin_sense.Checked = True) And (num_vin_stop.Value > num_vin_max.Value) Then
            btn_lineR_vin.Enabled = False
            error_message("The set value is larger than ""VIN MAX""!")
        Else
            btn_lineR_vin.Enabled = True
        End If
    End Sub

    Private Sub num_vin_step_ValueChanged(sender As Object, e As EventArgs) Handles num_vin_step.ValueChanged
        If (check_vin_sense.Checked = True) And ((num_vin_start.Value + num_vin_step.Value) > num_vin_max.Value) Then
            btn_lineR_vin.Enabled = False
            error_message("The set value is larger than ""VIN MAX""!")
        Else
            btn_lineR_vin.Enabled = True
        End If
    End Sub
End Class
