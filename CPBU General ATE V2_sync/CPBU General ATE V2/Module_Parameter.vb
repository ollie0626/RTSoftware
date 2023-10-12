Module Module_Parameter

    ' Meter
    Public Relay_OK As Boolean = False
    Public in_high_id, in_low_id, in_middle_id, in_io_id As Integer

    Public in_high_id2, in_low_id2, in_middle_id2, in_io_id2 As Integer


    Public out_high_id, out_low_id, out_middle_id, out_io_id As Integer
    'Public resolution_input As Double
    'Public resolution_output As Double

    Public in_high_comp, in_low_comp, in_middle_comp As Integer
    Public out_high_comp, out_low_comp, out_middle_comp As Integer

    Public in_high_resolution, in_low_resolution, in_middle_resolution As Double
    Public out_high_resolution, out_low_resolution, out_middle_resolution As Double

    Public Meter_H As Double = 9.5
    Public Meter_L As Double = 0.18
    Public DUT2_en As Boolean = False


    Public add_dut2 As String = "_DUT2"

    ' RTBB handle2
    Public hDevice2 As Integer

    ' some variable
    Public meas1 As Integer = 1
    Public meas2 As Integer = 2
    Public meas3 As Integer = 3
    Public meas4 As Integer = 4
    Public meas5 As Integer = 5
    Public meas6 As Integer = 6
    Public meas7 As Integer = 7
    Public meas8 As Integer = 8

    Public DAQ_resolution2 As String = "DEF" 'DEF=5 1/2; MIN=6 1/2; MAX=4 1/2
    Public vin_daq2 As String = ""
    Public vout_daq2 As String = ""
    Public Eff_vout_daq2 As String = ""

    ' meter handle
    Public Meter_iin_addr2 As String
    Public Meter_iin_dev2 As Integer
    Public Meter_iout_addr2 As String
    Public Meter_iout_dev2 As Integer
    Public iout_meas2 As Double
    Public iin_meas2 As Double
    Public iin_meter_change2 As Double
    Public iout_meter_change2 As Double


    ' vin handle
    Public vin_addr2 As String
    Public vin_Dev2 As Integer
    Public Power_Dev2 As Integer
    Public vin_device2 As String
    Public Vin_out2 As String
    Public vin_meas2 As Double
    Public vout_meas2 As Double
    Public monitor_vout2 As Boolean
    Public vin_dev_ch2 As Integer


    ' scope channel
    Public lx2_ch As Integer
    Public vout2_ch As Integer
    Public iout2_ch As Integer

    ' jitter 
    Public Jitter_folder2 As String

    ' Stability
    Public Beta_folder2 As String
    Public Error_folder2 As String
    Public error_pic_path2 As String

    ' Efficiency
    Public Eff_vout_meas2 As Double

    ' Eload
    Public Load_ch2 As Integer = 1
    Public Load_ch_set2(0) As Integer
    Public DCload_ch2(3) As Boolean
    Public Iout_board_EN2 As Boolean = False



    Public error_pic_num2 As Integer
    Public error_pic_col2, error_pic_row2 As Integer
    Public hyperlink_col2, hyperlink_row2 As Integer


End Module
