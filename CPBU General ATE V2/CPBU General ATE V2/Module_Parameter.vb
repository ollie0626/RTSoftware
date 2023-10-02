Module Module_Parameter

    ' Meter
    Public Relay_OK As Boolean = False
    Public in_high_id, in_low_id, in_middle_id, in_io_id As Integer
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

End Module
