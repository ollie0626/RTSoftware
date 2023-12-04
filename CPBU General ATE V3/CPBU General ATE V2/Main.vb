Imports System.Runtime.InteropServices
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.IO
Imports System.Runtime.InteropServices.Marshal
Imports RTBBLibDotNet
Public Class Main



    Function Scan_Instrument() As Integer

        Dim temp() As String
        Dim i As Integer
        Dim name() As String
        Dim addr() As String
        Dim addr_temp As Integer
        Dim TCP_addr As String

        Me.Enabled = False


        status_led_error.Visible = True
        status_led_run.Visible = False

        temp = visa_scan()

        RS_Scope = False

        RS_Scope_Dev = ""
        Scope_Addr = 0
        Power_num = 0
        Meter_num = 0
        FG_num = 0
        Scope_num = 0

       
        Load_device = no_device
        DCLOAD_63600 = False
        Temp_name = no_device
        DAQ_name = no_device

        data_GPIB.Rows.Clear()

        addr_temp = 0
        TCP_addr = ""


        ReDim LOAD_63600_CCH(0)
        ReDim LOAD_63600_CCL(0)
        ReDim LOAD_63600_Watt_L(0)
        ReDim LOAD_63600_Watt_M(0)
        ReDim LOAD_63600_Watt_H(0)
        ReDim LOAD_63600_Model(0)

        If visa_count > 0 Then

            For i = 0 To temp.Length - 1  'visa_count - 1
                addr = Split(temp(i), "::")



                Select Case Mid(temp(i), 1, 3)


                    Case "TCP"



                        name = visa_name(temp(i))
                        Select Case Mid(name(1), 1, 3)
                            Case "RTE"

                                ReDim Preserve Scope_name(Scope_num)
                                ReDim Preserve Scope_IF(Scope_num)
                                Scope_name(Scope_num) = name(1)
                                Scope_IF(Scope_num) = temp(i)
                                RS_Scope_Dev = temp(i)
                                RS_Scope = True

                                Scope_num = Scope_num + 1

                                If TCP_addr <> addr(1) Then
                                    TCP_addr = addr(1)
                                    data_GPIB.Rows.Add("Scope", name(1), temp(i))
                                End If


                            Case "RTO"
                                ReDim Preserve Scope_name(Scope_num)
                                ReDim Preserve Scope_IF(Scope_num)
                                Scope_name(Scope_num) = name(1)
                                Scope_IF(Scope_num) = temp(i)

                                RS_Scope_Dev = temp(i)
                                RS_Scope = True
                                Scope_num = Scope_num + 1
                                If TCP_addr <> addr(1) Then
                                    TCP_addr = addr(1)
                                    data_GPIB.Rows.Add("Scope", name(1), temp(i))
                                End If

                            Case "DPO"
                                ReDim Preserve Scope_name(Scope_num)
                                ReDim Preserve Scope_IF(Scope_num)
                                Scope_name(Scope_num) = name(1)
                                Scope_IF(Scope_num) = temp(i)

                                Scope_Addr = addr(1)
                                Scope_num = Scope_num + 1
                                data_GPIB.Rows.Add("Scope", name(1), temp(i))



                        End Select


                    Case "USB"
                        name = visa_name(temp(i))
                        Select Case name(1)
                            Case " 2230-30-1"
                                ReDim Preserve Power_name(Power_num)
                                ReDim Preserve Power_addr(Power_num)
                                Power_name(Power_num) = name(1)
                                Power_addr(Power_num) = temp(i)
                                Power_num = Power_num + 1
                                data_GPIB.Rows.Add("Power_" & Power_num, name(1), temp(i))
                        End Select

                    Case "GPI"


                        If addr_temp = 0 Then
                            addr_temp = addr(1)
                        Else

                            If addr(1) <> addr_temp Then
                                addr_temp = addr(1)
                            Else
                                error_message("GPIB Address:" & addr(1) & " has duplicates!")
                                Exit Function

                            End If
                        End If



                        If addr(1) = 3 Then

                            Temp_name = "4350B"
                            Temp_addr = addr(1)
                            data_GPIB.Rows.Add("Chamber", Temp_name, temp(i))

                        Else
                            name = visa_name(temp(i))

                            Select Case name(1)


                                Case "34970A"

                                    DAQ_name = name(1)
                                    DAQ_addr = addr(1)

                                    data_GPIB.Rows.Add("DAQ", name(1), addr(0) & "::" & addr(1) & "::INSTR")
                                Case "N6705B"
                                    ReDim Preserve Power_name(Power_num)
                                    ReDim Preserve Power_addr(Power_num)
                                    Power_name(Power_num) = name(1)
                                    Power_addr(Power_num) = addr(0) & "::" & addr(1) & "::INSTR"
                                    Power_num = Power_num + 1
                                    '-----------------------------------------------------------------------------
                                    data_GPIB.Rows.Add("Power_" & Power_num, name(1), addr(0) & "::" & addr(1) & "::INSTR")
                                Case "E3631A"
                                    ReDim Preserve Power_name(Power_num)
                                    ReDim Preserve Power_addr(Power_num)
                                    Power_name(Power_num) = name(1)
                                    Power_addr(Power_num) = addr(0) & "::" & addr(1) & "::INSTR"
                                    Power_num = Power_num + 1
                                    data_GPIB.Rows.Add("Power_" & Power_num, name(1), addr(0) & "::" & addr(1) & "::INSTR")
                                Case "E3632A"
                                    ReDim Preserve Power_name(Power_num)
                                    ReDim Preserve Power_addr(Power_num)
                                    Power_name(Power_num) = name(1)
                                    Power_addr(Power_num) = addr(0) & "::" & addr(1) & "::INSTR"
                                    Power_num = Power_num + 1
                                    data_GPIB.Rows.Add("Power_" & Power_num, name(1), addr(0) & "::" & addr(1) & "::INSTR")


                                Case "6210-40"
                                    ReDim Preserve Power_name(Power_num)
                                    ReDim Preserve Power_addr(Power_num)
                                    Power_name(Power_num) = name(1)
                                    Power_addr(Power_num) = addr(0) & "::" & addr(1) & "::INSTR"
                                    Power_num = Power_num + 1
                                    data_GPIB.Rows.Add("Power_" & Power_num, name(1), addr(0) & "::" & addr(1) & "::INSTR")

                                Case "MODEL 2400"

                                    ReDim Preserve Power_name(Power_num)
                                    ReDim Preserve Power_addr(Power_num)
                                    Power_name(Power_num) = name(1)
                                    Power_addr(Power_num) = addr(0) & "::" & addr(1) & "::INSTR"
                                    Power_num = Power_num + 1
                                    data_GPIB.Rows.Add("Power_" & Power_num, name(1), addr(0) & "::" & addr(1) & "::INSTR")



                                Case "MODEL 2410"
                                    ReDim Preserve Power_name(Power_num)
                                    ReDim Preserve Power_addr(Power_num)
                                    Power_name(Power_num) = name(1)
                                    Power_addr(Power_num) = addr(0) & "::" & addr(1) & "::INSTR"
                                    Power_num = Power_num + 1
                                    data_GPIB.Rows.Add("Power_" & Power_num, name(1), addr(0) & "::" & addr(1) & "::INSTR")

                                Case "E36312A"
                                    ReDim Preserve Power_name(Power_num)
                                    ReDim Preserve Power_addr(Power_num)
                                    Power_name(Power_num) = name(1)
                                    Power_addr(Power_num) = addr(0) & "::" & addr(1) & "::INSTR"
                                    Power_num = Power_num + 1
                                    data_GPIB.Rows.Add("Power_" & Power_num, name(1), addr(0) & "::" & addr(1) & "::INSTR")

                                Case "6301"
                                    Load_device = name(1)
                                    Load_addr = addr(1)
                                    data_GPIB.Rows.Add("DC Load", name(1), addr(0) & "::" & addr(1) & "::INSTR")
                                Case "6304"
                                    Load_device = name(1)
                                    Load_addr = addr(1)
                                    data_GPIB.Rows.Add("DC Load", name(1), addr(0) & "::" & addr(1) & "::INSTR")
                                Case "6312A"
                                    Load_device = name(1)
                                    Load_addr = addr(1)
                                    data_GPIB.Rows.Add("DC Load", name(1), addr(0) & "::" & addr(1) & "::INSTR")
                                Case "6312"
                                    Load_device = name(1)
                                    Load_addr = addr(1)
                                    data_GPIB.Rows.Add("DC Load", name(1), addr(0) & "::" & addr(1) & "::INSTR")

                                Case "63600-1"
                                    Load_device = name(1)
                                    Load_addr = addr(1)
                                    data_GPIB.Rows.Add("DC Load", name(1), addr(0) & "::" & addr(1) & "::INSTR")

                                    load_model_check(1, 0)


                                    DCLOAD_63600 = True
                                Case "63600-2"
                                    Load_device = name(1)
                                    Load_addr = addr(1)
                                    data_GPIB.Rows.Add("DC Load", name(1), addr(0) & "::" & addr(1) & "::INSTR")

                                    load_model_check(1, 0)

                                    ReDim Preserve LOAD_63600_CCH(1)
                                    ReDim Preserve LOAD_63600_CCL(1)
                                    ReDim Preserve LOAD_63600_Watt_L(1)
                                    ReDim Preserve LOAD_63600_Watt_M(1)
                                    ReDim Preserve LOAD_63600_Watt_H(1)
                                    ReDim Preserve LOAD_63600_Model(1)

                                    load_model_check(3, 1)

                                    DCLOAD_63600 = True
                                Case "DMM4050"
                                    ReDim Preserve Meter_name(Meter_num)
                                    ReDim Preserve Meter_addr(Meter_num)
                                    Meter_name(Meter_num) = name(1)
                                    Meter_addr(Meter_num) = addr(0) & "::" & addr(1) & "::INSTR"
                                    Meter_num = Meter_num + 1
                                    data_GPIB.Rows.Add("Meter_" & Meter_num, name(1), addr(0) & "::" & addr(1) & "::INSTR")

                                Case "DMM4040"
                                    ReDim Preserve Meter_name(Meter_num)
                                    ReDim Preserve Meter_addr(Meter_num)
                                    Meter_name(Meter_num) = name(1)
                                    Meter_addr(Meter_num) = addr(0) & "::" & addr(1) & "::INSTR"
                                    Meter_num = Meter_num + 1
                                    data_GPIB.Rows.Add("Meter_" & Meter_num, name(1), addr(0) & "::" & addr(1) & "::INSTR")

                                Case "MODEL DMM6500"
                                    ReDim Preserve Meter_name(Meter_num)
                                    ReDim Preserve Meter_addr(Meter_num)
                                    Meter_name(Meter_num) = "DMM6500"
                                    Meter_addr(Meter_num) = addr(0) & "::" & addr(1) & "::INSTR"
                                    Meter_num = Meter_num + 1
                                    data_GPIB.Rows.Add("Meter_" & Meter_num, name(1), addr(0) & "::" & addr(1) & "::INSTR")

                                Case "34450A"
                                    ReDim Preserve Meter_name(Meter_num)
                                    ReDim Preserve Meter_addr(Meter_num)
                                    Meter_name(Meter_num) = name(1)
                                    Meter_addr(Meter_num) = addr(0) & "::" & addr(1) & "::INSTR"
                                    Meter_num = Meter_num + 1
                                    data_GPIB.Rows.Add("Meter_" & Meter_num, name(1), addr(0) & "::" & addr(1) & "::INSTR")



                            End Select

                            Select Case Mid(name(1), 1, 6)

                                Case "62006P"
                                    ReDim Preserve Power_name(Power_num)
                                    ReDim Preserve Power_addr(Power_num)
                                    Power_name(Power_num) = name(1)
                                    Power_addr(Power_num) = addr(0) & "::" & addr(1) & "::INSTR"
                                    Power_num = Power_num + 1
                                    data_GPIB.Rows.Add("Power_" & Power_num, name(1), addr(0) & "::" & addr(1) & "::INSTR")

                                Case "62012P"
                                    ReDim Preserve Power_name(Power_num)
                                    ReDim Preserve Power_addr(Power_num)
                                    Power_name(Power_num) = name(1)
                                    Power_addr(Power_num) = addr(0) & "::" & addr(1) & "::INSTR"
                                    Power_num = Power_num + 1
                                    data_GPIB.Rows.Add("Power_" & Power_num, name(1), addr(0) & "::" & addr(1) & "::INSTR")
                            End Select



                            Select Case Mid(name(1), 1, 3)

                                Case "DPO"
                                    ReDim Preserve Scope_name(Scope_num)
                                    ReDim Preserve Scope_IF(Scope_num)
                                    Scope_name(Scope_num) = name(1)
                                    Scope_IF(Scope_num) = temp(i)

                                    Scope_Addr = addr(1)
                                    Scope_num = Scope_num + 1
                                    data_GPIB.Rows.Add("Scope", name(1), addr(0) & "::" & addr(1) & "::INSTR")

                                Case "AFG"
                                    ReDim Preserve FG_name(FG_num)
                                    ReDim Preserve FG_Addr(FG_num)
                                    FG_name(FG_num) = name(1)
                                    FG_Addr(FG_num) = addr(0) & "::" & addr(1) & "::INSTR"
                                    FG_num = FG_num + 1
                                    '-----------------------------------------------------------------------------
                                    data_GPIB.Rows.Add("FG_" & FG_num, name(1), addr(0) & "::" & addr(1) & "::INSTR")



                                Case "332"
                                    ReDim Preserve FG_name(FG_num)
                                    ReDim Preserve FG_Addr(FG_num)
                                    FG_name(FG_num) = name(1)
                                    FG_Addr(FG_num) = addr(0) & "::" & addr(1) & "::INSTR"
                                    FG_num = FG_num + 1
                                    '-----------------------------------------------------------------------------
                                    data_GPIB.Rows.Add("FG_" & FG_num, name(1), addr(0) & "::" & addr(1) & "::INSTR")


                                Case "RTE"
                                    ReDim Preserve Scope_name(Scope_num)
                                    ReDim Preserve Scope_IF(Scope_num)
                                    Scope_name(Scope_num) = name(1)
                                    Scope_IF(Scope_num) = temp(i)

                                    RS_Scope_Dev = temp(i)
                                    RS_Scope = True
                                    Scope_num = Scope_num + 1
                                    data_GPIB.Rows.Add("Scope", name(1), temp(i))

                                Case "RTO"
                                    ReDim Preserve Scope_name(Scope_num)
                                    ReDim Preserve Scope_IF(Scope_num)
                                    Scope_name(Scope_num) = name(1)
                                    Scope_IF(Scope_num) = temp(i)

                                    RS_Scope_Dev = temp(i)
                                    RS_Scope = True
                                    Scope_num = Scope_num + 1
                                    data_GPIB.Rows.Add("Scope", name(1), temp(i))

                            End Select

                        End If

                End Select

            Next
        End If

        If data_GPIB.Rows.Count > 0 Then
            status_led_error.Visible = False
            status_led_run.Visible = True
        End If


        '--------------------------------------------
        'Power Supply



        If cbox_ven.SelectedIndex = -1 Then
            'init
            cbox_ven.Items.Clear()

            If Power_num = 0 Then

                cbox_ven.Items.Add(no_device)

            Else
                cbox_ven.Items.AddRange(Power_name)

            End If

            cbox_ven.SelectedIndex = cbox_ven.Items.Count - 1
        Else
            If device_select_same(cbox_ven, txt_ven_Addr, Power, False) = False Then
                ven_dev_ch = 0
            End If
            cbox_ven_ch.SelectedIndex = ven_dev_ch
        End If

        '--------------------------------------------
        If RS_Scope = True Then
            RS_visa(True)
            'RS_View()
        End If

        Me.Enabled = True

    End Function


    Function main_reset() As Integer
        Dim f As Form
        Dim open_num As Integer
        data_Test.Rows.Clear()
        open_num = My.Application.OpenForms.Count - 1
        While My.Application.OpenForms.Count > 1
            f = My.Application.OpenForms(My.Application.OpenForms.Count - 1)
            If f.Name <> Me.Name Then
                f.Close()
            End If
        End While
        PartI_num = 0
    End Function

    Function main_set() As Integer
        Dim i As Integer
        row = 1
        col = 1


        xlSheet.Cells.Font.Name = "Arial"
        xlSheet.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlLeft

        '//------------------------------------------------------------------------------------//
        'Main Page
        xlSheet.Cells(row, col) = "BOM Table"
        title_set()
        row = row + 1
        xlSheet.Cells(row, col) = txt_vin_cap.Text
        xlSheet.Cells(row, col + 1) = num_vin_cap.Value
        row = row + 1
        xlSheet.Cells(row, col) = txt_vout_cap.Text
        xlSheet.Cells(row, col + 1) = num_vout_cap.Value
        row = row + 1
        xlSheet.Cells(row, col) = txt_inductor.Text
        xlSheet.Cells(row, col + 1) = num_inductor.Value
        row = row + 1
        xlSheet.Cells(row, col) = txt_full_load.Text
        xlSheet.Cells(row, col + 1) = num_full_load.Value
        row = row + 1
        '------------------------------------------------------------------------------------
        xlSheet.Cells(row, col) = "3T Part"
        title_set()
        row = row + 1
        xlSheet.Cells(row, col) = check_TA_en.Checked
        xlSheet.Cells(row, col + 1) = num_delay_Temp.Value
        row = row + 1

        data_test_set(data_Temp)

      
        '------------------------------------------------------------------------------------

        xlSheet.Cells(row, col) = "I2C Initial Setting"
        title_set()
        row = row + 1

        data_test_set(data_i2c)

      
        '//------------------------------------------------------------------------------------//
        'Global Page

        '------------------------------------------------------------------------------------
        xlSheet.Cells(row, col) = "Enable Requirement Table"
        title_set()
        row = row + 1
        xlSheet.Cells(row, col) = check_en.Checked
        xlSheet.Cells(row, col + 1) = txt_EN.Text
        xlSheet.Cells(row, col + 2) = num_en_delay.Value


        row = row + 1

        xlSheet.Cells(row, col) = "Reload EN"
        xlSheet.Cells(row, col + 1) = check_EN_off.Checked

        row = row + 1


        xlSheet.Cells(row, col) = "MODE"
        xlSheet.Cells(row, col + 1) = cbox_en_mode.SelectedItem
        row = row + 1

        If cbox_en_mode.SelectedIndex = 0 Then
            'Power Supply
            xlSheet.Cells(row, col) = "ON"
            xlSheet.Cells(row, col + 1) = num_EN_ON.Value
            xlSheet.Cells(row, col + 2) = cbox_ven.SelectedItem
            xlSheet.Cells(row, col + 3) = cbox_ven_ch.SelectedItem

            row = row + 1

            xlSheet.Cells(row, col) = "OFF"
            xlSheet.Cells(row, col + 1) = num_EN_OFF.Value


            row = row + 1
        Else
            xlSheet.Cells(row, col) = "ON"
            xlSheet.Cells(row, col + 1) = txt_EN_set_ON.Text

            row = row + 1

            xlSheet.Cells(row, col) = "OFF"
            xlSheet.Cells(row, col + 1) = txt_EN_set_OFF.Text


            row = row + 1
        End If


        '//------------------------------------------------------------------------------------//
        'Global Page

        xlSheet.Cells(row, col) = "Fs Requirement Table"
        title_set()
        row = row + 1
        xlSheet.Cells(row, col) = cbox_fs_ctr.SelectedItem
        row = row + 1

        xlSheet.Cells(row, col) = data_fs.Columns(0).HeaderText
        xlSheet.Cells(row + 1, col) = data_fs.Columns(1).HeaderText

        If data_fs.Rows.Count = 0 Then
            xlSheet.Cells(row, col + 1) = num_fs_set.Value * 1000
            xlSheet.Cells(row + 1, col + 1) = ""

        Else

            For i = 0 To data_fs.Rows.Count - 1
                xlSheet.Cells(row, col + 1 + i) = data_fs.Rows(i).Cells(0).Value
                xlSheet.Cells(row + 1, col + 1 + i) = data_fs.Rows(i).Cells(1).Value
            Next
        End If
        row = row + 1
        row = row + 1
        '------------------------------------------------------------------------------------
        xlSheet.Cells(row, col) = "VOUT Requirement Table"
        title_set()
        row = row + 1

        xlSheet.Cells(row, col) = cbox_vout_ctr.SelectedItem
        row = row + 1
        xlSheet.Cells(row, col) = data_vout.Columns(0).HeaderText
        xlSheet.Cells(row + 1, col) = data_vout.Columns(1).HeaderText

        If data_vout.Rows.Count = 0 Then
            xlSheet.Cells(row, col + 1) = num_vout_set.Value
            xlSheet.Cells(row + 1, col + 1) = ""
        Else

            For i = 0 To data_vout.Rows.Count - 1
                xlSheet.Cells(row, col + 1 + i) = data_vout.Rows(i).Cells(0).Value
                xlSheet.Cells(row + 1, col + 1 + i) = data_vout.Rows(i).Cells(1).Value
            Next


        End If
        row = row + 1
        row = row + 1
      

        '//------------------------------------------------------------------------------------//

        xlSheet.Cells(row, col) = "Test"
        title_set()
        row = row + 1
        xlSheet.Cells(row, col) = "Send Email"
        xlSheet.Cells(row, col + 1) = check_email.Checked
        xlSheet.Cells(row, col + 2) = txt_email_to.Text
        xlSheet.Cells(row, col + 3) = check_report.Checked
        row = row + 1
        xlSheet.Cells(row, col) = check_excel_visible.Text
        xlSheet.Cells(row, col + 1) = check_excel_visible.Checked

        '//------------------------------------------------------------------------------------//

        xlSheet.Cells(row, col) = "Report"
        title_set()
        row = row + 1

        xlSheet.Cells(row, col) = "Row"
        xlSheet.Cells(row, col + 1) = num_test_row.Value
        xlSheet.Cells(row, col + 2) = num_row_space.Value
        xlSheet.Cells(row, col + 3) = "Column"
        xlSheet.Cells(row, col + 4) = num_test_col.Value
        xlSheet.Cells(row, col + 5) = num_col_space.Value
        row = row + 1
        xlSheet.Cells(row, col) = "Chart"
        xlSheet.Cells(row, col + 1) = num_chart_height.Value
        xlSheet.Cells(row, col + 2) = num_chart_width.Value
        xlSheet.Cells(row, col + 3) = "Picture"
        xlSheet.Cells(row, col + 4) = num_pic_height.Value
        xlSheet.Cells(row, col + 5) = num_pic_width.Value
        row = row + 1

        xlSheet.Cells(row, col) = "Data Color"
        xlSheet.Cells(row, col + 1) = num_data_color.Value
        xlSheet.Cells(row, col + 2) = "Chart Color"
        xlSheet.Cells(row, col + 3) = num_chart_color.Value

        row = row + 1


        '//------------------------------------------------------------------------------------//
        'Relay Board

        '
        xlSheet.Cells(row, col) = "Relay Board"
        title_set()
        row = row + 1

        xlSheet.Cells(row, col) = "IIN Rshunt (Ω)"
        xlSheet.Cells(row, col + 1) = num_IIN_Rshunt_L.Value
        xlSheet.Cells(row, col + 2) = num_IIN_Rshunt_H.Value

        row = row + 1
        xlSheet.Cells(row, col) = "IOUT Rshunt (Ω)"
        xlSheet.Cells(row, col + 1) = num_IOUT_Rshunt_L.Value
        xlSheet.Cells(row, col + 2) = num_IOUT_Rshunt_H.Value

        row = row + 1

        xlSheet.Cells(row, col) = "Averaging Mode :"
        xlSheet.Cells(row, col + 1) = cbox_INA226_b11_9.SelectedIndex
        xlSheet.Cells(row, col + 2) = txt_INA226_00h.Text

        row = row + 1

        xlSheet.Columns(1).AutoFit()
        FinalReleaseComObject(xlSheet)
        xlSheet = Nothing

        xlBook.Save()


    End Function

    Function main_import() As Integer
        Dim i As Integer
        Dim last_col As Integer
        Dim temp As String
        Dim import_ok As Boolean = False
        row = 1
        col = 1

        last_col = xlSheet.Range(ConvertToLetter(col) & row).CurrentRegion.Columns.Count
        '//------------------------------------------------------------------------------------//

        '"BOM Table"
        row = row + 1

        num_vin_cap.Value = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 1).Value
        row = row + 1

        num_vout_cap.Value = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 1).Value
        row = row + 1

        num_inductor.Value = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 1).Value
        row = row + 1

        num_full_load.Value = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 1).Value
        row = row + 1
        '------------------------------------------------------------------------------------
        '"3T Part"
        row = row + 1
        check_TA_en.Checked = xlSheet.Range(ConvertToLetter(col) & row).Value
        num_delay_Temp.Value = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 1).Value
        row = row + 1
        data_test_import(data_Temp, last_col)

        TA_set()
        'data_Temp.Rows.Clear()

        'For i = 0 To last_col
        '    temp = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 1 + i).Value
        '    If temp <> Nothing Then

        '        data_Temp.Rows.Add(temp)

        '    End If
        'Next


        'row = row + 1

        '------------------------------------------------------------------------------------
        ' "I2C Device"
        row = row + 1


        data_test_import(data_i2c, last_col)

        'data_i2c.Rows.Clear()

        'For i = 0 To last_col
        '    temp = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 1 + i).Value
        '    If temp <> Nothing Then

        '        data_i2c.Rows.Add(temp, xlSheet.Range(ConvertToLetter(col) & row + 1).Offset(, 1 + i).Value, xlSheet.Range(ConvertToLetter(col) & row + 2).Offset(, 1 + i).Value)

        '    End If
        'Next





        'row = row + 1
        'row = row + 1
        'row = row + 1

        '//------------------------------------------------------------------------------------//

        '"Enable Requirement Table"

        row = row + 1

        check_en.Checked = xlSheet.Range(ConvertToLetter(col) & row).Value
        txt_EN.Text = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 1).Value
        num_en_delay.Value = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 2).Value


        row = row + 1

        check_EN_off.Checked = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 1).Value

        row = row + 1

        ' xlSheet.Cells(row, col) = "MODE"
        cbox_en_mode.SelectedItem = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 1).Value
        row = row + 1


        import_ok = False
        If cbox_en_mode.SelectedIndex = 0 Then
            num_EN_ON.Value = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 1).Value
            'Power Supply
            For i = 0 To cbox_ven.Items.Count - 1
                If cbox_ven.Items(i) = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 2).Value Then
                    cbox_ven.SelectedIndex = i

                    import_ok = True
                    Exit For
                End If
            Next

            If import_ok = False Then
                cbox_ven.SelectedIndex = 0
            End If


            cbox_ven_ch.SelectedItem = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 3).Value

            'OFF
            row = row + 1
            num_EN_OFF.Value = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 1).Value



        Else

            txt_EN_set_ON.Text = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 1).Value

            'OFF
            row = row + 1
            txt_EN_set_OFF.Text = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 1).Value
        End If



        row = row + 1

        '------------------------------------------------------------------------------------
        '"Fs Requirement Table"
        row = row + 1
        cbox_fs_ctr.SelectedItem = xlSheet.Range(ConvertToLetter(col) & row).Value
        row = row + 1

        data_fs.Rows.Clear()


        If xlSheet.Range(ConvertToLetter(col) & row).Offset(1, 1).Value = Nothing Then
            num_fs_set.Value = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 1).Value / 1000
            txt_fs_set.Text = ""
        Else

            For i = 0 To last_col
                temp = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 1 + i).Value
                If temp <> Nothing Then
                    data_fs.Rows.Add(xlSheet.Range(ConvertToLetter(col) & row).Offset(, 1 + i).Value, xlSheet.Range(ConvertToLetter(col) & row).Offset(1, 1 + i).Value)
                End If
            Next
        End If


        row = row + 1
        row = row + 1

        '------------------------------------------------------------------------------------
        '"VOUT Requirement Table"

        row = row + 1

      
        cbox_vout_ctr.SelectedItem = xlSheet.Range(ConvertToLetter(col) & row).Value
        row = row + 1

        data_vout.Rows.Clear()


        If xlSheet.Range(ConvertToLetter(col) & row).Offset(1, 1).Value = Nothing Then
            num_vout_set.Value = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 1).Value
            txt_vout_set.Text = ""
        Else

            For i = 0 To last_col
                temp = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 1 + i).Value
                If temp <> Nothing Then
                    data_vout.Rows.Add(xlSheet.Range(ConvertToLetter(col) & row).Offset(, 1 + i).Value, xlSheet.Range(ConvertToLetter(col) & row).Offset(1, 1 + i).Value)
                End If
            Next
        End If


        row = row + 1
        row = row + 1

        '//------------------------------------------------------------------------------------//

        ' "Test"

        row = row + 1

        check_email.Checked = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 1).Value
        txt_email_to.Text = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 2).Value
        check_report.Checked = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 3).Value

        row = row + 1

        check_excel_visible.Checked = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 1).Value


        '//------------------------------------------------------------------------------------//
        'xlSheet.Cells(row, col) = "Report"
        'title_set()
        row = row + 1

        'xlSheet.Cells(row, col) = "Row"
        num_test_row.Value = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 1).Value
        num_row_space.Value = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 2).Value
        ' xlSheet.Cells(row, col + 3) = "Column"
        num_test_col.Value = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 4).Value
        num_col_space.Value = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 5).Value
        row = row + 1
        ' xlSheet.Cells(row, col) = "Chart"
        num_chart_height.Value = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 1).Value
        num_chart_width.Value = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 2).Value
        ' xlSheet.Cells(row, col + 3) = "Picture"
        num_pic_height.Value = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 4).Value
        num_pic_width.Value = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 5).Value
        row = row + 1

        'xlSheet.Cells(row, col) = "Data Color"
        num_data_color.Value = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 1).Value
        'xlSheet.Cells(row, col + 2) = "Chart Color"
        num_chart_color.Value = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 3).Value

        row = row + 1



        '//------------------------------------------------------------------------------------//
        'Relay Board

        '
        'xlSheet.Cells(row, col) = "Relay Board"
        'title_set()
        row = row + 1

        'xlSheet.Cells(row, col) = "IIN Rshunt (Ω)"
        num_IIN_Rshunt_L.Value = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 1).Value
        num_IIN_Rshunt_H.Value = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 2).Value

        row = row + 1
        ' xlSheet.Cells(row, col) = "IOUT Rshunt (Ω)"
        num_IOUT_Rshunt_L.Value = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 1).Value
        num_IOUT_Rshunt_H.Value = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 2).Value

        row = row + 1

        ' xlSheet.Cells(row, col) = "Averaging Mode :"
        cbox_INA226_b11_9.SelectedIndex = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 1).Value
        txt_INA226_00h.Text = xlSheet.Range(ConvertToLetter(col) & row).Offset(, 2).Value




        FinalReleaseComObject(xlSheet)
        xlSheet = Nothing

    End Function

    Function export_set() As Integer

        Dim sheet_first As Boolean = True
        Dim open_num As Integer
        Dim f As Form
        Dim i, ii As Integer

        'dlgSave.Filter = "Excel 97-2003 Worksheets|*.xls|Excel Worksheets|*.xlsx"
        dlgSave.Filter = "Excel Worksheets|*.xlsx|Excel 97-2003 Worksheets|*.xls"
        dlgSave.FilterIndex = 1
        dlgSave.RestoreDirectory = True
        dlgSave.DefaultExt = ".xlsx"
        dlgSave.FileName = "CPBU ATE_SET_" & DateTime.Now.ToString("MMdd")
        If dlgSave.ShowDialog() = System.Windows.Forms.DialogResult.OK Then

            sf_name = dlgSave.FileName

            dlgSave.Dispose()

            Delay(300)
            'Dim xlApp As New Excel.Application
            xlApp = CreateObject("Excel.Application") '?萄遣EXCEL撠情
            xlApp.DisplayAlerts = False
            xlApp.Visible = True



            xlBook = xlApp.Workbooks.Add

            xlBook.SaveAs(sf_name)

            Delay(100)


            open_num = My.Application.OpenForms.Count - 1

            If data_Test.Rows.Count > 0 Then
                For i = data_Test.Rows.Count - 1 To 0 Step -1

                    If sheet_first = True Then

                        xlSheet = xlBook.ActiveSheet
                        sheet_first = False

                    Else
                        xlSheet = xlBook.Sheets.Add
                    End If

                    xlSheet.Name = data_Test.Rows(i).Cells(1).Value

                    xlSheet.Cells.Font.Name = "Arial"
                    xlSheet.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlLeft

                    Save_set = True

                    For ii = 0 To open_num

                        f = My.Application.OpenForms(ii)
                        If f.Name = data_Test.Rows(i).Cells(1).Value Then

                            f.Show()
                            f.Hide()
                            Exit For

                        End If


                    Next





                    Save_set = False


                    'xlSheet.Columns(1).AutoFit()
                    'xlBook.Save()

                Next
            End If


            If sheet_first = True Then

                xlSheet = xlBook.ActiveSheet
                sheet_first = False

            Else
                xlSheet = xlBook.Sheets.Add
            End If

            xlSheet.Name = sheet_main
            main_set()
            'GC.Collect()
            'GC.WaitForPendingFinalizers()
            excel_close()

        End If
    End Function

    Function import_set() As Integer
        Dim i As Integer
        Dim temp() As String
        Dim check_temp As Boolean = False
        Dim test_name As String

        'dlgOpen.Filter = "Excel 97-2003 Worksheets|*.xls|Excel Worksheets|*.xlsx"
        dlgOpen.Filter = "Excel Worksheets|*.xlsx|Excel 97-2003 Worksheets|*.xls"
        dlgOpen.FilterIndex = 1
        dlgOpen.RestoreDirectory = True
        dlgOpen.DefaultExt = ".xlsx"
        dlgOpen.FileName = ""
        If dlgOpen.ShowDialog() = System.Windows.Forms.DialogResult.OK Then


            check_file_open(dlgOpen.FileName)

            dlgOpen.Dispose()

            Delay(300)

            main_reset()
            'Dim xlApp As New Excel.Application
            xlApp = CreateObject("Excel.Application") '?萄遣EXCEL撠情

            xlApp.DisplayAlerts = False
            xlApp.Visible = False



            xlBook = xlApp.Workbooks.Open(dlgOpen.FileName)
            xlBook.Activate()

            'Check Set File

            For i = 1 To xlBook.Sheets.Count
                xlSheet = xlBook.Sheets(i)
                xlSheet.Activate()
                test_name = xlSheet.Name
                FinalReleaseComObject(xlSheet)
                xlSheet = Nothing

                If test_name = sheet_main Then
                    check_temp = True
                    Exit For
                End If


            Next


            If check_temp = False Then
                error_message("This file path is not set file!")
            Else

                For i = 1 To xlBook.Sheets.Count
                    xlSheet = xlBook.Sheets(i)
                    xlSheet.Activate()
                    test_name = xlSheet.Name
                    If test_name = sheet_main Then

                        main_import()
                        GC.Collect()
                        GC.WaitForPendingFinalizers()
                    Else


                        temp = Split(test_name, "_")

                        Select Case temp(0)
                            Case "PartI"

                                cbox_test.SelectedIndex = 0

                                Open_set = True
                                data_test_now = data_Test.Rows.Count

                                If PartI_num > 0 Then
                                    Dim f As New PartI
                                    f.Name = PartI_test & "_" & PartI_num
                                    f.Show()
                                    f.Hide()
                                Else
                                    PartI.Show()
                                    PartI.Hide()
                                End If


                                data_Test.Rows.Add(True, test_name)
                                PartI_num = PartI_num + 1

                                Open_set = False




                        End Select
                        GC.Collect()
                        GC.WaitForPendingFinalizers()
                    End If
                Next

            End If




            excel_close()



        End If
    End Function


    Function fs_vout_set() As Integer
        Dim i As Integer


        If (data_fs.Rows.Count = 0) Or (cbox_fs_ctr.SelectedItem = no_device) Then
            ReDim fs_value(0)
            ReDim fs_set(0)

            fs_value(0) = num_fs_set.Value * 10 ^ 3
            fs_set(0) = ""
        Else
            ReDim fs_value(data_fs.Rows.Count - 1)
            ReDim fs_set(data_fs.Rows.Count - 1)

            For i = 0 To data_fs.Rows.Count - 1
                fs_value(i) = data_fs.Rows(i).Cells(0).Value
                fs_set(i) = data_fs.Rows(i).Cells(1).Value
            Next
        End If

        If (data_vout.Rows.Count = 0) Or (cbox_vout_ctr.SelectedItem = no_device) Then
            ReDim vout_value(0)
            ReDim vout_set(0)

            vout_value(0) = num_vout_set.Value
            vout_set(0) = ""
        Else
            ReDim vout_value(data_vout.Rows.Count - 1)
            ReDim vout_set(data_vout.Rows.Count - 1)

            For i = 0 To data_vout.Rows.Count - 1
                vout_value(i) = data_vout.Rows(i).Cells(0).Value
                vout_set(i) = data_vout.Rows(i).Cells(1).Value
            Next

        End If




    End Function

    Function excel_init() As Integer
        Dim temp As String

        dlgSave.Filter = "Excel Worksheets|*.xlsx|Excel 97-2003 Worksheets|*.xls"
        dlgSave.FilterIndex = 1
        dlgSave.RestoreDirectory = True
        dlgSave.DefaultExt = ".xlsx"

        sf_name = ""
        If num_fs_set.Value <> 0 Then
            temp = "fs=" & num_fs_set.Value & "KHz_"
        End If

        If num_vout_set.Value <> 0 Then
            temp = temp & "vout=" & num_vout_set.Value & "V_"
        End If

        dlgSave.FileName = "CPBU ATE_" & temp & DateTime.Now.ToString("MMdd")

        If dlgSave.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            sf_name = dlgSave.FileName

        Else
            'sf_name = Environment.CurrentDirectory & "\" & cbox_test.SelectedItem & "_" & DateTime.Now.ToString("MMdd") & ".xlsx" 'Now.Month & Now.Day
            run = False
            Exit Function
        End If

        Dim testFile As System.IO.FileInfo
        testFile = My.Computer.FileSystem.GetFileInfo(sf_name)
        txt_file.Text = sf_name

        folderPath = testFile.DirectoryName



        xlApp = CreateObject("Excel.Application") '?萄遣EXCEL撠情
        xlApp.DisplayAlerts = False

        xlApp.Visible = False

        xlApp.AutoRecover.Enabled = False


        xlBook = xlApp.Workbooks.Add
        xlBook.Activate()



        xlApp.Calculation = Excel.XlCalculation.xlCalculationManual

        xlBook.SaveAs(sf_name)

        excel_close()

        dlgSave.Dispose()


    End Function


    'Function RUN_stop() As Integer
    '    'Bridgeboard Stop




    '    run = False
    '    note_display = False
    '    '-------------------------------------------------
    '    'Chamber
    '    If Temp_Dev <> 0 Then
    '        temp_off()
    '        Chamber_off()
    '        ibonl(Temp_Dev, 0)

    '    End If
    '    '-------------------------------------------------
    '    'Power

    '    If vin_Dev <> 0 Then
    '        ibonl(vin_Dev, 0)
    '    End If

    '    '-------------------------------------------------
    '    'Ven
    '    If (Ven_dev <> 0) And (Ven_dev <> vin_Dev) Then
    '        ibonl(Ven_dev, 0)
    '    End If

    '    'Other Power


    '    '-------------------------------------------------
    '    'Load

    '    If Load_Dev <> 0 Then
    '        load_onoff("OFF")
    '        ibonl(Load_Dev, 0)
    '    End If

    '    '-------------------------------------------------
    '    'Meter

    '    If Meter_iin_dev <> 0 Then
    '        ibonl(Meter_iin_dev, 0)
    '    End If

    '    If Meter_iout_dev <> 0 Then
    '        ibonl(Meter_iout_dev, 0)
    '    End If



    '    '-------------------------------------------------
    '    'Scope
    '    If Scope_Dev <> 0 Then
    '        ibonl(Scope_Dev, 0)
    '    End If

    '    If RS_Scope = True Then
    '        RS_visa(False)
    '        RS_Local()
    '    End If

    '    '-------------------------------------------------
    '    'FG
    '    If FG_Dev <> 0 Then
    '        ibonl(FG_Dev, 0)
    '    End If

    '    '-------------------------------------------------
    '    'DAQ
    '    If DAQ_Dev <> 0 Then
    '        ibonl(DAQ_Dev, 0)
    '    End If

    '    btn_RUN.Visible = True
    '    btn_stop.Visible = False
    '    TabControl1.Enabled = True

    'End Function

    Function Chamber_Temp(ByVal Temp As Integer) As Integer
        Dim temp_now As Integer


        set_temp(Temp, 0)

        temp_now = read_temp()

        If Temp > temp_now Then

            While temp_now < Temp
                System.Windows.Forms.Application.DoEvents()
                If run = False Then
                    Exit While
                End If
                temp_now = read_temp()

            End While




        Else
            While temp_now > Temp
                System.Windows.Forms.Application.DoEvents()
                If run = False Then
                    Exit While
                End If

                temp_now = read_temp()


            End While

        End If


    End Function


    Private Sub Main_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        status_Version.Text = My.Application.Info.Version.Major & "." & My.Application.Info.Version.Minor & My.Application.Info.Version.Build
        Scan_Instrument()
        Check_Eagleboard()
        If data_meas.Rows.Count > 0 Then
            cbox_INA226_b11_9.SelectedIndex = cbox_INA226_b11_9.Items.Count - 1
        Else
            cbox_INA226_b11_9.SelectedIndex = 0
        End If
        cbox_test.SelectedIndex = 0
        cbox_vout_ctr.SelectedIndex = 0
        cbox_fs_ctr.SelectedIndex = 0
        cbox_test.SelectedIndex = 0
        cbox_ven.SelectedIndex = 0
        cbox_ven_ch.SelectedIndex = 0
        cbox_en_mode.SelectedIndex = 0
        First = False
    End Sub

    Private Sub btn_temp_add_Click(sender As Object, e As EventArgs) Handles btn_temp_add.Click
        Dim check_ok As Boolean = True
        If data_Temp.Rows.Count > 0 Then
            For i = 0 To data_Temp.Rows.Count - 1
                If num_Temp.Value = data_Temp.Rows(i).Cells(0).Value Then
                    check_ok = False
                    critical_message("Repeated temperature setting!")
                    Exit For
                End If
            Next
        End If
        If check_ok = True Then
            data_value_add(data_Temp, num_Temp, 0)
            TA_set()
        End If

    End Sub

    Private Sub btn_vout_add_Click(sender As Object, e As EventArgs) Handles btn_vout_add.Click
        data_vout.Rows.Add(Format(num_vout_set.Value, "#0.000"), txt_vout_set.Text)
        data_vout.CurrentCell = data_vout.Rows(data_vout.Rows.Count - 1).Cells(0)
        fs_vout_set()
    End Sub

    Private Sub btn_fs_add_Click(sender As Object, e As EventArgs) Handles btn_fs_add.Click



        data_fs.Rows.Add(num_fs_set.Value * (10 ^ 3), txt_fs_set.Text)
        data_fs.CurrentCell = data_fs.Rows(data_fs.Rows.Count - 1).Cells(0)
        fs_vout_set()
    End Sub



    Private Sub btn_i2c_add_Click(sender As Object, e As EventArgs) Handles btn_i2c_add.Click

        data_i2c.Rows.Add(hex_data(num_ID.Value, 2), hex_data(num_addr.Value, 2), hex_data(num_data.Value, 2))
        data_i2c.CurrentCell = data_i2c.Rows(data_i2c.Rows.Count - 1).Cells(0)
    End Sub

    Private Sub num_fs_set_ValueChanged(sender As Object, e As EventArgs) Handles num_fs_set.ValueChanged
        fs_vout_set()
    End Sub

    Private Sub num_vout_set_ValueChanged(sender As Object, e As EventArgs) Handles num_vout_set.ValueChanged
        fs_vout_set()
    End Sub

    Private Sub data_fs_RowsRemoved(sender As Object, e As DataGridViewRowsRemovedEventArgs) Handles data_fs.RowsRemoved
        fs_vout_set()
    End Sub

    Private Sub data_vout_RowsRemoved(sender As Object, e As DataGridViewRowsRemovedEventArgs) Handles data_vout.RowsRemoved
        fs_vout_set()
    End Sub


    Private Sub cbox_ven_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbox_ven.SelectedIndexChanged
        Dim addr() As String
       
        power_channel_set(cbox_ven, cbox_ven_ch)
        If cbox_ven.SelectedItem = no_device Then
            txt_ven_Addr.Text = ""
            ven_dev_ch = 0
        Else
            addr = Split(Power_addr(cbox_ven.SelectedIndex), "::")
            txt_ven_Addr.Text = addr(1)
        End If

        cbox_ven_ch.SelectedIndex = ven_dev_ch
   
    End Sub

    Private Sub cbox_en_mode_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbox_en_mode.SelectedIndexChanged
        If cbox_en_mode.SelectedIndex = 0 Then

            Panel_power_EN_OFF.Visible = True
            Panel_power_EN.Visible = True

            Panel_other_EN_OFF.Visible = False
            Panel_other_EN.Visible = False


        Else
            Panel_power_EN_OFF.Visible = False
            Panel_power_EN.Visible = False

            Panel_other_EN_OFF.Visible = True
            Panel_other_EN.Visible = True


        End If
    End Sub

    Private Sub btn_test_Add_Click(sender As Object, e As EventArgs) Handles btn_test_Add.Click


        Add_test = True
        Select Case cbox_test.SelectedIndex
            Case 0

                Dim f1 As New PartI


                If PartI_num > 0 Then
                    f1.Name = PartI_test & "_" & PartI_num
                Else
                    f1.Name = PartI_test
                End If

                f1.Show()

                'Case 1

                '    Dim f2 As New PartII


                '    If PartII_num > 0 Then
                '        f2.Name = PartII_test & "_" & PartI_num
                '    Else
                '        f2.Name = PartII_test
                '    End If

                '    f2.Show()
        End Select
    End Sub





    Private Sub btn_scan_Click(sender As Object, e As EventArgs) Handles btn_scan.Click
        Scan_Instrument()
        Check_Eagleboard()

    End Sub



    Private Sub data_Test_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles data_Test.CellClick

        If e.RowIndex >= 0 Then
            If data_Test.Rows(e.RowIndex).Cells(1).Selected Then
                Review_set = True
                For Each f As Form In My.Application.OpenForms
                    If f.Name = data_Test.Rows(e.RowIndex).Cells(1).Value Then
                        f.Show()
                        data_test_now = e.RowIndex
                        Exit For
                    End If
                Next
                Review_set = False
            End If

        End If




    End Sub

    Private Sub btn_clear_Click(sender As Object, e As EventArgs) Handles btn_clear.Click

        main_reset()

    End Sub


    Private Sub btn_test_import_Click(sender As Object, e As EventArgs) Handles btn_test_import.Click
        If data_GPIB.Rows.Count = 0 Then
            Scan_Instrument()
        End If

        'If (status_bridgeboad.Text = no_device) Or (txt_ID.Text = no_slave) Then
        '    Check_Eagleboard()
        'End If


        import_set()
        Delay(100)
        GC.Collect()
        GC.WaitForPendingFinalizers()

        Delay(100)
        System.Windows.Forms.Application.DoEvents()

    End Sub

    Private Sub btn_test_export_Click(sender As Object, e As EventArgs) Handles btn_test_export.Click
        export_set()
        Delay(100)
        GC.Collect()
        GC.WaitForPendingFinalizers()
        Delay(100)
        System.Windows.Forms.Application.DoEvents()
    End Sub


    Private Sub btn_RUN_Click(sender As Object, e As EventArgs) Handles btn_RUN.Click
        Dim i As Integer
  
        Dim start_test_time As Date
        Dim txt_test_time As String
        Dim title As String
        Dim style As MsgBoxStyle
        Dim response As MsgBoxResult
        Dim addr() As String
        Dim meter_temp As Integer = 0


        If data_Test.Rows.Count = 0 Then
            Exit Sub
        End If
        GC.Collect()
        GC.WaitForPendingFinalizers()

        btn_RUN.Visible = False
        btn_stop.Visible = True
        btn_pause.Visible = True
        'TabControl1.Enabled = False

        txt_test_run.Text = "Check Setting!" & vbNewLine

        ''---------------------------------------------------------------------------------
        'Test Check
        If (txt_email_to.Text = "") And ((check_email.Checked = True) Or (check_report.Checked = True)) Then
            error_message("Please enter the recipient of the email!!")
        End If

        'I2C INIT
        If (txt_ID.Text = no_device) And (data_i2c.Rows.Count > 0) Then
            error_message("Please enter the I2C Initial set value!!")
            RUN_stop()
            Exit Sub
        End If


        If Meter_iin_relay_check = True Or Meter_iout_relay_check = True Or check_multi.Checked = True Then

            If RTBB_board = False Then
                error_message("Bridgebaord is not detected!!")
                RUN_stop()
                Exit Sub
            End If
        End If

        If data_meas.Rows.Count > 0 Then
            current_monitor_init()
        End If

        ''---------------------------------------------------------------------------------
        'Instrument Check
        ''---------------------------------------------------------------------------------

        'Power

        If Power_num = 0 Then
            error_message("GPIB connection to the Power Supply is not detected!!")
            RUN_stop()
            Exit Sub

        End If

        'En
        If (check_en.Checked = True) And (cbox_en_mode.SelectedIndex = 0) Then
            addr = Split(Power_addr(cbox_ven.SelectedIndex), "::")
            ven_addr = addr(1)
            Ven_Dev = ildev(BDINDEX, addr(1), NO_SECONDARY_ADDR, TIMEOUT, EOTMODE, EOSMODE)
            ven_device = cbox_ven.SelectedItem
            Ven_out = Power_channel(ven_device, cbox_ven_ch.SelectedIndex)
        End If



        ''---------------------------------------------------------------------------------
        'DAQ
        If DAQ_addr = 0 Then
            error_message("GPIB connection to the Data Acquisition is not detected!!")
            RUN_stop()
            Exit Sub
        Else
            DAQ_Dev = ildev(BDINDEX, DAQ_addr, NO_SECONDARY_ADDR, TIMEOUT, EOTMODE, EOSMODE)
        End If





        For i = 0 To data_Test.Rows.Count - 1

            If data_Test.Rows(i).Cells(0).Value = True Then


                'DC Load

                If (Load_check(i) = True) And (Load_addr = 0) Then
                    error_message("GPIB connection to the DC Load is not detected!!")
                    RUN_stop()
                    Exit Sub
                End If

                ''---------------------------------------------------------------------------------
                'Scope

                If (Scope_check(i) = True) And (Scope_num = 0) Then
                    error_message("GPIB connection to the Scope is not detected!!")
                    RUN_stop()
                    Exit Sub

                End If

                ''---------------------------------------------------------------------------------
                'Meter
                meter_temp = 0
                If (Meter_iin_check(i) = True) Then
                    If Meter_num = meter_temp Then
                        error_message("GPIB connection to the Meter is not detected!!")
                        RUN_stop()
                        Exit Sub
                    End If
                    meter_temp = meter_temp + 1
                End If


                If (Meter_iout_check(i) = True) Then
                    If Meter_num = meter_temp Then
                        error_message("GPIB connection to the Meter is not detected!!")
                        RUN_stop()
                        Exit Sub
                    End If
                    meter_temp = meter_temp + 1
                End If

                If (Meter_icc_check(i) = True) Then
                    If Meter_num = meter_temp Then
                        error_message("GPIB connection to the Meter is not detected!!")
                        RUN_stop()
                        Exit Sub
                    End If
                    meter_temp = meter_temp + 1
                End If




                'If (Relay_iin_check(i) = True) Or (Relay_iout_check(i) = True) Then



                '    If data_meas.Rows.Count = 0 Then
                '        error_message("Current Monitor of relay Board is not detected!!")
                '        RUN_stop()
                '        Exit Sub
                '    End If

                'End If

            End If
        Next



        ''---------------------------------------------------------------------------------
        'Chamber

        If (check_TA_en.Checked = True) Then

            If check_multi.Checked = False Or (check_multi.Checked = True And rbtn_Master.Checked = True) Then
                If Temp_addr = 0 Then
                    error_message("GPIB connection to the Chamber is not detected!!")
                    RUN_stop()
                    Exit Sub
                ElseIf data_Temp.Rows.Count = 0 Then
                    error_message("Please enter the Temp test value!!")
                    RUN_stop()
                    Exit Sub
                End If

                Temp_Dev = ildev(BDINDEX, Temp_addr, NO_SECONDARY_ADDR, TIMEOUT, EOTMODE, EOSMODE)
            End If
            TA_num = data_Temp.Rows.Count - 1

        Else
            TA_num = 0
            TA_now = "25"
        End If

     
        ''Report
        excel_init()

        If sf_name = "" Then
            RUN_stop()
            Exit Sub

        End If


        GC.Collect()
        GC.WaitForPendingFinalizers()

        check_file_open(sf_name)
        run = True
        ''---------------------------------------------------------------------------------
        start_test_time = Now

        Main_test_run()
        System.Windows.Forms.Application.DoEvents()

        RUN_stop()


        GC.Collect()
        GC.WaitForPendingFinalizers()

        System.Windows.Forms.Application.DoEvents()
        ' ''---------------------------------------------------------------------------------  '------------------------------------------------------------



        txt_test_time = "Test Time:" & vbNewLine
        txt_test_time = txt_test_time & "Start= " & FormatDateTime(start_test_time, DateFormat.ShortDate) & ":" & FormatDateTime(start_test_time, DateFormat.LongTime) & vbNewLine
        txt_test_time = txt_test_time & "Stop= " & FormatDateTime(Now, DateFormat.ShortDate) & ":" & FormatDateTime(Now, DateFormat.LongTime) & vbNewLine
        txt_test_time = txt_test_time & "Total=" & DateDiff(DateInterval.Second, start_test_time, Now) & "s" & vbNewLine

        Final_update()

        GC.Collect()
        GC.WaitForPendingFinalizers()

        If (check_email.Checked = True) And (check_report.Checked = True) Then
            SendEmail(txt_email_to.Text & txt_email.Text, "CPBU General ATE Test: End test", txt_test_time, sf_name, Path.GetFileName(sf_name))
        ElseIf (check_email.Checked = True) And (check_report.Checked = False) Then
            SendEmail(txt_email_to.Text & txt_email.Text, "CPBU General ATE Test: End test", txt_test_time, "", "")
        ElseIf (check_email.Checked = False) And (check_report.Checked = True) Then
            SendEmail(txt_email_to.Text & txt_email.Text, "CPBU General ATE Test" & "(" & FormatDateTime(Now, DateFormat.GeneralDate) & ")", txt_test_time, sf_name, Path.GetFileName(sf_name))
        End If


        System.Windows.Forms.Application.DoEvents()


        style = MsgBoxStyle.Information Or MsgBoxStyle.OkOnly

        title = "Finish Test"
        response = MsgBox(txt_test_time, style, title)

        If response = MsgBoxResult.Ok Then

            GC.Collect()
            GC.WaitForPendingFinalizers()
            System.Windows.Forms.Application.DoEvents()
            Delay(100)
        End If

        ' ''---------------------------------------------------------------------------------  '------------------------------------------------------------





        Delay(100)

    End Sub

    Function release_excel() As Integer
        GC.Collect()
        GC.WaitForPendingFinalizers()
    End Function


    Function Main_test_run() As Integer
        Dim t As Integer
        Dim status As Integer
        Dim i As Integer
        Dim temp_bit(7) As Integer

        If check_multi.Checked = True Then

            'Init
            If rbtn_Slave.Checked = True Then
                For i = 0 To Master_GPIO.Length - 1
                    RTBB_GPIOSingleSetIODirection(hDevice, 32 + Master_GPIO(i), False) 'Input
                Next
                RTBB_GPIOSingleSetIODirection(hDevice, 32 + Slave_GPIO, True) 'Output
                GPIO_single_write(Slave_GPIO, 0) 'Output=0
             
            Else

                RTBB_GPIOSingleSetIODirection(hDevice, 32 + Slave_GPIO, False) 'Input
                For i = 0 To Master_GPIO.Length - 1
                    RTBB_GPIOSingleSetIODirection(hDevice, 32 + Master_GPIO(i), True) 'Output
                    GPIO_single_write(Master_GPIO(i), 0) 'Output=0
                Next

            End If
         
          
        End If


        txt_master.Text = ""
        txt_slave.Text = ""



        Delay(50)

        report_sheet_first = True

        txt_test_run.Text = txt_test_run.Text & "Test Start!" & vbNewLine

        For t = 0 To TA_num
            TA_Test_num = t


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

            If check_TA_en.Checked = True Then

                TA_now = data_Temp.Rows(t).Cells(0).Value
                txt_test_run.Text = txt_test_run.Text & "TA= " & TA_now & vbNewLine

                txt_master.Text = ""
                txt_slave.Text = ""

                If check_multi.Checked = True Then



                    If rbtn_Master.Checked = True Then
                        'Master
                        If TA_Test_num > 0 Then
                            'check slave=1?


                            status = GPIO_single_read(Slave_GPIO)

                            While status = 0
                                System.Windows.Forms.Application.DoEvents()
                                If run = False Then
                                    Exit While
                                End If
                                status = GPIO_single_read(Slave_GPIO)

                            End While

                            System.Windows.Forms.Application.DoEvents()
                            txt_slave.Text = "Test" & t & ":OK!"

                            'Master Init
                            'P2.6=0

                            For i = 0 To Master_GPIO.Length - 1
                                GPIO_single_write(Master_GPIO(i), 0) 'Output=0
                            Next

                            txt_master.Text = ""
                        End If

                        'set Chamber
                        Chamber_Temp(TA_now)
                        Delay_s(num_delay_Temp.Value)

                        'Temp ok
                        'P2.0~P2.2
                        '傳給Slave目前是第幾個溫度ready

                        temp_bit = data_set(t + 1)

                        For i = 0 To Master_GPIO.Length - 1
                            GPIO_single_write(Master_GPIO(i), temp_bit(i)) 'Output=0
                        Next
                        System.Windows.Forms.Application.DoEvents()
                        txt_master.Text = "Temp" & t & ": OK!"


                        'RUN
                        RUN_test()

                      

                        Else
                            'Slave
                        '確認Master的溫度是否達到?
                        'P2.0~P2.2


                        For i = 0 To Master_GPIO.Length - 1
                            temp_bit(i) = GPIO_single_read(Master_GPIO(i)) 'Output=0
                        Next
                        status = temp_bit(2) * 2 ^ 2 + temp_bit(1) * 2 ^ 1 + temp_bit(0)

                        While status <> (t + 1)

                            System.Windows.Forms.Application.DoEvents()
                            If run = False Then
                                Exit While
                            End If

                            For i = 0 To Master_GPIO.Length - 1
                                temp_bit(i) = GPIO_single_read(Master_GPIO(i)) 'Output=0
                            Next
                            status = temp_bit(2) * 2 ^ 2 + temp_bit(1) * 2 ^ 1 + temp_bit(0)

                        End While
                        System.Windows.Forms.Application.DoEvents()
                        txt_master.Text = "Temp" & t & ": OK!"

                        'Slave Init
                        'P2.3=0
                        GPIO_single_write(Slave_GPIO, 0)


                        txt_slave.Text = ""
                        'RUN
                        RUN_test()

                        txt_test_run.Text = txt_test_run.Text & TA_now & vbNewLine
                        'End

                        GPIO_single_write(Slave_GPIO, 1)
                        System.Windows.Forms.Application.DoEvents()
                        txt_slave.Text = "Test" & t & ": OK!"

                        End If
                    'Init
               



                    'RUN
                  


                    Else
                        'Normal


                    'set Chamber

                        Chamber_Temp(TA_now)
                        Delay_s(num_delay_Temp.Value)

                    'RUN Test

                    RUN_test()



                    End If


            Else
                'check_TA_en.Checked = False
                RUN_test()
                Exit For

            End If



        Next


        ' ''---------------------------------------------------------------------------------  '------------------------------------------------------------

        For i = 0 To Master_GPIO.Length - 1
            RTBB_GPIOSingleSetIODirection(hDevice, 32 + Master_GPIO(i), True) 'Output
            GPIO_single_write(Master_GPIO(i), 0) 'Output=0
        Next
        txt_test_run.Text = txt_test_run.Text & "-Test Finish! " & vbNewLine

    End Function

    Function RUN_test() As Integer
        Dim i, ii As Integer

        Dim open_num As Integer
        Dim f As Form
        open_num = My.Application.OpenForms.Count - 1


        For i = 0 To data_Test.Rows.Count - 1
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


            'ProgressBar1.Value = ((1 + i + data_Test.Rows.Count * t) / (data_Test.Rows.Count * (TA_num + 1))) * 100



            Test_run = True
            TestITem_run_now = False
            If data_Test.Rows(i).Cells(0).Value = True Then



                For ii = 0 To open_num

                    System.Windows.Forms.Application.DoEvents()
                    f = My.Application.OpenForms(ii)
                    If f.Name = data_Test.Rows(i).Cells(1).Value Then
                        txt_test_run.Text = txt_test_run.Text & "Test Item:" & vbNewLine
                        txt_test_run.Text = txt_test_run.Text & data_Test.Rows(i).Cells(1).Value & vbNewLine
                        f.Show()

                        System.Windows.Forms.Application.DoEvents()
                        f.Hide()

                        Exit For

                    End If


                Next




            End If
            TestITem_run_now = False
            Test_run = False



        Next

        GC.Collect()
        GC.WaitForPendingFinalizers()
        Delay(100)
    End Function


    Private Sub btn_stop_Click(sender As Object, e As EventArgs) Handles btn_stop.Click
        run = False
        System.Windows.Forms.Application.DoEvents()
    End Sub

    Private Sub num_inductor_ValueChanged(sender As Object, e As EventArgs) Handles num_inductor.ValueChanged
        'L = num_inductor.Value * (10 ^ -6) 'uH
        TA_set()
    End Sub

    Private Sub num_full_load_ValueChanged(sender As Object, e As EventArgs) Handles num_full_load.ValueChanged
        'Full_load = num_full_load.Value
        'Stability.txt_Full_load.Text = Full_load
    End Sub

    Private Sub num_vout_cap_ValueChanged(sender As Object, e As EventArgs) Handles num_vout_cap.ValueChanged
        'Cout = num_vout_cap.Value * (10 ^ -6) 'uF
        TA_set()
    End Sub

    Private Sub num_vin_cap_ValueChanged(sender As Object, e As EventArgs) Handles num_vin_cap.ValueChanged
        'Cin = num_vin_cap.Value * (10 ^ -6) 'uF
        TA_set()
    End Sub

    Private Sub check_EN_off_CheckedChanged(sender As Object, e As EventArgs) Handles check_EN_off.CheckedChanged
        If check_EN_off.Checked = True Then
            check_en.Checked = True
        End If

    End Sub


    Private Sub Link_relay_sch_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs)
        Open_file("Relay_SCH.pdf")
    End Sub

   

    Private Sub txt_ID_TextChanged(sender As Object, e As EventArgs) Handles txt_ID.TextChanged
        'If txt_ID.Text = no_slave Then
        '    btn_i2c_add.Enabled = False
        'Else
        '    btn_i2c_add.Enabled = True
        'End If
    End Sub

    Private Sub num_test_col_ValueChanged(sender As Object, e As EventArgs) Handles num_test_col.ValueChanged
        test_col = num_test_col.Value
    End Sub

    Private Sub num_test_row_ValueChanged(sender As Object, e As EventArgs) Handles num_test_row.ValueChanged
        test_row = num_test_row.Value
    End Sub
    Private Sub num_pic_height_ValueChanged(sender As Object, e As EventArgs) Handles num_pic_height.ValueChanged
        pic_height = num_pic_height.Value
    End Sub

    Private Sub num_pic_width_ValueChanged(sender As Object, e As EventArgs) Handles num_pic_width.ValueChanged
        pic_width = num_pic_width.Value
    End Sub
  
    Private Sub num_chart_width_ValueChanged(sender As Object, e As EventArgs) Handles num_chart_width.ValueChanged
        chart_width = num_chart_width.Value

    End Sub
    Private Sub num_chart_height_ValueChanged(sender As Object, e As EventArgs) Handles num_chart_height.ValueChanged
        chart_height = num_chart_height.Value
    End Sub

    Private Sub num_row_space_ValueChanged(sender As Object, e As EventArgs) Handles num_row_space.ValueChanged
        row_Space = num_row_space.Value
    End Sub

    Private Sub num_col_space_ValueChanged(sender As Object, e As EventArgs) Handles num_col_space.ValueChanged
        col_Space = num_col_space.Value
    End Sub

    Private Sub num_chart_color_ValueChanged(sender As Object, e As EventArgs) Handles num_chart_color.ValueChanged
        chart_title_color = num_chart_color.Value
    End Sub

    Private Sub num_data_color_ValueChanged(sender As Object, e As EventArgs) Handles num_data_color.ValueChanged
        data_title_color = num_data_color.Value
    End Sub

   
    Private Sub btn_pause_Click(sender As Object, e As EventArgs) Handles btn_pause.Click
        Dim title As String
        Dim style As MsgBoxStyle
        Dim response As MsgBoxResult

        pause = True
        'Power_OFF_set()
        'DCLoad_ONOFF("OFF")
        style = MsgBoxStyle.Question Or MsgBoxStyle.YesNo


        title = "Interrupt Test"
        response = MsgBox("Do you want to continue?", style, title)

        Select Case response

            Case vbNo   '中止
                run = False
                RUN_stop()

            Case vbYes '繼續
                'Power_ON_set()
                'DCLoad_ONOFF("ON")
                pause = False


        End Select

        System.Windows.Forms.Application.DoEvents()
    End Sub

  
 

    Function Final_update() As Integer
        Dim open_num As Integer
        Dim f As Form
        'xlApp = CreateObject("Excel.Application") '?萄遣EXCEL撠情
        'xlApp.DisplayAlerts = False

        ''開啟或放大檔案會變大
        'xlApp.WindowState = Excel.XlWindowState.xlMinimized

        'xlApp.Visible = False
        'xlBook = xlApp.Workbooks.Open(sf_name)
        'xlBook.Activate()


        open_num = My.Application.OpenForms.Count - 1
     
        For i = 0 To data_Test.Rows.Count - 1




            If data_Test.Rows(i).Cells(0).Value = True Then

                report_run = True
                TestITem_run_now = False

                For ii = 0 To open_num

                    f = My.Application.OpenForms(ii)
                    If f.Name = data_Test.Rows(i).Cells(1).Value Then

                        f.Show()

                        System.Windows.Forms.Application.DoEvents()
                        f.Hide()
                        GC.Collect()
                        GC.WaitForPendingFinalizers()

                        Exit For

                    End If


                Next


                report_run = False
                TestITem_run_now = False


            End If




        Next



        'excel_close(True, True)

    End Function


    Function TA_set() As Integer
        Dim i As Integer
        Dim L, Cin, Cout As Double
        Dim temp As String

        Cin = num_vin_cap.Value * (10 ^ -6)
        Cout = num_vout_cap.Value * (10 ^ -6)
        L = num_inductor.Value * (10 ^ -6)


        If (check_TA_en.Checked = True) And data_Temp.Rows.Count > 0 Then
 
                TA_num = data_Temp.Rows.Count - 1
                ReDim TA_value(TA_num)
                ReDim Cin_Value(TA_num)
                ReDim Cout_Value(TA_num)
                ReDim L_Value(TA_num)

            For i = 0 To TA_num

                TA_now = data_Temp.Rows(i).Cells(0).Value
                TA_value(i) = TA_now
                'Cin
                temp = data_Temp.Rows(i).Cells(1).Value
                If Mid(temp, 1, 1) = "-" Then

                    Cin = Cin * (1 - Val(Mid(temp, 2)) / 100)
                ElseIf Mid(temp, 1, 1) = "+" Then
                    Cin = Cin * (1 + Val(Mid(temp, 2)) / 100)
                Else
                    Cin = Cin * (1 + Val(Mid(temp, 1)) / 100)

                End If
                'Cout
                temp = data_Temp.Rows(i).Cells(2).Value

                If Mid(temp, 1, 1) = "-" Then

                    Cout = Cout * (1 - Val(Mid(temp, 2)) / 100)
                ElseIf Mid(temp, 1, 1) = "+" Then
                    Cout = Cout * (1 + Val(Mid(temp, 2)) / 100)
                Else
                    Cout = Cout * (1 + Val(Mid(temp, 1)) / 100)

                End If
                'L
                temp = data_Temp.Rows(i).Cells(3).Value

                If Mid(temp, 1, 1) = "-" Then

                    L = L * (1 - Val(Mid(temp, 2)) / 100)
                ElseIf Mid(temp, 1, 1) = "+" Then
                    L = L * (1 + Val(Mid(temp, 2)) / 100)
                Else
                    L = L * (1 + Val(Mid(temp, 1)) / 100)

                End If

                Cin_Value(i) = Cin
                Cout_Value(i) = Cout
                L_Value(i) = L
            Next
          

        Else
            TA_num = 0
            TA_now = "25"
            ReDim TA_value(TA_num)
            ReDim Cin_Value(TA_num)
            ReDim Cout_Value(TA_num)
            ReDim L_Value(TA_num)
            TA_value(TA_num) = TA_now
            Cin_Value(TA_num) = Cin
            Cout_Value(TA_num) = Cout
            L_Value(TA_num) = L




        End If

    End Function

 
  
    Private Sub data_Temp_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles data_Temp.CellEndEdit
        Dim i As Integer
        Dim TA As Double


        If e.ColumnIndex = 0 Then
            TA = data_Temp.Rows(e.RowIndex).Cells(0).Value

            For i = 0 To data_Temp.Rows.Count - 1

                If i <> e.RowIndex Then

                    If TA = data_Temp.Rows(i).Cells(0).Value Then
                        data_Temp.Rows(i).Cells(0).Value = TA_value(i)
                        critical_message("Repeated temperature setting!")
                        Exit For

                    End If
                End If



            Next
        End If


    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs)
        ''Report
        excel_init()
        GC.Collect()
        GC.WaitForPendingFinalizers()

        check_file_open(sf_name)
        run = True
        ''---------------------------------------------------------------------------------


        Main_test_run()


        RUN_stop()

        GC.Collect()
        GC.WaitForPendingFinalizers()

    End Sub

    Private Sub num_IIN_Rshunt_L_ValueChanged(sender As Object, e As EventArgs) Handles num_IIN_Rshunt_L.ValueChanged
        INA226_Iin_max_L = 0.08 / num_IIN_Rshunt_L.Value
        txt_IIN_L.Text = "0~" & INA226_Iin_max_L & "A"
    End Sub

    Private Sub num_IIN_Rshunt_H_ValueChanged(sender As Object, e As EventArgs) Handles num_IIN_Rshunt_H.ValueChanged
        INA226_Iin_max_H = 0.08 / num_IIN_Rshunt_H.Value
        txt_IIN_H.Text = INA226_Iin_max_L & "~" & INA226_Iin_max_H & "A"
    End Sub

    Private Sub num_IOUT_Rshunt_L_ValueChanged_1(sender As Object, e As EventArgs) Handles num_IOUT_Rshunt_L.ValueChanged
        INA226_Iout_max_L = 0.08 / num_IOUT_Rshunt_L.Value
        txt_IOUT_L.Text = "0~" & INA226_Iout_max_L & "A"
    End Sub

    Private Sub num_IOUT_Rshunt_H_ValueChanged_1(sender As Object, e As EventArgs) Handles num_IOUT_Rshunt_H.ValueChanged
        INA226_Iout_max_H = 0.08 / num_IOUT_Rshunt_H.Value
        txt_IOUT_H.Text = INA226_Iout_max_L & "~" & INA226_Iout_max_H & "A"
    End Sub

 
    Private Sub cbox_INA226_b11_9_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbox_INA226_b11_9.SelectedIndexChanged

        INA226_config_data = (1 * 2 ^ 14 + cbox_INA226_b11_9.SelectedIndex * 2 ^ 9 + 1 * 2 ^ 8 + Val("&H27"))
        txt_INA226_00h.Text = INA226_config_data.ToString("X4")

    End Sub

    Private Sub btn_folder_Click(sender As Object, e As EventArgs) Handles btn_folder.Click
        FolderBrowserDialog1.SelectedPath = txt_folder.Text

        If (FolderBrowserDialog1.ShowDialog() = DialogResult.OK) Then
            txt_folder.Text = FolderBrowserDialog1.SelectedPath

        End If
    End Sub

    Private Sub btn_open_excel_Click(sender As Object, e As EventArgs) Handles btn_open_excel.Click
        dlgOpen.Filter = "Excel Worksheets|*.xlsx|Excel 97-2003 Worksheets|*.xls"
        dlgOpen.FilterIndex = 1
        dlgOpen.RestoreDirectory = True
        dlgOpen.DefaultExt = ".xlsx"
        dlgOpen.FileName = ""
        If dlgOpen.ShowDialog() = System.Windows.Forms.DialogResult.OK Then

            txt_file.Text = dlgOpen.FileName

            GC.Collect()
            GC.WaitForPendingFinalizers()

            dlgOpen.Dispose()
        End If
    End Sub

   

    Private Sub btn_update_Click(sender As Object, e As EventArgs) Handles btn_update.Click

        If txt_folder.Text = "" Then
            If (FolderBrowserDialog1.ShowDialog() = DialogResult.OK) Then
                txt_folder.Text = FolderBrowserDialog1.SelectedPath
            Else
                Exit Sub
            End If
        End If

        btn_update.Enabled = False

        If txt_file.Text = "" Then

            dlgSave.Filter = "Excel Worksheets|*.xlsx|Excel 97-2003 Worksheets|*.xls"
            dlgSave.FilterIndex = 1
            dlgSave.RestoreDirectory = True
            dlgSave.DefaultExt = ".xlsx"


           

            dlgSave.FileName = "CPBU ATE_" & DateTime.Now.ToString("MMdd")

            If dlgSave.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                sf_name = dlgSave.FileName

            Else
                sf_name = txt_folder.Text & "\" & cbox_test.SelectedItem & "_" & DateTime.Now.ToString("MMdd") & ".xlsx" 'Now.Month & Now.Day

            End If

            txt_file.Text = sf_name

        End If

        If txt_sheet.Text = "" Then
            txt_sheet.Text = "Test"
        End If
        update_pic2report(1, 4)

        btn_update.Enabled = True
    End Sub

    Private Sub txt_INA226_00h_TextChanged(sender As Object, e As EventArgs) Handles txt_INA226_00h.TextChanged

        If (RTBB_board = True) Then
            INA226_config()
        End If

    End Sub

  
    Private Sub cbox_ven_ch_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbox_ven_ch.SelectedIndexChanged
        ven_dev_ch = cbox_ven_ch.SelectedIndex
    End Sub

   
    Private Sub check_TA_en_CheckedChanged(sender As Object, e As EventArgs) Handles check_TA_en.CheckedChanged
        TA_set()

        If check_TA_en.Checked = True Then
            gbox_multi.Enabled = True
        Else
            gbox_multi.Enabled = False
        End If
    End Sub


    Private Sub Button1_Click_1(sender As Object, e As EventArgs)


        RS_Scope_Dev = "TCPIP::RTE1054-17222::INSTR"
        viOpenDefaultRM(defaultRM)
        viOpen(defaultRM, RS_Scope_Dev, VI_NO_LOCK, 5000, RS_vi)

        RS_Scope = True


        Scope_RUN(True)


        viClose(RS_vi)
        viClose(defaultRM)
    End Sub

    Private Sub check_multi_CheckedChanged(sender As Object, e As EventArgs) Handles check_multi.CheckedChanged
        If check_multi.Checked = False Then
            rbtn_Master.Checked = True
            rbtn_Master.Enabled = False
            rbtn_Slave.Enabled = False
        Else
            rbtn_Master.Enabled = True
            rbtn_Slave.Enabled = True
        End If
    End Sub


    Private Sub cbox_test_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbox_test.SelectedIndexChanged
        ListBox1.Items.Clear()

        Select Case cbox_test.SelectedIndex
            Case 0
                ListBox1.Items.Add("Efficiency")
                ListBox1.Items.Add("Load Regulation")
                ListBox1.Items.Add("Line Regulation")
                ListBox1.Items.Add("Stability")
                ListBox1.Items.Add("Jitter")


            Case 1

                ListBox1.Items.Add("Soft Start and Stop")
                ListBox1.Items.Add("Enable Pattern")
                ListBox1.Items.Add("Multi Inputs")
        End Select
    End Sub

    Private Sub txt_scope_folder_TextChanged(sender As Object, e As EventArgs) Handles txt_scope_folder.TextChanged
        Scope_folder = txt_scope_folder.Text
    End Sub

End Class
