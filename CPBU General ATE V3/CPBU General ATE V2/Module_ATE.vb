Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Runtime.InteropServices.Marshal
Module Module_ATE

    Public load_num As Integer = 0
    Public Load_ch_set(0) As Integer

    Public Test_Error As Boolean = False

    Public Power As String = "Power"
    Public Meter As String = "Meter"
    Public FG As String = "FG"

    Public TA_num As Integer
    Public TA_value() As String

    Public TestITem_run_now As Boolean
    Public data_test_now As Integer = 0
    Public Fs_control As String = "Fs"
    Public Vout_control As String = "Vout"

    Public report_sheet_first As Boolean = True
    Public sheet_main As String = "Main"
    Public File_set_temp As String = Environment.CurrentDirectory & "\Temp.xlsx"
    '//-----------------------------------------------------------------------------//
    'Test Item

    Public no_device As String = "NA"
    '//-----------------------------------------------------------------------------//

    'Bom
    Public Cin_Value() As Double
    Public Cout_Value() As Double
    Public L_Value() As Double



    '//-----------------------------------------------------------------------------//
    'unit
    Public ns As Double = 10 ^ -9
    Public us As Double = 10 ^ -6
    Public ms As Double = 10 ^ -3

    '//-----------------------------------------------------------------------------//
    'Grobal setting
    Public fs_value(0) As String
    Public fs_set(0) As String
    Public vout_value(0) As String
    Public vout_set(0) As String


    '//-----------------------------------------------------------------------------//
    'Test

    Public monitor_vout As Boolean
    Public vout_title As String = "Output Voltage (V)"
    Public vin_title As String = "Input Voltage (V)"
    Public Iout_title As String = "Load Current (A)"
    Public TA_name As String = "TA (℃)"
    Public Fs_name As String = "Fsw (kHz)"
    Public Vin_name As String = "VIN (V)"
    Public Iin_name As String = "IIN (A)"
    Public Vout_name As String = "VOUT (V)"
    Public Iout_name As String = "IOUT (A)"
    Public Vcc_name As String = "VCC (V)"
    Public Icc_name As String = "ICC (A)"
    Public En_name As String = "EN (V)"

    Public PartI_num As Integer = 0
    Public PartI_test As String = "PartI"
    Public PartII_num As Integer = 1
    Public PartII_test As String = "PartII"

    Public Add_test As Boolean = False
    Public Save_set As Boolean = False
    Public Open_set As Boolean = False
    Public Review_set As Boolean = False
    Public Test_run As Boolean = False

    Public TA_Test_num As Integer = 0

    Public Vout_TA_set As String = 100
    Public Power_recorve As Boolean = False

    '//-----------------------------------------------------------------------------//
    'Report

    Public pass_value_Max, pass_value_Min As Double
    Public test_row As Integer = 20
    Public test_col As Integer = 2
    Public col_Space As Integer = 3
    Public row_Space As Integer = 5
    Public data_title_color As Integer = 24
    Public chart_title_color As Integer = 19
    Public test_fail_color As Integer = 65535

    '--------------------------------------------------------
    'Picture

    Public pic_ByteSize As Long
    Public pic_height As Integer = 16
    Public pic_width As Integer = 8
    Public pic_text As String
    Public pic_top As String


    '--------------------------------------------------------
    'Chart
    Public chart_top As String
    Public chart_height As Integer = 16
    Public chart_width As Integer = 8
    Public chart_row_start As Integer
    Public chart_row_stop As Integer

    Public copy_row_start As Integer
    Public copy_row_stop As Integer
    Public paste_row_start As Integer

    '//-----------------------------------------------------------------------------//

    'Meter
    Public Meter_iin_check() As Boolean
    Public Meter_iin_relay_check As Boolean
    Public Relay_iin_check() As Boolean
    Public Meter_iin_dev As Integer
    Public Meter_iout_dev As Integer
    Public Meter_icc_dev As Integer

    Public Meter_iin_addr As Integer
    Public Meter_iout_addr As Integer
    Public Meter_icc_addr As Integer


    Public Meter_iout_check() As Boolean
    Public Relay_iout_check() As Boolean
    Public Meter_iout_relay_check As Boolean
    Public Meter_iin_relay(1) As Integer '(0)=High (1)=Low
    Public Meter_iout_relay(1) As Integer '(0)=High (1)=Low
    Public Meter_iin_Max As Double
    Public Meter_iout_Max As Double


    Public Iout_Meter_Max As Boolean = False
    Public Iin_Meter_Max As Boolean = False
    Public iin_meter_change As Double 'A
    Public iout_meter_change As Double

    Public Meter_icc_check() As Boolean

    '--------------------------------------------------------
    'GPIO
    ' Public gpio_b5_0() As Integer = {0, 0, 0, 0, 0, 0}
    'Public gpio_num As Integer = 6
    Public gpio_b3_0() As Integer = {0, 0, 0, 0}
    Public Master_GPIO() As Integer = {0, 1, 2} 'P2.0~P2.2
    Public Slave_GPIO As Integer = 3 'P2.3


    '//-----------------------------------------------------------------------------//

    'Power
    Public vin_Dev As Integer = 0
    Public vin_device As String
    Public Vin_out As String
    Public vin_addr As Integer
    Public vin_dev_ch As Integer = 0

    Public Ven_Dev As Integer = 0
    Public ven_device As String
    Public Ven_out As String
    Public ven_addr As Integer
    Public ven_dev_ch As Integer = 0

    Public VCC_Dev As Integer = 0
    Public VCC_device As String
    Public VCC_out As String
    Public vcc_addr As Integer
    Public vcc_dev_ch As Integer = 0

    Public Power_vcc_check() As Boolean



    '//-----------------------------------------------------------------------------//

    'DAQ

    Public vin_daq As String = ""
    Public vout_daq As String = ""
    Public vcc_daq As String = ""


    '//-----------------------------------------------------------------------------//
    Public vin_meas As Double
    Public iin_meas As Double

    Public iout_meas As Double
    Public vout_meas As Double
    Public Eff_vout_meas As Double
    Public vcc_meas As Double
    Public icc_meas As Double

    Public vcc_now As Double
    Public iout_now As Double
    Public vin_now As Double
    Public fs_now As Double
    Public vout_now As Double

  

    Public total_vcc() As Double
    Public total_fs() As Double
    Public total_vout() As Double
    'Public total_iout() As Double


    Public PASS As String = "PASS"
    Public FAIL As String = "FAIL"


    Public TA_now As String = ""

    Public vout_err As Integer = 90



    '//-----------------------------------------------------------------------------//
    'Scope
    'Scope Channel 
    Public Scope_check() As Boolean
    Public vin_ch As Integer = 1
    Public vout_ch As Integer = 2
    Public iout_ch As Integer = 4
    Public lx_ch As Integer = 3

    Public RL_value As Integer
    Public Wave_num As Integer
    Public Samplerate_num As Integer


    Public wave_pc_path As String = Environment.CurrentDirectory & "\wave.CSV"
    Public xlBook_wave As Excel.Workbook
    Public xlSheet_wave As Excel.Worksheet

    '//-----------------------------------------------------------------------------//
    'DC Load 

    Public Load_check() As Boolean
    Public iout_scale_set, iout_scale_now As Integer
    Public iout_scale_unit As String
    Public DCload_ch(3) As Boolean

    'load_onoff("OFF") 會將DCLoad_ON設成False，DCLoad_Iout判斷如果DCLoad_ON=False會自動將DCLoad ON並設為True
    Public DCLoad_ON As Boolean = False

    '//-----------------------------------------------------------------------------//

    'FG
    Public FG_check As Boolean = False

    '//-----------------------------------------------------------------------------//
    'Status
    Public note_display As Boolean
    Public note_delay As String = "Time(s):"
    Public note_count As String = "Count:"
    Public note_run As String = "Status:"
    Public note_value As Integer
    Public note_string As String = ""
    '//-----------------------------------------------------------------------------//
    'Report Note

    Public Note_EN_value As String = no_device

    Dim report_row As Integer = 1
    Dim report_col As Integer = 2

    Public report_run As Boolean = False

    Public folderPath As String




    Function update_pic2report(ByVal pic_start_col As Integer, ByVal pic_start_row As Integer) As Integer
        Dim pic_format As String = ".PNG"
        Dim num_temp As Integer
        Dim update_row, update_col As Integer
        Dim temp() As String
        Dim height_temp As Double
        Dim width_temp As Double
        Dim i As Integer
        Dim sheet_ok As Boolean = False


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

            For i = 1 To xlApp.Worksheets.Count
                If xlApp.Worksheets(i).Name = Main.txt_sheet.Text Then
                    xlSheet = xlBook.Sheets(Main.txt_sheet.Text)
                    sheet_ok = True
                    Exit For
                End If
            Next

            If sheet_ok = False Then
                xlSheet = xlBook.Sheets.Add
                xlSheet.Name = Main.txt_sheet.Text
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

            excel_close()


            FinalReleaseComObject(xlSheet)
            Note.Close()
        End If





    End Function

    Function data_test_import(ByVal data As Object, ByVal last_col As Integer) As Integer
        Dim temp As String


        data.Rows.Clear()

        xlrange = xlSheet.Range(ConvertToLetter(col) & row)

        For i = 0 To last_col

            temp = xlrange.Offset(, 1 + i).Value

            If (temp <> Nothing) And (temp <> " ") And (temp <> "") Then
                data.Rows.Add(temp)
                For ii = 1 To data.Columns.Count - 1
                    temp = xlrange.Offset(ii, 1 + i).Value
                    data.rows(data.rows.count - 1).cells(ii).value = temp

                Next

            End If
        Next

        For ii = 0 To data.Columns.Count - 1
            row = row + 1
        Next

        FinalReleaseComObject(xlrange)

        'xlrange = Nothing
    End Function


    Function data_test_set(ByVal data As Object) As Integer
        Dim i, ii As Integer

        For ii = 0 To data.Columns.Count - 1

            xlSheet.Cells(row + ii, col) = data.Columns(ii).HeaderText

        Next

        For i = 0 To data.Rows.Count - 1

            For ii = 0 To data.Columns.Count - 1
                xlSheet.Cells(row + ii, col + 1 + i) = data.Rows(i).Cells(ii).Value
            Next
        Next

        For ii = 0 To data.Columns.Count - 1
            row = row + 1
        Next
    End Function

    Function sheet_init(ByVal sheet_name As String) As Integer

        If report_sheet_first = True Then

            xlSheet = xlBook.ActiveSheet



            report_sheet_first = False

        Else
            xlSheet = xlBook.Sheets.Add
        End If

        xlSheet.Name = sheet_name

        xlSheet.Cells.Font.Name = "Arial"

    End Function


    Function Calculate_iout(ByVal data_iout As Object) As Double()
        Dim i, ii As Integer
        Dim iout, iout_start, iout_stop, iout_step As Double
        Dim iout_temp() As Double
        Dim iout_num As Integer = 0
        Dim iout_same As Boolean

        For i = 0 To data_iout.Rows.Count - 1
            iout_start = data_iout.Rows(i).Cells(0).Value
            iout_stop = data_iout.Rows(i).Cells(1).Value
            iout_step = data_iout.Rows(i).Cells(2).Value


            If iout_step = 0 Then

                iout_step = 1

            End If



            For iout = iout_start To iout_stop Step iout_step
                iout = Math.Round(iout, 9)


                If i = 0 Then
                    ReDim Preserve iout_temp(iout_num)
                    iout_temp(iout_num) = iout
                    iout_num = iout_num + 1

                Else

                    iout_same = False
                    For ii = 0 To iout_temp.Length - 1

                        If iout_temp(ii) = iout Then
                            iout_same = True
                        End If


                    Next

                    If iout_same = False Then

                        ReDim Preserve iout_temp(iout_num)
                        iout_temp(iout_num) = iout
                        iout_num = iout_num + 1

                    End If

                End If

                If data_iout.Rows(i).Cells(2).Value = 0 Or (iout_start = iout_stop) Then
                    Exit For
                End If

            Next
        Next

        '------------------------------------------------------------
        '由小排到大

        Array.Sort(iout_temp)

        Return iout_temp

        '------------------------------------------------------------
    End Function


    Function meas_type(ByVal cbox_type As Object, ByVal cbox_meas As Object) As Integer
        Dim Ampl() As String = {"AMPlitude", "PK2Pk", "RMS", "HIGH", "LOW", "MAXimum", "MINImum", "CRMs", "MEAN", "CMEan", "POVershoot", "NOVershoot"}
        Dim Time() As String = {"RISe", "FALL", "PWIdth", "NWIdth", "PERIod", "FREQuency", "PDUty", "NDUty", "DELay"}
        Dim i As Integer

        cbox_meas.Items.Clear()

        If cbox_type.SelectedIndex = 0 Then
            ''Ampl
            For i = 0 To Ampl.Length - 1
                cbox_meas.Items.Add(Ampl(i))
            Next
        Else
            For i = 0 To Time.Length - 1
                cbox_meas.Items.Add(Time(i))
            Next

        End If


        cbox_meas.Selectedindex = 0

    End Function

    Function excel_open() As Integer

        xlApp = CreateObject("Excel.Application") '?萄遣EXCEL撠情
        xlApp.DisplayAlerts = False

        If Main.check_excel_visible.Checked = True Then

            xlApp.WindowState = Excel.XlWindowState.xlMaximized
            xlApp.Visible = True
        Else
            xlApp.WindowState = Excel.XlWindowState.xlMinimized

            xlApp.Visible = False
        End If

        xlBook = xlApp.Workbooks.Open(sf_name)

    End Function

    Function excel_close_temp() As Integer
        xlBook.Save()
        xlBook.Close(True)

        FinalReleaseComObject(xlBook)
        xlApp.Quit() '結束EXCEL對象



        xlBook = Nothing
        xlSheet = Nothing

        FinalReleaseComObject(xlApp)
        xlApp = Nothing
    End Function


    Function excel_close() As Integer


        'Delay(100)

        excel_close_temp()

        GC.Collect()
        GC.WaitForPendingFinalizers()

        ' Kill(file_name)
        'Delay(100)
    End Function


    Function check_file_open(ByVal file_name As String) As Integer
        Dim IsError = True
        Dim ReadIO As IO.FileStream
        Dim Mcu_Save_CSV_File As String = file_name
        While IsError
            Try
                ReadIO = IO.File.OpenRead(Mcu_Save_CSV_File)       '/開啟檔案等待有無錯誤訊息
                IsError = False                                 '/預設沒有錯誤狀態
            Catch Have_Read As System.IO.IOException
                '/讀取Have_Read.GetType.ToString目前狀態
                '/當出現System.IO.FileNotFoundException為檔案路徑不存在
                '/當出現System.IO.IOException為檔案被占用
                If Have_Read.GetType.ToString = "System.IO.FileNotFoundException" Then
                    IO.File.Create(Mcu_Save_CSV_File).Dispose()
                    FileClose()
                Else
                    MsgBox("目前檔案被開啟，請關閉檔案後再試", 0)

                End If

                'MsgBox(Have_Read.GetType.ToString, 0)
            End Try
        End While
        ReadIO.Close()  '/關閉開啟的檔案
    End Function



    'Function Paste_scope_pic(ByVal paste_txt As String, ByVal paste_pic_col As Integer, ByVal paste_pic_row As Integer, ByVal pic_path As String) As Integer



    '    ' ''------------------------------------------------------------
    '    ' ''Update Picture-2

    '    pic_top = ConvertToLetter(paste_pic_col) & paste_pic_row

    '    xlrange = xlSheet.Range(pic_top)
    '    xlrange.Interior.ColorIndex = 45 'Orange
    '    xlrange.Value2 = paste_txt

    '    xlrange = xlSheet.Range(pic_top & ":" & ConvertToLetter(paste_pic_col + pic_width - 1) & paste_pic_row)
    '    xlrange.MergeCells = True
    '    FinalReleaseComObject(xlrange)

    '    pic_ByteSize = Hardcopy("PNG", pic_path)





    '    'xlrange = Nothing


    'End Function



    'Function update_pic(ByVal paste_pic_col As Integer, ByVal paste_pic_row As Integer) As Integer
    '    Dim height_temp As Double
    '    Dim width_temp As Double
    '    'Update picture


    '    pic_top = ConvertToLetter(paste_pic_col) & paste_pic_row

    '    xlrange = xlSheet.Range(pic_top & ":" & ConvertToLetter(paste_pic_col) & (paste_pic_row + pic_height - 1))
    '    height_temp = xlrange.Height
    '    xlrange = xlSheet.Range(pic_top & ":" & ConvertToLetter(paste_pic_col + pic_width - 1) & paste_pic_row)
    '    width_temp = xlrange.Width


    '    pic_ByteSize = Hardcopy("PNG", pic_path)



    '    If pic_ByteSize > 0 Then
    '        paste_picture(pic_path, pic_top, width_temp, height_temp)
    '        Delay(100)
    '    End If

    '    FinalReleaseComObject(xlrange)

    'End Function


    Function update_pic(ByVal paste_pic_col As Integer, ByVal paste_pic_row As Integer, ByVal pic_path As String) As Integer
        Dim height_temp As Double
        Dim width_temp As Double






        'Update picture


        pic_top = ConvertToLetter(paste_pic_col) & paste_pic_row
        xlrange = xlSheet.Range(pic_top & ":" & ConvertToLetter(paste_pic_col) & (paste_pic_row + pic_height - 1))
        height_temp = xlrange.Height
        xlrange = xlSheet.Range(pic_top & ":" & ConvertToLetter(paste_pic_col + pic_width - 1) & paste_pic_row)
        width_temp = xlrange.Width


        pic_ByteSize = Hardcopy("PNG", pic_path)



        If (pic_ByteSize > 0) Then
            paste_picture(pic_path, pic_top, width_temp, height_temp)
            Delay(100)
        End If


    End Function



    Function Power_EN(ByVal EN_ON As Boolean) As Integer
        Dim temp() As String
        Dim ii As Integer
        Dim ID, addr, data As Byte
        Dim control_bits As Integer = 4
        Dim test() As String


        If Main.check_en.Checked = True Then

            If Main.cbox_en_mode.SelectedIndex = 0 Then
                If Power_num = 0 Then
                    error_message("GPIB connection to the Power Supply is not detected!!")
                    RUN_stop()
                    Exit Function
                Else

                    If Ven_Dev = 0 Then
                        temp = Split(Power_addr(Main.cbox_ven.SelectedIndex), "::")
                        Ven_Dev = ildev(BDINDEX, temp(1), NO_SECONDARY_ADDR, TIMEOUT, EOTMODE, EOSMODE)
                        ven_device = Main.cbox_ven.SelectedItem
                        Ven_out = Power_channel(vin_device, Main.cbox_ven_ch.SelectedIndex)
                    End If



                    Power_Dev = Ven_Dev

                    If EN_ON = True Then
                        power_volt(ven_device, Ven_out, Main.num_EN_ON.Value)
                    Else
                        power_volt(ven_device, Ven_out, Main.num_EN_OFF.Value)
                    End If




                    If Main.num_en_delay.Value <> 0 Then
                        Delay(Main.num_en_delay.Value)
                    End If
                    power_on_off(ven_device, Ven_out, "ON")

                    Power_Dev = vin_Dev


                End If

            Else


                If EN_ON = True Then
                    temp = Split(Main.txt_EN_set_ON.Text, ",")
                Else
                    temp = Split(Main.txt_EN_set_OFF.Text, ",")
                End If

                If Main.num_en_delay.Value <> 0 Then
                    Delay(Main.num_en_delay.Value)
                End If



                If Main.status_bridgeboad.Text <> no_device Then

                    Select Case Main.cbox_en_mode.SelectedItem

                        Case "I2C"

                            For ii = 0 To temp.Length - 1

                                ' ID_Addr[Data],
                                test = Split(temp(ii), "_")
                                ID = Val("&H" & test(0))
                                addr = Val("&H" & Mid(test(1), 1, 2))
                                data = Val("&H" & Mid(test(1), 4, 2))

                                reg_write(ID, addr, data, device_sel)
                            Next

                        Case "GPIO"

                            For ii = 0 To temp.Length - 1

                                gpio_b3_0(ii) = temp(ii)
                            Next

                            GPIO_out(control_bits, gpio_b3_0, device_sel)

                    End Select



                Else
                    error_message("No Bridge Board detected!!")
                    RUN_stop()
                    Exit Function
                End If

            End If



        End If

    End Function



    'Function excel_data_copy(ByVal sheet_copy As String, ByVal copy_col As Integer, ByVal Sheet_paste As String, ByVal paste_col As Integer) As Integer
    '    'Copy Range: target_col&Start_row ~ target_col&final_row Ex: C2:C80
    '    'Paste: target_col&target_row Ex:B1
    '    Dim xlSheet_source As Excel.Worksheet
    '    xlSheet_source = xlBook.Worksheets(sheet_copy)
    '    xlSheet_source.Activate()
    '    xlSheet_source.Range(ConvertToLetter(copy_col) & copy_row_start & ":" & ConvertToLetter(copy_col) & copy_row_stop).Copy()



    '    xlSheet = xlBook.Worksheets(Sheet_paste)
    '    xlSheet.Activate()
    '    xlSheet.Range(ConvertToLetter(paste_col) & paste_row_start).Select()
    '    xlSheet.Paste()

    '    FinalReleaseComObject(xlSheet_source)
    'End Function

    Function vin_power_sense(ByVal vin_instrument As String, ByVal volt_step As Double, ByVal volt_max As Double, ByVal volt_sense As Double) As Integer
        Dim temp As Double
        Dim value As Double
        Dim new_volt As Double
        Dim i As Integer
        Dim power_now As Double

        Power_Dev = vin_Dev
        If vin_instrument = " 2230-30-1" Then


            power_now = Power2230_read(Vin_out, "VOLT")
        Else

            power_now = power_read(vin_instrument, Vin_out, "VOLT")
        End If

        For i = 0 To 4


            value = 0
            value = DAQ_read(vin_daq)

            If value > 0 Then
                Exit For
            End If


        Next



        temp = value - volt_sense
        new_volt = power_now - temp

        While (value > (volt_sense + volt_step) Or value < (volt_sense - volt_step))
            System.Windows.Forms.Application.DoEvents()

            If run = False Then
                Exit While
            End If

            If value > (volt_sense + volt_step) Then
                new_volt = new_volt - volt_step

                'If new_volt < volt - range_max Then
                If new_volt <= 0 Then
                    Exit While
                End If
            Else
                new_volt = new_volt + volt_step

                'If new_volt > volt + range_max Then
                If new_volt >= volt_max Then
                    Exit While
                End If
            End If

            If vin_instrument = " 2230-30-1" Then


                Power2230_set(Vin_out, new_volt)
            Else

                power_volt(vin_instrument, Vin_out, new_volt)
            End If

            value = DAQ_read(vin_daq)

            check_vout()

        End While





    End Function

    Function Remaining_time(ByVal stop_time As String) As Integer
        Dim stopDate As Date

        stopDate = CDate(stop_time)
        hour = DateDiff(DateInterval.Hour, Now, stopDate)
        minute = DateDiff(DateInterval.Minute, Now, stopDate) - (hour * 60)
        second = DateDiff(DateInterval.Second, Now, stopDate) - (hour * 60 * 60) - minute * 60
        'txt_remaining_time.Text = hour & ":" & minute & ":" & second & " (hh:mm:ss)"

        Return (hour * 60 * 60 + minute * 60 + second)

    End Function

    Function Delay_s(ByVal second As Integer) As Integer

        Information.information_run("Remaining time", note_delay)

        run_time = DateTime.Now.AddSeconds(second + 1)

        run_second = Remaining_time(run_time)

        note_value = run_second

        'Delay second
        While run_second > 0

            System.Windows.Forms.Application.DoEvents()

            If run = False Then
                Exit While
            End If


            run_second = Remaining_time(run_time)

            note_value = run_second

        End While

        note_display = False

    End Function



    Function report_test_info() As Integer
        Dim i, ii As Integer



        '------------------------------------------------------------------------------------
        'Initial Page

        report_title("Test Time", report_col, report_row, 2 * (TA_num + 2), 1, 44)

        For i = 0 To (TA_num + 1)
            For ii = 2 To 6
                report_title("", report_col + 2 * i, ii, 2, 1, 2)
            Next

        Next
        xlrange = xlSheet.Range(ConvertToLetter(report_col) & report_row)
        xlrange.Offset(1, 0).Value = "Temp. (℃):"
        xlrange.Offset(2, 0).Value = "Start time:"
        xlrange.Offset(3, 0).Value = "Stop time:"
        xlrange.Offset(4, 0).Value = "Total times(s):"
        xlrange.Offset(5, 0).Value = "Total points:"
        FinalReleaseComObject(xlrange)
        report_Group(report_col, 1, 2 * (TA_num + 2), 6)

    End Function



    Function report_test_update(ByVal start_test_time As Date, ByVal point As String) As Integer
        xlSheet.Activate()
        xlrange = xlSheet.Range(ConvertToLetter(report_col + 2 * (1 + TA_Test_num)) & report_row)
        xlrange.Offset(1, 0).Value = TA_now

        xlrange.Offset(2, 0).Value = FormatDateTime(start_test_time, DateFormat.LongTime)

        xlrange.Offset(3, 0).Value = FormatDateTime(Now, DateFormat.LongTime)

        xlrange.Offset(4, 0).Value = DateDiff(DateInterval.Second, start_test_time, Now)

        xlrange.Offset(5, 0).Value = point

        FinalReleaseComObject(xlrange)

    End Function




    Function monitor_count(ByVal num_counts As Integer, ByVal scope_stop As Boolean, ByVal test_item As String) As Integer

        Dim count_before As Integer
        Dim count_temp As Integer
        Dim measure_time As Date
        Dim error_monitor As Boolean = False
        Dim second_max As Integer = 30 '30 sec
        Dim trigger_error As Boolean = False

        Information.information_run("Monitor Count", note_count)



        Scope_measure_reset()


        'Information.information_run("Monitor Count", note_count)


        count_temp = Scope_measure_count(1)

        count_before = count_temp

        note_value = count_temp

        While count_temp <= num_counts


            System.Windows.Forms.Application.DoEvents()

            If run = False Then
                Exit While
            End If

            count_temp = Scope_measure_count(1)
            '----------------------------------------
            note_value = count_temp



            '----------------------------------------
            If count_temp = count_before Then


                If error_monitor = False Then
                    measure_time = Now
                    error_monitor = True
                End If

                If (DateDiff(DateInterval.Second, measure_time, Now) >= 2) And (trigger_error = False) Then
                    Select Case test_item

                        Case "Part I"
                            '交換trigger
                            If PartI.rbtn_vin_trigger.Checked = True Then
                                Trigger_auto_level(lx_ch, "R")
                            Else
                                Trigger_set(lx_ch, "R", vin_now / PartI.num_vin_trigger.Value)
                            End If

                    End Select

                    trigger_error = True

                    count_temp = Scope_measure_count(1)

                    measure_time = Now

                End If

                '-----------------------------------------------------------------
                '遇到異常重開2次還是異常直接中斷測試
                If DateDiff(DateInterval.Second, measure_time, Now) >= second_max Then

                    If monitor_vout = True Then
                        check_vout()
                    End If
                    critical_message("Detecting LX count timeout!")

                    Exit While

                End If
                '-----------------------------------------------------------------
            Else

                error_monitor = False
                count_before = count_temp
            End If



        End While


        If scope_stop = True Then
            Scope_RUN(False)


        End If

        If trigger_error = True Then
            '回設定值
            Select Case test_item

                Case "Part I"
                    If PartI.rbtn_vin_trigger.Checked = True Then
                        Trigger_set(lx_ch, "R", vin_now / PartI.num_vin_trigger.Value)
                    Else
                        Trigger_auto_level(lx_ch, "R")

                    End If
            End Select


        End If



        note_display = False



    End Function

    Function Iin_Meter_initial(ByVal check_iin As Object, ByVal cbox_IIN_meter As Object, ByVal cbox_IIN_relay As Object) As Integer

        '保持在高檔位
        If (check_iin.Checked = True) Then
            GPIO_single_write(Mid(cbox_IIN_relay.SelectedItem, 4, 1), Meter_iin_relay(0))
            Delay(20)
        End If

        'If Meter_iin_range = "4e-1" Then
        If Meter_iin_range <> "MAX" Then
            Meter_iin_range = "MAX"
        End If
        Iin_Meter_Max = True

        If Meter_iin_dev <> 0 Then
            meter_config(cbox_IIN_meter.SelectedItem, Meter_iin_dev, Meter_iin_range)
        End If





    End Function



    Function Iout_Meter_initial(ByVal check_iout As Object, ByVal cbox_Iout_meter As Object, ByVal cbox_Iout_relay As Object) As Integer


        If (check_iout.Checked = True) Then

            GPIO_single_write(Mid(cbox_Iout_relay.SelectedItem, 4, 1), Meter_iout_relay(0))
            Delay(20)
        End If

        'If Meter_iout_range = "4e-1" Then
        If Meter_iout_range <> "MAX" Then
            Meter_iout_range = "MAX"
        End If

        If Meter_iout_dev <> 0 Then
            meter_config(cbox_Iout_meter.SelectedItem, Meter_iout_dev, Meter_iout_range)

        End If

        Iout_Meter_Max = True




    End Function

    Function GPIB_reset(ByVal dev As Integer) As Integer


        ts = "*RST" & Convert.ToChar(10)

        ilwrt(dev, ts, CInt(Len(ts)))

        Delay(1000)
    End Function


    Function Iin_meter_set(ByVal check_iin As Object, ByVal cbox_IIN_meter As Object, ByVal cbox_iin_relay As Object) As Double

        If check_iin.Checked = True Then

            '以DC Load抽載的電流值來切檔位


            If (iout_now < iin_meter_change) And (Iin_Meter_Max = True) Then

                '切小檔位

                'DCLoad_ONOFF("OFF")
                Power_OFF_set()
                GPIO_single_write(Mid(cbox_iin_relay.SelectedItem, 4, 1), Meter_iin_relay(1))

                Delay(100)

                If Meter_iin_range = "MAX" Then
                    ' Meter_iin_range = Meter_iin_low
                    Meter_iin_range = "4e-1"
                End If

                Iin_Meter_Max = False

                meter_config(cbox_IIN_meter.SelectedItem, Meter_iin_dev, Meter_iin_range)




            ElseIf (iout_now >= iin_meter_change) And (Iin_Meter_Max = False) Then
                '切大檔位

                'DCLoad_ONOFF("OFF")
                Power_OFF_set()
                GPIO_single_write(Mid(cbox_iin_relay.SelectedItem, 4, 1), Meter_iin_relay(0))
                Delay(100)

                If Meter_iin_range <> "MAX" Then

                    Meter_iin_range = "MAX"
                End If

                Iin_Meter_Max = True
                meter_config(cbox_IIN_meter.SelectedItem, Meter_iin_dev, Meter_iin_range)


            End If


            If DCLoad_ON = False Then
                Power_ON_set()
                'DCLoad_ONOFF("ON")
                Delay(100)
            End If

        End If

        'meter_value = meter_average(Meter_iin_dev, 1, Meter_iin_range)

        'Return meter_value



    End Function

    Function Iout_meter_set(ByVal check_iout As Object, ByVal cbox_Iout_meter As Object, ByVal cbox_IOUT_relay As Object) As Double


        'meter_value = meter_average(Meter_dev(num), 1, Meter_range(num))

        '以DC Load抽載的電流值來切檔位

        If check_iout.Checked = True Then
            If (iout_now < iout_meter_change) And (Iout_Meter_Max = True) Then


                DCLoad_ONOFF("OFF")

                GPIO_single_write(Mid(cbox_IOUT_relay.SelectedItem, 4, 1), Meter_iout_relay(1))

                Delay(100)

                If Meter_iout_range = "MAX" Then
                    Meter_iout_range = "1e-4" '"4e-1"
                End If

                Iout_Meter_Max = False

                meter_config(cbox_Iout_meter.SelectedItem, Meter_iout_dev, Meter_iout_range)

                Delay(100)


            ElseIf (iout_now >= iout_meter_change) And (Iout_Meter_Max = False) Then

                DCLoad_ONOFF("OFF")
                GPIO_single_write(Mid(cbox_IOUT_relay.SelectedItem, 4, 1), Meter_iout_relay(0))


                Delay(100)


                If Meter_iout_range <> "MAX" Then
                    Meter_iout_range = "MAX"
                End If
                Iout_Meter_Max = True
                meter_config(cbox_Iout_meter.SelectedItem, Meter_iout_dev, Meter_iout_range)

            End If


            If DCLoad_ON = False Then
                DCLoad_ONOFF("ON")
                Delay(100)
            End If
        End If





        'meter_value = meter_average(Meter_iout_dev, 1, Meter_iout_range)

        'Return meter_value

    End Function

    Function DCLoad_ONOFF(ByVal onoff As String) As Integer
        Dim i As Integer




        For i = 0 To Load_ch_set.Length - 1

            Load_ch = Load_ch_set(i)

            If Iout_board_EN = True Then


                If onoff = "ON" Then
                    If Iout_Meter_Max = True Then

                        'Iout >80mA-> CH2,CH4

                        load_onoff("OFF")
                        Load_ch = Load_ch_set(i) + 1
                        load_onoff("ON")
                    Else
                        'Iout <80mA-> CH2,CH4

                        Load_ch = Load_ch_set(i) + 1
                        load_onoff("OFF")

                        Load_ch = Load_ch_set(i)
                        load_onoff("ON")




                    End If
                Else


                    load_onoff("OFF")
                    Load_ch = Load_ch_set(i) + 1
                    load_onoff("OFF")
                End If





            Else

                load_onoff(onoff)
            End If
        Next
    End Function
    'Function DCLoad_ONOFF(ByVal onoff As String) As Integer
    '    Dim ch_temp As Integer
    '    If Iout_board_EN = True Then

    '        If (DCload_ch(0) = True) Then
    '            Load_ch = 1
    '        Else
    '            Load_ch = 3
    '        End If
    '        If onoff = "ON" Then
    '            If Iout_Meter_Max = True Then

    '                'Iout >80mA-> CH2,CH4
    '                load_onoff("OFF")
    '                Load_ch = Load_ch + 1
    '                load_onoff("ON")
    '            Else
    '                'Iout <80mA-> CH2,CH4
    '                ch_temp = Load_ch
    '                Load_ch = Load_ch + 1
    '                load_onoff("OFF")

    '                Load_ch = ch_temp
    '                load_onoff("ON")




    '            End If
    '        Else
    '            load_onoff("OFF")
    '            Load_ch = Load_ch + 1
    '            load_onoff("OFF")
    '        End If


    '    Else

    '        load_onoff(onoff)
    '    End If

    'End Function

    Function DCLoad_check_range() As Integer
        Dim iout_temp As Double
        Dim i As Integer


        'DC Load Mode Change
        If DCLOAD_63600 = True Then

            'Check Watt

            For i = 0 To LOAD_63600_CCL.Length - 1

                DCLoad_CCL = LOAD_63600_CCL(i)
                'Check Low
                DCLoad_CCH = LOAD_63600_CCH(i)

                iout_temp = LOAD_63600_Watt_L(i) / vout_now

                If iout_temp < DCLoad_CCL Then
                    LOAD_63600_CCL(i) = iout_temp
                End If

                'check High

                iout_temp = LOAD_63600_Watt_M(i) / vout_now

                If iout_temp < DCLoad_CCH Then
                    LOAD_63600_CCH(i) = iout_temp
                End If

            Next



        End If

    End Function

    Function DCLoad_Iout(ByVal iout_now As Double, ByVal vout_check As Boolean) As Integer
        Dim change_mode As Boolean = False
        Dim module_sel As Integer
        Dim i As Integer

        For i = 0 To Load_ch_set.Length - 1

            Load_ch = Load_ch_set(i)

            'DC Load Mode Change
            If DCLOAD_63600 = True Then

                module_sel = Int((Load_ch_set(0) - 1) / 2)

                DCLoad_CCH = LOAD_63600_CCH(module_sel)
                DCLoad_CCL = LOAD_63600_CCL(module_sel)

                If iout_now > DCLoad_CCH And Load_range <> Load_range_H Then
                    Load_range = Load_range_H
                    change_mode = True '   load_init(Load_range)
                ElseIf iout_now <= DCLoad_CCH And iout_now > DCLoad_CCL And Load_range <> Load_range_M Then
                    Load_range = Load_range_M
                    change_mode = True '   load_init(Load_range)
                ElseIf iout_now <= DCLoad_CCL And Load_range <> Load_range_L Then
                    Load_range = Load_range_L
                    change_mode = True '   load_init(Load_range)
                End If

            Else


                If iout_now > DCLoad_CCH And Load_range <> Load_range_H Then
                    Load_range = Load_range_H
                    change_mode = True '   load_init(Load_range)
                ElseIf iout_now <= DCLoad_CCH And Load_range <> Load_range_L Then
                    Load_range = Load_range_L
                    change_mode = True '   load_init(Load_range)
                End If

            End If




            'Current monitor change

            If Iout_board_EN = True Then



                'CH1,CH2
                If iout_now > INA226_Iout_max_L Then
                    If Iout_Meter_Max <> True Then
                        'L->H
                        Iout_Meter_Max = True
                        load_onoff("OFF")
                        load_init(Load_range_L)
                        load_set(0)
                    End If

                    Load_ch = Load_ch_set(i) + 1

                Else
                    'H -> L

                    If Iout_Meter_Max <> False Then

                        Iout_Meter_Max = False

                        Load_ch = Load_ch_set(i) + 1
                        load_onoff("OFF")
                        load_init(Load_range_L)
                        load_set(0)
                        Load_ch = Load_ch_set(i)
                    End If

                End If

            End If


            If change_mode = True Then
                DCLoad_ONOFF("OFF")
                Load_ch = Load_ch_set(i)
                load_init(Load_range)
            End If





            load_set(iout_now / Load_ch_set.Length)


            If DCLoad_ON = False Then
                DCLoad_ONOFF("ON")

                Delay(100)
            End If
        Next

       

        If vout_check = True Then
            check_vout()
        End If



    End Function

    'Function DCLoad_Iout(ByVal iout_now As Double, ByVal vout_check As Boolean) As Integer
    '    Dim change_mode As Boolean = False
    '    Dim ch_temp As Integer
    '    Dim module_sel As Integer


    '    'DC Load Mode Change
    '    If DCLOAD_63600 = True Then

    '        module_sel = Int((Load_ch - 1) / 2)

    '        DCLoad_CCH = LOAD_63600_CCH(module_sel)
    '        DCLoad_CCL = LOAD_63600_CCL(module_sel)

    '        If iout_now > DCLoad_CCH And Load_range <> Load_range_H Then
    '            Load_range = Load_range_H
    '            change_mode = True '   load_init(Load_range)
    '        ElseIf iout_now <= DCLoad_CCH And iout_now > DCLoad_CCL And Load_range <> Load_range_M Then
    '            Load_range = Load_range_M
    '            change_mode = True '   load_init(Load_range)
    '        ElseIf iout_now <= DCLoad_CCL And Load_range <> Load_range_L Then
    '            Load_range = Load_range_L
    '            change_mode = True '   load_init(Load_range)
    '        End If

    '    Else


    '        If iout_now > DCLoad_CCH And Load_range <> Load_range_H Then
    '            Load_range = Load_range_H
    '            change_mode = True '   load_init(Load_range)
    '        ElseIf iout_now <= DCLoad_CCH And Load_range <> Load_range_L Then
    '            Load_range = Load_range_L
    '            change_mode = True '   load_init(Load_range)
    '        End If

    '    End If




    '    'Current monitor change

    '    If Iout_board_EN = True Then

    '        If (DCload_ch(0) = True) Then
    '            Load_ch = 1
    '        Else
    '            Load_ch = 3
    '        End If

    '        'CH1,CH2
    '        If iout_now > INA226_Iout_max_L Then
    '            If Iout_Meter_Max <> True Then
    '                'L->H
    '                Iout_Meter_Max = True
    '                load_onoff("OFF")
    '                load_init(Load_range_L)
    '                load_set(0)
    '            End If

    '            Load_ch = Load_ch + 1

    '        Else
    '            'H -> L

    '            If Iout_Meter_Max <> False Then

    '                Iout_Meter_Max = False
    '                ch_temp = Load_ch
    '                Load_ch = Load_ch + 1
    '                load_onoff("OFF")
    '                load_init(Load_range_L)
    '                load_set(0)
    '                Load_ch = ch_temp
    '            End If

    '        End If

    '    End If


    '    If change_mode = True Then
    '        DCLoad_ONOFF("OFF")
    '        load_init(Load_range)
    '    End If





    '    load_set(iout_now)


    '    If DCLoad_ON = False Then
    '        DCLoad_ONOFF("ON")

    '        Delay(100)
    '    End If

    '    If vout_check = True Then
    '        check_vout()
    '    End If



    'End Function



    Function check_iout_scale() As Integer


        'Check iout scale
        '高於offset=Scale/4 跳階，低於原來才降階
        'a. IOUT <= 600mA, Scale = 200mA 
        'b. 600mA<IOUT <= 3A, Scale = 1A
        'c. 3A <IOUT<= 6A, Scale = 2A
        'd. 6<IOUT<=15,scale=5A
        'e.15<IOUT<=60,scale=20A
        iout_scale_unit = "V"

        If iout_now <= 0.6 Then
            iout_scale_set = 200
            iout_scale_unit = "mV"
        ElseIf (iout_now > 0.6) And (iout_now <= 3) Then
            iout_scale_set = 1
        ElseIf (iout_now > 3) And (iout_now <= 6) Then
            iout_scale_set = 2
        ElseIf (iout_now > 6) And (iout_now <= 15) Then
            iout_scale_set = 5
        Else
            iout_scale_set = 20
        End If




        If iout_scale_set <> iout_scale_now Then


            '----------------------------------------------------------

            'IOUT

            CHx_scale(iout_ch, iout_scale_set, iout_scale_unit) 'a. IOUT < 600mA, Scale = 200mA, b. 600mA<IOUT < 3A, Sacle = 1A,c. 3A <IOUT< 6A, Scale = 2A




        End If

    End Function

    Function GPIB_refresh() As Integer
        '-------------------------------------------------
        'Chamber
        If Temp_Dev <> 0 Then

            ibonl(Temp_Dev, 0)
            Temp_Dev = ildev(BDINDEX, Temp_addr, NO_SECONDARY_ADDR, TIMEOUT, EOTMODE, EOSMODE)

        End If
        '-------------------------------------------------
        'DAQ
        If DAQ_Dev <> 0 Then
            ibonl(DAQ_Dev, 0)
            DAQ_Dev = ildev(BDINDEX, DAQ_addr, NO_SECONDARY_ADDR, TIMEOUT, EOTMODE, EOSMODE)

        End If
        'DC Load
        If Load_Dev <> 0 Then

            ibonl(Load_Dev, 0)
            Load_Dev = ildev(BDINDEX, Load_addr, NO_SECONDARY_ADDR, TIMEOUT, EOTMODE, EOSMODE)
        End If

        'Meter
        If Meter_iin_dev <> 0 Then

            ibonl(Meter_iin_dev, 0)
            Meter_iin_dev = ildev(BDINDEX, Meter_iin_addr, NO_SECONDARY_ADDR, TIMEOUT, EOTMODE, EOSMODE)
        End If


        If Meter_iout_dev <> 0 Then
            ibonl(Meter_iout_dev, 0)
            Meter_iout_dev = ildev(BDINDEX, Meter_iout_addr, NO_SECONDARY_ADDR, TIMEOUT, EOTMODE, EOSMODE)
        End If


        If Meter_icc_dev <> 0 Then
            ibonl(Meter_icc_dev, 0)
            Meter_icc_dev = ildev(BDINDEX, Meter_icc_addr, NO_SECONDARY_ADDR, TIMEOUT, EOTMODE, EOSMODE)
        End If


        '--------------------------------------------------
        'Power 
        If VCC_Dev <> 0 Then
            ibonl(VCC_Dev, 0)
            VCC_Dev = ildev(BDINDEX, vcc_addr, NO_SECONDARY_ADDR, TIMEOUT, EOTMODE, EOSMODE)

        End If

        If vin_Dev <> 0 Then
            ibonl(vin_Dev, 0)
            vin_Dev = ildev(BDINDEX, vin_addr, NO_SECONDARY_ADDR, TIMEOUT, EOTMODE, EOSMODE)

        End If

        If Ven_Dev <> 0 Then
            ibonl(Ven_Dev, 0)
            Ven_Dev = ildev(BDINDEX, ven_addr, NO_SECONDARY_ADDR, TIMEOUT, EOTMODE, EOSMODE)

        End If

        '-------------------------------------------------
        'Scope
        If Scope_Dev <> 0 Then
            ibonl(Scope_Dev, 0)
            Scope_Dev = ibdev32(BDINDEX, Scope_Addr, NO_SECONDARY_ADDR, TIMEOUT, EOTMODE, EOSMODE)
        End If



        '-------------------------------------------------
    End Function

    Function RUN_stop() As Integer

        pause = False
        note_display = False
        '-------------------------------------------------
        'Chamber
        If Temp_Dev <> 0 Then
            temp_off()
            Chamber_off()
            ibonl(Temp_Dev, 0)
            Temp_Dev = 0
        End If
        '-------------------------------------------------
        'DAQ
        If DAQ_Dev <> 0 Then
            ibonl(DAQ_Dev, 0)
            DAQ_Dev = 0

        End If
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

        '-------------------------------------------------




        Main.btn_RUN.Visible = True
        Main.btn_stop.Visible = False
        Main.btn_pause.Visible = False
        Main.TabControl1.Enabled = True
    End Function



    Function check_vout() As Integer

        If Val(TA_now) > Val(Vout_TA_set) Then
            If Power_recorve = True Then
                Power_OFF_set()
                Power_ON_set()
            End If
            Exit Function
        End If
        vout_meas = DAQ_read(vout_daq)
        If vout_meas < (vout_now * vout_err / 100) Then
            critical_message("Please confirm the output voltage!")
        End If
    End Function

    Function Power_OFF_set() As Integer
        DCLoad_ONOFF("OFF")
        Power_EN(False)

        'Vin

        Power_Dev = vin_Dev
        power_on_off(vin_device, Vin_out, "OFF")


    End Function

    Function Power_ON_set() As Integer

        ' ''----------------------------------------------------------------------------------
        'Vin
        Power_Dev = vin_Dev
        power_on_off(vin_device, Vin_out, "ON")
        Delay(100)
        ' ''----------------------------------------------------------------------------------
        'Power Enable
        'Enable
        Power_EN(True)
        Delay(100)
        ''---------------------------------------------------------------------------------
        'I2C Init
        If data_i2c_p.Rows.Count > 0 Then
            For i = 0 To Main.data_i2c.Rows.Count - 1
                System.Windows.Forms.Application.DoEvents()
                If run = False Then
                    Exit For
                End If
                reg_write(Val("&H" & data_i2c_p.Rows(i).Cells(0).Value),
                          Val("&H" & data_i2c_p.Rows(i).Cells(1).Value),
                          Val("&H" & data_i2c_p.Rows(i).Cells(2).Value))
            Next
        End If

        ''---------------------------------------------------------------------------------
        'Fs Set
        If cbox_fs_ctr_p.SelectedItem <> no_device Then
            Grobal_Control(Fs_control, fs_now,
                           data_fs_p, data_vout_p,
                           cbox_fs_ctr_p, cbox_vout_ctr_p)
        End If

        'Vout Set
        If cbox_vout_ctr_p.SelectedItem <> no_device Then
            Grobal_Control(Vout_control, vout_now,
                           data_fs_p, data_vout_p,
                           cbox_fs_ctr_p, cbox_vout_ctr_p)
        End If

        If Main.check_EN_off.Checked = True Then
            Power_EN(False)
            Power_EN(True)
        End If

        DCLoad_ONOFF("ON")
    End Function

    Function error_message(ByVal msg As String) As String
        Dim title As String
        Dim style As MsgBoxStyle
        Dim response As MsgBoxResult

        If (run = True) And (Main.check_email.Checked = True) Then
            SendEmail(Main.txt_email_to.Text, "CPBU General ATE Test: Error!!!", msg, "", "")
        End If


        style = MsgBoxStyle.Critical Or MsgBoxStyle.OkOnly

        title = "Error Message"
        response = MsgBox(msg, style, title)
        If response = vbOK Then
            If run = True Then
                Power_OFF_set()
                run = False
            End If

        End If



    End Function

    Function critical_message(ByVal msg As String) As String
        Dim title As String
        Dim style As MsgBoxStyle
        Dim response As MsgBoxResult
        pause = True


        If (run = True) And (Main.check_email.Checked = True) Then
            SendEmail(Main.txt_email_to.Text, "CPBU General ATE Test: Error!!!", msg, "", "")
        End If



        style = MsgBoxStyle.Critical Or MsgBoxStyle.RetryCancel


        title = "Error Message"
        response = MsgBox(msg, style, title)

        Select Case response

            Case vbCancel   '中止
                'Power_OFF_set()
                DCLoad_ONOFF("OFF")
                run = False


            Case vbRetry '重試
                'Power_OFF_set()
                Power_ON_set()
                pause = False


        End Select



    End Function



    Function check_message(ByVal msg As String) As Boolean
        Dim title As String
        Dim style As MsgBoxStyle
        Dim response As MsgBoxResult


        style = MsgBoxStyle.Question Or MsgBoxStyle.YesNo
        title = "Check Message"
        response = MsgBox(msg, style, title)
        If response = vbOK Then
            Return True
        Else
            Return False
        End If


    End Function


    Function error_capture(ByVal source_num As Integer, ByVal edge As String, ByVal value1 As Double, ByVal Addition As Boolean, ByVal value2 As Double, ByVal delay_ms As Integer) As Integer
        Dim status_temp As String
        Dim error_count As Integer = 0
        Dim error_num As Integer = 1

        error_num = 0


        'error pic

        'single trigger，讀取scope是否已經已經STOP
        '詢問100次沒有就再減(或加)1/10，直到5/10就停止
        '詢問之間可以設定delay
        If RS_Scope = True Then

            'RS_Local()
            RS_View(True)

        End If

        If Addition = True Then
            'Trigger的電壓以level1增加level2的1/10開始
            Trigger_set(source_num, edge, value1 + (value2 * (error_num / 10)))
        Else
            'Trigger的電壓以level1減少level2的1/10開始
            Trigger_set(source_num, edge, value1 - (value2 * (error_num / 10)))
        End If



        RUN_set("SEQuence")
        Scope_RUN(True)

        status_temp = Scope_status()

        error_count = 0

        While status_temp = "Running"

            If delay_ms > 0 Then
                Delay(delay_ms)
            End If


            System.Windows.Forms.Application.DoEvents()

            If run = False Then
                Exit While
            End If

            status_temp = Scope_status()

            If error_count >= 200 Then
                If error_num = 10 Then
                    Exit While
                End If
                error_num = error_num + 1

                If Addition = True Then
                    Trigger_set(source_num, edge, value1 + (value2 * (error_num / 10)))
                Else
                    Trigger_set(source_num, edge, value1 - (value2 * (error_num / 10)))
                End If
                error_count = 0
            Else
                error_count = error_count + 1
            End If
            Delay(10)

        End While

        If (status_temp = "Running") And (RS_Scope = True) Then

            Scope_RUN(False)

        End If

        'RUN_set("RUNSTop")

    End Function

    'Function Grobal_Control_check(ByVal type As String) As Double
    '    Dim temp() As String
    '    Dim test() As String
    '    Dim i, ii As Integer
    '    Dim ID, addr, data As Byte
    '    Dim control_bits As Integer = 4

    '    Dim cbox_ctr As Object
    '    Dim txt_set As Object



    '    If type = Fs_control Then
    '        txt_set = Main.txt_fs_set
    '        cbox_ctr = Main.cbox_fs_ctr
    '    Else

    '        txt_set = Main.txt_vout_set
    '        cbox_ctr = Main.cbox_vout_ctr

    '    End If



    '    temp = Split(txt_set.Text, ",")
    '    Select Case cbox_ctr.SelectedItem

    '        Case "I2C"
    '            For ii = 0 To temp.Length - 1
    '                ' ID_Addr[Data],
    '                test = Split(temp(ii), "_")
    '                ID = Val("&H" & test(0))
    '                addr = Val("&H" & Mid(test(1), 1, 2))
    '                data = Val("&H" & Mid(test(1), 4, 2))
    '                reg_write(ID, addr, data)
    '            Next



    '        Case "GPIO"

    '            For ii = 0 To temp.Length - 1

    '                gpio_b3_0(ii) = temp(ii)
    '            Next

    '            GPIO_out(control_bits, gpio_b3_0)



    '    End Select



    'End Function


    Function Grobal_Control(ByVal type As String, ByVal test_value As Double,
                            ByVal data_fs As DataGridView,
                            ByVal data_vout As DataGridView,
                            ByVal cbox_fs_ctr As ComboBox,
                            ByVal cbox_vout_ctr As ComboBox
                            ) As Double


        Dim temp() As String
        Dim test() As String
        Dim i, ii, n As Integer
        'Dim ID, addr, data As Byte
        Dim ID, addr As Byte
        Dim data() As Byte
        Dim data_num As Integer
        Dim data_temp As String

        Dim control_bits As Integer = 4

        Dim data_set As Object
        Dim cbox_ctr As Object

        If type = Fs_control Then
            data_set = data_fs
            cbox_ctr = cbox_fs_ctr
        Else

            data_set = data_vout
            cbox_ctr = cbox_vout_ctr

        End If

        For i = 0 To data_set.Rows.Count - 1

            If (test_value = data_set.Rows(i).Cells(0).Value) Then

                temp = Split(data_set.Rows(i).Cells(1).Value, ",")
                Select Case cbox_ctr.SelectedItem

                    Case "I2C"
                        For ii = 0 To temp.Length - 1
                            ' ID_Addr[Data],

                            test = Split(temp(ii), "_")
                            ID = Val("&H" & test(0))
                            'test(1) -> xx[xx,xx]

                            addr = Val("&H" & Mid(test(1), 1, 2))
                            data_num = test(1).Length - 4

                            data_temp = Mid(test(1), 4, data_num)

                            test = Split(data_temp, ":")

                            ReDim data(test.Length - 1)
                            For n = 0 To test.Length - 1
                                data(n) = Val("&H" & test(n))
                            Next

                            reg_write_multi(ID, addr, data, device_sel)

                            'data = Val("&H" & Mid(test(1), 4, 2))
                            'reg_write(ID, addr, data)


                        Next

                        Exit For

                    Case "GPIO"

                        For ii = 0 To temp.Length - 1

                            gpio_b3_0(ii) = temp(ii)
                        Next

                        GPIO_out(control_bits, gpio_b3_0, device_sel)

                        Exit For

                End Select
            End If

        Next




    End Function

    Function device_select_same(ByVal cbox_dev As Object, ByVal txt_addr As Object, ByVal mode As String, ByVal add_NA As Boolean) As Boolean
        Dim dev_temp As String
        Dim addr_temp As Integer
        Dim addr() As String
        Dim dev_num As Integer
        Dim dev_name() As String
        Dim dev_addr() As String
        Dim same As Boolean = False

        Select Case mode


            Case Power

                dev_num = Power_num
                dev_name = Power_name
                dev_addr = Power_addr

            Case Meter


                dev_num = Meter_num
                dev_name = Meter_name
                dev_addr = Meter_addr


            Case FG

                dev_num = FG_num
                dev_name = FG_name
                dev_addr = FG_Addr


        End Select


        dev_temp = cbox_dev.SelectedItem
        addr_temp = Val(txt_addr.Text)

        If dev_temp <> no_device Then
            For i = 0 To dev_name.Length - 1

                '先選相同名稱，相同addr
                If (dev_name(i) = dev_temp) Then
                    addr = Split(dev_addr(i), "::")
                    If addr_temp = addr(1) Then
                        Exit Function

                    End If
                End If

            Next
        End If




        cbox_dev.Items.Clear()


        If dev_num = 0 Then

            cbox_dev.Items.Add(no_device)
            cbox_dev.SelectedIndex = 0
        Else
            cbox_dev.Items.AddRange(dev_name)
            If add_NA = True Then
                cbox_dev.Items.Add(no_device)
            End If

            If dev_temp = no_device Then
                cbox_dev.SelectedItem = no_device
            Else
                For i = 0 To dev_name.Length - 1

                    '先選相同名稱，相同addr
                    If (dev_name(i) = dev_temp) Then
                        addr = Split(dev_addr(i), "::")
                        If addr_temp = addr(1) Then
                            cbox_dev.SelectedIndex = i
                            same = True
                            Exit For
                        End If
                    End If

                Next

                If same = False Then
                    For i = 0 To dev_name.Length - 1

                        '先選相同名稱，相同addr
                        If (dev_name(i) = dev_temp) Then
                            cbox_dev.SelectedIndex = i
                            Exit For
                        End If
                    Next
                End If

            End If


            If cbox_dev.SelectedIndex = -1 Then
                cbox_dev.SelectedIndex = 0
            End If

        End If



        Return same
    End Function

    Function data_value_add(ByVal data As Object, ByVal num As Object, ByVal double_num As Integer) As Integer
        Dim data_row As Integer
        Dim temp As Double

        If data.Rows.Count = 0 Then
            data_row = 0
        Else
            data_row = data.SelectedCells(0).RowIndex + 1
        End If

        temp = num.value

        data.Rows.Insert(data_row, temp.ToString("F" & double_num))
        data.CurrentCell = data.Rows(data_row).Cells(0)
    End Function




    Function Iout_meter_average(ByVal cbox_Iout_meter As Object, ByVal average As Integer, ByVal range As String) As Double
        Dim i As Integer
        Dim temp, total As Double
        Dim error_num As Integer = 5


        For i = 1 To average
            System.Windows.Forms.Application.DoEvents()

            If run = False Then
                Exit For
            End If



            temp = meter_meas(cbox_Iout_meter.SelectedItem, Meter_iout_dev, range, Meter_iin_low)


            '1.讀取過大的值

            While temp > (10 ^ 10)
                System.Windows.Forms.Application.DoEvents()

                If run = False Then
                    Exit While
                End If

                GPIB_reset(Meter_iout_dev)

                temp = meter_meas(cbox_Iout_meter.SelectedItem, Meter_iout_dev, range, Meter_iin_low)
                Delay(10)


            End While




            '2. 沒有回應或Measured IOUT < 0.9 x IOUT condition or Measured IOUT > 1.1 x IOUT condition

            If (temp = 0) Or (temp > (1.1 * iout_now)) Or (temp < (0.9 * iout_now)) Then
                i = i - 1

                If error_num = 0 Then
                    average = i
                    Exit For
                Else
                    error_num = error_num - 1
                End If

            End If


            If i = 1 Then
                total = temp
            Else
                total = total + temp
            End If

        Next
        If average > 0 Then
            total = total / average
        Else
            total = 0
        End If

        Return total
    End Function


    Function report_title(ByVal title As String, ByVal start_col As Integer, ByVal start_row As Integer, ByVal col_num As Integer, ByVal row_num As Integer, ByVal color As Integer) As Integer
        Dim top As String

        xlSheet.Activate()
        top = ConvertToLetter(start_col) & start_row & ":" & ConvertToLetter(start_col + col_num - 1) & (start_row + row_num - 1)
        xlSheet.Cells(start_row, start_col) = title

        xlrange = xlSheet.Range(top)
        xlrange.MergeCells = True

        'xlrange.Interior.ColorIndex = color
        xlrange.HorizontalAlignment = Excel.Constants.xlCenter
        xlrange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous
        xlSheet.Columns(start_col).AutoFit()

        FinalReleaseComObject(xlrange)
        'xlrange = Nothing
    End Function

    Function report_Group(ByVal start_col As Integer, ByVal start_row As Integer, ByVal col_num As Integer, ByVal row_num As Integer) As Integer
        Dim top As String

        xlSheet.Activate()
        top = ConvertToLetter(start_col) & start_row & ":" & ConvertToLetter(start_col + col_num - 1) & (start_row + row_num - 1)
        xlrange = xlSheet.Range(top)
        xlrange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous

        xlrange.Borders(Excel.XlBordersIndex.xlEdgeTop).Weight = Excel.XlBorderWeight.xlMedium
        xlrange.Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = Excel.XlBorderWeight.xlMedium
        xlrange.Borders(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlMedium
        xlrange.Borders(Excel.XlBordersIndex.xlEdgeLeft).Weight = Excel.XlBorderWeight.xlMedium

        xlrange.Borders(Excel.XlBordersIndex.xlInsideHorizontal).Weight = Excel.XlBorderWeight.xlThin
        xlrange.Borders(Excel.XlBordersIndex.xlInsideVertical).Weight = Excel.XlBorderWeight.xlThin

        xlrange.HorizontalAlignment = Excel.Constants.xlCenter

        FinalReleaseComObject(xlrange)
        'xlrange = Nothing
    End Function

    Function title_set() As Integer
        xlSheet.Activate()
        xlrange = xlSheet.Range(ConvertToLetter(col) & row)

        xlrange.Font.Bold = True
        xlrange.Interior.Color = 65535
        FinalReleaseComObject(xlrange)
        'xlrange = Nothing
    End Function

    Function pic_init(ByVal title_text As String, ByVal start_col As Integer, ByVal start_row As Integer, ByVal pic_num As Integer) As Integer
        Dim top As String
        xlSheet.Activate()
        'Update Title

        pic_row = start_row + 1
        pic_col = start_col + pic_width - 1
        report_title(title_text, start_col, start_row, pic_width, 1, chart_title_color)
        top = ConvertToLetter(start_col) & start_row & ":" & ConvertToLetter(pic_col) & (pic_row)

        xlrange = xlSheet.Range(top)
        xlrange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous

        xlrange.Borders(Excel.XlBordersIndex.xlEdgeTop).Weight = Excel.XlBorderWeight.xlMedium
        xlrange.Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = Excel.XlBorderWeight.xlMedium
        xlrange.Borders(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlMedium
        xlrange.Borders(Excel.XlBordersIndex.xlEdgeLeft).Weight = Excel.XlBorderWeight.xlMedium
        '-------------------------------------------------------------------------------------------------------------------
        For i = 0 To pic_num - 1

            pic_row = start_row + i * pic_height + 1
            top = ConvertToLetter(start_col) & (pic_row) & ":" & ConvertToLetter(pic_col) & (pic_row + pic_height - 1)
            xlrange = xlSheet.Range(top)
            xlrange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous

            xlrange.Borders(Excel.XlBordersIndex.xlEdgeTop).Weight = Excel.XlBorderWeight.xlMedium
            xlrange.Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = Excel.XlBorderWeight.xlMedium
            xlrange.Borders(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlMedium
            xlrange.Borders(Excel.XlBordersIndex.xlEdgeLeft).Weight = Excel.XlBorderWeight.xlMedium

            xlrange.Borders(Excel.XlBordersIndex.xlInsideHorizontal).LineStyle = Excel.XlLineStyle.xlLineStyleNone
            xlrange.Borders(Excel.XlBordersIndex.xlInsideVertical).LineStyle = Excel.XlLineStyle.xlLineStyleNone
        Next


        FinalReleaseComObject(xlrange)
        'xlrange = Nothing
    End Function

    Function chart_init(ByVal chart As Excel.ChartObject, ByVal title_text As String, ByVal chart_title As String, ByVal x_text As String, ByVal y_text As String, ByVal x_max As String, ByVal x_min As String, ByVal y_max As String, ByVal y_min As String, ByVal chart_type As String) As Integer
        Dim heigh As String
        Dim width As String
        Dim height_temp As Integer
        Dim width_temp As Integer
        Dim value As Double

        xlSheet.Activate()
        'Update Title
        report_title(title_text, chart_col, chart_row, chart_width, 1, chart_title_color)

        'Chart Init
        chart_row = chart_row + 1

        chart_top = ConvertToLetter(chart_col) & chart_row

        heigh = chart_top & ":" & ConvertToLetter(chart_col) & (chart_row + chart_height - 1)
        height_temp = xlSheet.Range(heigh).Height


        width = chart_top & ":" & ConvertToLetter(chart_col + chart_width - 1) & chart_row
        width_temp = xlSheet.Range(width).Width


        chart = xlSheet.ChartObjects.add(xlSheet.Range(chart_top).Left, xlSheet.Range(chart_top).Top, width_temp, height_temp)



        xlchart = chart.Chart



        xlchart.ChartType = Excel.XlChartType.xlXYScatterSmoothNoMarkers

        With xlchart
            .HasTitle = True
            .ChartTitle.Text = chart_title
            .ChartTitle.Font.Bold = False
            .ChartTitle.Font.Size = 14

        End With


        With xlchart.Axes(1)
            .HasTitle = True
            .HasMajorGridlines = True
            .AxisTitle.Text = x_text
            .AxisTitle.Font.Bold = False


            If chart_type = "Log" Then
                .ScaleType = Excel.XlScaleType.xlScaleLogarithmic
                .HasMinorGridlines = True
                .MinimumScale = 0.000001
                .CrossesAt = 0.000001

            Else
                .ScaleType = Excel.XlScaleType.xlScaleLinear

                If x_min = "" Then
                    .MinimumScaleIsAuto = True
                Else
                    value = Val(x_min)
                    .MinimumScale = value
                End If


            End If

            If x_max = "" Then
                .MaximumScaleIsAuto = True
            Else
                value = Val(x_max)
                .MaximumScale = value
            End If


        End With


        With xlchart.Axes(2)
            .HasTitle = True
            .AxisTitle.Text = y_text
            .AxisTitle.Font.Bold = False
            If y_min = "" Then
                .MinimumScaleIsAuto = True
            Else
                value = Val(y_min)
                .MinimumScale = value
            End If
            If y_max = "" Then
                .MaximumScaleIsAuto = True
            Else
                value = Val(y_max)
                .MaximumScale = value
            End If
            '.MinimumScale = y_min
            '.MaximumScale = y_max
            '.MinimumScaleIsAuto =
            '.MaximumScaleIsAuto = 
        End With

        FinalReleaseComObject(xlchart)
        xlchart = Nothing
    End Function


    Function chart_add_series(ByVal Test_sheet As String, ByVal chart_name As Excel.ChartObject, ByVal chart_num As Integer, ByVal series_text As String, ByVal x_col As Integer, ByVal y_col As Integer, ByVal linedash As Boolean) As Integer
        Dim num As Integer
        xlSheet.Activate()
        chart_name = xlSheet.ChartObjects(chart_num)
        xlchart = chart_name.Chart
        xlchart.SeriesCollection.NewSeries()
        num = xlchart.SeriesCollection.Count
        xlchart.FullSeriesCollection(num).Name = series_text
        xlchart.FullSeriesCollection(num).XValues = "='" & Test_sheet & "'!$" & ConvertToLetter(x_col) & "$" & chart_row_start & ":$" & ConvertToLetter(x_col) & "$" & chart_row_stop
        xlchart.FullSeriesCollection(num).Values = "='" & Test_sheet & "'!$" & ConvertToLetter(y_col) & "$" & chart_row_start & ":$" & ConvertToLetter(y_col) & "$" & chart_row_stop

        If linedash = True Then

            With xlchart.FullSeriesCollection(num).Format.Line
                .Visible = Microsoft.Office.Core.MsoTriState.msoTrue
                .DashStyle = Microsoft.Office.Core.MsoLineDashStyle.msoLineDash
                '    .Weight = 2
            End With

        End If

        FinalReleaseComObject(xlchart)
        xlchart = Nothing
    End Function




End Module
