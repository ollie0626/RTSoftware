
Imports System.IO
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Runtime.InteropServices.Marshal

Module Module_SaveData


    Public starup_path As String = Application.StartupPath
    Public stability_file As String = starup_path & "\stability_data.txt"
    Public efficiency_file As String = starup_path & "\efficiency_data.txt"
    Public loadR_file As String = starup_path & "\loadR_data.txt"
    Public jitter_file As String = starup_path & "\jitter_data.txt"
    Public line_file As String = starup_path & "\line_data.txt"


    Public stability_sheet As String
    Public jitter_sheet As String
    Public eff_sheet As String
    Public loadR_sheet As String
    Public line_sheet As String

    Public stable_sel As Integer = 0
    Public eff_sel As Integer = 1
    Public loadR_sel As Integer = 2
    Public jitter_sel As Integer = 3
    Public line_sel As Integer = 4

    Public test_file As String = starup_path & "\text.txt"
    Public test_sheet As String = "工作表1"
    Public test_sel As Integer = 5

    Public stable_col_len As Integer

    Public jitter_start_row As List(Of Integer) = New List(Of Integer)()
    Public jitter_start_col As Integer
    Public jitter_stop_col As Integer 'jitter_stop_col = jitter_start_col + jitter_col.Length

    Public lineR_start_row As List(Of Integer) = New List(Of Integer)()
    Public lineR_vin_col As List(Of Integer) = New List(Of Integer)()
    Public lineR_col_len As Integer

    Public loadR_start_row As List(Of Integer) = New List(Of Integer)
    Public loadR_vin_col As List(Of Integer) = New List(Of Integer)()
    Public loadR_col_len As Integer

    Public eff_start_row As List(Of Integer) = New List(Of Integer)()
    Public eff_vin_col As List(Of Integer) = New List(Of Integer)()
    Public eff_vin_col_len As Integer


    Public add_dut2 As String = "_dut2"

    ' param item :  0 -> stability
    '               1 -> efficiency
    '               2 -> load regulation
    '               3 -> jitter
    '               4 -> line regulation
    '               5 -> test
    Function SaveDataToFile(ByVal data_list As List(Of String), ByVal pass_fail As String, ByVal item_sel As Integer, Optional ByVal dut_sel As Integer = 0) As Boolean
        Dim data_buf As String = ""
        Dim sw As StreamWriter
        Dim sr As StreamReader
        Dim path_sel As String = ""

        For Each item As String In data_list : data_buf += item & vbTab : Next
        data_buf += pass_fail

        Select Case item_sel
            Case stable_sel : path_sel = stability_file
            Case eff_sel : path_sel = efficiency_file
            Case loadR_sel : path_sel = loadR_file
            Case jitter_sel : path_sel = jitter_file
            Case line_sel : path_sel = line_file
            Case 5 : path_sel = test_file
        End Select
        If path_sel = "" Then : Return False : End If

        Dim buf As String = ""
        If dut_sel = 1 Then : path_sel = path_sel & add_dut2 : End If
        Try

            If File.Exists(path_sel) Then
                sr = New StreamReader(path_sel)
                buf = sr.ReadToEnd()
                sr.Close()
            End If

            sw = New StreamWriter(path_sel)
            buf += data_buf
            sw.Write(buf & vbNewLine)
            sw.Close()

            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    ' param item :  0 -> stability
    '               1 -> efficiency
    '               2 -> load regulation
    '               3 -> jitter
    '               4 -> line regulation
    '               5 -> test
    Function TxtToExcel(ByVal item_sel As Integer, ByVal start_row As Integer, ByVal idx As Integer) As Boolean
        Dim txt_path As String = ""
        Dim sheet_name As String = ""
        Dim sr As StreamReader
        Dim _range As Excel.Range

        Dim start_col As String = ""
        Dim end_col As String = ""

        Select Case item_sel
            Case stable_sel : txt_path = stability_file : sheet_name = stability_sheet : start_col = "M" : end_col = "AK"
            Case eff_sel : txt_path = efficiency_file : sheet_name = eff_sheet
            Case loadR_sel : txt_path = loadR_file : sheet_name = loadR_sheet
            Case jitter_sel : txt_path = jitter_file : sheet_name = jitter_sheet
            Case line_sel : txt_path = line_file : sheet_name = line_sheet
            Case 5 : txt_path = test_file : sheet_name = test_sheet : start_col = "A" : end_col = "G"
        End Select

        If txt_path = "" Then : Return False : End If
        If sheet_name = "" Then : Return False : End If

        sr = New StreamReader(txt_path)
        Dim line = sr.ReadToEnd()
        xlSheet = xlBook.Sheets(sheet_name)
        ' transfer string to double
        Dim str_ar() = line.Split(vbNewLine)
        Dim col_number As Integer = 0
        Dim res As String = ""

        ' transfer string to double
        For Each item As String In str_ar
            item = item.Replace(vbLf, "")
            Dim temp() = item.Split(vbTab)
            Dim dou_ar As List(Of Double) = New List(Of Double)()
            For Each data As String In temp

                If data = "" Then : Return False : End If
                If data = PASS Or data = FAIL Then
                    res = data
                Else
                    dou_ar.Add(Convert.ToDouble(data))
                End If
            Next

            Select Case item_sel
                Case stable_sel, 5
                    ' stability and test case
                    _range = xlSheet.Range(start_col & start_row, end_col & start_row) ' row, col
                    col_number = 38
                    _range.Value = dou_ar.ToArray()
                Case eff_sel
                    ' eff case
                    _range = xlSheet.Range(
                        ConvertToLetter(eff_vin_col(idx)) & start_row,
                        ConvertToLetter(eff_vin_col(idx) + eff_vin_col_len - 1) & start_row)
                    _range.Value = dou_ar.ToArray()
                    col_number = eff_vin_col(idx) + eff_vin_col_len - 1
                Case loadR_sel, line_sel
                    ' 2 : load regulation
                    ' 4 : line regulation
                    Dim _range_temp As Excel.Range
                    Dim _range_copy As Excel.Range
                    Dim col_sel As String = ""

                    If item_sel = loadR_sel Then : col_sel = ConvertToLetter(loadR_vin_col(idx)) & start_row : End If
                    If item_sel = line_sel Then : col_sel = ConvertToLetter(lineR_vin_col(idx)) & start_row : End If


                    _range = xlSheet.Range(col_sel)
                    _range.Value = dou_ar.ToArray()

                    If item_sel = loadR_sel Then : col_number = loadR_vin_col(0) + loadR_col_len + 1 : End If
                    If item_sel = line_sel Then : col_number = lineR_vin_col(0) + lineR_col_len + 1 : End If

                    'FinalReleaseComObject(_range_temp)
                    'FinalReleaseComObject(_range_copy)
                    FinalReleaseComObject(_range)
                Case jitter_sel
                    ' jitter case
                    _range = xlSheet.Range(ConvertToLetter(jitter_start_col) & start_row, ConvertToLetter(jitter_stop_col - 1) & start_row)
                    col_number = jitter_stop_col - 1
                    _range.Value = dou_ar.ToArray()
                    FinalReleaseComObject(_range)
            End Select


            PassAndFailToExcel(item_sel, col_number, start_row, res)
            start_row += 1
        Next


        FinalReleaseComObject(xlSheet)
        xlSheet = Nothing
        xlBook.Save()
        sr.Close()
        Return True
    End Function

    Function TxttoExcel_v2(ByVal item_sel As Integer, ByVal start_row As Integer, ByVal data_idx As Integer, ByVal data_len As Integer, Optional ByVal dut_sel As Integer = 0) As Boolean

        Dim txt_path As String = ""
        Dim sheet_name As String = ""
        Dim sr As StreamReader
        Dim start_col As String = ""
        Dim end_col As String = ""

        Select Case item_sel
            Case stable_sel : txt_path = stability_file : sheet_name = stability_sheet : start_col = "M" : end_col = "AL"
            Case jitter_sel : txt_path = jitter_file : sheet_name = jitter_sheet : start_col = ConvertToLetter(jitter_start_col) : end_col = ConvertToLetter(jitter_stop_col - 1)
            Case eff_sel : txt_path = efficiency_file : sheet_name = eff_sheet : start_col = ConvertToLetter(eff_vin_col(idx)) : end_col = ConvertToLetter(eff_vin_col(idx) + eff_vin_col_len - 1)
        End Select

        If txt_path = "" Then : Return False : End If
        If sheet_name = "" Then : Return False : End If

        ' open text file
        sr = New StreamReader(txt_path)
        Dim line = sr.ReadToEnd()
        ' transfer string to double
        Dim str_ar() = line.Split(vbNewLine)
        Dim col_number As Integer = 0
        Dim res As String = ""
        sr.Close()

        Dim start_data_pos As Integer = data_idx * data_len
        Dim stop_data_pos As Integer = data_idx * data_len + data_len
        Dim all_data As List(Of List(Of String)) = New List(Of List(Of String))()

        For line_idx = 0 To str_ar.Length - 1
            If line_idx >= start_data_pos And line_idx < stop_data_pos Then
                Dim item As String = str_ar(line_idx).Replace(vbLf, "")
                ' string data
                Dim temp() = item.Split(vbTab)
                all_data.Add(temp.ToList())
                ' converter to double
                'Dim doubleAry As Double() = Array.ConvertAll(item.Split(vbTab), New Converter(Of String, Double)(AddressOf Double.Parse))
                'Dim dou_ar As List(Of List(Of Double)) = New List(Of List(Of Double))()
                'dou_ar.Add(doubleAry.ToList())
            End If
        Next

        If dut_sel = 1 Then
            sheet_name = sheet_name & add_dut2
        End If

        ' past data to Excel
        xlSheet = xlBook.Sheets(sheet_name)
        xlrange = xlSheet.Range(start_col & start_row, end_col & (start_row + data_len))
        xlrange.Value = all_data.ToArray()

        FinalReleaseComObject(xlSheet)
        FinalReleaseComObject(xlrange)

        xlSheet = Nothing
        xlBook.Save()
        sr.Close()
        Return True
    End Function



    Function PassAndFailToExcel(ByVal item_sel As Integer,
                                ByVal col As Integer,
                                ByVal row As Integer,
                                ByVal pass_fail As String)
        Dim _range As Excel.Range
        Dim sheet_name As String = ""

        Select Case item_sel
            Case stable_sel : sheet_name = stability_sheet
            Case eff_sel : sheet_name = eff_sheet
            Case loadR_sel : sheet_name = loadR_sheet
            Case jitter_sel : sheet_name = jitter_sheet
            Case line_sel : sheet_name = line_sheet
            Case 5 : sheet_name = test_sheet
        End Select

        xlSheet = xlBook.Sheets(sheet_name)

        _range = xlSheet.Range(ConvertToLetter(col) & row)
        _range.Value = pass_fail
        If pass_fail = FAIL Then
            _range.Interior.Color = test_fail_color
        End If
    End Function




    Function ClearTxtFile(ByVal item_sel As Integer) As Boolean
        Dim sw As StreamWriter
        Dim path_sel As String = ""
        Dim data_buf As String = ""

        Select Case item_sel
            Case 0 : path_sel = stability_file
            Case 1 : path_sel = efficiency_file
            Case 2 : path_sel = loadR_file
            Case 3 : path_sel = jitter_file
            Case 4 : path_sel = line_file
            Case 5 : path_sel = test_file
        End Select

        If path_sel = "" Then : Return False : End If

        Try
            sw = New StreamWriter(path_sel)
            sw.Write(data_buf)
            sw.Close()
            Return True
        Catch ex As Exception
            Return False
        End Try

        Return True
    End Function

    Public Sub Clear0To4TxtFile()
        For i As Integer = 0 To 4
            ClearTxtFile(i)
        Next
    End Sub


End Module


'_range_temp = xlSheet.Range("A1")
'_range_temp.Value = dou_ar.ToArray()

' past data to A1 and transfer row to columns
'With xlSheet
'    .Range("A1", _range_temp.End(Excel.XlDirection.xlToRight)).Copy()
'    .Range(col_sel).PasteSpecial(
'                            Excel.XlPasteType.xlPasteAll,
'                            Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone,
'                            False, True)
'End With

'_range_copy = xlSheet.Range("A1", _range_temp.End(Excel.XlDirection.xlToRight))
'_range_copy.Delete()

'col_number = IIf(item_sel = loadR_sel,
'                loadR_vin_col(idx) + loadR_col_len,
'                lineR_vin_col(idx) + lineR_col_len)

