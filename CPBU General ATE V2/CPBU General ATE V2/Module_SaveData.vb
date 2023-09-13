
Imports System.IO
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Runtime.InteropServices.Marshal

Module Module_SaveData


    Public starup_path As String = Application.StartupPath
    Public stability_file As String = starup_path & "\\stability_data.txt"
    Public efficiency_file As String = starup_path & "\\efficiency_data.txt"
    Public loadR_file As String = starup_path & "\\loadR_data.txt"
    Public jitter_file As String = starup_path & "\\jitter_data.txt"
    Public line_file As String = starup_path & "\\line_data.txt"

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


    ' param item :  0 -> stability
    '               1 -> efficiency
    '               2 -> load regulation
    '               3 -> jitter
    '               4 -> line regulation
    Function SaveDataToFile(ByVal data_list As List(Of Double), ByVal pass_fail As String, ByVal item_sel As Integer) As Boolean
        Dim data_buf As String = ""
        Dim sw As StreamWriter
        Dim path_sel As String = ""

        For Each item As String In data_list : data_buf += item & "\t" : Next
        data_buf += pass_fail

        Select Case item_sel
            Case 0 : path_sel = stability_file
            Case 1 : path_sel = efficiency_file
            Case 2 : path_sel = loadR_file
            Case 3 : path_sel = jitter_file
        End Select
        If path_sel = "" Then : Return False : End If
        Try
            sw = New StreamWriter(path_sel)
            sw.WriteLine(data_buf)
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
    Function TxtToExcel(ByVal item_sel As Integer, ByVal row As Integer, ByVal eff_idx As Integer) As Boolean
        Dim txt_path As String = ""
        Dim sheet_name As String = ""
        Dim sr As StreamReader
        Dim _range As Excel.Range
        Dim start_col As String = ""
        Dim end_col As String = ""

        Dim eff_start() As String = New String() {"M", "U", "AC", "AK"}
        Dim eff_end() As String = New String() {"S", "AA", "AI", "AQ"}

        Select Case item_sel
            Case 0 : txt_path = stability_file : sheet_name = stability_sheet : start_col = "M" : end_col = "AL"
            Case 1 : txt_path = efficiency_file : sheet_name = eff_sheet : start_col = eff_start(eff_idx) : end_col = eff_end(eff_idx)
            Case 2 : txt_path = loadR_file : sheet_name = loadR_sheet : start_col = "M" : end_col = "R"
            Case 3 : txt_path = jitter_file : sheet_name = jitter_sheet : start_col = "CG" : end_col = "CQ"
            Case 4 : txt_path = line_file : sheet_name = line_sheet : start_col = "M" : end_col = "Q"
        End Select

        If txt_path = "" Then : Return False : End If
        If sheet_name = "" Then : Return False : End If

        sr = New StreamReader(txt_path)
        Dim line = sr.ReadLine()
        xlSheet = xlBook.Sheets(sheet_name)

        Dim str_ar() = line.Split(New String() {"\r\n"}, StringSplitOptions.None)
        For Each item As String In str_ar
            _range = xlSheet.Range(start_col & row, end_col & row) ' row, col
            _range.Value = item
            row += 1
        Next
        FinalReleaseComObject(xlSheet)
        xlSheet = Nothing
        xlBook.Save()
        Return True
    End Function

End Module
