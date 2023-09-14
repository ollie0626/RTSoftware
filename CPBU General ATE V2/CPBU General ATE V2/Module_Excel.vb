
Imports System.Runtime.InteropServices
Imports System.Text.RegularExpressions
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Runtime.InteropServices.Marshal
Module Module_Excel


    Public xlApp As Excel.Application
    Public xlBook As Excel.Workbook

    Public xlSheet As Excel.Worksheet
    Public shpPic As Excel.Shape

    Public xlrange As Excel.Range
    Public col As Integer = 1
    Public row As Integer = 1
    Public data_start_row As Integer = 20

    Public row_index As Integer = 2
    Public excel_check As Integer
    Public xlchart As Excel.Chart
    Public myChart As Excel.ChartObject

    Public chart_col As Integer
    Public chart_row As Integer = test_row

    Public pic_col As Integer
    Public pic_row As Integer = test_row
  
    ' 取得最後的row:      last_row= xlSheet_wave.Range(ConvertToLetter(1) & 1).CurrentRegion.Rows.Count
    ' 取得最後的col:       last_col = xlSheet.Range(ConvertToLetter(col) & row).CurrentRegion.Columns.Count
    ' Dim time_volt As Object
    '  time_volt = xlSheet_wave.Range(xlApp.Cells(start_row, start_col), xlApp.Cells(stop_row, Stop_col)).Value()
    '由1開始，time_volt (1,1); 若為兩個矩陣就是 time_volt (1,1),time_volt (1,2)


    Public Function UsedRows(ByVal FileName As String, ByVal SheetName As String) As Integer
        '取有多少的row
        Dim RowsUsed As Integer = -1

        If IO.File.Exists(FileName) Then
            Dim xlAplication As Excel.Application = Nothing
            Dim xlWorkBooks As Excel.Workbooks = Nothing
            Dim xlWorkBook As Excel.Workbook = Nothing
            Dim xlWorkSheet As Excel.Worksheet = Nothing
            Dim xlWorkSheets As Excel.Sheets = Nothing

            xlAplication = New Excel.Application
            xlAplication.DisplayAlerts = False
            xlWorkBooks = xlAplication.Workbooks
            xlWorkBook = xlWorkBooks.Open(FileName)

            xlAplication.Visible = False

            xlWorkSheets = xlWorkBook.Sheets

            For x As Integer = 1 To xlWorkSheets.Count

                xlWorkSheet = CType(xlWorkSheets(x), Excel.Worksheet)

                If xlWorkSheet.Name = SheetName Then
                    Dim xlCells As Excel.Range = Nothing
                    xlCells = xlWorkSheet.Cells

                    Dim xlTempRange As Excel.Range = xlCells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell)

                    RowsUsed = xlTempRange.Row
                    Runtime.InteropServices.Marshal.FinalReleaseComObject(xlTempRange)
                    xlTempRange = Nothing

                    Runtime.InteropServices.Marshal.FinalReleaseComObject(xlCells)
                    xlCells = Nothing

                    Exit For
                End If

                Runtime.InteropServices.Marshal.FinalReleaseComObject(xlWorkSheet)
                xlWorkSheet = Nothing

            Next

            xlWorkBook.Close()
            xlAplication.UserControl = True
            xlAplication.Quit()



            ReleaseComObject(xlWorkSheets)
            ReleaseComObject(xlWorkSheet)
            ReleaseComObject(xlWorkBook)
            ReleaseComObject(xlWorkBooks)
            ReleaseComObject(xlAplication)


        Else
            Throw New Exception("'" & FileName & "' not found.")
        End If

        Return RowsUsed

    End Function


    Public Sub ReleaseComObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        End Try
    End Sub

    Function paste_picture(ByVal picture_file As String, ByVal pic_top As String, ByVal width_temp As Double, ByVal height_temp As Double) As Integer


        shpPic = xlSheet.Shapes.AddPicture(picture_file, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, xlSheet.Range(pic_top).Left, xlSheet.Range(pic_top).Top, width_temp, height_temp)

    End Function


    
   



End Module
