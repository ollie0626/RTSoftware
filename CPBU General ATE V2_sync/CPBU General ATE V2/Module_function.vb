Imports System.Threading
Imports System.Drawing


Module Module_function
    Public time_stop As Integer = 0
    Public First As Boolean = True
    Public run As Boolean = False
    Public pause As Boolean = False
    Public run_stop_num As Integer = 1

    Public Resolution As Integer
    Public sf_name As String
    Public pf_name As String


    'Time
    Public hour As String
    Public minute As String
    Public second As String

    Public run_second As Integer
    Public run_time As String
    Public test_time As String


    'Math.Ceiling() 無條件進位, Math.Floor() 捨去小數

    Function Open_file(ByVal file_name As String) As Integer
        Process.Start(Environment.CurrentDirectory & "\" & file_name)
    End Function
    Function SendEmail(ByVal send_to As String, ByVal subject As String, ByVal Body As String, ByVal send_file_path As String, ByVal send_file As String) As Integer
        ' Create an Outlook application.
        Dim oApp As Object
        oApp = CreateObject("Outlook.application")

        ' Create a new MailItem.
        Dim oMsg As Object
        oMsg = oApp.CreateItem(0)
        oMsg.Subject = subject
        oMsg.Body = Body
        oMsg.To = send_to

        If send_file_path = "" Or send_file_path = " " Then
            ' Send
            oMsg.Send()


            ' Clean up
            oApp = Nothing
            oMsg = Nothing
        Else

            Dim sSource As String = send_file_path
            Dim sDisplayName As String = send_file

            Dim sBodyLen As String = oMsg.Body.Length
            Dim oAttachs As Object = oMsg.Attachments
            Dim oAttach As Object
            oAttach = oAttachs.Add(sSource, , sBodyLen + 1, sDisplayName)

            ' Send
            oMsg.Send()

            ' Clean up
            oApp = Nothing
            oMsg = Nothing


            oAttach = Nothing
            oAttachs = Nothing

        End If



    End Function


    Function Delay(ByVal minisecond As Integer) As Integer

        Dim stopwatch As Stopwatch = stopwatch.StartNew
        Thread.Sleep(minisecond)
        stopwatch.Stop()

    End Function

  


    '將hex的data統一化,不足的補0. Ex: hex=F, data_byte=2 --> hex_data="0F" 
    Function hex_data(ByVal dec_data As Integer, ByVal data_byte As Integer) As String
        'Dim i As Integer
        'hex_data = "0x"
        'hex_data = ""
        'If data_byte <> Len(Hex(dec_data)) Then
        '    For i = 1 To data_byte - Len(Hex(dec_data))
        '        hex_data = hex_data & "0"
        '    Next
        'End If

        'hex_data = hex_data & Hex(dec_data)

        hex_data = dec_data.ToString("X" & data_byte)


        Return hex_data
    End Function


    ' I2C bit set (按下"1"-> "0", 按"0"->"1"，並將算出來的值作運算)
    Function bit2data(ByVal bit_set As String, ByVal device As Object, ByVal bit As Integer) As String


        If bit_set = "0" Then
            bit_set = "1"
            device.Value = device.Value + 2 ^ bit
        Else
            bit_set = "0"
            device.Value = device.Value - 2 ^ bit

        End If

        Return bit_set

    End Function


    Function data2bit(ByVal device() As Object, ByVal data As Integer) As Integer
        Dim i As Integer
        For i = 0 To 7
            device(i).Text = data Mod 2
            data = Int(data / 2)
        Next
    End Function


    'I2C data set (十進制轉二進制)
    Function data_set(ByVal data As Integer) As Integer()

        Dim i, dec As Integer
        Dim a(7) As Integer
        dec = data
        For i = 0 To 7
            a(i) = dec Mod 2
            dec = Int(dec / 2)
        Next
        Return a
    End Function

    '將dec的data統一化,不足的補0. 
    Function dec_data(ByVal data As String, ByVal bits As Integer) As String

        Dim i As Integer
        Dim dec, max As Integer

        dec = Len(data)
        max = Len(CStr(2 ^ bits - 1))
        dec_data = ""

        If dec < max Then
            For i = 1 To max - dec
                dec_data = dec_data & "0"
            Next
        End If

        dec_data = dec_data & data
        Return dec_data

    End Function

    'check before register bit
    Function bit_check(ByVal data As Integer, ByVal bit_start As Integer, ByVal bit_end As Integer) As Integer
        'data=register data
        'bit_data=check register bit  Ex:bit_data=3 (0b00000011)
        Dim bit_data As Integer = 0
        Dim i As Integer
        Dim temp As Integer

        For i = bit_start To bit_end
            bit_data = bit_data + 2 ^ i
        Next

        temp = data Or (255 - bit_data)
        temp = temp And bit_data
        temp = temp / (2 ^ bit_start)
        Return temp

    End Function
    'check before register data
    Function data_check(ByVal data As Integer, ByVal bit_start As Integer, ByVal bit_end As Integer) As Integer
        'data=register data
        'bit_data=check register bit  Ex:bit_data=3 (0b00000011)
        Dim bit_data As Integer = 0
        Dim i As Integer
        Dim temp As Integer

        For i = bit_start To bit_end
            bit_data = bit_data + 2 ^ i
        Next

        temp = data Or (255 - bit_data)
        temp = temp And bit_data
        Return temp

    End Function


    '將word分成兩個byte
    Function word2byte_data(ByVal firstbyte As String, ByVal word_data As Integer) As Byte()
        Dim Hbyte, Lbyte As Integer
        Dim data(1) As Byte

        Hbyte = Int(word_data / 256)
        Lbyte = word_data Mod 256


        'High byte first
        If firstbyte = "H" Then

            data(0) = Hbyte
            data(1) = Lbyte
        Else
            data(0) = Lbyte
            data(1) = Hbyte

        End If



        Return data
    End Function

    Function ConvertToLetter(ByRef iCol As Integer) As String

        Dim dividend As Integer = iCol
        Dim columnName As String = String.Empty
        Dim modulo As Integer

        While dividend > 0
            modulo = (dividend - 1) Mod 26
            columnName = Convert.ToChar(65 + modulo).ToString() & columnName
            dividend = CInt((dividend - modulo) / 26)
        End While

        Return columnName



    End Function


    Function GetExcelColumnName(ByVal columnNumber As Integer) As String
        Dim dividend As Integer = columnNumber
        Dim columnName As String = String.Empty
        Dim modulo As Integer

        While dividend > 0
            modulo = (dividend - 1) Mod 26
            columnName = Convert.ToChar(65 + modulo).ToString() & columnName
            dividend = CInt((dividend - modulo) / 26)
        End While

        Return columnName
    End Function


End Module
