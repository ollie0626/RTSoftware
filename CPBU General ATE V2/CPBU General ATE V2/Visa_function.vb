Module Visa_function

    Public defaultRM As Integer
    Public visa_status As Integer

    Public visa_response As String = Space(VI_FIND_BUFLEN)
    Public visa_descriptor As String = Space(VI_FIND_BUFLEN)

    Public vi As Integer
    Public visa_count As Integer

    Public retcount As Integer


    Function visa_write(ByVal vi_desc As String, ByVal vi_dev As Integer, ByVal ts As String) As Integer
        Dim i As Integer


        visa_status = viWrite(vi_dev, ts, Len(ts), retcount)

        ' visa write will repeat if write fail.
        ' each times interval 100ms.
        If visa_status <> VI_SUCCESS Then
                For i = 0 To 5

                    System.Windows.Forms.Application.DoEvents()
                    If run = False Then
                        Exit For

                    End If
                    Delay(100)
                    visa_status = viWrite(vi_dev, ts, Len(ts), retcount)
                    If visa_status = VI_SUCCESS Then
                        Exit For
                    ElseIf visa_status = VI_ERROR_CONN_LOST Then
                        viOpen(defaultRM, vi_desc, VI_NO_LOCK, 2000, vi_dev)
                    End If
                Next
            End If





    End Function


  


  


    Function visa_scan() As String()

        Dim i As Integer = 1
        Dim temp(3) As String
        Dim temp1() As String
        Dim num As Integer = 0
        visa_status = viOpenDefaultRM(defaultRM)

        'viStatusDesc(defaultRM, visa_status, visa_response)
        viFindRsrc(defaultRM, "?*INSTR", vi, visa_count, visa_descriptor)
        'temp(0) = visa_descriptor

        For i = 0 To visa_count - 1
            If i > 0 Then
                viFindNext(vi, visa_descriptor)
            End If
            temp1 = Split(visa_descriptor, "::")
            If (Mid(temp1(0), 1, 3) = "TCP") Then
                If (temp1(2) = "inst0") Then
                    ReDim Preserve temp(num)
                    temp(num) = visa_descriptor
                    num = num + 1
                End If
              
            ElseIf (Mid(temp1(0), 1, 4) <> "ASRL") Then
                ReDim Preserve temp(num)
                temp(num) = visa_descriptor
                num = num + 1

            End If
         


        Next

        'While (i < visa_count)
        '    viFindNext(vi, visa_descriptor)

        '    temp1 = Split(visa_descriptor, "::")
        '    If temp1(2) = "inst0" Then
        '        temp(i) = visa_descriptor

        '    End If

        '    i = i + 1

        'End While


        viClose(defaultRM)

        Return temp

    End Function


    Function visa_name(ByVal visa As String) As String()
        Dim ts As String

        Dim temp() As String = {"", "", "", "", "", "", "", "", "", "", "", "", ""}

        viOpenDefaultRM(defaultRM)
        viOpen(defaultRM, visa, VI_NO_LOCK, 2000, vi)




        ts = "*IDN?"


        viWrite(vi, ts, Len(ts), retcount)

        visa_status = viRead(vi, visa_response, Len(visa_response), retcount)

        If visa_status = VI_SUCCESS Then
            If retcount > 0 Then
                temp = Split(visa_response, ",")

            End If
        Else
            temp(0) = "Error"
        End If



        viClose(vi)
        viClose(defaultRM)

        Return temp

    End Function

End Module
