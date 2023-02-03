
Imports System.IO

Public Class Form1

    Dim main_ver As String = "IN528 Tool v1.4"
    Dim byt_Bin1(255) As Byte
    Dim byt_Bin2(255) As Byte
    Dim flagTable() As Boolean = {True, True, True, True, True, True, True, True}

    Dim WRBuffer() As Byte = New Byte(&H1F + 1) {}
    Dim RDBuffer() As Byte = New Byte(&H1F + 1) {}

    Dim byt00() As String = {"AVDD_EN", "VGL1_EN", "VGL2_EN", "VGH_EN", "HAVDD_EN", "VCORE_EN", "VIO_EN", "VDO_EN"}
    Dim byt01() As String = {"GMA_EN", "RESET_EN", "VCOM_EN", "NTC_EN"}
    Dim byt02() As String = {"AVDD"}
    Dim byt03() As String = {"VGL1"}
    Dim byt04() As String = {"VGL2"}
    Dim byt05() As String = {"VGH"}
    Dim byt06() As String = {"VGHT"}
    Dim byt07() As String = {"HAVDD"}
    Dim byt08() As String = {"VCORE"}
    Dim byt09() As String = {"VIO"}
    Dim byt0A() As String = {""}
    Dim byt0B() As String = {"LDO"}
    Dim byt0C() As String = {"GMA1"}
    Dim byt0D() As String = {"GMA2"}
    Dim byt0E() As String = {"VCOM"}
    Dim byt0F() As String = {"RESET"}
    Dim byt10() As String = {"AVDD Freq", "AVDD SR"}
    Dim byt11() As String = {"AVDD Dly", "AVDD SST"}
    Dim byt12() As String = {"AVDD OCP"}
    Dim byt13() As String = {"VGL Dly", "VGL SST", "VGL Freq", "VGL Mode"}
    Dim byt14() As String = {"VGL2 Dly", "VGL2 SST"}
    Dim byt15() As String = {"VGH Dly", "VGH SST", "VGH Freq"}
    Dim byt16() As String = {"VGH SR"}
    Dim byt17() As String = {"VCORE Dly", "VCORE SST", "VCORE Freq"}
    Dim byt18() As String = {"VIO Dly", "VIO SST", "VIO Freq"}
    Dim byt19() As String = {"LDO Dly", "VCOM Dly"}
    Dim byt1A() As String = {"Reset Dly"}
    Dim byt1B() As String = {"VCOM power off", "VCOM Disc", "VGL Disc", "VGH Disc", "AVDD Disc", "LDO Disc", "VCORE Disc", "VIO Disc"}
    Dim byt1C() As String = {"VCOM Disc SR", "VCOM SST", "AVDD Mode", "VGH Mode", "VCORE Mode", "VIO Mode"}

    Dim MapList As New List(Of List(Of String))
    Dim hDevice As Integer

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Text = main_ver
        'RegSetting()

        MapList.Add(byt00.ToList())
        MapList.Add(byt01.ToList())
        MapList.Add(byt02.ToList())
        MapList.Add(byt03.ToList())
        MapList.Add(byt04.ToList())
        MapList.Add(byt05.ToList())
        MapList.Add(byt06.ToList())
        MapList.Add(byt07.ToList())
        MapList.Add(byt08.ToList())
        MapList.Add(byt09.ToList())
        MapList.Add(byt0A.ToList())
        MapList.Add(byt0B.ToList())
        MapList.Add(byt0C.ToList())
        MapList.Add(byt0D.ToList())
        MapList.Add(byt0E.ToList())
        MapList.Add(byt0F.ToList())
        MapList.Add(byt10.ToList())
        MapList.Add(byt11.ToList())
        MapList.Add(byt12.ToList())
        MapList.Add(byt13.ToList())
        MapList.Add(byt14.ToList())
        MapList.Add(byt15.ToList())
        MapList.Add(byt16.ToList())
        MapList.Add(byt17.ToList())
        MapList.Add(byt18.ToList())
        MapList.Add(byt19.ToList())
        MapList.Add(byt1A.ToList())
        MapList.Add(byt1B.ToList())
        MapList.Add(byt1C.ToList())
        ' 2-D List get element
        'MessageBox.Show(MapList(0)(0))
        CB_A0.SelectedIndex = 0
        hDevice = RTBB_ConnectToBridgeByIndex(0)
        If hDevice <> 0 Then
            'MsgBox("Link RTBridge Board Successful!!", main_ver, MsgBoxStyle.DefaultButton1)
            MessageBox.Show("Link RTBridge Board Successful!!", main_ver, MessageBoxButtons.OK)
        Else
            'MsgBox("Please Link RTBridge Board!!", main_ver, MsgBoxStyle.DefaultButton1)
            MessageBox.Show("Please Link RTBridge Board!!", main_ver, MessageBoxButtons.OK)
        End If

        GUIInitial()
    End Sub

    Private Sub CalculateResult()

        Dim Wr_res As Integer
        Dim Rd_res As Integer

        For i As Integer = 0 To WRBuffer.Length - 1
            Wr_res += WRBuffer(i)
            Rd_res += RDBuffer(i)
        Next

        CheckSum_resall.Text = (Wr_res Xor Rd_res).ToString("X")

        Wr_res -= WRBuffer(&HE)
        Rd_res -= RDBuffer(&HE)

        CheckSum_resvcom.Text = (Wr_res Xor Rd_res).ToString("X")

    End Sub

    Private Sub CalculateFlagNum()
        Dim res As Integer = 0

        For i As Integer = 0 To flagTable.Length - 1
            If Not (flagTable(i)) Then
                res += 1
            End If
        Next
        D8_label.Text = res.ToString()
    End Sub

    Private Sub GUIInitial()
        W00_0.Maximum = 1
        W00_1.Maximum = 1
        W00_2.Maximum = 1
        W00_3.Maximum = 1
        W00_4.Maximum = 1
        W00_5.Maximum = 1
        W00_6.Maximum = 1
        W00_7.Maximum = 1

        W01_0.Maximum = 1
        W01_1.Maximum = 1
        W01_4.Maximum = 1
        W01_5.Maximum = 1

        W02.Maximum = &H7F
        W03.Maximum = &H7F
        W04.Maximum = &H7F
        W05.Maximum = &H3F
        W06.Maximum = &H1F
        W07.Maximum = &H2F
        W08.Maximum = &H1F
        W09.Maximum = &H2F
        W0B.Maximum = &HF

        W0C.Maximum = &H1F
        W0D.Maximum = &H1F
        W0E.Maximum = &HFF
        W0F.Maximum = &H7

        W10_0.Maximum = &H7
        W10_3.Maximum = &H7
        W11_0.Maximum = &H7
        W11_3.Maximum = &H7

        W12_0.Maximum = &H7
        W13_0.Maximum = &H7
        W13_3.Maximum = &H3
        W13_5.Maximum = &H1
        W13_6.Maximum = &H3

        W14_0.Maximum = &H7
        W14_3.Maximum = &H3

        W15_0.Maximum = &H7
        W15_3.Maximum = &H3
        W15_5.Maximum = &H7

        W16_0.Maximum = &H7
        W17_0.Maximum = &H3
        W17_2.Maximum = &H7
        W17_5.Maximum = &H3

        W18_0.Maximum = &H3
        W18_2.Maximum = &H7
        W18_5.Maximum = &H3

        W19_0.Maximum = &H3
        W19_2.Maximum = &H1F

        W1A.Maximum = &HF
        W1B_0.Maximum = &H1
        W1B_1.Maximum = &H1
        W1B_2.Maximum = &H1
        W1B_3.Maximum = &H1
        W1B_4.Maximum = &H1
        W1B_5.Maximum = &H1
        W1B_6.Maximum = &H1
        W1B_7.Maximum = &H1

        W1C_0.Maximum = &H3
        W1C_2.Maximum = &H3
        W1C_4.Maximum = &H1
        W1C_5.Maximum = &H1
        W1C_6.Maximum = &H1
        W1C_7.Maximum = &H1

        W1D.Maximum = &HF
        W1E.Maximum = &HFF
        W1F.Maximum = &HFF

        Bar_HAVDD.Maximum = &H2F
        Bar_VIO.Maximum = &H2F
        Bar_VGL2Dly.Maximum = &H7
        Bar_VGHDly.Maximum = &H7


        W00_0.Value = 1
        W00_1.Value = 1
        W00_2.Value = 1
        W00_3.Value = 1
        W00_4.Value = 1
        W00_5.Value = 1
        W00_6.Value = 1
        W00_7.Value = 1
        W01_0.Value = 1
        W01_1.Value = 1
        W01_4.Value = 1
        W01_5.Value = 1
        W02.Value = 1
        W03.Value = 1
        W04.Value = 1
        W05.Value = 1
        W06.Value = 1
        W07.Value = 1
        W08.Value = 1
        W09.Value = 1
        W0B.Value = 1
        W0C.Value = 1
        W0D.Value = 1
        W0E.Value = 1
        W0F.Value = 1
        W10_0.Value = 1
        W10_3.Value = 1
        W11_0.Value = 1
        W11_3.Value = 1
        W12_0.Value = 1
        W13_0.Value = 1
        W13_3.Value = 1
        W13_5.Value = 1
        W13_6.Value = 1
        W14_0.Value = 1
        W14_3.Value = 1
        W15_0.Value = 1
        W15_3.Value = 1
        W15_5.Value = 1
        W16_0.Value = 1
        W17_0.Value = 1
        W17_2.Value = 1
        W17_5.Value = 1
        W18_0.Value = 1
        W18_2.Value = 1
        W18_5.Value = 1
        W19_0.Value = 1
        W19_2.Value = 1
        W1A.Value = 1
        W1B_0.Value = 1
        W1B_1.Value = 1
        W1B_2.Value = 1
        W1B_3.Value = 1
        W1B_4.Value = 1
        W1B_5.Value = 1
        W1B_6.Value = 1
        W1B_7.Value = 1
        W1C_0.Value = 1
        W1C_2.Value = 1
        W1C_4.Value = 1
        W1C_5.Value = 1
        W1C_6.Value = 1
        W1C_7.Value = 1
        W1D.Value = 1
        W1E.Value = 1
        W1F.Value = 1


        W00_0.Value = 0
        W00_1.Value = 0
        W00_2.Value = 0
        W00_3.Value = 0
        W00_4.Value = 0
        W00_5.Value = 0
        W00_6.Value = 0
        W00_7.Value = 0
        W01_0.Value = 0
        W01_1.Value = 0
        W01_4.Value = 0
        W01_5.Value = 0
        W02.Value = 0
        W03.Value = 0
        W04.Value = 0
        W05.Value = 0
        W06.Value = 0
        W07.Value = 0
        W08.Value = 0
        W09.Value = 0
        W0B.Value = 0
        W0C.Value = 0
        W0D.Value = 0
        W0E.Value = 0
        W0F.Value = 0
        W10_0.Value = 0
        W10_3.Value = 0
        W11_0.Value = 0
        W11_3.Value = 0
        W12_0.Value = 0
        W13_0.Value = 0
        W13_3.Value = 0
        W13_5.Value = 0
        W13_6.Value = 0
        W14_0.Value = 0
        W14_3.Value = 0
        W15_0.Value = 0
        W15_3.Value = 0
        W15_5.Value = 0
        W16_0.Value = 0
        W17_0.Value = 0
        W17_2.Value = 0
        W17_5.Value = 0
        W18_0.Value = 0
        W18_2.Value = 0
        W18_5.Value = 0
        W19_0.Value = 0
        W19_2.Value = 0
        W1A.Value = 0
        W1B_0.Value = 0
        W1B_1.Value = 0
        W1B_2.Value = 0
        W1B_3.Value = 0
        W1B_4.Value = 0
        W1B_5.Value = 0
        W1B_6.Value = 0
        W1B_7.Value = 0
        W1C_0.Value = 0
        W1C_2.Value = 0
        W1C_4.Value = 0
        W1C_5.Value = 0
        W1C_6.Value = 0
        W1C_7.Value = 0
        W1D.Value = 0
        W1E.Value = 0
        W1F.Value = 0
    End Sub

    Private Sub CK00_7_CheckedChanged(sender As Object, e As EventArgs) Handles CK00_7.CheckedChanged, CK00_6.CheckedChanged, CK00_5.CheckedChanged, CK00_4.CheckedChanged, CK00_3.CheckedChanged, CK00_2.CheckedChanged, CK00_1.CheckedChanged, CK00_0.CheckedChanged
        Dim byt00_bit() As NumericUpDown = {W00_0, W00_1, W00_2, W00_3, W00_4, W00_5, W00_6, W00_7}
        Dim byt00_ck() As CheckBox = {CK00_0, CK00_1, CK00_2, CK00_3, CK00_4, CK00_5, CK00_6, CK00_7}
        Dim byt00_name() As Label = {AVDD_EN, VGL1_EN, VGL2_EN, VGH_EN, HAVDD_EN, VCORE_EN, VIO_EN, LDO_EN}

        Dim bit7 As Byte = CK00_7.Checked And &H1
        Dim bit6 As Byte = CK00_6.Checked And &H1
        Dim bit5 As Byte = CK00_5.Checked And &H1
        Dim bit4 As Byte = CK00_4.Checked And &H1
        Dim bit3 As Byte = CK00_3.Checked And &H1
        Dim bit2 As Byte = CK00_2.Checked And &H1
        Dim bit1 As Byte = CK00_1.Checked And &H1
        Dim bit0 As Byte = CK00_0.Checked And &H1

        Dim byte_data() As Byte = {bit0, bit1, bit2, bit3, bit4, bit5, bit6, bit7}

        For i As Integer = 0 To 7
            byt00_bit(i).Value = byte_data(i)
        Next

        For i As Integer = 0 To 7
            If byt00_ck(i).Checked Then
                byt00_name(i).Text = "Enable"
                byt00_name(i).ForeColor = Color.Black
            Else
                byt00_name(i).Text = "Disable"
                byt00_name(i).ForeColor = Color.Red
            End If
        Next
    End Sub

    Private Sub W00_7_ValueChanged(sender As Object, e As EventArgs) Handles W00_7.ValueChanged, W00_6.ValueChanged, W00_5.ValueChanged, W00_4.ValueChanged, W00_3.ValueChanged, W00_2.ValueChanged, W00_1.ValueChanged, W00_0.ValueChanged
        Dim byt00_bit() As NumericUpDown = {W00_0, W00_1, W00_2, W00_3, W00_4, W00_5, W00_6, W00_7}
        Dim byt00_ck() As CheckBox = {CK00_0, CK00_1, CK00_2, CK00_3, CK00_4, CK00_5, CK00_6, CK00_7}
        Dim byt00_name() As Label = {LDO_EN, VIO_EN, VCORE_EN, HAVDD_EN, VGH_EN, VGL2_EN, VGL1_EN, AVDD_EN}

        For i As Integer = 0 To 7
            byt00_ck(i).Checked = byt00_bit(i).Value And &H1
        Next
    End Sub

    Private Sub CK01_5_CheckedChanged(sender As Object, e As EventArgs) Handles CK01_5.CheckedChanged, CK01_4.CheckedChanged, CK01_1.CheckedChanged, CK01_0.CheckedChanged
        Dim byt01_bit() As NumericUpDown = {W01_0, W01_1, W01_4, W01_5}
        Dim byt01_ck() As CheckBox = {CK01_0, CK01_1, CK01_4, CK01_5}
        Dim byt01_name() As Label = {lab01_0, lab01_1, lab01_4, lab01_5}

        W01_0.Value = CK01_0.Checked And &H1
        W01_1.Value = CK01_1.Checked And &H1
        W01_4.Value = CK01_4.Checked And &H1
        W01_5.Value = CK01_5.Checked And &H1

        For i As Integer = 0 To 3
            If byt01_ck(i).Checked Then
                byt01_name(i).Text = "Enable"
                byt01_name(i).ForeColor = Color.Black
            Else
                byt01_name(i).Text = "Disable"
                byt01_name(i).ForeColor = Color.Red
            End If
        Next
    End Sub

    Private Sub Bar_AVDD_Scroll(sender As Object, e As ScrollEventArgs) Handles Bar_AVDD.Scroll
        W02.Value = Bar_AVDD.Value
    End Sub

    Private Sub W02_ValueChanged(sender As Object, e As EventArgs) Handles W02.ValueChanged
        Bar_AVDD.Value = W02.Value
        Dim AVDD As Double = (Bar_AVDD.Value * 0.1) + 7

        If AVDD > 14 Then
            AVDD = 14
        End If
        AVDD_V.Text = String.Format("{0:0.0}V", AVDD)

        Dim offset As Double
        If W0C.Value <= &H12 Then
            offset = (W0C.Value + 2) * 0.05
        Else
            offset = (&H12 + 2) * 0.05
        End If
        GMA1V.Text = String.Format("{0:0.00}V", AVDD - offset)
    End Sub

    Private Sub Bar_VGL1_Scroll(sender As Object, e As ScrollEventArgs) Handles Bar_VGL1.Scroll
        W03.Value = Bar_VGL1.Value
    End Sub

    Private Sub W03_ValueChanged(sender As Object, e As EventArgs) Handles W03.ValueChanged
        Bar_VGL1.Value = W03.Value
        Dim VGL1 As Double = (Bar_VGL1.Value * 0.1) + 2
        If VGL1 > 14.5 Then
            VGL1 = 14.5
        End If
        VGL1V.Text = String.Format("-{0:0.0}V", VGL1)
    End Sub

    Private Sub Bar_VGL2_Scroll(sender As Object, e As ScrollEventArgs) Handles Bar_VGL2.Scroll
        W04.Value = Bar_VGL2.Value
    End Sub

    Private Sub W04_ValueChanged(sender As Object, e As EventArgs) Handles W04.ValueChanged
        Bar_VGL2.Value = W04.Value
        Dim VGL2 As Double = Bar_VGL2.Value * 0.1 + 2
        If VGL2 > 14.5 Then
            VGL2 = 14.5
        End If
        VGL2V.Text = String.Format("-{0:0.0}V", VGL2)
    End Sub

    Private Sub Bar_VGH_Scroll(sender As Object, e As ScrollEventArgs) Handles Bar_VGH.Scroll
        W05.Value = Bar_VGH.Value
    End Sub

    Private Sub W05_ValueChanged(sender As Object, e As EventArgs) Handles W05.ValueChanged
        Bar_VGH.Value = W05.Value
        Dim VGH As Double = (Bar_VGH.Value * 0.5) + 5
        If VGH > 36 Then
            VGH = 36
        End If
        VGHV.Text = String.Format("{0:0.0}V", VGH)


        '-----------------------------------------------------------------
        Dim VGHT As Double = (Bar_VGHT.Value * 1) + 5
        If VGHT >= VGH Then
            D2_label.Text = "True"
            flagTable(1) = True
        Else
            D2_label.Text = "False"
            flagTable(1) = False
        End If
        CalculateFlagNum()
    End Sub

    Private Sub Bar_VGHT_Scroll(sender As Object, e As ScrollEventArgs) Handles Bar_VGHT.Scroll
        W06.Value = Bar_VGHT.Value
    End Sub

    Private Sub Bar_HAVDD_Scroll(sender As Object, e As ScrollEventArgs) Handles Bar_HAVDD.Scroll
        W07.Value = Bar_HAVDD.Value
    End Sub

    Private Sub Bar_VCORE_Scroll(sender As Object, e As ScrollEventArgs) Handles Bar_VCORE.Scroll
        W08.Value = Bar_VCORE.Value
    End Sub

    Private Sub Bar_VIO_Scroll(sender As Object, e As ScrollEventArgs) Handles Bar_VIO.Scroll
        W09.Value = Bar_VIO.Value
    End Sub

    Private Sub Bar_LDO_Scroll(sender As Object, e As ScrollEventArgs) Handles Bar_LDO.Scroll
        W0B.Value = Bar_LDO.Value
    End Sub

    Private Sub Bar_GMA1_Scroll(sender As Object, e As ScrollEventArgs) Handles Bar_GMA1.Scroll
        W0C.Value = Bar_GMA1.Value
    End Sub

    Private Sub Bar_GMA2_Scroll(sender As Object, e As ScrollEventArgs) Handles Bar_GMA2.Scroll
        W0D.Value = Bar_GMA2.Value
    End Sub

    Private Sub Bar_VCOM_Scroll(sender As Object, e As ScrollEventArgs) Handles Bar_VCOM.Scroll
        W0E.Value = Bar_VCOM.Value
    End Sub

    Private Sub Bar_Reset_Scroll(sender As Object, e As ScrollEventArgs) Handles Bar_Reset.Scroll
        W0F.Value = Bar_Reset.Value
    End Sub

    Private Sub Bar_AVDDLx_Scroll(sender As Object, e As ScrollEventArgs) Handles Bar_AVDDLx.Scroll, Bar_AVDDLxSR.Scroll
        Dim byte10 As Byte = Bar_AVDDLx.Value Or (Bar_AVDDLxSR.Value << 3)
        W10_0.Value = Bar_AVDDLx.Value
        W10_3.Value = Bar_AVDDLxSR.Value
    End Sub

    Private Sub Bar_AVDDDly_Scroll(sender As Object, e As ScrollEventArgs) Handles Bar_AVDDDly.Scroll, Bar_AVDDSST.Scroll
        W11_0.Value = Bar_AVDDDly.Value
        W11_3.Value = Bar_AVDDSST.Value
    End Sub

    Private Sub Bar_AVDDOCP_Scroll(sender As Object, e As ScrollEventArgs) Handles Bar_AVDDOCP.Scroll
        W12_0.Value = Bar_AVDDOCP.Value
    End Sub

    Private Sub Bar_VGLDly_Scroll(sender As Object, e As ScrollEventArgs) Handles Bar_VGLDly.Scroll, Bar_VGLSST.Scroll, Bar_VGLMode.Scroll, Bar_VGLFreq.Scroll
        W13_0.Value = Bar_VGLDly.Value
        W13_3.Value = Bar_VGLSST.Value
        W13_5.Value = Bar_VGLFreq.Value
        W13_6.Value = Bar_VGLMode.Value
    End Sub

    Private Sub Bar_VGL2Dly_Scroll(sender As Object, e As ScrollEventArgs) Handles Bar_VGL2Dly.Scroll
        W14_0.Value = Bar_VGL2Dly.Value
    End Sub

    Private Sub Bar_VGHDly_Scroll(sender As Object, e As ScrollEventArgs) Handles Bar_VGHDly.Scroll, Bar_VGHSST.Scroll, Bar_VGHFreq.Scroll
        W15_0.Value = Bar_VGHDly.Value
        W15_3.Value = Bar_VGHSST.Value
        W15_5.Value = Bar_VGHFreq.Value
    End Sub

    Private Sub Bar_VGHSR_Scroll(sender As Object, e As ScrollEventArgs) Handles Bar_VGHSR.Scroll
        W16_0.Value = Bar_VGHSR.Value
    End Sub

    Private Sub Bar_VcoreDly_Scroll(sender As Object, e As ScrollEventArgs) Handles Bar_VcoreDly.Scroll, Bar_VcoreSR.Scroll, Bar_VcoreFreq.Scroll
        W17_0.Value = Bar_VcoreDly.Value
        W17_2.Value = Bar_VcoreFreq.Value
        W17_5.Value = Bar_VcoreSR.Value
    End Sub

    Private Sub Bar_VioDly_Scroll(sender As Object, e As ScrollEventArgs) Handles Bar_VioDly.Scroll, Bar_VIOSR.Scroll, Bar_VIOFreq.Scroll
        W18_0.Value = Bar_VioDly.Value
        W18_2.Value = Bar_VIOFreq.Value
        W18_5.Value = Bar_VIOSR.Value
    End Sub

    Private Sub Bar_LDODly_Scroll(sender As Object, e As ScrollEventArgs) Handles Bar_LDODly.Scroll, Bar_VcomDly.Scroll
        W19_0.Value = Bar_LDODly.Value
        W19_2.Value = Bar_VcomDly.Value
    End Sub

    Private Sub Bar_ResetDly_Scroll(sender As Object, e As ScrollEventArgs) Handles Bar_ResetDly.Scroll
        W1A.Value = Bar_ResetDly.Value
    End Sub

    Private Sub Bar_Vcom_power_off_Scroll(sender As Object, e As ScrollEventArgs) Handles Bar_Vcom_power_off.Scroll, Bar_VIODisc.Scroll, Bar_VGLDisc.Scroll, Bar_VGHDisc.Scroll, Bar_VcoreDisc.Scroll, Bar_VcomDisc.Scroll, Bar_LDODisc.Scroll, Bar_AVDDDisc.Scroll
        W1B_0.Value = Bar_Vcom_power_off.Value
        W1B_1.Value = Bar_VcomDisc.Value
        W1B_2.Value = Bar_VGLDisc.Value
        W1B_3.Value = Bar_VGHDisc.Value
        W1B_4.Value = Bar_AVDDDisc.Value
        W1B_5.Value = Bar_LDODisc.Value
        W1B_6.Value = Bar_VcoreDisc.Value
        W1B_7.Value = Bar_VIODisc.Value
    End Sub

    Private Sub Bar_VcomDiscSR_Scroll(sender As Object, e As ScrollEventArgs) Handles Bar_VcomDiscSR.Scroll, Bar_VGHMode.Scroll, Bar_VcoreMode.Scroll, Bar_VcomSST.Scroll, Bar_AVDDMode.Scroll, Bar_VIOMode.Scroll
        W1C_0.Value = Bar_VcomDiscSR.Value
        W1C_2.Value = Bar_VcomSST.Value
        W1C_4.Value = Bar_AVDDMode.Value
        W1C_5.Value = Bar_VGHMode.Value
        W1C_6.Value = Bar_VcoreMode.Value
        W1C_7.Value = Bar_VIOMode.Value

    End Sub

    Private Sub Bar_Inx1_Scroll(sender As Object, e As ScrollEventArgs) Handles Bar_Inx1.Scroll
        W1D.Value = Bar_Inx1.Value
    End Sub

    Private Sub Bar_Inx2_Scroll(sender As Object, e As ScrollEventArgs) Handles Bar_Inx2.Scroll
        W1E.Value = Bar_Inx2.Value
    End Sub

    Private Sub Bar_Inx3_Scroll(sender As Object, e As ScrollEventArgs) Handles Bar_Inx3.Scroll
        W1F.Value = Bar_Inx3.Value
    End Sub

    Private Sub W07_ValueChanged(sender As Object, e As EventArgs) Handles W07.ValueChanged
        Bar_HAVDD.Value = W07.Value
        Dim code As Integer = W07.Value
        Dim HAVDD As Double
        If code <= 3 Then
            HAVDD = code * 0.1 + 3.5
        Else
            HAVDD = (code - 3) * 0.05 + 3.8
        End If
        HAVDDV.Text = String.Format("{0:0.00}V", HAVDD)
        'Dim HAVDD As Double = Bar_HAVDD.Value * 
    End Sub

    Private Sub W08_ValueChanged(sender As Object, e As EventArgs) Handles W08.ValueChanged
        Bar_VCORE.Value = W08.Value
        Dim Vcore As Double = (0.05 * Bar_VCORE.Value) + 0.8
        If Vcore > 2 Then
            Vcore = 2
        End If
        VCOREV.Text = String.Format("{0:0.00}V", Vcore)

    End Sub

    Private Sub W09_ValueChanged(sender As Object, e As EventArgs) Handles W09.ValueChanged
        Bar_VIO.Value = W09.Value
        Dim VIO As Double = (Bar_VIO.Value * 0.05) + 1
        If VIO > 2.8 Then
            VIO = 2.8
        End If

        VIOV.Text = String.Format("{0:0.0}V", VIO)
    End Sub

    Private Sub W0B_ValueChanged(sender As Object, e As EventArgs) Handles W0B.ValueChanged
        Bar_LDO.Value = W0B.Value
        Dim LDO As Double = (Bar_LDO.Value * 0.1) + 1.7
        If LDO > 2.8 Then
            LDO = 2.8
        End If
        LDOV.Text = String.Format("{0:0.0}V", LDO)

    End Sub

    Private Sub W06_ValueChanged(sender As Object, e As EventArgs) Handles W06.ValueChanged
        Bar_VGHT.Value = W06.Value
        Dim VGHT As Double = (Bar_VGHT.Value * 1) + 5
        VGHTV.Text = String.Format("{0:0}V", VGHT)

        '--------------------------------------------------------
        Dim VGH As Double = (Bar_VGH.Value * 0.5) + 5
        If VGH > 36 Then
            VGH = 36
        End If

        If VGHT >= VGH Then
            D2_label.Text = "True"
        Else
            D2_label.Text = "False"
        End If
        CalculateFlagNum()

    End Sub

    Private Sub W0D_ValueChanged(sender As Object, e As EventArgs) Handles W0D.ValueChanged
        Bar_GMA2.Value = W0D.Value
        Dim GMA2 As Double = (Bar_GMA2.Value * 0.05) + 0.1
        If GMA2 > 1 Then
            GMA2 = 1
        End If
        GMA2V.Text = String.Format("{0:0.0}V", GMA2)
    End Sub

    Private Sub W0E_ValueChanged(sender As Object, e As EventArgs) Handles W0E.ValueChanged
        Bar_VCOM.Value = W0E.Value
        Dim Vcom As Double = (Bar_VCOM.Value * 0.02) + 1.4
        VCOMV.Text = String.Format("{0:0.00}V", Vcom)
    End Sub

    Private Sub W0F_ValueChanged(sender As Object, e As EventArgs) Handles W0F.ValueChanged
        Bar_Reset.Value = W0F.Value
        Dim Reset As Double = (Bar_Reset.Value * 0.1) + 2
        RESETV.Text = String.Format("{0:0.0}V", Reset)
    End Sub

    Private Sub W10_0_ValueChanged(sender As Object, e As EventArgs) Handles W10_0.ValueChanged
        Bar_AVDDLx.Value = W10_0.Value
        Dim AVDDLxDes() As String = {"600kHz", "715kHz", "800kHz", "933kHz", "1000kHz", "1225kHz", "1225kHz", "1225kHz"}
        AVDD_Lx.Text = AVDDLxDes(W10_0.Value)
        W13_5_ValueChanged(Nothing, Nothing)
    End Sub

    Private Sub W10_3_ValueChanged(sender As Object, e As EventArgs) Handles W10_3.ValueChanged
        Bar_AVDDLxSR.Value = W10_3.Value
        Dim AVDDSRDes() As String = {"120%", "120%", "100%", "100%", "80%", "80%", "60%", "60%"}
        AVDD_SR.Text = AVDDSRDes(W10_3.Value)
    End Sub

    Private Sub W11_0_ValueChanged(sender As Object, e As EventArgs) Handles W11_0.ValueChanged
        Bar_AVDDDly.Value = W11_0.Value
        Dim AVDDDlyDes() As String = {"1mS", "2mS", "3mS", "4mS", "5mS", "6mS", "7mS", "12mS"}
        AVDD_Dly.Text = AVDDDlyDes(W11_0.Value)

        ' -------------------------------------------------------------------------------
        Dim AVDDDly() As Integer = {1, 2, 3, 4, 5, 6, 7, 12}
        Dim VGHDly() As Integer = {2, 7, 18, 25, 34, 50, 100, 150}
        Dim VGLDly As Integer = 5 * Bar_VGLDly.Value

        If AVDDDly(W11_0.Value) <= VGHDly(W15_0.Value) Then
            D1_lable.Text = "True"
            flagTable(0) = True
        Else
            D1_lable.Text = "False"
            flagTable(0) = False
        End If

        If VGLDly >= AVDDDly(W11_0.Value) Then
            D5_label.Text = "True"
            flagTable(4) = True
        Else
            D5_label.Text = "False"
            flagTable(4) = True
        End If
        CalculateFlagNum()
    End Sub

    Private Sub W11_3_ValueChanged(sender As Object, e As EventArgs) Handles W11_3.ValueChanged
        Bar_AVDDSST.Value = W11_3.Value
        AVDD_SST.Text = String.Format("{0:0}mS", (W11_3.Value * 2) + 2)
    End Sub

    Private Sub W12_0_ValueChanged(sender As Object, e As EventArgs) Handles W12_0.ValueChanged
        Bar_AVDDOCP.Value = W12_0.Value
        Dim AVDDOCPDes() As String = {"2A", "1.5A", "1A", "0.5A", "0.4A", "0.3A", "0.2", "0.2"}
        AVDD_OCP.Text = AVDDOCPDes(W12_0.Value)
    End Sub

    Private Sub W13_0_ValueChanged(sender As Object, e As EventArgs) Handles W13_0.ValueChanged
        Bar_VGLDly.Value = W13_0.Value
        VGL_Dly.Text = String.Format("{0:0}mS", 5 * Bar_VGLDly.Value)

        '---------------------------------------------------------------
        Dim VGHDly() As Integer = {2, 7, 18, 25, 34, 50, 100, 150}
        Dim VGLDly As Integer = 5 * Bar_VGLDly.Value
        Dim AVDDDly() As Integer = {1, 2, 3, 4, 5, 6, 7, 12}
        Dim VGL2Dly As Integer = Bar_VGL2Dly.Value * 5

        If VGLDly >= VGHDly(W15_0.Value) Then
            D3_label.Text = "True"
            flagTable(2) = True
        Else
            D3_label.Text = "False"
            flagTable(2) = False
        End If


        If VGLDly >= AVDDDly(W11_0.Value) Then
            D5_label.Text = "True"
            flagTable(4) = True
        Else
            D5_label.Text = "False"
            flagTable(4) = False
        End If

        If VGLDly <= VGL2Dly Then
            D7_label.Text = "True"
            flagTable(6) = True
        Else
            D7_label.Text = "False"
            flagTable(6) = False
        End If

        CalculateFlagNum()
    End Sub

    Private Sub W13_3_ValueChanged(sender As Object, e As EventArgs) Handles W13_3.ValueChanged
        Bar_VGLSST.Value = W13_3.Value
        VGL_SST.Text = String.Format("{0:0}mS", 2 * Bar_VGLSST.Value)
    End Sub

    Private Sub W13_5_ValueChanged(sender As Object, e As EventArgs) Handles W13_5.ValueChanged
        Bar_VGLFreq.Value = W13_5.Value
        Dim AVDDFreq() As Double = {600, 715, 800, 933, 1000, 1225, 1225, 1225}
        Dim VGLFreX() As Double = {0.5, 1}
        VGL_Freq.Text = String.Format("{0:0}kHz", AVDDFreq(Bar_AVDDLx.Value) * VGLFreX(Bar_VGLFreq.Value))
    End Sub

    Private Sub W13_6_ValueChanged(sender As Object, e As EventArgs) Handles W13_6.ValueChanged
        Bar_VGLMode.Value = W13_6.Value
        Dim Arch() As String = {"mode1 VGL > -(AVDD-0.5V)", "mode2 -14.5V ≦ VGL ≦ -(AVDD-0.5V)", "mode3 VGL<-14.5V", "mode3 VGL<-14.5V"}
        VGL_Mode.Text = Arch(Bar_VGLMode.Value)
    End Sub

    Private Sub W14_0_ValueChanged(sender As Object, e As EventArgs) Handles W14_0.ValueChanged
        Bar_VGL2Dly.Value = W14_0.Value
        VGL2_Dly.Text = String.Format("{0:0}mS", Bar_VGL2Dly.Value * 5)

        Dim VGHDly() As Integer = {2, 7, 18, 25, 34, 50, 100, 150}
        Dim AVDDDly() As Integer = {1, 2, 3, 4, 5, 6, 7, 12}
        Dim VGL2Dly As Integer = Bar_VGL2Dly.Value * 5
        Dim VGLDly As Integer = 5 * Bar_VGLDly.Value
        If VGHDly(W15_0.Value) <= VGL2Dly Then
            D4_label.Text = "True"
            flagTable(3) = True
        Else
            D4_label.Text = "False"
            flagTable(4) = False
        End If

        If AVDDDly(W11_0.Value) <= VGL2Dly Then
            D6_label.Text = "True"
            flagTable(5) = True
        Else
            D6_label.Text = "False"
            flagTable(5) = False
        End If

        If VGLDly <= VGL2Dly Then
            D7_label.Text = "True"
            flagTable(6) = True
        Else
            D7_label.Text = "False"
            flagTable(6) = False
        End If

        CalculateFlagNum()
    End Sub

    Private Sub W14_3_ValueChanged(sender As Object, e As EventArgs) Handles W14_3.ValueChanged
        Bar_VGL2SST.Value = W14_3.Value
        VGL2_SST.Text = String.Format("{0:0}mS", Bar_VGL2SST.Value * 2 + 2)
    End Sub

    Private Sub W15_0_ValueChanged(sender As Object, e As EventArgs) Handles W15_0.ValueChanged
        Bar_VGHDly.Value = W15_0.Value
        Dim VGHDlyDes() As String = {"2mS", "7mS", "18mS", "25mS", "34mS", "50mS", "100mS", "150mS"}
        VGH_Dly.Text = VGHDlyDes(Bar_VGHDly.Value)

        '-------------------------------------------------------------------------------------------
        Dim AVDDDly() As Integer = {1, 2, 3, 4, 5, 6, 7, 12}
        Dim VGHDly() As Integer = {2, 7, 18, 25, 34, 50, 100, 150}
        If AVDDDly(W11_0.Value) <= VGHDly(W15_0.Value) Then
            D1_lable.Text = "True"
            flagTable(0) = True
        Else
            D1_lable.Text = "False"
            flagTable(0) = True
        End If

        Dim VGLDly As Integer = 5 * Bar_VGLDly.Value
        If VGLDly >= VGHDly(W15_0.Value) Then
            D3_label.Text = "True"
            flagTable(2) = True
        Else
            D3_label.Text = "False"
            flagTable(2) = False
        End If

        If VGHDly(W15_0.Value) <= Bar_VGL2Dly.Value * 5 Then
            D4_label.Text = "True"
            flagTable(3) = True
        Else
            D4_label.Text = "False"
            flagTable(3) = True
        End If

        CalculateFlagNum()
    End Sub

    Private Sub W15_3_ValueChanged(sender As Object, e As EventArgs) Handles W15_3.ValueChanged
        Bar_VGHSST.Value = W15_3.Value
        VGH_SST.Text = String.Format("{0:0}mS", Bar_VGHSST.Value * 2 + 2)
    End Sub

    Private Sub W15_5_ValueChanged(sender As Object, e As EventArgs) Handles W15_5.ValueChanged
        Bar_VGHFreq.Value = W15_5.Value
        Dim VGHFreqDes() As String = {"600kHz", "715kHz", "800kHz", "933kHz", "1000kHz", "1225kHz", "1225kHz", "1225kHz"}
        VGH_Freq.Text = VGHFreqDes(Bar_VGHFreq.Value)
    End Sub

    Private Sub W16_0_ValueChanged(sender As Object, e As EventArgs) Handles W16_0.ValueChanged
        Bar_VGHSR.Value = W16_0.Value
        Dim VGHSRDes() As String = {"120%", "120%", "100%", "100%", "80%", "80%", "60%", "60%"}
        VGH_SR.Text = VGHSRDes(Bar_VGHSR.Value)
    End Sub

    Private Sub Bar_VGL2SST_Scroll(sender As Object, e As ScrollEventArgs) Handles Bar_VGL2SST.Scroll
        W14_3.Value = Bar_VGL2SST.Value
    End Sub

    Private Sub W17_0_ValueChanged(sender As Object, e As EventArgs) Handles W17_0.ValueChanged
        Bar_VcoreDly.Value = W17_0.Value
        Vcore_Dly.Text = String.Format("{0:0}mS", Bar_VcoreDly.Value * 3)
    End Sub

    Private Sub W17_2_ValueChanged(sender As Object, e As EventArgs) Handles W17_2.ValueChanged
        Bar_VcoreFreq.Value = W17_2.Value
        Dim VcoreFreDes() As String = {"600kHz", "715kHz", "800kHz", "933kHz", "1000kHz", "1225kHz", "1225kHz", "1225kHz"}
        Vcore_Freq.Text = VcoreFreDes(Bar_VcoreFreq.Value)
    End Sub

    Private Sub W17_5_ValueChanged(sender As Object, e As EventArgs) Handles W17_5.ValueChanged
        Bar_VcoreSR.Value = W17_5.Value
        Dim VcoreSRDes() As String = {"120%", "100%", "80%", "60%"}
        Vcore_SR.Text = VcoreSRDes(Bar_VcoreSR.Value)
    End Sub

    Private Sub W18_0_ValueChanged(sender As Object, e As EventArgs) Handles W18_0.ValueChanged
        Bar_VioDly.Value = W18_0.Value
        Vio_Dly.Text = String.Format("{0}mS", Bar_VioDly.Value * 3)
    End Sub

    Private Sub W18_2_ValueChanged(sender As Object, e As EventArgs) Handles W18_2.ValueChanged
        Bar_VIOFreq.Value = W18_2.Value
        Dim VioFreDes() As String = {"600kHz", "715kHz", "800kHz", "933kHz", "1000kHz", "1225kHz", "1225kHz", "1225kHz"}
        Vio_freq.Text = VioFreDes(Bar_VIOFreq.Value)
    End Sub

    Private Sub W18_5_ValueChanged(sender As Object, e As EventArgs) Handles W18_5.ValueChanged
        Bar_VIOSR.Value = W18_5.Value
        Dim VioSRDes() As String = {"120%", "100%", "80%", "60%"}
        VIO_SR.Text = VioSRDes(Bar_VIOSR.Value)
    End Sub

    Private Sub W19_0_ValueChanged(sender As Object, e As EventArgs) Handles W19_0.ValueChanged
        Bar_LDODly.Value = W19_0.Value
        Dim LDODlyDes() As String = {"0mS", "15mS", "34mS", "45mS"}
        LDO_Dly.Text = LDODlyDes(Bar_LDODly.Value)
    End Sub

    Private Sub W19_2_ValueChanged(sender As Object, e As EventArgs) Handles W19_2.ValueChanged
        Bar_VcomDly.Value = W19_2.Value
        VCOM_Dly.Text = String.Format("{0}mS", 5 * Bar_VcomDly.Value)
    End Sub

    Private Sub W1A_ValueChanged(sender As Object, e As EventArgs) Handles W1A.ValueChanged
        Bar_ResetDly.Value = W1A.Value
        Reset_Dly.Text = String.Format("{0}mS", 5 * Bar_ResetDly.Value)
    End Sub

    Private Sub W1B_0_ValueChanged(sender As Object, e As EventArgs) Handles W1B_0.ValueChanged
        Bar_Vcom_power_off.Value = W1B_0.Value
        Dim power_off_Des() As String = {"UVLO_F", "Reset"}
        Vcom_power_off.Text = power_off_Des(Bar_Vcom_power_off.Value)
    End Sub

    Private Sub W1B_1_ValueChanged(sender As Object, e As EventArgs) Handles W1B_1.ValueChanged
        Bar_VcomDisc.Value = W1B_1.Value
        Dim Des() As String = {"Disable", "Enable"}
        Vcom_disc.Text = Des(Bar_VcomDisc.Value)
    End Sub

    Private Sub W1B_2_ValueChanged(sender As Object, e As EventArgs) Handles W1B_2.ValueChanged
        Bar_VGLDisc.Value = W1B_2.Value
        Dim Des() As String = {"Disable", "Enable"}
        VGL_Disc.Text = Des(Bar_VGLDisc.Value)
    End Sub

    Private Sub W1B_3_ValueChanged(sender As Object, e As EventArgs) Handles W1B_3.ValueChanged
        Bar_VGHDisc.Value = W1B_3.Value
        Dim Des() As String = {"Disable", "Enable"}
        VGH_Disc.Text = Des(Bar_VGHDisc.Value)
    End Sub

    Private Sub W1B_4_ValueChanged(sender As Object, e As EventArgs) Handles W1B_4.ValueChanged
        Bar_AVDDDisc.Value = W1B_4.Value
        Dim Des() As String = {"Disable", "Enable"}
        AVDD_Disc.Text = Des(Bar_AVDDDisc.Value)
    End Sub

    Private Sub W1B_5_ValueChanged(sender As Object, e As EventArgs) Handles W1B_5.ValueChanged
        Bar_LDODisc.Value = W1B_5.Value
        Dim Des() As String = {"Disable", "Enable"}
        LDO_Disc.Text = Des(Bar_LDODisc.Value)
    End Sub

    Private Sub W1B_6_ValueChanged(sender As Object, e As EventArgs) Handles W1B_6.ValueChanged
        Bar_VcoreDisc.Value = W1B_6.Value
        Dim Des() As String = {"Disable", "Enable"}
        Vcore_Disc.Text = Des(Bar_VcoreDisc.Value)
    End Sub

    Private Sub W1B_7_ValueChanged(sender As Object, e As EventArgs) Handles W1B_7.ValueChanged
        Bar_VIODisc.Value = W1B_7.Value
        Dim Des() As String = {"Disable", "Enable"}
        Vio_Disc.Text = Des(Bar_VIODisc.Value)
    End Sub

    Private Sub W1C_0_ValueChanged(sender As Object, e As EventArgs) Handles W1C_0.ValueChanged
        Bar_VcomDiscSR.Value = W1C_0.Value
        Dim Des() As String = {"Slowest", "Slow", "Normal", "Fastest"}
        Vcom_Disc_SR.Text = Des(Bar_VcomDiscSR.Value)
    End Sub

    Private Sub W1C_2_ValueChanged(sender As Object, e As EventArgs) Handles W1C_2.ValueChanged
        Bar_VcomSST.Value = W1C_2.Value
        Dim Des() As String = {"Slowest", "Slow", "Normal", "Fastest"}
        Vcom_SST.Text = Des(Bar_VcomSST.Value)
    End Sub

    Private Sub W1C_4_ValueChanged(sender As Object, e As EventArgs) Handles W1C_4.ValueChanged
        Bar_AVDDMode.Value = W1C_4.Value
        Dim Des() As String = {"PSM Mode", "PWM Mode"}
        AVDD_Mode.Text = Des(Bar_AVDDMode.Value)
    End Sub

    Private Sub W1C_5_ValueChanged(sender As Object, e As EventArgs) Handles W1C_5.ValueChanged
        Bar_VGHMode.Value = W1C_5.Value
        Dim Des() As String = {"PSM Mode", "PWM Mode"}
        VGH_Mode.Text = Des(Bar_VGHMode.Value)
    End Sub

    Private Sub W1C_6_ValueChanged(sender As Object, e As EventArgs) Handles W1C_6.ValueChanged
        Bar_VcoreMode.Value = W1C_6.Value
        Dim Des() As String = {"PSM Mode", "PWM Mode"}
        Vcore_Mode.Text = Des(Bar_VcoreMode.Value)
    End Sub

    Private Sub W1C_7_ValueChanged(sender As Object, e As EventArgs) Handles W1C_7.ValueChanged
        Bar_VIOMode.Value = W1C_7.Value
        Dim Des() As String = {"PSM Mode", "PWM Mode"}
        Vio_Mode.Text = Des(Bar_VIOMode.Value)
    End Sub

    Private Sub LoadBin1ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LoadBin1ToolStripMenuItem.Click
        Dim openDlg As OpenFileDialog = New OpenFileDialog()
        openDlg.Filter = "Bin Files (*.bin) | *.bin|All files (*.*)|*.*"
        openDlg.RestoreDirectory = True
        If openDlg.ShowDialog() = DialogResult.OK Then

            Dim FileInfo As New FileInfo(openDlg.FileName) 'get bin file data number
            Dim strPath As String = openDlg.FileName 'file path
            Dim fs As New FileStream(strPath, FileMode.Open, FileAccess.Read) 'file stream mode -> read
            Dim bytBytes() As Byte = New Byte(FileInfo.Length) {} 'create bin file buffer

            fs.Read(bytBytes, 0, bytBytes.Length) 'read bin file to buffer
            Array.Copy(bytBytes, byt_Bin1, bytBytes.Length)

            fs.Close()
        End If
    End Sub

    Private Sub LoadBin2ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LoadBin2ToolStripMenuItem.Click
        Dim openDlg As OpenFileDialog = New OpenFileDialog()
        openDlg.Filter = "Bin Files (*.bin) | *.bin|All files (*.*)|*.*"
        openDlg.RestoreDirectory = True
        If openDlg.ShowDialog() = DialogResult.OK Then

            Dim FileInfo As New FileInfo(openDlg.FileName) 'get bin file data number
            Dim strPath As String = openDlg.FileName 'file path
            Dim fs As New FileStream(strPath, FileMode.Open, FileAccess.Read) 'file stream mode -> read
            Dim bytBytes() As Byte = New Byte(FileInfo.Length) {} 'create bin file buffer

            fs.Read(bytBytes, 0, bytBytes.Length) 'read bin file to buffer
            Array.Copy(bytBytes, byt_Bin2, bytBytes.Length)

            fs.Close()
        End If
    End Sub

    Private Sub AddCompare_Info(ByVal Addr As Integer, ByRef idx As Integer, ByVal bit_loop() As Byte)

        Dim X_axis As Integer = 30
        Dim Y_axis As Integer = 30
        Dim X_axis_addr As Integer = X_axis + 95
        Dim X_axis_value As Integer = X_axis_addr + 95
        Dim X_axis_info As Integer = X_axis_value + 95

        For bit_idx As Integer = 0 To bit_loop.Length - 1
            If bit_loop(bit_idx) <> 0 Then
                ' print to GUI
                Dim Reg_name As Label = New Label()
                Reg_name.Location = New Point(X_axis, Y_axis * idx)
                Reg_name.Text = MapList(Addr)(bit_idx)
                Reg_name.AutoSize = True
                Panel1.Controls.Add(Reg_name)

                Dim Reg_addr As Label = New Label()
                Reg_addr.Location = New Point(X_axis_addr, Y_axis * idx)
                Reg_addr.Text = String.Format("{0:X2}h", Addr)
                Reg_addr.AutoSize = True
                Panel1.Controls.Add(Reg_addr)

                Dim Reg_value As Label = New Label()
                Reg_value.Location = New Point(X_axis_value, Y_axis * idx)
                Reg_value.Text = String.Format("0x{0:X2}", bit_loop(bit_idx))
                Reg_value.AutoSize = True
                Panel1.Controls.Add(Reg_value)


                Dim Reg_info As Label = New Label()
                Dim info As String
                Reg_info.Location = New Point(X_axis_info, Y_axis * idx)
                Reg_info.AutoSize = True
                info = ""

                Select Case (Addr)
                    Case &H0, &H1B, &H1
                        If Addr = &H1B And bit_idx = 0 Then
                            If bit_loop(bit_idx) = &H1 Then
                                info = "Reset"
                            Else
                                info = "UVLO_F"
                            End If
                        Else
                            If bit_loop(bit_idx) = &H1 Then
                                info = "Enable"
                            Else
                                info = "Disable"
                                Reg_info.ForeColor = Color.Red
                            End If
                        End If
                    Case &H2
                        Dim AVDD As Double = (bit_loop(bit_idx) * 0.1) + 7
                        If AVDD > 14 Then
                            AVDD = 14
                        End If
                        info = String.Format("{0:0.0}V", AVDD)
                    Case &H3
                        Dim VGL1 As Double = (bit_loop(bit_idx) * 0.1) + 2
                        If VGL1 > 14.5 Then
                            VGL1 = 14.5
                        End If
                        info = String.Format("-{0:0.0}V", VGL1)
                    Case &H4
                        Dim VGL2 As Double = (bit_loop(bit_idx) * 0.1) + 2
                        If VGL2 > 14.5 Then
                            VGL2 = 14.5
                        End If
                        info = String.Format("-{0:0.0}V", VGL2)
                    Case &H5
                        Dim Vol As Double = (bit_loop(bit_idx) * 0.5) + 5
                        If Vol > 35 Then
                            Vol = 35
                        End If
                        info = String.Format("{0:0.0}V", Vol)
                    Case &H6
                        Dim Vol As Double = (bit_loop(bit_idx) * 1) + 5
                        If Vol > 36 Then
                            Vol = 36
                        End If
                        info = String.Format("{0:0.0}V", Vol)
                    Case &H7 ' havdd need excel
                        Dim HAVDD As Double
                        Dim code As Integer = bit_loop(bit_idx)
                        If code <= 3 Then
                            HAVDD = code * 0.1 + 3.5
                        Else
                            HAVDD = (code - 3) * 0.05 + 3.8
                        End If
                        info = String.Format("{0:0.00}V", HAVDD)
                    Case &H8
                        Dim Vol As Double = (bit_loop(bit_idx) * 0.05) + 0.8
                        If Vol > 2 Then
                            Vol = 2
                        End If
                        info = String.Format("{0:0.0}V", Vol)
                    Case &H9
                        Dim Vol As Double = (bit_loop(bit_idx) * 0.05) + 1
                        If Vol > 2.8 Then
                            Vol = 2.8
                        End If
                        info = String.Format("{0:0.0}V", Vol)
                    Case &HA
                    Case &HB
                        Dim Vol As Double = (bit_loop(bit_idx) * 0.1) + 1.7
                        If Vol > 2.8 Then
                            Vol = 2.8
                        End If
                        info = String.Format("{0:0.0}V", Vol)
                    Case &HC ' need avdd
                        Dim offset As Double = (bit_loop(bit_idx) * 0.05)
                        info = String.Format("AVDD - {0:0.0}V", offset)

                    Case &HD
                        Dim Vol As Double = (bit_loop(bit_idx) * 0.05) + 0.1
                        If Vol > 1 Then
                            Vol = 1
                        End If
                        info = String.Format("{0:0.0}V", Vol)
                    Case &HE
                        Dim Vol As Double = (bit_loop(bit_idx) * 0.02) + 1.4
                        If Vol > 6.5 Then
                            Vol = 6.5
                        End If
                        info = String.Format("{0:0.0}V", Vol)
                    Case &HF
                        Dim Vol As Double = (bit_loop(bit_idx) * 0.1) + 2
                        If Vol > 2.7 Then
                            Vol = 2.7
                        End If
                        info = String.Format("{0:0.0}V", Vol)
                    Case &H10
                        Dim AVDDLxDes() As String = {"600kHz", "715kHz", "800kHz", "933kHz", "1000kHz", "1225kHz", "1225kHz", "1225kHz"}
                        Dim AVDDSRDes() As String = {"120%", "120%", "100%", "100%", "80%", "80%", "60%", "60%"}
                        Select Case (bit_idx)
                            Case 0
                                info = AVDDLxDes(bit_loop(bit_idx))
                            Case 1
                                info = AVDDSRDes(bit_loop(bit_idx))
                        End Select
                    Case &H11
                        Dim AVDDDlyDes() As String = {"1mS", "2mS", "3mS", "4mS", "5mS", "6mS", "7mS", "12mS"}
                        Select Case (bit_idx)
                            Case 0
                                info = AVDDDlyDes(bit_loop(bit_idx))
                            Case 1
                                info = String.Format("{0}mS", (bit_loop(bit_idx) * 2) + 2)
                        End Select
                    Case &H12
                        Dim AVDDOCPDes() As String = {"2A", "1.5A", "1A", "0.5A", "0.4A", "0.3A", "0.2A", "0.2A"}
                        info = AVDDOCPDes(bit_loop(bit_idx))
                    Case &H13
                        Select Case (bit_idx)
                            Case 0
                                info = String.Format("{0:0}mS", 5 * bit_loop(bit_idx))
                            Case 1
                                info = String.Format("{0:0}mS", 2 * bit_loop(bit_idx))
                            Case 2
                                Dim AVDDFreq() As Double = {600, 715, 800, 933, 1000, 1225, 1225, 1225}
                                Dim VGLFreX() As Double = {0.5, 1}
                                info = String.Format("{0:0}kHz", AVDDFreq(Bar_AVDDLx.Value) * VGLFreX(Bar_VGLFreq.Value))
                            Case 3
                                Dim Arch() As String = {"mode1 VGL > -(AVDD-0.5V)", "mode2 -14.5V ≦ VGL ≦ -(AVDD-0.5V)", "mode3 VGL<-14.5V", "mode3 VGL<-14.5V"}
                                info = Arch(bit_loop(bit_idx))
                        End Select

                    Case &H14
                        Select Case (bit_idx)
                            Case 0
                                info = String.Format("{0:0}mS", bit_loop(bit_idx) * 5)
                            Case 1
                                info = String.Format("{0:0}mS", bit_loop(bit_idx) * 2 + 2)
                        End Select
                    Case &H15
                        Select Case (bit_idx)
                            Case 0
                                Dim VGHDlyDes() As String = {"2mS", "7mS", "18mS", "25mS", "34mS", "50mS", "100mS", "150mS"}
                                info = VGHDlyDes(bit_loop(bit_idx))
                            Case 1
                                info = String.Format("{0:0}mS", bit_loop(bit_idx) * 2 + 2)
                            Case 2
                                Dim VGHFreqDes() As String = {"600kHz", "715kHz", "800kHz", "933kHz", "1000kHz", "1225kHz", "1225kHz", "1225kHz"}
                                info = VGHFreqDes(bit_loop(bit_idx))
                        End Select
                    Case &H16
                        Dim VGHSRDes() As String = {"120%", "120%", "100%", "100%", "80%", "80%", "60%", "60%"}
                        info = VGHSRDes(bit_loop(bit_idx))
                    Case &H17
                        Select Case (bit_idx)
                            Case 0
                                info = String.Format("{0:0}mS", bit_loop(bit_idx) * 3)
                            Case 1
                                Dim VcoreFreDes() As String = {"600kHz", "715kHz", "800kHz", "933kHz", "1000kHz", "1225kHz", "1225kHz", "1225kHz"}
                                info = VcoreFreDes(bit_loop(bit_idx))
                            Case 2
                                Dim VcoreSRDes() As String = {"120%", "100%", "80%", "60%"}
                                info = VcoreSRDes(bit_loop(bit_idx))
                        End Select
                    Case &H18
                        Select Case (bit_idx)
                            Case 0
                                info = String.Format("{0}mS", bit_loop(bit_idx) * 3)
                            Case 1
                                Dim VioFreDes() As String = {"600kHz", "715kHz", "800kHz", "933kHz", "1000kHz", "1225kHz", "1225kHz", "1225kHz"}
                                info = VioFreDes(bit_loop(bit_idx))
                            Case 2
                                Dim VioSRDes() As String = {"120%", "100%", "80%", "60%"}
                                info = VioSRDes(bit_loop(bit_idx))
                        End Select
                    Case &H19

                        Select Case bit_idx
                            Case 0
                                Dim LDODlyDes() As String = {"0mS", "15mS", "34mS", "45mS"}
                                info = LDODlyDes(bit_loop(bit_idx))
                            Case 1
                                info = String.Format("{0}mS", 5 * bit_loop(bit_idx))
                        End Select
                    Case &H1A
                        info = String.Format("{0}mS", 5 * bit_loop(bit_idx))
                    Case &H1C
                        Select Case bit_idx
                            Case 0
                                Dim Des1() As String = {"Slowest", "Slow", "Normal", "Fastest"}
                                info = Des1(bit_loop(bit_idx))
                            Case 1
                                Dim Des2() As String = {"Slowest", "Slow", "Normal", "Fastest"}
                                info = Des2(bit_loop(bit_idx))
                            Case 2
                                Dim Des3() As String = {"PSM Mode", "PWM Mode"}
                                info = Des3(bit_loop(bit_idx))
                            Case 3
                                Dim Des4() As String = {"PSM Mode", "PWM Mode"}
                                info = Des4(bit_loop(bit_idx))
                            Case 4
                                Dim Des5() As String = {"PSM Mode", "PWM Mode"}
                                info = Des5(bit_loop(bit_idx))
                            Case 5
                                Dim Des6() As String = {"PSM Mode", "PWM Mode"}
                                info = Des6(bit_loop(bit_idx))
                        End Select
                End Select
                Reg_info.Text = info
                Panel1.Controls.Add(Reg_info)
                idx += 1
            End If
        Next

    End Sub

    Private Sub BT_Compare_Click(sender As Object, e As EventArgs) Handles BT_Compare.Click

        Panel1.Controls.Clear() ' Clear object
        Dim idx As Integer = 1

        ' test code
        'For i As Integer = 0 To byt_Bin1.Length - 1
        '    byt_Bin1(i) = &HFF
        'Next


        For i As Integer = 0 To (&H1D - &H1)

            If byt_Bin1(i) <> byt_Bin2(i) Then

                Select Case (i)
                    Case &H0, &H1B
                        Dim bit0 As Byte = ((byt_Bin1(i) Xor byt_Bin2(i)) And &H1) >> 0 ' detect bit0
                        Dim bit1 As Byte = ((byt_Bin1(i) Xor byt_Bin2(i)) And &H2) >> 1 ' detect bit1
                        Dim bit2 As Byte = ((byt_Bin1(i) Xor byt_Bin2(i)) And &H4) >> 2 ' detect bit2
                        Dim bit3 As Byte = ((byt_Bin1(i) Xor byt_Bin2(i)) And &H8) >> 3 ' detect bit3
                        Dim bit4 As Byte = ((byt_Bin1(i) Xor byt_Bin2(i)) And &H10) >> 4 ' detect bit4
                        Dim bit5 As Byte = ((byt_Bin1(i) Xor byt_Bin2(i)) And &H20) >> 5 ' detect bit5
                        Dim bit6 As Byte = ((byt_Bin1(i) Xor byt_Bin2(i)) And &H40) >> 6 ' detect bit6
                        Dim bit7 As Byte = ((byt_Bin1(i) Xor byt_Bin2(i)) And &H80) >> 7 ' detect bit7
                        Dim bit_loop() As Byte = {bit0, bit1, bit2, bit3, bit4, bit5, bit6, bit7}
                        AddCompare_Info(i, idx, bit_loop)
                    Case &H1
                        Dim bit0 As Byte = ((byt_Bin1(i) Xor byt_Bin2(i)) And &H1) >> 0 ' detect bit0
                        Dim bit1 As Byte = ((byt_Bin1(i) Xor byt_Bin2(i)) And &H2) >> 1 ' detect bit1
                        Dim bit4 As Byte = ((byt_Bin1(i) Xor byt_Bin2(i)) And &H10) >> 4 ' detect bit4
                        Dim bit5 As Byte = ((byt_Bin1(i) Xor byt_Bin2(i)) And &H20) >> 5 ' detect bit5
                        Dim bit_loop() As Byte = {bit0, bit1, bit4, bit5}
                        AddCompare_Info(i, idx, bit_loop)
                    Case &H2, &H3, &H4, &H5, &H6, &H7, &H8, &H9, &HA, &HB, &HC, &HD, &HE, &HF
                        AddCompare_Info(i, idx, New Byte() {byt_Bin1(i)})
                    Case &H10, &H11
                        Dim bit2_0 As Byte = ((byt_Bin1(i) Xor byt_Bin2(i)) And &H7) >> 0
                        Dim bit5_3 As Byte = ((byt_Bin1(i) Xor byt_Bin2(i)) And &H38) >> 3
                        Dim bit_loop() As Byte = {bit2_0, bit5_3}
                        AddCompare_Info(i, idx, bit_loop)
                    Case &H12, &H16
                        Dim bit2_0 As Byte = ((byt_Bin1(i) Xor byt_Bin2(i)) And &H7) >> 0
                        Dim bit_loop() As Byte = {bit2_0}
                        AddCompare_Info(i, idx, bit_loop)
                    Case &H13
                        Dim bit2_0 As Byte = ((byt_Bin1(i) Xor byt_Bin2(i)) And &H7) >> 0
                        Dim bit4_3 As Byte = ((byt_Bin1(i) Xor byt_Bin2(i)) And &H18) >> 3
                        Dim bit5 As Byte = ((byt_Bin1(i) Xor byt_Bin2(i)) And &H20) >> 5
                        Dim bit7_6 As Byte = ((byt_Bin1(i) Xor byt_Bin2(i)) And &HC0) >> 6
                        Dim bit_loop() As Byte = {bit2_0, bit4_3, bit5, bit7_6}
                        AddCompare_Info(i, idx, bit_loop)
                    Case &H14
                        Dim bit2_0 As Byte = ((byt_Bin1(i) Xor byt_Bin2(i)) And &H7) >> 0
                        Dim bit4_3 As Byte = ((byt_Bin1(i) Xor byt_Bin2(i)) And &H18) >> 3
                        Dim bit_loop() As Byte = {bit2_0, bit4_3}
                        AddCompare_Info(i, idx, bit_loop)
                    Case &H15
                        Dim bit2_0 As Byte = ((byt_Bin1(i) Xor byt_Bin2(i)) And &H7) >> 0
                        Dim bit4_3 As Byte = ((byt_Bin1(i) Xor byt_Bin2(i)) And &H18) >> 3
                        Dim bit7_5 As Byte = ((byt_Bin1(i) Xor byt_Bin2(i)) And &HE0) >> 5
                        Dim bit_loop() As Byte = {bit2_0, bit4_3, bit7_5}
                        AddCompare_Info(i, idx, bit_loop)
                    Case &H17, &H18
                        Dim bit1_0 As Byte = ((byt_Bin1(i) Xor byt_Bin2(i)) And &H3) >> 0
                        Dim bit4_2 As Byte = ((byt_Bin1(i) Xor byt_Bin2(i)) And &H1C) >> 2
                        Dim bit6_5 As Byte = ((byt_Bin1(i) Xor byt_Bin2(i)) And &H60) >> 5
                        Dim bit_loop() As Byte = {bit1_0, bit4_2, bit6_5}
                        AddCompare_Info(i, idx, bit_loop)
                    Case &H19
                        Dim bit1_0 As Byte = ((byt_Bin1(i) Xor byt_Bin2(i)) And &H3) >> 0
                        Dim bit6_2 As Byte = ((byt_Bin1(i) Xor byt_Bin2(i)) And &H7C) >> 2
                        Dim bit_loop() As Byte = {bit1_0, bit6_2}
                        AddCompare_Info(i, idx, bit_loop)
                    Case &H1A
                        Dim bit3_0 As Byte = ((byt_Bin1(i) Xor byt_Bin2(i)) And &HF) >> 0
                        Dim bit_loop() As Byte = {bit3_0}
                        AddCompare_Info(i, idx, bit_loop)
                    Case &H1C
                        Dim bit1_0 As Byte = ((byt_Bin1(i) Xor byt_Bin2(i)) And &H3) >> 0
                        Dim bit3_2 As Byte = ((byt_Bin1(i) Xor byt_Bin2(i)) And &HC) >> 2
                        Dim bit4 As Byte = ((byt_Bin1(i) Xor byt_Bin2(i)) And &H10) >> 4
                        Dim bit5 As Byte = ((byt_Bin1(i) Xor byt_Bin2(i)) And &H20) >> 5
                        Dim bit6 As Byte = ((byt_Bin1(i) Xor byt_Bin2(i)) And &H40) >> 6
                        Dim bit7 As Byte = ((byt_Bin1(i) Xor byt_Bin2(i)) And &H80) >> 7
                        Dim bit_loop() As Byte = {bit1_0, bit3_2, bit4, bit5, bit6, bit7}
                        AddCompare_Info(i, idx, bit_loop)
                End Select

            End If
        Next

    End Sub

    Private Sub BT_WriteAll_Click(sender As Object, e As EventArgs) Handles BT_WriteAll.Click

        Dim len As Integer = WRBuffer.Length
        WRBuffer(&H0) = W00_0.Value Or (W00_1.Value << 1) Or (W00_2.Value << 2) Or (W00_3.Value << 3) Or (W00_4.Value << 4) Or (W00_5.Value << 5) Or (W00_6.Value << 6) Or (W00_7.Value << 7)
        WRBuffer(&H1) = W01_0.Value Or (W01_1.Value << 1) Or (W01_4.Value << 4) Or (W01_5.Value << 5)
        WRBuffer(&H2) = W02.Value
        WRBuffer(&H3) = W03.Value
        WRBuffer(&H4) = W04.Value
        WRBuffer(&H5) = W05.Value
        WRBuffer(&H6) = W06.Value
        WRBuffer(&H7) = W07.Value
        WRBuffer(&H8) = W08.Value
        WRBuffer(&H9) = W09.Value
        WRBuffer(&HA) = 0
        WRBuffer(&HB) = W0B.Value
        WRBuffer(&HC) = W0C.Value
        WRBuffer(&HD) = W0D.Value
        WRBuffer(&HE) = W0E.Value
        WRBuffer(&HF) = W0F.Value
        WRBuffer(&H10) = W10_0.Value Or W10_3.Value << 3
        WRBuffer(&H11) = W11_0.Value Or W11_3.Value << 3
        WRBuffer(&H12) = W12_0.Value
        WRBuffer(&H13) = W13_0.Value Or (W13_3.Value << 3) Or (W13_5.Value << 5) Or (W13_6.Value << 6)
        WRBuffer(&H14) = W14_0.Value Or (W14_3.Value << 3)
        WRBuffer(&H15) = W15_0.Value Or (W15_3.Value << 3) Or (W15_5.Value << 5)
        WRBuffer(&H16) = W16_0.Value
        WRBuffer(&H17) = W17_0.Value Or W17_2.Value << 2 Or W17_5.Value
        WRBuffer(&H18) = W18_0.Value Or W18_2.Value << 2 Or W18_5.Value << 5
        WRBuffer(&H19) = W19_0.Value Or W19_2.Value << 2
        WRBuffer(&H1A) = W1A.Value
        WRBuffer(&H1B) = W1B_0.Value Or (W1B_1.Value << 1) Or (W1B_2.Value << 2) Or (W1B_3.Value << 3) Or (W1B_4.Value << 4) Or (W1B_5.Value << 5) Or (W1B_6.Value << 6) Or (W1B_7.Value << 7)
        WRBuffer(&H1C) = W1C_0.Value Or (W1C_2.Value << 2) Or (W1C_4.Value << 4) Or (W1C_5.Value << 5) Or (W1C_6.Value << 6) Or (W1C_7.Value << 7)
        WRBuffer(&H1D) = W1D.Value
        WRBuffer(&H1E) = W1E.Value
        WRBuffer(&H1F) = W1F.Value

        RTBB_I2CWrite(hDevice, 0, Int(NuSlave.Value / 2), &H1, &H0, len, WRBuffer(0))

        Dim res As Integer = 0

        For i As Integer = 0 To WRBuffer.Length - 1
            res += WRBuffer(i)
        Next

        CheckSum_wrall.Text = res.ToString("X")
        CheckSum_wrvcom.Text = (res - WRBuffer(&HE)).ToString("X")

        CalculateResult()

    End Sub

    Private Sub BT_ReadAll_Click(sender As Object, e As EventArgs) Handles BT_ReadAll.Click

        ' DAC read command
        Dim addr As Integer = &HFF
        Dim Data() As Byte = New Byte() {&H0}
        RTBB_I2CWrite(hDevice, 0, Int(NuSlave.Value / 2), &H1, addr, Data.Length, Data(0))


        ' read data
        Dim len As Integer = RDBuffer.Length - 1
        RTBB_I2CRead(hDevice, 0, Int(NuSlave.Value / 2), &H1, &H0, len, RDBuffer(0))

        Dim res As Integer = 0

        For i As Integer = 0 To RDBuffer.Length - 1
            res += RDBuffer(i)
        Next

        CheckSum_rdall.Text = res.ToString("X")
        CheckSum_rdvcom.Text = (res - RDBuffer(&HE)).ToString("X")

        W00_0.Value = (RDBuffer(0) And &H1) >> 0
        W00_1.Value = (RDBuffer(0) And &H2) >> 1
        W00_2.Value = (RDBuffer(0) And &H4) >> 2
        W00_3.Value = (RDBuffer(0) And &H8) >> 3
        W00_4.Value = (RDBuffer(0) And &H10) >> 4
        W00_5.Value = (RDBuffer(0) And &H20) >> 5
        W00_6.Value = (RDBuffer(0) And &H40) >> 6
        W00_7.Value = (RDBuffer(0) And &H80) >> 7

        W01_0.Value = (RDBuffer(1) And &H1) >> 0
        W01_1.Value = (RDBuffer(1) And &H2) >> 1
        W01_4.Value = (RDBuffer(1) And &H10) >> 4
        W01_5.Value = (RDBuffer(1) And &H20) >> 5

        W02.Value = RDBuffer(2)
        W03.Value = RDBuffer(3)
        W04.Value = RDBuffer(4)
        W05.Value = RDBuffer(5)
        W06.Value = RDBuffer(6)
        W07.Value = RDBuffer(7)
        W08.Value = RDBuffer(8)
        W09.Value = RDBuffer(9)
        'W0a.Value = RDBuffer(4)
        W0B.Value = RDBuffer(&HB)
        W0C.Value = RDBuffer(&HC)
        W0D.Value = RDBuffer(&HD)
        W0E.Value = RDBuffer(&HE)
        W0F.Value = RDBuffer(&HF)

        W10_0.Value = RDBuffer(&H10) And &H7
        W10_3.Value = (RDBuffer(&H10) And &H38) >> 3

        W11_0.Value = RDBuffer(&H11) And &H7
        W11_3.Value = (RDBuffer(&H11) And &H38) >> 3

        W12_0.Value = RDBuffer(&H12)
        W13_0.Value = RDBuffer(&H13) And &H7
        W13_3.Value = (RDBuffer(&H13) And &H18) >> 3
        W13_5.Value = (RDBuffer(&H13) And &H20) >> 5
        W13_6.Value = (RDBuffer(&H13) And &HC0) >> 6

        W14_0.Value = RDBuffer(&H14) And &H7
        W14_3.Value = (RDBuffer(&H14) And &H18) >> 3

        W15_0.Value = RDBuffer(&H15) And &H7
        W15_3.Value = (RDBuffer(&H15) And &H18) >> 3
        W15_5.Value = (RDBuffer(&H15) And &HE0) >> 5

        W16_0.Value = RDBuffer(&H16)

        W17_0.Value = (RDBuffer(&H17) And &H3)
        W17_2.Value = (RDBuffer(&H17) And &H1C) >> 2
        W17_5.Value = (RDBuffer(&H17) And &H60) >> 5

        W18_0.Value = (RDBuffer(&H18) And &H3)
        W18_2.Value = (RDBuffer(&H18) And &H1C) >> 2
        W18_5.Value = (RDBuffer(&H18) And &H60) >> 5

        W19_0.Value = (RDBuffer(&H19) And &H3)
        W19_2.Value = (RDBuffer(&H19) And &H7C) >> 2
        W1A.Value = RDBuffer(&H1A)

        W1B_0.Value = (RDBuffer(&H1B) And &H1) >> 0
        W1B_1.Value = (RDBuffer(&H1B) And &H2) >> 1
        W1B_2.Value = (RDBuffer(&H1B) And &H4) >> 2
        W1B_3.Value = (RDBuffer(&H1B) And &H8) >> 3
        W1B_4.Value = (RDBuffer(&H1B) And &H10) >> 4
        W1B_5.Value = (RDBuffer(&H1B) And &H20) >> 5
        W1B_6.Value = (RDBuffer(&H1B) And &H40) >> 6
        W1B_7.Value = (RDBuffer(&H1B) And &H80) >> 7

        W1C_0.Value = (RDBuffer(&H1C) And &H3)
        W1C_2.Value = (RDBuffer(&H1C) And &HC) >> 2
        W1C_4.Value = (RDBuffer(&H1C) And &H10) >> 4
        W1C_5.Value = (RDBuffer(&H1C) And &H20) >> 5
        W1C_6.Value = (RDBuffer(&H1C) And &H40) >> 6
        W1C_7.Value = (RDBuffer(&H1C) And &H80) >> 7

        W1D.Value = RDBuffer(&H1D)
        W1E.Value = RDBuffer(&H1E)
        W1F.Value = RDBuffer(&H1F)



        CalculateResult()

    End Sub

    Private Sub BT_WriteMTP_Click(sender As Object, e As EventArgs) Handles BT_WriteMTP.Click

        BT_WriteAll_Click(Nothing, Nothing)

        ' MTP program command
        ' Address 0x80 data 0x80
        Dim addr As Integer = &HFF
        Dim Data() As Byte = New Byte() {&H80}

        RTBB_I2CWrite(hDevice, 0, Int(NuSlave.Value / 2), &H1, addr, Data.Length, Data(0))
        System.Threading.Thread.Sleep(500)
    End Sub

    Private Sub BT_ReadMTP_Click(sender As Object, e As EventArgs) Handles BT_ReadMTP.Click

        System.Threading.Thread.Sleep(500)
        ' MTP read command
        ' Address 0x80 data 0x01
        Dim addr As Integer = &HFF
        Dim Data() As Byte = New Byte() {&H1}
        RTBB_I2CWrite(hDevice, 0, Int(NuSlave.Value / 2), &H1, addr, Data.Length, Data(0))


        ' read data
        Dim len As Integer = RDBuffer.Length - 1
        Array.Clear(RDBuffer, 0, RDBuffer.Length)
        RTBB_I2CRead(hDevice, 0, Int(NuSlave.Value / 2), &H1, &H0, len, RDBuffer(0))

        Dim res As Integer = 0

        For i As Integer = 0 To RDBuffer.Length - 1
            res += RDBuffer(i)
        Next

        CheckSum_rdall.Text = res.ToString("X")
        CheckSum_rdvcom.Text = (res - RDBuffer(&HE)).ToString("X")

        W00_0.Value = (RDBuffer(0) And &H1) >> 0
        W00_1.Value = (RDBuffer(0) And &H2) >> 1
        W00_2.Value = (RDBuffer(0) And &H4) >> 2
        W00_3.Value = (RDBuffer(0) And &H8) >> 3
        W00_4.Value = (RDBuffer(0) And &H10) >> 4
        W00_5.Value = (RDBuffer(0) And &H20) >> 5
        W00_6.Value = (RDBuffer(0) And &H40) >> 6
        W00_7.Value = (RDBuffer(0) And &H80) >> 7

        W01_0.Value = (RDBuffer(1) And &H1) >> 0
        W01_1.Value = (RDBuffer(1) And &H2) >> 1
        W01_4.Value = (RDBuffer(1) And &H10) >> 4
        W01_5.Value = (RDBuffer(1) And &H20) >> 5

        W02.Value = RDBuffer(2)
        W03.Value = RDBuffer(3)
        W04.Value = RDBuffer(4)
        W05.Value = RDBuffer(5)
        W06.Value = RDBuffer(6)
        W07.Value = RDBuffer(7)
        W08.Value = RDBuffer(8)
        W09.Value = RDBuffer(9)
        'W0a.Value = RDBuffer(4)
        W0B.Value = RDBuffer(&HB)
        W0C.Value = RDBuffer(&HC)
        W0D.Value = RDBuffer(&HD)
        W0E.Value = RDBuffer(&HE)
        W0F.Value = RDBuffer(&HF)

        W10_0.Value = RDBuffer(&H10) And &H7
        W10_3.Value = (RDBuffer(&H10) And &H38) >> 3

        W11_0.Value = RDBuffer(&H11) And &H7
        W11_3.Value = (RDBuffer(&H11) And &H38) >> 3

        W12_0.Value = RDBuffer(&H12)
        W13_0.Value = RDBuffer(&H13) And &H7
        W13_3.Value = (RDBuffer(&H13) And &H18) >> 3
        W13_5.Value = (RDBuffer(&H13) And &H20) >> 5
        W13_6.Value = (RDBuffer(&H13) And &HC0) >> 6

        W14_0.Value = RDBuffer(&H14) And &H7
        W14_3.Value = (RDBuffer(&H14) And &H18) >> 3

        W15_0.Value = RDBuffer(&H15) And &H7
        W15_3.Value = (RDBuffer(&H15) And &H18) >> 3
        W15_5.Value = (RDBuffer(&H15) And &HE0) >> 5

        W16_0.Value = RDBuffer(&H16)

        W17_0.Value = (RDBuffer(&H17) And &H3)
        W17_2.Value = (RDBuffer(&H17) And &H1C) >> 2
        W17_5.Value = (RDBuffer(&H17) And &H60) >> 5

        W18_0.Value = (RDBuffer(&H18) And &H3)
        W18_2.Value = (RDBuffer(&H18) And &H1C) >> 2
        W18_5.Value = (RDBuffer(&H18) And &H60) >> 5

        W19_0.Value = (RDBuffer(&H19) And &H3)
        W19_2.Value = (RDBuffer(&H19) And &H7C) >> 2
        W1A.Value = RDBuffer(&H1A)

        W1B_0.Value = (RDBuffer(&H1B) And &H1) >> 0
        W1B_1.Value = (RDBuffer(&H1B) And &H2) >> 1
        W1B_2.Value = (RDBuffer(&H1B) And &H4) >> 2
        W1B_3.Value = (RDBuffer(&H1B) And &H8) >> 3
        W1B_4.Value = (RDBuffer(&H1B) And &H10) >> 4
        W1B_5.Value = (RDBuffer(&H1B) And &H20) >> 5
        W1B_6.Value = (RDBuffer(&H1B) And &H40) >> 6
        W1B_7.Value = (RDBuffer(&H1B) And &H80) >> 7

        W1C_0.Value = (RDBuffer(&H1C) And &H3)
        W1C_2.Value = (RDBuffer(&H1C) And &HC) >> 2
        W1C_4.Value = (RDBuffer(&H1C) And &H10) >> 4
        W1C_5.Value = (RDBuffer(&H1C) And &H20) >> 5
        W1C_6.Value = (RDBuffer(&H1C) And &H40) >> 6
        W1C_7.Value = (RDBuffer(&H1C) And &H80) >> 7

        W1D.Value = RDBuffer(&H1D)
        W1E.Value = RDBuffer(&H1E)
        W1F.Value = RDBuffer(&H1F)



        CalculateResult()

    End Sub

    Private Sub BT_SingleWrite_Click(sender As Object, e As EventArgs) Handles BT_SingleWrite.Click
        Dim addr As Integer = NuAddr.Value
        Dim Data() As Byte = New Byte() {NuWrData.Value}
        RTBB_I2CWrite(hDevice, 0, Int(NuSlave.Value / 2), &H1, addr, Data.Length, Data(0))
    End Sub

    Private Sub BT_SingleRead_Click(sender As Object, e As EventArgs) Handles BT_SingleRead.Click
        ' DAC Read command
        Dim addr As Integer = &HFF
        Dim Data() As Byte = New Byte() {&H0}
        RTBB_I2CWrite(hDevice, 0, Int(NuSlave.Value / 2), &H1, addr, Data.Length, Data(0))

        addr = NuAddr.Value
        RTBB_I2CRead(hDevice, 0, Int(NuSlave.Value / 2), &H1, addr, Data.Length, Data(0))
        NuRdData.Value = Data(0)
    End Sub

    Private Sub BT_Single_ReadMTP_Click(sender As Object, e As EventArgs) Handles BT_Single_ReadMTP.Click
        ' MTP Read command
        Dim addr As Integer = &HFF
        Dim Data() As Byte = New Byte() {&H1}
        RTBB_I2CWrite(hDevice, 0, Int(NuSlave.Value / 2), &H1, addr, Data.Length, Data(0))

        addr = NuAddr.Value
        RTBB_I2CRead(hDevice, 0, Int(NuSlave.Value / 2), &H1, addr, Data.Length, Data(0))
        NuMTPRd.Value = Data(0)
    End Sub

    Private Sub W0C_ValueChanged(sender As Object, e As EventArgs) Handles W0C.ValueChanged
        Bar_GMA1.Value = W0C.Value
        Dim offset As Double
        Dim AVDD As Double = (W02.Value * 0.1) + 7
        If (AVDD > 14) Then
            AVDD = 14
        End If

        If W0C.Value <= &H12 Then
            offset = (W0C.Value + 2) * 0.05
        Else
            offset = (&H12 + 2) * 0.05
        End If
        GMA1V.Text = String.Format("{0:0.00}V", AVDD - offset)
    End Sub

    Private Sub CB_A0_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CB_A0.SelectedIndexChanged
        Select Case CB_A0.SelectedIndex
            Case 0
                NuSlave.Value = &H9C
            Case 1
                NuSlave.Value = &H9E
        End Select
    End Sub

    Private Sub SaveBinToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SaveBinToolStripMenuItem.Click
        WRBuffer(0) = W00_0.Value Or (W00_1.Value << 1) Or (W00_2.Value << 2) Or (W00_3.Value << 3) Or (W00_4.Value << 4) Or (W00_5.Value << 5) Or (W00_6.Value << 6) Or (W00_7.Value << 7)
        WRBuffer(1) = W01_0.Value Or (W01_1.Value << 1) Or (W01_4.Value << 4) Or (W01_5.Value << 5)
        WRBuffer(2) = W02.Value
        WRBuffer(3) = W03.Value
        WRBuffer(4) = W04.Value
        WRBuffer(5) = W05.Value
        WRBuffer(6) = W06.Value
        WRBuffer(7) = W07.Value
        WRBuffer(8) = W08.Value
        WRBuffer(9) = W09.Value
        WRBuffer(&HA) = 0
        WRBuffer(&HB) = W0B.Value
        WRBuffer(&HC) = W0C.Value
        WRBuffer(&HD) = W0D.Value
        WRBuffer(&HE) = W0E.Value
        WRBuffer(&HF) = W0F.Value
        WRBuffer(&H10) = W10_0.Value Or W10_3.Value << 3
        WRBuffer(&H11) = W11_0.Value Or W11_3.Value << 3
        WRBuffer(&H12) = W12_0.Value
        WRBuffer(&H13) = W13_0.Value Or (W13_3.Value << 3) Or (W13_5.Value << 5) Or (W13_6.Value << 6)
        WRBuffer(&H14) = W14_0.Value Or (W14_3.Value << 3)
        WRBuffer(&H15) = W15_0.Value Or (W15_3.Value << 3) Or (W15_5.Value << 5)
        WRBuffer(&H16) = W16_0.Value
        WRBuffer(&H17) = W17_0.Value Or W17_2.Value << 2 Or W17_5.Value
        WRBuffer(&H18) = W18_0.Value Or W18_2.Value << 2 Or W18_5.Value << 5
        WRBuffer(&H19) = W19_0.Value Or W19_2.Value << 2
        WRBuffer(&H1A) = W1A.Value
        WRBuffer(&H1B) = W1B_0.Value Or (W1B_1.Value << 1) Or (W1B_2.Value << 2) Or (W1B_3.Value << 3) Or (W1B_4.Value << 4) Or (W1B_5.Value << 5) Or (W1B_6.Value << 6) Or (W1B_7.Value << 7)
        WRBuffer(&H1C) = W1C_0.Value Or (W1C_2.Value << 2) Or (W1C_4.Value << 4) Or (W1C_5.Value << 5) Or (W1C_6.Value << 6) Or (W1C_7.Value << 7)
        WRBuffer(&H1D) = W1D.Value
        WRBuffer(&H1E) = W1E.Value
        WRBuffer(&H1F) = W1F.Value

        Dim savedlg As SaveFileDialog = New SaveFileDialog()
        savedlg.Filter = "Bin Files (*.bin) | *.bin|All files (*.*)|*.*"
        savedlg.RestoreDirectory = True
        If savedlg.ShowDialog() = DialogResult.OK Then
            Dim myFile As FileStream = New FileStream(savedlg.FileName, FileMode.OpenOrCreate)
            Dim bwr As BinaryWriter = New BinaryWriter(myFile)

            bwr.Write(WRBuffer, 0, WRBuffer.Length)
            bwr.Close()
            myFile.Close()
        End If
    End Sub

    Private Sub LoadBinToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LoadBinToolStripMenuItem.Click


        Dim opendlg As OpenFileDialog = New OpenFileDialog()
        opendlg.Filter = "Bin Files (*.bin) | *.bin|All files (*.*)|*.*"
        opendlg.RestoreDirectory = True
        If opendlg.ShowDialog() = DialogResult.OK Then
            Dim myFile As FileStream = New FileStream(opendlg.FileName, FileMode.OpenOrCreate)
            Dim brd As BinaryReader = New BinaryReader(myFile)

            brd.Read(RDBuffer, 0, RDBuffer.Length)

            W00_0.Value = (RDBuffer(0) And &H1) >> 0
            W00_1.Value = (RDBuffer(0) And &H2) >> 1
            W00_2.Value = (RDBuffer(0) And &H4) >> 2
            W00_3.Value = (RDBuffer(0) And &H8) >> 3
            W00_4.Value = (RDBuffer(0) And &H10) >> 4
            W00_5.Value = (RDBuffer(0) And &H20) >> 5
            W00_6.Value = (RDBuffer(0) And &H40) >> 6
            W00_7.Value = (RDBuffer(0) And &H80) >> 7

            W01_0.Value = (RDBuffer(1) And &H1) >> 0
            W01_1.Value = (RDBuffer(1) And &H2) >> 1
            W01_4.Value = (RDBuffer(1) And &H10) >> 4
            W01_5.Value = (RDBuffer(1) And &H20) >> 5

            W02.Value = RDBuffer(2)
            W03.Value = RDBuffer(3)
            W04.Value = RDBuffer(4)
            W05.Value = RDBuffer(5)
            W06.Value = RDBuffer(6)
            W07.Value = RDBuffer(7)
            W08.Value = RDBuffer(8)
            W09.Value = RDBuffer(9)
            'W0a.Value = RDBuffer(4)
            W0B.Value = RDBuffer(&HB)
            W0C.Value = RDBuffer(&HC)
            W0D.Value = RDBuffer(&HD)
            W0E.Value = RDBuffer(&HE)
            W0F.Value = RDBuffer(&HF)

            W10_0.Value = RDBuffer(&H10) And &H7
            W10_3.Value = (RDBuffer(&H10) And &H38) >> 3

            W11_0.Value = RDBuffer(&H11) And &H7
            W11_3.Value = (RDBuffer(&H11) And &H38) >> 3

            W12_0.Value = RDBuffer(&H12)
            W13_0.Value = RDBuffer(&H13) And &H7
            W13_3.Value = (RDBuffer(&H13) And &H18) >> 3
            W13_5.Value = (RDBuffer(&H13) And &H20) >> 5
            W13_6.Value = (RDBuffer(&H13) And &HC0) >> 6

            W14_0.Value = RDBuffer(&H14) And &H7
            W14_3.Value = (RDBuffer(&H14) And &H18) >> 3

            W15_0.Value = RDBuffer(&H15) And &H7
            W15_3.Value = (RDBuffer(&H15) And &H18) >> 3
            W15_5.Value = (RDBuffer(&H15) And &HE0) >> 5

            W16_0.Value = RDBuffer(&H16)

            W17_0.Value = (RDBuffer(&H17) And &H3)
            W17_2.Value = (RDBuffer(&H17) And &H1C) >> 2
            W17_5.Value = (RDBuffer(&H17) And &H60) >> 5

            W18_0.Value = (RDBuffer(&H18) And &H3)
            W18_2.Value = (RDBuffer(&H18) And &H1C) >> 2
            W18_5.Value = (RDBuffer(&H18) And &H60) >> 5

            W19_0.Value = (RDBuffer(&H19) And &H3)
            W19_2.Value = (RDBuffer(&H19) And &H7C) >> 2
            W1A.Value = RDBuffer(&H1A)

            W1B_0.Value = (RDBuffer(&H1B) And &H1) >> 0
            W1B_1.Value = (RDBuffer(&H1B) And &H2) >> 1
            W1B_2.Value = (RDBuffer(&H1B) And &H4) >> 2
            W1B_3.Value = (RDBuffer(&H1B) And &H8) >> 3
            W1B_4.Value = (RDBuffer(&H1B) And &H10) >> 4
            W1B_5.Value = (RDBuffer(&H1B) And &H20) >> 5
            W1B_6.Value = (RDBuffer(&H1B) And &H40) >> 6
            W1B_7.Value = (RDBuffer(&H1B) And &H80) >> 7

            W1C_0.Value = (RDBuffer(&H1C) And &H3)
            W1C_2.Value = (RDBuffer(&H1C) And &HC) >> 2
            W1C_4.Value = (RDBuffer(&H1C) And &H10) >> 4
            W1C_5.Value = (RDBuffer(&H1C) And &H20) >> 5
            W1C_6.Value = (RDBuffer(&H1C) And &H40) >> 6
            W1C_7.Value = (RDBuffer(&H1C) And &H80) >> 7

            W1D.Value = RDBuffer(&H1D)
            W1E.Value = RDBuffer(&H1E)
            W1F.Value = RDBuffer(&H1F)

            brd.Close()
            myFile.Close()

        End If


    End Sub

    Private Sub W01_5_ValueChanged(sender As Object, e As EventArgs) Handles W01_5.ValueChanged, W01_4.ValueChanged, W01_1.ValueChanged, W01_0.ValueChanged
        Dim byt01_bit() As NumericUpDown = {W01_0, W01_1, W01_4, W01_5}
        Dim byt01_ck() As CheckBox = {CK01_0, CK01_1, CK01_4, CK01_5}

        For i As Integer = 0 To byt01_bit.Length - 1
            byt01_ck(i).Checked = byt01_bit(i).Value And &H1
        Next
    End Sub

    Private Sub W1D_ValueChanged(sender As Object, e As EventArgs) Handles W1D.ValueChanged
        Bar_Inx1.Value = W1D.Value
    End Sub

    Private Sub W1E_ValueChanged(sender As Object, e As EventArgs) Handles W1E.ValueChanged
        Bar_Inx2.Value = W1E.Value
    End Sub

    Private Sub W1F_ValueChanged(sender As Object, e As EventArgs) Handles W1F.ValueChanged
        Bar_Inx3.Value = W1F.Value
    End Sub

    Private Sub LinkBridgeBoardToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LinkBridgeBoardToolStripMenuItem.Click
        hDevice = RTBB_ConnectToBridgeByIndex(0)
        If hDevice <> 0 Then
            'MsgBox("Link RTBridge Board Successful!!", main_ver, MsgBoxStyle.DefaultButton1)
            MessageBox.Show("Link RTBridge Board Successful!!", main_ver, MessageBoxButtons.OK)
        Else
            'MsgBox("Please Link RTBridge Board!!", main_ver, MsgBoxStyle.DefaultButton1)
            MessageBox.Show("Please Link RTBridge Board!!", main_ver, MessageBoxButtons.OK)
        End If
    End Sub
End Class
