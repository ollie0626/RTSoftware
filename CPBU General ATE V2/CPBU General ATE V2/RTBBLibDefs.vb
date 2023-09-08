Module RTBBLibDefs

    Public hDevice As Integer
    Public I2CBus As Integer
    Public hEnum As Integer
    Public pEnumBoardInfo As Integer
    Public BoardCount As Integer
    Public FreeEnumBoard As Boolean
    Public strVendorName As String
    Public strControllerName As String
    Public strLibraryName As String
    Public strLibraryPath As String
    Public strFirmwareInfo As String
    Public strDevicePath As String
    Public strBoardName As String
    Public nMcuIndexOfDevice As Integer
    Public nMcuVID As UInteger
    Public nMcuPID As UInteger
    Public nMcuCapability As Long
    Public nMcuGPIOBitsType As Long
    Public nMcuGPIOPinCount As Long
    Public nMcuI2CCount As Long
    Public nMcuSPICount As Long
    Public nMcuUARTCount As Long

    'GPIO
    Public nResult0, nResult1 As Integer

    Public nLength As UInteger = 4

    Public nMask As UInteger = CInt("&h00FF")

    Public bCheckAck As Boolean = False

    Public GPIO_Data(100) As Short



    'RT_BRIDGE_BOARD_DEFS_H

    Public Const RT_BB_SUCCESS As Short = 0
    Public Const RT_BB_BAD_PARAMETER As Short = -1
    Public Const RT_BB_HARDWARE_NOT_FOUND As Short = -2
    Public Const RT_BB_SLAVE_DEVICE_NOT_FOUND As Short = -3
    Public Const RT_BB_TRANSACTION_FAILED As Short = -4
    Public Const RT_BB_SLAVE_OPENNING_FOR_WRITE_FAILED As Short = -5
    Public Const RT_BB_SLAVE_OPENNING_FOR_READ_FAILED As Short = -6
    Public Const RT_BB_SENDING_MEMORY_ADDRESS_FAILED As Short = -7
    Public Const RT_BB_SENDING_DATA_FAILED As Short = -8
    Public Const RT_BB_NOT_IMPLEMENTED As Short = -9
    Public Const RT_BB_NO_ACK As Short = -10
    Public Const RT_BB_DEVICE_BUSY As Short = -11
    Public Const RT_BB_MEMORY_ERROR As Short = -12
    Public Const RT_BB_UNKNOWN_ERROR As Short = -13
    Public Const RT_BB_I2C_TIMEOUT As Short = -14
    Public Const RT_BB_IDLE As Short = -15
    Public Const RT_BB_NO_DATA As Short = -16
    Public Const RT_BB_BUFFER_OVERFLOW As Short = -17
    Public Const RT_BB_HARDWARE_NOT_SUPPORT As Short = -18
    Public Const RT_BB_MEMORY_ACCESS_ERROR As Short = -19
    Public Const RT_BB_GSMW_PIN_MASK_ERROR As Short = -20
    Public Const RT_BB_SVI2_TIMEOUT As Short = -21
    Public Const RT_BB_SVI2_NO_POWER_OK As Short = -22
    Public Const RT_BB_SVI2_ALREADY_BOOTUP As Short = -23
    Public Const RT_BB_SVI2_ALREADY_POWEROFF As Short = -24
    Public Const RT_BB_USB_COMMUNICATION_FAILED As Short = -25







    'RTI2C_DEFS_H
    Public Const RT_I2C_FREQ_FAST As Short = &H1
    Public Const RT_I2C_FREQ_STD As Short = &H2
    Public Const RT_I2C_FREQ_83KHZ As Short = &H4
    Public Const RT_I2C_FREQ_71KHZ As Short = &H8
    Public Const RT_I2C_FREQ_62KHZ As Short = &H10
    Public Const RT_I2C_FREQ_50KHZ As Short = &H20
    Public Const RT_I2C_FREQ_25KHZ As Short = &H40
    Public Const RT_I2C_FREQ_10KHZ As Short = &H80
    Public Const RT_I2C_FREQ_5KHZ As Short = &H100
    Public Const RT_I2C_FREQ_2KHZ As Short = &H200
    Public Const RT_I2C_FREQ_1MHZ As Short = &H400
    Public Const RT_I2C_FREQ_HS As Short = &H800
    Public Const RT_I2C_FREQ_CUSTOM As Integer = &H40000000 '(1 << 30)

    'RTEXTGPIOMISC_DEFS_H
    Public Const GPIOMISC_PINMODE_NORMAL As Short = 0
    Public Const GPIOMISC_PINMODE_PULLUP As Short = 1
    Public Const GPIOMISC_PINMODE_PULLDOWN As Short = 2
    Public Const GPIOMISC_PINMODE_REPEATER As Short = 3

    Public Const GPIOMISC_PINSEL_GPIOFUNC As Short = 0
    Public Const GPIOMISC_PINSEL_I2CFUNC As Short = 1
    Public Const GPIOMISC_PINSEL_SPIFUNC As Short = 2
    Public Const GPIOMISC_PINSEL_USBFUNC As Short = 3
    Public Const GPIOMISC_PINSEL_UART As Short = 4
    Public Const GPIOMISC_PINSEL_PWM As Short = 5
    Public Const GPIOMISC_PINSEL_I2S As Short = 6
    Public Const GPIOMISC_PINSEL_ADC As Short = 7
    Public Const GPIOMISC_PINSEL_DAC As Short = 8
    Public Const GPIOMISC_PINSEL_CAN As Short = 9
    Public Const GPIOMISC_PINSEL_OTHER As Short = 15

    'RTSPI_DEFS_H
    Public Const RT_SPI_FREQ_400KHZ As Short = &H1 '(1 << 0)
    Public Const RT_SPI_FREQ_200KHZ As Short = &H2 '(1 << 1)
    Public Const RT_SPI_FREQ_100KHZ As Short = &H4 '(1 << 2)
    Public Const RT_SPI_FREQ_83KHZ As Short = &H8 '(1 << 3)
    Public Const RT_SPI_FREQ_71KHZ As Short = &H10 '(1 << 4)
    Public Const RT_SPI_FREQ_62KHZ As Short = &H20 '(1 << 5)
    Public Const RT_SPI_FREQ_50KHZ As Short = &H40 '(1 << 6)
    Public Const RT_SPI_FREQ_25KHZ As Short = &H80 '(1 << 7)
    Public Const RT_SPI_FREQ_10KHZ As Short = &H100 ' (1 << 8)
    Public Const RT_SPI_FREQ_5KHZ As Short = &H200 '(1 << 9)
    Public Const RT_SPI_FREQ_2KHZ As Short = &H400 '(1 << 10)
    Public Const RT_SPI_FREQ_1MHZ As Short = &H800 '(1 << 11)
    Public Const RT_SPI_FREQ_2MHZ As Short = &H1000 '(1 << 12)
    Public Const RT_SPI_FREQ_3MHZ As Short = &H2000 '(1 << 13)
    Public Const RT_SPI_FREQ_5MHZ As Short = &H4000 '(1 << 14)
    Public Const RT_SPI_FREQ_6MHZ As Integer = &H8000 '(1 << 15)
    Public Const RT_SPI_FREQ_10MHZ As Integer = &H10000 '(1 << 16)
    Public Const RT_SPI_FREQ_12MHZ As Integer = &H20000 '(1 << 17)
    Public Const RT_SPI_FREQ_15MHZ As Integer = &H40000 '(1 << 18)
    Public Const RT_SPI_FREQ_20MHZ As Integer = &H80000 '(1 << 19)
    Public Const RT_SPI_FREQ_24MHZ As Integer = &H100000 '(1 << 20)
    Public Const RT_SPI_FREQ_30MHZ As Integer = &H200000 '(1 << 21)
    Public Const RT_SPI_FREQ_36MHZ As Integer = &H400000 '(1 << 22)
    Public Const RT_SPI_FREQ_60MHZ As Integer = &H800000 '(1 << 23)
    Public Const RT_SPI_FREQ_CUSTOM As Integer = &H40000000 '(1 << 30)

    Function i2c_status(ByVal num As Integer) As String
        Dim status As String = ""

        Select Case num
            Case 0
                status = "RT_BB_SUCCESS"
            Case -1
                status = "RT_BB_BAD_PARAMETER"
            Case -2
                status = "RT_BB_HARDWARE_NOT_FOUND"
            Case -3
                status = "RT_BB_SLAVE_DEVICE_NOT_FOUND"
            Case -4
                status = "RT_BB_TRANSACTION_FAILED"
            Case -5
                status = "RT_BB_SLAVE_OPENNING_FOR_WRITE_FAILED"
            Case -6
                status = "RT_BB_SLAVE_OPENNING_FOR_READ_FAILED"
            Case -7
                status = "RT_BB_SENDING_MEMORY_ADDRESS_FAILED"
            Case -8
                status = "RT_BB_SENDING_DATA_FAILED"
            Case -9
                status = "RT_BB_NOT_IMPLEMENTED"
            Case -10
                status = "RT_BB_NO_ACK"
            Case -11
                status = "RT_BB_DEVICE_BUSY"
            Case -12
                status = "RT_BB_MEMORY_ERROR"
            Case -13
                status = "RT_BB_UNKNOWN_ERROR"
            Case -14
                status = "RT_BB_I2C_TIMEOUT"
            Case -15
                status = "RT_BB_IDLE"
            Case -16
                status = "RT_BB_NO_DATA"
            Case -17
                status = "RT_BB_BUFFER_OVERFLOW"
            Case -18
                status = "RT_BB_HARDWARE_NOT_SUPPORT"
            Case -19
                status = "RT_BB_MEMORY_ACCESS_ERROR"
            Case -20
                status = "RT_BB_GSMW_PIN_MASK_ERROR"
            Case -21
                status = "RT_BB_SVI2_TIMEOUT"
            Case -22
                status = "RT_BB_SVI2_NO_POWER_OK"
            Case -23
                status = "RT_BB_SVI2_ALREADY_BOOTUP"
            Case -24
                status = "RT_BB_SVI2_ALREADY_POWEROFF"
            Case -25
                status = "RT_BB_USB_COMMUNICATION_FAILED"

        End Select

        Return status

    End Function


End Module
